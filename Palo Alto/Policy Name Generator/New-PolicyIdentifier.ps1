<#
.SYNOPSIS
    Generates an opaque, collision-checked random identifier for a PAN-OS policy
    by checking live Panorama configuration directly (no local tracking file).

.DESCRIPTION
    Prompts for a policy type, then performs a single bulk fetch of every
    existing rule name of that type across all device groups (pre-rulebase
    and post-rulebase) and the shared rulebase. Generates an 8-character
    random alphanumeric string (excluding visually ambiguous characters:
    0, O, 1, I, L), and checks it locally against the fetched name set —
    regenerating without further API calls if a collision occurs. No local
    CSV/database is used as the source of truth for uniqueness.

.NOTES
    Full Policy Name format: <PREFIX>-POL-XXXXXXXX  (e.g., SEC-POL-7F3KQ9NX)

    CONFIRMED against live Panorama:
      - Bulk <rules> container fetch via action=get returns
        $result.response.result.rules.entry as expected — multiple <entry
        name="..."> nodes directly under <rules>.
      - Rulebase tag names for all nine listed policy types (security, nat,
        decryption, authentication, dos, pbf, qos, sdwan, tunnel-inspect)
        confirmed directly from a live "type=Complete" query against Panorama.
      - An empty rulebase (zero rules of a given type) returns status="success"
        with code="7" and an empty <result/> — no <rules> child at all. The
        bulk-fetch logic explicitly checks for the <rules> node's presence
        before reading .entry rather than relying on incidental null-property
        behavior.
#>

param (
    [string]$PanoramaIP = $env:PaloAltoDeviceIP,
    [string]$ApiKey     = $env:PaloAltoAPIKey
)

# ── Validate required connection info ─────────────────────────────────────────

if ([string]::IsNullOrWhiteSpace($PanoramaIP)) {
    Write-Error "PanoramaIP is not set. Provide -PanoramaIP or set `$env:PaloAltoDeviceIP."
    exit 1
}

if ([string]::IsNullOrWhiteSpace($ApiKey)) {
    Write-Error "ApiKey is not set. Provide -ApiKey or set `$env:PaloAltoAPIKey."
    exit 1
}

# ── SSL bypass (version-aware, consistent with sync script) ──────────────────

$SkipCertSplat = @{}
if ($PSVersionTable.PSVersion.Major -ge 6) {
    $SkipCertSplat['SkipCertificateCheck'] = $true
} else {
    [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
}

$BaseUrl = "https://$PanoramaIP/api/"

# ── Helper: invoke PAN-OS XML API ─────────────────────────────────────────────

function Invoke-PanosApi {
    param (
        [hashtable]$QueryParams
    )

    $queryString = ($QueryParams.GetEnumerator() | ForEach-Object {
        "$($_.Key)=$([System.Uri]::EscapeDataString($_.Value))"
    }) -join "&"

    $uri = "$BaseUrl`?$queryString"

    try {
        $response = Invoke-WebRequest -Uri $uri -Method GET -UseBasicParsing -ErrorAction Stop @SkipCertSplat
        [xml]$response.Content
    }
    catch {
        Write-Warning "API call failed for URI: $uri`n  Error: $_"
        $null
    }
}

# ── Helper: get device group names ────────────────────────────────────────────
# Verified against sync-script logic — see comment inside function body.

function Get-DeviceGroupNames {
    $dgXml = Invoke-PanosApi -QueryParams @{
        type = "op"
        cmd  = "<show><devicegroups></devicegroups></show>"
        key  = $ApiKey
    }

    if ($null -eq $dgXml -or $dgXml.response.status -ne "success") {
        Write-Error "Failed to retrieve device group list from Panorama."
        exit 1
    }

    # Verified structure (reused from sync script): top-level device-group entries
    # each carry a 'name' attribute. The nested <devices><entry> nodes under each
    # device group are member-firewall sync status and are not relevant here.
    $dgEntries = @($dgXml.response.result.devicegroups.entry)
    $dgEntries | ForEach-Object { $_.name }
}

# ── Policy type selection ─────────────────────────────────────────────────────

$PolicyTypes = [ordered]@{
    "1"  = @{ Name = "Security";                  Prefix = "SEC";  Tag = "security" }
    "2"  = @{ Name = "NAT";                       Prefix = "NAT";  Tag = "nat" }
    "3"  = @{ Name = "Decryption";                Prefix = "DEC";  Tag = "decryption" }
    "4"  = @{ Name = "Authentication";             Prefix = "AUTH"; Tag = "authentication" }
    "5"  = @{ Name = "DoS Protection";             Prefix = "DOS";  Tag = "dos" }
    "6"  = @{ Name = "Policy-Based Forwarding";    Prefix = "PBF";  Tag = "pbf" }
    "7"  = @{ Name = "QoS";                        Prefix = "QOS";  Tag = "qos" }
    "8"  = @{ Name = "SD-WAN";                     Prefix = "SDW";  Tag = "sdwan" }
    "9"  = @{ Name = "Tunnel Inspection";          Prefix = "TUN";  Tag = "tunnel-inspect" }
    "10" = @{ Name = "Application Override";       Prefix = "APP";  Tag = "application-override" }
    "11" = @{ Name = "Network Packet Broker";      Prefix = "NPB";  Tag = "network-packet-broker" }
}

Write-Host ""
Write-Host "Select a policy type:" -ForegroundColor Cyan
foreach ($key in $PolicyTypes.Keys) {
    Write-Host ("  [{0}] {1}" -f $key, $PolicyTypes[$key].Name)
}
Write-Host "  [Q] Quit"
Write-Host ""

do {
    $selection = Read-Host "Enter the number of the policy type (or Q to quit)"
    if ($selection -eq "Q" -or $selection -eq "q") {
        Write-Host "Exiting." -ForegroundColor Cyan
        exit 0
    }
    $validSelection = (-not [string]::IsNullOrWhiteSpace($selection)) -and $PolicyTypes.Contains($selection)
    if (-not $validSelection) {
        Write-Host "Invalid selection. Please choose a number from the list above, or Q to quit." -ForegroundColor Red
    }
} until ($validSelection)

$policyTypeName = $PolicyTypes[$selection].Name
$policyPrefix   = $PolicyTypes[$selection].Prefix
$rulebaseTag    = $PolicyTypes[$selection].Tag

# ── Random string generation ──────────────────────────────────────────────────

$RandomStringLength = 8
$AllowedChars = [char[]]"ABCDEFGHJKMNPQRSTUVWXYZ23456789"  # excludes 0,O,1,I,L

function New-RandomString {
    param ([int]$Length, [char[]]$CharSet)
    -join (1..$Length | ForEach-Object { $CharSet | Get-Random })
}

# ── Bulk fetch of existing rule names (list-then-compare) ────────────────────
# Pulls every existing rule name for the selected policy type once — across all
# device groups (pre/post-rulebase) and shared — rather than issuing a targeted
# per-name lookup on every generation attempt. Candidates are then checked
# locally with no further API calls.
#
# Structure confirmed against live Panorama: $result.response.result.rules.entry
# returns multiple <entry name="..."> nodes directly under <rules>. An empty
# rulebase returns status="success" with no <rules> child at all — handled
# explicitly below rather than assumed.

function Get-ExistingRuleNames {
    param (
        [string]$RulebaseTag,
        [string[]]$DeviceGroups
    )

    $existingNames = [System.Collections.Generic.HashSet[string]]::new()

    foreach ($dg in $DeviceGroups) {
        foreach ($section in @("pre-rulebase", "post-rulebase")) {
            $xpath = "/config/devices/entry[@name='localhost.localdomain']/device-group/entry[@name='$dg']/$section/$RulebaseTag/rules"
            $result = Invoke-PanosApi -QueryParams @{
                type   = "config"
                action = "get"
                xpath  = $xpath
                key    = $ApiKey
            }
            # Confirmed: an empty rulebase returns status="success" with an
            # empty <result/> (no <rules> child at all) — not a non-success
            # status. Check explicitly for the <rules> node before reading
            # .entry, rather than relying on incidental null-property behavior.
            if ($null -ne $result -and $result.response.status -eq "success" -and $result.response.result.rules) {
                $entries = @($result.response.result.rules.entry)
                foreach ($e in $entries) {
                    if ($e.name) { [void]$existingNames.Add($e.name) }
                }
            }
        }
    }

    # Shared rulebase
    $sharedXpath = "/config/shared/$RulebaseTag/rules"
    $sharedResult = Invoke-PanosApi -QueryParams @{
        type   = "config"
        action = "get"
        xpath  = $sharedXpath
        key    = $ApiKey
    }
    if ($null -ne $sharedResult -and $sharedResult.response.status -eq "success" -and $sharedResult.response.result.rules) {
        $entries = @($sharedResult.response.result.rules.entry)
        foreach ($e in $entries) {
            if ($e.name) { [void]$existingNames.Add($e.name) }
        }
    }

    ,$existingNames
}

# ── Main generation loop ──────────────────────────────────────────────────────

Write-Host ""
Write-Host "Retrieving device groups from Panorama ($PanoramaIP)..." -ForegroundColor Cyan
$deviceGroups = Get-DeviceGroupNames
Write-Host "Found $($deviceGroups.Count) device group(s)." -ForegroundColor Cyan

Write-Host "Retrieving existing $rulebaseTag rule names (all device groups + shared)..." -ForegroundColor Cyan
$existingNames = Get-ExistingRuleNames -RulebaseTag $rulebaseTag -DeviceGroups $deviceGroups
Write-Host "Found $($existingNames.Count) existing rule name(s) of this type." -ForegroundColor Cyan

$fullPolicyName = $null
$maxAttempts = 50
$attempt = 0

do {
    $candidateString = New-RandomString -Length $RandomStringLength -CharSet $AllowedChars
    $candidateFullName = "$policyPrefix-POL-$candidateString"
    $attempt++

    if (-not $existingNames.Contains($candidateFullName)) {
        $fullPolicyName = $candidateFullName
    } else {
        Write-Host "Duplicate found: $candidateFullName already exists. Generating a new random string (attempt $attempt)..." -ForegroundColor Yellow
        if ($attempt -ge $maxAttempts) {
            Write-Error "Unable to generate a unique name after $maxAttempts attempts. Aborting."
            exit 1
        }
    }
} while (-not $fullPolicyName)

# ── Output result ──────────────────────────────────────────────────────────────

Write-Host ""
Write-Host "Generated unique identifier:" -ForegroundColor Green
Write-Host ("  Policy Type:      {0}" -f $policyTypeName)
Write-Host ("  Full Policy Name: {0}" -f $fullPolicyName)
