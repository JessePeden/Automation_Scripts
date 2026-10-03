<#
.SYNOPSIS
    Prompts for a Panorama-managed firewall and retrieves the physical
    media type for every port on that device, using a single wildcard
    XML API call forwarded through Panorama.

.DESCRIPTION
    Queries Panorama for "show devices connected" to build a list of
    managed firewalls (hostname + serial). Displays that list and prompts
    for a selection. Using the selected device's serial number as the
    Panorama API "target" parameter, issues a single op command with the
    wildcard filter "sys.s*.p*.phy" against that firewall, then splits
    the response into per-port blocks and extracts media/type from each.

.PARAMETER PanoramaIP
    Management IP or FQDN of Panorama.

.PARAMETER ApiKey
    Panorama XML API key. If omitted, you'll be prompted for a username
    and password and a key will be generated for you.

.EXAMPLE
    .\Get-PANInterfaceMedia_v21.ps1 -PanoramaIP panorama.company.local -ApiKey $key
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$PanoramaIP,

    [Parameter(Mandatory = $false)]
    [string]$ApiKey
)

$ErrorActionPreference = 'Stop'

# Suppress SSL/TLS certificate errors for self-signed certs (common on PAN-OS).
# Remove this block if your Panorama uses a trusted certificate.
$SkipCertSplat = @{}
if ($PSVersionTable.PSVersion.Major -ge 6) {
    # PowerShell 6+ (Core) — use the built-in parameter
    $SkipCertSplat['SkipCertificateCheck'] = $true
} else {
    # Windows PowerShell 5.1 — use the ServicePointManager callback
    [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
}

# --- Obtain an API key if one wasn't supplied ---
if (-not $ApiKey) {
    $Credential = Get-Credential -Message "Enter Panorama admin credentials for $PanoramaIP"
    $userEnc = [uri]::EscapeDataString($Credential.UserName)
    $passEnc = [uri]::EscapeDataString($Credential.GetNetworkCredential().Password)
    $keyUri  = "https://$PanoramaIP/api/?type=keygen&user=$userEnc&password=$passEnc"

    try {
        $keyResponse = Invoke-RestMethod -Uri $keyUri -Method Get @SkipCertSplat
        $ApiKey = $keyResponse.response.result.key
        if (-not $ApiKey) { throw "No key returned in response." }
    }
    catch {
        Write-Error "Failed to generate API key from $PanoramaIP : $_"
        return
    }
}

# --- Retrieve managed devices from Panorama ---
$devicesCmd = "<show><devices><connected></connected></devices></show>"
$devicesCmdEnc = [uri]::EscapeDataString($devicesCmd)
$devicesUri = "https://$PanoramaIP/api/?type=op&cmd=$devicesCmdEnc&key=$ApiKey"

try {
    $devicesResponse = Invoke-RestMethod -Uri $devicesUri -Method Get @SkipCertSplat
}
catch {
    Write-Error "Failed to query connected devices from $PanoramaIP : $_"
    return
}

$deviceEntries = $devicesResponse.response.result.devices.entry

if (-not $deviceEntries) {
    Write-Warning "No connected devices found on $PanoramaIP."
    return
}

$devices = @()
foreach ($entry in $deviceEntries) {
    $serial = $entry.name
    if ([string]::IsNullOrWhiteSpace($serial)) { continue }

    $hostname = if ($entry.hostname) { ([string]$entry.hostname).ToUpper() } else { $serial }
    $model    = [string]$entry.model

    $devices += [PSCustomObject]@{
        Hostname = $hostname
        Serial   = $serial
        Model    = $model
    }
}

if ($devices.Count -eq 0) {
    Write-Warning "No devices with valid serial numbers found."
    return
}

$devices = $devices | Sort-Object -Property Hostname

# --- Display list and prompt for selection ---
Write-Host ""
Write-Host "Managed devices on $PanoramaIP :"
for ($i = 0; $i -lt $devices.Count; $i++) {
    Write-Host ("  [{0}] {1}  ({2}, {3})" -f ($i + 1), $devices[$i].Hostname, $devices[$i].Model, $devices[$i].Serial)
}
Write-Host "  [Q] Quit"
Write-Host ""

$selection = $null
do {
    $selectionInput = Read-Host "Select a device by number (or Q to quit)"
    if ($selectionInput -eq "Q" -or $selectionInput -eq "q") {
        Write-Host "Exiting." -ForegroundColor Cyan
        exit 0
    }

    $parsed = 0
    if ([int]::TryParse($selectionInput, [ref]$parsed) -and $parsed -ge 1 -and $parsed -le $devices.Count) {
        $selection = $parsed
    }
    else {
        Write-Host "Invalid selection. Enter a number between 1 and $($devices.Count), or Q to quit." -ForegroundColor Red
    }
} until ($selection)

$targetDevice = $devices[$selection - 1]
Write-Host ""
Write-Host "Querying $($targetDevice.Hostname) ($($targetDevice.Serial))..."
Write-Host ""

# --- Query the selected device via Panorama target forwarding ---
$phyCmd = "<show><system><state><filter-pretty>sys.s*.p*.phy</filter-pretty></state></system></show>"
$phyCmdEnc = [uri]::EscapeDataString($phyCmd)
$phyUri = "https://$PanoramaIP/api/?type=op&cmd=$phyCmdEnc&target=$($targetDevice.Serial)&key=$ApiKey"

try {
    $phyResponse = Invoke-RestMethod -Uri $phyUri -Method Get @SkipCertSplat
    $rawText = $phyResponse.response.result.InnerText
}
catch {
    Write-Error "Failed to query $($targetDevice.Hostname) via $PanoramaIP : $_"
    return
}

# --- Split raw text into per-port blocks ---
# Each block starts at a line like: sys.s1.p9.phy: {
$blockPattern = '(?m)^sys\.s(\d+)\.p(\d+)\.phy:'
$blockMatches = [regex]::Matches($rawText, $blockPattern)

if ($blockMatches.Count -eq 0) {
    Write-Warning "No port data returned for $($targetDevice.Hostname)."
    return
}

$results = @()

for ($i = 0; $i -lt $blockMatches.Count; $i++) {
    $slot = $blockMatches[$i].Groups[1].Value
    $port = $blockMatches[$i].Groups[2].Value

    $blockStart = $blockMatches[$i].Index
    $blockEnd = if ($i -lt $blockMatches.Count - 1) { $blockMatches[$i + 1].Index } else { $rawText.Length }
    $block = $rawText.Substring($blockStart, $blockEnd - $blockStart)

    $entry = [ordered]@{
        Interface = "ethernet$slot/$port"
        Slot      = [int]$slot
        Port      = [int]$port
        Media     = $null
        Type      = $null
        Connector = $null
    }

    if ($block -match '(?m)^\s*media:\s*([^\s,]+)') {
        $entry.Media = $Matches[1].TrimEnd(',')
    }
    if ($block -match '(?m)^\s*type:\s*([^\s,]+)') {
        $entry.Type = $Matches[1].TrimEnd(',')
    }

    $results += [PSCustomObject]$entry
}

# --- Query port capability (speed list) to infer connector type ---
# CAT5 ports are already unambiguous from the media field. For everything
# else, the max supported speed distinguishes SFP/SFP+/SFP28 (<= 25Gb/s)
# from QSFP+/QSFP28 (40Gb/s or 100Gb/s). 1G/10G overlap between SFP+ and
# SFP28 ports means those two can't be split further by speed alone, so
# they're reported as a combined label.
$capCmd = "<show><system><state><filter-pretty>sys.s*.p*.capability</filter-pretty></state></system></show>"
$capCmdEnc = [uri]::EscapeDataString($capCmd)
$capUri = "https://$PanoramaIP/api/?type=op&cmd=$capCmdEnc&target=$($targetDevice.Serial)&key=$ApiKey"

try {
    $capResponse = Invoke-RestMethod -Uri $capUri -Method Get @SkipCertSplat
    $capRawText = $capResponse.response.result.InnerText
}
catch {
    Write-Warning "Failed to query port capabilities from $($targetDevice.Hostname) via $PanoramaIP : $_. Connector type will be left blank."
    $capRawText = $null
}

$maxSpeedByPort = @{}

if ($capRawText) {
    $capBlockPattern = '(?m)^sys\.s(\d+)\.p(\d+)\.capability:'
    $capBlockMatches = [regex]::Matches($capRawText, $capBlockPattern)

    for ($i = 0; $i -lt $capBlockMatches.Count; $i++) {
        $capSlot = $capBlockMatches[$i].Groups[1].Value
        $capPort = $capBlockMatches[$i].Groups[2].Value

        $capBlockStart = $capBlockMatches[$i].Index
        $capBlockEnd = if ($i -lt $capBlockMatches.Count - 1) { $capBlockMatches[$i + 1].Index } else { $capRawText.Length }
        $capBlock = $capRawText.Substring($capBlockStart, $capBlockEnd - $capBlockStart)

        $speedMatches = [regex]::Matches($capBlock, '(\d+(?:\.\d+)?)(Mb|Gb)/s-full')
        $maxMbps = 0
        foreach ($sm in $speedMatches) {
            $value = [double]$sm.Groups[1].Value
            $mbps = if ($sm.Groups[2].Value -eq 'Gb') { $value * 1000 } else { $value }
            if ($mbps -gt $maxMbps) { $maxMbps = $mbps }
        }

        $maxSpeedByPort["$capSlot.$capPort"] = $maxMbps
    }
}

foreach ($row in $results) {
    if ($row.Media -eq 'CAT5') {
        $row.Connector = 'Copper/RJ-45'
        continue
    }

    $key = "$($row.Slot).$($row.Port)"
    if ($maxSpeedByPort.ContainsKey($key) -and $maxSpeedByPort[$key] -gt 0) {
        $maxMbps = $maxSpeedByPort[$key]
        if ($maxMbps -ge 40000) {
            $row.Connector = 'QSFP+/QSFP28'
        }
        else {
            $row.Connector = 'SFP/SFP+/SFP28'
        }
    }
    else {
        $row.Connector = 'Unknown'
    }
}

# --- Identify QSFP breakout parent/child relationships ---
# Parent ports are physical QSFP cages (Media starts with "QSFP"). The
# breakout-capable lanes are the 4x(parent count) ports immediately below
# the lowest-numbered parent port. This is positional, not based on the
# per-port "identifier" field, since that field is only present when a
# transceiver is actually installed and would undercount empty lanes.
$parentPorts = $results | Where-Object { $_.Media -like 'QSFP*' } | Sort-Object Port
$childToParent = @{}
$parentToChildren = @{}

if ($parentPorts.Count -gt 0) {
    $parentCount = $parentPorts.Count
    $minParentPort = $parentPorts[0].Port
    $laneCount = 4 * $parentCount
    $laneStart = $minParentPort - $laneCount
    $laneEnd = $minParentPort - 1

    $laneCandidates = $results |
        Where-Object { $_.Port -ge $laneStart -and $_.Port -le $laneEnd -and $_.Media -ne 'CAT5' } |
        Sort-Object Port

    if ($laneCandidates.Count -eq $laneCount) {
        for ($g = 0; $g -lt $parentCount; $g++) {
            $groupLanes = $laneCandidates[($g * 4)..($g * 4 + 3)]
            $parentPort = $parentPorts[$g].Port
            $parentToChildren[$parentPort] = $groupLanes
            foreach ($lane in $groupLanes) { $childToParent[$lane.Port] = $parentPort }
        }
    }
    else {
        Write-Warning "Expected $laneCount breakout-lane ports below port $minParentPort but found $($laneCandidates.Count). Skipping breakout grouping."
    }
}

# --- Output, sorted numerically by port, with breakout children indented under their parent ---
Write-Host "Media types for $($targetDevice.Hostname):"
Write-Host ""
Write-Host ("{0,-14} {1}" -f "Interface", "Connector")
Write-Host ("{0,-14} {1}" -f "---------", "---------")
foreach ($row in ($results | Sort-Object Slot, Port)) {
    if ($childToParent.ContainsKey($row.Port)) { continue }

    Write-Host ("{0,-14} {1}" -f $row.Interface, $row.Connector)

    if ($parentToChildren.ContainsKey($row.Port)) {
        $children = $parentToChildren[$row.Port] | Sort-Object Port
        Write-Host "  |"
        foreach ($child in $children) {
            Write-Host ("  +-- {0,-10} {1}  (breakout of $($row.Interface))" -f $child.Interface, $child.Connector)
        }
    }
}
