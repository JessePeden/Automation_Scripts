#Requires -Version 5.1
# Version: 66
<#
.SYNOPSIS
    Queries all Panorama-managed firewalls for HA running-sync state and Panorama sync status.

.DESCRIPTION
    Uses the PAN-OS XML API to:
      1. Retrieve all connected managed devices from Panorama (serial/hostname map).
      2. Query 'show devicegroups' on Panorama to get shared policy sync status per device.
      3. Query 'show templates' on Panorama to get template sync status per device.
      4. Query 'show high-availability state' on Panorama itself to check Panorama HA sync.
      5. Issue 'show high-availability state' to each managed device via the Panorama proxy.
      6. Write all results sorted by hostname to an XLSX file.

    Panorama itself appears as a row in the output with 'Not Applicable' in the
    Device-Group Status and Template Status columns.

.PARAMETER PanoramaIP
    IP address or FQDN of the Panorama instance.
    Defaults to the environment variable $env:PaloAltoDeviceIP if not specified.

.PARAMETER ApiKey
    PAN-OS API key with at least read access.
    Defaults to the environment variable $env:PaloAltoAPIKey if not specified.

.PARAMETER OutputFolder
    Folder path where the XLSX file will be saved.
    Defaults to the current directory.

.PARAMETER OutputFileName
    Name of the XLSX file (including .xlsx extension).
    Defaults to <PanoramaHostname>_Sync_Status.xlsx, derived at runtime from Panorama system info.

.EXAMPLE
    .\Get-PANSyncState.ps1

.EXAMPLE
    .\Get-PANSyncState.ps1 -PanoramaIP 192.168.1.1 -ApiKey "LUFRPT1x..." -OutputFolder C:\Reports -OutputFileName MyReport.xlsx

.NOTES
    Requires the ImportExcel module: Install-Module -Name ImportExcel -Scope CurrentUser
    If -OutputFileName is not specified, the script queries Panorama system info to obtain
    the hostname and names the file <hostname>_Sync_Status.xlsx.
    If -OutputFolder is not specified, the file is saved in the current directory.
#>

[CmdletBinding()]
param (
    [string]$PanoramaIP = $env:PaloAltoDeviceIP,
    [string]$ApiKey     = $env:PaloAltoAPIKey,
    [string]$OutputFolder   = "",
    [string]$OutputFileName = ""
)

# ── Preflight checks ──────────────────────────────────────────────────────────

if ([string]::IsNullOrWhiteSpace($PanoramaIP)) {
    Write-Error "PanoramaIP is not set. Provide -PanoramaIP or set `$env:PaloAltoDeviceIP."
    exit 1
}

if ([string]::IsNullOrWhiteSpace($ApiKey)) {
    Write-Error "ApiKey is not set. Provide -ApiKey or set `$env:PaloAltoAPIKey."
    exit 1
}

# Suppress SSL/TLS certificate errors for self-signed certs (common on PAN-OS).
# Remove this block if your Panorama uses a trusted certificate.
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

# ── Resolve output folder and filename ────────────────────────────────────────────

if ([string]::IsNullOrWhiteSpace($OutputFolder)) {
    $OutputFolder = (Get-Location).Path
}

if ([string]::IsNullOrWhiteSpace($OutputFileName)) {
    $sysInfoXml = Invoke-PanosApi -QueryParams @{
        type = "op"
        cmd  = "<show><system><info></info></system></show>"
        key  = $ApiKey
    }

    $panoramaHostname = if ($null -ne $sysInfoXml -and $sysInfoXml.response.status -eq "success") {
        ([string]$sysInfoXml.response.result.system.hostname).ToUpper()
    } else {
        "PANORAMA"
    }

$panoramaSerial = if ($null -ne $sysInfoXml -and $sysInfoXml.response.status -eq "success") {
        ([string]$sysInfoXml.response.result.system.serial)
    } else {
        "Not Applicable"
    }

$panoramaModel = if ($null -ne $sysInfoXml -and $sysInfoXml.response.status -eq "success") {
        ([string]$sysInfoXml.response.result.system.model)
    } else {
        ""
    }

    $OutputFileName = "${panoramaHostname}_Sync_Status.xlsx"
}

$OutputXlsx = Join-Path $OutputFolder $OutputFileName
Write-Host "Output file will be: $OutputXlsx`n" -ForegroundColor Cyan

# ── Step 1: Get connected managed devices ─────────────────────────────────────

Write-Host "Retrieving all managed devices from Panorama ($PanoramaIP)..." -ForegroundColor Cyan

$devicesXml = Invoke-PanosApi -QueryParams @{
    type = "op"
    cmd  = "<show><devices><all></all></devices></show>"
    key  = $ApiKey
}

if ($null -eq $devicesXml -or $devicesXml.response.status -ne "success") {
    Write-Error "Failed to retrieve device list from Panorama."
    exit 1
}

$deviceEntries = @($devicesXml.response.result.devices.entry)

if ($deviceEntries.Count -eq 0) {
    Write-Warning "No managed devices found. Exiting."
    exit 0
}

Write-Host "Found $($deviceEntries.Count) managed device(s).`n" -ForegroundColor Green

# Build serial -> hostname, serial -> connected, and serial -> model lookups
$hostnameMap  = @{}
$connectedMap = @{}
$modelMap     = @{}
foreach ($d in $deviceEntries) {
    $hostnameMap[$d.name]  = if ($d.hostname) { ([string]$d.hostname).ToUpper() } else { ([string]$d.name).ToUpper() }
    $connectedMap[$d.name] = if ($d.connected -eq "yes") { "Connected" } else { "Disconnected" }
    $modelMap[$d.name]     = [string]$d.model
}

# ── Step 2: Query device group sync status from Panorama ──────────────────────

Write-Host "Querying device group sync status from Panorama..." -ForegroundColor Cyan

$dgXml = Invoke-PanosApi -QueryParams @{
    type = "op"
    cmd  = "<show><devicegroups></devicegroups></show>"
    key  = $ApiKey
}

# Build a serial -> "DeviceGroup:Status" string (a device can only belong to one device group)
$dgStatusMap = @{}

if ($null -ne $dgXml -and $dgXml.response.status -eq "success") {
    $dgEntries = @($dgXml.response.result.devicegroups.entry)
    foreach ($dg in $dgEntries) {
        $dgName    = $dg.name
        $dgDevices = @($dg.devices.entry)
        foreach ($dev in $dgDevices) {
            $serial = $dev.name
            if ([string]::IsNullOrWhiteSpace($serial)) { continue }
            $status = if ($dev.'shared-policy-status') { $dev.'shared-policy-status' } else { "Unknown" }
            $dgStatusMap[$serial] = "${dgName}:${status}"
        }
    }
} else {
    Write-Warning "Failed to retrieve device group info — DeviceGroupStatus will show as error."
}

# ── Step 3: Query template sync status from Panorama ──────────────────────────

Write-Host "Querying template sync status from Panorama..." -ForegroundColor Cyan

$tmplXml = Invoke-PanosApi -QueryParams @{
    type = "op"
    cmd  = "<show><templates></templates></show>"
    key  = $ApiKey
}

# Build a serial -> "TemplateStack:Status" string (a device can only have one template stack applied)
$tmplStatusMap = @{}

if ($null -ne $tmplXml -and $tmplXml.response.status -eq "success") {
    $tmplEntries = @($tmplXml.response.result.templates.entry)
    foreach ($tmpl in $tmplEntries) {
        $tmplName    = $tmpl.name
        $tmplDevices = @($tmpl.devices.entry)
        foreach ($dev in $tmplDevices) {
            $serial = $dev.name
            if ([string]::IsNullOrWhiteSpace($serial)) { continue }
            $status = if ($dev.'template-status') { $dev.'template-status' } else { "Unknown" }
            $tmplStatusMap[$serial] = "${tmplName}:${status}"
        }
    }
} else {
    Write-Warning "Failed to retrieve template info — TemplateStatus will show as error."
}

# ── Helper: extract HA compat and connection status fields ───────────────────────

$haFieldMapping = @{
    'build-compat'     = 'HA PAN-OS Version'
    'app-compat'       = 'HA App Version'
    'iot-compat'       = 'HA Device Dictionary'
    'av-compat'        = 'HA Antivirus Version'
    'threat-compat'    = 'HA Threat Version'
    'vpnclient-compat' = 'HA Legacy VPN Client Version'
    'gpclient-compat'  = 'HA GlobalProtect Version'
    'DLP'              = 'HA DLP Version'
    'Cloud-Services'   = 'HA Cloud Services'
    'cloudconnector'   = 'HA Cloud Connector'
    'OC'               = 'HA OpenConfig'
    'VMS'              = 'HA VM Series'
}

$haConnMapping = @{
    'conn-ha1'        = 'HA1 Connection Status'
    'conn-ha1-backup' = 'HA1 Backup Connection Status'
    'conn-ha2'        = 'HA2 Connection Status'
    'conn-ha2-backup' = 'HA2 Backup Connection Status'
}

function Get-HaCompatFields {
    param ($HaNode)
    $fields = [ordered]@{}
    if ($null -eq $HaNode) { return $fields }

    # Extract compat fields from local-info — includes any field in the mapping or ending in '-compat'
    $localInfoNodes = $HaNode.SelectNodes('local-info/*')
    if ($null -ne $localInfoNodes) {
        foreach ($node in $localInfoNodes) {
            $fieldName = $node.Name
            if ($haFieldMapping.ContainsKey($fieldName)) {
                $header = if ($haFieldMapping.ContainsKey($fieldName)) {
                    $haFieldMapping[$fieldName]
                } else {
                    $fieldParts = $fieldName -split '-'
                    $fieldTitle = ($fieldParts | ForEach-Object { $_.Substring(0,1).ToUpper() + $_.Substring(1) }) -join '-'
                    "HA $fieldTitle"
                }
                $fields[$header] = [string]$node.InnerText
            }
        }
    }

    # Extract conn-status from conn-ha* nodes in peer-info using XPath
    $connNodes = $HaNode.SelectNodes("peer-info/*[starts-with(name(), 'conn-ha')]")
    if ($null -ne $connNodes) {
        foreach ($connNode in $connNodes) {
            $connStatusNode = $connNode.SelectSingleNode('conn-status')
            if ($null -ne $connStatusNode) {
                $header = if ($haConnMapping.ContainsKey($connNode.Name)) {
                    $haConnMapping[$connNode.Name]
                } else {
                    $connParts = ($connNode.Name -replace '^conn-', '') -split '-'
                    $connTitle = ($connParts | ForEach-Object { $_.Substring(0,1).ToUpper() + $_.Substring(1) }) -join '-'
                    "HA $connTitle Conn-Status"
                }
                $fields[$header] = [string]$connStatusNode.InnerText
            }
        }
    }

    return $fields
}

# ── Helper: query uncommitted changes and return distinct admin list ──────────────────────────────

function Get-UncommittedChanges {
    param (
        [hashtable]$QueryParams,
        [scriptblock]$ApiFunction
    )
    $xml = & $ApiFunction $QueryParams
    if ($null -eq $xml -or $xml.response.status -ne "success") { return "Unknown" }
    $entries = $xml.response.result.journal.entry
    if ($null -eq $entries) { return "None" }
    $entries = @($entries)
    if ($entries.Count -eq 0) { return "None" }
    $admins = ($entries | ForEach-Object { [string]$_.owner } | Where-Object { $_ -ne '__dlp' -and $_ -ne '' } | Select-Object -Unique | Sort-Object) -join ", "
    if ([string]::IsNullOrWhiteSpace($admins)) { return "None" }
    return $admins
}

# ── Step 4: Query Panorama’s own HA state and peer hostname ─────────────────────────────────────────────────────────────────

Write-Host "Querying Panorama HA state ($panoramaHostname)..." -ForegroundColor Cyan

$panoramaHaXml = Invoke-PanosApi -QueryParams @{
    type = "op"
    cmd  = "<show><high-availability><state></state></high-availability></show>"
    key  = $ApiKey
}

$panoramaRunningSync = "error"
$panoramaHaEnabled   = "no"

if ($null -ne $panoramaHaXml -and $panoramaHaXml.response.status -eq "success") {
    $panoramaHaEnabled = $panoramaHaXml.response.result.enabled
    if ($panoramaHaEnabled -eq "no") {
        $panoramaRunningSync = "HA not enabled"
    } else {
        $panoramaRunningSync = if ($panoramaHaXml.response.result."running-sync") { $panoramaHaXml.response.result."running-sync" } else { "N/A" }
        $panoramaHaState     = if ($panoramaHaXml.response.result."local-info".state) { $panoramaHaXml.response.result."local-info".state } else { "N/A" }
    }
} else {
    Write-Warning "  API error or no response for Panorama HA state — marking as error."
}

if (-not $panoramaHaState) { $panoramaHaState = "HA not enabled" }

$panoramaHaCompatFields = Get-HaCompatFields -HaNode $panoramaHaXml.response.result

# Query uncommitted changes on primary Panorama
$panoramaUncommittedAdmins = Get-UncommittedChanges -QueryParams @{
    type = "op"
    cmd  = "<show><config><list><changes></changes></list></config></show>"
    key  = $ApiKey
} -ApiFunction { param($p) Invoke-PanosApi -QueryParams $p }


# ── Step 5: Query HA state on each managed device ─────────────────────────────

$results             = [System.Collections.Generic.List[PSCustomObject]]::new()
$uncommittedComments = @{}

# Add the primary Panorama row
$panoramaUncommittedDisplay = if ($panoramaUncommittedAdmins -eq 'None') { 'None' } else { 'Yes — see comment' }
$panoramaProps = [ordered]@{
    'Hostname'            = [string]$panoramaHostname
    'Model'               = [string]$panoramaModel
    'Serial Number'       = [string]$panoramaSerial
    'Connection Status'   = [string]'Not Applicable'
    'Device-Group Status' = [string]'Not Applicable'
    'Template Status'     = [string]'Not Applicable'
    'HA Running-Sync'     = [string]$panoramaRunningSync
    'HA State'            = [string]$panoramaHaState
    'Uncommitted Changes' = [string]$panoramaUncommittedDisplay
}
foreach ($key in $panoramaHaCompatFields.Keys) { $panoramaProps[$key] = $panoramaHaCompatFields[$key] }
$results.Add([PSCustomObject]$panoramaProps)
$uncommittedComments[$results.Count - 1] = $panoramaUncommittedAdmins

# Only query the peer if HA is enabled on the primary Panorama
if ($panoramaHaEnabled -ne "no") {

    # Query peer.cfg.hostname to get the peer Panorama hostname
    $panoramaPeerHostnameXml = Invoke-PanosApi -QueryParams @{
        type = "op"
        cmd  = "<show><system><state><filter>peer.cfg.hostname</filter></state></system></show>"
        key  = $ApiKey
    }

    $panoramaPeerHostname = $null
    if ($null -ne $panoramaPeerHostnameXml -and $panoramaPeerHostnameXml.response.status -eq "success") {
        $stateText = $panoramaPeerHostnameXml.response.result.InnerText.Trim()
        if ($stateText -match "peer\.cfg\.hostname:\s*(\S+)") {
            $panoramaPeerHostname = $Matches[1].Trim().ToUpper()
        }
    }

    if ([string]::IsNullOrWhiteSpace($panoramaPeerHostname)) {
        Write-Warning "  Could not resolve peer hostname — skipping peer queries."
    } else {
        Write-Host "  Peer hostname : $panoramaPeerHostname"

        # Query the peer directly for its serial number and HA sync status
        $peerBaseUrl = "https://$panoramaPeerHostname/api/"

        function Invoke-PeerPanosApi {
            param (
                [hashtable]$QueryParams
            )

            $queryString = ($QueryParams.GetEnumerator() | ForEach-Object {
                "$($_.Key)=$([System.Uri]::EscapeDataString($_.Value))"
            }) -join "&"

            $uri = "$peerBaseUrl`?$queryString"

            try {
                $response = Invoke-WebRequest -Uri $uri -Method GET -UseBasicParsing -ErrorAction Stop @SkipCertSplat
                [xml]$response.Content
            }
            catch {
                Write-Warning "API call failed for URI: $uri`n  Error: $_"
                $null
            }
        }

        # Get peer serial number
        $peerSysInfoXml = Invoke-PeerPanosApi -QueryParams @{
            type = "op"
            cmd  = "<show><system><info></info></system></show>"
            key  = $ApiKey
        }

        $peerSerial = if ($null -ne $peerSysInfoXml -and $peerSysInfoXml.response.status -eq "success") {
            ([string]$peerSysInfoXml.response.result.system.serial)
        } else { "" }

        $peerModel = if ($null -ne $peerSysInfoXml -and $peerSysInfoXml.response.status -eq "success") {
            ([string]$peerSysInfoXml.response.result.system.model)
        } else { "" }

        # Get peer HA sync status
        $peerHaXml = Invoke-PeerPanosApi -QueryParams @{
            type = "op"
            cmd  = "<show><high-availability><state></state></high-availability></show>"
            key  = $ApiKey
        }

        $peerRunningSync = "error"
        if ($null -ne $peerHaXml -and $peerHaXml.response.status -eq "success") {
            $peerHaEnabled = $peerHaXml.response.result.enabled
            if ($peerHaEnabled -eq "no") {
                $peerRunningSync = "HA not enabled"
            } else {
                $peerRunningSync = if ($peerHaXml.response.result."running-sync") { $peerHaXml.response.result."running-sync" } else { "N/A" }
                $peerHaState     = if ($peerHaXml.response.result."local-info".state) { $peerHaXml.response.result."local-info".state } else { "N/A" }
            }
        } else {
            Write-Warning "  API error or no response for peer Panorama HA state — marking as error."
        }

        if (-not $peerHaState) { $peerHaState = "HA not enabled" }

        $peerHaCompatFields = Get-HaCompatFields -HaNode $peerHaXml.response.result

        # Query uncommitted changes on peer Panorama
        $peerUncommittedAdmins = Get-UncommittedChanges -QueryParams @{
            type = "op"
            cmd  = "<show><config><list><changes></changes></list></config></show>"
            key  = $ApiKey
        } -ApiFunction { param($p) Invoke-PeerPanosApi -QueryParams $p }


        # Add the peer Panorama row
        $peerUncommittedDisplay = if ($peerUncommittedAdmins -eq 'None') { 'None' } else { 'Yes — see comment' }
        $peerProps = [ordered]@{
            'Hostname'            = [string]$panoramaPeerHostname
            'Model'               = [string]$peerModel
            'Serial Number'       = [string]$peerSerial
            'Connection Status'   = [string]'Not Applicable'
            'Device-Group Status' = [string]'Not Applicable'
            'Template Status'     = [string]'Not Applicable'
            'HA Running-Sync'     = [string]$peerRunningSync
            'HA State'            = [string]$peerHaState
            'Uncommitted Changes' = [string]$peerUncommittedDisplay
        }
        foreach ($key in $peerHaCompatFields.Keys) { $peerProps[$key] = $peerHaCompatFields[$key] }
        $results.Add([PSCustomObject]$peerProps)
        $uncommittedComments[$results.Count - 1] = $peerUncommittedAdmins
    }
} else {
    Write-Host "  Panorama HA is not enabled — skipping peer queries.`n" -ForegroundColor Yellow
}
$i = 0

foreach ($device in $deviceEntries) {
    $i++
    $serial   = $device.name
    $hostname = $hostnameMap[$serial]

    $connectedStatus = $connectedMap[$serial]
    $model           = $modelMap[$serial]
    $dgStatus   = if ($dgStatusMap.ContainsKey($serial))   { $dgStatusMap[$serial] } else { "Not in any device group" }
    $tmplStatus = if ($tmplStatusMap.ContainsKey($serial))  { $tmplStatusMap[$serial] } else { "Not in any template" }

    Write-Host "[$i/$($deviceEntries.Count)] Querying $hostname ($serial)"

    $runningSync = "Unknown"
    $haState     = "Unknown"

    if ($connectedStatus -eq "Connected") {
        $haXml = Invoke-PanosApi -QueryParams @{
            type   = "op"
            cmd    = "<show><high-availability><state></state></high-availability></show>"
            target = $serial
            key    = $ApiKey
        }

        $runningSync = "error"

        if ($null -ne $haXml -and $haXml.response.status -eq "success") {
            $haEnabled = $haXml.response.result.enabled
            if ($haEnabled -eq "no") {
                $runningSync = "HA not enabled"
            } else {
                $group = $haXml.response.result.group
                $runningSync = if ($group."running-sync") { $group."running-sync" } else { "N/A" }
                $haState     = if ($group."local-info".state) { $group."local-info".state } else { "N/A" }
            }
        } else {
            Write-Warning "  API error or no response for $hostname — marking as error."
        }

        if (-not $haState -or $haState -eq "Unknown") { $haState = "HA not enabled" }
        $haCompatFields = Get-HaCompatFields -HaNode $haXml.response.result.group
    } else {
        $haCompatFields = [ordered]@{}
    }

    # Query uncommitted changes on managed firewall via Panorama proxy
    $fwUncommittedAdmins = "Unknown"
    if ($connectedStatus -eq "Connected") {
        $fwUncommittedAdmins = Get-UncommittedChanges -QueryParams @{
            type   = "op"
            cmd    = "<show><config><list><changes></changes></list></config></show>"
            target = $serial
            key    = $ApiKey
        } -ApiFunction { param($p) Invoke-PanosApi -QueryParams $p }
    }


    $fwUncommittedDisplay = if ($fwUncommittedAdmins -eq 'None') { 'None' } else { 'Yes — see comment' }
    $fwProps = [ordered]@{
        'Hostname'            = [string]$hostname
        'Model'               = [string]$model
        'Serial Number'        = [string]$serial
        'Connection Status'    = [string]$connectedStatus
        'Device-Group Status'  = [string]$dgStatus
        'Template Status'      = [string]$tmplStatus
        'HA Running-Sync'      = [string]$runningSync
        'HA State'             = [string]$haState
        'Uncommitted Changes'  = [string]$fwUncommittedDisplay
    }
    foreach ($key in $haCompatFields.Keys) { $fwProps[$key] = $haCompatFields[$key] }
    $results.Add([PSCustomObject]$fwProps)
    $uncommittedComments[$results.Count - 1] = $fwUncommittedAdmins
}

# Normalize all result rows to the same property set so Export-Excel and
# Format-Table see a consistent schema regardless of which HA fields each
# device returned. Rows missing a property get Not Applicable.
$allPropertyNames = $results | ForEach-Object { $_.PSObject.Properties.Name } | Select-Object -Unique

# Apply explicit preferred column order
$preferredOrder = @(
    'Hostname', 'Model', 'Serial Number', 'Connection Status',
    'Device-Group Status', 'Template Status', 'Uncommitted Changes', 'HA Running-Sync', 'HA State',
    'HA DLP Version', 'HA PAN-OS Version', 'HA App Version',
    'HA Antivirus Version', 'HA Threat Version', 'HA Legacy VPN Client Version',
    'HA GlobalProtect Version', 'HA Cloud Services', 'HA Cloud Connector',
    'HA Device Dictionary', 'HA OpenConfig', 'HA VM Series',
    'HA1 Connection Status', 'HA1 Backup Connection Status',
    'HA2 Connection Status', 'HA2 Backup Connection Status'
)
# Start with all preferred columns unconditionally, then append any unknown discovered columns
$allPropertyNames = @(
    $preferredOrder
    $allPropertyNames | Where-Object { $preferredOrder -notcontains $_ }
)

foreach ($row in $results) {
    foreach ($propName in $allPropertyNames) {
        if (-not $row.PSObject.Properties[$propName]) {
            $row | Add-Member -MemberType NoteProperty -Name $propName -Value "Not Applicable"
        }
    }
}

# ── Step 6: Export to XLSX ─────────────────────────────────────────────────────────────────────────────

if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Error "ImportExcel module is not installed. Run: Install-Module -Name ImportExcel -Scope CurrentUser"
    exit 1
}

$panoramaHostnames = @($panoramaHostname)
if ($panoramaPeerHostname) { $panoramaHostnames += $panoramaPeerHostname }

$panoramaRows = @($results | Where-Object { $panoramaHostnames -contains $_.'Hostname' })
$firewallRows = @($results | Where-Object { $panoramaHostnames -notcontains $_.'Hostname' } | Sort-Object -Property 'Hostname')
$sorted = @($panoramaRows + $firewallRows)

# Build hostname -> result index map for uncommitted changes comment lookup
$hostnameToResultIndex = @{}
for ($idx = 0; $idx -lt $results.Count; $idx++) {
    $hostnameToResultIndex[$results[$idx].Hostname] = $idx
}

try {
    $excel = $sorted | Select-Object -Property $allPropertyNames | Export-Excel -Path $OutputXlsx `
        -WorksheetName "PAN Sync State" `
        -ClearSheet `
        -FreezeTopRowFirstColumn `
        -BoldTopRow `
        -TableName "PANSyncState" `
        -TableStyle Medium2 `
        -PassThru `
        -ErrorAction Stop

    $ws      = $excel.Workbook.Worksheets["PAN Sync State"]
    $lastRow = $ws.Dimension.End.Row
    $lastCol = $ws.Dimension.End.Column

    # Re-write serial number column (B) as text to preserve leading zeros.
    # Values are sourced from $sorted (not from the cells) to avoid reading
    # already-coerced numeric values back from EPPlus.
    for ($row = 0; $row -lt $sorted.Count; $row++) {
        $rowNum = $row + 2
        $cell = $ws.Cells["C$rowNum"]
        $cell.Style.Numberformat.Format = "@"
        $cell.Value = [string]$sorted[$row].'Serial Number'
    }

    # Add cell comments for rows with uncommitted changes
    # Find the Uncommitted Changes column index once
    $uncommittedColIndex = $null
    for ($col = 1; $col -le $lastCol; $col++) {
        if ([string]$ws.Cells[1, $col].Value -eq 'Uncommitted Changes') {
            $uncommittedColIndex = $col
            break
        }
    }

    if ($null -ne $uncommittedColIndex) {
        for ($row = 0; $row -lt $sorted.Count; $row++) {
            $cellValue = [string]$sorted[$row].'Uncommitted Changes'
            if ($cellValue -eq 'Yes — see comment') {
                $resultIdx = $hostnameToResultIndex[$sorted[$row].Hostname]
                $admins    = $uncommittedComments[$resultIdx]
                if (-not [string]::IsNullOrWhiteSpace($admins)) {
                    $rowNum  = $row + 2
                    $colLetter = [char](64 + $uncommittedColIndex)
                    $cell    = $ws.Cells["${colLetter}${rowNum}"]
                    $comment = $cell.AddComment($admins, "Get-PANSyncState")
                    $comment.AutoFit = $true
                }
            }
        }
    }

    # Manual column sizing — iterate every column and set width to the longest cell value
    for ($col = 1; $col -le $lastCol; $col++) {
        $maxLen = 0
        for ($row = 1; $row -le $lastRow; $row++) {
            $cellValue = [string]$ws.Cells[$row, $col].Value
            if ($cellValue.Length -gt $maxLen) {
                $maxLen = $cellValue.Length
            }
        }
        $ws.Column($col).Width = [Math]::Max($maxLen + 2, 10)
    }

    # Disable auto-filter on the table
    $ws.Tables["PANSyncState"].ShowFilter = $false

    # Apply all borders to the entire data range
    $dataRange = $ws.Cells[1, 1, $lastRow, $lastCol]
    $dataRange.Style.Border.Top.Style    = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
    $dataRange.Style.Border.Bottom.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
    $dataRange.Style.Border.Left.Style   = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
    $dataRange.Style.Border.Right.Style  = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin

    # Column index map (1-based): A=Hostname, B=Model, C=Serial Number, D=Connection Status, E=Device-Group Status, F=Template Status, G=Uncommitted Changes, H=HA Running-Sync, I=HA State
    $colConn    = "D"
    $colDG      = "E"
    $colTmpl    = "F"
    $colHA      = "H"
    $colHAState = "I"

    $green     = [System.Drawing.Color]::FromArgb(198, 239, 206)   # Excel "Good" green
    $red       = [System.Drawing.Color]::FromArgb(255, 199, 206)   # Excel "Bad" red
    $yellow    = [System.Drawing.Color]::FromArgb(255, 235, 156)   # Excel "Neutral" yellow
    $gray      = [System.Drawing.Color]::FromArgb(217, 217, 217)   # Excel "Not Applicable" light gray
    $peach     = [System.Drawing.Color]::FromArgb(249, 169, 131)   # Peach (active)
    $lightPeach = [System.Drawing.Color]::FromArgb(251, 226, 213)  # Light peach (passive)

    # ── Conditional formatting rule definitions ──────────────────────────────────
    # Static columns: each entry maps a fixed column letter to an ordered list of
    # (Value, Color) pairs. Order matters where one value is a substring of another
    # (e.g. "not synchronized" must precede "synchronized").
    # Find Uncommitted Changes column letter dynamically
    $colUncommitted = $null
    for ($col = 1; $col -le $lastCol; $col++) {
        if ([string]$ws.Cells[1, $col].Value -eq 'Uncommitted Changes') {
            $colUncommitted = [char](64 + $col)
            break
        }
    }

    $staticFormattingRules = @(
        @{ Column = $colConn;    Rules = @(
            @{ Value = "Disconnected";          Color = $red }
            @{ Value = "Connected";             Color = $green }
            @{ Value = "Not Applicable";        Color = $gray }
        ) }
        @{ Column = $colDG;      Rules = @(
            @{ Value = "Out of Sync";           Color = $red }
            @{ Value = "In Sync";               Color = $green }
            @{ Value = "Not Applicable";        Color = $gray }
            @{ Value = "Not in any device group"; Color = $gray }
        ) }
        @{ Column = $colTmpl;    Rules = @(
            @{ Value = "Out of Sync";           Color = $red }
            @{ Value = "In Sync";               Color = $green }
            @{ Value = "Not Applicable";        Color = $gray }
            @{ Value = "Not in any template";   Color = $gray }
        ) }
        @{ Column = $colHA;      Rules = @(
            @{ Value = "not synchronized";      Color = $red }
            @{ Value = "synchronized";          Color = $green }
            @{ Value = "HA not enabled";        Color = $yellow }
            @{ Value = "Unknown";               Color = $gray }
        ) }
        @{ Column = $colHAState; Rules = @(
            @{ Value = "active";                Color = $peach }
            @{ Value = "passive";               Color = $lightPeach }
            @{ Value = "non-functional";        Color = $red }
            @{ Value = "HA not enabled";        Color = $yellow }
            @{ Value = "Unknown";               Color = $gray }
        ) }
    )

    foreach ($colDef in $staticFormattingRules) {
        $col = $colDef.Column
        foreach ($rule in $colDef.Rules) {
            Add-ConditionalFormatting -Worksheet $ws -Range "${col}2:${col}${lastRow}" -RuleType ContainsText -ConditionValue $rule.Value -BackgroundColor $rule.Color
        }
    }

    # Uncommitted Changes column — applied separately since column letter is discovered at runtime
    if ($colUncommitted) {
        $uncommittedRules = @(
            @{ Value = "Yes";            Color = $red }
            @{ Value = "None";           Color = $green }
            @{ Value = "Unknown";        Color = $gray }
            @{ Value = "Not Applicable"; Color = $gray }
        )
        foreach ($rule in $uncommittedRules) {
            Add-ConditionalFormatting -Worksheet $ws -Range "${colUncommitted}2:${colUncommitted}${lastRow}" -RuleType ContainsText -ConditionValue $rule.Value -BackgroundColor $rule.Color
        }
    }

    # Dynamic HA compat and connection status columns — column letters are discovered
    # at runtime by header name, so each rule set has no Column attached.
    $dynamicCompatHeaders = @(
        'HA PAN-OS Version', 'HA App Version',
        'HA Device Dictionary', 'HA Antivirus Version', 'HA Threat Version',
        'HA Legacy VPN Client Version', 'HA GlobalProtect Version', 'HA DLP Version',
        'HA Cloud Services', 'HA Cloud Connector', 'HA OpenConfig', 'HA VM Series'
    )
    $dynamicConnHeaders = @(
        'HA1 Connection Status', 'HA1 Backup Connection Status', 'HA2 Connection Status', 'HA2 Backup Connection Status'
    )

    $compatRules = @(
        @{ Value = "Mismatch";       Color = $red }
        @{ Value = "Match";          Color = $green }
        @{ Value = "Unknown";        Color = $gray }
        @{ Value = "Not Applicable"; Color = $gray }
    )
    $connRules = @(
        @{ Value = "down";           Color = $red }
        @{ Value = "up";             Color = $green }
        @{ Value = "Unknown";        Color = $gray }
        @{ Value = "Not Applicable"; Color = $gray }
    )

    for ($col = 1; $col -le $lastCol; $col++) {
        $colLetter = [char](64 + $col)
        $header    = [string]$ws.Cells[1, $col].Value

        $ruleSet = if ($dynamicCompatHeaders -contains $header) {
            $compatRules
        } elseif ($dynamicConnHeaders -contains $header) {
            $connRules
        } else {
            $null
        }

        if ($null -ne $ruleSet) {
            foreach ($rule in $ruleSet) {
                Add-ConditionalFormatting -Worksheet $ws -Range "${colLetter}2:${colLetter}${lastRow}" -RuleType ContainsText -ConditionValue $rule.Value -BackgroundColor $rule.Color
            }
        }
    }

        Close-ExcelPackage $excel

    Write-Host "Results written to: $(Resolve-Path $OutputXlsx)" -ForegroundColor Green
} catch {
    Write-Error "Failed to write XLSX: $_ `n$($_.ScriptStackTrace)"
    exit 1
}
# ── Summary table to console ──────────────────────────────────────────────────

Write-Host "`nSummary:" -ForegroundColor Cyan
foreach ($row in $sorted) {
    $header = "$($row.Hostname) ($($row.'Serial Number'))"
    Write-Host $header
    Write-Host ('-' * 60)
    foreach ($prop in $allPropertyNames) {
        if ($prop -ne 'Hostname' -and $prop -ne 'Serial Number') {
            $displayValue = $row.$prop
            if ($prop -eq 'Uncommitted Changes' -and $displayValue -eq 'Yes — see comment') {
                $resultIdx    = $hostnameToResultIndex[$row.Hostname]
                $admins       = if ($null -ne $resultIdx -and $uncommittedComments.ContainsKey($resultIdx)) { $uncommittedComments[$resultIdx] } else { 'Unknown' }
                $displayValue = "Yes ($admins)"
            }
            Write-Host ("  {0,-40} {1}" -f "${prop}:", $displayValue)
        }
    }
    Write-Host ""
}
