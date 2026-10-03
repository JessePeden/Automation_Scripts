#Requires -Version 5.1
# Version: 1
<#
.SYNOPSIS
    Reads Panorama sync and HA status XLSX files and generates a plain-text
    problem summary for use in email notification bodies.

.DESCRIPTION
    Reads all XLSX files in the specified folder, identifies any rows where
    fields are in a problem state, and outputs a grouped plain-text summary.

    Problem conditions:
      Connection Status     : Disconnected
      Device-Group Status   : contains "Out of Sync"
      Template Status       : contains "Out of Sync"
      HA Running-Sync       : contains "not synchronized"
      HA State              : non-functional
      HA compat columns     : Mismatch
      HA conn columns       : down
      Uncommitted Changes   : Yes — see comment (admins read from cell comment)

.PARAMETER XlsxFolder
    Folder containing the XLSX files to read.
    Defaults to the current directory.

.PARAMETER OutputFile
    Path to write the problem summary text to.
    If not specified, output is written to stdout only.

.EXAMPLE
    .\Get-SyncProblems.ps1 -XlsxFolder ./powershell -OutputFile ./powershell/summary.txt
#>

[CmdletBinding()]
param (
    [string]$XlsxFolder = ".",
    [string]$OutputFile  = ""
)

# ── Problem condition definitions ─────────────────────────────────────────────

$problemConditions = [ordered]@{
    'Connection Status'             = { param($v) $v -eq 'Disconnected' }
    'Device-Group Status'           = { param($v) $v -like '*Out of Sync*' }
    'Template Status'               = { param($v) $v -like '*Out of Sync*' }
    'Uncommitted Changes'           = { param($v) $v -eq 'Yes — see comment' }
    'HA Running-Sync'               = { param($v) $v -like '*not synchronized*' }
    'HA State'                      = { param($v) $v -eq 'non-functional' }
    'HA DLP Version'                = { param($v) $v -eq 'Mismatch' }
    'HA PAN-OS Version'             = { param($v) $v -eq 'Mismatch' }
    'HA App Version'                = { param($v) $v -eq 'Mismatch' }
    'HA Antivirus Version'          = { param($v) $v -eq 'Mismatch' }
    'HA Threat Version'             = { param($v) $v -eq 'Mismatch' }
    'HA Legacy VPN Client Version'  = { param($v) $v -eq 'Mismatch' }
    'HA GlobalProtect Version'      = { param($v) $v -eq 'Mismatch' }
    'HA Cloud Services'             = { param($v) $v -eq 'Mismatch' }
    'HA Cloud Connector'            = { param($v) $v -eq 'Mismatch' }
    'HA Device Dictionary'          = { param($v) $v -eq 'Mismatch' }
    'HA OpenConfig'                 = { param($v) $v -eq 'Mismatch' }
    'HA VM Series'                  = { param($v) $v -eq 'Mismatch' }
    'HA1 Connection Status'         = { param($v) $v -eq 'down' }
    'HA1 Backup Connection Status'  = { param($v) $v -eq 'down' }
    'HA2 Connection Status'         = { param($v) $v -eq 'down' }
    'HA2 Backup Connection Status'  = { param($v) $v -eq 'down' }
}

# ── Check for ImportExcel ─────────────────────────────────────────────────────

if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Error "ImportExcel module is not installed. Run: Install-Module -Name ImportExcel -Scope CurrentUser"
    exit 1
}

# ── Read XLSX files ───────────────────────────────────────────────────────────

$xlsxFiles = Get-ChildItem -Path $XlsxFolder -Filter "*.xlsx" | Sort-Object Name

if ($xlsxFiles.Count -eq 0) {
    Write-Warning "No XLSX files found in: $XlsxFolder"
    exit 0
}

$outputLines = @()

foreach ($file in $xlsxFiles) {
    $envName = $file.BaseName -replace '_Sync_Status$', ''

    # Open workbook via EPPlus to access both cell values and comments
    $excel = Open-ExcelPackage -Path $file.FullName
    $ws    = $excel.Workbook.Worksheets["PAN Sync State"]

    if ($null -eq $ws) {
        Write-Warning "Worksheet 'PAN Sync State' not found in $($file.Name) — skipping."
        Close-ExcelPackage $excel -NoSave
        continue
    }

    $lastRow = $ws.Dimension.End.Row
    $lastCol = $ws.Dimension.End.Column

    # Build header -> column index map
    $headerMap = @{}
    for ($col = 1; $col -le $lastCol; $col++) {
        $header = [string]$ws.Cells[1, $col].Value
        if (-not [string]::IsNullOrWhiteSpace($header)) {
            $headerMap[$header] = $col
        }
    }

    $hostnameCol = $headerMap['Hostname']
    $modelCol    = $headerMap['Model']

    if ($null -eq $hostnameCol) {
        Write-Warning "Could not find Hostname column in $($file.Name) — skipping."
        Close-ExcelPackage $excel -NoSave
        continue
    }

    $problemLines = @()

    for ($row = 2; $row -le $lastRow; $row++) {
        $hostname = [string]$ws.Cells[$row, $hostnameCol].Value
        if ([string]::IsNullOrWhiteSpace($hostname)) { continue }
        $model = if ($modelCol) { [string]$ws.Cells[$row, $modelCol].Value } else { '' }

        $rowProblems = @()

        foreach ($field in $problemConditions.Keys) {
            if (-not $headerMap.ContainsKey($field)) { continue }
            $col   = $headerMap[$field]
            $value = [string]$ws.Cells[$row, $col].Value

            if (& $problemConditions[$field] $value) {
                if ($field -eq 'Uncommitted Changes') {
                    # Read admin names from the cell comment
                    $comment = $ws.Cells[$row, $col].Comment
                    $admins  = if ($null -ne $comment) { $comment.Text.Trim() } else { 'unknown admins' }
                    $rowProblems += "$field`: Yes ($admins)"
                } else {
                    # Strip "Name:" prefix from Device-Group and Template Status values
                    $displayValue = $value
                    if ($displayValue -match ':') {
                        $displayValue = ($displayValue -split ':')[-1].Trim()
                    }
                    $rowProblems += "$field`: $displayValue"
                }
            }
        }

        if ($rowProblems.Count -gt 0) {
            $entry = "  $hostname"
            if (-not [string]::IsNullOrWhiteSpace($model)) { $entry += " ($model)" }
            $entry += ' — ' + ($rowProblems -join ', ')
            $problemLines += $entry
        }
    }

    Close-ExcelPackage $excel -NoSave

    if ($problemLines.Count -gt 0) {
        $outputLines += ""
        $outputLines += "$envName"
        $outputLines += ("-" * $envName.Length)
        $outputLines += $problemLines
    }
}

# ── Build final output ────────────────────────────────────────────────────────

$summaryText = if ($outputLines.Count -gt 0) {
    "The following issues were detected:`n" + ($outputLines -join "`n")
} else {
    "No issues detected."
}

Write-Host $summaryText

if (-not [string]::IsNullOrWhiteSpace($OutputFile)) {
    $summaryText | Out-File -FilePath $OutputFile -Encoding UTF8 -NoNewline
    Write-Host "`nSummary written to: $OutputFile"
}
