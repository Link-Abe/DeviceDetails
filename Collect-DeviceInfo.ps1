<#
.SYNOPSIS
    Collects hardware and battery information from a Windows machine and exports it to an Excel file.

.PARAMETER SavePath
    The directory where the Excel file will be saved. Created if it does not exist.

.PARAMETER FileName
    The name of the output Excel file. The .xlsx extension is appended automatically if absent.
#>
param(
    [string]$SavePath,
    [string]$FileName
)

# --- Module check ---
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    throw "The 'ImportExcel' module is not installed. Install it by running: Install-Module ImportExcel -Scope CurrentUser"
}

# --- Normalize FileName: append .xlsx if missing ---
if (-not $FileName.EndsWith('.xlsx')) {
    $FileName = "$FileName.xlsx"
}

# --- Ensure SavePath directory exists ---
if (-not (Test-Path -Path $SavePath)) {
    try {
        New-Item -Path $SavePath -ItemType Directory -Force -ErrorAction Stop | Out-Null
    }
    catch {
        throw "Failed to create directory '$SavePath': $_"
    }
}

# --- Function stubs (implemented in later tasks) ---

function Get-HardwareInfo {
    $result = @{
        Manufacturer = "Unknown"
        Model        = "Unknown"
        SerialNumber = "Unknown"
    }

    try {
        $cs = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop
        if (-not [string]::IsNullOrWhiteSpace($cs.Manufacturer)) {
            $result.Manufacturer = $cs.Manufacturer
        }
        if (-not [string]::IsNullOrWhiteSpace($cs.Model)) {
            $result.Model = $cs.Model
        }
    }
    catch {
        Write-Warning "Failed to query Win32_ComputerSystem: $_"
    }

    try {
        $bios = Get-CimInstance -ClassName Win32_BIOS -ErrorAction Stop
        if (-not [string]::IsNullOrWhiteSpace($bios.SerialNumber)) {
            $result.SerialNumber = $bios.SerialNumber
        }
    }
    catch {
        Write-Warning "Failed to query Win32_BIOS: $_"
    }

    return $result
}

function Parse-BatteryHtml {
    <#
    .SYNOPSIS
        Parses a powercfg /batteryreport HTML string and returns a hashtable with
        BatteryName, BatteryChemistry, DesignCapacity_mWh, and FullChargeCapacity_mWh.
        Falls back to "Unknown" for unparseable name/chemistry and "N/A" for capacity fields.
    #>
    param(
        [string]$Html
    )

    $result = @{
        BatteryName             = "Unknown"
        BatteryChemistry        = "Unknown"
        DesignCapacity_mWh      = "N/A"
        FullChargeCapacity_mWh  = "N/A"
    }

    if ([string]::IsNullOrWhiteSpace($Html)) {
        return $result
    }

    # --- Parse BatteryName ---
    if ($Html -match '(?i)<td[^>]*>\s*BATTERY NAME\s*</td>\s*<td[^>]*>(.*?)</td>') {
        $name = $Matches[1].Trim()
        if (-not [string]::IsNullOrWhiteSpace($name)) {
            $result.BatteryName = $name
        }
        else {
            Write-Warning "Battery name field is empty; storing 'Unknown'."
        }
    }
    else {
        Write-Warning "Could not parse battery name from report; storing 'Unknown'."
    }

    # --- Parse BatteryChemistry ---
    if ($Html -match '(?i)<td[^>]*>\s*CHEMISTRY\s*</td>\s*<td[^>]*>(.*?)</td>') {
        $chem = $Matches[1].Trim()
        if (-not [string]::IsNullOrWhiteSpace($chem)) {
            $result.BatteryChemistry = $chem
        }
        else {
            Write-Warning "Battery chemistry field is empty; storing 'Unknown'."
        }
    }
    else {
        Write-Warning "Could not parse battery chemistry from report; storing 'Unknown'."
    }

    # --- Parse DesignCapacity_mWh ---
    if ($Html -match '(?i)<td[^>]*>\s*DESIGN CAPACITY\s*</td>\s*<td[^>]*>(.*?)</td>') {
        $raw = $Matches[1].Trim()
        $cleaned = $raw -replace ',', '' -replace '\s*mWh\s*', '' -replace '\s+', ''
        if ([string]::IsNullOrWhiteSpace($cleaned)) {
            Write-Warning "Could not parse design capacity value '$raw'; storing 'N/A'."
        }
        else {
            try {
                $result.DesignCapacity_mWh = [int]$cleaned
            }
            catch {
                Write-Warning "Could not parse design capacity value '$raw'; storing 'N/A'."
            }
        }
    }
    else {
        Write-Warning "Could not parse design capacity from report; storing 'N/A'."
    }

    # --- Parse FullChargeCapacity_mWh ---
    if ($Html -match '(?i)<td[^>]*>\s*FULL CHARGE CAPACITY\s*</td>\s*<td[^>]*>(.*?)</td>') {
        $raw = $Matches[1].Trim()
        $cleaned = $raw -replace ',', '' -replace '\s*mWh\s*', '' -replace '\s+', ''
        if ([string]::IsNullOrWhiteSpace($cleaned)) {
            Write-Warning "Could not parse full charge capacity value '$raw'; storing 'N/A'."
        }
        else {
            try {
                $result.FullChargeCapacity_mWh = [int]$cleaned
            }
            catch {
                Write-Warning "Could not parse full charge capacity value '$raw'; storing 'N/A'."
            }
        }
    }
    else {
        Write-Warning "Could not parse full charge capacity from report; storing 'N/A'."
    }

    return $result
}

function Get-BatteryHealthPercent {
    <#
    .SYNOPSIS
        Calculates battery health as a percentage.
        Returns [math]::Round(($FullChargeCapacity / $DesignCapacity) * 100, 2),
        or "N/A" if DesignCapacity is 0.
    #>
    param(
        [int]$DesignCapacity,
        [int]$FullChargeCapacity
    )

    if ($DesignCapacity -eq 0) {
        return "N/A"
    }

    return [math]::Round(($FullChargeCapacity / $DesignCapacity) * 100, 2)
}

function Read-FileText {
    <#
    .SYNOPSIS
        Thin wrapper around [System.IO.File]::ReadAllText so it can be mocked in tests.
    #>
    param([string]$Path)
    return [System.IO.File]::ReadAllText($Path)
}

function Get-BatteryInfo {
    $result = @{
        BatteryName             = "N/A"
        BatteryChemistry        = "N/A"
        DesignCapacity_mWh      = "N/A"
        FullChargeCapacity_mWh  = "N/A"
        BatteryHealth_Percent   = "N/A"
    }

    # Generate a temp file path for the battery report HTML
    $tempReport = [System.IO.Path]::GetTempFileName() + ".html"

    try {
        $proc = Start-Process -FilePath "powercfg" `
            -ArgumentList "/batteryreport /output `"$tempReport`"" `
            -Wait -PassThru -NoNewWindow -ErrorAction Stop

        if ($proc.ExitCode -ne 0) {
            Write-Warning "powercfg exited with code $($proc.ExitCode); returning N/A for all battery fields."
            return $result
        }
    }
    catch {
        Write-Warning "Failed to run powercfg: $_; returning N/A for all battery fields."
        return $result
    }

    if (-not (Test-Path -Path $tempReport)) {
        Write-Warning "Battery report file not found at '$tempReport'; returning N/A for all battery fields."
        return $result
    }

    try {
        $html = Read-FileText -Path $tempReport
    }
    catch {
        Write-Warning "Failed to read battery report: $_; returning N/A for all battery fields."
        return $result
    }
    finally {
        # Clean up temp file
        if (Test-Path -Path $tempReport) {
            Remove-Item -Path $tempReport -Force -ErrorAction SilentlyContinue
        }
    }

    # Delegate HTML parsing to the testable helper
    $parsed = Parse-BatteryHtml -Html $html

    $result.BatteryName            = $parsed.BatteryName
    $result.BatteryChemistry       = $parsed.BatteryChemistry
    $result.DesignCapacity_mWh     = $parsed.DesignCapacity_mWh
    $result.FullChargeCapacity_mWh = $parsed.FullChargeCapacity_mWh

    # --- Calculate BatteryHealth_Percent ---
    $designCapacity     = $parsed.DesignCapacity_mWh
    $fullChargeCapacity = $parsed.FullChargeCapacity_mWh

    if ($designCapacity -ne "N/A" -and $fullChargeCapacity -ne "N/A") {
        $health = Get-BatteryHealthPercent -DesignCapacity $designCapacity -FullChargeCapacity $fullChargeCapacity
        if ($health -eq "N/A") {
            Write-Warning "Design capacity is 0; cannot calculate battery health, storing 'N/A'."
        }
        $result.BatteryHealth_Percent = $health
    }
    else {
        Write-Warning "Cannot calculate battery health due to missing capacity values; storing 'N/A'."
    }

    return $result
}

function Build-OutputRow {
    <#
    .SYNOPSIS
        Combines hardware and battery hashtables with a Timestamp into a PSCustomObject
        with all nine required columns in fixed order.
    #>
    param(
        [hashtable]$HardwareInfo,
        [hashtable]$BatteryInfo
    )

    return [PSCustomObject]@{
        Timestamp               = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        Manufacturer            = $HardwareInfo.Manufacturer
        Model                   = $HardwareInfo.Model
        SerialNumber            = $HardwareInfo.SerialNumber
        BatteryName             = $BatteryInfo.BatteryName
        BatteryChemistry        = $BatteryInfo.BatteryChemistry
        DesignCapacity_mWh      = $BatteryInfo.DesignCapacity_mWh
        FullChargeCapacity_mWh  = $BatteryInfo.FullChargeCapacity_mWh
        BatteryHealth_Percent   = $BatteryInfo.BatteryHealth_Percent
    }
}

function Export-ToExcel {
    param(
        [PSCustomObject]$Row,
        [string]$FullPath
    )

    try {
        if (-not (Test-Path -Path $FullPath)) {
            $Row | Export-Excel -Path $FullPath -WorksheetName "DeviceInfo" -AutoSize
        }
        else {
            $Row | Export-Excel -Path $FullPath -WorksheetName "DeviceInfo" -Append
        }
    }
    catch [System.IO.IOException] {
        throw "Failed to write to Excel file '$FullPath': the file may be locked by another process. $_"
    }
    catch {
        throw "Failed to write to Excel file '$FullPath': $_"
    }
}

function Resolve-OutputPath {
    <#
    .SYNOPSIS
        Normalises FileName (appending .xlsx if absent) and returns the full output path
        by combining SavePath and FileName via Join-Path.
    #>
    param(
        [string]$SavePath,
        [string]$FileName
    )

    if (-not $FileName.EndsWith('.xlsx')) {
        $FileName = "$FileName.xlsx"
    }

    return Join-Path $SavePath $FileName
}

# --- Main script body ---
$fullPath = Resolve-OutputPath -SavePath $SavePath -FileName $FileName
$hw       = Get-HardwareInfo
$battery  = Get-BatteryInfo
$row      = Build-OutputRow -HardwareInfo $hw -BatteryInfo $battery
Export-ToExcel -Row $row -FullPath $fullPath
