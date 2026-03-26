# TestHelpers.ps1
# Random data generator functions for property-based tests.

<#
.SYNOPSIS
    Returns a random WMI-like response string for Manufacturer, Model, or SerialNumber.
    The returned value may be: a valid non-empty string, an empty string, a whitespace-only
    string, or $null — mirroring the full range of values WMI can return in practice.
#>
function New-RandomWmiString {
    $roll = Get-Random -Minimum 0 -Maximum 4
    switch ($roll) {
        0 {
            # Valid non-empty string (letters, digits, spaces, hyphens)
            $chars  = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789 -'
            $length = Get-Random -Minimum 1 -Maximum 32
            -join (1..$length | ForEach-Object { $chars[(Get-Random -Maximum $chars.Length)] })
        }
        1 { '' }                          # Empty string
        2 { '   ' }                       # Whitespace-only
        3 { $null }                       # Null
    }
}

<#
.SYNOPSIS
    Builds a mock CIM object (PSCustomObject) that resembles a Win32_ComputerSystem or
    Win32_BIOS instance, with random values for Manufacturer, Model, and SerialNumber.
#>
function New-RandomWmiObject {
    [PSCustomObject]@{
        Manufacturer = New-RandomWmiString
        Model        = New-RandomWmiString
        SerialNumber = New-RandomWmiString
    }
}

<#
.SYNOPSIS
    Generates a random battery HTML fragment that mimics a powercfg /batteryreport table.

    Returns a hashtable with two keys:
      Html     - the generated HTML string (may be empty/null to test absent-HTML path)
      Expected - a hashtable of the expected parsed values:
                   BatteryName, BatteryChemistry, DesignCapacity_mWh, FullChargeCapacity_mWh

    Each field is independently randomised to be present (valid), present-but-empty,
    or absent entirely, so the generator covers all fallback branches.
#>
function New-RandomBatteryHtml {

    # Helper: random printable string (no HTML special chars to keep regex simple)
    function _RandStr {
        $chars  = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789 -_'
        $length = Get-Random -Minimum 1 -Maximum 40
        -join (1..$length | ForEach-Object { $chars[(Get-Random -Maximum $chars.Length)] })
    }

    # Helper: random positive integer capacity (1 – 99999)
    function _RandCapacity {
        Get-Random -Minimum 1 -Maximum 100000
    }

    # Helper: format a capacity integer as the HTML report does, e.g. "45,000 mWh"
    function _FormatCapacity([int]$v) {
        # Insert comma thousands separator manually (no culture dependency)
        $s = $v.ToString()
        if ($s.Length -gt 3) {
            $s = $s.Substring(0, $s.Length - 3) + ',' + $s.Substring($s.Length - 3)
        }
        "$s mWh"
    }

    # Decide presence for each field: 0=present+valid, 1=present+empty, 2=absent
    $nameMode  = Get-Random -Minimum 0 -Maximum 3
    $chemMode  = Get-Random -Minimum 0 -Maximum 3
    $designMode = Get-Random -Minimum 0 -Maximum 3
    $fullMode   = Get-Random -Minimum 0 -Maximum 3

    $nameVal   = if ($nameMode  -eq 0) { _RandStr }    else { '' }
    $chemVal   = if ($chemMode  -eq 0) { _RandStr }    else { '' }
    $designVal = if ($designMode -eq 0) { _RandCapacity } else { $null }
    $fullVal   = if ($fullMode   -eq 0) { _RandCapacity } else { $null }

    # Build HTML rows only for present fields
    $rows = ''
    if ($nameMode -ne 2) {
        $rows += "<tr><td>BATTERY NAME</td><td>$nameVal</td></tr>`n"
    }
    if ($chemMode -ne 2) {
        $rows += "<tr><td>CHEMISTRY</td><td>$chemVal</td></tr>`n"
    }
    if ($designMode -ne 2) {
        $designHtml = if ($null -ne $designVal) { _FormatCapacity $designVal } else { '' }
        $rows += "<tr><td>DESIGN CAPACITY</td><td>$designHtml</td></tr>`n"
    }
    if ($fullMode -ne 2) {
        $fullHtml = if ($null -ne $fullVal) { _FormatCapacity $fullVal } else { '' }
        $rows += "<tr><td>FULL CHARGE CAPACITY</td><td>$fullHtml</td></tr>`n"
    }

    $html = "<table>`n$rows</table>"

    # Compute expected parsed values (mirrors Parse-BatteryHtml fallback logic)
    $expName  = if ($nameMode  -eq 0) { $nameVal.Trim() }  else { 'Unknown' }
    $expChem  = if ($chemMode  -eq 0) { $chemVal.Trim() }  else { 'Unknown' }
    $expDesign = if ($designMode -eq 0 -and $null -ne $designVal) { $designVal } else { 'N/A' }
    $expFull   = if ($fullMode   -eq 0 -and $null -ne $fullVal)   { $fullVal }   else { 'N/A' }

    return @{
        Html     = $html
        Expected = @{
            BatteryName            = $expName
            BatteryChemistry       = $expChem
            DesignCapacity_mWh     = $expDesign
            FullChargeCapacity_mWh = $expFull
        }
    }
}

<#
.SYNOPSIS
    Generates a random SavePath / FileName pair for property-based testing of output
    path construction.

    Returns a hashtable with:
      SavePath  - a random directory-like path string (no trailing slash)
      FileName  - a random filename, randomly with or without the .xlsx extension
      HasXlsx   - $true if FileName already ends with .xlsx, $false otherwise
#>
function New-RandomPathPair {
    # Random path segment characters (safe for Join-Path on Windows)
    $chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789_-'

    function _RandSegment {
        $len = Get-Random -Minimum 1 -Maximum 16
        -join (1..$len | ForEach-Object { $chars[(Get-Random -Maximum $chars.Length)] })
    }

    # Build a 1-3 segment path (e.g. "C:\Reports\Q1" style, but without a drive letter
    # so it stays portable and doesn't require the path to exist)
    $depth = Get-Random -Minimum 1 -Maximum 4
    $segments = 1..$depth | ForEach-Object { _RandSegment }
    $savePath = $segments -join [System.IO.Path]::DirectorySeparatorChar

    # Random base filename
    $baseName = _RandSegment

    # 50 % chance the caller already supplies the .xlsx extension
    $hasXlsx = (Get-Random -Minimum 0 -Maximum 2) -eq 0
    $fileName = if ($hasXlsx) { "$baseName.xlsx" } else { $baseName }

    return @{
        SavePath = $savePath
        FileName = $fileName
        HasXlsx  = $hasXlsx
    }
}

<#
.SYNOPSIS
    Generates a random hardware hashtable and battery hashtable, then calls Build-OutputRow
    to produce a PSCustomObject with all nine required columns.

    Returns the PSCustomObject directly.
#>
function New-RandomDataRow {
    # Random non-empty string helper
    function _RandStr {
        $chars  = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789 -_'
        $length = Get-Random -Minimum 1 -Maximum 32
        -join (1..$length | ForEach-Object { $chars[(Get-Random -Maximum $chars.Length)] })
    }

    # Random capacity value: either a positive integer or "N/A"
    function _RandCapacity {
        if ((Get-Random -Minimum 0 -Maximum 2) -eq 0) {
            return Get-Random -Minimum 1 -Maximum 100000
        }
        return 'N/A'
    }

    # Random health: either a decimal or "N/A"
    function _RandHealth {
        if ((Get-Random -Minimum 0 -Maximum 2) -eq 0) {
            return [math]::Round((Get-Random -Minimum 1 -Maximum 10000) / 100.0, 2)
        }
        return 'N/A'
    }

    $hardware = @{
        Manufacturer = _RandStr
        Model        = _RandStr
        SerialNumber = _RandStr
    }

    $battery = @{
        BatteryName            = _RandStr
        BatteryChemistry       = _RandStr
        DesignCapacity_mWh     = _RandCapacity
        FullChargeCapacity_mWh = _RandCapacity
        BatteryHealth_Percent  = _RandHealth
    }

    return Build-OutputRow -HardwareInfo $hardware -BatteryInfo $battery
}
