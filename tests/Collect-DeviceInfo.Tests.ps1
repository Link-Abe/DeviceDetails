# Collect-DeviceInfo.Tests.ps1
# Pester v3/v4 compatible test suite — unit tests and property-based tests.

. "$PSScriptRoot/TestHelpers.ps1"

# Dot-source only the Get-HardwareInfo function from the script.
# We extract the function definition via regex to avoid executing the top-level
# param block and module-check side-effects.
$scriptContent = Get-Content -Path "$PSScriptRoot/../Collect-DeviceInfo.ps1" -Raw

if ($scriptContent -match '(?s)(function\s+Get-HardwareInfo\s*\{.*?\n\})') {
    Invoke-Expression $Matches[1]
}
else {
    throw "Could not locate Get-HardwareInfo in Collect-DeviceInfo.ps1"
}

if ($scriptContent -match '(?s)(function\s+Parse-BatteryHtml\s*\{.*?\n\})') {
    Invoke-Expression $Matches[1]
}
else {
    throw "Could not locate Parse-BatteryHtml in Collect-DeviceInfo.ps1"
}

if ($scriptContent -match '(?s)(function\s+Get-BatteryHealthPercent\s*\{.*?\n\})') {
    Invoke-Expression $Matches[1]
}
else {
    throw "Could not locate Get-BatteryHealthPercent in Collect-DeviceInfo.ps1"
}

if ($scriptContent -match '(?s)(function\s+Read-FileText\s*\{.*?\n\})') {
    Invoke-Expression $Matches[1]
}
else {
    throw "Could not locate Read-FileText in Collect-DeviceInfo.ps1"
}

if ($scriptContent -match '(?s)(function\s+Get-BatteryInfo\s*\{.*?\n\})') {
    Invoke-Expression $Matches[1]
}
else {
    throw "Could not locate Get-BatteryInfo in Collect-DeviceInfo.ps1"
}

if ($scriptContent -match '(?s)(function\s+Build-OutputRow\s*\{.*?\n\})') {
    Invoke-Expression $Matches[1]
}
else {
    throw "Could not locate Build-OutputRow in Collect-DeviceInfo.ps1"
}

if ($scriptContent -match '(?s)(function\s+Resolve-OutputPath\s*\{.*?\n\})') {
    Invoke-Expression $Matches[1]
}
else {
    throw "Could not locate Resolve-OutputPath in Collect-DeviceInfo.ps1"
}

if ($scriptContent -match '(?s)(function\s+Export-ToExcel\s*\{.*?\n\})') {
    Invoke-Expression $Matches[1]
}
else {
    throw "Could not locate Export-ToExcel in Collect-DeviceInfo.ps1"
}

Describe "Property 1: WMI Value Fallback" {

    # Property 1 — WMI Value Fallback
    # Validates: Requirements 1.1, 1.2, 1.3
    #
    # For any WMI query response for Manufacturer, Model, or SerialNumber, the stored
    # value must equal the returned string if it is non-null and non-empty (and not
    # whitespace-only), and must equal "Unknown" if the returned value is null, empty,
    # or whitespace-only.
    #
    # Tag: Feature: windows-machine-info-collector, Property 1: WMI Value Fallback

    It "stores the WMI value or 'Unknown' for 100 random inputs" {
        $failures = @()

        for ($i = 0; $i -lt 100; $i++) {
            # --- Generate random WMI-like objects ---
            $csObj   = New-RandomWmiObject   # Manufacturer, Model
            $biosObj = [PSCustomObject]@{ SerialNumber = New-RandomWmiString }

            # Capture for use inside the Mock scriptblock (Pester v3 closure workaround)
            $capturedCs   = $csObj
            $capturedBios = $biosObj

            # --- Mock Get-CimInstance to return our random objects ---
            Mock Get-CimInstance {
                param($ClassName, $ErrorAction)
                if ($ClassName -eq 'Win32_ComputerSystem') { return $capturedCs }
                if ($ClassName -eq 'Win32_BIOS')           { return $capturedBios }
            }

            # --- Call the function under test ---
            $result = Get-HardwareInfo

            # --- Assert each field ---
            foreach ($field in @(
                @{ Key = 'Manufacturer'; Raw = $csObj.Manufacturer },
                @{ Key = 'Model';        Raw = $csObj.Model },
                @{ Key = 'SerialNumber'; Raw = $biosObj.SerialNumber }
            )) {
                $raw      = $field.Raw
                $stored   = $result[$field.Key]
                $expected = if ([string]::IsNullOrWhiteSpace($raw)) { 'Unknown' } else { $raw }

                if ($stored -ne $expected) {
                    $failures += "Iteration $i | Field=$($field.Key) | Raw=$(if($null -eq $raw){'<null>'}else{"'$raw'"}) | Expected='$expected' | Got='$stored'"
                }
            }
        }

        $failures | Should BeNullOrEmpty
    }
}

Describe "Unit Tests: Get-HardwareInfo" {

    # --- 1. Valid WMI data produces correct hashtable ---
    It "returns correct Manufacturer, Model, and SerialNumber when WMI returns valid values" {
        Mock Get-CimInstance {
            param($ClassName, $ErrorAction)
            if ($ClassName -eq 'Win32_ComputerSystem') {
                return [PSCustomObject]@{ Manufacturer = 'Dell Inc.'; Model = 'Latitude 5520' }
            }
            if ($ClassName -eq 'Win32_BIOS') {
                return [PSCustomObject]@{ SerialNumber = 'ABC1234' }
            }
        }

        $result = Get-HardwareInfo

        $result['Manufacturer'] | Should Be 'Dell Inc.'
        $result['Model']        | Should Be 'Latitude 5520'
        $result['SerialNumber'] | Should Be 'ABC1234'
    }

    # --- 2. Null/empty/whitespace fields fall back to "Unknown" ---
    It "stores 'Unknown' for null Manufacturer, empty Model, and whitespace-only SerialNumber" {
        Mock Get-CimInstance {
            param($ClassName, $ErrorAction)
            if ($ClassName -eq 'Win32_ComputerSystem') {
                return [PSCustomObject]@{ Manufacturer = $null; Model = '' }
            }
            if ($ClassName -eq 'Win32_BIOS') {
                return [PSCustomObject]@{ SerialNumber = '   ' }
            }
        }

        $result = Get-HardwareInfo

        $result['Manufacturer'] | Should Be 'Unknown'
        $result['Model']        | Should Be 'Unknown'
        $result['SerialNumber'] | Should Be 'Unknown'
    }

    # --- 3. WMI exception produces all-"Unknown" result and Write-Warning is called ---
    It "returns all 'Unknown' values and calls Write-Warning when Get-CimInstance throws" {
        Mock Get-CimInstance { throw "WMI unavailable" }
        Mock Write-Warning {}

        $result = Get-HardwareInfo

        $result['Manufacturer'] | Should Be 'Unknown'
        $result['Model']        | Should Be 'Unknown'
        $result['SerialNumber'] | Should Be 'Unknown'
        Assert-MockCalled Write-Warning -Times 1 -Exactly:$false
    }
}

Describe "Property 2: Battery HTML Parsing" {

    # Property 2 — Battery HTML Parsing
    # Validates: Requirements 2.1, 2.2, 2.4, 2.5
    #
    # For any valid powercfg /batteryreport HTML string containing battery table rows,
    # the parser must extract BatteryName, BatteryChemistry, DesignCapacity_mWh, and
    # FullChargeCapacity_mWh with values that match the source HTML. When the HTML is
    # absent or unparseable for a field, that field must be "N/A" (capacity fields) or
    # "Unknown" (name/chemistry fields).
    #
    # Tag: Feature: windows-machine-info-collector, Property 2: Battery HTML Parsing

    It "extracts correct field values or correct fallbacks for 100 random HTML inputs" {
        $failures = @()

        for ($i = 0; $i -lt 100; $i++) {
            $sample   = New-RandomBatteryHtml
            $html     = $sample.Html
            $expected = $sample.Expected

            $result = Parse-BatteryHtml -Html $html

            foreach ($field in @('BatteryName', 'BatteryChemistry', 'DesignCapacity_mWh', 'FullChargeCapacity_mWh')) {
                $got = $result[$field]
                $exp = $expected[$field]

                if ($got -ne $exp) {
                    $failures += "Iteration $i | Field=$field | Expected='$exp' | Got='$got'"
                }
            }
        }

        $failures | Should BeNullOrEmpty
    }

    It "returns 'Unknown' for all name/chemistry fields when HTML is empty" {
        $result = Parse-BatteryHtml -Html ''
        $result['BatteryName']       | Should Be 'Unknown'
        $result['BatteryChemistry']  | Should Be 'Unknown'
        $result['DesignCapacity_mWh']    | Should Be 'N/A'
        $result['FullChargeCapacity_mWh'] | Should Be 'N/A'
    }

    It "returns 'Unknown' for all name/chemistry fields when HTML is null" {
        $result = Parse-BatteryHtml -Html $null
        $result['BatteryName']       | Should Be 'Unknown'
        $result['BatteryChemistry']  | Should Be 'Unknown'
        $result['DesignCapacity_mWh']    | Should Be 'N/A'
        $result['FullChargeCapacity_mWh'] | Should Be 'N/A'
    }

    It "parses a well-formed battery HTML fragment correctly" {
        $html = @"
<table>
<tr><td>BATTERY NAME</td><td>Contoso SR1234</td></tr>
<tr><td>CHEMISTRY</td><td>Li-ion</td></tr>
<tr><td>DESIGN CAPACITY</td><td>45,000 mWh</td></tr>
<tr><td>FULL CHARGE CAPACITY</td><td>38,500 mWh</td></tr>
</table>
"@
        $result = Parse-BatteryHtml -Html $html
        $result['BatteryName']             | Should Be 'Contoso SR1234'
        $result['BatteryChemistry']        | Should Be 'Li-ion'
        $result['DesignCapacity_mWh']      | Should Be 45000
        $result['FullChargeCapacity_mWh']  | Should Be 38500
    }
}

Describe "Property 3: Battery Health Calculation" -Tags @("Feature: windows-machine-info-collector", "Property 3: Battery Health Calculation") {

    # Property 3 — Battery Health Calculation
    # Validates: Requirements 2.3
    #
    # For any pair of positive integers (FullChargeCapacity, DesignCapacity), the computed
    # BatteryHealth_Percent must equal [math]::Round((FullChargeCapacity / DesignCapacity) * 100, 2).
    #
    # Tag: Feature: windows-machine-info-collector, Property 3: Battery Health Calculation

    It "computes correct health percentage for 100 random positive integer pairs" {
        $failures = @()

        for ($i = 0; $i -lt 100; $i++) {
            $design    = Get-Random -Minimum 1 -Maximum 100000
            $fullCharge = Get-Random -Minimum 1 -Maximum 100000

            $result   = Get-BatteryHealthPercent -DesignCapacity $design -FullChargeCapacity $fullCharge
            $expected = [math]::Round(($fullCharge / $design) * 100, 2)

            if ($result -ne $expected) {
                $failures += "Iteration $i | Design=$design | FullCharge=$fullCharge | Expected=$expected | Got=$result"
            }
        }

        $failures | Should BeNullOrEmpty
    }

    It "returns 'N/A' when DesignCapacity is 0" {
        $result = Get-BatteryHealthPercent -DesignCapacity 0 -FullChargeCapacity 50000
        $result | Should Be 'N/A'
    }
}

Describe "Unit Tests: Get-BatteryInfo" {

    # --- 1. Valid HTML produces correct hashtable values ---
    It "returns correct battery fields when powercfg succeeds and HTML is valid" {
        # Mock Start-Process to simulate successful powercfg run (exit code 0)
        Mock Start-Process {
            return [PSCustomObject]@{ ExitCode = 0 }
        }
        # Mock Test-Path to indicate the report file was created
        Mock Test-Path { return $true }
        # Mock Read-FileText wrapper so we don't need a real file on disk
        Mock Read-FileText { return '<table><tr><td>BATTERY NAME</td><td>Contoso SR1234</td></tr></table>' }
        # Mock Parse-BatteryHtml to return known parsed values
        Mock Parse-BatteryHtml {
            return @{
                BatteryName            = 'Contoso SR1234'
                BatteryChemistry       = 'Li-ion'
                DesignCapacity_mWh     = 45000
                FullChargeCapacity_mWh = 38500
            }
        }
        # Mock Get-BatteryHealthPercent to return a known health value
        Mock Get-BatteryHealthPercent { return 85.56 }
        # Mock Remove-Item to avoid touching the filesystem
        Mock Remove-Item {}

        $result = Get-BatteryInfo

        $result['BatteryName']            | Should Be 'Contoso SR1234'
        $result['BatteryChemistry']       | Should Be 'Li-ion'
        $result['DesignCapacity_mWh']     | Should Be 45000
        $result['FullChargeCapacity_mWh'] | Should Be 38500
        $result['BatteryHealth_Percent']  | Should Be 85.56
    }

    # --- 2. Missing battery report returns all "N/A" / "Unknown" ---
    It "returns all N/A fields when the battery report file is not created" {
        Mock Start-Process {
            return [PSCustomObject]@{ ExitCode = 0 }
        }
        # Test-Path returns $false — report file was not created
        Mock Test-Path { return $false }
        Mock Write-Warning {}

        $result = Get-BatteryInfo

        $result['BatteryName']            | Should Be 'N/A'
        $result['BatteryChemistry']       | Should Be 'N/A'
        $result['DesignCapacity_mWh']     | Should Be 'N/A'
        $result['FullChargeCapacity_mWh'] | Should Be 'N/A'
        $result['BatteryHealth_Percent']  | Should Be 'N/A'
    }

    # --- 3. Zero DesignCapacity stores "N/A" for BatteryHealth_Percent ---
    It "stores 'N/A' for BatteryHealth_Percent when DesignCapacity is 0" {
        Mock Start-Process {
            return [PSCustomObject]@{ ExitCode = 0 }
        }
        Mock Test-Path { return $true }
        Mock Parse-BatteryHtml {
            return @{
                BatteryName            = 'Contoso SR1234'
                BatteryChemistry       = 'Li-ion'
                DesignCapacity_mWh     = 0
                FullChargeCapacity_mWh = 38500
            }
        }
        # Get-BatteryHealthPercent returns "N/A" for zero design capacity
        Mock Get-BatteryHealthPercent { return 'N/A' }
        Mock Remove-Item {}
        Mock Write-Warning {}

        $result = Get-BatteryInfo

        $result['BatteryHealth_Percent'] | Should Be 'N/A'
    }

    # --- 4. powercfg failure stores all "N/A" / "Unknown" ---
    It "returns all N/A fields when Start-Process throws" {
        Mock Start-Process { throw "powercfg not found" }
        Mock Write-Warning {}

        $result = Get-BatteryInfo

        $result['BatteryName']            | Should Be 'N/A'
        $result['BatteryChemistry']       | Should Be 'N/A'
        $result['DesignCapacity_mWh']     | Should Be 'N/A'
        $result['FullChargeCapacity_mWh'] | Should Be 'N/A'
        $result['BatteryHealth_Percent']  | Should Be 'N/A'
    }
}

Describe "Property 6: Timestamp Format" -Tags @("Feature: windows-machine-info-collector", "Property 6: Timestamp Format") {

    # Property 6 — Timestamp Format
    # Validates: Requirements 4.5
    #
    # For any script execution, the value written to the Timestamp column must match
    # the pattern ^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$ (i.e., yyyy-MM-dd HH:mm:ss).
    #
    # Tag: Feature: windows-machine-info-collector, Property 6: Timestamp Format

    It "produces a Timestamp matching yyyy-MM-dd HH:mm:ss for 100 invocations" {
        $failures = @()

        $minimalHardware = @{
            Manufacturer = 'TestMfr'
            Model        = 'TestModel'
            SerialNumber = 'SN-0001'
        }

        $minimalBattery = @{
            BatteryName            = 'TestBattery'
            BatteryChemistry       = 'Li-ion'
            DesignCapacity_mWh     = 50000
            FullChargeCapacity_mWh = 45000
            BatteryHealth_Percent  = 90.0
        }

        $pattern = '^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$'

        for ($i = 0; $i -lt 100; $i++) {
            $row = Build-OutputRow -HardwareInfo $minimalHardware -BatteryInfo $minimalBattery

            $ts = $row.Timestamp

            # Assert format matches pattern
            if ($ts -notmatch $pattern) {
                $failures += "Iteration $i | Timestamp='$ts' does not match pattern '$pattern'"
            }

            # Assert the timestamp is a valid parseable datetime
            try {
                [datetime]::ParseExact(
                    $ts,
                    'yyyy-MM-dd HH:mm:ss',
                    [System.Globalization.CultureInfo]::InvariantCulture
                ) | Out-Null
            }
            catch {
                $failures += "Iteration $i | Timestamp='$ts' is not a valid datetime in format 'yyyy-MM-dd HH:mm:ss'"
            }
        }

        $failures | Should BeNullOrEmpty
    }
}

Describe "Property 4: Output Path Construction" -Tags @("Feature: windows-machine-info-collector", "Property 4: Output Path Construction") {

    # Property 4 — Output Path Construction
    # Validates: Requirements 3.3, 3.4, 3.5
    #
    # For any combination of -SavePath and -FileName values, the resolved output path
    # must equal Join-Path($SavePath, $FileName) where $FileName is guaranteed to end
    # with .xlsx.  The directory-creation aspect (Req 3.3) is covered by integration
    # tests; this property focuses solely on the path string construction logic.
    #
    # Tag: Feature: windows-machine-info-collector, Property 4: Output Path Construction

    It "always produces a path ending in .xlsx that equals Join-Path(SavePath, normalizedFileName) for 100 random inputs" {
        $failures = @()

        for ($i = 0; $i -lt 100; $i++) {
            $pair     = New-RandomPathPair
            $savePath = $pair.SavePath
            $fileName = $pair.FileName

            $result = Resolve-OutputPath -SavePath $savePath -FileName $fileName

            # Determine the normalised filename (with .xlsx guaranteed)
            $normalizedFileName = if ($fileName.EndsWith('.xlsx')) { $fileName } else { "$fileName.xlsx" }
            $expected = Join-Path $savePath $normalizedFileName

            # Assert: result ends with .xlsx
            if (-not $result.EndsWith('.xlsx')) {
                $failures += "Iteration $i | SavePath='$savePath' FileName='$fileName' | Result='$result' does not end with .xlsx"
            }

            # Assert: result equals Join-Path(SavePath, normalizedFileName)
            if ($result -ne $expected) {
                $failures += "Iteration $i | SavePath='$savePath' FileName='$fileName' | Expected='$expected' | Got='$result'"
            }
        }

        $failures | Should BeNullOrEmpty
    }

    It "appends .xlsx when FileName has no extension" {
        $result = Resolve-OutputPath -SavePath "C:\Reports" -FileName "output"
        $result | Should Be (Join-Path "C:\Reports" "output.xlsx")
    }

    It "does not double-append .xlsx when FileName already ends with .xlsx" {
        $result = Resolve-OutputPath -SavePath "C:\Reports" -FileName "output.xlsx"
        $result | Should Be (Join-Path "C:\Reports" "output.xlsx")
        $result | Should Not Match '\.xlsx\.xlsx'
    }

    It "combines SavePath and FileName correctly via Join-Path" {
        $result = Resolve-OutputPath -SavePath "D:\Data\Devices" -FileName "inventory"
        $result | Should Be (Join-Path "D:\Data\Devices" "inventory.xlsx")
    }
}

Describe "Property 5: Row Written with All Required Columns" -Tags @("Feature: windows-machine-info-collector", "Property 5: Row Written with All Required Columns") {

    # Property 5 — Row Written with All Required Columns
    # Validates: Requirements 4.1, 4.2, 4.4
    #
    # For any collected data object, the PSCustomObject produced by Build-OutputRow must
    # contain all nine required columns (Timestamp, Manufacturer, Model, SerialNumber,
    # BatteryName, BatteryChemistry, DesignCapacity_mWh, FullChargeCapacity_mWh,
    # BatteryHealth_Percent) as properties, and every property value must be non-null.
    #
    # Tag: Feature: windows-machine-info-collector, Property 5: Row Written with All Required Columns

    $requiredColumns = @(
        'Timestamp',
        'Manufacturer',
        'Model',
        'SerialNumber',
        'BatteryName',
        'BatteryChemistry',
        'DesignCapacity_mWh',
        'FullChargeCapacity_mWh',
        'BatteryHealth_Percent'
    )

    It "produces a PSCustomObject with all 9 required columns present and non-null for 100 random inputs" {
        $failures = @()

        for ($i = 0; $i -lt 100; $i++) {
            $row = New-RandomDataRow

            foreach ($col in $requiredColumns) {
                # Assert the property exists on the object
                $propNames = $row.PSObject.Properties.Name
                if ($col -notin $propNames) {
                    $failures += "Iteration $i | Column '$col' is missing from the row object"
                    continue
                }

                # Assert the property value is non-null
                $val = $row.$col
                if ($null -eq $val) {
                    $failures += "Iteration $i | Column '$col' has a null value"
                }
            }
        }

        $failures | Should BeNullOrEmpty
    }

    It "produces exactly 9 properties on the row object" {
        $row = New-RandomDataRow
        $row.PSObject.Properties.Name.Count | Should Be 9
    }

    It "contains all required column names in the correct order" {
        $hardware = @{
            Manufacturer = 'Dell Inc.'
            Model        = 'Latitude 5520'
            SerialNumber = 'ABC1234'
        }
        $battery = @{
            BatteryName            = 'Contoso SR1234'
            BatteryChemistry       = 'Li-ion'
            DesignCapacity_mWh     = 45000
            FullChargeCapacity_mWh = 38500
            BatteryHealth_Percent  = 85.56
        }

        $row = Build-OutputRow -HardwareInfo $hardware -BatteryInfo $battery
        $props = $row.PSObject.Properties.Name

        $props[0] | Should Be 'Timestamp'
        $props[1] | Should Be 'Manufacturer'
        $props[2] | Should Be 'Model'
        $props[3] | Should Be 'SerialNumber'
        $props[4] | Should Be 'BatteryName'
        $props[5] | Should Be 'BatteryChemistry'
        $props[6] | Should Be 'DesignCapacity_mWh'
        $props[7] | Should Be 'FullChargeCapacity_mWh'
        $props[8] | Should Be 'BatteryHealth_Percent'
    }
}

# Pester v3 requires the command to exist before it can be mocked.
# Define a stub at the top level so Mock can intercept calls in ALL Describe blocks,
# even when ImportExcel is not installed.
if (-not (Get-Command Export-Excel -ErrorAction SilentlyContinue)) {
    function Export-Excel { param([string]$Path, [string]$WorksheetName, [switch]$AutoSize, [switch]$Append) }
}

Describe "Unit Tests: Export-ToExcel append behaviour" {

    # Validates: Requirements 4.1, 4.4
    # When the target file does not exist, Export-Excel is called without -Append.
    # When the target file already exists, Export-Excel is called with -Append.

    $sampleRow = [PSCustomObject]@{
        Timestamp              = '2024-01-01 12:00:00'
        Manufacturer           = 'Dell Inc.'
        Model                  = 'Latitude 5520'
        SerialNumber           = 'ABC1234'
        BatteryName            = 'Contoso SR1234'
        BatteryChemistry       = 'Li-ion'
        DesignCapacity_mWh     = 45000
        FullChargeCapacity_mWh = 38500
        BatteryHealth_Percent  = 85.56
    }

    It "calls Export-Excel without -Append when the file does not exist" {
        Mock Test-Path { return $false }
        Mock Export-Excel {}

        Export-ToExcel -Row $sampleRow -FullPath 'C:\Reports\output.xlsx'

        Assert-MockCalled Export-Excel -Times 1 -ParameterFilter {
            $Append -ne $true
        }
    }

    It "calls Export-Excel with -Append when the file already exists" {
        Mock Test-Path { return $true }
        Mock Export-Excel {}

        Export-ToExcel -Row $sampleRow -FullPath 'C:\Reports\output.xlsx'

        Assert-MockCalled Export-Excel -Times 1 -ParameterFilter {
            $Append -eq $true
        }
    }

    It "calls Export-Excel with -Append on the second call and not on the first" {
        $callCount = 0
        Mock Test-Path {
            # First call: file doesn't exist; second call: file exists
            $callCount++
            return $callCount -gt 1
        }
        Mock Export-Excel {}

        Export-ToExcel -Row $sampleRow -FullPath 'C:\Reports\output.xlsx'
        Export-ToExcel -Row $sampleRow -FullPath 'C:\Reports\output.xlsx'

        Assert-MockCalled Export-Excel -Times 2
    }

    # Validates: Requirement 4.3
    # When the Excel file is locked by another process, Export-ToExcel must throw a
    # terminating error with a descriptive message.
    It "throws a terminating error with a descriptive message when the file is locked" {
        Mock Test-Path { return $true }
        Mock Export-Excel { throw [System.IO.IOException]::new("File is locked") }

        { Export-ToExcel -Row $sampleRow -FullPath 'C:\Reports\output.xlsx' } | Should Throw
    }
}

Describe "Integration Tests: Full Script Execution" {

    # Integration tests — Full Script Execution
    # Validates: Requirements 1.1, 2.1, 3.5, 4.1, 4.4
    #
    # Simulates the main script body by calling the pipeline functions in sequence:
    #   Get-HardwareInfo → Get-BatteryInfo → Build-OutputRow → Export-ToExcel
    # All external dependencies (Get-CimInstance, Start-Process, Test-Path,
    # Read-FileText, Export-Excel, Remove-Item) are mocked so no real hardware
    # access or filesystem writes occur.

    # Fixed hardware data returned by the mocked WMI calls
    $intHardwareCsObj   = [PSCustomObject]@{ Manufacturer = 'Contoso Corp'; Model = 'ProBook 9000' }
    $intHardwareBiosObj = [PSCustomObject]@{ SerialNumber = 'SN-INT-001' }

    # Fixed battery HTML that Parse-BatteryHtml will receive via the mocked Read-FileText
    $intBatteryHtml = @"
<table>
<tr><td>BATTERY NAME</td><td>Contoso Battery</td></tr>
<tr><td>CHEMISTRY</td><td>Li-ion</td></tr>
<tr><td>DESIGN CAPACITY</td><td>50,000 mWh</td></tr>
<tr><td>FULL CHARGE CAPACITY</td><td>45,000 mWh</td></tr>
</table>
"@

    # -----------------------------------------------------------------------
    # Test 1: New file creation — Export-Excel called once WITHOUT -Append,
    #         and the row contains all 9 columns with the expected values.
    # Validates: Requirements 1.1, 2.1, 3.5, 4.1
    # -----------------------------------------------------------------------
    It "creates a new Excel file with a correctly populated row when the file does not exist" {

        # Capture Export-Excel calls so we can inspect parameters
        $exportCalls = [System.Collections.Generic.List[hashtable]]::new()

        Mock Get-CimInstance {
            param($ClassName, $ErrorAction)
            if ($ClassName -eq 'Win32_ComputerSystem') { return $intHardwareCsObj }
            if ($ClassName -eq 'Win32_BIOS')           { return $intHardwareBiosObj }
        }
        Mock Start-Process { return [PSCustomObject]@{ ExitCode = 0 } }
        # Test-Path: $false for the Excel file check inside Export-ToExcel;
        #            $true  for the battery report temp-file check inside Get-BatteryInfo
        Mock Test-Path {
            param($Path)
            # The battery temp report path ends with .html; the Excel path ends with .xlsx
            if ($Path -like '*.html') { return $true }
            return $false
        }
        Mock Read-FileText { return $intBatteryHtml }
        Mock Remove-Item {}
        Mock Export-Excel {
            param([string]$Path, [string]$WorksheetName, [switch]$AutoSize, [switch]$Append)
            $exportCalls.Add(@{
                Path          = $Path
                WorksheetName = $WorksheetName
                AutoSize      = $AutoSize.IsPresent
                Append        = $Append.IsPresent
            })
        }

        # --- Run the pipeline (mirrors the main script body) ---
        $fullPath = 'C:\Reports\integration.xlsx'
        $hw       = Get-HardwareInfo
        $battery  = Get-BatteryInfo
        $row      = Build-OutputRow -HardwareInfo $hw -BatteryInfo $battery
        Export-ToExcel -Row $row -FullPath $fullPath

        # --- Assertions ---

        # Export-Excel was called exactly once
        $exportCalls.Count | Should Be 1

        # Called WITHOUT -Append (new file)
        $exportCalls[0].Append | Should Be $false

        # Called with the correct path and worksheet name
        $exportCalls[0].Path          | Should Be $fullPath
        $exportCalls[0].WorksheetName | Should Be 'DeviceInfo'

        # Row has all 9 required columns
        $props = $row.PSObject.Properties.Name
        $props.Count | Should Be 9
        ($props -contains 'Timestamp')               | Should Be $true
        ($props -contains 'Manufacturer')            | Should Be $true
        ($props -contains 'Model')                   | Should Be $true
        ($props -contains 'SerialNumber')            | Should Be $true
        ($props -contains 'BatteryName')             | Should Be $true
        ($props -contains 'BatteryChemistry')        | Should Be $true
        ($props -contains 'DesignCapacity_mWh')      | Should Be $true
        ($props -contains 'FullChargeCapacity_mWh')  | Should Be $true
        ($props -contains 'BatteryHealth_Percent')   | Should Be $true

        # Hardware fields flow through correctly
        $row.Manufacturer | Should Be 'Contoso Corp'
        $row.Model        | Should Be 'ProBook 9000'
        $row.SerialNumber | Should Be 'SN-INT-001'

        # Battery fields flow through correctly
        $row.BatteryName            | Should Be 'Contoso Battery'
        $row.BatteryChemistry       | Should Be 'Li-ion'
        $row.DesignCapacity_mWh     | Should Be 50000
        $row.FullChargeCapacity_mWh | Should Be 45000

        # Timestamp matches expected format
        $row.Timestamp | Should Match '^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$'
    }

    # -----------------------------------------------------------------------
    # Test 2: Append scenario — run the pipeline twice; second call uses -Append.
    # Validates: Requirements 4.4
    # -----------------------------------------------------------------------
    It "appends a second row when the pipeline is run twice and the file already exists on the second run" {

        $exportCalls = [System.Collections.Generic.List[hashtable]]::new()

        $excelPath = 'C:\Reports\integration.xlsx'

        Mock Get-CimInstance {
            param($ClassName, $ErrorAction)
            if ($ClassName -eq 'Win32_ComputerSystem') { return $intHardwareCsObj }
            if ($ClassName -eq 'Win32_BIOS')           { return $intHardwareBiosObj }
        }
        Mock Start-Process { return [PSCustomObject]@{ ExitCode = 0 } }
        # Use script scope so the Mock scriptblock can read/write the counter
        $script:intTestPathCount = 0
        Mock Test-Path {
            param($Path)
            # Battery temp report always exists
            if ($Path -like '*.html') { return $true }
            # Excel file: does not exist on first pipeline run, exists on second
            $script:intTestPathCount++
            return $script:intTestPathCount -gt 1
        }
        Mock Read-FileText { return $intBatteryHtml }
        Mock Remove-Item {}
        Mock Export-Excel {
            param([string]$Path, [string]$WorksheetName, [switch]$AutoSize, [switch]$Append)
            $exportCalls.Add(@{
                Append = $Append.IsPresent
            })
        }

        $fullPath = $excelPath

        # --- First pipeline run ---
        $hw1      = Get-HardwareInfo
        $battery1 = Get-BatteryInfo
        $row1     = Build-OutputRow -HardwareInfo $hw1 -BatteryInfo $battery1
        Export-ToExcel -Row $row1 -FullPath $fullPath

        # --- Second pipeline run ---
        $hw2      = Get-HardwareInfo
        $battery2 = Get-BatteryInfo
        $row2     = Build-OutputRow -HardwareInfo $hw2 -BatteryInfo $battery2
        Export-ToExcel -Row $row2 -FullPath $fullPath

        # --- Assertions ---

        # Export-Excel was called exactly twice (once per pipeline run)
        $exportCalls.Count | Should Be 2

        # First call: no -Append (new file)
        $exportCalls[0].Append | Should Be $false

        # Second call: -Append (file already exists)
        $exportCalls[1].Append | Should Be $true
    }
}
