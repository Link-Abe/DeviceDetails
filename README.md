# DeviceDetails
Pulling out Device Details To Collect Information for future use. Information such as Device Make, Model, Serial Number, Battery Health and Type

## Usage

Run `Collect-DeviceInfo.ps1` from a PowerShell session on a Windows 10/11 machine. The `ImportExcel` module must be installed (`Install-Module ImportExcel`).

### Parameters

- `-SavePath` *(string, required)* — The directory where the Excel file will be saved. If the directory does not exist, the script creates it automatically.
- `-FileName` *(string, required)* — The name of the Excel file. The `.xlsx` extension is appended automatically if omitted.

### Examples

```powershell
# Save to C:\Reports with a specific filename
.\Collect-DeviceInfo.ps1 -SavePath "C:\Reports" -FileName "DeviceInfo"

# Save to a network share, filename already includes extension
.\Collect-DeviceInfo.ps1 -SavePath "\\server\share\IT" -FileName "DeviceInfo.xlsx"
```

Each run appends a new row to the Excel file (or creates the file if it does not exist). The output contains the following columns: `Timestamp`, `Manufacturer`, `Model`, `SerialNumber`, `BatteryName`, `BatteryChemistry`, `DesignCapacity_mWh`, `FullChargeCapacity_mWh`, `BatteryHealth_Percent`.
