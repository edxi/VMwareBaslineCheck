# VMwareBaslineCheck

[![Build status](https://ci.appveyor.com/api/projects/status/p778wr2eg3hvmcoc?svg=true)](https://ci.appveyor.com/project/edxi/vmwarebaslinecheck)

The purpose of this script which is provides VMware baseline check functions.
It reads baseline items from a predefined excel file. And it based on [Vester](https://github.com/WahlNetwork/Vester) to get server configurations.
> Vester is a community project that provides an extremely light-weight approach to configuration management of your VMware environment.
Exports the Vester tests output as check results to the predefined excel file.

## Features

* Read/Write items in predefined excel file's tables. Only following predefined table headers related, regardless other table contents.
  * Test Item - Compares Vester test results with this header. If the content match to a result, test result will write to below two headers.
  * Difference - Writes configuration difference if Vester test does not passed.
  * Compliance - Yes/No by Vester tests passed.
* Connect to vcenter server(s) and run Vester tests.
* Optionally generate a new Vester config from exist server(s). (Run `New-Vester` before `Invoke-Vester`)

## Examples

### Run without parameter

```powershell
PS C:\>Import-Module VMware.VimAutomation.Vds
PS C:\>Import-Module PSExcel
PS C:\>Compare-VesterOutput
```

Run Directly will prompt provide excel file, and vcenter connection.
Vester tests will use Vester module default config file and Tests scripts.
In most cases, PSExcel and VMware VDs module must be imported explicitly.

### Generate config first

```powershell
PS C:\>Compare-VesterOutput -ReadNewConfig
```

A new Vester configuration will be generated at VMwareBaselineCheck module's config folder.
Vester tests will use this config file.

### Pipeline, Multiple vCenters and baseline files

```powershell
PS C:\>$xlsxfiles = @(".\baseline1.xlsx","c:\temp\baseline.xlsx")
PS C:\>$vCenters = @("vcenter1.vmlab.com","192.168.100.10")
PS C:\>$Credential = Get-Credential
PS C:\>$xlsxfiles | Compare-VesterOutput -vCenter $vCenters -Credential $Credential -Test ".\Tests"
```

It connects two vcenters and compares two excel files.

## Feedback

Please send your feedback to <https://github.com/edxi/BaselineCheck/issues>