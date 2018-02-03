<#
    .SYNOPSIS
    Load pre-defined .xlsx file(s) to check baseline items with Vester outputs.

    .DESCRIPTION
    Read baseline check excel file.
    Connect to vcenter server(s) and run Vester tests.
    Optionally generate a new Vester config from exist server(s).
    Compare items in .xlsx file with Vester outputs.
    Fill compare results up to excel file.

    .PARAMETER xlsxfile
    Input .xlsx file(s) which pre-defined format.

    .PARAMETER vCenter
    Connect vCenter Server(s) to check configurations.

    .PARAMETER Credential
    vCenter Credential

    .PARAMETER ReadNewConfig
    Generate new Vester config before run Vester tests.

    .PARAMETER OutputFolder
    Folder pass to New-VesterConfig while need generate new Vester config.

    .PARAMETER Config
    Config file pass to Invoke-Vester. It will be overwrited if ReadNewConfig parameter used.

    .PARAMETER Test
    Test folder or file pass to Invoke-Vester.

    .EXAMPLE
    PS C:\>Import-Module VMware.VimAutomation.Vds
    PS C:\>Import-Module PSExcel
    PS C:\>Compare-VesterOutput
    Run Directly will prompt provide excel file, and vcenter connection.
    Vester tests will use Vester module default config file and Tests scripts.
    In most cases, PSExcel and VMware VDs module must be imported explicitly.

    .EXAMPLE
    PS C:\>Compare-VesterOutput -ReadNewConfig
    A new Vester configuration will be generated at VMwareBaselineCheck module's config folder.
    Vester tests will use this config file.

    .EXAMPLE
    PS C:\>$xlsxfiles = @(".\baseline1.xlsx","c:\temp\baseline.xlsx")
    PS C:\>$vCenters = @("vcenter1.vmlab.com","192.168.100.10")
    PS C:\>$Credential = Get-Credential
    PS C:\>$xlsxfiles | Compare-VesterOutput -vCenter $vCenters -Credential $Credential -Test ".\Tests"
    It connects two vcenters and compare two excel files.

    .OUTPUTS
    The compare result write to predefined excel file's corresponded columns.
#>

function Compare-VesterOutput {
    param (
        [CmdletBinding(SupportsShouldProcess = $true, DefaultParameterSetName = 'xlsxfile')]
        [Parameter(
            Position = 0,
            Mandatory = $true,
            ParameterSetName = 'xlsxfile',
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true,
            HelpMessage = 'Path to one or more xlsx file.')]
        [ValidateNotNullOrEmpty()]
        [string[]]
        $xlsxfile,
        [Parameter(Mandatory = $false, HelpMessage = 'Generate new config JSON')]
        [switch]$ReadNewConfig = $false,
        [Parameter(Mandatory = $false, HelpMessage = 'Folder pass to New-VesterConfig')]
        [ValidateScript( {Test-Path $_ -PathType Container})]
        [object]$OutputFolder = "$(Split-Path -Parent $PSScriptRoot)\Configs",
        [Parameter(Mandatory = $false, HelpMessage = 'Config file pass to Invoke-Vester')]
        [object[]]$Config = $null,
        [Parameter(Mandatory = $false, HelpMessage = 'Test folder or file pass to Invoke-Vester')]
        [object[]]$Test = $null
    )

    begin {
        if ($ReadNewConfig) {
            New-VesterConfig -OutputFolder $OutputFolder
            $Config = "$OutputFolder\Config.json"
        }

        $InvokeVesterArgs = @{ Config = $Config; Test = $Test }
        $ht2 = $InvokeVesterArgs.Clone()
        $ht2.GetEnumerator()|ForEach-Object {if ($_.value -eq $null) {$InvokeVesterArgs.Remove($_.key)}}
        $VesterResult = Invoke-Vester @InvokeVesterArgs -PassThru
    }

    process {
        foreach ($axlsxfile in $xlsxfile) {
            $ExcelVar = New-Excel -Path $axlsxfile
            foreach ($ExcelSheet in $ExcelVar.Workbook.Worksheets) {
                $ExcelSheet.Tables | ForEach-Object {
                    $Coordinates = $_.address.address
                    $ColumnStart = ($($Coordinates -split ":")[0] -replace "[0-9]", "").ToUpperInvariant()
                    $ColumnEnd = ($($Coordinates -split ":")[1] -replace "[0-9]", "").ToUpperInvariant()
                    [int]$RowStart = $($Coordinates -split ":")[0] -replace "[a-zA-Z]", ""
                    [int]$RowEnd = $($Coordinates -split ":")[1] -replace "[a-zA-Z]", ""
                    $Rows = $RowEnd - $RowStart + 1
                    $ColumnStart = Get-ExcelColumnInt $ColumnStart
                    $ColumnEnd = Get-ExcelColumnInt $ColumnEnd
                    $Columns = $ColumnEnd - $ColumnStart + 1

                    for ($i = $ColumnStart; $i -le $Columns; $i++) {
                        if ($ExcelSheet.GetValue($RowStart, $i) -eq "Test Item") {
                            $TestItemCol = $i
                        }
                        if ($ExcelSheet.GetValue($RowStart, $i) -eq "Difference") {
                            $DifferenceCol = $i
                        }
                        if ($ExcelSheet.GetValue($RowStart, $i) -eq "Compliance") {
                            $ComplianceCol = $i
                        }
                    }

                    if ($TestItemCol -ne $null -and $DifferenceCol -ne $null -and $ComplianceCol -ne $null) {
                        for ($i = $RowStart + 1; $i -lt $Rows; $i++) {
                            if ($ExcelSheet.GetValue($i, $TestItemCol) -ne $null) {
                                $VesterResult.TestResult | ForEach-Object {
                                    if ($_.Name -match $ExcelSheet.GetValue($i, $TestItemCol)) {
                                        $Compliance = @{$true = 'Yes'; $false = 'No'}
                                        $ExcelSheet.SetValue($i, $ComplianceCol, $Compliance[$_.Passed])
                                        $ExcelSheet.SetValue($i, $DifferenceCol, $_.FailureMessage)
                                    }
                                }
                            }
                        }
                    }
                }
            }
            $ExcelVar | Save-Excel -Close
        }
    }

    end {
    }
}
