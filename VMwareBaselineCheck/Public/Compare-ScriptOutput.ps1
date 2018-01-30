<#
    .SYNOPSIS
    Load a pre-defined .xlsx file to check baseline item with Script outputs.

    .DESCRIPTION
    Read input from .xlsx file
    Run script block which defined in .xlsx file, to generate outputs.
    Compare items in .xlsx file with script outputs.
    Output compare results to report xlsx file with hostname-datetime.

    .PARAMETER xlsxfile
    Input a .xlsx file which pre-defined format.

    .EXAMPLE
    Compare-ScriptOutput baselineGP.xlsx

    .OUTPUTS
    The compare results to report .xlsx file with hostname-datetime.

    .NOTES
    Author: Xi ErDe
    Date:   Jan 14, 2018
#>

function Compare-ScriptOutput {
    param (
        [CmdletBinding(SupportsShouldProcess = $true, DefaultParameterSetName = 'NoCredential')]
        [Parameter(
            Position = 0,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true,
            HelpMessage = 'Path to one or more xlsx file.')]
        [ValidateNotNullOrEmpty()]
        [SupportsWildcards()]
        [string[]]
        $xlsxfile = '',
        [Parameter(ParameterSetName = 'NoCredential', Mandatory = $false, HelpMessage = 'vCenter server')]
        [Parameter(ParameterSetName = 'WithCredential', Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string] $vCenter = '',
        [Parameter(ParameterSetName = 'WithCredential', Mandatory = $false, HelpMessage = 'vCenter server Credential')]
        [ValidateNotNullOrEmpty()]
        [pscredential]$Credential,
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
        while ($xlsxfile[0] -eq '') {$xlsxfile = Get-FileName}

        if ($vCenter -eq '' -and $Credential -eq $null) {
            Connect-VIServer
        }
        else {
            Connect-VIServer -Server $vCenter -Credential $Credential
        }

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
