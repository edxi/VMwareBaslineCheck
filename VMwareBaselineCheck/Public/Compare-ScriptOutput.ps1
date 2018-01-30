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
        $ht2=$InvokeVesterArgs.Clone()
        $ht2.GetEnumerator()|ForEach-Object{if($_.value -eq $null){$InvokeVesterArgs.Remove($_.key)}}
        $VesterResult = Invoke-Vester @InvokeVesterArgs -PassThru
    }

    # process {
    #     foreach ($axlsxfile in $xlsxfile) {
    #         $allBaselineSettings += Import-xlsx -Path $axlsxfile
    #     }

    #     $allBaselineSettings | Where-Object {$_.Script -ne '' -and $_.'Baseline Value' -ne ''} | ForEach-Object {
    #         $ScriptReturn = &([Scriptblock]::Create($_.Script))
    #         $_.'Actual Value' = $ScriptReturn['Actual Value']
    #         $_.'Check Result' = $ScriptReturn['Check Result']
    #     }
    # }

    # end {
    #     $allBaselineSettings | Export-xlsx -Path "$env:TEMP\$env:COMPUTERNAME-$(Get-Date -UFormat "%Y%m%d-%H%M%S").xlsx" -NoTypeInformation
    # }
}
