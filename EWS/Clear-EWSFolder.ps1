function Clear-EWSFolder {
    [CmdletBinding(
        SupportsShouldProcess,
        ConfirmImpact = 'High'
    )]
    param (
        [Microsoft.Exchange.WebServices.Data.DeleteMode]$DeleteMode = 'MoveToDeletedItems',

        [Parameter(
            ValueFromPipelineByPropertyName
        )]
        [Microsoft.Exchange.WebServices.Data.ExchangeService]$Service = $Script:exchangeService,
        
        [Parameter(
            ValueFromPipeline,
            Mandatory
        )]
        [Microsoft.Exchange.WebServices.Data.Folder]$Folder,
        
        [switch]$IncludeSubfolders
    )

    process {
        if (-not $Service) {
            return
        }
        if ($PSCmdlet.ShouldProcess(
            "Folder: $($Folder.DisplayName)",
            "Clear (empty) EWS Folder with$('out' * (-not $IncludeSubfolders)) subfolders"
        )) {
            $Folder.Empty($DeleteMode, $IncludeSubfolders)
        }
    }

    end {
        if (-not $Service) {
            Write-Warning 'No connection defined. Use Connect-EWSService first!'
            return
        }
    }
}
