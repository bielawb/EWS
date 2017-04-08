function Remove-EWSFolder {
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
        [Microsoft.Exchange.WebServices.Data.Folder]$Folder
    )

    process {
        if (-not $Service) {
            return
        }
        if ($PSCmdlet.ShouldProcess(
                "Folder: $($Folder.DisplayName)",
                'Remove EWS Folder'
        )) {
            $Folder.Delete($DeleteMode)
        }
    }

    end {
        if (-not $Service) {
            Write-Warning 'No connection defined. Use Connect-EWSService first!'
            return
        }
    }
}
