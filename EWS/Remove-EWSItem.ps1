function Remove-EWSItem {
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
        [Microsoft.Exchange.WebServices.Data.Item]$Item
    )
    
    process {
        if (-not $Service) {
            return
        }
        if ($PSCmdlet.ShouldProcess(
                "Item: $($Item.Subject)",
                "Remove using $DeleteMode mode"
        )) {
            $Item.Delete($DeleteMode)
        }
    }
    

    end {
        if (-not $Service) {
            Write-Warning 'No connection defined. Use Connect-EWSService first!'
            return
        }
    }
}
