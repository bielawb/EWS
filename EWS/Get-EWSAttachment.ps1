function Get-EWSAttachment {
    [OutputType('Microsoft.Exchange.WebServices.Data.FileAttachment')]
    [CmdletBinding()]
    param (
        [Parameter(
                ValueFromPipeline
        )]
        [Microsoft.Exchange.WebServices.Data.Item[]]$Item,
        
        [Parameter(
                ValueFromPipelineByPropertyName
        )]
        [Microsoft.Exchange.WebServices.Data.ExchangeService]$Service = $script:exchangeService

    )

    process {
        if (-not $Service) {
            return
        }
        
        $null = $Service.LoadPropertiesForItems(
            $Item,
            [Microsoft.Exchange.WebServices.Data.PropertySet]::FirstClassProperties
        )

        foreach ($singleItem in $Item) {
            foreach ($attachment in $singleItem.Attachments) {
                $attachment
            }
        }
    }
    
    end {
        if (-not $Service) {
            Write-Warning 'No connection defined. Use Connect-EWSService first!'
        }
    }
}
