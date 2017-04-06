function Connect-EWSService {
    [OutputType('Microsoft.Exchange.WebServices.Data.ExchangeService')]
    [CmdletBinding()]
    param (
		
        [Parameter(
                Mandatory
        )]
        [String]$Mailbox,
        

        [Microsoft.Exchange.WebServices.Data.ExchangeVersion]
        $Version = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1,

        [Management.Automation.PSCredential]
        [Management.Automation.Credential()]
        $Credential = [Management.Automation.PSCredential]::Empty
    )

    $exchangeService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService $Version

    if ($Credential -ne [Management.Automation.PSCredential]::Empty) {
        $exchangeService.UseDefaultCredentials = $false
        $exchangeService.Credentials = $Credential.GetNetworkCredential()   
    } else {
        $exchangeService.UseDefaultCredentials = $true
    }
    $exchangeService.AutodiscoverUrl($Mailbox)
    $Script:exchangeService = $exchangeService
    $Script:connections.Add(
        $exchangeService,
        $Mailbox
    )
    $exchangeService
}
