function Connect-EWSService {
    [OutputType('Microsoft.Exchange.WebServices.Data.ExchangeService')]
    [CmdletBinding()]
    param (
		
        # Mailbox that command will connect to.
        [Parameter(
                Mandatory
        )]
        [String]$Mailbox,
        

        # Version of exchange server that we will connect to.
        [Microsoft.Exchange.WebServices.Data.ExchangeVersion]
        $Version = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1,

        # Credentials that will be used to connect to mailbox.
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
