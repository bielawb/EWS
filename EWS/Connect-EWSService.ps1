function Connect-EWSService {
    [OutputType('Microsoft.Exchange.WebServices.Data.ExchangeService')]
    [CmdletBinding()]
    param (		
        [Parameter(
                Mandatory
        )]
        [String]$Mailbox,

        [String]$ServiceUrl,

        [Microsoft.Exchange.WebServices.Data.ExchangeVersion]
        $Version = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1,
        
        [switch]$AllowRedirect,
        
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

    if ($ServiceUrl) {
        $exchangeService.Url = $ServiceUrl
    } else {
        try {
            $exchangeService.AutodiscoverUrl($Mailbox,{[bool]$AllowRedirect})
        } catch {
            throw "Failed to identify Url for $Mailbox - $_"
        }
    }

    $Script:exchangeService = $exchangeService
    $Script:connections.Add(
        $exchangeService,
        $Mailbox
    )
    $exchangeService
}
