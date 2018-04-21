function Get-EWSService {
    [OutputType('Microsoft.Exchange.WebServices.Data.ExchangeService')]
    [CmdletBinding()]
    Param(
        
    )
    if ($Script:exchangeService){
        $Script:exchangeService
    } else {
        Write-Information -MessageData "No Active Session Found" -InformationAction Continue
    }
}