function New-EWSContact {
    [OutputType('Microsoft.Exchange.WebServices.Data.Contact')]
    param (
        [string]$GivenName,

        [string]$Surname,

        [string]$MiddleName,

        [string]$Company,

        [ValidateScript({
                    try {
                        foreach ($key in $_.Keys) {
                            [Microsoft.Exchange.WebServices.Data.PhoneNumberKey]$key 
                        }
                        $true
                    } catch {
                        throw "Error: $key - $_"
                    }
        })]
        [hashtable]$Phone,

        [ValidateScript({
                    foreach ($hash in $_) {
                        foreach ($key in $hash.Keys) {
                            if ('DisplayName','email' -notcontains $key) {
                                throw "Wrong key in e-mail hashtable: $key"
                            }
                        }
                    }
                    $true
        })]
        [hashtable[]]$Email,

        [ValidateScript({
                    try {
                        foreach ($key in $_.Keys) {
                            [Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]$key 
                        }
                    } catch {
                        throw "Error: $key - $_"
                    }

                    foreach ($hash in $_.Values) {
                        foreach ($key in $hash.Keys) {
                            if ('City','CountryOrRegion','PostalCode','State','Street' -notcontains $key) {
                                throw "Wrong key in physical address hashtable: $key"
                            }
                        }
                    }
                    $true
        })]
        [hashtable]$PhysicalAddress,
        
        [Microsoft.Exchange.WebServices.Data.FolderId]$FolderId,
        
        [Parameter(
                ValueFromPipelineByPropertyName
        )]
        [Microsoft.Exchange.WebServices.Data.ExchangeService]$Service = $script:ExchangeService
    )
    if (-not $Service) {
        Write-Warning 'No connection defined. Use Connect-EWSService first!'
        return
    }

    $contact = New-Object Microsoft.Exchange.WebServices.Data.Contact $Service -Property @{
        GivenName = $GivenName
        MiddleName = $MiddleName
        Surname = $Surname
        CompanyName = $Company
        FileAsMapping = [Microsoft.Exchange.WebServices.Data.FileAsMapping]::SurnameCommaGivenName
    }

    foreach ($key in $Phone.Keys) {
        $Contact.PhoneNumbers[$key] = $Phone.$key
    }

    $index = 1
    foreach ($hash in $Email) {
        if ($index -le 3) {
            if ($emailAddress = $hash.email) {
                if (! ($displayName = $hash.DisplayName)) {
                    $displayName = $hash.email
                }
                $contact.EmailAddresses["EmailAddress$index"] =
                    New-Object Microsoft.Exchange.WebServices.Data.EmailAddress -ArgumentList @(
                        $displayName
                        $emailAddress
                    )
                $index++
            }
        }
    }
    foreach ($key in $PhysicalAddress.Keys) {
        $contact.PhysicalAddresses[$key] = 
            New-Object Microsoft.Exchange.WebServices.Data.PhysicalAddressEntry -Property @{
                City = $PhysicalAddress[$key].City
                CountryOrRegion = $PhysicalAddress[$key].CountryOrRegion
                PostalCode = $PhysicalAddress[$key].PostalCode
                State = $PhysicalAddress[$key].State
                Street = $PhysicalAddress[$key].Street
            }
    }
    if ($FolderId) {
        $contact.Save($FolderId)
    }
    else {
        $contact.Save()
    }
    $contact
}
