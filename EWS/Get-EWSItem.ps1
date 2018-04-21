function Get-EWSItem {
    [OutputType('Microsoft.Exchange.WebServices.Data.Item')]
    [CmdletBinding(
            DefaultParameterSetName = 'byId'
    )]
    param (

        [Parameter(
                Position = 0
        )]
        [string]$Filter,

        [Parameter(
                ParameterSetName = 'byId',
                ValueFromPipelineByPropertyName,
                Mandatory
        )]
        [Microsoft.Exchange.WebServices.Data.FolderId]$Id,

        [Parameter(
                ParameterSetName = 'byName',
                Mandatory
        )]
        [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]$Name,
    
        [Parameter(
                ValueFromPipelineByPropertyName
        )]
        [Microsoft.Exchange.WebServices.Data.ExchangeService]$Service = $Script:exchangeService,

        [Microsoft.Exchange.WebServices.Data.BasePropertySet]$PropertySet,

        [ValidateRange(1,1000)]
        [Int]$PageSize = 100,

        [Int]$Limit
    )

    process {
        if (-not $Service) {
            return
        }

        $secondParameter = switch ($PSCmdlet.ParameterSetName) {
            byId {
                $Id
            }
            byName {
                $Name
            }
        }
        
        $folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind(
            $Service,
            $secondParameter
        )

        if ($Limit -and $Limit -lt $PageSize) {
            $PageSize = $Limit
        }

        $view = New-Object Microsoft.Exchange.WebServices.Data.ItemView $PageSize, 0
        do {   
            if ($Filter) {
                $list = $Folder.FindItems($Filter, $view)
            }
            else {
                $list = $Folder.FindItems($view)
            }
        
            if ($PropertySet -and $list.TotalCount) {
                $set = New-Object Microsoft.Exchange.WebServices.Data.PropertySet $PropertySet
                $Service.LoadPropertiesForItems(
                    $list,
                    $set
                ) | Out-Null
            }
            $list
            $view.Offset += $list.Items.Count
            if ($view.Offset -ge $Limit) {
                break
            }
        } while ($list.MoreAvailable)
    }

    end {
        if (-not $Service) {
            Write-Warning 'No connection defined. Use Connect-EWSService first!'
        }
    }
}
