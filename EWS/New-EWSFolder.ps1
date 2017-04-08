function New-EWSFolder {
    [OutputType('Microsoft.Exchange.WebServices.Data.Folder')]
    [CmdletBinding(
            DefaultParameterSetName = 'byName'
    )]
    param (
        [Parameter(
                Mandatory,
                ValueFromPipelineByPropertyName
        )]
        [string]$DisplayName,

        [Parameter(
                ValueFromPipelineByPropertyName
        )]
        [ValidateSet(
                'Calendar',
                'Contacts',
                'Tasks',
                'Mail'
        )]
        [string]$FolderType = 'Mail',

        [Parameter(
                ParameterSetName = 'byId',
                ValueFromPipelineByPropertyName,
                Mandatory
        )]
        [Alias('Id')]
        [Microsoft.Exchange.WebServices.Data.FolderId]$ParentId,

        [Parameter(
                ParameterSetName = 'byName',
                Mandatory
        )]
        [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]$ParentName,

        [Parameter(
                ValueFromPipelineByPropertyName
        )]
        [Microsoft.Exchange.WebServices.Data.ExchangeService]$Service = $Script:exchangeService,
        
        [Microsoft.Exchange.WebServices.Data.Mailbox]$Mailbox

    )

    begin {
        $type = 'Microsoft.Exchange.WebServices.Data.{0}Folder' -f "$(
            if ($FolderType -ne 'Mail') {
            $FolderType
            }
        )"
    }

    process {
        if (-not $Service) {
            return
        }

        if ($PSCmdlet.ParameterSetName -eq 'byName') {
            if (-not $PSBoundParameters.ContainsKey('Mailbox')) {
                if (-not ($Mailbox = $Connections[$Service])) {
                    $Mailbox = $Script:Mailbox
                }
            }

            $folder = New-Object Microsoft.Exchange.WebServices.Data.FolderId $ParentName, $Mailbox
            $parentId = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service, $ParentName).Id
        }

        $newFolder = New-Object $Type $Service
        if ($FolderType -eq 'Mail') {
            $newFolder.FolderClass = 'IPF.Note'
        }
        $newFolder.DisplayName = $DisplayName
        $newFolder.Save($parentId)
    
        $name = [Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName
        $root = [Microsoft.Exchange.WebServices.Data.Folder]::Bind(
            $Service,
            $parentId
        )

        $filter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo -ArgumentList @(
            $name,
            $DisplayName
        )
        try {
            $root.FindFolders($filter,1).Folders[0]
        } catch {}
    }
    end {
        if (-not $Service) {
            Write-Warning 'No connection defined. Use Connect-EWSService first!'
            return
        }
    }
}
