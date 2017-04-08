function Move-EWSItem {
    [CmdletBinding(
            ConfirmImpact = 'Medium',
            SupportsShouldProcess
    )]
    param (
        [Parameter(
                Mandatory
        )]
        [ValidateScript({
                    $root, $rest = $_ -split '\\|/|:|\|'
                    try { 
                        [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]$root 
                        $true
                    } catch {
                        throw "Root folder must belong to WellKnownFolderName: $_"
                    }
        })]
        [string]$DestinationPath,

        [Parameter(
                ValueFromPipeline,
                Mandatory
        )]
        [Microsoft.Exchange.WebServices.Data.Item]$Item,
        
        [Parameter(
                ValueFromPipelineByPropertyName
        )]
        [Microsoft.Exchange.WebServices.Data.ExchangeService]$Service = $Script:exchangeService

    )

    begin {

        $Mailbox = $Connections[$Service]
        $name = [Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName
        $rootName, $remaingPath = $DestinationPath -split '\\|/|:|\|'
        try {
            $rootId = New-Object Microsoft.Exchange.WebServices.Data.FolderId $rootName, $Mailbox
        } catch {
            throw "Unexpected error occured: $_"
        }
        $root = [Microsoft.Exchange.WebServices.Data.Folder]::Bind(
            $Service,
            $rootId
        )

        foreach ($pathItem in $remaingPath) {
            $filter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo -ArgumentList @(
                $name,
                $pathItem
            )
            try {
                $new = $root.FindFolders($filter,1).Folders[0]
                if ($new) {
                    $root = $new
                }
            } catch {
                Write-Error "Could not find sub-folder: $pathItem"
                $root
                return
            }
        }
        $destination = $root
    }

    process {
        if (-not $Service) {
            return
        }

        if ($PSCmdlet.ShouldProcess(
                $Item.Subject,
                "Move to $DestinationPath"
        )) { 
            $Item.Move($destination.Id)
        }
    }
    end {
        if (-not $Service) {
            Write-Warning 'No connection defined. Use Connect-EWSService first!'
            return
        }
    }
}
