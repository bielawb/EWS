function Get-EWSFolder  {
    [OutputType('Microsoft.Exchange.WebServices.Data.Folder')]
    [CmdletBinding()]
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
                        throw "Root folder must belong to collection of WellKnownFolderName: $_"
                    }
        })]
        [string]$Path,
        [Microsoft.Exchange.WebServices.Data.ExchangeService]$Service = $Script:exchangeService,
        [Microsoft.Exchange.WebServices.Data.Mailbox]$Mailbox
    )
    
    if (-not $Service) {
        Write-Warning 'No connection defined. Use Connect-EWSService first!'
        return
    }

    if (-not $PSBoundParameters.ContainsKey('Mailbox')) {
        if (-not ($Mailbox = $connections[$Service])) {
            $Mailbox = $Script:Mailbox
        }
    }
    $name = [Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName
    $rootFolderName, $remaingPath = $Path -split '\\|/|:|\|'
    try {
        $rootFolder = New-Object Microsoft.Exchange.WebServices.Data.FolderId $rootFolderName, $Mailbox
    } catch {
        throw "Unexpected error occured: $_"
    }
    $root = [Microsoft.Exchange.WebServices.Data.Folder]::Bind(
        $Service,
        $rootFolder
    )

    foreach ($item in $remaingPath) {
        $filter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo -ArgumentList @(
            $name,
            $item
        )
        $lastItem = 
            New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring -ArgumentList @(
                $name,
                $item
            )
        try {
            $new = $root.FindFolders($filter,1).Folders[0]
            if ($new) {
                $root = $new
            } else {
                $view = New-Object Microsoft.Exchange.WebServices.Data.FolderView 10, 0
                do {
                    $list = $root.FindFolders($lastItem,$view)
                    $list
                    $view.Offset += $list.Items.Count
                } while ($list.MoreAvailable)
            }
        } catch {
            Write-Error "Could not find subfolder: $item"
            $root
            return
        }
    }
    $root
}
