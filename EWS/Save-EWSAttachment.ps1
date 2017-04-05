function Save-EWSAttachment {
    [CmdletBinding(
            SupportsShouldProcess,
            ConfirmImpact = 'Low'
    )]
    param (
        [string]$Path = '.',
        [Parameter(
                ValueFromPipelineByPropertyName,
                ValueFromPipeline
        )]
        [Microsoft.Exchange.WebServices.Data.FileAttachment[]]$Attachment,
        [Parameter(
                ValueFromPipelineByPropertyName
        )]
        [Microsoft.Exchange.WebServices.Data.ExchangeService]$Service = $script:exchangeService
    )

    process {
        if (-not $Service) {
            return
        }

        foreach ($singleAttachment in $Attachment) {
            $fullPath = Join-Path -Path $Path -ChildPath $singleAttachment.Name
            if ($PSCmdlet.ShouldProcess(
                    $singleAttachment.Name,
                    "Save to $Path"
            )) { 
                $singleAttachment.Load(
                    $fullPath
                )
            }
        }
    }
    
    end {
        if (-not $Service) {
            Write-Warning 'No connection defined. Use Connect-EWSService first!'
        }
    }
}
