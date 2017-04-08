function New-EWSMessage {
    [OutputType('Microsoft.Exchange.WebServices.Data.EmailMessage')]
    [CmdletBinding(
            DefaultParameterSetName = 'inline'
    )]
    param (
        [Parameter(
                Mandatory
        )]
        [string[]]$To,

        [string[]]$Cc,

        [string[]]$Bcc,

        [Parameter(
                Mandatory
        )]
        [string]$Subject,

        [Parameter(
                Mandatory,
                ParameterSetName = 'inline'
        )]
        [string]$Body,

        [Microsoft.Exchange.WebServices.Data.BodyType]$BodyType = 'Text',

        [string[]]$Attachment,

        [Parameter(
                ValueFromPipeline,
                Mandatory,
                ParameterSetName = 'pipe'
        )]
        $InputObject,

        [Parameter(
                ParameterSetName = 'pipe'
        )]
        [switch]$IsHtml,
        
        [Parameter(
                ValueFromPipelineByPropertyName
        )]
        [Microsoft.Exchange.WebServices.Data.ExchangeService]$Service = $script:exchangeService
    )

    begin {
        $data = New-Object Collections.ArrayList
    }

    process {
        if (-not $Service) {
            return
        }

        if ($InputObject) {
            $null = $data.Add($InputObject)
        }
    }

    end {
        if (-not $Service) {
            Write-Warning 'No connection defined. Use Connect-EWSService first!'
            return
        }

        if ($data.Count) {
            if (! $IsHtml -and $BodyType -eq 'HTML') {
                $body = $data | ConvertTo-Html | Out-String
            } else {
                $body = $data | Out-String
            }
        }
        $message = New-Object Microsoft.Exchange.WebServices.Data.EmailMessage $Service
        $message.Subject = $Subject
        $message.Body = $body
        $message.Body.BodyType = $BodyType
        foreach ($file in $Attachment) {
            $null = $message.Attachments.AddFileAttachment($file)
        }
        foreach ($recipient in $To) {
            $null = $message.ToRecipients.Add(
                $recipient    
            )
        }
        foreach ($recipient in $Cc) {
            $null = $message.CcRecipients.Add(
                $recipient
            )
        }
        foreach ($recipient in $Bcc) {
            $null = $message.BccRecipients.Add(
                $recipient
            )
        }
        $message.SendAndSaveCopy()
        $message
    }
}
