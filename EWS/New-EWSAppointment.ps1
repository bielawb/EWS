function New-EWSAppointment {
    [OutputType('Microsoft.Exchange.WebServices.Data.Appointment')]
    [CmdletBinding(
            DefaultParameterSetName = 'inline'
    )]
    param (
        [string[]]$Required,

        [string[]]$Optional,

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

        [string]$Location,

        [DateTime]$Start,

        [TimeSpan]$Duration,

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
            if (-not $IsHtml -and $BodyType -eq 'HTML') {
                $body = $data | ConvertTo-Html | Out-String
            } else {
                $body = $data | Out-String
            }
        }
        $appointment = New-Object Microsoft.Exchange.WebServices.Data.Appointment $Service -Property @{
            Subject = $Subject
            Body = $body
            Start = $Start
            End = $Start + $Duration
            Location = $Location
        }
        $appointment.Body.BodyType = $BodyType
        foreach ($file in $Attachment) {
            $null = $appointment.Attachments.AddFileAttachment($file)
        }
        foreach ($attendee in $Required) {
            $null = $appointment.RequiredAttendees.Add(
                $attendee
            )
        }
        foreach ($attendee in $Optional) {
            $null = $appointment.OptionalAttendees.Add(
                $attendee
            )
        }
        if ($Required -or $Optional) {
            $method = [Microsoft.Exchange.WebServices.Data.SendInvitationsMode]::SendToAllAndSaveCopy
        } else {
            $method = [Microsoft.Exchange.WebServices.Data.SendInvitationsMode]::SendToNone
        }
        $appointment.Save($method)
        $appointment
    }
}
