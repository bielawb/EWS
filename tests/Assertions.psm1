Add-AssertionOperator -Name ExportCommand -Test {
    param ($Module, $CommandName, [switch]$Negate)
    $out = if ($Module -is [psmoduleinfo]) {
        $Module.ExportedCommands.ContainsKey($CommandName)
    } else {
        Get-Command -Module $Module -Name $CommandName
    }
    $failure = "{{Module $($Module.Name)}} {0} export {{$CommandName}}" -f $(
        if ($Negate) {
            $out = -not $out
            'does'
        } else {
            "doesn't"
        }
    )
    [PSCustomObject]@{
        Succeeded = $out
        FailureMessage = $failure
    }
}
