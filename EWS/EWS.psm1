param (
    [ValidateSet(
        'useScriptBlock',
        'useScript'
    )]
    [String]$DotSourceMethod = 'useScriptBlock'
)

foreach ($file in Get-ChildItem -Path $PSScriptRoot\*.ps1) {
    if ($DotSourceMethod -eq 'useScriptBlock') {
        # Avoiding bug in WMF5/Install-Module
        # See: https://constantinekokkinos.com/articles/64/troubleshooting-powershell-and-net-when-error-messages-are-not-enough
        $ExecutionContext.InvokeCommand.InvokeScript(
            $false, 
            (
                [scriptblock]::Create(
                    [io.file]::ReadAllText(
                        $file.FullName,
                        [Text.Encoding]::UTF8
                    )
                )
            ), 
            $null, 
            $null
        )
    } else {
        . $file.FullName
    }
}

$connections = @{}