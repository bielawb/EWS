param (
    [ValidateSet(
        'useScriptBlock',
        'useScript'
    )]
    [String]$DotSourceMethod = 'useScriptBlock'
)

foreach ($file in Get-ChildItem -Path $PSScriptRoot\*.ps1) {
    if ($DotSourceMethod -eq 'useScriptBlock') {
        . (
            [scriptblock]::Create(
                [System.IO.File]::ReadAllText(
                    $file.FullName,
                    [System.Text.Encoding]::UTF8
                )
            )
        )
    } else {
        . $file.FullName
    }
}

$connections = @{}