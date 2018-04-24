$prefix = 'Global tests of EWS module'

Describe "$prefix - manifest" {
    Context 'Dot-sourced files' {
        foreach ($file in Get-ChildItem -Path .\EWS\*.ps1) {
            $fileName = $file.Name
            $baseName = $file.BaseName
            $module = Get-Module -Name EWS
            It "$fileName should have name like PS function that it contains" {
                $file.Name | Should -Match ^\w+-EWS\w+\.ps1
            }
            
            It "$fileName should have name that starts with approved verb" {
                $file.Name -replace '^(.*?)-.*', '$1' | Should -BeIn (Get-Verb).Verb
            }
            
            It "$baseName should be exported as a command from the module" {
                $module | Should -Export -CommandName $baseName 
            }
        }
    }
}