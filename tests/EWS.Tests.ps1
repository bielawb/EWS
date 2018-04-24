$prefix = 'Global tests of EWS module'
$module = Get-Module -Name EWS

Describe "$prefix - manifest" {
    Context 'Module properties' {
        It 'Should have description' {
            $module.Description | Should -Not -BeNullOrEmpty
        }

        It 'Should have an author' {
            $module.Author | Should -Not -BeNullOrEmpty
        }

        It 'Should have LicenseUri' {
            $module.LicenseUri | Should -Not -BeNullOrEmpty
        }

        It 'Should have ProjectUri' {
            $module.ProjectUri | Should -Not -BeNullOrEmpty
        }

        It 'Uris should point to github project' {
            $githubUri = 'https://github.com/bielawb/EWS*'
            $module.LicenseUri | Should -BeLike $githubUri
            $module.ProjectUri | Should -BeLike $githubUri
        }
    }

    Context 'Dot-sourced files' {
        foreach ($file in Get-ChildItem -Path .\EWS\*.ps1) {
            $fileName = $file.Name
            $baseName = $file.BaseName
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

Describe "$prefix - Help" {
    Context 'About topics' {
        It 'Should contain global about topic' {
            '.\EWS\en-US\about_EWS.help.txt' | Should -Exist
        }

        It 'Should contain advanced query syntax about topic' {
            '.\EWS\en-US\about_AQS.help.txt' | Should -Exist
        }
    }

    Context 'Commands help' {
        foreach ($command in $module.ExportedCommands.Keys) {
            $help = Get-Help -Name $command -ErrorAction SilentlyContinue

            It "Command $command should have help/Synopsis" {
                $help.Synopsis | Should -Not -BeNullOrEmpty
            }

            It "Command $command should have help/Description" {
                $help.Description | Should -Not -BeNullOrEmpty
            }

            It "Command $command should have at least two examples in help" {
                $help.examples.example.Count | Should -BeGreaterOrEqual 2
            }
        }
    }
}