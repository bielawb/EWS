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
    $placeholderPattern = '{{.*Placeholder.*}}'
    
    Context 'About topics' {
        $aboutTestCases = @(
            @{
                Name = 'global'
                File = 'about_EWS'
            }
            @{
                Name = 'advanced query syntax'
                File = 'about_AQS'
            }
        )

        It 'Should contain <Name> about topic' {
            param ($Name, $File)
            ".\EWS\en-US\$File.help.txt" | Should -Exist
        } -TestCases $aboutTestCases

        It 'Should not have placeholders in <Name> about topic' {
            param ($Name, $File)
            ".\EWS\en-US\$File.help.txt" | Should -Not -FileContentMatch $placeholderPattern
        } -TestCases $aboutTestCases
    }

    Context 'Commands help' {
        foreach ($command in $module.ExportedCommands.Keys) {
            $help = Get-Help -Name $command -ErrorAction SilentlyContinue
            
            $testCases = @(
                @{
                    Section = 'Description'
                }
                @{
                    Section = 'Synopsis'
                }
                @{
                    Section = 'Examples'
                }
            )

            It "Command $command should have help/<Section>" {
                param ($Section)
                $help.$Section | Should -Not -BeNullOrEmpty
            } -TestCases $testCases

            It "Command $command <Section> should not contain placeholders" {
                param ($Section)
                $help.$Section | Should -Not -Match $placeholderPattern
            } -TestCases $testCases

            It "Command $command should have at least two examples in help" {
                $help.examples.example.Count | Should -BeGreaterOrEqual 2
            }

            $parametersTestCases = @(
                foreach ($parameter in $help.parameters.parameter) {
                    @{
                        Name = $parameter.Name
                        Description = $parameter.Description
                    }
                }
            )
            
            It 'Parameter <Name> should have description' {
                param ($Name, $Description)
                $Description | Should -Not -BeNullOrEmpty
            } -TestCases $parametersTestCases

            It 'Description of parameter <Name> should not contain placeholders' {
                param ($Name, $Description)
                $Description | Should -Not -Match $placeholderPattern
            } -TestCases $parametersTestCases
        }
    }
}