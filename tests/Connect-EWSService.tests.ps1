Import-Module -Name $PSScriptRoot\..\EWS\EWS.psd1 -ArgumentList useScript

$testMailbox = 'test@outlook.com'
$testServiceUrl = 'https://test.it/EWS/Exchange.asmx'

Describe 'Testing function Connect-EWSService' {
    Context 'Error handling' {
        BeforeAll {
            Mock -CommandName New-Object -MockWith {
                $broken = [PSCustomObject]@{
                    UseDefaultCredentials = $true
                    Credentials = [pscredential]::Empty
                    Url = ''
                    RequestedServerVersion = $ArgumentList[0]
                }

                Add-Member -InputObject $broken -MemberType ScriptMethod -Name AutodiscoverUrl -Value {
                    throw 'Oops?'
                } -PassThru
            } -ModuleName EWS
        }

        It 'Should write error and return when default ErrorAction is used' {
            $result = Connect-EWSService -Mailbox $testMailbox 2>&1
            $result.Exception.Message | Should -Match Oops
        }

        It 'Should throw when ErrorAction Stop is used' {
            {
                Connect-EWSService -Mailbox $testMailbox -ErrorAction Stop
            } | Should -Throw -ExpectedMessage Oops
        }

        It 'Should use Write-Error when error happens' {
            Mock -CommandName Write-Error -ModuleName EWS
            $null = Connect-EWSService -Mailbox $testMailbox
            Assert-MockCalled -CommandName Write-Error -Times 1 -Exactly -Scope It -ModuleName EWS
        }
    }

    Context 'Running with correct parametrs' {
        BeforeAll {
            Mock -CommandName New-Object -MockWith {
                $mockObject = [PSCustomObject]@{
                    UseDefaultCredentials = $true
                    Credentials = [pscredential]::Empty
                    Url = ''
                    RequestedServerVersion = $ArgumentList[0]
                }

                Add-Member -InputObject $mockObject -MemberType ScriptMethod -Name AutodiscoverUrl -Value {
                    param (
                        [String]$Mailbox,
                        [scriptblock]$AutodiscoverRedirectionUrlValidationCallback
                    )
                    Set-Content -Path testdrive:\$Mailbox -Value (& $AutodiscoverRedirectionUrlValidationCallback)
                    $this.Url = 'discovered'
                } -PassThru
            } -ModuleName EWS

            Mock -CommandName Write-Error
        }

        AfterEach {
            if (Test-Path -LiteralPath testdrive:\$testMailbox) {
                Remove-Item -Force -LiteralPath testdrive:\$testMailbox
            }
        }

        It 'Should use default credentials if no credentials were passed' {
            $result = Connect-EWSService -Mailbox $testMailbox
            $result.UseDefaultCredentials | Should -BeTrue
        }

        It 'Should inject specified credentials if user passed them' {
            $credentials = [pscredential]::new('luser', (ConvertTo-SecureString -AsPlainText -Force -String lpass))
            $result = Connect-EWSService -Mailbox $testMailbox -Credential $credentials
            $result.UseDefaultCredentials | Should -BeFalse
            $result.Credentials.UserName | Should -BeExactly luser
            $result.Credentials.Password | Should -BeExactly lpass
        }

        It 'Should run autodiscovery if ServiceUrl is not specified' {
            $result = Connect-EWSService -Mailbox $testMailbox
            "testdrive:\$testMailbox" | Should -Exist
            "testdrive:\$testMailbox" | Should -FileContentMatch False
            $result.Url | Should -BeExactly discovered
        }

        It 'Should run autodiscovery and allow redirecting if AllowRedirect switch is used' {
            $result = Connect-EWSService -Mailbox $testMailbox -AllowRedirect
            "testdrive:\$testMailbox" | Should -Exist
            "testdrive:\$testMailbox" | Should -FileContentMatch True
            $result.Url | Should -BeExactly discovered
        }

        It 'Should not run autodiscovery when ServiceUrl is specified' {
            $result = Connect-EWSService -Mailbox $testmailbox -ServiceUrl $testServiceUrl
            "testdrive:\$testmailbox" | Should -Not -Exist
            $result.Url | Should -BeExactly $testServiceUrl
        }

        It 'Should use specified Exchange version' {
            $result = Connect-EWSService -Mailbox $testMailbox -Version Exchange2013
            $result.RequestedServerVersion | Should -BeExactly Exchange2013
        }
    }
}