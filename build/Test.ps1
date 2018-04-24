Import-Module -Name Pester
Import-Module -Name .\EWS
Import-Module -Name .\tests\Assertions.psm1
Invoke-Pester -Script .\tests\ -EnableExit