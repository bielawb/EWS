Import-Module -Name Pester
Import-Module -Name .\EWS
Invoke-Pester -Script .\tests\ -EnableExit