<#
###
The purpose of this PowerShell scripting tool is to backup a user's various settings, such as...
Lync user and RGS settings, Exchange and AD group membership, etc...

###
Copyright (c) 2016 Greg Besso

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
#>
#
# created 10/24/2016, Greg Besso
# last modified 10/24/2016, Greg Besso
#
#
function Get-ExchangeServersInSite {
    $ADSite = [System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]
    $siteDN = $ADSite::GetComputerSite().GetDirectoryEntry().distinguishedName
    $configNC=([ADSI]"LDAP://RootDse").configurationNamingContext
    $search = new-object DirectoryServices.DirectorySearcher([ADSI]"LDAP://$configNC")
    $objectClass = "objectClass=msExchExchangeServer"
    $version = "versionNumber>=1937801568"
    $site = "msExchServerSite=$siteDN"
    $search.Filter = "(&($objectClass)($version)($site))"
    $search.PageSize=1000
    [void] $search.PropertiesToLoad.Add("name")
    [void] $search.PropertiesToLoad.Add("msexchcurrentserverroles")
    [void] $search.PropertiesToLoad.Add("networkaddress")
    $search.FindAll() | %{
        New-Object PSObject -Property @{
            Name = $_.Properties.name[0]
            FQDN = $_.Properties.networkaddress |
                %{if ($_ -match "ncacn_ip_tcp") {$_.split(":")[1]}}
            Roles = $_.Properties.msexchcurrentserverroles[0]
        }
    }
}

function BackupExistingConfiguration(){
    Param (
        [Parameter(Mandatory=$True,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        [string]$SamAccountName
    )

    Begin {}
    Process {
    
        Try {
            $samLength = $SamAccountName.Length
            If ($samLength -gt 1) {                
                $error.Clear()

                #connect to domain controller
                $getDC = ([ADSI]"LDAP://RootDSE").dnshostname
                If (!($sessionDC)) { $sessionDC = New-PSSession -ComputerName $getDC}
                $getInfo = Invoke-Command -Session $sessionDC -ScriptBlock {
                    # get input from function calling remote session
                    Param ($SamAccountName)

                    # do stuff...
                    Import-Module ActiveDirectory
                    $getInfo = Get-ADPrincipalGroupMembership $SamAccountName
                    $getInfo

                } -ArgumentList $SamAccountName
                $getInfo | Export-CliXml "\\server\share\UserManagementAuto\Requests\UserConfigurations\$SamAccountName-ADPrincipalGroupMembership.xml"
                $sessionDC | Remove-PSSession

                #connect to Exchange
                $getExServers = Get-ExchangeServersInSite
                $getExch = $getExServers[0].Name
                $sessionExchange = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$getExch/PowerShell/" -Authentication Kerberos
                Import-PSSession $sessionExchange -DisableNameChecking -AllowClobber | Out-Null
                $getMailbox = Get-Mailbox -Identity "$SamAccountName"
                $getUMMailbox = Get-UMMailbox -Identity "$SamAccountName"
                $getMailboxStatistics = Get-MailboxStatistics -Identity "$SamAccountName"
                $getMailbox | Export-CliXml "\\server\share\UserManagementAuto\Requests\UserConfigurations\$SamAccountName-GetMailbox.xml"
                $getUMMailbox | Export-CliXml "\\server\share\UserManagementAuto\Requests\UserConfigurations\$SamAccountName-GetUMMailbox.xml"
                $getMailboxStatistics | Export-CliXml "\\server\share\UserManagementAuto\Requests\UserConfigurations\$SamAccountName-GetMailboxStatistics.xml"
                $sessionExchange | Remove-PSSession



                $sessionLync = New-PSSession -ConnectionURI "https://<yourlyncServerOrPoolURL/OcsPowershell" -Authentication NegotiateWithImplicitCredential
                Import-PsSession $sessionLync -DisableNameChecking -AllowClobber | Out-Null
                $getCsUser = Get-CSUser -Identity "$SamAccountName"
                $getCsRgs = Get-CsRgsAgentGroup | Where-Object {$_.AgentsByUri -like "*$SamAccountName*"}
                $getCsUser | Export-CliXml "\\server\share\UserManagementAuto\Requests\UserConfigurations\$SamAccountName-GetCSUser.xml"
                $getCsRgs | Export-CliXml "\\server\share\UserManagementAuto\Requests\UserConfigurations\$SamAccountName-GetCSRgs.xml"
                $sessionLync | Remove-PSSession



                # if errors, send an email...
                If ($errors.Exception -ne $null) {                    
                    Return "In Progress"
                } Else {
                    Return "Completed"
                }
            }
            
        } Catch {

        }
    }
    End {}
}