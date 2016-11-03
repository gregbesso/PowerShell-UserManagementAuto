<#
###
The purpose of this PowerShell scripting tool is to forward a user's email to their manager
#

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
# created 10/25/2016, Greg Besso
# last modified 10/25/2016, Greg Besso
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

function EmailForwardToManager(){
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
                $getManager = Invoke-Command -Session $sessionDC -ScriptBlock {
                    # get input from function calling remote session
                    Param ($SamAccountName)

                    # do stuff...
                    Import-Module ActiveDirectory
                    $getManager = (Get-ADUser (Get-ADUser $SamAccountName -Properties Manager).Manager).SamAccountName
                    $getManager

                } -ArgumentList $SamAccountName                
                $sessionDC | Remove-PSSession


                If ($getManager.Length -gt 1) {
                    #connect to Exchange
                    $getExServers = Get-ExchangeServersInSite
                    $getExch = $getExServers[0].Name
                    $sessionExchange = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$getExch/PowerShell/" -Authentication Kerberos
                    Import-PSSession $sessionExchange -DisableNameChecking -AllowClobber | Out-Null
  
                    $getMailbox = Get-Mailbox -Identity "$getManager"
                    $managerMail = $getMailbox.PrimarySmtpAddress

                    Set-Mailbox -Identity "$SamAccountName" -ForwardingSMTPAddress $Null
                    Set-Mailbox -Identity "$SamAccountName" -ForwardingAddress $Null
                    Set-Mailbox -Identity "$SamAccountName" -ForwardingAddress "$managerMail"

                    $checkThis = Get-Mailbox -Identity "$SamAccountName"
                    $checkThis = $checkThis.ForwardingAddress  
                    
                    $checkThis = (Get-Mailbox $checkThis).PrimarySmtpAddress
                                      
                    $sessionExchange | Remove-PSSession
                } Else { 
                    $checkThis = $managerMail = "" 
                    $thisEmailBody = "<body><div style='font-family: calibri;'>Hello there, <br/><br/>The user $SamAccountName does not have a manager specified. This user's email will need to be manually forwarded to someone. <br/><br/> Thanks, <br/>Virtual assistant</div></body>"
                    send-mailmessage `
                    -to "You@yourdomain.com" `
                    -from "noreply@yourdomain.com" `
                    -subject "Offboarding user missing manager field..." `
                    -body $thisEmailBody `
                    -smtpserver "YourSMTPinternalServer" `
                    -BodyAsHtml

                }

                If ($checkThis.Length -gt 1) {
                    Return "Completed"
                } Else {
                    Return "In Progress"
                }
            }
            
        } Catch {

        }
    }
    End {}
}