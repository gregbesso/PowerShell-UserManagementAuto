<#
###
The purpose of this PowerShell scripting tool is to...
# remove a user from all domain security and distribution groups
# Result is user only a member of a new shell group called "Disabled Offboarding".
# This group has to be created in AD before the script will work properly.
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
# created 10/24/2016, Greg Besso
# last modified 10/24/2016, Greg Besso
#
#

function RemoveUserFromGroups(){
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

                    # first change user's primary group so can be removed from domain users...
                    Add-ADGroupMember -Identity "Disabled Offboarding" -Member "$SamAccountName"
                    $group = Get-ADGRoup -Identity "Disabled Offboarding"
                    $groupSID = $group.SID
                    [int]$groupID = $groupSID.Value.Substring($groupSID.Value.LastIndexOf("-")+1)
                    Get-ADUser -Identity "$SamAccountName" | Set-ADObject -Replace @{primaryGroupID="$groupID"}


                    # then remove user from all real groups...
                    $ADgroups = Get-ADPrincipalGroupMembership -Identity "$SamAccountName" | where {$_.Name -ne "Disabled Offboarding"}
                    Remove-ADPrincipalGroupMembership -Identity "$SamAccountName" -MemberOf $ADgroups -Confirm:$false

                    # then check to ensure all is set...
                    $getGroups = Get-ADPrincipalGroupMembership -Identity "$SamAccountName"
                    $getName = $getGroups.Name
                    $getName

                } -ArgumentList $SamAccountName
                $sessionDC | Remove-PSSession

                # if all set update task to completed...
                If ($getInfo -eq "Disabled Offboarding") {
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