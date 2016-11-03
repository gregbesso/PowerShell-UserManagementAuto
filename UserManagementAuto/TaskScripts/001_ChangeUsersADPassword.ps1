<#
###
The purpose of this PowerShell scripting tool is to reset an AD account's password to something known
so that IT can login as that account during off-boarding processes.


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
# created 10/28/2016, Greg Besso
# last modified 10/28/2016, Greg Besso
#
#

function ChangeUsersADPassword(){
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

                    Try {
                        Set-ADAccountPassword -Identity "$SamAccountName" -Reset -NewPassword (ConvertTo-SecureString -AsPlainText "YourGenericPWchoiceHere" -Force)
                        $getResults = "Completed"
                    } Catch {
                        $getResults = "In Progress"
                    }
                    $getResults
                } -ArgumentList $SamAccountName
                $sessionDC | Remove-PSSession

                # if all set update task to completed...
                If ($getInfo -eq "Completed") {
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