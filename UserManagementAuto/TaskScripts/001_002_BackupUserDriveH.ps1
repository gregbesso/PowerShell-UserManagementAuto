<#
###
The purpose of this PowerShell scripting tool is to...
# backup a user's network drive contents to a separate network share.
# source typically is something like \\server\share\Users\$samAccountName
# destination typically is something like \\server\share$\exUsers\$samAccountName
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
# created 10/18/2016, Greg Besso
# last modified 10/24/2016, Greg Besso
#
#

function BackupUserDriveH(){
    Param (
        [Parameter(Mandatory=$True,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        [string]$SamAccountName
    )

    Begin {}
    Process {
    
        Try {
            $samLength = $SamAccountName.Length
            If ($samLength -gt 1) {
                $sourceFolder = "\\server\share\users\$samAccountName"
                $destinationFolder = "\\server\share\TermBackup\$samAccountName"

                # create the destination folder...
                If (-Not (Test-Path $destinationFolder)) { New-Item -ItemType Directory -Force -Path $destinationFolder }

                # copy the contents...
                $error.Clear()
                Copy-Item $sourceFolder $destinationFolder -Recurse -Force -errorVariable errors -PassThru -ErrorAction SilentlyContinue
                
                # if errors, send an email...
                If ($errors.Exception -ne $null) {
                    $thisEmailBody = "<body><div style='font-family: calibri;'>For the off-boarding process of: $thisSam, the task for backing up the network drive had an issue... <br/><br/>Exception: $error<br/><br/> Thanks, <br/>Your faithful virtual assistant</div></body>"
                    send-mailmessage `
                    -to "you@yourdomain.com" `
                    -from "noreply@yourdomain.com" `
                    -subject "Offboarding error" `
                    -body $thisEmailBody `
                    -smtpserver "YourSMPTinternalServer" `
                    -BodyAsHtml

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