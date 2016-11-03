#
#
#
#
#
#
#
#
# import the Offboarding global variables, to avoid having to define them 
# in multiple places by the handful of script files. probably the only place with a manually specified path...
#
. "\\server\share\UserManagementAuto\ServerScripts\OffboardingImports.ps1" 


# function that gets existing task templates...
function Get-UserMgtTasks(){
    Param (
        [Parameter(Mandatory=$True,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        [string]$whichList
    )

    Begin {}
    Process {
    
        Try {
            # load existing info - web, site, lists, fields for specific list...
            $web = $context.Web
            $lists = $web.Lists
            $list = $web.Lists.GetByTitle("$whichList");
            $qry = [Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery()
            $items = $list.GetItems($qry)
            $context.Load($web)
            $context.Load($lists)
            $context.Load($list)
            $context.Load($items)
            $context.ExecuteQuery()
                
            $newArray = @()
            ForEach ($item in $Items) { 
            
                $thisSam = $item.FieldValues.SamAccountName
                $thisStatus = $item.FieldValues.Status
                $thisTaskType = $item.FieldValues.TaskType

                If (($thisStatus -ne 'Completed') -And ($thisTaskType -eq 'Automatic')) {              
                    $newItem = New-Object PSObject -Property @{
                        ID=$item.FieldValues.ID
                        Title=$item.FieldValues.Title
                        AssignedTo=$item.FieldValues.AssignedTo
                        StartDate=$item.FieldValues.StartDate
                        TaskScript=$item.FieldValues.TaskScript
                        TaskType=$item.FieldValues.TaskType
                        TaskDay=$item.FieldValues.TaskDay
                        TaskOrder=$item.FieldValues.TaskOrder
                        SamAccountName=$item.FieldValues.SamAccountName
                        EmailSubject=$item.FieldValues.EmailSubject
                        CompletedDate=$item.FieldValues.CompletedDate         
                    }
                    $newArray += $newItem
                }
            }

            Return $newArray

        } Catch {
            Write-Warning "Error occurred: $_.Exception.Message"
        }
    }
    End {}
}

# function that creates any missing columns in a specified list...
function Update-UserMgtTasks(){
    Param (
        [Parameter(Mandatory=$True,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        [string]$whichList,
        [string]$whichStatus,
        [int]$whichID
    )


    Begin {}
    Process {
    
        Try {
            If ($whichID -gt 0) {
                # load existing info - web, site, lists, fields for specific list...
                $web = $context.Web
                $lists = $web.Lists
                $list = $web.Lists.GetByTitle("$whichList");
                $qry = [Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery()
                $items = $list.GetItems($qry)
                $item = $items.GetByID($whichID);                
                $item["Status"] = "$whichStatus";

                If ($whichStatus -eq 'Completed') {
                $item["CompletedDate"] = Get-Date;
                }

                $item.Update()
                $context.ExecuteQuery()
            }

        } Catch {
            Write-Warning "Error occurred: $_.Exception.Message"
        }
    
    
    
    }
    End {}




}

# function that does the stuff...
function Update-PerformTasks() {

    [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client")
    [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime")
    #Bind to site collection
        $Context = New-Object Microsoft.SharePoint.Client.ClientContext($global:usrMgmtSpWeb)
        $Creds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($global:usrMgmtSpUser,$global:usrMgmtCredSpOnline)
        $Context.Credentials = $Creds


    $getTasks = Get-UserMgtTasks -whichList "$global:usrMgmtTasksListName"
    $getTasks = $GetTasks | Sort-Object TaskDay, TaskOrder

    # loop through all the tasks
    $getTasks | ForEach-Object {
        Try {
            $thisTask = $_
            $daysTillReady = ($thisTask.StartDate - (Get-Date)).Days    

            If ($daysTillReady -lt 1) {
                $thisSam = $thisTask.SamAccountName
                $thisSubject = $thisTask.EmailSubject
                $thisScript = $thisTask.TaskScript
                $thisID = $thisTask.ID
                $thisTitle = $thisTask.Title
        

                # parse the function from within the script that is to be loaded (name matches file name)...
                $thisFunction = $thisScript.Replace(".ps1","")
                $thisFunction = $thisFunction.Split("_")
                $thisFunction = $thisFunction[$thisFunction.Length-1]

                # load the script that contains the function to be executed...
                . "$thisScript"


                # form the expression that needs to beg called...
                $doThis = "$thisFunction"
                $doThis += " -SamAccountName ""$thisSam"""

                # run the function stored in the loaded script, with parameters sent to it...
                $getResult = Invoke-Expression $doThis
                If ($getResult -ne "Completed") {
                    $getResult = $getResult[$getResult.Length-1]
                }

                # send an email once done...
                If ($getResult -eq 'Completed') {
                    Send-MailMessage -to "yourTicketingsystem@yourdomain.com" -from 'noreply@yourdomain.com' -subject "$thisSubject" -body "For the off-boarding process of: $thisSam, the following task is now completed: $thisTitle" -smtpserver 'yourSMTPserver'
                } Else {
                    Send-MailMessage -to "you@yourdomain.com" -from 'noreply@yourdomain.com' -subject "$thisSubject" -body "For the off-boarding process of: $thisSam, the following task had an issue: $thisTitle" -smtpserver 'yourSMTPserver'
                }

                # add section to update task as completed...
                Update-UserMgtTasks -whichList "$global:usrMgmtTasksListName" -whichID $thisID -whichStatus "$getResult"
            }
        } Catch {
            #do something
        }
    }



}
