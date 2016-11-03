#
# for example of loading lists and fields, https://sharepoint.stackexchange.com/questions/159749/retrieve-all-fields-from-a-list-using-powershell-csom/159751
# for example of creating columns, https://blogs.technet.microsoft.com/fromthefield/2015/03/03/office-365-create-a-list-and-add-custom-fields-using-csom/
# another example of creating columns, https://karinebosch.wordpress.com/my-articles/creating-fields-using-csom/
# for field elements info, https://msdn.microsoft.com/en-us/library/office/ms437580.aspx
# for passing variable to .ps1 file, http://stackoverflow.com/questions/5592531/how-to-pass-an-argument-to-a-powershell-script
#
# $secureStringSpOnline = Read-Host -AsSecureString "Enter your service account password"
# $secureStringSpOnline | ConvertFrom-SecureString | Out-File -FilePath "$global:usrMgmtServerSource\Cred\usrMgmtSpOnlineVV.txt"
#
# import the Offboarding global variables, to avoid having to define them 
# in multiple places by the handful of script files. probably the only place with a manually specified path...
#

# function that checks if lists need to be created...
function New-UserMgtListCheck(){
    Begin {}
    Process {
    
        Try {
            # load existing info - web, site, lists, fields for specific list...
            $web = $context.Web
            $site = $context.Site 
            $lists = $web.Lists
            $context.Load($web)
            $context.Load($site)
            $context.Load($lists)
            $context.ExecuteQuery()

            New-UserMgtList -whichList "$global:usrMgmtTemplatesListName"
            New-UserMgtList -whichList "$global:usrMgmtTasksListName"           


        } Catch {
            Write-Warning "Error occurred: $_.Exception.Message"
        }
    }
    End {}




}

# function that creates the lists if they don't exist yet...
function New-UserMgtList(){
    Param (
        [Parameter(Mandatory=$True,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        [string]$whichList
    )


    Begin {}
    Process {
    
        Try {
            # load existing info - web, site, lists, fields for specific list...
            $web = $context.Web
            $site = $context.Site 
            $lists = $web.Lists
            $context.Load($web)
            $context.Load($site)
            $context.ExecuteQuery()

            
            # if list doesn't exist, create it
                $spListCheck = 0
                ForEach ($list in $Lists) {
                    $getTitle = $list.Title
                    If ($getTitle -eq "$whichList") { $spListCheck = 1 }

                }
                If ($spListCheck -eq 0) {
                    #Create list with "custom" list template
                    $ListInfo = New-Object Microsoft.SharePoint.Client.ListCreationInformation
                    $ListInfo.Title = "$whichList"
                    $ListInfo.TemplateType = "107" #using tasks list template, others available  https://msdn.microsoft.com/en-us/library/microsoft.sharepoint.splisttemplatetype.aspx
                    $List = $Context.Web.Lists.Add($ListInfo)
                    $List.Description = $ListTitle
                    $List.OnQuickLaunch = "True"
                    $List.Update()
                    $Context.ExecuteQuery()



                    # remove built-in columns that are not needed...

                    $removeMe = $List.Fields.GetByInternalNameOrTitle("Predecessors")
                    $removeMe.DeleteObject()
                    $removeMe2 = $List.Fields.GetByInternalNameOrTitle("Priority")
                    $removeMe2.DeleteObject()
                    $removeMe3 = $List.Fields.GetByInternalNameOrTitle("Related Items")
                    $removeMe3.DeleteObject()
                    $removeMe4 = $List.Fields.GetByInternalNameOrTitle("Due Date")
                    $removeMe4.DeleteObject()
                    $removeMe5 = $List.Fields.GetByInternalNameOrTitle("Description")
                    $removeMe5.DeleteObject()
                    $Context.ExecuteQuery()
                }

        } Catch {
            Write-Warning "Error occurred: $_.Exception.Message"
        }
    }
    End {}




}

# function that creates any missing columns in a specified list...
function Update-UserMgtListColumns(){
    Param (
        [Parameter(Mandatory=$True,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        [string]$whichList
    )


    Begin {}
    Process {
    
        Try {

            # load existing info - web, site, lists, fields for specific list...
            $web = $context.Web
            $site = $context.Site 
            $lists = $web.Lists
            $list = $web.Lists.GetByTitle("$whichList");
            $fields = $list.Fields;
            $context.Load($web)
            $context.Load($site)
            $context.Load($list)
            $context.Load($fields)
            $context.ExecuteQuery()


            # add TaskScript field, if not yet existing...
                $spFieldName = 'TaskScript'
                $spFieldCheck = 0
                ForEach ($field in $fields) {
                    $getTitle = $field.Title
                    If ($getTitle -eq "$spFieldName") { $spFieldCheck = 1 }
                }

                If ($spFieldCheck -eq 0) {    
                    $a = $List.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='TaskScript' MaxLength='255'></Field>",$true,[Microsoft.SharePoint.Client.AddFieldOptions]::AddFieldToDefaultView)
                    $List.Update()
                    $Context.ExecuteQuery()
                }

            # add TaskType field, if not yet existing...
                $spFieldName = 'TaskType'
                $spFieldCheck = 0
                ForEach ($field in $fields) {
                    $getTitle = $field.Title
                    If ($getTitle -eq "$spFieldName") { $spFieldCheck = 1 }
                }

                If ($spFieldCheck -eq 0) {    
                    $a = $List.Fields.AddFieldAsXml("<Field Type='Choice' DisplayName='TaskType'>
                        <CHOICES>
                            <CHOICE>Automatic</CHOICE>
                            <CHOICE>Manual</CHOICE>
                        </CHOICES></Field>",$true,[Microsoft.SharePoint.Client.AddFieldOptions]::AddFieldToDefaultView)
                    $List.Update()
                    $Context.ExecuteQuery()
                }

            # add TaskDay field, if not yet existing...
                $spFieldName = 'TaskDay'
                $spFieldCheck = 0
                ForEach ($field in $fields) {
                    $getTitle = $field.Title
                    If ($getTitle -eq "$spFieldName") { $spFieldCheck = 1 }
                }

                If ($spFieldCheck -eq 0) {    
                    $a = $List.Fields.AddFieldAsXml("<Field Type='Number' DisplayName='TaskDay' Decimals='0' Percentage='False' Min='0' Max='360'></Field>",$true,[Microsoft.SharePoint.Client.AddFieldOptions]::AddFieldToDefaultView)
                    $List.Update()
                    $Context.ExecuteQuery()
                }

            # add TaskOrder field, if not yet existing...
                $spFieldName = 'TaskOrder'
                $spFieldCheck = 0
                ForEach ($field in $fields) {
                    $getTitle = $field.Title
                    If ($getTitle -eq "$spFieldName") { $spFieldCheck = 1 }
                }

                If ($spFieldCheck -eq 0) {    
                    $a = $List.Fields.AddFieldAsXml("<Field Type='Number' DisplayName='TaskOrder' Decimals='0' Percentage='False' Min='0' Max='50'><Default>50</Default></Field>",$true,[Microsoft.SharePoint.Client.AddFieldOptions]::AddFieldToDefaultView)
                    $List.Update()
                    $Context.ExecuteQuery()
                }



            # add SamAccountName field, if not yet existing...
                $spFieldName = 'SamAccountName'
                $spFieldCheck = 0
                ForEach ($field in $fields) {
                    $getTitle = $field.Title
                    If ($getTitle -eq "$spFieldName") { $spFieldCheck = 1 }
                }

                If ($spFieldCheck -eq 0) {    
                    $a = $List.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='SamAccountName' MaxLength='20'></Field>",$true,[Microsoft.SharePoint.Client.AddFieldOptions]::AddFieldToDefaultView)
                    $List.Update()
                    $Context.ExecuteQuery()
                }

            # add EmailSubject field, if not yet existing...
                $spFieldName = 'EmailSubject'
                $spFieldCheck = 0
                ForEach ($field in $fields) {
                    $getTitle = $field.Title
                    If ($getTitle -eq "$spFieldName") { $spFieldCheck = 1 }
                }

                If ($spFieldCheck -eq 0) {    
                    $a = $List.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='EmailSubject' MaxLength='255'></Field>",$true,[Microsoft.SharePoint.Client.AddFieldOptions]::AddFieldToDefaultView)
                    $List.Update()
                    $Context.ExecuteQuery()
                }

            # add CompletedDate field, if not yet existing...
                $spFieldName = 'CompletedDate'
                $spFieldCheck = 0
                ForEach ($field in $fields) {
                    $getTitle = $field.Title
                    If ($getTitle -eq "$spFieldName") { $spFieldCheck = 1 }
                }

                If ($spFieldCheck -eq 0) {    
                    $a = $List.Fields.AddFieldAsXml("<Field Type='DateTime' DisplayName='CompletedDate' Format='DateOnly'></Field>",$true,[Microsoft.SharePoint.Client.AddFieldOptions]::AddFieldToDefaultView)
                    $List.Update()
                    $Context.ExecuteQuery()
                }

            # add TaskDetails field, if not yet existing...
                $spFieldName = 'TaskDetails'
                $spFieldCheck = 0
                ForEach ($field in $fields) {
                    $getTitle = $field.Title
                    If ($getTitle -eq "$spFieldName") { $spFieldCheck = 1 }
                }

                If ($spFieldCheck -eq 0) {    
                    $a = $List.Fields.AddFieldAsXml("<Field Type='Note' DisplayName='TaskDetails' NumLines='6' RichText='False' Sortable='False'></Field>",$true,[Microsoft.SharePoint.Client.AddFieldOptions]::AddFieldToDefaultView)
                    $List.Update()
                    $Context.ExecuteQuery()
                }


        } Catch {
            Write-Warning "Error occurred: $_.Exception.Message"
        }
    
    
    
    }
    End {}




}

# function that gets existing task templates...
function Get-UserMgtTemplates(){
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
                $newItem = New-Object PSObject -Property @{
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

            Return $newArray

        } Catch {
            Write-Warning "Error occurred: $_.Exception.Message"
        }
    }
    End {}




}

# function that gets existing task templates...
function Update-UserMgtTemplates(){
    Param (
        [Parameter(Mandatory=$True,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        [object]$items,
        [string]$samAccountName,
        [string]$spWeb,
        [string]$emailSubject
    )

    Begin {}
    Process {
    
        Try {

            #Add-PSSnapin Microsoft.SharePoint.PowerShell
            #$spWeb = Get-SPWeb $spWeb
            $newItems = $items

            $newItems | ForEach-Object {
                $_.SamAccountName = "$samAccountName"
                $_.EmailSubject = "$emailSubject"
                $thisDay = $_.TaskDay
                $thisTaskType = $_.TaskType
                If ($thisDay -gt 1) {
                    $_.StartDate = (Get-Date).AddDays($thisDay)
                } Else {
                    $_.StartDate = Get-Date
                }
                If ($thisTaskType -eq 'Automatic') {
                    $assignedTo = Import-CliXML "$global:usrMgmtServerSource\People\usrMgmtSpUser.xml"
                    $_.AssignedTo = $assignedTo
                } Else {
                    $assignedTo = Import-CliXML "$global:usrMgmtServerSource\People\usrMgmtAssignedManual.xml"
                    $_.AssignedTo = $assignedTo
                }
            }
            Return $newItems

        } Catch {
            Write-Warning "Error occurred: $_.Exception.Message"
        }
    }
    End {}




}

# function that creates new tasks in SharePoint for the user off-boarding...
function New-UserMgtOffboardingTasks(){
    Param (
        [Parameter(Mandatory=$True,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        [object]$items,
        [string]$whichList,
        [string]$spWeb
    )

    Begin {}
    Process {
    
        Try {

            Add-PSSnapin Microsoft.SharePoint.PowerShell

            # load existing info - web, site, lists, fields for specific list...
            $web = $context.Web
            $site = $context.Site 
            $lists = $web.Lists
            $list = $web.Lists.GetByTitle("$whichList");
            $fields = $list.Fields;
            $context.Load($web)
            $context.Load($site)
            $context.Load($list)
            $context.Load($fields)
            $context.ExecuteQuery()


            # items $updatedTemplates -spWeb "$global:usrMgmtSpWeb" $whichList = "$global:usrMgmtTasksListName"


            $items | ForEach-Object {

                $thisItem = $_


                $newItemInfo = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation
                $newItem = $list.AddItem($newItemInfo)
                $newItem["TaskType"] = $thisItem."TaskType"
                $newItem["TaskDay"] = $thisItem."TaskDay"
                $newItem["TaskOrder"] = $thisItem."TaskOrder"
                $newItem["EmailSubject"] = $thisItem."EmailSubject"
                $newItem["StartDate"] = $thisItem."StartDate"
                $newItem["AssignedTo"] = $thisItem."AssignedTo".LookupId
                $newItem["Title"] = $thisItem."Title"
                $newItem["TaskScript"] = $thisItem."TaskScript"
                $newItem["SamAccountName"] = $thisItem."SamAccountName"
                $newItem.Update()
                $Context.ExecuteQuery()
            }



        } Catch {
            Write-Warning "Error occurred: $_.Exception.Message"
        }
    }
    End {}




}

# function to get things started...
function New-UserMgtTasksProjectStart() {

    $getFiles = Get-ChildItem "$global:usrMgmtServerSource\Requests\*" -Include "*.xml"

    [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client")
    [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime")
    #Bind to site collection
        $Context = New-Object Microsoft.SharePoint.Client.ClientContext($global:usrMgmtSpWeb)
        $Creds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($global:usrMgmtSpUser,$global:usrMgmtCredSpOnline)
        $Context.Credentials = $Creds

    # check that lists exist already and are updated with all needed columns...
        New-UserMgtListCheck
        Update-UserMgtListColumns -whichList "$global:usrMgmtTemplatesListName"
        Update-UserMgtListColumns -whichList "$global:usrMgmtTasksListName"

    # load existing info - web, site, lists, fields for specific list...
        $web = $context.Web
        $site = $context.Site 
        $lists = $web.Lists
        $listTemplates = $web.Lists.GetByTitle("$global:usrMgmtTemplatesListName");
        $listTasks = $web.Lists.GetByTitle("$global:usrMgmtTasksListName");
        $fieldsTemplates = $listTemplates.Fields;
        $fieldsTasks = $listTasks.Fields;
        $qry = [Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery()
        $items = $listTemplates.GetItems($qry)
        $context.Load($web)
        $context.Load($site)
        $context.Load($listTemplates)
        $context.Load($listTasks)
        $context.Load($lists)
        $context.Load($fieldsTemplates)
        $context.Load($fieldsTasks)
        $context.Load($items)
        $context.ExecuteQuery()


    # load and tweak the existing template tasks...
        $getTemplates = Get-UserMgtTemplates -whichList "$global:usrMgmtTemplatesListName"    

    # create new tasks in SharePoint for the user being off-boarded...

        $getFiles | ForEach-Object {
            Try {
                $thisFileName = $_.Name
                $thisFileNow = "$global:usrMgmtServerSource\Requests\$thisFileName"
                $thisFileMoved = "$global:usrMgmtServerSource\Requests\Processed\$thisFileName"

                $thisFileContents = Import-CliXml -Path $thisFileNow
                $thisSam = $thisFileContents.SamAccountName
                $thisEmailSubject = $thisFileContents.EmailSubject

                $updatedTemplates = Update-UserMgtTemplates -items $getTemplates -spWeb "$global:usrMgmtSpWeb" -samAccountName "$thisSam" -emailSubject "$thisEmailSubject"
                New-UserMgtOffboardingTasks -items $updatedTemplates -spWeb "$global:usrMgmtSpWeb" -whichList "$global:usrMgmtTasksListName"

                Move-Item "$thisFileNow" "$thisFileMoved"


                $thisEmailBody = "<body><div style='font-family: calibri;'>Hello there, <br/><br/>The off-boarding scripts have been run for the following user: $thisSam. Keep an eye out on <a href='https://yoursharepointURL/sites/Offboarding/Lists/Offboarding%20Tasks/AllItems.aspx'>the SharePoint Online site</a> for any tasks that need to be performed...<br/><br/> Thanks, <br/>Virtual assistant</div></body>"
                send-mailmessage `
                -to "you@yourdomain.com" `
                -from "noreply@yourdomain.com" `
                -subject "Launch Pad New Off-boarding" `
                -body $thisEmailBody `
                -smtpserver "yourSMTPserver" `
                -BodyAsHtml


            } Catch {
                Write-Warning "Error occurred: $_.Exception.Message"
            }
        
        }   
}
