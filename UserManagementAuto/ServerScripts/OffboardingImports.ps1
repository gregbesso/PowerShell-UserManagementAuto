#
# This file has some repeatedly used variables and other items for the Offboarding Automation project...
#

# this folder will contain the folders used by the project. Cred, People, Requests, ServerScripts, TaskScripts...
$global:usrMgmtServerSource = '\\tremor.local\data\IT Admin\PowerShell\UserManagementAuto'

# One-time step for installing
#
# $secureStringSpOnline = Read-Host -AsSecureString "Enter your service account password"
# $secureStringSpOnline | ConvertFrom-SecureString | Out-File -FilePath "$global:usrMgmtServerSource\Cred\usrMgmtCredential.txt"


# this is to pull a valid domain controller name for use in certain places...
$global:usrMgmtDomainController = ([ADSI]"LDAP://RootDSE").dnshostname
# this is for the domain netbios name, used in one or two places...
$global:usrMgmtDomainNetbios = "$env:userdomain"
# this is for when the fqdn of the domain is needed...
$global:usrMgmtDomainFqdn = "$env:userdnsdomain"
# this is the local process server, where your scheduled tasks will be created/run from...
$global:usrMgmtProcessServer = 'NYC-SPF01'
# this will be the local folder where the process server stores things
# i use folder from another existing script but you can use anything...
$global:usrMgmtLocalPath = 'C:\psmanage'
# same as above, but what it is when connecting remotely - for creating scheduled tasks...
$global:usrMgmtRemotePath = 'c$\psmanage'
# this is the folder that the IT generated projects are initially stored in XML format...
$global:usrMgmtRequestsShare = "$global:usrMgmtServerSource\Requests"
# this is the SharePoint Online site that will store all the tasks...
$global:usrMgmtSpWeb = 'https://tremorvideo.sharepoint.com/sites/ITTEAM/Offboarding' 
# this is the service account that is used to connect to SharePoint Online...
$global:usrMgmtSpUser = 'virtualvicky@tremorvideo.com'
# this is the user or group that you want to use for assigning manual tasks to...
$global:usrMgmtAssignedManual = 'IT@tremorvideo.com'
# this is the local service account that runs the scheduled tasks. i use same as service account above...
$global:usrMgmtProcessUser = 'virtualvicky'
# this line imports the pre-stored password for automatic authentication, if on the process server...
If ($env:computername -eq "$global:usrMgmtProcessServer") {
    $global:usrMgmtCredSpOnline = Get-Content -Path "$global:usrMgmtServerSource\Cred\usrMgmtCredential.txt" | ConvertTo-SecureString
}
# this is the name of the task list that stores pre-defined tasks for each offboarding project...
$global:usrMgmtTemplatesListName = 'Offboarding Templates'
# this is the name of the list that new tasks are created for each offboarding project, using copies of the templates...
$global:usrMgmtTasksListName = 'Offboarding Tasks'
# this is the local process server, where your scheduled tasks will be created/run from...
$global:usrMgmtProcessServer = 'NYC-SPF01'
# this is the name of the scheduled task that creates tasks on SPO...
$global:usrMgmtScheduledCreate = 'OffboardingTasksCreate'
# this is the name of the scheduled task that runs the tasks...
$global:usrMgmtScheduledPerform = 'OffboardingTasksPerform'




# function that setups up the scheduled tasks on the processing server
# only need to run this manually when first "installing" these
function New-UserMgmtScheduledTasks() {
    #
    # get some random variables for the start and repeat times for the scheduled task 
    # (so not every computer updates at the same time and kills SharePoint hehe)
    #
    
    BEGIN{}
    PROCESS{
        If (Test-Connection -ComputerName $global:usrMgmtProcessServer -Quiet) {
            # this one i set to something like 9pm for tasks process...
            $startHour = "21"
            $startMinute = "00"



            #
            # Add the service account to the local administrators group on the system
            #
            Try {
                $addToAdmins = [ADSI]"WinNT://$global:usrMgmtProcessServer/Administrators,group" 
                $addToAdmins.psbase.Invoke("Add",([ADSI]"WinNT://$global:usrMgmtDomainNetbios/$global:usrMgmtProcessUser").path)
            } Catch {}


            #
            # Copy the scheduled task files to the process server
            #
            Try {
                If (Test-Path "$global:usrMgmtServerSource\ServerScripts\OffboardingTasksCreate.ps1") { 
                    If (!(Test-Path "\\$global:usrMgmtProcessServer\$global:usrMgmtRemotePath")) { New-Item -ItemType Directory -Path "\\$global:usrMgmtProcessServer\$global:usrMgmtRemotePath" }
                    If (!(Test-Path "\\$global:usrMgmtProcessServer\$global:usrMgmtRemotePath\Server")) { New-Item -ItemType Directory -Path "\\$global:usrMgmtProcessServer\$global:usrMgmtRemotePath\Server" }
                    Copy-Item -Path "$global:usrMgmtServerSource\ServerScripts\OffboardingImports.ps1" -Destination "\\$global:usrMgmtProcessServer\$global:usrMgmtRemotePath\Server\OffboardingImports.ps1" -Force
                    Copy-Item -Path "$global:usrMgmtServerSource\ServerScripts\OffboardingTasksCreateLoader.ps1" -Destination "\\$global:usrMgmtProcessServer\$global:usrMgmtRemotePath\Server\OffboardingTasksCreateLoader.ps1" -Force
                    Copy-Item -Path "$global:usrMgmtServerSource\ServerScripts\OffboardingTasksCreate.ps1" -Destination "\\$global:usrMgmtProcessServer\$global:usrMgmtRemotePath\Server\OffboardingTasksCreate.ps1" -Force
                    Copy-Item -Path "$global:usrMgmtServerSource\ServerScripts\OffboardingTasksPerformLoader.ps1" -Destination "\\$global:usrMgmtProcessServer\$global:usrMgmtRemotePath\Server\OffboardingTasksPerformLoader.ps1" -Force
                    Copy-Item -Path "$global:usrMgmtServerSource\ServerScripts\OffboardingTasksPerform.ps1" -Destination "\\$global:usrMgmtProcessServer\$global:usrMgmtRemotePath\Server\OffboardingTasksPerform.ps1" -Force
                }
            } Catch {}


            #
            # Create the scheduled task on the computer
            #
            
            #First, check then create (if missing) for the create tasks job...
            Try {
                $existingQueries = Schtasks.exe /S $global:usrMgmtProcessServer /Query /TN "$global:usrMgmtScheduledCreate"
            } Catch {}

            Try {
                If ($existingQueries.Length -lt 1) {
                    $tempPW = [System.Runtime.InteropServices.marshal]::PtrToStringAuto([System.Runtime.InteropServices.marshal]::SecureStringToBSTR($global:usrMgmtCredSpOnline))
                    # if running without schedule, for create...
                    Schtasks.exe /S $global:usrMgmtProcessServer /Create /RU "$global:usrMgmtProcessUser@$global:usrMgmtDomainFqdn" /RP "$tempPW" /SC ONCE /ST $startHour":"$startMinute /TN "$global:usrMgmtScheduledCreate" /TR "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy Bypass -File $global:usrMgmtLocalPath\Server\OffboardingTasksCreateLoader.ps1" /RL HIGHEST
                    $tempPW = "nothing"
                }
            } Catch {}


            #Second, check then create (if missing) for the process tasks job...
            Try {
                $existingQueries = Schtasks.exe /S $global:usrMgmtProcessServer /Query /TN "$global:usrMgmtScheduledPerform"
            } Catch {}

            Try {
                If ($existingQueries.Length -lt 1) {
                    $tempPW = [System.Runtime.InteropServices.marshal]::PtrToStringAuto([System.Runtime.InteropServices.marshal]::SecureStringToBSTR($global:usrMgmtCredSpOnline))
                    # if running daily, for process...
                    Schtasks.exe /S $global:usrMgmtProcessServer /Create /RU "$global:usrMgmtProcessUser@$global:usrMgmtDomainFqdn" /RP "$tempPW" /SC DAILY /ST $startHour":"$startMinute /MO 1 /TN "$global:usrMgmtScheduledPerform" /TR "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy Bypass -File $global:usrMgmtLocalPath\Server\OffboardingTasksPerformLoader.ps1" /RL HIGHEST
                    $tempPW = "nothing"
                }
            } Catch {}
        }
    }
    End {}
}

# function to get your "assigned to" users from SharePoint Online and store their user ID for use later
# only need to run this manually when first "installing" these, or when changing the users you may reference...
function Get-UserMgmtAssignToUsers() {


    [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client")
    [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime")
    #Bind to site collection
        $Context = New-Object Microsoft.SharePoint.Client.ClientContext($global:usrMgmtSpWeb)
        $Creds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($global:usrMgmtSpUser,$global:usrMgmtCredSpOnline)
        $Context.Credentials = $Creds


        $Users = $Context.Web.SiteUsers
        $context.Load($Users)
        $context.ExecuteQuery()

        # get the user ID for your service account's account in SPO...
        $getUserID = ''
        $users | ForEach-Object {
            $thisEmail = $_.Email
            If ($thisEmail -eq $global:usrMgmtSpUser) {                
                $template = Import-CliXML "$global:usrMgmtServerSource\People\template.xml"
                $template.Email = $_.Email
                $template.TypeId = $_.UserId.TypeId
                $template.LookupId = $_.Id
                $template.LookupValue = $_.Title
            }
        }
        If ($template.LookupId.Length -gt 0) {
            $template | Export-CliXML "$global:usrMgmtServerSource\People\usrMgmtSpUser.xml" -Force
        }

        # get the user/group ID for your manual tasks to be assigned to...
        $getUserID = ''
        $users | ForEach-Object {
            $thisEmail = $_.Email
            If ($thisEmail -eq $global:usrMgmtAssignedManual) {
                $template = Import-CliXML "$global:usrMgmtServerSource\People\template.xml"
                $template.Email = $_.Email
                $template.TypeId = $_.UserId.TypeId
                $template.LookupId = $_.Id
                $template.LookupValue = $_.Title
            }
        }
        If ($template.LookupId.Length -gt 0) {
            $template | Export-CliXML "$global:usrMgmtServerSource\People\usrMgmtAssignedManual.xml" -Force
        }
}
