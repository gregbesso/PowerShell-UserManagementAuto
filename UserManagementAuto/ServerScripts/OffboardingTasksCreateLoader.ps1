#OffboardingTasksCreateLoader

# had to put these here so they can be used, duplicates of variables in the import but are needed for now...
$global:usrMgmtServerSource = '\\server\share\UserManagementAuto'
$global:usrMgmtLocalPath = 'C:\psmanage'

#
# Check if network copy of scripts is accessible. If it is, copy down latest version before importing...
#
Try {
    If (Test-Path "$global:usrMgmtServerSource\ServerScripts\OffboardingTasksCreate.ps1") { 
        If (!(Test-Path "$global:usrMgmtLocalPath")) { New-Item -ItemType Directory -Path "$global:usrMgmtLocalPath" }
        If (!(Test-Path "$global:usrMgmtLocalPath\Server")) { New-Item -ItemType Directory -Path "$global:usrMgmtLocalPath\Server" }
        Copy-Item -Path "$global:usrMgmtServerSource\ServerScripts\OffboardingImports.ps1" -Destination "$global:usrMgmtLocalPath\Server\OffboardingImports.ps1" -Force
        Copy-Item -Path "$global:usrMgmtServerSource\ServerScripts\OffboardingTasksCreateLoader.ps1" -Destination "$global:usrMgmtLocalPath\Server\OffboardingTasksCreateLoader.ps1" -Force
        Copy-Item -Path "$global:usrMgmtServerSource\ServerScripts\OffboardingTasksCreate.ps1" -Destination "$global:usrMgmtLocalPath\Server\OffboardingTasksCreate.ps1" -Force
    }
} Catch {}


# then load the locally copied script and get started...
If (Test-Path "$global:usrMgmtLocalPath\Server\OffboardingTasksCreate.ps1") { 
    #  
    # Import local copy of scripts to be used below, and then call the control script to get things moving along...
    #
    . "$global:usrMgmtLocalPath\Server\OffboardingImports.ps1"
    . "$global:usrMgmtLocalPath\Server\OffboardingTasksCreate.ps1"


    New-UserMgtTasksProjectStart
}