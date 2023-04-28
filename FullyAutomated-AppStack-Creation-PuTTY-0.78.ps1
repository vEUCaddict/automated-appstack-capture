#################################**####################################
############################***********################################
########################*******************############################
#####################***********((((************#######################
################************((((((((((((/***********###################
############************(((((((((((((((((((((***********###############
########***********/(((((((((***((((***((((((((((************##########
####***********((((((((((*******((((********(((((((((************######
####********(((((((((/***********((((************(((((((((/********####
####******(((((((****************((((****************((((((((******####
####******(((((******************((((******************((((((******####
####******(((((******************((((******************((((((******####
####******(((((******************((((******************((((((******####
####******(((((******************((((******************((((((******####
####******(((((******************((((******************((((((******####
####******(((((****************(((((((((***************((((((******####
####******(((((***********((((((((((((((((((***********((((((******####
####******(((((*******(((((((((********/((((((((/******((((((******####
####******(((((**/(((((((((*****************(((((((((**((((((******####
####******(((((((((((((*************************(((((((((((((******####
####******((((((((/*********************************(((((((((******####
####*******/((((((((*******************************((((((((********####
#####************((((((((**********************/((((((((***********####
########************((((((((**************((((((((/***********#########
#############***********((((((((/*****((((((((************#############
################***********/(((((((((((((************##################
####################************(((((************######################
########################********************###########################
#############################***********###############################
################################***####################################
#######################################################################
#######################################################################
#######################################################################
#### AUTHOR: Sidney Laan (LANOS-GreenIT)     ##########################
#### VERSION: 001                            ##########################
#### RELEASE DATE: 28 apr. 2023              ##########################
#### FILE NAME: FullyAutomated-AppStack-     ##########################
#### Creation-PuTTY-0.78.ps1 	         	 ##########################
#######################################################################

#######################################################################
############################## README #################################
#######################################################################

# Edit the variable section to your environment needs
# Open a PowerShell window (preferred not as administrator)
# Execute: Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
# Execute: cd %scriptlocation%
# Execute: .\FullyAutomated-AppStack-Creation-PuTTY-0.78.ps1 -AV_CAP_VM_NAME %VMNAME%


######################################################################
######################### RELEASE NOTES ##############################
######################################################################

# V001 = The main goal of this powershell script is to automate the creation of App Volumes 4.x applications and application package. Keep in mind you have some prerequisites to start this script!


#####################################################################
#################### INPUT PARAMTERS SECTION ########################
#####################################################################

param (
	$AV_CAP_VM_NAME
)


#####################################################################
####################### VARIABLES SECTION ###########################
#####################################################################

#####################################################################
####################### STATIC VARIABLES ############################
################# NOTE: NORMALLY SET ONE TIME #######################
#####################################################################

# Username of the AD packaging service account (don't put domain in front of the username!) (User is App Volumes Administrator, Custom vCenter Role and is a local administrator on VMware App Volumes Capture VM)
$AV_CAP_USR = "packageuser"

# Fully Qualified Domain Name (FQDN) of your Windows AD environment
$AV_CAP_DOM = "dummy.local"

# Request password for the App Volumes Package service account
$AV_CAP_USR_CRED = Get-Credential -Message "Enter the password of the App Volumes Package service account" -UserName "$AV_CAP_DOM\$AV_CAP_USR"
$AV_CAP_PWD = $AV_CAP_USR_CRED.getnetworkcredential().password

# Name of the vCenter Server Appliance (FQDN)
$VC_SRV = "vc01.dummy.local"

# Name of the App Volumes Manager Server (FQDN)
$AV_MGR_SRV = "appvolumes.dummy.local"

# Communication protocol to App Volumes Managers REST API (OPTIONS are http or https)
$AV_MGR_REST_API_PROT = "https"

# Virtual Machine name of the VMware App Volumes Capture VM in VMware vCenter/ESXi (NOTE: THIS IS NOT ALWAYS THE HOSTNAME OR FQDN!)
#$AV_CAP_VM_NAME = ""

# Windows variable COMPUTERNAME of the VMware App Volumes Capture VM
$AV_CAP_VM_COMPNAME = "$AV_CAP_VM_NAME.$AV_CAP_DOM"

# Name of the snapshot from the clean VMware App Volumes Capture VM state
$AV_CAP_VM_SS = "====== START HERE ======"

# Stage of the App Volumes - Application Package IN 4.0 & 4.0.1 (1 = TESTED, 3 = PUBLISHED, 4 = RETIRED, 5 = NEW) = CHANGED IN APP VOLUMES 4.1 (1 = NEW, 2 = TESTED, 3 = PUBLISHED, 4 = RETIRED)
$AV_APP_PKG_STAGE = "1"

# Name of the VMware vCenter Data Center for the App Volumes Application Package creation
$AV_APP_PKG_DC = "Datacenter"

# Name of the VMware datastore where in the App Volumes Application Packages are being created
$AV_APP_PKG_DS = "datastore"

# Path where the App Volumes Applications Packages resides
$AV_APP_PKG_PATH = "appvolumes/packages"

# Processing behaviour of the App Volumes Application Package (true = run in the background, false = wait for completion)
$AV_APP_PKG_BKG = "true"

# Path where the App Stack Application Package Templates resides
$AV_APP_TMP_PATH = "appvolumes/packages_templates"

# Name of the Application Package Template to be used (including *.vmdk extension)
$AV_APP_TMP_NAME = "template.vmdk"

# Path of the software repository (network share, where at least everyone/authenticated users has read permissions)
$AV_SOURCE_REPO = "\\dummy.local\DFS\Sources\AppStack"

# Path were the PowerShell script starts from to create logging
$SCRIPT_PATH = Get-Location

# Reboot time in seconds (monitor with a stopwatch, it really depends on your environment)
$AV_CAP_VM_REBOOT_TIME = 45

#####################################################################
######################## DYNAMIC VARIABLES ##########################
##################### NOTE: CHANGE FREQUENTLY #######################
#####################################################################


# Name of the VMware App Volumes 4.x Application (Inventory > Applications)
$AV_APP_NAME = "PuTTY"

# Description of the App Volumes - Application (optional)
$AV_APP_NAME_DESC = ""

# Name of the VMware App Volumes 4.x Application Application Package (Inventory > Packages)
$AV_APP_PKG_NAME = "PuTTY 0.78"

# Description of the App Volumes - Application Package (optional)
$AV_APP_PKG_DESC = ""

# Path which comes after the software repository ($AV_SOURCE_REPO) where the software installers resides
$AV_SOURCE_PATH = "Simon Tatham\PuTTY\0.78"

# Name of the installer file (including *.msi/*.exe extension)
$AV_SOURCE_INST = "putty-64bit-0.78-installer.msi"

# Name of the MSI transform file (including *.mst extension), if exists else keep empty
$AV_SOURCE_MST = ""

# (Silent) Installation Parameters to pre-configure the software
# EXE example: "/S"
# MSI example: "/i `"C:\Windows\Logs\$AV_SOURCE_INST`" TRANSFORMS=`"C:\Windows\Logs\$AV_SOURCE_MST`" REBOOT=`"ReallySuppress`" /qb- /l*v `"C:\Windows\Logs\install-$AV_APP_PKG_NAME.log`""
$AV_SOURCE_PARM = "/i `"C:\Windows\Logs\$AV_SOURCE_INST`" REBOOT=`"ReallySuppress`" /qb- /l*v `"C:\Windows\Logs\install-$AV_APP_PKG_NAME.log`""

# Post Configuration tasks which are executed after the installation (copy the examples commands after the : into the $AV_SOURCE_PSTCFG_OPTIONS which builds up an array of tasks)
# List of examples:

# Example to run a patch or upgrade installer just after the main installer (*.msi):					'Start-Process -Wait "msiexec.exe" -ArgumentList "/i "C:\Windows\Logs\' + $AV_SOURCE_INST + '" TRANSFORMS="C:\Windows\Logs\' + $AV_SOURCE_MST + '" REBOOT="ReallySuppress" /qb- /l*v "C:\Windows\Logs\install-' + $AV_APP_PKG_NAME + '.log""
# Example to run a patch or upgrade installer just after the main installer (*.exe):					'Start-Process -Wait -FilePath "C:\Windows\Logs\Updater\Setup.exe" -ArgumentList "/q"'
# Example to delete a Start Menu Folder: 																'Remove-Item -Path "' + ${env:ProgramData} + '\Microsoft\Windows\Start Menu\Programs\7-Zip" -Recurse -Force'
# Example to delete a Shart Menu Shortcut (include *.lnk extension): 									'Remove-Item -Path "' + ${env:ProgramData} + '\Microsoft\Windows\Start Menu\Programs\Google Chrome.lnk" -Force -Confirm:$false'
# Example to delete a Desktop Shortcut (include *.lnk extension): 										'Remove-Item -Path "C:\Users\Public\Desktop\Google Chrome.lnk" -Force -Confirm:$false'
# Example to delete a Folder inside the installation directory: 										'Remove-Item -Path "' + ${env:ProgramFiles} + '\NotePad++\updater" -Recurse -Force'
# Example to delete all files inside a subfolder from the installation directory: 						'Remove-Item -Path "' + ${env:ProgramFiles(x86)} + '\WinSCP\Translations\*.*" -Force -Confirm:$false'
# Example to delete a registry key in the HKEY_LOCAL_MACHINE hive: 										'Remove-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run" -Name "Greenshot" -Force -Confirm:$false'
# Example to copy a folder from the temporary installer directory to the installation directory: 		'Copy-Item -Path "C:\Windows\Logs\ComparePlugin" -Destination "' + ${env:ProgramFiles} + '\Notepad++\plugins" -Recurse -Force'
# Example to copy a file from the temporary installer directory to the installation directory : 		'Copy-Item -Path "C:\Windows\Logs\WinSCP.nl" -Destination "' + ${env:ProgramFiles(x86)} + '\WinSCP\Translations" -Force -Confirm:$false'
# Example to add a registry key in the HKEY_LOCAL_MACHINE hive: 										'New-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\SURF\SURFdrive" -Value 1 -PropertyType dword -Name "skipUpdateCheck" -Force'

# Array of post configuration tasks to execute after the installation process succeeds
$AV_SOURCE_PSTCFG_OPTIONS = @(
	'Remove-Item -Path "' + ${env:ProgramData} + '\Microsoft\Windows\Start Menu\Programs\PuTTY (64-bit)\PuTTY Web Site.lnk" -Force -Confirm:$false'
	'Remove-Item -Path "' + ${env:ProgramData} + '\Microsoft\Windows\Start Menu\Programs\PuTTY (64-bit)\PuTTY Manual.lnk" -Force -Confirm:$false'
)


#####################################################################
###################### START LOGGING OUTPUT #########################
#####################################################################

$DateTime = $(Get-Date -Format "yyyy-MM-dd_HH.mm")
IF( -Not (Test-Path -Path "$SCRIPT_PATH\_Logs" ) ) { New-Item -ItemType directory -Path "\\$SCRIPT_PATH\_Logs" | Out-Null }
Start-Transcript -Path "$SCRIPT_PATH\_Logs\$($DateTime)_Auto-AppStack-Creation-of-$($AV_APP_PKG_NAME).txt" -IncludeInvocationHeader -Force -Confirm:$false


#####################################################################
################## POWERSHELL MODULE(S) SECTION #####################
#####################################################################

# IMPORT VMWARE.VIMAUTOMATION.CORE - POWERCLI MODULE
Write-Host "Trying to import VMware PowerCLI module(s)" -ForegroundColor "Cyan"
TRY {
    Import-Module VMware.VimAutomation.Core
	Write-Host "Import of the VMware.VimAutomation.Core PowerShell module is succesful!" -ForegroundColor "Green"
	Write-Host "`n"
}
CATCH {
    Write-Host "VMware PowerCLI Module already loaded..." -ForegroundColor "Yellow"
	Write-Host "`n"
}


#####################################################################
###################### FUNCTION(S) SECTION ##########################
#####################################################################

FUNCTION Start-Sleep-ProgressBar($seconds) {
    $doneDT = (Get-Date).AddSeconds($seconds)
    while($doneDT -gt (Get-Date)) {
        $secondsLeft = $doneDT.Subtract((Get-Date)).TotalSeconds
        $percent = ($seconds - $secondsLeft) / $seconds * 100
        Write-Progress -Activity "Waiting till the script can continue." -Status "Running..." -SecondsRemaining $secondsLeft -PercentComplete $percent
        [System.Threading.Thread]::Sleep(500)
    }
    Write-Progress -Activity "Waiting till the script can continue." -Status "Running..." -SecondsRemaining 0 -Completed
}

#####################################################################
###################### START SCRIPT SECTION #########################
###################  DON'T EDIT SCRIPT BELOW,  ######################
##############  UNLESS YOU KNOW WHAT YOU ARE DOING! #################
#####################################################################

# SETUP REST API CONNECTION TO APP VOLUMES MANAGER SERVER
Write-Host "Connect to $AV_MGR_SRV via REST API..." -ForegroundColor "Cyan"
TRY {
    $body = @{
        username = "$AV_CAP_DOM\$AV_CAP_USR"
        password = $AV_CAP_PWD
    }
    Invoke-RestMethod -Method Post -Uri "$AV_MGR_REST_API_PROT`://$av_mgr_srv/cv_api/sessions" -Body $body -SessionVariable Login | Out-Null
    Write-Host "Connection to App Volumes REST API is successful!" -ForegroundColor "Green"
	Write-Host "`n"
}
CATCH {
    $_.Exception.Message
    Write-Host "Connection to App Volumes REST API failed!" -ForegroundColor "Red"
	Write-Host "`n"
    Stop-Transcript
    pause
	exit
}


# GET APP VOLUMES PACKAGE USER FOR OWNERSHIP OF APPLICATION
$AV_USR = ((Invoke-WebRequest -WebSession $Login -Method Get -Uri "$AV_MGR_REST_API_PROT`://$av_mgr_srv/app_volumes/owners?name=$AV_CAP_USR&filter=contains&recursive=0&type=user").content | ConvertFrom-Json).data
$AV_USR_GUID = $AV_USR.object_guid


# CREATE APPLICATION AND WAIT FOR COMPLETION
Write-Host "STEP 1: Create an App Volumes 4.x - Application, if not already exists..." -ForegroundColor "Cyan"
$AV_APP = ((Invoke-WebRequest -WebSession $Login -Method Get -Uri "$AV_MGR_REST_API_PROT`://$av_mgr_srv/app_volumes/app_products").content | ConvertFrom-Json).data | Where-Object {$_.name -eq $AV_APP_NAME}
IF ($null -eq $AV_APP.Name) {
	Write-Host "Application does not exists yet, Application: $AV_APP_NAME is being created..." -ForegroundColor "Yellow"
	Invoke-WebRequest -WebSession $Login -Method Post -Uri "$AV_MGR_REST_API_PROT`://$av_mgr_srv/app_volumes/app_products?data[name]=$AV_APP_NAME&data[description]=$AV_APP_NAME_DESC&data[owner_guid=$AV_USR_GUID]" | Out-Null

    # WAIT TILL APPLICATION CREATION IS COMPLETED
    DO {
        Write-Host "Waiting for the application creation job is finished..."
        $pending_jobs = ((Invoke-WebRequest -WebSession $Login -Method Get -Uri "$AV_MGR_REST_API_PROT`://$av_mgr_srv/cv_api/jobs/pending").content | ConvertFrom-Json)
        Start-Sleep -s 2
    }
    UNTIL ($pending_jobs.pending -eq "0")
    Write-Host "Application creation job is finished, continue to next step..." -ForegroundColor "Green"
	Write-Host "`n"
}
ELSE {
	Write-Host "Application already exists, skipping this step..." -ForegroundColor "Green"
	Write-Host "`n"
}


# CREATE APPLICATION PACKAGE AND WAIT FOR COMPLETION
Write-Host "STEP 2: Create an App Volumes 4.x - Application Package..." -ForegroundColor "Cyan"
$AV_APP_PKG = ((Invoke-WebRequest -WebSession $Login -Method Get -Uri "$AV_MGR_REST_API_PROT`://$av_mgr_srv/app_volumes/app_packages").content | ConvertFrom-Json).data | Where-Object {$_.name -eq $AV_APP_PKG_NAME}
IF ($AV_APP_PKG.Name -eq $AV_APP_PKG_NAME) {
    Write-Host "Application Package with this name already exists, abort the script and will not continue..." -ForegroundColor "Red"
    Write-Host "Press Ctrl + C to abort the script..." -ForegroundColor "Cyan"
	Write-Host "`n"
    Stop-Transcript
    pause
	exit
}
ELSE {
    Write-Host "Application Package does not exists yet, Application Package: $AV_APP_PKG_NAME is being created..." -ForegroundColor "Yellow"
    $AV_APP = ((Invoke-WebRequest -WebSession $Login -Method Get -Uri "$AV_MGR_REST_API_PROT`://$av_mgr_srv/app_volumes/app_products").content | ConvertFrom-Json).data | Where-Object {$_.name -eq $AV_APP_NAME}
    $AV_APP_ID = $AV_APP.id
    Invoke-WebRequest -WebSession $Login -Method Post -Uri "$AV_MGR_REST_API_PROT`://$av_mgr_srv/app_volumes/app_packages?data[name]=$AV_APP_PKG_NAME&data[app_product_id]=$AV_APP_ID&data[lifecycle_stage_id]=$AV_APP_PKG_STAGE&data[description]=$AV_APP_PKG_DESC&data[datacenter]=$AV_APP_PKG_DC&data[datastore]=$AV_APP_PKG_DS&data[path]=$AV_APP_PKG_PATH&data[background]=$AV_APP_PKG_BKG&data[template_path]=$AV_APP_TMP_PATH&data[template_name]=$AV_APP_TMP_NAME" | Out-Null
    DO {
        Write-Host "Waiting for the application package creation job is finished..."
        $pending_jobs = ((Invoke-WebRequest -WebSession $Login -Method Get -Uri "$AV_MGR_REST_API_PROT`://$av_mgr_srv/cv_api/jobs/pending").content | ConvertFrom-Json)
        Start-Sleep -s 2
    }
    UNTIL($pending_jobs.pending -eq "0")
    Write-Host "Application Package creation job is finished, continue to next step..." -ForegroundColor "Green"
	Write-Host "`n"
}


# CONNECT TO VCENTER SERVER VIA POWERCLI
Write-Host "Connect to vCenter Server via PowerCLI..." -ForegroundColor "Cyan"
Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Scope Session -Confirm:$false | Out-Null
TRY {
	Connect-VIServer -Server "$VC_SRV" -User "$AV_CAP_DOM\$AV_CAP_USR" -Password "$AV_CAP_PWD" | Out-Null
	Write-Host "Connection to vCenter Server is succesful!" -ForegroundColor "Green"
	Write-Host "`n"
}
CATCH {
	Write-Host "Connection to vCenter failed, check the vCenter name, credentials or permissions..." -ForegroundColor "Red"
	Write-Host "Press Ctrl + C to abort the script..." -ForegroundColor "Cyan"
	Write-Host "`n"
    Stop-Transcript
    pause
	exit
}


# REVERT SNAPSHOT OF APP VOLUMES CAPTURE VM TO CLEAN STATE
Write-Host "Reverting: $AV_CAP_VM_NAME to a clean state..." -ForegroundColor "Cyan"
TRY {
	Set-VM -VM "$AV_CAP_VM_NAME" -SnapShot "$AV_CAP_VM_SS" -Confirm:$false | Out-Null
	Write-Host "Reverting to a clean state is completed!" -ForegroundColor "Green"
	Write-Host "`n"
}
CATCH {
	Write-Host "Reverting to a clean state failed..." -ForegroundColor "Red"
	Write-Host "Press Ctrl + C to abort the script..." -ForegroundColor "Cyan"
	Write-Host "`n"
    Stop-Transcript
    pause
	exit
}


# POWER ON THE APP VOLUMES CAPTURE VM
Write-Host "Powering on App Volumes Capture VM: $AV_CAP_VM_NAME ..." -ForegroundColor "Cyan"
TRY {
	Start-VM -VM "$AV_CAP_VM_NAME" | Out-Null
	Write-Host "App Volumes Capture VM: $AV_CAP_VM_NAME succesfully powered on!" -ForegroundColor "Green"
	Write-Host "`n"
}
CATCH {
	Write-Host "App Volumes Capture VM: $AV_CAP_VM_NAME failed to power on..." -ForegroundColor "Red"
	Write-Host "Press Ctrl + C to abort the script..." -ForegroundColor "Cyan"
	Write-Host "`n"
    Stop-Transcript
    pause
	exit	
}


# OPEN VMWARE CONSOLE WINDOW TO FOLLOW THE CAPTURING PROCESS (NOT NECESSARY, USEFUL FOR TROUBLESHOOTING)
Write-Host "Opening VMware (WEB) Console window to follow the capturing process..." -ForegroundColor "Cyan"
Write-Host "WARN: Don`'t click with the mouse in the VM, else you can interupt the capturing process!!!" -ForegroundColor "Yellow"
Write-Host "`n"
Open-VMConsoleWindow -VM "$AV_CAP_VM_NAME" -Confirm:$false
Start-Sleep-ProgressBar -s 30 # Higher up the number in seconds if your capture VM boot takes longer!


# COPY SOFTWARE (INSTALLER) TO APP VOLUMES CAPTURE VM TEMPORARY LOCATION
Write-Host "Copy $AV_APP_NAME installer to C:\Windows\Logs directory..." -ForegroundColor "Cyan"
Write-Host "`n"
IF( -Not (Test-Path -Path "\\$AV_CAP_VM_COMPNAME\C$\Windows\Logs" ) ) { New-Item -ItemType directory -Path "\\$AV_CAP_VM_COMPNAME\C$\Windows\Logs" | Out-Null }
$SourceCopied = Start-Job -ScriptBlock { Copy-Item -Path "$using:AV_SOURCE_REPO\$using:AV_SOURCE_PATH\*" -Destination "\\$using:AV_CAP_VM_COMPNAME\C$\Windows\Logs" -Force -Recurse -Confirm:$false }
$SourceCopied | Wait-Job


# SETUP A REMOTE POWERSHELL SESSION TO APP VOLUMES CAPTURE VM
$CredPass = ConvertTo-SecureString -String $AV_CAP_PWD -AsPlainText -Force
$Credentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList ($AV_CAP_USR, $CredPass)
$PSSession = New-PSSession -ComputerName "$AV_CAP_VM_COMPNAME" -Credential $Credentials


# WAITING TILL WINDOWS REMOTE MANAGEMENT SERVICE (WINRM) IS STARTED
Invoke-Command -Session $PSSession -ScriptBlock {
    Write-Host "The Windows Remote Management Service (WinRM) must run before continue with the script!" -ForegroundColor "Cyan"
    DO {
        $WinRMState = Get-Service -Name WinRM
        Write-Host "Waiting till Windows Remote Management service (WinRM) is running on $AV_CAP_VM_COMPNAME, status is $($WinRMState.Status)..."
        Start-Sleep -s 5
    }
    UNTIL ($WinRMState.Status -eq "Running")
    Write-Host "The WinRM service is running, script now continues..." -ForegroundColor "Green"
    Write-Host "`n"
}


# CREATE TEMPORARY SCRIPT TO HANDLE LOCAL TASKS OF FINISHING THE APPLICATION PACKAGE
Invoke-Command -Session $PSSession -ScriptBlock {  
    Add-Content "C:\Windows\Temp\AppForeground_and_OK.ps1" "Start-Sleep -s 8"
    Add-Content "C:\Windows\Temp\AppForeground_and_OK.ps1" "`$wshell = New-Object -ComObject wscript.shell;"
    Add-Content "C:\Windows\Temp\AppForeground_and_OK.ps1" "`$wshell.AppActivate(`'VMware App Volumes`')"
    Add-Content "C:\Windows\Temp\AppForeground_and_OK.ps1" "Start-Sleep -s 5"
    Add-Content "C:\Windows\Temp\AppForeground_and_OK.ps1" "`$wshell.SendKeys(`'~`')"
    Add-Content "C:\Windows\Temp\AppForeground_and_OK.ps1" "`$wshell.AppActivate(`'VMware App Volumes`')"
    Add-Content "C:\Windows\Temp\AppForeground_and_OK.ps1" "Start-Sleep -s 5"
    Add-Content "C:\Windows\Temp\AppForeground_and_OK.ps1" "`$wshell.SendKeys(`'~`')"
    Add-Content "C:\Windows\Temp\AppForeground_and_OK.ps1" "`$wshell.AppActivate(`'VMware App Volumes`')"
    Add-Content "C:\Windows\Temp\AppForeground_and_OK.ps1" "Start-Sleep -s 25"
    Add-Content "C:\Windows\Temp\AppForeground_and_OK.ps1" "`$wshell.SendKeys(`'~`')"
    Add-Content "C:\Windows\Temp\AppForeground_and_OK.ps1" "`$wshell.AppActivate(`'VMware App Volumes`')"
    Add-Content "C:\Windows\Temp\AppForeground_and_OK.ps1" "Start-Sleep -s 5"
    Add-Content "C:\Windows\Temp\AppForeground_and_OK.ps1" "`$wshell.SendKeys(`'~`')"
}


# CREATE A SCHEDULED TASK FOR RUNNING THE LOCAL TASKS FOR FINISHING THE APPLICATION PACKAGE
Invoke-Command -Session $PSSession -ScriptBlock {
    $User= "$using:AV_CAP_DOM\$using:AV_CAP_USR" # Specify the account to run the script
    $Action= New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"C:\Windows\Temp\AppForeground_and_OK.ps1`""
    Register-ScheduledTask -TaskName "FinishingAVCapture" -User $User -Action $Action -RunLevel Highest –Force | Out-Null
}


# GET APP VOLUMES CAPTURE VM ID & UUID
$AV_CAP_VM_API = ((Invoke-WebRequest -WebSession $Login -Method Get -Uri "$AV_MGR_REST_API_PROT`://$av_mgr_srv/cv_api/machines").content | ConvertFrom-Json).machines | Where-Object {$_.name -eq $AV_CAP_VM_NAME -and $_.status -eq "Existing"}
$AV_CAP_VM_API_ID = $AV_CAP_VM_API.id
$AV_CAP_VM_API_UUID = $AV_CAP_VM_API.identifier


# GET APP VOLUMES APPLICATION PACKAGE ID
$AV_APP_PKG_API = ((Invoke-WebRequest -WebSession $Login -Method Get -Uri "$AV_MGR_REST_API_PROT`://$av_mgr_srv/app_volumes/app_packages").content | ConvertFrom-Json).data | Where-Object {$_.name -eq $AV_APP_PKG_NAME}
$AV_APP_PKG_API_ID = $AV_APP_PKG_API.id


# START THE APPLICATION CAPTURE (FROM THIS MOMENT EVERYTHING WILL BE CAPTURED INSIDE THE APPLICATION PACKAGE (APPSTACK) !)
Write-Host "From this moment everything will be captured inside the application package..." -ForegroundColor "Yellow"
Write-Host "`n"
Invoke-WebRequest -WebSession $Login -Method Post -Uri "$AV_MGR_REST_API_PROT`://$av_mgr_srv/app_volumes/app_packages/$AV_APP_PKG_API_ID/start_package?data[computer_id]=$AV_CAP_VM_API_ID&data[uuid]=$AV_CAP_VM_API_UUID" | Out-Null
Start-Sleep -s 10


# INSTALL SOFTWARE
Write-Host "Starting the installation of $AV_APP_PKG_NAME on $AV_CAP_VM_NAME..."
IF ($AV_SOURCE_INST -like "*.msi") { 
    $InstallExitCode = Invoke-Command -Session $PSSession -ScriptBlock {
        $InstallResult = (Start-Process -Wait -PassThru "msiexec.exe" -ArgumentList "$using:AV_SOURCE_PARM").ExitCode
        return $InstallResult
    }
	IF ($InstallExitCode -eq 0) {
        Write-Host "MSI Installation successful with exitcode: $InstallExitCode !" -ForegroundColor "Green"
        IF ($AV_SOURCE_PSTCFG_OPTIONS) {
            Write-Host "Executing post-configuration steps to make application VDI ready, if they are entered..." -ForegroundColor "Cyan"
		    Write-Host "`n"
            Invoke-Command -Session $PSSession -ScriptBlock {
			
                ForEach ($AV_SOURCE_PSTCFG_OPTION in $using:AV_SOURCE_PSTCFG_OPTIONS) {
				    Invoke-Expression "$AV_SOURCE_PSTCFG_OPTION"
			    }
            }
        }
	}
    ELSEIF ($InstallExitCode -eq 3010) {
        Write-Host "MSI Installation successful with exitcode: $InstallExitCode , reboot is required and take place automatically!" -ForegroundColor "Green"
        IF ($AV_SOURCE_PSTCFG_OPTIONS) {
            Write-Host "Executing post-configuration steps to make application VDI ready, if they are entered..." -ForegroundColor "Cyan"
		    Write-Host "`n"
            Invoke-Command -Session $PSSession -ScriptBlock {		
                ForEach ($AV_SOURCE_PSTCFG_OPTION in $using:AV_SOURCE_PSTCFG_OPTIONS) {
				    Invoke-Expression "$AV_SOURCE_PSTCFG_OPTION"
			    }
                Write-Host "Waiting for the App Volumes Capture VM is booted and the WMI + WinRM services are running!" -ForegroundColor "Yellow"
                shutdown /r /t 5
            }
            # WAITING TILL APP VOLUMES CAPTURE VM IS REBOOTED
            Start-Sleep-ProgressBar -s $AV_CAP_VM_REBOOT_TIME
            
            # WAITING TILL WINDOWS REMOTE MANAGEMENT SERVICE (WINRM) IS STARTED
            Invoke-Command -Session $PSSession -ScriptBlock {
                Write-Host "The Windows Remote Management Service (WinRM) must run before continue with the script!" -ForegroundColor "Cyan"
                DO {
                    $WinRMState = Get-Service -Name WinRM
                    Write-Host "Waiting till Windows Remote Management service (WinRM) is running on $AV_CAP_VM_COMPNAME, status is $($WinRMState.Status)..."
                    Start-Sleep -s 5
                }
                UNTIL ($WinRMState.Status -eq "Running")
                Write-Host "The WinRM service is running, script now continues..." -ForegroundColor "Green"
                Write-Host "`n"
            }
        }
        ELSE {																				 
            Write-Host "Waiting for the App Volumes Capture VM is booted and the WMI + WinRM services are running!" -ForegroundColor "Yellow"
            Invoke-Command -Session $PSSession -ScriptBlock {
                shutdown /r /t 5
            }
            # WAITING TILL APP VOLUMES CAPTURE VM IS REBOOTED
            Start-Sleep-ProgressBar -s $AV_CAP_VM_REBOOT_TIME
            
            # WAITING TILL WINDOWS REMOTE MANAGEMENT SERVICE (WINRM) IS STARTED
            Invoke-Command -Session $PSSession -ScriptBlock {
                Write-Host "The Windows Remote Management Service (WinRM) must run before continue with the script!" -ForegroundColor "Cyan"
                DO {
                    $WinRMState = Get-Service -Name WinRM
                    Write-Host "Waiting till Windows Remote Management service (WinRM) is running on $AV_CAP_VM_COMPNAME, status is $($WinRMState.Status)..."
                    Start-Sleep -s 5
                }
                UNTIL ($WinRMState.Status -eq "Running")
                Write-Host "The WinRM service is running, script now continues..." -ForegroundColor "Green"
                Write-Host "`n"
            }
        }
    }
	ELSE {
		Write-Host "MSI Installation failed with exitcode: $InstallExitCode !" -ForegroundColor "Red"
		Write-Host "`n"
        Stop-Transcript
        pause
	    exit
	}
} 
ELSEIF ($AV_SOURCE_INST -like "*.exe") {
    $InstallExitCode = Invoke-Command -Session $PSSession -ScriptBlock {
        $InstallResult = (Start-Process -Wait -PassThru -FilePath "C:\Windows\Logs\$using:AV_SOURCE_INST" -ArgumentList "$using:AV_SOURCE_PARM").ExitCode
        return $InstallResult
    }
    IF ($InstallExitCode -eq 0) {
	    Write-Host "EXE Installation successful with exitcode: $InstallExitCode !" -ForegroundColor "Green"    
        IF ($AV_SOURCE_PSTCFG_OPTIONS) {
            Write-Host "Executing post-configuration steps to make application VDI ready, if they are entered..." -ForegroundColor "Cyan"
		    Write-Host "`n"
            Invoke-Command -Session $PSSession -ScriptBlock {
			
                ForEach ($AV_SOURCE_PSTCFG_OPTION in $using:AV_SOURCE_PSTCFG_OPTIONS) {
				    Invoke-Expression "$AV_SOURCE_PSTCFG_OPTION"
			    }
            }
        }
    }
    ELSEIF ($InstallExitCode -eq 3010) {
        Write-Host "EXE Installation successful with exitcode: $InstallExitCode , reboot is required and take place automatically!" -ForegroundColor "Green"
        IF ($AV_SOURCE_PSTCFG_OPTIONS) {
            Write-Host "Executing post-configuration steps to make application VDI ready, if they are entered..." -ForegroundColor "Cyan"
		    Write-Host "`n"
            Invoke-Command -Session $PSSession -ScriptBlock { 
                ForEach ($AV_SOURCE_PSTCFG_OPTION in $using:AV_SOURCE_PSTCFG_OPTIONS) {
				    Invoke-Expression "$AV_SOURCE_PSTCFG_OPTION"
			    }
                Write-Host "Waiting for the App Volumes Capture VM is booted and the WMI + WinRM services are running!" -ForegroundColor "Yellow"
                shutdown /r /t 5
            }
            # WAITING TILL APP VOLUMES CAPTURE VM IS REBOOTED
            Start-Sleep-ProgressBar -s $AV_CAP_VM_REBOOT_TIME
            
            # WAITING TILL WINDOWS REMOTE MANAGEMENT SERVICE (WINRM) IS STARTED
            Invoke-Command -Session $PSSession -ScriptBlock {
                Write-Host "The Windows Remote Management Service (WinRM) must run before continue with the script!" -ForegroundColor "Cyan"
                DO {
                    $WinRMState = Get-Service -Name WinRM
                    Write-Host "Waiting till Windows Remote Management service (WinRM) is running on $AV_CAP_VM_COMPNAME, status is $($WinRMState.Status)..."
                    Start-Sleep -s 5
                }
                UNTIL ($WinRMState.Status -eq "Running")
                Write-Host "The WinRM service is running, script now continues..." -ForegroundColor "Green"
                Write-Host "`n"
            }
        }
        ELSE {																					 
            Write-Host "Waiting for the App Volumes Capture VM is booted and the WMI + WinRM services are running!" -ForegroundColor "Yellow"
            Invoke-Command -Session $PSSession -ScriptBlock {
                shutdown /r /t 5
            }
            # WAITING TILL APP VOLUMES CAPTURE VM IS REBOOTED
            Start-Sleep-ProgressBar -s $AV_CAP_VM_REBOOT_TIME
            
            # WAITING TILL WINDOWS REMOTE MANAGEMENT SERVICE (WINRM) IS STARTED
            Invoke-Command -Session $PSSession -ScriptBlock {
                Write-Host "The Windows Remote Management Service (WinRM) must run before continue with the script!" -ForegroundColor "Cyan"
                DO {
                    $WinRMState = Get-Service -Name WinRM
                    Write-Host "Waiting till Windows Remote Management service (WinRM) is running on $AV_CAP_VM_COMPNAME, status is $($WinRMState.Status)..."
                    Start-Sleep -s 5
                }
                UNTIL ($WinRMState.Status -eq "Running")
                Write-Host "The WinRM service is running, script now continues..." -ForegroundColor "Green"
                Write-Host "`n"
            }
        }
    } 
	ELSE {
		Write-Host "EXE Installation failed with exitcode: $InstallExitCode !" -ForegroundColor "Red"
		Write-Host "`n"
        Stop-Transcript
        pause
	    exit
	}
}
ELSE {
    Write-Host "No setup and/or installers found!" -ForegroundColor "Red"
	Write-Host "`n"
    Stop-Transcript
    pause
	exit
}


# SETUP A REMOTE POWERSHELL SESSION TO APP VOLUMES CAPTURE VM AFTER REBOOT
$Credentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList ($AV_CAP_USR, $CredPass)
$PSSession = New-PSSession -ComputerName "$AV_CAP_VM_COMPNAME" -Credential $Credentials


# BRING APP VOLUMES CAPTURE WINDOW TO FOREGROUND AND FINISH CAPTURE PROCESS
Invoke-Command -Session $PSSession -ScriptBlock {
    Start-ScheduledTask -TaskName "FinishingAVCapture" | Out-Null
}
Write-Host "Waiting till scheduled task to finish the capture process, before continue the script..." -ForegroundColor "Cyan"
Start-Sleep-ProgressBar -s 50 # Higher up the seconds if scheduledtask takes a longer time.
Write-Host "The scheduled task is completed and will automatically reboot..." -ForegroundColor "Green"
Start-Sleep-ProgressBar -s $AV_CAP_VM_REBOOT_TIME
Write-Host "`n"


# WAITING TILL THE WINDOWS REMOTE MANAGEMENT (WINRM) SERVICE AND APP VOLUMES CAPTURE PROCESS STARTED AFTER REBOOT
Write-Host "The App Volumes capture process must run to finish the App Volumes capture process correctly!" -ForegroundColor "Cyan"
WHILE ($null -eq (Get-CimInstance Win32_Process -ComputerName $AV_CAP_VM_COMPNAME -Filter "Name = 'svservice.exe'" -ErrorAction SilentlyContinue)) {
    Write-Host "Waiting till the Windows Remote Management (WinRM) service and the App Volumes capture process is started on $AV_CAP_VM_COMPNAME..."
    Start-Sleep -s 5
}											
Write-Host "The App Volumes service running, script now continues..." -ForegroundColor "Green"
Write-Host "The App Volumes capture process is complete!" -ForegroundColor "Green"
Write-Host "`n"


# REVERT SNAPSHOT OF APP VOLUMES CAPTURE VM TO CLEAN STATE
Write-Host "Reverting: $AV_CAP_VM_NAME to a clean state..." -ForegroundColor "Cyan"
TRY {
	Set-VM -VM "$AV_CAP_VM_NAME" -SnapShot "$AV_CAP_VM_SS" -Confirm:$false | Out-Null
	Write-Host "Reverting to a clean state is completed!" -ForegroundColor "Green"
	Write-Host "`n"
}
CATCH {
	Write-Host "Reverting to a clean state failed..." -ForegroundColor "Red"
	Write-Host "Press Ctrl + C to abort the script..." -ForegroundColor "Cyan"
	Write-Host "`n"
    Stop-Transcript
    pause
	exit
}


# CLOSE VMWARE REMOTE CONSOLE WINDOW
IF ($null -eq (Get-Process -ProcessName vmrc -ErrorAction SilentlyContinue)) {
    Write-Host "The VMware Remote Console (VMRC) is process is not found, probably already closed!" -ForegroundColor "Yellow"
	Write-Host "`n"
}
ELSE {
    Get-Process -ProcessName vmrc | Where-Object { $_.CloseMainWindow() | Out-Null }
    Write-Host "The VMware Remote Console (VMRC) is closed succesfully!" -ForegroundColor "Green"
	Write-Host "`n"
}


# SCRIPT FINISHED
Write-Host "The App Volumes 4.x - Application (Package) is now ready for use, assign the application to an Active Directory user or user-group!" -ForegroundColor "Green"


#####################################################################
######################## STOP LOGGING OUTPUT ########################
#####################################################################

Stop-Transcript

#####################################################################
######################## END SCRIPT SECTION #########################
#####################################################################