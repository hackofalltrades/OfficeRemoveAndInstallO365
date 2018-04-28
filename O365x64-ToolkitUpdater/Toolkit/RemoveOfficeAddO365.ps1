<#
.SYNOPSIS
 This script performs the installation or uninstallation of an application(s).
.DESCRIPTION
 The script is provided as a template to perform an install or uninstall of an application(s).
 The script either performs an "Install" deployment type or an "Uninstall" deployment type.
 The install deployment type is broken down in to 3 main sections/phases: Pre-Install, Install, and Post-Install.
 The script dot-sources the AppDeployToolkitMain.ps1 script which contains the logic and functions required to install or uninstall an application.
 To access the help section,
.EXAMPLE
 Deploy-Application.ps1
.EXAMPLE
 Deploy-Application.ps1 -DeploymentType "Silent"
.EXAMPLE
 Deploy-Application.ps1 -AllowRebootPassThru -AllowDefer
.EXAMPLE
 Deploy-Application.ps1 -Uninstall
.PARAMETER DeploymentType
 The type of deployment to perform. [Default is "Install"]
.PARAMETER DeployMode
 Specifies whether the installation should be run in Interactive, Silent or NonInteractive mode.
 Interactive = Default mode
 Silent = No dialogs
 NonInteractive = Very silent, i.e. no blocking apps. Noninteractive mode is automatically set if an SCCM task sequence or session 0 is detected.
.PARAMETER AllowRebootPassThru
 Allows the 3010 return code (requires restart) to be passed back to the parent process (e.g. SCCM) if detected from an installation.
 If 3010 is passed back to SCCM a reboot prompt will be triggered.
.NOTES
.LINK
 Http://psappdeploytoolkit.codeplex.com
"#>
Param (
 [ValidateSet("Install","Uninstall")]
 [string] $DeploymentType = "Install",
 [ValidateSet("Interactive","Silent","NonInteractive")]
 [string] $DeployMode = "Interactive",
 [switch] $AllowRebootPassThru = $false
 
)
 
#*===============================================
#* VARIABLE DECLARATION
#Try {
#*===============================================
 
#*===============================================
# Variables: Application
 
$appVendor = "Microsoft"
$appName = "Office"
$appVersion = "365 Pro Plus x64"
$appArch = "x64"
$appLang = "EN"
$appRevision = "15.0.4551.1011"
$appScriptVersion = "2.0.1"
$appScriptDate = "04/14/2017"
$appScriptAuthor = "Greg King"
 
#*===============================================
# Variables: Script - Do not modify this section
 
$deployAppScriptFriendlyName = "Deploy Application"
$deployAppScriptVersion = "3.0.6"
$deployAppScriptDate = "10/10/2013"
$deployAppScriptParameters = $psBoundParameters
 
# Variables: Environment
$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Definition
# Dot source the App Deploy Toolkit Functions
."$scriptDirectory\AppDeployToolkit\AppDeployToolkitMain.ps1"
 
# Office Directory
$dirOffice = Join-Path "$envProgramFilesX86" "Microsoft Office"
 
#*===============================================
#* END VARIABLE DECLARATION
#*===============================================
 
#*===============================================
#* PRE-INSTALLATION
If ($deploymentType -ne "uninstall") { $installPhase = "Pre-Installation"
#*===============================================
 
 # Show Welcome Message, close Internet Explorer if required, allow up to 3 deferrals, and verify there is enough disk space to complete the install
 Show-InstallationWelcome -CloseApps "iexplore,PWConsole,excel,groove,onenote,infopath,onenote,outlook,mspub,powerpnt,winword,communicator,lync" -BlockExecution -AllowDefer -DeferTimes 3 -CheckDiskSpace
 
# Check whether anything might prevent us from running the cleanup
 If (($isServerOS -eq $true)) {
 Write-Log "Installation of components has been skipped as one of the following options are enabled. isServerOS: $isServerOS"
 }
 
 # Display Pre-Install cleanup status
 Show-InstallationProgress "Performing Pre-Install cleanup. This may take some time. Please wait..."
 
# Remove any previous version of Office (if required)
 $officeExecutables = @("excel.exe", "groove.exe", "onenote.exe", "infopath.exe", "onenote.exe", "outlook.exe", "mspub.exe", "powerpnt.exe", "winword.exe", "winproj.exe", "lync.exe")
 ForEach ($officeExecutable in $officeExecutables) {
 If (Test-Path (Join-Path $dirOffice "Office12\$officeExecutable")) {
 Write-Log "Microsoft Office 2007 was detected. Will be uninstalled."
 Execute-Process -FilePath "CScript.Exe" -Arguments "`"$dirSupportFiles\OffScrub07.vbs`" STANDARD,PROPLUS,PROOFKIT /S /Q /NoCancel" -WindowStyle Hidden -IgnoreExitCodes "1,2,3"
 Break
 }
 }
 ForEach ($officeExecutable in $officeExecutables) {
 If (Test-Path (Join-Path $dirOffice "Office14\$officeExecutable")) {
 Write-Log "Microsoft Office 2010 was detected. Will be uninstalled."
 Execute-Process -FilePath "CScript.Exe" -Arguments "`"$dirSupportFiles\OffScrub10.vbs`" STANDARD,PROPLUS,PROOFKIT /S /Q /NoCancel" -WindowStyle Hidden -IgnoreExitCodes "1,2,3"
 Break
 }
 }
 ForEach ($officeExecutable in $officeExecutables) {
 If (Test-Path (Join-Path $dirOffice "Office15\$officeExecutable")) {
 Write-Log "Microsoft Office 2013 was detected. Will be uninstalled."
 Execute-Process -FilePath "CScript.Exe" -Arguments "`"$dirSupportFiles\OffScrub13.vbs`" ALL /S /Q /NoCancel" -WindowStyle Hidden -IgnoreExitCodes "1,2,3"
 Break
 }
 }
 ForEach ($officeExecutable in $officeExecutables) {
 If (Test-Path (Join-Path $dirOffice "Office16\$officeExecutable")) {
 Write-Log "Microsoft Office 2016 was detected. Will be uninstalled."
 Execute-Process -FilePath "CScript.Exe" -Arguments "`"$dirSupportFiles\OffScrub16.vbs`" CLIENTALL /S /Q /NoCancel" -WindowStyle Hidden -IgnoreExitCodes "1,2,3"
 Break
 }
 }
  ForEach ($officeExecutable in $officeExecutables) {
 If (Test-Path "C:\Program Files (x86)\Microsoft Office\root\Office16\$officeexecutable") {
 Write-Log "Microsoft Office 2016 was detected. Will be uninstalled."
 Execute-Process -FilePath "$dirFiles\Office 365 ProPlus\setup.exe" -Arguments " /configure `"$dirFiles\Office 365 ProPlus\uninstall.xml`"" -WindowStyle Hidden -IgnoreExitCodes "3010"
 Break
 }
 }

 # Remove Microsoft Guide to The Ribbion
 Remove-MSIApplications "Microsoft Guide to the Ribbon"
 
 # Remove Microsoft Office 2007 Help Tab
 Remove-MSIApplications "Microsoft Office 2007 Help Tab"
 
 # Remove Microsoft Conferencing Add-in for Microsoft Office Outlook
 Remove-MSIApplications "Microsoft Conferencing Add-in for Microsoft Office Outlook"
 
 # Remove Microsoft Office Live Meeting 2007
 Remove-MSIApplications "Microsoft Office Live Meeting 2007"
 
 # Remove Microsoft Office 2010 Interactive Guide
 Remove-MSIApplications "Microsoft Office 2010 Interactive Guide"
 
 # Remove Microsoft Office 2010 User Resources
 Remove-MSIApplications "Office 2010 User Resources"
 
# Remove Microsoft Office Communicator 2007
 Remove-MSIApplications "Microsoft Office Communicator 2007"
 
#*===============================================
#* INSTALLATION
$installPhase = "Installation"
#*===============================================
 
# Installing Office 365 Pro Plus
 Show-InstallationProgress "Installing Office 365 Pro Plus x64. This may take some time. Please wait..."
 Execute-Process -FilePath "$dirFiles\Office 365 ProPlus\setup.exe" -Arguments " /configure `"$dirFiles\Office 365 ProPlus\configuration.xml`"" -WindowStyle Hidden -IgnoreExitCodes "3010"
 
#*===============================================
#* POST-INSTALLATION
$installPhase = "Post-Installation"
#*===============================================
 
# Install Office 2013 Proofing Tools with all languages
# Show-InstallationProgress "Installing Office Proofing Tools. This may take some time. Please wait..."
# Execute-Process -FilePath "$dirFiles\ProofingTools\setup.exe" -Arguments " /config `"$dirFiles\ProofingTools\W7_MSOffice2013_ProfingTools_config.xml`"" -WindowStyle Hidden -IgnoreExitCodes "3010"
 
# Remove Office 2013 Proofing Tools IME. Removes Asian keyboard from default setup
# Show-InstallationProgress "Removing Office 2013 Proofing Tools IME"
# Execute-Process -FilePath "regedit.exe" -Arguments " /s `"$dirSupportFiles\Proofing_Tools_2013_Remove_IME.reg`""
 
 # Add Custom Outlook 2013 Shortcut for NK2 migration.
# Show-InstallationProgress "Installing Office 365 Shortcuts"
# Copy-File -Path "$dirSupportFiles\Outlook 2013.lnk" -Destination "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office 2013\Outlook 2013.lnk"
 
 # Disable Lync Server Check. Makes it possible for Lync to communicate with the Communicatoer 2007 R2 Server Pool
# Show-InstallationProgress "Disabling Lync Server Check"
## Execute-Process -FilePath "regedit.exe" -Arguments " /s `"$dirSupportFiles\Lync_Disable_Server_Check.reg`""
 
# Install Microsoft Conferencing Add-in for Microsoft Office Outlook
# If ($deploymentType -eq "KeepCommunicator") {
# Show-InstallationProgress "Installing Conferencing Add-in for Microsoft Office Outlook. This may take some time. Please wait..."
# Execute-MSI -Action Install -Path "$dirFiles\Microsoft Conferencing Add-in\ConfAddin.msi" -Arguments " /q ALLUSERS=2 /i /l*v `"C:\Windows\WW-Group\SWDLogs\ConferencingAddinOutlook806362202.log`""
 }
 # Prompt for a restart (if running as a user, not installing components and not running on a server)
 #If (($deployMode -eq "Interactive") -and ($IsServerOS -eq $false)) {
 #Show-InstallationRestartPrompt -Countdownseconds 14400 -CountdownNoHideSeconds 300
#}
#*===============================================
#* UNINSTALLATION
#} ElseIf ($deploymentType -eq "uninstall") { $installPhase = "Uninstallation"
#*===============================================
 
# Show Welcome Message, close applications if required with a 60 second countdown before automatically closing
 Show-InstallationWelcome -CloseApps "excel,groove,onenote,infopath,onenote,outlook,mspub,powerpnt,winword,winproj,visio,communicator,lync"
 
# Show Progress Message (with the default message)
 #Show-InstallationProgress "Removing Office 365 Pro Plus with Proofing Tools"
 #Execute-Process -FilePath "$dirFiles\Office 365 ProPlus\setup.exe" -Arguments " /configure `"$dirFiles\Office 365 ProPlus\Remove.xml`"" -WindowStyle Hidden -IgnoreExitCodes "3010"
 
# Remove Microsoft Office Proofing Tools 2013
 #Remove-MSIApplications "Microsoft Office Proofing Tools 2013"
 
#*===============================================
#* END SCRIPT BODY
#} } Catch {$exceptionMessage = "$($_.Exception.Message) `($($_.ScriptStackTrace)`)"; Write-Log "$exceptionMessage"; Show-DialogBox -Text $exceptionMessage -Icon "Stop"; Exit-Script -ExitCode 1} # Catch any errors in this script
Exit-Script -ExitCode 0 # Otherwise call the Exit-Script function to perform final cleanup operations
#*===============================================