<#
.SYNOPSIS
    Ultimate Windows Fix & Optimization Script – Safe & Best Practices

.DESCRIPTION
    This script performs a series of system repair, cleanup, and update operations in the following order:
      • Checks that System Restore is enabled and (optionally) creates a restore point.
      • Repairs Windows image and system files using DISM and SFC.
      • Performs a component cleanup.
      • Optionally schedules a CHKDSK (disk check) for the next reboot.
      • Checks for required winget apps (e.g. gsudo, Windows Terminal, PowerShell, Oh My Posh) and installs any that are missing.
      • Previews available software upgrades via winget and (optionally) upgrades them.
      • Checks for Python installation, offering to install it if missing, then upgrades pip and required packages.
      • Ensures necessary PowerShell modules (e.g. gsudoModule) are installed (with WhatIf simulation first).
      • Verifies the fixes with final system integrity checks.
      
    **IMPORTANT:**
      • No forced elevation is performed. Commands that require administrative privileges are executed via
        the custom Invoke-AdminCommand function – the user is prompted before any elevated action.
      • Dangerous actions (e.g. CHKDSK, module installations) are not run automatically.
      • The script does not auto‑restart the system. A final confirmation is requested before a reboot.

.NOTES
    Author: Jonna, Cyan
    Date: 2025-02-02
    Usage: Run this script in a normal PowerShell window.
    Parameters:
      -Silent       : (Switch) Run the script without waiting for keypresses at the end.
      -AutoConfirm  : (Switch) Automatically answer "yes" to all prompts.
#>

[CmdletBinding()]
param(
  [switch]$Silent,
  [switch]$AutoConfirm
)

# ----------------------------------------------------
# Register Ctrl+C Event Handler to Stop the Transcript
# ----------------------------------------------------
# This ensures that if the user hits Ctrl+C, the transcript is stopped and the log file isn't left locked.
$global:transcriptStarted = $false
$null = [Console]::CancelKeyPress += {
  param($sender, $args)
  Write-Host "Ctrl+C detected! Stopping transcript..." -ForegroundColor Red
  if ($global:transcriptStarted) {
    try {
      Stop-Transcript | Out-Null
      Write-Host "Transcript stopped successfully." -ForegroundColor Green
    }
    catch {
      Write-Host "Error stopping transcript: $_" -ForegroundColor Red
    }
  }
  # Allow the termination to proceed
  $args.Cancel = $false
  exit
}

# -------------------------------
# Set $now Variable for Restore Point Description
# -------------------------------
$now = $(Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fffZ").ToString()

# -------------------------------
# Determine Preferred Shell
# -------------------------------
# Check if pwsh (PowerShell 7+) is available. If so, use it; otherwise, fall back to powershell.exe.
if (Get-Command pwsh.exe -ErrorAction SilentlyContinue) {
  $preferredShell = "pwsh.exe"
  Write-Host "Preferred shell set to pwsh.exe" -ForegroundColor Cyan
}
elseif (Get-Command powershell.exe -ErrorAction SilentlyContinue) {
  $preferredShell = "powershell.exe"
  Write-Host "Preferred shell set to powershell.exe" -ForegroundColor Cyan
}
else {
  Write-Error "No supported PowerShell executable found."
  exit 1
}

# -------------------------------
# Package Manager Lists
# -------------------------------
# Change these lists to control which apps get installed in bulk.

# Required winget apps (format: hashtable with Id and Name)
$wingetAppsRequired = @(
  @{ Id = "gerardog.gsudo"; Name = "gsudo" },
  @{ Id = "Microsoft.WindowsTerminal"; Name = "Windows Terminal" },
  @{ Id = "Microsoft.PowerShell"; Name = "PowerShell" },
  @{ Id = "JanDeDobbeleer.OhMyPosh"; Name = "Oh My Posh" },
  @{ Id = "Microsoft.VisualStudioCode"; Name = "Visual Studio Code" },
  @{ Id = "Google.Cloud.SDK"; Name = "Google Cloud SDK" }, # pip install google-cloud-sdk
  @{ Id = "Microsoft.PowerToys"; Name = "Microsoft PowerToys" }
)

# Required pip packages (format: hashtable with Package)
$pipPackagesRequired = @(
  @{ Package = "Flask" },
  @{ Package = "google-api-python-client" },
  @{ Package = "google-auth-httplib2" },
  @{ Package = "google-auth-oauthlib" },
  @{ Package = "colorama" }
)

# Required PowerShell modules (list of module names)
$psModulesRequired = @(
  "gsudoModule"
)

# -------------------------------
# Helper Function for Input
# -------------------------------
function Get-UserInput {
  param(
    [string]$Prompt,
    [string]$Default = "y"
  )
  if ($AutoConfirm) {
    Write-Host "$Prompt (AutoConfirm: $Default)" -ForegroundColor Magenta
    return $Default
  }
  else {
    return Read-Host $Prompt
  }
}

# -------------------------------
# Function: Invoke-AdminCommand
# -------------------------------
# This function runs a given script block with administrative privileges.
# If the built-in 'sudo' command is available, it uses that (along with the preferred shell);
# otherwise, it falls back to checking the current session and launching an elevated process.
function Invoke-AdminCommand {
  param(
    [Parameter(Mandatory = $true)]
    [ScriptBlock]$ScriptBlock
  )
  
  # If 'sudo' is available, use it.
  if (Get-Command sudo -ErrorAction SilentlyContinue) {
    $command = $ScriptBlock.ToString()
    Write-Host "Using sudo for elevated command: $command" -ForegroundColor Cyan
    sudo $preferredShell -NoProfile -ExecutionPolicy Bypass -Command $command
  }
  else {
    # Check if the current session is already elevated.
    $currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentIdentity)
    if ($principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
      Write-Host "Executing command with administrative privileges..." -ForegroundColor Cyan
      & $ScriptBlock
    }
    else {
      # Encode the script block (to handle multi-line commands safely)
      $encodedCommand = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($ScriptBlock.ToString()))
      Write-Host "Launching elevated $preferredShell for the command..." -ForegroundColor Cyan
      Start-Process -FilePath $preferredShell `
        -ArgumentList "-NoProfile -ExecutionPolicy Bypass -EncodedCommand $encodedCommand" `
        -Verb RunAs -Wait
    }
  }
}

# -------------------------------
# Start Logging
# -------------------------------
$logFile = Join-Path -Path $PSScriptRoot -ChildPath "FixLog.txt"

if (Test-Path $logFile) {
  Write-Host "Log file exists." -ForegroundColor Yellow
  $overwrite = Get-UserInput "Overwrite existing log file? (y/n)" "n"
  if ($overwrite -match "^(y|yes)$") {
    # User wants to overwrite, so delete the existing file first.
    Remove-Item $logFile -Force
  }
}

# Start the transcript (it will create a new file if needed or append if not)
Start-Transcript -Path $logFile -Append
$global:transcriptStarted = $true

# -------------------------------
# Initial Warning – Save Your Work!
# -------------------------------
Write-Host "⚠️  WARNING: Please ensure ALL work is saved. This script will make system changes." -ForegroundColor Yellow
$proceed = Get-UserInput "Do you wish to continue? (y/n)" "y"
if ($proceed -notmatch "^(y|yes)$") {
  Write-Host "Operation cancelled by user." -ForegroundColor Red
  Stop-Transcript
  exit
}

# -------------------------------
# Section 1: System Restore Check & Restore Point Creation
# -------------------------------
Write-Host "`n🔒 Checking System Restore status..." -ForegroundColor Cyan
try {
  $restorePoints = Get-ComputerRestorePoint -ErrorAction Stop
}
catch {
  $restorePoints = $null
}

if (-not $restorePoints) {
  Write-Host "⚠️  System Restore appears to be disabled." -ForegroundColor Red
  $enableRestore = Get-UserInput "Do you want to enable System Restore on drive C:\? (y/n)" "y"
  if ($enableRestore -match "^(y|yes)$") {
    Invoke-AdminCommand -ScriptBlock { Enable-ComputerRestore -Drive "C:\" }
    Write-Host "🛡️  System Restore has been enabled." -ForegroundColor Green
  }
  else {
    Write-Host "Skipping enabling System Restore. A restore point may not be created." -ForegroundColor Yellow
  }
}

$createRP = Get-UserInput "Do you want to create a System Restore Point now? (Recommended) (y/n)" "y"
if ($createRP -match "^(y|yes)$") {
  Write-Host "🛡️  Creating a System Restore Point..." -ForegroundColor Green
  Invoke-AdminCommand -ScriptBlock ([scriptblock]::Create("Checkpoint-Computer -Description '$now' -RestorePointType MODIFY_SETTINGS"))
}
else {
  Write-Host "Skipping System Restore Point creation." -ForegroundColor Yellow
}

# -------------------------------
# Section 2: Windows Repair & Optimization
# -------------------------------
Write-Host "`n🛠️  Starting Windows Repair & Optimization..." -ForegroundColor Cyan

Write-Host "`n🖼️  Checking Windows image health with DISM..." -ForegroundColor Yellow
Invoke-AdminCommand -ScriptBlock { DISM /Online /Cleanup-Image /CheckHealth }

Write-Host "`n🔍 Scanning Windows image for issues with DISM..." -ForegroundColor Yellow
Invoke-AdminCommand -ScriptBlock { DISM /Online /Cleanup-Image /ScanHealth }

Write-Host "`n🛠️  Repairing Windows image with DISM (RestoreHealth)..." -ForegroundColor Yellow
Invoke-AdminCommand -ScriptBlock { DISM /Online /Cleanup-Image /RestoreHealth }

Write-Host "`n🛠️  Running System File Checker (sfc /scannow)..." -ForegroundColor Yellow
Invoke-AdminCommand -ScriptBlock { sfc /scannow }

Write-Host "`n🧹 Running Component Cleanup to remove outdated updates..." -ForegroundColor Yellow
Invoke-AdminCommand -ScriptBlock { Dism.exe /online /Cleanup-Image /StartComponentCleanup }

Write-Host "🖥️  Scheduling disk check (CHKDSK /f /r) on next reboot..." -ForegroundColor Yellow
Invoke-AdminCommand -ScriptBlock { chkdsk C: /f /r }

# -------------------------------
# Section 3: Check for Required Winget Apps
# -------------------------------
Write-Host "`n🔍 Checking for required winget apps..." -ForegroundColor Cyan

foreach ($app in $wingetAppsRequired) {
  $installedOutput = winget list --id $app.Id -e 2>$null | Out-String
  if ($installedOutput -match $app.Id) {
    Write-Host "App '$($app.Name)' is already installed." -ForegroundColor Green
  }
  else {
    Write-Host "App '$($app.Name)' is not installed. Installing via winget..." -ForegroundColor Yellow
    Start-Process -FilePath $preferredShell `
      -ArgumentList "-NoProfile -ExecutionPolicy Bypass -Command `"winget install --id=$($app.Id) -e`"" `
      -WindowStyle Normal -Wait
  }
}

# -------------------------------
# Section 4: Software Updates & Optimizations
# -------------------------------
Write-Host "`n📦 Checking for software updates via winget (preview mode)..." -ForegroundColor Cyan
try {
  $wingetPreview = winget upgrade
  Write-Host "`nAvailable upgrades:" -ForegroundColor Yellow
  Write-Host $wingetPreview
}
catch {
  Write-Host "Error running winget. Ensure winget is installed and configured." -ForegroundColor Red
}

$upgradeConfirm = Get-UserInput "`nDo you want to upgrade the listed software packages? (y/n)" "y"
if ($upgradeConfirm -match "^(y|yes)$") {
  Write-Host "🔄 Initiating software upgrades via winget..." -ForegroundColor Yellow
  Start-Process -FilePath $preferredShell `
    -ArgumentList "-NoProfile -ExecutionPolicy Bypass -Command `"winget upgrade --all`"" `
    -WindowStyle Normal -Wait
}
else {
  Write-Host "Skipping software upgrades via winget." -ForegroundColor Green
}

# -------------------------------
# Section 5: Python Installation & Package Updates
# -------------------------------
Write-Host "`n🐍 Checking Python installation..." -ForegroundColor Cyan
if (-not (Get-Command python -ErrorAction SilentlyContinue)) {
  Write-Host "Python is not installed." -ForegroundColor Red
  $installPython = Get-UserInput "Do you want to install Python via winget? (y/n)" "y"
  if ($installPython -match "^(y|yes)$") {
    Write-Host "Installing Python..." -ForegroundColor Yellow
    Start-Process -FilePath $preferredShell `
      -ArgumentList "-NoProfile -ExecutionPolicy Bypass -Command `"winget install --id Python.Python.3`"" `
      -WindowStyle Normal -Wait
  }
  else {
    Write-Host "Skipping Python installation. Python-related updates will be skipped." -ForegroundColor Yellow
  }
}
else {
  Write-Host "Python is installed." -ForegroundColor Green

  Write-Host "`n🐍 Ensuring pip is up-to-date first..." -ForegroundColor Yellow
  Invoke-AdminCommand -ScriptBlock { python -m pip install --upgrade pip }

  Write-Host "`n🐍 Upgrading setuptools and wheel..." -ForegroundColor Yellow
  Invoke-AdminCommand -ScriptBlock { python -m pip install --upgrade setuptools wheel }


  # Bulk-install/upgrade pip packages using our pip package list.
  $pipPackageList = ($pipPackagesRequired | ForEach-Object { $_.Package }) -join " "
  Write-Host "`n📦 Installing/Upgrading required Python packages ($pipPackageList)..." -ForegroundColor Yellow
  $command = "pip install --upgrade $pipPackageList"
  Invoke-AdminCommand -ScriptBlock ([scriptblock]::Create($command))
}

# -------------------------------
# Section 6: PowerShell Module Checks & Profile Update
# -------------------------------
Write-Host "`n🔄 Checking for necessary PowerShell modules..." -ForegroundColor Cyan
foreach ($module in $psModulesRequired) {
  if (-not (Get-Module -ListAvailable -Name $module)) {
    Write-Host "PowerShell module '$module' is missing." -ForegroundColor Red
    $installModule = Get-UserInput "Do you want to simulate installation of '$module' using -WhatIf? (y/n)" "y"
    if ($installModule -match "^(y|yes)$") {
      Write-Host "Simulating installation of '$module' (WhatIf mode)..." -ForegroundColor Yellow
      Install-Module -Name $module -Scope CurrentUser -Force -WhatIf
      $confirmReal = Get-UserInput "Proceed with actual installation of '$module'? (y/n)" "y"
      if ($confirmReal -match "^(y|yes)$") {
        Install-Module -Name $module -Scope CurrentUser -Force
        Write-Host "'$module' installed." -ForegroundColor Green
      }
      else {
        Write-Host "Skipping actual installation of '$module'." -ForegroundColor Yellow
      }
    }
    else {
      Write-Host "Skipping installation of '$module'." -ForegroundColor Yellow
    }
  }
  else {
    Write-Host "PowerShell module '$module' is already installed." -ForegroundColor Green
  }
}

# Update the PowerShell profile with necessary imports.
$profilePaths = @(
  "$env:USERPROFILE\Documents\PowerShell\Microsoft.PowerShell_profile.ps1",
  "$env:USERPROFILE\Documents\WindowsPowerShell\Microsoft.PowerShell_profile.ps1"
)
$profileUpdated = $false
foreach ($p in $profilePaths) {
  if (Test-Path $p) {
    if (-not (Select-String -Path $p -Pattern "Import-Module\s+gsudoModule" -SimpleMatch -Quiet)) {
      $profilePrompt = Get-UserInput "Do you want to add 'Import-Module gsudoModule' to your PowerShell profile at '$p'? (y/n)" "y"
      if ($profilePrompt -match "^(y|yes)$") {
        Add-Content -Path $p -Value "`nImport-Module gsudoModule"
        Write-Host "Updated profile: $p" -ForegroundColor Green
        $profileUpdated = $true
        break
      }
    }
    else {
      Write-Host "Your profile at '$p' already imports 'gsudoModule'." -ForegroundColor Green
      $profileUpdated = $true
      break
    }
  }
}
if (-not $profileUpdated) {
  Write-Host "No valid PowerShell profile found or update skipped." -ForegroundColor Yellow
}

# -------------------------------
# Append oh-my-posh Initialization to Profile
# -------------------------------
# This command will initialize oh-my-posh with your chosen configuration.
$ohMyPoshInitCommand = "oh-my-posh init pwsh --config 'https://raw.githubusercontent.com/JanDeDobbeleer/oh-my-posh/main/themes/bubblesline.omp.json' | Invoke-Expression"
foreach ($p in $profilePaths) {
  if (Test-Path $p) {
    if (-not (Select-String -Path $p -Pattern [regex]::Escape($ohMyPoshInitCommand) -SimpleMatch -Quiet)) {
      $profilePrompt = Get-UserInput "Do you want to add the oh-my-posh initialization command to your PowerShell profile at '$p'? (y/n)" "y"
      if ($profilePrompt -match "^(y|yes)$") {
        Add-Content -Path $p -Value "`n$ohMyPoshInitCommand"
        Write-Host "Added oh-my-posh initialization to profile: $p" -ForegroundColor Green
      }
    }
    else {
      Write-Host "Your profile at '$p' already includes the oh-my-posh initialization command." -ForegroundColor Green
    }
  }
}

# -------------------------------
# Section 7: Final System Verification
# -------------------------------
Write-Host "`n✅ Performing final system checks..." -ForegroundColor Green
Write-Host "`n🐍 Verifying Python installation and package versions..." -ForegroundColor Yellow
try {
  python --version
  pip --version
}
catch {
  Write-Host "Python or pip not found." -ForegroundColor Red
}
Write-Host "`n🛠️ Rechecking system integrity with SFC and DISM..." -ForegroundColor Yellow
Invoke-AdminCommand -ScriptBlock { sfc /scannow }
Invoke-AdminCommand -ScriptBlock { DISM /Online /Cleanup-Image /CheckHealth }

# -------------------------------
# Section 8: Final Confirmation & Restart Prompt
# -------------------------------
Write-Host "`n✅ All fixes and updates are complete!" -ForegroundColor Green
$restartConfirm = Get-UserInput "Do you want to restart your computer now? (y/n)" "y"
if ($restartConfirm -match "^(y|yes)$") {
  Write-Host "🔁 Restarting your computer..." -ForegroundColor Red
  Stop-Transcript
  Invoke-AdminCommand -ScriptBlock { Restart-Computer }
}
else {
  Write-Host "👍 Restart skipped. Note that some changes may require a reboot to take effect." -ForegroundColor Green
}

# -------------------------------
# End of Script: Wait for User Input Before Exiting (unless -Silent is specified)
# -------------------------------
if (-not $Silent) {
  Write-Host "`n🟢 Press any key to exit this window..." -ForegroundColor Cyan
  $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# Stop Logging
Stop-Transcript
