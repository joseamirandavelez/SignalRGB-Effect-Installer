#Requires -Version 5.1

# --- Versioning ---
$ScriptVersion = "1.6"
$RepoOwner = "joseamirandavelez"
$RepoName = "SignalRGB-Effect-Installer"

# --- Load Windows Forms Assembly ---
# Load these first to ensure all UI elements (even error popups) get modern styling
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic

# --- ADD THIS BLOCK FOR DARK TITLE BAR ---
try {
    $dwmapiCode = @"
    using System;
    using System.Runtime.InteropServices;
    public class DwmApi {
        [DllImport("dwmapi.dll", PreserveSig = true)]
        public static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);
    }
"@
    Add-Type -TypeDefinition $dwmapiCode -Language CSharp
}
catch {
    # Ignore error if type is already added in current session
}

# Set high DPI awareness and visual styles
try {
    [System.Windows.Forms.Application]::SetHighDpiMode('SystemAware')
}
catch {}
[System.Windows.Forms.Application]::EnableVisualStyles()


<#
.SYNOPSIS
    A GUI-based installer for SignalRGB effects and components.
    Auto-detects file type (.html, .json, .zip) and installs to the correct folder.
    Effects are installed to .../Effects/[EffectName]/[EffectName].html
    Components are installed to .../Components/[ComponentName].json
    Includes a unified uninstaller.
#>

# --- Script-Wide Variables ---
$Global:ScriptDirectory = $null
$Global:ScriptFullPath = $null
$Global:DesktopShortcutPath = $null
$Global:StartMenuShortcutPath = $null
$Global:SendToShortcutPath = $null
$script:hue = 0 # NEW: Global variable for RGB animation

try {
    # Check if we are running as an EXE or a script
    if ($PSModuleInfo -ne $null -or $MyInvocation.MyCommand.Path -ne $null) {
        $Global:ScriptFullPath = $MyInvocation.MyCommand.Path
    }
    else {
        # Fallback for EXE wrappers or unconventional execution
        $Global:ScriptFullPath = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
    }

    # Ensure we actually got a path before trying to split it
    if ($null -eq $Global:ScriptFullPath) {
        throw "Script path could not be detected via standard methods."
    }

    $Global:ScriptDirectory = Split-Path -Path $Global:ScriptFullPath -Parent
    
    # Define shortcut paths
    $desktopPath = [Environment]::GetFolderPath("Desktop")
    $startMenuPath = [Environment]::GetFolderPath("Programs")
    $sendToPath = Join-Path -Path ([Environment]::GetFolderPath("ApplicationData")) -ChildPath "Microsoft\Windows\SendTo"

    $Global:DesktopShortcutPath = Join-Path -Path $desktopPath -ChildPath "Effect Installer.lnk"
    $Global:StartMenuShortcutPath = Join-Path -Path $startMenuPath -ChildPath "SignalRGB Tools\Effect Installer.lnk"
    $Global:SendToShortcutPath = Join-Path -Path $sendToPath -ChildPath "SignalRGB Installer.lnk"
}
catch {
    [System.Windows.Forms.MessageBox]::Show("CRITICAL ERROR: Could not determine script's own path. Shortcuts will fail. `n$($_.Exception.Message)", "Error", "OK", "Error")
    return
}

# --- Registry and App Configuration ---
$RegKey = "HKCU:\Software\WhirlwindFX\SignalRgb"
$RegValue = "UserDirectory"
$AppName = "SignalRGB"
$EffectsSubFolder = "Effects" # The subfolder inside UserDirectory
$ComponentsSubFolder = "Components" # The subfolder for components

# --- Load Remaining Assemblies ---
Add-Type -AssemblyName System.IO.Compression.FileSystem

# --- Helper Functions (Ensuring all are defined before first call) ---

# --- NEW: RGB COLOR FUNCTION ---
function Get-RGBColor {
    param ([double]$Hue)
    # Saturation 0.6, Lightness 0.15 (Deep/Dark color cycle)
    $S = 0.6; $L = 0.15; $C = (1 - [Math]::Abs(2 * $L - 1)) * $S
    $X = $C * (1 - [Math]::Abs(($Hue / 60) % 2 - 1)); $m = $L - $C / 2
    
    if ($Hue -lt 60) { $R = $C; $G = $X; $B = 0 }
    elseif ($Hue -lt 120) { $R = $X; $G = $C; $B = 0 }
    elseif ($Hue -lt 180) { $R = 0; $G = $C; $B = $X }
    elseif ($Hue -lt 240) { $R = 0; $G = $X; $B = $C }
    elseif ($Hue -lt 300) { $R = $X; $G = 0; $B = $C }
    else { $R = $C; $G = 0; $B = $X }
    
    return [System.Drawing.Color]::FromArgb(255, [byte](($R + $m) * 255), [byte](($G + $m) * 255), [byte](($B + $m) * 255))
}
# --- END NEW RGB FUNCTION ---

function Set-WindowDarkMode {
    param([System.Windows.Forms.Form]$Form)
    try {
        $hwnd = $Form.Handle
        [int]$attribute = 20 # DWMWA_USE_IMMERSIVE_DARK_MODE
        [int]$value = 1      # True
        # Call the API defined at the top of the script
        [DwmApi]::DwmSetWindowAttribute($hwnd, $attribute, [ref]$value, 4)
    }
    catch {
        # Fails silently on Windows 7/8
    }
}

function Check-ForUpdates {
    Write-Status "Checking for updates..."
    
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

        # 1. Get Latest Release Info from GitHub
        $apiUrl = "https://api.github.com/repos/$RepoOwner/$RepoName/releases/latest"
        $latestRelease = Invoke-RestMethod -Uri $apiUrl -ErrorAction Stop

        # 2. Compare Versions
        $latestTag = $latestRelease.tag_name -replace "^v", ""
        
        if ([System.Version]$latestTag -gt [System.Version]$ScriptVersion) {
            
            $msg = "A new version (v$latestTag) is available.`n`nCurrent: $ScriptVersion`nNew: $latestTag`n`nUpdate now? The app will restart."
            $result = [System.Windows.Forms.MessageBox]::Show($msg, "Update Available", "YesNo", "Information")

            if ($result -eq 'Yes') {
                Write-Status "Preparing update..."

                # --- Find the EXE Asset (Looking for .exe instead of .zip) ---
                $asset = $latestRelease.assets | Where-Object { $_.name -like "*.exe" } | Select-Object -First 1

                if (-not $asset) {
                    Write-Status "ERROR: No .exe asset found in release v$latestTag."
                    [System.Windows.Forms.MessageBox]::Show("Update failed: No executable file found in the latest release.", "Error", "OK", "Error") | Out-Null
                    return
                }

                $downloadUrl = $asset.browser_download_url
                
                # Get the path of the currently running EXE
                $currentExe = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
                $tempExe = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath "SignalRGB_Update.exe"

                # 3. Download the new EXE to Temp
                Write-Status "Downloading update..."
                try {
                    Invoke-WebRequest -Uri $downloadUrl -OutFile $tempExe -ErrorAction Stop
                }
                catch {
                    Write-Status "Download failed: $($_.Exception.Message)"
                    [System.Windows.Forms.MessageBox]::Show("Download failed.`n$($_.Exception.Message)", "Error", "OK", "Error") | Out-Null
                    return
                }

                # 4. Execute the File Swap
                # We spawn a hidden PowerShell process that:
                # - Waits for this app to close (Start-Sleep)
                # - Force-moves the temp EXE over the current one
                # - Restarts the new EXE
                Write-Status "Installing update..."
                
                $swapCommand = "Start-Sleep -s 2; Move-Item -Path '$tempExe' -Destination '$currentExe' -Force; Start-Process '$currentExe'"
                Start-Process powershell.exe -ArgumentList "-NoProfile -WindowStyle Hidden -Command `"$swapCommand`""

                # 5. Exit immediately to free up the file handle
                $Global:mainForm.Close()
                [System.Windows.Forms.Application]::Exit()
                Stop-Process -Id $PID -Force
            }
        }
        else {
            Write-Status "You are on the latest version ($ScriptVersion)."
        }
    }
    catch {
        Write-Status "Update check failed: $($_.Exception.Message)"
    }
}

function Write-Status {
    param (
        [string]$Message,
        [string]$Type = "INFO"
    )
    
    # Check if the global variable is a Forms control and has a created handle
    if ($script:txtStatus -is [System.Windows.Forms.TextBox] -and $script:txtStatus.IsHandleCreated) {
        $timestamp = Get-Date -Format "HH:mm:ss"
        $formattedMessage = "$timestamp - $Type - $Message`r`n"
        
        # --- MODIFICATION ---
        # Use BeginInvoke for thread-safe UI updates
        # Cast to [void] to prevent the IAsyncResult object from leaking to the pipeline
        [void]$script:txtStatus.BeginInvoke([Action[string]] {
                param ($msg)
                $script:txtStatus.AppendText($msg)
                $script:txtStatus.ScrollToCaret()
            }, $formattedMessage)
    }
    else {
        # Fallback for headless mode
        Write-Host "${Type}: $Message"
    }
}

function Set-OpenWithRegistryKeys {
    param (
        [string]$AppName,
        [string]$AppPath,
        [string[]]$FileExtensions
    )
    
    # 1. Register the application itself under App Paths
    $appExeName = "$($AppName).exe"
    $appKeyPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\App Paths\$appExeName"
    Write-Status "Registering app path for Open With..."
    
    try {
        # FIX: Ensure we create the App Paths key
        New-Item -Path $appKeyPath -Force -ErrorAction Stop | Out-Null
        
        # Point the default value to the PowerShell executable
        Set-ItemProperty -Path $appKeyPath -Name "(Default)" -Value "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" -Type String -Force

        # Set the 'Path' property to define the command line arguments
        Set-ItemProperty -Path $appKeyPath -Name "Path" -Value "powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$AppPath`"" -Type String -Force | Out-Null
        Write-Status "App Path created: $appExeName"
    }
    catch {
        Write-Status "ERROR setting App Path registry key: $($_.Exception.Message)"
        return $false
    }

    # 2. Link file extensions to the registered app
    Write-Status "Linking extensions ($($FileExtensions -join ', ')) to the app..."
    
    foreach ($ext in $FileExtensions) {
        $extKeyPath = "HKCU:\Software\Classes\$ext\OpenWithList"
        try {
            # Ensure OpenWithList key exists for the extension
            if (-not (Test-Path -Path $extKeyPath)) {
                New-Item -Path $extKeyPath -Force | Out-Null
            }
            
            # Add the app name to the OpenWithList
            Set-ItemProperty -Path $extKeyPath -Name "$appExeName" -Value "" -Type String -Force -ErrorAction Stop | Out-Null
            
            # CRITICAL FIX for Windows 11: Set the association in the subkey directly
            # This often forces the association to appear immediately.
            $assocKeyPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\ApplicationAssociationToasts"
            New-ItemProperty -Path $assocKeyPath -Name "$appExeName$ext" -Value "Application" -Type String -Force | Out-Null
            
            Write-Status "Added $appExeName to Open With for ${ext}"
        }
        catch {
            # FIX: Using ${ext} to resolve parser confusion near the colon.
            Write-Status "WARNING: Could not link Open With for ${ext}: $($_.Exception.Message)"
        }
    }
    return $true
}

function Show-CreateShortcutWindow {
    param (
        [string]$ScriptDirectory,
        [string]$IconPath,
        [bool]$CheckDesktop = $true,
        [bool]$CheckStartMenu = $true,
        [bool]$CheckSendTo = $true,
        [bool]$CheckOpenWith = $true
    )

    $shortcutForm = New-Object System.Windows.Forms.Form
    $shortcutForm.Text = "Create Shortcuts"
    $shortcutForm.Size = New-Object System.Drawing.Size(350, 260) # Set height for 4 options + button
    $shortcutForm.FormBorderStyle = 'FixedDialog'
    $shortcutForm.MaximizeBox = $false
    $shortcutForm.MinimizeBox = $false
    $shortcutForm.StartPosition = 'CenterParent'
    if ($Global:mainForm -and $Global:mainForm.Icon) { $shortcutForm.Icon = $Global:mainForm.Icon }

    # Dark Mode/RGB Fixes
    $shortcutForm.BackColor = $Global:mainForm.BackColor
    $shortcutForm.ForeColor = $Global:mainForm.ForeColor
    
    $layout = New-Object System.Windows.Forms.TableLayoutPanel
    $layout.Dock = 'Fill'
    $layout.Padding = 10
    $layout.ColumnCount = 1
    # Total of 6 rows: Label (0) + 4 Checkboxes (1-4) + Button (5)
    $layout.RowCount = 6 
    $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null # 0: Label
    $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null # 1: Desktop
    $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null # 2: Start Menu
    $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null # 3: Send To
    $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null # 4: Open With
    $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null # 5: Button/Spacer
    $shortcutForm.Controls.Add($layout)

    $lblInfo = New-Object System.Windows.Forms.Label
    $lblInfo.Text = "Create shortcuts and registry keys for the Effect Installer:"
    $lblInfo.Dock = 'Fill'
    # Dark Mode/RGB Fixes
    $lblInfo.BackColor = [System.Drawing.Color]::Transparent
    $lblInfo.ForeColor = [System.Drawing.Color]::White
    $layout.Controls.Add($lblInfo, 0, 0) # Row 0

    $chkDesktop = New-Object System.Windows.Forms.CheckBox
    $chkDesktop.Text = "On the Desktop"
    $chkDesktop.Checked = $CheckDesktop
    $chkDesktop.Dock = 'Fill'
    # Dark Mode/RGB Fixes
    $chkDesktop.BackColor = [System.Drawing.Color]::Transparent
    $chkDesktop.ForeColor = [System.Drawing.Color]::White
    $layout.Controls.Add($chkDesktop, 0, 1) # Row 1

    $chkStartMenu = New-Object System.Windows.Forms.CheckBox
    $chkStartMenu.Text = "In the Start Menu (under 'SignalRGB Tools')"
    $chkStartMenu.Checked = $CheckStartMenu
    $chkStartMenu.Dock = 'Fill'
    # Dark Mode/RGB Fixes
    $chkStartMenu.BackColor = [System.Drawing.Color]::Transparent
    $chkStartMenu.ForeColor = [System.Drawing.Color]::White
    $layout.Controls.Add($chkStartMenu, 0, 2) # Row 2

    $chkSendTo = New-Object System.Windows.Forms.CheckBox
    $chkSendTo.Text = "In the 'Send To' menu (for quick installs)"
    $chkSendTo.Checked = $CheckSendTo
    $chkSendTo.Dock = 'Fill'
    # Dark Mode/RGB Fixes
    $chkSendTo.BackColor = [System.Drawing.Color]::Transparent
    $chkSendTo.ForeColor = [System.Drawing.Color]::White
    $layout.Controls.Add($chkSendTo, 0, 3) # Row 3

    $chkOpenWith = New-Object System.Windows.Forms.CheckBox
    $chkOpenWith.Text = "Add to 'Open with' context menu"
    $chkOpenWith.Checked = $CheckOpenWith
    $chkOpenWith.Dock = 'Fill'
    # Dark Mode/RGB Fixes
    $chkOpenWith.BackColor = [System.Drawing.Color]::Transparent
    $chkOpenWith.ForeColor = [System.Drawing.Color]::White
    $layout.Controls.Add($chkOpenWith, 0, 4) # Row 4

    $btnCreate = New-Object System.Windows.Forms.Button
    $btnCreate.Text = "Create"
    $btnCreate.DialogResult = [System.Windows.Forms.DialogResult]::OK # Set DialogResult for clean closing
    $btnCreate.Dock = 'Fill'
    $btnCreate.Anchor = 'Top'
    $btnCreate.Height = 30
    $btnCreate.Margin = [System.Windows.Forms.Padding]::new(0, 10, 0, 0)
    # Dark Mode/RGB Fixes
    $btnCreate.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnCreate.ForeColor = [System.Drawing.Color]::White
    $btnCreate.BackColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
    $btnCreate.FlatAppearance.BorderSize = 1
    $btnCreate.FlatAppearance.BorderColor = [System.Drawing.Color]::Gray
    $layout.Controls.Add($btnCreate, 0, 5) # Row 5

    $btnCreate.Add_Click({
            # The execution of this script block performs the action and then closes the dialog
            try {
                # WScript.Shell is the object that creates shortcuts
                $wsShell = New-Object -ComObject WScript.Shell

                # Get the full path to the launch.vbs script
                $targetFile = Join-Path -Path $Global:ScriptDirectory -ChildPath "SignalRGB_Installer.exe"
                # The VBScript handles all arguments, so this should be empty
                $arguments = ""

                if ($chkDesktop.Checked) {
                    try {
                        Write-Status "Creating Desktop shortcut..."
                        $shortcut = $wsShell.CreateShortcut($Global:DesktopShortcutPath)
                        $shortcut.TargetPath = $targetFile
                        $shortcut.Arguments = $arguments
                        $shortcut.WorkingDirectory = $Global:ScriptDirectory
                        if (Test-Path -Path $IconPath) {
                            $shortcut.IconLocation = $IconPath
                        }
                        $shortcut.Save()
                        Write-Status "Desktop shortcut created."
                    }
                    catch {
                        Write-Status "ERROR creating Desktop shortcut: $($_.Exception.Message)"
                    }
                }
            
                if ($chkStartMenu.Checked) {
                    try {
                        Write-Status "Creating Start Menu shortcut..."
                        $startMenuFolder = Split-Path -Path $Global:StartMenuShortcutPath -Parent
                        if (-not (Test-Path -Path $startMenuFolder)) {
                            New-Item -Path $startMenuFolder -ItemType Directory -Force | Out-Null
                        }
                        $shortcut = $wsShell.CreateShortcut($Global:StartMenuShortcutPath)
                        $shortcut.TargetPath = $targetFile
                        $shortcut.Arguments = $arguments
                        $shortcut.WorkingDirectory = $Global:ScriptDirectory
                        if (Test-Path -Path $IconPath) {
                            $shortcut.IconLocation = $IconPath
                        }
                        $shortcut.Save()
                        Write-Status "Start Menu shortcut created."
                    }
                    catch {
                        Write-Status "ERROR creating Start Menu shortcut: $($_.Exception.Message)"
                    }
                }

                if ($chkSendTo.Checked) {
                    try {
                        Write-Status "Creating 'Send To' shortcut..."
                        $shortcut = $wsShell.CreateShortcut($Global:SendToShortcutPath)
                        $shortcut.TargetPath = $targetFile
                        $shortcut.Arguments = $arguments
                        $shortcut.WorkingDirectory = $Global:ScriptDirectory
                        if (Test-Path -Path $IconPath) {
                            $shortcut.IconLocation = $IconPath
                        }
                        $shortcut.Save()
                        Write-Status "'Send To' shortcut created: $($Global:SendToShortcutPath)"
                    }
                    catch {
                        Write-Status "ERROR creating 'Send To' shortcut: $($_.Exception.Message)"
                    }
                }
                
                if ($chkOpenWith.Checked) {
                    # MODIFIED: Add .json to the Open With list
                    Set-OpenWithRegistryKeys -AppName "SRGB Installer" -AppPath $Global:ScriptFullPath -FileExtensions @(".zip", ".html", ".json")
                }
            
                # Dialog will close cleanly because DialogResult is set on the button
            }
            catch {
                Write-Status "CRITICAL ERROR in shortcut creation click handler: $($_.Exception.Message)"
                [System.Windows.Forms.MessageBox]::Show("A critical error occurred while executing the creation script: $($_.Exception.Message)", "Fatal Error", "OK", "Error") | Out-Null
            }
        })

    # ShowDialog returns the DialogResult of the button that closed it.
    $shortcutForm.ShowDialog($Global:mainForm) | Out-Null
    $shortcutForm.Dispose()
}

function Show-UninstallWindow {
    param (
        [string]$UserDirectory # MODIFIED: Pass base user directory
    )

    # Define paths based on UserDirectory
    $effectsBasePath = Join-Path -Path $UserDirectory -ChildPath $EffectsSubFolder
    $componentsBasePath = Join-Path -Path $UserDirectory -ChildPath $ComponentsSubFolder

    $uninstallForm = New-Object System.Windows.Forms.Form
    $uninstallForm.Text = "Uninstall Effects / Components"
    $uninstallForm.Size = New-Object System.Drawing.Size(550, 400) # Increased width for 2 columns
    $uninstallForm.MinimumSize = New-Object System.Drawing.Size(400, 200)
    $uninstallForm.FormBorderStyle = 'Sizable'
    $uninstallForm.MaximizeBox = $true
    $uninstallForm.MinimizeBox = $true
    $uninstallForm.StartPosition = 'CenterParent'
    if ($Global:mainForm -and $Global:mainForm.Icon) { $uninstallForm.Icon = $Global:mainForm.Icon }

    # Dark Mode/RGB Fixes
    $uninstallForm.BackColor = $Global:mainForm.BackColor
    $uninstallForm.ForeColor = $Global:mainForm.ForeColor

    $layout = New-Object System.Windows.Forms.TableLayoutPanel
    $layout.Dock = 'Fill'
    $layout.Padding = 10
    $layout.ColumnCount = 1
    $layout.RowCount = 3
    $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
    $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
    $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
    $uninstallForm.Controls.Add($layout)

    # --- Header: Label and Refresh Button ---
    $headerPanel = New-Object System.Windows.Forms.TableLayoutPanel
    $headerPanel.Dock = 'Fill'
    $headerPanel.AutoSize = $true
    $headerPanel.ColumnCount = 2
    $headerPanel.RowCount = 1
    $headerPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
    $headerPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
    $layout.Controls.Add($headerPanel, 0, 0)
    
    $lblInfo = New-Object System.Windows.Forms.Label
    $lblInfo.Text = "Select items to delete (moves to Recycle Bin):" # MODIFIED
    $lblInfo.Dock = 'Fill'
    $lblInfo.TextAlign = 'MiddleLeft'
    # Dark Mode/RGB Fixes
    $lblInfo.BackColor = [System.Drawing.Color]::Transparent
    $lblInfo.ForeColor = [System.Drawing.Color]::White
    $headerPanel.Controls.Add($lblInfo, 0, 0)

    $btnRefresh = New-Object System.Windows.Forms.Button
    $btnRefresh.Text = "Refresh List"
    $btnRefresh.Dock = 'None'
    $btnRefresh.Anchor = 'Right'
    # Dark Mode/RGB Fixes
    $btnRefresh.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnRefresh.ForeColor = [System.Drawing.Color]::White
    $btnRefresh.BackColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
    $btnRefresh.FlatAppearance.BorderSize = 1
    $btnRefresh.FlatAppearance.BorderColor = [System.Drawing.Color]::Gray
    $headerPanel.Controls.Add($btnRefresh, 1, 0)

    # --- SCRIPT CHANGE: Create a 2-column layout for the lists ---
    $listsLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $listsLayout.Dock = 'Fill'
    $listsLayout.Margin = [System.Windows.Forms.Padding]::new(0, 5, 0, 5)
    $listsLayout.ColumnCount = 2
    $listsLayout.RowCount = 1
    $listsLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50))) | Out-Null
    $listsLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50))) | Out-Null
    $layout.Controls.Add($listsLayout, 0, 1)

    # --- Column 0: Effects List ---
    $effectsLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $effectsLayout.Dock = 'Fill'
    $effectsLayout.ColumnCount = 1
    $effectsLayout.RowCount = 2
    $effectsLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
    $effectsLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
    $listsLayout.Controls.Add($effectsLayout, 0, 0)

    $lblEffects = New-Object System.Windows.Forms.Label
    $lblEffects.Text = "Effects:"
    $lblEffects.Font = New-Object System.Drawing.Font($lblInfo.Font, [System.Drawing.FontStyle]::Bold)
    $lblEffects.Dock = 'Fill'
    # Dark Mode/RGB Fixes
    $lblEffects.BackColor = [System.Drawing.Color]::Transparent
    $lblEffects.ForeColor = [System.Drawing.Color]::White
    $effectsLayout.Controls.Add($lblEffects, 0, 0)

    $clbEffects = New-Object System.Windows.Forms.CheckedListBox
    $clbEffects.Dock = 'Fill'
    $clbEffects.CheckOnClick = $true
    $clbEffects.Margin = [System.Windows.Forms.Padding]::new(0, 3, 5, 0) # Add margin to the right
    # Dark Mode/RGB Fixes
    $clbEffects.BackColor = [System.Drawing.Color]::FromArgb(40, 40, 40)
    $clbEffects.ForeColor = [System.Drawing.Color]::White
    $effectsLayout.Controls.Add($clbEffects, 0, 1)

    # --- Column 1: Components List ---
    $componentsLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $componentsLayout.Dock = 'Fill'
    $componentsLayout.ColumnCount = 1
    $componentsLayout.RowCount = 2
    $componentsLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
    $componentsLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
    $listsLayout.Controls.Add($componentsLayout, 1, 0)

    $lblComponents = New-Object System.Windows.Forms.Label
    $lblComponents.Text = "Components:"
    $lblComponents.Font = New-Object System.Drawing.Font($lblInfo.Font, [System.Drawing.FontStyle]::Bold)
    $lblComponents.Dock = 'Fill'
    # Dark Mode/RGB Fixes
    $lblComponents.BackColor = [System.Drawing.Color]::Transparent
    $lblComponents.ForeColor = [System.Drawing.Color]::White
    $componentsLayout.Controls.Add($lblComponents, 0, 0)

    $clbComponents = New-Object System.Windows.Forms.CheckedListBox
    $clbComponents.Dock = 'Fill'
    $clbComponents.CheckOnClick = $true
    $clbComponents.Margin = [System.Windows.Forms.Padding]::new(5, 3, 0, 0) # Add margin to the left
    
    # --- MODIFICATION: Tell the listbox which property to display ---
    $clbComponents.DisplayMember = "DisplayName"
    # Dark Mode/RGB Fixes
    $clbComponents.BackColor = [System.Drawing.Color]::FromArgb(40, 40, 40)
    $clbComponents.ForeColor = [System.Drawing.Color]::White
    
    $componentsLayout.Controls.Add($clbComponents, 0, 1)
    # --- End of layout change ---

    # --- Footer: Delete Button ---
    $btnDelete = New-Object System.Windows.Forms.Button
    $btnDelete.Text = "Delete Selected"
    $btnDelete.Dock = 'Fill'
    $btnDelete.Height = 30
    # Dark Mode/RGB Fixes
    $btnDelete.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnDelete.ForeColor = [System.Drawing.Color]::White
    $btnDelete.BackColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
    $btnDelete.FlatAppearance.BorderSize = 1
    $btnDelete.FlatAppearance.BorderColor = [System.Drawing.Color]::Gray
    $layout.Controls.Add($btnDelete, 0, 2)

    # --- Functions for this window ---
    $populateList = {
        # --- SCRIPT CHANGE: Populate two separate lists ---
        $clbEffects.Items.Clear()
        $clbComponents.Items.Clear()
        $totalItems = 0

        Write-Status "Refreshing effect list from: $effectsBasePath"
        if (Test-Path -Path $effectsBasePath) {
            try {
                $effects = Get-ChildItem -Path $effectsBasePath -Directory | ForEach-Object { $_.Name } | Sort-Object
                $clbEffects.Items.AddRange($effects)
                $totalItems += $effects.Count
                Write-Status "Found $($effects.Count) effects."
            }
            catch {
                Write-Status "ERROR scanning for effects: $($_.Exception.Message)"
            }
        }
        else {
            Write-Status "Effects folder does not exist. Nothing to list."
        }

        Write-Status "Refreshing component list from: $componentsBasePath"
        if (Test-Path -Path $componentsBasePath) {
            try {
                # --- MODIFICATION: Get component names from JSON files ---
                $componentFiles = Get-ChildItem -Path $componentsBasePath -Filter "*.json"
                $componentList = @()
                
                foreach ($file in $componentFiles) {
                    $displayName = Get-ComponentNameFromJson -JsonFilePath $file.FullName
                    # Store an object with both the display name and the real filename
                    $componentList += [PSCustomObject]@{
                        DisplayName = $displayName
                        FileName    = $file.Name # e.g., "MyComponent.json"
                    }
                }
                
                # Sort the list of objects by their display name
                $sortedList = $componentList | Sort-Object -Property DisplayName
                
                $clbComponents.Items.AddRange($sortedList)
                $totalItems += $sortedList.Count
                Write-Status "Found $($sortedList.Count) components."
                # --- End of MODIFICATION ---
            }
            catch {
                Write-Status "ERROR scanning for components: $($_.Exception.Message)"
            }
        }
        else {
            Write-Status "Components folder does not exist. Nothing to list."
        }
        
        Write-Status "Found $totalItems total items."
    }

    $btnRefresh.Add_Click({
            $populateList.Invoke() # Use .Invoke() for script blocks
        })

    $btnDelete.Add_Click({
            # --- SCRIPT CHANGE: Get checked items from both lists ---
            $selectedItems = @()
            
            foreach ($item in $clbEffects.CheckedItems) {
                $selectedItems += "[Effect] $item"
            }
            
            # --- MODIFICATION: Get the FileName property from the selected object ---
            foreach ($item in $clbComponents.CheckedItems) {
                # $item is now a PSCustomObject
                $selectedItems += "[Component] $($item.FileName)"
            }
            # --- End of MODIFICATION ---
            
            if ($selectedItems.Count -eq 0) {
                [System.Windows.Forms.MessageBox]::Show("Please select one or more items to delete.", "No Selection", "OK", "Information") | Out-Null
                return
            }

            # --- Logic to determine if the active effect is being deleted ---
            # --- SCRIPT CHANGE: Get effect folders from the $clbEffects list ---
            $originalEffectFolders = @($clbEffects.Items)
            $currentAlwaysTitle = ""
            $currentAlwaysFolder = $null
        
            try {
                $currentAlwaysTitle = (Get-ItemProperty -Path "HKCU:\Software\WhirlwindFX\SignalRgb\effects\selected" -Name "always" -ErrorAction SilentlyContinue).always
            }
            catch {
                Write-Status "Could not read current 'always' key. Will not update registry on delete."
            }

            if (-not [string]::IsNullOrWhiteSpace($currentAlwaysTitle)) {
                # Find the folder name that corresponds to the active title
                foreach ($folderName in $originalEffectFolders) {
                    $effectHtmlPath = Join-Path -Path $effectsBasePath -ChildPath "$folderName\$folderName.html"
                    $title = Get-EffectTitleFromHtml -HtmlFilePath $effectHtmlPath
                    
                    if (-not [string]::IsNullOrWhiteSpace($title) -and $currentAlwaysTitle.Equals($title, [StringComparison]::OrdinalIgnoreCase)) {
                        $currentAlwaysFolder = $folderName
                        break
                    }
                }
            }
        
            $activeEffectFolderWasDeleted = $false
            $selectedEffectItems = $selectedItems | Where-Object { $_ -like '[Effect]*' }
            if ($currentAlwaysFolder -and $selectedEffectItems -contains "[Effect] $currentAlwaysFolder") {
                $activeEffectFolderWasDeleted = $true
                Write-Status "Active effect '$currentAlwaysTitle' is scheduled for deletion."
            }
            # --- End new logic block ---

            # --- Set new active effect *before* deleting files ---
            if ($activeEffectFolderWasDeleted) {
                Write-Status "Updating active effect registry keys..."
                $remainingEffectFolders = @($originalEffectFolders | Where-Object { $_ -notin ($selectedEffectItems | ForEach-Object { $_ -replace '^\[Effect\] ', '' }) })
            
                if ($remainingEffectFolders.Count -eq 0) {
                    # No effects left, set to empty
                    Write-Status "All effects deleted. Setting active effect to empty."
                    Set-ActiveEffectRegistryKeys -NewEffectTitle ""
                }
                else {
                    # Find the effect that was alphabetically before the one we deleted
                    $originalIndex = [array]::IndexOf($originalEffectFolders, $currentAlwaysFolder)
                    $newIndex = $originalIndex - 1
                
                    if ($newIndex -lt 0) {
                        # It was the first item, wrap around to the end of the *remaining* list
                        $newIndex = $remainingEffectFolders.Count - 1
                    }
                
                    # Check that the new index is valid for the remaining list
                    if ($newIndex -ge $remainingEffectFolders.Count) {
                        # This can happen if we delete the last item. Default to the new last item.
                        $newIndex = $remainingEffectFolders.Count - 1
                    }

                    $newEffectFolder = $remainingEffectFolders[$newIndex]
                    $newEffectHtmlPath = Join-Path -Path $effectsBasePath -ChildPath "$newEffectFolder\$newEffectFolder.html"
                    $newEffectTitle = Get-EffectTitleFromHtml -HtmlFilePath $newEffectHtmlPath
                
                    Write-Status "Setting new active effect to: '$newEffectTitle'"
                    Set-ActiveEffectRegistryKeys -NewEffectTitle $newEffectTitle
                }
            }
            # --- End of moved block ---

            # --- SCRIPT FIX: Use FileSystem.Delete... to send to Recycle Bin ---
            foreach ($fullItemName in $selectedItems) {
                
                $itemName = $null
                $itemType = "Unknown"

                try {
                    if ($fullItemName.StartsWith("[Effect] ")) {
                        $itemName = $fullItemName.Substring(9)
                        $itemType = "Effect"
                        $itemFolder = Join-Path -Path $effectsBasePath -ChildPath $itemName
                        
                        Write-Status "Moving ${itemType} to Recycle Bin: $itemName"
                        if (Test-Path -Path $itemFolder) {
                            [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteDirectory(
                                $itemFolder, 
                                [Microsoft.VisualBasic.FileIO.UIOption]::OnlyErrorDialogs, 
                                [Microsoft.VisualBasic.FileIO.RecycleOption]::SendToRecycleBin)
                            Write-Status "Successfully recycled $itemType '$itemName' (folder)."
                        }
                        else {
                            Write-Status "Folder '$itemFolder' not found for deletion. Skipping."
                        }
                    }
                    elseif ($fullItemName.StartsWith("[Component] ")) {
                        $itemName = $fullItemName.Substring(12) # $itemName is now "MyComponent.json"
                        $itemType = "Component"
                        
                        $jsonFilePath = Join-Path -Path $componentsBasePath -ChildPath $itemName
                        $pngFilePath = Join-Path -Path $componentsBasePath -ChildPath "$([System.IO.Path]::GetFileNameWithoutExtension($itemName)).png"
                        
                        Write-Status "Moving ${itemType} to Recycle Bin: $itemName"
                        
                        if (Test-Path -Path $jsonFilePath) {
                            [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile(
                                $jsonFilePath, 
                                [Microsoft.VisualBasic.FileIO.UIOption]::OnlyErrorDialogs, 
                                [Microsoft.VisualBasic.FileIO.RecycleOption]::SendToRecycleBin)
                            Write-Status "Successfully recycled file: $itemName"
                        }
                        else {
                            Write-Status "File '$jsonFilePath' not found for deletion. Skipping."
                        }
                        
                        # Conditionally delete matching .png
                        if (Test-Path -Path $pngFilePath) {
                            [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile(
                                $pngFilePath, 
                                [Microsoft.VisualBasic.FileIO.UIOption]::OnlyErrorDialogs, 
                                [Microsoft.VisualBasic.FileIO.RecycleOption]::SendToRecycleBin)
                            Write-Status "Successfully recycled matching PNG: $(Split-Path $pngFilePath -Leaf)"
                        }
                    }
                    else {
                        Write-Status "Skipping unknown item in delete list: $fullItemName"
                    }
                }
                catch {
                    Write-Status "ERROR deleting '$itemName': $($_.Exception.Message)"
                    [System.Windows.Forms.MessageBox]::Show("Error deleting '$itemName':`n$($_.Exception.Message)", "Error", "OK", "Error") | Out-Null
                }
            }
            # --- End of script fix ---
        
            Write-Status "Deletion complete."
        
            # Refresh the list
            $populateList.Invoke()
        })

    # --- Initial Load ---
    $populateList.Invoke()
    $uninstallForm.Add_Shown({
            Set-WindowDarkMode -Form $this
        })
    $uninstallForm.ShowDialog($Global:mainForm) | Out-Null
    $uninstallForm.Dispose()
}


function Get-EffectTitleFromHtml {
    param (
        [string]$HtmlFilePath
    )
    
    # Fallback title is the filename without extension
    $fallbackTitle = [System.IO.Path]::GetFileNameWithoutExtension($HtmlFilePath)
    
    try {
        $content = Get-Content -Path $HtmlFilePath -Raw
        $regex = [regex]'(?i)<title.*?>(.*?)</title>'
        $match = $regex.Match($content)
        
        if ($match.Success -and -not [string]::IsNullOrWhiteSpace($match.Groups[1].Value)) {
            return $match.Groups[1].Value.Trim()
        }
        else {
            Write-Status "WARNING: Could not find <title> tag in $HtmlFilePath. Using filename as fallback."
            return $fallbackTitle
        }
    }
    catch {
        Write-Status "ERROR: Could not read $HtmlFilePath to find title. Using filename as fallback. $($_.Exception.Message)"
        return $fallbackTitle
    }
}

function Get-ComponentNameFromJson {
    param (
        [string]$JsonFilePath
    )
    
    # Fallback name is the filename
    $fallbackName = [System.IO.Path]::GetFileName($JsonFilePath)
    
    try {
        # Remove non-breaking space characters (U+00A0) before parsing
        $jsonString = (Get-Content -Path $JsonFilePath -Raw) -replace "\u00A0", " "
        $content = $jsonString | ConvertFrom-Json
        
        if ($content.PSObject.Properties.Name -contains 'ProductName' -and -not [string]::IsNullOrWhiteSpace($content.ProductName)) {
            # Return the internal product name
            return $content.ProductName.Trim()
        }
        else {
            Write-Status "WARNING: No 'ProductName' in $JsonFilePath. Using filename as display name."
            return $fallbackName
        }
    }
    catch {
        Write-Status "ERROR: Could not parse $JsonFilePath to find component name. Using filename. $($_.Exception.Message)"
        return $fallbackName
    }
}

function Find-ComponentConflict {
    param (
        [string]$NewComponentJsonPath,
        [string]$ComponentsBasePath
    )
    
    # 1. Get the product name from the new file
    $newProductName = Get-ComponentNameFromJson -JsonFilePath $NewComponentJsonPath
    if ($newProductName -eq [System.IO.Path]::GetFileName($NewComponentJsonPath)) {
        # Get-ComponentNameFromJson returns filename on failure.
        # If it fails to parse the *new* file, we can't check for product name conflicts.
        Write-Status "Could not parse new component, skipping product name conflict check."
        return $null 
    }
    
    if (-not (Test-Path -Path $ComponentsBasePath)) {
        return $null # No components folder, so no conflict
    }
    
    # 2. Iterate through all existing component files
    $allJsonFiles = Get-ChildItem -Path $ComponentsBasePath -Recurse -Filter "*.json"
    foreach ($file in $allJsonFiles) {
        
        # Don't compare the file against itself (this handles drag-dropping a file from its install location)
        if ($file.FullName -eq $NewComponentJsonPath) {
            continue
        }

        $existingProductName = Get-ComponentNameFromJson -JsonFilePath $file.FullName
        
        if ($existingProductName -and ($newProductName.Equals($existingProductName, [StringComparison]::OrdinalIgnoreCase))) {
            # Found a match! Return the path of the *existing* file.
            Write-Status "Product name conflict found: '$newProductName' is used by $($file.Name)"
            return $file.FullName 
        }
    }
    
    return $null # No conflict found
}

function Find-EffectTitleConflict {
    param (
        [string]$NewEffectTitle,
        [string]$EffectsBasePath
    )
    
    if (-not (Test-Path -Path $EffectsBasePath)) {
        return $null # No effects folder, so no conflict
    }
    
    $allHtmlFiles = Get-ChildItem -Path $EffectsBasePath -Recurse -Filter "*.html"
    foreach ($file in $allHtmlFiles) {
        $existingTitle = Get-EffectTitleFromHtml -HtmlFilePath $file.FullName
        
        if ($existingTitle -and ($NewEffectTitle.Equals($existingTitle, [StringComparison]::OrdinalIgnoreCase))) {
            return $file.FullName # Return path of conflicting file
        }
    }
    
    return $null # No conflict found
}

function Show-ConflictDialog {
    param (
        [string]$Message
    )
    
    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = "Conflict Detected"
    $dialog.Size = New-Object System.Drawing.Size(400, 170) # Increased height
    $dialog.FormBorderStyle = 'FixedDialog'
    $dialog.MaximizeBox = $false
    $dialog.MinimizeBox = $false
    $dialog.StartPosition = 'CenterParent'
    if ($Global:mainForm -and $Global:mainForm.Icon) { $dialog.Icon = $Global:mainForm.Icon }
    
    # Dark Mode/RGB Fixes
    $dialog.BackColor = $Global:mainForm.BackColor
    $dialog.ForeColor = $Global:mainForm.ForeColor

    $mainLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $mainLayout.Dock = 'Fill'
    $mainLayout.Padding = 10
    $mainLayout.ColumnCount = 1
    $mainLayout.RowCount = 2
    $mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
    $mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
    $dialog.Controls.Add($mainLayout)

    $lblMessage = New-Object System.Windows.Forms.Label
    $lblMessage.Text = $Message
    $lblMessage.Dock = 'Fill'
    
    # --- MODIFICATION ---
    $lblMessage.TextAlign = 'MiddleLeft' # Was 'MiddleCenter'
    $lblMessage.BackColor = [System.Drawing.Color]::Transparent
    $lblMessage.ForeColor = [System.Drawing.Color]::White
    
    $mainLayout.Controls.Add($lblMessage, 0, 0)

    $buttonLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $buttonLayout.Dock = 'Fill'
    $buttonLayout.AutoSize = $true
    $buttonLayout.ColumnCount = 3
    $buttonLayout.RowCount = 1
    $buttonLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33.3))) | Out-Null
    $buttonLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33.3))) | Out-Null
    $buttonLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33.3))) | Out-Null
    $mainLayout.Controls.Add($buttonLayout, 0, 1)

    $btnOverwrite = New-Object System.Windows.Forms.Button
    $btnOverwrite.Text = "Overwrite"
    $btnOverwrite.DialogResult = 'OK'
    $btnOverwrite.Dock = 'None'
    $btnOverwrite.Anchor = 'Top, Left, Right'
    $btnOverwrite.Height = 30
    # Dark Mode/RGB Fixes
    $btnOverwrite.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnOverwrite.ForeColor = [System.Drawing.Color]::White
    $btnOverwrite.BackColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
    $btnOverwrite.FlatAppearance.BorderSize = 1
    $btnOverwrite.FlatAppearance.BorderColor = [System.Drawing.Color]::Gray
    $buttonLayout.Controls.Add($btnOverwrite, 0, 0)
    
    $btnRename = New-Object System.Windows.Forms.Button
    $btnRename.Text = "Rename"
    $btnRename.DialogResult = 'Retry'
    $btnRename.Dock = 'None'
    $btnRename.Anchor = 'Top, Left, Right'
    $btnRename.Height = 30
    # Dark Mode/RGB Fixes
    $btnRename.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnRename.ForeColor = [System.Drawing.Color]::White
    $btnRename.BackColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
    $btnRename.FlatAppearance.BorderSize = 1
    $btnRename.FlatAppearance.BorderColor = [System.Drawing.Color]::Gray
    $buttonLayout.Controls.Add($btnRename, 1, 0)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.DialogResult = 'Cancel'
    $btnCancel.Dock = 'None'
    $btnCancel.Anchor = 'Top, Left, Right'
    $btnCancel.Height = 30
    # Dark Mode/RGB Fixes
    $btnCancel.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnCancel.ForeColor = [System.Drawing.Color]::White
    $btnCancel.BackColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
    $btnCancel.FlatAppearance.BorderSize = 1
    $btnCancel.FlatAppearance.BorderColor = [System.Drawing.Color]::Gray
    $buttonLayout.Controls.Add($btnCancel, 2, 0)

    $result = $dialog.ShowDialog($Global:mainForm)
    $dialog.Dispose()
    
    # Map DialogResult to simple strings
    if ($result -eq 'OK') { return 'Overwrite' }
    if ($result -eq 'Retry') { return 'Rename' }
    return 'Cancel'
}

function Show-RenameDialog {
    param (
        [string]$OldName
    )
    
    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = "Rename Item"
    $dialog.Size = New-Object System.Drawing.Size(350, 150)
    $dialog.FormBorderStyle = 'FixedDialog'
    $dialog.MaximizeBox = $false
    $dialog.MinimizeBox = $false
    $dialog.StartPosition = 'CenterParent'
    if ($Global:mainForm -and $Global:mainForm.Icon) { $dialog.Icon = $Global:mainForm.Icon }

    # Dark Mode/RGB Fixes
    $dialog.BackColor = $Global:mainForm.BackColor
    $dialog.ForeColor = $Global:mainForm.ForeColor

    $lblInfo = New-Object System.Windows.Forms.Label
    $lblInfo.Text = "Enter a new name for the item:"
    $lblInfo.Location = New-Object System.Drawing.Point(20, 20)
    $lblInfo.AutoSize = $true
    # Dark Mode/RGB Fixes
    $lblInfo.BackColor = [System.Drawing.Color]::Transparent
    $lblInfo.ForeColor = [System.Drawing.Color]::White
    $dialog.Controls.Add($lblInfo)
    
    $txtNewName = New-Object System.Windows.Forms.TextBox
    $txtNewName.Text = $OldName
    $txtNewName.Location = New-Object System.Drawing.Point(20, 50)
    $txtNewName.Size = New-Object System.Drawing.Size(300, 20)
    # --- DARK MODE FIX ---
    $txtNewName.BackColor = [System.Drawing.Color]::FromArgb(40, 40, 40)
    $txtNewName.ForeColor = [System.Drawing.Color]::White
    $txtNewName.BorderStyle = 'FixedSingle'
    # ---------------------
    $dialog.Controls.Add($txtNewName)
        
    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "OK"
    $btnOK.DialogResult = 'OK'
    $btnOK.Location = New-Object System.Drawing.Point(160, 90)
    # Dark Mode/RGB Fixes
    $btnOK.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnOK.ForeColor = [System.Drawing.Color]::White
    $btnOK.BackColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
    $btnOK.FlatAppearance.BorderSize = 1
    $btnOK.FlatAppearance.BorderColor = [System.Drawing.Color]::Gray
    $dialog.Controls.Add($btnOK)
    
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.DialogResult = 'Cancel'
    $btnCancel.Location = New-Object System.Drawing.Point(240, 90)
    # Dark Mode/RGB Fixes
    $btnCancel.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnCancel.ForeColor = [System.Drawing.Color]::White
    $btnCancel.BackColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
    $btnCancel.FlatAppearance.BorderSize = 1
    $btnCancel.FlatAppearance.BorderColor = [System.Drawing.Color]::Gray
    $dialog.Controls.Add($btnCancel)
    
    $dialog.AcceptButton = $btnOK
    $dialog.CancelButton = $btnCancel

    if ($dialog.ShowDialog($Global:mainForm) -eq 'OK') {
        $newName = $txtNewName.Text.Trim()
        # Sanitize name to be a valid file/folder name
        $invalidChars = [IO.Path]::GetInvalidFileNameChars() -join ''
        $regex = "[$([regex]::Escape($invalidChars))]"
        $sanitizedName = $newName -replace $regex, ''
        
        $dialog.Dispose()
        
        if ([string]::IsNullOrWhiteSpace($sanitizedName)) {
            Write-Status "ERROR: New name is invalid or empty."
            return $null
        }
        return $sanitizedName
    }
    
    $dialog.Dispose()
    return $null
}

function Set-EffectTitleInHtml {
    param (
        [string]$HtmlFilePath,
        [string]$NewTitle
    )
    
    try {
        $content = Get-Content -Path $HtmlFilePath -Raw
        $regex = [regex]'(?i)(<title.*?>)(.*?)(</title>)'
        
        if ($regex.IsMatch($content)) {
            $newContent = $regex.Replace($content, ('$1' + $NewTitle + '$3'), 1)
            Set-Content -Path $HtmlFilePath -Value $newContent -Encoding UTF8
            Write-Status "Successfully updated <title> in $HtmlFilePath to '$NewTitle'."
            return $true
        }
        else {
            Write-Status "WARNING: Could not find <title> tag in $HtmlFilePath to update it."
            return $false
        }
    }
    catch {
        Write-Status "ERROR: Could not read or write to $HtmlFilePath to update title. $($_.Exception.Message)"
        return $false
    }
}

# --- NEW: Function to validate a component JSON file ---
function Test-IsComponentJson {
    param (
        [string]$JsonFilePath
    )
    
    try {
        # Remove non-breaking space characters (U+00A0) before parsing
        $jsonString = (Get-Content -Path $JsonFilePath -Raw) -replace "\u00A0", " "
        $content = $jsonString | ConvertFrom-Json
        
        # Check for key properties that identify a component
        if ($content.PSObject.Properties.Name -contains 'ProductName' -and
            $content.PSObject.Properties.Name -contains 'LedCount' -and
            $content.PSObject.Properties.Name -contains 'LedMapping') {
            
            Write-Status "JSON file confirmed as SignalRGB Component: $($content.ProductName)"
            return $true
        }
        else {
            Write-Status "ERROR: JSON file is missing required component properties (e.g., ProductName, LedCount)."
            return $false
        }
    }
    catch {
        Write-Status "ERROR: Could not parse JSON file. $($_.Exception.Message)"
        return $false
    }
}

# --- NEW: Function to update component JSON file on rename ---
function Set-ComponentNamesInJson {
    param (
        [string]$JsonFilePath,
        [string]$NewName
    )
    
    try {
        # Remove non-breaking space characters (U+00A0) before parsing
        $jsonString = (Get-Content -Path $JsonFilePath -Raw) -replace "\u00A0", " "
        $component = $jsonString | ConvertFrom-Json
        
        $component.ProductName = $NewName
        $component.DisplayName = $NewName
        
        # Convert back to JSON with sufficient depth for coordinate arrays
        $component | ConvertTo-Json -Depth 10 | Set-Content -Path $JsonFilePath -Encoding UTF8
        
        Write-Status "Successfully updated component JSON with new name: $NewName"
        return $true
    }
    catch {
        Write-Status "ERROR: Could not read or write to $JsonFilePath to update component names. $($_.Exception.Message)"
        return $false
    }
}

function Set-ActiveEffectRegistryKeys {
    param(
        [string]$NewEffectTitle
    )
    
    $keyPath = "HKCU:\Software\WhirlwindFX\SignalRgb\effects\selected"
    Write-Status "Setting active effect registry keys..."
    
    try {
        # Ensure the path exists
        if (-not (Test-Path -Path $keyPath)) {
            New-Item -Path $keyPath -Force | Out-Null
            Write-Status "Created registry key path: $keyPath"
        }
        
        # 1. Get the current 'always' value
        $currentAlways = Get-ItemProperty -Path $keyPath -Name "always" -ErrorAction SilentlyContinue
        
        if ($currentAlways) {
            # 2. Set 'previous' key to the 'always' value
            Set-ItemProperty -Path $keyPath -Name "previous" -Value $currentAlways.always
            Write-Status "Set 'previous' key to: $($currentAlways.always)"
        }
        else {
            Write-Status "No existing 'always' key found. Skipping 'previous' key set."
        }
        
        # 3. Set 'always' key to the new effect title
        Set-ItemProperty -Path $keyPath -Name "always" -Value $NewEffectTitle
        Write-Status "Set 'always' key to: $NewEffectTitle"
        
    }
    catch {
        Write-Status "ERROR: Could not update registry keys. $($_.Exception.Message)"
    }
}

# --- NEW: Disclaimer Window Function ---
function Show-DisclaimerWindow {
    
    # Define the disclaimer text
    $disclaimerText = @"
SRGB Effect Installer - Terms of Use

1. No Warranty
This Tool is provided "as-is", without any warranties of any kind, express or implied. The developer makes no guarantees regarding its functionality, reliability, or suitability for any particular purpose.

2. No Affiliation
The developer of this Tool is not associated with SignalRGB or its parent company, WhirlwindFX. This is an unofficial, third-party application.

3. Limitation of liability
In no event shall the developer be liable for any claim, damages, or other liability arising from the use of, or inability to use, the Tool.

4. User Responsibility
You are solely responsible for the effects you download and install. It is your responsibility to ensure they comply with the terms of service of any platform or application they are used with, including SignalRGB.

5. Content Disclaimer
The developer is not responsible for any content, including effects, installed by users of this Tool. Some user-generated effects may be inappropriate or unsuitable for certain audiences. Use caution when downloading or loading effects from others.

6. License (GNU GPLv3.0)
This project is licensed under the GNU General Public License v3.0. For the full license text, please see the project's GitHub page:
https://github.com/joseamirandavelez/SignalRGB-Effect-Installer
"@

    $disclaimerForm = New-Object System.Windows.Forms.Form
    $disclaimerForm.Text = "Disclaimer & Terms of Use"
    $disclaimerForm.Size = New-Object System.Drawing.Size(550, 450)
    $disclaimerForm.FormBorderStyle = 'FixedDialog'
    $disclaimerForm.MaximizeBox = $false
    $disclaimerForm.MinimizeBox = $false
    $disclaimerForm.StartPosition = 'CenterParent'
    if ($Global:mainForm -and $Global:mainForm.Icon) { $disclaimerForm.Icon = $Global:mainForm.Icon }

    # Dark Mode/RGB Fixes
    $disclaimerForm.BackColor = $Global:mainForm.BackColor
    $disclaimerForm.ForeColor = $Global:mainForm.ForeColor

    $layout = New-Object System.Windows.Forms.TableLayoutPanel
    $layout.Dock = 'Fill'
    $layout.Padding = 10
    $layout.ColumnCount = 1
    $layout.RowCount = 2
    $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
    $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
    $disclaimerForm.Controls.Add($layout)

    $rtbDisclaimer = New-Object System.Windows.Forms.RichTextBox
    $rtbDisclaimer.Dock = 'Fill'
    $rtbDisclaimer.ReadOnly = $true
    $rtbDisclaimer.DetectUrls = $true
    $rtbDisclaimer.Text = $disclaimerText
    $rtbDisclaimer.Margin = [System.Windows.Forms.Padding]::new(0, 0, 0, 10)
    # Dark Mode/RGB Fixes
    $rtbDisclaimer.BackColor = [System.Drawing.Color]::FromArgb(40, 40, 40)
    $rtbDisclaimer.ForeColor = [System.Drawing.Color]::White
    $rtbDisclaimer.BorderStyle = 'FixedSingle'
    $layout.Controls.Add($rtbDisclaimer, 0, 0)

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "OK"
    $btnOK.DialogResult = 'OK'
    $btnOK.Dock = 'None'
    $btnOK.Anchor = 'Right'
    $btnOK.Height = 30
    $btnOK.Width = 80
    # Dark Mode/RGB Fixes
    $btnOK.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnOK.ForeColor = [System.Drawing.Color]::White
    $btnOK.BackColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
    $btnOK.FlatAppearance.BorderSize = 1
    $btnOK.FlatAppearance.BorderColor = [System.Drawing.Color]::Gray
    $layout.Controls.Add($btnOK, 0, 1)

    $rtbDisclaimer.Add_LinkClicked({
            param($s, $e) # FIX: Renamed $sender to $s
            # Open the link in the default browser
            try {
                [System.Diagnostics.Process]::Start($e.LinkText)
            }
            catch {
                Write-Status "ERROR: Could not open URL: $($e.LinkText)"
            }
        })

    $disclaimerForm.AcceptButton = $btnOK
    $disclaimerForm.Add_Shown({
            Set-WindowDarkMode -Form $this
        })
    $disclaimerForm.ShowDialog($Global:mainForm) | Out-Null
    $disclaimerForm.Dispose()
}


# --- Main Installation Logic ---

function Start-Installation {
    param (
        [string]$FilePath
    )
    
    Write-Status "Starting installation for: $FilePath"
    
    # --- MODIFIED: restartRequired logic
    # Effects: Only require restart for New/Rename.
    # Components: ALWAYS require restart.
    $restartRequired = $false 
    $isOverwrite = $false
    
    # 1. Get SignalRGB User Directory from Registry
    $userDir = $null
    try {
        Write-Status "Reading registry key: $RegKey"
        $userDir = (Get-ItemProperty -Path $RegKey -Name $RegValue).$RegValue
        if ([string]::IsNullOrWhiteSpace($userDir) -or -not (Test-Path -Path $userDir)) {
            Write-Status "ERROR: Registry key found, but folder is invalid or missing: $userDir"
            [System.Windows.Forms.MessageBox]::Show("Error: SignalRGB UserDirectory not found or is invalid.`nChecked: $userDir", "Registry Error", "OK", "Error") | Out-Null
            return $false # Return false on failure
        }
        Write-Status "Found folder: $userDir"
    }
    catch {
        Write-Status "ERROR: Could not read registry key: $RegKey. $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Error: Could not read registry key for $AppName.`n$($_.Exception.Message)", "Registry Error", "OK", "Error") | Out-Null
        return $false # Return false on failure
    }

    # --- Define the correct base install folders ---
    $effectsBasePath = Join-Path -Path $userDir -ChildPath $EffectsSubFolder
    $componentsBasePath = Join-Path -Path $userDir -ChildPath $ComponentsSubFolder
    
    # --- Auto-detection variables ---
    $installType = $null # 'Effect' or 'Component'
    $installBasePath = $null
    $sourceHtmlFile = $null
    $sourceJsonFile = $null
    $sourcePngFile = $null
    $itemName = $null # The base name, e.g., "MyEffect" or "MyComponent"
    $tempExtractFolder = $null
    
    $extension = [System.IO.Path]::GetExtension($FilePath).ToLower()
    
    try {
        # 2. Prepare source files based on input (zip, html, or json)
        
        if ($extension -eq ".zip") {
            Write-Status "Zip file detected. Extracting..."
            $tempExtractFolder = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ([System.IO.Path]::GetRandomFileName())
            [System.IO.Compression.ZipFile]::ExtractToDirectory($FilePath, $tempExtractFolder)
            Write-Status "Extracted to: $tempExtractFolder"
            
            # Find the html or json files inside
            $sourceHtmlFile = (Get-ChildItem -Path $tempExtractFolder -Recurse -Filter "*.html" | Select-Object -First 1).FullName
            $sourceJsonFile = (Get-ChildItem -Path $tempExtractFolder -Recurse -Filter "*.json" | Select-Object -First 1).FullName

            if ($sourceHtmlFile) {
                # --- HTML file found, it's an Effect ---
                $installType = 'Effect'
                $itemName = [System.IO.Path]::GetFileNameWithoutExtension($sourceHtmlFile)
                $sourcePngFile = (Get-ChildItem -Path (Split-Path $sourceHtmlFile) -Filter "$itemName.png" | Select-Object -First 1).FullName
                Write-Status "Found Effect .html in zip: $itemName"
            }
            elseif ($sourceJsonFile) {
                # --- No HTML, but JSON found, it's a Component ---
                # Validate the JSON
                if (-not (Test-IsComponentJson -JsonFilePath $sourceJsonFile)) {
                    Write-Status "ERROR: .json file in zip is not a valid SignalRGB Component."
                    [System.Windows.Forms.MessageBox]::Show("Error: The .json file in the zip archive is not a valid SignalRGB Component.", "Zip Error", "OK", "Error") | Out-Null
                    return $false
                }
                
                $installType = 'Component'
                $itemName = [System.IO.Path]::GetFileNameWithoutExtension($sourceJsonFile)
                $sourcePngFile = (Get-ChildItem -Path (Split-Path $sourceJsonFile) -Filter "$itemName.png" | Select-Object -First 1).FullName
                Write-Status "Found Component .json in zip: $itemName"
            }
            else {
                Write-Status "ERROR: No .html or .json file found in the zip archive."
                [System.Windows.Forms.MessageBox]::Show("Error: No .html or .json file found in the zip archive.", "Zip Error", "OK", "Error") | Out-Null
                return $false # Return false on failure
            }
        }
        elseif ($extension -eq ".html") {
            # --- Standalone HTML, it's an Effect ---
            $installType = 'Effect'
            $sourceHtmlFile = $FilePath
            $itemName = [System.IO.Path]::GetFileNameWithoutExtension($sourceHtmlFile)
            $sourcePngFile = Join-Path -Path (Split-Path $FilePath) -ChildPath "$itemName.png"
            if (-not (Test-Path -Path $sourcePngFile)) { $sourcePngFile = $null }
            Write-Status "Detected standalone Effect file: $itemName"
        }
        elseif ($extension -eq ".json") {
            # --- Standalone JSON, check if it's a Component ---
            if (-not (Test-IsComponentJson -JsonFilePath $FilePath)) {
                Write-Status "ERROR: This .json file is not a valid SignalRGB Component."
                [System.Windows.Forms.MessageBox]::Show("Error: The selected .json file is not a valid SignalRGB Component.", "Invalid File", "OK", "Error") | Out-Null
                return $false
            }
            
            $installType = 'Component'
            $sourceJsonFile = $FilePath
            $itemName = [System.IO.Path]::GetFileNameWithoutExtension($sourceJsonFile)
            $sourcePngFile = Join-Path -Path (Split-Path $FilePath) -ChildPath "$itemName.png"
            if (-not (Test-Path -Path $sourcePngFile)) { $sourcePngFile = $null }
            Write-Status "Detected standalone Component file: $itemName"
        }
        else {
            Write-Status "ERROR: Invalid file type. Please select a .zip, .html, or .json file."
            [System.Windows.Forms.MessageBox]::Show("Invalid file type. Please select a .zip, .html, or .json file.", "Invalid File", "OK", "Error") | Out-Null
            return $false # Return false on failure
        }

        # 3. Conflict Detection Loop
        $installConfirmed = $false
        $currentItemName = $itemName
        
        # --- Set the correct path based on detected type ---
        if ($installType -eq 'Effect') {
            $installBasePath = $effectsBasePath
        }
        elseif ($installType -eq 'Component') {
            $installBasePath = $componentsBasePath
        }

        # --- Ensure the target base folder ('Effects' or 'Components') exists ---
        if (-not (Test-Path -Path $installBasePath)) {
            Write-Status "'$($installBasePath | Split-Path -Leaf)' folder not found. Creating it..."
            New-Item -Path $installBasePath -ItemType Directory -Force | Out-Null
            Write-Status "Created folder: $installBasePath"
        }

        
        while (-not $installConfirmed) {
            
            $titleConflictFile = $null
            $productNameConflictFile = $null
            $fileNameConflict = $false
            $conflictMessage = "A conflict was found.`n`n"

            # --- Type-specific conflict logic ---
            if ($installType -eq 'Effect') {
                # Effects check for folder names AND <title> tags
                $destFolder = Join-Path -Path $installBasePath -ChildPath $currentItemName
                $fileNameConflict = Test-Path -Path $destFolder # For effects, "fileNameConflict" means folder name conflict
                
                $currentEffectTitle = Get-EffectTitleFromHtml -HtmlFilePath $sourceHtmlFile
                Write-Status "Effect Title (from HTML): $currentEffectTitle"
                $titleConflictFile = Find-EffectTitleConflict -NewEffectTitle $currentEffectTitle -EffectsBasePath $installBasePath
                
                if ($fileNameConflict) { $conflictMessage += "- Folder '$currentItemName' already exists.`n" }
                
                # --- MODIFICATION ---
                if ($titleConflictFile) { $conflictMessage += "- Title '$currentEffectTitle' is already used by `n  $(Split-Path $titleConflictFile -Leaf)`n" }
            }
            elseif ($installType -eq 'Component') {
                # --- MODIFIED COMPONENT CONFLICT LOGIC ---
                $destJsonFile = Join-Path -Path $installBasePath -ChildPath ($currentItemName + ".json")
                $fileNameConflict = Test-Path -Path $destJsonFile # This is a file name conflict
                
                # Check for *internal* product name conflict
                $componentName = Get-ComponentNameFromJson -JsonFilePath $sourceJsonFile
                $productNameConflictFile = Find-ComponentConflict -NewComponentJsonPath $sourceJsonFile -ComponentsBasePath $installBasePath
                
                # If product name matches, but it's the *same file*, just show the file conflict.
                if ($productNameConflictFile -and $fileNameConflict -and ($productNameConflictFile -eq $destJsonFile)) {
                    $productNameConflictFile = $null
                }

                if ($fileNameConflict) { $conflictMessage += "- File '$($currentItemName + ".json")' already exists.`n" }
                if ($productNameConflictFile) { $conflictMessage += "- Component Name '$componentName' is already used by `n  $(Split-Path $productNameConflictFile -Leaf)`n" }
                # --- END MODIFIED LOGIC ---
            }
            
            if ($fileNameConflict -or $titleConflictFile -or $productNameConflictFile) {
                # MODIFIED to include product name
                # Conflict found!
                $conflictMessage += "`nWould you like to Overwrite, Rename, or Cancel?"
                
                Write-Status "Conflict detected: $conflictMessage"
                $userChoice = Show-ConflictDialog -Message $conflictMessage
                
                if ($userChoice -eq 'Overwrite') {
                    Write-Status "User chose to Overwrite."
                    
                    # --- NEW: Handle component product name conflict on Overwrite ---
                    # This deletes the *old* file that has the conflicting name.
                    if ($installType -eq 'Component' -and $productNameConflictFile) {
                        $oldJsonPath = $productNameConflictFile
                        $oldPngPath = Join-Path -Path (Split-Path $oldJsonPath) -ChildPath "$([System.IO.Path]::GetFileNameWithoutExtension($oldJsonPath)).png"

                        Write-Status "Overwrite: Deleting old conflicting component file: $(Split-Path $oldJsonPath -Leaf)"
                        try {
                            [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile(
                                $oldJsonPath, 
                                [Microsoft.VisualBasic.FileIO.UIOption]::OnlyErrorDialogs, 
                                [Microsoft.VisualBasic.FileIO.RecycleOption]::SendToRecycleBin)
                            
                            if (Test-Path -Path $oldPngPath) {
                                Write-Status "Overwrite: Deleting old matching PNG: $(Split-Path $oldPngPath -Leaf)"
                                [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile(
                                    $oldPngPath, 
                                    [Microsoft.VisualBasic.FileIO.UIOption]::OnlyErrorDialogs, 
                                    [Microsoft.VisualBasic.FileIO.RecycleOption]::SendToRecycleBin)
                            }
                            Write-Status "Successfully recycled old component files."
                        }
                        catch {
                            Write-Status "ERROR: Could not delete old conflicting component '($productNameConflictFile)'. $($_.Exception.Message)"
                            [System.Windows.Forms.MessageBox]::Show("Error: Could not delete the old conflicting file '$productNameConflictFile'.`nInstallation aborted.", "Error", "OK", "Error") | Out-Null
                            return $false # Abort installation
                        }
                    }
                    # --- END NEW LOGIC ---
                    
                    $isOverwrite = $true
                    $installConfirmed = $true
                }
                elseif ($userChoice -eq 'Rename') {
                    Write-Status "User chose to Rename."
                    $newName = Show-RenameDialog -OldName $currentItemName
                    
                    if ([string]::IsNullOrWhiteSpace($newName) -or $newName.Equals($currentItemName, [StringComparison]::OrdinalIgnoreCase)) {
                        Write-Status "Rename cancelled or name is unchanged."
                        # If name is unchanged, loop will repeat. If null, we cancel.
                        if ([string]::IsNullOrWhiteSpace($newName)) { return $false }
                        continue
                    }

                    Write-Status "New name selected: $newName"
                    $currentItemName = $newName
                    
                    # --- Type-specific rename logic ---
                    if ($installType -eq 'Effect') {
                        # Update HTML <title> tag
                        Set-EffectTitleInHtml -HtmlFilePath $sourceHtmlFile -NewTitle $newName | Out-Null
                    }
                    elseif ($installType -eq 'Component') {
                        # Update JSON 'ProductName' and 'DisplayName' fields
                        Set-ComponentNamesInJson -JsonFilePath $sourceJsonFile -NewName $newName | Out-Null
                    }
                }
                else {
                    # User chose Cancel
                    Write-Status "Installation cancelled by user."
                    return $false # Return false on user cancel
                }
            }
            else {
                # No conflict (New installation)
                Write-Status "No conflicts found. Proceeding with installation."
                $installConfirmed = $true
            }
        } # End conflict loop

        # 4. Perform Installation
        
        # --- Type-specific file copy ---
        if ($installType -eq 'Effect') {
            # --- SCRIPT FIX: Create folder *inside* effect block ---
            $destFolder = Join-Path -Path $installBasePath -ChildPath $currentItemName
            Write-Status "Installing '$currentItemName' to $destFolder"
            New-Item -Path $destFolder -ItemType Directory -Force | Out-Null

            $finalHtmlPath = Join-Path -Path $destFolder -ChildPath ($currentItemName + ".html")
            Copy-Item -Path $sourceHtmlFile -Destination $finalHtmlPath -Force
            Write-Status "Copied HTML to: $finalHtmlPath"
            
            # 5. Set Registry Keys (Only for Effects)
            $finalEffectTitle = Get-EffectTitleFromHtml -HtmlFilePath $finalHtmlPath
            Set-ActiveEffectRegistryKeys -NewEffectTitle $finalEffectTitle
            
            # 6. Determine restart necessity (Only for Effects)
            if (-not $isOverwrite) {
                $restartRequired = $true
                Write-Status "Installation finished. Restart is required for this effect."
            }
            else {
                Write-Status "Overwrite finished. No restart required."
            }
        }
        elseif ($installType -eq 'Component') {
            # --- SCRIPT FIX: Install component as flat file ---
            $finalJsonPath = Join-Path -Path $installBasePath -ChildPath ($currentItemName + ".json")
            Write-Status "Installing '$currentItemName' to $finalJsonPath"
            Copy-Item -Path $sourceJsonFile -Destination $finalJsonPath -Force
            Write-Status "Copied JSON to: $finalJsonPath"
            
            # 6. Determine restart necessity (Components)
            $restartRequired = $true # Components always require a restart
            Write-Status "Component installation finished. Restart is required."
        }
        
        # --- Copy PNG (for both types, with correct paths) ---
        $finalPngDestPath = $null
        if ($installType -eq 'Effect') {
            $destFolder = Join-Path -Path $installBasePath -ChildPath $currentItemName # Re-define for this scope
            $finalPngDestPath = Join-Path -Path $destFolder -ChildPath ($currentItemName + ".png")
        }
        elseif ($installType -eq 'Component') {
            $finalPngDestPath = Join-Path -Path $installBasePath -ChildPath ($currentItemName + ".png")
        }
        
        if ($finalPngDestPath -and $sourcePngFile -and (Test-Path -Path $sourcePngFile)) {
            Copy-Item -Path $sourcePngFile -Destination $finalPngDestPath -Force
            Write-Status "Copied PNG to: $finalPngDestPath"
        }
        elseif ($sourcePngFile -and -not (Test-Path -Path $sourcePngFile)) {
            Write-Status "No matching .png file found at $sourcePngFile. Skipping PNG copy."
        }
        
        [System.Windows.Forms.MessageBox]::Show("'$currentItemName' ($installType) installed/updated successfully.", "Installation Complete", "OK", "Information") | Out-Null
        
        return $restartRequired

    }
    catch {
        $errorMsg = $_.Exception.Message
        Write-Status "FATAL ERROR during installation: $errorMsg"
        if ($_.Exception.InnerException) {
            Write-Status "INNER EXCEPTION: $($_.Exception.InnerException.Message)"
        }
        [System.Windows.Forms.MessageBox]::Show("A fatal error occurred during installation of '$currentItemName':`n$errorMsg", "Fatal Error", "OK", "Error") | Out-Null
        return $false # Return false on fatal error
    }
    finally {
        # 7. Clean up temp folder if we created one
        if ($tempExtractFolder -and (Test-Path -Path $tempExtractFolder)) {
            Write-Status "Cleaning up temporary folder: $tempExtractFolder"
            Remove-Item -Path $tempExtractFolder -Recurse -Force
        }
    }
}

# --- Headless Drag-and-Drop/Send To Handler (Now after functions) ---
if ($args.Count -gt 0) {
    # We assume all arguments are files passed by the OS when using drag-and-drop or Send To.
    
    $batchRestartRequired = $false
    $AppName = "SignalRGB" # Needs to be defined in this scope

    Write-Host "Starting headless installation for $($args.Count) file(s)..."
    
    foreach ($droppedFile in $args) {
        $droppedFile = $droppedFile.Trim()
        
        if (Test-Path -Path $droppedFile -PathType Leaf) {
            $ext = [System.IO.Path]::GetExtension($droppedFile).ToLower()
            
            if ($ext -in @(".zip", ".html", ".json")) {
                Write-Host "Processing: $droppedFile"
                $restartNeeded = $false
                
                try {
                    $restartNeeded = Start-Installation -FilePath $droppedFile
                } 
                catch {
                    [System.Windows.Forms.MessageBox]::Show("A critical error occurred during installation of $($droppedFile): $($_.Exception.Message)", "Fatal Error", "OK", "Error") | Out-Null
                    continue
                }
                
                if ($restartNeeded) {
                    $batchRestartRequired = $true
                }
            }
            else {
                Write-Host "Skipping invalid file type: $droppedFile"
            }
        }
    }
    
    # Final Restart Prompt (if at least one installed file required it)
    if ($batchRestartRequired) {
        $restartResult = [System.Windows.Forms.MessageBox]::Show("Batch installation complete.`n`n$AppName must be restarted to load the new item(s). Restart now?", "Restart Required", "YesNo", "Question")
    
        if ($restartResult -eq 'Yes') {
            Write-Host "User chose to restart."
            try {
                Start-Process "signalrgb://app/restart" -ErrorAction Stop
                Write-Host "Restart signal sent to $AppName."
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Could not trigger the restart via URI. Please restart $AppName manually.", "Restart Failed", "OK", "Warning") | Out-Null
            }
        }
    }
    else {
        [System.Windows.Forms.MessageBox]::Show("Batch installation complete. No restart was required.", "Installation Complete", "OK", "Information") | Out-Null
    }

    # Exit the script after a successful headless install
    exit 0
}
# --- End Headless Handler ---


# --- GUI Definition ---

# High DPI and Visual Styles are already set at the top of the script.

# --- Main Form ---
$Global:mainForm = New-Object System.Windows.Forms.Form
$Global:mainForm.Text = "SRGB Effect & Component Installer - v$ScriptVersion"
$Global:mainForm.Size = New-Object System.Drawing.Size(650, 500) 
$Global:mainForm.FormBorderStyle = 'FixedSingle'
$Global:mainForm.MaximizeBox = $false
$Global:mainForm.StartPosition = 'CenterScreen'
$Global:mainForm.ShowInTaskbar = $true

# --- DARK MODE/RGB FIXES ---
$Global:mainForm.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30) # Default Dark Grey
$Global:mainForm.ForeColor = [System.Drawing.Color]::White # White Text
# --- END DARK MODE/RGB FIXES ---

# --- Main Layout Table ---
$mainLayout = New-Object System.Windows.Forms.TableLayoutPanel
$mainLayout.Dock = 'Fill'
$mainLayout.Padding = 10
$mainLayout.ColumnCount = 1
$mainLayout.RowCount = 4 
$mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null 
$mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null 
$mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null 
$mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null 

# FIX 1: Set Transparency so the RGB animation shows through
$mainLayout.BackColor = [System.Drawing.Color]::Transparent 

# FIX 2: Add the layout to the FORM, not to itself
$Global:mainForm.Controls.Add($mainLayout) 

# --- Row 0: Logo ---
$logoBase64 = "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"
$Global:picLogo = New-Object System.Windows.Forms.PictureBox
$Global:picLogo.SizeMode = 'Zoom'
$Global:picLogo.Size = New-Object System.Drawing.Size(128, 128)
$Global:picLogo.Dock = 'None'
$Global:picLogo.Anchor = 'Top'
$Global:picLogo.Margin = [System.Windows.Forms.Padding]::new(0, 0, 0, 10)
$Global:picLogo.BackColor = [System.Drawing.Color]::Transparent
$mainLayout.Controls.Add($Global:picLogo, 0, 0)

if ($logoBase64) {
    try {
        $imageBytes = [System.Convert]::FromBase64String($logoBase64)
        $ms = New-Object System.IO.MemoryStream($imageBytes, 0, $imageBytes.Length)
        $Global:picLogo.Image = [System.Drawing.Image]::FromStream($ms)
    }
    catch {
        Write-Status "Failed to load embedded logo string."
    }
}

# These paths are no longer needed for the executable version
$localIconPath = Join-Path -Path $Global:ScriptDirectory -ChildPath "icon.ico"
$localLogoPath = Join-Path -Path $Global:ScriptDirectory -ChildPath "logo.png"

# Load Logo
$logoLoaded = $false
if (Test-Path -Path $localLogoPath) {
    try {
        $logoStream = New-Object System.IO.FileStream($localLogoPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
        $Global:picLogo.Image = [System.Drawing.Image]::FromStream($logoStream)
        $logoStream.Close()
        $logoStream.Dispose()
        $mainLayout.Controls.Add($Global:picLogo, 0, 0)
        $logoLoaded = $true
    }
    catch {
        Write-Status "ERROR: Could not load logo.png: $($_.Exception.Message)"
    }
}

# --- Row 1: File Input ---
$fileInputLayout = New-Object System.Windows.Forms.TableLayoutPanel
$fileInputLayout.Dock = 'Fill'
$fileInputLayout.AutoSize = $true
$fileInputLayout.ColumnCount = 3
$fileInputLayout.RowCount = 4
$fileInputLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
$fileInputLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
$fileInputLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 95))) | Out-Null

# FIX 3: Transparency for the input area
$fileInputLayout.BackColor = [System.Drawing.Color]::Transparent
$mainLayout.Controls.Add($fileInputLayout, 0, 1)

# --- ENABLE DROP ZONE ---
$fileInputLayout.AllowDrop = $true

$lblFile = New-Object System.Windows.Forms.Label
$lblFile.Text = "Install File(s):" 
$lblFile.Dock = 'Fill'
$lblFile.TextAlign = 'MiddleLeft'
$lblFile.Margin = [System.Windows.Forms.Padding]::new(0, 0, 5, 0)
$lblFile.BackColor = [System.Drawing.Color]::Transparent
$lblFile.ForeColor = [System.Drawing.Color]::White
$fileInputLayout.Controls.Add($lblFile, 0, 0)

$txtFilePath = New-Object System.Windows.Forms.TextBox
$txtFilePath.Dock = 'Fill'
$txtFilePath.AllowDrop = $true
# --- DARK MODE FIX ---
$txtFilePath.BackColor = [System.Drawing.Color]::FromArgb(40, 40, 40)
$txtFilePath.ForeColor = [System.Drawing.Color]::White
$txtFilePath.BorderStyle = 'FixedSingle' # Makes it look flatter/modern
# ---------------------
$fileInputLayout.Controls.Add($txtFilePath, 1, 0)

$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = "Browse..."
$btnBrowse.Dock = 'None' 
$btnBrowse.AutoSize = $true 
$btnBrowse.Anchor = 'Top, Left'
$btnBrowse.Margin = [System.Windows.Forms.Padding]::new(5, 0, 0, 0)
$btnBrowse.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnBrowse.ForeColor = [System.Drawing.Color]::White
$btnBrowse.BackColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
$btnBrowse.FlatAppearance.BorderSize = 1
$btnBrowse.FlatAppearance.BorderColor = [System.Drawing.Color]::Gray
$fileInputLayout.Controls.Add($btnBrowse, 2, 0)

$lblHint = New-Object System.Windows.Forms.Label
$lblHint.Text = "Drag-and-drop a .zip, .html, or .json file here, or click Browse." 
$lblHint.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
$lblHint.Dock = 'Fill'
$lblHint.Margin = [System.Windows.Forms.Padding]::new(0, 0, 0, 5)
$lblHint.BackColor = [System.Drawing.Color]::Transparent
$lblHint.ForeColor = [System.Drawing.Color]::White
$fileInputLayout.Controls.Add($lblHint, 1, 1) 
$fileInputLayout.SetColumnSpan($lblHint, 2)

# --- Add LinkLabels (Builder) ---
$lnkBuilder = New-Object System.Windows.Forms.LinkLabel
$lnkBuilder.Text = "Create your own effects with the Effect Builder"
$lnkBuilder.Links.Add(32, 14, "https://effectbuilder.github.io/") | Out-Null 
$lnkBuilder.Dock = 'Fill'
$lnkBuilder.TextAlign = 'MiddleLeft'
$lnkBuilder.Margin = [System.Windows.Forms.Padding]::new(0, 0, 0, 5)
$lnkBuilder.BackColor = [System.Drawing.Color]::Transparent
# Fix Link Color for Dark Mode
$lnkBuilder.LinkColor = [System.Drawing.Color]::Cyan 
$fileInputLayout.Controls.Add($lnkBuilder, 1, 2)
$fileInputLayout.SetColumnSpan($lnkBuilder, 2)

$lnkBuilder.Add_LinkClicked({
        param($s, $e)
        try { [System.Diagnostics.Process]::Start($e.Link.LinkData); $e.Link.Visited = $true } catch {}
    })

$lnkComponentBuilder = New-Object System.Windows.Forms.LinkLabel
$lnkComponentBuilder.Text = "Create your own components with the Component Builder"
$lnkComponentBuilder.Links.Add(35, 17, "https://effectbuilder.github.io/builder/") | Out-Null 
$lnkComponentBuilder.Dock = 'Fill'
$lnkComponentBuilder.TextAlign = 'MiddleLeft'
$lnkComponentBuilder.Margin = [System.Windows.Forms.Padding]::new(0, 0, 0, 5)
$lnkComponentBuilder.BackColor = [System.Drawing.Color]::Transparent
# Fix Link Color for Dark Mode
$lnkComponentBuilder.LinkColor = [System.Drawing.Color]::Cyan
$fileInputLayout.Controls.Add($lnkComponentBuilder, 1, 3) 
$fileInputLayout.SetColumnSpan($lnkComponentBuilder, 2)

$lnkComponentBuilder.Add_LinkClicked({
        param($s, $e)
        try { [System.Diagnostics.Process]::Start($e.Link.LinkData); $e.Link.Visited = $true } catch {}
    })

# --- Row 2: Buttons ---
$buttonLayout = New-Object System.Windows.Forms.TableLayoutPanel
$buttonLayout.Dock = 'Fill'
$buttonLayout.AutoSize = $true
$buttonLayout.ColumnCount = 4 
$buttonLayout.RowCount = 1
$buttonLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
$buttonLayout.ColumnStyles.Clear()
$buttonLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 25))) | Out-Null
$buttonLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 25))) | Out-Null
$buttonLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 25))) | Out-Null
$buttonLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 25))) | Out-Null
$buttonLayout.Margin = [System.Windows.Forms.Padding]::new(0, 5, 0, 10)

# FIX 4: Transparency for the button container
$buttonLayout.BackColor = [System.Drawing.Color]::Transparent

$mainLayout.Controls.Add($buttonLayout, 0, 2)

$btnInstall = New-Object System.Windows.Forms.Button
$btnInstall.Text = "Install Item(s)" # MODIFIED
$btnInstall.Dock = 'None' 
$btnInstall.Anchor = 'Top, Left, Right' 
$btnInstall.Height = 30 
# Dark Mode/RGB Fixes
$btnInstall.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnInstall.ForeColor = [System.Drawing.Color]::White
$btnInstall.BackColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
$btnInstall.FlatAppearance.BorderSize = 1
$btnInstall.FlatAppearance.BorderColor = [System.Drawing.Color]::Gray
$buttonLayout.Controls.Add($btnInstall, 0, 0)

$btnUninstall = New-Object System.Windows.Forms.Button
$btnUninstall.Text = "Uninstall an Item..." # MODIFIED
$btnUninstall.Dock = 'None' 
$btnUninstall.Anchor = 'Top, Left, Right' 
$btnUninstall.Height = 30 
# Dark Mode/RGB Fixes
$btnUninstall.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnUninstall.ForeColor = [System.Drawing.Color]::White
$btnUninstall.BackColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
$btnUninstall.FlatAppearance.BorderSize = 1
$btnUninstall.FlatAppearance.BorderColor = [System.Drawing.Color]::Gray
$buttonLayout.Controls.Add($btnUninstall, 1, 0)

# --- Disclaimer Button ---
$btnDisclaimer = New-Object System.Windows.Forms.Button
$btnDisclaimer.Text = "Disclaimer..."
$btnDisclaimer.Dock = 'None'
$btnDisclaimer.Anchor = 'Top, Left, Right'
$btnDisclaimer.Height = 30
# Dark Mode/RGB Fixes
$btnDisclaimer.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnDisclaimer.ForeColor = [System.Drawing.Color]::White
$btnDisclaimer.BackColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
$btnDisclaimer.FlatAppearance.BorderSize = 1
$btnDisclaimer.FlatAppearance.BorderColor = [System.Drawing.Color]::Gray
$buttonLayout.Controls.Add($btnDisclaimer, 2, 0)

# --- NEW: Open Folder Button ---
$btnOpenFolder = New-Object System.Windows.Forms.Button
$btnOpenFolder.Text = "Open WhirlwindFX Folder"
$btnOpenFolder.Dock = 'None'
$btnOpenFolder.Anchor = 'Top, Left, Right'
$btnOpenFolder.Height = 30
# Dark Mode/RGB Fixes
$btnOpenFolder.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnOpenFolder.ForeColor = [System.Drawing.Color]::White
$btnOpenFolder.BackColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
$btnOpenFolder.FlatAppearance.BorderSize = 1
$btnOpenFolder.FlatAppearance.BorderColor = [System.Drawing.Color]::Gray
$buttonLayout.Controls.Add($btnOpenFolder, 3, 0)

# --- Row 3: Status Box ---
$script:txtStatus = New-Object System.Windows.Forms.TextBox
$script:txtStatus.Multiline = $true
$script:txtStatus.ReadOnly = $true
$script:txtStatus.ScrollBars = 'Vertical'
$script:txtStatus.Dock = 'Fill'
$script:txtStatus.Font = New-Object System.Drawing.Font("Consolas", 9)
# Dark Mode/RGB Fixes
# Set to pure black to contrast with the colorful animated background
$script:txtStatus.BackColor = [System.Drawing.Color]::Black 
# Brighter, crisper green
$script:txtStatus.ForeColor = [System.Drawing.Color]::FromArgb(50, 255, 50) 
$script:txtStatus.BorderStyle = 'Fixed3D' # Adds a bit of depth
$script:txtStatus.BorderStyle = 'FixedSingle'
$mainLayout.Controls.Add($script:txtStatus, 0, 3)

# --- Form and Control Event Handlers ---

$Global:mainForm.Add_Shown({
    # Enable Dark Mode
    Set-WindowDarkMode -Form $Global:mainForm

    # 1. Handle the Window Title Bar Icon (Extract from EXE)
    try {
        $currentProcess = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
        $extractedIcon = [System.Drawing.Icon]::ExtractAssociatedIcon($currentProcess)
        $Global:mainForm.Icon = $extractedIcon
        Write-Status "Window icon loaded from executable resources."
    }
    catch {
        Write-Status "Resource Error (Icon): $($_.Exception.Message)"
    }

    # 2. Handle the PictureBox Logo (Prefer Base64 -> then File -> then Icon Fallback)
    $logoSuccessfullySet = $false

    # Try loading from the Base64 string first (Most reliable for EXE)
    if (-not [string]::IsNullOrEmpty($logoBase64)) {
        try {
            $imageBytes = [System.Convert]::FromBase64String($logoBase64)
            $ms = New-Object System.IO.MemoryStream($imageBytes, 0, $imageBytes.Length)
            $Global:picLogo.Image = [System.Drawing.Image]::FromStream($ms)
            $logoSuccessfullySet = $true
            Write-Status "Logo loaded from embedded Base64 string."
        }
        catch {
            Write-Status "Base64 Logo Error: $($_.Exception.Message)"
        }
    }

    # If Base64 failed, try loading from local file
    if (-not $logoSuccessfullySet -and (Test-Path -Path $localLogoPath)) {
        try {
            $bytes = [System.IO.File]::ReadAllBytes($localLogoPath)
            $ms = New-Object System.IO.MemoryStream($bytes)
            $Global:picLogo.Image = [System.Drawing.Image]::FromStream($ms)
            $logoSuccessfullySet = $true
            Write-Status "Logo loaded from local file: logo.png"
        }
        catch {
            Write-Status "Local Logo Error: $($_.Exception.Message)"
        }
    }

    # Final Fallback: If still nothing, use the extracted Icon for the PictureBox
    if (-not $logoSuccessfullySet -and $null -ne $Global:mainForm.Icon) {
        $Global:picLogo.Image = $Global:mainForm.Icon.ToBitmap()
        Write-Status "Logo fell back to extracted executable icon."
    }

    # --- Check for missing items (Robust Logic) ---
    Write-Status "--- Checking for Installer Shortcuts and Registry Keys ---"

    # Check Shortcut Files
    $desktopExists = (Test-Path -Path $Global:DesktopShortcutPath -PathType Leaf)
    $startMenuExists = (Test-Path -Path $Global:StartMenuShortcutPath -PathType Leaf)
    $sendToExists = (Test-Path -Path $Global:SendToShortcutPath -PathType Leaf)

    # Check Registry Key (Open With)
    $openWithKeyPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\App Paths\SignalRGB Installer.exe"
    $openWithExists = Test-Path -Path $openWithKeyPath

    $isAnyMissing = (-not $desktopExists -or -not $startMenuExists -or -not $sendToExists -or -not $openWithExists)

    Write-Status "Desktop: $desktopExists, Start Menu: $startMenuExists, Send To: $sendToExists, Open With Reg: $openWithExists"

    if ($isAnyMissing) {
        Write-Status "One or more desired setup items are missing. Prompting user."
        $promptResult = [System.Windows.Forms.MessageBox]::Show("Would you like to create shortcuts on your Desktop, Start Menu, 'Send To' menu, and/or 'Open with' context menu?", "Create Shortcut/Registry?", "YesNo", "Question")

        if ($promptResult -eq 'Yes') {
            # Pre-check the boxes for the missing ones only
            Show-CreateShortcutWindow -ScriptDirectory $Global:ScriptDirectory -IconPath $localIconPath -CheckDesktop (-not $desktopExists) -CheckStartMenu (-not $startMenuExists) -CheckSendTo (-not $sendToExists) -CheckOpenWith (-not $openWithExists)
        }
    }

    # Initial Log of Registry Key
    try {
        Write-Status "Reading registry key: $RegKey"
        $userDir = (Get-ItemProperty -Path $RegKey -Name $RegValue).$RegValue
        Write-Status "Found SignalRGB User Directory: $userDir"
        $effectsDir = Join-Path -Path $userDir -ChildPath $EffectsSubFolder
        $componentsDir = Join-Path -Path $userDir -ChildPath $ComponentsSubFolder
        Write-Status "Effect install directory set to: $effectsDir"
        Write-Status "Component install directory set to: $componentsDir"
    }
    catch {
        Write-Status "ERROR: Could not read SignalRGB registry key on startup."
    }

    # --- NEW: RGB ANIMATION TIMER ---
    $rgbTimer = New-Object System.Windows.Forms.Timer
    $rgbTimer.Interval = 100 # Update every 100ms
    $rgbTimer.Add_Tick({
            # Increment Hue (0 to 360)
            $script:hue += 1
            if ($script:hue -ge 360) { $script:hue = 0 }
    
            # Calculate new dark color
            $newColor = Get-RGBColor -Hue $script:hue
    
            # Apply to Form
            $Global:mainForm.BackColor = $newColor
        })
    $rgbTimer.Start()
    # --- END NEW RGB ANIMATION TIMER ---

    # Run update check slightly delayed to allow UI to render first
    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 1000 # 1 seconds
    $timer.Add_Tick({
            $this.Stop() 
            Check-ForUpdates
        })
    $timer.Start()
})

# --- Enhanced Drag and Drop Logic (Visual Drop Zone) ---

# 0. Enable Drop on the larger layout container
$fileInputLayout.AllowDrop = $true

# 1. Define the "Drag Enter" Visuals (Green Glow + Text Change)
$onDragEnter = {
    param($s, $e)
    if ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
        $files = $e.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
    
        # Check for valid extensions
        $hasValidFile = $files | Where-Object { 
            $ext = [System.IO.Path]::GetExtension($_).ToLower()
            $ext -in @(".zip", ".html", ".json")
        }
    
        if ($hasValidFile) {
            $e.Effect = [System.Windows.Forms.DragDropEffects]::Copy
            
            # --- VISUAL FEEDBACK START ---
            # Light up the layout background with semi-transparent Green
            $fileInputLayout.BackColor = [System.Drawing.Color]::FromArgb(50, 0, 255, 0) 
            
            # Make the hint text bold and yellow
            $lblHint.Text = "!!! DROP FILES HERE !!!"
            $lblHint.ForeColor = [System.Drawing.Color]::Yellow
            $lblHint.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
            # --- VISUAL FEEDBACK END ---
        }
    }
}

# 2. Define "Drag Leave" (Reset Visuals if user cancels)
$onDragLeave = {
    param($s, $e)
    # --- RESET VISUALS ---
    $fileInputLayout.BackColor = [System.Drawing.Color]::Transparent
    $lblHint.Text = "Drag-and-drop a .zip, .html, or .json file here, or click Browse."
    $lblHint.ForeColor = [System.Drawing.Color]::White
    $lblHint.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
}

# 3. Define "Drag Drop" (Process Files + Reset Visuals)
$onDragDrop = {
    param($s, $e)
    
    # Immediately reset the visuals to normal
    $onDragLeave.Invoke($s, $e)

    # Process the dropped files
    $files = $e.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
    if ($files.Count -gt 0) {
        $validFiles = @()
        foreach ($file in $files) {
            $ext = [System.IO.Path]::GetExtension($file).ToLower()
            if ($ext -in @(".zip", ".html", ".json")) {
                $validFiles += $file
            }
        }
    
        if ($validFiles.Count -gt 0) {
            $fileList = $validFiles -join ";"
            $txtFilePath.Text = $fileList
            Write-Status "Files loaded: $($validFiles.Count) valid files."
            
            # Update hint to show readiness
            $lblHint.Text = "Files loaded! Click **'Install Item(s)'** to finish."
            $lblHint.ForeColor = [System.Drawing.Color]::Cyan
        }
        else {
            Write-Status "Ignored invalid files."
        }
    }
}

# 4. Attach these events to BOTH the TextBox and the Layout Panel
# This creates the "Big Drop Target" feel
$txtFilePath.Add_DragEnter($onDragEnter)
$txtFilePath.Add_DragLeave($onDragLeave)
$txtFilePath.Add_DragDrop($onDragDrop)

$fileInputLayout.Add_DragEnter($onDragEnter)
$fileInputLayout.Add_DragLeave($onDragLeave)
$fileInputLayout.Add_DragDrop($onDragDrop)

# Browse Button
$btnBrowse.Add_Click({
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        # MODIFIED: Updated filter to include .json
        $openFileDialog.Filter = "SignalRGB Files (*.zip, *.html, *.json)|*.zip;*.html;*.json|All Files (*.*)|*.*"
        $openFileDialog.Title = "Select Effect or Component File(s)"

        $openFileDialog.Multiselect = $true

        if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            # The FileNames property returns an array of selected files.
            $selectedFiles = $openFileDialog.FileNames
    
            # Filter for valid file types and join them with the semicolon separator
            $validFiles = @()
            foreach ($file in $selectedFiles) {
                # MODIFIED: Check for .json as well
                $ext = [System.IO.Path]::GetExtension($file).ToLower()
                if ($ext -in @(".zip", ".html", ".json")) {
                    $validFiles += $file
                }
            }
    
            if ($validFiles.Count -gt 0) {
                $fileList = $validFiles -join ";"
                $txtFilePath.Text = $fileList
                Write-Status "Files selected via browse: $($validFiles.Count) valid files selected."
        
                # Update hint label to match batch installation mode
                if ($validFiles.Count -gt 1) {
                    $lblHint.Text = "Multiple files loaded. Click **'Install Item(s)'** to process them sequentially."
                }
                else {
                    $lblHint.Text = "Drag-and-drop a .zip, .html, or .json file here, or click Browse."
                }
            }
            else {
                [System.Windows.Forms.MessageBox]::Show("You must select at least one .zip, .html, or .json file.", "No Valid Selection", "OK", "Warning") | Out-Null
                $txtFilePath.Text = ""
                Write-Status "Browse operation canceled or no valid files selected."
            }
        }
    })

# Install Button
$btnInstall.Add_Click({
        $filePathText = $txtFilePath.Text.Trim()

        if ([string]::IsNullOrWhiteSpace($filePathText)) {
            [System.Windows.Forms.MessageBox]::Show("Please select a file(s) to install first.", "No File Selected", "OK", "Warning") | Out-Null
            return
        }

        $filesToInstall = $filePathText.Split(";") | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

        if ($filesToInstall.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No valid file paths found.", "Error", "OK", "Error") | Out-Null
            return
        }

        # Tracking variable for the whole batch
        $batchRestartRequired = $false 

        # Disable button during install
        $btnInstall.Enabled = $false
        $btnInstall.Text = "Installing..."
        $Global:mainForm.Refresh()

        try {
            foreach ($file in $filesToInstall) {
                if (-not (Test-Path -Path $file)) {
                    Write-Status "ERROR: Skipped. File not found: $file"
                    [System.Windows.Forms.MessageBox]::Show("Skipping installation for '$file': File not found.", "File Not Found", "OK", "Error") | Out-Null
                    continue
                }
                Write-Status "--- Starting Sequential Install for: $file ---"
        
                # Capture the return status from Start-Installation
                $isRestartNeededForFile = Start-Installation -FilePath $file
        
                if ($isRestartNeededForFile) {
                    $batchRestartRequired = $true
                }

                Write-Status "--- Finished Install for: $file ---"
            }
    
            # FINAL RESTART PROMPT (Only executed if at least one file needed a restart)
            if ($batchRestartRequired) {
                Write-Status "One or more new items were installed. Prompting user to restart $AppName."
                $restartResult = [System.Windows.Forms.MessageBox]::Show("All files processed.`n`n$AppName must be restarted to load the new item(s). Restart now?", "Restart Required", "YesNo", "Question")
        
                if ($restartResult -eq 'Yes') {
                    Write-Status "User chose to restart."
                    Write-Status "Attempting to stop $AppName to force reload..."
                    try {
                        $process = Get-Process -Name $AppName -ErrorAction Stop
                        if ($process) {
                            Stop-Process -Name $AppName -Force
                            Write-Status "$AppName process stopped. It should restart automatically after a few seconds. Start it manually otherwise."
                        }
                        else {
                            Write-Status "$AppName is not running."
                        }
                    }
                    catch {
                        Write-Status "WARNING: Could not stop $AppName. It may be running as Administrator."
                        Write-Status "Please restart $AppName manually to see the new item(s)."
                        [System.Windows.Forms.MessageBox]::Show("Could not stop $AppName (it may be running as Administrator).`n`nPlease restart it manually to see the new item(s).", "Restart Failed", "OK", "Warning") | Out-Null
                    }
                }
                else {
                    Write-Status "User declined automatic restart. Manual restart required."
                    [System.Windows.Forms.MessageBox]::Show("Batch installation complete.`n`nManual restart of $AppName is required to see all new item(s).", "Manual Restart Needed", "OK", "Information") | Out-Null
                }
            }
            else {
                Write-Status "Batch installation complete. No restart was required."
                [System.Windows.Forms.MessageBox]::Show("Batch installation complete.`n`nNo restart of $AppName was required.", "Installation Complete", "OK", "Information") | Out-Null
            }
    
        }
        catch {
            Write-Status "Unhandled exception during file iteration: $($_.Exception.Message)"
        }
        finally {
            # Re-enable button
            $btnInstall.Enabled = $true
            $btnInstall.Text = "Install Item(s)" # MODIFIED
    
            # Clear text box after batch installation is complete
            $txtFilePath.Text = ""
            $lblHint.Text = "Drag-and-drop a .zip, .html, or .json file here, or click Browse." # MODIFIED
        }
    })

# Uninstall Button
$btnUninstall.Add_Click({
        try {
            $userDir = (Get-ItemProperty -Path $RegKey -Name $RegValue).$RegValue
            # MODIFIED: Pass the base UserDirectory to the uninstall window
            Show-UninstallWindow -UserDirectory $userDir
        }
        catch {
            Write-Status "ERROR: Could not get SignalRGB User Directory for uninstaller. $($_.Exception.Message)"
            [System.Windows.Forms.MessageBox]::Show("Could not find the SignalRGB User Directory.`nHave you run SignalRGB at least once?", "Error", "OK", "Error") | Out-Null
        }
    })

# Disclaimer Button
$btnDisclaimer.Add_Click({
        Show-DisclaimerWindow
    })

# --- NEW: Open Folder Button Event Handler ---
$btnOpenFolder.Add_Click({
        Write-Status "Attempting to open user folder..."
        $userDir = $null
        try {
            $userDir = (Get-ItemProperty -Path $RegKey -Name $RegValue).$RegValue
        
            if ($userDir -and (Test-Path -Path $userDir)) {
                Write-Status "Opening folder in Explorer: $userDir"
                Invoke-Item -Path $userDir
            }
            else {
                Write-Status "ERROR: WhirlwindFX folder path not found or is invalid: $userDir"
                [System.Windows.Forms.MessageBox]::Show("Could not find the WhirlwindFX folder path from the registry.", "Error", "OK", "Error") | Out-Null
            }
        }
        catch {
            Write-Status "ERROR: Could not read registry key to find folder. $($_.Exception.Message)"
            [System.Windows.Forms.MessageBox]::Show("Could not read the registry to find the SignalRGB User Directory.", "Error", "OK", "Error") | Out-Null
        }
    })

# --- Show the Form ---
# Write-Status "GUI Initialized. Showing main window."
[System.Windows.Forms.Application]::Run($Global:mainForm)