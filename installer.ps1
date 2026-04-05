#Requires -Version 5.1

# --- Versioning ---
$ScriptVersion = "1.8"
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
                        if (-not [string]::IsNullOrWhiteSpace($IconPath) -and (Test-Path -LiteralPath $IconPath)) {
                            $shortcut.IconLocation = "$IconPath, 0"
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
                        if (-not [string]::IsNullOrWhiteSpace($IconPath) -and (Test-Path -LiteralPath $IconPath)) {
                            $shortcut.IconLocation = "$IconPath, 0"
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
                        if (-not [string]::IsNullOrWhiteSpace($IconPath) -and (Test-Path -LiteralPath $IconPath)) {
                            $shortcut.IconLocation = "$IconPath, 0"
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
                    Set-OpenWithRegistryKeys -AppName "RGBJunkie Installer" -AppPath $Global:ScriptFullPath -FileExtensions @(".zip", ".html", ".json")
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
        if (Test-Path -LiteralPath $effectsBasePath) {
            try {
                # FIX: Force the result into an array to prevent PowerShell from unrolling a single string into characters
                $effects = @(Get-ChildItem -LiteralPath $effectsBasePath -Directory | ForEach-Object { $_.Name } | Sort-Object)
                if ($effects.Count -gt 0) {
                    $clbEffects.Items.AddRange([object[]]$effects)
                }
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
        if (Test-Path -LiteralPath $componentsBasePath) {
            try {
                # --- MODIFICATION: Get component names from JSON files ---
                $componentFiles = @(Get-ChildItem -LiteralPath $componentsBasePath -Filter "*.json")
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
                $sortedList = @($componentList | Sort-Object -Property DisplayName)
                
                if ($sortedList.Count -gt 0) {
                    $clbComponents.Items.AddRange([object[]]$sortedList)
                }
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
            $originalEffectFolders = @($clbEffects.Items)
            $currentAlwaysTitle = ""
            $currentAlwaysFolder = $null
    
            try {
                $currentAlwaysTitle = (Get-ItemProperty -LiteralPath "HKCU:\Software\WhirlwindFX\SignalRgb\effects\selected" -Name "always" -ErrorAction SilentlyContinue).always
            }
            catch {
                Write-Status "Could not read current 'always' key. Will not update registry on delete."
            }

            if (-not [string]::IsNullOrWhiteSpace($currentAlwaysTitle)) {
                # Find the folder name that corresponds to the active title
                foreach ($folderName in $originalEffectFolders) {
                
                    # SCRIPT FIX: Join-Path uses -Path, Get-ChildItem uses -LiteralPath
                    $folderPath = Join-Path -Path $effectsBasePath -ChildPath $folderName
                    $htmlFiles = @(Get-ChildItem -LiteralPath $folderPath -Filter "*.html" -ErrorAction SilentlyContinue)
                
                    if ($htmlFiles.Count -gt 0) {
                        # Check the first HTML file found
                        $title = Get-EffectTitleFromHtml -HtmlFilePath $htmlFiles[0].FullName
                    
                        if (-not [string]::IsNullOrWhiteSpace($title) -and $currentAlwaysTitle.Equals($title, [StringComparison]::OrdinalIgnoreCase)) {
                            $currentAlwaysFolder = $folderName
                            break
                        }
                    }
                }
            }
    
            $activeEffectFolderWasDeleted = $false
            # Ensure selectedEffectItems is an array even if only 1 item is selected
            $selectedEffectItems = @($selectedItems | Where-Object { $_ -like '[Effect]*' })
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
                
                    # SCRIPT FIX: Join-Path uses -Path, Get-ChildItem uses -LiteralPath
                    $newFolderPath = Join-Path -Path $effectsBasePath -ChildPath $newEffectFolder
                    $newHtmlFiles = @(Get-ChildItem -LiteralPath $newFolderPath -Filter "*.html" -ErrorAction SilentlyContinue)
                
                    if ($newHtmlFiles.Count -gt 0) {
                        $newEffectTitle = Get-EffectTitleFromHtml -HtmlFilePath $newHtmlFiles[0].FullName
                        Write-Status "Setting new active effect to: '$newEffectTitle'"
                        Set-ActiveEffectRegistryKeys -NewEffectTitle $newEffectTitle
                    }
                    else {
                        Write-Status "WARNING: Could not find an HTML file in '$newEffectFolder' to set as active."
                    }
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
                        # FIX: Use LiteralPath to prevent bracket errors on folders
                        if (Test-Path -LiteralPath $itemFolder) {
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
                    
                        # FIX: Use LiteralPath
                        if (Test-Path -LiteralPath $jsonFilePath) {
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
                        if (Test-Path -LiteralPath $pngFilePath) {
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
RGBJunkie Effect Installer - Terms of Use

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
$Global:mainForm.Text = "RGBJunkie Effect & Component Installer - v$ScriptVersion"
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
# $localIconPath = Join-Path -Path $Global:ScriptDirectory -ChildPath "rgbjunkielogo2.ico"
# $localLogoPath = Join-Path -Path $Global:ScriptDirectory -ChildPath "rgbjunkie.png"

# Load Logo
$logoLoaded = $false
if (-not [string]::IsNullOrWhiteSpace($localLogoPath) -and (Test-Path -Path $localLogoPath)) {
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
                Show-CreateShortcutWindow -ScriptDirectory $Global:ScriptDirectory -IconPath $Global:ScriptFullPath -CheckDesktop (-not $desktopExists) -CheckStartMenu (-not $startMenuExists) -CheckSendTo (-not $sendToExists) -CheckOpenWith (-not $openWithExists)
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