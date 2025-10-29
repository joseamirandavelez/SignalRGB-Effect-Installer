#Requires -Version 5.1

# --- Load Windows Forms Assembly ---
# Load these first to ensure all UI elements (even error popups) get modern styling
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Set high DPI awareness and visual styles
try {
    [System.Windows.Forms.Application]::SetHighDpiMode('SystemAware')
}
catch {}
[System.Windows.Forms.Application]::EnableVisualStyles()


<#
.SYNOPSIS
    A GUI-based installer for SignalRGB effects.
    Handles .html, .png, and .zip files.
    Manages folder creation, conflict detection (file and <title> tag),
    overwrite/rename logic, and app restart.
    Includes an uninstaller and shortcut creation.
#>

# --- Script-Wide Variables ---
$Global:ScriptDirectory = $null
$Global:ScriptFullPath = $null
$Global:DesktopShortcutPath = $null
$Global:StartMenuShortcutPath = $null
$Global:SendToShortcutPath = $null

try {
    # Set global paths. This is more reliable than $PSScriptRoot
    $Global:ScriptFullPath = $MyInvocation.MyCommand.Path
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

# --- Load Remaining Assemblies ---
Add-Type -AssemblyName System.IO.Compression.FileSystem

# --- Helper Functions (Ensuring all are defined before first call) ---

function Write-Status {
    param (
        [string]$Message,
        [string]$Type = "INFO"
    )
    
    # Check if the global variable is a Forms control and has a created handle
    if ($script:txtStatus -is [System.Windows.Forms.TextBox] -and $script:txtStatus.IsHandleCreated) {
        $timestamp = Get-Date -Format "HH:mm:ss"
        $formattedMessage = "$timestamp - $Type - $Message`r`n"
        
        # Use BeginInvoke for thread-safe UI updates
        $script:txtStatus.BeginInvoke([Action[string]] {
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
    $layout.Controls.Add($lblInfo, 0, 0) # Row 0

    $chkDesktop = New-Object System.Windows.Forms.CheckBox
    $chkDesktop.Text = "On the Desktop"
    $chkDesktop.Checked = $CheckDesktop
    $chkDesktop.Dock = 'Fill'
    $layout.Controls.Add($chkDesktop, 0, 1) # Row 1

    $chkStartMenu = New-Object System.Windows.Forms.CheckBox
    $chkStartMenu.Text = "In the Start Menu (under 'SignalRGB Tools')"
    $chkStartMenu.Checked = $CheckStartMenu
    $chkStartMenu.Dock = 'Fill'
    $layout.Controls.Add($chkStartMenu, 0, 2) # Row 2

    $chkSendTo = New-Object System.Windows.Forms.CheckBox
    $chkSendTo.Text = "In the 'Send To' menu (for quick installs)"
    $chkSendTo.Checked = $CheckSendTo
    $chkSendTo.Dock = 'Fill'
    $layout.Controls.Add($chkSendTo, 0, 3) # Row 3

    $chkOpenWith = New-Object System.Windows.Forms.CheckBox
    $chkOpenWith.Text = "Add to 'Open with' context menu"
    $chkOpenWith.Checked = $CheckOpenWith
    $chkOpenWith.Dock = 'Fill'
    $layout.Controls.Add($chkOpenWith, 0, 4) # Row 4

    $btnCreate = New-Object System.Windows.Forms.Button
    $btnCreate.Text = "Create"
    $btnCreate.DialogResult = [System.Windows.Forms.DialogResult]::OK # Set DialogResult for clean closing
    $btnCreate.Dock = 'Fill'
    $btnCreate.Anchor = 'Top'
    $btnCreate.Height = 30
    $btnCreate.Margin = [System.Windows.Forms.Padding]::new(0, 10, 0, 0)
    $layout.Controls.Add($btnCreate, 0, 5) # Row 5

    $btnCreate.Add_Click({
            # The execution of this script block performs the action and then closes the dialog
            try {
                # WScript.Shell is the object that creates shortcuts
                $wsShell = New-Object -ComObject WScript.Shell

                $targetFile = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
                $arguments = "-WindowStyle Hidden -ExecutionPolicy Bypass -File ""$Global:ScriptFullPath"""

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
                    Set-OpenWithRegistryKeys -AppName "SignalRGB Installer" -AppPath $Global:ScriptFullPath -FileExtensions @(".zip", ".html")
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
        [string]$EffectsBasePath
    )

    $uninstallForm = New-Object System.Windows.Forms.Form
    $uninstallForm.Text = "Uninstall Effects"
    $uninstallForm.Size = New-Object System.Drawing.Size(500, 400)
    $uninstallForm.MinimumSize = New-Object System.Drawing.Size(300, 200)
    $uninstallForm.FormBorderStyle = 'Sizable'
    $uninstallForm.MaximizeBox = $true
    $uninstallForm.MinimizeBox = $true
    $uninstallForm.StartPosition = 'CenterParent'
    if ($Global:mainForm -and $Global:mainForm.Icon) { $uninstallForm.Icon = $Global:mainForm.Icon }

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
    $lblInfo.Text = "Select effects to delete:"
    $lblInfo.Dock = 'Fill'
    $lblInfo.TextAlign = 'MiddleLeft'
    $headerPanel.Controls.Add($lblInfo, 0, 0)

    $btnRefresh = New-Object System.Windows.Forms.Button
    $btnRefresh.Text = "Refresh List"
    $btnRefresh.Dock = 'None'
    $btnRefresh.Anchor = 'Right'
    $headerPanel.Controls.Add($btnRefresh, 1, 0)

    # --- Effect List ---
    $clbEffects = New-Object System.Windows.Forms.CheckedListBox
    $clbEffects.Dock = 'Fill'
    $clbEffects.Margin = [System.Windows.Forms.Padding]::new(0, 5, 0, 5)
    $layout.Controls.Add($clbEffects, 0, 1)

    # --- Footer: Delete Button ---
    $btnDelete = New-Object System.Windows.Forms.Button
    $btnDelete.Text = "Delete Selected"
    $btnDelete.Dock = 'Fill'
    $btnDelete.Height = 30
    $layout.Controls.Add($btnDelete, 0, 2)

    # --- Functions for this window ---
    $populateList = {
        $clbEffects.Items.Clear()
        Write-Status "Refreshing effect list from: $EffectsBasePath"
        if (-not (Test-Path -Path $EffectsBasePath)) {
            Write-Status "Effects folder does not exist. Nothing to list."
            return
        }
        
        try {
            # Find all .html files within subdirectories of the Effects folder
            $effects = Get-ChildItem -Path $EffectsBasePath -Recurse -Filter "*.html" | ForEach-Object {
                [PSCustomObject]@{
                    Name     = $_.Directory.Name # Get the folder name
                    HtmlFile = $_.FullName
                    PngFile  = Join-Path -Path $_.Directory.FullName -ChildPath ($_.BaseName + ".png")
                    Folder   = $_.Directory.FullName
                }
            } | Sort-Object Name
            
            # Add unique folder names to the list
            $uniqueFolders = $effects | Select-Object -ExpandProperty Name -Unique
            foreach ($folderName in $uniqueFolders) {
                $clbEffects.Items.Add($folderName, $false) | Out-Null
            }
            Write-Status "Found $($uniqueFolders.Count) effects."
        }
        catch {
            Write-Status "ERROR scanning for effects: $($_.Exception.Message)"
        }
    }

    $btnRefresh.Add_Click({
            Invoke-Command $populateList
        })

    $btnDelete.Add_Click({
            $selectedItems = $clbEffects.CheckedItems
            if ($selectedItems.Count -eq 0) {
                [System.Windows.Forms.MessageBox]::Show("Please select one or more effects to delete.", "No Selection", "OK", "Information") | Out-Null
                return
            }

            # --- NEW: Logic to determine if the active effect is being deleted ---
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
                    $effectHtmlPath = Join-Path -Path $EffectsBasePath -ChildPath "$folderName\$folderName.html"
                    $title = Get-EffectTitleFromHtml -HtmlFilePath $effectHtmlPath
                    # --- FIX: Reversed the .Equals() call to prevent null reference error ---
                    if (-not [string]::IsNullOrWhiteSpace($title) -and $currentAlwaysTitle.Equals($title, [StringComparison]::OrdinalIgnoreCase)) {
                        $currentAlwaysFolder = $folderName
                        break
                    }
                }
            }
        
            $activeEffectFolderWasDeleted = $false
            if ($currentAlwaysFolder -and $selectedItems -contains $currentAlwaysFolder) {
                $activeEffectFolderWasDeleted = $true
                Write-Status "Active effect '$currentAlwaysTitle' is scheduled for deletion."
            }
            # --- End new logic block ---

            # --- MOVED: Set new active effect *before* deleting files ---
            if ($activeEffectFolderWasDeleted) {
                Write-Status "Updating active effect registry keys..."
                $remainingEffectFolders = @($originalEffectFolders | Where-Object { $_ -notin $selectedItems })
            
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
                    $newEffectHtmlPath = Join-Path -Path $EffectsBasePath -ChildPath "$newEffectFolder\$newEffectFolder.html"
                    $newEffectTitle = Get-EffectTitleFromHtml -HtmlFilePath $newEffectHtmlPath
                
                    Write-Status "Setting new active effect to: '$newEffectTitle'"
                    Set-ActiveEffectRegistryKeys -NewEffectTitle $newEffectTitle
                }
            }
            # --- End of moved block ---

            foreach ($itemName in $selectedItems) {
                $effectFolder = Join-Path -Path $EffectsBasePath -ChildPath $itemName
                Write-Status "Deleting effect: $itemName"
                try {
                    # Find all html and png files in that folder
                    # Fix: Ensure $filesToDelete is always an array by using @()
                    $filesToDelete = @(Get-ChildItem -Path $effectFolder -Filter "*.html")
                    $filesToDelete += @(Get-ChildItem -Path $effectFolder -Filter "*.png")

                    foreach ($file in $filesToDelete) {
                        Write-Status "Deleting file: $($file.Name)"
                        Remove-Item -Path $file.FullName -Force
                    }
                
                    # Check if folder is now empty
                    if ((Get-ChildItem -Path $effectFolder -Force | Measure-Object).Count -eq 0) {
                        Write-Status "Folder '$itemName' is empty, deleting it."
                        Remove-Item -Path $effectFolder -Recurse -Force
                    }
                    else {
                        Write-Status "Folder '$itemName' is not empty, will not delete folder."
                    }
                    Write-Status "Successfully deleted effect '$itemName'."
                }
                catch {
                    Write-Status "ERROR deleting '$itemName': $($_.Exception.Message)"
                    [System.Windows.Forms.MessageBox]::Show("Error deleting '$itemName':`n$($_.Exception.Message)", "Error", "OK", "Error") | Out-Null
                }
            }
        
            Write-Status "Deletion complete."
        
            # Refresh the list
            Invoke-Command $populateList
        })

    # --- Initial Load ---
    Invoke-Command $populateList
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
        # --- FIX: Reversed the .Equals() call to prevent null reference error ---
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
    
    # --- FIX: Rebuilt this function with TableLayoutPanel ---
    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = "Conflict Detected"
    $dialog.Size = New-Object System.Drawing.Size(400, 170) # Increased height
    $dialog.FormBorderStyle = 'FixedDialog'
    $dialog.MaximizeBox = $false
    $dialog.MinimizeBox = $false
    $dialog.StartPosition = 'CenterParent'
    if ($Global:mainForm -and $Global:mainForm.Icon) { $dialog.Icon = $Global:mainForm.Icon }
    
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
    $lblMessage.TextAlign = 'MiddleCenter'
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
    $buttonLayout.Controls.Add($btnOverwrite, 0, 0)
    
    $btnRename = New-Object System.Windows.Forms.Button
    $btnRename.Text = "Rename"
    $btnRename.DialogResult = 'Retry'
    $btnRename.Dock = 'None'
    $btnRename.Anchor = 'Top, Left, Right'
    $btnRename.Height = 30
    $buttonLayout.Controls.Add($btnRename, 1, 0)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.DialogResult = 'Cancel'
    $btnCancel.Dock = 'None'
    $btnCancel.Anchor = 'Top, Left, Right'
    $btnCancel.Height = 30
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
    $dialog.Text = "Rename Effect"
    $dialog.Size = New-Object System.Drawing.Size(350, 150)
    $dialog.FormBorderStyle = 'FixedDialog'
    $dialog.MaximizeBox = $false
    $dialog.MinimizeBox = $false
    $dialog.StartPosition = 'CenterParent'
    if ($Global:mainForm -and $Global:mainForm.Icon) { $dialog.Icon = $Global:mainForm.Icon }

    $lblInfo = New-Object System.Windows.Forms.Label
    $lblInfo.Text = "Enter a new name for the effect:"
    $lblInfo.Location = New-Object System.Drawing.Point(20, 20)
    $lblInfo.AutoSize = $true
    $dialog.Controls.Add($lblInfo)
    
    $txtNewName = New-Object System.Windows.Forms.TextBox
    $txtNewName.Text = $OldName
    $txtNewName.Location = New-Object System.Drawing.Point(20, 50)
    $txtNewName.Size = New-Object System.Drawing.Size(300, 20)
    $dialog.Controls.Add($txtNewName)
    
    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "OK"
    $btnOK.DialogResult = 'OK'
    $btnOK.Location = New-Object System.Drawing.Point(160, 90)
    $dialog.Controls.Add($btnOK)
    
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.DialogResult = 'Cancel'
    $btnCancel.Location = New-Object System.Drawing.Point(240, 90)
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
SignalRGB Effect Installer - Terms of Use

1. No Warranty
This Tool is provided "as-is", without any warranties of any kind, express or implied. The developer makes no guarantees regarding its functionality, reliability, or suitability for any particular purpose.

2. No Affiliation
The developer of this Tool is not associated with SignalRGB or its parent company, WhirlwindFX. This is an unofficial, third-party application.

3. Limitation of Liability
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
    $layout.Controls.Add($rtbDisclaimer, 0, 0)

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "OK"
    $btnOK.DialogResult = 'OK'
    $btnOK.Dock = 'None'
    $btnOK.Anchor = 'Right'
    $btnOK.Height = 30
    $btnOK.Width = 80
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
    $disclaimerForm.ShowDialog($Global:mainForm) | Out-Null
    $disclaimerForm.Dispose()
}


# --- Main Installation Logic ---

function Start-Installation {
    param (
        [string]$FilePath
    )
    
    Write-Status "Starting installation for: $FilePath"
    
    # Flag to indicate if SignalRGB must be restarted after this installation.
    # A restart is required for New Installs and Renames, but NOT for Overwrites.
    $restartRequired = $false 
    
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

    # --- Define the correct base install folder ---
    $installBasePath = Join-Path -Path $userDir -ChildPath $EffectsSubFolder
    
    # --- Ensure the 'Effects' folder exists ---
    if (-not (Test-Path -Path $installBasePath)) {
        Write-Status "'$EffectsSubFolder' folder not found. Creating it..."
        try {
            New-Item -Path $installBasePath -ItemType Directory -Force | Out-Null
            Write-Status "Created folder: $installBasePath"
        }
        catch {
            Write-Status "ERROR: Could not create folder: $installBasePath. $($_.Exception.Message)"
            [System.Windows.Forms.MessageBox]::Show("Error: Could not create folder:`n$installBasePath", "Error", "OK", "Error") | Out-Null
            return $false # Return false on failure
        }
    }
    
    $sourceHtmlFile = $null
    $sourcePngFile = $null
    $effectName = $null
    $tempExtractFolder = $null
    
    $extension = [System.IO.Path]::GetExtension($FilePath).ToLower()
    
    try {
        # 2. Prepare source files based on input (zip, html, or png)
        
        if ($extension -eq ".zip") {
            Write-Status "Zip file detected. Extracting..."
            $tempExtractFolder = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ([System.IO.Path]::GetRandomFileName())
            [System.IO.Compression.ZipFile]::ExtractToDirectory($FilePath, $tempExtractFolder)
            Write-Status "Extracted to: $tempExtractFolder"
            
            # Find the html and png files inside
            $sourceHtmlFile = (Get-ChildItem -Path $tempExtractFolder -Recurse -Filter "*.html" | Select-Object -First 1).FullName
            if (-not $sourceHtmlFile) {
                Write-Status "ERROR: No .html file found in the zip archive."
                [System.Windows.Forms.MessageBox]::Show("Error: No .html file found in the zip archive.", "Zip Error", "OK", "Error") | Out-Null
                return $false # Return false on failure
            }
            
            $effectName = [System.IO.Path]::GetFileNameWithoutExtension($sourceHtmlFile)
            $sourcePngFile = (Get-ChildItem -Path (Split-Path $sourceHtmlFile) -Filter "$effectName.png" | Select-Object -First 1).FullName
            if (-not $sourcePngFile) {
                Write-Status "No matching .png found in zip. This might be ok."
            }
            
        }
        elseif ($extension -eq ".html") {
            $sourceHtmlFile = $FilePath
            $effectName = [System.IO.Path]::GetFileNameWithoutExtension($sourceHtmlFile)
            $sourcePngFile = Join-Path -Path (Split-Path $FilePath) -ChildPath "$effectName.png"
            if (-not (Test-Path -Path $sourcePngFile)) {
                $sourcePngFile = $null # Don't try to copy a non-existent file
                Write-Status "No matching .png found for $effectName. This might be ok."
            }
            
        }
        else {
            Write-Status "ERROR: Invalid file type. Please select a .zip or .html file."
            [System.Windows.Forms.MessageBox]::Show("Invalid file type. Please select a .zip or .html file.", "Invalid File", "OK", "Error") | Out-Null
            return $false # Return false on failure
        }

        if (-not $sourceHtmlFile) {
            Write-Status "ERROR: Could not determine source .html file."
            return $false # Return false on failure
        }

        Write-Status "Effect Name (from file): $effectName"
        Write-Status "Source HTML: $sourceHtmlFile"
        Write-Status "Source PNG: $sourcePngFile"

        # 3. Conflict Detection Loop
        $installConfirmed = $false
        $currentEffectName = $effectName
        $currentHtmlFile = $sourceHtmlFile
        $currentPngFile = $sourcePngFile
        $isOverwrite = $false 
        
        while (-not $installConfirmed) {
            $destFolder = Join-Path -Path $installBasePath -ChildPath $currentEffectName
            # FIX: Removed assignment to unused variable 'destHtmlFile'
            Join-Path -Path $destFolder -ChildPath ($currentEffectName + ".html") | Out-Null

            # Get title from HTML
            $currentEffectTitle = Get-EffectTitleFromHtml -HtmlFilePath $currentHtmlFile
            # No fallback needed here, Get-EffectTitleFromHtml will return filename if title is missing
            Write-Status "Effect Title (from HTML): $currentEffectTitle"
            
            # Check for conflicts
            $folderExists = Test-Path -Path $destFolder
            $titleConflictFile = Find-EffectTitleConflict -NewEffectTitle $currentEffectTitle -EffectsBasePath $installBasePath
            
            if ($folderExists -or $titleConflictFile) {
                # Conflict found!
                $conflictMessage = "A conflict was found.`n`n"
                if ($folderExists) { $conflictMessage += "- Folder '$currentEffectName' already exists.`n" }
                if ($titleConflictFile) { $conflictMessage += "- Title '$currentEffectTitle' is already used by `n  $titleConflictFile`n" }
                $conflictMessage += "`nWould you like to Overwrite, Rename, or Cancel?"
                
                Write-Status "Conflict detected: $conflictMessage"
                $userChoice = Show-ConflictDialog -Message $conflictMessage
                
                if ($userChoice -eq 'Overwrite') {
                    Write-Status "User chose to Overwrite."
                    $isOverwrite = $true
                    $installConfirmed = $true
                    
                }
                elseif ($userChoice -eq 'Rename') {
                    Write-Status "User chose to Rename."
                    $newName = Show-RenameDialog -OldName $currentEffectName
                    
                    if ([string]::IsNullOrWhiteSpace($newName)) {
                        Write-Status "Rename cancelled by user."
                        return $false # Return false on user cancel
                    }
                    
                    if ($newName.Equals($currentEffectName, [StringComparison]::OrdinalIgnoreCase)) {
                        Write-Status "New name is the same as the old name. No changes made."
                        # Loop will repeat
                        continue
                    }

                    Write-Status "New name selected: $newName"
                    $currentEffectName = $newName
                    
                    # Update HTML <title> tag
                    Set-EffectTitleInHtml -HtmlFilePath $currentHtmlFile -NewTitle $newName | Out-Null
                    # Note: We will re-read this new title at the start of the next loop.
                    
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
        Write-Status "Installing '$currentEffectName' to $destFolder"
        
        # Ensure destination folder exists
        if (-not (Test-Path -Path $destFolder)) {
            New-Item -Path $destFolder -ItemType Directory -Force | Out-Null
        }
        
        # Define final destination file paths
        $finalHtmlPath = Join-Path -Path $destFolder -ChildPath ($currentEffectName + ".html")
        $finalPngPath = Join-Path -Path $destFolder -ChildPath ($currentEffectName + ".png")

        # Copy HTML
        Copy-Item -Path $currentHtmlFile -Destination $finalHtmlPath -Force
        Write-Status "Copied HTML to: $finalHtmlPath"
        
        # Copy PNG if it exists
        if ($currentPngFile -and (Test-Path -Path $currentPngFile)) {
            Copy-Item -Path $currentPngFile -Destination $finalPngPath -Force
            Write-Status "Copied PNG to: $finalPngPath"
        }
        
        # 5. Set Registry Keys (Always set the registry key to make it the active effect)
        $finalEffectTitle = Get-EffectTitleFromHtml -HtmlFilePath $finalHtmlPath
        
        Set-ActiveEffectRegistryKeys -NewEffectTitle $finalEffectTitle
        
        # 6. Determine restart necessity
        if (-not $isOverwrite) {
            # A new installation or a rename occurred, so a restart is required.
            $restartRequired = $true
            Write-Status "Installation finished. Restart is required for this effect."
        }
        else {
            # It was an overwrite. Only registry change needed.
            $restartRequired = $false
            Write-Status "Overwrite finished. No restart required."
        }
        
        [System.Windows.Forms.MessageBox]::Show("Effect '$currentEffectName' installed/updated successfully and set as active.", "Installation Complete", "OK", "Information") | Out-Null
        
        return $restartRequired

    }
    catch {
        Write-Status "FATAL ERROR during installation: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("A fatal error occurred during installation of '$currentEffectName':`n$($_.Exception.Message)", "Fatal Error", "OK", "Error") | Out-Null
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
            
            if ($ext -eq ".zip" -or $ext -eq ".html") {
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
        $restartResult = [System.Windows.Forms.MessageBox]::Show("Batch installation complete.`n`n$AppName must be restarted to load the new effect(s). Restart now?", "Restart Required", "YesNo", "Question")
        
        if ($restartResult -eq 'Yes') {
            Write-Host "User chose to restart."
            try {
                Stop-Process -Name $AppName -Force -ErrorAction Stop
                Write-Host "$AppName process stopped. It should restart automatically."
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Could not stop $AppName. Please restart it manually.", "Restart Failed", "OK", "Warning") | Out-Null
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
$Global:mainForm.Text = "SignalRGB Effect Installer"
$Global:mainForm.Size = New-Object System.Drawing.Size(650, 500) # Increased width from 640
$Global:mainForm.FormBorderStyle = 'FixedSingle'
$Global:mainForm.MaximizeBox = $false
$Global:mainForm.StartPosition = 'CenterScreen'

# --- Main Layout Table ---
$mainLayout = New-Object System.Windows.Forms.TableLayoutPanel
$mainLayout.Dock = 'Fill'
$mainLayout.Padding = 10
$mainLayout.ColumnCount = 1
$mainLayout.RowCount = 4 # Row for Logo, Row for File Input, Row for Buttons, Row for Status Box
$mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null # Logo
$mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null # File Input
$mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null # Buttons
$mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null # Status Box
$Global:mainForm.Controls.Add($mainLayout)

# --- Row 0: Logo ---
$picLogo = New-Object System.Windows.Forms.PictureBox
$picLogo.SizeMode = 'Zoom'
$picLogo.Size = New-Object System.Drawing.Size(128, 128)
$picLogo.Dock = 'None'
$picLogo.Anchor = 'Top'
$picLogo.Margin = [System.Windows.Forms.Padding]::new(0, 0, 0, 10)

$localIconPath = Join-Path -Path $Global:ScriptDirectory -ChildPath "icon.ico"
$localLogoPath = Join-Path -Path $Global:ScriptDirectory -ChildPath "logo.png"

# Load Logo
$logoLoaded = $false
if (Test-Path -Path $localLogoPath) {
    try {
        $logoStream = New-Object System.IO.FileStream($localLogoPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
        $picLogo.Image = [System.Drawing.Image]::FromStream($logoStream)
        $logoStream.Close()
        $logoStream.Dispose()
        $mainLayout.Controls.Add($picLogo, 0, 0)
        $logoLoaded = $true
    }
    catch {
        Write-Status "ERROR: Could not load logo.png: $($_.Exception.Message)"
    }
}
else {
    Write-Status "No logo.png found. Hiding logo area."
}

# --- Row 1: File Input ---
$fileInputLayout = New-Object System.Windows.Forms.TableLayoutPanel
$fileInputLayout.Dock = 'Fill'
$fileInputLayout.AutoSize = $true
$fileInputLayout.ColumnCount = 3
$fileInputLayout.RowCount = 3
$fileInputLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
$fileInputLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
$fileInputLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 95))) | Out-Null # Increased from 85
$mainLayout.Controls.Add($fileInputLayout, 0, 1)

$lblFile = New-Object System.Windows.Forms.Label
$lblFile.Text = "Effect File:"
$lblFile.Dock = 'Fill'
$lblFile.TextAlign = 'MiddleLeft'
$lblFile.Margin = [System.Windows.Forms.Padding]::new(0, 0, 5, 0)
$fileInputLayout.Controls.Add($lblFile, 0, 0)

$txtFilePath = New-Object System.Windows.Forms.TextBox
$txtFilePath.Dock = 'Fill'
$txtFilePath.AllowDrop = $true
$fileInputLayout.Controls.Add($txtFilePath, 1, 0)

$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = "Browse..."
$btnBrowse.Dock = 'None' # Changed from Fill
$btnBrowse.AutoSize = $true # Let button size to its text
$btnBrowse.Anchor = 'Top, Left' # Set Anchor
$btnBrowse.Margin = [System.Windows.Forms.Padding]::new(5, 0, 0, 0)
$fileInputLayout.Controls.Add($btnBrowse, 2, 0)

$lblHint = New-Object System.Windows.Forms.Label
$lblHint.Text = "Drag-and-drop a .zip or .html file here, or click Browse."
$lblHint.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
$lblHint.Dock = 'Fill'
$lblHint.Margin = [System.Windows.Forms.Padding]::new(0, 0, 0, 5)
$fileInputLayout.Controls.Add($lblHint, 1, 1) # Span 2
$fileInputLayout.SetColumnSpan($lblHint, 2)

# --- NEW: Add LinkLabel for Effect Builder ---
$lnkBuilder = New-Object System.Windows.Forms.LinkLabel
$lnkBuilder.Text = "Create your own effects with the Effect Builder"
# Link starts at char 32 ("Effect Builder"), length 14
$lnkBuilder.Links.Add(32, 14, "https://effectbuilder.github.io/") | Out-Null 
$lnkBuilder.Dock = 'Fill'
$lnkBuilder.TextAlign = 'MiddleLeft'
$lnkBuilder.Margin = [System.Windows.Forms.Padding]::new(0, 0, 0, 5)
$fileInputLayout.Controls.Add($lnkBuilder, 1, 2) # Add to new row 2
$fileInputLayout.SetColumnSpan($lnkBuilder, 2)

$lnkBuilder.Add_LinkClicked({
        param($s, $e) # FIX: Renamed $sender to $s
        try {
            [System.Diagnostics.Process]::Start($e.Link.LinkData)
            $e.Link.Visited = $true
        }
        catch {
            Write-Status "ERROR: Could not open URL: $($e.Link.LinkData)"
        }
    })
# --- End of NEW LinkLabel ---


# --- Row 2: Buttons ---
$buttonLayout = New-Object System.Windows.Forms.TableLayoutPanel
$buttonLayout.Dock = 'Fill'
$buttonLayout.AutoSize = $true
$buttonLayout.ColumnCount = 3 # Install, Uninstall, Spacer
$buttonLayout.RowCount = 1
$buttonLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null # Added this line
$buttonLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33.3))) | Out-Null
$buttonLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33.3))) | Out-Null
$buttonLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33.3))) | Out-Null
$buttonLayout.Margin = [System.Windows.Forms.Padding]::new(0, 5, 0, 10)
$mainLayout.Controls.Add($buttonLayout, 0, 2)

$btnInstall = New-Object System.Windows.Forms.Button
$btnInstall.Text = "Install Effect"
$btnInstall.Dock = 'None' # Changed from Fill
$btnInstall.Anchor = 'Top, Left, Right' # Added Anchor
$btnInstall.Height = 30 # Reduced height
$buttonLayout.Controls.Add($btnInstall, 0, 0)

$btnUninstall = New-Object System.Windows.Forms.Button
$btnUninstall.Text = "Uninstall an Effect..."
$btnUninstall.Dock = 'None' # Changed from Fill
$btnUninstall.Anchor = 'Top, Left, Right' # Added Anchor
$btnUninstall.Height = 30 # Reduced height
$buttonLayout.Controls.Add($btnUninstall, 1, 0)

# --- NEW: Disclaimer Button ---
$btnDisclaimer = New-Object System.Windows.Forms.Button
$btnDisclaimer.Text = "Disclaimer..."
$btnDisclaimer.Dock = 'None'
$btnDisclaimer.Anchor = 'Top, Left, Right'
$btnDisclaimer.Height = 30
$buttonLayout.Controls.Add($btnDisclaimer, 2, 0)

# --- Row 3: Status Box ---
$script:txtStatus = New-Object System.Windows.Forms.TextBox
$script:txtStatus.Multiline = $true
$script:txtStatus.ReadOnly = $true
$script:txtStatus.ScrollBars = 'Vertical'
$script:txtStatus.Dock = 'Fill'
$script:txtStatus.Font = New-Object System.Drawing.Font("Consolas", 9)
$mainLayout.Controls.Add($script:txtStatus, 0, 3)

# --- Form and Control Event Handlers ---

# --- FIX: Changed from Add_Load to Add_Shown ---
# Add_Shown fires *after* the form is visible and the handle is created,
# preventing the BeginInvoke error.
$Global:mainForm.Add_Shown({
        # Logging will work here as the form is visible and the handle is created
        Write-Status "Application started."
        Write-Status "Looking for resources in: $Global:ScriptDirectory"
    
        # Set Window Icon
        if (Test-Path -Path $localIconPath) {
            try {
                $iconStream = New-Object System.IO.FileStream($localIconPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
                $Global:mainForm.Icon = New-Object System.Drawing.Icon($iconStream)
                $iconStream.Close()
                $iconStream.Dispose()
                Write-Status "Successfully set window icon."
            }
            catch {
                Write-Status "WARNING: Could not load icon.ico: $($_.Exception.Message)"
            }
        }
        else {
            Write-Status "No icon.ico found. Using default icon."
        }
    
        # Set logo status
        if ($logoLoaded) {
            Write-Status "Logo loaded successfully from local file."
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
            Write-Status "Found folder: $userDir"
            $effectsDir = Join-Path -Path $userDir -ChildPath $EffectsSubFolder
            Write-Status "Effect install directory set to: $effectsDir"
        }
        catch {
            Write-Status "ERROR: Could not read SignalRGB registry key on startup."
        }
    })

# Drag and Drop Event
$txtFilePath.Add_DragEnter({
        param($s, $e) # FIX: Renamed $sender to $s
        if ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
            $files = $e.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
        
            # Check if at least one of the dropped files is a valid type (.zip or .html)
            $hasValidFile = $files | Where-Object { 
                $ext = [System.IO.Path]::GetExtension($_).ToLower()
                $ext -eq ".zip" -or $ext -eq ".html"
            }
        
            if ($hasValidFile) {
                $e.Effect = [System.Windows.Forms.DragDropEffects]::Copy
            }
        }
    })

$txtFilePath.Add_DragDrop({
        param($s, $e) # FIX: Renamed $sender to $s
        $files = $e.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
    
        if ($files.Count -gt 0) {
            # *** MODIFICATION: Collect and display ALL valid files, separated by a semicolon ***
            $validFiles = @()
        
            foreach ($file in $files) {
                $ext = [System.IO.Path]::GetExtension($file).ToLower()
                if ($ext -eq ".zip" -or $ext -eq ".html") {
                    $validFiles += $file
                }
                else {
                    Write-Status "Skipped file due to invalid type: $file"
                }
            }
        
            if ($validFiles.Count -gt 0) {
                $fileList = $validFiles -join ";"
                $txtFilePath.Text = $fileList
                Write-Status "Files selected by drag-drop: $($validFiles.Count) valid files selected."
            
                # Since multiple files are selected, update the hint label to prompt the user to click Install
                $lblHint.Text = "Multiple files loaded. Click **'Install Effect'** to process them sequentially."
            }
            else {
                Write-Status "No valid .zip or .html files were dropped."
                $txtFilePath.Text = ""
            }
        }
    })

# Browse Button
$btnBrowse.Add_Click({
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.Filter = "Effect Files (*.zip, *.html)|*.zip;*.html|All Files (*.*)|*.*"
        $openFileDialog.Title = "Select Effect File(s)"
    
        # *** MODIFICATION: Set Multiselect property to allow multiple files ***
        $openFileDialog.Multiselect = $true
    
        if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            # The FileNames property returns an array of selected files.
            $selectedFiles = $openFileDialog.FileNames
        
            # Filter for valid file types and join them with the semicolon separator
            $validFiles = @()
            foreach ($file in $selectedFiles) {
                $ext = [System.IO.Path]::GetExtension($file).ToLower()
                if ($ext -eq ".zip" -or $ext -eq ".html") {
                    $validFiles += $file
                }
            }
        
            if ($validFiles.Count -gt 0) {
                $fileList = $validFiles -join ";"
                $txtFilePath.Text = $fileList
                Write-Status "Files selected via browse: $($validFiles.Count) valid files selected."
            
                # Update hint label to match batch installation mode
                if ($validFiles.Count -gt 1) {
                    $lblHint.Text = "Multiple files loaded. Click **'Install Effect'** to process them sequentially."
                }
                else {
                    $lblHint.Text = "Drag-and-drop a .zip or .html file here, or click Browse."
                }
            }
            else {
                [System.Windows.Forms.MessageBox]::Show("You must select at least one .zip or .html file.", "No Valid Selection", "OK", "Warning") | Out-Null
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
                Write-Status "One or more new/renamed effects were installed. Prompting user to restart $AppName."
                $restartResult = [System.Windows.Forms.MessageBox]::Show("All files processed.`n`n$AppName must be restarted to load the new effect(s). Restart now?", "Restart Required", "YesNo", "Question")
            
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
                        Write-Status "Please restart $AppName manually to see the new effect(s)."
                        [System.Windows.Forms.MessageBox]::Show("Could not stop $AppName (it may be running as Administrator).`n`nPlease restart it manually to see the new effect(s).", "Restart Failed", "OK", "Warning") | Out-Null
                    }
                }
                else {
                    Write-Status "User declined automatic restart. Manual restart required."
                    [System.Windows.Forms.MessageBox]::Show("Batch installation complete.`n`nManual restart of $AppName is required to see all new effect(s).", "Manual Restart Needed", "OK", "Information") | Out-Null
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
            $btnInstall.Text = "Install Effect"
        
            # Clear text box after batch installation is complete
            $txtFilePath.Text = ""
            $lblHint.Text = "Drag-and-drop a .zip or .html file here, or click Browse. Multiple files allowed."
        }
    })

# Uninstall Button
$btnUninstall.Add_Click({
        try {
            $userDir = (Get-ItemProperty -Path $RegKey -Name $RegValue).$RegValue
            $effectsDir = Join-Path -Path $userDir -ChildPath $EffectsSubFolder
            Show-UninstallWindow -EffectsBasePath $effectsDir
        }
        catch {
            Write-Status "ERROR: Could not get effects directory for uninstaller. $($_.Exception.Message)"
            [System.Windows.Forms.MessageBox]::Show("Could not find the SignalRGB Effects directory.`nHave you run SignalRGB at least once?", "Error", "OK", "Error") | Out-Null
        }
    })

# --- NEW: Disclaimer Button Event Handler ---
$btnDisclaimer.Add_Click({
        Show-DisclaimerWindow
    })

# --- Show the Form ---
Write-Status "GUI Initialized. Showing main window."
[System.Windows.Forms.Application]::Run($Global:mainForm)