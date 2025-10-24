#Requires -Version 5.1

# --- Load Windows Forms Assembly ---
# Load these first to ensure all UI elements (even error popups) get modern styling
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Set high DPI awareness and visual styles
try {
    [System.Windows.Forms.Application]::SetHighDpiMode('SystemAware')
} catch {}
[System.Windows.Forms.Application]::EnableVisualStyles()


<#
.SYNOPSIS
    A GUI-based installer for SignalRGB effects.
    Handles .html, .png, and .zip files.
    Manages folder creation, conflict detection (file and <title> tag),
    overwrite/rename logic, and app restart.
    Includes an uninstaller and shortcut creation.

.DESCRIPTION
    This script provides a user-friendly graphical interface for
    installing custom effects into SignalRGB.

    Features:
    - GUI: A simple, clean interface.
    - Drag and Drop: Drop files directly onto the file box.
    - File Types: Supports .html, .png, and .zip files.
    - Conflict Resolution:
        - Checks if the destination folder exists.
        - Checks if any existing effect uses the same <title> tag.
        - Prompts user to Overwrite, Rename, or Cancel.
    - Rename Logic: Renames the folder, .html, .png, and the <title> tag inside the .html file.
    - Uninstaller: A separate window to view and delete installed effects.
    - Auto-Restart: Kills the SignalRGB process to force a refresh.
    - Auto-Shortcut: Prompts user to create shortcuts if they don't exist.
    - Logo/Icon: Loads local logo.png and icon.ico from the script's directory.
#>

# --- Self-Elevation block removed ---
# We no longer force administrator privileges.
# This will allow drag-and-drop to function correctly.
# The only side-effect is that stopping the SignalRGB process
# might fail if it is running with higher privileges.

# --- Script-Wide Variables ---
$Global:ScriptDirectory = $null
$Global:ScriptFullPath = $null
$Global:DesktopShortcutPath = $null
$Global:StartMenuShortcutPath = $null

try {
    # Set global paths. This is more reliable than $PSScriptRoot
    $Global:ScriptFullPath = $MyInvocation.MyCommand.Path
    $Global:ScriptDirectory = Split-Path -Path $Global:ScriptFullPath -Parent
    
    # Define shortcut paths
    $desktopPath = [Environment]::GetFolderPath("Desktop")
    $startMenuPath = [Environment]::GetFolderPath("Programs")
    $Global:DesktopShortcutPath = Join-Path -Path $desktopPath -ChildPath "Effect Installer.lnk"
    $Global:StartMenuShortcutPath = Join-Path -Path $startMenuPath -ChildPath "SignalRGB Tools\Effect Installer.lnk"
    
} catch {
    [System.Windows.Forms.MessageBox]::Show("CRITICAL ERROR: Could not determine script's own path. Shortcuts will fail. `n$($_.Exception.Message)", "Error", "OK", "Error")
    # Exit if we can't even find our own path
    return
}

# --- Registry and App Configuration ---
$RegKey = "HKCU:\Software\WhirlwindFX\SignalRgb"
$RegValue = "UserDirectory"
$AppName = "SignalRGB"
$AppExe = "SignalRGB.exe"
$EffectsSubFolder = "Effects" # The subfolder inside UserDirectory

# --- Load Remaining Assemblies ---
Add-Type -AssemblyName System.IO.Compression.FileSystem

# --- Helper Functions ---

function Log-Status {
    param (
        [string]$Message,
        [string]$Type = "INFO"
    )
    
    # This function assumes $txtStatus (the log box) is available in a scope it can access.
    # In PowerShell Forms, this usually works if the function is defined before the form is shown.
    if ($script:txtStatus) {
        $timestamp = Get-Date -Format "HH:mm:ss"
        $formattedMessage = "$timestamp - $Type - $Message`r`n"
        
        # Use BeginInvoke for thread-safe UI updates
        $script:txtStatus.BeginInvoke([Action[string]] {
            param ($msg)
            $script:txtStatus.AppendText($msg)
            $script:txtStatus.ScrollToCaret()
        }, $formattedMessage)
    } else {
        Write-Host "${Type}: $Message"
    }
}

function Show-CreateShortcutWindow {
    param (
        [string]$ScriptDirectory,
        [string]$IconPath,
        [bool]$CheckDesktop = $true,
        [bool]$CheckStartMenu = $true
    )

    $shortcutForm = New-Object System.Windows.Forms.Form
    $shortcutForm.Text = "Create Shortcuts"
    $shortcutForm.Size = New-Object System.Drawing.Size(350, 200)
    $shortcutForm.FormBorderStyle = 'FixedDialog'
    $shortcutForm.MaximizeBox = $false
    $shortcutForm.MinimizeBox = $false
    $shortcutForm.StartPosition = 'CenterParent'
    if ($Global:mainForm.Icon) { $shortcutForm.Icon = $Global:mainForm.Icon }

    $layout = New-Object System.Windows.Forms.TableLayoutPanel
    $layout.Dock = 'Fill'
    $layout.Padding = 10
    $layout.ColumnCount = 1
    $layout.RowCount = 4
    $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
    $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
    $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
    $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null # Spacer
    $shortcutForm.Controls.Add($layout)

    $lblInfo = New-Object System.Windows.Forms.Label
    $lblInfo.Text = "Create shortcuts for the Effect Installer:"
    $lblInfo.Dock = 'Fill'
    $layout.Controls.Add($lblInfo, 0, 0)

    $chkDesktop = New-Object System.Windows.Forms.CheckBox
    $chkDesktop.Text = "On the Desktop"
    $chkDesktop.Checked = $CheckDesktop
    $chkDesktop.Dock = 'Fill'
    $layout.Controls.Add($chkDesktop, 0, 1)

    $chkStartMenu = New-Object System.Windows.Forms.CheckBox
    $chkStartMenu.Text = "In the Start Menu (under 'SignalRGB Tools')"
    $chkStartMenu.Checked = $CheckStartMenu
    $chkStartMenu.Dock = 'Fill'
    $layout.Controls.Add($chkStartMenu, 0, 2)

    $btnCreate = New-Object System.Windows.Forms.Button
    $btnCreate.Text = "Create"
    $btnCreate.Dock = 'Fill'
    $btnCreate.Anchor = 'Top'
    $btnCreate.Height = 30
    $btnCreate.Margin = [System.Windows.Forms.Padding]::new(0, 10, 0, 0)
    $layout.Controls.Add($btnCreate, 0, 3)

    $btnCreate.Add_Click({
        try {
            # WScript.Shell is the object that creates shortcuts
            $wsShell = New-Object -ComObject WScript.Shell

            $targetFile = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
            $arguments = "-WindowStyle Hidden -ExecutionPolicy Bypass -File ""$Global:ScriptFullPath"""

            if ($chkDesktop.Checked) {
                try {
                    Log-Status "Creating Desktop shortcut..."
                    $shortcut = $wsShell.CreateShortcut($Global:DesktopShortcutPath)
                    $shortcut.TargetPath = $targetFile
                    $shortcut.Arguments = $arguments
                    $shortcut.WorkingDirectory = $Global:ScriptDirectory
                    if (Test-Path -Path $IconPath) {
                        $shortcut.IconLocation = $IconPath
                    }
                    $shortcut.Save()
                    Log-Status "Desktop shortcut created."
                } catch {
                    Log-Status "ERROR creating Desktop shortcut: $($_.Exception.Message)"
                }
            }
            
            if ($chkStartMenu.Checked) {
                try {
                    Log-Status "Creating Start Menu shortcut..."
                    # Ensure the folder exists
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
                    Log-Status "Start Menu shortcut created."
                } catch {
                    Log-Status "ERROR creating Start Menu shortcut: $($_.Exception.Message)"
                }
            }
            $shortcutForm.Close()
        } catch {
            Log-Status "ERROR: Could not create shortcuts. $($_.Exception.Message)"
            [System.Windows.Forms.MessageBox]::Show("Error creating shortcut: $($_.Exception.Message)", "Error", "OK", "Error") | Out-Null
        }
    })

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
    if ($Global:mainForm.Icon) { $uninstallForm.Icon = $Global:mainForm.Icon }

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
        Log-Status "Refreshing effect list from: $EffectsBasePath"
        if (-not (Test-Path -Path $EffectsBasePath)) {
            Log-Status "Effects folder does not exist. Nothing to list."
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
            Log-Status "Found $($uniqueFolders.Count) effects."
        } catch {
            Log-Status "ERROR scanning for effects: $($_.Exception.Message)"
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
        } catch {
            Log-Status "Could not read current 'always' key. Will not update registry on delete."
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
            Log-Status "Active effect '$currentAlwaysTitle' is scheduled for deletion."
        }
        # --- End new logic block ---

        # --- MOVED: Set new active effect *before* deleting files ---
        if ($activeEffectFolderWasDeleted) {
            Log-Status "Updating active effect registry keys..."
            $remainingEffectFolders = @($originalEffectFolders | Where-Object { $_ -notin $selectedItems })
            
            if ($remainingEffectFolders.Count -eq 0) {
                # No effects left, set to empty
                Log-Status "All effects deleted. Setting active effect to empty."
                Set-ActiveEffectRegistryKeys -NewEffectTitle ""
            } else {
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
                
                Log-Status "Setting new active effect to: '$newEffectTitle'"
                Set-ActiveEffectRegistryKeys -NewEffectTitle $newEffectTitle
            }
        }
        # --- End of moved block ---

        foreach ($itemName in $selectedItems) {
            $effectFolder = Join-Path -Path $EffectsBasePath -ChildPath $itemName
            Log-Status "Deleting effect: $itemName"
            try {
                # Find all html and png files in that folder
                # Fix: Ensure $filesToDelete is always an array by using @()
                $filesToDelete = @(Get-ChildItem -Path $effectFolder -Filter "*.html")
                $filesToDelete += @(Get-ChildItem -Path $effectFolder -Filter "*.png")

                foreach ($file in $filesToDelete) {
                    Log-Status "Deleting file: $($file.Name)"
                    Remove-Item -Path $file.FullName -Force
                }
                
                # Check if folder is now empty
                if ((Get-ChildItem -Path $effectFolder -Force | Measure-Object).Count -eq 0) {
                    Log-Status "Folder '$itemName' is empty, deleting it."
                    Remove-Item -Path $effectFolder -Recurse -Force
                } else {
                    Log-Status "Folder '$itemName' is not empty, will not delete folder."
                }
                Log-Status "Successfully deleted effect '$itemName'."
            } catch {
                Log-Status "ERROR deleting '$itemName': $($_.Exception.Message)"
                [System.Windows.Forms.MessageBox]::Show("Error deleting '$itemName':`n$($_.Exception.Message)", "Error", "OK", "Error") | Out-Null
            }
        }
        
        Log-Status "Deletion complete."
        
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
        } else {
            Log-Status "WARNING: Could not find <title> tag in $HtmlFilePath. Using filename as fallback."
            return $fallbackTitle
        }
    } catch {
        Log-Status "ERROR: Could not read $HtmlFilePath to find title. Using filename as fallback. $($_.Exception.Message)"
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
    
    # This custom dialog is necessary because MessageBox doesn't support 3 custom buttons.
    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = "Conflict Detected"
    $dialog.Size = New-Object System.Drawing.Size(400, 150)
    $dialog.FormBorderStyle = 'FixedDialog'
    $dialog.MaximizeBox = $false
    $dialog.MinimizeBox = $false
    $dialog.StartPosition = 'CenterParent'
    if ($Global:mainForm.Icon) { $dialog.Icon = $Global:mainForm.Icon }
    
    $lblMessage = New-Object System.Windows.Forms.Label
    $lblMessage.Text = $Message
    $lblMessage.Dock = 'Fill'
    $lblMessage.TextAlign = 'MiddleCenter'
    $dialog.Controls.Add($lblMessage)

    $btnOverwrite = New-Object System.Windows.Forms.Button
    $btnOverwrite.Text = "Overwrite"
    $btnOverwrite.DialogResult = 'OK'
    $btnOverwrite.Location = New-Object System.Drawing.Point(30, 80)
    $dialog.Controls.Add($btnOverwrite)
    
    $btnRename = New-Object System.Windows.Forms.Button
    $btnRename.Text = "Rename"
    $btnRename.DialogResult = 'Retry'
    $btnRename.Location = New-Object System.Drawing.Point(150, 80)
    $dialog.Controls.Add($btnRename)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.DialogResult = 'Cancel'
    $btnCancel.Location = New-Object System.Drawing.Point(270, 80)
    $dialog.Controls.Add($btnCancel)

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
    if ($Global:mainForm.Icon) { $dialog.Icon = $Global:mainForm.Icon }

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
            Log-Status "ERROR: New name is invalid or empty."
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
            Log-Status "Successfully updated <title> in $HtmlFilePath to '$NewTitle'."
            return $true
        } else {
            Log-Status "WARNING: Could not find <title> tag in $HtmlFilePath to update it."
            return $false
        }
    } catch {
        Log-Status "ERROR: Could not read or write to $HtmlFilePath to update title. $($_.Exception.Message)"
        return $false
    }
}

function Set-ActiveEffectRegistryKeys {
    param(
        [string]$NewEffectTitle
    )
    
    $keyPath = "HKCU:\Software\WhirlwindFX\SignalRgb\effects\selected"
    Log-Status "Setting active effect registry keys..."
    
    try {
        # Ensure the path exists
        if (-not (Test-Path -Path $keyPath)) {
            New-Item -Path $keyPath -Force | Out-Null
            Log-Status "Created registry key path: $keyPath"
        }
        
        # 1. Get the current 'always' value
        $currentAlways = Get-ItemProperty -Path $keyPath -Name "always" -ErrorAction SilentlyContinue
        
        if ($currentAlways) {
            # 2. Set 'previous' key to the 'always' value
            Set-ItemProperty -Path $keyPath -Name "previous" -Value $currentAlways.always
            Log-Status "Set 'previous' key to: $($currentAlways.always)"
        } else {
            Log-Status "No existing 'always' key found. Skipping 'previous' key set."
        }
        
        # 3. Set 'always' key to the new effect title
        Set-ItemProperty -Path $keyPath -Name "always" -Value $NewEffectTitle
        Log-Status "Set 'always' key to: $NewEffectTitle"
        
    } catch {
        Log-Status "ERROR: Could not update registry keys. $($_.Exception.Message)"
    }
}

# --- REMOVED Swap-ActiveEffectRegistryKeys function ---

# --- Main Installation Logic ---

function Start-Installation {
    param (
        [string]$FilePath
    )
    
    Log-Status "Starting installation for: $FilePath"
    
    # 1. Get SignalRGB User Directory from Registry
    $userDir = $null
    try {
        Log-Status "Reading registry key: $RegKey"
        $userDir = (Get-ItemProperty -Path $RegKey -Name $RegValue).$RegValue
        if ([string]::IsNullOrWhiteSpace($userDir) -or -not (Test-Path -Path $userDir)) {
            Log-Status "ERROR: Registry key found, but folder is invalid or missing: $userDir"
            [System.Windows.Forms.MessageBox]::Show("Error: SignalRGB UserDirectory not found or is invalid.`nChecked: $userDir", "Registry Error", "OK", "Error") | Out-Null
            return
        }
        Log-Status "Found folder: $userDir"
    } catch {
        Log-Status "ERROR: Could not read registry key: $RegKey. $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Error: Could not read registry key for $AppName.`n$($_.Exception.Message)", "Registry Error", "OK", "Error") | Out-Null
        return
    }

    # --- Define the correct base install folder ---
    $installBasePath = Join-Path -Path $userDir -ChildPath $EffectsSubFolder
    
    # --- Ensure the 'Effects' folder exists ---
    if (-not (Test-Path -Path $installBasePath)) {
        Log-Status "'$EffectsSubFolder' folder not found. Creating it..."
        try {
            New-Item -Path $installBasePath -ItemType Directory -Force | Out-Null
            Log-Status "Created folder: $installBasePath"
        } catch {
            Log-Status "ERROR: Could not create folder: $installBasePath. $($_.Exception.Message)"
            [System.Windows.Forms.MessageBox]::Show("Error: Could not create folder:`n$installBasePath", "Error", "OK", "Error") | Out-Null
            return
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
            Log-Status "Zip file detected. Extracting..."
            $tempExtractFolder = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ([System.IO.Path]::GetRandomFileName())
            [System.IO.Compression.ZipFile]::ExtractToDirectory($FilePath, $tempExtractFolder)
            Log-Status "Extracted to: $tempExtractFolder"
            
            # Find the html and png files inside
            $sourceHtmlFile = (Get-ChildItem -Path $tempExtractFolder -Recurse -Filter "*.html" | Select-Object -First 1).FullName
            if (-not $sourceHtmlFile) {
                Log-Status "ERROR: No .html file found in the zip archive."
                [System.Windows.Forms.MessageBox]::Show("Error: No .html file found in the zip archive.", "Zip Error", "OK", "Error") | Out-Null
                return
            }
            
            $effectName = [System.IO.Path]::GetFileNameWithoutExtension($sourceHtmlFile)
            $sourcePngFile = (Get-ChildItem -Path (Split-Path $sourceHtmlFile) -Filter "$effectName.png" | Select-Object -First 1).FullName
            if (-not $sourcePngFile) {
                Log-Status "No matching .png found in zip. This might be ok."
            }
            
        } elseif ($extension -eq ".html") {
            $sourceHtmlFile = $FilePath
            $effectName = [System.IO.Path]::GetFileNameWithoutExtension($sourceHtmlFile)
            $sourcePngFile = Join-Path -Path (Split-Path $FilePath) -ChildPath "$effectName.png"
            if (-not (Test-Path -Path $sourcePngFile)) {
                $sourcePngFile = $null # Don't try to copy a non-existent file
                Log-Status "No matching .png found for $effectName. This might be ok."
            }
            
        } else {
            Log-Status "ERROR: Invalid file type. Please select a .zip or .html file."
            [System.Windows.Forms.MessageBox]::Show("Invalid file type. Please select a .zip or .html file.", "Invalid File", "OK", "Error") | Out-Null
            return
        }

        if (-not $sourceHtmlFile) {
            Log-Status "ERROR: Could not determine source .html file."
            return
        }

        Log-Status "Effect Name (from file): $effectName"
        Log-Status "Source HTML: $sourceHtmlFile"
        Log-Status "Source PNG: $sourcePngFile"

        # 3. Conflict Detection Loop
        $installConfirmed = $false
        $currentEffectName = $effectName
        $currentHtmlFile = $sourceHtmlFile
        $currentPngFile = $sourcePngFile
        
        while (-not $installConfirmed) {
            $destFolder = Join-Path -Path $installBasePath -ChildPath $currentEffectName
            $destHtmlFile = Join-Path -Path $destFolder -ChildPath ($currentEffectName + ".html")
            
            # Get title from HTML
            $currentEffectTitle = Get-EffectTitleFromHtml -HtmlFilePath $currentHtmlFile
            # No fallback needed here, Get-EffectTitleFromHtml will return filename if title is missing
            Log-Status "Effect Title (from HTML): $currentEffectTitle"
            
            # Check for conflicts
            $folderExists = Test-Path -Path $destFolder
            $titleConflictFile = Find-EffectTitleConflict -NewEffectTitle $currentEffectTitle -EffectsBasePath $installBasePath
            
            if ($folderExists -or $titleConflictFile) {
                # Conflict found!
                $conflictMessage = "A conflict was found.`n`n"
                if ($folderExists) { $conflictMessage += "- Folder '$currentEffectName' already exists.`n" }
                if ($titleConflictFile) { $conflictMessage += "- Title '$currentEffectTitle' is already used by `n  $titleConflictFile`n" }
                $conflictMessage += "`nWould you like to Overwrite, Rename, or Cancel?"
                
                Log-Status "Conflict detected: $conflictMessage"
                $userChoice = Show-ConflictDialog -Message $conflictMessage
                
                if ($userChoice -eq 'Overwrite') {
                    Log-Status "User chose to Overwrite."
                    $installConfirmed = $true
                    
                } elseif ($userChoice -eq 'Rename') {
                    Log-Status "User chose to Rename."
                    $newName = Show-RenameDialog -OldName $currentEffectName
                    
                    if ([string]::IsNullOrWhiteSpace($newName)) {
                        Log-Status "Rename cancelled by user."
                        # Loop will repeat, user will be prompted again
                        return
                    }
                    
                    if ($newName.Equals($currentEffectName, [StringComparison]::OrdinalIgnoreCase)) {
                        Log-Status "New name is the same as the old name. No changes made."
                        # Loop will repeat
                        continue
                    }

                    Log-Status "New name selected: $newName"
                    $currentEffectName = $newName
                    
                    # Update HTML <title> tag
                    Set-EffectTitleInHtml -HtmlFilePath $currentHtmlFile -NewTitle $newName | Out-Null
                    # Note: We will re-read this new title at the start of the next loop.
                    
                } else {
                    # User chose Cancel
                    Log-Status "Installation cancelled by user."
                    return
                }
                
            } else {
                # No conflict
                Log-Status "No conflicts found. Proceeding with installation."
                $installConfirmed = $true
            }
        } # End conflict loop

        # 4. Perform Installation
        Log-Status "Installing '$currentEffectName' to $destFolder"
        
        # Ensure destination folder exists
        if (-not (Test-Path -Path $destFolder)) {
            New-Item -Path $destFolder -ItemType Directory -Force | Out-Null
        }
        
        # Define final destination file paths
        $finalHtmlPath = Join-Path -Path $destFolder -ChildPath ($currentEffectName + ".html")
        $finalPngPath = Join-Path -Path $destFolder -ChildPath ($currentEffectName + ".png")

        # Copy HTML
        Copy-Item -Path $currentHtmlFile -Destination $finalHtmlPath -Force
        Log-Status "Copied HTML to: $finalHtmlPath"
        
        # Copy PNG if it exists
        if ($currentPngFile -and (Test-Path -Path $currentPngFile)) {
            Copy-Item -Path $currentPngFile -Destination $finalPngPath -Force
            Log-Status "Copied PNG to: $finalPngPath"
        }
        
        # 5. Set Registry Keys
        $finalEffectTitle = Get-EffectTitleFromHtml -HtmlFilePath $finalHtmlPath
        # No fallback needed, function handles it
        
        Set-ActiveEffectRegistryKeys -NewEffectTitle $finalEffectTitle
        
        # 6. Stop SignalRGB to force it to restart and reload effects
        Log-Status "Attempting to stop $AppName to force reload..."
        try {
            $process = Get-Process -Name $AppName -ErrorAction Stop
            if ($process) {
                Stop-Process -Name $AppName -Force
                Log-Status "$AppName process stopped. It should restart automatically."
            } else {
                Log-Status "$AppName is not running."
            }
        } catch {
            Log-Status "WARNING: Could not stop $AppName. It may be running as Administrator."
            Log-Status "Please restart $AppName manually to see the new effect."
        }
        
        Log-Status "SUCCESS: Effect '$currentEffectName' installed."
        [System.Windows.Forms.MessageBox]::Show("Effect '$currentEffectName' installed successfully.`n`n$AppName will now restart to load the new effect.", "Success", "OK", "Information") | Out-Null
        
    } catch {
        Log-Status "FATAL ERROR during installation: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("A fatal error occurred during installation:`n$($_.Exception.Message)", "Fatal Error", "OK", "Error") | Out-Null
    } finally {
        # 7. Clean up temp folder if we created one
        if ($tempExtractFolder -and (Test-Path -Path $tempExtractFolder)) {
            Log-Status "Cleaning up temporary folder: $tempExtractFolder"
            Remove-Item -Path $tempExtractFolder -Recurse -Force
        }
    }
}

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
    } catch {
        Log-Status "ERROR: Could not load logo.png: $($_.Exception.Message)"
    }
} else {
    Log-Status "No logo.png found. Hiding logo area."
}

# --- Row 1: File Input ---
$fileInputLayout = New-Object System.Windows.Forms.TableLayoutPanel
$fileInputLayout.Dock = 'Fill'
$fileInputLayout.AutoSize = $true
$fileInputLayout.ColumnCount = 3
$fileInputLayout.RowCount = 2
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
$lblHint.Text = "Drag & drop a .zip or .html file here, or click Browse."
$lblHint.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
$lblHint.Dock = 'Fill'
$lblHint.Margin = [System.Windows.Forms.Padding]::new(0, 0, 0, 5)
$fileInputLayout.Controls.Add($lblHint, 1, 1) # Span 2
$fileInputLayout.SetColumnSpan($lblHint, 2)

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

# Spacer
$pnlSpacer = New-Object System.Windows.Forms.Panel
$pnlSpacer.Dock = 'Fill'
$pnlSpacer.Height = 30 # Added this line to fix row height
$buttonLayout.Controls.Add($pnlSpacer, 2, 0)

# --- Row 3: Status Box ---
$script:txtStatus = New-Object System.Windows.Forms.TextBox
$script:txtStatus.Multiline = $true
$script:txtStatus.ReadOnly = $true
$script:txtStatus.ScrollBars = 'Vertical'
$script:txtStatus.Dock = 'Fill'
$script:txtStatus.Font = New-Object System.Drawing.Font("Consolas", 9)
$mainLayout.Controls.Add($script:txtStatus, 0, 3)

# --- Form and Control Event Handlers ---

# Form Load
$Global:mainForm.Add_Load({
    Log-Status "Application started."
    Log-Status "Looking for resources in: $Global:ScriptDirectory"
    
    # Set Window Icon
    if (Test-Path -Path $localIconPath) {
        try {
            $iconStream = New-Object System.IO.FileStream($localIconPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
            $Global:mainForm.Icon = New-Object System.Drawing.Icon($iconStream)
            $iconStream.Close()
            $iconStream.Dispose()
            Log-Status "Successfully set window icon."
        } catch {
            Log-Status "WARNING: Could not load icon.ico: $($_.Exception.Message)"
        }
    } else {
        Log-Status "No icon.ico found. Using default icon."
    }
    
    # Set logo status
    if ($logoLoaded) {
        Log-Status "Logo loaded successfully from local file."
    }

    # --- NEW: Check for existing shortcuts ---
    Log-Status "Checking for Desktop shortcut at: $Global:DesktopShortcutPath"
    # -PathType Leaf ensures we are looking for a file, not a directory or a recycled item
    $desktopExists = (Test-Path -Path $Global:DesktopShortcutPath -PathType Leaf)
    Log-Status "Desktop shortcut exists: $desktopExists"
    
    Log-Status "Checking for Start Menu shortcut at: $Global:StartMenuShortcutPath"
    $startMenuExists = (Test-Path -Path $Global:StartMenuShortcutPath -PathType Leaf)
    Log-Status "Start Menu shortcut exists: $startMenuExists"

    if (-not $desktopExists -or -not $startMenuExists) {
        Log-Status "One or more shortcuts are missing. Prompting user."
        $promptResult = [System.Windows.Forms.MessageBox]::Show("Would you like to create a shortcut for this installer on your Desktop and/or Start Menu?", "Create Shortcut?", "YesNo", "Question")
        
        if ($promptResult -eq 'Yes') {
            # Call the shortcut window, pre-checking the boxes for the missing ones
            Show-CreateShortcutWindow -ScriptDirectory $Global:ScriptDirectory -IconPath $localIconPath -CheckDesktop (-not $desktopExists) -CheckStartMenu (-not $startMenuExists)
        }
    }

    # Initial Log of Registry Key
    try {
        Log-Status "Reading registry key: $RegKey"
        $userDir = (Get-ItemProperty -Path $RegKey -Name $RegValue).$RegValue
        Log-Status "Found folder: $userDir"
        $effectsDir = Join-Path -Path $userDir -ChildPath $EffectsSubFolder
        Log-Status "Effect install directory set to: $effectsDir"
    } catch {
        Log-Status "ERROR: Could not read SignalRGB registry key on startup."
    }
})

# Drag and Drop Event
$txtFilePath.Add_DragEnter({
    param($sender, $e)
    if ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
        $e.Effect = [System.Windows.Forms.DragDropEffects]::Copy
    }
})

$txtFilePath.Add_DragDrop({
    param($sender, $e)
    $files = $e.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
    if ($files.Count -gt 0) {
        $firstFile = $files[0]
        $ext = [System.IO.Path]::GetExtension($firstFile).ToLower()
        if ($ext -eq ".zip" -or $ext -eq ".html") {
             $txtFilePath.Text = $firstFile
             Log-Status "File selected by drag-drop: $firstFile"
        } else {
            Log-Status "Invalid file type dropped. Please drop a .zip or .html file."
        }
    }
})

# Browse Button
$btnBrowse.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "Effect Files (*.zip, *.html)|*.zip;*.html|All Files (*.*)|*.*"
    $openFileDialog.Title = "Select an Effect File"
    
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtFilePath.Text = $openFileDialog.FileName
        Log-Status "File selected via browse: $($openFileDialog.FileName)"
    }
})

# Install Button
$btnInstall.Add_Click({
    if ([string]::IsNullOrWhiteSpace($txtFilePath.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Please select a file to install first.", "No File Selected", "OK", "Warning") | Out-Null
        return
    }
    
    if (-not (Test-Path -Path $txtFilePath.Text)) {
        [System.Windows.Forms.MessageBox]::Show("The selected file does not exist. Please check the path.", "File Not Found", "OK", "Error") | Out-Null
        return
    }
    
    # Disable button during install
    $btnInstall.Enabled = $false
    $btnInstall.Text = "Installing..."
    $Global:mainForm.Refresh()
    
    try {
        Start-Installation -FilePath $txtFilePath.Text
    } catch {
        Log-Status "Unhandled exception in Start-Installation: $($_.Exception.Message)"
    } finally {
        # Re-enable button
        $btnInstall.Enabled = $true
        $btnInstall.Text = "Install Effect"
    }
})

# Uninstall Button
$btnUninstall.Add_Click({
    try {
        $userDir = (Get-ItemProperty -Path $RegKey -Name $RegValue).$RegValue
        $effectsDir = Join-Path -Path $userDir -ChildPath $EffectsSubFolder
        Show-UninstallWindow -EffectsBasePath $effectsDir
    } catch {
        Log-Status "ERROR: Could not get effects directory for uninstaller. $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Could not find the SignalRGB Effects directory.`nHave you run SignalRGB at least once?", "Error", "OK", "Error") | Out-Null
    }
})

# --- Show the Form ---
Log-Status "GUI Initialized. Showing main window."
[System.Windows.Forms.Application]::Run($Global:mainForm)
