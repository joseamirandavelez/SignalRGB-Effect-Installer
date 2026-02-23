# $bytes = [System.IO.File]::ReadAllBytes(".\logo.png")
# $base64 = [System.Convert]::ToBase64String($bytes)
# $base64 | Out-File "encoded_logo.txt"

Invoke-ps2exe -inputFile ".\installer.ps1" `
              -outputFile ".\bin\SignalRGB_Installer.exe" `
              -icon ".\icon.ico" `
              -title "SRGB Effect Installer" `
              -description "Utility to install custom effects and components to SignalRGB" `
              -company "José Miranda" `
              -version "1.6.0.0" `
              -noConsole