# $bytes = [System.IO.File]::ReadAllBytes("C:\Users\josea\OneDrive\Documents\EffectInstaller\SignalRGB-Effect-Installer\rgbjunkielogo.png")
# $base64 = [System.Convert]::ToBase64String($bytes)
# $base64 | Out-File "encoded_logo.txt"

Invoke-ps2exe -inputFile "$PSScriptRoot\installer.ps1" `
              -outputFile "$PSScriptRoot\bin\SignalRGB_Installer.exe" `
              -icon "$PSScriptRoot\rgbjunkielogo2.ico" `
              -title "RGBJunkie Effect Installer" `
              -description "Utility to install custom effects and components to SignalRGB" `
              -company "José Miranda" `
              -version "1.6.0.0" `
              -noConsole