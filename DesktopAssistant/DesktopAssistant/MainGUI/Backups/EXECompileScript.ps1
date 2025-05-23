Remove-Item -Path "C:\Users\jewilliams1\Documents\DesktopAssistant\MainGUI\DesktopAssistant.exe"

ps2exe.ps1 -inputFile "C:\Users\jewilliams1\Documents\DesktopAssistant\MainGUI\MainGUI.ps1" -outputFile "C:\Users\jewilliams1\Documents\DesktopAssistant\MainGUI\DesktopAssistant.exe" -noConsole -icon "C:\Users\jewilliams1\Documents\DesktopAssistant\MainGUI\AlliedLogo.ico"

Invoke-Item -Path "C:\Users\jewilliams1\Documents\DesktopAssistant\MainGUI\DesktopAssistant.exe"