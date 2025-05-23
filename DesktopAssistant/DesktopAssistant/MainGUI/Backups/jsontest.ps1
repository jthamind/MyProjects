$ConfigFile = Get-Content 'C:\Users\jewilliams1\OneDrive - Allied Solutions\Documents\Allied\Scripts\DesktopAssistant\MainGUI\config.json' | ConvertFrom-Json
$userProfilePath = [Environment]::GetEnvironmentVariable("USERPROFILE")
$ConfigFile.csvPath = $ConfigFile.csvPath.Replace("{USERPROFILE}", $userProfilePath)

Write-Host $ConfigFile.csvPath # This should print the expanded path