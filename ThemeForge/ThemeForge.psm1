# Paths to settings file and custom themes
$script:themesPath = Join-Path -Path $Env:LOCALAPPDATA -ChildPath "Packages\Microsoft.WindowsTerminal_8wekyb3d8bbwe\LocalState"
$script:settingsPath = Join-Path -Path $script:themesPath -ChildPath "settings.json"

function Show-ThemeForge {
    $helpText = @"
=============================================================================================
                                  (                           
  *   )    )                      )\ )                        
` )  /( ( /(    (     )      (   (()/(      (    (  (     (   
 ( )(_)))\())  ))\   (      ))\   /(_)) (   )(   )\))(   ))\  
(_(_())((_)\  /((_)  )\  ' /((_) (_))_| )\ (()\ ((_))\  /((_) 
|_   _|| |(_)(_))  _((_)) (_))   | |_  ((_) ((_) (()(_)(_))   
  | |  | ' \ / -_)| '  \()/ -_)  | __|/ _ \| '_|/ _` | / -_)  
  |_|  |_||_|\___||_|_|_| \___|  |_|  \___/|_|  \__, | \___|  
                                                |___/         
Author: Jeremiah Williams

This module allows you to create, manage, and share themes for Windows Terminal. These 
settings are stored in Terminal's settings.json file which you can access by pressing ctrl+, 
while in Terminal. This module works by creating a copy of the current theme's settings in a 
new directory called .\<theme_name>\settings.json. If the theme has a background image, 
then that file will also be copied to the new directory. 

Tab autocomplete is supported for the below commands with an asterisk(*) after the parameter.

- New-ThemeForge    (ntf) <Name>	Create a new theme
- Get-ThemeForge    (gtf)		Get a list of current and available themes
- Change-ThemeForge (ctf) <Name> *     	Change terminal theme
- Update-ThemeForge (utf) <Name> *     	Update theme with current settings
- Remove-ThemeForge (rtf) <Name> *     	Remove theme
- Export-ThemeForge (etf) <Name> *     	Share theme to T:\Software\WindowsTerminalThemes
- Show-ThemeForge   (stf) <Name>	Show this help utility

Check out these websites for a wide range of free, custom themes:

https://windowsterminalthemes.dev
https://terminalsplash.com
https://github.com/rjcarneiro/windows-terminals

=============================================================================================
"@
    Write-Host $helpText
}
function Change-ThemeForge {
    param (
        [Parameter(Mandatory, HelpMessage = "Specify the name of the theme to change to.")]
	[ValidateNotNullOrEmpty()]
        [string]$Name
    )

    # Verify the theme directory exists
    $themePath = Join-Path -Path $script:themesPath -ChildPath $Name
    if (-not (Test-Path -Path $themePath)) {
        Write-Error "'$Name' theme not found. Ensure the directory exists at $themePath."
        return
    }

    # Verify the theme's settings.json file exists
    $themeSettingsPath = Join-Path -Path $themePath -ChildPath "settings.json"
    if (-not (Test-Path -Path $themeSettingsPath)) {
        Write-Error "settings.json file not found for theme '$Name'."
        return
    }

    # Copy the theme's settings.json to the active settings.json
    Copy-Item -Path $themeSettingsPath -Destination $script:settingsPath -Force

    Write-Host "'$Name' is now active." -ForegroundColor Green
}
function Get-ThemeForge {
    # Read the settings.json file to get the active colorScheme
    $settings = Get-Content -Path $settingsPath -Raw | ConvertFrom-Json
    $activeTheme = $settings.profiles.defaults.colorScheme

    # Get the themes in the directory
    $themes = Get-ChildItem -Path $script:themesPath -Directory | Select-Object Name

    # Display the active theme and available themes
    Write-Host "===================================================`n"
    Write-Host "|Active|" -ForegroundColor Green
	Write-Host "$activeTheme`n"

    Write-Host "|Available|" -ForegroundColor Yellow
    $themes | ForEach-Object {
        if ($_.Name -eq $activeTheme) {
            # Don't display the active theme again in the available list
            $_.Name
        } else {
            $_.Name
        }
    }
    Write-Host "`n==================================================="
}
function New-ThemeForge {
    param (
	[Parameter(Mandatory, HelpMessage = "Specify the name of the new theme.")]
	[ValidateNotNullOrEmpty()]
        [string]$Name
    )

    # If no theme name is passed, default to the colorScheme from settings.json
    if (-not $Name) {
        $settings = Get-Content -Path $script:settingsPath -Raw | ConvertFrom-Json
        $Name = $settings.profiles.defaults.colorScheme
    }

    # Check if the theme folder already exists
    $themePath = Join-Path -Path $script:themesPath -ChildPath $Name
    if (Test-Path -Path $themePath) {
        Write-Host "'$Name' theme already exists." -ForegroundColor Red
        return
    }

    # Create the new theme folder
    $newDirectoryOutput = New-Item -Path $themePath -ItemType Directory

    # Optionally: copy over settings.json or create a default structure for the new theme
    $templateSettings = Join-Path -Path $script:themesPath -ChildPath "settings.json"
    if (Test-Path $templateSettings) {
        Copy-Item -Path $templateSettings -Destination $themePath
    }

    # Check if the profiles>defaults section contains a background image
    $settings = Get-Content -Path $script:settingsPath -Raw | ConvertFrom-Json
    $backgroundImagePath = $settings.profiles.defaults.backgroundImage

    if ($backgroundImagePath) {
        # Check if the image file exists
        if (Test-Path -Path $backgroundImagePath) {
            # Define the new path for the image in the theme folder
            $newBackgroundPath = Join-Path -Path $themePath -ChildPath (Split-Path -Leaf $backgroundImagePath)

            # Copy the image to the new theme folder
            Copy-Item -Path $backgroundImagePath -Destination $newBackgroundPath
        } else {
            Write-Host "Background image '$backgroundImagePath' not found." -ForegroundColor Yellow
        }
    }

    Write-Host "'$Name' theme created at $themePath" -ForegroundColor Green
}
function Remove-ThemeForge {
    param (
        [Parameter(Mandatory, HelpMessage = "Specify the name of the theme to remove.")]
	[ValidateNotNullOrEmpty()]
        [string]$Name
    )

    # Define the path to the theme folder
    $themePath = Join-Path -Path $script:themesPath -ChildPath $Name

    # Check if the theme folder exists
    if (Test-Path -Path $themePath) {
        # Remove the theme folder and all its contents
        Remove-Item -Path $themePath -Recurse -Force
        Write-Host "'$Name' theme has been removed successfully." -ForegroundColor Green
    } else {
        Write-Host "'$Name' theme does not exist." -ForegroundColor Red
    }
}
function Export-ThemeForge {
    param (
	[Parameter(Mandatory, HelpMessage = "Specify the name of the theme to export.")]
	[ValidateNotNullOrEmpty()]
        [string]$Name
    )

    # Define the source and destination paths
    $sourcePath = Join-Path -Path $Env:LOCALAPPDATA -ChildPath "Packages\Microsoft.WindowsTerminal_8wekyb3d8bbwe\LocalState\$Name"
    $destinationPath = "T:\Software\WindowsTerminalThemes\$Name"

    # Check if the theme folder exists in the source path
    if (Test-Path -Path $sourcePath) {
        # Check if the destination folder already exists
        if (Test-Path -Path $destinationPath) {
            Write-Host "The theme '$Name' already exists at $destinationPath." -ForegroundColor Yellow
        } else {
            # Create the destination folder if it doesn't exist
            New-Item -ItemType Directory -Path $destinationPath

            # Copy the theme folder to the destination
            Copy-Item -Path $sourcePath\* -Destination $destinationPath -Recurse -Force

            # Read the settings.json file from the source folder
            $settingsFilePath = Join-Path $sourcePath "settings.json"
            if (Test-Path -Path $settingsFilePath) {
                # Read the entire settings.json file
                $settingsJson = Get-Content -Path $settingsFilePath | ConvertFrom-Json

                # Create a new object with only the 'profiles.defaults' and 'schemes' sections
                $newSettings = @{
                    profiles = @{
                        defaults = $settingsJson.profiles.defaults
                    }
                    schemes = $settingsJson.schemes
                }

                # Save the new settings to the destination
                $newSettings | ConvertTo-Json -Depth 10 | Set-Content -Path (Join-Path $destinationPath "settings.json")

                Write-Host "'$Name' theme has been shared to $destinationPath" -ForegroundColor Green
            } else {
                Write-Host "The 'settings.json' file was not found in $sourcePath" -ForegroundColor Red
            }
        }
    } else {
        Write-Host "The '$Name' theme was not found in $sourcePath" -ForegroundColor Red
    }
}
function Update-ThemeForge {
    param (
	[Parameter(Mandatory, HelpMessage = "Specify the name of the theme to update.")]
	[ValidateNotNullOrEmpty()]
        [string]$Name
    )

    $themePath = Join-Path -Path $script:themesPath -ChildPath $Name

    # Check if the theme folder exists
    if (-not (Test-Path -Path $themePath)) {
        Write-Host "'$Name' theme does not exist." -ForegroundColor Red
        return
    }

    # Copy the current settings.json file to the theme folder
    if (Test-Path -Path $script:settingsPath) {
        Copy-Item -Path $script:settingsPath -Destination $themePath -Force
        Write-Host "settings.json has been copied to '$themePath'" -ForegroundColor Green
    } else {
        Write-Host "The settings.json file was not found." -ForegroundColor Yellow
    }

    # Check if the profiles>defaults section contains a background image
    $settings = Get-Content -Path $settingsPath -Raw | ConvertFrom-Json
    $backgroundImagePath = $settings.profiles.defaults.backgroundImage

    if ($backgroundImagePath) {
        # Check if the image file exists
        if (Test-Path -Path $backgroundImagePath) {
            # Define the new path for the image in the theme folder
            $newBackgroundPath = Join-Path -Path $themePath -ChildPath (Split-Path -Leaf $backgroundImagePath)

            # Copy the image to the new theme folder
            Copy-Item -Path $backgroundImagePath -Destination $newBackgroundPath -Force
            Write-Host "Background image has been copied to '$newBackgroundPath'" -ForegroundColor Green
        } else {
            Write-Host "Background image '$backgroundImagePath' not found." -ForegroundColor Yellow
        }
    }
}
# Register tab autocomplete functionality for change, export, update, and remove functions
Register-ArgumentCompleter -CommandName Change-ThemeForge -ParameterName Name -ScriptBlock {
    param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameters)

    # Get theme names from the directory
    Get-ChildItem -Path $script:themesPath -Directory | Where-Object { $_.Name -like "$wordToComplete*" } | ForEach-Object {
        [System.Management.Automation.CompletionResult]::new($_.Name, $_.Name, 'ParameterValue', $_.Name)
    }
}
Register-ArgumentCompleter -CommandName Export-ThemeForge -ParameterName Name -ScriptBlock {
    param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameters)

    # Get theme names from the directory
    Get-ChildItem -Path $script:themesPath -Directory | Where-Object { $_.Name -like "$wordToComplete*" } | ForEach-Object {
        [System.Management.Automation.CompletionResult]::new($_.Name, $_.Name, 'ParameterValue', $_.Name)
    }
}
Register-ArgumentCompleter -CommandName Update-ThemeForge -ParameterName Name -ScriptBlock {
    param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameters)

    # Get theme names from the directory
    Get-ChildItem -Path $script:themesPath -Directory | Where-Object { $_.Name -like "$wordToComplete*" } | ForEach-Object {
        [System.Management.Automation.CompletionResult]::new($_.Name, $_.Name, 'ParameterValue', $_.Name)
    }
}
Register-ArgumentCompleter -CommandName Remove-ThemeForge -ParameterName Name -ScriptBlock {
    param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameters)

    # Get theme names from the directory
    Get-ChildItem -Path $script:themesPath -Directory | Where-Object { $_.Name -like "$wordToComplete*" } | ForEach-Object {
        [System.Management.Automation.CompletionResult]::new($_.Name, $_.Name, 'ParameterValue', $_.Name)
    }
}
# Aliases for functions
Set-Alias -Name stf -Value Show-ThemeForge
Set-Alias -Name ctf -Value Change-ThemeForge
Set-Alias -Name gtf -Value Get-ThemeForge
Set-Alias -Name ntf -Value New-ThemeForge
Set-Alias -Name rtf -Value Remove-ThemeForge
Set-Alias -Name etf -Value Export-ThemeForge
Set-Alias -Name utf -Value Update-ThemeForge
Export-ModuleMember -Function New-ThemeForge, Change-ThemeForge, Get-ThemeForge, Remove-ThemeForge, Export-ThemeForge, Update-ThemeForge, Show-ThemeForge
Export-ModuleMember -Alias stf, ctf, gtf, ntf, rtf, etf, utf
