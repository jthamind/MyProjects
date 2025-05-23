```text
                                  (                           
  *   )    )                      )\ )                        
` )  /( ( /(    (     )      (   (()/(      (    (  (     (   
 ( )(_)))\())  ))\   (      ))\   /(_)) (   )(   )\))(   ))\  
(_(_())((_)\  /((_)  )\  ' /((_) (_))_| )\ (()\ ((_))\  /((_) 
|_   _|| |(_)(_))  _((_)) (_))   | |_  ((_) ((_) (()(_)(_))   
  | |  | ' \ / -_)| '  \()/ -_)  | __|/ _ \| '_|/ _` | / -_)  
  |_|  |_||_|\___||_|_|_| \___|  |_|  \___/|_|  \__, | \___|  
                                                |___/         
```
**Author: Jeremiah Williams**

This module allows you to create, manage, and share themes for Windows Terminal. These 
settings are stored in Terminal's settings.json file which you can access by pressing ctrl+, 
while in Terminal. This module works by creating a copy of the current theme's settings in a 
new directory called .\<theme_name>\settings.json. If the theme has a background image, 
then that file will also be copied to the new directory. 

Tab autocomplete is supported for the below commands with an asterisk(*) after the parameter.
```text
- New-ThemeForge    (ntf) <Name>	Create a new theme
- Get-ThemeForge    (gtf)		Get a list of current and available themes
- Change-ThemeForge (ctf) <Name> *     	Change terminal theme
- Update-ThemeForge (utf) <Name> *     	Update theme with current settings
- Remove-ThemeForge (rtf) <Name> *     	Remove theme
- Export-ThemeForge (etf) <Name> *     	Share theme to T:\Software\WindowsTerminalThemes
- Show-ThemeForge   (stf) <Name>	Show this help utility
```
Check out these websites for a wide range of free, custom themes:

https://windowsterminalthemes.dev

https://terminalsplash.com

https://github.com/rjcarneiro/windows-terminals

## Quick Start Guide

1. [Download ThemeForge](https://github.com/allied-solutions/DevOps/releases/latest/download/ThemeForge.zip)
2. Drop the ThemeForge folder into your PowerShell module folder (run $env:PSModulePath -split ';' to find your module directories)
3. Open Windows Terminal, then run Show-ThemeForge (stf) to confirm it's installed
4. In Windows Terminal, press ctrl and comma, then click Open JSON File at the bottom
5. Customize the hex color values under the 'schemes' section (use the websites above for ideas)
6. Save the settings.json file
7. In Windows Terminal, run New-ThemeForge (ntf) <Name> to create the new theme
8. Repeat steps 4-7 to add as many themes as you like!
