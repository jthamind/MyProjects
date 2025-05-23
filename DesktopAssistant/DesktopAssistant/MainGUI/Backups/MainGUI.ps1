Add-Type -AssemblyName System.Data
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationCore
[System.Windows.Forms.Application]::EnableVisualStyles()

#* Version History
#* 0.0.0 - 06/26/2023 - Initial Concept
#* 1.0.0 - 09/26/2023 - Initial Release
#* 1.1.0 - 10/02/2023 - Added Lender LFP module
#* 1.2.0 - 10/06/2023 - Added Premium themes
#* 1.3.0 - 10/11/2023 - Added Reverse IP Lookup

#* Table of contents for this script
#* Use Control + F to search for the section you want to jump to

#*  1. Global Variables and Functions
#*  2. Main GUI
#*  3. Restarts GUI
#*  4. NSLookup
#*  5. Server Ping
#*  6. Reverse IP Lookup
#*  7. Prod Support Tool
#*  8. Add Lender to LFP Services
#*  9. Password Manager
#* 10. Create HDTStorage Table
#* 11. Documentation Creator
#* 12. Ticket Manager
#* 13. Form Build

# ================================================================================================

# todo - List of to-do items

# todo - CURRENT ISSUES
# todo - ADD BETTER LOGGING
# todo - Put Lender LFP and Create HDTStorage inside of a runspace
# todo - Add icons for each sports team to the Form icon
# todo - Make HDTStorage modular

# todo - FUTURE RELEASE FEATURES
# todo - Make startup messages team-specific
# todo - Add Theme Builder
# todo - EC2 monitoring (unlikely to happen since SAs can't get access keys in AWS)

<#
? **********************************************************************************************************************
? START OF GLOBAL VARIABLES AND FUNCTIONS
? **********************************************************************************************************************
#>

# Sync hashtable for cross-threading
$Global:synchash = [hashtable]::Synchronized(@{})
$synchash.CWD = if ($PSScriptRoot) { $PSScriptRoot }
else { Split-Path -LiteralPath ([Environment]::GetCommandLineArgs()[0])}

# Set location to script directory
Set-location $synchash.CWD

# Synchash variables
$synchash.$OutText = $OutText

# Set timestamp function used throughout the script
function Get-Timestamp {
    return Get-Date -Format "yyyy/MM/dd hh:mm:ss"
}

$Icon = $null
$global:SecurePW = $null
$global:AltSecurePW = $null
$global:IsLenderLFPPopupActive = $false
$global:IsHDTStoragePopupActive = $false
$global:IsFeedbackPopupActive = $false
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition

# Get color themes from json file
$ColorTheme = Get-Content -Path .\ColorThemes.json | ConvertFrom-Json

# Get config values from json file
$ConfigValues = Get-Content -Path .\Config.json | ConvertFrom-Json
$userProfilePath = [Environment]::GetEnvironmentVariable("USERPROFILE")
$AlliedLogoFile = $ConfigValues.AlliedLogo.Replace("{USERPROFILE}", $UserProfilePath)
$TestingWebhookURL = $ConfigValues.TestingWebhook
$script:WorkhorseServer = $ConfigValues.WorkhorseServer

# Initialize the tooltip
$ToolTip = New-Object System.Windows.Forms.ToolTip
$ToolTip.InitialDelay = 100

# Check if DefaultUserTheme has a value or is null
if ($null -ne $ConfigValues.DefaultUserTheme -and $ConfigValues.DefaultUserTheme -ne "") {
    $SelectedTheme = $ColorTheme.PSObject.Properties | Where-Object { $_.Value.$($ConfigValues.DefaultUserTheme) } | Select-Object -ExpandProperty Name
    $themeColors = $ColorTheme.$SelectedTheme.$($ConfigValues.DefaultUserTheme)
    if ($SelectedTheme -eq 'Premium') {
        $global:DisabledBackColor = $themeColors.DisabledColor
    }
    else {
        $global:DisabledBackColor = '#A9A9A9'
    }
}
$global:DisabledForeColor = '#FFFFFF'

# Function to get the quote for the selected theme
function Get-ThemeQuote{
    $script:Quote = $ColorTheme.$SelectedTheme.$($ConfigValues.DefaultUserTheme).Quote
    $OutText.AppendText("$(Get-Timestamp) - $($ConfigValues.DefaultUserTheme) theme is now active.`r`n")
    if ($script:Quote) {
        $OutText.AppendText("$(Get-Timestamp) - $script:Quote`r`n")
    }
}

# Check if the theme falls under NBA, NFL, or MLB and set the icon accordingly
# If the theme is not one of those, set the icon to the Allied logo
function Set-FormIcon {
    $SelectedTheme = $ColorTheme.PSObject.Properties | Where-Object { $_.Value.$($ConfigValues.DefaultUserTheme) } | Select-Object -ExpandProperty Name
    if ($SelectedTheme -eq 'NBA' -or $SelectedTheme -eq 'NFL' -or $SelectedTheme -eq 'MLB') {
        $IconPath = "$scriptPath\TeamIcons\$SelectedTheme\$($ConfigValues.DefaultUserTheme).ico"

        if (Test-Path $IconPath) {
            $Icon = New-Object System.Drawing.Icon($IconPath)
            $Form.Icon = $Icon
            if ($IsLenderLFPPopupActive) {
                $script:LenderLFPPopup.Icon = $Icon
            }
            if ($IsHDTStoragePopupActive) {
                $script:HDTStoragePopup.Icon = $Icon
            }
            if ($IsFeedbackPopupActive) {
                $script:FeedbackForm.Icon = $Icon
            }
        }
    }
    else {
        $Form.Icon = $ConfigValues.AlliedLogo.Replace("{USERPROFILE}", $UserProfilePath)
        if ($IsLenderLFPPopupActive) {
            $script:LenderLFPPopup.Icon = $ConfigValues.AlliedLogo.Replace("{USERPROFILE}", $UserProfilePath)
        }
        if ($IsHDTStoragePopupActive) {
            $script:HDTStoragePopup.Icon = $ConfigValues.AlliedLogo.Replace("{USERPROFILE}", $UserProfilePath)
        }
        if ($IsFeedbackPopupActive) {
            $script:FeedbackForm.Icon = $ConfigValues.AlliedLogo.Replace("{USERPROFILE}", $UserProfilePath)
        }
    }
}

# Function to enable tooltips
function Enable-ToolTips {
	$ToolTip.SetToolTip($ClearOutTextButton, "Click to clear the output text box")
	$ToolTip.SetToolTip($SaveOutTextButton, "Click to save the output to a text file")
	$ToolTip.SetToolTip($ServersListBox, "Select a server to show its Services, Sites, and App Pools")
	$ToolTip.SetToolTip($ServicesListBox, "Select a service to check its status")
	$Tooltip.SetToolTip($IISSitesListBox, "Select a site to check its status")
	$Tooltip.SetToolTip($AppPoolsListBox, "Select an app pool to check its status")
	$ToolTip.SetToolTip($AppListCombo, "Select an application")
	$ToolTip.SetToolTip($RestartButton, "Restart selected item(s)")
	$ToolTip.SetToolTip($StartButton, "Start selected item(s)")
	$ToolTip.SetToolTip($StopButton, "Stop selected item(s)")
	$ToolTip.SetToolTip($OpenSiteButton, "Open selected site in Windows Explorer")
	$ToolTip.SetToolTip($RestartIISButton, "Restart IIS on selected server")
	$ToolTip.SetToolTip($StartIISButton, "Start IIS on selected server")
	$ToolTip.SetToolTip($StopIISButton, "Stop IIS on selected server")
	$Tooltip.SetToolTip($NSLookupButton, "Run nslookup")
	$Tooltip.SetToolTip($script:NSLookupTextBox, "Enter a hostname to resolve")
    $ToolTip.SetToolTip($ReverseIPButton, "Run reverse IP lookup")
    $ToolTip.SetToolTip($ReverseIPTextBox, "Enter a DNS name to resolve")
	$ToolTip.SetToolTip($ServerPingButton, "Click to ping server")
	$ToolTip.SetToolTip($ServerPingTextBox, "Enter a server name or IP address to test the connection")
	$ToolTip.SetToolTip($PSTCombo, "Select the environment to run the Prod Support Tool in")
	$ToolTip.SetToolTip($SelectEnvButton, "Click to switch environment to run the Prod Support Tool in")
	$ToolTip.SetToolTip($ResetEnvButton, "Click to reset the environment configuration")
	$ToolTip.SetToolTip($RunPSTButton, "Click to run the Prod Support Tool")
	$ToolTip.SetToolTip($RefreshPSTButton, "Click to refresh the Prod Support Tool files")
    $ToolTip.SetToolTip($LaunchLFPWizardButton, "Click to launch the Add Lender to LFP wizard")
	$ToolTip.SetToolTip($PWTextBox, "Enter a password to set for the remainder of the session")
	$ToolTip.SetToolTip($SetPWButton, "Click to set your password")
	$ToolTip.SetToolTip($AltSetPWButton, "Click to set an alternate password")
	$ToolTip.SetToolTip($GetPWButton, "Click to retrieve your password")
	$ToolTip.SetToolTip($AltGetPWButton, "Click to retrieve your alternate password")
	$ToolTip.SetToolTip($ClearPWButton, "Click to clear your password")
	$ToolTip.SetToolTip($AltClearPWButton, "Click to clear your alternate password")
	$ToolTip.SetToolTip($GenPWButton, "Click to generate a 16 character password with at least 1 uppercase, 1 lowercase, 1 number, and 1 special character")
	$ToolTip.SetToolTip($LaunchHDTStorageButton, "Click to launch the HDTStorage table creator")
	$ToolTip.SetToolTip($NewDocTextBox, "Enter the name of the new documentation you want to create")
	$ToolTip.SetToolTip($NewDocButton, "Click to create new documentation template")
	$ToolTip.SetToolTip($NewTicketButton, "Click to create a new ticket")
	$ToolTip.SetToolTip($NewTicketTextBox, "Enter Jira ticket number, i.e. AIH-12345")
	$ToolTip.SetToolTip($RenameTicketButton, "Click to rename the selected ticket")
	$ToolTip.SetToolTip($RenameTicketTextBox, "Enter a new ticket number to rename the selected ticket")
	$ToolTip.SetToolTip($FolderContentsListBox, "Displays the selected ticket's folder contents")
	$ToolTip.SetToolTip($CompleteTicketButton, "Click to mark a ticket complete")
	$ToolTip.SetToolTip($ReactivateTicketButton, "Click to reactivate a ticket")
	$ToolTip.SetToolTip($OpenFolderButton, "Click to open a ticket folder in Windows Explorer")
	$ToolTip.SetToolTip($ActiveTicketsListBox, "Click to select one or more tickets")
	$ToolTip.SetToolTip($CompletedTicketsListBox, "Click to select one or more tickets")
	
	$ConfigValues.HoverToolTips = "Enabled"
	$UpdatedToolTipValue = ConvertTo-Json -InputObject $ConfigValues -Depth 100
	Set-Content -Path .\Config.json -Value $UpdatedToolTipValue
 }

# Function to remove pre-existing bullet points from theme menu
function Remove-AllBulletPoints {
    foreach ($menu in @($MLBThemes, $NBAThemes, $NFLThemes, $PremiumThemes, $TraditionalThemes)) {
        foreach ($menuItem in $menu.DropDownItems) {
            $menuItem.Text = $menuItem.Text -replace '^\•\s', ''
        }
    }
}

# Function to assign the appropriate theme color based on whether there is a default theme or a user-selected theme
function Get-AppropriateColor {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ColorType # "ForeColor" or "BackColor"
    )

    if ($null -ne $global:CurrentTeamBackColor -and $null -ne $global:CurrentTeamForeColor) {
        return Get-Variable -Name ("CurrentTeam" + $ColorType) -Scope Global -ValueOnly
    }
    elseif ($null -ne $themeColors.$ColorType) {
        return $themeColors.$ColorType
    }
    else {
        return Get-Variable -Name ("Disabled" + $ColorType) -Scope Global -ValueOnly  # Default to disabled color
    }
}

# Function to update the main theme
function Update-MainTheme {
    param (
        [string]$Team,
        [string]$Category,
        [object]$ColorData
    )

    $global:CurrentTeamBackColor = $ColorData.$Category.$Team.BackColor
    $global:CurrentTeamForeColor = $ColorData.$Category.$Team.ForeColor
    
    if ($Category -eq 'Premium') {
        $global:CurrentTeamAccentColor = $ColorData.$Category.$Team.AccentColor
            $OutText.BackColor = $global:CurrentTeamAccentColor
            $ServersListBox.Backcolor = $global:CurrentTeamAccentColor
            $ServicesListBox.Backcolor = $global:CurrentTeamAccentColor
            $IISSitesListBox.Backcolor = $global:CurrentTeamAccentColor
            $AppPoolsListBox.Backcolor = $global:CurrentTeamAccentColor
            $FolderContentsListBox.BackColor = $global:CurrentTeamAccentColor
            $ActiveTicketsListBox.BackColor = $global:CurrentTeamAccentColor
            $CompletedTicketsListBox.BackColor = $global:CurrentTeamAccentColor
            $ServerPingTextBox.BackColor = $global:CurrentTeamAccentColor
            $script:NSLookupTextBox.BackColor = $global:CurrentTeamAccentColor
            $ReverseIPTextBox.BackColor = $global:CurrentTeamAccentColor
            $PWTextBox.BackColor = $global:CurrentTeamAccentColor
            $NewDocTextBox.BackColor = $global:CurrentTeamAccentColor
            $NewTicketTextBox.BackColor = $global:CurrentTeamAccentColor
            $RenameTicketTextBox.BackColor = $global:CurrentTeamAccentColor
            $global:DisabledBackColor = $ColorData.$Category.$Team.DisabledColor
        }
        else {
            $OutText.BackColor = [System.Drawing.SystemColors]::Control
            $ServersListBox.Backcolor = [System.Drawing.SystemColors]::Control
            $ServicesListBox.Backcolor = [System.Drawing.SystemColors]::Control
            $IISSitesListBox.Backcolor = [System.Drawing.SystemColors]::Control
            $AppPoolsListBox.Backcolor = [System.Drawing.SystemColors]::Control
            $FolderContentsListBox.BackColor = [System.Drawing.SystemColors]::Control
            $ActiveTicketsListBox.BackColor = [System.Drawing.SystemColors]::Control
            $CompletedTicketsListBox.BackColor = [System.Drawing.SystemColors]::Control
            $ServerPingTextBox.BackColor = [System.Drawing.SystemColors]::Control
            $script:NSLookupTextBox.BackColor = [System.Drawing.SystemColors]::Control
            $ReverseIPTextBox.BackColor = [System.Drawing.SystemColors]::Control
            $PWTextBox.BackColor = [System.Drawing.SystemColors]::Control
            $NewDocTextBox.BackColor = [System.Drawing.SystemColors]::Control
            $NewTicketTextBox.BackColor = [System.Drawing.SystemColors]::Control
            $RenameTicketTextBox.BackColor = [System.Drawing.SystemColors]::Control
            $global:DisabledBackColor = '#A9A9A9'
        }  


    # Update the default theme in the config file
    $ConfigValues.DefaultUserTheme = $Team
    $UpdatedUserTheme = ConvertTo-Json -InputObject $ConfigValues -Depth 100
    Set-Content -Path .\Config.json -Value $UpdatedUserTheme
    Get-ThemeQuote

    $Form.BackColor = $ColorData.$Category.$Team.BackColor
    $MenuStrip.BackColor = $ColorData.$Category.$Team.BackColor
    $MenuStrip.ForeColor = $ColorData.$Category.$Team.ForeColor
    $MainFormTabControl.BackColor = $ColorData.$Category.$Team.BackColor
    $SaveOutTextButton.BackColor = $ColorData.$Category.$Team.ForeColor
    $SaveOutTextButton.ForeColor = $ColorData.$Category.$Team.BackColor
    $ClearOutTextButton.BackColor = $ColorData.$Category.$Team.ForeColor
    $ClearOutTextButton.ForeColor = $ColorData.$Category.$Team.BackColor
    $SysAdminTab.BackColor = $ColorData.$Category.$Team.BackColor
    $SupportTab.BackColor = $ColorData.$Category.$Team.BackColor
    $TicketManagerTab.BackColor = $ColorData.$Category.$Team.BackColor
    $AppListCombo.BackColor = $OutText.BackColor
    $AppListLabel.ForeColor = $ColorData.$Category.$Team.ForeColor
    $OpenSiteButton.BackColor = $ColorData.$Category.$Team.ForeColor
    $OpenSiteButton.ForeColor = $ColorData.$Category.$Team.BackColor
    $RestartIISButton.BackColor = $ColorData.$Category.$Team.ForeColor
    $RestartIISButton.ForeColor = $ColorData.$Category.$Team.BackColor
    $StartIISButton.BackColor = $ColorData.$Category.$Team.ForeColor
    $StartIISButton.ForeColor = $ColorData.$Category.$Team.BackColor
    $StopIISButton.BackColor = $ColorData.$Category.$Team.ForeColor
    $StopIISButton.ForeColor = $ColorData.$Category.$Team.BackColor
    $RestartsSeparator.BackColor = $ColorData.$Category.$Team.ForeColor
    $NSLookupLabel.ForeColor = $ColorData.$Category.$Team.ForeColor
    $ServerPingLabel.ForeColor = $ColorData.$Category.$Team.ForeColor
    $ReverseIPLabel.ForeColor = $ColorData.$Category.$Team.ForeColor
    $PSTCombo.BackColor = $OutText.BackColor
    $RefreshPSTButton.BackColor = $ColorData.$Category.$Team.ForeColor
    $RefreshPSTButton.ForeColor = $ColorData.$Category.$Team.BackColor
    $LaunchLFPWizardButton.BackColor = $ColorData.$Category.$Team.ForeColor
    $LaunchLFPWizardButton.ForeColor = $ColorData.$Category.$Team.BackColor
    $LaunchLFPWizardLabel.ForeColor = $ColorData.$Category.$Team.ForeColor
    $PSTComboLabel.ForeColor = $ColorData.$Category.$Team.ForeColor
    $PSTSeparator.BackColor = $ColorData.$Category.$Team.ForeColor
    $PWTextBoxLabel.ForeColor = $ColorData.$Category.$Team.ForeColor
    $GenPWButton.BackColor = $ColorData.$Category.$Team.ForeColor
    $GenPWButton.ForeColor = $ColorData.$Category.$Team.BackColor
    $GenPWLabel.ForeColor = $ColorData.$Category.$Team.ForeColor
    $PWManagerSeparator.BackColor = $ColorData.$Category.$Team.ForeColor
    $LaunchHDTStorageButton.BackColor = $ColorData.$Category.$Team.ForeColor
    $LaunchHDTStorageButton.ForeColor = $ColorData.$Category.$Team.BackColor
    $LaunchHDTStorageLabel.ForeColor = $ColorData.$Category.$Team.ForeColor
    $NewDocLabel.ForeColor = $ColorData.$Category.$Team.ForeColor
    $NewTicketLabel.ForeColor = $ColorData.$Category.$Team.ForeColor
    $RenameTicketLabel.ForeColor = $ColorData.$Category.$Team.ForeColor

    if ($RestartButton.Enabled) {
        $RestartButton.BackColor = $ColorData.$Category.$Team.ForeColor
        $RestartButton.ForeColor = $ColorData.$Category.$Team.BackColor
    } else {
        $RestartButton.BackColor = $global:DisabledBackColor
        $RestartButton.ForeColor = $global:DisabledForeColor
    }

    if ($StartButton.Enabled) {
        $StartButton.BackColor = $ColorData.$Category.$Team.ForeColor
        $StartButton.ForeColor = $ColorData.$Category.$Team.BackColor
    } else {
        $StartButton.BackColor = $global:DisabledBackColor
        $StartButton.ForeColor = $global:DisabledForeColor
    }

    if ($StopButton.Enabled) {
        $StopButton.BackColor = $ColorData.$Category.$Team.ForeColor
        $StopButton.ForeColor = $ColorData.$Category.$Team.BackColor
    } else {
        $StopButton.BackColor = $global:DisabledBackColor
        $StopButton.ForeColor = $global:DisabledForeColor
    }

    if ($OpenSiteButton.Enabled) {
        $OpenSiteButton.BackColor = $ColorData.$Category.$Team.ForeColor
        $OpenSiteButton.ForeColor = $ColorData.$Category.$Team.BackColor
    } else {
        $OpenSiteButton.BackColor = $global:DisabledBackColor
        $OpenSiteButton.ForeColor = $global:DisabledForeColor
    }

    if ($ServerPingButton.Enabled) {
        $ServerPingButton.BackColor = $ColorData.$Category.$Team.ForeColor
        $ServerPingButton.ForeColor = $ColorData.$Category.$Team.BackColor
    } else {
        $ServerPingButton.BackColor = $global:DisabledBackColor
        $ServerPingButton.ForeColor = $global:DisabledForeColor
    }

    if ($NSLookupButton.Enabled) {
        $NSLookupButton.BackColor = $ColorData.$Category.$Team.ForeColor
        $NSLookupButton.ForeColor = $ColorData.$Category.$Team.BackColor
    } else {
        $NSLookupButton.BackColor = $global:DisabledBackColor
        $NSLookupButton.ForeColor = $global:DisabledForeColor
    }

    if ($ReverseIPButton.Enabled) {
        $ReverseIPButton.BackColor = $ColorData.$Category.$Team.ForeColor
        $ReverseIPButton.ForeColor = $ColorData.$Category.$Team.BackColor
    } else {
        $ReverseIPButton.BackColor = $global:DisabledBackColor
        $ReverseIPButton.ForeColor = $global:DisabledForeColor
    }

    if ($SelectEnvButton.Enabled) {
        $SelectEnvButton.BackColor = $ColorData.$Category.$Team.ForeColor
        $SelectEnvButton.ForeColor = $ColorData.$Category.$Team.BackColor
    } else {
        $SelectEnvButton.BackColor = $global:DisabledBackColor
        $SelectEnvButton.ForeColor = $global:DisabledForeColor
    }

    if ($ResetEnvButton.Enabled) {
        $ResetEnvButton.BackColor = $ColorData.$Category.$Team.ForeColor
        $ResetEnvButton.ForeColor = $ColorData.$Category.$Team.BackColor
    } else {
        $ResetEnvButton.BackColor = $global:DisabledBackColor
        $ResetEnvButton.ForeColor = $global:DisabledForeColor
    }

    if ($RefreshPSTButton.Enabled) {
        $RefreshPSTButton.BackColor = $ColorData.$Category.$Team.ForeColor
        $RefreshPSTButton.ForeColor = $ColorData.$Category.$Team.BackColor
    } else {
        $RefreshPSTButton.BackColor = $global:DisabledBackColor
        $RefreshPSTButton.ForeColor = $global:DisabledForeColor
    }

    if ($RunPSTButton.Enabled) {
        $RunPSTButton.BackColor = $ColorData.$Category.$Team.ForeColor
        $RunPSTButton.ForeColor = $ColorData.$Category.$Team.BackColor
    } else {
        $RunPSTButton.BackColor = $global:DisabledBackColor
        $RunPSTButton.ForeColor = $global:DisabledForeColor
    }

    if ($SetPWButton.Enabled) {
        $SetPWButton.BackColor = $ColorData.$Category.$Team.ForeColor
        $SetPWButton.ForeColor = $ColorData.$Category.$Team.BackColor
    } else {
        $SetPWButton.BackColor = $global:DisabledBackColor
        $SetPWButton.ForeColor = $global:DisabledForeColor
    }

    if ($AltSetPWButton.Enabled) {
        $AltSetPWButton.BackColor = $ColorData.$Category.$Team.ForeColor
        $AltSetPWButton.ForeColor = $ColorData.$Category.$Team.BackColor
    } else {
        $AltSetPWButton.BackColor = $global:DisabledBackColor
        $AltSetPWButton.ForeColor = $global:DisabledForeColor
    }

    if ($GetPWButton.Enabled) {
        $GetPWButton.BackColor = $ColorData.$Category.$Team.ForeColor
        $GetPWButton.ForeColor = $ColorData.$Category.$Team.BackColor
    } else {
        $GetPWButton.BackColor = $global:DisabledBackColor
        $GetPWButton.ForeColor = $global:DisabledForeColor
    }

    if ($ClearPWButton.Enabled) {
        $ClearPWButton.BackColor = $ColorData.$Category.$Team.ForeColor
        $ClearPWButton.ForeColor = $ColorData.$Category.$Team.BackColor
    } else {
        $ClearPWButton.BackColor = $global:DisabledBackColor
        $ClearPWButton.ForeColor = $global:DisabledForeColor
    }

    if ($AltGetPWButton.Enabled) {
        $AltGetPWButton.BackColor = $ColorData.$Category.$Team.ForeColor
        $AltGetPWButton.ForeColor = $ColorData.$Category.$Team.BackColor
    } else {
        $AltGetPWButton.BackColor = $global:DisabledBackColor
        $AltGetPWButton.ForeColor = $global:DisabledForeColor
    }

    if ($AltClearPWButton.Enabled) {
        $AltClearPWButton.BackColor = $ColorData.$Category.$Team.ForeColor
        $AltClearPWButton.ForeColor = $ColorData.$Category.$Team.BackColor
    } else {
        $AltClearPWButton.BackColor = $global:DisabledBackColor
        $AltClearPWButton.ForeColor = $global:DisabledForeColor
    }

    if ($NewDocButton.Enabled) {
        $NewDocButton.BackColor = $ColorData.$Category.$Team.ForeColor
        $NewDocButton.ForeColor = $ColorData.$Category.$Team.BackColor
    } else {
        $NewDocButton.BackColor = $global:DisabledBackColor
        $NewDocButton.ForeColor = $global:DisabledForeColor
    }

    if ($NewTicketButton.Enabled) {
        $NewTicketButton.BackColor = $ColorData.$Category.$Team.ForeColor
        $NewTicketButton.ForeColor = $ColorData.$Category.$Team.BackColor
    } else {
        $NewTicketButton.BackColor = $global:DisabledBackColor
        $NewTicketButton.ForeColor = $global:DisabledForeColor
    }

    if ($RenameTicketButton.Enabled) {
        $RenameTicketButton.BackColor = $ColorData.$Category.$Team.ForeColor
        $RenameTicketButton.ForeColor = $ColorData.$Category.$Team.BackColor
    } else {
        $RenameTicketButton.BackColor = $global:DisabledBackColor
        $RenameTicketButton.ForeColor = $global:DisabledForeColor
    }

    if ($ReactivateTicketButton.Enabled) {
        $ReactivateTicketButton.BackColor = $ColorData.$Category.$Team.ForeColor
        $ReactivateTicketButton.ForeColor = $ColorData.$Category.$Team.BackColor
    } else {
        $ReactivateTicketButton.BackColor = $global:DisabledBackColor
        $ReactivateTicketButton.ForeColor = $global:DisabledForeColor
    }

    if ($CompleteTicketButton.Enabled) {
        $CompleteTicketButton.BackColor = $ColorData.$Category.$Team.ForeColor
        $CompleteTicketButton.ForeColor = $ColorData.$Category.$Team.BackColor
    } else {
        $CompleteTicketButton.BackColor = $global:DisabledBackColor
        $CompleteTicketButton.ForeColor = $global:DisabledForeColor
    }

    if ($OpenFolderButton.Enabled) {
        $OpenFolderButton.BackColor = $ColorData.$Category.$Team.ForeColor
        $OpenFolderButton.ForeColor = $ColorData.$Category.$Team.BackColor
    } else {
        $OpenFolderButton.BackColor = $global:DisabledBackColor
        $OpenFolderButton.ForeColor = $global:DisabledForeColor
    }

    if ($ClearOutTextButton.Enabled) {
        $ClearOutTextButton.BackColor = $ColorData.$Category.$Team.ForeColor
        $ClearOutTextButton.ForeColor = $ColorData.$Category.$Team.BackColor
    } else {
        $ClearOutTextButton.BackColor = $global:DisabledBackColor
        $ClearOutTextButton.ForeColor = $global:DisabledForeColor
    }

    if ($SaveOutTextButton.Enabled) {
        $SaveOutTextButton.BackColor = $ColorData.$Category.$Team.ForeColor
        $SaveOutTextButton.ForeColor = $ColorData.$Category.$Team.BackColor
    } else {
        $SaveOutTextButton.BackColor = $global:DisabledBackColor
        $SaveOutTextButton.ForeColor = $global:DisabledForeColor
    }

    # Updates the Lender LFP popup controls if the popup is active
    # This is to prevent an error where the system tries to update the controls before they are created
    if ($global:IsLenderLFPPopupActive) {
        $script:LenderLFPPopup.BackColor = $ColorData.$Category.$Team.BackColor
        $script:LenderLFPPopup.ForeColor = $ColorData.$Category.$Team.ForeColor
        $script:LenderLFPComboLabel.ForeColor = $ColorData.$Category.$Team.ForeColor
        $script:LenderLFPTextBoxLabel.ForeColor = $ColorData.$Category.$Team.ForeColor
        $script:LenderLFPTicketTextBoxLabel.ForeColor = $ColorData.$Category.$Team.ForeColor
        if ($script:AddLenderLFPButton.Enabled) {
            $script:AddLenderLFPButton.BackColor = $ColorData.$Category.$Team.ForeColor
            $script:AddLenderLFPButton.ForeColor = $ColorData.$Category.$Team.BackColor
        } else {
            $script:AddLenderLFPButton.BackColor = $global:DisabledBackColor
            $script:AddLenderLFPButton.ForeColor = $global:DisabledForeColor
        }
        $script:LenderLFPPopup.Refresh()
    }

    # Updates the HDTStorage popup controls if the popup is active
    # This is to prevent an error where the system tries to update the controls before they are created
    if ($IsHDTStoragePopupActive) {
        $script:HDTStoragePopup.BackColor = $ColorData.$Category.$Team.BackColor
        $script:HDTStoragePopup.ForeColor = $ColorData.$Category.$Team.ForeColor
        $script:HDTStorageFileButton.BackColor = $ColorData.$Category.$Team.ForeColor
        $script:HDTStorageFileButton.ForeColor = $ColorData.$Category.$Team.BackColor
        $script:FileLocationLabel.Backcolor = $ColorData.$Category.$Team.BackColor
        $script:FileLocationLabel.ForeColor = $ColorData.$Category.$Team.ForeColor
        $script:DBServerLabel.Backcolor = $ColorData.$Category.$Team.BackColor
        $script:DBServerLabel.ForeColor = $ColorData.$Category.$Team.ForeColor
        $script:TableNameLabel.Backcolor = $ColorData.$Category.$Team.BackColor
        $script:TableNameLabel.ForeColor = $ColorData.$Category.$Team.ForeColor
        $script:SecurePasswordLabel.Backcolor = $ColorData.$Category.$Team.BackColor
        $script:SecurePasswordLabel.ForeColor = $ColorData.$Category.$Team.ForeColor
        if ($script:RunScriptButton.Enabled) {
            $script:RunScriptButton.BackColor = $ColorData.$Category.$Team.ForeColor
            $script:RunScriptButton.ForeColor = $ColorData.$Category.$Team.BackColor
        } else {
            $script:RunScriptButton.BackColor = $global:DisabledBackColor
            $script:RunScriptButton.ForeColor = $global:DisabledForeColor
        }
        $script:HDTStoragePopup.Refresh()
    }

    # Updates the feedback popup controls if the popup is active
    # This is to prevent an error where the system tries to update the controls before they are created
    if ($IsFeedbackPopupActive) {
        $script:FeedbackForm.BackColor = $ColorData.$Category.$Team.BackColor
        $script:FeedbackForm.ForeColor = $ColorData.$Category.$Team.ForeColor
        $script:UserNameLabel.BackColor = $ColorData.$Category.$Team.BackColor
        $script:UserNameLabel.ForeColor = $ColorData.$Category.$Team.ForeColor
        $script:FeedbackLabel.BackColor = $ColorData.$Category.$Team.BackColor
        $script:FeedbackLabel.ForeColor = $ColorData.$Category.$Team.ForeColor
        if ($script:SubmitFeedbackButton.Enabled) {
            $script:SubmitFeedbackButton.BackColor = $ColorData.$Category.$Team.ForeColor
            $script:SubmitFeedbackButton.ForeColor = $ColorData.$Category.$Team.BackColor
        } else {
            $script:SubmitFeedbackButton.BackColor = $global:DisabledBackColor
            $script:SubmitFeedbackButton.ForeColor = $global:DisabledForeColor
        }
        $script:FeedbackForm.Refresh()
    }

    Set-FormIcon

    $Form.Refresh()

    Remove-AllBulletPoints

    # Add bullet point to the active theme
    $activeTheme = $null
    switch ($Category) {
        'MLB' { $activeTheme = $MLBThemes.DropDownItems | Where-Object { $_.Text -eq $Team } }
        'NBA' { $activeTheme = $NBAThemes.DropDownItems | Where-Object { $_.Text -eq $Team } }
        'NFL' { $activeTheme = $NFLThemes.DropDownItems | Where-Object { $_.Text -eq $Team } }
        'Premium' { $activeTheme = $PremiumThemes.DropDownItems | Where-Object { $_.Text -eq $Team } }
        'Traditional' { $activeTheme = $TraditionalThemes.DropDownItems | Where-Object { $_.Text -eq $Team } }
    }

    if ($activeTheme) {
        $activeTheme.Text = "• " + $activeTheme.Text
    }
    return $global:CurrentTeamBackColor, $global:CurrentTeamForeColor
}

    # Function to evaluate if the submit feedback button should be enabled
    function Enable-SubmitFeedbackButton {
        if ($null -ne $FeedbackTextBox.Text -and $FeedbackTextBox.Text -ne '') {
            $script:SubmitFeedbackButton.Enabled = $true

            $backColor = Get-AppropriateColor -ColorType "BackColor"
            $foreColor = Get-AppropriateColor -ColorType "ForeColor"

            $script:SubmitFeedbackButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
            $script:SubmitFeedbackButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
        }
        else {
            $script:SubmitFeedbackButton.Enabled = $false
            $script:SubmitFeedbackButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
            $script:SubmitFeedbackButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
        }
    }

    # Function to evaluate if the run script button should be enabled
    function Enable-AddLenderButton {
        if ($null -ne $script:LenderLFPCombo.SelectedItem -and $script:LenderLFPIdTextBox.Text -ne '' -and $script:LenderLFPTicketTextBox.Text -ne '') {
            $script:AddLenderLFPButton.Enabled = $true

            $backColor = Get-AppropriateColor -ColorType "BackColor"
            $foreColor = Get-AppropriateColor -ColorType "ForeColor"

            $script:AddLenderLFPButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
            $script:AddLenderLFPButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
        }
        else {
            $script:AddLenderLFPButton.Enabled = $false
            $script:AddLenderLFPButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
            $script:AddLenderLFPButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
        }
    }

    # Function to evaluate if the run script button should be enabled
    function Enable-RunScriptButton {
        if ($null -ne $script:HDTStoragePopup.Tag -and $script:DBServerTextBox.Text -ne '' -and $script:TableNameTextBox.Text -ne '' -and $script:SecurePasswordTextBox.Text -ne '') {
            $script:RunScriptButton.Enabled = $true

            $backColor = Get-AppropriateColor -ColorType "BackColor"
            $foreColor = Get-AppropriateColor -ColorType "ForeColor"

            $script:RunScriptButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
            $script:RunScriptButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
    }
        else {
            $script:RunScriptButton.Enabled = $false
            $script:RunScriptButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
            $script:RunScriptButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
        }
    }

<#
? **********************************************************************************************************************
? END OF GLOBAL VARIABLES AND FUNCTIONS
? **********************************************************************************************************************
#>
<#
? **********************************************************************************************************************
? START OF MAIN GUI 
? **********************************************************************************************************************
#>

# Initialize form
$Form = New-Object System.Windows.Forms.Form
$Form.Text = "ETG Desktop Assistant"
$Form.Size = New-Object System.Drawing.Size(1000, 600)
$Form.ShowInTaskbar = $True
$Form.KeyPreview = $True
$Form.AutoSize = $True
$Form.FormBorderStyle = "Fixed3D"
$Form.MaximizeBox = $False
$Form.MinimizeBox = $True
$Form.ControlBox = $True
$Form.Icon = $Icon
$Form.StartPosition = "CenterScreen"

# Menu strip
$MenuStrip = New-Object System.Windows.Forms.MenuStrip
$MenuStrip.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
$MenuStrip.AutoSize = $True
$MenuStrip.Dock = "Top"

# File menu
$FileMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$FileMenu.Text = "File"
$SubmitFeedback = New-Object System.Windows.Forms.ToolStripMenuItem
$SubmitFeedback.Text = "Submit Feedback"
$MenuQuit = New-Object System.Windows.Forms.ToolStripMenuItem
$MenuQuit.Text = "Quit"

# Options menu
$OptionsMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$OptionsMenu.Text = "Options"
$MenuToolTips = New-Object System.Windows.Forms.ToolStripMenuItem
$MenuToolTips.Text = "Tool Tips"
$MenuColorTheme = New-Object System.Windows.Forms.ToolStripMenuItem
$MenuColorTheme.Text = "Select Theme"

# Show Tool Tips menu
$ShowHelpBannerMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$ShowHelpBannerMenu.Text = "Show Help Banner"
$ShowToolTipsMenu = New-Object System.Windows.Forms.ToolStripMenuItem

# Color theme sub-menu
$MLBThemes = New-Object System.Windows.Forms.ToolStripMenuItem
$MLBThemes.Text = "MLB Teams"
foreach ($team in $ColorTheme.MLB.PSObject.Properties) {
    $MenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $MenuItem.Text = $team.Name
    if ($MenuItem.Text -eq $ConfigValues.DefaultUserTheme) {
        $MenuItem.Text = "• " + $MenuItem.Text
    }
    $MenuItem.add_Click({
        Update-MainTheme -Team $this.Text -Category 'MLB' -ColorData $ColorTheme
    })    
    $MLBThemes.DropDownItems.Add($MenuItem) | Out-Null
}
$NBAThemes = New-Object System.Windows.Forms.ToolStripMenuItem
$NBAThemes.Text = "NBA Teams"
foreach ($team in $ColorTheme.NBA.PSObject.Properties) {
    $MenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $MenuItem.Text = $team.Name
    if ($MenuItem.Text -eq $ConfigValues.DefaultUserTheme) {
        $MenuItem.Text = "• " + $MenuItem.Text
    }
    $MenuItem.add_Click({
        Update-MainTheme -Team $this.Text -Category 'NBA' -ColorData $ColorTheme
    })
    $NBAThemes.DropDownItems.Add($MenuItem) | Out-Null
}
$NFLThemes = New-Object System.Windows.Forms.ToolStripMenuItem
$NFLThemes.Text = "NFL Teams"
foreach ($team in $ColorTheme.NFL.PSObject.Properties) {
    $MenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $MenuItem.Text = $team.Name
    if ($MenuItem.Text -eq $ConfigValues.DefaultUserTheme) {
        $MenuItem.Text = "• " + $MenuItem.Text
    }
    $MenuItem.add_Click({
        Update-MainTheme -Team $this.Text -Category 'NFL' -ColorData $ColorTheme
    })
    $NFLThemes.DropDownItems.Add($MenuItem) | Out-Null
}
$PremiumThemes = New-Object System.Windows.Forms.ToolStripMenuItem
$PremiumThemes.Text = "Premium"
foreach ($theme in $ColorTheme.Premium.PSObject.Properties) {
    $MenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $MenuItem.Text = $theme.Name
    if ($MenuItem.Text -eq $ConfigValues.DefaultUserTheme) {
        $MenuItem.Text = "• " + $MenuItem.Text
    }
    $MenuItem.add_Click({
        Update-MainTheme -Team $this.Text -Category 'Premium' -ColorData $ColorTheme
    })
    $PremiumThemes.DropDownItems.Add($MenuItem) | Out-Null
}

$TraditionalThemes = New-Object System.Windows.Forms.ToolStripMenuItem
$TraditionalThemes.Text = "Traditional"
foreach ($theme in $ColorTheme.Traditional.PSObject.Properties) {
    $MenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $MenuItem.Text = $theme.Name
    $MenuItem.add_Click({
        Update-MainTheme -Team $this.Text -Category 'Traditional' -ColorData $ColorTheme
    })
    $TraditionalThemes.DropDownItems.Add($MenuItem) | Out-Null
}

# About menu
$AboutMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$AboutMenu.Text = "About"
$MenuGitHub = New-Object System.Windows.Forms.ToolStripMenuItem
$MenuGitHub.Text = "GitHub Repo"
$MenuAboutItem = New-Object System.Windows.Forms.ToolStripMenuItem
$MenuAboutItem.Text = "About This App"

# Acknowledgements form
$AboutForm = New-Object System.Windows.Forms.Form
$AboutForm.Text = "About"
$AboutForm.Size = New-Object System.Drawing.Size(400, 300)
$AboutForm.ShowInTaskbar = $False
$AboutForm.KeyPreview = $True
$AboutForm.AutoSize = $True
$AboutForm.FormBorderStyle = "Fixed3D"
$AboutForm.MaximizeBox = $False
$AboutForm.MinimizeBox = $False

# Tab control creation
$MainFormTabControl = New-object System.Windows.Forms.TabControl
$MainFormTabControl.Size = "590,500"
$MainFormTabControl.Location = "5,65"

# System Administator Tools Tab
$SysAdminTab = New-Object System.Windows.Forms.TabPage
$SysAdminTab.DataBindings.DefaultDataSourceUpdateMode = 0
$SysAdminTab.UseVisualStyleBackColor = $True
$SysAdminTab.Name = 'SysAdminTools'
$SysAdminTab.Text = 'SysAdmin Tools'
$SysAdminTab.Font = [System.Drawing.Font]::new("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)

# Support Tools Tab
$SupportTab = New-Object System.Windows.Forms.TabPage
$SupportTab.DataBindings.DefaultDataSourceUpdateMode = 0
$SupportTab.UseVisualStyleBackColor = $True
$SupportTab.Name = 'SupportTools'
$SupportTab.Text = 'Support Tools'
$SupportTab.Font = [System.Drawing.Font]::new("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)

# Ticket Manager Tab
$TicketManagerTab = New-Object System.Windows.Forms.TabPage
$TicketManagerTab.DataBindings.DefaultDataSourceUpdateMode = 0
$TicketManagerTab.UseVisualStyleBackColor = $True
$TicketManagerTab.Name = 'TicketManager'
$TicketManagerTab.Text = 'Ticket Manager'
$TicketManagerTab.Font = [System.Drawing.Font]::new("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)

# Create a textbox for logging
$OutText = New-Object System.Windows.Forms.TextBox
$OutText.Size = New-Object System.Drawing.Size(400, 450)
$OutText.Location = New-Object System.Drawing.Point(600, 85)
$OutText.ScrollBars = "Vertical"
$OutText.Multiline = $true
$OutText.Enabled = $True
$OutText.ReadOnly = $True

# Button for clearing logging text box output
$ClearOutTextButton = New-Object System.Windows.Forms.Button
$ClearOutTextButton.Location = New-Object System.Drawing.Point(660, 540)
$ClearOutTextButton.Width = 100
$ClearOutTextButton.FlatStyle = "Popup"
$ClearOutTextButton.Text = "Clear Output"
$ClearOutTextButton.BackColor = $global:DisabledBackColor
$ClearOutTextButton.ForeColor = $global:DisabledForeColor
$ClearOutTextButton.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
$ClearOutTextButton.Enabled = $false

# Button for saving logging text box output
$SaveOutTextButton = New-Object System.Windows.Forms.Button
$SaveOutTextButton.Location = New-Object System.Drawing.Point(840, 540)
$SaveOutTextButton.Width = 100
$SaveOutTextButton.FlatStyle = "Popup"
$SaveOutTextButton.Text = "Save Output"
$SaveOutTextButton.BackColor = $global:DisabledBackColor
$SaveOutTextButton.ForeColor = $global:DisabledForeColor
$SaveOutTextButton.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
$SaveOutTextButton.Enabled = $false

# Logic for enabling/disabling the Save button
$OutText.add_TextChanged({
    if ($OutText.Text.Length -gt 0) {
        $backColor = Get-AppropriateColor -ColorType "BackColor"
        $foreColor = Get-AppropriateColor -ColorType "ForeColor"
        $SaveOutTextButton.Enabled = $true
        $ClearOutTextButton.Enabled = $true
		$SaveOutTextButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
        $SaveOutTextButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
		$ClearOutTextButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
        $ClearOutTextButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
		
    } else {
        $SaveOutTextButton.Enabled = $false
        $ClearOutTextButton.Enabled = $false
        $SaveOutTextButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
        $SaveOutTextButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
        $ClearOutTextButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
        $ClearOutTextButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
    }
})

# Event handler for clearing the output text box
$ClearOutTextButton.Add_Click({
        $OutText.Clear()
})

# Event handler for saving the output text box
$SaveOutTextButton.Add_Click({
        $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $SaveFileDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
        $SaveFileDialog.Title = "Save Output"
        $SaveFileDialog.ShowDialog()
        if ($SaveFileDialog.FileName -ne "") {
            $OutText.Text | Out-File -FilePath $SaveFileDialog.FileName
    }
})

# Click event for the File menu Quit option
$MenuQuit.add_Click({ $Form.Close() })

# Click event for the Options menu Show Tool Tips option
$ShowToolTipsMenu.add_Click({
    if ($ShowToolTipsMenu.Text -eq "Show Tool Tips") {
        $ConfigValues.HoverToolTips = "Enabled"
        $UpdatedToolTipValue = ConvertTo-Json -InputObject $ConfigValues -Depth 100
        Set-Content -Path .\Config.json -Value $UpdatedToolTipValue
        Enable-ToolTips
        $ShowToolTipsMenu.Text = "Hide Tool Tips"
    }
    else {
        $ConfigValues.HoverToolTips = "Disabled"
        $UpdatedToolTipValue = ConvertTo-Json -InputObject $ConfigValues -Depth 100
        Set-Content -Path .\Config.json -Value $UpdatedToolTipValue
        $ToolTip.RemoveAll()
        $ShowToolTipsMenu.Text = "Show Tool Tips"
    }
})

# Click event for the Submit Feedback menu option
$SubmitFeedback.add_Click({
    $script:FeedbackForm = New-Object System.Windows.Forms.Form
    $script:FeedbackForm.Text = "Submit Feedback"
    $script:FeedbackForm.Size = New-Object System.Drawing.Size(300, 350)
    $script:FeedbackForm.ShowInTaskbar = $False
    $script:FeedbackForm.KeyPreview = $True
    $script:FeedbackForm.AutoSize = $True
    $script:FeedbackForm.FormBorderStyle = "Fixed3D"
    $script:FeedbackForm.MaximizeBox = $False
    $script:FeedbackForm.MinimizeBox = $False
    $global:IsFeedbackPopupActive = $true

    # Label for providing UserName
    $script:UserNameLabel = New-Object System.Windows.Forms.Label
    $script:UserNameLabel.Location = New-Object System.Drawing.Size(20, 25)
    $script:UserNameLabel.Size = New-Object System.Drawing.Size(200, 20)
    $script:UserNameLabel.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Bold)
    $script:UserNameLabel.Text = "Add Your Name?"

    # Radio button for remaining anonymous
    $script:AnonymousRadioButton = New-Object System.Windows.Forms.RadioButton
    $script:AnonymousRadioButton.Location = New-Object System.Drawing.Point(20, 45)
    $script:AnonymousRadioButton.Size = New-Object System.Drawing.Size(200, 20)
    $script:AnonymousRadioButton.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
    $script:AnonymousRadioButton.Text = 'Remain Anonymous'
    $script:AnonymousRadioButton.Checked = $true

    # Radio button for providing name
    $script:UserNameRadioButton = New-Object System.Windows.Forms.RadioButton
    $script:UserNameRadioButton.Location = New-Object System.Drawing.Point(20, 70)
    $script:UserNameRadioButton.Size = New-Object System.Drawing.Size(200, 20)
    $script:UserNameRadioButton.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
    $script:UserNameRadioButton.Text = 'Provide Name'
    $script:UserNameRadioButton.Checked = $false

    # UserName textbox
    $script:UserNameTextBox = New-Object System.Windows.Forms.TextBox
    $script:UserNameTextBox.Location = New-Object System.Drawing.Size(20, 100)
    $script:UserNameTextBox.Size = New-Object System.Drawing.Size(170, 20)
    $script:UserNameTextBox.Enabled = $false
    $script:UserNameTextBox.Text = ''
    $script:UserNameTextBox.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)

    # Label for feedback form
    $script:FeedbackLabel = New-Object System.Windows.Forms.Label
    $script:FeedbackLabel.Location = New-Object System.Drawing.Size(20, 165)
    $script:FeedbackLabel.Size = New-Object System.Drawing.Size(200, 20)
    $script:FeedbackLabel.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Bold)
    $script:FeedbackLabel.Text = "Enter your feedback below:"
    
    # Textbox for feedback form
    $script:FeedbackTextBox = New-Object System.Windows.Forms.TextBox
    $script:FeedbackTextBox.Location = New-Object System.Drawing.Size(20, 185)
    $script:FeedbackTextBox.Size = New-Object System.Drawing.Size(240, 50)
    $script:FeedbackTextBox.Multiline = $true
    $script:FeedbackTextBox.Enabled = $true
    $script:FeedbackTextBox.Text = ''
    $script:FeedbackTextBox.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
    $script:FeedbackTextBox.ScrollBars = "Vertical"

    # Button for submitting feedback
    $script:SubmitFeedbackButton = New-Object System.Windows.Forms.Button
    $script:SubmitFeedbackButton.Location = New-Object System.Drawing.Point(20, 260)
    $script:SubmitFeedbackButton.Width = 100
    $script:SubmitFeedbackButton.FlatStyle = "Popup"
    $script:SubmitFeedbackButton.Text = "Submit"
    $script:SubmitFeedbackButton.BackColor = $global:DisabledBackColor
    $script:SubmitFeedbackButton.ForeColor = $global:DisabledForeColor
    $script:SubmitFeedbackButton.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
    $script:SubmitFeedbackButton.Enabled = $false

    # Check if DefaultUserTheme has a value or is null
    if ($null -ne $ConfigValues.DefaultUserTheme -and $ConfigValues.DefaultUserTheme -ne "") {
        # Get theme
        $SelectedTheme = $ColorTheme.PSObject.Properties | Where-Object { $_.Value.$($ConfigValues.DefaultUserTheme) } | Select-Object -ExpandProperty Name
        $themeColors = $ColorTheme.$SelectedTheme.$($ConfigValues.DefaultUserTheme)
        
        $script:FeedbackForm.BackColor = $themeColors.BackColor
        $script:FeedbackForm.ForeColor = $themeColors.ForeColor
        $script:UserNameLabel.ForeColor = $themeColors.ForeColor
        $script:FeedbackLabel.BackColor = $themeColors.BackColor
        $script:FeedbackLabel.ForeColor = $themeColors.ForeColor
        $script:SubmitFeedbackButton.BackColor = $themeColors.ForeColor
        $script:SubmitFeedbackButton.ForeColor = $themeColors.BackColor
    }

    # Call the function to set the form icon
    Set-FormIcon

    # Logic to enable the UserName textbox if the user selects the UserName radio button
    $script:UserNameRadioButton.add_CheckedChanged({
        if ($script:UserNameRadioButton.Checked) {
            $script:UserNameTextBox.Enabled = $true
        } else {
            $script:UserNameTextBox.Enabled = $false
            $script:UserNameTextBox.Text = ''
        }
    })

    # Call the Enable-SubmitFeedbackButton function at startup to see if button should be enabled
    Enable-SubmitFeedbackButton

    # Event handler for enabling/disabling the submit button
    $script:FeedbackTextBox.add_TextChanged({
        Enable-SubmitFeedbackButton
    })

    # Event handler for submitting feedback
    $script:SubmitFeedbackButton.add_Click({
        # $FeedbackUserName = ' - ' + $script:UserNameTextBox.Text
        $UserFeedback = $script:FeedbackTextBox.Text
        if ($script:UserNameRadioButton.Checked) {
            $FeedbackUserName = ' - ' + $script:UserNameTextBox.Text
        } else {
            $FeedbackUserName = ' - ' + 'Anonymous'
        }
        try {
            Invoke-Command -ComputerName $script:WorkhorseServer -ScriptBlock {
                param(
                    $TestingWebhookURL, $UserFeedback, $FeedbackUserName
                )
                    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
                    # Create the message body
                    $messageBody = @{
                        "@type" = "MessageCard"
                        "@context" = "http://schema.org/extensions"
                        "summary" = "ETG Desktop Assistant Feedback"
                        "sections" = @(
                            @{
                                "text" = ("Desktop Assistant Feedback: $($UserFeedback)$($FeedbackUserName)`r`n")
                            }
                        )
                    } | ConvertTo-Json
        
                    # Post the message to Teams
                    Invoke-RestMethod -Uri $TestingWebhookURL -Method Post -Body $messageBody -ContentType "application/json" | Out-Null
            } -ArgumentList $TestingWebhookURL, $UserFeedback, $FeedbackUserName
            $OutText.AppendText("$(Get-Timestamp) - Teams message sent for feedback.`r`n")
        } catch {
            $OutText.AppendText("$(Get-Timestamp) - An unhandled exception occurred: $($_.Exception.Message)`r`n")
        }
    })

    # Button click event for closing the feedback form
    $script:FeedbackForm.add_FormClosing({
        $global:IsFeedbackPopupActive = $false
    })

    # Feedback form build
    $script:FeedbackForm.Controls.Add($script:UserNameLabel)
    $script:FeedbackForm.Controls.Add($script:AnonymousRadioButton)
    $script:FeedbackForm.Controls.Add($script:UserNameRadioButton)
    $script:FeedbackForm.Controls.Add($script:UserNameTextBox)
    $script:FeedbackForm.Controls.Add($script:FeedbackLabel)
    $script:FeedbackForm.Controls.Add($script:FeedbackTextBox)
    $script:FeedbackForm.Controls.Add($script:SubmitFeedbackButton)

    $script:FeedbackForm.Show() | Out-Null
})

# Click event for the GitHub Repo menu option
$MenuGitHub.add_Click({ Start-Process "https://github.com/jthamind/DesktopAssistant" })

# Click event for the Acknowledgements menu option
$MenuAboutItem.add_Click({

    # Label for acknowledgements form
    $AboutLabel = New-Object System.Windows.Forms.Label
    $AboutLabel.Location = New-Object System.Drawing.Size(115, 75)
    $AboutLabel.Size = New-Object System.Drawing.Size(150, 20)
    $AboutLabel.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Bold)
    $AboutLabel.Text = 'ETG Desktop Assistant'

    # Picture box for Allied Solutions logo
    $AlliedLogo = New-Object System.Windows.Forms.PictureBox
    $AlliedLogo.Location = New-Object System.Drawing.Size(50, 5)
    $AlliedLogo.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
    $AlliedLogo.Size = New-Object System.Drawing.Size(275, 100)
    $AlliedLogoImage = [System.Drawing.Image]::FromFile($AlliedLogoFile)
    $AlliedLogo.Image = $AlliedLogoImage

    # Label for version numbers
    $VersionLabel = New-Object System.Windows.Forms.Label
    $VersionLabel.Location = New-Object System.Drawing.Size(145, 105)
    $VersionLabel.Size = New-Object System.Drawing.Size(150, 20)
    $VersionLabel.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
    $VersionLabel.Text = "Version: 1.1.0"

    # Label for upcoming features
    $UpcomingFeaturesLabel = New-Object System.Windows.Forms.Label
    $UpcomingFeaturesLabel.Location = New-Object System.Drawing.Size(125, 140)
    $UpcomingFeaturesLabel.Size = New-Object System.Drawing.Size(200, 75)
    $UpcomingFeaturesLabel.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Bold)
    $UpcomingFeaturesLabel.Text = "Upcoming Features`r`n`r`n• Theme Builder`r`n• AWS EC2 Monitoring"

    # Label for Allied IMPACT
    $AlliedImpactLabel = New-Object System.Windows.Forms.Label
    $AlliedImpactLabel.Location = New-Object System.Drawing.Size(40, 250)
    $AlliedImpactLabel.Width = 300
    $AlliedImpactLabel.Font = [System.Drawing.Font]::new("Arial", 7, [System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Italic)
    $AlliedImpactLabel.Text = "Made with Passion to make an IMPACT at Allied Solutions, LLC."

    # Check if DefaultUserTheme has a value or is null
    if ($null -ne $ConfigValues.DefaultUserTheme -and $ConfigValues.DefaultUserTheme -ne "") {
        # Get theme
        $SelectedTheme = $ColorTheme.PSObject.Properties | Where-Object { $_.Value.$($ConfigValues.DefaultUserTheme) } | Select-Object -ExpandProperty Name
        $themeColors = $ColorTheme.$SelectedTheme.$($ConfigValues.DefaultUserTheme)
        
		$AboutForm.BackColor = $themeColors.BackColor
		$AboutForm.ForeColor = $themeColors.ForeColor
        $AboutLabel.ForeColor = $themeColors.ForeColor
    }
    $AboutForm.Controls.Add($AboutLabel)
    $AboutForm.Controls.Add($AlliedLogo)
    $AboutForm.Controls.Add($VersionLabel)
    $AboutForm.Controls.Add($UpcomingFeaturesLabel)
    $AboutForm.Controls.Add($AlliedImpactLabel)
    $AboutForm.ShowDialog()
})

<#
? **********************************************************************************************************************
? END OF MAIN GUI 
? **********************************************************************************************************************
#>
<#
? **********************************************************************************************************************
? START OF RESTARTS GUI
? **********************************************************************************************************************
#>

# Restarts tab control creation
$RestartsTabControl = New-object System.Windows.Forms.TabControl
$RestartsTabControl.Size = "250,250"
$RestartsTabControl.Location = "225,75"

# Individual servers list box
$ServersListBox = New-Object System.Windows.Forms.ListBox
$ServersListBox.Location = New-Object System.Drawing.Point(5,95)
$ServersListBox.Size = New-Object System.Drawing.Size(200,240)
$ServersListBox.SelectionMode = 'One'

# Services list box
$ServicesListBox = New-Object System.Windows.Forms.ListBox
$ServicesListBox.Location = New-Object System.Drawing.Point(0,0)
$ServicesListBox.Size = New-Object System.Drawing.Size(245,240)
$ServicesListBox.SelectionMode = 'MultiExtended'

# IIS Sites list box
$IISSitesListBox = New-Object System.Windows.Forms.ListBox
$IISSitesListBox.Location = New-Object System.Drawing.Point(0,0)
$IISSitesListBox.Size = New-Object System.Drawing.Size(245,240)
$IISSitesListBox.SelectionMode = 'MultiExtended'

# IIS App Pools list box
$AppPoolsListBox = New-Object System.Windows.Forms.ListBox
$AppPoolsListBox.Location = New-Object System.Drawing.Point(0,0)
$AppPoolsListBox.Size = New-Object System.Drawing.Size(245,240)
$AppPoolsListBox.SelectionMode = 'MultiExtended'

# Combobox for application selection
$AppListCombo = New-Object System.Windows.Forms.ComboBox
$AppListCombo.Location = New-Object System.Drawing.Point(5,65)
$AppListCombo.Size = New-Object System.Drawing.Size(200, 200)
$AppListCombo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$AppListCombo.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)


# Label applist combo box
$AppListLabel = New-Object System.Windows.Forms.Label
$AppListLabel.Location = New-Object System.Drawing.Size(5, 40)
$AppListLabel.Size = New-Object System.Drawing.Size(150, 20)
$AppListLabel.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Bold)
$AppListLabel.Text = 'Select a Server'

# Tab for services list
$ServicesTab = New-Object System.Windows.Forms.TabPage
$ServicesTab.DataBindings.DefaultDataSourceUpdateMode = 0
$ServicesTab.UseVisualStyleBackColor = $True
$ServicesTab.Name = 'ServicesTab'
$ServicesTab.Text = 'Services'
$ServicesTab.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)

# Tab for IIS sites list
$IISSitesTab = New-Object System.Windows.Forms.TabPage
$IISSitesTab.DataBindings.DefaultDataSourceUpdateMode = 0
$IISSitesTab.UseVisualStyleBackColor = $True
$IISSitesTab.Name = 'IISSites'
$IISSitesTab.Text = 'IIS Sites'
$IISSitesTab.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)

# Tab for IIS App Pools list
$AppPoolsTab = New-Object System.Windows.Forms.TabPage
$AppPoolsTab.DataBindings.DefaultDataSourceUpdateMode = 0
$AppPoolsTab.UseVisualStyleBackColor = $True
$AppPoolsTab.Name = 'IISAppPools'
$AppPoolsTab.Text = 'App Pools'
$AppPoolsTab.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)

# Button for restarting services
$RestartButton = New-Object System.Windows.Forms.Button
$RestartButton.Location = New-Object System.Drawing.Point(490, 95)
$RestartButton.Width = 75
$RestartButton.BackColor = $global:DisabledBackColor
$RestartButton.ForeColor = $global:DisabledForeColor
$RestartButton.FlatStyle = "Popup"
$RestartButton.Text = "Restart"
$RestartButton.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
$RestartButton.Enabled = $false

# Button for starting services
$StartButton = New-Object System.Windows.Forms.Button
$StartButton.Location = New-Object System.Drawing.Point(490, 125)
$StartButton.Width = 75
$StartButton.BackColor = $global:DisabledBackColor
$StartButton.ForeColor = $global:DisabledForeColor
$StartButton.FlatStyle = "Popup"
$StartButton.Text = "Start"
$StartButton.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
$StartButton.Enabled = $false

# Button for stopping services
$StopButton = New-Object System.Windows.Forms.Button
$StopButton.Location = New-Object System.Drawing.Point(490, 155)
$StopButton.Width = 75
$StopButton.BackColor = $global:DisabledBackColor
$StopButton.ForeColor = $global:DisabledForeColor
$StopButton.FlatStyle = "Popup"
$StopButton.Text = "Stop"
$StopButton.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
$StopButton.Enabled = $false

# Button for for opening IIS Site in Windows Explorer
$OpenSiteButton = New-Object System.Windows.Forms.Button
$OpenSiteButton.Location = New-Object System.Drawing.Point(490, 185)
$OpenSiteButton.Width = 75
$OpenSiteButton.BackColor = $global:DisabledBackColor
$OpenSiteButton.ForeColor = $global:DisabledForeColor
$OpenSiteButton.FlatStyle = "Popup"
$OpenSiteButton.Text = "Open"
$OpenSiteButton.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
$OpenSiteButton.Enabled = $false
$OpenSiteButton.Visible = $false

# Button for restarting IIS on server
$RestartIISButton = New-Object System.Windows.Forms.Button
$RestartIISButton.Location = New-Object System.Drawing.Point(490, 215)
$RestartIISButton.Width = 75
$RestartIISButton.BackColor = $global:DisabledBackColor
$RestartIISButton.ForeColor = $global:DisabledForeColor
$RestartIISButton.FlatStyle = "Popup"
$RestartIISButton.Text = "Restart IIS"
$RestartIISButton.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
$RestartIISButton.Enabled = $false
$RestartIISButton.Visible = $false

# Button for starting IIS on server
$StartIISButton = New-Object System.Windows.Forms.Button
$StartIISButton.Location = New-Object System.Drawing.Point(490, 245)
$StartIISButton.Width = 75
$StartIISButton.BackColor = $global:DisabledBackColor
$StartIISButton.ForeColor = $global:DisabledForeColor
$StartIISButton.FlatStyle = "Popup"
$StartIISButton.Text = "Start IIS"
$StartIISButton.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
$StartIISButton.Enabled = $false
$StartIISButton.Visible = $false

# Button for stopping IIS on server
$StopIISButton = New-Object System.Windows.Forms.Button
$StopIISButton.Location = New-Object System.Drawing.Point(490, 275)
$StopIISButton.Width = 75
$StopIISButton.BackColor = $global:DisabledBackColor
$StopIISButton.ForeColor = $global:DisabledForeColor
$StopIISButton.FlatStyle = "Popup"
$StopIISButton.Text = "Stop IIS"
$StopIISButton.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
$StopIISButton.Enabled = $false
$StopIISButton.Visible = $false

# Separator line under Restarts GUI
$RestartsSeparator = New-Object System.Windows.Forms.Label
$RestartsSeparator.Location = New-Object System.Drawing.Size(0, 350)
$RestartsSeparator.Size = New-Object System.Drawing.Size(1000, 2)

# Disable restart button if no options are selected in the ServicesListBox. Enable if options are selected.
$ServicesListBox.add_SelectedIndexChanged({
    if ($ServicesListBox.SelectedItems.Count -gt 0) {
        $backColor = Get-AppropriateColor -ColorType "BackColor"
        $foreColor = Get-AppropriateColor -ColorType "ForeColor"
        $RestartButton.Enabled = $true
        $StartButton.Enabled = $true
        $StopButton.Enabled = $true
        $RestartButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
        $RestartButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
        $StartButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
        $StartButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
        $StopButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
        $StopButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
    } else {
        $RestartButton.Enabled = $false
        $StartButton.Enabled = $false
        $StopButton.Enabled = $false
        $RestartButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
        $RestartButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
        $StartButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
        $StartButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
        $StopButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
        $StopButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
    }
})

# Disable restart button if no options are selected in the IISSitesListBox. Enable if options are selected.
$IISSitesListBox.add_SelectedIndexChanged({
    if ($IISSitesListBox.SelectedItems.Count -gt 0) {
        $backColor = Get-AppropriateColor -ColorType "BackColor"
        $foreColor = Get-AppropriateColor -ColorType "ForeColor"
        $RestartButton.Enabled = $true
        $StartButton.Enabled = $true
        $StopButton.Enabled = $true
        $OpenSiteButton.Enabled = $true
        $RestartButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
        $RestartButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
        $StartButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
        $StartButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
        $StopButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
        $StopButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
        $OpenSiteButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
        $OpenSiteButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
    } else {
        $RestartButton.Enabled = $false
        $StartButton.Enabled = $false
        $StopButton.Enabled = $false
        $OpenSiteButton.Enabled = $false
        $RestartButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
        $RestartButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
        $StartButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
        $StartButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
        $StopButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
        $StopButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
        $OpenSiteButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
        $OpenSiteButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
    }
})

# Disable restart button if no options are selected in the AppPoolsListBox. Enable if options are selected.
$AppPoolsListBox.add_SelectedIndexChanged({
    if ($AppPoolsListBox.SelectedItems.Count -gt 0) {
        $backColor = Get-AppropriateColor -ColorType "BackColor"
        $foreColor = Get-AppropriateColor -ColorType "ForeColor"
        $RestartButton.Enabled = $true
        $StartButton.Enabled = $true
        $StopButton.Enabled = $true
        $RestartButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
        $RestartButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
        $StartButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
        $StartButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
        $StopButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
        $StopButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
    } else {
        $RestartButton.Enabled = $false
        $StartButton.Enabled = $false
        $StopButton.Enabled = $false
        $RestartButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
        $RestartButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
        $StartButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
        $StartButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
        $StopButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
        $StopButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
    }
})

# Set Open button to invisible if the selected tab is not IIS Sites
$RestartsTabControl.add_SelectedIndexChanged({
    $SelectedTab = $RestartsTabControl.SelectedTab.Text
    $synchash.SelectedTab = $SelectedTab
    if ($SelectedTab -eq "IIS Sites") {
        $OpenSiteButton.Visible = $true
    } else {
        $OpenSiteButton.Visible = $false
    }
})

# Variables for importing server list from CSV
$csvPath = $ConfigValues.csvPath.Replace("{USERPROFILE}", $userProfilePath)
$ServerCSV = Import-CSV $csvPath
$csvHeaders = ($ServerCSV | Get-Member -MemberType NoteProperty).name

# Populate the AppListCombo with the list of applications from the CSV
foreach ($header in $csvHeaders) {
    [void]$AppListCombo.Items.Add($header)
}

# Event handler for selecting a server from ServersListBox
$ServersListBox.add_SelectedIndexChanged({
    if ($ServersListBox.SelectedIndex -ge 0) {
        $backColor = Get-AppropriateColor -ColorType "BackColor"
        $foreColor = Get-AppropriateColor -ColorType "ForeColor"
        $StartIISButton.Visible = $true
        $StopIISButton.Visible = $true
        $RestartIISButton.Visible = $true
        $StartIISButton.Enabled = $true
        $StopIISButton.Enabled = $true
        $RestartIISButton.Enabled = $true
        $StartIISButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
        $StartIISButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
        $StopIISButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
        $StopIISButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
        $RestartIISButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
        $RestartIISButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
    } else {
        $StartIISButton.Visible = $false
        $StartIISButton.Enabled = $false
        $StopIISButton.Visible = $false
        $StopIISButton.Enabled = $false
        $RestartIISButton.Visible = $false
        $RestartIISButton.Enabled = $false
        $StartIISButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
        $StartIISButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
        $StopIISButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
        $StopIISButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
        $RestartIISButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
        $RestartIISButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
    }
})

# Function to open an async runspace for populating the list boxes
function Open-PopulateListBoxRunspace {
    param (
        [System.Windows.Forms.TabControl]$RestartsTabControl,
        [System.Windows.Forms.ListBox]$ServersListBox,
        [System.Windows.Forms.ListBox]$ServicesListBox,
        [System.Windows.Forms.ListBox]$IISSitesListBox,
        [System.Windows.Forms.ListBox]$AppPoolsListBox,
        [System.Windows.Forms.TextBox]$OutText,
        [hashtable]$synchash
    )

    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.Open()

    $psCmd = [PowerShell]::Create().AddScript({
        param (
            [System.Windows.Forms.TabControl]$RestartsTabControl,
            [System.Windows.Forms.ListBox]$ServersListBox,
            [System.Windows.Forms.ListBox]$ServicesListBox,
            [System.Windows.Forms.ListBox]$IISSitesListBox,
            [System.Windows.Forms.ListBox]$AppPoolsListBox,
            [System.Windows.Forms.TextBox]$OutText,
            [hashtable]$synchash
        )

        function Get-Timestamp {
            return (Get-Date -Format "yyyy/MM/dd HH:mm:ss")
        }

        $SelectedTab = $RestartsTabControl.SelectedTab.Text
        $synchash.SelectedTab = $SelectedTab

        $SelectedServer = $ServersListBox.SelectedItem
        if ($null -eq $SelectedServer) {
            return
        }

        switch ($SelectedTab) {
            "Services" {
                $OutText.AppendText("$((Get-Timestamp)) - Retrieving services on $SelectedServer...`r`n")
                $ServicesListBox.Items.Clear()
                $Services = Invoke-Command -ComputerName $SelectedServer -ScriptBlock {
                    Get-Service | ForEach-Object { $_.DisplayName } | Sort-Object
                }
                foreach ($service in $Services) {
                    [void]$ServicesListBox.Items.Add($service)
                }
            }
            "IIS Sites" {
                $OutText.AppendText("$((Get-Timestamp)) - Retrieving sites on $SelectedServer...`r`n")
                $IISSitesListBox.Items.Clear()
                $Sites = Invoke-Command -ComputerName $SelectedServer -ScriptBlock {
                    Import-Module WebAdministration
                    Get-Website | ForEach-Object { $_.Name } | Sort-Object
                }
                foreach ($site in $Sites) {
                    [void]$IISSitesListBox.Items.Add($site)
                }
            }
            "App Pools" {
                $OutText.AppendText("$((Get-Timestamp)) - Retrieving AppPools on $SelectedServer...`r`n")
                $AppPoolsListBox.Items.Clear()
                $AppPools = Invoke-Command -ComputerName $SelectedServer -ScriptBlock {
                    Import-Module WebAdministration
                    Get-IISAppPool | ForEach-Object { $_.Name } | Sort-Object
                }
                foreach ($apppool in $AppPools) {
                    [void]$AppPoolsListBox.Items.Add($apppool)
                }
            }
        }
    }).AddParameters(@{
        RestartsTabControl = $RestartsTabControl
        ServersListBox = $ServersListBox
        ServicesListBox = $ServicesListBox
        IISSitesListBox = $IISSitesListBox
        AppPoolsListBox = $AppPoolsListBox
        OutText = $OutText
        synchash = $synchash
    })
    
    $psCmd.Runspace = $runspace
    $null = $psCmd.BeginInvoke()
}

# Event handler for the AppListCombo's SelectedIndexChanged event
$AppListCombo_SelectedIndexChanged = {
    $selectedHeader = $AppListCombo.SelectedItem.ToString()
    $ServicesListBox.Items.Clear()
    $IISSitesListBox.Items.Clear()
    $AppPoolsListBox.Items.Clear()
    $OutText.AppendText("$(Get-Timestamp) - Selected server list: $selectedHeader`r`n")
    $servers = $ServerCSV | ForEach-Object { $_.$selectedHeader } | Where-Object { $_ -ne '' }
    $ServersListBox.Items.Clear()
    $servers | ForEach-Object {
        [void]$ServersListBox.Items.Add($_)
    }
}

# Button click event handler for restarting one or more services, IIS sites, or app pools in an async runspace pool
$RestartButton.Add_Click({
    $SelectedServer = $ServersListBox.SelectedItem
    $SelectedTab = $RestartsTabControl.SelectedTab.Text
    $synchash.SelectedTab = $SelectedTab

    $runspacePool = [runspacefactory]::CreateRunspacePool(1, 5)
    $runspacePool.Open()
    $runspaces = @()

    $updateUI = {
        param($message)
        $OutText.AppendText("$(Get-Timestamp) - $message`r`n")
    }

    $scriptblock = {
        param($SelectedServer, $item, $SelectedTab, $updateUI)
        try {
            $msg = "Restarting $item on $SelectedServer..."
            & $updateUI $msg
            
            $result = Invoke-Command -ComputerName $SelectedServer -ScriptBlock {
                param($item, $SelectedTab)
                if ($SelectedTab -eq 'Services') {
                    Restart-Service -Name $item
                } elseif ($SelectedTab -eq 'IIS Sites') {
                    Import-Module WebAdministration
                    Restart-WebSite -Name $item
                } elseif ($SelectedTab -eq 'App Pools') {
                    Import-Module WebAdministration
                    Restart-WebAppPool -Name $item
                }
                return "Restarted $item"
            } -ArgumentList $item, $SelectedTab
            
            & $updateUI $result
        } catch {
            & $updateUI "An error occurred: $($_.Exception.Message)"
        }
    }

    if ($SelectedTab -match "Services|IIS Sites|App Pools") {
        $listBox = switch ($SelectedTab) {
            'Services' { $ServicesListBox }
            'IIS Sites' { $IISSitesListBox }
            'App Pools' { $AppPoolsListBox }
        }
        
        foreach ($item in $listBox.SelectedItems) {
            $runspace = [powershell]::Create()
            $runspace.RunspacePool = $runspacePool
            $runspace.AddScript($scriptblock).AddArgument($SelectedServer).AddArgument($item).AddArgument($SelectedTab).AddArgument($updateUI)
            $runspaces += [PSCustomObject]@{ Pipe = $runspace; Status = $runspace.BeginInvoke() }
        }

        Register-ObjectEvent -InputObject $runspacePool -EventName 'StateChanged' -Action {
            if ($runspacePool.State -eq [System.Management.Automation.Runspaces.RunspacePoolState]::Closed) {
                $runspacePool.Dispose()
            }
        }
    }
})

# Button click event handler for starting services, IIS sites, or app pools in an async runspace pool
$StartButton.Add_Click({
    $SelectedServer = $ServersListBox.SelectedItem
    $SelectedTab = $RestartsTabControl.SelectedTab.Text
    $synchash.SelectedTab = $SelectedTab

    $runspacePool = [runspacefactory]::CreateRunspacePool(1, 5)
    $runspacePool.Open()
    $runspaces = @()

    $updateUI = {
        param($message)
        $OutText.AppendText("$(Get-Timestamp) - $message`r`n")
    }

    $scriptblock = {
        param($SelectedServer, $item, $SelectedTab, $updateUI)
        try {
            $msg = "Starting $item on $SelectedServer..."
            & $updateUI $msg
            
            $result = Invoke-Command -ComputerName $SelectedServer -ScriptBlock {
                param($item, $SelectedTab)
                if ($SelectedTab -eq 'Services') {
                    Start-Service -Name $item
                } elseif ($SelectedTab -eq 'IIS Sites') {
                    Import-Module WebAdministration
                    Start-WebSite -Name $item
                } elseif ($SelectedTab -eq 'App Pools') {
                    Import-Module WebAdministration
                    Start-WebAppPool -Name $item
                }
                return "Started $item"
            } -ArgumentList $item, $SelectedTab
            
            & $updateUI $result
        } catch {
            & $updateUI "An error occurred: $($_.Exception.Message)"
        }
    }

    if ($SelectedTab -match "Services|IIS Sites|App Pools") {
        $listBox = switch ($SelectedTab) {
            'Services' { $ServicesListBox }
            'IIS Sites' { $IISSitesListBox }
            'App Pools' { $AppPoolsListBox }
        }
        
        foreach ($item in $listBox.SelectedItems) {
            $runspace = [powershell]::Create()
            $runspace.RunspacePool = $runspacePool
            $runspace.AddScript($scriptblock).AddArgument($SelectedServer).AddArgument($item).AddArgument($SelectedTab).AddArgument($updateUI)
            $runspaces += [PSCustomObject]@{ Pipe = $runspace; Status = $runspace.BeginInvoke() }
        }

        Register-ObjectEvent -InputObject $runspacePool -EventName 'StateChanged' -Action {
            if ($runspacePool.State -eq [System.Management.Automation.Runspaces.RunspacePoolState]::Closed) {
                $runspacePool.Dispose()
            }
        }
    }
})

# Button click event handler for stopping one or more services, IIS sites, or app pools in an async runspace pool
$StopButton.Add_Click({
    $SelectedServer = $ServersListBox.SelectedItem
    $SelectedTab = $RestartsTabControl.SelectedTab.Text
    $synchash.SelectedTab = $SelectedTab

    $runspacePool = [runspacefactory]::CreateRunspacePool(1, 5)
    $runspacePool.Open()
    $runspaces = @()

    $updateUI = {
        param($message)
        $OutText.AppendText("$(Get-Timestamp) - $message`r`n")
    }

    $scriptblock = {
        param($SelectedServer, $item, $SelectedTab, $updateUI)
        try {
            $msg = "Stopping $item on $SelectedServer..."
            & $updateUI $msg
            
            $result = Invoke-Command -ComputerName $SelectedServer -ScriptBlock {
                param($item, $SelectedTab)
                if ($SelectedTab -eq 'Services') {
                    Stop-Service -Name $item
                } elseif ($SelectedTab -eq 'IIS Sites') {
                    Import-Module WebAdministration
                    Stop-WebSite -Name $item
                } elseif ($SelectedTab -eq 'App Pools') {
                    Import-Module WebAdministration
                    Stop-WebAppPool -Name $item
                }
                return "Stopped $item"
            } -ArgumentList $item, $SelectedTab
            
            & $updateUI $result
        } catch {
            & $updateUI "An error occurred: $($_.Exception.Message)"
        }
    }

    if ($SelectedTab -match "Services|IIS Sites|App Pools") {
        $listBox = switch ($SelectedTab) {
            'Services' { $ServicesListBox }
            'IIS Sites' { $IISSitesListBox }
            'App Pools' { $AppPoolsListBox }
        }
        
        foreach ($item in $listBox.SelectedItems) {
            $runspace = [powershell]::Create()
            $runspace.RunspacePool = $runspacePool
            $runspace.AddScript($scriptblock).AddArgument($SelectedServer).AddArgument($item).AddArgument($SelectedTab).AddArgument($updateUI)
            $runspaces += [PSCustomObject]@{ Pipe = $runspace; Status = $runspace.BeginInvoke() }
        }

        Register-ObjectEvent -InputObject $runspacePool -EventName 'StateChanged' -Action {
            if ($runspacePool.State -eq [System.Management.Automation.Runspaces.RunspacePoolState]::Closed) {
                $runspacePool.Dispose()
            }
        }
    }
})

# Button click event handler for opening one or more IIS site directories in an async runspace pool
$OpenSiteButton.Add_Click({
    $SelectedServer = $ServersListBox.SelectedItem
    $SelectedTab = $RestartsTabControl.SelectedTab.Text
    $synchash.SelectedTab = $SelectedTab
    
    $runspacePool = [runspacefactory]::CreateRunspacePool(1, 5)
    $runspacePool.Open()
    $runspaces = @()

    $updateUI = {
        param($message)
        $OutText.AppendText("$(Get-Timestamp) - $message`r`n")
    }
    
    $scriptblock = {
        param($SelectedServer, $item, $updateUI)
        
        try {
            $msg = "Opening $item on $SelectedServer..."
            & $updateUI $msg
            
            $remoteSiteFolder = Invoke-Command -ComputerName $SelectedServer -ScriptBlock {
                param($item)
                Import-Module WebAdministration
                return "\\$($env:COMPUTERNAME)\e$\inetpub\$item"
            } -ArgumentList $item

            Start-Process explorer.exe -ArgumentList $remoteSiteFolder
            & $updateUI "Opened $item"
        } catch {
            & $updateUI "An error occurred: $($_.Exception.Message)"
        }
    }
    
    if ($SelectedTab -eq 'IIS Sites') {
        foreach ($item in $IISSitesListBox.SelectedItems) {
            $runspace = [powershell]::Create()
            $runspace.RunspacePool = $runspacePool
            $runspace.AddScript($scriptblock).AddArgument($SelectedServer).AddArgument($item).AddArgument($updateUI)
            $runspaces += [PSCustomObject]@{ Pipe = $runspace; Status = $runspace.BeginInvoke() }
        }

        Register-ObjectEvent -InputObject $runspacePool -EventName 'StateChanged' -Action {
            if ($runspacePool.State -eq [System.Management.Automation.Runspaces.RunspacePoolState]::Closed) {
                $runspacePool.Dispose()
            }
        }
    }
})

# Event handler for restarting IIS on a server
$RestartIISButton.Add_Click({
    $SelectedServer = $ServersListBox.SelectedItem
    $SelectedTab = $RestartsTabControl.SelectedTab.Text
    $synchash.SelectedTab = $SelectedTab

    $psCmd = [PowerShell]::Create().AddScript({
        param (
            [string]$SelectedServer,
            [System.Windows.Forms.TextBox]$OutText
        )
        
        function Get-Timestamp {
            return Get-Date -Format "yyyy/MM/dd hh:mm:ss"
        }

        try {
            Invoke-Command -ComputerName $SelectedServer -ScriptBlock {
                iisreset /restart
            }
            $OutText.AppendText("$(Get-Timestamp) - Restarted IIS on $SelectedServer successfully.`r`n")
        } catch {
            $OutText.AppendText("$(Get-Timestamp) - Failed to restart IIS on ${SelectedServer}: $($_.Exception.Message)`r`n")
        }
    }).AddParameters(@{SelectedServer = $SelectedServer; OutText = $OutText})

    $runspace = [RunspaceFactory]::CreateRunspace()
    $psCmd.Runspace = $runspace

    try {
        $runspace.Open()
        $OutText.AppendText("$(Get-Timestamp) - Restarting IIS on $SelectedServer...`r`n")

        $psCmd.BeginInvoke()
    } catch {
        $OutText.AppendText("$(Get-Timestamp) - An error occurred while invoking the command: $($_.Exception.Message)`r`n")
    } finally {
        Register-ObjectEvent -InputObject $psCmd -EventName InvocationStateChanged -Action {
            $Sender.Dispose()
            $Event.SourceEventArgs.Runspace.Close()
            $Event.SourceEventArgs.Runspace.Dispose()
        }
    }
})

# Button click event handler for starting IIS on a server
$StartIISButton.Add_Click({
    $SelectedServer = $ServersListBox.SelectedItem
    $SelectedTab = $RestartsTabControl.SelectedTab.Text
    $synchash.SelectedTab = $SelectedTab
    
        $psCmd = [PowerShell]::Create().AddScript({
            param (
                [string]$SelectedServer,
                [System.Windows.Forms.TextBox]$OutText
            )

            function Get-Timestamp {
                return Get-Date -Format "yyyy/MM/dd hh:mm:ss"
            }

            try {
                Invoke-Command -ComputerName $SelectedServer -ScriptBlock {
                    iisreset /start
                }
                $OutText.AppendText("$(Get-Timestamp) - Started IIS on $SelectedServer successfully.`r`n")
            } catch {
                $OutText.AppendText("$(Get-Timestamp) - Failed to start IIS on ${SelectedServer}: $($_.Exception.Message)`r`n")
            }
        }).AddParameters(@{SelectedServer = $SelectedServer; OutText = $OutText})

        $runspace = [RunspaceFactory]::CreateRunspace()
        $psCmd.Runspace = $runspace

        try {
            $runspace.Open()
            $OutText.AppendText("$(Get-Timestamp) - Starting IIS on $SelectedServer...`r`n")

            $psCmd.BeginInvoke()
        } catch {
            $OutText.AppendText("$(Get-Timestamp) - An error occurred while invoking the command: $($_.Exception.Message)`r`n")
        } finally {
            Register-ObjectEvent -InputObject $psCmd -EventName InvocationStateChanged -Action {
                $Sender.Dispose()
                $Event.SourceEventArgs.Runspace.Close()
                $Event.SourceEventArgs.Runspace.Dispose()
            }
        }
})

# Button click event handler for starting IIS on a server
$StopIISButton.Add_Click({
    $SelectedServer = $ServersListBox.SelectedItem
    $SelectedTab = $RestartsTabControl.SelectedTab.Text
    $synchash.SelectedTab = $SelectedTab
    
        $psCmd = [PowerShell]::Create().AddScript({
            param (
                [string]$SelectedServer,
                [System.Windows.Forms.TextBox]$OutText
            )

            function Get-Timestamp {
                return Get-Date -Format "yyyy/MM/dd hh:mm:ss"
            }

            try {
                Invoke-Command -ComputerName $SelectedServer -ScriptBlock {
                    iisreset /stop
                }
                $OutText.AppendText("$(Get-Timestamp) - Stopped IIS on $SelectedServer successfully.`r`n")
            } catch {
                $OutText.AppendText("$(Get-Timestamp) - Failed to stop IIS on ${SelectedServer}: $($_.Exception.Message)`r`n")
            }
        }).AddParameters(@{SelectedServer = $SelectedServer; OutText = $OutText})

        $runspace = [RunspaceFactory]::CreateRunspace()
        $psCmd.Runspace = $runspace

        try {
            $runspace.Open()
            $OutText.AppendText("$(Get-Timestamp) - Stopping IIS on $SelectedServer...`r`n")

            $psCmd.BeginInvoke()
        } catch {
            $OutText.AppendText("$(Get-Timestamp) - An error occurred while invoking the command: $($_.Exception.Message)`r`n")
        } finally {
            Register-ObjectEvent -InputObject $psCmd -EventName InvocationStateChanged -Action {
                $Sender.Dispose()
                $Event.SourceEventArgs.Runspace.Close()
                $Event.SourceEventArgs.Runspace.Dispose()
            }
        }
})

# Function to handle the event when a server is selected in the $ServersListBox
function OnServerSelected {
    $SelectedServer = $ServersListBox.SelectedItem
    if ($null -ne $SelectedServer) {
        # Call the Open-PopulateListBoxRunspace function passing the selected server
        Open-PopulateListBoxRunspace -OutText $OutText -RestartsTabControl $RestartsTabControl -ServersListBox $ServersListBox -ServicesListBox $ServicesListBox -IISSitesListBox $IISSitesListBox -AppPoolsListBox $AppPoolsListBox
    }
}

# Function to show service status when a server is selected
function OnServiceSelected {
    $SelectedServer = $ServersListBox.SelectedItem
    if ($null -ne $SelectedServer) {
        $SelectedService = $ServicesListBox.SelectedItem
        if ($null -ne $SelectedService) {
            $psCmd = [PowerShell]::Create().AddScript({
                param (
                    [string]$SelectedServer,
                    [string]$SelectedService,
                    [System.Windows.Forms.TextBox]$OutText
                )

                function Get-Timestamp {
                    return Get-Date -Format "yyyy/MM/dd hh:mm:ss"
                }

                try {
                    $ServiceStatus = Invoke-Command -ComputerName $SelectedServer -ScriptBlock {
                        Get-Service -DisplayName $using:SelectedService
                    }
                    $OutText.AppendText("$(Get-Timestamp) - $SelectedService status: $($ServiceStatus.Status)`r`n")
                } catch {
                    $OutText.AppendText("$(Get-Timestamp) - Error retrieving service status for ${SelectedService}: $($_.Exception.Message)`r`n")
                }
            }).AddParameters(@{SelectedServer = $SelectedServer; SelectedService = $SelectedService; OutText = $OutText})

            $runspace = [RunspaceFactory]::CreateRunspace()
            $psCmd.Runspace = $runspace

            try {
                $runspace.Open()
                $OutText.AppendText("$(Get-Timestamp) - Retrieving $SelectedService status on $SelectedServer...`r`n")

                $psCmd.BeginInvoke()
            } catch {
                $OutText.AppendText("$(Get-Timestamp) - An error occurred while invoking the command: $($_.Exception.Message)`r`n")
            } finally {
                Register-ObjectEvent -InputObject $psCmd -EventName InvocationStateChanged -Action {
                    $Sender.Dispose()
                    $Event.SourceEventArgs.Runspace.Close()
                    $Event.SourceEventArgs.Runspace.Dispose()
                }
            }
        }
    }
}

# Function to show service status when an IIS site is selected
function OnIISSiteSelected {
    $SelectedServer = $ServersListBox.SelectedItem
    if ($null -ne $SelectedServer) {
        $SelectedIISSite = $IISsitesListBox.SelectedItem
        if ($null -ne $SelectedIISSite) {
            $psCmd = [PowerShell]::Create().AddScript({
                param (
                    [string]$SelectedServer,
                    [string]$SelectedIISSite,
                    [System.Windows.Forms.TextBox]$OutText
                )

                function Get-Timestamp {
                    return Get-Date -Format "yyyy/MM/dd hh:mm:ss"
                }

                try {
                    $IISSiteStatus = Invoke-Command -ComputerName $SelectedServer -ScriptBlock {
                        Import-Module WebAdministration
                        Get-Website -Name $using:SelectedIISSite
                    }
                    $OutText.AppendText("$(Get-Timestamp) - $SelectedIISSite status: $($IISSiteStatus.State)`r`n")
                } catch {
                    $OutText.AppendText("$(Get-Timestamp) - Error retrieving IIS site status for ${SelectedIISSite}: $($_.Exception.Message)`r`n")
                }
            }).AddParameters(@{SelectedServer = $SelectedServer; SelectedIISSite = $SelectedIISSite; OutText = $OutText})

            $runspace = [RunspaceFactory]::CreateRunspace()
            $psCmd.Runspace = $runspace

            try {
                $runspace.Open()
                $OutText.AppendText("$(Get-Timestamp) - Retrieving $SelectedIISSite status on $SelectedServer...`r`n")

                $psCmd.BeginInvoke()
            } catch {
                $OutText.AppendText("$(Get-Timestamp) - An error occurred while invoking the command: $($_.Exception.Message)`r`n")
            } finally {
                Register-ObjectEvent -InputObject $psCmd -EventName InvocationStateChanged -Action {
                    $Sender.Dispose()
                    $Event.SourceEventArgs.Runspace.Close()
                    $Event.SourceEventArgs.Runspace.Dispose()
                }
            }
        }
    }
}

# Function to show AppPool status when a server is selected
function OnAppPoolSelected {
    $SelectedServer = $ServersListBox.SelectedItem
    if ($null -ne $SelectedServer) {
        $SelectedAppPool = $AppPoolsListBox.SelectedItem
        if ($null -ne $SelectedAppPool) {
            $psCmd = [PowerShell]::Create().AddScript({
                param (
                    [string]$SelectedServer,
                    [string]$SelectedAppPool,
                    [System.Windows.Forms.TextBox]$OutText
                )

                function Get-Timestamp {
                    return Get-Date -Format "yyyy/MM/dd hh:mm:ss"
                }

                try {
                    $AppPoolStatus = Invoke-Command -ComputerName $SelectedServer -ScriptBlock {
                        Import-Module WebAdministration
                        Get-WebAppPoolState -Name $using:SelectedAppPool
                    }
                    $OutText.AppendText("$(Get-Timestamp) - ${SelectedAppPool} status: $($AppPoolStatus.Value)`r`n")
                } catch {
                    $OutText.AppendText("$(Get-Timestamp) - Error retrieving AppPool status for ${SelectedAppPool}: $($_.Exception.Message)`r`n")
                }
            }).AddParameters(@{SelectedServer = $SelectedServer; SelectedAppPool = $SelectedAppPool; OutText = $OutText})

            $runspace = [RunspaceFactory]::CreateRunspace()
            $psCmd.Runspace = $runspace

            try {
                $runspace.Open()
                $OutText.AppendText("$(Get-Timestamp) - Retrieving ${SelectedAppPool} status on $SelectedServer...`r`n")

                $psCmd.BeginInvoke()
            } catch {
                $OutText.AppendText("$(Get-Timestamp) - An error occurred while invoking the command: $($_.Exception.Message)`r`n")
            } finally {
                Register-ObjectEvent -InputObject $psCmd -EventName InvocationStateChanged -Action {
                    $Sender.Dispose()
                    $Event.SourceEventArgs.Runspace.Close()
                    $Event.SourceEventArgs.Runspace.Dispose()
                }
            }
        }
    }
}

# Add the event handler to the $ServersListBox
$ServersListBox.add_SelectedIndexChanged({ OnServerSelected })

# Store the script block in a variable
$OnServiceSelectedAction = { OnServiceSelected }

# Add the event using the script block
$ServicesListBox.add_SelectedIndexChanged($OnServiceSelectedAction)

# Add the event handler to the $IISSitesListBox
$IISSitesListBox.add_SelectedIndexChanged({ OnIISSiteSelected })

# Add the event handler to the $AppPoolsListBox
$AppPoolsListBox.add_SelectedIndexChanged({ OnAppPoolSelected })

# Event handler for TabControl's SelectedIndexChanged event
$RestartsTabControl_SelectedIndexChanged = {
    $ServicesListBox.ClearSelected()
    $IISSitesListBox.ClearSelected()
    $AppPoolsListBox.ClearSelected()
    $RestartsTabControl.SelectedTab.Text
    Open-PopulateListBoxRunspace -OutText $OutText -RestartsTabControl $RestartsTabControl -ServersListBox $ServersListBox -ServicesListBox $ServicesListBox -IISSitesListBox $IISSitesListBox -AppPoolsListBox $AppPoolsListBox

}

$RestartsTabControl.add_SelectedIndexChanged($RestartsTabControl_SelectedIndexChanged)
$RestartsTabControl.Controls.Add($ServicesTab)
$RestartsTabControl.Controls.Add($IISSitesTab)
$RestartsTabControl.Controls.Add($AppPoolsTab)
$AppListCombo.add_SelectedIndexChanged($AppListCombo_SelectedIndexChanged)
$ServersListBox.add_SelectedIndexChanged($ServersListBox_SelectedIndexChanged)

<#
? **********************************************************************************************************************
? END OF RESTARTS GUI
? **********************************************************************************************************************
#>
<#
? **********************************************************************************************************************
? START OF NSLOOKUP 
? **********************************************************************************************************************
#>

# Button to run Resolve-DnsName
$NSLookupButton = New-Object System.Windows.Forms.Button
$NSLookupButton.Location = New-Object System.Drawing.Size(300, 400)
$NSLookupButton.Width = 65
$NSLookupButton.BackColor = $global:DisabledBackColor
$NSLookupButton.ForeColor = $global:DisabledForeColor
$NSLookupButton.FlatStyle = "Popup"
$NSLookupButton.Text = "Get DNS"
$NSLookupButton.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
$NSLookupButton.Enabled = $false

# Text box for Resolve-DnsName
$script:NSLookupTextBox = New-Object System.Windows.Forms.TextBox
$script:NSLookupTextBox.Location = New-Object System.Drawing.Size(195, 400)
$script:NSLookupTextBox.Size = New-Object System.Drawing.Size(100, 20)
$script:NSLookupTextBox.Text = ''

# Label for Enter computer text box
$NSLookupLabel = New-Object System.Windows.Forms.Label
$NSLookupLabel.Location = New-Object System.Drawing.Size(195, 375)
$NSLookupLabel.Size = New-Object System.Drawing.Size(150, 20)
$NSLookupLabel.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Bold)
$NSLookupLabel.Text = 'NSlookup'

# If NSLookup text box is empty, disable the Test button
$script:NSLookupTextBox.Add_TextChanged({
    if ($script:NSLookupTextBox.Text.Length -eq 0) {
        $NSLookupButton.Enabled = $False
        $NSLookupButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
        $NSLookupButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
    }
    else {
        $NSLookupButton.Enabled = $True
        
        $backColor = Get-AppropriateColor -ColorType "BackColor"
        $foreColor = Get-AppropriateColor -ColorType "ForeColor"

        $NSLookupButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
        $NSLookupButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
    }
})

# Function to open a runspace and run Resolve-DnsName
function Open-NSLookupRunspace {
    param (
        [System.Windows.Forms.TextBox]$OutText,
        [System.Windows.Forms.TextBox]$NSLookupTextBox
    )

    try {
        $NSLookupRunspace = [runspacefactory]::CreateRunspace()
        $NSLookupRunspace.Open()

        $syncHash = @{}
        $syncHash.OutText = $OutText

        $NSLookupRunspace.SessionStateProxy.SetVariable("syncHash", $syncHash)

        $psCmd = [PowerShell]::Create().AddScript({
            function Get-Timestamp {
                return Get-Date -Format "yyyy/MM/dd hh:mm:ss"
            }

            $Pattern = "^(?=.{1,255}$)(?!-)[a-zA-Z0-9-]{1,63}(?<!-)(\.[a-zA-Z0-9-]{1,63})*$"

            if ($args[0] -match $Pattern) {
                try {
                    $SelectedObjects = Resolve-DnsName -Name $args[0] -ErrorAction Stop
                    foreach ($SelectedObject in $SelectedObjects) {
                        if ($SelectedObject -and $SelectedObject.IPAddress) {
                            $IPAddress = $SelectedObject.IPAddress -as [ipaddress]
                            $IPAddressType = if ($IPAddress.AddressFamily -eq 'InterNetworkV6') {'IPv6'} else {'IPv4'}
                            $syncHash.OutText.AppendText("$(Get-Timestamp) - $($args[0]) $IPAddressType = $($SelectedObject.IPAddress)`r`n")
                        }
                    }
                }
                catch {
                    $syncHash.OutText.AppendText("$(Get-Timestamp) - An error occurred while resolving the DNS name. Please ensure you're entering a valid hostname.`r`n")
                }
            }
            else {
                $syncHash.OutText.AppendText("$(Get-Timestamp) - Please ensure you're entering a valid hostname`r`n")
            }
        }).AddArgument($NSLookupTextBox.Text)

        $psCmd.Runspace = $NSLookupRunspace

        $null = $psCmd.BeginInvoke()

    } catch {
        $OutText.AppendText("$(Get-Date -Format "yyyy/MM/dd hh:mm:ss") - An error occurred: $($_.Exception.Message)`r`n")
    } finally {
        $NSLookupTextBox.Text = ''
        Register-ObjectEvent -InputObject $psCmd -EventName InvocationStateChanged -Action {
            $Sender.Runspace.Dispose()
            $Sender.Dispose()
        }
    }
}

# Event handler for the RunLookup button
$NSLookupButton.Add_Click({
    Open-NSLookupRunspace -OutText $OutText -NSLookupTextBox $script:NSLookupTextBox
})

# Even handler for the RunLookup text box
$script:NSLookupTextBox.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") {
        Open-NSLookupRunspace -OutText $OutText -NSLookupTextBox $script:NSLookupTextBox
        $_.SuppressKeyPress = $true
    }
})

<#
? **********************************************************************************************************************
? END OF NSLOOKUP
? **********************************************************************************************************************
#>
<#
? **********************************************************************************************************************
? START OF SERVER PING 
? **********************************************************************************************************************
#>

# Button for testing server connection
$ServerPingButton = New-Object System.Windows.Forms.Button
$ServerPingButton.Location = New-Object System.Drawing.Point(110, 400)
$ServerPingButton.Width = 45
$ServerPingButton.BackColor = $global:DisabledBackColor
$ServerPingButton.ForeColor = $global:DisabledForeColor
$ServerPingButton.FlatStyle = "Popup"
$ServerPingButton.Text = "Ping"
$ServerPingButton.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
$ServerPingButton.Enabled = $false

# Text box for testing server connection
$ServerPingTextBox = New-Object System.Windows.Forms.TextBox
$ServerPingTextBox.Location = New-Object System.Drawing.Size(5, 400)
$ServerPingTextBox.Size = New-Object System.Drawing.Size(100, 20)
$ServerPingTextBox.Text = ''
$ServerPingTextBox.ShortcutsEnabled = $True

# Label for testing server connection text box
$ServerPingLabel = New-Object System.Windows.Forms.Label
$ServerPingLabel.Location = New-Object System.Drawing.Size(5, 375)
$ServerPingLabel.Size = New-Object System.Drawing.Size(150, 20)
$ServerPingLabel.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Bold)
$ServerPingLabel.Text = "Ping a Server"

# If server ping text box is empty, disable the Test button
$ServerPingTextBox.Add_TextChanged({
    if ($ServerPingTextBox.Text.Length -eq 0) {
        $ServerPingButton.Enabled = $False
        $ServerPingButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
        $ServerPingButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
    }
    else {
        $ServerPingButton.Enabled = $True
		$backColor = Get-AppropriateColor -ColorType "BackColor"
        $foreColor = Get-AppropriateColor -ColorType "ForeColor"
		
		$ServerPingButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
        $ServerPingButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
    }
})

# Function for opening a runspace and running Test-Connection
function Open-ServerPingRunspace {
    param (
        [System.Windows.Forms.TextBox]$OutText,
        [System.Windows.Forms.TextBox]$ServerPingTextBox
    )
    
    try {
        $ServerPingRunspace = [runspacefactory]::CreateRunspace()
        $ServerPingRunspace.Open()
        
        $syncHash = @{}
        $syncHash.OutText = $OutText

        $ServerPingRunspace.SessionStateProxy.SetVariable("syncHash", $syncHash)

        $psCmd = [PowerShell]::Create().AddScript({
            function Get-Timestamp {
                return Get-Date -Format "yyyy/MM/dd hh:mm:ss"
            }

            $syncHash.OutText.AppendText("$(Get-Timestamp) - Testing connection to $($args[0])...`r`n")
            $PingResult = Test-Connection -ComputerName $args[0] -Quiet
            if ($PingResult) {
                $syncHash.OutText.AppendText("$(Get-Timestamp) - Connection to $($args[0]) successful.`r`n")
            }
            else {
                $syncHash.OutText.AppendText("$(Get-Timestamp) - Connection to $($args[0]) failed.`r`n")
            }
        }).AddArgument($ServerPingTextBox.Text)

        $psCmd.Runspace = $ServerPingRunspace
        
        $null = $psCmd.BeginInvoke()

    } catch {
        $OutText.AppendText("$(Get-Date -Format "yyyy/MM/dd hh:mm:ss") - An error occurred: $($_.Exception.Message)`r`n")
    } finally {
        $ServerPingTextBox.Text = ''
        Register-ObjectEvent -InputObject $psCmd -EventName InvocationStateChanged -Action {
            $Sender.Runspace.Dispose()
            $Sender.Dispose()
        }
    }
}

# Event handler for testing server connection button
$ServerPingButton.Add_Click({
    Open-ServerPingRunspace -OutText $OutText -ServerPingTextBox $ServerPingTextBox
})

# Server Ping text box Enter key logic
$ServerPingTextBox.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") {
        Open-ServerPingRunspace -OutText $OutText -ServerPingTextBox $ServerPingTextBox
        $_.SuppressKeyPress = $true
    }
})

<#
? **********************************************************************************************************************
? END OF SERVER PING 
? **********************************************************************************************************************
#>
<#
? **********************************************************************************************************************
? START OF REVERSE IP LOOKUP
? **********************************************************************************************************************
#>

# Text box for Reverse IP Lookup
$ReverseIPTextBox = New-Object System.Windows.Forms.TextBox
$ReverseIPTextBox.Location = New-Object System.Drawing.Size(400, 400)
$ReverseIPTextBox.Size = New-Object System.Drawing.Size(100, 20)
$ReverseIPTextBox.Text = ''
$ReverseIPTextBox.ShortcutsEnabled = $True

# Label for Reverse IP Lookup text box
$ReverseIPLabel = New-Object System.Windows.Forms.Label
$ReverseIPLabel.Location = New-Object System.Drawing.Size(400, 375)
$ReverseIPLabel.Size = New-Object System.Drawing.Size(150, 20)
$ReverseIPLabel.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Bold)
$ReverseIPLabel.Text = "IP Lookup"

# Button for Reverse IP Lookup
$ReverseIPButton = New-Object System.Windows.Forms.Button
$ReverseIPButton.Location = New-Object System.Drawing.Point(510, 400)
$ReverseIPButton.Width = 65
$ReverseIPButton.BackColor = $global:DisabledBackColor
$ReverseIPButton.ForeColor = $global:DisabledForeColor
$ReverseIPButton.FlatStyle = "Popup"
$ReverseIPButton.Text = "Get IP"
$ReverseIPButton.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
$ReverseIPButton.Enabled = $false

# Event handler for enabling the ReverseIPButton
$ReverseIPTextBox.Add_TextChanged({
    if ($ReverseIPTextBox.Text.Length -gt 0) {
        $ReverseIPButton.Enabled = $true

        $backColor = Get-AppropriateColor -ColorType "BackColor"
        $foreColor = Get-AppropriateColor -ColorType "ForeColor"

        $script:ReverseIPButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
        $script:ReverseIPButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
    }
    else {
        $ReverseIPButton.Enabled = $false
        $ReverseIPButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
        $ReverseIPButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
    }
})

# Function to perform a reverse IP lookup
function Get-IPAddress {
    param (
        [string]$IPAddress
    )

    try {
        $hostEntry = [System.Net.Dns]::GetHostEntry($IPAddress)
        return $hostEntry.HostName
    }
    catch {
        $OutText.AppendText("$(Get-Timestamp) - An error occurred while performing a reverse IP lookup: $($_.Exception.Message)")
        return $null
    }
}

# Button click event handler for Reverse IP Lookup
$ReverseIPButton.add_Click({
    $ip = $ReverseIPTextBox.Text

    # Validate IP Address
    $isValidIP = [ipaddress]::TryParse($ip, [ref]$null)

    if ($isValidIP) {
        $hostname = Get-IPAddress -IPAddress $ip

        if ($hostname) {
            $OutText.AppendText("$(Get-Timestamp) - The hostname for IP address $ip is $hostname.`r`n")
        } else {
            $OutText.AppendText("$(Get-Timestamp) - Could not find the hostname for IP address $ip.`r`n")
        }
    }
    else {
        $OutText.AppendText("$(Get-Timestamp) - The provided IP address ($ip) is invalid. Please enter a valid IP address.`r`n")
    }
})

<#
? **********************************************************************************************************************
? END OF REVERSE IP LOOKUP
? **********************************************************************************************************************
#>
<#
? **********************************************************************************************************************
? START OF PROD SUPPORT TOOL 

? **********************************************************************************************************************
#>

# PST Environment variables
$RemoteSupportTool = $ConfigValues.RemoteSupportTool
$LocalSupportTool = $ConfigValues.LocalSupportTool.Replace("{USERPROFILE}", $userProfilePath)
$LocalConfigs = $ConfigValues.LocalConfigs.Replace("{USERPROFILE}", $userProfilePath)

# Combobox for environment selection
$PSTCombo = New-Object System.Windows.Forms.ComboBox
$PSTCombo.Location = New-Object System.Drawing.Point(5,65)
$PSTCombo.Size = New-Object System.Drawing.Size(200, 200)
$PSTCombo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$PSTCombo.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
@('QA', 'Stage', 'Production') | ForEach-Object { [void]$PSTCombo.Items.Add($_) }
$PSTCombo.BackColor = $OutText.BackColor

# Label for PST combo box
$PSTComboLabel = New-Object System.Windows.Forms.Label
$PSTComboLabel.Location = New-Object System.Drawing.Size(5, 40)
$PSTComboLabel.Size = New-Object System.Drawing.Size(150, 20)
$PSTComboLabel.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Bold)
$PSTComboLabel.Text = "Production Support Tool"

# Button for switching environment
$SelectEnvButton = New-Object System.Windows.Forms.Button
$SelectEnvButton.Location = New-Object System.Drawing.Point(210, 35)
$SelectEnvButton.Width = 150
$SelectEnvButton.BackColor = $global:DisabledBackColor
$SelectEnvButton.ForeColor = $global:DisabledForeColor
$SelectEnvButton.FlatStyle = "Popup"
$SelectEnvButton.Text = "Select Environment"
$SelectEnvButton.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
$SelectEnvButton.Enabled = $false

# Button for resetting environment
$ResetEnvButton = New-Object System.Windows.Forms.Button
$ResetEnvButton.Location = New-Object System.Drawing.Point(210, 65)
$ResetEnvButton.Width = 150
$ResetEnvButton.FlatStyle = "Popup"
$ResetEnvButton.BackColor = $global:DisabledBackColor
$ResetEnvButton.ForeColor = $global:DisabledForeColor
$ResetEnvButton.Text = "Reset Environment"
$ResetEnvButton.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
$ResetEnvButton.Enabled = $false

# Button for running the PST
$RunPSTButton = New-Object System.Windows.Forms.Button
$RunPSTButton.Location = New-Object System.Drawing.Point(210, 95)
$RunPSTButton.Width = 150
$RunPSTButton.FlatStyle = "Popup"
$RunPSTButton.BackColor = $global:DisabledBackColor
$RunPSTButton.ForeColor = $global:DisabledForeColor
$RunPSTButton.Text = "Run Prod Support Tool"
$RunPSTButton.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
$RunPSTButton.Enabled = $false

# Button for refreshing the PST files
$RefreshPSTButton = New-Object System.Windows.Forms.Button
$RefreshPSTButton.Location = New-Object System.Drawing.Point(5, 95)
$RefreshPSTButton.Width = 150
$RefreshPSTButton.FlatStyle = "Popup"
$RefreshPSTButton.Text = $RefreshButtonText
$RefreshPSTButton.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
$RefreshPSTButton.Enabled = $true

# Separator line under PST
$PSTSeparator = New-Object System.Windows.Forms.Label
$PSTSeparator.Location = New-Object System.Drawing.Size(0, 170)
$PSTSeparator.Size = New-Object System.Drawing.Size(1000, 2)

# Get the last write time of the remote and local PST folders
if (Test-Path $RemoteSupportTool) {
    $RemoteSupportToolFolderDate = Get-ChildItem $RemoteSupportTool -Exclude *.txt -Recurse | ForEach-Object {$_.LastWriteTime} | Sort-Object -Descending | Select-Object -First 1
} else {
    $OutText.AppendText("$(Get-Timestamp) - Remote Prod Support Tool folder not found.`r`n")
}

# Get the last write time and count of the local PST folder
if (Test-Path $LocalSupportTool) {
    $LocalSupportToolFolderDate = Get-ChildItem $LocalSupportTool -Exclude *.txt -Recurse | ForEach-Object {$_.LastWriteTime} | Sort-Object -Descending | Select-Object -First 1
    $LocalSupportToolFolderCount = Get-ChildItem $LocalSupportTool | Measure-Object
} else {
    $OutText.AppendText("$(Get-Timestamp) - Local Prod Support Tool folder not found.`r`n")
}

$RefreshButtonText = if ($LocalSupportToolFolderCount -eq 0) { "Import PST Files" } else { "Refresh PST Files" }

# Function for resetting the selected environment
Function Reset-Environment {
    Remove-Item -Path "$LocalSupportTool\ProductionSupportTool.exe.config"
    Rename-Item "$LocalSupportTool\ProductionSupportTool.exe.config.old" -NewName "$LocalSupportTool\ProductionSupportTool.exe.config"
    $OutText.AppendText("$(Get-Timestamp) - Environment has been reset`r`n")
    $ResetEnvButton.Enabled = $false
    $RunPSTButton.Enabled = $false
    $RunPSTButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
    $RunPSTButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
    $ResetEnvButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
    $ResetEnvButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
    $SelectEnvButton.Enabled = $true

    $backColor = Get-AppropriateColor -ColorType "BackColor"
    $foreColor = Get-AppropriateColor -ColorType "ForeColor"

    $SelectEnvButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
    $SelectEnvButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)

    $PSTCombo.Enabled = $true
}

# Event handler for enabling the Select Environment button
$PSTCombo.Add_SelectedIndexChanged({
    if ($PSTCombo.SelectedItem -ne $null) {
        $SelectEnvButton.Enabled = $true

        $backColor = Get-AppropriateColor -ColorType "BackColor"
        $foreColor = Get-AppropriateColor -ColorType "ForeColor"

        $script:SelectEnvButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
        $script:SelectEnvButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
    }
    else {
        $SelectEnvButton.Enabled = $false
        $SelectEnvButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
        $SelectEnvButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
    }
})

# Event handler for the reset environment button
$ResetEnvButton.Add_Click({
    Reset-Environment
})

# Event handler for the environment switch button
$SelectEnvButton.Add_Click({
    If (Test-Path "$LocalSupportTool\ProductionSupportTool.exe.config.old") {
        $OutText.AppendText("$(Get-Timestamp) - Existing config found; cleaning up`r`n")   
        Reset-Environment
    }
    else {
        $ResetEnvButton.Enabled = $false
		$ResetEnvButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
        $ResetEnvButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
    }
    # Adding code to a button, so that when clicked, it switches environments            
    $Environment = $PSTCombo.SelectedItem
        switch ($Environment) {
            "QA" {
                $OutText.AppendText("$(Get-Timestamp) - Entering QA environment`r`n")
                $RunningConfig = "$LocalConfigs\ProductionSupportTool.exe_QA.config"
            }
            "Stage" {
                $OutText.AppendText("$(Get-Timestamp) - Entering Staging environment`r`n")
                $RunningConfig = "$LocalConfigs\ProductionSupportTool.exe_Stage.config"
            }
            "Production" {
                $OutText.AppendText("$(Get-Timestamp) - Entering Production environment`r`n")
                $RunningConfig = "$LocalConfigs\ProductionSupportTool.exe_Prod.config"
            }
            # Behavior if no option is selected
            Default {
                $OutText.AppendText("$(Get-Timestamp) - Please make a valid selection or reset`r`n")
                throw "No selection made"
            }
        }
        Start-Sleep -Seconds 1
        # Rename Current Running config and Copy configuration file for correct environment 
        Rename-Item "$LocalSupportTool\ProductionSupportTool.exe.config" -NewName "$LocalSupportTool\ProductionSupportTool.exe.config.old"
        Copy-Item $RunningConfig -Destination "$LocalSupportTool\ProductionSupportTool.exe.config"
        $OutText.AppendText("$(Get-Timestamp) - You are ready to run in the $Environment environment`r`n")
        $PSTCombo.Enabled = $false
        $ResetEnvButton.Enabled = $true
        $RunPSTButton.Enabled = $true
		
		$backColor = Get-AppropriateColor -ColorType "BackColor"
        $foreColor = Get-AppropriateColor -ColorType "ForeColor"
		
		$ResetEnvButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
        $ResetEnvButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
        $RunPSTButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
        $RunPSTButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
		
        $SelectEnvButton.Enabled = $false

        $SelectEnvButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
        $SelectEnvButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
})

# Logic to determine if the Refresh PST button should say "Import PST Files" or "Refresh PST Files"
if (Test-Path $LocalSupportTool) {
    $RefreshPSTButton.Text = "Refresh PST Files"
} else {
    $RefreshPSTButton.Text = "Import PST Files"
}

# Event handler for the run PST button
$RunPSTButton.Add_Click({
    $OutText.AppendText("$(Get-Timestamp) - Launching Prod Support Tool`r`n")
    Start-Process -FilePath "$LocalSupportTool\ProductionSupportTool.exe"
})

# Function for opening a runspace and running Refresh/Import PST files
function Update-PSTFiles {
    $RefreshPSTRunspace = [runspacefactory]::CreateRunspace()
    $RefreshPSTRunspace.Open()
    $syncHash = @{}
    $syncHash.OutText = $OutText
    $RefreshPSTRunspace.SessionStateProxy.SetVariable("syncHash", $syncHash)

    $psCmd = [PowerShell]::Create().AddScript({
        # Initilize the timestamp function in the runspace
        function Get-Timestamp {
            return Get-Date -Format "yyyy/MM/dd hh:mm:ss"
        }

        $LocalSupportToolFolderCount = (Get-ChildItem $LocalSupportTool -Exclude *.txt -Recurse).Count
    
    if ($LocalSupportToolFolderCount -eq 0) {
        $RefreshPSTButton.Text = "Import PST Files"
    } else {
        $RefreshPSTButton.Text = "Refresh PST Files"
    }

    $RemoteSupportToolFolderDate = Get-ChildItem $RemoteSupportTool -Exclude *.txt -Recurse | ForEach-Object {$_.LastWriteTime} | Sort-Object -Descending | Select-Object -First 1
    $LocalSupportToolFolderDate = Get-ChildItem $LocalSupportTool -Exclude *.txt -Recurse | ForEach-Object {$_.LastWriteTime} | Sort-Object -Descending | Select-Object -First 1

    if ($RemoteSupportToolFolderDate -gt $LocalSupportToolFolderDate) {
        $synchash.OutText.AppendText("$(Get-Timestamp) - Remote Prod Support Tool folder is newer than local folder.`r`n")
        $synchash.OutText.AppendText("$(Get-Timestamp) - Copying files from remote folder to local folder.`r`n")

        Copy-Item -Path $RemoteSupportTool* -Destination $LocalSupportTool -Recurse -Force

        $synchash.OutText.AppendText("$(Get-Timestamp) - Files successfully copied.`r`n")
        $synchash.OutText.AppendText("$(Get-Timestamp) - Local Prod Support Tool folder is now up to date.`r`n")
    } else {
        $synchash.OutText.AppendText("$(Get-Timestamp) - Local Prod Support Tool folder is up to date.`r`n")
    }
    })

    $psCmd.Runspace = $RefreshPSTRunspace
    $null = $psCmd.BeginInvoke()
}

# Event handler for importing/refreshing the PST files
$RefreshPSTButton.Add_Click({
    Update-PSTFiles
})

<#
? **********************************************************************************************************************
? END OF PROD SUPPORT TOOL
? **********************************************************************************************************************
#>
<#
? **********************************************************************************************************************
? START OF ADD LENDER TO LFP SERVICES
? **********************************************************************************************************************
#>

# Tooltips for LFP popup form
$LFPToolTip = New-Object System.Windows.Forms.ToolTip
$LFPToolTip.InitialDelay = 100

# Button for launching LFP wizard
$LaunchLFPWizardButton = New-Object System.Windows.Forms.Button
$LaunchLFPWizardButton.Location = New-Object System.Drawing.Point(400, 65)
$LaunchLFPWizardButton.Width = 150
$LaunchLFPWizardButton.FlatStyle = "Popup"
$LaunchLFPWizardButton.Text = "Launch LFP Wizard"
$LaunchLFPWizardButton.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
$LaunchLFPWizardButton.Enabled = $true

# Label for LFP wizard button
$LaunchLFPWizardLabel = New-Object System.Windows.Forms.Label
$LaunchLFPWizardLabel.Location = New-Object System.Drawing.Size(413, 40)
$LaunchLFPWizardLabel.Size = New-Object System.Drawing.Size(150, 20)
$LaunchLFPWizardLabel.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Bold)
$LaunchLFPWizardLabel.Text = "Lender LFP Services"

# Click event handler for launching the LFP wizard
$LaunchLFPWizardButton.Add_Click({
    $OutText.AppendText("$(Get-Timestamp) - Launching LFP Wizard...`r`n")

    # New popup window
    $script:LenderLFPPopup = New-Object System.Windows.Forms.Form
    $script:LenderLFPPopup.Text = "Add Lender to LFP Services"
    $script:LenderLFPPopup.Size = New-Object System.Drawing.Size(300, 350)
    $script:LenderLFPPopup.ShowInTaskbar = $True
    $script:LenderLFPPopup.KeyPreview = $True
    $script:LenderLFPPopup.AutoSize = $True
    $script:LenderLFPPopup.FormBorderStyle = 'Fixed3D'
    $script:LenderLFPPopup.MaximizeBox = $False
    $script:LenderLFPPopup.MinimizeBox = $True
    $script:LenderLFPPopup.ControlBox = $True
    $script:LenderLFPPopup.Icon = $Icon
    $script:LenderLFPPopup.StartPosition = "CenterScreen"
    $global:IsLenderLFPPopupActive = $true

    # Combo box for selecting the environment where lender will be added
    $script:LenderLFPCombo = New-Object System.Windows.Forms.ComboBox
    $script:LenderLFPCombo.Location = New-Object System.Drawing.Point(5, 65)
    $script:LenderLFPCombo.Size = New-Object System.Drawing.Size(200, 200)
    $script:LenderLFPCombo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $script:LenderLFPCombo.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
    @('QA', 'Stage', 'Production') | ForEach-Object { [void]$script:LenderLFPCombo.Items.Add($_) }

    # Label for Lender LFP combo box
    $script:LenderLFPComboLabel = New-Object System.Windows.Forms.Label
    $script:LenderLFPComboLabel.Location = New-Object System.Drawing.Size(5, 40)
    $script:LenderLFPComboLabel.Size = New-Object System.Drawing.Size(150, 20)
    $script:LenderLFPComboLabel.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Bold)
    $script:LenderLFPComboLabel.Text = "Select Environment"

    # Text box for entering the lender Id
    $script:LenderLFPIdTextBox = New-Object System.Windows.Forms.TextBox
    $script:LenderLFPIdTextBox.Location = New-Object System.Drawing.Size(5, 125)
    $script:LenderLFPIdTextBox.Size = New-Object System.Drawing.Size(150, 20)
    $script:LenderLFPIdTextBox.Text = ''
    $script:LenderLFPIdTextBox.ShortcutsEnabled = $True

    # Label for Lender LFP text box
    $script:LenderLFPTextBoxLabel = New-Object System.Windows.Forms.Label
    $script:LenderLFPTextBoxLabel.Location = New-Object System.Drawing.Size(5, 100)
    $script:LenderLFPTextBoxLabel.Size = New-Object System.Drawing.Size(150, 20)
    $script:LenderLFPTextBoxLabel.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Bold)
    $script:LenderLFPTextBoxLabel.Text = "Lender ID"

    # Text box for entering ticket number
    $script:LenderLFPTicketTextBox = New-Object System.Windows.Forms.TextBox
    $script:LenderLFPTicketTextBox.Location = New-Object System.Drawing.Size(5, 185)
    $script:LenderLFPTicketTextBox.Size = New-Object System.Drawing.Size(150, 20)
    $script:LenderLFPTicketTextBox.Text = ''
    $script:LenderLFPTicketTextBox.ShortcutsEnabled = $True

    # Label for Lender LFP ticket text box
    $script:LenderLFPTicketTextBoxLabel = New-Object System.Windows.Forms.Label
    $script:LenderLFPTicketTextBoxLabel.Location = New-Object System.Drawing.Size(5, 160)
    $script:LenderLFPTicketTextBoxLabel.Size = New-Object System.Drawing.Size(150, 20)
    $script:LenderLFPTicketTextBoxLabel.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Bold)
    $script:LenderLFPTicketTextBoxLabel.Text = "Jira Ticket Number"

    # Button for adding lender to LFP services
    $script:AddLenderLFPButton = New-Object System.Windows.Forms.Button
    $script:AddLenderLFPButton.Location = New-Object System.Drawing.Point(5, 215)
    $script:AddLenderLFPButton.Width = 100
    $script:AddLenderLFPButton.BackColor = $global:DisabledBackColor
    $script:AddLenderLFPButton.ForeColor = $global:DisabledForeColor
    $script:AddLenderLFPButton.FlatStyle = "Popup"
    $script:AddLenderLFPButton.Text = "Add Lender"
    $script:AddLenderLFPButton.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
    $script:AddLenderLFPButton.Enabled = $false

    # Check if DefaultUserTheme has a value or is null
    if ($null -ne $ConfigValues.DefaultUserTheme -and $ConfigValues.DefaultUserTheme -ne "") {
        # Get theme
        $SelectedTheme = $ColorTheme.PSObject.Properties | Where-Object { $_.Value.$($ConfigValues.DefaultUserTheme) } | Select-Object -ExpandProperty Name
        $themeColors = $ColorTheme.$SelectedTheme.$($ConfigValues.DefaultUserTheme)
        
        $script:LenderLFPPopup.BackColor = $themeColors.BackColor
        $script:LenderLFPPopup.ForeColor = $themeColors.ForeColor
        $script:LenderLFPCombo.BackColor = $OutText.BackColor
        $script:LenderLFPComboLabel.ForeColor = $themeColors.ForeColor
        $script:LenderLFPTextBoxLabel.ForeColor = $themeColors.ForeColor
        $script:LenderLFPTicketTextBoxLabel.ForeColor = $themeColors.ForeColor
    }

    # Add tooltips if the user has them enabled
    if ($ConfigValues.HoverToolTips -eq "Enabled" -or $null -eq $ConfigValues.HoverToolTips) {
        $LFPToolTip.SetToolTip($script:LenderLFPCombo, "Select an environment to add lender")
        $LFPToolTip.SetToolTip($script:LenderLFPIdTextBox, "Enter a numeric lender Id, i.e. 6969")
        $LFPToolTip.SetToolTip($script:LenderLFPTicketTextBox, "Enter the ticket in Jira format, i.e. AIH-42069")
        $LFPToolTip.SetToolTip($script:AddLenderLFPButton, "Click to add lender to LFP services")
    }

    # Call the function to set form icon
    Set-FormIcon

    # Event handler for the Lender LFP combo box
    $script:LenderLFPCombo.Add_SelectedIndexChanged({
        Enable-AddLenderButton
    })

    # Event handler for Lender LFP text box
    $script:LenderLFPIdTextBox.Add_TextChanged({
        Enable-AddLenderButton
    })

    # Event handler for Lender LFP ticket text box
    $script:LenderLFPTicketTextBox.Add_TextChanged({
        Enable-AddLenderButton
    })

    # Button click event for adding lender to LFP services
    $script:AddLenderLFPButton.add_Click({
        $AddLenderScript = $ConfigValues.AddLenderToLFPServicesScript.Replace("{USERPROFILE}", $userProfilePath)
        if ($script:LenderLFPIdTextBox.Text -match '^\d+$') {
            if ($script:LenderLFPTicketTextBox.Text -match '^[A-Za-z]{3}-\d{5}$') {
                if (Test-Path $AddLenderScript){
                    try {
                        & $AddLenderScript -LenderId $script:LenderLFPIdTextBox.Text -TicketNumber $script:LenderLFPTicketTextBox.Text -Environment $script:LenderLFPCombo.SelectedItem -OutTextControl $OutText -TimestampFunction ${function:Get-Timestamp}
                    }
                    catch {
                        $OutText.AppendText("$(Get-Timestamp) - An error occurred while executing the AddLenderScript: $_`r`n")
                    }            
                } else {
                    $OutText.AppendText("$(Get-Timestamp) - AddLendertoLFPServices.ps1 not found`r`n")
                }
            }
            else {
                $OutText.AppendText("$(Get-Timestamp) - Please enter the ticket in Jira format, i.e. AIH-42069`r`n")
            }
        }
        else {
            $OutText.AppendText("$(Get-Timestamp) - Please enter a valid numeric Lender Id, i.e. 6969`r`n")
        }
    })

    # Click event for form close; sets global variable to false
    $script:LenderLFPPopup.Add_FormClosed({
        $global:IsLenderLFPPopupActive = $false
    })

    $script:LenderLFPPopup.Controls.Add($script:LenderLFPCombo)
    $script:LenderLFPPopup.Controls.Add($script:LenderLFPComboLabel)
    $script:LenderLFPPopup.Controls.Add($script:LenderLFPIdTextBox)
    $script:LenderLFPPopup.Controls.Add($script:LenderLFPTextBoxLabel)
    $script:LenderLFPPopup.Controls.Add($script:LenderLFPTicketTextBox)
    $script:LenderLFPPopup.Controls.Add($script:LenderLFPTicketTextBoxLabel)
    $script:LenderLFPPopup.Controls.Add($script:AddLenderLFPButton)
    $script:LenderLFPPopup.Show()
})

<#
? **********************************************************************************************************************
? END OF ADD LENDER TO LFP SERVICES
? **********************************************************************************************************************
#>
<#
? **********************************************************************************************************************
? START OF PASSWORD MANAGER
? **********************************************************************************************************************
#>

# Text box for setting passwords
$PWTextBox = New-Object System.Windows.Forms.TextBox
$PWTextBox.Location = New-Object System.Drawing.Size(5, 250)
$PWTextBox.Size = New-Object System.Drawing.Size(150, 20)
$PWTextBox.Text = ''
$PWTextBox.ShortcutsEnabled = $True
$PWTextBox.PasswordChar = '*'

# Label for setting passwords text box
$PWTextBoxLabel = New-Object System.Windows.Forms.Label
$PWTextBoxLabel.Location = New-Object System.Drawing.Size(5, 225)
$PWTextBoxLabel.Size = New-Object System.Drawing.Size(150, 20)
$PWTextBoxLabel.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Bold)
$PWTextBoxLabel.Text = 'Enter Secure Password'

# Button for setting your own password
$SetPWButton = New-Object System.Windows.Forms.Button
$SetPWButton.Location = New-Object System.Drawing.Point(160, 250)
$SetPWButton.Width = 100
$SetPWButton.BackColor = $global:DisabledBackColor
$SetPWButton.ForeColor = $global:DisabledForeColor
$SetPWButton.FlatStyle = "Popup"
$SetPWButton.Text = "Set Your PW"
$SetPWButton.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
$SetPWButton.Enabled = $false

# Button for setting your alternate password
$AltSetPWButton = New-Object System.Windows.Forms.Button
$AltSetPWButton.Location = New-Object System.Drawing.Point(160, 275)
$AltSetPWButton.Width = 100
$AltSetPWButton.BackColor = $global:DisabledBackColor
$AltSetPWButton.ForeColor = $global:DisabledForeColor
$AltSetPWButton.FlatStyle = "Popup"
$AltSetPWButton.Text = "Set Alt PW"
$AltSetPWButton.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
$AltSetPWButton.Enabled = $false

# Button for getting your own password
$GetPWButton = New-Object System.Windows.Forms.Button
$GetPWButton.Location = New-Object System.Drawing.Point(160, 250)
$GetPWButton.Width = 100
$GetPWButton.BackColor = $global:DisabledBackColor
$GetPWButton.ForeColor = $global:DisabledForeColor
$GetPWButton.FlatStyle = "Popup"
$GetPWButton.Text = "Get Your PW"
$GetPWButton.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
$GetPWButton.Enabled = $false

# Button for getting your alternate password
$AltGetPWButton = New-Object System.Windows.Forms.Button
$AltGetPWButton.Location = New-Object System.Drawing.Point(160, 275)
$AltGetPWButton.Width = 100
$AltGetPWButton.BackColor = $global:DisabledBackColor
$AltGetPWButton.ForeColor = $global:DisabledForeColor
$AltGetPWButton.FlatStyle = "Popup"
$AltGetPWButton.Text = "Get Alt PW"
$AltGetPWButton.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
$AltGetPWButton.Enabled = $false

# Button for clearing your own password
$ClearPWButton = New-Object System.Windows.Forms.Button
$ClearPWButton.Location = New-Object System.Drawing.Point(265, 250)
$ClearPWButton.Width = 100
$ClearPWButton.BackColor = $global:DisabledBackColor
$ClearPWButton.ForeColor = $global:DisabledForeColor
$ClearPWButton.FlatStyle = "Popup"
$ClearPWButton.Text = "Clear Your PW"
$ClearPWButton.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
$ClearPWButton.Enabled = $false

# Button for clearing your alternate password
$AltClearPWButton = New-Object System.Windows.Forms.Button
$AltClearPWButton.Location = New-Object System.Drawing.Point(265, 275)
$AltClearPWButton.Width = 100
$AltClearPWButton.BackColor = $global:DisabledBackColor
$AltClearPWButton.ForeColor = $global:DisabledForeColor
$AltClearPWButton.FlatStyle = "Popup"
$AltClearPWButton.Text = "Clear Alt PW"
$AltClearPWButton.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
$AltClearPWButton.Enabled = $false

# Button for generating a random password
$GenPWButton = New-Object System.Windows.Forms.Button
$GenPWButton.Location = New-Object System.Drawing.Point(425, 250)
$GenPWButton.Width = 100
$GenPWButton.FlatStyle = "Popup"
$GenPWButton.Text = "Generate"
$GenPWButton.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
$GenPWButton.Enabled = $True

# Label for generating a random password
$GenPWLabel = New-Object System.Windows.Forms.Label
$GenPWLabel.Location = New-Object System.Drawing.Size(417, 225)
$GenPWLabel.Size = New-Object System.Drawing.Size(150, 20)
$GenPWLabel.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Bold)
$GenPWLabel.Text = 'Generate Password'

# Separator line for password manager
$PWManagerSeparator = New-Object System.Windows.Forms.Label
$PWManagerSeparator.Location = New-Object System.Drawing.Size(0, 360)
$PWManagerSeparator.Size = New-Object System.Drawing.Size(1000, 2)

# Logic for enabling the Set PW buttons if text box is populated
$PWTextBox.Add_TextChanged({
    if ($PWTextBox.Text.Length -eq 0) {
        $SetPWButton.Enabled = $false
        $SetPWButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
        $SetPWButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
        $AltSetPWButton.Enabled = $false
        $AltSetPWButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
        $AltSetPWButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
    }
    else {
        $SetPWButton.Enabled = $true
		$AltSetPWButton.Enabled = $true
		
		$backColor = Get-AppropriateColor -ColorType "BackColor"
        $foreColor = Get-AppropriateColor -ColorType "ForeColor"
		
        $SetPWButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
        $SetPWButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
        $AltSetPWButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
        $AltSetPWButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
    }
})

# Function for setting your own password
function Set-MyPassword{
    $PWTextBox.Text | Set-Clipboard
    $global:SecurePW = $PWTextBox.Text | ConvertTo-SecureString -AsPlainText -Force
    $OutText.AppendText("$(Get-Timestamp) - Your password has been set and copied to the clipboard.`r`n")
    $PWTextBox.Text = ''
    $SetPWButton.Enabled = $False
    $SetPWButton.Visible = $False
    $GetPWButton.Enabled = $True
    $GetPWButton.Visible = $True
    $backColor = Get-AppropriateColor -ColorType "BackColor"
    $foreColor = Get-AppropriateColor -ColorType "ForeColor"
    $GetPWButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
    $GetPWButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
    $ClearPWButton.Enabled = $True
    $ClearPWButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
    $ClearPWButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
    return $global:SecurePW
}

# Function for setting your alternate password
function Set-AltPassword{
    $PWTextBox.Text | Set-Clipboard
    $global:AltSecurePW = $PWTextBox.Text | ConvertTo-SecureString -AsPlainText -Force
    $OutText.AppendText("$(Get-Timestamp) - Your alternate password has been set and copied to the clipboard.`r`n")
    $PWTextBox.Text = ''
    $AltSetPWButton.Enabled = $False
    $AltSetPWButton.Visible = $False
    $AltGetPWButton.Enabled = $True
    $AltGetPWButton.Visible = $True
    $backColor = Get-AppropriateColor -ColorType "BackColor"
    $foreColor = Get-AppropriateColor -ColorType "ForeColor"
    $AltGetPWButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
    $AltGetPWButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
    $AltClearPWButton.Enabled = $True
    $AltClearPWButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
    $AltClearPWButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
    return $global:AltSecurePW
}

$SetPWButton.Add_Click({
    Set-MyPassword
})

$AltSetPWButton.Add_Click({
    Set-AltPassword
})

# Clear my password button logic
$ClearPWButton.Add_Click({
    $PWTextBox.Text = ''
    $global:SecurePW = $null
    Set-Clipboard -Value ''
    $ClearPWButton.Enabled = $False
    $ClearPWButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
    $ClearPWButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
    $SetPWButton.Enabled = $False
    $SetPWButton.Visible = $True
    $SetPWButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
    $SetPWButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
    $GetPWButton.Enabled = $False
    $GetPWButton.Visible = $False
    $OutText.AppendText("$(Get-Timestamp) - Your password has been cleared.`r`n")
})

# Clear my alternate password button logic
$AltClearPWButton.Add_Click({
    $PWTextBox.Text = ''
    $global:AltSecurePW = $null
    Set-Clipboard -Value ''
    $AltClearPWButton.Enabled = $False
    $AltClearPWButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
    $AltClearPWButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
    $AltSetPWButton.Enabled = $False
    $AltSetPWButton.Visible = $True
    $AltSetPWButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
    $AltSetPWButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
    $AltGetPWButton.Enabled = $False
    $AltGetPWButton.Visible = $False
    $OutText.AppendText("$(Get-Timestamp) - Your alternate password has been cleared.`r`n")
})

# Copy my password button logic
$GetPWButton.Add_Click({
    $PlainTxtPW = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($global:SecurePW))
    $PlainTxtPW | Set-Clipboard
    $OutText.AppendText("$(Get-Timestamp) - Your password has been copied to the clipboard.`r`n")
})

# Copy my alternate password button logic
$AltGetPWButton.Add_Click({
    $PlainTxtPW = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($global:AltSecurePW))
    $PlainTxtPW | Set-Clipboard
    $OutText.AppendText("$(Get-Timestamp) - Your alternate password has been copied to the clipboard.`r`n")
})

# Function for generating a random 16 character password
function New-Password {
    $Length = 16
    $Lowercase = "abcdefghijklmnopqrstuvwxyz"
    $Uppercase = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    $Numbers = "1234567890"
    $SpecialChars = "!@#$%^&*()-_=+"

    # Ensure at least one character from each category is included
    $GeneratedPassword = -join (
        ($Lowercase | Get-Random -Count 1),
        ($Uppercase | Get-Random -Count 1),
        ($Numbers | Get-Random -Count 1),
        ($SpecialChars | Get-Random -Count 1)
    )

    # Fill the rest of the password length with random characters from all categories
    $AllChars = $Lowercase + $Uppercase + $Numbers + $SpecialChars
    $GeneratedPassword += -join ($AllChars | Get-Random -Count ($Length - 4))

    # Shuffle the password to ensure the random distribution of characters
    $GeneratedPassword = $GeneratedPassword.ToCharArray() | Get-Random -Count $Length

    return -join $GeneratedPassword
}

# Generate password button logic
$GenPWButton.Add_Click({
    New-Password | Set-Clipboard
    $OutText.AppendText("$(Get-Timestamp) - A 16 character password has been generated and copied to the clipboard.`r`n")
})

<#
? **********************************************************************************************************************
? END OF PASSWORD MANAGER
? **********************************************************************************************************************
#>
<#
? **********************************************************************************************************************
? START OF CREATE HDTSTORAGE TABLE
? **********************************************************************************************************************
#>

# Button for launching the HDTStorage table creator
$LaunchHDTStorageButton = New-Object System.Windows.Forms.Button
$LaunchHDTStorageButton.Location = New-Object System.Drawing.Point(5, 400)
$LaunchHDTStorageButton.Width = 150
$LaunchHDTStorageButton.FlatStyle = "Popup"
$LaunchHDTStorageButton.Text = "Launch HDT Wizard"
$LaunchHDTStorageButton.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
$LaunchHDTStorageButton.Enabled = $true

# Label for launching the HDTStorage table creator
$LaunchHDTStorageLabel = New-Object System.Windows.Forms.Label
$LaunchHDTStorageLabel.Location = New-Object System.Drawing.Size(5, 375)
$LaunchHDTStorageLabel.Size = New-Object System.Drawing.Size(175, 20)
$LaunchHDTStorageLabel.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Bold)
$LaunchHDTStorageLabel.Text = 'Create HDTStorage Table'

# Event handler for the launch HDTStorage button
$LaunchHDTStorageButton.Add_Click({
    $OutText.AppendText("$(Get-Timestamp) - Launching HDTStorage Wizard...`r`n")
    if ((Get-WSManCredSSP).State -ne "Enabled") {
        Enable-WSManCredSSP -Role Client -DelegateComputer $script:WorkhorseServer -Force
    }
    # New popup window
    $script:HDTStoragePopup = New-Object System.Windows.Forms.Form
    $script:HDTStoragePopup.Text = "Create HDTStorage Table"
    $script:HDTStoragePopup.Size = New-Object System.Drawing.Size(300, 350)
    $script:HDTStoragePopup.ShowInTaskbar = $True
    $script:HDTStoragePopup.KeyPreview = $True
    $script:HDTStoragePopup.AutoSize = $True
    $script:HDTStoragePopup.FormBorderStyle = 'Fixed3D'
    $script:HDTStoragePopup.MaximizeBox = $False
    $script:HDTStoragePopup.MinimizeBox = $True
    $script:HDTStoragePopup.ControlBox = $True
    $script:HDTStoragePopup.Icon = $Icon
    $script:HDTStoragePopup.StartPosition = "CenterScreen"
    $global:IsHDTStoragePopupActive = $true

    # Button for selecting the file to upload
    $script:HDTStorageFileButton = New-Object System.Windows.Forms.Button
    $script:HDTStorageFileButton.Location = New-Object System.Drawing.Size(20, 30)
    $script:HDTStorageFileButton.Width = 75
    $script:HDTStorageFileButton.Text = "Browse"
    $script:HDTStorageFileButton.Enabled = $true

    # Label for the file location button
    $script:FileLocationLabel = New-Object System.Windows.Forms.Label
    $script:FileLocationLabel.Location = New-Object System.Drawing.Size(20, 10)
    $script:FileLocationLabel.Size = New-Object System.Drawing.Size(150, 20)
    $script:FileLocationLabel.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Bold)
    $script:FileLocationLabel.Text = "File Location"

    # Text box for entering the SQL instance
    $script:DBServerTextBox = New-Object System.Windows.Forms.TextBox
    $script:DBServerTextBox.Location = New-Object System.Drawing.Size(20, 80)
    $script:DBServerTextBox.Size = New-Object System.Drawing.Size(200, 20)
    $script:DBServerTextBox.Text = ''
    $script:DBServerTextBox.ShortcutsEnabled = $True

    # Label for the SQL instance text box
    $script:DBServerLabel = New-Object System.Windows.Forms.Label
    $script:DBServerLabel.Location = New-Object System.Drawing.Size(20, 60)
    $script:DBServerLabel.Size = New-Object System.Drawing.Size(150, 20)
    $script:DBServerLabel.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Bold)
    $script:DBServerLabel.Text = "SQL Instance"

    # Text box for entering the table name
    $script:TableNameTextBox = New-Object System.Windows.Forms.TextBox
    $script:TableNameTextBox.Location = New-Object System.Drawing.Size(20, 130)
    $script:TableNameTextBox.Size = New-Object System.Drawing.Size(200, 20)
    $script:TableNameTextBox.Text = ''
    $script:TableNameTextBox.ShortcutsEnabled = $True

    # Label for the table name text box
    $script:TableNameLabel = New-Object System.Windows.Forms.Label
    $script:TableNameLabel.Location = New-Object System.Drawing.Size(20, 110)
    $script:TableNameLabel.Size = New-Object System.Drawing.Size(150, 20)
    $script:TableNameLabel.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Bold)
    $script:TableNameLabel.Text = "Table Name"

    # Text box for entering secure password
    $script:SecurePasswordTextBox = New-Object System.Windows.Forms.TextBox
    $script:SecurePasswordTextBox.Location = New-Object System.Drawing.Size(20, 180)
    $script:SecurePasswordTextBox.Size = New-Object System.Drawing.Size(200, 20)
    $script:SecurePasswordTextBox.Text = ''
    $script:SecurePasswordTextBox.ShortcutsEnabled = $True
    $script:SecurePasswordTextBox.PasswordChar = '*'

    # Label for the secure password text box
    $script:SecurePasswordLabel = New-Object System.Windows.Forms.Label
    $script:SecurePasswordLabel.Location = New-Object System.Drawing.Size(20, 160)
    $script:SecurePasswordLabel.Size = New-Object System.Drawing.Size(150, 20)
    $script:SecurePasswordLabel.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Bold)
    $script:SecurePasswordLabel.Text = "Your Allied Password"

    # Checkbox to use the user's PW entered in Password Manager
    $script:HDTStoragePWCheckbox = New-Object System.Windows.Forms.CheckBox
    $script:HDTStoragePWCheckbox.Location = New-Object System.Drawing.Size(20, 210)
    $script:HDTStoragePWCheckbox.Size = New-Object System.Drawing.Size(200, 20)
    $script:HDTStoragePWCheckbox.Text = "Use Your Saved PW"
    $script:HDTStoragePWCheckbox.Enabled = -not [string]::IsNullOrEmpty($global:SecurePW)

    # Button for running the script
    $script:RunScriptButton = New-Object System.Windows.Forms.Button
    $script:RunScriptButton.Location = New-Object System.Drawing.Size(20, 250)
    $script:RunScriptButton.BackColor = $global:DisabledBackColor
    $script:RunScriptButton.ForeColor = $global:DisabledForeColor
    $script:RunScriptButton.FlatStyle = "Popup"
    $script:RunScriptButton.Width = 75
    $script:RunScriptButton.Text = "Create"
    $script:RunScriptButton.Enabled = $false

    # Check if DefaultUserTheme has a value or is null
    if ($null -ne $ConfigValues.DefaultUserTheme -and $ConfigValues.DefaultUserTheme -ne "") {
        # Get theme
        $SelectedTheme = $ColorTheme.PSObject.Properties | Where-Object { $_.Value.$($ConfigValues.DefaultUserTheme) } | Select-Object -ExpandProperty Name
        $themeColors = $ColorTheme.$SelectedTheme.$($ConfigValues.DefaultUserTheme)
        
        $script:HDTStoragePopup.BackColor = $themeColors.BackColor
        $script:HDTStoragePopup.ForeColor = $themeColors.ForeColor
        $script:HDTStorageFileButton.BackColor = $themeColors.ForeColor
        $script:HDTStorageFileButton.ForeColor = $themeColors.BackColor
        $script:FileLocationLabel.BackColor = $themeColors.BackColor
        $script:FileLocationLabel.ForeColor = $themeColors.ForeColor
        $script:DBServerLabel.BackColor = $themeColors.BackColor
        $script:DBServerLabel.ForeColor = $themeColors.ForeColor
        $script:TableNameLabel.BackColor = $themeColors.BackColor
        $script:TableNameLabel.ForeColor = $themeColors.ForeColor
        $script:SecurePasswordLabel.BackColor = $themeColors.BackColor
        $script:SecurePasswordLabel.ForeColor = $themeColors.ForeColor
    }

    # Call the function to set form icon
    Set-FormIcon

    # Add tooltips if the user has them enabled
    if ($ConfigValues.HoverToolTips -eq "Enabled" -or $null -eq $ConfigValues.HoverToolTips) {
        $ToolTip.SetToolTip($script:HDTStorageFileButton, "Click to select a file to upload")
        $ToolTip.SetToolTip($script:DBServerTextBox, "Enter the SQL instance name")
        $ToolTip.SetToolTip($script:TableNameTextBox, "Enter the table name, i.e. CSH12345_info")
        $ToolTip.SetToolTip($script:SecurePasswordTextBox, "Enter your Allied network password (this will not be saved)")
        $ToolTip.SetToolTip($script:HDTStoragePWCheckbox, "Check to use your saved password from Password Manager")
        $ToolTip.SetToolTip($script:RunScriptButton, "Click to create the HDTStorage table")
    }

    # Event handler for the file location button
    $script:HDTStorageFileButton.Add_Click({
        $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog
        $FileBrowser.InitialDirectory = "C:\"
        $FileBrowser.Filter = "Excel Files (*.xlsx)|*.xlsx"
        $FileBrowserResult = $FileBrowser.ShowDialog()
        if ($FileBrowserResult -eq 'OK') {
            $script:HDTStoragePopup.Tag = $FileBrowser.FileName
            Enable-RunScriptButton
        }
    })

    # Event handler for the SQL instance text box
    $script:DBServerTextBox.Add_TextChanged({
        Enable-RunScriptButton
    })

    # Event handler for the table name text box
    $script:TableNameTextBox.Add_TextChanged({
        Enable-RunScriptButton
    })

    # Event handler for the secure password text box
    $script:SecurePasswordTextBox.Add_TextChanged({
        Enable-RunScriptButton
    })

    # Event handler for the use saved password checkbox
    $script:HDTStoragePWCheckbox.Add_CheckedChanged({
        if ($script:HDTStoragePWCheckbox.Checked) {
            $script:SecurePasswordTextBox.Text = $global:SecurePW
        } else {
            $script:SecurePasswordTextBox.Text = ''
        }
    })

    # Event handler for the run script button
    $script:RunScriptButton.Add_Click({
        $DestinationPath = $ConfigValues.WorkhorseDirectoryPath
        $File = $script:HDTStoragePopup.Tag
        $OutPath = Join-Path -Path "\\$($ConfigValues.WorkhorseServer)" -ChildPath $DestinationPath  # Note: Using $ConfigValues.WorkhorseServer
        $DestinationFile = Join-Path -Path $OutPath -ChildPath (Get-Item $File).Name  # Combine destination directory and filename
        $LoginUsername = [System.Environment]::UserName
        $LoginPassword = ConvertTo-SecureString $script:SecurePasswordTextBox.Text -AsPlainText -Force
        $LoginCredentials = New-Object System.Management.Automation.PSCredential ($LoginUsername, $LoginPassword)
        Copy-Item -Path $File -Destination $DestinationFile -Force

        $Instance = $script:DBServerTextBox.Text
        $TableName = $script:TableNameTextBox.Text
        try {
            Invoke-Command -ComputerName $ConfigValues.WorkhorseServer -Authentication Credssp -Credential $LoginCredentials -ScriptBlock {
                param(
                    $File, 
                    $Instance, 
                    $TableName
                )
                $Database = "HDTStorage"
                foreach($sheet in Get-ExcelSheetInfo $File) {
                    $data = Import-Excel -Path $File -WorksheetName $sheet.name | ConvertTo-DbaDataTable
                    Write-DbaDataTable -SqlInstance $Instance -Database $Database -InputObject $data -AutoCreateTable -Table $TableName
                }
            } -ArgumentList $DestinationFile, $Instance, $TableName
            $OutText.AppendText("$(Get-Timestamp) - HDTStorage table $($TableName) created.`r`n")
        } catch {
            $OutText.AppendText("$(Get-Timestamp) - SQL Command failed to execute. Error: " + $_.Exception.Message + "`r`n")
        }        
    })


    # Click event for form close; sets global variable to false
    $script:HDTStoragePopup.Add_FormClosed({
        $global:IsHDTStoragePopupActive = $false
    })

    $script:HDTStoragePopup.Controls.Add($script:HDTStorageFileButton)
    $script:HDTStoragePopup.Controls.Add($script:DBServerTextBox)
    $script:HDTStoragePopup.Controls.Add($script:TableNameTextBox)
    $script:HDTStoragePopup.Controls.Add($script:FileLocationLabel)
    $script:HDTStoragePopup.Controls.Add($script:DBServerLabel)
    $script:HDTStoragePopup.Controls.Add($script:TableNameLabel)
    $script:HDTStoragePopup.Controls.Add($script:SecurePasswordTextBox)
    $script:HDTStoragePopup.Controls.Add($script:SecurePasswordLabel)
    $script:HDTStoragePopup.Controls.Add($script:HDTStoragePWCheckbox)
    $script:HDTStoragePopup.Controls.Add($script:RunScriptButton)
    $script:HDTStoragePopup.Show()
})

<#
? **********************************************************************************************************************
? END OF CREATE HDTSTORAGE TABLE
? **********************************************************************************************************************
#>
<#
? **********************************************************************************************************************
? START OF DOCUMENTATION CREATOR
? **********************************************************************************************************************
#>

# Text box for creating new documentation
$NewDocTextBox = New-Object System.Windows.Forms.TextBox
$NewDocTextBox.Location = New-Object System.Drawing.Size(265, 400)
$NewDocTextBox.Size = New-Object System.Drawing.Size(140, 30)
$NewDocTextBox.Text = ''
$NewDocTextBox.ShortcutsEnabled = $True

# Button for creating new documentation
$NewDocButton = New-Object System.Windows.Forms.Button
$NewDocButton.Location = New-Object System.Drawing.Point(425, 400)
$NewDocButton.Width = 100
$NewDocButton.BackColor = $global:DisabledBackColor
$NewDocButton.ForeColor = $global:DisabledForeColor
$NewDocButton.FlatStyle = "Popup"
$NewDocButton.Text = "Create"
$NewDocButton.Enabled = $false

# Label for new documentation text box
$NewDocLabel = New-Object System.Windows.Forms.Label
$NewDocLabel.Location = New-Object System.Drawing.Size(265, 375)
$NewDocLabel.Size = New-Object System.Drawing.Size(150, 20)
$NewDocLabel.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Bold)
$NewDocLabel.Text = 'Create Documentation'

# If Create Documentation text box is empty, disable the Create button
$NewDocTextBox.Add_TextChanged({
    if ($NewDocTextBox.Text.Length -eq 0) {
        $NewDocButton.Enabled = $False
        $NewDocButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
        $NewDocButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
    }
    else {
        $NewDocButton.Enabled = $True

		$backColor = Get-AppropriateColor -ColorType "BackColor"
        $foreColor = Get-AppropriateColor -ColorType "ForeColor"

		$NewDocButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
		$NewDocButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
        }
})

# New documentation function
Function New-DocTemplate {
    param (
        [string]$DocTopic
    )
# Variables
$TemplateFile = $ConfigValues.TemplateFile.Replace("{USERPROFILE}", $userProfilePath)
$FindText = "Documentation Template"
$DocTopic = $NewDocTextBox.Text
$NewFile = Join-Path "C:" "$DocTopic.docx"
$MatchCase = $false
$MatchWholeWorld = $true
$MatchWildcards = $false
$MatchSoundsLike = $false
$MatchAllWordForms = $false
$Forward = $false
$Wrap = 1
$Format = $false
$Replace = 2

# Creates a new instance of the Word application using the Component Object Model (COM)
$Word = New-Object -ComObject Word.Application

# Open the document
$Document = $Word.Documents.Open("$TemplateFile")

# Get the first section of the document
$section = $document.Sections.Item(1)

# Get the header of the first section
$header = $section.Headers.Item(1)

# Find and replace the date in the header
$searchText = "1/1/2023"
$NewDate = (Get-Date).ToString("M/d/yyyy")
$header.Range.Find.Execute($searchText, $false, $false, $false, $false, $false, $true, 1, $false, $NewDate, 2)

# Find and replace text using the variables above
$Document.Content.Find.Execute($FindText, $MatchCase, $MatchWholeWorld, $MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $Wrap, $Format, $DocTopic, $Replace)

try {

    # Save a new copy of the document
    $Document.SaveAs("$NewFile")
} 
catch {
    # Catch any exceptions that occur during file save
    $OutText.AppendText("$(Get-Timestamp) - An error occurred while trying to create the new Word document. This is likely due to the document title exceeding 255 characters.`r`n")
    $Document.Close([Microsoft.Office.Interop.Word.WdSaveOptions]::wdDoNotSaveChanges)
}

# Close the Word application
$Word.Quit()

# Release the COM object
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word)

# Set the variable to $null
$Word = $null

$DocumentationPath = $ConfigValues.DocumentationPath.Replace("{USERPROFILE}", $userProfilePath)
New-item -Path $DocumentationPath -Name "$DocTopic" -ItemType "directory"
Move-Item -Path $NewFile -Destination "$DocumentationPath\$DocTopic"
Invoke-Item -Path "$DocumentationPath\$DocTopic\$DocTopic.docx"
$OutText.AppendText("$(Get-Timestamp) - New document created at $DocumentationPath\$DocTopic\$DocTopic.docx`r`n")
}

# Function to test for invalid characters in the doc name
function Test-ValidFileName
{
    param([string]$DocTopic)

    $IndexOfInvalidChar = $DocTopic.IndexOfAny([System.IO.Path]::GetInvalidFileNameChars())

    # IndexOfAny() returns the value -1 to indicate no such character was found
    return $IndexOfInvalidChar -eq -1
}

# Function to check document name length
function Test-DocLength
{
    param([string]$DocTopic)

    $DocTopic = $NewDocTextBox.Text
    if ($DocTopic.length -gt 200) {
        $OutText.AppendText("$(Get-Timestamp) - Please enter a document name less than 200 characters`r`n")
        return
    }
    elseif (Test-ValidFileName $DocTopic) {
        $OutText.AppendText("$(Get-Timestamp) - Creating documentation for $DocTopic...`r`n")
        New-DocTemplate -DocTopic $DocTopic
    }
    else {
        $OutText.AppendText("$(Get-Timestamp) - Please enter a document name without any of the following characters: \ / : * ? < > |`r`n")
    }  
}

# New Document button press logic
$NewDocButton.Add_Click({
    Test-DocLength 
})

# New Document text box Enter key logic
$NewDocTextBox.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") {
        Test-DocLength        
    }
})

<#
? **********************************************************************************************************************
? END OF DOCUMENTATION CREATOR
? **********************************************************************************************************************
#>
<#
? **********************************************************************************************************************
? START OF TICKET MANAGER 
? **********************************************************************************************************************
#>

# Global variables for ticket manager
$TicketsPath = $ConfigValues.TicketsPath.Replace("{USERPROFILE}", $userProfilePath)

# Button for new ticket
$NewTicketButton = New-Object System.Windows.Forms.Button
$NewTicketButton.Location = New-Object System.Drawing.Point(150, 40)
$NewTicketButton.Width = 75
$NewTicketButton.BackColor = $global:DisabledBackColor
$NewTicketButton.ForeColor = $global:DisabledForeColor
$NewTicketButton.FlatStyle = "Popup"
$NewTicketButton.Text = "New Ticket"
$NewTicketButton.Enabled = $false

# Text box for new ticket
$NewTicketTextBox = New-Object System.Windows.Forms.TextBox
$NewTicketTextBox.Location = New-Object System.Drawing.Point(5, 40)
$NewTicketTextBox.Size = New-Object System.Drawing.Size(140, 30)
$NewTicketTextBox.Text = ''
$NewTicketTextBox.ShortcutsEnabled = $True

# Label for new ticket text box
$NewTicketLabel = New-Object System.Windows.Forms.Label
$NewTicketLabel.Location = New-Object System.Drawing.Size(5, 15)
$NewTicketLabel.Size = New-Object System.Drawing.Size(150, 20)
$NewTicketLabel.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Bold)
$NewTicketLabel.Text = 'Create a New Ticket'

# Button for renaming a ticket
$RenameTicketButton = New-Object System.Windows.Forms.Button
$RenameTicketButton.Location = New-Object System.Drawing.Point(455, 40)
$RenameTicketButton.Width = 75
$RenameTicketButton.BackColor = $global:DisabledBackColor
$RenameTicketButton.ForeColor = $global:DisabledForeColor
$RenameTicketButton.FlatStyle = "Popup"
$RenameTicketButton.Text = "Rename"
$RenameTicketButton.Enabled = $false

# Text box for renaming a ticket
$RenameTicketTextBox = New-Object System.Windows.Forms.TextBox
$RenameTicketTextBox.Location = New-Object System.Drawing.Point(310, 40)
$RenameTicketTextBox.Size = New-Object System.Drawing.Size(140, 30)
$RenameTicketTextBox.Text = ''
$RenameTicketTextBox.ShortcutsEnabled = $True

# Label applist combo box
$RenameTicketLabel = New-Object System.Windows.Forms.Label
$RenameTicketLabel.Location = New-Object System.Drawing.Size(310, 15)
$RenameTicketLabel.Size = New-Object System.Drawing.Size(150, 20)
$RenameTicketLabel.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Bold)
$RenameTicketLabel.Text = 'Rename a Ticket'

# Tab control for ticket manager
$TicketManagerTabControl = New-Object System.Windows.Forms.TabControl
$TicketManagerTabControl.Location = "5,95"
$TicketManagerTabControl.Size = "220,250"

# Tab for active tickets
$ActiveTicketsTab = New-Object System.Windows.Forms.TabPage
$ActiveTicketsTab.DataBindings.DefaultDataSourceUpdateMode = 0
$ActiveTicketsTab.UseVisualStyleBackColor = $True
$ActiveTicketsTab.Name = 'ActiveTickets'
$ActiveTicketsTab.Text = 'Active Tickets'
$ActiveTicketsTab.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)

# Tab for completed tickets
$CompletedTicketsTab = New-Object System.Windows.Forms.TabPage
$CompletedTicketsTab.DataBindings.DefaultDataSourceUpdateMode = 0
$CompletedTicketsTab.UseVisualStyleBackColor = $True
$CompletedTicketsTab.Name = 'CompletedTickets'
$CompletedTicketsTab.Text = 'Completed Tickets'
$CompletedTicketsTab.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)

# List box to show folder contents
$FolderContentsListBox = New-Object System.Windows.Forms.ListBox
$FolderContentsListBox.Location = New-Object System.Drawing.Point(310, 95)
$FolderContentsListBox.Size = New-Object System.Drawing.Size(220, 255)
$FolderContentsListBox.BackColor = $OutText.BackColor

# Complete ticket button
$CompleteTicketButton = New-Object System.Windows.Forms.Button
$CompleteTicketButton.Location = New-Object System.Drawing.Point(5, 360)
$CompleteTicketButton.Width = 75
$CompleteTicketButton.BackColor = $global:DisabledBackColor
$CompleteTicketButton.ForeColor = $global:DisabledForeColor
$CompleteTicketButton.FlatStyle = "Popup"
$CompleteTicketButton.Text = "Complete"
$CompleteTicketButton.Enabled = $false
$CompleteTicketButton.Visible = $true

# Reactivate ticket button
$ReactivateTicketButton = New-Object System.Windows.Forms.Button
$ReactivateTicketButton.Location = New-Object System.Drawing.Point(5, 360)
$ReactivateTicketButton.Width = 75
$ReactivateTicketButton.BackColor = $global:DisabledBackColor
$ReactivateTicketButton.ForeColor = $global:DisabledForeColor
$ReactivateTicketButton.FlatStyle = "Popup"
$ReactivateTicketButton.Text = "Reactivate"
$ReactivateTicketButton.Enabled = $false
$ReactivateTicketButton.Visible = $false

# Open folder button
$OpenFolderButton = New-Object System.Windows.Forms.Button
$OpenFolderButton.Location = New-Object System.Drawing.Point(150, 360)
$OpenFolderButton.Width = 75
$OpenFolderButton.BackColor = $global:DisabledBackColor
$OpenFolderButton.ForeColor = $global:DisabledForeColor
$OpenFolderButton.FlatStyle = "Popup"
$OpenFolderButton.Text = "Open"
$OpenFolderButton.Enabled = $false

# List box for active tickets
$ActiveTicketsListBox = New-Object System.Windows.Forms.ListBox
$ActiveTicketsListBox.Location = New-Object System.Drawing.Point(0,0)
$ActiveTicketsListBox.Size = New-Object System.Drawing.Size(215,240)
$ActiveTicketsListBox.SelectionMode = 'MultiExtended'
$ActiveTicketsListBox.BackColor = $OutText.BackColor

# List box for completed tickets
$CompletedTicketsListBox = New-Object System.Windows.Forms.ListBox
$CompletedTicketsListBox.Location = New-Object System.Drawing.Point(0,0)
$CompletedTicketsListBox.Size = New-Object System.Drawing.Size(215,240)
$CompletedTicketsListBox.SelectionMode = 'MultiExtended'
$CompletedTicketsListBox.BackColor = $OutText.BackColor

# Creates the active and completed tickets path variables on startup
# Calls function immediately after
function Start-Setup {
    $ActiveTicketsPath = "$TicketsPath\Active"
    $CompletedTicketsPath = "$TicketsPath\Completed"
    if (!(Test-Path -Path $ActiveTicketsPath)) {
        mkdir $ActiveTicketsPath | Out-Null
    }   
    if (!(Test-Path -Path $CompletedTicketsPath)) {
        mkdir $CompletedTicketsPath | Out-Null
    }
}
Start-Setup

# Function to populate active tickets list box
# Calls function immediately after
Function Get-ActiveListItems {
    param($listbox)
    if ($listbox.items.Count -gt 0){
        $listBox.Items.Clear()
    }
$tickets = @(Get-ChildItem "$TicketsPath\Active" | Select-Object -ExpandProperty Name)
foreach ($ticket in $tickets) {
    [void]$ActiveTicketsListBox.Items.Add($ticket)
}
}
Get-ActiveListItems($ActiveTicketsListBox)

# Function to populate completed tickets list box
# Calls function immediately after
Function Get-CompletedListItems {
    param($listbox)
    if ($listbox.items.Count -gt 0){
        $listBox.Items.Clear()
    }
$tickets = @(Get-ChildItem "$TicketsPath\Completed" | Select-Object -ExpandProperty Name)
foreach ($ticket in $tickets) {
    [void]$CompletedTicketsListBox.Items.Add($ticket)
}
}
Get-CompletedListItems($CompletedTicketsListBox)

# Logic for turning off the rename functionality
function Set-RenameOff {
    $RenameTicketButton.Enabled = $false
    $RenameTicketTextBox.Text = ''
    $RenameTicketTextBox.Enabled = $false
    $FolderContentsListBox.Items.Clear()
}

# Logic to enable new ticket button when text box is not empty
$NewTicketTextBox.Add_TextChanged({
    if ($NewTicketTextBox.Text.Length -eq 0) {
        $NewTicketButton.Enabled = $False
        $NewTicketButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
        $NewTicketButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
    }
    else {
        $NewTicketButton.Enabled = $True
		
		$backColor = Get-AppropriateColor -ColorType "BackColor"
        $foreColor = Get-AppropriateColor -ColorType "ForeColor"
		
        $NewTicketButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
        $NewTicketButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
    }
})

# Function for new ticket logic
function New-Ticket {
    $TicketNumber = $NewTicketTextBox.Text
    mkdir "$TicketsPath\Active\$TicketNumber"
    if ($TicketNumber -like "DS*" ) {
        $DSTicket = $TicketNumber.Substring(3, 5)
        $UrlPath = "http://tfs-sharepoint.alliedsolutions.net/Sites/Unitrac/Lists/Database%20Scripting/DispForm.aspx?ID=$DSTicket"
    }
    else {
        $UrlPath = "https://alliedsolutions.atlassian.net/browse/$TicketNumber"
    }
    $wshShell = New-Object -ComObject "WScript.Shell"
    $URLShortcut = $wshShell.CreateShortcut(
        "$TicketsPath\Active\$TicketNumber\$TicketNumber.url")
    $URLShortcut.TargetPath = $UrlPath
    $URLShortcut.Save()
    $NewTicketTextBox.Text = ''
    $OutText.AppendText("$(Get-Timestamp) - New ticket created: $TicketNumber`r`n")
    Invoke-Item "$TicketsPath\Active\$TicketNumber"
    Get-ActiveListItems($ActiveTicketsListBox)
    Get-CompletedListItems($CompletedTicketsListBox)
    Set-RenameOff
}

# Logic to enable rename ticket button when text box is not empty
$RenameTicketTextBox.Add_TextChanged({
    if ($RenameTicketTextBox.Text.Length -eq 0) {
        $RenameTicketButton.Enabled = $False
        $RenameTicketButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
        $RenameTicketButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
    }
    else {
        $RenameTicketButton.Enabled = $True
		
		$backColor = Get-AppropriateColor -ColorType "BackColor"
        $foreColor = Get-AppropriateColor -ColorType "ForeColor"
		
        $RenameTicketButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
        $RenameTicketButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
    }
})

# Logic for buttons when active tickets are selected
$ActiveTicketsListBox.Add_SelectedIndexChanged({
    if ($ActiveTicketsListBox.SelectedItems.Count -gt 0) {
        $CompleteTicketButton.Enabled = $true
		$OpenFolderButton.Enabled = $true
		$ReactivateTicketButton.Enabled = $false
        $RenameTicketButton.Enabled = $true
		$RenameTicketTextBox.Enabled = $true
		
		$backColor = Get-AppropriateColor -ColorType "BackColor"
        $foreColor = Get-AppropriateColor -ColorType "ForeColor"
		
        $CompleteTicketButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
        $CompleteTicketButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
        $OpenFolderButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
        $OpenFolderButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
        $RenameTicketButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
        $RenameTicketButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
        $RenameTicketTextBox.Text = $ActiveTicketsListBox.SelectedItem
    } else {
        $CompleteTicketButton.Enabled = $false
		$OpenFolderButton.Enabled = $false
		$ReactivateTicketButton.Enabled = $false
		$RenameTicketTextBox.Enabled = $false
        $RenameTicketButton.Enabled = $false
        $CompleteTicketButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
        $CompleteTicketButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
        $OpenFolderButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
        $OpenFolderButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
        $ReactivateTicketButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
        $ReactivateTicketButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
        $RenameTicketButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
        $RenameTicketButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
    }
})

# Logic for buttons when completed tickets are selected
$CompletedTicketsListBox.Add_SelectedIndexChanged({
    if ($CompletedTicketsListBox.SelectedItems.Count -gt 0) {
        $CompleteTicketButton.Enabled = $false
		$OpenFolderButton.Enabled = $true
		$ReactivateTicketButton.Enabled = $true
		$RenameTicketButton.Enabled = $true
		$RenameTicketTextBox.Enabled = $true
		
		$backColor = Get-AppropriateColor -ColorType "BackColor"
        $foreColor = Get-AppropriateColor -ColorType "ForeColor"
		
        $CompleteTicketButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
        $CompleteTicketButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
        $OpenFolderButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
        $OpenFolderButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
        $ReactivateTicketButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
        $ReactivateTicketButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
        $RenameTicketButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
        $RenameTicketButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
        $RenameTicketTextBox.Text = $CompletedTicketsListBox.SelectedItem
    } else {
        $CompleteTicketButton.Enabled = $false
		$OpenFolderButton.Enabled = $false
		$ReactivateTicketButton.Enabled = $false
		$RenameTicketTextBox.Enabled = $false
        $RenameTicketButton.Enabled = $false
        $CompleteTicketButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
        $CompleteTicketButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
        $OpenFolderButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
        $OpenFolderButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
        $ReactivateTicketButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
        $ReactivateTicketButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
        $RenameTicketButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
        $RenameTicketButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
    }
})

# Event handler for displaying folder contents of selected active ticket
$ActiveTicketsListBox_SelectedIndexChanged = {
    $selectedItem = $ActiveTicketsListBox.SelectedItem
    $FolderContentsListBox.Items.Clear()
    $ActivePath = Join-Path $TicketsPath "Active\$selectedItem"
    Get-ChildItem -Path $ActivePath | ForEach-Object {
        $FolderContentsListBox.Items.Add($_.Name)
    }
}

# Event handler for displaying folder contents of selected completed ticket
$CompletedTicketsListBox_SelectedIndexChanged = {
    $selectedItem = $CompletedTicketsListBox.SelectedItem
    $FolderContentsListBox.Items.Clear()
    $CompletedPath = Join-Path $TicketsPath "Completed\$selectedItem"
    Get-ChildItem -Path $CompletedPath | ForEach-Object {
        $FolderContentsListBox.Items.Add($_.Name)
    }
}

# Event handler for clearing everything when switching between tabs
$TicketManagerTabControl_SelectedIndexChanged = {
    $ActiveTicketsListBox.ClearSelected()
    $CompletedTicketsListBox.ClearSelected()
    $RenameTicketTextBox.Clear()
    $FolderContentsListBox.Items.Clear()
    $OpenFolderButton.Enabled = $false
    $OpenFolderButton.BackColor = $global:DisabledBackColor
    $OpenFolderButton.ForeColor = $global:DisabledForeColor
}

# Function to check if a duplicate ticket exists in the Active and Completed directories
function Find-DupeTickets {
    param (
        [string]$TicketNumber
    )
    $Directories = Get-ChildItem -Path $TicketsPath -Recurse -Directory
    $FolderExists = $false

    foreach ($Directory in $Directories) {
        if ($Directory.Name -eq $TicketNumber) {
            $FolderExists = $true
            if ($Directory.FullName -match '\\Active\\') {
                $OutText.AppendText("$(Get-Timestamp) - Ticket $TicketNumber already exists in the Active tickets folder`r`n")
            }
            elseif ($Directory.FullName -match '\\Completed\\') {
                $OutText.AppendText("$(Get-Timestamp) - Ticket $TicketNumber already exists in the Completed tickets folder`r`n")
            }
            else {
                $OutText.AppendText("$(Get-Timestamp) - Ticket $TicketNumber already exists`r`n")
            }
            break
        }
    }
    return $FolderExists
}

# Event handler for new ticket button
$NewTicketButton.Add_Click({
    $TicketNumber = $NewTicketTextBox.Text
    $FolderExists = Find-DupeTickets -TicketNumber $TicketNumber
    if (-not $FolderExists) {
        New-Ticket
    }
})

# Event handler for rename ticket button
$RenameTicketButton.Add_Click({
    $TicketNumber = $RenameTicketTextBox.Text
    $FolderExists = Find-DupeTickets -TicketNumber $TicketNumber
    if (-not $FolderExists) {
        $ticket = $ActiveTicketsListBox.SelectedItem
        $NewName  = $RenameTicketTextBox.Text
        if ($ticket -ne $NewName) {
            Rename-Item -Path "$TicketsPath\Active\$ticket" -NewName $NewName
            $OutText.AppendText("$(Get-Timestamp) - Ticket $ticket renamed to $NewName`r`n")
            Set-RenameOff
            Get-ActiveListItems($ActiveTicketsListBox)
            Get-CompletedListItems($CompletedTicketsListBox)
        }
        else {
            $OutText.AppendText("$(Get-Timestamp) - Please enter a new ticket name`r`n")
        }
    }
})

# Event handler for complete ticket button
$CompleteTicketButton.Add_Click({
    $tickets = $ActiveTicketsListBox.SelectedItems
    foreach ($ticket in $tickets){
        Move-Item -Path "$TicketsPath\Active\$ticket" -Destination "$TicketsPath\Completed\"
        $OutText.AppendText("$(Get-Timestamp) - Ticket $ticket completed`r`n")
    }
    Set-RenameOff
    Get-ActiveListItems($ActiveTicketsListBox)
    Get-CompletedListItems($CompletedTicketsListBox)
})

# Event handler for open folder button
$OpenFolderButton.Add_Click({
    if ($ActiveTicketsListBox.SelectedItems.Count -gt 0) {
        $tickets = $ActiveTicketsListBox.SelectedItems
        foreach ($ticket in $tickets){
            Invoke-Item "$TicketsPath\Active\$ticket"
            $OutText.Appendtext("$(Get-Timestamp) - Opened folder for ticket $ticket`r`n")
        }
    } else {
        $tickets = $CompletedTicketsListBox.SelectedItems
        foreach ($ticket in $tickets){
            Invoke-Item "$TicketsPath\Completed\$ticket"
            $OutText.Appendtext("$(Get-Timestamp) - Opened folder for ticket $ticket`r`n")
        }
    }
})

# Event handler for reactivate ticket button
$ReactivateTicketButton.Add_Click({
    $tickets = $CompletedTicketsListBox.SelectedItems
    foreach ($ticket in $tickets){
        Move-Item -Path "$TicketsPath\Completed\$ticket" -Destination "$TicketsPath\Active\"
        $OutText.AppendText("$(Get-Timestamp) - Ticket $ticket reactivated`r`n")
    }
    Set-RenameOff
    Get-ActiveListItems($ActiveTicketsListBox)
    Get-CompletedListItems($CompletedTicketsListBox)
})

# Event handler for Complete/Reactivate button visibility logic
$TicketManagerTabControl.Add_SelectedIndexChanged({
    if ($TicketManagerTabControl.SelectedTab -eq $ActiveTicketsTab) {
        $CompleteTicketButton.Visible = $true
        $ReactivateTicketButton.Visible = $false
    } else {
        $CompleteTicketButton.Visible = $false
        $ReactivateTicketButton.Visible = $true
    }
})

# Event handler for Enter key logic on New Ticket text box
$NewTicketTextBox.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") {
        $TicketNumber = $NewTicketTextBox.Text
        $FolderExists = Find-DupeTickets -TicketNumber $TicketNumber
        if (-not $FolderExists) {
            New-Ticket
        }
        $_.SuppressKeyPress = $true
    }
})

# Event handler for Enter key logic on Rename Ticket text box
$RenameTicketTextBox.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") {
        $TicketNumber = $RenameTicketTextBox.Text
        $FolderExists = Find-DupeTickets -TicketNumber $TicketNumber
        if (-not $FolderExists) {
            $ticket = $ActiveTicketsListBox.SelectedItem
            $NewName  = $RenameTicketTextBox.Text
            if ($ticket -ne $NewName) {
                Rename-Item -Path "$TicketsPath\Active\$ticket" -NewName $NewName
                Set-RenameOff
                Get-ActiveListItems($ActiveTicketsListBox)
                Get-CompletedListItems($CompletedTicketsListBox)
            }
            else {
                $OutText.AppendText("$(Get-Timestamp) - Please enter a new ticket name`r`n")
            }
        }
        $_.SuppressKeyPress = $true
    }
})

# Register the Ticket Manager event handlers
# Store the script block in a variable
$OnActiveTicketSelectedAction = $ActiveTicketsListBox_SelectedIndexChanged

# Add the event using the script block
$ActiveTicketsListBox.add_SelectedIndexChanged($OnActiveTicketSelectedAction)

$ActiveTicketsListBox.add_SelectedIndexChanged($ActiveTicketsListBox_SelectedIndexChanged)
$CompletedTicketsListBox.add_SelectedIndexChanged($CompletedTicketsListBox_SelectedIndexChanged)
$TicketManagerTabControl.add_SelectedIndexChanged($TicketManagerTabControl_SelectedIndexChanged)

<#
? **********************************************************************************************************************
? END OF TICKET MANAGER 
? **********************************************************************************************************************
#>
<#
? **********************************************************************************************************************
? START OF FORM BUILD
? **********************************************************************************************************************
#>

# Build form
$Form.Controls.Add($MainFormTabControl)
$MainFormTabControl.Controls.Add($SysAdminTab)
$MainFormTabControl.Controls.Add($SupportTab)
$MainFormTabControl.Controls.Add($TicketManagerTab)
$Form.Controls.Add($OutText)
$Form.Controls.Add($ClearOutTextButton)
$Form.Controls.Add($SaveOutTextButton)
$Form.Controls.Add($MenuStrip)
$MenuStrip.Items.Add($FileMenu) | Out-Null
$MenuStrip.Items.Add($OptionsMenu) | Out-Null
$MenuStrip.Items.Add($AboutMenu) | Out-Null
$FileMenu.DropDownItems.Add($SubmitFeedback) | Out-Null
$FileMenu.DropDownItems.Add($MenuQuit) | Out-Null
$OptionsMenu.DropDownItems.Add($MenuToolTips) | Out-Null
$MenuToolTips.DropDownItems.Add($ShowHelpBannerMenu) | Out-Null
$MenuToolTips.DropDownItems.Add($ShowToolTipsMenu) | Out-Null
$OptionsMenu.DropDownItems.Add($MenuColorTheme) | Out-Null
$AboutMenu.DropDownItems.Add($MenuGitHub) | Out-Null
$AboutMenu.DropDownItems.Add($MenuAboutItem) | Out-Null
$MenuColorTheme.DropDownItems.Add($MLBThemes) | Out-Null
$MenuColorTheme.DropDownItems.Add($NBAThemes) | Out-Null
$MenuColorTheme.DropDownItems.Add($NFLThemes) | Out-Null
$MenuColorTheme.DropDownItems.Add($PremiumThemes) | Out-Null
$MenuColorTheme.DropDownItems.Add($TraditionalThemes) | Out-Null
$SysAdminTab.Controls.Add($RestartsTabControl)
$SysAdminTab.Controls.Add($ServersListBox)
$SysAdminTab.Controls.Add($AppListCombo)
$SysAdminTab.Controls.Add($RestartButton)
$SysAdminTab.Controls.Add($StartButton)
$SysAdminTab.Controls.Add($StopButton)
$SysAdminTab.Controls.Add($OpenSiteButton)
$SysAdminTab.Controls.Add($RestartIISButton)
$SysAdminTab.Controls.Add($StartIISButton)
$SysAdminTab.Controls.Add($StopIISButton)
$SysAdminTab.Controls.Add($ServerPingButton)
$SysAdminTab.Controls.Add($ServerPingTextBox)
$SysAdminTab.Controls.Add($NSLookupButton)
$SysAdminTab.Controls.Add($script:NSLookupTextBox)
$SysAdminTab.Controls.Add($ServerPingLabel)
$SysAdminTab.Controls.Add($NSLookupLabel)
$SysAdminTab.Controls.Add($ReverseIPTextBox)
$SysAdminTab.Controls.Add($ReverseIPLabel)
$SysAdminTab.Controls.Add($ReverseIPButton)
$SysAdminTab.Controls.Add($AppListLabel)
$SysAdminTab.Controls.Add($RestartsSeparator)
$ServicesTab.Controls.Add($ServicesListBox)
$IISSitesTab.Controls.Add($IISSitesListBox)
$AppPoolsTab.Controls.Add($AppPoolsListBox)
$SupportTab.Controls.Add($PSTCombo)
$SupportTab.Controls.Add($PSTComboLabel)
$SupportTab.Controls.Add($SelectEnvButton)
$SupportTab.Controls.Add($ResetEnvButton)
$SupportTab.Controls.Add($RunPSTButton)
$SupportTab.Controls.Add($RefreshPSTButton)
$SupportTab.Controls.Add($LaunchLFPWizardButton)
$SupportTab.Controls.Add($LaunchLFPWizardLabel)
$SupportTab.Controls.Add($PSTSeparator)
$SupportTab.Controls.Add($LaunchHDTStorageButton)
$SupportTab.Controls.Add($PWTextBox)
$SupportTab.Controls.Add($PWLabel)
$SupportTab.Controls.Add($SetPWButton)
$SupportTab.Controls.Add($AltSetPWButton)
$SupportTab.Controls.Add($GetPWButton)
$SupportTab.Controls.Add($AltGetPWButton)
$SupportTab.Controls.Add($ClearPWButton)
$SupportTab.Controls.Add($AltClearPWButton)
$SupportTab.Controls.Add($GenPWButton)
$SupportTab.Controls.Add($PWTextBoxLabel)
$SupportTab.Controls.Add($GenPWLabel)
$SupportTab.Controls.Add($PWManagerSeparator)
$SupportTab.Controls.Add($LaunchHDTStorageLabel)
$SupportTab.Controls.Add($NewDocTextBox)
$SupportTab.Controls.Add($NewDocButton)
$SupportTab.Controls.Add($NewDocLabel)
$TicketManagerTab.Controls.Add($NewTicketButton)
$TicketManagerTab.Controls.Add($NewTicketTextBox)
$TicketManagerTab.Controls.Add($RenameTicketButton)
$TicketManagerTab.Controls.Add($RenameTicketTextBox)
$TicketManagerTab.Controls.Add($TicketManagerTabControl)
$TicketManagerTab.Controls.Add($FolderContentsListBox)
$TicketManagerTab.Controls.Add($CompleteTicketButton)
$TicketManagerTab.Controls.Add($ReactivateTicketButton)
$TicketManagerTab.Controls.Add($OpenFolderButton)
$TicketManagerTab.Controls.Add($NewTicketLabel)
$TicketManagerTab.Controls.Add($RenameTicketLabel)
$TicketManagerTabControl.Controls.Add($ActiveTicketsTab)
$TicketManagerTabControl.Controls.Add($CompletedTicketsTab)
$ActiveTicketsTab.Controls.Add($ActiveTicketsListBox)
$CompletedTicketsTab.Controls.Add($CompletedTicketsListBox)

# Check if DefaultUserTheme has a value or is null
if ($null -ne $ConfigValues.DefaultUserTheme -and $ConfigValues.DefaultUserTheme -ne "") {
    $SelectedTheme = $ColorTheme.PSObject.Properties | Where-Object { $_.Value.$($ConfigValues.DefaultUserTheme) } | Select-Object -ExpandProperty Name
    $themeColors = $ColorTheme.$SelectedTheme.$($ConfigValues.DefaultUserTheme)

    # Check if the theme falls under Premium and set the AccentColor accordingly
    if ($SelectedTheme -eq 'Premium') {
        $OutText.BackColor = $themeColors.AccentColor
        $ServersListBox.Backcolor = $themeColors.AccentColor
        $ServicesListBox.Backcolor = $themeColors.AccentColor
        $IISSitesListBox.Backcolor = $themeColors.AccentColor
        $AppPoolsListBox.Backcolor = $themeColors.AccentColor
        $FolderContentsListBox.BackColor = $themeColors.AccentColor
        $ActiveTicketsListBox.BackColor = $themeColors.AccentColor
        $CompletedTicketsListBox.BackColor = $themeColors.AccentColor
        $ServerPingTextBox.BackColor = $themeColors.AccentColor
        $script:NSLookupTextBox.BackColor = $themeColors.AccentColor
        $ReverseIPTextBox.BackColor = $themeColors.AccentColor
        $PWTextBox.BackColor = $themeColors.AccentColor
        $NewDocTextBox.BackColor = $themeColors.AccentColor
        $NewTicketTextBox.BackColor = $themeColors.AccentColor
        $RenameTicketTextBox.BackColor = $themeColors.AccentColor
        $global:DisabledBackColor = $themeColors.DisabledColor        
    }
    else {
        $OutText.BackColor = [System.Drawing.SystemColors]::Control
        $ServersListBox.Backcolor = [System.Drawing.SystemColors]::Control
        $ServicesListBox.Backcolor = [System.Drawing.SystemColors]::Control
        $IISSitesListBox.Backcolor = [System.Drawing.SystemColors]::Control
        $AppPoolsListBox.Backcolor = [System.Drawing.SystemColors]::Control
        $FolderContentsListBox.BackColor = [System.Drawing.SystemColors]::Control
        $ActiveTicketsListBox.BackColor = [System.Drawing.SystemColors]::Control
        $CompletedTicketsListBox.BackColor = [System.Drawing.SystemColors]::Control
        $ServerPingTextBox.BackColor = [System.Drawing.SystemColors]::Control
        $script:NSLookupTextBox.BackColor = [System.Drawing.SystemColors]::Control
        $ReverseIPTextBox.BackColor = [System.Drawing.SystemColors]::Control
        $PWTextBox.BackColor = [System.Drawing.SystemColors]::Control
        $NewDocTextBox.BackColor = [System.Drawing.SystemColors]::Control
        $NewTicketTextBox.BackColor = [System.Drawing.SystemColors]::Control
        $RenameTicketTextBox.BackColor = [System.Drawing.SystemColors]::Control
        $global:DisabledBackColor = '#A9A9A9'
    }

    $global:DisabledForeColor = '#FFFFFF'

    # Set the colors for the form
    $Form.BackColor = $themeColors.BackColor
    $MenuStrip.BackColor = $themeColors.BackColor
    $MenuStrip.ForeColor = $themeColors.ForeColor
    $MainFormTabControl.BackColor = $themeColors.BackColor
    $SysAdminTab.BackColor = $themeColors.BackColor
    $SupportTab.BackColor = $themeColors.BackColor
    $TicketManagerTab.BackColor = $themeColors.BackColor
    $ServersListBox.Backcolor = $OutText.BackColor
    $ServicesListBox.Backcolor = $OutText.BackColor
    $IISSitesListBox.Backcolor = $OutText.BackColor
    $AppPoolsListBox.Backcolor = $OutText.BackColor
    $AppListCombo.BackColor = $OutText.BackColor
    $AppListLabel.ForeColor = $themeColors.ForeColor
    $RestartsSeparator.BackColor = $themeColors.ForeColor
    $NSLookupLabel.ForeColor = $themeColors.ForeColor
    $ServerPingLabel.ForeColor = $themeColors.ForeColor
    $ReverseIPLabel.ForeColor = $themeColors.ForeColor
    $PSTComboLabel.ForeColor = $themeColors.ForeColor
    $PSTCombo.BackColor = $OutText.BackColor
    $RunPSTButton.BackColor = $themeColors.ForeColor
    $RunPSTButton.ForeColor = $themeColors.BackColor
    $RefreshPSTButton.BackColor = $themeColors.ForeColor
    $RefreshPSTButton.ForeColor = $themeColors.BackColor
    $LaunchLFPWizardButton.BackColor = $themeColors.ForeColor
    $LaunchLFPWizardButton.ForeColor = $themeColors.BackColor
    $LaunchLFPWizardLabel.ForeColor = $themeColors.ForeColor
    $PSTSeparator.BackColor = $themeColors.ForeColor
    $PWTextBoxLabel.ForeColor = $themeColors.ForeColor
    $GenPWButton.BackColor = $themeColors.ForeColor
    $GenPWButton.ForeColor = $themeColors.BackColor
    $GenPWLabel.ForeColor = $themeColors.ForeColor
    $PWManagerSeparator.BackColor = $themeColors.ForeColor
    $LaunchHDTStorageButton.BackColor = $themeColors.ForeColor
    $LaunchHDTStorageButton.ForeColor = $themeColors.BackColor
    $LaunchHDTStorageLabel.ForeColor = $themeColors.ForeColor
    $NewDocLabel.ForeColor = $themeColors.ForeColor
    $NewTicketLabel.ForeColor = $themeColors.ForeColor
    $RenameTicketLabel.ForeColor = $themeColors.ForeColor
    $FolderContentsListBox.BackColor = $OutText.BackColor
    $ActiveTicketsListBox.BackColor = $OutText.BackColor
    $CompletedTicketsListBox.BackColor = $OutText.BackColor

    # Check if the theme falls under NBA, NFL, or MLB and set the icon accordingly
    if ($SelectedTheme -eq 'NBA' -or $SelectedTheme -eq 'NFL' -or $SelectedTheme -eq 'MLB') {
        $IconPath = "$scriptPath\TeamIcons\$SelectedTheme\$($ConfigValues.DefaultUserTheme).ico"

        if (Test-Path $IconPath) {
            $Icon = New-Object System.Drawing.Icon($IconPath)
        }
    }
    else {
        $IconPath = $ConfigValues.AlliedLogo.Replace("{USERPROFILE}", $UserProfilePath)

        if (Test-Path $IconPath) {
            $Icon = New-Object System.Drawing.Icon($IconPath)
        }
    }
}

if ($ConfigValues.HoverToolTips -eq "Enabled" -or $null -eq $ConfigValues.HoverToolTips) {
	Enable-ToolTips
    $ShowToolTipsMenu.Text = "Hide Tool Tips"
}
else {
    $ShowToolTipsMenu.Text = "Show Tool Tips"
}

$OutText.AppendText("$(Get-Timestamp) - Welcome to the ETG Desktop Assistant!`r`n")

Get-ThemeQuote

<# $script:Quote = $ColorTheme.$SelectedTheme.$($ConfigValues.DefaultUserTheme).Quote
$OutText.AppendText("$(Get-Timestamp) - $($ConfigValues.DefaultUserTheme) theme is now active.`r`n")
if ($script:Quote) {
    $OutText.AppendText("$(Get-Timestamp) - $script:Quote`r`n")
} #>

# When the form loads, set the icon if $icon is not $null
$Form.add_Load({
    if ($Icon) {
        $This.Icon = $Icon
    }
})

# Show Form
$Form.ShowDialog() | Out-Null

<#
? **********************************************************************************************************************
? END OF FORM BUILD
? **********************************************************************************************************************
#>

# Clean up sync hash when form is closed
# Reset the PST config file
$Form.Add_FormClosing({
    $synchash.Closed = $True  
    Reset-Environment
})