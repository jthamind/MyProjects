Add-Type -AssemblyName System.Data
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationCore
[System.Windows.Forms.Application]::EnableVisualStyles()

#* Version History
#* 0.0.0 - 06/26/2023 - Initial Concept
#* 1.0.0 - 09/26/2023 - Initial Release
#* 1.1.0 - 10/02/2023 - Lender LFP module beta for QA
#* 1.2.0 - 10/06/2023 - Premium themes
#* 1.3.0 - 10/11/2023 - Reverse IP Lookup
#* 1.4.0 - 10/12/2023 - Theme-specific quotes
#* 1.5.0 - 10/18/2023 - Full Lender LFP module for QA, Staging, and Prod
#* 2.0.0 - 11/02/2023 - Theme Builder
#* 2.0.1 - 11/08/2023 - Minor rework of Prod Support Tool logic to check/refresh files

#* Table of contents for this script
#* Use Control + F to search for the section you want to jump to

#*  1. Global Variables and Functions
#*  2. Main GUI
#*  3. Menu Strip
#*  4. Restarts GUI
#*  5. NSLookup
#*  6. Server Ping
#*  7. Reverse IP Lookup
#*  8. Prod Support Tool
#*  9. Add Lender to LFP Services
#* 10. Password Manager
#* 11. Create HDTStorage Table
#* 12. Documentation Creator
#* 13. Ticket Manager
#* 14. Form Build

# ================================================================================================

# todo - List of to-do items

# todo - CURRENT ISSUES
# todo - None!

# todo - TECHNICAL DEBT
# todo - Move documentation, tickets, and Prod Support Tool inside the script directory

# todo - FUTURE RELEASE FEATURES
# todo - Allies theme: text colors switch between the pride flag colors every letter. Light and Dark versions. Pride flag made out of label lines (ETA mid November 2023)
# todo - Full documentation for this app (ETA Q1 2024)
# todo - Module for remote PowerShell sessions on servers (ETA late November 2023)
# todo - Module for restarting DirectoryWatcher services (ETA mid November 2023)
# todo - Guided walkthrough for first time users (ETA late November 2023)
# todo - EC2 monitoring (need to see if access keys are possible, ETA unknown)

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

# synchash variables
$synchash.$OutText = $OutText

# Set timestamp function used throughout the script
function Get-Timestamp {
    return Get-Date -Format "yyyy/MM/dd hh:mm:ss"
}

$Icon = $null
$global:SecurePW = $null
$global:AltSecurePW = $null
$global:IsThemeBuilderPopupActive = $false
$global:IsThemeApplied = $false
$global:IsAboutPopupActive = $false
$global:IsLenderLFPPopupActive = $false
$global:IsHDTStoragePopupActive = $false
$global:IsFeedbackPopupActive = $false
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition
$script:PSTEnvironment = $null # Initializing variable for holding the selected PST Environment so it can be used across multiple event handlers
$global:WasMainThemeActivated = $false # Initializing variable for tracking whether the main theme has been activated yet
$global:WasThemePremium = $false # Initializing variable for tracking whether the theme is premium or not

# Initialize the tooltip
$ToolTip = New-Object System.Windows.Forms.ToolTip
$ToolTip.InitialDelay = 100

# Get color themes from json file
$ColorTheme = Get-Content -Path .\ColorThemes.json | ConvertFrom-Json

# Path to the config file
$configFilePath = '.\Config.json'

# Check if the config file exists
if (-not (Test-Path -Path $configFilePath)) {
    $OutText.AppendText("$((Get-Timestamp)) - The configuration file '$configFilePath' does not exist.`r`n")
    return
}
else {
    # Get config values from json file
    $ConfigValues = Get-Content -Path $configFilePath | ConvertFrom-Json
    $userProfilePath = [Environment]::GetEnvironmentVariable("USERPROFILE")
}

# Construct the paths for icon and logo
$AlliedIcon = Join-Path -Path $PSScriptRoot -ChildPath $ConfigValues.AlliedIcon
$AlliedLogo = Join-Path -Path $PSScriptRoot -ChildPath $ConfigValues.AlliedLogo

$TestingWebhookURL = $ConfigValues.TestingWebhook
$script:WorkhorseServer = $ConfigValues.WorkhorseServer

# Check if DefaultUserTheme has a value or is null
if ($null -ne $ConfigValues.DefaultUserTheme -and $ConfigValues.DefaultUserTheme -ne "") {
    $SelectedTheme = $ColorTheme.PSObject.Properties | Where-Object { $_.Value.$($ConfigValues.DefaultUserTheme) } | Select-Object -ExpandProperty Name
    $themeColors = $ColorTheme.$SelectedTheme.$($ConfigValues.DefaultUserTheme)
    if ($SelectedTheme -eq 'Premium' -or $SelectedTheme -eq 'Custom') {
        $global:DisabledBackColor = $themeColors.DisabledColor
    }
    else {
        $global:DisabledBackColor = '#A9A9A9'
    }
}
$global:DisabledForeColor = '#FFFFFF'

# Function to get the quote for the selected theme
function Get-ThemeQuote {
    $script:Quote = $script:ColorTheme.$script:SelectedTheme.$($ConfigValues.DefaultUserTheme).Quote
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
            if ($global:IsAboutPopupActive) {
                $script:AboutForm.Icon = $Icon
            }
            if ($global:IsLenderLFPPopupActive) {
                $script:LenderLFPPopup.Icon = $Icon
            }
            if ($global:IsHDTStoragePopupActive) {
                $script:HDTStoragePopup.Icon = $Icon
            }
            if ($global:isfeedbackpopupactive) {
                $script:FeedbackForm.Icon = $Icon
            }
            if ($global:IsThemeBuilderPopupActive) {
                $script:ThemeBuilderForm.Icon = $Icon
            }
        }
        else {
            $OutText.AppendText("$((Get-Timestamp)) - The icon file '$IconPath' does not exist.`r`n")
        }
    }
    else {
        if (Test-Path $AlliedIcon) {
            $Form.Icon = $AlliedIcon
            if ($global:IsAboutPopupActive) {
                $script:AboutForm.Icon = $AlliedIcon
            }
            if ($global:IsLenderLFPPopupActive) {
                $script:LenderLFPPopup.Icon = $AlliedIcon
            }
            if ($global:IsHDTStoragePopupActive) {
                $script:HDTStoragePopup.Icon = $AlliedIcon
            }
            if ($global:isfeedbackpopupactive) {
                $script:FeedbackForm.Icon = $AlliedIcon
            }
            if ($global:IsThemeBuilderPopupActive) {
                $script:ThemeBuilderForm.Icon = $AlliedIcon
            }
        }
        else {
            $OutText.AppendText("$((Get-Timestamp)) - The Allied icon file '$AlliedIcon' does not exist.`r`n")
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
	$ToolTip.SetToolTip($script:RestartButton, "Restart selected item(s)")
	$ToolTip.SetToolTip($script:StartButton, "Start selected item(s)")
	$ToolTip.SetToolTip($script:StopButton, "Stop selected item(s)")
	$ToolTip.SetToolTip($script:OpenSiteButton, "Open selected site in Windows Explorer")
	$ToolTip.SetToolTip($script:RestartIISButton, "Restart IIS on selected server")
	$ToolTip.SetToolTip($script:StartIISButton, "Start IIS on selected server")
	$ToolTip.SetToolTip($script:StopIISButton, "Stop IIS on selected server")
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
    foreach ($menu in @($script:CustomThemes, $script:MLBThemes, $script:NBAThemes, $NFLThemes, $PremiumThemes)) {
        foreach ($menuItem in $menu.DropDownItems) {
            $menuItem.Text = $menuItem.Text -replace '^\•\s', ''
        }
    }
}

# Function to update the menu items across all categories
function Update-ThemeMenuItems {
    param (
        [Parameter(Mandatory=$true)]
        [System.Windows.Forms.ToolStripMenuItem[]]$MenuCategories
    )

    foreach ($menuCategory in $MenuCategories) {
        foreach ($item in $menuCategory.DropDownItems) {
            $cleanText = $item.Text -replace "• ", ""  # Remove bullet point if present
            $item.Text = $cleanText  # Update the text to the clean version without the bullet

            if ($cleanText -eq $ConfigValues.DefaultUserTheme) {
                $item.Text = "• " + $cleanText  # Add bullet point to the current theme
                $item.Enabled = $false  # Disable the current theme menu item
            } else {
                $item.Enabled = $true  # Enable all other menu items
            }
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

# Function to enable white text for USA theme
function Enable-USAThemeTextColor {
    
    $OutText.ForeColor = 'White'
    $script:NSLookupTextBox.ForeColor = 'White'
    $ReverseIPTextBox.ForeColor = 'White'
    $ServerPingTextBox.ForeColor = 'White'
    $PWTextBox.ForeColor = 'White'
    $NewDocTextBox.ForeColor = 'White'
    $NewTicketTextBox.ForeColor = 'White'
    $RenameTicketTextBox.ForeColor = 'White'

    # Define common properties for all controls
    $script:whiteBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::White)
    $script:fontArialRegular9 = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
    $script:itemHeight = 15

    # An array of all controls that need to be owner-drawn
    $allControls = @($PSTCombo, $ActiveTicketsListBox, $CompletedTicketsListBox, $FolderContentsListBox, 
                     $ServersListBox, $ServicesListBox, $IISSitesListBox, $AppPoolsListBox, $AppListCombo)

    # Apply changes to all controls
	foreach ($control in $allControls) {
		$control.SuspendLayout() # Suspend layout logic

		$control.DrawMode = [System.Windows.Forms.DrawMode]::OwnerDrawFixed
		$control.Font = $script:fontArialRegular9
		$control.ItemHeight = $script:itemHeight

		# Remove existing DrawItem event handlers
		$control.remove_DrawItem($control.DrawItem)

		# Add the DrawItem event handler
		$control.add_DrawItem({
			param($s, $e)

			# Draw the background and focus rectangle
			$e.DrawBackground()
			$e.DrawFocusRectangle()

			# Draw the text for the item
			$point = New-Object System.Drawing.PointF($e.Bounds.X, $e.Bounds.Y)
			$e.Graphics.DrawString($s.Items[$e.Index].ToString(), $e.Font, $script:whiteBrush, $point)

			# Draw the focus rectangle if the list box has focus
			$e.DrawFocusRectangle()
		})

		$control.ResumeLayout($false) # Resume layout logic
		$control.Refresh() # Refresh the control to apply changes
		$control.Invalidate() # Force a complete redraw of the control
	}
    # Explicitly set the size of the ListBox after the theme change
    $ServicesListBox.Size = New-Object System.Drawing.Size(245,240)
    $IISSitesListBox.Size = New-Object System.Drawing.Size(245,240)
    $AppPoolsListBox.Size = New-Object System.Drawing.Size(245,240)
    $ServersListBox.Size = New-Object System.Drawing.Size(200,240)
    $ActiveTicketsListBox.Size = New-Object System.Drawing.Size(215,240)
    $CompletedTicketsListBox.Size = New-Object System.Drawing.Size(215,240)
    $FolderContentsListBox.Size = New-Object System.Drawing.Size(220,255)
}

# Function to disable white text for USA theme
function Disable-USAThemeTextColor {
    # Reset text boxes' ForeColor to default (typically black or control text color)
    $OutText.ForeColor = 'Black'
    $script:NSLookupTextBox.ForeColor = 'Black'
    $ReverseIPTextBox.ForeColor = 'Black'
    $ServerPingTextBox.ForeColor = 'Black'
    $PWTextBox.ForeColor = 'Black'
    $NewDocTextBox.ForeColor = 'Black'
    $NewTicketTextBox.ForeColor = 'Black'
    $RenameTicketTextBox.ForeColor = 'Black'

    # An array of all combo and list boxes to revert to normal drawing mode
    $allControls = @($PSTCombo, $ActiveTicketsListBox, $CompletedTicketsListBox, $FolderContentsListBox, 
                    $ServersListBox, $ServicesListBox, $IISSitesListBox, $AppPoolsListBox, $AppListCombo)

    # Loop through each control and reset its properties
    foreach ($control in $allControls) {
        $control.DrawMode = [System.Windows.Forms.DrawMode]::Normal
        $control.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
        $control.ItemHeight = 13

        # Remove any DrawItem event handlers associated with the control
        $control.remove_DrawItem($control.DrawItem)

        # Refresh the control to re-draw with updated settings
        $control.Refresh()
    }
    # Explicitly set the size of the ListBox after the theme change
    $ServicesListBox.Size = New-Object System.Drawing.Size(245,240)
    $IISSitesListBox.Size = New-Object System.Drawing.Size(245,240)
    $AppPoolsListBox.Size = New-Object System.Drawing.Size(245,240)
    $ServersListBox.Size = New-Object System.Drawing.Size(200,240)
    $ActiveTicketsListBox.Size = New-Object System.Drawing.Size(215,240)
    $CompletedTicketsListBox.Size = New-Object System.Drawing.Size(215,240)
    $FolderContentsListBox.Size = New-Object System.Drawing.Size(220,255)
}

# Function to update the main theme
function Update-MainTheme {
    param (
        [string]$Team,
        [string]$Category,
        [object]$ColorData
    )

    $script:ColorTheme = $ColorData
    $script:SelectedTheme = $Category
    $ConfigValues.DefaultUserTheme = $Team

    $global:CurrentTeamBackColor = $ColorData.$Category.$Team.BackColor
    $global:CurrentTeamForeColor = $ColorData.$Category.$Team.ForeColor

    # Update the synchronized hashtable with the new values
    $synchash['CurrentTeamBackColor'] = $global:CurrentTeamBackColor
    $synchash['CurrentTeamForeColor'] = $global:CurrentTeamForeColor

    # Call the function to enable white text for USA theme or disable it for all other themes
    if ($Team -eq 'USA') {
        Enable-USAThemeTextColor
    }
    else {
        Disable-USAThemeTextColor
    }

    if ($Team -eq 'Test Allies') {
        Enable-PrideTheme
    }
    
    if ($Category -eq 'Premium' -or $Category -eq 'Custom') {
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
    
    # Update the synchronized hashtable with the new values
    $synchash['CurrentTeamAccentColor'] = $global:CurrentTeamAccentColor


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
    $script:OpenSiteButton.BackColor = $ColorData.$Category.$Team.ForeColor
    $script:OpenSiteButton.ForeColor = $ColorData.$Category.$Team.BackColor
    $script:RestartIISButton.BackColor = $ColorData.$Category.$Team.ForeColor
    $script:RestartIISButton.ForeColor = $ColorData.$Category.$Team.BackColor
    $script:StartIISButton.BackColor = $ColorData.$Category.$Team.ForeColor
    $script:StartIISButton.ForeColor = $ColorData.$Category.$Team.BackColor
    $script:StopIISButton.BackColor = $ColorData.$Category.$Team.ForeColor
    $script:StopIISButton.ForeColor = $ColorData.$Category.$Team.BackColor
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

    if ($script:RestartButton.Enabled) {
        $script:RestartButton.BackColor = $ColorData.$Category.$Team.ForeColor
        $script:RestartButton.ForeColor = $ColorData.$Category.$Team.BackColor
    } else {
        $script:RestartButton.BackColor = $global:DisabledBackColor
        $script:RestartButton.ForeColor = $global:DisabledForeColor
    }

    if ($script:StartButton.Enabled) {
        $script:StartButton.BackColor = $ColorData.$Category.$Team.ForeColor
        $script:StartButton.ForeColor = $ColorData.$Category.$Team.BackColor
    } else {
        $script:StartButton.BackColor = $global:DisabledBackColor
        $script:StartButton.ForeColor = $global:DisabledForeColor
    }

    if ($script:StopButton.Enabled) {
        $script:StopButton.BackColor = $ColorData.$Category.$Team.ForeColor
        $script:StopButton.ForeColor = $ColorData.$Category.$Team.BackColor
    } else {
        $script:StopButton.BackColor = $global:DisabledBackColor
        $script:StopButton.ForeColor = $global:DisabledForeColor
    }

    if ($script:OpenSiteButton.Enabled) {
        $script:OpenSiteButton.BackColor = $ColorData.$Category.$Team.ForeColor
        $script:OpenSiteButton.ForeColor = $ColorData.$Category.$Team.BackColor
    } else {
        $script:OpenSiteButton.BackColor = $global:DisabledBackColor
        $script:OpenSiteButton.ForeColor = $global:DisabledForeColor
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

    # Updates the About menu popup controls if the popup is active
    # This is to prevent an error where the system tries to update the controls before they are created
    if ($global:IsAboutPopupActive) {
        $script:AboutForm.BackColor = $ColorData.$Category.$Team.BackColor
        $script:AboutForm.ForeColor = $ColorData.$Category.$Team.ForeColor
        $script:AboutForm.Refresh()
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
    if ($global:IsHDTStoragePopupActive) {
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
        if ($script:CreateHDTStorageButton.Enabled) {
            $script:CreateHDTStorageButton.BackColor = $ColorData.$Category.$Team.ForeColor
            $script:CreateHDTStorageButton.ForeColor = $ColorData.$Category.$Team.BackColor
        } else {
            $script:CreateHDTStorageButton.BackColor = $global:DisabledBackColor
            $script:CreateHDTStorageButton.ForeColor = $global:DisabledForeColor
        }
        $script:HDTStoragePopup.Refresh()
    }

    # Updates the feedback popup controls if the popup is active
    # This is to prevent an error where the system tries to update the controls before they are created
    if ($global:isfeedbackpopupactive) {
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
        'Custom' { $activeTheme = $script:CustomThemes.DropDownItems | Where-Object { $_.Text -eq $Team } }
        'MLB' { $activeTheme = $script:MLBThemes.DropDownItems | Where-Object { $_.Text -eq $Team } }
        'NBA' { $activeTheme = $script:NBAThemes.DropDownItems | Where-Object { $_.Text -eq $Team } }
        'NFL' { $activeTheme = $NFLThemes.DropDownItems | Where-Object { $_.Text -eq $Team } }
        'Premium' { $activeTheme = $PremiumThemes.DropDownItems | Where-Object { $_.Text -eq $Team } }
    }

    if ($activeTheme) {
        $activeTheme.Text = "• " + $activeTheme.Text
    }
    # Update the synchronized hashtable
    $Global:synchash['WasMainThemeActivated'] = $true
    $synchash['SelectedTheme'] = $Category
    # Return the current theme colors
    return $global:CurrentTeamBackColor, $global:CurrentTeamForeColor, $global:CurrentTeamAccentColor
}

# Function to evaluate if the submit feedback button should be enabled
function Enable-SubmitFeedbackButton {
    if ($script:UserNameRadioButton.Checked -eq $true) {
        if ($null -ne $FeedbackTextBox.Text -and $FeedbackTextBox.Text -ne '' -and
            $null -ne $script:UserNameTextBox.Text -and $script:UserNameTextBox.Text -ne '') {
            $script:SubmitFeedbackButton.Enabled = $true
        } else {
            $script:SubmitFeedbackButton.Enabled = $false
        }
    } else {
        if ($null -ne $FeedbackTextBox.Text -and $FeedbackTextBox.Text -ne '') {
            $script:SubmitFeedbackButton.Enabled = $true
        } else {
            $script:SubmitFeedbackButton.Enabled = $false
        }
    }

    if ($script:SubmitFeedbackButton.Enabled) {
        $backColor = Get-AppropriateColor -ColorType "BackColor"
        $foreColor = Get-AppropriateColor -ColorType "ForeColor"

        $script:SubmitFeedbackButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
        $script:SubmitFeedbackButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
    } else {
        $script:SubmitFeedbackButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
        $script:SubmitFeedbackButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
    }
}

# Function to open an asynchronous runspace and submit feedback through a Teams message
function Open-SubmitFeedbackRunspace {
    param (
        [string]$ConfigValuesSubmitFeedbackScript,
        [string]$UserProfilePath,
        [string]$FeedbackText,
        [bool]$IsUserNameChecked,
        [string]$UserNameText,
        [System.Windows.Forms.TextBox]$OutText,
        [scriptblock]$TimestampFunction,
        [string]$TestingWebhookURL
    )

    $SubmitFeedbackScript = Join-Path -Path $PSScriptRoot -ChildPath $ConfigValues.SubmitFeedbackScript
    $UserFeedback = $FeedbackText

    if ($IsUserNameChecked) {
        $FeedbackUserName = ' - ' + $UserNameText
    } else {
        $FeedbackUserName = ' - ' + 'Anonymous'
    }

    # Create a runspace to execute the script in a separate thread and keep the main GUI responsive
    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.Open()

    $psCmd = [PowerShell]::Create().AddScript({
        param (
            [string]$SubmitFeedbackScript,
            [string]$TestingWebhookURL,
            [string]$UserFeedback,
            [string]$FeedbackUserName,
            [System.Windows.Forms.TextBox]$OutText,
            [scriptblock]$TimestampFunction
        )
        try {
            & $SubmitFeedbackScript -TestingWebhookURL $TestingWebhookURL -UserFeedback $UserFeedback -FeedbackUserName $FeedbackUserName -OutTextControl $OutText -TimestampFunction $TimestampFunction
        } catch {
            $OutText.AppendText("$($TimestampFunction.Invoke()) - An unhandled exception occurred: $($_.Exception.Message)`r`n")
        }
    }).AddParameters(@{
        SubmitFeedbackScript = $SubmitFeedbackScript
        TestingWebhookURL = $TestingWebhookURL
        UserFeedback = $UserFeedback
        FeedbackUserName = $FeedbackUserName
        OutText = $OutText
        TimestampFunction = $TimestampFunction
    })

    # Start the script in the runspace
    $psCmd.Runspace = $runspace
    $null = $psCmd.BeginInvoke()

    # Register event to clean up resources once the script completes
    $psCmd.add_InvocationStateChanged({
        if ($_.InvocationStateInfo.State -eq [System.Management.Automation.PSInvocationState]::Completed) {
            $psCmd.Dispose()
            $runspace.Close()
            $runspace.Dispose()
        }
    })
}

# Function to evaluate if the Theme Builder buttons should be enabled
function Enable-ThemeBuilderButtons {
    if ($null -ne $script:ThemeBuilderBackColorTextBox.Text -and $script:ThemeBuilderBackColorTextBox.Text -ne '' -and $null -ne $script:ThemeBuilderForeColorTextBox.Text -and $script:ThemeBuilderForeColorTextBox.Text -ne '' -and $null -ne $script:ThemeBuilderAccentColorTextBox.Text -and $script:ThemeBuilderAccentColorTextBox.Text -ne '' -and $null -ne $script:ThemeBuilderDisabledColorTextBox.Text -and $script:ThemeBuilderDisabledColorTextBox.Text -ne '') {
        $script:ThemeBuilderApplyThemeButton.Enabled = $true
		$script:ThemeBuilderSaveThemeButton.Enabled = $true
		$script:ThemeBuilderResetThemeButton.Enabled = $true
		if ($global:IsThemeApplied) {
			$script:ThemeBuilderApplyThemeButton.BackColor = $script:CustomForeColor
			$script:ThemeBuilderApplyThemeButton.ForeColor = $script:CustomBackColor
			$script:ThemeBuilderSaveThemeButton.BackColor = $script:CustomForeColor
			$script:ThemeBuilderSaveThemeButton.ForeColor = $script:CustomBackColor
			$script:ThemeBuilderResetThemeButton.BackColor = $script:CustomForeColor
			$script:ThemeBuilderResetThemeButton.ForeColor = $script:CustomBackColor
		}
    }
    else {
        $script:ThemeBuilderApplyThemeButton.Enabled = $false
        $script:ThemeBuilderSaveThemeButton.Enabled = $false
        $script:ThemeBuilderResetThemeButton.Enabled = $false
		if ($global:IsThemeApplied) {
			$script:ThemeBuilderApplyThemeButton.BackColor = $script:CustomDisabledColor
			$script:ThemeBuilderSaveThemeButton.BackColor = $script:CustomDisabledColor
			$script:ThemeBuilderResetThemeButton.BackColor = $script:CustomDisabledColor
		}
    }
}

# Function to evaluate if the user-entered color is valid
function Test-ColorInput {
    param (
        [string]$ColorInput,
        [string]$ColorType
    )

    # Check if the input is a valid named color
    $namedColors = [enum]::GetNames([System.Drawing.KnownColor])
    if ($ColorInput -in $namedColors) {
        return $true
    }

    # Check if the input is a valid hex color (#RRGGBB)
    if ($ColorInput -match '^#[0-9a-fA-F]{6}$') {
        return $true
    }

    return $false
}

# Function to evaluate if the Theme Builder Save Theme buttons should be enabled
function Enable-SaveThemeButtons {
    if ($null -ne $script:ThemeBuilderThemeNameTextBox.Text -and $script:ThemeBuilderThemeNameTextBox.Text -ne '' -and $null -ne $script:ThemeBuilderQuoteTextBox.Text -and $script:ThemeBuilderQuoteTextBox.Text -ne '') {
        $script:ThemeBuilderSaveThemePopupButton.Enabled = $true
		$script:ThemeBuilderSaveAndApplyThemeButton.Enabled = $true
		if ($global:IsThemeApplied) {
			$script:ThemeBuilderSaveThemePopupButton.BackColor = $script:CustomForeColor
			$script:ThemeBuilderSaveThemePopupButton.ForeColor = $script:CustomBackColor
			$script:ThemeBuilderSaveAndApplyThemeButton.BackColor = $script:CustomForeColor
			$script:ThemeBuilderSaveAndApplyThemeButton.ForeColor = $script:CustomBackColor
		}
    }
    else {
        $script:ThemeBuilderSaveThemePopupButton.Enabled = $false
        $script:ThemeBuilderSaveAndApplyThemeButton.Enabled = $false
		if ($global:IsThemeApplied) {
			$script:ThemeBuilderSaveThemePopupButton.BackColor = $script:CustomDisabledColor
			$script:ThemeBuilderSaveAndApplyThemeButton.BackColor = $script:CustomDisabledColor
		}
    }
}

# Function for the Theme Builder Save Theme popup
Function Save-CustomTheme {
    param (
        [string]$ThemeName,
        [string]$BackColor,
        [string]$ForeColor,
        [string]$AccentColor,
        [string]$DisabledColor,
        [string]$Quote,
        [ref]$ColorTheme
    )

    try {
        # Convert PSCustomObject to Hashtable if needed
        if ($ColorTheme.Value.Custom -is [PSCustomObject]) {
            $tempHashTable = @{}
            $ColorTheme.Value.Custom.PSObject.Properties | ForEach-Object {
                $tempHashTable[$_.Name] = $_.Value
            }
            $ColorTheme.Value.Custom = $tempHashTable
        }

        # Create new theme object with ordered hashtable
        $CustomTheme = [ordered]@{
            BackColor      = $BackColor
            ForeColor      = $ForeColor
            AccentColor    = $AccentColor
            DisabledColor  = $DisabledColor
            Quote          = $Quote
        }

        # Check if the "Custom" section is null, and initialize if needed
        if ($null -eq $ColorTheme.Value.Custom) {
            $ColorTheme.Value.Custom = @{}
        }

        # Check if the theme name already exists
        if ($ThemeName -in $ColorTheme.Value.Custom.Keys) {
            $script:ThemeBuilderOutText.AppendText("$(Get-Timestamp) - Theme name '$ThemeName' already exists. Please choose a different name.`r`n")
            return $false
        }

        # Add the new theme to the "Custom" section
        $ColorTheme.Value.Custom[$ThemeName] = $CustomTheme

        # Sort the "Custom" section alphabetically by key
        $sortedCustom = [ordered]@{}
        $ColorTheme.Value.Custom.Keys | Sort-Object | ForEach-Object {
            $sortedCustom[$_]= $ColorTheme.Value.Custom[$_]
        }
        $ColorTheme.Value.Custom = $sortedCustom

        # Convert the updated PowerShell object back to a JSON string
        $jsonString = $ColorTheme.Value | ConvertTo-Json -Depth 5

        # Write the updated JSON string back to the file
        Set-Content -Path .\ColorThemes.json -Value $jsonString

        $script:ThemeBuilderOutText.AppendText("$(Get-Timestamp) - Custom theme $script:NewThemeName was saved successfully.`r`n")
        return $true
    } catch {
        $script:ThemeBuilderOutText.AppendText("$(Get-Timestamp) - An error occurred while saving the theme: $_`r`n")
        return $false
    }
}

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

# Function to open a runspace and run Resolve-DnsName
function Open-NSLookupRunspace {
    param (
        [System.Windows.Forms.TextBox]$OutText,
        [System.Windows.Forms.TextBox]$NSLookupTextBox
    )

    try {
        $NSLookupRunspace = [runspacefactory]::CreateRunspace()
        $NSLookupRunspace.Open()

        $synchash = @{}
        $synchash.OutText = $OutText

        $NSLookupRunspace.SessionStateProxy.SetVariable("synchash", $synchash)

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
                            $synchash.OutText.AppendText("$(Get-Timestamp) - $($args[0]) $IPAddressType = $($SelectedObject.IPAddress)`r`n")
                        }
                    }
                }
                catch {
                    $synchash.OutText.AppendText("$(Get-Timestamp) - An error occurred while resolving the DNS name. Please ensure you're entering a valid hostname.`r`n")
                }
            }
            else {
                $synchash.OutText.AppendText("$(Get-Timestamp) - Please ensure you're entering a valid hostname`r`n")
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

# Function for opening a runspace and running Test-Connection
function Open-ServerPingRunspace {
    param (
        [System.Windows.Forms.TextBox]$OutText,
        [System.Windows.Forms.TextBox]$ServerPingTextBox
    )
    
    try {
        $ServerPingRunspace = [runspacefactory]::CreateRunspace()
        $ServerPingRunspace.Open()
        
        $synchash = @{}
        $synchash.OutText = $OutText

        $ServerPingRunspace.SessionStateProxy.SetVariable("synchash", $synchash)

        $psCmd = [PowerShell]::Create().AddScript({
            function Get-Timestamp {
                return Get-Date -Format "yyyy/MM/dd hh:mm:ss"
            }

            $synchash.OutText.AppendText("$(Get-Timestamp) - Testing connection to $($args[0])...`r`n")
            $PingResult = Test-Connection -ComputerName $args[0] -Quiet
            if ($PingResult) {
                $synchash.OutText.AppendText("$(Get-Timestamp) - Connection to $($args[0]) successful.`r`n")
            }
            else {
                $synchash.OutText.AppendText("$(Get-Timestamp) - Connection to $($args[0]) failed.`r`n")
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

# Function for resetting the selected environment
Function Reset-Environment {
    $PSTPath = Join-Path -Path $ResolvedLocalSupportTool -ChildPath "ProductionSupportTool.exe.config"
    $PSTOldPath = Join-Path -Path $ResolvedLocalSupportTool -ChildPath "ProductionSupportTool.exe.config.old"
    
    if (Test-Path $PSTPath) {
        Remove-Item -Path $PSTPath
    }
    
    if (Test-Path $PSTOldPath) {
        Rename-Item $PSTOldPath -NewName $PSTPath
    }
    <# Remove-Item -Path "$ResolvedLocalSupportTool\ProductionSupportTool.exe.config"
    Rename-Item "$ResolvedLocalSupportTool\ProductionSupportTool.exe.config.old" -NewName "$ResolvedLocalSupportTool\ProductionSupportTool.exe.config" #>
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

# Function for opening a runspace and running Refresh/Import PST files
function Update-PSTFiles {
    param (
        [string]$Team,
        [string]$Category
    )
    
    $RefreshPSTRunspace = [runspacefactory]::CreateRunspace()
    $RefreshPSTRunspace.Open()
    $synchash.OutText = $OutText
    $synchash.CurrentTeamBackColor = $global:CurrentTeamBackColor
    $synchash.CurrentTeamForeColor = $global:CurrentTeamForeColor
    $synchash.CurrentTeamAccentColor = $global:CurrentTeamAccentColor
    $synchash.ResolvedLocalSupportTool = $ResolvedLocalSupportTool
    $synchash.ResolvedLocalConfigs = $ResolvedLocalConfigs
    $synchash.RefreshPSTButton = $RefreshPSTButton
    $synchash.PSTCombo = $PSTCombo
    $synchash.SelectedTheme = $SelectedTheme
    $synchash.ThemeColors = $themeColors
    $synchash.DisabledBackColor = $global:DisabledBackColor
    $synchash.DisabledForeColor = $global:DisabledForeColor
    $RefreshPSTRunspace.SessionStateProxy.SetVariable("synchash", $synchash)

    $psCmd = [PowerShell]::Create().AddScript({
        param($synchash, $RemoteSupportTool, $LocalSupportTool, $RemoteConfigs, $LocalConfigs)

        # Initilize the timestamp function in the runspace
        function Get-Timestamp {
            return Get-Date -Format "yyyy/MM/dd hh:mm:ss"
        }

        # Define the Update-ButtonTheme function inside the script block
        function Update-ButtonTheme {
            param([hashtable]$synchash)
            
            if ($synchash['WasMainThemeActivated']) {
                $synchash.RefreshPSTButton.BackColor = $synchash.CurrentTeamForeColor
                $synchash.RefreshPSTButton.ForeColor = $synchash.CurrentTeamBackColor
                # Set the flag back to $false after updating the theme
                $synchash['WasMainThemeActivated'] = $false
                $synchash.PSTCombo.Enabled = $true
                if ($synchash.SelectedTheme -eq 'Custom' -or $synchash.SelectedTheme -eq 'Premium') {
                    $synchash.PSTCombo.BackColor = $synchash.CurrentTeamAccentColor
                }
            } else {
                $synchash.RefreshPSTButton.BackColor = $synchash.ThemeColors.ForeColor
                $synchash.RefreshPSTButton.ForeColor = $synchash.ThemeColors.BackColor
                $synchash.PSTCombo.Enabled = $true
                if ($synchash.SelectedTheme -eq 'Custom' -or $synchash.SelectedTheme -eq 'Premium') {
                    $synchash.PSTCombo.BackColor = $synchash.ThemeColors.AccentColor
                }

            }
        }

        # Disable button with appropriate colors
        $synchash.PSTCombo.Enabled = $false
        $synchash.RefreshPSTButton.Enabled = $false
        $synchash.RefreshPSTButton.BackColor = $synchash.DisabledBackColor
        $synchash.RefreshPSTButton.ForeColor = $synchash.DisabledForeColor

        if (!(Test-Path $LocalConfigs)) {
            $synchash.OutText.AppendText("$(Get-Timestamp) - Local Configs folder does not exist. Creating folder and copying files...`r`n")
            Copy-Item -Path $RemoteConfigs* -Destination $LocalConfigs -Recurse -Force
            $synchash.OutText.AppendText("$(Get-Timestamp) - Local Configs folder created successfully.`r`n")
        }
        elseif ((Get-Item $RemoteConfigs).LastWriteTime -gt (Get-Item $LocalConfigs).LastWriteTime) {
            $synchash.OutText.AppendText("$(Get-Timestamp) - Local Configs folder is out of date. Updating folder...`r`n")
            Copy-Item -Path $RemoteConfigs* -Destination $LocalConfigs -Recurse -Force
            $synchash.OutText.AppendText("$(Get-Timestamp) - Local Configs folder updated successfully.`r`n")
        }
        else {
            $synchash.OutText.AppendText("$(Get-Timestamp) - Local Configs folder is up to date.`r`n")
        }

        if (-not (Test-Path $LocalSupportTool)) {
            $synchash.OutText.AppendText("$(Get-Timestamp) - Local Prod Support Tool folder does not exist. Creating folder and copying files. This will take a few minutes...`r`n")
            $synchash.RefreshPSTButton.Text = "Copying PST Files..."

            Copy-Item -Path $RemoteSupportTool* -Destination $LocalSupportTool -Recurse -Force

            $synchash.RefreshPSTButton.Text = "Refresh PST Files"
            $synchash.OutText.AppendText("$(Get-Timestamp) - Local Prod Support Tool folder created successfully. You can now use the Prod Support Tool.`r`n")
        } elseif ($RemoteSupportToolFolderDate -gt $LocalSupportToolFolderDate) {
            $synchash.OutText.AppendText("$(Get-Timestamp) - Local Prod Support Tool folder is out of date. Updating folder...`r`n")
            $synchash.RefreshPSTButton.Text = "Copying PST Files..."

            # Get all files from the remote directory
            $remoteFiles = Get-ChildItem -Path $RemoteSupportTool -Recurse -File

            foreach ($file in $remoteFiles) {
                # Construct the full path for the corresponding local file
                $localFile = Join-Path -Path $LocalSupportTool -ChildPath $file.Name

                # Check if the local file exists and if the remote file is newer
                if (-not (Test-Path $localFile) -or (Get-Item $localFile).LastWriteTime -lt $file.LastWriteTime) {
                    # Copy the newer file to the local directory
                    Copy-Item -Path $file.FullName -Destination $localFile -Force
                    $synchash.OutText.AppendText("$(Get-Timestamp) - Copied updated file: $($file.Name)`r`n")
                }
        }
        $synchash.RefreshPSTButton.Text = "Refresh PST Files"
        $synchash.OutText.AppendText("$(Get-Timestamp) - Local Prod Support Tool folder updated successfully.`r`n")
            }
            else {
                $synchash.OutText.AppendText("$(Get-Timestamp) - Local Prod Support Tool folder is up to date.`r`n")
            }
        # Enable buttons with appropriate colors
        Update-ButtonTheme -synchash $synchash
        $synchash.RefreshPSTButton.Enabled = $true
        $synchash.PSTCombo.Enabled = $true
        # Update Resolved Paths
        $LocalSupportTool = Join-Path -Path $PSScriptRoot -ChildPath $ConfigValues.LocalSupportTool
        if (Test-Path $LocalSupportTool) {
            $synchash.ResolvedLocalSupportTool = Resolve-Path $LocalSupportTool
        } else {
            $synchash.ResolvedLocalSupportTool = $null
        }
        $LocalConfigs = Join-Path -Path $PSScriptRoot -ChildPath $ConfigValues.LocalConfigs
        if (Test-Path $LocalConfigs) {
            $synchash.ResolvedLocalConfigs = Resolve-Path $LocalConfigs
        } else {
            $synchash.ResolvedLocalConfigs = $null
        }

        # Output the updated synchash for retrieval in the main runspace
        $synchash

    }).AddArgument($synchash).AddArgument($RemoteSupportTool).AddArgument($LocalSupportTool).AddArgument($RemoteConfigs).AddArgument($LocalConfigs)

    $psCmd.Runspace = $RefreshPSTRunspace
    
    # Start the asynchronous operation
    $AsyncResult = $psCmd.BeginInvoke()

    # Wait for the asynchronous operation to complete
    $AsyncResult.AsyncWaitHandle.WaitOne()

    # Retrieve the results of the runspace operation
    $synchash = $psCmd.EndInvoke($AsyncResult)

    # Update the main script variables with the results
    $script:ResolvedLocalSupportTool = $synchash.ResolvedLocalSupportTool
    $script:ResolvedLocalConfigs = $synchash.ResolvedLocalConfigs

    # Clean up
    $psCmd.Dispose()
    $RefreshPSTRunspace.Close()
    $RefreshPSTRunspace.Dispose()

    # Display the updated variables
    Write-Host "ResolvedLocalSupportTool: $script:ResolvedLocalSupportTool"
    Write-Host "ResolvedLocalConfigs: $script:ResolvedLocalConfigs"
}

# Function to evaluate if the run script button should be enabled
function Enable-AddLenderButton {
    $backColor = Get-AppropriateColor -ColorType "BackColor"
    $foreColor = Get-AppropriateColor -ColorType "ForeColor"
    if ($null -ne $script:LenderLFPCombo.SelectedItem) {
		if ($script:LenderLFPCombo.SelectedItem -eq 'Production'){
			if ($null -ne $script:LenderLFPCombo.SelectedItem -and $script:LenderLFPIdTextBox.Text -ne '' -and $script:LenderLFPTicketTextBox.Text -ne '' -and $script:ProductionCheckbox.Checked -eq $true){
				$script:AddLenderLFPButton.Enabled = $true
                $script:AddLenderLFPButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
                $script:AddLenderLFPButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
			}
			else {
				$script:AddLenderLFPButton.Enabled = $false
                $script:AddLenderLFPButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
                $script:AddLenderLFPButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
			}
		}
		elseif ($null -ne $script:LenderLFPCombo.SelectedItem -and $script:LenderLFPIdTextBox.Text -ne '' -and $script:LenderLFPTicketTextBox.Text -ne ''){
			$script:AddLenderLFPButton.Enabled = $true
            $script:AddLenderLFPButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
            $script:AddLenderLFPButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
		}
	}
	else {
		$script:AddLenderLFPButton.Enabled = $false
        $script:AddLenderLFPButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
        $script:AddLenderLFPButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
	}
	
}

# Function to open an asynchronous runspace and run the AddLenderScript
function Open-AddLenderScriptRunspace {
    param (
        [string]$AddLenderScript,
        [string]$LenderId,
        [string]$TicketNumber,
        [string]$Environment,
        [System.Windows.Forms.TextBox]$OutText,
        [scriptblock]$TimestampFunction,
        [scriptblock]$ActiveTicketsFunction
    )
    
    # Create a runspace to execute the script in a separate thread and keep the main GUI responsive
    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.Open()

    $psCmd = [PowerShell]::Create().AddScript({
        param (
            [string]$AddLenderScript,
            [string]$LenderId,
            [string]$TicketNumber,
            [string]$Environment,
            [System.Windows.Forms.TextBox]$OutText,
            [scriptblock]$TimestampFunction,
            [scriptblock]$ActiveTicketsFunction
        )

        try {
            . $AddLenderScript -LenderId $LenderId -TicketNumber $TicketNumber -Environment $Environment -OutTextControl $OutText -TimestampFunction $TimestampFunction -ActiveTicketsFunction $ActiveTicketsFunction
        } catch {
            $OutText.AppendText("$($TimestampFunction.Invoke()) - An error occurred while executing the AddLenderScript: $($_.Exception.Message)`r`n")
        }
    }).AddParameters(@{
        AddLenderScript = $AddLenderScript
        LenderId = $LenderId
        TicketNumber = $TicketNumber
        Environment = $Environment
        OutText = $OutText
        TimestampFunction = $TimestampFunction
    })

    # Start the script in the runspace
    $psCmd.Runspace = $runspace
    $null = $psCmd.BeginInvoke()

    # Register event to clean up resources once the script completes
    $psCmd.add_InvocationStateChanged({
        if ($_.InvocationStateInfo.State -eq [System.Management.Automation.PSInvocationState]::Completed) {
            $psCmd.Dispose()
            $runspace.Close()
            $runspace.Dispose()
        }
    })
}

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

# New documentation function
Function New-DocTemplate {
    param (
        [string]$DocTopic
    )
    # Variables
    $TemplateFile = Join-Path -Path $PSScriptRoot -ChildPath $ConfigValues.TemplateFile
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

    $DocumentationPath = Join-Path -Path $PSScriptRoot -ChildPath $ConfigValues.DocumentationPath
    New-item -Path $DocumentationPath -Name "$DocTopic" -ItemType "directory"
    Move-Item -Path $NewFile -Destination "$DocumentationPath\$DocTopic"
    Invoke-Item -Path "$DocumentationPath\$DocTopic\$DocTopic.docx"
    $OutText.AppendText("$(Get-Timestamp) - New document created at $DocumentationPath\$DocTopic\$DocTopic.docx`r`n")
}

# Function to evaluate if the run script button should be enabled
function Enable-CreateHDTStorageButton {
    if ($null -ne $script:HDTStoragePopup.Tag -and $script:DBServerTextBox.Text -ne '' -and $script:TableNameTextBox.Text -ne '' -and $script:SecurePasswordTextBox.Text -ne '') {
        $script:CreateHDTStorageButton.Enabled = $true

        $backColor = Get-AppropriateColor -ColorType "BackColor"
        $foreColor = Get-AppropriateColor -ColorType "ForeColor"

        $script:CreateHDTStorageButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
        $script:CreateHDTStorageButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
    }
    else {
        $script:CreateHDTStorageButton.Enabled = $false
        $script:CreateHDTStorageButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
        $script:CreateHDTStorageButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
    }
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

# Function to open an asynchronous runspace and create an HDDT storage table
function Open-CreateHDTStorageRunspace {
    param (
        [string]$ConfigValuesCreateHDTStorageScript,
        [string]$UserProfilePath,
        [string]$WorkhorseDirectoryPath,
        [string]$WorkhorseServer,
        [string]$FileTag,
        [string]$DBServerText,
        [string]$TableNameText,
        [pscredential]$LoginCredentials,
        [System.Windows.Forms.TextBox]$OutText,
        [scriptblock]$TimestampFunction
    )

    $CreateHDTStorageScript = Join-Path -Path $PSScriptRoot -ChildPath $ConfigValuesCreateHDTStorageScript
    $DestinationPath = $WorkhorseDirectoryPath
    $File = $FileTag
    $OutPath = Join-Path -Path "\\$($WorkhorseServer)" -ChildPath $DestinationPath
    $DestinationFile = Join-Path -Path $OutPath -ChildPath (Get-Item $File).Name
    $TableName = $TableNameText
    $SQLInstance = $DBServerText

    # Create a runspace to execute the script in a separate thread and keep the main GUI responsive
    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.Open()

    $psCmd = [PowerShell]::Create().AddScript({
        param (
            [string]$CreateHDTStorageScript,
            [string]$File,
            [string]$DestinationFile,
            [string]$OutPath,
            [System.Management.Automation.PSCredential]$LoginCredentials,
            [string]$SQLInstance,
            [string]$TableName,
            [System.Windows.Forms.TextBox]$OutText,
            [scriptblock]$TimestampFunction
        )

        try {
            & $CreateHDTStorageScript -File $File -DestinationFile $DestinationFile -OutPath $OutPath -LoginCredentials $LoginCredentials -Instance $SQLInstance -TableName $TableName -OutTextControl $OutText -TimestampFunction $TimestampFunction
        } catch {
            $OutText.AppendText("$($TimestampFunction.Invoke()) - An error occurred: $($_.Exception.Message)`r`n")
        }
    }).AddParameters(@{
        CreateHDTStorageScript = $CreateHDTStorageScript
        File = $File
        DestinationFile = $DestinationFile
        OutPath = $OutPath
        LoginCredentials = $LoginCredentials
        SQLInstance = $SQLInstance
        TableName = $TableName
        OutText = $OutText
        TimestampFunction = $TimestampFunction
    })

    # Start the script in the runspace
    $psCmd.Runspace = $runspace
    $null = $psCmd.BeginInvoke()

    # Register event to clean up resources once the script completes
    $psCmd.add_InvocationStateChanged({
        if ($_.InvocationStateInfo.State -eq [System.Management.Automation.PSInvocationState]::Completed) {
            $psCmd.Dispose()
            $runspace.Close()
            $runspace.Dispose()
        }
    })
}

<# # Function to perform a reverse IP lookup
function Get-IPAddress {
    param (
        [string]$IPAddress
    )

    try {
        $hostEntry = [System.Net.Dns]::GetHostEntry($IPAddress)
        return $hostEntry.HostName
    }
    catch [System.Net.Sockets.SocketException] {
        if ($_.Exception.ErrorCode -eq 11001) { # Host not found error code
            throw "No hostname found for IP address $IPAddress. It might not have a reverse DNS entry."
        } else {
            throw "An error occurred while performing a reverse IP lookup: $($_.Exception.Message)"
        }
    }
    catch {
        throw "An unexpected error occurred: $($_.Exception.Message)"
    }
}

# Function for obtaining an IP from a given hostname
function Invoke-ReverseIPLookup {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ip,
        [Parameter(Mandatory=$true)]
        [System.Windows.Forms.TextBox]$OutText
    )

    # Validate IP Address
    $isValidIP = [ipaddress]::TryParse($ip, [ref]$null)

    if ($isValidIP) {
        try {
            $hostname = Get-IPAddress -IPAddress $ip
            $OutText.AppendText("$(Get-Timestamp) - The hostname for IP address $ip is $hostname.`r`n")
        } catch {
            $OutText.AppendText("$(Get-Timestamp) - $_`r`n")
        }
    }
    else {
        $OutText.AppendText("$(Get-Timestamp) - The provided IP address ($ip) is invalid. Please enter a valid IP address.`r`n")
    }
} #>

# Function to perform a reverse IP lookup asynchronously
function Invoke-ReverseIPLookupRunspace {
    param (
        [string]$ip,
        [System.Windows.Forms.TextBox]$OutText,
        [scriptblock]$TimestampFunction
    )

    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.Open()

    $psCmd = [PowerShell]::Create().AddScript({
        param (
            [string]$ip,
            [System.Windows.Forms.TextBox]$OutText,
            [scriptblock]$TimestampFunction
        )

        function Get-IPAddress {
            param ([string]$IPAddress)
            try {
                # Await the asynchronous call to GetHostEntryAsync
                $hostEntryTask = [System.Net.Dns]::GetHostEntryAsync($IPAddress)
                while (-not $hostEntryTask.IsCompleted) {
                    Start-Sleep -Milliseconds 100
                }
                if ($hostEntryTask.Status -eq [System.Threading.Tasks.TaskStatus]::RanToCompletion) {
                    $hostEntry = $hostEntryTask.Result
                    # Ensure we are returning a string
                    return [string]$hostEntry.HostName
                } elseif ($hostEntryTask.Status -eq [System.Threading.Tasks.TaskStatus]::Faulted) {
                    # Extract the exception message from the faulted task
                    $errorMessage = $hostEntryTask.Exception.InnerException.Message
                    throw "DNS lookup failed: $errorMessage"
                } else {
                    throw "The task did not complete successfully. Status: $($hostEntryTask.Status)"
                }
            } catch {
                $errorMessage = $_.Exception.Message
                if (-not $errorMessage) {
                    $errorMessage = "An unspecified error occurred."
                }
                # Return error as a string to ensure consistency in return type
                return "Error: $errorMessage"
            }
        }                

        # Validate IP Address
        $isValidIP = [System.Net.IPAddress]::TryParse($ip, [ref]$null)

        if ($isValidIP) {
            $OutText.AppendText("$($TimestampFunction.Invoke()) - Starting reverse IP lookup for $ip...`r`n")
            $hostname = Get-IPAddress -IPAddress $ip
            if ($hostname -notlike "Error:*") {
                $OutText.AppendText("$($TimestampFunction.Invoke()) - The hostname for IP address $ip is $hostname.`r`n")
            } else {
                # If an error occurred, $hostname contains the error message
                $OutText.AppendText("$($TimestampFunction.Invoke()) - $hostname`r`n")
            }
        } else {
            $OutText.AppendText("$($TimestampFunction.Invoke()) - The provided IP address ($ip) is invalid. Please enter a valid IP address.`r`n")
        }
    }).AddParameters(@{
        ip = $ip
        OutText = $OutText
        TimestampFunction = $TimestampFunction
    })

    # Start the script in the runspace
    $psCmd.Runspace = $runspace
    $null = $psCmd.BeginInvoke()

    # Register event to clean up resources once the script completes
    $psCmd.add_InvocationStateChanged({
        if ($_.InvocationStateInfo.State -eq [System.Management.Automation.PSInvocationState]::Completed) {
            $psCmd.Dispose()
            $runspace.Close()
            $runspace.Dispose()
        }
    })
}

# Creates the active and completed tickets path variables on startup
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

# Function to populate active tickets list box
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

# Function to populate completed tickets list box
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

# Logic for turning off the rename functionality
function Set-RenameOff {
    $RenameTicketButton.Enabled = $false
    $RenameTicketTextBox.Text = ''
    $RenameTicketTextBox.Enabled = $false
    $FolderContentsListBox.Items.Clear()
}

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

# Function to perform the directory checks
function Get-DirectoryData {
    param($path, $excludePattern)
    $data = @{}

    if (Test-Path $path) {
        $data['LastWriteTime'] = Get-ChildItem $path -Exclude $excludePattern -Recurse |
                                 Sort-Object LastWriteTime -Descending |
                                 Select-Object -First 1 LastWriteTime

        $data['ItemCount'] = (Get-ChildItem $path -File | Measure-Object).Count
    } else {
        $data['Error'] = "$path not found."
    }

    return $data
}

# Function to start the job and handle its output
function Start-DirectoryCheckJob {
    param([string]$path, [string]$excludePattern, [string]$type)

    $job = Start-Job -ScriptBlock ${function:Get-DirectoryData} -ArgumentList $path, $excludePattern

    # Use Register-ObjectEvent to handle job completion
    Register-ObjectEvent -InputObject $job -EventName StateChanged -Action {
        if ($job.State -eq 'Completed') {
            # Job is completed, retrieve results
            $result = Receive-Job -Job $job

            $script:OutText.Invoke({
                if ($result['Error']) {
                    $script:OutText.AppendText("$($result['Error'])`r`n")
                } else {
                    $timeStamp = Get-Timestamp
                    if ($type -eq 'Remote') {
                        $script:RemoteSupportToolFolderDate = $result['LastWriteTime']
                        $script:OutText.AppendText("$timeStamp - Remote folder last updated: $($result['LastWriteTime'])`r`n")
                    } elseif ($type -eq 'Local') {
                        $script:LocalSupportToolFolderDate = $result['LastWriteTime']
                        $script:LocalSupportToolFolderCount = $result['ItemCount']
                        $script:OutText.AppendText("$timeStamp - Local folder last updated: $($result['LastWriteTime']), item count: $($result['ItemCount'])`r`n")
                    }
                }
            })

            # Clean up the job
            Remove-Job -Job $job
        }
    } | Out-Null # We don't need the event subscriber object
}

function Enable-PrideTheme {
    # Define Pride colors
    $prideColorsHex = @('#E40303', '#FF8C00', '#FFED00', '#008026', '#24408E', '#732982')

    # Convert hex color codes to Color objects
    $prideColorObjects = $prideColorsHex | ForEach-Object {
        [System.Drawing.ColorTranslator]::FromHtml($_)
    }

    # Check if the prideColorObjects array is populated correctly
    if ($prideColorObjects.Length -eq 0) {
        throw 'Pride color objects array is empty.'
    }

    # Define common properties for all controls
    $fontArialRegular9 = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
    $itemHeight = 15

    # An array of all controls that need to be owner-drawn
    $allControls = @($ActiveTicketsListBox, $CompletedTicketsListBox, $FolderContentsListBox, 
                     $ServersListBox, $ServicesListBox, $IISSitesListBox, $AppPoolsListBox, $AppListCombo)

    # Apply changes to all owner-drawn controls
    foreach ($control in $allControls) {
        $control.SuspendLayout() # Suspend layout logic
        $control.DrawMode = [System.Windows.Forms.DrawMode]::OwnerDrawFixed
        $control.Font = $fontArialRegular9
        $control.ItemHeight = $itemHeight

        # Remove existing DrawItem event handlers
        $control.remove_DrawItem($control.DrawItem)

        # Add the DrawItem event handler
        $control.add_DrawItem({
            param($s, $e)

            # Draw the background and focus rectangle
            $e.DrawBackground()
            $e.DrawFocusRectangle()

            # Calculate the initial X position for drawing
            $initialX = $e.Bounds.X

            # Draw each character in the item with the appropriate Pride color
            for ($i = 0; $i -lt $s.Items[$e.Index].ToString().Length; $i++) {
                $character = $s.Items[$e.Index].ToString()[$i]
                $colorIndex = $i % $prideColorObjects.Length
                $brush = New-Object System.Drawing.SolidBrush ($prideColorObjects[$colorIndex])

                # Measure the width of the character
                $characterSize = [System.Windows.Forms.TextRenderer]::MeasureText($character, $e.Font)
                $point = New-Object System.Drawing.PointF($initialX, $e.Bounds.Y)

                # Draw the character
                $e.Graphics.DrawString($character, $e.Font, $brush, $point)

                # Increment the X position by the width of the character
                $initialX += $characterSize.Width
            }

            # Draw the focus rectangle if the list box has focus
            $e.DrawFocusRectangle()
        })

        $control.ResumeLayout($false) # Resume layout logic
        $control.Refresh() # Refresh the control to apply changes
        $control.Invalidate() # Force a complete redraw of the control
    }
}

function Disable-PrideTheme {
    # Reset to the original theme
    # This function should revert all changes made by Enable-PrideTheme
    # Similar to the Disable-USAThemeTextColor function you provided earlier
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
$SysAdminTab.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)

# Support Tools Tab
$SupportTab = New-Object System.Windows.Forms.TabPage
$SupportTab.DataBindings.DefaultDataSourceUpdateMode = 0
$SupportTab.UseVisualStyleBackColor = $True
$SupportTab.Name = 'SupportTools'
$SupportTab.Text = 'Support Tools'
$SupportTab.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)

# Ticket Manager Tab
$TicketManagerTab = New-Object System.Windows.Forms.TabPage
$TicketManagerTab.DataBindings.DefaultDataSourceUpdateMode = 0
$TicketManagerTab.UseVisualStyleBackColor = $True
$TicketManagerTab.Name = 'TicketManager'
$TicketManagerTab.Text = 'Ticket Manager'
$TicketManagerTab.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)

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

<#
? **********************************************************************************************************************
? END OF MAIN GUI
? **********************************************************************************************************************
#>
<#
? **********************************************************************************************************************
? START OF MENU STRIP
? **********************************************************************************************************************
#>

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

# Click event for the Submit Feedback menu option
$SubmitFeedback.add_Click({
    $OutText.AppendText("$(Get-Timestamp) - Launching feedback form...`r`n")
    $script:FeedbackForm = New-Object System.Windows.Forms.Form
    $script:FeedbackForm.Text = "Submit Feedback"
    $script:FeedbackForm.Size = New-Object System.Drawing.Size(300, 350)
    $script:FeedbackForm.ShowInTaskbar = $False
    $script:FeedbackForm.KeyPreview = $True
    $script:FeedbackForm.AutoSize = $True
    $script:FeedbackForm.FormBorderStyle = "Fixed3D"
    $script:FeedbackForm.Icon = $Icon
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

    # Event handler for anonymous radio button to check if the submit button should be enabled
    $script:AnonymousRadioButton.add_CheckedChanged({
        Enable-SubmitFeedbackButton
    })

    # Event handler for the provide username button to check if the submit button should be enabled
    $script:UserNameRadioButton.add_CheckedChanged({
        Enable-SubmitFeedbackButton
    })

    # Event handler for feeback textbox to check if the submit button should be enabled
    $script:FeedbackTextBox.add_TextChanged({
        Enable-SubmitFeedbackButton
    })

    # Event handler for username textbox to check if the submit button should be enabled
    $script:UserNameTextBox.add_TextChanged({
        Enable-SubmitFeedbackButton
    })

    # Event handler for the submit feedback button
    $script:SubmitFeedbackButton.add_Click({
        Open-SubmitFeedbackRunspace -ConfigValuesSubmitFeedbackScript $ConfigValues.SubmitFeedbackScript -UserProfilePath $userProfilePath -FeedbackText $script:FeedbackTextBox.Text -IsUserNameChecked $script:UserNameRadioButton.Checked -UserNameText $script:UserNameTextBox.Text -OutText $OutText -TimestampFunction ${function:Get-Timestamp} -TestingWebhookURL $script:TestingWebhookURL
    })

    # Button click event for closing the feedback form
    $script:FeedbackForm.add_FormClosed({
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

# Click event for the File menu Quit option
$MenuQuit.add_Click({ $Form.Close() })

# Options menu
$OptionsMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$OptionsMenu.Text = "Options"
$MenuColorTheme = New-Object System.Windows.Forms.ToolStripMenuItem
$MenuColorTheme.Text = "Select Theme"
$MenuThemeBuilder = New-Object System.Windows.Forms.ToolStripMenuItem
$MenuThemeBuilder.Text = "Theme Builder"
$MenuToolTips = New-Object System.Windows.Forms.ToolStripMenuItem
$MenuToolTips.Text = "Tool Tips"
$ShowHelpBannerMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$ShowHelpBannerMenu.Text = "Show Help Banner"
$ShowToolTipsMenu = New-Object System.Windows.Forms.ToolStripMenuItem

# Color theme sub-menu
$script:CustomThemes = New-Object System.Windows.Forms.ToolStripMenuItem
$script:CustomThemes.Text = "Custom Themes"
# Initial population of the theme menu
foreach ($team in $ColorTheme.Custom.PSObject.Properties) {
    $MenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $MenuItem.Text = $team.Name
    $MenuItem.add_Click({
        $selectedTheme = $this.Text -replace "• ", ""  # Remove bullet point if present
        if ($selectedTheme -eq $ConfigValues.DefaultUserTheme) {
            $OutText.AppendText("$(Get-Timestamp) - The selected theme is already active.`r`n")
        }
        else {
            Update-MainTheme -Team $selectedTheme -Category 'Custom' -ColorData $ColorTheme
            $ConfigValues.DefaultUserTheme = $selectedTheme  # Update the current theme
            Update-ThemeMenuItems  -MenuCategories @($script:CustomThemes, $script:MLBThemes, $script:NBAThemes, $NFLThemes, $PremiumThemes)
        }
    })
    $script:CustomThemes.DropDownItems.Add($MenuItem) | Out-Null
}
$script:MLBThemes = New-Object System.Windows.Forms.ToolStripMenuItem
$script:MLBThemes.Text = "MLB Teams"
foreach ($team in $ColorTheme.MLB.PSObject.Properties) {
    $MenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $MenuItem.Text = $team.Name
    $MenuItem.add_Click({
        $selectedTheme = $this.Text -replace "• ", ""  # Remove bullet point if present
        if ($selectedTheme -eq $ConfigValues.DefaultUserTheme) {
            $OutText.AppendText("$(Get-Timestamp) - The selected theme is already active.`r`n")
        }
        else {
            Update-MainTheme -Team $selectedTheme -Category 'MLB' -ColorData $ColorTheme
            $ConfigValues.DefaultUserTheme = $selectedTheme  # Update the current theme
            Update-ThemeMenuItems  -MenuCategories @($script:CustomThemes, $script:MLBThemes, $script:NBAThemes, $NFLThemes, $PremiumThemes)
        }
    })
    $script:MLBThemes.DropDownItems.Add($MenuItem) | Out-Null
}
$script:NBAThemes = New-Object System.Windows.Forms.ToolStripMenuItem
$script:NBAThemes.Text = "NBA Teams"
foreach ($team in $ColorTheme.NBA.PSObject.Properties) {
    $MenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $MenuItem.Text = $team.Name
    $MenuItem.add_Click({
        $selectedTheme = $this.Text -replace "• ", ""  # Remove bullet point if present
        if ($selectedTheme -eq $ConfigValues.DefaultUserTheme) {
            $OutText.AppendText("$(Get-Timestamp) - The selected theme is already active.`r`n")
        }
        else {
            Update-MainTheme -Team $selectedTheme -Category 'NBA' -ColorData $ColorTheme
            $ConfigValues.DefaultUserTheme = $selectedTheme  # Update the current theme
            Update-ThemeMenuItems  -MenuCategories @($script:CustomThemes, $script:MLBThemes, $script:NBAThemes, $NFLThemes, $PremiumThemes)
        }
    })
    $script:NBAThemes.DropDownItems.Add($MenuItem) | Out-Null
}
$NFLThemes = New-Object System.Windows.Forms.ToolStripMenuItem
$NFLThemes.Text = "NFL Teams"
foreach ($team in $ColorTheme.NFL.PSObject.Properties) {
    $MenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $MenuItem.Text = $team.Name
    $MenuItem.add_Click({
        $selectedTheme = $this.Text -replace "• ", ""  # Remove bullet point if present
        if ($selectedTheme -eq $ConfigValues.DefaultUserTheme) {
            $OutText.AppendText("$(Get-Timestamp) - The selected theme is already active.`r`n")
        }
        else {
            Update-MainTheme -Team $selectedTheme -Category 'NFL' -ColorData $ColorTheme
            $ConfigValues.DefaultUserTheme = $selectedTheme  # Update the current theme
            Update-ThemeMenuItems  -MenuCategories @($script:CustomThemes, $script:MLBThemes, $script:NBAThemes, $NFLThemes, $PremiumThemes)
        }
    })
    $NFLThemes.DropDownItems.Add($MenuItem) | Out-Null
}
$PremiumThemes = New-Object System.Windows.Forms.ToolStripMenuItem
$PremiumThemes.Text = "Premium"
foreach ($team in $ColorTheme.Premium.PSObject.Properties) {
    $MenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $MenuItem.Text = $team.Name
    $MenuItem.add_Click({
        $selectedTheme = $this.Text -replace "• ", ""  # Remove bullet point if present
        if ($selectedTheme -eq $ConfigValues.DefaultUserTheme) {
            $OutText.AppendText("$(Get-Timestamp) - The selected theme is already active.`r`n")
        }
        else {
            Update-MainTheme -Team $selectedTheme -Category 'Premium' -ColorData $ColorTheme
            $ConfigValues.DefaultUserTheme = $selectedTheme  # Update the current theme
            Update-ThemeMenuItems  -MenuCategories @($script:CustomThemes, $script:MLBThemes, $script:NBAThemes, $NFLThemes, $PremiumThemes)
        }
    })
    $PremiumThemes.DropDownItems.Add($MenuItem) | Out-Null
}

# Call `Update-ThemeMenuItems` after all menus are initialized
Update-ThemeMenuItems -MenuCategories @($script:CustomThemes, $script:MLBThemes, $script:NBAThemes, $NFLThemes, $PremiumThemes)

# Event handler for Theme Builder menu option
$MenuThemeBuilder.add_Click({
    $OutText.AppendText("$(Get-Timestamp) - Launching Theme Builder...`r`n")

    # Set global variable to indicate that the Theme Builder popup is active
    $global:IsThemeBuilderPopupActive = $true
    
    # Theme Builder form
    $script:ThemeBuilderForm = New-Object System.Windows.Forms.Form
    $script:ThemeBuilderForm.Text = "Theme Builder"
    $script:ThemeBuilderForm.Size = New-Object System.Drawing.Size(750, 720)
    $script:ThemeBuilderForm.ShowInTaskbar = $True
    $script:ThemeBuilderForm.KeyPreview = $True
    $script:ThemeBuilderForm.AutoSize = $True
    $script:ThemeBuilderForm.FormBorderStyle = "Fixed3D"
    $script:ThemeBuilderForm.Icon = $Icon
    $script:ThemeBuilderForm.MaximizeBox = $False
    $script:ThemeBuilderForm.MinimizeBox = $True
    $script:ThemeBuilderForm.ControlBox = $True
    $script:ThemeBuilderForm.StartPosition = "CenterScreen"

    # Theme Builder Tab control creation
    $script:ThemeBuilderMainFormTabControl = New-object System.Windows.Forms.TabControl
    $script:ThemeBuilderMainFormTabControl.Size = "442,375"
    $script:ThemeBuilderMainFormTabControl.Location = "4,49"

    # Theme Builder System Administrator Tools Tab
    $script:ThemeBuilderSysAdminTab = New-Object System.Windows.Forms.TabPage
    $script:ThemeBuilderSysAdminTab.DataBindings.DefaultDataSourceUpdateMode = 0
    $script:ThemeBuilderSysAdminTab.UseVisualStyleBackColor = $True
    $script:ThemeBuilderSysAdminTab.Name = 'ThemeBuilderSysAdminTools'
    $script:ThemeBuilderSysAdminTab.Text = 'SysAdmin Tools'
    $script:ThemeBuilderSysAdminTab.Font = [System.Drawing.Font]::new("Arial", 8, [System.Drawing.FontStyle]::Regular)

    # Theme Builder textbox for logging
    $script:ThemeBuilderOutText = New-Object System.Windows.Forms.TextBox
    $script:ThemeBuilderOutText.Size = New-Object System.Drawing.Size(300, 335)
    $script:ThemeBuilderOutText.Location = New-Object System.Drawing.Point(450, 64)
    $script:ThemeBuilderOutText.ScrollBars = "Vertical"
    $script:ThemeBuilderOutText.Multiline = $true
    $script:ThemeBuilderOutText.Enabled = $true
    $script:ThemeBuilderOutText.ReadOnly = $true
    $script:ThemeBuilderOutText.Font = [System.Drawing.Font]::new("Arial", 8, [System.Drawing.FontStyle]::Regular)
    $script:ThemeBuilderOutText.AppendText("$(Get-Timestamp) - Foreground colors are used for labels and buttons.`r`n")
    $script:ThemeBuilderOutText.AppendText("$(Get-Timestamp) - Disabled colors are for disabled buttons, like the 'Start' and 'Clear Output' buttons bordering this box.`r`n")

    # Theme Builder button for clearing logging text box output
    $script:ThemeBuilderClearOutTextButton = New-Object System.Windows.Forms.Button
    $script:ThemeBuilderClearOutTextButton.Location = New-Object System.Drawing.Point(495, 400)
    $script:ThemeBuilderClearOutTextButton.Width = 75
    $script:ThemeBuilderClearOutTextButton.FlatStyle = "Popup"
    $script:ThemeBuilderClearOutTextButton.Text = "Clear Output"
    $script:ThemeBuilderClearOutTextButton.Font = [System.Drawing.Font]::new("Arial", 8, [System.Drawing.FontStyle]::Regular)
    $script:ThemeBuilderClearOutTextButton.Enabled = $false

    # Theme Builder button for saving logging text box output
    $script:ThemeBuilderSaveOutTextButton = New-Object System.Windows.Forms.Button
    $script:ThemeBuilderSaveOutTextButton.Location = New-Object System.Drawing.Point(630, 400)
    $script:ThemeBuilderSaveOutTextButton.Width = 75
    $script:ThemeBuilderSaveOutTextButton.FlatStyle = "Popup"
    $script:ThemeBuilderSaveOutTextButton.Text = "Save Output"
    $script:ThemeBuilderSaveOutTextButton.Font = [System.Drawing.Font]::new("Arial", 8, [System.Drawing.FontStyle]::Regular)
    $script:ThemeBuilderSaveOutTextButton.Enabled = $true

    # Theme Builder Restarts tab control creation
    $ThemeBuilderRestartsTabControl = New-object System.Windows.Forms.TabControl
    $ThemeBuilderRestartsTabControl.Size = "187,187"
    $ThemeBuilderRestartsTabControl.Location = "169,56"

    # Theme Builder Individual servers list box
    $script:ThemeBuilderServersListBox = New-Object System.Windows.Forms.ListBox
    $script:ThemeBuilderServersListBox.Location = New-Object System.Drawing.Point(4,71)
    $script:ThemeBuilderServersListBox.Size = New-Object System.Drawing.Size(150,180)
    $script:ThemeBuilderServersListBox.SelectionMode = 'One'
    $script:ThemeBuilderServersListBox.Font = [System.Drawing.Font]::new("Arial", 8, [System.Drawing.FontStyle]::Regular)
    $script:ThemeBuilderServersListBox.Items.Add("Accent colors are used")
    $script:ThemeBuilderServersListBox.Items.Add("for list boxes like this.")

    # Theme Builder Services list box
    $script:ThemeBuilderServicesListBox = New-Object System.Windows.Forms.ListBox
    $script:ThemeBuilderServicesListBox.Location = New-Object System.Drawing.Point(0,0)
    $script:ThemeBuilderServicesListBox.Size = New-Object System.Drawing.Size(184,180)
    $script:ThemeBuilderServicesListBox.SelectionMode = 'MultiExtended'
    $script:ThemeBuilderServicesListBox.Font = [System.Drawing.Font]::new("Arial", 8, [System.Drawing.FontStyle]::Regular)
    $script:ThemeBuilderServicesListBox.Items.Add("Accent colors are also")
    $script:ThemeBuilderServicesListBox.Items.Add("used for text boxes.")

    # Theme Builder Combobox for application selection
    $script:ThemeBuilderAppListCombo = New-Object System.Windows.Forms.ComboBox
    $script:ThemeBuilderAppListCombo.Location = New-Object System.Drawing.Point(4,49)
    $script:ThemeBuilderAppListCombo.Size = New-Object System.Drawing.Size(150, 150)
    $script:ThemeBuilderAppListCombo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $script:ThemeBuilderAppListCombo.Font = [System.Drawing.Font]::new("Arial", 8, [System.Drawing.FontStyle]::Regular)

    # Populate the ThemeBuilderAppListCombo with the list of applications from the CSV
    foreach ($header in $csvHeaders) {
        [void]$script:ThemeBuilderAppListCombo.Items.Add($header)
    }

    # Theme Builder Label applist combo box
    $script:ThemeBuilderAppListLabel = New-Object System.Windows.Forms.Label
    $script:ThemeBuilderAppListLabel.Location = New-Object System.Drawing.Size(4, 30)
    $script:ThemeBuilderAppListLabel.Size = New-Object System.Drawing.Size(113, 15)
    $script:ThemeBuilderAppListLabel.Font = [System.Drawing.Font]::new("Arial", 8, [System.Drawing.FontStyle]::Bold)
    $script:ThemeBuilderAppListLabel.Text = 'Select a Server'

    # ThemeBuilder tab for services list
    $ThemeBuilderServicesTab = New-Object System.Windows.Forms.TabPage
    $ThemeBuilderServicesTab.DataBindings.DefaultDataSourceUpdateMode = 0
    $ThemeBuilderServicesTab.UseVisualStyleBackColor = $True
    $ThemeBuilderServicesTab.Name = 'ServicesTab'
    $ThemeBuilderServicesTab.Text = 'Services'
    $ThemeBuilderServicesTab.Font = [System.Drawing.Font]::new("Arial", 8, [System.Drawing.FontStyle]::Regular)

    # ThemeBuilder Button for restarting services
    $script:ThemeBuilderRestartButton = New-Object System.Windows.Forms.Button
    $script:ThemeBuilderRestartButton.Location = New-Object System.Drawing.Point(370, 56)
    $script:ThemeBuilderRestartButton.Width = 56
    $script:ThemeBuilderRestartButton.FlatStyle = "Popup"
    $script:ThemeBuilderRestartButton.Text = "Restart"
    $script:ThemeBuilderRestartButton.Font = [System.Drawing.Font]::new("Arial", 8, [System.Drawing.FontStyle]::Regular)
    $script:ThemeBuilderRestartButton.Enabled = $true

    # ThemeBuilder Button for starting services
    $script:ThemeBuilderStartButton = New-Object System.Windows.Forms.Button
    $script:ThemeBuilderStartButton.Location = New-Object System.Drawing.Point(370, 86)
    $script:ThemeBuilderStartButton.Width = 56
    $script:ThemeBuilderStartButton.FlatStyle = "Popup"
    $script:ThemeBuilderStartButton.Text = "Start"
    $script:ThemeBuilderStartButton.Font = [System.Drawing.Font]::new("Arial", 8, [System.Drawing.FontStyle]::Regular)
    $script:ThemeBuilderStartButton.Enabled = $false

    # ThemeBuilder Button for stopping services
    $script:ThemeBuilderStopButton = New-Object System.Windows.Forms.Button
    $script:ThemeBuilderStopButton.Location = New-Object System.Drawing.Point(370, 116)
    $script:ThemeBuilderStopButton.Width = 56
    $script:ThemeBuilderStopButton.FlatStyle = "Popup"
    $script:ThemeBuilderStopButton.Text = "Stop"
    $script:ThemeBuilderStopButton.Font = [System.Drawing.Font]::new("Arial", 8, [System.Drawing.FontStyle]::Regular)
    $script:ThemeBuilderStopButton.Enabled = $true

    # Label for entering text in color text boxes
    $script:ThemeBuilderHelpLabel = New-Object System.Windows.Forms.Label
    $script:ThemeBuilderHelpLabel.Location = New-Object System.Drawing.Point(0, 450)
    $script:ThemeBuilderHelpLabel.Size = New-Object System.Drawing.Size(750, 20)
    $script:ThemeBuilderHelpLabel.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Bold)
    $script:ThemeBuilderHelpLabel.Text = 'Enter colors in Hex format (e.g. #0000FF) or as a color name (e.g. Blue)'
    $script:ThemeBuilderHelpLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter

    # Text box for entering BackColor
    $script:ThemeBuilderBackColorTextBox = New-Object System.Windows.Forms.TextBox
    $script:ThemeBuilderBackColorTextBox.Location = New-Object System.Drawing.Size(60, 510)
    $script:ThemeBuilderBackColorTextBox.Size = New-Object System.Drawing.Size(100, 20)

    # Label for BackColor text box
    $script:ThemeBuilderBackColorLabel = New-Object System.Windows.Forms.Label
    $script:ThemeBuilderBackColorLabel.Location = New-Object System.Drawing.Size(60, 490)
    $script:ThemeBuilderBackColorLabel.Size = New-Object System.Drawing.Size(100, 20)
    $script:ThemeBuilderBackColorLabel.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Bold)
    $script:ThemeBuilderBackColorLabel.Text = 'Background'

    # Text box for entering ForeColor
    $script:ThemeBuilderForeColorTextBox = New-Object System.Windows.Forms.TextBox
    $script:ThemeBuilderForeColorTextBox.Location = New-Object System.Drawing.Size(236, 510)
    $script:ThemeBuilderForeColorTextBox.Size = New-Object System.Drawing.Size(100, 20)

    # Label for ForeColor text box
    $script:ThemeBuilderForeColorLabel = New-Object System.Windows.Forms.Label
    $script:ThemeBuilderForeColorLabel.Location = New-Object System.Drawing.Size(236, 490)
    $script:ThemeBuilderForeColorLabel.Size = New-Object System.Drawing.Size(100, 20)
    $script:ThemeBuilderForeColorLabel.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Bold)
    $script:ThemeBuilderForeColorLabel.Text = 'Foreground'

    # Text box for entering AccentColor
    $script:ThemeBuilderAccentColorTextBox = New-Object System.Windows.Forms.TextBox
    $script:ThemeBuilderAccentColorTextBox.Location = New-Object System.Drawing.Size(412, 510)
    $script:ThemeBuilderAccentColorTextBox.Size = New-Object System.Drawing.Size(100, 20)

    # Label for AccentColor text box
    $script:ThemeBuilderAccentColorLabel = New-Object System.Windows.Forms.Label
    $script:ThemeBuilderAccentColorLabel.Location = New-Object System.Drawing.Size(412, 490)
    $script:ThemeBuilderAccentColorLabel.Size = New-Object System.Drawing.Size(100, 20)
    $script:ThemeBuilderAccentColorLabel.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Bold)
    $script:ThemeBuilderAccentColorLabel.Text = 'Accent'

    # Text box for entering DisabledColor
    $script:ThemeBuilderDisabledColorTextBox = New-Object System.Windows.Forms.TextBox
    $script:ThemeBuilderDisabledColorTextBox.Location = New-Object System.Drawing.Size(588, 510)
    $script:ThemeBuilderDisabledColorTextBox.Size = New-Object System.Drawing.Size(100, 20)

    # Label for DisabledColor text box
    $script:ThemeBuilderDisabledColorLabel = New-Object System.Windows.Forms.Label
    $script:ThemeBuilderDisabledColorLabel.Location = New-Object System.Drawing.Size(588, 490)
    $script:ThemeBuilderDisabledColorLabel.Size = New-Object System.Drawing.Size(100, 20)
    $script:ThemeBuilderDisabledColorLabel.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Bold)
    $script:ThemeBuilderDisabledColorLabel.Text = 'Disabled'

    # Button for applying user-defined colors
    $script:ThemeBuilderApplyThemeButton = New-Object System.Windows.Forms.Button
    $script:ThemeBuilderApplyThemeButton.Location = New-Object System.Drawing.Point(120, 560)
    $script:ThemeBuilderApplyThemeButton.Width = 100
    $script:ThemeBuilderApplyThemeButton.FlatStyle = "Popup"
    $script:ThemeBuilderApplyThemeButton.Text = "Apply Theme"
    $script:ThemeBuilderApplyThemeButton.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
    $script:ThemeBuilderApplyThemeButton.Enabled = $false

    # Button for saving user-defined colors
    $script:ThemeBuilderSaveThemeButton = New-Object System.Windows.Forms.Button
    $script:ThemeBuilderSaveThemeButton.Location = New-Object System.Drawing.Point(325, 560)
    $script:ThemeBuilderSaveThemeButton.Width = 100
    $script:ThemeBuilderSaveThemeButton.FlatStyle = "Popup"
    $script:ThemeBuilderSaveThemeButton.Text = "Save Theme"
    $script:ThemeBuilderSaveThemeButton.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
    $script:ThemeBuilderSaveThemeButton.Enabled = $false

    # Button for clearing all text boxes
    $script:ThemeBuilderResetThemeButton = New-Object System.Windows.Forms.Button
    $script:ThemeBuilderResetThemeButton.Location = New-Object System.Drawing.Point(530, 560)
    $script:ThemeBuilderResetThemeButton.Width = 100
    $script:ThemeBuilderResetThemeButton.FlatStyle = "Popup"
    $script:ThemeBuilderResetThemeButton.Text = "Reset Theme"
    $script:ThemeBuilderResetThemeButton.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
    $script:ThemeBuilderResetThemeButton.Enabled = $false

    # Button for deleting custom themes
    $script:ThemeBuilderDeleteThemesButton = New-Object System.Windows.Forms.Button
    $script:ThemeBuilderDeleteThemesButton.Location = New-Object System.Drawing.Point(300, 625)
    $script:ThemeBuilderDeleteThemesButton.Width = 150
    $script:ThemeBuilderDeleteThemesButton.FlatStyle = "Popup"
    $script:ThemeBuilderDeleteThemesButton.Text = "Delete Custom Themes"
    $script:ThemeBuilderDeleteThemesButton.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
    $script:ThemeBuilderDeleteThemesButton.Enabled = if ($ColorTheme.Custom.PSObject.Properties.Count -gt 0) {
        $true
    } else {
        $false
    }

    # Event handler for checking text changed in BackColor text box
    $script:ThemeBuilderBackColorTextBox.add_TextChanged({
        Enable-ThemeBuilderButtons
    })

    # Event handler for checking text changed in ForeColor text box
    $script:ThemeBuilderForeColorTextBox.add_TextChanged({
        Enable-ThemeBuilderButtons
    })

    # Event handler for checking text changed in AccentColor text box
    $script:ThemeBuilderAccentColorTextBox.add_TextChanged({
        Enable-ThemeBuilderButtons
    })

    # Event handler for checking text changed in DisabledColor text box
    $script:ThemeBuilderDisabledColorTextBox.add_TextChanged({
        Enable-ThemeBuilderButtons
    })

    # Event handler for clicking the Apply Theme button
    $script:ThemeBuilderApplyThemeButton.add_Click({

        $isValidBackColor = Test-ColorInput -ColorInput $script:ThemeBuilderBackColorTextBox.Text -ColorType 'BackColor'
        $isValidForeColor = Test-ColorInput -ColorInput $script:ThemeBuilderForeColorTextBox.Text -ColorType 'ForeColor'
        $isValidAccentColor = Test-ColorInput -ColorInput $script:ThemeBuilderAccentColorTextBox.Text -ColorType 'AccentColor'
        $isValidDisabledColor = Test-ColorInput -ColorInput $script:ThemeBuilderDisabledColorTextBox.Text -ColorType 'DisabledColor'
        
        # Check if all colors are valid
        if ($isValidBackColor -and $isValidForeColor -and $isValidAccentColor -and $isValidDisabledColor) {

        # Set global Is Theme Applied variable to true
        $global:IsThemeApplied = $true

        $script:CustomBackColor = $script:ThemeBuilderBackColorTextBox.Text
        $script:CustomForeColor = $script:ThemeBuilderForeColorTextBox.Text
        $script:CustomAccentColor = $script:ThemeBuilderAccentColorTextBox.Text
        $script:CustomDisabledColor = $script:ThemeBuilderDisabledColorTextBox.Text

        # Apply colors to the Theme Builder form
        $script:ThemeBuilderForm.BackColor = $script:CustomBackColor
        $script:ThemeBuilderMainFormTabControl.BackColor = $script:CustomBackColor
        $script:ThemeBuilderOutText.BackColor = $script:CustomAccentColor
        $script:ThemeBuilderClearOutTextButton.BackColor = $script:CustomDisabledColor
        $script:ThemeBuilderSaveOutTextButton.BackColor = $script:CustomForeColor
        $script:ThemeBuilderSaveOutTextButton.ForeColor = $script:CustomBackColor
        $script:ThemeBuilderSysAdminTab.BackColor = $script:CustomBackColor
        $script:ThemeBuilderServersListBox.BackColor = $script:CustomAccentColor
        $script:ThemeBuilderAppListCombo.BackColor = $script:CustomAccentColor
        $script:ThemeBuilderAppListLabel.BackColor = $script:CustomBackColor
        $script:ThemeBuilderAppListLabel.ForeColor = $script:CustomForeColor
        $script:ThemeBuilderServicesListBox.BackColor = $script:CustomAccentColor
        $script:ThemeBuilderRestartButton.BackColor = $script:CustomForeColor
        $script:ThemeBuilderRestartButton.ForeColor = $script:CustomBackColor
        $script:ThemeBuilderStartButton.BackColor = $script:CustomDisabledColor
        $script:ThemeBuilderStopButton.BackColor = $script:CustomForeColor
        $script:ThemeBuilderHelpLabel.BackColor = $script:CustomBackColor
        $script:ThemeBuilderHelpLabel.ForeColor = $script:CustomForeColor
        $script:ThemeBuilderBackColorTextBox.BackColor = $script:CustomAccentColor
        $script:ThemeBuilderForeColorTextBox.BackColor = $script:CustomAccentColor
        $script:ThemeBuilderAccentColorTextBox.BackColor = $script:CustomAccentColor
        $script:ThemeBuilderDisabledColorTextBox.BackColor = $script:CustomAccentColor
        $script:ThemeBuilderBackColorLabel.BackColor = $script:CustomBackColor
        $script:ThemeBuilderBackColorLabel.ForeColor = $script:CustomForeColor
        $script:ThemeBuilderForeColorLabel.BackColor = $script:CustomBackColor
        $script:ThemeBuilderForeColorLabel.ForeColor = $script:CustomForeColor
        $script:ThemeBuilderAccentColorLabel.BackColor = $script:CustomBackColor
        $script:ThemeBuilderAccentColorLabel.ForeColor = $script:CustomForeColor
        $script:ThemeBuilderDisabledColorLabel.BackColor = $script:CustomBackColor
        $script:ThemeBuilderDisabledColorLabel.ForeColor = $script:CustomForeColor
        $script:ThemeBuilderApplyThemeButton.BackColor = $script:CustomForeColor
        $script:ThemeBuilderApplyThemeButton.ForeColor = $script:CustomBackColor
        $script:ThemeBuilderSaveThemeButton.BackColor = $script:CustomForeColor
        $script:ThemeBuilderSaveThemeButton.ForeColor = $script:CustomBackColor
        $script:ThemeBuilderResetThemeButton.BackColor = $script:CustomForeColor
        $script:ThemeBuilderResetThemeButton.ForeColor = $script:CustomBackColor
        $script:ThemeBuilderDeleteThemesButton.BackColor = $script:CustomForeColor
        $script:ThemeBuilderDeleteThemesButton.ForeColor = $script:CustomBackColor
        $script:ThemeBuilderOutText.AppendText("$(Get-Timestamp) - Colors applied to Theme Builder form.`r`n")
        } else {
            # Create a list of invalid color inputs
            $invalidColors = @()
            if (-not $isValidBackColor) { $invalidColors += "BackColor" }
            if (-not $isValidForeColor) { $invalidColors += "ForeColor" }
            if (-not $isValidAccentColor) { $invalidColors += "AccentColor" }
            if (-not $isValidDisabledColor) { $invalidColors += "DisabledColor" }

            # Create a string from the list
            $invalidColorsString = $invalidColors -join ', '

            # Append to OutText
            $script:ThemeBuilderOutText.AppendText("$(Get-Timestamp) - Invalid color(s) in: $invalidColorsString. Please check your input and try again.`r`n")
        }
    })

    # Event handler for the Save Theme button
    $script:ThemeBuilderSaveThemeButton.add_Click({

        $isValidBackColor = Test-ColorInput -ColorInput $script:ThemeBuilderBackColorTextBox.Text -ColorType 'BackColor'
        $isValidForeColor = Test-ColorInput -ColorInput $script:ThemeBuilderForeColorTextBox.Text -ColorType 'ForeColor'
        $isValidAccentColor = Test-ColorInput -ColorInput $script:ThemeBuilderAccentColorTextBox.Text -ColorType 'AccentColor'
        $isValidDisabledColor = Test-ColorInput -ColorInput $script:ThemeBuilderDisabledColorTextBox.Text -ColorType 'DisabledColor'
        
        # Check if all colors are valid
        if ($isValidBackColor -and $isValidForeColor -and $isValidAccentColor -and $isValidDisabledColor) {

            $script:ThemeBuilderOutText.AppendText("$(Get-Timestamp) - Launching Theme Saver...`r`n")

            $script:CustomBackColor = $script:ThemeBuilderBackColorTextBox.Text
            $script:CustomForeColor = $script:ThemeBuilderForeColorTextBox.Text
            $script:CustomAccentColor = $script:ThemeBuilderAccentColorTextBox.Text
            $script:CustomDisabledColor = $script:ThemeBuilderDisabledColorTextBox.Text

            # New popup window to save the theme
            $script:ThemeBuilderSaveThemeForm = New-Object System.Windows.Forms.Form
            $script:ThemeBuilderSaveThemeForm.Text = "Save Theme"
            $script:ThemeBuilderSaveThemeForm.Size = New-Object System.Drawing.Size(300, 300)
            $script:ThemeBuilderSaveThemeForm.ShowInTaskbar = $False
            $script:ThemeBuilderSaveThemeForm.KeyPreview = $True
            $script:ThemeBuilderSaveThemeForm.FormBorderStyle = "Fixed3D"
            $script:ThemeBuilderSaveThemeForm.MaximizeBox = $False
            $script:ThemeBuilderSaveThemeForm.MinimizeBox = $False
            $script:ThemeBuilderSaveThemeForm.ControlBox = $True
            $script:ThemeBuilderSaveThemeForm.StartPosition = "CenterScreen"

            # Text box for theme name
            $script:ThemeBuilderThemeNameTextBox = New-Object System.Windows.Forms.TextBox
            $script:ThemeBuilderThemeNameTextBox.Location = New-Object System.Drawing.Point(20, 40)
            $script:ThemeBuilderThemeNameTextBox.Size = New-Object System.Drawing.Size(150, 20)

            # Label for theme name text box
            $script:ThemeBuilderThemeNameLabel = New-Object System.Windows.Forms.Label
            $script:ThemeBuilderThemeNameLabel.Location = New-Object System.Drawing.Point(20, 20)
            $script:ThemeBuilderThemeNameLabel.Size = New-Object System.Drawing.Size(150, 20)
            $script:ThemeBuilderThemeNameLabel.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Bold)
            $script:ThemeBuilderThemeNameLabel.Text = 'Theme Name'

            # Text box for quote
            $script:ThemeBuilderQuoteTextBox = New-Object System.Windows.Forms.TextBox
            $script:ThemeBuilderQuoteTextBox.Location = New-Object System.Drawing.Point(20, 120)
            $script:ThemeBuilderQuoteTextBox.Size = New-Object System.Drawing.Size(150, 20)

            # Label for quote text box
            $script:ThemeBuilderQuoteLabel = New-Object System.Windows.Forms.Label
            $script:ThemeBuilderQuoteLabel.Location = New-Object System.Drawing.Point(20, 100)
            $script:ThemeBuilderQuoteLabel.Size = New-Object System.Drawing.Size(150, 20)
            $script:ThemeBuilderQuoteLabel.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Bold)
            $script:ThemeBuilderQuoteLabel.Text = 'Quote'

            # Button for saving the theme
            $script:ThemeBuilderSaveThemePopupButton = New-Object System.Windows.Forms.Button
            $script:ThemeBuilderSaveThemePopupButton.Location = New-Object System.Drawing.Point(20, 180)
            $script:ThemeBuilderSaveThemePopupButton.Width = 100
            $script:ThemeBuilderSaveThemePopupButton.FlatStyle = "Popup"
            $script:ThemeBuilderSaveThemePopupButton.Text = "Save Theme"
            $script:ThemeBuilderSaveThemePopupButton.Font = [System.Drawing.Font]::new("Arial", 8, [System.Drawing.FontStyle]::Regular)
            $script:ThemeBuilderSaveThemePopupButton.Enabled = $false

            # Button for saving and applying the theme
            $script:ThemeBuilderSaveAndApplyThemeButton = New-Object System.Windows.Forms.Button
            $script:ThemeBuilderSaveAndApplyThemeButton.Location = New-Object System.Drawing.Point(160, 180)
            $script:ThemeBuilderSaveAndApplyThemeButton.Width = 100
            $script:ThemeBuilderSaveAndApplyThemeButton.FlatStyle = "Popup"
            $script:ThemeBuilderSaveAndApplyThemeButton.Text = "Save && Apply"
            $script:ThemeBuilderSaveAndApplyThemeButton.Font = [System.Drawing.Font]::new("Arial", 8, [System.Drawing.FontStyle]::Regular)
            $script:ThemeBuilderSaveAndApplyThemeButton.Enabled = $false

            # Apply theme colors to form if theme has been applied
            if ($global:IsThemeApplied) {
                $script:ThemeBuilderSaveThemeForm.BackColor = $script:CustomBackColor
                $script:ThemeBuilderThemeNameTextBox.BackColor = $script:CustomAccentColor
                $script:ThemeBuilderThemeNameLabel.BackColor = $script:CustomBackColor
                $script:ThemeBuilderThemeNameLabel.ForeColor = $script:CustomForeColor
                $script:ThemeBuilderQuoteTextBox.BackColor = $script:CustomAccentColor
                $script:ThemeBuilderQuoteLabel.BackColor = $script:CustomBackColor
                $script:ThemeBuilderQuoteLabel.ForeColor = $script:CustomForeColor
                $script:ThemeBuilderSaveThemePopupButton.BackColor = $script:CustomDisabledColor
                $script:ThemeBuilderSaveAndApplyThemeButton.BackColor = $script:CustomDisabledColor
            }

            # Event handler for checking text changed in Theme Name text box
            $script:ThemeBuilderThemeNameTextBox.add_TextChanged({
                Enable-SaveThemeButtons
            })

            # Event handler for checking text changed in Quote text box
            $script:ThemeBuilderQuoteTextBox.add_TextChanged({
                Enable-SaveThemeButtons
            })

            # Event handler for the Save Theme button
            $script:ThemeBuilderSaveThemePopupButton.add_Click({
                $script:NewThemeName = $script:ThemeBuilderThemeNameTextBox.Text
                $script:NewThemeQuote = $script:ThemeBuilderQuoteTextBox.Text

                # Call the Save-CustomTheme function and check if the save was successful
                $saveSuccessful = Save-CustomTheme -ThemeName $script:NewThemeName -BackColor $script:CustomBackColor -ForeColor $script:CustomForeColor -AccentColor $script:CustomAccentColor -DisabledColor $script:CustomDisabledColor -Quote $script:NewThemeQuote -ColorTheme ([ref]$ColorTheme)

                if ($saveSuccessful) {
                    # Create a new menu item for the new theme
                    $NewMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
                    $NewMenuItem.Text = $script:NewThemeName

                    # Add the new menu item to the CustomThemes dropdown
                    $script:CustomThemes.DropDownItems.Add($NewMenuItem) | Out-Null

                    # Sort the dropdown items alphabetically
                    $SortedItems = $script:CustomThemes.DropDownItems | Sort-Object { $_.Text -replace '• ', '' }
                    $script:CustomThemes.DropDownItems.Clear()
                    $script:CustomThemes.DropDownItems.AddRange($SortedItems)

                    #Close the Save Theme popup form
                    $script:ThemeBuilderSaveThemeForm.Close()
                }
            })

            # Event handler for the Save & Apply button
            $script:ThemeBuilderSaveAndApplyThemeButton.add_Click({
                $script:NewThemeName = $script:ThemeBuilderThemeNameTextBox.Text
                $script:NewThemeQuote = $script:ThemeBuilderQuoteTextBox.Text

                # Call the Save-CustomTheme function and check if the save was successful
                $saveSuccessful = Save-CustomTheme -ThemeName $script:NewThemeName -BackColor $script:CustomBackColor -ForeColor $script:CustomForeColor -AccentColor $script:CustomAccentColor -DisabledColor $script:CustomDisabledColor -Quote $script:NewThemeQuote -ColorTheme ([ref]$ColorTheme)

                if ($saveSuccessful) {
                    # Create a new menu item for the new theme
                    $NewMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
                    $NewMenuItem.Text = $script:NewThemeName
                    $NewMenuItem.add_Click({
                        # Re-read the JSON file to update $ColorTheme
                        $ColorTheme = Get-Content -Path .\ColorThemes.json | ConvertFrom-Json

                        Update-MainTheme -Team $this.Text -Category 'Custom' -ColorData $ColorTheme
                    })

                    # Add the new menu item to the CustomThemes dropdown
                    $script:CustomThemes.DropDownItems.Add($NewMenuItem) | Out-Null

                    # Sort the dropdown items alphabetically
                    $SortedItems = $script:CustomThemes.DropDownItems | Sort-Object Text
                    $script:CustomThemes.DropDownItems.Clear()
                    $script:CustomThemes.DropDownItems.AddRange($SortedItems)

                    # Apply the new theme
                    Update-MainTheme -Team $script:NewThemeName -Category 'Custom' -ColorData $ColorTheme
                    $script:ThemeBuilderSaveThemeForm.Close()
                }
            })

            # Add controls to Save Theme popup form
            $script:ThemeBuilderSaveThemeForm.Controls.Add($script:ThemeBuilderThemeNameTextBox)
            $script:ThemeBuilderSaveThemeForm.Controls.Add($script:ThemeBuilderThemeNameLabel)
            $script:ThemeBuilderSaveThemeForm.Controls.Add($script:ThemeBuilderQuoteTextBox)
            $script:ThemeBuilderSaveThemeForm.Controls.Add($script:ThemeBuilderQuoteLabel)
            $script:ThemeBuilderSaveThemeForm.Controls.Add($script:ThemeBuilderSaveThemePopupButton)
            $script:ThemeBuilderSaveThemeForm.Controls.Add($script:ThemeBuilderSaveAndApplyThemeButton)

            # Open the form
            $script:ThemeBuilderSaveThemeForm.ShowDialog() | Out-Null
        } else {
            # Create a list of invalid color inputs
            $invalidColors = @()
            if (-not $isValidBackColor) { $invalidColors += "BackColor" }
            if (-not $isValidForeColor) { $invalidColors += "ForeColor" }
            if (-not $isValidAccentColor) { $invalidColors += "AccentColor" }
            if (-not $isValidDisabledColor) { $invalidColors += "DisabledColor" }

            # Create a string from the list
            $invalidColorsString = $invalidColors -join ', '

            # Append to OutText
            $script:ThemeBuilderOutText.AppendText("$(Get-Timestamp) - Invalid color(s) in: $invalidColorsString. Please check your input and try again.`r`n")
        }
    })

    # Event handler for the Reset Theme button
    $script:ThemeBuilderResetThemeButton.add_Click({

        # Set global Is Theme Applied variable to true
        $global:IsThemeApplied = $false

        $script:ThemeBuilderBackColorTextBox.Text = ''
        $script:ThemeBuilderForeColorTextBox.Text = ''
        $script:ThemeBuilderAccentColorTextBox.Text = ''
        $script:ThemeBuilderDisabledColorTextBox.Text = ''
        $script:ThemeBuilderApplyThemeButton.Enabled = $false
        $script:ThemeBuilderSaveThemeButton.Enabled = $false
        $script:ThemeBuilderResetThemeButton.Enabled = $false
        $script:ThemeBuilderForm.BackColor = [System.Drawing.SystemColors]::Control
        $script:ThemeBuilderMainFormTabControl.BackColor = [System.Drawing.SystemColors]::Control
        $script:ThemeBuilderOutText.BackColor = [System.Drawing.SystemColors]::Control
        $script:ThemeBuilderClearOutTextButton.BackColor = [System.Drawing.SystemColors]::Control
        $script:ThemeBuilderSaveOutTextButton.BackColor = [System.Drawing.SystemColors]::Control
        $script:ThemeBuilderSaveOutTextButton.ForeColor = [System.Drawing.SystemColors]::ControlText
        $script:ThemeBuilderSysAdminTab.BackColor = [System.Drawing.SystemColors]::Control
        $script:ThemeBuilderServersListBox.BackColor = [System.Drawing.SystemColors]::Control
        $script:ThemeBuilderAppListCombo.BackColor = [System.Drawing.SystemColors]::Control
        $script:ThemeBuilderAppListLabel.BackColor = [System.Drawing.SystemColors]::Control
        $script:ThemeBuilderAppListLabel.ForeColor = [System.Drawing.SystemColors]::ControlText
        $script:ThemeBuilderServicesListBox.BackColor = [System.Drawing.SystemColors]::Control
        $script:ThemeBuilderRestartButton.BackColor = [System.Drawing.SystemColors]::Control
        $script:ThemeBuilderRestartButton.ForeColor = [System.Drawing.SystemColors]::ControlText
        $script:ThemeBuilderStartButton.BackColor = [System.Drawing.SystemColors]::Control
        $script:ThemeBuilderStopButton.BackColor = [System.Drawing.SystemColors]::Control
        $script:ThemeBuilderHelpLabel.BackColor = [System.Drawing.SystemColors]::Control
        $script:ThemeBuilderHelpLabel.ForeColor = [System.Drawing.SystemColors]::ControlText
        $script:ThemeBuilderBackColorTextBox.BackColor = [System.Drawing.SystemColors]::Control
        $script:ThemeBuilderForeColorTextBox.BackColor = [System.Drawing.SystemColors]::Control
        $script:ThemeBuilderAccentColorTextBox.BackColor = [System.Drawing.SystemColors]::Control
        $script:ThemeBuilderDisabledColorTextBox.BackColor = [System.Drawing.SystemColors]::Control
        $script:ThemeBuilderBackColorLabel.BackColor = [System.Drawing.SystemColors]::Control
        $script:ThemeBuilderBackColorLabel.ForeColor = [System.Drawing.SystemColors]::ControlText
        $script:ThemeBuilderForeColorLabel.BackColor = [System.Drawing.SystemColors]::Control
        $script:ThemeBuilderForeColorLabel.ForeColor = [System.Drawing.SystemColors]::ControlText
        $script:ThemeBuilderAccentColorLabel.BackColor = [System.Drawing.SystemColors]::Control
        $script:ThemeBuilderAccentColorLabel.ForeColor = [System.Drawing.SystemColors]::ControlText
        $script:ThemeBuilderDisabledColorLabel.BackColor = [System.Drawing.SystemColors]::Control
        $script:ThemeBuilderDisabledColorLabel.ForeColor = [System.Drawing.SystemColors]::ControlText
        $script:ThemeBuilderApplyThemeButton.BackColor = [System.Drawing.SystemColors]::Control
        $script:ThemeBuilderApplyThemeButton.ForeColor = [System.Drawing.SystemColors]::ControlText
        $script:ThemeBuilderSaveThemeButton.BackColor = [System.Drawing.SystemColors]::Control
        $script:ThemeBuilderSaveThemeButton.ForeColor = [System.Drawing.SystemColors]::ControlText
        $script:ThemeBuilderResetThemeButton.BackColor = [System.Drawing.SystemColors]::Control
        $script:ThemeBuilderResetThemeButton.ForeColor = [System.Drawing.SystemColors]::ControlText
        $script:ThemeBuilderDeleteThemesButton.BackColor = [System.Drawing.SystemColors]::Control
        $script:ThemeBuilderDeleteThemesButton.ForeColor = [System.Drawing.SystemColors]::ControlText
    })

    # Event handler for the Delete Custom Themes button
    $script:ThemeBuilderDeleteThemesButton.add_Click({
        # New popup window to delete custom themes
        $script:ThemeBuilderDeleteThemesForm = New-Object System.Windows.Forms.Form
        $script:ThemeBuilderDeleteThemesForm.Text = "Delete Custom Themes"
        $script:ThemeBuilderDeleteThemesForm.Size = New-Object System.Drawing.Size(300, 380)
        $script:ThemeBuilderDeleteThemesForm.ShowInTaskbar = $False
        $script:ThemeBuilderDeleteThemesForm.KeyPreview = $True
        $script:ThemeBuilderDeleteThemesForm.FormBorderStyle = "Fixed3D"
        $script:ThemeBuilderDeleteThemesForm.MaximizeBox = $False
        $script:ThemeBuilderDeleteThemesForm.MinimizeBox = $False
        $script:ThemeBuilderDeleteThemesForm.ControlBox = $True
        $script:ThemeBuilderDeleteThemesForm.StartPosition = "CenterScreen"

        # Label for theme name list box
        $script:ThemeBuilderThemeNameListBoxLabel = New-Object System.Windows.Forms.Label
        $script:ThemeBuilderThemeNameListBoxLabel.Location = New-Object System.Drawing.Point(65, 20)
        $script:ThemeBuilderThemeNameListBoxLabel.Size = New-Object System.Drawing.Size(150, 40)
        $script:ThemeBuilderThemeNameListBoxLabel.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Bold)
        $script:ThemeBuilderThemeNameListBoxLabel.Text = 'Permanently delete one or more custom themes'
        $script:ThemeBuilderThemeNameListBoxLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter

        # List box for theme names
        $script:ThemeBuilderThemeNameListBox = New-Object System.Windows.Forms.ListBox
        $script:ThemeBuilderThemeNameListBox.Location = New-Object System.Drawing.Point(65, 70)
        $script:ThemeBuilderThemeNameListBox.Size = New-Object System.Drawing.Size(150, 200)
        $script:ThemeBuilderThemeNameListBox.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
        $script:ThemeBuilderThemeNameListBox.SelectionMode = "MultiExtended"

        # Button for deleting themes
        $script:ThemeBuilderDeleteThemesPopupButton = New-Object System.Windows.Forms.Button
        $script:ThemeBuilderDeleteThemesPopupButton.Location = New-Object System.Drawing.Point(90, 290)
        $script:ThemeBuilderDeleteThemesPopupButton.Width = 100
        $script:ThemeBuilderDeleteThemesPopupButton.FlatStyle = "Popup"
        $script:ThemeBuilderDeleteThemesPopupButton.Text = "Delete Theme"
        $script:ThemeBuilderDeleteThemesPopupButton.Font = [System.Drawing.Font]::new("Arial", 8, [System.Drawing.FontStyle]::Regular)
        $script:ThemeBuilderDeleteThemesPopupButton.Enabled = $false

        # Apply theme colors to form if theme has been applied
        if ($global:IsThemeApplied) {
            $script:ThemeBuilderDeleteThemesForm.BackColor = $script:CustomBackColor
            $script:ThemeBuilderThemeNameListBoxLabel.BackColor = $script:CustomBackColor
            $script:ThemeBuilderThemeNameListBoxLabel.ForeColor = $script:CustomForeColor
            $script:ThemeBuilderThemeNameListBox.BackColor = $script:CustomAccentColor
            $script:ThemeBuilderDeleteThemesPopupButton.BackColor = $script:CustomDisabledColor
        }

        # Assuming $ColorTheme is the parsed JSON object
        $script:CustomThemes = $ColorTheme.Custom | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name

        foreach ($themeName in $script:CustomThemes) {
            $script:ThemeBuilderThemeNameListBox.Items.Add($themeName)
        }

        # Logic for enabling the Delete Theme(s) button
        $script:ThemeBuilderThemeNameListBox.add_SelectedIndexChanged({
            if ($script:ThemeBuilderThemeNameListBox.SelectedItems.Count -gt 0) {
                $script:ThemeBuilderDeleteThemesPopupButton.Enabled = $true
                if ($global:IsThemeApplied) {
                    $script:ThemeBuilderDeleteThemesPopupButton.BackColor = $script:CustomForeColor
                    $script:ThemeBuilderDeleteThemesPopupButton.ForeColor = $script:CustomBackColor
                }
            }
            else {
                $script:ThemeBuilderDeleteThemesPopupButton.Enabled = $false
                if ($global:IsThemeApplied) {
                    $script:ThemeBuilderDeleteThemesPopupButton.BackColor = $script:CustomDisabledColor
                }
            }
        })

        # Event handler for the Delete Theme(s) button
        $script:ThemeBuilderDeleteThemesPopupButton.add_Click({
            $script:ThemesToDelete = @($script:ThemeBuilderThemeNameListBox.SelectedItems | ForEach-Object { $_.ToString() })

            # Ensure $ColorTheme.Custom is not null
            if ($null -ne $ColorTheme.Custom) {
                foreach ($ThemeToDelete in $script:ThemesToDelete) {
                    $ColorTheme.Custom.PSObject.Properties.Remove($ThemeToDelete)
                }
            }

            # Append text to output
            $script:ThemeBuilderOutText.AppendText("$(Get-Timestamp) - Successfully deleted the below custom theme(s):`r`n")
            $script:ThemeBuilderOutText.AppendText("$(Get-Timestamp) - $($script:ThemesToDelete -join "`r`n")`r`n")

            # Save the changes back to the JSON file
            $jsonString = $ColorTheme | ConvertTo-Json -Depth 5
            Set-Content -Path .\ColorThemes.json -Value $jsonString

            # Re-load the $ColorTheme from the file or refresh it here
            $ColorTheme = Get-Content -Path .\ColorThemes.json | ConvertFrom-Json

            # Check if $script:CustomThemes and its DropDownItems are not null
            if ($null -ne $script:CustomThemes -and $null -ne $script:CustomThemes.DropDownItems) {
                $script:CustomThemes.DropDownItems.Clear()
                foreach ($Theme in $ColorTheme.Custom.PSObject.Properties) {
                    $NewMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
                    $NewMenuItem.Text = $Theme.Name
                    $NewMenuItem.add_Click({
                        Update-MainTheme -Team $this.Text -Category 'Custom' -ColorData $ColorTheme
                    })
                    $script:CustomThemes.DropDownItems.Add($NewMenuItem) | Out-Null
                }

                # Sort and re-add the dropdown items
                $SortedItems = $script:CustomThemes.DropDownItems | Sort-Object Text
                $script:CustomThemes.DropDownItems.Clear()
                $script:CustomThemes.DropDownItems.AddRange($SortedItems)
            }

            # Update the Delete Custom Themes button
            if ($ColorTheme.Custom.PSObject.Properties.Count -gt 0) {
                $script:ThemeBuilderDeleteThemesButton.Enabled = $true
            } else {
                $script:ThemeBuilderDeleteThemesButton.Enabled = $false
            }

            # Clear and repopulate the list box
            $script:ThemeBuilderThemeNameListBox.Items.Clear()
            foreach ($Theme in $ColorTheme.Custom.PSObject.Properties) {
                $script:ThemeBuilderThemeNameListBox.Items.Add($Theme.Name)
            }
        })

        # Add controls to the Delete Custom Themes form
        $script:ThemeBuilderDeleteThemesForm.Controls.Add($script:ThemeBuilderThemeNameListBoxLabel)
        $script:ThemeBuilderDeleteThemesForm.Controls.Add($script:ThemeBuilderThemeNameListBox)
        $script:ThemeBuilderDeleteThemesForm.Controls.Add($script:ThemeBuilderDeleteThemesPopupButton)

        # Show form
        $script:ThemeBuilderDeleteThemesForm.ShowDialog() | Out-Null
    })

    # Add controls to the Theme Builder form
    $script:ThemeBuilderForm.Controls.Add($script:ThemeBuilderMainFormTabControl)
    $script:ThemeBuilderMainFormTabControl.Controls.Add($script:ThemeBuilderSysAdminTab)
    $script:ThemeBuilderSysAdminTab.Controls.Add($ThemeBuilderRestartsTabControl)
    $script:ThemeBuilderSysAdminTab.Controls.Add($script:ThemeBuilderServersListBox)
    $script:ThemeBuilderSysAdminTab.Controls.Add($script:ThemeBuilderAppListCombo)
    $script:ThemeBuilderSysAdminTab.Controls.Add($script:ThemeBuilderAppListLabel)
    $ThemeBuilderRestartsTabControl.Controls.Add($ThemeBuilderServicesTab)
    $ThemeBuilderServicesTab.Controls.Add($script:ThemeBuilderServicesListBox)
    $script:ThemeBuilderSysAdminTab.Controls.Add($script:ThemeBuilderRestartButton)
    $script:ThemeBuilderSysAdminTab.Controls.Add($script:ThemeBuilderStartButton)
    $script:ThemeBuilderSysAdminTab.Controls.Add($script:ThemeBuilderStopButton)
    $script:ThemeBuilderForm.Controls.Add($script:ThemeBuilderHelpLabel)
    $script:ThemeBuilderForm.Controls.Add($script:ThemeBuilderBackColorTextBox)
    $script:ThemeBuilderForm.Controls.Add($script:ThemeBuilderBackColorLabel)
    $script:ThemeBuilderForm.Controls.Add($script:ThemeBuilderForeColorTextBox)
    $script:ThemeBuilderForm.Controls.Add($script:ThemeBuilderForeColorLabel)
    $script:ThemeBuilderForm.Controls.Add($script:ThemeBuilderAccentColorTextBox)
    $script:ThemeBuilderForm.Controls.Add($script:ThemeBuilderAccentColorLabel)
    $script:ThemeBuilderForm.Controls.Add($script:ThemeBuilderDisabledColorTextBox)
    $script:ThemeBuilderForm.Controls.Add($script:ThemeBuilderDisabledColorLabel)
    $script:ThemeBuilderForm.Controls.Add($script:ThemeBuilderApplyThemeButton)
    $script:ThemeBuilderForm.Controls.Add($script:ThemeBuilderSaveThemeButton)
    $script:ThemeBuilderForm.Controls.Add($script:ThemeBuilderResetThemeButton)
    $script:ThemeBuilderForm.Controls.Add($script:ThemeBuilderDeleteThemesButton)
    $script:ThemeBuilderForm.Controls.Add($script:ThemeBuilderOutText)
    $script:ThemeBuilderForm.Controls.Add($script:ThemeBuilderClearOutTextButton)
    $script:ThemeBuilderForm.Controls.Add($script:ThemeBuilderSaveOutTextButton)

    # Show the Theme Builder form
    $script:ThemeBuilderForm.Show() | Out-Null

    # Click event for form close; sets global variable to false and clears color variables
    $script:ThemeBuilderForm.Add_FormClosed({
        $global:IsThemeBuilderPopupActive = $false
        $script:CustomBackColor = $null
        $script:CustomForeColor = $null
        $script:CustomAccentColor = $null
        $script:CustomDisabledColor = $null
        $global:IsThemeApplied = $false
    })
})

# About menu
$AboutMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$AboutMenu.Text = "About"
$MenuAboutItem = New-Object System.Windows.Forms.ToolStripMenuItem
$MenuAboutItem.Text = "About This App"
$MenuGitHub = New-Object System.Windows.Forms.ToolStripMenuItem
$MenuGitHub.Text = "GitHub Repo"

# Click event for the Options menu Show Tool Tips option
$ShowToolTipsMenu.add_Click({
    if ($ShowToolTipsMenu.Text -eq "Show Tool Tips") {
        $ConfigValues.HoverToolTips = "Enabled"
        $UpdatedToolTipValue = ConvertTo-Json -InputObject $ConfigValues -Depth 100
        Set-Content -Path .\Config.json -Value $UpdatedToolTipValue
        Enable-ToolTips
        $ShowToolTipsMenu.Text = "Hide Tool Tips"
        $OutText.AppendText("$(Get-Timestamp) - Tool tips have been enabled.`r`n")
    }
    else {
        $ConfigValues.HoverToolTips = "Disabled"
        $UpdatedToolTipValue = ConvertTo-Json -InputObject $ConfigValues -Depth 100
        Set-Content -Path .\Config.json -Value $UpdatedToolTipValue
        $ToolTip.RemoveAll()
        $ShowToolTipsMenu.Text = "Show Tool Tips"
        $OutText.AppendText("$(Get-Timestamp) - Tool tips have been disabled.`r`n")
    }
})

# Click event for the GitHub Repo menu option
$MenuGitHub.add_Click({ Start-Process "https://github.com/jthamind/DesktopAssistant" })

# Click event for the About menu option
$MenuAboutItem.add_Click({
    $OutText.AppendText("$(Get-Timestamp) - Launching About form...`r`n")
    # About form
    $script:AboutForm = New-Object System.Windows.Forms.Form
    $script:AboutForm.Text = "About"
    $script:AboutForm.Size = New-Object System.Drawing.Size(400, 300)
    $script:AboutForm.ShowInTaskbar = $False
    $script:AboutForm.KeyPreview = $True
    $script:AboutForm.AutoSize = $True
    $script:AboutForm.FormBorderStyle = "Fixed3D"
    $script:AboutForm.Icon = $Icon
    $script:AboutForm.MaximizeBox = $False
    $script:AboutForm.MinimizeBox = $False
    $global:IsAboutPopupActive = $true

    # Label for About form
    $AboutLabel = New-Object System.Windows.Forms.Label
    $AboutLabel.Location = New-Object System.Drawing.Size(115, 120)
    $AboutLabel.Size = New-Object System.Drawing.Size(150, 20)
    $AboutLabel.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Bold)
    $AboutLabel.Text = 'ETG Desktop Assistant'

    # Picture box for Allied Solutions logo
    $AboutMenuAlliedLogo = New-Object System.Windows.Forms.PictureBox
    $AboutMenuAlliedLogo.Location = New-Object System.Drawing.Size(50, 5)
    $AboutMenuAlliedLogo.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
    $AboutMenuAlliedLogo.Size = New-Object System.Drawing.Size(275, 100)
    $AboutMenuAlliedLogoImage = [System.Drawing.Image]::FromFile($AlliedLogo)
    $AboutMenuAlliedLogo.Image = $AboutMenuAlliedLogoImage

    # Label for version numbers
    $VersionLabel = New-Object System.Windows.Forms.Label
    $VersionLabel.Location = New-Object System.Drawing.Size(145, 145)
    $VersionLabel.Size = New-Object System.Drawing.Size(150, 20)
    $VersionLabel.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
    $VersionLabel.Text = "Version: 2.0.1"

    # Label for upcoming features
    $UpcomingFeaturesLabel = New-Object System.Windows.Forms.Label
    $UpcomingFeaturesLabel.Location = New-Object System.Drawing.Size(125, 180)
    $UpcomingFeaturesLabel.Size = New-Object System.Drawing.Size(200, 75)
    $UpcomingFeaturesLabel.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Bold)
    $UpcomingFeaturesLabel.Text = "Upcoming Features`r`n• Remote PowerShell Sessions`r`n• Billing Service Restarts"

    # Label for Allied IMPACT
    $AlliedImpactLabel = New-Object System.Windows.Forms.Label
    $AlliedImpactLabel.Location = New-Object System.Drawing.Size(40, 260)
    $AlliedImpactLabel.Size = New-Object System.Drawing.Size(300, 20)
    $AlliedImpactLabel.Width = 300
    $AlliedImpactLabel.Font = [System.Drawing.Font]::new("Arial", 7, [System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Italic)
    $AlliedImpactLabel.Text = "Made with Passion to make an IMPACT at Allied Solutions, LLC."

    # Check if DefaultUserTheme has a value or is null
    if ($null -ne $ConfigValues.DefaultUserTheme -and $ConfigValues.DefaultUserTheme -ne "") {
        # Get theme
        $SelectedTheme = $ColorTheme.PSObject.Properties | Where-Object { $_.Value.$($ConfigValues.DefaultUserTheme) } | Select-Object -ExpandProperty Name
        $themeColors = $ColorTheme.$SelectedTheme.$($ConfigValues.DefaultUserTheme)
        
		$script:AboutForm.BackColor = $themeColors.BackColor
		$script:AboutForm.ForeColor = $themeColors.ForeColor
        $AboutLabel.BackColor = $themeColors.BackColor
        $AboutLabel.ForeColor = $themeColors.ForeColor
    }

    # Click event for form close; sets global variable to false
    $script:AboutForm.Add_FormClosed({
        $global:IsAboutPopupActive = $false
    })

    $script:AboutForm.Controls.Add($AboutLabel)
    $script:AboutForm.Controls.Add($AboutMenuAlliedLogo)
    $script:AboutForm.Controls.Add($VersionLabel)
    $script:AboutForm.Controls.Add($UpcomingFeaturesLabel)
    $script:AboutForm.Controls.Add($AlliedImpactLabel)
    $script:AboutForm.Show()
})

<#
? **********************************************************************************************************************
? END OF MENU STRIP
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
$RestartsTabControl.AutoSize = $False

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
$AppListCombo.SelectedIndex = -1

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
$script:RestartButton = New-Object System.Windows.Forms.Button
$script:RestartButton.Location = New-Object System.Drawing.Point(490, 95)
$script:RestartButton.Width = 75
$script:RestartButton.BackColor = $global:DisabledBackColor
$script:RestartButton.ForeColor = $global:DisabledForeColor
$script:RestartButton.FlatStyle = "Popup"
$script:RestartButton.Text = "Restart"
$script:RestartButton.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
$script:RestartButton.Enabled = $false

# Button for starting services
$script:StartButton = New-Object System.Windows.Forms.Button
$script:StartButton.Location = New-Object System.Drawing.Point(490, 125)
$script:StartButton.Width = 75
$script:StartButton.BackColor = $global:DisabledBackColor
$script:StartButton.ForeColor = $global:DisabledForeColor
$script:StartButton.FlatStyle = "Popup"
$script:StartButton.Text = "Start"
$script:StartButton.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
$script:StartButton.Enabled = $false

# Button for stopping services
$script:StopButton = New-Object System.Windows.Forms.Button
$script:StopButton.Location = New-Object System.Drawing.Point(490, 155)
$script:StopButton.Width = 75
$script:StopButton.BackColor = $global:DisabledBackColor
$script:StopButton.ForeColor = $global:DisabledForeColor
$script:StopButton.FlatStyle = "Popup"
$script:StopButton.Text = "Stop"
$script:StopButton.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
$script:StopButton.Enabled = $false

# Button for for opening IIS Site in Windows Explorer
$script:OpenSiteButton = New-Object System.Windows.Forms.Button
$script:OpenSiteButton.Location = New-Object System.Drawing.Point(490, 185)
$script:OpenSiteButton.Width = 75
$script:OpenSiteButton.BackColor = $global:DisabledBackColor
$script:OpenSiteButton.ForeColor = $global:DisabledForeColor
$script:OpenSiteButton.FlatStyle = "Popup"
$script:OpenSiteButton.Text = "Open"
$script:OpenSiteButton.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
$script:OpenSiteButton.Enabled = $false
$script:OpenSiteButton.Visible = $false

# Button for restarting IIS on server
$script:RestartIISButton = New-Object System.Windows.Forms.Button
$script:RestartIISButton.Location = New-Object System.Drawing.Point(490, 215)
$script:RestartIISButton.Width = 75
$script:RestartIISButton.BackColor = $global:DisabledBackColor
$script:RestartIISButton.ForeColor = $global:DisabledForeColor
$script:RestartIISButton.FlatStyle = "Popup"
$script:RestartIISButton.Text = "Restart IIS"
$script:RestartIISButton.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
$script:RestartIISButton.Enabled = $false
$script:RestartIISButton.Visible = $false

# Button for starting IIS on server
$script:StartIISButton = New-Object System.Windows.Forms.Button
$script:StartIISButton.Location = New-Object System.Drawing.Point(490, 245)
$script:StartIISButton.Width = 75
$script:StartIISButton.BackColor = $global:DisabledBackColor
$script:StartIISButton.ForeColor = $global:DisabledForeColor
$script:StartIISButton.FlatStyle = "Popup"
$script:StartIISButton.Text = "Start IIS"
$script:StartIISButton.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
$script:StartIISButton.Enabled = $false
$script:StartIISButton.Visible = $false

# Button for stopping IIS on server
$script:StopIISButton = New-Object System.Windows.Forms.Button
$script:StopIISButton.Location = New-Object System.Drawing.Point(490, 275)
$script:StopIISButton.Width = 75
$script:StopIISButton.BackColor = $global:DisabledBackColor
$script:StopIISButton.ForeColor = $global:DisabledForeColor
$script:StopIISButton.FlatStyle = "Popup"
$script:StopIISButton.Text = "Stop IIS"
$script:StopIISButton.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
$script:StopIISButton.Enabled = $false
$script:StopIISButton.Visible = $false

# Separator line under Restarts GUI
$RestartsSeparator = New-Object System.Windows.Forms.Label
$RestartsSeparator.Location = New-Object System.Drawing.Size(0, 350)
$RestartsSeparator.Size = New-Object System.Drawing.Size(1000, 2)

# Disable restart button if no options are selected in the ServicesListBox. Enable if options are selected.
$ServicesListBox.add_SelectedIndexChanged({
    if ($ServicesListBox.SelectedItems.Count -gt 0) {
        $backColor = Get-AppropriateColor -ColorType "BackColor"
        $foreColor = Get-AppropriateColor -ColorType "ForeColor"
        $script:RestartButton.Enabled = $true
        $script:StartButton.Enabled = $true
        $script:StopButton.Enabled = $true
        $script:RestartButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
        $script:RestartButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
        $script:StartButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
        $script:StartButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
        $script:StopButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
        $script:StopButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
    } else {
        $script:RestartButton.Enabled = $false
        $script:StartButton.Enabled = $false
        $script:StopButton.Enabled = $false
        $script:RestartButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
        $script:RestartButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
        $script:StartButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
        $script:StartButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
        $script:StopButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
        $script:StopButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
    }
})

# Disable restart button if no options are selected in the IISSitesListBox. Enable if options are selected.
$IISSitesListBox.add_SelectedIndexChanged({
    if ($IISSitesListBox.SelectedItems.Count -gt 0) {
        $backColor = Get-AppropriateColor -ColorType "BackColor"
        $foreColor = Get-AppropriateColor -ColorType "ForeColor"
        $script:RestartButton.Enabled = $true
        $script:StartButton.Enabled = $true
        $script:StopButton.Enabled = $true
        $script:OpenSiteButton.Enabled = $true
        $script:RestartButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
        $script:RestartButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
        $script:StartButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
        $script:StartButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
        $script:StopButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
        $script:StopButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
        $script:OpenSiteButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
        $script:OpenSiteButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
    } else {
        $script:RestartButton.Enabled = $false
        $script:StartButton.Enabled = $false
        $script:StopButton.Enabled = $false
        $script:OpenSiteButton.Enabled = $false
        $script:RestartButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
        $script:RestartButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
        $script:StartButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
        $script:StartButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
        $script:StopButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
        $script:StopButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
        $script:OpenSiteButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
        $script:OpenSiteButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
    }
})

# Disable restart button if no options are selected in the AppPoolsListBox. Enable if options are selected.
$AppPoolsListBox.add_SelectedIndexChanged({
    if ($AppPoolsListBox.SelectedItems.Count -gt 0) {
        $backColor = Get-AppropriateColor -ColorType "BackColor"
        $foreColor = Get-AppropriateColor -ColorType "ForeColor"
        $script:RestartButton.Enabled = $true
        $script:StartButton.Enabled = $true
        $script:StopButton.Enabled = $true
        $script:RestartButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
        $script:RestartButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
        $script:StartButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
        $script:StartButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
        $script:StopButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
        $script:StopButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
    } else {
        $script:RestartButton.Enabled = $false
        $script:StartButton.Enabled = $false
        $script:StopButton.Enabled = $false
        $script:RestartButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
        $script:RestartButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
        $script:StartButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
        $script:StartButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
        $script:StopButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
        $script:StopButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
    }
})

# Set Open button to invisible if the selected tab is not IIS Sites
$RestartsTabControl.add_SelectedIndexChanged({
    $SelectedTab = $RestartsTabControl.SelectedTab.Text
    $synchash.SelectedTab = $SelectedTab
    if ($SelectedTab -eq "IIS Sites") {
        $script:OpenSiteButton.Visible = $true
    } else {
        $script:OpenSiteButton.Visible = $false
    }
})

# Variables for importing server list from CSV
$csvPath = Join-Path -Path $PSScriptRoot -ChildPath $ConfigValues.csvPath
$ServerCSV = Import-CSV $csvPath
$csvHeaders = ($ServerCSV | Get-Member -MemberType NoteProperty).name

# Populate the AppListCombo with the list of applications from the CSV
foreach ($header in $csvHeaders) {
    [void]$AppListCombo.Items.Add($header)
}

# Event handler for selecting a server from ServersListBox
$ServersListBox.add_SelectedIndexChanged({
    $ServicesListBox.ClearSelected()
    $IISSitesListBox.ClearSelected()
    $AppPoolsListBox.ClearSelected()
    $backColor = Get-AppropriateColor -ColorType "BackColor"
    $foreColor = Get-AppropriateColor -ColorType "ForeColor"
    if ($ServersListBox.SelectedIndex -ge 0) {
        $script:StartIISButton.Visible = $true
        $script:StopIISButton.Visible = $true
        $script:RestartIISButton.Visible = $true
        $script:StartIISButton.Enabled = $true
        $script:StopIISButton.Enabled = $true
        $script:RestartIISButton.Enabled = $true
        $script:StartIISButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
        $script:StartIISButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
        $script:StopIISButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
        $script:StopIISButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
        $script:RestartIISButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
        $script:RestartIISButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
    } else {
        $script:StartIISButton.Visible = $false
        $script:StartIISButton.Enabled = $false
        $script:StopIISButton.Visible = $false
        $script:StopIISButton.Enabled = $false
        $script:RestartIISButton.Visible = $false
        $script:RestartIISButton.Enabled = $false
        $script:StartIISButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
        $script:StartIISButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
        $script:StopIISButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
        $script:StopIISButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
        $script:RestartIISButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
        $script:RestartIISButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
    }
})

# Event handler for the AppListCombo's SelectedIndexChanged event
$AppListCombo_SelectedIndexChanged = {
    $selectedHeader = $AppListCombo.SelectedItem.ToString()
    $ServersListBox.SelectedIndex = -1
    $ServersListBox.Items.Clear()
    $ServicesListBox.Items.Clear()
    $IISSitesListBox.Items.Clear()
    $AppPoolsListBox.Items.Clear()
    $OutText.AppendText("$(Get-Timestamp) - Selected server list: $selectedHeader`r`n")
    $servers = $ServerCSV | ForEach-Object { $_.$selectedHeader } | Where-Object { $_ -ne '' }
    $servers | ForEach-Object {
        [void]$ServersListBox.Items.Add($_)
    }
}

# Button click event handler for restarting one or more services, IIS sites, or app pools in an async runspace pool
$script:RestartButton.Add_Click({
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
$script:StartButton.Add_Click({
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
$script:StopButton.Add_Click({
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
$script:OpenSiteButton.Add_Click({
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
$script:RestartIISButton.Add_Click({
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
$script:StartIISButton.Add_Click({
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
$script:StopIISButton.Add_Click({
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
$NSLookupButton.Text = "Get IP"
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
$ReverseIPButton.Text = "Get DNS"
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

# Button click event handler for Reverse IP Lookup
$ReverseIPButton.add_Click({
    $ip = $ReverseIPTextBox.Text
    $ReverseIPTextBox.Text = ''
    Invoke-ReverseIPLookupRunspace -ip $ip -OutText $OutText -TimestampFunction ${function:Get-Timestamp}
})

# Event handler for pressing Enter in the Reverse IP Lookup text box
$ReverseIPTextBox.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") {
        $ip = $ReverseIPTextBox.Text
        $ReverseIPTextBox.Text = ''
        Invoke-ReverseIPLookupRunspace -ip $ip -OutText $OutText -TimestampFunction ${function:Get-Timestamp}
        $_.SuppressKeyPress = $true
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
$RemoteConfigs = $ConfigValues.RemoteConfigs
$LocalSupportTool = Join-Path -Path $PSScriptRoot -ChildPath $ConfigValues.LocalSupportTool
# Resolve the relative path to an absolute path
if (Test-Path "$LocalSupportTool") {
    $ResolvedLocalSupportTool = Resolve-Path $LocalSupportTool
}
$ResolvedLocalSupportTool = Resolve-Path $LocalSupportTool
$LocalConfigs = Join-Path -Path $PSScriptRoot -ChildPath $ConfigValues.LocalConfigs
# Resolve the relative path to an absolute path
if (Test-Path "$LocalConfigs") {
    $ResolvedLocalConfigs = Resolve-Path $LocalConfigs
}

# Combobox for environment selection
$PSTCombo = New-Object System.Windows.Forms.ComboBox
$PSTCombo.Location = New-Object System.Drawing.Point(5,65)
$PSTCombo.Size = New-Object System.Drawing.Size(150, 200)
$PSTCombo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$PSTCombo.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
@('QA', 'Stage', 'Production') | ForEach-Object { [void]$PSTCombo.Items.Add($_) }
$PSTCombo.BackColor = $OutText.BackColor
$PSTCombo.SelectedIndex = -1

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

# Check if the PST files exist locally
if (!(Test-Path "$LocalSupportTool") -or !(Test-Path "$LocalConfigs")) {
    $RefreshButtonText = "Import PST Files"
    $PSTCombo.Enabled = $false
}
else {
    $RefreshButtonText = "Refresh PST Files"
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

<# # Event handler for the environment switch button
$SelectEnvButton.Add_Click({
    if (!(Test-Path "$ResolvedLocalSupportTool")) {
        $OutText.AppendText("$(Get-Timestamp) - Local Prod Support Tool does not exist. Please import the PST files first.`r`n")
    }
    else {
        If (Test-Path "$ResolvedLocalSupportTool\ProductionSupportTool.exe.config.old") {
            $OutText.AppendText("$(Get-Timestamp) - Existing config found; cleaning up`r`n")   
            Reset-Environment
        }
        else {
            $ResetEnvButton.Enabled = $false
            $ResetEnvButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
            $ResetEnvButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
        }
                
        $script:PSTEnvironment = $PSTCombo.SelectedItem
            switch ($script:PSTEnvironment) {
                "QA" {
                    $OutText.AppendText("$(Get-Timestamp) - Entering QA environment`r`n")
                    $RunningConfig = "$ResolvedLocalConfigs\ProductionSupportTool.exe_QA.config"
                }
                "Stage" {
                    $OutText.AppendText("$(Get-Timestamp) - Entering Staging environment`r`n")
                    $RunningConfig = "$ResolvedLocalConfigs\ProductionSupportTool.exe_Stage.config"
                }
                "Production" {
                    $OutText.AppendText("$(Get-Timestamp) - Entering Production environment`r`n")
                    $RunningConfig = "$ResolvedLocalConfigs\ProductionSupportTool.exe_Prod.config"
                }
                # Behavior if no option is selected
                Default {
                    $OutText.AppendText("$(Get-Timestamp) - Please make a valid selection or reset`r`n")
                    throw "No selection made"
                }
            }
            Start-Sleep -Seconds 1
            # Rename Current Running config and Copy configuration file for correct environment
            Rename-Item "$ResolvedLocalSupportTool\ProductionSupportTool.exe.config" -NewName "$ResolvedLocalSupportTool\ProductionSupportTool.exe.config.old"
            Copy-Item $RunningConfig -Destination "$ResolvedLocalSupportTool\ProductionSupportTool.exe.config"
            $OutText.AppendText("$(Get-Timestamp) - You are ready to run in the $script:PSTEnvironment environment`r`n")
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
    }
}) #>

$SelectEnvButton.Add_Click({
    if (!$ResolvedLocalConfigs) {
        $OutText.AppendText("$(Get-Timestamp) - Local Configs path is not set. Please ensure the configs path is resolved.`r`n")
        return
    }

    if (!(Test-Path $ResolvedLocalSupportTool)) {
        $OutText.AppendText("$(Get-Timestamp) - Local Prod Support Tool does not exist. Please import the PST files first.`r`n")
        return
    }

    if (Test-Path "$ResolvedLocalSupportTool\ProductionSupportTool.exe.config.old") {
        $OutText.AppendText("$(Get-Timestamp) - Existing config found; cleaning up`r`n")   
        Reset-Environment
    } else {
        $ResetEnvButton.Enabled = $false
        $ResetEnvButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
        $ResetEnvButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
    }
            
    $script:PSTEnvironment = $PSTCombo.SelectedItem
    switch ($script:PSTEnvironment) {
        "QA" {
            $OutText.AppendText("$(Get-Timestamp) - Entering QA environment`r`n")
            $RunningConfig = Join-Path -Path $ResolvedLocalConfigs -ChildPath "ProductionSupportTool.exe_QA.config"
        }
        "Stage" {
            $OutText.AppendText("$(Get-Timestamp) - Entering Staging environment`r`n")
            $RunningConfig = Join-Path -Path $ResolvedLocalConfigs -ChildPath "ProductionSupportTool.exe_Stage.config"
        }
        "Production" {
            $OutText.AppendText("$(Get-Timestamp) - Entering Production environment`r`n")
            $RunningConfig = Join-Path -Path $ResolvedLocalConfigs -ChildPath "ProductionSupportTool.exe_Prod.config"
        }
        Default {
            $OutText.AppendText("$(Get-Timestamp) - Please make a valid selection or reset`r`n")
            throw "No selection made"
        }
    }

    if (!(Test-Path $RunningConfig)) {
        $OutText.AppendText("$(Get-Timestamp) - The configuration file does not exist: $RunningConfig`r`n")
        return
    }

    Start-Sleep -Seconds 1

    try {
        Rename-Item "$ResolvedLocalSupportTool\ProductionSupportTool.exe.config" -NewName "$ResolvedLocalSupportTool\ProductionSupportTool.exe.config.old" -ErrorAction Stop
        Copy-Item $RunningConfig -Destination "$ResolvedLocalSupportTool\ProductionSupportTool.exe.config" -ErrorAction Stop
        $OutText.AppendText("$(Get-Timestamp) - You are ready to run in the $script:PSTEnvironment environment`r`n")
    } catch {
        $OutText.AppendText("$(Get-Timestamp) - An error occurred: $_`r`n")
        return
    }

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
if ((Test-Path $LocalSupportTool) -and (Test-Path $LocalConfigs)) {
    $RefreshPSTButton.Text = "Refresh PST Files"
} else {
    $RefreshPSTButton.Text = "Import PST Files"
}

# Event handler for the run PST button
$RunPSTButton.Add_Click({
    $OutText.AppendText("$(Get-Timestamp) - Launching Prod Support Tool in $script:PSTEnvironment.`r`n")
    Start-Process -FilePath "$ResolvedLocalSupportTool\ProductionSupportTool.exe"
})

# Event handler for importing/refreshing the PST files
$RefreshPSTButton.Add_Click({
    Update-PSTFiles -Team $Team -Category $Category
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
    @('QA', 'Staging', 'Production') | ForEach-Object { [void]$script:LenderLFPCombo.Items.Add($_) }

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

    # Checkbox to confirm Production runs
    $script:ProductionCheckBox = New-Object System.Windows.Forms.CheckBox
    $script:ProductionCheckBox.Location = New-Object System.Drawing.Size(5, 215)
    $script:ProductionCheckBox.Size = New-Object System.Drawing.Size(225, 20)
    $script:ProductionCheckBox.Text = "Check to confirm PRODUCTION run"
    $script:ProductionCheckBox.Checked = $false
    $script:ProductionCheckBox.Enabled = $false

    # Button for adding lender to LFP services
    $script:AddLenderLFPButton = New-Object System.Windows.Forms.Button
    $script:AddLenderLFPButton.Location = New-Object System.Drawing.Point(5, 250)
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

    # Event handler for checking to see if the Production checkbox should be enabled
    $script:LenderLFPCombo.Add_SelectedIndexChanged({
        if ($script:LenderLFPCombo.SelectedItem -eq "Production") {
            $script:LenderLFPPopup.Controls.Add($script:ProductionCheckBox)
            $script:ProductionCheckBox.Enabled = $true
        }
        else {
            $script:LenderLFPPopup.Controls.Remove($script:ProductionCheckBox)
            $script:ProductionCheckBox.Enabled = $false
            $script:ProductionCheckBox.Checked = $false
        }
    })

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

    # Event handler for Productrion checkbox
    $script:ProductionCheckBox.Add_CheckedChanged({
        Enable-AddLenderButton
    })

    # Button click event for adding lender to LFP services
    $script:AddLenderLFPButton.add_Click({
        $AddLenderScript = Join-Path -Path $PSScriptRoot -ChildPath $ConfigValues.AddLenderToLFPServicesScript
        $LenderId = $script:LenderLFPIdTextBox.Text
        $TicketNumber = $script:LenderLFPTicketTextBox.Text
        $Environment = $script:LenderLFPCombo.SelectedItem

        if ($script:LenderLFPIdTextBox.Text -match '^\d+$') {
            if ($script:LenderLFPTicketTextBox.Text -match '^[A-Za-z]{3}-\d{5}$') {
                if (Test-Path $AddLenderScript){
                    $script:LenderLFPIdTextBox.Text = ''
                    $script:LenderLFPTicketTextBox.Text = ''
                    $script:LenderLFPCombo.SelectedItem = $null
                    $OutText.AppendText("$(Get-Timestamp) - Beginning process to add lender to LFP services...`r`n")
                    $OutText.AppendText("$(Get-Timestamp) - Lender to be added: $($LenderId)`r`n")
                    $OutText.AppendText("$(Get-Timestamp) - Environment: $($Environment)`r`n")
                    $OutText.AppendText("$(Get-Timestamp) - Jira ticket number: $($TicketNumber)`r`n")
                    Open-AddLenderScriptRunspace -AddLenderScript $AddLenderScript -LenderId $LenderId -TicketNumber $TicketNumber -Environment $Environment -OutText $OutText -TimestampFunction ${function:Get-Timestamp} -ActiveTicketsFunction ${function:Get-ActiveListItems}
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
    $script:CreateHDTStorageButton = New-Object System.Windows.Forms.Button
    $script:CreateHDTStorageButton.Location = New-Object System.Drawing.Size(20, 250)
    $script:CreateHDTStorageButton.BackColor = $global:DisabledBackColor
    $script:CreateHDTStorageButton.ForeColor = $global:DisabledForeColor
    $script:CreateHDTStorageButton.FlatStyle = "Popup"
    $script:CreateHDTStorageButton.Width = 75
    $script:CreateHDTStorageButton.Text = "Create"
    $script:CreateHDTStorageButton.Enabled = $false

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
        $ToolTip.SetToolTip($script:CreateHDTStorageButton, "Click to create the HDTStorage table")
    }

    # Event handler for the file location button
    $script:HDTStorageFileButton.Add_Click({
        $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog
        $FileBrowser.InitialDirectory = "C:\"
        $FileBrowser.Filter = "Excel Files (*.xlsx)|*.xlsx"
        $FileBrowserResult = $FileBrowser.ShowDialog()
        if ($FileBrowserResult -eq 'OK') {
            $script:HDTStoragePopup.Tag = $FileBrowser.FileName
            Enable-CreateHDTStorageButton
        }
    })

    # Event handler for the SQL instance text box
    $script:DBServerTextBox.Add_TextChanged({
        Enable-CreateHDTStorageButton
    })

    # Event handler for the table name text box
    $script:TableNameTextBox.Add_TextChanged({
        Enable-CreateHDTStorageButton
    })

    # Event handler for the secure password text box
    $script:SecurePasswordTextBox.Add_TextChanged({
        Enable-CreateHDTStorageButton
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
    $script:CreateHDTStorageButton.Add_Click({
        $LoginUsername = [System.Environment]::UserName
        $LoginPassword = ConvertTo-SecureString $script:SecurePasswordTextBox.Text -AsPlainText -Force
        $LoginCredentials = New-Object System.Management.Automation.PSCredential ($LoginUsername, $LoginPassword)

        # Set variables so values can be cleared from form
        $script:DBServerText = $script:DBServerTextBox.Text
        $script:TableNameText = $script:TableNameTextBox.Text

        # Clear form values
        $script:DBServerTextBox.Text = ''
        $script:TableNameTextBox.Text = ''
        $script:SecurePasswordTextBox.Text = ''
        $script:HDTStoragePWCheckbox.Checked = $false

        $OutText.AppendText("$(Get-Timestamp) - Creating HDTStorage table with the following values:`r`n")
        $OutText.AppendText("$(Get-Timestamp) - File Location: $($script:HDTStoragePopup.Tag)`r`n")
        $OutText.AppendText("$(Get-Timestamp) - SQL Instance: $script:DBServerText`r`n")
        $OutText.AppendText("$(Get-Timestamp) - Table Name: $script:TableNameText`r`n")

        Open-CreateHDTStorageRunspace -ConfigValuesCreateHDTStorageScript $ConfigValues.CreateHDTStorageScript -UserProfilePath $UserProfilePath -WorkhorseDirectoryPath $ConfigValues.WorkhorseDirectoryPath -WorkhorseServer $script:WorkhorseServer -FileTag $script:HDTStoragePopup.Tag -DBServerText $script:DBServerText -TableNameText $script:TableNameText -LoginCredentials $LoginCredentials -OutText $OutText -TimestampFunction ${function:Get-Timestamp}
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
    $script:HDTStoragePopup.Controls.Add($script:CreateHDTStorageButton)
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
$TicketsPath = Join-Path -Path $PSScriptRoot -ChildPath $ConfigValues.TicketsPath

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

# Initialize the ticket manager
Start-Setup
Get-ActiveListItems($ActiveTicketsListBox)
Get-CompletedListItems($CompletedTicketsListBox)

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
    $ActiveTicketsListBox.ClearSelected()
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
    $CompletedTicketsListBox.ClearSelected()
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
$OptionsMenu.DropDownItems.Add($MenuColorTheme) | Out-Null
$OptionsMenu.DropDownItems.Add($MenuThemeBuilder) | Out-Null
$OptionsMenu.DropDownItems.Add($MenuToolTips) | Out-Null
$MenuToolTips.DropDownItems.Add($ShowHelpBannerMenu) | Out-Null
$MenuToolTips.DropDownItems.Add($ShowToolTipsMenu) | Out-Null
$AboutMenu.DropDownItems.Add($MenuAboutItem) | Out-Null
$AboutMenu.DropDownItems.Add($MenuGitHub) | Out-Null
$MenuColorTheme.DropDownItems.Add($script:CustomThemes) | Out-Null
$MenuColorTheme.DropDownItems.Add($script:MLBThemes) | Out-Null
$MenuColorTheme.DropDownItems.Add($script:NBAThemes) | Out-Null
$MenuColorTheme.DropDownItems.Add($NFLThemes) | Out-Null
$MenuColorTheme.DropDownItems.Add($PremiumThemes) | Out-Null
$SysAdminTab.Controls.Add($RestartsTabControl)
$SysAdminTab.Controls.Add($ServersListBox)
$SysAdminTab.Controls.Add($AppListCombo)
$SysAdminTab.Controls.Add($script:RestartButton)
$SysAdminTab.Controls.Add($script:StartButton)
$SysAdminTab.Controls.Add($script:StopButton)
$SysAdminTab.Controls.Add($script:OpenSiteButton)
$SysAdminTab.Controls.Add($script:RestartIISButton)
$SysAdminTab.Controls.Add($script:StartIISButton)
$SysAdminTab.Controls.Add($script:StopIISButton)
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

    # Call the function to enable white text for USA theme
    if ($ConfigValues.DefaultUserTheme -eq 'USA') {
        Enable-USAThemeTextColor
    }

    # Check if the theme falls under Premium and set the AccentColor accordingly
    if ($SelectedTheme -eq 'Premium' -or $SelectedTheme -eq 'Custom') {
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
        $IconPath = Join-Path -Path $PSScriptRoot -ChildPath $ConfigValues.AlliedIcon

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

# When the form loads, set the icon if $icon is not $null
$Form.add_Load({
    if ($Icon) {
        $This.Icon = $Icon
    }
})

# Confirm there are no active popups before closing the form
# Clean up sync hash when form is closed
# Reset the PST config file
$Form.Add_FormClosing({
    try {
        if ($global:IsAboutPopupActive -or $global:IsLenderLFPPopupActive -or $global:IsHDTStoragePopupActive -or $global:IsFeedbackPopupActive -or $global:IsThemeBuilderPopupActive) {
            $_.Cancel = $true
            $OutText.AppendText("$(Get-Timestamp) - Please close any open popup windows before closing the main form.`r`n")
        }
        else {
            $synchash.Closed = $True
            if (Test-Path "$ResolvedLocalSupportTool\ProductionSupportTool.exe.config.old") {
                Reset-Environment
            }
        }
    }
    catch {
        $OutText.AppendText("$(Get-Timestamp) - Error closing form: $_`r`n")
    }
})

# Start the directory check jobs before showing the form
$RemoteSupportToolFolderDate = Start-DirectoryCheckJob -path $RemoteSupportTool -excludePattern "*.txt"
$LocalSupportToolFolderDate =  Start-DirectoryCheckJob -path $LocalSupportTool -excludePattern "*.txt"

# Show Form
$Form.ShowDialog() | Out-Null

<#
? **********************************************************************************************************************
? END OF FORM BUILD
? **********************************************************************************************************************
#>