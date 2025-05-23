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
#* 1.3.1 - 10/12/2023 - Theme-specific quotes
#* 1.4.0 - 10/18/2023 - Full Lender LFP module for QA, Staging, and Prod
#* 2.0.0 - 11/02/2023 - Theme Builder
#* 2.0.1 - 11/08/2023 - Minor rework of Prod Support Tool logic to check/refresh files
#* 2.1.0 - 11/21/2023 - Billing Service Restarts
#* 2.2.0 - 12/09/2023 - Cert Check
#* 2.2.1 - 12/13/2023 - Minor rework of Restarts GUI logic
#* 2.3.0 - 12/19/2023 - Color Picker for Theme Builder
#* 3.0.0 - 12/28/2023 - Help Icons
#* 4.0.0 - 01/14/2024 - AWS Monitoring GUI
#* 4.1.0 - 01/23/2024 - Font Picker

#* Table of contents for this script
#* Use Control + F to search for the section you want to jump to

#*  1. Global Variables and Functions
#*  2. Main GUI
#*  3. Menu Strip
#*  4. Restarts GUI
#*  5. NSLookup
#*  6. Server Ping
#*  7. Reverse IP Lookup
#*  8. Cert Check
#*  9. AWS Monitoring GUI
#* 10. Prod Support Tool
#* 11. Add Lender to LFP Services
#* 12. Billing Service Restarts
#* 13. Password Manager
#* 14. Create HDTStorage Table
#* 15. Documentation Creator
#* 16. Ticket Manager
#* 17. Form Build

# ================================================================================================

# todo - List of to-do items

# todo - CURRENT ISSUES
# todo - None!

# todo - TECHNICAL DEBT
# todo - Add section in Cert Check to choose the expiration date cutoff OR return all certs regardless of expiration date
# todo - Add logic to check if popup windows are open when enabling/disabling tooltips
# todo - For AWS SSO login, automatically close runspaces after X amount of time
# todo - Create lists for all controls and use foreach loops to set properties
# todo - Return AWS CPU Metrics form status to main GUI for Update-MainTheme logic

# todo - UPCOMING FEATURES
# todo - Add size combo box to Font Picker
# todo - Add logic to rebuild servers.csv file from AWS Monitoring GUI
# todo - Search Event Viewer
# todo - Change function for setting text to white to accept any color choice
# todo - Module for adding EDI Trading Partners
# todo - Terraform Admin Tools

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

# Variables to keep track of Desktop Assistant form's position
$global:MouseDown = $false
$global:FormOffset = New-Object System.Drawing.Point

$global:SecurePW = $null
$global:AltSecurePW = $null
$global:IsThemeBuilderPopupActive = $false
$global:IsColorPickerPopupActive = $false
$global:IsFontPickerPopupActive = $false
$global:IsThemeApplied = $false
$global:IsLenderLFPPopupActive = $false
$global:IsHDTStoragePopupActive = $false
$global:IsFeedbackPopupActive = $false
$global:IsBillingRestartPopupActive = $false
$global:IsCertCheckPopupActive = $false
$script:PSTEnvironment = $null # Initializing variable for holding the selected PST Environment so it can be used across multiple event handlers
$global:WasMainThemeActivated = $false # Initializing variable for tracking whether the main theme has been activated yet
$global:WasThemePremium = $false # Initializing variable for tracking whether the theme is premium or not

# Hashtable for storing user's selections in Color Picker
$script:SelectedColors = @{}

# Initialize the tooltip object
$ToolTip = New-Object System.Windows.Forms.ToolTip
$ToolTip.InitialDelay = 100

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

# Get color themes from json file
$ColorTheme = Get-Content -Path .\ColorThemes.json | ConvertFrom-Json

# Not currently in use, here for reference
$DefaultFont = 'Arial'

$global:DefaultUserFont = $ConfigValues.DefaultUserFont

# Variables for form building functions
$ControlColor = [System.Drawing.SystemColors]::Control
$ControlColorText = [System.Drawing.SystemColors]::ControlText
$global:NormalFont = [System.Drawing.Font]::new($global:DefaultUserFont, 9, [System.Drawing.FontStyle]::Regular)
$global:NormalSmallFont = [System.Drawing.Font]::new($global:DefaultUserFont, 8, [System.Drawing.FontStyle]::Regular)
$global:NormalBoldFont = [System.Drawing.Font]::new($global:DefaultUserFont, 9, [System.Drawing.FontStyle]::Bold)
$global:NormalItalicFont = [System.Drawing.Font]::new($global:DefaultUserFont, 9, [System.Drawing.FontStyle]::Italic)
$global:NormalItalicBoldFont = [System.Drawing.Font]::new($global:DefaultUserFont, 9, [System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Italic)
$global:SmallItalicFont = [System.Drawing.Font]::new($global:DefaultUserFont, 8, [System.Drawing.FontStyle]::Italic)
$global:SmallBoldFont = [System.Drawing.Font]::new($global:DefaultUserFont, 8, [System.Drawing.FontStyle]::Bold)
$global:ExtraSmallFont = [System.Drawing.Font]::new($global:DefaultUserFont, 7, [System.Drawing.FontStyle]::Regular)
$global:ExtraSmallItalicBoldFont = [System.Drawing.Font]::new($global:DefaultUserFont, 7, [System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Italic)
$global:ExtraLargeFont = [System.Drawing.Font]::new($global:DefaultUserFont, 16, [System.Drawing.FontStyle]::Regular)
$DefaultTextAlign = [System.Drawing.ContentAlignment]::TopLeft
$MiddleTextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$TopRightTextAlign = [System.Drawing.ContentAlignment]::TopRight
$ArrowCursor = [System.Windows.Forms.Cursors]::Arrow
$HandCursor = [System.Windows.Forms.Cursors]::Hand
$AutoSizeMode = [System.Windows.Forms.PictureBoxSizeMode]::AutoSize
$ZoomSizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom

# Construct paths for icons and logos
$AlliedIcon = Join-Path -Path $PSScriptRoot -ChildPath $ConfigValues.AlliedIcon
$AlliedLogo = Join-Path -Path $PSScriptRoot -ChildPath $ConfigValues.AlliedLogo
$AboutMenuAlliedLogoImage = [System.Drawing.Image]::FromFile($AlliedLogo)
$HelpIconPath = Join-Path -Path $PSScriptRoot -ChildPath $ConfigValues.QuestionMarkIcon
$HelpIcon = [System.Drawing.Image]::FromFile($HelpIconPath)
$HelpIconImage = New-Object System.Drawing.Bitmap($HelpIcon)

# Webhook URL for sending Teams messages; currently used for Submit Feedback module
$TestingWebhookURL = $ConfigValues.TestingWebhook

# Get the path of the Workhorse server
$script:WorkhorseServer = $ConfigValues.WorkhorseServer

# Check if DefaultUserTheme has a value or is null
if ($null -ne $ConfigValues.DefaultUserTheme -and $ConfigValues.DefaultUserTheme -ne '') {
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

#Template for creating a new Windows form
function New-WindowsForm {
    param(
        [int]$SizeX,
        [int]$SizeY,
		[string]$Text,
		[System.Windows.Forms.FormStartPosition]$StartPosition,
		[bool]$TopMost,
		[bool]$ShowInTaskbar,
		[bool]$KeyPreview,
		[bool]$MinimizeBox
    )

    $Form = New-Object System.Windows.Forms.Form
    $Form.Size = New-Object System.Drawing.Size($SizeX, $SizeY)
	$Form.Text = $Text
	$Form.StartPosition = $StartPosition
	$Form.TopMost = $TopMost
	$Form.FormBorderStyle = 'FixedSingle'
	$Form.ControlBox = $false
	$Form.ShowInTaskbar = $ShowInTaskbar
	$Form.KeyPreview = $KeyPreview
	$Form.AutoSize = $true
	$Form.MaximizeBox = $false
	$Form.MinimizeBox = $MinimizeBox

    return $Form
}

# Template for creating a new form button
function New-FormButton {
    param(
        [string]$Text,
        [int]$LocationX,
        [int]$LocationY,
        [int]$Width,
        [System.Drawing.Color]$BackColor,
        [System.Drawing.Color]$ForeColor,
        [System.Drawing.Font]$Font,
        [bool]$Enabled,
        [bool]$Visible
    )

    $Button = New-Object System.Windows.Forms.Button
    $Button.Location = New-Object System.Drawing.Point($LocationX, $LocationY)
    $Button.Width = $Width
	$Button.BackColor = $BackColor
    $Button.ForeColor = $ForeColor
    $Button.FlatStyle = "Popup"
    $Button.Text = $Text
    $Button.Font = $Font
    $Button.Enabled = $Enabled
    $Button.Visible = $Visible

    return $Button
}

# Template for creating a new form label
function New-FormLabel {
    param(
        [int]$LocationX,
        [int]$LocationY,
		[int]$SizeX,
		[int]$SizeY,
		[string]$Text,
        [System.Drawing.Font]$Font,
		[System.Drawing.ContentAlignment]$TextAlign
    ) 

    $Label = New-Object System.Windows.Forms.Label
    $Label.Location = New-Object System.Drawing.Point($LocationX, $LocationY)
    $Label.Size = New-Object System.Drawing.Size($SizeX, $SizeY)
    $Label.Text = $Text
    $Label.Font = $Font
	$Label.TextAlign = $TextAlign

    return $Label
}

# Template for creating a new form textbox
function New-FormTextBox {
    param(
        [int]$LocationX,
        [int]$LocationY,
		[int]$SizeX,
		[int]$SizeY,
		[System.Windows.Forms.ScrollBars]$ScrollBars,
        [bool]$Multiline,
		[bool]$Enabled,
		[bool]$ReadOnly,
		[string]$Text,
		[System.Drawing.Font]$Font
    )

    $TextBox = New-Object System.Windows.Forms.TextBox
    $TextBox.Location = New-Object System.Drawing.Point($LocationX, $LocationY)
    $TextBox.Size = New-Object System.Drawing.Size($SizeX, $SizeY)
    $TextBox.ScrollBars = $ScrollBars
    $TextBox.Multiline = $Multiline
    $TextBox.Enabled = $Enabled
    $TextBox.ReadOnly = $ReadOnly
    $TextBox.Text = $Text
    $TextBox.Font = $Font
    $TextBox.ShortcutsEnabled = $true

    return $TextBox
}

# Template for creating a new form listbox
function New-FormListBox {
    param(
        [int]$LocationX,
        [int]$LocationY,
        [int]$SizeX,
        [int]$SizeY,
        [System.Drawing.Font]$Font,
        [System.Windows.Forms.SelectionMode]$SelectionMode
    )

    $ListBox = New-Object System.Windows.Forms.ListBox
    $ListBox.Location = New-Object System.Drawing.Point($LocationX, $LocationY)
    $ListBox.Size = New-Object System.Drawing.Size($SizeX, $SizeY)
    $ListBox.Font = $Font
    $ListBox.SelectionMode = $SelectionMode

    return $ListBox
}

# Template for creating a new form combobox
function New-FormComboBox {
    param(
        [int]$LocationX,
        [int]$LocationY,
        [int]$SizeX,
        [int]$SizeY,
        [System.Drawing.Font]$Font
    )

    $ComboBox = New-Object System.Windows.Forms.ComboBox
    $ComboBox.Location = New-Object System.Drawing.Point($LocationX, $LocationY)
    $ComboBox.Size = New-Object System.Drawing.Size($SizeX, $SizeY)
	$ComboBox.Font = $Font
    $ComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
	$ComboBox.SelectedIndex = -1

    return $ComboBox
}

# Function for creating a new form picturebox
function New-FormPictureBox {
    param(
        [int]$LocationX,
        [int]$LocationY,
        [int]$SizeX,
        [int]$SizeY,
        [System.Drawing.Image]$Image,
        [System.Windows.Forms.PictureBoxSizeMode]$SizeMode,
        [System.Windows.Forms.Cursor]$Cursor
    )

    $PictureBox = New-Object System.Windows.Forms.PictureBox
    $PictureBox.Location = New-Object System.Drawing.Point($LocationX, $LocationY)
    $PictureBox.Size = New-Object System.Drawing.Size($SizeX, $SizeY)
    $PictureBox.Image = $Image
    $PictureBox.SizeMode = $SizeMode
    $PictureBox.Cursor = $Cursor

    return $PictureBox
}

# Template for creating a new form checkbox
function New-FormCheckbox {
    param(
		[int]$LocationX,
		[int]$LocationY,
        [int]$SizeX,
        [int]$SizeY,
		[string]$Text,
		[System.Drawing.Font]$Font,
		[bool]$Checked,
		[bool]$Enabled
    )

    $CheckBox = New-Object System.Windows.Forms.CheckBox
	$CheckBox.Location = New-Object System.Drawing.Point($LocationX, $LocationY)
    $CheckBox.Size = New-Object System.Drawing.Size($SizeX, $SizeY)
	$CheckBox.Text = $Text
	$CheckBox.Font = $Font
	$CheckBox.Checked = $Checked
	$CheckBox.Enabled = $Enabled

    return $CheckBox
}

# Template for creating a new form TabPage
function New-FormTabPage {
    param(
		[System.Drawing.Font]$Font,
		[string]$Name,
		[string]$Text
    )

    $TabPage = New-Object System.Windows.Forms.TabPage
	$TabPage.DataBindings.DefaultDataSourceUpdateMode = 0
	$TabPage.UseVisualStyleBackColor = $true
	$TabPage.Font = $Font
	$TabPage.Name = $Name
	$TabPage.Text = $Text

    return $TabPage
}

# Template for creating a new form timer
function New-FormTimer {
    param(
        [int]$Interval,
        [bool]$Enabled
    )

    $Timer = New-Object System.Windows.Forms.Timer
    $Timer.Interval = $Interval
    $Timer.Enabled = $Enabled

    return $Timer
}

# Function to return the current time in EST
function Get-Timestamp {
    # Get the current time (this will be in the time zone of the system where the script is running)
    $currentTime = Get-Date

    # Define the EST time zone
    $easternTimeZone = [System.TimeZoneInfo]::FindSystemTimeZoneById("Eastern Standard Time")

    # Convert the current time to EST
    $currentTimeInEST = [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId($currentTime, $easternTimeZone.Id)

    # Return the time in EST
    return $currentTimeInEST.ToString("yyyy/MM/dd hh:mm:ss")
}

# Function to show the Help Form popup near the icon
function Show-HelpForm {
    param(
        [System.Windows.Forms.PictureBox]$PictureBox,
        [string]$HelpText
    )
    # Create the label with help text
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $HelpText
    $label.AutoSize = $true
    $label.MaximumSize = New-Object System.Drawing.Size(180, 0)
    $label.Padding = New-Object System.Windows.Forms.Padding(10)

    # Create the second label for the closing instructions
    $InstructionLabel = New-Object System.Windows.Forms.Label
    $InstructionLabel.Text = "Click anywhere outside this popup to close it."
    $InstructionLabel.Font = New-Object System.Drawing.Font($InstructionLabel.Font, [System.Drawing.FontStyle]::Italic)
    $InstructionLabel.AutoSize = $true
    $InstructionLabel.Padding = New-Object System.Windows.Forms.Padding(10)
    $InstructionLabel.MaximumSize = New-Object System.Drawing.Size(180, 0)

    if ($null -ne $ConfigValues.DefaultUserTheme -and $ConfigValues.DefaultUserTheme -ne '') {
        $SelectedTheme = $ColorTheme.PSObject.Properties | Where-Object { $_.Value.$($ConfigValues.DefaultUserTheme) } | Select-Object -ExpandProperty Name
        $themeColors = $ColorTheme.$SelectedTheme.$($ConfigValues.DefaultUserTheme)
        $HelpForm.BackColor = $themeColors.BackColor
        $HelpForm.ForeColor = $themeColors.ForeColor
        $label.ForeColor = $themeColors.ForeColor
    }

    # Clear existing controls and add the new label to the popup form
    $HelpForm.Controls.Clear()
    $HelpForm.Controls.Add($label)
    $HelpForm.Controls.Add($InstructionLabel)

    # Adjust the label positions and the popup form's size
    $label.Location = New-Object System.Drawing.Point(10, 5)
    $InstructionLabelY = [int]($label.Height + $label.Top)  # Calculate Y position for the instruction label
    $InstructionLabel.Location = New-Object System.Drawing.Point(10, $InstructionLabelY)


    # Calculate the total height of the content
    $totalContentHeight = $label.Height + $InstructionLabel.Height + 10  # 10 for padding

    # Ensure the label dimensions are integers
    $labelWidth = [int]$label.Width
    $labelHeight = [int]$totalContentHeight

    # Check if label dimensions are valid before proceeding
    if ($labelWidth -gt 0 -and $labelHeight -gt 0) {
        # Adjust the popup form's size based on the label's size
        $HelpFormSizeWidth = $labelWidth + 10
        $HelpFormSizeHeight = $labelHeight + 10
        $HelpForm.Size = New-Object System.Drawing.Size($HelpFormSizeWidth, $HelpFormSizeHeight)

        # Calculate and set the popup form's location
        $x = $DesktopAssistantForm.Location.X + $PictureBox.Location.X + $PictureBox.Width
        $y = $DesktopAssistantForm.Location.Y + $PictureBox.Location.Y + $PictureBox.Height
        $HelpForm.Location = New-Object System.Drawing.Point($x, $y)

        # Show the popup form
        $HelpForm.Show()
    } else {
        $OutText.AppendText("$(Get-Timestamp) - Invalid label dimensions: Width=$labelWidth, Height=$labelHeight")
    }
}

# Function for enabling the Font Picker Apply Choice button
function Enable-FontPickerApplyChoiceButton {
    if ($FontPickerListBox.SelectedIndex -ge 0) {
        $script:FontPickerApplyButton.Enabled = $true
        $script:FontPickerApplyButton
    }
    else {
        $script:FontPickerApplyButton.Enabled = $false
    }

    if ($script:FontPickerApplyButton.Enabled) {
        $backColor = Get-AppropriateColor -ColorType "BackColor"
        $foreColor = Get-AppropriateColor -ColorType "ForeColor"

        $script:FontPickerApplyButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
        $script:FontPickerApplyButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
    } else {
        $script:FontPickerApplyButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
        $script:FontPickerApplyButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
    }
}

# Function for passing the new chosen font to all form controls
function Update-FormFonts {
    param(
        [System.Windows.Forms.Control]$Control,
        [hashtable]$FontMapping
    )

    $controlName = if ([string]::IsNullOrWhiteSpace($Control.Name)) { "Unnamed control" } else { $Control.Name }

    if ($Control -is [System.Windows.Forms.Control] -and $Control.GetType().GetProperty('Font')) {
        $oldFont = $Control.Font
        $fontStyleString = $oldFont.Style.ToString()
        $fontStyleString = $fontStyleString -replace ' ', '' -replace ',', ', '
        
        # Ensure the size is formatted correctly, e.g., as a whole number or a specific decimal
        $fontSizeFormatted = "{0:N1}" -f $oldFont.Size  # Formats to 1 decimal place, adjust as needed
        
        $fontKey = "$fontSizeFormatted-$fontStyleString"
        $newFont = $FontMapping[$fontKey]

        if ($null -ne $newFont) {
            $Control.Font = $newFont
        } else {
            Write-Host "No matching font found for $controlName with key $fontKey. Available keys: $($FontMapping.Keys -join ', ')"
        }
    }

    if ($Control.HasChildren) {
        foreach ($childControl in $Control.Controls) {
            Update-FormFonts -Control $childControl -FontMapping $FontMapping
        }
    }
}

# Function to get the quote for the selected theme
function Get-ThemeQuote {
    $script:Quote = $script:ColorTheme.$script:SelectedTheme.$($ConfigValues.DefaultUserTheme).Quote
    $OutText.AppendText("$(Get-Timestamp) - $($ConfigValues.DefaultUserTheme) theme is now active.`r`n")
    if ($script:Quote) {
        $OutText.AppendText("$(Get-Timestamp) - $script:Quote`r`n")
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
    $ToolTip.SetToolTip($RestartsGUIHelpIcon, "Click to learn more about Server Restarts")
	$ToolTip.SetToolTip($script:RestartButton, "Click to restart selected item(s)")
	$ToolTip.SetToolTip($script:StartButton, "Click to start selected item(s)")
	$ToolTip.SetToolTip($script:StopButton, "Click to stop selected item(s)")
	$ToolTip.SetToolTip($script:OpenSiteButton, "Click to open selected site in Windows Explorer")
	$ToolTip.SetToolTip($script:RestartIISButton, "Click to restart IIS on selected server")
	$ToolTip.SetToolTip($script:StartIISButton, "Click to start IIS on selected server")
	$ToolTip.SetToolTip($script:StopIISButton, "Click to stop IIS on selected server")
	$Tooltip.SetToolTip($NSLookupButton, "Click to run nslookup")
	$Tooltip.SetToolTip($script:NSLookupTextBox, "Enter a hostname to resolve")
    $ToolTip.SetToolTip($ReverseIPButton, "Click to run reverse IP lookup")
    $ToolTip.SetToolTip($ReverseIPTextBox, "Enter a DNS name to resolve")
	$ToolTip.SetToolTip($ServerPingButton, "Click to ping server")
    $ToolTip.SetToolTip($ServerPingHelpIcon, "Click to learn more about Server Ping")
    $ToolTip.SetToolTip($NSLookupHelpIcon, "Click to learn more about NSLookup")
    $ToolTip.SetToolTip($ReverseIPHelpIcon, "Click to learn more about Reverse IP Lookup")
	$ToolTip.SetToolTip($ServerPingTextBox, "Enter a server name or IP address to test the connection")
    $ToolTip.SetToolTip($CertCheckWizardButton, "Click to launch the Cert Check Wizard")
    $ToolTip.SetToolTip($CertCheckHelpIcon, "Click to learn more about Cert Check Wizard")
    $ToolTip.SetToolTip($script:AWSSSOLoginButton, "Click to login to your AWS account using SSO; This will open a new browser window")
    $ToolTip.SetToolTip($script:AWSSSOLoginTimerLabel, "The amount of time remaining in your AWS session (default is 1 hour)")
    $ToolTip.SetToolTip($script:ListAWSAccountsButton, "Click to list all AWS accounts")
    $ToolTip.SetToolTip($AWSMonitoringGUIHelpIcon, "Click to learn more about AWS Admin Tools")
    $ToolTip.SetToolTip($script:AWSAccountsListBox, "Select an AWS account to list its child accounts")
    $ToolTip.SetToolTip($script:AWSInstancesListBox, "Select one or more instances to take action on")
    $ToolTip.SetToolTip($script:RebootAWSInstancesButton, "Click to reboot one or more AWS EC2 instances")
    $ToolTip.SetToolTip($script:StartAWSInstancesButton, "Click to start one or more AWS EC2 instances")
    $ToolTip.SetToolTip($script:StopAWSInstancesButton, "Click to stop one or more AWS EC2 instances")
    $ToolTip.SetToolTip($script:AWSScreenshotButton, "Click to take a screenshot of an AWS EC2 instance")
    $ToolTip.SetToolTip($script:AWSCPUMetricsButton, "Click to gather CPU metrics of an AWS EC2 instance for a given timeframe")
    $ToolTip.SetToolTip($PSTHelpIcon, "Click to learn more about UniTrac Prod Support Tool")
	$ToolTip.SetToolTip($PSTCombo, "Select the environment to run the Prod Support Tool in")
	$ToolTip.SetToolTip($SelectEnvButton, "Click to switch environment to run the Prod Support Tool in")
	$ToolTip.SetToolTip($ResetEnvButton, "Click to reset the environment configuration")
	$ToolTip.SetToolTip($RunPSTButton, "Click to run the Prod Support Tool")
	$ToolTip.SetToolTip($RefreshPSTButton, "Click to refresh the Prod Support Tool files")
    $ToolTip.SetToolTip($LaunchLFPWizardButton, "Click to launch the Add Lender to LFP wizard")
    $ToolTip.SetToolTip($LaunchBillingRestartButton, "Click to launch the Billing Restarts wizard")
    $ToolTip.SetToolTip($LFPWizardHelpIcon, "Click to learn more about LFP Wizard")
    $ToolTip.SetToolTip($BillingRestartHelpIcon, "Click to learn more about Billing Restarts Wizard")
	$ToolTip.SetToolTip($PWTextBox, "Enter a password to set for the remainder of the session")
	$ToolTip.SetToolTip($SetPWButton, "Click to set your password")
	$ToolTip.SetToolTip($AltSetPWButton, "Click to set an alternate password")
	$ToolTip.SetToolTip($GetPWButton, "Click to retrieve your password")
	$ToolTip.SetToolTip($AltGetPWButton, "Click to retrieve your alternate password")
	$ToolTip.SetToolTip($ClearPWButton, "Click to clear your password")
	$ToolTip.SetToolTip($AltClearPWButton, "Click to clear your alternate password")
	$ToolTip.SetToolTip($GenPWButton, "Click to generate a 16 character password with at least 1 uppercase, 1 lowercase, 1 number, and 1 special character")
	$ToolTip.SetToolTip($LaunchHDTStorageButton, "Click to launch the HDTStorage table creator wizard")
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
    $ToolTip.SetToolTip($PWManagerHelpIcon, "Click to learn more about Password Manager")
    $ToolTip.SetToolTip($GenPWHelpIcon, "Click to learn more about Password Generator")
    $ToolTip.SetToolTip($HDTStorageHelpIcon, "Click to learn more about HDTStorage Table Wizard")
    $ToolTip.SetToolTip($NewDocHelpIcon, "Click to learn more about Documentation Creator")
    $ToolTip.SetToolTip($NewTicketHelpIcon, "Click to learn more about creating tickets in Ticket Manager")
    $ToolTip.SetToolTip($RenameTicketHelpIcon, "Click to learn more about renaming tickets in Ticket Manager")
	
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
            $cleanText = $item.Text -replace "• ", ''  # Remove bullet point if present
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

    if ($global:IsFeedbackPopupActive) {
        $script:UserNameTextBox.ForeColor = 'White'
        $script:FeedbackTextBox.ForeColor = 'White'
    }

    if ($global:IsCertCheckPopupActive) {
        $script:CertCheckServerTextBox.ForeColor = 'White'
    }

    if ($global:IsLenderLFPPopupActive) {
        $script:LenderLFPIdTextBox.ForeColor = 'White'
        $script:LenderLFPTicketTextBox.ForeColor = 'White'
    }

    if ($global:IsHDTStoragePopupActive) {
        $script:DBServerTextBox.ForeColor = 'White'
        $script:TableNameTextBox.ForeColor = 'White'
        $script:SecurePasswordTextBox.ForeColor = 'White'
    }

    # Define common properties for all controls
    $script:whiteBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::White)
    $script:fontArialRegular9 = [System.Drawing.Font]::new($global:DefaultUserFont, 9, [System.Drawing.FontStyle]::Regular)
    $script:itemHeight = 15

    # An array of all controls that need to be owner-drawn
    $allControls = @($PSTCombo, $ActiveTicketsListBox, $CompletedTicketsListBox, $FolderContentsListBox,
                     $ServersListBox, $ServicesListBox, $IISSitesListBox, $AppPoolsListBox, $AppListCombo, $script:AWSAccountsListBox, $script:AWSInstancesListBox)
    if ($global:IsLenderLFPPopupActive){                 
    $allControls += $script:LenderLFPCombo
    }
    if ($global:IsBillingRestartPopupActive){
        $allControls += $script:BillingRestartCombo
    }
    if ($global:IsFontPickerPopupActive) {
        $allControls += $script:FontPickerListBox
    }

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
    if ($global:IsLenderLFPPopupActive){
        $script:LenderLFPCombo.Size = New-Object System.Drawing.Size(200,200)
    }
    if ($global:IsBillingRestartPopupActive){
        $script:BillingRestartCombo.Size = New-Object System.Drawing.Size(150,200)
    }
    if ($global:IsFontPickerPopupActive) {
        $script:FontPickerListBox.Size = New-Object System.Drawing.Size(370,540)
    }
    $ServicesListBox.Size = New-Object System.Drawing.Size(245,240)
    $IISSitesListBox.Size = New-Object System.Drawing.Size(245,240)
    $AppPoolsListBox.Size = New-Object System.Drawing.Size(245,240)
    $ServersListBox.Size = New-Object System.Drawing.Size(200,240)
    $ActiveTicketsListBox.Size = New-Object System.Drawing.Size(215,240)
    $CompletedTicketsListBox.Size = New-Object System.Drawing.Size(215,240)
    $FolderContentsListBox.Size = New-Object System.Drawing.Size(220,255)
    $script:AWSAccountsListBox.Size = New-Object System.Drawing.Size(200,240)
    $script:AWSInstancesListBox.Size = New-Object System.Drawing.Size(200,240)
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

    if ($global:IsFeedbackPopupActive) {
        $script:UserNameTextBox.ForeColor = 'Black'
        $script:FeedbackTextBox.ForeColor = 'Black'
    }

    if ($global:IsCertCheckPopupActive) {
        $script:CertCheckServerTextBox.ForeColor = 'Black'
    }

    if ($global:IsLenderLFPPopupActive) {
        $script:LenderLFPIdTextBox.ForeColor = 'Black'
        $script:LenderLFPTicketTextBox.ForeColor = 'Black'
    }

    if ($global:IsHDTStoragePopupActive) {
        $script:DBServerTextBox.ForeColor = 'Black'
        $script:TableNameTextBox.ForeColor = 'Black'
        $script:SecurePasswordTextBox.ForeColor = 'Black'
    }

    # An array of all combo and list boxes to revert to normal drawing mode
    $allControls = @($PSTCombo, $ActiveTicketsListBox, $CompletedTicketsListBox, $FolderContentsListBox, 
                    $ServersListBox, $ServicesListBox, $IISSitesListBox, $AppPoolsListBox, $AppListCombo, $script:AWSAccountsListBox, $script:AWSInstancesListBox)
    if ($global:IsLenderLFPPopupActive){                 
        $allControls += $script:LenderLFPCombo
    }
    if ($global:IsBillingRestartPopupActive){
        $allControls += $script:BillingRestartCombo
    }
    if ($global:IsFontPickerPopupActive) {
        $allControls += $script:FontPickerListBox
    }

    # Loop through each control and reset its properties
    foreach ($control in $allControls) {
        $control.DrawMode = [System.Windows.Forms.DrawMode]::Normal
        $control.Font = [System.Drawing.Font]::new($global:DefaultUserFont, 9, [System.Drawing.FontStyle]::Regular)
        $control.ItemHeight = 13

        # Remove any DrawItem event handlers associated with the control
        $control.remove_DrawItem($control.DrawItem)

        # Refresh the control to re-draw with updated settings
        $control.Refresh()
    }
    # Explicitly set the size of the ListBox after the theme change
    if ($global:IsLenderLFPPopupActive){
        $script:LenderLFPCombo.Size = New-Object System.Drawing.Size(200,200)
    }
    if ($global:IsBillingRestartPopupActive){
        $script:BillingRestartCombo.Size = New-Object System.Drawing.Size(150,200)
    }
    if ($global:IsFontPickerPopupActive) {
        $script:FontPickerListBox.Size = New-Object System.Drawing.Size(370,540)
    }
    $ServicesListBox.Size = New-Object System.Drawing.Size(245,240)
    $IISSitesListBox.Size = New-Object System.Drawing.Size(245,240)
    $AppPoolsListBox.Size = New-Object System.Drawing.Size(245,240)
    $ServersListBox.Size = New-Object System.Drawing.Size(200,240)
    $ActiveTicketsListBox.Size = New-Object System.Drawing.Size(215,240)
    $CompletedTicketsListBox.Size = New-Object System.Drawing.Size(215,240)
    $FolderContentsListBox.Size = New-Object System.Drawing.Size(220,255)
    $script:AWSAccountsListBox.Size = New-Object System.Drawing.Size(200,240)
    $script:AWSInstancesListBox.Size = New-Object System.Drawing.Size(200,240)
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

    if ($ConfigValues.DefaultUserTheme -eq "Manny's") {
        # Add the Deadpool--YES, FREAKING DEADPOOL--to the form
        $DesktopAssistantForm.Controls.Add($DeadpoolsVeryOwnPictureBoxTM)
    }
    else {
        # Remove the Deadpool from the form
        $DesktopAssistantForm.Controls.Remove($DeadpoolsVeryOwnPictureBoxTM)
    }
    
    if ($Category -eq 'Premium' -or $Category -eq 'Custom') {
        $global:CurrentTeamAccentColor = $ColorData.$Category.$Team.AccentColor
        $OutText.BackColor = $global:CurrentTeamAccentColor
        $ServersListBox.Backcolor = $global:CurrentTeamAccentColor
        $ServicesListBox.Backcolor = $global:CurrentTeamAccentColor
        $IISSitesListBox.Backcolor = $global:CurrentTeamAccentColor
        $AppPoolsListBox.Backcolor = $global:CurrentTeamAccentColor
        $script:AWSAccountsListBox.Backcolor = $global:CurrentTeamAccentColor
        $script:AWSInstancesListBox.Backcolor = $global:CurrentTeamAccentColor
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
        $script:AWSAccountsListBox.Backcolor = [System.Drawing.SystemColors]::Control
        $script:AWSInstancesListBox.Backcolor = [System.Drawing.SystemColors]::Control
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

    $MainFormMinimizeButton.FlatAppearance.MouseOverBackColor = if ($Category -eq 'Premium' -or $Category -eq 'Custom') { $global:CurrentTeamAccentColor } else { $global:CurrentTeamForeColor }
    $MainFormMinimizeButton.Add_MouseEnter({ if ($Category -eq 'Premium' -or $Category -eq 'Custom') { $MainFormMinimizeButton.ForeColor = $global:CurrentTeamForeColor } else { $MainFormMinimizeButton.ForeColor = $global:CurrentTeamBackColor } })
    $MainFormMinimizeButton.Add_MouseLeave({ $MainFormMinimizeButton.ForeColor = $global:CurrentTeamForeColor })

    $MainFormCloseButton.FlatAppearance.MouseOverBackColor = if ($Category -eq 'Premium' -or $Category -eq 'Custom') { $global:CurrentTeamAccentColor } else { $global:CurrentTeamForeColor }
    $MainFormCloseButton.Add_MouseEnter({ if ($Category -eq 'Premium' -or $Category -eq 'Custom') { $MainFormCloseButton.ForeColor = $global:CurrentTeamForeColor } else { $MainFormCloseButton.ForeColor = $global:CurrentTeamBackColor } })
    $MainFormCloseButton.Add_MouseLeave({ $MainFormCloseButton.ForeColor = $global:CurrentTeamForeColor })

    $DesktopAssistantForm.BackColor = $ColorData.$Category.$Team.BackColor
    $MainFormMinimizeButton.BackColor = $ColorData.$Category.$Team.BackColor
    $MainFormMinimizeButton.ForeColor = $ColorData.$Category.$Team.ForeColor
    $MainFormCloseButton.BackColor = $ColorData.$Category.$Team.BackColor
    $MainFormCloseButton.ForeColor = $ColorData.$Category.$Team.ForeColor
    $MenuStrip.BackColor = $ColorData.$Category.$Team.BackColor
    $MenuStrip.ForeColor = $ColorData.$Category.$Team.ForeColor
    $MainFormTabControl.BackColor = $ColorData.$Category.$Team.BackColor
    $SaveOutTextButton.BackColor = $ColorData.$Category.$Team.ForeColor
    $SaveOutTextButton.ForeColor = $ColorData.$Category.$Team.BackColor
    $ClearOutTextButton.BackColor = $ColorData.$Category.$Team.ForeColor
    $ClearOutTextButton.ForeColor = $ColorData.$Category.$Team.BackColor
    $SysAdminTab.BackColor = $ColorData.$Category.$Team.BackColor
    $AWSAdminTab.BackColor = $ColorData.$Category.$Team.BackColor
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
    $CertCheckWizardButton.BackColor = $ColorData.$Category.$Team.ForeColor
    $CertCheckWizardButton.ForeColor = $ColorData.$Category.$Team.BackColor
    $CertCheckLabel.ForeColor = $ColorData.$Category.$Team.ForeColor
    $script:AWSSSOLoginTimerLabel.ForeColor = $ColorData.$Category.$Team.ForeColor
    $PSTCombo.BackColor = $OutText.BackColor
    $RefreshPSTButton.BackColor = $ColorData.$Category.$Team.ForeColor
    $RefreshPSTButton.ForeColor = $ColorData.$Category.$Team.BackColor
    $LaunchLFPWizardButton.BackColor = $ColorData.$Category.$Team.ForeColor
    $LaunchLFPWizardButton.ForeColor = $ColorData.$Category.$Team.BackColor
    $LaunchLFPWizardLabel.ForeColor = $ColorData.$Category.$Team.ForeColor
    $LaunchBillingRestartButton.BackColor = $ColorData.$Category.$Team.ForeColor
    $LaunchBillingRestartButton.ForeColor = $ColorData.$Category.$Team.BackColor
    $LaunchBillingRestartLabel.ForeColor = $ColorData.$Category.$Team.ForeColor
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

    if ($script:AWSSSOLoginButton.Enabled) {
        $script:AWSSSOLoginButton.BackColor = $ColorData.$Category.$Team.ForeColor
        $script:AWSSSOLoginButton.ForeColor = $ColorData.$Category.$Team.BackColor
    } else {
        $script:AWSSSOLoginButton.BackColor = $global:DisabledBackColor
        $script:AWSSSOLoginButton.ForeColor = $global:DisabledForeColor
    }

    if ($script:ListAWSAccountsButton.Enabled) {
        $script:ListAWSAccountsButton.BackColor = $ColorData.$Category.$Team.ForeColor
        $script:ListAWSAccountsButton.ForeColor = $ColorData.$Category.$Team.BackColor
    } else {
        $script:ListAWSAccountsButton.BackColor = $global:DisabledBackColor
        $script:ListAWSAccountsButton.ForeColor = $global:DisabledForeColor
    }

    if ($script:RebootAWSInstancesButton.Enabled) {
        $script:RebootAWSInstancesButton.BackColor = $ColorData.$Category.$Team.ForeColor
        $script:RebootAWSInstancesButton.ForeColor = $ColorData.$Category.$Team.BackColor
    } else {
        $script:RebootAWSInstancesButton.BackColor = $global:DisabledBackColor
        $script:RebootAWSInstancesButton.ForeColor = $global:DisabledForeColor
    }

    if ($script:StartAWSInstancesButton.Enabled) {
        $script:StartAWSInstancesButton.BackColor = $ColorData.$Category.$Team.ForeColor
        $script:StartAWSInstancesButton.ForeColor = $ColorData.$Category.$Team.BackColor
    } else {
        $script:StartAWSInstancesButton.BackColor = $global:DisabledBackColor
        $script:StartAWSInstancesButton.ForeColor = $global:DisabledForeColor
    }

    if ($script:StopAWSInstancesButton.Enabled) {
        $script:StopAWSInstancesButton.BackColor = $ColorData.$Category.$Team.ForeColor
        $script:StopAWSInstancesButton.ForeColor = $ColorData.$Category.$Team.BackColor
    } else {
        $script:StopAWSInstancesButton.BackColor = $global:DisabledBackColor
        $script:StopAWSInstancesButton.ForeColor = $global:DisabledForeColor
    }

    if ($script:AWSScreenshotButton.Enabled) {
        $script:AWSScreenshotButton.BackColor = $ColorData.$Category.$Team.ForeColor
        $script:AWSScreenshotButton.ForeColor = $ColorData.$Category.$Team.BackColor
    } else {
        $script:AWSScreenshotButton.BackColor = $global:DisabledBackColor
        $script:AWSScreenshotButton.ForeColor = $global:DisabledForeColor
    }

    if ($script:AWSCPUMetricsButton.Enabled) {
        $script:AWSCPUMetricsButton.BackColor = $ColorData.$Category.$Team.ForeColor
        $script:AWSCPUMetricsButton.ForeColor = $ColorData.$Category.$Team.BackColor
    } else {
        $script:AWSCPUMetricsButton.BackColor = $global:DisabledBackColor
        $script:AWSCPUMetricsButton.ForeColor = $global:DisabledForeColor
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

    if ($OpenItemButton.Enabled) {
        $OpenItemButton.BackColor = $ColorData.$Category.$Team.ForeColor
        $OpenItemButton.ForeColor = $ColorData.$Category.$Team.BackColor
    } else {
        $OpenItemButton.BackColor = $global:DisabledBackColor
        $OpenItemButton.ForeColor = $global:DisabledForeColor
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

    # Updates the Font Picker menu popup controls if the popup is active
    # This is to prevent an error where the system tries to update the controls before they are created
    if ($global:IsFontPickerPopupActive) {

        $script:FontPickerPopup.BackColor = $ColorData.$Category.$Team.BackColor
        $script:FontPickerPanel.BackColor = $ColorData.$Category.$Team.BackColor

        if ($Category -eq 'Premium' -or $Category -eq 'Custom') {
            $global:CurrentTeamAccentColor = $ColorData.$Category.$Team.AccentColor
            $script:FontPickerListBox.BackColor = $global:CurrentTeamAccentColor
        }
        else {
            $script:FontPickerListBox.BackColor = [System.Drawing.SystemColors]::Control
            $global:DisabledBackColor = '#A9A9A9'
        }

        if ($script:FontPickerApplyButton.Enabled) {
            $script:FontPickerApplyButton.BackColor = $ColorData.$Category.$Team.ForeColor
            $script:FontPickerApplyButton.ForeColor = $ColorData.$Category.$Team.BackColor
        } else {
            $script:FontPickerApplyButton.BackColor = $global:DisabledBackColor
            $script:FontPickerApplyButton.ForeColor = $global:DisabledForeColor
        }
        # Reset the DrawMode to trigger a redraw. Otherwise, the listbox items will all revert to the default font
        $script:FontPickerListBox.DrawMode = 'Normal'
        $script:FontPickerListBox.DrawMode = 'OwnerDrawFixed'
        $script:FontPickerListBox.ItemHeight = 20

        # Refresh the form
        $script:FontPickerPopup.Refresh()
    }

    # Updates the Cert Check popup controls if the popup is active
    # This is to prevent an error where the system tries to update the controls before they are created
    if ($global:IsCertCheckPopupActive) {
        $script:CertCheckMinimizeButton.BackColor = $ColorData.$Category.$Team.BackColor
        $script:CertCheckMinimizeButton.ForeColor = $ColorData.$Category.$Team.ForeColor
        $script:CertCheckCloseButton.BackColor = $ColorData.$Category.$Team.BackColor
        $script:CertCheckCloseButton.ForeColor = $ColorData.$Category.$Team.ForeColor
        $script:CertCheckPopup.BackColor = $ColorData.$Category.$Team.BackColor
        $script:CertCheckPopup.ForeColor = $ColorData.$Category.$Team.ForeColor
        $script:ServerChoiceLabel.ForeColor = $ColorData.$Category.$Team.ForeColor
        $script:OutputLabel.ForeColor = $ColorData.$Category.$Team.ForeColor

        $script:CertCheckMinimizeButton.FlatAppearance.MouseOverBackColor = $global:CurrentTeamAccentColor
        $script:CertCheckMinimizeButton.Add_MouseEnter({ if ($Category -eq 'Premium' -or $Category -eq 'Custom') { $script:CertCheckMinimizeButton.ForeColor = $global:CurrentTeamForeColor } else { $script:CertCheckMinimizeButton.ForeColor = $global:CurrentTeamBackColor } })
        $script:CertCheckMinimizeButton.Add_MouseLeave({ $script:CertCheckMinimizeButton.ForeColor = $global:CurrentTeamForeColor })

        $script:CertCheckCloseButton.FlatAppearance.MouseOverBackColor = $global:CurrentTeamAccentColor
        $script:CertCheckCloseButton.Add_MouseEnter({ if ($Category -eq 'Premium' -or $Category -eq 'Custom') { $script:CertCheckCloseButton.ForeColor = $global:CurrentTeamForeColor } else { $script:CertCheckCloseButton.ForeColor = $global:CurrentTeamBackColor } })
        $script:CertCheckCloseButton.Add_MouseLeave({ $script:CertCheckCloseButton.ForeColor = $global:CurrentTeamForeColor })

        if ($Category -eq 'Premium' -or $Category -eq 'Custom') {
            $global:CurrentTeamAccentColor = $ColorData.$Category.$Team.AccentColor
            $script:CertCheckServerTextBox.BackColor = $global:CurrentTeamAccentColor
        }
        else {
            $script:CertCheckServerTextBox.BackColor = [System.Drawing.SystemColors]::Control
            $global:DisabledBackColor = '#A9A9A9'
        }

        if ($script:ChooseTextFileButton.Enabled) {
            $script:ChooseTextFileButton.BackColor = $ColorData.$Category.$Team.ForeColor
            $script:ChooseTextFileButton.ForeColor = $ColorData.$Category.$Team.BackColor
        } else {
            $script:ChooseTextFileButton.BackColor = $global:DisabledBackColor
            $script:ChooseTextFileButton.ForeColor = $global:DisabledForeColor
        }

        if ($script:RunCertCheckButton.Enabled) {
            $script:RunCertCheckButton.BackColor = $ColorData.$Category.$Team.ForeColor
            $script:RunCertCheckButton.ForeColor = $ColorData.$Category.$Team.BackColor
        } else {
            $script:RunCertCheckButton.BackColor = $global:DisabledBackColor
            $script:RunCertCheckButton.ForeColor = $global:DisabledForeColor
        }

        $script:CertCheckPopup.Refresh()
    }

    # Updates the Lender LFP popup controls if the popup is active
    # This is to prevent an error where the system tries to update the controls before they are created
    if ($global:IsLenderLFPPopupActive) {
        $script:LenderLFPPopup.BackColor = $ColorData.$Category.$Team.BackColor
        $script:LenderLFPPopup.ForeColor = $ColorData.$Category.$Team.ForeColor
        $script:LenderLFPMinimizeButton.BackColor = $ColorData.$Category.$Team.BackColor
        $script:LenderLFPMinimizeButton.ForeColor = $ColorData.$Category.$Team.ForeColor
        $script:LenderLFPCloseButton.BackColor = $ColorData.$Category.$Team.BackColor
        $script:LenderLFPCloseButton.ForeColor = $ColorData.$Category.$Team.ForeColor
        $script:LenderLFPComboLabel.ForeColor = $ColorData.$Category.$Team.ForeColor
        $script:LenderLFPTextBoxLabel.ForeColor = $ColorData.$Category.$Team.ForeColor
        $script:LenderLFPTicketTextBoxLabel.ForeColor = $ColorData.$Category.$Team.ForeColor

        $script:LenderLFPMinimizeButton.FlatAppearance.MouseOverBackColor = if ($Category -eq 'Premium' -or $Category -eq 'Custom') { $global:CurrentTeamAccentColor } else { $global:CurrentTeamForeColor }
        $script:LenderLFPMinimizeButton.Add_MouseEnter({ if ($Category -eq 'Premium' -or $Category -eq 'Custom') { $script:LenderLFPMinimizeButton.ForeColor = $global:CurrentTeamForeColor } else { $script:LenderLFPMinimizeButton.ForeColor = $global:CurrentTeamBackColor } })
        $script:LenderLFPMinimizeButton.Add_MouseLeave({ $script:LenderLFPMinimizeButton.ForeColor = $global:CurrentTeamForeColor })

        $script:LenderLFPCloseButton.FlatAppearance.MouseOverBackColor = if ($Category -eq 'Premium' -or $Category -eq 'Custom') { $global:CurrentTeamAccentColor } else { $global:CurrentTeamForeColor }
        $script:LenderLFPCloseButton.Add_MouseEnter({ if ($Category -eq 'Premium' -or $Category -eq 'Custom') { $script:LenderLFPCloseButton.ForeColor = $global:CurrentTeamForeColor } else { $script:LenderLFPCloseButton.ForeColor = $global:CurrentTeamBackColor } })
        $script:LenderLFPCloseButton.Add_MouseLeave({ $script:LenderLFPCloseButton.ForeColor = $global:CurrentTeamForeColor })

        if ($Category -eq 'Premium' -or $Category -eq 'Custom') {
            $global:CurrentTeamAccentColor = $ColorData.$Category.$Team.AccentColor
            $script:LenderLFPCombo.BackColor = $global:CurrentTeamAccentColor
            $script:LenderLFPIdTextBox.BackColor = $global:CurrentTeamAccentColor
            $script:LenderLFPTicketTextBox.BackColor = $global:CurrentTeamAccentColor
        }
        else {
            $script:LenderLFPCombo.BackColor = [System.Drawing.SystemColors]::Control
            $script:LenderLFPIdTextBox.BackColor = [System.Drawing.SystemColors]::Control
            $script:LenderLFPTicketTextBox.BackColor = [System.Drawing.SystemColors]::Control
            $global:DisabledBackColor = '#A9A9A9'
        }

        if ($script:AddLenderLFPButton.Enabled) {
            $script:AddLenderLFPButton.BackColor = $ColorData.$Category.$Team.ForeColor
            $script:AddLenderLFPButton.ForeColor = $ColorData.$Category.$Team.BackColor
        } else {
            $script:AddLenderLFPButton.BackColor = $global:DisabledBackColor
            $script:AddLenderLFPButton.ForeColor = $global:DisabledForeColor
        }

        $script:LenderLFPPopup.Refresh()
    }

    # Update the Billing Restart popup controls if the popup is active
    # This is to prevent an error where the system tries to update the controls before they are created
    if ($global:IsBillingRestartPopupActive) {
        $script:BillingRestartPopup.BackColor = $ColorData.$Category.$Team.BackColor
        $script:BillingRestartPopup.ForeColor = $ColorData.$Category.$Team.ForeColor
        $script:BillingRestartMinimizeButton.BackColor = $ColorData.$Category.$Team.BackColor
        $script:BillingRestartMinimizeButton.ForeColor = $ColorData.$Category.$Team.ForeColor
        $script:BillingRestartCloseButton.BackColor = $ColorData.$Category.$Team.BackColor
        $script:BillingRestartCloseButton.ForeColor = $ColorData.$Category.$Team.ForeColor
        $script:BillingRestartComboLabel.ForeColor = $ColorData.$Category.$Team.ForeColor

        $script:BillingRestartMinimizeButton.FlatAppearance.MouseOverBackColor = if ($Category -eq 'Premium' -or $Category -eq 'Custom') { $global:CurrentTeamAccentColor } else { $global:CurrentTeamForeColor }
        $script:BillingRestartMinimizeButton.Add_MouseEnter({ if ($Category -eq 'Premium' -or $Category -eq 'Custom') { $script:BillingRestartMinimizeButton.ForeColor = $global:CurrentTeamForeColor } else { $script:BillingRestartMinimizeButton.ForeColor = $global:CurrentTeamBackColor } })
        $script:BillingRestartMinimizeButton.Add_MouseLeave({ $script:BillingRestartMinimizeButton.ForeColor = $global:CurrentTeamForeColor })

        $script:BillingRestartCloseButton.FlatAppearance.MouseOverBackColor = if ($Category -eq 'Premium' -or $Category -eq 'Custom') { $global:CurrentTeamAccentColor } else { $global:CurrentTeamForeColor }
        $script:BillingRestartCloseButton.Add_MouseEnter({ if ($Category -eq 'Premium' -or $Category -eq 'Custom') { $script:BillingRestartCloseButton.ForeColor = $global:CurrentTeamForeColor } else { $script:BillingRestartCloseButton.ForeColor = $global:CurrentTeamBackColor } })
        $script:BillingRestartCloseButton.Add_MouseLeave({ $script:BillingRestartCloseButton.ForeColor = $global:CurrentTeamForeColor })

        if ($Category -eq 'Premium' -or $Category -eq 'Custom') {
            $global:CurrentTeamAccentColor = $ColorData.$Category.$Team.AccentColor
            $script:BillingRestartCombo.BackColor = $global:CurrentTeamAccentColor
        }
        else {
            $script:BillingRestartCombo.BackColor = [System.Drawing.SystemColors]::Control
            $global:DisabledBackColor = '#A9A9A9'
        }

        if ($global:BillingRestartButton.Enabled) {
            $script:BillingRestartButton.BackColor = $ColorData.$Category.$Team.ForeColor
            $script:BillingRestartButton.ForeColor = $ColorData.$Category.$Team.BackColor
        } else {
            $script:BillingRestartButton.BackColor = $global:DisabledBackColor
            $script:BillingRestartButton.ForeColor = $global:DisabledForeColor
        }

        $script:BillingRestartPopup.Refresh()
    }

    # Updates the HDTStorage popup controls if the popup is active
    # This is to prevent an error where the system tries to update the controls before they are created
    if ($global:IsHDTStoragePopupActive) {
        $script:HDTStoragePopup.BackColor = $ColorData.$Category.$Team.BackColor
        $script:HDTStoragePopup.ForeColor = $ColorData.$Category.$Team.ForeColor
        $script:HDTStorageMinimizeButton.BackColor = $ColorData.$Category.$Team.BackColor
        $script:HDTStorageMinimizeButton.ForeColor = $ColorData.$Category.$Team.ForeColor
        $script:HDTStorageCloseButton.BackColor = $ColorData.$Category.$Team.BackColor
        $script:HDTStorageCloseButton.ForeColor = $ColorData.$Category.$Team.ForeColor
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

        $script:HDTStorageMinimizeButton.FlatAppearance.MouseOverBackColor = if ($Category -eq 'Premium' -or $Category -eq 'Custom') { $global:CurrentTeamAccentColor } else { $global:CurrentTeamForeColor }
        $script:HDTStorageMinimizeButton.Add_MouseEnter({ if ($Category -eq 'Premium' -or $Category -eq 'Custom') { $script:HDTStorageMinimizeButton.ForeColor = $global:CurrentTeamForeColor } else { $script:HDTStorageMinimizeButton.ForeColor = $global:CurrentTeamBackColor } })
        $script:HDTStorageMinimizeButton.Add_MouseLeave({ $script:HDTStorageMinimizeButton.ForeColor = $global:CurrentTeamForeColor })

        $script:HDTStorageCloseButton.FlatAppearance.MouseOverBackColor = if ($Category -eq 'Premium' -or $Category -eq 'Custom') { $global:CurrentTeamAccentColor } else { $global:CurrentTeamForeColor }
        $script:HDTStorageCloseButton.Add_MouseEnter({ if ($Category -eq 'Premium' -or $Category -eq 'Custom') { $script:HDTStorageCloseButton.ForeColor = $global:CurrentTeamForeColor } else { $script:HDTStorageCloseButton.ForeColor = $global:CurrentTeamBackColor } })
        $script:HDTStorageCloseButton.Add_MouseLeave({ $script:HDTStorageCloseButton.ForeColor = $global:CurrentTeamForeColor })

        if ($Category -eq 'Premium' -or $Category -eq 'Custom') {
            $global:CurrentTeamAccentColor = $ColorData.$Category.$Team.AccentColor
            $script:DBServerTextBox.BackColor = $ColorData.$Category.$Team.AccentColor
            $script:TableNameTextBox.BackColor = $ColorData.$Category.$Team.AccentColor
            $script:SecurePasswordTextBox.BackColor = $ColorData.$Category.$Team.AccentColor
        }
        else {
            $script:DBServerTextBox.BackColor = [System.Drawing.SystemColors]::Control
            $script:TableNameTextBox.BackColor = [System.Drawing.SystemColors]::Control
            $script:SecurePasswordTextBox.BackColor = [System.Drawing.SystemColors]::Control
            $global:DisabledBackColor = '#A9A9A9'
        }

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
    if ($global:IsFeedbackPopupActive) {
        $script:FeedbackPopup.BackColor = $ColorData.$Category.$Team.BackColor
        $script:FeedbackPopup.ForeColor = $ColorData.$Category.$Team.ForeColor
        $script:FeedbackPopupMinimizeButton.BackColor = $ColorData.$Category.$Team.BackColor
        $script:FeedbackPopupMinimizeButton.ForeColor = $ColorData.$Category.$Team.ForeColor
        $script:FeedbackPopupCloseButton.BackColor = $ColorData.$Category.$Team.BackColor
        $script:FeedbackPopupCloseButton.ForeColor = $ColorData.$Category.$Team.ForeColor
        $script:UserNameLabel.BackColor = $ColorData.$Category.$Team.BackColor
        $script:UserNameLabel.ForeColor = $ColorData.$Category.$Team.ForeColor
        $script:FeedbackLabel.BackColor = $ColorData.$Category.$Team.BackColor
        $script:FeedbackLabel.ForeColor = $ColorData.$Category.$Team.ForeColor

        $script:FeedbackPopupMinimizeButton.FlatAppearance.MouseOverBackColor = if ($Category -eq 'Premium' -or $Category -eq 'Custom') { $global:CurrentTeamAccentColor } else { $global:CurrentTeamForeColor }
        $script:FeedbackPopupMinimizeButton.Add_MouseEnter({ if ($Category -eq 'Premium' -or $Category -eq 'Custom') { $script:FeedbackPopupMinimizeButton.ForeColor = $global:CurrentTeamForeColor } else { $script:FeedbackPopupMinimizeButton.ForeColor = $global:CurrentTeamBackColor } })
        $script:FeedbackPopupMinimizeButton.Add_MouseLeave({ $script:FeedbackPopupMinimizeButton.ForeColor = $global:CurrentTeamForeColor })

        $script:FeedbackPopupCloseButton.FlatAppearance.MouseOverBackColor = if ($Category -eq 'Premium' -or $Category -eq 'Custom') { $global:CurrentTeamAccentColor } else { $global:CurrentTeamForeColor }
        $script:FeedbackPopupCloseButton.Add_MouseEnter({ if ($Category -eq 'Premium' -or $Category -eq 'Custom') { $script:FeedbackPopupCloseButton.ForeColor = $global:CurrentTeamForeColor } else { $script:FeedbackPopupCloseButton.ForeColor = $global:CurrentTeamBackColor } })
        $script:FeedbackPopupCloseButton.Add_MouseLeave({ $script:FeedbackPopupCloseButton.ForeColor = $global:CurrentTeamForeColor })

        if ($Category -eq 'Premium' -or $Category -eq 'Custom') {
            $global:CurrentTeamAccentColor = $ColorData.$Category.$Team.AccentColor
            $script:UserNameTextBox.BackColor = $ColorData.$Category.$Team.AccentColor
            $script:FeedbackTextBox.BackColor = $ColorData.$Category.$Team.AccentColor
        }
        else {
            $script:UserNameTextBox.BackColor = [System.Drawing.SystemColors]::Control
            $script:FeedbackTextBox.BackColor = [System.Drawing.SystemColors]::Control
            $global:DisabledBackColor = '#A9A9A9'
        }

        if ($script:SubmitFeedbackButton.Enabled) {
            $script:SubmitFeedbackButton.BackColor = $ColorData.$Category.$Team.ForeColor
            $script:SubmitFeedbackButton.ForeColor = $ColorData.$Category.$Team.BackColor
        } else {
            $script:SubmitFeedbackButton.BackColor = $global:DisabledBackColor
            $script:SubmitFeedbackButton.ForeColor = $global:DisabledForeColor
        }

        $script:FeedbackPopup.Refresh()
    }

    # Update the AWS CPU Monitoring graph popup controls if the popup is active
    # This is to prevent an error where the system tries to update the controls before they are created
    if ($null -ne $AWSCPUMetricsForm -and $AWSCPUMetricsForm.Visible -eq $true) {
        $AWSCPUMetricsChart.BackColor = $global:CurrentTeamBackColor
        $AWSCPUMetricsChartArea.AxisX.TitleForeColor = $global:CurrentTeamAccentColor
        $AWSCPUMetricsChartArea.AxisX.LabelStyle.ForeColor = $global:CurrentTeamAccentColor
        $AWSCPUMetricsChartArea.AxisX.LineColor = $global:CurrentTeamAccentColor
        $AWSCPUMetricsChartArea.AxisY.TitleForeColor = $global:CurrentTeamAccentColor
        $AWSCPUMetricsChartArea.AxisY.LabelStyle.ForeColor = $global:CurrentTeamAccentColor
        $AWSCPUMetricsChartArea.AxisY.LineColor = $global:CurrentTeamAccentColor
        $AWSCPUMetricsChartArea.BackColor = $global:CurrentTeamBackColor
        $series.Color = $global:CurrentTeamForeColor
    }

    $DesktopAssistantForm.Refresh()

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
        if ($null -ne $script:FeedbackTextBox.Text -and $script:FeedbackTextBox.Text -ne '' -and
            $null -ne $script:UserNameTextBox.Text -and $script:UserNameTextBox.Text -ne '') {
            $script:SubmitFeedbackButton.Enabled = $true
        } else {
            $script:SubmitFeedbackButton.Enabled = $false
        }
    } else {
        if ($null -ne $script:FeedbackTextBox.Text -and $script:FeedbackTextBox.Text -ne '') {
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

# Function for setting the default colors for Theme Builder
function Set-ThemeBuilderDefaultColors {
    $script:ThemeBuilderBackColorTextBox.Text = ''
    $script:ThemeBuilderForeColorTextBox.Text = ''
    $script:ThemeBuilderAccentColorTextBox.Text = ''
    $script:ThemeBuilderDisabledColorTextBox.Text = ''
    $script:ThemeBuilderApplyThemeButton.Enabled = $false
    $script:ThemeBuilderSaveThemeButton.Enabled = $false
    $script:ThemeBuilderResetThemeButton.Enabled = $false
    $script:ThemeBuilderMinimizeButton.BackColor = [System.Drawing.SystemColors]::Control
    $script:ThemeBuilderMinimizeButton.ForeColor = [System.Drawing.SystemColors]::ControlText
    $script:ThemeBuilderCloseButton.BackColor = [System.Drawing.SystemColors]::Control
    $script:ThemeBuilderCloseButton.ForeColor = [System.Drawing.SystemColors]::ControlText
    $script:ThemeBuilderMinimizeButton.FlatAppearance.MouseOverBackColor = $ControlColorText
    $script:ThemeBuilderMinimizeButton.Add_MouseEnter({ $script:ThemeBuilderMinimizeButton.ForeColor = $ControlColor })
    $script:ThemeBuilderMinimizeButton.Add_MouseLeave({ $script:ThemeBuilderMinimizeButton.ForeColor = $themeColors.ForeColor })
	$script:ThemeBuilderCloseButton.FlatAppearance.MouseOverBackColor = $ControlColorText
    $script:ThemeBuilderCloseButton.Add_MouseEnter({ $script:ThemeBuilderCloseButton.ForeColor = $ControlColor })
    $script:ThemeBuilderCloseButton.Add_MouseLeave({ $script:ThemeBuilderCloseButton.ForeColor = $ControlColorText })
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
    $script:ThemeBuilderColorPickerButton.BackColor = [System.Drawing.SystemColors]::Control
    $script:ThemeBuilderColorPickerButton.ForeColor = [System.Drawing.SystemColors]::ControlText
    $script:ThemeBuilderBackColorTextBox.BackColor = 'White'
    $script:ThemeBuilderForeColorTextBox.BackColor = 'White'
    $script:ThemeBuilderAccentColorTextBox.BackColor = 'White'
    $script:ThemeBuilderDisabledColorTextBox.BackColor = 'White'
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
        [hashtable]$synchash,
        [scriptblock]$TimestampFunction
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
            [hashtable]$synchash,
            [scriptblock]$TimestampFunction
        )

        $SelectedTab = $RestartsTabControl.SelectedTab.Text
        $synchash.SelectedTab = $SelectedTab

        $SelectedServer = $ServersListBox.SelectedItem
        if ($null -eq $SelectedServer) {
            $OutText.AppendText("$($TimestampFunction.Invoke()) - DEBUGGING: No server selected.`r`n")
            return
        }

        switch ($SelectedTab) {
            "Services" {
                $OutText.AppendText("$($TimestampFunction.Invoke()) - Retrieving services on $SelectedServer...`r`n")
                $ServicesListBox.Items.Clear()
                $Services = Invoke-Command -ComputerName $SelectedServer -ScriptBlock {
                    Get-Service | ForEach-Object { $_.DisplayName } | Sort-Object
                } -Authentication Negotiate
                foreach ($service in $Services) {
                    [void]$ServicesListBox.Items.Add($service)
                }
            }
            "IIS Sites" {
                $OutText.AppendText("$($TimestampFunction.Invoke()) - Retrieving sites on $SelectedServer...`r`n")
                $IISSitesListBox.Items.Clear()
                $Sites = Invoke-Command -ComputerName $SelectedServer -ScriptBlock {
                    Import-Module WebAdministration
                    Get-Website | ForEach-Object { $_.Name } | Sort-Object
                } -Authentication Negotiate
                foreach ($site in $Sites) {
                    [void]$IISSitesListBox.Items.Add($site)
                }
            }
            "App Pools" {
                $OutText.AppendText("$($TimestampFunction.Invoke()) - Retrieving AppPools on $SelectedServer...`r`n")
                $AppPoolsListBox.Items.Clear()
                $AppPools = Invoke-Command -ComputerName $SelectedServer -ScriptBlock {
                    Import-Module WebAdministration
                    Get-IISAppPool | ForEach-Object { $_.Name } | Sort-Object
                } -Authentication Negotiate
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
        TimestampFunction = $TimestampFunction
    })
    
    $psCmd.Runspace = $runspace
    $null = $psCmd.BeginInvoke()
}

# Function to handle the event when a server is selected in the $ServersListBox
function OnServerSelected {
    $SelectedServer = $ServersListBox.SelectedItem
    if ($null -ne $SelectedServer) {
        # Call the Open PopulateListBoxRunspace function passing the selected server
        Open-PopulateListBoxRunspace -OutText $OutText -RestartsTabControl $RestartsTabControl -ServersListBox $ServersListBox -ServicesListBox $ServicesListBox -IISSitesListBox $IISSitesListBox -AppPoolsListBox $AppPoolsListBox -TimestampFunction ${function:Get-Timestamp}
    }
}

# Function to show service status when a server is selected
function OnServiceSelected {
    param(
        [scriptblock]$TimestampFunction
    )
    $SelectedServer = $ServersListBox.SelectedItem
    if ($null -ne $SelectedServer) {
        $SelectedService = $ServicesListBox.SelectedItem
        if ($null -ne $SelectedService) {
            $psCmd = [PowerShell]::Create().AddScript({
                param (
                    [string]$SelectedServer,
                    [string]$SelectedService,
                    [System.Windows.Forms.TextBox]$OutText,
                    [scriptblock]$TimestampFunction
                )

                try {
                    $ServiceStatus = Invoke-Command -ComputerName $SelectedServer -ScriptBlock {
                        Get-Service -DisplayName $using:SelectedService
                    } -Authentication Negotiate
                    $OutText.AppendText("$($TimestampFunction.Invoke()) - $SelectedService status: $($ServiceStatus.Status)`r`n")
                } catch {
                    $OutText.AppendText("$($TimestampFunction.Invoke()) - Error retrieving service status for ${SelectedService}: $($_.Exception.Message)`r`n")
                }
            }).AddParameters(@{SelectedServer = $SelectedServer; SelectedService = $SelectedService; OutText = $OutText; TimestampFunction = $TimestampFunction})

            $runspace = [RunspaceFactory]::CreateRunspace()
            $psCmd.Runspace = $runspace

            try {
                $runspace.Open()
                $OutText.AppendText("$($TimestampFunction.Invoke()) - Retrieving $SelectedService status on $SelectedServer...`r`n")

                $psCmd.BeginInvoke()
            } catch {
                $OutText.AppendText("$($TimestampFunction.Invoke()) - An error occurred while invoking the command: $($_.Exception.Message)`r`n")
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
    param(
        [scriptblock]$TimestampFunction
        )
    $SelectedServer = $ServersListBox.SelectedItem
    if ($null -ne $SelectedServer) {
        $SelectedIISSite = $IISsitesListBox.SelectedItem
        if ($null -ne $SelectedIISSite) {
            $psCmd = [PowerShell]::Create().AddScript({
                param (
                    [string]$SelectedServer,
                    [string]$SelectedIISSite,
                    [System.Windows.Forms.TextBox]$OutText,
                    [scriptblock]$TimestampFunction
                )

                try {
                    $IISSiteStatus = Invoke-Command -ComputerName $SelectedServer -ScriptBlock {
                        Import-Module WebAdministration
                        Get-Website -Name $using:SelectedIISSite
                    } -Authentication Negotiate
                    $OutText.AppendText("$($TimestampFunction.Invoke()) - $SelectedIISSite status: $($IISSiteStatus.State)`r`n")
                } catch {
                    $OutText.AppendText("$($TimestampFunction.Invoke()) - Error retrieving IIS site status for ${SelectedIISSite}: $($_.Exception.Message)`r`n")
                }
            }).AddParameters(@{SelectedServer = $SelectedServer; SelectedIISSite = $SelectedIISSite; OutText = $OutText; TimestampFunction = $TimestampFunction})

            $runspace = [RunspaceFactory]::CreateRunspace()
            $psCmd.Runspace = $runspace

            try {
                $runspace.Open()
                $OutText.AppendText("$($TimestampFunction.Invoke()) - Retrieving $SelectedIISSite status on $SelectedServer...`r`n")

                $psCmd.BeginInvoke()
            } catch {
                $OutText.AppendText("$($TimestampFunction.Invoke()) - An error occurred while invoking the command: $($_.Exception.Message)`r`n")
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
    param(
        [scriptblock]$TimestampFunction
    )
    $SelectedServer = $ServersListBox.SelectedItem
    if ($null -ne $SelectedServer) {
        $SelectedAppPool = $AppPoolsListBox.SelectedItem
        if ($null -ne $SelectedAppPool) {
            $psCmd = [PowerShell]::Create().AddScript({
                param (
                    [string]$SelectedServer,
                    [string]$SelectedAppPool,
                    [System.Windows.Forms.TextBox]$OutText,
                    [scriptblock]$TimestampFunction
                )

                try {
                    $AppPoolStatus = Invoke-Command -ComputerName $SelectedServer -ScriptBlock {
                        Import-Module WebAdministration
                        Get-WebAppPoolState -Name $using:SelectedAppPool
                    } -Authentication Negotiate
                    $OutText.AppendText("$($TimestampFunction.Invoke()) - ${SelectedAppPool} status: $($AppPoolStatus.Value)`r`n")
                } catch {
                    $OutText.AppendText("$($TimestampFunction.Invoke()) - Error retrieving AppPool status for ${SelectedAppPool}: $($_.Exception.Message)`r`n")
                }
            }).AddParameters(@{SelectedServer = $SelectedServer; SelectedAppPool = $SelectedAppPool; OutText = $OutText; TimestampFunction = $TimestampFunction})

            $runspace = [RunspaceFactory]::CreateRunspace()
            $psCmd.Runspace = $runspace

            try {
                $runspace.Open()
                $OutText.AppendText("$($TimestampFunction.Invoke()) - Retrieving ${SelectedAppPool} status on $SelectedServer...`r`n")

                $psCmd.BeginInvoke()
            } catch {
                $OutText.AppendText("$($TimestampFunction.Invoke()) - An error occurred while invoking the command: $($_.Exception.Message)`r`n")
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

# Function to open an asynchronous runspace and restart one or more services/iis sites/apppools on a server
function Open-RestartItemsRunspace {
    param (
		[string]$ConfigValuesRestartItemsScript,
        [string]$Action,
		[string]$SelectedServer,
		[string]$SelectedTab,
        [array]$SelectedItems,
        [System.Windows.Forms.TextBox]$OutText,
        [scriptblock]$TimestampFunction
    )

    $RestartItemsScript = Join-Path -Path $PSScriptRoot -ChildPath $ConfigValues.RestartItemsScript

    # Create a runspace to execute the script in a separate thread and keep the main GUI responsive
    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.Open()

    $psCmd = [PowerShell]::Create().AddScript({
        param (
            [string]$RestartItemsScript,
            [string]$Action,
            [string]$SelectedServer,
            [string]$SelectedTab,
            [array]$SelectedItems,
            [System.Windows.Forms.TextBox]$OutText,
            [scriptblock]$TimestampFunction
        )
        try {
            & $RestartItemsScript -Action $Action -SelectedServer $SelectedServer -SelectedTab $SelectedTab -SelectedItems $SelectedItems -OutTextControl $OutText -TimestampFunction $TimestampFunction
        } catch {
            $OutText.AppendText("$($TimestampFunction.Invoke()) - An unhandled exception occurred: $($_.Exception.Message)`r`n")
        }
    }).AddParameters(@{
		RestartItemsScript = $RestartItemsScript
        Action = $Action
		SelectedServer = $SelectedServer
		SelectedTab = $SelectedTab
        SelectedItems = $SelectedItems
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

# Function to open an asynchronous runspace and restart IIS on a server
function Open-RestartIISRunspace {
    param (
        [string]$SelectedServer,
        [System.Windows.Forms.TextBox]$OutText,
        [scriptblock]$TimestampFunction
    )

    $scriptblock = {
        param($SelectedServer, $TimestampFunction, $OutText)

        # Function to update UI from within the runspace
        $updateUI = {
            param($message)
                $OutText.AppendText("$($TimestampFunction.Invoke()) - $message`r`n")
        }

        try {
            & $updateUI "Restarting IIS on $SelectedServer..."
            
            $result = Invoke-Command -ComputerName $SelectedServer -ScriptBlock {
                param($ServerName)
                try {
                    # Check if IIS service exists and is running
                    $iisService = Get-Service -Name W3SVC -ErrorAction SilentlyContinue
                    if ($null -eq $iisService) {
                        return "IIS is not installed on $ServerName."
                    }
            
                    iisreset /restart > $null
                    return "Restarted IIS on $ServerName successfully."
                } catch {
                    return "An error occurred: $($_.Exception.Message)"
                }
            } -Authentication Negotiate -ArgumentList $SelectedServer
            
            & $updateUI $result
        } catch {
            & $updateUI "An error occurred: $($_.Exception.Message)"
        }
    }

    $runspace = [powershell]::Create()
    $runspace.AddScript($scriptblock).AddArgument($SelectedServer).AddArgument($TimestampFunction).AddArgument($OutText)
    $runspace.BeginInvoke()

    Register-ObjectEvent -InputObject $runspace -EventName InvocationStateChanged -Action {
        $Sender.Dispose()
        $Event.SourceEventArgs.Runspace.Close()
        $Event.SourceEventArgs.Runspace.Dispose()
    }
}

# Function to open an asynchronous runspace and restart IIS on a server
function Open-StartIISRunspace {
    param (
        [string]$SelectedServer,
        [System.Windows.Forms.TextBox]$OutText,
        [scriptblock]$TimestampFunction
    )

    $scriptblock = {
        param($SelectedServer, $TimestampFunction, $OutText)

        # Function to update UI from within the runspace
        $updateUI = {
            param($message)
                $OutText.AppendText("$($TimestampFunction.Invoke()) - $message`r`n")
        }

        try {
            & $updateUI "Starting IIS on $SelectedServer..."
            
            $result = Invoke-Command -ComputerName $SelectedServer -ScriptBlock {
                param($ServerName)
                try {
                    # Check if IIS service exists and is running
                    $iisService = Get-Service -Name W3SVC -ErrorAction SilentlyContinue
                    if ($null -eq $iisService) {
                        return "IIS is not installed on $ServerName."
                    }
            
                    iisreset /start > $null
                    return "Started IIS on $ServerName successfully."
                } catch {
                    return "An error occurred: $($_.Exception.Message)"
                }
            } -Authentication Negotiate -ArgumentList $SelectedServer
            
            & $updateUI $result
        } catch {
            & $updateUI "An error occurred: $($_.Exception.Message)"
        }
    }

    $runspace = [powershell]::Create()
    $runspace.AddScript($scriptblock).AddArgument($SelectedServer).AddArgument($TimestampFunction).AddArgument($OutText)
    $runspace.BeginInvoke()

    Register-ObjectEvent -InputObject $runspace -EventName InvocationStateChanged -Action {
        $Sender.Dispose()
        $Event.SourceEventArgs.Runspace.Close()
        $Event.SourceEventArgs.Runspace.Dispose()
    }
}

# Function to open an asynchronous runspace and restart IIS on a server
function Open-StopIISRunspace {
    param (
        [string]$SelectedServer,
        [System.Windows.Forms.TextBox]$OutText,
        [scriptblock]$TimestampFunction
    )

    $scriptblock = {
        param($SelectedServer, $TimestampFunction, $OutText)

        # Function to update UI from within the runspace
        $updateUI = {
            param($message)
                $OutText.AppendText("$($TimestampFunction.Invoke()) - $message`r`n")
        }

        try {
            & $updateUI "Stopping IIS on $SelectedServer..."
            
            $result = Invoke-Command -ComputerName $SelectedServer -ScriptBlock {
                param($ServerName)
                try {
                    # Check if IIS service exists and is running
                    $iisService = Get-Service -Name W3SVC -ErrorAction SilentlyContinue
                    if ($null -eq $iisService) {
                        return "IIS is not installed on $ServerName."
                    }
            
                    iisreset /stop > $null
                    return "Stopped IIS on $ServerName successfully."
                } catch {
                    return "An error occurred: $($_.Exception.Message)"
                }
            } -Authentication Negotiate -ArgumentList $SelectedServer
            
            & $updateUI $result
        } catch {
            & $updateUI "An error occurred: $($_.Exception.Message)"
        }
    }

    $runspace = [powershell]::Create()
    $runspace.AddScript($scriptblock).AddArgument($SelectedServer).AddArgument($TimestampFunction).AddArgument($OutText)
    $runspace.BeginInvoke()

    Register-ObjectEvent -InputObject $runspace -EventName InvocationStateChanged -Action {
        $Sender.Dispose()
        $Event.SourceEventArgs.Runspace.Close()
        $Event.SourceEventArgs.Runspace.Dispose()
    }
}

# Function to open a runspace and run Resolve-DnsName
function Open-NSLookupRunspace {
    param (
        [System.Windows.Forms.TextBox]$OutText,
        [System.Windows.Forms.TextBox]$NSLookupTextBox,
        [ScriptBlock]$TimestampFunction
    )

    try {
        $NSLookupRunspace = [runspacefactory]::CreateRunspace()
        $NSLookupRunspace.Open()

        $synchash = @{}
        $synchash.OutText = $OutText

        $NSLookupRunspace.SessionStateProxy.SetVariable("synchash", $synchash)

        $psCmd = [PowerShell]::Create().AddScript({
            param (
                [scriptblock]$TimestampFunction
            )

            $Pattern = "^(?=.{1,255}$)(?!-)[a-zA-Z0-9-]{1,63}(?<!-)(\.[a-zA-Z0-9-]{1,63})*$"
            $Hostnames = $args[0] -split ',' | ForEach-Object { $_.Trim() }

            foreach ($Hostname in $Hostnames) {
                if ($Hostname -match $Pattern) {
                    try {
                        $SelectedObjects = Resolve-DnsName -Name $Hostname -ErrorAction Stop
                        foreach ($SelectedObject in $SelectedObjects) {
                            if ($SelectedObject -and $SelectedObject.IPAddress) {
                                $IPAddress = $SelectedObject.IPAddress -as [ipaddress]
                                $IPAddressType = if ($IPAddress.AddressFamily -eq 'InterNetworkV6') {'IPv6'} else {'IPv4'}
                                $synchash.OutText.AppendText("$($TimestampFunction.Invoke()) - $Hostname $IPAddressType = $($SelectedObject.IPAddress)`r`n")
                            }
                        }
                    }
                    catch {
                        $synchash.OutText.AppendText("$($TimestampFunction.Invoke()) - An error occurred while resolving ${$Hostname}: $($_.Exception.Message)`r`n")
                    }
                }
                else {
                    $synchash.OutText.AppendText("$($TimestampFunction.Invoke()) - $Hostname is not a valid hostname`r`n")
                }
            }
        }).AddArgument($NSLookupTextBox.Text).AddParameters(@{
            TimestampFunction = $TimestampFunction
        })

        $psCmd.Runspace = $NSLookupRunspace

        $null = $psCmd.BeginInvoke()

    } catch {
        $OutText.AppendText("$($TimestampFunction.Invoke()) - An error occurred: $($_.Exception.Message)`r`n")
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
        [System.Windows.Forms.TextBox]$ServerPingTextBox,
        [ScriptBlock]$TimestampFunction
    )
     
    try {
        $ServerPingRunspace = [runspacefactory]::CreateRunspace()
        $ServerPingRunspace.Open()
        
        $synchash = @{}
        $synchash.OutText = $OutText
        $synchash.Servers = $ServerPingTextBox.Text -split ',' | ForEach-Object { $_.Trim() }

        $ServerPingRunspace.SessionStateProxy.SetVariable("synchash", $synchash)

        $psCmd = [PowerShell]::Create().AddScript({
            param (
            [scriptblock]$TimestampFunction
        )

            $synchash.Servers | ForEach-Object {
                try {
                    $server = $_
                    $synchash.OutText.AppendText("$($TimestampFunction.Invoke()) - Testing connection to $server...`r`n")
                    $PingResult = Test-Connection -ComputerName $server -Quiet -Count 3 -Timeout 1000

                    if ($PingResult) {
                        $synchash.OutText.AppendText("$($TimestampFunction.Invoke()) - Connection to $server successful.`r`n")
                    } else {
                        $synchash.OutText.AppendText("$($TimestampFunction.Invoke()) - Connection to $server failed.`r`n")
                    }
                } catch {
                    $synchash.OutText.AppendText("$($TimestampFunction.Invoke()) - Error pinging ${$server}: $($_.Exception.Message)`r`n")
                }
            }
        }).AddParameters(@{
            TimestampFunction = $TimestampFunction
        })

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

# Function for getting the latest AWS SSO cache file
function Get-AWSSSOCacheFile {
    param (
        [string]$AWSSSOCacheFilePath
    )
    $AWSSSOCacheFile = (Get-ChildItem -Path $AWSSSOCacheFilePath | Sort-Object LastWriteTime -Descending | Select-Object -First 1 -Property FullName).FullName
    $AWSSSOCacheJson = Get-Content -Path $AWSSSOCacheFile | Out-String
    $ExpiresAt = ($AWSSSOCacheJson | ConvertFrom-Json).ExpiresAt

    # Ensure UTC is respected in the target time
    try {
        # Parse the date using the format it appears in the file
        $script:TargetDateTime = [datetime]::ParseExact($ExpiresAt, "MM/dd/yyyy HH:mm:ss", $null)
    } catch {
        $OutText.AppendText("$(Get-Timestamp) - Error parsing date: $_`r`n")
    }

    return $script:TargetDateTime
}

# Function for retrieving the login confirmation code from the AWS SSO Login Output file
function Get-AWSSSOLoginOutput {
    param (
        [string]$AWSSSOLoginOutput,
        [System.Windows.Forms.TextBox]$OutText,
        [ScriptBlock]$TimestampFunction
    )

    $MatchPattern = "Then enter the code:\s*\r?\n\s*([A-Z0-9]{4}-[A-Z0-9]{4})"
    $StartTime = Get-Date
    $Timeout = New-Timespan -Minutes 1

    while ($true) {
        $AWSSSOLoginOutputCode = Get-Content -Path $AWSSSOLoginOutput -Raw
        if ($AWSSSOLoginOutputCode -match $MatchPattern) {
            $AWSSSOLoginOutputCode = $Matches[1]
            $OutText.AppendText("$($TimestampFunction.Invoke()) - AWS SSO Login Output Code: $AWSSSOLoginOutputCode`r`n")
            return $true  # Return true indicating the code was found and processed
        } else {
            Start-Sleep -Milliseconds 500  # Add a delay to prevent high CPU usage
        }

        # Break the loop and return false if timeout is reached
        if ((Get-Date) - $startTime -gt $timeout) {
            $OutText.AppendText("$($TimestampFunction.Invoke()) - Timeout reached while waiting for AWS SSO Login Output Code.`r`n")
            return $false
        }
    }
}

# Function for updating the AWS SSO Login Expiration timer
function Update-CountdownTimer {
    param(
        [System.Windows.Forms.Label]$script:AWSSSOLoginTimerLabel,
        [DateTime]$script:TargetDateTime,
        [System.Windows.Forms.Timer]$script:AWSSSOLoginTimer
    )

    $NowUTC = (Get-Date).ToUniversalTime()
    $TimeLeft = $script:TargetDateTime - $NowUTC

    if ($TimeLeft -le [TimeSpan]::Zero) {
        $script:AWSSSOLoginTimerLabel.Text = 'Session Expired'
        $script:AWSSSOLoginTimer.Stop()
        $script:AWSSSOLoginButton.Enabled = $true
        $script:ListAWSAccountsButton.Enabled = $false
    } else {
        $script:AWSSSOLoginTimerLabel.Text = 'Session expires in ' + ($TimeLeft.ToString("mm\:ss"))
        $script:AWSSSOLoginButton.Enabled = $false
        $script:ListAWSAccountsButton.Enabled = $true
    }

    if ($script:AWSSSOLoginButton.Enabled) {
        $backColor = Get-AppropriateColor -ColorType "BackColor"
        $foreColor = Get-AppropriateColor -ColorType "ForeColor"
    
        $script:AWSSSOLoginButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
        $script:AWSSSOLoginButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
    } else {
        $script:AWSSSOLoginButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
        $script:AWSSSOLoginButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
    }

    if ($script:ListAWSAccountsButton.Enabled) {
        $backColor = Get-AppropriateColor -ColorType "BackColor"
        $foreColor = Get-AppropriateColor -ColorType "ForeColor"
    
        $script:ListAWSAccountsButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
        $script:ListAWSAccountsButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
    } else {
        $script:ListAWSAccountsButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
        $script:ListAWSAccountsButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
    }
}

# Function to list all AWS accounts. The AWS SSO list-accounts command usually requires a re-auth, so the Access Token section is a workaround for that.
function Open-AWSSSOLoginRunspace {
    param(
        [string]$AWSSSOLoginOutput,
		[string]$AWSSSOProfile,
        [string]$AWSSSOCacheFilePath,
		[System.Windows.Forms.TextBox]$OutText,
        [scriptblock]$TimestampFunction
    )

    # Create a runspace to execute the script in a separate thread and keep the main GUI responsive
    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.Open()

    $psCmd = [PowerShell]::Create().AddScript({
        param(
            [string]$AWSSSOLoginOutput,
			[string]$AWSSSOProfile,
			[string]$AWSSSOCacheFilePath,
			[System.Windows.Forms.TextBox]$OutText,
			[scriptblock]$TimestampFunction
		)

        # Command to login to AWS using SSO
		# This will open a browser window for the user to allow access to the AWS account
		aws sso login --profile $AWSSSOProfile *> $AWSSSOLoginOutput
		
    }).AddParameters(@{
        AWSSSOLoginOutput = $AWSSSOLoginOutput
        AWSSSOProfile = $AWSSSOProfile
        AWSSSOCacheFilePath = $AWSSSOCacheFilePath
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

# Function to list all AWS accounts. The AWS SSO list-accounts command usually requires a re-auth, so the Access Token section is a workaround for that.
function Get-AWSAccounts {
    param (
        [string]$AWSSSOProfile,
        [string]$AWSAccountsFile,
        [string]$AWSAccessToken,
        [System.Windows.Forms.ListBox]$AWSAccountsListBox,
        [System.Windows.Forms.TextBox]$OutText,
        [scriptblock]$TimestampFunction
    )

    # Create a runspace to execute the script in a separate thread and keep the main GUI responsive
    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.Open()

    $psCmd = [PowerShell]::Create().AddScript({
        param (
            [string]$AWSSSOProfile,
            [string]$AWSAccountsFile,
            [string]$AWSAccessToken,
            [System.Windows.Forms.ListBox]$AWSAccountsListBox,
            [System.Windows.Forms.TextBox]$OutText,
            [scriptblock]$TimestampFunction
        )

        try {
            # Run the AWS SSO list-accounts command and capture the output in a json file
            aws sso list-accounts --profile $AWSSSOProfile --access-token $AWSAccessToken --output json > $AWSAccountsFile
        } catch {
            $OutText.AppendText("$($TimestampFunction.Invoke()) - An error occurred while listing AWS accounts: $($_.Exception.Message)`r`n")
        }
        
        try {
            $jsonString = Get-Content -Path $AWSAccountsFile -Raw
            $jsonString | ConvertFrom-Json -ErrorAction Stop | Select-Object -ExpandProperty accountList | Sort-Object -Property accountName |
            ForEach-Object {
                [void]$script:AWSAccountsListBox.Items.Add($_.accountName)
            }
        } catch {
            $OutText.AppendText("$($TimestampFunction.Invoke()) - Failed to load or parse JSON: $_`r`n")
        }

    }).AddParameters(@{
        AWSSSOProfile = $AWSSSOProfile
        AWSAccountsFile = $AWSAccountsFile
        AWSAccessToken = $AWSAccessToken
        AWSAccountsListBox = $AWSAccountsListBox
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

# Function for updating the Account ID in the AWS config file
# This is used to easily switch between accounts without needing to reauthenticate
function Update-AccountID {
    param (
        [string]$AccountId,
        [string]$AWSSSOProfile,
        [string]$AWSSSOConfigFilePath,
        [System.Windows.Forms.TextBox]$OutText,
        [scriptblock]$TimestampFunction
    )

    # Create a runspace to execute the script in a separate thread and keep the main GUI responsive
    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.Open()

    $psCmd = [PowerShell]::Create().AddScript({
        param (
            [string]$AccountId,
            [string]$AWSSSOProfile,
            [string]$AWSSSOConfigFilePath,
            [System.Windows.Forms.TextBox]$OutText,
            [scriptblock]$TimestampFunction
        )

        # Check if the config file exists
        if (Test-Path $AWSSSOConfigFilePath) {
            # Read the current configuration
            $ConfigContent = Get-Content $AWSSSOConfigFilePath -Raw

            # Regular expression to find the profile section and the sso_account_id line
            $ProfileSectionPattern = "\[profile\s+$AWSSSOProfile\](.*?)sso_account_id\s*=.*?(\r?\n)"
            $ReplacementText = "[profile $AWSSSOProfile]`$1sso_account_id = $AccountId`$2"

            # Update the sso_account_id in the profile section
            $UpdatedConfigContent = [regex]::Replace($ConfigContent, $ProfileSectionPattern, $ReplacementText, [System.Text.RegularExpressions.RegexOptions]::Singleline)

            try {
                # Write the updated configuration back to the file
                Set-Content -Path $AWSSSOConfigFilePath -Value $UpdatedConfigContent
                $OutText.AppendText("$($TimestampFunction.Invoke()) - Updated account ID to $AccountId in config file`r`n")
            } catch {
                $OutText.AppendText("$($TimestampFunction.Invoke()) - Error updating account ID in config file: $($_.Exception.Message)`r`n")
            }
        } else {
            $OutText.AppendText("$($TimestampFunction.Invoke()) - AWS config file not found.`r`n")
        }
    }).AddParameters(@{
        AccountId = $AccountId
        AWSSSOProfile = $AWSSSOProfile
        AWSSSOConfigFilePath = $AWSSSOConfigFilePath
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

# Function for running the AWS describe-instances command
function Invoke-DescribeInstances {
    param (
        [string]$AWSSSOProfile,
        [System.Windows.Forms.Listbox]$AWSInstancesListBox,
        [System.Windows.Forms.TextBox]$OutText,
        [scriptblock]$TimestampFunction
    )
    # Create a runspace to execute the script in a separate thread and keep the main GUI responsive
    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.Open()

    $psCmd = [PowerShell]::Create().AddScript({
        param (
            [string]$AWSSSOProfile,
            [System.Windows.Forms.Listbox]$AWSInstancesListBox,
            [System.Windows.Forms.TextBox]$OutText,
            [scriptblock]$TimestampFunction
        )
        try {
            # Describe the instances in the selected account
            $EC2Result = aws ec2 describe-instances --region us-east-2 --query "Reservations[*].Instances[*].Tags[?Key=='Name'].Value" --output text --profile $AWSSSOProfile

            $FormattedResult = $EC2Result -split "`r`n"
            
            # Adding items to the AWS instances ListBox
            try {

                if ($null -eq $script:AWSInstancesListBox) {
                    throw "ListBox is not initialized."
                }

                foreach ($line in $FormattedResult) {
                    if (-not [string]::IsNullOrWhiteSpace($line)) {
                        [void]$script:AWSInstancesListBox.Items.Add($line)
                    }
                }
            } catch {
                $OutText.AppendText("$($TimestampFunction.Invoke()) - Error adding instances to list box: $($_.Exception.Message)`r`n")
            }            

        } catch {
            $OutText.AppendText("$($TimestampFunction.Invoke()) - Failed to describe instances for account: $selectedAccountName`r`n")
        }
    }).AddParameters(@{
        AWSSSOProfile = $AWSSSOProfile
        AWSInstancesListBox = $AWSInstancesListBox
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

# Function for event when clicking on an account in the AWS Accounts list box
# This function is used to get the instances for the selected account
function OnAWSAccountSelected {
	param (
        [string]$AWSSSOProfile,
		[string]$AWSAccountsFile,
		[string]$AWSConfigFile,
		[System.Windows.Forms.ListBox]$AWSAccountsListBox,
        [System.Windows.Forms.Listbox]$AWSInstancesListBox,
		[System.Windows.Forms.TextBox]$OutText,
        [scriptblock]$TimestampFunction
    )

    # Clear current items from AWS instances ListBox
    $AWSInstancesListBox.Items.Clear()

    $OutText.AppendText("$(Get-Timestamp) - Describing instances for $($AWSAccountsListBox.SelectedItem)...`r`n")

	if ($null -ne $script:AWSAccountsListBox.SelectedItem) {
        $SelectedAccountName = $script:AWSAccountsListBox.SelectedItem.ToString()
		
		$jsonString = Get-Content -Path $AWSAccountsFile -Raw
        $AWSAccountsList = $jsonString | ConvertFrom-Json -ErrorAction Stop

        # Extract the accountId for the selected account from the accounts list
        $selectedAccount = $AWSAccountsList.accountList | Where-Object { $_.accountName -eq $SelectedAccountName }
        $selectedAccountId = $selectedAccount.accountId

        # Paths to AWS config files
        $AWSSSOConfigFilePath = Join-Path $env:USERPROFILE $AWSConfigFile

        Update-AccountID -accountId $selectedAccountId -AWSSSOProfile $AWSSSOProfile -AWSSSOConfigFilePath $AWSSSOConfigFilePath

        if ($selectedAccountId) {
            # Describe the instances in the selected account
            Invoke-DescribeInstances -AWSSSOProfile $AWSSSOProfile -AWSInstancesListBox $AWSInstancesListBox -OutText $OutText -TimestampFunction ${function:Get-Timestamp}
        }
    } else {
        $OutText.AppendText("$(Get-Timestamp) - Please select an account from the list.`r`n")
    }
}

# Function event when clicking on an instance in the AWS Instances list box
# This function is used to get the instance status
function OnAWSInstanceSelected {
    param (
        [string]$AWSSSOProfile,
        [System.Windows.Forms.ListBox]$AWSInstancesListBox,
        [System.Windows.Forms.TextBox]$OutText,
        [scriptblock]$TimestampFunction
    )

    $SelectedInstanceName = $AWSInstancesListBox.SelectedItem.ToString()

    $OutText.AppendText("$(Get-Timestamp) - Getting status for $SelectedInstanceName...`r`n")

    # Create a runspace to execute the script in a separate thread and keep the main GUI responsive
    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.Open()

    $psCmd = [PowerShell]::Create().AddScript({
        param (
			[string]$SelectedInstanceName,
            [string]$AWSSSOProfile,
            [System.Windows.Forms.TextBox]$OutText,
            [scriptblock]$TimestampFunction
        )
		
		try {
			$AWSInstanceStatus = aws ec2 describe-instances --region us-east-2 --filters "Name=tag:Name,Values=$SelectedInstanceName" --query "Reservations[*].Instances[*].State.Name" --profile $AWSSSOProfile --output text
			$OutText.AppendText("$($TimestampFunction.Invoke()) - $SelectedInstanceName is $AWSInstanceStatus.`r`n")
		} catch {
			$OutText.AppendText("$($TimestampFunction.Invoke()) - Error retrieving instance status: $_`r`n")
		}
		
    }).AddParameters(@{
		SelectedInstanceName = $SelectedInstanceName
        AWSSSOProfile = $AWSSSOProfile
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

# Function for running the AWS describe-instances command
function Open-RebootAWSInstancesRunspace {
    param (
		[string]$RestartAWSInstanceScript,
        [string]$AWSSSOProfile,
		[string]$Action,
		[string]$SelectedInstanceName,
        [System.Windows.Forms.TextBox]$OutText,
        [scriptblock]$TimestampFunction
    )
    # Create a runspace to execute the script in a separate thread and keep the main GUI responsive
    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.Open()

    $psCmd = [PowerShell]::Create().AddScript({
        param (
			[string]$RestartAWSInstanceScript,
			[string]$AWSSSOProfile,
			[string]$Action,
			[string]$SelectedInstanceName,
			[System.Windows.Forms.TextBox]$OutText,
			[scriptblock]$TimestampFunction
        )
		try {
			& $RestartAWSInstanceScript -AWSSSOProfile $AWSSSOProfile -Action $Action -SelectedInstanceName $SelectedInstanceName -OutTextControl $OutText -TimestampFunction $TimestampFunction
		} catch {
			$OutText.AppendText("$($TimestampFunction.Invoke()) - Error launching Restart AWS Instances script: $_`r`n")
		}
		
    }).AddParameters(@{
		RestartAWSInstanceScript = $RestartAWSInstanceScript
        AWSSSOProfile = $AWSSSOProfile
		Action = $Action
		SelectedInstanceName = $SelectedInstanceName
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

# Function logic for enabling the AWS Instance Screenshot button
function Enable-AWSGUIButtons {
    if ($null -ne $script:AWSInstancesListBox.SelectedItem) {
        $script:RebootAWSInstancesButton.Enabled = $true
        $script:StartAWSInstancesButton.Enabled = $true
        $script:StopAWSInstancesButton.Enabled = $true
        $script:AWSScreenshotButton.Enabled = $true
        $script:AWSCPUMetricsButton.Enabled = $true

        $backColor = Get-AppropriateColor -ColorType "BackColor"
        $foreColor = Get-AppropriateColor -ColorType "ForeColor"

        $script:RebootAWSInstancesButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
        $script:RebootAWSInstancesButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
        $script:StartAWSInstancesButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
        $script:StartAWSInstancesButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
        $script:StopAWSInstancesButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
        $script:StopAWSInstancesButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
        $script:AWSScreenshotButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
        $script:AWSScreenshotButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
        $script:AWSCPUMetricsButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
        $script:AWSCPUMetricsButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
    } else {
        $script:RebootAWSInstancesButton.Enabled = $false
        $script:StartAWSInstancesButton.Enabled = $false
        $script:StopAWSInstancesButton.Enabled = $false
        $script:AWSScreenshotButton.Enabled = $false

        $script:RebootAWSInstancesButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
        $script:RebootAWSInstancesButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
        $script:StartAWSInstancesButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
        $script:StartAWSInstancesButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
        $script:StopAWSInstancesButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
        $script:StopAWSInstancesButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
        $script:AWSScreenshotButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
        $script:AWSScreenshotButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
        $script:AWSCPUMetricsButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
        $script:AWSCPUMetricsButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
    }
}

# Function for running the AWS describe-instances command
function Get-InstanceScreenshot {
    param (
		[string]$SelectedInstanceName,
        [string]$AWSSSOProfile,
		[string]$InstanceScreenshotFile,
        [System.Windows.Forms.TextBox]$OutText,
        [scriptblock]$TimestampFunction
    )
    # Create a runspace to execute the script in a separate thread and keep the main GUI responsive
    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.Open()

    $psCmd = [PowerShell]::Create().AddScript({
        param (
			[string]$SelectedInstanceName,
            [string]$AWSSSOProfile,
			[string]$InstanceScreenshotFile,
            [System.Windows.Forms.TextBox]$OutText,
            [scriptblock]$TimestampFunction
        )
		
		try {
			$AWSInstanceId = aws ec2 describe-instances --region us-east-2 --filters "Name=tag:Name,Values=$SelectedInstanceName" --query "Reservations[*].Instances[*].InstanceId" --profile $AWSSSOProfile --output text
		} catch {
			$OutText.AppendText("$($TimestampFunction.Invoke()) - Error retrieving instance name: $_`r`n")
		}

		try {
			$InstanceScreenshot = aws ec2 get-console-screenshot --instance-id $AWSInstanceId --region us-east-2 --profile $AWSSSOProfile
		} catch {
			$OutText.AppendText("$($TimestampFunction.Invoke()) - Error taking snapshot of $($SelectedInstanceName): $_`r`n")
		}
		
		try {
			$ImageData = ($InstanceScreenshot | ConvertFrom-Json).ImageData
			$ImageBytes = [System.Convert]::FromBase64String($ImageData)
			[System.IO.File]::WriteAllBytes($InstanceScreenshotFile, $ImageBytes)

            # Resolve the path to the screenshot file to make the output more readable
            $InstanceScreenshotFile = Resolve-Path -Path $InstanceScreenshotFile -ErrorAction Stop

            $OutText.AppendText("$($TimestampFunction.Invoke()) - Successfully saved screenshot at $InstanceScreenshotFile.`r`n")
            $OutText.AppendText("$($TimestampFunction.Invoke()) - Opening screenshot...`r`n")
		} catch {
			$OutText.AppendText("$($TimestampFunction.Invoke()) - Error saving snapshot of $($SelectedInstanceName): $_`r`n")
		}
		
		Invoke-Item $InstanceScreenshotFile
		
    }).AddParameters(@{
		SelectedInstanceName = $SelectedInstanceName
        AWSSSOProfile = $AWSSSOProfile
		InstanceScreenshotFile = $InstanceScreenshotFile
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

# Function to open an asynchronous runspace and retrieve CPU metrics for the selected instances
function Open-GetAWSCPUMetricsRunspace {
    param (
		[string]$GetAWSCPUMetricsScript,
		[string]$AWSSSOProfile,
		[string]$SelectedInstanceName,
        [string]$StartTime,
        [string]$EndTime,
        [int]$PollingPeriod,
		[string]$BackColor,
		[string]$ForeColor,
		[string]$AccentColor,
        [System.Windows.Forms.TextBox]$OutText,
        [scriptblock]$TimestampFunction
    )

    # Create a runspace to execute the script in a separate thread and keep the main GUI responsive
    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.Open()

    $psCmd = [PowerShell]::Create().AddScript({
        param (
            [string]$GetAWSCPUMetricsScript,
			[string]$AWSSSOProfile,
			[string]$SelectedInstanceName,
            [string]$StartTime,
            [string]$EndTime,
            [int]$PollingPeriod,
			[string]$BackColor,
			[string]$ForeColor,
			[string]$AccentColor,
			[System.Windows.Forms.TextBox]$OutText,
			[scriptblock]$TimestampFunction
        )
        try {
            & $GetAWSCPUMetricsScript -AWSSSOProfile $AWSSSOProfile -SelectedInstanceName $SelectedInstanceName -StartTime $StartTime -EndTime $EndTime -PollingPeriod $PollingPeriod -BackColor $BackColor -ForeColor $ForeColor -AccentColor $AccentColor -OutTextControl $OutText -TimestampFunction $TimestampFunction
        } catch {
            $OutText.AppendText("$($TimestampFunction.Invoke()) - An unhandled exception occurred: $($_.Exception.Message)`r`n")
        }
    }).AddParameters(@{
        GetAWSCPUMetricsScript = $GetAWSCPUMetricsScript
        AWSSSOProfile = $AWSSSOProfile
        SelectedInstanceName = $SelectedInstanceName
        StartTime = $StartTime
        EndTime = $EndTime
        PollingPeriod = $PollingPeriod
        BackColor = $BackColor
		ForeColor = $ForeColor
		AccentColor = $AccentColor
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

# Function for resetting the selected environment in Prod Support Tool
Function Reset-Environment {
    $PSTPath = Join-Path -Path $ResolvedLocalSupportTool -ChildPath "ProductionSupportTool.exe.config"
    $PSTOldPath = Join-Path -Path $ResolvedLocalSupportTool -ChildPath "ProductionSupportTool.exe.config.old"
    
    if (Test-Path $PSTPath) {
        Remove-Item -Path $PSTPath
    }
    
    if (Test-Path $PSTOldPath) {
        Rename-Item $PSTOldPath -NewName $PSTPath
    }
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

# Function for opening a runspace and running Refresh/Import Prod Support Tool files
function Update-PSTFiles {
    param (
        [string]$Team,
        [string]$Category,
        [scriptblock]$TimestampFunction
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
        param($synchash, $RemoteSupportTool, $LocalSupportTool, $RemoteConfigs, $LocalConfigs, $TimestampFunction)

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
            $synchash.OutText.AppendText("$($TimestampFunction.Invoke()) - Local Configs folder does not exist. Creating folder and copying files...`r`n")
            Copy-Item -Path $RemoteConfigs* -Destination $LocalConfigs -Recurse -Force
            $synchash.OutText.AppendText("$($TimestampFunction.Invoke()) - Local Configs folder created successfully.`r`n")
        }
        elseif ((Get-Item $RemoteConfigs).LastWriteTime -gt (Get-Item $LocalConfigs).LastWriteTime) {
            $synchash.OutText.AppendText("$($TimestampFunction.Invoke()) - Local Configs folder is out of date. Updating folder...`r`n")
            Copy-Item -Path $RemoteConfigs* -Destination $LocalConfigs -Recurse -Force
            $synchash.OutText.AppendText("$($TimestampFunction.Invoke()) - Local Configs folder updated successfully.`r`n")
        }
        else {
            $synchash.OutText.AppendText("$($TimestampFunction.Invoke()) - Local Configs folder is up to date.`r`n")
        }

        if (-not (Test-Path $LocalSupportTool)) {
            $synchash.OutText.AppendText("$($TimestampFunction.Invoke()) - Local Prod Support Tool folder does not exist. Creating folder and copying files. This will take a few minutes...`r`n")
            $synchash.RefreshPSTButton.Text = "Copying PST Files..."

            Copy-Item -Path $RemoteSupportTool* -Destination $LocalSupportTool -Recurse -Force

            $synchash.RefreshPSTButton.Text = "Refresh PST Files"
            $synchash.OutText.AppendText("$($TimestampFunction.Invoke()) - Local Prod Support Tool folder created successfully. You can now use the Prod Support Tool.`r`n")
        } elseif ($RemoteSupportToolFolderDate -gt $LocalSupportToolFolderDate) {
            $synchash.OutText.AppendText("$($TimestampFunction.Invoke()) - Local Prod Support Tool folder is out of date. Updating folder...`r`n")
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
                    $synchash.OutText.AppendText("$($TimestampFunction.Invoke()) - Copied updated file: $($file.Name)`r`n")
                }
        }
        $synchash.RefreshPSTButton.Text = "Refresh PST Files"
        $synchash.OutText.AppendText("$($TimestampFunction.Invoke()) - Local Prod Support Tool folder updated successfully.`r`n")
            }
            else {
                $synchash.OutText.AppendText("$($TimestampFunction.Invoke()) - Local Prod Support Tool folder is up to date.`r`n")
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

    }).AddArgument($synchash).AddArgument($RemoteSupportTool).AddArgument($LocalSupportTool).AddArgument($RemoteConfigs).AddArgument($LocalConfigs).AddArgument($TimestampFunction)

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
}

# Function to evaluate if the run script button should be enabled
function Enable-AddLenderButton {
    $backColor = Get-AppropriateColor -ColorType "BackColor"
    $foreColor = Get-AppropriateColor -ColorType "ForeColor"
    if ($null -ne $script:LenderLFPCombo.SelectedItem) {
		if ($script:LenderLFPCombo.SelectedItem -eq 'Production'){
			if ($null -ne $script:LenderLFPCombo.SelectedItem -and $script:LenderLFPIdTextBox.Text -ne '' -and $script:LenderLFPTicketTextBox.Text -ne '' -and $script:LFPProductionCheckbox.Checked -eq $true){
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
        [scriptblock]$TimestampFunction
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
            [scriptblock]$TimestampFunction
        )

        try {
            . $AddLenderScript -LenderId $LenderId -TicketNumber $TicketNumber -Environment $Environment -OutTextControl $OutText -TimestampFunction $TimestampFunction
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

# Function to evaluate if the run script button should be enabled
function Enable-BillingServicesRestartsButton {
    $backColor = Get-AppropriateColor -ColorType "BackColor"
    $foreColor = Get-AppropriateColor -ColorType "ForeColor"
    if ($null -ne $script:BillingRestartCombo.SelectedItem) {
		if ($script:BillingRestartCombo.SelectedItem -eq 'Production'){
			if ($null -ne $script:BillingRestartCombo.SelectedItem -and $script:BillingRestartProductionCheckbox.Checked -eq $true){
				$script:BillingRestartButton.Enabled = $true
                $script:BillingRestartButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
                $script:BillingRestartButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
			}
			else {
				$script:BillingRestartButton.Enabled = $false
                $script:BillingRestartButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
                $script:BillingRestartButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
			}
		}
		elseif ($null -ne $script:BillingRestartCombo.SelectedItem){
			$script:BillingRestartButton.Enabled = $true
            $script:BillingRestartButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
            $script:BillingRestartButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
		}
	}
	else {
		$script:BillingRestartButton.Enabled = $false
        $script:BillingRestartButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
        $script:BillingRestartButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
	}
}

# Function to open an asynchronous runspace and run the BillingRestartsScript
function Open-BillingRestartsScriptRunspace {
    param (
        [string]$BillingRestartsScript,
        [string]$Environment,
        [System.Windows.Forms.TextBox]$OutText,
        [scriptblock]$TimestampFunction
    )
    
    # Create a runspace to execute the script in a separate thread and keep the main GUI responsive
    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.Open()

    $psCmd = [PowerShell]::Create().AddScript({
        param (
            [string]$BillingRestartsScript,
            [string]$Environment,
            [System.Windows.Forms.TextBox]$OutText,
            [scriptblock]$TimestampFunction
        )

        try {
            . $BillingRestartsScript -Environment $Environment -OutTextControl $OutText -TimestampFunction $TimestampFunction
        } catch {
            $OutText.AppendText("$($TimestampFunction.Invoke()) - An error occurred while executing the BillingRestartsScript: $($_.Exception.Message)`r`n")
        }
    }).AddParameters(@{
        BillingRestartsScript = $BillingRestartsScript
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
function Set-MyPassword {
    $TrimmedPassword = $PWTextBox.Text.Trim()
    $TrimmedPassword | Set-Clipboard
    $global:SecurePW = $TrimmedPassword | ConvertTo-SecureString -AsPlainText -Force
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
function Set-AltPassword {
    $TrimmedAltPassword = $PWTextBox.Text.Trim()
    $TrimmedAltPassword | Set-Clipboard
    $global:AltSecurePW = $TrimmedAltPassword | ConvertTo-SecureString -AsPlainText -Force
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

# Function to open an asynchronous runspace and create an HDDT storage table
function Open-CreateHDTStorageRunspace {
    param (
        [string]$ConfigValuesCreateHDTStorageTableScript,
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

    $CreateHDTStorageTableScript = Join-Path -Path $PSScriptRoot -ChildPath $ConfigValuesCreateHDTStorageTableScript
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
            [string]$CreateHDTStorageTableScript,
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
            & $CreateHDTStorageTableScript -File $File -DestinationFile $DestinationFile -OutPath $OutPath -LoginCredentials $LoginCredentials -Instance $SQLInstance -TableName $TableName -OutTextControl $OutText -TimestampFunction $TimestampFunction
        } catch {
            $OutText.AppendText("$($TimestampFunction.Invoke()) - An error occurred: $($_.Exception.Message)`r`n")
        }
    }).AddParameters(@{
        CreateHDTStorageTableScript = $CreateHDTStorageTableScript
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

# Function to test for invalid characters in the doc name
function Test-ValidFileName {
    param([string]$DocTopic)

    $IndexOfInvalidChar = $DocTopic.IndexOfAny([System.IO.Path]::GetInvalidFileNameChars())

    # IndexOfAny() returns the value -1 to indicate no such character was found
    return $IndexOfInvalidChar -eq -1
}

# Function to check document name length
function Test-DocLength {
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

# Function to perform a reverse IP lookup asynchronously
function Open-ReverseIPLookupRunspace {
    param (
        [string]$ip,
        [System.Windows.Forms.TextBox]$OutText,
        [scriptblock]$TimestampFunction
    )

    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.Open()
    
    # Create a synchronized hashtable
    $synchash = [hashtable]::Synchronized(@{})
    $synchash.OutText = $OutText
    $synchash.Servers = $ip -split ',' | ForEach-Object { $_.Trim() }

    # Define the script block to run in the runspace
    $psCmd = [PowerShell]::Create().AddScript({
        param (
            [System.Collections.Hashtable]$synchash,
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

        $synchash.Servers | ForEach-Object {
            $ip = $_  # Capture the current IP address

            # Validate IP Address
            $isValidIP = [System.Net.IPAddress]::TryParse($ip, [ref]$null)

            if ($isValidIP) {
                $synchash.OutText.AppendText("$($TimestampFunction.Invoke()) - Starting reverse IP lookup for $ip...`r`n")
                $hostname = Get-IPAddress -IPAddress $ip
                if ($hostname -notlike "Error:*") {
                    $synchash.OutText.AppendText("$($TimestampFunction.Invoke()) - The hostname for IP address $ip is $hostname.`r`n")
                } else {
                    # If an error occurred, $hostname contains the error message
                    $synchash.OutText.AppendText("$($TimestampFunction.Invoke()) - $hostname`r`n")
                }
            } else {
                $synchash.OutText.AppendText("$($TimestampFunction.Invoke()) - The provided IP address ($ip) is invalid. Please enter a valid IP address.`r`n")
            }
        }
    }).AddParameters(@{
        synchash = $synchash
        TimestampFunction = $TimestampFunction
    })

    # Start the script in the runspace
    $psCmd.Runspace = $runspace
    $null = $psCmd.BeginInvoke()

    # Clean up resources once the script completes
    $psCmd.add_InvocationStateChanged({
        if ($_.InvocationStateInfo.State -eq [System.Management.Automation.PSInvocationState]::Completed) {
            $psCmd.Dispose()
            $runspace.Close()
            $runspace.Dispose()
        }
    })
}

# Function to check if any radio button in a panel is selected
function IsAnyRadioButtonChecked {
    param (
        [System.Windows.Forms.Panel]$Panel
    )

    foreach ($control in $Panel.Controls) {
        if ($control -is [System.Windows.Forms.RadioButton] -and $control.Checked) {
            return $true
        }
    }
    return $false
}

# Function to evaluate if the Cert Check button should be enabled
function Enable-CertCheckObjects {
    if ($script:TextFileRadioButton.Checked) {
        $script:ChooseTextFileButton.Enabled = $true
    } else {
        $script:ChooseTextFileButton. Enabled = $false
        }
	
	if ($script:UserServersRadioButton.Checked) {
		$script:CertCheckServerTextBox.Enabled = $true
	} else {
        $script:CertCheckServerTextBox.Enabled = $false
    }

    if ((IsAnyRadioButtonChecked -Panel $script:ServerChoicePanel) -and (IsAnyRadioButtonChecked -Panel $script:OutputPanel)) {
        if ($script:ServerCSVRadioButton.Checked -and $script:OutTextRadioButton.Checked) {
            $script:RunCertCheckButton.Enabled = $true
        } elseif ($script:ServerCSVRadioButton.Checked -and $script:TextFileRadioButton.Checked -and $script:ChooseTextFileResult -eq [System.Windows.Forms.DialogResult]::OK) {
            $script:RunCertCheckButton.Enabled = $true
        } elseif ($script:UserServersRadioButton.Checked -and $script:OutTextRadioButton.Checked) {
            $script:RunCertCheckButton.Enabled = $true
        } elseif ($script:UserServersRadioButton.Checked -and $script:TextFileRadioButton.Checked -and $script:ChooseTextFileResult -eq [System.Windows.Forms.DialogResult]::OK) {
            $script:RunCertCheckButton.Enabled = $true
        } else {
            $script:RunCertCheckButton.Enabled = $false
        }
    }

    if ($script:ChooseTextFileButton.Enabled) {
        $backColor = Get-AppropriateColor -ColorType "BackColor"
        $foreColor = Get-AppropriateColor -ColorType "ForeColor"

        $script:ChooseTextFileButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
        $script:ChooseTextFileButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
    } else {
        $script:ChooseTextFileButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
        $script:ChooseTextFileButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
    }

    if ($script:RunCertCheckButton.Enabled) {
        $backColor = Get-AppropriateColor -ColorType "BackColor"
        $foreColor = Get-AppropriateColor -ColorType "ForeColor"

        $script:RunCertCheckButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
        $script:RunCertCheckButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
    } else {
        $script:RunCertCheckButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
        $script:RunCertCheckButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
    }
}

# Function to open an asynchronous runspace and submit feedback through a Teams message
function Open-CertCheckRunspace {
    param (
		[string]$ConfigValuesCertCheckScript,
        [string]$UserChosenServers,
        [string]$OutputFilePath,
        [string]$IgnoreFailedServers,
        [System.Windows.Forms.TextBox]$OutText,
        [scriptblock]$TimestampFunction
    )

    $CertCheckScript = Join-Path -Path $PSScriptRoot -ChildPath $ConfigValues.CertCheckScript

    # Create a runspace to execute the script in a separate thread and keep the main GUI responsive
    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.Open()

    $psCmd = [PowerShell]::Create().AddScript({
        param (
            [string]$CertCheckScript,
            [string]$UserChosenServers,
            [string]$OutputFilePath,
            [string]$IgnoreFailedServers,
            [System.Windows.Forms.TextBox]$OutText,
            [scriptblock]$TimestampFunction
        )
        try {
            & $CertCheckScript -UserChosenServers $UserChosenServers -OutputFilePath $OutputFilePath -IgnoreFailedServers $IgnoreFailedServers -OutTextControl $OutText -TimestampFunction $TimestampFunction
        } catch {
            $OutText.AppendText("$($TimestampFunction.Invoke()) - An unhandled exception occurred: $($_.Exception.Message)`r`n")
        }
    }).AddParameters(@{
		CertCheckScript = $CertCheckScript
        UserChosenServers = $UserChosenServers
        OutputFilePath = $OutputFilePath
        IgnoreFailedServers = $IgnoreFailedServers
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
    $TicketNumber = $NewTicketTextBox.Text.Trim()
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
    } | Out-Null
}

<#
? **********************************************************************************************************************
? END OF GLOBAL VARIABLES AND FUNCTIONS
? **********************************************************************************************************************
#>
<#
? **********************************************************************************************************************
? START OF STARTUP GUI 
? **********************************************************************************************************************
#>

# Initialize the startup form
$StartupForm = New-WindowsForm -SizeX 475 -SizeY 400 -Text '' -StartPosition 'CenterScreen' -TopMost $false -ShowInTaskbar $true -KeyPreview $false -MinimizeBox $false

# Label for choosing Prod Support Tool path
$PSTPathLabel = New-FormLabel -LocationX 30 -LocationY 20 -SizeX 300 -SizeY 20 -Text "Choose path for UniTrac Prod Support Tool:" -Font $global:NormalFont -TextAlign $DefaultTextAlign

# Text box for Prod Support Tool path
$PSTPathTextBox = New-FormTextBox -LocationX 30 -LocationY 40 -SizeX 300 -SizeY 20 -ScrollBars 'None' -Multiline $false -Enabled $true -ReadOnly $false -Text '' -Font $global:NormalFont

# Checkbox for choosing whether to install the Prod Support Tool
$InstallPSTCheckBox = New-FormCheckBox -LocationX 30 -LocationY 70 -SizeX 300 -SizeY 20 -Text "I don't need the Prod Support Tool" -Font $global:NormalSmallFont -Checked $false -Enabled $true

# Button for choosing Prod Support Tool path
$ChoosePSTPathButton = New-FormButton -Text "Choose Path" -LocationX 340 -LocationY 40 -Width 100 -BackColor ([System.Drawing.SystemColors]::Control) -ForeColor ([System.Drawing.SystemColors]::ControlText) -Font $global:NormalFont -Enabled $true -Visible $true
$ChoosePSTPathButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$StartupForm.AcceptButton = $ChoosePSTPathButton

# Label for choosing Documentation path
$DocumentationPathLabel = New-FormLabel -LocationX 30 -LocationY 120 -SizeX 300 -SizeY 20 -Text "Choose path for documentation files:" -Font $global:NormalFont -TextAlign $DefaultTextAlign

# Text box for Documentation path
$DocumentationPathTextBox = New-FormTextBox -LocationX 30 -LocationY 140 -SizeX 300 -SizeY 20 -ScrollBars 'None' -Multiline $false -Enabled $true -ReadOnly $false -Text '' -Font $global:NormalFont

# Button for choosing Documentation path
$ChooseDocumentationPathButton = New-FormButton -Text "Choose Path" -LocationX 340 -LocationY 140 -Width 100 -BackColor ([System.Drawing.SystemColors]::Control) -ForeColor ([System.Drawing.SystemColors]::ControlText) -Font $global:NormalFont -Enabled $true -Visible $true

# Label for choosing path for Ticket Manager
$TicketManagerPathLabel = New-FormLabel -LocationX 30 -LocationY 190 -SizeX 300 -SizeY 20 -Text "Choose path for Ticket Manager:" -Font $global:NormalFont -TextAlign $DefaultTextAlign

# Text box for Ticket Manager path
$TicketManagerPathTextBox = New-FormTextBox -LocationX 30 -LocationY 210 -SizeX 300 -SizeY 20 -ScrollBars 'None' -Multiline $false -Enabled $true -ReadOnly $false -Text '' -Font $global:NormalFont

# Button for choosing Ticket Manager path
$ChooseTicketManagerPathButton = New-FormButton -Text "Choose Path" -LocationX 340 -LocationY 210 -Width 100 -BackColor ([System.Drawing.SystemColors]::Control) -ForeColor ([System.Drawing.SystemColors]::ControlText) -Font $global:NormalFont -Enabled $true -Visible $true

# Add controls to startup form
$StartupForm.Controls.Add($PSTPathLabel)
$StartupForm.Controls.Add($PSTPathTextBox)
$StartupForm.Controls.Add($ChoosePSTPathButton)
$StartupForm.Controls.Add($InstallPSTCheckBox)
$StartupForm.Controls.Add($DocumentationPathLabel)
$StartupForm.Controls.Add($DocumentationPathTextBox)
$StartupForm.Controls.Add($ChooseDocumentationPathButton)
$StartupForm.Controls.Add($TicketManagerPathLabel)
$StartupForm.Controls.Add($TicketManagerPathTextBox)
$StartupForm.Controls.Add($ChooseTicketManagerPathButton)

# Display the startup form
#$StartupForm.ShowDialog() | Out-Null

<#
? **********************************************************************************************************************
? END OF STARTUP GUI 
? **********************************************************************************************************************
#>
<#
? **********************************************************************************************************************
? START OF MAIN GUI 
? **********************************************************************************************************************
#>

# Initialize Help popup form
$HelpForm = New-WindowsForm -SizeX 200 -SizeY 200 -Text '' -StartPosition 'Manual' -TopMost $true -ShowInTaskbar $false -KeyPreview $false -MinimizeBox $false

# Event handler for hiding the Help popup form
$HelpForm.add_Deactivate({
    $HelpForm.Hide()
})

# Initialize the main Desktop Assistant form
$DesktopAssistantForm = New-WindowsForm -SizeX 1000 -SizeY 570 -Text '' -StartPosition 'CenterScreen' -TopMost $false -ShowInTaskbar $true -KeyPreview $true -MinimizeBox $false

# Mouse Down event handler for moving the main Desktop Assistant form
# This is needed since the form has no ControlBox
$DesktopAssistantForm_MouseDown = {
    $global:MouseDown = $true
    $global:MouseClickPoint = [System.Windows.Forms.Cursor]::Position
}

# Mouse Down event handler for moving the main Desktop Assistant form
# This is needed since the form has no ControlBox
$DesktopAssistantForm_MouseMove = {
    if ($global:MouseDown) {
        $CurrentCursorPosition = [System.Windows.Forms.Cursor]::Position
        $FormLocation = $DesktopAssistantForm.Location
        
        $newX = $FormLocation.X + ($CurrentCursorPosition.X - $global:MouseClickPoint.X)
        $newY = $FormLocation.Y + ($CurrentCursorPosition.Y - $global:MouseClickPoint.Y)
        
        $DesktopAssistantForm.Location = New-Object System.Drawing.Point($newX, $newY)
        $global:MouseClickPoint = $CurrentCursorPosition
    }
}

# Mouse Down event handler for moving the main Desktop Assistant form
# This is needed since the form has no ControlBox
$DesktopAssistantForm_MouseUp = {
    $global:MouseDown = $false
}

# Add event handlers to the form
$DesktopAssistantForm.Add_MouseDown($DesktopAssistantForm_MouseDown)
$DesktopAssistantForm.Add_MouseMove($DesktopAssistantForm_MouseMove)
$DesktopAssistantForm.Add_MouseUp($DesktopAssistantForm_MouseUp)

# Button for minimizing the main Desktop Assistant form
$MainFormMinimizeButton = New-FormButton -Text "─" -LocationX 920 -LocationY 0 -Width 40 -BackColor $themeColors.BackColor -ForeColor $themeColors.ForeColor -Font $global:NormalBoldFont -Enabled $true -Visible $true
$MainFormMinimizeButton.Add_Click({ $DesktopAssistantForm.WindowState = 'Minimized' })
$MainFormMinimizeButton.FlatStyle = 'Flat'
$MainFormMinimizeButton.FlatAppearance.BorderSize = 0
$MainFormMinimizeButton.FlatAppearance.MouseOverBackColor = if ($SelectedTheme -eq 'Premium' -or $SelectedTheme -eq 'Custom') { $themeColors.AccentColor } else { $themeColors.ForeColor }
$MainFormMinimizeButton.Add_MouseEnter({ if ($SelectedTheme -eq 'Premium' -or $SelectedTheme -eq 'Custom') { $MainFormMinimizeButton.ForeColor = $themeColors.ForeColor } else { $MainFormMinimizeButton.ForeColor = $themeColors.BackColor } })
$MainFormMinimizeButton.Add_MouseLeave({ $MainFormMinimizeButton.ForeColor = $themeColors.ForeColor })

# Button for closing the main Desktop Assistant form
$MainFormCloseButton = New-FormButton -Text "X" -LocationX 960 -LocationY 0 -Width 40 -BackColor $themeColors.BackColor -ForeColor $themeColors.ForeColor -Font $global:NormalBoldFont -Enabled $true -Visible $true
$MainFormCloseButton.Add_Click({ $DesktopAssistantForm.Close() })
$MainFormCloseButton.FlatStyle = 'Flat'
$MainFormCloseButton.FlatAppearance.BorderSize = 0
$MainFormCloseButton.FlatAppearance.MouseOverBackColor = if ($SelectedTheme -eq 'Premium' -or $SelectedTheme -eq 'Custom') { $themeColors.AccentColor } else { $themeColors.ForeColor }
$MainFormCloseButton.Add_MouseEnter({ if ($SelectedTheme -eq 'Premium' -or $SelectedTheme -eq 'Custom') { $MainFormCloseButton.ForeColor = $themeColors.ForeColor } else { $MainFormCloseButton.ForeColor = $themeColors.BackColor } })
$MainFormCloseButton.Add_MouseLeave({ $MainFormCloseButton.ForeColor = $themeColors.ForeColor })

# Tab control creation
$MainFormTabControl = New-Object System.Windows.Forms.TabControl
$MainFormTabControl.Size = "590,500"
$MainFormTabControl.Location = "5,65"

# System Administator Tools Tab
$SysAdminTab = New-FormTabPage -Font $global:NormalFont -Name "SysAdminTools" -Text "SysAdmin Tools"

# Mouse event handlers for moving the main Desktop Assistant form
$SysAdminTab.Add_MouseDown($DesktopAssistantForm_MouseDown)
$SysAdminTab.Add_MouseMove($DesktopAssistantForm_MouseMove)
$SysAdminTab.Add_MouseUp($DesktopAssistantForm_MouseUp)

# AWS Administrator Tools Tab
$AWSAdminTab = New-FormTabPage -Font $global:NormalFont -Name "AWSAdminTools" -Text "AWS Admin Tools"

# Mouse event handlers for moving the main Desktop Assistant form
$AWSAdminTab.Add_MouseDown($DesktopAssistantForm_MouseDown)
$AWSAdminTab.Add_MouseMove($DesktopAssistantForm_MouseMove)
$AWSAdminTab.Add_MouseUp($DesktopAssistantForm_MouseUp)

# Support Tools Tab
$SupportTab = New-FormTabPage -Font $global:NormalFont -Name "SupportTools" -Text "Support Tools"

# Mouse down event handlers for moving the main Desktop Assistant form
$SupportTab.Add_MouseDown($DesktopAssistantForm_MouseDown)
$SupportTab.Add_MouseMove($DesktopAssistantForm_MouseMove)
$SupportTab.Add_MouseUp($DesktopAssistantForm_MouseUp)

# Ticket Manager Tab
$TicketManagerTab = New-FormTabPage -Font $global:NormalFont -Name "TicketManager" -Text "Ticket Manager"

# Mouse down event handlers for moving the main Desktop Assistant form
$TicketManagerTab.Add_MouseDown($DesktopAssistantForm_MouseDown)
$TicketManagerTab.Add_MouseMove($DesktopAssistantForm_MouseMove)
$TicketManagerTab.Add_MouseUp($DesktopAssistantForm_MouseUp)

# Create a textbox for logging
$OutText = New-FormTextBox -LocationX 600 -LocationY 85 -SizeX 400 -SizeY 450 -ScrollBars 'Vertical' -Multiline $true -Enabled $true -ReadOnly $true -Text '' -Font $global:NormalFont

# Button for clearing logging text box output
$ClearOutTextButton = New-FormButton -Text "Clear Output" -LocationX 660 -LocationY 540 -Width 100 -BackColor $global:DisabledBackColor -ForeColor $global:DisabledForeColor -Font $global:NormalFont -Enabled $false -Visible $true

# Button for saving logging text box output
$SaveOutTextButton = New-FormButton -Text "Save Output" -LocationX 840 -LocationY 540 -Width 100 -BackColor $global:DisabledBackColor -ForeColor $global:DisabledForeColor -Font $global:NormalFont -Enabled $false -Visible $true

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
        if ($SaveFileDialog.FileName -ne '') {
            $OutText.Text | Out-File -FilePath $SaveFileDialog.FileName
    }
})

# Create a picture box for Deadpool! Chimichangas not included. Don't be sad. You can always make your own.
$DeadpoolsVeryOwnPictureBoxTM = New-Object System.Windows.Forms.PictureBox
$DeadpoolsVeryOwnPictureBoxTM.Location = New-Object System.Drawing.Point(700, (-20)) # Using a negative number for Y location to move the picture box up in order to show more of the image
$DeadpoolsVeryOwnPictureBoxTM.Size = New-Object System.Drawing.Size(234, 174)
$DeadpoolsVeryOwnPictureBoxTM.SizeMode = 'Zoom'
$DeadpoolsVeryOwnPictureBoxTM.BackColor = 'Transparent'

# Load the .png image of Ryan Reynolds I mean Deadpool
$DeadpoolPath = Join-Path -Path $PSScriptRoot -ChildPath $ConfigValues.DeadPoolWaitingToBePaintedLikeOneOfYourFrenchGirls
$DeadPoolWaitingToBePaintedLikeOneOfYourFrenchGirls = [System.Drawing.Image]::FromFile($DeadpoolPath)
$DeadpoolsVeryOwnPictureBoxTM.Image = $DeadPoolWaitingToBePaintedLikeOneOfYourFrenchGirls

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
$MenuStrip.Font = [System.Drawing.Font]::new($global:DefaultUserFont, 9, [System.Drawing.FontStyle]::Regular)
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
    if ($global:IsFeedbackPopupActive) {
        $OutText.AppendText("$(Get-Timestamp) - Feedback form is already open.`r`n")
        $script:FeedbackPopup.Activate()
        return
    }
    $OutText.AppendText("$(Get-Timestamp) - Launching feedback form...`r`n")

    # Create the Feedback form
    $script:FeedbackPopup = New-WindowsForm -SizeX 300 -SizeY 350 -Text '' -StartPosition 'CenterScreen' -TopMost $true -ShowInTaskbar $true -KeyPreview $true -MinimizeBox $false
    $global:IsFeedbackPopupActive = $true

    # Mouse Down event handler for moving the main Desktop Assistant form
    # This is needed since the form has no ControlBox
    $script:FeedbackPopup_MouseDown = {
        $global:MouseDown = $true
        $global:MouseClickPoint = [System.Windows.Forms.Cursor]::Position
    }

    # Mouse Down event handler for moving the main Desktop Assistant form
    # This is needed since the form has no ControlBox
    $script:FeedbackPopup_MouseMove = {
        if ($global:MouseDown) {
            $CurrentCursorPosition = [System.Windows.Forms.Cursor]::Position
            $FormLocation = $script:FeedbackPopup.Location
            
            $newX = $FormLocation.X + ($CurrentCursorPosition.X - $global:MouseClickPoint.X)
            $newY = $FormLocation.Y + ($CurrentCursorPosition.Y - $global:MouseClickPoint.Y)
            
            $script:FeedbackPopup.Location = New-Object System.Drawing.Point($newX, $newY)
            $global:MouseClickPoint = $CurrentCursorPosition
        }
    }

    # Mouse Down event handler for moving the main Desktop Assistant form
    # This is needed since the form has no ControlBox
    $script:FeedbackPopup_MouseUp = {
        $global:MouseDown = $false
    }

    # Add event handlers to the form
    $script:FeedbackPopup.Add_MouseDown($script:FeedbackPopup_MouseDown)
    $script:FeedbackPopup.Add_MouseMove($script:FeedbackPopup_MouseMove)
    $script:FeedbackPopup.Add_MouseUp($script:FeedbackPopup_MouseUp)

    # Button for minimizing the Submit Feedback popup form
    $script:FeedbackPopupMinimizeButton = New-FormButton -Text "─" -LocationX 220 -LocationY 0 -Width 40 -BackColor $themeColors.BackColor -ForeColor $themeColors.ForeColor -Font $global:NormalBoldFont -Enabled $true -Visible $true
    $script:FeedbackPopupMinimizeButton.Add_Click({ $script:FeedbackPopup.WindowState = 'Minimized' })
    $script:FeedbackPopupMinimizeButton.FlatStyle = 'Flat'
    $script:FeedbackPopupMinimizeButton.FlatAppearance.BorderSize = 0
    $script:FeedbackPopupMinimizeButton.FlatAppearance.MouseOverBackColor = if ($SelectedTheme -eq 'Premium' -or $SelectedTheme -eq 'Custom') { $themeColors.AccentColor } else { $themeColors.ForeColor }
    $script:FeedbackPopupMinimizeButton.Add_MouseEnter({ if ($SelectedTheme -eq 'Premium' -or $SelectedTheme -eq 'Custom') { $script:FeedbackPopupMinimizeButton.ForeColor = $themeColors.ForeColor } else { $script:FeedbackPopupMinimizeButton.ForeColor = $themeColors.BackColor } })
    $script:FeedbackPopupMinimizeButton.Add_MouseLeave({ $script:FeedbackPopupMinimizeButton.ForeColor = $themeColors.ForeColor })

    # Button for closing the Submit Feedback popup form
    $script:FeedbackPopupCloseButton = New-FormButton -Text "X" -LocationX 260 -LocationY 0 -Width 40 -BackColor $themeColors.BackColor -ForeColor $themeColors.ForeColor -Font $global:NormalBoldFont -Enabled $true -Visible $true
    $script:FeedbackPopupCloseButton.Add_Click({ $script:FeedbackPopup.Close() })
    $script:FeedbackPopupCloseButton.FlatStyle = 'Flat'
    $script:FeedbackPopupCloseButton.FlatAppearance.BorderSize = 0
    $script:FeedbackPopupCloseButton.FlatAppearance.MouseOverBackColor = if ($SelectedTheme -eq 'Premium' -or $SelectedTheme -eq 'Custom') { $themeColors.AccentColor } else { $themeColors.ForeColor }
    $script:FeedbackPopupCloseButton.Add_MouseEnter({ if ($SelectedTheme -eq 'Premium' -or $SelectedTheme -eq 'Custom') { $script:FeedbackPopupCloseButton.ForeColor = $themeColors.ForeColor } else { $script:FeedbackPopupCloseButton.ForeColor = $themeColors.BackColor } })
    $script:FeedbackPopupCloseButton.Add_MouseLeave({ $script:FeedbackPopupCloseButton.ForeColor = $themeColors.ForeColor })

    # Label for providing UserName
    $script:UserNameLabel = New-FormLabel -LocationX 20 -LocationY 40 -SizeX 200 -SizeY 20 -Text "Add Your Name?" -Font $global:NormalBoldFont -TextAlign $DefaultTextAlign

    # Radio button for remaining anonymous
    $script:AnonymousRadioButton = New-Object System.Windows.Forms.RadioButton
    $script:AnonymousRadioButton.Location = New-Object System.Drawing.Point(20, 60)
    $script:AnonymousRadioButton.Size = New-Object System.Drawing.Size(200, 20)
    $script:AnonymousRadioButton.Font = [System.Drawing.Font]::new($global:DefaultUserFont, 9, [System.Drawing.FontStyle]::Regular)
    $script:AnonymousRadioButton.Text = 'Remain Anonymous'
    $script:AnonymousRadioButton.Checked = $true

    # Radio button for providing name
    $script:UserNameRadioButton = New-Object System.Windows.Forms.RadioButton
    $script:UserNameRadioButton.Location = New-Object System.Drawing.Point(20, 85)
    $script:UserNameRadioButton.Size = New-Object System.Drawing.Size(200, 20)
    $script:UserNameRadioButton.Font = [System.Drawing.Font]::new($global:DefaultUserFont, 9, [System.Drawing.FontStyle]::Regular)
    $script:UserNameRadioButton.Text = 'Provide Name'
    $script:UserNameRadioButton.Checked = $false

    # UserName textbox
    $script:UserNameTextBox = New-FormTextBox -LocationX 20 -LocationY 115 -SizeX 170 -SizeY 20 -ScrollBars 'None' -Multiline $false -Enabled $false -ReadOnly $false -Text '' -Font $global:NormalFont

    # Label for feedback form
    $script:FeedbackLabel = New-FormLabel -LocationX 20 -LocationY 180 -SizeX 200 -SizeY 20 -Text 'Enter your feedback below:' -Font $global:NormalBoldFont -TextAlign $DefaultTextAlign

    # Textbox for feedback form
    $script:FeedbackTextBox = New-FormTextBox -LocationX 20 -LocationY 200 -SizeX 240 -SizeY 50 -ScrollBars 'Vertical' -Multiline $true -Enabled $true -ReadOnly $false -Text '' -Font $global:NormalFont

    # Button for submitting feedback
    $script:SubmitFeedbackButton = New-FormButton -Text "Submit" -LocationX 20 -LocationY 275 -Width 100 -BackColor $global:DisabledBackColor -ForeColor $global:DisabledForeColor -Font $global:NormalFont -Enabled $false -Visible $true

    # Check if DefaultUserTheme has a value or is null
    if ($null -ne $ConfigValues.DefaultUserTheme -and $ConfigValues.DefaultUserTheme -ne '') {
        # Get theme
        $SelectedTheme = $ColorTheme.PSObject.Properties | Where-Object { $_.Value.$($ConfigValues.DefaultUserTheme) } | Select-Object -ExpandProperty Name
        $themeColors = $ColorTheme.$SelectedTheme.$($ConfigValues.DefaultUserTheme)

        if ($SelectedTheme -eq 'Premium' -or $SelectedTheme -eq 'Custom') {
            $script:UserNameTextBox.BackColor = $themeColors.AccentColor
            $script:FeedbackTextBox.BackColor = $themeColors.AccentColor
        } else {
            $script:UserNameTextBox.BackColor = [System.Drawing.SystemColors]::Control
            $script:FeedbackTextBox.BackColor = [System.Drawing.SystemColors]::Control
            $global:DisabledBackColor = '#A9A9A9'
        }

        if ($ConfigValues.DefaultUserTheme -eq 'USA') {
            Enable-USAThemeTextColor
        }
        
        $script:FeedbackPopup.BackColor = $themeColors.BackColor
        $script:FeedbackPopup.ForeColor = $themeColors.ForeColor
        $script:FeedbackPopupMinimizeButton.BackColor = $themeColors.BackColor
        $script:FeedbackPopupMinimizeButton.ForeColor = $themeColors.ForeColor
        $script:FeedbackPopupCloseButton.BackColor = $themeColors.BackColor
        $script:FeedbackPopupCloseButton.ForeColor = $themeColors.ForeColor
        $script:UserNameLabel.ForeColor = $themeColors.ForeColor
        $script:FeedbackLabel.BackColor = $themeColors.BackColor
        $script:FeedbackLabel.ForeColor = $themeColors.ForeColor
    }

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
        $FeedbackText = $script:FeedbackTextBox.Text
        $IsUserNameChecked = $script:UserNameRadioButton.Checked
        $UserNameText = $script:UserNameTextBox.Text
        $script:FeedbackTextBox.Text = ''
        $script:UserNameRadioButton.Checked = $false
        $script:AnonymousRadioButton.Checked = $false
        $script:FeedbackPopup.Close()
        Open-SubmitFeedbackRunspace -ConfigValuesSubmitFeedbackScript $ConfigValues.SubmitFeedbackScript -UserProfilePath $userProfilePath -FeedbackText $FeedbackText -IsUserNameChecked $IsUserNameChecked -UserNameText $UserNameText -OutText $OutText -TimestampFunction ${function:Get-Timestamp} -TestingWebhookURL $script:TestingWebhookURL
    })

    # Button click event for closing the feedback form
    $script:FeedbackPopup.add_FormClosed({
        $global:IsFeedbackPopupActive = $false
    })

    # Feedback form build
    $script:FeedbackPopup.Controls.Add($script:FeedbackPopupMinimizeButton)
    $script:FeedbackPopup.Controls.Add($script:FeedbackPopupCloseButton)
    $script:FeedbackPopup.Controls.Add($script:UserNameLabel)
    $script:FeedbackPopup.Controls.Add($script:AnonymousRadioButton)
    $script:FeedbackPopup.Controls.Add($script:UserNameRadioButton)
    $script:FeedbackPopup.Controls.Add($script:UserNameTextBox)
    $script:FeedbackPopup.Controls.Add($script:FeedbackLabel)
    $script:FeedbackPopup.Controls.Add($script:FeedbackTextBox)
    $script:FeedbackPopup.Controls.Add($script:SubmitFeedbackButton)

    $script:FeedbackPopup.Show() | Out-Null
})

# Click event for the File menu Quit option
$MenuQuit.add_Click({ $DesktopAssistantForm.Close() })

# Options menu
$OptionsMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$OptionsMenu.Text = "Options"
$MenuColorTheme = New-Object System.Windows.Forms.ToolStripMenuItem
$MenuColorTheme.Text = "Select Theme"
$MenuThemeBuilder = New-Object System.Windows.Forms.ToolStripMenuItem
$MenuThemeBuilder.Text = "Theme Builder"
$MenuFontPicker = New-Object System.Windows.Forms.ToolStripMenuItem
$MenuFontPicker.Text = "Font Picker"
$MenuHelpOptions = New-Object System.Windows.Forms.ToolStripMenuItem
$MenuHelpOptions.Text = "Help Options"
$ShowHelpIconsMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$ShowToolTipsMenu = New-Object System.Windows.Forms.ToolStripMenuItem

# Color theme sub-menu
$script:CustomThemes = New-Object System.Windows.Forms.ToolStripMenuItem
$script:CustomThemes.Text = "Custom Themes"
# Initial population of the theme menu
foreach ($team in $ColorTheme.Custom.PSObject.Properties) {
    $MenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $MenuItem.Text = $team.Name
    $MenuItem.add_Click({
        $selectedTheme = $this.Text -replace "• ", ''  # Remove bullet point if present
        if ($selectedTheme -eq $ConfigValues.DefaultUserTheme) {
            $OutText.AppendText("$(Get-Timestamp) - The selected theme is already active.`r`n")
        }
        else {
            Update-MainTheme -Team $selectedTheme -Category 'Custom' -ColorData $ColorTheme
            $ConfigValues.DefaultUserTheme = $selectedTheme  # Update the current theme
            # Update-ThemeMenuItems  -MenuCategories @($script:CustomThemes, $script:MLBThemes, $script:NBAThemes, $NFLThemes, $PremiumThemes)
            # Create a list specifically for ToolStripMenuItem objects
            $menuCategoriesList = New-Object 'System.Collections.Generic.List[System.Windows.Forms.ToolStripMenuItem]'
            # Add each menu item to the list
            $menuCategoriesList.Add($script:CustomThemes)
            $menuCategoriesList.Add($script:MLBThemes)
            $menuCategoriesList.Add($script:NBAThemes)
            $menuCategoriesList.Add($NFLThemes)
            $menuCategoriesList.Add($PremiumThemes)

            # Pass the list to the function
            Update-ThemeMenuItems -MenuCategories $menuCategoriesList.ToArray()
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
        $selectedTheme = $this.Text -replace "• ", ''  # Remove bullet point if present
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
        $selectedTheme = $this.Text -replace "• ", ''  # Remove bullet point if present
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
        $selectedTheme = $this.Text -replace "• ", ''  # Remove bullet point if present
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
        $selectedTheme = $this.Text -replace "• ", ''  # Remove bullet point if present
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

    # Create the Theme Builder form
    $script:ThemeBuilderForm = New-WindowsForm -SizeX 750 -SizeY 720 -Text '' -StartPosition 'CenterScreen' -TopMost $false -ShowInTaskbar $true -KeyPreview $true -MinimizeBox $true
    $global:IsThemeBuilderPopupActive = $true

    # Button for minimizing the Submit Feedback popup form
    $script:ThemeBuilderMinimizeButton = New-FormButton -Text "─" -LocationX 670 -LocationY 0 -Width 40 -BackColor $ControlColor -ForeColor $ControlColorText -Font $global:NormalBoldFont -Enabled $true -Visible $true
    $script:ThemeBuilderMinimizeButton.Add_Click({ $script:ThemeBuilderForm.WindowState = 'Minimized' })
    $script:ThemeBuilderMinimizeButton.FlatStyle = 'Flat'
    $script:ThemeBuilderMinimizeButton.FlatAppearance.BorderSize = 0
    $script:ThemeBuilderMinimizeButton.FlatAppearance.MouseOverBackColor = $ControlColorText
    $script:ThemeBuilderMinimizeButton.Add_MouseEnter({ $script:ThemeBuilderMinimizeButton.ForeColor = $ControlColor })
    $script:ThemeBuilderMinimizeButton.Add_MouseLeave({ $script:ThemeBuilderMinimizeButton.ForeColor = $ControlColorText })

    # Button for closing the Submit Feedback popup form
    $script:ThemeBuilderCloseButton = New-FormButton -Text "X" -LocationX 710 -LocationY 0 -Width 40 -BackColor $ControlColor -ForeColor $ControlColorText -Font $global:NormalBoldFont -Enabled $true -Visible $true
    $script:ThemeBuilderCloseButton.Add_Click({ $script:ThemeBuilderForm.Close() })
    $script:ThemeBuilderCloseButton.FlatStyle = 'Flat'
    $script:ThemeBuilderCloseButton.FlatAppearance.BorderSize = 0
    $script:ThemeBuilderCloseButton.FlatAppearance.MouseOverBackColor = $ControlColorText
    $script:ThemeBuilderCloseButton.Add_MouseEnter({ $script:ThemeBuilderCloseButton.ForeColor = $ControlColor })
    $script:ThemeBuilderCloseButton.Add_MouseLeave({ $script:ThemeBuilderCloseButton.ForeColor = $ControlColorText })

    # Mouse Down event handler for moving the Theme Builder form
    # This is needed since the form has no ControlBox
    $script:ThemeBuilderForm_MouseDown = {
        $global:MouseDown = $true
        $global:MouseClickPoint = [System.Windows.Forms.Cursor]::Position
    }

    # Mouse Down event handler for moving the Theme Builder form
    # This is needed since the form has no ControlBox
    $script:ThemeBuilderForm_MouseMove = {
        if ($global:MouseDown) {
            $CurrentCursorPosition = [System.Windows.Forms.Cursor]::Position
            $FormLocation = $script:ThemeBuilderForm.Location
            
            $newX = $FormLocation.X + ($CurrentCursorPosition.X - $global:MouseClickPoint.X)
            $newY = $FormLocation.Y + ($CurrentCursorPosition.Y - $global:MouseClickPoint.Y)
            
            $script:ThemeBuilderForm.Location = New-Object System.Drawing.Point($newX, $newY)
            $global:MouseClickPoint = $CurrentCursorPosition
        }
    }

    # Mouse Down event handler for moving the Theme Builder form
    # This is needed since the form has no ControlBox
    $script:ThemeBuilderForm_MouseUp = {
        $global:MouseDown = $false
    }

    # Add event handlers to the form
    $script:ThemeBuilderForm.Add_MouseDown($script:ThemeBuilderForm_MouseDown)
    $script:ThemeBuilderForm.Add_MouseMove($script:ThemeBuilderForm_MouseMove)
    $script:ThemeBuilderForm.Add_MouseUp($script:ThemeBuilderForm_MouseUp)

    # Theme Builder Tab control creation
    $script:ThemeBuilderMainFormTabControl = New-object System.Windows.Forms.TabControl
    $script:ThemeBuilderMainFormTabControl.Size = "442,375"
    $script:ThemeBuilderMainFormTabControl.Location = "4,49"

    # Theme Builder System Administrator Tools Tab
    $script:ThemeBuilderSysAdminTab = New-FormTabPage -Font $global:NormalSmallFont -Name "ThemeBuilderSysAdminTools" -Text "SysAdmin Tools"

    # Mouse Down event handler for moving the Theme Builder form
    # This is needed since the form has no ControlBox
    $script:ThemeBuilderSysAdminTab_MouseDown = {
        $global:MouseDown = $true
        $global:MouseClickPoint = [System.Windows.Forms.Cursor]::Position
    }

    # Mouse Down event handler for moving the Theme Builder form
    # This is needed since the form has no ControlBox
    $script:ThemeBuilderSysAdminTab_MouseMove = {
        if ($global:MouseDown) {
            $CurrentCursorPosition = [System.Windows.Forms.Cursor]::Position
            $FormLocation = $script:ThemeBuilderForm.Location
            
            $newX = $FormLocation.X + ($CurrentCursorPosition.X - $global:MouseClickPoint.X)
            $newY = $FormLocation.Y + ($CurrentCursorPosition.Y - $global:MouseClickPoint.Y)
            
            $script:ThemeBuilderForm.Location = New-Object System.Drawing.Point($newX, $newY)
            $global:MouseClickPoint = $CurrentCursorPosition
        }
    }

    # Mouse Down event handler for moving the Theme Builder form
    # This is needed since the form has no ControlBox
    $script:ThemeBuilderSysAdminTab_MouseUp = {
        $global:MouseDown = $false
    }

    # Add event handlers to the form
    $script:ThemeBuilderSysAdminTab.Add_MouseDown($script:ThemeBuilderForm_MouseDown)
    $script:ThemeBuilderSysAdminTab.Add_MouseMove($script:ThemeBuilderForm_MouseMove)
    $script:ThemeBuilderSysAdminTab.Add_MouseUp($script:ThemeBuilderForm_MouseUp)

    # Theme Builder textbox for logging
    $script:ThemeBuilderOutText = New-FormTextBox -LocationX 450 -LocationY 64 -SizeX 300 -SizeY 335 -ScrollBars 'Vertical' -Multiline $true -Enabled $true -ReadOnly $true -Text '' -Font $global:NormalSmallFont
    $script:ThemeBuilderOutText.AppendText("$(Get-Timestamp) - Foreground colors are used for labels and buttons.`r`n")
    $script:ThemeBuilderOutText.AppendText("$(Get-Timestamp) - Disabled colors are for disabled buttons, like the 'Start' and 'Clear Output' buttons bordering this box.`r`n")

    # Theme Builder button for clearing logging text box output
    $script:ThemeBuilderClearOutTextButton = New-FormButton -Text "Clear Output" -LocationX 495 -LocationY 400 -Width 75 -BackColor $ControlColor -ForeColor $ControlColorText -Font $global:NormalSmallFont -Enabled $false -Visible $true

    # Theme Builder button for saving logging text box output
    $script:ThemeBuilderSaveOutTextButton = New-FormButton -Text "Save Output" -LocationX 630 -LocationY 400 -Width 75 -BackColor $ControlColor -ForeColor $ControlColorText -Font $global:NormalSmallFont -Enabled $true -Visible $true

    # Theme Builder Restarts tab control creation
    $ThemeBuilderRestartsTabControl = New-object System.Windows.Forms.TabControl
    $ThemeBuilderRestartsTabControl.Size = "187,187"
    $ThemeBuilderRestartsTabControl.Location = "169,56"

    # Theme Builder Individual servers list box
    $script:ThemeBuilderServersListBox = New-FormListBox -LocationX 4 -LocationY 71 -SizeX 150 -SizeY 180 -SelectionMode 'One' -Font $global:NormalSmallFont
    $script:ThemeBuilderServersListBox.Items.Add("Accent colors are used")
    $script:ThemeBuilderServersListBox.Items.Add("for list boxes like this.")

    # Theme Builder Services list box
    $script:ThemeBuilderServicesListBox = New-FormListBox -LocationX 0 -LocationY 0 -SizeX 184 -SizeY 180 -SelectionMode 'MultiExtended' -Font $global:NormalSmallFont
    $script:ThemeBuilderServicesListBox.Items.Add("Accent colors are also")
    $script:ThemeBuilderServicesListBox.Items.Add("used for text boxes.")

    # Theme Builder Combobox for application selection
    $script:ThemeBuilderAppListCombo = New-FormComboBox -LocationX 4 -LocationY 49 -SizeX 150 -SizeY 150 -Font $global:NormalSmallFont

    # Populate the ThemeBuilderAppListCombo with the list of applications from the CSV
    foreach ($header in $csvHeaders) {
        [void]$script:ThemeBuilderAppListCombo.Items.Add($header)
    }

    # Theme Builder Label applist combo box
    $script:ThemeBuilderAppListLabel = New-FormLabel -LocationX 4 -LocationY 30 -SizeX 113 -SizeY 15 -Text "Select a Server" -Font $global:SmallBoldFont -TextAlign $DefaultTextAlign

    # ThemeBuilder tab for services list
    $ThemeBuilderServicesTab = New-FormTabPage -Font $global:NormalSmallFont -Name "ServicesTab" -Text "Services"

    # ThemeBuilder Button for restarting services
    $script:ThemeBuilderRestartButton = New-FormButton -Text "Restart" -LocationX 370 -LocationY 56 -Width 56 -BackColor $ControlColor -ForeColor $ControlColorText -Font $global:NormalSmallFont -Enabled $true -Visible $true

    # ThemeBuilder Button for starting services
    $script:ThemeBuilderStartButton = New-FormButton -Text "Start" -LocationX 370 -LocationY 86 -Width 56 -BackColor $ControlColor -ForeColor $ControlColorText -Font $global:NormalSmallFont -Enabled $false -Visible $true

    # ThemeBuilder Button for stopping services
    $script:ThemeBuilderStopButton = New-FormButton -Text "Stop" -LocationX 370 -LocationY 116 -Width 56 -BackColor $ControlColor -ForeColor $ControlColorText -Font $global:NormalSmallFont -Enabled $true -Visible $true

    # Label for entering text in color text boxes
    $script:ThemeBuilderHelpLabel = New-FormLabel -LocationX 0 -LocationY 450 -SizeX 750 -SizeY 20 -Text "Enter colors in Hex format (e.g. #0000FF) or as a color name (e.g. Blue)" -Font $global:NormalBoldFont -TextAlign $MiddleTextAlign

    # Button for launching Color Picker window
    $script:ThemeBuilderColorPickerButton = New-FormButton -Text "Color Picker" -LocationX 325 -LocationY 480 -Width 100 -BackColor $ControlColor -ForeColor $ControlColorText -Font $global:NormalFont -Enabled $true -Visible $true

    # Text box for entering BackColor
    $script:ThemeBuilderBackColorTextBox = New-FormTextBox -LocationX 60 -LocationY 550 -SizeX 100 -SizeY 20 -ScrollBars 'None' -Multiline $false -Enabled $true -ReadOnly $false -Text '' -Font $global:NormalFont

    # Label for BackColor text box
    $script:ThemeBuilderBackColorLabel = New-FormLabel -LocationX 60 -LocationY 520 -SizeX 100 -SizeY 20 -Text 'BackColor' -Font $global:NormalBoldFont -TextAlign $DefaultTextAlign

    # Text box for entering ForeColor
    $script:ThemeBuilderForeColorTextBox = New-FormTextBox -LocationX 236 -LocationY 550 -SizeX 100 -SizeY 20 -ScrollBars 'None' -Multiline $false -Enabled $true -ReadOnly $false -Text '' -Font $global:NormalFont

    # Label for ForeColor text box
    $script:ThemeBuilderForeColorLabel = New-FormLabel -LocationX 236 -LocationY 520 -SizeX 100 -SizeY 20 -Text 'ForeColor' -Font $global:NormalBoldFont -TextAlign $DefaultTextAlign

    # Text box for entering AccentColor
    $script:ThemeBuilderAccentColorTextBox = New-FormTextBox -LocationX 412 -LocationY 550 -SizeX 100 -SizeY 20 -ScrollBars 'None' -Multiline $false -Enabled $true -ReadOnly $false -Text '' -Font $global:NormalFont

    # Label for AccentColor text box
    $script:ThemeBuilderAccentColorLabel = New-FormLabel -LocationX 412 -LocationY 520 -SizeX 100 -SizeY 20 -Text 'Accent' -Font $global:NormalBoldFont -TextAlign $DefaultTextAlign

    # Text box for entering DisabledColor
    $script:ThemeBuilderDisabledColorTextBox = New-FormTextBox -LocationX 588 -LocationY 550 -SizeX 100 -SizeY 20 -ScrollBars 'None' -Multiline $false -Enabled $true -ReadOnly $false -Text '' -Font $global:NormalFont

    # Label for DisabledColor text box
    $script:ThemeBuilderDisabledColorLabel = New-FormLabel -LocationX 588 -LocationY 520 -SizeX 100 -SizeY 20 -Text 'Disabled' -Font $global:NormalBoldFont -TextAlign $DefaultTextAlign

    # Button for applying user-defined colors
    $script:ThemeBuilderApplyThemeButton = New-FormButton -Text "Apply Theme" -LocationX 120 -LocationY 600 -Width 100 -BackColor $ControlColor -ForeColor $ControlColorText -Font $global:NormalFont -Enabled $false -Visible $true

    # Button for saving user-defined colors
    $script:ThemeBuilderSaveThemeButton = New-FormButton -Text "Save Theme" -LocationX 325 -LocationY 600 -Width 100 -BackColor $ControlColor -ForeColor $ControlColorText -Font $global:NormalFont -Enabled $false -Visible $true

    # Button for clearing all text boxes
    $script:ThemeBuilderResetThemeButton = New-FormButton -Text "Reset Theme" -LocationX 530 -LocationY 600 -Width 100 -BackColor $ControlColor -ForeColor $ControlColorText -Font $global:NormalFont -Enabled $false -Visible $true

    $script:ThemeBuilderDeleteEnabled = if ($ColorTheme.Custom.PSObject.Properties.Count -gt 0) {
        $true
    } else {
        $false
    }

    # Button for deleting custom themes
    $script:ThemeBuilderDeleteThemesButton = New-FormButton -Text "Delete Custom Themes" -LocationX 300 -LocationY 640 -Width 150 -BackColor $ControlColor -ForeColor $ControlColorText -Font $global:NormalFont -Enabled $script:ThemeBuilderDeleteEnabled -Visible $true

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

    # Event handler for clicking the Color Picker button
    $script:ThemeBuilderColorPickerButton.add_Click({
        $OutText.AppendText("$(Get-Timestamp) - Launching Color Picker...`r`n")

        # Launch the Color Picker form
        $script:ColorPickerPopup = New-WindowsForm -SizeX 380 -SizeY 660 -Text '' -StartPosition 'CenterScreen' -TopMost $false -ShowInTaskbar $true -KeyPreview $true -MinimizeBox $false
        $global:IsColorPickerPopupActive = $true

        # Mouse Down event handler for moving the Theme Builder form
        # This is needed since the form has no ControlBox
        $script:ColorPickerPopup_MouseDown = {
            $global:MouseDown = $true
            $global:MouseClickPoint = [System.Windows.Forms.Cursor]::Position
        }

        # Mouse Down event handler for moving the Theme Builder form
        # This is needed since the form has no ControlBox
        $script:ColorPickerPopup_MouseMove = {
            if ($global:MouseDown) {
                $CurrentCursorPosition = [System.Windows.Forms.Cursor]::Position
                $FormLocation = $script:ColorPickerPopup.Location
                
                $newX = $FormLocation.X + ($CurrentCursorPosition.X - $global:MouseClickPoint.X)
                $newY = $FormLocation.Y + ($CurrentCursorPosition.Y - $global:MouseClickPoint.Y)
                
                $script:ColorPickerPopup.Location = New-Object System.Drawing.Point($newX, $newY)
                $global:MouseClickPoint = $CurrentCursorPosition
            }
        }

        # Mouse Down event handler for moving the Theme Builder form
        # This is needed since the form has no ControlBox
        $script:ColorPickerPopup_MouseUp = {
            $global:MouseDown = $false
        }

        # Add event handlers to the form
        $script:ColorPickerPopup.Add_MouseDown($script:ColorPickerPopup_MouseDown)
        $script:ColorPickerPopup.Add_MouseMove($script:ColorPickerPopup_MouseMove)
        $script:ColorPickerPopup.Add_MouseUp($script:ColorPickerPopup_MouseUp)

        # Button for minimizing the Color Picker popup form
        $script:ColorPickerMinimizeButton = New-FormButton -Text "─" -LocationX 295 -LocationY 0 -Width 40 -BackColor $ControlColor -ForeColor $ControlColorText -Font $global:NormalBoldFont -Enabled $true -Visible $true
        $script:ColorPickerMinimizeButton.Add_Click({ $script:ColorPickerPopup.WindowState = 'Minimized' })
        $script:ColorPickerMinimizeButton.FlatStyle = 'Flat'
        $script:ColorPickerMinimizeButton.FlatAppearance.BorderSize = 0
        $script:ColorPickerMinimizeButton.FlatAppearance.MouseOverBackColor = $ControlColorText
        $script:ColorPickerMinimizeButton.Add_MouseEnter({ $script:ColorPickerMinimizeButton.ForeColor = $ControlColor })
        $script:ColorPickerMinimizeButton.Add_MouseLeave({ $script:ColorPickerMinimizeButton.ForeColor =  $ControlColorText })

        # Button for closing the Color Picker popup form
        $script:ColorPickerCloseButton = New-FormButton -Text "X" -LocationX 335 -LocationY 0 -Width 40 -BackColor $ControlColor -ForeColor $ControlColorText -Font $global:NormalBoldFont -Enabled $true -Visible $true
        $script:ColorPickerCloseButton.Add_Click({ $script:ColorPickerPopup.Close() })
        $script:ColorPickerCloseButton.FlatStyle = 'Flat'
        $script:ColorPickerCloseButton.FlatAppearance.BorderSize = 0
        $script:ColorPickerCloseButton.FlatAppearance.MouseOverBackColor = $ControlColorText
        $script:ColorPickerCloseButton.Add_MouseEnter({ $script:ColorPickerCloseButton.ForeColor = $ControlColor })
        $script:ColorPickerCloseButton.Add_MouseLeave({ $script:ColorPickerCloseButton.ForeColor = $ControlColorText })
 
        # Create a panel with scrollbars
        $script:ColorPickerPanel = New-Object System.Windows.Forms.Panel
        $script:ColorPickerPanel.Location = New-Object System.Drawing.Point(5, 75)
        $script:ColorPickerPanel.Size = New-Object System.Drawing.Size(370, 575)
        $script:ColorPickerPanel.AutoScroll = $true

        # Button for applying color choices
        $script:ColorPickerApplyButton = New-FormButton -Text "Apply Choices" -LocationX 59 -LocationY 45 -Width 100 -BackColor $ControlColor -ForeColor $ControlColorText -Font $global:NormalFont -Enabled $true -Visible $true

        # Button for clearing color choices
        $ColorPickerClearButton = New-FormButton -Text "Clear Choices" -LocationX 219 -LocationY 45 -Width 100 -BackColor $ControlColor -ForeColor $ControlColorText -Font $global:NormalFont -Enabled $true -Visible $true

        # Define label size and spacing
        $LabelWidth = 200
        $LabelHeight = 20
        $ColorLineHeight = 20
        $ColorLineWidth = 40
        $VerticalSpacing = 5
        $yCoordinate = 10

        # Script block for updating the ThemeBuilder text boxes based on the selected colors
        $UpdateThemeBuilderTextBoxScriptBlock = {
            param($ComboBoxSender, $e)
        
            # Get the selected color type and color
            $selectedColorType = $ComboBoxSender.SelectedItem
            if ($selectedColorType -ne $null) {
                $colorLabel = $ComboBoxSender.Parent.Controls | Where-Object { $_ -is [System.Windows.Forms.Label] -and $_.Location.Y -eq $ComboBoxSender.Location.Y }
                $selectedColor = $colorLabel.Text.Split('-')[0].Trim()
        
                # Store the selection in the hashtable
                $SelectedColors[$selectedColorType] = $selectedColor
            }
        }

        $ManageComboBoxItemsScriptBlock = {
            param($ComboSender, $e)
        
            $currentSelection = $ComboSender.SelectedItem
            $previousSelection = $ComboSender.Tag

            # Function to update items in a ComboBox
            function Update-ComboBoxItems {
                param($comboBox, $itemToAdd, $itemToRemove)
                if ($itemToAdd) {
                    # Remove the item if it already exists to avoid duplicates
                    $comboBox.Items.Remove($itemToAdd)
                    # Find the correct index to insert the item based on its predefined order
                    $index = 0
                    for (; $index -lt $comboBox.Items.Count; $index++) {
                        if ($script:ColorTypeOrder[$comboBox.Items[$index]] -gt $script:ColorTypeOrder[$itemToAdd]) {
                            break
                        }
                    }
                    # Insert the item at the correct position
                    $comboBox.Items.Insert($index, $itemToAdd)
                }
                if ($itemToRemove) {
                    $comboBox.Items.Remove($itemToRemove)
                }
            }
        
            # Iterate through all ComboBoxes to manage items
            foreach ($otherComboBox in $Global:comboBoxes) {
                if ($otherComboBox -ne $ComboSender) {  # Skip the sender itself
                    Update-ComboBoxItems -comboBox $otherComboBox -itemToAdd $previousSelection -itemToRemove $currentSelection
                }
            }
        
            # Update the Tag with the current selection for the next event
            $ComboSender.Tag = $currentSelection
        }

        # Define color types
        $script:ColorTypes = @('BackColor', 'ForeColor', 'AccentColor', 'DisabledColor')

        # Define color types with their order
        $script:ColorTypeOrder = @{
            'BackColor' = 1;
            'ForeColor' = 2;
            'AccentColor' = 3;
            'DisabledColor' = 4
        }

        # Initialize the global list of ComboBoxes
        $Global:comboBoxes = New-Object System.Collections.ArrayList

        # Add color names and hex values
        foreach ($ColorName in [enum]::GetNames([System.Drawing.KnownColor])) {
            $color = [System.Drawing.Color]::FromName($ColorName)
            $hex = "#{0:X2}{1:X2}{2:X2}" -f $color.R, $color.G, $color.B

            # Create label for color name and hex value
            $ColorLabel = New-FormLabel -LocationX 5 -LocationY $yCoordinate -SizeX $LabelWidth -SizeY $LabelHeight -Text "$ColorName - $hex" -Font $global:NormalFont -TextAlign $DefaultTextAlign
            $script:ColorPickerPanel.Controls.Add($ColorLabel)

            # Calculate X-coordinate for color line
            $ColorLineX = 20 + $LabelWidth

            # Create label for color line
            $ColorLine = New-FormLabel -LocationX $ColorLineX -LocationY $yCoordinate -SizeX $ColorLineWidth -SizeY $ColorLineHeight -Text "" -Font $global:NormalFont -TextAlign $DefaultTextAlign
            $ColorLine.BackColor = $color
            $script:ColorPickerPanel.Controls.Add($ColorLine)

            # Calculate X-coordinate for combobox
            $ComboX = $ColorLineX + $ColorLineWidth + 10

            # Combobox for selecting this color
            $ColorComboBox = New-FormComboBox -LocationX $ComboX -LocationY $yCoordinate -SizeX 80 -SizeY 20 -Font $global:NormalSmallFont

            # Add color types to the ComboBox
            foreach ($ColorType in $script:ColorTypes) {
                [void]$ColorComboBox.Items.Add($ColorType) | Out-Null
            }

            [void]$Global:comboBoxes.Add($ColorComboBox) | Out-Null

            $ColorComboBox.add_SelectedIndexChanged($UpdateThemeBuilderTextBoxScriptBlock)
            $ColorComboBox.add_SelectedIndexChanged($ManageComboBoxItemsScriptBlock)
            $script:ColorPickerPanel.Controls.Add($ColorComboBox)

            # Increment Y coordinate for next item
            $yCoordinate += $LabelHeight + $VerticalSpacing
        }

        # Event handler for the Apply button
        $script:ColorPickerApplyButton.add_Click({
            # Update the ThemeBuilder TextBoxes based on the stored selections
            if ($script:SelectedColors.ContainsKey("BackColor")) {
                $script:ThemeBuilderBackColorTextBox.Text = $script:SelectedColors["BackColor"]
            }
            if ($script:SelectedColors.ContainsKey("ForeColor")) {
                $script:ThemeBuilderForeColorTextBox.Text = $script:SelectedColors["ForeColor"]
            }
            if ($script:SelectedColors.ContainsKey("AccentColor")) {
                $script:ThemeBuilderAccentColorTextBox.Text = $script:SelectedColors["AccentColor"]
            }
            if ($script:SelectedColors.ContainsKey("DisabledColor")) {
                $script:ThemeBuilderDisabledColorTextBox.Text = $script:SelectedColors["DisabledColor"]
            }
        
            # Clear the hashtable after applying the colors
            $script:SelectedColors.Clear()
        })

        # Event handler for the Clear button
        $ColorPickerClearButton.add_Click({
            $OutText.AppendText("$(Get-Timestamp) - Clearing color selections.`r`n")
            # Loop through all controls in the ColorPickerPanel
            foreach ($control in $script:ColorPickerPanel.Controls) {
                # Check if the control is a ComboBox
                if ($control -is [System.Windows.Forms.ComboBox]) {
                    # Reset the SelectedIndex to -1 (no selection)
                    $control.SelectedIndex = -1
                }
            }
        })

        # Event handler for closing the Color Picker popup
        $script:ColorPickerPopup.Add_FormClosed({
            # Set global variable to indicate that the Color Picker popup is no longer active
            $global:IsColorPickerPopupActive = $false
        })

        # Add the panel to the form
        $script:ColorPickerPopup.Controls.Add($script:ColorPickerMinimizeButton)
        $script:ColorPickerPopup.Controls.Add($script:ColorPickerCloseButton)
        $script:ColorPickerPopup.Controls.Add($script:ColorPickerPanel)
        $script:ColorPickerPopup.Controls.Add($script:ColorPickerApplyButton)
        $script:ColorPickerPopup.Controls.Add($ColorPickerClearButton)

        # Show the form
        $script:ColorPickerPopup.Show()
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
            $script:ThemeBuilderMinimizeButton.BackColor = $script:CustomBackColor
            $script:ThemeBuilderMinimizeButton.ForeColor = $script:CustomForeColor
            $script:ThemeBuilderCloseButton.BackColor = $script:CustomBackColor
            $script:ThemeBuilderCloseButton.ForeColor = $script:CustomForeColor

            $script:ThemeBuilderMinimizeButton.FlatAppearance.MouseOverBackColor = $script:CustomAccentColor
            $script:ThemeBuilderMinimizeButton.Add_MouseEnter({ $script:ThemeBuilderMinimizeButton.ForeColor = $script:CustomForeColor })
            $script:ThemeBuilderMinimizeButton.Add_MouseLeave({ $script:ThemeBuilderMinimizeButton.ForeColor = $script:CustomForeColor })
            $script:ThemeBuilderCloseButton.FlatAppearance.MouseOverBackColor = $script:CustomAccentColor
            $script:ThemeBuilderCloseButton.Add_MouseEnter({ $script:ThemeBuilderCloseButton.ForeColor = $script:CustomForeColor })
            # This scriptblock is needed for the Close button when the form is closing; otherwise, it will throw an exception trying to set ForeColor
            $script:ThemeBuilderMouseLeaveScriptBlock = { 
                $color = if ($script:CustomForeColor) { $script:CustomForeColor } else { [System.Drawing.Color]::Black }
                $script:ThemeBuilderCloseButton.ForeColor = $color
            }
            $script:ThemeBuilderCloseButton.Add_MouseLeave($script:ThemeBuilderMouseLeaveScriptBlock)

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
            $script:ThemeBuilderColorPickerButton.BackColor = $script:CustomForeColor
            $script:ThemeBuilderColorPickerButton.ForeColor = $script:CustomBackColor
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

            # Custom object for holding color values
            $SaveThemePopupColors = New-Object -TypeName PSCustomObject

            # Add properties to the object
            $SaveThemePopupColors | Add-Member -MemberType NoteProperty -Name "BackColor" -Value ""
            $SaveThemePopupColors | Add-Member -MemberType NoteProperty -Name "ForeColor" -Value ""
            $SaveThemePopupColors | Add-Member -MemberType NoteProperty -Name "AccentColor" -Value ""
            $SaveThemePopupColors | Add-Member -MemberType NoteProperty -Name "DisabledColor" -Value ""

            switch ($global:IsThemeApplied) {
                $true {
                    $SaveThemePopupColors.BackColor = $script:CustomBackColor
                    $SaveThemePopupColors.ForeColor = $script:CustomForeColor
                    $SaveThemePopupColors.AccentColor = $script:CustomAccentColor
                    $SaveThemePopupColors.DisabledColor = $script:CustomDisabledColor
                }
                $false {
                    $SaveThemePopupColors.BackColor = $ControlColor
                    $SaveThemePopupColors.ForeColor = $ControlColorText
                    $SaveThemePopupColors.AccentColor = $ControlColorText
                    $SaveThemePopupColors.DisabledColor = $ControlColorText
                }
            }

            # New popup window to save the theme
            $script:ThemeBuilderSaveThemeForm = New-WindowsForm -SizeX 300 -SizeY 320 -Text '' -StartPosition 'CenterScreen' -TopMost $false -ShowInTaskbar $true -KeyPreview $true -MinimizeBox $false
            # Variable for showing popup is active is not needed since form is shown with ShowDialog()
            # This means the form will have focus while it's open, so the user can't interact with the main form

            # Mouse Down event handler for moving the Theme Builder form
            # This is needed since the form has no ControlBox
            $script:ThemeBuilderSaveThemeForm_MouseDown = {
                $global:MouseDown = $true
                $global:MouseClickPoint = [System.Windows.Forms.Cursor]::Position
            }

            # Mouse Down event handler for moving the Theme Builder form
            # This is needed since the form has no ControlBox
            $script:ThemeBuilderSaveThemeForm_MouseMove = {
                if ($global:MouseDown) {
                    $CurrentCursorPosition = [System.Windows.Forms.Cursor]::Position
                    $FormLocation = $script:ThemeBuilderSaveThemeForm.Location
                    
                    $newX = $FormLocation.X + ($CurrentCursorPosition.X - $global:MouseClickPoint.X)
                    $newY = $FormLocation.Y + ($CurrentCursorPosition.Y - $global:MouseClickPoint.Y)
                    
                    $script:ThemeBuilderSaveThemeForm.Location = New-Object System.Drawing.Point($newX, $newY)
                    $global:MouseClickPoint = $CurrentCursorPosition
                }
            }

            # Mouse Down event handler for moving the Theme Builder form
            # This is needed since the form has no ControlBox
            $script:ThemeBuilderSaveThemeForm_MouseUp = {
                $global:MouseDown = $false
            }

            # Add event handlers to the form
            $script:ThemeBuilderSaveThemeForm.Add_MouseDown($script:ThemeBuilderSaveThemeForm_MouseDown)
            $script:ThemeBuilderSaveThemeForm.Add_MouseMove($script:ThemeBuilderSaveThemeForm_MouseMove)
            $script:ThemeBuilderSaveThemeForm.Add_MouseUp($script:ThemeBuilderSaveThemeForm_MouseUp)

            # Button for closing the Save Theme popup form
            $script:SaveThemeCloseButton = New-FormButton -Text "X" -LocationX 260 -LocationY 0 -Width 40 -BackColor $SaveThemePopupColors.BackColor -ForeColor $SaveThemePopupColors.ForeColor -Font $global:NormalBoldFont -Enabled $true -Visible $true
            $script:SaveThemeCloseButton.Add_Click({ $script:ThemeBuilderSaveThemeForm.Close() })
            $script:SaveThemeCloseButton.FlatStyle = 'Flat'
            $script:SaveThemeCloseButton.FlatAppearance.BorderSize = 0
            $script:SaveThemeCloseButton.FlatAppearance.MouseOverBackColor = $SaveThemePopupColors.ForeColor
            $script:SaveThemeCloseButton.Add_MouseEnter({ $script:SaveThemeCloseButton.ForeColor = $SaveThemePopupColors.BackColor })
            # This scriptblock is needed for the Close button when the form is closing; otherwise, it will throw an exception trying to set ForeColor
            $script:ThemeBuilderSaveThemeFormScriptBlock = { 
                $color = if ($script:CustomForeColor) { $SaveThemePopupColors.ForeColor } else { [System.Drawing.Color]::Black }
                $script:SaveThemeCloseButton.ForeColor = $color
            }
            $script:SaveThemeCloseButton.Add_MouseLeave($script:ThemeBuilderSaveThemeFormScriptBlock)

            # Text box for theme name
            $script:ThemeBuilderThemeNameTextBox = New-FormTextBox -LocationX 20 -LocationY 60 -SizeX 150 -SizeY 20 -ScrollBars 'None' -Multiline $false -Enabled $true -ReadOnly $false -Text '' -Font $global:NormalFont

            # Label for theme name text box
            $script:ThemeBuilderThemeNameLabel = New-FormLabel -LocationX 20 -LocationY 40 -SizeX 150 -SizeY 20 -Text "Theme Name" -Font $global:NormalBoldFont -TextAlign $DefaultTextAlign

            # Text box for quote
            $script:ThemeBuilderQuoteTextBox = New-FormTextBox -LocationX 20 -LocationY 140 -SizeX 150 -SizeY 20 -ScrollBars 'None' -Multiline $false -Enabled $true -ReadOnly $false -Text '' -Font $global:NormalFont

            # Label for quote text box
            $script:ThemeBuilderQuoteLabel = New-FormLabel -LocationX 20 -LocationY 120 -SizeX 150 -SizeY 20 -Text "Quote" -Font $global:NormalBoldFont -TextAlign $DefaultTextAlign

            # Button for saving the theme
            $script:ThemeBuilderSaveThemePopupButton = New-FormButton -Text "Save Theme" -LocationX 20 -LocationY 200 -Width 100 -BackColor $ControlColor -ForeColor $ControlColorText -Font $global:NormalSmallFont -Enabled $false -Visible $true

            # Button for saving and applying the theme
            $script:ThemeBuilderSaveAndApplyThemeButton = New-FormButton -Text "Save && Apply" -LocationX 160 -LocationY 200 -Width 100 -BackColor $ControlColor -ForeColor $ControlColorText -Font $global:NormalSmallFont -Enabled $false -Visible $true

            # Apply theme colors to form if theme has been applied
            if ($global:IsThemeApplied) {
                $script:ThemeBuilderSaveThemeForm.BackColor = $SaveThemePopupColors.BackColor
                $script:ThemeBuilderThemeNameTextBox.BackColor = $SaveThemePopupColors.AccentColor
                $script:ThemeBuilderThemeNameLabel.BackColor = $SaveThemePopupColors.BackColor
                $script:ThemeBuilderThemeNameLabel.ForeColor = $SaveThemePopupColors.ForeColor
                $script:ThemeBuilderQuoteTextBox.BackColor = $SaveThemePopupColors.AccentColor
                $script:ThemeBuilderQuoteLabel.BackColor = $SaveThemePopupColors.BackColor
                $script:ThemeBuilderQuoteLabel.ForeColor = $SaveThemePopupColors.ForeColor
                $script:ThemeBuilderSaveThemePopupButton.BackColor = $SaveThemePopupColors.DisabledColor
                $script:ThemeBuilderSaveAndApplyThemeButton.BackColor = $SaveThemePopupColors.DisabledColor
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
            $script:ThemeBuilderSaveThemeForm.Controls.Add($script:SaveThemeCloseButton)
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

        Set-ThemeBuilderDefaultColors
    })

    # Event handler for the Delete Custom Themes button
    $script:ThemeBuilderDeleteThemesButton.add_Click({
        # New popup window to delete custom themes
        $script:ThemeBuilderDeleteThemesForm = New-WindowsForm -SizeX 300 -SizeY 380 -Text '' -StartPosition 'CenterScreen' -TopMost $false -ShowInTaskbar $true -KeyPreview $true -MinimizeBox $true
        # Variable for showing popup is active is not needed since form is shown with ShowDialog()
        # This means the form will have focus while it's open, so the user can't interact with the main form

        $script:CustomBackColor = $script:ThemeBuilderBackColorTextBox.Text
        $script:CustomForeColor = $script:ThemeBuilderForeColorTextBox.Text
        $script:CustomAccentColor = $script:ThemeBuilderAccentColorTextBox.Text
        $script:CustomDisabledColor = $script:ThemeBuilderDisabledColorTextBox.Text

        # Custom object for holding color values
        $DeleteThemePopupColors = New-Object -TypeName PSCustomObject

        # Add properties to the object
        $DeleteThemePopupColors | Add-Member -MemberType NoteProperty -Name "BackColor" -Value ""
        $DeleteThemePopupColors | Add-Member -MemberType NoteProperty -Name "ForeColor" -Value ""
        $DeleteThemePopupColors | Add-Member -MemberType NoteProperty -Name "AccentColor" -Value ""
        $DeleteThemePopupColors | Add-Member -MemberType NoteProperty -Name "DisabledColor" -Value ""

        switch ($global:IsThemeApplied) {
            $true {
                $DeleteThemePopupColors.BackColor = $script:CustomBackColor
                $DeleteThemePopupColors.ForeColor = $script:CustomForeColor
                $DeleteThemePopupColors.AccentColor = $script:CustomAccentColor
                $DeleteThemePopupColors.DisabledColor = $script:CustomDisabledColor
            }
            $false {
                $DeleteThemePopupColors.BackColor = $ControlColor
                $DeleteThemePopupColors.ForeColor = $ControlColorText
                $DeleteThemePopupColors.AccentColor = $ControlColorText
                $DeleteThemePopupColors.DisabledColor = $ControlColorText
            }
        }

        # Button for closing the Save Theme popup form
        $script:DeleteThemeCloseButton = New-FormButton -Text "X" -LocationX 260 -LocationY 0 -Width 40 -BackColor $DeleteThemePopupColors.BackColor -ForeColor $DeleteThemePopupColors.ForeColor -Font $global:NormalBoldFont -Enabled $true -Visible $true
        $script:DeleteThemeCloseButton.Add_Click({ $script:ThemeBuilderDeleteThemesForm.Close() })
        $script:DeleteThemeCloseButton.FlatStyle = 'Flat'
        $script:DeleteThemeCloseButton.FlatAppearance.BorderSize = 0
        $script:DeleteThemeCloseButton.FlatAppearance.MouseOverBackColor = $DeleteThemePopupColors.ForeColor
        $script:DeleteThemeCloseButton.Add_MouseEnter({ $script:DeleteThemeCloseButton.ForeColor = $DeleteThemePopupColors.BackColor })
        # This scriptblock is needed for the Close button when the form is closing; otherwise, it will throw an exception trying to set ForeColor
        $script:ThemeBuilderDeleteThemeFormScriptBlock = { 
        $color = if ($script:CustomForeColor) { $DeleteThemePopupColors.ForeColor } else { [System.Drawing.Color]::Black }
        $script:DeleteThemeCloseButton.ForeColor = $color
        }
        $script:DeleteThemeCloseButton.Add_MouseLeave($script:ThemeBuilderDeleteThemeFormScriptBlock)

        # Mouse Down event handler for moving the Theme Builder form
        # This is needed since the form has no ControlBox
        $script:ThemeBuilderDeleteThemesForm_MouseDown = {
            $global:MouseDown = $true
            $global:MouseClickPoint = [System.Windows.Forms.Cursor]::Position
        }

        # Mouse Down event handler for moving the Theme Builder form
        # This is needed since the form has no ControlBox
        $script:ThemeBuilderDeleteThemesForm_MouseMove = {
            if ($global:MouseDown) {
                $CurrentCursorPosition = [System.Windows.Forms.Cursor]::Position
                $FormLocation = $script:ThemeBuilderDeleteThemesForm.Location
                
                $newX = $FormLocation.X + ($CurrentCursorPosition.X - $global:MouseClickPoint.X)
                $newY = $FormLocation.Y + ($CurrentCursorPosition.Y - $global:MouseClickPoint.Y)
                
                $script:ThemeBuilderDeleteThemesForm.Location = New-Object System.Drawing.Point($newX, $newY)
                $global:MouseClickPoint = $CurrentCursorPosition
            }
        }

        # Mouse Down event handler for moving the Theme Builder form
        # This is needed since the form has no ControlBox
        $script:ThemeBuilderDeleteThemesForm_MouseUp = {
            $global:MouseDown = $false
        }

        # Add event handlers to the form
        $script:ThemeBuilderDeleteThemesForm.Add_MouseDown($script:ThemeBuilderDeleteThemesForm_MouseDown)
        $script:ThemeBuilderDeleteThemesForm.Add_MouseMove($script:ThemeBuilderDeleteThemesForm_MouseMove)
        $script:ThemeBuilderDeleteThemesForm.Add_MouseUp($script:ThemeBuilderDeleteThemesForm_MouseUp)

        # Label for theme name list box
        $script:ThemeBuilderThemeNameListBoxLabel = New-FormLabel -LocationX 65 -LocationY 20 -SizeX 150 -SizeY 40 -Text 'Permanently delete one or more custom themes' -Font $global:NormalBoldFont -TextAlign $MiddleTextAlign

        # List box for theme names
        $script:ThemeBuilderThemeNameListBox = New-FormListBox -LocationX 65 -LocationY 70 -SizeX 150 -SizeY 200 -Font $global:NormalFont -SelectionMode 'MultiExtended'

        # Button for deleting themes
        $script:ThemeBuilderDeleteThemesPopupButton = New-FormButton -Text "Delete Theme" -LocationX 90 -LocationY 290 -Width 100 -BackColor $ControlColor -ForeColor $ControlColorText -Font $global:NormalSmallFont -Enabled $false -Visible $true

        # Apply theme colors to form if theme has been applied
        if ($global:IsThemeApplied) {
            $script:ThemeBuilderDeleteThemesForm.BackColor = $DeleteThemePopupColors.BackColor
            $script:ThemeBuilderThemeNameListBoxLabel.BackColor = $DeleteThemePopupColors.BackColor
            $script:ThemeBuilderThemeNameListBoxLabel.ForeColor = $DeleteThemePopupColors.ForeColor
            $script:ThemeBuilderThemeNameListBox.BackColor = $DeleteThemePopupColors.AccentColor
            $script:ThemeBuilderDeleteThemesPopupButton.BackColor = $DeleteThemePopupColors.DisabledColor
        }

        # Get ColorTheme data from JSON file
        $ColorTheme = Get-Content -Path .\ColorThemes.json | ConvertFrom-Json
        $script:ThemeNames = $ColorTheme.Custom | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name

        foreach ($theme in $script:ThemeNames) {
            $script:ThemeBuilderThemeNameListBox.Items.Add($theme)
        }

        # Logic for enabling the Delete Theme(s) button
        $script:ThemeBuilderThemeNameListBox.add_SelectedIndexChanged({
            if ($script:ThemeBuilderThemeNameListBox.SelectedItems.Count -gt 0) {
                $script:ThemeBuilderDeleteThemesPopupButton.Enabled = $true
                if ($global:IsThemeApplied) {
                    $script:ThemeBuilderDeleteThemesPopupButton.BackColor = $script:CustomForeColor
                    $script:ThemeBuilderDeleteThemesPopupButton.ForeColor = $script:CustomBackColor
                }
                if ($script:ThemeBuilderThemeNameListBox.SelectedItems.Count -eq 1) {
                    $script:ThemeBuilderDeleteThemesPopupButton.Text = "Delete Theme"
                }
                else {
                    $script:ThemeBuilderDeleteThemesPopupButton.Text = "Delete Themes"
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

            # Re-load the $ColorTheme from the file
            $ColorTheme = Get-Content -Path .\ColorThemes.json | ConvertFrom-Json

            # Dispose the existing CustomThemes menu item
            if ($null -ne $script:CustomThemes) {
                $script:CustomThemes.Dispose()
            }

            # Recreate the CustomThemes menu item
            $script:CustomThemes = New-Object System.Windows.Forms.ToolStripMenuItem
            $script:CustomThemes.Text = "Custom Themes"

            # Repopulate the Custom Themes menu
            foreach ($Theme in $ColorTheme.Custom.PSObject.Properties) {
                $NewMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
                $NewMenuItem.Text = $Theme.Name
                $NewMenuItem.add_Click({
                    $selectedTheme = $this.Text -replace "• ", ''
                    if ($selectedTheme -eq $ConfigValues.DefaultUserTheme) {
                        $OutText.AppendText("$(Get-Timestamp) - The selected theme is already active.`r`n")
                    }
                    else {
                        Update-MainTheme -Team $selectedTheme -Category 'Custom' -ColorData $ColorTheme
                        $ConfigValues.DefaultUserTheme = $selectedTheme
                        # Update other menu items and logic as needed
                    }
                })
                $script:CustomThemes.DropDownItems.Add($NewMenuItem) | Out-Null
            }

            # Sort and re-add the dropdown items if needed
            $SortedItems = $script:CustomThemes.DropDownItems | Sort-Object Text
            $script:CustomThemes.DropDownItems.Clear()
            $script:CustomThemes.DropDownItems.AddRange($SortedItems)

            $MenuColorTheme.DropDownItems.Insert(0, $script:CustomThemes) | Out-Null

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
        $script:ThemeBuilderDeleteThemesForm.Controls.Add($script:DeleteThemeCloseButton)
        $script:ThemeBuilderDeleteThemesForm.Controls.Add($script:ThemeBuilderThemeNameListBoxLabel)
        $script:ThemeBuilderDeleteThemesForm.Controls.Add($script:ThemeBuilderThemeNameListBox)
        $script:ThemeBuilderDeleteThemesForm.Controls.Add($script:ThemeBuilderDeleteThemesPopupButton)

        # Show form
        $script:ThemeBuilderDeleteThemesForm.ShowDialog() | Out-Null
    })

    # Add controls to the Theme Builder form
    $script:ThemeBuilderForm.Controls.Add($script:ThemeBuilderMinimizeButton)
    $script:ThemeBuilderForm.Controls.Add($script:ThemeBuilderCloseButton)
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
    $script:ThemeBuilderForm.Controls.Add($script:ThemeBuilderColorPickerButton)
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

    Set-ThemeBuilderDefaultColors

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

        # Close Color Picker as well if it's open
        if ($global:IsColorPickerPopupActive) {
            $script:ColorPickerPopup.Close()
        }
    })
})

# Click event handler for the Font Picker popup menu
$MenuFontPicker.add_Click({
    if ($global:IsFontPickerPopupActive) {
        $OutText.AppendText("$(Get-Timestamp) - Font Picker is already open.`r`n")
        $script:FontPickerPopup.Activate()
        return
    }
    $OutText.AppendText("$(Get-Timestamp) - Launching Font Picker...`r`n")

    # Launch the Font Picker form
    $script:FontPickerPopup = New-WindowsForm -SizeX 380 -SizeY 635 -Text '' -StartPosition 'CenterScreen' -TopMost $false -ShowInTaskbar $true -KeyPreview $true -MinimizeBox $true
    $global:IsFontPickerPopupActive = $true

    # Create a panel with scrollbars
    $script:FontPickerPanel = New-Object System.Windows.Forms.Panel
    $script:FontPickerPanel.Location = New-Object System.Drawing.Point(5, 40)
    $script:FontPickerPanel.Size = New-Object System.Drawing.Size(370, 540)
    $script:FontPickerPanel.AutoScroll = $true

    # Button for applying color choices
    $script:FontPickerApplyButton = New-FormButton -Text "Apply Choices" -LocationX 140 -LocationY 10 -Width 100 -BackColor $DisabledBackColor -ForeColor $DisabledForeColor -Font $global:NormalFont -Enabled $false -Visible $true

    # Create a list box to hold the font names
    $script:FontPickerListBox = New-FormListBox -LocationX 0 -LocationY 0 -SizeX 370 -SizeY 540 -Font $Font -SelectionMode 'One'
    $script:FontPickerListBox.ItemHeight = 20
    $script:FontPickerListBox.DrawMode = 'OwnerDrawFixed'

    # Check if DefaultUserTheme has a value or is null
    if ($null -ne $ConfigValues.DefaultUserTheme -and $ConfigValues.DefaultUserTheme -ne '') {
        # Get theme
        $SelectedTheme = $ColorTheme.PSObject.Properties | Where-Object { $_.Value.$($ConfigValues.DefaultUserTheme) } | Select-Object -ExpandProperty Name
        $themeColors = $ColorTheme.$SelectedTheme.$($ConfigValues.DefaultUserTheme)
        
        $script:FontPickerPopup.BackColor = $themeColors.BackColor
        $script:FontPickerPanel.BackColor = $themeColors.BackColor

        if ($SelectedTheme -eq 'Premium' -or $SelectedTheme -eq 'Custom') {
            $script:FontPickerListBox.BackColor = $themeColors.AccentColor
        }
        else {
            $script:FontPickerListBox.BackColor = [System.Drawing.SystemColors]::Control
        }

        if ($ConfigValues.DefaultUserTheme -eq 'USA') {
            Enable-USAThemeTextColor
        }
    }

    # Event handler for drawing each item
    $script:FontPickerListBox.Add_DrawItem({
        param($FontPickerSender, $e)

        # Draw the background
        $e.DrawBackground()

        # Create a font from the list item
        try {
            $fontName = $FontPickerSender.Items[$e.Index]
            $font = New-Object System.Drawing.Font($fontName, 10, [System.Drawing.FontStyle]::Regular)
        } catch {
            # If there's an error creating the font, use the default font
            $font = $e.Font
        }

        # Set the brush for drawing the font name
        $brush = New-Object System.Drawing.SolidBrush($e.ForeColor)

        # Draw the font name
        $e.Graphics.DrawString($fontName, $font, $brush, $e.Bounds.X, $e.Bounds.Y)

        # Draw a focus rectangle if the item has focus
        $e.DrawFocusRectangle()
    })

    # Define a list of known symbol font names
    $SymbolFontNames = @(
        'Bookshelf Symbol 7',
        'Marlett', 
        'Webdings', 
        'Wingdings', 
        'Wingdings 2', 
        'Wingdings 3', 
        'Symbol', 
        'ZapfDingbats', 
        'HoloLens MDL2 Assets', 
        'MS Outlook', 
        'MS Reference Specialty', 
        'MT Extra', 
        'Segoe MDL2 Assets', 
        'SimSun-ExtB'
    )

    # Populate the list box with font names, excluding symbol fonts
    foreach ($fontFamily in [System.Drawing.FontFamily]::Families) {
        try {
            $fontName = $fontFamily.Name

            # Skip symbol fonts
            if ($fontName -in $SymbolFontNames) {
                continue
            }

            # Check if the font can be displayed with GDI
            $fontTest = New-Object System.Drawing.Font($fontName, 10, [System.Drawing.FontStyle]::Regular)
            if ($fontTest.GdiCharSet -ne 0) {
                $script:FontPickerListBox.Items.Add($fontName) | Out-Null
            }
        } catch {
            # If there's an error creating the font, skip it
            continue
        }
    }

    # Event handler for checking if a font is selected
    $script:FontPickerListBox.add_SelectedIndexChanged({
        Enable-FontPickerApplyChoiceButton
    })

    # Button click event handler for applying font choice
    $script:FontPickerApplyButton.add_Click({
        $SelectedFont = $script:FontPickerListBox.SelectedItem

        # Load the JSON content from ColorThemes.json
        $jsonContent = Get-Content -Path .\Config.json -Raw | ConvertFrom-Json

        # Update the DefaultUserFont with the selected item from the ListBox
        $jsonContent.DefaultUserFont = $SelectedFont #$script:FontPickerListBox.SelectedItem.ToString()

        # Convert the updated object back to JSON and save it to ColorThemes.json
        $jsonContent | ConvertTo-Json -Depth 100 | Set-Content -Path .\Config.json

        $global:DefaultUserFont = $SelectedFont

        $global:NormalFont = [System.Drawing.Font]::new($global:DefaultUserFont, 9, [System.Drawing.FontStyle]::Regular)
        $global:NormalSmallFont = [System.Drawing.Font]::new($global:DefaultUserFont, 8, [System.Drawing.FontStyle]::Regular)
        $global:NormalBoldFont = [System.Drawing.Font]::new($global:DefaultUserFont, 9, [System.Drawing.FontStyle]::Bold)
        $global:SmallBoldFont = [System.Drawing.Font]::new($global:DefaultUserFont, 8, [System.Drawing.FontStyle]::Bold)
        $global:ExtraSmallFont = [System.Drawing.Font]::new($global:DefaultUserFont, 7, [System.Drawing.FontStyle]::Regular)
        $global:ExtraSmallItalicBoldFont = [System.Drawing.Font]::new($global:DefaultUserFont, 7, [System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Italic)
        $global:ExtraLargeFont = [System.Drawing.Font]::new($global:DefaultUserFont, 16, [System.Drawing.FontStyle]::Regular)

        # Create a hashtable for font size and style mapping
        $FontMapping = @{
            "9.0-Regular" = $global:NormalFont
            "8.0-Regular" = $global:NormalSmallFont
            "9.0-Bold" = $global:NormalBoldFont
            "8.0-Bold" = $global:SmallBoldFont
            "7.0-Regular" = $global:ExtraSmallFont
            "7.0-Bold, Italic" = $global:ExtraSmallItalicBoldFont
            "16.0-Regular" = $global:ExtraLargeFont
        }        

        # Update the font of the form and all child controls
        Update-FormFonts -Control $DesktopAssistantForm -FontMapping $FontMapping

        # Update the font for other forms if they're open
        if ($global:IsLenderLFPPopupActive) {
            Update-FormFonts -Control $script:LenderLFPPopup -FontMapping $FontMapping
            $script:LenderLFPPopup.Refresh()
        }
        if ($global:IsBillingRestartPopupActive) {
            Update-FormFonts -Control $script:BillingRestartPopup -FontMapping $FontMapping
            $script:BillingRestartPopup.Refresh()
        }
        if ($global:IsHDTStoragePopupActive) {
            Update-FormFonts -Control $script:HDTStoragePopup -FontMapping $FontMapping
            $script:HDTStoragePopup.Refresh()
        }
        if ($global:IsFeedbackPopupActive) {
            Update-FormFonts -Control $script:FeedbackPopup -FontMapping $FontMapping
            $script:FeedbackPopup.Refresh()
        }
        if ($global:IsThemeBuilderPopupActive) {
            Update-FormFonts -Control $script:ThemeBuilderForm -FontMapping $FontMapping
            $script:ThemeBuilderForm.Refresh()
        }
        if ($global:IsColorPickerPopupActive) {
            Update-FormFonts -Control $script:ColorPickerPopup -FontMapping $FontMapping
            $script:ColorPickerPopup.Refresh()
        }
        if ($global:IsCertCheckPopupActive) {
            Update-FormFonts -Control $script:CertCheckPopup -FontMapping $FontMapping
            $script:CertCheckPopup.Refresh()
        }
        if ($global:IsFontPickerPopupActive) {
            Update-FormFonts -Control $script:FontPickerPopup -FontMapping $FontMapping
            $script:FontPickerPopup.Refresh()
        }

        $DesktopAssistantForm.Refresh()

        $OutText.AppendText("$(Get-Timestamp) - Font changed to: $SelectedFont`r`n")
    })

    # Event handler for form close
    $script:FontPickerPopup.Add_FormClosed({
        $global:IsFontPickerPopupActive = $false
    })

    # Add controls to the form
    $script:FontPickerPopup.Controls.Add($script:FontPickerApplyButton)
    $script:FontPickerPopup.Controls.Add($script:FontPickerListBox)
    $script:FontPickerPopup.Controls.Add($script:FontPickerPanel)
    $script:FontPickerPanel.Controls.Add($script:FontPickerListBox)

    # Show the form
    $script:FontPickerPopup.Show() | Out-Null
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

# Click event for the Options menu Show Help Icons option
$ShowHelpIconsMenu.add_Click({
    if ($ShowHelpIconsMenu.Text -eq "Show Help Icons") {
        $SysAdminTab.Controls.Add($ServerPingHelpIcon)
        $SysAdminTab.Controls.Add($NSLookupHelpIcon)
        $SysAdminTab.Controls.Add($ReverseIPHelpIcon)
        $SysAdminTab.Controls.Add($RestartsGUIHelpIcon)
        $SysAdminTab.Controls.Add($CertCheckHelpIcon)
        $AWSAdminTab.Controls.Add($AWSMonitoringGUIHelpIcon)
        $SupportTab.Controls.Add($PSTHelpIcon)
        $SupportTab.Controls.Add($LFPWizardHelpIcon)
        $SupportTab.Controls.Add($BillingRestartHelpIcon)
        $SupportTab.Controls.Add($PWManagerHelpIcon)
        $SupportTab.Controls.Add($GenPWHelpIcon)
        $SupportTab.Controls.Add($HDTStorageHelpIcon)
        $SupportTab.Controls.Add($NewDocHelpIcon)
        $TicketManagerTab.Controls.Add($NewTicketHelpIcon)
        $TicketManagerTab.Controls.Add($RenameTicketHelpIcon)
        $ServerPingHelpIcon.BringToFront()
        $ConfigValues.HelpIcons = "Enabled"
        $UpdatedHelpIconValue = ConvertTo-Json -InputObject $ConfigValues -Depth 100
        Set-Content -Path .\Config.json -Value $UpdatedHelpIconValue
        $ShowHelpIconsMenu.Text = "Hide Help Icons"
        $OutText.AppendText("$(Get-Timestamp) - Help icons have been enabled.`r`n")
    }
    else {
        $SysAdminTab.Controls.Remove($ServerPingHelpIcon)
        $SysAdminTab.Controls.Remove($NSLookupHelpIcon)
        $SysAdminTab.Controls.Remove($ReverseIPHelpIcon)
        $SysAdminTab.Controls.Remove($RestartsGUIHelpIcon)
        $SysAdminTab.Controls.Remove($CertCheckHelpIcon)
        $AWSAdminTab.Controls.Remove($AWSMonitoringGUIHelpIcon)
        $SupportTab.Controls.Remove($PSTHelpIcon)
        $SupportTab.Controls.Remove($LFPWizardHelpIcon)
        $SupportTab.Controls.Remove($BillingRestartHelpIcon)
        $SupportTab.Controls.Remove($PWManagerHelpIcon)
        $SupportTab.Controls.Remove($GenPWHelpIcon)
        $SupportTab.Controls.Remove($HDTStorageHelpIcon)
        $SupportTab.Controls.Remove($NewDocHelpIcon)
        $TicketManagerTab.Controls.Remove($NewTicketHelpIcon)
        $TicketManagerTab.Controls.Remove($RenameTicketHelpIcon)
        $ConfigValues.HelpIcons = "Disabled"
        $UpdatedHelpIconValue = ConvertTo-Json -InputObject $ConfigValues -Depth 100
        Set-Content -Path .\Config.json -Value $UpdatedHelpIconValue
        $ShowHelpIconsMenu.Text = "Show Help Icons"
        $OutText.AppendText("$(Get-Timestamp) - Help icons have been disabled.`r`n")
    }
})

# Click event for the GitHub Repo menu option
$MenuGitHub.add_Click({ Start-Process "https://github.com/jthamind/DesktopAssistant" })

# Click event for the About menu option
$MenuAboutItem.add_Click({
    $OutText.AppendText("$(Get-Timestamp) - Launching About form...`r`n")

    # Create the About form
    $script:AboutForm = New-WindowsForm -SizeX 400 -SizeY 300 -Text '' -StartPosition 'CenterScreen' -TopMost $false -ShowInTaskbar $true -KeyPreview $true -MinimizeBox $false

    # Close About form when deactivated
    $script:AboutForm.Add_Deactivate({
        $script:AboutForm.Close()
    })

    # Mouse Down event handler for moving the About form
    # This is needed since the form has no ControlBox
    $script:AboutForm_MouseDown = {
        $global:MouseDown = $true
        $global:MouseClickPoint = [System.Windows.Forms.Cursor]::Position
    }

    # Mouse Down event handler for moving the About form
    # This is needed since the form has no ControlBox
    $script:AboutForm_MouseMove = {
        if ($global:MouseDown) {
            $CurrentCursorPosition = [System.Windows.Forms.Cursor]::Position
            $FormLocation = $script:AboutForm.Location
            
            $newX = $FormLocation.X + ($CurrentCursorPosition.X - $global:MouseClickPoint.X)
            $newY = $FormLocation.Y + ($CurrentCursorPosition.Y - $global:MouseClickPoint.Y)
            
            $script:AboutForm.Location = New-Object System.Drawing.Point($newX, $newY)
            $global:MouseClickPoint = $CurrentCursorPosition
        }
    }

    # Mouse Down event handler for moving the About form
    # This is needed since the form has no ControlBox
    $script:AboutForm_MouseUp = {
        $global:MouseDown = $false
    }

    # Add event handlers to the form
    $script:AboutForm.Add_MouseDown($script:AboutForm_MouseDown)
    $script:AboutForm.Add_MouseMove($script:AboutForm_MouseMove)
    $script:AboutForm.Add_MouseUp($script:AboutForm_MouseUp)

    # Picture box for Allied Solutions logo
    $AboutMenuAlliedLogo = New-FormPictureBox -LocationX 50 -LocationY 30 -SizeX 275 -SizeY 100 -Image $AboutMenuAlliedLogoImage -SizeMode $ZoomSizeMode -Cursor $ArrowCursor

    # Label for About form
    $AboutLabel = New-FormLabel -LocationX 115 -LocationY 150 -SizeX 150 -SizeY 20 -Text 'ETG Desktop Assistant' -Font $global:NormalBoldFont -TextAlign $DefaultTextAlign

    # Label for version numbers
    $VersionLabel = New-FormLabel -LocationX 145 -LocationY 175 -SizeX 150 -SizeY 20 -Text "Version: 3.0.0" -Font $global:NormalFont -TextAlign $DefaultTextAlign

    # Label for Allied IMPACT
    $AlliedImpactLabel = New-FormLabel -LocationX 40 -LocationY 200 -SizeX 300 -SizeY 40 -Text "Made with Passion to make an IMPACT`r`nat Allied Solutions, LLC." -Font $global:NormalItalicBoldFont -TextAlign $MiddleTextAlign

    # About form instruction label
    $AboutFormInstructionLabel = New-FormLabel -LocationX 75 -LocationY 260 -SizeX 180 -SizeY 40 -Text "Click anywhere outside this popup to close it." -Font $global:SmallItalicFont -TextAlign $MiddleTextAlign
    $AboutFormInstructionLabel.AutoSize = $true
    $AboutFormInstructionLabel.Padding = New-Object System.Windows.Forms.Padding(10)

    # Check if DefaultUserTheme has a value or is null
    if ($null -ne $ConfigValues.DefaultUserTheme -and $ConfigValues.DefaultUserTheme -ne '') {
        # Get theme
        $SelectedTheme = $ColorTheme.PSObject.Properties | Where-Object { $_.Value.$($ConfigValues.DefaultUserTheme) } | Select-Object -ExpandProperty Name
        $themeColors = $ColorTheme.$SelectedTheme.$($ConfigValues.DefaultUserTheme)
        
		$script:AboutForm.BackColor = $themeColors.BackColor
		$script:AboutForm.ForeColor = $themeColors.ForeColor
        $AboutLabel.BackColor = $themeColors.BackColor
        $AboutLabel.ForeColor = $themeColors.ForeColor
    }

    $script:AboutForm.Controls.Add($AboutLabel)
    $script:AboutForm.Controls.Add($AboutMenuAlliedLogo)
    $script:AboutForm.Controls.Add($VersionLabel)
    $script:AboutForm.Controls.Add($AlliedImpactLabel)
    $script:AboutForm.Controls.Add($AboutFormInstructionLabel)
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
$RestartsTabControl.AutoSize = $false

# Individual servers list box
$ServersListBox = New-FormListBox -LocationX 5 -LocationY 95 -SizeX 200 -SizeY 240 -Font $global:NormalFont -SelectionMode 'One'

# Services list box
$ServicesListBox = New-FormListBox -LocationX 0 -LocationY 0 -SizeX 245 -SizeY 240 -Font $global:NormalFont -SelectionMode 'MultiExtended'

# IIS Sites list box
$IISSitesListBox = New-FormListBox -LocationX 0 -LocationY 0 -SizeX 245 -SizeY 240 -Font $global:NormalFont -SelectionMode 'MultiExtended'

# IIS App Pools list box
$AppPoolsListBox = New-FormListBox -LocationX 0 -LocationY 0 -SizeX 245 -SizeY 240 -Font $global:NormalFont -SelectionMode 'MultiExtended'

# Combobox for application selection
$AppListCombo = New-FormComboBox -LocationX 5 -LocationY 65 -SizeX 200 -SizeY 200 -Font $global:NormalFont

# Label applist combo box
$AppListLabel = New-FormLabel -LocationX 5 -LocationY 40 -SizeX 95 -SizeY 20 -Text 'Select a Server' -Font $global:NormalBoldFont -TextAlign $DefaultTextAlign

# Help Icon for restarts GUI
$RestartsGUIHelpIcon = New-FormPictureBox -LocationX 100 -LocationY 40 -SizeX 20 -SizeY 20 -Image $HelpIconImage -SizeMode $AutoSizeMode -Cursor $HandCursor

# Event handler for the Restarts GUI Help Icon
$RestartsGUIHelpIcon.add_Click({
    Show-HelpForm -PictureBox $RestartsGUIHelpIcon -HelpText "Start, Stop, and Restart one or more Windows services, Websites, Application Pools, or IIS itself on the chosen server. The servers.csv file used for this module can be found at .\servers.csv."
})

# Tab for services list
$ServicesTab = New-FormTabPage -Font $global:NormalFont -Name "ServicesTab" -Text "Services"

# Tab for IIS sites list
$IISSitesTab = New-FormTabPage -Font $global:NormalFont -Name "IISSitesTab" -Text "IIS Sites"

# Tab for IIS App Pools list
$AppPoolsTab = New-FormTabPage -Font $global:NormalFont -Name "AppPoolsTab" -Text "App Pools"

# Button for restarting services
$script:RestartButton = New-FormButton -Text "Restart" -LocationX 490 -LocationY 95 -Width 75 -BackColor $global:DisabledBackColor -ForeColor $global:DisabledForeColor -Font $global:NormalFont -Enabled $false -Visible $true

# Button for starting services
$script:StartButton = New-FormButton -Text "Start" -LocationX 490 -LocationY 125 -Width 75 -BackColor $global:DisabledBackColor -ForeColor $global:DisabledForeColor -Font $global:NormalFont -Enabled $false -Visible $true

# Button for stopping services
$script:StopButton = New-FormButton -Text "Stop" -LocationX 490 -LocationY 155 -Width 75 -BackColor $global:DisabledBackColor -ForeColor $global:DisabledForeColor -Font $global:NormalFont -Enabled $false -Visible $true

# Button for for opening IIS Site in Windows Explorer
$script:OpenSiteButton = New-FormButton -Text "Open" -LocationX 490 -LocationY 185 -Width 75 -BackColor $global:DisabledBackColor -ForeColor $global:DisabledForeColor -Font $global:NormalFont -Enabled $false -Visible $false

# Button for restarting IIS on server
$script:RestartIISButton = New-FormButton -Text "Restart IIS" -LocationX 490 -LocationY 215 -Width 75 -BackColor $global:DisabledBackColor -ForeColor $global:DisabledForeColor -Font $global:NormalFont -Enabled $false -Visible $false

# Button for starting IIS on server
$script:StartIISButton = New-FormButton -Text "Start IIS" -LocationX 490 -LocationY 245 -Width 75 -BackColor $global:DisabledBackColor -ForeColor $global:DisabledForeColor -Font $global:NormalFont -Enabled $false -Visible $false

# Button for stopping IIS on server
$script:StopIISButton = New-FormButton -Text "Stop IIS" -LocationX 490 -LocationY 275 -Width 75 -BackColor $global:DisabledBackColor -ForeColor $global:DisabledForeColor -Font $global:NormalFont -Enabled $false -Visible $false

# Separator line under Restarts GUI
$RestartsSeparator = New-FormLabel -LocationX 0 -LocationY 340 -SizeX 1000 -SizeY 2 -Text '' -Font $global:NormalFont -TextAlign $DefaultTextAlign

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

$script:RestartButton.Add_Click({
    $Action = 'Restart'
    $SelectedServer = $ServersListBox.SelectedItem
    $SelectedTab = $RestartsTabControl.SelectedTab.Text

    if ($SelectedTab -match "Services|IIS Sites|App Pools") {
        $listBox = switch ($SelectedTab) {
            'Services' { $ServicesListBox }
            'IIS Sites' { $IISSitesListBox }
            'App Pools' { $AppPoolsListBox }
        }
    }
    $SelectedItems = @()
    foreach ($item in $listBox.SelectedItems) {
        $SelectedItems += $item
    }
    
    Open-RestartItemsRunspace -ConfigValuesRestartItemsScript $ConfigValues.RestartItemsScript -Action $Action -SelectedServer $SelectedServer -SelectedTab $SelectedTab -SelectedItems $SelectedItems -OutText $OutText -TimestampFunction ${function:Get-Timestamp}
})

# Button click event handler for starting services, IIS sites, or app pools in an async runspace pool
$script:StartButton.Add_Click({
    $Action = 'Start'
    $SelectedServer = $ServersListBox.SelectedItem
    $SelectedTab = $RestartsTabControl.SelectedTab.Text

    if ($SelectedTab -match "Services|IIS Sites|App Pools") {
        $listBox = switch ($SelectedTab) {
            'Services' { $ServicesListBox }
            'IIS Sites' { $IISSitesListBox }
            'App Pools' { $AppPoolsListBox }
        }
    }
    $SelectedItems = @()
    foreach ($item in $listBox.SelectedItems) {
        $SelectedItems += $item
    }

    Open-RestartItemsRunspace -ConfigValuesStartItemsScript $ConfigValues.StartItemsScript -Action $Action -SelectedServer $SelectedServer -SelectedTab $SelectedTab -SelectedItems $SelectedItems -OutText $OutText -TimestampFunction ${function:Get-Timestamp}
})

# Button click event handler for stopping one or more services, IIS sites, or app pools in an async runspace pool
$script:StopButton.Add_Click({
    $Action = 'Stop'
    $SelectedServer = $ServersListBox.SelectedItem
    $SelectedTab = $RestartsTabControl.SelectedTab.Text

    if ($SelectedTab -match "Services|IIS Sites|App Pools") {
        $listBox = switch ($SelectedTab) {
            'Services' { $ServicesListBox }
            'IIS Sites' { $IISSitesListBox }
            'App Pools' { $AppPoolsListBox }
        }
    }
    $SelectedItems = @()
    foreach ($item in $listBox.SelectedItems) {
        $SelectedItems += $item
    }

    Open-RestartItemsRunspace -ConfigValuesStopItemsScript $ConfigValues.StopItemsScript -Action $Action -SelectedServer $SelectedServer -SelectedTab $SelectedTab -SelectedItems $SelectedItems -OutText $OutText -TimestampFunction ${function:Get-Timestamp}
})

# Button click event handler for opening one or more IIS site directories in an async runspace pool
$script:OpenSiteButton.Add_Click({
    $SelectedServer = $ServersListBox.SelectedItem
    $SelectedTab = $RestartsTabControl.SelectedTab.Text
    $synchash.SelectedTab = $SelectedTab
    
    $runspacePool = [runspacefactory]::CreateRunspacePool(1, 10)
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
            } -ArgumentList $item -Authentication Negotiate

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

    Open-RestartIISRunspace -SelectedServer $SelectedServer -OutText $OutText -TimestampFunction ${function:Get-Timestamp}
})

# Event handler for restarting IIS on a server
$script:StartIISButton.Add_Click({
    $SelectedServer = $ServersListBox.SelectedItem

    Open-StartIISRunspace -SelectedServer $SelectedServer -OutText $OutText -TimestampFunction ${function:Get-Timestamp}
})

# Event handler for restarting IIS on a server
$script:StopIISButton.Add_Click({
    $SelectedServer = $ServersListBox.SelectedItem

    Open-StopIISRunspace -SelectedServer $SelectedServer -OutText $OutText -TimestampFunction ${function:Get-Timestamp}
})

# Add the event handler to the $ServersListBox
$ServersListBox.add_SelectedIndexChanged({ OnServerSelected })

# Store the script block in a variable
$OnServiceSelectedAction = { OnServiceSelected -TimestampFunction ${function:Get-Timestamp} }

# Add the event using the script block
$ServicesListBox.add_SelectedIndexChanged($OnServiceSelectedAction)

# Add the event handler to the $IISSitesListBox
$IISSitesListBox.add_SelectedIndexChanged({ OnIISSiteSelected -TimestampFunction ${function:Get-Timestamp} })

# Add the event handler to the $AppPoolsListBox
$AppPoolsListBox.add_SelectedIndexChanged({ OnAppPoolSelected -TimestampFunction ${function:Get-Timestamp} })

# Event handler for TabControl's SelectedIndexChanged event
$RestartsTabControl_SelectedIndexChanged = {
    $ServicesListBox.ClearSelected()
    $IISSitesListBox.ClearSelected()
    $AppPoolsListBox.ClearSelected()
    $RestartsTabControl.SelectedTab.Text
    Open-PopulateListBoxRunspace -OutText $OutText -RestartsTabControl $RestartsTabControl -ServersListBox $ServersListBox -ServicesListBox $ServicesListBox -IISSitesListBox $IISSitesListBox -AppPoolsListBox $AppPoolsListBox -TimestampFunction ${function:Get-Timestamp}

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

# Button to run NSLookup
$NSLookupButton = New-FormButton -Text "Get IP" -LocationX 300 -LocationY 375 -Width 50 -BackColor $global:DisabledBackColor -ForeColor $global:DisabledForeColor -Font $global:NormalFont -Enabled $false -Visible $true

# Text box for NSLookup
$script:NSLookupTextBox = New-FormTextBox -LocationX 195 -LocationY 375 -SizeX 100 -SizeY 20 -ScrollBars 'None' -Multiline $false -Enabled $true -ReadOnly $false -Text '' -Font $global:NormalFont

# Label for NSLookup text box
$NSLookupLabel = New-FormLabel -LocationX 195 -LocationY 350 -SizeX 65 -SizeY 20 -Text 'NSlookup' -Font $global:NormalBoldFont -TextAlign $DefaultTextAlign

# Help icon for NSLookup text box
$NSLookupHelpIcon = New-FormPictureBox -LocationX 260 -LocationY 350 -SizeX 20 -SizeY 20 -Image $HelpIconImage -SizeMode $AutoSizeMode -Cursor $HandCursor

# Event handler for NSLookup Help Icon
$NSLookupHelpIcon.Add_Click({
    Show-HelpForm -PictureBox $NSLookupHelpIcon -HelpText "Use PowerShell's native Resolve-DnsName cmdlet to get the IPv4/IPv6 address of a given server. Accepts multiple DNS names, each separated by commas."
})

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
    Open-NSLookupRunspace -OutText $OutText -NSLookupTextBox $script:NSLookupTextBox -TimestampFunction ${function:Get-Timestamp}
})

# Even handler for the RunLookup text box
$script:NSLookupTextBox.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") {
        Open-NSLookupRunspace -OutText $OutText -NSLookupTextBox $script:NSLookupTextBox -TimestampFunction ${function:Get-Timestamp}
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
$ServerPingButton = New-FormButton -Text "Ping" -LocationX 110 -LocationY 375 -Width 45 -BackColor $global:DisabledBackColor -ForeColor $global:DisabledForeColor -Font $global:NormalFont -Enabled $false -Visible $true

# Text box for testing server connection
$ServerPingTextBox = New-FormTextBox -LocationX 5 -LocationY 375 -SizeX 100 -SizeY 20 -ScrollBars 'None' -Multiline $false -Enabled $true -ReadOnly $false -Text '' -Font $global:NormalFont

# Label for testing server connection text box
$ServerPingLabel = New-FormLabel -LocationX 5 -LocationY 350 -SizeX 150 -SizeY 20 -Text 'Ping a Server' -Font $global:NormalBoldFont -TextAlign $DefaultTextAlign

# Help icon for testing server connection
$ServerPingHelpIcon = New-FormPictureBox -LocationX 90 -LocationY 350 -SizeX 20 -SizeY 20 -Image $HelpIconImage -SizeMode $AutoSizeMode -Cursor $HandCursor

# Event handler for Server Ping Help Icon
$ServerPingHelpIcon.Add_Click({
    Show-HelpForm -PictureBox $ServerPingHelpIcon -HelpText "Use PowerShell's native Test-Connection cmdlet to ping a server three times with a 1000ms timeout duration. Accepts multiple server names or IP addresses, each separated by commas."
})

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
    Open-ServerPingRunspace -OutText $OutText -ServerPingTextBox $ServerPingTextBox -TimestampFunction ${function:Get-Timestamp}
})

# Server Ping text box Enter key logic
$ServerPingTextBox.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") {
        Open-ServerPingRunspace -OutText $OutText -ServerPingTextBox $ServerPingTextBox -TimestampFunction ${function:Get-Timestamp}
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

# Button for Reverse IP Lookup
$ReverseIPButton = New-FormButton -Text "Get DNS" -LocationX 510 -LocationY 375 -Width 65 -BackColor $global:DisabledBackColor -ForeColor $global:DisabledForeColor -Font $global:NormalFont -Enabled $false -Visible $true

# Text box for Reverse IP Lookup
$ReverseIPTextBox = New-FormTextBox -LocationX 400 -LocationY 375 -SizeX 100 -SizeY 20 -ScrollBars 'None' -Multiline $false -Enabled $true -ReadOnly $false -Text '' -Font $global:NormalFont

# Label for Reverse IP Lookup text box
$ReverseIPLabel = New-FormLabel -LocationX 400 -LocationY 350 -SizeX 65 -SizeY 20 -Text 'IP Lookup' -Font $global:NormalBoldFont -TextAlign $DefaultTextAlign

# Help Icon for Reverse IP Lookup
$ReverseIPHelpIcon = New-FormPictureBox -LocationX 465 -LocationY 350 -SizeX 20 -SizeY 20 -Image $HelpIconImage -SizeMode $AutoSizeMode -Cursor $HandCursor

# Event handler for Reverse IP Lookup Help Icon
$ReverseIPHelpIcon.Add_Click({
    Show-HelpForm -PictureBox $ReverseIPHelpIcon -HelpText "Perform a reverse IP lookup using the System.Net.Dns.GetHostEntryAsync method to resolve an IP address to a hostname. Accepts multiple IP addresses, each separated by commas."
})

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
    Open-ReverseIPLookupRunspace -ip $ip -OutText $OutText -TimestampFunction ${function:Get-Timestamp}
})

# Event handler for pressing Enter in the Reverse IP Lookup text box
$ReverseIPTextBox.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") {
        $ip = $ReverseIPTextBox.Text
        $ReverseIPTextBox.Text = ''
        Open-ReverseIPLookupRunspace -ip $ip -OutText $OutText -TimestampFunction ${function:Get-Timestamp}
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
? START OF CERT CHECK
? **********************************************************************************************************************
#>

# Button for launching Cert Check Wizard
$CertCheckWizardButton = New-FormButton -Text "Launch CC Wizard" -LocationX 5 -LocationY 440 -Width 150 -BackColor $ControlColor -ForeColor $ControlColor -Font $global:NormalFont -Enabled $true -Visible $true

# Label for Cert Check Wizard
$CertCheckLabel = New-FormLabel -LocationX 27 -LocationY 415 -SizeX 108 -SizeY 20 -Text 'Run Cert Checks' -Font $global:NormalBoldFont -TextAlign $DefaultTextAlign

# Help Icon for Cert Check Wizard
$CertCheckHelpIcon = New-FormPictureBox -LocationX 135 -LocationY 415 -SizeX 20 -SizeY 20 -Image $HelpIconImage -SizeMode $AutoSizeMode -Cursor $HandCursor

# Event handler for Cert Check Help Icon
$CertCheckHelpIcon.Add_Click({
    Show-HelpForm -PictureBox $CertCheckHelpIcon -HelpText "Use the included servers.csv file or enter one or more server names to check SSL certificates bound in IIS and their expiration dates. Results can be output to the main logging textbox or saved to a file, and servers with no results can optionally be ignored."
})

# Event handler for launching Cert Check Wizard
$CertCheckWizardButton.add_Click({
    if ($global:IsCertCheckPopupActive -eq $true) {
        $OutText.AppendText("$(Get-Timestamp) - Cert Check Wizard is already open.`r`n")
        $CertCheckPopup.Activate()
        return
    }
    $OutText.AppendText("$(Get-Timestamp) - Launching Cert Check Wizard...`r`n")

    # Create the Cert Check popup form
    $script:CertCheckPopup = New-WindowsForm -SizeX 300 -SizeY 600 -Text '' -StartPosition 'CenterScreen' -TopMost $false -ShowInTaskbar $true -KeyPreview $true -MinimizeBox $true
    $global:IsCertCheckPopupActive = $true

    # Mouse Down event handler for moving the Theme Builder form
    # This is needed since the form has no ControlBox
    $script:CertCheckPopup_MouseDown = {
        $global:MouseDown = $true
        $global:MouseClickPoint = [System.Windows.Forms.Cursor]::Position
    }

    # Mouse Down event handler for moving the Theme Builder form
    # This is needed since the form has no ControlBox
    $script:CertCheckPopup_MouseMove = {
        if ($global:MouseDown) {
            $CurrentCursorPosition = [System.Windows.Forms.Cursor]::Position
            $FormLocation = $script:CertCheckPopup.Location
            
            $newX = $FormLocation.X + ($CurrentCursorPosition.X - $global:MouseClickPoint.X)
            $newY = $FormLocation.Y + ($CurrentCursorPosition.Y - $global:MouseClickPoint.Y)
            
            $script:CertCheckPopup.Location = New-Object System.Drawing.Point($newX, $newY)
            $global:MouseClickPoint = $CurrentCursorPosition
        }
    }

    # Mouse Down event handler for moving the Theme Builder form
    # This is needed since the form has no ControlBox
    $script:CertCheckPopup_MouseUp = {
        $global:MouseDown = $false
    }

    # Add event handlers to the form
    $script:CertCheckPopup.Add_MouseDown($script:CertCheckPopup_MouseDown)
    $script:CertCheckPopup.Add_MouseMove($script:CertCheckPopup_MouseMove)
    $script:CertCheckPopup.Add_MouseUp($script:CertCheckPopup_MouseUp)

    # Button for minimizing the Save Theme popup form
    $script:CertCheckMinimizeButton = New-FormButton -Text "─" -LocationX 220 -LocationY 0 -Width 40 -BackColor $themeColors.BackColor -ForeColor $themeColors.ForeColor -Font $global:NormalBoldFont -Enabled $true -Visible $true
    $script:CertCheckMinimizeButton.Add_Click({ $script:CertCheckPopup.WindowState = 'Minimized' })
    $script:CertCheckMinimizeButton.FlatStyle = 'Flat'
    $script:CertCheckMinimizeButton.FlatAppearance.BorderSize = 0
    $script:CertCheckMinimizeButton.FlatAppearance.MouseOverBackColor = if ($SelectedTheme -eq 'Premium' -or $SelectedTheme -eq 'Custom') { $themeColors.AccentColor } else { $themeColors.ForeColor }
    $script:CertCheckMinimizeButton.Add_MouseEnter({ if ($SelectedTheme -eq 'Premium' -or $SelectedTheme -eq 'Custom') { $script:CertCheckMinimizeButton.ForeColor = $themeColors.ForeColor } else { $script:CertCheckMinimizeButton.ForeColor = $themeColors.BackColor } })
    $script:CertCheckMinimizeButton.Add_MouseLeave({ $script:CertCheckMinimizeButton.ForeColor = $themeColors.ForeColor })

    # Button for closing the Save Theme popup form
    $script:CertCheckCloseButton = New-FormButton -Text "X" -LocationX 260 -LocationY 0 -Width 40 -BackColor $themeColors.BackColor -ForeColor $themeColors.ForeColor -Font $global:NormalBoldFont -Enabled $true -Visible $true
    $script:CertCheckCloseButton.Add_Click({ $script:CertCheckPopup.Close() })
    $script:CertCheckCloseButton.FlatStyle = 'Flat'
    $script:CertCheckCloseButton.FlatAppearance.BorderSize = 0
    $script:CertCheckCloseButton.FlatAppearance.MouseOverBackColor = if ($SelectedTheme -eq 'Premium' -or $SelectedTheme -eq 'Custom') { $themeColors.AccentColor } else { $themeColors.ForeColor }
    $script:CertCheckCloseButton.Add_MouseEnter({ if ($SelectedTheme -eq 'Premium' -or $SelectedTheme -eq 'Custom') { $script:CertCheckCloseButton.ForeColor = $themeColors.ForeColor } else { $script:CertCheckCloseButton.ForeColor = $themeColors.BackColor } })
    $script:CertCheckCloseButton.Add_MouseLeave({ $script:CertCheckCloseButton.ForeColor = $themeColors.ForeColor })

    # Label for server choice radio buttons
    $script:ServerChoiceLabel = New-FormLabel -LocationX 20 -LocationY 40 -SizeX 200 -SizeY 20 -Text 'Choose servers to check:' -Font $global:NormalBoldFont -TextAlign $DefaultTextAlign

    # Radio button for using Server CSV file
    $script:ServerCSVRadioButton = New-Object System.Windows.Forms.RadioButton
    $script:ServerCSVRadioButton.Location = New-Object System.Drawing.Point(0, 0)
    $script:ServerCSVRadioButton.Size = New-Object System.Drawing.Size(200, 20)
    $script:ServerCSVRadioButton.Font = [System.Drawing.Font]::new($global:DefaultUserFont, 9, [System.Drawing.FontStyle]::Regular)
    $script:ServerCSVRadioButton.Text = "Use Server CSV File"
    $script:ServerCSVRadioButton.Checked = $false

    # Radio button for user to enter servers
    $script:UserServersRadioButton = New-Object System.Windows.Forms.RadioButton
    $script:UserServersRadioButton.Location = New-Object System.Drawing.Point(0, 20)
    $script:UserServersRadioButton.Size = New-Object System.Drawing.Size(200, 40)
    $script:UserServersRadioButton.Font = [System.Drawing.Font]::new($global:DefaultUserFont, 9, [System.Drawing.FontStyle]::Regular)
    $script:UserServersRadioButton.Text = "Enter one or more servers,`r`neach on a new line"
    $script:UserServersRadioButton.Checked = $false

    # Panel for server choice radio buttons
    $script:ServerChoicePanel = New-Object System.Windows.Forms.Panel
    $script:ServerChoicePanel.Location = New-Object System.Drawing.Size(20, 60)
    $script:ServerChoicePanel.Size = New-Object System.Drawing.Size(200, 60)
    $script:ServerChoicePanel.Controls.Add($script:ServerCSVRadioButton)
    $script:ServerChoicePanel.Controls.Add($script:UserServersRadioButton)

    # Text box for user to enter servers
    $script:CertCheckServerTextBox = New-FormTextBox -LocationX 20 -LocationY 120 -SizeX 150 -SizeY 50 -ScrollBars 'Vertical' -Multiline $true -Enabled $false -ReadOnly $false -Text '' -Font $global:NormalFont

    # Label for output radio buttons
    $script:OutputLabel = New-FormLabel -LocationX 20 -LocationY 185 -SizeX 200 -SizeY 20 -Text 'Output results to:' -Font $global:NormalBoldFont -TextAlign $DefaultTextAlign

    # Radio button for outputting results to the output text box
    $script:OutTextRadioButton = New-Object System.Windows.Forms.RadioButton
    $script:OutTextRadioButton.Location = New-Object System.Drawing.Point(0, 0)
    $script:OutTextRadioButton.Size = New-Object System.Drawing.Size(200, 20)
    $script:OutTextRadioButton.Font = [System.Drawing.Font]::new($global:DefaultUserFont, 9, [System.Drawing.FontStyle]::Regular)
    $script:OutTextRadioButton.Text = 'Output Text Box'
    $script:OutTextRadioButton.Checked = $false

    # Radio button for outputting results to a text file
    $script:TextFileRadioButton = New-Object System.Windows.Forms.RadioButton
    $script:TextFileRadioButton.Location = New-Object System.Drawing.Point(0, 20)
    $script:TextFileRadioButton.Size = New-Object System.Drawing.Size(200, 20)
    $script:TextFileRadioButton.Font = [System.Drawing.Font]::new($global:DefaultUserFont, 9, [System.Drawing.FontStyle]::Regular)
    $script:TextFileRadioButton.Text = 'Text File'
    $script:TextFileRadioButton.Checked = $false

    # Button for choosing a text file to save results to
    $script:ChooseTextFileButton = New-FormButton -Text "Browse" -LocationX 20 -LocationY 245 -Width 75 -BackColor $global:DisabledBackColor -ForeColor $global:DisabledForeColor -Font $global:NormalFont -Enabled $false -Visible $true

    # Event handler for the file location button
    $script:ChooseTextFileButton.Add_Click({
        $CertCheckFileBrowser = New-Object System.Windows.Forms.OpenFileDialog
        $CertCheckFileBrowser.InitialDirectory = "C:\"
        $CertCheckFileBrowserResult = $CertCheckFileBrowser.ShowDialog()
        if ($CertCheckFileBrowserResult -eq 'OK') {
            $script:CertCheckPopup.Tag = $CertCheckFileBrowser.FileName
            Write-Host "File path set to $($CertCheckFileBrowser.FileName)"
            $script:OutputFileTextbox.Text = $CertCheckFileBrowser.FileName
            #Enable-CreateHDTStorageButton
        }
    })

    # Textbox for choosing a text file to save results to
    $script:OutputFileTextbox = New-FormTextBox -LocationX 20 -LocationY 290 -SizeX 150 -SizeY 20 -ScrollBars 'None' -Multiline $true -Enabled $true -ReadOnly $false -Text '' -Font $global:NormalFont

    # Label for cert choice radio buttons
    $script:CertChoiceLabel = New-FormLabel -LocationX 20 -LocationY 355 -SizeX 200 -SizeY 20 -Text 'Choose certificates to check:' -Font $global:NormalBoldFont -TextAlign $DefaultTextAlign

    # Panel for cert choice radio buttons
    $script:CertChoicePanel = New-Object System.Windows.Forms.Panel
    $script:CertChoicePanel.Location = New-Object System.Drawing.Size(20, 375)
    $script:CertChoicePanel.Size = New-Object System.Drawing.Size(200, 60)
    $script:CertChoicePanel.Controls.Add($script:AllCertsRadioButton)
    $script:CertChoicePanel.Controls.Add($script:ExpiringCertsRadioButton)

    # Radio button for returning all certs
    $script:AllCertsRadioButton = New-Object System.Windows.Forms.RadioButton
    $script:AllCertsRadioButton.Location = New-Object System.Drawing.Point(0, 0)
    $script:AllCertsRadioButton.Size = New-Object System.Drawing.Size(200, 20)
    $script:AllCertsRadioButton.Font = [System.Drawing.Font]::new($global:DefaultUserFont, 9, [System.Drawing.FontStyle]::Regular)
    $script:AllCertsRadioButton.Text = 'All Certificates'

    # Radio button for returning certs expiring within the user-chosen timeframe
    $script:ExpiringCertsRadioButton = New-Object System.Windows.Forms.RadioButton
    $script:ExpiringCertsRadioButton.Location = New-Object System.Drawing.Point(0, 20)
    $script:ExpiringCertsRadioButton.Size = New-Object System.Drawing.Size(200, 20)
    $script:ExpiringCertsRadioButton.Font = [System.Drawing.Font]::new($global:DefaultUserFont, 9, [System.Drawing.FontStyle]::Regular)
    $script:ExpiringCertsRadioButton.Text = 'Expiring Certificates'

    # Panel for output radio buttons
    $script:OutputPanel = New-Object System.Windows.Forms.Panel
    $script:OutputPanel.Location = New-Object System.Drawing.Size(20, 205)
    $script:OutputPanel.Size = New-Object System.Drawing.Size(260, 80)
    $script:OutputPanel.Controls.Add($script:TextFileRadioButton)
    $script:OutputPanel.Controls.Add($script:OutTextRadioButton)

    # Checkbox for specifying whether to include servers with no results
    $script:IgnoreNoResultsCheckBox = New-FormCheckbox -LocationX 20 -LocationY 450 -SizeX 200 -SizeY 20 -Text 'Ignore servers with no results?' -Font $global:NormalFont -Checked $false -Enabled $true

    # Button for running Cert Check
    $script:RunCertCheckButton = New-FormButton -Text "Run Checks" -LocationX 20 -LocationY 500 -Width 100 -BackColor $global:DisabledBackColor -ForeColor $global:DisabledForeColor -Font $global:NormalFont -Enabled $false -Visible $true

    # Event handling for user to select a text file to save results to
    $script:ChooseTextFileButton.add_Click({
        $script:ChooseTextFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $script:ChooseTextFileDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
        $script:ChooseTextFileDialog.Filter = "Text Files (*.txt)|*.txt"
        $script:ChooseTextFileResult = $script:ChooseTextFileDialog.ShowDialog()

        if ($script:ChooseTextFileResult -eq [System.Windows.Forms.DialogResult]::OK) {
            $script:OutputFilePath = $script:ChooseTextFileDialog.FileName
            $OutText.AppendText("$(Get-Timestamp) - Output file path set to $script:OutputFilePath`r`n")
            Enable-CertCheckObjects
        }
    })

    # Event handler for the Server CSV radio button
    $script:ServerCSVRadioButton.add_CheckedChanged({
        $script:CertCheckServerTextBox.Enabled = $false
        $script:CertCheckServerTextBox.Text = ''
        Enable-CertCheckObjects
    })

    # Event handler for the User Servers radio button
    $script:UserServersRadioButton.add_CheckedChanged({
        Enable-CertCheckObjects
    })

    # Event handler for the Servers text box
    $script:CertCheckServerTextBox.add_TextChanged({
        Enable-CertCheckObjects
    })

    # Event handler for the Output Text Box radio button
    $script:OutTextRadioButton.add_CheckedChanged({
        $script:OutputFilePath = $null
        Enable-CertCheckObjects
    })

    # Event handler for the Text File radio button
    $script:TextFileRadioButton.add_CheckedChanged({
        Enable-CertCheckObjects
    })

    # Event handler for running Cert Checks
    $script:RunCertCheckButton.add_Click({
        if ($script:CertCheckServerTextBox.Text -eq '') {
            $UserChosenServers = $null
        } else {
            $UserChosenServers = $script:CertCheckServerTextBox.Text
        }
        if ($script:IgnoreNoResultsCheckBox.Checked) {
            $IgnoreFailedServers = $true
        }
        else {
            $IgnoreFailedServers = $false
        }
        $script:CertCheckPopup.Close()
        $OutText.AppendText("$(Get-Timestamp) - Starting Cert Check Process...`r`n")
        Open-CertCheckRunspace -ConfigValuesCertCheckScript $ConfigValues.CertCheckScript -UserChosenServers $UserChosenServers -OutputFilePath $script:OutputFilePath -IgnoreFailedServers $IgnoreFailedServers -OutText $OutText -TimestampFunction ${function:Get-Timestamp}
    })

    # Check if DefaultUserTheme has a value or is null
    if ($null -ne $ConfigValues.DefaultUserTheme -and $ConfigValues.DefaultUserTheme -ne '') {
        # Get theme
        $SelectedTheme = $ColorTheme.PSObject.Properties | Where-Object { $_.Value.$($ConfigValues.DefaultUserTheme) } | Select-Object -ExpandProperty Name
        $themeColors = $ColorTheme.$SelectedTheme.$($ConfigValues.DefaultUserTheme)

        if ($SelectedTheme -eq 'Premium' -or $SelectedTheme -eq 'Custom') {
            $script:CertCheckServerTextBox.BackColor = $themeColors.AccentColor
        }
        else {
            $script:CertCheckServerTextBox.BackColor = [System.Drawing.SystemColors]::Control
            $global:DisabledBackColor = '#A9A9A9'
        }

        if ($ConfigValues.DefaultUserTheme -eq 'USA') {
            Enable-USAThemeTextColor
        }
        
        $script:ServerChoiceLabel.ForeColor = $themeColors.ForeColor
        $script:CertCheckPopup.BackColor = $themeColors.BackColor
        $script:CertCheckPopup.ForeColor = $themeColors.ForeColor
        $script:OutputLabel.ForeColor = $themeColors.ForeColor
    }

    # Click event for form close; sets global variable to false
    $script:CertCheckPopup.Add_FormClosed({
        $global:IsCertCheckPopupActive = $false
    })

    # Cert Check form build
    $script:CertCheckPopup.Controls.Add($script:CertCheckMinimizeButton)
    $script:CertCheckPopup.Controls.Add($script:CertCheckCloseButton)
    $script:CertCheckPopup.Controls.Add($script:ServerChoiceLabel)
    $script:CertCheckPopup.Controls.Add($script:ServerChoicePanel)
    $script:CertCheckPopup.Controls.Add($script:CertCheckServerTextBox)
    $script:CertCheckPopup.Controls.Add($script:OutputLabel)
    $script:CertCheckPopup.Controls.Add($script:ChooseTextFileButton)
    $script:CertCheckPopup.Controls.Add($script:OutputPanel)
    $script:CertCheckPopup.Controls.Add($script:OutputFileTextbox)
    $script:CertCheckPopup.Controls.Add($script:CertChoiceLabel)
    $script:CertCheckPopup.Controls.Add($script:CertChoicePanel)
    $script:CertCheckPopup.Controls.Add($script:IgnoreNoResultsCheckBox)
    $script:CertCheckPopup.Controls.Add($script:RunCertCheckButton)

    $script:CertCheckPopup.Show() | Out-Null
})

<#
? **********************************************************************************************************************
? END OF CERT CHECK
? **********************************************************************************************************************
#>
<#
? **********************************************************************************************************************
? START OF AWS MONITORING GUI
? **********************************************************************************************************************
#>

# AWS variables
$AWSSSSOCache = $ConfigValues.AWSSSOCache
$AWSConfigFile = $ConfigValues.AWSConfigFile
$AWSSSOProfile = $ConfigValues.AWSSSOProfile
$AWSSSOLoginOutput = Join-Path -Path $PSScriptRoot -ChildPath $ConfigValues.AWSSSOLoginOutput
$script:IsAWSSSOLoginOutputCodeRetrieved = $false
$AWSSSOCacheFilePath = Join-Path $env:USERPROFILE $AWSSSSOCache
$AWSSSOCacheFile = (Get-ChildItem -Path $AWSSSOCacheFilePath | Sort-Object LastWriteTime -Descending | Select-Object -First 1 -Property FullName).FullName
$AWSSSOCacheJson = Get-Content -Path $AWSSSOCacheFile | Out-String
$AWSAccountsFile = Join-Path -Path $PSScriptRoot -ChildPath $ConfigValues.AWSAccountsFile
$ExpiresAt = ($AWSSSOCacheJson | ConvertFrom-Json).ExpiresAt
$script:TargetDateTime = Get-AWSSSOCacheFile -AWSSSOCacheFilePath $AWSSSOCacheFilePath
$AWSScreenshotPath = Join-Path -Path $PSScriptRoot -ChildPath $ConfigValues.AWSScreenshotPath
$RestartAWSInstanceScript = Join-Path -Path $PSScriptRoot -ChildPath $ConfigValues.RestartAWSInstanceScript
$GetAWSCPUMetricsScript = Join-Path -Path $PSScriptRoot -ChildPath $ConfigValues.GetAWSCPUMetricsScript

# Button for running the AWS SSO login command
$script:AWSSSOLoginButton = New-FormButton -Text 'Login With SSO' -LocationX 440 -LocationY 20 -Width 125 -BackColor $global:DisabledBackColor -ForeColor $global:DisabledForeColor -Font $global:NormalFont -Enabled $false -Visible $true

# Label for displaying the AWS login countdown timer
$script:AWSSSOLoginTimerLabel = New-FormLabel -LocationX 180 -LocationY 20 -SizeX 250 -SizeY 150 -Text '' -Font $global:ExtraLargeFont -TextAlign $TopRightTextAlign

# Timer for AWS SSO login expiration countdown
$script:AWSSSOLoginTimer = New-FormTimer -Interval 1000 -Enabled $true

# Initialize a timer for checking the file update
$script:FileCheckTimer = New-FormTimer -Interval 500 -Enabled $true

# Button for listing AWS accounts
$script:ListAWSAccountsButton = New-FormButton -Text 'List AWS Accounts' -LocationX 5 -LocationY 65 -Width 150 -BackColor $global:DisabledBackColor -ForeColor $global:DisabledForeColor -Font $global:NormalFont -Enabled $false -Visible $true

# Help Icon for AWS Monitoring GUI
$AWSMonitoringGUIHelpIcon = New-FormPictureBox -LocationX 160 -LocationY 67 -SizeX 20 -SizeY 20 -Image $HelpIconImage -SizeMode $AutoSizeMode -Cursor $HandCursor

# Event handler for the AWS Monitoring GUI Help Icon
$AWSMonitoringGUIHelpIcon.Add_Click({
    Show-HelpForm -PictureBox $AWSMonitoringGUIHelpIcon -HelpText "View and manage EC2 instances in all Allied AWS accounts. Login with the AWS SSO button, use the button to the left to retrieve a list of accounts, then select an account to view its instances. Select one or more instances to perform actions on them."
})

# List box for AWS accounts
$script:AWSAccountsListBox = New-FormListBox -LocationX 5 -LocationY 95 -SizeX 200 -SizeY 240 -Font $global:NormalFont -SelectionMode 'One'

# List box for AWS instances for a specific account
$script:AWSInstancesListBox = New-FormListBox -LocationX 225 -LocationY 95 -SizeX 200 -SizeY 240 -Font $global:NormalFont -SelectionMode 'MultiExtended'

# Button for restarting AWS instances
$script:RebootAWSInstancesButton = New-FormButton -Text 'Reboot Instance' -LocationX 440 -LocationY 95 -Width 125 -BackColor $global:DisabledBackColor -ForeColor $global:DisabledForeColor -Font $global:NormalFont -Enabled $false -Visible $true

# Button for starting AWS instances
$script:StartAWSInstancesButton = New-FormButton -Text 'Start Instance' -LocationX 440 -LocationY 135 -Width 125 -BackColor $global:DisabledBackColor -ForeColor $global:DisabledForeColor -Font $global:NormalFont -Enabled $false -Visible $true

# Button for stopping AWS instances
$script:StopAWSInstancesButton = New-FormButton -Text 'Stop Instance' -LocationX 440 -LocationY 175 -Width 125 -BackColor $global:DisabledBackColor -ForeColor $global:DisabledForeColor -Font $global:NormalFont -Enabled $false -Visible $true

# Button for taking a screenshot of an AWS instance
$script:AWSScreenshotButton = New-FormButton -Text 'Get Screenshot' -LocationX 440 -LocationY 215 -Width 125 -BackColor $global:DisabledBackColor -ForeColor $global:DisabledForeColor -Font $global:NormalFont -Enabled $false -Visible $true

# Button for retrieving CPU metrics for an AWS instance
$script:AWSCPUMetricsButton = New-FormButton -Text 'Get CPU Metrics' -LocationX 440 -LocationY 255 -Width 125 -BackColor $global:DisabledBackColor -ForeColor $global:DisabledForeColor -Font $global:NormalFont -Enabled $false -Visible $true

# Context menu for AWS CPU Metrics button
$script:AWSCPUMetricsContextMenuStrip = New-Object System.Windows.Forms.ContextMenuStrip
$CPUMetricsOneHour = New-Object System.Windows.Forms.ToolStripMenuItem
$CPUMetricsOneHour.Text = 'Past Hour'
$CPUMetricsSixHours = New-Object System.Windows.Forms.ToolStripMenuItem
$CPUMetricsSixHours.Text = 'Past 6 Hours'
$CPUMetricsTwelveHours = New-Object System.Windows.Forms.ToolStripMenuItem
$CPUMetricsTwelveHours.Text = 'Past 12 Hours'
$CPUMetricsDay = New-Object System.Windows.Forms.ToolStripMenuItem
$CPUMetricsDay.Text = 'Past 24 Hours'

# Add the CPU Metrics menu items to the context menu strip
$script:AWSCPUMetricsContextMenuStrip.Items.AddRange(@($CPUMetricsOneHour, $CPUMetricsSixHours, $CPUMetricsTwelveHours, $CPUMetricsDay))

# Add the context menu strip to the CPU Metrics button
$script:AWSCPUMetricsButton.ContextMenuStrip = $script:AWSCPUMetricsContextMenuStrip

# Click event handler to show the CPU Metrics context menu when the button is clicked
$script:AWSCPUMetricsButton.Add_Click({
    $script:AWSCPUMetricsContextMenuStrip.Show($script:AWSCPUMetricsButton, $script:AWSCPUMetricsButton.PointToClient([System.Windows.Forms.Cursor]::Position))
})

# Add tick event for AWS SSO login timer
$script:AWSSSOLoginTimer.Add_Tick({
    Update-CountdownTimer $script:AWSSSOLoginTimerLabel $script:TargetDateTime $script:AWSSSOLoginTimer
})

# Start the AWS SSO login expiration timer
$script:AWSSSOLoginTimer.Start()

# Click event for AWS SSO login button
$script:AWSSSOLoginButton.Add_Click({

    # Call the function to perform AWS SSO login
    Open-AWSSSOLoginRunspace -AWSSSOLoginOutput $AWSSSOLoginOutput -AWSSSOProfile $AWSSSOProfile -AWSSSOCacheFilePath $AWSSSOCacheFilePath -OutText $OutText -TimestampFunction ${function:Get-Timestamp}

    $script:FileCheckTimer.Add_Tick({
            $script:TargetDateTime = Get-AWSSSOCacheFile -AWSSSOCacheFilePath $AWSSSOCacheFilePath

            if ($script:IsAWSSSOLoginOutputCodeRetrieved -eq $false) {
                $script:IsAWSSSOLoginOutputCodeRetrieved = Get-AWSSSOLoginOutput -AWSSSOLoginOutput $AWSSSOLoginOutput -OutText $OutText -TimestampFunction ${function:Get-Timestamp}
            }

            Update-CountdownTimer $script:AWSSSOLoginTimerLabel $script:TargetDateTime $script:AWSSSOLoginTimer
    })

    # Start the file check timer
    $script:FileCheckTimer.Start()
})

# Button click event handler for listing AWS accounts
$script:ListAWSAccountsButton.Add_Click({

    $OutText.AppendText("$(Get-Timestamp) - Retrieving list of AWS accounts...`r`n")

    # Get the most recent cache file
    $AWSSSOCacheFiles = Get-ChildItem -Path $AWSSSOCacheFilePath -Filter "*.json" | Sort-Object LastWriteTime -Descending
    $AccessTokenFile = $AWSSSOCacheFiles | Select-Object -First 1

    # Read the access token from the cache file
    $AccessTokenContent = Get-Content -Path $AccessTokenFile.FullName | ConvertFrom-Json
    $AWSAccessToken = $AccessTokenContent.accessToken

    Get-AWSAccounts -AWSSSOProfile $AWSSSOProfile -AWSAccessToken $AWSAccessToken -AWSAccountsFile $AWSAccountsFile -AWSAccountsListBox $script:AWSAccountsListBox -OutText $OutText -TimestampFunction ${function:Get-Timestamp}
})

# Store the Describe Instances functions in a variable
$OnAWSAccountSelectedAction = { OnAWSAccountSelected -AWSSSOProfile $AWSSSOProfile -AWSAccountsFile $AWSAccountsFile -AWSConfigFile $AWSConfigFile -AWSAccountsListBox $AWSAccountsListBox -AWSInstancesListBox $AWSInstancesListBox -OutText $OutText -TimestampFunction ${function:Get-Timestamp} }

# Add the event handler for the AWS Accounts list box
$script:AWSAccountsListBox.Add_SelectedIndexChanged($OnAWSAccountSelectedAction)

# Store the function to get instance status in a variable
$OnAWSInstanceSelectedAction = { OnAWSInstanceSelected -AWSSSOProfile $AWSSSOProfile -AWSInstancesListBox $AWSInstancesListBox -OutText $OutText -TimestampFunction ${function:Get-Timestamp}}

# Add the event handler for the AWS Instances list box
$script:AWSInstancesListBox.Add_SelectedIndexChanged($OnAWSInstanceSelectedAction)

# Button click event handler for rebooting AWS Instances in an async runspace pool
$script:RebootAWSInstancesButton.Add_Click({
    $Action = 'Reboot'

    $SelectedInstanceName = @()
    foreach ($Instance in $script:AWSInstancesListBox.SelectedItems) {
        $SelectedInstanceName += $Instance
    }

    Open-RebootAWSInstancesRunspace -RestartAWSInstanceScript $RestartAWSInstanceScript -AWSSSOProfile $AWSSSOProfile -Action $Action -SelectedInstanceName $SelectedInstanceName -OutText $OutText -TimestampFunction ${function:Get-Timestamp}
})

# Button click event handling for starting AWS Instances in an async runspace pool
$script:StartAWSInstancesButton.Add_Click({
    $Action = 'Start'

    $SelectedInstanceName = @()
    foreach ($Instance in $script:AWSInstancesListBox.SelectedItems) {
        $SelectedInstanceName += $Instance
    }

    Open-RebootAWSInstancesRunspace -RestartAWSInstanceScript $RestartAWSInstanceScript -AWSSSOProfile $AWSSSOProfile -Action $Action -SelectedInstanceName $SelectedInstanceName -OutText $OutText -TimestampFunction ${function:Get-Timestamp}
})

# Button click event handling for stopping AWS Instances in an async runspace pool
$script:StopAWSInstancesButton.Add_Click({
    $Action = 'Stop'

    $SelectedInstanceName = @()
    foreach ($Instance in $script:AWSInstancesListBox.SelectedItems) {
        $SelectedInstanceName += $Instance
    }

    Open-RebootAWSInstancesRunspace -RestartAWSInstanceScript $RestartAWSInstanceScript -AWSSSOProfile $AWSSSOProfile -Action $Action -SelectedInstanceName $SelectedInstanceName -OutText $OutText -TimestampFunction ${function:Get-Timestamp}
})

# Event logic for selecting an item in AWS Instances list box
$script:AWSInstancesListBox.Add_SelectedIndexChanged({
    Enable-AWSGUIButtons
})

# Click event handler for the AWS Instance screenshot button
$script:AWSScreenshotButton.Add_Click({

    $SelectedInstanceName = $script:AWSInstancesListBox.SelectedItem.ToString()

    $InstanceScreenshotFile = Join-Path -Path $AWSScreenshotPath -ChildPath "$SelectedInstanceName.png"

    $OutText.AppendText("$(Get-Timestamp) - Retrieving screenshot of $SelectedInstanceName. This may take a moment...`r`n")

    Get-InstanceScreenshot -SelectedInstanceName $SelectedInstanceName -AWSSSOProfile $AWSSSOProfile -InstanceScreenshotFile $InstanceScreenshotFile -OutText $OutText -TimestampFunction ${function:Get-Timestamp}
})

# Click event for the CPU Metrics DAY context menu item
$CPUMetricsDay.add_Click({
    if ($null -ne $ConfigValues.DefaultUserTheme -and $ConfigValues.DefaultUserTheme -ne '') {
        $SelectedTheme = $ColorTheme.PSObject.Properties | Where-Object { $_.Value.$($ConfigValues.DefaultUserTheme) } | Select-Object -ExpandProperty Name
        $themeColors = $ColorTheme.$SelectedTheme.$($ConfigValues.DefaultUserTheme)
        $BackColor = $themeColors.BackColor
        $ForeColor = $themeColors.ForeColor
        
        # Check if the theme falls under Premium and set the AccentColor accordingly
        if ($SelectedTheme -eq 'Premium' -or $SelectedTheme -eq 'Custom') {
            $AccentColor = $themeColors.AccentColor
        }
        else {
            $AccentColor = 'White'
        }
    }

    $SelectedInstanceName = $script:AWSInstancesListBox.SelectedItem.ToString()

    # Get Start and End times for the CPU Metrics
    $24HourTimespan = New-TimeSpan -Hours 24
    $DayValue = (Get-Date) - $24HourTimespan
    $StartTime = $DayValue.ToString('yyyy-MM-ddTHH:mm:ss')

    $EndTime = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss')

    $PollingPeriod = 3600

    Open-GetAWSCPUMetricsRunspace -GetAWSCPUMetricsScript $GetAWSCPUMetricsScript -AWSSSOProfile $AWSSSOProfile -SelectedInstanceName $SelectedInstanceName -StartTime $StartTime -EndTime $EndTime -PollingPeriod $PollingPeriod -BackColor $BackColor -ForeColor $ForeColor -AccentColor $AccentColor -OutText $OutText -TimestampFunction ${function:Get-Timestamp}
})

# Click event for the CPU Metrics 12 Hour context menu item
$CPUMetricsTwelveHours.Add_Click({
    if ($null -ne $ConfigValues.DefaultUserTheme -and $ConfigValues.DefaultUserTheme -ne '') {
        $SelectedTheme = $ColorTheme.PSObject.Properties | Where-Object { $_.Value.$($ConfigValues.DefaultUserTheme) } | Select-Object -ExpandProperty Name
        $themeColors = $ColorTheme.$SelectedTheme.$($ConfigValues.DefaultUserTheme)
        $BackColor = $themeColors.BackColor
        $ForeColor = $themeColors.ForeColor
        
        # Check if the theme falls under Premium and set the AccentColor accordingly
        if ($SelectedTheme -eq 'Premium' -or $SelectedTheme -eq 'Custom') {
            $AccentColor = $themeColors.AccentColor
        }
        else {
            $AccentColor = 'White'
        }
    }

    $SelectedInstanceName = $script:AWSInstancesListBox.SelectedItem.ToString()

    # Get Start and End times for the CPU Metrics
    $12HourTimespan = New-TimeSpan -Hours 12
    $12HourValue = (Get-Date) - $12HourTimespan
    $StartTime = $12HourValue.ToString('yyyy-MM-ddTHH:mm:ss')

    $EndTime = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss')

    $PollingPeriod = 1800

    Open-GetAWSCPUMetricsRunspace -GetAWSCPUMetricsScript $GetAWSCPUMetricsScript -AWSSSOProfile $AWSSSOProfile -SelectedInstanceName $SelectedInstanceName -StartTime $StartTime -EndTime $EndTime -PollingPeriod $PollingPeriod -BackColor $BackColor -ForeColor $ForeColor -AccentColor $AccentColor -OutText $OutText -TimestampFunction ${function:Get-Timestamp}
})

# Click event for the CPU Metrics 6 Hour context menu item
$CPUMetricsSixHours.Add_Click({
    if ($null -ne $ConfigValues.DefaultUserTheme -and $ConfigValues.DefaultUserTheme -ne '') {
        $SelectedTheme = $ColorTheme.PSObject.Properties | Where-Object { $_.Value.$($ConfigValues.DefaultUserTheme) } | Select-Object -ExpandProperty Name
        $themeColors = $ColorTheme.$SelectedTheme.$($ConfigValues.DefaultUserTheme)
        $BackColor = $themeColors.BackColor
        $ForeColor = $themeColors.ForeColor
        
        # Check if the theme falls under Premium and set the AccentColor accordingly
        if ($SelectedTheme -eq 'Premium' -or $SelectedTheme -eq 'Custom') {
            $AccentColor = $themeColors.AccentColor
        }
        else {
            $AccentColor = 'White'
        }
    }

    $SelectedInstanceName = $script:AWSInstancesListBox.SelectedItem.ToString()

    # Get Start and End times for the CPU Metrics
    $6HourTimespan = New-TimeSpan -Hours 6
    $6HourValue = (Get-Date) - $6HourTimespan
    $StartTime = $6HourValue.ToString('yyyy-MM-ddTHH:mm:ss')

    $EndTime = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss')

    $PollingPeriod = 900

    Open-GetAWSCPUMetricsRunspace -GetAWSCPUMetricsScript $GetAWSCPUMetricsScript -AWSSSOProfile $AWSSSOProfile -SelectedInstanceName $SelectedInstanceName -StartTime $StartTime -EndTime $EndTime -PollingPeriod $PollingPeriod -BackColor $BackColor -ForeColor $ForeColor -AccentColor $AccentColor -OutText $OutText -TimestampFunction ${function:Get-Timestamp}
})

# Click event for the CPU Metrics 1 Hour context menu item
$CPUMetricsOneHour.Add_Click({
    if ($null -ne $ConfigValues.DefaultUserTheme -and $ConfigValues.DefaultUserTheme -ne '') {
        $SelectedTheme = $ColorTheme.PSObject.Properties | Where-Object { $_.Value.$($ConfigValues.DefaultUserTheme) } | Select-Object -ExpandProperty Name
        $themeColors = $ColorTheme.$SelectedTheme.$($ConfigValues.DefaultUserTheme)
        $BackColor = $themeColors.BackColor
        $ForeColor = $themeColors.ForeColor
        
        # Check if the theme falls under Premium and set the AccentColor accordingly
        if ($SelectedTheme -eq 'Premium' -or $SelectedTheme -eq 'Custom') {
            $AccentColor = $themeColors.AccentColor
        }
        else {
            $AccentColor = 'White'
        }
    }

    $SelectedInstanceName = $script:AWSInstancesListBox.SelectedItem.ToString()

    # Get Start and End times for the CPU Metrics
    $1HourTimespan = New-TimeSpan -Hours 1
    $1HourValue = (Get-Date) - $1HourTimespan
    $StartTime = $1HourValue.ToString('yyyy-MM-ddTHH:mm:ss')

    $EndTime = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss')

    $PollingPeriod = 300

    Open-GetAWSCPUMetricsRunspace -GetAWSCPUMetricsScript $GetAWSCPUMetricsScript -AWSSSOProfile $AWSSSOProfile -SelectedInstanceName $SelectedInstanceName -StartTime $StartTime -EndTime $EndTime -PollingPeriod $PollingPeriod -BackColor $BackColor -ForeColor $ForeColor -AccentColor $AccentColor -OutText $OutText -TimestampFunction ${function:Get-Timestamp}
})

<#
? **********************************************************************************************************************
? END OF AWS MONITORING GUI
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
$PSTCombo = New-FormComboBox -LocationX 5 -LocationY 65 -SizeX 150 -SizeY 200 -Font $global:NormalFont
@('QA', 'Stage', 'Production') | ForEach-Object { [void]$PSTCombo.Items.Add($_) }

# Label for PST combo box
$PSTComboLabel = New-FormLabel -LocationX 5 -LocationY 40 -SizeX 112 -SizeY 20 -Text 'Prod Support Tool' -Font $global:NormalBoldFont -TextAlign $DefaultTextAlign

# Help Icon for Prod Support Tool
$PSTHelpIcon = New-FormPictureBox -LocationX 116 -LocationY 40 -SizeX 20 -SizeY 20 -Image $HelpIconImage -SizeMode $AutoSizeMode -Cursor $HandCursor

# Event handler for Prod Support Tool Help Icon
$PSTHelpIcon.Add_Click({
    Show-HelpForm -PictureBox $PSTHelpIcon -HelpText "Launches a local installation of the UniTrac Prod Support Tool, which allows for multiple concurrent users. Also checks the workhorse server for more recent files each time Desktop Assistant starts up. The selected environment is reset each time the main form is closed."
})

# Button for switching environment
$SelectEnvButton = New-FormButton -Text "Select Environment" -LocationX 170 -LocationY 35 -Width 150 -BackColor $global:DisabledBackColor -ForeColor $global:DisabledForeColor -Font $global:NormalFont -Enabled $false -Visible $true

# Button for resetting environment
$ResetEnvButton = New-FormButton -Text "Reset Environment" -LocationX 170 -LocationY 65 -Width 150 -BackColor $global:DisabledBackColor -ForeColor $global:DisabledForeColor -Font $global:NormalFont -Enabled $false -Visible $true

# Button for running the PST
$RunPSTButton = New-FormButton -Text "Run Prod Support Tool" -LocationX 170 -LocationY 115 -Width 150 -BackColor $global:DisabledBackColor -ForeColor $global:DisabledForeColor -Font $global:NormalFont -Enabled $false -Visible $true

# Button for refreshing the PST files
$RefreshPSTButton = New-FormButton -Text $RefreshButtonText -LocationX 5 -LocationY 115 -Width 150 -BackColor $ControlColor -ForeColor $ControlColor -Font $global:NormalFont -Enabled $true -Visible $true

# Separator line under PST
$PSTSeparator = New-FormLabel -LocationX 0 -LocationY 170 -SizeX 1000 -SizeY 2 -Text '' -Font $global:NormalFont -TextAlign $DefaultTextAlign

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
    Update-PSTFiles -Team $Team -Category $Category -TimestampFunction ${function:Get-Timestamp}
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
$LaunchLFPWizardButton = New-FormButton -Text "Launch LFP Wizard" -LocationX 400 -LocationY 35 -Width 150 -BackColor $ControlColor -ForeColor $ControlColor -Font $global:NormalFont -Enabled $true -Visible $true

# Label for LFP wizard button
$LaunchLFPWizardLabel = New-FormLabel -LocationX 413 -LocationY 10 -SizeX 127 -SizeY 20 -Text 'Lender LFP Services' -Font $global:NormalBoldFont -TextAlign $DefaultTextAlign

# Help Icon for LFP wizard
$LFPWizardHelpIcon = New-FormPictureBox -LocationX 540 -LocationY 10 -SizeX 20 -SizeY 20 -Image $HelpIconImage -SizeMode $AutoSizeMode -Cursor $HandCursor

# Event handler for LFP wizard help icon
$LFPWizardHelpIcon.Add_Click({
    Show-HelpForm -PictureBox $LFPWizardHelpIcon -HelpText "Add a lender to the Lender File Processing (LFP) services in QA, Staging, or Production. The services to be restarted and script to run are located at ..\AddLendertoLFPServices\. Also integrates with Ticket Manager to create a new ticket and move the SQL script to its respective folder."
})

# Click event handler for launching the LFP wizard
$LaunchLFPWizardButton.Add_Click({
    if ($global:IsLenderLFPPopupActive -eq $true) {
        $OutText.AppendText("$(Get-Timestamp) - LFP Wizard is already open.`r`n")
        $script:LenderLFPPopup.Activate()
        return
    }
    $OutText.AppendText("$(Get-Timestamp) - Launching LFP Wizard...`r`n")

    # Create the Lender LFP popup form
    $script:LenderLFPPopup = New-WindowsForm -SizeX 300 -SizeY 350 -Text '' -StartPosition 'CenterScreen' -TopMost $false -ShowInTaskbar $true -KeyPreview $true -MinimizeBox $true
    $global:IsLenderLFPPopupActive = $true

    # Mouse Down event handler for moving the Lender LFP form
    # This is needed since the form has no ControlBox
    $script:LenderLFPPopup_MouseDown = {
        $global:MouseDown = $true
        $global:MouseClickPoint = [System.Windows.Forms.Cursor]::Position
    }

    # Mouse Down event handler for moving the Lender LFP form
    # This is needed since the form has no ControlBox
    $script:LenderLFPPopup_MouseMove = {
        if ($global:MouseDown) {
            $CurrentCursorPosition = [System.Windows.Forms.Cursor]::Position
            $FormLocation = $script:LenderLFPPopup.Location
            
            $newX = $FormLocation.X + ($CurrentCursorPosition.X - $global:MouseClickPoint.X)
            $newY = $FormLocation.Y + ($CurrentCursorPosition.Y - $global:MouseClickPoint.Y)
            
            $script:LenderLFPPopup.Location = New-Object System.Drawing.Point($newX, $newY)
            $global:MouseClickPoint = $CurrentCursorPosition
        }
    }

    # Mouse Down event handler for moving the Lender LFP form
    # This is needed since the form has no ControlBox
    $script:LenderLFPPopup_MouseUp = {
        $global:MouseDown = $false
    }

    # Add event handlers to the form
    $script:LenderLFPPopup.Add_MouseDown($script:LenderLFPPopup_MouseDown)
    $script:LenderLFPPopup.Add_MouseMove($script:LenderLFPPopup_MouseMove)
    $script:LenderLFPPopup.Add_MouseUp($script:LenderLFPPopup_MouseUp)

    # Button for minimizing the Save Theme popup form
    $script:LenderLFPMinimizeButton = New-FormButton -Text "─" -LocationX 220 -LocationY 0 -Width 40 -BackColor $themeColors.BackColor -ForeColor $themeColors.ForeColor -Font $global:NormalBoldFont -Enabled $true -Visible $true
    $script:LenderLFPMinimizeButton.Add_Click({ $script:LenderLFPPopup.WindowState = 'Minimized' })
    $script:LenderLFPMinimizeButton.FlatStyle = 'Flat'
    $script:LenderLFPMinimizeButton.FlatAppearance.BorderSize = 0
    $script:LenderLFPMinimizeButton.FlatAppearance.MouseOverBackColor = if ($SelectedTheme -eq 'Premium' -or $SelectedTheme -eq 'Custom') { $themeColors.AccentColor } else { $themeColors.ForeColor }
    $script:LenderLFPMinimizeButton.Add_MouseEnter({ if ($SelectedTheme -eq 'Premium' -or $SelectedTheme -eq 'Custom') { $script:LenderLFPMinimizeButton.ForeColor = $themeColors.ForeColor } else { $script:LenderLFPMinimizeButton.ForeColor = $themeColors.BackColor } })
    $script:LenderLFPMinimizeButton.Add_MouseLeave({ $script:LenderLFPMinimizeButton.ForeColor = $themeColors.ForeColor })

    # Button for closing the Save Theme popup form
    $script:LenderLFPCloseButton = New-FormButton -Text "X" -LocationX 260 -LocationY 0 -Width 40 -BackColor $themeColors.BackColor -ForeColor $themeColors.ForeColor -Font $global:NormalBoldFont -Enabled $true -Visible $true
    $script:LenderLFPCloseButton.Add_Click({ $script:LenderLFPPopup.Close() })
    $script:LenderLFPCloseButton.FlatStyle = 'Flat'
    $script:LenderLFPCloseButton.FlatAppearance.BorderSize = 0
    $script:LenderLFPCloseButton.FlatAppearance.MouseOverBackColor = if ($SelectedTheme -eq 'Premium' -or $SelectedTheme -eq 'Custom') { $themeColors.AccentColor } else { $themeColors.ForeColor }
    $script:LenderLFPCloseButton.Add_MouseEnter({ if ($SelectedTheme -eq 'Premium' -or $SelectedTheme -eq 'Custom') { $script:LenderLFPCloseButton.ForeColor = $themeColors.ForeColor } else { $script:LenderLFPCloseButton.ForeColor = $themeColors.BackColor } })
    $script:LenderLFPCloseButton.Add_MouseLeave({ $script:LenderLFPCloseButton.ForeColor = $themeColors.ForeColor })

    # Combo box for selecting the environment where lender will be added
    $script:LenderLFPCombo = New-FormComboBox -LocationX 5 -LocationY 65 -SizeX 100 -SizeY 200 -Font $global:NormalFont
    @('QA', 'Staging', 'Production') | ForEach-Object { [void]$script:LenderLFPCombo.Items.Add($_) }

    # Label for Lender LFP combo box
    $script:LenderLFPComboLabel = New-FormLabel -LocationX 5 -LocationY 40 -SizeX 150 -SizeY 20 -Text 'Select Environment' -Font $global:NormalBoldFont -TextAlign $DefaultTextAlign

    # Text box for entering the lender Id
    $script:LenderLFPIdTextBox = New-FormTextBox -LocationX 5 -LocationY 125 -SizeX 150 -SizeY 20 -ScrollBars 'None' -Multiline $false -Enabled $true -ReadOnly $false -Text '' -Font $global:NormalFont

    # Label for Lender LFP text box
    $script:LenderLFPTextBoxLabel = New-FormLabel -LocationX 5 -LocationY 100 -SizeX 150 -SizeY 20 -Text 'Lender ID' -Font $global:NormalBoldFont -TextAlign $DefaultTextAlign

    # Text box for entering ticket number
    $script:LenderLFPTicketTextBox = New-FormTextBox -LocationX 5 -LocationY 185 -SizeX 150 -SizeY 20 -ScrollBars 'None' -Multiline $false -Enabled $true -ReadOnly $false -Text '' -Font $global:NormalFont

    # Label for Lender LFP ticket text box
    $script:LenderLFPTicketTextBoxLabel = New-FormLabel -LocationX 5 -LocationY 160 -SizeX 150 -SizeY 20 -Text 'Jira Ticket Number' -Font $global:NormalBoldFont -TextAlign $DefaultTextAlign

    # Checkbox to confirm Production runs
    $script:LFPProductionCheckbox = New-FormCheckbox -LocationX 5 -LocationY 215 -SizeX 225 -SizeY 20 -Text 'Check to confirm PRODUCTION run' -Font $global:NormalFont -Checked $false -Enabled $false

    # Button for adding lender to LFP services
    $script:AddLenderLFPButton = New-FormButton -Text "Add Lender" -LocationX 5 -LocationY 250 -Width 100 -BackColor $global:DisabledBackColor -ForeColor $global:DisabledForeColor -Font $global:NormalFont -Enabled $false -Visible $true

    # Check if DefaultUserTheme has a value or is null
    if ($null -ne $ConfigValues.DefaultUserTheme -and $ConfigValues.DefaultUserTheme -ne '') {
        # Get theme
        $SelectedTheme = $ColorTheme.PSObject.Properties | Where-Object { $_.Value.$($ConfigValues.DefaultUserTheme) } | Select-Object -ExpandProperty Name
        $themeColors = $ColorTheme.$SelectedTheme.$($ConfigValues.DefaultUserTheme)

        if ($SelectedTheme -eq 'Premium' -or $SelectedTheme -eq 'Custom') {
            $script:LenderLFPCombo.BackColor = $themeColors.AccentColor
            $script:LenderLFPIdTextBox.BackColor = $themeColors.AccentColor
            $script:LenderLFPTicketTextBox.BackColor = $themeColors.AccentColor
        }
        else {
            $script:LenderLFPCombo.BackColor = [System.Drawing.SystemColors]::Control
            $script:LenderLFPIdTextBox.BackColor = [System.Drawing.SystemColors]::Control
            $script:LenderLFPTicketTextBox.BackColor = [System.Drawing.SystemColors]::Control
            $global:DisabledBackColor = '#A9A9A9'
        }

        if ($ConfigValues.DefaultUserTheme -eq 'USA') {
            Enable-USAThemeTextColor
        }
        
        $script:LenderLFPPopup.BackColor = $themeColors.BackColor
        $script:LenderLFPPopup.ForeColor = $themeColors.ForeColor
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
            $script:LenderLFPPopup.Controls.Add($script:LFPProductionCheckbox)
            $script:LFPProductionCheckbox.Enabled = $true
        }
        else {
            $script:LenderLFPPopup.Controls.Remove($script:LFPProductionCheckbox)
            $script:LFPProductionCheckbox.Enabled = $false
            $script:LFPProductionCheckbox.Checked = $false
        }
    })

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
    $script:LFPProductionCheckbox.Add_CheckedChanged({
        Enable-AddLenderButton
    })

    # Button click event for adding lender to LFP services
    $script:AddLenderLFPButton.add_Click({
        $AddLenderScript = Join-Path -Path $PSScriptRoot -ChildPath $ConfigValues.AddLenderToLFPServicesScript
        $LenderId = $script:LenderLFPIdTextBox.Text.Trim()
        $TicketNumber = $script:LenderLFPTicketTextBox.Text.Trim()
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
                    $FolderExists = Find-DupeTickets -TicketNumber $TicketNumber
                    if (-not $FolderExists) {
                        mkdir "$TicketsPath\Active\$TicketNumber"
                        Get-ActiveListItems($ActiveTicketsListBox)
                        $OutText.AppendText("$(Get-Timestamp) - Folder created for $TicketNumber`r`n")
                    }
                    $script:LenderLFPPopup.Close()
                    Open-AddLenderScriptRunspace -AddLenderScript $AddLenderScript -LenderId $LenderId -TicketNumber $TicketNumber -Environment $Environment -OutText $OutText -TimestampFunction ${function:Get-Timestamp}
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
    
    $script:LenderLFPPopup.Controls.Add($script:LenderLFPMinimizeButton)
    $script:LenderLFPPopup.Controls.Add($script:LenderLFPCloseButton)
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
? START OF BILLING SERVICE RESTARTS
? **********************************************************************************************************************
#>

# Tooltips for Billing Restart popup form
$BillingRestartToolTip = New-Object System.Windows.Forms.ToolTip
$BillingRestartToolTip.InitialDelay = 100

# Button for launching the restart billing file services wizard
$LaunchBillingRestartButton = New-FormButton -Text "Launch Restart Wizard" -LocationX 400 -LocationY 115 -Width 150 -BackColor $ControlColor -ForeColor $ControlColor -Font $global:NormalFont -Enabled $true -Visible $true

# Label for restart billing file services button
$LaunchBillingRestartLabel = New-FormLabel -LocationX 405 -LocationY 90 -SizeX 143 -SizeY 20 -Text 'Restart Billing Services' -Font $global:NormalBoldFont -TextAlign $DefaultTextAlign

# Help Icon for restart billing file services button
$BillingRestartHelpIcon = New-FormPictureBox -LocationX 540 -LocationY 90 -SizeX 20 -SizeY 20 -Image $HelpIconImage -SizeMode $AutoSizeMode -Cursor $HandCursor

# Event handler for the Billing Service Restart Help icon
$BillingRestartHelpIcon.Add_Click({
    Show-HelpForm -PictureBox $BillingRestartHelpIcon -HelpText "Restart the Billing File Services for QA, Staging, or Production. Servers and service names are located in ..\BillingServicesRestarts\BillingServiceEnvironments.json"
})

# Button click event handler for launching the Billing Service Restarts module
$LaunchBillingRestartButton.add_Click({
    if ($global:IsBillingRestartPopupActive) {
        $OutText.AppendText("$(Get-Timestamp) - Billing Services Restart Wizard is already open.`r`n")
        $script:BillingRestartPopup.Activate()
        return
    }
    $OutText.AppendText("$(Get-Timestamp) - Launching Billing File Services Restart Wizard...`r`n")

    # Create the Billing Restart popup form
    $script:BillingRestartPopup = New-WindowsForm -SizeX 250 -SizeY 235 -Text '' -StartPosition 'CenterScreen' -TopMost $false -ShowInTaskbar $true -KeyPreview $true -MinimizeBox $true
    $global:IsBillingRestartPopupActive = $true

    # Mouse Down event handler for moving the Billing Restart form
    # This is needed since the form has no ControlBox
    $script:BillingRestartPopup_MouseDown = {
        $global:MouseDown = $true
        $global:MouseClickPoint = [System.Windows.Forms.Cursor]::Position
    }

    # Mouse Down event handler for moving the Billing Restart form
    # This is needed since the form has no ControlBox
    $script:BillingRestartPopup_MouseMove = {
        if ($global:MouseDown) {
            $CurrentCursorPosition = [System.Windows.Forms.Cursor]::Position
            $FormLocation = $script:BillingRestartPopup.Location
            
            $newX = $FormLocation.X + ($CurrentCursorPosition.X - $global:MouseClickPoint.X)
            $newY = $FormLocation.Y + ($CurrentCursorPosition.Y - $global:MouseClickPoint.Y)
            
            $script:BillingRestartPopup.Location = New-Object System.Drawing.Point($newX, $newY)
            $global:MouseClickPoint = $CurrentCursorPosition
        }
    }

    # Mouse Down event handler for moving the Billing Restart form
    # This is needed since the form has no ControlBox
    $script:BillingRestartPopup_MouseUp = {
        $global:MouseDown = $false
    }

    # Add event handlers to the form
    $script:BillingRestartPopup.Add_MouseDown($script:BillingRestartPopup_MouseDown)
    $script:BillingRestartPopup.Add_MouseMove($script:BillingRestartPopup_MouseMove)
    $script:BillingRestartPopup.Add_MouseUp($script:BillingRestartPopup_MouseUp)

    # Button for minimizing the HDTStorage popup form
    $script:BillingRestartMinimizeButton = New-FormButton -Text "─" -LocationX 170 -LocationY 0 -Width 40 -BackColor $themeColors.BackColor -ForeColor $themeColors.ForeColor -Font $global:NormalBoldFont -Enabled $true -Visible $true
    $script:BillingRestartMinimizeButton.Add_Click({ $script:BillingRestartPopup.WindowState = 'Minimized' })
    $script:BillingRestartMinimizeButton.FlatStyle = 'Flat'
    $script:BillingRestartMinimizeButton.FlatAppearance.BorderSize = 0
    $script:BillingRestartMinimizeButton.FlatAppearance.MouseOverBackColor = if ($SelectedTheme -eq 'Premium' -or $SelectedTheme -eq 'Custom') { $themeColors.AccentColor } else { $themeColors.ForeColor }
    $script:BillingRestartMinimizeButton.Add_MouseEnter({ if ($SelectedTheme -eq 'Premium' -or $SelectedTheme -eq 'Custom') { $script:BillingRestartMinimizeButton.ForeColor = $themeColors.ForeColor } else { $script:BillingRestartMinimizeButton.ForeColor = $themeColors.BackColor } })
    $script:BillingRestartMinimizeButton.Add_MouseLeave({ $script:BillingRestartMinimizeButton.ForeColor = $themeColors.ForeColor })

    # Button for closing the HDTStorage popup form
    $script:BillingRestartCloseButton = New-FormButton -Text "X" -LocationX 210 -LocationY 0 -Width 40 -BackColor $themeColors.BackColor -ForeColor $themeColors.ForeColor -Font $global:NormalBoldFont -Enabled $true -Visible $true
    $script:BillingRestartCloseButton.Add_Click({ $script:BillingRestartPopup.Close() })
    $script:BillingRestartCloseButton.FlatStyle = 'Flat'
    $script:BillingRestartCloseButton.FlatAppearance.BorderSize = 0
    $script:BillingRestartCloseButton.FlatAppearance.MouseOverBackColor = if ($SelectedTheme -eq 'Premium' -or $SelectedTheme -eq 'Custom') { $themeColors.AccentColor } else { $themeColors.ForeColor }
    $script:BillingRestartCloseButton.Add_MouseEnter({ if ($SelectedTheme -eq 'Premium' -or $SelectedTheme -eq 'Custom') { $script:BillingRestartCloseButton.ForeColor = $themeColors.ForeColor } else { $script:BillingRestartCloseButton.ForeColor = $themeColors.BackColor } })
    $script:BillingRestartCloseButton.Add_MouseLeave({ $script:BillingRestartCloseButton.ForeColor = $themeColors.ForeColor })

    # Label for Billing Restart combo box
    $script:BillingRestartComboLabel = New-FormLabel -LocationX 55 -LocationY 50 -SizeX 150 -SizeY 20 -Text 'Select Environment' -Font $global:NormalBoldFont -TextAlign $DefaultTextAlign

    # Combo box for selecting the environment where lender will be added
    $script:BillingRestartCombo = New-FormComboBox -LocationX 40 -LocationY 75 -SizeX 150 -SizeY 200 -Font $global:NormalFont
    @('QA', 'Staging', 'Production') | ForEach-Object { [void]$script:BillingRestartCombo.Items.Add($_) }

    # Checkbox to confirm Production runs
    $script:BillingRestartProductionCheckbox = New-FormCheckbox -LocationX 25 -LocationY 125 -SizeX 200 -SizeY 20 -Text 'Check to confirm PRODUCTION run' -Font $global:ExtraSmallFont -Checked $false -Enabled $false

    # Button for restarting billing file services
    $script:BillingRestartButton = New-FormButton -Text "Restart" -LocationX 75 -LocationY 160 -Width 75 -BackColor $global:DisabledBackColor -ForeColor $global:DisabledForeColor -Font $global:NormalFont -Enabled $false -Visible $true

    # Check if DefaultUserTheme has a value or is null
    if ($null -ne $ConfigValues.DefaultUserTheme -and $ConfigValues.DefaultUserTheme -ne '') {
        # Get theme
        $SelectedTheme = $ColorTheme.PSObject.Properties | Where-Object { $_.Value.$($ConfigValues.DefaultUserTheme) } | Select-Object -ExpandProperty Name
        $themeColors = $ColorTheme.$SelectedTheme.$($ConfigValues.DefaultUserTheme)

        if ($SelectedTheme -eq 'Premium' -or $SelectedTheme -eq 'Custom') {
            $script:BillingRestartCombo.BackColor = $themeColors.AccentColor
        }
        else {
            $script:BillingRestartCombo.BackColor = [System.Drawing.SystemColors]::Control
            $global:DisabledBackColor = '#A9A9A9'
        }

        if ($ConfigValues.DefaultUserTheme -eq 'USA') {
            Enable-USAThemeTextColor
        }
        
        $script:BillingRestartPopup.BackColor = $themeColors.BackColor
        $script:BillingRestartPopup.ForeColor = $themeColors.ForeColor
        $script:BillingRestartCombo.BackColor = $OutText.BackColor
        $script:BillingRestartComboLabel.ForeColor = $themeColors.ForeColor
    }

    # Add tooltips if the user has them enabled
    if ($ConfigValues.HoverToolTips -eq "Enabled" -or $null -eq $ConfigValues.HoverToolTips) {
        $BillingRestartToolTip.SetToolTip($script:BillingRestartCombo, "Select an environment to restart billing services")
        $BillingRestartToolTip.SetToolTip($script:BillingRestartProductionCheckbox, "Check to confirm PRODUCTION run")
        $BillingRestartToolTip.SetToolTip($script:BillingRestartButton, "Click to restart billing services in selected environment")
    }

    # Event handler for checking to see if the Production checkbox should be enabled
    $script:BillingRestartCombo.Add_SelectedIndexChanged({
        if ($script:BillingRestartCombo.SelectedItem -eq "Production") {
            $script:BillingRestartPopup.Controls.Add($script:BillingRestartProductionCheckbox)
            $script:BillingRestartProductionCheckbox.Enabled = $true
        }
        else {
            $script:BillingRestartPopup.Controls.Remove($script:BillingRestartProductionCheckbox)
            $script:BillingRestartProductionCheckbox.Enabled = $false
            $script:BillingRestartProductionCheckbox.Checked = $false
        }
    })

    # Event handler for the Lender LFP combo box
    $script:BillingRestartCombo.Add_SelectedIndexChanged({
        Enable-BillingServicesRestartsButton
    })
	
    # Event handler for Productrion checkbox
    $script:BillingRestartProductionCheckbox.Add_CheckedChanged({
        Enable-BillingServicesRestartsButton
    })

    # Button click event for adding lender to LFP services
    $script:BillingRestartButton.add_Click({
        $BillingRestartsScript = Join-Path -Path $PSScriptRoot -ChildPath $ConfigValues.BillingServicesRestartScript
        $Environment = $script:BillingRestartCombo.SelectedItem

        if (Test-Path $BillingRestartsScript){
            $script:BillingRestartCombo.SelectedItem = $null
            $OutText.AppendText("$(Get-Timestamp) - Beginning process to restart billing file services...`r`n")
            $OutText.AppendText("$(Get-Timestamp) - Environment: $($Environment)`r`n")
            $script:BillingRestartPopup.Close()
            Open-BillingRestartsScriptRunspace -BillingRestartsScript $BillingRestartsScript -Environment $Environment -OutText $OutText -TimestampFunction ${function:Get-Timestamp}
        } else {
            $OutText.AppendText("$(Get-Timestamp) - BillingServicesRestarts.ps1 not found`r`n")
        }
    })

    # Click event for form close; sets global variable to false
    $script:BillingRestartPopup.Add_FormClosed({
        $global:IsBillingRestartPopupActive = $false
    })

    $script:BillingRestartPopup.Controls.Add($script:BillingRestartMinimizeButton)
    $script:BillingRestartPopup.Controls.Add($script:BillingRestartCloseButton)
    $script:BillingRestartPopup.Controls.Add($script:BillingRestartCombo)
    $script:BillingRestartPopup.Controls.Add($script:BillingRestartComboLabel)
    $script:BillingRestartPopup.Controls.Add($script:BillingRestartButton)
    $script:BillingRestartPopup.Show()
})

<#
? **********************************************************************************************************************
? END OF BILLING SERVICE RESTARTS
? **********************************************************************************************************************
#>
<#
? **********************************************************************************************************************
? START OF PASSWORD MANAGER
? **********************************************************************************************************************
#>

# Text box for setting passwords
$PWTextBox = New-FormTextBox -LocationX 5 -LocationY 250 -SizeX 150 -SizeY 20 -ScrollBars 'None' -Multiline $false -Enabled $true -ReadOnly $false -Text '' -Font $global:NormalFont
$PWTextBox.PasswordChar = '*' # Set password masking character

# Label for setting passwords text box
$PWTextBoxLabel = New-FormLabel -LocationX 5 -LocationY 225 -SizeX 145 -SizeY 20 -Text 'Enter Secure Password' -Font $global:NormalBoldFont -TextAlign $DefaultTextAlign

# Help Icon for Password Manager
$PWManagerHelpIcon = New-FormPictureBox -LocationX 150 -LocationY 225 -SizeX 20 -SizeY 20 -Image $HelpIconImage -SizeMode $AutoSizeMode -Cursor $HandCursor

# Event handler for the Password Manager Help icon
$PWManagerHelpIcon.Add_Click({
    Show-HelpForm -PictureBox $PWManagerHelpIcon -HelpText "Store up to two passwords as secure strings in memory and retrieve them as plaintext with the click of a button. Passwords are NEVER stored in Desktop Assistant."
})

# Button for setting your own password
$SetPWButton = New-FormButton -Text "Set Your PW" -LocationX 160 -LocationY 250 -Width 100 -BackColor $global:DisabledBackColor -ForeColor $global:DisabledForeColor -Font $global:NormalFont -Enabled $false -Visible $true

# Button for setting your alternate password
$AltSetPWButton = New-FormButton -Text "Set Alt PW" -LocationX 160 -LocationY 275 -Width 100 -BackColor $global:DisabledBackColor -ForeColor $global:DisabledForeColor -Font $global:NormalFont -Enabled $false -Visible $true

# Button for getting your own password
$GetPWButton = New-FormButton -Text "Get Your PW" -LocationX 160 -LocationY 250 -Width 100 -BackColor $global:DisabledBackColor -ForeColor $global:DisabledForeColor -Font $global:NormalFont -Enabled $false -Visible $true

# Button for getting your alternate password
$AltGetPWButton = New-FormButton -Text "Get Alt PW" -LocationX 160 -LocationY 275 -Width 100 -BackColor $global:DisabledBackColor -ForeColor $global:DisabledForeColor -Font $global:NormalFont -Enabled $false -Visible $true

# Button for clearing your own password
$ClearPWButton = New-FormButton -Text "Clear Your PW" -LocationX 265 -LocationY 250 -Width 100 -BackColor $global:DisabledBackColor -ForeColor $global:DisabledForeColor -Font $global:NormalFont -Enabled $false -Visible $true

# Button for clearing your alternate password
$AltClearPWButton = New-FormButton -Text "Clear Alt PW" -LocationX 265 -LocationY 275 -Width 100 -BackColor $global:DisabledBackColor -ForeColor $global:DisabledForeColor -Font $global:NormalFont -Enabled $false -Visible $true

# Button for generating a random password
$GenPWButton = New-FormButton -Text "Generate" -LocationX 425 -LocationY 250 -Width 100 -BackColor $ControlColor -ForeColor $ControlColor -Font $global:NormalFont -Enabled $true -Visible $true

# Label for generating a random password
$GenPWLabel = New-FormLabel -LocationX 417 -LocationY 225 -SizeX 120 -SizeY 20 -Text 'Generate Password' -Font $global:NormalBoldFont -TextAlign $DefaultTextAlign

# Help Icon for generating a random password
$GenPWHelpIcon = New-FormPictureBox -LocationX 540 -LocationY 225 -SizeX 20 -SizeY 20 -Image $HelpIconImage -SizeMode $AutoSizeMode -Cursor $HandCursor

# Event handler for the Password Generator Help icon
$GenPWHelpIcon.Add_Click({
    Show-HelpForm -PictureBox $GenPWHelpIcon -HelpText "Generate a random, 16-character password with at least one lowercase letter, one uppercase letter, one number, and one of the following special characters: !@#$%^&*()-_=+ The password is NEVER stored in Desktop Assistant."
})

# Separator line for password manager
$PWManagerSeparator = New-FormLabel -LocationX 0 -LocationY 360 -SizeX 1000 -SizeY 2 -Text '' -Font $global:NormalFont -TextAlign $DefaultTextAlign

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
$LaunchHDTStorageButton = New-FormButton -Text "Launch HDT Wizard" -LocationX 5 -LocationY 400 -Width 150 -BackColor $ControlColor -ForeColor $ControlColor -Font $global:NormalFont -Enabled $true -Visible $true

# Label for launching the HDTStorage table creator
$LaunchHDTStorageLabel = New-FormLabel -LocationX 5 -LocationY 375 -SizeX 153 -SizeY 20 -Text 'Create HDTStorage Table' -Font $global:NormalBoldFont -TextAlign $DefaultTextAlign

# Help Icon for launching the HDTStorage table creator
$HDTStorageHelpIcon = New-FormPictureBox -LocationX 158 -LocationY 375 -SizeX 20 -SizeY 20 -Image $HelpIconImage -SizeMode $AutoSizeMode -Cursor $HandCursor

# Event handler for the HDTStorage Help icon
$HDTStorageHelpIcon.Add_Click({
    Show-HelpForm -PictureBox $HDTStorageHelpIcon -HelpText "Create an HDTStorage table in the Database of your choice with your own .xlsx file. This module runs SQL on the Workhorse server and requires your Allied password to connect. Passwords are used solely for the SQL connection and are NEVER stored in Desktop Assistant."
})

# Event handler for the launch HDTStorage button
$LaunchHDTStorageButton.Add_Click({
    if ($global:IsHDTStoragePopupActive) {
        $OutText.AppendText("$(Get-Timestamp) - HDTStorage Wizard is already open.`r`n")
        $script:HDTStoragePopup.Activate()
        return
    }
    $OutText.AppendText("$(Get-Timestamp) - Launching HDTStorage Wizard...`r`n")
    if ((Get-WSManCredSSP).State -ne "Enabled") {
        Enable-WSManCredSSP -Role Client -DelegateComputer $script:WorkhorseServer -Force
    }

    # Create the HDTStorage popup form
    $script:HDTStoragePopup = New-WindowsForm -SizeX 300 -SizeY 350 -Text '' -StartPosition 'CenterScreen' -TopMost $false -ShowInTaskbar $true -KeyPreview $true -MinimizeBox $true
    $global:IsHDTStoragePopupActive = $true

    # Mouse Down event handler for moving the HDTStorage Table form
    # This is needed since the form has no ControlBox
    $script:HDTStoragePopup_MouseDown = {
        $global:MouseDown = $true
        $global:MouseClickPoint = [System.Windows.Forms.Cursor]::Position
    }

    # Mouse Down event handler for moving the HDTStorage Table form
    # This is needed since the form has no ControlBox
    $script:HDTStoragePopup_MouseMove = {
        if ($global:MouseDown) {
            $CurrentCursorPosition = [System.Windows.Forms.Cursor]::Position
            $FormLocation = $script:HDTStoragePopup.Location
            
            $newX = $FormLocation.X + ($CurrentCursorPosition.X - $global:MouseClickPoint.X)
            $newY = $FormLocation.Y + ($CurrentCursorPosition.Y - $global:MouseClickPoint.Y)
            
            $script:HDTStoragePopup.Location = New-Object System.Drawing.Point($newX, $newY)
            $global:MouseClickPoint = $CurrentCursorPosition
        }
    }

    # Mouse Down event handler for moving the HDTStorage Table form
    # This is needed since the form has no ControlBox
    $script:HDTStoragePopup_MouseUp = {
        $global:MouseDown = $false
    }

    # Add event handlers to the form
    $script:HDTStoragePopup.Add_MouseDown($script:HDTStoragePopup_MouseDown)
    $script:HDTStoragePopup.Add_MouseMove($script:HDTStoragePopup_MouseMove)
    $script:HDTStoragePopup.Add_MouseUp($script:HDTStoragePopup_MouseUp)

    # Button for minimizing the HDTStorage popup form
    $script:HDTStorageMinimizeButton = New-FormButton -Text "─" -LocationX 220 -LocationY 0 -Width 40 -BackColor $themeColors.BackColor -ForeColor $themeColors.ForeColor -Font $global:NormalBoldFont -Enabled $true -Visible $true
    $script:HDTStorageMinimizeButton.Add_Click({ $script:HDTStoragePopup.WindowState = 'Minimized' })
    $script:HDTStorageMinimizeButton.FlatStyle = 'Flat'
    $script:HDTStorageMinimizeButton.FlatAppearance.BorderSize = 0
    $script:HDTStorageMinimizeButton.FlatAppearance.MouseOverBackColor = if ($SelectedTheme -eq 'Premium' -or $SelectedTheme -eq 'Custom') { $themeColors.AccentColor } else { $themeColors.ForeColor }
    $script:HDTStorageMinimizeButton.Add_MouseEnter({ if ($SelectedTheme -eq 'Premium' -or $SelectedTheme -eq 'Custom') { $script:HDTStorageMinimizeButton.ForeColor = $themeColors.ForeColor } else { $script:HDTStorageMinimizeButton.ForeColor = $themeColors.BackColor } })
    $script:HDTStorageMinimizeButton.Add_MouseLeave({ $script:HDTStorageMinimizeButton.ForeColor = $themeColors.ForeColor })

    # Button for closing the HDTStorage popup form
    $script:HDTStorageCloseButton = New-FormButton -Text "X" -LocationX 260 -LocationY 0 -Width 40 -BackColor $themeColors.BackColor -ForeColor $themeColors.ForeColor -Font $global:NormalBoldFont -Enabled $true -Visible $true
    $script:HDTStorageCloseButton.Add_Click({ $script:HDTStoragePopup.Close() })
    $script:HDTStorageCloseButton.FlatStyle = 'Flat'
    $script:HDTStorageCloseButton.FlatAppearance.BorderSize = 0
    $script:HDTStorageCloseButton.FlatAppearance.MouseOverBackColor = if ($SelectedTheme -eq 'Premium' -or $SelectedTheme -eq 'Custom') { $themeColors.AccentColor } else { $themeColors.ForeColor }
    $script:HDTStorageCloseButton.Add_MouseEnter({ if ($SelectedTheme -eq 'Premium' -or $SelectedTheme -eq 'Custom') { $script:HDTStorageCloseButton.ForeColor = $themeColors.ForeColor } else { $script:HDTStorageCloseButton.ForeColor = $themeColors.BackColor } })
    $script:HDTStorageCloseButton.Add_MouseLeave({ $script:HDTStorageCloseButton.ForeColor = $themeColors.ForeColor })

    # Label for the file location button
    $script:FileLocationLabel = New-FormLabel -LocationX 20 -LocationY 40 -SizeX 150 -SizeY 20 -Text 'File Location' -Font $global:NormalBoldFont -TextAlign $DefaultTextAlign

    # Button for selecting the file to upload
    $script:HDTStorageFileButton = New-FormButton -Text "Browse" -LocationX 20 -LocationY 60 -Width 75 -BackColor $ControlColor -ForeColor $ControlColor -Font $global:NormalFont -Enabled $true -Visible $true

    # Text box for entering the SQL instance
    $script:DBServerTextBox = New-FormTextBox -LocationX 20 -LocationY 110 -SizeX 200 -SizeY 20 -ScrollBars 'None' -Multiline $false -Enabled $true -ReadOnly $false -Text '' -Font $global:NormalFont

    # Label for the SQL instance text box
    $script:DBServerLabel = New-FormLabel -LocationX 20 -LocationY 90 -SizeX 150 -SizeY 20 -Text 'SQL Instance' -Font $global:NormalBoldFont -TextAlign $DefaultTextAlign

    # Text box for entering the table name
    $script:TableNameTextBox = New-FormTextBox -LocationX 20 -LocationY 160 -SizeX 200 -SizeY 20 -ScrollBars 'None' -Multiline $false -Enabled $true -ReadOnly $false -Text '' -Font $global:NormalFont

    # Label for the table name text box
    $script:TableNameLabel = New-FormLabel -LocationX 20 -LocationY 140 -SizeX 150 -SizeY 20 -Text 'Table Name' -Font $global:NormalBoldFont -TextAlign $DefaultTextAlign

    # Text box for entering secure password
    $script:SecurePasswordTextBox = New-FormTextBox -LocationX 20 -LocationY 210 -SizeX 200 -SizeY 20 -ScrollBars 'None' -Multiline $false -Enabled $true -ReadOnly $false -Text '' -Font $global:NormalFont
    $script:SecurePasswordTextBox.PasswordChar = '*'

    # Label for the secure password text box
    $script:SecurePasswordLabel = New-FormLabel -LocationX 20 -LocationY 190 -SizeX 150 -SizeY 20 -Text 'Your Allied Password' -Font $global:NormalBoldFont -TextAlign $DefaultTextAlign

    # Checkbox to use the user's PW entered in Password Manager
    $script:HDTStoragePWCheckbox = New-FormCheckbox -LocationX 20 -LocationY 240 -SizeX 200 -SizeY 20 -Text 'Use Your Saved PW?' -Font $global:NormalFont -Checked $false -Enabled (-not [string]::IsNullOrEmpty($global:SecurePW))

    # Button for running the script to create the HDTStorage table
    $script:CreateHDTStorageButton = New-FormButton -Text "Create" -LocationX 20 -LocationY 280 -Width 75 -BackColor $global:DisabledBackColor -ForeColor $global:DisabledForeColor -Font $global:NormalFont -Enabled $false -Visible $true

    # Check if DefaultUserTheme has a value or is null
    if ($null -ne $ConfigValues.DefaultUserTheme -and $ConfigValues.DefaultUserTheme -ne '') {
        # Get theme
        $SelectedTheme = $ColorTheme.PSObject.Properties | Where-Object { $_.Value.$($ConfigValues.DefaultUserTheme) } | Select-Object -ExpandProperty Name
        $themeColors = $ColorTheme.$SelectedTheme.$($ConfigValues.DefaultUserTheme)

        # Check if the theme falls under Premium and set the AccentColor accordingly
        if ($SelectedTheme -eq 'Premium' -or $SelectedTheme -eq 'Custom') {
            $script:DBServerTextBox.BackColor = $themeColors.AccentColor
            $script:TableNameTextBox.BackColor = $themeColors.AccentColor
            $script:SecurePasswordTextBox.BackColor = $themeColors.AccentColor
        } else {
            $script:DBServerTextBox.BackColor = [System.Drawing.SystemColors]::Control
            $script:TableNameTextBox.BackColor = [System.Drawing.SystemColors]::Control
            $script:SecurePasswordTextBox.BackColor = [System.Drawing.SystemColors]::Control
        }

        if ($ConfigValues.DefaultUserTheme -eq 'USA') {
            Enable-USAThemeTextColor
        }
        
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
        $LoginPassword = ConvertTo-SecureString $script:SecurePasswordTextBox.Text.Trim() -AsPlainText -Force
        $LoginCredentials = New-Object System.Management.Automation.PSCredential ($LoginUsername, $LoginPassword)

        # Set variables so values can be cleared from form
        $script:DBServerText = $script:DBServerTextBox.Text.Trim()
        $script:TableNameText = $script:TableNameTextBox.Text.Trim()

        # Clear form values
        $script:DBServerTextBox.Text = ''
        $script:TableNameTextBox.Text = ''
        $script:SecurePasswordTextBox.Text = ''
        $script:HDTStoragePWCheckbox.Checked = $false

        $OutText.AppendText("$(Get-Timestamp) - Creating HDTStorage table with the following values:`r`n")
        $OutText.AppendText("$(Get-Timestamp) - File Location: $($script:HDTStoragePopup.Tag)`r`n")
        $OutText.AppendText("$(Get-Timestamp) - SQL Instance: $script:DBServerText`r`n")
        $OutText.AppendText("$(Get-Timestamp) - Table Name: $script:TableNameText`r`n")
        $script:HDTStoragePopup.Close()
        Open-CreateHDTStorageRunspace -ConfigValuesCreateHDTStorageTableScript $ConfigValues.CreateHDTStorageTableScript -UserProfilePath $UserProfilePath -WorkhorseDirectoryPath $ConfigValues.WorkhorseDirectoryPath -WorkhorseServer $script:WorkhorseServer -FileTag $script:HDTStoragePopup.Tag -DBServerText $script:DBServerText -TableNameText $script:TableNameText -LoginCredentials $LoginCredentials -OutText $OutText -TimestampFunction ${function:Get-Timestamp}
    })

    # Click event for form close; sets global variable to false
    $script:HDTStoragePopup.Add_FormClosed({
        $global:IsHDTStoragePopupActive = $false
    })

    $script:HDTStoragePopup.Controls.Add($script:HDTStorageMinimizeButton)
    $script:HDTStoragePopup.Controls.Add($script:HDTStorageCloseButton)
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
$NewDocTextBox = New-FormTextBox -LocationX 265 -LocationY 400 -SizeX 140 -SizeY 20 -ScrollBars 'None' -Multiline $false -Enabled $true -ReadOnly $false -Text '' -Font $global:NormalFont

# Button for creating new documentation
$NewDocButton = New-FormButton -Text "Create" -LocationX 425 -LocationY 400 -Width 100 -BackColor $global:DisabledBackColor -ForeColor $global:DisabledForeColor -Font $global:NormalFont -Enabled $false -Visible $true

# Label for new documentation text box
$NewDocLabel = New-FormLabel -LocationX 265 -LocationY 375 -SizeX 137 -SizeY 20 -Text 'Create Documentation' -Font $global:NormalBoldFont -TextAlign $DefaultTextAlign

# Help Icon for new documentation text box
$NewDocHelpIcon = New-FormPictureBox -LocationX 402 -LocationY 375 -SizeX 20 -SizeY 20 -Image $HelpIconImage -SizeMode $AutoSizeMode -Cursor $HandCursor

# Event handler for the Documentation Help icon
$NewDocHelpIcon.Add_Click({
    Show-HelpForm -PictureBox $NewDocHelpIcon -HelpText "Create a new Microsoft Word document using the template in ..\Documentation\NewTemplate. The template will be updated with the title, today's date, and your name."
})

# If Documentation text box is empty, disable the Create button
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
$NewTicketButton = New-FormButton -Text "New Ticket" -LocationX 150 -LocationY 40 -Width 75 -BackColor $global:DisabledBackColor -ForeColor $global:DisabledForeColor -Font $global:NormalFont -Enabled $false -Visible $true

# Text box for new ticket
$NewTicketTextBox = New-FormTextBox -LocationX 5 -LocationY 40 -SizeX 140 -SizeY 20 -ScrollBars 'None' -Multiline $false -Enabled $true -ReadOnly $false -Text '' -Font $global:NormalFont

# Label for new ticket text box
$NewTicketLabel = New-FormLabel -LocationX 5 -LocationY 15 -SizeX 115 -SizeY 20 -Text 'Create New Ticket' -Font $global:NormalBoldFont -TextAlign $DefaultTextAlign

# Help Icon for new ticket text box
$NewTicketHelpIcon = New-FormPictureBox -LocationX 120 -LocationY 15 -SizeX 20 -SizeY 20 -Image $HelpIconImage -SizeMode $AutoSizeMode -Cursor $HandCursor

# Event handler for the New Ticket Help icon
$NewTicketHelpIcon.Add_Click({
    Show-HelpForm -PictureBox $NewTicketHelpIcon -HelpText "Create an .html file for a Jira ticket and place it inside a folder with the ticket name. Add scripts, logs, and other files to the ticket's folder to track work. Enter items by their Jira ticket number, i.e. CSH-12345, AIH-12345, etc."
})

# Button for renaming a ticket
$RenameTicketButton = New-FormButton -Text "Rename" -LocationX 455 -LocationY 40 -Width 75 -BackColor $global:DisabledBackColor -ForeColor $global:DisabledForeColor -Font $global:NormalFont -Enabled $false -Visible $true

# Text box for renaming a ticket
$RenameTicketTextBox = New-FormTextBox -LocationX 310 -LocationY 40 -SizeX 140 -SizeY 20 -ScrollBars 'None' -Multiline $false -Enabled $true -ReadOnly $false -Text '' -Font $global:NormalFont

# Label for renaming a ticket text box
$RenameTicketLabel = New-FormLabel -LocationX 310 -LocationY 15 -SizeX 95 -SizeY 20 -Text 'Rename Ticket' -Font $global:NormalBoldFont -TextAlign $DefaultTextAlign

# Help Icon for renaming a ticket text box
$RenameTicketHelpIcon = New-FormPictureBox -LocationX 405 -LocationY 15 -SizeX 20 -SizeY 20 -Image $HelpIconImage -SizeMode $AutoSizeMode -Cursor $HandCursor

# Event handler for the Rename Ticket Help icon
$RenameTicketHelpIcon.Add_Click({
    Show-HelpForm -PictureBox $RenameTicketHelpIcon -HelpText "Rename an active or completed Jira ticket .html file and its respective folder. Enter items by their Jira ticket number, i.e. CSH-12345, AIH-12345, etc."
})

# Tab control for ticket manager
$TicketManagerTabControl = New-Object System.Windows.Forms.TabControl
$TicketManagerTabControl.Location = "5,95"
$TicketManagerTabControl.Size = "220,255"

# Tab for active tickets
$ActiveTicketsTab = New-FormTabPage -Font $global:NormalFont -Name "ActiveTickets" -Text "Active Tickets"

# Tab for completed tickets
$CompletedTicketsTab = New-FormTabPage -Font $global:NormalFont -Name "CompletedTickets" -Text "Completed Tickets"

# List box to show folder contents
$FolderContentsListBox = New-FormListBox -LocationX 310 -LocationY 95 -SizeX 220 -SizeY 259 -Font $global:NormalFont -SelectionMode 'MultiExtended'

# Complete ticket button
$CompleteTicketButton = New-FormButton -Text "Complete" -LocationX 5 -LocationY 360 -Width 75 -BackColor $global:DisabledBackColor -ForeColor $global:DisabledForeColor -Font $global:NormalFont -Enabled $false -Visible $true

# Reactivate ticket button
$ReactivateTicketButton = New-FormButton -Text "Reactivate" -LocationX 5 -LocationY 360 -Width 75 -BackColor $global:DisabledBackColor -ForeColor $global:DisabledForeColor -Font $global:NormalFont -Enabled $false -Visible $false

# Open folder button
$OpenFolderButton = New-FormButton -Text "Open" -LocationX 150 -LocationY 360 -Width 75 -BackColor $global:DisabledBackColor -ForeColor $global:DisabledForeColor -Font $global:NormalFont -Enabled $false -Visible $true

# Button for opening an item in a folder
$OpenItemButton = New-FormButton -Text "Open" -LocationX 455 -LocationY 360 -Width 75 -BackColor $global:DisabledBackColor -ForeColor $global:DisabledForeColor -Font $global:NormalFont -Enabled $false -Visible $true

# List box for active tickets
$ActiveTicketsListBox = New-FormListBox -LocationX 0 -LocationY 0 -SizeX 215 -SizeY 240 -Font $global:NormalFont -SelectionMode 'MultiExtended'

# List box for completed tickets
$CompletedTicketsListBox = New-FormListBox -LocationX 0 -LocationY 0 -SizeX 215 -SizeY 240 -Font $global:NormalFont -SelectionMode 'MultiExtended'

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

# Logic for enabling the Open Item button
$FolderContentsListBox.Add_SelectedIndexChanged({
    if ($FolderContentsListBox.SelectedItems.Count -gt 0) {
        $OpenItemButton.Enabled = $true
        
        $backColor = Get-AppropriateColor -ColorType "BackColor"
        $foreColor = Get-AppropriateColor -ColorType "ForeColor"
        
        $OpenItemButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml($foreColor)
        $OpenItemButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($backColor)
    } else {
        $OpenItemButton.Enabled = $false
        $OpenItemButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledBackColor" -Scope Global -ValueOnly))
        $OpenItemButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml((Get-Variable -Name "DisabledForeColor" -Scope Global -ValueOnly))
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

# Event handler for opening one or more items in a folder
$OpenItemButton.Add_Click({
    if ($ActiveTicketsListBox.SelectedItems.Count -gt 0) {
        $tickets = $ActiveTicketsListBox.SelectedItems
        foreach ($ticket in $tickets){
            $selectedItems = $FolderContentsListBox.SelectedItems
            foreach ($selectedItem in $selectedItems){
                Invoke-Item "$TicketsPath\Active\$ticket\$selectedItem"
                $OutText.Appendtext("$(Get-Timestamp) - Opened item $selectedItem in folder for ticket $ticket`r`n")
            }
        }
    } else {
        $tickets = $CompletedTicketsListBox.SelectedItems
        foreach ($ticket in $tickets){
            $selectedItems = $FolderContentsListBox.SelectedItems
            foreach ($selectedItem in $selectedItems){
                Invoke-Item "$TicketsPath\Completed\$ticket\$selectedItem"
                $OutText.Appendtext("$(Get-Timestamp) - Opened item $selectedItem in folder for ticket $ticket`r`n")
            }
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
        $TicketNumber = $RenameTicketTextBox.Text.Trim()
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
$DesktopAssistantForm.Controls.Add($MainFormMinimizeButton)
$DesktopAssistantForm.Controls.Add($MainFormCloseButton)
$DesktopAssistantForm.Controls.Add($MainFormTabControl)
$MainFormTabControl.Controls.Add($SysAdminTab)
$MainFormTabControl.Controls.Add($AWSAdminTab)
$MainFormTabControl.Controls.Add($SupportTab)
$MainFormTabControl.Controls.Add($TicketManagerTab)
$DesktopAssistantForm.Controls.Add($OutText)
$DesktopAssistantForm.Controls.Add($ClearOutTextButton)
$DesktopAssistantForm.Controls.Add($SaveOutTextButton)
$DesktopAssistantForm.Controls.Add($MenuStrip)
$MenuStrip.Items.Add($FileMenu) | Out-Null
$MenuStrip.Items.Add($OptionsMenu) | Out-Null
$MenuStrip.Items.Add($AboutMenu) | Out-Null
$FileMenu.DropDownItems.Add($SubmitFeedback) | Out-Null
$FileMenu.DropDownItems.Add($MenuQuit) | Out-Null
$OptionsMenu.DropDownItems.Add($MenuColorTheme) | Out-Null
$OptionsMenu.DropDownItems.Add($MenuThemeBuilder) | Out-Null
$OptionsMenu.DropDownItems.Add($MenuFontPicker) | Out-Null
$OptionsMenu.DropDownItems.Add($MenuHelpOptions) | Out-Null
$MenuHelpOptions.DropDownItems.Add($ShowHelpIconsMenu) | Out-Null
$MenuHelpOptions.DropDownItems.Add($ShowToolTipsMenu) | Out-Null
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
$SysAdminTab.Controls.Add($CertCheckWizardButton)
$SysAdminTab.Controls.Add($CertCheckLabel)
$SysAdminTab.Controls.Add($AppListLabel)
$SysAdminTab.Controls.Add($RestartsSeparator)
$AWSAdminTab.Controls.Add($script:AWSAccountsListBox)
$AWSAdminTab.Controls.Add($script:AWSInstancesListBox)
$AWSAdminTab.Controls.Add($script:RebootAWSInstancesButton)
$AWSAdminTab.Controls.Add($script:StartAWSInstancesButton)
$AWSAdminTab.Controls.Add($script:StopAWSInstancesButton)
$AWSAdminTab.Controls.Add($script:AWSSSOLoginButton)
$AWSAdminTab.Controls.Add($script:AWSSSOLoginTimerLabel)
$AWSAdminTab.Controls.Add($script:ListAWSAccountsButton)
$AWSAdminTab.Controls.Add($script:AWSScreenshotButton)
$AWSAdminTab.Controls.Add($script:AWSCPUMetricsButton)
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
$SupportTab.Controls.Add($LaunchBillingRestartButton)
$SupportTab.Controls.Add($LaunchBillingRestartLabel)
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
$TicketManagerTab.Controls.Add($OpenItemButton)
$TicketManagerTab.Controls.Add($NewTicketLabel)
$TicketManagerTab.Controls.Add($RenameTicketLabel)
$TicketManagerTabControl.Controls.Add($ActiveTicketsTab)
$TicketManagerTabControl.Controls.Add($CompletedTicketsTab)
$ActiveTicketsTab.Controls.Add($ActiveTicketsListBox)
$CompletedTicketsTab.Controls.Add($CompletedTicketsListBox)

# Check if DefaultUserTheme has a value or is null
if ($null -ne $ConfigValues.DefaultUserTheme -and $ConfigValues.DefaultUserTheme -ne '') {
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
        $script:AWSAccountsListBox.Backcolor = $themeColors.AccentColor
        $script:AWSInstancesListBox.Backcolor = $themeColors.AccentColor
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
        $script:AWSAccountsListBox.Backcolor = [System.Drawing.SystemColors]::Control
        $script:AWSInstancesListBox.Backcolor = [System.Drawing.SystemColors]::Control
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
    $DesktopAssistantForm.BackColor = $themeColors.BackColor
    $MenuStrip.BackColor = $themeColors.BackColor
    $MenuStrip.ForeColor = $themeColors.ForeColor
    $MainFormTabControl.BackColor = $themeColors.BackColor
    $SysAdminTab.BackColor = $themeColors.BackColor
    $AWSAdminTab.BackColor = $themeColors.BackColor
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
    $CertCheckWizardButton.BackColor = $themeColors.ForeColor
    $CertCheckWizardButton.ForeColor = $themeColors.BackColor
    $CertCheckLabel.ForeColor = $themeColors.ForeColor
    $ReverseIPLabel.ForeColor = $themeColors.ForeColor
    $script:AWSSSOLoginTimerLabel.ForeColor = $themeColors.ForeColor
    $script:AWSAccountsListBox.BackColor = $OutText.BackColor
    $PSTComboLabel.ForeColor = $themeColors.ForeColor
    $PSTCombo.BackColor = $OutText.BackColor
    $RefreshPSTButton.BackColor = $themeColors.ForeColor
    $RefreshPSTButton.ForeColor = $themeColors.BackColor
    $LaunchLFPWizardButton.BackColor = $themeColors.ForeColor
    $LaunchLFPWizardButton.ForeColor = $themeColors.BackColor
    $LaunchLFPWizardLabel.ForeColor = $themeColors.ForeColor
    $LaunchBillingRestartButton.BackColor = $themeColors.ForeColor
    $LaunchBillingRestartButton.ForeColor = $themeColors.BackColor
    $LaunchBillingRestartLabel.ForeColor = $themeColors.ForeColor
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

    if ($ConfigValues.DefaultUserTheme -eq "Manny's") {
        # Add the Deadpool--YES, FREAKING DEADPOOL--to the form
        $DesktopAssistantForm.Controls.Add($DeadpoolsVeryOwnPictureBoxTM)
    }
}

# Enable tooltips if config value is set to Enabled
if ($ConfigValues.HoverToolTips -eq "Enabled" -or $null -eq $ConfigValues.HoverToolTips) {
	Enable-ToolTips
    $ShowToolTipsMenu.Text = "Hide Tool Tips"
}
else {
    $ShowToolTipsMenu.Text = "Show Tool Tips"
}

# Enable Help Icons if config value is set to Enabled
if ($ConfigValues.HelpIcons -eq "Enabled" -or $null -eq $ConfigValues.HelpIcons) {
    $SysAdminTab.Controls.Add($ServerPingHelpIcon)
	$SysAdminTab.Controls.Add($NSLookupHelpIcon)
	$SysAdminTab.Controls.Add($ReverseIPHelpIcon)
	$SysAdminTab.Controls.Add($RestartsGUIHelpIcon)
    $SysAdminTab.Controls.Add($CertCheckHelpIcon)
    $AWSAdminTab.Controls.Add($AWSMonitoringGUIHelpIcon)
	$SupportTab.Controls.Add($PSTHelpIcon)
	$SupportTab.Controls.Add($LFPWizardHelpIcon)
	$SupportTab.Controls.Add($BillingRestartHelpIcon)
	$SupportTab.Controls.Add($PWManagerHelpIcon)
	$SupportTab.Controls.Add($GenPWHelpIcon)
	$SupportTab.Controls.Add($HDTStorageHelpIcon)
	$SupportTab.Controls.Add($NewDocHelpIcon)
	$TicketManagerTab.Controls.Add($NewTicketHelpIcon)
	$TicketManagerTab.Controls.Add($RenameTicketHelpIcon)
    $ServerPingHelpIcon.BringToFront()
    $ShowHelpIconsMenu.Text = "Hide Help Icons"
}
else {
    $ShowHelpIconsMenu.Text = "Show Help Icons"
}

$OutText.AppendText("$(Get-Timestamp) - Welcome to the ETG Desktop Assistant!`r`n")

Get-ThemeQuote

# Confirm there are no active popups before closing the form
# Clean up sync hash when form is closed
# Reset the PST config file
$DesktopAssistantForm.Add_FormClosing({
    try {
        if ($global:IsLenderLFPPopupActive -or $global:IsBillingRestartPopupActive -or $global:IsHDTStoragePopupActive -or $global:IsFeedbackPopupActive -or $global:IsThemeBuilderPopupActive -or $global:IsColorPickerPopupActive -or $global:IsCertCheckPopupactive) {
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

    # Close Font Picker as well if it's active
    if ($global:IsFontPickerPopupActive) {
        $script:FontPickerPopup.Close()
    }
})

# Start the directory check jobs before showing the form
$RemoteSupportToolFolderDate = Start-DirectoryCheckJob -path $RemoteSupportTool -excludePattern "*.txt"
$LocalSupportToolFolderDate =  Start-DirectoryCheckJob -path $LocalSupportTool -excludePattern "*.txt"

# Show Form
$DesktopAssistantForm.ShowDialog() | Out-Null

<#
? **********************************************************************************************************************
? END OF FORM BUILD
? **********************************************************************************************************************
#>