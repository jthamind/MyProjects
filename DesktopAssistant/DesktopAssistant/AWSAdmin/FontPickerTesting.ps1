Add-Type -AssemblyName System.Data
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationCore
[System.Windows.Forms.Application]::EnableVisualStyles()

# Variables for form building functions
$ControlColor = [System.Drawing.SystemColors]::Control
$ControlColorText = [System.Drawing.SystemColors]::ControlText
$NormalFont = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
$NormalSmallFont = [System.Drawing.Font]::new("Arial", 8, [System.Drawing.FontStyle]::Regular)
$NormalBoldFont = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Bold)
$SmallBoldFont = [System.Drawing.Font]::new("Arial", 8, [System.Drawing.FontStyle]::Bold)
$ExtraSmallFont = [System.Drawing.Font]::new("Arial", 7, [System.Drawing.FontStyle]::Regular)
$ExtraSmallItalicBoldFont = [System.Drawing.Font]::new("Arial", 7, [System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Italic)
$ExtraLargeFont = [System.Drawing.Font]::new("Arial", 16, [System.Drawing.FontStyle]::Regular)
$DefaultTextAlign = [System.Drawing.ContentAlignment]::TopLeft
$MiddleTextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$TopRightTextAlign = [System.Drawing.ContentAlignment]::TopRight
$ArrowCursor = [System.Windows.Forms.Cursors]::Arrow
$HandCursor = [System.Windows.Forms.Cursors]::Hand
$AutoSizeMode = [System.Windows.Forms.PictureBoxSizeMode]::AutoSize
$ZoomSizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom

#Template for creating a new Windows form
function New-WindowsForm {
    param(
        [int]$SizeX,
        [int]$SizeY,
		[string]$Text,
		[System.Windows.Forms.FormStartPosition]$StartPosition,
		[bool]$TopMost,
		[System.Windows.Forms.FormBorderStyle]$FormBorderStyle,
		[bool]$ControlBox,
		[bool]$ShowInTaskbar,
		[bool]$KeyPreview,
		[bool]$MinimizeBox,
		[System.Drawing.Icon]$Icon
    )

    $Form = New-Object System.Windows.Forms.Form
    $Form.Size = New-Object System.Drawing.Size($SizeX, $SizeY)
	$Form.Text = $Text
	$Form.StartPosition = $StartPosition
	$Form.TopMost = $TopMost
	$Form.FormBorderStyle = $FormBorderStyle
	$Form.ControlBox = $ControlBox
	$Form.ShowInTaskbar = $ShowInTaskbar
	$Form.KeyPreview = $KeyPreview
	$Form.AutoSize = $true
	$Form.MaximizeBox = $false
	$Form.MinimizeBox = $MinimizeBox
	$Form.Icon = $Icon

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

# ! ==================================================================================================
# ! ==================================================================================================
# ! ==================================================================================================
# ! ==================================================================================================
# ! ==================================================================================================

# Launch the Font Picker form
$script:FontPickerPopup = New-WindowsForm -SizeX 380 -SizeY 635 -Text 'Font Picker' -StartPosition 'CenterScreen' -TopMost $false -FormBorderStyle 'Fixed3D' -ControlBox $true -ShowInTaskbar $true -KeyPreview $true -MinimizeBox $true -Icon $TeamIcon

# Create a panel with scrollbars
$script:FontPickerPanel = New-Object System.Windows.Forms.Panel
$script:FontPickerPanel.Location = New-Object System.Drawing.Point(5, 40)
$script:FontPickerPanel.Size = New-Object System.Drawing.Size(370, 540)
$script:FontPickerPanel.AutoScroll = $true

# Button for applying color choices
$script:FontPickerApplyButton = New-FormButton -Text "Apply Choices" -LocationX 140 -LocationY 10 -Width 100 -BackColor $ControlColor -ForeColor $ControlColorText -Font $NormalFont -Enabled $false -Visible $true

# Create a list box to hold the font names
$script:FontListBox = New-FormListBox -LocationX 0 -LocationY 0 -SizeX 370 -SizeY 540 -Font $Font -SelectionMode 'One'
$script:FontListBox.ItemHeight = 20
$script:FontListBox.DrawMode = 'OwnerDrawFixed'

# Add the list box to the form
$script:FontPickerPopup.Controls.Add($script:FontPickerApplyButton)
$script:FontPickerPopup.Controls.Add($script:FontListBox)
$script:FontPickerPopup.Controls.Add($script:FontPickerPanel)
$script:FontPickerPanel.Controls.Add($script:FontListBox)

# Event handler for drawing each item
$script:FontListBox.Add_DrawItem({
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
            $script:FontListBox.Items.Add($fontName) | Out-Null
        }
    } catch {
        # Handle exceptions if necessary
    }
}

# Show the form
$script:FontPickerPopup.ShowDialog() | Out-Null