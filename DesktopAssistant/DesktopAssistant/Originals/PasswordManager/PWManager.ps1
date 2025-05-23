Add-Type -AssemblyName System.Data
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationCore
[System.Windows.Forms.Application]::EnableVisualStyles()

# Sync hashtable for cross-threading
$Global:synchash = [hashtable]::Synchronized(@{})
$synchash.CWD = if ($PSScriptRoot) { $PSScriptRoot }
else { Split-Path -LiteralPath ([Environment]::GetCommandLineArgs()[0])}
$synchash.Closed = $False

# Function to clean up existing runspaces
function Stop-Runspace {
    $runspaces = Get-Runspace | Where-Object { $_.Id -gt 4 -and $_.Name -and $_.RunspaceAvailability -like "Available" }
 
    foreach ( $runspace in $runspaces ) {
        try {
            [void]$runspace.Close()
            [void]$runspace.Dispose()
        }
        catch {
            $_
        }
    }
}

# Cleans up any existing runspaces
Stop-Runspace

# Creates a new runspace
$newRunspace = [runspacefactory]::CreateRunspace()
$newRunspace.ApartmentState = "STA"
$newRunspace.ThreadOptions = "ReuseThread"
$newRunspace.Name = "PasswordManager"
$newRunspace.Open()

$GUIPowershell = [PowerShell]::Create().AddScript({

# Set location to script directory
Set-location $synchash.CWD
# Set timestamp function
function Get-Timestamp {
    return Get-Date -Format "yyyy/MM/dd hh:mm:ss"
}
    
# Declare the password variables in the global scope, allowing them to pass between functions and event handlers without losing assigned values.
$SecurePW = $null
$AltSecurePW = $null

# Initialize form
$Form = New-Object System.Windows.Forms.Form
$Form.Text = "Mr. Brobot"
$Form.Size = New-Object System.Drawing.Size(600, 300)
$Form.ShowInTaskbar = $True
$Form.KeyPreview = $True
$Form.AutoSize = $True
$Form.FormBorderStyle = 'Fixed3D'
$Form.MaximizeBox = $False
$Form.MinimizeBox = $True
$Form.ControlBox = $True
$Form.BackColor = "Black"
$Form.Icon = $Icon
$Form.TopMost = $True
$Form.StartPosition = "CenterScreen"

# Output text box
$OutText = New-Object System.Windows.Forms.TextBox
$OutText.Location = New-Object System.Drawing.Size(25, 75)
$OutText.Size = New-Object System.Drawing.Size(350, 200)
$OutText.Multiline = $true
$OutText.ScrollBars = "Vertical"
$OutText.Enabled = $True
$OutText.ReadOnly = $True
$OutText.BackColor = "White"
$OutText.ForeColor = "Black"

# Set password text box
$PWTxtBx = New-Object System.Windows.Forms.TextBox
$PWTxtBx.Location = New-Object System.Drawing.Size(10, 25)
$PWTxtBx.Text = ''
$PWTxtBx.BackColor = "Yellow"
$PWTxtBx.ForeColor = "Black"
$PWTxtBx.ShortcutsEnabled = $True

# Button for setting your own password
$SetPWBtn = New-Object System.Windows.Forms.Button
$SetPWBtn.Location = New-Object System.Drawing.Size(260, 25)
$SetPWBtn.Width = 100
$SetPWBtn.Text = "Set Your PW"
$SetPWBtn.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
$SetPWBtn.Enabled = $false
$SetPWBtn.BackColor = "Yellow"
$SetPWBtn.ForeColor = "Black"

# Button for setting the alt password
$AltSetPWBtn = New-Object System.Windows.Forms.Button
$AltSetPWBtn.Location = New-Object System.Drawing.Size(260, 50)
$AltSetPWBtn.Width = 100
$AltSetPWBtn.Text = "Set Alt PW"
$AltSetPWBtn.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
$AltSetPWBtn.Enabled = $false
$AltSetPWBtn.BackColor = "Yellow"
$AltSetPWBtn.ForeColor = "Black"

# Button for copying your own password
$CopyPWBtn = New-Object System.Windows.Forms.Button
$CopyPWBtn.Location = New-Object System.Drawing.Size(260, 25)
$CopyPWBtn.Width = 100
$CopyPWBtn.Text = "Copy Your PW"
$CopyPWBtn.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
$CopyPWBtn.Enabled = $false
$CopyPWBtn.Visible = $false
$CopyPWBtn.BackColor = "Yellow"
$CopyPWBtn.ForeColor = "Black"

# Button for setting the alt password
$AltCopyPWBtn = New-Object System.Windows.Forms.Button
$AltCopyPWBtn.Location = New-Object System.Drawing.Size(260, 50)
$AltCopyPWBtn.Width = 100
$AltCopyPWBtn.Text = "Copy Alt PW"
$AltCopyPWBtn.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
$AltCopyPWBtn.Enabled = $false
$AltCopyPWBtn.Visible = $false
$AltCopyPWBtn.BackColor = "Yellow"
$AltCopyPWBtn.ForeColor = "Black"

# Button for clearing your own password
$ClearPWBtn = New-Object System.Windows.Forms.Button
$ClearPWBtn.Location = New-Object System.Drawing.Size(150, 25)
$ClearPWBtn.Width = 90
$ClearPWBtn.Text = "Clear Yours"
$ClearPWBtn.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
$ClearPWBtn.Enabled = $false
$ClearPWBtn.BackColor = "Yellow"
$ClearPWBtn.ForeColor = "Black"

# Button for clearing the alt password
$AltClearPWBtn = New-Object System.Windows.Forms.Button
$AltClearPWBtn.Location = New-Object System.Drawing.Size(150, 50)
$AltClearPWBtn.Width = 90
$AltClearPWBtn.Text = "Clear Alt"
$AltClearPWBtn.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
$AltClearPWBtn.Enabled = $false
$AltClearPWBtn.BackColor = "Yellow"
$AltClearPWBtn.ForeColor = "Black"


# If text box is empty, disable the Set Password button
$PWTxtBx.Add_TextChanged({
    if ($PWTxtBx.Text -ne '') {
        $SetPWBtn.Enabled = $True
        $AltSetPWBtn.Enabled = $True
    }
    else {
        $SetPWBtn.Enabled = $False
        $AltSetPWBtn.Enabled = $False
    }
})

function Set-MyPassword {
    $PWTxtBx.Text | Set-Clipboard
    $global:SecurePW = $PWTxtBx.Text | ConvertTo-SecureString -AsPlainText -Force
    $OutText.AppendText("$(Get-Timestamp) - Your password has been set and copied to the clipboard.`r`n")
    $PWTxtBx.Text = ''
    $SetPWBtn.Enabled = $False
    $SetPWBtn.Visible = $False
    $CopyPWBtn.Enabled = $True
    $CopyPWBtn.Visible = $True
    $ClearPWBtn.Enabled = $True
}

function Set-AltPassword {
    $PWTxtBx.Text | Set-Clipboard
    $global:AltSecurePW = $PWTxtBx.Text | ConvertTo-SecureString -AsPlainText -Force
    $OutText.AppendText("$(Get-Timestamp) - Your password has been set and copied to the clipboard.`r`n")
    $PWTxtBx.Text = ''
    $AltSetPWBtn.Enabled = $False
    $AltSetPWBtn.Visible = $False
    $AltCopyPWBtn.Enabled = $True
    $AltCopyPWBtn.Visible = $True
    $AltClearPWBtn.Enabled = $True
}

# Set my password button logic
$SetPWBtn.Add_Click({
    Set-MyPassword
})

# Set alt password button logic
$AltSetPWBtn.Add_Click({
    Set-AltPassword
})

# Clear my password button logic
$ClearPWBtn.Add_Click({
    $PWTxtBx.Text = ''
    $global:SecurePW = $null
    Set-Clipboard -Value ''
    $ClearPWBtn.Enabled = $False
    $SetPWBtn.Enabled = $True
    $SetPWBtn.Visible = $True
    $CopyPWBtn.Enabled = $False
    $CopyPWBtn.Visible = $False
    $OutText.AppendText("$(Get-Timestamp) - Your password has been cleared.`r`n")
})

# Clear alt password button logic
$AltClearPWBtn.Add_Click({
    $PWTxtBx.Text = ''
    $global:AltSecurePW = $null
    Set-Clipboard -Value ''
    $AltClearPWBtn.Enabled = $False
    $AltSetPWBtn.Enabled = $True
    $AltSetPWBtn.Visible = $True
    $AltCopyPWBtn.Enabled = $False
    $AltCopyPWBtn.Visible = $False
    $OutText.AppendText("$(Get-Timestamp) - Your password has been cleared.`r`n")
})

# Copy my password button logic
$CopyPWBtn.Add_Click({
    $PlainTxtPW = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($global:SecurePW))
    $PlainTxtPW | Set-Clipboard
    $OutText.AppendText("$(Get-Timestamp) - Your password has been copied to the clipboard.`r`n")
})

# Copy alt password button logic
$AltCopyPWBtn.Add_Click({
    $PlainTxtAltPW = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($global:AltSecurePW))
    $PlainTxtAltPW | Set-Clipboard
    $OutText.AppendText("$(Get-Timestamp) - Alt password has been copied to the clipboard.`r`n")
})

<# ! No longer viable since there are two buttons that can set the password
# Password text box Enter key logic
$PWTxtBx.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") {
        Set-MyPassword        
    }
})
#>

# Password text box Ctrl+A logic
$PWTxtBx.Add_KeyDown({
    if ($_.KeyCode -eq "A" -and $_.Control -eq "True") {
        $PWTxtBx.SelectAll()    
    }
    })

# Build form
$Form.Controls.Add($OutText)
$Form.Controls.Add($SetPWBtn)
$Form.Controls.Add($AltSetPWBtn)
$Form.Controls.Add($ClearPWBtn)
$Form.Controls.Add($AltClearPWBtn)
$Form.Controls.Add($PWTxtBx)
$Form.Controls.Add($CopyPWBtn)
$Form.Controls.Add($AltCopyPWBtn)

# Show Form
$Form.ShowDialog() | Out-Null
})

$GUIPowershell.Runspace = $newRunspace
$async = $GUIPowershell.BeginInvoke()
$GuiPowerShell.EndInvoke($async)