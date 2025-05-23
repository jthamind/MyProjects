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

# Test computer connection button
$TestConnButton = New-Object System.Windows.Forms.Button
$TestConnButton.Location = New-Object System.Drawing.Size(125, 25)
$TestConnButton.Text = "Test"
$TestConnButton.Enabled = $False
$TestConnButton.BackColor = "Yellow"
$TestConnButton.ForeColor = "Black"

# Enter computer text box
$EnterCompTextBox = New-Object System.Windows.Forms.TextBox
$EnterCompTextBox.Location = New-Object System.Drawing.Size(10, 25)
$EnterCompTextBox.Text = ''
$EnterCompTextBox.BackColor = "Yellow"
$EnterCompTextBox.ForeColor = "Black"
$EnterCompTextBox.ShortcutsEnabled = $True

# Computer text box logic
# If text box is empty, disable the Test button
$EnterCompTextBox.Add_TextChanged({
    if ($EnterCompTextBox.Text -ne '') {
        $TestConnButton.Enabled = $True 
    }
    else {
        $TestConnButton.Enabled = $False
    }
})

# Function to test computer connection
function Test-ComputerConnection {
    param($ComputerName)
    Test-Connection -ComputerName $ComputerName
}

# Test connection button logic
$TestConnButton.Add_Click({
    $ComputerName = $EnterCompTextBox.Text
    $OutText.AppendText("$(Get-Timestamp) - Testing connection to $ComputerName...`r`n")
    $PingResult = Test-ComputerConnection -ComputerName $ComputerName
    # $OutText.AppendText($PingResult)
    if ($PingResult -ne $null) {
        $OutText.AppendText("$(Get-Timestamp) - Connection to $ComputerName successful.`r`n")
    }
    else {
        $OutText.AppendText("$(Get-Timestamp) - Connection to $ComputerName failed.`r`n")
    }
})

# Computer text box Enter key logic
$EnterCompTextBox.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") {
        Test-ComputerConnection -ComputerName $ComputerName       
    }
})

# Computer text box Ctrl+A logic
$EnterCompTextBox.Add_KeyDown({
    if ($_.KeyCode -eq "A" -and $_.Control -eq "True") {
        $EnterCompTextBox.SelectAll()
    }
})

# Build form
$Form.Controls.Add($OutText)
$Form.Controls.Add($TestConnButton)
$Form.Controls.Add($EnterCompTextBox)

# Show Form
$Form.ShowDialog() | Out-Null

})

$GUIPowershell.Runspace = $newRunspace
$async = $GUIPowershell.BeginInvoke()
$GuiPowerShell.EndInvoke($async)