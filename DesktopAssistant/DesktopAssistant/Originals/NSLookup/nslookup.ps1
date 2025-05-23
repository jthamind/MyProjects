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
$newRunspace.Name = "GetServerIP"
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

# Button to run Resolve-DnsName
$RunLookupButton = New-Object System.Windows.Forms.Button
$RunLookupButton.Location = New-Object System.Drawing.Size(125, 25)
$RunLookupButton.Text = "Lookup"
$RunLookupButton.Enabled = $False
$RunLookupButton.BackColor = "Yellow"
$RunLookupButton.ForeColor = "Black"

# Run Resolve-DnsName text box
$RunLookupTextBox = New-Object System.Windows.Forms.TextBox
$RunLookupTextBox.Location = New-Object System.Drawing.Size(10, 25)
$RunLookupTextBox.Text = ''
$RunLookupTextBox.BackColor = "Black"
$RunLookupTextBox.ForeColor = "Yellow"

# If text box is empty, disable the Create button
$RunLookupTextBox.Add_TextChanged({
    if ($RunLookupTextBox.Text -ne '') {
        $RunLookupButton.Enabled = $True 
    }
    else {
        $RunLookupButton.Enabled = $False
    }
})

# Function to resolve a DNS name and return an IP address
Function Get-DnsName {
    param (
        [string]$Computer
    )
    try {
        $SelectedObject = Resolve-DnsName -Name $Computer -ErrorAction Stop
        if ($SelectedObject -and $SelectedObject.IPAddress) {
            $IPAddress = $SelectedObject.IPAddress
            $OutText.AppendText("$(Get-Timestamp) - IP Address = $IPAddress`r`n")
        } else {
            $OutText.AppendText("$(Get-Timestamp) - An error occurred while resolving the DNS name. Please ensure you're entering a valid hostname.`r`n")
        }
    }
    catch {
        $OutText.AppendText("$(Get-Timestamp) - An error occurred while resolving the DNS name. Please ensure you're entering a valid hostname.`r`n")
    }
}

# Function to test hostname pattern and run Resolve-DnsName
function Test-Hostname {
    param (
        [string]$Hostname
    )
    $Computer = $RunLookupTextBox.Text
    $Pattern = "^(?=.{1,255}$)(?!-)[a-zA-Z0-9-]{1,63}(?<!-)(\.[a-zA-Z0-9-]{1,63})*$"

    if ($Computer -match $Pattern) {
        Get-DnsName -Computer $Computer
    }
    else {
        $OutText.AppendText("$(Get-Timestamp) - Please ensure you're entering a valid hostname`r`n")
    }
}

# Resolve-DnsName button press logic
$RunLookupButton.Add_Click({
    Test-Hostname -Hostname $RunLookupTextBox.Text
})

# Resolve-DnsName text box Enter key logic
$RunLookupTextBox.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") {
        Test-Hostname -Hostname $RunLookupTextBox.Text        
    }
})

# Build Form
$Form.Controls.Add($RunLookupButton)
$Form.Controls.Add($RunLookupTextBox)
$Form.Controls.Add($OutText)

# Show Form
$Form.ShowDialog() | Out-Null
})

$GUIPowershell.Runspace = $newRunspace
$async = $GUIPowershell.BeginInvoke()
$GuiPowerShell.EndInvoke($async)