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
$newRunspace.SessionStateProxy.SetVariable("synchash", $synchash)

$MainGUI = [PowerShell]::Create().AddScript({

# Set location to script directory
Set-location $synchash.CWD
# Set timestamp function
function Get-Timestamp {
    return Get-Date -Format "yyyy/MM/dd hh:mm:ss"
}

************************************************************************************
<# ? START OF MAIN GUI BUILD #>
************************************************************************************

# Initialize form
$Form = New-Object System.Windows.Forms.Form
$Form.Text = "ETG Desktop Assistant"
$Form.Size = New-Object System.Drawing.Size(1000, 600)
$Form.ShowInTaskbar = $True
$Form.KeyPreview = $True
$Form.AutoSize = $True
$Form.FormBorderStyle = "Fixed3D"
$Form.BorderColor = "#0060a9"
$Form.MaximizeBox = $False
$Form.MinimizeBox = $True
$Form.ControlBox = $True
$Form.Icon = $Icon
$Form.TopMost = $True
$Form.StartPosition = "CenterScreen"
$Form.BackColor = "#0060a9"

# Tab control creation
$MainFormTabControl = New-object System.Windows.Forms.TabControl
$MainFormTabControl.Size = "590,500"
$MainFormTabControl.Location = "5,65"
$MainFormTabControl.BackColor = "#0060a9"
$MainFormTabControl.BorderThickness = "0"
$MainFormTabControl.BorderStyle = "None"

# System Administator Tools Tab
$SysAdminTab = New-Object System.Windows.Forms.TabPage
$SysAdminTab.DataBindings.DefaultDataSourceUpdateMode = 0
$SysAdminTab.UseVisualStyleBackColor = $True
$SysAdminTab.BackColor = "#0060a9"
$SysAdminTab.Name = 'SysAdminTools'
$SysAdminTab.Text = 'SysAdmin Tools'
$SysAdminTab.Font = [System.Drawing.Font]::new("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)

# Support Tools Tab
$SupportTab = New-Object System.Windows.Forms.TabPage
$SupportTab.DataBindings.DefaultDataSourceUpdateMode = 0
$SupportTab.UseVisualStyleBackColor = $True
$SupportTab.BackColor = "#0060a9"
$SupportTab.Name = 'SupportTools'
$SupportTab.Text = 'Support Tools'
$SupportTab.Font = [System.Drawing.Font]::new("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)

# Ticket Manager Tab
$TicketManagerTab = New-Object System.Windows.Forms.TabPage
$TicketManagerTab.DataBindings.DefaultDataSourceUpdateMode = 0
$TicketManagerTab.UseVisualStyleBackColor = $True
$TicketManagerTab.BackColor = "#0060a9"
$TicketManagerTab.Name = 'TicketManager'
$TicketManagerTab.Text = 'Ticket Manager'
$TicketManagerTab.Font = [System.Drawing.Font]::new("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)

# Create a custom Panel control for the transparent banner
$bannerPanel = New-Object System.Windows.Forms.Panel
$bannerPanel.Size = New-Object System.Drawing.Size($Form.Width, 60)
$bannerPanel.Location = New-Object System.Drawing.Point(0, 0)
$bannerPanel.BackColor = [System.Drawing.Color]::FromArgb(0, 0, 0, 0)  # Transparent

# Create a Label control inside the transparent banner
$bannerLabel = New-Object System.Windows.Forms.Label
$bannerLabel.Text = "ETG Desktop Assistant"
$bannerLabel.Font = New-Object System.Drawing.Font("Segoe UI", 20, [System.Drawing.FontStyle]::Bold)
$bannerLabel.UseCompatibleTextRendering = $true  # Enable text antialiasing
$bannerLabel.ForeColor = [System.Drawing.Color]::Red
$bannerLabel.Dock = [System.Windows.Forms.DockStyle]::Fill
$bannerLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter

# Create a textbox for logging
$OutText = New-Object System.Windows.Forms.TextBox
$OutText.Size = New-Object System.Drawing.Size(400, 480)
$OutText.Location = New-Object System.Drawing.Point(600, 85)
$OutText.UseCompatibleTextRendering = $true  # Enable text antialiasing
$OutText.Multiline = $true
$OutText.ScrollBars = "Vertical"
$OutText.Enabled = $True
$OutText.ReadOnly = $True

************************************************************************************
<# ? END OF MAIN GUI BUILD #>
************************************************************************************

************************************************************************************
<# ? START OF RESTARTS GUI #>
************************************************************************************

# Tab control creation
$RestartsTabControl = New-object System.Windows.Forms.TabControl
$RestartsTabControl.Size = "250,250"
$RestartsTabControl.Location = "225,75"

# Individual servers list box
$ServersListBox = New-Object System.Windows.Forms.ListBox
$ServersListBox.Location = New-Object System.Drawing.Point(5,95)
$ServersListBox.Size = New-Object System.Drawing.Size(200,240)
$ServersListBox.DockStyle = "Fill"
$ServersListBox.SelectionMode = 'Single'

# Services list box
$ServicesListBox = New-Object System.Windows.Forms.ListBox
$ServicesListBox.Location = New-Object System.Drawing.Point(0,0)
$ServicesListBox.Size = New-Object System.Drawing.Size(245,240)
$ServicesListBox.DockStyle = "Fill"
$ServicesListBox.SelectionMode = 'MultiExtended'

# IIS App Pools list box
$AppPoolsListBox = New-Object System.Windows.Forms.ListBox
$AppPoolsListBox.Location = New-Object System.Drawing.Point(0,0)
$AppPoolsListBox.Size = New-Object System.Drawing.Size(245,240)
$AppPoolsListBox.DockStyle = "Fill"
$AppPoolsListBox.SelectionMode = 'MultiExtended'

# Combobox for application selection
$AppListCombo = New-Object System.Windows.Forms.ComboBox
$AppListCombo.Location = New-Object System.Drawing.Point(5,65)
$AppListCombo.Size = New-Object System.Drawing.Size(200, 200)
$AppListCombo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$AppListCombo.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)

# Tab for services list
$ServicesTab = New-Object System.Windows.Forms.TabPage
$ServicesTab.DataBindings.DefaultDataSourceUpdateMode = 0
$ServicesTab.UseVisualStyleBackColor = $True
$ServicesTab.Name = 'ServicesTab'
$ServicesTab.Text = 'Services'
$ServicesTab.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)

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
$RestartButton.Width = 65
$RestartButton.BackColor = "White"
$RestartButton.ForeColor = "#0060a9"
$RestartButton.FlatStyle = "Popup"
$RestartButton.Text = "Restart"
$RestartButton.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
$RestartButton.Enabled = $false

# Button for starting services
$StartButton = New-Object System.Windows.Forms.Button
$StartButton.Location = New-Object System.Drawing.Point(490, 125)
$StartButton.Width = 65
$StartButton.BackColor = "White"
$StartButton.ForeColor = "#0060a9"
$StartButton.FlatStyle = "Popup"
$StartButton.Text = "Start"
$StartButton.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
$StartButton.Enabled = $false

# Button for stopping services
$StopButton = New-Object System.Windows.Forms.Button
$StopButton.Location = New-Object System.Drawing.Point(490, 155)
$StopButton.Width = 65
$StopButton.BackColor = "White"
$StopButton.ForeColor = "#0060a9"
$StopButton.FlatStyle = "Popup"
$StopButton.Text = "Stop"
$StopButton.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
$StopButton.Enabled = $false

# Disable restart button if no options are selected in the ServicesListBox. Enable if options are selected.
$ServicesListBox.add_SelectedIndexChanged({
    if ($ServicesListBox.SelectedItems.Count -gt 0) {
        $RestartButton.Enabled = $true
        $StartButton.Enabled = $true
        $StopButton.Enabled = $true
    } else {
        $RestartButton.Enabled = $false
        $StartButton.Enabled = $false
        $StopButton.Enabled = $false
    }
})

# Disable restart button if no options are selected in the AppPoolsListBox. Enable if options are selected.
$AppPoolsListBox.add_SelectedIndexChanged({
    if ($AppPoolsListBox.SelectedItems.Count -gt 0) {
        $RestartButton.Enabled = $true
        $StartButton.Enabled = $true
        $StopButton.Enabled = $true
    } else {
        $RestartButton.Enabled = $false
        $StartButton.Enabled = $false
        $StopButton.Enabled = $false
    }
})

# Variables for importing server list from CSV
$csvPath = "$env:USERPROFILE\OneDrive - Allied Solutions\Documents\Documentation\servers.csv"
$ServerCSV = Import-CSV $csvPath
$csvHeaders = ($ServerCSV | Get-Member -MemberType NoteProperty).name

# Populate the AppListCombo with the list of applications from the CSV
foreach ($header in $csvHeaders) {
    [void]$AppListCombo.Items.Add($header)
}

function PopulateListBox {
    $SelectedTab = $FormTabControl.SelectedTab.Text
    $synchash.SelectedTab = $SelectedTab

    # Get the selected server from the $ServersListBox
    $SelectedServer = $ServersListBox.SelectedItem
    if ($null -eq $SelectedServer) {
        return
    }

    switch ($synchash.SelectedTab) {
        "Services" {
            $ServicesListBox.Items.Clear()  # Clear the list box before populating with fresh data
            try {
                $Services = Invoke-Command -ComputerName $SelectedServer -ScriptBlock {
                    Get-Service | ForEach-Object { $_.DisplayName } | Sort-Object
                }
                foreach ($service in $Services) {
                    [void]$ServicesListBox.Items.Add($service)
                }
            } catch {
                $synchash.OutText.AppendText("$(Get-Timestamp) - Error retrieving service: $($_.Exception.Message)`r`n")
            }
        }
       <# ! Logic to add Sites to list box - NOT IN USE
        "Sites" {
            try {
                $IISSites = Invoke-Command -ComputerName $SelectedServer -ScriptBlock {
                    Import-Module WebAdministration
                    Get-IISSite | ForEach-Object { $_.Name } | Sort-Object
                }
                foreach ($item in $IISSites) {
                    [void]$IISSitesListBox.Items.Add($item)
                }
            } catch {
                $synchash.OutText.AppendText("$(Get-Timestamp) - Error retrieving IIS Sites: $($_.Exception.Message)`r`n")
            }
        }#>
        "App Pools" {
            $AppPoolsListBox.Items.Clear()  # Clear the list box before populating with fresh data
            try {
                $AppPools = Invoke-Command -ComputerName $SelectedServer -ScriptBlock {
                    Import-Module WebAdministration
                    Get-IISAppPool | ForEach-Object { $_.Name } | Sort-Object
                }
                if ($null -eq $AppPools) {
                    [void]$AppPoolsListBox.Items.Add("No AppPools found")
                }
                foreach ($apppool in $AppPools) {
                    [void]$AppPoolsListBox.Items.Add($apppool)
                }
            } catch {
                $synchash.OutText.AppendText("$(Get-Timestamp) - Error retrieving AppPools: $($_.Exception.Message)`r`n")
            }
        }
    }
}

# Event handler for the AppListCombo's SelectedIndexChanged event
$AppListCombo_SelectedIndexChanged = {
    $selectedHeader = $AppListCombo.SelectedItem.ToString()
    $servers = $ServerCSV | ForEach-Object { $_.$selectedHeader } | Where-Object { $_ -ne '' }
    $ServersListBox.Items.Clear()
    $servers | ForEach-Object {
        [void]$ServersListBox.Items.Add($_)
    }
}

# Button click event handler for restarting services
$RestartButton.Add_Click({
    $OutText.AppendText("Restarting $services on $SelectedServer `r`n")
    if ($SelectedTab -eq "Services") {
        foreach ($item in $ServicesListBox.SelectedItems) {
            try {
                Invoke-Command -ComputerName $SelectedServer -ScriptBlock {
                    Restart-Service -Name $using:item
                }
                $OutText.AppendText("$(Get-Timestamp) - Restarted $item`r`n")
            } catch {
                $OutText.AppendText("$(Get-Timestamp) - Error restarting service: $($_.Exception.Message)`r`n")
            }
        }
    } elseif ($SelectedTab -eq "App Pools") {
        foreach ($item in $AppPoolsListBox.SelectedItems) {
            try {
                Invoke-Command -ComputerName $SelectedServer -ScriptBlock {
                    Import-Module WebAdministration
                    Restart-WebAppPool -Name $using:item
                }
                $OutText.AppendText("$(Get-Timestamp) - Restarted $item`r`n")
            } catch {
                $OutText.AppendText("$(Get-Timestamp) - Error restarting AppPool: $($_.Exception.Message)`r`n")
            }
        }
    }
})

# Button click event handler for starting services
$StartButton.Add_Click({
    if ($SelectedTab -eq "Services") {
        foreach ($item in $ServicesListBox.SelectedItems) {
            try {
                Invoke-Command -ComputerName $SelectedServer -ScriptBlock {
                    Start-Service -Name $using:item
                }
                $OutText.AppendText("$(Get-Timestamp) - Started $item`r`n")
            } catch {
                $OutText.AppendText("$(Get-Timestamp) - Error starting service: $($_.Exception.Message)`r`n")
            }
        }
    } elseif ($SelectedTab -eq "App Pools") {
        foreach ($item in $AppPoolsListBox.SelectedItems) {
            try {
                Invoke-Command -ComputerName $SelectedServer -ScriptBlock {
                    Import-Module WebAdministration
                    Start-WebAppPool -Name $using:item
                }
                $OutText.AppendText("$(Get-Timestamp) - Started $item`r`n")
            } catch {
                $OutText.AppendText("$(Get-Timestamp) - Error starting AppPool: $($_.Exception.Message)`r`n")
            }
        }
    }
})

# Button click event handler for stopping services
$StopButton.Add_Click({
    if ($SelectedTab -eq "Services") {
        foreach ($item in $ServicesListBox.SelectedItems) {
            try {
                Invoke-Command -ComputerName $SelectedServer -ScriptBlock {
                    Stop-Service -Name $using:item
                }
                $OutText.AppendText("$(Get-Timestamp) - Stopped $item`r`n")
            } catch {
                $OutText.AppendText("$(Get-Timestamp) - Error stopping service: $($_.Exception.Message)`r`n")
            }
        }
    } elseif ($SelectedTab -eq "App Pools") {
        foreach ($item in $AppPoolsListBox.SelectedItems) {
            try {
                Invoke-Command -ComputerName $SelectedServer -ScriptBlock {
                    Import-Module WebAdministration
                    Stop-WebAppPool -Name $using:item
                }
                $OutText.AppendText("$(Get-Timestamp) - Stopped $item`r`n")
            } catch {
                $OutText.AppendText("$(Get-Timestamp) - Error stopping AppPool: $($_.Exception.Message)`r`n")
            }
        }
    }
})

# Function to handle the event when a server is selected in the $ServersListBox
function OnServerSelected {
    $SelectedServer = $ServersListBox.SelectedItem
    if ($null -ne $SelectedServer) {
        # Call the PopulateListBox function passing the selected server
        PopulateListBox $SelectedServer
    }
}

# Function to show service status when a server is selected
function OnServiceSelected {
    $SelectedServer = $ServersListBox.SelectedItem
    if ($null -ne $SelectedServer) {
        $SelectedService = $ServicesListBox.SelectedItem
        if ($null -ne $SelectedService) {
            try {
                $ServiceStatus = Invoke-Command -ComputerName $SelectedServer -ScriptBlock {
                    Get-Service -DisplayName $using:SelectedService
                }
                $OutText.AppendText("$(Get-Timestamp) - $SelectedService status: $($ServiceStatus.Status)`r`n")
            } catch {
                $OutText.AppendText("$(Get-Timestamp) - Error retrieving service status: $($_.Exception.Message)`r`n")
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
            try {
                $AppPoolStatus = Invoke-Command -ComputerName $SelectedServer -ScriptBlock {
                    Import-Module WebAdministration
                    Get-WebAppPoolState -Name $using:SelectedAppPool
                }
                $OutText.AppendText("$(Get-Timestamp) - $SelectedAppPool status: $($AppPoolStatus.Value)`r`n")
            } catch {
                $OutText.AppendText("$(Get-Timestamp) - Error retrieving AppPool status: $($_.Exception.Message)`r`n")
            }
        }
    }
}

# Add the event handler to the $ServersListBox
$ServersListBox.add_SelectedIndexChanged({ OnServerSelected })

# Add the event handler to the $ServicesListBox
$ServicesListBox.add_SelectedIndexChanged({ OnServiceSelected })

# Add the event handler to the $AppPoolsListBox
$AppPoolsListBox.add_SelectedIndexChanged({ OnAppPoolSelected })

# Event handler for TabControl's SelectedIndexChanged event
$RestartsTabControl_SelectedIndexChanged = {
    $RestartsTabControl.SelectedTab.Text
    PopulateListBox
}

$RestartsTabControl.add_SelectedIndexChanged($RestartsTabControl_SelectedIndexChanged)

$RestartsTabControl.Controls.Add($ServicesTab)
$RestartsTabControl.Controls.Add($AppPoolsTab)
$AppListCombo.add_SelectedIndexChanged($AppListCombo_SelectedIndexChanged)

************************************************************************************
<# ? END OF RESTARTS GUI #>
************************************************************************************

************************************************************************************
<# ? START OF NSLOOKUP GUI #>
************************************************************************************

# Button to run Resolve-DnsName
$RunLookupButton = New-Object System.Windows.Forms.Button
$RunLookupButton.Location = New-Object System.Drawing.Size(410, 400)
$RunLookupButton.Width = 65
$RunLookupButton.BackColor = "White"
$RunLookupButton.ForeColor = "#0060a9"
$RunLookupButton.FlatStyle = "Popup"
$RunLookupButton.Text = "Lookup"
$RunLookupButton.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
$RunLookupButton.Enabled = $false

# Run Resolve-DnsName text box
$RunLookupTextBox = New-Object System.Windows.Forms.TextBox
$RunLookupTextBox.Location = New-Object System.Drawing.Size(225, 400)
$RunLookupTextBox.Size = New-Object System.Drawing.Size(175, 20)
$RunLookupTextBox.Text = ''

# If text box is empty, disable the Lookup button
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

************************************************************************************
<# ? END OF NSLOOKUP #>
************************************************************************************

************************************************************************************
<# ? START OF SERVER PING #>
************************************************************************************

# Test computer connection button
$TestConnButton = New-Object System.Windows.Forms.Button
$TestConnButton.Location = New-Object System.Drawing.Point(160, 400)
$TestConnButton.Width = 45
$TestConnButton.BackColor = "White"
$TestConnButton.ForeColor = "#0060a9"
$TestConnButton.FlatStyle = "Popup"
$TestConnButton.Text = "Test"
$TestConnButton.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
$TestConnButton.Enabled = $false

# Enter computer text box
$EnterCompTextBox = New-Object System.Windows.Forms.TextBox
$EnterCompTextBox.Location = New-Object System.Drawing.Size(5, 400)
$EnterCompTextBox.Size = New-Object System.Drawing.Size(150, 20)
$EnterCompTextBox.Text = ''
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

************************************************************************************
<# ? END OF SERVER PING #>
************************************************************************************

$SysAdminTab.Controls.Add($RestartsTabControl)
$SysAdminTab.Controls.Add($ServersListBox)
$SysAdminTab.Controls.Add($AppListCombo)
$SysAdminTab.Controls.Add($RestartButton)
$SysAdminTab.Controls.Add($StartButton)
$SysAdminTab.Controls.Add($StopButton)
$SysAdminTab.Controls.Add($TestConnButton)
$SysAdminTab.Controls.Add($EnterCompTextBox)
$SysAdminTab.Controls.Add($RunLookupButton)
$SysAdminTab.Controls.Add($RunLookupTextBox)
$RestartsTabControl.Controls.Add($ServersListBox)
$RestartsTabControl.Controls.Add($AppPoolsListBox)
$ServicesTab.Controls.Add($ServicesListBox)
$AppPoolsTab.Controls.Add($AppPoolsListBox)

# Build form
$Form.Controls.Add($MainFormTabControl)
$Form.Controls.Add($bannerPanel)
$bannerPanel.Controls.Add($bannerLabel)
$MainFormTabControl.Controls.Add($SysAdminTab)
$MainFormTabControl.Controls.Add($SupportTab)
$MainFormTabControl.Controls.Add($TicketManagerTab)
$Form.Controls.Add($OutText)

# Show Form
$Form.ShowDialog() | Out-Null

# Clean up sync hash when form is closed
$Form.Add_FormClosing({
    $synchash.Closed = $True   
})

})

$MainGUI.Runspace = $newRunspace
$async = $MainGUI.BeginInvoke()
$MainGUI.EndInvoke($async)