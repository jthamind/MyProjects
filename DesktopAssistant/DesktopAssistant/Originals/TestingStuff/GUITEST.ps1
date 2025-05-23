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

# Restarts tab control creation
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

# Label applist combo box
$AppListLabel = New-Object System.Windows.Forms.Label
$AppListLabel.Location = New-Object System.Drawing.Size(5, 40)
$AppListLabel.Size = New-Object System.Drawing.Size(150, 20)
$AppListLabel.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Bold)
$AppListLabel.ForeColor = "White"
$AppListLabel.Text = 'Select a Server:'

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

# Function to populate the list boxes based on the selected server
function PopulateListBox {
    $SelectedTab = $RestartsTabControl.SelectedTab.Text
    $synchash.SelectedTab = $SelectedTab

    # Get the selected server from the $ServersListBox
    $SelectedServer = $ServersListBox.SelectedItem

    if ($null -eq $SelectedServer) {
        return
    }

    switch ($SelectedTab) {
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
                $OutText.AppendText("$(Get-Timestamp) - Error retrieving service: $($_.Exception.Message)`r`n")
            }
        }
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
                $OutText.AppendText("$(Get-Timestamp) - Error retrieving AppPools: $($_.Exception.Message)`r`n")
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
    $SelectedServer = $ServersListBox.SelectedItem
    $SelectedTab = $RestartsTabControl.SelectedTab.Text
    $synchash.SelectedTab = $SelectedTab
    if ($SelectedTab -eq "Services") {
        foreach ($item in $ServicesListBox.SelectedItems) {
            try {
                $OutText.AppendText("$(Get-Timestamp) - Restarting $item on $SelectedServer `r`n")
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
                $OutText.AppendText("$(Get-Timestamp) - Restarting $item on $SelectedServer `r`n")
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
    $SelectedServer = $ServersListBox.SelectedItem
    $SelectedTab = $RestartsTabControl.SelectedTab.Text
    $synchash.SelectedTab = $SelectedTab
    if ($SelectedTab -eq "Services") {
        foreach ($item in $ServicesListBox.SelectedItems) {
            try {
                $OutText.AppendText("$(Get-Timestamp) - Starting $item on $SelectedServer `r`n")
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
                $OutText.AppendText("$(Get-Timestamp) - Starting $item on $SelectedServer `r`n")
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
    $SelectedServer = $ServersListBox.SelectedItem
    $SelectedTab = $RestartsTabControl.SelectedTab.Text
    $synchash.SelectedTab = $SelectedTab
    if ($SelectedTab -eq "Services") {
        foreach ($item in $ServicesListBox.SelectedItems) {
            try {
                $OutText.AppendText("$(Get-Timestamp) - Stopping $item on $SelectedServer `r`n")
                Invoke-Command -ComputerName $SelectedServer -ScriptBlock {
                    $service = Get-Service -Name $using:item
                    Stop-Service -Name $using:item -NoWait

                    $timeout = 10 # seconds
                    $sw = [System.Diagnostics.Stopwatch]::StartNew()
                    while ($service.Status -ne 'Stopped' -and $sw.Elapsed.TotalSeconds -lt $timeout) {
                        Start-Sleep -Seconds 1
                        $service.Refresh()
                    }

                    if ($service.Status -ne 'Stopped') {
                        # If service did not stop, force kill the process
                        $process = Get-Process -Id $service.ProcessId -ErrorAction SilentlyContinue
                        if ($process) {
                            $process.Kill()
                        }
                    }
                }
                $OutText.AppendText("$(Get-Timestamp) - Stopped $item`r`n")
            } catch {
                $OutText.AppendText("$(Get-Timestamp) - Error stopping service: $($_.Exception.Message)`r`n")
            }
        }
    } elseif ($SelectedTab -eq "App Pools") {
        foreach ($item in $AppPoolsListBox.SelectedItems) {
            try {
                $OutText.AppendText("$(Get-Timestamp) - Stopping $item on $SelectedServer `r`n")
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

# Text box for Resolve-DnsName
$RunLookupTextBox = New-Object System.Windows.Forms.TextBox
$RunLookupTextBox.Location = New-Object System.Drawing.Size(225, 400)
$RunLookupTextBox.Size = New-Object System.Drawing.Size(175, 20)
$RunLookupTextBox.Text = ''

# Label for Enter computer text box
$RunLookupLabel = New-Object System.Windows.Forms.Label
$RunLookupLabel.Location = New-Object System.Drawing.Size(225, 375)
$RunLookupLabel.Size = New-Object System.Drawing.Size(150, 20)
$RunLookupLabel.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Bold)
$RunLookupLabel.ForeColor = "White"
$RunLookupLabel.Text = 'Run nslookup'

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
        $_.SuppressKeyPress = $true
    }
})

************************************************************************************
<# ? END OF NSLOOKUP #>
************************************************************************************

************************************************************************************
<# ? START OF SERVER PING #>
************************************************************************************

# Button for testing server connection
$ServerPingButton = New-Object System.Windows.Forms.Button
$ServerPingButton.Location = New-Object System.Drawing.Point(160, 400)
$ServerPingButton.Width = 45
$ServerPingButton.BackColor = "White"
$ServerPingButton.ForeColor = "#0060a9"
$ServerPingButton.FlatStyle = "Popup"
$ServerPingButton.Text = "Test"
$ServerPingButton.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
$ServerPingButton.Enabled = $false

# Text box for testing server connection
$ServerPingTextBox = New-Object System.Windows.Forms.TextBox
$ServerPingTextBox.Location = New-Object System.Drawing.Size(5, 400)
$ServerPingTextBox.Size = New-Object System.Drawing.Size(150, 20)
$ServerPingTextBox.Text = ''
$ServerPingTextBox.ShortcutsEnabled = $True

# Label for testing server connection text box
$ServerPingLabel = New-Object System.Windows.Forms.Label
$ServerPingLabel.Location = New-Object System.Drawing.Size(5, 375)
$ServerPingLabel.Size = New-Object System.Drawing.Size(150, 20)
$ServerPingLabel.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Bold)
$ServerPingLabel.ForeColor = "White"
$ServerPingLabel.Text = 'Test Server Connection'

# Computer text box logic
# If text box is empty, disable the Test button
$ServerPingTextBox.Add_TextChanged({
    if ($ServerPingTextBox.Text -ne '') {
        $ServerPingButton.Enabled = $True 
    }
    else {
        $ServerPingButton.Enabled = $False
    }
})

# Function to test computer connection
function Test-ComputerConnection {
    param($ComputerName)
    $OutText.AppendText("$(Get-Timestamp) - Testing connection to $ComputerName...`r`n")
    $PingResult = Test-Connection -ComputerName $ComputerName
    if ($PingResult -ne $null) {
        $OutText.AppendText("$(Get-Timestamp) - Connection to $ComputerName successful.`r`n")
    }
    else {
        $OutText.AppendText("$(Get-Timestamp) - Connection to $ComputerName failed.`r`n")
    }
}

# Test connection button logic
$ServerPingButton.Add_Click({
    $ComputerName = $ServerPingTextBox.Text
    Test-ComputerConnection -ComputerName $ComputerName
})

# Computer text box Enter key logic
$ServerPingTextBox.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") {
        $ComputerName = $ServerPingTextBox.Text
        Test-ComputerConnection -ComputerName $ComputerName
        $_.SuppressKeyPress = $true
    }
})

************************************************************************************
<# ? END OF SERVER PING #>
************************************************************************************

************************************************************************************
<# ? START OF TICKET MANAGER #>
************************************************************************************

# Global variables for ticket manager
$TicketsPath = "$env:USERPROFILE\OneDrive - Allied Solutions\Documents\Allied\Tickets"

# Button for new ticket
$NewTicketButton = New-Object System.Windows.Forms.Button
$NewTicketButton.Location = New-Object System.Drawing.Point(170, 40)
$NewTicketButton.Width = 75
$NewTicketButton.BackColor = "White"
$NewTicketButton.ForeColor = "#0060a9"
$NewTicketButton.FlatStyle = "Popup"
$NewTicketButton.Text = "New Ticket"
$NewTicketButton.Enabled = $false

# Text box for new ticket
$NewTicketTextBox = New-Object System.Windows.Forms.TextBox
$NewTicketTextBox.Location = New-Object System.Drawing.Point(25, 40)
$NewTicketTextBox.Size = New-Object System.Drawing.Size(140, 30)
$NewTicketTextBox.Text = ''
$NewTicketTextBox.ShortcutsEnabled = $True

# Label applist combo box
$NewTicketLabel = New-Object System.Windows.Forms.Label
$NewTicketLabel.Location = New-Object System.Drawing.Size(25, 15)
$NewTicketLabel.Size = New-Object System.Drawing.Size(150, 20)
$NewTicketLabel.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Bold)
$NewTicketLabel.ForeColor = "White"
$NewTicketLabel.Text = 'Create a New Ticket'

# Button for renaming a ticket
$RenameTicketButton = New-Object System.Windows.Forms.Button
$RenameTicketButton.Location = New-Object System.Drawing.Point(455, 40)
$RenameTicketButton.Width = 75
$RenameTicketButton.BackColor = "White"
$RenameTicketButton.ForeColor = "#0060a9"
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
$RenameTicketLabel.ForeColor = "White"
$RenameTicketLabel.Text = 'Rename a Ticket'

# Tab control for ticket manager
$TicketManagerTabControl = New-Object System.Windows.Forms.TabControl
$TicketManagerTabControl.Location = "25,95"
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

# Complete ticket button
$CompleteTicketButton = New-Object System.Windows.Forms.Button
$CompleteTicketButton.Location = New-Object System.Drawing.Point(25, 360)
$CompleteTicketButton.Width = 75
$CompleteTicketButton.BackColor = "White"
$CompleteTicketButton.ForeColor = "#0060a9"
$CompleteTicketButton.FlatStyle = "Popup"
$CompleteTicketButton.Text = "Complete"
$CompleteTicketButton.Enabled = $false
$CompleteTicketButton.Visible = $true

# Reactivate ticket button
$ReactivateTicketButton = New-Object System.Windows.Forms.Button
$ReactivateTicketButton.Location = New-Object System.Drawing.Point(25, 360)
$ReactivateTicketButton.Width = 75
$ReactivateTicketButton.BackColor = "White"
$ReactivateTicketButton.ForeColor = "#0060a9"
$ReactivateTicketButton.FlatStyle = "Popup"
$ReactivateTicketButton.Text = "Reactivate"
$ReactivateTicketButton.Enabled = $false
$ReactivateTicketButton.Visible = $false

# Open folder button
$OpenFolderButton = New-Object System.Windows.Forms.Button
$OpenFolderButton.Location = New-Object System.Drawing.Point(170, 360)
$OpenFolderButton.Width = 75
$OpenFolderButton.BackColor = "White"
$OpenFolderButton.ForeColor = "#0060a9"
$OpenFolderButton.FlatStyle = "Popup"
$OpenFolderButton.Text = "Open"
$OpenFolderButton.Enabled = $false

# List box for active tickets
$ActiveTicketsListBox = New-Object System.Windows.Forms.ListBox
$ActiveTicketsListBox.Location = New-Object System.Drawing.Point(0,0)
$ActiveTicketsListBox.Size = New-Object System.Drawing.Size(215,240)
$ActiveTicketsListBox.DockStyle = "Fill"
$ActiveTicketsListBox.SelectionMode = 'MultiExtended'
$ActiveTicketsListBox.ScrollBars = 'Vertical'

# List box for completed tickets
$CompletedTicketsListBox = New-Object System.Windows.Forms.ListBox
$CompletedTicketsListBox.Location = New-Object System.Drawing.Point(0,0)
$CompletedTicketsListBox.Size = New-Object System.Drawing.Size(215,240)
$CompletedTicketsListBox.DockStyle = "Fill"
$CompletedTicketsListBox.SelectionMode = 'MultiExtended'
$CompletedTicketsListBox.ScrollBars = 'Vertical'

# Creates the active and completed tickets path variables on startup
# Calls function immediately after
function Start-Setup {
    $ActiveTicketsPath = "$TicketsPath\Active"
    $CompletedTicketsPath = "$TicketsPath\Completed"
    if (!(Test-Path -Path $ActiveTicketsPath)) {
    (mkdir $ActiveTicketsPath)
    }   
    if (!(Test-Path -Path $CompletedTicketsPath)) {
    
        (mkdir $CompletedTicketsPath)
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
    if ($NewTicketTextBox.Text -ne '') {
        $NewTicketButton.Enabled = $True
    }
    else {
        $NewTicketButton.Enabled = $False
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
    Invoke-Item "$TicketsPath\Active\$TicketNumber"
    Get-ActiveListItems($ActiveTicketsListBox)
    Get-CompletedListItems($CompletedTicketsListBox)
    Set-RenameOff
}

# Logic to enable rename ticket button when text box is not empty
$RenameTicketTextBox.Add_TextChanged({
    if ($RenameTicketTextBox.Text -ne '') {
        $RenameTicketButton.Enabled = $True
    }
    else {
        $RenameTicketButton.Enabled = $False
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
        $RenameTicketTextBox.Text = $ActiveTicketsListBox.SelectedItem
    } else {
        $CompleteTicketButton.Enabled = $false
        $ReactivateTicketButton.Enabled = $false
        $RenameTicketTextBox.Enabled = $false
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
        $RenameTicketTextBox.Text = $CompletedTicketsListBox.SelectedItem
    } else {
        $CompleteTicketButton.Enabled = $false
        $ReactivateTicketButton.Enabled = $false
        $RenameTicketTextBox.Enabled = $false
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
}

# Event handler for new ticket button
$NewTicketButton.Add_Click({
    New-Ticket
})

# Event handler for rename ticket button
$RenameTicketButton.Add_Click({
    $ticket = $ActiveTicketsListBox.SelectedItem
    $NewName  = $RenameTicketTextBox.Text
    if ($ticket -ne $NewName) {
        Rename-Item -Path "$TicketsPath\Active\$ticket" -NewName $NewName
        Set-RenameOff
        Get-ActiveListItems($ActiveTicketsListBox)
        Get-CompletedListItems($CompletedTicketsListBox)
    }
    else {
        $OutText.AppendText("$(Get-Timestamp) - Please enter a new ticket name`r`n")}
})

# Event handler for complete ticket button
$CompleteTicketButton.Add_Click({
    $tickets = $ActiveTicketsListBox.SelectedItems
    foreach ($ticket in $tickets){
        Move-Item -Path "$TicketsPath\Active\$ticket" -Destination "$TicketsPath\Completed\"
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
        }
    } else {
        $tickets = $CompletedTicketsListBox.SelectedItems
        foreach ($ticket in $tickets){
            Invoke-Item "$TicketsPath\Completed\$ticket"
        }
    }
})

# Event handler for reactivate ticket button
$ReactivateTicketButton.Add_Click({
    $tickets = $CompletedTicketsListBox.SelectedItems
    foreach ($ticket in $tickets){
        Move-Item -Path "$TicketsPath\Completed\$ticket" -Destination "$TicketsPath\Active\"
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

$NewTicketTextBox.Add_KeyDown({
    if ($_.KeyCode -eq "Enter" -and $NewTicketTextBox.Text.Trim() -ne "") {
        New-Ticket
        $_.SuppressKeyPress = $true
    }
    else {
        $OutText.AppendText("$(Get-Timestamp) - Please enter a ticket number`r`n")
    }
})

# Event handler for Enter key logic on Rename Ticket text box
$RenameTicketTextBox.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") {
        $ticket = $ActiveTicketsListBox.SelectedItem
        $NewName  = $RenameTicketTextBox.Text
        Rename-Item -Path "$TicketsPath\Active\$ticket" -NewName $NewName
        Set-RenameOff
        Get-ActiveListItems($ActiveTicketsListBox)
        Get-CompletedListItems($CompletedTicketsListBox)
        $_.SuppressKeyPress = $true
    }
})

# Register the Ticket Manager event handlers
$ActiveTicketsListBox.add_SelectedIndexChanged($ActiveTicketsListBox_SelectedIndexChanged)
$CompletedTicketsListBox.add_SelectedIndexChanged($CompletedTicketsListBox_SelectedIndexChanged)
$TicketManagerTabControl.add_SelectedIndexChanged($TicketManagerTabControl_SelectedIndexChanged)

************************************************************************************
<# ? END OF TICKET MANAGER #>
************************************************************************************

# Build form
$Form.Controls.Add($MainFormTabControl)
$Form.Controls.Add($bannerPanel)
$bannerPanel.Controls.Add($bannerLabel)
$MainFormTabControl.Controls.Add($SysAdminTab)
$MainFormTabControl.Controls.Add($SupportTab)
$MainFormTabControl.Controls.Add($TicketManagerTab)
$Form.Controls.Add($OutText)
$SysAdminTab.Controls.Add($RestartsTabControl)
$SysAdminTab.Controls.Add($ServersListBox)
$SysAdminTab.Controls.Add($AppListCombo)
$SysAdminTab.Controls.Add($RestartButton)
$SysAdminTab.Controls.Add($StartButton)
$SysAdminTab.Controls.Add($StopButton)
$SysAdminTab.Controls.Add($ServerPingButton)
$SysAdminTab.Controls.Add($ServerPingTextBox)
$SysAdminTab.Controls.Add($RunLookupButton)
$SysAdminTab.Controls.Add($RunLookupTextBox)
$SysAdminTab.Controls.Add($ServerPingLabel)
$SysAdminTab.Controls.Add($RunLookupLabel)
$SysAdminTab.Controls.Add($AppListLabel)
$RestartsTabControl.Controls.Add($ServersListBox)
$RestartsTabControl.Controls.Add($AppPoolsListBox)
$ServicesTab.Controls.Add($ServicesListBox)
$AppPoolsTab.Controls.Add($AppPoolsListBox)
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