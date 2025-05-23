Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public static class Win32Helpers {
    [DllImport("gdi32.dll")]
    public static extern IntPtr CreateRoundRectRgn(int nLeftRect, int nTopRect,
        int nRightRect, int nBottomRect, int nWidthEllipse, int nHeightEllipse);
}
"@
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
$newRunspace.Name = "ServiceRestarts"
$newRunspace.Open()
$newRunspace.SessionStateProxy.SetVariable("synchash", $synchash)

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
$Form.Size = New-Object System.Drawing.Size(900, 600)
$Form.ShowInTaskbar = $True
$Form.KeyPreview = $True
$Form.AutoSize = $True
$Form.FormBorderStyle = "Fixed3D"
$Form.MaximizeBox = $False
$Form.MinimizeBox = $True
$Form.ControlBox = $True
$Form.Icon = $Icon
$Form.TopMost = $True
$Form.StartPosition = "CenterScreen"

# Tab control creation
$FormTabControl = New-object System.Windows.Forms.TabControl
$FormTabControl.Size = "200,250"
$FormTabControl.Location = "0,150"

# Tab for services list
$ServicesTab = New-Object System.Windows.Forms.TabPage
$ServicesTab.DataBindings.DefaultDataSourceUpdateMode = 0
$ServicesTab.UseVisualStyleBackColor = $True
$ServicesTab.Name = 'ServicesTab'
$ServicesTab.Text = 'Services'
$ServicesTab.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)

<# ! IIS Sites stuff - NOT IN USE
# Tab for IIS Sites list
$IISSitesTab = New-Object System.Windows.Forms.TabPage
$IISSitesTab.DataBindings.DefaultDataSourceUpdateMode = 0
$IISSitesTab.UseVisualStyleBackColor = $True
$IISSitesTab.Name = 'IISSitesTab'
$IISSitesTab.Text = 'Sites'
$IISSitesTab.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
#>

# Tab for IIS App Pools list
$AppPoolsTab = New-Object System.Windows.Forms.TabPage
$AppPoolsTab.DataBindings.DefaultDataSourceUpdateMode = 0
$AppPoolsTab.UseVisualStyleBackColor = $True
$AppPoolsTab.Name = 'IISAppPools'
$AppPoolsTab.Text = 'App Pools'
$AppPoolsTab.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)

# Combobox for application selection
$AppListCombo = New-Object System.Windows.Forms.ComboBox
$AppListCombo.Location = New-Object System.Drawing.Point(5,30)
$AppListCombo.Size = New-Object System.Drawing.Size(260, 200)
$AppListCombo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$AppListCombo.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)

# Output text box
$OutText = New-Object System.Windows.Forms.TextBox
$OutText.Location = New-Object System.Drawing.Size(600, 180)
$OutText.Size = New-Object System.Drawing.Size(260, 200)
$OutText.Multiline = $true
$OutText.ScrollBars = "Vertical"
$OutText.Enabled = $True
$OutText.ReadOnly = $True
$synchash.OutText = $OutText

# Individual servers list box
$ServersListBox = New-Object System.Windows.Forms.ListBox
$ServersListBox.Location = New-Object System.Drawing.Point(220,180)
$ServersListBox.Size = New-Object System.Drawing.Size(200,225)
$ServersListBox.DockStyle = "Fill"
$ServersListBox.SelectionMode = 'Single'

# Services list box
$ServicesListBox = New-Object System.Windows.Forms.ListBox
$ServicesListBox.Location = New-Object System.Drawing.Point(5,10)
$ServicesListBox.Size = New-Object System.Drawing.Size(180,200)
$ServicesListBox.DockStyle = "Fill"
$ServicesListBox.SelectionMode = 'MultiExtended'

<# ! IIS Sites stuff - NOT IN USE
# IIS Sites list box
$IISSitesListBox = New-Object System.Windows.Forms.ListBox
$IISSitesListBox.Location = New-Object System.Drawing.Point(5,10)
$IISSitesListBox.Size = New-Object System.Drawing.Size(260,200)
$IISSitesListBox.DockStyle = "Fill"
$IISSitesListBox.SelectionMode = 'MultiExtended'
#>

# IIS App Pools list box
$AppPoolsListBox = New-Object System.Windows.Forms.ListBox
$AppPoolsListBox.Location = New-Object System.Drawing.Point(5,10)
$AppPoolsListBox.Size = New-Object System.Drawing.Size(180,200)
$AppPoolsListBox.DockStyle = "Fill"
$AppPoolsListBox.SelectionMode = 'MultiExtended'

# Button for restarting services
$RestartButton = New-Object System.Windows.Forms.Button
$RestartButton.Location = New-Object System.Drawing.Point(500, 180)
$RestartButton.Width = 75
$RestartButton.Text = "Restart"
$RestartButton.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
$RestartButton.Enabled = $false

# Button for starting services
$StartButton = New-Object System.Windows.Forms.Button
$StartButton.Location = New-Object System.Drawing.Point(500, 200)
$StartButton.Width = 75
$StartButton.Text = "Start"
$StartButton.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
$StartButton.Enabled = $false

# Button for stopping services
$StopButton = New-Object System.Windows.Forms.Button
$StopButton.Location = New-Object System.Drawing.Point(500, 220)
$StopButton.Width = 75
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
    $SelectedServer = $synchash.SelectedServer
    if ($synchash.SelectedTab -eq "Services") {
        foreach ($item in $ServicesListBox.SelectedItems) {
            try {
                Invoke-Command -ComputerName $SelectedServer -ScriptBlock {
                    Restart-Service -Name $using:item
                }
                $synchash.OutText.AppendText("$(Get-Timestamp) - Restarted $item`r`n")
            } catch {
                $synchash.OutText.AppendText("$(Get-Timestamp) - Error restarting service: $($_.Exception.Message)`r`n")
            }
        }
    } elseif ($synchash.SelectedTab -eq "App Pools") {
        foreach ($item in $AppPoolsListBox.SelectedItems) {
            try {
                Invoke-Command -ComputerName $SelectedServer -ScriptBlock {
                    Import-Module WebAdministration
                    Restart-WebAppPool -Name $using:item
                }
                $synchash.OutText.AppendText("$(Get-Timestamp) - Restarted $item`r`n")
            } catch {
                $synchash.OutText.AppendText("$(Get-Timestamp) - Error restarting AppPool: $($_.Exception.Message)`r`n")
            }
        }
    }
})

# Button click event handler for starting services
$StartButton.Add_Click({
    $SelectedServer = $synchash.SelectedServer
    if ($synchash.SelectedTab -eq "Services") {
        foreach ($item in $ServicesListBox.SelectedItems) {
            try {
                Invoke-Command -ComputerName $SelectedServer -ScriptBlock {
                    Start-Service -Name $using:item
                }
                $synchash.OutText.AppendText("$(Get-Timestamp) - Started $item`r`n")
            } catch {
                $synchash.OutText.AppendText("$(Get-Timestamp) - Error starting service: $($_.Exception.Message)`r`n")
            }
        }
    } elseif ($synchash.SelectedTab -eq "App Pools") {
        foreach ($item in $AppPoolsListBox.SelectedItems) {
            try {
                Invoke-Command -ComputerName $SelectedServer -ScriptBlock {
                    Import-Module WebAdministration
                    Start-WebAppPool -Name $using:item
                }
                $synchash.OutText.AppendText("$(Get-Timestamp) - Started $item`r`n")
            } catch {
                $synchash.OutText.AppendText("$(Get-Timestamp) - Error starting AppPool: $($_.Exception.Message)`r`n")
            }
        }
    }
})

# Button click event handler for stopping services
$StopButton.Add_Click({
    $SelectedServer = $synchash.SelectedServer
    if ($synchash.SelectedTab -eq "Services") {
        foreach ($item in $ServicesListBox.SelectedItems) {
            try {
                Invoke-Command -ComputerName $SelectedServer -ScriptBlock {
                    Stop-Service -Name $using:item
                }
                $synchash.OutText.AppendText("$(Get-Timestamp) - Stopped $item`r`n")
            } catch {
                $synchash.OutText.AppendText("$(Get-Timestamp) - Error stopping service: $($_.Exception.Message)`r`n")
            }
        }
    } elseif ($synchash.SelectedTab -eq "App Pools") {
        foreach ($item in $AppPoolsListBox.SelectedItems) {
            try {
                Invoke-Command -ComputerName $SelectedServer -ScriptBlock {
                    Import-Module WebAdministration
                    Stop-WebAppPool -Name $using:item
                }
                $synchash.OutText.AppendText("$(Get-Timestamp) - Stopped $item`r`n")
            } catch {
                $synchash.OutText.AppendText("$(Get-Timestamp) - Error stopping AppPool: $($_.Exception.Message)`r`n")
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
        $synchash.SelectedServer = $SelectedServer
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
                $synchash.OutText.AppendText("$(Get-Timestamp) - $SelectedService status: $($ServiceStatus.Status)`r`n")
            } catch {
                $synchash.OutText.AppendText("$(Get-Timestamp) - Error retrieving service status: $($_.Exception.Message)`r`n")
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
                $synchash.OutText.AppendText("$(Get-Timestamp) - $SelectedAppPool status: $($AppPoolStatus.Value)`r`n")
            } catch {
                $synchash.OutText.AppendText("$(Get-Timestamp) - Error retrieving AppPool status: $($_.Exception.Message)`r`n")
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
$FormTabControl_SelectedIndexChanged = {
    PopulateListBox
}

# Function to populate the list boxes on form load
$Form_Load = {
    PopulateListBox
}

# Register the event handlers
$FormTabControl.add_SelectedIndexChanged($FormTabControl_SelectedIndexChanged)
$Form.add_Load($Form_Load)

# Add rounded corners to the form
$Form.add_Load({
    $hrgn = [Win32Helpers]::CreateRoundRectRgn(0, 0, $Form.Width, $Form.Height, 20, 20)
    $Form.Region = [System.Drawing.Region]::FromHrgn($hrgn)
})

# Build form
$Form.Controls.Add($OutText)
$Form.Controls.Add($ServicesListBox)
$Form.Controls.Add($RestartButton)
$Form.Controls.Add($StartButton)
$Form.Controls.Add($StopButton)
$Form.Controls.Add($FormTabControl)
$Form.Controls.Add($AppPoolsListBox)
$FormTabControl.Controls.Add($ServicesTab)
$FormTabControl.Controls.Add($AppPoolsTab)
$Form.Controls.Add($AppListCombo)
$Form.Controls.Add($ServersListBox)
$AppListCombo.add_SelectedIndexChanged($AppListCombo_SelectedIndexChanged)
# ! $Form.Control.Add($IISSitesListBox)
# ! $FormTabControl.Controls.Add($IISSitesTab)

# Add ServicesListBox to each tab
$ServicesTab.Controls.Add($ServicesListBox)
$AppPoolsTab.Controls.Add($AppPoolsListBox)
# ! $IISSitesTab.Controls.Add($IISSitesListBox)

# Show Form
$Form.ShowDialog() | Out-Null

# Clean up sync hash when form is closed
$Form.Add_FormClosing({
    $synchash.Closed = $True   
})

})

$GUIPowershell.Runspace = $newRunspace
$async = $GUIPowershell.BeginInvoke()
$GuiPowerShell.EndInvoke($async)