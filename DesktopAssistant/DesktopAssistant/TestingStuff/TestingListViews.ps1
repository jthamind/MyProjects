Add-Type -AssemblyName System.Data
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationCore
[System.Windows.Forms.Application]::EnableVisualStyles()

# Initialize form
$Form = New-Object System.Windows.Forms.Form
$Form.Text = "Mr. Brobot"
$Form.Size = New-Object System.Drawing.Size(900, 600)

# Create a ListView control
$ServerList = New-Object Windows.Forms.ListView
$ServerList.Size = New-Object Drawing.Size(350, 250)
$ServerList.Location = New-Object Drawing.Point(10, 10)
$ServerList.View = [System.Windows.Forms.View]::Details

# Create ListView columns and add headers
$ServerList.Columns.Add("Group", 100)
$ServerList.Columns.Add("Server Name", 150)

# Create a function to add groups to the ListView
function Add-ListViewGroup {
    param (
        [System.Windows.Forms.ListView]$ListView,
        [string]$GroupName
    )

    $group = New-Object System.Windows.Forms.ListViewGroup($GroupName)
    $ListView.Groups.Add($group)
}

# Create ListView groups and add items to groups
Add-ListViewGroup -ListView $ServerList -GroupName "Group 1"
Add-ListViewGroup -ListView $ServerList -GroupName "Group 2"

# Add items for group headers
$Group1Header = New-Object Windows.Forms.ListViewItem
$Group1Header.Text = "Group 1"
$Group1Header.BackColor = [System.Drawing.Color]::Gray
$ServerList.Items.Add($Group1Header)

$Group2Header = New-Object Windows.Forms.ListViewItem
$Group2Header.Text = "Group 2"
$Group2Header.BackColor = [System.Drawing.Color]::Gray
$ServerList.Items.Add($Group2Header)

$ListItem1 = New-Object Windows.Forms.ListViewItem("Server1")
$ListItem2 = New-Object Windows.Forms.ListViewItem("Server2")
$ListItem3 = New-Object Windows.Forms.ListViewItem("Server3")
$ListItem4 = New-Object Windows.Forms.ListViewItem("Server4")

# Assign items to groups
$ListItem1.Group = $ServerList.Groups["Group 1"]
$ListItem2.Group = $ServerList.Groups["Group 1"]
$ListItem3.Group = $ServerList.Groups["Group 2"]
$ListItem4.Group = $ServerList.Groups["Group 2"]

# Add the items to the ListView
$ServerList.Items.Add($ListItem1)
$ServerList.Items.Add($ListItem2)
$ServerList.Items.Add($ListItem3)
$ServerList.Items.Add($ListItem4)

# Build Form
$Form.Controls.Add($ServerList)

# Show Form
$Form.ShowDialog() | Out-Null
