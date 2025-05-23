Add-Type -AssemblyName System.Data
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationCore
<# Global Variables #>
$ParentPath = "$env:USERPROFILE\OneDrive - Allied Solutions\Documents\Allied\Tickets"
<# Functions #>
function Start-Setup {
$ParentPath = "$env:USERPROFILE\OneDrive - Allied Solutions\Documents\"
$AlliedTicketPath = "$ParentPath\Allied"
$activeTicketsPath = "$AlliedTicketPath\Tickets\Active"
$CompletedTicketsPath = "$AlliedTicketPath\Tickets\Completed"
if (!(Test-Path -Path $activeTicketsPath)) {
    mkdir $activeTicketsPath | Out-Null
}   
if (!(Test-Path -Path $activeTicketsPath)) {

    mkdir $CompletedTicketsPath | Out-Null
}
}
Start-Setup
Function Get-ActiveListItems {
    param($listbox)
    if ($listbox.items.Count -gt 0){
        $listBox.Items.Clear()
    }
$tickets = @(Get-ChildItem "$ParentPath\Active" | Select-Object -ExpandProperty Name)
foreach ($ticket in $tickets) {
    [void]$activeTicketsLstBx.Items.Add($ticket)
}
}
function Get-CompletedListItems {
    param($listbox)
    if ($listbox.items.Count -gt 0){
        $listBox.Items.Clear()
    }
    $tickets = @(Get-ChildItem "$ParentPath\Completed" | Select-Object -ExpandProperty Name)
    foreach ($ticket in $tickets) {
        [void]$CompleteTicketsLstBx.Items.Add($ticket)
    }
}
<# End Functions  #>

<# Initialize Form #>
$Form = New-Object System.Windows.Forms.Form
$Form.Text = "Ticket Manager"
$Form.Size = New-Object System.Drawing.Size(600, 300)
$Form.ShowInTaskbar = $True
$Form.KeyPreview = $True
$Form.AutoSize = $True
$Form.FormBorderStyle = 'Fixed3D'
$Form.MaximizeBox = $False
$Form.MinimizeBox = $True
$Form.ControlBox = $True
$Form.Icon = $Icon

<# * New Ticket Button #>
$newTicketBtn = New-Object System.Windows.Forms.Button
$newTicketBtn.Location = New-Object System.Drawing.Size(125, 25)
$newTicketBtn.Text = "New Ticket"
$newTicketBtn.Enabled = $False

<# * New Ticket Text Box #>
$newTicketTxtBx = New-Object System.Windows.Forms.TextBox
$newTicketTxtBx.Location = New-Object System.Drawing.Size(10, 25)
$newTicketTxtBx.Text = ''

<# * Rename Ticket Text Box #>
$RenameTicketTxtBx = New-Object System.Windows.Forms.TextBox
$RenameTicketTxtBx.Location = New-Object System.Drawing.Size(250, 25)
$RenameTicketTxtBx.Text = ''
$RenameTicketTxtBx.Enabled = $False

<# * Rename Ticket Button #>
$RenameTicketBtn = New-Object System.Windows.Forms.Button
$RenameTicketBtn.Location = New-Object System.Drawing.Size(365, 25)
$RenameTicketBtn.Text = "Rename"
$RenameTicketBtn.Enabled = $False

<#** Tab control Creation  **#>
$FormTabControl = New-object System.Windows.Forms.TabControl
$FormTabControl.Size = "600,300"
$FormTabControl.Location = "0,75"

<#**  List box to show folder contents  #>
$FolderContentsLstBx = New-Object System.Windows.Forms.ListBox
$FolderContentsLstBx.Location = New-Object System.Drawing.Point(300, 125)
$FolderContentsLstBx.Size = New-Object System.Drawing.Size(225, 225)
$FolderContentsLstBx.SelectionMode = 'MultiExtended'
$Form.Controls.Add($FolderContentsLstBx)

<# Active Tickets Tab #>
$ActiveTab = New-object System.Windows.Forms.Tabpage
$ActiveTab.DataBindings.DefaultDataSourceUpdateMode = 0
$ActiveTab.UseVisualStyleBackColor = $True
$ActiveTab.Name = 'ActiveTickets'
$ActiveTab.Text = 'Active Tickets'
$ActiveTab.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
$FormtabControl.Controls.add($ActiveTab)  

<# Active Items Listbox #>
$activeTicketsLstBx = New-Object System.Windows.Forms.ListBox
$activeTicketsLstBx.Location = New-Object System.Drawing.Point(25, 25)
$activeTicketsLstBx.Size = New-Object System.Drawing.Size(225, 225)
$activeTicketsLstBx.SelectionMode = 'MultiExtended'
$ActiveTab.Controls.Add($activeTicketsLstBx)
Get-ActiveListItems($activeTicketsLstBx)

<# Complete Ticket Button #>
$CompleteTicketBtn = New-Object System.Windows.Forms.Button
$CompleteTicketBtn.Location = New-Object System.Drawing.Size(25, 245)
$completeTicketBtn.Text = "Complete"
$completeTicketBtn.Width = 100
$CompleteTicketBtn.Enabled = $False
$ActiveTab.Controls.Add($CompleteTicketBtn)

<# Open folder Ticket Button #>
$OpenFolderBtn = New-Object System.Windows.Forms.Button
$OpenFolderBtn.Location = New-Object System.Drawing.Size(150, 245)
$OpenFolderBtn.Width = 100
$OpenFolderBtn.Text = "Open Folder"
$OpenFolderBtn.Enabled = $False
$ActiveTab.Controls.Add($OpenFolderBtn)

<# Completed Tickets Tab #>
$CompletedTab = New-object System.Windows.Forms.Tabpage
$CompletedTab.DataBindings.DefaultDataSourceUpdateMode = 0
$CompletedTab.UseVisualStyleBackColor = $True
$CompletedTab.Name = "CompletedTickets"
$CompletedTab.Text = "Completed Tickets"
$CompletedTab.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
$FormtabControl.Controls.add($CompletedTab)  

<# Completed Items Listbox #>
$CompleteTicketsLstBx = New-Object System.Windows.Forms.ListBox
$CompleteTicketsLstBx.Location = New-Object System.Drawing.Point(25, 25)
$CompleteTicketsLstBx.Size = New-Object System.Drawing.Size(225, 225)
$CompleteTicketsLstBx.SelectionMode = 'MultiExtended'
$CompletedTab.Controls.Add($CompleteTicketsLstBx)
Get-CompletedListItems($CompleteTicketsLstBx)

<# Reactivate Ticket Button #>
$ReactivateTicketBtn = New-Object System.Windows.Forms.Button
$ReactivateTicketBtn.Location = New-Object System.Drawing.Size(25, 245)
$ReactivateTicketBtn.Text = "Reactivate"
$ReactivateTicketBtn.Enabled = $False
$CompletedTab.Controls.Add($ReactivateTicketBtn)

<# ************END FORM BUILD Section *************** #>

<# Validation Logic #>
$newTicketTxtBx.Add_TextChanged({
        if ($newTicketTxtBx.Text -ne '') {
            $newTicketBtn.Enabled = $True 
        }
        else {
            $newTicketBtn.Enabled = $False
        }
    })

$activeTicketsLstBx.Add_SelectedIndexChanged({
 
        If ($activeTicketsLstBx.SelectedItems.Count -gt 0) {
            $CompleteTicketBtn.Enabled = $True
        }
        else {
            $CompleteTicketBtn.Enabled = $false
        }
        If ($activeTicketsLstBx.SelectedItems.Count -eq 1) {
            $RenameTicketBtn.Enabled = $True
            $RenameTicketTxtBx.Enabled = $True
            $RenameTicketTxtBx.Text = $activeTicketsLstBx.SelectedItem
        }
        else {
            $RenameTicketBtn.Enabled = $false
            $RenameTicketTxtBx.Text = ''
            $RenameTicketTxtBx.Enabled = $false
        }
        If ($activeTicketsLstBx.SelectedItems.Count -eq 1) {
            $OpenFolderBtn.Enabled = $True
        }
        else {
            $OpenFolderBtn.Enabled = $false
        }

    })
$CompleteTicketsLstBx.Add_SelectedIndexChanged({
        If ($CompleteTicketsLstBx.SelectedItems -gt 0) {
            $ReactivateTicketBtn.Enabled = $True
        }
        else {
            $ReactivateTicketBtn.Enabled = $false
        }
    })

<# Button Logic #>
$newTicketBtn.Add_Click({
        $TicketNumber = $newTicketTxtBx.Text
        mkdir "$ParentPath\Active\$TicketNumber"
        if ($TicketNumber -like "DS*" ) {
            $DSTicket = $TicketNumber.Substring(3, 5)
            $UrlPath = "http://tfs-sharepoint.alliedsolutions.net/Sites/Unitrac/Lists/Database%20Scripting/DispForm.aspx?ID=$DSTicket"
        }
        else {
            $UrlPath = "https://alliedsolutions.atlassian.net/browse/$TicketNumber"
        }
        $wshShell = New-Object -ComObject "WScript.Shell"
        $urlShortcut = $wshShell.CreateShortcut(
            "$ParentPath\Active\$TicketNumber\$TicketNumber.url")
        $urlShortcut.TargetPath = $UrlPath
        $urlShortcut.Save()
        $newTicketTxtBx.Text = ''
        Invoke-Item "$ParentPath\Active\$TicketNumber"
        Get-ActiveListItems($ActiveTicketsLstBx)
        Get-CompletedListItems($CompleteTicketsLstBx)

    })

# Event handler for displaying folder contents of selected active ticket
$activeTicketsLstBx.Add_Click({
    $ticket = $ActiveTicketsLstBx.SelectedItem
    $FolderContentsLstBx.Items.Clear()
    $ticketPath = Join-Path $ParentPath "Active\$ticket"
    $ticketFiles = Get-ChildItem -Path $ticketPath
    foreach ($file in $ticketFiles) {
        [void]$FolderContentsLstBx.Items.Add($file.Name)
    }
})

# Event handler for displaying folder contents of selected completed ticket
$CompleteTicketsLstBx.Add_Click({
    $ticket = $CompleteTicketsLstBx.SelectedItem
    $FolderContentsLstBx.Items.Clear()
    $ticketPath = Join-Path $ParentPath "Completed\$ticket"
    $ticketFiles = Get-ChildItem -Path $ticketPath
    foreach ($file in $ticketFiles) {
        [void]$FolderContentsLstBx.Items.Add($file.Name)
    }
})

$CompleteTicketBtn.Add_Click({
    $tickets = $activeTicketsLstBx.SelectedItems
    foreach ($ticket in $tickets){
        Move-Item -Path "$ParentPath\Active\$ticket" -Destination "$ParentPath\Completed\"
    }
    $RenameTicketBtn.Enabled = $false
    $RenameTicketTxtBx.Text = ''
    $RenameTicketTxtBx.Enabled = $false
    Get-ActiveListItems($ActiveTicketsLstBx)
    Get-CompletedListItems($CompleteTicketsLstBx)
})

$RenameTicketBtn.Add_Click({
    $ticket = $ActiveTicketsLstBx.SelectedItem
    $newName  = $RenameTicketTxtBx.Text
    Rename-Item -Path "$ParentPath\Active\$ticket" -NewName $NewName
    $RenameTicketBtn.Enabled = $false
    $RenameTicketTxtBx.Text = ''
    $RenameTicketTxtBx.Enabled = $false
    Get-ActiveListItems($ActiveTicketsLstBx)
    Get-CompletedListItems($CompleteTicketsLstBx)
})

$OpenFolderBtn.Add_Click({
    $ticket = $activeTicketsLstBx.SelectedItem
    Invoke-Item "$ParentPath\Active\$ticket"
    $RenameTicketBtn.Enabled = $false
    $RenameTicketTxtBx.Text = ''
    $RenameTicketTxtBx.Enabled = $false
})

$ReactivateTicketBtn.Add_Click({
    $tickets = $CompleteTicketsLstBx.SelectedItems
    foreach ($ticket in $tickets){
        Move-Item -Path "$ParentPath\Completed\$ticket" -Destination "$ParentPath\Active\"
    }
    $RenameTicketBtn.Enabled = $false
    $RenameTicketTxtBx.Text = ''
    $RenameTicketTxtBx.Enabled = $false
    Get-ActiveListItems($ActiveTicketsLstBx)
    Get-CompletedListItems($CompleteTicketsLstBx)
})

$ActiveTab.Add_Enter({
    $FolderContentsLstBx.Items.Clear()
    $activeTicketsLstBx.SelectedItems.Clear()
})

$CompletedTab.Add_Enter({
    $FolderContentsLstBx.Items.Clear()
    $CompleteTicketsLstBx.SelectedItems.Clear()
})


<# Build Form #>
$Form.Controls.Add($newTicketTxtBx)
$Form.Controls.Add($newTicketBtn)
$Form.Controls.Add($RenameTicketTxtBx)
$Form.Controls.Add($RenameTicketBtn)
$Form.Controls.Add($FormTabControl)
<# * Shows form  #>
$Form.ShowDialog() | Out-Null