Add-Type -AssemblyName System.Data
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationCore
[System.Windows.Forms.Application]::EnableVisualStyles()

$LoginUsername = [System.Environment]::UserName
$WorkhorseServer = 'unitrac-wh001'
if ((Get-WSManCredSSP).State -ne "Enabled") {
    Enable-WSManCredSSP -Role Client -DelegateComputer $WorkhorseServer -Force
}

# Initialize form
$Form = New-Object System.Windows.Forms.Form
$Form.Text = "Create HDTStorage Table"
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

# New documentation template button
$CreateTableButton = New-Object System.Windows.Forms.Button
$CreateTableButton.Location = New-Object System.Drawing.Size(125, 25)
$CreateTableButton.Width = 175
$CreateTableButton.Text = "Create HDTStorage Table"
$CreateTableButton.Enabled = $true
$CreateTableButton.BackColor = "Yellow"
$CreateTableButton.ForeColor = "Black"

$CreateTableButton.Add_Click({
    # New popup window
    $HDTStoragePopup = New-Object System.Windows.Forms.Form
    $HDTStoragePopup.Text = "Eyyyy lmao"
    $HDTStoragePopup.Size = New-Object System.Drawing.Size(600, 300)
    $HDTStoragePopup.ShowInTaskbar = $True
    $HDTStoragePopup.KeyPreview = $True
    $HDTStoragePopup.AutoSize = $True
    $HDTStoragePopup.FormBorderStyle = 'Fixed3D'
    $HDTStoragePopup.MaximizeBox = $False
    $HDTStoragePopup.MinimizeBox = $True
    $HDTStoragePopup.ControlBox = $True
    $HDTStoragePopup.BackColor = "Gray"
    $HDTStoragePopup.Icon = $Icon
    $HDTStoragePopup.TopMost = $True
    $HDTStoragePopup.StartPosition = "CenterScreen"

    # Button for selecting the file location
    $HDTStorageFileButton = New-Object System.Windows.Forms.Button
    $HDTStorageFileButton.Location = New-Object System.Drawing.Size(10, 25)
    $HDTStorageFileButton.Width = 75
    $HDTStorageFileButton.Text = "Browse"
    $HDTStorageFileButton.Enabled = $true
    $HDTStorageFileButton.BackColor = "Yellow"
    $HDTStorageFileButton.ForeColor = "Black"

    # Label for the file location button
    $FileLocationLabel = New-Object System.Windows.Forms.Label
    $FileLocationLabel.Location = New-Object System.Drawing.Size(10, 5)
    $FileLocationLabel.Text = "File Location"
    $FileLocationLabel.AutoSize = $True
    $FileLocationLabel.BackColor = "Gray"
    $FileLocationLabel.ForeColor = "Black"

    # Text box for entering the SQL instance
    $DBServerTextBox = New-Object System.Windows.Forms.TextBox
    $DBServerTextBox.Location = New-Object System.Drawing.Size(10, 75)
    $DBServerTextBox.Text = ''
    $DBServerTextBox.BackColor = "White"
    $DBServerTextBox.ForeColor = "Black"
    $DBServerTextBox.ShortcutsEnabled = $True
    $DBServerTextBox.Width = 400

    # Label for the SQL instance text box
    $DBServerLabel = New-Object System.Windows.Forms.Label
    $DBServerLabel.Location = New-Object System.Drawing.Size(10, 55)
    $DBServerLabel.Text = "SQL Instance"
    $DBServerLabel.AutoSize = $True
    $DBServerLabel.BackColor = "Gray"
    $DBServerLabel.ForeColor = "Black"

    # Text box for entering the table name
    $TableNameTextBox = New-Object System.Windows.Forms.TextBox
    $TableNameTextBox.Location = New-Object System.Drawing.Size(10, 125)
    $TableNameTextBox.Text = ''
    $TableNameTextBox.BackColor = "White"
    $TableNameTextBox.ForeColor = "Black"
    $TableNameTextBox.ShortcutsEnabled = $True
    $TableNameTextBox.Width = 400

    # Label for the table name text box
    $TableNameLabel = New-Object System.Windows.Forms.Label
    $TableNameLabel.Location = New-Object System.Drawing.Size(10, 105)
    $TableNameLabel.Text = "Table Name"
    $TableNameLabel.AutoSize = $True
    $TableNameLabel.BackColor = "Gray"
    $TableNameLabel.ForeColor = "Black"

    # Text box for entering secure password
    $SecurePasswordTextBox = New-Object System.Windows.Forms.TextBox
    $SecurePasswordTextBox.Location = New-Object System.Drawing.Size(10, 175)
    $SecurePasswordTextBox.Text = ''
    $SecurePasswordTextBox.BackColor = "White"
    $SecurePasswordTextBox.ForeColor = "Black"
    $SecurePasswordTextBox.ShortcutsEnabled = $True
    $SecurePasswordTextBox.Width = 400
    $SecurePasswordTextBox.PasswordChar = '*'

    # Label for the secure password text box
    $SecurePasswordLabel = New-Object System.Windows.Forms.Label
    $SecurePasswordLabel.Location = New-Object System.Drawing.Size(10, 155)
    $SecurePasswordLabel.Text = "Secure Password"
    $SecurePasswordLabel.AutoSize = $True
    $SecurePasswordLabel.BackColor = "Gray"
    $SecurePasswordLabel.ForeColor = "Black"

    # Button for running the script
    $RunScriptButton = New-Object System.Windows.Forms.Button
    $RunScriptButton.Location = New-Object System.Drawing.Size(10, 200)
    $RunScriptButton.Width = 175
    $RunScriptButton.Text = "Create Table"
    $RunScriptButton.Enabled = $true
    $RunScriptButton.BackColor = "Yellow"
    $RunScriptButton.ForeColor = "Black"

    # Event handler for the file location button
    $HDTStorageFileButton.Add_Click({
        $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog
        $FileBrowser.InitialDirectory = "C:\"
        $FileBrowser.Filter = "Excel Files (*.xlsx)|*.xlsx"
        $FileBrowserResult = $FileBrowser.ShowDialog()
        if ($FileBrowserResult -eq 'OK') {
            $Form.Tag = $FileBrowser.FileName
        }
    })

    <#
    # Event handler for the run script button
    $RunScriptButton.Add_Click({        
        $DestinationPath = "\\$WorkhorseServer\e$\AdminAppFiles\HDTStorageTables\" # Change this to your destination path
        $LoginPassword = ConvertTo-SecureString $SecurePasswordTextBox.Text -AsPlainText -Force
        $LoginCredentials = New-Object System.Management.Automation.PSCredential ($LoginUsername, $LoginPassword)
        $File = $Form.Tag
        $DestinationFile = Join-Path -Path $DestinationPath -ChildPath (Get-Item -Path $File).Name
        Copy-Item -Path $File -Destination $DestinationFile -Force
        $File = $DestinationFile
        $Instance = $DBServerTextBox.Text
        $TableName = $TableNameTextBox.Text
        Invoke-Command -ComputerName $WorkhorseServer -Authentication Credssp -Credential $LoginCredentials -ScriptBlock {
            param(
                $File, 
                $Instance, 
                $TableName
            )
            $Database = "HDTStorage"
            foreach($sheet in Get-ExcelSheetInfo $File) {
                $data = Import-Excel -Path $File -WorksheetName $sheet.name | ConvertTo-DbaDataTable
                Write-DbaDataTable -SqlInstance $Instance -Database $Database -InputObject $data -AutoCreateTable -Table $TableName
            }
        } -ArgumentList $File, $Instance, $TableName        
    })
    #>

    # Event handler for the run script button
    $RunScriptButton.Add_Click({        
        $DestinationPath = "\\$WorkhorseServer\e$\AdminAppFiles\HDTStorageTables\" # Change this to your destination path
        $LoginPassword = ConvertTo-SecureString $SecurePasswordTextBox.Text -AsPlainText -Force
        $LoginCredentials = New-Object System.Management.Automation.PSCredential ($LoginUsername, $LoginPassword)
        $File = $Form.Tag
        $DestinationFile = Join-Path -Path $DestinationPath -ChildPath (Get-Item -Path $File).Name
        Copy-Item -Path $File -Destination $DestinationFile -Force
        $File = $DestinationFile
        $Instance = $DBServerTextBox.Text
        $TableName = $TableNameTextBox.Text
        try {
            Invoke-Command -ComputerName $WorkhorseServer -Authentication Credssp -Credential $LoginCredentials -ScriptBlock {
                param(
                    $File, 
                    $Instance, 
                    $TableName
                )
                $Database = "HDTStorage"
                foreach($sheet in Get-ExcelSheetInfo $File) {
                    $data = Import-Excel -Path $File -WorksheetName $sheet.name | ConvertTo-DbaDataTable
                    Write-DbaDataTable -SqlInstance $Instance -Database $Database -InputObject $data -AutoCreateTable -Table $TableName
                }
            } -ArgumentList $File, $Instance, $TableName
    
            # Logic for successful execution
            [System.Windows.Forms.MessageBox]::Show("SQL Command executed successfully.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } catch {
            # Logic for failure
            [System.Windows.Forms.MessageBox]::Show("SQL Command failed to execute. Error: " + $_.Exception.Message, "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }        
    })

    $HDTStoragePopup.Controls.Add($HDTStorageFileButton)
    $HDTStoragePopup.Controls.Add($DBServerTextBox)
    $HDTStoragePopup.Controls.Add($TableNameTextBox)
    $HDTStoragePopup.Controls.Add($FileLocationLabel)
    $HDTStoragePopup.Controls.Add($DBServerLabel)
    $HDTStoragePopup.Controls.Add($TableNameLabel)
    $HDTStoragePopup.Controls.Add($SecurePasswordTextBox)
    $HDTStoragePopup.Controls.Add($SecurePasswordLabel)
    $HDTStoragePopup.Controls.Add($RunScriptButton)
    $HDTStoragePopup.ShowDialog()
    
    })

$Form.Controls.Add($CreateTableButton)
$Form.ShowDialog()