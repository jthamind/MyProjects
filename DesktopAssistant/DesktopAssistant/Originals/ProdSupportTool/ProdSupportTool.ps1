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
$newRunspace.Name = "PSTool"
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
$Form.FormBorderStyle = "Fixed3D"
$Form.MaximizeBox = $False
$Form.MinimizeBox = $True
$Form.ControlBox = $True
$Form.I1n = $Icon
$Form.TopMost = $True
$Form.StartPosition = "CenterScreen"

# Environment variables
$RemoteEnvSelect = "\\unitrac-wh001\E$\Software\ProdSupportEnvSelect\"
$RemoteSupportTool = "\\unitrac-wh001\E$\Software\ProdSupportTool\"
$LocalSupportTool = "$env:USERPROFILE\OneDrive - Allied Solutions\Documents\Software\ProdSupportTool"
$LocalEnvSelect = "$env:USERPROFILE\OneDrive - Allied Solutions\Documents\Software\ProdSupportEnvSelect\"
$RemoteConfigs = "\\unitrac-wh001\E$\AdminAppFiles\ProdSupportConfigs\"
$LocalConfigs = "$env:USERPROFILE\OneDrive - Allied Solutions\Documents\Software\ProdSupportConfigs"

# List box to select environment
$List = New-Object System.Windows.Forms.ComboBox
$List.Text = 'Select Environment'
$List.Width = 150
$List.AutoSize = $true
$List.Location = New-Object System.Drawing.Size(10, 25)


# Add items to the list box
@('QA', 'Stage', 'Production') | ForEach-Object { [void]$List.Items.Add($_) }
$List.SelectedIndex = 0
$List.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)

# Text box for displaying output
$OutText = New-Object System.Windows.Forms.TextBox
$OutText.Location = New-Object System.Drawing.Size(10, 75)
$OutText.Size = New-Object System.Drawing.Size(450, 200)
$OutText.Multiline = $true
$OutText.ScrollBars = "Vertical"
$OutText.Enabled = $false

# Button for starting ProdSupportTool executable
$LaunchProdSupportToolBtn = New-Object System.Windows.Forms.Button
$LaunchProdSupportToolBtn.Location = New-Object System.Drawing.Size(260, 50)
$LaunchProdSupportToolBtn.Width = 150
$LaunchProdSupportToolBtn.Text = "Run Prod Support Tool"
$LaunchProdSupportToolBtn.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)

# Button for switching environment
$SwitchBtn = New-Object System.Windows.Forms.Button
$SwitchBtn.Location = New-Object System.Drawing.Size(260, 25)
$SwitchBtn.Width = 150
$SwitchBtn.Text = "Select Environment"
$SwitchBtn.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)

# Button for resetting environment
$ResetBtn = New-Object System.Windows.Forms.Button
$ResetBtn.Location = New-Object System.Drawing.Size(175, 25)
$ResetBtn.Width = 75
$ResetBtn.Text = "Reset"
$ResetBtn.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
$ResetBtn.Enabled = $false

# Button for checking for newer version of Prod Support Tool
$VersionCheckBtn = New-Object System.Windows.Forms.Button
$VersionCheckBtn.Location = New-Object System.Drawing.Size(10, 50)
$VersionCheckBtn.Width = 75
$VersionCheckBtn.Text = "Refresh"
$VersionCheckBtn.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Regular)
$VersionCheckBtn.Enabled = $true

Function Reset-Environment {
    Remove-Item -Path "$LocalSupportTool\ProductionSupportTool.exe.config"
    Rename-Item "$LocalSupportTool\ProductionSupportTool.exe.config.old" -NewName "$LocalSupportTool\ProductionSupportTool.exe.config"
    $OutText.AppendText("$(Get-Timestamp) - Environment has been reset. Please select an environment.`r`n")
    $ResetBtn.Enabled = $false
    $SwitchBtn.Enabled = $true
    $List.Enabled = $true
}

$SwitchBtn.Add_Click({
    If (Test-Path "$LocalSupportTool\ProductionSupportTool.exe.config.old") {
        $OutText.AppendText("$(Get-Timestamp) - Existing config found. Cleaning up.`r`n")   
        Reset-Environment
    }
    else {
        $ResetBtn.Enabled = $false
        $OutText.AppendText("$(Get-Timestamp) - No existing config found. Please select environment.`r`n")
               
    }
    # Adding code to a button, so that when clicked, it switches environments            
    $environment = $List.SelectedItem
    $Form.Text = "Prod Support Tool Environment Select - $environment"
    try {
        switch ($environment) {
            "QA" {
                $OutText.AppendText("$(Get-Timestamp) - Entering QA environment.`r`n")
                $runningConfig = "$LocalConfigs\ProductionSupportTool.exe_QA.config"
            }
            "Stage" {
                $OutText.AppendText("$(Get-Timestamp) - Entering Staging environment.`r`n")
                $runningConfig = "$LocalConfigs\ProductionSupportTool.exe_Stage.config"
            }
            "Production" {
                $OutText.AppendText("$(Get-Timestamp) - Entering Production environment.`r`n")
                $runningConfig = "$LocalConfigs\ProductionSupportTool.exe_Prod.config"
            }
            # Behavior if no option is selected
            Default {
                $OutText.AppendText("$(Get-Timestamp) - Please make a valid selection or reset.`r`n")
                throw "No selection Made"
            }
        }
        Start-Sleep -Seconds 1
        # Rename Current Running config and Copy configuration file for correct environment 
        Rename-Item "$LocalSupportTool\ProductionSupportTool.exe.config" -NewName "$LocalSupportTool\ProductionSupportTool.exe.config.old"
        Copy-Item $runningConfig -Destination "$LocalSupportTool\ProductionSupportTool.exe.config"
        $OutText.AppendText("$(Get-Timestamp) - You are good to run in the $environment environment.`r`n")
        $List.Enabled = $false
        $ResetBtn.Enabled = $true
        $SwitchBtn.Enabled = $false
    }
    catch {
        $OutText.Text = $_
    }
})

$ResetBtn.Add_Click({
    Reset-Environment
    
})

$VersionCheckBtn.Add_Click({
    # If the remote folder is newer than the local folder, copy the files from the remote folder
if ($RemoteSupportToolFolderDate -gt $LocalSupportToolFolderDate) {
    $OutText.AppendText("$(Get-Timestamp) - Remote Prod Support Tool folder is newer than local folder.`r`n")
    $OutText.AppendText("$(Get-Timestamp) - Copying files from remote folder to local folder.`r`n")

    Copy-Item -Path $RemoteSupportTool* -Destination $LocalSupportTool -Recurse -Force

    $OutText.AppendText("$(Get-Timestamp) - Files successfully copied.`r`n")
    $OutText.AppendText("$(Get-Timestamp) - Local Prod Support Tool folder is now up to date.`r`n")
}
else {
    $OutText.AppendText("$(Get-Timestamp) - Local Prod Support Tool folder is up to date.`r`n")
}
})
    

# Get the last write time of the remote and local Prod Support Tool folders
$RemoteSupportToolFolderDate = Get-ChildItem $RemoteSupportTool -Exclude *.txt -Recurse | ForEach-Object {$_.LastWriteTime} | Sort-Object -Descending | Select-Object -First 1
$LocalSupportToolFolderDate = Get-ChildItem $LocalSupportTool -Exclude *.txt -Recurse | ForEach-Object {$_.LastWriteTime} | Sort-Object -Descending | Select-Object -First 1

# Get the number of files in the local folders for later comparison
$LocalSupportToolFolderCount = Get-ChildItem $LocalSupportTool | Measure-Object
$LocalEnvSelectFolderCount = Get-ChildItem $LocalEnvSelect | Measure-Object
$LocalConfigsFolderCount = Get-ChildItem $LocalConfigs | Measure-Object

# If the local Environment Select folder is empty, copy the files from the remote folder
if ($LocalEnvSelectFolderCount.count -eq 0) {
    $OutText.AppendText("$(Get-Timestamp) - Local folder is empty.`r`n")
    $OutText.AppendText("$(Get-Timestamp) - Copying files from remote Environment Select folder to local folder.`r`n")

    Copy-Item -Path $RemoteEnvSelect* -Destination $LocalEnvSelect -Recurse -Force

    $OutText.AppendText("$(Get-Timestamp) - Files copied successfully.`r`n")
    $OutText.AppendText("$(Get-Timestamp) - Local folder is now up to date.`r`n")
}
else {
    $OutText.AppendText("$(Get-Timestamp) - Local Environment Select folder is up to date.`r`n")
}

# If the local Prod Support Tool folder is empty, copy the files from the remote folder
if ($LocalSupportToolFolderCount.count -eq 0) {
    $OutText.AppendText("$(Get-Timestamp) - Local Prod Support Tool folder is empty.`r`n")
    $OutText.AppendText("$(Get-Timestamp) - Copying files from remote Prod Support Tool folder to local folder.`r`n")

    Copy-Item -Path $RemoteSupportTool* -Destination $LocalSupportTool -Recurse -Force

    $OutText.AppendText("$(Get-Timestamp) - Files successfully copied.`r`n")
    $OutText.AppendText("$(Get-Timestamp) - Local Prod Support Tool folder is now up to date.`r`n")
}
else {
    $OutText.AppendText("$(Get-Timestamp) - Local Prod Support Tool files are present.`r`n")
}

# If local ProdSupportConfigs folder is empty, copy the files from the remote folder
if ($LocalConfigsFolderCount.count -eq 0) {
    $OutText.AppendText("$(Get-Timestamp) - Local Prod Support Configs folder is empty.`r`n")
    $OutText.AppendText("$(Get-Timestamp) - Copying files from remote folder to local folder.`r`n")

    Copy-Item -Path $RemoteConfigs* -Destination $LocalConfigs -Recurse -Force

    $OutText.AppendText("$(Get-Timestamp) - Files successfully copied.`r`n")
    $OutText.AppendText("$(Get-Timestamp) - Local Prod Support Configs folder is now up to date.`r`n")
}
else {
    $OutText.AppendText("$(Get-Timestamp) - Local Prod Support Configs folder is up to date.`r`n")
}

# Button for launching Prod Support Tool
$LaunchProdSupportToolBtn.Add_Click({
    Start-Process -FilePath "$LocalSupportTool\ProductionSupportTool.exe"
})

# Add controls to form
$Form.Controls.Add($List)
$Form.Controls.Add($SwitchBtn)
$Form.Controls.Add($ResetBtn)
$Form.Controls.Add($LaunchProdSupportToolBtn)
$Form.Controls.Add($OutText)
$Form.Controls.Add($VersionCheckBtn)

# Show Form
$Form.ShowDialog() | Out-Null
})

$GUIPowershell.Runspace = $newRunspace
$async = $GUIPowershell.BeginInvoke()
$GuiPowerShell.EndInvoke($async)