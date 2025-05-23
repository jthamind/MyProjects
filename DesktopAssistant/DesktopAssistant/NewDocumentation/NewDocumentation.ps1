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
$newRunspace.Name = "NewDocumentation"
$newRunspace.Open()


$GUIPowershell = [PowerShell]::Create().AddScript({

# Set location to script directory
Set-location $synchash.CWD
# Set timestamp function
function Get-Timestamp {
    return Get-Date -Format "yyyy/MM/dd hh:mm:ss"
}

# Initial setup for the directories and Word documentation template
function Test-Directory {
    $ParentPath = "$env:USERPROFILE\OneDrive - Allied Solutions\Documents\Documentation\NewTemplate"
    $NetworkPath = "\\dfssprdawfs01\infotechshare\DocumentationTemplate"
    if (!(Test-Path -Path $ParentPath)) {
    mkdir $ParentPath
    Copy-Item -Path "$NetworkPath\DocTemplate.docx" -Destination $ParentPath
    }
}

Test-Directory

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

# New documentation template button
$NewDocButton = New-Object System.Windows.Forms.Button
$NewDocButton.Location = New-Object System.Drawing.Size(125, 25)
$NewDocButton.Text = "Create"
$NewDocButton.Enabled = $False
$NewDocButton.BackColor = "Yellow"
$NewDocButton.ForeColor = "Black"

# Create new documentation template text box
$NewDocTextBox = New-Object System.Windows.Forms.TextBox
$NewDocTextBox.Location = New-Object System.Drawing.Size(10, 25)
$NewDocTextBox.Text = ''
$NewDocTextBox.BackColor = "Yellow"
$NewDocTextBox.ForeColor = "Black"
$NewDocTextBox.ShortcutsEnabled = $True

# New Documentation Template text box logic
# If text box is empty, disable the New Document button
$NewDocTextBox.Add_TextChanged({
    if ($NewDocTextBox.Text -ne '') {
        $NewDocButton.Enabled = $True 
    }
    else {
        $NewDocButton.Enabled = $False
    }
})


# New documentation template function
Function New-DocTemplate {
    param (
        [string]$DocTopic
    )
# Variables
$TemplateFile = "C:\Users\jewilliams1\OneDrive - Allied Solutions\Documents\Documentation\NewTemplate\DocTemplate.docx"
$FindText = "Documentation Template"
$DocTopic = $NewDocTextBox.Text
$NewFile = "C:\$DocTopic.docx"
$MatchCase = $false
$MatchWholeWorld = $true
$MatchWildcards = $false
$MatchSoundsLike = $false
$MatchAllWordForms = $false
$Forward = $false
$Wrap = 1
$Format = $false
$Replace = 2

# Creates a new instance of the Word application using the Component Object Model (COM)
$Word = New-Object -ComObject Word.Application

# Open the document
$Document = $Word.Documents.Open("$TemplateFile")

# Get the first section of the document
$section = $document.Sections.Item(1)

# Get the header of the first section
$header = $section.Headers.Item(1)

# Find and replace the date in the header
$searchText = "1/1/2023"
$NewDate = (Get-Date).ToString("M/d/yyyy")
$header.Range.Find.Execute($searchText, $false, $false, $false, $false, $false, $true, 1, $false, $NewDate, 2)

# Find and replace text using the variables above
$Document.Content.Find.Execute($FindText, $MatchCase, $MatchWholeWorld, $MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $Wrap, $Format, $DocTopic, $Replace)

try {

    # Save a new copy of the document
    $Document.SaveAs("$NewFile")
} 
catch {
    # Catch any exceptions that occur during file save
    $OutText.AppendText("$(Get-Timestamp) - An error occurred while trying to create the new Word document. This is likely due to the file size exceeding 255 characters.`r`n")
    $Document.Close([Microsoft.Office.Interop.Word.WdSaveOptions]::wdDoNotSaveChanges)
}

# Close the Word application
$Word.Quit()

# Release the COM object
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word)

# Set the variable to $null
$Word = $null

$DocumentationPath = "C:\Users\jewilliams1\OneDrive - Allied Solutions\Documents\Documentation"
New-item -Path $DocumentationPath -Name "$DocTopic" -ItemType "directory"
Move-Item -Path $NewFile -Destination "$DocumentationPath\$DocTopic"
Invoke-Item -Path "$DocumentationPath\$DocTopic\$DocTopic.docx"
$OutText.AppendText("$(Get-Timestamp) - New document created at $DocumentationPath\$DocTopic\$DocTopic.docx`r`n")

}

function Test-ValidFileName
{
    param([string]$DocTopic)

    $IndexOfInvalidChar = $DocTopic.IndexOfAny([System.IO.Path]::GetInvalidFileNameChars())

    # IndexOfAny() returns the value -1 to indicate no such character was found
    return $IndexOfInvalidChar -eq -1
}

# Function to check doc length name
function Test-DocLength
{
    param([string]$DocTopic)

    $DocTopic = $NewDocTextBox.Text
    if ($DocTopic.length -gt 200) {
        $OutText.AppendText("$(Get-Timestamp) - Please enter a document name less than 200 characters`r`n")
        return
    }
    elseif (Test-ValidFileName $DocTopic) {
        $OutText.AppendText("$(Get-Timestamp) - $DocTopic is heckin valid`r`n")
        New-DocTemplate -DocTopic $DocTopic
    }
    else {
        $OutText.AppendText("$(Get-Timestamp) - Please enter a document name without any of the following characters: \ / : * ? < > |`r`n")
    }  
}

# New Document button press logic
$NewDocButton.Add_Click({
    Test-DocLength 
})

# New Document text box Enter key logic
$NewDocTextBox.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") {
        Test-DocLength        
    }
})

# New Document text box Ctrl+A logic
$NewDocTextBox.Add_KeyDown({
    if ($_.KeyCode -eq "A" -and $_.Control -eq "True") {
        $NewDocTextBox.SelectAll()    
    }
    })

# Build Form
$Form.Controls.Add($NewDocButton)
$Form.Controls.Add($NewDocTextBox)
$Form.Controls.Add($OutText)

# Show Form
$Form.ShowDialog() | Out-Null

})

$GUIPowershell.Runspace = $newRunspace
$async = $GUIPowershell.BeginInvoke()
$GuiPowerShell.EndInvoke($async)