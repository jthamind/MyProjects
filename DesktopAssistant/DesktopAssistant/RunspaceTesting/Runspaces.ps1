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

$Textbox = New-Object System.Windows.Forms.TextBox
$Textbox.Location = New-Object System.Drawing.Size(25, 75)
$Textbox.Size = New-Object System.Drawing.Size(350, 200)
$Textbox.Multiline = $true
$Textbox.ScrollBars = "Vertical"
$Textbox.Enabled = $True
$Textbox.ReadOnly = $True
$Textbox.BackColor = "White"
$Textbox.ForeColor = "Black"

# New documentation template button
$NewDocButton = New-Object System.Windows.Forms.Button
$NewDocButton.Location = New-Object System.Drawing.Size(125, 25)
$NewDocButton.Text = "Create"
$NewDocButton.Enabled = $False
$NewDocButton.BackColor = "Yellow"
$NewDocButton.ForeColor = "Black"

# Create new documentation template text box
$newDocumentTxtBx = New-Object System.Windows.Forms.TextBox
$newDocumentTxtBx.Location = New-Object System.Drawing.Size(10, 25)
$newDocumentTxtBx.Text = ''
$newDocumentTxtBx.BackColor = "Yellow"
$newDocumentTxtBx.ForeColor = "Black"

# New Documentation Template text box logic
# If text box is empty, disable the New Document button
$newDocumentTxtBx.Add_TextChanged({
    if ($newDocumentTxtBx.Text -ne '') {
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
$DocTopic = $newDocumentTxtBx.Text
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
    <#Do this if a terminating exception happens#>
    Write-Host "An error occurred while trying to create the new Word document"
    Write-Host "This is likely due to the file size exceeding 255 characters"
    $Document.Close([Microsoft.Office.Interop.Word.WdSaveOptions]::wdDoNotSaveChanges)
}

# Close the Word application
$Word.Quit()

# Release the COM object
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word)

# Set the variable to $null
$Word = $null

$DocumentationPath = "C:\Users\jewilliams1\OneDrive - Allied Solutions\Documents\Documentation"
Write-Host "DocumentationPath is $DocumentationPath"
Write-Host "New directory is $DocumentationPath/$DocTopic"
New-item -Path $DocumentationPath -Name "$DocTopic" -ItemType "directory"
Move-Item -Path $NewFile -Destination "$DocumentationPath\$DocTopic"
Invoke-Item -Path "$DocumentationPath\$DocTopic\$DocTopic.docx"

}

function Test-ValidFileName
{
    param([string]$DocTopic)

    $IndexOfInvalidChar = $DocTopic.IndexOfAny([System.IO.Path]::GetInvalidFileNameChars())

    # IndexOfAny() returns the value -1 to indicate no such character was found
    return $IndexOfInvalidChar -eq -1
}

# New Document button logic
$NewDocButton.Add_Click({
    $DocTopic = $newDocumentTxtBx.Text
    if ($DocTopic.length -gt 200) {
        Write-Host "$DocTopic is too long"
        return
    }
    elseif (Test-ValidFileName $DocTopic) {
        Write-Host "$DocTopic is heckin valid"
        New-DocTemplate -DocTopic $DocTopic
    }
    else {
        Write-Host "$DocTopic contains invalid characters"
    }  
})


# Build Form
$Form.Controls.Add($NewDocButton)
$Form.Controls.Add($newDocumentTxtBx)
$Form.Controls.Add($Textbox)

# Show Form
$Form.ShowDialog() | Out-Null

})

$GUIPowershell.Runspace = $newRunspace
$async = $GUIPowershell.BeginInvoke() 


<#

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

$Textbox = New-Object System.Windows.Forms.TextBox
$Textbox.Location = New-Object System.Drawing.Size(25, 75)
$Textbox.Size = New-Object System.Drawing.Size(350, 200)
$Textbox.Multiline = $true
$Textbox.ScrollBars = "Vertical"
$Textbox.Enabled = $True
$Textbox.ReadOnly = $True
$Textbox.BackColor = "White"
$Textbox.ForeColor = "Black"

# New documentation template button
$NewDocButton = New-Object System.Windows.Forms.Button
$NewDocButton.Location = New-Object System.Drawing.Size(125, 25)
$NewDocButton.Text = "Create"
$NewDocButton.Enabled = $False
$NewDocButton.BackColor = "Yellow"
$NewDocButton.ForeColor = "Black"

# Create new documentation template text box
$newDocumentTxtBx = New-Object System.Windows.Forms.TextBox
$newDocumentTxtBx.Location = New-Object System.Drawing.Size(10, 25)
$newDocumentTxtBx.Text = ''
$newDocumentTxtBx.BackColor = "Yellow"
$newDocumentTxtBx.ForeColor = "Black"

# New Documentation Template text box logic
# If text box is empty, disable the New Document button
$newDocumentTxtBx.Add_TextChanged({
    if ($newDocumentTxtBx.Text -ne '') {
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
$DocTopic = $newDocumentTxtBx.Text
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
    # Do this if a terminating exception happens
    Write-Host "An error occurred while trying to create the new Word document"
    Write-Host "This is likely due to the file size exceeding 255 characters"
    $Document.Close([Microsoft.Office.Interop.Word.WdSaveOptions]::wdDoNotSaveChanges)
}

# Close the Word application
$Word.Quit()

# Release the COM object
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word)

# Set the variable to $null
$Word = $null

$DocumentationPath = "C:\Users\jewilliams1\OneDrive - Allied Solutions\Documents\Documentation"
Write-Host "DocumentationPath is $DocumentationPath"
Write-Host "New directory is $DocumentationPath/$DocTopic"
New-item -Path $DocumentationPath -Name "$DocTopic" -ItemType "directory"
Move-Item -Path $NewFile -Destination "$DocumentationPath\$DocTopic"
Invoke-Item -Path "$DocumentationPath\$DocTopic\$DocTopic.docx"

}

function Test-ValidFileName
{
    param([string]$DocTopic)

    $IndexOfInvalidChar = $DocTopic.IndexOfAny([System.IO.Path]::GetInvalidFileNameChars())

    # IndexOfAny() returns the value -1 to indicate no such character was found
    return $IndexOfInvalidChar -eq -1
}

# New Document button logic
$NewDocButton.Add_Click({
    $DocTopic = $newDocumentTxtBx.Text
    if ($DocTopic.length -gt 200) {
        Write-Host "$DocTopic is too long"
        return
    }
    elseif (Test-ValidFileName $DocTopic) {
        Write-Host "$DocTopic is heckin valid"
        New-DocTemplate -DocTopic $DocTopic
    }
    else {
        Write-Host "$DocTopic contains invalid characters"
    }  
})


# Build Form
$Form.Controls.Add($NewDocButton)
$Form.Controls.Add($newDocumentTxtBx)
$Form.Controls.Add($Textbox)

# Show Form
$Form.ShowDialog() | Out-Null

})

$GUIPowershell.Runspace = $newRunspace
$async = $GUIPowershell.BeginInvoke() 

#>