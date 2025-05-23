<# ! This section works and lists each column header and then the values below it#>
<#
$csvPath = "C:\Users\jewilliams1\Downloads\servers.csv"
$csv = Import-CSV $csvPath
$csvHeaders = ($csv | Get-Member -MemberType NoteProperty).name

foreach ($header in $csvHeaders) {
    # Display the column header
    Write-Host "Column header: $header"

    # Get the rest of the values in the column and filter out empty values
    $restOfValues = ($csv | Select-Object -ExpandProperty $header) | Where-Object { $_ -ne '' }
    Write-Host "Rest of the values: $($restOfValues -join ', ')"
}
#>


Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create the form
$form = New-Object Windows.Forms.Form
$form.Text = "ComboBox with Headers and Values"
$form.Size = New-Object Drawing.Size(400, 300)

# Create the ComboBox
$comboBox = New-Object Windows.Forms.ComboBox
$comboBox.Location = New-Object Drawing.Point(10, 10)
$comboBox.Width = 380
$combobox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList

# Add headers as bold items and the rest of the values for each column
$csvPath = "C:\Users\jewilliams1\Downloads\servers.csv"
$csv = Import-CSV $csvPath
$csvHeaders = ($csv | Get-Member -MemberType NoteProperty).name

foreach ($header in $csvHeaders) {
    # Add the header as a bold item
    $headerItem = New-Object Windows.Forms.ComboBoxItem
    $headerItem.Text = $header
    $headerItem.Font = New-Object Drawing.Font("Microsoft Sans Serif", 10, [System.Drawing.FontStyle]::Bold)
    $comboBox.Items.Add($header)

    # Get the rest of the values in the column and filter out empty values
    $restOfValues = ($csv | Select-Object -ExpandProperty $header) | Where-Object { $_ -ne '' }

    # Add the rest of the values under the header item
    $comboBox.Items.AddRange($restOfValues)
}

# Show the form
$form.Controls.Add($comboBox)
$form.ShowDialog()
