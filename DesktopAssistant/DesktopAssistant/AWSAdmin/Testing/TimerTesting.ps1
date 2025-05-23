Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

# Read the target date and time from the text file
$filePath = (Get-ChildItem -Path "C:\Users\jewilliams1\.aws\sso\cache" | Sort-Object LastWriteTime -Descending | Select-Object -First 1 -Property FullName).FullName
$json = Get-Content -Path $filePath | Out-String
$expiresAt = ($json | ConvertFrom-Json).expiresAt

# Ensure UTC is respected in the target time
try {
    # Parse the date using the format it appears in the file
    $targetDateTime = [datetime]::ParseExact($expiresAt, "MM/dd/yyyy HH:mm:ss", $null)
} catch {
    Write-Host "Error parsing date: $_"
}

# Create a form
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Countdown Timer'
$form.Size = New-Object System.Drawing.Size(300,200)

# Create a AWSSSOLoginTimerLabel to display the countdown
$AWSSSOLoginTimerLabel = New-Object System.Windows.Forms.Label
$AWSSSOLoginTimerLabel.Location = New-Object System.Drawing.Point(10,10)
$AWSSSOLoginTimerLabel.Size = New-Object System.Drawing.Size(280,150)
$AWSSSOLoginTimerLabel.Font = New-Object System.Drawing.Font('Arial',16)
$form.Controls.Add($AWSSSOLoginTimerLabel)

# Timer to update the AWSSSOLoginTimerLabel every second
$AWSSSOLoginTimer = New-Object System.Windows.Forms.Timer
$AWSSSOLoginTimer.Interval = 1000 # Update every second
$AWSSSOLoginTimer.Add_Tick({
    $NowUTC = (Get-Date).ToUniversalTime()
    $TimeLeft = $targetDateTime - $NowUTC

    if ($TimeLeft -le [TimeSpan]::Zero) {
        $AWSSSOLoginTimerLabel.Text = "Login Expired"
        $AWSSSOLoginTimer.Stop()
    } else {
        $AWSSSOLoginTimerLabel.Text = ($TimeLeft.ToString("dd\.hh\:mm\:ss"))
    }
})
$AWSSSOLoginTimer.Start()

# Show the form
$form.ShowDialog()

$form.Add_FormClosing({
    $AWSSSOLoginTimer.Stop()
    $AWSSSOLoginTimer.Dispose()
})