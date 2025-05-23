Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

# Create the form
$Form = New-Object System.Windows.Forms.Form
$Form.Text = 'AWS EC2 Instance Describer'
$Form.Size = New-Object System.Drawing.Size(600,800)
$Form.StartPosition = 'CenterScreen'

# Create the ListBox for account selection
$script:AWSAccountsListBox = New-Object System.Windows.Forms.ListBox
$script:AWSAccountsListBox.Location = New-Object System.Drawing.Point(10,10)
$script:AWSAccountsListBox.Size = New-Object System.Drawing.Size(560,150)
$Form.Controls.Add($script:AWSAccountsListBox)

# Create the TextBox to display the output
$TextBox = New-Object System.Windows.Forms.TextBox
$TextBox.Location = New-Object System.Drawing.Point(10,200)
$TextBox.Size = New-Object System.Drawing.Size(560,540)
$TextBox.Multiline = $true
$TextBox.ScrollBars = 'Vertical'
$TextBox.ReadOnly = $true
$Form.Controls.Add($TextBox)

# Create the Button to describe instances
$Button = New-Object System.Windows.Forms.Button
$Button.Location = New-Object System.Drawing.Point(10,170)
$Button.Size = New-Object System.Drawing.Size(120,30)
$Button.Text = 'Describe Instances'

# Function for updating the Account ID in the AWS config file
# This is used to easily switch between accounts without needing to reauthenticate
function Update-AccountID {
    param (
        [string]$AccountId,
        [string]$AWSSSOProfile,
        [string]$AWSConfigFile
    )

    $OutText.AppendText("$(Get-Timestamp) - Updating account ID to $AccountId`r`n")

    # Paths to AWS config files
    $AWSSSOConfigFilePath = Join-Path $env:USERPROFILE $AWSConfigFile

    # Print config file path
    $OutText.AppendText("$(Get-Timestamp) - Config file path: $AWSSSOConfigFilePath`r`n")

    # Check if the config file exists
    if (Test-Path $AWSSSOConfigFilePath) {
        # Read the current configuration
        $ConfigContent = Get-Content $AWSSSOConfigFilePath -Raw

        # Regular expression to find the profile section and the sso_account_id line
        $ProfileSectionPattern = "\[profile\s+$AWSSSOProfile\](.*?)sso_account_id\s*=.*?(\r?\n)"
        $ReplacementText = "[profile $AWSSSOProfile]`$1sso_account_id = $AccountId`$2"

        # Update the sso_account_id in the profile section
        $UpdatedConfigContent = [regex]::Replace($ConfigContent, $ProfileSectionPattern, $ReplacementText, [System.Text.RegularExpressions.RegexOptions]::Singleline)

        # Write the updated configuration back to the file
        Set-Content -Path $AWSSSOConfigFilePath -Value $UpdatedConfigContent
        $OutText.AppendText("$(Get-Timestamp) - Updated account ID to $AccountId in config file`r`n")
    } else {
        $OutText.AppendText("$(Get-Timestamp) - AWS config file not found.`r`n")
    }
}

# Function for running the AWS describe-instances command
function Invoke-DescribeInstances {
    param (
        [string]$AWSSSOProfile
    )

    # Describe the instances in the selected account
    try {
        # Ensure you have the necessary permissions to use this filter from your current account
        $EC2Result = aws ec2 describe-instances --region us-east-2 --query "Reservations[*].Instances[*].[Tags[?Key=='Name'].Value | [0], State.Name]" --output text --profile $AWSSSOProfile
        $ResultforTextbox = $EC2Result | Out-String
        $OutText.AppendText("$(Get-Timestamp) - $ResultforTextbox`r`n")
    } catch {
        $OutText.AppendText("$(Get-Timestamp) - Failed to describe instances for account: $selectedAccountName`r`n")
    }
}

$Button.Add_Click({
    if ($script:AWSAccountsListBox.SelectedItem -ne $null) {
        $selectedAccountName = $script:AWSAccountsListBox.SelectedItem.ToString()

        # Extract the accountId for the selected account from the accounts list
        $selectedAccount = $AWSAccountsList.accountList | Where-Object { $_.accountName -eq $selectedAccountName }
        $selectedAccountId = $selectedAccount.accountId

        $OutText.AppendText("$($TimestampFunction.Invoke()) - Updating account ID to $AccountId`r`n")

        # Paths to AWS config files
        $AWSSSOConfigFilePath = Join-Path $env:USERPROFILE $AWSConfigFile
    
        # Print config file path
        $OutText.AppendText("$($TimestampFunction.Invoke()) - Config file path: $AWSSSOConfigFilePath`r`n")

        Update-AccountID -accountId $selectedAccountId -AWSSSOProfile $AWSSSOProfile -AWSSSOConfigFilePath $AWSSSOConfigFilePath

        if ($selectedAccountId) {
            # Describe the instances in the selected account
            Invoke-DescribeInstances -AWSSSOProfile $AWSSSOProfile -OutText $OutText -TimestampFunction ${function:Get-Timestamp}
        }
    } else {
        $OutText.AppendText("$(Get-Timestamp) - Please select an account from the list.`r`n")
    }
})

$Form.Controls.Add($Button)

# Function to list all AWS accounts
function Get-AWSAccounts {
    param (
        [string]$AWSSSOProfile,
        [string]$AccessToken,
        [string]$AWSAccountsFile,
        [string]$AWSSSOCacheFilePath,
        [System.Windows.Forms.TextBox]$OutText
    )

    # Get the most recent cache file
    $AWSSSOCacheFiles = Get-ChildItem -Path $AWSSSOCacheFilePath -Filter "*.json" | Sort-Object LastWriteTime -Descending
    $AccessTokenFile = $AWSSSOCacheFiles | Select-Object -First 1

    # Read the access token from the cache file
    $AccessTokenContent = Get-Content -Path $AccessTokenFile.FullName | ConvertFrom-Json
    $AccessToken = $AccessTokenContent.accessToken

    # Run the AWS SSO list-accounts command and capture the output in a json file
    aws sso list-accounts --profile $AWSSSOProfile --access-token $AccessToken --output json > $AWSAccountsFile

    try {
        $jsonString = Get-Content -Path $AWSAccountsFile -Raw
        $AWSAccountsList = $jsonString | ConvertFrom-Json -ErrorAction Stop
        $AWSAccountsList.accountList | ForEach-Object {
            [void]$script:AWSAccountsListBox.Items.Add($_.accountName)
        }
    } catch {
        $OutText.AppendText("$(Get-Timestamp) - Failed to load or parse JSON: $_`r`n")
    }
}

Get-AWSAccounts -AWSSSOProfile $AWSSSOProfile -AccessToken $AccessToken -AWSAccountsFile $AWSAccountsFile -OutText $OutText

# Show the form
$Form.ShowDialog()