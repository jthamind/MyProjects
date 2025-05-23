param(
    [string]$AWSSSOLoginScript,
    [string]$AWSSSOProfile,
    [string]$AWSSSOLoginConfirm
)
# The above parameters are passed in from the AddLendertoLFPServices.ps1 script

return "inside subscript"

# Start AWS SSO login and redirect output to a file
Start-Process -FilePath $AWSSSOLoginScript -ArgumentList "sso login --profile $AWSSSOProfile" -RedirectStandardOutput $AWSSSOLoginConfirm -NoNewWindow

# Continuously read the file
while (Test-Path $AWSSSOLoginConfirm) {
    # Read new content from the file
    $LoginConfirmCode = Get-Content $outputFile -Wait -Tail 1

    # Delete the AWSSSOLoginConfirm file to exit the loop in the next iteration
    Remove-Item $AWSSSOLoginConfirm -Force
}

$LoginConfirmCode


# ! THIS WORKS AS BASE LOGIC
<# $profileName = "DesktopAssistant"
$outputFile = "C:\Users\jewilliams1\Downloads\aws-sso-output.txt"
if (!(Test-Path $outputFile)) {
    New-Item $outputFile -ItemType File | Out-Null
}

# Start AWS SSO login and redirect output to a file
#Start-Process -FilePath "aws" -ArgumentList "sso login --profile $profileName" -RedirectStandardOutput $outputFile -NoNewWindow

aws sso login --profile $profileName *> $outputFile

# Continuously read the file
while ($true) {
    # Read new content from the file
    Get-Content $outputFile -Wait -Tail 1
} #>