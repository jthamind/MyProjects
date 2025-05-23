# Get config values from json file
$ConfigValues = Get-Content -Path "C:\Users\jewilliams1\Documents\DesktopAssistant\MainGUI\Config.json" | ConvertFrom-Json
$UserProfilePath = [Environment]::GetEnvironmentVariable("USERPROFILE")

$StopScriptPath = $ConfigValues.StopLFPServicesQAScript.Replace("{USERPROFILE}", $UserProfilePath)
$StartScriptPath = $ConfigValues.StartLFPServicesQAScript.Replace("{USERPROFILE}", $UserProfilePath)
$TestingTeamsWebhook = $ConfigValues.TestingWebhook


# Post Teams message using the UniTrac Deployments webhook
$WebhookURL = $TestingTeamsWebhook

# Create the message body
$messageBody = @{
    "@type" = "MessageCard"
    "@context" = "http://schema.org/extensions"
    "summary" = "Deployment message for stopping QA services in order to add a lender to the LFP services"
    "sections" = @(
        @{
            "text" = "This is going to post a message and then stop the LFP services in QA"
        }
    )
} | ConvertTo-Json

# Post the message to Teams
Invoke-RestMethod -Uri $WebhookURL -Method Post -Body $messageBody -ContentType "application/json" | Out-Null


if (Test-Path $StopScriptPath) {
    & $StopScriptPath
    if ($LASTEXITCODE -eq 0) {
        Write-Host "External script executed successfully."
    } else {
        Write-Host "External script failed to execute."
    }
} else {
    Write-Host "The script $StopScriptPath does not exist."
}

Write-Host "Waiting for 10 seconds before continuing..."

Start-Sleep -Seconds 10

Write-Host "Continuing..."

if (Test-Path $StartScriptPath) {
    & $StartScriptPath
    if ($LASTEXITCODE -eq 0) {
        Write-Host "External script executed successfully."
    } else {
        Write-Host "External script failed to execute."
    }
} else {
    Write-Host "The script $StartScriptPath does not exist."
}

# Post Teams message using the UniTrac Deployments webhook
$messageBody = @{
    "@type" = "MessageCard"
    "@context" = "http://schema.org/extensions"
    "summary" = "Deployment message for stopping QA services in order to add a lender to the LFP services"
    "sections" = @(
        @{
            "text" = "I just started the QA services back up. :)"
        }
    )
} | ConvertTo-Json

# Post the message to Teams
Invoke-RestMethod -Uri $WebhookURL -Method Post -Body $messageBody -ContentType "application/json" | Out-Null