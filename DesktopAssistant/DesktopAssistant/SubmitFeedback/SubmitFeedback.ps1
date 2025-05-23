param (
    [string]$TestingWebhookURL,
    [string]$UserFeedback,
    [string]$FeedbackUserName,
    $OutTextControl,
    [scriptblock]$TimestampFunction
)

# Get the directory of the currently executing script
$ScriptDir = $PSScriptRoot

# Build the path to the JSON file
$ConfigFilePath = "..\MainGUI\Config.json"

# Combine them to get the full path to the JSON file
$FullJsonFilePath = Join-Path -Path $ScriptDir -ChildPath $ConfigFilePath

# Get config values from json file
$ConfigValues = Get-Content -Path $FullJsonFilePath | ConvertFrom-Json
$WorkhorseServer = $ConfigValues.WorkhorseServer

try {
    Invoke-Command -ComputerName $WorkhorseServer -ScriptBlock {
        param(
                    $TestingWebhookURL, $UserFeedback, $FeedbackUserName, $TimestampFunction
            )
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    # Create the message body
    $messageBody = @{
        "@type" = "MessageCard"
        "@context" = "http://schema.org/extensions"
        "summary" = "ETG Desktop Assistant Feedback"
        "sections" = @(
            @{
                "text" = ("Desktop Assistant Feedback: $($UserFeedback)$($FeedbackUserName)`r`n")
            }
        )
    } | ConvertTo-Json

    # Post the message to Teams
    Invoke-RestMethod -Uri $TestingWebhookURL -Method Post -Body $messageBody -ContentType "application/json" | Out-Null
    } -ArgumentList $TestingWebhookURL, $UserFeedback, $FeedbackUserName, $TimestampFunction
    $OutTextControl.AppendText("$((& $TimestampFunction)) - Teams message sent for feedback.`r`n")
} catch {
    $OutTextControl.AppendText("$((& $TimestampFunction)) - An error occurred: $($_.Exception.Message)`r`n")
    throw $errorMessage
}