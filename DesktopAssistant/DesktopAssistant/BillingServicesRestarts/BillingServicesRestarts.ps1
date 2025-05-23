param(
    [string]$Environment,
    $OutTextControl,
    [scriptblock]$TimestampFunction
)
# The above parameters are passed in from the MainGUI.ps1 script

Add-Type -AssemblyName "System.Data"

# Get the directory of the currently executing script
$ScriptDir = $PSScriptRoot

# Build the path to the JSON file
$ConfigFilePath = "..\MainGUI\Config.json"

# Combine them to get the full path to the JSON file
$FullJsonFilePath = Join-Path -Path $ScriptDir -ChildPath $ConfigFilePath

# Get config values from json file
$ConfigValues = Get-Content -Path $FullJsonFilePath | ConvertFrom-Json
$RestartBillingServicesScriptPath = Join-Path -Path $PSScriptRoot -ChildPath $ConfigValues.RestartBillingServicesScript

$OutTextControl.AppendText("$((& $TimestampFunction)) - Restarting Billing File Services...`r`n")

if (Test-Path $RestartBillingServicesScriptPath) {
    & $RestartBillingServicesScriptPath -Environment $Environment -TimestampFunction $TimestampFunction -OutTextControl $OutTextControl
    if ($LASTEXITCODE -eq 0) {
        $OutTextControl.AppendText("$((& $TimestampFunction)) - Restart Billing File Services script executed successfully.`r`n")
    } else {
        $OutTextControl.AppendText("$((& $TimestampFunction)) - Restart Billing File Services script failed to execute.`r`n")
    }
} else {
    $OutTextControl.AppendText("$((& $TimestampFunction)) - The script $RestartBillingServicesScriptPath does not exist.`r`n")
}