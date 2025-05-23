param(
    [string]$Environment
)
# The above parameters are passed in from the AddLendertoLFPServices.ps1 script

# Get the directory of the currently executing script
$ScriptDir = $PSScriptRoot

# Build the path to the JSON file
$ConfigFilePath = "..\MainGUI\Config.json"

# Combine them to get the full path to the JSON file
$FullJsonFilePath = Join-Path -Path $ScriptDir -ChildPath $ConfigFilePath

# Get config values from json file
$ConfigValues = Get-Content -Path $FullJsonFilePath | ConvertFrom-Json
$BillingRestartsJson = Join-Path -Path $PSScriptRoot -ChildPath $ConfigValues.BillingRestartsEnvironments
$BillingRestartsEnvironments = Get-Content -Path $BillingRestartsJson | ConvertFrom-Json

$EnvironmentData = $null
switch ($Environment) {
    'QA' {
        $EnvironmentData = $BillingRestartsEnvironments.Environments.QA
    }
    'Staging' {
        $EnvironmentData = $BillingRestartsEnvironments.Environments.Staging
    }
    'Production' {
        $EnvironmentData = $BillingRestartsEnvironments.Environments.Prod
    }
    default { throw "Restart Services script: Invalid Environment specified: $Environment" }
}

# Initialize the exit code to 0 (success); this will be updated to 1 (failure) if any of the services fail to stop
$ExitCode = 0

# Function to remotely restart the billing file services
function Restart-RemoteService {
    param(
        [string]$computerName,
        [string]$serviceName,
        $OutText
    )

    $scriptBlock = {
        param($serviceName)
        $result = @()
        # Function to return the current time in EST
        function Get-Timestamp {
            # Get the current time (this will be in the time zone of the system where the script is running)
            $currentTime = Get-Date

            # Define the EST time zone
            $easternTimeZone = [System.TimeZoneInfo]::FindSystemTimeZoneById("Eastern Standard Time")

            # Convert the current time to EST
            $currentTimeInEST = [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId($currentTime, $easternTimeZone.Id)

            # Return the time in EST
            return $currentTimeInEST.ToString("yyyy/MM/dd hh:mm:ss")
        }
        Try {
            $service = Get-Service -Name $serviceName
            $service | Restart-Service
            $result += "$(Get-Timestamp) - $serviceName restarted."
        } Catch {
            $result += "$(Get-Timestamp) - $serviceName failed to restart."
            $result += "$(Get-Timestamp) - $_.Exception.Message"
            $script:ExitCode = 1
        }
        return $result -join "`r`n"
    }

    $remoteOutput = Invoke-Command -ComputerName $computerName -ScriptBlock $scriptBlock -ArgumentList $serviceName
    $OutText.AppendText("$remoteOutput`r`n")
}

if ($null -ne $EnvironmentData) {
    foreach ($entry in $EnvironmentData) {
        $computerName = $entry.ComputerName
        $serviceName = $entry.ServiceName
        Restart-RemoteService -computerName $computerName -serviceName $serviceName -OutText $OutText
    }
} else {
    $OutTextControl.AppendText("$((& $TimestampFunction)) - Restart Services script: No environment data found for $Environment.`r`n")
    $ExitCode = 1
}

exit $ExitCode