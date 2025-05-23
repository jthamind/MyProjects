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
$UserProfilePath = [Environment]::GetEnvironmentVariable("USERPROFILE")
$LenderLFPJson = $ConfigValues.LenderLFPEnvironments.Replace("{USERPROFILE}", $UserProfilePath)
$LenderLFPEnvironments = Get-Content -Path $LenderLFPJson | ConvertFrom-Json

$EnvironmentData = $null
switch ($Environment) {
    'QA' {
        $EnvironmentData = $LenderLFPEnvironments.Environments.QA
    }
    'Staging' {
        $EnvironmentData = $LenderLFPEnvironments.Environments.Staging
    }
    'Production' {
        $EnvironmentData = $LenderLFPEnvironments.Environments.Prod
    }
    default { throw "Start Services script: Invalid Environment specified: $Environment" }
}

# Initialize the exit code to 0 (success); this will be updated to 1 (failure) if any of the services fail to stop
$ExitCode = 0

function Start-RemoteService {
    param(
        [string]$computerName,
        [string]$serviceName,
        $OutText
    )

    $scriptBlock = {
        param($serviceName)
        $result = @()
        Try {
            $service = Get-Service -Name $serviceName
            if ($service.Status -eq 'Stopped') {
                $service | Start-Service
            }
            Start-Sleep -Seconds 5
            $service = Get-Service -Name $serviceName
            if ($service.Status -eq 'Running') {
                $result += "$(Get-Date -Format "yyyy/MM/dd hh:mm:ss") - $serviceName started successfully."
            } else {
                $result += "$(Get-Date -Format "yyyy/MM/dd hh:mm:ss") - $serviceName failed to start."
                $script:ExitCode = 1
            }
        } Catch {
            $result += "$(Get-Date -Format "yyyy/MM/dd hh:mm:ss") - $serviceName failed to start."
            $result += "$(Get-Date -Format "yyyy/MM/dd hh:mm:ss") - $_.Exception.Message"
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
        # Call your function to start the service
        Start-RemoteService -computerName $computerName -serviceName $serviceName -OutText $OutText
    }
} else {
    $OutText.AppendText("$(Get-Timestamp) - Start Services script: No environment data found for $Environment.`r`n")
    $ExitCode = 1
}

<# # Call the function for each service you want to start
Start-RemoteService -computerName 'UTQA2-APP2' -serviceName 'LDHServicePRCPA' -OutText $OutText
Start-RemoteService -computerName 'UTQA2-APP4' -serviceName 'MSGSRVREXTINFO' -OutText $OutText
Start-RemoteService -computerName 'UTQA2-APP2' -serviceName 'LDHServiceEDI' -OutText $OutText
Start-RemoteService -computerName 'UTQA2-APP2' -serviceName 'LDHServiceUSD' -OutText $OutText #>

exit $ExitCode