param(
    [string]$LenderId,
    [string]$TicketNumber,
    [string]$Environment,
    $OutTextControl,
    [scriptblock]$TimestampFunction
)
# The above parameters are passed in from the MainGUI.ps1 script

Add-Type -AssemblyName "System.Data"

$SQLInstance = switch ($Environment) {
    'QA' {
        'UTQA-SQL-14'
    }
    'Staging' {
        'UT-SQLSTG-01'
    }
    'Production' {
        'UT-PRD-LISTENER'
    }
    default { throw "Invalid Environment specified: $Environment" }
}

$OriginalTicketNumber = $TicketNumber

# Check and remove any dashes in the ticket number
if ($OriginalTicketNumber -match "-") {
    $UpdatedTicketNumber = $OriginalTicketNumber -replace "-", ""
}

# Get the directory of the currently executing script
$ScriptDir = $PSScriptRoot

# Build the path to the JSON file
$ConfigFilePath = "..\MainGUI\Config.json"

# Combine them to get the full path to the JSON file
$FullJsonFilePath = Join-Path -Path $ScriptDir -ChildPath $ConfigFilePath

# Get config values from json file
$ConfigValues = Get-Content -Path $FullJsonFilePath | ConvertFrom-Json
$StopScriptPath = Join-Path -Path $PSScriptRoot -ChildPath $ConfigValues.StopLFPServicesScript
$StartScriptPath = Join-Path -Path $PSScriptRoot -ChildPath $ConfigValues.StartLFPServicesScript
$AddLenderToLFPServicesSQLScript = Join-Path -Path $PSScriptRoot -ChildPath $ConfigValues.AddLenderToLFPServicesSQL
$TicketsPath = Join-Path -Path $PSScriptRoot -ChildPath $ConfigValues.TicketsPath
$ResolvedTicketsPath = Resolve-Path $TicketsPath

$OutTextControl.AppendText("$((& $TimestampFunction)) - Stopping LFP services...`r`n")

if (Test-Path $StopScriptPath) {
    & $StopScriptPath -Environment $Environment
    if ($LASTEXITCODE -eq 0) {
        $OutTextControl.AppendText("$((& $TimestampFunction)) - Stop Lender Services script executed successfully.`r`n")
    } else {
        $OutTextControl.AppendText("$((& $TimestampFunction)) - Stop Lender Services script failed to execute.`r`n")
    }
} else {
    $OutTextControl.AppendText("$((& $TimestampFunction)) - The script $StopScriptPath does not exist.`r`n")
}

$OutTextControl.AppendText("$((& $TimestampFunction)) - Waiting for 10 seconds before continuing...`r`n")

Start-Sleep -Seconds 10

# ! START CALL LFP SQL FILE

# Read the LFP SQL template file
$LFPSQLFileContent = Get-Content -Path $AddLenderToLFPServicesSQLScript -Raw

# Replace the default values (CSH12345 for ticket and 6969 [nice] for lender ID)
$ReplacedLFPSQLContent = $LFPSQLFileContent -replace [regex]::Escape("CSH12345"), $UpdatedTicketNumber -replace [regex]::Escape("6969"), $LenderId

# Construct the target directory path
$TargetDirectory = "$ResolvedTicketsPath\Active\$($OriginalTicketNumber)"

# Resolves the target directory path to an absolute path
$ResolvedTargetDirectory = Resolve-Path $TargetDirectory

# Set new file path for the modified SQL file
$UpdatedSQLFilePath = "$ResolvedTargetDirectory\AddLenderToLFP$($OriginalTicketNumber).sql"

# Write the modified content back to the file
Set-Content -Path $UpdatedSQLFilePath -Value $ReplacedLFPSQLContent

# Initialize the connection string
# Integrated Security=True means use Windows Authentication
$ConnectionString = "Data Source=$SQLInstance;Initial Catalog=UniTrac;Integrated Security=True"

# Initialize the query
$Query = Get-Content -Path $UpdatedSQLFilePath | Out-String

# Initialize the SQL connection
$Connection = New-Object System.Data.SqlClient.SqlConnection
$Connection.ConnectionString = $ConnectionString

# Event handler for capturing SQL Server messages printed in the .SQL file
$Connection.add_InfoMessage({
    $OutTextControl.AppendText("$((& $TimestampFunction)) - SQL Server Message: " + $_.Message + "`r`n")
})

# Open the SQL connection
$Connection.Open()

# Initialize the SQL command
$Command = $Connection.CreateCommand()
$Command.CommandText = $Query

# Execute the SQL command
$Command.ExecuteNonQuery() | Out-Null

# Close the SQL connection
$Connection.Close()

# ! END CALL LFP SQL FILE

$OutTextControl.AppendText("$((& $TimestampFunction)) - Continuing with service restarts...`r`n")

if (Test-Path $StartScriptPath) {
    & $StartScriptPath -Environment $Environment
    if ($LASTEXITCODE -eq 0) {
        $OutTextControl.AppendText("$((& $TimestampFunction)) - Start Lender Services script executed successfully.`r`n")
    } else {
        $OutTextControl.AppendText("$((& $TimestampFunction)) - Start Lender Services script failed to execute.`r`n")
    }
} else {
    $OutTextControl.AppendText("$((& $TimestampFunction)) - The script $StartScriptPath does not exist.`r`n")
}

$OutTextControl.AppendText("$((& $TimestampFunction)) - Restarted LFP services.`r`n")
$OutTextControl.AppendText("$((& $TimestampFunction)) - Lender $LenderId successfully added to $Environment`r`n")