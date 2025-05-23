param (
    [string]$File,
    [string]$DestinationFile,
    [string]$OutPath,
    [pscredential]$LoginCredentials,
    [string]$Instance,
    [string]$TableName,
    $OutTextControl,
    [scriptblock]$TimestampFunction
)
# The above parameters are passed in from the MainGUI.ps1 script

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
    Invoke-Command -ComputerName $WorkhorseServer -Authentication Credssp -Credential $LoginCredentials -ScriptBlock {
        param(
            $File, 
            $Instance, 
            $TableName,
            $TimestampFunction
        )
        $Database = "HDTStorage"
        foreach($sheet in Get-ExcelSheetInfo $File) {
            $data = Import-Excel -Path $File -WorksheetName $sheet.name | ConvertTo-DbaDataTable
            Write-DbaDataTable -SqlInstance $Instance -Database $Database -InputObject $data -AutoCreateTable -Table $TableName
        }
    } -ArgumentList $DestinationFile, $Instance, $TableName, $TimestampFunction
    $OutTextControl.AppendText("$((& $TimestampFunction)) - HDTStorage.dbo.$($TableName) created.`r`n")
} catch {
    $ErrorMessage = $OutTextControl.AppendText("$((& $TimestampFunction)) - SQL Command failed to execute. Error: $($_.Exception.Message)")
    throw $ErrorMessage
}