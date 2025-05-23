$ServerCSV = Import-CSV $csvPath
# Create an empty list to store the values
$ServerList = New-Object System.Collections.Generic.List[System.Object]

# Iterate through each column in the CSV
foreach ($column in $ServerCSV[0].PSObject.Properties.Name) {
    # Iterate through each row in the column
    foreach ($row in $ServerCSV) {
        # Check if the value is not null or empty
        if (![string]::IsNullOrEmpty($row.$column)) {
            # Add the value to the list
            $ServerList.Add($row.$column)
        }
    }
}

Write-Host "Server List: $ServerList"

# Get the directory of the currently executing script
$ScriptDir = $PSScriptRoot

# Variables for importing server list from CSV
$csvPath = Join-Path -Path $ScriptDir -ChildPath "..\MainGUI\servers.csv"
$ServerCSV = Import-CSV $csvPath
$csvHeaders = ($ServerCSV | Get-Member -MemberType NoteProperty).name

# ScriptBlock to be executed on each server
$scriptBlock = {
    param($server)

    try {
        # Attempt to import the WebAdministration module
        Import-Module WebAdministration -ErrorAction Stop
    }
    catch {
        # Write the error message and exit the script block
        Write-Output "The server $server failed to import WebAdministration module: $_"
        return
    }

    # Function to get certificate expiration date by thumbprint
    function Get-CertExpirationDate {
        param ([string]$thumbprint)
        $store = New-Object System.Security.Cryptography.X509Certificates.X509Store -ArgumentList "My", "LocalMachine"
        $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadOnly)
        $cert = $store.Certificates | Where-Object { $_.Thumbprint -eq $thumbprint }
        $store.Close()

        if ($null -ne $cert) {
            return $cert.NotAfter
        } else {
            return $null
        }
    }

    $today = Get-Date
    $threshold = $today.AddDays(365)

    # Retrieve all sites and their bindings for each server
    $sites = Get-ChildItem IIS:\Sites

    foreach ($site in $sites) {
        foreach ($binding in $site.bindings.Collection) {
            if ($binding.protocol -eq "https") {
                $thumbprint = $binding.certificateHash
                $expirationDate = Get-CertExpirationDate -thumbprint $thumbprint

                if ($expirationDate -ne $null -and $expirationDate -le $threshold) {
                    Write-Output "Server: $server, Site: $($site.Name), Certificate Expiring Soon: $expirationDate"
                }
            }
        }
    }
}

foreach ($header in $csvHeaders) {
    foreach ($server in $ServerCSV.$header) {
        if ($null -ne $server -and $server -ne "") {
            # Execute the script block on each server
            Invoke-Command -ComputerName $server -ScriptBlock $scriptBlock -ArgumentList $server -Authentication Negotiate
        }
    }
}


# ! LOGIC FOR DISPLAYING THE RESULTS IN TABLE FORMAT
<# # Mock data for demonstration
$certsExpiringThisMonth = @(
    [PsCustomObject]@{ Site = "Site1"; ExpirationDate = "2023-11-20" },
    [PsCustomObject]@{ Site = "Site2"; ExpirationDate = "2023-11-25" }
)

$certsExpiringNext3Months = @(
    [PsCustomObject]@{ Site = "Site3"; ExpirationDate = "2023-12-15" },
    [PsCustomObject]@{ Site = "Site4"; ExpirationDate = "2024-01-10" }
)

$certsExpiringThisYear = @(
    [PsCustomObject]@{ Site = "Site5"; ExpirationDate = "2023-12-30" },
    [PsCustomObject]@{ Site = "Site6"; ExpirationDate = "2023-10-05" } # Example of an expired certificate
)

# Function to display the data in a table format
function Display-CertInfo {
    param($title, $data)
    if ($data.Count -eq 0) {
        Write-Host "$($title): None"
    } else {
        Write-Host $title
        $data | Format-Table -AutoSize
    }
}

# Display the collected data
Display-CertInfo "Certs expiring this month" $certsExpiringThisMonth
Display-CertInfo "Certs expiring in the next 3 months" $certsExpiringNext3Months
Display-CertInfo "Certs expiring this year" $certsExpiringThisYear #>