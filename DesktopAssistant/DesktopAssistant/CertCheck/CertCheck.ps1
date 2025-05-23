param (
    [string]$UserChosenServers,
    [string]$OutputFilePath,
    [string]$IgnoreFailedServers,
    $OutTextControl,
    [scriptblock]$TimestampFunction
)

# Import CSV
$FullCsvFilePath = Join-Path -Path (Resolve-Path -Path "$PSScriptRoot\..\MainGUI") -ChildPath "servers.csv"
$ServerCSV = Import-Csv -Path $FullCsvFilePath

$ServerList = New-Object System.Collections.Generic.List[System.Object]

if ($null -ne $UserChosenServers -and $UserChosenServers -ne '') {
    $UserChosenServers -split '\r?\n' | ForEach-Object {
        if (![string]::IsNullOrEmpty($_)) {
            $ServerList.Add($_.Trim())
        }
    }
} elseif (@($ServerCSV).Count -gt 0) {
	foreach ($row in $ServerCSV) {
		# Process each column in the row
		foreach ($columnName in $row.PSObject.Properties.Name) {
			# Get the server name from the column
			$serverName = $row.$columnName
			# Add non-empty server names to the list
			if (![string]::IsNullOrEmpty($serverName)) {
				$ServerList.Add($serverName.Trim())
			}
		}
	}
} else {
    $OutTextControl.AppendText("$((& $TimestampFunction)) - No server(s) specified. If you used the servers.csv file, make sure it's correctly configured. If you entered server names, make sure they're valid. Exiting Cert Check.`r`n")
    return
}

$runspacePool = [runspacefactory]::CreateRunspacePool(1, [Math]::Min(1, $ServerList.Count))
$runspacePool.Open()

$scriptblock = {
    param($server, $IgnoreFailedServers)
    try {
        $result = Invoke-Command -ComputerName $server -ScriptBlock {
            try {
                Import-Module WebAdministration -ErrorAction Stop
            } catch {
                throw "WebAdministration module not loaded"
            }

            function Get-CertExpirationDate {
                param ([string]$thumbprint)
                $store = New-Object System.Security.Cryptography.X509Certificates.X509Store "My", "LocalMachine"
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
            $sites = Get-ChildItem IIS:\Sites
            $output = @()

            foreach ($site in $sites) {
                foreach ($binding in $site.bindings.Collection) {
                    if ($binding.protocol -eq "https") {
                        $thumbprint = $binding.certificateHash
                        $expirationDate = Get-CertExpirationDate -thumbprint $thumbprint

                        if ($expirationDate -ne $null -and $expirationDate -le $threshold) {
                            $output += "Site: $($site.Name), Certificate Expiring Soon: $expirationDate"
                        }
                    }
                }
            }
            return $output
        } -ArgumentList $server -Authentication Negotiate

        if ($result -eq $null -and $IgnoreFailedServers -eq $false) {
            return "$($server): No response (possible access issue or command execution error)."
        } elseif ($result -match "WebAdministration module not loaded" -and $IgnoreFailedServers -eq $false) {
            return "$($server): The specified module 'WebAdministration' was not loaded."
        } else {
            $criticalResults = $result | Where-Object { $_ -match "Certificate Expiring Soon" }
            if ($criticalResults) {
                $criticalResults | ForEach-Object { "$($server): $_" }
            } else {
                return "$($server): No certificates near expiration threshold."
            }
        }
    } catch {
        if ($IgnoreFailedServers -eq $false) {
            return "$($server): An error occurred during remote command execution: $($_.Exception.Message)"
        }
    }
}

$runspaces = @()

foreach ($server in $ServerList) {
    $runspace = [powershell]::Create().AddScript($scriptblock).AddArgument($server).AddArgument($IgnoreFailedServers)
    $runspace.RunspacePool = $runspacePool
    $runspaces += [PSCustomObject]@{
        Server = $server
        Runspace = $runspace
        PowerShell = $runspace.BeginInvoke()
    }
}

# Collecting results after all runspaces have completed
foreach ($r in $runspaces) {
    $result = $r.Runspace.EndInvoke($r.PowerShell)
    $r.Runspace.Dispose()
    foreach ($msg in $result) {
        if ($null -ne $OutputFilePath -and $OutputFilePath -ne '') {
            "$((& $TimestampFunction)) - $msg`r`n" | Out-File -FilePath $OutputFilePath -Append
        } else {
            $OutText.AppendText("$((& $TimestampFunction)) - $msg`r`n")
        }
    }
}

$runspacePool.Close()
$runspacePool.Dispose()
$OutText.AppendText("$((& $TimestampFunction)) - Cert Check Process complete.`r`n")