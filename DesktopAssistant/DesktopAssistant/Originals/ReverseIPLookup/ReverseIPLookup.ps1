# Function to perform a reverse IP lookup
function ReverseIPLookup {
    param (
        [string]$IPAddress
    )

    try {
        $hostEntry = [System.Net.Dns]::GetHostEntry($IPAddress)
        return $hostEntry.HostName
    }
    catch {
        Write-Host "Failed to perform reverse IP lookup."
        return $null
    }
}

# Example usage
$ip = "172.31.146.23"
$hostname = ReverseIPLookup -IPAddress $ip

if ($hostname) {
    Write-Host "The hostname for IP address $ip is $hostname."
} else {
    Write-Host "Could not find the hostname for IP address $ip."
}