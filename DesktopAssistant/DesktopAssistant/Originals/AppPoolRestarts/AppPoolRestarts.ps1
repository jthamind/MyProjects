try {
    Invoke-Command -ComputerName 'CP-WEBSVCPRD-01' -ScriptBlock {
        $AppPools = Get-IISAppPool | ForEach-Object { $_.Name } | Sort-Object
        return $AppPools
    }
    foreach ($item in $AppPools) {
        Write-Host "$item"
    }
} catch {
    $OutText.AppendText("Error retrieving AppPools: $($_.Exception.Message)`r`n")
}