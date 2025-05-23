try {
    Invoke-Command -ComputerName 'CP-WEBSVCPRD-01' -ScriptBlock {
        $IISSites = Get-IISSite | ForEach-Object { $_.Name } | Sort-Object
        return $IISSites
    }
    foreach ($item in $IISSites) {
        Write-Host "$item"
    }
} catch {
    $OutText.AppendText("Error retrieving IIS Sites: $($_.Exception.Message)`r`n")
}