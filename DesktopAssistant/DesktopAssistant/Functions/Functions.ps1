# Function to populate the list boxes based on the selected server
function PopulateListBox {
    $SelectedTab = $RestartsTabControl.SelectedTab.Text
    $synchash.SelectedTab = $SelectedTab

    # Get the selected server from the $ServersListBox
    $SelectedServer = $ServersListBox.SelectedItem
    $OutText.AppendText("Selected Server is: $SelectedServer`r`n")

    if ($null -eq $SelectedServer) {
        return
    }

    switch ($synchash.SelectedTab) {
        "Services" {
            $ServicesListBox.Items.Clear()  # Clear the list box before populating with fresh data
            try {
                $Services = Invoke-Command -ComputerName $SelectedServer -ScriptBlock {
                    Get-Service | ForEach-Object { $_.DisplayName } | Sort-Object
                }
                foreach ($service in $Services) {
                    [void]$ServicesListBox.Items.Add($service)
                }
            } catch {
                $synchash.OutText.AppendText("$(Get-Timestamp) - Error retrieving service: $($_.Exception.Message)`r`n")
            }
        }
        "App Pools" {
            $AppPoolsListBox.Items.Clear()  # Clear the list box before populating with fresh data
            try {
                $AppPools = Invoke-Command -ComputerName $SelectedServer -ScriptBlock {
                    Import-Module WebAdministration
                    Get-IISAppPool | ForEach-Object { $_.Name } | Sort-Object
                }
                if ($null -eq $AppPools) {
                    [void]$AppPoolsListBox.Items.Add("No AppPools found")
                }
                foreach ($apppool in $AppPools) {
                    [void]$AppPoolsListBox.Items.Add($apppool)
                }
            } catch {
                $synchash.OutText.AppendText("$(Get-Timestamp) - Error retrieving AppPools: $($_.Exception.Message)`r`n")
            }
        }
    }
}