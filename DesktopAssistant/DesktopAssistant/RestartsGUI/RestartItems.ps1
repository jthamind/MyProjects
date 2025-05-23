param (
    [string]$Action,
    [string]$SelectedServer,
    [string]$SelectedTab,
    [array]$SelectedItems,
    $OutTextControl,
    [scriptblock]$TimestampFunction
)

if ($SelectedItems.Count -gt 2) {
    $itemList = $SelectedItems[0..($SelectedItems.Count - 2)] -join ', '
    $itemList += ", and $($SelectedItems[-1])"
} elseif ($SelectedItems.Count -eq 2) {
    $itemList = "$($SelectedItems[0]) and $($SelectedItems[1])"
} else {
    $itemList = $SelectedItems[0]
}

# Adjusted action string for display in present continuous tense
$DisplayActionPresentTense = if ($Action -eq 'Stop') { 'Stopping' } else { "${Action}ing" }
# Adjusted action string for display in past tense
$DisplayActionPastTense = if ($Action -eq 'Stop') { 'Stopped' } else { "${Action}ed" }

# Append the formatted string to the output control
$OutTextControl.AppendText("$((& $TimestampFunction)) - $DisplayActionPresentTense $itemList on $SelectedServer`r`n")

$runspacePool = [runspacefactory]::CreateRunspacePool(1, [Math]::Min(1, $SelectedItems.Count))
$runspacePool.Open()

$scriptblock = {
    param($Action, $SelectedServer, $item, $SelectedTab, $DisplayActionPresentTense, $DisplayActionPastTense)

    # Remote script block to be executed on the selected server
    $remoteScriptBlock = {
        param($Action, $SelectedServer, $item, $SelectedTab, $DisplayActionPresentTense, $DisplayActionPastTense)
        switch ($SelectedTab) {
            'Services' {
                $cmdlet = switch ($Action) {
                    'Stop' { 'Stop-Service' }
                    'Start' { 'Start-Service' }
                    'Restart' { 'Restart-Service' }
                }
                $DisplayName = $item
                $ServiceName = (Get-WmiObject -Class Win32_Service | Where-Object { $_.DisplayName -eq $DisplayName }).Name
                $DisabledCheck = Get-WmiObject -Class Win32_Service -Filter "Name='$ServiceName'"
                if ($DisabledCheck.StartMode -eq 'Disabled') {
                        return "$item is disabled and cannot be $DisplayActionPastTense"
                    } else {
                        try {
                            & $cmdlet -Name $ServiceName
                            return "Successfully $DisplayActionPastTense $item."
                        } catch {
                            return "Failed to $DisplayActionPresentTense $($item): $($_.Exception.Message)"
                        }
                }
            }
            'IIS Sites' {
                Import-Module WebAdministration
                $cmdlet = switch ($Action) {
                    'Stop' { 'Stop-WebSite' }
                    'Start' { 'Start-WebSite' }
                    'Restart' { 'Restart-WebSite' }
                }
                try {
                    $SiteStatus = Get-WebSite -Name $item | Select-Object -ExpandProperty State
                    if ($SiteStatus -eq 'Stopped' -and $Action -eq 'Restart') {
                        Start-WebSite -Name $item
                        return "$item was already stopped. Started $item."
                    } else {
                        & $cmdlet -Name $item
                        return "Successfully $DisplayActionPastTense $item."
                    }
                } catch {
                    return "Failed to $DisplayActionPresentTense $($item): $($_.Exception.Message)"
                }
            }
            'App Pools' {
                Import-Module WebAdministration
                $cmdlet = switch ($Action) {
                    'Stop' { 'Stop-WebAppPool' }
                    'Start' { 'Start-WebAppPool' }
                    'Restart' { 'Restart-WebAppPool' }
                }
                try {
                    $AppPoolStatus = (Get-WebAppPoolState -Name $item).Value
                    if ($AppPoolStatus -eq 'Stopped' -and $Action -eq 'Restart') {
                        Start-WebAppPool -Name $item
                        return "$item was already stopped. Started $item."
                    } else {
                        & $cmdlet -Name $item
                        return "Successfully $DisplayActionPastTense $item."
                    }
                } catch {
                    return "Failed to $DisplayActionPresentTense $($item): $($_.Exception.Message)"
                }
            }
        }
    }

    # Invoke the command on the selected server
    $result = Invoke-Command -ComputerName $SelectedServer -ScriptBlock $remoteScriptBlock -Authentication Negotiate -ArgumentList $Action, $SelectedServer, $item, $SelectedTab, $DisplayActionPresentTense, $DisplayActionPastTense

    return $result
}

$runspaces = @()

foreach ($item in $SelectedItems) {
    $runspace = [powershell]::Create().AddScript($scriptblock).AddArgument($Action).AddArgument($SelectedServer).AddArgument($item).AddArgument($SelectedTab).AddArgument($DisplayActionPresentTense).AddArgument($DisplayActionPastTense)
    $runspace.RunspacePool = $runspacePool
    $runspaces += [PSCustomObject]@{
        Runspace = $runspace
        PowerShell = $runspace.BeginInvoke()
    }
}

# Collecting results after all runspaces have completed
foreach ($r in $runspaces) {
    $result = $r.Runspace.EndInvoke($r.PowerShell)
    $r.Runspace.Dispose()
    foreach ($msg in $result) {
        $OutText.AppendText("$((& $TimestampFunction)) - $msg`r`n")
    }
}

$runspacePool.Close()
$runspacePool.Dispose()