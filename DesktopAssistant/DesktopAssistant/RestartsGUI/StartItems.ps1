param (
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

# Append the formatted string to the output control
$OutTextControl.AppendText("$((& $TimestampFunction)) - Starting $itemList on $SelectedServer`r`n")

$runspacePool = [runspacefactory]::CreateRunspacePool(1, [Math]::Min(1, $SelectedItems.Count))
$runspacePool.Open()

$scriptblock = {
    param($SelectedServer, $item, $SelectedTab, $OutTextControl, $TimestampFunction)
    $result = Invoke-Command -ComputerName $SelectedServer -ScriptBlock {
        param($SelectedServer, $item, $SelectedTab)
        if ($SelectedTab -eq 'Services') {
            Start-Service -Name $item
        } elseif ($SelectedTab -eq 'IIS Sites') {
            Import-Module WebAdministration
            Start-WebSite -Name $item
        } elseif ($SelectedTab -eq 'App Pools') {
            Import-Module WebAdministration
            Start-WebAppPool -Name $item
        }
        return "Successfully started $item"
    } -Authentication Negotiate -ArgumentList $SelectedServer, $item, $SelectedTab

    return $result
}

$runspaces = @()

foreach ($item in $SelectedItems) {
    $runspace = [powershell]::Create().AddScript($scriptblock).AddArgument($SelectedServer).AddArgument($item).AddArgument($SelectedTab).AddArgument($OutTextControl).AddArgument($TimestampFunction)
    $runspace.RunspacePool = $runspacePool
    $runspaces += [PSCustomObject]@{
        Item = $item
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