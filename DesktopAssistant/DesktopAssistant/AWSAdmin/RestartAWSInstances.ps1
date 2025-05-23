param (
    [string]$AWSSSOProfile,
	[string]$Action,
	[string]$SelectedInstanceName,
    $OutTextControl,
    [scriptblock]$TimestampFunction
)
# The above parameters are passed in from the main script

# Check if $SelectedInstanceName is an array
if ($SelectedInstanceName -is [array]) {
    if ($SelectedInstanceName.Count -gt 2) {
        $InstanceList = $SelectedInstanceName[0..($SelectedInstanceName.Count - 2)] -join ', '
        $InstanceList += ", and $($SelectedInstanceName[-1])"
    } elseif ($SelectedInstanceName.Count -eq 2) {
        $InstanceList = "$($SelectedInstanceName[0]) and $($SelectedInstanceName[1])"
    } else {
        $InstanceList = $SelectedInstanceName[0]
    }
} else {
    # If $SelectedInstanceName is not an array, it's a single string
    $InstanceList = $SelectedInstanceName
}

# Adjusted action string for display in present continuous tense
$DisplayActionPresentTense = if ($Action -eq 'Stop') { 'Stopping' } else { "${Action}ing" }
# Adjusted action string for display in past tense
$DisplayActionPastTense = if ($Action -eq 'Stop') { 'Stopped' } else { "${Action}ed" }

$OutTextControl.AppendText("$((& $TimestampFunction)) - $DisplayActionPresentTense $InstanceList`r`n")

$runspacePool = [runspacefactory]::CreateRunspacePool(1, [Math]::Min(1, $SelectedInstanceName.Count))
$runspacePool.Open()

$scriptblock = {
    param($Action, $AWSSSOProfile, $Item, $DisplayActionPresentTense, $DisplayActionPastTense)

	$CheckStatusMessage = 'This process may take a minute, so please check the status of the instance for updates.'

    try {
        # Get the instance ID and status as a single string, separated by a delimiter
        $instanceInfoString = aws ec2 describe-instances --region us-east-2 --filters "Name=tag:Name,Values=$Item" --query "Reservations[*].Instances[*].join(':', [InstanceId, State.Name])" --profile $AWSSSOProfile --output text

        # Split the string to extract ID and status
        $splitInstanceInfo = $instanceInfoString -split ':'
        $AWSInstanceId = @{
            Id = $splitInstanceInfo[0]
            Status = $splitInstanceInfo[1]
        }
    } catch {
        return "Failed to retrieve InstanceId and Status for $($Item): $($_.Exception.Message)"
    }

    switch ($Action) {
        'Reboot' {
            if ($AWSInstanceId.Status -eq 'running') {
                try {
                    aws ec2 reboot-instances --region us-east-2 --instance-ids $AWSInstanceId.Id --profile $AWSSSOProfile --output text | Out-Null
                    return "Successfully $DisplayActionPastTense $Item. $CheckStatusMessage"
                } catch {
                    return "Failed to $Action $($Item): $($_.Exception.Message)"
                }
            }
            elseif ($AWSInstanceId.Status -eq 'stopped') {
                try {
                    aws ec2 start-instances --region us-east-2 --instance-ids $AWSInstanceId.Id --profile $AWSSSOProfile --output text | Out-Null
                    return "$Item was already $DisplayActionPastTense. Started $Item. $CheckStatusMessage"
                } catch {
                    return "Failed to $Action $($Item): $($_.Exception.Message)"
                }
            }
            else {
                return "Unable to $Action $Item. Check the instance state as it may be in the process of starting up or shutting down."
            }
        }
        'Start' {
            if ($AWSInstanceId.Status -eq 'running') {
                return "$Item is already running."
            }
            elseif ($AWSInstanceId.Status -eq 'stopped') {
                try {
                    aws ec2 start-instances --region us-east-2 --instance-ids $AWSInstanceId.Id --profile $AWSSSOProfile --output text | Out-Null
                    return "Successfully $DisplayActionPastTense $Item. $CheckStatusMessage"
                } catch {
                    return "Failed to $Action $($Item): $($_.Exception.Message)"
                }
            }
            else {
                return "Unable to $Action $Item. Check the instance state as it may be in the process of starting up or shutting down."
            }
        }
        'Stop' {
            if ($AWSInstanceId.Status -eq 'running') {
                try {
                    aws ec2 stop-instances --region us-east-2 --instance-ids $AWSInstanceId.Id --profile $AWSSSOProfile --output text | Out-Null
                    return "Successfully $DisplayActionPastTense $Item. $CheckStatusMessage"
                } catch {
                    return "Failed to $Action $($Item): $($_.Exception.Message)"
                }
            }
            elseif ($AWSInstanceId.Status -eq 'stopped') {
                return "$Item is already $DisplayActionPastTense."
            }
            else {
                return "Unable to $Action $Item. Check the instance state as it may be in the process of starting up or shutting down."
            }
        }
    }
}

$runspaces = @()

foreach ($Item in $SelectedInstanceName) {
    $runspace = [powershell]::Create().AddScript($scriptblock).AddArgument($Action).AddArgument($AWSSSOProfile).AddArgument($Item).AddArgument($DisplayActionPresentTense).AddArgument($DisplayActionPastTense)
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