param (
    [string]$AWSSSOProfile,
	[string]$SelectedInstanceName,
    [string]$StartTime,
    [string]$EndTime,
    [int]$PollingPeriod,
    [string]$BackColor,
	[string]$ForeColor,
	[string]$AccentColor,
    $OutTextControl,
    [scriptblock]$TimestampFunction
)
# The above parameters are passed in from the main script

# Load required assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms.DataVisualization

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

$OutTextControl.AppendText("$((& $TimestampFunction)) - Gathering CPU metrics for $InstanceList...`r`n")

$runspacePool = [runspacefactory]::CreateRunspacePool(1, [Math]::Min(1, $SelectedInstanceName.Count))
$runspacePool.Open()

$scriptblock = {
    param($AWSSSOProfile, $Item, $StartTime, $EndTime, $PollingPeriod, $BackColor, $ForeColor, $AccentColor)

    try {
        $InstanceId = aws ec2 describe-instances --region us-east-2 --filters "Name=tag:Name,Values=$Item" --query "Reservations[*].Instances[*].InstanceId" --profile $AWSSSOProfile --output text
    } catch {
        return "Failed to retrieve InstanceId for $($Item): $($_.Exception.Message)"
    }

    # Convert the string to a datetime object
    $StartTime = [datetime]::ParseExact($StartTime, 'yyyy-MM-ddTHH:mm:ss', $null)
    $EndTime = [datetime]::ParseExact($EndTime, 'yyyy-MM-ddTHH:mm:ss', $null)

    # Add 5 hours to each
    $StartTime = $StartTime.AddHours(5)
    $EndTime = $EndTime.AddHours(5)

    try {
        $jsonResponse = aws cloudwatch get-metric-statistics --metric-name CPUUtilization --start-time $StartTime --end-time $EndTime --period $PollingPeriod --namespace AWS/EC2 --statistics Maximum --dimensions "Name=InstanceId,Value=$InstanceId" --profile $AWSSSOProfile
    } catch {
        return "Failed to retrieve CPU metrics for $($Item): $($_.Exception.Message)"
    }

    # Parse the AWS CLI Output
    $awsData = $jsonResponse | ConvertFrom-Json

    # Prepare Data for Charting
    $cpuUtilizationDataPoints = @()
    foreach ($Datapoint in $awsData.Datapoints) {
        # Create a hashtable for each datapoint with timestamp and maximum value
        $cpuUtilizationDataPoints += @{
            Timestamp = [datetime]$Datapoint.Timestamp
            Maximum = $Datapoint.Maximum
        }
    }

    # Sort the data points by the Timestamp
    $sortedDataPoints = $cpuUtilizationDataPoints | Sort-Object Timestamp

    # Extract the sorted data into separate arrays
    $cpuUtilizationData = @($sortedDataPoints | ForEach-Object { $_.Maximum })
    $cpuUtilizationTimestamps = @($sortedDataPoints | ForEach-Object { $_.Timestamp })

    # Sort the timestamps again after applying the timezone offset
    $cpuUtilizationTimestamps = $cpuUtilizationTimestamps | Sort-Object

    # Create a form to hold the chart
    $AWSCPUMetricsForm = New-Object Windows.Forms.Form
    $AWSCPUMetricsForm.Text = 'CPU Utilization Graph'
    $AWSCPUMetricsForm.Width = 900
    $AWSCPUMetricsForm.Height = 630
    $AWSCPUMetricsForm.MaximizeBox = $false

    # Create the chart for AWS CPU metrics
    $AWSCPUMetricsChart = New-Object System.Windows.Forms.DataVisualization.Charting.Chart
    $AWSCPUMetricsChart.Width = 900
    $AWSCPUMetricsChart.Height = 600
    $AWSCPUMetricsChart.BackColor = $BackColor

    # Create and configure chart area
    $AWSCPUMetricsChartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
    $AWSCPUMetricsChartArea.AxisX.Title = "Time" # Use a variable that shows the date/time range
    $AWSCPUMetricsChartArea.AxisX.TitleForeColor = $AccentColor
    $AWSCPUMetricsChartArea.AxisX.LabelStyle.ForeColor = $AccentColor
    $AWSCPUMetricsChartArea.AxisX.LineColor = $AccentColor
    $AWSCPUMetricsChartArea.AxisY.Title = "CPU %"
    $AWSCPUMetricsChartArea.AxisY.TitleForeColor = $AccentColor
    $AWSCPUMetricsChartArea.AxisY.LabelStyle.ForeColor = $AccentColor
    $AWSCPUMetricsChartArea.AxisY.LineColor = $AccentColor
    $AWSCPUMetricsChartArea.BackColor = $BackColor
    $AWSCPUMetricsChart.ChartAreas.Add($AWSCPUMetricsChartArea)

    # Set the minimum and maximum for the Y axis to 0 and 100
    $AWSCPUMetricsChartArea.AxisY.Minimum = 0
    $AWSCPUMetricsChartArea.AxisY.Maximum = 100

    # Create and add data to the series
    $series = New-Object System.Windows.Forms.DataVisualization.Charting.Series
    $series.ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Line
    $series.MarkerStyle = [System.Windows.Forms.DataVisualization.Charting.MarkerStyle]::Circle
    $series.MarkerSize = 5
    $series.BorderWidth = 2
    $series.Color = $ForeColor

    # Clear existing series points if any
    $series.Points.Clear()

    # Bind the sorted and adjusted data to the chart series
    foreach ($i in 0..($cpuUtilizationData.Count - 1)) {
        $point = New-Object System.Windows.Forms.DataVisualization.Charting.DataPoint
        $point.SetValueXY($cpuUtilizationTimestamps[$i], $cpuUtilizationData[$i])
        $series.Points.Add($point)
    }

    # Configure chart area and X-axis to show all points with hour labels
    $AWSCPUMetricsChartArea.AxisX.IntervalType = [System.Windows.Forms.DataVisualization.Charting.DateTimeIntervalType]::Hours
    $AWSCPUMetricsChartArea.AxisX.Interval = 1  # Set interval to 1 to show every hour
    $AWSCPUMetricsChartArea.AxisX.LabelStyle.Format = "HH:mm"  # Format label to show hour and minute

    # Find the minimum and maximum timestamps from the data
    $minTimestamp = $cpuUtilizationTimestamps | Measure-Object -Minimum | Select-Object -ExpandProperty Minimum
    $maxTimestamp = $cpuUtilizationTimestamps | Measure-Object -Maximum | Select-Object -ExpandProperty Maximum

    # Set the minimum and maximum values for the AxisX using the actual data range
    $AWSCPUMetricsChartArea.AxisX.Minimum = $minTimestamp.ToOADate()
    $AWSCPUMetricsChartArea.AxisX.Maximum = $maxTimestamp.ToOADate()

    # Set the IntervalType based on the range you have
    $intervalHours = ($maxTimestamp - $minTimestamp).TotalHours
    $intervalType = [System.Windows.Forms.DataVisualization.Charting.DateTimeIntervalType]::Auto

    # Based on the intervalHours, set the interval type and interval value to ensure proper labeling and readability
    if ($intervalHours -le 1) {
        $intervalType = [System.Windows.Forms.DataVisualization.Charting.DateTimeIntervalType]::Minutes
        $AWSCPUMetricsChartArea.AxisX.Interval = 5 # example for a 1-hour range, adjust as needed
    } elseif ($intervalHours -le 6) {
        $intervalType = [System.Windows.Forms.DataVisualization.Charting.DateTimeIntervalType]::Minutes
        $AWSCPUMetricsChartArea.AxisX.Interval = 30 # example for a 6-hour range, adjust as needed
    } elseif ($intervalHours -le 12) {
        $intervalType = [System.Windows.Forms.DataVisualization.Charting.DateTimeIntervalType]::Hours
        $AWSCPUMetricsChartArea.AxisX.Interval = 1 # example for a 12-hour range, adjust as needed
    } else {
        $intervalType = [System.Windows.Forms.DataVisualization.Charting.DateTimeIntervalType]::Hours
        $AWSCPUMetricsChartArea.AxisX.Interval = 2 # default for more than 12-hour range, adjust as needed
    }

    $AWSCPUMetricsChartArea.AxisX.IntervalType = $intervalType


    $AWSCPUMetricsChart.Series.Add($series)

    # Set the chart area's AxisX LabelStyle format to show labels in HH:mm format
    $AWSCPUMetricsChart.ChartAreas[0].AxisX.LabelStyle.Format = "HH:mm"

    # Add the chart to the form
    $AWSCPUMetricsForm.Controls.Add($AWSCPUMetricsChart)

    # Display the form
    $AWSCPUMetricsForm.ShowDialog() | Out-Null

}

$runspaces = @()

foreach ($Item in $SelectedInstanceName) {
    $runspace = [powershell]::Create().AddScript($scriptblock).AddArgument($AWSSSOProfile).AddArgument($Item).AddArgument($StartTime).AddArgument($EndTime).AddArgument($PollingPeriod).AddArgument($BackColor).AddArgument($ForeColor).AddArgument($AccentColor)
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