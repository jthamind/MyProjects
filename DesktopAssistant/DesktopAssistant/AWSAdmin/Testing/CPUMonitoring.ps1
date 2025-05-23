# Load required assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms.DataVisualization

# Create a form to hold the chart
$AWSCPUMetricsForm = New-Object Windows.Forms.Form
$AWSCPUMetricsForm.Text = 'CPU Utilization Graph'
$AWSCPUMetricsForm.Width = 880
$AWSCPUMetricsForm.Height = 630

# Create the chart for AWS CPU metrics
$AWSCPUMetricsChart = New-Object System.Windows.Forms.DataVisualization.Charting.Chart
$AWSCPUMetricsChart.Width = 900
$AWSCPUMetricsChart.Height = 610
$AWSCPUMetricsChart.BackColor = "#230A04"

# ! Need to add a label above the chart

# Create and configure chart area
$AWSCPUMetricsChartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
$AWSCPUMetricsChartArea.AxisX.Title = "Time" # Use a variable that shows the date/time range
$AWSCPUMetricsChartArea.AxisX.TitleForeColor = "#63CFDA"
$AWSCPUMetricsChartArea.AxisX.LabelStyle.ForeColor = "#63CFDA"
$AWSCPUMetricsChartArea.AxisX.LineColor = "#63CFDA"
$AWSCPUMetricsChartArea.AxisY.Title = "CPU %"
$AWSCPUMetricsChartArea.AxisY.TitleForeColor = "#63CFDA"
$AWSCPUMetricsChartArea.AxisY.LabelStyle.ForeColor = "#63CFDA"
$AWSCPUMetricsChartArea.AxisY.LineColor = "#63CFDA"
$AWSCPUMetricsChartArea.BackColor = "#230A04"
$AWSCPUMetricsChart.ChartAreas.Add($AWSCPUMetricsChartArea)

# Create and add data to the series
$AWSCPUMetricsSeries = New-Object System.Windows.Forms.DataVisualization.Charting.Series
$AWSCPUMetricsSeries.ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Line
$AWSCPUMetricsSeries.MarkerStyle = [System.Windows.Forms.DataVisualization.Charting.MarkerStyle]::Circle
$AWSCPUMetricsSeries.MarkerSize = 5
$AWSCPUMetricsSeries.BorderWidth = 2
$AWSCPUMetricsSeries.Color = "#CE450D"

$AWSSSOProfile = 'DesktopAssistant'

$InstanceId = aws ec2 describe-instances --region us-east-2 --filters "Name=tag:Name,Values=CP-WEBDEV-01" --query "Reservations[*].Instances[*].InstanceId" --profile $AWSSSOProfile --output text

# Convert the string to a datetime object
$StartTime = [datetime]::ParseExact('2024-01-14T00:00:00', 'yyyy-MM-ddTHH:mm:ss', $null)
$EndTime = [datetime]::ParseExact('2024-01-15T00:00:00', 'yyyy-MM-ddTHH:mm:ss', $null)

# Add 5 hours to each
$StartTime = $StartTime.AddHours(5)
$EndTime = $EndTime.AddHours(5)

# Parse the AWS CLI Output
$jsonResponse = aws cloudwatch get-metric-statistics --metric-name CPUUtilization --start-time $StartTime --end-time $EndTime --period 3600 --namespace AWS/EC2 --statistics Maximum --dimensions "Name=InstanceId,Value=$InstanceId" --profile $AWSSSOProfile
$awsData = $jsonResponse | ConvertFrom-Json

# Prepare Data for Charting
$cpuUtilizationDataPoints = @()
foreach ($datapoint in $awsData.Datapoints) {
    # Create a hashtable for each datapoint with timestamp and maximum value
    $cpuUtilizationDataPoints += @{
        Timestamp = [datetime]$datapoint.Timestamp
        Maximum = $datapoint.Maximum
    }
}

# Sort the data points by the Timestamp
$sortedDataPoints = $cpuUtilizationDataPoints | Sort-Object Timestamp

# Extract the sorted data into separate arrays
$cpuUtilizationData = @($sortedDataPoints | ForEach-Object { $_.Maximum })
$cpuUtilizationTimestamps = @($sortedDataPoints | ForEach-Object { $_.Timestamp })

# Sort the timestamps again after applying the timezone offset
$cpuUtilizationTimestamps = $cpuUtilizationTimestamps | Sort-Object

# Now bind the sorted and adjusted data to the chart series
# Clear existing series points if any
$AWSCPUMetricsSeries.Points.Clear()

foreach ($i in 0..($cpuUtilizationData.Count - 1)) {
    $point = New-Object System.Windows.Forms.DataVisualization.Charting.DataPoint
    $point.SetValueXY($cpuUtilizationTimestamps[$i], $cpuUtilizationData[$i])
    $AWSCPUMetricsSeries.Points.Add($point)
}

# Configure chart area and X-axis to show all points with hour labels
$AWSCPUMetricsChartArea.AxisX.IntervalType = [System.Windows.Forms.DataVisualization.Charting.DateTimeIntervalType]::Hours
$AWSCPUMetricsChartArea.AxisX.Interval = 1  # Set interval to 1 to show every hour
$AWSCPUMetricsChartArea.AxisX.LabelStyle.Format = "HH:mm"  # Format label to show hour and minute

# After adding all points to the series
$AWSCPUMetricsChartArea.AxisX.Minimum = [datetime]::ParseExact('2024-01-14T00:00:00', 'yyyy-MM-ddTHH:mm:ss', $null).ToOADate()
$AWSCPUMetricsChartArea.AxisX.Maximum = [datetime]::ParseExact('2024-01-15T00:00:00', 'yyyy-MM-ddTHH:mm:ss', $null).ToOADate()

$AWSCPUMetricsChart.Series.Add($AWSCPUMetricsSeries)

# Set the chart area's AxisX LabelStyle format to show labels in HH:mm format
$AWSCPUMetricsChart.ChartAreas[0].AxisX.LabelStyle.Format = "HH:mm"

# Add the chart to the form
$AWSCPUMetricsForm.Controls.Add($AWSCPUMetricsChart)

# Display the form
$AWSCPUMetricsForm.ShowDialog()