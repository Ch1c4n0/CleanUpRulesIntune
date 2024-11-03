Add-Type -AssemblyName PresentationFramework

# Create the main window
$window = New-Object System.Windows.Window
$window.Title = "Device List Clean Up v.1 - Marcelo Goncalves"
$window.Width = 400
$window.Height = 300

# Create a stack panel to hold the buttons
$stackPanel = New-Object System.Windows.Controls.StackPanel
$stackPanel.Orientation = "Vertical"
$stackPanel.HorizontalAlignment = "Center"
$stackPanel.VerticalAlignment = "Center"

# Create the Login button
$loginButton = New-Object System.Windows.Controls.Button
$loginButton.Content = "Login"
$loginButton.Background = [System.Windows.Media.Brushes]::Yellow
$loginButton.Margin = "10"
$loginButton.Add_Click({
    try {
        Import-Module Microsoft.Graph.Beta.DeviceManagement
        Connect-MgGraph -Scopes "DeviceManagementManagedDevices.Read.All" -NoWelcome
        [System.Windows.MessageBox]::Show("Login successful!", "Info", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        $deviceListButton.IsEnabled = $true
    } catch {
        [System.Windows.MessageBox]::Show("Login failed: $_", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
})
$stackPanel.Children.Add($loginButton)

# Create the Device List for Clean Up button
$deviceListButton = New-Object System.Windows.Controls.Button
$deviceListButton.Content = "Device List for Clean Up"
$deviceListButton.Margin = "10"
$deviceListButton.IsEnabled = $false
$deviceListButton.Add_Click({
    $thresholdDate = (Get-Date).AddDays(-10)

    $devices = Get-MgBetaDeviceManagementManagedDevice |
        Where-Object { $_.LastSyncDateTime -lt $thresholdDate } |
        Select-Object LastSyncDateTime, 
                      @{Name="DaysSinceLastSync"; Expression={(Get-Date) - $_.LastSyncDateTime | Select-Object -ExpandProperty Days}},
                      UserPrincipalName, 
                      DeviceName, 
                      Manufacturer, 
                      Model, 
                      OperatingSystem, 
                      OSVersion, 
                      ComplianceState, 
                      ManagementState, 
                      ManagedDeviceOwnerType, 
                      JoinType, 
                      ManagementCertificateExpirationDate |
        Sort-Object LastSyncDateTime

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Device List for Clean Up</title>
    <style>
        table {
            width: 100%;
            border-collapse: collapse;
        }
        th, td {
            border: 1px solid black;
            padding: 8px;
            text-align: left;
        }
        .green {
            background-color: green;
        }
        .yellow {
            background-color: yellow;
        }
        .red {
            background-color: red;
        }
    </style>
</head>
<body>
    <h1>Device List for Clean Up</h1>
    <p>Access Log: Login successful at $(Get-Date)</p>
    <table>
        <tr>
            <th>Last Sync Date Time</th>
            <th>Days Since Last Sync</th>
            <th>User Principal Name</th>
            <th>Device Name</th>
            <th>Manufacturer</th>
            <th>Model</th>
            <th>Operating System</th>
            <th>OS Version</th>
            <th>Compliance State</th>
            <th>Management State</th>
            <th>Managed Device Owner Type</th>
            <th>Join Type</th>
            <th>Management Certificate Expiration Date</th>
        </tr>
"@

    foreach ($device in $devices) {
        $daysSinceLastSync = [int]$device.DaysSinceLastSync
        $class = ""
        if ($daysSinceLastSync -ge 1 -and $daysSinceLastSync -le 20) {
            $class = "green"
        } elseif ($daysSinceLastSync -gt 20 -and $daysSinceLastSync -le 30) {
            $class = "yellow"
        } elseif ($daysSinceLastSync -gt 30) {
            $class = "red"
        }

        $html += @"
        <tr class="$class">
            <td>$($device.LastSyncDateTime)</td>
            <td>$($device.DaysSinceLastSync)</td>
            <td>$($device.UserPrincipalName)</td>
            <td>$($device.DeviceName)</td>
            <td>$($device.Manufacturer)</td>
            <td>$($device.Model)</td>
            <td>$($device.OperatingSystem)</td>
            <td>$($device.OSVersion)</td>
            <td>$($device.ComplianceState)</td>
            <td>$($device.ManagementState)</td>
            <td>$($device.ManagedDeviceOwnerType)</td>
            <td>$($device.JoinType)</td>
            <td>$($device.ManagementCertificateExpirationDate)</td>
        </tr>
"@
    }

    $html += @"
    </table>
</body>
</html>
"@

    $htmlFilePath = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "DeviceListForCleanUp.html")
    $html | Out-File -FilePath $htmlFilePath -Encoding UTF8

    Start-Process $htmlFilePath
})
$stackPanel.Children.Add($deviceListButton)

# Add the stack panel to the main window
$window.Content = $stackPanel

# Show the main window
$window.ShowDialog()