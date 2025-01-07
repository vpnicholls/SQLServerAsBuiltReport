#requires -Module dbatools

# Parameters

param(
    [System.Object]$Server
)

# Gather variable values

If ($server -eq $True) {
    $Instance = Find-DbaInstance -ComputerName $Server
    } else {
    $Instance = Find-DbaInstance -ComputerName $ENV:computername
}

$Script:Index = 1

$header = @"
<style>
    h1 {
        font-family: Arial, Helvetica, sans-serif;
        color: #e68a00;
        font-size: 28px;
    }
    
    h2 {
        font-family: Arial, Helvetica, sans-serif;
        color: #000099;
        font-size: 16px;
    }

    table {
		font-size: 12px;
		border: 0px; 
		font-family: Arial, Helvetica, sans-serif;
	} 
	
    td {
		padding: 4px;
		margin: 0px;
		border: 0;
	}
	
    th {
        background: #395870;
        background: linear-gradient(#49708f, #293f50);
        color: #fff;
        font-size: 11px;
        text-transform: uppercase;
        padding: 10px 15px;
        vertical-align: middle;
	}

    tbody tr:nth-child(even) {
        background: #f0f0f2;
    }

    #CreationDate {
        font-family: Arial, Helvetica, sans-serif;
        color: #ff3300;
        font-size: 12px;
    }

    .RunningStatus {
        color: #008000;
    }

    .StopStatus {
        color: #ff0000;
    }
</style>
"@

# Queries required for script
$QuerySystemDataFiles = @"
SELECT
	DB_NAME([database_id]) AS [Database],
	CASE
		WHEN [type_desc] = 'ROWS' THEN 'Data'
		WHEN [type_desc] = 'LOG' THEN 'Log'
	END AS [Type],
	[Name] AS [Logical Name],
	REPLACE([physical_name], 'Microsoft SQL Server\MSSQL16.MSSQLSERVER\MSSQL', '...') AS [Physical Name],
	[size]/128 AS [Initial Size (MB)],
	CASE
		WHEN [is_percent_growth] = 0 THEN CONVERT(NVARCHAR(10), [growth]/128) + 'MB'
		WHEN [is_percent_growth] = 1 THEN CONVERT(NVARCHAR(10), [growth]) + '%'
	END AS [Growth]
FROM [master].[sys].[master_files]
WHERE 
	DB_NAME([database_id]) IN ('master', 'model', 'msdb', 'distribution', 'tempdb') AND
	[type_desc] = 'ROWS';
"@

$QuerySystemLogFiles = @"
SELECT
	DB_NAME([database_id]) AS [Database],
	CASE
		WHEN [type_desc] = 'ROWS' THEN 'Data'
		WHEN [type_desc] = 'LOG' THEN 'Log'
	END AS [Type],
	[Name] AS [Logical Name],
	REPLACE([physical_name], 'Microsoft SQL Server\MSSQL16.MSSQLSERVER\MSSQL', '...') AS [Physical Name],
	[size]/128 AS [Initial Size (MB)],
	CASE
		WHEN [is_percent_growth] = 0 THEN CONVERT(NVARCHAR(10), [growth]/128) + 'MB'
		WHEN [is_percent_growth] = 1 THEN CONVERT(NVARCHAR(10), [growth]) + '%'
	END AS [Growth]
FROM [master].[sys].[master_files]
WHERE 
	DB_NAME([database_id]) IN ('master', 'model', 'msdb', 'distribution', 'tempdb') AND
	[type_desc] = 'LOG';
"@

#The command below will get the name of the computer. If a value was passed in for the -Server parameter, then that value will be used instead.
if ($Server -eq $True) {
    $ComputerName = "<h1>Computer Name: $Server</h1>"
} else {
    $ComputerName = "<h1>Computer Name: $env:computername</h1>"
}

if ($Instance -eq $null) {
    Write-Error "There are no SQL Server instances on this host. Please try running this against a with a SQL Server instance." -ErrorAction Stop
}

#The command below gathers some basic details about the host server
$HostServerProperties = @(
    (Get-DbaInstanceProperty -SqlInstance $Instance | Select-Object Name, Value | Where-Object {$_.Name -eq "FullyQualifiedNetName"}),
    (Get-DbaInstanceProperty -SqlInstance $Instance | Select-Object Name, Value | Where-Object {$_.Name -eq "HostDistribution"}),
    (Get-DbaInstanceProperty -SqlInstance $Instance | Select-Object Name, Value | Where-Object {$_.Name -eq "HostRelease"})
    (Get-DbaInstanceProperty -SqlInstance $Instance | Select-Object Name, Value | Where-Object {$_.Name -eq "OSVersion"}),
    (Get-DbaInstanceProperty -SqlInstance $Instance | Select-Object Name, Value | Where-Object {$_.Name -eq "PhysicalMemory"})
    (Get-DbaInstanceProperty -SqlInstance $Instance | Select-Object Name, Value | Where-Object {$_.Name -eq "Processors"})
) | ConvertTo-Html ID, Name, Value -Fragment -PreContent "<h2>Host Server Properties</h2>"

#The command below will get the Operating System information, convert the result to HTML code as table and store it to a variable
$OSinfo = Get-CimInstance -Class Win32_OperatingSystem | ConvertTo-Html -As List -Property Version,Caption,BuildNumber,Manufacturer -Fragment -PreContent "<h2>Operating System Information</h2>"

#The command below will get the Processor information, convert the result to HTML code as table and store it to a variable
$ProcessInfo = Get-CimInstance -ClassName Win32_Processor | ConvertTo-Html -As List -Property DeviceID,Name,Caption,MaxClockSpeed,SocketDesignation,Manufacturer -Fragment -PreContent "<h2>Processor Information</h2>"

#The command below will get the BIOS information, convert the result to HTML code as table and store it to a variable
$BiosInfo = Get-CimInstance -ClassName Win32_BIOS | ConvertTo-Html -As List -Property SMBIOSBIOSVersion,Manufacturer,Name,SerialNumber -Fragment -PreContent "<h2>BIOS Information</h2>"

#The command below will get the details of Disk, convert the result to HTML code as table and store it to a variable
$DiskInfo = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType=3" | ConvertTo-Html -Property `
    @{Label="Device ID"; Expression = {$_.DeviceID}},
    @{Label="Drive Type"; Expression = {$_.DriveType}},
    @{Label="Provider Name"; Expression = {$_.ProviderName}},
    @{Label="Volume Name"; Expression = {$_.VolumeName}},
    @{Label="Size"; Expression = {$_.Size/1GB}},
    @{Label="Free Space"; Expression = {$_.FreeSpace/1GB}} `
    -Fragment -PreContent "<h2>Disk Information</h2>"

#The command below will get first 10 services information, convert the result to HTML code as table and store it to a variable
$ServicesInfo = Get-DbaService | ConvertTo-Html -Property `
    @{Label = "Service Name"; Expression = "ServiceName"}, `
    @{Label = "Display Name"; Expression = "DisplayName"}, `
    @{Label = "Service Account"; Expression = "StartName"}, `
    @{Label = "Start Mode"; Expression = "StartMode"}, `
    @{Label = "State"; Expression = "State"} `
    -Fragment -PreContent "<h2>Services Information</h2>"
$ServicesInfo = $ServicesInfo -replace '<td>Running</td>','<td class="RunningStatus">Running</td>' 
$ServicesInfo = $ServicesInfo -replace '<td>Stopped</td>','<td class="StopStatus">Stopped</td>'

#The command below will get instance-level properties of most interest, convert the result to HTML code as table and store it to a variable
$InstanceProperties = @(
    (Get-DbaInstanceProperty -SqlInstance $Instance | Select-Object Name, Value | Where-Object {$_.Name -eq "Edition"}),
    (Get-DbaInstanceProperty -SqlInstance $Instance | Select-Object Name, Value | Where-Object {$_.Name -eq "VersionString"}),
    (Get-DbaInstanceProperty -SqlInstance $Instance | Select-Object Name, Value | Where-Object {$_.Name -eq "Collation"})
    (Get-DbaInstanceProperty -SqlInstance $Instance | Select-Object Name, Value | Where-Object {$_.Name -eq "IsFullTextInstalled"}),
    (Get-DbaInstanceProperty -SqlInstance $Instance | Select-Object Name, Value | Where-Object {$_.Name -eq "RootDirectory"})
    (Get-DbaInstanceProperty -SqlInstance $Instance | Select-Object Name, Value | Where-Object {$_.Name -eq "MasterDBPath"}),
    (Get-DbaInstanceProperty -SqlInstance $Instance | Select-Object Name, Value | Where-Object {$_.Name -eq "MasterDBLogPath"}),
    (Get-DbaInstanceProperty -SqlInstance $Instance | Select-Object Name, Value | Where-Object {$_.Name -eq "DefaultLog"})
    (Get-DbaInstanceProperty -SqlInstance $Instance | Select-Object Name, Value | Where-Object {$_.Name -eq "BackupDirectory"}),
    (Get-DbaInstanceProperty -SqlInstance $Instance | Select-Object Name, Value | Where-Object {$_.Name -eq "LoginMode"})
) | ConvertTo-Html ID, Name, Value -Fragment -PreContent "<h2>SQL Server Instance Properties</h2>"

#The command below will details of the system databases' data files, convert the result to HTML code as table and store it to a variable
$SystemDataFiles = Invoke-DbaQuery -SqlInstance $Instance -Query $QuerySystemDataFiles | ConvertTo-Html -Property `
    Database, Type, "Logical Name", "Physical Name", "Initial Size (MB)", Growth `
    -Fragment -PreContent "<h2>System Data Files</h2>"

#The command below will details of the system databases' data files, convert the result to HTML code as table and store it to a variable
$SystemLogFiles = Invoke-DbaQuery -SqlInstance $Instance -Query $QuerySystemDataFiles | ConvertTo-Html -Property `
    Database, Type, "Logical Name", "Physical Name", "Initial Size (MB)", Growth `
    -Fragment -PreContent "<h2>System Log Files</h2>"
  
#The command below will combine all the information gathered into a single HTML report
$Report = ConvertTo-HTML -Body "$ComputerName $HostServerProperties $OSinfo $ProcessInfo $BiosInfo $DiskInfo $ServicesInfo $InstanceProperties $SystemDataFiles $SystemLogFiles" `
    -Title "Computer Information" -Head $header -PostContent "<p id='CreationDate'>Creation Date: $(Get-Date); Report Version: $ReportVersion;<p>"

#The command below will generate the report to an HTML file
$Report | Out-File .\AsBuiltReport.html
