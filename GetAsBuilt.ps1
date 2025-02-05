<#
.SYNOPSIS
    Generates an "As Built" document for each SQL Server instance provided.

.DESCRIPTION
    This script compiles detailed information about one or more SQL Server instances into individual structured documents, 
    which can be used for documentation, auditing, or migration planning. It uses PowerShell and dbatools to query SQL Server 
    and constructs documents in markdown or another specified format.

.PARAMETER SQLServerInstances
    An array of SQL Server instance names or connection strings to generate documents for.

.PARAMETER OutputPath
    The directory where the generated documents will be saved.

.PARAMETER DocumentFormat
    The format of the document to be generated. Currently supports 'markdown'. 

.PARAMETER ScriptEventLogPath
    The directory where the log file will be saved.

.NOTES
    Author: Vaughan Nicholls
    Date:   January 07, 2025

.LINK
    https://github.com/vpnicholls/SQLServerAsBuiltReport
#>

#################
### The Setup ###
#################

param (
    [Parameter(Mandatory=$true)]
    [string[]]$SQLServerInstances,
    [string]$OutputPath = "$env:USERPROFILE\Documents\AsBuiltDocs",
    [string]$ScriptEventLogPath = "$($OutputPath)\Logs",
    [string]$DocumentFormat = "markdown"

)

# Generate log file name with datetime stamp
$logFileName = Join-Path -Path $ScriptEventLogPath -ChildPath "SQLAsBuiltDocLog_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"

# Define the function to write to the log file
function Write-Log {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Message,

        [Parameter(Mandatory=$false)]
        [ValidateSet("INFO", "WARNING", "ERROR", "SUCCESS", "DEBUG", "VERBOSE", "FATAL")]
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp [$Level] $Message" | Out-File -FilePath $logFileName -Append
}

# Import SQL credentials
if (-not $HasDomainAccount) {
    try {
        $SqlCredential = Import-Clixml -Path ".\Credentials\SqlCredentials.xml"
    }
    catch {
        Write-Log -Message "Failed to import SQL credentials from .\Credentials\SqlCredentials.xml: $_" -Level ERROR
        throw "SQL Credential import failed. Please ensure the credentials file is present and accessible."
    }
}

# Collect Windows credentials for each host server if not using domain account
$ServerCredentials = @{}
if (-not $HasDomainAccount) {
    foreach ($instance in $SQLServerInstances) {
        # Extract server name from instance name
        $serverName = if ($instance.Contains('\')) {
            $instance.Split('\')[0]
        } else {
            $instance
        }

        try {
            $ServerCredentials[$serverName] = Import-Clixml -Path ".\Credentials\$serverName.xml"
            Write-Log -Message "Successfully imported Windows credentials for $serverName." -Level "INFO"
        }
        catch {
            Write-Log -Message "Failed to import Windows credentials for $serverName from .\Credentials\$serverName.xml: $_" -Level "ERROR"
            # Instead of throwing an error, log it and continue processing other instances if possible
        }
    }
}
else {
    Write-Log -Message "Using domain account for authentication." -Level INFO
}

# Define function get basic properties for host server 
Function Get-HostServerProperties {
    Param (
        [string]$InstanceName
    )

    try {
        Write-Log -Message "Retrieving host server properties for $InstanceName" -Level INFO
        $allProperties = Get-DbaInstanceProperty -SqlInstance $InstanceName
        $HostServerProperties = $allProperties | Where-Object {
            $_.Name -in @("FullyQualifiedNetName", "HostDistribution", "HostRelease", "OSVersion", "PhysicalMemory", "Processors")
        } | Select-Object Name, Value
        return $HostServerProperties
    } catch {
        Write-Log -Message "Failed to retrieve host server properties for $InstanceName. Error: $_" -Level ERROR
        return @()
    }
}

# Define function to get server configuration using dbatools
function Get-SQLServerConfig {
    param (
        [Parameter(Mandatory=$true)]
        [string]$InstanceName
    )

    try {
        Write-Log -Message "Retrieving configuration for $InstanceName" -Level INFO
        $serverInfo = Get-DbaInstanceProperty -SqlInstance $InstanceName -ErrorAction Stop
        $config = @{}
        @('VersionString', 'Edition', 'Collation', 'IsClustered') | ForEach-Object {
            $property = $_
            $value = ($serverInfo | Where-Object { $_.Name -eq $property } | Select-Object -ExpandProperty Value -ErrorAction SilentlyContinue)
            if ($value) {
                $config[$property] = $value
            } else {
                Write-Log -Message "Property $property not found for $InstanceName" -Level WARNING
            }
        }
        return $config
    }
    catch {
        Write-Log -Message "Failed to retrieve configuration for $InstanceName. Error: $_" -Level ERROR
        return $null
    }
}

# Function to get databases details using dbatools
function Get-SQLDatabases {
    param (
        [Parameter(Mandatory=$true)]
        [string]$InstanceName
    )

    try {
        Write-Log -Message "Retrieving database details for $InstanceName" -Level INFO
        $databases = Get-DbaDatabase -SqlInstance $InstanceName -ErrorAction Stop
        
        $systemDatabases = $databases | Where-Object { $_.IsSystemObject }
        $userDatabases = $databases | Where-Object { -not $_.IsSystemObject }

        # Function to get detailed info for databases
        function Get-DatabaseDetails {
            param ($db)
            $dbFiles = Get-DbaDbFile -SqlInstance $InstanceName -Database $db.Name
            $totalSizeBytes = ($dbFiles | Measure-Object -Property Size -Sum).Sum
            
            return @{
                'Name' = $db.Name
                'Status' = $db.Status
                'RecoveryModel' = $db.RecoveryModel
                'CompatibilityLevel' = $db.Compatibility
                'Owner' = $db.Owner
                'CreationDate' = $db.CreateDate.ToString("yyyy-MM-dd HH:mm:ss")
                'LastBackupDate' = if ($db.LastBackupDate) { $db.LastBackupDate.ToString("yyyy-MM-dd HH:mm:ss") } else { "Never" }
                'AutoClose' = $db.AutoClose
                'AutoShrink' = $db.AutoShrink
                'IsReadCommittedSnapshotOn' = $db.IsReadCommittedSnapshotOn
                'IsAutoCreateStatisticsEnabled' = $db.AutoCreateStatisticsEnabled
                'IsAutoUpdateStatisticsEnabled' = $db.AutoUpdateStatisticsEnabled
            }
        }

        $systemDbInfo = $systemDatabases | ForEach-Object { Get-DatabaseDetails -db $_ }
        $userDbInfo = $userDatabases | ForEach-Object { Get-DatabaseDetails -db $_ }

        return @{
            'SystemDatabases' = $systemDbInfo
            'UserDatabases' = $userDbInfo
        }
    }
    catch {
        Write-Log -Message "Failed to retrieve databases for $InstanceName. Error: $_" -Level ERROR
        return @{
            'SystemDatabases' = @()
            'UserDatabases' = @()
        }
    }
}

# Function to Availability Group listeners' and databases' details
function Add-AgListenerAndDatabaseDetails {
    param (
        [Parameter(Mandatory=$true)]
        [string]$InstanceName,
        [Parameter(Mandatory=$true)]
        [System.Management.Automation.PSCredential]$SqlCredential
    )

    try {
        Write-Log -Message "Retrieving AG Listener details for $InstanceName" -Level INFO
        $AGLs = Get-DbaAgListener -SqlInstance $InstanceName -SqlCredential $SqlCredential
        $listenerDetails = $AGLs | Select-Object AvailabilityGroup, @{Name='ListenerName';Expression={$_.Name}}, ClusterIPConfiguration, PortNumber

        Write-Log -Message "Retrieving AG Database details for $InstanceName" -Level INFO
        $AGDBs = Get-DbaAgDatabase -SqlInstance $InstanceName -SqlCredential $SqlCredential

        # Group databases by Availability Group
        $groupedDatabases = $AGDBs | Group-Object -Property AvailabilityGroup

        $details = @{
            'Listeners' = $listenerDetails
            'Databases' = $groupedDatabases
        }

        return $details
    }
    catch {
        Write-Log -Message "Failed to retrieve AG Listener and Database details for $InstanceName. Error: $_" -Level ERROR
        return @{
            'Listeners' = @()
            'Databases' = @()
        }
    }
}

#######################################################
### Main function to generate the As Built Document ###
#######################################################

function Generate-AsBuiltDoc {
    param (
        [Parameter(Mandatory=$true)]
        [string[]]$Instances,
        [string]$OutputPath,
        [string]$DocumentFormat
    )

    foreach ($instance in $Instances) {
        Write-Log -Message "Starting document generation for $instance" -Level INFO
        
        $hostserver = Get-HostServerProperties -InstanceName $instance
        $config = Get-SQLServerConfig -InstanceName $instance
        $databases = Get-SQLDatabases -InstanceName $instance
        
        # Get disk properties
        $DiskInfo = Get-CimInstance -ClassName Win32_Volume | Where-Object {($_.DriveLetter).Length -eq 2} | ForEach-Object {
            $diskData = @{
                'Name' = $_.Name
                'Label' = if ($_.Name -eq 'C:\' -and [string]::IsNullOrEmpty($_.Label)) { 'OS' } else { $_.Label }
                'Size (GB)' = [math]::Round($_.Capacity / 1GB, 2)
                'Free Space (GB)' = [math]::Round($_.FreeSpace / 1GB, 2)
                'Block Size' = $_.BlockSize
            }
            $diskData
        }

        # Get services properties
        $ServicesInfo = Get-DbaService | ForEach-Object {
            $ServicesData = @{
                'Service Name' = $_.ServiceName
                'Display Name' = $_.DisplayName
                'Service Account' = $_.StartName
                'Start Mode' = $_.StartMode
                'State' = $_.State
            }
            $ServicesData
        }

        # Get network protocols
        $NetworkProtocols = try {
            Write-Log -Message "Retrieving network protocol information for $instance" -Level INFO
            $protocols = Get-DbaInstanceProtocol -ErrorAction Stop
            $protocols | ForEach-Object {
                @{
                    'DisplayName' = $_.DisplayName
                    'Enabled' = $_.IsEnabled
                    'Port' = if ($_.Name -eq "Tcp") { (Get-DbaTcpPort -SqlInstance $instance).Port } else { "N/A" }
                }
            }
        }
        catch {
            Write-Log -Message "Failed to retrieve network protocols for $instance. Error: $_" -Level ERROR
            @()
        }

        # Get linked servers
        $LinkedServers = try {
            Write-Log -Message "Retrieving linked server information for $instance" -Level INFO
            Get-DbaLinkedServer -SqlInstance $instance -ErrorAction Stop | ForEach-Object {
                @{
                    'Linked Server Name' = $_.Name
                    'Product Name' = $_.ProductName
                    'Provider Name' = $_.ProviderName
                    'Data Source' = $_.DataSource
                    'Location' = $_.Location
                    'Provider String' = $_.ProviderString
                    'Is Remote Login Enabled' = $_.IsRemoteLoginEnabled
                    'Is RPC Out Enabled' = $_.IsRpcOutEnabled
                }
            }
        }
        catch {
            Write-Log -Message "Failed to retrieve linked server information for $instance. Error: $_" -Level ERROR
            @()
        }

        if ($DocumentFormat -eq "markdown") {
            $documentContent = "h2. SQL Server: $instance @ $(get-date -format "dd MMMM yyyy")`n"

            # Set host server properties in document
            $documentContent += "`nh3. Host Server Properties`n"
            $documentContent += "| *Property* | *Value* |`n"
            foreach ($property in $hostserver) {
                $documentContent += "| $($property.Name) | $($property.Value) |`n"
            }

            # Set SQL Server properties in document
            $documentContent += "h3. SQL Server Instance Properties`n"
            $documentContent += "| *Property* | *Value* |`n"
            foreach ($key in $config.Keys) {
                $documentContent += "| $key | $($config[$key]) |`n"
            }

            # Set disk properties in document
            $documentContent += "`nh3. Disk Properties`n"
            $documentContent += "| *Name* | *Label* | *Size (GB)* | *Free Space (GB)* | *Block Size* |`n"
            foreach ($disk in $DiskInfo) {
                $documentContent += "| $($disk.'Name') | $($disk.'Label') | $($disk.'Size (GB)') | $($disk.'Free Space (GB)') | $($disk.'Block Size') |`n"
            }

            # Set service properties in document
            $documentContent += "`nh3. Services Properties`n"
            $documentContent += "| *Service Name* | *Display Name* | *Service Account* | *Start Mode* | *State* |`n"
            foreach ($Service in $ServicesInfo) {
                $documentContent += "| $($Service.'Service Name') | $($Service.'Display Name') | $($Service.'Service Account') | $($Service.'Start Mode') | $($Service.'State') |`n"
            }

            # Set network protocols in document
            $documentContent += "`nh3. Network Protocols`n"
            $documentContent += "| *Name* | *Enabled* | *Port* |`n"
            foreach ($protocol in $NetworkProtocols) {
                $documentContent += "| $($protocol.'DisplayName') | $($protocol.'Enabled') | $($protocol.'Port') |`n"
            }

            # Get system databases' file properties --
            $SystemDbFiles = try {
                Write-Log -Message "Retrieving system database file information for $instance" -Level INFO
                Get-DbaDbFile -SqlInstance $instance -Database (Get-DbaDatabase -SqlInstance $instance | Where-Object {$_.IsSystemObject}).Name | ForEach-Object {
                    @{
                        'Database' = $_.Database
                        'File Type' = $_.Type
                        'Logical Name' = $_.Name
                        'Physical Name' = $_.PhysicalName
                        'Size' = [math]::Round($_.Size / 1MB, 2)  # Size in MB
                        'Growth' = if ($_.GrowthType -eq 'Percent') {
                            "$($_.Growth)%"
                        } else {
                            "$($_.Growth) MB"
                        }
                    }
                }
            }
            catch {
                Write-Log -Message "Failed to retrieve system database file information for $instance. Error: $_" -Level ERROR
                @()
            }

            # Set system databases' properties in document
            $documentContent += "`nh3. System Databases`n"
            $orderedPropertiesForSystem = @('Owner', 'RecoveryModel', 'IsAutoCreateStatisticsEnabled', 'LastBackupDate', 'IsReadCommittedSnapshotOn', 'AutoShrink', 'Status', 'SizeMB', 'AutoClose', 'IsAutoUpdateStatisticsEnabled')
            foreach ($db in $databases.SystemDatabases) {
                $documentContent += "`nh4. $($db.Name)`n"
                $documentContent += "| *Property* | *Value* |`n"
                foreach ($prop in $orderedPropertiesForSystem) {
                    if ($db.ContainsKey($prop)) {
                        $documentContent += "| $prop | $($db[$prop]) |`n"
                    }
                }
            }

            $documentContent += "`nh3. System Database File Sizes and Growth`n"
            $documentContent += "| *Database* | *File Type* | *Logical Name* | *Physical Name* | *Size (MB)* | *Growth* |`n"
            foreach ($file in $SystemDbFiles) {
                $documentContent += "| $($file.'Database') | $($file.'File Type') | $($file.'Logical Name') | $($file.'Physical Name') | $($file.'Size') | $($file.'Growth') |`n"
            }

            # Get user databases' file properties --
            $UserDbFiles = try {
                Write-Log -Message "Retrieving system database file information for $instance" -Level INFO
                Get-DbaDbFile -SqlInstance $instance -Database (Get-DbaDatabase -SqlInstance $instance | Where-Object {-not $_.IsSystemObject}).Name | ForEach-Object {
                    @{
                        'Database' = $_.Database
                        'File Type' = $_.Type
                        'Logical Name' = $_.Name
                        'Physical Name' = $_.PhysicalName
                        'Size' = [math]::Round($_.Size / 1MB, 2)  # Size in MB
                        'Growth' = if ($_.GrowthType -eq 'Percent') {
                            "$($_.Growth)%"
                        } else {
                            "$($_.Growth) MB"
                        }
                    }
                }
            }
            catch {
                Write-Log -Message "Failed to retrieve system database file information for $instance. Error: $_" -Level ERROR
                @()
            }

            # Set user databases' file properties in document
            $documentContent += "`nh3. User Database File Sizes and Growth`n"
            $documentContent += "| *Database* | *File Type* | *Logical Name* | *Physical Name* | *Size (MB)* | *Growth* |`n"
            foreach ($file in $UserDbFiles) {
                $documentContent += "| $($file.'Database') | $($file.'File Type') | $($file.'Logical Name') | $($file.'Physical Name') | $($file.'Size') | $($file.'Growth') |`n"
            }

            # Set user databases' properties in document
            $documentContent += "`nh3. User Databases`n"
            $orderedPropertiesForUser = @('Owner', 'RecoveryModel', 'IsAutoCreateStatisticsEnabled', 'LastBackupDate', 'IsReadCommittedSnapshotOn', 'AutoShrink', 'Status', 'SizeMB', 'CreationDate', 'AutoClose', 'CompatibilityLevel', 'IsAutoUpdateStatisticsEnabled')
            foreach ($db in $databases.UserDatabases) {
                $documentContent += "`nh4. $($db.Name)`n"
                $documentContent += "| *Property* | *Value* |`n"
                foreach ($prop in $orderedPropertiesForUser) {
                    if ($db.ContainsKey($prop)) {
                        $documentContent += "| $prop | $($db[$prop]) |`n"
                    }
                }
            }

            # Set Availability Group properties in document
            $agDetails = Add-AgListenerAndDatabaseDetails -InstanceName $instance -SqlCredential $SqlCredential

            if ($DocumentFormat -eq "markdown") {
                # Markdown formatting for Listeners
                $documentContent += "`nh3. Availability Group Listeners`n"
                $documentContent += "| AvailabilityGroup | ListenerName | ClusterIPConfiguration | PortNumber |`n"
                foreach ($listener in $agDetails.Listeners) {
                    $documentContent += "| $($listener.AvailabilityGroup) | $($listener.ListenerName) | $($listener.ClusterIPConfiguration) | $($listener.PortNumber) |`n"
                }

                # Markdown formatting for Databases
                $documentContent += "`nh3. Availability Group Databases`n"
                $documentContent += "| AvailabilityGroup | Databases | LocalReplicaRole | SynchronizationState |`n"
                foreach ($group in $agDetails.Databases) {
                    # Use -join operator instead of Join-String
                    $databases = ($group.Group | Sort-Object -Property Name | ForEach-Object { $_.Name }) -join ", "
                    $sampleDb = $group.Group | Select-Object -First 1
                    $documentContent += "| $($group.Name) | $databases | $($sampleDb.LocalReplicaRole) | $($sampleDb.SynchronizationState) |`n"
                }
            } else {
                Write-Log -Message "Document format $DocumentFormat not implemented for AG details yet." -Level WARNING
            }

            # Set Linked Server properties in document
            $documentContent += "`nh3. Linked Servers`n"
            $documentContent += "| *Linked Server Name* | *Product Name* | *Provider Name* | *Data Source* | *Location* | *Provider String* | *Is Remote Login Enabled* | *Is RPC Out Enabled* |`n"
            foreach ($server in $LinkedServers) {
                $documentContent += "| $($server.'Linked Server Name') | $($server.'Product Name') | $($server.'Provider Name') | $($server.'Data Source') | $($server.'Location') | $($server.'Provider String') | $($server.'Is Remote Login Enabled') | $($server.'Is RPC Out Enabled') |`n"
            }

            # Generate unique filename with server name and date
            $dateStamp = Get-Date -Format "yyyyMMdd"
            $outputFile = Join-Path -Path $OutputPath -ChildPath "$($instance)_$dateStamp.md"
            $documentContent | Out-File -FilePath $outputFile
            Write-Log -Message "Document for $instance saved to $outputFile" -Level SUCCESS
        }
        else {
            Write-Log -Message "Document format $DocumentFormat not supported yet." -Level WARNING
        }
    }
}

######################
### Main execution ###
######################

Write-Log -Message "Starting SQL Server As Built Document Generation" -Level INFO
try {
    Generate-AsBuiltDoc -Instances $SQLServerInstances -OutputPath $OutputPath -DocumentFormat $DocumentFormat
}
catch {
    Write-Log -Message "An error occurred during document generation: $_" -Level "ERROR"
}
Write-Log -Message "Finished SQL Server As Built Document Generation" -Level INFO