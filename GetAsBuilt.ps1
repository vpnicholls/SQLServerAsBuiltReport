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
        $serverName = if ($instance.Contains('\')) { $instance.Split('\')[0] } else { $instance }
        try {
            $ServerCredentials[$serverName] = Import-Clixml -Path ".\Credentials\$serverName.xml"
            Write-Log -Message "Successfully imported Windows credentials for $serverName." -Level "INFO"
        }
        catch {
            Write-Log -Message "Failed to import Windows credentials for $serverName from .\Credentials\$serverName.xml: $_" -Level "ERROR"
        }
    }
}
else {
    Write-Log -Message "Using domain account for authentication." -Level INFO
}

# Define function to get basic properties for host server 
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
# Define function to get server configuration using dbatools
function Get-SQLServerConfig {
    param (
        [Parameter(Mandatory=$true)]
        [string]$InstanceName
    )
    try {
        Write-Log -Message "Retrieving configuration for $InstanceName" -Level INFO
        
        # Initialize configuration hashtable
        $config = @{}

        # Get properties from Get-DbaInstanceProperty
        $serverInfo = Get-DbaInstanceProperty -SqlInstance $InstanceName -ErrorAction Stop
        @('VersionString', 'Edition', 'Collation', 'Language', 'IsFullTextInstalled', 'IsHADREnabled', 'LoginMode') | ForEach-Object {
            $property = $_
            $value = ($serverInfo | Where-Object { $_.Name -eq $property } | Select-Object -ExpandProperty Value -ErrorAction SilentlyContinue)
            if ($null -ne $value) { 
                $config[$property] = $value 
            } else { 
                Write-Log -Message "Property $property not found for $InstanceName" -Level WARNING 
            }
        }

        # Get properties from Get-DbaSpConfigure
        $spConfig = Get-DbaSpConfigure -SqlInstance $InstanceName -ErrorAction Stop | 
            Where-Object { $_.DisplayName -in @(
                "backup checksum default", 
                "backup compression default", 
                "blocked process threshold (s)", 
                "max server memory (MB)", 
                "min server memory (MB)", 
                "max degree of parallelism", 
                "cost threshold for parallelism"
            )} | Sort-Object DisplayName

        # Map DisplayName to user-friendly property names and add to config
        foreach ($item in $spConfig) {
            $propertyName = switch ($item.DisplayName) {
                "backup checksum default" { "Backup Checksum Default" }
                "backup compression default" { "Backup Compression Default" }
                "blocked process threshold (s)" { "Blocked Process Threshold (s)" }
                "max server memory (MB)" { "Max Server Memory (MB)" }
                "min server memory (MB)" { "Min Server Memory (MB)" }
                "max degree of parallelism" { "Max Degree of Parallelism" }
                "cost threshold for parallelism" { "Cost Threshold for Parallelism" }
                default { $item.DisplayName }
            }
            # Format boolean values as Yes/No for consistency
            $value = if ($item.DisplayName -in @("backup checksum default", "backup compression default")) {
                if ($item.ConfiguredValue -eq 1) { "Yes" } else { "No" }
            } else {
                $item.ConfiguredValue
            }
            $config[$propertyName] = $value
        }

        return $config
    }
    catch {
        Write-Log -Message "Failed to retrieve configuration for $InstanceName. Error: $_" -Level ERROR
        return $null
    }
}

# Function to get databases' properties
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
        function Get-DatabaseDetails {
            param ($db)
            $dbFiles = Get-DbaDbFile -SqlInstance $InstanceName -Database $db.Name -WarningAction SilentlyContinue
            if (-not $dbFiles) {
                Write-Log -Message "Database '$($db.Name)' on $InstanceName is not readable (likely an AG secondary)" -Level INFO
            }
            $totalSizeBytes = ($dbFiles | Measure-Object -Property Size -Sum).Sum
            return @{
                'Name' = $db.Name
                'Status' = $db.Status
                'RecoveryModel' = $db.RecoveryModel
                'CompatibilityLevel' = $db.Compatibility -replace 'Version',''
                'Owner' = $db.Owner
                'CreationDate' = $db.CreateDate.ToString("yyyy-MM-dd HH:mm:ss")
                'LastBackupDate' = if ($db.Name -eq 'tempdb') { "N/A " } elseif ($db.LastBackupDate) { $db.LastBackupDate.ToString("yyyy-MM-dd HH:mm:ss") } else { "Never" }
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

# Function to get Availability Group listeners' and databases' properties
function Add-AgListenerAndDatabaseDetails {
    param (
        [Parameter(Mandatory=$true)]
        [string]$InstanceName,
        [Parameter(Mandatory=$true)]
        [System.Management.Automation.PSCredential]$SqlCredential
    )
    try {
        Write-Log -Message "Retrieving AG Listener details for $InstanceName" -Level INFO
        $AGLs = Get-DbaAgListener -SqlInstance $InstanceName -SqlCredential $SqlCredential -WarningAction SilentlyContinue
        if (-not $AGLs) {
            Write-Log -Message "No Availability Group listeners detected on $InstanceName" -Level INFO
        }
        $listenerDetails = $AGLs | ForEach-Object {
            $ipMatches = [regex]::Matches($_.ClusterIPConfiguration, 'IP Address: (\d+\.\d+\.\d+\.\d+)')
            $ipList = $ipMatches | ForEach-Object { $_.Groups[1].Value }
            [PSCustomObject]@{
                AvailabilityGroup = $_.AvailabilityGroup
                ListenerName = $_.Name
                ClusterIPConfiguration = $ipList
                PortNumber = $_.PortNumber
            }
        }
        Write-Log -Message "Retrieving AG Database details for $InstanceName" -Level INFO
        $AGDBs = Get-DbaAgDatabase -SqlInstance $InstanceName -SqlCredential $SqlCredential -WarningAction SilentlyContinue
        if (-not $AGDBs -and -not $AGLs) {
            Write-Log -Message "Availability Groups (HADR) not configured for $InstanceName" -Level INFO
        }
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

# Function to get Database Mail details
function Get-DatabaseMailDetails {
    param (
        [Parameter(Mandatory=$true)]
        [string]$InstanceName,
        [System.Management.Automation.PSCredential]$SqlCredential
    )

    try {
        Write-Log -Message "Retrieving Database Mail details for $InstanceName" -Level INFO

        # Get Mail Accounts
        $mailAccounts = Get-DbaDbMailAccount -SqlInstance $InstanceName -SqlCredential $SqlCredential | 
            Select-Object Name, EmailAddress, DisplayName, ReplyToAddress, @{Name='SmtpServer';Expression={$_.MailServers.ServerName}}

        # Get Mail Profiles
        $mailProfiles = Get-DbaDbMailProfile -SqlInstance $InstanceName -SqlCredential $SqlCredential | 
            Select-Object Name, Description, IsDefault

        # Get Mail Servers (linked to accounts)
        $mailServers = Get-DbaDbMailServer -SqlInstance $InstanceName -SqlCredential $SqlCredential | 
            Select-Object AccountName, ServerName, Port, EnableSsl, @{Name='AuthenticationType';Expression={$_.AuthenticationType.ToString()}}

        # Get Mail Configuration (optional, for key settings)
        $mailConfig = Get-DbaDbMailConfig -SqlInstance $InstanceName -SqlCredential $SqlCredential | 
            Select-Object Name, Value

        if (-not $mailAccounts -and -not $mailProfiles -and -not $mailServers) {
            Write-Log -Message "No Database Mail configuration detected on $InstanceName" -Level INFO
            return $null
        }

        $mailDetails = [PSCustomObject]@{
            Accounts = $mailAccounts
            Profiles = $mailProfiles
            Servers = $mailServers
            Config = $mailConfig
        }

        Write-Log -Message "Database Mail details collected: $($mailAccounts.Count) accounts, $($mailProfiles.Count) profiles, $($mailServers.Count) servers" -Level DEBUG
        return $mailDetails
    }
    catch {
        Write-Log -Message "Failed to retrieve Database Mail details for $InstanceName. Error: $_" -Level ERROR
        return $null
    }
}

# Function to get replication publisher details
function Get-ReplicationPublisherDetails {
    param (
        [Parameter(Mandatory=$true)]
        [string]$InstanceName,
        [System.Management.Automation.PSCredential]$SqlCredential
    )

    try {
        Write-Log -Message "Retrieving replication publisher details for $InstanceName" -Level INFO

        # Get only user databases that are published
        $publishedDatabases = Get-DbaDatabase -SqlInstance $InstanceName -SqlCredential $SqlCredential -ExcludeSystem | 
            Where-Object { $_.ReplicationOptions -eq "Published" } | 
            Select-Object -ExpandProperty Name
        Write-Log -Message "Found $($publishedDatabases.Count) published user databases: $($publishedDatabases -join ', ')" -Level DEBUG

        # SQL query for publisher details
        $pubQuery = @"
        SELECT 
            [p].[name] AS [PublicationName],
            DB_NAME() AS [Database],
            [p].[snapshot_in_defaultfolder] AS [DefaultSnapshotFolder],
            [p].[alt_snapshot_folder] AS [AlternateSnapshotFolder]
        FROM dbo.syspublications AS [p]
"@

        # Initialize as an empty array
        $publisherDetails = @()
        foreach ($db in $publishedDatabases) {
            Write-Log -Message "Checking database '$db' for publications" -Level DEBUG
            $pubs = Invoke-DbaQuery -SqlInstance $InstanceName -SqlCredential $SqlCredential -Database $db -Query $pubQuery -WarningAction SilentlyContinue -ErrorAction Stop
            if (-not $pubs) {
                Write-Log -Message "Database '$db' on $InstanceName is not readable for replication details (likely an AG secondary)" -Level INFO
            }
            if ($pubs) {
                Write-Log -Message "Found $($pubs.Count) publications in '$db': $($pubs.PublicationName -join ', ')" -Level INFO
                $publisherDetails += $pubs
            }
        }

        # Log final results
        Write-Log -Message "Final publisherDetails count: $($publisherDetails.Count)" -Level DEBUG
        if ($publisherDetails) {
            $pubNames = ($publisherDetails | ForEach-Object { $_.PublicationName }) -join ', '
            Write-Log -Message "Publisher details collected: $pubNames" -Level DEBUG
            Write-Log -Message "Detected publisher role with $($publisherDetails.Count) publications⁠" -Level INFO
        } else {
            Write-Log -Message "No publisher role detected on $InstanceName" -Level INFO
        }

        return $publisherDetails
    }
    catch {
        Write-Log -Message "Failed to retrieve replication publisher details for $InstanceName. Error: $_" -Level ERROR
        return @()
    }
}

# Function to get replication distributor details
function Get-ReplicationDistributorDetails {
    param (
        [Parameter(Mandatory=$true)]
        [string]$InstanceName,
        [System.Management.Automation.PSCredential]$SqlCredential
    )

    try {
        Write-Log -Message "Retrieving replication distributor details for $InstanceName" -Level INFO

        # Check if this instance is a distributor using dbatools
        $distributorInfo = Get-DbaRepDistributor -SqlInstance $InstanceName -SqlCredential $SqlCredential
        if (-not $distributorInfo -or $distributorInfo.IsDistributor -ne $true) {
            Write-Log -Message "No distributor role detected on $InstanceName" -Level INFO
            return $null
        }

        # Log distributor detection
        Write-Log -Message "Distributor role detected on $InstanceName" -Level INFO

        # Get distribution database name
        $distDbQuery = @"
        SELECT name AS DistributionDatabase
        FROM master.sys.databases
        WHERE is_distributor = 1
"@
        $distDb = Invoke-DbaQuery -SqlInstance $InstanceName -SqlCredential $SqlCredential -Database master -Query $distDbQuery -ErrorAction Stop
        if (-not $distDb) {
            Write-Log -Message "No distribution database found in master.sys.databases" -Level WARNING
            return $null
        }
        $distDbName = $distDb.DistributionDatabase

        # Get publishers from MSdistpublishers
        $publishersQuery = @"
        SELECT [name] AS PublisherName
        FROM msdb.dbo.MSdistpublishers
        WHERE active = 1
"@
        $publishers = Invoke-DbaQuery -SqlInstance $InstanceName -SqlCredential $SqlCredential -Database msdb -Query $publishersQuery -ErrorAction Stop

        # Get publications from MSpublications in the distribution database
        $publicationsQuery = @"
        SELECT 
            publication AS PublicationName,
            publisher_db AS PublisherDatabase,
            CASE 
                WHEN EXISTS (SELECT 1 FROM msdb.dbo.MSdistpublishers mp WHERE mp.name = '$($publishers.PublisherName)' AND mp.active = 1)
                THEN '$($publishers.PublisherName)'
                ELSE 'Unknown'
            END AS PublisherName
        FROM [$distDbName].dbo.MSpublications
"@
        $publications = Invoke-DbaQuery -SqlInstance $InstanceName -SqlCredential $SqlCredential -Database $distDbName -Query $publicationsQuery -ErrorAction Stop

        if (-not $publishers -and -not $publications) {
            Write-Log -Message "No active publishers or publications found for $InstanceName" -Level INFO
            $distributorDetails = [PSCustomObject]@{
                Publishers = @([PSCustomObject]@{ PublisherName = "None"; PublicationName = "None" })
            }
        } else {
            # Compile details, prioritizing publications and falling back to publishers without publications
            $pubList = @()
            if ($publications) {
                $pubList = $publications | ForEach-Object {
                    [PSCustomObject]@{
                        PublisherName = $_.PublisherName
                        PublicationName = $_.PublicationName
                    }
                }
            } else {
                # If no publications, list publishers with "None"
                $pubList = $publishers | ForEach-Object {
                    [PSCustomObject]@{
                        PublisherName = $_.PublisherName
                        PublicationName = "None"
                    }
                }
            }

            $distributorDetails = [PSCustomObject]@{
                Publishers = $pubList
            }
        }

        Write-Log -Message "Distributor details collected: $(($distributorDetails.Publishers | ForEach-Object { "$($_.PublisherName):$($_.PublicationName)" }) -join ', ')" -Level DEBUG
        Write-Log -Message "Detected distributor role with $(($distributorDetails.Publishers).Count) publisher entries" -Level INFO

        return $distributorDetails
    }
    catch {
        Write-Log -Message "Failed to retrieve replication distributor details for $InstanceName. Error: $_" -Level ERROR
        return $null
    }
}

# Function to get replication subscriber details
function Get-ReplicationSubscriptionDetails {
    param (
        [Parameter(Mandatory=$true)]
        [string]$InstanceName,
        [System.Management.Automation.PSCredential]$SqlCredential,
        [PSCustomObject]$DistributorDetails
    )

    try {
        Write-Log -Message "Retrieving replication subscription details for $InstanceName" -Level INFO

        # Check if this instance is a distributor
        $distributorInfo = Get-DbaRepDistributor -SqlInstance $InstanceName -SqlCredential $SqlCredential
        $isDistributor = $distributorInfo -and $distributorInfo.IsDistributor -eq $true

        if (-not $isDistributor) {
            Write-Log -Message "Instance $InstanceName is not a Distributor; skipping subscription details" -Level INFO
            return $null
        }

        # Initialize subscription collection
        $subscriptions = @()

        # Try getting subscriptions from the Distributor
        $subscriptions = Get-DbaReplSubscription -SqlInstance $InstanceName -SqlCredential $SqlCredential | 
            Select-Object PublicationName, SubscriptionType, DatabaseName, SubscriptionDBName
        
        if (-not $subscriptions -and $DistributorDetails -and $DistributorDetails.Publishers) {
            # If no subscriptions found on Distributor, query each Publisher
            Write-Log -Message "No subscriptions found on Distributor $InstanceName, checking Publishers" -Level INFO
            foreach ($pub in $DistributorDetails.Publishers | Where-Object { $_.PublisherName -ne "None" }) {
                $pubInstance = $pub.PublisherName
                Write-Log -Message "Querying subscriptions from Publisher $pubInstance" -Level DEBUG
                $pubSubs = Get-DbaReplSubscription -SqlInstance $pubInstance -SqlCredential $SqlCredential -ErrorAction Stop | 
                    Select-Object PublicationName, SubscriptionType, DatabaseName, SubscriptionDBName
                if ($pubSubs) {
                    $subscriptions += $pubSubs
                }
            }
        }

        if (-not $subscriptions) {
            Write-Log -Message "No subscriptions found for $InstanceName or its Publishers" -Level INFO
            $subscriptionDetails = [PSCustomObject]@{
                Subscriptions = @([PSCustomObject]@{
                    PublisherName = "None"
                    PublicationName = "None"
                    SubscriptionType = "None"
                    DatabaseName = "None"
                    SubscriptionDBName = "None"
                })
            }
        } else {
            # Cross-reference with DistributorDetails to add PublisherName
            $subscriptionDetails = [PSCustomObject]@{
                Subscriptions = $subscriptions | ForEach-Object {
                    $pubName = $_.PublicationName
                    $publisher = if ($DistributorDetails -and $DistributorDetails.Publishers) {
                        $matchingPub = $DistributorDetails.Publishers | Where-Object { $_.PublicationName -eq $pubName }
                        if ($matchingPub) { $matchingPub.PublisherName } else { "Unknown" }
                    } else { "Unknown" }
                    [PSCustomObject]@{
                        PublisherName = $publisher
                        PublicationName = $_.PublicationName
                        SubscriptionType = $_.SubscriptionType
                        DatabaseName = $_.DatabaseName
                        SubscriptionDBName = $_.SubscriptionDBName
                    }
                }
            }
        }

        Write-Log -Message "Subscription details collected: $(($subscriptionDetails.Subscriptions | ForEach-Object { "$($_.PublisherName):$($_.PublicationName):$($_.SubscriptionDBName)" }) -join ', ')" -Level DEBUG
        Write-Log -Message "Detected $(($subscriptionDetails.Subscriptions).Count) subscription entries" -Level INFO

        return $subscriptionDetails
    }
    catch {
        Write-Log -Message "Failed to retrieve replication subscription details for $InstanceName. Error: $_" -Level ERROR
        return $null
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
        $DiskInfo = Get-CimInstance -ClassName Win32_Volume | Where-Object { 
            ($_.DriveLetter).Length -eq 2 -and 
            $_.Capacity -ne $null -and 
            $_.Capacity -gt 0 
        } | ForEach-Object {
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

        # Get Endpoint properties
        $EndpointProperties = try {
            Write-Log -Message "Retrieving Endpoint properties for $instance" -Level INFO
            Get-DbaEndpoint -SqlInstance $instance -SqlCredential $SqlCredential | Where-Object {$_.IsSystemObject -eq $False} -ErrorAction Stop | ForEach-Object {
                @{
                    'EndpointType' = $_.EndpointType
                    'Owner' = $_.Owner
                    'ProtocolType' = $_.ProtocolType
                    'Name' = $_.Name
                    'Port' = $_.Port
                }
            }
        } catch {
            Write-Log -Message "Failed to retrieve Endpoint properties for $instance. Error: $_" -Level ERROR
            @()
        }

        # Get database certificates
        $DatabaseCertificates = try {
            Write-Log -Message "Retrieving Database Certificate properties for $instance" -Level INFO
            $certificates = Get-DbaDbCertificate -SqlInstance $instance -SqlCredential $SqlCredential -WarningAction SilentlyContinue | Where-Object {$_.PrivateKeyEncryptionType -ne "NoKey"} -ErrorAction Stop
            if (-not $certificates) {
                Write-Log -Message "No readable database certificates detected on $instance (databases may be AG secondaries)" -Level INFO
            }
            $certificates | ForEach-Object {
                Write-Log -Message "Retrieved certificate '$($_.Name)' from database '$($_.Database)' on $instance" -Level DEBUG
                @{
                    'Database' = $_.Database
                    'Name' = $_.Name
                    'Issuer' = $_.Issuer
                    'PrivateKeyEncryptionType' = $_.PrivateKeyEncryptionType
                    'StartDate' = $_.StartDate.ToString("yyyy-MM-dd")
                    'ExpirationDate' = $_.ExpirationDate.ToString("yyyy-MM-dd")
                }
            }
        } catch {
            Write-Log -Message "Failed to retrieve Database Certificate properties for $instance. Error: $_" -Level ERROR
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
                    'Impersonate' = $_.Impersonate
                    'RpcOut' = $_.RpcOut
                }
            }
        }
        catch {
            Write-Log -Message "Failed to retrieve linked server information for $instance. Error: $_" -Level ERROR
            @()
        }

        # Get Database Mail details
        $mailDetails = Get-DatabaseMailDetails -InstanceName $instance -SqlCredential $SqlCredential

        # Get replication publisher details
        $publisherDetails = Get-ReplicationPublisherDetails -InstanceName $instance -SqlCredential $SqlCredential

        # Get replication distributor details
        $distributorDetails = Get-ReplicationDistributorDetails -InstanceName $instance -SqlCredential $SqlCredential

        # Get replication subscription details
        $subscriptionDetails = Get-ReplicationSubscriptionDetails -InstanceName $instance -SqlCredential $SqlCredential -DistributorDetails $distributorDetails

        # Get system databases' file properties
        $SystemDbFiles = try {
            Write-Log -Message "Retrieving system database file information for $instance" -Level INFO
            Get-DbaDbFile -SqlInstance $instance -Database (Get-DbaDatabase -SqlInstance $instance | Where-Object {$_.IsSystemObject}).Name | ForEach-Object {
                $sizeMB = [math]::Round($_.Size.Megabyte, 2)
                $growthValue = if ($_.GrowthType -eq 'Percent') { "$($_.Growth)%" } else { [math]::Round($_.Growth / 1024, 2) }
                Write-Log -Message "Raw data for $($_.LogicalName): Size = $($_.Size.Megabyte) MB, Growth = $($_.Growth), GrowthType = $($_.GrowthType), Converted Size = $sizeMB MB, Converted Growth = $growthValue" -Level DEBUG
                @{
                    'Database' = $_.Database
                    'File Type' = $_.TypeDescription
                    'Logical Name' = $_.LogicalName
                    'Physical Name' = $_.PhysicalName
                    'Size' = $sizeMB
                    'Growth' = $growthValue
                }
            }
        }
        catch {
            Write-Log -Message "Failed to retrieve system database file information for $instance. Error: $_" -Level ERROR
            @()
        }

        # Get user databases' file properties
        $UserDbFiles = try {
            Write-Log -Message "Retrieving user database file information for $instance" -Level INFO
            $userDbs = Get-DbaDatabase -SqlInstance $instance | Where-Object {-not $_.IsSystemObject}
            $files = Get-DbaDbFile -SqlInstance $instance -Database $userDbs.Name -WarningAction SilentlyContinue | ForEach-Object {
                $sizeMB = [math]::Round($_.Size.Megabyte, 2)
                $growthValue = if ($_.GrowthType -eq 'Percent') { "$($_.Growth)%" } else { [math]::Round($_.Growth / 1024, 2) }
                Write-Log -Message "Raw data for $($_.LogicalName): Size = $($_.Size.Megabyte) MB, Growth = $($_.Growth), GrowthType = $($_.GrowthType), Converted Size = $sizeMB MB, Converted Growth = $growthValue" -Level DEBUG
                @{
                    'Database' = $_.Database
                    'File Type' = $_.TypeDescription
                    'Logical Name' = $_.LogicalName
                    'Physical Name' = $_.PhysicalName
                    'Size' = $sizeMB
                    'Growth' = $growthValue
                }
            }
            if (-not $files -and $userDbs) {
                Write-Log -Message "No readable user database files retrieved from $instance. (Databases may be Availability Group secondaries)." -Level INFO
            }
            $files
        }
        catch {
            Write-Log -Message "Failed to retrieve user database file information for $instance. Error: $_" -Level ERROR
            @()
        }

        if ($DocumentFormat -eq "markdown") {
            $documentContent = "h2. SQL Server: $instance @ $(Get-Date -Format 'dd MMMM yyyy')`n"

            # Set host server properties in document
            $documentContent += "`nh3. Host Server Properties`n"
            $documentContent += "|*Property*|*Value*|`n"
            foreach ($property in $hostserver) {
                if ($property.Name -eq "PhysicalMemory") {
                    $mbValue = "{0:N0}" -f $property.Value
                    $gbValue = "{0:N1}" -f ($property.Value / 1024)
                    $formattedValue = "$mbValue MB / $gbValue GB"
                    $documentContent += "|$($property.Name)|$formattedValue|`n"
                } else {
                    $documentContent += "|$($property.Name)|$($property.Value)|`n"
                }
            }

            # Set SQL Server properties in document
            $documentContent += "`nh3. SQL Server Instance Properties`n"
            $documentContent += "|*Property*|*Value*|`n"
            foreach ($key in ($config.Keys | Sort-Object)) {
                $documentContent += "|$key|$($config[$key])|`n"
            }

            # Set disk properties in document
            $documentContent += "`nh3. Disk Properties`n"
            $documentContent += "|*Name*|*Label*|*Size (GB)*|*Free Space (GB)*|*Block Size (KB)*|`n"
            foreach ($disk in $DiskInfo) {
                $spacedName = "$($disk.'Name') "  # Trailing space for backslash fix
                $sizeGB = "{0:N1}" -f $disk.'Size (GB)'
                $freeSpaceGB = "{0:N1}" -f $disk.'Free Space (GB)'
                $blockSizeKB = "{0:N0}" -f $disk.'Block Size'
                $documentContent += "|$spacedName|$($disk.'Label')|$sizeGB GB|$freeSpaceGB GB|$blockSizeKB|`n"
            }

            # Set service properties in document
            $documentContent += "`nh3. Services Properties`n"
            $documentContent += "|*Service Name*|*Display Name*|*Service Account*|*Start Mode*|*State*|`n"
            foreach ($Service in $ServicesInfo) {
                $documentContent += "|$($Service.'Service Name')|$($Service.'Display Name')|$($Service.'Service Account')|$($Service.'Start Mode')|$($Service.'State')|`n"
            }

            # Set network protocols in document
            $documentContent += "`nh3. Network Protocols`n"
            $documentContent += "|*Name*|*Enabled*|*Port*|`n"
            foreach ($protocol in $NetworkProtocols) {
                $documentContent += "|$($protocol.'DisplayName')|$($protocol.'Enabled')|$($protocol.'Port')|`n"
            }

            # Set system databases' properties in document
            $documentContent += "`nh3. System Databases`n"
            $orderedPropertiesForSystem = @('Owner', 'RecoveryModel', 'IsAutoCreateStatisticsEnabled', 'LastBackupDate', 'IsReadCommittedSnapshotOn', 'AutoShrink', 'Status', 'SizeMB', 'AutoClose', 'IsAutoUpdateStatisticsEnabled')
            foreach ($db in $databases.SystemDatabases) {
                $documentContent += "`nh4. $($db.Name)`n"
                $documentContent += "|*Property*|*Value*|`n"
                foreach ($prop in $orderedPropertiesForSystem) {
                    if ($db.ContainsKey($prop)) {
                        $documentContent += "|$prop|$($db[$prop])|`n"
                    }
                }
            }

            # Set system databases' file properties in document
            $documentContent += "`nh3. System Database File Sizes and Growth`n"
            $documentContent += "|*Database*|*File Type*|*Logical Name*|*Physical Name*|*Size*|*Growth*|`n"
            foreach ($file in $SystemDbFiles) {
                $sizeMB = $file.'Size'
                if ($sizeMB -ge 1024) {
                    $sizeValue = "{0:N1} GB" -f ($sizeMB / 1024)
                } else {
                    $sizeValue = "{0:N0} MB" -f $sizeMB
                }
                $growthValue = if ($file.'Growth' -match '%') { 
                    $file.'Growth'
                } else {
                    $growthMB = [double]$file.'Growth'
                    if ($growthMB -ge 1024) {
                        "{0:N1} GB" -f ($growthMB / 1024)
                    } else {
                        "{0:N0} MB" -f $growthMB
                    }
                }
                $abridgedPhysicalName = $file.'Physical Name' -replace 'Microsoft SQL Server\\MSSQL\d{2}\.MSSQLSERVER\\MSSQL', '...'
                $documentContent += "|$($file.'Database')|$($file.'File Type')|$($file.'Logical Name')|$abridgedPhysicalName|$sizeValue|$growthValue|`n"
            }

            # Set user databases' properties in document
            $documentContent += "`nh3. User Databases`n"
            $orderedPropertiesForUser = @('Owner', 'RecoveryModel', 'IsAutoCreateStatisticsEnabled', 'LastBackupDate', 'IsReadCommittedSnapshotOn', 'AutoShrink', 'Status', 'SizeMB', 'CreationDate', 'AutoClose', 'CompatibilityLevel', 'IsAutoUpdateStatisticsEnabled')
            foreach ($db in $databases.UserDatabases) {
                $documentContent += "`nh4. $($db.Name)`n"
                $documentContent += "|*Property*|*Value*|`n"
                foreach ($prop in $orderedPropertiesForUser) {
                    if ($db.ContainsKey($prop)) {
                        $value = $db[$prop]
                        # Strip 'Version' from CompatibilityLevel
                        if ($prop -eq 'CompatibilityLevel') {
                            $value = $value -replace 'Version', ''
                        }
                        $documentContent += "|$prop|$value|`n"
                    }
                }
            }

            $documentContent += "`nh3. User Database File Sizes and Growth`n"
            if ($UserDbFiles) {
                $documentContent += "|*Database*|*File Type*|*Logical Name*|*Physical Name*|*Size*|*Growth*|`n"
                foreach ($file in $UserDbFiles) {
                    $sizeMB = $file.'Size'
                    if ($sizeMB -ge 1024) {
                        $sizeValue = "{0:N1} GB" -f ($sizeMB / 1024)
                    } else {
                        $sizeValue = "{0:N0} MB" -f $sizeMB
                    }
                    $growthValue = if ($file.'Growth' -match '%') { 
                        $file.'Growth'
                    } else {
                        $growthMB = [double]$file.'Growth'
                        if ($growthMB -ge 1024) {
                            "{0:N1} GB" -f ($growthMB / 1024)
                        } else {
                            "{0:N0} MB" -f $growthMB
                        }
                    }
                    $documentContent += "|$($file.'Database')|$($file.'File Type')|$($file.'Logical Name')|$($file.'Physical Name')|$sizeValue|$growthValue|`n"
                }
            } else {
                $documentContent += "No readable user database configuration detected on this instance. Databases may be Availability Group secondaries.`n"
            }

            # Set Availability Group properties in document
            $agDetails = Add-AgListenerAndDatabaseDetails -InstanceName $instance -SqlCredential $SqlCredential
            $documentContent += "`nh3. Availability Group Listeners`n"
            if ($agDetails.Listeners) {
                $documentContent += "|*AvailabilityGroup*|*ListenerName*|*ClusterIPConfiguration*|*PortNumber*|`n"
                foreach ($listener in $agDetails.Listeners) {
                    $ipString = $listener.ClusterIPConfiguration -join ', '  # Join IPs with comma and space
                    $documentContent += "|$($listener.AvailabilityGroup)|$($listener.ListenerName)|$ipString|$($listener.PortNumber)|`n"
                }
            } else {
                $documentContent += "No Availability Group listeners detected on this instance.`n"
            }

            $documentContent += "`nh3. Availability Group Databases`n"
            if ($agDetails.Databases) {
                $documentContent += "|*AvailabilityGroup*|*Databases*|*LocalReplicaRole*|*SynchronizationState*|`n"
                foreach ($group in $agDetails.Databases) {
                    $databasesList = ($group.Group | Sort-Object -Property Name | ForEach-Object { $_.Name }) -join ", "
                    $sampleDb = $group.Group | Select-Object -First 1
                    $documentContent += "|$($group.Name)|$databasesList|$($sampleDb.LocalReplicaRole)|$($sampleDb.SynchronizationState)|`n"
                }
            } else {
                $documentContent += "No Availability Group databases detected on this instance.`n"
            }

            # Set Endpoint properties in document
            $documentContent += "`nh3. Endpoints`n"
            if ($EndpointProperties) {
                $documentContent += "|*Endpoint Type*|*Owner*|*Protocol Type*|*Name*|*Port*|`n"
                foreach ($Endpoint in $EndpointProperties) {
                    $documentContent += "|$($Endpoint.EndpointType)|$($Endpoint.Owner)|$($Endpoint.ProtocolType)|$($Endpoint.Name)|$($Endpoint.Port)|`n"
                }
            } else {
                $documentContent += "No endpoints detected on this instance.`n"
            }

            # Set Certificate properties in document
            $documentContent += "`nh3. Database Certificates`n"
            if ($DatabaseCertificates) {
                $documentContent += "|*Database*|*Name*|*Issuer*|*PrivateKeyEncryptionType*|*Start Date*|*Expiration Date*|`n"
                foreach ($DatabaseCertificate in $DatabaseCertificates) {
                    $documentContent += "|$($DatabaseCertificate.Database)|$($DatabaseCertificate.Name)|$($DatabaseCertificate.Issuer)|$($DatabaseCertificate.PrivateKeyEncryptionType)|$($DatabaseCertificate.StartDate)|$($DatabaseCertificate.ExpirationDate)|`n"
                }
            } else {
                $documentContent += "No database certificates detected on this instance.`n"
            }

            # Set Linked Server properties in document
            $documentContent += "`nh3. Linked Servers`n"
            if ($LinkedServers) {
                $documentContent += "|*Linked Server Name*|*Product Name*|*Provider Name*|*Data Source*|*Impersonate*|*RpcOut*|`n"
                foreach ($server in $LinkedServers) {
                    $documentContent += "|$($server.'Linked Server Name')|$($server.'Product Name')|$($server.'Provider Name')|$($server.'Data Source')|$($server.'Impersonate')|$($server.'RpcOut')|`n"
                }
            } else {
                $documentContent += "No linked servers detected on this instance.`n"
            }

            # Set Database Mail properties in document
            $documentContent += "`nh3. Database Mail Configuration`n"
            if ($mailDetails -and ($mailDetails.Accounts -or $mailDetails.Profiles -or $mailDetails.Servers -or $mailDetails.Config)) {
                # Mail Accounts
                $documentContent += "`nh4. Mail Accounts`n"
                if ($mailDetails.Accounts) {
                    $documentContent += "|*Name*|*Email Address*|*Display Name*|*Reply-To Address*|*SMTP Server*|`n"
                    foreach ($account in $mailDetails.Accounts) {
                        $name = if ([string]::IsNullOrEmpty($account.Name)) { "N/A " } else { $account.Name }
                        $email = if ([string]::IsNullOrEmpty($account.EmailAddress)) { "N/A " } else { $account.EmailAddress }
                        $displayName = if ([string]::IsNullOrEmpty($account.DisplayName)) { "N/A " } else { $account.DisplayName }
                        $replyTo = if ([string]::IsNullOrEmpty($account.ReplyToAddress)) { "N/A " } else { $account.ReplyToAddress }
                        $smtpServer = if ([string]::IsNullOrEmpty($account.SmtpServer)) { "N/A " } else { $account.SmtpServer }
                        $documentContent += "|$name|$email|$displayName|$replyTo|$smtpServer|`n"
                    }
                } else {
                    $documentContent += "No mail accounts configured.`n"
                }

                # Mail Profiles
                $documentContent += "`nh4. Mail Profiles`n"
                if ($mailDetails.Profiles) {
                    $documentContent += "|*Name*|*Description*|*Is Default*|`n"
                    foreach ($profile in $mailDetails.Profiles) {
                        $name = if ([string]::IsNullOrEmpty($profile.Name)) { "N/A " } else { $profile.Name }
                        $description = if ([string]::IsNullOrEmpty($profile.Description)) { "N/A " } else { $profile.Description }
                        $isDefault = if ([string]::IsNullOrEmpty($profile.IsDefault)) { "N/A " } else { $profile.IsDefault }
                        $documentContent += "|$name|$description|$isDefault|`n"
                    }
                } else {
                    $documentContent += "No mail profiles configured.`n"
                }

                # Mail Servers
                $documentContent += "`nh4. Mail Servers`n"
                if ($mailDetails.Servers) {
                    $documentContent += "|*Account Name*|*Server Name*|*Port*|*Enable SSL*|*Authentication Type*|`n"
                    foreach ($server in $mailDetails.Servers) {
                        $accountName = if ([string]::IsNullOrEmpty($server.AccountName)) { "N/A " } else { $server.AccountName }
                        $serverName = if ([string]::IsNullOrEmpty($server.ServerName)) { "N/A " } else { $server.ServerName }
                        $port = if ([string]::IsNullOrEmpty($server.Port)) { "N/A " } else { $server.Port }
                        $enableSsl = if ([string]::IsNullOrEmpty($server.EnableSsl)) { "N/A " } else { $server.EnableSsl }
                        $authType = if ([string]::IsNullOrEmpty($server.AuthenticationType)) { "N/A " } else { $server.AuthenticationType }
                        $documentContent += "|$accountName|$serverName|$port|$enableSsl|$authType|`n"
                    }
                } else {
                    $documentContent += "No mail servers configured.`n"
                }

                # Mail Configuration Settings
                $documentContent += "`nh4. Mail Configuration Settings`n"
                if ($mailDetails.Config) {
                    $documentContent += "|*Name*|*Value*|`n"
                    foreach ($config in $mailDetails.Config) {
                        $documentContent += "|$($config.Name)|$($config.Value)|`n"
                    }
                } else {
                    $documentContent += "No additional configuration settings found.`n"
                }
            } else {
                Write-Log -Message "No Database Mail details to display" -Level DEBUG
                $documentContent += "No Database Mail configuration detected on this instance.`n"
            }

            # Set Replication Publisher properties in document
            $documentContent += "`nh3. Replication Publisher Configuration`n"
            if ($publisherDetails) {
                Write-Log -Message "Adding publisher table with $($publisherDetails.Count) entries" -Level DEBUG
                $documentContent += "|*Publication Name*|*Database*|*Default Snapshot Folder*|*Alternate Snapshot Folder*|`n"
                foreach ($pub in $publisherDetails) {
                    $alternateSnapshotFolder = "$($pub.AlternateSnapshotFolder) "  # Add trailing space
                    $documentContent += "|$($pub.PublicationName)|$($pub.Database)|$($pub.DefaultSnapshotFolder)|$alternateSnapshotFolder|`n"
                }
            } else {
                Write-Log -Message "No publisher details to display" -Level DEBUG
                $documentContent += "No replication publisher configuration detected on this instance.`n"
            }

            # Set Replication Distributor properties in document
            $documentContent += "`nh3. Replication Distributor Configuration`n"
            if ($distributorDetails) {
                Write-Log -Message "Adding distributor table with details" -Level DEBUG
                $documentContent += "|*Publisher*|*Publication*|`n"
                foreach ($pub in $distributorDetails.Publishers) {
                    $documentContent += "|$($pub.PublisherName)|$($pub.PublicationName)|`n"
                }
            } else {
                Write-Log -Message "No distributor details to display" -Level DEBUG
                $documentContent += "No replication distributor configuration detected on this instance.`n"
            }

            # Set Replication Subscription properties in document
            $documentContent += "`nh3. Replication Subscription Configuration`n"
            if ($subscriptionDetails -and $subscriptionDetails.Subscriptions) {
                Write-Log -Message "Adding subscription table with details" -Level DEBUG
                $documentContent += "|*Publisher*|*Publication*|*Subscription Type*|*Database Name*|*Subscription DB Name*|`n"
                foreach ($sub in $subscriptionDetails.Subscriptions) {
                    $documentContent += "|$($sub.PublisherName)|$($sub.PublicationName)|$($sub.SubscriptionType)|$($sub.DatabaseName)|$($sub.SubscriptionDBName)|`n"
                }
            } else {
                Write-Log -Message "No subscription details to display" -Level DEBUG
                $documentContent += "No replication subscription configuration detected on this instance.`n"
            }

            # Generate unique filename with server name and date
            $dateStamp = Get-Date -Format "yyyyMMdd"
            $outputFile = Join-Path -Path $OutputPath -ChildPath "$($instance)_$dateStamp.md"
            $documentContent | Out-File -FilePath $outputFile -Encoding UTF8
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

# Ensure output directories exist
if (-not (Test-Path $OutputPath)) { New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null }
if (-not (Test-Path $ScriptEventLogPath)) { New-Item -Path $ScriptEventLogPath -ItemType Directory -Force | Out-Null }

foreach ($instance in $SQLServerInstances) {
    try {
        # Use stored SqlCredential if no domain account
        if (-not $HasDomainAccount) {
            $SqlCredential = Import-Clixml -Path ".\Credentials\SqlCredentials.xml"
            Write-Log -Message "Successfully imported SQL credentials for $instance" -Level INFO
        }

        Generate-AsBuiltDoc -Instances $instance -OutputPath $OutputPath -DocumentFormat $DocumentFormat
    }
    catch {
        Write-Log -Message "An error occurred during document generation for $($instance): $_" -Level ERROR
    }
}

Write-Log -Message "Finished SQL Server As Built Document Generation" -Level INFO
