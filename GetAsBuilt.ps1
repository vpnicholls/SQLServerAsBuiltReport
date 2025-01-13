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
    [string]$DocumentFormat = "markdown",
    [string]$ScriptEventLogPath = "$env:USERPROFILE\Documents\AsBuiltDocsLogs"
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

# Function to get server configuration using dbatools
function Get-SQLServerConfig {
    param (
        [Parameter(Mandatory=$true)]
        [string]$InstanceName
    )

    try {
        Write-Log -Message "Retrieving configuration for $InstanceName" -Level INFO
        $serverInfo = Get-DbaInstanceProperty -SqlInstance $InstanceName -ErrorAction Stop
        return @{
            'Version' = $serverInfo | Where-Object { $_.Name -eq 'VersionString' } | Select-Object -ExpandProperty Value
            'Edition' = $serverInfo | Where-Object { $_.Name -eq 'Edition' } | Select-Object -ExpandProperty Value
            'Collation' = $serverInfo | Where-Object { $_.Name -eq 'Collation' } | Select-Object -ExpandProperty Value
            'IsClustered' = $serverInfo | Where-Object { $_.Name -eq 'IsClustered' } | Select-Object -ExpandProperty Value
        }
    }
    catch {
        Write-Log -Message "Failed to retrieve configuration for $InstanceName. Error: $_" -Level ERROR
        return $null
    }
}

function Add-DatabaseInfoToDoc {
    param (
        [Parameter(Mandatory=$true)]
        [Array]$Databases,
        [Parameter(Mandatory=$true)]
        [ref]$DocumentContent
    )

    foreach ($db in $Databases) {
        $DocumentContent.Value += "`n#### $($db.Name)`n"
        $DocumentContent.Value += "| Property | Value |`n"
        $DocumentContent.Value += "| --- | --- |`n"
        foreach ($prop in $db.Keys | Where-Object { $_ -ne "Name" }) {
            $DocumentContent.Value += "| $prop | $($db[$prop]) |`n"
        }
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
                'SizeMB' = [math]::Round($totalSizeBytes / 1048576, 2)
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
        
        $config = Get-SQLServerConfig -InstanceName $instance
        $databases = Get-SQLDatabases -InstanceName $instance

        if ($DocumentFormat -eq "markdown") {
            $documentContent = "## SQL Server: $instance`n"
            $documentContent += "### Configuration`n"
            $documentContent += "| Key | Value |`n"
            $documentContent += "| --- | --- |`n"
            foreach ($key in $config.Keys) {
                $documentContent += "| $key | $($config[$key]) |`n"
            }

            $documentContent += "`n### System Databases`n"
            foreach ($db in $databases.SystemDatabases) {
                $documentContent += "`n#### $($db.Name)`n"
                $documentContent += "| Property | Value |`n"
                $documentContent += "| --- | --- |`n"

                # Define the order of properties for system databases, excluding CreationDate and CompatibilityLevel
                $orderedPropertiesForSystem = @('Owner', 'RecoveryModel', 'IsAutoCreateStatisticsEnabled', 'LastBackupDate', 'IsReadCommittedSnapshotOn', 'AutoShrink', 'Status', 'SizeMB', 'AutoClose', 'IsAutoUpdateStatisticsEnabled')
                
                foreach ($prop in $orderedPropertiesForSystem) {
                    if ($db.ContainsKey($prop)) {
                        $documentContent += "| $prop | $($db[$prop]) |`n"
                    }
                }
            }

            $documentContent += "`n### User Databases`n"
            foreach ($db in $databases.UserDatabases) {
                $documentContent += "`n#### $($db.Name)`n"
                $documentContent += "| Property | Value |`n"
                $documentContent += "| --- | --- |`n"

                # Use a different array for user databases if needed, here we're keeping all properties for example
                $orderedPropertiesForUser = @('Owner', 'RecoveryModel', 'IsAutoCreateStatisticsEnabled', 'LastBackupDate', 'IsReadCommittedSnapshotOn', 'AutoShrink', 'Status', 'SizeMB', 'CreationDate', 'AutoClose', 'CompatibilityLevel', 'IsAutoUpdateStatisticsEnabled')
                
                foreach ($prop in $orderedPropertiesForUser) {
                    if ($db.ContainsKey($prop)) {
                        $documentContent += "| $prop | $($db[$prop]) |`n"
                    }
                }
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