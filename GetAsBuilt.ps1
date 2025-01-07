<#
.SYNOPSIS
    Generates an "As Built" document for each SQL Server instance provided.

.DESCRIPTION
    This script compiles detailed information about one or more SQL Server instances into individual structured documents, 
    which can be used for documentation, auditing, or migration planning. It uses PowerShell to query SQL Server 
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

param (
    [Parameter(Mandatory=$true)]
    [string[]]$SQLServerInstances,
    [string]$OutputPath = "$env:USERPROFILE\Documents\AsBuiltDocs",
    [string]$DocumentFormat = "markdown",
    [string]$ScriptEventLogPath = "$env:USERPROFILE\Documents\Logs"
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

# Function to get server configuration
function Get-SQLServerConfig {
    param (
        [Parameter(Mandatory=$true)]
        [string]$InstanceName
    )

    try {
        Write-Log -Message "Retrieving configuration for $InstanceName" -Level INFO
        $server = New-Object Microsoft.SqlServer.Management.Smo.Server($InstanceName)
        $config = @{
            'Version' = $server.Version
            'Edition' = $server.Edition
            'Collation' = $server.Collation
            'IsClustered' = $server.IsClustered
        }
        return $config
    }
    catch {
        Write-Log -Message "Failed to retrieve configuration for $InstanceName. Error: $_" -Level ERROR
        return $null
    }
}

# Function to get databases details
function Get-SQLDatabases {
    param (
        [Parameter(Mandatory=$true)]
        [string]$InstanceName
    )

    try {
        Write-Log -Message "Retrieving database details for $InstanceName" -Level INFO
        $server = New-Object Microsoft.SqlServer.Management.Smo.Server($InstanceName)
        $databases = @()
        foreach ($db in $server.Databases) {
            $databases += @{
                'Name' = $db.Name
                'SizeMB' = [math]::Round($db.Size / 1MB, 2)
                'Status' = $db.Status
                'RecoveryModel' = $db.RecoveryModel
            }
        }
        return $databases
    }
    catch {
        Write-Log -Message "Failed to retrieve databases for $InstanceName. Error: $_" -Level ERROR
        return @()
    }
}

# Main function to generate the As Built Document
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
            $documentContent += "`n### Databases`n"
            $documentContent += "| Name | Size (MB) | Status | Recovery Model |`n"
            $documentContent += "| --- | --- | --- | --- |`n"
            foreach ($db in $databases) {
                $documentContent += "| $($db.Name) | $($db.SizeMB) | $($db.Status) | $($db.RecoveryModel) |`n"
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

# Main execution
Write-Log -Message "Starting SQL Server As Built Document Generation" -Level INFO
Generate-AsBuiltDoc -Instances $SQLServerInstances -OutputPath $OutputPath -DocumentFormat $DocumentFormat
Write-Log -Message "Finished SQL Server As Built Document Generation" -Level INFO