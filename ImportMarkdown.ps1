# Define variables
$ConfluenceEmail = "<Set your Confluence account email address here>"
$ConfluenceAPIToken = '<Set your API token here>'
$PageID = "<Set your Confluence page ID here>"
$MarkdownFile = "<Set the location of your *.md markdown file here>"
$Title = "Set the title of yor Confluence document here"

# Use basic authentication
$base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f $ConfluenceEmail,$ConfluenceAPIToken)))

$content = Get-Content -Path $MarkdownFile -Raw -Encoding UTF8
$content = $content.Trim()

$Body = @{
    representation = "wiki"
    value = $content
}

# GET the current page details to get the version number
$getPageOptions = @{
    Uri = "<Set your Confluence page address here>$($PageID)"
    Headers = @{
        'Accept' = 'application/json'
        'Authorization' = ('Basic {0}' -f $base64AuthInfo)
    }
    Method = 'GET'
}

try {
    $pageResponse = Invoke-RestMethod @getPageOptions
    Write-Output "Page retrieved successfully:"
    Write-Output $pageResponse
} catch {
    Write-Error "Failed to retrieve page: $_"
    exit  # Exit script if we can't get the page details
}

# Increment the version number
$pageVersion = [int]$pageResponse.version.number + 1
Write-Verbose "The old version number is $($pageResponse.version.number)."
Write-Verbose "The new version number is $($pageVersion)."

# Prepare for the PUT request
$putPageOptions = @{
    Uri = "<Set your Confluence page address here>/$PageID"
    Headers = @{
        'Accept' = 'application/json'
        'Content-Type' = 'application/json'
        'Authorization' = ('Basic {0}' -f $base64AuthInfo)
    }
    Method = 'PUT'
    Body = @{
        id = $PageID
        status = "current"
        title = $Title
        body = $Body
        version = @{
            number = $pageVersion
        }
    } | ConvertTo-Json
}

# Execute the PUT request
try {
    $updateResponse = Invoke-RestMethod @putPageOptions
    Write-Output "Page updated successfully."
} catch {
    Write-Error "Failed to update the page: $_"
}
