# Initilize variables
$UserID = $env:USERNAME
$localPath = $PSScriptRoot
$errorLogPath = Join-Path -Path $localPath -ChildPath "GetProjectID_errorlog.txt"

# Define the URI for the Web API
$uri = "http://msmstestcurrentapp/MSMS.WebApi/api/user/getroles/$UserID"

# Create a new WebClient object and set it to use the default credentials
$webClient = New-Object System.Net.WebClient
$webClient.UseDefaultCredentials = $true

try {
    # Download the JSON response from the Web API
    $response = $webClient.DownloadString($uri)

    # Convert the JSON response to a PowerShell object
    $roles = ConvertFrom-Json $response

    # Initialize array to store project IDs
    $projectID = @()

    foreach ($project in $roles) {
        $tempProjectIDs = $project.ProjectId.Id

        # Check to see if ID is in the array, add if not
        if (-not ($projectID -contains $tempProjectIDs)) {
        $projectID += $tempProjectIDs
       }
    } 

    # Convert the array of project IDs to JSON
    $jsonOutput = $projectID | ConvertTo-Json -Compress

    # Write the JSON output to a file
    $jsonOutput | Set-Content -Path "$localPath\$userID-projects.json"

} catch {
    # Write the error message to the log file
    "{0}: Failed to retrieve data from the Web API: $_" -f (Get-Date) | Out-File -Append -FilePath $localPath
}
#Write-Output "The value of projectID is:"
#$projectID | ForEach-Object { Write-Output $_ }
"{0}: Failed to retrieve data from the Web API: $_" -f (Get-Date) | Out-File -Append -FilePath $localPath