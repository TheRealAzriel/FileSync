# This is a program written in collaboration by Brendan Carroll and Chad Cattel for use in copying files.
# This program was written on 2/5/25 and established on 2/25/25.
#
# This program uses a method to copy files from a given source to a destination.
# When this program is openned, it creates the form based on the directives listed below. 
# The method in takes both a source and destination. I use variables for both to allow flexibility of use.
# This program does not copy out empty directories or copy the original directories, it builds the necessary directories.
# This allows the files to retain their date metadata without altering the folders from the source. 
param(
    [string]$mode = "all"
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName "System.Net.Http"

# Running this first will prevent the script from opening multiple instances

#$ScriptName="FileSync"

# Current Script Version for logging purposes
$FileSyncScript = "1.0"

# Get System Information
$UserID = $env:USERNAME
$compName = $env:COMPUTERNAME
$startTime = $(Get-Date)

# Base network path excluding the dynamic folder part
$baseNetworkPath = "\\msms-fs\Deployment\Source"

# this variable is only here to distiguish field lapotps from other machines
$allUserLogin = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp\alluserlogin.lnk"

# API calls
$MSMSProd = "http://msmsprodapp/MSMS.WebApi/api/user/getroles/$UserID"
$MSMSTestC = "http://msmstestcurrentapp/MSMS.WebApi/api/user/getroles/$UserID"
# $MSMSQA = "http://msmsqaapp/MSMS.WebApi/api/user/getroles/$UserID"
# $MSMSTestNext = "http://msmstestnextapp/MSMS.WebApi/api/user/getroles/$UserID"
$script:apiFailed=$false

#File Paths
#Core File paths are at the top with project paths towards the end of this list

#All Files in core
$sourceCORE = "\\msms-fs\Deployment\Source\Core"
$destinationCORE = "C:\"

# Put files onto the Desktop.
$sourceDestkop = "\\msms-fs\Deployment\Source\Desktop"
$destinationDesktop = "C:\Users\$UserID\desktop"

# Project File Path's
$SourceProjectPath = "\\msms-fs\Deployment\Source\Project"
$destinationProjects = "C:\"
$sourceBuild = "\\msms-fs\Deployment\Source\SMS"
$destinationBuild = "C:\"

# Variable to store the path to the CSV file containing the project folder mappings
$csvFilePath = "\\msms-fs\Deployment\Source\Project\ProjectCrossRef.csv"

# Variable to store the path to the text file containing the folder names to download
$foldersToDownloadPath = "C:\SRO\Apps\FileSync\logs\Project\FoldersToDownload-$UserID.txt" 
$foldersToDownloadDirectory = Split-Path $foldersToDownloadPath 

# Custom variables
$sourceCustomNetworkPath = "\\msms-fs\Deployment\Source\Custom\CustomUserFiles"

# Specify the directory containing the text files
$customTextFilesDirectory = "\\msms-fs\Deployment\Source\Custom\UserLists"

# are they a custom user?
$global:customUser = $false

# Specify the username to search for
$usernameToSearch = $UserID

# User files
$sourceUser = "\\msms-fs\Deployment\Source\User\$UserID"
$destinationUser = "C:\"

# Global counters for tracking operations
$global:totalOperations = 0
$global:completedOperations = 0
$beginning = 2

# Define the log file path
$logFilePath = "C:\SRO\Apps\FileSync\logs\FileSync-$UserID-$compName.txt"
$logDirectory = Split-Path $logFilePath

# If the files to download path does not exist, create it.
if (-not (Test-Path -Path $foldersToDownloadDirectory)) {
    New-Item -Path $foldersToDownloadDirectory -ItemType Directory -Force | Out-Null
}

# Check if the log directory exists, if not, create it
if (-not (Test-Path -Path $logDirectory)) {
    New-Item -Path $logDirectory -ItemType Directory -Force | Out-Null
}

# Check if the log file exists, if not, create it
if (-not (Test-Path -Path $logFilePath)) {
    New-Item -Path $logFilePath -ItemType File -Force | Out-Null
}

#######################################################################################################################################################################################

# Writing a header to the log file
Add-Content -Path $logFilePath -Value "===================================================================================================="
Add-Content -Path $logFilePath -Value "[START] Log started on $startTime"
Add-Content -Path $logFilePath -Value "[INFO] Syncing from PC: $compName"
Add-Content -Path $logFilePath -Value "[INFO] Syncing by: $UserID"
Add-Content -Path $logFilePath -Value "[INFO] Script Version: $FileSyncScript"
Add-Content -Path $logFilePath -Value "[INFO] Script Path: $logFilePath `n"

#######################################################################################################################################################################################

# Create the progress bar form
# Form code
$form = New-Object System.Windows.Forms.Form
$form.Text = "FileSync"
$form.Width = 320
$form.Height = 200
$form.StartPosition = 'CenterScreen'
$form.TopMost = $true

# Load the PNG icon and set it as the form's icon
$icoPath = "\\msms-fs\Deployment\Source\FileSync.ico"
if (Test-Path -Path $icoPath) {
    $form.Icon = New-Object System.Drawing.Icon($icoPath)
}
else {
    Write-Host "Icon file not found: $icoPath"
}

# Progress bar code
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Minimum = 0
$progressBar.Maximum = 100
$progressBar.Width = 250
$progressBar.Height = 30
$progressBar.Top = 30
$progressBar.Left = 20
$form.Controls.Add($progressBar)

# Label code
$label = New-Object System.Windows.Forms.Label
$label.Width = 250
$label.Height = 40
$label.Top = 70
$label.Left = 20
$form.Controls.Add($label)

# Phase Label Code
$phaseLabel = New-Object System.Windows.Forms.Label
$phaseLabel.Width = 250
$phaseLabel.Height = 30
$phaseLabel.Top = 10
$phaseLabel.Left = 20
$form.Controls.Add($phaseLabel)

# Label to display the status of operations
$operationStatusLabel = New-Object System.Windows.Forms.Label
$operationStatusLabel.Width = 250
$operationStatusLabel.Height = 20
$operationStatusLabel.Top = 110
$operationStatusLabel.Left = 20
$form.Controls.Add($operationStatusLabel)

# All the above code is then put together to create the UI for the user to know that the program is running

# Method to check VPN connectivity
function CheckNetworkConnectivity {
    param (
        [string]$testPath
    )

    if (-not (Test-Path -Path $testPath)) {
        # Log the error message
        $timestamp = Get-Date -Format "MM-dd-yyyy HH:mm:ss"
        Add-Content -Path $logFilePath -Value "[ERROR] - $timestamp - VPN connection unavailable. Please connect to the VPN."

        # Notify the user
        [System.Windows.Forms.MessageBox]::Show("VPN connection unavailable. Please connect to the VPN.", "VPN Error", 'OK', 'Error')

        # Exit the script
        $mutex.ReleaseMutex()
        $mutex.Dispose()
        exit
    }
}

# Function to set total number of operations
function SetTotalOperations {
    param (
        [int]$count
    )
    $global:totalOperations = $count
    $operationStatusLabel.Text = "Step: $completedOperations of $totalOperations"
}

# Function to copy files with GUI progress bar
function CopyFilesWithGuiProgressBar {
    param (
        [string]$source,
        [string]$destination
    )

    # Ensure the form is visible
    $form.Visible = $true

    # Get a list of files to copy
    $files = Get-ChildItem -Path $source -File -Recurse

    # Get the total number of files
    $totalFiles = $files.Count
    
    # Initialize file counter
    $currentFileIndex = 0

    #Logging Source
    Add-Content -Path $logFilePath -Value "Files Downloaded from: $source `n"

    foreach ($file in $files) {
        # Increment the file counter
        $currentFileIndex++
        
        # Calculate the percentage complete
        $percentComplete = [math]::Floor(($currentFileIndex / $totalFiles) * 100)

        # Update the progress bar and label
        $progressBar.Value = $percentComplete
        #Write-host "File: $file to be copied"
        $label.Text = "Syncing $($file.Name) ($percentComplete%)"

        # Update the form
        $form.Update()
        $form.Refresh()
         
        # Define the destination file path
        $destinationFile = Join-Path $destination ($file.FullName.Substring($source.Length))
 
        # Always ensure the directory exists
        $destinationDirectory = Split-Path $destinationFile
        if (-not (Test-Path -Path $destinationDirectory)) {
            New-Item -Path $destinationDirectory -ItemType Directory | Out-Null
        }
 
        # Attempt to copy the file with error handling
        try {
            # Check if the file should be copied based on timestamps
            if (-not (Test-Path -Path $destinationFile) -or (Get-Item $file.FullName).LastWriteTime -gt (Get-Item $destinationFile).LastWriteTime -or (Get-Item $file.FullName).LastWriteTime -lt (Get-Item $destinationFile).LastWriteTime) {
                # Copy and overwrite the current file if source is newer, older, or does not exist
                Copy-Item -Path $file.FullName -Destination $destinationFile -Force

                $lastModified = (Get-Item $file.FullName).LastWriteTime.ToString("MM-dd-yyyy HH:mm:ss")
                # Log the copied file
                Add-Content -Path $logFilePath -Value "$destinationFile $lastmodified"
            }
        }
        catch {
            CheckNetworkConnectivity -testPath $baseNetworkPath
            # Log the error message
            $errorMessage = "[Error] Failed to copy '$($file.FullName)': $($_.Exception.Message)"
            Add-Content -Path $logFilePath -Value $errorMessage
        }
        
        Start-Sleep -Milliseconds .001

        # Calculate the percentage complete
        #$percentComplete = [math]::Floor(($currentFileIndex / $totalFiles) * 100)
 
        # Update the progress bar and label
        #$progressBar.Value = $percentComplete
        #$label.Text = "Syncing $($file.Name) ($percentComplete%)"
 
        # Update the form
        #$form.Update()
        #$form.Refresh()
    }

    # Update the completed operation count
    # $global:completedOperations++
    # $operationStatusLabel.Text = "Operations: $completedOperations of $totalOperations completed"
    $form.Refresh()
    
    # Update final message
    $label.Text = "File copy completed."
    Add-Content -Path $logFilePath -Value "----------------------------------------------------------------------------------------------------"
}

function updateOperations {
    $global:completedOperations++
    $operationStatusLabel.Text = "Step: $completedOperations of $totalOperations"
    $form.Refresh()
}

function getUserProjectsST {

    # Initilize variables
    $errorLogPath = "C:\SRO\Apps\FileSync\logs\Project\GetProjectID_surveytrak_errorlog.txt" 
        
    # Define the log file path
    $STuserProjectsDirectory = "C:\SRO\Apps\FileSync\logs\Project"
    $STuserProjectsPathProd = "C:\SRO\Apps\FileSync\logs\Project\FileSyncProjects_SurveyTrak_Prod-{0}.txt" -f $UserID
    $STuserProjectsPathTest = "C:\SRO\Apps\FileSync\logs\Project\FileSyncProjects_SurveyTrak_Test-{0}.txt" -f $UserID
    $userProjectsPathST = "C:\SRO\Apps\FileSync\logs\Project\FileSyncProjects_SurveyTrak_All-{0}.txt" -f $UserID 

    # Check if the log directory exists, if not, create it
    if (-not (Test-Path -Path $STuserProjectsDirectory)) {
        New-Item -Path $STuserProjectsDirectory -ItemType Directory -Force | Out-Null
    }

    # Check if the ST prod log files exists, if not, create it
    if (-not (Test-Path -Path $STuserProjectsPathProd)) {
        New-Item -Path $STuserProjectsPathProd -ItemType File -Force | Out-Null
    }

        # Check if the ST test log files exists, if not, create it
    if (-not (Test-Path -Path $STuserProjectsPathTest)) {
        New-Item -Path $STuserProjectsPathTest -ItemType File -Force | Out-Null
    }
    
    try {             
        $rolesProd = @()
        $rolesTest = @()

        if (Test-Path -Path ("C:\SRO\Apps\FileSync\logs\Project\FileSyncProjects_SurveyTrak_Prod-{0}.txt" -f $UserID)) {
            $rolesProd = Get-Content -Path ("C:\SRO\Apps\FileSync\logs\Project\FileSyncProjects_SurveyTrak_Prod-{0}.txt" -f $UserID)
        }

        if (Test-Path -Path ("C:\SRO\Apps\FileSync\logs\Project\FileSyncProjects_SurveyTrak_Test-{0}.txt" -f $UserID)) {
            $rolesTest = Get-Content -Path ("C:\SRO\Apps\FileSync\logs\Project\FileSyncProjects_SurveyTrak_Test-{0}.txt" -f $UserID)
        }

        $projectID = $rolesProd + $rolesTest

        $projectID | Out-File -FilePath $userProjectsPathST 

    }
    catch {
        # Write the error message to the log file
        "{0}: Failed to retrieve data from the SurveyTrak projects file: $_" -f (Get-Date) | Out-File -Append -FilePath $errorLogPath
    }
            
    # Read ProjectCrossRef.csv, parse Project ID and Field name to generate text file to use to identify folders to download
        
    try {
        if (Test-Path -Path $csvFilePath) {
            # Import the CSV file           
            $csvFile = Import-Csv -Path $csvFilePath
            if (Test-Path -Path  $userProjectsPathST) {
                $userProjectsToScan = Get-Content -Path $userProjectsPathST                
                    
                # Create array to store folder names to download
                $folderNames = @()
                        
                # Process each line in the CSV
                foreach ($csvEntry in $csvFile) {
                    $projectIDFromCsv = $csvEntry.'Project ID'.Trim()
                    $folderNameFromCsv = $csvEntry.'Folder Name'.Trim()
        
                    if (-not [string]::IsNullOrEmpty($projectIDFromCsv)) {
                        foreach ($id in $userProjectsToScan) {
                            if ($id -like "*$projectIDFromCsv*") {
                                                
                                if (-not ($folderNames -contains $folderNameFromCsv)) {                                
                                    $folderNames += $folderNameFromCsv
                                }
                            }
                        }
                    }
                }
            }
        
            # Output the folder names to a text file for FileSync            
            $folderNames | Out-File -Append -FilePath $foldersToDownloadPath
        }
        else {
            "{0}: ProjectFolderMappings.csvnot found: $_" -f (Get-Date) | Out-File -Append -FilePath $errorLogPath
        }
        
    }
        catch {
            # Write the error message to the log file
            "{0}: Failed to create foldersToDownload: $_" -f (Get-Date) | Out-File -Append -FilePath $errorLogPath
        }                
}

function getuserprojects {
    param ([string] $api)

    # Get API substring for logging purposes
    # Trim the http://
    $apiLog = $api -split "//" | Select-Object -Last 1 
    # Trim everything after the first part of the api
    $projectLog = $apiLog -split "/" | Select-Object -first 1    
    
    # Define the log file path
    $userProjectsPath = "C:\SRO\Apps\FileSync\logs\Project\FileSyncProjects-{0}-{1}.txt" -f $projectLog, $UserID
    $userProjectsDirectory = Split-Path $userProjectsPath

    # Check if the log directory exists, if not, create it
    if (-not (Test-Path -Path $userProjectsDirectory)) {
        New-Item -Path $userProjectsDirectory -ItemType Directory -Force | Out-Null
    }

    # Check if the log file exists, if not, create it
    if (-not (Test-Path -Path $userProjectsPath)) {
        New-Item -Path $userProjectsPath -ItemType File -Force | Out-Null
    }

    # Initilize variables
    $errorLogPath = "C:\SRO\Apps\FileSync\logs\Project\GetProjectID_{0}_errorlog.txt" -f $projectLog

    try {
        # Create a new instance of HttpClientHandler
        $handler = [System.Net.Http.HttpClientHandler]::new()

        # Optionally use default credentials
        $handler.UseDefaultCredentials = $true

        # Create the HttpClient using the handler
        $client = [System.Net.Http.HttpClient]::new($handler)

        # Set the timeout (e.g., 3 seconds)
        $client.Timeout = [TimeSpan]::FromSeconds(3)

        # Set User-Agent header if needed
        $client.DefaultRequestHeaders.UserAgent.ParseAdd("MyPowerShellClient/1.0")

        # URI for the request
        $uri = $api

        try {
            # Synchronously send a GET request
            $task = $client.GetAsync($uri)
            $task.Wait()

            # Check if the request was successful
            if ($task.Result.IsSuccessStatusCode) {
                # Clear previous project list and start fresh on a successful response
                # Clearing the content here preserves the last successfull api sync
                Clear-Content -Path $userProjectsPath

                $result = $task.Result.Content.ReadAsStringAsync().Result
                #Write-Host "Task completed successfully. Result:"
                #Write-Host $result
            }
            else {
                # If request was not successful display the following error:
                Write-Host "Error: $($task.Result.StatusCode)"
                "[Error] $($task.Result.StatusCode)" | Out-File -Append -FilePath $errorLogPath
                $script:apiFailed=$true
            }
        }
        catch [System.Threading.Tasks.TaskCanceledException] {
            # Handle any cancelations from an api taking too long
            Write-Host "Exception: $projectLog was unable to resolve within timeout."
            "[Exception] $($innerException.Message) $projectLog was unable to resolve within timeout." | Out-File -Append -FilePath $errorLogPath
            $script:apiFailed=$true
        }
        catch [System.AggregateException] {
            # Handle aggregate exceptions from asynchronous operations
            foreach ($innerException in $_.Exception.InnerExceptions) {
                Write-Host "Exception: $($innerException.Message)"
                "[Exception] $($innerException.Message)" | Out-File -Append -FilePath $errorLogPath
            }
            $script:apiFailed=$true
        }
        catch {
            Write-Host "An unexpected exception occurred: $($_.Exception.Message)"
            "[Exception] An unexpected exception occurred: $($_.Exception.Message)" | Out-File -Append -FilePath $errorLogPath
            $script:apiFailed=$true
        }

        # Convert the JSON response to a PowerShell object
        $roles = ConvertFrom-Json $result

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
        #$jsonOutput = $projectID | ConvertTo-Json -Compress

        # Write the JSON output to a file
        $projectID | Out-File -FilePath $userProjectsPath

    }
    catch {
        # Write the error message to the log file
        "[Error] {0}: Failed to retrieve data from the Web API: $_" -f (Get-Date) | Out-File -Append -FilePath $errorLogPath
    }
    
    # Read ProjectCrossRef.csv, parse Project ID and Field name to generate text file to use to identify folders to download

    try {
        if (Test-Path -Path $csvFilePath) {
            # Import the CSV file           
            $csvFile = Import-Csv -Path $csvFilePath
            if (Test-Path -Path  $userProjectsPath) {
                $userProjectsToScan = Get-Content -Path $userProjectsPath                
            
                # Create array to store folder names to download
                $folderNames = @()
                
                # Process each line in the CSV
                foreach ($csvEntry in $csvFile) {
                    $projectIDFromCsv = $csvEntry.'Project ID'.Trim()
                    $folderNameFromCsv = $csvEntry.'Folder Name'.Trim()

                    if (-not [string]::IsNullOrEmpty($projectIDFromCsv)) {
                        foreach ($id in $userProjectsToScan) {
                            if ($id -like "*$projectIDFromCsv*") {
                                        
                                if (-not ($folderNames -contains $folderNameFromCsv)) {                                
                                    $folderNames += $folderNameFromCsv
                                }
                            }
                        }
                    }
                }
            }

            # Output the folder names to a text file for FileSync            
            $folderNames | Out-File -Append -FilePath $foldersToDownloadPath
        }
        else {
            "[Error] {0}: ProjectFolderMappings.csvnot found: $_" -f (Get-Date) | Out-File -Append -FilePath $errorLogPath
        }

    }
    catch {
        # Write the error message to the log file
        "[Error] {0}: Failed to create foldersToDownload: $_" -f (Get-Date) | Out-File -Append -FilePath $errorLogPath
    }
    
    # Cleanup
    $client.Dispose()
}
function CoreFiles {
    #Invoke the method with a source and destination. I use variables to avoid having to write the same destination down here incase they change later down the line.

    $phaseLabel.Text = "Downloading Core Files"
    updateOperations

    if (Test-Path -path $allUserLogin) {
        CopyFilesWithGuiProgressBar -source $sourceDestkop -destination $destinationDesktop
    }
    
    CopyFilesWithGuiProgressBar -source $sourceCORE -destination $destinationCORE

    # Start-Sleep -Seconds 2
}

function projectFilesMode {
    param ([String] $modeSwitch)
    
    # Check if the project file list exists and clear it if it does
    if (Test-Path -Path $foldersToDownloadPath) {
        Clear-Content -Path $foldersToDownloadPath
    }

    switch ($modeSwitch) {
        "SurveyTrak" { 
            "Sync all SurveyTrak files"
            getUserProjectsST
            ; break    
        }
        "MSMS" {
            "Sync all MSMS files"
            getUserProjects -api $MSMSProd
            getuserprojects -api $MSMSTestC
            #getuserprojects -api $MSMSQA
            #getuserprojects -api $MSMSTestNext; 
            ; break
        }
        "build" {
            "Sync all build files"
            buildFiles
            ; break 
        }
        default {
            "Sync all project files"
            getUserProjectsST
            getUserProjects -api $MSMSProd
            getuserprojects -api $MSMSTestC
            buildFiles
            ; break 
        }
    }
}

Function ProjectFiles {
    
    $phaseLabel.Text = "Downloading Project Files"
    updateOperations

    $projectFoldersFile = $foldersToDownloadPath 

    # Attempt to read the folder names from the text file
    try {
        $projectFolderNames = Get-Content -Path $projectFoldersFile
    }
    catch {
        $mutex.ReleaseMutex()
        $mutex.Dispose()
        # Log if the folder does not exist
        $timestamp = Get-Date -Format "MM-dd-yyyy HH:mm:ss"
        Add-Content -Path $logFilePath -Value "[Error] $timestamp - A projectFoldersFile was unreadable or not created. Please check C:\SRO\Apps\FileSync\logs\Project\FoldersToDownload-$UserID.txt"
        Add-Content -Path $logFilePath -Value "----------------------------------------------------------------------------------------------------"
        Write-Host "Error reading folders file: $($_.Exception.Message)"
        exit
    }

    # Copy files from each network source folder to the local destination
    foreach ($projectFolderName in $projectFolderNames) {
        # Construct the full network path
        $sourceProjectDesktop = Join-Path (Join-Path $SourceProjectPath $projectFolderName) "Desktop"
        $sourceProjectRoot = Join-Path (Join-Path $SourceProjectPath $projectFolderName) "root"

        if (($projectFolderName -eq "SMS") -and ((Test-Path -path $allUserLogin))) {
            CopyFilesWithGuiProgressBar -source $sourceBuild -destination $destinationBuild
        }
        elseif (($projectFolderName -eq "SMS") -and (-not (Test-Path -path $allUserLogin))) {
            # Skip
        }
        else {
            # Copy files from the constructed path to the destination
            CopyFilesWithGuiProgressBar -source $sourceProjectRoot -destination $destinationProjects

            if (Test-Path -path $allUserLogin) {
                CopyFilesWithGuiProgressBar -source $sourceProjectDesktop -destination $destinationDesktop
            }
        }

    }
        
    # Start-Sleep -Seconds 2
}

function isCustomUser {
    # Get all text files from the specified directory
    $customTextFiles = Get-ChildItem -Path $customTextFilesDirectory -Filter *.txt
    # Process each text file in the directory
    foreach ($customtextFile in $customTextFiles) {
        # Read the contents of the text file
        $fileContents = Get-Content -Path $customTextFile.FullName
        # Check if the username is in the file contents
        if ($fileContents -match $usernameToSearch) {
            # If the username is in a list, set the boolean to true (default false)
            $global:customUser = $true
            # This write host is only here to prevent a false positive from the above line
            Write-Host $customUser
            # return out of the function since we now know that the user is on atleast one custom list
            return
        }
    }
}

function CustomFiles {
    # If the user is on a custom list, run the following code, otherwise skip.
    if ($global:customUser -eq $true) {
        $phaseLabel.Text = "Downloading Custom Files"
        updateOperations
        # Get all text files from the specified directory
        $customTextFiles = Get-ChildItem -Path $customTextFilesDirectory -Filter *.txt
        # Process each text file in the directory
        foreach ($customtextFile in $customTextFiles) {
            # Read the contents of the text file
            $fileContents = Get-Content -Path $customTextFile.FullName
    
            # Check if the username is in the file contents
            if ($fileContents -match $usernameToSearch) {
                # Use the text file's name (without extension) as part of the path
                $CustomfolderName = [System.IO.Path]::GetFileNameWithoutExtension($customTextFile.Name)
            
                # Construct the full network path
                $sourceCustomDesktop = Join-Path (Join-Path $sourceCustomNetworkPath $CustomfolderName) "Desktop"
                $sourceCustomPath = Join-Path (Join-Path $sourceCustomNetworkPath $CustomfolderName) "root"
            
                if ((Test-Path -path $sourceCustomDesktop) -and (Test-Path -path $allUserLogin)) {
                    # Perform the copy operation
                    CopyFilesWithGuiProgressBar -source $sourceCustomDesktop -destination $destinationDesktop
                }
                if (Test-Path -path $sourceCustomPath) {
                    # Perform the copy operation
                    CopyFilesWithGuiProgressBar -source $sourceCustomPath -destination $destinationProjects
                }
                else {
                    # Log if the folder does not exist
                    $timestamp = Get-Date -Format "MM-dd-yyyy HH:mm:ss"
                    Add-Content -Path $logFilePath -Value "[Error] $timestamp - A custom folder does not exist for text file: $customTextFile"
                    Add-Content -Path $logFilePath -Value "----------------------------------------------------------------------------------------------------"
                }
            }
        }
        # Start-Sleep -Seconds 2
    }
}

function UserFiles {
    if (Test-Path $sourceUser) {
        $phaseLabel.Text = "Downloading User Files"
        updateOperations
        $form.Refresh()

        $sourceUserRoot = Join-Path $sourceUser "Root"
        $sourceUserDesktop = Join-Path $sourceUser "Desktop"

        CopyFilesWithGuiProgressBar -source $sourceUserRoot -destination $destinationUser 
        CopyFilesWithGuiProgressBar -source $sourceUserDesktop -destination $destinationDesktop
    } 
}

function buildFiles {
    # Check if the project file list exists and clear it if it does
    if ((Test-Path -Path $foldersToDownloadPath) -and ($mode -eq "build")) {
        Clear-Content -Path $foldersToDownloadPath
    }
    "SMS" | Out-File -Append -FilePath $foldersToDownloadPath
}

function finishProgram {
    Start-Sleep -Seconds 2

    $endTime = $(Get-Date)
    $timeDiff = $endTime - $startTime
    Add-Content -Path $logFilePath -Value "[Finish] Log Finished in $timeDiff at $EndTime"
    Add-Content -Path $logFilePath -Value "===================================================================================================="

    $mutex.ReleaseMutex()
    $mutex.Dispose()

    # Optionally close the form
    $form.Close()

    #Write-Host $script:apiFailed

    if($script:apiFailed -eq $true){
        [System.Windows.Forms.MessageBox]::Show("MSMS file sync has failed. Please try again in 15 minutes.", "API Error", 'OK', 'Error')
    }

    #& "C:\SRO\SRUD\Send Receive Upload Download.exe"
}

######################################################################################################################################################################################
# Program Start

# create the mutex
$mutexName = "Global\FileSync"

# Discover which project files mode will be run.
projectFilesMode -modeSwitch $mode
# Discover if the user is on a custom list
isCustomUser

# Declare the variable before using [ref]
$createdNew = $false
$createdNewRef = [ref] $createdNew

# Try to create the mutex
$mutex = [System.Threading.Mutex]::new($true, $mutexName, $createdNewRef)

try {
    if (-not $createdNewRef.Value) {
        #Write-Host "Another instance of the script is already running. Exiting."
        [System.Windows.Forms.MessageBox]::Show("Another instance of the script is already running. Please Wait...", "Instance Error", 'OK', 'Error')
        exit
    }

    # Check if connected to the network path
    CheckNetworkConnectivity -testPath $baseNetworkPath

    # Set the number of operations you plan to run
    SetTotalOperations -count $beginning

    # If user has a user folder add one more progress bar
    if (Test-Path $sourceUser) {
        $beginning ++
        SetTotalOperations -count $beginning
    }

    # If the user is on a custom list, increment the total opperations
    if ($global:customUser -eq $true) {
        $beginning ++
        SetTotalOperations -count $beginning
    }

    # Show the form by invoking the method multiple times
    $form.Show()

    # Each Operation in order: 

    CoreFiles

    ProjectFiles

    CustomFiles

    UserFiles

    finishProgram
}
finally {
    if ($createdNew.Value) {
        # Release the mutex when done
        $mutex.ReleaseMutex()
        $mutex.Dispose()
    }
}

# Program End
#########################################################################################################################################################################################