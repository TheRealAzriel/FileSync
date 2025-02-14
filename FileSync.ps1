# This is a program written by Brendan Carroll for use in copying files.
# This program was written on 2/5/25.
#
# This program uses a method to copy files from a given source to a destination.
# When this program is openned, it creates the form based on the directives listed below. 
# The method in takes both a source and destination. I use variables for both to allow flexibility of use.
# This program does not copy out empty directories or copy the original directories, it builds the necessary directories.
# This allows the files to retain their date metadata without altering the folders from the source. 
param(
    [string]$mode = "MSMS"
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Running this first will prevent the script from opening multiple instances

#$ScriptName="FileSync"

# Current Script Version for logging purposes
$FileSyncScript="1.0"

$UserID = $env:USERNAME
$compName=$env:COMPUTERNAME
$startTime=$(Get-Date)

# Base network path excluding the dynamic folder part
$baseNetworkPath = "\\msms-fs\Deployment\Source"

# this variable is only here to distiguish field lapotps from other machines
$allUserLogin="C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp\alluserlogin.lnk"

# Custom variables
$sourceCustomNetworkPath="\\msms-fs\Deployment\Source\Custom\CustomUserFiles"

# Specify the directory containing the text files
$customTextFilesDirectory = "\\msms-fs\Deployment\Source\Custom\UserLists"

# Specify the username to search for
$usernameToSearch = $UserID

# API calls
# Define the URI for the Web API
#$uri = "http://msmstestcurrentapp/MSMS.WebApi/api/user/getroles/$UserID"
#$uri = "http://msmsprodapp/MSMS.WebApi/api/user/getroles/$UserID"
$MSMSProd = "http://msmsprodapp/MSMS.WebApi/api/user/getroles/$UserID"
$MSMSTestC = "http://msmstestcurrentapp/MSMS.WebApi/api/user/getroles/$UserID"
$MSMSQA = "http://msmsqaapp/MSMS.WebApi/api/user/getroles/$UserID"
$MSMSTestNext = "http://msmstestnextapp/MSMS.WebApi/api/user/getroles/$UserID"

#File Paths
#Core File paths are at the top with project paths towards the end of this list

# Project File Path's
$SourceProjectPath = "\\msms-fs\Deployment\Source\Project"
$destinationProjects="C:\"

#All Files in core
$sourceCORE="\\msms-fs\Deployment\Source\Core"
$destinationCORE="C:\"

# Put files onto the Desktop.
$sourceDestkop = "\\msms-fs\Deployment\Source\Desktop"
$destinationDesktop = "C:\Users\$UserID\desktop"

# User files
$sourceUser = "\\msms-fs\Deployment\Source\User\$UserID"
$destinationUser = "C:\"

# Global counters for tracking operations
$global:totalOperations = 0
$global:completedOperations = 0
$beginning=3

# Define the log file path
$logFilePath = "C:\SRO\Apps\FileSync\logs\FileSync-$UserID-$compName.txt"
$logDirectory = Split-Path $logFilePath

# Check if the log directory exists, if not, create it
if (-not (Test-Path -Path $logDirectory)) {
    New-Item -Path $logDirectory -ItemType Directory -Force | Out-Null
}

# Check if the log file exists, if not, create it
if (-not (Test-Path -Path $logFilePath)) {
    New-Item -Path $logFilePath -ItemType File -Force | Out-Null
}

# Writing a header to the log file
Add-Content -Path $logFilePath -Value "===================================================================================================="
Add-Content -Path $logFilePath -Value "[START] Log started on $startTime"
Add-Content -Path $logFilePath -Value "[INFO] Syncing from PC: $compName"
Add-Content -Path $logFilePath -Value "[INFO] Syncing by: $UserID"
Add-Content -Path $logFilePath -Value "[INFO] Script Version: $FileSyncScript"
Add-Content -Path $logFilePath -Value "[INFO] Script Path: $logFilePath `n"

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
} else {
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
            if (-not (Test-Path -Path $destinationFile) -or (Get-Item $file.FullName).LastWriteTime -gt (Get-Item $destinationFile).LastWriteTime -or (Get-Item $file.FullName).LastWriteTime -lt (Get-Item $destinationFile).LastWriteTime)
                {
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

function updateOperations{
    $global:completedOperations++
    $operationStatusLabel.Text = "Step: $completedOperations of $totalOperations"
    $form.Refresh()
}

# Variable to store the path to the JSON file containing the project folder mappings
$jsonfilePath = "\\msms-fs\Deployment\Source\Project\ProjectFolderMappings.json" 

# Variable to store the path to the text file containing the folder names to download
$foldersToDownloadPath = "C:\SRO\Apps\FileSync\logs\Project\FoldersToDownload-$UserID.txt"  

# Check if the file exists and clear it if it does
if (Test-Path -Path $foldersToDownloadPath) {
    Clear-Content -Path $foldersToDownloadPath
}
function getuserprojects{
    param ([string] $api)

    # Get API substring for logging purposes
    # Trim the http://
    $apiLog = $api -split "//" | Select-Object -Last 1 
    # Trim everything after the first part of the api
    $projectLog=$apiLog -split "/" | Select-Object -first 1    
    
    # Define the log file path
    $userProjectsPath = "C:\SRO\Apps\FileSync\logs\Project\FileSyncProjects-{0}-{1} .txt" -f $projectLog, $UserID
    $userProjectsDirectory = Split-Path $userProjectsPath

    # Check if the log directory exists, if not, create it
    if (-not (Test-Path -Path $userProjectsDirectory)) {
        New-Item -Path $userProjectsDirectory -ItemType Directory -Force | Out-Null
    }

    # Check if the log file exists, if not, create it
    if (-not (Test-Path -Path $userProjectsPath)) {
        New-Item -Path $userProjectsPath -ItemType File -Force | Out-Null
    }

    # Clear previous project list and start fresh.
    Clear-Content -Path $userProjectsPath

    # Initilize variables
    $localPath = $userProjectsPath
    $errorLogPath = Join-Path -Path $localPath -ChildPath "GetProjectID_errorlog.txt"

    # Create a new WebClient object and set it to use the default credentials
    $webClient = New-Object System.Net.WebClient
    $webClient.UseDefaultCredentials = $true

    try {
        # Download the JSON response from the Web API
        $response = $webClient.DownloadString($api)

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
        #$jsonOutput = $projectID | ConvertTo-Json -Compress

        # Write the JSON output to a file
        $projectID | Out-File -FilePath $localPath

    } catch {
        # Write the error message to the log file
        "{0}: Failed to retrieve data from the Web API: $_" -f (Get-Date) | Out-File -Append -FilePath $errorLogPath
    }
    
    # Generate text file to use to identify folders to download

    try {
        if (Test-Path -Path $jsonfilePath) {
            # Import the JSON file           
            $jsonFile = Get-Content -Path $jsonfilePath | ConvertFrom-Json

            # Create array to store folder names to download
            $folderNames = @()

            foreach ($jsonEntry in $jsonFile) {
                #if ($projectID -contains $jsonEntry.ProjectID) {
                if ($projectID -like "*$($jsonEntry.ProjectID)*") {
                    if (-not ($folderNames -contains $jsonEntry.FolderName)) {
                        $folderNames += $jsonEntry.FolderName
                    }
                }
            }

            # Output the folder names to a text file for FileSync            
            $folderNames | Out-File -Append -FilePath $foldersToDownloadPath
        } else {
            "{0}: ProjectFolderMappings.json not found: $_" -f (Get-Date) | Out-File -Append -FilePath $errorLogPath
        }

    } catch {
        # Write the error message to the log file
        "{0}: Failed to create foldersToDownload: $_" -f (Get-Date) | Out-File -Append -FilePath $errorLogPath
    }
}

function CoreFiles {
    #Invoke the method with a source and destination. I use variables to avoid having to write the same destination down here incase they change later down the line.

    $phaseLabel.Text = "Downloading Core Files"
    updateOperations

    if (Test-Path -path $allUserLogin){
        CopyFilesWithGuiProgressBar -source $sourceDestkop -destination $destinationDesktop
    }
    
    CopyFilesWithGuiProgressBar -source $sourceCORE -destination $destinationCORE

    # Start-Sleep -Seconds 2
}

function projectFilesMode {
    param 
    (
        [String] $modeSwitch
    )
    
    switch ($modeSwitch){
        "SurveyTrak" {"Sync all SurveyTrak files"; break} 
        "MSMS" 
        {
            "Synch all MSMS files"
            getUserProjects -api $MSMSProd
            getuserprojects -api $MSMSTestC
            getuserprojects -api $MSMSQA
            getuserprojects -api $MSMSTestNext; 
            break
        }
        "build"{"Sync all build files"; break}
        default {"Sync all project files"; break}
    }


}

Function ProjectFiles {
    $phaseLabel.Text = "Downloading Project Files"
    updateOperations

        # Copy files from each network source folder to the local destination
        foreach ($folderName in $folderNames) {
            # Construct the full network path
            $sourceProjectDesktop = Join-Path (Join-Path $SourceProjectPath $folderName) "Desktop"
            $sourceProjectRoot = Join-Path (Join-Path $SourceProjectPath $folderName) "root"

            # Copy files from the constructed path to the destination
            CopyFilesWithGuiProgressBar -source $sourceProjectRoot -destination $destinationProjects
            CopyFilesWithGuiProgressBar -source $sourceProjectDesktop -destination $destinationDesktop
        }
        
        # Start-Sleep -Seconds 2
}

function CustomFiles {
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
            
            if((Test-Path -path $sourceCustomDesktop) -and (Test-Path -path $allUserLogin)){
                # Perform the copy operation
                CopyFilesWithGuiProgressBar -source $sourceCustomDesktop -destination $destinationDesktop
            }
            if(Test-Path -path $sourceCustomPath){
                # Perform the copy operation
                CopyFilesWithGuiProgressBar -source $sourceCustomPath -destination $destinationProjects
            }else{
                # Log if the folder does not exist
                $timestamp = Get-Date -Format "MM-dd-yyyy HH:mm:ss"
                Add-Content -Path $logFilePath -Value "[Error] $timestamp - A custom folder does not exist for text file: $customTextFile"
                Add-Content -Path $logFilePath -Value "----------------------------------------------------------------------------------------------------"
            }
        }
    }
    # Start-Sleep -Seconds 2
}

function UserFiles{
    if (Test-Path $sourceUser){
        $phaseLabel.Text = "Downloading User Files"
        updateOperations
        $form.Refresh()
        CopyFilesWithGuiProgressBar -source $sourceUser -destination $destinationUser
        
    }
    
}

function finishProgram {
    Start-Sleep -Seconds 2

    $endTime=$(Get-Date)
    $timeDiff=$endTime-$startTime
    Add-Content -Path $logFilePath -Value "[Finish] Log Finished in $timeDiff at $EndTime"
    Add-Content -Path $logFilePath -Value "===================================================================================================="

    $mutex.ReleaseMutex()
    $mutex.Dispose()
    # Optionally close the form
    $form.Close()

    #& "C:\SRO\SRUD\Send Receive Upload Download.exe"
}

######################################################################################################################################################################################
# Program Start

$mutexName = "Global\FileSync"

# Declare the variable before using [ref]
$createdNew = $false
$createdNewRef = [ref] $createdNew

# Try to create the mutex
$mutex = [System.Threading.Mutex]::new($true, $mutexName, $createdNewRef)

try {
    if (-not $createdNewRef.Value)
    {
        #Write-Host "Another instance of the script is already running. Exiting."
        [System.Windows.Forms.MessageBox]::Show("Another instance of the script is already running. Please Wait...", "Instance Error", 'OK', 'Error')
        exit
    }

    # Check if connected to the network path
    CheckNetworkConnectivity -testPath $baseNetworkPath

    # Set the number of operations you plan to run
    SetTotalOperations -count $beginning

    # If user has a user folder add one more progress bar
    if (Test-Path $sourceUser){
        $beginning ++
        SetTotalOperations -count $beginning
    }

    # Show the form by invoking the method multiple times
    $form.Show()

    # Each Operation in order: 

    CoreFiles

    ProjectFiles
    projectFilesMode -modeSwitch $mode

    CustomFiles

    UserFiles

    finishProgram

}finally {
    if ($createdNew.Value) 
    {
        # Release the mutex when done
        $mutex.ReleaseMutex()
        $mutex.Dispose()
    }
}

# getuserprojects log file

<# # Define the log file path
$userProjectsPath = "C:\SRO\Apps\FileSync\logs\FileSynchProjects_MSMS_Prod-$UserID.txt"
$userProjectsDirectory = Split-Path $userProjectsPath

# Check if the log directory exists, if not, create it
if (-not (Test-Path -Path $userProjectsDirectory)) {
    New-Item -Path $userProjectsDirectory -ItemType Directory -Force | Out-Null
}

# Check if the log file exists, if not, create it
if (-not (Test-Path -Path $userProjectsPath)) {
    New-Item -Path $userProjectsPath -ItemType File -Force | Out-Null
} #>


#fileSync_Main -mode msms

# Specify the path to the text file containing folder names
<# $foldersFile = $userProjectsPath

# Attempt to read the folder names from the text file
try {
    $folderNames = Get-Content -Path $foldersFile
} catch {
    $mutex.ReleaseMutex()
    $mutex.Dispose()
    Write-Host "Error reading folders file: $($_.Exception.Message)"
    exit
} #>

# For every folder in the project list, add two more progress bars
<# foreach ($folderName in $folderNames) {
    # $global:totalOperations+=2
} #>
