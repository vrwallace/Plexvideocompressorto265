<#
Video Compression Script Overview
This PowerShell script is designed to automate the process of compressing video files using HandBrake. It recursively searches for video files in a specified directory, compresses them using predefined settings, and saves the compressed versions in an output directory.
Key Variables and Parameters

$HandbrakeCliPath: Path to the HandBrake CLI executable
Default: "C:\Program Files\HandBrake\HandBrakeCLI.exe"
$RootDirectory: The source directory where the script will search for video files to compress
Default: "l:"
$OutputDirectory: The directory where compressed videos will be saved
Default: "d:\CompressedVideos2"
$TempDirectory: A temporary directory used during the compression process
Default: "C:\Temp\VideoProcessing"
$LogFile: Path to the log file where the script will write its operation logs
Default: "d:\CompressedVideos\compression_log.txt"
$MaxRetries: Maximum number of retry attempts for failed operations
Default: 3
$RetryDelay: Delay in seconds between retry attempts
Default: 5
$MaxParallelJobs: Maximum number of parallel compression jobs (currently set to 1)
Default: 1
$SkipExisting: A switch parameter to skip files that already exist in the output directory
Default: Not set (false)
$MinimumSizeBytes: Minimum file size in bytes for videos to be considered for compression
Default: 0 (no minimum size)
$EmailFrom, $EmailTo, $SmtpServer: Email notification settings (optional)
Default: Empty strings

Main Functions

Clear-TempDirectory: Cleans up the temporary directory before starting the compression process.
Test-NetworkShare: Checks if the network share (root directory) is accessible.
Process-VideoFile: Handles the compression of individual video files.
Start-VideoCompression: The main function that orchestrates the entire compression process.

Process Flow

The script starts by cleaning the temporary directory.
It checks if the root directory (network share) is accessible.
It creates the output and temporary directories if they don't exist.
The script then searches for video files in the root directory.
Each video file is processed:

Copied to the temp directory
Compressed using HandBrake
Moved to the output directory


The script logs its progress and any errors encountered.
Upon completion, it generates a summary report and can send an email notification.

This script is designed to be configurable and robust, with error handling and logging to ensure reliable operation when processing large numbers of video files.


#>



# Script parameters

#root is source directory it will recursely grab files
#put temp on high speed storage nvme ssd etc

param (
    [string]$HandbrakeCliPath = "C:\Program Files\HandBrake\HandBrakeCLI.exe",
    [string]$RootDirectory = "l:\james",
    [string]$OutputDirectory = "l:\CompressedVideos2",
    [string]$TempDirectory = "C:\Temp\VideoProcessing",
    [string]$LogFile = "l:\CompressedVideos2\compression_log.txt",
    [int]$MaxRetries = 3,
    [int]$RetryDelay = 5,
    [switch]$SkipExisting,
    [long]$MinimumSizeBytes = 0,
    [string]$EmailFrom = "",
    [string]$EmailTo = "",
    [string]$SmtpServer = ""
)

# Increase the maximum function count and memory limit
$MaximumFunctionCount = 32768
[System.GC]::Collect()
$MaxMemoryPerShellMB = 2048

# Enable strict mode for better error handling
Set-StrictMode -Version Latest

# Error action preference
$ErrorActionPreference = "Stop"

# Import required modules
Import-Module Microsoft.PowerShell.Utility

# Transcript logging
$transcriptPath = Join-Path $OutputDirectory "compression_transcript.txt"
Start-Transcript -Path $transcriptPath -Append

# HandBrake encoding settings
$handbrakeSettings = @{
    Encoder = "nvenc_h265"
    Quality = 20
    PeakFrameRate = $true
    MaxWidth = 0  # 0 means no scaling
    MaxHeight = 0  # 0 means no scaling
    AudioEncoders = @("copy")  # Use passthrough to keep original audio
    SubtitleLangList = "eng"
    SubtitleBurned = $false
    SubtitleDefault = $false
    Crop = "none"
    Decomb = $false
    Detelecine = $false
    Deinterlace = $false
    Denoise = "off"
    ChromaSmooth = $false
    ChapterMarkers = $true
    Format = "av_mkv"  # Changed to MKV container
    Align_AV = $true
}


# Function to get user confirmation
function Get-UserConfirmation {
    param (
        [string]$Prompt
    )
    $confirmation = Read-Host -Prompt $Prompt
    return $confirmation -eq 'y'
}

function Clear-TempDirectory {
    param (
        [string]$TempDirectory,
        [string]$LogFile
    )

    Write-Log "Checking temp directory: $TempDirectory" -LogFile $LogFile

    if (Test-Path $TempDirectory) {
        $files = @(Get-ChildItem -Path $TempDirectory -File -Depth 0 -ErrorAction SilentlyContinue)
        $fileCount = $files.Count

        if ($fileCount -gt 0) {
            Write-Host "The temp directory contains $fileCount file(s) at the root level."
            
            # First confirmation
            $confirm1 = Get-UserConfirmation "Do you want to clear these files from the temp directory? (y/n)"
            
            if ($confirm1) {
                # Second confirmation
                $confirm2 = Get-UserConfirmation "Are you sure you want to clear these files? This action cannot be undone. (y/n)"
                
                if ($confirm2) {
                    Write-Log "User confirmed. Cleaning up files in temp directory." -LogFile $LogFile
                    try {
                        foreach ($file in $files) {
                            try {
                                Remove-Item $file.FullName -Force
                                Write-Log "Removed temp file: $($file.FullName)" -Level "DEBUG" -LogFile $LogFile
                            } catch {
                                Write-Log "Failed to remove temp file: $($file.FullName). Error: $_" -Level "WARN" -LogFile $LogFile
                            }
                        }
                        Write-Log "Temp directory file cleanup completed." -LogFile $LogFile
                    } catch {
                        Write-Log "Error during temp directory file cleanup: $_" -Level "ERROR" -LogFile $LogFile
                    }
                } else {
                    Write-Log "User declined second confirmation. Temp directory not cleared." -LogFile $LogFile
                }
            } else {
                Write-Log "User declined first confirmation. Temp directory not cleared." -LogFile $LogFile
            }
        } else {
            Write-Log "No files found in the root of the temp directory. No cleanup needed." -LogFile $LogFile
        }
    } else {
        Write-Log "Temp directory does not exist. No cleanup needed." -LogFile $LogFile
    }
}
# Function to write log entries
function Write-Log {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Message,
        [Parameter(Mandatory=$false)]
        [ValidateSet("INFO", "WARN", "ERROR", "DEBUG")]
        [string]$Level = "INFO",
        [Parameter(Mandatory=$false)]
        [string]$LogFile
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "$timestamp [$Level] $Message"
    
    if ($LogFile) {
        $mutex = New-Object System.Threading.Mutex($false, "Global\LogFileMutex")
        $mutex.WaitOne() | Out-Null
        try {
            Add-Content -Path $LogFile -Value $logEntry
        }
        finally {
            $mutex.ReleaseMutex()
        }
    }

    switch ($Level) {
        "INFO"  { Write-Host $logEntry -ForegroundColor White }
        "WARN"  { Write-Host $logEntry -ForegroundColor Yellow }
        "ERROR" { Write-Host $logEntry -ForegroundColor Red }
        "DEBUG" { Write-Host $logEntry -ForegroundColor Gray }
    }
}

# Function to check network share accessibility
function Test-NetworkShare {
    param ([string]$Path)
    try {
        if (Test-Path $Path) {
            Write-Log "Network share $Path is accessible." -LogFile $LogFile
            return $true
        } else {
            Write-Log "Network share $Path is not accessible." -Level "ERROR" -LogFile $LogFile
            return $false
        }
    } catch {
        Write-Log "Error accessing network share ${Path}: $_" -Level "ERROR" -LogFile $LogFile
        return $false
    }
}

# Function to send email
function Send-EmailNotification {
    param (
        [string]$Subject,
        [string]$Body
    )
    if ($EmailFrom -and $EmailTo -and $SmtpServer) {
        try {
            Send-MailMessage -From $EmailFrom -To $EmailTo -Subject $Subject -Body $Body -SmtpServer $SmtpServer
            Write-Log "Email notification sent successfully." -LogFile $LogFile
        } catch {
            Write-Log "Failed to send email notification: $_" -Level "ERROR" -LogFile $LogFile
        }
    }
}

# Function to clean up in case of errors
function Invoke-Cleanup {
    param (
        [string]$TempPath,
        [string]$OutputPath
    )
    if (Test-Path $TempPath) {
        Remove-Item -Path $TempPath -Force
    }
    if (Test-Path $OutputPath) {
        Remove-Item -Path $OutputPath -Force
    }
}

# Function to copy file with retry logic
function Copy-FileWithRetry {
    param (
        [string]$SourcePath,
        [string]$DestinationPath,
        [string]$LogFile,
        [int]$MaxRetries = 5,
        [int]$RetryDelay = 10
    )

    $retry = 0
    $success = $false

    while (-not $success -and $retry -lt $MaxRetries) {
        try {
            Write-Log "Copying file to temp location: $DestinationPath" -LogFile $LogFile
            Copy-Item -Path $SourcePath -Destination $DestinationPath -Force -ErrorAction Stop

            # Verify the copy is complete
            if (!(Test-Path $DestinationPath)) {
                throw "Destination file not found after copy"
            }

            $sourceFileInfo = Get-Item $SourcePath
            $destFileInfo = Get-Item $DestinationPath

            if ($destFileInfo.Length -ne $sourceFileInfo.Length) {
                throw "File sizes don't match after copy"
            }

            $success = $true
            Write-Log "File successfully copied to temp location: $DestinationPath" -LogFile $LogFile
        }
        catch {
            $retry++
            Write-Log "File copy to temp failed. Attempt $retry of $MaxRetries. Error: $_" -Level "WARN" -LogFile $LogFile
            if ($retry -lt $MaxRetries) {
                Write-Log "Retrying in $RetryDelay seconds..." -LogFile $LogFile
                Start-Sleep -Seconds $RetryDelay
            }
            else {
                Write-Log "File copy failed after $MaxRetries attempts." -Level "ERROR" -LogFile $LogFile
                throw
            }
        }
    }
}

function Build-HandbrakeArguments {
    param (
        [string]$InputPath,
        [string]$OutputPath,
        [hashtable]$Settings
    )

    $args = @("-i", $InputPath, "-o", $OutputPath)

    foreach ($key in $Settings.Keys) {
        $value = $Settings[$key]
        switch ($key) {
            "Encoder" { $args += "-e", $value }
            "Quality" { $args += "-q", $value }
            "PeakFrameRate" { $args += $value ? "--pfr" : "--cfr" }
            "MaxWidth" { if ($value -ne 0) { $args += "-X", $value } }
            "MaxHeight" { if ($value -ne 0) { $args += "-Y", $value } }
            "AudioEncoders" { $args += "-E", ($value -join ",") }
            "SubtitleLangList" { 
                $args += "--subtitle-lang-list", $value 
                $args += "--all-subtitles"
            }
            "SubtitleBurned" { $args += "--subtitle-burned=none" }
            "SubtitleDefault" { $args += "--subtitle-default=none" }
            "Crop" { $args += "--crop", $value }
            "Decomb" { $args += $value ? "--decomb" : "--no-decomb" }
            "Detelecine" { $args += $value ? "--detelecine" : "--no-detelecine" }
            "Deinterlace" { $args += $value ? "--deinterlace" : "--no-deinterlace" }
            "Denoise" { if ($value -ne "off") { $args += "--denoise", $value } }
            "Sharpen" { if ($value -ne "off") { $args += "--sharpen", $value } }
            "ChromaSmooth" { $args += $value ? "--comb-detect" : "--no-comb-detect" }
            "ChapterMarkers" { $args += $value ? "--markers" : "--no-markers" }
            "Format" { $args += "-f", $value }
            "Align_AV" { $args += $value ? "--align-av" : "--no-align-av" }
            default { Write-Log "Unknown HandBrake setting: $key" -Level "WARN" -LogFile $LogFile }
        }
    }

    # Add subtitle tracks
    $args += "--subtitle-lang-list", "all"
    $args += "--all-subtitles"

    return $args
}

# Function to run HandBrake with retry logic
function Run-HandbrakeWithRetry {
    param (
        [string]$HandbrakeCliPath,
        [array]$Arguments,
        [string]$OutputPath,
        [string]$LogFile,
        [int]$MaxRetries = 3,
        [int]$RetryDelay = 5
    )

    $retry = 0
    $success = $false

    while (-not $success -and $retry -lt $MaxRetries) {
        try {
            # Check if HandBrake CLI exists
            if (-not (Test-Path $HandbrakeCliPath)) {
                throw "HandBrake CLI not found at path: $HandbrakeCliPath"
            }

            # Log the full HandBrake command
            $fullCommand = "& `"$HandbrakeCliPath`" $($Arguments -join ' ')"
            Write-Log "Executing HandBrake command: $fullCommand" -LogFile $LogFile

            # Run HandBrake and capture output
            $handbrakeOutput = & $HandbrakeCliPath $Arguments 2>&1
            
            # Log the HandBrake output
            $handbrakeOutput | ForEach-Object { Write-Log "HandBrake output: $_" -LogFile $LogFile }

            if ($LASTEXITCODE -ne 0) {
                throw "HandBrake encoding failed with exit code $LASTEXITCODE. See log for full output."
            }

            # Verify HandBrake output
            if (!(Test-Path $OutputPath)) {
                throw "HandBrake output file not found at: $OutputPath"
            }

            $outputFileInfo = Get-Item $OutputPath
            if ($outputFileInfo.Length -eq 0) {
                throw "HandBrake output file is empty: $OutputPath"
            }

            $success = $true
            Write-Log "HandBrake encoding completed successfully: $OutputPath" -LogFile $LogFile
        }
        catch {
            $retry++
            Write-Log "HandBrake encoding failed. Attempt $retry of $MaxRetries. Error: $_" -Level "WARN" -LogFile $LogFile
            if ($retry -lt $MaxRetries) {
                Write-Log "Retrying in $RetryDelay seconds..." -LogFile $LogFile
                Start-Sleep -Seconds $RetryDelay
            }
            else {
                Write-Log "HandBrake encoding failed after $MaxRetries attempts." -Level "ERROR" -LogFile $LogFile
                throw
            }
        }
    }
}

# Function to process a single video file
function Process-VideoFile {
    param (
        [System.IO.FileInfo]$File,
        [string]$RootDirectory,
        [string]$OutputDirectory,
        [string]$TempDirectory,
        [string]$HandbrakeCliPath,
        [hashtable]$HandbrakeSettings,
        [string]$LogFile
    )

    $relativePath = $File.FullName.Substring($RootDirectory.Length)
    $fileName = [System.IO.Path]::GetFileNameWithoutExtension($File.Name)
    $outputFileName = "${fileName}_optimized.mkv"
    $outputPath = Join-Path -Path $OutputDirectory -ChildPath $relativePath
    $outputPath = Join-Path -Path (Split-Path $outputPath -Parent) -ChildPath $outputFileName
    $outputFolder = Split-Path -Path $outputPath -Parent
    $tempPath = Join-Path -Path $TempDirectory -ChildPath ($File.Name + ".temp")
    $tempOutputPath = Join-Path -Path $TempDirectory -ChildPath $outputFileName

    # Check if output file already exists
    if (Test-Path $outputPath) {
        Write-Log "Skipping $($File.Name) - optimized file already exists at $outputPath" -LogFile $LogFile
        return @{
            FileName = $File.Name
            OriginalSize = $File.Length
            OptimizedSize = 0
            CompressionRatio = 0
            ProcessingTime = 0
            Status = "Skipped"
        }
    }

    # Cleanup any existing partial files
    if (Test-Path $tempOutputPath) {
        Remove-Item -Path $tempOutputPath -Force
    }

    if (!(Test-Path -Path $outputFolder)) {
        New-Item -ItemType Directory -Path $outputFolder -Force | Out-Null
        Write-Log "Created output folder: $outputFolder" -Level "DEBUG" -LogFile $LogFile
    }

    Write-Log "Processing file: $($File.Name)" -LogFile $LogFile

    $result = @{
        FileName = $File.Name
        OriginalSize = $File.Length
        OptimizedSize = 0
        CompressionRatio = 0
        ProcessingTime = 0
        Status = "Failed"
    }

    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    try {
        # Verify source file
        if (!(Test-Path $File.FullName)) {
            throw "Source file not found: $($File.FullName)"
        }

        # Copy to temp and ensure it's complete
        Copy-FileWithRetry -SourcePath $File.FullName -DestinationPath $tempPath -LogFile $LogFile

        # Verify temp file
        if (!(Test-Path $tempPath)) {
            throw "Temp file not found after copy: $tempPath"
        }

        # Run HandBrake with temp output path
        $handbrakeArgs = Build-HandbrakeArguments -InputPath $tempPath -OutputPath $tempOutputPath -Settings $HandbrakeSettings
        Run-HandbrakeWithRetry -HandbrakeCliPath $HandbrakeCliPath -Arguments $handbrakeArgs -OutputPath $tempOutputPath -LogFile $LogFile
        
        # Verify output file
        if (!(Test-Path $tempOutputPath)) {
            throw "Output file not found after HandBrake encoding: $tempOutputPath"
        }

        $outputFileInfo = Get-Item $tempOutputPath
        if ($outputFileInfo.Length -lt 1MB) {  # Adjust this threshold as needed
            throw "HandBrake output file is suspiciously small: $($outputFileInfo.Length) bytes"
        }

        # Move the encoded file from temp to final destination
        Move-Item -Path $tempOutputPath -Destination $outputPath -Force

        # Remove temp file
        Remove-Item -Path $tempPath -Force -ErrorAction Stop
        Write-Log "Temp file removed successfully: $tempPath" -LogFile $LogFile

        $optimizedFile = Get-Item $outputPath
        $result.OptimizedSize = $optimizedFile.Length
        $result.CompressionRatio = [math]::Round(($File.Length - $optimizedFile.Length) / $File.Length * 100, 2)
        $result.Status = "Success"

        Write-Log "Optimized $($File.Name): Original size: $($result.OriginalSize) bytes, Optimized size: $($result.OptimizedSize) bytes, Reduction: $($result.CompressionRatio)%" -LogFile $LogFile
    }
    catch {
        Write-Log "Error processing $($File.Name): $_" -Level "ERROR" -LogFile $LogFile
        Invoke-Cleanup -TempPath $tempPath -OutputPath $tempOutputPath
        Invoke-Cleanup -TempPath $tempPath -OutputPath $outputPath
    }
    finally {
        $stopwatch.Stop()
        $result.ProcessingTime = $stopwatch.Elapsed.TotalSeconds
    }

    return $result
}

# Main compression function
function Start-VideoCompression {

    Write-Log "Starting video compression script" -LogFile $LogFile

    # Clean up temp directory before starting, with user confirmation
    Clear-TempDirectory -TempDirectory $TempDirectory -LogFile $LogFile

    if (-not (Test-NetworkShare $RootDirectory)) {
        Write-Log "Cannot access the network share. Exiting script." -Level "ERROR" -LogFile $LogFile
        return
    }

    

    

    @($OutputDirectory, $TempDirectory) | ForEach-Object {
        if (!(Test-Path -Path $_)) {
            New-Item -ItemType Directory -Path $_ -Force | Out-Null
            Write-Log "Created directory: $_" -LogFile $LogFile
        }
    }

    Write-Log "Searching for video files in $RootDirectory" -LogFile $LogFile
    $videoFiles = Get-ChildItem -Path $RootDirectory -Recurse -Include @("*.mp4", "*.avi", "*.mkv", "*.mov") |
        Where-Object { $_.Length -ge $MinimumSizeBytes }
    Write-Log "Found $($videoFiles.Count) video files to process" -LogFile $LogFile

    $totalFiles = $videoFiles.Count
    $processedFiles = 0
    $summary = @()

    foreach ($file in $videoFiles) {
        $result = Process-VideoFile -File $file -RootDirectory $RootDirectory -OutputDirectory $OutputDirectory `
                                    -TempDirectory $TempDirectory -HandbrakeCliPath $HandbrakeCliPath `
                                    -HandbrakeSettings $handbrakeSettings -LogFile $LogFile

        $summary += $result
        $processedFiles++
        Write-Progress -Activity "Optimizing Videos" -Status "Processing $processedFiles of $totalFiles" -PercentComplete (($processedFiles / $totalFiles) * 100)
    }

    # Generate summary report
    $reportPath = Join-Path $OutputDirectory "OptimizationSummary.csv"
    $summary | Export-Csv -Path $reportPath -NoTypeInformation

    $totalSavings = ($summary | Measure-Object -Property OriginalSize -Sum).Sum - ($summary | Measure-Object -Property OptimizedSize -Sum).Sum
    $averageCompressionRatio = ($summary | Measure-Object -Property CompressionRatio -Average).Average

    $completionMessage = @"
Video optimization completed.
Total files processed: $totalFiles
Total space saved: $([math]::Round($totalSavings / 1GB, 2)) GB
Average compression ratio: $([math]::Round($averageCompressionRatio, 2))%
Summary report saved to: $reportPath
"@

    Write-Log $completionMessage -LogFile $LogFile
    Send-EmailNotification -Subject "Video Optimization Completed" -Body $completionMessage
}

# Call the main function to start the script
Start-VideoCompression

# Stop the transcript
Stop-Transcript