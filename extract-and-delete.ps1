<#
.SYNOPSIS
    Extract multi-part RAR archives with automatic part deletion
.DESCRIPTION
    This script extracts multi-part RAR archives and automatically deletes each part
    after it's no longer needed, saving disk space during extraction.
.PARAMETER FirstPartPath
    Path to the first part of the RAR archive (e.g., file.part1.rar)
.EXAMPLE
    .\extract-and-delete.ps1
    .\extract-and-delete.ps1 "C:\Downloads\archive.part1.rar"
#>

param(
    [Parameter(Mandatory=$false)]
    [string]$FirstPartPath
)

#region Configuration
$CONFIG = @{
    RarPath = "C:\Program Files\WinRAR\rar.exe"
    DeleteDelayMs = 500  # Delay before deleting a part to ensure it's not in use
    ErrorDisplaySeconds = 5  # How long to show error messages before exiting
    PartFilePattern = "*.part*.rar"  # Pattern to match RAR parts
    TestBeforeExtract = $true  # Whether to test archive integrity before extraction
}
#endregion

#region Helper Functions

function Write-ColorMessage {
    <#
    .SYNOPSIS
        Write colored console messages with consistent formatting
    #>
    param(
        [string]$Message,
        [ValidateSet('Info', 'Success', 'Warning', 'Error', 'Gray')]
        [string]$Type = 'Info'
    )
    
    $colors = @{
        Info = 'Cyan'
        Success = 'Green'
        Warning = 'Yellow'
        Error = 'Red'
        Gray = 'Gray'
    }
    
    Write-Host $Message -ForegroundColor $colors[$Type]
}

function Get-RarExecutable {
    <#
    .SYNOPSIS
        Validate that WinRAR is installed and accessible
    #>
    if (-not (Test-Path $CONFIG.RarPath)) {
        Write-ColorMessage "WinRAR not found at: $($CONFIG.RarPath)" -Type Error
        Write-ColorMessage "Please install WinRAR or update the path in the script configuration." -Type Error
        return $null
    }
    return $CONFIG.RarPath
}

function Get-SourceArchive {
    <#
    .SYNOPSIS
        Get the source archive path via file picker or parameter
    #>
    param([string]$ProvidedPath)
    
    if ($ProvidedPath) {
        $resolvedPath = Resolve-Path $ProvidedPath -ErrorAction SilentlyContinue
        if ($resolvedPath) {
            return $resolvedPath.Path
        }
        Write-ColorMessage "Provided file not found: $ProvidedPath" -Type Error
        return $null
    }
    
    # Show file picker
    Add-Type -AssemblyName System.Windows.Forms
    
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "RAR Part 1 Files (*.part1.rar;*.part01.rar)|*.part1.rar;*.part01.rar|All RAR Files (*.rar)|*.rar"
    $openFileDialog.Title = "Select the first part of the RAR archive (part1.rar)"
    $openFileDialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')
    
    $result = $openFileDialog.ShowDialog()
    
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $openFileDialog.FileName
    }
    
    Write-ColorMessage "No file selected." -Type Warning
    return $null
}

function Get-ExtractionDestination {
    <#
    .SYNOPSIS
        Get the extraction destination via folder picker
    #>
    param([string]$DefaultPath)
    
    Add-Type -AssemblyName System.Windows.Forms
    
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "Select where to extract the files"
    $folderBrowser.RootFolder = [System.Environment+SpecialFolder]::MyComputer
    $folderBrowser.SelectedPath = $DefaultPath
    
    $result = $folderBrowser.ShowDialog()
    
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $folderBrowser.SelectedPath
    }
    
    Write-ColorMessage "No destination selected." -Type Warning
    return $null
}

function Get-ArchiveParts {
    <#
    .SYNOPSIS
        Find all parts of a multi-part RAR archive
    #>
    param(
        [string]$FirstPartPath
    )
    
    $fileInfo = Get-Item $FirstPartPath
    $directory = $fileInfo.DirectoryName
    $baseName = $fileInfo.Name -replace '\.part\d+\.rar$', ''
    
    $allParts = Get-ChildItem -Path $directory -Filter "$baseName.part*.rar" | Sort-Object Name
    
    if ($allParts.Count -eq 0) {
        Write-ColorMessage "No multi-part RAR files found matching: $baseName.part*.rar" -Type Error
        return $null
    }
    
    return @{
        Parts = $allParts
        Directory = $directory
        BaseName = $baseName
    }
}

function Test-ArchiveIntegrity {
    <#
    .SYNOPSIS
        Test the integrity of the RAR archive before extraction
    #>
    param(
        [string]$RarPath,
        [string]$FirstPartPath,
        [string]$WorkingDirectory
    )
    
    Write-ColorMessage "`nTesting archive integrity..." -Type Info
    
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $RarPath
    $psi.Arguments = "t `"$FirstPartPath`""
    $psi.WorkingDirectory = $WorkingDirectory
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.CreateNoWindow = $true
    
    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $psi
    
    try {
        [void]$process.Start()
        
        # Read and display test output
        while (-not $process.StandardOutput.EndOfStream) {
            $line = $process.StandardOutput.ReadLine()
            if ($line) {
                Write-ColorMessage "  $line" -Type Gray
            }
        }
        
        $process.WaitForExit()
        $errorOutput = $process.StandardError.ReadToEnd()
        
        if ($process.ExitCode -ne 0) {
            Write-ColorMessage "`nArchive test FAILED! The archive may be corrupted." -Type Error
            if ($errorOutput) {
                Write-ColorMessage $errorOutput -Type Error
            }
            return $false
        }
        
        Write-ColorMessage "`nArchive test passed successfully!" -Type Success
        return $true
    }
    finally {
        $process.Dispose()
    }
}

function Start-ExtractionWithCleanup {
    <#
    .SYNOPSIS
        Extract the archive and delete parts progressively
    #>
    param(
        [string]$RarPath,
        [string]$FirstPartPath,
        [string]$DestinationPath,
        [array]$AllParts
    )
    
    Write-ColorMessage "`nExtracting to: $DestinationPath" -Type Info
    Write-ColorMessage "Starting extraction with auto-delete...`n" -Type Success
    
    # Setup extraction process
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $RarPath
    $psi.Arguments = "x `"$FirstPartPath`""
    $psi.WorkingDirectory = $DestinationPath
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.CreateNoWindow = $true
    
    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $psi
    
    $currentPart = 1
    $deletedParts = @()
    
    try {
        [void]$process.Start()
        
        # Monitor extraction progress and delete parts
        while (-not $process.StandardOutput.EndOfStream) {
            $line = $process.StandardOutput.ReadLine()
            
            if ($line) {
                Write-Host $line
                
                # Detect when RAR moves to a new part
                if ($line -match 'Extracting from .*\.part(\d+)\.rar') {
                    $newPart = [int]$matches[1]
                    
                    if ($newPart -gt $currentPart) {
                        $partToDelete = $AllParts | Where-Object { $_.Name -match "\.part$currentPart\.rar$" }
                        
                        if ($partToDelete -and (Test-Path $partToDelete.FullName)) {
                            Start-Sleep -Milliseconds $CONFIG.DeleteDelayMs
                            
                            try {
                                Remove-Item $partToDelete.FullName -Force -ErrorAction Stop
                                Write-ColorMessage "  [DELETED] $($partToDelete.Name)" -Type Warning
                                $deletedParts += $partToDelete.Name
                            }
                            catch {
                                Write-ColorMessage "  [ERROR] Could not delete $($partToDelete.Name): $_" -Type Error
                            }
                        }
                        
                        $currentPart = $newPart
                    }
                }
            }
        }
        
        $process.WaitForExit()
        
        # Read any error output
        $errorOutput = $process.StandardError.ReadToEnd()
        if ($errorOutput) {
            Write-ColorMessage $errorOutput -Type Error
        }
        
        # Delete the last part if extraction was successful
        if ($process.ExitCode -eq 0) {
            Start-Sleep -Seconds 1
            
            $lastPart = $AllParts | Where-Object { $_.Name -match "\.part$currentPart\.rar$" }
            if ($lastPart -and (Test-Path $lastPart.FullName)) {
                try {
                    Remove-Item $lastPart.FullName -Force -ErrorAction Stop
                    Write-ColorMessage "  [DELETED] $($lastPart.Name)" -Type Warning
                    $deletedParts += $lastPart.Name
                }
                catch {
                    Write-ColorMessage "  [ERROR] Could not delete $($lastPart.Name): $_" -Type Error
                }
            }
            
            Write-ColorMessage "`nExtraction completed successfully!" -Type Success
            Write-ColorMessage "Deleted $($deletedParts.Count) part(s)" -Type Info
            return $true
        }
        else {
            Write-ColorMessage "`nExtraction failed with exit code: $($process.ExitCode)" -Type Error
            return $false
        }
    }
    finally {
        if ($process -and -not $process.HasExited) {
            $process.Kill()
        }
        $process.Dispose()
    }
}

#endregion

#region Main Execution

try {
    # Step 1: Validate WinRAR installation
    $rarPath = Get-RarExecutable
    if (-not $rarPath) {
        Start-Sleep -Seconds $CONFIG.ErrorDisplaySeconds
        exit 1
    }
    
    # Step 2: Get source archive
    $sourcePath = Get-SourceArchive -ProvidedPath $FirstPartPath
    if (-not $sourcePath) {
        Start-Sleep -Seconds 2
        exit 0
    }
    
    # Step 3: Get archive parts
    $archiveInfo = Get-ArchiveParts -FirstPartPath $sourcePath
    if (-not $archiveInfo) {
        Start-Sleep -Seconds $CONFIG.ErrorDisplaySeconds
        exit 1
    }
    
    Write-ColorMessage "Found $($archiveInfo.Parts.Count) parts to extract" -Type Info
    foreach ($part in $archiveInfo.Parts) {
        Write-ColorMessage "  $($part.Name)" -Type Gray
    }
    
    # Step 4: Get extraction destination
    $destinationPath = Get-ExtractionDestination -DefaultPath $archiveInfo.Directory
    if (-not $destinationPath) {
        Start-Sleep -Seconds 2
        exit 0
    }
    
    # Step 5: Test archive integrity
    if ($CONFIG.TestBeforeExtract) {
        $testPassed = Test-ArchiveIntegrity -RarPath $rarPath -FirstPartPath $sourcePath -WorkingDirectory $archiveInfo.Directory
        if (-not $testPassed) {
            Write-ColorMessage "Aborting extraction to prevent data loss." -Type Error
            Start-Sleep -Seconds $CONFIG.ErrorDisplaySeconds
            exit 1
        }
    }
    
    # Step 6: Extract with progressive cleanup
    $success = Start-ExtractionWithCleanup -RarPath $rarPath -FirstPartPath $sourcePath -DestinationPath $destinationPath -AllParts $archiveInfo.Parts
    
    if ($success) {
        exit 0
    }
    else {
        exit 1
    }
}
catch {
    Write-ColorMessage "`nUnexpected error occurred:" -Type Error
    Write-ColorMessage $_.Exception.Message -Type Error
    Write-ColorMessage $_.ScriptStackTrace -Type Gray
    Start-Sleep -Seconds $CONFIG.ErrorDisplaySeconds
    exit 1
}

#endregion
