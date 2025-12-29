# Auto-delete RAR parts after extraction
# Usage: .\extract-and-delete.ps1 "path\to\file.part1.rar"

param(
    [Parameter(Mandatory=$false)]
    [string]$FirstPartPath
)

# If no path provided, show file picker
if (-not $FirstPartPath) {
    Add-Type -AssemblyName System.Windows.Forms
    
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "RAR Part 1 Files (*.part1.rar;*.part01.rar)|*.part1.rar;*.part01.rar|All RAR Files (*.rar)|*.rar"
    $openFileDialog.Title = "Select the first part of the RAR archive (part1.rar)"
    $openFileDialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')
    
    $result = $openFileDialog.ShowDialog()
    
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $FirstPartPath = $openFileDialog.FileName
    }
    else {
        Write-Host "No file selected. Exiting..." -ForegroundColor Yellow
        Start-Sleep -Seconds 2
        exit 0
    }
}

# Ask for extraction destination
Add-Type -AssemblyName System.Windows.Forms

$folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
$folderBrowser.Description = "Select where to extract the files"
$folderBrowser.RootFolder = [System.Environment+SpecialFolder]::MyComputer
$folderBrowser.SelectedPath = (Get-Item $FirstPartPath).DirectoryName

$folderResult = $folderBrowser.ShowDialog()

if ($folderResult -eq [System.Windows.Forms.DialogResult]::OK) {
    $extractPath = $folderBrowser.SelectedPath
}
else {
    Write-Host "No destination selected. Exiting..." -ForegroundColor Yellow
    Start-Sleep -Seconds 2
    exit 0
}

# Get WinRAR path
$rarPath = "C:\Program Files\WinRAR\rar.exe"
if (-not (Test-Path $rarPath)) {
    Write-Host "WinRAR not found at $rarPath" -ForegroundColor Red
    exit 1
}

# Resolve full path
$FirstPartPath = Resolve-Path $FirstPartPath -ErrorAction Stop

# Validate first part exists
if (-not (Test-Path $FirstPartPath)) {
    Write-Host "File not found: $FirstPartPath" -ForegroundColor Red
    exit 1
}

# Get base name and directory
$fileInfo = Get-Item $FirstPartPath
$directory = $fileInfo.DirectoryName
$baseName = $fileInfo.Name -replace '\.part\d+\.rar$', ''

# Find all parts
$allParts = Get-ChildItem -Path $directory -Filter "$baseName.part*.rar" | Sort-Object Name
if ($allParts.Count -eq 0) {
    Write-Host "No multi-part RAR files found matching: $baseName.part*.rar" -ForegroundColor Red
    exit 1
}

Write-Host "Found $($allParts.Count) parts to extract" -ForegroundColor Cyan
foreach ($part in $allParts) {
    Write-Host "  $($part.Name)" -ForegroundColor Gray
}
Write-Host ""
Write-Host "Extracting to: $extractPath" -ForegroundColor Cyan
Write-Host "Starting extraction with auto-delete..." -ForegroundColor Green
Write-Host ""

# Start RAR extraction process
$psi = New-Object System.Diagnostics.ProcessStartInfo
$psi.FileName = $rarPath
$psi.Arguments = "x -ad `"$FirstPartPath`""
$psi.WorkingDirectory = $extractPath
$psi.UseShellExecute = $false
$psi.RedirectStandardOutput = $true
$psi.RedirectStandardError = $true
$psi.CreateNoWindow = $true

$process = New-Object System.Diagnostics.Process
$process.StartInfo = $psi

# Track current part
$currentPart = 1
$deletedParts = @()

# Start the process
[void]$process.Start()

try {
    # Read output line by line synchronously
    while (-not $process.StandardOutput.EndOfStream) {
        $line = $process.StandardOutput.ReadLine()
        
        if ($line) {
            Write-Host $line
            
            # Check if extracting from a new part
            if ($line -match 'Extracting from .*\.part(\d+)\.rar') {
                $newPart = [int]$matches[1]
                
                # If moved to a new part, delete the previous one
                if ($newPart -gt $currentPart) {
                    $partToDelete = $allParts | Where-Object { $_.Name -match "\.part$currentPart\.rar$" }
                    
                    if ($partToDelete -and (Test-Path $partToDelete.FullName)) {
                        Start-Sleep -Milliseconds 500
                        try {
                            Remove-Item $partToDelete.FullName -Force
                            Write-Host "  [DELETED] $($partToDelete.Name)" -ForegroundColor Yellow
                            $deletedParts += $partToDelete.Name
                        }
                        catch {
                            Write-Host "  [ERROR] Could not delete $($partToDelete.Name): $_" -ForegroundColor Red
                        }
                    }
                    
                    $currentPart = $newPart
                }
            }
        }
    }
    
    # Wait for process to complete
    $process.WaitForExit()
    
    # Read any remaining error output
    $errorOutput = $process.StandardError.ReadToEnd()
    if ($errorOutput) {
        Write-Host $errorOutput -ForegroundColor Red
    }
    
    # Delete the last part if extraction was successful
    if ($process.ExitCode -eq 0) {
        Start-Sleep -Seconds 1
        
        $lastPart = $allParts | Where-Object { $_.Name -match "\.part$currentPart\.rar$" }
        if ($lastPart -and (Test-Path $lastPart.FullName)) {
            try {
                Remove-Item $lastPart.FullName -Force
                Write-Host "  [DELETED] $($lastPart.Name)" -ForegroundColor Yellow
                $deletedParts += $lastPart.Name
            }
            catch {
                Write-Host "  [ERROR] Could not delete $($lastPart.Name): $_" -ForegroundColor Red
            }
        }
        
        Write-Host ""
        Write-Host "Extraction completed successfully!" -ForegroundColor Green
        Write-Host "Deleted $($deletedParts.Count) part(s)" -ForegroundColor Cyan
    }
    else {
        Write-Host ""
        Write-Host "Extraction failed with exit code: $($process.ExitCode)" -ForegroundColor Red
    }
    
    exit $process.ExitCode
}
finally {
    if ($process -and -not $process.HasExited) {
        $process.Kill()
    }
    $process.Dispose()
}
