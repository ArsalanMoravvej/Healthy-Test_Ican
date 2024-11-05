# Function to get the Windows version
function Get-WindowsVersion {
    $os = Get-CimInstance -ClassName Win32_OperatingSystem
    return "$($os.Caption) - Version $($os.Version)"
}

# Function to get the free and total space of a specified drive
function Get-DriveSpace {
    param (
        [string]$driveName
    )
    
    $drive = Get-PSDrive -Name $driveName
    if ($null -eq $drive) {
        throw "Drive $driveName not found."
    }

    $freeSpaceGB = [math]::Round($drive.Free / 1GB, 0)
    $totalSpaceGB = [math]::Round(($drive.Used + $drive.Free) / 1GB, 0)
    
    return [pscustomobject]@{
        FreeSpaceGB  = $freeSpaceGB
        TotalSpaceGB = $totalSpaceGB
    }
}

# Function to get the total memory size in MB
function Get-TotalMemoryMB {
    $totalMemoryMB = (Get-CimInstance -ClassName Win32_PhysicalMemory | Measure-Object -Property Capacity -Sum).Sum / 1MB
    return [math]::Round($totalMemoryMB, 0)
}

# Function to get the average RAM usage in percentage
function Get-AverageRamUsagePercent {
    $ramUsageSamples = Get-Counter '\Memory\Committed Bytes' -SampleInterval 1 -MaxSamples 5
    $averageRamUsageMB = ($ramUsageSamples.CounterSamples | Measure-Object -Property CookedValue -Average).Average / 1MB
    $totalMemoryMB = Get-TotalMemoryMB
    $averageRamUsagePercent = ($averageRamUsageMB / $totalMemoryMB) * 100
    return [math]::Round($averageRamUsagePercent, 1)
}

# Function to get the average CPU usage percentage
function Get-AverageCpuUsagePercent {
    $cpuUsageSamples = Get-Counter '\Processor(_Total)\% Processor Time' -SampleInterval 1 -MaxSamples 5
    $averageCpuUsage = ($cpuUsageSamples.CounterSamples | Measure-Object -Property CookedValue -Average).Average
    return [math]::Round($averageCpuUsage, 1)
}

# Function to update a bookmark in a Word document without reopening Word each time
function Update-WordBookmark {
    param (
        $wordApp,       # Word Application object
        $document,      # Word Document object
        [string]$bookmarkName,  # Name of the bookmark
        [string]$textToInsert   # Text to insert at the bookmark
    )

    if ($document.Bookmarks.Exists($bookmarkName)) {
        $bookmark = $document.Bookmarks.Item($bookmarkName)
        $range = $bookmark.Range
        $range.Text = $textToInsert
        $document.Bookmarks.Add($bookmarkName, $range)
        
        Write-Host "Successfully updated the bookmark '$bookmarkName' with text: $textToInsert"
    }
    else {
        Write-Host "Bookmark '$bookmarkName' does not exist in the document."
    }
}

# Main script
$docPath = Join-Path -Path $PSScriptRoot -ChildPath "ApplicationServerHealthyCheckList.docx"

# Define bookmarks for Drive C, Drive E, Windows version, RAM, and CPU usage
$bookmarkFreeSpaceC = "FreeSpaceCBookmark"
$bookmarkTotalSpaceC = "TotalSpaceCBookmark"
$bookmarkFreeSpaceE = "FreeSpaceEBookmark"
$bookmarkTotalSpaceE = "TotalSpaceEBookmark"
$bookmarkWindowsVersion = "WindowsVersionBookmark"
$bookmarkTotalMemory = "TotalMemoryBookmark"
$bookmarkAverageRamUsage = "AverageRamUsageBookmark"
$bookmarkAverageCpuUsage = "AverageCpuUsageBookmark"

try {
    # Create Word Application COM object and open the document once
    $wordApp = New-Object -ComObject Word.Application
    if ($null -eq $wordApp) {
        throw "Failed to initialize Word COM object."
    }
    
    $wordApp.Visible = $false
    $document = $wordApp.Documents.Open($docPath)
    if ($null -eq $document) {
        throw "Failed to open the document at path: $docPath"
    }

    # Get system information
    $driveSpaceC = Get-DriveSpace -driveName "C"
    $driveSpaceE = Get-DriveSpace -driveName "E"
    $windowsVersion = Get-WindowsVersion
    $totalMemory = Get-TotalMemoryMB
    $averageRamUsage = Get-AverageRamUsagePercent
    $averageCpuUsage = Get-AverageCpuUsagePercent

    # Update bookmarks for Drive C
    Update-WordBookmark -wordApp $wordApp -document $document -bookmarkName $bookmarkFreeSpaceC -textToInsert "$($driveSpaceC.FreeSpaceGB) GB"
    Update-WordBookmark -wordApp $wordApp -document $document -bookmarkName $bookmarkTotalSpaceC -textToInsert "$($driveSpaceC.TotalSpaceGB) GB"

    # Update bookmarks for Drive E
    Update-WordBookmark -wordApp $wordApp -document $document -bookmarkName $bookmarkFreeSpaceE -textToInsert "$($driveSpaceE.FreeSpaceGB) GB"
    Update-WordBookmark -wordApp $wordApp -document $document -bookmarkName $bookmarkTotalSpaceE -textToInsert "$($driveSpaceE.TotalSpaceGB) GB"

    # Update the bookmark for Windows version
    Update-WordBookmark -wordApp $wordApp -document $document -bookmarkName $bookmarkWindowsVersion -textToInsert $windowsVersion

    # Update bookmarks for RAM and CPU information
    Update-WordBookmark -wordApp $wordApp -document $document -bookmarkName $bookmarkTotalMemory -textToInsert "$totalMemory MB"
    Update-WordBookmark -wordApp $wordApp -document $document -bookmarkName $bookmarkAverageRamUsage -textToInsert "$averageRamUsage %"
    Update-WordBookmark -wordApp $wordApp -document $document -bookmarkName $bookmarkAverageCpuUsage -textToInsert "$averageCpuUsage %"

    # Save and close the document
    $document.Save()
    $document.Close()
}
catch {
    Write-Host "An error occurred: $_"
}
finally {
    # Quit Word Application
    if ($wordApp) {
        $wordApp.Quit()
    }
}
