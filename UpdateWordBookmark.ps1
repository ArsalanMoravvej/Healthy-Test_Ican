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

    $freeSpaceGB = [math]::Round($drive.Free / 1GB, 2)
    $totalSpaceGB = [math]::Round(($drive.Used + $drive.Free) / 1GB, 2)
    
    return [pscustomobject]@{
        FreeSpaceGB  = $freeSpaceGB
        TotalSpaceGB = $totalSpaceGB
    }
}

# Function to update a bookmark in a Word document without reopening Word each time
function Update-WordBookmark {
    param (
        [Microsoft.Office.Interop.Word.Application]$wordApp, # Word Application object
        [Microsoft.Office.Interop.Word.Document]$document,   # Word Document object
        [string]$bookmarkName,                               # Name of the bookmark
        [string]$textToInsert                                # Text to insert at the bookmark
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
$docPath = Join-Path -Path $PSScriptRoot -ChildPath "document.docx"   # Replace "document.docx" with your Word document name

# Define bookmarks for Drive C, Drive E, and Windows version
$bookmarkFreeSpaceC = "FreeSpaceCBookmark"
$bookmarkTotalSpaceC = "TotalSpaceCBookmark"
$bookmarkFreeSpaceE = "FreeSpaceEBookmark"
$bookmarkTotalSpaceE = "TotalSpaceEBookmark"
$bookmarkWindowsVersion = "WindowsVersionBookmark"

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

    # Update bookmarks for Drive C
    Update-WordBookmark -wordApp $wordApp -document $document -bookmarkName $bookmarkFreeSpaceC -textToInsert "$($driveSpaceC.FreeSpaceGB) GB"
    Update-WordBookmark -wordApp $wordApp -document $document -bookmarkName $bookmarkTotalSpaceC -textToInsert "$($driveSpaceC.TotalSpaceGB) GB"

    # Update bookmarks for Drive E
    Update-WordBookmark -wordApp $wordApp -document $document -bookmarkName $bookmarkFreeSpaceE -textToInsert "$($driveSpaceE.FreeSpaceGB) GB"
    Update-WordBookmark -wordApp $wordApp -document $document -bookmarkName $bookmarkTotalSpaceE -textToInsert "$($driveSpaceE.TotalSpaceGB) GB"

    # Update the bookmark for Windows version
    Update-WordBookmark -wordApp $wordApp -document $document -bookmarkName $bookmarkWindowsVersion -textToInsert $windowsVersion

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
