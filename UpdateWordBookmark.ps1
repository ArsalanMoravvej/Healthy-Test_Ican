# Function to get the Windows version
function Get-WindowsVersion {
    $os = Get-CimInstance -ClassName Win32_OperatingSystem
    return "$($os.Caption) - Version $($os.Version)"
}

# Function to get the free and total space of drive C separately
function Get-DriveSpace {
    $drive = Get-PSDrive -Name C
    if ($null -eq $drive) {
        throw "Drive C: not found."
    }

    $freeSpaceGB = [math]::Round($drive.Free / 1GB, 2)
    $totalSpaceGB = [math]::Round($drive.Used / 1GB + $freeSpaceGB, 2)
    
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
$bookmarkFreeSpace = "FreeSpaceBookmark"                              # Replace with the bookmark for free space
$bookmarkTotalSpace = "TotalSpaceBookmark"                            # Replace with the bookmark for total space
$bookmarkWindowsVersion = "WindowsVersionBookmark"                    # Replace with the bookmark for Windows version

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
    $driveSpace = Get-DriveSpace
    $windowsVersion = Get-WindowsVersion

    # Update bookmarks
    Update-WordBookmark -wordApp $wordApp -document $document -bookmarkName $bookmarkFreeSpace -textToInsert "$($driveSpace.FreeSpaceGB) GB"
    Update-WordBookmark -wordApp $wordApp -document $document -bookmarkName $bookmarkTotalSpace -textToInsert "$($driveSpace.TotalSpaceGB) GB"
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
