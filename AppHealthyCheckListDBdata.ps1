    # Load .NET SQL types
Add-Type -AssemblyName "System.Data"

# Define SQL Server connection details
$serverName = "$($env:COMPUTERNAME)"  # Adjust instance name if necessary
$database = "eOrganization"           # Database name
$username = "farzindb"                # Replace with your SQL Server username
$password = "123@com"                 # Replace with your SQL Server password

# Connection string
$connectionString = "Server=$serverName;Database=$database;Integrated Security=False;User ID=$username;Password=$password;"

# SQL query
$query = @"
SELECT 
      [MaxSearchResult],
      [RecordPerPage],
      [TopResultCategory],
      CASE 
          WHEN [SendSearchType] = 0 THEN N'غیر اتومات'
          WHEN [SendSearchType] = 1 THEN N'اتومات'
          ELSE N'0'
      END AS [SendSearchTypeDescription],
      CASE 
          WHEN [EnableSessionTimeOut] = 0 THEN N'غیر فعال'
          WHEN [EnableSessionTimeOut] = 1 THEN N'فعال'
          ELSE N'0'
      END AS [EnableSessionTimeOutDescription],
      [SessionTimeOut]
FROM [eOrganization].[dbo].[System_Settings]
WHERE [Code] = 1;
"@

# Function to update a bookmark in a Word document
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

try {
    # Connect to SQL Server and execute query
    $connection = New-Object System.Data.SqlClient.SqlConnection($connectionString)
    $connection.Open()

    $command = $connection.CreateCommand()
    $command.CommandText = $query

    $adapter = New-Object System.Data.SqlClient.SqlDataAdapter $command
    $datatable = New-Object System.Data.DataTable
    $adapter.Fill($datatable)

    # Create Word Application COM object and open the document
    $wordApp = New-Object -ComObject Word.Application
    if ($null -eq $wordApp) {
        throw "Failed to initialize Word COM object."
    }

    $wordApp.Visible = $false
    $document = $wordApp.Documents.Open($docPath)
    if ($null -eq $document) {
        throw "Failed to open the document at path: $docPath"
    }

    # Loop through the data rows and update bookmarks with SQL query results
    foreach ($row in $datatable.Rows) {
        Update-WordBookmark -wordApp $wordApp -document $document -bookmarkName "MaxSearchResultBookmark" -textToInsert "$($row.MaxSearchResult)"
        Update-WordBookmark -wordApp $wordApp -document $document -bookmarkName "RecordPerPageBookmark" -textToInsert "$($row.RecordPerPage)"
        Update-WordBookmark -wordApp $wordApp -document $document -bookmarkName "TopResultCategoryBookmark" -textToInsert "$($row.TopResultCategory)"
        Update-WordBookmark -wordApp $wordApp -document $document -bookmarkName "SendSearchTypeBookmark" -textToInsert "$($row.SendSearchTypeDescription)"
        Update-WordBookmark -wordApp $wordApp -document $document -bookmarkName "EnableSessionTimeOutBookmark" -textToInsert "$($row.EnableSessionTimeOutDescription)"
        Update-WordBookmark -wordApp $wordApp -document $document -bookmarkName "SessionTimeOutBookmark" -textToInsert "$($row.SessionTimeOut)"
    }

    # Save and close the document
    $document.Save()
    $document.Close()
}
catch {
    Write-Host "An error occurred: $_"
}
finally {
    # Clean up
    if ($connection.State -eq 'Open') {
        $connection.Close()
    }

    if ($wordApp) {
        $wordApp.Quit()
    }
}
