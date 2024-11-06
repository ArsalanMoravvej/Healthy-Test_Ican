# Load .NET SQL types
Add-Type -AssemblyName "System.Data"

# Define SQL Server connection details
$serverName = "$($env:COMPUTERNAME)"  # Adjust instance name if necessary
$username = "farzindb"                # Replace with your SQL Server username
$password = "123@com"                 # Replace with your SQL Server password

# Possible database names
$databases = @("eOrganization", "ICANSBPMS")

# SQL query
$query = @"
SELECT 
      [MaxSearchResult],
      [RecordPerPage],
      [TopResultCategory],
      CASE 
          WHEN [SendSearchType] = 0 THEN N'غیر اتومات'
          WHEN [SendSearchType] = 1 THEN N'اتومات'
          ELSE N'نامشخص'
      END AS [SendSearchTypeDescription],
      CASE 
          WHEN [EnableSessionTimeOut] = 0 THEN N'غیر فعال'
          WHEN [EnableSessionTimeOut] = 1 THEN N'فعال'
          ELSE N'نامشخص'
      END AS [EnableSessionTimeOutDescription],
      [SessionTimeOut]
FROM [dbo].[System_Settings]
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

# Function to attempt connection to a list of databases
function Connect-ToDatabase {
    param (
        [array]$dbNames,     # List of possible database names
        [string]$serverName, # Server name
        [string]$username,   # Username
        [string]$password    # Password
    )

    foreach ($dbName in $dbNames) {
        $connectionString = "Server=$serverName;Database=$dbName;Integrated Security=False;User ID=$username;Password=$password;"
        
        # Try connecting
        $connection = New-Object System.Data.SqlClient.SqlConnection($connectionString)
        try {
            $connection.Open()
            if ($connection.State -eq "Open") {
                Write-Host "Connected to database '$dbName' successfully."
                return $connection  # Return the successful connection
            }
        }
        catch {
            Write-Host "Could not connect to database '$dbName'. Trying next..."
        }
    }
    throw "Unable to connect to any specified database."
}

# Main script
$docPath = Join-Path -Path $PSScriptRoot -ChildPath "ApplicationServerHealthyCheckList.docx"

try {
    # Attempt to connect to one of the databases
    $connection = Connect-ToDatabase -dbNames $databases -serverName $serverName -username $username -password $password

    # Create SQL command
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
