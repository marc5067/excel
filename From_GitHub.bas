Attribute VB_Name = "From_GitHub"
Sub PullAllVBAFromGitHubAndImport()
    Dim repoUrl As String
    Dim localPath As String
    Dim filesList As Object
    Dim http As Object
    Dim file As Object
    Dim fileName As String
    Dim fileUrl As String
    Dim filePath As String
    Dim fileNum As Integer

    ' Define the GitHub repository URL and local path
    repoUrl = "https://api.github.com/repos/marc5067/excel/contents/" ' Adjust to your repo
    localPath = "C:\Users\marce\OneDrive\eva_excel\v2\repo\"  ' Local path to temporarily save the downloaded files
    
        ' Ensure the export path exists
    If Dir(repoPath, vbDirectory) = "" Then
        MsgBox "Repository path does not exist. Please check the path and try again."
        Exit Sub
    End If
    
        ' Delete all files in the repository folder
    fileName = Dir(repoPath & "*.*")
    Do While fileName <> ""
        Kill repoPath & fileName
        fileName = Dir
    Loop
    
    ' Create an instance of the MSXML2.XMLHTTP object
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    ' Make an API request to get the list of files in the repository
    http.Open "GET", repoUrl, False
    http.send
    
    ' Check if the request was successful
    If http.Status = 200 Then
        ' Parse the JSON response to get file names
        Set filesList = JsonConverter.ParseJson(http.responseText)
        
        ' Loop through the files in the repository
        For Each file In filesList
            fileName = file("name")
            fileUrl = file("download_url")
            filePath = localPath & fileName
            
            ' Download the file
            http.Open "GET", fileUrl, False
            http.send
            
            ' Save the file locally
            If http.Status = 200 Then
                fileNum = FreeFile
                Open filePath For Binary As #fileNum
                Put #fileNum, 1, http.responseBody
                Close #fileNum
                
                ' Import the file if it's not an .frx file (which is binary data for forms)
                If Right(fileName, 4) <> ".frx" Then
                    ThisWorkbook.VBProject.VBComponents.Import filePath
                End If
            Else
                MsgBox "Failed to download " & fileName, vbCritical
            End If
        Next file
        
        MsgBox "VBA code pulled from GitHub and imported into Excel successfully!"
    Else
        MsgBox "Failed to retrieve file list from GitHub", vbCritical
    End If
End Sub

