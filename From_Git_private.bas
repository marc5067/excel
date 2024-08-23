Attribute VB_Name = "From_Git_private"
Sub ImportVBAFromGitHub()
    Dim repoOwner As String
    Dim repoName As String
    Dim branchName As String
    Dim token As String
    Dim url As String
    Dim http As Object
    Dim response As String
    Dim fileName As String
    Dim fileContent As String
    Dim ws As Worksheet
    Dim i As Integer
    Dim json As Object

    ' Set your GitHub repository details and personal access token
    repoOwner = "marc5067"  ' Replace with your GitHub repository owner
    repoName = "excel"    ' Replace with your GitHub repository name
    branchName = "work"            ' Replace with the branch name (e.g., "main" or "master")
    token = "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"    ' Replace with your GitHub personal access token

    ' Set up the URL to get the list of files in the repository
    url = "https://api.github.com/repos/" & repoOwner & "/" & repoName & "/contents?ref=" & branchName

    ' Create the XMLHTTP object to perform the HTTP request
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    ' Fetch the list of files
    http.Open "GET", url, False
    http.setRequestHeader "Authorization", "token " & token
    http.send
    response = http.responseText
    
    ' Parse the JSON response
    Set json = JsonConverter.ParseJson(response)
    
    ' Loop through each file in the repository
    For i = 1 To json.Count
        ' Get the file name and download URL
        fileName = json(i)("name")
        fileContent = GetFileContent(json(i)("download_url"), token)
        
        ' Import the file content into a new module or form
        Select Case True
            Case fileName Like "*.bas"
                ' Import standard module
                ImportVBAComponent fileName, fileContent, vbext_ct_StdModule
            Case fileName Like "*.cls"
                ' Import class module
                ImportVBAComponent fileName, fileContent, vbext_ct_ClassModule
            Case fileName Like "*.frm"
                ' Import user form
                ImportVBAComponent fileName, fileContent, vbext_ct_MSForm
        End Select
    Next i
    
    MsgBox "VBA code imported successfully!"
End Sub

Function GetFileContent(downloadUrl As String, token As String) As String
    Dim http As Object
    Dim response As String
    
    ' Create the XMLHTTP object to perform the HTTP request
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    ' Fetch the file content
    http.Open "GET", downloadUrl, False
    http.setRequestHeader "Authorization", "token " & token
    http.send
    response = http.responseText
    
    GetFileContent = response
End Function

Sub ImportVBAComponent(fileName As String, fileContent As String, componentType As vbext_ComponentType)
    Dim vbComp As VBComponent
    Dim tempFilePath As String
    
    ' Create a temporary file path to save the component content
    tempFilePath = Environ("TEMP") & "\" & fileName
    
    ' Save the content to the temporary file
    Dim fileNumber As Integer
    fileNumber = FreeFile
    Open tempFilePath For Output As #fileNumber
    Print #fileNumber, fileContent
    Close #fileNumber
    
    ' Import the component from the temporary file
    Set vbComp = ThisWorkbook.VBProject.VBComponents.Import(tempFilePath)
    
    ' Clean up temporary file
    Kill tempFilePath
End Sub

