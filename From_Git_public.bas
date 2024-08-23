Attribute VB_Name = "From_Git_public"
Function DownloadVBAFromGitHub(ByVal GitHubRawURL As String) As String
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")

    On Error GoTo ErrHandler
    http.Open "GET", GitHubRawURL, False
    http.send

    If http.Status = 200 Then
        DownloadVBAFromGitHub = http.responseText
    Else
        DownloadVBAFromGitHub = "Error: " & http.Status & " - " & http.statusText
    End If

    Exit Function
ErrHandler:
    DownloadVBAFromGitHub = "Error: " & Err.Description
End Function


Sub UpdateVBAFromGitHub()
    Dim code As String
    Dim moduleName As String
    Dim vbComp As VBComponent
    
    ' Define the URL to your raw file on GitHub
    Dim GitHubURL As String
    GitHubURL = "https://raw.githubusercontent.com/marc5067/excel/main/module2.bas"
    
    ' Download the code
    code = DownloadVBAFromGitHub(GitHubURL)
    
    If Left(code, 5) = "Error" Then
        MsgBox code
        Exit Sub
    End If
    
    ' Define the module you want to replace (or add)
    moduleName = "Module2"
    
    ' Check if the module exists, and remove it if it does
    On Error Resume Next
    Set vbComp = ThisWorkbook.VBProject.VBComponents(moduleName)
    If Not vbComp Is Nothing Then
        ThisWorkbook.VBProject.VBComponents.Remove vbComp
    End If
    On Error GoTo 0
    
    ' Add a new module and insert the downloaded code
    Set vbComp = ThisWorkbook.VBProject.VBComponents.Add(vbext_ct_StdModule)
    vbComp.Name = moduleName
    vbComp.CodeModule.AddFromString code
    
    MsgBox "VBA code has been updated successfully."
End Sub

