Attribute VB_Name = "To_2_GitHub"
Sub ExportVBAAndPushToGitHub()
    Dim vbComp As VBComponent
    Dim repoPath As String
    Dim gitCommand As String
    Dim shellResult As Variant
    Dim fileName As String
    Dim filePath As String
    Dim fileNum As Integer
    
    ' Define the path to your local Git repository
    repoPath = "C:\Users\marce\OneDrive\eva_excel\v2\repo\"  ' Ensure the trailing backslash
    
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
    
    ' Export each VBA component
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Select Case vbComp.Type
            Case vbext_ct_StdModule, vbext_ct_ClassModule
                ' Export the component as .bas file
                filePath = repoPath & vbComp.Name & ".bas"
                vbComp.Export filePath
            
            Case vbext_ct_MSForm
                ' Export the component as .bas and .frx files
                filePath = repoPath & vbComp.Name & ".bas"
                vbComp.Export filePath
                
                ' Create .frx file for the form
                filePath = repoPath & vbComp.Name & ".frx"
                fileNum = FreeFile
                Open filePath For Binary Access Write As #fileNum
                Close #fileNum  ' Creating an empty file as placeholder
        End Select
    Next vbComp
    
    ' Git add command to stage all changes
    gitCommand = "cmd.exe /c cd """ & repoPath & """ && git add ."
    shellResult = Shell(gitCommand, vbNormalFocus)
    
    ' Git commit command to commit changes
    gitCommand = "cmd.exe /c cd """ & repoPath & """ && git commit -m ""Updated VBA code from Excel"""
    shellResult = Shell(gitCommand, vbNormalFocus)
    
    ' Git push command to push changes to the remote repository
    gitCommand = "cmd.exe /c cd """ & repoPath & """ && git push origin main"
    shellResult = Shell(gitCommand, vbNormalFocus)
    
    MsgBox "VBA code exported and pushed to GitHub successfully!"
End Sub

