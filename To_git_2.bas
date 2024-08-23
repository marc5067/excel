Attribute VB_Name = "To_git_2"
Sub ExportVBAAndPushToGitHub()
    Dim vbComp As VBComponent
    Dim repoPath As String
    Dim gitCommand As String
    Dim shellResult As Variant
    
    ' Define the path to your local Git repository
    repoPath = "D:\PROGRAMING\excel\"  ' Path to your local Git repository
    
    ' Ensure the export path exists
    If Dir(repoPath, vbDirectory) = "" Then
        MsgBox "Repository path does not exist. Please check the path and try again."
        Exit Sub
    End If
    
    ' Export each VBA component that is not a sheet or workbook module
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Select Case vbComp.Type
            Case vbext_ct_StdModule, vbext_ct_ClassModule, vbext_ct_MSForm
                ' Export the component
                vbComp.Export repoPath & vbComp.Name & ".bas"
        End Select
    Next vbComp
    
    ' Git add command to stage all changes
    gitCommand = "cmd.exe /c cd " & repoPath & " && git add ."
    shellResult = Shell(gitCommand, vbNormalFocus)
    
    ' Check for errors in the git add command
    If shellResult <> 0 Then
        MsgBox "Failed to add files to Git. Check the command output for details."
        Exit Sub
    End If
    
    ' Git commit command to commit changes
    gitCommand = "cmd.exe /c cd " & repoPath & " && git commit -m ""Updated VBA code from Excel"""
    shellResult = Shell(gitCommand, vbNormalFocus)
    
    ' Check for errors in the git commit command
    If shellResult <> 0 Then
        MsgBox "Failed to commit changes to Git. Check the command output for details."
        Exit Sub
    End If
    
    ' Git push command to push changes to the remote repository
    gitCommand = "cmd.exe /c cd " & repoPath & " && git push origin work" ' Replace "work" with your branch name
    shellResult = Shell(gitCommand, vbNormalFocus)
    
    ' Check for errors in the git push command
    If shellResult <> 0 Then
        MsgBox "Failed to push changes to GitHub. Check the command output for details."
        Exit Sub
    End If
    
    MsgBox "VBA code exported and pushed to GitHub successfully!"
End Sub

