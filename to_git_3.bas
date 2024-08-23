Attribute VB_Name = "to_git_3"
Sub ExportVBAAndPushToGitHub()
    Dim vbComp As VBComponent
    Dim repoPath As String
    Dim gitCommand As String
    Dim shellResult As Variant
    Dim filePath As String
    
    ' Define the path to your local Git repository
    repoPath = "D:\PROGRAMING\excel\"  ' Path to your local Git repository
    
    ' Ensure the export path exists
    If Dir(repoPath, vbDirectory) = "" Then
        MsgBox "Repository path does not exist. Please check the path and try again."
        Exit Sub
    End If
    
    ' Export each VBA component, handling extensions appropriately
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        ' Define the file path for the component
        filePath = repoPath & vbComp.Name
        
        Select Case vbComp.Type
            Case vbext_ct_StdModule
                ' Remove existing .bas files to avoid duplication
                If Dir(filePath & ".bas") <> "" Then Kill filePath & ".bas"
                ' Export standard modules with .bas extension
                vbComp.Export filePath & ".bas"
                
            Case vbext_ct_ClassModule
                ' Remove existing .cls files to avoid duplication
                If Dir(filePath & ".cls") <> "" Then Kill filePath & ".cls"
                ' Export class modules with .cls extension
                vbComp.Export filePath & ".cls"
                
            Case vbext_ct_MSForm
                ' Remove existing .frm files to avoid duplication
                If Dir(filePath & ".frm") <> "" Then Kill filePath & ".frm"
                ' Export user forms with .frm extension
                vbComp.Export filePath & ".frm"
                
                ' Remove existing .frx files to avoid duplication
                If Dir(filePath & ".frx") <> "" Then Kill filePath & ".frx"
                ' The .frx file is automatically created during the .frm export
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
