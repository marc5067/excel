Attribute VB_Name = "Test_GitHub_refrences"
'v tools/refrences moras vklopiti Microsoft Visual Basic for Applications Extensibility 5.3
'runnas ta program
'excel/file/trust center/trust center settings/macro settings obklukaj trust access to the vba project object model
Sub CheckVBALibraryReference()
    Dim ref As Reference
    For Each ref In ThisWorkbook.VBProject.References
        If ref.Name = "VBIDE" Then
            MsgBox "Reference is enabled."
            Exit Sub
        End If
    Next ref
    MsgBox "Reference is not enabled. Please enable 'Microsoft Visual Basic for Applications Extensibility 5.3'."
End Sub

