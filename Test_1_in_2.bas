Attribute VB_Name = "Test_1_in_2"
Sub Test()
    Dim olddate As String
    Dim newDate As Date
    
    On Error Resume Next
    olddate = "12.12.2024"
    newDate = Replace(olddate, ".", "/")
    thisdate = DateAdd("d", 30, newDate)
   
    
    If Err.Number <> 0 Then
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Error Description: " & Err.Description, vbExclamation, "Error"
    Err.Clear ' Clear the error object
    Else
    MsgBox thisdate
    End If
    On Error GoTo 0 ' Reset error handling
    
End Sub
Sub test2()

Dim olddate As String

    Dim box_valutadni As Double
    Dim box_datumizdaje As String
    Dim datum As Date
    
    box_datumizdaje = "12.12.2024"
    
    
    On Error Resume Next
    datum = CDate(Replace(box_datumizdaje, ".", "/"))
    
If Err.Number <> 0 Then
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Error Description: " & Err.Description, vbExclamation, "Error"
    Err.Clear ' Clear the error object
Else
    box_rokplacila = DateAdd("d", box_valutadni, datum)
    MsgBox box_rokplacila
End If
On Error GoTo 0 ' Reset error handling

    
End Sub


  


