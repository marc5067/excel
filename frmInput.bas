VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInput 
   Caption         =   "Input Form"
   ClientHeight    =   9645.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19365
   OleObjectBlob   =   "frmInput.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public uredi As Integer
Public st As Integer
Public myrow

Public sheetDela As Worksheet
Public sheetStranke As Worksheet
Public sheetRacuni As Worksheet
Public sheetEvidence As Worksheet
Public sheetKopija As Worksheet
Public once As Integer

Private Sub btn_cancel_Click()
 Call UserForm_Initialize
 my_pages_01.Value = "2"
 my_pages_02.Value = "0"
End Sub

Private Sub btn_delete_Click()
    Set sheetStranke = Sheets("Database_stranke")
    Row = sheetStranke.Range("B:B").Find(box_new_name).Row
    sheetStranke.Rows(Row).Delete
    Call UserForm_Initialize
    my_pages_01.Value = "2"
    my_pages_02.Value = "0"
End Sub

Private Sub btn_ustvari_izpisek_Click()
    my_pages_01.Value = "3"
    
   
End Sub

Private Sub btn_topdf_Click()
    Set sheetKopija = Sheets("KOPIJA")
    Set sheetRacuni = Sheets("Racuni")
    
    strac = ListBox1.Value
    Row = sheetRacuni.Range("A:A").Find(strac).Row
    
    'kopira vrednosti v kopijo
    sheetKopija.Cells(9, "A").Value = sheetRacuni.Cells(Row, 3).Value
    sheetKopija.Cells(10, "A").Value = sheetRacuni.Cells(Row, 4).Value
    sheetKopija.Cells(11, "A").Value = sheetRacuni.Cells(Row, 5).Value
    sheetKopija.Cells(13, "B").Value = sheetRacuni.Cells(Row, 6).Value
    
    sheetKopija.Cells(16, "B").Value = sheetRacuni.Cells(Row, 1).Value
    
    sheetKopija.Cells(20, "A").Value = sheetRacuni.Cells(Row, 7).Value
    sheetKopija.Cells(20, "C").Value = sheetRacuni.Cells(Row, 9).Value
    sheetKopija.Cells(20, "E").Value = sheetRacuni.Cells(Row, 8).Value
    
    sheetKopija.Cells(8, "F").Value = sheetRacuni.Cells(Row, 11).Value
    sheetKopija.Cells(10, "F").Value = sheetRacuni.Cells(Row, 12).Value
    sheetKopija.Cells(11, "F").Value = sheetRacuni.Cells(Row, 13).Value
    sheetKopija.Cells(12, "F").Value = sheetRacuni.Cells(Row, 14).Value
    
    'exporta as pdf
    Dim sheetToPdf As Worksheet
    Dim savePath As String
    Dim pdfname As String

    ' Set the worksheet you want to export to PDF
    Set sheetToPdf = ThisWorkbook.Sheets("KOPIJA")
    
    pdfname = sheetKopija.Cells(34, "B").Value
    
    'Specify the path where you want to save the PDF
        'pogleda kater operacijski sistem je
    If InStr(1, Application.OperatingSystem, "Windows") > 0 Then
    savePath = ThisWorkbook.Path & "\" & pdfname & ".pdf"
    Else
    savePath = ThisWorkbook.Path & "/" & pdfname & ".pdf"
    End If
    
    ' Export the worksheet as PDF
    sheetToPdf.ExportAsFixedFormat Type:=xlTypePDF, fileName:=savePath, Quality:=xlQualityStandard

    ' Notify the user once the export is complete
    MsgBox "PDF saved at: " & savePath, vbInformation
End Sub

Private Sub btn_reset_1_Click()
    Call UserForm_Initialize
End Sub

Private Sub btn_reset_2_Click()
    Call UserForm_Initialize
End Sub

Private Sub btn_uredistranko_Click()
    Set sheetStranke = Sheets("Database_stranke")
    
    st_stranke = ListBox2.Value
    If (IsNumeric(st_stranke)) Then
    
    myrow = sheetStranke.Range("A:A").Find(st_stranke).Row
    
    my_pages_02.Value = "1"
    
    If Not (sheetStranke.Cells(myrow, 5).Value = "") Then
    box_new_davcnast.Visible = True
    toggle_pravna = True
    End If
    
    btn_reset_2.Visible = False
    
    btn_reset_2.Visible = False
    btn_delete.Visible = True
    
    box_new_name = sheetStranke.Cells(myrow, 2).Value
    box_new_naslov = sheetStranke.Cells(myrow, 3).Value
    box_new_postnast = sheetStranke.Cells(myrow, 4).Value
    box_new_davcnast = sheetStranke.Cells(myrow, 5).Value
    
    myjob = sheetStranke.Cells(myrow, "F").Value
    Dim optButton As Variant
    For Each optButton In Array("option_1", "option_2", "option_3", "option_4")
        If Me.Controls(optButton).Caption = myjob Then
            Me.Controls(optButton).Value = True
            Exit For
        End If
    Next optButton
    
    uredi = 1
    ListBox2.Value = Null
    End If
End Sub

Private Sub button_new_confirm_Click()
    'precekira ce si vnesel podatke
    
    check1 = IsDate(box_datumizdaje.Value)
    check2 = IsDate(box_datumstoritve.Value)
    check3 = IsDate(box_rokplacila.Value)
    check4 = IsDate(box_valutadni.Value)
    
    check5 = IsNumeric(box_cena.Value)
    check6 = IsNumeric(box_kolicinaur.Value)
    
    If Not box_new_name = "" And Not box_new_naslov = "" And Not box_new_postnast = "" Then
        
        Dim optButton As Variant
    
        If (uredi = 0) Then

            Dim lastRow As Long
    
            lastRow = sheetStranke.Cells(sheetStranke.Rows.Count, "A").End(xlUp).Row
            firstEmptyRowNum = lastRow + 1
    
            lastNum = sheetStranke.Cells(lastRow, "A").Value
            If Not IsNumeric(lastNum) Then
                lastNum = 0
            End If
    
            lastNum = lastNum + 1
    
            sheetStranke.Cells(firstEmptyRowNum, "A").Value = lastNum
            sheetStranke.Cells(firstEmptyRowNum, "B").Value = box_new_name.Value
            sheetStranke.Cells(firstEmptyRowNum, "C").Value = box_new_naslov.Value
            sheetStranke.Cells(firstEmptyRowNum, "D").Value = box_new_postnast.Value
    
            For Each optButton In Array("option_1", "option_2", "option_3", "option_4")
                If Me.Controls(optButton).Value = True Then
                    sheetStranke.Cells(firstEmptyRowNum, "F").Value = Me.Controls(optButton).Caption
                Exit For
                End If
            Next optButton
    
            If toggle_pravna = True Then
                sheetStranke.Cells(firstEmptyRowNum, "E").Value = box_new_davcnast.Value
            End If
    
        Else
            sheetStranke.Cells(myrow, "B").Value = box_new_name.Value
            sheetStranke.Cells(myrow, "C").Value = box_new_naslov.Value
            sheetStranke.Cells(myrow, "D").Value = box_new_postnast.Value
         
            For Each optButton In Array("option_1", "option_2", "option_3", "option_4")
            If Me.Controls(optButton).Value = True Then
                sheetStranke.Cells(myrow, "F").Value = Me.Controls(optButton).Caption
                Exit For
            End If
            Next optButton
    
            If toggle_pravna = True Then
            sheetStranke.Cells(myrow, "E").Value = box_new_davcnast.Value
            End If

        End If
        
        my_pages_02.Value = "0"
    
        Call UserForm_Initialize
        
        'TUKI NI TREBA TOK VELIKEGA IF NAREST LOH SAM IF PA POL SPREMENIS ST VRSTICE !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    
    Else
        Dim msgValue As VbMsgBoxResult
        msgValue = MsgBox("Vpisi vse podatke!")
        
    End If
    
End Sub


Private Sub combo_filter_racuni_Change()
    Select Case combo_filter_racuni.Value
    Case "Leto"
    lbl_izpisek.Caption = "Leto:"
    myColumn = "K"
    Case "Stranka"
    lbl_izpisek.Caption = "Stranka:"
    myColumn = "C"
    Case "Mesec"
    lbl_izpisek.Caption = "Mesec:"
    myColumn = "K"
    End Select
    
    combo_izberi_racuni.Clear
    lastRow = sheetRacuni.Cells(sheetRacuni.Rows.Count, "A").End(xlUp).Row
    Dim rng2 As Range

    myrg = myColumn + "2:" + myColumn
    Set rng2 = sheetRacuni.Range(myrg & lastRow)
    rng2.Select
    
    Dim cell2 As Range
    For Each cell2 In rng2
    Select Case combo_filter_racuni.Value
    Case "Leto"
        datum = CDate(Replace(cell2.Value, ".", "/"))
        mysel = Year(datum)
    Case "Stranka"
        mysel = cell2.Value
    Case "Mesec"
        datum = CDate(Replace(cell2.Value, ".", "/"))
        mysel = Month(toDate(cell2.Value))
    End Select
        combo_izberi_racuni.AddItem mysel
    Next cell2
    
    
End Sub

Private Sub combo_izberi_racuni_Change()
Select Case lbl_izpisek.Caption
    Case "Leto"
        myColumn = "K"
    Case "Stranka"
        myColumn = "C"
    Case "Mesec"
        myColumn = "K"
    End Select
    
    ListBox3.Clear
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    lastRow = sheetRacuni.Cells(sheetRacuni.Rows.Count, "A").End(xlUp).Row
    
    Dim rng2 As Range
    myrg = myColumn + "2:" + myColumn
    Set rng2 = sheetRacuni.Range(myrg & lastRow)
    rng2.Select
    
    'Dim cell2 As Range
    'For Each cell2 In rng2
    'If cell2.Value = combo_izberi_racuni.Value Then
    '    combo_izberi_racuni.AddItem cell2
    'End If
    'Next cell2
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ListBox3.AddItem
End Sub

Private Sub toggle_pravna_Change()
    If toggle_pravna = True Then
    box_new_davcnast.Visible = True
    lbl_davcnast.Visible = True
    Else
    box_new_davcnast.Visible = False
    lbl_davcnast.Visible = False
    End If
    
End Sub


Private Sub btn_dodajstranko_Click()
    my_pages_02.Value = "1"
End Sub

Private Sub btn_izbrisivnos_Click()
   
    strac = ListBox1.Value
    
    Row = sheetRacuni.Range("A:A").Find(strac).Row
    
    sheetEvidence.Rows(Row).Delete
    sheetRacuni.Rows(Row).Delete
    
    Call UserForm_Initialize
    
End Sub

Private Sub btn_now_racun_Click()
    my_pages_01.Value = "0"
End Sub

Private Sub btn_racuni_Click()
    my_pages_01.Value = "1"
End Sub

Private Sub btn_stranke_Click()
    my_pages_01.Value = "2"
End Sub

Private Sub button_confirm_Click()
    'precekiras ce je vse napisano v box-e
    check1 = IsDate(Replace(box_datumizdaje.Value, ".", "/"))
    check2 = IsValidDateRange(box_datumstoritve.Value)
    check3 = IsDate(Replace(box_rokplacila.Value, ".", "/"))
    check4 = IsNumeric(box_valutadni.Value)
        
    check5 = IsNumeric(box_cena.Value)
    check6 = IsNumeric(box_kolicinaur.Value)
   
    If check1 And check2 And check3 And check4 And check5 And check6 And Not choose_stranko.Value = "Izberi stranko" Then
    
        'stevilka zadnje vrstice
    firstEmptyRowNum = sheetRacuni.Cells(sheetRacuni.Rows.Count, "A").End(xlUp).Row + 1
    
        'vstavimo novo vrstico
    sheetRacuni.Rows(firstEmptyRowNum).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

    'kopiras vrednosti v racuni
    stranka = sheetStranke.Range("B:B").Find(choose_stranko.Value).Offset(0, -1).Value
        
    sheetRacuni.Cells(firstEmptyRowNum, "A").Value = box_strac.Value
    sheetRacuni.Cells(firstEmptyRowNum, "B").Value = "SI00 " & stranka & "-" & box_strac.Value
    sheetRacuni.Cells(firstEmptyRowNum, "C").Value = choose_stranko.Value
    sheetRacuni.Cells(firstEmptyRowNum, "D").Value = box_naslov.Value
    sheetRacuni.Cells(firstEmptyRowNum, "E").Value = box_postnast.Value
    sheetRacuni.Cells(firstEmptyRowNum, "F").Value = box_ddv.Value
    sheetRacuni.Cells(firstEmptyRowNum, "G").Value = choose_delo.Value
    sheetRacuni.Cells(firstEmptyRowNum, "H").Value = box_cena.Value
    sheetRacuni.Cells(firstEmptyRowNum, "I").Value = box_kolicinaur.Value
    sheetRacuni.Cells(firstEmptyRowNum, "J").Value = box_cena.Value * box_kolicinaur.Value
    sheetRacuni.Cells(firstEmptyRowNum, "K").Value = box_datumizdaje.Value
    sheetRacuni.Cells(firstEmptyRowNum, "L").Value = box_datumstoritve.Value
    sheetRacuni.Cells(firstEmptyRowNum, "M").Value = box_valutadni.Value
    sheetRacuni.Cells(firstEmptyRowNum, "N").Value = box_rokplacila.Value
    
    'kopiras vrednosti v evidence
        'stevilka zadnje vrstice
    firstEmptyRowNum = sheetEvidence.Cells(sheetEvidence.Rows.Count, "A").End(xlUp).Row + 2
    
        'vstavimo novo vrstico
    sheetEvidence.Rows(firstEmptyRowNum).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

        'kopiras vrednosti
    sheetEvidence.Cells(firstEmptyRowNum, "A").Value = box_strac.Value - 24000
    sheetEvidence.Cells(firstEmptyRowNum, "B").Value = box_strac.Value
    sheetEvidence.Cells(firstEmptyRowNum, "C").Value = box_datumizdaje.Value
    sheetEvidence.Cells(firstEmptyRowNum, "D").Value = box_datumstoritve.Value
    sheetEvidence.Cells(firstEmptyRowNum, "E").Value = choose_stranko.Value
    sheetEvidence.Cells(firstEmptyRowNum, "F").Value = box_naslov.Value
    sheetEvidence.Cells(firstEmptyRowNum, "G").Value = box_ddv.Value
    sheetEvidence.Cells(firstEmptyRowNum, "H").Value = box_rokplacila.Value
    sheetEvidence.Cells(firstEmptyRowNum, "J").Value = box_cena.Value * box_kolicinaur.Value
    
    Dim msgValue As VbMsgBoxResult
    msgValue = MsgBox("Racun shranjen")
    
    Call UserForm_Initialize
    
    Else
    Dim msgValue1 As VbMsgBoxResult
    msgValue1 = MsgBox("Vpsis vse podatke!")
    
    End If
    
End Sub

Private Sub choose_delo_Change()
   
    If Not choose_delo.Value = ". . ." Then
    Dim Job As String
    Job = choose_delo.Value
    Dim cena As String
    cena = sheetDela.Range("A:A").Find(Job).Offset(0, 1).Value
    box_cena.Value = cena
    End If
End Sub

Private Sub choose_stranko_Change()
    If Not choose_stranko.Value = "Izberi stranko" Then
    
    Dim customerName As String
    customerName = choose_stranko.Value
    
    Dim naslov As String
    Dim postst As String
    Dim ddv As String
    Dim delo As String
    
    naslov = sheetStranke.Range("B:B").Find(customerName).Offset(0, 1).Value
    postst = sheetStranke.Range("B:B").Find(customerName).Offset(0, 2).Value
    ddv = sheetStranke.Range("B:B").Find(customerName).Offset(0, 3).Value
    delo = sheetStranke.Range("B:B").Find(customerName).Offset(0, 4).Value
    
    box_naslov.Value = naslov
    box_postnast.Value = postst
    box_ddv.Value = ddv
    choose_delo.Value = delo
    
    End If
    
End Sub

Private Sub UserForm_Initialize()

        'full screen mode
    'With Application
    '.WindowState = xlMaximized
    'Zoom = Int(.Width / Me.Width * 100)
    'Width = .Width
    'Height = .Height
    'End With
    
    If (once = 0) Then
        Call initializeVariables
    End If
   
    once = 1

        'najdi zadnjo prazno vrstico
    lastRowRacuni = sheetRacuni.Cells(sheetRacuni.Rows.Count, "A").End(xlUp).Row
    lastRowEvidence = sheetEvidence.Cells(sheetEvidence.Rows.Count, "H").End(xlUp).Row
    lastRowStranke = sheetStranke.Cells(sheetStranke.Rows.Count, "A").End(xlUp).Row
    lastRowDela = sheetDela.Cells(sheetDela.Rows.Count, "A").End(xlUp).Row

        'dodamo st racuna
    lastValue = sheetRacuni.Cells(lastRowRacuni, 1).Value
    If Not IsNumeric(lastValue) Then
        lastValue = 24000
    End If
    box_strac.Value = lastValue + 1
   
        'dodamo stranke
        
    choose_stranko.Clear
    
    Dim rng1 As Range
    Set rng1 = sheetStranke.Range("B3:B" & lastRowStranke)
    
    Dim cell1 As Range
    For Each cell1 In rng1
        choose_stranko.AddItem cell1.Value
    Next cell1
    
        'dodamo dela
    choose_delo.Clear
    Dim rng2 As Range
    Set rng2 = sheetDela.Range("A3:A" & lastRowDela)
    
    Dim cell2 As Range
    For Each cell2 In rng2
        choose_delo.AddItem cell2.Value
    Next cell2
    
        'default stranka in delo
    choose_stranko.Value = "Izberi stranko"
    choose_delo.Value = ". . ."
    
        'davcna ni vidna
    box_new_davcnast.Visible = False
    lbl_davcnast.Visible = False
    
        'vpisemo podatke v listbox
    ListBox1.ColumnCount = "14"
    ListBox1.ColumnHeads = True
    ListBox1.RowSource = "Racuni!A2:N" & lastRowRacuni
    
    ListBox2.ColumnCount = "6"
    ListBox2.ColumnHeads = True
    ListBox2.RowSource = "Database_stranke!A3:F" & lastRowStranke
    
        'resetira nov racun
    box_kolicinaur.Value = Null
    box_datumizdaje.Value = Null
    box_datumstoritve.Value = Null
    box_valutadni.Value = Null
    box_rokplacila.Value = Null
    box_cena.Value = Null
    choose_stranko.Value = "Izberi stranko"
    box_naslov.Value = Null
    box_postnast.Value = Null
    box_ddv.Value = Null
    
    box_new_name = Null
    box_new_naslov = Null
    box_new_postnast = Null
    box_new_davcnast = Null
    
    'default strani
    If Not st = 1 Then
    my_pages_01.Value = "0"
    my_pages_02.Value = "0"
    End If
    
    st = 1
    
    toggle_pravna = False
    btn_reset_2.Visible = True
    btn_delete.Visible = False
    uredi = 0
    
    combo_filter_racuni.AddItem "Leto"
    combo_filter_racuni.AddItem "Stranka"
    combo_filter_racuni.AddItem "Mesec"
 
End Sub

Sub iz_confirm_button_click()
    
    'PREPISES V KOPIJO
    sheetKopija.Cells(9, "A").Value = choose_stranko.Value
    sheetKopija.Cells(10, "A").Value = box_naslov.Value
    sheetKopija.Cells(11, "A").Value = box_postnast.Value
    sheetKopija.Cells(13, "B").Value = box_ddv.Value
    
    sheetKopija.Cells(16, "B").Value = box_strac.Value
    
    sheetKopija.Cells(20, "A").Value = choose_delo.Value
    sheetKopija.Cells(20, "C").Value = box_kolicinaur.Value
    sheetKopija.Cells(20, "E").Value = box_cena.Value
    
    sheetKopija.Cells(8, "F").Value = box_datumizdaje.Value
    sheetKopija.Cells(10, "F").Value = box_datumstoritve.Value
    sheetKopija.Cells(11, "F").Value = box_valutadni.Value
    sheetKopija.Cells(12, "F").Value = box_rokplacila.Value
        
    Set mycell = sheetKopija.Range("B16")
    mycell.ClearContents
    Set mycell = sheetKopija.Range("F8:F12")
    mycell.ClearContents
    Set mycell = sheetKopija.Range("C20")
    mycell.ClearContents
    Set mycell = sheetKopija.Range("A9")
    mycell.ClearContents
    Set mycell = sheetKopija.Range("A20")
    mycell.ClearContents
            'PODVOJIS SHEET KOPIJA
    n = Sheets.Count
    sheetKopija.Copy after:=Sheets(n)
    
          'Clear clipboard
    Application.CutCopyMode = False
    
        'PREIMENUJES PODVOJEN (ZADNJI) SHEET
    lastSheetIndex = Sheets.Count
    Sheets(lastSheetIndex).Name = sheetKopija.Range("B34").Value
    
End Sub

Function IsValidDateRange(dateRange As String) As Boolean
    Dim dateArray() As String
    Dim fromDate As Date
    Dim toDate As Date
    
    ' Split the string by hyphen
    dateArray = Split(dateRange, "-")
    
    ' Check if there are two elements after splitting
    If UBound(dateArray) = 1 Then
        ' Attempt to convert the parts into dates with the specific format
        If IsDate(Replace(dateArray(0), ".", "/")) And IsDate(Replace(dateArray(1), ".", "/")) Then
            fromDate = CDate(Replace(dateArray(0), ".", "/"))
            toDate = CDate(Replace(dateArray(1), ".", "/"))
            
            ' Check if the first date is before or equal to the second date
            If fromDate < toDate Then
                IsValidDateRange = True
            End If
        End If
    End If
End Function

Function initializeVariables()
    Set sheetDela = Sheets("Database_dela")
    Set sheetStranke = Sheets("Database_stranke")
    Set sheetRacuni = Sheets("Racuni")
    Set sheetEvidence = Sheets("Evidence")
    Set sheetKopija = Sheets("KOPIJA")
End Function

Function toDate(myDate As String) As Date
    toDate = CDate(Replace(myDate, ".", "/"))
End Function
'Private Sub btn_add_Click()
'    Dim datum As Date
'
'    'On Error Resume Next
'    datum = CDate(Replace(box_datumizdaje.Value, ".", "/"))
'
'If Err.Number <> 0 Then
'    MsgBox "Error Number: " & Err.Number & vbCrLf & "Error Description: " & Err.Description, vbExclamation, "Error"
'    Err.Clear ' Clear the error object
'Else
'    box_rokplacila.Value = DateAdd("d", box_valutadni.Value, datum)
'End If
'On Error GoTo 0 ' Reset error handling
'End Sub

Private Sub btn_add_Click()

firstEmptyRowNum = sheetEvidence.Cells(sheetEvidence.Rows.Count, "A").End(xlUp).Row + 2

MsgBox firstEmptyRowNum
End Sub
