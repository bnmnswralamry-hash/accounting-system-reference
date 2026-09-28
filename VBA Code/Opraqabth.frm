VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Opraqabth 
   ClientHeight    =   10644
   ClientLeft      =   -336
   ClientTop       =   -1308
   ClientWidth     =   17136
   OleObjectBlob   =   "Opraqabth.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Opraqabth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub cbokop_AfterUpdate()
    
    Dim selectedVal As String
    selectedVal = Trim(Me.CboKOP.Value)
    
    If selectedVal = "" Then Exit Sub
    
    If Not IsInList(Me.CboKOP, selectedVal) Then
        MsgBox "«·‰Ê⁄ «·–Ì √œŒ· Â €Ì— „ÊÃÊœ", vbExclamation, " ‰»ÌÂ"
        Me.CboKOP.Value = ""
        Me.CboKOP.SetFocus
    End If

End Sub

Private Sub Cbokop_Change()

    On Error Resume Next
    Me.CboKOP.DropDown
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim accType As String
    Dim jFilter As String
    
    Set ws = ThisWorkbook.Sheets("judje")
    Me.Co_cashac.clear
    
    If Me.CboKOP.ListIndex = -1 Then Exit Sub
    
    accType = Trim(Me.CboKOP.Value)
    
    '  ÕœÌœ ﬁÌ„… «·⁄„Êœ J Õ”» «Œ Ì«— cbokop
    Select Case UCase(accType)
        Case "”‰œ ﬁ»÷ ‰ﬁœÌ"
            jFilter = "’‰œÊﬁ"
        Case "”‰œ ﬁ»÷ »‰ﬂÌ"
            jFilter = "»‰ﬂ"
        Case "ﬁ»÷ ⁄Âœ…"
            jFilter = "⁄Âœ"
        Case Else
            Exit Sub
    End Select
    
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    
    With Me.Co_cashac
        For i = 3 To lastRow
            ' ‘—ÿ 1: ⁄„Êœ J = «·ﬁÌ„… «·„Õœœ… „‰ cbokop
            ' ‘—ÿ 2: ⁄„Êœ F = " Õ·Ì·Ì"
            If UCase(Trim(ws.Cells(i, "J").Value)) = UCase(jFilter) _
               And UCase(Trim(ws.Cells(i, "F").Value)) = " Õ·Ì·Ì" Then
                
                If Trim(ws.Cells(i, "C").Value) <> "" Then
                    .AddItem ws.Cells(i, "C").Value
                End If
            End If
        Next i
    End With
    
    ' «Œ Ì«—Ì: «Œ Ì«— √Ê· ⁄‰’—  ·ﬁ«∆Ì«
   ' If Me.Co_cashac.listCount > 0 Then
   '     Me.Co_cashac.ListIndex = 0
   ' End If


    
   
  
    Dim maxNum As Long
    Dim num As Long
    Dim prefix As String
    
    On Error GoTo ErrHandler
    
    
    If Trim(Me.CboKOP.Text) = "" Then
        Me.Tex_NU.Value = ""
        Exit Sub
    End If
    
    Set ws = ThisWorkbook.Worksheets("Daily Operation")
    maxNum = 0
    
    
    Select Case Trim(Me.CboKOP.Text)
        Case "”‰œ ﬁ»÷ ‰ﬁœÌ", "ﬁ»÷ ⁄Âœ…", "”‰œ ﬁ»÷ »‰ﬂÌ"
            prefix = "RC"
       
    End Select
    
    
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    If lastRow < 3 Then lastRow = 3
    
    
    For i = 3 To lastRow
        If UCase(Left(Trim(ws.Cells(i, "B").Value), Len(prefix))) = prefix Then
            num = val(Mid(Trim(ws.Cells(i, "B").Value), Len(prefix) + 1))
            If num > maxNum Then maxNum = num
        End If
    Next i
    
    
    Me.Tex_NU.Value = prefix & Format(maxNum + 1, "000")
    Exit Sub
    
ErrHandler:
    MsgBox "Œÿ√ ›Ì  Ê·Ìœ —ﬁ„ «·⁄„·Ì…: " & Err.Description, vbExclamation
    Me.Tex_NU.Value = prefix & "001"

End Sub

Private Sub CboMostafed_Change()
Me.CboMostafed.DropDown
End Sub

Private Sub cbosenter_Change()
Me.cbosenter.DropDown

End Sub

'Private Sub co_cashac_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
Private Sub Co_cashac_AfterUpdate()
    Dim selectedAcc As String
    selectedAcc = Trim(Me.Co_cashac.Value)
    
    If selectedAcc = "" Then Exit Sub
    
    ' ›Õ’ ·Ê «Œ «— ‰›” «·Õ”«» «··Ì ›Ì «·„œÌ‰
    If UCase(Trim(selectedAcc)) = UCase(Trim(Me.Co_dpacc.Value)) Then
        MsgBox "·« Ì„ﬂ‰ «·ﬁ»÷ „‰ ‰›” «·Õ”«»", vbExclamation, " ‰»ÌÂ"
        Me.Co_cashac.Value = ""
        Me.Co_cashac.SetFocus
        Exit Sub
    End If
    
    ' »«ﬁÌ ﬂÊœ ›Õ’ ÊÃÊœ «·Õ”«»
    If Not IsInList(Me.Co_cashac, selectedAcc) Then
        MsgBox "«·Õ”«» «·„œŒ· €Ì— „ÊÃÊœ", vbExclamation, " ‰»ÌÂ"
        Me.Co_cashac.Value = ""
        Me.Co_cashac.SetFocus
    End If
   '
End Sub

Private Sub Co_cashac_Change()
Me.Co_cashac.DropDown

End Sub
Private Sub co_cur_Change()
Dim ws As Worksheet
    Dim rng As Range
    Dim MINP As Double
    Dim MAXP As Double
    
    Set ws = ThisWorkbook.Worksheets("data")
    Set rng = ws.Range("c:c").Find(Co_cur.Value)
    
    If Not rng Is Nothing Then
      MINP = ws.Cells(rng.Row, "f").Value
       MAXP = ws.Cells(rng.Row, "g").Value
       
       Tex_rate.Value = ""
        Tex_rate.Tag = MINP & "I" & MAXP
       Else
        Tex_rate.Value = ""
        Tex_rate.Tag = ""
       End If
       
       
      '  ÕœÌœ ”⁄— «·’—›
   
    Dim curName As String
    Dim foundRow As Long
    
    If Me.Co_cur.ListIndex = -1 Then Exit Sub
    
    Set ws = ThisWorkbook.Worksheets("data")
    curName = Trim(Me.Co_cur.Text)
    
    ' œÊ— ⁄·Ï «·⁄„·… ›Ì ⁄„Êœ A
    foundRow = 0
    On Error Resume Next
    foundRow = Application.Match(curName, ws.Range("c3:c" & ws.Cells(ws.Rows.Count, "c").End(xlUp).Row), 0)
    On Error GoTo 0
    
    If foundRow > 0 Then
        foundRow = foundRow + 2
        
        
        If LCase(Trim(ws.Cells(foundRow, "H").Value)) = "⁄„·… „Õ·Ì…" Then
            Me.Tex_rate.Text = "1"
            Me.Tex_rate.Enabled = False
            Me.Tex_Amount.SetFocus
        Else
            
            Me.Tex_rate.Text = ws.Cells(foundRow, "E").Value
            Me.Tex_rate.Enabled = True
            Me.Tex_rate.SetFocus
        End If
    Else
        Me.Tex_rate.Text = ""
        Me.Tex_rate.Enabled = True
    End If

End Sub
Private Sub cobkindmrga_AfterUpdate()
Dim selectedVal As String
    selectedVal = Trim(Me.cobkindmrga.Value)
    
    If selectedVal = "" Then Exit Sub
    
    If Not IsInList(Me.cobkindmrga, selectedVal) Then
        MsgBox "«·‰Ê⁄ «·–Ì √œŒ· Â €Ì— „ÊÃÊœ", vbExclamation, " ‰»ÌÂ"
        Me.cobkindmrga.Value = ""
        Me.cobkindmrga.SetFocus
    End If
End Sub

Private Sub cobkindmrga_Change()

    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim dict As Object
    Dim key As Variant
    Dim filterVal As String
    
    Set ws = ThisWorkbook.Sheets("daily operation")
    Set dict = CreateObject("Scripting.Dictionary")
    
    Me.cobnamemrga.clear
    
    Select Case Me.cobkindmrga.Value
        Case "«” ·«„ ‰ﬁœÌ"
           ' Me.LablNameMrga.Visible = False
           ' Me.cobnamemrga.Visible = False
            Me.cobnamemrga.Value = "”‰œ «” ·«„"
            Me.cobnamemrga.Enabled = False
           ' Me.Tex_MrgaNo.SetFocus
        Case "ÕÊ«·Â"
            Me.LablNameMrga.Visible = True
            Me.LablNameMrga.Caption = "«”‹‹„ «·»‰ﬂ"
            Me.cobnamemrga.Visible = True
            Me.cobnamemrga.Enabled = True
            filterVal = "ÕÊ«·Â"
            Call FillFilteredList(filterVal)
            
             
            
        Case Else
            Me.LablNameMrga.Visible = False
            Me.cobnamemrga.Visible = False
    End Select
End Sub

Private Sub FillFilteredList(filterVal As String)
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim dict As Object
    Dim key As Variant
    
    Set ws = ThisWorkbook.Sheets("daily operation")
    Set dict = CreateObject("Scripting.Dictionary")
    
    lastRow = ws.Cells(ws.Rows.Count, "O").End(xlUp).Row
    
    For i = 3 To lastRow
        If UCase(Trim(ws.Cells(i, "O").Value)) = UCase(filterVal) Then
            If Trim(ws.Cells(i, "P").Value) <> "" Then
                If Not dict.Exists(Trim(ws.Cells(i, "P").Value)) Then
                    dict.Add Trim(ws.Cells(i, "P").Value), ""
                End If
            End If
        End If
    Next i
    
    Me.cobnamemrga.clear
    For Each key In dict.Keys
        Me.cobnamemrga.AddItem CStr(key)
    Next key
        Me.cobnamemrga.Value = ""
        Me.cobnamemrga.ListIndex = -1
        
End Sub



Private Sub cobnamemrga_Change()
If Me.Enabled = True Then
If Len(Me.cobnamemrga.Text) > 0 Then
Me.cobnamemrga.DropDown
End If
End If

End Sub


Private Sub save_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
If Tex_da.Value = "" Or Tex_NU.Value = "" Or Co_cashac.Value = "" Or cboBian.Value = "" Or Tex_Amount.Value = "" Or Tex_rate.Value = "" Or cbosenter.Value = "" Or Co_cur.Value = "" Or Co_dpacc.Value = "" Then

MsgBox prompt:="«·—Ã«¡ «· √ﬂœ „‰ ≈œŒ«· Ã„Ì⁄ »Ì«‰«  «·≈–‰", Title:="Œÿ√⁄œ„ «ﬂ „«· «·»Ì«‰« "
Else
Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual

Me.ListBox1.RowSource = ""
''''''''''''«÷«›…''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sheets("Daily Operation").Activate
 With ThisWorkbook.Sheets("Daily Operation")
    .Range("A3:Z3").Select
    Selection.ListObject.ListRows.Add (1)

    .Range("A3").Value = "=row()-2"
    .Range("B3").Value = Me.Tex_NU.Value
    .Range("C3").Value = Me.Tex_da
    .Range("C3").NumberFormat = "yyyy-mm-dd"
    .Range("I3").Value = Me.Co_cashac.Text
    .Range("D3").Value = Me.CboKOP.Value
    .Range("g3").Value = Me.cboBian.Value
    .Range("F3").Value = Me.Co_dpacc.Value
    .Range("J3").Value = Me.Co_cur.Value
    .Range("K3").Value = Me.Tex_Amount.Text
    .Range("L3").Value = Me.Tex_rate.Value
    .Range("n3").Value = Me.CboMostafed.Value
    .Range("O3").Value = Me.cobkindmrga.Value
    .Range("P3").Value = Me.cobnamemrga.Text
    .Range("Q3").Value = Me.Tex_MrgaNo.Text
    .Range("R3").Value = Me.cbosenter.Text
    .Range("U3").Value = Format(Now(), "yyyy-mm-dd hh:mm:ss")
    
        End With
        
       
    Dim ws As Worksheet
    Dim lastRow As Long
    
    Set ws = ThisWorkbook.Worksheets("Daily Journal")
    
   
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    If lastRow < 3 Then lastRow = 2
    
   
    ' ????? ?????: Tex_NU + co_dpacc
    ws.Cells(lastRow + 1, "C").Value = Me.Tex_NU.Value
    ws.Cells(lastRow + 1, "G").Value = Me.Co_cashac.Value
    
    ' ????? ??????: Tex_NU + co_cashac
    ws.Cells(lastRow + 2, "C").Value = Me.Tex_NU.Value
    ws.Cells(lastRow + 2, "G").Value = Me.Co_dpacc.Value
    
    
          Application.Calculation = xlCalculationAutomatic
          Application.EnableEvents = True
          Application.ScreenUpdating = True
   
        
     MsgBox (" „ «÷«›… «·≈–‰ »‰Ã«Õ")
   
   Me.Co_cashac.Value = ""
   Me.CboKOP.Value = ""
   Me.cboBian.Value = ""
   Me.Co_dpacc.Value = ""
   Me.Co_cur.Value = ""
   Me.Tex_Amount.Text = ""
   Me.Tex_rate.Text = ""
   Me.CboMostafed.Value = ""
   Me.cobkindmrga.Value = ""
   Me.cobnamemrga.Value = ""
   Me.Tex_MrgaNo.Value = ""
   Me.cbosenter.Value = ""
  
      End If
      
       Me.ListBox1.RowSource = "Tabledaily"

    
      
End Sub

Private Sub Tex_Amount_Change()
  Call Calc_Amount
End Sub

Private Sub Tex_da_AfterUpdate()
'  ﬂ”  «· «—ÌŒ
 Dim dateValue As Date
    Dim dateString As String
    
    dateString = Tex_da.Value
   
    If IsNumeric(dateString) Then
        dateString = Format(dateString, "00000000")
        dateString = Left(dateString, 4) & "-" & Mid(dateString, 5, 2) & "-" & Right(dateString, 2)
    End If

    On Error Resume Next
    dateValue = CDate(dateString)
    On Error GoTo 0
    
    If dateValue = 0 Then
        MsgBox " «—ÌŒ €Ì— ’ÕÌÕ", vbExclamation, "Œÿ√"
        Tex_da.Value = Format(Date, "yyyy-mm-dd")
        Exit Sub
    End If
 
    If dateValue > Date Then
        MsgBox "«· «—ÌŒ «ﬂ»— „‰  «—ÌŒ «·ÌÊ„", vbInformation, " ‰»ÌÂ"
    End If
     
    If Month(dateValue) < 1 Or Month(dateValue) > 12 Then
        MsgBox "«·‘Â— €Ì— ’ÕÌÕ", vbExclamation, "Œÿ√"
        Tex_da.Value = Format(Date, "yyyy-mm-dd")
        Exit Sub
    End If
    
    
    If Day(dateValue) < 1 Or Day(dateValue) > Day(DateSerial(Year(dateValue), Month(dateValue) + 1, 0)) Then
        MsgBox "«·ÌÊ„ €Ì— ’ÕÌÕ", vbExclamation, "Œÿ«¡"
        Tex_da.Value = Format(Date, "yyyy-mm-dd")
        Exit Sub
    End If
    ' date
    Tex_da.Value = Format(dateValue, "yyyy-mm-dd")
      
   
    
End Sub
Private Sub Tex_rate_AfterUpdate()
 
    Call Calc_Amount
  Dim PRICE() As String
    Dim MINP As Double
    Dim MAXP As Double
    
 PRICE = Split(Tex_rate.Tag, "I")
 
  If UBound(PRICE) = 1 Then
  MINP = PRICE(0)
  MAXP = PRICE(1)
  
  If Tex_rate.Value = "" Then
     MsgBox "«·—Ã«¡ «œŒ«· ”⁄— «·’—›", vbExclamation, "Œÿ√"
    ElseIf Not IsNumeric(Tex_rate.Value) Then
    MsgBox "«·—Ã«¡ «œŒ«· ﬁÌ„… —ﬁ„Ì…", vbExclamation, "Œÿ«¡"
    Tex_rate.Value = ""
    ElseIf CDbl(Tex_rate.Value) < MINP Or CDbl(Tex_rate.Value) > MAXP Then
    MsgBox "”⁄— «·’—› ÌÃ» «‰ ÌﬂÊ‰ »Ì‰" & MINP & "Ê" & MAXP, vbExclamation, "Œÿ√"
    Tex_rate.Value = ""
    Tex_rate.SetFocus
    End If
    End If
    
End Sub


Private Sub UserForm_Initialize()

 Tex_da.Value = Format(Date, "yyyy-mm-dd")
 
  ' full cbokop
    With Me.CboKOP
    .clear
    .AddItem "”‰œ ﬁ»÷ ‰ﬁœÌ"
    .AddItem "”‰œ ﬁ»÷ »‰ﬂÌ"
     .AddItem "ﬁ»÷ ⁄Âœ…"
    End With
    
    With Me.cobkindmrga
        .clear
        .AddItem "«” ·«„ ‰ﬁœÌ"
        .AddItem "ÕÊ«·Â"
        
        
    End With
   
        
    'cbmostafi
   

    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim dict As Object
    Dim key As Variant
    
    Set ws = ThisWorkbook.Sheets("daily operation")
    Set dict = CreateObject("Scripting.Dictionary")
    
    Me.CboMostafed.clear
    
    lastRow = ws.Cells(ws.Rows.Count, "N").End(xlUp).Row
    
    '
    For i = 3 To lastRow
        If Trim(ws.Cells(i, "N").Value) <> "" Then
            If Not dict.Exists(Trim(ws.Cells(i, "N").Value)) Then
                dict.Add Trim(ws.Cells(i, "N").Value), ""
            End If
        End If
    Next i
    
    
    For Each key In dict.Keys
        Me.CboMostafed.AddItem CStr(key)
    Next key
    '''''''''''''''''''
    
    

'''''''''''''''''''''''''''''

    Me.cboBian.Style = fmStyleDropDownCombo
    Me.cboBian.MatchEntry = fmMatchEntryNone
    Call FillCbobianFull
    
    
    
    
   
    Me.Co_dpacc.Style = fmStyleDropDownCombo
    Me.Co_dpacc.MatchEntry = fmMatchEntryNone
    Call FillCoDpAccFull
End Sub

Private Sub co_dpacc_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long, j As Long
    Dim searchTxt As String, itemTxt As String, typedTxt As String
    Dim tempList() As String, words() As String
    Dim found As Boolean, listCount As Long

    On Error GoTo CleanExit
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyEscape Then Exit Sub

    Set ws = ThisWorkbook.Worksheets("judje")
    typedTxt = Me.Co_dpacc.Text
    searchTxt = Trim(typedTxt)

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    If searchTxt = "" Then
        Call FillCoDpAccFull
        GoTo CleanExit
    End If

    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    ReDim tempList(0 To 0)
    listCount = 0

    For i = 3 To lastRow
        If ws.Cells(i, "F").Value = " Õ·Ì·Ì" Then
            itemTxt = ws.Cells(i, "C").Value
            words = Split(itemTxt, " ")
            found = False
            
            For j = LBound(words) To UBound(words)
                If InStr(1, UCase(words(j)), UCase(searchTxt), vbTextCompare) > 0 Then
                    found = True
                    Exit For
                End If
            Next j
            
            If found Then
                tempList(listCount) = itemTxt
                listCount = listCount + 1
                ReDim Preserve tempList(0 To listCount)
            End If
        End If
    Next i

    If listCount > 0 Then
        ReDim Preserve tempList(0 To listCount - 1)
        Me.Co_dpacc.List = tempList
        Me.Co_dpacc.DropDown
    Else
        Me.Co_dpacc.clear
    End If

    Me.Co_dpacc.Text = typedTxt
    Me.Co_dpacc.SelStart = Len(typedTxt)

CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    On Error GoTo 0
End Sub

Private Sub FillCoDpAccFull()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim tempList() As String
    Dim listCount As Long

    Set ws = ThisWorkbook.Worksheets("judje")
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    ReDim tempList(0 To 0)
    listCount = 0

    For i = 3 To lastRow
        If ws.Cells(i, "F").Value = " Õ·Ì·Ì" Then
            tempList(listCount) = ws.Cells(i, "C").Value
            listCount = listCount + 1
            ReDim Preserve tempList(0 To listCount)
        End If
    Next i

    If listCount > 0 Then
        ReDim Preserve tempList(0 To listCount - 1)
        Me.Co_dpacc.List = tempList
    End If
End Sub

Private Sub co_dpacc_AfterUpdate()
    Dim selectedAcc As String
    selectedAcc = Trim(Me.Co_dpacc.Value)
    
    If selectedAcc = "" Then Exit Sub
    
    ' ·« Ì„ﬂ‰ «·’—› „‰ ‰›” «·Õ”«»
    If selectedAcc = Trim(Me.Co_cashac.Value) Then
        MsgBox "·« Ì„ﬂ‰ «·’—› „‰ ‰›” «·Õ”«»", vbExclamation
        Me.Co_dpacc.Value = ""
        Me.Co_dpacc.SetFocus
        Exit Sub
    End If
    
    ' «·Õ”«» €Ì— „ÊÃÊœ
    If Not IsInList(Me.Co_dpacc, selectedAcc) Then
        MsgBox "«·Õ”«» «·„œŒ· €Ì— „ÊÃÊœ", vbExclamation
        Me.Co_dpacc.Value = ""
        Me.Co_dpacc.SetFocus
    End If
End Sub

Private Function IsInList(cmb As MSForms.ComboBox, val As String) As Boolean
    Dim i As Long
    For i = 0 To cmb.listCount - 1
        If Trim(cmb.List(i)) = val Then
            IsInList = True
            Exit Function
        End If
    Next i
    IsInList = False

    
End Function

''''''''''''''''''''
Private Sub Cbobian_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim searchTxt As String, itemTxt As String, typedTxt As String
    Dim tempList() As String
    Dim dict As Object
    Dim listCount As Long
    Dim searchWords() As String
    
    On Error GoTo CleanExit
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyEscape Then Exit Sub
    
    Set ws = ThisWorkbook.Worksheets("daily operation")
    typedTxt = Me.cboBian.Text
    searchTxt = Trim(typedTxt)
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    If searchTxt = "" Then
        Call FillCbobianFull
        GoTo CleanExit
    End If
    
    Set dict = CreateObject("Scripting.Dictionary")
    searchWords = Split(UCase(searchTxt), " ")
    lastRow = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
    ReDim tempList(0 To 0)
    listCount = 0
    
    For i = 3 To lastRow
        itemTxt = Trim(ws.Cells(i, "G").Value)
        If itemTxt <> "" Then
            If Not dict.Exists(itemTxt) Then
                If MatchExactStart(UCase(itemTxt), searchWords) Then
                    dict.Add itemTxt, ""
                    tempList(listCount) = itemTxt
                    listCount = listCount + 1
                    ReDim Preserve tempList(0 To listCount)
                End If
            End If
        End If
    Next i
    
    If listCount > 0 Then
        ReDim Preserve tempList(0 To listCount - 1)
        Me.cboBian.List = tempList
        Me.cboBian.DropDown
    Else
        Me.cboBian.clear
    End If
    
    Me.cboBian.Text = typedTxt
    Me.cboBian.SelStart = Len(typedTxt)
    
CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    On Error GoTo 0
End Sub

Private Sub FillCbobianFull()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim tempList() As String
    Dim dict As Object
    Dim listCount As Long
    
    Set ws = ThisWorkbook.Worksheets("daily operation")
    Set dict = CreateObject("Scripting.Dictionary")
    
    lastRow = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
    ReDim tempList(0 To 0)
    listCount = 0
    
    For i = 3 To lastRow
        If Trim(ws.Cells(i, "G").Value) <> "" Then
            If Not dict.Exists(Trim(ws.Cells(i, "G").Value)) Then
                dict.Add Trim(ws.Cells(i, "G").Value), ""
                tempList(listCount) = Trim(ws.Cells(i, "G").Value)
                listCount = listCount + 1
                ReDim Preserve tempList(0 To listCount)
            End If
        End If
    Next i
    
    If listCount > 0 Then
        ReDim Preserve tempList(0 To listCount - 1)
        Me.cboBian.List = tempList
    End If
End Sub

Private Function MatchExactStart(itemText As String, searchWords() As String) As Boolean
    Dim itemWords() As String
    Dim i As Long, j As Long
    
    itemWords = Split(itemText, " ")
    j = LBound(searchWords)
    
    For i = LBound(itemWords) To UBound(itemWords)
        If Len(itemWords(i)) >= Len(searchWords(j)) Then
            If Left(itemWords(i), Len(searchWords(j))) = searchWords(j) Then
                j = j + 1
                If j > UBound(searchWords) Then
                    MatchExactStart = True
                    Exit Function
                End If
            End If
        End If
    Next i
    
    MatchExactStart = False
End Function

Private Sub Calc_Amount()
    If Trim(Me.Tex_rate.Text) = "" Or Trim(Me.Tex_Amount.Text) = "" Then
        Me.TXTAMOSR.Text = ""
    Else
        Me.TXTAMOSR.Text = Format(val(Me.Tex_Amount.Text) / val(Me.Tex_rate.Text), "0.00")
    End If
End Sub
