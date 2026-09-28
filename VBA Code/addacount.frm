VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} addacount 
   ClientHeight    =   11364
   ClientLeft      =   -336
   ClientTop       =   -1308
   ClientWidth     =   17124
   OleObjectBlob   =   "addacount.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "addacount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

 Private isUpdating As Boolean

Private Sub AddNew_Click()
 
    Dim wsJudje As Worksheet
    Dim lastRow As Long, i As Long
    Dim chAccValue As String
    Dim targetLevel As Integer
    
    Set wsJudje = ThisWorkbook.Sheets("judje")
    chAccValue = Trim(Me.Cobchacc.Value)
    
    If chAccValue = "" Then Exit Sub
    
    Application.EnableEvents = False
    isUpdating = True
    
    ' „”Õ «·ﬁÌ„ «·ﬁœÌ„… »”
    Me.cboparent.clear
    Me.txtAccountName.Text = ""
    Me.cboGroupCode.Value = ""
    Me.cboAnalytical.Value = ""
    Me.cboClasscode.Value = ""
    Me.txtLevel.Value = ""
    Me.cboType.Value = ""
    
    Select Case chAccValue
        Case "⁄«„"
            Me.cboType.Value = "⁄«„"
            Me.txtLevel.Value = 2
            targetLevel = 1
            Me.cboparent.Enabled = True
            Me.txtAccountName.Enabled = True
            Me.cboGroupCode.Enabled = True
            Me.Cobchacc.Enabled = False
            Me.cboparent.SetFocus
        Case "„”«⁄œ"
            Me.cboType.Value = "„”«⁄œ"
            Me.txtLevel.Value = 3
            targetLevel = 2
             Me.Cobchacc.Enabled = False
            Me.cboparent.Enabled = True
            Me.txtAccountName.Enabled = True
            Me.cboGroupCode.Enabled = True
            Me.cboparent.SetFocus
            
        Case "›—⁄Ì"
            Me.cboType.Value = "›—⁄Ì"
            Me.txtLevel.Value = 4
            targetLevel = 3
            Me.cboparent.Enabled = True
            Me.txtAccountName.Enabled = True
            Me.cboGroupCode.Enabled = True
             Me.Cobchacc.Enabled = False
             Me.cboparent.SetFocus
        Case " Õ·Ì·Ì"
            Me.cboType.Value = " Õ·Ì·Ì"
            Me.txtLevel.Value = 5
            targetLevel = 4
            Me.cboparent.Enabled = True
            Me.txtAccountName.Enabled = True
            Me.cboGroupCode.Enabled = True
            Me.cboAnalytical.Enabled = True
            Me.cboClasscode.Enabled = True
            Me.Cobchacc.Enabled = False
            Me.cboparent.SetFocus
    End Select
    
    '  ⁄»∆… cboparent „‰ judje ⁄„Êœ C Õ”» „” ÊÏ E
    If targetLevel > 0 Then
        lastRow = wsJudje.Cells(wsJudje.Rows.Count, "C").End(xlUp).Row
        For i = 3 To lastRow
            If val(wsJudje.Cells(i, "E").Value) = targetLevel Then
                If Trim(wsJudje.Cells(i, "C").Value) <> "" Then
                    Me.cboparent.AddItem wsJudje.Cells(i, "C").Value
                End If
            End If
        Next i
    End If
    
    Application.EnableEvents = True
    isUpdating = False
    Me.cboparent.SetFocus

End Sub

' œ«·… „”«⁄œ… Cboparent›Ì  ⁄»∆…
Private Sub FillCboparentByLevel(ws As Worksheet, level As Integer)
    Dim lastRow As Long, i As Long
    
    cboparent.clear
    lastRow = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
    
    For i = 3 To lastRow
        If val(ws.Cells(i, "E").Value) = level Then
            cboparent.AddItem ws.Cells(i, "C").Value
        End If
    Next i

End Sub
 Private Sub cboparent_Change()
' ⁄»∆…  ﬂ”  —ﬁ„ «·Õ”«» «·Õ”«» «·—∆Ì”Ì Ê  ﬂ”  — »… «·Õ”«»
 

    Dim wsJudje As Worksheet
    Dim parentName As String
   
    Dim i As Long, lastRow As Long
    
    If isUpdating Then Exit Sub
    If Trim(Me.cboparent.Value) = "" Then Exit Sub
    
    Set wsJudje = ThisWorkbook.Sheets("judje")
    parentName = Trim(Me.cboparent.Value)
   
    Application.EnableEvents = False
    isUpdating = True
    
    Me.TxtprenterNo.Value = ""
    lastRow = wsJudje.Cells(wsJudje.Rows.Count, "C").End(xlUp).Row
    
    For i = 3 To lastRow
        If Trim(wsJudje.Cells(i, "C").Value) = parentName Then
            Me.TxtprenterNo.Value = wsJudje.Cells(i, "B").Value
            Exit For
        End If
    Next i
    
    Application.EnableEvents = True
    isUpdating = False

        
    ''''''''''''''''''''''' —ﬁÌ„ «·Õ”«» «·ÃœÌœ
  
    
    
   
      
End Sub

Private Sub clear_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Me.cboparent.Text = ""
    Me.txtAccountNo.Text = ""
    Me.txtAccountName.Text = ""
    Me.cboClasscode.Text = ""
    Me.cboType.Text = ""
    Me.cboAnalytical.Text = ""
    Me.txtLevel.Text = ""
    Me.cboNature.Text = ""
    Me.cboReportType.Text = ""
    Me.chckStatus.Value = False
    Me.cboGroupCode.Value = ""
    Me.Cobchacc.SetFocus
    Me.save.Visible = True
    Me.edit.Visible = False
    Me.del.Visible = False
      Me.cboparent.Enabled = False
            Me.txtAccountName.Enabled = False
            Me.cboGroupCode.Enabled = False
            Me.cboAnalytical.Enabled = False
            Me.cboClasscode.Enabled = False
            Me.Cobchacc.Enabled = True
    
End Sub
Private Sub save_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
   
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim accName As String
     Dim accType As String
    Set ws = ThisWorkbook.Worksheets("judje")
    accName = Trim(Me.txtAccountName.Text)
    
  
accType = LCase(Trim(Me.cboType.Text))

' 1. «· Õﬁﬁ „‰ «”„ «·Õ”«» - œ«Ì„«
If accName = "" Then
    MsgBox "ÌÃ» «œŒ«· «”„ «·Õ”«»", vbExclamation
    Me.txtAccountName.SetFocus: Exit Sub
End If

' 2. «· Õﬁﬁ «·√”«”Ì ·ﬂ· «·√‰Ê«⁄
If Trim(Me.cboparent.Text) = "" Then
    MsgBox "Ì—ÃÏ «Œ Ì«— «·Õ”«» «·—∆Ì”Ì", vbExclamation
    Me.cboparent.SetFocus: Exit Sub
End If
If Trim(Me.txtAccountNo.Text) = "" Then
    MsgBox "ÌÃ» «œŒ«· —ﬁ„ «·Õ”«»", vbExclamation
    Me.txtAccountNo.SetFocus: Exit Sub
End If
If Trim(Me.TxtprenterNo.Text) = "" Then
    MsgBox "ÌÃ» «œŒ«· —ﬁ„ «·Õ”«» «·—∆Ì”Ì", vbExclamation
    Me.TxtprenterNo.SetFocus: Exit Sub
End If
If Trim(Me.cboType.Text) = "" Then
    MsgBox "ÌÃ»  ÕœÌœ ‰Ê⁄ «·Õ”«»", vbExclamation
    Me.cboType.SetFocus: Exit Sub
End If
If Trim(Me.cboNature.Text) = "" Then
    MsgBox "ÌÃ»  ÕœÌœ ÿ»Ì⁄… «·Õ”«»", vbExclamation
    Me.cboNature.SetFocus: Exit Sub
End If
If Trim(Me.cboReportType.Text) = "" Then
    MsgBox "ÌÃ» «Œ Ì«— ‰Ê⁄ «· ﬁ—Ì—", vbExclamation
    Me.cboReportType.SetFocus: Exit Sub
End If

' 3. ‘—ÿ ≈÷«›Ì »” ·Ê «·‰Ê⁄  Õ·Ì
If accType = " Õ·Ì·Ì" Then
    If Trim(Me.cboAnalytical.Text) = "" Then
        MsgBox "ÌÃ» «Œ Ì«— «·»‰œ «· Õ·Ì·Ì", vbExclamation
        Me.cboAnalytical.SetFocus: Exit Sub
    End If
End If
    Me.ListBox1.RowSource = ""
    ' 2. «· Õﬁﬁ „‰  ﬂ—«— «”„ «·Õ”«» ›Ì ⁄„Êœ C
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    For i = 2 To lastRow
        If LCase(Trim(ws.Cells(i, "C").Value)) = LCase(accName) Then
            MsgBox "Â–« «·Õ”«» „ÊÃÊœ „”»ﬁ”«", vbExclamation
            Me.txtAccountName.SetFocus
            GoTo CleanExit
            Exit Sub
        End If
    Next i
    
    ' 3. «÷«›… «·»Ì«‰«  ›Ì «Œ— ”ÿ— ›«÷Ì
    lastRow = lastRow + 1
    
    ws.Cells(lastRow, "D").Value = Me.TxtprenterNo.Text        ' «·⁄„Êœ D
    ws.Cells(lastRow, "B").Value = Me.txtAccountNo.Text       ' «·⁄„Êœ B
    ws.Cells(lastRow, "C").Value = accName                    ' «·⁄„Êœ C
    ws.Cells(lastRow, "L").Value = Me.cboClasscode.Text       ' «·⁄„Êœ L
    ws.Cells(lastRow, "F").Value = Me.cboType.Text            ' «·⁄„Êœ F
    ws.Cells(lastRow, "H").Value = Me.cboGroupCode.Text       ' «·⁄„Êœ H
    ws.Cells(lastRow, "J").Value = Me.cboAnalytical.Text      ' «·⁄„Êœ J
    ws.Cells(lastRow, "E").Value = Me.txtLevel.Text           ' «·⁄„Êœ E
    ws.Cells(lastRow, "I").Value = Me.cboNature.Text          ' «·⁄„Êœ I
    ws.Cells(lastRow, "G").Value = Me.cboReportType.Text      ' «·⁄„Êœ G
    
    ' «·⁄„Êœ K Õ”» Õ«·… «·‘Ìﬂ »Êﬂ”
    If Me.chckStatus.Value = True Then
        ws.Cells(lastRow, "K").Value = "„Êﬁ›"
    Else
        ws.Cells(lastRow, "K").Value = "‰‘ÿ"
    End If
    Me.ListBox1.RowSource = "TableJudje"
    
    MsgBox " „ «÷«›… «·Õ”«» »‰Ã«Õ", vbInformation
    
    ' 4. „”Õ «·ÕﬁÊ· »⁄œ «·Õ›Ÿ
    Me.cboparent.Text = ""
    Me.txtAccountNo.Text = ""
    Me.txtAccountName.Text = ""
    Me.cboClasscode.Text = ""
    Me.cboType.Text = ""
    Me.cboAnalytical.Text = ""
    Me.txtLevel.Text = ""
    Me.cboNature.Text = ""
    Me.cboReportType.Text = ""
    Me.chckStatus.Value = False
    Me.cboGroupCode.Value = ""
     Me.cboparent.Enabled = False
     Me.cboClasscode.Enabled = False
     Me.cboAnalytical.Enabled = False
     Me.cboGroupCode.Enabled = False
     Me.txtAccountName.Enabled = False
     Me.Cobchacc.Enabled = True
     Me.Cobchacc.SetFocus
CleanExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub
    
ErrHandler:
    MsgBox "ÕœÀ Œÿ√" & Err.Description & vbCrLf & "—ﬁ„ «·Œÿ√" & Err.number, vbCritical
    Resume CleanExit
    
End Sub



Private Sub txtprenterNo_Change()
 

     Dim wsJudje As Worksheet
    Dim lastRow As Long, i As Long
    Dim cnt As Long
    Dim mainAcc As String
    Dim accType As String
    
    If Trim(Me.TxtprenterNo.Value) = "" Then Exit Sub
    
    Set wsJudje = ThisWorkbook.Sheets("judje")
    mainAcc = Trim(Me.TxtprenterNo.Value)
    accType = LCase(Trim(Me.cboType.Value))
    
    
    ' ÃÌ» ÿ»Ì⁄… «·Õ”«» Ê‰Ê⁄ «· ﬁ—Ì— „‰ «·Õ”«» «·—∆Ì”Ì ›Ì ⁄„Êœ B
lastRow = wsJudje.Cells(wsJudje.Rows.Count, "B").End(xlUp).Row
For i = 3 To lastRow
    If Trim(wsJudje.Cells(i, "B").Value) = mainAcc Then
        Me.cboNature.Value = Trim(wsJudje.Cells(i, "I").Value)
        Me.cboReportType.Value = Trim(wsJudje.Cells(i, "G").Value)
        Exit For
    End If
Next i

    If accType <> "⁄«„" And accType <> "„”«⁄œ" And accType <> "›—⁄Ì" And accType <> " Õ·Ì·Ì" Then Exit Sub
    
    cnt = 0
    lastRow = wsJudje.Cells(wsJudje.Rows.Count, "D").End(xlUp).Row
    
    '
    For i = 3 To lastRow
        If Trim(wsJudje.Cells(i, "D").Value) = mainAcc _
           And LCase(Trim(wsJudje.Cells(i, "F").Value)) = accType Then
            cnt = cnt + 1
        End If
    Next i
    
    
    If accType = "⁄«„" Then
        If cnt + 1 > 9 Then
            MsgBox "·«Ì„ﬂ‰ «÷«›… Õ”«» «Œ— Â–Â «·„Ã„Ê⁄… „ﬂ „·… 9 Õ”«»« ", vbExclamation
            Me.txtAccountNo.Value = ""
            Exit Sub
        End If
        Me.txtAccountNo.Value = mainAcc & CStr(cnt + 1)
    
   
    Else
        If cnt + 1 > 99 Then
            MsgBox "«Ì„ﬂ‰ «÷«›… Õ”«» «Œ— Â–Â «·„Ã„Ê⁄… „ﬂ „·… 99 Õ”«»«  " & Me.cboType.Value, vbExclamation
            Me.txtAccountNo.Value = ""
            Exit Sub
        End If
        Me.txtAccountNo.Value = mainAcc & Format(cnt + 1, "00")
    End If


End Sub



Private Sub UserForm_Initialize()

    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    
    Set ws = ThisWorkbook.Sheets("data")
    
    Application.EnableEvents = False
    isUpdating = True
    
    Me.Cobchacc.clear
    
    lastRow = ws.Cells(ws.Rows.Count, "AL").End(xlUp).Row
    
    For i = 4 To lastRow
        If Trim(ws.Cells(i, "AL").Value) <> "" Then
            Me.Cobchacc.AddItem ws.Cells(i, "AL").Value
        End If
    Next i
    
    Application.EnableEvents = True
    isUpdating = False

End Sub

