VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Cod 
   Caption         =   "≈÷«›… ⁄„·‹‹‹‹‹…"
   ClientHeight    =   9540.001
   ClientLeft      =   -276
   ClientTop       =   -1068
   ClientWidth     =   13716
   OleObjectBlob   =   "Cod.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Cod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cboCurrType_Change()
    If cboCurrType.Value = "⁄„·… „Õ·Ì…" Then
     
        MinRate.Value = 1
        MaxRate.Value = 1
        txt_CrrRate.Value = 1
        
        MinRate.Locked = True
        MaxRate.Locked = True
        txt_CrrRate.Locked = True
        
        
        MinRate.BackColor = RGB(240, 240, 240)
        MaxRate.BackColor = RGB(240, 240, 240)
        txt_CrrRate.BackColor = RGB(240, 240, 240)
        
    ElseIf cboCurrType.Value = "⁄„·… «Ã‰»Ì…" Then
     
        MinRate.Locked = False
        MaxRate.Locked = False
        txt_CrrRate.Locked = False
        
        
        MinRate.BackColor = RGB(255, 255, 255)
        MaxRate.BackColor = RGB(255, 255, 255)
        txt_CrrRate.BackColor = RGB(255, 255, 255)
        
    
       ' MinRate.Value = ""
       ' MaxRate.Value = ""
       ' txt_CrrRate.Value = ""
    End If

End Sub

Private Sub clear_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub clear_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Me.txtCurrID.Value = ""
        Me.txtCurrName.Value = ""
        Me.txtCurrCode.Value = ""
        Me.txt_CrrRate.Value = ""
        Me.txt_CrrRate.Value = ""
        Me.MinRate.Value = ""
        Me.MaxRate.Value = ""
        Me.txtCurrName.Enabled = True
        Me.txtCurrCode.Enabled = True
        Me.save.Visible = True
        Me.edit.Visible = False
       
    Dim ws As Worksheet
    Dim MaxID As Long
    Dim i As Long
    
    Set ws = ThisWorkbook.Worksheets("data")
    
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    MaxID = 0
   
    For i = 3 To lastRow
        If IsNumeric(ws.Cells(i, "B").Value) Then
            If ws.Cells(i, "B").Value > MaxID Then
                MaxID = ws.Cells(i, "B").Value
            End If
        End If
    Next i
    
    txtCurrID.Value = MaxID + 1
    
    Me.txtCurrName.SetFocus
        
End Sub
Private Sub edit_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
On Error Resume Next
Dim rown As Integer
rown = Me.ListBox1.List(TRANS, 0) + 1
Me.ListBox1.ListIndex = -1
Me.ListBox1.RowSource = ""
 With ThisWorkbook.Sheets("data")

    .Range("A" & rown).Value = "=row()-2"
    .Range("C" & rown).Value = Me.txtCurrName.Value
    .Range("D" & rown).Value = Me.txtCurrCode.Value
    .Range("B" & rown).Value = Me.txtCurrID.Text
    .Range("E" & rown).Value = Me.txt_CrrRate.Value
    .Range("g" & rown).Value = Me.MaxRate.Text
    .Range("f" & rown).Value = Me.MinRate.Text
    .Range("H" & rown).Value = Me.cboCurrType.Text
    
    End With
    MsgBox " „  ⁄œÌ· «·⁄„·Ì… »‰Ã«Õ"
    
     Me.txtCurrID.Value = ""
        Me.txtCurrName.Value = ""
        Me.txtCurrCode.Value = ""
        Me.txt_CrrRate.Value = ""
        Me.txt_CrrRate.Value = ""
        Me.MinRate.Value = ""
        Me.MaxRate.Value = ""
     Me.txtCurrName.Enabled = True
     Me.txtCurrCode.Enabled = True
     Me.ListBox1.RowSource = "tablecurrncy"
     
    Me.txtCurrName.SetFocus
    Dim MaxID As Long
    Dim i As Long
    
    Set ws = ThisWorkbook.Worksheets("data")
    
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    MaxID = 0
   
    For i = 3 To lastRow
        If IsNumeric(ws.Cells(i, "B").Value) Then
            If ws.Cells(i, "B").Value > MaxID Then
                MaxID = ws.Cells(i, "B").Value
            End If
        End If
    Next i
    
    txtCurrID.Value = MaxID + 1

End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Me.save.Visible = False

Me.txtCurrName.Enabled = False
Me.txtCurrCode.Enabled = False
Me.TRANS.Value = Me.ListBox1.List(Me.ListBox1.ListIndex, 0)
Me.txtCurrID.Value = Me.ListBox1.List(Me.ListBox1.ListIndex, 1)
Me.txtCurrName.Value = Me.ListBox1.List(Me.ListBox1.ListIndex, 2)
Me.txtCurrCode.Value = Me.ListBox1.List(Me.ListBox1.ListIndex, 3)
Me.txt_CrrRate.Text = Me.ListBox1.List(Me.ListBox1.ListIndex, 4)
Me.MinRate.Text = Me.ListBox1.List(Me.ListBox1.ListIndex, 5)
Me.MaxRate.Value = Me.ListBox1.List(Me.ListBox1.ListIndex, 6)
Me.cboCurrType.Value = Me.ListBox1.List(Me.ListBox1.ListIndex, 7)
Me.del.Visible = True
Me.edit.Visible = True
End Sub

Private Sub MinRate_AfterUpdate()

   
    Dim MinR As Double
    Dim MaxR As Double
    Dim Rate As Double
    
   
    If cboCurrType.Value = "⁄„·… „Õ·Ì…" Then Exit Sub
    
    If Trim(MinRate.Value) = "" Then Exit Sub
    If Not IsNumeric(MinRate.Value) Then
        MsgBox "«·Õœ «·«œ‰Ï ÌÃ» «‰ ÌﬂÊ‰ —ﬁ„", vbExclamation
        MinRate.Value = ""
        MinRate.SetFocus
        Exit Sub
    End If
    
    MinR = val(MinRate.Value)
    
  
    If MinR <= 1 Then
        MsgBox "«·Õœ «·«œ‰Ï ÌÃ» «‰ ÌﬂÊ‰ «ﬂ»— „‰ 1", vbExclamation
        MinRate.Value = ""
        MinRate.SetFocus
        Exit Sub
    End If
    
    
    If Trim(txt_CrrRate.Value) <> "" Then
        Rate = val(txt_CrrRate.Value)
        If MinR > Rate Then
            MsgBox "«·Õœ «·«œ‰Ï ·«Ì„ﬂ‰ «‰ ÌﬂÊ‰ «ﬂ»— „‰ ”⁄— «·’—›", vbExclamation
            MinRate.Value = ""
            MinRate.SetFocus
            Exit Sub
        End If
    End If
    
    
    If Trim(MaxRate.Value) <> "" Then
        MaxR = val(MaxRate.Value)
        If MinR >= MaxR Then
            MsgBox "«·Õœ «·«œ‰Ï ÌÃ» «‰ ÌﬂÊ‰ «ﬁ· „‰ «·Õœ «·«⁄·Ï", vbExclamation
            MinRate.Value = ""
            MinRate.SetFocus
            Exit Sub
        End If
    End If
End Sub

Private Sub MaxRate_AfterUpdate()
    Dim MinR As Double
    Dim MaxR As Double
    Dim Rate As Double
    
    If cboCurrType.Value = "⁄„·… „Õ·Ì…" Then Exit Sub
    
    If Trim(MaxRate.Value) = "" Then Exit Sub
    If Not IsNumeric(MaxRate.Value) Then
        MsgBox "«·Õœ «·«⁄·Ï ÌÃ» «‰ ÌﬂÊ‰ —ﬁ„", vbExclamation
        MaxRate.Value = ""
        MaxRate.SetFocus
        Exit Sub
    End If
    
    MaxR = val(MaxRate.Value)
    
    
    If Trim(MinRate.Value) <> "" Then
        MinR = val(MinRate.Value)
        If MaxR <= MinR Then
            MsgBox "«·Õœ «·«⁄·Ï ÌÃ» «‰ ÌﬂÊ‰ «ﬂ»— „‰ «·Õœ «·«œ‰Ï", vbExclamation
            MaxRate.Value = ""
            MaxRate.SetFocus
            Exit Sub
        End If
    End If
    
    
    If Trim(txt_CrrRate.Value) <> "" Then
        Rate = val(txt_CrrRate.Value)
        If MaxR < Rate Then
            MsgBox "«·Õœ «·«⁄·Ï ÌÃ» «‰ ÌﬂÊ‰ «ﬂ»— «Ê Ì”«ÊÌ ”⁄— «·’—›", vbExclamation
            MaxRate.Value = ""
            MaxRate.SetFocus
            Exit Sub
        End If
    End If
End Sub
Private Sub save_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
   Dim ws As Worksheet
   Dim lastRow As Long
   Dim r As Range
   Dim curName As String
   
   Set ws = ThisWorkbook.Worksheets("data")
   curName = Trim(txtCurrName.Value)
   
   If Trim(txtCurrID.Value) = "" Or Trim(txtCurrCode.Value) = "" Or curName = "" Or Trim(cboCurrType.Value) = "" Then
        MsgBox prompt:="«·—Ã«¡ «· √ﬂœ „‰ ≈œŒ«· Ã„Ì⁄ »Ì«‰«  «·≈–‰", Title:="Œÿ√⁄œ„ «ﬂ „«· «·»Ì«‰« "
    Exit Sub
    End If
    
    Set r = ws.Range("c:c").Find(what:=curName, lookat:=xlWhole, MatchCase:=False)
       If Not r Is Nothing Then
       MsgBox "«·⁄„·… „œŒ·… ”«»ﬁ«", vbExclamation, " ‰»ÌÂ"
    Exit Sub
    End If
    
 If cboCurrType.Value = "" Then
        If Trim(txt_CrrRate.Value) = "" Or Not IsNumeric(txt_CrrRate.Value) Then
        MsgBox "Ì—ÃÏ «œŒ«· ”⁄— «·’—› ﬂ—ﬁ„", vbExclamation, " ‰»ÌÂ"
        txt_CrrRate.SetFocus
    Exit Sub
    End If
    
    If Trim(MinRate.Value) = "" Or Not IsNumeric(MinRate.Value) Then
            MsgBox "Ì—ÃÏ «œŒ«· «·Õœ «·«œ‰Ï ﬂ—ﬁ„", vbExclamation, " ‰»ÌÂ"
            MinRate.SetFocus
    Exit Sub
    End If
    
    If Trim(MaxRate.Value) = "" Or Not IsNumeric(MaxRate.Value) Then
            MsgBox "Ì—ÃÏ «œŒ«· «·Õœ «·«⁄Ï ﬂ—ﬁ„", vbExclamation, " ‰»ÌÂ"
            MaxRate.SetFocus
        Exit Sub
        End If
End If
Me.ListBox1.RowSource = ""
''''''''''''«÷«›…''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sheets("data").Activate
 With ThisWorkbook.Sheets("data")
    .Range("A3:h3").Select
    Selection.ListObject.ListRows.Add (1)
  
    .Range("B3").Value = Me.txtCurrID.Value
    .Range("C3").Value = Me.txtCurrName.Value
    .Range("d3").Value = Me.txtCurrCode.Value
    .Range("e3").Value = Me.txt_CrrRate.Value
    .Range("f3").Value = Me.MinRate.Value
    .Range("g3").Value = Me.MaxRate.Value
    .Range("h3").Value = Me.cboCurrType.Value
         End With
         
          MsgBox (" „ «÷«›… «·⁄„·… »‰Ã«Õ")
          
        Me.txtCurrID.Value = ""
        Me.txtCurrName.Value = ""
        Me.txtCurrCode.Value = ""
        Me.txt_CrrRate.Value = ""
        Me.txt_CrrRate.Value = ""
        Me.MinRate.Value = ""
        Me.MaxRate.Value = ""
        
        Me.ListBox1.RowSource = "tablecurrncy"
          
  
    Dim MaxID As Long
    Dim i As Long
    
    Set ws = ThisWorkbook.Worksheets("data")
    
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    MaxID = 0
   
    For i = 3 To lastRow
        If IsNumeric(ws.Cells(i, "B").Value) Then
            If ws.Cells(i, "B").Value > MaxID Then
                MaxID = ws.Cells(i, "B").Value
            End If
        End If
    Next i
    
    txtCurrID.Value = MaxID + 1

End Sub

Private Sub txt_CrrRate_AfterUpdate()

    Dim MinR As Double
    Dim MaxR As Double
    Dim Rate As Double
    

    If cboCurrType.Value = "⁄„·… „Õ·Ì…" Then Exit Sub
    
  
    If Trim(txt_CrrRate.Value) = "" Then Exit Sub
    
  
    If Not IsNumeric(txt_CrrRate.Value) Then
        MsgBox "”⁄— «·’—› ÌÃ» «‰ ÌﬂÊ‰ —ﬁ„", vbExclamation
        txt_CrrRate.Value = ""
        txt_CrrRate.SetFocus
        Exit Sub
    End If
    
    Rate = val(txt_CrrRate.Value)
    
  
    If Trim(MinRate.Value) = "" Or Trim(MaxRate.Value) = "" Then
        MsgBox "Ì—ÃÏ  ÕœÌœ «·Õœ «·«œ‰Ï Ê«·Õœ «·«⁄·Ï «Ê·«", vbExclamation
        txt_CrrRate.Value = ""
        MinRate.SetFocus
        Exit Sub
    End If
    
    MinR = val(MinRate.Value)
    MaxR = val(MaxRate.Value)
    
    
    If Rate > MaxR Then
        MsgBox "”⁄— «·’—› «·–Ì «œŒ· Â «ﬂ»— „‰ «·Õœ «·«⁄·Ï", vbExclamation
        txt_CrrRate.Value = ""
        txt_CrrRate.SetFocus
        Exit Sub
    End If
    

    If Rate < MinR Then
        MsgBox "”⁄— «·’—› «·–Ì «œŒ· Â «ﬁ· „‰ «·Õœ «·«œ‰Ï", vbExclamation
        txt_CrrRate.Value = ""
        txt_CrrRate.SetFocus
        Exit Sub
    End If
End Sub



Private Sub UserForm_Initialize()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim MaxID As Long
    Dim i As Long
      Dim HasLocal As Boolean
   ' Set ws = ThisWorkbook.Worksheets("data")
    
   ' LastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
   ' MaxID = 0
   
   ' For i = 3 To LastRow
   '     If IsNumeric(ws.Cells(i, "B").Value) Then
   '         If ws.Cells(i, "B").Value > MaxID Then
   '             MaxID = ws.Cells(i, "B").Value
   '         End If
   '     End If
   ' Next i
    
   ' txtCurrID.Value = MaxID + 1


''
  
    
    Set ws = ThisWorkbook.Worksheets("data")
    
    
    cboCurrType.clear
    cboCurrType.AddItem "⁄„·… „Õ·Ì…"
    cboCurrType.AddItem "⁄„·… «Ã‰»Ì…"
    
   
    lastRow = ws.Cells(ws.Rows.Count, "H").End(xlUp).Row
    HasLocal = False
    
    For i = 3 To lastRow
        If Trim(ws.Cells(i, "H").Value) = "⁄„·… „Õ·Ì…" Then
            HasLocal = True
            Exit For
        End If
    Next i
    
   
    If HasLocal Then
        cboCurrType.Value = "⁄„·… «Ã‰»Ì…"
        cboCurrType.Enabled = False
    Else
        cboCurrType.Enabled = True
    End If
    
    
    MaxID = 0
    For i = 2 To ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
        If IsNumeric(ws.Cells(i, "B").Value) Then
            If ws.Cells(i, "B").Value > MaxID Then MaxID = ws.Cells(i, "B").Value
        End If
    Next i
    txtCurrID.Value = MaxID + 1

End Sub
