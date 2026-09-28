VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RPACCOUNT 
   Caption         =   "ﬂ‘‹‹› Õ”‹‹‹«»"
   ClientHeight    =   4392
   ClientLeft      =   -864
   ClientTop       =   -3072
   ClientWidth     =   12852
   OleObjectBlob   =   "RPACCOUNT.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RPACCOUNT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit

Private Sub Cod_AfterUpdate()
 '  Dim i, lr As Long
 '
 '  lr = judje.Cells(judje.Rows.Count, "B").End(xlUp).Row
 '  For i = 3 To lr
 'If val(Me.Cod) = judje.Cells(i, "B") Then
 '     Me.itms.Value = judje.Cells(i, "C").Value
 'End If
 '  Next i
End Sub

Private Sub ComClear_Click()
Me.itms.Value = ""
Me.Cod.Value = ""

End Sub

Private Sub ComSearch_Click()
  On Error Resume Next
  If Trim(Me.itms.Value) = "" Or Trim(Me.Cod.Value) = "" Or Trim(Me.da.Value) = "" Or Trim(Me.da1.Value) = "" Then
      MsgBox "«·—Ã«¡ «œŒ«· ﬂ«›… «·»Ì«‰« ", vbExclamation, " ‰»ÌÂ"
      Exit Sub
      End If
      
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("report")
    
    With ws
     .Cells(2, "f").Value = Me.itms.Value
      .Cells(2, "h").Value = Me.Cod.Value
    .Cells(2, "t").Value = Format(Me.da.Value, "yyyy/mm/dd")
    .Cells(3, "t").Value = Format(Me.da1.Value, "yyyy/mm/dd")
     End With
        ws.Activate
        ws.Range("c7").Select
        Unload Me
        
End Sub

Private Sub Frame1_Click()

End Sub

Private Sub itms_AfterUpdate()
Dim i, lr As Long
Dim judje As Worksheet
  Set judje = ThisWorkbook.Sheets("judje")
  
lr = judje.Cells(judje.Rows.Count, "C").End(xlUp).Row
For i = 3 To lr
If Me.itms.Value = judje.Cells(i, "C").Value Then
Me.Cod.Value = judje.Cells(i, "B").Value
End If

Next i

End Sub

Private Sub da_AfterUpdate()
da = Format(da, "yyyy/mm/dd")
End Sub

Private Sub da1_AfterUpdate()
da1 = Format(da1, "yyyy/mm/dd")
End Sub
Private Sub UserForm_Initialize()
Me.Height = 247
Me.Width = 655

Me.da.Value = Format(report.Cells(2, "S").Value, "yyyy/mm/dd")
Me.da1.Value = Format(report.Cells(3, "s").Value, "yyyy/mm/dd")


Dim i As Long, lastRow As Long

Me.itms.clear

lastRow = Sheets("judje").Cells(Rows.Count, "C").End(xlUp).Row

For i = 3 To lastRow
    If Sheets("judje").Cells(i, "E").Value = 5 Then
        Me.itms.AddItem Sheets("judje").Cells(i, "C").Value
    End If
Next i

End Sub
