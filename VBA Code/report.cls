VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim KeyCells As Range
    Set KeyCells = Me.Range("H2,F3,H3")
    
    If Intersect(Target, KeyCells) Is Nothing Then Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Dim i As Long
    Dim lastRow As Long

    ' ‰ÃÌ» ¬Œ— ’› ›ÌÂ „œÌ‰ √Ê œ«∆‰
    lastRow = Application.Max( _
        Me.Cells(Me.Rows.Count, "K").End(xlUp).Row, _
        Me.Cells(Me.Rows.Count, "L").End(xlUp).Row)
    
    If lastRow < 8 Then lastRow = 8

    ' ‰„”Õ ⁄„Êœ M ﬂ·Â „‰ ’› 8 ·¬Œ— «·‘Ì 
    Me.Range("M8:M" & Me.Rows.Count).ClearContents

    ' ‰ √ﬂœ ≈‰ M7 —ﬁ„
    If Not IsNumeric(Me.Range("M7").Value) Then Me.Range("M7").Value = 0

    ' ‰Õÿ «·„⁄«œ·… »” ··’›Ê› «··Ì ›ÌÂ« „œÌ‰ √Ê œ«∆‰
    Application.Calculation = xlCalculationManual
    For i = 8 To lastRow
        If IsNumeric(Me.Cells(i, "K").Value) Or IsNumeric(Me.Cells(i, "L").Value) Then
            Me.Cells(i, "M").Formula = "=M" & i - 1 & "+K" & i & "-L" & i
        End If
    Next i
    Application.Calculate
    Application.Calculation = xlCalculationAutomatic

    Me.Range("M8:M" & lastRow).NumberFormat = "#,##0.00 ""—.”"""
    Me.Rows("8:" & lastRow).AutoFit

    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
'Private Sub Worksheet_Calculate()
'    Dim lastRow As Long
'    Dim i As Long
'
'    Application.EnableEvents = False
'     Me.Range("M8:M100000").ClearContents
'    lastRow = Me.Cells(Me.Rows.Count, "K").End(xlUp).Row
'    If lastRow < 8 Then lastRow = 8
    
'    For i = 8 To lastRow
''      If Me.Cells(i, "K").Value <> "" Or Me.Cells(i, "L").Value <> "" Then
 '      ' If i = 8 Then
 '      '     Me.Cells(i, "M").Formula = "=K" & i & "-L" & i
 '      '     Else
 '           Me.Cells(i, "M").Formula = "=M" & i - 1 & "+K" & i & "-L" & i
 '       End If
 '          Me.Rows("8:" & lastRow).WrapText = True
 '          Me.Rows("8:" & lastRow).EntireRow.AutoFit
 ''
       ' End If
 '   Next i
 '
 '   Application.EnableEvents = True
'End Sub
