Attribute VB_Name = "FunaccountNo"
Option Explicit

Function GenAccNo(parentNo As String, AccLevel As Integer) As String
    
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim lastSeq As Long
    Dim prefix As String
    
    On Error GoTo ErrHandler  '  Ã«Â· «·√Œÿ«¡
    
    Set ws = ThisWorkbook.Worksheets("judje")
    
    Select Case AccLevel
        
        Case 1
            GenAccNo = Format(parentNo, "0")
        
        Case 2
            prefix = Format(parentNo, "0")
            lastSeq = 0
            lastRow = ws.Cells(Rows.Count, "B").End(xlUp).Row
            For i = 3 To lastRow
                If Len(ws.Cells(i, "B").Value) = 2 Then
                    If Left(ws.Cells(i, "B").Value, 1) = prefix Then
                        lastSeq = Application.Max(lastSeq, CLng(Right(ws.Cells(i, "B").Value, 1)))
                    End If
                End If
            Next i
            GenAccNo = prefix & Format(lastSeq + 1, "0")
        
        Case 3
            prefix = Format(parentNo, "00")
            lastSeq = 0
            lastRow = ws.Cells(Rows.Count, "B").End(xlUp).Row
            For i = 3 To lastRow
                If Len(ws.Cells(i, "B").Value) = 3 Then
                    If Left(ws.Cells(i, "B").Value, 2) = prefix Then
                        lastSeq = Application.Max(lastSeq, CLng(Right(ws.Cells(i, "B").Value, 1)))
                    End If
                End If
            Next i
            GenAccNo = prefix & Format(lastSeq + 1, "0")
        
        Case 4
            prefix = parentNo
            lastSeq = 0
            lastRow = ws.Cells(Rows.Count, "B").End(xlUp).Row
            For i = 3 To lastRow
                If Len(ws.Cells(i, "B").Value) = 5 Then
                    If Left(ws.Cells(i, "B").Value, Len(prefix)) = prefix Then
                        lastSeq = Application.Max(lastSeq, CLng(Right(ws.Cells(i, "B").Value, 2)))
                    End If
                End If
            Next i
            GenAccNo = prefix & Format(lastSeq + 1, "00")
        
        Case 5
            prefix = parentNo
            lastSeq = 0
            lastRow = ws.Cells(Rows.Count, "B").End(xlUp).Row
            For i = 3 To lastRow
                If Len(ws.Cells(i, "B").Value) = 7 Then
                    If Left(ws.Cells(i, "B").Value, Len(prefix)) = prefix Then
                        lastSeq = Application.Max(lastSeq, CLng(Right(ws.Cells(i, "B").Value, 2)))
                    End If
                End If
            Next i
            GenAccNo = prefix & Format(lastSeq + 1, "00")
        
        Case 6
            GenAccNo = parentNo & "1"
        
        Case Else
            GenAccNo = ""
    End Select
    Exit Function

ErrHandler:
    GenAccNo = ""  ' ·Ê ’«— Œÿ√ Ì—Ã⁄ ›«÷Ì
End Function
