Attribute VB_Name = "ClearWorksheetContent"
Option Explicit

'This module helps delete data from target worksheet (workbook must be activated first)
'It could keep data header or any formula within the worksheet (formula can only be on the right side of data)


'Developed and last updated by
'West Wang, 2022.7.11

Sub ClearWS(ByVal ShtName As String, Optional ByVal firstCell As String = "A1", _
                                     Optional ByVal withHeader As Boolean = False, _
                                     Optional ByVal keepFormula As Boolean = False)
'firstCell: topleft of the data range
                                     

Dim LastClear As Integer, RightClear As Integer
Dim topLeft As Range, BottomRight As Range

' Clear worksheet ShtName
On Error Resume Next:

With Application.ThisWorkbook.Worksheets(ShtName)

    LastClear = .Cells(.Rows.Count, Range(firstCell).Column).End(xlUp).Row
    RightClear = .Range(firstCell).End(xlToRight).Column
    
    Set topLeft = .Range(firstCell)
    Set BottomRight = .Cells(LastClear, RightClear)
    
    ' Clear data but keep header and row 2 formula
    If withHeader And keepFormula Then
        If LastClear >= .Range(firstCell).Row Then 'there are data
            .Range(topLeft.Offset(2, 0), BottomRight.Offset(1, 0)).ClearContents
            topLeft.Offset(1, 0).Resize(1, RightClear - topLeft.Column).SpecialCells(xlCellTypeConstants, 23).ClearContents
        End If
    ElseIf Not withHeader And keepFormula Then
        .Range(topLeft.Offset(1, 0), BottomRight.Offset(1, 0)).ClearContents
        topLeft.Resize(1, RightClear - topLeft.Column).SpecialCells(xlCellTypeConstants, 23).ClearContents 'remove non-formula cells
    ElseIf withHeader Then
        If LastClear >= .Range(firstCell).Row Then 'there are data
            .Range(topLeft.Offset(1, 0), .Cells(LastClear, RightClear)).ClearContents
        End If
    Else
        If LastClear >= .Range(firstCell).Row Then 'there are data
            .Range(topLeft, .Cells(LastClear, RightClear)).ClearContents
        End If
    End If
    
    Set topLeft = Nothing
    Set BottomRight = Nothing

End With

End Sub


Sub TrimWSCol(ByRef targetWSName As String, ByVal colLetter As String)
    
    'Column will be set to String format and trimmed
    Dim rng As Range

    With ThisWorkbook.Worksheets(targetWSName)
        
        Set rng = Range(colLetter & ":" & colLetter)
        rng.NumberFormat = "@"
        rng.Value = Application.Trim(rng)
        Set rng = Nothing
    
    End With

End Sub
