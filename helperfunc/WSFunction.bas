Attribute VB_Name = "WSFunction"
Option Explicit

Sub ClearWS(ByVal ShtName As String, Optional ByVal firstCell As String = "A1", _
                                     Optional ByVal withHeader As Boolean = False, _
                                     Optional ByVal keepFormula As Boolean = False)

Dim LastClear As Integer, RightClear As Integer
Dim TopLeft As Range, BottomRight As Range

' Clear worksheet ShtName
On Error Resume Next:

With Application.ThisWorkbook.Worksheets(ShtName)

    LastClear = .Cells(.Rows.Count, Range(firstCell).Column).End(xlUp).row
    RightClear = .Range(firstCell).End(xlToRight).Column
    
    Set TopLeft = .Range(firstCell)
    Set BottomRight = .Cells(LastClear, RightClear)
    
    ' Clear data but keep header and row 2 formula
    If withHeader And keepFormula Then
        If LastClear > .Range(firstCell).row Then 'there are data
            .Range(TopLeft.Offset(2, 0), BottomRight.Offset(1, 0)).ClearContents
            TopLeft.Offset(1, 0).Resize(1, RightClear - TopLeft.Column).SpecialCells(xlCellTypeConstants, 23).ClearContents
        End If
    ElseIf Not withHeader And keepFormula Then
        .Range(TopLeft.Offset(1, 0), BottomRight.Offset(1, 0)).ClearContents
        TopLeft.Resize(1, RightClear - TopLeft.Column).SpecialCells(xlCellTypeConstants, 23).ClearContents
    ElseIf withHeader Then
        If LastClear > .Range(firstCell).row Then 'there are data
            .Range(TopLeft.Offset(1, 0), .Cells(LastClear, RightClear)).ClearContents
        End If
    Else
        If LastClear >= .Range(firstCell).row Then 'there are data
            .Range(TopLeft, .Cells(LastClear, RightClear)).ClearContents
        End If
    End If
    
    Set TopLeft = Nothing
    Set BottomRight = Nothing

End With

End Sub

