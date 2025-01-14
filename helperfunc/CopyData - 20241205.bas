Attribute VB_Name = "CopyData"
Sub CopyDatatoWS(ByVal src_wb As String, ByVal tgt_ws As String, _
                Optional ByVal src_ws As String = "", Optional ByVal last_col As String = "", _
                Optional ByVal ColumnNameFilterRange As Range = Nothing)
    ' if last_col <> "", there are formulas on the right-hand columns
    ' ColumnNameFilterRange contains two columns, one for colname, one for filter. If not Empty then filter source file
    Dim sourceWorkbook As Workbook
    Dim sourceWorksheet As Worksheet, currentWorksheet As Worksheet
    Dim current_last_col As Long, current_last_row As Long
    Dim source_last_row As Long
    Dim loopCount As Long, filterLastRow As Long, fieldNum As Long
    Dim filterRow As Variant
    Dim srcColName As String
    Dim srcFilValue As Variant
    
    Set sourceWorkbook = Workbooks.Open(src_wb)
    If src_ws = "" Then
        Set sourceWorksheet = sourceWorkbook.Sheets(1)
    Else
        Set sourceWorksheet = sourceWorkbook.Sheets(src_ws)
    End If
    Set currentWorksheet = ThisWorkbook.Sheets(tgt_ws)
    
    If Not (ColumnNameFilterRange Is Nothing) Then 'filter is not empty
        loopCount = 0
        For Each filterRow In ColumnNameFilterRange.Rows
            If loopCount > 0 Then
                srcColName = filterRow.Cells(1, 1).Value
                srcFilValue = Split(filterRow.Cells(1, 2).Value, ",")
                fieldNum = Application.Match(srcColName, sourceWorksheet.Range("1:1"), 0)
                sourceWorksheet.Cells.AutoFilter field:=fieldNum, Criteria1:=srcFilValue, Operator:=xlFilterValues
            End If
            loopCount = loopCount + 1
        Next filterRow
    End If
    
    'Clear current worksheet
    If last_col = "" Then
        currentWorksheet.Cells.ClearContents
        sourceWorksheet.UsedRange.Copy
        currentWorksheet.Range("A1").PasteSpecial Paste:=xlPasteAll
    Else
        If currentWorksheet.Range("A1").End(xlDown).row > 2 Then
            currentWorksheet.Range("3:" & currentWorksheet.Range("A1").End(xlDown).row).ClearContents
        End If
        currentWorksheet.Range("A:" & last_col).Clear
        
        source_last_row = sourceWorksheet.Cells(sourceWorksheet.Rows.Count, "A").End(xlUp).row
        sourceWorksheet.Range("A1:" & last_col & source_last_row).Copy
        currentWorksheet.Range("A1").PasteSpecial Paste:=xlPasteAll
        
        current_last_row = currentWorksheet.Range("A1").End(xlDown).row
        current_last_col = currentWorksheet.Range(last_col & "1").End(xlToRight).Column
        
        'copy down formula
        If currentWorksheet.Cells(sourceWorksheet.Rows.Count, "A").End(xlUp).row > 2 Then
            currentWorksheet.Range(currentWorksheet.Cells(2, last_col).Offset(0, 1), currentWorksheet.Cells(current_last_row, current_last_col)).FillDown
        End If
    End If
    
    Application.CutCopyMode = False
    
    sourceWorkbook.Close False
    
    Set sourceWorkbook = Nothing
    Set sourceWorksheet = Nothing
    Set currentWorksheet = Nothing
    
End Sub


Sub CopyMultiWSData(ByVal src_wb As String, ByVal src_wss As Variant, ByVal tgt_wss As Variant)
    
    Dim sourceWorkbook As Workbook
    Dim i As Long
    
    Set sourceWorkbook = Workbooks.Open(src_wb)
    
    For i = LBound(src_wss) To UBound(src_wss)
        Set sourceWorksheet = sourceWorkbook.Sheets(src_wss(i))
        Set currentWorksheet = ThisWorkbook.Sheets(tgt_wss(i))
        
        currentWorksheet.Cells.Clear
        sourceWorksheet.Cells.Copy currentWorksheet.Range("A1")
    Next i
    
    sourceWorkbook.Close False
    Set sourceWorkbook = Nothing
    Set sourceWorksheet = Nothing
    Set currentWorksheet = Nothing

End Sub

Function TrimFilterRange(ByVal fulRange As Range) As Range
    Dim last_row As Long
    
    last_row = fulRange.End(xlDown).row
    
    Set TrimFilterRange = fulRange.Cells(1, 1).Resize(last_row, 2)

End Function





