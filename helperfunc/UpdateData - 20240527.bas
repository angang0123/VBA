Attribute VB_Name = "UpdateData"
Sub UpdateDatatoWS()

    Dim srcWS As Worksheet, tgtWS As Worksheet
    Dim dataRange As Range
    Dim row As Long
    
    Dim matchResult As Variant

    Set tgtWS = ThisWorkbook.Worksheets("Data")
    '***********Set Copy data range*************
    Set srcDataRange = ThisWorkbook.Worksheets("Data Extract").Range("B3:AX3")
    
    matchResult = Application.Match(ThisWorkbook.Names("fund_date").RefersToRange, tgtWS.Range("A:A"), 0)
    
    If Not IsError(matchResult) And ThisWorkbook.Names("override").RefersToRange.Value = "Yes" Then
        'override current data
        row = matchResult
        srcDataRange.Copy
        tgtWS.Cells(row, 1).PasteSpecial Paste:=xlPasteValues
        
    ElseIf IsError(matchResult) Then
        'not exist, add data
        row = tgtWS.Range("A1").End(xlDown).row + 1
        If row > 1000000 Then
            row = 2
        End If
        srcDataRange.Copy
        tgtWS.Cells(row, 1).PasteSpecial Paste:=xlPasteValues
        
'        tgtWS.UsedRange.Select
'        tgtWS.Sort.SortFields.Add2 Key:=Range("A:A"), SortOn:=xlSortOnValues, Order:=xlAscending
        With tgtWS.Sort
            .SortFields.Clear
            .SortFields.Add Key:=tgtWS.Range("A:A"), Order:=xlAscending
            .SetRange tgtWS.UsedRange
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .Apply
        End With
        
    ElseIf Not IsError(matchResult) And ThisWorkbook.Names("override").RefersToRange.Value = "No" Then
        MsgBox "Fund date already has data and not override."
    
    End If
    

    Set tgtWS = Nothing
    Set srcDataRange = Nothing

End Sub
