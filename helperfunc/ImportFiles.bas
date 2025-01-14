Attribute VB_Name = "ImportFiles"
Sub ImportFilestoWS()

    Application.DisplayStatusBar = True
'    Application.ScreenUpdating = False
'    Application.DisplayAlerts = False
''

    Dim src_wb As String, tgt_wss As String
    Dim controlWS As Worksheet, filterWS As Worksheet
    Dim reg_file As String

    Set controlWS = ThisWorkbook.Worksheets("Control")
    Set filterWS = ThisWorkbook.Worksheets("Filter")

'    Copy NAV File
    src_wb = controlWS.Range("nav_eod_path")
    tgt_ws = "NAV EOD"
    CopyDatatoWS src_wb, tgt_ws

    'Copy Holding File, Regex***
    reg_file = Dir(controlWS.Range("hld_eod_path") & "\SS LUX Positions by Settle Location_" & "*")
    src_wb = controlWS.Range("hld_eod_path") & "\" & reg_file
    tgt_ws = "Holding"
    CopyDatatoWS src_wb, tgt_ws, "I", TrimFilterRange(filterWS.Range("D:E"))
    
    'Copy Valuation File
    src_wb = controlWS.Range("val_path")
    tgt_ws = "Valuation"
    CopyDatatoWS src_wb, tgt_ws, "AL", TrimFilterRange(filterWS.Range("A:B"))
    
    'Copy Index Close File
    src_wb = controlWS.Range("idx_close_path")
    tgt_ws = "Index(Close)"
    CopyDatatoWS src_wb, tgt_ws, "AI"
    
    'Copy Index Open File
    src_wb = controlWS.Range("idx_open_path")
    tgt_ws = "Index(Open)"
    CopyDatatoWS src_wb, tgt_ws, "AI"
    
    'Copy Tracker File
    src_wb = controlWS.Range("idx_tracker_path")
    tgt_ws = "Index(Tracker)"
    CopyDatatoWS src_wb, tgt_ws
    
    'Copy 5D Tracker File
    src_wb = controlWS.Range("idx_5d_tracker_path")
    tgt_ws = "Index(5D Tracker)"
    CopyDatatoWS src_wb, tgt_ws
    
    'Copy FX File
    src_wb = controlWS.Range("idx_fx_path")
    tgt_ws = "Index(FX)"
    CopyDatatoWS src_wb, tgt_ws
    
    controlWS.Activate
        
    Set controlWS = Nothing
    Set filterWS = Nothing

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox "Source files are loaded."

End Sub
