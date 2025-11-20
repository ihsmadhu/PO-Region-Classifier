Attribute VB_Name = "Module9"
Sub Clear_POData()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("POData")
    
    ' Clear PO numbers + Region column (A & B, below header)
    ws.Range("A2:B" & ws.Rows.Count).ClearContents
    
    ' Clear vertical totals only (E:F)
    ws.Range("E1:F50").ClearContents
    
    MsgBox "POData cleared (PO numbers and totals).", vbInformation
End Sub

