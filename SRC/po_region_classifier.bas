Attribute VB_Name = "Module12"
Sub Classify_ByRegion_ApacSet()
    Dim wsData As Worksheet, wsMap As Worksheet
    Dim lastRow As Long, lastMap As Long
    Dim i As Long, j As Long
    Dim prefix As String, region As String
    Dim dict As Object, totals As Object
    Dim key As Variant, totalsRow As Long
    Dim order As Variant
    
    Set wsData = ThisWorkbook.Sheets("POData")
    Set wsMap = ThisWorkbook.Sheets("POMappings")
    
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    lastMap = wsMap.Cells(wsMap.Rows.Count, "A").End(xlUp).Row
    
    ' Build dictionary: Prefix ? GlobalRegion (Col C)
    Set dict = CreateObject("Scripting.Dictionary")
    For j = 2 To lastMap
        If wsMap.Cells(j, 1).Value <> "" Then
            dict(UCase(wsMap.Cells(j, 1).Value)) = UCase(wsMap.Cells(j, 3).Value)
        End If
    Next j
    
    ' Clear old totals (ONLY vertical region block)
    wsData.Range("E1:F50").ClearContents
    
    ' Classify each PO into Column B
    For i = 2 To lastRow
        prefix = UCase(Left(wsData.Cells(i, 1).Value, 2))
        region = IIf(dict.Exists(prefix), dict(prefix), "Unassigned")
        wsData.Cells(i, 2).Value = region
    Next i
    
    ' Build totals
    Set totals = CreateObject("Scripting.Dictionary")
    totals("AMER") = 0
    totals("APAC") = 0
    totals("EMEA") = 0
    totals("Unassigned") = 0
    
    For i = 2 To lastRow
        region = wsData.Cells(i, 2).Value
        If totals.Exists(region) Then
            totals(region) = totals(region) + 1
        Else
            totals("Unassigned") = totals("Unassigned") + 1
        End If
    Next i
    
    ' Fixed order
    order = Array("AMER", "APAC", "EMEA", "Unassigned")
    
    ' Vertical table ONLY
    wsData.Range("E1").Value = "Region"
    wsData.Range("F1").Value = "Count"
    
    totalsRow = 2
    For Each key In order
        wsData.Cells(totalsRow, 5).Value = key
        wsData.Cells(totalsRow, 6).Value = totals(key)
        totalsRow = totalsRow + 1
    Next key
    
    wsData.Cells(totalsRow, 5).Value = "Total"
    wsData.Cells(totalsRow, 6).Value = _
        Application.Sum(wsData.Range("F2:F" & totalsRow - 1))

    MsgBox "Classification complete.", vbInformation
End Sub


