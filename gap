Sub Create_Market_Gap_Analysis()

    Dim wsData As Worksheet
    Dim wsOut As Worksheet
    Dim lastRow As Long
    Dim lastUsedRow As Long
    Dim lastUsedCol As Long
    Dim dictMarket As Object
    Dim dictSteps As Object
    Dim i As Long, r As Long
    Dim key As Variant
    Dim stepKey As String
    Dim lastMarketCol As Long
    Dim lastOutputRow As Long
    Dim dataRange As Range
    Dim firstDataCol As Long
    Dim avgCol As Long
    
    Set wsData = Sheets("Sheet1")
    Set wsOut = Sheets("Output")
    
    '-----------------------------------
    ' CLEAR ONLY GENERATED AREA
    '-----------------------------------
    lastUsedRow = wsOut.Cells(wsOut.Rows.Count, 1).End(xlUp).Row
    lastUsedCol = wsOut.Cells(1, wsOut.Columns.Count).End(xlToLeft).Column
    
    If lastUsedRow > 1 Then
        wsOut.Range(wsOut.Cells(1, 1), wsOut.Cells(lastUsedRow, lastUsedCol)).Clear
        wsOut.Range(wsOut.Cells(1, 1), wsOut.Cells(lastUsedRow, lastUsedCol)).FormatConditions.Delete
    End If
    
    Set dictMarket = CreateObject("Scripting.Dictionary")
    Set dictSteps = CreateObject("Scripting.Dictionary")
    
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    
    '-----------------------------------
    ' 1?? Collect Unique Markets
    '-----------------------------------
    For i = 2 To lastRow
        If LCase(Trim(wsData.Cells(i, 11).Value)) = "yes" Then
            If Not dictMarket.exists(wsData.Cells(i, 6).Value) Then
                dictMarket.Add wsData.Cells(i, 6).Value, dictMarket.Count + 8
            End If
        End If
    Next i
    
    '-----------------------------------
    ' 2?? Collect Unique Step Combination
    '-----------------------------------
    For i = 2 To lastRow
        If LCase(Trim(wsData.Cells(i, 11).Value)) = "yes" Then
            
            stepKey = wsData.Cells(i, 2).Value & "|" & _
                      wsData.Cells(i, 3).Value & "|" & _
                      wsData.Cells(i, 4).Value & "|" & _
                      wsData.Cells(i, 5).Value & "|" & _
                      wsData.Cells(i, 7).Value & "|" & _
                      wsData.Cells(i, 8).Value & "|" & _
                      wsData.Cells(i, 9).Value
                      
            If Not dictSteps.exists(stepKey) Then
                dictSteps.Add stepKey, dictSteps.Count + 2
            End If
        End If
    Next i
    
    '-----------------------------------
    ' 3?? Write Headers
    '-----------------------------------
    wsOut.Cells(1, 1).Value = "Tower"
    wsOut.Cells(1, 2).Value = "L2"
    wsOut.Cells(1, 3).Value = "L3"
    wsOut.Cells(1, 4).Value = "L4"
    wsOut.Cells(1, 5).Value = "Entity Variant"
    wsOut.Cells(1, 6).Value = "Broad Step"
    wsOut.Cells(1, 7).Value = "Step Name"
    
    For Each key In dictMarket.Keys
        wsOut.Cells(1, dictMarket(key)).Value = key
    Next key
    
    '-----------------------------------
    ' 4?? Write Step Rows
    '-----------------------------------
    For Each key In dictSteps.Keys
        
        r = dictSteps(key)
        
        wsOut.Cells(r, 1).Value = Split(key, "|")(0)
        wsOut.Cells(r, 2).Value = Split(key, "|")(1)
        wsOut.Cells(r, 3).Value = Split(key, "|")(2)
        wsOut.Cells(r, 4).Value = Split(key, "|")(3)
        wsOut.Cells(r, 5).Value = Split(key, "|")(4)
        wsOut.Cells(r, 6).Value = Split(key, "|")(5)
        wsOut.Cells(r, 7).Value = Split(key, "|")(6)
        
    Next key
    
    '-----------------------------------
    ' 5?? Fill Duration
    '-----------------------------------
    For i = 2 To lastRow
        If LCase(Trim(wsData.Cells(i, 11).Value)) = "yes" Then
            
            stepKey = wsData.Cells(i, 2).Value & "|" & _
                      wsData.Cells(i, 3).Value & "|" & _
                      wsData.Cells(i, 4).Value & "|" & _
                      wsData.Cells(i, 5).Value & "|" & _
                      wsData.Cells(i, 7).Value & "|" & _
                      wsData.Cells(i, 8).Value & "|" & _
                      wsData.Cells(i, 9).Value
                      
            wsOut.Cells(dictSteps(stepKey), _
                        dictMarket(wsData.Cells(i, 6).Value)).Value = _
                        wsData.Cells(i, 10).Value
        End If
    Next i
    
    '-----------------------------------
    ' 6?? Add Average Column
    '-----------------------------------
    lastMarketCol = wsOut.Cells(1, wsOut.Columns.Count).End(xlToLeft).Column
    wsOut.Cells(1, lastMarketCol + 1).Value = "Average"
    
    lastOutputRow = wsOut.Cells(wsOut.Rows.Count, 1).End(xlUp).Row
    
    For r = 2 To lastOutputRow
        wsOut.Cells(r, lastMarketCol + 1).Formula = _
            "=IFERROR(AVERAGEIF(" & _
            wsOut.Cells(r, 8).Address(False, False) & ":" & _
            wsOut.Cells(r, lastMarketCol).Address(False, False) & _
            ","">0""),"""")"
    Next r
    
    '-----------------------------------
    ' 7?? Conditional Formatting
    '-----------------------------------
    firstDataCol = 8
    avgCol = lastMarketCol + 1
    
    Set dataRange = wsOut.Range(wsOut.Cells(2, firstDataCol), _
                                wsOut.Cells(lastOutputRow, lastMarketCol))
    
    dataRange.FormatConditions.Delete
    
    dataRange.FormatConditions.Add Type:=xlExpression, _
        Formula1:="=LEN(H2)=0"
    dataRange.FormatConditions(1).Interior.Color = RGB(255, 0, 0)
    
    dataRange.FormatConditions.Add Type:=xlExpression, _
        Formula1:="=AND(H2<>"""",ABS(H2-$" & _
        Split(wsOut.Cells(1, avgCol).Address, "$")(1) & _
        "2)>0.1*$" & _
        Split(wsOut.Cells(1, avgCol).Address, "$")(1) & "2)"
    dataRange.FormatConditions(2).Interior.Color = RGB(255, 255, 0)
    
    wsOut.Rows(1).Font.Bold = True
    wsOut.Columns.AutoFit
    
    MsgBox "Report Refreshed Without Affecting Other Data!", vbInformation

End Sub

