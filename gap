Sub Create_Market_Gap_Analysis()

    Dim wsData As Worksheet
    Dim wsOut As Worksheet
    Dim lastRow As Long
    Dim dictMarket As Object
    Dim dictSteps As Object
    Dim i As Long, r As Long
    Dim stepKey As String
    Dim marketKey As String
    Dim key As Variant
    Dim lastMarketCol As Long
    Dim lastOutputRow As Long
    Dim dataRange As Range
    Dim avgCol As Long
    Dim lastUsedRow As Long
    Dim lastUsedCol As Long
    Dim firstDataCol As Long
    
    Set wsData = Sheets("Raw Data")
    Set wsOut = Sheets("Output")
    
    lastUsedRow = wsOut.Cells(wsOut.Rows.Count, 1).End(xlUp).Row
    lastUsedCol = wsOut.Cells(1, wsOut.Columns.Count).End(xlToLeft).Column
    
    If lastUsedRow > 1 Then
        wsOut.Range(wsOut.Cells(1, 1), wsOut.Cells(lastUsedRow, lastUsedCol)).Clear
        wsOut.Range(wsOut.Cells(1, 1), wsOut.Cells(lastUsedRow, lastUsedCol)).FormatConditions.Delete
    End If
    
    Set dictMarket = CreateObject("Scripting.Dictionary")
    Set dictSteps = CreateObject("Scripting.Dictionary")
    
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row

'====================================================
' 1?? UNIQUE MARKET_ENTITY_VARIANT (F + E + G)
'====================================================
    
    For i = 2 To lastRow
        
        If LCase(Trim(wsData.Cells(i, 11).Value)) = "yes" Then
            
            marketKey = Trim(wsData.Cells(i, 6).Value) & "_" & _
                        Trim(wsData.Cells(i, 5).Value) & "_" & _
                        Trim(wsData.Cells(i, 7).Value)
            
            If Not dictMarket.exists(marketKey) Then
                dictMarket.Add marketKey, dictMarket.Count + 6
            End If
            
        End If
        
    Next i

'====================================================
' 2?? UNIQUE STEP ROW (A,B,C,H,I)
'====================================================
    
    For i = 2 To lastRow
        
        If LCase(Trim(wsData.Cells(i, 11).Value)) = "yes" Then
            
            stepKey = Trim(wsData.Cells(i, 1).Value) & "|" & _
                      Trim(wsData.Cells(i, 2).Value) & "|" & _
                      Trim(wsData.Cells(i, 3).Value) & "|" & _
                      Trim(wsData.Cells(i, 8).Value) & "|" & _
                      Trim(wsData.Cells(i, 9).Value)
            
            If Not dictSteps.exists(stepKey) Then
                dictSteps.Add stepKey, dictSteps.Count + 2
            End If
            
        End If
        
    Next i

'====================================================
' 3?? WRITE HEADERS
'====================================================
    
    wsOut.Cells(1, 1).Value = "Tower"
    wsOut.Cells(1, 2).Value = "L2"
    wsOut.Cells(1, 3).Value = "L3"
    wsOut.Cells(1, 4).Value = "Broad Steps"
    wsOut.Cells(1, 5).Value = "Steps Name"
    
    For Each key In dictMarket.Keys
        wsOut.Cells(1, dictMarket(key)).Value = key
    Next key

'====================================================
' 4?? WRITE STEP ROWS
'====================================================
    
    For Each key In dictSteps.Keys
        
        r = dictSteps(key)
        
        wsOut.Cells(r, 1).Value = Split(key, "|")(0)
        wsOut.Cells(r, 2).Value = Split(key, "|")(1)
        wsOut.Cells(r, 3).Value = Split(key, "|")(2)
        wsOut.Cells(r, 4).Value = Split(key, "|")(3)
        wsOut.Cells(r, 5).Value = Split(key, "|")(4)
        
    Next key

'====================================================
' 5?? FILL DURATION (Column J)
'====================================================
    
    For i = 2 To lastRow
        
        If LCase(Trim(wsData.Cells(i, 11).Value)) = "yes" Then
            
            stepKey = Trim(wsData.Cells(i, 1).Value) & "|" & _
                      Trim(wsData.Cells(i, 2).Value) & "|" & _
                      Trim(wsData.Cells(i, 3).Value) & "|" & _
                      Trim(wsData.Cells(i, 8).Value) & "|" & _
                      Trim(wsData.Cells(i, 9).Value)
            
            marketKey = Trim(wsData.Cells(i, 6).Value) & "_" & _
                        Trim(wsData.Cells(i, 5).Value) & "_" & _
                        Trim(wsData.Cells(i, 7).Value)
            
            If dictSteps.exists(stepKey) And dictMarket.exists(marketKey) Then
                
                wsOut.Cells(dictSteps(stepKey), _
                            dictMarket(marketKey)).Value = _
                            wsData.Cells(i, 10).Value
                            
            End If
            
        End If
        
    Next i

'====================================================
' 6?? AVERAGE (Ignore 0 & Blank)
'====================================================
    
    lastMarketCol = wsOut.Cells(1, wsOut.Columns.Count).End(xlToLeft).Column
    wsOut.Cells(1, lastMarketCol + 1).Value = "Average"
    
    lastOutputRow = wsOut.Cells(wsOut.Rows.Count, 1).End(xlUp).Row
    
    For r = 2 To lastOutputRow
        
        wsOut.Cells(r, lastMarketCol + 1).Formula = _
        "=IFERROR(AVERAGEIF(" & _
        wsOut.Cells(r, 6).Address(False, False) & ":" & _
        wsOut.Cells(r, lastMarketCol).Address(False, False) & _
        ","">0""),"""")"
        
    Next r

'====================================================
' 7?? CONDITIONAL FORMATTING
'====================================================
    
    firstDataCol = 6
    avgCol = lastMarketCol + 1
    
    Set dataRange = wsOut.Range(wsOut.Cells(2, firstDataCol), _
                                wsOut.Cells(lastOutputRow, lastMarketCol))
    
    dataRange.FormatConditions.Delete
    
    'Blank ? Red
    dataRange.FormatConditions.Add Type:=xlExpression, _
        Formula1:="=LEN(F2)=0"
    dataRange.FormatConditions(1).Interior.Color = RGB(255, 0, 0)
    
    '±10% from Average ? Yellow
    dataRange.FormatConditions.Add Type:=xlExpression, _
        Formula1:="=AND(F2<>"""",ABS(F2-$" & _
        Split(wsOut.Cells(1, avgCol).Address, "$")(1) & _
        "2)>0.1*$" & _
        Split(wsOut.Cells(1, avgCol).Address, "$")(1) & "2)"
    
    dataRange.FormatConditions(2).Interior.Color = RGB(255, 255, 0)

    wsOut.Rows(1).Font.Bold = True
    wsOut.Columns.AutoFit
    
    MsgBox "Final Correct Report Generated Successfully!", vbInformation

End Sub

