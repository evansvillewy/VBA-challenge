Sub forEachWs()

    Dim ws As Worksheet
    Dim Rows As Double
    Dim Columns As Double

    'interate through each worksheet in the workbook
    For Each ws In ActiveWorkbook.Worksheets

        Rows = 0
        Columns = 0

        'Get the number of rows and columns
        Call getRange(ws, Rows, Columns)

        'Sort the data
        Call sortData(ws, Rows, Columns)

        'Get the stock totals and exceptions
        Call getTotals(ws, Rows)
    Next
    
    'Activate the first sheet of the workbook
    ActiveWorkbook.Worksheets(1).Activate

End Sub

Sub getRange(ws, Rows, Columns)

   Rows = ws.UsedRange.Rows.Count
   Columns = ws.UsedRange.Columns.Count

End Sub

Sub sortData(ws, Rows, Columns)
    Dim startRange As String
    Dim endRange As String
    
    'identify the start/end range and sort keys
    With ws
        startRange = .Cells(1, 1).Address(RowAbsolute:=False, ColumnAbsolute:=False)
        sortKey1 = "A1"
        sortKey2 = "B1"
        endRange = .Cells(Rows, Columns).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    End With
    
    'sort the data in ascending order
    With ws.Sort
        .SortFields.Add Key:=Range(sortKey1), Order:=xlAscending
        .SortFields.Add Key:=Range(sortKey2), Order:=xlAscending
        .SetRange Range(startRange + ":" + endRange)
        .Header = xlYes
        .Apply
    End With
End Sub

Sub getTotals(ws, Rows)
    Dim total As Double
    Dim stockName As String
    Dim nextstockName As String
    Dim openData As String
    Dim stockNameOut As String
    Dim stockTotalOut As String
    Dim volumeData As String
    Dim closeData As String
    Dim yearlyChange As Double
    Dim changePCT As Double
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim greatestPCTInc As Double
    Dim greatestPCTDec As Double
    Dim greatestTotal As Double
    Dim greatestPCTIncName As String
    Dim greatestPCTDecName As String
    Dim greatestTotalName As String
    Dim totalCounter As Double
    
    'initialize variables
    openData = "C"
    closeData = "F"
    stockNameOut = "I"
    yearlyChangeOut = "J"
    changePCTOut = "K"
    stockTotalOut = "L"
    volumeData = "G"
    
    stockName = ""
    nextstockName = ""
    total = 0
    totalCounter = 1
    yearlyChange = 0
    changePCT = 0
    greatestPCTInc = 0
    greatestPCTDec = 0
    greatestTotal = 0
    greatestPCTIncName = ""
    greatestPCTDecName = ""
    greatestTotalName = ""
    
    
    With ws
        'output headers
        .Range(Replace(stockNameOut + Str(totalCounter), " ", "")).Value = "Ticker"
        .Range(Replace(yearlyChangeOut + Str(totalCounter), " ", "")).Value = "Yearly Change"
        .Range(Replace(changePCTOut + Str(totalCounter), " ", "")).Value = "Percent Change"
        .Range(Replace(stockTotalOut + Str(totalCounter), " ", "")).Value = "Total Stock Volume"
        
        'get the opening price
        openingPrice = .Range(Replace(openData + Str(2), " ", "")).Value
        
        'interate through all the ticker date/price records
        For i = 2 To Rows
            'only perform the operations if there is anopening price
            If openingPrice <> 0 Then
                stockName = .Cells(i, 1).Value
                nextstockName = .Cells(i + 1, 1).Value
                
                'if the stock is changing capture totals per stock
                If stockName <> nextstockName Then
                
                    'increment the stock total and atock counter
                    totalCounter = totalCounter + 1
                    total = total + .Range(Replace(volumeData + Str(i), " ", "")).Value
                
                    'capture the stock closing price
                    closingPrice = .Range(Replace(closeData + Str(i), " ", "")).Value
                    
                    'calculate the change over the year and % of the change
                    yearlyChange = closingPrice - openingPrice
                    changePCT = yearlyChange / openingPrice
            
                    'capture the next stokl openeing price
                    openingPrice = .Range(Replace(openData + Str(i + 1), " ", "")).Value
    
                    'output the totals per stock
                    .Range(Replace(stockNameOut + Str(totalCounter), " ", "")).Value = stockName
                    .Range(Replace(yearlyChangeOut + Str(totalCounter), " ", "")).Value = yearlyChange
                    .Range(Replace(changePCTOut + Str(totalCounter), " ", "")).Value = changePCT
                    .Range(Replace(changePCTOut + Str(totalCounter), " ", "")).NumberFormat = "0.00%"
                    .Range(Replace(stockTotalOut + Str(totalCounter), " ", "")).Value = total
                    
                    'format the yearly change based on its value
                    If yearlyChange > 0 Then
                       .Range(Replace(yearlyChangeOut + Str(totalCounter), " ", "")).Interior.ColorIndex = 4
                    ElseIf yearlyChange < 0 Then
                        .Range(Replace(yearlyChangeOut + Str(totalCounter), " ", "")).Interior.ColorIndex = 3
                    End If
                    
                    'capture greatest/least increase, decrease, total per stock
                    If changePCT > 0 Then
                        If greatestPCTInc < changePCT Then
                            greatestPCTInc = changePCT
                            greatestPCTIncName = stockName
                        End If
                    Else
                        If greatestPCTDec < Abs(changePCT) Then
                            greatestPCTDec = Abs(changePCT)
                            greatestPCTDecName = stockName
                        End If
                    End If
                    
                    If greatestTotal < total Then
                        greatestTotal = total
                        greatestTotalName = stockName
                    End If
                    
                    'reset the stock total
                    total = 0
                Else
                    total = total + .Range(Replace(volumeData + Str(i), " ", "")).Value
                End If
            Else
                'capture the next opening price
                openingPrice = .Range(Replace(openData + Str(i + 1), " ", "")).Value
            End If
           
        Next i
        
        'Output the headers and greatest increase, decrease and total stock for the year
        .Range("O1").Value = "Ticker"
        .Range("P1").Value = "Value"
        .Range("N2").Value = "Greatest % Increase"
        .Range("N3").Value = "Greatest % Decrease"
        .Range("N4").Value = "Greatest Total Volume"
        .Range("O2").Value = greatestPCTIncName
        .Range("O3").Value = greatestPCTDecName
        .Range("O4").Value = greatestTotalName
        .Range("P2").Value = greatestPCTInc
        .Range("P2").NumberFormat = "0.00%"
        .Range("P3").Value = greatestPCTDec * -1
        .Range("P3").NumberFormat = "0.00%"
        .Range("P4").Value = greatestTotal
        
        'AutoFit All Columns on Worksheet
        .Cells.EntireColumn.AutoFit
        
        'Activate/scroll to the first cell for each sheet
        .Activate
        'Select cell A1 in active worksheet
        .Range("A1").Select
        'Zoom to first cell
        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 1

    End With
    
End Sub

