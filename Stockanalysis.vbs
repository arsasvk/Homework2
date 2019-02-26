Sub stockanalysis()

Dim ws As Worksheet
Dim rowCount As Long
Dim i As Long
Dim tick As String
Dim mintick As String
Dim voltick As String
Dim j As Long
Dim stockVolume As LongLong
Dim openPrice As Double
Dim closePrice As Double
Dim priceDiff As Double
Dim maxPercentIncr As Double
Dim minPercentDec As Double
Dim maxVolume As LongLong

'Iterate through worksheets, one per year

For Each ws In Worksheets
    ws.Activate
    ws.Range("I1:Q80000").ClearFormats
    ws.Range("I1:Q80000").ClearContents

    'Add summary columns
    Columns("I:J").Insert Shift:=xlToRight
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Total Stock Volume"
    Range("K1").Value = "Yearly Change"
    Range("L1").Value = "Percent Change"
    'Iterate through ticker symbols
    rowCount = Cells(ws.Rows.Count, 1).End(xlUp).Row
    tick = Cells(2, 1).Value
    openPrice = Cells(2, 3).Value
    j = 2
    stockVolume = 0
    For i = 2 To rowCount
        'Calculate total stock volume and assign a ticker accordingly
        If tick <> Cells(i, 1).Value Then
            Cells(j, 9).Value = tick
            Cells(j, 10).Value = stockVolume
            closePrice = Cells(i - 1, 6).Value
            
            'Calculate yearly change
            Cells(j, 11).NumberFormat = "0.000000000"
            priceDiff = closePrice - openPrice
            Cells(j, 11).Value = priceDiff
            If priceDiff > 0 Then
                Cells(j, 11).Interior.ColorIndex = 4
            Else
                Cells(j, 11).Interior.ColorIndex = 3
            End If
            
            'Calculate percent change
            Cells(j, 12).NumberFormat = "0.00%"
            If openPrice = 0 Then
                openPrice = closePrice
            Else
                Cells(j, 12).Value = (priceDiff) / openPrice
            End If
            
            openPrice = Cells(i, 3).Value
            
            stockVolume = 0
            j = j + 1
            tick = Cells(i, 1).Value
        End If
        stockVolume = stockVolume + Cells(i, 7).Value
        Next i
    
    ' Locating the stock with the Greatest % increase, Greatest % Decrease and Greatest total volume
    
    maxPercentIncr = 0
    minPercentDec = 0
    maxVolume = 0
    
    rowCount = j - 1
    tick = Cells(2, 9).Value
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 17).NumberFormat = "0.00%"
    Cells(3, 17).NumberFormat = "0.00%"
    Cells(2, 15).Value = "Max % Increase"
    Cells(3, 15).Value = "Max % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    
    For i = 2 To rowCount
        'Calculating Greatest % increase
        If maxPercentIncr < Cells(i, 12).Value Then
            maxPercentIncr = Cells(i, 12).Value
            tick = Cells(i, 9).Value
        End If
        'Calculating Greatest % decrease
        If minPercentDec > Cells(i, 12).Value Then
            minPercentDec = Cells(i, 12).Value
            mintick = Cells(i, 9).Value
        End If
        'Calculating Greatest total volume
        If maxVolume < Cells(i, 10).Value Then
            maxVolume = Cells(i, 10).Value
            voltick = Cells(i, 9).Value
        End If
    
        Next i
    'Assigning calculated values
    Cells(2, 16).Value = tick
    Cells(2, 17).Value = maxPercentIncr
    Cells(3, 16).Value = mintick
    Cells(3, 17).Value = minPercentDec
    Cells(4, 16).Value = voltick
    Cells(4, 17).Value = maxVolume
    Next
End Sub

