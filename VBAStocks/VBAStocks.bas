Attribute VB_Name = "Module1"
Sub VBAStocks()

For Each ws In Worksheets 'looping through each worksheet

    Dim count As Integer 'count for ticker
    Dim lastrow As Long 'find last row of ticker in column A
    Dim open_price As Double
    Dim closing_price As Double
    Dim count2 As Integer 'count2 used to determine open_price
    Dim dif As Double 'dif is yearly changed
    Dim percent_change As Double
    Dim total As Double 'total volume
    Dim lastrow2 As Long 'find last row in new ticker list generated in column I
    Dim greatest_percent_increase As Double
    Dim greatest_persent_decrease As Double
    Dim greatest_volume As Double
    Dim ticker_increase As String 'ticker name of greatest increase%
    Dim ticker_decrease As String 'ticker name of greatest decrease%
    Dim ticker_volume As String 'ticker name of greatest total volume

    count = 1
    count2 = 0

    lastrow = ws.Cells(Rows.count, "A").End(xlUp).Row
    
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("k1") = "Percent Change"
    ws.Range("L1") = "Total Volume"

    'loops through current worksheet to populate columns I to L with ticker, yearly changed, % changed, and total volume
    For i = 2 To lastrow
        count2 = count2 + 1
        total = total + Cells(i, 7) 'add up total volume
        If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then 'Separates ticker symbols from each other
    
            count = count + 1
            ws.Cells(count, 9) = ws.Cells(i, 1)
            open_price = ws.Cells((i + 1) - count2, 3)
            closing_price = ws.Cells(i, 6)
            dif = closing_price - open_price
            ws.Cells(count, 10) = dif
        
            If dif >= 0 Then 'Formats cells in column J with colors with green or red
                ws.Cells(count, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(count, 10).Interior.ColorIndex = 3
            End If
        
            If open_price <> 0 Then 'prevents from dividing by 0
                percent_change = dif / open_price
            End If
        
            ws.Cells(count, 11) = percent_change
            ws.Cells(count, 11).NumberFormat = "0.00%" 'format cells in percent changed as a percent
            count2 = 0
            ws.Cells(count, 12) = total
            total = 0 'resets volume to 0 for next ticker
        End If

    Next i

    ws.Range("O2") = "Greatest % Increase"
    ws.Range("O3") = "Greatest % Decrease"
    ws.Range("o4") = "Greatest Volume"
    ws.Range("P1") = "Ticker"
    ws.Range("Q1") = "Value"


    lastrow2 = ws.Cells(Rows.count, "K").End(xlUp).Row
    greatest_percent_increase = ws.Cells(2, 11)
    greatest_percent_decrease = ws.Cells(2, 11)
    greatest_volume = ws.Cells(2, 12)


    'loop to find greatest percent increase, greatest percent decrease, and greatest total volume
    For j = 3 To lastrow2
    
        If greatest_percent_increase < ws.Cells(j, 11) Then
        
            greatest_percent_increase = ws.Cells(j, 11)
            ticker_increase = ws.Cells(j, 9)
       
        End If
    
        If greatest_percent_decrease > ws.Cells(j, 11) Then
        
            greatest_percent_decrease = ws.Cells(j, 11)
            ticker_decrease = ws.Cells(j, 9)
       
        End If
    
        If greatest_volume < ws.Cells(j, 12) Then
        
            greatest_volume = ws.Cells(j, 12)
            ticker_volume = ws.Cells(j, 9)
       
        End If
    Next j
    ws.Range("P2") = ticker_increase
    ws.Range("P3") = ticker_decrease
    ws.Range("P4") = ticker_volume
    ws.Range("Q2") = greatest_percent_increase
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("Q3") = greatest_percent_decrease
    ws.Range("Q3").NumberFormat = "0.00%"
    ws.Range("Q4") = greatest_volume

Next ws

End Sub
