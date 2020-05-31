Attribute VB_Name = "Module1"
Sub StockData()

Dim ticker As String
Dim yearopen As Double
Dim yearclose As Double
Dim yearchange As Double
Dim percentage As Double
Dim vol As Double
Dim summary_table_row As Integer
Dim increament As Integer


For Each ws In Worksheets
    
    vol = 0
    increament = 0
    summary_table_row = 2
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
            
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'ws.Range("A2:G" & lastrow).Sort Key1:=Range("A2:A" & lastrow), Order1:=xlAscending, Header:=xlNo
    'ws.Range("A2:G" & lastrow).Sort Key2:=Range("B2:B" & lastrow), Order2:=xlAscending, Header:=xlNo
    
    For i = 2 To lastrow
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            ticker = ws.Cells(i, 1).Value
            vol = vol + ws.Cells(i, 7).Value
            yearopen = ws.Cells(i - increament, 3).Value
            yearclose = ws.Cells(i, 6).Value
            yearchange = yearclose - yearopen
            
            If yearclose > 0 Then
                percentage = yearchange / yearclose
            Else
                percentage = 0
            End If
            
            ws.Cells(summary_table_row, 9).Value = ticker
            ws.Cells(summary_table_row, 10).Value = yearchange
            ws.Cells(summary_table_row, 11).Value = percentage
            ws.Cells(summary_table_row, 12).Value = vol

            If ws.Cells(summary_table_row, 10) > 0 Then
                ws.Cells(summary_table_row, 10).Interior.ColorIndex = 4
            ElseIf ws.Cells(summary_table_row, 10) < 0 Then
                ws.Cells(summary_table_row, 10).Interior.ColorIndex = 3
            End If

            summary_table_row = summary_table_row + 1
            vol = 0
            increament = 0
        
        
        Else
            increament = increament + 1
            vol = vol + ws.Cells(i, 7).Value
            
        End If

    Next i
    
        
    ws.Columns("K:K").NumberFormat = "0.00%"
    
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greater % Increase"
    ws.Cells(3, 15).Value = "Greater % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
    max_inc = WorksheetFunction.Max(ws.Range("K:K"))
    max_incS = WorksheetFunction.Match(max_inc, ws.Range("K:K"), 0)
    ws.Cells(2, 16).Value = ws.Cells(max_incS + 1, 9)
    ws.Cells(2, 17).Value = max_inc

    max_dec = WorksheetFunction.Min(ws.Range("K:K"))
    max_decS = WorksheetFunction.Match(max_dec, ws.Range("K:K"), 0)
    ws.Cells(3, 16).Value = ws.Cells(max_decS + 1, 9)
    ws.Cells(3, 17).Value = max_dec
        
    max_vol = WorksheetFunction.Max(ws.Range("L:L"))
    max_volS = WorksheetFunction.Match(max_vol, ws.Range("L:L"), 0)
    ws.Cells(4, 16).Value = ws.Cells(max_volS + 1, 9)
    ws.Cells(4, 17).Value = max_vol
        
    ws.Range("Q2:Q3").NumberFormat = "0.00%"
    ws.Columns("I:Q").AutoFit

Next ws

End Sub
