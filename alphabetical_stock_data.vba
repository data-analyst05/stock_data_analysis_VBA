Sub alphabetical_testing_StockSummaryAndFormat()
'Published by Ese Ekrebe
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String
    Dim greatestIncreasePercent As Double
    Dim greatestDecreasePercent As Double
    Dim greatestVolume As Double
    
    ' Loop through each specified sheet
    For Each ws In ThisWorkbook.Sheets(Array("A", "B", "C", "D", "F"))
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        ' Initialize variables
        ticker = ws.Cells(2, 1).Value
        openingPrice = ws.Cells(2, 3).Value
        totalVolume = 0
        
        ' Initialize variables for greatest % increase, decrease, and volume
        greatestIncreaseTicker = ""
        greatestDecreaseTicker = ""
        greatestVolumeTicker = ""
        greatestIncreasePercent = 0
        greatestDecreasePercent = 0
        greatestVolume = 0
        
        ' Loop through the rows in the sheet
        For i = 2 To lastRow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' Set closing price and calculate changes
                closingPrice = ws.Cells(i, 6).Value
                yearlyChange = closingPrice - openingPrice
                
                ' Handle potential division by zero for percent change
                If openingPrice <> 0 Then
                    percentChange = (yearlyChange / openingPrice) * 100
                Else
                    percentChange = 0
                End If
                
                ' Add to total volume
                totalVolume = totalVolume + ws.Cells(i, 7).Value
                
                ' Output results in next available row
                ws.Cells(1, 9).Value = "Ticker"
                ws.Cells(1, 10).Value = "Yearly Change"
                ws.Cells(1, 11).Value = "Percent Change"
                ws.Cells(1, 12).Value = "Total Stock Volume"
                
                ws.Cells(ws.Rows.Count, 9).End(xlUp).Offset(1, 0).Value = ticker
                ws.Cells(ws.Rows.Count, 10).End(xlUp).Offset(1, 0).Value = yearlyChange
                ws.Cells(ws.Rows.Count, 11).End(xlUp).Offset(1, 0).Value = percentChange
                ws.Cells(ws.Rows.Count, 12).End(xlUp).Offset(1, 0).Value = totalVolume
                
                ' Update greatest % increase, decrease, and volume
                If percentChange > greatestIncreasePercent Then
                    greatestIncreasePercent = percentChange
                    greatestIncreaseTicker = ticker
                End If
                If percentChange < greatestDecreasePercent Then
                    greatestDecreasePercent = percentChange
                    greatestDecreaseTicker = ticker
                End If
                If totalVolume > greatestVolume Then
                    greatestVolume = totalVolume
                    greatestVolumeTicker = ticker
                End If
                
                ' Reset variables for the next stock
                ticker = ws.Cells(i + 1, 1).Value
                openingPrice = ws.Cells(i + 1, 3).Value
                totalVolume = 0
            Else
                ' Add to total volume
                totalVolume = totalVolume + ws.Cells(i, 7).Value
            End If
        Next i
        
        ' Output greatest % increase, decrease, and volume in columns O, P, and Q
        ws.Cells(1, 15).Value = "Greatest % Increase"
        ws.Cells(2, 15).Value = "Greatest % Decrease"
        ws.Cells(3, 15).Value = "Greatest Total Volume"
        
        ws.Cells(1, 16).Value = greatestIncreaseTicker
        ws.Cells(2, 16).Value = greatestDecreaseTicker
        ws.Cells(3, 16).Value = greatestVolumeTicker
        
        ws.Cells(1, 17).Value = greatestIncreasePercent
        ws.Cells(2, 17).Value = greatestDecreasePercent
        ws.Cells(3, 17).Value = greatestVolume
        
        ' Apply conditional formatting to Yearly Change column (Column K)
        Dim cell As Range
        For Each cell In ws.Range("K2:K" & lastRow)
            If cell.Value > 0 Then
                cell.Interior.Color = RGB(0, 255, 0) ' Green
            ElseIf cell.Value < 0 Then
                cell.Interior.Color = RGB(255, 0, 0) ' Red
            Else
                cell.Interior.ColorIndex = xlNone ' No formatting for zero changes
            End If
        Next cell
    Next ws
End Sub

