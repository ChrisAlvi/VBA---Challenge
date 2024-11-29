Sub StockDataSummary()
    Dim S As Worksheet
    Dim lastDataRow As Long, currentRow As Long, summaryRow As Long
    Dim symbol As String, openPrice As Double, closePrice As Double
    Dim totalVolume As Double, priceChange As Double, percentChange As Double
    Dim highestIncrease As Double, highestDecrease As Double, highestVolume As Double
    Dim lastSummaryRow As Long
    
    ' Loop through all sheets in the workbook
    For Each S In ThisWorkbook.Worksheets
        ' Get the last row with data
        lastDataRow = S.Cells(S.Rows.Count, 1).End(xlUp).Row
        
        ' Initialize variables for calculations
        summaryRow = 2
        totalVolume = 0
        highestIncrease = -1E+30
        highestDecrease = 1E+30
        highestVolume = 0
        
        ' Set headers for summary table
        S.Cells(1, 9).Value = "Ticker"
        S.Cells(1, 10).Value = "Quarterly Change" ' Changed header to "Quarterly Change"
        S.Cells(1, 11).Value = "Percent Change"
        S.Cells(1, 12).Value = "Total Volume"
        S.Cells(2, 14).Value = "Greatest % Increase"
        S.Cells(3, 14).Value = "Greatest % Decrease"
        S.Cells(4, 14).Value = "Greatest Total Volume"
        S.Cells(1, 15).Value = "Ticker"
        S.Cells(1, 16).Value = "Value"
        
        ' Loop through all data rows in the sheet
        For currentRow = 2 To lastDataRow
            If S.Cells(currentRow - 1, 1).Value <> S.Cells(currentRow, 1).Value Then
                ' Capture the opening price when a new symbol is encountered
                openPrice = S.Cells(currentRow, 3).Value
            ElseIf S.Cells(currentRow + 1, 1).Value <> S.Cells(currentRow, 1).Value Then
                ' Capture stock symbol and closing price at the last row of the symbol
                symbol = S.Cells(currentRow, 1).Value
                closePrice = S.Cells(currentRow, 6).Value
                
                ' Calculate quarterly change and percent change
                priceChange = closePrice - openPrice
                percentChange = priceChange / openPrice
                
                ' Add the volume for the current symbol
                totalVolume = totalVolume + S.Cells(currentRow, 7).Value
                
                ' Write the summary data to the worksheet
                S.Cells(summaryRow, 9).Value = symbol
                S.Cells(summaryRow, 10).Value = priceChange ' Display quarterly change
                S.Cells(summaryRow, 11).Value = percentChange
                S.Cells(summaryRow, 12).Value = totalVolume
                
                ' Apply percentage format to the percent change column
                S.Columns("K:K").NumberFormat = "0.00%"
                
                ' Move to the next row in the summary
                summaryRow = summaryRow + 1
                
                ' Reset total volume for next symbol
                totalVolume = 0
            Else
                ' Accumulate the total volume for the current symbol
                totalVolume = totalVolume + S.Cells(currentRow, 7).Value
            End If
        Next currentRow
        
        ' Identify the row where the summary data ends
        lastSummaryRow = S.Cells(S.Rows.Count, 10).End(xlUp).Row
        
        ' Loop through the summary rows to calculate greatest values
        For currentRow = 2 To lastSummaryRow
            ' Apply color formatting to the Quarterly Change column (Column J)
            If S.Cells(currentRow, 10).Value >= 0 Then
                S.Cells(currentRow, 10).Interior.Color = RGB(0, 255, 0) ' Green for positive
            Else
                S.Cells(currentRow, 10).Interior.Color = RGB(255, 0, 0) ' Red for negative
            End If

            ' Update the greatest percentage increase if applicable
            If S.Cells(currentRow, 11).Value > highestIncrease Then
                highestIncrease = S.Cells(currentRow, 11).Value
                S.Cells(2, 16).Value = S.Cells(currentRow, 9).Value
                S.Cells(2, 17).Value = highestIncrease
                S.Cells(2, 17).NumberFormat = "0.00%"
            End If

            ' Update the greatest percentage decrease if applicable
            If S.Cells(currentRow, 11).Value < highestDecrease Then
                highestDecrease = S.Cells(currentRow, 11).Value
                S.Cells(3, 16).Value = S.Cells(currentRow, 9).Value
                S.Cells(3, 17).Value = highestDecrease
                S.Cells(3, 17).NumberFormat = "0.00%"
            End If

            ' Update the greatest volume if applicable
            If S.Cells(currentRow, 12).Value > highestVolume Then
                highestVolume = S.Cells(currentRow, 12).Value
                S.Cells(4, 16).Value = S.Cells(currentRow, 9).Value
                S.Cells(4, 17).Value = highestVolume
            End If
        Next currentRow
    Next S
End Sub
