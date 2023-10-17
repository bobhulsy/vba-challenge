Attribute VB_Name = "Module1"
Sub ExtractStockMetrics()

    ' Declare variables for the current worksheet, starting/ending rows of stock data, stock details, and the output row number
    Dim targetSheet As Worksheet
    Dim stockStartRow As Long
    Dim stockEndRow As Long
    Dim stockTicker As String
    Dim yearStartPrice As Double
    Dim yearEndPrice As Double
    Dim yearVol As Double
    Dim outputRow As Long
    
    ' Loop through each worksheet in the workbook
    For Each targetSheet In ThisWorkbook.Worksheets
        
        ' Initialize starting row for each new worksheet and output row for the summary table
        stockStartRow = 2
        outputRow = 2
        
        ' Set header titles for the summary table
        With targetSheet
            .Cells(1, 9).Value = "Ticker"
            .Cells(1, 10).Value = "Yearly Change"
            .Cells(1, 11).Value = "Percent Change"
            .Cells(1, 12).Value = "Total Stock Volume"
        End With

        ' Process each row until an empty cell is found in the ticker column
        Do While Not IsEmpty(targetSheet.Cells(stockStartRow, 1))
            
            ' Get the stock ticker name
            stockTicker = targetSheet.Cells(stockStartRow, 1).Value
            ' Get the opening price for the year
            yearStartPrice = targetSheet.Cells(stockStartRow, 3).Value
            ' Reset the yearly volume
            yearVol = 0
            
            ' Continue processing until the ticker changes (end of data for the current stock)
            Do
                ' Accumulate the volume for the year
                yearVol = yearVol + targetSheet.Cells(stockStartRow, 7).Value
                ' Move to the next row
                stockStartRow = stockStartRow + 1
            Loop Until targetSheet.Cells(stockStartRow, 1).Value <> stockTicker
            
            ' Get the closing price for the year for the current stock
            yearEndPrice = targetSheet.Cells(stockStartRow - 1, 6).Value
            
            ' Output stock data and apply conditional formatting
            With targetSheet
                ' Output ticker name
                .Cells(outputRow, 9).Value = stockTicker
                ' Calculate and output yearly change
                .Cells(outputRow, 10).Value = yearEndPrice - yearStartPrice
                
                ' Conditional Formatting for Yearly Change
                If .Cells(outputRow, 10).Value < 0 Then
                    ' If negative, color it Red
                    .Cells(outputRow, 10).Interior.ColorIndex = 3
                Else
                    ' If positive, color it Green
                    .Cells(outputRow, 10).Interior.ColorIndex = 4
                End If
                
                ' Calculate and output Percent Change, handling divide by zero
                If yearStartPrice <> 0 Then
                    .Cells(outputRow, 11).Value = Format((yearEndPrice - yearStartPrice) / yearStartPrice, "Percent")
                Else
                    .Cells(outputRow, 11).Value = "0%"
                End If
                ' Remove any color formatting from the Percent Change column
                .Cells(outputRow, 11).Interior.ColorIndex = -4142
                
                ' Output total volume for the year
                .Cells(outputRow, 12).Value = yearVol
            End With
            
            ' Move to the next row for the next stock's output
            outputRow = outputRow + 1
        Loop

    Next targetSheet

End Sub

