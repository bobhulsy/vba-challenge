Attribute VB_Name = "Module2"
Sub AnalyzeMetrics()

    ' Declare variables for the current worksheet, maximum increase/decrease, maximum volume, and row counters
    Dim targetSheet As Worksheet
    Dim maxIncrease As Double
    Dim minIncrease As Double
    Dim maxVol As Double
    Dim finalRow As Long
    Dim i As Long
    
    Dim maxIncreaseTicker As String
    Dim minIncreaseTicker As String
    Dim maxVolTicker As String

    ' Loop through each worksheet in the workbook
    For Each targetSheet In ThisWorkbook.Worksheets
        
        ' Initialize metrics for each worksheet
        maxIncrease = 0
        minIncrease = 1E+30  ' Start with a very high number for minimum detection
        maxVol = 0
        
        ' Get the last row of the summary table
        finalRow = targetSheet.Cells(targetSheet.Rows.Count, 9).End(xlUp).Row
        
        ' Loop through each row of the summary table to find maximums and minimums
        For i = 2 To finalRow
            With targetSheet
                ' Detect greatest percent increase
                If .Cells(i, 11).Value > maxIncrease Then
                    maxIncrease = .Cells(i, 11).Value
                    maxIncreaseTicker = .Cells(i, 9).Value
                End If
                
                ' Detect greatest percent decrease
                If .Cells(i, 11).Value < minIncrease Then
                    minIncrease = .Cells(i, 11).Value
                    minIncreaseTicker = .Cells(i, 9).Value
                End If
                
                ' Detect greatest total volume
                If .Cells(i, 12).Value > maxVol Then
                    maxVol = .Cells(i, 12).Value
                    maxVolTicker = .Cells(i, 9).Value
                End If
            End With
        Next i

        ' Output results
        With targetSheet
            ' Set headers
            .Cells(1, 15).Value = "Metric"
            .Cells(1, 16).Value = "Ticker"
            .Cells(1, 17).Value = "Value"

            ' Greatest % Increase
            .Cells(2, 15).Value = "Greatest % Increase"
            .Cells(2, 16).Value = maxIncreaseTicker
            .Cells(2, 17).Value = Format(maxIncrease, "Percent")

            ' Greatest % Decrease
            .Cells(3, 15).Value = "Greatest % Decrease"
            .Cells(3, 16).Value = minIncreaseTicker
            .Cells(3, 17).Value = Format(minIncrease, "Percent")

            ' Greatest Total Volume
            .Cells(4, 15).Value = "Greatest Total Volume"
            .Cells(4, 16).Value = maxVolTicker
            .Cells(4, 17).Value = maxVol
        End With

    Next targetSheet

End Sub

