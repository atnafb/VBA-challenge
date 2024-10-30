Attribute VB_Name = "Module1"
Sub QuarterlyStockAnalysis()

    ' Declaring variables
    Dim totalVolume As Double
    Dim row As Long
    Dim rowCount As Long
    Dim quarterlyChange As Double
    Dim PercentChange As Double
    Dim Stat_Table As Long
    Dim stockStartRow As Long
    Dim startValue As Long
    Dim lastTicker As String
    Dim findValue As Long
    Dim Column As Long
    
Dim ws As Worksheet
 ' Perform the same actions on each worksheet
For Each ws In Worksheets
   
            ' Set the headers for Ticker, Quarterly Change, Percent Change, and Total Stock Volume
            ws.Range("I1").Value = "Ticker"
            ws.Range("J1").Value = "Quarterly Change"
            ws.Range("K1").Value = "Percent Change"
            ws.Range("L1").Value = "Total Stock Volume"
        
            ' Set up the title row of the Aggregate Section
            ws.Range("P1").Value = "Ticker"
            ws.Range("Q1").Value = "Value"
            ws.Range("O2").Value = "Greatest % Increase"
            ws.Range("O3").Value = "Greatest % Decrease"
            ws.Range("O4").Value = "Greatest Total Volume"
        
            ' Initialize values
            totalVolume = 0
            stockStartRow = 2
            Stat_Table = 0
            startValue = 2
            quarterlyChange = 0
        
            ' Get the last row in the sheet
            rowCount = ws.Cells(Rows.Count, "A").End(xlUp).row
            lastTicker = ws.Cells(rowCount, 1).Value ' Find the last ticker to that we can break out of the loop
        
            ' Loop until the end of the sheet
            For row = 2 To rowCount
                ' Check if the current row is the first entry of a new quarter or the last entry of the current quarter
                If ws.Cells(row + 1, 1).Value <> ws.Cells(row, 1).Value Then
                    totalVolume = totalVolume + ws.Cells(row, 7).Value
                    If totalVolume = 0 Then ' Print the result in the Stat table from column I - L
                        ws.Range("I" & 2 + Stat_Table).Value = ws.Cells(row, 1).Value
                        ws.Range("J" & 2 + Stat_Table).Value = 0
                        ws.Range("K" & 2 + Stat_Table).Value = 0
                        ws.Range("L" & 2 + Stat_Table).Value = totalVolume
                    Else
                        If ws.Cells(startValue, 3) = 0 Then ' Find the first non-0 value for the stock
                            For findValue = startValue To row
                                If ws.Cells(findValue, 3).Value <> 0 Then
                                    startValue = findValue
                                    Exit For ' Break the loop
                                End If
                            Next findValue
                        End If
        
                        ' Calculate quarterly change and percent change
                        quarterlyChange = ws.Cells(row, 6).Value - ws.Cells(startValue, 3).Value
                        PercentChange = quarterlyChange / ws.Cells(startValue, 3).Value
        
                        ' Print results in the Stat table
                        ws.Range("I" & 2 + Stat_Table).Value = ws.Cells(row, 1).Value
                        ws.Range("J" & 2 + Stat_Table).Value = quarterlyChange
                        ws.Range("K" & 2 + Stat_Table).Value = PercentChange
                        ws.Range("L" & 2 + Stat_Table).Value = totalVolume
        
                        ' Formatting for Quarterly Change
                        If quarterlyChange > 0 Then
                            ws.Range("J" & 2 + Stat_Table).Interior.ColorIndex = 4
                        ElseIf quarterlyChange < 0 Then
                            ws.Range("J" & 2 + Stat_Table).Interior.ColorIndex = 3
                        Else
                            ws.Range("J" & 2 + Stat_Table).Interior.ColorIndex = 0
                        End If
        
                        ' Reset values for the next ticker
                        totalVolume = 0
                        quarterlyChange = 0
                        startValue = row + 1
                        Stat_Table = Stat_Table + 1
                    End If
                Else
                    totalVolume = totalVolume + ws.Cells(row, 7).Value
                End If
            Next row
        
            ' Clean up to avoid extra data in the stat section
            Stat_Table = ws.Cells(Rows.Count, "I").End(xlUp).row
        
            ' Clear data in the extra rows from Quarterly Change, Percent Change, and Total Stock columns
            Dim lastExtraRow As Long
            lastExtraRow = Cells(Rows.Count, "J").End(xlUp).row
        
            For j = Stat_Table To lastExtraRow
                For Column = 9 To 12
                    ws.Cells(j, Column).Value = ""
                    ws.Cells(j, Column).Interior.ColorIndex = 0
                Next Column
            Next j
        
            ' Display the Greatest % Increase, Greatest % Decrease, and Greatest Total Volume in the summary table
            ws.Range("Q2").Value = WorksheetFunction.Max(ws.Range("K2:K" & Stat_Table + 2))
            ws.Range("Q3").Value = WorksheetFunction.Min(ws.Range("K2:K" & Stat_Table + 2))
            ws.Range("Q4").Value = WorksheetFunction.Max(ws.Range("L2:L" & Stat_Table + 2))
        
            ' Find rows for Greatest % Increase, Greatest % Decrease, and Greatest Total Volume
            Dim greatestIncreaseRow As Long
            Dim greatestDecreaseRow As Long
            Dim greatestTotalVolRow As Long
        
            ' Use Match() to find row numbers
            greatestIncreaseRow = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & Stat_Table + 2)), ws.Range("K2:K" & Stat_Table + 2), 0)
            greatestDecreaseRow = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & Stat_Table + 2)), ws.Range("K2:K" & Stat_Table + 2), 0)
            greatestTotalVolRow = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & Stat_Table + 2)), ws.Range("L2:L" & Stat_Table + 2), 0)
        
            ' Display ticker symbols for Greatest % Increase, Greatest % Decrease, and Greatest Total Volume
            ws.Range("P2").Value = ws.Cells(greatestIncreaseRow + 1, 9).Value
            ws.Range("P3").Value = ws.Cells(greatestDecreaseRow + 1, 9).Value
            ws.Range("P4").Value = ws.Cells(greatestTotalVolRow + 1, 9).Value
        
        
                    ' Formatinging the Stat Table
            For s = 0 To Stat_Table
            ws.Range("J" & 2 + s).NumberFormat = "0.00" 'to display two decimal places on the quartely change column
            ws.Range("K" & 2 + s).NumberFormat = "0.00%" ' to display a percentage with two decimal placeson percentchange column
            ws.Range("L" & 2 + s).NumberFormat = "#,###" ' to include a thousands separator on thetotal stock volume column
            Next s
            'Format the state aggregates
            ws.Range("Q2").NumberFormat = "0.00%"  ' format the greatest % increase
            ws.Range("Q3").NumberFormat = "0.00%"
            ws.Range("Q4").NumberFormat = "#,###"
                    
            ' Autofit columns for better readability
            ws.Columns("A:Q").AutoFit
Next ws
End Sub
