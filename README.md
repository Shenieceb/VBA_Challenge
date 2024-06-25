Background
You are well on your way to becoming a programmer and Excel expert! In this homework assignment, you will use VBA scripting to analyze generated stock market data.

Before You Begin
Create a new repository for this project called VBA-challenge. Do not add this assignment to an existing repository.

Inside the new repository that you just created, add any VBA files that you use for this assignment. These will be the main scripts to run for each analysis.

The code below was used for VBA scripting


Sub QuarterlyStockAnalysis()

    ' Loop through all worksheets in the workbook
    For Each ws In ThisWorkbook.Worksheets
                    
        ' Initialize variables
        Dim Ticker As String
        Dim QuarterlyChange As Double
        Dim PercentChange As Double
        Dim TotalVolume As Double
        Dim OpeningPrice As Double
        Dim ClosingPrice As Double
        Dim LastRow As Long
        Dim SummaryRow As Long

        ' Create column headers in summary table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        ' Get the last row of data in the worksheet
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Set first row of data in summary table
        SummaryRow = 2
        
        ' Loop through all rows of data
        For i = 2 To LastRow
        
            ' If it is the first time the ticker symbol appears
            If i = 2 Or ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                ' Get the opening price
                OpeningPrice = ws.Cells(i, 3).Value
                ' Restart the total volume
                TotalVolume = 0
            End If
            
            ' Add to the total volume
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value

            ' If it is the last time the ticker symbol appears
            If i = LastRow Or ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then

                ' Get the ticker symbol
                Ticker = ws.Cells(i, 1).Value

                ' Get the closing price
                ClosingPrice = ws.Cells(i, 6).Value

                ' Calculate the quarterly change
                QuarterlyChange = ClosingPrice - OpeningPrice

                ' Calculate the percent change
                If OpeningPrice <> 0 Then
                    PercentChange = (QuarterlyChange / OpeningPrice)
                Else
                    PercentChange = 0
                End If

                ' Add the ticker, quarterly change, percent change, and total volume to the summary table
                ws.Cells(SummaryRow, 9).Value = Ticker
                ws.Cells(SummaryRow, 10).Value = QuarterlyChange
                ws.Cells(SummaryRow, 11).Value = PercentChange
                ws.Cells(SummaryRow, 12).Value = TotalVolume

                ' Format the values in the summary table
                ws.Cells(SummaryRow, 10).NumberFormat = "0.00"
                ws.Cells(SummaryRow, 11).NumberFormat = "0.00%"

                ' Set conditional formatting for positive and negative quarterly changes
                If QuarterlyChange > 0 Then
                    ws.Cells(SummaryRow, 10).Interior.Color = RGB(0, 255, 0)
                ElseIf QuarterlyChange < 0 Then
                    ws.Cells(SummaryRow, 10).Interior.Color = RGB(255, 0, 0)
                End If

                ' Increment the summary table row
                SummaryRow = SummaryRow + 1
                
             End If
       
        ' Go to next row
        Next i

        ' ------------END SUMMARY TABLE------------
        
        ' ------------START MAX % INCREASE/DECREASE & MAX VOLUME TABLE ------------
        ' ------------Stock with the greatest percent increase, decrease, volume------------

        ' Set initial variables
        Dim MaxPercentIncrease As Double
        Dim MaxPercentDecrease As Double
        Dim MaxVolume As Double
        Dim MaxPercentIncreaseTicker As String
        Dim MaxPercentDecreasTicker As String
        Dim MaxVolumeTicker As String
    
        MaxPercentIncrease = 0
        MaxPercentDecrease = 0
        MaxVolume = 0
    
        ' Get the last row of data in the summary table
        SummaryLastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
        ' Loop through all rows of data in summary table
        For i = 2 To SummaryLastRow
        
            ' Find greatest percent increase
            If ws.Cells(i, 11).Value > MaxPercentIncrease Then
                MaxPercentIncrease = ws.Cells(i, 11).Value
                MaxPercentIncreaseTicker = ws.Cells(i, 9).Value
            End If
        
            ' Find greatest percent decrease
            If ws.Cells(i, 11).Value < MaxPercentDecrease Then
                MaxPercentDecrease = ws.Cells(i, 11).Value
                MaxPercentDecreaseTicker = ws.Cells(i, 9).Value
            End If
        
            ' Find greatest volume
            If ws.Cells(i, 12).Value > MaxVolume Then
                MaxVolume = ws.Cells(i, 12).Value
                MaxVolumeTicker = ws.Cells(i, 9).Value
            End If
        
        ' Go to next row
        Next i
    
        ' Create the summary table
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(2, 16).Value = MaxPercentIncreaseTicker
        ws.Cells(2, 17).Value = MaxPercentIncrease
        
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(3, 16).Value = MaxPercentDecreaseTicker
        ws.Cells(3, 17).Value = MaxPercentDecrease
        
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(4, 16).Value = MaxVolumeTicker
        ws.Cells(4, 17).Value = MaxVolume
    
        ' Format the values in the summary table
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 17).NumberFormat = "0.00%"
        ws.Cells(4, 17).NumberFormat = "0.00E+0"
        
        ' ------------END MAX % INCREASE/DECREASE & MAX VOLUME TABLE ------------
        
        ' Autofit column width for all created columns
        ws.Columns("I:Q").AutoFit
    
    ' go to next worksheet and repeat
    Next ws

End Sub


