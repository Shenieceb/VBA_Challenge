Background
You are well on your way to becoming a programmer and Excel expert! In this homework assignment, you will use VBA scripting to analyze generated stock market data.

Before You Begin
Create a new repository for this project called VBA-challenge. Do not add this assignment to an existing repository.

Inside the new repository that you just created, add any VBA files that you use for this assignment. These will be the main scripts to run for each analysis.

The code below was used for VBA scripting
Sub Counter()
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim tableRow As Long
    Dim ticker As String
    Dim firstTicker As Long
    Dim lastTicker As Long
    Dim Total As Double
    Dim Count As Long
    Dim lastRowL As Long
    Dim lastRowM As Long
    Dim MaxL As Double
    Dim MinL As Double
    Dim MaxTotal As Double
    
    For Each ws In Worksheets
    
        
        ws.Cells(1, 10).Value = "Tickers"
        ws.Cells(1, 10).Font.Bold = True
        ws.Cells(1, 11).Value = "Yearly Change"
        ws.Cells(1, 11).Font.Bold = True
        ws.Cells(1, 12).Value = "Percent change"
        ws.Cells(1, 12).Font.Bold = True
        ws.Cells(1, 13).Value = "Total Stock Volum"
        ws.Cells(1, 13).Font.Bold = True
        
        ws.Cells(2, 16).Value = "Greatest % Increase"
        ws.Cells(2, 16).Font.Bold = True
        ws.Cells(3, 16).Value = "Greatest % Decrease"
        ws.Cells(3, 16).Font.Bold = True
        ws.Cells(4, 16).Value = "Greatest Total Volum"
        ws.Cells(4, 16).Font.Bold = True
        
        ws.Cells(1, 17).Value = "Ticker"
        ws.Cells(1, 17).Font.Bold = True
        ws.Cells(1, 18).Value = "Value"
        ws.Cells(1, 18).Font.Bold = True
        
        
        tableRow = 2
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        Total = 0

For i = 2 To lastRow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                Total = Total + ws.Range("G" & i).Value
                
                Count = 0
                For j = 1 To lastRow
                    If ws.Cells(j, 1).Value = ticker Then
                        Count = Count + 1
                    End If
                Next j
                
                firstTicker = i - Count + 1
                lastTicker = i
                
                ws.Range("J" & tableRow).Value = ticker
                ws.Range("K" & tableRow).Value = ws.Cells(lastTicker, 6).Value - ws.Cells(firstTicker, 3).Value
                ws.Range("L" & tableRow).Value = (ws.Cells(lastTicker, 6).Value - ws.Cells(firstTicker, 3).Value) / ws.Cells(firstTicker, 3).Value
                ws.Range("L" & tableRow).NumberFormat = "0.00%"
                ws.Range("M" & tableRow).Value = Total
                
                tableRow = tableRow + 1
                Total = 0
            Else
                Total = Total + ws.Range("G" & i).Value
            End If
        Next i
        For tableRow = 2 To lastRow
            
            If ws.Range("K" & tableRow).Value > 0 Then
                ws.Cells(tableRow, 11).Interior.ColorIndex = 4
            ElseIf ws.Range("K" & tableRow).Value < 0 Then
                ws.Cells(tableRow, 11).Interior.ColorIndex = 3
            End If
        Next tableRow
        
        
        lastRowL = ws.Cells(ws.Rows.Count, "L").End(xlUp).Row
        lastRowM = ws.Cells(ws.Rows.Count, "M").End(xlUp).Row
        MaxL = WorksheetFunction.Max(ws.Range("L2:L" & lastRowL))
        MinL = WorksheetFunction.Min(ws.Range("L2:L" & lastRowL))
        MaxTotal = WorksheetFunction.Max(ws.Range("M2:M" & lastRowM))
        
        For l = 2 To lastRowL
            If ws.Cells(l, 12).Value = MaxL Then
                ws.Cells(2, 18).Value = MaxL
                ws.Cells(2, 18).NumberFormat = "0.00%"
                ws.Cells(2, 17).Value = ws.Cells(l, 12).Offset(, -2).Value
            End If
        Next l
        
        For l = 2 To lastRowL
                
            If ws.Cells(l, 12).Value = MinL Then
                ws.Cells(3, 18).Value = MinL
                ws.Cells(3, 18).NumberFormat = "0.00%"
                ws.Cells(3, 17).Value = ws.Cells(l, 12).Offset(, -2).Value
                Exit For
            End If
        
        Next l
        
        For m = 2 To lastRowM
            If ws.Cells(m, 13).Value = MaxTotal Then
                ws.Cells(4, 18).Value = MaxTotal
                ws.Cells(4, 17).Value = ws.Cells(m, 13).Offset(, -3).Value
                Exit For
            End If
        Next m
    Next ws
End Sub



