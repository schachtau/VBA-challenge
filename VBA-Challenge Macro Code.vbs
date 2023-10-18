Sub StockAnalysisAllSheets()
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim Ticker As String
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalVolume As Double
    Dim SummaryRow As Long
    Dim MaxPercentIncrease As Double
    Dim MaxPercentDecrease As Double
    Dim MaxTotalVolume As Double ' Change to Double
    Dim MaxPercentIncreaseTicker As String
    Dim MaxPercentDecreaseTicker As String
    Dim MaxTotalVolumeTicker As String
    
    ' Initialize summary variables for the entire workbook
    MaxPercentIncrease = 0
    MaxPercentDecrease = 0
    MaxTotalVolume = 0
    MaxPercentIncreaseTicker = ""
    MaxPercentDecreaseTicker = ""
    MaxTotalVolumeTicker = ""
    
    ' Loop through all worksheets (years) in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Initialize variables for the current worksheet
        SummaryRow = 2
        TotalVolume = 0
        OpeningPrice = ws.Cells(2, 3).Value
        ' Find the last row in the current worksheet
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        ' Loop through the rows in the current worksheet
        For i = 2 To LastRow
            Ticker = ws.Cells(i, 1).Value
            ClosingPrice = ws.Cells(i, 6).Value
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value
            
            ' Check if the current row has a different ticker symbol
            If ws.Cells(i + 1, 1).Value <> Ticker Then
                ' Calculate the yearly change and percent change
                YearlyChange = ClosingPrice - OpeningPrice
                If OpeningPrice <> 0 Then
                    PercentChange = (YearlyChange / OpeningPrice) * 100
                Else
                    PercentChange = 0
                End If
                ' Conditional formatting for positive and negative changes
                If YearlyChange >= 0 Then
                    ws.Cells(SummaryRow, 10).Interior.Color = RGB(0, 255, 0) ' Green
                Else
                    ws.Cells(SummaryRow, 10).Interior.Color = RGB(255, 0, 0) ' Red
                End If
                ' Output data for the current year
                ws.Cells(SummaryRow, 9).Value = Ticker ' Ticker
                ws.Cells(SummaryRow, 10).Value = YearlyChange
                ws.Cells(SummaryRow, 11).Value = PercentChange & "%" ' Adding percent sign here
                ws.Cells(SummaryRow, 12).Value = TotalVolume
                
                ' Update summary variables for the current year
                If PercentChange > MaxPercentIncrease Then
                    MaxPercentIncrease = PercentChange
                    MaxPercentIncreaseTicker = Ticker
                End If
                
                If PercentChange < MaxPercentDecrease Then
                    MaxPercentDecrease = PercentChange
                    MaxPercentDecreaseTicker = Ticker
                End If
                
                If TotalVolume > MaxTotalVolume Then
                    MaxTotalVolume = TotalVolume
                    MaxTotalVolumeTicker = Ticker
                End If
                
                ' Reset variables for the next ticker
                SummaryRow = SummaryRow + 1
                TotalVolume = 0
                OpeningPrice = ws.Cells(i + 1, 3).Value ' Set opening price for the next ticker
            End If
        Next i
        
        
        
        ' Output "Greatest % Increase," "Greatest % Decrease," and "Greatest Total Volume" for the current year
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(2, 16).Value = MaxPercentIncreaseTicker
        ws.Cells(2, 17).Value = MaxPercentIncrease & "%"
        
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(3, 16).Value = MaxPercentDecreaseTicker
        ws.Cells(3, 17).Value = MaxPercentDecrease & "%"
        
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(4, 16).Value = MaxTotalVolumeTicker
        ws.Cells(4, 17).Value = MaxTotalVolume
        ' Add header "Ticker" in cell P1 and the corresponding value in cell Q1
          
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 9).Font.Bold = True
        ws.Cells(1, 10).Font.Bold = True
        ws.Cells(1, 11).Font.Bold = True
        ws.Cells(1, 12).Font.Bold = True
        ' Add headers and format summary table
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        ' Reset summary variables for the entire workbook
        MaxPercentIncrease = 0
        MaxPercentDecrease = 0
        MaxTotalVolume = 0
        MaxPercentIncreaseTicker = ""
        MaxPercentDecreaseTicker = ""
        MaxTotalVolumeTicker = ""
    Next ws
End Sub