Attribute VB_Name = "Module1"
Sub stocks_hw()

    Dim tickerName As String
    Dim openValue As Double
    Dim highValue As Double
    Dim lowValue As Double
    Dim closeValue As Double
    Dim volume As Double

    Dim lastRow As Long
    Dim counter As Long
    counter = 2
    Dim i As Long
    Dim yearPercent As Double
    
    
    Dim GreatVol As Double
    Dim GreatIncr As Double
    Dim GreatDecr As Double
    Dim TickerGreatestIncr As String
    Dim TickerGreatestDecr As String
    Dim TickerGreatestVol As String

    Dim ws As Worksheet
    Set ws = ActiveSheet

    lastRow = ws.UsedRange.Rows(ws.UsedRange.Rows.Count).Row

ws.Cells(1, 10).Value = "Ticker"
ws.Cells(1, 11).Value = "Quarterly Change"
ws.Cells(1, 12).Value = "Percent Change"
ws.Cells(1, 13).Value = "Total Stock Volume"

    For i = 1 To lastRow

        If i <= 1 Then
            tickerName = ws.Cells(2, 1)
            ws.Cells(counter, 10) = tickerName
            openValue = ws.Cells(2, 3)
        Else
            If ws.Cells(i, 1).Value = tickerName Then
                volume = volume + ws.Cells(i, 7)
            Else
                closeValue = ws.Cells(i - 1, 6)
                ws.Cells(counter, 11) = (closeValue - openValue)
                yearPercent = (closeValue - openValue) / openValue * 100
                ws.Cells(counter, 12).NumberFormat = "0.00%"
                ws.Cells(counter, 12) = yearPercent & "%"
                ws.Cells(counter, 13) = volume
                
                If ws.Cells(counter, 11).Value < 0 Then
                   ws.Cells(counter, 11).Interior.Color = RGB(255, 0, 0)
                Else
                    ws.Cells(counter, 11).Interior.Color = RGB(0, 255, 0)
                End If
                
                ' Check for greatest increase, decrease, and volume
                If yearPercent > GreatIncr Then
                    GreatIncr = yearPercent
                    TickerGreatestIncr = tickerName
                  
                End If
                
                If yearPercent < GreatDecr Then
                    GreatDecr = yearPercent
                    TickerGreatestDecr = tickerName
                End If
                
                If volume > GreatVol Then
                    GreatVol = volume
                    TickerGreatestVol = tickerName
                End If
                
                counter = counter + 1
                
                ' Reset values for the next stock
                tickerName = ws.Cells(i, 1).Value
                ws.Cells(counter, 10) = tickerName
                openValue = ws.Cells(i, 3)
                volume = 0
                volume = ws.Cells(i, 7)
            End If
        End If
    Next i
    
    ' Output the results for greatest increase, decrease, and volume
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
    
    
    ws.Cells(2, 16).Value = TickerGreatestIncr
    ws.Cells(2, 17).Value = Format(GreatIncr / 100, "0.00%")
    ws.Cells(3, 16).Value = TickerGreatestDecr
    ws.Cells(3, 17).Value = Format(GreatDecr / 100, "0.00%")
    ws.Cells(4, 16).Value = TickerGreatestVol
    ws.Cells(4, 17).Value = Format(GreatVol, "Scientific")

End Sub

