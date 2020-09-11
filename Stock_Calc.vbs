Sub CalcStock()
    'Declaring required variables
    Dim tableRow As Integer
    Dim openPrice, closePrice, yearlyChange, percentChange As Double
    Dim TotVolume As Double
    Dim ticker As String
    
    'Variables for Greatest increase, decrease and total
    '===============================================
    Dim highInc, highDec As Double
    Dim highVol As Double
    Dim hiTickInc, hiTickDec, hiTickVol As String
    '===============================================
    
    'looping through each worksheets
    For Each ws In Worksheets
    
        'Setting up title names for each column for the summary table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        'Set initial value
        tableRow = 2
        openPrice = ws.Cells(2, 3).Value
        ticker = ws.Cells(2, 1).Value
        highInc = 0
        highDec = 0
        highVol = 0
        
        'Begin looping through data
        For rowStock = 2 To Range("A:A").End(xlDown).Row
            TotVolume = TotVolume + (ws.Cells(rowStock, 7))
            
            If ws.Cells(rowStock, 1) <> ws.Cells(rowStock + 1, 1) Then
                closePrice = ws.Cells(rowStock, 6).Value
                ws.Cells(tableRow, 9).Value = ticker
                yearlyChange = closePrice - openPrice
                ws.Cells(tableRow, 12).Value = TotVolume
                
                'Change color of cell based on whether it is positive(green) or negative(red) yearly changes
                If (yearlyChange > 0) Then
                    With ws.Cells(tableRow, 10)
                        .Value = yearlyChange
                        .Interior.ColorIndex = 4
                    End With
                ElseIf (yearlyChange < 0) Then
                    With ws.Cells(tableRow, 10)
                        .Value = yearlyChange
                        .Interior.ColorIndex = 3
                    End With
                Else
                    ws.Cells(tableRow, 10).Value = yearlyChange
                End If
                
                'A check to make sure VBA does not divide by 0
                If (openPrice = 0) Then
                    percentChange = 0
                    ws.Cells(tableRow, 11).Value = "undefined"
                Else
                    percentChange = yearlyChange / openPrice
                    With ws.Cells(tableRow, 11)
                        .Value = percentChange
                        .NumberFormat = "0.00%"
                    End With
                End If
                
                'Holding the greatest value per iteration
                '==============================================================
                
                'Determining the greatest increase and greatest decrease
                If (percentChange >= 0) Then
                    If (highInc < percentChange) Then
                        highInc = percentChange
                        hiTickInc = ticker
                    End If
                Else
                    If (highDec > percentChange) Then
                        highDec = percentChange
                        hiTickDec = ticker
                    End If
                End If
                
                'Determining greatest volume
                If (highVol < TotVolume) Then
                    highVol = TotVolume
                    hiTickVol = ticker
                End If
                
                '================================================================
                
                
                
                'Reinitiate the value for next iteration
                openPrice = ws.Cells(rowStock + 1, 3).Value
                TotVolume = 0
                ticker = ws.Cells(rowStock + 1, 1).Value
                tableRow = tableRow + 1
            End If
            
        Next rowStock
        
        'Populate the cells for greatest inc, dec and vol
        '=======================================================================
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("P2").Value = hiTickInc
        ws.Range("P3").Value = hiTickDec
        ws.Range("P4").Value = hiTickVol
        ws.Range("Q1").Value = "Value"
        With ws.Range("Q2")
            .Value = highInc
            .NumberFormat = "0.00%"
        End With
        With ws.Range("Q3")
            .Value = highDec
            .NumberFormat = "0.00%"
        End With
        ws.Range("Q4").Value = highVol
        '=======================================================================
    
    Next ws
    
    MsgBox ("Finish")
    
End Sub

