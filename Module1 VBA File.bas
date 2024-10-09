Attribute VB_Name = "Module1"
Sub QuarterlyStockAnalysis()
    Dim ws As Worksheet
    Dim sheetNames As Variant
    Dim lastRow As Long
    Dim quarterlyChange As Double
    Dim percentageChange As Double
    Dim totalVolume As Double
    Dim openPrice As Double
    Dim closePrice As Double
    Dim currentTicker As Variant
    Dim previousTicker As Variant
    Dim i As Long
    Dim outputRow As Long
    Dim sheetName As Variant
    Dim tickerDict As Object
    Dim dictKeys As Variant
    Dim ticker As Variant
    Dim startDate As Date
    Dim endDate As Date
    Dim currentDate As Date
    Dim firstOpenPrice As Boolean

    
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String

  
    startDate = DateSerial(2022, 1, 1)
    endDate = DateSerial(2022, 12, 31)
    
   
    sheetNames = Array("Q1", "Q2", "Q3", "Q4")
    
   
    For Each sheetName In sheetNames
     
        Set tickerDict = CreateObject("Scripting.Dictionary")
        
    
        greatestIncrease = -99999
        greatestDecrease = 99999
        greatestVolume = 0

        Set ws = ThisWorkbook.Sheets(sheetName)
        
        
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        
        previousTicker = ""
        firstOpenPrice = True
        For i = 2 To lastRow
            currentTicker = ws.Cells(i, 1).Value
            currentDate = ws.Cells(i, 2).Value
            
            
            If currentDate >= startDate And currentDate <= endDate Then
                If currentTicker <> previousTicker And previousTicker <> "" Then
                  
                    If Not tickerDict.exists(previousTicker) Then
                        tickerDict.Add previousTicker, Array(quarterlyChange, percentageChange, totalVolume, 1)
                    Else
                        tickerDict(previousTicker)(0) = tickerDict(previousTicker)(0) + quarterlyChange
                        tickerDict(previousTicker)(1) = tickerDict(previousTicker)(1) + percentageChange
                        tickerDict(previousTicker)(2) = tickerDict(previousTicker)(2) + totalVolume
                        tickerDict(previousTicker)(3) = tickerDict(previousTicker)(3) + 1
                    End If
                    
                    
                    quarterlyChange = 0
                    percentageChange = 0
                    totalVolume = 0
                    firstOpenPrice = True
                End If
                
                
                If firstOpenPrice Then
                    openPrice = ws.Cells(i, 3).Value
                    firstOpenPrice = False
                End If
                closePrice = ws.Cells(i, 6).Value
                quarterlyChange = closePrice - openPrice
                percentageChange = ((closePrice - openPrice) / openPrice)
                totalVolume = totalVolume + ws.Cells(i, 7).Value
                
                
                previousTicker = currentTicker
            End If
        Next i
        
      
        If previousTicker <> "" Then
            If Not tickerDict.exists(previousTicker) Then
                tickerDict.Add previousTicker, Array(quarterlyChange, percentageChange, totalVolume, 1)
            Else
                tickerDict(previousTicker)(0) = tickerDict(previousTicker)(0) + quarterlyChange
                tickerDict(previousTicker)(1) = tickerDict(previousTicker)(1) + percentageChange
                tickerDict(previousTicker)(2) = tickerDict(previousTicker)(2) + totalVolume
                tickerDict(previousTicker)(3) = tickerDict(previousTicker)(3) + 1
            End If
        End If
        
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percentage Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        outputRow = 2
        dictKeys = tickerDict.Keys
        For Each ticker In dictKeys
            ws.Cells(outputRow, 9).Value = ticker
            ws.Cells(outputRow, 10).Value = tickerDict(ticker)(0)
            ws.Cells(outputRow, 11).Value = tickerDict(ticker)(1)
            ws.Cells(outputRow, 12).Value = tickerDict(ticker)(2)
            
            
            If tickerDict(ticker)(1) > greatestIncrease Then
                greatestIncrease = tickerDict(ticker)(1)
                greatestIncreaseTicker = ticker
            End If
            If tickerDict(ticker)(1) < greatestDecrease Then
                greatestDecrease = tickerDict(ticker)(1)
                greatestDecreaseTicker = ticker
            End If
            If tickerDict(ticker)(2) > greatestVolume Then
                greatestVolume = tickerDict(ticker)(2)
                greatestVolumeTicker = ticker
            End If
            
            
            With ws.Cells(outputRow, 10)
                If tickerDict(ticker)(0) > 0 Then
                    .Interior.Color = RGB(0, 255, 0)
                ElseIf tickerDict(ticker)(0) < 0 Then
                    .Interior.Color = RGB(255, 0, 0)
                Else
                    .Interior.ColorIndex = xlNone
                End If
            End With
            
            outputRow = outputRow + 1
        Next ticker
        
        
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        ws.Cells(2, 15).Value = greatestIncreaseTicker
        ws.Cells(3, 15).Value = greatestDecreaseTicker
        ws.Cells(4, 15).Value = greatestVolumeTicker
        ws.Cells(2, 16).Value = greatestIncrease
        ws.Cells(3, 16).Value = greatestDecrease
        ws.Cells(4, 16).Value = greatestVolume
    Next sheetName
    
    MsgBox "Quarterly stock analysis complete!"
End Sub

