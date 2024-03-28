Sub vbachallenge()

For Each ws In Worksheets

Dim worksheetname As String
worksheetname = ws.Name

Dim lastrow As Long
lastrow = 0


Dim Ticker As String
Dim tickerrow As Double
tickerrow = 2

Dim stockvolume As LongLong
Dim stockvolumerow As Double
stockvolume = 0

Dim percentchange As Double
Dim openvalue As Double
Dim closevalue As Double
Dim opencounter As Double


Dim maxvalue As Double
Dim maxticker As String
maxvalue = 0

Dim minvalue As Double
Dim minticker As String

minvalue = 999


Dim maxvaluestock As Double
Dim maxvolumeticker As String
maxvaluestock = 0

Dim k As Double
k = 2

ws.Cells(1, 9).Value = "Ticker"
    
ws.Cells(1, 10).Value = "Yearly Change"
      
ws.Cells(1, 11).Value = "Percent Change"
    
ws.Cells(1, 12).Value = "Total Stock Volume"

ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"

ws.Cells(1, 16).Value = "Ticker"

ws.Cells(1, 17).Value = "Value"


lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row


 For i = 2 To lastrow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
            
            Ticker = ws.Cells(i, 1).Value
            stockvolume = stockvolume + ws.Cells(i, 7).Value
            
            closevalue = ws.Cells(i, 6).Value

            ws.Range("i" & tickerrow).Value = Ticker
            ws.Range("L" & tickerrow).Value = stockvolume
            
            If stockvolume > maxvaluestock Then
                maxvaluestock = stockvolume
                maxvolumeticker = ws.Cells(i, 1).Value
            End If
            
            stockvolume = 0
            closevalue = ws.Cells(i, 6).Value
            openvalue = ws.Cells(i - opencounter, 3).Value
            
        
            percentchange = ((closevalue - openvalue) / openvalue)
            
                If percentchange > maxvalue Then
                    maxvalue = percentchange
                    maxticker = ws.Cells(i, 1).Value
                    
                End If
                
                If percentchange < minvalue Then
                    minvalue = percentchange
                    minticker = ws.Cells(i, 1).Value
                End If
                
                
            ws.Range("k" & tickerrow).Value = percentchange
            ws.Range("j" & tickerrow).Value = closevalue - openvalue
            
                If ws.Cells(k, 10).Value >= 0 Then
                   ws.Cells(k, 10).Interior.ColorIndex = 4
                Else
                   ws.Cells(k, 10).Interior.ColorIndex = 3
                End If
            k = k + 1
            
            tickerrow = tickerrow + 1
            opencounter = 0
            
            Else
                
                opencounter = opencounter + 1
                stockvolume = stockvolume + ws.Cells(i, 7).Value
                

                
            End If
            ws.Cells(2, 17).Value = maxvalue
            ws.Cells(3, 17).Value = minvalue
            
            ws.Cells(4, 16).Value = maxvolumeticker
            ws.Cells(4, 17).Value = maxvaluestock
            
            ws.Cells(3, 16).Value = minticker
            ws.Cells(2, 16).Value = maxticker
        Next i

        ws.Range("k:k").NumberFormat = "#0.00%"
        ws.Range("q2:q3").NumberFormat = "#0.00%"

    Next ws
    


End Sub
