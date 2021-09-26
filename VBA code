Sub stockStats():

' creating variables
Dim ticker As String
Dim openPrice As Double
Dim closePrice As Double
Dim vol As Integer
Dim lastRow As Long
Dim yearlyChange As Double
Dim percentChange As Double
Dim totalStockVol As Double
Dim nextTicker As Integer

' placing each variable as heaaders on each worksheet
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

For Each ws In Worksheets
    
    ' finding the last row in each worksheet to keep from having to manually calculate it
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' initializing nextTicker and ticker at 0 and an empty set
     nextTicker = 0
     ticker = ""
    
    ' starting the range at 2 so that headers are not included in the calculations
    For i = 2 To lastRow
        
        ' Get the start of the year opening price for the ticker
        If openPrice = 0 Then
            openPrice = Cells(i, 3).Value
        End If
        
        ' calculating the total stock vol for each ticker
        totalStockVol = totalStockVol + Cells(i, 7).Value
        
        ' stating the variable "ticker" is located in column 1
        ticker = Cells(i, 1).Value
        
        ' if when looping through column 1 the value of tickers are not equal, 
        ' then we want to add the ticker to column 9 and move onto the next ticker
        If Cells(i + 1, 1) <> ticker Then
            nextTicker = nextTicker + 1
            Cells(nextTicker + 1, 9) = ticker
            
            ' stating the variable "closePrice" is located in column 6
            closePrice = Cells(i, 6).Value
            
            ' equation for finding yearly change
            yearlyChange = closePrice - openPrice
            
            ' placing values of yearly change for each ticker in column 10, and increasing the row each time there's a new ticker
            Cells(nextTicker + 1, 10) = yearlyChange
            
            ' if yearlyChange is less than 0, we want the color of the cell to be red
             If yearlyChange < 0 Then
                Cells(nextTicker + 1, 10).Interior.ColorIndex = 3
           ' else we want the color of the cell to be green
           Else:
                Cells(nextTicker + 1, 10).Interior.ColorIndex = 4
            End If
            
            ' since we cannot divide by 0, if openPrice is 0, then the percentChange is also going to be 0
            If openPrice = 0 Then
                percentChange = 0
            ' equation for finding percent change
            Else
                percentChange = (closePrice - openPrice) / openPrice
            End If
            
            ' placing the values of percent change for each ticker in column 11, and increasing the row each time there's a new ticker
            Cells(nextTicker + 1, 11).Value = percentChange
            ' changing the format of the percent change values to percent form
            Cells(nextTicker + 1, 11).Value = Format(percentChange, "Percent")
            ' placing the values of percent change for each ticker in column 12, and increasing the row each time there's a new ticker
            Cells(nextTicker + 1, 12).Value = totalStockVol
            
            ' setting openPrice and totalStockVol back to 0 when loop starts over
            openPrice = 0
            totalStockVol = 0
        
        End If
        
    Next i
    
Next ws


End Sub
