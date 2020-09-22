
Sub fn_Ticker()

Dim lastRow As Long
 lastRow = Cells(Rows.Count, 1).End(xlUp).Row
 MsgBox ("last row for Year " & lastRow)

'Loop through rows for each year
 TickerYear = "2014"

currentTickerTtlCtr = 0
currentTickerName = ""
currentTickerOpenValue = 0
currentTickerCloseValue = 0
NewTicker = "Yes"
currentTotalTickerStockVol = 0
CntPrintTicker = 2

For y = 2 To lastRow
      currentTickerName = Cells(y, 1).Value
      NextRowTickerName = Cells(y + 1, 1).Value
      If currentTickerName = NextRowTickerName Then
                 
        If NewTicker = "Yes" Then
         NewTicker = "No"
                 currentTickerOpenValue = Cells(y, 3).Value
         End If
         
        currentTotalTickerStockVol = currentTotalTickerStockVol + Cells(y, 7).Value
         
      Else
      
       currentTickerCloseValue = Cells(y, 6).Value
      
       currentTotalTickerStockVol = currentTotalTickerStockVol + Cells(y, 7).Value
       Cells(CntPrintTicker, 13).Value = currentTickerName
                     
       Cells(CntPrintTicker, 14).Value = currentTickerOpenValue
       Cells(CntPrintTicker, 15).Value = currentTickerCloseValue
       Cells(CntPrintTicker, 18).Value = currentTotalTickerStockVol
       YearlyChange = currentTickerCloseValue - currentTickerOpenValue
       
       Cells(CntPrintTicker, 16).Value = YearlyChange
       
       If YearlyChange >= 0 Then
            Cells(CntPrintTicker, 16).Interior.ColorIndex = 4
       Else
            Cells(CntPrintTicker, 16).Interior.ColorIndex = 3
       End If
       
           percent_Chg = (currentTickerCloseValue - currentTickerOpenValue) / currentTickerOpenValue
 
           Cells(CntPrintTicker, 17).Value = percent_Chg
       
       CntPrintTicker = CntPrintTicker + 1
       currentTotalTickerStockVol = 0
       currentTickerOpenValue = ""
       currentTickerCloseValue = ""
       NewTicker = "Yes"
       YearlyChange = ""
       
    End If
    

Next y

 


End Sub




