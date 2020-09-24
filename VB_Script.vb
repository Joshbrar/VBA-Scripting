
'Code to run for all worksheets

Sub runAll()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call fn_Ticker
    Next
    Application.ScreenUpdating = True
End Sub


' This function will do all the calculations and display the readings

Sub fn_Ticker()

Dim lastRow As Long



'get the last row of the worksheet

 lastRow = Cells(Rows.Count, 1).End(xlUp).Row
  ' MsgBox ("last row for Year " & lastRow)  ' read the data to process and displaytemp

'assign variables

currentTickerTtlCtr = 0
currentTickerName = ""
currentTickerOpenValue = 0
currentTickerCloseValue = 0
NewTicker = "Yes"
currentTotalTickerStockVol = 0
CntPrintTicker = 2

'Print Output Headers
       Cells(1, 13).Value = "Ticker"
       Cells(1, 14).Value = "Yearly Change"
       Cells(1, 15).Value = "Change Percentage"
       Cells(1, 16).Value = "Total Stock Volume"
       
       '  Cells(1, 17).Value = "openValue"   ' for debugging
       '  Cells(1, 18).Value = "CloseValue"   ' for debugging

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
      
      'Cells(CntPrintTicker, 14).Value = currentTickerOpenValue
      ' Cells(CntPrintTicker, 15).Value = currentTickerCloseValue
       
       Cells(CntPrintTicker, 16).Value = currentTotalTickerStockVol
       YearlyChange = currentTickerCloseValue - currentTickerOpenValue
       
       Cells(CntPrintTicker, 14).Value = YearlyChange
       
       ' code for changing color of cell if yearly change is positive or negative
       
       If YearlyChange >= 0 Then
            Cells(CntPrintTicker, 14).Interior.ColorIndex = 4
       Else
            Cells(CntPrintTicker, 14).Interior.ColorIndex = 3
       End If
         
         ' Make sure the TickerOpen Value is not zero   - question for Instructor and TA
         If currentTickerOpenValue > 0 Then
           percent_Chg = (currentTickerCloseValue - currentTickerOpenValue) / currentTickerOpenValue
 
           'Cells(CntPrintTicker, 15).Value = percent_Chg * 100 & "%"
              Cells(CntPrintTicker, 15).Value = FormatPercent(percent_Chg, , , , vbFalse)
         Else
         End If
         
           
       CntPrintTicker = CntPrintTicker + 1
       currentTotalTickerStockVol = 0
       currentTickerOpenValue = ""
       currentTickerCloseValue = ""
       NewTicker = "Yes"
       YearlyChange = ""
       
    End If
    

Next y

' GetMaxpercentage MinPercentage , Max Total Stock Values
 Cells(1, 18).Value = "Ticker"
 Cells(1, 19).Value = "Value"
 
 Cells(3, 17).Value = "Greatest % Increase :  "
 varGPI = Application.WorksheetFunction.Max(Range("O:O"))
 varRowNoMaxPercentVal = Application.WorksheetFunction.Match(Range("M:M").Application.WorksheetFunction.Max(Range("O:O")), Range("O:O"), 0)
'MsgBox (varRowNoMaxPercentVal)
 Cells(3, 18).Value = Cells(varRowNoMaxPercentVal, 13).Value
 Cells(3, 19).Value = FormatPercent(varGPI, , , , vbFalse)


  Cells(4, 17).Value = "Greatest % Decrease :  "
 
  varGPD = Application.WorksheetFunction.Min(Range("O:O"))
 
 varRowNoMinPercentVal = Application.WorksheetFunction.Match(Range("M:M").Application.WorksheetFunction.Min(Range("O:O")), Range("O:O"), 0)
'MsgBox (varRowNoMinPercentVal)
 
  Cells(4, 18).Value = Cells(varRowNoMinPercentVal, 13).Value
 
 Cells(4, 19).Value = FormatPercent(varGPD, , , , vbFalse)



 Cells(5, 17).Value = "Greatest Total Stock Value :  "
 Cells(5, 19).Value = Application.WorksheetFunction.Max(Range("P:P"))
 
 varRowNoMaxTotalStkVal = Application.WorksheetFunction.Match(Range("M:M").Application.WorksheetFunction.Max(Range("P:P")), Range("P:P"), 0)
 Cells(5, 18).Value = Cells(varRowNoMaxTotalStkVal, 13).Value
 

 
End Sub


