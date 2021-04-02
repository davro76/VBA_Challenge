Sub StockMarketAnalyst():

Dim Sheet1 As Worksheet

Set Sheet1 = Worksheets("A")

'Created new columns'

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(2, 14).Value = "Greatest % Increase"
Cells(3, 14).Value = "Greatest % Decrease"
Cells(4, 14).Value = "Greatest Total Volume"
Cells(1, 15).Value = "Ticker"
Cells(1, 16).Value = "Value"


'Changed background of created columns'

Range("I1:L1").Interior.ColorIndex = 17

'Changed Font Color for created columns'

Range("I1:L1").Font.ColorIndex = 2

'Declaring variables to hold Ticker, Yearly Change, Percent Change, Last Row Number, and Total Stock Volume'
 Dim Sticker As String
 Dim LastRowNumber As Double
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim YearlyChange As Double
Dim PercentChange As Double
 LastRowNumber = Cells(Rows.Count, 1).End(xlUp).Row
 Dim TotalVolumeStock As Double
 TotalVolumeStock = 0
 
 ' Search Tricker in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  ' Loop through all credit card purchases
  For i = 2 To LastRowNumber

    ' Check if we are still within the same credit card brand, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the Sticker name
      Sticker = Cells(i, 1).Value

      ' Add to the Total Volume Stock
      TotalVolumeStock = TotalVolumeStock + Cells(i, 7).Value
    
      'Yearly Change
      OpenPrice = Cells(i, 3).Value
      ClosePrice = Cells(i, 6).Value
      YearlyChange = ClosePrice - OpenPrice
      'Percent Change
      PercentChange = (YearlyChange / OpenPrice) * 100
      
       
       ' Populated the Summary Table
      Range("I" & Summary_Table_Row).Value = Sticker
      Range("L" & Summary_Table_Row).Value = TotalVolumeStock
      Range("J" & Summary_Table_Row).Value = YearlyChange
      Range("K" & Summary_Table_Row).Value = PercentChange
    
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Total Volume Stock
      TotalVolumeStock = 0

    ' If the cell immediately following a row is the same brand...
    Else

      ' Add to the Total Volume Stock
      TotalVolumeStock = TotalVolumeStock + Cells(i, 7).Value
      
   
    End If
   

  Next i
  
  '
      
    'Greatest % increase, Greatest % decrease, Greatest Total Volume
    
    GreatestPercentIncrease = WorksheetFunction.Max(Sheet1.Range("K2:K290"))
    TickerGreatestPercentIncrease = WorksheetFunction.Match(GreatestPercentIncrease, Sheet1.Range("K2:K290"), 0)
    Sheet1.Cells(2, 15).Value = Sheet1.Cells(TickerGreatestPercentIncrease + 1, 1)
   Sheet1.Cells(2, 16).Value = GreatestPercentIncrease
   
   GreatestPercentDecrease = WorksheetFunction.Min(Range("K2:K290"))
    TickerGreatestPercentDecrease = WorksheetFunction.Match(GreatestPercentDecrease, Range("K2:K290"), 0)
    Sheet1.Cells(3, 15).Value = Sheet1.Cells(TickerGreatestPercentDecrease + 1, 1)
   Sheet1.Cells(3, 16).Value = GreatestPercentDecrease

GreatestTotalVolume = WorksheetFunction.Max(Sheet1.Range("L2:L290"))
    TickerGreatestTotalVolume = WorksheetFunction.Match(GreatestTotalVolume, Sheet1.Range("L2:L290"), 0)
    Sheet1.Cells(4, 15).Value = Sheet1.Cells(TickerGreatestTotalVolume + 1, 1)
   Sheet1.Cells(4, 16).Value = GreatestTotalVolume

End Sub




































