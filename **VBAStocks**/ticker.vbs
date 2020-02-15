' * Create a script that will loop through all the stocks for one year for each run and take the following information.

' * The ticker symbol.

' * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

' * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

' * The total stock volume of the stock.

' * You should also have conditional formatting that will highlight positive change in green and negative change in red.

Sub ticker()

  Dim summary_row As Long
  summary_row = 2
  Dim totalVolume As LongLong
  totalVolume = 0
  ' For Each ws In Worksheets
  Dim OpeningPrice As Variant
  OpeningPrice = Cells(2, 3).Value
  Dim ClosingPrice As Variant
  ' ClosingPrice = 0
  Dim YearlyChange As Variant
  ' YearlyChange = 0
  Dim PercentChange As Double
  ' PercentChange = 0

  lastrow = Cells(Rows.Count, 1).End(xlUp).Row

  ' Columns("A:G").Sort key1:=Range("B2"), _
  '   order1:=xlAscending, Header:=xlYes
  ' Columns("A:G").Sort key1:=Range("A2"), _
  '   order1:=xlAscending, Header:=xlYes
  

  Range("I1").Value = "Ticker"
  Range("J1").Value = "$ Change"
  Range("K1").Value = "% Change"
  Range("L1").Value = "Total Volume"

  For Row = 2 To lastrow
    If Cells(Row, 1).Value <> Cells(Row + 1, 1).Value Then
      ClosingPrice = Cells(Row, 6).Value
      YearlyChange = ClosingPrice - OpeningPrice

      If OpeningPrice = 0 Then
        PercentChange = 0
      Else
        PercentChange = YearlyChange / OpeningPrice
      End If
      ' Cells(row,10).NumberFormat = "$00.00"
      ' Cells(row,11).NumberFormat = "0.00%"
      totalVolume = totalVolume + Cells(Row, 7).Value

      Range("I" & summary_row).Value = Cells(Row, 1).Value
      Range("J" & summary_row).Value = YearlyChange 
      Range("K" & summary_row).Value = PercentChange
      Range("L" & summary_row).Value = totalVolume
      
      summary_row = summary_row + 1
      totalVolume = 0
      OpeningPrice = Cells(Row + 1, 3)
    Else
      totalVolume = totalVolume + Cells(Row, 7).Value
    End If
  Next Row

End Sub