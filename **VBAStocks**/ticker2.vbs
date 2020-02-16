  Sub ticker()

    ' For Each ws In Worksheets

    Dim summary_row As Long
    summary_row = 2

    Dim totalVolume As LongLong
    totalVolume = 0

    Dim OpeningPrice As Variant
    OpeningPrice = Cells(2, 3).Value

    Dim ClosingPrice As Variant
    ' ClosingPrice = 0

    Dim YearlyChange As Variant
    ' YearlyChange = 0

    Dim PercentChange As Variant
    ' PercentChange = 0

    Dim GreatPercentInc As Variant
    GreatPercentInc = 0
    Dim GreatPercentIncTiker As String

    Dim GreatPercentDec As Variant
    GreatPercentDec = 0
    Dim GreatPercentDecTiker As String

    Dim GreatTotalVol As Variant
    GreatTotalVol = 0
    Dim GreatTotalVolTiker As String

    lastrow = Cells(Rows.Count, 1).End(xlUp).Row    

    ' Label Headers
    Range("I1").Value = "Ticker"
    Range("J1").Value = "$ Change"
    Range("K1").Value = "% Change"
    Range("L1").Value = "Total Volume"
    Range("N2").Value = "Greatest % Increase"
    Range("N3").Value = "Greatest % Decrease"
    Range("N4").Value = "Greatest Total Volume"
    Range("P1").Value = "Challenge Ticker"
    Range("Q1").Value = "Challenge Value"


    For Row = 2 To lastrow
      ' format $ & % symbols
      Cells(row,11).NumberFormat = "0.00%"
      Cells(row,10).NumberFormat = "$0.00"

      ' Main Section
      If Cells(Row, 1).Value <> Cells(Row + 1, 1).Value Then
        ClosingPrice = Cells(Row, 6).Value

        'Calculate Price Change
        YearlyChange = ClosingPrice - OpeningPrice

        ' Calculate % Change
        If OpeningPrice = 0 Then
          PercentChange = 0
        Else
          PercentChange = YearlyChange / OpeningPrice
        End If

        ' Calculate Total Volume
        totalVolume = totalVolume + Cells(Row, 7).Value

        ' Greatest Percent Increase
        ' If OpeningPrice <> 0 Then
        '   If YearlyChange / OpeningPrice > GreatPercentInc Then
        '     GreatPercentInc = Cells(Row, 11).Value
        '     GreatPercentIncTiker = Cells(Row, 9).Value
        '   End If
        ' End If

        ' Greatest Percent Decrease
        ' If OpeningPrice <> 0 Then
        '   If YearlyChange / OpeningPrice < GreatPercentDec Then
        '       GreatPercentDec = Cells(Row, 11).Value
        '       GreatPercentDecSet = Cells(Row, 9).Value
        '   End If
        ' End If
        
        ' ' Greatest Total Volume
        ' If totalVolume > GreatTotalVol Then
        '   GreatTotalVol = Cells(Row, 11).Value
        '   GreatTotalVolTiker = Cells(Row, 9).Value
        ' End If

        ' Place Result Values
        Range("I" & summary_row).Value = Cells(Row, 1).Value
        Range("J" & summary_row).Value = YearlyChange 
        Range("K" & summary_row).Value = PercentChange
        Range("L" & summary_row).Value = totalVolume

        ' ' Place Greatest % Increase
        ' Range("P2").Value = GreatPercentIncTiker
        ' Range("Q2").value = GreatPercentInc
        ' Range("Q2").NumberFormat = "0.00%"

        ' ' Place Greatest % Decrease
        ' Range("P3").Value = GreatPercentDecTiker
        ' Range("Q3").value = GreatPercentDec
        ' Range("Q3").NumberFormat = "0.00%"

        ' ' Place Greatest Total Volume
        ' Range("P4").Value = GreatTotalVolTiker
        ' Range("Q4").value = GreatTotalVol

        summary_row = summary_row + 1
        totalVolume = 0
        OpeningPrice = Cells(Row + 1, 3)
      Else
        totalVolume = totalVolume + Cells(Row, 7).Value
      End If
    Next Row

    ' Add colors for positive or negative Price Change
    For Row = 2 to lastrow
      If Cells(Row, 10) < 0 THEN
        Cells(Row, 10).Interior.ColorIndex = 3
      Else
        Cells(Row, 10).Interior.ColorIndex = 4
      End If
    Next Row

    ' Next ws

  End Sub