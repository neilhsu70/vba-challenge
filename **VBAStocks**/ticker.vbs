  Sub ticker()

    For Each ws In Worksheets

    Dim summary_row As Long
    summary_row = 2

    Dim totalVolume As LongLong
    totalVolume = 0

    Dim OpeningPrice As Variant
    OpeningPrice = ws.Cells(2, 3).Value

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

    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row    

    ' Label Headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "$ Change"
    ws.Range("K1").Value = "% Change"
    ws.Range("L1").Value = "Total Volume"
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Challenge Ticker"
    ws.Range("Q1").Value = "Challenge Value"


    For Row = 2 To lastrow
      ' format $ & % symbols
      ws.Cells(row,11).NumberFormat = "0.00%"
      ws.Cells(row,10).NumberFormat = "$0.00"

      ' Main Section
      If ws.Cells(Row, 1).Value <> ws.Cells(Row + 1, 1).Value Then
        ClosingPrice = ws.Cells(Row, 6).Value

        'Calculate Price Change
        YearlyChange = ClosingPrice - OpeningPrice

        ' Calculate % Change
        If OpeningPrice = 0 Then
          PercentChange = 0
        Else
          PercentChange = YearlyChange / OpeningPrice
        End If

        ' Calculate Total Volume
        totalVolume = totalVolume + ws.Cells(Row, 7).Value

        ' Greatest Percent Increase
        If OpeningPrice <> 0 Then
          If YearlyChange / OpeningPrice > GreatPercentInc Then
            GreatPercentInc = ws.Cells(Row, 11).Value
            GreatPercentIncTiker = ws.Cells(Row, 9).Value
          End If
        End If

        ' Greatest Percent Decrease
        If OpeningPrice <> 0 Then
          If YearlyChange / OpeningPrice < GreatPercentDec Then
              GreatPercentDec = ws.Cells(Row, 11).Value
              GreatPercentDecSet = ws.Cells(Row, 9).Value
          End If
        End If
        
        ' Greatest Total Volume
        If totalVolume > GreatTotalVol Then
          GreatTotalVol = ws.Cells(Row, 11).Value
          GreatTotalVolTiker = ws.Cells(Row, 9).Value
        End If

        ' Place Result Values
        ws.Range("I" & summary_row).Value = ws.Cells(Row, 1).Value
        ws.Range("J" & summary_row).Value = YearlyChange 
        ws.Range("K" & summary_row).Value = PercentChange
        ws.Range("L" & summary_row).Value = totalVolume

        ' Place Greatest % Increase
        ws.Range("P2").Value = GreatPercentIncTiker
        ws.Range("Q2").value = GreatPercentInc
        ws.Range("Q2").NumberFormat = "0.00%"

        ' Place Greatest % Decrease
        ws.Range("P3").Value = GreatPercentDecTiker
        ws.Range("Q3").value = GreatPercentDec
        ws.Range("Q3").NumberFormat = "0.00%"

        ' Place Greatest Total Volume
        ws.Range("P4").Value = GreatTotalVolTiker
        ws.Range("Q4").value = GreatTotalVol

        summary_row = summary_row + 1
        totalVolume = 0
        OpeningPrice = ws.Cells(Row + 1, 3)
      Else
        totalVolume = totalVolume + ws.Cells(Row, 7).Value
      End If
    Next Row

    ' Add colors for positive or negative Price Change
    For Row = 2 to lastrow
      If ws.Cells(Row, 10) < 0 THEN
        ws.Cells(Row, 10).Interior.ColorIndex = 3
      Else
        ws.Cells(Row, 10).Interior.ColorIndex = 4
      End If
    Next Row

    Next ws

  End Sub