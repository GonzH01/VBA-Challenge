Sub Stockdata()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long, outputRow As Long
    Dim ticker As String, rowDate As Long
    Dim openingPrice As Variant, closingPrice As Variant
    Dim yearlyChange As Double, percentageChange As Double
    Dim formattedPercentageChange As String
    Dim volume As Variant, totalVolume As Double
    Dim foundStart As Boolean, foundEnd As Boolean

    ' Define the start and end dates
    Const startDate As Long = 20180102
    Const endDate As Long = 20181231

    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("2018")

    ' Find the last row with data in column 1 (ticker column)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Initialize output row for column I (9) to start from the second row
    outputRow = 2

    ' Loop through each ticker
    For i = 1 To lastRow
        ticker = ws.Cells(i, 1).Value
        rowDate = Val(ws.Cells(i, 2).Value) ' Convert date to numeric value

        ' Check for the start date and get opening price
        If rowDate = startDate And ws.Cells(i, 1).Value = ticker Then
            openingPrice = ws.Cells(i, 3).Value
            foundStart = True
        End If

        ' Check for the end date and get closing price
        If rowDate = endDate And ws.Cells(i, 1).Value = ticker Then
            closingPrice = ws.Cells(i, 6).Value
            foundEnd = True
        End If

        ' Accumulate the total volume for the ticker
        If ws.Cells(i, 1).Value = ticker Then
            volume = ws.Cells(i, 7).Value
            If IsNumeric(volume) Then
                totalVolume = totalVolume + volume
            End If
        End If

        ' If both start and end dates are found for the ticker, calculate the change and percentage
        If foundStart And foundEnd Then
            If IsNumeric(openingPrice) And IsNumeric(closingPrice) Then
                yearlyChange = closingPrice - openingPrice
                If openingPrice <> 0 Then
                    percentageChange = (yearlyChange / openingPrice) * 100
                Else
                    percentageChange = 0
                End If
                formattedPercentageChange = Format(percentageChange, "0.00") & "%"
            Else
                yearlyChange = 0
                formattedPercentageChange = "N/A"
            End If

            ' Output the ticker, yearly change, formatted percentage change, and total volume
            ws.Cells(outputRow, 9).Value = ticker
            ws.Cells(outputRow, 10).Value = yearlyChange
            ws.Cells(outputRow, 11).Value = formattedPercentageChange
            ws.Cells(outputRow, 12).Value = totalVolume

            ' Move to the next output row and reset flags and total volume
            outputRow = outputRow + 1
            foundStart = False
            foundEnd = False
            totalVolume = 0
        End If
    Next i
End Sub
