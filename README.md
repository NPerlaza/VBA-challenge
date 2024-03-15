# VBA-challenge
Module 2 Challenge
Sub CalculateStockInfo()

    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets

        Dim startRow As Long
        Dim endRow As Long
        Dim outputRow As Long
        Dim ticker As String
        Dim yearlyChange As Double
        Dim percentChange As Double
        Dim totalVolume As Double

        ' Set headers for new columns
        ws.Range("H1").Value = "Ticker"
        ws.Range("I1").Value = "Yearly Change"
        ws.Range("J1").Value = "Percentage Change"
        ws.Range("K1").Value = "Total Stock Volume"

        ' Initialize startRow and outputRow
        startRow = 2
        outputRow = 2

        ' Loop through all rows
        For i = 2 To 753001
            ' Check if we're still on the same ticker
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                endRow = i
                ticker = ws.Cells(startRow, 1).Value
                yearlyChange = ws.Cells(endRow, 6).Value - ws.Cells(startRow, 3).Value
                percentChange = yearlyChange / ws.Cells(startRow, 3).Value
                totalVolume = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(startRow, 7), ws.Cells(endRow, 7)))

                ' Output to new columns
                ws.Cells(outputRow, 8).Value = ticker
                ws.Cells(outputRow, 9).Value = yearlyChange
                ws.Cells(outputRow, 10).Value = percentChange
                ws.Cells(outputRow, 11).Value = totalVolume

                ' Reset startRow and increment outputRow
                startRow = i + 1
                outputRow = outputRow + 1
            End If
        Next i
