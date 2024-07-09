Attribute VB_Name = "Module1"
Sub CalculateQuarterlySummaryForAllQuarters()
    Dim ws As Worksheet
    Dim summaryRow As Long
    Dim colOffset As Integer
    Dim lastRow As Long
    Dim startRow As Long
    Dim endRow As Long
    Dim i As Long
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim totalVolume As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim rng As Range
    Dim quarterSheets As Variant
    Dim sheetName As Variant
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String
    Dim greatestIncreaseValue As Double
    Dim greatestDecreaseValue As Double
    Dim greatestVolumeValue As Double
    
    ' Define the sheet names for each quarter
    quarterSheets = Array("Q1", "Q2", "Q3", "Q4")
    
    ' Loop through each quarter sheet
    For Each sheetName In quarterSheets
        Set ws = Worksheets(sheetName)
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        ' Initialize summary sheet headers
        colOffset = 8 ' Start from column I
        ws.Cells(1, colOffset + 1).Value = "Ticker"
        ws.Cells(1, colOffset + 2).Value = "Quarterly Change"
        ws.Cells(1, colOffset + 3).Value = "Percent Change"
        ws.Cells(1, colOffset + 4).Value = "Total Stock Volume"
        
        summaryRow = 2
        
        ' Initialize variables
        startRow = 2
        greatestIncreaseValue = -1
        greatestDecreaseValue = 1
        greatestVolumeValue = -1

        ' Loop through each row in the data
        For i = 2 To lastRow
            ' Check if the ticker has changed or if it's the last row
            If ws.Cells(i, 1).Value <> ws.Cells(startRow, 1).Value Or i = lastRow Then
                ' Process the previous group of data
                If i = lastRow Then endRow = i Else endRow = i - 1
                ticker = ws.Cells(startRow, 1).Value
                openPrice = ws.Cells(startRow, 3).Value
                closePrice = ws.Cells(endRow, 6).Value
                totalVolume = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(startRow, 7), ws.Cells(endRow, 7)))
                quarterlyChange = closePrice - openPrice
                percentChange = (quarterlyChange / openPrice)

                ' Write the result to the summary section in the same sheet
                ws.Cells(summaryRow, colOffset + 1).Value = ticker
                ws.Cells(summaryRow, colOffset + 2).Value = quarterlyChange
                ws.Cells(summaryRow, colOffset + 3).Value = Format(percentChange, "0.00%")
                ws.Cells(summaryRow, colOffset + 4).Value = totalVolume
                summaryRow = summaryRow + 1

                ' Check for greatest % increase, % decrease, and total volume
                If percentChange > greatestIncreaseValue Then
                    greatestIncreaseValue = percentChange
                    greatestIncreaseTicker = ticker
                End If
                If percentChange < greatestDecreaseValue Then
                    greatestDecreaseValue = percentChange
                    greatestDecreaseTicker = ticker
                End If
                If totalVolume > greatestVolumeValue Then
                    greatestVolumeValue = totalVolume
                    greatestVolumeTicker = ticker
                End If

                ' Reset variables for the new group
                startRow = i
            End If
        Next i

        ' Apply conditional formatting to the summary section
        With ws
            ' Define the range for quarterly change
            Set rng = .Range(.Cells(2, colOffset + 2), .Cells(summaryRow - 1, colOffset + 2))
            ' Apply conditional formatting to highlight positive changes in green
            With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
                .Interior.Color = RGB(0, 255, 0)
            End With
            ' Apply conditional formatting to highlight negative changes in red
            With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
                .Interior.Color = RGB(255, 0, 0)
            End With
        End With

        ' Output the greatest values
        ws.Cells(1, colOffset + 6).Value = "Ticker"
        ws.Cells(1, colOffset + 7).Value = "Value"
        ws.Cells(2, colOffset + 5).Value = "Greatest % Increase"
        ws.Cells(2, colOffset + 6).Value = greatestIncreaseTicker
        ws.Cells(2, colOffset + 7).Value = Format(greatestIncreaseValue, "0.00%")
        ws.Cells(3, colOffset + 5).Value = "Greatest % Decrease"
        ws.Cells(3, colOffset + 6).Value = greatestDecreaseTicker
        ws.Cells(3, colOffset + 7).Value = Format(greatestDecreaseValue, "0.00%")
        ws.Cells(4, colOffset + 5).Value = "Greatest Total Volume"
        ws.Cells(4, colOffset + 6).Value = greatestVolumeTicker
        ws.Cells(4, colOffset + 7).Value = greatestVolumeValue

    Next sheetName

    MsgBox "Quarterly summary has been calculated for all quarters!", vbInformation
End Sub

