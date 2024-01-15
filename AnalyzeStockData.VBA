Sub AnalyzeStockData()

    'Declare Variables
    
    Dim stockSymbol As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim summaryRow As Double
    Dim greatestIncreaseStock As String
    Dim greatestDecreaseStock As String
    Dim ws As Worksheet
    Dim greatestVolumeStock As String
    Dim yearlyDifference As Double
    Dim percentageChange As Double
    Dim totalVolume As Double
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim openingRow As Double
    Dim summaryTable As Double

    'Loop through all worksheets to perform data analysis
    For Each ws In Worksheets

        'Set Column Headers
        
        ws.Range("I1").Value = "Stock Symbol"
        ws.Range("J1").Value = "Yearly Difference"
        ws.Range("K1").Value = "Percentage Change"
        ws.Range("L1").Value = "Total Stock Volume"

        'Initialize Variables
        summaryTable = 1
        totalVolume = 0
        openingRow = 2

        'Find the last row of the worksheet
        lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row

        'Analyze yearly change, percentage change, and total stock volume for each stock symbol
        For i = 2 To lastRow

            'Check if stock symbols are different
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                stockSymbol = ws.Cells(i, 1).Value
                summaryTable = summaryTable + 1

                'Retrieve values for the opening day and last day of the year
                yearOpen = ws.Cells(summaryTable, 3).Value
                yearClose = ws.Cells(i, 6).Value

                'Loop to calculate total stock volume for the stock symbol
                For j = summaryTable To i
                    totalVolume = totalVolume + ws.Cells(j, 7).Value
                Next j

                'Check for any 0 value in the opening column
                If yearOpen = 0 Then
                    percentageChange = yearClose
                Else
                    yearlyDifference = yearClose - yearOpen
                    percentageChange = yearlyDifference / yearOpen
                End If

                'Update values under respective headers
                    ws.Cells(openingRow, 9).Value = stockSymbol ' Changed from startRow
                    ws.Cells(openingRow, 10).Value = yearlyDifference ' Changed from startRow
                    ws.Cells(openingRow, 11).Value = percentageChange ' Changed from startRow

                'Format the Percentage Change column to percentage format
                    ws.Cells(openingRow, 11).NumberFormat = "0.00%"
                    ws.Cells(openingRow, 12).Value = totalVolume ' Changed from startRow

                'Format the Yearly Difference column
                    If yearlyDifference > 0 Then
                        ws.Cells(openingRow, 10).Interior.ColorIndex = 4 ' Green
                Else
                        ws.Cells(openingRow, 10).Interior.ColorIndex = 3 ' Red
                End If

                openingRow = openingRow + 1 ' Changed from startRow

                'Reset values to 0
                
                totalVolume = 0
                yearlyDifference = 0
                percentageChange = 0

                'Move i number to variable summaryTable
                summaryTable = i

            End If

        Next i
    Next ws

    ' Create a second summary table for Greatest % Increase, Greatest % Decrease, and Greatest Total Volume
    ' Count the number of rows in Percentage Change column in each worksheet
    For Each ws In Worksheets

        kLastRow = ws.Cells(Rows.Count, "K").End(xlUp).Row

        ' Initialize variables
        greatestIncrease = ws.Cells(2, 11).Value
        greatestDecrease = ws.Cells(2, 11).Value
        greatestVolume = ws.Cells(2, 12).Value
        greatestIncreaseStock = ws.Cells(2, 9).Value
        greatestDecreaseStock = ws.Cells(2, 9).Value
        greatestVolumeStock = ws.Cells(2, 9).Value

        'Loop through each row in column K (Percentage Change)
        For i = 2 To kLastRow
        
            'Update the greatest increase
            
            If ws.Cells(i, 11).Value > greatestIncrease Then
            
                greatestIncrease = ws.Cells(i, 11).Value
                greatestIncreaseStock = ws.Cells(i, 9).Value
                
            End If

            'Update greatest total volume
            
            If ws.Cells(i, 12).Value > greatestVolume Then
            
                greatestVolume = ws.Cells(i, 12).Value
                greatestVolumeStock = ws.Cells(i, 9).Value
                
            End If
            
        Next i

        
    'Output results in the summary table
        
            ws.Range("O2").Value = "Greatest % Increase"
            ws.Range("O3").Value = "Greatest % Decrease"
            ws.Range("O4").Value = "Greatest Total Volume"
            ws.Range("P1").Value = "Stock Symbol"
            ws.Range("Q1").Value = "Value"
            ws.Range("P2").Value = greatestIncreaseStock
            ws.Range("P3").Value = greatestDecreaseStock
            ws.Range("P4").Value = greatestVolumeStock
            ws.Range("Q2").Value = greatestIncrease
            ws.Range("Q3").Value = greatestDecrease
            ws.Range("Q4").Value = greatestVolume

    'Format percentage for percent change values
    
            ws.Range("Q2:Q3").NumberFormat = "0.00%"

    Next ws

End Sub
