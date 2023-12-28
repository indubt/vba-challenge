'Create a script that loops through all the stocks for one year and outputs the following information:

'The ticker symbol

'Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.

'The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.

'The total stock volume of the stock. The result should match the following image:


Sub Stock_Rpt()


    For Each ws In Worksheets

        Dim WorksheetName As String
        
        WorksheetName = ws.Name
        
         ' Determine the Last Row
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Add Columns Headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        ' Add Labels
        
        ws.Range("R4").Value = "Greatest % Increase"
        ws.Range("R5").Value = "Greatest % Decrease"
        ws.Range("R6").Value = "Greatest Total Volume"
        
        ws.Range("S3").Value = "Ticker"
        ws.Range("T3").Value = "Value"

        
        ' Variables
        StockReportTableRow = 2
        
        Dim openValue As Variant
        
        openValue = Null
        
        totalStockVolume = 0
        
        For i = 2 To lastRow
    
        
            If (ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value) Then
            
                ' Read close value for each Ticker
                closeValue = ws.Range("F" & i).Value
            
                ' Populate Ticker
                ws.Range("I" & StockReportTableRow).Value = ws.Cells(i, 1).Value
                
                ' Calculate Yearly Change
                ws.Range("J" & StockReportTableRow).Value = closeValue - openValue
                
                ' Calculate Percent Change
                ws.Range("K" & StockReportTableRow).Value = (closeValue - openValue) / openValue
                ws.Range("K" & StockReportTableRow).NumberFormat = "0.00%"
                
                ' Calculate Total Stock Volume
                totalStockVolume = totalStockVolume + ws.Range("G" & i).Value
                ws.Range("L" & StockReportTableRow).Value = totalStockVolume
                ws.Range("L" & StockReportTableRow).NumberFormat = "0"
                
                ' Update Variables
                StockReportTableRow = StockReportTableRow + 1
                openValue = Null
                totalStockVolume = 0
            
            ElseIf (IsNull(openValue)) Then
            
                ' Read openValue for each Ticker in the beginning of the year
                openValue = ws.Range("C" & i).Value
                
                ' Calculate Total Stock Volume
                totalStockVolume = totalStockVolume + ws.Range("G" & i).Value
                ws.Range("L" & StockReportTableRow).Value = totalStockVolume
                
            Else
            
                ' Calculate Total Stock Volume
                totalStockVolume = totalStockVolume + ws.Range("G" & i).Value
                ws.Range("L" & StockReportTableRow).Value = totalStockVolume
            
            
            End If
        
        Next i
        
        ' Determine last row from the Stock Report summary table
        With ws.Range("I2:L" & lastRow).CurrentRegion
            stockReportLastRow = .Rows(.Rows.Count).Row
        End With

        ' Determine Greatest Increase, Decrease, Total Volume
        greatestIncreaseValue = ws.Application.WorksheetFunction.Max(ws.Range("K2:K" & stockReportLastRow))
        
        greatestDecreaseValue = ws.Application.WorksheetFunction.Min(ws.Range("K2:K" & stockReportLastRow))
        
        greatestTotalVolume = ws.Application.WorksheetFunction.Max(ws.Range("L2:L" & stockReportLastRow))
        
        
        ' Determine the ticker symbol for the Greatest Increase, Decrease and Total Volume
        For i = 2 To stockReportLastRow
        
            If (ws.Range("K" & i).Value = greatestIncreaseValue) Then
                
                greatestIncreaseTicker = ws.Range("I" & i).Value
                
            ElseIf (ws.Range("K" & i).Value = greatestDecreaseValue) Then
                    
                greatestDecreaseTicker = ws.Range("I" & i).Value
                
            End If
            
            If (ws.Range("L" & i).Value = greatestTotalVolume) Then
            
                greatestVolumeTicker = ws.Range("I" & i).Value
            
            End If
        
        Next i
        
        ' Populate Greatest Values
        ws.Range("S4").Value = greatestIncreaseTicker
        ws.Range("T4").Value = greatestIncreaseValue
        ws.Range("T4").NumberFormat = "0.00%"
        
        ws.Range("S5").Value = greatestDecreaseTicker
        ws.Range("T5").Value = greatestDecreaseValue
        ws.Range("T5").NumberFormat = "0.00%"
        
        ws.Range("S6").Value = greatestVolumeTicker
        ws.Range("T6").Value = greatestTotalVolume
        ws.Range("T6").NumberFormat = "0"
        
        
        ' Conditional Formatting for Yearly Change and Percent Change
        Dim rg As Range
        Dim postiveCondition As FormatCondition, negativeCondition As FormatCondition
        
        Set rg = ws.Range("J2:K" & stockReportLastRow)
        rg.FormatConditions.Delete
        
        Set postiveCondition = rg.FormatConditions.Add(xlCellValue, xlGreaterEqual, "=0")
        Set negativeCondition = rg.FormatConditions.Add(xlCellValue, xlLess, "=0")
    
        'define conditional formatting to use
        With postiveCondition
        .Interior.Color = vbGreen
        End With
        
        With negativeCondition
        .Interior.Color = vbRed
        End With
    
    Next ws


End Sub

