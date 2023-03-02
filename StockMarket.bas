Attribute VB_Name = "StockMarket"
Sub stock_market()

' Add a summary table to each worksheet that specifies the Yearly Change,
' Percent Change and Total Stock Volumn for each Ticker (or Ticker Name)
' Ticker (ticker name) will be in column I or 9
' Yearly Change will be in column J or 10
' Percent Change will be in column K or 11
' Total Stock Volume will be in column L or 12

' After this initial table is added, provide further analysis by adding
' another table to summarize data from the new table
 


' Go through each worksheet in the workbook
For Each ws In Worksheets
    
    'Set variable LastRow to contain the number of the last row in the worksheet
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Write header names for the 4 new columns across the first row of
    ' previously specified columns
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    ' Change style of Percent Change column to percent
    ws.Range("K2:K" & LastRow).Style = "Percent"
    
    
    ' Declare and initialize variables to hold values for the 4 new columns
    ' and other variables needed to calculate those values
    
    
    ' New column #1:  Ticker
    Dim ticker_name As String
    
    
    ' New column #2:  Yearly Change
    Dim yearly_change As Double
    
    ' Variables used to calculate yearly_change:  opening_price and closing_price
    
    ' Opening price for a ticker will be the value in the "open" column (C)
    ' the first time the ticker name appears in the data set
    Dim opening_price As Double
    
    ' Initialize opening_price to the value in the "open" column
    ' for the first ticker in the data set
    opening_price = ws.Range("C2")
    
    ' Closing price for a ticker will be the value in the "close" column (F)
    ' the last time the ticker name appears in the data set
    ' Note that this variable cannot be initialized because the last row
    ' for the first ticker has not yet been determined
    Dim closing_price As Double
    
    
    ' New column #3:  Percent Change
    Dim percent_change As Double
    
    
    ' New column #4:  Total Stock Volume
    Dim total_stock_volume As Double
    
    ' Initialize total_stock_volume to 0 for the first ticker
    ' total_stock_volume will be increased as each new row for the ticker is processed
    ' by adding the value in the "vol" column (G) to the running total in total_stock_volume
    total_stock_volume = 0
    
    
    ' Declare variable to track the current row being written to for new columns
    Dim summary_table_row
    
    ' Initialize summary_table_row to 2 since that will be the
    ' first row written to in the new summary table
    ' Note that this value is set to 2 to accommodate the table headers in row 1
    summary_table_row = 2
    
    
    ' Loop through all rows in the worksheet starting with row 2 (row 1 contains headers)
    ' and ending with the last row of the worksheet as already set in the LastRow variable
    ' Note that the i variable will contain the current row
    
    For i = 2 To LastRow
        
        ' If the row after the current row contains a different ticker_name (as contained
        ' in column 1), that indicates that the current row is the last row containing
        ' data for the current ticker_name and information about the current ticker_name
        ' should be determined and written to the new summary table
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            ' Set name of current ticker
            ticker_name = ws.Cells(i, 1).Value
            
            ' Add current ticker_name to summary table
            ws.Range("I" & summary_table_row) = ticker_name

            
            ' Calculate yearly_change and add it to summary table
            ' yearly_change = closing_price - opening_price
            
            ' closing_price will be the value in the "close" column (F or 6) in this last row
            ' for the ticker_name
            closing_price = ws.Cells(i, 6).Value
            
            ' calculate the yearly_change
            ' note that opening_price was set for the first ticker_name when the variable
            ' was declared and set for each subsequent ticker_name after the last data row for
            ' the previous ticker_name was processed
            yearly_change = closing_price - opening_price
            
            ' Add yearly change to summary table
            ws.Range("J" & summary_table_row) = yearly_change
            
            ' If the yearly_change is positive, then set the interior color of the cell to green (4)
            ' If the yearly_change is negative, then set the interior color of the cell to red (3)
            ' Otherwise (meaning the yearly_change was 0), do not change the interior color
            If yearly_change > 0 Then
                ws.Range("J" & summary_table_row).Interior.ColorIndex = 4
            ElseIf yearly_change < 0 Then
                ws.Range("J" & summary_table_row).Interior.ColorIndex = 3
            End If
                            
            
            ' Calculate Percent Change
            percent_change = yearly_change / opening_price
            
            ' Add percent_change to summary table
            ws.Range("K" & summary_table_row) = percent_change
            
             
            ' Add stock volume (from column G or 7) of this row to the running total stock volume
            total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
            
            ' Add total_stock_volume to summary table total
            ws.Range("L" & summary_table_row) = total_stock_volume
            
            ' Increase the summary_table_row by 1 to move to the next row of the summary table
            summary_table_row = summary_table_row + 1
            
            ' Reset the total_stock_volume back to 0 for the next ticker_name
            total_stock_volume = 0
            
        Else
        
        ' This row is NOT the last row for the current Ticker.  Increase the running total stock volume only.
          
            ' Add stock volume (from column G or 7) of this row to the running total stock volume
            total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
            
        End If
        
    Next i
    
    
    ' Create additional table to further summarize previous new table's information
    ' For this table add Greatest % Increase, Greatest % Decrease, and Greatest Total Volume
    
    ' Write row label names (Greatest %Increase, Greatest %Decrease, and Greatest Total Volume)
    ' and column label names (Ticker and Value)
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    ' Declare and set variables for Greatest % Increase
    Dim max_percent_increase As Double
    Dim max_percent_increase_index As Double
    Dim max_percent_increase_ticker As String
    max_percent_increase = WorksheetFunction.Max(ws.Range("K2:K" & LastRow))
    max_percent_increase_index = WorksheetFunction.Match(max_percent_increase, ws.Range("K2:K" & LastRow), 0)
    max_percent_increase_ticker = ws.Range("I" & (max_percent_increase_index + 1)).Value
    
    ' Declare and set variables for Greatest % Decrease
    Dim max_percent_decrease As Double
    Dim max_percent_decrease_index As Double
    Dim max_percent_decrease_ticker As String
    max_percent_decrease = WorksheetFunction.Min(ws.Range("K2:K" & LastRow))
    max_percent_decrease_index = WorksheetFunction.Match(max_percent_decrease, ws.Range("K2:K" & LastRow), 0)
    max_percent_decrease_ticker = ws.Range("I" & (max_percent_decrease_index + 1)).Value
      
    ' Declare and set variables for Greatest Total Volume
    Dim max_total_stock_volume As Double
    Dim max_total_stock_volume_index As Double
    Dim max_total_stock_volume_ticker As String
    max_total_stock_volume = WorksheetFunction.Max(ws.Range("L2:L" & LastRow))
    max_total_stock_volume_index = WorksheetFunction.Match(max_total_stock_volume, ws.Range("L2:L" & LastRow), 0)
    max_total_stock_volume_ticker = ws.Range("I" & max_percent_increase_index + 1).Value
    
    'Add new summary values to new table area on spreadsheet
    
    'Add ticker names
    ws.Range("P2").Value = max_percent_increase_ticker
    ws.Range("P3").Value = max_percent_decrease_ticker
    ws.Range("P4").Value = max_total_stock_volume_ticker
    
    'Add newly calculated values (some require changing style of cell)
    ws.Range("Q2").Style = "percent"
    ws.Range("Q2").Value = max_percent_increase
    ws.Range("Q3").Style = "percent"
    ws.Range("Q3").Value = max_percent_decrease
    ws.Range("Q4").Value = max_total_stock_volume
    
    ' Adjust autofit for all new columns added for the new table data for full readability
    ws.Range("I1:I" & LastRow).Columns.AutoFit
    ws.Range("J1:J" & LastRow).Columns.AutoFit
    ws.Range("K1:K" & LastRow).Columns.AutoFit
    ws.Range("L1:L" & LastRow).Columns.AutoFit
    ws.Range("O1:O4").Columns.AutoFit
    ws.Range("P1:P4").Columns.AutoFit
    ws.Range("Q1:Q4").Columns.AutoFit
     
Next ws
            
End Subt
