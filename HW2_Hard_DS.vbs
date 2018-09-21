'HW2 ## Moderate
'* Create a script that will loop through all the stocks and take the following info.
'   * Yearly change from what the stock opened the year at to what the closing price was.
'   * The percent change from the what it opened the year at to what it closed.
'   * The total Volume of the stock
'   * Ticker symbol
'* You should also have conditional formatting that will highlight positive change in green and negative change in red.
'* The result should look as follows.![moderate_solution](Images/moderate_solution.png)

'#####HARD#####
'Your solution will include everything from the moderate challenge.
'Your solution will also be able to locate the stock with the
'-->"Greatest % increase",
'-->"Greatest % Decrease" and
'-->"Greatest total volume".
'--------
'SOLUTION
'--------

'Start with the code for one sheet
Sub tickerloop()
    'Setting up the stage

        'Set a variable for holding the ticker name, the column of interest
        Dim tickername As String
    
        'Set a varable for holding a total count on the total volume of trade
        Dim tickervolume As Double
        tickervolume = 0

        'Keep track of the location for each ticker name in the summary table
        Dim summary_ticker_row As Integer
        summary_ticker_row = 2
        
        'Note: Yearly Change is simply the difference: (Close Price at the end of a trading year - Open Price at the beginning of the trading year)
        'Percent change is a simple percent change -->((Close - Open)/Open)*100
        Dim open_price As Double
        'Set initial open_price. Other opening prices will be determined in the conditional loop.
        open_price = Cells(2, 3).Value
        
        Dim close_price As Double
        Dim yearly_change As Double
        Dim percent_change As Double

        'Label the Summary Table headers
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"

        'Count the number of rows in the first column.
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row

        'Loop through the rows by the ticker names
        'Make sure that the ticker names are sorted and are alpha-numeric/string variables.
        'Do a manual check.

        For i = 2 To lastrow

            'Searches for when the value of the next cell is different than that of the current cell
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
              'Set the ticker name
              tickername = Cells(i, 1).Value

              'Add the volume of trade
              tickervolume = tickervolume + Cells(i, 7).Value

              'Print the ticker name in the summary table
              Range("I" & summary_ticker_row).Value = tickername

              'Print the trade volume for each ticker in the summary table
              Range("L" & summary_ticker_row).Value = tickervolume

              'Now collect information about closing price
              close_price = Cells(i, 6).Value

              'Calculate yearly change
               yearly_change = (close_price - open_price)
              
              'Print the yearly change for each ticker in the summary table
              Range("J" & summary_ticker_row).Value = yearly_change

              'Check for the non-divisibilty condition when calculating the percent change
                If open_price = 0 Then
                    percent_change = 0
                Else
                    percent_change = yearly_change / open_price
                End If

              'Print the yearly change for each ticker in the summary table
              Range("K" & summary_ticker_row).Value = percent_change
              Range("K" & summary_ticker_row).NumberFormat = "0.00%"
   
              'Reset the row counter. Add one to the summary_ticker_row
              summary_ticker_row = summary_ticker_row + 1

              'Reset volume of trade to zero
              tickervolume = 0

              'Reset the opening price
              open_price = Cells(i + 1, 3)
            
            Else
              
               'Add the volume of trade
              tickervolume = tickervolume + Cells(i, 7).Value

            
            End If
        
        Next i

    'Conditional formatting that will highlight positive change in green and negative change in red
    'First find the last row of the summary table

    lastrow_summary_table = Cells(Rows.Count, 9).End(xlUp).Row
    
    'Color code yearly change
        For i = 2 To lastrow_summary_table
            If Cells(i, 10).Value > 0 Then
                Cells(i, 10).Interior.ColorIndex = 10
            Else
                Cells(i, 10).Interior.ColorIndex = 3
            End If
        Next i

    'Highlight the stock price changes
    'First label the cells according to the sample .png provided in the assignment

        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"

    'Determine the max and min values in column "Percent Change" and just max in column "Total Stock Volume"
    'Then collect the ticker name, and the corresponding values for the percent change and total volume of trade for that ticker
    '
        For i = 2 To lastrow_summary_table
            'Find the maximum percent change
            If Cells(i, 11).Value = Application.WorksheetFunction.Max(Range("K2:K" & lastrow_summary_table)) Then
                Cells(2, 16).Value = Cells(i, 9).Value
                Cells(2, 17).Value = Cells(i, 11).Value
                Cells(2, 17).NumberFormat = "0.00%"

            'Find the minimum percent change
            ElseIf Cells(i, 11).Value = Application.WorksheetFunction.Min(Range("K2:K" & lastrow_summary_table)) Then
                Cells(3, 16).Value = Cells(i, 9).Value
                Cells(3, 17).Value = Cells(i, 11).Value
                Cells(3, 17).NumberFormat = "0.00%"
            
            'Find the maximum volume of trade
            ElseIf Cells(i, 12).Value = Application.WorksheetFunction.Max(Range("L2:L" & lastrow_summary_table)) Then
                Cells(4, 16).Value = Cells(i, 9).Value
                Cells(4, 17).Value = Cells(i, 12).Value
            
            End If
        
        Next i
        
End Sub

