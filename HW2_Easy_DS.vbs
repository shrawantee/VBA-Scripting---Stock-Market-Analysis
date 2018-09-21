'HW2  -### Easy
'Instructions

'* Create a script that will loop through each year of stock data and grab the total amount of volume each stock had over the year.
'* You will also need to display the ticker symbol to coincide with the total volume.
'* Your result should look as follows (note: all solution images are for 2015 data).
'![easy_solution](Images/easy_solution.png)

'--------
'SOLUTION
'--------

'Start with the code for one sheet
    
sub tickerloop()

    'Setting up the stage 

        'Set a variable for holding the ticker name, the column of interest
        Dim tickername As String
    
        'Set a varable for holding a total count on the total volume of trade 
        Dim tickervolume As Double
        tickervolume = 0

        'Keep track of the location for each ticker name in the summary table
        Dim summary_ticker_row As Integer
        summary_ticker_row = 2

        'Label the Summary Table headers
        Cells(1,9).Value = "Ticker"
        Cells(1,10).Value = "Total Stock Volume"

        'Count the number of rows in the first column. 
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row

        'Loop through the rows by the ticker names
        'Make sure that all the ticker names are sorted and are alpha-numeric/string variables. 
        'Do a manual check. 
            
        For i = 2 to lastrow

            'Searches for when the value of the next cell is different than that of the current cell
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
                'Set the ticker name
                tickername = Cells(i,1).Value 

                'Add the volume of trade
                tickervolume = tickervolume + Cells(i, 7).Value

                'Print the ticker name in the summary table
                Range("I" & summary_ticker_row).Value = tickername

                'Print the trade volume for each ticker in the summary table
                Range("J" & summary_ticker_row).Value = tickervolume

                'Add one to the summary_ticker_row 
                summary_ticker_row = summary_ticker_row + 1

                'Reset tickervolume to zero
                tickervolume = 0

            Else
              
                'Add the volume of trade
                tickervolume = tickervolume + Cells(i, 7).Value

            End if
        
        Next i

End Sub

