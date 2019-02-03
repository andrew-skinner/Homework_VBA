Sub stock_counter()

'Set an initial variable for each ticker
Dim ticker As String

'Set an initial variable for holding the total per ticker symbol
Dim ticker_total As Double
ticker_total = 0

'Keep track of the location for each ticker in the summary table 
Dim summary_table_row As Double
Summary_table_row = 2

'Set header row for summary table
Cells(1,9).value = ("Ticker")
Cells(1,10).value = ("Ticker Volume")

'Loop through all ticker symbols
For t = 2 to 1000000

    'Check if we are still within the same ticker, if not then
    If Cells(t+1,1).Value <> Cells(t,1).Value then

        'Set the ticker name
        ticker = Cells(t, 1).Value

        'Add to the ticker total
        ticker_total = ticker_total + Cells(t,7).Value

        'Print the ticke in the summary table
        Range("I" & Summary_table_row).Value = ticker

        'Print the ticker total in the Summary Table
        Range("J" & Summary_table_row).Value = ticker_total

        'Add one to the summary table row
        Summary_table_row = summary_table_row + 1

        'Reset the ticker total
        ticker_total = 0

        'If the cell immediately following a row is the same brand
        Else

        'Add to the ticker total
        ticker_total = ticker_total + cells(t,7).Value

        End If
    
    Next t

End Sub

