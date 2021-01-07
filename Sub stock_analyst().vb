Sub stock_analyst()
'Stock Analyst Function
'Pramod Philip

'To fill the new ticker column, set up a for loop to move through each row and compare the
'current row value with the previous row value. If they are the same, it should continue parsing.
'If they are different, the previous ticker symbol should be saved in the new Ticker column. The ticker variable used to
'reference the row of the Ticker column is then increased by one, and the for loop continues.

'For the Yearly Change column, a yearly change variable will be initialized with
'the difference between the opening and closing price of the stock on the first listed stock
'and will continually increased by the yearly change of the following listed stocks provided that
'these stocks match in ticker symbol. If the ticker symbols for two consectively listed stocks are different,
'the yearly change variable will be listed in the Yearly Change column next to the respective ticker symbol
'to represent the total yearly change for that given stock type. The yearly change variable is then reset to
'equal the yearly change of the stock on the current row in the for loop.
'If the total yearly change for a given stock is positive, its cell is filled green.
'Contrarily, if the total yearly change for a given stock is negative, its cell is filled red.

'For the Percent Change column, a percent change variable will be initialized with
'the percentage difference between the opening and closing price of the stock on the first listed stock
'and will continually increased by the percent change of the following listed stocks provided that
'these stocks match in ticker symbol. If the ticker symbols for two consectively listed stocks are different,
'the percent change variable will be listed in the Percent Change column next to the respective ticker symbol
'to represent the total percent change for that given stock type. The percent change variable is then reset to
'equal the percent change of the stock on the current row in the for loop.

'For the Total Stock Volume column, a stock volume variable is initialized with the stock volume
'of the first stock listed. Afterwards, this variable will be continually increased by the stock volumes
'of the stocks that follow provided that these stocks all share the same ticker symbol.
'Should a stock happen to have a different ticker symbol from those previous, the stock volume will be listed in
'the corresponding row in the Total Stock Volume column.
'The stock volume variable is then reset with the value of the stock volume of the current row in the for loop.

'Declare worksheet loop counter variable
Dim ws As Worksheet

'Set up worksheet loop
For Each ws In ActiveWorkbook.Worksheets

'Declares script variables

'Variable for for loop
Dim i As Long
'Variable that keeps track of the row count for the new Ticker column
'This variable is reused for the Yearly Change, Percent Change, and Total Stock Volume
Dim tick_count As Integer
'Variable that holds Ticker symbols
Dim ticker As String
'Variable that contains the number of used rows in the worksheet
Dim row_count As Long
'Opening price at beginning of year for stock
Dim opn As Double
'Closing price at end of year for stock
Dim clse As Double
'Variable that keeps track of the yearly change between opening and closing stock prices
Dim change As Double
'Variable that keeps track of the percent change
Dim per_change As Double
'Variable that keeps track of total stock volume for each stock
Dim vol As Double
'Variable that holds the greatest percent increase overall
Dim max_per_i As Double
'Variable that holds the ticker symbol with the greatest percent increase overall
Dim max_per_i_tick As String
'Variable that holds the greatest percent decrease overall
Dim max_per_d As Double
'Variable that holds the ticker symbol with the greatest percent decrease overall
Dim max_per_d_tick As String
'Variable that holds the maximum total stock volume overall
Dim max_vol As Double
'Variable that holds the ticket symbol with the largest overall stock volume
Dim max_vol_tick As String
'Variable that holds the number of rows in the second Ticker column (column I)
Dim sec_row_count As Integer

'Initializes the first ticker symbol as a starting point for the for loop to check
ticker = ws.Cells(2, 1).Value

'Establishes the number of rows as a maximum boundary for the loop
row_count = 1 + ws.Cells(Rows.Count, "A").End(xlUp).Row

'Initializes the ticker variable to save the position of each new ticker symbol for the new
'Ticker column
tick_count = 2

'Initializes the variable for yearly change between opening and closing price in row 2
change = ws.Range("F2").Value - ws.Range("C2").Value

'Initializes the variable for percent change between opening and closing price
per_change = (ws.Range("F2").Value / ws.Range("C2").Value) - 1

'Initializes the variable for stock volume
vol = ws.Range("G2").Value

'Creates header for new Ticker column
ws.Range("I1").Value = "Ticker"

'Creates header for new Yearly Change column
ws.Range("J1").Value = "Yearly Change"

'Creates header for new Percent Change column
ws.Range("K1").Value = "Percent Change"

'Creates header for new Total Stock Volume column
ws.Range("L1").Value = "Total Stock Volume"

'Initializes the beginning of year opening price variable
opn = ws.Cells(2, 3).Value

'Sets up for loop which loops through every row on the active sheet, starting from 3rd row
For i = 2 To (row_count - 1)

'If the current ticker is different from the previous ticker, the previous
'ticker value will be saved in the new Ticker column in column I
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        ws.Cells(tick_count, 9).Value = ws.Cells(i - 1, 1).Value
        
        'Setting the end year closing price for a stock
        clse = ws.Cells(i, 6).Value

        'The yearly change variable is reset to equal the first yearly change for the next stock type
        change = clse - opn
        
        'The yearly change for that particular stock will be saved in column 10
        ws.Cells(tick_count, 10).Value = change
        
        'Sets condition if the opening price for a stock equals 0
        If opn = 0 Then
        
            'The percent change for that particular stock will be saved in column 11
            ws.Cells(tick_count, 11).Value = FormatPercent(0)
        
        'Sets else condition if the opening price is > than 0
        Else
        
            'The percent change variable is reset to equal the first percent change for the next stock type
            per_change = (clse / opn) - 1
            'The percent change for that particular stock will be saved in column 11
            ws.Cells(tick_count, 11).Value = FormatPercent(per_change)
            
        End If
        
    'The percent change for that particular stock will be saved in column 11
    ws.Cells(tick_count, 11).Value = FormatPercent(per_change)
        
        'Resets opening price as the opening price of next stock
        opn = ws.Cells(i + 1, 3).Value
        
        'Determines if percent change is positive or negative
        If change >= 0 Then
            'If positive, assigns a green cell fill
            ws.Cells(tick_count, 10).Interior.ColorIndex = 4
        Else
            'If negative, assigns a red cell fill
            ws.Cells(tick_count, 10).Interior.ColorIndex = 3
        End If
    
    'The total stock volume for that particular stock will be saved in column 12
        ws.Cells(tick_count, 12).Value = vol
    'The stock volume variable is reset to equal the first list stock volume for the next stock type
        vol = ws.Cells(i + 1, 7).Value
    'The row count for the new columns is increased by 1
        tick_count = tick_count + 1
        
    Else
    'Stock volume is increased by the stock volume value on the ith row assuming
    'the two ticker symbols on the current and previous row are equal
        vol = vol + ws.Cells(i + 1, 7).Value
        
    End If

'Processes first inner loop
Next i

'Header for Greatest % Increase
ws.Range("O2").Value = "Greatest % Increase"
'Header for Greatest % Decrease
ws.Range("O3").Value = "Greatest % Decrease"
'Header for Greatest Total Volume
ws.Range("O4").Value = "Greatest Total Volume"
'Header for corresponding ticker symbols of previous three headers
ws.Range("P1").Value = "Ticker"
'Header for corresponding values of previous three headers
ws.Range("Q1").Value = "Value"

'Initializes the sec_row_count variable
sec_row_count = 1 + ws.Cells(Rows.Count, "I").End(xlUp).Row

'Initializes the greatest percent  increase variable
max_per_i = ws.Range("K2").Value

'Initializes the greatest percent decrease variable
max_per_d = ws.Range("K2").Value

'Initializes the greatest stock volume variable
max_vol = ws.Range("L2").Value

'Begins the second innter loop that parses through columns I through L
For i = 3 To sec_row_count

    'Checks to see if percent change on row i is larger than the current greatest percent increase
    If ws.Cells(i, 11).Value > max_per_i Then
    'If yes, sets percent change on row i as new greatest percent increase
        max_per_i = ws.Cells(i, 11).Value
    'Grabs ticker symbol corresponding to percent change on row i
        max_per_i_tick = ws.Cells(i, 9).Value
    'Checks to see if percent change on row i is smaller than the current greatest percent decrease
    ElseIf ws.Cells(i, 11).Value < max_per_d Then
    'If yes, sets percent change on row i as new greatest percent decrease
        max_per_d = ws.Cells(i, 11).Value
    'Grabs ticker symbol corresponding to percent change on row i
        max_per_d_tick = ws.Cells(i, 9).Value
    End If
        
    'Checks to see if total stock volume on row i is
    'larger than the current total stock volume
    If ws.Cells(i, 12).Value > max_vol Then
    'If yes, sets stock volume on row i as new greatest total stock volume
        max_vol = ws.Cells(i, 12).Value
    'Grabs ticker symbol corresponding to stock volume on row i
        max_vol_tick = ws.Cells(i, 9).Value
    End If

'Processes second inner for loop
Next i


'Posts the ticker symbol of greatest percent increase to P2
ws.Range("P2").Value = max_per_i_tick
'Posts the value of greatest percent increase to Q2
ws.Range("Q2").Value = FormatPercent(max_per_i)
'Posts the ticker symbol of greatest percent decrease to P3
ws.Range("P3").Value = max_per_d_tick
'Posts the value of greatest percent decrease to Q3
ws.Range("Q3").Value = FormatPercent(max_per_d)
'Posts the ticker symbol of largest total stock volume to P4
ws.Range("P4").Value = max_vol_tick
'Posts the value of largest total stock volume to Q4
ws.Range("Q4").Value = max_vol

Next ws

End Sub
