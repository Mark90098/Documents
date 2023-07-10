Sub stock_market()


For Each ws In Worksheets   'loops through all worksheets in document. Need to ad "ws." before action code
'Declaring variables
Dim ticker As String
'Dim open_price As Double   'this works as a double
'Dim close_price As String    'for some reason this variable has be be a string. Using a double doesn't work so I don't declare it at all and it works
Dim ticker_counter As Integer

'define initial variable values
ticker_counter = 2
open_price = Cells(2, 3).Value  ' need initial condition for open price because this starts for loop on row 2
total_stock_vol = 0


'Name title cells. Put at end to overwrite 1st iteration of title that appears
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "% Change"
ws.Cells(1, 12).Value = "Total Stock Vol"


'define what the lastrow of the worksheet it
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row


'Starting for loop to run down column 1 checking for a ticker change and if so, defining open and closing prices
For i = 2 To lastrow
ticker = ws.Cells(i, 1).Value 'define ticker symbol

stock_vol = ws.Cells(i, 7).Value
total_stock_vol = stock_vol + total_stock_vol  'add all stock volumes to get the sum for each ticker symbol

    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1) Then
    
    close_price = ws.Cells(i, 6).Value   'defines close price
    
    'Display the 1st 3 desired column information
    'Cells(ticker_counter, 14).Value = close_price 'I want each row in column 10 to be the next company symbol (ticker).
    ws.Cells(ticker_counter, 9).Value = ticker
   ws.Cells(ticker_counter, 10).Value = (close_price - open_price)
    ws.Cells(ticker_counter, 11).Value = (close_price - open_price) / open_price * 100
    
    
    
    'Conditional statements to change cell colors green or red
        If ws.Cells(ticker_counter, 10).Value > 0 Then
        ws.Cells(ticker_counter, 10).Interior.ColorIndex = 4 'green
        
        ElseIf ws.Cells(ticker_counter, 10).Value < 0 Then
        ws.Cells(ticker_counter, 10).Interior.ColorIndex = 3 'red
        
        End If
        
        
    
    ws.Cells(ticker_counter, 12).Value = total_stock_vol  '4th desired column
    total_stock_vol = 0 'reset total stock volume for next ticker symbol
    
    ticker_counter = ticker_counter + 1   'this counts the # of changes of company symbol

   
    open_price = ws.Cells(i + 1, 3).Value 'defines open price.This has to after ticker_counter because the open and close price needs to both be from the same ticker symbol
    'Cells(ticker_counter, 13).Value = open_price
    
    End If



Next i

Next ws

End Sub
