### VBA-challenge

## General Information
In this Challenge I used VBA script to analyze generated stock market data for three years. The VBA script I developed allows someone to loop through stock marekt data to find relevent information regarding stocks including Ticker Names, Yearly Change, Percent Change, and Total Stock Volume.

## Technologies Used
* VBA Script


## Process
1. I started off with a for loop to loop through all the worksheets.
2. Created and defined all of my variables
3. Created a row counter to know where the last row in each sheet is located
4. created a loop with a conditional to find Total Stock Volume
```
For i = 2 To LastRow
    'Conditional to add total stock volume for each particular ticker name
    If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
         Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
        
        
    Else
        Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
        Summary_Table_row = Summary_Table_row + 1
        ws.Cells(Summary_Table_row, 12).Value = Total_Stock_Volume
    End If
Next i
```
5. Created a loop with a conditional to list Ticker Names
```
For c = 2 To LastRow
    'Conditional to move from one ticker to the next
    If ws.Cells(c + 1, 1).Value <> ws.Cells(c, 1).Value Then
        Ticker_Name = ws.Cells(c, 1).Value
        Summary_Table_row = Summary_Table_row + 1
        ws.Range("I" & Summary_Table_row).Value = Ticker_Name
    End If
Next c
```
6. Created a loop with a conditional to calulate Yearly Change and Percent Change. This also Formated Percent Change as a percent and colored Increase/decrease yearly change with Green/Red respectively
```
For s = 2 To LastRow
        'Conditonal to 1. find yearly change then 2. Use yearly change to find percent change
        If ws.Cells(s + 1, 1).Value <> ws.Cells(s, 1).Value Then
        
            Summary_Table_row = Summary_Table_row + 1
            Stock_Close = ws.Cells(s, 6).Value
            Yearly_Change = Stock_Close - Stock_Open
            ws.Cells(Summary_Table_row, 10).Value = Yearly_Change
            'Sets the cell colors for the yearly change column
                If ws.Cells(Summary_Table_row, 10).Value > 0 Then
                    ws.Cells(Summary_Table_row, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(Summary_Table_row, 10).Interior.ColorIndex = 3
                End If
        'Finds the percent change in each ticker
            ws.Cells(Summary_Table_row, 11).Value = (ws.Cells(Summary_Table_row, 10).Value / Stock_Open)
            ws.Cells(Summary_Table_row, 11).NumberFormat = "0.00%"
        'Resets stock open for next ticker
            Stock_Open = ws.Cells(s + 1, 3).Value
        End If
```        
  
 
