VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub VBA_Challenge()
'Loop through all the worksheets
For Each ws In Worksheets

Dim WorkesheetName As String
Dim LastRow As Long
LastRow = ws.Range("A1").End(xlDown).Row
Worksheetname = ws.Name
'MsgBox Worksheetname

'Setting up all Variables

Dim Ticker_Name As String
Dim Stock_Open As Double
Dim Stock_Close As Double
Dim Total_Stock_Volume As LongLong
Dim Percent_Change As Double
Dim Yearly_Change As Double

'Define variables
Stock_Open = ws.Cells(2, 3).Value
Stock_Close = 0
Total_Stock_Volume = 0

'tracks row count and helps define if the loop moves from one ticker to the next.
Dim Summary_Table_row As Integer
 Summary_Table_row = 1
'Sets up headers for the new columns
Header = VBA.Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")
ws.Range("I1:L1").Value = Header


'Loop for Total Stock Volume
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


'Reset Summary Table Row
Summary_Table_row = 1
'Loop for Ticker names
For c = 2 To LastRow
    'Conditional to move from one ticker to the next
    If ws.Cells(c + 1, 1).Value <> ws.Cells(c, 1).Value Then
        Ticker_Name = ws.Cells(c, 1).Value
        Summary_Table_row = Summary_Table_row + 1
        ws.Range("I" & Summary_Table_row).Value = Ticker_Name
    End If
Next c



'Reset Summary Table Row
Summary_Table_row = 1
'Redefines Stock open
Stock_Open = ws.Cells(2, 3).Value
'Loop for Yearly Change/Percent change
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
        
Next s


Next ws
End Sub

