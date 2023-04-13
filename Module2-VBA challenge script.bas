Attribute VB_Name = "Module1"
Sub Stocks():
'Create a script that loops through all data and outputs the ticker symbol, yearly change, percentage change,
    'and total stock volume
 For Each ws In Worksheets
'Declare Variables
    Dim Ticker As String
    'Dim ws As Worksheets
    
'Count how many rows there are in this worksheet
    LastRow = Cells(Rows.Count, "A").End(xlUp).Row
    MsgBox (LastRow)
       
'Insert New Columns
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly_Change"
    Cells(1, 11).Value = "Percent_Change"
    Cells(1, 12).Value = "Total_Stock_Volume"
        
    
'Set the starting points
    Total_Stock_Volume = 0
    Start = 2
    Openprice = Cells(2, 3)
    
'Loop through column A and locate the unique ticker symbols and insert them into column I
    j = 1
    For i = 2 To LastRow
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        Ticker = Cells(i, 1).Value
        Range("I" & 1 + j).Value = Ticker
        
    'Find the opening and closing price for the ticker
        Closeprice = Range("F" & i).Value
      
        
    'Subtract those two prices and insert them into their column
        'Openprice = Cells(2, 3)
        Yearly_Change = Closeprice - Openprice
        Range("J" & 1 + j).Value = Yearly_Change
        'Cells(i + 1, 3).Value = Openprice
        Openprice = Cells(i + 1, 3)
    'Divide the two prices and insert them into their column for percent changed
        Percent_Change = (Closeprice - Openprice) - 1
        Range("K" & 1 + j).Value = Percent_Change & "%"
        'Cells(i + 1, 3).Value = Openprice
        
    'Add all the values in each <vol> cell for each ticker and put it in their column
        Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
        Range("L" & 1 + j).Value = Total_Stock_Volume
        Start = i + 1
        Total_Stock_Volume = 0
        
    j = j + 1
    
    Else
         Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
        
    End If
    Next i
    
'Add conditional formatting to change the cell color of the Yearly_Change column
    For i = 2 To LastRow
    If Range("J" & 1 + i).Value >= 0 Then
        Cells(i, "J").Interior.ColorIndex = 4 'Green
    Else
        Cells(i, "J").Interior.ColorIndex = 3 'Red
    End If
    Next i
    
'Insert new cells
    Cells(2, 14).Value = "Greatest % Increase"
    Cells(3, 14).Value = "Greatest %Decrease"
    Cells(4, 14).Value = "Greatest Total Volume"
    Cells(1, 15).Value = "Ticker"
    Cells(1, 16).Value = "Value"
    
'Find greatest % increase
    Cells(2, 15).Value = WorksheetFunction.Max(Range("K2:K" & LastRow))
    Cells(2, 15).NumberFormat = "0.00%"
    increase_index = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & LastRow)), Range("K2:K" & LastRow), 0)
    Range("P2") = Cells(increase_index + 1, 9)
    
'Find greatest % decrease
    Cells(3, 15).Value = WorksheetFunction.Min(Range("K2:K" & LastRow))
    Cells(3, 15).NumberFormat = "0.00%"
    decrease_index = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & LastRow)), Range("K2:K" & LastRow), 0)
    Range("P3") = Cells(decrease_index + 1, 9)
    
'Find the greatest volume
    Cells(4, 15).Value = WorksheetFunction.Max(Range("L2:L" & LastRow))
    volume_index = WorksheetFunction.Match(WorksheetFunction.Min(Range("L2:L" & LastRow)), Range("L2:L" & LastRow), 0)
    Range("P4") = Cells(volume_index + 1, 9)
    
'Loop through all worksheets
Next ws
    
End Sub


