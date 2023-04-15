Attribute VB_Name = "Module1"
Sub Stocks():
'Create a script that loops through all data and outputs the ticker symbol, yearly change, percentage change,
    'and total stock volume
 For Each ws In Worksheets
    ws.Activate
    
'Declare Variables
    Dim Ticker As String
    'Dim ws As Worksheets
    
'Count how many rows there are in this worksheet
    LastRow = Cells(Rows.Count, "A").End(xlUp).Row
       
'Insert New Columns
    Cells(1, "I").Value = "Ticker"
    Cells(1, "J").Value = "Yearly_Change"
    Cells(1, "K").Value = "Percent_Change"
    Cells(1, "L").Value = "Total_Stock_Volume"
        
    
'Set the starting points
    Total_Stock_Volume = 0
    Start = 2
    Openprice_pointer = 2
    'Openprice = Cells(2, 3)
    
'Loop through column A and locate the unique ticker symbols and insert them into column I
    
    For i = 2 To LastRow
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            Total_Stock_Volume = Total_Stock_Volume + Cells(i, "G").Value
            Ticker = Cells(i, 1).Value
            Openprice = Cells(Openprice_pointer, "C").Value
            
            
        'Find the opening and closing price for the ticker
            Closeprice = Range("F" & i).Value
          
            
        'Subtract those two prices and insert them into their column
            Yearly_Change = Closeprice - Openprice
            
           
        'Divide the two prices and insert them into their column for percent changed
            Percent_change = (Closeprice - Openprice) - 1
            
            Cells(Start, "I").Value = Ticker
            Cells(Start, "J").Value = Yearly_Change
            Cells(Start, "K").Value = "%" & Percent_change
            Cells(Start, "L").Value = Total_Stock_Volume
            
        If Yearly_Change >= 0 Then
            Cells(Start, "J").Interior.ColorIndex = 4 'Green
        Else
            Cells(Start, "J").Interior.ColorIndex = 3 'Red
        End If
            
            
            'Cells(i + 1, 3).Value = Openprice
            
        'Add all the values in each <vol> cell for each ticker and put it in their column
            Total_Stock_Volume = 0
            Start = Start + 1
            Openprice_pointer = i + 1
            
    
        Else
            Total_Stock_Volume = Total_Stock_Volume + Cells(i, "G").Value
        
        End If
    Next i
    
    
'Insert new cells
    Cells(2, "N").Value = "Greatest % Increase"
    Cells(3, "N").Value = "Greatest %Decrease"
    Cells(4, "N").Value = "Greatest Total Volume"
    Cells(1, "O").Value = "Ticker"
    Cells(1, "P").Value = "Value"
    
'Find greatest % increase
    Cells(2, "P").Value = WorksheetFunction.Max(Range("K2:K" & LastRow))
    Cells(2, "P").NumberFormat = "0.00%"
    increase_index = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & LastRow)), Range("K2:K" & LastRow), 0)
    Range("O2") = Cells(increase_index + 1, "I")
    
'Find greatest % decrease
    Cells(3, "P").Value = WorksheetFunction.Min(Range("K2:K" & LastRow))
    Cells(3, "P").NumberFormat = "0.00%"
    decrease_index = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & LastRow)), Range("K2:K" & LastRow), 0)
    Range("O3") = Cells(decrease_index + 1, "I")
    
'Find the greatest volume
    Cells(4, "P").Value = WorksheetFunction.Max(Range("L2:L" & LastRow))
    volume_index = WorksheetFunction.Match(WorksheetFunction.Min(Range("L2:L" & LastRow)), Range("L2:L" & LastRow), 0)
    Range("O4") = Cells(volume_index + 1, "I")
    
'Loop through all worksheets
Next ws
MsgBox ("Done")
    
End Sub


