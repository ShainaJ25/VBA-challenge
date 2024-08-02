Attribute VB_Name = "Module1"
Sub Multiple_year_stock_data_solution()

    'loop through all worksheets.
    For Each ws In Worksheets

    'set an initial variable for holding the ticker symbol
    Dim Ticker_Symbol As String

    'set an initial variable for holding the total stock volume
    Dim Ticker_Volume As Double
    Ticker_Volume = 0
    
    'keep track of the location for each ticker symbol in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    'define more variables
    Dim Open_Price As Double
    
    Open_Price = ws.Cells(2, 3).Value
    
    Dim Close_Price As Double
    Dim Quarterly_Change As Double
    Dim Percent_Change As Double
    

    'set the summary table names
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Quarterly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    
    'Counts the number of rows
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'loop through all ticker symbols
    For i = 2 To lastrow

    'Check if we are still within the same tickery symbol, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
   
        'set the ticker symbol
       Ticker_Symbol = ws.Cells(i, 1).Value

        'add to the stock volume total
       Ticker_Volume = Ticker_Volume + ws.Cells(i, 7).Value
    
        'print the ticker symbol in the Summary Table
        ws.Range("I" & Summary_Table_Row).Value = Ticker_Symbol
    
         'print the stock volume amount to the summary table
        ws.Range("L" & Summary_Table_Row).Value = Ticker_Volume
    
         'label close price
        Close_Price = ws.Cells(i, 6).Value
     
         'find quarterly change
        Quarterly_Change = (Close_Price - Open_Price)
 
        'print quarterly change for each ticker symbol in the summary table
        ws.Range("J" & Summary_Table_Row).Value = Quarterly_Change
    
            'find percent change
             If Open_Price = 0 Then
                Percent_Change = 0
                
            Else
            
            Percent_Change = (Quarterly_Change / Open_Price)
            
            End If
        
        
         'print percent change for each tickery symbol in the summary table
         ws.Range("K" & Summary_Table_Row).Value = Percent_Change
         ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
     
         'add one to the summary table row
         Summary_Table_Row = Summary_Table_Row + 1
    
         'reset volume of stock total to 0
         Ticker_Volume = 0

         'reset the opening price
         Open_Price = ws.Cells(i + 1, 3)

     'If the cell immediately following a row is the same brand...
     Else

         'add to the volume of stock total
         Ticker_Volume = Ticker_Volume + ws.Cells(i, 7).Value
     
      End If
    
Next i

 'adding colors

 'find last row

 lastrow_Summary_Table = ws.Cells(Rows.Count, 9).End(xlUp).Row

    For i = 2 To lastrow_Summary_Table
    
            If ws.Cells(i, 10).Value > 0 Then
            
               ws.Cells(i, 10).Interior.ColorIndex = 4
               
            ElseIf ws.Cells(i, 10).Value < 0 Then
            
               ws.Cells(i, 10).Interior.ColorIndex = 3
            
            Else
            
               ws.Cells(i, 10).Interior.ColorIndex = xlNone
              
            End If
            
    Next i
    
   'add functionality to script

   'name cells

   ws.Cells(2, 15).Value = "Greatest % Increase"
   ws.Cells(3, 15).Value = "Greatest % Decrease"
   ws.Cells(4, 15).Value = "Greatest Total Volume"
   ws.Cells(1, 16).Value = "Ticker"
   ws.Cells(1, 17).Value = "Value"


   For i = 2 To lastrow_Summary_Table
    
    'greatest percent increase
    If ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & lastrow_Summary_Table)) Then
       ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
       ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
       ws.Cells(2, 17).NumberFormat = "0.00%"
     
     
    'greatest percent decrease
    ElseIf ws.Cells(i, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & lastrow_Summary_Table)) Then
             ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
             ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
             ws.Cells(3, 17).NumberFormat = "0.00%"
     
     'greatest volume of stock total
     ElseIf ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastrow_Summary_Table)) Then
              ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
              ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
  
     End If
  
  Next i
     
  Next ws

End Sub
