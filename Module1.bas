Attribute VB_Name = "Module1"
Sub Stockmetrics()
'MsgBox "Processing sheet: " & ws.Name
'Dim ws As Worksheet
'For Each ws In ThisWorkbook.Worksheets

'Variable to hold column name
Dim column As Integer
column = 1

'Variable to increment the Ticker symbols in column I
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

' Variable to Sum Stock Volume for each ticker
Dim Total_Stock_Volume As Double
Total_Stock_Volume = 0

'Variables for calculating Quarterly Change/Percent Change
Dim Open_Price As Double
Dim End_Price As Double
Dim Quarterly_Change As Double
Dim Percent_Change As Double

'Variables for calculating Greatest Increase %, Greatest Decrease %, Greatest Volume
Dim Greatest_Increase As Double
Dim Greatest_Decrease As Double
Dim Greatest_Volume As Double
Dim GI_row As Integer
Dim GD_row As Integer
Dim GV_row As Integer


'Loop Counter
Dim loop_count As Integer
counter = 0

'Dim i As Integer
Dim Ticker_Symbol As String


'Find Last Row
lastRow = Range("A" & Rows.Count).End(xlUp).Row
'lastRow = Cells(Rows.Count, 1).End(x1Up).Row (This syntax doesn't work for some reason)

'Create Column Headers for Ticker,etc.
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Quarterly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

'Create Column Headers for Greatest Increase, Decrease, etc.
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"

'Format Percent_Change column
Range("K:K").NumberFormat = "0.00%"
'Initialize Open Price
Open_Price = Cells(2, 3).Value
Range("J" & Summary_Table_Row).Value = Open_Price

'Loop through rows in the column
For i = 2 To lastRow
'Searches for when the value of the next cell is different than the current cell
    If Cells(i + 1, column).Value <> Cells(i, column).Value Then
    
'Write Ticker Symbol to Column I
        Ticker_Symbol = Cells(i, 1).Value
        Range("I" & Summary_Table_Row).Value = Ticker_Symbol
        
'Write Total Stock Volume to Column L
        Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
        
'Calculate Open Price
        Open_Price = Cells(i - loop_count, 3).Value
        
'Calculate End_Price
        End_Price = Cells(i, 6).Value
             
'Write Quarterly Change to Colum J
        Quarterly_Change = End_Price - Open_Price
        
        Range("J" & Summary_Table_Row).Value = Quarterly_Change
        
'Format Quarterly Change fill color
                If (Quarterly_Change > 0) Then
                    Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
            
                ElseIf (Quarterly_Change < 0) Then
                    Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
            
                End If
        
        
'Write Percent Change to Column K
        Percent_Change = Quarterly_Change / Open_Price
        Range("K" & Summary_Table_Row).Value = Percent_Change
        
'Increment Summary_Table_Row
        
        Summary_Table_Row = Summary_Table_Row + 1
        
'Reset loop_count to zero
        loop_count = 0
'Reset Total_Stock_Volume to zero
        Total_Stock_Volume = 0
    
    Else
'Calculate Total Stock Volume
        Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
        
'Increment loop_count variable
        loop_count = loop_count + 1
   
    End If
Next i

'Find last row of new columns
lastRow_newcol = Range("I" & Rows.Count).End(xlUp).Row

'Find Greatest % Increase, Greatest % Decrease, Greatest Total Volume
 
'Dim GI_row As Integer
'Dim GD_row As Integer
'Dim GV_row As Integer
    
    Greatest_Increase = WorksheetFunction.Max(Range("K2:K" & lastRow_newcol))
    Cells(2, 17).Value = Greatest_Increase
    GI_row = Application.WorksheetFunction.Match(Greatest_Increase, Range("K2:K" & lastRow_newcol), 0) + 1
    Cells(2, 16).Value = Cells(GI_row, "I").Value
    
    Greatest_Decrease = WorksheetFunction.Min(Range("K2:K" & lastRow_newcol))
    Cells(3, 17).Value = Greatest_Decrease
    GD_row = Application.WorksheetFunction.Match(Greatest_Decrease, Range("K2:K" & lastRow_newcol), 0) + 1
    Cells(3, 16).Value = Cells(GD_row, "I").Value
    
    Greatest_Volume = WorksheetFunction.Max(Range("L2:L" & lastRow_newcol))
    Cells(4, 17).Value = Greatest_Volume
    GV_row = Application.WorksheetFunction.Match(Greatest_Volume, Range("L2:L" & lastRow_newcol), 0) + 1
    Cells(4, 16).Value = Cells(GV_row, "I").Value
    
'Next ws

End Sub
