Sub Ticker()

'sort each worksheet by Columns A and Columns B
For Each ws In Worksheets
ws.Activate
Columns.Sort Key1:=Columns("A"), Order1:=xlAscending, Key2:=Columns("B"), Order2:=xlAscending, Header:=xlYes
Next ws

'apply macro to each worksheet
For Each ws In Worksheets

'define variables
Dim Ticker As String
Dim Ticker_Volume As Double
Dim Summary_Table_Row As Integer
Dim Open_Price As Double
Dim Close_Price As Double
Dim Yearly_Change As Double
Dim Percent_Change As Double

'place headers for summary table
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

'create Ticker Volume counter
Ticker_Volume = 0

'create summary row counter to place data in
Summary_Table_Row = 2

'set Open Price
Open_Price = ws.Cells(2, 3).Value

'find last row of dataset
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'begin for loop to search data
For i = 2 To lastrow

    'if the next value in column does not match the previous value then do the following...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        'assign the Ticker variable the value of the cell in column A that does not match the next value
        Ticker = ws.Cells(i, 1).Value
        
        'place the Ticker value into the summary table column I
        ws.Range("I" & Summary_Table_Row).Value = Ticker
        
        'assign the Close_Price variable the value of the close price in column F that does not match the next value
        Close_Price = ws.Cells(i, 6).Value
        
        'determine the yearly change value by substracting the close price from the open price
        Yearly_Change = (Close_Price - Open_Price)
        
        'place the yearly change value in the summary table
        ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
        
        'determine the percent change by dividing the yearly change from the open price
        Percent_Change = Yearly_Change / Open_Price
        
        'place the percent change in the summary table and format to percentage
        ws.Range("K" & Summary_Table_Row).Value = Percent_Change
            ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
        
        'add the volume of the ticker for the first row that does not match to the Ticker Volume counter
        Ticker_Volume = Ticker_Volume + ws.Cells(i, 7).Value
        
        'place the value of the Ticker Volume counter into column L of the summary table
        ws.Range("L" & Summary_Table_Row).Value = Ticker_Volume
        
        'add another row to the summary table
        Summary_Table_Row = Summary_Table_Row + 1
        
        'reset the Ticker Volume
        Ticker_Volume = 0
        
        'set the Open Price to the value in column C of the next row that does not match
        Open_Price = ws.Cells(i + 1, 3)
    
    'if the next value matches the previous value then add the value in column G to the Ticker Volume counter
    Else
        Ticker_Volume = Ticker_Volume + ws.Cells(i, 7).Value
    
    End If

Next i

'find last row of summary table
lastrow_Summary = ws.Cells(Rows.Count, 9).End(xlUp).Row

'for loop for conditional formating
For i = 2 To lastrow_Summary

    'if the value in column J is greater than 0 then...
    If ws.Cells(i, 10).Value > 0 Then
        'format the cell green
        ws.Cells(i, 10).Interior.ColorIndex = 4
    Else
        'if not format the cell red
        ws.Cells(i, 10).Interior.ColorIndex = 3
    End If
Next i

'assign headers for greatest table
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"

'for loop to determine greatest increase, decrease, and greatest total volume
For i = 2 To lastrow_Summary

    'if the value in column K is equal to the maximum value in the range of K2:K then...
    If ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & lastrow_Summary)) Then
        'place the value of column I into cell P2
        ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
        'place the value of column K into cell Q2 and format the cell to percentage
        ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
            ws.Cells(2, 17).NumberFormat = "0.00%"
    
    'if the value in column K is equal to minimum value in the range of K2:K then...
    ElseIf ws.Cells(i, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & lastrow_Summary)) Then
        'place the value of column I into cell P3
        ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
        'place the value of column K into cell Q3 and format the cell to percentage
        ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
            ws.Cells(3, 17).NumberFormat = "0.00%"
        
    'if the value in column L is equal to the maximum value in the range of L2:L then...
    ElseIf ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastrow_Summary)) Then
        'place the value of column I into cell P4
        ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
        'place the value of column L into cell Q4
        ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
    End If

Next i

Next ws

End Sub


