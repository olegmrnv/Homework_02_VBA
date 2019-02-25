Attribute VB_Name = "Module1"
Sub total_counter()

'declaring variables we will be using
Dim row_count, i, Total, open_value, close_value, yearly_change As Double
Dim k As Integer

'creating loop for each worksheet in this file
For Each ws In Worksheets
    
    'will be using "k" variable for the row where we print next stock ticker and total value
    k = 2
    Total = 0
    
    'printing table headers and applying format
    ws.Cells(1, 9) = "Ticker"
    ws.Cells(1, 12) = "Total Stock Volume"
    ws.Cells(1, 10) = "Yearly Change"
    ws.Cells(1, 11) = "Percent Change"
    ws.Cells(2, 15) = "Greatest % increase"
    ws.Cells(3, 15) = "Greatest % Decrease"
    ws.Cells(4, 15) = "Greatest total volume"
    ws.Cells(1, 16) = "Ticker"
    ws.Cells(1, 17) = "Value"
    ws.Cells(2, 17).NumberFormat = "0.00%"
    ws.Cells(3, 17).NumberFormat = "0.00%"

    'detecting amount of rows
    row_count = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'creating a loop to go through each row
    For i = 2 To row_count
        
        'detecting the open value of current stock
        If ws.Cells(i, 1) <> ws.Cells(i - 1, 1) Then open_value = ws.Cells(i, 3)
        
        'summarizing total volume of current stock
        If ws.Cells(i, 1) = ws.Cells(i + 1, 1) Then
            Total = Total + ws.Cells(i, 7)
        
        Else
            'if ticker changes we reading close value, calculating yearly change and yearly change percent
            close_value = ws.Cells(i, 6)
            yearly_change = close_value - open_value
            If open_value <> 0 Then yearly_change_percent = yearly_change / open_value
            
            'adding last volume to total, printing info to summary table, applying conditional formatting
            Total = Total + ws.Cells(i, 7)
            ws.Cells(k, 9) = ws.Cells(i, 1)
            ws.Cells(k, 12) = Total
            ws.Cells(k, 10) = yearly_change
            If yearly_change < 0 Then ws.Cells(k, 10).Interior.Color = RGB(255, 0, 0)
            If yearly_change > 0 Then ws.Cells(k, 10).Interior.Color = RGB(0, 255, 0)
            ws.Cells(k, 11) = yearly_change_percent
            ws.Cells(k, 11).NumberFormat = "0.00%"
            
            'checking if current yearly change percent is highest and then printing to apropriate cell
            If yearly_change_percent > ws.Cells(2, 17) Then
                ws.Cells(2, 17) = yearly_change_percent
                ws.Cells(2, 16) = ws.Cells(i, 1)
            End If
            'checking if current yearly change percent is lowest and then printing to apropriate cell
            If yearly_change_percent < ws.Cells(3, 17) Then
                ws.Cells(3, 17) = yearly_change_percent
                ws.Cells(3, 16) = ws.Cells(i, 1)
            End If
            'checking if current total volume is highest and printing it to appropriate cell
            If Total > ws.Cells(4, 17) Then
                ws.Cells(4, 17) = Total
                ws.Cells(4, 16) = ws.Cells(i, 1)
            End If
            
            'resetting total and going to next cell in summary table
            Total = 0
            k = k + 1
        End If
        
    Next i
    'applying formatting to entire document
    ws.Columns("A:Q").AutoFit

Next ws
End Sub


