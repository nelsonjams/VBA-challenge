Attribute VB_Name = "Module1"
Sub stock_data()

'Start loop through each worksheet
Dim ws As Worksheet
For Each ws In Sheets
    Dim LastRow As Long
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Open Price"
    ws.Cells(1, 11).Value = "Closing Price"
    ws.Cells(1, 12).Value = "Yr. Px Chg"
    ws.Cells(1, 13).Value = "% Px Chg"
    ws.Cells(1, 14).Value = "Total Volume"
    ws.Cells(2, 10).Value = ws.Cells(2, 3).Value
    ws.Range("L2:L" & LastRow).Formula = "=K2-J2"
    ws.Range("M2:M" & LastRow).Formula = "=((K2/J2)-1)"
    ws.Range("M2:M" & LastRow).NumberFormat = "0.00%"

'Set a variable for the ticker
    Dim ticker As String

'Set a variable for the opening price
     Dim open_px As Double

'Set a variable for the closing price
    Dim close_px As Double
         
'Set a variable for the total volume
    Dim volume_total As Double
    'volume_total = 0

'Keep track of each ticker in the summary table
    Dim ticker_table_row As Integer
    ticker_table_row = 2
    
'Loop through the entire ticker table
For i = 2 To LastRow:
'Check that we are still in the same ticker and if it is not
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

        'Set the ticker
        ticker = ws.Cells(i, 1).Value
        'Print ticker in the ticker table
        ws.Range("I" & ticker_table_row).Value = ticker
        
        'Set the opening price
        open_px = ws.Cells(i + 1, 3).Value
        'Print opening price in the ticker table
        ws.Range("J" & ticker_table_row + 1).Value = open_px
        
        'Set the closing price
        close_px = ws.Cells(i, 6).Value
        'Print closing price in the ticker table
        ws.Range("K" & ticker_table_row).Value = close_px
        
        
        'Add volume to the volume_total
        volume_total = volume_total + ws.Cells(i, 7).Value
        
        'Print the total volume of ticker in the ticker table
        ws.Range("N" & ticker_table_row).Value = volume_total
            
        'Add 1 to the ticker table row
        ticker_table_row = ticker_table_row + 1
    
        'reset the volume total
        volume_total = CDbl(0)
        
                          
    Else
    ' Add to the volume total
        volume_total = CDbl(volume_total + ws.Cells(i, 7).Value)
    End If

Next i
'Perform conditional formatting for percent column
Dim rng As Range
Set rng = ws.Range("M2:M" & LastRow)
    For Each rng In rng:
        If WorksheetFunction.IsNumber(rng) Then
            If rng.Value >= 0 Then
                rng.Cells.Interior.ColorIndex = 4
            ElseIf rng.Value < 0 Then
                rng.Cells.Interior.ColorIndex = 3
            End If
            End If
    Next rng

'Put the 1000 separator in the volume column
ws.Range("N2:N" & LastRow).NumberFormat = "#,##0"

'Conditional formatting for positive and negative percent

    
Next ws

End Sub

