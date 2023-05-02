Sub ticker_symbol()

Dim i As Long
Dim total_volume As Double
Dim yearly_change As Double
Dim percent_change As Double

'set new variables to calculate the yearly change and percentage change

Dim initial_open As Double
Dim initial_row As Long
Dim final_close As Double
Dim final_row As Long

Dim ticker As String
Dim summary_row As Integer

Dim LastRow As Long
Dim LastSummary As Long

total_volume = 0
summary_row = 2
yearly_change = 0
percent_change = 0

initial_row = 2

'Loop through all worksheet and identify the last row

    For Each ws In Worksheets
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Volume"
    
        ' iterate through each row starting on the second row
        
        For i = 2 To LastRow
        
            ' start on the initial row and check if it is diff than the next row
            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                
                'print ticker to summary table
                ticker = ws.Cells(i, 1).Value
                
                final_row = i
                
                ' add total_volume together
                total_volume = total_volume + ws.Cells(i, 7).Value
                
                'capture initial open
                initial_open = ws.Cells(initial_row, 3).Value
                
                'capture final close to calculate the yearly change value
                final_close = ws.Cells(final_row, 6).Value
                
                'Calculate yearly changeby
                yearly_change = final_close - initial_open
                
                'Calculate percent change
                If yearly_change = 0 Then
                    percent_change = 0
                Else
                    percent_change = yearly_change / initial_open
                End If
                
                 'Set yearly change to red or green
                If yearly_change < 0 Then
                    ws.Cells(summary_row, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(summary_row, 10).Interior.ColorIndex = 4
                End If
                'Set percent change tored or green
                If percent_change < 0 Then
                    ws.Cells(summary_row, 11).Interior.ColorIndex = 3
                Else
                    ws.Cells(summary_row, 11).Interior.ColorIndex = 4
                End If
                
                'print to ticker to summary section
                ws.Cells(summary_row, 9).Value = ticker
                'print the yearly change, percentage change, and total volume
                ws.Cells(summary_row, 10).Value = yearly_change
                ws.Cells(summary_row, 11).Value = Str(percent_change * 100) + "%"
                ws.Cells(summary_row, 12).Value = total_volume
                
        
            'reset the values to zero to allow the ne ticker values to be aggregated
            total_volume = 0
            yearly_change = 0
            percent_change = 0
            
            'increment these variable to allow the new aggregate numbers to be printed in the summary table
            summary_row = summary_row + 1
            
            'establishes the new ticker's open price to calculate the yearly change value
            initial_row = i + 1
        Else
            
            total_volume = total_volume + Cells(i, 7).Value
           
        
           
        End If
    Next i
    
    'Determine Greatest Summary Stats
    
    Dim Max_Inc As Double
    Dim Max_Dec As Double
    Dim Max_Vol As Double
    
    Dim Max_Inc_Row As Integer
    Dim Max_Dec_Row As Integer
    Dim Max_Vol_Row As Integer
    
    Max_Inc_Row = 2
    Max_Dec_Row = 2
    Max_Vol_Row = 2
    
    'setting the value to compare the max % decrease and increase
    Max_Inc = 0
    Max_Dec = 0
    Max_Volume = 0
    
    'finding the last row of the summary table for the For loop to iterate through
    LastSummary = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    'printing row and column headers for the greatest aggregate stats
    ws.Cells(2, 14).Value = "Greatest % Increase"
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    ws.Cells(4, 14).Value = "Greatest Volume"
    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value"

    For i = 2 To LastSummary
    
    'finding max % inc with If statement
        If ws.Cells(i, 11).Value > Max_Inc Then
            Max_Inc = ws.Cells(i, 11).Value
            Max_Inc_Row = i
        'finding max % decrease with if statement
        ElseIf ws.Cells(i, 11).Value < Max_Dec Then
            Max_Dec = ws.Cells(i, 11).Value
            Max_Dec_Row = i
        Else
            Max_Inc = Max_Inc
            Max_Dec = Max_Dec
        End If
        
        'find max volume
         If ws.Cells(i, 12).Value > Max_Vol Then
            Max_Vol = ws.Cells(i, 12).Value
            Max_Vol_Row = i
        End If
        
        'print max inc ticker
        ws.Cells(2, 15).Value = ws.Cells(Max_Inc_Row, 9).Value
        'print max % inc value
        ws.Cells(2, 16).Value = Str(Max_Inc * 100) + "%"
        
        'print max dec ticker
        ws.Cells(3, 15).Value = ws.Cells(Max_Dec_Row, 9).Value
        'print max % dec value
        ws.Cells(3, 16).Value = Str(Max_Dec * 100) + "%"
        
        'print max volume ticker and value
        ws.Cells(4, 15).Value = ws.Cells(Max_Vol_Row, 9).Value
        ws.Cells(4, 16).Value = Max_Vol
    Next i
            
'reset values for new sheet
initial_row = 2
summary_row = 2
Next ws
End Sub

