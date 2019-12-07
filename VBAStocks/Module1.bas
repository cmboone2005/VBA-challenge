Attribute VB_Name = "Module1"
Sub runtime()
Dim ticker_symbol As String
Dim summary_table_row As Double
Dim lastRow As Double

summary_table_row = 2

'LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
lastRow = Cells(Rows.Count, 1).End(xlUp).Row

Dim yearly_change As Double
'Dim percent_change As Double
Dim starting_price As Double
starting_price = Cells(2, 3).Value


Dim ending_price As Double




Dim total_volume As Double
total_volume = 0

Dim greatest_percent As Double
greatest_percent = 0

Dim lowest_percent As Double
lowest_percent = 0

Dim greatest_volume As Double
greatest_volume = 0

Dim greatest_ticker As String
Dim lowest_ticker As String
Dim volume_ticker As String


For I = 2 To lastRow




    If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
        Dim percent_change As Double
        
        ticker_symbol = Cells(I, 1).Value
        total_volume = total_volume + Cells(I, 7).Value
        ending_price = Cells(I, 6).Value
        yearly_change = ending_price - starting_price
        If starting_price = 0 Then
        precent_change = 0
        Else
        percent_change = yearly_change / starting_price
        End If
        
        
        Range("J" & summary_table_row).Value = yearly_change
            If Range("J" & summary_table_row).Value > 0 Then
                Range("j" & summary_table_row).Interior.ColorIndex = 4
                Range("j" & summary_table_row).Font.ColorIndex = 1
                ElseIf Range("j" & summary_table_row).Value < 0 Then
                Range("j" & summary_table_row).Interior.ColorIndex = 3
                Range("j" & summary_table_row).Font.ColorIndex = 1
                End If
                
        Range("k" & summary_table_row).Value = FormatPercent(percent_change)
        Range("I" & summary_table_row).Value = ticker_symbol
        Range("L" & summary_table_row).Value = total_volume
        
        If percent_change > greatest_percent Then
        greatest_percent = percent_change
        greatest_ticker = ticker_symbol
        End If
        
        If percent_change < lowest_percent Then
        lowest_percent = percent_change
        lowest_ticker = ticker_symbol
        End If
        
        If total_volume > greatest_volume Then
        greatest_volume = total_volume
        volume_ticker = ticker_symbol
        End If

        summary_table_row = summary_table_row + 1
        starting_price = Cells(I + 1, 3).Value
        

        
        total_volume = 0

        Else
        total_volume = total_volume + Cells(I, 7).Value

        

End If
        
Next I


        
        
Range("p" & 2).Value = FormatPercent(greatest_percent)
Range("p" & 3).Value = FormatPercent(lowest_percent)
Range("p" & 4).Value = greatest_volume
Range("o" & 2).Value = greatest_ticker
Range("o" & 3).Value = lowest_ticker
Range("o" & 4).Value = volume_ticker

End Sub
