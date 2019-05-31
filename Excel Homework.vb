Sub StockCounter()

Dim Ticker As String
Dim Volume As Double
Dim Open_Price As Double
Dim Close_Price As Double
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim ws As Worksheet


Dim Sum_Ticker As String
Dim Sum_Volume As Double
Sum_Volume = 0

Dim Sum_Ticker_Row As Integer
Sum_Ticker_Row = 2

For Each ws In Worksheets
ws.Activate
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
 
Range("I1") = "Ticker"
Range("J1") = "Yearly Change"
Range("K1") = "Percent Change"
Range("L1") = "Total Stock Volume"
Range("O2") = "Greatest % Increase"
Range("O3") = "Greatest % Decrease"
Range("O4") = "Greatest Total Volume"
Range("P1") = "Ticker"
Range("Q1") = "Value"


For i = 2 To lastrow

    If i = 2 Then
        Open_Price = Cells(i, 3).Value
    ElseIf Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
        Open_Price = Cells(i, 3).Value
    End If
    
    
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        Close_Price = Cells(i, 6).Value
    ElseIf i = lastrow Then
        Close_Price = Cells(i, 6).Value
    End If
    
Yearly_Change = Close_Price - Open_Price

    If Open_Price > 0 Then
        Percent_Change = (Close_Price - Open_Price) / Open_Price
    ElseIf Open_Price = 0 Then
        Percent_Change = 0
    End If
    
   

    If Percent_Change > 0 Then
        Range("J" & Sum_Ticker_Row).Interior.ColorIndex = 4
    ElseIf Percent_Change < 0 Then
        Range("J" & Sum_Ticker_Row).Interior.ColorIndex = 3
    End If

        
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        Sum_Ticker = Cells(i, 1).Value
        Sum_Volume = Sum_Volume + Cells(i, 7).Value
        Range("I" & Sum_Ticker_Row).Value = Sum_Ticker
        Range("J" & Sum_Ticker_Row).Value = Yearly_Change
        Range("K" & Sum_Ticker_Row).Value = FormatPercent(Percent_Change, 2)
        Range("L" & Sum_Ticker_Row).Value = Sum_Volume
        Sum_Ticker_Row = Sum_Ticker_Row + 1
        Sum_Volume = 0
    Else
        Sum_Volume = Sum_Volume + Cells(i, 7).Value
    End If

Next i

Dim Max As Double
Dim Min As Double
Dim Vol As Double
Dim Max_Percent As Double
Dim Min_Percent As Double
Dim Max_Vol As Double
Max = 0
Min = 0
Vol = 0

lRow = ws.Cells(Rows.Count, 12).End(xlUp).Row
Dim rng As Range
Set rng = Range("I2:L" & lRow)

For i = 2 To lRow
    
    If Cells(i, 11).Value > Max Then
       Ticker = Cells(i, 9).Value
       Max_Percent = Cells(i, 11).Value
       Range("P2") = Ticker
       Range("Q2") = FormatPercent(Max_Percent, 2)
    ElseIf Cells(i, 11).Value < Min Then
       Ticker = Cells(i, 9).Value
       Min_Percent = Cells(i, 11).Value
       Range("P3") = Ticker
       Range("Q3") = FormatPercent(Min_Percent, 2)
    End If
    
    If Cells(i, 12) > Vol Then
        Ticker = Cells(i, 9).Value
        Max_Vol = Cells(i, 12).Value
        Range("P4") = Ticker
        Range("Q4") = Max_Vol
    End If
    
    Max = Max_Percent
    Min = Min_Percent
    Vol = Max_Vol
    
Next i

Range("I1:L1").Columns.AutoFit
Range("O:Q").Columns.AutoFit
Sum_Ticker_Row = 2

Next ws

End Sub




