
Sub stockmarket()

Dim WS As Worksheet

For Each WS In Worksheets

WS.Range("I1").Value = "Ticker"
WS.Range("J1").Value = "Yearly Change"
WS.Range("k1").Value = "Percent Change"
WS.Range("l1").Value = "Total Stock Volume"


WS.Range("N2").Value = "Greatest % Increase"
WS.Range("N3").Value = "Greatest % Decrease"
WS.Range("N4").Value = "Greatest Total Volume"
WS.Range("O1").Value = "Ticker"
WS.Range("P1").Value = "Value"



Dim Max_Percentage As Double
Dim Min_Percentage As Double
Dim Max_Volume As Double
Dim Ticker As String
Dim Max_Percentage_Ticker As String
Dim Min_Percentage_Ticker As String
Dim Max_Volume_Ticker As String


Dim Volume As Double

Dim Change As Double

Dim Percentage As Double

Dim summary_row_counter As Double

Dim opening_price_counter As Double

summary_row_counter = 2
opening_price_counter = 2

Volume = 0



RowCount = WS.Cells(Rows.Count, "A").End(xlUp).Row
RowCount2 = WS.Cells(Rows.Count, "J").End(xlUp).Row


For i = 2 To RowCount
Ticker = WS.Cells(i, 1).Value

If Ticker <> WS.Cells(i + 1, 1).Value Then

  Volume = Volume + WS.Cells(i, 7).Value

Change = WS.Cells(i, 6).Value - WS.Cells(opening_price_counter, 3).Value

Percentage = Change / WS.Cells(opening_price_counter, 3).Value
If Percentage > Max_Percentage Then
Max_Percentage = Percentage
Max_Percentage_Ticker = Ticker

ElseIf Percentage < Min_Percentage Then
Min_Percentage = Percentage
Min_Percentage_Ticker = Ticker

ElseIf Volume > Max_Volume Then
Max_Volume = Volume
Max_Volume_Ticker = Ticker

End If

For j = 2 To RowCount2
If WS.Cells(j, 10).Value >= 0 Then
WS.Cells(j, 10).Interior.ColorIndex = 4

Else
WS.Cells(j, 10).Interior.ColorIndex = 3


End If
Next j

WS.Range("I" & summary_row_counter).Value = Ticker
WS.Range("J" & summary_row_counter).Value = Change
WS.Range("J" & summary_row_counter).NumberFormat = "0.00"
WS.Range("K" & summary_row_counter).Value = Percentage
WS.Range("K" & summary_row_counter).NumberFormat = "0.00%"
WS.Range("L" & summary_row_counter).Value = Volume

Volume = 0
Change = 0

summary_row_counter = summary_row_counter + 1

opening_price_counter = i + 1


Else
    Volume = Volume + WS.Cells(i, 7).Value
    



End If

Next i


WS.Range("O2").Value = Max_Percentage_Ticker
WS.Range("P2").Value = Max_Percentage
WS.Range("P2").NumberFormat = "0.00%"
WS.Range("O3").Value = Min_Percentage_Ticker
WS.Range("P3").Value = Min_Percentage
WS.Range("P3").NumberFormat = "0.00%"
WS.Range("O4").Value = Max_Volume_Ticker
WS.Range("P4").Value = Max_Volume

WS.Columns("I:P").AutoFit

Next WS



End Sub

