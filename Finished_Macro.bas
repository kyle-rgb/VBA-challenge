Attribute VB_Name = "Module1"
Sub Yearly_Res()
Attribute Yearly_Res.VB_ProcData.VB_Invoke_Func = "X\n14"

Dim iter As Double
iter = 2
Dim open_price As Double
open_price = Cells(2, 3).Value
Dim close_price As Double
Dim percentage As Currency
percentage = 0
Dim total_volume As Double
total_volume = 0

TotalRows = Range("A2", Range("A2").End(xlDown).Offset(1, 0)).Rows.Count
Range("A2").Select

Range("I1:L1").Value = Array("Ticker", "Yearly Change", "Percentage Change", "Total Stock Volume")
Range("I1:L1").WrapText = True
Range("I1:L1").HorizontalAlignment = xlCenter
Range("I1:L1").Font.Bold = True

For x = 2 To TotalRows

    If Cells(x, 1).Value <> Cells(x + 1, 1).Value Then
        total_volume = Cells(x, 7).Value + total_volume
        Cells(iter, 9).Value = Cells(x, 1).Value
        close_price = Cells(x, 6).Value
        Cells(iter, 10).Value = close_price - open_price
        If Cells(iter, 10).Value > 0 Then
            Cells(iter, 10).Interior.ColorIndex = 4
        Else:
            Cells(iter, 10).Interior.ColorIndex = 3
            End If
        If Cells(iter, 10).Value <> 0 And open_price <> 0 Then
            Cells(iter, 11).Value = FormatPercent(Cells(iter, 10).Value / open_price)
        ElseIf open_price = 0 And Cells(iter, 10).Value <> 0 Then
            Cells(iter, 11).Value = FormatPercent(1)
        Else:
            Cells(iter, 11).Value = 0
            End If
        Cells(iter, 12).Value = total_volume
        open_price = Cells(x + 1, 3).Value
        total_volume = Cells(x + 1, 7).Value
        x = x + 1
        iter = iter + 1
    Else:
        total_volume = total_volume + Cells(x, 7).Value
        End If
    Next

SumRows = Range("I2", Range("I2").End(xlDown)).Rows.Count

Dim maxreturn As Double
maxreturn = Cells(2, 11).Value
Dim max_name As String
max_name = Cells(2, 9).Value
Dim minreturn As Double
minreturn = Cells(2, 11).Value
Dim min_name As String
min_name = Cells(2, 9).Value
Dim large_vol As LongLong
large_vol = Cells(2, 12).Value
Dim vol_name As String
vol_name = Cells(2, 9).Value

For t = 2 To SumRows:
    If maxreturn < Cells(t + 1, 11).Value Then
        max_name = Cells(t + 1, 9).Value
        maxreturn = Cells(t + 1, 11).Value
    ElseIf minreturn > Cells(t + 1, 11).Value Then
        min_name = Cells(t + 1, 9).Value
        minreturn = Cells(t + 1, 11).Value
    ElseIf large_vol < Cells(t + 1, 12).Value Then
        vol_name = Cells(t + 1, 9).Value
        large_vol = Cells(t + 1, 12).Value
    End If
    Next
    
Range("O1:Q1").Value = Array("", "Ticker", "Value")
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"
Cells(2, 17).Value = FormatPercent(maxreturn)
Cells(2, 16).Value = max_name
Cells(3, 17).Value = FormatPercent(minreturn)
Cells(3, 16).Value = min_name
Cells(4, 17).Value = large_vol
Cells(4, 16).Value = vol_name

End Sub
