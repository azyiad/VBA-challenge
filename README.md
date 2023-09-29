# VBA-challenge

Had issues with Next for compile error.  Please let me kmow what I missed in my code.

Sub StockAnalysis()


'Create Variables
Dim Ticker As String
Dim YearEnd As Integer
Dim Percent As Double
Dim DailyVolume As Long
Dim Opmkt As Integer
Dim Clmkt As Integer
Dim Lastrow As Long
Dim num As Integer
Dim Volume As Long
Dim ws As Worksheet
Dim MaxIncrease As Double
Dim MaxDecrease As Double
Dim MaxVolume As Long
Dim TickerIncrease As String
Dim TickerDecrease As String
Dim TickerVolume As String


'Initialize variables for maximum values
MaxIncrease = 0
MaxDecrease = 0
MaxVolume = 0

'Loop through each worksheet for Each ws in ThisWorkbook.Worksheets
'Check if worksheets name is "2018, "2019", or "2020"

If ws.Name = "2018" Or ws.Name = "2019" Or ws.Name = "2020" Then

'Activate the worksheet ws.Activate

'Create lastrow funtion
Lastrow = Cells(Rows.Count, 1).End(xlUp).Row

num = 2
YearEnd = 0
Percent = 0
Volume = 0


'Create columns on spreadsheet
Range("i1:l1").Value = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")

'choosing start of row and end of row
For i = 2 To Lastrow Step 1

'choosing start of columns to move across columns in loop
    For j = 2 To 7
        Ticker = Cells(i, 1).Value
        Opmkt = Cells(i, 3).Value
        Clmkt = Cells(i, 6).Value
        DailyVolume = Cells(i, 7).Value
        YearEnd = YearEnd + Opmkt - Clmkt
        Percent = Percent + Clmkt / Opmkt
        Volume = Volume + DailyVolume
        
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        Cells(num, 9).Value = Ticker
        Cells(num, 10).Value = YearEnd
        Cells(num, 11).Value = Percent
        Cells(num, 12).Value = Volume
        num = num + 1
        YearEnd = 0
        Percent = 0
        Volume = 0
  End If
  
  Next j
  
  Next i

    
'check if current worksheet has maximum values
If Cells(num, 11).Value > MaxIncrease Then
    MaxIncrease = Cells(num, 11).Value
    TickerIncrease = Cells(num, 9).Value
    
    End If
    
If Cells(num, 11).Value < MaxDecrease Then
    MaxDecrease = Cells(num, 11).Value
    TickerIncrease = Cells(num, 9).Value
    
    End If

If Cells(um, 12).Value > MaxVolume Then
    MaxVolume = Cells(num, 12).Value
    TickerVolume = Cells(num, 9).Value
    End If
    
'Apply conditional formatting to highliight positive yearly cfhange in green and negative yearly change in red
Dim rng As Range
Set rng = Range("J2:J" & num - 1)

'Clear existing format conditions

rng.FormatConditions.Delete

'Format postive values in green rng.formatconditions.add

rng.FormatConditions(rng.FormatConditions.Count).SetFirstPriority
rng.FormatConditions(1).Interior.Color = RGB(0, 255, 0)

'Format negative values in red rng.formatConditions
rng.FormatConditions(rng.FormatConditions.Count).SetFirstPriority
rng.FormatConditions(1).Interior.Color = RGB(255, 0, 0)

'Clear any other existing format conditions
rng.FormatConditions(1).StopIfTrue = False
 

    End If
    
  
    
    Next ws


'Print the stocks with the greatest % increase, greatest % derease, and greates total volume

MsgBox " Stocks with the greatest % increase: " & TickerIncrease & "(" & Format(MaxIncrease, "0.00%") & ")" & vbCrLf & "Stock with the greatest total volume: " & TickerVolume & "(" & MaxVolumr & ")"

End Sub
