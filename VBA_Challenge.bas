Attribute VB_Name = "Module1"
Sub VBAChallenge():

'Worksheet looping

Dim WS As Worksheet

For Each WS In Worksheets

'Defines variables

Dim Ticker As String
Dim YearStart As Double
Dim YearEnd As Double
Dim TotalVolume As Double
Dim TickerCount As LongLong

TickerCount = 1


'Formats Worksheet

WS.Cells(1, 9).Value = "Ticker"
WS.Cells(1, 10).Value = "Yearly Change"
WS.Cells(1, 11).Value = "Percent Change"
WS.Cells(1, 12).Value = "Total Stock Volume"
WS.Cells(2, 14).Value = "Greatest % Increase"
WS.Cells(3, 14).Value = "Greatest % Decrease"
WS.Cells(4, 14).Value = "Greatest Total Volume"
WS.Cells(1, 15).Value = "Ticker"
WS.Cells(1, 16).Value = "Value"


'Find the number of rows in the sheet

NumRows = WS.Range("A1", WS.Range("A1").End(xlDown)).Rows.Count
i = 2
For i = 2 To NumRows
    
    'Compares current ticker to the ticker in the row above
    
    If WS.Cells(i, 1).Value <> WS.Cells(i - 1, 1) Then
    YearStart = WS.Cells(i, 3).Value
    End If
    
    'Compares current ticker to the ticker in the row bellow
    
    If WS.Cells(i, 1).Value = WS.Cells(i + 1, 1).Value Then
        
    'Adds volume to the total volume
    
    TotalVolume = TotalVolume + WS.Cells(i, 7).Value
    
    Else
    'Sets year end as close value of the current row
    
    YearEnd = WS.Cells(i, 6).Value
       
    'Adds volume to the total volume
    
    TotalVolume = TotalVolume + WS.Cells(i, 7).Value
    'Increases each time the ticker changes
    
    TickerCount = TickerCount + 1
    
    'Prints values to the cells
    
    WS.Cells(TickerCount, 9).Value = WS.Cells(i, 1).Value
    WS.Cells(TickerCount, 10).Value = YearEnd - YearStart
    If YearStart = 0 Then   'Prevents div0 error
    WS.Cells(TickerCount, 11).Value = "N/A"
    Else
    WS.Cells(TickerCount, 11).Value = (YearEnd - YearStart) / YearStart
    End If
    WS.Cells(TickerCount, 12).Value = TotalVolume
    
    'Resets TotalVolume
    
    TotalVolume = 0

    End If
    
    Next i

'Formatting
OutputRows = WS.Range("I1", WS.Range("I1").End(xlDown)).Rows.Count

WS.Range("K1", WS.Cells(OutputRows, 11)).NumberFormat = "0.00%"

WS.Range("P2:P3").NumberFormat = "0.00%"

Dim range1 As Range
Dim ConFormat1 As FormatCondition
Dim ConFormat2 As FormatCondition

Set range1 = WS.Range("J2", WS.Range("J2").End(xlDown))

Set ConFormat1 = range1.FormatConditions.Add(xlCellValue, xlGreater, "=0")
Set ConFormat2 = range1.FormatConditions.Add(xlCellValue, xlLess, "=0")

With ConFormat1
    .Interior.Color = vbGreen

End With

With ConFormat2
    .Interior.Color = vbRed

End With

'Finds max
Dim PRng As Range

Dim VRng As Range

Set PRng = WS.Range("K1", WS.Cells(OutputRows, 11))

Set VRng = WS.Range("L1", WS.Cells(OutputRows, 12))

Dim MaxI As Double

Dim MaxD As Double

Dim MaxV As Double

MaxI = Application.WorksheetFunction.Max(PRng)

MaxD = Application.WorksheetFunction.Min(PRng)

MaxV = Application.WorksheetFunction.Max(VRng)

'Finds rows of max
iRow = WorksheetFunction.Match(MaxI, PRng, 0)

dRow = WorksheetFunction.Match(MaxD, PRng, 0)

VRow = WorksheetFunction.Match(MaxV, VRng, 0)
'Writes values

WS.Range("O2").Value = WS.Cells(iRow, 9)

WS.Range("O3").Value = WS.Cells(dRow, 9)

WS.Range("O4").Value = WS.Cells(VRow, 9)

WS.Range("P2").Value = MaxI

WS.Range("P3").Value = MaxD

WS.Range("P4").Value = MaxV


Next WS

End Sub

