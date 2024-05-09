Attribute VB_Name = "Module1"
Sub StockData()

Dim ws As Worksheet
For Each ws In Worksheets

ws.Cells(1, "J").Value = "Ticker"
ws.Cells(1, "K").Value = "Quarterly Change"
ws.Cells(1, "L").Value = "Percent Change"
ws.Cells(1, "M").Value = "Total Stock Volume"

ws.Cells(1, "Q").Value = "Ticker"
ws.Cells(1, "R").Value = "Value"
ws.Cells(2, "P").Value = "Greatest % increase"
ws.Cells(3, "P").Value = "Greatest % decrease"
ws.Cells(4, "P").Value = "Greatest Total Volume"

ws.Columns("J:R").AutoFit

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
Dim TickerRow As Integer
TickerRow = 2

Dim Ticker As String
Dim Total_Stock_Volume As Double
Total_Stock_Volume = 0
Dim QuarterlyChange As Double
QuarterlyChange = 0
Dim PercentChange As Double
PercentChange = 0


Dim YearOpen As Double
Dim YearClose As Double

Dim Start As Double
Start = 2

    For i = 2 To LastRow
    
        
'Calculating  the Ticker, Quarterly Change, Percent Change, & Total Stock Volume


        If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
    
        Ticker = ws.Cells(i, 1).Value
        ws.Cells(TickerRow, "J").Value = Ticker
        
        YearClose = ws.Cells(i, 6).Value
        
        YearOpen = ws.Cells(Start, 3)
        
        Start = i + 1
        
        Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, "G").Value
        ws.Cells(TickerRow, "M").Value = Total_Stock_Volume
        
        
        TickerRow = TickerRow + 1
        
        Total_Stock_Volume = 0
        
        
        
       
        
        QuarterlyChange = YearClose - YearOpen
        ws.Cells(TickerRow - 1, "K").Value = QuarterlyChange
        ws.Cells(TickerRow - 1, "K").NumberFormat = "0.00"
        
        PercentChange = QuarterlyChange / YearOpen
        ws.Cells(TickerRow - 1, "L").Value = PercentChange
        ws.Cells(TickerRow - 1, "L").NumberFormat = "0.00%"
        
        
        Else
        Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, "G").Value
        ws.Cells(TickerRow, "M").Value = Total_Stock_Volume

        End If
        

        
    Next i
Next ws

End Sub

Sub ColorIndex()

Dim ws As Worksheet
For Each ws In Worksheets

LastRow = ws.Cells(Rows.Count, "J").End(xlUp).Row

'Applying conditional formatting

    For i = 2 To LastRow
 
        If ws.Cells(i, "K").Value > 0 Then

        ws.Cells(i, "K").Interior.ColorIndex = 4

        ElseIf ws.Cells(i, "K").Value < 0 Then

        ws.Cells(i, "K").Interior.ColorIndex = 3

        End If
        
        If ws.Cells(i, "L").Value > 0 Then

        ws.Cells(i, "L").Interior.ColorIndex = 4

        ElseIf ws.Cells(i, "L").Value < 0 Then

        ws.Cells(i, "L").Interior.ColorIndex = 3

        End If
        
    Next i
Next ws
End Sub


Sub SummaryTable()

Dim ws As Worksheet
For Each ws In Worksheets

LastRow = ws.Cells(Rows.Count, "J").End(xlUp).Row
'Calculating Greatest % increase, Greatest % decrease, and Greatest total volume
    
    For i = 2 To LastRow
    
    'Greatest % increase
    
        ws.Range("R2").Value = WorksheetFunction.Max(ws.Range("L:L"))
        ws.Range("R2").NumberFormat = "0.00%"
        
        If ws.Cells(i, "L").Value = ws.Range("R2").Value Then
        
        Ticker = ws.Cells(i, "J").Value
        ws.Cells(2, "Q").Value = Ticker
        
                
        End If
 
    'Greatest % decrease
    
        ws.Range("R3").Value = WorksheetFunction.Min(ws.Range("L:L"))
        ws.Range("R3").NumberFormat = "0.00%"
        
        If ws.Cells(i, "L").Value = ws.Range("R3").Value Then
        
        Ticker = ws.Cells(i, "J").Value
        ws.Cells(3, "Q").Value = Ticker
        
               
        End If
        
    'Greatest total volume
    
        ws.Range("R4").Value = WorksheetFunction.Max(ws.Range("M:M"))
                
        If ws.Cells(i, "M").Value = ws.Range("R4").Value Then
        
        Ticker = ws.Cells(i, "J").Value
        ws.Cells(4, "Q").Value = Ticker
        
              
        End If
        
        
    Next i
    
Next ws
End Sub

Sub ResetButton()
Dim ws As Worksheet
For Each ws In Worksheets

    ws.Columns("J:R").ClearContents
    
    
    ws.Columns("K:L").Interior.ColorIndex = 0
    

Next ws
End Sub

Sub Calculate()

Call StockData

Call ColorIndex

Call SummaryTable

MsgBox "Voila!!"
End Sub
