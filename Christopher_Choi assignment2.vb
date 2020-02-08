Sub assignment2()

'loop through all sheets
For Each ws In Worksheets
  
'input header for the assignment required items
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"



'input values from data columns to assignment columns
Dim ticker As String

Dim yearlyChange As Double
yearlyChange = 0

Dim percentChange As Double
percentChange = 0

Dim totalStockVolume As Double
totalStockVolume = 0

Dim summaryTableRow As Integer
summaryTableRow = 2
'this is just for equation clarity and will not be displayed on summary table
Dim openPrice As Double

Dim closePrice As Double


'find out where's the last row
Dim lastrow As Long
lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row


For i = 2 To lastrow
    'if a row's value in column A is different to the previous one
    If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
    
        'picking the first <open> value if vba realises the ticker value of the current row it's checking is different than the previous row
        openPrice = ws.Cells(i, 3).Value
        
    End If
    
    'if a row's value in column A is different to the next one
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
    ticker = ws.Cells(i, 1).Value
    
    closePrice = ws.Cells(i, 6).Value
        
        
    yearlyChange = closePrice - openPrice

        If openPrice = 0 Then
        percentChange = 0
    
        Else
        percentChange = (yearlyChange / openPrice)
    
        End If
    
    totalStockVolume = totalStockVolume + ws.Cells(i, 7).Value
    
    'printing values into proper cells
    ws.Range("I" & summaryTableRow).Value = ticker
    ws.Range("J" & summaryTableRow).Value = yearlyChange
    ws.Range("K" & summaryTableRow).Value = percentChange
    ws.Range("L" & summaryTableRow).Value = totalStockVolume
    
    'move on to next row
    summaryTableRow = summaryTableRow + 1
    
    'resetting values
    yearlyChange = 0
    openPrice = 0
    closePrice = 0
    percentChange = 0
    totalStockVolume = 0
    
    Else
    
    totalStockVolume = totalStockVolume + ws.Cells(i, 7).Value
    
    End If
    
Next i
    
'input conditional colouring for 'Yearly Change'
Dim lastrow2 As Long
lastrow2 = ws.Cells(Rows.Count, "J").End(xlUp).Row

        For x = 2 To lastrow2
            'positive = green
            If ws.Cells(x, 10).Value > 0 Then
                ws.Cells(x, 10).Interior.ColorIndex = 4
            '0 = white
            ElseIf ws.Cells(x, 10).Value = 0 Then
                ws.Cells(x, 10).Interior.ColorIndex = 2
            'negative = red
            Else
                ws.Cells(x, 10).Interior.ColorIndex = 3
            End If
        Next x

'changing column k's number format to %
Dim lastrow3 As Long
lastrow3 = ws.Cells(Rows.Count, "K").End(xlUp).Row
ws.Range("K2:K" & lastrow3).Style = "Percent"
ws.Range("K2:K" & lastrow3).NumberFormat = "0.00%"

Next ws

End Sub
