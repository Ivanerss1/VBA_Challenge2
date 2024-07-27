Attribute VB_Name = "Module2"
Sub Greatest_Value()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim percentChange As Double
    Dim maxIncrease As Double
    Dim maxDecrease As Double
    Dim maxVolume As Double
    Dim maxIncreaseTicker As String
    Dim maxDecreaseTicker As String
    Dim maxVolumeTicker As String
    Dim tickerVolume As Collection
    Dim tickerVolumeArrays As Variant
    Dim ticker As String
    Dim vol As Double
    Dim found As Boolean
    Dim j As Long
    
    
For Each ws In ThisWorkbook.Worksheets
lastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
    Set tickerVolume = New Collection
    
    'variables
maxIncrease = -1000000
maxDecrease = 1000000
maxVolume = 0

'loop through each row
For i = 2 To lastRow
    percentChange = ws.Cells(i, 11).Value
    
    ' maximum increae
    If percentChange > maxIncrease Then
    maxIncrease = percentChange
    maxIncreaseTicker = ws.Cells(i, 9).Value
End If
    
        'maximum Decrease
    If percentChange < maxDecrease Then
    maxDecrease = percentChange
    maxDecreaseTicker = ws.Cells(i, 9).Value
End If

        'total volume
    ticker = ws.Cells(i, 9).Value
    vol = ws.Cells(i, 13).Value
    found = False
    
    For j = 1 To tickerVolume.Count
    Set tickerVolumeArrays = tickerVolume(j)
    If tickerVolumeArrays(1) = ticker Then
    tickerVolumeArrays(2) = tickerVolumeArrays(2) + vol
    tickerVolume.Remove j
    tickerVolume.Add tickerVolumeArrays
    found = True
    End If
Next j

If Not found Then Set tickerVolumeArrays = New Collection
tickerVolumeArrays.Add ticker
tickerVolumeArrays.Add vol
tickerVolume.Add tickerVolumeArrays

    Next i

        'Find ticker with max vol.
    For j = 1 To tickerVolume.Count
    Set tickerVolumeArrays = tickerVolume(j)
    If tickerVolumeArrays(2) > maxVolume Then
    maxVolume = tickerVolumeArrays(2)
    maxVolumeTicker = tickerVolumeArrays(1)
End If
        Next j
        
        'results output
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatesr % Increase"
    ws.Cells(2, 16).Value = maxIncreaseTicker
    ws.Cells(2, 17).Value = Format(maxIncrease, "0.00") & "%"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(3, 16).Value = maxDecreaseTicker
    ws.Cells(3, 17).Value = Format(maxDecrease, "0.00") & "%"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(4, 16).Value = maxVolumeTicker
    ws.Cells(4, 17).Value = Format(maxVolume, "0.00") & "%"
    
     Next ws
    
        
    


End Sub
