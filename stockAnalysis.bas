Sub stockAnalysis():
Dim ws As Worksheet
Dim i,count as Integer
'loop through each work
For Each ws In Worksheets
    'set column titles of new table starting in K1
    ws.Range("K1").Value = "Ticker"
    ws.Range("L1").Value = "Total Change In Year"
    ws.Range("M1").Value = "Percentage change in Year"
    ws.Range("N1").Value = "Volume traded in Year"
    ws.Range("R1").Value = "Ticker"
    ws.Range("S1").Value = "Value"
    ws.Range("Q2").Value = "Greatest % Increase"
    ws.Range("Q3").Value = "Greated % Decrease"
    ws.Range("Q4").Value = "Greatest Total Volume"
    'Start at row 2 for rows to the new table
    count = 2
    Dim volume,maxVolume As LongLong
    volume = 0&
    maxVolume = 0&
    Dim opening, closing, maxIncrease, maxDecrease As Double
    maxIncrease = 0
    maxDecrease = 100000 
    Dim currentTicker, newTicker, maxVolumeTicker, maxIncreaseTicker,maxDecreaseTicker As String
    'loop from row 2 until the first empty row
    i = 2
    Do Until isEmpty(ws.Cells(i,1))
        'if this is a new stock
        If ws.Cells(i,1).Value <> currentTicker Then
            'if it is the first overall ticker on the worksheet, no data to output
            If i = 2 Then
                'no data to output here
            Else
                'output the data and then reset volume to 0, and increment count so the next row adds below
                ws.Cells(count, 11).Value = currentTicker
                ws.Cells(count, 12).Value = closing - opening
                If opening = 0 Then
                    ws.Cells(count, 13).Value = "n/a"
                    ws.Cells(count, 13).Interior.ColorIndex = 6
                Else
                    ws.Cells(count, 13).Value = (closing/opening)
                    ws.Cells(count, 13).NumberFormat = "0.00%"
                    If (closing/opening) > maxIncrease Then
                        maxIncrease = closing/opening
                        maxIncreaseTicker = currentTicker
                    End If
                    If  closing/opening < maxDecrease Then
                        maxDecrease = closing/opening
                        maxDecreaseTicker = currentTicker
                    End If
                End If
                ws.Cells(count, 14).Value = volume
                count = count + 1
                If volume > maxVolume Then
                    maxVolume = volume
                    maxVolumeTicker = currentTicker
                End If
                volume = 0&
            End If
            'set the new current ticker, opening and add the daily volume and closing
            currentTicker = ws.Cells(i,1).Value
            opening = CDbl(ws.Cells(i,3).Value)
            closing = CDbl(ws.Cells(i,6).Value)
            volume = CLngLng(ws.Cells(i,7).Value)
        Else 
            'not a new stock, so we don't set the current ticker or opening, just the closing and add the volume
            closing = CDbl(ws.Cells(i,6).Value)
            volume = volume + CLngLng(ws.Cells(i,7).Value)
        End If
        i = i +1
        'if it is the last stock (i is already incremented, so this is checing the next one), then output the data
        If isEmpty(ws.Cells(i,1)) Then
            ws.Cells(count, 11).Value = currentTicker
            ws.Cells(count, 12).Value = closing - opening
            If opening = 0 Then
                ws.Cells(count, 13).Value = "n/a"
            Else
                ws.Cells(count, 13).Value = (closing/opening)
                ws.Cells(count, 13).NumberFormat = "0.00%"
                If closing/opening > maxIncrease Then
                        maxIncrease = closing/opening
                        maxIncreaseTicker = currentTicker
                End If
                If closing/opening < maxDecrease Then
                        maxDecrease = closing/opening
                        maxDecreaseTicker = currentTicker
                End If
            End If
            ws.Cells(count, 14).Value = volume
            count = count + 1
            If volume > maxVolume Then
                    maxVolume = volume
                    maxVolumeTicker = currentTicker
            End If
            volume = 0&
        End If
    Loop
    ws.Range("R2").Value = maxIncreaseTicker
    ws.Range("R3").Value = maxDecreaseTicker
    ws.Range("R4").Value = maxVolumeTicker
    ws.Range("S2").Value = maxIncrease
    ws.Range("S2").NumberFormat = "0.00%"
    ws.Range("S2").Interior.Colorindex = 4
    ws.Range("S3").Value = maxDecrease
    ws.Range("S3").NumberFormat = "0.00%"
    ws.Range("S3").Interior.Colorindex = 3
    ws.Range("S4").Value = maxVolume
    'set the conditional formatting
    Dim range As Range
    Dim condition1,condtion2,condition3 As FormatCondition
    Set range = ws.Range("L:L")
    range.FormatConditions.Delete
    
    Set condition1 = range.FormatConditions.Add(xlCellValue,xlLess,"=0")
    Set condition2 = range.FormatConditions.Add(xlCellValue,xlGreater,"=0")
    With condition1
        .Font.Color = vbRed
        .Font.Bold = True
    End With
    
    With condition2
        .Font.Color = vbGreen
        .Font.Bold = True
    End With

    Set range = ws.Range("M:M")
    range.FormatConditions.Delete
    Set condition1 = range.FormatConditions.Add(xlCellValue,xlLess,"1")
    Set condition2 = range.FormatConditions.Add(xlCellValue,xlGreater,"1")
    Set condition3 = range.FormatConditions.Add(xlCellValue,xlEqual,"n/a")

    With condition1
        .Font.Color = vbRed
        .Font.Bold = True
    End With
    
    With condition2
        .Font.Color = vbGreen
        .Font.Bold = True
    End With

    With condition3
        .Font.Color = vbYellow
        .Font.Bold = True
    End With

    Set range = ws.Range("L1:M1")
    range.FormatConditions.Delete
    
Next


End Sub
