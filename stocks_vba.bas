Attribute VB_Name = "Module1"
Sub StockTicket()
    
    Dim tNm As String
    Dim tVol As Double
    Dim lrow As Long
    Dim sRow As Integer
    Dim tOpen As Double
    Dim tClose As Double
    Dim tDelta As Double
    Dim pDelta As Double
    Dim wksht As Worksheet
    Dim s_MaxDelta As Double
    Dim s_MaxNm As String
    Dim s_MinDelta As Double
    Dim s_MinNm As String
    Dim s_MaxVol As Double
    Dim s_MaxVolNm As String
    Dim s_lrow As Long
    
    
    For Each wksht In ThisWorkbook.Worksheets
        wksht.Activate
    
    'Presets
        tVol = 0
        lrow = Cells(Rows.Count, 1).End(xlUp).Row
        sRow = 2
    
    'Headers of Summary
        Cells(1, 9).Value = "TIcker"
        'Cells(1, 10).Value = "Total Volume" --Easy
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
    
    'Loop for values
        For i = 2 To lrow
            
            'Get open value
            If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
                tOpen = Cells(i, 3).Value
            End If
            
            'Remaining Values and Summary
            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                tNm = Cells(i, 1).Value
                tVol = tVol + Cells(i, 7).Value
                tClose = Cells(i, 6).Value
                If tOpen = 0 And tClose = 0 Then
                    tDelta = 0
                    pDelta = 0
                ElseIf tOpen = 0 Then
                    tDelta = tClose - tOpen
                    pDelta = (tDelta / 0.01)
                Else
                tDelta = tClose - tOpen
                pDelta = (tDelta / tOpen)
                End If
                
                'tDelta = (1 + pDelta) * tOpen --to show close amount
                
                '26.6 -- start
                '17.83 --End
                
                
                'fill summary
                Cells(sRow, 9).Value = tNm
                Cells(sRow, 10).Value = tDelta
                Cells(sRow, 11).Value = pDelta
                Cells(sRow, 12).Value = tVol
                
                'Coloring
                If pDelta > 0 Then
                    Cells(sRow, 10).Interior.ColorIndex = 4
                ElseIf pDelta < 0 Then
                    Cells(sRow, 10).Interior.ColorIndex = 3
                End If
                
                Cells(sRow, 11).NumberFormat = "0.00%"
                'next sum row
                sRow = sRow + 1
                
                'Reset volume
                tVol = 0
                'Debug.Print tVol
            Else
                tVol = tVol + Cells(i, 7).Value
            End If
        Next i
        'hard - Filter through values in summary to find max/min
        
        'Summary Last row
        s_lrow = Cells(Rows.Count, 9).End(xlUp).Row
        
        'Loop through summary comparing max/min
        s_MaxDelta = 0
    'Dim s_MaxNm As String
        s_MinDelta = 0
    'Dim s_MinNm As String
        s_MaxVol = 0
    'Dim s_MaxVolNm As String
    
        For s = 2 To s_lrow
            If Cells(s, 11).Value > s_MaxDelta Then
                s_MaxDelta = Cells(s, 11).Value
                s_MaxNm = Cells(s, 9).Value
            End If
            If Cells(s, 11).Value < s_MinDelta Then
                s_MinDelta = Cells(s, 11).Value
                s_MinNm = Cells(s, 9).Value
            End If
            If Cells(s, 12).Value > s_MaxVol Then
                s_MaxVol = Cells(s, 12).Value
                s_MaxVolNm = Cells(s, 9).Value
            End If
        Next s
        
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        Cells(2, 16).Value = s_MaxNm
        Cells(2, 17).Value = s_MaxDelta
        Cells(2, 17).NumberFormat = "0.00%"
        Cells(3, 16).Value = s_MinNm
        Cells(3, 17).Value = s_MinDelta
        Cells(3, 17).NumberFormat = "0.00%"
        Cells(4, 16).Value = s_MaxVolNm
        Cells(4, 17).Value = s_MaxVol
        
        
    Next
    
End Sub
