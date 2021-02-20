Attribute VB_Name = "Module1"
Sub StockAnalysis()
    Dim labels_1 As Variant
    labels_1 = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")
    For i = 0 To UBound(labels_1)
        Cells(1, 9 + i).Value = labels_1(i)
    Next i
    
    Dim labels_2 As Variant
    labels_2 = Array("Ticker", "Value")
    For i = 0 To UBound(labels_2)
        Cells(1, 16 + i).Value = labels_2(i)
    Next i
    
    Dim labels_3 As Variant
    labels_3 = Array("Greatest % Increase", "Greatest % Decrease", "Greatest Total Volume")
    For i = 0 To UBound(labels_3)
        Cells(2 + i, 15).Value = labels_3(i)
    Next i
    
    
    Dim LastRowStocks As Long
    LastRowStocks = Cells(Rows.Count, 1).End(xlUp).Row
    
    Dim PrintRow As Long
    Dim CurTicker As String
    Dim CurStartRow As Long
    Dim CurOpen As Double
    Dim CurClose As Double
    Dim CurChange As Double
    Dim CurPercentChg As Double
    Dim CurVolume As Double
    Dim MaxIncrease As Double
    Dim MaxIncTick As String
    Dim MaxDecrease As Double
    Dim MaxDecTick As String
    Dim MaxVolume As Double
    Dim MaxVolTick As String
    
    PrintRow = 2
    CurTicker = Cells(2, 1).Value
    CurStartRow = 2
    CurOpen = Cells(2, 3).Value
    MaxIncrease = 0
    MaxDecrease = 0
    MaxVolume = 0
    
    For i = 2 To LastRowStocks
        
        If CurTicker <> Cells(i + 1, 1).Value Then
            CurClose = Cells(i, 6).Value
            CurChange = CurClose - CurOpen
            Cells(PrintRow, 9).Value = CurTicker
            Cells(PrintRow, 10).Value = CurChange
            
            If CurChange >= 0 Then
                Cells(PrintRow, 10).Interior.Color = RGB(0, 255, 0)
            Else
                Cells(PrintRow, 10).Interior.Color = RGB(255, 0, 0)
            End If
            
            If CurChange = 0 Then
                CurPercentChg = 0
            ElseIf CurChange <> 0 And CurOpen = 0 Then
                CurPercentChg = 1
            Else
                CurPercentChg = CurChange / CurOpen
            End If
            
            Cells(PrintRow, 11).Value = CurPercentChg
            Cells(PrintRow, 11).NumberFormat = "0.00%"
            
            If CurPercentChg >= MaxIncrease Then
                MaxIncrease = CurPercentChg
                MaxIncTick = CurTicker
            ElseIf CurPercentChg <= MaxDecrease Then
                MaxDecrease = CurPercentChg
                MaxDecTick = CurTicker
            End If
            
            CurVolume = Application.Sum(Range(Cells(CurStartRow, 7), Cells(i, 7)))
            Cells(PrintRow, 12).Value = CurVolume
            If CurVolume >= MaxVolume Then
                MaxVolume = CurVolume
                MaxVolTick = CurTicker
            End If
           
            CurStartRow = i + 1
            CurTicker = Cells(i + 1, 1).Value
            CurOpen = Cells(i + 1, 3).Value
            PrintRow = PrintRow + 1
        End If
    Next i
    
    Range("P2").Value = MaxIncTick
    Range("Q2").Value = MaxIncrease
    Range("P3").Value = MaxDecTick
    Range("Q3").Value = MaxDecrease
    Range("Q2:Q3").NumberFormat = "0.00%"
    Range("P4").Value = MaxVolTick
    Range("Q4").Value = MaxVolume
    
    Columns("I:Q").AutoFit
End Sub
