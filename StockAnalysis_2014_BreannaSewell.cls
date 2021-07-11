VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub StocksAnalysis()


    ' set the variables
    Dim i As Long
    Dim Ticker As String
    Dim SummaryTableRow As Long
    Dim YearRange As Double
    Dim PercentChange As Double
    Dim StockVolume As Double
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim vDate As Long
    
    ' create headers
    Range("J1").Value = "Ticker"
    Range("K1").Value = "Year Range"
    Range("L1").Value = "Percent Change"
    Range("M1").Value = "Stock Volume - Thousands"
    Range("N1").Value = "Stock Volume"
    Range("O1").Value = "Close Price"
    Range("P1").Value = "Open Price"
    
    'value the variables
    SummaryTableRow = 2
    StockVolume = 0
    vDate = 20140101
    
    ' determine last row
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    ' loop through the stock instances
    For i = 2 To LastRow
        
        ' check if it is a new ticker
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            ' add ticker to the summary table
            Ticker = Cells(i, 1).Value
            
            ' add to total stock volume
            StockVolume = StockVolume + Cells(i, 8).Value
                
            ' populate summary table
            Range("J" & SummaryTableRow).Value = Ticker
            Range("M" & SummaryTableRow).Value = StockVolume
            
            ' reset the volume total
            StockVolume = 0
            
            'find close price
            ClosePrice = Cells(i, 6).Value
            
            'create close price columns
            Range("O" & SummaryTableRow).Value = ClosePrice
            
            ' start a new row in the summary table
            SummaryTableRow = SummaryTableRow + 1
            
        ElseIf Cells(i, 2) = vDate Then
            OpenPrice = Cells(i, 3).Value
            Range("P" & SummaryTableRow).Value = OpenPrice
            
        
        ' if the ticker has not changed...
        Else
            
            Ticker = Cells(i, 1).Value
             
            'add the volume to the volume total
            StockVolume = StockVolume + Cells(i, 8).Value
            
            Range("J" & SummaryTableRow).Value = Ticker
            Range("M" & SummaryTableRow).Value = StockVolume
            
        End If
    
    Next i
    
    
    For i = 2 To 290
        TotalVolume = Range("M" & i).Value * 1000
        Range("N" & i) = TotalVolume
        
        PriceRange = Range("O" & i).Value - Range("P" & i).Value
        Range("K" & i) = PriceRange
        
        PercentChange = PriceRange / Range("P" & i).Value
        Range("L" & i) = PercentChange
        
            If Range("K" & i).Value > 0 Then
                Range("K" & i).Interior.ColorIndex = 4
            
            Else
                Range("K" & i).Interior.ColorIndex = 3
            
            End If
        
    Next i
       

End Sub



