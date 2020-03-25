Attribute VB_Name = "Module2"
Sub HW2()
   
   
    'Set Headers
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Total Stock Volume"
   
    'Counter for print statement
    Dim Summary_Table_Row As Double
    Summary_Table_Row = 2
   
    'Set variables to hold ticker name & totalStockVolume
    Dim tickerName As String
    Dim tickerTotalVolume As Double
    tickerTotalVolume = 0
   
    For Each ws In Worksheets
       
        'Set tickerName
        tickerName = ws.Name
   
        'Find Last Row in current worksheet
        Dim lastRow As Double
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
   
        'Loop through and sum volumes
        Dim j As Double
        For j = 2 To lastRow
            tickerTotalVolume = tickerTotalVolume + ws.Cells(j, 7).Value
        Next j
   
        'add ticker name to table
        Range("I" & Summary_Table_Row).Value = tickerName
   
        'add volume to table
        Range("J" & Summary_Table_Row).Value = tickerTotalVolume
   
   
        'increment summary table row
        Summary_Table_Row = Summary_Table_Row + 1
       
        'reset tickerTotalVolume
        tickerTotalVolume = 0
   
    Next ws
   
   
   
End Sub
