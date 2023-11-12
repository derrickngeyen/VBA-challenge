Attribute VB_Name = "Module1"
Sub StockAnalysis():
       
       Dim Ticker As String
       Dim YearlyChange As Double
       Dim PercentChange As Double
       Dim TotalVolume As Double
       Dim LastRow As Long
       Dim SummaryRow As Long
       Dim OpeningPrice As Double
       Dim ClosingPrice As Double
       Dim ws As Worksheet
       Set ws = Sheets(1)
    
       
       
       
       
       
       
       
       
       
   For Each ws In ThisWorkbook.Sheets
       ws.Cells(1, 9).Value = "Ticker"
       ws.Cells(1, 10).Value = "Yearly Change"
       ws.Cells(1, 11).Value = "Percent Change"
       ws.Cells(1, 12).Value = "Total Stock Volume"
        
           
       
       LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
       
       
       SummaryRow = 2
      
       
  
       For i = 2 To LastRow
                 
           If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
               
               Ticker = ws.Cells(i, 1).Value
               
              
               OpeningPrice = ws.Cells(i, 3).Value
               
               
               ClosingPrice = ws.Cells(i, 6).Value
               
              
               YearlyChange = ClosingPrice - OpeningPrice
               
               
               If OpeningPrice <> 0 Then
                   PercentChange = (YearlyChange / OpeningPrice) * 100
               Else
                   PercentChange = 0
               End If
               
               
               ws.Cells(SummaryRow, 9).Value = Ticker
               ws.Cells(SummaryRow, 10).Value = YearlyChange
               ws.Cells(SummaryRow, 11).Value = PercentChange
               ws.Cells(SummaryRow, 12).Value = TotalVolume
              
               
               
               ws.Cells(SummaryRow, 11).NumberFormat = "0.00%"
               
               
               If YearlyChange > 0 Then
                   ws.Cells(SummaryRow, 10).Interior.Color = RGB(0, 255, 0)
               ElseIf YearlyChange < 0 Then
                   ws.Cells(SummaryRow, 10).Interior.Color = RGB(255, 0, 0)
               End If
               
               
               SummaryRow = SummaryRow + 1
                ws.Range("P1").Value = "Ticker"
                ws.Range("Q1").Value = "Value"
                ws.Range("O2").Value = "Greatest % Increase"
                ws.Range("O3").Value = "Greatest % Decrease"
                ws.Range("O4").Value = "Greatest Total Volume"

               'Define Ticker variable
                Ticker = " "
                Dim Ticker_volume As Double
                Ticker_volume = 0
               
               TotalVolume = 0
           End If
           
    
    
           TotalVolume = TotalVolume + ws.Cells(i, 7).Value
       Next i
       
     
    
       SummaryLastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
       
       
       Dim MaxPercentIncrease As Double
       Dim MaxPercent
     
     Next ws
    End Sub
    
      
