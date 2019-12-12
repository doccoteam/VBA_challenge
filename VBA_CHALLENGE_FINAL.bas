Attribute VB_Name = "Module1"
Sub StockChanges()

    Dim sht As Worksheet
    Dim i As Long
    Dim t As String
    Dim LastRow As Long
    Dim Ticker As String
    Dim Yearly_Change_Total As Double
    Dim TickerResult
    Dim FirstPrice, LastPrice As Double
    Dim TotalVolume As Double
    Dim TotalItems As Long
    Dim Summary_Table_Row As Integer
    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim GreatestTotalVolume As Double
    Dim rng As Range, cell As Range
    Dim sht2 As Worksheet
    Dim grincr As Double
    Dim tickergrincr As String
    Dim grdecr As Double
    Dim tickergrdecr As String
    Dim grtotalvolume As Double
    Dim tickergrtotalvolume As String
    
    
    
        
    grincr = 0
    tickergrincr = 0
    grdecr = 0
    tickergrdecr = 0
    grtotalvolume = 0
      
For Each ws In Worksheets
      
      Summary_Table_Row = 2
    
       
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        FirstPrice = ws.Cells(2, 3).Value
        
        TotalItems = 0
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
  
          
     For i = 2 To LastRow
     
        TotalVolume = TotalVolume + ws.Cells(i, 7).Value
    '    TotalItems = TotalItems + 1
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                
    
            Ticker = ws.Cells(i, 1).Value
                
            ws.Range("I" & Summary_Table_Row).Value = Ticker
            LastPrice = ws.Cells(i, 6).Value
            ws.Range("J" & Summary_Table_Row).Value = LastPrice - FirstPrice
            Yearly_Change_Total = ws.Range("J" & Summary_Table_Row).Value
            
            If FirstPrice <> 0 Then
            
                ws.Range("K" & Summary_Table_Row).Value = (Yearly_Change_Total / FirstPrice) * 1
                
              Else
              
              ws.Range("K" & Summary_Table_Row).Value = 0
              
            End If
              
            ws.Range("L" & Summary_Table_Row).Value = TotalVolume
            FisrtPrice = ws.Cells(i + 1, 1).Value
            
                ' Reset the variables
        TotalVolume = 0
        TotalItems = 0
    
            
            
      '      TickerResult = ws.Range("I" & Summary_Table_Row).Value
            
        
           'Add one to the summary table row
          Summary_Table_Row = Summary_Table_Row + 1
          
         
              
        End If
        
       
        
        Set sht1 = ActiveWorkbook.Sheets(1)
    
    If ws.Cells(i, 10).Value > 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 3
            
            Else
            ws.Cells(i, 10).Interior.ColorIndex = 4
            
        End If
                   
        If ws.Cells(i, 11).Value > grincr Then
        
                      
    
            grincr = ws.Cells(i, 11).Value
            tickergrincr = ws.Cells(i, 9).Value

                                
        
        End If
        
        If ws.Cells(i, 11).Value < grdecr Then
        
                      
    
            grdecr = ws.Cells(i, 11).Value
            tickergrdecr = ws.Cells(i, 9).Value
            
        End If
     
        If ws.Cells(i, 12).Value > grtotalvolume Then
        
    
            grtotalvolume = ws.Cells(i, 12).Value
            tickergrtotalvolume = ws.Cells(i, 9).Value
            
        End If
        
        
              
       
              
      Next i
      
      
        ws.Columns("K").NumberFormat = "0.00%"
        ws.Columns("I:L").AutoFit
        
             
                        
    Next ws
    
    sht1.Activate
               

        sht1.Range("O2").Value = "Greatest % Increase"
        sht1.Range("O3").Value = "Greatest % Decrease"
        sht1.Range("O4").Value = "Greatest Total Volume"
        sht1.Range("P1").Value = "Ticker"
        sht1.Range("Q1").Value = "Value"
        sht1.Range("P2").Value = tickergrincr
        sht1.Range("Q2").Value = grincr
        sht1.Range("Q2").NumberFormat = "0.00%"
        sht1.Range("P3").Value = tickergrdecr
        sht1.Range("Q3").Value = grdecr
        sht1.Range("Q3").NumberFormat = "0.00%"
        sht1.Range("P4").Value = tickergrtotalvolume
        sht1.Range("Q4").Value = grtotalvolume
    
    MsgBox ("Data Successfully Calculated")
    
 
       
End Sub

Sub challenge()

Dim grincr As Double
Dim tickergrincr As String
Dim grdecr As Double
Dim tickergrdecr As String
Dim grtotalvolume As Double
Dim tickergrtotalvolume As String




  ' Set sht = ActiveSheet
     
    
    Set sht1 = ActiveWorkbook.Sheets(1)
    
    
    grincr = 0
    tickergrincr = 0
    grdecr = 0
    tickergrdecr = 0
    grtotalvolume = 0
    
    
    For Each ws In Worksheets
    
        
    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
      For i = 2 To LastRow
                  
        If ws.Cells(i, 11).Value > grincr Then
        
                      
    
            grincr = ws.Cells(i, 11).Value
            tickergrincr = ws.Cells(i, 9).Value

                                
        
        End If
        
        If ws.Cells(i, 11).Value < grdecr Then
        
                      
    
            grdecr = ws.Cells(i, 11).Value
            tickergrdecr = ws.Cells(i, 9).Value
            
        End If
     
        If ws.Cells(i, 12).Value > grtotalvolume Then
        
    
            grtotalvolume = ws.Cells(i, 12).Value
            tickergrtotalvolume = ws.Cells(i, 9).Value
            
        End If
        
        
              
       Next i
       
   
    Next ws
    
       
    
       sht1.Activate
               

        sht1.Range("O2").Value = "Greatest % Increase"
        sht1.Range("O3").Value = "Greatest % Decrease"
        sht1.Range("O4").Value = "Greatest Total Volume"
        sht1.Range("P1").Value = "Ticker"
        sht1.Range("Q1").Value = "Value"
        sht1.Range("P2").Value = tickergrincr
        sht1.Range("Q2").Value = grincr
        sht1.Range("Q2").NumberFormat = "0.00%"
        sht1.Range("P3").Value = tickergrdecr
        sht1.Range("Q3").Value = grdecr
        sht1.Range("Q3").NumberFormat = "0.00%"
        sht1.Range("P4").Value = tickergrtotalvolume
        sht1.Range("Q4").Value = grtotalvolume
        
    
End Sub
