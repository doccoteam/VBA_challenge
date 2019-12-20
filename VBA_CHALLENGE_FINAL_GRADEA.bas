Attribute VB_Name = "Module1"
Sub StockChanges()

    Dim ws1 As Worksheet
    Dim sht As Worksheet
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim t As String
    Dim LastRow As Long
    Dim Ticker As String
    Dim Yearly_Change_Total As Double
    Dim TickerResult
    Dim FirstPrice, LastPrice As Double
    Dim TotalVolume As Double
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
        
         
       
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
  
          
     For i = 2 To LastRow
     
      
        FirstPrice = ws.Cells(i, 3).Value
        
        For j = 2 To LastRow
        
        TotalVolume = TotalVolume + ws.Cells(j, 7).Value
    
        If ws.Cells(j + 1, 1).Value <> ws.Cells(j, 1).Value Then
                
            
    
            Ticker = ws.Cells(j, 1).Value
                
            ws.Range("I" & Summary_Table_Row).Value = Ticker
            LastPrice = ws.Cells(j, 6).Value
            ws.Range("J" & Summary_Table_Row).Value = LastPrice - FirstPrice
            Yearly_Change_Total = ws.Range("J" & Summary_Table_Row).Value
            
            
            If FirstPrice <> 0 Then
            
                ws.Range("K" & Summary_Table_Row).Value = (Yearly_Change_Total / FirstPrice) * 1
                
              Else
              
              ws.Range("K" & Summary_Table_Row).Value = 0
              
            End If
              
            ws.Range("L" & Summary_Table_Row).Value = TotalVolume
                 
            i = j + 1
            
           ' Reset the variables
        TotalVolume = 0
            
        
           'Add one to the summary table row
          Summary_Table_Row = Summary_Table_Row + 1
        
              
        End If
        
        Next j
              
      Next i
        
    Next ws
        
For Each ws1 In Worksheets
        
        For k = 2 To LastRow
        
                ' Changing colors
                
                If ws1.Cells(k, 10).Value > 0 Then
                    ws1.Cells(k, 10).Interior.ColorIndex = 4
                        
                 Else
                    ws1.Cells(k, 10).Interior.ColorIndex = 3
                        
                 End If
                               
                 'greatest increase value
                 If ws1.Cells(k, 11).Value > grincr Then
                    
                                  
                    grincr = ws1.Cells(k, 11).Value
                    tickergrincr = ws1.Cells(k, 9).Value
            
                                           
                    
                End If
                
                 'greatest decrease value
                    
                If ws1.Cells(k, 11).Value < grdecr Then
                    
                                  
                
                    grdecr = ws1.Cells(k, 11).Value
                    tickergrdecr = ws1.Cells(k, 9).Value
                
                End If
                 
                 'greatest total volume value
                 
                If ws1.Cells(k, 12).Value > grtotalvolume Then
                    
                
                    grtotalvolume = ws1.Cells(k, 12).Value
                    tickergrtotalvolume = ws1.Cells(k, 9).Value
                        
                  End If
            
        Next k
              
 
        ws1.Columns("K").NumberFormat = "0.00%"
        ws1.Columns("I:L").AutoFit
        
             
                        
Next ws1
    
    Set sht1 = ActiveWorkbook.Sheets(1)
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
