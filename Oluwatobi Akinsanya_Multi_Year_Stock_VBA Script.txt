Sub StockMarket2()

For Each ws In Worksheets


' Created a Variable to Hold File Name, Last Row, Last Column, and Year

Dim Ticker As String
Dim yearly_change As String
Dim percent_change As Double
Dim openprice As Double
Dim closeprice As Double

Dim Tick As String
Dim Tick1 As String
Dim Tick2 As String
Dim Greatest_increase As Double
Dim Greatest_Decrease As Double


Dim LastRow As Long

total_stock = 0

Dim summary_row As Integer
summary_row = 2

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"




LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To LastRow
    
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
        Ticker = ws.Cells(i, 1).Value
        
        closeprice = ws.Cells(i, 6).Value
        
        yearly_change = closeprice - openprice
        
            If openprice <> 0 Then
            
                percent_change = (yearly_change / openprice) * 100
                
            Else
            
                percent_change = 0
            End If
                
        total_stock = total_stock + ws.Cells(i, 7).Value
        
        ws.Range("I" & summary_row).Value = Ticker
        
        ws.Range("J" & summary_row).Value = yearly_change
        
        ws.Range("K" & summary_row).Value = percent_change & "%"
        
        ws.Range("L" & summary_row).Value = total_stock
        
        If (ws.Range("J" & summary_row).Value < 0) Then
    
            ws.Range("J" & summary_row).Interior.ColorIndex = 3
         Else
    
            ws.Range("J" & summary_row).Interior.ColorIndex = 4
            
        End If
    
        summary_row = summary_row + 1
        
        total_stock = 0
    
    
    ElseIf (ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value) Then
        openprice = ws.Cells(i, 3).Value
    
    
    Else
    
        total_stock = total_stock + ws.Cells(i, 7).Value
    

    End If

  Next i
   


Last_row_2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
Greatest_increase = 0
Greatest_Decrease = 0
Greatest_Total = ws.Range("L2").Value
For j = 2 To Last_row_2
    If (ws.Range("K" & j).Value > Greatest_increase) Then
        Greatest_increase = ws.Range("K" & j).Value
        Tick = ws.Range("I" & j).Value
    
    
    ElseIf (ws.Range("K" & j).Value < Greatest_Decrease) Then
        Greatest_Decrease = ws.Range("K" & j).Value
        Tick1 = ws.Range("I" & j).Value
   
    End If
    
    If (ws.Range("L" & j).Value > Greatest_Total) Then
        Greatest_Total = ws.Range("L" & j).Value
        Tick2 = ws.Range("I" & j).Value
    End If
  
Next j
ws.Range("P2").Value = Tick
ws.Range("Q2").Value = Greatest_increase * 100 & "%"
ws.Range("P3").Value = Tick1
ws.Range("Q3").Value = Greatest_Decrease * 100 & "%"
ws.Range("P4").Value = Tick2
ws.Range("Q4").Value = Greatest_Total
ws.Range("I:Q").EntireColumn.AutoFit
Next ws

End Sub

