Attribute VB_Name = "Module1"
Sub Stock():

For Each ws In Worksheets


ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
    


Dim Ticker As String
Dim LastRow As Long
        
Dim Total_Stock_Volume As Double
Total_Stock_Volume = 0
        
Dim Summary_Table As Long
Summary_Table = 2
        
Dim Open_Price As Double
Dim Close_Price As Double
        
Dim Yearly_Change As Double
Dim Percent_Change As Double

Dim Last_Price As Long
Last_Price = 2
       
        


LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
For i = 2 To LastRow

Total_Stock_Volume = Total_Stocke_Volume + ws.Cells(i, 7).Value

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
    Ticker = ws.Cells(i, 1).Value
          
    ws.Range("I" & Summary_Table).Value = Ticker
       
    ws.Range("L" & Summary_Table).Value = Total_Stock_Volume

    Total_Stock_Volume = 0


    Open_Price = ws.Range("C" & Last_Price)
    Close_Price = ws.Range("F" & i)
    Yearly_Change = Close_Price - Open_Price
    ws.Range("J" & Summary_Table).Value = Yearly_Change

    
If Open_Price = 0 Then
    Percent_Change = 0
Else
    Open_Price = ws.Range("C" & Last_Price)
    Percent_Change = Yearly_Change / Open_Price
End If

ws.Range("K" & Summary_Table).NumberFormat = "00.00%"
ws.Range("K" & Summary_Table).Value = Percent_Change

                
If ws.Range("J" & Summary_Table).Value >= 0 Then
    ws.Range("J" & Summary_Table).Interior.ColorIndex = 4
Else
    ws.Range("J" & Summary_Table).Interior.ColorIndex = 3
End If
            
              
Summary_Table = Summary_Table + 1
Last_Price = i + 1
End If
Next i

LastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
    
            
        
Next ws

End Sub

