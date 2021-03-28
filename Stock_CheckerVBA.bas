Attribute VB_Name = "Module1"
Sub stock_Counter():



Dim Summary_Table_Row As Integer
Summary_Table_Row = 2
Dim Total_Volume As Double
Total_Volume = 0
Dim Ticker As String
Dim year_open As Double
Dim year_closed As Double

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly_Change"
Cells(1, 11).Value = "Percent_Change"
Cells(1, 12).Value = "TotVolume"

 lastrow = Cells(Rows.Count, 1).End(xlUp).Row
 
 For i = 2 To lastrow
       
       If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
       
            Ticker = Cells(i, 1).Value
            TotVolume = TotVolume + Cells(i, 7).Value
            close_price = Cells(i, 5).Value
            Open_price = Cells(i, 2).Value
            Yearly_Change = (close_price - Open_price)
            Percent_Change = (Yearly_Change / Open_price)
            
            
            
            Range("I" & Summary_Table_Row).Value = Ticker
            Range("L" & Summary_Table_Row).Value = TotVolume
            Range("J" & Summary_Table_Row).Value = Yearly_Change
            Range("K" & Summary_Table_Row).Value = Percent_Change
            
            Summary_Table_Row = Summary_Table_Row + 1
            
            
         Else
            TotVolume = TotVolume + Cells(i, 7).Value
           
        End If
        Cells(i + 1, "K") = Format("0.00%")
        TotVolume = 0
        Open_price = 0
        close_price = 0
        Percent_Change = 0
        
        Next i
        
            
         
   
  
End Sub
 
