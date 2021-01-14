Attribute VB_Name = "Module2"
Sub VBA_HW()

Dim ticker As String
  Dim lastrow As Long
  Dim OpenPrice As Double
  Dim Closeprice As Double
  Dim percentage As Variant
  Dim j As Integer
  Dim YearlyChange As Double
  Dim TotalVolume As Variant
  
  For Each Ws In Sheets
  If Ws.Visible Then Ws.Select
  
  
  'Initialize the variables
   TotalVolume = 0
   j = 2
   
   
   'Find the lastrow
   lastrow = Cells(Rows.Count, 1).End(xlUp).Row
   'Loop through the column A
   For i = 2 To lastrow
    If (Cells(i + 1, 1).Value <> Cells(i, 1).Value) Then
        ticker = Cells(i, "A").Value
        Closeprice = Cells(i, "F").Value
        TotalVolume = TotalVolume + Cells(i, 7).Value
        
        'calculate yearly change and percentage
        YearlyChange = Closeprice - OpenPrice
        If (OpenPrice <> 0) Then
            percentage = ((OpenPrice / Closeprice) - 1) * 100
        Else
           percentage = 0
        End If
        
        'Print on column I , J , K
        Range("I" & j).Value = ticker
        Range("J" & j).Value = YearlyChange
        Range("K" & j).Value = percentage
        Range("L" & j).Value = TotalVolume
        
        'inititalize for next ticker
        TotalVolume = 0
        j = j + 1
    Else
         OpenPrice = Cells(i, "C").Value
         TotalVolume = TotalVolume + Cells(i, 7).Value
         
    If YearlyChange < 0 Then
    
        Range("J" & j).Interior.ColorIndex = 3
        Else
        
        Range("J" & j).Interior.ColorIndex = 4
        End If
         
         
         
         
         
         
    End If
   Next i
   
   Next Ws

   
   
   
   
   

End Sub
