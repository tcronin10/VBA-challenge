Attribute VB_Name = "Module1"
Sub StockData():

Dim i, z As Long
Dim totalvolume As Double
Dim ticker As String
Dim openprice As Double
Dim closeprice As Double
Dim yearchange As Double
Dim percentchange As Double
Dim lastrow As Long
Dim openprice_row As Long


For Each ws In Worksheets


ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

totalvolume = 0
z = 2
openprice_row = 2

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastrow
    If ws.Range("A" & i + 1).Value = ws.Range("A" & i).Value Then
        totalvolume = totalvolume + ws.Range("G" & i).Value
    
    Else
        ticker = ws.Range("A" & i).Value
        openprice = ws.Range("C" & openprice_row)
        closeprice = ws.Range("F" & i)
        yearchange = closeprice - openprice
    
    If openprice = 0 Then
        percentchange = 0
    Else
        percentchange = yearchange / openprice
    End If
    
    ws.Range("I" & z).Value = ticker
    ws.Range("L" & z).Value = totalvolume + ws.Range("G" & i).Value
    ws.Range("J" & z).Value = yearchange
    ws.Range("K" & z).Value = percentchange
    ws.Range("K" & z).NumberFormat = "0.00%"
    
    z = z + 1
    totalvolume = 0
    openprice_row = i + 1
    
         
   End If
   
   If ws.Range("J" & z).Value > 0 Then
          ws.Range("J" & z).Interior.ColorIndex = 4
        Else
          ws.Range("J" & z).Interior.ColorIndex = 3
    End If
  
Next i
Next ws
End Sub
  


        
    

    
    








