Attribute VB_Name = "Module2"

Sub VBA()

Dim i As Long
Dim lastrow As Long
Dim Ticker As String



Range("I1") = "Ticker"


Range("J1") = "Yearly_Change"
Range("K1") = "Opening_Price"
Range("L1") = "Closing_Price"
Range("M1") = "Percent_Change"
Range("N1") = "Total_Stock_Volume"

lastrow = Cells(Rows.Count, "A").End(xlUp).Row

j = 2
    
    For i = 2 To lastrow
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            Ticker = Cells(i, 1)
            Cells(j, 9) = Ticker
            
            Closing_Price = Cells(i, 6).Value
            Cells(j, 12) = Closing_Price

        ElseIf Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
        
            Opening_Price = Cells(i, 3)
            Cells(j + 1, 11) = Opening_Price
            
        j = j + 1

        End If
    Next i

End Sub

Sub YearlyChange()
Dim Percent_Change As Double
lastrow = Cells(Rows.Count, "I").End(xlUp).Row

    For i = 3 To lastrow
        Yearly_Change = Cells(i, 12).Value - Cells(i, 11).Value
        Cells(i, 10) = Yearly_Change
        
        Percent_Change = Cells(i, 10) / Cells(i, 11)
        Cells(i, 13) = Percent_Change
    
        
    Next i
End Sub


Sub TotalVolume()
Range("N1") = "Total_Stock_Volume"
lastrow = Cells(Rows.Count, "A").End(xlUp).Row

j = 2
    
    For i = 2 To lastrow
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            Total_Stock_Volume = Cells(j + 1, 7).Value - Cells(j, 7).Value
            Cells(j + 1, 14) = Total_Stock_Volume
            
        j = j + 1
    
      End If
    Next i

End Sub

 
Sub Color()
lastrow = Cells(Rows.Count, "I").End(xlUp).Row
j = 3

  For i = 2 To lastrow
    If Cells(i, 13) > 0 Then
            Cells(i, 13).Interior.ColorIndex = 4
    Else
        Cells(i, 13).Interior.ColorIndex = 3

            
                
        j = j + 1
        
        End If
    Next i
        
End Sub

