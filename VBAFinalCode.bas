Attribute VB_Name = "Module1"
Sub Test_Stock()
 Dim x As String
    Dim YearChange As Single
    Dim PercentageChange As Single
    Dim vol As Double
    Dim i As Double
    Dim cont1 As Double
    Dim cont2 As Double
    Dim cont3 As Double
    Dim cont4 As Double
    Dim z As Double
    Dim avg1 As Double
    Dim avg2 As Double
    
    z = 2
    x = Cells(2, 1)
    
    cont1 = Cells(2, 3)
    cont2 = 0
    cont3 = 0
    cont4 = 0
    
    For i = 2 To 70000
    

     If Cells(i + 1, 1) = x Then
     cont3 = cont3 + Cells(i, 7)
    
    
    Else
    cont2 = Cells(i, 6)
    YearChange = cont2 - cont1
    PercentageChange = ((cont2 - cont1) / cont2) * 100
    vol = cont3
    Cells(z, 9) = x
    Cells(z, 10) = YearChange
    Cells(z, 11) = PercentageChange
    Cells(z, 12) = vol
    x = Cells((i + 1), 1)
   cont1 = Cells(i + 1, 3)
   
    z = z + 1
    
    End If
    
     Next i
   
    
End Sub
