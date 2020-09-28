Attribute VB_Name = "Module1"
Sub alpha_test()

'Create a script that will loop through all the stocks for one year and output the following information.
'   The ticker symbol.
'   Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
'   The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
'   The total stock volume of the stock.
'   You should also have conditional formatting that will highlight positive change in green and negative change in red.

Dim lastrow As Double
Dim lastcol As Integer
'variable to track ticker changes
Dim ticker As String
'this will track the next available row to store incoming values in
Dim rowtrack As Double

'Dim ws As Worksheet

'Price change variables
Dim price1 As Double
Dim price2 As Double
'Percent change variables
Dim perc1 As Double
Dim perc2 As Double
'total volume
Dim volu As Double


'record initial ticker
Cells(2, 9).Value = Cells(2, 1).Value
ticker = Cells(2, 1).Value
lastrow = Cells(Rows.Count, "A").End(xlUp).Row

'loop through column A until a different value is detected
For i = 2 To lastrow
    
    'calculate new row to store values
    rowtrack = Cells(Rows.Count, "I").End(xlUp).Row + 1
    'save first price point if applicable
    If price1 = 0 Then
        price1 = Cells(i, "C").Value
    End If
    
    'begin recording stock volume
    volu = volu + Cells(i, "G").Value
    
    
    
    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
        'record newly detected ticker
        Cells(rowtrack, "I").Value = Cells(i + 1, 1).Value
        'record price change and reset price1 & 2 to zero
        
        
        'record total stock volume and reset volu to zero
        Cells(rowtrack - 1, "L").Value = volu
        volu = 0
        
        Exit For
    End If
Next i
   


End Sub
