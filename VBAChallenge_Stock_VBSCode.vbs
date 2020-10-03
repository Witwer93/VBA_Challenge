Attribute VB_Name = "Module1"
Sub find_useful_stockdata()

    'Create a script that will loop through all the stocks for one year and output the following information.
    '   The ticker symbol.
    '   Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
    '   The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
    '   The total stock volume of the stock.
    '   You should also have conditional formatting that will highlight positive change in green and negative change in red.
    
    
    
    'this will track the next available row to store incoming values in
    Dim rowtrack As Double
    
    'necessary
    Dim ws As Worksheet
    Dim lastrow As Double
    
    'Price change variables
    Dim price1 As Double
    Dim price2 As Double
    Dim newprice As Double
    
    'percent change variable
    Dim perc As Double
    
    'total volume
    Dim volu As Double
    
    'variables for storing largest values
    Dim bigup As Double
    Dim bigdown As Double
    Dim bigvol As Double
    Dim bigtickI As String
    Dim bigtickD As String
    Dim bigtickV As String
    
    
    
    
    
    
    For Each ws In ThisWorkbook.Worksheets
        
        ws.Activate
        'calculate and save # of rows
        lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row
        'record initial ticker
        Cells(2, 9).Value = Cells(2, 1).Value
        'record initial opening price
        price1 = Cells(2, 3).Value
        'Create headers
        Cells(1, "I").Value = "Ticker"
        Cells(1, "J").Value = "Price Change"
        Cells(1, "K").Value = "Percent Change"
        Cells(1, "L").Value = "Total Stock Volume"
        Cells(2, "O").Value = "Greatest % Increase"
        Cells(3, "O").Value = "Greatest % Decrease"
        Cells(4, "O").Value = "Greatest Total Volume"
        Cells(1, "P").Value = "Ticker"
        Cells(1, "Q").Value = "Value"
        
        'variable to track current row storing data
        rowtrack = 2
        
        'the loop that does all the work
        'For i = 2 To 1000 ---> (for testing)
        For i = 2 To lastrow
            'begin recording stock volume
            volu = volu + Cells(i, "G").Value
            
            'when a new ticker is found, calculate/record/reset
            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                'record newly detected ticker
                Cells(rowtrack + 1, "I").Value = Cells(i + 1, 1).Value
                'record new opening price
                newprice = Cells(i + 1, "C").Value
                'record closing price
                price2 = Cells(i, "F").Value
                
                'temporarily save opening price in perc to calculate difference
                perc = price1
                
                'record price change and format
                price1 = price2 - price1
                Cells(rowtrack, "J").Value = price1
                If price1 > 0 Then
                    Cells(rowtrack, "J").Interior.ColorIndex = 4
                ElseIf price1 < 0 Then
                    Cells(rowtrack, "J").Interior.ColorIndex = 3
                End If
                
                'calculate and record percent change, if statement to avoid division by 0
                If perc = 0 Then
                    Cells(rowtrack, "K").Value = 0
                Else
                perc = (price1 / perc)
                Cells(rowtrack, "K").Value = perc
                End If
                
                'check for biggest/smallest percent changes
                If perc > bigup Then
                    bigup = perc
                    bigtickI = Cells(i, 1).Value
                ElseIf perc < bigdown Then
                    bigdown = perc
                    bigtickD = Cells(i, 1).Value
                End If
                
                'record total stock volume
                Cells(rowtrack, "L").Value = volu
                'check for biggest volume
                If volu > bigvol Then
                    bigvol = volu
                    bigtickV = Cells(i, 1)
                End If
                
                'reset variables
                volu = 0
                price1 = newprice
                rowtrack = rowtrack + 1
                
                'Exit For ----> (for testing)
            End If
        Next i
        'add greatest increase/decrease & volume values
        Cells(2, "P").Value = bigtickI
        Cells(3, "P").Value = bigtickD
        Cells(4, "P").Value = bigtickV
        Cells(2, "Q").Value = bigup
        Cells(3, "Q").Value = bigdown
        Cells(4, "Q").Value = bigvol
        
        'format percent change into percentages
        Range("K:K").NumberFormat = "0.00%"
        Cells(2, "Q").NumberFormat = "0.00%"
        Cells(3, "Q").NumberFormat = "0.00%"
        
        'make it all look pretty
        Range("I:Q").Columns.AutoFit
        'reset variables
        bigtickI = ""
        bigtickD = ""
        bigtickV = ""
        bigup = 0
        bigdown = 0
        bigvol = 0
        
    Next ws
End Sub

