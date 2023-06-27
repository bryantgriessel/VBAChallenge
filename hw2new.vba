Dim Array1(300000) As Double
Dim x As Double
Dim a As Integer
Dim i As Double
Dim Totalprice As Double
Dim Closeprice As Double
Dim Openprice As Double
Dim Gmax As Double
Dim Gmin As Double
Dim Gtot As Double
Dim Tvalmin As String
Dim Tvalmax As String
Dim Tvaltot As String
Dim Lastr As Double
'Dim Grange As Range
'Dim Frange As Range
'Dim Grow As Long
'Dim Tval As String

'Dim Gadr As String




Sub stockstuff2()

    For Each ws In Worksheets
        'assign/init some variables
        first = 0
        x = 0
        a = 2
        Lastr = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'print headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        'The following code uses the first For loop to change collumns and the second for loop to change rows
        For i = 2 To Lastr
            'Check if the current cell in row i of for loop is the same as the next cell
            If Not (ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value) Then
                Array1(x) = ws.Cells(i, 3).Value
                
                If first = 0 Then
                    first = 1
                    Firstvol = Cells(i, 7).Address
                End If
                x = x + 1
            'end of ticker so lets assign values and write to cells for each row a
    
            Else
                Array1(x) = ws.Cells(i, 3).Value
                'Assign values to variables
                x = 0
                first = 0
                ticker = ws.Cells(i, 1).Value
                Openprice = CDbl(Array1(0))
                Closeprice = ws.Cells(i, 6).Value
                Totalprice = Closeprice - Openprice
                Percentprice = (Closeprice - Openprice) / Openprice
                Currentvol = ws.Cells(i, 7).Address
                Totalvol = Application.WorksheetFunction.Sum(ws.Range(Firstvol, Currentvol))
                'Assign value to cells
                ws.Cells(a, 9).Value = ticker
                ws.Cells(a, 10).Value = Totalprice
                ws.Cells(a, 11).Value = FormatPercent(Percentprice)
                ws.Cells(a, 12).Value = Totalvol
                a = a + 1
            End If
    
        Next i
        Erase Array1()
        'print cell titles
        
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        Gmax = ws.Cells(2, 11).Value
        Gmin = ws.Cells(2, 11).Value
        Gtot = ws.Cells(2, 12).Value
        
        'determine last row in col 11 and then loop to assign compare cell values to previous values
        
        Lastr = ws.Cells(Rows.Count, 11).End(xlUp).Row
        For j = 2 To Lastr
            If (ws.Cells(j, 11).Value > Gmax) Then
                Gmax = ws.Cells(j, 11).Value
                Tvalmin = ws.Cells(j, 9)
            End If
            If (ws.Cells(j, 11).Value < Gmin) Then
                Gmin = ws.Cells(j, 11).Value
                Tvalmax = ws.Cells(j, 9)
            End If
            If (ws.Cells(j, 12).Value > Gtot) Then
                Gtot = ws.Cells(j, 12).Value
                Tvaltot = ws.Cells(j, 9)
            End If
            If ws.Cells(j, 10).Value > 0 Then
                ws.Cells(j, 10).Interior.Color = vbGreen
            ElseIf ws.Cells(j, 10).Value < 0 Then
                ws.Cells(j, 10).Interior.Color = vbRed
            End If
        Next j
        'BUNCH OF MUMBO JUMBO I DID NOT WANT TO  DELETE YET want to figure out another method of doing this
        
        'Set Grange = Range("K2", ws.Cells(Lastr, 11).Address)
        'Gmax = Application.WorksheetFunction.Max(Grange)
        'Set Frange = Grange.Find(what:=Gmax)
        'Grow = Frange.Row
       ' Tval = ws.Cells(Grow, 9).Value
        'If Frange Is Nothing Then
        '    Debug.Print "Name was not found."
        'Else
         '   Debug.Print "Name found in :" & Frange.Address
        'End If
        'Gadr = Frange.Address
        'Tadr = Range(Gmax).Offset(0, 2).Value
        
        'END OF MUMBO JUMBO
        
        
        ws.Cells(2, 16).Value = Tvalmax
        ws.Cells(3, 16).Value = Tvalmin
        ws.Cells(4, 16).Value = Tvaltot
        ws.Cells(2, 17).Value = FormatPercent(Gmax)
        ws.Cells(3, 17).Value = FormatPercent(Gmin)
        ws.Cells(4, 17).Value = Gtot
        
        'Makes the columns wider to fit whatever data is inside
        ws.Columns("I:L").AutoFit
        ws.Columns("O:Q").AutoFit
        
    Next ws
                  
    
End Sub




