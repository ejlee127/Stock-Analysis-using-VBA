'BootCamp - June, 2020. VBA to analyze meta stock data

Sub YearlyChange_Color()

    Dim ticker As String
    Dim openprice, closeprice As Double
    Dim vol As Long
    Dim i As Long
    
    
    Dim tickername As String
    Dim changeprice As Double
    Dim totalvol As LongLong
    Dim ti As Long
    
    'To set up the result table
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    'To read the first row
    i = 2
    ticker = Cells(i, 1).Value
    
    'To initialize the values for the result table
    ti = 2
    tickername = ticker
    openprice = 0
    closeprice = 0
    totalvol = 0

    While (ticker <> "")
 
        'Case when new ticker appears
        If (tickername <> ticker) Then
        
            'Finishing up writing the result table for the current ticker
            
            Cells(ti, 9).Value = tickername
            
            changeprice = closeprice - openprice
            Cells(ti, 10).Value = changeprice
            
            'Fill the cell with red or green according to the sign of changeprice
            If changeprice < 0 Then
                Cells(ti, 10).Interior.ColorIndex = 3
            Else
                Cells(ti, 10).Interior.ColorIndex = 4
            End If
            
            'Calculate the percentage of yearly change from open price
            If openprice = 0 Then
                MsgBox ("Oh, No! Zero Open Price!" & tickername)
            Else
                'Calculate only the ratio because Cell format is percentage
                Cells(ti, 11).Value = (changeprice / openprice)
                Cells(ti, 11).NumberFormat = "0.00%"
            End If
            
            Cells(ti, 12) = totalvol
            
            'Update the index and initialize variables for the new ticker
            ti = ti + 1
            tickername = ticker
            openprice = 0
            closeprice = 0
            totalvol = 0
            
        End If
        
       'Read the price and vol for the current ticker
        If openprice = 0 Then
            openprice = Cells(i, 3).Value
        End If
        closeprice = Cells(i, 6).Value
        vol = Cells(i, 7).Value
        totalvol = totalvol + vol
 
        'Update the index and read new ticker
        i = i + 1
        ticker = Cells(i, 1).Value
              
    Wend
    
End Sub



