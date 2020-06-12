Attribute VB_Name = "Module1"
'BootCamp - June, 2020. VBA to analyze meta stock data

'This script will loop through all the stocks for one year and output the following information.
'   * The ticker symbol.
'   * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
'   * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
'   * The total stock volume of the stock.
'   * You should also have conditional formatting that will highlight positive change in green and negative change in red.

Sub YearlyChange()

    Dim ticker As String
    Dim openprice, closeprice As Double
    Dim vol As Long
    Dim i As Long
        
    Dim tickername As String
    Dim changeprice As Double
    Dim totalvol As LongLong
    Dim ti As Long
    
    'Do each worksheet
    For Each ws In Worksheets
    
        'To set up the result table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        'To initialize the values for the result table
        ti = 2
        openprice = 0
        closeprice = 0
        totalvol = 0
   
        'To read the first row
        i = 2
        ticker = ws.Cells(i, 1).Value
    
        'Set the first ticker name as the current(first) ticker
        tickername = ticker
 
        While (ticker <> "")
 
            'Case when new ticker appears
            If (tickername <> ticker) Then
        
                'Finishing up writing the result table for the current ticker
            
                ws.Cells(ti, 9).Value = tickername
            
                changeprice = closeprice - openprice
                ws.Cells(ti, 10).Value = changeprice
            
                'Calculate the percentage of yearly change from open price
                ws.Cells(ti, 11).NumberFormat = "0.00%"
                If openprice = 0 Then
                    MsgBox ("Oh, No! Zero Open Price! : " & tickername & " in Sheet:" & ws.Name)
                    ws.Cells(ti, 11).Value = changeprice
                Else
                    'Calculate only the ratio because Cell format is percentage
                    ws.Cells(ti, 11).Value = (changeprice / openprice)
                End If
            
                ws.Cells(ti, 12) = totalvol
                
                'Fill the cell with red or green according to the sign of changeprice
                If changeprice < 0 Then
                    ws.Cells(ti, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(ti, 10).Interior.ColorIndex = 4
                End If
            
                'Update the index and initialize variables for the new ticker
                ti = ti + 1
                openprice = 0
                closeprice = 0
                totalvol = 0
                tickername = ticker
                
        End If
        
       'Read the price and vol for the current ticker
        If openprice = 0 Then
            openprice = ws.Cells(i, 3).Value
        End If
        closeprice = ws.Cells(i, 6).Value
        vol = ws.Cells(i, 7).Value
        totalvol = totalvol + vol
 
        'Update the index and read new ticker
        i = i + 1
        ticker = ws.Cells(i, 1).Value
              
    Wend
    
    Next ws
    
End Sub

