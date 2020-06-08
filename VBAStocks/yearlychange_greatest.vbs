Sub YearlyChange_Greatest()

    Dim ticker As String
    Dim openprice, closeprice As Double
    Dim vol As Long
    Dim i As Long
        
    Dim tickername As String
    Dim changeprice As Double
    Dim percent As Double
    Dim totalvol As LongLong
    Dim ti As Long
    
    Dim H_increase As Double
    Dim H_decrease As Double
    Dim H_totalvol As LongLong
    Dim i_ticker As String
    Dim d_ticker As String
    Dim v_ticker As String

    
    'To set up the result table
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    'To read the first row
    i = 2
    ticker = Cells(i, 1).Value
    
    'To initialize the values for the yearly change for each ticker
    ti = 2
    tickername = ticker
    openprice = 0
    closeprice = 0
    totalvol = 0

    'To initialize the values for the greatest values
    H_increase = 0
    H_decrease = 0
    H_totalvol = 0

    While (ticker <> "")
    
    'Debug.Print ticker
    
    
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
                percent = (changeprice / openprice)
                Cells(ti, 11).Value = percent
                Cells(ti, 11).NumberFormat = "0.00%"
            End If
            
            Cells(ti, 12) = totalvol
            
            'Check whether the values hit the greatest value
            If H_increase < percent Then
                H_increase = percent
                i_ticker = tickername
            End If
            
            If H_decrease > percent Then
                H_decrease = percent
                d_ticker = tickername
            End If
            
            If H_totalvol < totalvol Then
                H_totalvol = totalvol
                v_ticker = tickername
            End If
            
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
    
    Cells(1, 15).Value = ""
    Cells(1, 16).Value = "ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(2, 16).Value = i_ticker
    Cells(2, 17).Value = H_increase
    Cells(2, 17).NumberFormat = "0.00%"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(3, 16).Value = d_ticker
    Cells(3, 17).Value = H_decrease
    Cells(3, 17).NumberFormat = "0.00%"
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(4, 16).Value = v_ticker
    Cells(4, 17).Value = H_totalvol

End Sub
    

