Attribute VB_Name = "Module2"

'Let's find monthly stock averages per the year

Sub MonthlyAvg()

    Dim mth As String
    Dim mcnt(12) As Long
    Dim mvol(12) As LongLong
    
    Dim date_cell As String
    Dim i As Long
    
    Dim sheet As Worksheet
    Dim nrow As Integer
    
    Set sheet = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
    sheet.Name = "Monthly Summary"
 
    'In a new worksheet, write the result - month, total count for the month, total vol for the month
    sheet.Range("A" & 1).Value = "Year"
    sheet.Range("B" & 1).Value = "Month"
    sheet.Range("C" & 1).Value = "Tickers_Count"
    sheet.Range("D" & 1).Value = "Total Volume"
    sheet.Range("E" & 1).Value = "Avg. Volume"
    
    For Each ws In Worksheets
        Debug.Print ws.Name
        Debug.Print ws.Index
        
        If ws.Name = "Monthly Summary" Then
            Exit For
        End If
        
        i = 2
        
        'Read the first row in <date> column
        date_cell = ws.Range("B" & i).Value
        
        'Loop by reading the date column until it is empty string
        While date_cell <> ""
               
             'Take month from the date
             mth = Mid(date_cell, 5, 2)
             
             'Read the volume
             vol = ws.Range("G" & i).Value
             
            'update the sum and count
            mcnt(Int(mth) - 1) = mcnt(Int(mth) - 1) + 1
            mvol(Int(mth) - 1) = mvol(Int(mth) - 1) + vol
        
            i = i + 1
            date_cell = ws.Range("B" & i).Value
        Wend

    'For each worksheet, the monthly average is written in the new wsheet.
        For j = 0 To 11
            'Set the row index for each month of the year(given by sheet name)
            nrow = (ws.Index - 1) * 12 + j + 2
            
            sheet.Range("A" & nrow).Value = ws.Name
            sheet.Range("B" & nrow).Value = j + 1
            sheet.Range("C" & nrow).Value = mcnt(j)
            sheet.Range("D" & nrow).Value = mvol(j)
            sheet.Range("E" & nrow).Value = mvol(j) / mcnt(j)
        Next j

        'Initialize mcnt and mvol for the next worksheet
        For j = 0 To 11
            mcnt(j) = 0
            mvol(j) = 0
        Next j
    Next ws

End Sub

