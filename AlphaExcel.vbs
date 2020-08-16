Sub alpha()
'Setting up worksheet as a variable to go through all sheets in excel file at once
        Dim WorksheetName As String
        For Each ws In Worksheets
        WorksheetName = ws.Name
        ws.Activate
        
        'Determining the last row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Header Names
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        
        'Declaring loop *In progress*
        For i = 2 To LastRow
         ' Creating if then statement
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

            'Ticker
                TickerName = Cells(i, 1).Value
                Cells(2, 9).Value = TickerName
                
            'Closing and Open Price
                ClosingPrice = Cells(i, 6).Value
                OpenPrice = Cells(i, 3).Value
                
            'Yearly Change
                YearlyChange = ClosingPrice - OpenPrice
                Cells(2, 10).Value = YearlyChange
                
             'Percent Change *In progress*
                    
                ' Volume
                Volume = Volume + Cells(i, 7).Value
                Cells(2, 12).Value = Volume
            Else
               Volume = Volume + Cells(i, 7).Value
            End If
           
        Next i
        
        
    
       'Conditional Formatting for Yearly Change column
       'Declaring last row
        YearlyChangeLastRow = ws.Cells(Rows.Count, 10).End(xlUp).Row
        'Setting the Cell Colors
        For j = 2 To YearlyChangeLastRow
            If Cells(j, 10).Value >= 0 Then
                Cells(j, 10).Interior.ColorIndex = 4
            ElseIf Cells(j, 10).Value < 0 Then
                Cells(j, 10).Interior.ColorIndex = 3
            End If
        Next j

    
       'All work in first worksheet done, moving on to next worksheet
       Next ws


End Sub
