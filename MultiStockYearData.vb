Sub Stock_Market_Analysis()

    For Each ws In Worksheets
    
    'Add Column Headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Value"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "% Change"
    ws.Range("L1").Value = "Total Stock Vol"
    ws.Range("N2").Value = "Greatest % Incr"
    ws.Range("N3").Value = "Greatest % Decr"
    ws.Range("N4").Value = "Greatest Total Vol"
    
    
    'Variable Declaration
    Dim Ticker_symbol As String
    Dim LastRowA As Long
    Dim LastRowK As Long
    Dim Total_Volume As Double
    Total_Volume = 0
    Dim SummaryTableRow As Long
    SummaryTableRow = 2
    Dim Open_Price As Double
    Dim Close_Price As Double
    Dim YearlyChange As Double
    Dim PreviousAmount As Long
    PreviousAmount = 2
    Dim Percent_Change As Double
    Dim Greatest_increase As Double
    Greatest_increase = 0
    Dim Greatest_decrease As Double
    Greatest_decrease = 0
    Dim LastRowValue As Long
    Dim Greatest_Total_Volume As Double
    Greatest_Total_Volume = 0
    
    'Determine last row
    Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
    For i = 2 To Lastrow

        'Add values to Ticker volume
        Total_Volume = Total_Volume + ws.Cells(i, 7).Value
    
        'Check if ticker is same or different
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

            'Get ticker name
            Ticker_symbol = ws.Cells(i, 1).Value
                
            'Print ticker Name in Summary Table
            ws.Range("I" & SummaryTableRow).Value = Ticker_symbol
                
            'Print Total Ticker Volume in Summary Table
            ws.Range("L" & SummaryTableRow).Value = Total_Volume
               
            'Reset Total Ticker Volume
            Total_Volume = 0

            'Opening Price
            Open_Price = ws.Range("C" & PreviousAmount)
                
            'Closing Price
            Close_Price = ws.Range("F" & i)
                
            'Set Value of Yearly Change
            YearlyChange = Close_Price - Open_Price
            ws.Range("J" & SummaryTableRow).Value = YearlyChange
                
            'Change format of Column J to Accounting with "$"
            ws.Range("J" & SummaryTableRow).NumberFormat = "$0.00"

            'Determine Percent Change, if Yearly Open Price is 0, then Percent Change is 0
            If Open_Price = 0 Then
                Percent_Change = 0
                    
                'Otherwise, set Percent Change to Yearly Change divided by Yearly Open Price
                Else
                YearlyOpen = ws.Range("C" & PreviousAmount)
                Percent_Change = YearlyChange / Open_Price
                        
            End If
                
            'Print % Change to Column K
            ws.Range("K" & SummaryTableRow).Value = Percent_Change
                
            'Change format of Column K to Percentage
            ws.Range("K" & SummaryTableRow).NumberFormat = "0.00%"

            'If value is Positive, fill cell with Green(4)
            If ws.Range("J" & SummaryTableRow).Value >= 0 Then
            ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 4
                    
                Else
                'If value is Negative, fill cell with Red(3)
                ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 3
                
            End If
            
            'Update Summary Table Row
            SummaryTableRow = SummaryTableRow + 1
              
            Previous Amount
            PreviousAmount = i + 1
                
        End If
                
        Next i

    Next ws

End Sub
