Attribute VB_Name = "Module1"
Sub MultipleYearStockData()
    
    'Declare "ws" as Worksheet
    Dim ws As Worksheet
    
    'Loop through each worksheet
    For Each ws In Worksheets
    
    'Label Column Headers and Tables
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    'Declare variables and set counter to default amounts
    Dim TickerName As String
    Dim LastRowA As Long
    Dim LastRowK As Long
    Dim TotalTickerVolume As Double
    TotalTickerVolume = 0
    Dim SummaryTableRow As Long
    SummaryTableRow = 2
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim YearlyChange As Double
    Dim PreviousAmount As Long
    PreviousAmount = 2
    Dim PercentChange As Double
    Dim GreatestIncrease As Double
    GreatestIncrease = 0
    Dim GreatestDecrease As Double
    GreatestDecrease = 0
    Dim LastRowValue As Long
    Dim GreatestTotalVolume As Double
    GreatestTotalVolume = 0

    'Determine value of the last row by finding the last non-blank cell in column A
    LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
    'Loop through rows
    For i = 2 To LastRowA

        'Add values to Total Ticker Volume
        TotalTickerVolume = TotalTickerVolume + ws.Cells(i, 7).Value
    
        'Check if the next row has the same ticker name as the previous row
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

            'Set Ticker Name for the first column
            TickerName = ws.Cells(i, 1).Value
                
            'Print Ticker Name in Summary Table at Column I
            ws.Range("I" & SummaryTableRow).Value = TickerName
                
            'Print Total Ticker Volume in Summary Table at Column L
            ws.Range("L" & SummaryTableRow).Value = TotalTickerVolume
               
            'Reset Total Ticker Volume
            TotalTickerVolume = 0

            'Set Yearly Open Price
            OpenPrice = ws.Range("C" & PreviousAmount)
                
            'Set Yearly Close Price
            ClosePrice = ws.Range("F" & i)
                
            'Set Value of Yearly Change
            YearlyChange = ClosePrice - OpenPrice
            ws.Range("J" & SummaryTableRow).Value = YearlyChange
                
            'Change format of Column J to Accounting with "$"
            ws.Range("J" & SummaryTableRow).NumberFormat = "$0.00"

            'Determine Percent Change, if Yearly Open Price is 0, then Percent Change is 0
            If OpenPrice = 0 Then
                PercentChange = 0
                    
                'Otherwise, set Percent Change to Yearly Change divided by Yearly Open Price
                Else
                YearlyOpen = ws.Range("C" & PreviousAmount)
                PercentChange = YearlyChange / OpenPrice
                        
            End If
                
            'Print Percent Change to Column K
            ws.Range("K" & SummaryTableRow).Value = PercentChange
                
            'Change format of Column K to Percentage with "%" and to the hundredths decimal place
            ws.Range("K" & SummaryTableRow).NumberFormat = "0.00%"

            'Conditional Formatting, if value is Positive, fill cell with Green
            If ws.Range("J" & SummaryTableRow).Value >= 0 Then
            ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 4
                    
                Else
                'Conditional Formatting, if value is Negative, fill cell with Red
                ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 3
                
            End If
            
            'Add 1 to Summary Table Row
            SummaryTableRow = SummaryTableRow + 1
              
            'Set Previous Amount
            PreviousAmount = i + 1
                
        End If
                
        'Go to next row
        Next i

        'Determine value of the last row by finding the last non-blank cell in column K
        LastRowK = ws.Cells(Rows.Count, 11).End(xlUp).Row
        
        'Loop through rows for final result table
        For i = 2 To LastRowK
            
            'Determine Greatest % Increase
            If ws.Range("K" & i).Value > ws.Range("Q2").Value Then
                ws.Range("Q2").Value = ws.Range("K" & i).Value
                ws.Range("P2").Value = ws.Range("I" & i).Value
                
            End If

            'Determine Greatest % Decrease
            If ws.Range("K" & i).Value < ws.Range("Q3").Value Then
                ws.Range("Q3").Value = ws.Range("K" & i).Value
                ws.Range("P3").Value = ws.Range("I" & i).Value
                    
            End If

            'Determine Greatest Total Volume
            If ws.Range("L" & i).Value > ws.Range("Q4").Value Then
                ws.Range("Q4").Value = ws.Range("L" & i).Value
                ws.Range("P4").Value = ws.Range("I" & i).Value
                    
            End If

            Next i
            
        'Change format of Q2 and Q3 to Percentage with "%" and to the hundredths decimal place
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"

    Next ws

End Sub
