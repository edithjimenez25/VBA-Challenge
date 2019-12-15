Sub StockMarketAnalysis()
'Sub to create ticker tle and calculate the values per Ticker
'Declare the variables on the sub
'Teacher and TAs I'm declare Dim when I require in the code, it is more easy for me to understand the sequence
Dim ws As Worksheet

'loping through all worksheets in this workbook

For Each ws In ThisWorkbook.Worksheets
    'I need to exclude the Instructions worksheet in my workbook
    If ws.Name <> "Instructions" Then
        ws.Activate
    'start to write the title for my table
        ws.Cells(1, 10).Value = "Ticker"
        ws.Cells(1, 11).Value = "Yearly Change"
        ws.Cells(1, 12).Value = "% Change"
        ws.Cells(1, 13).Value = "Total Stock Volume"
            
        'Declare the variable for Ticker loop
      
        Dim TickerName As String
        Dim OpenPrice As Double
        Dim ClosePrice As Double
       
        
        'declare variable r for rows
        Dim r As Long 'Data row
        Dim i As Long 'Summary row
        i = 1
        
  'looping through each row to find the last row
        For r = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
            
            'define start row and find the change on the ticker and do calcs
            'add loop to look for ticker row
            
            If ws.Cells(r - 1, 1).Value <> ws.Cells(r, 1).Value Then
                TickerName = ws.Cells(r, 1).Value
                OpenPrice = ws.Cells(r, 3).Value
            
            'define variable for the top row
                Dim rTop As Long
                rTop = r
            End If
            
            If ws.Cells(r + 1, 1).Value <> ws.Cells(r, 1).Value Then
                i = i + 1
                
                ClosePrice = ws.Cells(r, 6).Value
                
                'First column is Ticker
                ws.Range("J" & i).Value = TickerName
                
                'Second Column is Yearly Change
                ws.Range("K" & i).Value = ClosePrice - OpenPrice
                
                'Third Column is Percent Change
                'define condition when openPrice =0 and avoid the error of dividing by zero
                
                If OpenPrice = 0 Then
                    ws.Range("L" & i).Value = 0
                
                'continue to the next tickers and calculate the % Change for the ticker
                
                Else
                    ws.Range("L" & i).Value = (ClosePrice - OpenPrice) / OpenPrice
                    'change the Style of the Column %Change to Percent
                    ws.Range("L" & i).Style = "Percent"
                End If
                
                'determine the Total Stock volume transaction
                ws.Range("M" & i).Value = Excel.WorksheetFunction.Sum(ws.Range("G" & rTop & ":G" & r))
                
                'Use the sub CellFormat to format column YearlyChange
                'Red when value is lower than zero
                'Green when value is equal o higher than zero
                
                'CFormat ws.Range("K" & i)
                
          
            End If
        
        Next r
        'define variable for cell to format
            Dim YearlyChange As Long
            YearlyChange = ws.Cells(Rows.Count, 11).End(xlUp).Row
            
            'include a loop for format the cells in column 11 until the last row on the table.
            'the row start in 2 until the last row
            'Red when value is lower than zero (code is 3)
            'Green when value is equal o higher than zero (Code is 4)
            For i = 2 To YearlyChange
                If ws.Cells(i, 11).Value >= 0 Then
                    ws.Cells(i, 11).Interior.ColorIndex = 4
                Else
                    ws.Cells(i, 11).Interior.ColorIndex = 3
                End If
            Next i
            
    End If

'======================================
'Challenge
'Write the titles of the table
        'Columns
        ws.Cells(2, 16).Value = "Greatest % Increase"
        ws.Cells(3, 16).Value = "Greatest % Decrease"
        ws.Cells(4, 16).Value = "Greatest Total Volume"
        'Rows
        ws.Cells(1, 17).Value = "Ticker"
        ws.Cells(1, 18).Value = "Value"
'Formatting the Titles
        'Columns
        ws.Cells(2, 16).Font.FontStyle = "Bold"
        ws.Cells(3, 16).Font.FontStyle = "Bold"
        ws.Cells(4, 16).Font.FontStyle = "Bold"
        'Rows
        ws.Cells(1, 17).Font.FontStyle = "Bold"
        ws.Cells(1, 18).Font.FontStyle = "Bold"

'Define Variable to be used in the calculus of the Maximum and Minimum Value
    Dim RowLastPercentValue As Long
        RowLastPercentValue = ws.Cells(Rows.Count, 12).End(xlUp).Row
    Dim MaximumValue As Double
        MaximumValue = 0
    Dim MinimumValue As Double
        MinimumValue = 0
    
'start a loop to find the maximum and minimum values
    For i = 2 To RowLastPercentValue

'creating a conditional to check the maximum values
        If MaximumValue < ws.Cells(i, 12).Value Then
            MaximumValue = ws.Cells(i, 12).Value
            ws.Cells(2, 18).Value = MaximumValue
            ws.Cells(2, 17).Value = ws.Cells(i, 10).Value
        ElseIf MinimumValue > ws.Cells(i, 12).Value Then
            MinimumValue = ws.Cells(i, 12).Value
            ws.Cells(3, 18).Value = MinimumValue
            ws.Cells(3, 17).Value = ws.Cells(i, 10).Value
                  
            'Format Maximum and Mimimun Values in the Column R
                ws.Cells(2, 18).Style = "Percent"
                ws.Cells(3, 18).Style = "Percent"
        End If

    Next i

'======================================
'
'Define Variable to be used in the greatest value
    Dim GreatestTotalVolumeRow As Long
        GreatestTotalVolumeRow = ws.Cells(Rows.Count, 13).End(xlUp).Row
    Dim MaximumTotalVolume As Double
        MaximumTotalVolume = 0

'start a loop to find the Greatest Total Volume
    For i = 2 To GreatestTotalVolumeRow
 
'creating a conditional to check the Greatest Total Volume Row
        If MaximumTotalVolume < ws.Cells(i, 13).Value Then
            MaximumTotalVolume = ws.Cells(i, 13).Value
            ws.Cells(4, 18).Value = MaximumTotalVolume
            ws.Cells(4, 17).Value = ws.Cells(i, 10).Value
        End If
    
    Next i

'======================================

Next ws

MsgBox "===== Work Done!!!! ===Enjoy Using VBA==="

End Sub


