Sub Title_Add_Headers()
'Sub to create title and calculate the values per Ticker
'Declare the variables on the sub
Dim ws As Worksheet

'loping through all worksheets in this workbook

For Each ws In ThisWorkbook.Worksheets
    'I need to exclude the Instructions workshet in my workbook
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
                
                'Use the sub CFormat to format column YearlyChange
                'Red when value is lower than zero
                'Green when value is eaqul o higher than zero
                
                CFormat ws.Range("K" & i)
                
          
            
            End If
        
        Next r
    End If
Next ws
MsgBox "Done"
End Sub

'Create a Sub to format the table

Sub CFormat(rng As Range)
    rng.Select
    rng.FormatConditions.Delete
    
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 5287936
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
End Sub
