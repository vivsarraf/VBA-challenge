Attribute VB_Name = "Module1"
Sub Ticker_StockData():

       ' --------------------------------------------
       ' LOOP THROUGH ALL SHEETS
       ' --------------------------------------------
      
        For Each ws In Worksheets
        
        'Variable to hold Worskheets
        Dim WorksheetExcel As String
         
        'Variable to hold Current row of Ticker Stockdata
        Dim i As Long
        
        'Variable to hold the Starting row of Ticker Stockdata
        Dim j As Long
        
        'Variable to hold Ticker
        Dim Ticker As Long
        
        'Variable for greatest total volume
        Dim GreatVolume As Double
        
        'Variable to hold Last row column A
        Dim LastRowColA As Long
        
        'Variable to hold last row column I
        Dim LastRowColI As Long
        
        'Variable for percent change calculation
        Dim PercentChange As Double
        
        'Variable for greatest increase calculation
        Dim GreatIncrease As Double
        
        'Variable for greatest decrease calculation
        Dim GreatDecrease As Double
                
        'Grabbed the WorksheetName
        WorksheetExcel = ws.Name
        
        'Create column headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        'Set Ticker Counter to first row
        Ticker = 2
        
        'Set start row to 2
        j = 2
        
        'Find the last non-blank cell in column A
        LastRowColA = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'MsgBox ("Last row in column A is " & LastRowA)
        
            'Loop through all rows
            For i = 2 To LastRowColA
            
                'Check if ticker name changed
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'Write ticker in column I (#9)
                ws.Cells(Ticker, 9).Value = ws.Cells(i, 1).Value
                
                'Calculate and write Yearly Change (close Price - openprice)
                ws.Cells(Ticker, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                
                'Add the Currency Format to Yearly Change
                ws.Cells(Ticker, 10).Style = "Currency"
                
                    'Conditional formating
                    If ws.Cells(Ticker, 10).Value < 0 Then
                
                    'Set cell background color to red
                    ws.Cells(Ticker, 10).Interior.ColorIndex = 3
                
                    Else
                
                    'Set cell background color to green
                    ws.Cells(Ticker, 10).Interior.ColorIndex = 4
                
                    End If
                    
                    'Calculate and write percent change in column K
                    If ws.Cells(j, 3).Value <> 0 Then
                    PercentChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                    
                    'Percent formating
                    ws.Cells(Ticker, 11).Value = Format(PercentChange, "Percent")
                    
                    Else
                    
                    ws.Cells(Ticker, 11).Value = Format(0, "Percent")
                    
                    End If
                    
                'Calculate and write total Volume of Stock in column L
                ws.Cells(Ticker, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                
                'Increase Ticker Count by 1
                Ticker = Ticker + 1
                
                'Set new start row of the ticker block
                j = i + 1
                
                End If
            
            Next i
            
        'Find last non-blank cell in column I
        LastRowColI = ws.Cells(Rows.Count, 9).End(xlUp).Row
        'MsgBox ("Last row in column I is " & LastRowI)
        
        'Prepare for summary
        GreatVolume = ws.Cells(2, 12).Value
        GreatIncrease = ws.Cells(2, 11).Value
        GreatDecrease = ws.Cells(2, 11).Value
        
            'Loop for summary
            For i = 2 To LastRowColI
            
                'For greatest total volume--check if next value is larger--if yes take over a new value and populate ws.Cells
                If ws.Cells(i, 12).Value > GreatVolume Then
                GreatVolume = ws.Cells(i, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatVolume = GreatVolume
                
                End If
                
                'For greatest increase--check if next value is larger--if yes take over a new value and populate ws.Cells
                If ws.Cells(i, 11).Value > GreatIncrease Then
                GreatIncrease = ws.Cells(i, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatIncrease = GreatIncrease
                
                End If
                
                'For greatest decrease--check if next value is smaller--if yes take over a new value and populate ws.Cells
                If ws.Cells(i, 11).Value < GreatDecrease Then
                GreatDecrease = ws.Cells(i, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatDecrease = GreatDecrease
                
                End If
                
            'Write summary results in ws.Cells
            ws.Cells(2, 17).Value = Format(GreatIncrease, "Percent")
            ws.Cells(3, 17).Value = Format(GreatDecrease, "Percent")
            ws.Cells(4, 17).Value = Format(GreatVolume, "Scientific")
            
            Next i
            
        'ADjust column width automatically
        Worksheets(WorksheetExcel).Columns("A:Z").AutoFit
            
    Next ws
        
End Sub


