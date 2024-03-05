Attribute VB_Name = "Module1"
Sub tickerinfo()

'Begin the worksheet loop, with all cells and ranges needing ws.
    For Each ws In Worksheets

'Set variables for: ticker, annual opening and closing price, yearly change, % change, volume
    Dim Ticker As String
    Dim Open_Annual As Double
    Dim Close_Annual As Double
    Dim Yearly_Change As Double
    Dim Volume_Total As Double
    Dim Percent_Change As Double
 
'Set starting position for stock's opening price for the year
    'Cells(2, 3).Value = Open_Annual
    Open_Annual = ws.Cells(2, 3).Value

'Set an initial variable for holding the stock volume per ticker
    Volume_Total = 0

'Keep track of the location for each ticker in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2

'Variable for counting the number of rows
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
 
'Loop through each ticker each day ***LIMITED LOOP RANGE FOR INITIAL TESTING
    For I = 2 To lastrow

'Check if still the same ticker, if it is NOT <> ...
    If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
    
'Store the value of the Ticker, Annual Open and Annual Close
            Ticker = ws.Cells(I, 1).Value

'Store the value of the Annual Close, since this is the last row for the ticker per if<>
            Close_Annual = ws.Cells(I, 6).Value
            
'Before overwriting the ticker's annual opening price...
'Calculate yearly and % changes, print and format
            Yearly_Change = (Close_Annual - Open_Annual)
            ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
            
'To address possible divide by zero error when annual opening price is 0
            'If Open_Annual = 0 Then
                ws.Range("K" & Summary_Table_Row).Value = "Cannot calculate with annual opening price of zero"
            'Else
            'Calculate % Change
                Percent_Change = (Yearly_Change / Open_Annual)
            'Print % Change
                ws.Range("K" & Summary_Table_Row).Value = Percent_Change
            'Format for percentage data type
                ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
            'Format color for yearly change
                If Yearly_Change >= 0 Then
                    ws.Range("J" & Summary_Table_Row).Interior.Color = vbGreen
                    
                Else
                    ws.Range("J" & Summary_Table_Row).Interior.Color = vbRed
                
                End If
                
            'Close nested if to catch div/0 errors
            'End If
                                    
'Add to the stock volume
            Volume_Total = Volume_Total + ws.Cells(I, 7).Value
    
      'Print the Ticker in the summary table
            ws.Range("I" & Summary_Table_Row).Value = Ticker

      ' Print the stock volume to the summary table
      ws.Range("L" & Summary_Table_Row).Value = Volume_Total

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the stock volume total
      Volume_Total = 0
      
      'Store the next ticker Annual Open for when the loop iterates again through if<>
        Open_Annual = ws.Cells(I + 1, 3)

    ' If the cell immediately following a row is the same ticker, add to the stock volume total
    Else
      Volume_Total = Volume_Total + ws.Cells(I, 7).Value
    
    'Close the if conditional
    End If
    
    'Close the for loop
    Next I

    'Add the table column headers, etc.
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
'Define variables for add'l table
        Dim increase_greatest As Double
        Dim decrease_greatest As Double
        Dim volume_greatest As Double

        'increase_greatest = 0
        
 'new loop for add'l table
        'For j = 2 To lastrow
    
    'find and return max increase w/format
            increase_greatest = WorksheetFunction.max(ws.Range("K2:K5000"))
            ws.Range("Q2").Value = increase_greatest
        
        '**need to return corresponding tickers
        
        'find and return min increase w/format
            decrease_greatest = WorksheetFunction.Min(ws.Range("K2:K5000"))
            'ws.Range("P3").Value = ws.Range("K" & Summary_Table_Row).Value
            ws.Range("Q3").Value = decrease_greatest
            
            'format percentages
            ws.Range("Q2:Q3").NumberFormat = "0.00%"
            
        'find and return max volume
             volume_greatest = WorksheetFunction.max(ws.Range("L2:L5000"))
             ws.Range("Q4").Value = volume_greatest
    
    'Print the Ticker in the summary table
       '     ws.Range("I" & Summary_Table_Row).Value = Ticker

      ' Print the stock volume to the summary table
      'ws.Range("L" & q).Value = Volume_Total
    
    'Volume_Total
    'Percent_Change
        
        'Next j
        
'Close the For looping across worksheets
    Next ws

'Confirming number of wkshts is recognized, trying to debug why no loop across the wkbk...
'Debug.Print (Worksheets.Count)

End Sub


Sub testclear()

For Each ws In Worksheets

'Clear the table results to help while testing
    ws.Range("I1:L1000").clear

Next ws

End Sub
