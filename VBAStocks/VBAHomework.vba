' ################################################
' Paolo Vega
' VBA Script for the tickers in Walls Streets
' Bootcamp
' Assignment 2 - VBA
' VBA Stocks
' Version 1.0 - 02/29/2020
' Version 1.1 - 02/29/2020
'#################################################


' ================ Main routine =====================
Sub TotalStocks()
    
    Dim ws As Worksheet
	
	' Variables to store the data for each Ticker
    Dim ticker As String	
    Dim stock_vol As Double	
    Dim initial_year_ind As Double
    Dim final_year_ind As Double
	
	' Index to print the results for each sheet
    Dim results_idx As Double
	
	' Variables for the challenge
    Dim greatest_inc As Double
    Dim greatest_dec As Double
    Dim greatest_val As Double
    
    
    ' Iterate through the whole Worksheets to get inidividual differences over year
    For Each ws In ActiveWorkbook.Worksheets
        
        ' Initialize variables to start couting all over again
        stock_vol = 0
        initial_year_ind = 0
        final_year_ind = 0
        results_idx = 2
        ticker = ""
        
		' Set the titles for the rersults
        setHeaders ws
                     
        ' Iterate through the whole rows to read the ticker's values
        For i = 2 To ws.UsedRange.Rows.Count
		
            'Read the current ticker
            ticker = ws.Cells(i, 1).Value
			
			' Add its stock value
            stock_vol = ws.Cells(i, 7).Value + stock_vol
			
            'Compare the following ticker to check to see if it is the same or a new one
            If (ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value) Then
				
				' Store information of that ticker for future references
                ticker = ws.Cells(i, 1).Value
				' If the next row is a new ticker, we're reading the last value of the year
                final_year_ind = ws.Cells(i, 6).Value
                
                ' Print Results for the unique ticker
                ws.Cells(results_idx, 9).Value = ticker
				' Apply some formatting
                ws.Cells(results_idx, 9).HorizontalAlignment = xlCenter
                
				' Calculate the yearly change for that ticker
                ws.Cells(results_idx, 10).Value = final_year_ind - initial_year_ind
                
				' Assess if the division is possible to calculate the % yearly change
                If (initial_year_ind <> 0) Then
                    ws.Cells(results_idx, 11).Value = (final_year_ind / initial_year_ind) - 1
                Else
                    ws.Cells(results_idx, 11).Value = 0
                End If
                
				' Apply formatting to read easily the variations in the yearly change
                If (ws.Cells(results_idx, 10) >= 0) Then
                    ws.Cells(results_idx, 10).Interior.ColorIndex = 10 ' Green
                ElseIf ws.Cells(results_idx, 10) < 0 Then
                    ws.Cells(results_idx, 10).Interior.ColorIndex = 30 ' Red
                End If
                
				' Set white font color to contrast against green or red
                ws.Cells(results_idx, 10).Font.Color = vbWhite
                
				' Set the accumulative stock in the appropiate location for the ticker
                ws.Cells(results_idx, 12).Value = stock_vol
				' Set the initial value for the ticker (validation purposes)
                'ws.Cells(results_idx, 13).Value = initial_year_ind
				' Set the final value for the ticker (validation purposes)
                'ws.Cells(results_idx, 14).Value = final_year_ind
                
                ' Enhance the formatting of the numbers
                ws.Cells(results_idx, 10).NumberFormat = "#,#00.00"
                ws.Cells(results_idx, 10).HorizontalAlignment = xlCenter
                ws.Cells(results_idx, 11).NumberFormat = "#,#00.00%"
                ws.Cells(results_idx, 11).HorizontalAlignment = xlCenter
                ws.Cells(results_idx, 12).NumberFormat = "#,#00"
                ws.Cells(results_idx, 12).HorizontalAlignment = xlCenter
                
				' Initialize variables to count it for the next ticker
                stock_vol = 0
                initial_year_ind = 0
                final_year_ind = 0
				' Go below the record to continue with the results
                results_idx = results_idx + 1
				
            ' Assess if this is the first row of the ticker to store its initial value
            ElseIf (initial_year_ind = 0) Then
                initial_year_ind = ws.Cells(i, 3).Value
            End If
		
        Next i
        
    Next ws
    
End Sub

' Set the titles for the final results per Ticker
Sub setWorksheetHeaders(ws As Worksheet)
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "% Yearly Change"
    ws.Cells(1, 12).Value = "Total Stock"
    ws.Cells(1, 13).Value = "Open"
    ws.Cells(1, 14).Value = "Close"
End Sub








