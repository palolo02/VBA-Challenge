' ################################################
' Paolo Vega
' VBA Script for the tickers in Walls Streets
' Bootcamp
' Assignment 2 - VBA
' Challenge VBA Stocks
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
        setWorksheetHeaders ws
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
    
    
    
    ' Iterate through the whole Worksheets again to get the greatest numbers among the tickers
    For Each ws In ActiveWorkbook.Worksheets
		
		' Initialize variables to start couting all over again in every sheet
        greatest_inc = 0
        greatest_dec = 0
        greatest_val = 0
        results_idx = 2
        ticker = ""
        
		' Iterate through the whole rows after being summarized
        For i = 2 To CInt(ws.Cells(Rows.Count, "I").End(xlUp).Row)
			' Set the current ticker we're analyzing
            ticker = ws.Cells(i, "I").Value
            
			' Compare the values to determine the highest percentage against the previous ticker
            If (ws.Cells(i, "K").Value > greatest_inc) Then
				' Update the greatest %
                greatest_inc = ws.Cells(i, "K").Value
				' Update the information in the cell and apply formatting to facilite how to read it
                ws.Cells(2, 18).Value = ws.Cells(i, "K").Value
                ws.Cells(2, 18).HorizontalAlignment = xlCenter
                ws.Cells(2, 18).NumberFormat = "#,#00.00%"
                ws.Cells(2, 18).Value = greatest_inc
                ws.Cells(2, 17).Value = ticker
            End If
            ' Compare the values to determine the lowest percentage against the previous ticker
            If (ws.Cells(i, "K").Value < greatest_dec) Then
				' Update the lowest %
                greatest_dec = ws.Cells(i, "K").Value
				' Update the information in the cell and apply formatting to facilite how to read it
                ws.Cells(3, 18).Value = ws.Cells(i, "K").Value
                ws.Cells(3, 18).HorizontalAlignment = xlCenter
                ws.Cells(3, 18).NumberFormat = "#,#00.00%"
                ws.Cells(3, 18).Value = greatest_dec
                ws.Cells(3, 17).Value = ticker
            End If
            ' Compare the values to determine the lowest stock value against the previous ticker
            If (ws.Cells(i, "L").Value > greatest_val) Then
				' Update the greatest %
                greatest_val = ws.Cells(i, "L").Value
				' Update the information in the cell and apply formatting to facilite how to read it
                ws.Cells(4, 18).Value = ws.Cells(i, "L").Value
                ws.Cells(4, 18).NumberFormat = "#,#00"
                ws.Cells(4, 18).HorizontalAlignment = xlCenter
                ws.Cells(4, 18).Value = greatest_val
                ws.Cells(4, 17).Value = ticker
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


' Set the labels and titles for the challenge per Worksheet
Sub setHeaders(ws As Worksheet)
    
	ws.Cells(1, 17).Value = "Ticker"
	ws.Cells(1, 18).Value = "Value"	
    ws.Cells(2, 16).Value = "Greatest % Increase"
    ws.Cells(3, 16).Value = "Greatest % Decrease"
    ws.Cells(4, 16).Value = "Greatest Total Volume"
    
End Sub







