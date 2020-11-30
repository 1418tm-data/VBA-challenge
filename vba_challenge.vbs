Sub stock_data():
    '-------------------------------------------------------------------------------------
    'Start loop through worksheets
    Dim ws As Worksheet
    Dim starting_ws As Worksheet
    Set starting_ws = ActiveSheet           'Remember which worksheet is active at the start

    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
    '--------------------------------------------------------------------------------------
    
    'This code runs on every worksheet
        Dim ticker_name As String               'Declare variable to store tickers
        
        Dim stock_Vol_Sum As Double             'Declare variable to store Total Stock volume and initialise it to 0
        stock_Vol_Sum = 0
        
        Dim ticker As Integer                   'Declare variable to store Total Stock volume and initialise it to 0
        ticker = 2
        
        Dim lastrow As Long                     'Declare last row of the dataset and set the active sheet to the
        Set wksheet = ActiveSheet               'current worksheet
        lastrow = wksheet.Cells(wksheet.Rows.Count, "A").End(xlUp).Row
        
        Dim open_price As Double                'Declare and initialise variabbles to store starting price, closing price
        open_price = -5                         'year change and % change
        Dim close_price As Double
        Dim year_change As Double
        year_change = 0
        Dim year_perc_change As Double
            '----------------------------------------------
            'Insert the values for the summary table
            Range("P2").Value = "Greatest % Increase"
            Range("P3").Value = "Greatest % Decrease"
            Range("P4").Value = "Greatest Total Volume"
            Range("Q1").Value = "Ticker"
            Range("R1").Value = "Value"
            '-----------------------------------------------
            'Set headers for the output columns
            Range("J1").Value = "Ticker"
            Range("K1").Value = "Yearly Change"
            Range("L1").Value = "Percentage change"
            Range("M1").Value = "Total Stock Volume"
            
    For i = 2 To lastrow
        'loop through the first column and check if the next cell different to the current
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            'if true updte the variables...
            ticker_name = Cells(i, 1).Value
            stock_Vol_Sum = stock_Vol_Sum + Cells(i, 7).Value
            close_price = Cells(i, 6).Value
            year_change = close_price - open_price
                
                'check for 0 values in the opening/ closing price columns
                If year_change = 0 Or open_price = 0 Then
                    year_perc_change = 0
                Else
                    year_perc_change = year_change / open_price
                End If
                
                    'Conditional formating of the year change column
                    If year_change < 0 Then
                        Cells(ticker, 11).Interior.ColorIndex = 3
                    Else
                        Cells(ticker, 11).Interior.ColorIndex = 4
                    End If
                    
            'Update the columns with the summary data stored in the variables
            Cells(ticker, 10).Value = ticker_name
            Cells(ticker, 11).Value = year_change
            Cells(ticker, 12).Value = year_perc_change
            Cells(ticker, 12).NumberFormat = "0.00%"
            Cells(ticker, 13).Value = stock_Vol_Sum
        
            'Reset the variables ready for the next iteration
            ticker = ticker + 1 '(this one however increases by 1)
            stock_Vol_Sum = 0
            year_change = 0
            open_price = -5
            
        'if the current sell and next are the same then...
        Else
                If open_price = -5 Then
                    open_price = Cells(i, 3).Value
                End If
            'be sure to include totals of the last row
            stock_Vol_Sum = stock_Vol_Sum + Cells(i, 7).Value
        End If
    Next i
'--------------------------------------------------------------------------------------------
'Call private subroutine to run the summary
    Call great_summary
Next

starting_ws.Activate    'Activate the worksheet that was originally active
End Sub

'Sub to create the summary
Private Sub great_summary():
    'decare variables to hold the summary data
    Dim Total_Vol As Double
    Dim ticker_p_inc, ticker_p_dec, ticker_total As String
    Dim percentage_inc, percentage_dec As Double
    
    Dim lastrow As Long                     'Declare last row of the dataset and set the active sheet to the
        Set wksheet = ActiveSheet           'current worksheet
        lastrow = wksheet.Cells(wksheet.Rows.Count, "L").End(xlUp).Row
    
    For i = 2 To lastrow
        'loop through the new summary columns to find the min and max values
        percentage_inc = Application.WorksheetFunction.Max(Range("l:l"))
        percentage_dec = Application.WorksheetFunction.Min(Range("l:l"))
        Total_Vol = Application.WorksheetFunction.Max(Range("m:m"))
            'if the max or min is found, place the ticker corresponding in the variables
            If Cells(i, 12) = percentage_inc Then
                ticker_p_inc = Cells(i, 10).Value
            End If
            
            If Cells(i, 12) = percentage_dec Then
                ticker_p_dec = Cells(i, 10).Value
            End If
            
            If Cells(i, 13) = Total_Vol Then
                ticker_total = Cells(i, 10).Value
            End If
        'populate the cells with the data in the variables
        Range("Q2").Value = ticker_p_inc
        Range("Q3").Value = ticker_p_dec
        Range("Q4").Value = ticker_total
        Range("R2").Value = percentage_inc
        Range("R2").NumberFormat = "0.00%"
        Range("R3").Value = percentage_dec
        Range("R3").NumberFormat = "0.00%"
        Range("R4").Value = Total_Vol
    Next i
End Sub