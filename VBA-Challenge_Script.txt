Sub WorksheetLoop()

'==================================================='
'CODE TO LOOP THROUGH THE WORKSHEETS IN THE WORKBOOK'
'==================================================='

    'Define WS as a variable'
    Dim WS As Worksheet

        'Cycle through the worksheets'
         For Each WS In ActiveWorkbook.Worksheets
    
        'Use the macro below in each worksheet'
        Call StockLoop(WS)
        
    Next
    
End Sub


Sub StockLoop(WS As Worksheet)

With WS
'================================================================================'
'CREATE COLUMNS FOR TICKER, YEARLY CHANGE, PERCENT CHANGE, AND TOTAL STOCK VOLUME'
'================================================================================'
    
    .Range("I1").Value = "Ticker"
    .Range("J1").Value = "Yearly Change"
    .Range("K1").Value = "Percent Change"
    .Range("L1").Value = "Total Stock Volume"
    

'============================================================================================'
'CREATE ROWS AND COLUMNS FOR GREATEST % INCREASE, % DECREASE, TOTAL VOLUME, TICKER, AND VALUE'
'============================================================================================'
        
        'Create the row labels first'
        .Range("O2").Value = "Greatest % Increase"
        .Range("O3").Value = "Greatest % Decrease"
        .Range("O4").Value = "Greaest Total Volume"
        
        'Create the Column Labels'
        .Range("P1").Value = "Ticker"
        .Range("Q1").Value = "Value"
        
        
'======================================================================================'
'     CREATE LOOP TO GET TICKERS, YEARLY CHANGE, PERCENT CHANGE, AND TOTAL STOCK VOLUME'
'======================================================================================'

    'Create a Variable to represent the second row where we will begin inputting data'
    Dim DataTableRow As Integer
    DataTableRow = 2
    
    'Create a Variables'
    Dim Ticker As String
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalStock As LongLong
    Dim SheetName As String

        SheetName = .Name
    
    'Total Stock will start at 0 for each ticker'
    TotalStock = 0
        

    'Determine what the last row of the data is'
    Dim LR As Long
    LR = .Cells(Rows.Count, 1).End(xlUp).Row
    
    'Create Variables for Year and Start Date'
    Dim Year As String
    Dim StartDate As String
    
        'Assign the name of the worksheet to the year variable'
        Year = .Name
        'MsgBox (Year)
        
        'Assign the year plus the first day of the fiscal year into the StartDate variable'
        StartDate = Year & "0102"
        'MsgBox (StartDate)
        
    
    'Create a vertical loop to find ticker, opening price, closing price, and total stock volume'
    For I = 2 To LR
    
        'Create an if statement to find the opening price value based on the start date'
        If .Cells(I, 2).Value = StartDate Then
        'Need to figure out a way to just look for 0102 not the whole thing as the year changes'
        
            'Set the OpeningPrice value'
            OpeningPrice = .Cells(I, 3).Value
            'MsgBox ("Opening Price for" & " " & Ticker & ":" & " " & OpeningPrice)
            
        End If
    
        'If a cell value in column A does not equal the next cell'
        If .Cells(I, 1).Value <> .Cells(I + 1, 1).Value Then
            
            'Set the ticker value'
            Ticker = .Cells(I, 1).Value
            
            'Set ClosingPrice Value'
            ClosingPrice = .Cells(I, 6).Value
            'MsgBox ("Closing Price for" & " " & Ticker & ":" & " " & ClosingPrice)
            
            'Set the YearlyChange Value'
            YearlyChange = ClosingPrice - OpeningPrice
            
            'Set the PercentChange Value'
            PercentChange = ((ClosingPrice - OpeningPrice) / OpeningPrice)
            
            'Add to the TotalStock value'
            TotalStock = TotalStock + .Cells(I, 7).Value
            
            'Add Ticker Value into the data table'
            .Range("I" & DataTableRow).Value = Ticker
            
            'Add the yearlychange value to data table'
            .Range("J" & DataTableRow).Value = YearlyChange
            
                'Change the values in the yearlchange column to currency format'
                .Range("J" & DataTableRow).NumberFormat = "$0.00"
            
            'If statement to apply conditional formatting to the yearly change column'
            If .Range("J" & DataTableRow).Value >= 0 Then
            
                'Set the cell color of positive values as green'
                .Range("J" & DataTableRow).Interior.ColorIndex = 4
                
            Else
                'set the cell color of negative values as red'
                .Range("J" & DataTableRow).Interior.ColorIndex = 3
                    
            End If
                              
            
            'Add the PercentChange Vaue into the data table'
            .Range("K" & DataTableRow).Value = PercentChange
            
                'Change the format to be a percentage'
                .Range("K" & DataTableRow).NumberFormat = "0.00%"
                
                
            'If statement to apply conditional formatting to the yearly change column'
            If .Range("K" & DataTableRow).Value >= 0 Then
            
                'Set the cell color of positive values as green'
                .Range("K" & DataTableRow).Interior.ColorIndex = 4
                
            Else
                'set the cell color of negative values as red'
                .Range("K" & DataTableRow).Interior.ColorIndex = 3
                    
            End If
                
            'Add the TotalStock value into the data table'
            .Range("L" & DataTableRow).Value = TotalStock

            'Go to the next row of the data table'
            DataTableRow = DataTableRow + 1
            
            'Reset the TotalStock Value for the next ticker'
            TotalStock = 0
            
        'If the cells next to each other have the same ticker'
        Else
            
            'Add the TotalStock'
            TotalStock = TotalStock + .Cells(I, 7).Value
            
        End If
        
    Next I
                
'======================================================================================================='
'CREATE LOOP TO FIND TICKERS AND VALUES FOR GREATEST % INCREASE, % DECREASE, AND TOTAL VOLUME'
'======================================================================================================='
            
        'Create a Second LR for the second dataset'
        Dim LR2 As Long
            LR2 = .Cells(Rows.Count, 10).End(xlUp).Row
            
        'Find the Greatest % Increase Value and change the format to a percentage'
        .Range("Q2").Value = Application.WorksheetFunction.Max(.Range("K2" & ":" & "K" & LR2).Value)
        .Range("Q2").NumberFormat = "0.00%"
              
        'Find the Greatest % Decrease Value and change the format to a percentage'
        .Range("Q3").Value = Application.WorksheetFunction.Min(.Range("K2" & ":" & "K" & LR2).Value)
        .Range("Q3").NumberFormat = "0.00%"
        
        'Find the Greatest Total Volume'
        .Range("Q4").Value = Application.WorksheetFunction.Max(.Range("L2" & ":" & "L" & LR2).Value)
           
           
        'Create a nested for loop'
        For m = 2 To LR2
        
            
            'Create an if statement to find the ticker that goes along with Greatest % increase'
            If .Cells(m, 11).Value = .Range("Q2").Value Then
                
                'Assigning the ticker value'
                .Range("P2").Value = .Cells(m, 9).Value
                
        
            End If
            
            'Create an if statement to find the ticker that goes along with Greatest % deccrease'
            If .Cells(m, 11).Value = .Range("Q3").Value Then
            
                'Assigning the ticker value'
                .Range("P3").Value = .Cells(m, 9).Value
                
            End If
            
            
            'Create an if statement to find the ticker that goes along with Greatest Total volume'
            If .Cells(m, 12).Value = .Range("Q4").Value Then
            
                'Assigning the ticker value'
                .Range("P4").Value = .Cells(m, 9).Value
                
            End If
            
        Next m
        
    End With
        
End Sub