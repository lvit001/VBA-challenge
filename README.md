# VBA-challenge
Challenge 2 Assignment for Bootcamp - Sources for Code

Found Variable "LongLong" here https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/data-type-summary


Found .NumberFormat = "0.00%" here https://stackoverflow.com/questions/42844778/vba-for-each-cell-in-range-format-as-percentage

  Applied this same logic when changing the format to currency


Found Application.worksheetfunction.max(range("a:a")) here https://stackoverflow.com/questions/42633273/finding-max-of-a-column-in-vba


For looping through worksheets code, that was found here under option 2 here: https://support.microsoft.com/en-us/topic/macro-to-loop-through-all-worksheets-in-a-workbook-feef14e3-97cf-00e2-538b-5da40186e2b0

    'Define WS as a variable'
    Dim WS As Worksheet

        'Cycle through the worksheets'
         For Each WS In ActiveWorkbook.Worksheets
    
        'Use the macro below in each worksheet'
        Call StockLoop(WS)
        
    Next
    
End Sub


Also used this source for help with looping through worksheets and using Call and With functions https://stackoverflow.com/questions/21918166/excel-vba-for-each-worksheet-loop

Sub StockLoop(WS As Worksheet)

With WS

Adding . before lines of code to reference the WS

End With
