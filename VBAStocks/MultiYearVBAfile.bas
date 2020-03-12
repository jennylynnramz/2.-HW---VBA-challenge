Attribute VB_Name = "Module1"
Sub alphatest():

'Running on multiple sheets

Dim xsheet As Worksheet
For Each xsheet In ThisWorkbook.Worksheets
 xsheet.Select




'Find Last Row
       Dim lastrow As Long
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        
'Add column title "Ticker", define column values
        Cells(1, 9).Value = "Ticker"
        Dim ticker As String

'Add column title "Price Change", define column values
        Cells(1, 10).Value = "Price Change"
        Dim price_change As Double
   
'Add column title "Percent Change"
        Cells(1, 11).Value = "Percent Change"
        Dim percent_change As Double
        Dim open_price As Double
        Dim close_price As Double
        
'Add column title "Total Volume"
        Cells(1, 12).Value = "Total Volume"
        Dim total_volume As Long
        

        
'add tickers. dynamic array
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        price_change = 0
        close_price = 0
        open_price = 0
        percent_change = 0
        stock_volume = 0
        
'Time to do all the complex maths
        For i = 2 To lastrow
            
            If Cells(i + 1, 1).Value = Cells(i, 1).Value Then
                    tickers = Cells(i, 1).Value
                    ticker_value = Cells(i, 1).Value
                If Cells(i + 1, 1).Value = Cells(i, 1).Value Then open_price = open_price + Cells(i, 3).Value 'works
                If Cells(i + 1, 1).Value = Cells(i, 1).Value Then close_price = close_price + Cells(i, 6).Value 'works
                If Cells(i + 1, 1).Value = Cells(i, 1).Value Then price_change = (close_price - open_price) 'works
                If Cells(i + 1, 1).Value = Cells(i, 1).Value Then percent_change = ((close_price - open_price) / open_price) 'works. but needs percent sign still
                If Cells(i + 1, 1).Value = Cells(i, 1).Value Then stock_volume = stock_volume + Cells(i, 7).Value 'works
            
             ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then Range("I" & Summary_Table_Row).Value = tickers 'works
                Range("J" & Summary_Table_Row).Value = price_change 'worked
                Range("K" & Summary_Table_Row).Value = FormatPercent(percent_change)
                Range("L" & Summary_Table_Row).Value = stock_volume 'worked
               
            
                Summary_Table_Row = Summary_Table_Row + 1 'works
               
            
             End If
             
         Next i
            
'Reset the variables
            price_change = 0
            close_price = 0
            open_price = 0
            percent_change = 0
            stock_volume = 0
           
'Add conditional formatting, positive green, negative red
            
            For f = 2 To lastrow
                If Cells(f, 10) < 0 Then Cells(f, 10).Interior.ColorIndex = 3
                    If Cells(f, 10) > 0 Then Cells(f, 10).Interior.ColorIndex = 4
             'zero left intentionally unformatted
            Next f
                
'Create additional table
   
        'Create Titles for Data
        Cells(2, 14).Value = "Greatest % Increase"
        Cells(3, 14).Value = "Greatest % Decrease"
        Cells(4, 14).Value = "Greatest Total Volume"
        Cells(1, 15).Value = "Ticker"
        Cells(1, 16).Value = "Value"
        
'Define values for challenge table

'Greatest Increase AS PERCENT FORMATgreatest_increase = Application.WorksheetFunction.Max(Range("k:k"))

        greatest_increase = Application.WorksheetFunction.Max(Range("k:k"))
       For i = 2 To lastrow
          If Cells(i, 11).Value = greatest_increase Then Cells(2, 16).Value = FormatPercent(greatest_increase)
           If Cells(i, 11).Value = greatest_increase Then Cells(2, 15).Value = Cells(i, 9).Value
        
                
            Next i
            

'Greates decrease AS PERCENT FORMAT
        greatest_decrease = Application.WorksheetFunction.Min(Range("k:k"))
       For j = 2 To lastrow
          If Cells(j, 11).Value = greatest_decrease Then Cells(3, 16).Value = FormatPercent(greatest_decrease)
           If Cells(j, 11).Value = greatest_decrease Then Cells(3, 15).Value = Cells(j, 9).Value
        
                
            Next j
        
        
'Greatest total colume as LONG
        
        greatest_volume = Application.WorksheetFunction.Max(Range("L:L"))
       For k = 2 To lastrow
        If Cells(k, 12).Value = greatest_volume Then Cells(4, 16).Value = greatest_volume
           If Cells(k, 12).Value = greatest_volume Then Cells(4, 15).Value = Cells(k, 9).Value
       
       
            Next k
        
        
        
        
 'formatting
        'Need to make these cells autofit the content.works
        Columns("A:Q").AutoFit
        'Bold format the first row bc I like the way it looks.works
        Range("A1:Q1").Font.Bold = True
        Range("N2:N4").Font.Bold = True
        'More formatting just to make myself happy. works
        Range("A1:G1").Interior.ColorIndex = 48
        Range("I1:L1").Interior.ColorIndex = 48
        Range("N1:P1").Interior.ColorIndex = 48
        Range("N2:N4").Interior.ColorIndex = 48
        
        
        
        
                
                
                
                
            

'To do:
'Challenges:
    

'Adjust to run on every worksheet
'Look into formatting to ignore blank cells in grey fill

    

        
  
       
        
Next xsheet



    
End Sub









