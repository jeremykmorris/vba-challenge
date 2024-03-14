Attribute VB_Name = "Module2"
'Start with the code for one sheet
Sub tickerloop()
     'Set variables
     Dim tickername As String
     Dim tickervolume As Double
        tickervolume = 0
  
     Dim summary_ticker_row As Integer
        summary_ticker_row = 2
        
     Dim open_price As Double
        open_price = Cells(2, 3).Value
     Dim close_price As Double
     Dim yearly_change As Double
     Dim percent_change As Double

     'Label Summary Table headers
     Cells(1, 9).Value = "Ticker"
     Cells(1, 10).Value = "Yearly Change"
     Cells(1, 11).Value = "Percent Change"
     Cells(1, 12).Value = "Total Stock Volume"

     'Find last row
     lastRow = Cells(Rows.Count, 1).End(xlUp).Row

     'Loop through  data by the ticker name until last row
     For i = 2 To lastRow

         If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
     
           'Update tickername and tickervolume
           tickername = Cells(i, 1).Value
           tickervolume = tickervolume + Cells(i, 7).Value

           'Print summary table
           Range("I" & summary_ticker_row).Value = tickername
           Range("L" & summary_ticker_row).Value = tickervolume

           'closing price
           close_price = Cells(i, 6).Value
           
           'yearly change
            yearly_change = (close_price - open_price)
           
           Range("J" & summary_ticker_row).Value = yearly_change

             If open_price = 0 Then
                 percent_change = 0
             Else
                 percent_change = yearly_change / open_price
             End If

           'Print yearly change for each ticker in summary table
           Range("K" & summary_ticker_row).Value = percent_change
           Range("K" & summary_ticker_row).NumberFormat = "0.00%"

           'Reset
           summary_ticker_row = summary_ticker_row + 1

           tickervolume = 0
           
           open_price = Cells(i + 1, 3)
         
         Else
           
            'Add trade volume
           tickervolume = tickervolume + Cells(i, 7).Value

         
         End If
     
     Next i

 'Conditional formatting for yearly change in summary table

 lastrow_summary_table = Cells(Rows.Count, 9).End(xlUp).Row
 
 'Color code yearly change
     For i = 2 To lastrow_summary_table
         If Cells(i, 10).Value > 0 Then
             Cells(i, 10).Interior.Color = RGB(0, 255, 0)
         Else
             Cells(i, 10).Interior.Color = RGB(255, 0, 0)
         End If
     Next i

 'changes in stock price max min
     Cells(2, 15).Value = "Greatest % Increase"
     Cells(3, 15).Value = "Greatest % Decrease"
     Cells(4, 15).Value = "Greatest Total Volume"
     Cells(1, 16).Value = "Ticker"
     Cells(1, 17).Value = "Value"

    'summary table

     For i = 2 To lastrow_summary_table
         'max percent change
         If Cells(i, 11).Value = Application.WorksheetFunction.Max(Range("K2:K" & lastrow_summary_table)) Then
             Cells(2, 16).Value = Cells(i, 9).Value
             Cells(2, 17).Value = Cells(i, 11).Value
             Cells(2, 17).NumberFormat = "0.00%"

         'min percent change
         ElseIf Cells(i, 11).Value = Application.WorksheetFunction.Min(Range("K2:K" & lastrow_summary_table)) Then
             Cells(3, 16).Value = Cells(i, 9).Value
             Cells(3, 17).Value = Cells(i, 11).Value
             Cells(3, 17).NumberFormat = "0.00%"
         
         'max trade volume
         ElseIf Cells(i, 12).Value = Application.WorksheetFunction.Max(Range("L2:L" & lastrow_summary_table)) Then
             Cells(4, 16).Value = Cells(i, 9).Value
             Cells(4, 17).Value = Cells(i, 12).Value
         
         End If
     
     Next i
     
End Sub
'run on each worksheet
Sub FormatWorksheets()
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        ' Your formatting code goes here
        ' For example, you can add code to format each worksheet
        ' This is where you would put the code that you want to run on each worksheet
        
        ' For demonstration purposes, let's change the background color of cell A1 on each worksheet
        ws.Range("A1").Interior.Color = RGB(255, 192, 203) ' Light pink color
    Next ws

End Sub
