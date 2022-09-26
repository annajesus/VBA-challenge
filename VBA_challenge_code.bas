Attribute VB_Name = "Module1"
Option Explicit

Sub Alphabetical_testing()

'Declarations
        Dim wsheet As Worksheet
        Dim ticker As String
        Dim Total_vol As Double
        Dim yr_open As Double
        Dim yr_close As Double
        Dim yr_change As Double
        Dim percent_change As Double
        Dim summary_tbl As Integer
        Dim rng As Range
        Dim j As Long
        Dim k As Long
        Dim cell_color As Range
        
        
        
'Create a loop to set up all column headings through each worksheet.
        For Each wsheet In Worksheets
            'input headers for all worksheets
            wsheet.Range("I1").Value = "Ticker"
            wsheet.Range("J1").Value = "Yearly Change"
            wsheet.Range("K1").Value = "Percent Change"
            wsheet.Range("L1").Value = "Total Stock Volume"


'Perform a loop of wsheets to the last row to determine opening price for 1st ticker and closing price for last ticker

'Determine last row for loop
        Dim lst_row As Long
            lst_row = Cells(Rows.Count, 1).End(xlUp).Row


'Set up integer and value for loop
summary_tbl = 2
Total_vol = 0

Dim i As Long
Dim ticker_row As Long
ticker_row = 1

      For i = 2 To lst_row
      
      If wsheet.Cells(i + 1, 1).Value <> wsheet.Cells(i, 1).Value Then
            'find all the values of the tickers
            ticker = wsheet.Cells(i, 1).Value
            Total_vol = wsheet.Cells(i, 7).Value
            'define year opening price
            yr_open = wsheet.Cells(i, 3).Value
            'Define year closing price
            yr_close = wsheet.Cells(i, 6).Value
            'calculate changes in prices
            yr_change = yr_open - yr_close
            percent_change = (1 - (yr_close / yr_open)) * 100
            
            
            'Insert values into summary_tbl
            wsheet.Cells(summary_tbl, 9).Value = ticker
            wsheet.Cells(summary_tbl, 10).Value = yr_change
            wsheet.Cells(summary_tbl, 11).Value = percent_change
            wsheet.Cells(summary_tbl, 12).Value = Total_vol
            summary_tbl = summary_tbl + 1
        
        End If

Next i


'conditional formatting for pos_change_green and neg_change_red


        
        'define which range to find the change in
        Set rng = Range("J2", Range("J2").End(xlDown))
        k = rng.Cells.Count
        For j = 1 To k
        Set cell_color = rng(j)
        Select Case cell_color
            Case Is >= 0
            With cell_color
                .Interior.ColorIndex = 4
                End With
            Case Is < 0
            With cell_color
                .Interior.ColorIndex = 3
                
                End With
        
        End Select
        
Next j


'loop through the following worksheets
Next wsheet

End Sub

