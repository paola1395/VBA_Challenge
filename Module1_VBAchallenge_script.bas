Attribute VB_Name = "Module1"
Sub alphabetical_testing()
    
    'Set variables
    Dim Ticker As String
    Dim ws As Worksheet
    Dim last As Long
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    summary_row = 2
    Total_volume = 0
    Open_val = Cells(2, "C").Value
    
    
    Cells(1, "i").Value = "Ticker"
    Cells(1, "j").Value = "Yearly Change"
    Cells(1, "k").Value = "Percent Change"
    Cells(1, "l").Value = "Total Stock Volume"
    
    'Loop through all stocks
    For i = 2 To lastRow
        Total_volume = Total_volume + Cells(i, "g").Value
    
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            Ticker = Cells(i, 1).Value
            
            close_val = Cells(i, "f").Value
            yearly_change = close_val - Open_val
            Percent_change = yearly_change / close_val
            
            
            Cells(summary_row, "i").Value = Ticker
            Cells(summary_row, "l").Value = Total_volume
            Cells(summary_row, "j").Value = yearly_change
            Cells(summary_row, "k").Value = Percent_change
            Cells(summary_row, "k").Style = "Percent"
            
            summary_row = summary_row + 1
            Total_volume = 0
            Open_val = Cells(i + 1, "C").Value
            
        End If
    
  
       'conditional formatting for yearly change column
        If Cells(summary_row, "j").Value > 0 Then
             Cells(summary_row, "j").Interior.ColorIndex = 4
        
        ElseIf Cells(summary_row, "j").Value < 0 Then
            Cells(summary_row, "j").Interior.ColorIndex = 3
        
        Else:
            Cells(summary_row, "j").Interior.ColorIndex = 0
        
        End If
            
    Next i


End Sub
