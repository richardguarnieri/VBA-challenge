Attribute VB_Name = "VBAHomeWork_NonArrays"
Option Explicit

Sub VBA_HomeWork_NonArrays()

    'variables declaration
    Dim ticker_name As String
    Dim open_price As Double
    Dim close_price As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim total_volume As Double
    Dim last_row As Double
    Dim row_tracker As Double
    Dim ws As Worksheet
    Dim i As Long
    
    'bonus - variables declaration
    Dim greatest_increase As Double
    Dim greatest_increase_ticker As String
    Dim greatest_decrease As Double
    Dim greatest_decrease_ticker As String
    Dim greatest_volume As Double
    Dim greatest_volume_ticker As String
    
    For Each ws In Worksheets
    
        'setting up the headers
        ws.Range("I1:L1").Value = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"
        ws.Range("N2").Value = "Greatest % increase"
        ws.Range("N3").Value = "Greatest % decrease"
        ws.Range("N4").Value = "Greatest total volume"
        
        'resets variables every new ws
        last_row = ws.Range("A1", ws.Range("A1").End(xlDown)).Count
        open_price = ws.Range("C2").Value
        
        'bonus - resets variables every new ws
        greatest_increase = 0
        greatest_increase_ticker = ""
        greatest_decrease = 0
        greatest_decrease_ticker = ""
        greatest_volume = 0
        greatest_volume_ticker = ""
        
        For i = 1 To last_row
        
            'tracks the number of rows to output results
            row_tracker = ws.Range("I" & Rows.Count).End(xlUp).Offset(1, 0).Row
            
            'checks if open_price = 0, if it is then assigns next row value as open_price
            If open_price = 0 Then
                open_price = ws.Range("A1").Offset(i, 2).Value
            End If
            
            'checks if ticker following current row is different than current ticker
            'if not, add row volume to total_volume variable and
            'continue looping until tickers are different
            If ws.Range("A1").Offset(i + 1, 0) = ws.Range("A1").Offset(i, 0) Then
                total_volume = total_volume + ws.Range("A1").Offset(i, 6).Value
                
            'if tickers are different, add current row values to variables and make calculations
            Else
                ticker_name = ws.Range("A1").Offset(i, 0).Value
                close_price = ws.Range("A1").Offset(i, 5).Value
                total_volume = total_volume + ws.Range("A1").Offset(i, 6).Value
                yearly_change = close_price - open_price
                
                '---validates if a 0/0 division takes place, if so, output 0
                If open_price = 0 And close_price = 0 Then
                    percent_change = 0
                Else
                    percent_change = (close_price / open_price) - 1
                End If
                '---
                
                'output values into headers
                ws.Range("I" & row_tracker).Value = ticker_name
                ws.Range("J" & row_tracker).Value = yearly_change
                ws.Range("K" & row_tracker).Value = percent_change
                ws.Range("L" & row_tracker).Value = total_volume
                
                'bonus - variables assignments
                If percent_change > greatest_increase Then
                    greatest_increase = percent_change
                    greatest_increase_ticker = ticker_name
                ElseIf percent_change < greatest_decrease Then
                    greatest_decrease = percent_change
                    greatest_decrease_ticker = ticker_name
                End If
                If total_volume > greatest_volume Then
                    greatest_volume = total_volume
                    greatest_volume_ticker = ticker_name
                End If
                
                'resets variables
                open_price = ws.Range("A1").Offset(i + 1, 2).Value
                total_volume = 0
            End If
            
        Next i
        
        'bonus - variables outputs
        ws.Range("O2").Value = greatest_increase_ticker
        ws.Range("O3").Value = greatest_decrease_ticker
        ws.Range("O4").Value = greatest_volume_ticker
        ws.Range("P2").Value = greatest_increase
        ws.Range("P3").Value = greatest_decrease
        ws.Range("P4").Value = greatest_volume
        
        'bonus - applying conditional formatting
        ws.Range("P2").NumberFormat = "0.00%"
        ws.Range("P3").NumberFormat = "0.00%"
        
        'applying conditional formatting
        ws.Range("K2", ws.Range("K2").End(xlDown)).NumberFormat = "0.00%"
        
        For i = 1 To ws.Range("J2", ws.Range("J2").End(xlDown)).Count
            If ws.Range("J1").Offset(i, 0).Value < 0 Then
                ws.Range("J1").Offset(i, 0).Interior.Color = rgbRed
                ws.Range("J1").Offset(i, 0).Font.Color = rgbWhite
            Else
                ws.Range("J1").Offset(i, 0).Interior.Color = rgbGreen
                ws.Range("J1").Offset(i, 0).Font.Color = rgbWhite
            End If
        Next i
    
    Next ws
    
End Sub
