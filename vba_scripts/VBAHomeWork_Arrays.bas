Attribute VB_Name = "VBAHomeWork_Arrays"
Option Explicit

Sub VBA_HomeWork_Arrays()

    'declaring variables
    Dim ws As Worksheet
    Dim r As Range
    Dim rangeOfCells As Range
    Dim headerRange As Range
    Dim i As Long
    Dim j As Long
    Dim counter As Long
    Dim unique As Boolean
    'declaring arrays
    Dim uArray() As Variant
    Dim headers() As Variant
        
    headers() = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")
    
    'initiating For Each loop
    For Each ws In Worksheets
    
    'setting the initial sizing of the uArray and it's 0 value
    ReDim uArray(0)
    uArray(UBound(uArray)) = ws.Range("A2").Value
    
    'setting the headers and range of cells considering the active worksheet
    'as well as setting the uArray UBound counter to 0 every ws loop
    Set headerRange = ws.Range("I1:L1")
    Set rangeOfCells = ws.Range("A2", ws.Range("A2").End(xlDown))
    headerRange.Value = headers()
    counter = 0
                
        For Each r In rangeOfCells
        
            unique = True
            'For Loop to check if r.Value is found within the uArray.
            'if found, exits loop and moves towards next r.Value
            For i = LBound(uArray) To UBound(uArray)
                If r.Value = uArray(i) Then
                    unique = False
                    Exit For
                End If
            Next i
            
            'if r.Value is not found within the array,
            'increase uArray size by 1 and assigns the r.Value
            If unique = True Then
                counter = counter + 1
                ReDim Preserve uArray(counter)
                uArray(UBound(uArray)) = r.Value
            End If
                      
        Next r
        
        'extracting uArray values into the worksheet
        For i = LBound(uArray) To UBound(uArray)
            ws.Range("I2").Offset(i, 0).Value = uArray(i)
        Next i
        
        'clearing uArray from memory to restart the next ws loop
        Erase uArray
        
        '-----------------Part Two!!!--------------------

        'declaring new variables
        Dim numberOfTickers As Long
        'declare one Array for each column (yearly change, percent change and total volume)
        Dim oArray() As Variant
        Dim cArray() As Variant
        Dim vArray() As Variant
        
        ReDim oArray(0)
        ReDim cArray(0)
        ReDim vArray(0)
        
        numberOfTickers = ws.Range("I2", ws.Range("I2").End(xlDown)).Count
        
        'modify rangeOfCells variable for this step
        Set rangeOfCells = ws.Range("A2", ws.Range("A2").End(xlDown).Offset(1, 0))
               
        For i = 0 To numberOfTickers - 1
                                 
            counter = 0
                                 
            'For Loop to check if r.Value matches the value in the Ticker column (range I)
            'if found, adds values to each array and moves towards next row of r.Value
            For Each r In rangeOfCells
                unique = False
                If r.Value = ws.Range("I" & i + 2).Value Then
                    unique = True
                    counter = counter + 1
                    ReDim Preserve oArray(counter)
                    ReDim Preserve cArray(counter)
                    ReDim Preserve vArray(counter)
                    oArray(UBound(oArray)) = r.Offset(0, 2).Value
                    cArray(UBound(cArray)) = r.Offset(0, 5).Value
                    vArray(UBound(vArray)) = r.Offset(0, 6).Value
                End If
                
                'if r.Value does not match the value in the Ticker column (range I)
                'make calculations and print out the values in the Arrays
                'after values are printed, clear Arrays from memory and restart the loop to next Ticker value
                If unique = False And counter <> 0 Then
                
                    'checks if 1st oArray value is 0, if so loops until it's not 0 and assigns that value as oArrayMin
                    Dim oArrayMin As Double
                    oArrayMin = 0
                    
                        For j = 1 To UBound(oArray)
                            If oArray(j) <> 0 Then
                                oArrayMin = oArray(j)
                                Exit For
                            End If
                        Next j
                        
                    ws.Range("I" & i + 2).Offset(0, 1).Value = cArray(UBound(cArray)) - oArrayMin
                    
                    '---checks if division is 0/0, if it is then output 0 as result
                    If oArrayMin = 0 And oArray(UBound(oArray)) = 0 Then
                        ws.Range("I" & i + 2).Offset(0, 2).Value = 0
                    Else
                        ws.Range("I" & i + 2).Offset(0, 2).Value = (cArray(UBound(cArray)) / oArrayMin) - 1
                    End If
                    '---
                    ws.Range("I" & i + 2).Offset(0, 3).Value = Application.WorksheetFunction.Sum(vArray)
                    Erase oArray, cArray, vArray
                    Exit For
                End If
            
            Next r
            
        Next i
        
        'applying conditional formatting (red/green) to "Yearly Change" column
        For i = 0 To numberOfTickers - 1
            If ws.Range("J2").Offset(i, 0).Value < 0 Then
                ws.Range("J2").Offset(i, 0).Interior.Color = rgbRed
                ws.Range("J2").Offset(i, 0).Font.Color = rgbWhite
            Else
                ws.Range("J2").Offset(i, 0).Interior.Color = rgbGreen
                ws.Range("J2").Offset(i, 0).Font.Color = rgbWhite
            End If
        Next i
        
        'applying percentage number format to "Percent Change" column
        ws.Range("K2", ws.Range("K2").End(xlDown)).NumberFormat = "0.00%"
               
               
        '-----------------Bonus Solution!--------------------
        
        'declaring new variables
        Dim gIncrease As Double
        Dim gDecrease As Double
        Dim gVolume As Double
    
        gIncrease = Application.WorksheetFunction.Max(ws.Range("K2", ws.Range("K2").End(xlDown)))
        gDecrease = Application.WorksheetFunction.Min(ws.Range("K2", ws.Range("K2").End(xlDown)))
        gVolume = Application.WorksheetFunction.Max(ws.Range("L2", ws.Range("L2").End(xlDown)))
        
        'print out headers in the workshet
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"
        ws.Range("N2").Value = "Greatest % increase"
        ws.Range("N3").Value = "Greatest % decrease"
        ws.Range("N4").Value = "Greatest total volume"
        
        'look for MAX/MIN value and print it out
        ws.Range("P2").Value = gIncrease
        ws.Range("P3").Value = gDecrease
        ws.Range("P4").Value = gVolume
            
        'looking up the values
        ws.Range("O2").Value = Application.WorksheetFunction.Index(ws.Range("I2", ws.Range("I2").End(xlDown)), _
                            Application.WorksheetFunction.Match(gIncrease, ws.Range("K2", ws.Range("K2").End(xlDown)), 0))
        ws.Range("O3").Value = Application.WorksheetFunction.Index(ws.Range("I2", ws.Range("I2").End(xlDown)), _
                            Application.WorksheetFunction.Match(gDecrease, ws.Range("K2", ws.Range("K2").End(xlDown)), 0))
        ws.Range("O4").Value = Application.WorksheetFunction.Index(ws.Range("I2", ws.Range("I2").End(xlDown)), _
                            Application.WorksheetFunction.Match(gVolume, ws.Range("L2", ws.Range("L2").End(xlDown)), 0))
    
        'applying percentage number format greatest % cells
        ws.Range("P2", "P3").NumberFormat = "0.00%"
                 
    Next ws
           
End Sub
