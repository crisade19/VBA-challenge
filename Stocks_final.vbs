Attribute VB_Name = "Stocks"
Sub Stocks()

Dim ticker As String
Dim prev_ticker As String
Dim LastRowA As Long
Dim LastRowM As Long
Dim j As Long
Dim x As Long
Dim min As Long
Dim max As Long
Dim open_val As Double
Dim close_val As Double
Dim pop_total As Double

'Empty initialisation
prev_ticker = ""

'Incremental row for generated table
j = 2

'Incremental min
min = 2

'Check last non empty row
LastRowA = Range("A" & Rows.Count).End(xlUp).Row

For i = 2 To LastRowA
    ticker = Cells(i, 1).Value
    
    If prev_ticker = "" Then
        'Initialise the first ticker name
        prev_ticker = ticker
    End If
    
    'If the previous ticker is different from the new one do...
    If prev_ticker <> ticker Or i = LastRowA Then
    
        'When we get to last row max should be equal to i ELSE max should be i - 1
        If i = LastRow Then
            max = i
        Else
            max = i - 1
        End If
        
        'Populate the first ticker in the new table
        Cells(j, 11).Value = prev_ticker
        
        'Calculate the Total Stock Volume
        pop_total = WorksheetFunction.Sum(Range("G" & min & ":G" & max))
        'Populate the Total Stock Volume value into the new table
        Cells(j, 14).Value = pop_total
        
        'Get the beginning of the year open value + end of the year close value
        open_val = Cells(min, 3).Value
        close_val = Cells(max, 6).Value
        
        'Find out the Yearly Change
        If open_val > close_val Then
            Cells(j, 12).Value = close_val - open_val
            Cells(j, 12).Interior.ColorIndex = 3
        ElseIf open_val < close_val Then
            Cells(j, 12).Value = close_val - open_val
            Cells(j, 12).Interior.ColorIndex = 4
        End If
        
        'Find out the Percentage Change
        Cells(j, 13).Formula = "=" & close_val & "/" & open_val & "-1"
        Cells(j, 13).NumberFormat = "0.00%"
        
        'Move to the next row inside the new table
        j = j + 1
        'Change the new min to the old max + 1
        min = max + 1
        'Change the previous ticker to new ticker
        prev_ticker = ticker
          
    End If
Next


'Check last non empty row
LastRowM = Range("M" & Rows.Count).End(xlUp).Row

'Calculate percentage increase, decrease and greatest total vol
increase = WorksheetFunction.max(Range("M2 : M" & LastRowM))
decrease = WorksheetFunction.min(Range("M2 : M" & LastRowM))
total_vol = WorksheetFunction.max(Range("N2 : N" & LastRowM))

'Format cells value to percentage
Range("R2:R3").NumberFormat = "0.00%"

'Set new table values
Cells(2, 18).Value = increase
Cells(3, 18).Value = decrease
Cells(4, 18).Value = total_vol

For x = 2 To LastRowM

    'Look for ticker name which value matches increase var
    If Cells(x, 13).Value = increase Then
        Cells(2, 17).Value = Cells(x, 11).Value
    End If
    
    'Look for ticker name which value matches decrease var
    If Cells(x, 13).Value = decrease Then
        Cells(3, 17).Value = Cells(x, 11).Value
    End If
    
    'Look for ticker name which value matches total_vol var
    If Cells(x, 14).Value = total_vol Then
        Cells(4, 17).Value = Cells(x, 11).Value
    End If
    
Next


    
End Sub

