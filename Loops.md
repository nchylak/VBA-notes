# Loops

## With... End With

```visual basic
Sub exo_with()

    Dim myRange As Range
    Set myRange = Range("A10", "A" & Cells(Rows.Count, 1).End(xlUp).Row)

    With myRange.Font
        .Name = "Calibri"
        .Size = 12
        .Bold = True
    End With

End Sub
```

## For Each `x` in `y`.... Next `x`

```visual basic
Sub protect()

	Dim WS As Worksheet
    
    For Each WS In ThisWorkbook.Worksheets    
        WS.protect       
    Next WS

End Sub
```

## For `i` = 0 To 10 Step `1`... Next `i`

```visual basic
Sub simple_for()

    Dim i As Long
    Dim lastRow As Long
    Dim cellValue As Double
    Const startRow As Integer = 10
    
    lastRow = Range("A" & Rows.Count).End(xlUp).Row

    For i = startRow To lastRow
        cellValue = Range("F" & i).Value
        If cellValue > 400 Then
            Range("F" & i).Value = cellValue + 10
        End If      
    Next i

End Sub
```

:exclamation: `Exit For` if you want to leave earlier

## If... Then... ElseIf... Then....Else...End If

```visual basic
Sub long_if()

    Dim WS As Worksheet
    
    For Each WS In ThisWorkbook.Worksheets    
        If WS.Name = "Purpose" Then
            WS.protect        
        ElseIf WS.CodeName = "Sheet1" Then
            'do nothing        
        Else
            WS.protect , , , , , True, True, True        
        End If    
    Next WS

End Sub
```

## Select Case... Case 1...Case 2... Case Else... End Select

```visual basic
Sub exo_case()

    Select Case Range("B3").Value    
        Case 1 To 200
            Range("C3").Value = "Good"        
        Case 0
            Range("C3").Value = ""        
        Case Is > 200
            Range("C3").Value = "Excellent"            
        Case Else
            Range("C3").Value = "Bad"        
    End Select

End Sub
```

## Do While...Loop / Do Until... Loop / Do... Loop

:exclamation: `Exit Do` for leaving early

```vb
Sub do_until()

    Const startRow As Integer = 8
    Dim nextRow As Long
    
    nextRow = startRow
    
    Do Until Sheet3.Range("A" & nextRow).Value = ""
        Sheet3.Range("B" & nextRow).Value = Sheet3.Range("A" & nextRow).Value + 10
        nextRow = nextRow + 1
    Loop

End Sub
'-------------------------------------------------------------------------------
Sub do_while()

    Const startRow As Integer = 8
    Dim nextRow As Long
    
    nextRow = startRow
    
    Do While Sheet3.Range("A" & nextRow).Value <> ""
        Sheet3.Range("C" & nextRow).Value = Sheet3.Range("A" & nextRow).Value + 10
        nextRow = nextRow + 1
    Loop

End Sub
'-------------------------------------------------------------------------------
Sub do_loop()

    Const startRow As Integer = 8
    Dim nextRow As Long
    
    nextRow = startRow
    
    Do
        If Sheet3.Range("A" & nextRow).Value = "" Then Exit Do
        Sheet3.Range("D" & nextRow).Value = Sheet3.Range("A" & nextRow).Value + 10
        nextRow = nextRow + 1
    Loop

End Sub
```

## GoTo `blah` .... `blah`: (for error handling)

```visual basic
Sub exo_goto()

    Range("D3").Value = ""

    If VBA.IsError(Range("B3").Value) Then GoTo Leave
    
    'if no error
    Range("C3").Value = Range("B3").Value
    Exit Sub
            
'if error       
Leave:
    Range("C3").Value = ""
    Range("D3").Value = "ERROR"

End Sub
```

:exclamation: Jumps through the code so need early exit if jump is not taken (`Exit Sub`)

## Find method

```vb
Dim cellID As Range

Set cellID = Sheet4.Range("A:A").Find(Sheet4.Range("B3"), , xlValues, xlWhole) 'Like pressing CTRL+F
Set cellID = Sheet4.Range("A:A").FindNext(cellID) ' To find the the next occurence. Warning, it loops again from the beginning once it is done so a variable needs to be introduced in order to track when we are back at the beginning (like pressing CTRL+F endlessly)
```

