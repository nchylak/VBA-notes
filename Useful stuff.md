# Useful stuff

## Make code faster

The below makes the code faster by: 

* Disabling screen updating
* Preventing interim calculations
* Preventing warning pop up windows
* Emptying the clipboard at the end of the macro.

```vb
Private Sub Entry_Point()

    With Application
        .StatusBar = "Running..."
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .DisplayAlerts = False   
    End With
    
End Sub

Private Sub Exit_Point()

    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .DisplayAlerts = True
        .StatusBar = ""
        'in case you have used the copy and pastespecial methods, you could have a lot of data on the clipboard.
        .Application.CutCopyMode = False
    End With

End Sub

Sub Do_Stuff()
    
    Call Entry_Point

	'CODE
    
	Call Exit_Point

End Sub
```

## Find last row

```visual basic
' If no blank cells (like CTRL + up/down)
Range("K6").Value = Range("A" & Rows.Count).End(xlUp).Row
' Find row with last used cells
Range("K11").Value = Cells.SpecialCells(xlCellTypeLastCell).Row
' Find the nb of rows of the sheet's used range
Range("K12").Value = ActiveSheet.UsedRange.Rows.Count
```

## Copy filtered table to a new worksheet

```visual basic
ActiveSheet.AutoFilter.Range.Copy
Worksheets.Add
Range("A1").PasteSpecial
```

