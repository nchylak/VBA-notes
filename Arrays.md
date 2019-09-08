# Arrays

## One-dimensional arrays

### Fill arrays

```vb
Dim MonthArray(1 To 12) As String
Dim i As Integer

' Fill an array

For i = LBound(MonthArray) To UBound(MonthArray) 'i = 1 to 12
	MonthArray(i) = Range("A" & i + 4).Value
Next i

For i = LBound(MonthArray) To UBound(MonthArray)
	MonthArray(i) = Range("Months").Cells(i, 1).Value 'With named ranges
Next i
```

Or faster with variant arrays:

```vb
Dim MonthArray As Variant

MonthArray = Range("Months").Value
```

### Write arrays

```vb
Range("C5:N5").Value = MonthArray 'horizontally
Range("C5:C16").Value = Application.WorksheetFunction.Transpose(MonthArray) 'vertically (take the transpose)
```

Or faster with variant arrays:

```vb
Range("Months") = MonthArray 
```

### Dynamic arrays

```vb
Dim MonthArray() As String 'No dimension given
Dim UBarray As Long 'Upper bound is dynamic

UBarray = Range("A3").CurrentRegion.Rows.Count - 1
ReDim MonthArray(1 To UBarray) 'Redimension of array
```

:exclamation: `ReDim` erases the content of the array!! In order to keep the content, use `Redim Preserve`:

```vb
ReDim Preserve MonthArray(1 To UBarray + 2)
```

## Two-dimensional arrays

### Fill arrays

```vb
Dim MonthArray(1 To 12, 1 To 2) As Variant
Dim i As Integer, j As Integer

For i = LBound(MonthArray) To UBound(MonthArray)

    For j = 1 To 2
    	MonthArray(i, j) = Sheet4.Cells(i + 4, j).Value
    Next j

Next i

For i = LBound(MonthArray) To UBound(MonthArray)

    For j = 1 To 2
    	MonthArray(i, j) = Range("myRange").Cells(i, j).Value 'With named ranges
    Next j

Next i
```

