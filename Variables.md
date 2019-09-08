# Variables

## Types

* Byte: 0 to 255
* Boolean: True or False
* Integer: -32k to 32k
* Long: -2B to 2B
* Double: Very large floats
* String: Strings
* Variant: Anything

## Declare variables

```visual basic
Dim myText As String '1 variable
Dim myText As String, myLong As Long '2 variables at once
Dim myArray(1 To 10) As String '1-dim array
Dim myArray(1 To 10, 1 To 2) As String '2-dim array

Const myConstant as String = "Nadia" 'Constant values

Dim myWS as WorkSheet
Set myWS = WorkSheets("Bla") 'Object variables need to be SET
```

Depending on where the variable is declared it will be available to use:

* In the **procedure** if declared in the the procedure
* In the **module** if declared outside the procedure
* In **all modules and procedures** if declared outside a procedure with `Public` instead of `Dim`