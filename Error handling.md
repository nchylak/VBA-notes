# Error handling

## Step into

Debug > Step Into or F8, F8, F8, ...

## Breakpoints

Debug > Toggle Breakpoints or F9 then F8, F8, F8, ...

## Immediate window

View > Immediate Window

The immediate window prints out the result of `Debug.Print` . You can also type code directly in this window using the `?`.

## Watch window

View > Watch Window

The watch window is for keeping an eye on variables. Variables can be added to the watch window with right click + add to watch or by drag-and-dropping.

Specify the property you would like to keep an eye on, e.g. `myVariable.value`.

## Hover over variables

To  see the value currently contained by the variable.

## GoTo

```vb
On Error GoTo errorHandle
    
'Blah

Exit Sub ' For not going through errorHandle
    
errorHandle:
    Select Case Err.Number
        Case 424 'Cancel is selected
            Exit Sub
        Case Else
            MsgBox "An error occurred."
    End Select
```

