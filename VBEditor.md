# VBEditor

## Open editor

Alt + F11

## Access documentation

* Click on “Object Browser” (or `F2`)

* Click on problematic word and then `F1` to get help
* Go to the “Object Browser” AND click on a method/property and press `F1` for help

:exclamation: You can add libraries in Tools > References...

## Syntax check

Tools > Options > Untick if you do not want systematic check

## Auto-completion

Ctrl + Space

## Comment out

Click “Comment block” and “Uncomment block” in the editing functions

## Debug

### Execute row by row

F8, F8, F8, etc.

### Show result in immediate window (console)

```visual basic
Debug.Print ActiveWorkbook.Path
```

## Subs

```vb
Sub my_sub()
    Code.code.code _ ' SPACE + _ to go to next line
    code.continuation
    ' Comment
End Sub
```

