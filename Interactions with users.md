# Interactions with users

## Message boxes

```visual basic
Sub welcome()

    MsgBox "Welcome " & Excel.Application.UserName & "!", , "Welcome"

End Sub
```

:exclamation: No parenthesis!!

## Message boxes with multiple choices

```visual basic
Sub clear_yes_no()

    Dim Answer As Byte
    Answer = MsgBox("Are you sure?", vbYesNo + vbQuestion, "Please confirm.")
    
    If Answer = vbYes Then
        Range("A7").CurrentRegion.Clear
    End If

End Sub
```

:exclamation: Parenthesis obligatory!!

## Input boxes

### VBA input boxes

VBA input boxes always return a string.

```vb
Dim Resp As String
Dim lastRow As Long

Resp = VBA.InputBox("Type the customer's name:")

If Resp <> "" Then
    lastRow = Sheet5.Range("A" & Sheet5.Rows.Count).End(xlUp).Row + 1
    Sheet5.Range("A" & lastRow).Value = Excel.WorksheetFunction.Proper(Resp)
End If
```

### Excel input boxes

```vb
Dim Resp As Long
Dim lastRow As Long

Resp = Application.InputBox("Type an amount:", "Amount", , , , , , 1) 'Type:=1 is for numbers

If Not IsEmpty(Resp) Then
    lastRow = Sheet6.Range("A" & Sheet6.Rows.Count).End(xlUp).Row + 1
    Sheet6.Range("A" & lastRow).Value = Resp
End If
```

Possible types:

* 0: Formula

* 1: Number
* 2: String
* 8:Range
* 1+2=3: Number + String

:exclamation: Excel performs validation of data type