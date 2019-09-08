# Object-oriented VBA

## Objects

The following hierarchy applies:

* Application.Workbooks (“Name”).Worksheets(“ WSName”).Range(“A1”)

:exclamation: When not specified, the object is ASSUMED by Excel to be the active object

## Methods

```vb
Range("A2").Copy Range("B3") 'Destination is an optional argument to Copy
Sheet1.Copy After:=Sheet2
'OR
Sheet1.Copy ,Sheet2 'comma to signify to interpret as second argument
```

## Properties

```vb
Range("A1").Interior.Color = vbRed 'Assign a value to a property
Range("A1").Interior 'A property can return an object (here an interior object)
```

