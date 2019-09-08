# Ranges, Worksheets and Workbooks

## Select a range

```visual basic
Range("A2")
Range("A2,A4") 'A2, A4
Range("A2","A4") 'A2, A3, A4
```

## Select a worksheet

```visual basic
Sheet1.Range("A1").etc 'Code name of sheet (can be changed in properties)
Worksheets("Blabla").Range("A1").etc 'by name given by user (if user changes name does not work anymore!)
Worksheets(6).Range("A1").etc 'Use the 6th worksheet (if order is changed does not work anymore!)
```

## Select a workbook

```visual basic
ThisWorkBook
Workbooks("YourWorkBookName.xlsx")
```

