
- Workbook is `CleaningUpData.xlsm` and Sheet is `RECORDING`
- I recorded a macro `AddHeader`. I added new row using Ctrl+ "+". Then added name to each column and stopped the macro.
- we recorded another macro 'FormatHeaders'. We changed the followings:
	- Font Color
	- Background Color
	- Thick Borders
	- Bold the Headers
- Now we added the following code to apply above recorded procedures to all the sheets in the current workbook:

``` vba
Public Sub CleanUpData()

    Dim i As Integer
    i = 1
    Do While i <= Worksheets.Count
        Worksheets(i).Select
        AddHeader
        FormatHeaders
        i = i + 1


    Loop


End Sub
```
- 