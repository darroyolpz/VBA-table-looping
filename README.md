# Table looping

I've been coding VBA for quite some time in my company. So far our files only had about 3000 rows top, so there was no concern about performance nor speed. Waiting time was not a struggle.

However, I decided to test some code to check by myself how to get some improvements.

## For-loop approach

I've been using this structure for many months. It was easy to code, fast and familiar for me. Sadly it's not the best way to loop through a 1M rows since it takes about 6 seconds to complete.

```vba
x = 7644667220

LastRow = Sheets(1).Cells(Cells.Rows.Count, 1).End(xlUp).Row
For i = 2 To LastRow
    Value = Sheets(1).Cells(i, 1).Value
    If Value = x Then
        Debug.Print Sheets(1).Cells(i, 3)
        Exit For
    End If
Next i
```
## Application-loop approach

This method is far superior to the previous one. It takes advantage of the pre-built Excel functions, so run time is extremely fast. It takes only 0.03 seconds to run (about 200 times faster)

```vba
Set tData = Range("table").ListObject

If Not IsError(Application.Match(x, tData.ListColumns("Column1").Range, 0)) Then
    index_value = Application.Match(x, tData.ListColumns("Column1").Range, 0)
    cell_value = Application.Index(tData.ListColumns("Column3").Range, index_value, 1)
    Debug.Print cell_value
Else
    Debug.Print "Value does not exist"
End If
```

This uses one of the most common functions in Excel, index & match, to get values from a table.

There is no doubt I will be using this application approach in my next codes. One does not get 200 times faster very often.
