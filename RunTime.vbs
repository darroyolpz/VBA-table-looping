Sub RunTime()

Dim StartTime As Double
Dim SecondsElapsed As Double
Dim tData As Object

StartTime = Timer
'*****************************
x = 7644667220#

LastRow = Sheets(1).Cells(Cells.Rows.Count, 1).End(xlUp).Row
For i = 2 To LastRow
    Value = Sheets(1).Cells(i, 1).Value
    If Value = x Then
        Debug.Print Sheets(1).Cells(i, 3)
        Exit For
    End If
Next i
'*****************************
ForElapsed = Round(Timer - StartTime, 2)

StartTime = Timer
'*****************************
Set tData = Range("table").ListObject

If Not IsError(Application.Match(x, tData.ListColumns("Column1").Range, 0)) Then
    index_value = Application.Match(x, tData.ListColumns("Column1").Range, 0)
    cell_value = Application.Index(tData.ListColumns("Column3").Range, index_value, 1)
    Debug.Print cell_value
Else
    Debug.Print "Value does not exist"
End If
'*****************************
AppElapsed = Round(Timer - StartTime, 2)

MsgBox "First For pass in " & ForElapsed & " seconds", vbInformation
MsgBox "Second App pass in " & AppElapsed & " seconds", vbInformation

End Sub