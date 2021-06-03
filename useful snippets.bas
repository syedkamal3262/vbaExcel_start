Attribute VB_Name = "Module8"
Option Explicit

Sub Example1()

ActiveCell.Value = "abc"
End Sub

Sub Example2()
ActiveCell.Copy Destination:=Range("A2:A7")
End Sub

Sub MethodExample()

Worksheets.Add before:=ActiveSheet
End Sub


Sub RangeExample()
Worksheets("sheet1").Range("A1:B9").Value = "22"
Worksheets(1).Range("C2,C7").Value = "33"
ActiveSheet.Range("D4").Value = "99"

End Sub

Sub activeCellExample()
Range("A1", "A4").Select
Range(ActiveCell, "K2").Select
End Sub

Sub offsetExample()
ActiveCell.Offset(0, 1).Range("A1:C1").Value = "ABC"
End Sub

Sub variablesExample()
Dim MyNumber As Long
MyNumber = Worksheets(1).Rows.Count
ActiveCell.Value = MyNumber
End Sub

Sub IFstatement()
    Dim num As Integer
    num = 30
    If num = 20 Then
        MsgBox ("bumber is 20")
    ElseIf num = 30 Then
        MsgBox ("bumber is 30")
    Else
        MsgBox ("number is not 20")
    End If
End Sub

Sub CaseTest()
 Dim myNum As Integer
 myNum = ActiveCell.Value
 Select Case myNum
    Case 10 To 20
        MsgBox "number is betwenn 10 to 20"
    Case 30
        MsgBox "num is 30"
    Case Else
        MsgBox "NUM IS NOT IN BETWEEN 10 TO 30 AND IS ALSO NOT 30"
End Select
End Sub

Sub formatCell()

With ActiveCell.Font
    .Bold = True
    .Color = vbBlue
    .Italic = True
    .Name = "Arial"
End With
End Sub


Sub StringTest()
Dim myString As String
myString = Range("A2").Value
Range("A2").Offset(0, 1).Value = LCase(myString)
Range("A2").Offset(0, 2).Value = UCase(myString)
Range("A2").Offset(0, 3).Value = Len(myString)
Range("A2").Offset(0, 4).Value = Trim(myString)
Range("A2").Offset(0, 5).Value = Left(myString, 2)
Range("A2").Offset(0, 6).Value = Right(myString, 2)
Range("A2").Offset(0, 7).Value = Mid(myString, 2, 5)
End Sub

Sub forNextloop()
Dim startNumber As Integer
For startNumber = 1 To 5
    MsgBox (startNumber)
    Next startNumber
    MsgBox "loop end"
End Sub
Sub FORCELLS()
Dim startNum As Integer, endNum As Integer
endNum = ActiveSheet.UsedRange.Rows.Count - 1
For startNum = 1 To endNum
    Cells(startNum + 1, "A").Value = startNum
Next startNum

End Sub
Sub forEachloop()
Dim wrksht As Worksheet

For Each wrksht In Worksheets
    MsgBox wrksht.Name
Next wrksht

End Sub

Sub doountitilloopcheck()
    Range("D2").Select
    Do Until ActiveCell = ""
        Active.Value = ActiveCell.Value * 2
        'satarted 259.
        ActiveCell.Offset(1, 0).Select
    Loop
End Sub

Sub doWhileloopcheck()
    Range("D2").Select
    Do While ActiveCell <> ""
        Active.Value = ActiveCell.Value * 2
        'satarted 259.
        ActiveCell.Offset(1, 0).Select
    Loop
End Sub


Sub testSubProcedures()
If ActiveCell.Value > 20 Then
    Call formatCell ' CALLING ABOVE ALREADY ASUB PROCEDURE
End If
End Sub


Sub testExitSub()

If IsNumeric(ActiveCell.Value) Then
    MsgBox "number is numeric"
    Exit Sub
    Else
    Call formatCell ' CALLING ABOVE ALREADY ASUB PROCEDURE
    MsgBox "number is not numeric"
End If
End Sub



Sub testVByesNoBox()

If MsgBox("do you like vba ?", vbYesNo + vbQuestion, "hello there ? ") = vbYes Then
    MsgBox "thats good!"
    Else
    MsgBox "thats bad "
End If
End Sub


Sub testInputBOX()
Dim department As Variant
department = InputBox("enter department", "required ?", "HR")
MsgBox (department)
End Sub


Function ABmultiply(A As Integer, B As Integer) As Integer
    ABmultiply = A + B
End Function

Sub usefunc()
    ActiveCell = ABmultiply(Range("A1"), Range("A2"))
End Sub


Sub dimObject()

Dim rangearea As Range
Set rangearea = Range("A1:B7")
rangearea.Select

End Sub
