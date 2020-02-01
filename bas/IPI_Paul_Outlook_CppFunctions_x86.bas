Attribute VB_Name = "IPI_Paul_Outlook_CppFunctions"
Private Declare PtrSafe Function multiplyBy Lib "C:\Users\Paul\Documents\Source Files\dll\multiplyBy x86.dll" (ByRef x As Double, ByRef y As Double) As Double

Public Function cppMultiplyBy(ByRef x As Double, ByRef y As Double)
    Dim result
    multiplyBy x, y
    cppMultiplyBy = y
End Function

Function normVal(val)
    For Each itm In Array(",", "'")
        val = Replace(val, itm, "")
    Next
    normVal = val
End Function

Sub retCppMultiplyBy()
    Dim x As Double, y As Double
0:
    vals = InputBox("Please enter the two values to multiply separated by a space", "C++ multiplyBy Linked Function", "2.1 2,324.41")
    If vals > "" Then
        vals = Split(normVal(vals), " ")
        x = vals(0)
        y = vals(1)
        If MsgBox(x & " * " & vals(1) & " = " & cppMultiplyBy(x, y) & vbCrLf & vbCrLf & "Do you want to calculate another?", vbYesNo, "C++ multiplyBy Linked Function Result") = vbYes Then GoTo 0
    End If
End Sub

Sub viewForm_CppMultiplyBy()
    IPI_Paul_Outlook_CppMultiPlyBy.Show 0
End Sub
