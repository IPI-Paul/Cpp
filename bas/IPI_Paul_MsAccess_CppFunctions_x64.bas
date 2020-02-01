Attribute VB_Name = "IPI_Paul_MsAccess_CppFunctions"
Private Declare PtrSafe Function multiplyBy Lib "C:\Users\Paul\Documents\Source Files\dll\multiplyBy x64.dll" (ByRef x As Double, ByRef y As Double) As Double

Function cppMultiplyBy(ByRef x As Double, ByRef y As Double)
    Dim result
    multiplyBy x, y
    cppMultiplyBy = y
End Function
