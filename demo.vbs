Const MY_CONSTANTS = "Some super const"
Dim AnotherVariable As String

Sub MyTestSub(ByVal arg1 As String)
    ' The below real-world used query contains nested quotes/apostrophes
    ' That could confuse `removeComments` routine
    Query = "SELECT * FROM __InstanceModificationEvent WITHIN 60 " _
    & "WHERE TargetInstance ISA 'Win32_PerfFormattedData_PerfOS_System' " _
    & "AND TargetInstance.SystemUpTime >= 200 AND " _
    & "TargetInstance.SystemUpTime < 320"

    Dim arr As Variant
    arr = Array(1, 2, 3, 4, 5, 6)   ' This is an inline comment

    Dim var1, var2
    var1 = "Testing Short string"
    var2 = "short"

    '
    ' This is a comment
    '
    Dim longString
    longString = "1. this is some very long string concatenated."
    longString = longString + "2. this is some very long string concatenated."
    longString = longString + "3. this is some very long string concatenated."
    longString = longString + "4. this is some very long string concatenated."

    longString = longString + "5. this is some very long string concatenated." _
    & "6. this is some very long string concatenated." _
    & "7. this is some very long string concatenated." _
    & "8. this is some very long string concatenated." _
    & "9. this is some very long string concatenated."

    Dim somecondition As Boolean
    somecondition = False
    If somecondition <> False Then
        Exit Sub
    End If

    MsgBox ("Test1(Constant): " & MY_CONSTANTS)
    MsgBox ("Test2(Query): " & Query)
    MsgBox ("Test3(var1 + var2): " & var1 & var2)
    MsgBox ("Test4(longString): " & longString)
    MsgBox ("Test5(Array's contents): " & Join(arr, ","))

End Sub

Sub test()
    MyTestSub (0)
End Sub
