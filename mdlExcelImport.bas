Attribute VB_Name = "mdlExcelImport"
' Given a list of tasks imported from Excel,
' where the tasks are leading-space delimited to indicate
' the outline level, set the task outline level and strip
' the excess spaces.
Public Sub ConvertIndentedTasksToMSProjectOutline()
    Dim t As Task
    For Each t In ThisProject.Tasks
        If Not (t Is Nothing) Then
            If (t.OutlineLevel = 1) Then
                t.OutlineLevel = get_outline_level(t.Name, 4)
                t.Name = strip_leading_spaces(t.Name)
            End If
        End If
    Next
End Sub


Private Function strip_leading_spaces(s As String) As String
    If (s = "") Then
        strip_leading_spaces = ""
        Exit Function
    End If
    
    Dim n As Integer
    Dim found As Integer
    found = 0
    For n = 1 To Len(s)
        If (Mid(s, n, 1) <> " ") Then
            found = n
            Exit For
        End If
    Next n
    strip_leading_spaces = Mid(s, found)
End Function

Private Function get_outline_level(s As String, n As Integer) As Integer
    Dim tmp As String
    tmp = s
    Dim ret As Integer
    ret = 1
    Do While (Left(tmp, n) = String(n, " "))
        ret = ret + 1
        tmp = Mid(tmp, n + 1)
    Loop
    get_outline_level = ret
End Function


' ///////////////////////////
' TESTS

Private Sub run_tests()
    assert_outline_level_equals "", 1
    assert_outline_level_equals "a", 1
    assert_outline_level_equals " a", 1
    assert_outline_level_equals "  a", 1
    assert_outline_level_equals "   a", 1
    assert_outline_level_equals "    b", 2
    assert_outline_level_equals "       b2", 2
    assert_outline_level_equals "        c3", 3
    assert_outline_level_equals "           d3", 3
    assert_outline_level_equals String(4 * 5, " ") + "a", 6
    
    assert_strip_equals "", ""
    assert_strip_equals "a", "a"
    assert_strip_equals "a  ", "a  "
    assert_strip_equals "a b", "a b"
    assert_strip_equals "    x", "x"
    assert_strip_equals String(1000, " ") & "hello", "hello"
    Debug.Print "Done"
End Sub

Private Function assert_strip_equals(s As String, expected As String)
    Dim actual As String
    actual = strip_leading_spaces(s)
    If (actual <> expected) Then
        Dim msg As String
        msg = "For " & s & ", expected " & expected & " but got " & actual
        Err.Raise vbObjectError, , msg
    End If
End Function

Private Function assert_outline_level_equals(s As String, expected As Integer)
    Dim actual As Integer
    actual = get_outline_level(s, 4)
    If (actual <> expected) Then
        Dim msg As String
        msg = "For " & s & ", expected " & expected & " but got " & actual
        Err.Raise vbObjectError, , msg
    End If
End Function


