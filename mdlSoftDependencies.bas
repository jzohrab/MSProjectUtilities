Attribute VB_Name = "mdlSoftDependencies"
' Methods to add "Soft dependencies" to an MS Project file.
' Based on <http://www.mpug.com/articles/hard-soft-dependencies/>.
'
' Outline
'
' This code assumes that the soft dependencies are stored in custom field *Text29*
' (reason: it's assumed that people may already be using Text1 and other Textn fields for some custom data).
' When the code runs, it clears the current soft dependencies from the hard dependencies,
' and re-adds the soft dependencies as hard dependencies.
'
' Note that the old soft dependencies are stored in custom field Text30.
' This is because the prior values are needed to completely break any changed soft dependencies.
' For example, if a user enters soft dependencies "5" for a field, and runs `AddSoftDependencies`,
' the value of "5" is entered as a (regular) hard dependency.  If the user changed that "5" to "14",
' the now-invalid hard dependency of "5" needs to be deleted as well, and then the hard dependency
' "14" is added.
'
' Configuration/set-up
' - add this module to MS Project file
' - fix areas in the "configuration" section of the code
' - add custom field *Text29* to the MS Project fields, rename to something like `Soft predecessor UniqueID(s)`
' - add field `UniqueID` to the field list as well, since the soft predecessors use that field, and not the row number.
'
' Usage
'
' - add UniqueID of the soft dependency predecessor(s) to the Text29 field
' - run `AddSoftPredecessors`


' ///////////////////////////////////////////////////////
' "Configuration" - edit the below as needed.

' Region-specific delimiter
Private Const DELIM = ","

' Get task's soft dependencies.
Private Function get_soft_dependencies(tsk As Task) As String
    get_soft_dependencies = tsk.Text29
End Function

' Get previous value of soft dependencies.
Private Function get_OLD_soft_deps(tsk As Task) As String
    get_OLD_soft_deps = tsk.Text30
End Function

' Set previous value of soft dependencies.
Private Sub set_OLD_soft_deps(tsk As Task, s As String)
    tsk.Text30 = s
End Sub

' ///////////////////////////////////////////////////////
' Public API
' Add/remove dependencies

' Removes the existing soft predecessors, and adds the existing ones.
Public Sub AddSoftPredecessors()
    On Error GoTo err_handler
    
    Dim errors As Collection
    Set errors = New Collection
    
    Call RemoveSoftPredecessors
    
    Dim tsk As Task
    For Each tsk In ActiveProject.Tasks
        If Not tsk Is Nothing Then
            Call do_add_soft_predecessors(tsk, errors)
            Call set_OLD_soft_deps(tsk, get_soft_dependencies(tsk))
        End If
    Next tsk
    
    Call report_errors(errors)
    Exit Sub
    
err_handler:
    MsgBox Err.Description, vbCritical, "Exception"
End Sub


Public Sub RemoveSoftPredecessors()
    On Error GoTo err_handler

    Dim tsk As Task
    For Each tsk In ActiveProject.Tasks
        If Not tsk Is Nothing Then
            tsk.UniqueIDPredecessors = remove_deps(tsk.UniqueIDPredecessors, get_soft_dependencies(tsk))
            tsk.UniqueIDPredecessors = remove_deps(tsk.UniqueIDPredecessors, get_OLD_soft_deps(tsk))
        End If
    Next tsk
    Exit Sub
    
err_handler:
    MsgBox Err.Description, vbCritical, "Exception"
End Sub


' /////////////////////////////////////////////////////////
' Utility methods

' Add soft predecessors, on failure add error message to errors collector.
Private Sub do_add_soft_predecessors(ByRef tsk As Task, ByRef errors As Collection)
    On Error GoTo err_handler
    Dim deps As String
    deps = tsk.UniqueIDPredecessors & DELIM & get_soft_dependencies(tsk)
    deps = trim_delimiters(deps)
    tsk.UniqueIDPredecessors = deps
    Exit Sub
    
err_handler:
    msg = "Task " + Str(tsk.ID) + " (" + tsk.Name + "): " + get_soft_dependencies(tsk) + " => " + Err.Description
    errors.Add msg
End Sub


Private Sub report_errors(c As Collection)
    If (c.Count = 0) Then
        Exit Sub
    End If
    
    Dim s As String
    For Each v In c
        s = s + v + vbCrLf
    Next v
    MsgBox s, , "Errors adding soft dependencies"
End Sub


Private Function remove_deps(dependencies As String, remove_dep As String) As String
    lhs_a = Split(Replace(dependencies, " ", ""), DELIM)
    rhs_a = Split(Replace(remove_dep, " ", ""), DELIM)
    remove_deps = Join(array_subtract(lhs_a, rhs_a), DELIM)
End Function


' Remove the rhs array elements from the lhs array.
Private Function array_subtract(lhs As Variant, rhs As Variant) As Variant
    Dim c As Collection
    Set c = New Collection
    For i = LBound(lhs) To UBound(lhs)
        found = False
        For j = LBound(rhs) To UBound(rhs)
            If (lhs(i) = rhs(j)) Then
                found = True
            End If
        Next j
        If (found = False) Then
            c.Add lhs(i)
        End If
    Next i
    
    Dim ret As Variant
    ret = Array()
    If (c.Count > 0) Then
        ReDim ret(c.Count - 1)
    End If
    For i = 0 To c.Count - 1
        ret(i) = c.Item(i + 1)
    Next i
    array_subtract = ret
End Function


Private Function trim_delimiters(s As String)
    Dim ret As String
    ret = Trim(s)
    If (Left(ret, Len(DELIM)) = DELIM) Then
        ret = Right(ret, Len(ret) - 1)
    End If
    If (Right(ret, Len(DELIM)) = DELIM) Then
        ret = Left(ret, Len(ret) - 1)
    End If
    trim_delimiters = ret
End Function


' ////////////////////////////////////////////////////
' Tests during development

Private Sub test_suite()
    On Error GoTo err_handler
    Call test_array_subtract
    Debug.Print "OK"
    Exit Sub
err_handler:
    Debug.Print "FAILED: " + Err.Description
End Sub

Private Sub test_array_subtract()
    assert_subtract_equals "1", "2", "1", "not found"
    assert_subtract_equals "1", "1", "", "found single number"
    assert_subtract_equals "1,2", "1", "2", "found at the start"
    assert_subtract_equals "2,1", "1", "2", "found at the end"
    assert_subtract_equals "2,1,3", "1", "2,3", "found in the middle"
    
    assert_subtract_equals "11,2", "1", "11,2", "NOT found at the start"
    assert_subtract_equals "2,11", "1", "2,11", "NOT found at the end"
    assert_subtract_equals "2,11,3", "1", "2,11,3", "NOT found in the middle"
    
    assert_subtract_equals "", "1", "", "empty lhs"
    assert_subtract_equals "1", "", "1", "empty rhs"
    
    assert_subtract_equals "1,2,3", "1,2", "3", "multiple"
    assert_subtract_equals "1,2,3", "1,3", "2", "multi middle left"
    assert_subtract_equals "1,2,3,4,5", "5,4,3,1", "2", "out of order"
    assert_subtract_equals "1,2,3", "1,2,3", "", "complete deletion"
    
    assert_subtract_equals "a, b, c", "b", "a,c", "spaces ignored"
End Sub

Private Sub assert_subtract_equals(lhs As String, rhs As String, expected As String, Optional msg As String = "")
    actual = remove_deps(lhs, rhs)
    If (actual <> expected) Then
        Err.Raise vbObjectError, , msg + ": Expected " + expected + ", but got " + actual
    End If
End Sub

