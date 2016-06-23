Attribute VB_Name = "mdlFindRedundantDependencies"
' Finds redundant predecessors.
' Redundants are shown in a message box, and printed to the Debug window.
Public Sub FindRedundantDependencies()
    Dim t As Task
    Dim c As Collection
    Set c = New Collection
    For Each t In ActiveProject.Tasks
        ' Debug.Print t.Name
        Call task_has_redundant_predecessors(t, c)
    Next t
    
    If (c.Count = 0) Then
        MsgBox "No redundant dependencies."
        Exit Sub
    End If
    
    Dim s As Variant
    Dim msg As String
    msg = ""
    For Each s In c
        msg = msg & s & vbCrLf
    Next s
    MsgBox msg, Title:="Redundant dependencies."
End Sub


Private Sub task_has_redundant_predecessors(t As Task, collector As Collection)
    If (t Is Nothing) Then
        Exit Sub
    End If
        
    If (t.PredecessorTasks.Count = 0) Then
        Exit Sub
    End If
    
    Dim pt As Task
    Dim other_pt As Task
    For Each pt In t.PredecessorTasks
        For Each other_pt In t.PredecessorTasks
            If (pt.ID <> other_pt.ID) Then
                If (check_has_predecessor(other_pt, pt.ID)) Then
                    Dim msg As String
                    msg = "Task TASKID (TASKNAME): Predecessor OTHERID (OTHERNAME) already contains PTID (PTNAME)"
                    msg = Replace(msg, "TASKID", t.ID)
                    msg = Replace(msg, "TASKNAME", t.Name)
                    msg = Replace(msg, "PTID", CStr(pt.ID))
                    msg = Replace(msg, "PTNAME", pt.Name)
                    msg = Replace(msg, "OTHERID", CStr(other_pt.ID))
                    msg = Replace(msg, "OTHERNAME", other_pt.Name)
                    Debug.Print msg
                    collector.Add msg
                End If
            End If
        Next other_pt
    Next pt
End Sub


Private Function check_has_predecessor(t As Task, task_id As Integer) As Boolean
    ' Debug.Print "Checking task " + CStr(t.ID) + "(" + t.Name + ") for predecessor " + CStr(task_id)
    
    If (t Is Nothing) Then
        check_has_predecessor = False
        Exit Function
    End If
    
    If (t.ID = task_id) Then
        check_has_predecessor = True
        Exit Function
    End If
    
    If (t.PredecessorTasks.Count = 0) Then
        check_has_predecessor = False
        Exit Function
    End If
        
    Dim pt As Task
    For Each pt In t.PredecessorTasks
        If (task_id = pt.ID) Then
            check_has_predecessor = True
            Exit Function
        End If
        
        If (check_has_predecessor(pt, task_id)) Then
            check_has_predecessor = True
            Exit Function
        End If
    Next pt
    
    check_has_predecessor = False
End Function
