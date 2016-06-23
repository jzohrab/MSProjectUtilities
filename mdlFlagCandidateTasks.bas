Attribute VB_Name = "mdlFlagCandidateTasks"
' Flag all tasks with no parents, or whose parents have all been completed.
' These tasks are tasks that can potentially be started.
Public Sub FlagUnblockedTasks()
    Dim t As Task
    For Each t In ThisProject.Tasks
        If Not (t Is Nothing) Then
            If Not t.Summary Then
                Dim should_start As Boolean
                should_start = ((t.PercentComplete <> 100) And Not taskIsBlocked(t))
                t.Text1 = IIf(should_start, "Yes", "")
            End If
        End If
    Next 't
End Sub

Private Function taskIsBlocked(t As Task) As Boolean
    If t.PercentComplete = 100 Then
        taskIsBlocked = False
        Exit Function
    End If
    
    ' If this is part of a summary task, the summary
    ' task may itself be blocked due predecessors.
    ' If that is the case, this task should be blocked.
    If (t.OutlineLevel > 1) Then
        If (taskIsBlocked(t.OutlineParent)) Then
            taskIsBlocked = True
            Exit Function
        End If
    End If
    
    ' Debug.Print t.Name
    Dim p As TaskDependency
    For Each p In t.TaskDependencies
        If (p.Type = pjFinishToStart And p.From.ID <> t.ID) Then
            ' Debug.Print p.From.Name
            If p.From.PercentComplete <> 100 Then
                taskIsBlocked = True
                Exit Function
            Else
                ' If parent is blocked, this is blocked.
                If taskIsBlocked(p.From) Then
                    taskIsBlocked = True
                    Exit Function
                End If
            End If
        End If
    Next 'p
    
    ' If reached here, all the parent tasks are complete,
    ' or we don't have parent tasks.
    taskIsBlocked = False
End Function

