Attribute VB_Name = "dsAppGlobals"
Option Compare Database
Option Explicit
Option Base 0


' --- dsAppGlobals Module ---

#If AutomationSupport = 1 Then
    Private pActiveTask As Variant
    Private pFailedTask As Variant
    Private pRunnableTasks As ArrayListEx
    Private pLiveEd As dsLiveEd
#End If

Private pCustomVars As DictionaryEx


Public Property Get CustomVars() As DictionaryEx
    If pCustomVars Is Nothing Then InitializeCustomVars
    Set CustomVars = pCustomVars
End Property

Public Property Get CustomVar(ByVal VarName As String) As Variant
    If pCustomVars Is Nothing Then InitializeCustomVars
    If Not (Left(VarName, 2) = "${") Then VarName = "${" & VarName & "}"
    CustomVar = pCustomVars.GetValue(VarName, vbNullString)
End Property

Public Property Let CustomVar(ByVal VarName As String, ByVal VarValue As Variant)
    If pCustomVars Is Nothing Then InitializeCustomVars
    If Not (Left(VarName, 2) = "${") Then VarName = "${" & VarName & "}"
    pCustomVars(VarName) = VarValue
End Property

Public Property Get ApplyCustomVarsOn(ByVal Target As Variant) As Variant
    If pCustomVars Is Nothing Then InitializeCustomVars
    If pCustomVars.Count = 0 Then GoTo NoCustomVars
    
    If IsObject(Target) Then
        Set ApplyCustomVarsOn = ReplaceCustomVars(Target)
    Else
        ApplyCustomVarsOn = ReplaceCustomVars(Target)
    End If
    
    Exit Property
NoCustomVars:
    If IsObject(Target) Then
        Set ApplyCustomVarsOn = Target
    Else
        ApplyCustomVarsOn = Target
    End If
End Property


#If AutomationSupport = 1 Then
    
    Public Property Get RunnableTasks() As ArrayListEx
        If pRunnableTasks Is Nothing Then Set pRunnableTasks = ArrayListEx.Create()
        Set RunnableTasks = pRunnableTasks
    End Property
    
    Public Property Get RunAll() As Boolean
        Dim t As Variant, shouldContinue As Boolean, g As Variant
        pFailedTask = Empty
        
        For Each t In RunnableTasks
            pActiveTask = t
            Set pLiveEd = Nothing
            
            Set pLiveEd = New dsLiveEd
            shouldContinue = pLiveEd.ImportPreset(ApplyCustomVarsOn(CStr(t(0))))
            If shouldContinue Then
                shouldContinue = pLiveEd.TryApply(g)
                If Not shouldContinue Then
                    If CBool(t(2)) Then shouldContinue = True   ' Where `t(2)` is the `OnErrorResumeNext` (optional) parameter provided to the `AddRunnableTask()` call.
                End If
            End If
            
            If shouldContinue Then
                pActiveTask = Empty
                RunnableTasks.RemoveAt 0
            Else
                pFailedTask = pActiveTask
                pActiveTask = Empty
                Exit For
            End If
        Next t
        
        Set pLiveEd = Nothing
        RunAll = shouldContinue
    End Property
    
    
    
    Public Function IsTaskRunning(Optional ByVal TaskNamePattern As String = "*") As Boolean
        If IsEmpty(pActiveTask) Then Exit Function
        IsTaskRunning = ((pActiveTask(0) Like "*" & TaskNamePattern & "*") Or (pActiveTask(1) Like "*" & TaskNamePattern & "*"))
    End Function
    
    Public Function HasFailedToRunAllTasks() As Boolean
        HasFailedToRunAllTasks = Not (IsEmpty(pFailedTask))
    End Function
    
    Public Sub SetCustomVar(ByVal VarName As String, ByVal VarValue As Variant)
        CustomVar(VarName) = VarValue
    End Sub
    
    ' Adds a preset file to the runnable tasks queue
    Public Sub AddRunnableTask(ByVal TargetPath As String, Optional ByVal RunnableTaskName As String = "", Optional ByVal OnErrorResumeNext As Boolean = False)
        RunnableTasks.Add Array(TargetPath, RunnableTaskName, OnErrorResumeNext)
    End Sub
    
    Public Sub ClearAllRunnableTasks()
        RunnableTasks.Clear
    End Sub
    
    ' Sequentially executes all runnable tasks in queue
    Public Sub RunAllAsync()
        DoCmd.OpenForm "DS_ASYNC_RUNNER", WindowMode:=acHidden
        Forms("DS_ASYNC_RUNNER").RunAsync
    End Sub
    
    Public Function NumTasksInQueue() As Long
        NumTasksInQueue = RunnableTasks.Count
    End Function

#Else
    
    Public Property Get RunAll() As Boolean
        Err.Raise 425
    End Property
    
#End If


Public Function CreateDSLiveEd() As dsLiveEd
    Set CreateDSLiveEd = New dsLiveEd
End Function


' --- CustomVars ---

Private Sub InitializeCustomVars()
    Set pCustomVars = DictionaryEx.Create()
    
    pCustomVars.Add "${Timestamp}", CStr(DateDiff("s", DateValue("1970-01-01"), Now()))
    pCustomVars.Add "${Date}", VBA.Format$(Date, "yyyymmdd")
    pCustomVars.Add "${DateTime}", VBA.Format$(Now(), "yyyymmddhhMMss")
    pCustomVars.Add "${Year}", VBA.Format$(Date, "yyyy")
    pCustomVars.Add "${Month}", VBA.Format$(Date, "mm")
    pCustomVars.Add "${Day}", VBA.Format$(Date, "dd")
    pCustomVars.Add "${Username}", Nz(CreateObject("WScript.Network").UserName, "")
    pCustomVars.Add "${UserProfile}", VBA.Environ$("USERPROFILE")
    pCustomVars.Add "${Temp}", VBA.Environ$("TEMP")
End Sub

Private Function ReplaceCustomVars(ByVal Target As Variant) As Variant
    Select Case CLng(VarType(Target))
        Case Is = CLng(vbString)
            ReplaceCustomVars = ReplaceCustomVarsOnString(Target)
        Case Is >= CLng(vbArray)
            ReplaceCustomVars = ReplaceCustomVarsOnIterable(Target)
        Case Else
            If IsObject(Target) Then
                Set ReplaceCustomVars = Target
            Else
                ReplaceCustomVars = Target
            End If
    End Select
End Function

Private Function ReplaceCustomVarsOnIterable(ByVal Target As Variant) As Variant
    Dim aX As ArrayListEx, Item As Variant
    Set aX = ArrayListEx.Create()
    
    For Each Item In Target
        aX.Add ReplaceCustomVars(Item)
    Next Item
    
    ReplaceCustomVarsOnIterable = aX.ToArray()
End Function

Private Function ReplaceCustomVarsOnString(ByVal Target As String) As String
    Dim cKey As Variant
    
    For Each cKey In pCustomVars.Keys()
        If InStr(1, Target, CStr(cKey), vbTextCompare) > 0 Then
            Target = Replace(Target, CStr(cKey), CStr(pCustomVars(cKey)), Compare:=vbTextCompare)
        End If
    Next cKey
    
    ReplaceCustomVarsOnString = Target
End Function

' ---



