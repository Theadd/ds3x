Attribute VB_Name = "ds3xGlobals"
Option Compare Database
Option Explicit
Option Base 0


' --- ACCESS WINDOW HIDE / SHOW ---

Global Const SW_HIDE = 0
Global Const SW_SHOWNORMAL = 1
Global Const SW_SHOWMINIMIZED = 2
Global Const SW_SHOWMAXIMIZED = 3
Global Const SW_SHOW = 5
'Forces a top-level window onto the taskbar when the window is visible.
Public Const WS_EX_APPWINDOW As Long = &H40000


' --- ScreenLib Types ---

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type BOUNDS
    X As Long
    Y As Long
    W As Long
    h As Long
End Type


' --- MemoryLib Types ---

Public Type REMOTE_MEMORY
    memValue As Variant
    remoteVT As Variant
    isInitialized As Boolean
End Type

Public Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type

Public Type SAFEARRAY_1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    #If Win64 Then
        dummyPadding As Long
        pvData As LongLong
    #Else
        pvData As Long
    #End If
    rgsabound0 As SAFEARRAYBOUND
End Type


' --- Automation ---

Public Function RunApplicationCommandArgs()
    #If AutomationSupport = 1 Then
        dsApp.ExecuteAutomationCommandArgs VBA.Command$()
    #End If
End Function

#If AutomationSupport = 1 Then

    Public Function IsTaskRunning(Optional ByVal TaskNamePattern As String = "*") As Boolean
        If IsEmpty(dsApp.ActiveTask) Then Exit Function
        IsTaskRunning = ((dsApp.ActiveTask(0) Like "*" & TaskNamePattern & "*") Or (dsApp.ActiveTask(1) Like "*" & TaskNamePattern & "*"))
    End Function
    
    Public Function HasFailedToRunAllTasks() As Boolean
        HasFailedToRunAllTasks = Not (IsEmpty(dsApp.FailedTask))
    End Function
    
    Public Sub SetCustomVar(ByVal VarName As String, ByVal VarValue As Variant)
        dsApp.CustomVar(VarName) = VarValue
    End Sub
    
    ' Adds a preset file to the runnable tasks queue
    Public Sub AddRunnableTask(ByVal TargetPath As String, Optional ByVal RunnableTaskName As String = "", Optional ByVal OnErrorResumeNext As Boolean = False)
        dsApp.RunnableTasks.Add Array(TargetPath, RunnableTaskName, OnErrorResumeNext)
    End Sub
    
    Public Sub ClearAllRunnableTasks()
        dsApp.RunnableTasks.Clear
    End Sub
    
    ' Sequentially executes all runnable tasks in queue
    Public Sub RunAllAsync()
        DoCmd.OpenForm "DS_ASYNC_RUNNER", WindowMode:=acHidden
        Forms("DS_ASYNC_RUNNER").RunAsync
    End Sub
    
    Public Function NumTasksInQueue() As Long
        NumTasksInQueue = dsApp.RunnableTasks.Count
    End Function
    
#End If


' --- Utility Functions ---

' USAGE: Printf("Name: %1, Age: %2", "John", 32) -> "Name: John, Age: 32"
Public Function Printf(ByVal mask As String, ParamArray Tokens() As Variant) As String
    Dim parts() As String: parts = Split(mask, "%")
    Dim i As Long, j As Long, isFound As Boolean, s As String
    
    For i = LBound(parts) + 1 To UBound(parts)
        If LenB(parts(i)) = 0 Then
            parts(i) = "%"
        Else
            isFound = False
            For j = UBound(Tokens) To LBound(Tokens) Step -1
                s = CStr(j + 1)
                If Left$(parts(i), Len(s)) = s Then
                    parts(i) = Tokens(j) & Right$(parts(i), Len(parts(i)) - Len(s))
                    isFound = True
                    Exit For
                End If
            Next j
            If Not isFound Then
                parts(i) = "%" & parts(i)
            End If
        End If
    Next i
    
    Printf = Join(parts, vbNullString)
End Function


' --- TO REMOVE ---

Public Sub TestArraySlicingStuff()
    Dim t() As Variant, sgX As ArraySliceGroup, a2dX As Array2dEx, sg2 As ArraySliceGroup, sl3x As ArraySlice, vReverse As Variant, b2dX As Array2dEx, c2dX As Array2dEx
    ReDim t(0 To 5, 0 To 3)
    
    t(0, 0) = "not zero":    t(0, 1) = 300:  t(0, 2) = "ZERO":   t(0, 3) = 130
    t(1, 0) = "-1":   t(1, 1) = 310:  t(1, 2) = "ONE":    t(1, 3) = 131
    t(2, 0) = "2ac":    t(2, 1) = 320:  t(2, 2) = "TWO":    t(2, 3) = 132
    t(3, 0) = "f -3a":   t(3, 1) = 330:  t(3, 2) = "THREE":  t(3, 3) = 133
    t(4, 0) = 4:    t(4, 1) = 340:  t(4, 2) = "FOUR":   t(4, 3) = 134
    t(5, 0) = "It's -5":   t(5, 1) = 350:  t(5, 2) = "FIVE":   t(5, 3) = 135

    Set a2dX = Array2dEx.Create()
    'a2dX.Instance = CollectionsLib.GetArrayByRef(t)
    a2dX.Instance = t
    
    Set sgX = ArraySliceGroup.Create(a2dX)
    Set sg2 = sgX.GetRange(1, 3, Array(0, 1, 2))

    Set c2dX = sg2.Instance

    Debug.Print JSON.Stringify(c2dX)
End Sub


