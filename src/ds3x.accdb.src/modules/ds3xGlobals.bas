Attribute VB_Name = "ds3xGlobals"
Option Compare Database
Option Explicit
Option Base 0


' --- ACCESS WINDOW HIDE / SHOW ---

Public Const SW_HIDE = 0
Public Const SW_SHOWNORMAL = 1
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOW = 5


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


' --- MemoryLib ---

#If Win64 Then
    Public Const PTR_SIZE As Long = 8
    Public Const VARIANT_SIZE As Long = 24
#Else
    Public Const PTR_SIZE As Long = 4
    Public Const VARIANT_SIZE As Long = 16
#End If

Public Const BYTE_SIZE As Long = 1
Public Const INT_SIZE As Long = 2
Public Const VT_SPACING As Long = VARIANT_SIZE / INT_SIZE 'VarType spacing in an array of Variants

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

Public Type SAFEARRAY_2D
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
    rgsabound0(0 To 1) As SAFEARRAYBOUND
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

