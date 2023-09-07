Attribute VB_Name = "ds3xGlobals"
Option Compare Database
Option Explicit
Option Base 0

Public GLOBAL_TIMER_ACC As Single
Public GLOBAL_rX As RecordsetEx
Public GLOBAL_dsT As dsTable
Public GLOBAL_GC_dbQ As ArrayListEx

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


' --- DEBUG ---

Public Sub AllTestsRecordsetEx()
    dsApp.GlobalVMemAnalysis "INIT"
'    PreTestCreateArrayList
'    PreTestCreateFirstTable
    TestRecordsetExVMemAlloc
    SecondTestRecordsetEx
    ThirdTestRecordsetEx
End Sub

Public Sub PreTestCreateArrayList()
    dsApp.GlobalVMemAnalysis "BEGIN PRE aList TEST"
    Dim aList As ArrayList
    dsApp.GlobalVMemAnalysis "PRE aList - 0"
    Set aList = New ArrayList
    dsApp.GlobalVMemAnalysis "PRE aList - 1"
    aList.Add "Hello"
    dsApp.GlobalVMemAnalysis "PRE aList - 2"
    
End Sub


Public Sub PreTestCreateFirstTable()
    dsApp.GlobalVMemAnalysis "BEGIN PRE TEST"
    
    Dim dsT As dsTable, aX As ArrayListEx
    
    dsApp.GlobalVMemAnalysis "PRE - 0"
    
    Set aX = ArrayListEx.Create() _
        .Add(Array(0, "Hello AA", 1113)) _
        .Add(Array(1, "Hello AA2", 22433)) _
        .Add(Array(2, "Hello AA3", 33)) _
        .Add(Array(3, "Hello AA4 44", 4444)) _
        .Add(Array(4, "Hello AA5.# DONE!", 55055))
        
    dsApp.GlobalVMemAnalysis "PRE - 1"
        
    Set dsT = dsTable.Create(aX, False)
    dsApp.GlobalVMemAnalysis "PRE - 2"
    dsT.SetHeaders (Array("ID", "TEXT", "VALUE"))
    dsApp.GlobalVMemAnalysis "PRE - 3"
    
End Sub

Public Sub TestRecordsetExVMemAlloc()
    Dim dbQ As New dbQuery, rX As RecordsetEx
    
'    dbQ.CloseOnTerminate = False
    
    dsApp.GlobalVMemAnalysis "BEGIN"
    Set GLOBAL_rX = RecordsetEx.Create(dbQ.Create("SELECT TOP (1000000) * FROM [dbo].[T_ATM_TEST_CUSTOMERS]"))
    dsApp.GlobalVMemAnalysis "GLOBAL_rX CREATED"
    Set GLOBAL_dsT = dsTable.Create(GLOBAL_rX, True)
    dsApp.GlobalVMemAnalysis "GLOBAL_dsT CREATED"
    
    Set GLOBAL_GC_dbQ = ArrayListEx.Create()
    GLOBAL_GC_dbQ.Add dbQ
End Sub

Public Sub SecondTestRecordsetEx()
    dsApp.GlobalVMemAnalysis "BEGIN SECOND TEST"
    
    Set GLOBAL_dsT = Nothing
    dsApp.GlobalVMemAnalysis "GLOBAL_dsT = Nothing"
    
'    GLOBAL_rX.Instance.ActiveConnection.Close
    Set GLOBAL_rX.Instance.ActiveConnection = Nothing
    Set GLOBAL_rX.Instance = Nothing
    Set GLOBAL_rX = Nothing
    dsApp.GlobalVMemAnalysis "GLOBAL_rX = Nothing"
    
End Sub

Public Sub ThirdTestRecordsetEx()
    dsApp.GlobalVMemAnalysis "BEGIN THIRD TEST"
    GLOBAL_GC_dbQ(0).Connection.Close
    GLOBAL_GC_dbQ.Clear
    dsApp.GlobalVMemAnalysis "THIRD - 0"
    Set GLOBAL_GC_dbQ = Nothing
    dsApp.GlobalVMemAnalysis "THIRD - 1"
End Sub

