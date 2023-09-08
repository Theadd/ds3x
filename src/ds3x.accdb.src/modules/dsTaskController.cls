VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "dsTaskController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder "ds3x.LiveEditor"
Option Compare Database
Option Explicit
Option Base 0

' dsTaskController

' Err.Raise:
' 9 - Subscript out of range (Array index out of bounds)
' 17 - Can't perform requested operation
' 13 - Type mismatch
' 51 - Internal error
' 70 - Permission denied
' 327 - Data value named not found
' 328 - Illegal parameter; can't write arrays
' 380 - Invalid property value
' 419 - Permission to use object denied
' 422 - Property not found
' 423 - Property or method not found
' 424 - Object Required
' 425 - Invalid object use
' 448 - Named Not Argument
' 449 - Argument not optional or invalid property assignment
' 450 - Wrong number of arguments or invalid property assignment
' 31001 - Out of memory
' 31004 - No Object
' 65536 - Invalid procedure call or argument

Public Event OnTableChange()
Public Event OnDataSourceChange(ByVal dsId As String)
Public Event OnRebuildSequenceChange()
Public Event OnActiveSequenceIndexChange(ByVal TargetIndex As Long, ByVal HasFailed As Boolean)
Public Event OnRebuildSequenceFail(ByVal FailedSequenceIndex As Long)
Public Event OnBusyChange(ByVal IsBusy As Boolean)

Private pRebuildSequence As New ArrayListEx
Private pSequenceIndex As Long
' The lowest index within the RebuildSequence array that needs rebuilding.
Private pDirtyRebuildIndex As Long
Private pTaskResults As DictionaryEx
Private pAvailableTasks As dsTaskDefinitions
Private pTable As dsTable
Private pWorksheet As xlSheetsEx
Private pDataSources As DictionaryEx
Private pGarbageCollector As DictionaryEx

'    {
'      "AddRowNumbers": [
'        "AddRowNumbers",
'        [
'          "[Source]: Table, [FirstNumber = 1]: Long",
'          2,
'          [
'            ["[Source]", "Table"],
'            ["[FirstNumber = 1]", "Long"]
'          ]
'        ],
'        "",            /* Description */
'        "",            /* Usage */
'        "T",           /* TaskType */
'        "[TD]*"        /* Requires */
'      ],
'      ...
'    }
' TaskInfo(TaskName)(DS_T_PARAMS)(2)

Const DS_T_NAME As Long = 0
Const DS_T_PARAMS As Long = 1
Const DS_T_DESC As Long = 2
Const DS_T_USAGE As Long = 3
Const DS_T_TYPE As Long = 4
Const DS_T_REQUIRES As Long = 5



Public Property Get RebuildSequence() As ArrayListEx: Set RebuildSequence = pRebuildSequence: End Property
Public Property Set RebuildSequence(ByRef Value As ArrayListEx): SetRebuildSequence Value: End Property
Public Property Get Table() As dsTable: Set Table = pTable: End Property
Public Property Get IsValidTable() As Boolean: IsValidTable = Not (pTable Is Nothing): End Property
Public Property Get Worksheet() As xlSheetsEx: Set Worksheet = pWorksheet: End Property
Public Property Get DataSource(ByVal ID As String) As Object: Set DataSource = pDataSources(ID): End Property
Public Property Set DataSource(ByVal ID As String, ByRef Value As Object): SetDataSource ID, Value: End Property
Public Property Get SequenceIndex() As Long: SequenceIndex = pSequenceIndex: End Property
Public Property Let SequenceIndex(ByVal Value As Long): SetActiveSequenceIndex Value: End Property
Public Property Get Initialized() As Boolean: Initialized = (pRebuildSequence.Count > 0): End Property
Public Property Get TaskInfo(ByVal TaskName As String) As Variant: TaskInfo = pAvailableTasks.Tasks(TaskName): End Property
Public Property Get AvailableTasks() As dsTaskDefinitions: Set AvailableTasks = pAvailableTasks: End Property
Public Property Get DirtyRebuildIndex() As Long: DirtyRebuildIndex = pDirtyRebuildIndex: End Property
Public Property Get GeneratedValue() As Variant: Assign GeneratedValue, pTaskResults(RebuildSequence(SequenceIndex)("Id")): End Property



Public Sub AddTask(ByVal TaskName As String, ByVal TaskValues As Variant, Optional ByVal TaskId As String = "")
    Dim sIndex As Long: sIndex = SequenceIndex
    
    If sIndex < RebuildSequence.Count - 1 Then
        RebuildSequence().Insert sIndex + 1, GenerateTask(TaskName, TaskValues, TaskId)
    Else
        RebuildSequence().Add GenerateTask(TaskName, TaskValues, TaskId)
    End If
    
    MakeRebuildSequenceDirtyOn sIndex + 1

    RaiseEvent OnRebuildSequenceChange
    SequenceIndex = sIndex + 1
End Sub

Public Sub RemoveTask(ByVal TaskId As String)
    Dim aX As ArrayListEx, TaskIndex As Long, iAux As Long
    
    Set aX = ArrayListEx.Create(RebuildSequence)
    If TryGetIndexOfTaskId(TaskId, TaskIndex) Then
        aX.RemoveAt TaskIndex
        Set RebuildSequence = aX
        
        If Not TryGetIndexOfTaskId(TaskId, iAux) Then
            If pTaskResults.Exists(TaskId) Then
                Set pTaskResults(TaskId) = Nothing
                pTaskResults.Remove TaskId
            End If
            If pDataSources.Exists(TaskId) Then
                Set pDataSources(TaskId) = Nothing
                pDataSources.Remove TaskId
            End If
            Set pTable = Nothing
            MakeRebuildSequenceDirtyOn -1
            Rebuild -1
        End If
    End If
    
End Sub

' Used to update an existing task at specified index
Public Sub SetTask(ByVal Task As Scripting.Dictionary, ByVal Index As Long)
    Set pRebuildSequence(Index) = Task
    MakeRebuildSequenceDirtyOn Index
    
    RaiseEvent OnRebuildSequenceChange
    SequenceIndex = Index
End Sub

Public Sub RebuildAll()
    Dim i As Long
    pTaskResults.RemoveAll
    MakeRebuildSequenceDirtyOn 0
    SequenceIndex = RebuildSequence.Count - 1
End Sub

Private Function Rebuild(ByVal TargetIndex As Long) As Boolean
    Dim i As Long, oTable As dsTable, dX As New DictionaryEx, oWorksheet As xlSheetsEx, bFailed As Boolean
    
    RaiseEvent OnBusyChange(True)
'    dsApp.GlobalVMemAnalysis "@Rebuild(" & CStr(TargetIndex) & ")"
    If TargetIndex >= pDirtyRebuildIndex Then
        If TargetIndex > pRebuildSequence.Count - 1 Then Err.Raise 9, , "Subscript out of range (Array index out of bounds)"
        If TryGetSourceTableForSequence(pDirtyRebuildIndex, oTable) Then
            Set pTable = oTable
        End If
        If TryGetSourceWorksheetForSequence(pDirtyRebuildIndex, oWorksheet) Then
            Set pWorksheet = oWorksheet
        End If

        On Error GoTo RebuildTaskFail
        For i = pDirtyRebuildIndex To TargetIndex
            pSequenceIndex = i
            RebuildTask dX.Bind(pRebuildSequence.Item(i)).Instance
            pDirtyRebuildIndex = i + 1
        Next i
    End If
    
    On Error GoTo 0
    If TryGetSourceTableForSequence(TargetIndex + 1, oTable) Then
        Set pTable = oTable
        pSequenceIndex = TargetIndex
        If TryGetSourceWorksheetForSequence(TargetIndex + 1, oWorksheet) Then
            Set pWorksheet = oWorksheet
        End If
    End If
    
ExitRebuild:
'    dsApp.GlobalVMemAnalysis "@Rebuild(" & CStr(TargetIndex) & ") BEFORE OnTableChange"
    RaiseEvent OnTableChange
'    dsApp.GlobalVMemAnalysis "@Rebuild(" & CStr(TargetIndex) & ") BEFORE OnTableChange"
    RaiseEvent OnBusyChange(False)
    Rebuild = Not bFailed
    Exit Function
RebuildTaskFail:
    bFailed = True
    MakeRebuildSequenceDirtyOn pSequenceIndex
    RaiseEvent OnRebuildSequenceFail(pSequenceIndex)
    Resume ExitRebuild
End Function

Private Sub RebuildTask(ByRef Task As Scripting.Dictionary)
    Static dX As New DictionaryEx
    Static cb As New Callback
    Dim t() As Variant, TaskName As String, TaskParams As Variant
    
    dsApp.GlobalVMemAnalysis "", False
    dX.Bind Task
    TaskName = dX("TaskName")
    TaskParams = TaskInfo(TaskName)(DS_T_PARAMS)(2)
    cb.Bind pAvailableTasks, TaskName
    
    If TryRebuildTaskValues(t, TaskParams, dX.GetValue("Values", Array())) Then
        Select Case GetReturnTypeOf(dX("TaskName"))
            Case "Table"
                Set pTable = cb.ExecuteOnArray(t).Result
                Set pTaskResults(dX("Id")) = pTable
            Case "Worksheet"
                Set pWorksheet = cb.ExecuteOnArray(t).Result
                Set pTaskResults(dX("Id")) = pWorksheet
            Case Else
                Err.Raise 13
        End Select
    Else
        Err.Raise 13
    End If
    
    dsApp.GlobalVMemAnalysis "@RebuildTask(" & dX("Id") & " " & TaskName & ")"
'    VMemAnalysis "@" & dX("Id") & " " & TaskName
End Sub




Private Sub Class_Initialize()
    pDirtyRebuildIndex = 0
    Set pTaskResults = DictionaryEx.Create()
    Set pGarbageCollector = DictionaryEx.Create()
    Set pDataSources = DictionaryEx.Create()
    Set pTable = Nothing
    Set pAvailableTasks = New dsTaskDefinitions
    Set pAvailableTasks.TaskController = Me
End Sub

Private Sub Class_Terminate()
    Dispose
End Sub

Public Sub Dispose()
    On Error Resume Next
'    Debug.Print "[INFO] @dsTaskController.Dispose()"
    Set pTable = Nothing
    pTaskResults.RemoveAll
    pDataSources.RemoveAll
    ' Run GarbageCollector
    FreeUnlinkedResources
    pDataSources.Unbind
    Set pTaskResults = Nothing
    Set pDataSources = Nothing
    Set pGarbageCollector = Nothing
    Set pAvailableTasks.TaskController = Nothing
    Set pAvailableTasks = Nothing
End Sub




Private Sub SetRebuildSequence(ByRef Value As ArrayListEx)
    Dim ds1 As dsTable, rs1 As xlSheetsEx
    
    If pDataSources.Exists("DS1") Then
        Set ds1 = pDataSources("DS1")
    End If
    If pDataSources.Exists("XS1") Then
        Set rs1 = pDataSources("XS1")
    End If
    pTaskResults.RemoveAll
    Set pTable = Nothing
    Set pWorksheet = Nothing
    pSequenceIndex = -1
    MakeRebuildSequenceDirtyOn 0
    pDataSources.Unbind
    pDataSources.Bind DictionaryEx.Create().Instance
    If Not ds1 Is Nothing Then
        Set pDataSources("DS1") = ds1
    End If
    If Not rs1 Is Nothing Then
        Set pDataSources("XS1") = rs1
    End If
    Set pRebuildSequence = Value
    RaiseEvent OnRebuildSequenceChange
    SequenceIndex = 0
    FreeUnlinkedResources
End Sub

Private Sub SetDataSource(ByVal TaskId As String, ByRef Value As Object)
    On Error Resume Next
    Dim Index As Long, dsT As dsTable, xlS As xlSheetsEx, sourceType As String
    
    Select Case TypeName(Value)
        Case "dsTable"
            Set dsT = Value
            sourceType = "DS"
        Case "xlSheetsEx"
            Set xlS = Value
            sourceType = "XS"
        Case Else
            Stop
    End Select
    
'    Select Case TaskId
'        Case "DS1", "XS1"
'            pTaskResults.RemoveAll
'            pSequenceIndex = -1
'            Set pTable = Nothing
'            Set pWorksheet = Nothing
'            MakeRebuildSequenceDirtyOn 0
'            Set pDataSources(TaskId) = Nothing
'            Select Case sourceType
'                Case "DS"
'                    Set pDataSources(TaskId) = dsT
'                    Set pTable = dsT
'                Case "XS"
'                    Set pDataSources(TaskId) = xlS
'                    Set pWorksheet = xlS
'                Case Else: Stop
'            End Select
'        Case Else
            If TryGetIndexOfTaskId(TaskId, Index) Then
                MakeRebuildSequenceDirtyOn Index
                Select Case sourceType
                    Case "DS"
                        Set pDataSources(TaskId) = dsT
                    Case "XS"
                        Set pDataSources(TaskId) = xlS
                    Case Else: Stop
                End Select
            Else
                Err.Raise 327, , "DataSource """ & TaskId & """ not found."
            End If
'    End Select
    
    RaiseEvent OnDataSourceChange(TaskId)
End Sub

Friend Function TryGetIndexOfTaskId(ByVal TaskId As String, ByRef OutIndex As Long) As Boolean
    Dim i As Long
    On Error GoTo Finally
    
    For i = 0 To pRebuildSequence.Count - 1
        If pRebuildSequence(i)("Id") = TaskId Then
            OutIndex = i
            TryGetIndexOfTaskId = True
            Exit Function
        End If
    Next i
    
Finally:
End Function

Private Sub SetActiveSequenceIndex(ByVal Value As Long)
    Dim rebuildFail As Boolean
    If Value < 0 Then Exit Sub
    If Value >= RebuildSequence.Count Then Exit Sub
    On Error GoTo Finally
    
    pSequenceIndex = Value
    rebuildFail = Not Rebuild(Value)
Finally:
    On Error GoTo 0
    RaiseEvent OnActiveSequenceIndexChange(Value, rebuildFail)
End Sub

Private Function GetNextSequenceId(ByVal IdPrefix As String) As String
    Dim Item As Variant, dX As New DictionaryEx, curId As Long, maxId As Long
    
    For Each Item In RebuildSequence
        dX.Bind Item
        If dX("Id") Like IdPrefix & "*" Then
            curId = CLng(Val(VBA.Mid(dX("Id"), Len(IdPrefix) + 1)))
            If curId > maxId Then maxId = curId
        End If
    Next Item
    
    GetNextSequenceId = IdPrefix & CStr(maxId + 1)
End Function

Friend Function GenerateTask(ByVal TaskName As String, ByVal TaskValues As Variant, Optional ByVal TaskId As String = "") As DictionaryEx
    If TaskId = "" Then TaskId = GetNextSequenceId(TaskInfo(TaskName)(DS_T_TYPE))
    Set GenerateTask = DictionaryEx.Create() _
        .Add("Id", TaskId) _
        .Add("TaskName", TaskName) _
        .Add("Values", TaskValues)
End Function

Private Sub MakeRebuildSequenceDirtyOn(ByVal Index As Long)
    Dim i As Long
    
    If pDirtyRebuildIndex > Index Then
        pDirtyRebuildIndex = IIf(Index >= 0, Index, 0)
        
        On Error Resume Next
        For i = pDirtyRebuildIndex To pRebuildSequence.Count - 1
            pTaskResults.Remove pRebuildSequence(i)("Id")
        Next i
    End If
End Sub

' Provides the dsTable object from the previous sequence to be used at the specified sequence index or .DataSource("DS1") when Index is 0.
Private Function TryGetSourceTableForSequence(ByVal Index As Long, ByRef OutTable As dsTable) As Boolean
    On Error GoTo Finally
    
    Do While True
        If Index <= 0 Then
            Set OutTable = DataSource("DS1")
            Exit Do
        Else
            With RebuildSequence(Index - 1)
                If GetReturnTypeOf(.Item("TaskName")) = "Table" Then
                    Set OutTable = pTaskResults(.Item("Id"))
                    Exit Do
                End If
            End With
            Index = Index - 1
        End If
    Loop
    
    TryGetSourceTableForSequence = Not (OutTable Is Nothing)
Finally:
End Function

' Provides the xlSheetsEx object from the previous sequence to be used at the specified sequence index.
Private Function TryGetSourceWorksheetForSequence(ByVal Index As Long, ByRef OutWorksheet As xlSheetsEx) As Boolean
    On Error GoTo Finally
    
    Do While True
        If Index <= 0 Then
            Set OutWorksheet = DataSource("XS1")
            Exit Do
        Else
            With RebuildSequence(Index - 1)
                If GetReturnTypeOf(.Item("TaskName")) = "Worksheet" Then
                    Set OutWorksheet = pTaskResults(.Item("Id"))
                    Exit Do
                End If
            End With
            Index = Index - 1
        End If
    Loop
    
    TryGetSourceWorksheetForSequence = Not (OutWorksheet Is Nothing)
Finally:
End Function

Public Function GetReturnTypeOf(ByVal TaskName As String) As String
    ' Is there any task returning something other than "Table/Worksheet"?
    ' If so, implement it in RebuildTask() and here
    Select Case Left(TaskInfo(TaskName)(DS_T_TYPE), 1)
        Case "X"
            GetReturnTypeOf = "Worksheet"
        Case Else
            GetReturnTypeOf = "Table"
    End Select
End Function

Private Sub Assign(ByRef Target, Optional ByRef Value As Variant)
    If Not IsMissing(Value) Then
        If IsObject(Value) Then
            Set Target = Value
        Else
            Target = Value
        End If
    End If
End Sub

Private Function TryRebuildTaskValues(ByRef t() As Variant, ByVal TaskParams As Variant, ByVal TaskValues As Variant) As Boolean
    On Error GoTo 0
    Dim i As Long, vItem As Variant, shouldGuessValue As Boolean
    
    If UBound(TaskParams) < 0 Then
        t = Array()
    Else
        ReDim t(0 To UBound(TaskParams))
    End If
    
    For i = 0 To UBound(TaskParams)
        shouldGuessValue = False
        If i > UBound(TaskValues) Then
            If i = 0 And TaskParams(0)(1) = "Table" Then
                Set t(i) = pTable
            ElseIf i = 0 And TaskParams(0)(1) = "Worksheet" Then
                Set t(i) = pWorksheet
            End If
            Exit For
        End If
        If IsObject(TaskValues(i)) Then
            Set vItem = dsApp.ApplyCustomVarsOn(TaskValues(i))
        Else
            If Not IsError(TaskValues(i)) Then
                vItem = dsApp.ApplyCustomVarsOn(TaskValues(i))
            Else
                vItem = TaskValues(i)
            End If
        End If

        If Not IsError(vItem) Then
            If IsArray(vItem) Then
                Select Case vItem(0)
                    Case "CURRENT"
                        Select Case vItem(1)
                            Case "Table"
                                Set t(i) = pTable
                            Case "Worksheet"
                                Set t(i) = pWorksheet
                            Case Else
                                Err.Raise 13
                        End Select
                    Case "VALUE"
                        If IsObject(vItem(1)) Then
                            Set t(i) = vItem(1)
                        Else
                            t(i) = vItem(1)
                        End If
                    Case "CALLABLE"
                        Set t(i) = Callback.Create(, vItem(1))
                    Case "REFERENCE"
                        If IsObject(pTaskResults(vItem(1))) Then
                            Set t(i) = pTaskResults(vItem(1))
                        Else
                            t(i) = pTaskResults(vItem(1))
                        End If
                    Case Else
                        shouldGuessValue = True
                End Select
            ElseIf IsObject(vItem) Then
                Set t(i) = vItem
            Else
                shouldGuessValue = True
            End If
            
            If shouldGuessValue Then
                Select Case TaskParams(i)(1)
                    Case "Table", "Worksheet"
                        Set t(i) = pTaskResults(vItem)
                    Case "Callable"
                        Set t(i) = Callback.Create(, vItem)
                    Case "Callable|Dictionary"
                        If Left(LTrim(vItem), 1) = "{" Then
                            Set t(i) = JSON.Parse(vItem, True, True)
                        Else
                            Set t(i) = Callback.Create(, vItem)
                        End If
                    Case Else
                        t(i) = vItem
                End Select
            End If
        Else
            If i = 0 And TaskParams(0)(1) = "Table" Then
                Set t(i) = pTable
            ElseIf i = 0 And TaskParams(0)(1) = "Worksheet" Then
                Set t(i) = pWorksheet
            Else
                t(i) = vItem
            End If
        End If
    Next i
    
    TryRebuildTaskValues = True
End Function



' --- GARBAGE COLLECTOR ---

Public Function AddResource(ByRef TargetResource As Object) As Object
    If Not TargetResource Is Nothing Then
        pGarbageCollector.Add CStr(ObjPtr(TargetResource)), TargetResource
    End If
    Set AddResource = TargetResource
End Function

' Run GarbageCollector
Public Sub FreeUnlinkedResources(Optional ByVal DisposeAll As Boolean = False)
    Dim i As Long, gcKeys As Variant, iMax As Long, cKey As LongPtr, isMatch As Boolean
    Dim iTR As Long, iDS As Long, trKeys As Variant, dsKeys As Variant
    
    gcKeys = pGarbageCollector.Keys()
    trKeys = pTaskResults.Keys()
    dsKeys = pDataSources.Keys()
    iMax = UBound(gcKeys)
    
    On Error GoTo StopOnError
'    Debug.Print "[DEBUG] Running GarbageCollector:"
'    Debug.Print "    TaskResults.Keys: " & JSON.Stringify(trKeys)
'    Debug.Print "    DataSources.Keys: " & JSON.Stringify(dsKeys)
'    Debug.Print "    GarbageCollector.Keys: " & JSON.Stringify(gcKeys)
    
    For i = iMax To 0 Step -1
        cKey = gcKeys(i)
        isMatch = False
        
        For iTR = 0 To UBound(trKeys)
            If IsObject(pTaskResults(trKeys(iTR))) Then
                If CStr(ObjPtr(pTaskResults(trKeys(iTR)))) = cKey Then
                    isMatch = True
                    Exit For
                End If
            End If
        Next iTR
        
        If Not isMatch Then
            For iDS = 0 To UBound(dsKeys)
                If IsObject(pDataSources(dsKeys(iDS))) Then
                    If CStr(ObjPtr(pDataSources(dsKeys(iDS)))) = cKey Then
                        isMatch = True
                        Exit For
                    End If
                End If
            Next iDS
        End If
        
        If Not isMatch Then
            If SafeDisposeResource(pGarbageCollector(gcKeys(i)), DisposeAll) Then
                On Error Resume Next
                Set pGarbageCollector(gcKeys(i)) = Nothing
                pGarbageCollector.Remove gcKeys(i)
                On Error GoTo 0
            End If
        End If
    Next i
    
    Exit Sub
StopOnError:
    Debug.Print "[ERROR] @dsTaskController.FreeUnlinkedResources() - " & Err.Description
    Resume Next
End Sub

Private Function SafeDisposeResource(ByRef TargetResource As Object, Optional ByVal DisposeAll As Boolean = False) As Boolean
    On Error GoTo DisposeFailed
    
    If Not DisposeAll And TypeOf TargetResource Is dbQuery Then Exit Function
    TargetResource.Dispose
Finally:
    SafeDisposeResource = True
    Exit Function
DisposeFailed:
    Debug.Print "[ERROR] @dsTaskController.SafeDisposeResource() - TypeName: " & TypeName(TargetResource)
    Resume Finally
End Function

'Private Sub VMemAnalysis(Optional ByVal Message As Variant)
'    Static LastVMemFree As Long, sVMemCallCount As Long
'    Dim cVMemFree As Long
'
'    cVMemFree = FileSystemLib.GetAvailableVirtualMemory()
'    sVMemCallCount = sVMemCallCount + 1
'    If LastVMemFree <> 0 Then
'        Debug.Print Printf("[VMem %4] %1 MB -> %2 MB (%3 MB) %5", LastVMemFree, cVMemFree, LastVMemFree - cVMemFree, sVMemCallCount, IIf(IsMissing(Message), "", "- " & CStr(Message)))
'    End If
'    LastVMemFree = cVMemFree
'End Sub


