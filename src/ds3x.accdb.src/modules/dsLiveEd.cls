VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "dsLiveEd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit


Public Event OnDataSourceChange()
Public Event OnSave()
Public Event OnClose()

Private pTaskController As New dsTaskController
Private pGridController As dsGridController

'Private pSourceRecordsetEx As New RecordsetEx
'Private pTargetExcelSheetEx As xlSheetsEx

Private pLiveEditor As Access.Form

Private pPresetObject As Scripting.Dictionary
Private pPresetFile As String



Public Property Get TaskController() As dsTaskController: Set TaskController = pTaskController: End Property
Public Property Get GridController() As dsGridController: Set GridController = GetGridController: End Property

Public Property Get LiveEditor() As Access.Form: Set LiveEditor = pLiveEditor: End Property
Public Property Set LiveEditor(ByRef Value As Access.Form): Set pLiveEditor = Value: End Property

Public Property Get PresetObject() As Scripting.Dictionary: Set PresetObject = pPresetObject: End Property
Public Property Get PresetFile() As String: PresetFile = pPresetFile: End Property
Public Property Let PresetFile(ByVal Value As String): pPresetFile = Value: End Property


Public Property Get Visible() As Boolean
    If pLiveEditor Is Nothing Then Exit Property
    Visible = pLiveEditor.Visible
End Property

Public Property Let Visible(ByVal Value As Boolean)
    If pLiveEditor Is Nothing Then
        If Value Then
            Dim Self As dsLiveEd
            Set Self = Me
            Set pLiveEditor = New Form_DS_LIVE_EDITOR
            Set pLiveEditor.Controller = Self
            pLiveEditor.Visible = True
            WindowAndTaskbarIconAsVisible pLiveEditor
        End If
    Else
        pLiveEditor.Visible = Value
        If Value Then WindowAndTaskbarIconAsVisible pLiveEditor
    End If
End Property

'Public Property Set SourceRecordset(ByRef rs As ADODB.Recordset)
''    pSourceRecordsetEx.Bind rs
'    Set TaskController.DataSource("DS1") = dsTable.Create(pSourceRecordsetEx).SetHeaders(pSourceRecordsetEx)
'    RaiseEvent OnDataSourceChange
'End Property



'Public Property Set SourceExcelRange(ByRef xlRange As Excel.Range)
'    With xlRange
'        Set TaskController.DataSource("DS1") = dsTable _
'            .Create(ArrayListEx.Create(.Range(.Cells(2, 1), .Cells(.Rows.Count, .Columns.Count)))) _
'            .SetHeaders(ArrayListEx.Create(.Range(.Cells(1, 1), .Cells(1, .Columns.Count)))(0))
'    End With
'    RaiseEvent OnDataSourceChange
'End Property

Public Sub AddSource(ByRef SourceObject As Object, Optional ByVal TaskId As Variant)
    Dim rX As RecordsetEx, dsT As dsTable, xlS As xlSheetsEx
    
    If TypeName(SourceObject) Like "*Recordset*" Then
        If TypeOf SourceObject Is ADODB.Recordset Then
            Set rX = RecordsetEx.Create(SourceObject.Clone)
        ElseIf TypeOf SourceObject Is RecordsetEx Then
            Set rX = SourceObject.Clone
        Else
            Err.Raise 13
        End If
        If IsMissing(TaskId) Then TaskId = "DS1"
        If Len(CStr(TaskId)) = 0 Then TaskId = "DS1"
        Set TaskController.DataSource(TaskId) = dsTable.Create(rX).SetHeaders(rX)
    Else
        If TypeOf SourceObject Is Excel.Worksheet Then
            Set xlS = xlSheetsEx.CreateFrom(SourceObject)
        ElseIf TypeOf SourceObject Is Excel.Range Then
            Set xlS = xlSheetsEx.CreateFrom(SourceObject)
        ElseIf TypeOf SourceObject Is xlSheetsEx Then
            Set xlS = SourceObject
        Else
            If TypeOf SourceObject Is dsTable Then
                Set dsT = SourceObject
                If IsMissing(TaskId) Then TaskId = "DS1"
                If Len(CStr(TaskId)) = 0 Then TaskId = "DS1"
                Set TaskController.DataSource(TaskId) = dsT
            Else
                Err.Raise 13
            End If
        End If

        If Not (xlS Is Nothing) Then
            If IsMissing(TaskId) Then TaskId = "XS1"
            If Len(CStr(TaskId)) = 0 Then TaskId = "XS1"
            Set TaskController.DataSource(TaskId) = xlS
        End If
    End If
    
    RaiseEvent OnDataSourceChange
End Sub

Public Function TryGenerateAsExcel(Optional ByRef xlSheet As xlSheetsEx = Nothing, Optional ByRef xlRange As Excel.Range = Nothing) As Boolean
    Dim vGenerated As Variant, dsT As dsTable, xlSX As xlSheetsEx, xlR As Excel.Range
    
    If TryRebuildAll Then
        
        Assign vGenerated, TaskController.GeneratedValue
        If IsObject(vGenerated) Then
            If Not vGenerated Is Nothing Then
                Select Case TypeName(vGenerated)
                    Case "dsTable"
                        If xlSheet Is Nothing Then
                            Set xlSheet = xlSheetsEx.Create()
                            Set xlRange = xlSheet.Range("A1")
                        Else
                            If xlRange Is Nothing Then Set xlRange = xlSheet.Range("A1")
                        End If
                        Set dsT = vGenerated
                        dsT.CopyToRange xlRange, True
                        ' TODO: Apply basic in-Excel post processing tasks
                        TryGenerateAsExcel = True
                        
                    Case "xlSheetsEx"
                        If xlSheet Is Nothing Then
                            Set xlSheet = vGenerated
                            Set xlRange = xlSheet.UsedRange
                        Else
                            Set xlSX = vGenerated
                            xlSX.Instance.Select
                            xlR = xlSX.UsedRange
                            xlR.Copy
                            xlSheet.Instance.Select
                            If xlRange Is Nothing Then
                                xlRange = xlSheet.Range("A1")
                            End If
                            xlRange.Select
                            xlSheet.Instance.Paste
                        End If
                        TryGenerateAsExcel = True
                        
                    Case Else
                        Debug.Print "[ERROR] @dsLiveEd.TryGenerate"
                End Select
            End If
        End If
    End If
End Function

Public Function TryRebuildAll() As Boolean
    On Error GoTo Finally
    With TaskController
        .SequenceIndex = .RebuildSequence.Count - 1
        TryRebuildAll = (.DirtyRebuildIndex = .RebuildSequence.Count)
    End With
Finally:
End Function



Public Sub Dispose()
    Destroy
End Sub

Private Sub Class_Initialize()
'    Set pSourceRecordsetEx = Nothing
    Set pGridController = Nothing
    Set pTaskController = Nothing
    ' Set GridController.TaskController = TaskController
End Sub

Private Sub Class_Terminate()
    Debug.Print "[INFO] dsLiveEd.Terminate()"
    Destroy
End Sub

Private Function Destroy()
    On Error Resume Next
    Visible = False
    Set pPresetObject = Nothing
    pPresetFile = ""
'    pSourceRecordsetEx.Unbind
    If Not pGridController Is Nothing Then Set pGridController.TaskController = Nothing
    Set pGridController = Nothing
    pTaskController.Dispose
    Set pTaskController = Nothing
    Set pLiveEditor = Nothing
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

Private Function GetGridController() As dsGridController
    If pGridController Is Nothing Then
        Set pGridController = New dsGridController
        Set pGridController.TaskController = TaskController
    End If
    Set GetGridController = pGridController
End Function


' --- PRESETS ---

' Updates PresetObject dictionary with the current state in the editor, if that state derives from
' a previously imported preset by providing a dictionary object while importing it, that dictionary
' will be updated accordingly while keeping any other values unrelated to dsLiveEd.
' If it was imported from a json file, it will also update that file unless Target specifies another
' path, which in that case, it will be saved to the specified file if Target points to a full path to
' a file, if it points to a directory instead, it will show the file save dialog on that directory.
' If Target is not specified or invalid and current state doesn't derive from a previously imported
' file, it will not be saved to disk.
'
' Triggers the OnSave event before it returns.
'
' Where Target can be:
'   - Missing (No parameter provided)
'   - A full path to a JSON file or directory
' Returns whether save succeeded or not
Public Function SavePreset(Optional ByVal Target As Variant) As Boolean
    Dim TargetPath As String
    
    If Not UpdatePresetObject Then Exit Function
    
    If IsMissing(Target) Then Target = pPresetFile
    If Trim(Target) <> "" Then
        If TryWaitFolderExists(Target, 0) Then
            TargetPath = Target
            If Not TryWriteAccessOfSaveAsDialog(TargetPath, "*.json", "Save As") Then Exit Function
            Target = TargetPath
        End If
        If Not TryWaitFileWriteAccess(Target, 20) Then Exit Function
        If Not TryWriteTextToFile(Target, JSON.Stringify(pPresetObject), True) Then Exit Function
        pPresetFile = Target
    End If
    
    RaiseEvent OnSave
    SavePreset = True
End Function

' Where Target can be:
'   - Missing (No parameter provided)   -> Shows an open file dialog to select a *.json file
'   - A full path to a JSON file        -> Loads the preset in that file (unicode), if it points to a directory instead, shows the open file dialog on that path
'   - A JSON string                     -> Parses that JSON and loads it as preset
'   - A Dictionary/DictionaryEx object  -> The preset itself
' Returns true if preset import was a success, false otherwise
Public Function ImportPreset(Optional ByVal Target As Variant) As Boolean
    Dim TargetPath As String, Aux As String
    If IsMissing(Target) Then Target = VBA.Environ$("USERPROFILE")
    
    If Not IsObject(Target) Then
        If Not Left(LTrim(Target), 1) = "{" Then
            If TryWaitFolderExists(Target, 0) Then
                If Not TryFileOpenDialog(TargetPath, Target) Then Exit Function
                Target = TargetPath
            End If
            If Not TryWaitFileExists(Target, 10) Then Exit Function
            TargetPath = "" & Target
            If Not TryReadAllTextInFile(Target, Aux) Then Exit Function
            Target = Aux
        End If
        If Left(LTrim(Target), 1) = "{" Then
            Set Target = JSON.Parse(Target, True)
        End If
    End If
    If IsObject(Target) Then
        If Target Is Nothing Then Set Target = DictionaryEx.Create().Instance
        If TypeOf Target Is DictionaryEx Then Set Target = Target.Instance
        If TypeOf Target Is Scripting.Dictionary Then
            ImportPreset = SetPresetObject(Target)
        End If
    End If
    
    If ImportPreset Then
        pPresetFile = TargetPath
    End If
End Function

Private Function SetPresetObject(ByVal Target As Scripting.Dictionary) As Boolean
    On Error GoTo Finally
    Dim dX As New DictionaryEx, rollbackDenied As Boolean, seqList As ArrayListEx, dEntry As Scripting.Dictionary, restoreGridController As Boolean
    If Target Is Nothing Then Set Target = DictionaryEx.Create().Instance
    
    dX.Bind Target
    dX.Add "dsLiveEd", JSON.Parse(JSON.Stringify(dX.GetValue("dsLiveEd", GetDefaultDsLiveEdPreset)), True, True)
    
    Set dEntry = GetNamedEntryFromEntries(dX)
    Set seqList = ArrayListEx.Create(JSON.Parse(JSON.Stringify(dEntry("RebuildSequence")), True, True))
    
    ' Any errors that occur after this point will prevent restoring to the previous state, resetting the editor state
    rollbackDenied = True
    
    Set TaskController.RebuildSequence = seqList
    ' TODO (no other config yet): load config values from preset
    
    Set pPresetObject = dX.Instance
    SetPresetObject = True
    
    Exit Function
Finally:
    If rollbackDenied Then
        rollbackDenied = Visible
        restoreGridController = Not (pGridController Is Nothing)
        Destroy
        If restoreGridController Then
            Set GridController.TaskController = TaskController
        End If
        If rollbackDenied Then Visible = True
    End If
End Function

Private Function UpdatePresetObject() As Boolean
    Dim dX As New DictionaryEx, Entry As New DictionaryEx
    If pPresetObject Is Nothing Then Set pPresetObject = DictionaryEx.Create().Instance
    
    dX.Bind pPresetObject
    
    dX.Add "dsLiveEd", JSON.Parse(JSON.Stringify(dX.GetValue("dsLiveEd", GetDefaultDsLiveEdPreset)), True, True)
    
    Entry.Bind GetNamedEntryFromEntries(dX)
    If Entry.Instance Is Nothing Then Entry.Bind DictionaryEx.Create().Instance
    
    Entry.Add "RebuildSequence", TaskController.RebuildSequence
    ' TODO: Error handling / rollback
    If Not SetNamedEntryToEntries(dX, Entry) Then Exit Function
    
    dX.Add "dsLiveEd", JSON.Parse(JSON.Stringify(dX("dsLiveEd")), True, True)
    UpdatePresetObject = True
End Function

Private Function GetDefaultDsLiveEdPreset() As Scripting.Dictionary
    Set GetDefaultDsLiveEdPreset = JSON.Parse("{""Config"": {""EntryPoint"": ""default""}, ""Entries"": [{""Name"": ""default"", ""RebuildSequence"": []}]}", True, True)
End Function

Private Function GetNamedEntryFromEntries(ByRef dX As DictionaryEx, Optional ByVal EntryName As Variant) As Scripting.Dictionary
    Dim i As Long, eList As ArrayListEx
    
    If IsMissing(EntryName) Then EntryName = dX("dsLiveEd")("Config")("EntryPoint")
    Set eList = ArrayListEx.Create(dX("dsLiveEd")("Entries"))
    
    For i = 0 To eList.Count - 1
        If eList(i)("Name") = EntryName Then
            Set GetNamedEntryFromEntries = eList(i)
            Exit Function
        End If
    Next i
End Function

' If EntryName is provided, sets it to Entry.Name regardless of it having a different value or not existing at all. Otherwise, the
' current Entry.Name value is kept when exists or it defaults to config's EntryPoint.
Private Function SetNamedEntryToEntries(ByRef dX As DictionaryEx, ByRef Entry As DictionaryEx, Optional ByVal EntryName As Variant) As Boolean
    On Error GoTo Finally
    Dim i As Long, eList As ArrayListEx, isMatch As Boolean
    
    If IsMissing(EntryName) Then EntryName = Entry.GetValue("Name", dX("dsLiveEd")("Config")("EntryPoint"))
    Entry.Add "Name", EntryName
    
    Set eList = ArrayListEx.Create(dX("dsLiveEd")("Entries"))
    
    For i = 0 To eList.Count - 1
        If eList(i)("Name") = EntryName Then
            Set eList(i) = Entry.Instance
            isMatch = True
        End If
    Next i
    
    If Not isMatch Then
        eList.Add Entry.Instance
    End If
    
    DictionaryEx.Bind(dX("dsLiveEd")).Add("Entries", eList).Unbind
    
    SetNamedEntryToEntries = True
Finally:
End Function

Public Sub TriggerLiveEditorWindowClose(Optional ByVal SkipAutoClose As Boolean = False)
    On Error Resume Next
    
    If Not SkipAutoClose Then
        Me.Visible = False
        Set pLiveEditor = Nothing
    End If
    
    RaiseEvent OnClose
End Sub



' --- SAMPLE MOCKS FOR DEVELOPMENT ---


Public Sub LoadSampleValues()
    ' SELECT U.UR, E.DEMANDA, E.CONTRATO, E.TAREA_PROCESO, E.FECHA_ENTRADA, E.FECHA_ULTIMO_HITO, E.INFORME_FISCAL, U.DIRECCION, U.CODIGO_POSTAL, U.PROVINCIA, U.TIPO_INMUEBLE, U.VALOR_TASACION, U.PRINCIPAL, U.TOTAL_DEUDA, U.VALOR_SUBASTA, U.PORCENTAJE_PROPIEDAD, U.FECHA_PETICION_CALCULOS, U.TIPO, E.OBSERVACIONES
    AddSource _
        dbQuery.Create( _
            "SELECT U.UR, E.DEMANDA, E.FECHA_ENTRADA, U.PRINCIPAL, U.TIPO, E.CONTRATO, E.TAREA_PROCESO, E.FECHA_ULTIMO_HITO, U.TOTAL_DEUDA " & _
            "FROM [dbo].[T_ATM_ISCS_URS_LOTE] UL INNER JOIN [dbo].[T_ATM_ISCS_EXPEDIENTES] E ON E.DEMANDA = UL.DEMANDA INNER JOIN [dbo].[T_ATM_ISCS_URS] U ON U.UR = UL.UR " & _
            "WHERE e.OBSERVACIONES Is Not Null And U.DIRECCION Is Not Null And U.PRINCIPAL Is Not Null " _
        )
End Sub

Public Sub LoadADifferentSetOfSampleValues()
    AddSource dbQuery.Create("EXEC [dbo].[P_ATM_ISCS_APP_USER_MONTHLY_TIME]")
End Sub

Public Sub MockAddTasks()
    Dim aX As ArrayListEx
    
    Set aX = ArrayListEx.Create(TaskController.RebuildSequence)
    With aX
        .Add DictionaryEx.Create() _
                            .Add("Id", "DS1") _
                            .Add("TaskName", "DataSource") _
                            .Add("Values", Array())
                            
'        .Add DictionaryEx.Create() _
'                            .Add("Id", "T1") _
'                            .Add("TaskName", "AddRowNumbers") _
'                            .Add("Values", Array( _
'                                Array("CURRENT", "Table"), _
'                                Array("VALUE", 13)))
'        .Add DictionaryEx.Create() _
'                            .Add("Id", "T2") _
'                            .Add("TaskName", "SetColumnName") _
'                            .Add("RETURNS", "Table") _
'                            .Add("Values", Array( _
'                                Array("CURRENT", "Table"), _
'                                Array("VALUE", 3), _
'                                Array("VALUE", "NEW_COLUMN_NAME")))
'        .Add DictionaryEx.Create() _
'                            .Add("Id", "T3") _
'                            .Add("TaskName", "SetColumnName") _
'                            .Add("RETURNS", "Table") _
'                            .Add("Values", Array( _
'                                Array("CURRENT", "Table"), _
'                                Array("VALUE", 5), _
'                                Array("VALUE", "YET_ANOTHER_NEW_COLUMN_NAME")))
'        .Add DictionaryEx.Create() _
'                            .Add("Id", "T4") _
'                            .Add("TaskName", "Filter") _
'                            .Add("RETURNS", "Table") _
'                            .Add("Values", Array( _
'                                Array("CURRENT", "Table"), _
'                                Array("VALUE", "[TAREA_PROCESO] LIKE '35*'")))

        ' GetRange(ByRef TargetTable As Variant, Optional ByVal Index As Variant, Optional ByVal GetCount As Variant, Optional ByVal ColumnIndexes As Variant) As dsTable
        
        .Add DictionaryEx.Create() _
                            .Add("Id", "T2") _
                            .Add("TaskName", "GetRange") _
                            .Add("Values", Array(, , , Array(2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32)))
        
'        .Add DictionaryEx.Create() _
'                            .Add("Id", "T2") _
'                            .Add("TaskName", "GetRange") _
'                            .Add("Values", Array( _
'                                Array("CURRENT", "Table"), _
'                                , _
'                                , _
'                                Array("VALUE", Array(3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33))))

        .Add DictionaryEx.Create() _
                            .Add("Id", "T3") _
                            .Add("TaskName", "ReplaceValues") _
                            .Add("Values", Array(, "SecondsToHMS"))


        .Add DictionaryEx.Create() _
                            .Add("Id", "T4") _
                            .Add("TaskName", "GetRange") _
                            .Add("Values", Array("DS1", , , Array(0, 1)))
        
'        .Add DictionaryEx.Create() _
'                            .Add("Id", "T4") _
'                            .Add("TaskName", "GetRange") _
'                            .Add("Values", Array( _
'                                Array("REFERENCE", "T1"), _
'                                , _
'                                , _
'                                Array("VALUE", Array(0, 1, 2))))
        
        .Add DictionaryEx.Create() _
                            .Add("Id", "T5") _
                            .Add("TaskName", "Join") _
                            .Add("Values", Array(, "T3"))
                                
'        .Add DictionaryEx.Create() _
'                            .Add("Id", "T5") _
'                            .Add("TaskName", "Join") _
'                            .Add("Values", Array( _
'                                Array("CURRENT", "Table"), _
'                                Array("REFERENCE", "T3")))
        
        .Add DictionaryEx.Create() _
                            .Add("Id", "T6") _
                            .Add("TaskName", "ReplaceColumnNames") _
                            .Add("Values", Array( _
                                , _
                                "XDaysAgo2ColumnName"))

        ' XDaysAgo2ColumnName
'        .Add DictionaryEx.Create() _
'                            .Add("Id", "T6") _
'                            .Add("TaskName", "ReplaceColumnNames") _
'                            .Add("Values", Array( _
'                                Array("CURRENT", "Table"), _
'                                Array("CALLABLE", "XDaysAgo2ColumnName")))

'        .Add DictionaryEx.Create() _
'                            .Add("Id", "T7") _
'                            .Add("TaskName", "SetColumnName") _
'                            .Add("RETURNS", "Table") _
'                            .Add("Values", Array( _
'                                Array("CURRENT", "Table"), _
'                                Array("VALUE", 3), _
'                                Array("VALUE", "NEW_COLUMN_NAME")))

    End With
    
    Set TaskController.RebuildSequence = aX
End Sub







