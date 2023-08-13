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

Private pLiveEditor As Access.Form

Private pPresetObject As Scripting.Dictionary
Private pPresetFile As String

Private pDisposeOnClose As Boolean



Public Property Get TaskController() As dsTaskController: Set TaskController = pTaskController: End Property
Public Property Get GridController() As dsGridController: Set GridController = GetGridController: End Property

Public Property Get LiveEditor() As Access.Form: Set LiveEditor = pLiveEditor: End Property
Public Property Set LiveEditor(ByRef Value As Access.Form): Set pLiveEditor = Value: End Property

Public Property Get PresetObject() As Scripting.Dictionary: Set PresetObject = pPresetObject: End Property
Public Property Get PresetFile() As String: PresetFile = pPresetFile: End Property
Public Property Let PresetFile(ByVal Value As String): pPresetFile = Value: End Property

Public Property Get DisposeOnClose() As Boolean: DisposeOnClose = pDisposeOnClose: End Property
Public Property Let DisposeOnClose(ByVal Value As Boolean): pDisposeOnClose = Value: End Property

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
                            Set xlR = xlSX.UsedRange
                            xlR.Copy
                            xlSheet.Instance.Select
                            If xlRange Is Nothing Then
                                Set xlRange = xlSheet.Range("A1")
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

Public Function TryApplyPresetOn(ByVal Preset As Variant, ByRef SourceInput As Variant, ByRef GeneratedOutput As Variant, Optional ByVal AutoHeaders As Boolean = True) As Boolean
    On Error GoTo Fallback
    
    ImportPreset Preset
    TryApplyPresetOn = TryApplyOn(SourceInput, GeneratedOutput, AutoHeaders)
    
Finally:
    Exit Function
Fallback:
    Debug.Print "[ERROR] @dsLiveEd.TryApplyPresetOn() - " & Err.Description
    Resume Finally
End Function

Public Function TryApplyOn(ByRef SourceInput As Variant, ByRef GeneratedOutput As Variant, Optional ByVal AutoHeaders As Boolean = True) As Boolean
    Dim srcType As String, inType As String, ds1Idx As Long, xs1Idx As Long, xlSheet As xlSheetsEx
    On Error GoTo Fallback
    
    srcType = TypeName(SourceInput)
    If Not TaskController.TryGetIndexOfTaskId("DS1", ds1Idx) Then ds1Idx = 999
    If Not TaskController.TryGetIndexOfTaskId("XS1", xs1Idx) Then xs1Idx = 999
    inType = IIf(xs1Idx < ds1Idx, "Worksheet", "Table")
    
    Select Case inType
        Case "Table"
            If srcType Like "*Recordset*" Then
                AddSource C2Object(SourceInput)
            ElseIf srcType Like "*Table*" Then
                AddSource C2Object(SourceInput)
            Else
                AddSource dsTable.Create(SourceInput, AutoHeaders)
            End If
        Case "Worksheet"
            If IsObject(SourceInput) Then
                AddSource C2Object(SourceInput)
            Else
                If VarType(SourceInput) = vbString Then
                    If TryWaitFileExists(SourceInput, 10) Then
                        Set xlSheet = xlSheetsEx.Create(, SourceInput)
                        AddSource xlSheet
                    Else
                        Err.Raise 17
                    End If
                Else
                    Err.Raise 17
                End If
            End If
        Case Else
            Err.Raise 17
    End Select
    
    TryApplyOn = TryApply(GeneratedOutput)
    
Finally:
    Exit Function
Fallback:
    Debug.Print "[ERROR] @dsLiveEd.TryApplyOn() - " & Err.Description
    Resume Finally
End Function

' Returns wether current dsLiveEd setup successfully applied or not. While GeneratedOutput references to the final object being generated.
Public Function TryApply(ByRef GeneratedOutput As Variant) As Boolean
    On Error GoTo Fallback
    
    If TryRebuildAll Then
        Assign GeneratedOutput, TaskController.GeneratedValue
        If IsObject(GeneratedOutput) Then
            If Not GeneratedOutput Is Nothing Then
                TryApply = True
            End If
        End If
    End If
    
Finally:
    Exit Function
Fallback:
    Debug.Print "[ERROR] @dsLiveEd.TryApply() - " & Err.Description
    Resume Finally
End Function



Public Sub Dispose()
    Destroy
End Sub

Private Function C2Object(ByRef Value As Variant) As Object: Set C2Object = Value: End Function


Private Sub Class_Initialize()
'    Set pSourceRecordsetEx = Nothing
    Set pGridController = Nothing
    Set pTaskController = Nothing
    ' Set GridController.TaskController = TaskController
End Sub

Private Sub Class_Terminate()
    Debug.Print "[INFO] dsLiveEd.Terminate()"
'    On Error Resume Next
'    If Not pTaskController Is Nothing Then
'        Destroy
'    End If
    Destroy
End Sub

Private Function Destroy()
    On Error Resume Next
    Visible = False
    Set pPresetObject = Nothing
    pPresetFile = ""
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

Public Function GetBlankPreset() As Scripting.Dictionary
    Dim dX As DictionaryEx
    Set dX = DictionaryEx.Create()
    Set dX("dsLiveEd") = GetDefaultDsLiveEdPreset()
    Set GetBlankPreset = dX.Instance
End Function

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

' ---


Friend Sub TriggerLiveEditorWindowClose(Optional ByVal SkipAutoClose As Boolean = False)
    On Error Resume Next
    
    If Not SkipAutoClose Then
        Me.Visible = False
        Set pLiveEditor = Nothing
    End If
    
    RaiseEvent OnClose
End Sub


