Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "ds3x.UI.LiveEditor"
Option Compare Database
Option Explicit

Private WithEvents pController As dsGridController
Attribute pController.VB_VarHelpID = -1
Private WithEvents pContinuousList As Form_DS_RICHTEXT_LIST
Attribute pContinuousList.VB_VarHelpID = -1
Private WithEvents pCVarsScrollview As Form_DS_SCROLLVIEW
Attribute pCVarsScrollview.VB_VarHelpID = -1

Private pTasks As DictionaryEx
Private pSelectedTask As String
Private pTaskNames As ArrayListEx
Private pTaskGroups As DictionaryEx
Private pActiveTaskName As String
Private pActiveTaskIndex As Long
Private pEditMode As Boolean
Private pEnableAllTasks As Boolean

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
'        "T"            /* TaskType */
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




Property Get IsSubform() As Boolean
    On Error Resume Next
    IsSubform = Len(Me.Parent.Name) > 0
    On Error GoTo 0
End Property

Public Property Get ActiveTaskName() As String: ActiveTaskName = pActiveTaskName: End Property
Public Property Let ActiveTaskName(ByVal Value As String): pActiveTaskName = Value: End Property
Public Property Get ActiveTaskIndex() As Long: ActiveTaskIndex = pActiveTaskIndex: End Property
Public Property Let ActiveTaskIndex(ByVal Value As Long): pActiveTaskIndex = Value: End Property
Public Property Get EditMode() As Boolean: EditMode = pEditMode: End Property
Public Property Let EditMode(ByVal Value As Boolean): pEditMode = Value: End Property

'Private pActiveTaskIndex As Long
'Private pEditMode As Boolean

Public Property Get Controller() As dsGridController
    Set Controller = pController
End Property

Public Property Set Controller(ByRef Value As dsGridController)
    Set pController = Value
    pSelectedTask = ""
    RebuildTasksList
End Property



Private Sub DS_ADD_TASK_BUTTON_Click()
    If pSelectedTask <> "" Then
        AddTaskUsingCurrentValuesAs pSelectedTask
    End If
End Sub

Private Sub DS_CANCEL_BUTTON_Click()
    DoCmd.Close acForm, "DS_CREATE_TASK", acSaveNo
End Sub

Public Sub RebuildTasksList()
    Dim TaskName As Variant, i As Long, TaskType As String, TaskList As New ArrayListEx, ActiveTaskType As String
    Dim TListType As Variant, dsT As dsTable, rsX As RecordsetEx, RequiresPattern As String, ValidRequiredType As Boolean

    Set pTasks = pController.TaskController.AvailableTasks.Tasks
    Set pTaskNames = ArrayListEx.Create(pTasks.Keys)
    Set pTaskGroups = DictionaryEx.Create()
    If pActiveTaskName <> vbNullString Then
        ActiveTaskType = pTasks(pActiveTaskName)(DS_T_TYPE)
    End If

    For i = 0 To pTaskNames.Count - 1
        TaskName = pTaskNames(i)
        TaskType = pTasks(TaskName)(DS_T_TYPE)
        If TaskType = "XS" Then TaskType = "DS"
        RequiresPattern = pTasks(TaskName)(DS_T_REQUIRES)
        If Not pTaskGroups.Exists(TaskType) Then
            pTaskGroups.Add TaskType, ArrayListEx.Create()
        End If
        ValidRequiredType = pEnableAllTasks Or pEditMode Or (ActiveTaskType Like RequiresPattern)
        pTaskGroups(TaskType).Add Array(i, TaskName, IIf(ValidRequiredType, 1, 4))
    Next i

    For Each TListType In pTaskGroups.Keys
        TaskList.Add GetTaskTypeHeader(TListType)
        TaskList.AddRange pTaskGroups(TListType)
        TaskList.Add GetTaskTypeHeader("")
    Next TListType

    Set dsT = dsTable.Create(TaskList).SetHeaders(Array("DS_LIST_ITEM_INDEX", "DS_LIST_ITEM", "DS_LIST_ITEM_STATE"))
    Set rsX = RecordsetEx.Create(dsT.NamedRecordset)
    Set Me.DS_CREATE_TASK_LIST.Form.Recordset = rsX.Instance
    If pContinuousList Is Nothing Then
        Set pContinuousList = Me.DS_CREATE_TASK_LIST.Form
        pContinuousList.SetActiveIndex pTaskNames.IndexOf(IIf(pActiveTaskName = "", "DataSource", pActiveTaskName))
    End If
End Sub

Private Function GetTaskTypeHeader(ByVal TaskType As String) As Variant
    Static IndexCountdown As Long
    Dim Title As String
    IndexCountdown = IndexCountdown - 1

    Select Case TaskType
        Case "XL": Title = "XL FORMATTING"
        Case "XLG": Title = "XL GENERATION"
        Case "DS": Title = "DATA SOURCES"
        Case "T": Title = "TRANSFORMATION"
        Case "DSG": Title = "GENERATION"
        Case "DBG": Title = "DEV / DEBUG"
        Case Else
            GetTaskTypeHeader = Array(IndexCountdown, " ", 4)
            Exit Function
    End Select

    GetTaskTypeHeader = Array(IndexCountdown, CreateListItemTextHeader(Title), 4)
End Function

Private Function CreateListItemTextHeader(ByVal Title As String) As String
    CreateListItemTextHeader = _
        "<div align=center><font face=""Consolas"" size=1 color=""#A5A5A5""><b>" & Title & "</b></font><font color=white size=2>&nbsp;&nbsp;.<br>" & _
        "</font><font style=""BACKGROUND-COLOR:#808080"">______ _______ _________ ________ ________ _______ ____ ______ _____</font></div>"
End Function

Private Sub DS_ENABLE_ALL_TASKS_BUTTON_Click()
    pEnableAllTasks = Not pEnableAllTasks
    RebuildTasksList
End Sub

Private Sub DS_SHOW_CUSTOM_VARS_BUTTON_Click()
    ShowAvailableCustomVars
End Sub

Private Sub DS_TASK_PARAM_0_Change()
    Dim Target As String

    If GetControlText(Me.DS_TASK_PARAM_0) = "< Select... >" Then
        If FileSystemLib.TryFileOpenDialog(Target) Then
            SetControlText Me.DS_TASK_PARAM_0, Target
        Else
            SetControlText Me.DS_TASK_PARAM_0, ""
        End If
    End If
End Sub

Private Sub DS_TASK_PARAM_1_Change()
    Dim Target As String, sExtension As String

    If GetControlText(Me.DS_TASK_PARAM_1) = "< Select... >" Then
        If pSelectedTask Like "*JSON*" Then
            sExtension = "*.json"
        ElseIf pSelectedTask Like "*CSV*" Then
            sExtension = "*.csv"
        Else
            sExtension = "*.xlsx"
        End If
        If FileSystemLib.TrySaveAsDialog(Target, sExtension) Then
            SetControlText Me.DS_TASK_PARAM_1, Target
        Else
            SetControlText Me.DS_TASK_PARAM_1, ""
        End If
    End If
End Sub

Private Sub Form_Load()
    ScreenLib.ScreenLib_Resync
    ScreenLib.WindowSizeTo Me, 12300, 6500
    ScreenLib.WindowAlwaysOnTop Me
    ScreenLib.SetControlAsEnabled Me.DS_ADD_TASK_BUTTON, False
End Sub

Private Sub Form_Resize()
    Me.DS_FOOTER_BG.HorizontalAnchor = acHorizontalAnchorBoth
    Me.DS_TASK_DEFINITION_SYNTAX.HorizontalAnchor = acHorizontalAnchorBoth
    Me.DS_TASK_TEXT_CONTENT.HorizontalAnchor = acHorizontalAnchorBoth
    Me.DS_TASK_TEXT_CONTENT.Height = (Me.InsideHeight - Me.DS_CREATE_TASK_LIST.Top - Me.DS_FOOTER_BG.Height - Me.DS_TASK_TEXT_CONTENT.Top)
End Sub

Private Function GetReturnTypeOf(ByVal TaskName As String) As String
    GetReturnTypeOf = pController.TaskController.GetReturnTypeOf(TaskName)
End Function

Private Sub RebuildUIForTask(ByVal TaskName As String)
    Dim Task As Variant, textContent As String, TaskParams As Variant, i As Long
    On Error GoTo Finally

    ResetAllTaskParamTextFields
    Task = pTasks(TaskName)

    Me.DS_TASK_DEFINITION_SYNTAX = Task(DS_T_NAME) & " (" & Task(DS_T_PARAMS)(0) & "): " & GetReturnTypeOf(TaskName)
    Me.DS_TASK_DEFINITION_SYNTAX.Visible = True

    textContent = Task(DS_T_DESC)
    If Task(DS_T_USAGE) <> "" Then
        If textContent <> "" Then textContent = textContent & vbNewLine & " " & vbNewLine
        textContent = textContent & "<strong>Usage</strong>: " & Task(DS_T_USAGE)
    End If

    Me.DS_TASK_TEXT_CONTENT = Replace(textContent, vbNewLine, "<br>")

    TaskParams = Task(DS_T_PARAMS)(2)
    For i = 0 To UBound(TaskParams)
        RefillDefaultParamValues TaskName, i, TaskParams
        If EditMode And TaskName = ActiveTaskName Then
            RefillExistingParamValues TaskName, i, TaskParams, ActiveTaskIndex
        End If
    Next i
    MoveTaskTextContentToFit UBound(TaskParams) + 1

    For i = UBound(TaskParams) + 1 To 4
        Me.Controls("DS_TASK_PARAM_" & CStr(i)).Visible = False
        Me.Controls("DS_LABEL_TASK_PARAM_" & CStr(i)).Visible = False
    Next i
Finally:
    Me.DS_ADD_TASK_BUTTON.Caption = IIf(EditMode, "EDIT TASK", "ADD TASK")
    ScreenLib.SetControlAsEnabled Me.DS_ADD_TASK_BUTTON, True
End Sub

Private Sub MoveTaskTextContentToFit(ByVal NumParams As Long)
    With Me.DS_TASK_TEXT_CONTENT
        .Height = 270
        .Top = 3000 - ((5 - NumParams) * 450)
        .Height = (Me.InsideHeight - Me.DS_CREATE_TASK_LIST.Top - Me.DS_FOOTER_BG.Height - .Top)
    End With
End Sub

Private Sub ResetAllTaskParamTextFields()
    Dim i As Long
    On Error Resume Next

    For i = 0 To 4
        SetControlText Me.Controls("DS_TASK_PARAM_" & CStr(i)), ""
    Next i
    On Error GoTo 0
End Sub

Private Sub RefillExistingParamValues(ByVal TaskName As String, ByVal ParamIndex As Long, ByVal TaskParams As Variant, ByVal TaskIndex As Long)
    On Error GoTo 0
    Dim TaskValues As Variant, v As Variant, MatchingValue As String

    TaskValues = pController.TaskController.RebuildSequence(TaskIndex)("Values")
    If IsArray(TaskValues) Then
        If ParamIndex <= UBound(TaskValues) Then
            Assign v, TaskValues(ParamIndex)
            If Not IsObject(v) Then
                If Not (IsNull(v) Or IsError(v) Or IsEmpty(v)) Then
                    If IsArray(v) Then
                        MatchingValue = JSON.Stringify(v)
                    Else
                        Select Case VarType(v)
                            Case vbBoolean: MatchingValue = IIf(v, "true", "false")
                            Case Else: MatchingValue = CStr(v)
                        End Select
                    End If
                End If
            Else
                If TypeOf v Is Scripting.Dictionary Then
                    MatchingValue = JSON.Stringify(v)
                Else
                    Err.Raise 17
                End If
            End If
            If MatchingValue <> "" Then
                Me.Controls("DS_TASK_PARAM_" & CStr(ParamIndex)) = MatchingValue
            End If
        End If
    End If
End Sub

Private Sub RefillDefaultParamValues(ByVal TaskName As String, ByVal ParamIndex As Long, ByVal TaskParams As Variant)
    Dim Item As Variant, isRequired As Boolean, ParamName As String

    With Me.Controls("DS_TASK_PARAM_" & CStr(ParamIndex))
        .RowSourceType = "Value List"
        .RowSource = vbNullString
        .ColumnCount = 1
        .Visible = True
        ParamName = TaskParams(ParamIndex)(0)
        isRequired = Not (VBA.Mid$(ParamName, 1, 1) = "[")

        '"SetNumberFormat", "[Source]: Table, [ColumnIndexes]: Long|Array(), [NumberFormat]: String", _

        If ParamName Like "*ColumnIndexes*" Then
            If pController.SelectedColumnIndexes.Count >= 1 Then
                .Value = JSON.Stringify(pController.SelectedColumnIndexes)
            Else
                .Value = ""
            End If
        ElseIf ParamName Like "*ColumnIndex*" Then
            If pController.SelectedColumnIndexes.Count >= 1 Then
                .Value = CStr(pController.SelectedColumnIndexes(0))
            Else
                .Value = ""
            End If
        ElseIf ParamName Like "*TargetFile*" Then
            .AddItem ""
            .AddItem "< Select... >"
        ElseIf ParamName Like "*NumberFormat*" Then
            AddDefaultNumberFormatItems Me.Controls("DS_TASK_PARAM_" & CStr(ParamIndex))
            .Value = GetSelectedColumnNumberFormat()
        Else
            ' Select by PARAM TYPE
            Select Case TaskParams(ParamIndex)(1)
                Case "Table", "Worksheet"
                    If Not isRequired Then .AddItem ""
                    For Each Item In GetListOfPreviousTaskResults(TaskParams(ParamIndex)(1)).Items()
                        .AddItem CStr(Item)
                    Next Item
                Case "String|Dictionary"
                    ' Being PARAM TYPE "String|Dictionary", select by TASK NAME
                    Select Case TaskName
                        Case "LoadFromExcelFile", "OpenWorksheetFromFile"
                            Item = "{""UpdateLinks"": false, ""ReadOnly"": true, ""Local"": true, ""UTF8"": false, ""NoTextQualifier"": true}"
                            .AddItem ""
                            .AddItem Item
                            .Value = Item
                        Case "LoadFromFileAsCSV"
                            Item = "{""Delimiter"": ""Auto"", ""InLocalFormat"": false, ""AutoHeaders"": true, ""TextQualifier"": ""\""""}"
                            .AddItem ""
                            .AddItem Item
                            .Value = Item
                        Case Else
                            ' ...
                    End Select
                Case "String"
                    ' Being PARAM TYPE "String", select by PARAM NAME
                    Select Case TaskParams(ParamIndex)(0)
                        Case "TableStyle"
                            For Each Item In GetExcelTableStyles(): .AddItem CStr(Item): Next Item
                            .Value = "TableStyleMedium6"
                        Case "[ConnectionString]", "ConnectionString"
                            Item = dsQuery.ConnectionString
                            .AddItem ""
                            .AddItem CStr(Item)
                            .Value = CStr(Item)
                        Case Else
                            ' ...
                    End Select
                Case "Boolean"
                    .AddItem "true"
                    .AddItem "false"
                    If Not isRequired Then
                        On Error Resume Next
                        Item = LCase(Trim(Split(TaskParams(ParamIndex)(0), "=")(1)))
                        Item = Left(Item, Len(Item) - 1)
                        .Value = CStr(Item)
                        On Error GoTo 0
                    End If
                Case Else
                    ' ...
            End Select
        End If
    End With
    With Me.Controls("DS_LABEL_TASK_PARAM_" & CStr(ParamIndex))
        .Visible = True
        .Caption = TaskParams(ParamIndex)(0)
    End With

End Sub

Private Function GetListOfPreviousTaskResults(Optional ByVal TaskReturnType As String = "") As ArrayListEx
    Dim t As New ArrayListEx, i As Long, aList As ArrayListEx

    On Error GoTo Finally
    With pController.TaskController
        Set aList = .RebuildSequence

        For i = 0 To .SequenceIndex
            If TaskReturnType = "" Then
                t.Add aList(i)("Id")
            Else
                If .GetReturnTypeOf(aList(i)("TaskName")) = TaskReturnType Then
                    t.Add aList(i)("Id")
                End If
            End If
        Next i
    End With
Finally:
    Set GetListOfPreviousTaskResults = t
End Function

Private Sub AddTaskUsingCurrentValuesAs(ByVal TaskName As String)
    Dim Task As Variant, t() As Variant, TaskParams As Variant, i As Long, Item As Variant, TaskId As String
    Task = pTasks(TaskName)
    TaskParams = Task(DS_T_PARAMS)(2)

    If UBound(TaskParams) < 0 Then
        t = Array()
    Else
        ReDim t(0 To UBound(TaskParams))

        For i = 0 To UBound(TaskParams)
            If i < 5 Then
                If Nz(Me.Controls("DS_TASK_PARAM_" & CStr(i)), "") <> "" Then
                    Assign Item, ParseParamValue(Nz(Me.Controls("DS_TASK_PARAM_" & CStr(i)), ""))
                    If IsObject(Item) Then
                        Set t(i) = Item
                    Else
                        Select Case TaskParams(i)(1)
                            Case "Long": t(i) = CLng(Item)
                            Case "Boolean": t(i) = CBool(Item)
                            Case Else: t(i) = Item
                        End Select
                    End If
                Else
                    t(i) = CVErr(0)
                End If
            Else
                t(i) = CVErr(0)
            End If
        Next i
    End If

    If EditMode Then
        TaskId = pController.TaskController.RebuildSequence(ActiveTaskIndex)("Id")
        pController.TaskController.SetTask pController.TaskController.GenerateTask(TaskName, t, TaskId).Instance, ActiveTaskIndex
        pController.TaskController.SequenceIndex = ActiveTaskIndex
    Else
        i = IIf(pController.TaskController.RebuildSequence.Count - 1 < pController.TaskController.SequenceIndex, pController.TaskController.RebuildSequence.Count - 1, pController.TaskController.SequenceIndex)
        pController.TaskController.AddTask TaskName, t
        pController.TaskController.SequenceIndex = i + 1
    End If
    DoCmd.Close acForm, "DS_CREATE_TASK", acSaveNo
End Sub

Private Function ParseParamValue(ByVal Value As String) As Variant
    On Error GoTo Fallback

    Select Case VBA.Mid$(LTrim(Value), 1, 1)
        Case "[", "{"
            Assign ParseParamValue, JSON.Parse(Value, True, True)
        Case Else
            ParseParamValue = Value
    End Select

    Exit Function
Fallback:
    ParseParamValue = Value
End Function

Private Sub pContinuousList_OnActiveIndexChange(ByVal ActiveIndex As Long)
    pSelectedTask = pTaskNames(ActiveIndex)
    RebuildUIForTask pSelectedTask
End Sub

Private Sub pController_OnChange()
    If pSelectedTask <> "" Then
        RebuildUIForTask pSelectedTask
    End If
End Sub

Private Function GetControlText(ByRef TargetControl As Access.control) As String
    On Error Resume Next
    GetControlText = TargetControl.Value
    GetControlText = TargetControl.Text
    On Error GoTo 0
End Function

Private Sub SetControlText(ByRef TargetControl As Access.control, ByVal Text As String)
    On Error Resume Next
    TargetControl.Value = Text
    TargetControl.Text = Text
    On Error GoTo 0
End Sub

Private Sub Assign(ByRef Target, Optional ByRef Value As Variant)
    If Not IsMissing(Value) Then
        If IsObject(Value) Then
            Set Target = Value
        Else
            Target = Value
        End If
    End If
End Sub

Private Function GetExcelTableStyles() As Variant
    Dim i As Long

    With ArrayListEx.Create()
        For i = 1 To 21: .Add "TableStyleLight" & CStr(i): Next i
        For i = 1 To 28: .Add "TableStyleMedium" & CStr(i): Next i
        For i = 1 To 11: .Add "TableStyleDark" & CStr(i): Next i
        GetExcelTableStyles = .ToArray()
    End With
End Function

Private Sub ShowAvailableCustomVars()
    Dim dsT As dsTable, r As ds3xGlobals.RECT, rScreen As ds3xGlobals.RECT

    If pCVarsScrollview Is Nothing Then
        Set pCVarsScrollview = New Form_DS_SCROLLVIEW
        Set dsT = dsTable.Create(dsApp.CustomVars).SetHeaders(Array("CustomVar", "Value"))

        pCVarsScrollview.Visible = True
        ScreenLib.WindowSizeTo pCVarsScrollview, 5160, 8000
        ScreenLib.WindowAlwaysOnTop pCVarsScrollview
        r = ScreenLib.GetWindowRect(Me)
        rScreen = ScreenLib.GetScreenRectOfPoint(ScreenLib.PointInRect(r, DirectionType.Center))
        If CLng(r.Left) - CLng(rScreen.Left) > 5160 And r.Left > rScreen.Left Then
            ScreenLib.WindowMoveTo pCVarsScrollview, CLng(r.Left) - 5160, CLng(r.Top)
        ElseIf CLng(rScreen.Right) - CLng(r.Right) > 5160 And rScreen.Right > r.Right Then
            ScreenLib.WindowMoveTo pCVarsScrollview, CLng(r.Right), CLng(r.Top)
        End If
        pCVarsScrollview.Caption = "Custom Vars"
        Set pCVarsScrollview.Table = dsT
    Else
        pCVarsScrollview.Visible = False
        pCVarsScrollview.Visible = True
    End If
End Sub

Private Sub pCVarsScrollview_OnWindowClose(Cancel As Integer)
    Set pCVarsScrollview = Nothing
End Sub

Private Sub AddDefaultNumberFormatItems(ByVal TargetControl As Access.ComboBox)
    With TargetControl
        .AddItem ""
        .AddItem "General"
        .AddItem "@"
        .AddItem "0"
        .AddItem "0%"
        .AddItem "0.00%"
        .AddItem "#,##0.00;-#,##0.00"
        .AddItem "m/d/yyyy"
        .AddItem "mm/dd/yyyy"
        .AddItem "h:mm"
        .AddItem "h:mm:ss"
        .AddItem "[h]:mm:ss"
        .AddItem "mm:ss"
        .AddItem "hh:mm"
        .AddItem "hh:mm:ss"
        .AddItem "m/d/yyyy h:mm"
        .AddItem "m/d/yyyy h:mm:ss"
        .AddItem "mm/dd/yyyy hh:mm"
        .AddItem "mm/dd/yyyy hh:mm:ss"
    End With
End Sub

Private Function GetSelectedColumnNumberFormat() As String
    Dim cIndex As Long, cIndexValid As Boolean
    On Error GoTo Finally

    With pController.SelectedColumnIndexes
        If .Count > 0 Then
            cIndex = CLng(.Item(0))
            cIndexValid = True
        End If
    End With

    If cIndexValid Then
        GetSelectedColumnNumberFormat = CollectionsLib.ArrayItem(pController.Table.Headers.Row(0)(cIndex), 1, "General")
    End If
Finally:
End Function
