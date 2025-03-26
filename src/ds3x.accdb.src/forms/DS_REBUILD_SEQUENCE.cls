Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "ds3x.UI.LiveEditor"
Option Compare Database
Option Explicit

Private WithEvents pGridController As dsGridController
Attribute pGridController.VB_VarHelpID = -1
Private WithEvents pController As dsTaskController
Attribute pController.VB_VarHelpID = -1
Private WithEvents pContinuousList As Form_DS_RICHTEXT_LIST
Attribute pContinuousList.VB_VarHelpID = -1
Private pFailedSequenceIndex As Long
Private rsX As RecordsetEx
Private pIndexCountdown As Long

Property Get IsSubform() As Boolean
    On Error Resume Next
    IsSubform = Len(Me.Parent.Name) > 0
    On Error GoTo 0
End Property

Public Property Get Controller() As dsTaskController: Set Controller = pController: End Property
Public Property Set Controller(ByRef Value As dsTaskController): SetTaskController Value: End Property
Public Property Get GridController() As dsGridController: Set GridController = pGridController: End Property
Public Property Set GridController(ByRef Value As dsGridController): Set pGridController = Value: End Property


Private Sub SetTaskController(ByRef Value As dsTaskController)
    Set pController = Value
    RebuildSequenceList
End Sub

Private Sub DS_LIST_ITEM_ADD_Click()
    Dim frm As Access.Form, ActiveTaskName As String

    If pController.Initialized Then
        ActiveTaskName = pController.RebuildSequence(pController.SequenceIndex)("TaskName")
    Else
        ActiveTaskName = ""
    End If
    CloseFormIfOpen "DS_CREATE_TASK"
    DoCmd.OpenForm "DS_CREATE_TASK"
    Set frm = Forms("DS_CREATE_TASK")
    ScreenLib.WindowCenterTo frm, ScreenLib.GetScreenRectOfPoint(ScreenLib.PointInRect(ScreenLib.GetWindowRect(Me), DirectionType.Center))
    frm.ActiveTaskName = ActiveTaskName
    frm.ActiveTaskIndex = pController.SequenceIndex
    frm.EditMode = False
    Set frm.Controller = pGridController
    ScreenLib.WindowAndTaskbarIconAsVisible frm
End Sub

Private Sub DS_LIST_ITEM_EDIT_Click()
    ShowCreateTaskFormInEditMode pController.SequenceIndex
End Sub

Private Sub ShowCreateTaskFormInEditMode(ByVal TaskIndex As Long)
    Dim frm As Access.Form, ActiveTaskName As String

    If pController.Initialized Then
        ActiveTaskName = pController.RebuildSequence(TaskIndex)("TaskName")
    Else
        ActiveTaskName = ""
    End If
    CloseFormIfOpen "DS_CREATE_TASK"
    DoCmd.OpenForm "DS_CREATE_TASK"
    Set frm = Forms("DS_CREATE_TASK")
    ScreenLib.WindowCenterTo frm, ScreenLib.GetScreenRectOfPoint(ScreenLib.PointInRect(ScreenLib.GetWindowRect(Me), DirectionType.Center))
    frm.ActiveTaskName = ActiveTaskName
    frm.ActiveTaskIndex = TaskIndex
    frm.EditMode = True
    Set frm.Controller = pGridController
    ScreenLib.WindowAndTaskbarIconAsVisible frm
End Sub

Private Sub CloseFormIfOpen(ByVal FormName As String)
    On Error Resume Next
    DoCmd.Close acForm, FormName, acSaveNo
    On Error GoTo 0
End Sub

Private Sub DS_LIST_ITEM_REMOVE_Click()
    Focus Me
    pController.RemoveTask pController.RebuildSequence(pController.SequenceIndex)("Id")
End Sub

Private Sub Form_Load()
    pFailedSequenceIndex = -1
End Sub

Private Function CreateListItemHeader(Optional ByVal Title As String = "") As Variant
    pIndexCountdown = pIndexCountdown - 1

    If Trim(Title) = "" Then
        CreateListItemHeader = Array(pIndexCountdown, Title, 4)
    Else
        CreateListItemHeader = Array(pIndexCountdown, _
            "<div align=center><font face=""Consolas"" size=1 color=""#A5A5A5"">" & Title & "</font><font color=white size=2>&nbsp;&nbsp;.<br>" & _
            "</font><font style=""BACKGROUND-COLOR:#808080"">______ _______ _________ ________ ________ _______ ____ ______ _____</font></div>", _
            4)
    End If
End Function

Public Sub RebuildSequenceList()
    Dim dX As New DictionaryEx, i As Long, qList As New ArrayListEx, activeTaskGroup As String, sAux As String, dsT As dsTable, dirtyIndex As Long, seqIndex As Long, lState As Long
    seqIndex = pController.SequenceIndex
    dirtyIndex = pController.DirtyRebuildIndex
    pIndexCountdown = 0

    For i = 0 To pController.RebuildSequence.Count - 1
        dX.Bind pController.RebuildSequence(i)
        If pFailedSequenceIndex = i Then
            lState = 3
        Else
            lState = IIf(i >= dirtyIndex, 0, 1)
        End If
        Select Case Left(pController.TaskInfo(dX("TaskName"))(4), 1)
            Case "X": sAux = "X"
            Case Else: sAux = "T"
        End Select
        If activeTaskGroup <> sAux Then
            qList.Add CreateListItemHeader(" ")
            Select Case sAux
                Case "X": qList.Add CreateListItemHeader("XL FORMATTING TASKS")
                Case Else: qList.Add CreateListItemHeader("TRANSFORMATION TASKS")
            End Select
            activeTaskGroup = sAux
        End If
        qList.Add Array(i, "<b>" & dX("Id") & "</b> " & dX("TaskName"), lState)
    Next i
    qList.Add CreateListItemHeader(" ")
    qList.RemoveAt 0

    Set dsT = dsTable.Create(qList).SetHeaders(Array("DS_LIST_ITEM_INDEX", "DS_LIST_ITEM", "DS_LIST_ITEM_STATE"))
    Set rsX = RecordsetEx.Create(dsT.NamedRecordset)
    Set Me.DS_SEQUENCE_LIST.Form.Recordset = rsX.Instance
    If pContinuousList Is Nothing Then
        Set pContinuousList = Me.DS_SEQUENCE_LIST.Form
        pContinuousList.SetActiveIndex seqIndex
    End If
    SetButtonAsClickable Me.DS_LIST_ITEM_REMOVE, (qList.Count > 0), 5
    SetButtonAsClickable Me.DS_LIST_ITEM_EDIT, (qList.Count > 0), 8
    If IsSubform Then Me.Parent.RefreshClickableButtons
End Sub

Private Sub Form_Resize()
    Me.DS_SEQUENCE_LIST.HorizontalAnchor = acHorizontalAnchorBoth
    Me.DS_SEQUENCE_LIST.VerticalAnchor = acVerticalAnchorBoth
    Me.DS_FOOTER_CELL_GROW.HorizontalAnchor = acHorizontalAnchorBoth
End Sub

Private Sub pContinuousList_OnActiveIndexChange(ByVal ActiveIndex As Long)
    On Error Resume Next
    pController.SequenceIndex = ActiveIndex
    On Error GoTo 0
End Sub

Private Sub pContinuousList_OnItemDoubleClick(ByVal ActiveIndex As Long)
    ShowCreateTaskFormInEditMode ActiveIndex
End Sub

Private Sub pContinuousList_OnNoActiveIndexChange(ByVal ActiveIndex As Long)
    pController.SequenceIndex = ActiveIndex
End Sub

Private Sub pController_OnActiveSequenceIndexChange(ByVal TargetIndex As Long, ByVal HasFailed As Boolean)
    On Error Resume Next
    If (Not HasFailed) And pFailedSequenceIndex <= TargetIndex Then pFailedSequenceIndex = -1
    RebuildSequenceList
    pContinuousList.SetActiveIndex pController.SequenceIndex, True
    On Error GoTo 0
End Sub

Private Sub pController_OnRebuildSequenceChange()
    pFailedSequenceIndex = -1
    RebuildSequenceList
End Sub

Private Sub pController_OnRebuildSequenceFail(ByVal FailedSequenceIndex As Long)
    On Error Resume Next
    pFailedSequenceIndex = FailedSequenceIndex
    RebuildSequenceList
    pContinuousList.SetActiveIndex pFailedSequenceIndex, True
    On Error GoTo 0
End Sub

Public Function Focus(ByRef frm As Access.Form)
    On Error Resume Next
    frm.Controls("HiddenControl").SetFocus
    DoEvents
    On Error GoTo 0
End Function

Private Sub SetButtonAsClickable(ByRef TargetControl As Access.CommandButton, ByVal Value As Boolean, Optional ByVal ThemeColorIndexOnNormal As Long = 8)
    ScreenLib.SetClickableStyle TargetControl, Value
    ScreenLib.SetThemeColorStyle TargetControl, IIf(Value, ThemeColorIndexOnNormal, 6)
End Sub
