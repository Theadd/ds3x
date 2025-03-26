Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "ds3x.UI.LiveEditor"
Option Compare Database
Option Explicit

Private WithEvents pController As dsLiveEd
Attribute pController.VB_VarHelpID = -1

Private pInitialized As Boolean

Property Get IsSubform() As Boolean
    On Error Resume Next
    IsSubform = Len(Me.Parent.Name) > 0
    On Error GoTo 0
End Property

Public Property Get Controller() As dsLiveEd: Set Controller = pController: End Property
Public Property Set Controller(ByRef Value As dsLiveEd): SetController Value: End Property




Private Sub DS_CLOSE_BUTTON_Click()
    RemoveFocus
    On Error Resume Next
    Me.Parent.Dispose
    On Error GoTo 0
End Sub

Private Sub DS_EDIT_BUTTON_Click()
    RemoveFocus
End Sub

Private Sub DS_GENERATE_BUTTON_Click()
    On Error Resume Next
    RemoveFocus
    Dim xlSheet As xlSheetsEx

    If Controller.TryGenerateAsExcel(xlSheet) Then
        xlSheet.WindowVisibility = True
    End If
    On Error GoTo 0
End Sub

Private Sub DS_NEW_BUTTON_Click()
    RemoveFocus
    Dim dsEditor As New dsLiveEd

    dsEditor.Visible = True
    Set dsEditor.LiveEditor.Controller = dsEditor
End Sub

Private Sub DS_OPEN_BUTTON_Click()
    On Error Resume Next
    RemoveFocus
    Controller.ImportPreset
    On Error GoTo 0
End Sub

Private Sub DS_SAVE_AS_BUTTON_Click()
    On Error Resume Next
    RemoveFocus
    Controller.SavePreset VBA.Environ$("USERPROFILE")
    On Error GoTo 0
End Sub

Private Sub DS_SAVE_BUTTON_Click()
    On Error Resume Next
    RemoveFocus
    Controller.SavePreset
    On Error GoTo 0
End Sub

Private Sub Form_Load()
    SetButtonAsClickable Me.DS_EDIT_BUTTON, False, 8
    ScreenLib.SetControlAsEnabled Me.DS_ENTRY_NAME, False
    SetButtonAsClickable Me.DS_GENERATE_BUTTON, False, 8
    SetButtonAsClickable Me.DS_SAVE_BUTTON, False, 8
    SetButtonAsClickable Me.DS_SAVE_AS_BUTTON, False, 8
End Sub

Private Sub pController_OnDataSourceChange()
    Rebuild
End Sub

Private Sub SetController(ByVal TargetController As dsLiveEd)
    Set pController = TargetController
    Set Me.DS_REBUILD_SEQUENCE.Form.GridController = pController.GridController
    Set Me.DS_REBUILD_SEQUENCE.Form.Controller = pController.TaskController

    If pController.TaskController.Initialized Then Rebuild
End Sub

Private Sub Rebuild()
    Me.DS_REBUILD_SEQUENCE.Form.RebuildSequenceList
End Sub

Private Sub RemoveFocus()
    Me.HiddenControl.SetFocus
    DoEvents
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

Public Sub RefreshClickableButtons()
    Dim hasTasks As Boolean
    On Error GoTo Finally
    hasTasks = (pController.TaskController.RebuildSequence.Count > 0)
Finally:
    On Error Resume Next
    SetButtonAsClickable Me.DS_GENERATE_BUTTON, hasTasks, 8
    SetButtonAsClickable Me.DS_SAVE_BUTTON, hasTasks, 8
    SetButtonAsClickable Me.DS_SAVE_AS_BUTTON, hasTasks, 8
    On Error GoTo 0
End Sub
