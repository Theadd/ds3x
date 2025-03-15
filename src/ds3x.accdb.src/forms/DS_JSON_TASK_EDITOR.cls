Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "ds3x.UI.LiveEditor"
Option Compare Database
Option Explicit

Private WithEvents pController As dsTaskController
Attribute pController.VB_VarHelpID = -1

Private pOriginalValue As String
Private pSequenceIndex As Long
Private pExpandedMode As Boolean


Property Get IsSubform() As Boolean
    On Error Resume Next
    IsSubform = Len(Me.Parent.Name) > 0
    On Error GoTo 0
End Property

Public Property Get Controller() As dsTaskController
    Set Controller = pController
End Property

Public Property Set Controller(ByRef Value As dsTaskController)
    Set pController = Value
End Property

Public Property Get ExpandedMode() As Boolean: ExpandedMode = pExpandedMode: End Property
Public Property Let ExpandedMode(ByVal Value As Boolean): SetExpandedMode Value: End Property



Private Sub SetExpandedMode(ByVal Value As Boolean)
    If pExpandedMode <> Value Then
        pExpandedMode = Value
        ResizeToFit IIf(pExpandedMode, 6, 1)
        If Not IsSubform Then Exit Sub
        Me.Parent.ResizeFormHeaderToSize Me.Detalle.Height
    End If
End Sub

Private Sub ResizeToFit(ByVal NumLines As Long)
    Const vPadding As Long = 120
    Const lineHeight As Long = 345

    Me.DS_JSON_EDITOR.Height = NumLines * lineHeight
    Me.Detalle.Height = (2 * vPadding) + (NumLines * lineHeight)
    Me.DS_JSON_EDITOR.Height = NumLines * lineHeight
End Sub



Private Sub DS_BUTTON_TOGGLE_EXPAND_GotFocus()
    Me.HiddenControl.SetFocus
    ExpandedMode = Not ExpandedMode
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyReturn And Shift = 0 Then
        KeyCode = 0
        Me.HiddenControl.SetFocus
        DoEvents
    End If
    On Error GoTo 0
End Sub

Private Sub Form_Resize()
    Me.DS_JSON_EDITOR.HorizontalAnchor = acHorizontalAnchorBoth
    Me.DS_JSON_EDITOR.VerticalAnchor = acVerticalAnchorBoth
End Sub

Private Sub pController_OnActiveSequenceIndexChange(ByVal TargetIndex As Long, ByVal HasFailed As Boolean)
    On Error GoTo Finally

    pSequenceIndex = TargetIndex    ' Controller.SequenceIndex
    pOriginalValue = Encode(JSON.Stringify(Controller.RebuildSequence(pSequenceIndex)))
    Me.DS_JSON_EDITOR = pOriginalValue
    ScreenLib.SetGridlineAsValid Me.DS_JSON_EDITOR, True

    Exit Sub
Finally:
    Debug.Print "[UNHANDLED ERROR] @DS_JSON_TASK_EDITOR.pController_OnActiveSequenceIndexChange()"
    Resume Next
End Sub

Private Sub DS_JSON_EDITOR_LostFocus()
    ApplyChanges
End Sub

Private Sub ApplyChanges()
    Dim sValue As String, d As Scripting.Dictionary
    Debug.Print GetControlText(Me.DS_JSON_EDITOR)
    sValue = Decode(Application.PlainText(GetControlText(Me.DS_JSON_EDITOR)))
    If pOriginalValue <> "" And pOriginalValue <> sValue And Trim(sValue) <> "" Then
        Debug.Print sValue
        If TryParseJsonAsTask(sValue, d) Then
            ScreenLib.SetGridlineAsValid Me.DS_JSON_EDITOR, True
            ReplaceTaskInRebuildSequence d, pSequenceIndex
        Else
            ScreenLib.SetGridlineAsValid Me.DS_JSON_EDITOR, False
        End If
    End If
End Sub

Private Function TryParseJsonAsTask(ByVal JSONString As String, ByRef OutTask As Scripting.Dictionary) As Boolean
    On Error GoTo Finally

    Set OutTask = JSON.Parse(JSONString, True, True)
    TryParseJsonAsTask = True
Finally:
End Function

Private Sub ReplaceTaskInRebuildSequence(ByVal Task As Scripting.Dictionary, ByVal Index As Long)
    Controller.SetTask Task, Index
End Sub

Private Function Encode(ByVal JSONString As String) As String
    Dim s As String

    s = Replace(JSONString, """TaskName"":", "TaskName:")
    s = Replace(s, """Id"":", "Id:")
    s = Replace(s, """Values"":", "Values:")
    s = Replace(s, "\\", "\")

    Encode = s
End Function

Private Function Decode(ByVal RenderedString As String) As String
    Dim s As String

    s = RenderedString
    s = Replace(s, "\", "\\")

    Decode = s
End Function

Private Function GetControlText(ByRef TargetControl As Access.control) As String
    On Error Resume Next
    GetControlText = TargetControl.Value
    GetControlText = TargetControl.Text
    On Error GoTo 0
End Function
