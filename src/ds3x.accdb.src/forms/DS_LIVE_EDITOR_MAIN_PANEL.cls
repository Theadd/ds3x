Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "ds3x.UI.LiveEditor"
Option Compare Database
Option Explicit

Private WithEvents pController As dsLiveEd
Attribute pController.VB_VarHelpID = -1

Property Get IsSubform() As Boolean
    On Error Resume Next
    IsSubform = Len(Me.Parent.Name) > 0
    On Error GoTo 0
End Property

Public Property Get Controller() As dsLiveEd: Set Controller = pController: End Property
Public Property Set Controller(ByRef Value As dsLiveEd): SetController Value: End Property



Private Sub Form_Resize()
    'LogMe "[DEBUG] @DS_LIVE_EDITOR_MAIN_PANEL.Resize()", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Debug.Print "[INFO] @DS_LIVE_EDITOR_MAIN_PANEL.Unload()"
End Sub

Private Sub pController_OnDataSourceChange()
    Rebuild
End Sub

Private Sub SetController(ByVal TargetController As dsLiveEd)
    Set pController = TargetController
    Set pController.GridController.Scrollview = Me.DS_SCROLLVIEW.Form
    Set Me.DS_JSON_TASK_EDITOR.Form.Controller = pController.TaskController
    Me.DS_JSON_TASK_EDITOR.Visible = True
    Me.DS_SCROLLVIEW.Visible = True
    Me.DS_SCROLLVIEW.Form.EnableOutOfRangeScrolling = True
    ' Me.DS_SCROLLVIEW.Form.InvertScrollInputAxis = True
    Me.DS_SCROLLVIEW.Form.KeepScrollPositionOnTableChange = True
    Rebuild
End Sub

Private Sub Rebuild()
    With pController.GridController
        .SelectionClear
        .TriggerOnChangeEvent
        .SelectionApply
    End With
End Sub


Public Sub ResizeFormHeaderToSize(ByVal h As Long)
    Me.DS_JSON_TASK_EDITOR.Height = h
    Me.EncabezadoDelFormulario.Height = h
    Me.DS_JSON_TASK_EDITOR.Height = h
End Sub
