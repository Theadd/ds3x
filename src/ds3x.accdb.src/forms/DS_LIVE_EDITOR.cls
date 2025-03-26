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
Public Property Set Controller(ByRef Value As dsLiveEd): Bind Value: End Property

Public Property Let WindowVisibility(ByVal Value As Boolean)
    Me.Visible = Value
    If Value Then ScreenLib.WindowAndTaskbarIconAsVisible Me
End Property


Private Sub Form_Load()
    ScreenLib.ScreenLib_Resync

    ScreenLib.WindowSizeTo Me, 22050, 11000
    ScreenLib.WindowCenterTo Me, ScreenLib.PrimaryScreen

    Me.TimerInterval = 1
End Sub

Private Sub Bind(ByVal TargetController As dsLiveEd)

    If (Not TargetController Is Nothing) And (Not pController Is Nothing) Then
        ' More than one DS_LIVE_EDITOR windows can point to the same dsLiveEd controller
        ' But NOT more than one dsLiveEd instance can point to the same window, so:
        ' Unbind this form from controller's pLiveEditor if ObjPtr's of TargetController and pController differ
        If ObjPtr(pController) <> ObjPtr(TargetController) Then
            If Not pController.LiveEditor Is Nothing Then
                If ObjPtr(pController.LiveEditor) = ObjPtr(Me) Then
                    Set pController.LiveEditor = Nothing
                End If
            End If
            Set pController = Nothing
        End If
    End If
    Set pController = TargetController

    Me.DS_MAIN_PANEL.SourceObject = "Form.DS_LIVE_EDITOR_MAIN_PANEL"
    Me.DS_SIDE_PANEL.SourceObject = "Form.DS_LIVE_EDITOR_SIDE_PANEL"
    Set Me.DS_MAIN_PANEL.Form.Controller = pController
    Set Me.DS_SIDE_PANEL.Form.Controller = pController
    Me.DS_SIDE_PANEL.Visible = True
    Me.DS_MAIN_PANEL.Visible = True
End Sub

Private Sub Form_Resize()
    ResizeToFitAllControls
End Sub

Private Sub Form_Timer()
    Me.TimerInterval = 0

    If pController Is Nothing Then
        Dim d As dsLiveEd
        Set d = New dsLiveEd
        Bind d
        Set d.LiveEditor = Me
        d.DisposeOnClose = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    pController.TriggerLiveEditorWindowClose True
    Set pController = Nothing
    On Error GoTo 0
End Sub

Public Sub Dispose()
    On Error Resume Next
    Focus Me
    Me.SetFocus
    DoCmd.Close
    On Error GoTo 0
End Sub

Private Sub ResizeToFitAllControls()
    Dim parentWidth As Long, pSize As Long

    Focus Me
    parentWidth = Me.InsideWidth

    pSize = Me.DS_SIDE_PANEL.Width
    Me.DS_SIDE_PANEL.Left = parentWidth - pSize
    pSize = pSize + Me.DS_VERTICAL_SPLITTER.Width
    Me.DS_VERTICAL_SPLITTER.Left = parentWidth - pSize
    Me.DS_MAIN_PANEL.Width = parentWidth - (pSize + Me.DS_MAIN_PANEL.Left)
    ' Force resize event on side panel subform as quick fix for incorrect rendering
    If Me.DS_SIDE_PANEL.Width Mod 2 = 0 Then
        Me.DS_SIDE_PANEL.Width = Me.DS_SIDE_PANEL.Width - 15
    Else
        Me.DS_SIDE_PANEL.Width = Me.DS_SIDE_PANEL.Width + 15
    End If
End Sub

Private Sub pController_OnDataSourceChange()
    If Not Me.DS_MAIN_PANEL.Visible Then Me.DS_MAIN_PANEL.Visible = True
End Sub

Private Sub Focus(ByRef frm As Access.Form)
    On Error Resume Next
    frm.Controls("HiddenControl").SetFocus
    DoEvents
    On Error GoTo 0
End Sub
