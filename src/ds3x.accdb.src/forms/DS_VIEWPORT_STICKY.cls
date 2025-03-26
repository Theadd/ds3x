Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "ds3x.UI.Scrollview"
Option Compare Database
Option Explicit


Private pWorksheetNumbers As Form_DS_WORKSHEET_NUMBERS
Private pScrollview As Form_DS_SCROLLVIEW


Public Property Get WorksheetNumbers() As Form_DS_WORKSHEET_NUMBERS: Set WorksheetNumbers = pWorksheetNumbers: End Property
Public Property Set WorksheetNumbers(ByRef Value As Form_DS_WORKSHEET_NUMBERS): Set pWorksheetNumbers = Value: End Property

Public Property Get Scrollview() As Form_DS_SCROLLVIEW: Set Scrollview = pScrollview: End Property
Public Property Set Scrollview(ByRef Value As Form_DS_SCROLLVIEW): Set pScrollview = Value: End Property


' --- FORM EVENTS ---

Private Sub Form_MouseWheel(ByVal Page As Boolean, ByVal Count As Long)
    On Error Resume Next
    pScrollview.PropagateMouseWheel Page, Count
    On Error GoTo 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer): pScrollview.OnKeyDownHandler KeyCode, Shift: End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer): pScrollview.OnKeyUpHandler KeyCode, Shift: End Sub


' --- SETUP / BINDING ---

Public Sub Setup()
    Dim r As ds3xGlobals.RECT, B As ds3xGlobals.BOUNDS, t As Long

    Set WorksheetNumbers = Me.DS_WORKSHEET_NUMBERS.Form

    r = ScreenLib.GetScreenRectOfPoint(ScreenLib.PointInRect(ScreenLib.GetWindowRect(Me), DirectionType.Center), True)
    B = ScreenLib.RectToBounds(r)

    t = pWorksheetNumbers.MaxContentWidthLimit
    Me.Width = t
    Me.FormDetail.Height = CLng(Min(B.h * 1.95, 31500))
    With Me.DS_WORKSHEET_NUMBERS
        .Left = 0
        .Top = 0
        .Width = t
        .Height = CLng(Min(B.h * 1.95, 31500))
    End With
    Me.Width = t
End Sub


' --- EVENTS ---

Private Sub DS_FULL_AXIS_SCROLLING_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    pScrollview.PointerCapture = True
End Sub

Private Sub DS_FULL_AXIS_SCROLLING_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Static lastPointerCapture As Long
    Dim pCaptured As Boolean: pCaptured = pScrollview.PointerCapture

    If CLng(pCaptured) - 1 <> lastPointerCapture Then
        lastPointerCapture = CLng(pCaptured) - 1
        If pCaptured Then ScreenLib.MouseMoveCursor = True
    Else
        If Not pCaptured Then ScreenLib.MouseHelpCursor = True
    End If
    DoEvents
End Sub

Private Sub DS_FULL_AXIS_SCROLLING_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    pScrollview.PointerCapture = False
End Sub

Private Sub DS_HC_2_0_Click()
    pScrollview.TriggerClickOnInvertSelection
End Sub

Private Sub DS_HC_2_0_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    ScreenLib.MouseHandCursor = True
    DoEvents
End Sub

Private Sub DS_HC_1_0_Click()
    Me.HiddenControl.SetFocus
    DoEvents

    pScrollview.TriggerClickOnSelectAll
End Sub


' --- HELPERS ---

Private Function Max(x As Variant, Y As Variant) As Variant: Max = IIf(x > Y, x, Y): End Function
Private Function Min(x As Variant, Y As Variant) As Variant: Min = IIf(x < Y, x, Y): End Function
