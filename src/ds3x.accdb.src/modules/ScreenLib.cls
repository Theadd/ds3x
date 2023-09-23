VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ScreenLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder "ds3x.Libraries"
' ScreenLib Module
Option Compare Database
Option Explicit


' --- API DECLARATIONS --

Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function GetForegroundWindow Lib "user32.dll" () As LongPtr
Private Declare PtrSafe Function GetSystemMetrics32 Lib "user32" Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long
Private Declare PtrSafe Function GetClassName32 Lib "user32" Alias "GetClassNameA" (ByVal hWnd As LongPtr, ByVal lpClassName As String, ByVal nMaxCount As LongPtr) As Long

Private Declare PtrSafe Function IsZoomed Lib "user32" (ByVal hWnd As LongPtr) As Integer
Private Declare PtrSafe Function IsIconic Lib "user32" (ByVal hWnd As LongPtr) As Integer

Private Declare PtrSafe Function GetClientRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As RECT) As Long
Private Declare PtrSafe Function GetWindowRect32 Lib "user32" Alias "GetWindowRect" (ByVal hWnd As LongPtr, lpRect As RECT) As Long

Private Declare PtrSafe Function ClientToScreen Lib "user32" (ByVal hWnd As LongPtr, lpPoint As POINTAPI) As Long
Private Declare PtrSafe Function ScreenToClient Lib "user32" (ByVal hWnd As LongPtr, lpPoint As POINTAPI) As Long

Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare PtrSafe Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare PtrSafe Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As LongPtr, ByVal lpCursorName As Long) As LongPtr
Private Declare PtrSafe Function SetCursor Lib "user32" (ByVal hCursor As LongPtr) As LongPtr

Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hDC As LongPtr) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As LongPtr, ByVal nIndex As Long) As Long

Private Declare PtrSafe Function GetWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal wCmd As Long) As LongPtr
Private Declare PtrSafe Function SetShowWindow Lib "user32" Alias "ShowWindow" (ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Long
Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Private Declare PtrSafe Function GetDesktopWindow Lib "user32" () As LongPtr

Private Declare PtrSafe Function CalculatePopupWindowPosition Lib "user32" (anchorPoint As POINTAPI, windowSize As SIZE, ByVal flags As Long, excludeRect As RECT, outPosition As RECT) As Boolean  ' https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-calculatepopupwindowposition
Private Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As LongPtr, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Private Declare PtrSafe Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long

' --- API CONSTANTS ---

Private Const LOGPIXELSX As Long = 88   ' Pixels per logical inch in X
Private Const LOGPIXELSY As Long = 90   ' Pixels per logical inch in Y
Private Const TWIPSPERINCH As Long = 1440
'Forces a top-level window onto the taskbar when the window is visible.
Private Const WS_EX_APPWINDOW As Long = &H40000

Private Const GW_HWNDNEXT = 2
Private Const GW_CHILD = 5

Private Const LWA_ALPHA     As Long = &H2
Private Const GWL_EXSTYLE   As Long = -20
Private Const GWL_STYLE     As Long = -16
Private Const WS_EX_LAYERED As Long = &H80000
Private Const WS_VISIBLE    As Long = &H10000000

' ANCHORING RECTS - HORIZONTAL ALIGNMENT
Private Const TPM_LEFTALIGN As Long = &H0       ' Positions the window so that its left edge is aligned with the anchorPoint
Private Const TPM_CENTERALIGN As Long = &H4
Private Const TPM_RIGHTALIGN As Long = &H8      ' Positions the window so that its right edge is aligned with the anchorPoint

' ANCHORING RECTS - VERTICAL ALIGNMENT
Private Const TPM_TOPALIGN As Long = &H0        ' Positions the window so that its top edge is aligned with the anchorPoint
Private Const TPM_VCENTERALIGN As Long = &H10
Private Const TPM_BOTTOMALIGN As Long = &H20    ' Positions the window so that its bottom edge is aligned with the anchorPoint

' ANCHORING RECTS - DISPLACEMENT DIRECTION WHEN COLISIONING
Private Const TPM_HORIZONTAL As Long = &H0
Private Const TPM_VERTICAL As Long = &H40

' SetWindowPos FLAGS
Private Const SWP_NOSIZE As Long = &H1          ' Retains the current size (ignores the cx and cy parameters).
Private Const SWP_NOMOVE As Long = &H2          ' Retains the current position (ignores X and Y parameters).
Private Const SWP_NOREPOSITION As Long = &H200  ' Does not change the owner window's position in the Z order.
Private Const SWP_NOZORDER As Long = &H4        ' Retains the current Z order (ignores the hWndInsertAfter parameter).
Private Const SWP_NOACTIVATE As Long = &H10     ' Does not activate the window. If this flag is not set, the window is activated and moved to the top of either the topmost or non-topmost.
Private Const HWND_TOP As Long = 0           ' Places the window at the top of the Z order.
Private Const HWND_TOPMOST As Long = -1      ' Places the window above all non-topmost windows. The window maintains its topmost position even when it is deactivated.
Private Const HWND_BOTTOM As Long = 1        ' Places the window at the bottom of the Z order.

' --- TYPE DEFINITIONS ---

Private Type SIZE
    cX As Long
    cY As Long
End Type

Public Enum DirectionType
    Center = 0
    North = 1
    South = 2
    East = 4
    West = 8
    NorthEast = 5
    SouthEast = 6
    SouthWest = 10
    NorthWest = 9
End Enum


' --- MEMBERS ---

Public TwipsPerPixelX As Long   ' Twips per pixel for screen in X
Public TwipsPerPixelY As Long   ' Twips per pixel for screen in Y
Private pScreenOffset As POINTAPI

Private pMdiHwnd As LongPtr
Private pDesktopHwnd As LongPtr
Private pItem As Variant


' --- MONITORS ----

Private PrimaryScreenRect As RECT
Private SecondaryScreenRect As RECT
Private VirtualScreenRect As RECT
Private TotalScreenMonitors As Long
Private SecondaryScreenDirection As DirectionType


' --- MOUSE CURSORS ----

Private pMouseMoveCursor As LongPtr
Private pMouseHandCursor As LongPtr
Private pMouseBusyCursor As LongPtr
Private pMouseLoadingCursor As LongPtr
Private pMouseUnavailableCursor As LongPtr
Private pMouseHelpCursor As LongPtr


' --- FORM CONTROLS STYLING ---

Private Const SC_GRAYED_COLOR As Long = 15461355
Private Const SC_GRAYED_SHADE As Long = 92
Private Const SC_GRAYED_THEME As Long = 1
Private Const SC_GRAYED_TINT As Long = 100

Private Const SC_WHITE_COLOR As Long = 16777215
Private Const SC_WHITE_SHADE As Long = 100
Private Const SC_WHITE_THEME As Long = 1
Private Const SC_WHITE_TINT As Long = 100

Private Const SC_INVALID_COLOR As Long = 2366701


' --- PROPERTIES ---

Public Property Get ScreenOffset() As POINTAPI: ScreenOffset = pScreenOffset: End Property
Public Property Get PrimaryScreen() As RECT: PrimaryScreen = PrimaryScreenRect: End Property
Public Property Get VirtualScreen() As RECT: VirtualScreen = VirtualScreenRect: End Property
Public Property Get TotalMonitors() As Long: TotalMonitors = TotalScreenMonitors: End Property

Public Property Get IsAccessWindowVisible() As Boolean
    On Error Resume Next
    IsAccessWindowVisible = ((GetWindowLong(Application.hWndAccessApp, GWL_STYLE) And WS_VISIBLE) > 0)
End Property

Public Property Get IsVBEMainWindowVisible() As Boolean
    On Error Resume Next
    Dim uHwnd As Long: uHwnd = Application.VBE.MainWindow.hWnd
    
    If uHwnd > 0 Then
        IsVBEMainWindowVisible = ((GetWindowLong(uHwnd, GWL_STYLE) And WS_VISIBLE) > 0)
    End If
End Property

Public Property Let IsVBEMainWindowVisible(ByVal ShouldBeVisible As Boolean)
    On Error Resume Next
    Dim uHwnd As Long: uHwnd = Application.VBE.MainWindow.hWnd
    
    If uHwnd > 0 Then
        SetShowWindow uHwnd, IIf(ShouldBeVisible, SW_SHOW, SW_HIDE)
    End If
End Property


' --- PROPERTIES: MOUSE CURSORS ---

' More cursor types at https://learn.microsoft.com/en-us/windows/win32/menurc/about-cursors
Public Property Let MouseMoveCursor(ByVal Value As Boolean): SetCursorType pMouseMoveCursor, 32646&, Value: End Property
Public Property Let MouseHandCursor(ByVal Value As Boolean): SetCursorType pMouseHandCursor, 32649&, Value: End Property
Public Property Let MouseBusyCursor(ByVal Value As Boolean): SetCursorType pMouseBusyCursor, 32514&, Value: End Property
Public Property Let MouseLoadingCursor(ByVal Value As Boolean): SetCursorType pMouseLoadingCursor, 32650&, Value: End Property
Public Property Let MouseUnavailableCursor(ByVal Value As Boolean): SetCursorType pMouseUnavailableCursor, 32648&, Value: End Property
Public Property Let MouseHelpCursor(ByVal Value As Boolean): SetCursorType pMouseHelpCursor, 32651&, Value: End Property


' --- PRIVATE FUNCTIONS ---

Private Function PointToSize(p As POINTAPI) As SIZE: PointToSize.cX = p.X: PointToSize.cY = p.Y: End Function

' --- PUBLIC FUNCTIONS ---

Public Function TwipsInPixelsX(PixelsX As Long) As Long: TwipsInPixelsX = PixelsX * TwipsPerPixelX: End Function
Public Function TwipsInPixelsY(PixelsY As Long) As Long: TwipsInPixelsY = PixelsY * TwipsPerPixelY: End Function

Public Function PointToString(p As POINTAPI) As String: PointToString = Printf("(%1, %2)", p.X, p.Y): End Function
Public Function RectToString(r As RECT) As String: RectToString = Printf("(%1, %2) -> (%3, %4)", r.Left, r.Top, r.Right, r.Bottom): End Function
Public Function SerializeRect(r As RECT) As String: SerializeRect = Printf("%1,%2,%3,%4", r.Left, r.Top, r.Right, r.Bottom): End Function

Public Function UnserializeRect(ByVal s As String) As RECT
    Dim v As Variant, i As Long

    With UnserializeRect
        For Each v In Split(s, ",")
            Select Case i
                Case 0: .Left = CLng(v)
                Case 1: .Top = CLng(v)
                Case 2: .Right = CLng(v)
                Case 3: .Bottom = CLng(v)
            End Select
            i = i + 1
        Next v
    End With
End Function

Public Function GetPoint(Optional ByVal X As Long = 0, Optional ByVal Y As Long = 0) As POINTAPI: GetPoint.X = X: GetPoint.Y = Y: End Function
Public Function GetRect(ByVal X As Long, ByVal Y As Long, ByVal X2 As Long, ByVal Y2 As Long) As RECT: GetRect.Left = X: GetRect.Top = Y: GetRect.Right = X2: GetRect.Bottom = Y2: End Function
Public Function GetBounds(ByVal X As Long, ByVal Y As Long, ByVal W As Long, ByVal h As Long) As BOUNDS: GetBounds.X = X: GetBounds.Y = Y: GetBounds.W = W: GetBounds.h = h: End Function
Public Function RectToBounds(r As RECT) As BOUNDS: RectToBounds.X = r.Left: RectToBounds.Y = r.Top: RectToBounds.W = (r.Right - r.Left): RectToBounds.h = (r.Bottom - r.Top): End Function
Public Function BoundsToRect(b As BOUNDS) As RECT: BoundsToRect.Left = b.X: BoundsToRect.Top = b.Y: BoundsToRect.Right = (b.X + b.W): BoundsToRect.Bottom = (b.Y + b.h): End Function

Public Function RectAsBounds(r As RECT) As BOUNDS: RectAsBounds.X = r.Left: RectAsBounds.Y = r.Top: RectAsBounds.W = r.Right: RectAsBounds.h = r.Bottom: End Function
Public Function BoundsAsRect(b As BOUNDS) As RECT: BoundsAsRect.Left = b.X: BoundsAsRect.Top = b.Y: BoundsAsRect.Right = b.W: BoundsAsRect.Bottom = b.h: End Function

Public Function InvertPoint(p As POINTAPI) As POINTAPI: InvertPoint.X = 0 - p.X: InvertPoint.Y = 0 - p.Y: End Function
Public Function DirectionPoint(d As DirectionType) As POINTAPI: DirectionPoint.X = Clamp101((d And East) - (d And West)): DirectionPoint.Y = Clamp101((d And South) - (d And North)): End Function

Public Function PointInRect(r As RECT, d As DirectionType) As POINTAPI
    With DirectionPoint(d)
        PointInRect.X = Switch(.X = -1, r.Left, .X = 0, r.Left + ((r.Right - r.Left) / 2), .X = 1, r.Right)
        PointInRect.Y = Switch(.Y = -1, r.Top, .Y = 0, r.Top + ((r.Bottom - r.Top) / 2), .Y = 1, r.Bottom)
    End With
End Function

Public Function AddOffsetToRect(r As RECT, p As POINTAPI) As RECT: With AddOffsetToRect: .Left = r.Left + p.X: .Top = r.Top + p.Y: .Right = r.Right + p.X: .Bottom = r.Bottom + p.Y: End With: End Function
Public Function RemoveOffsetFromRect(r As RECT, p As POINTAPI) As RECT: RemoveOffsetFromRect = AddOffsetToRect(r, InvertPoint(p)): End Function

Public Function AddOffsetToPoint(t As POINTAPI, p As POINTAPI) As POINTAPI: AddOffsetToPoint.X = t.X + p.X: AddOffsetToPoint.Y = t.Y + p.Y: End Function
Public Function RemoveOffsetFromPoint(t As POINTAPI, p As POINTAPI) As POINTAPI: RemoveOffsetFromPoint = AddOffsetToPoint(t, InvertPoint(p)): End Function

Public Function RectInTwips(r As RECT) As RECT: With RectInTwips: .Left = r.Left * TwipsPerPixelX: .Top = r.Top * TwipsPerPixelY: .Right = r.Right * TwipsPerPixelX: .Bottom = r.Bottom * TwipsPerPixelY: End With: End Function
Public Function PointInTwips(p As POINTAPI) As POINTAPI: PointInTwips.X = p.X * TwipsPerPixelX: PointInTwips.Y = p.Y * TwipsPerPixelY: End Function

Public Function RectInPixels(r As RECT) As RECT
    With RectInPixels
        .Left = CLng(r.Left / TwipsPerPixelX)
        .Top = CLng(r.Top / TwipsPerPixelY)
        .Right = CLng(r.Right / TwipsPerPixelX)
        .Bottom = CLng(r.Bottom / TwipsPerPixelY)
    End With
End Function

Public Function PointInPixels(p As POINTAPI) As POINTAPI
    With PointInPixels
        .X = CLng(p.X / TwipsPerPixelX)
        .Y = CLng(p.Y / TwipsPerPixelY)
    End With
End Function

Public Function GetCursorPosition() As POINTAPI: GetCursorPos GetCursorPosition: GetCursorPosition = PointInTwips(GetCursorPosition): End Function
Public Sub SetCursorPosition(p As POINTAPI)
    With PointInPixels(p)
        SetCursorPos .X, .Y
    End With
End Sub

Public Function GetWindowRect(f As Access.Form) As RECT: GetWindowRect32 f.hWnd, GetWindowRect: GetWindowRect = RectInTwips(GetWindowRect): End Function

Public Function GetFormInnerRect(f As Access.Form) As RECT
    Dim p As POINTAPI, r As RECT, r2 As RECT
    
    GetWindowRect32 f.hWnd, r
    p = GetPoint(r.Left, r.Top)
    ScreenToClient f.hWnd, p
    r = RemoveOffsetFromRect(r, p)
    GetClientRect f.hWnd, r2
    r.Right = r.Left + r2.Right
    r.Bottom = r.Top + r2.Bottom
    GetFormInnerRect = RectInTwips(r)
End Function

Public Function GetControlRect(t As Access.Control) As RECT
    Dim f As Access.Form: Set f = GetParentForm(t)
    GetControlRect = AddOffsetToRect(GetControlRectInForm(t), PointInRect(GetFormInnerRect(f), NorthWest))
End Function

Public Function GetScreenRectOfPoint(p As POINTAPI, Optional ByVal defaultToPrimaryScreen As Boolean = True) As RECT
    Select Case GetMonitorDirectionOfPoint(p)
        Case DirectionType.Center: GetScreenRectOfPoint = PrimaryScreenRect
        Case SecondaryScreenDirection: GetScreenRectOfPoint = SecondaryScreenRect
        Case Else
            If defaultToPrimaryScreen Then
                GetScreenRectOfPoint = PrimaryScreenRect
            End If
    End Select
End Function


' --- GET/SET FOREGROUND WINDOW ---

Public Function TrySetForegroundWindow(ByVal hWnd As Long) As Boolean
    On Error GoTo Finally
    
    TrySetForegroundWindow = (SetForegroundWindow(hWnd) <> 0)
    
Finally:
End Function

Public Function TryGetForegroundWindow(ByRef hWnd As LongPtr) As Boolean
    On Error GoTo Finally
    
    hWnd = GetForegroundWindow()
    TryGetForegroundWindow = True
    
Finally:
End Function


' --- MSAccess Window & Taskbar App Icons ---

' Hide/show/maximize/minimize/restore MSAccess Window.
'
' @EXAMPLE: Restore MSAccess Window:
'
'    SetAccessWindow SW_SHOWNORMAL
'
Public Function SetAccessWindow(nCmdShow As Long) As Boolean
    On Error Resume Next
    SetAccessWindow = (SetShowWindow(Application.hWndAccessApp, nCmdShow) <> 0)
End Function

Public Function ShowWindow(ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Long
    ShowWindow = SetShowWindow(hWnd, nCmdShow)
End Function

Public Function HideAppWindow(frm As Access.Form)
    On Error Resume Next
    SetWindowLong frm.hWnd, GWL_EXSTYLE, GetWindowLong(frm.hWnd, GWL_EXSTYLE) Or WS_EX_APPWINDOW
    SetShowWindow Application.hWndAccessApp, SW_HIDE
    SetShowWindow frm.hWnd, SW_SHOW
 End Function

Public Sub RestoreTaskbarIconForAllOpenWindows(Optional ByVal HideMainAccessWindow As Boolean = True)
    Dim f As Access.Form, i As Long
    On Error Resume Next
    
    For i = Forms.Count - 1 To 0 Step -1
        If i < Forms.Count Then
            Set f = Forms(i)
            If f.Name <> "F_TIMER" Then
                ' ShowTaskbarWindowIcon
                SetWindowLong f.hWnd, GWL_EXSTYLE, GetWindowLong(f.hWnd, GWL_EXSTYLE) Or WS_EX_APPWINDOW
            End If
        End If
    Next i
    
    If HideMainAccessWindow Then SetShowWindow Application.hWndAccessApp, SW_HIDE
    
    For i = Forms.Count - 1 To 0 Step -1
        If i < Forms.Count Then
            If Forms(i).Name <> "F_TIMER" Then
                SetShowWindow Forms(i).hWnd, SW_SHOW
            End If
        End If
    Next i
End Sub

Public Sub WindowAndTaskbarIconAsVisible(ByRef f As Access.Form)
    SetWindowLong f.hWnd, GWL_EXSTYLE, GetWindowLong(f.hWnd, GWL_EXSTYLE) Or WS_EX_APPWINDOW
    DoEvents
    SetShowWindow f.hWnd, SW_SHOW
End Sub


' --- TOOLTIP WINDOW ANCHORING ---

Public Function GetAnchoredRectRelativeTo(r As RECT, sizeW As Long, sizeH As Long, Optional ByVal dAnchor As DirectionType = NorthEast) As RECT
    Dim success As Boolean, rAnchored As RECT, pAnchor As POINTAPI, rSize As SIZE, rX As RECT, aFlags As Long
    
    rX = RectInPixels(r)
    pAnchor = PointInRect(rX, dAnchor)
    rSize = PointToSize(PointInPixels(GetPoint(sizeW, sizeH)))
    aFlags = GetAlignmentFlagsFromAnchorDirection(dAnchor)
    success = CalculatePopupWindowPosition(pAnchor, rSize, aFlags, rX, rAnchored)

    GetAnchoredRectRelativeTo = RectInTwips(rAnchored)
End Function


' --- WINDOW OPACITY ---

' @EXAMPLE:
'     Dim bgColor as Long: bgColor = RGB(0, 0, 0)
'     Me.THE_TRANSPARENT_CONTROL.BackColor = bgColor
'     FormColorOpacity Me, 0.5, bgColor
Public Sub FormColorOpacity(f As Access.Form, sngOpacity As Single, TColor As Long)
    SetWindowLong f.hWnd, GWL_EXSTYLE, (GetWindowLong(f.hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
    SetLayeredWindowAttributes f.hWnd, TColor, (sngOpacity * 255), LWA_ALPHA
End Sub


' --- WINDOW MAXIMIZED / MINIMIZED ---

Public Function IsMaximized(hWnd As Long) As Boolean
    IsMaximized = IsZoomed(hWnd) * -1
End Function

Public Function IsMinimized(hWnd As Long) As Boolean
    IsMinimized = IsIconic(hWnd) * -1
End Function


' --- SET WINDOW SIZE / POSITION ---

Public Sub WindowMoveTo(f As Access.Form, X As Long, Y As Long)
    SetWindowPos f.hWnd, HWND_TOP, CLng(X / TwipsPerPixelX), CLng(Y / TwipsPerPixelY), 0, 0, (SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE)
End Sub

Public Sub WindowSizeTo(f As Access.Form, W As Long, h As Long)
    SetWindowPos f.hWnd, HWND_TOP, 0, 0, CLng(W / TwipsPerPixelX), CLng(h / TwipsPerPixelY), (SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOACTIVATE)
End Sub

Public Sub WindowMoveSize(f As Access.Form, X As Long, Y As Long, W As Long, h As Long)
    SetWindowPos f.hWnd, HWND_TOP, CLng(X / TwipsPerPixelX), CLng(Y / TwipsPerPixelY), CLng(W / TwipsPerPixelX), CLng(h / TwipsPerPixelY), (SWP_NOZORDER Or SWP_NOACTIVATE)
End Sub

Public Sub WindowBringToTop(f As Access.Form)
    SetWindowPos f.hWnd, HWND_TOP, 0, 0, 0, 0, (SWP_NOMOVE Or SWP_NOSIZE)
End Sub

Public Sub WindowAlwaysOnTop(f As Access.Form)
    SetWindowPos f.hWnd, HWND_TOPMOST, 0, 0, 0, 0, (SWP_NOMOVE Or SWP_NOSIZE)
End Sub

Public Sub WindowSendToBack(f As Access.Form)
    SetWindowPos f.hWnd, HWND_BOTTOM, 0, 0, 0, 0, (SWP_NOMOVE Or SWP_NOSIZE)
End Sub

Public Sub WindowCenterTo(f As Access.Form, r As RECT)
    Dim c As POINTAPI, b As BOUNDS
    
    c = PointInRect(r, DirectionType.Center)
    b = RectToBounds(GetWindowRect(f))

    SetWindowPos f.hWnd, HWND_TOP, CLng((c.X - (b.W / 2)) / TwipsPerPixelX), CLng((c.Y - (b.h / 2)) / TwipsPerPixelY), 0, 0, (SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE)
End Sub

Public Sub WindowMaximize(f As Access.Form): SetShowWindow f.hWnd, SW_SHOWMAXIMIZED: End Sub
Public Sub WindowHide(f As Access.Form): SetShowWindow f.hWnd, SW_HIDE: End Sub
Public Sub WindowShow(f As Access.Form): SetShowWindow f.hWnd, SW_SHOW: End Sub

Public Sub WindowLockFlickering(f As Access.Form)
    LockWindowUpdate GetTopParentForm(f).hWnd
End Sub

Public Sub WindowUnlockFlickering()
    LockWindowUpdate 0
End Sub

Private Function GetTopParentForm(ByRef TargetForm As Access.Form, Optional ByRef UpToParentHwnd As Long = -1) As Access.Form
    On Error GoTo Finally
    
    If TargetForm.hWnd = UpToParentHwnd Then GoTo Finally
    
    If TypeOf TargetForm.Parent Is Form Then
        Set GetTopParentForm = GetTopParentForm(TargetForm.Parent, UpToParentHwnd)
    Else
        Set GetTopParentForm = GetTopParentForm(TargetForm.Parent.Parent, UpToParentHwnd)
    End If
    
    Exit Function
Finally:
    Set GetTopParentForm = TargetForm
End Function


' --- INITIALIZE / RESYNC ---

Public Sub ScreenLib_Resync()
    Dim lgDC As LongPtr
    
    pMdiHwnd = GetMDIClientHwnd
    pDesktopHwnd = GetDesktopWindow
    lgDC = GetDC(pDesktopHwnd)
    TwipsPerPixelX = TWIPSPERINCH / GetDeviceCaps(lgDC, LOGPIXELSX)
    TwipsPerPixelY = TWIPSPERINCH / GetDeviceCaps(lgDC, LOGPIXELSY)
    ReleaseDC pDesktopHwnd, lgDC
    pScreenOffset = GetPoint
    ScreenToClient pMdiHwnd, pScreenOffset
    pScreenOffset = PointInTwips(pScreenOffset)
    PrimaryScreenRect = RectInTwips(GetRect(0, 0, CLng(GetSystemMetrics32(0)), CLng(GetSystemMetrics32(1))))
    VirtualScreenRect = RectInTwips(BoundsToRect(GetBounds(GetSystemMetrics32(76), GetSystemMetrics32(77), GetSystemMetrics32(78), GetSystemMetrics32(79))))
    TotalScreenMonitors = GetSystemMetrics32(80)
    If TotalScreenMonitors > 1 Then
        SecondaryScreenDirection = GetSecondaryScreenDirection
        SecondaryScreenRect = GetSecondaryScreenRect
    End If
End Sub


' --- PRIVATE ---

Private Function GetControlRectInForm(t As Access.Control) As RECT
    Dim p As POINTAPI: p = GetPoint
    Dim f As Access.Form: Set f = GetParentForm(t)
    
    If t.Section <> acHeader Then
        p = AddOffsetToPoint(p, GetFormSectionAsOffset(f, acHeader))
        If t.Section <> acDetail Then
            p = AddOffsetToPoint(p, GetFormSectionAsOffset(f, acDetail))
        End If
    End If
    
    GetControlRectInForm = AddOffsetToRect(GetControlRectInSection(t, f), p)
End Function

Private Function GetControlRectInSection(t As Access.Control, f As Access.Form) As RECT
    Dim lMod As Long, wMod As Long, tMod As Long, hMod As Long
    
    Select Case t.HorizontalAnchor
        Case acHorizontalAnchorRight: wMod = Max(f.InsideWidth, f.Width) - f.Width: lMod = wMod
        Case acHorizontalAnchorBoth: wMod = Max(f.InsideWidth, f.Width) - f.Width: lMod = 0
        Case Else: lMod = 0: wMod = 0
    End Select
    
    Select Case t.VerticalAnchor
        Case acVerticalAnchorBottom: hMod = Max(f.InsideHeight, f.Height) - f.Height: tMod = hMod
        Case acVerticalAnchorBoth: hMod = Max(f.InsideHeight, f.Height) - f.Height: tMod = 0
        Case Else: tMod = 0: hMod = 0
    End Select
    
    GetControlRectInSection = GetRect(t.Left + lMod, t.Top + tMod, t.Left + wMod + t.Width, t.Top + hMod + t.Height)
End Function


Private Function GetFormSectionAsOffset(f As Access.Form, t As AcSection) As POINTAPI
    On Error GoTo Finally
    
    GetFormSectionAsOffset = GetPoint
    With f.Section(t)
        If .Visible Then GetFormSectionAsOffset.Y = .Height
    End With
    
Finally:
End Function

Private Function GetParentForm(ByRef bControl As Access.Control) As Access.Form
    On Error GoTo HandleError
    Dim bForm As Access.Form
    
    Set bForm = bControl.Parent
    Set GetParentForm = bForm
    Exit Function
    
HandleError:
    Set bForm = GetParentForm(bControl.Parent)
    Set GetParentForm = bForm
End Function

' DEPRECATED, BETTER USE Control.Section INSTEAD.
Private Function IsControlInFormSection(f As Access.Form, t As AcSection, c As Access.Control) As Boolean
    On Error GoTo Finally
    
    For Each pItem In f.Section(t).Controls
        If pItem.Name = c.Name Then
            IsControlInFormSection = True
            Exit For
        End If
    Next pItem
    
Finally:
End Function

Private Function Clamp101(n As Variant) As Integer
    On Error Resume Next
    If n <> 0 Then Clamp101 = (n / Abs(n))
End Function

Private Function GetAlignmentFlagsFromAnchorDirection(ByVal d As DirectionType) As Long
    Dim v As Long
    
    Select Case d
        Case DirectionType.North: v = (TPM_BOTTOMALIGN Or TPM_CENTERALIGN)
        Case DirectionType.NorthEast: v = (TPM_BOTTOMALIGN Or TPM_LEFTALIGN)
        Case DirectionType.East: v = (TPM_VCENTERALIGN Or TPM_LEFTALIGN)
        Case DirectionType.SouthEast: v = (TPM_TOPALIGN Or TPM_LEFTALIGN)
        Case DirectionType.South: v = (TPM_TOPALIGN Or TPM_CENTERALIGN)
        Case DirectionType.SouthWest: v = (TPM_TOPALIGN Or TPM_RIGHTALIGN)
        Case DirectionType.West: v = (TPM_VCENTERALIGN Or TPM_RIGHTALIGN)
        Case DirectionType.NorthWest: v = (TPM_BOTTOMALIGN Or TPM_RIGHTALIGN)
    End Select
    
    GetAlignmentFlagsFromAnchorDirection = v
End Function

Private Function GetMDIClientHwnd() As LongPtr
    'Returns the handle of Access's MDI background
    Dim lgHwnd As LongPtr, stName As String
    lgHwnd = GetWindow(Application.hWndAccessApp, GW_CHILD)
    
    'Get class name of child windows
    Do
        stName = GetClassName(lgHwnd)
        If LCase(stName) = "mdiclient" Then
            GetMDIClientHwnd = lgHwnd
            Exit Function
        End If
        lgHwnd = GetWindow(lgHwnd, GW_HWNDNEXT)
    Loop While lgHwnd <> 0
    
End Function

Private Function GetClassName(ByVal lgHwnd As LongPtr) As String
    Dim stBuf As String, dl As Long
    'Initialize space
    stBuf = String$(255, 0)
    dl = GetClassName32(lgHwnd, stBuf, 255)
    If InStr(stBuf, Chr$(0)) Then stBuf = Left$(stBuf, InStr(stBuf, Chr$(0)) - 1)
    GetClassName = stBuf
End Function

Private Function SetCursorType(ByRef TargetCursor As LongPtr, ByVal CursorId As Long, ByVal ActiveState As Boolean) As LongPtr
    If ActiveState Then
        If Not TargetCursor > 0 Then TargetCursor = LoadCursor(0, CLng(CursorId))
        If TargetCursor > 0 Then
            SetCursorType = SetCursor(TargetCursor)
            Exit Function
        End If
    End If
    Screen.MousePointer = 1
End Function


' --- Multi-displays ---

Private Function GetSecondaryScreenRect() As RECT
    Dim vBounds As BOUNDS, pBounds As BOUNDS, sBounds As BOUNDS
    If Not CLng(SecondaryScreenDirection) > CLng(DirectionType.Center) Then Exit Function
    
    pBounds = RectToBounds(PrimaryScreenRect)
    vBounds = RectToBounds(VirtualScreenRect)
    sBounds = GetBounds(0, 0, vBounds.W - pBounds.W, vBounds.h - pBounds.h)
    If sBounds.W = 0 Then sBounds.W = pBounds.W
    If sBounds.h = 0 Then sBounds.h = pBounds.h
    
    If TotalScreenMonitors > 2 Then
        sBounds.W = Min(Max(pBounds.W, CLng(sBounds.W / (TotalScreenMonitors - 1))), sBounds.W)
        sBounds.h = Min(Max(pBounds.h, CLng(sBounds.h / (TotalScreenMonitors - 1))), sBounds.h)
    End If
    
    Select Case SecondaryScreenDirection
        Case DirectionType.North: GetSecondaryScreenRect = GetRect(0, 0 - sBounds.h, sBounds.W, 0)
        Case DirectionType.West: GetSecondaryScreenRect = GetRect(0 - sBounds.W, 0, 0, sBounds.h)
        Case DirectionType.East: GetSecondaryScreenRect = GetRect(pBounds.W, 0, pBounds.W + sBounds.W, sBounds.h)
        Case DirectionType.South: GetSecondaryScreenRect = GetRect(0, pBounds.h, sBounds.W, pBounds.h + sBounds.h)
    End Select
End Function

Private Function GetSecondaryScreenDirection() As DirectionType
    GetSecondaryScreenDirection = DirectionType.Center
    If TotalScreenMonitors <= 1 Then Exit Function
    
    Select Case True
        Case (VirtualScreenRect.Left < 0 And VirtualScreenRect.Top = 0): GetSecondaryScreenDirection = DirectionType.West
        Case (VirtualScreenRect.Left = 0 And VirtualScreenRect.Top < 0): GetSecondaryScreenDirection = DirectionType.North
        Case (VirtualScreenRect.Right > PrimaryScreenRect.Right And VirtualScreenRect.Bottom = PrimaryScreenRect.Bottom): GetSecondaryScreenDirection = DirectionType.East
        Case (VirtualScreenRect.Right = PrimaryScreenRect.Right And VirtualScreenRect.Bottom > PrimaryScreenRect.Bottom): GetSecondaryScreenDirection = DirectionType.South
        Case (VirtualScreenRect.Top < 0): GetSecondaryScreenDirection = DirectionType.North
        Case (VirtualScreenRect.Bottom > PrimaryScreenRect.Bottom): GetSecondaryScreenDirection = DirectionType.South
    End Select
End Function

Private Function GetMonitorDirectionOfPoint(p As POINTAPI) As DirectionType
    GetMonitorDirectionOfPoint = DirectionType.Center
    
    Select Case p.X
        Case Is < PrimaryScreenRect.Left: GetMonitorDirectionOfPoint = GetMonitorDirectionOfPoint Or DirectionType.West
        Case Is > PrimaryScreenRect.Right: GetMonitorDirectionOfPoint = GetMonitorDirectionOfPoint Or DirectionType.East
    End Select
    
    Select Case p.Y
        Case Is < PrimaryScreenRect.Top: GetMonitorDirectionOfPoint = GetMonitorDirectionOfPoint Or DirectionType.North
        Case Is > PrimaryScreenRect.Bottom: GetMonitorDirectionOfPoint = GetMonitorDirectionOfPoint Or DirectionType.South
    End Select
End Function


' --- FORM CONTROLS STYLING ---

Private Function SetGrayedStyle(ByRef TargetControl As Access.Control, ByVal asGrayed As Boolean)
    Dim TargetLabel As Access.Label, styleLabel As Boolean

    If TargetControl.BackColor <> IIf(asGrayed, SC_WHITE_COLOR, SC_GRAYED_COLOR) Then Exit Function
    
    If TryGetAttachedLabel(TargetControl, TargetLabel) Then
        styleLabel = (TargetLabel.BackColor = TargetControl.BackColor)
    End If
    
    With TargetControl
        .BackColor = IIf(asGrayed, SC_GRAYED_COLOR, SC_WHITE_COLOR)
        .BackShade = IIf(asGrayed, SC_GRAYED_SHADE, SC_WHITE_SHADE)
        .BackThemeColorIndex = IIf(asGrayed, SC_GRAYED_THEME, SC_WHITE_THEME)
        .BackTint = IIf(asGrayed, SC_GRAYED_TINT, SC_WHITE_TINT)
        .BorderColor = IIf(asGrayed, SC_GRAYED_COLOR, SC_WHITE_COLOR)
        .BorderShade = IIf(asGrayed, SC_GRAYED_SHADE, SC_WHITE_SHADE)
        .BorderThemeColorIndex = IIf(asGrayed, SC_GRAYED_THEME, SC_WHITE_THEME)
        .BorderTint = IIf(asGrayed, SC_GRAYED_TINT, SC_WHITE_TINT)
    End With
    
    If styleLabel Then
        With TargetLabel
            .BackColor = IIf(asGrayed, SC_GRAYED_COLOR, SC_WHITE_COLOR)
            .BackShade = IIf(asGrayed, SC_GRAYED_SHADE, SC_WHITE_SHADE)
            .BackThemeColorIndex = IIf(asGrayed, SC_GRAYED_THEME, SC_WHITE_THEME)
            .BackTint = IIf(asGrayed, SC_GRAYED_TINT, SC_WHITE_TINT)
            .BorderColor = IIf(asGrayed, SC_GRAYED_COLOR, SC_WHITE_COLOR)
            .BorderShade = IIf(asGrayed, SC_GRAYED_SHADE, SC_WHITE_SHADE)
            .BorderThemeColorIndex = IIf(asGrayed, SC_GRAYED_THEME, SC_WHITE_THEME)
            .BorderTint = IIf(asGrayed, SC_GRAYED_TINT, SC_WHITE_TINT)
        End With
    End If
End Function

Public Function SetGridlineAsValid(ByRef TargetControl As Access.Control, Optional ByVal esValido As Boolean = True) As Boolean
    Dim TargetLabel As Access.Label
    On Error GoTo Finally
    
    If Not esValido Then
        If TargetControl.GridlineColor <> SC_INVALID_COLOR Then
            TargetControl.Tag = CStr(TargetControl.GridlineColor)
            TargetControl.GridlineColor = SC_INVALID_COLOR
            If TryGetAttachedLabel(TargetControl, TargetLabel) Then
                TargetLabel.GridlineColor = SC_INVALID_COLOR
            End If
        End If
    Else
        If TargetControl.GridlineColor = SC_INVALID_COLOR Then
            TargetControl.GridlineColor = CLng(TargetControl.Tag)
            If TryGetAttachedLabel(TargetControl, TargetLabel) Then
                TargetLabel.GridlineColor = CLng(TargetControl.Tag)
            End If
            TargetControl.Tag = ""
        End If
    End If
    
Finally:
    SetGridlineAsValid = esValido
End Function

Public Function SetControlAsEnabled(ByRef TargetControl As Access.Control, Optional ByVal Enable As Boolean = True) As Boolean
    SetControlAsEnabled = Enable
    SetGrayedStyle TargetControl, (Not Enable)
    On Error GoTo Fallback
    TargetControl.Locked = (Not Enable)
    Exit Function
Fallback:
    TargetControl.Enabled = Enable
End Function

Public Function TryGetAttachedLabel(ByRef TargetControl As Access.Control, ByRef AttachedLabel As Access.Label) As Boolean
    On Error GoTo MissingLabel
    
    If TypeName(TargetControl.Controls(0)) = "Label" Then
        Set AttachedLabel = TargetControl.Controls(0)
        TryGetAttachedLabel = True
        Exit Function
    End If

MissingLabel:
    TryGetAttachedLabel = False
End Function

Public Sub SetClickableStyle(ByRef TargetControl As Access.Control, ByVal Value As Boolean)
    On Error Resume Next
    If Value Then
        TargetControl.OnGotFocus = ""
        TargetControl.CursorOnHover = acCursorOnHoverHyperlinkHand
    Else
        TargetControl.OnGotFocus = "=Focus([Form])"
        TargetControl.CursorOnHover = acCursorOnHoverDefault
    End If
End Sub

Public Sub SetThemeColorStyle(ByRef TargetControl As Access.Control, ByVal ColorIndex As Long)
    On Error Resume Next
    With TargetControl
        .BackThemeColorIndex = ColorIndex
        .HoverThemeColorIndex = ColorIndex
        .BorderThemeColorIndex = ColorIndex
        .PressedThemeColorIndex = ColorIndex
    End With
End Sub


' --- UTILITY HELPERS ---

Private Function Max(X As Variant, Y As Variant) As Variant: Max = IIf(X > Y, X, Y): End Function
Private Function Min(X As Variant, Y As Variant) As Variant: Min = IIf(X < Y, X, Y): End Function




