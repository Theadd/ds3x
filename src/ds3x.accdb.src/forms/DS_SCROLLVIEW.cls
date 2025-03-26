Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "ds3x.UI.Scrollview"
Option Compare Database
Option Explicit

Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer


' --- PUBLIC EVENTS ---

Public Event OnRowNumberClick(ByVal RowIndex As Long, ByVal CtrlKey As Boolean, ByVal ShiftKey As Boolean)
Public Event OnColumnLetterClick(ByVal ColumnIndex As Long, ByVal CtrlKey As Boolean, ByVal ShiftKey As Boolean)
Public Event OnColumnNameChange(ByVal ColumnIndex As Long, ByVal FromValue As String, ByVal ToValue As String)
Public Event OnCellEnter(ByVal RowIndex As Long, ByVal ColumnIndex As Long, ByVal CtrlKey As Boolean, ByVal ShiftKey As Boolean)
Public Event OnClearSelectionRequest()
Public Event OnSelectAllRequest()
Public Event OnInvertSelectionRequest()
Public Event OnSelectionControlKeyDown(ByRef KeyCode As Integer, ByRef Shift As Integer)
Public Event OnWindowClose(ByRef Cancel As Integer)


' --- PRIVATE ---

Private pKeepScrollPositionOnTableChange As Boolean
Private pEnableOutOfRangeScrolling As Boolean
Private pInvertScrollInputAxis As Boolean

Private pTable As dsTable
Private pSelectedColumns As ArrayListEx
Private pSelectedRows As ArrayListEx

' The amount of overflowing scroll to the right after the last table column
Private pOutOfBoundsScrollX As Long
Private pOutOfBoundsScrollY As Long
Private pScrollPageSizeX As Long
Private pScrollPageSizeY As Long
Private pInitialized As Boolean
Private pReady As Boolean
Private pLastScrollX As Variant
Private pLastScrollY As Variant
Private pPointerCapture As Boolean
Private pCapturedPointerPosition As ds3xGlobals.POINTAPI
Private pScrollSpeedMultiplier As Long
Private pIgnoreScrollingEvents As Boolean

Private pMaxContentSizeX As Double
Private pMaxContentSizeY As Double
Private pLowScrollModX As Long
Private pLowScrollModY As Long
Private pScrollFactorPosX As Double
Private pScrollFactorPosY As Double
Private pWorksheetHeadersSizeY As Long
Private pIsScrollXQueueListening As Boolean
Private pIsScrollXQueueTriggering As Boolean
Private pIsScrollXInQueue As Boolean
Private pIsScrollYQueueListening As Boolean
Private pIsScrollYQueueTriggering As Boolean
Private pIsScrollYInQueue As Boolean

Private pViewport As Form_DS_VIEWPORT
Private pViewportSticky As Form_DS_VIEWPORT_STICKY
Private WithEvents pWorksheet As Form_DS_WORKSHEET
Attribute pWorksheet.VB_VarHelpID = -1
Private WithEvents pWorksheetHeaders As Form_DS_WORKSHEET_HEADERS
Attribute pWorksheetHeaders.VB_VarHelpID = -1
Private WithEvents pWorksheetNumbers As Form_DS_WORKSHEET_NUMBERS
Attribute pWorksheetNumbers.VB_VarHelpID = -1


' --- PUBLIC PROPERTIES ---

Public Property Get Table() As dsTable: Set Table = pTable: End Property
Public Property Set Table(ByRef Value As dsTable): SetTable Value: End Property

Public Property Get SelectedColumns() As Variant: SelectedColumns = pSelectedColumns.ToArray(): End Property
Public Property Let SelectedColumns(ByVal Value As Variant): SetSelectedColumns Value: End Property

Public Property Get SelectedRows() As Variant: SelectedRows = pSelectedRows.ToArray(): End Property
Public Property Let SelectedRows(ByVal Value As Variant): SetSelectedRows Value: End Property

Public Property Get KeepScrollPositionOnTableChange() As Boolean: KeepScrollPositionOnTableChange = pKeepScrollPositionOnTableChange: End Property
Public Property Let KeepScrollPositionOnTableChange(ByVal Value As Boolean): pKeepScrollPositionOnTableChange = Value: End Property
Public Property Get EnableOutOfRangeScrolling() As Boolean: EnableOutOfRangeScrolling = pEnableOutOfRangeScrolling: End Property
Public Property Let EnableOutOfRangeScrolling(ByVal Value As Boolean): pEnableOutOfRangeScrolling = Value: End Property
Public Property Get InvertScrollInputAxis() As Boolean: InvertScrollInputAxis = pInvertScrollInputAxis: End Property
Public Property Let InvertScrollInputAxis(ByVal Value As Boolean): pInvertScrollInputAxis = Value: End Property


' --- INTERNAL PROPERTIES ---

Public Property Get Worksheet() As Form_DS_WORKSHEET: Set Worksheet = pWorksheet: End Property
Public Property Set Worksheet(ByRef Value As Form_DS_WORKSHEET): Set pWorksheet = Value: End Property

Public Property Get WorksheetHeaders() As Form_DS_WORKSHEET_HEADERS: Set WorksheetHeaders = pWorksheetHeaders: End Property
Public Property Set WorksheetHeaders(ByRef Value As Form_DS_WORKSHEET_HEADERS): Set pWorksheetHeaders = Value: End Property

Public Property Get WorksheetNumbers() As Form_DS_WORKSHEET_NUMBERS: Set WorksheetNumbers = pWorksheetNumbers: End Property
Public Property Set WorksheetNumbers(ByRef Value As Form_DS_WORKSHEET_NUMBERS): Set pWorksheetNumbers = Value: End Property

Public Property Get Viewport() As Form_DS_VIEWPORT: Set Viewport = pViewport: End Property
Public Property Set Viewport(ByRef Value As Form_DS_VIEWPORT): Set pViewport = Value: End Property

Public Property Get ViewportSticky() As Form_DS_VIEWPORT_STICKY: Set ViewportSticky = pViewportSticky: End Property
Public Property Set ViewportSticky(ByRef Value As Form_DS_VIEWPORT_STICKY): Set pViewportSticky = Value: End Property

Public Property Get IsSubform() As Boolean: On Error Resume Next: IsSubform = Len(Me.Parent.Name) > 0: On Error GoTo 0: End Property
Public Property Get Initialized() As Boolean: Initialized = pInitialized: End Property

Public Property Get OutOfBoundsScrollX() As Long: OutOfBoundsScrollX = pOutOfBoundsScrollX: End Property
Public Property Let OutOfBoundsScrollX(ByVal Value As Long): pOutOfBoundsScrollX = Value: End Property
Public Property Get OutOfBoundsScrollY() As Long: OutOfBoundsScrollY = pOutOfBoundsScrollY: End Property
Public Property Let OutOfBoundsScrollY(ByVal Value As Long): pOutOfBoundsScrollY = Value: End Property

Public Property Get ScrollPageSizeX() As Long: ScrollPageSizeX = pScrollPageSizeX: End Property
Public Property Get ScrollPageSizeY() As Long: ScrollPageSizeY = pScrollPageSizeY: End Property

Public Property Get PointerCapture() As Boolean: PointerCapture = pPointerCapture: End Property
Public Property Let PointerCapture(ByVal Value As Boolean): SetPointerCapture Value: End Property

Friend Property Get IgnoreScrollingEvents() As Boolean: IgnoreScrollingEvents = pIgnoreScrollingEvents: End Property
Friend Property Let IgnoreScrollingEvents(ByVal Value As Boolean): pIgnoreScrollingEvents = Value: End Property

Public Property Get ScrollPosY() As Double
    ScrollPosY = (CDbl(Nz(Me.SCROLLBAR_Y, 0)) * pScrollFactorPosY) + pLowScrollModY
End Property

Public Property Let ScrollPosY(ByVal Value As Double)
    Dim t As Double, nMod As Double
    nMod = pLowScrollModY
    t = CDbl(Fix(Value / pScrollFactorPosY))
    On Error GoTo Fallback
    pLowScrollModY = CLng(Value - (t * pScrollFactorPosY))
    Me.SCROLLBAR_Y = CLng(t)
    Exit Property
Fallback:
    pLowScrollModY = nMod
End Property

Public Property Get ScrollPosX() As Double
    ScrollPosX = (CDbl(Nz(Me.SCROLLBAR_X, 0)) * pScrollFactorPosX) + pLowScrollModX
End Property

Public Property Let ScrollPosX(ByVal Value As Double)
    Dim t As Double, nMod As Double
    nMod = pLowScrollModX
    t = CDbl(Fix(Value / pScrollFactorPosX))
    On Error GoTo Fallback
    pLowScrollModX = CLng(Value - (t * pScrollFactorPosX))
    Me.SCROLLBAR_X = CLng(t)
    Exit Property
Fallback:
    pLowScrollModX = nMod
End Property

Private Sub QueueDelayedScrollX()
    On Error Resume Next
    If pIsScrollXQueueListening Then pIsScrollXInQueue = True
    If pIsScrollXInQueue Then
        If pIsScrollXQueueTriggering Then
            Do While pIsScrollXInQueue
                pIsScrollXQueueTriggering = False
                pIsScrollXQueueListening = False
                DoEvents
                pIsScrollXQueueListening = True
                pIsScrollXInQueue = False
                ApplyScrollbarX
                DoEvents
            Loop
            pIsScrollXQueueTriggering = True
            pIsScrollXQueueListening = True
        End If
    End If
    On Error GoTo 0
End Sub

Private Sub QueueDelayedScrollY()
    On Error Resume Next
    If pIsScrollYQueueListening Then pIsScrollYInQueue = True
    If pIsScrollYInQueue Then
        If pIsScrollYQueueTriggering Then
            Do While pIsScrollYInQueue
                pIsScrollYQueueTriggering = False
                pIsScrollYQueueListening = False
                DoEvents
                pIsScrollYQueueListening = True
                pIsScrollYInQueue = False
                ApplyScrollbarY
                DoEvents
            Loop
            pIsScrollYQueueTriggering = True
            pIsScrollYQueueListening = True
        End If
    End If
    On Error GoTo 0
End Sub

' --- FORM EVENTS ---

Private Sub Form_Load()
    ScreenLib.ScreenLib_Resync
    pScrollSpeedMultiplier = 1
    pLastScrollX = Array(0, 0)
    pLastScrollY = Array(0, 0)
    pScrollFactorPosX = 1
    pScrollFactorPosY = 1
    pIsScrollYQueueListening = True
    pIsScrollYQueueTriggering = True
    pIsScrollXQueueListening = True
    pIsScrollXQueueTriggering = True
    Set pSelectedColumns = ArrayListEx.Create()
    Set pSelectedRows = ArrayListEx.Create()
    Setup

    If Not IsSubform Then
        ScreenLib.WindowSizeTo Me, 12000, 8000
        ScreenLib.WindowCenterTo Me, ScreenLib.GetScreenRectOfPoint(ScreenLib.PointInRect(ScreenLib.GetWindowRect(Me), DirectionType.Center))

        SetupDevelopmentEnvironment
    End If
    Me.TimerInterval = 1
End Sub

Private Sub Form_Resize()
    ResizeFormContent
    If pInitialized Then
        On Error GoTo Finally
        pIgnoreScrollingEvents = True
        UpdateScrollbarX False
        UpdateScrollbarY False
        Viewport.ScrollTo ScrollPosX, ScrollPosY
Finally:
        pIgnoreScrollingEvents = False
    End If
End Sub

Private Sub Form_Timer()
    Me.TimerInterval = 0
    If pPointerCapture Then
        ScrollUsingLastCapturedPointerPosition
        If pPointerCapture Then Me.TimerInterval = 10
    End If
    If Not pReady Then
        pReady = True
        ApplyScrollbarY 0, True
        DoEvents
        ApplyScrollbarY 0, True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Me.TimerInterval = 0
    RaiseEvent OnWindowClose(Cancel)
    If Not CBool(Cancel) Then Dispose
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer): OnKeyDownHandler KeyCode, Shift: End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer): OnKeyUpHandler KeyCode, Shift: End Sub


' --- SETUP / BINDING ---

Private Sub Setup()
    Set Viewport = Me.DS_VIEWPORT.Form
    Set pViewport.Scrollview = Me
    Set ViewportSticky = Me.DS_VIEWPORT_STICKY.Form
    Set pViewportSticky.Scrollview = Me
    pViewportSticky.Setup
    Set WorksheetNumbers = pViewportSticky.WorksheetNumbers
    Set pViewport.WorksheetNumbers = pWorksheetNumbers
    pViewport.Setup
    Set Worksheet = pViewport.Worksheet
    Set WorksheetHeaders = pViewport.WorksheetHeaders
    pWorksheetHeadersSizeY = pViewport.FormHeaders.Height
End Sub

Private Sub SetTable(ByRef Value As dsTable)
    Set pTable = Value
    pInitialized = Not (pTable Is Nothing)

    If pInitialized Then
        On Error GoTo Finally
        pIgnoreScrollingEvents = True
        UpdateScrollbarX
        UpdateScrollbarY True
Finally:
        pIgnoreScrollingEvents = False
    End If
    On Error Resume Next
    Viewport.OnSourceTableChange
    On Error GoTo 0
End Sub

Private Sub SetSelectedColumns(ByVal Value As Variant)
    pSelectedColumns.Clear
    pSelectedColumns.AddRange Value
    pWorksheetHeaders.SetSelectedColumns pSelectedColumns
End Sub

Private Sub SetSelectedRows(ByVal Value As Variant)
    pSelectedRows.Clear
    pSelectedRows.AddRange Value
    pWorksheetNumbers.SetSelectedRows pSelectedRows
End Sub

Private Sub Dispose()
    On Error Resume Next
    Set pWorksheetNumbers.Viewport = Nothing
    Set pWorksheetHeaders.Viewport = Nothing
    Set pWorksheet.Viewport = Nothing
    Set pViewportSticky.WorksheetNumbers = Nothing
    Set pViewportSticky.Scrollview = Nothing
    Set pViewport.Worksheet = Nothing
    Set pViewport.WorksheetHeaders = Nothing
    Set pViewport.WorksheetNumbers = Nothing
    Set pViewport.Scrollview = Nothing
    Set pWorksheetNumbers = Nothing
    Set pWorksheetHeaders = Nothing
    Set pWorksheet = Nothing
    Set pViewportSticky = Nothing
    Set pViewport = Nothing
    On Error GoTo 0
End Sub


' --- CONTENT RESIZING ---

Private Sub ResizeFormContent()
    Dim mW As Long, mH As Long, c As Long, t As Long

    mW = Me.InsideWidth
    mH = Me.InsideHeight + 15
    c = Me.DS_VIEWPORT_STICKY.Form.DS_WORKSHEET_NUMBERS.Form.MaxContentWidthLimit

    With Me.SCROLLBAR_X
        .Height = 270
        .Left = 0
        .Width = mW - 270
    End With
    With Me.DS_VIEWPORT
        .Height = mH - (270 + .Top)
        .Width = mW - (270 + c)
        .Left = c
        Me.SCROLLBAR_X.Top = .Height + .Top
    End With
    With Me.DS_VIEWPORT_STICKY
        .Left = 0
        .Width = c
        .Height = Me.DS_VIEWPORT.Height
    End With
    With Me.SCROLLBAR_Y
        .Height = mH - (270 + .Top)
        .Width = 270
        .Left = mW - (270 + c)
    End With
End Sub


' --- SCROLLING ---

Private Sub ScrollUsingLastCapturedPointerPosition()
    Dim p As ds3xGlobals.POINTAPI, x As Long, Y As Long

    ScreenLib.MouseMoveCursor = True
    p = ScreenLib.GetCursorPosition
    x = p.x - pCapturedPointerPosition.x
    Y = p.Y - pCapturedPointerPosition.Y
    ScreenLib.SetCursorPosition pCapturedPointerPosition
    ScrollTo CDbl(Max(ScrollPosX + CDbl(x * pScrollSpeedMultiplier), 0)), CDbl(Max(ScrollPosY + CDbl(Y * pScrollSpeedMultiplier), 0))
End Sub

Public Sub ScrollTo(ByVal x As Double, ByVal Y As Double)
    On Error GoTo Finally
    pIgnoreScrollingEvents = True
    ApplyScrollbarX x
    ApplyScrollbarY Y, True
Finally:
    pIgnoreScrollingEvents = False
End Sub

Public Function PropagateMouseWheel(ByVal Page As Boolean, ByVal Count As Long)
    Dim sAxisX As Boolean, sMod As Long
    sAxisX = pInvertScrollInputAxis

    sMod = IIf(GetAsyncKeyState(vbKeyShift), 3, 1)
    sAxisX = IIf(GetAsyncKeyState(vbKeyControl) And GetAsyncKeyState(vbKeyControl), Not sAxisX, sAxisX)

    If sAxisX Then
        ScrollPosX = ScrollPosX + CDbl(Count * sMod * Fix(CDbl(Worksheet.GridCellSizeX) / 5#))
    Else
        ScrollPosY = ScrollPosY + CDbl(Count * sMod * Worksheet.GridCellSizeY)
    End If
End Function

Private Sub SCROLLBAR_X_Change(): QueueDelayedScrollX: End Sub
Private Sub SCROLLBAR_X_Scroll(): QueueDelayedScrollX: End Sub
Private Sub SCROLLBAR_Y_Change(): QueueDelayedScrollY: End Sub
Private Sub SCROLLBAR_Y_Scroll(): QueueDelayedScrollY: End Sub

Public Sub ApplyScrollbarY(Optional ByVal Value As Variant, Optional ByVal ExplicitCall As Boolean = False)
    If (Not Me.SCROLLBAR_Y.Visible) Or (Not pInitialized) Then Exit Sub
    Dim rawMax As Long, yVal As Double, nY As Double

    With Me.SCROLLBAR_Y
        nY = ScrollPosY
        yVal = CDbl(IIf(IsMissing(Value), nY, CDbl(Value)) / pScrollFactorPosY)
        If IsMissing(Value) Then Value = nY
        If yVal < 0 Then yVal = 0
        rawMax = .Max - pOutOfBoundsScrollY
        If yVal > .Max Then
            If pEnableOutOfRangeScrolling Then
                pOutOfBoundsScrollY = yVal - rawMax
                .Max = yVal
                .Value = yVal
            Else
                .Value = .Max
            End If
            UpdateScrollbarY ExplicitCall
        Else
            If pOutOfBoundsScrollY > 0 Then
                pOutOfBoundsScrollY = IIf(yVal - rawMax > 0, yVal - rawMax, 0)
                If .Max <> rawMax + pOutOfBoundsScrollY Then .Max = rawMax + pOutOfBoundsScrollY
                If nY <> CDbl(Value) Then ScrollPosY = CDbl(Value)
                UpdateScrollbarY ExplicitCall
            Else
                If nY <> CDbl(Value) Then ScrollPosY = CDbl(Value)
                If (Not pIgnoreScrollingEvents) Or ExplicitCall Then
                    nY = ScrollPosY
                    Viewport.ScrollTo ScrollPosX, nY
                    pLastScrollY(0) = nY
                End If
            End If
        End If
    End With
End Sub

Public Sub ApplyScrollbarX(Optional ByVal Value As Variant, Optional ByVal ExplicitCall As Boolean = False)
    If (Not Me.SCROLLBAR_X.Visible) Or (Not pInitialized) Then Exit Sub
    Dim rawMax As Long, xVal As Double, nX As Double

    With Me.SCROLLBAR_X
        nX = ScrollPosX
        xVal = CDbl(IIf(IsMissing(Value), nX, CDbl(Value)) / pScrollFactorPosX)
        If IsMissing(Value) Then Value = nX
        If xVal < 0 Then xVal = 0
        rawMax = .Max - pOutOfBoundsScrollX
        If xVal > .Max Then
            If pEnableOutOfRangeScrolling Then
                pOutOfBoundsScrollX = xVal - rawMax
                .Max = xVal
                .Value = xVal
            Else
                .Value = .Max
            End If
            UpdateScrollbarX ExplicitCall
        Else
            If pOutOfBoundsScrollX > 0 Then
                pOutOfBoundsScrollX = IIf(xVal - rawMax > 0, xVal - rawMax, 0)
                If .Max <> rawMax + pOutOfBoundsScrollX Then .Max = rawMax + pOutOfBoundsScrollX
                If nX <> CDbl(Value) Then ScrollPosX = CDbl(Value)
                UpdateScrollbarX ExplicitCall
            Else
                If nX <> CDbl(Value) Then ScrollPosX = CDbl(Value)
                If (Not pIgnoreScrollingEvents) Or ExplicitCall Then
                    nX = ScrollPosX
                    Viewport.ScrollTo nX, ScrollPosY
                    pLastScrollX(0) = nX
                End If
            End If
        End If
    End With
End Sub

Private Sub UpdateScrollbarY(Optional ByVal ExplicitCall As Boolean = False)
    If Not Me.SCROLLBAR_Y.Visible Then Exit Sub
    Dim yMax As Long, cellSizeY As Long, nY As Double

    cellSizeY = Worksheet.GridCellSizeY
    pScrollPageSizeY = CLng(Int((Me.InsideHeight - 270) / cellSizeY)) * cellSizeY
    pMaxContentSizeY = CDbl(pTable.Count * CDbl(cellSizeY)) + CDbl(pWorksheetHeadersSizeY)
    pScrollFactorPosY = IIf(pMaxContentSizeY + CDbl(OutOfBoundsScrollY) > 2 ^ 30, 1000, 1)
    yMax = CLng(Fix(Max((pMaxContentSizeY + CDbl(OutOfBoundsScrollY)) - CDbl(Me.InsideHeight - 270), 0) / pScrollFactorPosY))

    With Me.SCROLLBAR_Y
        .Max = yMax
        .LargeChange = IIf(CLng(pScrollPageSizeY / pScrollFactorPosY) > yMax, yMax, CLng(pScrollPageSizeY / pScrollFactorPosY))
        .SmallChange = IIf(CLng(cellSizeY / pScrollFactorPosY) < 1, 1, CLng(cellSizeY / pScrollFactorPosY))
        If (Not pIgnoreScrollingEvents) Or ExplicitCall Then
            nY = ScrollPosY
            If nY <> pLastScrollY(0) Or yMax <> pLastScrollY(1) Then
                Viewport.ScrollTo ScrollPosX, nY
                pLastScrollY(0) = nY
                pLastScrollY(1) = yMax
            End If
        End If
    End With
End Sub

Private Sub UpdateScrollbarX(Optional ByVal ExplicitCall As Boolean = False)
    If Not Me.SCROLLBAR_X.Visible Then Exit Sub

    ' ____________* ScrollView.ScrollPageSizeX _______* ScrollView.OutOfBoundsScrollX
    '|____________|___________________________|_______|
    '|________________________________________|
    '                                         * Viewport.ViewportContentFullWidth

    Dim xMax As Long, cellSizeX As Long, viewSizeX As Long, nX As Double

    viewSizeX = Me.InsideWidth - 270 - Me.DS_VIEWPORT.Left
    cellSizeX = Worksheet.GridCellSizeX
    pScrollPageSizeX = CLng(Int(viewSizeX / cellSizeX)) * cellSizeX
    pMaxContentSizeX = Max(CDbl(pTable.ColumnCount * CDbl(cellSizeX)), viewSizeX)
    pScrollFactorPosX = IIf(pMaxContentSizeX + CDbl(OutOfBoundsScrollX) > 2 ^ 30, 20000, 1)
    xMax = CLng(Fix(Max((pMaxContentSizeX + CDbl(OutOfBoundsScrollX)) - CDbl(viewSizeX), 0) / pScrollFactorPosX))

    With Me.SCROLLBAR_X
        .Max = xMax
        .LargeChange = IIf(CLng(pScrollPageSizeX / pScrollFactorPosX) > xMax, xMax, CLng(pScrollPageSizeX / pScrollFactorPosX))
        .SmallChange = IIf(CLng(CLng(cellSizeX / 5) / pScrollFactorPosX) < 1, 1, CLng(CLng(cellSizeX / 5) / pScrollFactorPosX))
        If (Not pIgnoreScrollingEvents) Or ExplicitCall Then
            nX = ScrollPosX
            If nX <> pLastScrollX(0) Or xMax <> pLastScrollX(1) Then
                Viewport.ScrollTo nX, ScrollPosY
                pLastScrollX(0) = nX
                pLastScrollX(1) = xMax
            End If
        End If
    End With
End Sub


' --- SELECTION ---

Public Sub MoveTo(ByVal RowIndex As Long, ByVal ColumnIndex As Long, Optional ByVal PropagateEvent As Boolean = True)
    Dim x As Double, Y As Double

    Select Case CLng(0 - 1)
        Case RowIndex
            Y = ScrollPosY
            x = pViewport.GetScrollXTo(ColumnIndex)
        Case ColumnIndex
            x = ScrollPosX
            Y = pViewport.GetScrollYTo(RowIndex)
        Case Else
            x = pViewport.GetScrollXTo(ColumnIndex)
            Y = pViewport.GetScrollYTo(RowIndex)
    End Select

    ScrollTo x, Y

    If PropagateEvent Then
        Select Case CLng(0 - 1)
            Case RowIndex
                RaiseEvent OnColumnLetterClick(ColumnIndex, False, False)
            Case ColumnIndex
                RaiseEvent OnRowNumberClick(RowIndex, False, False)
            Case Else
                RaiseEvent OnCellEnter(RowIndex, ColumnIndex, False, False)
        End Select
    End If
End Sub

Public Sub TriggerClickOnSelectAll()
    RaiseEvent OnSelectAllRequest
End Sub

Public Sub TriggerClickOnInvertSelection()
    RaiseEvent OnInvertSelectionRequest
End Sub


' --- EVENTS ---

Private Sub pWorksheetHeaders_OnColumnLetterClick(ByVal ColumnIndex As Long, ByVal CtrlKey As Boolean, ByVal ShiftKey As Boolean)
    pWorksheetHeaders.RemoveFocus
    RaiseEvent OnColumnLetterClick(ColumnIndex, GetAsyncKeyState(vbKeyControl), GetAsyncKeyState(vbKeyShift))
End Sub

Private Sub pWorksheetHeaders_OnColumnNameWillChange(ByVal ColumnIndex As Long, ByVal Value As String)
    Dim oldValue As String

    On Error GoTo Finally
    oldValue = CollectionsLib.ArrayItem(Table().Headers.Row(0)(ColumnIndex), 0)
    If StrComp(oldValue, Value, vbBinaryCompare) <> 0 Then
        On Error GoTo 0
        RaiseEvent OnColumnNameChange(ColumnIndex, oldValue, Value)
    End If
Finally:
End Sub

Private Sub pWorksheetNumbers_OnRowNumberClick(ByVal RowIndex As Long, ByVal CtrlKey As Boolean, ByVal ShiftKey As Boolean)
    pWorksheetHeaders.RemoveFocus
    RaiseEvent OnRowNumberClick(RowIndex, GetAsyncKeyState(vbKeyControl), GetAsyncKeyState(vbKeyShift))
End Sub

Private Sub pWorksheet_OnCellEnter(ByVal RowIndex As Long, ByVal ColumnIndex As Long, ByVal CtrlKey As Boolean, ByVal ShiftKey As Boolean)
    pWorksheetHeaders.RemoveFocus
    RaiseEvent OnCellEnter(RowIndex, ColumnIndex, GetAsyncKeyState(vbKeyControl), GetAsyncKeyState(vbKeyShift))
End Sub


' --- EVENTS: Internal Event Actions

Public Sub OnKeyDownHandler(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyUp, vbKeyRight, vbKeyDown, vbKeyLeft, vbKeyPageUp, vbKeyPageDown, vbKeyHome, vbKeyEnd, vbKeyEscape, vbKeyTab, vbKeyReturn
            RaiseEvent OnSelectionControlKeyDown(KeyCode, Shift)
    End Select

    Select Case True
        Case (KeyCode = 0): ' Ignore
        Case (KeyCode = vbKeyDown And Shift = 2): KeyCode = 0
        Case (KeyCode = vbKeyP And Shift = 3): Stop
        Case (KeyCode = vbKeyShift): pScrollSpeedMultiplier = 3
        Case (KeyCode = vbKeyC And Shift = 2): CopySelectionToClipboard
        Case (KeyCode = vbKeyPageDown): KeyCode = 0
        Case (KeyCode = vbKeyPageUp): KeyCode = 0
        Case Else
            SetPointerCapture (KeyCode = vbKeySpace And (Shift = 2 Or Shift = 3))
    End Select
End Sub

Public Sub OnKeyUpHandler(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyShift Then
        pScrollSpeedMultiplier = 1
    ElseIf pPointerCapture Then
        SetPointerCapture False
    End If
End Sub

Public Sub TriggerOnClearSelectionRequest()
    RaiseEvent OnClearSelectionRequest
End Sub

Private Sub SetPointerCapture(ByVal Value As Boolean)
    If pPointerCapture = Value Then Exit Sub
    Me.TimerInterval = 0
    pPointerCapture = Value
    If Value Then
        pCapturedPointerPosition = ScreenLib.GetCursorPosition
        Me.TimerInterval = 100
    End If
End Sub

' TODO: IncludeHeaders
Private Sub CopySelectionToClipboard(Optional ByVal IncludeHeaders As Boolean = False)
    Const dsChunkSize As Long = 1000
    Dim i As Long, iMax As Long, nRows As Long, rStart As Long, rEnd As Long, rTake As Long, rTakeFrom As Long, aX As ArraySliceGroup, sCols As Variant, sValues As String
    On Error GoTo Finally
    pSelectedColumns.Sort
    sCols = Empty
    If pSelectedColumns.Count > 0 Then sCols = pSelectedColumns.ToArray()

    If pSelectedRows.Count > 0 Then
        pSelectedRows.Sort
        rStart = pSelectedRows(0)
        rEnd = pSelectedRows(pSelectedRows.Count - 1)
        nRows = 1 + rEnd - rStart
    Else
        rStart = 0
        nRows = pTable.Count
        rEnd = 0
    End If
    iMax = Int(CLng(nRows - 1) / dsChunkSize)

    For i = 0 To iMax
        rTake = IIf(i = iMax, nRows - (dsChunkSize * i), dsChunkSize)
        rTakeFrom = rStart + (i * dsChunkSize)
        Set aX = CollectionsLib.AsArraySliceGroup(pTable.Records.GetRange(rTakeFrom, rTake, sCols))
        If rEnd = 0 Then
            sValues = sValues & aX.ToExcel()
        Else
            sValues = sValues & aX.GetRows(pSelectedRows, 0 - rTakeFrom).ToExcel()
        End If
    Next i
Finally:
    FileSystemLib.SystemClipboard sValues
End Sub


' --- HELPERS ---

Private Function Max(x As Variant, Y As Variant) As Variant: Max = IIf(x > Y, x, Y): End Function
Private Function Min(x As Variant, Y As Variant) As Variant: Min = IIf(x < Y, x, Y): End Function


' --- TESTING / DEVELOPMENT ---

Private Sub SetupDevelopmentEnvironment()
    pEnableOutOfRangeScrolling = True
    Set Table = dsTable.CreateBlank(10, 0)
End Sub
