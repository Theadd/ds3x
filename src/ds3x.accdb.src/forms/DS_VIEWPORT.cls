Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "ds3x.UI.Scrollview"
Option Compare Database
Option Explicit


Private pWorksheet As Form_DS_WORKSHEET
Private pWorksheetHeaders As Form_DS_WORKSHEET_HEADERS
Private pWorksheetNumbers As Form_DS_WORKSHEET_NUMBERS
Private pScrollview As Form_DS_SCROLLVIEW

' The current ColumnsToLargeChangeTrack value used in CachedTracks, CachedTracks resets when this value changes
Private pTrackColumnSizesInCache As Long
Const NumPagesInLargeChangeRows As Long = 5
Const PageSize As Long = 10
Const PageCount As Long = 10

Private Type TViewportState
    ScrollPosX As Double
    ScrollPosY As Double
    ' Index of current visible track in Viewport
    TrackIndex As Long
    ' Index of current visible page in Viewport
    PageIndex As Long
    ' Index of the first visible column in **Table**
    FirstVisibleColumn As Long
    ' Index of the first visible row in **Table**
    FirstVisibleRow As Long
    ' The distance between the start of the first visible column to the viewport left edge (must be less than GridCellSizeX)
    FirstVisibleColumnPositionModX As Long
    ' The distance between the start of the first visible row to the viewport top edge (must be less than GridCellSizeY)
    FirstVisibleRowPositionModY As Long
    ' Index of the first visible column relative to current visible **Track**
    FirstVisibleColumnInTrack As Long
    ' Index of the first visible row relative to current visible **Page**
    FirstVisibleRowInPage As Long
    ' Number of columns as the distance between track switching
    ColumnsToLargeChangeTrack As Long
    ' The distance from the current track left edge to the viewport left edge
    TrackPositionModX As Long
    ' The distance from the current page top edge to the viewport top edge
    PagePositionModY As Long
End Type

Private This As TViewportState

' ---

Public Property Get Worksheet() As Form_DS_WORKSHEET: Set Worksheet = pWorksheet: End Property
Public Property Set Worksheet(ByRef Value As Form_DS_WORKSHEET): Set pWorksheet = Value: End Property

Public Property Get WorksheetHeaders() As Form_DS_WORKSHEET_HEADERS: Set WorksheetHeaders = pWorksheetHeaders: End Property
Public Property Set WorksheetHeaders(ByRef Value As Form_DS_WORKSHEET_HEADERS): Set pWorksheetHeaders = Value: End Property

Public Property Get WorksheetNumbers() As Form_DS_WORKSHEET_NUMBERS: Set WorksheetNumbers = pWorksheetNumbers: End Property
Public Property Set WorksheetNumbers(ByRef Value As Form_DS_WORKSHEET_NUMBERS): Set pWorksheetNumbers = Value: End Property

Public Property Get Scrollview() As Form_DS_SCROLLVIEW: Set Scrollview = pScrollview: End Property
Public Property Set Scrollview(ByRef Value As Form_DS_SCROLLVIEW): Set pScrollview = Value: End Property

Friend Property Get FirstColumnIndex() As Long: FirstColumnIndex = This.TrackIndex * This.ColumnsToLargeChangeTrack: End Property

Property Get IsSubform() As Boolean: On Error Resume Next: IsSubform = Len(Me.Parent.Name) > 0: On Error GoTo 0: End Property


' --- FORM EVENTS ---

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Me.TimerInterval = 0
    On Error GoTo 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer): pScrollview.OnKeyDownHandler KeyCode, Shift: End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer): pScrollview.OnKeyUpHandler KeyCode, Shift: End Sub


' --- SETUP / BINDING ---

Public Sub Setup()
    Dim r As ds3xGlobals.RECT, B As ds3xGlobals.BOUNDS, t As Long, c As Long

    Set Worksheet = Me.DS_WORKSHEET.Form
    Set pWorksheet.Viewport = Me
    Set WorksheetHeaders = Me.DS_WORKSHEET_HEADERS.Form
    Set pWorksheetHeaders.Viewport = Me
    Set pWorksheetNumbers.Viewport = Me
    pWorksheet.Setup
    pWorksheetHeaders.Setup
    pWorksheetNumbers.Setup PageSize * PageCount

    r = ScreenLib.GetScreenRectOfPoint(ScreenLib.PointInRect(ScreenLib.GetWindowRect(Me), DirectionType.Center), True)
    B = ScreenLib.RectToBounds(r)

    t = pWorksheet.MaxContentWidthLimit
    Me.Width = t
    With Me.DS_WORKSHEET_HEADERS
        .Left = 0
        .Top = 0
        .Width = t
        Me.FormHeaders.Height = .Height
    End With
    Me.FormDetail.Height = CLng(Min(B.h * 1.95, 31500))
    With Me.DS_WORKSHEET
        .Left = 0
        .Top = 0
        .Width = t
        .Height = CLng(Min(B.h * 1.95, 31500))
    End With
End Sub


' --- SCROLLING ---

Public Function PropagateMouseWheel(ByVal Page As Boolean, ByVal Count As Long)
    On Error Resume Next
    pScrollview.PropagateMouseWheel Page, Count
    On Error GoTo 0
End Function

Public Sub ScrollTo(ByVal x As Double, ByVal Y As Double)
    Dim sView As TViewportState
    sView = GetViewportStateAt(x, Y)

    If This.TrackIndex <> sView.TrackIndex Or This.PageIndex <> sView.PageIndex Or pTrackColumnSizesInCache <> sView.ColumnsToLargeChangeTrack Then
        pWorksheet.Painting = False
        pWorksheetHeaders.Painting = False
        pWorksheetNumbers.Painting = False
        ScreenLib.WindowMoveTo pWorksheet, 0 - sView.TrackPositionModX, 0 - sView.PagePositionModY
        ScreenLib.WindowMoveTo pWorksheetHeaders, 0 - sView.TrackPositionModX, 0
        ScreenLib.WindowMoveTo pWorksheetNumbers, 0, 0 - sView.PagePositionModY
        pTrackColumnSizesInCache = sView.ColumnsToLargeChangeTrack
        Set pWorksheet.Recordset = GetTrack(sView.TrackIndex, sView.PageIndex).Instance
        pWorksheet.SetupGrid sView.TrackIndex * sView.ColumnsToLargeChangeTrack, sView.PageIndex * PageSize * NumPagesInLargeChangeRows, pScrollview.Table
        pWorksheetHeaders.SetupGrid sView.TrackIndex * sView.ColumnsToLargeChangeTrack, sView.PageIndex * PageSize * NumPagesInLargeChangeRows, pScrollview.Table
        pWorksheetNumbers.SetupGrid sView.TrackIndex * sView.ColumnsToLargeChangeTrack, sView.PageIndex * PageSize * NumPagesInLargeChangeRows, pScrollview.Table
        pWorksheetNumbers.Painting = True
        pWorksheetHeaders.Painting = True
        pWorksheet.Painting = True
    Else
        ScreenLib.WindowMoveTo pWorksheet, 0 - sView.TrackPositionModX, 0 - sView.PagePositionModY
        ScreenLib.WindowMoveTo pWorksheetHeaders, 0 - sView.TrackPositionModX, 0
        ScreenLib.WindowMoveTo pWorksheetNumbers, 0, 0 - sView.PagePositionModY
    End If
    This = sView
End Sub


' --- STATE MANAGEMENT ---

Private Function GetViewportStateAt(ByVal x As Double, ByVal Y As Double) As TViewportState
    Dim t As TViewportState, maxTrackWidth As Long, cellWidth As Long, viewWidth As Long, cellHeight As Long

    t.ScrollPosX = x
    t.ScrollPosY = Y

    maxTrackWidth = Worksheet.MaxContentWidthLimit
    cellWidth = Worksheet.GridCellSizeX
    cellHeight = Worksheet.GridCellSizeY
    viewWidth = Scrollview.ScrollPageSizeX

    t.ColumnsToLargeChangeTrack = Max(CLng((maxTrackWidth - viewWidth) / cellWidth) - 1, 1)

    t.FirstVisibleColumn = CLng(Int(x / CDbl(cellWidth)))
    t.FirstVisibleRow = CLng(Int(Y / CDbl(cellHeight)))
    t.FirstVisibleColumnPositionModX = CLng(ModFunc(x, CDbl(cellWidth)))
    t.FirstVisibleRowPositionModY = CLng(ModFunc(Y, CDbl(cellHeight)))
    t.TrackIndex = CLng(Int(x / CDbl(cellWidth * t.ColumnsToLargeChangeTrack)))
    t.PageIndex = CLng(Int(t.FirstVisibleRow / (PageSize * NumPagesInLargeChangeRows)))
    t.FirstVisibleColumnInTrack = t.FirstVisibleColumn - (t.ColumnsToLargeChangeTrack * t.TrackIndex)
    t.FirstVisibleRowInPage = t.FirstVisibleRow - (t.PageIndex * PageSize * NumPagesInLargeChangeRows)
    t.TrackPositionModX = (t.FirstVisibleColumnInTrack * cellWidth) + t.FirstVisibleColumnPositionModX
    t.PagePositionModY = (t.FirstVisibleRowInPage * cellHeight) + t.FirstVisibleRowPositionModY

    GetViewportStateAt = t
End Function

Public Function GetScrollXTo(ByVal ColumnIndex As Long) As Double
    Dim cellWidth As Long, viewWidth As Long, curViewMinX As Double, curViewMaxX As Double
    Dim targetCellMinX As Double, targetCellMaxX As Double, x As Double

    cellWidth = Worksheet.GridCellSizeX
    viewWidth = Scrollview.ScrollPageSizeX

    x = CDbl(cellWidth) * CDbl(ColumnIndex)
    targetCellMinX = Max(x - CDbl(cellWidth), 0)
    targetCellMaxX = x + CDbl(3 * cellWidth)

    curViewMinX = This.ScrollPosX
    curViewMaxX = curViewMinX + CDbl(viewWidth)

    x = This.ScrollPosX
    If curViewMinX > targetCellMinX Then
        x = targetCellMinX
    ElseIf curViewMaxX < targetCellMaxX Then
        x = targetCellMaxX - CDbl(viewWidth)
    End If

    GetScrollXTo = x
End Function

Public Function GetScrollYTo(ByVal RowIndex As Long) As Double
    Dim cellHeight As Long, viewHeight As Long, curViewMinY As Double, curViewMaxY As Double
    Dim targetCellMinY As Double, targetCellMaxY As Double, Y As Double

    cellHeight = Worksheet.GridCellSizeY
    viewHeight = Scrollview.ScrollPageSizeY

    Y = CDbl(cellHeight) * CDbl(RowIndex)
    targetCellMinY = Max(Y - CDbl(cellHeight), 0)
    targetCellMaxY = Y + CDbl(3 * cellHeight)

    curViewMinY = This.ScrollPosY
    curViewMaxY = curViewMinY + CDbl(viewHeight)

    Y = This.ScrollPosY
    If curViewMinY > targetCellMinY Then
        Y = targetCellMinY
    ElseIf curViewMaxY < targetCellMaxY Then
        Y = targetCellMaxY - CDbl(viewHeight)
    End If

    GetScrollYTo = Y
End Function

Public Sub OnSourceTableChange()
    This.TrackIndex = -1
    If pScrollview.KeepScrollPositionOnTableChange Then
        ScrollTo This.ScrollPosX, This.ScrollPosY
    Else
        On Error GoTo Finally
        pScrollview.IgnoreScrollingEvents = True
        pScrollview.ScrollPosX = 0#
        pScrollview.ScrollPosY = 0#
        ScrollTo 0#, 0#
Finally:
        pScrollview.IgnoreScrollingEvents = False
    End If
End Sub

Private Function GetTrack(ByVal TrackIndex As Long, ByVal PageIndex As Long) As RecordsetEx
    Dim dsT As dsTable, rX As RecordsetEx, ColumnStartIndex As Long, nCols As Long, dsT2 As dsTable, dsT3 As dsTable, nRows As Long, RowStartIndex As Long

    Set dsT = Scrollview.Table
    ColumnStartIndex = TrackIndex * pTrackColumnSizesInCache
    nCols = Worksheet.MaxAvailColumns

    If ColumnStartIndex >= dsT.ColumnCount Then
        Set rX = RecordsetEx.CreateBlank(PageSize * PageCount, nCols)
    Else
        RowStartIndex = Min(PageSize * PageIndex * NumPagesInLargeChangeRows, dsT.Count)
        nRows = Min(dsT.Count - RowStartIndex, PageSize * PageCount)
        If dsT.ColumnCount - ColumnStartIndex > nCols Then
            If nRows < PageSize * PageCount Then
                Set dsT2 = dsT.GetRange(RowStartIndex, nRows, CollectionsLib.ArrayRange(ColumnStartIndex, nCols))
                Set dsT2 = dsT2.AddRange(dsTable.CreateBlank((PageSize * PageCount) - nRows, nCols))
                Set rX = RecordsetEx.Create(dsT2.IndexRecordset)
            Else
                Set rX = RecordsetEx.Create(dsT.CreateIndexRecordset(PageSize, PageIndex * NumPagesInLargeChangeRows, PageCount, ColumnStartIndex, nCols, True))
            End If
        Else
            Set dsT2 = dsT.GetRange(RowStartIndex, nRows, CollectionsLib.ArrayRange(ColumnStartIndex, dsT.ColumnCount))
            Set dsT3 = dsTable.CreateBlank(nRows, nCols - (dsT.ColumnCount - ColumnStartIndex))
            Set dsT2 = dsT2.Join(dsT3)
            If nRows < PageSize * PageCount Then
                Set dsT2 = dsT2.AddRange(dsTable.CreateBlank((PageSize * PageCount) - nRows, nCols))
            End If
            Set rX = RecordsetEx.Create(dsT2.IndexRecordset)
        End If
    End If

    Set GetTrack = rX
End Function


' --- HELPERS ---

Private Function Max(x As Variant, Y As Variant) As Variant: Max = IIf(x > Y, x, Y): End Function
Private Function Min(x As Variant, Y As Variant) As Variant: Min = IIf(x < Y, x, Y): End Function
Private Function ModFunc(x As Variant, Y As Variant) As Variant: ModFunc = x - (Fix(x / Y) * Y): End Function


' --- DEBUG ---

Public Sub Debugger()
    Stop
End Sub
