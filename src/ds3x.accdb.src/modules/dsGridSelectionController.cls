VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "dsGridSelectionController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "ds3x.LiveEditor.UI"
Option Compare Database
Option Explicit
Option Base 0

Private WithEvents pScrollview As Form_DS_SCROLLVIEW
Attribute pScrollview.VB_VarHelpID = -1

Private pLastClickedColumnIndex As Long
Private pSelectedColumnIndexes As ArrayListEx
Private pLastClickedRowIndex As Long
Private pSelectedRowIndexes As ArrayListEx
' Cell as: Array(RowIndex, ColumnIndex)
Private pLastEnteredCell As Variant
Private pLastSelectionTrigger As Long

' Selection Triggers
Private Const pST_NONE As Long = 0
Private Const pST_COLUMN As Long = 1
Private Const pST_ROW As Long = 2
Private Const pST_CELL As Long = 3

Private Type TScrollviewState
    ' Last row index in table
    LastRowIndex As Long
    ' Last column index in table
    LastColumnIndex As Long
    ' Tells if scrolling out of table bounds is enabled
    OutOfRangeScrolling As Boolean
    ShiftKey As Boolean
    CtrlKey As Boolean
    AltKey As Boolean
    ColumnIndex As Long
    RowIndex As Long
    ' Cell as: Array(RowIndex, ColumnIndex)
    Cell As Variant
    IsColumnTrigger As Boolean
    IsRowTrigger As Boolean
    IsCellTrigger As Boolean
End Type


' --- PROPERTIES ---

Public Property Get SelectedColumnIndexes() As ArrayListEx: Set SelectedColumnIndexes = pSelectedColumnIndexes: End Property
Public Property Get SelectedRowIndexes() As ArrayListEx: Set SelectedRowIndexes = pSelectedRowIndexes: End Property

Public Property Get Scrollview() As Access.Form: Set Scrollview = pScrollview: End Property
Public Property Set Scrollview(ByRef Value As Access.Form): Set pScrollview = Value: End Property


' --- PRIVATE PROPERTIES ---

'Private Property Get LastTableRowIndex() As Long: LastTableRowIndex = pScrollview.Table.Count - 1: End Property



' --- METHODS ---

Private Sub Class_Initialize()
'    pFixedColumnsOnTheLeft = 0
    pLastClickedColumnIndex = -1
    pLastClickedRowIndex = -1
    pLastEnteredCell = Array(0, 0)
    Set pSelectedColumnIndexes = ArrayListEx.Create()
    Set pSelectedRowIndexes = ArrayListEx.Create()
End Sub


' --- EVENTS ---

Private Sub pScrollview_OnColumnLetterClick(ByVal ColumnIndex As Long, ByVal CtrlKey As Boolean, ByVal ShiftKey As Boolean)
    pLastSelectionTrigger = pST_COLUMN
    If ShiftKey And pLastClickedColumnIndex <> -1 Then
        SelectionAddColumns pLastClickedColumnIndex, ColumnIndex
    Else
        If Not CtrlKey Then SelectionClear
        SelectionToggleColumn ColumnIndex
        pLastClickedColumnIndex = ColumnIndex
    End If
    SelectionApply
End Sub

Private Sub pScrollview_OnRowNumberClick(ByVal RowIndex As Long, ByVal CtrlKey As Boolean, ByVal ShiftKey As Boolean)
    pLastSelectionTrigger = pST_ROW
    If ShiftKey And pLastClickedRowIndex <> -1 Then
        SelectionAddRows pLastClickedRowIndex, RowIndex
    Else
        If Not CtrlKey Then SelectionClear
        SelectionToggleRow RowIndex
        pLastClickedRowIndex = RowIndex
    End If
    SelectionApply
End Sub

Private Sub pScrollview_OnCellEnter(ByVal RowIndex As Long, ByVal ColumnIndex As Long, ByVal CtrlKey As Boolean, ByVal ShiftKey As Boolean)
    pLastSelectionTrigger = pST_CELL
    If Not (ShiftKey Xor CtrlKey) Then SelectionClear
    SelectionToggleRow RowIndex
    SelectionToggleColumn ColumnIndex
    If ShiftKey Then
        SelectionAddRows pLastEnteredCell(0), RowIndex
        SelectionAddColumns pLastEnteredCell(1), ColumnIndex
    End If
    pLastEnteredCell = Array(RowIndex, ColumnIndex)
    SelectionApply
End Sub

Private Sub pScrollview_OnClearSelectionRequest()
    pLastSelectionTrigger = pST_NONE
    SelectionClear
    SelectionApply
End Sub

Private Sub pScrollview_OnSelectAllRequest()
    SelectionAddColumns 0, pScrollview.Table.ColumnCount - 1
    SelectionApply
End Sub

Private Sub pScrollview_OnInvertSelectionRequest()
    Dim i As Long, iMax As Long
    iMax = pScrollview.Table.ColumnCount - 1

    For i = 0 To iMax
        SelectionToggleColumn i
    Next i
    SelectionApply
End Sub


' --- SELECTION ---

Public Sub SelectionClear()
    pSelectedColumnIndexes.Clear
    pLastClickedColumnIndex = -1
    pSelectedRowIndexes.Clear
    pLastClickedRowIndex = -1
End Sub

Public Sub SelectionToggleRow(ByVal RowIndex As Long)
    ToggleElementInList pSelectedRowIndexes, RowIndex
End Sub

Public Sub SelectionToggleColumn(ByVal ColumnIndex As Long)
    ToggleElementInList pSelectedColumnIndexes, ColumnIndex
End Sub

Public Sub SelectionAddRow(ByVal RowIndex As Long)
    AddElementIfNotExists pSelectedRowIndexes, RowIndex
End Sub

Public Sub SelectionAddColumn(ByVal ColumnIndex As Long)
    AddElementIfNotExists pSelectedColumnIndexes, ColumnIndex
End Sub

Public Sub SelectionAddRows(ByVal FromRowIndex As Long, ByVal ToRowIndex As Long)
    AddAllNonExistingInRange pSelectedRowIndexes, FromRowIndex, ToRowIndex
End Sub

Public Sub SelectionAddColumns(ByVal FromColumnIndex As Long, ByVal ToColumnIndex As Long)
    AddAllNonExistingInRange pSelectedColumnIndexes, FromColumnIndex, ToColumnIndex
End Sub

Public Sub SelectionApply()
    pScrollview.SelectedColumns = pSelectedColumnIndexes.ToArray
    pScrollview.SelectedRows = pSelectedRowIndexes.ToArray
End Sub


' --- SELECTION: Movers

Public Sub MoveTo(ByVal RowIndex As Long, ByVal ColumnIndex As Long, ByVal PropagateEvent As Boolean)
    pScrollview.MoveTo RowIndex, ColumnIndex, PropagateEvent
End Sub


' --- EVENTS: Keyboard ---

Private Sub pScrollview_OnSelectionControlKeyDown(KeyCode As Integer, Shift As Integer)
    If pLastSelectionTrigger = pST_NONE Then Exit Sub
    Dim sView As TScrollviewState: sView = GetScrollviewState(Shift)
    Dim iAux As Long, cIndex As Long, rIndex As Long, sKeyCode As Integer: sKeyCode = KeyCode
    KeyCode = 0

'    Debug.Print Printf("%1 %2 CALLING OnSelectionControlKeyDown(%3, %4)", TimerSpan, vbTab, sKeyCode, Shift)

    ' NOTE: use KeyCode = sKeyCode to avoid "preventing default action"
    With sView
        Select Case sKeyCode
            ' KeyDown Arrow
            Case vbKeyDown
                Select Case True
                    Case .IsColumnTrigger
                        If Shift = 0 Then   ' No key mods
                            SelectionClear
                            MoveTo 0, .ColumnIndex, True
                        ElseIf Shift = acCtrlMask Then   ' Only CtrlKey mod
                            SelectionClear
                            MoveTo .LastRowIndex, .ColumnIndex, True
                        Else
                            ' No other key mod applicable here
                        End If
                    Case .IsRowTrigger, .IsCellTrigger
                        cIndex = IIf(.IsRowTrigger, -1, .Cell(1))
                        rIndex = IIf(.IsRowTrigger, .RowIndex, .Cell(0))
                        If Shift = 0 Then
                            SelectionClear
                            MoveTo rIndex + 1, cIndex, True
                        ElseIf Shift = acCtrlMask Then
                            SelectionClear
                            MoveTo GetNextRowIndexInFullJump(sView), cIndex, True
                        ElseIf Shift = (acCtrlMask Or acShiftMask) Then ' Ctrl + Shift + ArrowDown
                            iAux = rIndex
                            rIndex = GetNextRowIndexInFullJump(sView)
                            SelectionAddRows iAux, rIndex
                            SelectionApply
                            MoveTo rIndex, cIndex, False
                        ElseIf Shift = acShiftMask Then
                            If .OutOfRangeScrolling Or ((Not .OutOfRangeScrolling) And rIndex < .LastRowIndex) Then
                                rIndex = rIndex + 1
                                SelectionAddRow rIndex
                                SelectionApply
                                MoveTo rIndex, cIndex, False
                            End If
                        End If
                End Select
            ' KeyUp Arrow
            Case vbKeyUp
                Select Case True
                    Case .IsColumnTrigger
                        SelectionClear
                        MoveTo 0, .ColumnIndex, True
                    Case .IsRowTrigger, .IsCellTrigger
                        cIndex = IIf(.IsRowTrigger, -1, .Cell(1))
                        rIndex = IIf(.IsRowTrigger, .RowIndex, .Cell(0))
                        If Shift = 0 Then
                            SelectionClear
                            MoveTo rIndex - 1, cIndex, True
                        ElseIf Shift = acCtrlMask Then
                            SelectionClear
                            If rIndex > .LastRowIndex Then
                                MoveTo .LastRowIndex, cIndex, True
                            Else
                                MoveTo 0, cIndex, True
                            End If
                        ElseIf Shift = (acCtrlMask Or acShiftMask) Then ' Ctrl + Shift + ArrowDown
                            iAux = rIndex
                            rIndex = IIf(rIndex > .LastRowIndex, .LastRowIndex, 0)
                            SelectionAddRows iAux, rIndex
                            SelectionApply
                            MoveTo rIndex, cIndex, False
                        ElseIf Shift = acShiftMask Then
                            If rIndex > 0 Then
                                rIndex = rIndex - 1
                                SelectionAddRow rIndex
                                SelectionApply
                                MoveTo rIndex, cIndex, False
                            End If
                        End If
                End Select
            ' KeyRight Arrow
            Case vbKeyRight
                Select Case True
                    Case .IsColumnTrigger, .IsCellTrigger
                        cIndex = IIf(.IsColumnTrigger, .ColumnIndex, .Cell(1))
                        rIndex = IIf(.IsColumnTrigger, -1, .Cell(0))
                        If Shift = 0 Then
                            SelectionClear
                            MoveTo rIndex, cIndex + 1, True
                        ElseIf Shift = acCtrlMask Then
                            SelectionClear
                            MoveTo rIndex, GetNextColumnIndexInFullJump(sView), True
                        ElseIf Shift = (acCtrlMask Or acShiftMask) Then ' Ctrl + Shift
                            iAux = cIndex
                            cIndex = GetNextColumnIndexInFullJump(sView)
                            SelectionAddColumns iAux, cIndex
                            SelectionApply
                            MoveTo rIndex, cIndex, False
                        ElseIf Shift = acShiftMask Then
                            If .OutOfRangeScrolling Or ((Not .OutOfRangeScrolling) And cIndex < .LastColumnIndex) Then
                                cIndex = cIndex + 1
                                SelectionAddColumn cIndex
                                SelectionApply
                                MoveTo rIndex, cIndex, False
                            End If
                        End If
                    Case .IsRowTrigger
                        If Shift = 0 Then   ' No key mods
                            SelectionClear
                            MoveTo .RowIndex, 0, True
                        ElseIf Shift = acCtrlMask Then   ' Only CtrlKey mod
                            SelectionClear
                            MoveTo .RowIndex, .LastColumnIndex, True
                        End If
                End Select
            ' KeyLeft Arrow
            Case vbKeyLeft
                Select Case True
                    Case .IsColumnTrigger, .IsCellTrigger
                        cIndex = IIf(.IsColumnTrigger, .ColumnIndex, .Cell(1))
                        rIndex = IIf(.IsColumnTrigger, -1, .Cell(0))
                        If Shift = 0 Then
                            SelectionClear
                            MoveTo rIndex, cIndex - 1, True
                        ElseIf Shift = acCtrlMask Then
                            SelectionClear
                            If cIndex > .LastColumnIndex Then
                                MoveTo rIndex, .LastColumnIndex, True
                            Else
                                MoveTo rIndex, 0, True
                            End If
                        ElseIf Shift = (acCtrlMask Or acShiftMask) Then ' Ctrl + Shift + ArrowDown
                            iAux = cIndex
                            cIndex = IIf(cIndex > .LastColumnIndex, .LastColumnIndex, 0)
                            SelectionAddColumns iAux, cIndex
                            SelectionApply
                            MoveTo rIndex, cIndex, False
                        ElseIf Shift = acShiftMask Then
                            If cIndex > 0 Then
                                cIndex = cIndex - 1
                                SelectionAddColumn cIndex
                                SelectionApply
                                MoveTo rIndex, cIndex, False
                            End If
                        End If
                    Case .IsRowTrigger
                        SelectionClear
                        MoveTo .RowIndex, 0, True
                End Select
            Case vbKeyPageDown, vbKeyPageUp
                Select Case True
                    Case .IsColumnTrigger 'Ignore
                    Case .IsRowTrigger, .IsCellTrigger
                        cIndex = IIf(.IsRowTrigger, -1, .Cell(1))
                        rIndex = IIf(.IsRowTrigger, .RowIndex, .Cell(0))
                        iAux = IIf(sKeyCode = vbKeyPageUp, -1, 1)
                        iAux = iAux * IIf(.ShiftKey, 3, 1)
                        If .CtrlKey Then
                            iAux = iAux * CLng(Int((pScrollview.ScrollPageSizeX / pScrollview.Worksheet.GridCellSizeX) * 0.8))
                            MoveTo rIndex, cIndex + iAux, True
                        Else
                            iAux = iAux * CLng(Int((pScrollview.ScrollPageSizeY / pScrollview.Worksheet.GridCellSizeY) * 0.8))
                            MoveTo rIndex + iAux, cIndex, True
                        End If
                End Select
            Case Else
                KeyCode = sKeyCode
        End Select
    End With

    ' Update last clicked/entered row/col/cell for those calls to MoveTo without event propagation.
    If KeyCode = 0 Then UpdateSelectionStateAfterMoveTo sView, rIndex, cIndex, sKeyCode
End Sub

Private Sub UpdateSelectionStateAfterMoveTo(ByRef sView As TScrollviewState, ByVal rIndex As Long, ByVal cIndex As Long, ByVal KeyCode As Integer)
    With sView
        Select Case KeyCode
            Case vbKeyUp, vbKeyDown
                If .IsRowTrigger Then
                    If pLastClickedRowIndex = .RowIndex And .RowIndex <> rIndex Then
                        pLastClickedRowIndex = rIndex
                    End If
                ElseIf .IsCellTrigger Then
                    If (pLastEnteredCell(0) = .Cell(0) And .Cell(0) <> rIndex) Or _
                       (pLastEnteredCell(1) = .Cell(1) And .Cell(1) <> cIndex) Then
                        pLastEnteredCell = Array(rIndex, cIndex)
                    End If
                End If
            Case vbKeyRight, vbKeyLeft
                If .IsRowTrigger Then
                    If pLastClickedColumnIndex = .ColumnIndex And .ColumnIndex <> cIndex Then
                        pLastClickedColumnIndex = cIndex
                    End If
                ' TODO: refactor (DRY)
                ElseIf .IsCellTrigger Then
                    If (pLastEnteredCell(0) = .Cell(0) And .Cell(0) <> rIndex) Or _
                       (pLastEnteredCell(1) = .Cell(1) And .Cell(1) <> cIndex) Then
                        pLastEnteredCell = Array(rIndex, cIndex)
                    End If
                End If
        End Select
    End With
End Sub

Private Function GetNextRowIndexInFullJump(ByRef sView As TScrollviewState) As Long
    Dim rIndex As Long

    With sView
        rIndex = IIf(.IsRowTrigger, .RowIndex, .Cell(0))
        If rIndex >= .LastRowIndex Then
            If .OutOfRangeScrolling Then
                GetNextRowIndexInFullJump = CLng(Int((rIndex + 50000) / 50000)) * 50000
            Else
                GetNextRowIndexInFullJump = .LastRowIndex
            End If
        Else
            GetNextRowIndexInFullJump = .LastRowIndex
        End If
    End With
End Function

Private Function GetNextColumnIndexInFullJump(ByRef sView As TScrollviewState) As Long
    Dim cIndex As Long

    With sView
        cIndex = IIf(.IsColumnTrigger, .ColumnIndex, .Cell(1))
        If cIndex >= .LastColumnIndex Then
            If .OutOfRangeScrolling Then
                GetNextColumnIndexInFullJump = CLng(Int((cIndex + 50000) / 50000)) * 50000
            Else
                GetNextColumnIndexInFullJump = .LastColumnIndex
            End If
        Else
            GetNextColumnIndexInFullJump = .LastColumnIndex
        End If
    End With
End Function

Private Function GetScrollviewState(ByVal Shift As Integer) As TScrollviewState
    Dim t As TScrollviewState

    With pScrollview.Table
        t.LastRowIndex = .Count - 1
        t.LastColumnIndex = .ColumnCount - 1
    End With
    t.OutOfRangeScrolling = pScrollview.EnableOutOfRangeScrolling
    t.ShiftKey = ((Shift And acShiftMask) > 0)
    t.CtrlKey = ((Shift And acCtrlMask) > 0)
    t.AltKey = ((Shift And acAltMask) > 0)
    t.RowIndex = pLastClickedRowIndex
    t.ColumnIndex = pLastClickedColumnIndex
    t.Cell = pLastEnteredCell
    t.IsColumnTrigger = (pLastSelectionTrigger = pST_COLUMN)
    t.IsRowTrigger = (pLastSelectionTrigger = pST_ROW)
    t.IsCellTrigger = (pLastSelectionTrigger = pST_CELL)

    GetScrollviewState = t
End Function

Private Sub AddAllNonExistingInRange(ByRef aX As ArrayListEx, ByVal FromIndex As Long, ByVal ToIndex As Long)
    Dim i As Long, r As Long, c As Long: c = -1

    If ToIndex < FromIndex Then
        i = FromIndex
        FromIndex = ToIndex
        ToIndex = i
    End If

    With aX
        For i = FromIndex To ToIndex
            r = .BinarySearch(i, r, , Comparer)
            If r < 0 Then
                r = (Not r)
                If r = .Instance.Count Then
                    c = i
                    Exit For
                End If
                .Insert r, i
            End If
        Next i
        If c <> -1 Then
            For i = c To ToIndex
                .Add i
            Next i
        End If
    End With
End Sub

Private Sub ToggleElementInList(ByRef aX As ArrayListEx, ByVal Value As Long)
    Dim r As Long

    r = aX.BinarySearch(Value, 0, , Comparer)
    If r < 0 Then
        r = (Not r)
        If r = aX.Instance.Count Then
            aX.Add Value
        Else
            aX.Insert r, Value
        End If
    Else
        aX.RemoveAt r
    End If
End Sub

Private Sub AddElementIfNotExists(ByRef aX As ArrayListEx, ByVal Value As Long)
    Dim r As Long

    r = aX.BinarySearch(Value, 0, , Comparer)
    If r < 0 Then
        r = (Not r)
        If r = aX.Instance.Count Then
            aX.Add Value
        Else
            aX.Insert r, Value
        End If
    End If
End Sub
