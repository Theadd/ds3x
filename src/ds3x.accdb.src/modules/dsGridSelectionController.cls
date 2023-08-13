VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "dsGridSelectionController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

'Private Sub pScrollview_OnColumnNameChange(ByVal ColumnIndex As Long, ByVal FromValue As String, ByVal ToValue As String)
'    TaskController().AddTask _
'        "SetColumnName", _
'        Array(, ColumnIndex, ToValue)
'End Sub

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


'Friend Sub TriggerOnChangeEvent()
'    RaiseEvent OnChange
'    If pTaskController.IsValidTable Then
'        Set pScrollview.Table = pTaskController.Table
'    Else
'        Set pScrollview.Table = CreateBlankTable(100, 0)
'    End If
'End Sub


' --- SELECTION ---

Public Sub SelectionClear()
    pSelectedColumnIndexes.Clear
    pLastClickedColumnIndex = -1
    pSelectedRowIndexes.Clear
    pLastClickedRowIndex = -1
End Sub

Public Sub SelectionToggleRow(ByVal RowIndex As Long)
    If pSelectedRowIndexes.Contains(RowIndex) Then
        pSelectedRowIndexes.Remove RowIndex
    Else
        pSelectedRowIndexes.Add RowIndex
    End If
End Sub

Public Sub SelectionToggleColumn(ByVal ColumnIndex As Long)
    If pSelectedColumnIndexes.Contains(ColumnIndex) Then
        pSelectedColumnIndexes.Remove ColumnIndex
    Else
        pSelectedColumnIndexes.Add ColumnIndex
    End If
End Sub

Public Sub SelectionAddRow(ByVal RowIndex As Long)
    pSelectedRowIndexes.Add RowIndex
End Sub

Public Sub SelectionAddColumn(ByVal ColumnIndex As Long)
    pSelectedColumnIndexes.Add ColumnIndex
End Sub

Public Sub SelectionAddRows(ByVal FromRowIndex As Long, ByVal ToRowIndex As Long)
    Dim i As Long, iStep As Long
    
    iStep = IIf(FromRowIndex <= ToRowIndex, 1, -1)
    For i = FromRowIndex To ToRowIndex Step iStep
        pSelectedRowIndexes.Add i
    Next i
End Sub

Public Sub SelectionAddColumns(ByVal FromColumnIndex As Long, ByVal ToColumnIndex As Long)
    Dim i As Long, iStep As Long

    iStep = IIf(FromColumnIndex <= ToColumnIndex, 1, -1)
    For i = FromColumnIndex To ToColumnIndex Step iStep
        pSelectedColumnIndexes.Add i
    Next i
End Sub

Public Sub SelectionApply()
    If pSelectedColumnIndexes.Count > 1 Then
        pSelectedColumnIndexes.Unique
        pSelectedColumnIndexes.Sort
    End If
    If pSelectedRowIndexes.Count > 1 Then
        pSelectedRowIndexes.Unique
        pSelectedRowIndexes.Sort
    End If
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
    
    ' NOTE: use KeyCode = sKeyCode to avoid "preventing default action"
    With sView
        Select Case sKeyCode
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
            Case vbKeyUp
                
        End Select
    End With
    
    ' Update last clicked/entered row/col/cell for those calls to MoveTo without event propagation.
    If KeyCode = 0 Then UpdateSelectionStateAfterMoveTo sView, rIndex, cIndex
End Sub

Private Sub UpdateSelectionStateAfterMoveTo(ByRef sView As TScrollviewState, ByVal rIndex As Long, ByVal cIndex As Long)
    With sView
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

'Private Type TScrollviewState
'    ' Last row index in table
'    LastRowIndex As Long
'    ' Last column index in table
'    LastColumnIndex As Long
'    ' Tells if scrolling out of table bounds is enabled
'    OutOfRangeScrolling As Boolean
'    ShiftKey As Boolean
'    CtrlKey As Boolean
'    AltKey As Boolean
'    ColumnIndex As Long
'    RowIndex As Long
'    Cell As Variant
'    IsColumnTrigger As Boolean
'    IsRowTrigger As Boolean
'    IsCellTrigger As Boolean
'End Type

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





