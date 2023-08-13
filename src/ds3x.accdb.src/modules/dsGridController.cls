VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "dsGridController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Option Base 0

Public Event OnChange()

Private WithEvents pTaskController As dsTaskController
Attribute pTaskController.VB_VarHelpID = -1
Private WithEvents pScrollview As Form_DS_SCROLLVIEW
Attribute pScrollview.VB_VarHelpID = -1
Private pSelectionController As dsGridSelectionController

' DEPRECATED
Private pFixedColumnsOnTheLeft As Long
' ColumnIndex related to the last DS_HC_1_* CommandButton clicked
'Private pLastClickedColumnIndex As Long
'Private pSelectedColumnIndexes As New ArrayListEx
'Private pLastClickedRowIndex As Long
'Private pSelectedRowIndexes As New ArrayListEx
'Private pLastEnteredCell As Variant


' --- PROPERTIES ---
Public Property Get TaskController() As dsTaskController: Set TaskController = pTaskController: End Property
Public Property Set TaskController(ByRef Value As dsTaskController): Set pTaskController = Value: End Property
Public Property Get Table() As dsTable: Set Table = TaskController().Table: End Property

Public Property Get SelectedColumnIndexes() As ArrayListEx: Set SelectedColumnIndexes = pSelectionController.SelectedColumnIndexes: End Property
Public Property Get SelectedRowIndexes() As ArrayListEx: Set SelectedRowIndexes = pSelectionController.SelectedRowIndexes: End Property

Public Property Get Scrollview() As Access.Form: Set Scrollview = pScrollview: End Property
Public Property Set Scrollview(ByRef Value As Access.Form): Set pScrollview = Value: Set pSelectionController.Scrollview = Value: End Property

' BEGIN DEPRECATED
Public Property Get NumColumnsInTable() As Long: NumColumnsInTable = Table.ColumnCount - IIf(Table.RowNumbers, 1, 0): End Property
Public Property Get FixedColumnsOnTheLeft() As Long: FixedColumnsOnTheLeft = pFixedColumnsOnTheLeft: End Property
' END DEPRECATED

' --- METHODS ---

Private Sub Class_Initialize()
    Set pTaskController = Nothing
    Set pSelectionController = New dsGridSelectionController
    ' DEPRECATED
    pFixedColumnsOnTheLeft = 0
'    pLastClickedColumnIndex = -1
'    pLastClickedRowIndex = -1
'    pLastEnteredCell = Array(0, 0)
End Sub


' --- EVENTS: Scrollview ---

'Private Sub pScrollview_OnColumnLetterClick(ByVal ColumnIndex As Long, ByVal CtrlKey As Boolean, ByVal ShiftKey As Boolean)
'    HandleClickOnColumnHeader ColumnIndex, ShiftKey, CtrlKey
'End Sub

Private Sub pScrollview_OnColumnNameChange(ByVal ColumnIndex As Long, ByVal FromValue As String, ByVal ToValue As String)
    TaskController().AddTask _
        "SetColumnName", _
        Array(, ColumnIndex, ToValue)
End Sub

'Private Sub pScrollview_OnRowNumberClick(ByVal RowIndex As Long, ByVal CtrlKey As Boolean, ByVal ShiftKey As Boolean)
'    HandleClickOnRowHeader RowIndex, ShiftKey, CtrlKey
'End Sub
'
'Private Sub pScrollview_OnCellEnter(ByVal RowIndex As Long, ByVal ColumnIndex As Long, ByVal CtrlKey As Boolean, ByVal ShiftKey As Boolean)
'    HandleEnterOnCell RowIndex, ColumnIndex, ShiftKey, CtrlKey
'End Sub
'
'Private Sub pScrollview_OnClearSelectionRequest()
'    SelectionClear
'    SelectionApply
'End Sub

Public Sub SelectionClear(): pSelectionController.SelectionClear: End Sub
Public Sub SelectionApply(): pSelectionController.SelectionApply: End Sub

' --- EVENTS: TaskController ---

Private Sub pTaskController_OnTableChange()
    SelectionClear
    TriggerOnChangeEvent
    SelectionApply
End Sub

Public Function TryGetColumnHeaderAt(ByVal ColumnIndex As Long, ByRef OutColumnHeader As DictionaryEx) As Boolean
    On Error GoTo Finally
    
    Set OutColumnHeader = DictionaryEx.Bind(Table().HeaderList(0)(ColumnIndex)).Clone()
    TryGetColumnHeaderAt = True
Finally:
End Function

'Public Sub HandleClickOnColumnHeader(ByVal ColumnIndex As Long, ByVal ShiftKey As Boolean, ByVal CtrlKey As Boolean)
'    If ShiftKey And pLastClickedColumnIndex <> -1 Then
'        SelectionAddColumns pLastClickedColumnIndex, ColumnIndex
'    Else
'        If Not CtrlKey Then SelectionClear
'        SelectionToggleColumn ColumnIndex
'        pLastClickedColumnIndex = ColumnIndex
'    End If
'    SelectionApply
'End Sub
'
'Public Sub HandleClickOnRowHeader(ByVal RowIndex As Long, ByVal ShiftKey As Boolean, ByVal CtrlKey As Boolean)
'    If ShiftKey And pLastClickedRowIndex <> -1 Then
'        SelectionAddRows pLastClickedRowIndex, RowIndex
'    Else
'        If Not CtrlKey Then SelectionClear
'        SelectionToggleRow RowIndex
'        pLastClickedRowIndex = RowIndex
'    End If
'    SelectionApply
'End Sub
'
'Public Sub HandleEnterOnCell(ByVal RowIndex As Long, ByVal ColumnIndex As Long, ByVal ShiftKey As Boolean, ByVal CtrlKey As Boolean)
'    If Not (ShiftKey Xor CtrlKey) Then SelectionClear
''    If CtrlKey Then
'        SelectionToggleRow RowIndex
'        SelectionToggleColumn ColumnIndex
''    End If
'    If ShiftKey Then
'        SelectionAddRows pLastEnteredCell(0), RowIndex
'        SelectionAddColumns pLastEnteredCell(1), ColumnIndex
'    End If
'    pLastEnteredCell = Array(RowIndex, ColumnIndex)
'    SelectionApply
'End Sub

Friend Sub TriggerOnChangeEvent()
    RaiseEvent OnChange
    If pTaskController.IsValidTable Then
        Set pScrollview.Table = pTaskController.Table
    Else
        Set pScrollview.Table = CreateBlankTable(100, 0)
    End If
End Sub


' --- SELECTION ---

'Public Sub SelectionClear()
'    pSelectedColumnIndexes.Clear
'    pLastClickedColumnIndex = -1
'    pSelectedRowIndexes.Clear
'    pLastClickedRowIndex = -1
'End Sub
'
'Public Sub SelectionToggleRow(ByVal RowIndex As Long)
'    If pSelectedRowIndexes.Contains(RowIndex) Then
'        pSelectedRowIndexes.Remove RowIndex
'    Else
'        pSelectedRowIndexes.Add RowIndex
'    End If
'End Sub
'
'Public Sub SelectionToggleColumn(ByVal ColumnIndex As Long)
'    If pSelectedColumnIndexes.Contains(ColumnIndex) Then
'        pSelectedColumnIndexes.Remove ColumnIndex
'    Else
'        pSelectedColumnIndexes.Add ColumnIndex
'    End If
'End Sub
'
'Public Sub SelectionAddRow(ByVal RowIndex As Long)
'    pSelectedRowIndexes.Add RowIndex
'End Sub
'
'Public Sub SelectionAddColumn(ByVal ColumnIndex As Long)
'    pSelectedColumnIndexes.Add ColumnIndex
'End Sub
'
'Public Sub SelectionAddRows(ByVal FromRowIndex As Long, ByVal ToRowIndex As Long)
'    Dim i As Long, iStep As Long
'
'    iStep = IIf(FromRowIndex <= ToRowIndex, 1, -1)
'    For i = FromRowIndex To ToRowIndex Step iStep
'        pSelectedRowIndexes.Add i
'    Next i
'End Sub
'
'Public Sub SelectionAddColumns(ByVal FromColumnIndex As Long, ByVal ToColumnIndex As Long)
'    Dim i As Long, iStep As Long
'
'    iStep = IIf(FromColumnIndex <= ToColumnIndex, 1, -1)
'    For i = FromColumnIndex To ToColumnIndex Step iStep
'        pSelectedColumnIndexes.Add i
'    Next i
'End Sub
'
'Public Sub SelectionApply()
'    If pSelectedColumnIndexes.Count > 1 Then
'        pSelectedColumnIndexes.Unique
'        pSelectedColumnIndexes.Sort
'    End If
'    If pSelectedRowIndexes.Count > 1 Then
'        pSelectedRowIndexes.Unique
'        pSelectedRowIndexes.Sort
'    End If
'    pScrollview.SelectedColumns = pSelectedColumnIndexes.ToArray
'    pScrollview.SelectedRows = pSelectedRowIndexes.ToArray
'End Sub




