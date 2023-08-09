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

'Private WithEvents pGridForm As Access.Form
'Private WithEvents pGridContainer As Access.Form
Private WithEvents pTaskController As dsTaskController
Attribute pTaskController.VB_VarHelpID = -1
Private WithEvents pScrollview As Form_DS_SCROLLVIEW
Attribute pScrollview.VB_VarHelpID = -1

'Private Const pHeaderButtonBackColorSelected As Long = 12644351
'Private Const pHeaderButtonBackColorNormal As Long = 15132391

'Private Const pDefaultColumnWidth As Long = 1935
'Private Const pScrollIncrement As Long = 129

'Private pFirstColumnInViewport As Long
Private pFixedColumnsOnTheLeft As Long
'Private pViewportScrollX As Long

' ColumnIndex related to the last DS_HC_1_* CommandButton clicked
Private pLastClickedColumnIndex As Long
Private pSelectedColumnIndexes As New ArrayListEx
Private pLastClickedRowIndex As Long
Private pSelectedRowIndexes As New ArrayListEx


Public Event OnChange()



'Public Property Get GridForm() As Access.Form: Set GridForm = pGridForm: End Property
'Public Property Set GridForm(ByRef Value As Access.Form): Set pGridForm = Value: End Property
'Public Property Get GridContainer() As Access.Form: Set GridContainer = pGridContainer: End Property
'Public Property Set GridContainer(ByRef Value As Access.Form): Set pGridContainer = Value: End Property
Public Property Get TaskController() As dsTaskController: Set TaskController = pTaskController: End Property
Public Property Set TaskController(ByRef Value As dsTaskController): Set pTaskController = Value: End Property
Public Property Get Table() As dsTable: Set Table = TaskController().Table: End Property

'Public Property Get HeaderButtonBackColorSelected() As Long: HeaderButtonBackColorSelected = pHeaderButtonBackColorSelected: End Property
'Public Property Get HeaderButtonBackColorNormal() As Long: HeaderButtonBackColorNormal = pHeaderButtonBackColorNormal: End Property
Public Property Get SelectedColumnIndexes() As ArrayListEx: Set SelectedColumnIndexes = pSelectedColumnIndexes: End Property
Public Property Get SelectedRowIndexes() As ArrayListEx: Set SelectedRowIndexes = pSelectedRowIndexes: End Property

'Public Property Get DefaultColumnWidth() As Long: DefaultColumnWidth = pDefaultColumnWidth: End Property
'Public Property Get ColumnWidth(ByVal ColumnIndex As Long) As Long: ColumnWidth = GetColumnWidth(ColumnIndex): End Property
'Public Property Get ColumnLeft(ByVal ColumnIndex As Long) As Long: ColumnLeft = GetColumnLeft(ColumnIndex): End Property
'Public Property Get ScrollIncrement() As Long: ScrollIncrement = pScrollIncrement: End Property
'Public Property Get ViewportScrollX() As Long: ViewportScrollX = pViewportScrollX: End Property
'Public Property Let ViewportScrollX(ByVal Value As Long): pViewportScrollX = IIf(Value > 0, Value, 0): End Property

Public Property Get NumColumnsInTable() As Long: NumColumnsInTable = Table.ColumnCount - IIf(Table.RowNumbers, 1, 0): End Property
Public Property Get FixedColumnsOnTheLeft() As Long: FixedColumnsOnTheLeft = pFixedColumnsOnTheLeft: End Property

'Public Property Get FirstColumnInViewport() As Long
'    ' TODO: Implement variable column width
'    FirstColumnInViewport = CLng(Int(pViewportScrollX / pScrollIncrement))
'End Property

'Public Property Get VirtualScrollX() As Long
'    Dim a As Long, b As Long, c As Long
'
'    a = CLng(Int(pViewportScrollX / pScrollIncrement))
'    b = a * pScrollIncrement
'    c = pViewportScrollX - b
'    VirtualScrollX = CLng((GetColumnWidth(a + 1) / pScrollIncrement) * c)
'    Debug.Print Printf("a: %1, b: %2, c: %3, (%4 / %5) * %3: %6", a, b, c, GetColumnWidth(a + 1), pScrollIncrement, (GetColumnWidth(a + 1) / pScrollIncrement) * c)
'End Property

Public Property Get Scrollview() As Access.Form: Set Scrollview = pScrollview: End Property
Public Property Set Scrollview(ByRef Value As Access.Form): Set pScrollview = Value: End Property


Private Sub Class_Initialize()
'    Set pGridForm = Nothing
'    Set pGridContainer = Nothing
    Set pTaskController = Nothing
'    pFirstColumnInViewport = 0
    pFixedColumnsOnTheLeft = 0
    pLastClickedColumnIndex = -1
    pLastClickedRowIndex = -1
End Sub


Private Sub pScrollview_OnColumnLetterClick(ByVal ColumnIndex As Long, ByVal CtrlKey As Boolean, ByVal ShiftKey As Boolean)
    HandleClickOnColumnHeader ColumnIndex, ShiftKey, CtrlKey
End Sub

Private Sub pScrollview_OnColumnNameChange(ByVal ColumnIndex As Long, ByVal FromValue As String, ByVal ToValue As String)
    TaskController().AddTask _
        "SetColumnName", _
        Array(, ColumnIndex, ToValue)
End Sub

Private Sub pTaskController_OnTableChange()
    ClearSelection
    TriggerOnChangeEvent
End Sub

'Public Sub TriggerViewportSizeUpdate()
'
'    GridForm.ViewportSizeUpdate
'    GridContainer.ViewportSizeUpdate
'End Sub

Public Function TryGetColumnHeaderAt(ByVal ColumnIndex As Long, ByRef OutColumnHeader As DictionaryEx) As Boolean
    On Error GoTo Finally
    
    Set OutColumnHeader = DictionaryEx.Bind(Table().HeaderList(0)(ColumnIndex)).Clone()
    TryGetColumnHeaderAt = True
Finally:
End Function

Public Sub HandleClickOnColumnHeader(ByVal ColumnIndex As Long, ByVal ShiftKey As Boolean, ByVal CtrlKey As Boolean)
    Dim i As Long, iStep As Long
    
    If ShiftKey And pLastClickedColumnIndex <> -1 Then
        iStep = IIf(pLastClickedColumnIndex <= ColumnIndex, 1, -1)
        For i = pLastClickedColumnIndex To ColumnIndex Step iStep
            pSelectedColumnIndexes.Add i
        Next i
    Else
        If CtrlKey Then
            If pSelectedColumnIndexes.Contains(ColumnIndex) Then
                pSelectedColumnIndexes.Remove ColumnIndex
            Else
                pSelectedColumnIndexes.Add ColumnIndex
            End If
        Else
            pSelectedColumnIndexes.Clear
            pSelectedColumnIndexes.Add ColumnIndex
        End If
        
        pLastClickedColumnIndex = ColumnIndex
    End If
    
    If pSelectedColumnIndexes.Count > 1 Then
        pSelectedColumnIndexes.Unique
        pSelectedColumnIndexes.Sort
    End If
    pScrollview.SelectedColumns = pSelectedColumnIndexes.ToArray
End Sub

Public Sub HandleClickOnRowHeader(ByVal RowIndex As Long, ByVal ShiftKey As Boolean, ByVal CtrlKey As Boolean)
    Dim i As Long, iStep As Long
    
    If ShiftKey And pLastClickedRowIndex <> -1 Then
        iStep = IIf(pLastClickedRowIndex <= RowIndex, 1, -1)
        For i = pLastClickedRowIndex To RowIndex Step iStep
            pSelectedRowIndexes.Add i
        Next i
    Else
        If CtrlKey Then
            If pSelectedRowIndexes.Contains(RowIndex) Then
                pSelectedRowIndexes.Remove RowIndex
            Else
                pSelectedRowIndexes.Add RowIndex
            End If
        Else
            pSelectedRowIndexes.Clear
            pSelectedRowIndexes.Add RowIndex
        End If

        pLastClickedRowIndex = RowIndex
    End If

    If pSelectedRowIndexes.Count > 1 Then
        pSelectedRowIndexes.Unique
        pSelectedRowIndexes.Sort
    End If
    RaiseEvent OnChange
End Sub

Private Sub ClearSelection()
    pSelectedColumnIndexes.Clear
    pLastClickedColumnIndex = -1
    pSelectedRowIndexes.Clear
    pLastClickedRowIndex = -1
    pScrollview.SelectedColumns = pSelectedColumnIndexes.ToArray
End Sub

'Private Function GetColumnWidth(ByVal ColumnIndex As Long) As Long
'    ' TODO: Implement variable column width
'    Select Case ColumnIndex
'        Case 0
'            GetColumnWidth = 660
'        Case Else
'            GetColumnWidth = pDefaultColumnWidth
'    End Select
'End Function
'
'Private Function GetColumnLeft(ByVal ColumnIndex As Long) As Long
'    ' TODO: Implement variable column width
'    Select Case ColumnIndex
'        Case 0
'            GetColumnLeft = 0
'        Case Else
'            GetColumnLeft = 660 + ((ColumnIndex - 1) * pDefaultColumnWidth)
'    End Select
'End Function

Friend Sub TriggerOnChangeEvent()
    RaiseEvent OnChange
    If pTaskController.IsValidTable Then
        Set pScrollview.Table = pTaskController.Table
    Else
        Debug.Print "BLANK TABLE"
        Set pScrollview.Table = CreateBlankTable(50, 1)
    End If
    pScrollview.SelectedColumns = pSelectedColumnIndexes.ToArray
End Sub





