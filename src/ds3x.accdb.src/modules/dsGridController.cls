VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "dsGridController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder "ds3x.LiveEditor.UI"
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


' --- PROPERTIES ---
Public Property Get TaskController() As dsTaskController: Set TaskController = pTaskController: End Property
Public Property Set TaskController(ByRef Value As dsTaskController): Set pTaskController = Value: End Property
Public Property Get Table() As dsTable: Set Table = TaskController().Table: End Property

Public Property Get SelectedColumnIndexes() As ArrayListEx: Set SelectedColumnIndexes = pSelectionController.SelectedColumnIndexes: End Property
Public Property Get SelectedRowIndexes() As ArrayListEx: Set SelectedRowIndexes = pSelectionController.SelectedRowIndexes: End Property

Public Property Get Scrollview() As Access.Form: Set Scrollview = pScrollview: End Property
Public Property Set Scrollview(ByRef Value As Access.Form): Set pScrollview = Value: Set pSelectionController.Scrollview = Value: End Property


' --- METHODS ---

Private Sub Class_Initialize()
    Set pTaskController = Nothing
    Set pSelectionController = New dsGridSelectionController
End Sub


' --- EVENTS: Scrollview ---

Private Sub pScrollview_OnColumnNameChange(ByVal ColumnIndex As Long, ByVal FromValue As String, ByVal ToValue As String)
    TaskController().AddTask _
        "SetColumnName", _
        Array(, ColumnIndex, ToValue)
End Sub

Public Sub SelectionClear(): pSelectionController.SelectionClear: End Sub
Public Sub SelectionApply(): pSelectionController.SelectionApply: End Sub


' --- EVENTS: TaskController ---

Private Sub pTaskController_OnTableChange()
    SelectionClear
    TriggerOnChangeEvent
    SelectionApply
End Sub

Private Sub pTaskController_OnBusyChange(ByVal IsBusy As Boolean)
    ScreenLib.MouseLoadingCursor = IsBusy
End Sub


' --- OTHER ---

Friend Sub TriggerOnChangeEvent()
    RaiseEvent OnChange
    If pTaskController.IsValidTable Then
        Set pScrollview.Table = pTaskController.Table
    Else
        Set pScrollview.Table = dsTable.CreateBlank(100, 0)
    End If
End Sub


