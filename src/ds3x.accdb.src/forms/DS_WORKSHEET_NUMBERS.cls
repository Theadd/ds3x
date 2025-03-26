Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "ds3x.UI.Scrollview"
Option Compare Database
Option Explicit
Option Base 0

Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Public Event OnRowNumberClick(ByVal RowIndex As Long, ByVal CtrlKey As Boolean, ByVal ShiftKey As Boolean)

Private pViewport As Form_DS_VIEWPORT
Private pSelectedRows As ArrayListEx

Private Const pMaxAvailColumns As Long = 1
Private Const pGridCellSizeX As Long = 690
Private Const pGridCellSizeY As Long = 375

Private pFirstRowIndex As Long


Public Property Get Viewport() As Form_DS_VIEWPORT: Set Viewport = pViewport: End Property
Public Property Set Viewport(ByRef Value As Form_DS_VIEWPORT): Set pViewport = Value: End Property

Public Property Get MaxAvailColumns() As Long: MaxAvailColumns = pMaxAvailColumns: End Property
Public Property Get MaxContentWidthLimit() As Long: MaxContentWidthLimit = pGridCellSizeX * pMaxAvailColumns: End Property
' Default Cell/Column width
Public Property Get GridCellSizeX() As Long: GridCellSizeX = pGridCellSizeX: End Property
Public Property Get GridCellSizeY() As Long: GridCellSizeY = pGridCellSizeY: End Property


' --- FORM EVENTS ---

Private Sub Form_Load()
    Set pSelectedRows = ArrayListEx.Create()
End Sub

Private Sub Form_MouseWheel(ByVal Page As Boolean, ByVal Count As Long)
    On Error Resume Next
    pViewport.PropagateMouseWheel Page, Count
    On Error GoTo 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer): pViewport.Scrollview.OnKeyDownHandler KeyCode, Shift: End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer): pViewport.Scrollview.OnKeyUpHandler KeyCode, Shift: End Sub


' --- SETUP / BINDING ---

Public Sub Setup(ByVal RowsCount As Long)
    Set Me.Recordset = RecordsetEx.CreateBlank(RowsCount, 1).Instance
End Sub

Friend Sub SetupGrid(ByVal FirstColumnIndex As Long, ByVal FirstRowIndex As Long, ByRef dsT As dsTable)
    pFirstRowIndex = FirstRowIndex
    Me.Recalc
End Sub

Friend Sub SetSelectedRows(ByRef aX As ArrayListEx)
    Set pSelectedRows = aX
    Me.Recalc
End Sub


' --- HELPERS ---

Private Function Max(x As Variant, Y As Variant) As Variant: Max = IIf(x > Y, x, Y): End Function
Private Function Min(x As Variant, Y As Variant) As Variant: Min = IIf(x < Y, x, Y): End Function


' --- MISC ---

Public Function GetPageRowNum() As Long
    GetPageRowNum = CLng(Me.Bookmark(0)) + pFirstRowIndex
End Function

Public Function GetRowSelectedValue() As Long
    On Error GoTo Finally

    If pSelectedRows.Contains(CLng(Me.Bookmark(0)) + pFirstRowIndex - 1) Then
        GetRowSelectedValue = 1
    End If
Finally:
End Function

Public Function OnRowHeaderClick()
    Me.RowHiddenControl.SetFocus
    RaiseEvent OnRowNumberClick(CLng(Val(Nz(Me.DS_HR_NUMBER, 1))) - 1, GetAsyncKeyState(vbKeyControl), GetAsyncKeyState(vbKeyShift))
End Function
