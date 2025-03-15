Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "ds3x.UI.Scrollview"
Option Compare Database
Option Explicit
Option Base 0

Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Public Event OnCellEnter(ByVal RowIndex As Long, ByVal ColumnIndex As Long, ByVal CtrlKey As Boolean, ByVal ShiftKey As Boolean)

Private pViewport As Form_DS_VIEWPORT

Private Const pMaxAvailColumns As Long = 16
Private Const pGridCellSizeX As Long = 1965
Private Const pGridCellSizeY As Long = 375

Private pFirstInactiveRowIndex As Long
Private pFirstInactiveColumnIndex As Long
Private pFirstColumnIndex As Long
Private pFirstRowIndex As Long


Public Property Get Viewport() As Form_DS_VIEWPORT: Set Viewport = pViewport: End Property
Public Property Set Viewport(ByRef Value As Form_DS_VIEWPORT): Set pViewport = Value: End Property

Public Property Get MaxAvailColumns() As Long: MaxAvailColumns = pMaxAvailColumns: End Property
' Maximum available form width filled with table-related controls (almost 22 inches ~= 55.478cm ~= 31456 twips in my dev environment).
Public Property Get MaxContentWidthLimit() As Long: MaxContentWidthLimit = pGridCellSizeX * pMaxAvailColumns: End Property
' Default Cell/Column width
Public Property Get GridCellSizeX() As Long: GridCellSizeX = pGridCellSizeX: End Property
Public Property Get GridCellSizeY() As Long: GridCellSizeY = pGridCellSizeY: End Property


' --- FORM EVENTS ---

Private Sub Form_MouseWheel(ByVal Page As Boolean, ByVal Count As Long)
    On Error Resume Next
    pViewport.PropagateMouseWheel Page, Count
    On Error GoTo 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer): pViewport.Scrollview.OnKeyDownHandler KeyCode, Shift: End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer): pViewport.Scrollview.OnKeyUpHandler KeyCode, Shift: End Sub


' --- SETUP / BINDING ---

Public Sub Setup()
    SetupGridSizes
End Sub

Private Sub SetupGridSizes()
    Dim i As Long

    For i = 1 To pMaxAvailColumns
        If i = 1 Then
            Me.DS_RC_1_1.Left = 0
        End If
        Me.Controls("DS_RC_1_" & CStr(i)).Width = pGridCellSizeX
    Next i

    Me.DS_PAGE_ROW_NUM.Left = 0
    Me.DS_PAGE_ROW_NUM.Width = pGridCellSizeX * pMaxAvailColumns
    Me.Width = pGridCellSizeX * pMaxAvailColumns
End Sub

Friend Sub SetupGrid(ByVal FirstColumnIndex As Long, ByVal FirstRowIndex As Long, ByRef dsT As dsTable)
    pFirstColumnIndex = FirstColumnIndex
    pFirstRowIndex = FirstRowIndex
    pFirstInactiveColumnIndex = Max(Min(dsT.ColumnCount - FirstColumnIndex, pMaxAvailColumns), 0)
    pFirstInactiveRowIndex = Max(dsT.Count - FirstRowIndex, 0)
    SetupInactiveColumnsBackground
End Sub

Private Sub SetupInactiveColumnsBackground()
    With Me.DS_RC_BG_INACTIVE
        If pFirstInactiveColumnIndex < pMaxAvailColumns Then
            .Width = 15
            .Left = Me.Controls("DS_RC_1_" & CStr(pFirstInactiveColumnIndex + 1)).Left
            .Width = Me.Width - .Left - 150
            If Not .Visible Then .Visible = True
        Else
            If .Visible Then .Visible = False
        End If
    End With
End Sub


' --- EVENTS ---

Public Function OnGridCellEnter()
    Dim rIndex As Long, cIndex As Long
    On Error GoTo Finally

    rIndex = Me.SelTop - 1
    cIndex = Me.SelLeft - 2

    If rIndex >= 0 And cIndex >= 0 Then
        RaiseEvent OnCellEnter(rIndex + pFirstRowIndex, cIndex + pFirstColumnIndex, GetAsyncKeyState(vbKeyControl), GetAsyncKeyState(vbKeyShift))
    End If

'    Debug.Print Printf("[ENTER] rowBookmark: %1, pageRowNum: %2, sLeftPos: %3, sTopPos: %4", rowBookmark, pageRowNum, sLeftPos, sTopPos)
Finally:
End Function


' --- HELPERS ---

Private Function Max(x As Variant, Y As Variant) As Variant: Max = IIf(x > Y, x, Y): End Function
Private Function Min(x As Variant, Y As Variant) As Variant: Min = IIf(x < Y, x, Y): End Function


' --- MISC ---

Public Function GetPageRowNum() As Long
    GetPageRowNum = CLng(Me.Bookmark(0))
    If GetPageRowNum >= pFirstInactiveRowIndex + 1 Then GetPageRowNum = 0
End Function
