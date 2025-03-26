Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "ds3x.UI.Scrollview"
Option Compare Database
Option Explicit
Option Base 0

Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Public Event OnColumnLetterClick(ByVal ColumnIndex As Long, ByVal CtrlKey As Boolean, ByVal ShiftKey As Boolean)
Public Event OnColumnNameWillChange(ByVal ColumnIndex As Long, ByVal Value As String)

Private pViewport As Form_DS_VIEWPORT
Private pSelectedColumns As ArrayListEx

Private Const pMaxAvailColumns As Long = 16
Private Const pGridCellSizeX As Long = 1965
Private Const pGridCellSizeY As Long = 375

Private Const pHeaderButtonBackColorSelected As Long = 13434879
Private Const pHeaderButtonBackColorNormal As Long = 15132391
Private pFirstInactiveRowIndex As Long
Private pFirstInactiveColumnIndex As Long
Private pRemoveFocusNeeded As Boolean


Public Property Get Viewport() As Form_DS_VIEWPORT: Set Viewport = pViewport: End Property
Public Property Set Viewport(ByRef Value As Form_DS_VIEWPORT): Set pViewport = Value: End Property

Public Property Get MaxAvailColumns() As Long: MaxAvailColumns = pMaxAvailColumns: End Property
' Maximum available form width filled with table-related controls (almost 22 inches ~= 55.478cm ~= 31456 twips in my dev environment).
Public Property Get MaxContentWidthLimit() As Long: MaxContentWidthLimit = pGridCellSizeX * pMaxAvailColumns: End Property
' Default Cell/Column width
Public Property Get GridCellSizeX() As Long: GridCellSizeX = pGridCellSizeX: End Property
Public Property Get GridCellSizeY() As Long: GridCellSizeY = pGridCellSizeY: End Property


' --- FORM EVENTS ---

Private Sub Form_Load()
    Set pSelectedColumns = ArrayListEx.Create()
End Sub

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
            Me.DS_HC_1_1.Left = 0
            Me.DS_HC_2_1.Left = 0
        End If
        Me.Controls("DS_HC_1_" & CStr(i)).Width = pGridCellSizeX
        Me.Controls("DS_HC_2_" & CStr(i)).Width = pGridCellSizeX
    Next i

    Me.DS_HR_BG.Left = 0
    Me.DS_HR_BG.Width = pGridCellSizeX * pMaxAvailColumns
    Me.Width = pGridCellSizeX * pMaxAvailColumns
End Sub

Friend Sub SetupGrid(ByVal FirstColumnIndex As Long, ByVal FirstRowIndex As Long, ByRef dsT As dsTable)
    Dim i As Long, sNames As Variant

    sNames = dsT.GetColumnNames(FirstColumnIndex, pMaxAvailColumns)
    For i = 1 To pMaxAvailColumns
        With Me.Controls("DS_HC_1_" & CStr(i))
            .Caption = CollectionsLib.ColumnLetter(FirstColumnIndex + i - 1)
            .BackColor = IIf(pSelectedColumns.Contains(FirstColumnIndex + i - 1), pHeaderButtonBackColorSelected, pHeaderButtonBackColorNormal)
        End With
        Me.Controls("DS_HC_2_" & CStr(i)) = sNames(i - 1)
    Next i

    pFirstInactiveColumnIndex = Max(Min(dsT.ColumnCount - FirstColumnIndex, pMaxAvailColumns), 0)
    pFirstInactiveRowIndex = Max(dsT.Count - FirstRowIndex, 0)
End Sub


Friend Sub SetSelectedColumns(ByRef aX As ArrayListEx)
    Dim FirstColumnIndex As Long, i As Long

    Set pSelectedColumns = aX
    FirstColumnIndex = pViewport.FirstColumnIndex
    For i = 1 To pMaxAvailColumns
        Me.Controls("DS_HC_1_" & CStr(i)).BackColor = IIf(aX.Contains(FirstColumnIndex + i - 1), pHeaderButtonBackColorSelected, pHeaderButtonBackColorNormal)
    Next i
End Sub


' --- UI EVENT HANDLERS ---

Public Function OnColumnHeaderClick()
    Dim cIndex As Long

    cIndex = Me.SelLeft - 2
    Me.HiddenControl.SetFocus
    RaiseEvent OnColumnLetterClick(pViewport.FirstColumnIndex + cIndex, GetAsyncKeyState(vbKeyControl), GetAsyncKeyState(vbKeyShift))
End Function

Public Function OnColumnNameChange()
    Dim Target As Access.TextBox

    If TryGetActiveControl(Target) Then
        RaiseEvent OnColumnNameWillChange(pViewport.FirstColumnIndex + (Me.SelLeft - (pMaxAvailColumns + 0) - 2), GetControlText(Target))
    End If
End Function

Public Function OnEnterColumnName()
    pRemoveFocusNeeded = True
    DoEvents
    With pViewport.Scrollview
        .TriggerOnClearSelectionRequest
    End With
End Function

Public Sub RemoveFocus()
    If Not pRemoveFocusNeeded Then Exit Sub

    pRemoveFocusNeeded = False
    Me.HiddenControl.SetFocus
End Sub


' --- UTILITY HELPERS ---

Private Function GetControlText(ByRef TargetControl As Access.control) As String
    On Error Resume Next
    GetControlText = TargetControl.Value
    GetControlText = TargetControl.Text
    On Error GoTo 0
End Function

Private Function TryGetActiveControl(ByRef OutActiveControl As Variant) As Boolean
    On Error GoTo Finally

    Set OutActiveControl = Screen.ActiveControl
    TryGetActiveControl = True
Finally:
End Function

Private Function Max(x As Variant, Y As Variant) As Variant: Max = IIf(x > Y, x, Y): End Function
Private Function Min(x As Variant, Y As Variant) As Variant: Min = IIf(x < Y, x, Y): End Function
