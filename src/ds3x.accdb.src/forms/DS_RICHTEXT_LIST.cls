Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "ds3x.UI.Misc"
Option Compare Database
Option Explicit
Option Base 0


Private pLastActiveIndex As Long
Private pScrollPosY As Long

Public Event OnActiveIndexChange(ByVal ActiveIndex As Long)
Public Event OnNoActiveIndexChange(ByVal ActiveIndex As Long)
Public Event OnItemDoubleClick(ByVal ActiveIndex As Long)


Property Get IsSubform() As Boolean
    On Error Resume Next
    IsSubform = Len(Me.Parent.Name) > 0
    On Error GoTo 0
End Property



Public Sub SetActiveIndex(ByVal Index As Long, Optional ByVal SkipEventPropagation As Boolean = False)
    Dim TargetIndex As Long

    TargetIndex = RecordsetEx.Bind(Me.Recordset.Clone).IndexOf("[DS_LIST_ITEM_INDEX] = " & CStr(Index), -1)
    If TargetIndex <> -1 Then
        Me.SelTop = TargetIndex
        DoEvents
        SetCurrentAsActiveIndex SkipEventPropagation
        DoEvents
    End If
End Sub

Private Sub DS_LIST_ITEM_OVERLAY_Click()
    On Error GoTo Finally

    If CLng(Int(Val(Nz(Me.DS_LIST_ITEM_ROW_STATE)))) = 4 Then Exit Sub
    SaveScrollPosY
    SetCurrentAsActiveIndex
    RestoreScrollPosY
    Exit Sub
Finally:
    Debug.Print "[ERROR] @DS_LIST_ITEM_OVERLAY.Click() - " & Err.Description
End Sub

Private Sub SetCurrentAsActiveIndex(Optional ByVal SkipEventPropagation As Boolean = False)
    On Error GoTo Finally
    Dim cActiveIndex As Long

    cActiveIndex = CLng(Int(Val(Nz(Me.DS_LIST_ITEM_ROW_INDEX))))
    If cActiveIndex <> pLastActiveIndex Then
        pLastActiveIndex = cActiveIndex
        Me.DS_ACTIVE_INDEX = cActiveIndex
        If Not SkipEventPropagation Then
            RaiseEvent OnActiveIndexChange(pLastActiveIndex)
        End If
    Else
        If Not SkipEventPropagation Then
            RaiseEvent OnNoActiveIndexChange(pLastActiveIndex)
        End If
    End If

    Exit Sub
Finally:
    Debug.Print "[ERROR] @DS_RICHTEXT_LIST.SetCurrentAsActiveIndex() - " & Err.Description
End Sub


Private Sub DS_LIST_ITEM_OVERLAY_DblClick(Cancel As Integer)
    On Error Resume Next

    If CLng(Int(Val(Nz(Me.DS_LIST_ITEM_ROW_STATE)))) = 4 Then Exit Sub
    RaiseEvent OnItemDoubleClick(CLng(Int(Val(Nz(Me.DS_LIST_ITEM_ROW_INDEX)))))
    On Error GoTo 0
End Sub

Private Sub Form_Load()
    pLastActiveIndex = -1
End Sub


Private Sub SaveScrollPosY()
    pScrollPosY = Me.SelTop - Round(Me.CurrentSectionTop / Me.FormDetail.Height)
End Sub

Private Sub RestoreScrollPosY()
    On Error Resume Next
    Me.SelTop = Me.Recordset.RecordCount
    Me.SelTop = pScrollPosY
    On Error GoTo 0
End Sub
