VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder "ds3x.Collections"
Option Compare Database
Option Explicit

Implements IList

Private pItems() As Variant
Private pIndex As Long      ' _baseSize
Private pBaseIndex As Long  ' _baseIndex
Private pList As IList      ' _baseList
Private pVersion As Long
Private pFlags As Byte



'@DefaultMember
Public Property Get Item(ByVal Index As Long) As Variant
Attribute Item.VB_UserMemId = 0
    InternalUpdateRange
    Assign Item, pItems(Index)
End Property

Public Property Let Item(ByVal Index As Long, ByRef Value As Variant)
    InternalUpdateRange
    If Index < 0 Or Index >= pIndex Then Err.Raise 9
    pList.Item(pBaseIndex + Index) = Value
    pVersion = pVersion + 1
End Property

Public Property Set Item(ByVal Index As Long, ByRef Value As Variant)
    InternalUpdateRange
    If Index < 0 Or Index >= pIndex Then Err.Raise 9
    Set pList.Item(pBaseIndex + Index) = Value
    pVersion = pVersion + 1
End Property

Public Property Get Items() As Variant()
    Static sa As SAFEARRAY_1D
    With sa
        .cDims = 1
        .cbElements = VARIANT_SIZE
        .fFeatures = FADF_VARIANT Or FADF_AUTO
        .cLocks = 10
        .pvData = VarPtr(pItems(0))
        .rgsabound0.cElements = pIndex
    End With
    MemLongPtr(VarPtrArr(Items)) = VarPtr(sa)
End Property

Public Property Get Count() As Long: Count = pIndex: End Property
Private Property Let Count(ByVal Value As Long): pIndex = Value: SafeArraySetElementsCount pItems, Value: End Property
Public Property Get Capacity() As Long: Capacity = pList.Capacity: End Property
Public Property Let Capacity(ByVal Value As Long): End Property
Public Property Get IsFixedSize() As Boolean: IsFixedSize = (pFlags And 2 <> 0): End Property
Public Property Get IsReadOnly() As Boolean: IsReadOnly = (pFlags And 1 <> 0): End Property


Public Function Bind(ByRef TargetList As IList, Optional ByVal Index As Long = 0, Optional ByVal GetCount As Variant, Optional ByVal ArrayElementSize As Long = VARIANT_SIZE) As ListRange
    Static sa As SAFEARRAY_1D
    If IsMissing(GetCount) Then GetCount = TargetList.Count - Index
    pBaseIndex = Index
    pIndex = GetCount
    Set pList = TargetList
    pVersion = TargetList.Version
    pFlags = IIf(pList.IsFixedSize, CByte(2), CByte(0))
    If pList.IsReadOnly Then pFlags = (pFlags Or CByte(1))
    With sa
        .cDims = 1
        .cbElements = ArrayElementSize
        .fFeatures = IIf(ArrayElementSize = VARIANT_SIZE, FADF_VARIANT Or FADF_EMBEDDED, FADF_EMBEDDED)
        .cLocks = 10
        .pvData = TargetList.GetAddressOfItemAt(Index)
        .rgsabound0.cElements = GetCount
    End With
    MemLongPtr(VarPtrArr(pItems)) = VarPtr(sa)
    Set Bind = Me
End Function

'@Enumerator
Public Function NewEnum() As IEnumVARIANT
    Set NewEnum = GetNewEnum(Me, pIndex - 1, 0)
End Function

Public Function Add(ByRef Value As Variant) As Long
    InternalUpdateRange
    pList.Insert pBaseIndex + pIndex, Value
    Count = pIndex + 1
    Add = pIndex
    pVersion = pVersion + 1
End Function

Public Sub Insert(ByVal Index As Long, ByRef Value As Variant)
    InternalUpdateRange
    pList.Insert pBaseIndex + Index, Value
    Count = pIndex + 1
    pVersion = pVersion + 1
End Sub

Public Sub AddRange(ByRef ArrayLike As Variant)
    Dim GetCount As Long
    InternalUpdateRange
    If IsObject(ArrayLike) Then
        GetCount = ArrayLike.Count
    Else
        GetCount = (UBound(ArrayLike) - LBound(ArrayLike)) + 1
    End If
    If GetCount > 0 Then
        pList.InsertRange pBaseIndex + pIndex, ArrayLike
        Count = pIndex + GetCount
        pVersion = pVersion + 1
    End If
End Sub

Public Sub InsertRange(ByVal Index As Long, ByRef ArrayLike As Variant)
    Dim GetCount As Long
    InternalUpdateRange
    If IsObject(ArrayLike) Then
        GetCount = ArrayLike.Count
    Else
        GetCount = (UBound(ArrayLike) - LBound(ArrayLike)) + 1
    End If
    If GetCount > 0 Then
        pList.InsertRange pBaseIndex + Index, ArrayLike
        Count = pIndex + GetCount
        pVersion = pVersion + 1
    End If
End Sub

Public Function GetRange(ByVal Index As Long, ByVal GetCount As Long) As ListRange
    InternalUpdateRange
    If Index < 0 Or GetCount < 0 Or pIndex - Index < GetCount Then Err.Raise 9
    Set GetRange = New ListRange
    GetRange.Bind Me, Index, GetCount
End Function

Public Function GetEnumerator(Optional ByVal Index As Long = 0, Optional ByVal GetCount As Variant) As IEnumVARIANT
    If IsMissing(GetCount) Then GetCount = pIndex - Index
    Set GetEnumerator = GetNewEnum(Me, GetCount - 1, Index)
End Function

Public Function BinarySearch(ByVal Index As Long, ByVal GetCount As Long, ByRef Value As Variant, Optional ByRef Comparer As IComparer = Nothing) As Long
    InternalUpdateRange
    If Index < 0 Or GetCount < 0 Or pIndex - Index < GetCount Then Err.Raise 9
    Dim i As Long
    i = pList.BinarySearch(pBaseIndex + Index, GetCount, Value, Comparer)
    If i >= 0 Then
        BinarySearch = i - pBaseIndex
    Else
        BinarySearch = i + pBaseIndex
    End If
End Function

Public Sub Clear()
    InternalUpdateRange
    If pIndex > 0 Then
        pList.RemoveRange pBaseIndex, pIndex
        pVersion = pVersion + 1
        Count = 0
    End If
End Sub

'// Clones this ListRange as an ArrayList, doing a shallow copy.
Public Function Clone() As ArrayList
    Dim Target As New ArrayList
    C2IList(Target).CloneTo Target, C2IList(Me), 0, pIndex, pFlags
    Set Clone = Target
End Function

Public Function Contains(ByRef Value As Variant) As Boolean
    Dim i As Long
    If IsNull(Value) Then
        For i = 0 To pIndex - 1
            If IsNull(pItems(i)) Then
                Contains = True
                Exit For
            End If
        Next i
    ElseIf IsEmpty(Value) Then
        For i = 0 To pIndex - 1
            If IsEmpty(pItems(i)) Then
                Contains = True
                Exit For
            End If
        Next i
    ElseIf IsObject(Value) Then
        If Value Is Nothing Then
            For i = 0 To pIndex - 1
                If IsObject(pItems(i)) Then
                    If pItems(i) Is Nothing Then
                        Contains = True
                        Exit For
                    End If
                End If
            Next i
        Else
            For i = 0 To pIndex - 1
                If IsObject(pItems(i)) Then
                    If ObjPtr(pItems(i)) = ObjPtr(Value) Then
                        Contains = True
                        Exit For
                    End If
                End If
            Next i
        End If
    Else
        For i = 0 To pIndex - 1
            If pItems(i) = Value Then
                Contains = True
                Exit For
            End If
        Next i
    End If
End Function

Public Sub CopyTo(Optional ByVal Index As Long = 0, Optional ByRef TargetArray As Variant, Optional ByVal TargetIndex As Long = 0, Optional ByVal GetCount As Variant)
    ' TODO
    Err.Raise &H80004001
End Sub

Public Function FixedSize(ByRef ArrayLike As Variant) As ListRange
    ' TODO
    Err.Raise &H80004001
End Function

Public Function ReadOnly(ByRef ArrayLike As Variant) As ListRange
    ' TODO
    Err.Raise &H80004001
End Function

Public Function IndexOf(ByRef Value As Variant, Optional ByVal Index As Long = 0, Optional ByVal GetCount As Variant) As Long
    InternalUpdateRange
    If Index < 0 Or Index > pIndex Or GetCount < 0 Or (Index > pIndex - GetCount) Then Err.Raise 9
    Dim i As Long
    i = pList.IndexOf(Value, pBaseIndex + Index, GetCount)
    If i >= 0 Then
        IndexOf = i - pBaseIndex
    Else
        IndexOf = -1
    End If
End Function

Public Function LastIndexOf(ByRef Value As Variant, Optional ByVal Index As Long = 0, Optional ByVal GetCount As Variant) As Long
    InternalUpdateRange
    If Index >= pIndex Or Index < 0 Then Err.Raise 9
    Dim i As Long
    LastIndexOf = -1
    If pIndex > 0 Then
        i = pList.LastIndexOf(Value, pBaseIndex + Index, GetCount)
        If i >= 0 Then
            LastIndexOf = i - pBaseIndex
        Else
            LastIndexOf = -1
        End If
    End If
End Function

Public Sub Remove(ByRef Value As Variant)
    Dim Index As Long
    Index = IndexOf(Value)
    If Index >= 0 Then RemoveAt Index
End Sub

Public Sub RemoveAt(ByVal Index As Long)
    InternalUpdateRange
    If Index < 0 Or Index >= pIndex Then Err.Raise 9
    pList.RemoveAt pBaseIndex + Index
    Count = pIndex - 1
    pVersion = pVersion + 1
End Sub

Public Sub RemoveRange(ByVal Index As Long, ByVal GetCount As Long)
    InternalUpdateRange
    If Index < 0 Or GetCount < 0 Or (pIndex - Index < GetCount) Then Err.Raise 9
    If GetCount > 0 Then
        pList.RemoveRange pBaseIndex + Index, GetCount
        Count = pIndex - GetCount
        pVersion = pVersion + 1
    End If
End Sub

Public Sub Reverse(Optional ByVal Index As Long = 0, Optional ByVal GetCount As Variant)
    InternalUpdateRange
    If IsMissing(GetCount) Then GetCount = pIndex - Index
    If Index < 0 Or GetCount < 0 Or (pIndex - Index < GetCount) Then Err.Raise 9
    pList.Reverse pBaseIndex + Index, GetCount
    pVersion = pVersion + 1
End Sub

Public Sub SetRange(ByVal Index As Long, ByRef ArrayLike As Variant)
    InternalUpdateRange
    If Index < 0 Or Index >= pIndex Then Err.Raise 9
    pList.SetRange pBaseIndex + Index, ArrayLike
    pVersion = pVersion + 1
End Sub

Public Sub Sort(Optional ByVal Index As Long = 0, Optional ByVal GetCount As Variant, Optional ByRef Comparer As IComparer = Nothing)
    InternalUpdateRange
    If IsMissing(GetCount) Then GetCount = pIndex - Index
    If Index < 0 Or GetCount < 0 Or (pIndex - Index < GetCount) Then Err.Raise 9
    pList.Sort pBaseIndex + Index, GetCount, Comparer
    pVersion = pVersion + 1
End Sub

Public Function ToArray() As Variant()
    Dim t() As Variant
    If pIndex = 0 Then
        ToArray = Array()
        Exit Function
    End If
    ReDim t(0 To pIndex - 1)
    VariantArrayClone VarPtr(t(0)), VarPtr(pItems(0)), pIndex
    ReassignArrayTo ToArray, t
End Function



Private Sub Class_Terminate()
    On Error Resume Next
    MemLongPtr(VarPtrArr(pItems)) = CLngPtr(0)
    Set pList = Nothing
    On Error GoTo 0
End Sub

Private Sub InternalUpdateRange()
    If pVersion <> pList.Version Then Err.Raise vbObjectError + 988425, , "This ArrayList range is no longer valid."
    ' This array is fixed or temporarily locked (Error 10)
End Sub

Private Function GetAddressOfItemAt(ByVal Index As Long) As LongPtr
    GetAddressOfItemAt = VarPtr(pItems(Index))
End Function

Private Sub Assign(ByRef Target As Variant, ByRef Value As Variant)
    If IsObject(Value) Then
        Set Target = Value
    Else
        Target = Value
    End If
End Sub

Private Function CloneTo(ByVal Target As ArrayList, ByRef Source As IList, ByVal Index As Long, ByVal GetCount As Long, ByVal Flags As Byte) As Boolean
    Err.Raise &H80004001
End Function

Private Sub SafeArraySetElementsCount(ByRef TargetArray() As Variant, ByVal Value As Long)
    MemLong((Not Not TargetArray) + (VARIANT_SIZE / INT_SIZE) + 4 + PTR_SIZE) = Value
End Sub


Private Function C2IList(ByRef Value As Variant) As IList: Set C2IList = Value: End Function


' --- IList Interface ---

Private Property Get IList_Count() As Long: IList_Count = Count: End Property

Private Property Get IList_Capacity() As Long: IList_Capacity = Capacity: End Property

Private Property Let IList_Capacity(ByVal RHS As Long): Capacity = RHS: End Property

Private Property Get IList_IsFixedSize() As Boolean: IList_IsFixedSize = IsFixedSize: End Property

Private Property Get IList_IsReadOnly() As Boolean: IList_IsReadOnly = IsReadOnly: End Property

Private Property Get IList_IsSynchronized() As Boolean: IList_IsSynchronized = False: End Property

Private Property Set IList_Item(ByVal Index As Long, RHS As Variant): Set Item(Index) = RHS: End Property

Private Property Let IList_Item(ByVal Index As Long, RHS As Variant): Item(Index) = RHS: End Property

Private Property Get IList_Item(ByVal Index As Long) As Variant: Assign IList_Item, pItems(Index): End Property

Private Property Get IList_Items() As Variant(): IList_Items = Items(): End Property

Private Property Get IList_SyncRoot() As Object: End Property

Private Property Get IList_Version() As Long: IList_Version = pVersion: End Property

Private Function IList_Add(Value As Variant) As Long: IList_Add = Add(Value): End Function

Private Sub IList_AddRange(ArrayLike As Variant): AddRange ArrayLike: End Sub

Private Function IList_BinarySearch(ByVal Index As Long, ByVal GetCount As Long, Value As Variant, Optional Comparer As IComparer = Nothing) As Long: IList_BinarySearch = BinarySearch(Index, GetCount, Value, Comparer): End Function

Private Sub IList_Clear(): Clear: End Sub

Private Function IList_Clone() As IList: Set IList_Clone = Clone: End Function

Private Function IList_Contains(Value As Variant) As Boolean: IList_Contains = Contains(Value): End Function

Private Sub IList_CopyTo(Optional ByVal Index As Long = 0&, Optional TargetArray As Variant, Optional ByVal TargetIndex As Long = 0&, Optional ByVal GetCount As Variant): CopyTo Index, TargetArray, TargetIndex, GetCount: End Sub

Private Function IList_GetRange(ByVal Index As Long, ByVal GetCount As Long) As IList: Set IList_GetRange = GetRange(Index, GetCount): End Function

Private Function IList_IndexOf(Value As Variant, Optional ByVal Index As Long = 0&, Optional ByVal GetCount As Variant) As Long: IList_IndexOf = IndexOf(Value, Index, GetCount): End Function

Private Sub IList_Insert(ByVal Index As Long, Value As Variant): Insert Index, Value: End Sub

Private Sub IList_InsertRange(ByVal Index As Long, ArrayLike As Variant): InsertRange Index, ArrayLike: End Sub

Private Function IList_LastIndexOf(Value As Variant, Optional ByVal Index As Long = 0&, Optional ByVal GetCount As Variant) As Long: IList_LastIndexOf = LastIndexOf(Value, Index, GetCount): End Function

Private Sub IList_Remove(Value As Variant): Remove Value: End Sub

Private Sub IList_RemoveAt(ByVal Index As Long): RemoveAt Index: End Sub

Private Sub IList_RemoveRange(ByVal Index As Long, ByVal GetCount As Long): RemoveRange Index, GetCount: End Sub

Private Sub IList_Reverse(Optional ByVal Index As Long = 0&, Optional ByVal GetCount As Variant): Reverse Index, GetCount: End Sub

Private Sub IList_SetRange(ByVal Index As Long, ArrayLike As Variant): SetRange Index, ArrayLike: End Sub

Private Sub IList_Sort(Optional ByVal Index As Long = 0&, Optional ByVal GetCount As Variant, Optional Comparer As IComparer = Nothing): Sort Index, GetCount, Comparer: End Sub

Private Function IList_ToArray() As Variant(): IList_ToArray = ToArray: End Function

Private Function IList_GetAddressOfItemAt(ByVal Index As Long) As LongPtr: IList_GetAddressOfItemAt = GetAddressOfItemAt(Index): End Function

Private Function IList_CloneTo(ByVal Target As ArrayList, Source As IList, ByVal Index As Long, ByVal GetCount As Long, ByVal Flags As Byte) As Boolean: IList_CloneTo = CloneTo(Target, Source, Index, GetCount, Flags): End Function
