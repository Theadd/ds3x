VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder "ds3x.Collections"
'// Implements a variable-size List that uses an array of variants to store the
'// elements. A ArrayList has a capacity, which is the allocated length
'// of the internal array. As elements are added to a ArrayList, the capacity
'// of the ArrayList is automatically increased as required by reallocating the
'// internal array.

' ArrayList.cs: https://referencesource.microsoft.com/#mscorlib/system/collections/arraylist.cs
Option Compare Database
Option Explicit

Implements IList

Private pItems() As Variant
Private pIndex As Long
Private pVersion As Long
Private pFlags As Byte


'@DefaultMember
Public Property Get Item(ByVal Index As Long) As Variant
Attribute Item.VB_UserMemId = 0
    Assign Item, pItems(Index)
End Property

Public Property Let Item(ByVal Index As Long, ByRef Value As Variant)
    pItems(Index) = Value
    pVersion = pVersion + 1
End Property

Public Property Set Item(ByVal Index As Long, ByRef Value As Variant)
    Set pItems(Index) = Value
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
Public Property Get Capacity() As Long: Capacity = UBound(pItems) + 1: End Property
Public Property Let Capacity(ByVal Value As Long): ReDim Preserve pItems(0 To Value - 1): End Property
Public Property Get IsFixedSize() As Boolean: IsFixedSize = (pFlags And 2 <> 0): End Property
Public Property Get IsReadOnly() As Boolean: IsReadOnly = (pFlags And 1 <> 0): End Property
Public Property Get IsSynchronized() As Boolean: IsSynchronized = False: End Property
Public Property Get SyncRoot() As Object: End Property

'@Deprecated
Public Function Create() As ArrayList
    Set Create = New ArrayList
End Function

'@Enumerator
Public Function NewEnum() As IEnumVARIANT
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = GetNewEnum(Me, pIndex - 1, 0)
End Function

Public Function Add(ByRef Value As Variant) As Long
    If pIndex > UBound(pItems) Then GrowCapacity 1
    If IsObject(Value) Then
        Set pItems(pIndex) = Value
    Else
        pItems(pIndex) = Value
    End If
    Add = pIndex
    pIndex = pIndex + 1
    pVersion = pVersion + 1
End Function

Public Sub Insert(ByVal Index As Long, ByRef Value As Variant)
    ReserveSpaceForInsert Index, 1
    If IsObject(Value) Then
        Set pItems(Index) = Value
    Else
        pItems(Index) = Value
    End If
    pIndex = pIndex + 1
    pVersion = pVersion + 1
End Sub

Public Sub AddRange(ByRef ArrayLike As Variant)
    InsertRange pIndex, ArrayLike
End Sub

Public Sub InsertRange(ByVal Index As Long, ByRef ArrayLike As Variant)
    Dim v As Variant, GetCount As Long, isPending As Boolean

    If Not IsObject(ArrayLike) Then
        GetCount = (UBound(ArrayLike) - LBound(ArrayLike)) + 1
        If VarType(ArrayLike) = (vbArray + vbVariant) Then
            ' TODO: ReserveSpaceForInsert deallocates current pItems invalidating any ListRange of it
            ReserveSpaceForInsert Index, GetCount
            VariantArrayClone VarPtr(pItems(Index)), VarPtr(ArrayLike(LBound(ArrayLike))), GetCount
        Else
            isPending = True
        End If
    Else
        GetCount = ArrayLike.Count
        If TypeOf ArrayLike Is IList Then
            InsertListTo Index, GetCount, ArrayLike
        Else
            isPending = True
        End If
    End If
    If isPending Then
        ReserveSpaceForInsert Index, GetCount
        For Each v In ArrayLike
            If IsObject(v) Then
                Set pItems(Index) = v
            Else
                pItems(Index) = v
            End If
            Index = Index + 1
        Next v
    End If
    pIndex = pIndex + GetCount
    pVersion = pVersion + 1
End Sub

Public Function GetRange(ByVal Index As Long, ByVal GetCount As Long) As ListRange
    If Index < 0 Or GetCount < 0 Or pIndex - Index < GetCount Then Err.Raise 9
    Set GetRange = New ListRange
    GetRange.Bind Me, Index, GetCount
End Function

Public Function GetEnumerator(Optional ByVal Index As Long = 0, Optional ByVal GetCount As Variant) As IEnumVARIANT
    If IsMissing(GetCount) Then GetCount = pIndex - Index
    Set GetEnumerator = GetNewEnum(Me, GetCount - 1, Index)
End Function

'// Searches a section of the list. This method assumes that the given
'// section of the list is already sorted; if this is not the case, the
'// result will be incorrect.
'// The method returns the index of the given value in the list. If the
'// list does not contain the given value, the method returns a negative
'// integer. The bitwise complement operator (Not) can be applied to a
'// negative result to produce the index of the first element (if any) that
'// is larger than the given search value. This is also the index at which
'// the search value should be inserted into the list in order for the list
'// to remain sorted.
Public Function BinarySearch(ByVal Index As Long, ByVal GetCount As Long, ByRef Value As Variant, Optional ByRef Comparer As IComparer = Nothing) As Long
    Dim lo As Long, hi As Long, i As Long
    lo = Index
    hi = Index + GetCount - 1

    If Not (Comparer Is Nothing) Then
        Do While (lo <= hi)
            i = lo + ((hi - lo) \ 2)
            Select Case Comparer.Compare(pItems(i), Value)
                Case 0
                    BinarySearch = i
                    Exit Function
                Case Is < 0: lo = i + 1
                Case Else: hi = i - 1
            End Select
        Loop
    Else
        Do While (lo <= hi)
            i = lo + ((hi - lo) \ 2)
            Select Case pItems(i)
                Case Value
                    BinarySearch = i
                    Exit Function
                Case Is < Value: lo = i + 1
                Case Else: hi = i - 1
            End Select
        Loop
    End If
    BinarySearch = Not lo
End Function

Public Sub Clear()
    Erase pItems
    ReDim pItems(0 To 1)
    pIndex = 0
    pVersion = pVersion + 1
End Sub

'// Clones this ArrayList, doing a shallow copy.  (A copy is made of all
'// Object references in the ArrayList, but the Objects pointed to
'// are not cloned).
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
    ElseIf IsError(Value) Then
        For i = 0 To pIndex - 1
            If IsError(pItems(i)) Then
                If pItems(i) = Value Then
                    Contains = True
                    Exit For
                End If
            End If
        Next i
    Else
        For i = 0 To pIndex - 1
            If Not (IsObject(pItems(i)) Or IsError(pItems(i)) Or IsArray(pItems(i))) Then
                If pItems(i) = Value Then
                    Contains = True
                    Exit For
                End If
            End If
        Next i
    End If
End Function

'// Copies a section of this list to the given array at the given index.
Public Sub CopyTo(Optional ByVal Index As Long = 0, Optional ByRef TargetArray As Variant, Optional ByVal TargetIndex As Long = 0, Optional ByVal GetCount As Variant)
    ' TODO
    Err.Raise &H80004001
End Sub

'// Returns a list wrapper that is fixed at the current size. Operations
'// that add or remove items will fail, however, replacing items is allowed.
Public Function FixedSize(ByRef ArrayLike As Variant) As ListRange
    ' TODO
    Err.Raise &H80004001
End Function

'// Returns a read-only ArrayList wrapper for the given ArrayList.
Public Function ReadOnly(ByRef ArrayLike As Variant) As ListRange
    ' TODO
    Err.Raise &H80004001
End Function

'// Searches for the specified Value and returns the zero-based index of the
'// first occurrence within the range of elements in the ArrayList that starts
'// at the specified index and contains the specified number of elements.
Public Function IndexOf(ByRef Value As Variant, Optional ByVal Index As Long = 0, Optional ByVal GetCount As Variant) As Long
    Dim i As Long
    If IsMissing(GetCount) Then GetCount = pIndex - Index
    GetCount = Index + GetCount - 1
    IndexOf = -1
    For i = Index To GetCount
        If pItems(i) = Value Then
            IndexOf = i
            Exit For
        End If
    Next i
End Function

'// Searches for the specified Value and returns the zero-based index of the
'// last occurrence within the range of elements in the ArrayList that starts
'// at the specified index and contains the specified number of elements.
Public Function LastIndexOf(ByRef Value As Variant, Optional ByVal Index As Long = 0, Optional ByVal GetCount As Variant) As Long
    Dim i As Long
    If IsMissing(GetCount) Then GetCount = pIndex - Index
    GetCount = Index + GetCount - 1
    LastIndexOf = -1
    For i = GetCount To Index Step -1
        If pItems(i) = Value Then
            LastIndexOf = i
            Exit For
        End If
    Next i
End Function

Public Sub Remove(ByRef Value As Variant)
    Dim Index As Long
    Index = IndexOf(Value)
    If Index >= 0 Then RemoveAt Index
End Sub

'// Removes the element at the given index.
Public Sub RemoveAt(ByVal Index As Long)
    If Index < 0 Or Index >= pIndex Then Err.Raise 9
    pItems(Index) = Empty
    If Index < pIndex - 1 Then MemCopy VarPtr(pItems(Index)), VarPtr(pItems(Index + 1)), (pIndex - Index - 1) * VARIANT_SIZE
    ZeroMemory VarPtr(pItems(pIndex - 1)), VARIANT_SIZE
    pIndex = pIndex - 1
    pVersion = pVersion + 1
End Sub

'// Removes a range of elements from this list.
Public Sub RemoveRange(ByVal Index As Long, ByVal GetCount As Long)
    If Index < 0 Or GetCount < 0 Or (pIndex - Index < GetCount) Then Err.Raise 9
    Dim i As Long
    For i = 0 To GetCount - 1
        pItems(Index + i) = Empty
    Next i
    If Index < pIndex - 1 Then MemCopy VarPtr(pItems(Index)), VarPtr(pItems(Index + GetCount)), (pIndex - Index - GetCount) * VARIANT_SIZE
    ZeroMemory VarPtr(pItems(pIndex - GetCount)), GetCount * VARIANT_SIZE
    pIndex = pIndex - GetCount
    pVersion = pVersion + 1
End Sub

Public Function Repeat(ByRef Value As Variant, ByVal GetCount As Long) As ArrayList
    Dim t() As Variant, i As Long
    ReDim t(0 To GetCount - 1)
    If IsObject(Value) Then
        For i = 0 To GetCount - 1
            Set t(i) = Value
        Next i
    Else
        For i = 0 To GetCount - 1
            t(i) = Value
        Next i
    End If
    Set Repeat = New ArrayList
    Repeat.AddRange t
End Function

Public Sub Reverse(Optional ByVal Index As Long = 0, Optional ByVal GetCount As Variant)
    Dim i As Long, iMax As Long
    If IsMissing(GetCount) Then GetCount = pIndex - Index
    If Index < 0 Or GetCount < 0 Or (pIndex - Index < GetCount) Then Err.Raise 9
    iMax = GetCount \ 2
    For i = 0 To iMax
        Swap pItems(Index + i), pItems(Index + (GetCount - 1) - i)
    Next i
    pVersion = pVersion + 1
End Sub

Public Sub SetRange(ByVal Index As Long, ByRef ArrayLike As Variant)
    If IsObject(ArrayLike) Then
        If Index < 0 Or Index > pIndex - ArrayLike.Count Then Err.Raise 9
        VariantArrayClone VarPtr(pItems(Index)), VarPtr(ArrayLike.Items(0)), ArrayLike.Count
    Else
        If Index < 0 Or Index > pIndex - (1 + UBound(ArrayLike)) Then Err.Raise 9
        VariantArrayClone VarPtr(pItems(Index)), VarPtr(ArrayLike(0)), 1 + UBound(ArrayLike)
    End If
    pVersion = pVersion + 1
End Sub

Public Sub Sort_2(ByRef Comparer As IComparer): Sort 0, , Comparer: End Sub
Public Sub Sort(Optional ByVal Index As Long = 0, Optional ByVal GetCount As Variant, Optional ByRef Comparer As IComparer = Nothing)
    If IsMissing(GetCount) Then GetCount = pIndex - Index
    If Comparer Is Nothing Then
        QuickSortV2 Index, Index + GetCount - 1
    Else
        QuickSort Index, Index + GetCount - 1, Comparer
    End If
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


Private Sub Class_Initialize()
    ReDim pItems(0 To 1)
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    DoEvents
    Erase pItems
    On Error GoTo 0
End Sub

Private Function GetAddressOfItemAt(ByVal Index As Long) As LongPtr
    GetAddressOfItemAt = VarPtr(pItems(Index))
End Function

Private Sub Swap(ByRef A As Variant, ByRef B As Variant)
    Dim v As Variant
    If IsObject(A) Then
        Set v = A
    Else
        v = A
    End If
    If IsObject(B) Then
        Set A = B
    Else
        A = B
    End If
    If IsObject(v) Then
        Set B = v
    Else
        B = v
    End If
End Sub

Private Sub Assign(ByRef Target As Variant, ByRef Value As Variant)
    If IsObject(Value) Then
        Set Target = Value
    Else
        Target = Value
    End If
End Sub

Private Sub GrowCapacity(ByVal Increment As Long)
    If pIndex + Increment > UBound(pItems) Then
        If (UBound(pItems) * 2) + 2 >= pIndex + Increment Then
            ReDim Preserve pItems(0 To (UBound(pItems) * 2) + 2)
        Else
            ReDim Preserve pItems(0 To pIndex + Increment)
        End If
    End If
End Sub

Private Function GetCapacityFor(ByVal Increment As Long) As Long
    If pIndex + Increment > UBound(pItems) Then
        If (UBound(pItems) * 2) + 2 >= pIndex + Increment Then
            GetCapacityFor = (UBound(pItems) * 2) + 2
        Else
            GetCapacityFor = pIndex + Increment
        End If
    Else
        GetCapacityFor = UBound(pItems)
    End If
End Function

'Private Sub ReserveSpaceForInsert(ByVal Index As Long, ByVal GetCount As Long)
'    Dim t() As Variant, cap As Long
'    cap = GetCapacityFor(GetCount)
'    ReDim t(0 To cap)
'    If Index > 0 Then MemCopy VarPtr(t(0)), VarPtr(pItems(0)), Index * VARIANT_SIZE
'    MemCopy VarPtr(t(Index + GetCount)), VarPtr(pItems(Index)), (pIndex - Index) * VARIANT_SIZE
'    ZeroMemory VarPtr(pItems(0)), pIndex * VARIANT_SIZE
'    Erase pItems
'    ReassignArrayTo pItems, t
'End Sub

Private Sub ReserveSpaceForInsert(ByVal Index As Long, ByVal GetCount As Long)
    If pIndex + GetCount >= UBound(pItems) Then GrowCapacity GetCount
    ' CopyMemory ByVal VarPtr(pItems(Index + GetCount)), ByVal VarPtr(pItems(Index)), (pIndex - Index) * VARIANT_SIZE
    ChunkMemCopy VarPtr(pItems(Index + GetCount)), VarPtr(pItems(Index)), (pIndex - Index) * VARIANT_SIZE
    If GetCount = 1 Then
        MemInt(VarPtr(pItems(Index))) = vbEmpty
    Else
        ZeroMemory VarPtr(pItems(Index)), GetCount * VARIANT_SIZE
    End If
End Sub

Private Function IsContainedInMemoryRange(ByRef TargetList As IList, Optional ByVal Index As Long = 0, Optional ByVal GetCount As Variant) As Boolean
    If IsMissing(GetCount) Then GetCount = pIndex - Index
    If GetCount = 0 Then Exit Function
    IsContainedInMemoryRange = (TargetList.GetAddressOfItemAt(0) <= VarPtr(pItems(Index + GetCount - 1)) _
                                And TargetList.GetAddressOfItemAt(TargetList.Count - 1) >= VarPtr(pItems(Index)))
End Function

Private Sub InsertListTo(ByVal Index As Long, ByVal GetCount As Long, ByVal Target As IList)
    Dim pv0 As LongPtr, pv0base As LongPtr, pvIdx As LongPtr, t() As Variant, isContained As Boolean
    If GetCount > 0 Then
        pv0 = Target.GetAddressOfItemAt(0)
        pv0base = VarPtr(pItems(0))
        If IsContainedInMemoryRange(Target, Index) Then
            t = Target.ToArray()
            pv0 = VarPtr(t(0))
        Else
            isContained = IsContainedInMemoryRange(Target)
        End If
        GrowCapacity GetCount
        If Index < pIndex Then
            Select Case MemoryMovingMode
                Case MemoryMoveMode.MemCopyMode
                    MemCopy VarPtr(pItems(Index + GetCount)), VarPtr(pItems(Index)), (pIndex - Index) * VARIANT_SIZE
                    ZeroMemory VarPtr(pItems(Index)), GetCount * VARIANT_SIZE
                Case MemoryMoveMode.CopyMemoryMode
                    CopyMemory ByVal VarPtr(pItems(Index + GetCount)), ByVal VarPtr(pItems(Index)), (pIndex - Index) * VARIANT_SIZE
                    ZeroMemory VarPtr(pItems(Index)), GetCount * VARIANT_SIZE
                Case Else
                    ' Move memory `Target.Count` positions to the right, from `Index` position, due to memory overlapping,
                    ' it is done using a separate memory region as a helper. MemCopy automatically solves the memory
                    ' overlapping problem itself but I was getting app crashes when copying very large regions of memory.
                    If CreateMemoryCopy(pvIdx, VarPtr(pItems(Index)), (pIndex - Index) * VARIANT_SIZE) Then
                        ' SafeArrayCopyData releases any resources in destination array and those, if present, are
                        ' just copied to the right without increasing reference count so, we clear it beforehand to prevent
                        ' SafeArrayCopyData to release those resources.
                        ZeroMemory VarPtr(pItems(Index)), GetCount * VARIANT_SIZE
                        ' Copy back the elements from the temporal memory region to it's final position within the list.
                        Select Case MemoryMovingMode
                            Case MemoryMoveMode.MemAllocMemCopyMode
                                MemCopy VarPtr(pItems(Index + GetCount)), pvIdx, (pIndex - Index) * VARIANT_SIZE
                            Case MemoryMoveMode.MemAllocCopyMemoryMode
                                CopyMemory ByVal VarPtr(pItems(Index + GetCount)), ByVal pvIdx, (pIndex - Index) * VARIANT_SIZE
                            Case Else: Err.Raise 9, "ArrayList.InsertListTo", "Unexpected MemoryMoveMode"
                        End Select
                        FreeMemoryCopy pvIdx
                    Else
                        Err.Raise 7
                    End If
            End Select
        End If
        If VarPtr(pItems(0)) <> pv0base And isContained Then
            ' When `Target` derived from an array range contained within this list and, due to `Redim Preserve` from
            ' `GrowCapacity` above, the array gets reallocated to a different memory region, we've to update the pointer.
            pv0 = VarPtr(pItems(0)) + (pv0 - pv0base)
        End If
        ' Finally, insert (copy) all items in `Target` list to the region we just cleared for them.
        VariantArrayClone VarPtr(pItems(Index)), pv0, GetCount
    End If
End Sub

Private Sub QuickSort(l As Long, r As Long, ByRef Comparer As IComparer)
    Dim p As Variant, l0 As Long, r0 As Long
    l0 = l: r0 = r
    Assign p, pItems((l + r) \ 2)
    Do While (l0 <= r0)
        Comparer.Compare pItems(l0), p
        Do While (Comparer.Compare(pItems(l0), p) < 0 And l0 < r): l0 = l0 + 1: Loop
        Do While (Comparer.Compare(p, pItems(r0)) < 0 And r0 > l): r0 = r0 - 1: Loop
        If (l0 <= r0) Then
            Swap pItems(l0), pItems(r0)
            l0 = l0 + 1
            r0 = r0 - 1
        End If
    Loop
    If (l < r0) Then QuickSort l, r0, Comparer
    If (l0 < r) Then QuickSort l0, r, Comparer
End Sub

Private Sub QuickSortV2(l As Long, r As Long)
    Dim p As Variant, l0 As Long, r0 As Long
    l0 = l: r0 = r
    Assign p, pItems((l + r) \ 2)
    Do While (l0 <= r0)
        Do While (pItems(l0) < p And l0 < r): l0 = l0 + 1: Loop
        Do While (p < pItems(r0) And r0 > l): r0 = r0 - 1: Loop
        If (l0 <= r0) Then
            Swap pItems(l0), pItems(r0)
            l0 = l0 + 1
            r0 = r0 - 1
        End If
    Loop
    If (l < r0) Then QuickSortV2 l, r0
    If (l0 < r) Then QuickSortV2 l0, r
End Sub

Private Function CloneTo(ByVal Target As ArrayList, ByRef Source As IList, ByVal Index As Long, ByVal GetCount As Long, ByVal Flags As Byte) As Boolean
    pVersion = Source.Version
    Capacity = GetCount
    VariantArrayClone VarPtr(pItems(0)), Source.GetAddressOfItemAt(Index), GetCount
    pFlags = Flags
    pIndex = GetCount
End Function


Private Function C2ArrayList(ByRef Value As Variant) As ArrayList: Set C2ArrayList = Value: End Function
Private Function C2IList(ByRef Value As Variant) As IList: Set C2IList = Value: End Function


' --- IList Interface ---

Private Property Get IList_Count() As Long: IList_Count = Count: End Property

Private Property Get IList_Capacity() As Long: IList_Capacity = Capacity: End Property

Private Property Let IList_Capacity(ByVal RHS As Long): Capacity = RHS: End Property

Private Property Get IList_IsFixedSize() As Boolean: IList_IsFixedSize = IsFixedSize: End Property

Private Property Get IList_IsReadOnly() As Boolean: IList_IsReadOnly = IsReadOnly: End Property

Private Property Get IList_IsSynchronized() As Boolean: IList_IsSynchronized = IsSynchronized: End Property

Private Property Set IList_Item(ByVal Index As Long, RHS As Variant): Set Item(Index) = RHS: End Property

Private Property Let IList_Item(ByVal Index As Long, RHS As Variant): Item(Index) = RHS: End Property

Private Property Get IList_Item(ByVal Index As Long) As Variant: Assign IList_Item, Item(Index): End Property

Private Property Get IList_Items() As Variant(): IList_Items = Items(): End Property

Private Property Get IList_SyncRoot() As Object: Set IList_SyncRoot = SyncRoot: End Property

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
