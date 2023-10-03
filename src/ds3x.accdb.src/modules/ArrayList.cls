VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'// Implements a variable-size List that uses an array of variants to store the
'// elements. A ArrayList has a capacity, which is the allocated length
'// of the internal array. As elements are added to a ArrayList, the capacity
'// of the ArrayList is automatically increased as required by reallocating the
'// internal array.

' ArrayList.cs: https://referencesource.microsoft.com/#mscorlib/system/collections/arraylist.cs
Option Compare Database
Option Explicit

Private pItems() As Variant
Private pIndex As Long


'@DefaultMember
Public Property Get Item(ByVal Index As Long) As Variant
Attribute Item.VB_UserMemId = 0
    Assign Item, pItems(Index)
End Property

Public Property Let Item(ByVal Index As Long, ByRef Value As Variant)
    pItems(Index) = Value
End Property

Public Property Set Item(ByVal Index As Long, ByRef Value As Variant)
    Set pItems(Index) = Value
End Property

Public Property Get Items() As Variant()
    Static r As New ListRange
    Items = r.Bind(pItems, 0, pIndex).Items
End Property

Public Property Get Count() As Long: Count = pIndex: End Property
Public Property Get Capacity() As Long: Capacity = UBound(pItems) + 1: End Property
Public Property Let Capacity(ByVal Value As Long): ReDim Preserve pItems(0 To Value - 1): End Property
Public Property Get IsFixedSize() As Boolean: IsFixedSize = False: End Property
Public Property Get IsReadOnly() As Boolean: IsReadOnly = False: End Property
Public Property Get IsSynchronized() As Boolean: IsSynchronized = False: End Property
Public Property Get SyncRoot() As Object: End Property

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
End Function

Public Sub Insert(ByVal Index As Long, ByRef Value As Variant)
    ReserveSpaceForInsert Index, 1
    If IsObject(Value) Then
        Set pItems(Index) = Value
    Else
        pItems(Index) = Value
    End If
    pIndex = pIndex + 1
End Sub

Public Sub AddRange(ByRef ArrayLike As Variant)
    Dim v As Variant, GetCount As Long

    If Not IsObject(ArrayLike) Then
        GetCount = (UBound(ArrayLike) - LBound(ArrayLike)) + 1
        GrowCapacity GetCount
        If VarType(ArrayLike) = (vbArray + vbVariant) Then
            VariantArrayClone VarPtr(pItems(pIndex)), VarPtr(ArrayLike(LBound(ArrayLike))), GetCount
            pIndex = pIndex + GetCount
            Exit Sub
        End If
    Else
        GetCount = ArrayLike.Count
        GrowCapacity GetCount
        If TypeOf ArrayLike Is ArrayList Then
            VariantArrayClone VarPtr(pItems(pIndex)), C2ArrayList(ArrayLike).GetAddressOfItemAt(0), GetCount
            pIndex = pIndex + GetCount
            Exit Sub
        End If
    End If
    For Each v In ArrayLike
        If IsObject(v) Then
            Set pItems(pIndex) = v
        Else
            pItems(pIndex) = v
        End If
        pIndex = pIndex + 1
    Next v
End Sub

Public Sub InsertRange(ByVal Index As Long, ByRef ArrayLike As Variant)
    Dim v As Variant, GetCount As Long

    If Not IsObject(ArrayLike) Then
        GetCount = (UBound(ArrayLike) - LBound(ArrayLike)) + 1
        ReserveSpaceForInsert Index, GetCount
        If VarType(ArrayLike) = (vbArray + vbVariant) Then
            VariantArrayClone VarPtr(pItems(Index)), VarPtr(ArrayLike(LBound(ArrayLike))), GetCount
            pIndex = pIndex + GetCount
            Exit Sub
        End If
    Else
        GetCount = ArrayLike.Count
        ReserveSpaceForInsert Index, GetCount
        If TypeOf ArrayLike Is ArrayList Then
            VariantArrayClone VarPtr(pItems(Index)), C2ArrayList(ArrayLike).GetAddressOfItemAt(0), GetCount
            pIndex = pIndex + GetCount
            Exit Sub
        End If
    End If
    For Each v In ArrayLike
        If IsObject(v) Then
            Set pItems(Index) = v
        Else
            pItems(Index) = v
        End If
        Index = Index + 1
    Next v
    pIndex = pIndex + GetCount
End Sub

Public Function GetRange(ByVal Index As Long, ByVal GetCount As Long) As ArrayList
    Dim t() As Variant, sa As SAFEARRAY_1D
    Set GetRange = New ArrayList
    On Error GoTo Finally
    With sa
        .cDims = 1
        .cbElements = VARIANT_SIZE
        .pvData = VarPtr(pItems(Index))
        .rgsabound0.cElements = GetCount
    End With
    MemLongPtr(VarPtrArr(t)) = VarPtr(sa)
    GetRange.AddRange t
Finally:
    On Error Resume Next
    MemLongPtr(VarPtrArr(t)) = CLngPtr(0)
    On Error GoTo 0
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
End Sub

Public Function Clone() As ArrayList
    Set Clone = New ArrayList
    Clone.AddRange Me
End Function

Public Function Contains(ByRef Value As Variant) As Boolean
    Dim i As Long
    For i = 0 To pIndex - 1
        If pItems(i) = Value Then
            Contains = True
            Exit For
        End If
    Next i
End Function

Public Sub CopyTo(Optional ByVal Index As Long = 0, Optional ByRef TargetArray As Variant, Optional ByVal TargetIndex As Long = 0, Optional ByVal GetCount As Variant)
    ' TODO
    Err.Raise &H80004001
End Sub


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
    Dim i As Long
    For i = 0 To pIndex - 1
        If pItems(i) = Value Then
            pItems(i) = Empty
            If i < pIndex - 1 Then MemCopy VarPtr(pItems(i)), VarPtr(pItems(i + 1)), (pIndex - i - 1) * VARIANT_SIZE
            ZeroMemory VarPtr(pItems(pIndex - 1)), VARIANT_SIZE
            pIndex = pIndex - 1
            Exit For
        End If
    Next i
End Sub

Public Sub RemoveAt(ByVal Index As Long)
    pItems(Index) = Empty
    If Index < pIndex - 1 Then MemCopy VarPtr(pItems(Index)), VarPtr(pItems(Index + 1)), (pIndex - Index - 1) * VARIANT_SIZE
    ZeroMemory VarPtr(pItems(pIndex - 1)), VARIANT_SIZE
    pIndex = pIndex - 1
End Sub

Public Sub RemoveRange(ByVal Index As Long, ByVal GetCount As Long)
    Dim i As Long
    For i = 0 To GetCount - 1
        pItems(Index + i) = Empty
    Next i
    If Index < pIndex - 1 Then MemCopy VarPtr(pItems(Index)), VarPtr(pItems(Index + GetCount)), (pIndex - Index - GetCount) * VARIANT_SIZE
    ZeroMemory VarPtr(pItems(pIndex - GetCount)), GetCount * VARIANT_SIZE
    pIndex = pIndex - GetCount
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
    iMax = GetCount \ 2
    For i = 0 To iMax
        Swap pItems(Index + i), pItems(Index + (GetCount - 1) - i)
    Next i
End Sub

Public Sub SetRange(ByVal Index As Long, ByRef ArrayLike As Variant)
    If IsObject(ArrayLike) Then
        If Index < 0 Or Index > pIndex - ArrayLike.Count Then Err.Raise 9
        VariantArrayClone VarPtr(pItems(Index)), VarPtr(ArrayLike.Items(0)), ArrayLike.Count
    Else
        If Index < 0 Or Index > pIndex - (1 + UBound(ArrayLike)) Then Err.Raise 9
        VariantArrayClone VarPtr(pItems(Index)), VarPtr(ArrayLike(0)), 1 + UBound(ArrayLike)
    End If
End Sub

Public Sub Sort_2(ByRef Comparer As IComparer): Sort 0, , Comparer: End Sub
Public Sub Sort(Optional ByVal Index As Long = 0, Optional ByVal GetCount As Variant, Optional ByRef Comparer As IComparer = Nothing)
    If IsMissing(GetCount) Then GetCount = pIndex - Index
    If Comparer Is Nothing Then
        QuickSortV2 Index, Index + GetCount - 1
    Else
        QuickSort Index, Index + GetCount - 1, Comparer
    End If
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

Public Function GetAddressOfItemAt(ByVal Index As Long) As LongPtr
    GetAddressOfItemAt = VarPtr(pItems(Index))
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

Private Sub ReserveSpaceForInsert(ByVal Index As Long, ByVal GetCount As Long)
    Dim t() As Variant, cap As Long
    cap = GetCapacityFor(GetCount)
    ReDim t(0 To cap)
    If Index > 0 Then MemCopy VarPtr(t(0)), VarPtr(pItems(0)), Index * VARIANT_SIZE
    MemCopy VarPtr(t(Index + GetCount)), VarPtr(pItems(Index)), (pIndex - Index) * VARIANT_SIZE
    ZeroMemory VarPtr(pItems(0)), pIndex * VARIANT_SIZE
    Erase pItems
    ReassignArrayTo pItems, t
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

Private Function C2ArrayList(ByRef Value As Variant) As ArrayList: Set C2ArrayList = Value: End Function
