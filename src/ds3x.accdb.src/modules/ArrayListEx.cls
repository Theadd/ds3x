﻿VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayListEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder "ds3x.Collections"
Option Compare Database
Option Explicit
Option Base 0


Implements ICollectionEx

Public Instance As ArrayList

Private Const dsChunkSize As Long = 1000


' Gets the count of elements in the underlying ArrayList.
Public Property Get Count() As Variant
Attribute Count.VB_Description = "Gets the count of elements in the underlying mscorlib.ArrayList."
    Count = Instance.Count
End Property

' In table-like instances, gets the number of columns (the count of elements within the first element, which is expected to be an array).
Public Property Get ColumnCount() As Long
Attribute ColumnCount.VB_Description = "In table-like instances, gets the number of columns (the count of elements within the first element, which is expected to be an array)."
    On Error Resume Next
    ColumnCount = CountElementsIn(Instance(0))
    On Error GoTo 0
End Property

' Sets or gets the element at the specified index. It's the default class member. When used as getter, a negative Index is applied to .Count, being -1 the last element.
'@DefaultMember
Public Property Get Item(ByVal Index As Variant) As Variant
Attribute Item.VB_Description = "Sets or gets the element at the specified index. It's the default class member. When used as getter, a negative Index is applied to .Count, being -1 the last element."
Attribute Item.VB_UserMemId = 0
    Index = CLng(Index)
    Select Case Index
        Case Is < 0
            Index = Instance.Count + Index
            If Index < 0 Then Exit Property
            Assign Item, Instance(Index)
        Case Is < Instance.Count
            Assign Item, Instance(Index)
    End Select
End Property

Public Property Let Item(ByVal Index As Variant, Value As Variant)
    If Index >= Instance.Count Then Allocate Index
    Instance(CLng(Index)) = Value
End Property

Public Property Set Item(ByVal Index As Variant, Value As Variant)
    If Index >= Instance.Count Then Allocate Index
    Set Instance(CLng(Index)) = Value
End Property

Public Property Get Items() As Variant()
    Items = Instance.Items()
End Property

Public Property Get Row(ByVal Index As Long) As Variant
    Row = Instance(Index)
End Property


Public Function CreateBlank(ByVal RowsCount As Long, ByVal ColumnsCount As Long) As ArrayListEx: Set CreateBlank = CreateArrayListEx(RowsCount, ColumnsCount): End Function

' Returns an ArrayListEx object bound to a new ArrayList instance containing the elements
' found in the optional ArrayLike parameter, if specified.
' If specified, ArrayLike can be from one of the following types:
'   - ArrayListEx, ArrayList, ICollection or any other class implementing the ICollection interface.
'   - A single dimension Array (1D Array).
'   - A 2D Array.
'   - A Collection
'   - A Scripting.Dictionary or DictionaryEx.
'   - Any other class instance with support for enumeration (VB_UserMemId = -4) and default value (VB_UserMemId = 0).
'   - A Range, being converted as an ArrayList (rows) of Variant arrays (columns).
Public Function Create(Optional ByRef ArrayLike As Variant) As ArrayListEx
Attribute Create.VB_Description = "Returns an ArrayListEx object bound to a new ArrayList instance containing the elements found in the optional ArrayLike parameter, if specified."
    Set Create = New ArrayListEx
    If Not IsMissing(ArrayLike) Then Create.AddRange ArrayLike
End Function

Friend Function Bind(Optional ByRef TargetArrayList As Variant) As ArrayListEx
    Set Bind = Me
    Set Instance = Nothing
    If IsMissing(TargetArrayList) Then Exit Function
    If TargetArrayList Is Nothing Then Exit Function
    If TypeOf TargetArrayList Is ArrayListEx Then
        Stop ' An ArrayListEx instance should NEVER bind to another ArrayListEx instance! @SEE: .Dispose()
        Set Instance = TargetArrayList.Instance
    ElseIf TypeOf TargetArrayList Is ArrayList Then
        Set Instance = TargetArrayList
    End If
End Function

Public Function Unbind() As ArrayListEx
    Set Instance = Nothing: Set Unbind = Me
End Function

'@Enumerator
Public Function NewEnum() As IEnumVARIANT
Attribute NewEnum.VB_UserMemId = -4
'    Dim enumerator As IUnknown
'    Set enumerator = Instance.GetEnumerator(0, Instance.Count)
'    Set NewEnum = enumerator
    Set NewEnum = Instance.GetEnumerator(0, Instance.Count)
End Function

' Adds an element to the ArrayList (@see also: ArrayListEx.BinaryAdd).
Public Function Add(Value As Variant) As ArrayListEx
Attribute Add.VB_Description = "Adds an item to the ArrayList instance."
    Instance.Add Value
    Set Add = Me
End Function

Public Function AddRange(Optional ByRef ArrayLike As Variant) As ArrayListEx
    Dim Item As Variant
    Set AddRange = Me
    If IsMissing(ArrayLike) Then Exit Function

    If IsObject(ArrayLike) Then
        If ArrayLike Is Nothing Then Exit Function

        If TypeOf ArrayLike Is ArrayListEx Then
            Instance.AddRange ArrayLike.Instance
        ElseIf TypeOf ArrayLike Is ArrayList Then
            Instance.AddRange C2ArrayList(ArrayLike)
        ElseIf TypeOf ArrayLike Is Excel.Range Then
            AddFromExcelRange C2ExcelRange(ArrayLike)
        ElseIf TypeOf ArrayLike Is Scripting.Dictionary Then
            AddRange DictionaryEx.Create(C2Dictionary(ArrayLike))
        ElseIf TypeOf C2Object(ArrayLike) Is Array2dEx Then
            AddFrom2dArray ArrayLike.Instance
        Else
            For Each Item In ArrayLike
                Instance.Add Item
            Next Item
        End If
    ElseIf IsArray(ArrayLike) Then
        If Is2dArray(ArrayLike) Then
            AddFrom2dArray ArrayLike
        Else
            For Each Item In ArrayLike
                Instance.Add Item
            Next Item
        End If
    End If
End Function

Public Function Insert(ByVal Index As Long, ByVal Value As Variant) As ArrayListEx
    Instance.Insert Index, Value
    Set Insert = Me
End Function

Public Function ToArray() As Variant()
    ToArray = Instance.ToArray
End Function

Static Function IsArrayLike(Optional ByRef ArrayLike As Variant) As Boolean
    If IsMissing(ArrayLike) Then Exit Function
    If IsObject(ArrayLike) Then
        If ArrayLike Is Nothing Then Exit Function
    End If

    IsArrayLike = (IsObject(ArrayLike) Or IsArray(ArrayLike))
End Function

Static Function CountElementsIn(ByRef ArrayLike As Variant) As Long
    On Error Resume Next
    If IsObject(ArrayLike) Then CountElementsIn = ArrayLike.Count
    If IsArray(ArrayLike) Then CountElementsIn = UBound(ArrayLike) + 1
    On Error GoTo 0
End Function

' Removes an element by value from the ArrayList.
Public Function Remove(Value As Variant) As ArrayListEx
Attribute Remove.VB_Description = "Removes an element by value from the ArrayList."
    Instance.Remove Value
    Set Remove = Me
End Function

' Removes a range of elements from the ArrayList.
Public Function RemoveRange(ByVal Index As Long, ByVal RemoveCount As Long) As ArrayListEx
Attribute RemoveRange.VB_Description = "Removes a range of elements from the ArrayList."
    Instance.RemoveRange Index, RemoveCount
    Set RemoveRange = Me
End Function

' Removes each element found in ArrayLike from the ArrayList.
Public Function RemoveFrom(Optional ByVal ArrayLike As Variant) As ArrayListEx
    Dim Item As Variant
    Set RemoveFrom = Me

    If IsMissing(ArrayLike) Then Exit Function
    If IsObject(ArrayLike) Then
        If ArrayLike Is Nothing Then Exit Function
    End If

    If IsObject(ArrayLike) Or IsArray(ArrayLike) Then
        If TypeOf ArrayLike Is ArrayListEx Then
            For Each Item In ArrayLike.Items()
                Instance.Remove Item
            Next Item
        Else
            For Each Item In ArrayLike
                Instance.Remove Item
            Next Item
        End If
    End If
End Function

' Removes all elements from the ArrayList.
Public Function RemoveAll() As ArrayListEx
Attribute RemoveAll.VB_Description = "Removes all elements from the ArrayList. @see ArrayListEx.Clear()"
    Instance.Clear
    Set RemoveAll = Me
End Function

' Removes all duplicated elements from the ArrayList.
Public Function Unique() As ArrayListEx
Attribute Unique.VB_Description = "Removes all duplicated elements from the ArrayList."
    Dim oArray As ArrayListEx, Item As Variant

    Set Unique = Me
    Set oArray = ArrayListEx.Create(Instance)
    Instance.Clear

    For Each Item In oArray.Items()
        If Not Instance.Contains(Item) Then
            Instance.Add Item
        End If
    Next Item
End Function

' Removes the element at Index from the ArrayList.
Public Function RemoveAt(ByVal Index As Long) As ArrayListEx
Attribute RemoveAt.VB_Description = "Removes the element at Index from the ArrayList."
    Instance.RemoveAt Index
    Set RemoveAt = Me
End Function

' Copies all elements from the ArrayList to the specified Range object. The Array2dEx.CopyToRange implementation is vastly more efficient than this one.
Public Function CopyToRange(ByRef RangeObject As Excel.Range) As ArrayListEx
Attribute CopyToRange.VB_Description = "Copies all elements from the ArrayList to the specified Range object. The Array2dEx.CopyToRange implementation is vastly more efficient than this one."
    Dim Entry As Variant, i As Long, c As Long

    For i = 0 To Instance.Count - 1
        Assign Entry, Instance(i)

        If IsObject(Entry) Or IsArray(Entry) Then
            For c = 0 To CountElementsIn(Entry) - 1
                RangeObject.Cells(i + 1, c + 1).Value = Entry(c)
            Next c
        Else
            RangeObject.Cells(i + 1, 1).Value = Entry
        End If
    Next i

    Set CopyToRange = Me
End Function

' Returns an ArrayListEx which represents a subset of the elements in this instance.
' Works for both lists (1D) and tables (2D). If used with tables (2D), the elements
' represent the rows and together with the ColumnIndexes parameter, allows to obtain
' the values of a specific subset of columns.
'
' @param [Index] - The zero-based ArrayList row index at which the range starts.
' @param [GetCount] - The number of rows or elements in the range.
' @param [ColumnIndexes] - If specified, an array of column indexes to include in the range, otherwise, no columns are excluded.
Public Function GetRange(Optional ByVal Index As Variant, Optional ByVal GetCount As Variant, Optional ByVal ColumnIndexes As Variant) As ArrayListEx
Attribute GetRange.VB_Description = "Returns an ArrayListEx which represents a subset of the elements in this instance."
    Dim Aux As ArrayList, xRange As ArrayListEx, Entry As Variant, i As Long

    If IsMissing(Index) Or IsEmpty(Index) Then Index = 0
    If Index >= Instance.Count Then
        Index = IIf(Instance.Count - 1 < 0, 0, Instance.Count - 1)
        GetCount = 0
    Else
        If IsMissing(GetCount) Or IsEmpty(GetCount) Then GetCount = Instance.Count - Index
        If GetCount + Index > Instance.Count Then GetCount = Instance.Count - Index
    End If

    If IsMissing(ColumnIndexes) Or IsEmpty(ColumnIndexes) Then
        Set GetRange = ArrayListEx.Create(Instance.GetRange(Index, GetCount))
        'GetRange.Bind Instance.GetRange(Index, GetCount)
    Else
        If Not IsArray(ColumnIndexes) Then
            Err.Raise 5
        Else
            Set GetRange = New ArrayListEx
            Set Aux = New ArrayList
'            Set xRange = New ArrayListEx
'            xRange.Bind Instance.GetRange(Index, GetCount)
            Set xRange = ArrayListEx.Create(Instance.GetRange(Index, GetCount))
            For Each Entry In xRange.Instance
                Aux.Clear
                For i = 0 To UBound(ColumnIndexes)
                    Aux.Add Entry(ColumnIndexes(i))
                Next i
                GetRange.Add Aux.ToArray()
            Next Entry
        End If
    End If
End Function

' Returns an ArrayListEx containing only those rows specified in the RowIndexes array.
' Can also be used for: Reordering, cutting rows, duplicate, etc.
'
' @param RowIndexes - An array of row indices.
' @param [ModIndex] - A signed integer to shift values in the RowIndexes array.
Public Function GetRows(ByRef RowIndexes As Variant, Optional ByVal ModIndex As Long = 0) As ArrayListEx
Attribute GetRows.VB_Description = "Returns an ArrayListEx containing only those rows specified in the RowIndexes array."
    Dim aX As ArrayListEx, rIndex As Variant, nRows As Long, i As Long
    nRows = Me.Count
    Set aX = ArrayListEx.Create()

    For Each rIndex In RowIndexes
        i = CLng(rIndex) + ModIndex
        If i >= 0 And i < nRows Then
            aX.Add Instance(i)
        End If
    Next rIndex

    Set GetRows = aX
End Function

' Concatenate all elements of TargetArray as additional columns into a new ArrayListEx
Public Function Join(ByVal TargetArray As ArrayListEx) As ArrayListEx
Attribute Join.VB_Description = "Concatenate all elements of TargetArray as additional columns into a new ArrayListEx"
    Dim i As Long, c As Long, cMax As Long, cMax2 As Long, aX As ArrayListEx, Aux As ArrayListEx

    Set aX = ArrayListEx.Create()
    Set Aux = ArrayListEx.Create()
    cMax = Me.ColumnCount - 1
    cMax2 = TargetArray.ColumnCount - 1
    For i = 0 To Me.Count - 1
        Aux.Clear
        For c = 0 To cMax
            Aux.Add Me(i)(c)
        Next c
        For c = 0 To cMax2
            Aux.Add TargetArray(i)(c)
        Next c
        aX.Add Aux.ToArray()
    Next i

    Set Join = aX
End Function

' Removes all elements from the ArrayList.
Public Function Clear() As ArrayListEx
Attribute Clear.VB_Description = "Removes all elements from the ArrayList."
    Instance.Clear
    Set Clear = Me
End Function

' Whether a Value is contained in this ArrayListEx or not, using a STRICT equality comparer, which compares by value and value type. (Double !== Long)
Public Function Contains(ByVal Value As Variant) As Boolean
Attribute Contains.VB_Description = "Whether a Value is contained in this ArrayListEx or not, using a STRICT equality comparer, which compares by value and value type."
    Contains = Instance.Contains(Value)
End Function

Public Function IndexOf(ByVal Value As Variant, Optional ByVal startIndex As Variant) As Long
    If IsMissing(startIndex) Then startIndex = 0
    IndexOf = Instance.IndexOf(Value, startIndex)
End Function

Public Function Reverse() As ArrayListEx
    Instance.Reverse
    Set Reverse = Me
End Function

Public Function Clone() As ArrayListEx
    Set Clone = ArrayListEx.Create(Me)
End Function


' --- IComparer ---

Public Function Sort(Optional ByVal Comparer As IComparer = Nothing) As ArrayListEx
    If Comparer Is Nothing Then
        Instance.Sort
    Else
        Instance.Sort_2 Comparer
    End If
    Set Sort = Me
End Function

Public Function BinarySearch(ByVal Value As Variant, Optional ByVal startIndex As Long = 0, Optional ByVal SearchCount As Variant, Optional ByVal Comparer As IComparer = Nothing) As Long
    If IsMissing(SearchCount) Then SearchCount = Instance.Count - startIndex
    If Comparer Is Nothing Then
        BinarySearch = Instance.BinarySearch(startIndex, SearchCount, Value, Comparer)
    Else
        BinarySearch = Instance.BinarySearch(startIndex, SearchCount, Value, Comparer)
    End If
End Function

Public Function BinaryAdd(ByVal Value As Variant, Optional ByVal startIndex As Long = 0, Optional ByVal Comparer As IComparer = Nothing) As Long
    Dim r As Long

    If Comparer Is Nothing Then
        r = Instance.BinarySearch(startIndex, Instance.Count - startIndex, Value, Comparer)
    Else
        r = Instance.BinarySearch(startIndex, Instance.Count - startIndex, Value, Comparer)
    End If
    If r < 0 Then
        r = (Not r)
        If r = Instance.Count Then
            Add Value
        Else
            Insert r, Value
        End If
    End If

    BinaryAdd = r
End Function

Public Function BinaryToggle(ByVal Value As Variant, Optional ByVal startIndex As Long = 0, Optional ByVal Comparer As IComparer = Nothing) As Long
    Dim r As Long

    If Comparer Is Nothing Then
        r = Instance.BinarySearch(startIndex, Instance.Count - startIndex, Value, Comparer)
    Else
        r = Instance.BinarySearch(startIndex, Instance.Count - startIndex, Value, Comparer)
    End If
    If r < 0 Then
        r = (Not r)
        If r = Instance.Count Then
            Add Value
        Else
            Insert r, Value
        End If
    Else
        RemoveAt r
    End If

    BinaryToggle = r
End Function


' --- To String ---

Public Function PrintAll() As ArrayListEx
    Debug.Print JSON.Stringify(Instance, 2)
    Set PrintAll = Me
End Function

Public Function ToCSV(Optional ByVal Delimiter As String = ",", Optional ByVal InLocalFormat As Boolean = False) As String
    Dim t() As String, s() As String, i As Long, c As Variant, uc As Long, r As Variant
    uc = ColumnCount - 1
    ReDim t(0 To Count)
    t(UBound(t)) = ""
    If uc >= 0 Then
        ReDim s(0 To uc)
        If InLocalFormat Then
            For i = 0 To UBound(t) - 1
                For c = 0 To uc
                    s(c) = CollectionsLib.ValueToLocalString(Instance(i)(c))
                Next c
                t(i) = VBA.Join(s, Delimiter)
            Next i
            ToCSV = VBA.Join(t, vbNewLine)
        Else
            For i = 0 To UBound(t) - 1
                For c = 0 To uc
                    s(c) = CollectionsLib.ValueToString(Instance(i)(c))
                Next c
                t(i) = VBA.Join(s, Delimiter)
            Next i
            ToCSV = VBA.Join(t, vbLf)
        End If
    Else
        If InLocalFormat Then
            For i = 0 To UBound(t) - 1
                t(i) = CollectionsLib.ValueToLocalString(Instance(i))
            Next i
            ToCSV = VBA.Join(t, vbNewLine)
        Else
            For i = 0 To UBound(t) - 1
                t(i) = CollectionsLib.ValueToString(Instance(i))
            Next i
            ToCSV = VBA.Join(t, vbLf)
        End If
    End If
End Function

Public Function ToJSON() As String: ToJSON = JSON.Stringify(Me, 2): End Function

Public Function ToExcel() As String: ToExcel = ToCSV(vbTab, True): End Function


' --- PRIVATE METHODS ---

Private Sub Allocate(ByVal Index As Variant)
    Dim i As Long

    For i = Instance.Count To Index
        Instance.Add Empty
    Next i
End Sub

Private Sub Assign(ByRef Target, Optional ByRef Value As Variant)
    If Not IsMissing(Value) Then
        If IsObject(Value) Then
            Set Target = Value
        Else
            Target = Value
        End If
    End If
End Sub

Private Function CreateArrayListEx(ByVal RowsCount As Long, ByVal ColumnsCount As Long) As ArrayListEx
    Dim aX As ArrayListEx, i As Long
    Set aX = ArrayListEx.Create()
    Set CreateArrayListEx = ArrayListEx.Create()

    If ColumnsCount > 0 Then aX(ColumnsCount - 1) = Empty
    For i = 0 To RowsCount - 1
        CreateArrayListEx.Add aX.ToArray()
    Next i
End Function

Private Sub AddFromExcelRange(ByRef r As Excel.Range)
    Dim vItems As Variant, i As Long, c As Long, lc As Long, uc As Long, ui As Long, vRow As Variant
    Dim g As Long, gNumRows As Long, rCount As Long, rng2 As Excel.Range, rBase As Long

    rCount = r.Rows.Count
    gNumRows = rCount - (g * dsChunkSize)
    If gNumRows > dsChunkSize Then gNumRows = dsChunkSize
    If Not (gNumRows > 0) Then Exit Sub
    rBase = GetBaseRowIndexOfRange(r)

    Do While (gNumRows > 0)
        Set rng2 = r.Range(r.Cells(rBase + (g * dsChunkSize), 1), r.Cells((rBase + (g * dsChunkSize)) + (gNumRows - 1), r.Columns.Count))
        vItems = Empty
        vItems = rng2.Value
        lc = LBound(vItems, 2)
        uc = UBound(vItems, 2)
        ui = UBound(vItems, 1)
        ReDim vRow(0 To (uc - lc))

        For i = LBound(vItems, 1) To ui
            For c = lc To uc - 1
                vRow(c - lc) = vItems(i, c)
            Next c
            If VarType(vItems(i, uc)) = vbString Then
                vRow(uc - lc) = Trim(vItems(i, uc))
            Else
                vRow(uc - lc) = vItems(i, uc)
            End If
            Instance.Add vRow
        Next i

        g = g + 1
        gNumRows = rCount - (g * dsChunkSize)
        If gNumRows > dsChunkSize Then gNumRows = dsChunkSize
    Loop
    vItems = Empty
End Sub

Private Function GetBaseRowIndexOfRange(ByRef r As Excel.Range) As Long
    On Error GoTo Fallback

    GetBaseRowIndexOfRange = IIf(r.Range(r.Cells(0, 1), r.Cells(1, 2)).Rows.Count > 0, 0, 1)
    Exit Function
Fallback:
    GetBaseRowIndexOfRange = 1
End Function

Private Function AddFrom2dArray(ByRef vItems As Variant) As ArrayListEx
    Dim i As Long, c As Long, lc As Long, uc As Long, ui As Long, vRow As Variant

    lc = LBound(vItems, 2)
    uc = UBound(vItems, 2)
    ui = UBound(vItems, 1)
    ReDim vRow(0 To (uc - lc))

    For i = LBound(vItems, 1) To ui
        For c = lc To uc
            If IsObject(vItems(i, c)) Then
                Set vRow(c - lc) = vItems(i, c)
            Else
                vRow(c - lc) = vItems(i, c)
            End If
        Next c
        Instance.Add vRow
    Next i

    Set AddFrom2dArray = Me
End Function


' --- TYPE CASTING ---

Private Function C2Object(ByRef Value As Variant) As Object: Set C2Object = Value: End Function
Private Function C2ArrayList(ByRef Value As Variant) As ArrayList: Set C2ArrayList = Value: End Function
Private Function C2ExcelRange(ByRef Value As Variant) As Excel.Range: Set C2ExcelRange = Value: End Function
Private Function C2Dictionary(ByRef Value As Variant) As Scripting.Dictionary: Set C2Dictionary = Value: End Function


Private Sub Class_Initialize()
    Set Instance = New ArrayList
End Sub

Private Sub Class_Terminate()
    Set Instance = Nothing
End Sub

Public Sub Dispose()
    ' REMOVED
End Sub

Private Function Is2dArray(ByRef ArrayLike As Variant) As Boolean
    On Error GoTo Finally
    Dim t As Long

    t = UBound(ArrayLike, 1)
    t = UBound(ArrayLike, 2)
    Is2dArray = True
    t = UBound(ArrayLike, 3)
    Is2dArray = False

Finally:
End Function


' --- ICollectionEx ---

Private Property Get ICollectionEx_Count() As Long: ICollectionEx_Count = Count: End Property
Private Property Get ICollectionEx_ColumnCount() As Long: ICollectionEx_ColumnCount = ColumnCount: End Property
Private Property Get ICollectionEx_Item(ByVal Index As Long, Optional ByVal ColumnIndex As Long = 0&) As Variant: ICollectionEx_Item = Item(Index)(ColumnIndex): End Property
Private Property Get ICollectionEx_Row(ByVal Index As Long) As Variant: ICollectionEx_Row = Row(Index): End Property
Private Function ICollectionEx_CreateBlank(ByVal RowsCount As Long, ByVal ColumnsCount As Long) As ICollectionEx: Set ICollectionEx_CreateBlank = CreateBlank(RowsCount, ColumnsCount): End Function
Private Function ICollectionEx_Create(Optional ByRef FromTarget As Variant) As ICollectionEx: Set ICollectionEx_Create = Create(FromTarget): End Function
Private Function ICollectionEx_Bind(Optional ByRef Target As Variant) As ICollectionEx: Set ICollectionEx_Bind = Bind(Target): End Function
Private Function ICollectionEx_Unbind() As ICollectionEx: Set ICollectionEx_Unbind = Unbind(): End Function
Private Sub ICollectionEx_Dispose(): Dispose: End Sub
Private Function ICollectionEx_Join(ByVal Target As ICollectionEx) As ICollectionEx: Set ICollectionEx_Join = Join(Target): End Function
Private Function ICollectionEx_AddRange(ByVal Target As ICollectionEx) As ICollectionEx: Set ICollectionEx_AddRange = AddRange(Target): End Function
Private Function ICollectionEx_GetRange(Optional ByVal Index As Variant, Optional ByVal GetCount As Variant, Optional ByVal ColumnIndexes As Variant) As ICollectionEx: Set ICollectionEx_GetRange = GetRange(Index, GetCount, ColumnIndexes): End Function
Private Function ICollectionEx_ToCSV(Optional ByVal Delimiter As String = ",", Optional ByVal InLocalFormat As Boolean = False) As String: ICollectionEx_ToCSV = ToCSV(Delimiter, InLocalFormat): End Function
Private Function ICollectionEx_ToJSON() As String: ICollectionEx_ToJSON = ToJSON(): End Function
Private Function ICollectionEx_ToExcel() As String: ICollectionEx_ToExcel = ToExcel(): End Function
Private Function ICollectionEx_CopyToRange(ByRef RangeObject As Excel.Range, Optional ByVal ApplyUserLocale As Boolean = True, Optional ByVal WriteHeaders As Boolean = True) As ICollectionEx: Set ICollectionEx_CopyToRange = CopyToRange(RangeObject): End Function
