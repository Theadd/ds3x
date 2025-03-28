﻿VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Array2dEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder "ds3x.Collections"
Option Compare Database
Option Explicit
Option Base 0

Implements ICollectionEx


Public Instance As Variant
Private pColumnCountMod As Long
Private pInstance() As Variant


Public Property Get Count() As Long
    On Error Resume Next
    Count = 1 + (UBound(pInstance, 1) - LBound(pInstance, 1))
    On Error GoTo 0
End Property

Public Property Get ColumnCount() As Long
    On Error Resume Next
    ColumnCount = 1 + (UBound(pInstance, 2) - LBound(pInstance, 2)) + pColumnCountMod
    On Error GoTo 0
End Property

Public Property Let ColumnCount(ByVal Value As Long): pColumnCountMod = IIf(IsEmpty(Instance), Value, pColumnCountMod): End Property

'@DefaultMember
Public Property Get Item(ByVal Index As Long, Optional ByVal ColumnIndex As Long = 0) As Variant
Attribute Item.VB_UserMemId = 0
    Item = pInstance(Index, ColumnIndex)
End Property

Public Property Get Row(ByVal Index As Long) As Variant
    Dim t() As Variant, i As Long
    ReDim t(0 To UBound(pInstance, 2))

    For i = 0 To UBound(pInstance, 2)
        If IsObject(pInstance(Index, i)) Then
            Set t(i) = pInstance(Index, i)
        Else
            t(i) = pInstance(Index, i)
        End If
    Next i
    Row = t
End Property


Public Sub CloneSliceTo(ByVal Index As Long, ByVal GetCount As Long, ByVal ColumnIndex As Long, ByVal TargetMemoryAddress As LongPtr)
    VariantArrayClone TargetMemoryAddress, VarPtr(pInstance(Index, ColumnIndex)), GetCount
End Sub

Public Sub SetInstanceByRef(ByRef TargetArray() As Variant)
    If IsEmpty(Instance) Then
        ReDim pInstance(0 To 0, 0 To 0)
        Instance = GetArrayByRef(pInstance)
        ReassignArrayTo Instance, TargetArray
    Else
        Err.Raise 13, , "Can't reassign an Array2dEx Instance to another 2D Array once initialized."
    End If
End Sub

Public Function CreateBlank(ByVal RowsCount As Long, ByVal ColumnsCount As Long) As Array2dEx: Set CreateBlank = CreateArray2dEx(RowsCount, ColumnsCount): End Function

Public Function Create(Optional ByRef ArrayLike As Variant) As Array2dEx
    If IsMissing(ArrayLike) Then
        Set Create = New Array2dEx
    ElseIf IsObject(ArrayLike) Then
        Set Create = CreateFromObject(C2Object(ArrayLike))
    ElseIf IsArray(ArrayLike) Then
        Select Case GetArrayDimsCount(ArrayLike)
            Case 2
                Set Create = New Array2dEx
                Dim t() As Variant
                t = ArrayLike
                Create.SetInstanceByRef t
            Case 1
                Set Create = CreateFromJaggedArray(ArrayLike)
            Case Else
                Err.Raise 13
        End Select
    Else
        Err.Raise 13
    End If
End Function

Private Function CreateFromObject(ByVal ArrayLike As Object) As Array2dEx
    If TypeOf ArrayLike Is ADODB.Recordset Then
        Set CreateFromObject = CreateFromRecordset(C2Recordset(ArrayLike))
    ElseIf TypeOf ArrayLike Is RecordsetEx Then
        Set CreateFromObject = CreateFromRecordset(ArrayLike.Instance)
    ElseIf TypeOf ArrayLike Is Excel.Range Then
        Set CreateFromObject = CreateFromExcelRange(C2ExcelRange(ArrayLike))
    ElseIf TypeOf ArrayLike Is ADODB.Fields Then
        Set CreateFromObject = CreateFromRecordsetFields(C2RecordsetFields(ArrayLike))
    ElseIf TypeOf ArrayLike Is ArrayListEx Then
        With C2ArrayListEx(ArrayLike)
            If .ColumnCount = 0 Then
                If .Count = 0 Then
                    Set CreateFromObject = CreateArray2dEx(0, 0)
                Else
                    Set CreateFromObject = CreateFromSingleColumnArrayListEx(C2ArrayListEx(ArrayLike))
                End If
            Else
                Set CreateFromObject = CreateFromArrayList(.Instance)
            End If
        End With
    ElseIf TypeOf ArrayLike Is ArrayList Then
        Set CreateFromObject = CreateFromArrayList(C2ArrayList(ArrayLike))
    Else
        Err.Raise 13
    End If
End Function

Public Function Bind(Optional ByRef ArrayLike As Variant) As Array2dEx
    Set Bind = Me
    Err.Raise 17, , "Can't perform requested operation. Binding to an immutable object is not allowed."
End Function

Public Function Unbind() As Array2dEx
    Set Unbind = Me
    Err.Raise 17, , "Can't perform requested operation. Unbinding an Array2dEx can't be done."
End Function

Public Function CopyToRange(ByRef RangeObject As Excel.Range, Optional ByVal ApplyUserLocale As Boolean = True) As Array2dEx
    Set CopyToRange = Me
    If Not Count > 0 Then Exit Function

    With RangeObject.Resize(Me.Count, 1 + (UBound(pInstance, 2) - LBound(pInstance, 2)))
        If ApplyUserLocale Then
            On Error GoTo UserLocaleFail
            .FormulaR1C1Local = pInstance
        Else
            .Value = pInstance
        End If
    End With

    Exit Function
UserLocaleFail:
    Debug.Print "[ERROR] @Array2dEx.CopyToRange() -> ApplyUserLocale - " & Err.Description
    Resume Next
End Function

Public Function CopySliceToRange(ByVal Index As Long, ByVal GetCount As Long, ByVal ColumnIndex As Long, ByRef RangeObject As Excel.Range, Optional ByVal ApplyUserLocale As Boolean = True) As Array2dEx
    Dim t() As Variant
    Set CopySliceToRange = Me
    If Not (Count > 0 And GetCount > 0) Then Exit Function
    ReDim t(0 To GetCount - 1, 0 To 0)
    VariantArrayClone VarPtr(t(0, 0)), VarPtr(pInstance(Index, ColumnIndex)), GetCount

    With RangeObject.Resize(GetCount, 1)
        If ApplyUserLocale Then
            On Error GoTo UserLocaleFail
            .FormulaR1C1Local = t
        Else
            .Value = t
        End If
    End With

    Exit Function
UserLocaleFail:
    Debug.Print "[ERROR] @Array2dEx.CopySliceToRange() -> ApplyUserLocale - " & Err.Description
    Resume Next
End Function

Private Function CreateFromRecordset(ByRef rs As ADODB.Recordset) As Array2dEx
    Dim r As Long, c As Long, t() As Variant, ubRows As Long, ubCols As Long, k() As Variant

    ubRows = rs.RecordCount - 1
    ubCols = rs.Fields.Count - 1
    If ubRows >= 0 Then
        ' TODO: Restore rs.Bookmark instead of calling rs.MoveFirst below
        t = rs.GetRows(adGetRowsRest, adBookmarkFirst)
        ReDim k(0 To ubRows, 0 To ubCols)

        rs.MoveFirst
        For c = 0 To ubCols
            If rs.Fields(c).Type = adVarWChar Then
                For r = 0 To ubRows
                    k(r, c) = C2Date(t(c, r))
                Next r
            Else
                For r = 0 To ubRows
                    k(r, c) = t(c, r)
                Next r
            End If
        Next c
    End If

    Set CreateFromRecordset = New Array2dEx
    CreateFromRecordset.SetInstanceByRef k
End Function

Private Function CreateFromExcelRange(ByVal TargetRange As Excel.Range) As Array2dEx
    Dim t() As Variant

    Set CreateFromExcelRange = New Array2dEx
    t = TargetRange.Value

    UpdateLBound t, 1, 0
    UpdateLBound t, 2, 0

    CreateFromExcelRange.SetInstanceByRef t
End Function

Private Function CreateFromRecordsetFields(ByVal AllFields As ADODB.Fields) As Array2dEx
    Dim i As Long, t() As Variant
    ReDim t(0 To 0, 0 To AllFields.Count - 1)

    For i = 0 To AllFields.Count - 1
        t(0, i) = CStr(Nz(AllFields(i).Name))
    Next i

    Set CreateFromRecordsetFields = New Array2dEx
    CreateFromRecordsetFields.SetInstanceByRef t
End Function

Private Function CreateFromArrayList(ByVal TargetArray As ArrayList) As Array2dEx
    Dim Entry As Variant, i As Long, c As Long, lc As Long, uc As Long, t() As Variant
    If TargetArray.Count = 0 Then Exit Function

    lc = LBound(TargetArray(0))
    uc = UBound(TargetArray(0))
    ReDim t(0 To TargetArray.Count - 1, lc To uc)

    For i = 0 To TargetArray.Count - 1
        Entry = TargetArray(i)
        For c = lc To uc
            If IsObject(Entry(c)) Then
                Set t(i, c) = Entry(c)
            Else
                t(i, c) = Entry(c)
            End If
        Next c
    Next i

    Set CreateFromArrayList = New Array2dEx
    CreateFromArrayList.SetInstanceByRef t
End Function

Private Function CreateFromSingleColumnArrayListEx(ByVal TargetArray As ArrayListEx) As Array2dEx
    Dim t() As Variant, i As Long, v As Variant
    If TargetArray.Count = 0 Then Exit Function
    ReDim t(0 To TargetArray.Count - 1, 0 To 0)

    For Each v In TargetArray.ToArray()
        If IsObject(v) Then
            Set t(i, 0) = v
        Else
            t(i, 0) = v
        End If
        i = i + 1
    Next v

    Set CreateFromSingleColumnArrayListEx = New Array2dEx
    CreateFromSingleColumnArrayListEx.SetInstanceByRef t
End Function

Private Function CreateFromJaggedArray(ByRef TargetArray As Variant) As Array2dEx
    Dim Entry As Variant, i As Long, c As Long, lc As Long, uc As Long, t() As Variant
    Set CreateFromJaggedArray = New Array2dEx
    If UBound(TargetArray) < 0 Then Exit Function

    For Each Entry In TargetArray
        If IsArray(Entry) Then
            If UBound(Entry) - LBound(Entry) > uc Then
                uc = UBound(Entry)
                lc = LBound(Entry)
            End If
        End If
        i = i + 1
        If i > 1000 Then Exit For
    Next Entry

    ReDim t(LBound(TargetArray) To UBound(TargetArray), lc To uc)
    i = LBound(TargetArray)

    For Each Entry In TargetArray
        If IsArray(Entry) Then
            For c = lc To UBound(Entry)
                t(i, c) = Entry(c)
            Next c
        Else
            t(i, lc) = Entry
        End If
        i = i + 1
    Next Entry

    CreateFromJaggedArray.SetInstanceByRef t
End Function

' Concatenate all elements of TargetArray as additional columns into a new Array2dEx
Public Function Join(ByVal TargetArray As Array2dEx) As Array2dEx
    Dim t() As Variant, i As Long, c As Long, c2 As Long, uc1 As Long, uc2 As Long
    uc1 = Me.ColumnCount - 1
    uc2 = TargetArray.ColumnCount - 1
    ReDim t(0 To Me.Count - 1, 0 To (uc1 + uc2 + 1))

    For i = 0 To Me.Count - 1
        For c = 0 To uc1
            If IsObject(pInstance(i, c)) Then
                Set t(i, c) = pInstance(i, c)
            Else
                t(i, c) = pInstance(i, c)
            End If
        Next c
        For c2 = 0 To uc2
            If IsObject(TargetArray(i, c2)) Then
                Set t(i, uc1 + 1 + c2) = TargetArray(i, c2)
            Else
                t(i, uc1 + 1 + c2) = TargetArray(i, c2)
            End If
        Next c2
    Next i

    Set Join = New Array2dEx
    Join.SetInstanceByRef t
End Function

Public Function AddRange(ByVal TargetArray As Array2dEx) As Array2dEx
    Set AddRange = New Array2dEx
    ' TODO: Implement
    Err.Raise 17, , "Not implemented."
End Function

' @SEEALSO: ArrayListEx.GetRange([Index], [Count], [ColumnIndexes]) As ArrayListEx
Public Function GetRange(Optional ByVal Index As Variant, Optional ByVal GetCount As Variant, Optional ByVal ColumnIndexes As Variant) As Array2dEx
    Dim t() As Variant, i As Long, c As Variant, c0 As Long
    Set GetRange = New Array2dEx

    If IsMissing(Index) Or IsEmpty(Index) Then Index = 0
    If Index >= Me.Count Then
        Index = IIf(Me.Count - 1 < 0, 0, Me.Count - 1)
        GetCount = 0
    Else
        If IsMissing(GetCount) Or IsEmpty(GetCount) Then GetCount = Me.Count - Index
        If GetCount + Index > Me.Count Then GetCount = Me.Count - Index
    End If
    If IsMissing(ColumnIndexes) Or IsEmpty(ColumnIndexes) Then
        ColumnIndexes = CreateNumericIndexHeaders
    Else
        If Not IsArray(ColumnIndexes) Then ColumnIndexes = Array(ColumnIndexes)
    End If

    If GetCount > 0 Then
        ReDim t(0 To GetCount - 1, 0 To UBound(ColumnIndexes) - LBound(ColumnIndexes))

        For i = Index To Index + GetCount - 1
            c0 = 0
            For Each c In ColumnIndexes
                If IsObject(pInstance(i, CLng(c))) Then
                    Set t(i - Index, c0) = pInstance(i, CLng(c))
                Else
                    t(i - Index, c0) = pInstance(i, CLng(c))
                End If
                c0 = c0 + 1
            Next c
        Next i
    Else
        t = CreateBlank2dArray()
        GetRange.ColumnCount = -1
    End If

    GetRange.SetInstanceByRef t
End Function

' Returns an Array2dEx containing only those rows specified in the RowIndexes array.
' Can also be used for: Reordering, cutting rows, duplicate, etc.
'
' @param RowIndexes - An array of row indices.
' @param [ModIndex] - A signed integer to shift values in the RowIndexes array.
Public Function GetRows(ByRef RowIndexes As Variant, Optional ByVal ModIndex As Long = 0) As Array2dEx
    Dim t() As Variant, aX As ArrayListEx, rIndex As Variant, nRows As Long, nCols As Long, i As Long, c As Long
    nRows = Me.Count
    nCols = Me.ColumnCount
    Set aX = ArrayListEx.Create()
    ReDim t(0 To nCols - 1)

    For Each rIndex In RowIndexes
        i = CLng(rIndex) + ModIndex
        If i >= 0 And i < nRows Then
            For c = 0 To nCols - 1
                t(c) = pInstance(i, c)
            Next c
            aX.Add t
        End If
    Next rIndex

    Set GetRows = Array2dEx.Create(aX)
End Function

Public Function Clone() As Array2dEx
    Dim t() As Variant
    t = pInstance
    Set Clone = Array2dEx.Create(t)
End Function

Public Function ToCSV(Optional ByVal Delimiter As String = ",", Optional ByVal InLocalFormat As Boolean = False) As String
    Dim t() As String, s() As String, i As Long, c As Long, uc As Long
    uc = ColumnCount - 1
    ReDim t(0 To Count)
    ReDim s(0 To uc)
    t(UBound(t)) = ""

    If InLocalFormat Then
        For i = 0 To UBound(t) - 1
            For c = 0 To uc
                s(c) = CollectionsLib.ValueToLocalString(pInstance(i, c))
            Next c
            t(i) = VBA.Join(s, Delimiter)
        Next i
        ToCSV = VBA.Join(t, vbNewLine)
    Else
        For i = 0 To UBound(t) - 1
            For c = 0 To uc
                s(c) = CollectionsLib.ValueToString(pInstance(i, c))
            Next c
            t(i) = VBA.Join(s, Delimiter)
        Next i
        ToCSV = VBA.Join(t, vbLf)
    End If
End Function

Public Function ToJSON() As String: ToJSON = JSON.Stringify(Me, 2): End Function

Public Function ToExcel() As String: ToExcel = ToCSV(vbTab, True): End Function

' --- OTHER ---

Private Sub Class_Initialize()
    Instance = Empty
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Instance = Empty
    Erase pInstance
    On Error GoTo 0
End Sub

Public Sub Dispose()
    ' TODO
End Sub

Private Function ValueToCSV(ByVal Value As Variant) As String
    Select Case VarType(Value)
        Case vbString
            ValueToCSV = """" & Replace(CStr(Value), """", """""") & """"
        Case vbNull
            ValueToCSV = vbNullString
        Case Else
            ValueToCSV = CStr(Value)
    End Select
End Function


' --- UTILITIES ---

'@Description("Transposes the values in a 2d array. Rows become columns, columns become rows.")
Public Function Transpose() As Array2dEx
    Dim cRow As Long, lbRow As Long, ubRow As Long, cCol As Long, lbCol As Long, ubCol As Long, t() As Variant

    lbCol = LBound(pInstance, 1)
    ubCol = UBound(pInstance, 1)
    lbRow = LBound(pInstance, 2)
    ubRow = UBound(pInstance, 2)
    ReDim t(lbRow To ubRow, lbCol To ubCol)

    For cRow = lbRow To ubRow
        For cCol = lbCol To ubCol
            If IsObject(pInstance(cCol, cRow)) Then
                Set t(cRow, cCol) = pInstance(cCol, cRow)
            Else
                t(cRow, cCol) = pInstance(cCol, cRow)
            End If
        Next
    Next

    Set Transpose = New Array2dEx
    Transpose.SetInstanceByRef t
End Function


' --- CASTING / TYPE CHECKING ---

Private Function C2Date(Optional ByVal Value As Variant) As Variant
    If IsMissing(Value) Then Exit Function
    On Error Resume Next
    C2Date = Value
    C2Date = DateValue(Value)
    On Error GoTo 0
End Function

Private Function C2Recordset(ByVal Value As Variant) As ADODB.Recordset: Set C2Recordset = Value: End Function
Private Function C2ExcelRange(ByVal Value As Variant) As Excel.Range: Set C2ExcelRange = Value: End Function
Private Function C2RecordsetFields(ByVal Value As Variant) As ADODB.Fields: Set C2RecordsetFields = Value: End Function
Private Function C2ArrayList(ByVal Value As Variant) As ArrayList: Set C2ArrayList = Value: End Function
Private Function C2ArrayListEx(ByVal Value As Variant) As ArrayListEx: Set C2ArrayListEx = Value: End Function
Private Function C2Object(ByRef Value As Variant) As Object: Set C2Object = Value: End Function

'@Deprecated
Public Function Is2dArray(ByRef ArrayLike As Variant) As Boolean
    On Error GoTo Finally
    Dim t As Long

    t = UBound(ArrayLike, 1)
    t = UBound(ArrayLike, 2)
    Is2dArray = True
    t = UBound(ArrayLike, 3)
    Is2dArray = False
Finally:
End Function

'@Deprecated
Public Function Is1dArray(ByRef ArrayLike As Variant) As Boolean
    On Error GoTo Fallback

    Is1dArray = Not (UBound(ArrayLike, 2) >= 0)
    Exit Function

Fallback:
    Is1dArray = IsArray(ArrayLike)
End Function

Private Function CreateBlank2dArray(Optional ByVal NumRows As Long = 0) As Variant
    Dim Miss As Variant, t() As Variant, i As Long
    With CreateObject("Forms.ComboBox.1")
        If NumRows = 0 Then
            .List = Array()
        Else
            Miss = Array(, Null)(0)
            ReDim t(0 To NumRows - 1)
            For i = 0 To NumRows - 1
                t(i) = Miss
            Next i
            .List = t
        End If
        CreateBlank2dArray = .List
    End With
End Function

Private Sub Assign(ByRef Target, Optional ByRef Value As Variant)
    If Not IsMissing(Value) Then
        If IsObject(Value) Then
            Set Target = Value
        Else
            Target = Value
        End If
    End If
End Sub


' --- OTHER ---

Private Function CreateNumericIndexHeaders() As Variant
    Dim t As Variant, c As Long, cMax As Long
    cMax = Me.ColumnCount - 1
    ReDim t(0 To cMax)

    For c = 0 To cMax
        t(c) = c
    Next c

    CreateNumericIndexHeaders = t
End Function

Private Function CreateArray2dEx(ByVal RowsCount As Long, ByVal ColumnsCount As Long) As Array2dEx
    Dim t() As Variant

    Set CreateArray2dEx = New Array2dEx
    If RowsCount > 0 And ColumnsCount > 0 Then
        ReDim t(0 To RowsCount - 1, 0 To ColumnsCount - 1)
    Else
        t = CreateBlank2dArray(RowsCount)
        CreateArray2dEx.ColumnCount = -1
    End If

    CreateArray2dEx.SetInstanceByRef t
End Function


' --- ICollectionEx ---

Private Property Get ICollectionEx_Count() As Long: ICollectionEx_Count = Count: End Property
Private Property Get ICollectionEx_ColumnCount() As Long: ICollectionEx_ColumnCount = ColumnCount: End Property
Private Property Get ICollectionEx_Item(ByVal Index As Long, Optional ByVal ColumnIndex As Long = 0&) As Variant: ICollectionEx_Item = Item(Index, ColumnIndex): End Property
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
Private Function ICollectionEx_CopyToRange(ByRef RangeObject As Excel.Range, Optional ByVal ApplyUserLocale As Boolean = True, Optional ByVal WriteHeaders As Boolean = True) As ICollectionEx: Set ICollectionEx_CopyToRange = CopyToRange(RangeObject, ApplyUserLocale): End Function
