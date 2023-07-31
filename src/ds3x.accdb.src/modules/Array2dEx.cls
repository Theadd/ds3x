VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Array2dEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Option Base 0

' @SEE: https://github.com/cristianbuse/VBA-MemoryTools/blob/master/src/LibMemory.bas#L574


Public Instance As Variant


'Public Property Get ResolvedInstance() As Variant
'    Set ResolvedInstance = CollectionsLib.GetArrayByRef(Instance)
'End Property

Public Property Get Count() As Long
    On Error Resume Next
    Count = 1 + (UBound(Instance, 1) - LBound(Instance, 1))
End Property

Public Property Get ColumnCount() As Long
    On Error Resume Next
    ColumnCount = 1 + (UBound(Instance, 2) - LBound(Instance, 2))
End Property



Public Function Create(Optional ByRef ArrayLike As Variant) As Array2dEx
    If IsMissing(ArrayLike) Then
        Set Create = New Array2dEx
    ElseIf IsObject(ArrayLike) Then
        Set Create = CreateFromObject(C2Object(ArrayLike))
    ElseIf IsArray(ArrayLike) Then
        If Is2dArray(ArrayLike) Then
            Set Create = New Array2dEx
            Create.Instance = ArrayLike
        ElseIf Is1dArray(ArrayLike) Then
            Set Create = CreateFromJaggedArray(ArrayLike)
        Else
            Err.Raise 13
        End If
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
        Set CreateFromObject = CreateFromArrayList(ArrayLike.Instance)
    ElseIf TypeOf ArrayLike Is ArrayList Then
        Set CreateFromObject = CreateFromArrayList(C2ArrayList(ArrayLike))
    Else
        If DEBUG_MODE_ENABLED Then Stop
        Err.Raise 13
    End If
End Function

' WRONG DESCRIPTION: Points this Array2dEx's .Instance (underlying variant array) to reference the provided ArrayLike
' Currently, copies by value the 2d array found within ArrayLike to this .Instance
Public Function Bind(Optional ByRef ArrayLike As Variant) As Array2dEx
    Set Bind = Me
    ' TODO: FIXME: This overwrites the array not a pointer to an array.
    Instance = Empty
    
    If Not IsMissing(ArrayLike) Then
        If IsArray(ArrayLike) Then
            If Is2dArray(ArrayLike) Then
                Err.Raise 17
                'Instance = CollectionsLib.GetArrayByRef(ArrayLike)
            ElseIf Is1dArray(ArrayLike) Then
                Err.Raise 13
            End If
        ElseIf IsObject(ArrayLike) Then
            If TypeOf ArrayLike Is Array2dEx Then
                Err.Raise 17
                'Instance = CollectionsLib.GetArrayByRef(ArrayLike.Instance)
            End If
        End If
    End If
End Function



Public Function CopyToRange(ByRef RangeObject As Range, Optional ByVal ApplyUserLocale As Boolean = True) As Array2dEx
    Set CopyToRange = Me
    If Not Me.Count > 0 Then Exit Function

    With RangeObject.Resize(Me.Count, 1 + (UBound(Instance, 2) - LBound(Instance, 2)))
        .Value = Instance
        If ApplyUserLocale Then
            On Error GoTo UserLocaleFail
            .FormulaR1C1Local = .FormulaR1C1Local
        End If
    End With
    
    Exit Function
UserLocaleFail:
    Debug.Print "[ERROR] @Array2dEx.CopyToRange() -> ApplyUserLocale - " & Err.Description
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
        
        'On Error Resume Next
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
    CreateFromRecordset.Instance = CollectionsLib.GetArrayByRef(k)
End Function

Private Function CreateFromExcelRange(ByVal TargetRange As Excel.Range) As Array2dEx
    Set CreateFromExcelRange = New Array2dEx
    CreateFromExcelRange.Instance = TargetRange.Value
End Function

Private Function CreateFromRecordsetFields(ByVal AllFields As ADODB.Fields) As Array2dEx
    Dim i As Long, t() As Variant
    ReDim t(0 To 0, 0 To AllFields.Count - 1)
    
    For i = 0 To AllFields.Count - 1
        t(0, i) = CStr(Nz(AllFields(i).Name))
    Next i
    
    Set CreateFromRecordsetFields = New Array2dEx
    CreateFromRecordsetFields.Instance = CollectionsLib.GetArrayByRef(t)
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
    CreateFromArrayList.Instance = CollectionsLib.GetArrayByRef(t)
End Function


Private Function CreateFromJaggedArray(ByVal TargetArray As Variant) As Array2dEx
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
    
    CreateFromJaggedArray.Instance = CollectionsLib.GetArrayByRef(t)
End Function

' Concatenate all elements of TargetArray as additional columns into a new Array2dEx
Public Function Join(ByVal TargetArray As Array2dEx) As Array2dEx
    Dim t() As Variant, i As Long, c As Long, c2 As Long, uc1 As Long, uc2 As Long
    
    uc1 = Me.ColumnCount - 1
    uc2 = TargetArray.ColumnCount - 1
    ' ReDim t(0 To Me.Count - 1, 0 To (Me.ColumnCount - 1) + TargetArray.ColumnCount)
    ReDim t(0 To Me.Count - 1, 0 To (uc1 + uc2 + 1))
    
    For i = 0 To Me.Count - 1
        For c = 0 To uc1
            If IsObject(Instance(i, c)) Then
                Set t(i, c) = Instance(i, c)
            Else
                t(i, c) = Instance(i, c)
            End If
        Next c
        For c2 = 0 To uc2
            If IsObject(TargetArray.Instance(i, c2)) Then
                Set t(i, uc1 + 1 + c2) = TargetArray.Instance(i, c2)
            Else
                t(i, uc1 + 1 + c2) = TargetArray.Instance(i, c2)
            End If
        Next c2
    Next i
    
    Set Join = New Array2dEx
    Join.Instance = CollectionsLib.GetArrayByRef(t)
End Function

' @SEEALSO: ArrayListEx.GetRange([Index], [Count], [ColumnIndexes]) As ArrayListEx
Public Function GetRange(Optional ByVal Index As Variant, Optional ByVal GetCount As Variant, Optional ByVal ColumnIndexes As Variant) As Array2dEx
    Dim t() As Variant, i As Long, c As Variant, c0 As Long
    
    If IsMissing(Index) Or IsEmpty(Index) Then Index = 0
    If IsMissing(GetCount) Or IsEmpty(GetCount) Then GetCount = Me.Count
    If IsMissing(ColumnIndexes) Or IsEmpty(ColumnIndexes) Then
        ColumnIndexes = CreateNumericIndexHeaders
    Else
        If Not IsArray(ColumnIndexes) Then ColumnIndexes = Array(ColumnIndexes)
    End If
    
    ReDim t(0 To GetCount - 1, 0 To UBound(ColumnIndexes) - LBound(ColumnIndexes))
    
    'If UBound(ColumnIndexes) >= 0 Then c0 = CLng(ColumnIndexes(0))
    
    For i = Index To Index + GetCount - 1
        c0 = 0
        For Each c In ColumnIndexes
            If IsObject(Instance(i, CLng(c))) Then
                Set t(i, c0) = Instance(i, CLng(c))
            Else
                t(i, c0) = Instance(i, CLng(c))
            End If
            c0 = c0 + 1
        Next c
    Next i
    
    Set GetRange = New Array2dEx
    GetRange.Instance = CollectionsLib.GetArrayByRef(t)
End Function

'Public Function Resize(Optional ByVal NumRows As Variant, Optional ByVal NumColumns As Variant, Optional ByVal DefaultValue As Variant) As Array2dEx
'    Dim t As New Array2dEx
'End Function

Public Function ToCSV(Optional ByVal Delimiter As String = ",") As String
    Dim t() As String, s() As String, i As Long, c As Long, uc As Long
    uc = Me.ColumnCount - 1
    ReDim t(0 To Me.Count)
    ReDim s(0 To uc)
    
    For i = 0 To Me.Count - 1
        For c = 0 To uc
            s(c) = ValueToCSV(Instance(i, c))
        Next c
        
        t(i) = VBA.Join(s, Delimiter)
    Next i
    
    t(Me.Count) = ""
    ToCSV = VBA.Join(t, vbNewLine)
End Function

Private Function ValueToCSV(ByVal Value As Variant) As String
    Select Case VarType(Value)
        Case vbString
            ValueToCSV = """" & Replace(Replace(CStr(Value), """", """"""), vbCr, "") & """"
        Case vbNull
            ValueToCSV = ""
        Case Else
            ValueToCSV = CStr(Value)
    End Select
End Function

' --- OTHER ---

Private Sub Class_Initialize()
    Instance = Empty
End Sub

Private Sub Class_Terminate()
    Instance = Empty
End Sub





' --- UTILITIES ---

'@Description("Transposes the values in a 2d array. Rows become columns, columns become rows.")
Public Function Transpose() As Array2dEx
    Dim cRow As Long, lbRow As Long, ubRow As Long, cCol As Long, lbCol As Long, ubCol As Long, t() As Variant
    
    lbCol = LBound(Instance, 1)
    ubCol = UBound(Instance, 1)
    lbRow = LBound(Instance, 2)
    ubRow = UBound(Instance, 2)
    ReDim t(lbRow To ubRow, lbCol To ubCol)
    
    For cRow = lbRow To ubRow
        For cCol = lbCol To ubCol
            If IsObject(Instance(cCol, cRow)) Then
                Set t(cRow, cCol) = Instance(cCol, cRow)
            Else
                t(cRow, cCol) = Instance(cCol, cRow)
            End If
        Next
    Next
    
    Set Transpose = New Array2dEx
    Transpose.Instance = CollectionsLib.GetArrayByRef(t)
End Function


' --- CASTING / TYPE CHECKING ---

Private Function C2Date(Optional ByVal Value As Variant) As Variant
    If IsMissing(Value) Then Exit Function
    On Error Resume Next
    C2Date = Value
    C2Date = DateValue(Value)
End Function

Private Function C2Recordset(ByVal Value As Variant) As ADODB.Recordset
    Set C2Recordset = Value
End Function

Private Function C2ExcelRange(ByVal Value As Variant) As Excel.Range
    Set C2ExcelRange = Value
End Function

Private Function C2RecordsetFields(ByVal Value As Variant) As ADODB.Fields
    Set C2RecordsetFields = Value
End Function

Private Function C2ArrayList(ByVal Value As Variant) As ArrayList
    Set C2ArrayList = Value
End Function

Private Function C2Object(ByRef Value As Variant) As Object
    Set C2Object = Value
End Function

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

Private Function Is1dArray(ByRef ArrayLike As Variant) As Boolean
    On Error GoTo Fallback

    Is1dArray = Not (UBound(ArrayLike, 2) >= 0)
    Exit Function
    
Fallback:
    Is1dArray = IsArray(ArrayLike)
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





















'Private Sub ToDoImHere()
'    With Range("A1").CurrentRegion.Columns("EF:EJ")
'        .FormulaR1C1Local = .FormulaR1C1Local
'        .NumberFormatLocal = "#.##0,00;-#.##0,00"
'    End With
'End Sub

'Public Property Get ColumnNames() As Variant()
'    Dim i As Long, sNames() As Variant
'    ReDim sNames(Instance.Fields.Count - 1)
'
'    For i = 0 To Instance.Fields.Count - 1
'        sNames(i) = Instance.Fields(i).Name
'    Next i
'
'    ColumnNames = sNames
'End Property

'Private Function RAWCreateFromRecordset(ByVal RS As ADODB.Recordset) As Array2dEx
'    Dim r As Long, c As Long, t() As Variant, ubRows As Long, ubCols As Long
'
'    ubRows = RS.RecordCount - 1
'    ubCols = RS.Fields.Count - 1
'    If ubRows >= 0 Then
'        t = RS.GetRows(adGetRowsRest, adBookmarkFirst)
'        ReDim Instance(0 To ubRows, 0 To ubCols)
'
'        'On Error Resume Next
'        RS.MoveFirst
'        For c = 0 To ubCols
'            For r = 0 To ubRows
'                Instance(r, c) = t(c, r)
'            Next r
'        Next c
'    End If
'
'    Set RAWCreateFromRecordset = Me
'End Function

'Public Function RAWCopyToRange(ByVal RangeObject As Range) As Array2dEx
'    LogMe "IN Array2dEx.RAWCopyToRange()", True
'    Set RAWCopyToRange = Me
'    If Not Me.Count > 0 Then Exit Function
'
'    With RangeObject.Resize(Me.Count, 1 + (UBound(Instance, 2) - LBound(Instance, 2)))
'        .Value = Instance
'        .FormulaR1C1Local = .FormulaR1C1Local
'    End With
'    LogMe "OUT ...RAWCopyToRange()", True
'End Function

'Private Sub ToDoImHere()
'    With Range("A1").CurrentRegion.Columns("EF:EJ")
'        .FormulaR1C1Local = .FormulaR1C1Local
'        .NumberFormatLocal = "#.##0,00;-#.##0,00"
'    End With
'End Sub





