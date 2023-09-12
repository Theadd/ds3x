VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "dsTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder "ds3x.Collections"
Option Compare Database
Option Explicit
Option Base 0

Implements ICollectionEx

Private pRecords As ICollectionEx
Private pHeaders As ICollectionEx


Public Property Get Count() As Long
    On Error Resume Next
    Count = Records.Count
End Property

Public Property Get ColumnCount() As Long
    On Error Resume Next
    ColumnCount = Records.ColumnCount
    ColumnCount = pHeaders.ColumnCount
End Property

Public Property Get Row(ByVal Index As Long) As Variant
    Row = pRecords.Row(Index)
End Property

Public Property Get Instance() As Scripting.Dictionary
    Set Instance = DictionaryEx.Create() _
        .Add("dsTable", DictionaryEx.Create() _
            .Add("Headers", Headers) _
            .Add("Records", Records) _
            .Instance) _
        .Instance
End Property

Public Property Get Headers() As ICollectionEx: MakeSureHeadersExist: Set Headers = pHeaders: End Property
Public Property Get Records() As ICollectionEx: Set Records = pRecords: End Property

Public Property Get IndexRecordset(Optional ByVal PageSize As Variant, Optional ByVal PageIndex As Variant) As ADODB.Recordset: Set IndexRecordset = CreateIndexRecordset(PageSize, PageIndex): End Property
Public Property Get NamedRecordset() As ADODB.Recordset: Set NamedRecordset = CreateNamedRecordset: End Property


Public Function CreateBlank(ByVal RowsCount As Long, ByVal ColumnsCount As Long) As dsTable: Set CreateBlank = CreateTable(RowsCount, ColumnsCount): End Function

Public Function Create(Optional ByRef TableLike As Variant, Optional ByVal AutoHeaders As Boolean = False) As dsTable
    Set Create = New dsTable
    If Not IsMissing(TableLike) Then Create.Bind TableLike, AutoHeaders
End Function

Public Function Bind(Optional ByRef TableLike As Variant, Optional ByVal AutoHeaders As Boolean = False) As dsTable
    Set Bind = Unbind
    
    If Not IsMissing(TableLike) Then
        If Not IsObject(TableLike) Then
            If IsArray(TableLike) Then
                If Array2dEx.Is2dArray(TableLike) Then
                    Bind Array2dEx.Create(TableLike), AutoHeaders
                ElseIf Array2dEx.Is1dArray(TableLike) Then
                    Bind Array2dEx.Create(TableLike), AutoHeaders
                Else
                    Err.Raise 13
                End If
            ElseIf VarType(TableLike) = vbString Then
                If Left(LTrim(TableLike), 1) Like "[{[]" Then
                    Bind JSON.Parse(TableLike, True, True)
                Else
                    Err.Raise 13
                End If
            Else
                Err.Raise 13
            End If
        Else
            If TypeOf TableLike Is RecordsetEx Then
                Set pRecords = ArraySliceGroup.Create(C2RecordsetEx(TableLike).Clone)
                If AutoHeaders Then SetHeaders C2RecordsetEx(TableLike)
            ElseIf TypeOf TableLike Is ADODB.Recordset Then
                Set pRecords = ArraySliceGroup.Create(RecordsetEx.Create(C2Recordset(TableLike)).Clone)
                If AutoHeaders Then SetHeaders C2Recordset(TableLike)
            ElseIf TypeOf TableLike Is ArrayListEx Then
                Set pRecords = ArraySliceGroup.Create(Array2dEx.Create(TableLike))
            ElseIf TypeOf TableLike Is Array2dEx Then
                Set pRecords = ArraySliceGroup.Create(TableLike)
            ElseIf TypeOf TableLike Is Scripting.Dictionary Then
                Set Bind = CreateFromDictionary(C2Dictionary(TableLike), AutoHeaders)
            ElseIf TypeOf TableLike Is DictionaryEx Then
                Set Bind = CreateFromDictionary(TableLike.Instance, AutoHeaders)
            ElseIf TypeOf TableLike Is ArraySliceGroup Then
                Set pRecords = C2ArraySliceGroup(TableLike)
            Else
                If TypeOf TableLike Is xlSheetsEx Then
                    Set Bind = CreateFromExcelRange(C2xlSheetsEx(TableLike).UsedRange, AutoHeaders)
                ElseIf TypeOf TableLike Is Excel.Worksheet Then
                    Set Bind = CreateFromExcelRange(C2ExcelWorksheet(TableLike).UsedRange, AutoHeaders)
                ElseIf TypeOf TableLike Is Excel.Range Then
                    Set Bind = CreateFromExcelRange(C2ExcelRange(TableLike), AutoHeaders)
                Else
                    Err.Raise 13
                End If
            End If
        End If
    End If
    
End Function

Public Function Unbind() As dsTable
    On Error Resume Next
    Set pHeaders = Nothing
    Set pRecords = Nothing
    Set Unbind = Me
End Function

Public Sub Dispose()
    ' TODO
End Sub

Public Function SetHeaders(ByRef AllHeaders As Variant) As dsTable
    Dim colHeaders As ArrayListEx, i As Long, Aux As Variant
    If Not pHeaders Is Nothing Then
        Err.Raise 13
    Else
        If IsObject(AllHeaders) Then
            If TypeOf AllHeaders Is RecordsetEx Then
                Set pHeaders = CreateColumnHeadersFromRecordsetFields(AllHeaders.Fields)
            ElseIf TypeOf AllHeaders Is ADODB.Recordset Then
                Set pHeaders = CreateColumnHeadersFromRecordsetFields(C2Recordset(AllHeaders).Fields)
            ElseIf TypeOf AllHeaders Is ArrayListEx Then
                Set pHeaders = Array2dEx.Create(AllHeaders)
            ElseIf TypeOf AllHeaders Is Array2dEx Then
                Set pHeaders = AllHeaders
            ElseIf TypeOf AllHeaders Is Excel.Range Then
                Set pHeaders = CreateColumnHeadersFromExcelRange(C2ExcelRange(AllHeaders))
            Else
                Err.Raise 13
            End If
        ElseIf IsArray(AllHeaders) Then
            Set colHeaders = New ArrayListEx
            If UBound(AllHeaders) >= 0 Then
                ' Asuming a plain-text array of column names
                For i = 0 To UBound(AllHeaders)
                    colHeaders.Add CreateColumnHeader(CStr(AllHeaders(i)))
                Next i
                Aux = colHeaders.ToArray
                colHeaders.Clear
                colHeaders.Add Aux
                Set pHeaders = Array2dEx.Create(colHeaders)
            Else
                Set pHeaders = Array2dEx.CreateBlank(1, 0)
            End If
        End If
    End If
    
    Set SetHeaders = Me
End Function

' Concatenate all elements of dsTable as additional columns into a new dsTable
Public Function Join(ByVal TargetTable As dsTable) As dsTable
    Dim dsT As dsTable
    Set dsT = Create(Records.Join(TargetTable.Records))
    Set Join = dsT.SetHeaders(Headers.Join(TargetTable.Headers))
End Function

Public Function AddRange(ByVal TargetTable As dsTable) As dsTable
    Set AddRange = Me _
                .Create(Records.AddRange(TargetTable.Records)) _
                .SetHeaders(Headers.GetRange(0, Headers.Count))
End Function

' @SEEALSO: ArrayListEx.GetRange([Index], [Count], [ColumnIndexes]) As ArrayListEx
Public Function GetRange(Optional ByVal Index As Variant, Optional ByVal GetCount As Variant, Optional ByVal ColumnIndexes As Variant) As dsTable
    Set GetRange = Create(Records.GetRange(Index, GetCount, ColumnIndexes)) _
                     .SetHeaders(Headers.GetRange(0, Headers.Count, ColumnIndexes))
End Function

Public Function ToCSV(Optional ByVal Delimiter As String = ",") As String
    ToCSV = ArrayListEx.Create().Add(GetColumnNames).ToCSV(Delimiter) & Records.ToCSV(Delimiter)
End Function

Public Function ToJSON() As String: ToJSON = JSON.Stringify(Me, 2): End Function

Public Function ToExcel() As String: ToExcel = ToCSV(vbTab): End Function

Public Function ToMarkdownTable() As String
    On Error GoTo Finally
    Dim nCols As Long, nRows As Long, i As Long, k As Long, itemStr As String, itemLen As Long, vRow As Variant, sValue As String
    Dim cNamesX As ArrayListEx, cSizes As ArrayListEx, aX As ArrayListEx, cX As ICollectionEx

    Set aX = ArrayListEx.Create()
    Set cNamesX = ArrayListEx.Create(GetColumnNames())
    Set cSizes = ArrayListEx.Create()
    Set cX = pRecords
    
    nCols = ColumnCount
    nRows = Count

    For i = 0 To nCols - 1
        itemLen = Len(CStr(cNamesX(i)))
        cSizes(i) = CLng(itemLen + (4 - (itemLen Mod 4)))
    Next i

    For k = 0 To nRows - 1
        vRow = cX.Row(k)
        For i = 0 To nCols - 1
            itemLen = Len(CStr(Nz(vRow(i), vbNullString)))
            itemLen = itemLen + (4 - (itemLen Mod 4))
            If CLng(cSizes(i)) < itemLen Then cSizes(i) = itemLen
        Next i
    Next k
    
    ' PRINTING HEADERS
    itemStr = vbNullString
    For k = 0 To nCols - 1
        itemStr = itemStr & " | " & cNamesX(k) & Space(cSizes(k) - Len(cNamesX(k)))
    Next k
    aX.Add Mid$(itemStr, 4, Len(itemStr))
    aX.Add String(Len(aX(0)), "-")
    
    ' PRINTING TABLE ROWS
    For k = 0 To nRows - 1
        vRow = cX.Row(k)
        itemStr = vbNullString
        For i = 0 To nCols - 1
            sValue = CStr(Nz(vRow(i), vbNullString))
            itemStr = itemStr & " | " & sValue & Space(cSizes(i) - Len(sValue))
        Next i
        aX.Add Mid$(itemStr, 4, Len(itemStr))
    Next k

Finally:
    ToMarkdownTable = VBA.Join(aX.ToArray(), vbNewLine)
End Function

Private Sub ApplyFormatToExcelRange(ByVal RangeObject As Excel.Range, ByVal RangeHasHeaders As Boolean)
    Dim r As Excel.Range, i As Long, iX As Variant
    
    If RangeHasHeaders Then
        Set r = RangeObject.Range("A2").Resize(Count, ColumnCount)
    Else
        Set r = RangeObject.Range("A1").Resize(Count, ColumnCount)
    End If

    iX = Headers.Row(0)
    For i = 0 To UBound(iX)
        If CollectionsLib.ArrayItem(iX(i), 0, "General") <> "General" Then
            r.Columns(i + 1).NumberFormat = CollectionsLib.ArrayItem(iX(i), 1, "@")
        End If
    Next i
End Sub

Public Function CopyToRange(ByRef RangeObject As Excel.Range, Optional ByVal ApplyUserLocale As Boolean = True, Optional ByVal WriteHeaders As Boolean = True) As dsTable
    If WriteHeaders Then
        Array2dEx.Create(Array(GetColumnNames)).CopyToRange RangeObject, False
        ApplyFormatToExcelRange RangeObject, WriteHeaders
        Records.CopyToRange RangeObject.Range("A2"), ApplyUserLocale
    Else
        ApplyFormatToExcelRange RangeObject, WriteHeaders
        Records.CopyToRange RangeObject, ApplyUserLocale
    End If
End Function


Public Function CreateColumnHeader(Optional ByVal ColumnName As Variant, Optional ByVal NumberFormat As Variant, Optional ByVal DataType As Variant, Optional ByVal MaxSize As Variant) As Variant
    CreateColumnHeader = Array(ColumnName, NumberFormat, DataType, MaxSize)
End Function

Public Function CreateColumnHeaderFromRecordsetField(ByRef TargetField As ADODB.Field) As Variant
    Dim d As Variant
    
    With TargetField
        Select Case .Type
            Case adVarChar
                d = CreateColumnHeader(Nz(.Name), NumberFormat:="@", DataType:="Text", MaxSize:=.DefinedSize)
            Case adVarWChar
'                If .DefinedSize = 10 Then
'                    d = CreateColumnHeader(Nz(.Name), NumberFormat:="m/d/yyyy", DataType:="Date", MaxSize:=10)
'                Else
                    d = CreateColumnHeader(Nz(.Name), NumberFormat:="@", DataType:="Text", MaxSize:=.DefinedSize)
'                End If
            Case adInteger, adBigInt, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                d = CreateColumnHeader(Nz(.Name), NumberFormat:="0", DataType:="Numeric")
            Case adNumeric
                If .NumericScale < 16 And .NumericScale > 0 Then
                    d = CreateColumnHeader(Nz(.Name), NumberFormat:="0." & String(.NumericScale, "0"), DataType:="Numeric")
                Else
                    d = CreateColumnHeader(Nz(.Name), NumberFormat:="0", DataType:="Numeric")
                End If
            Case adCurrency
                d = CreateColumnHeader(Nz(.Name), NumberFormat:="0.00", DataType:="Currency")
            Case adDBTimeStamp
                d = CreateColumnHeader(Nz(.Name), NumberFormat:="m/d/yyyy h:mm", DataType:="DateTime", MaxSize:=.DefinedSize)
            Case adLongVarChar, adLongVarWChar
                d = CreateColumnHeader(Nz(.Name), NumberFormat:="@", DataType:="Text", MaxSize:=-1)
            Case Else
                d = CreateColumnHeader(Nz(.Name))
        End Select
        
        CreateColumnHeaderFromRecordsetField = d
    End With
End Function


' --- OTHER ---

Private Sub Class_Terminate()
    On Error Resume Next
    Unbind
End Sub

Private Function CreateTable(ByVal RowsCount As Long, ByVal ColumnsCount As Long) As dsTable
    Dim t() As Variant, i As Long, k() As Variant, dsT As dsTable, aX As ArrayListEx
    
    Set dsT = dsTable.Create(Array2dEx.CreateBlank(RowsCount, ColumnsCount), False)

    If ColumnsCount > 0 Then
        ReDim k(0 To ColumnsCount - 1)
        For i = 0 To ColumnsCount - 1
            k(i) = vbNullString
        Next i
        dsT.SetHeaders k
    Else
        dsT.SetHeaders Array()
    End If
    
    Set CreateTable = dsT
End Function

Private Function CreateFromExcelRange(ByRef xlRange As Excel.Range, Optional ByVal AutoHeaders As Boolean = False) As dsTable
    With xlRange
        If AutoHeaders Then
            Set pRecords = ArraySliceGroup.Create(.Range(.Cells(2, 1), .Cells(.Rows.Count, .Columns.Count)))
            SetHeaders xlRange
        Else
            Set pRecords = ArraySliceGroup.Create(.Range(.Cells(1, 1), .Cells(.Rows.Count, .Columns.Count)))
        End If
    End With
    
    Set CreateFromExcelRange = Me
End Function

Private Function CreateFromDictionary(ByRef Target As Scripting.Dictionary, Optional ByVal AutoHeaders As Boolean = False) As dsTable
    Dim dX As DictionaryEx
    On Error GoTo Fallback
    
    Set dX = DictionaryEx.Create(Target)
    If Not dX.Exists("dsTable") Then
        Set CreateFromDictionary = CreateFromKeyValueDictionary(Target, AutoHeaders)
        Exit Function
    End If
    Set dX = DictionaryEx.Create(dX("dsTable"))
    Set pRecords = ArraySliceGroup.Create(Array2dEx.Create(dX("Records")))
    Set pHeaders = Array2dEx.Create(ArrayListEx.Create().AddRange(dX("Headers")))
    
    Set CreateFromDictionary = Me
    Exit Function
Fallback:
    Err.Raise 13, , "Unable to create a dsTable from provided dictionary."
End Function

Private Function CreateFromKeyValueDictionary(ByRef Target As Scripting.Dictionary, Optional ByVal AutoHeaders As Boolean = False) As dsTable
    Dim dX As DictionaryEx, dKey As Variant, aX As ArrayListEx
    On Error GoTo Fallback
    
    Set dX = DictionaryEx.Create(Target)
    Set aX = ArrayListEx.Create()
    For Each dKey In dX.Keys()
        aX.Add Array(dKey, dX(dKey))
    Next dKey
    Set pRecords = ArraySliceGroup.Create(Array2dEx.Create(aX))
    If AutoHeaders Then
        SetHeaders Array("Key", "Value")
    End If
    
    Set CreateFromKeyValueDictionary = Me
    Exit Function
Fallback:
    Err.Raise 13, , "Unable to create a dsTable from provided dictionary."
End Function

Private Function CreateNamedRecordset() As ADODB.Recordset
    Dim rs As New ADODB.Recordset, i As Long, c As Long, cMax As Long, cMin As Long, iRow As Variant, AllValues As ArrayListEx, vHeads As Variant
    iRow = GetColumnNames
    cMax = UBound(iRow)
    cMin = LBound(iRow)
    Set AllValues = CollectionsLib.AsArrayListEx(Records)
    vHeads = Headers.Row(0)
    
    With rs
        For c = cMin To cMax
            Select Case CollectionsLib.ArrayItem(vHeads(c), 2, "Text")
                Case "Text"
                    i = CLng(CollectionsLib.ArrayItem(vHeads(c), 3, -1))
                    Select Case i
                        Case -1: .Fields.Append CStr(iRow(c)), adLongVarWChar, -1, adFldIsNullable Or adFldMayBeNull
                        Case Else: .Fields.Append CStr(iRow(c)), adBSTR, i, adFldIsNullable Or adFldMayBeNull
                    End Select
                Case "DateTime": .Fields.Append CStr(iRow(c)), adDBTimeStamp, CLng(CollectionsLib.ArrayItem(vHeads(c), 3, 16)), adFldIsNullable Or adFldMayBeNull Or adFldFixed
                Case Else: .Fields.Append CStr(iRow(c)), adLongVarWChar, -1, adFldIsNullable Or adFldMayBeNull
            End Select
        Next c
        .Open
        For i = 0 To Count - 1
            .AddNew FieldList:=iRow, Values:=AllValues(i)
        Next i
        If Count > 0 Then .MoveFirst
    End With
    
    Set CreateNamedRecordset = rs
End Function


Public Function CreateIndexRecordset( _
        Optional ByVal PageSize As Variant, _
        Optional ByVal PageIndex As Variant, _
        Optional ByVal PageCount As Long = 1, _
        Optional ByVal ColumnStartIndex As Variant, _
        Optional ByVal ColumnsCount As Variant, _
        Optional ByVal ColumnIndexStartsOnZero As Boolean = False _
        ) As ADODB.Recordset
    Dim rs As New ADODB.Recordset, i As Long, c As Long, cMax As Long, cMin As Long, iRow As Variant, AllValues As ArrayListEx, iMax As Long
    
    If IsMissing(PageSize) Then PageSize = Count
    If IsMissing(PageIndex) Then PageIndex = 0
    If IsMissing(ColumnStartIndex) Then ColumnStartIndex = 0
    If IsMissing(ColumnsCount) Then ColumnsCount = ColumnCount - ColumnStartIndex
    
    iRow = CreateIndexHeaders(ColumnStartIndex, ColumnsCount, ColumnIndexStartsOnZero)
    cMax = UBound(iRow)
    cMin = LBound(iRow)
    iMax = ((PageSize * (PageIndex + 1)) - 1) + ((PageCount - 1) * PageSize)
    If iMax > Count - 1 Then iMax = Count - 1
        
    If ColumnStartIndex <> 0 Or ColumnsCount <> ColumnCount Then
        Set AllValues = CollectionsLib.AsArrayListEx( _
                            Records.GetRange( _
                                PageSize * PageIndex, _
                                iMax - (PageSize * PageIndex) + 1, _
                                CreateIndexHeaders(ColumnStartIndex, ColumnsCount, False)))
    Else
        ' TODO: Slice rows up to requested ones
        Set AllValues = CollectionsLib.AsArrayListEx(Records)
    End If
    
    With rs
        For c = cMin To cMax
            .Fields.Append CStr(iRow(c)), adLongVarWChar, -1, adFldIsNullable
        Next c
        .Open
        For i = 0 To iMax - (PageSize * PageIndex)
            .AddNew FieldList:=iRow, Values:=AllValues(i)
        Next i
        If Count > 0 Then .MoveFirst
    End With
    
    Set CreateIndexRecordset = rs
End Function

Private Function CreateIndexHeaders(ByVal ColumnStartIndex As Long, ByVal ColumnsCount As Long, ByVal ColumnIndexStartsOnZero As Boolean) As Variant
    Dim t As Variant, c As Long, cMax As Long, cMin As Long
    cMax = (ColumnsCount - 1) ' + ColumnStartIndex
    
    If Not ColumnIndexStartsOnZero Then
        cMin = ColumnStartIndex
    End If
    
    ReDim t(0 To cMax)
    
    For c = 0 To cMax
        t(c) = CStr(c + cMin)
    Next c
    
    CreateIndexHeaders = t
End Function

Public Function GetColumnNames(Optional ByVal FirstColumnIndex As Long = 0, Optional ByVal GetCount As Variant) As Variant
    Dim t As Variant, c As Long, hItems As Variant, sLimit As Long
    
    If IsMissing(GetCount) Then GetCount = ColumnCount - FirstColumnIndex
    If GetCount = 0 Then
        GetColumnNames = Array()
        Exit Function
    End If
    ReDim t(0 To GetCount - 1)
    hItems = Headers.Row(0)
    sLimit = IIf(ColumnCount - FirstColumnIndex < GetCount, ColumnCount - FirstColumnIndex, GetCount)
    
    For c = 0 To sLimit - 1
        t(c) = CStr(CollectionsLib.ArrayItem(hItems(c + FirstColumnIndex), 0, vbNullString))
    Next c
    If sLimit < 0 Then sLimit = 0
    For c = sLimit To GetCount - 1
        t(c) = vbNullString
    Next c
    
    GetColumnNames = t
End Function


' --- Create Column Headers ---

Private Function CreateColumnHeadersFromRecordsetFields(ByVal AllFields As ADODB.Fields) As Array2dEx
    Dim i As Long, t() As Variant
    ReDim t(0 To 0, 0 To AllFields.Count - 1)
    
    For i = 0 To AllFields.Count - 1
        t(0, i) = CreateColumnHeaderFromRecordsetField(AllFields(i))
    Next i

    Set CreateColumnHeadersFromRecordsetFields = New Array2dEx
    CreateColumnHeadersFromRecordsetFields.Instance = CollectionsLib.GetArrayByRef(t)
End Function

Private Function CreateColumnHeadersFromExcelRange(ByVal TargetRange As Excel.Range) As Array2dEx
    Dim i As Long, t() As Variant, cFormats As Variant, cNames As Variant
    
    With TargetRange
        ReDim t(0 To 0, 0 To .Columns.Count - 1)
        cNames = ArrayListEx.Create(.Range(.Cells(1, 1), .Cells(1, .Columns.Count)))(0)
    End With
    cFormats = xlSheetsEx.GetColumnsAutoNumberFormats(TargetRange)
    
    For i = 0 To UBound(cFormats)
        Select Case UCase(Left(CStr(cFormats(i)(0)), 1))
            Case "G"
                t(0, i) = CreateColumnHeader(cNames(i))
            Case "0", "#"
                t(0, i) = CreateColumnHeader(cNames(i), NumberFormat:=cFormats(i)(0), DataType:="Numeric")
            Case "@"
                t(0, i) = CreateColumnHeader(cNames(i), NumberFormat:=cFormats(i)(0), DataType:="Text")
            Case "D", "M", "Y", "H"
                t(0, i) = CreateColumnHeader(cNames(i), NumberFormat:=cFormats(i)(0), DataType:="Date")
            Case Else
                t(0, i) = CreateColumnHeader(cNames(i), NumberFormat:=cFormats(i)(0))
        End Select
    Next i

    Set CreateColumnHeadersFromExcelRange = New Array2dEx
    CreateColumnHeadersFromExcelRange.Instance = CollectionsLib.GetArrayByRef(t)
End Function

Private Sub MakeSureHeadersExist()
    If Not pHeaders Is Nothing Then Exit Sub
    Dim t() As Variant, i As Long
    
    ReDim t(0 To 0, 0 To ColumnCount - 1)
    For i = 0 To ColumnCount - 1
        t(0, i) = Array()
    Next i
    SetHeaders Array2dEx.Create(t)
End Sub


' --- TYPE CASTING ---

Private Function C2Recordset(ByVal Value As Variant) As ADODB.Recordset: Set C2Recordset = Value: End Function
Private Function C2RecordsetEx(ByVal Value As Variant) As RecordsetEx: Set C2RecordsetEx = Value: End Function
Private Function C2ExcelRange(ByVal Value As Variant) As Excel.Range: Set C2ExcelRange = Value: End Function
Private Function C2ExcelWorksheet(ByVal Value As Variant) As Excel.Worksheet: Set C2ExcelWorksheet = Value: End Function
Private Function C2xlSheetsEx(ByVal Value As Variant) As xlSheetsEx: Set C2xlSheetsEx = Value: End Function
Private Function C2Dictionary(ByVal Value As Variant) As Scripting.Dictionary: Set C2Dictionary = Value: End Function
Private Function C2ArraySliceGroup(ByVal Value As Variant) As ArraySliceGroup: Set C2ArraySliceGroup = Value: End Function


' --- ICollectionEx ---

Private Property Get ICollectionEx_Count() As Long: ICollectionEx_Count = Count: End Property
Private Property Get ICollectionEx_ColumnCount() As Long: ICollectionEx_ColumnCount = ColumnCount: End Property
Private Property Get ICollectionEx_Item(ByVal Index As Long, Optional ByVal ColumnIndex As Long = 0&) As Variant: ICollectionEx_Item = Records(Index, ColumnIndex): End Property
Private Property Get ICollectionEx_Row(ByVal Index As Long) As Variant: ICollectionEx_Row = Row(Index): End Property
Private Function ICollectionEx_CreateBlank(ByVal RowsCount As Long, ByVal ColumnsCount As Long) As ICollectionEx: Set ICollectionEx_CreateBlank = CreateBlank(RowsCount, ColumnsCount): End Function
Private Function ICollectionEx_Create(Optional ByRef FromTarget As Variant) As ICollectionEx: Set ICollectionEx_Create = Create(FromTarget, True): End Function
Private Function ICollectionEx_Bind(Optional ByRef Target As Variant) As ICollectionEx: Set ICollectionEx_Bind = Bind(Target): End Function
Private Function ICollectionEx_Unbind() As ICollectionEx: Set ICollectionEx_Unbind = Unbind(): End Function
Private Sub ICollectionEx_Dispose(): Dispose: End Sub
Private Function ICollectionEx_Join(ByVal Target As ICollectionEx) As ICollectionEx: Set ICollectionEx_Join = Join(Target): End Function
Private Function ICollectionEx_AddRange(ByVal Target As ICollectionEx) As ICollectionEx: Set ICollectionEx_AddRange = AddRange(Target): End Function
Private Function ICollectionEx_GetRange(Optional ByVal Index As Variant, Optional ByVal GetCount As Variant, Optional ByVal ColumnIndexes As Variant) As ICollectionEx: Set ICollectionEx_GetRange = GetRange(Index, GetCount, ColumnIndexes): End Function
Private Function ICollectionEx_ToCSV(Optional ByVal Delimiter As String = ",") As String: ICollectionEx_ToCSV = ToCSV(Delimiter): End Function
Private Function ICollectionEx_ToJSON() As String: ICollectionEx_ToJSON = ToJSON(): End Function
Private Function ICollectionEx_ToExcel() As String: ICollectionEx_ToExcel = ToExcel(): End Function
Private Function ICollectionEx_CopyToRange(ByRef RangeObject As Excel.Range, Optional ByVal ApplyUserLocale As Boolean = True, Optional ByVal WriteHeaders As Boolean = True) As ICollectionEx: Set ICollectionEx_CopyToRange = CopyToRange(RangeObject, ApplyUserLocale, WriteHeaders): End Function

