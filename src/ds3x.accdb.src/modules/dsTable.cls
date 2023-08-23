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


Private pHeaderList As ArrayListEx
Private pHeaderArray As Array2dEx
Private pRecordList As ArrayListEx
Private pRecordArray As Array2dEx
Private pTableSource As Variant
Private Const dsChunkSize As Long = 4000



Public Property Get Count() As Long
    On Error Resume Next
    If pRecordArray Is Nothing Then
        If pRecordList Is Nothing Then
            Count = pTableSource.Count
        Else
            Count = pRecordList.Count
        End If
    Else
        Count = pRecordArray.Count
    End If
End Property

Public Property Get ColumnCount() As Long
    On Error Resume Next
    If pHeaderArray Is Nothing Then
        If pHeaderList Is Nothing Then
            ColumnCount = Me.Records.ColumnCount
        Else
            'ColumnCount = 1 + UBound(Me.HeaderList.Item(0)) - LBound(Me.HeaderList.Item(0))
            ColumnCount = pHeaderList.ColumnCount
        End If
    Else
        ColumnCount = pHeaderArray.ColumnCount
    End If
End Property

Public Property Get Instance() As Scripting.Dictionary
    Set Instance = DictionaryEx.Create() _
        .Add("dsTable", DictionaryEx.Create() _
            .Add("Headers", Me.Headers) _
            .Add("Records", Me.Records) _
            .Instance) _
        .Instance
End Property

Public Property Get HeaderList() As ArrayListEx
    If pHeaderList Is Nothing Then RebuildHeaderList
    Set HeaderList = pHeaderList
End Property

Public Property Get HeaderArray() As Array2dEx
    If pHeaderArray Is Nothing Then RebuildHeaderArray
    Set HeaderArray = pHeaderArray
End Property

Public Property Get RecordList() As ArrayListEx
    If pRecordList Is Nothing Then RebuildRecordList
    Set RecordList = pRecordList
End Property

Public Property Get RecordArray() As Array2dEx
    If pRecordArray Is Nothing Then RebuildRecordArray
    Set RecordArray = pRecordArray
End Property

Public Property Get Headers() As Variant
    If pHeaderList Is Nothing Then
        Set Headers = HeaderArray
    Else
        Set Headers = pHeaderList
    End If
End Property

Public Property Get Records() As Variant
    If pRecordList Is Nothing Then
        Set Records = RecordArray
    Else
        Set Records = pRecordList
    End If
End Property

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
                    Bind ArrayListEx.Create(TableLike), AutoHeaders
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
                Set pTableSource = TableLike    '.Clone
                If AutoHeaders Then Me.SetHeaders pTableSource
            ElseIf TypeOf TableLike Is ADODB.Recordset Then
                Set pTableSource = RecordsetEx.Create(C2Recordset(TableLike))
                If AutoHeaders Then Me.SetHeaders pTableSource
            ElseIf TypeOf TableLike Is ArrayListEx Then
                Set pRecordList = TableLike
            ElseIf TypeOf TableLike Is Array2dEx Then
                Set pRecordArray = TableLike
            ElseIf TypeOf TableLike Is Scripting.Dictionary Then
                Set Bind = CreateFromDictionary(C2Dictionary(TableLike), AutoHeaders)
            ElseIf TypeOf TableLike Is DictionaryEx Then
                Set Bind = CreateFromDictionary(TableLike.Instance, AutoHeaders)
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
    
    Set pHeaderList = Nothing
    Set pHeaderArray = Nothing
    Set pRecordList = Nothing
    Set pRecordArray = Nothing
    If IsObject(pTableSource) Then Set pTableSource = Nothing
    
    Set Unbind = Me
End Function

Public Sub Dispose()
    ' EDIT: DISPOSE
    ' Unbind
End Sub

Public Function SetHeaders(ByRef AllHeaders As Variant) As dsTable
    Dim colHeaders As ArrayListEx, i As Long, Aux As Variant
    If Not (pHeaderList Is Nothing And pHeaderArray Is Nothing) Then
        Err.Raise 13
    Else
        If IsObject(AllHeaders) Then
            If TypeOf AllHeaders Is RecordsetEx Then
                Set pHeaderArray = CreateColumnHeadersFromRecordsetFields(AllHeaders.Fields)
            ElseIf TypeOf AllHeaders Is ADODB.Recordset Then
                Set pHeaderArray = CreateColumnHeadersFromRecordsetFields(C2Recordset(AllHeaders).Fields)
            ElseIf TypeOf AllHeaders Is ArrayListEx Then
                Set pHeaderList = AllHeaders
            ElseIf TypeOf AllHeaders Is Array2dEx Then
                Set pHeaderArray = AllHeaders
            ElseIf TypeOf AllHeaders Is Excel.Range Then
                Set pHeaderArray = CreateColumnHeadersFromExcelRange(C2ExcelRange(AllHeaders))
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
            End If
            Aux = colHeaders.ToArray
            colHeaders.Clear
            colHeaders.Add Aux
            Set pHeaderList = colHeaders
        End If
    End If
    
    Set SetHeaders = Me
End Function


' Concatenate all elements of dsTable as additional columns into a new dsTable
Public Function Join(ByVal TargetTable As dsTable) As dsTable
    Dim dsT As dsTable
    
    Select Case 0
        Case Me.Count, Me.ColumnCount, TargetTable.ColumnCount, TargetTable.Count
            Set dsT = Me.Create(Me.RecordList.Join(TargetTable.RecordList))
        Case Else
            Set dsT = Me.Create(Me.RecordArray.Join(TargetTable.RecordArray))
    End Select

    Set Join = dsT.SetHeaders(Me.HeaderList.Join(TargetTable.HeaderList))
End Function

Public Function AddRange(ByVal TargetTable As dsTable) As dsTable
    Set AddRange = Me _
                .Create(Me.RecordList.AddRange(TargetTable.RecordList)) _
                .SetHeaders(Me.Headers.GetRange(0, Me.Headers.Count))
End Function

'
'
' @SEEALSO: ArrayListEx.GetRange([Index], [Count], [ColumnIndexes]) As ArrayListEx
Public Function GetRange(Optional ByVal Index As Variant, Optional ByVal GetCount As Variant, Optional ByVal ColumnIndexes As Variant) As dsTable
    Set GetRange = Me.Create(GetRangeOfRecords(Index, GetCount, ColumnIndexes)) _
                     .SetHeaders(Me.Headers.GetRange(0, Me.Headers.Count, ColumnIndexes))
End Function

Friend Function GetRangeOfRecords(Optional ByVal Index As Variant, Optional ByVal GetCount As Variant, Optional ByVal ColumnIndexes As Variant) As Variant
    If (pRecordList Is Nothing) And (pRecordArray Is Nothing) And (TypeOf pTableSource Is RecordsetEx) Then
        Set GetRangeOfRecords = C2RecordsetEx(pTableSource).GetRange(Index, GetCount, ColumnIndexes)
    Else
        Set GetRangeOfRecords = Me.Records.GetRange(Index, GetCount, ColumnIndexes)
    End If
End Function

'Public Function Resize(Optional ByVal NumRows As Variant, Optional ByVal NumColumns As Variant, Optional ByVal DefaultValue As Variant) As Array2dEx
'    Dim t As New Array2dEx
'End Function

Public Function ToCSV(Optional ByVal Delimiter As String = ",") As String
    ToCSV = ArrayListEx.Create().Add(GetColumnNames).ToCSV(Delimiter) & Me.Records.ToCSV(Delimiter)
End Function

Public Function ToJSON() As String: ToJSON = JSON.Stringify(Me, 2): End Function

Public Function ToExcel() As String: ToExcel = ToCSV(vbTab): End Function

Private Sub ApplyFormatToExcelRange(ByVal RangeObject As Excel.Range, ByVal RangeHasHeaders As Boolean)
    Dim r As Excel.Range, i As Long, iX As Variant, dX As DictionaryEx
    
    If RangeHasHeaders Then
        Set r = RangeObject.Range("A2").Resize(Me.Count, Me.ColumnCount)
    Else
        Set r = RangeObject.Range("A1").Resize(Me.Count, Me.ColumnCount)
    End If
    
    Set dX = DictionaryEx.Create()
    iX = Me.HeaderList(0)
    For i = 0 To UBound(iX)
        dX.Bind iX(i)
        If dX("NumberFormat") <> "General" Then
            r.Columns(i + 1).NumberFormat = dX("NumberFormat")
        End If
    Next i
End Sub

Public Function CopyToRange(ByRef RangeObject As Excel.Range, Optional ByVal ApplyUserLocale As Boolean = True, Optional ByVal WriteHeaders As Boolean = True) As dsTable
    Dim aX As ArrayListEx, cRow As Long, cRange As Excel.Range, i As Long, iMax As Long, rCount As Long, bX As ArrayListEx, a2dX As Array2dEx, iNumRows As Long
    
    If WriteHeaders Then
        Set aX = ArrayListEx.Create(Array(GetColumnNames()))
        Array2dEx.Create(aX).CopyToRange RangeObject, False
        aX.Dispose
        Set aX = Nothing
        cRow = cRow + 1
    End If
    ApplyFormatToExcelRange RangeObject, WriteHeaders
    
    ' TODO: Implement chunked copy of table source ranges when source recordset is way too large to be dealt at once in Win32 due to 2GB virtual memory limit.
    If pRecordArray Is Nothing And pRecordList Is Nothing Then RebuildRecordsFromTableSource

    If pRecordArray Is Nothing Then
        On Error GoTo ErrorHandler
        rCount = pRecordList.Count
        
        iNumRows = rCount - (i * dsChunkSize)
        If iNumRows > dsChunkSize Then iNumRows = dsChunkSize

        While (iNumRows > 0)
            ' TODO: To ~one-liner
            Set bX = pRecordList.GetRange(i * dsChunkSize, iNumRows)
            Set a2dX = Array2dEx.Create(bX)
            a2dX.CopyToRange RangeObject.Range("A" & CStr(1 + cRow)), ApplyUserLocale
            a2dX.Dispose
            bX.Dispose
            i = i + 1
            cRow = cRow + iNumRows
            iNumRows = rCount - (i * dsChunkSize)
            If iNumRows > dsChunkSize Then iNumRows = dsChunkSize
        Wend
        'GlobalVMemAnalysis "@dsTable.CopyToRange P2A"
    Else
        pRecordArray.CopyToRange RangeObject.Range("A" & CStr(1 + cRow)), ApplyUserLocale
        'GlobalVMemAnalysis "@dsTable.CopyToRange P2B"
    End If
    
    ' Me.RecordArray.CopyToRange RangeObject.Range("A2"), ApplyUserLocale
    
ExitFunc:
    Set CopyToRange = Me
    Exit Function
ErrorHandler:
    On Error GoTo ExitFunc
    Debug.Print CStr(Err.Number) & " - " & Err.Description
    Resume Next
End Function

Friend Function CopyToRangeNonChunked(ByRef RangeObject As Excel.Range, Optional ByVal ApplyUserLocale As Boolean = True) As dsTable
    Dim hList As New ArrayListEx
    
    hList.Add GetHeaderNames
    Array2dEx.Create(hList).CopyToRange RangeObject, False
    Me.RecordArray.CopyToRange RangeObject.Range("A2"), ApplyUserLocale
    
    Set CopyToRangeNonChunked = Me
End Function



' --- PUBLIC STATIC ---

Public Function CreateColumnHeader(ByVal ColumnName As String, Optional ByVal NumberFormat As String = "General", Optional ByVal DataType As String = "", Optional ByVal MaxSize As Long = -1) As Scripting.Dictionary
    Dim d As New Scripting.Dictionary
    d("ColumnName") = ColumnName
    d("NumberFormat") = NumberFormat
    d("DataType") = DataType
    d("MaxSize") = MaxSize
    Set CreateColumnHeader = d
End Function

Public Function CreateColumnHeaderFromRecordsetField(ByRef TargetField As ADODB.Field) As Scripting.Dictionary
    Dim d As Scripting.Dictionary
    
    With TargetField
        Select Case .Type
            Case adVarChar
                Set d = CreateColumnHeader(Nz(.Name), DataType:="Text", MaxSize:=.DefinedSize)
            Case adVarWChar
                If .DefinedSize = 10 Then
                    Set d = CreateColumnHeader(Nz(.Name), DataType:="Date", MaxSize:=10)
                Else
                    Set d = CreateColumnHeader(Nz(.Name), DataType:="Text", MaxSize:=.DefinedSize)
                End If
            Case adInteger, adBigInt, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                Set d = CreateColumnHeader(Nz(.Name), NumberFormat:="0", DataType:="Numeric")
            Case adNumeric
                If .NumericScale < 16 And .NumericScale > 0 Then
                    Set d = CreateColumnHeader(Nz(.Name), NumberFormat:="0." & String(.NumericScale, "0"), DataType:="Numeric")
                Else
                    Set d = CreateColumnHeader(Nz(.Name), NumberFormat:="0", DataType:="Numeric")
                End If
            Case adCurrency
                Set d = CreateColumnHeader(Nz(.Name), NumberFormat:="0.00", DataType:="Currency")
            Case adDBTimeStamp
                Set d = CreateColumnHeader(Nz(.Name), DataType:="DateTime") ' , MaxSize:=.DefinedSize
            Case adLongVarChar, adLongVarWChar
                Set d = CreateColumnHeader(Nz(.Name), DataType:="Text", MaxSize:=-1)
            Case Else
                Set d = CreateColumnHeader(Nz(.Name))
        End Select
        
        Set CreateColumnHeaderFromRecordsetField = d
    End With
End Function



' --- OTHER ---

Private Sub Class_Terminate()
    On Error Resume Next
    Unbind
End Sub





Private Sub RebuildHeaderList()
    If pHeaderArray Is Nothing Then
        Set pHeaderList = New ArrayListEx
    Else
        Set pHeaderList = ArrayListEx.Create(pHeaderArray)
    End If
End Sub

Private Sub RebuildHeaderArray()
    If pHeaderList Is Nothing Then
        Set pHeaderArray = New Array2dEx
    Else
        Set pHeaderArray = Array2dEx.Create(pHeaderList)
    End If
End Sub

Private Sub RebuildRecordList()
    If pRecordArray Is Nothing Then
        RebuildRecordsFromTableSource
    End If
    If pRecordList Is Nothing And Not pRecordArray Is Nothing Then
        Set pRecordList = ArrayListEx.Create(pRecordArray)
    End If
End Sub

Private Sub RebuildRecordArray()
    If pRecordList Is Nothing Then
        RebuildRecordsFromTableSource
    End If
    If pRecordArray Is Nothing And Not pRecordList Is Nothing Then
        Set pRecordArray = Array2dEx.Create(pRecordList)
    End If
End Sub

Private Sub RebuildRecordsFromTableSource()
    If IsObject(pTableSource) And Not pTableSource Is Nothing Then
        If TypeOf pTableSource Is ADODB.Recordset Or TypeOf pTableSource Is RecordsetEx Then
            Set pRecordArray = Array2dEx.Create(pTableSource)
        Else
            Err.Raise 13
        End If
    Else
        Set pRecordList = New ArrayListEx
        Set pRecordArray = New Array2dEx
    End If
End Sub

Private Function CreateTable(ByVal RowsCount As Long, ByVal ColumnsCount As Long) As dsTable
    Dim t() As Variant, i As Long, k() As Variant, dsT As dsTable, aX As ArrayListEx
    
    If RowsCount > 0 And ColumnsCount > 0 Then
        ReDim t(0 To RowsCount - 1, 0 To ColumnsCount - 1)
        Set dsT = dsTable.Create(Array2dEx.Create(t))
    Else
        Set aX = ArrayListEx.Create()
        If RowsCount > 0 Then
            For i = 0 To RowsCount - 1
                aX.Add Array()
            Next i
        End If
        Set dsT = dsTable.Create(aX)
    End If
    
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
            Set pRecordList = ArrayListEx.Create(.Range(.Cells(2, 1), .Cells(.Rows.Count, .Columns.Count)))
            Me.SetHeaders xlRange
        Else
            Set pRecordList = ArrayListEx.Create(.Range(.Cells(1, 1), .Cells(.Rows.Count, .Columns.Count)))
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
    Set pRecordList = ArrayListEx.Create()
    pRecordList.AddRange dX("Records")
    Set pHeaderList = ArrayListEx.Create()
    pHeaderList.AddRange dX("Headers")
    
    Set CreateFromDictionary = Me
    Exit Function
Fallback:
    Err.Raise 13, , "Unable to create a dsTable from provided dictionary."
End Function

Private Function CreateFromKeyValueDictionary(ByRef Target As Scripting.Dictionary, Optional ByVal AutoHeaders As Boolean = False) As dsTable
    Dim dX As DictionaryEx, dKey As Variant
    On Error GoTo Fallback
    
    Set dX = DictionaryEx.Create(Target)
    Set pRecordList = ArrayListEx.Create()
    For Each dKey In dX.Keys()
        pRecordList.Add Array(dKey, dX(dKey))
    Next dKey
    If AutoHeaders Then
        SetHeaders Array("Key", "Value")
    End If
    
    Set CreateFromKeyValueDictionary = Me
    Exit Function
Fallback:
    Err.Raise 13, , "Unable to create a dsTable from provided dictionary."
End Function

Private Function CreateNamedRecordset() As ADODB.Recordset
    Dim rs As New ADODB.Recordset, i As Long, c As Long, cMax As Long, cMin As Long, iRow As Variant, AllValues As ArrayListEx
    
    iRow = GetHeaderNames
    cMax = UBound(iRow)
    cMin = LBound(iRow)
    Set AllValues = Me.RecordList
    
    With rs
        For c = cMin To cMax
            .Fields.Append CStr(iRow(c)), adLongVarWChar, -1, adFldIsNullable
        Next c
        
        .Open
        
        For i = 0 To Me.Count - 1
            .AddNew FieldList:=iRow, Values:=AllValues(i)
        Next i
        
        If Me.Count > 0 Then .MoveFirst
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
    
    If IsMissing(PageSize) Then PageSize = Me.Count
    If IsMissing(PageIndex) Then PageIndex = 0
    If IsMissing(ColumnStartIndex) Then ColumnStartIndex = 0
    If IsMissing(ColumnsCount) Then ColumnsCount = Me.ColumnCount - ColumnStartIndex
    
    iRow = CreateIndexHeaders(ColumnStartIndex, ColumnsCount, ColumnIndexStartsOnZero)
    cMax = UBound(iRow)
    cMin = LBound(iRow)
    
    iMax = ((PageSize * (PageIndex + 1)) - 1) + ((PageCount - 1) * PageSize)
    If iMax > Me.Count - 1 Then iMax = Me.Count - 1
        
    If ColumnStartIndex <> 0 Or ColumnsCount <> Me.ColumnCount Then
        Set AllValues = Me.RecordList.GetRange(PageSize * PageIndex, iMax - (PageSize * PageIndex) + 1, CreateIndexHeaders(ColumnStartIndex, ColumnsCount, False))
    Else
        Set AllValues = Me.RecordList
    End If
    
    With rs
        For c = cMin To cMax
            .Fields.Append CStr(iRow(c)), adLongVarWChar, -1, adFldIsNullable
        Next c
        
        .Open
        For i = 0 To iMax - (PageSize * PageIndex)
            .AddNew FieldList:=iRow, Values:=AllValues(i)
        Next i
        
        If Me.Count > 0 Then .MoveFirst
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

' DEPRECATED, use GetColumnNames() instead.
Public Function GetHeaderNames() As Variant
    Dim t As Variant, c As Long, d As Scripting.Dictionary, hItems As Variant
    ReDim t(0 To Me.ColumnCount - 1)
    
    hItems = Me.HeaderList(0)
    
    For c = 0 To UBound(hItems)
        Set d = hItems(c)
        t(c) = CStr(d("ColumnName"))
    Next c
    
    GetHeaderNames = t
End Function

Public Function GetColumnNames(Optional ByVal FirstColumnIndex As Long = 0, Optional ByVal GetCount As Variant) As Variant
    Dim t As Variant, c As Long, hItems As Variant, sLimit As Long
    
    If IsMissing(GetCount) Then GetCount = Me.ColumnCount - FirstColumnIndex
    ReDim t(0 To GetCount - 1)
    hItems = Me.HeaderList(0)
    sLimit = IIf(Me.ColumnCount - FirstColumnIndex < GetCount, Me.ColumnCount - FirstColumnIndex, GetCount)
    
    For c = 0 To sLimit - 1
        t(c) = CStr(hItems(c + FirstColumnIndex)("ColumnName"))
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
        Set t(0, i) = CreateColumnHeaderFromRecordsetField(AllFields(i))
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
                Set t(0, i) = CreateColumnHeader(cNames(i))
            Case "0", "#"
                Set t(0, i) = CreateColumnHeader(cNames(i), NumberFormat:=cFormats(i)(0), DataType:="Numeric")
            Case "@"
                Set t(0, i) = CreateColumnHeader(cNames(i), NumberFormat:=cFormats(i)(0), DataType:="Text")
            Case "D", "M", "Y", "H"
                Set t(0, i) = CreateColumnHeader(cNames(i), NumberFormat:=cFormats(i)(0), DataType:="Date")
            Case Else
                Set t(0, i) = CreateColumnHeader(cNames(i), NumberFormat:=cFormats(i)(0))
        End Select
    Next i

    Set CreateColumnHeadersFromExcelRange = New Array2dEx
    CreateColumnHeadersFromExcelRange.Instance = CollectionsLib.GetArrayByRef(t)
End Function


' --- TYPE CASTING ---

Private Function C2Recordset(ByVal Value As Variant) As ADODB.Recordset: Set C2Recordset = Value: End Function
Private Function C2RecordsetEx(ByVal Value As Variant) As RecordsetEx: Set C2RecordsetEx = Value: End Function
Private Function C2ExcelRange(ByVal Value As Variant) As Excel.Range: Set C2ExcelRange = Value: End Function
Private Function C2ExcelWorksheet(ByVal Value As Variant) As Excel.Worksheet: Set C2ExcelWorksheet = Value: End Function
Private Function C2xlSheetsEx(ByVal Value As Variant) As xlSheetsEx: Set C2xlSheetsEx = Value: End Function
Private Function C2Dictionary(ByVal Value As Variant) As Scripting.Dictionary: Set C2Dictionary = Value: End Function


' --- ICollectionEx ---

Private Property Get ICollectionEx_Count() As Long: ICollectionEx_Count = Count: End Property
Private Property Get ICollectionEx_ColumnCount() As Long: ICollectionEx_ColumnCount = ColumnCount: End Property
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




