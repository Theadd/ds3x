VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "dsTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Option Base 0


Private pHeaderList As ArrayListEx
Private pHeaderArray As Array2dEx
Private pRecordList As ArrayListEx
Private pRecordArray As Array2dEx
Private pTableSource As Variant



Public Property Get Count() As Long
    On Error Resume Next
    If pRecordArray Is Nothing Then
        Count = Me.RecordList.Count
    Else
        Count = Me.RecordArray.Count
    End If
End Property

Public Property Get ColumnCount() As Long
    On Error Resume Next
    If pHeaderArray Is Nothing Then
        ColumnCount = 1 + UBound(Me.HeaderList.Item(0)) - LBound(Me.HeaderList.Item(0))
    Else
        ColumnCount = Me.HeaderArray.ColumnCount
    End If
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

Public Property Get IndexRecordset() As ADODB.Recordset: Set IndexRecordset = CreateIndexRecordset: End Property
Public Property Get NamedRecordset() As ADODB.Recordset: Set NamedRecordset = CreateNamedRecordset: End Property
Public Property Get RowNumbers() As Boolean: On Error Resume Next: RowNumbers = (GetHeaderNames()(0) = "${#}"): End Property



Public Function Create(Optional ByRef TableLike As Variant) As dsTable
    Set Create = New dsTable
    If Not IsMissing(TableLike) Then Create.Bind TableLike
End Function

Public Function Bind(Optional ByRef TableLike As Variant) As dsTable
    Set Bind = Unbind
    
    If Not IsMissing(TableLike) Then
        If Not IsObject(TableLike) Then
            Err.Raise 13
        Else
            If TypeOf TableLike Is RecordsetEx Then
                Set pTableSource = TableLike    '.Clone
            ElseIf TypeOf TableLike Is ArrayListEx Then
                Set pRecordList = TableLike
            ElseIf TypeOf TableLike Is Array2dEx Then
                Set pRecordArray = TableLike
            Else
                Err.Raise 13
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
    ' TODO
End Sub

Public Function SetHeaders(ByRef AllHeaders As Variant) As dsTable
    Dim colHeaders As ArrayListEx, i As Long, Aux As Variant
    If Not (pHeaderList Is Nothing And pHeaderArray Is Nothing) Then
        Err.Raise 13
    Else
        If IsObject(AllHeaders) Then
            If TypeOf AllHeaders Is RecordsetEx Then
                Set pHeaderArray = CreateColumnHeadersFromRecordsetFields(AllHeaders.Fields)
            ElseIf TypeOf AllHeaders Is ArrayListEx Then
                Set pHeaderList = AllHeaders
            ElseIf TypeOf AllHeaders Is Array2dEx Then
                Set pHeaderArray = AllHeaders
            Else
                Err.Raise 13
            End If
        ElseIf IsArray(AllHeaders) Then
            If UBound(AllHeaders) >= 0 Then
                ' Asuming a plain-text array of column names
                Set colHeaders = New ArrayListEx
                For i = 0 To UBound(AllHeaders)
                    colHeaders.Add CreateColumnHeader(CStr(AllHeaders(i)))
                Next i
                Aux = colHeaders.ToArray
                colHeaders.Clear
                colHeaders.Add Aux
                Set pHeaderList = colHeaders
            End If
        End If
    End If
    
    Set SetHeaders = Me
End Function

Private Function CreateColumnHeadersFromRecordsetFields(ByVal AllFields As ADODB.Fields) As Array2dEx
    Dim i As Long, t() As Variant
    ReDim t(0 To 0, 0 To AllFields.Count - 1)
    
    For i = 0 To AllFields.Count - 1
        Set t(0, i) = CreateColumnHeaderFromRecordsetField(AllFields(i))
    Next i

    Set CreateColumnHeadersFromRecordsetFields = New Array2dEx
    CreateColumnHeadersFromRecordsetFields.Instance = CollectionsLib.GetArrayByRef(t)
End Function



' Concatenate all elements of dsTable as additional columns into a new dsTable
Public Function Join(ByVal TargetTable As dsTable) As dsTable
    Set Join = Me _
                .Create(Me.RecordArray.Join(TargetTable.RecordArray)) _
                .SetHeaders(Me.HeaderArray.Join(TargetTable.HeaderArray))
End Function


'
'
' @SEEALSO: ArrayListEx.GetRange([Index], [Count], [ColumnIndexes]) As ArrayListEx
Public Function GetRange(Optional ByVal Index As Variant, Optional ByVal GetCount As Variant, Optional ByVal ColumnIndexes As Variant) As dsTable

    Set GetRange = Me.Create(Me.Records.GetRange(Index, GetCount, ColumnIndexes)) _
                     .SetHeaders(Me.Headers.GetRange(0, Me.Headers.Count, ColumnIndexes))
End Function

'Public Function Resize(Optional ByVal NumRows As Variant, Optional ByVal NumColumns As Variant, Optional ByVal DefaultValue As Variant) As Array2dEx
'    Dim t As New Array2dEx
'End Function

Public Function ToCSV(Optional ByVal Delimiter As String = ",") As String
    ToCSV = GetHeaderNames.ToCSV(Delimiter) & Me.Records.ToCSV(Delimiter)
End Function

Public Function CopyToRange(ByRef RangeObject As Range, Optional ByVal ApplyUserLocale As Boolean = True) As dsTable
    Dim hList As New ArrayListEx
    
    hList.Add GetHeaderNames
    Array2dEx.Create(hList).CopyToRange RangeObject, False
    Me.RecordArray.CopyToRange RangeObject.Range("A2"), ApplyUserLocale
    
    Set CopyToRange = Me
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

Private Sub Class_Initialize()
    
End Sub

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


Private Function CreateNamedRecordset() As ADODB.Recordset
    Dim rs As New ADODB.Recordset, i As Long, c As Long, cMax As Long, cMin As Long, iRow As Variant, allValues As ArrayListEx
    
'    If Me.Count = 0 Then
'        Set CreateNamedRecordset = rs
'        Exit Function
'    End If
    
    iRow = GetHeaderNames
    cMax = UBound(iRow)
    cMin = LBound(iRow)
    Set allValues = Me.RecordList
    
    With rs
        For c = cMin To cMax
            .Fields.Append CStr(iRow(c)), adLongVarWChar, -1, adFldIsNullable
        Next c
        
        .Open
        
        For i = 0 To Me.Count - 1
            .AddNew FieldList:=iRow, Values:=allValues(i)
        Next i
        
        If Me.Count > 0 Then .MoveFirst
    End With
    Set CreateNamedRecordset = rs
End Function

Private Function CreateIndexRecordset() As ADODB.Recordset
    Dim rs As New ADODB.Recordset, i As Long, c As Long, cMax As Long, cMin As Long, iRow As Variant, allValues As ArrayListEx
    
    If Me.Count = 0 Then
        Set CreateIndexRecordset = rs
        Exit Function
    End If
    
    iRow = CreateIndexHeaders
    cMax = UBound(iRow)
    cMin = LBound(iRow)
    Set allValues = Me.RecordList
    
    With rs
        For c = cMin To cMax
            .Fields.Append CStr(iRow(c)), adLongVarWChar, -1, adFldIsNullable
        Next c
        
        .Open
        
        For i = 0 To Me.Count - 1
            .AddNew FieldList:=iRow, Values:=allValues(i)
        Next i
        
        .MoveFirst
    End With
    Set CreateIndexRecordset = rs
End Function

Private Function CreateIndexHeaders() As Variant
    Dim t As Variant, c As Long, cMax As Long
    cMax = Me.ColumnCount - 1
    ReDim t(0 To cMax)
    
    For c = 0 To cMax
        t(c) = CStr(c)
    Next c
    
    CreateIndexHeaders = t
End Function

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



