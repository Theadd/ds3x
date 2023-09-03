VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArraySliceGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
Option Compare Database
Option Explicit
Option Base 0


Implements ICollectionEx

Private pGroups As ArrayListEx


Public Property Get Count() As Long
    On Error Resume Next
    Count = pGroups(0).Count
End Property

Public Property Get ColumnCount() As Long
    On Error Resume Next
    ColumnCount = pGroups.Count
End Property

Public Property Get Row(ByVal Index As Long) As Variant
    ' TODO
End Property

Public Property Get SliceGroups() As Variant
    If pGroups Is Nothing Then
        SliceGroups = Array()
    Else
        SliceGroups = pGroups.ToArray()
    End If
End Property

Public Property Let SliceGroups(ByRef Value As Variant)
    Dim v As Variant, r As Long
    If Not pGroups Is Nothing Then Err.Raise 70
    
    Set pGroups = ArrayListEx.Create(Value)
End Property

Public Property Get SliceAt(ByVal ColumnIndex As Long) As ArraySlice
    Set SliceAt = pGroups(ColumnIndex)
End Property



Public Function CreateBlank(ByVal RowsCount As Long, ByVal ColumnsCount As Long) As ArraySliceGroup: Set CreateBlank = Create(Array2dEx.CreateBlank(RowsCount, ColumnsCount)): End Function

Public Function Create(Optional ByRef ArrayLike As Variant) As ArraySliceGroup
    Set Create = New ArraySliceGroup
    If Not IsMissing(ArrayLike) Then
        Create.SliceGroups = CreateSliceGroups(ArrayLike).ToArray()
    End If
End Function



Public Function AddRange(ByRef Target As Variant) As ArraySliceGroup
    Dim t As ArraySliceGroup, c As Long, aX As ArrayListEx
    
    If IsObject(Target) Then
        If TypeOf Target Is ArraySliceGroup Then Set t = C2ArraySliceGroup(Target)
    End If
    If t Is Nothing Then Set t = Create(Target)
    Set aX = ArrayListEx.Create()
    
    For c = 0 To pGroups.Count - 1
        aX.Add C2ArraySlice(pGroups(c)).Add(t.SliceAt(c))
    Next c
    
    Set AddRange = New ArraySliceGroup
    AddRange.SliceGroups = aX.ToArray()
End Function

Public Function GetRange(Optional ByVal Index As Variant, Optional ByVal GetCount As Variant, Optional ByVal ColumnIndexes As Variant) As ArraySliceGroup
    Dim c As Variant, aX As ArrayListEx
    
    If IsMissing(Index) Or IsEmpty(Index) Then Index = 0
    If Index >= Count Then
        Index = IIf(Count - 1 < 0, 0, Count - 1)
        GetCount = 0
    Else
        If IsMissing(GetCount) Or IsEmpty(GetCount) Then GetCount = Count - Index
        If GetCount + Index > Count Then GetCount = Count - Index
    End If
    If IsMissing(ColumnIndexes) Or IsEmpty(ColumnIndexes) Then
        ColumnIndexes = CollectionsLib.ArrayRange(0, ColumnCount - 1)
    Else
        If Not IsArray(ColumnIndexes) Then ColumnIndexes = Array(ColumnIndexes)
    End If
    Set aX = ArrayListEx.Create()
    
    For Each c In ColumnIndexes
        aX.Add C2ArraySlice(pGroups(CLng(c))).Slice(Index, GetCount)
    Next c
    
    Set GetRange = New ArraySliceGroup
    GetRange.SliceGroups = aX.ToArray()
End Function

' Concatenate all slices from Target as additional slices into a new ArraySliceGroup
Public Function Join(ByVal Target As ArraySliceGroup) As ArraySliceGroup
    Set Join = New ArraySliceGroup
    Join.SliceGroups = ArrayListEx.Create(SliceGroups).AddRange(Target.SliceGroups).ToArray()
End Function


' Returns an ArraySliceGroup containing only those rows specified in the RowIndexes array.
' Can also be used for: Reordering, cutting rows, duplicate, etc.
'
' @param RowIndexes - An array of row indices.
' @param [ModIndex] - A signed integer to shift values in the RowIndexes array.
Public Function GetRows(ByRef RowIndexes As Variant, Optional ByVal ModIndex As Long = 0) As ArraySliceGroup
    Dim aX As ArrayListEx, vSlice As Variant, nRows As Long, i As Long, gX As ArrayListEx
    Dim r As Long, rMax As Long, gStartIndex As Long, isValid As Boolean, wasValid As Boolean, isSequence As Boolean, iLast As Long
    nRows = Count: iLast = -16
    rMax = ArrayListEx.CountElementsIn(RowIndexes) - 1
    Set aX = ArrayListEx.Create()
    
    For r = 0 To rMax
        wasValid = isValid
        i = CLng(RowIndexes(r)) + ModIndex
        isSequence = (r = 0 Or i - iLast = 1)
        isValid = (i >= 0 And i < nRows)
        
        If wasValid Then
            If (Not isValid Or Not isSequence) Then
                aX.Add Array(gStartIndex, iLast - gStartIndex + 1)
                isValid = False
            ElseIf r = rMax Then
                aX.Add Array(gStartIndex, i - gStartIndex + 1)
            End If
        ElseIf isValid And Not wasValid Then
            gStartIndex = i
        End If
        
        iLast = i
    Next r
    
    Set gX = ArrayListEx.Create()
    For Each vSlice In pGroups
        gX.Add C2ArraySlice(vSlice).SliceAllRanges(aX)
    Next vSlice
    
    Set GetRows = New ArraySliceGroup
    GetRows.SliceGroups = gX.ToArray()
End Function

Public Function CopyToRange(ByRef RangeObject As Range, Optional ByVal ApplyUserLocale As Boolean = True) As ArraySliceGroup
    Dim nRows As Long, c As Long, nCols As Long
    Set CopyToRange = Me
    nRows = Count
    nCols = ColumnCount
    If Not nRows > 0 Then Exit Function
    
    On Error GoTo UserLocaleFail
    
    For c = 0 To nCols - 1

        With RangeObject.Resize(nRows, 1).Offset(ColumnOffset:=c + 1)
            If ApplyUserLocale Then
                
                .FormulaR1C1Local = C2ArraySlice(pGroups(c)).Instance
            Else
                .Value = C2ArraySlice(pGroups(c)).Instance
            End If
        End With
        
    Next c
    
    Exit Function
UserLocaleFail:
    Debug.Print "[ERROR] @ArraySliceGroup.CopyToRange() -> " & Err.Description
    Resume Next
End Function

' Practically there's no point in cloning an immutable object, it's just exposed to share some common API with the other collection class modules.
Public Function Clone() As ArraySliceGroup
    Set Clone = New ArraySliceGroup
    Clone.SliceGroups = SliceGroups
End Function

Public Function ToCSV(Optional ByVal Delimiter As String = ",") As String
'    Dim t() As String, s() As String, i As Long, c As Long, uc As Long
'    uc = Me.ColumnCount - 1
'    ReDim t(0 To Me.Count)
'    ReDim s(0 To uc)
'
'    For i = 0 To Me.Count - 1
'        For c = 0 To uc
'            s(c) = ValueToCSV(Instance(i, c))
'        Next c
'
'        t(i) = VBA.Join(s, Delimiter)
'    Next i
'
'    t(Me.Count) = ""
'    ToCSV = VBA.Join(t, vbNewLine)
End Function

Public Function ToJSON() As String: ToJSON = JSON.Stringify(Me, 2): End Function

Public Function ToExcel() As String: ToExcel = ToCSV(vbTab): End Function



Private Function CreateSliceGroups(ByRef ArrayLike As Variant) As ArrayListEx
    Dim aX As ArrayListEx, nRows As Long, c As Long
    
    If IsObject(ArrayLike) Then
        If TypeOf ArrayLike Is ICollectionEx Then
            Set aX = ArrayListEx.Create()
            nRows = C2ICollectionEx(ArrayLike).Count
            For c = 0 To C2ICollectionEx(ArrayLike).ColumnCount - 1
                aX.Add ArraySlice.Create(ArrayLike, Index:=0, GetCount:=nRows, ColumnIndex:=c)
            Next c
            Set CreateSliceGroups = aX
            Exit Function
        End If
    End If
    Set CreateSliceGroups = CreateSliceGroups(Array2dEx.Create(ArrayLike))
End Function


Public Function Bind(Optional ByRef ArrayLike As Variant) As ArraySliceGroup
    Set Bind = Me
    Err.Raise 17, , "Can't perform requested operation. Binding to an immutable object is not allowed."
End Function

Public Function Unbind() As ArraySliceGroup
    Set Unbind = Me
    Err.Raise 17, , "Can't perform requested operation."
End Function

Public Sub Dispose()

End Sub

Private Function C2ArraySlice(ByRef Value As Variant) As ArraySlice: Set C2ArraySlice = Value: End Function
Private Function C2Array2dEx(ByRef Value As Variant) As Array2dEx: Set C2Array2dEx = Value: End Function
Private Function C2RecordsetEx(ByRef Value As Variant) As RecordsetEx: Set C2RecordsetEx = Value: End Function
Private Function C2ICollectionEx(ByRef Value As Variant) As ICollectionEx: Set C2ICollectionEx = Value: End Function
Private Function C2ArraySliceGroup(ByRef Value As Variant) As ArraySliceGroup: Set C2ArraySliceGroup = Value: End Function


' --- ICollectionEx ---

Private Property Get ICollectionEx_Count() As Long: ICollectionEx_Count = Count: End Property
Private Property Get ICollectionEx_ColumnCount() As Long: ICollectionEx_ColumnCount = ColumnCount: End Property
Private Property Get ICollectionEx_Row(ByVal Index As Long) As Variant: ICollectionEx_Row = Row(Index): End Property
Private Function ICollectionEx_CreateBlank(ByVal RowsCount As Long, ByVal ColumnsCount As Long) As ICollectionEx: Set ICollectionEx_CreateBlank = CreateBlank(RowsCount, ColumnsCount): End Function
Private Function ICollectionEx_Create(Optional ByRef FromTarget As Variant) As ICollectionEx: Set ICollectionEx_Create = Create(FromTarget): End Function
Private Function ICollectionEx_Bind(Optional ByRef Target As Variant) As ICollectionEx: Set ICollectionEx_Bind = Bind(Target): End Function
Private Function ICollectionEx_Unbind() As ICollectionEx: Set ICollectionEx_Unbind = Unbind(): End Function
Private Sub ICollectionEx_Dispose(): Dispose: End Sub
Private Function ICollectionEx_Join(ByVal Target As ICollectionEx) As ICollectionEx: Set ICollectionEx_Join = Join(Target): End Function
Private Function ICollectionEx_AddRange(ByVal Target As ICollectionEx) As ICollectionEx: Set ICollectionEx_AddRange = AddRange(Target): End Function
Private Function ICollectionEx_GetRange(Optional ByVal Index As Variant, Optional ByVal GetCount As Variant, Optional ByVal ColumnIndexes As Variant) As ICollectionEx: Set ICollectionEx_GetRange = GetRange(Index, GetCount, ColumnIndexes): End Function
Private Function ICollectionEx_ToCSV(Optional ByVal Delimiter As String = ",") As String: ICollectionEx_ToCSV = ToCSV(Delimiter): End Function
Private Function ICollectionEx_ToJSON() As String: ICollectionEx_ToJSON = ToJSON(): End Function
Private Function ICollectionEx_ToExcel() As String: ICollectionEx_ToExcel = ToExcel(): End Function
Private Function ICollectionEx_CopyToRange(ByRef RangeObject As Excel.Range, Optional ByVal ApplyUserLocale As Boolean = True, Optional ByVal WriteHeaders As Boolean = True) As ICollectionEx: Set ICollectionEx_CopyToRange = CopyToRange(RangeObject, ApplyUserLocale): End Function





