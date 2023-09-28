VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DictionaryEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder "ds3x.Collections"
Option Compare Database
Option Explicit

Implements ICollectionEx

Public Instance As Scripting.Dictionary


'Gets the key count of the underlying Scripting.Dictionary.
Public Property Get Count() As Variant
    Count = Instance.Count
End Property

Public Property Get ColumnCount() As Long: ColumnCount = 2: End Property

'Wrapper property that returns the value for the specified key if it exists.
'If it doesn't exist, it returns the default rather than raising an error.
Public Property Get GetValue(mKey As Variant, Optional ByRef mDefaultVal As Variant) As Variant
    If Instance.Exists(mKey) Then
        Assign GetValue, Instance.Item(mKey)
    Else
        If VarType(mKey) = vbString Then
            Assign GetValue, GetNestedItem(CStr(mKey))
        End If
        If IsEmpty(GetValue) Then
            If Not IsMissing(mDefaultVal) Then Assign GetValue, mDefaultVal
        End If
    End If
End Property

'Sets or returns an item for a specified key.
'@DefaultMember
Public Property Get Item(mKey As Variant) As Variant
Attribute Item.VB_UserMemId = 0
    If Instance.Exists(mKey) Then
        Assign Item, Instance.Item(mKey)
    Else
        If VarType(mKey) = vbString Then
            Assign Item, GetNestedItem(CStr(mKey))
        End If
    End If
End Property

Public Property Let Item(mKey As Variant, mVal As Variant)
    If VarType(mKey) = vbString Then
        If TrySetNestedItem(mKey, mVal) Then Exit Property
    End If
    Instance.Item(mKey) = mVal
End Property

Public Property Set Item(mKey As Variant, mVal As Variant)
    If VarType(mKey) = vbString Then
        If TrySetNestedItem(mKey, mVal) Then Exit Property
    End If
    Set Instance.Item(mKey) = mVal
End Property

'Updates a key, replacing the old key with the new key.
Public Property Let Key(mKey As Variant, mNewKey As Variant)
    Instance.Key(mKey) = mNewKey
End Property

Public Property Get Row(ByVal Index As Long) As Variant
    Dim dKey As Variant
    dKey = Instance.Keys()(Index)
    Row = Array(dKey, Instance.Item(dKey))
End Property


' --- PUBLIC METHODS ---

Public Function SetValue(mKey As Variant, ByRef Value As Variant) As DictionaryEx
    Set SetValue = Me
    If IsObject(Value) Then
        Set Item(mKey) = Value
    Else
        Item(mKey) = Value
    End If
End Function

Public Function CreateBlank(ByVal RowsCount As Long, ByVal ColumnsCount As Long) As DictionaryEx: Set CreateBlank = Create(): End Function

Public Function Create(Optional ByRef DictionaryLike As Variant) As DictionaryEx
    Set Create = New DictionaryEx
    
    If Not IsMissing(DictionaryLike) Then
        If IsObject(DictionaryLike) Then
            If Not DictionaryLike Is Nothing Then
                Create.Bind DictionaryLike
                Exit Function
            End If
        Else
            If VarType(DictionaryLike) = vbString Then
                If VBA.Left(LTrim(DictionaryLike), 1) = "{" Then
                    Create.Bind JSON.Parse(DictionaryLike, True, True)
                    Exit Function
                End If
            ElseIf IsArray(DictionaryLike) Then
                Create.Bind ArrayListEx.Create(DictionaryLike)
                Exit Function
'                If Array2dEx.Is2dArray(DictionaryLike) Then
'                    CreateFromIterable ArrayListEx.Create(DictionaryLike).ToArray()
'                Else
'                    CreateFromIterable DictionaryLike
'                End If
            End If
        End If
    End If
    
    Dim t As Scripting.Dictionary
    Set t = New Scripting.Dictionary
    Create.Bind t
End Function

Public Function Bind(ByRef DictionaryLike As Variant) As DictionaryEx
    Set Bind = Me
    If TypeOf DictionaryLike Is DictionaryEx Then
        Set Instance = DictionaryLike.Instance
    ElseIf TypeOf DictionaryLike Is Scripting.Dictionary Then
        Set Instance = DictionaryLike
    ElseIf TypeOf DictionaryLike Is ArrayListEx Then
        CreateFromIterable C2ArrayListEx(DictionaryLike).ToArray()
    ElseIf TypeOf DictionaryLike Is Array2dEx Then
        CreateFromIterable ArrayListEx.Create(DictionaryLike).ToArray()
    ElseIf TypeOf DictionaryLike Is ArrayList Then
        CreateFromIterable C2ArrayList(DictionaryLike).ToArray()
    End If
End Function

Public Function Unbind() As DictionaryEx
    Set Instance = Nothing: Set Unbind = Me
End Function

'Adds a key and item pair to the Scripting.Dictionary
Public Function Add(mKey As Variant, mVal As Variant) As DictionaryEx
    On Error Resume Next
    Instance.Add mKey, mVal

    If Err = 457 Then
        If IsObject(mVal) Then
            Set Instance.Item(mKey) = mVal
        Else
            Instance.Item(mKey) = mVal
        End If
    End If
    
    Set Add = Me
    On Error GoTo 0
End Function

Public Function AddRange(ByVal Target As Variant) As DictionaryEx
    Set AddRange = Append(Target)
End Function

Public Function GetRange(Optional ByVal Index As Variant, Optional ByVal GetCount As Variant) As DictionaryEx
    Set GetRange = DictionaryEx.Create(ArrayListEx.Create(Entries()).GetRange(Index, GetCount))
End Function

Public Function CopyToRange(ByRef RangeObject As Excel.Range, Optional ByVal ApplyUserLocale As Boolean = True, Optional ByVal WriteHeaders As Boolean = True) As DictionaryEx
    Set CopyToRange = Me
    dsTable.Create(ArrayListEx.Create(Entries()), False).SetHeaders(Array("Key", "Value")).CopyToRange RangeObject, ApplyUserLocale, WriteHeaders
End Function

'Appends another dictionary (Scripting.Dictionary or DictionaryEx) to the Scripting.Dictionary
Public Function Append(Optional ByRef mOther As Variant) As DictionaryEx
    Dim mKey As Variant
    
    If Not IsMissing(mOther) Then
        For Each mKey In mOther.Keys()
            Add mKey, mOther.Item(mKey)
        Next mKey
    End If
    
    Set Append = Me
End Function

Public Function Join(ByVal Target As RecordsetEx) As DictionaryEx
    Set Join = DictionaryEx.Create()
    ' TODO: Implement
    Err.Raise 17, , "Not implemented."
End Function

' Removes a key-value pair from the Scripting.Dictionary.
Public Function Remove(mKey As Variant) As DictionaryEx
    Instance.Remove mKey
    Set Remove = Me
End Function

' Removes all key-value pairs from the Scripting.Dictionary.
Public Function RemoveAll() As DictionaryEx
    Instance.RemoveAll
    Set RemoveAll = Me
End Function

' Removes all key-value pairs from the Scripting.Dictionary.
Public Function Clear() As DictionaryEx
    Instance.RemoveAll
    Set Clear = Me
End Function

'Returns True if a specified key exists in the Scripting.Dictionary, false otherwise.
Public Function Exists(Key As Variant) As Boolean
    Exists = Instance.Exists(Key)
End Function

'Returns an array containing all the items in the Scripting.Dictionary.
Public Function Items() As Variant()
    Items = Instance.Items
End Function

'Returns an array containing all existing keys in the Scripting.Dictionary.
Public Function Keys() As Variant()
    Keys = Instance.Keys
End Function

'Returns an iterable containing all the Key -> Value pair arrays in the Scripting.Dictionary.
Public Function Entries() As Variant
    Dim aX As ArrayListEx, dKey As Variant
    Set aX = ArrayListEx.Create()
    
    For Each dKey In Instance.Keys()
        aX.Add Array(dKey, Instance(dKey))
    Next dKey
    
    Entries = aX.ToArray()
End Function

Public Function Clone() As DictionaryEx
    Dim t As New DictionaryEx
    Set t.Instance = ShallowCopyDictionary(Instance)
    
    Set Clone = t
End Function

Public Function Duplicate() As DictionaryEx
    Set Duplicate = Unserialize(Serialize())
End Function

Public Function Serialize() As Variant
    Serialize = JSON.Stringify(Instance)
End Function

Public Function Unserialize(ByRef Serialized As Variant) As DictionaryEx
    Set Unserialize = New DictionaryEx
    Set Unserialize.Instance = New Scripting.Dictionary
    Set Unserialize.Instance = JSON.Parse(CStr(Serialized), True, True)
End Function

Public Function ToCSV(Optional ByVal Delimiter As String = ",", Optional ByVal InLocalFormat As Boolean = False) As String
    ToCSV = dsTable.Create(ArrayListEx.Create(Entries()), False).SetHeaders(Array("Key", "Value")).ToCSV(Delimiter, InLocalFormat)
End Function

Public Function ToJSON() As String: ToJSON = JSON.Stringify(Me, 2): End Function

Public Function ToExcel() As String: ToExcel = ToCSV(vbTab, True): End Function

Public Function ToString(Optional mPrettyIndent As Variant) As String
    If IsMissing(mPrettyIndent) Then
        ToString = JSON.Stringify(Instance)
    Else
        ToString = JSON.Stringify(Instance, mPrettyIndent)
    End If
End Function

Public Function PrintAll() As DictionaryEx
    Debug.Print JSON.Stringify(Instance, 2)
    Set PrintAll = Me
End Function


' --- PRIVATE ---

Private Function ShallowCopyDictionary(ByVal d As Scripting.Dictionary) As Scripting.Dictionary
    Dim s As New Scripting.Dictionary, k As Variant
    
    s.CompareMode = d.CompareMode
    For Each k In d.Keys
        s.Add k, d(k)
    Next k
    
    Set ShallowCopyDictionary = s
End Function

Private Function TrySetNestedItem(ByVal mKey As String, ByRef Value As Variant) As Boolean
    Dim t As Long, v As Variant, sAux As String
    
    t = InStrRev(mKey, ".")
    If t = 0 Then Exit Function
    
    Assign v, Item(VBA.Left$(mKey, t - 1))
    sAux = VBA.Mid$(mKey, t + 1, Len(mKey))
    If IsObject(Value) Then
        Set v(sAux) = Value
    Else
        v(sAux) = Value
    End If
    TrySetNestedItem = True
End Function

Private Function GetNestedItem(ByVal mKey As String) As Variant
    Dim t As Long, c As Long, c2 As Long, v As Variant, v2 As Variant, tEnd As Boolean, cEnd As Boolean
    On Error GoTo Finally
    
    Do
        t = InStr(1 + t, mKey, ".", vbBinaryCompare)
        c = InStr(1 + c, mKey, "[", vbBinaryCompare)
        If t = 0 Then tEnd = True
        If c = 0 Then cEnd = True
        If Not cEnd Then
            If Instance.Exists(VBA.Mid$(mKey, 1, c - 1)) Then
                c2 = InStr(1 + c, mKey, "]", vbBinaryCompare)
                If c2 <> 0 Then
                    v = VBA.Mid$(mKey, 1 + c, c2 - c - 1)
                    If CStr(Val(v)) = CStr(v) Then
                        c2 = CLng(Val(v))
                        Assign v2, Instance.Item(VBA.Mid$(mKey, 1, c - 1))
                        Assign v, v2(c2)
                        If Len(mKey) = (c + Len(CStr(c2)) + 1) Then
                            Assign GetNestedItem, v
                        Else
                            v2 = VBA.Mid$(mKey, 1 + c + Len(CStr(c2)) + 1, Len(mKey))
                            v2 = IIf(VBA.Left$(v2, 1) = ".", VBA.Mid$(v2, 2, Len(v2)), v2)
                            Assign GetNestedItem, ContinueNestedItem(v, v2)
                        End If
                        Exit Do
                    End If
                End If
            End If
        End If
        If Not tEnd Then
            If Instance.Exists(VBA.Mid$(mKey, 1, t - 1)) Then
                Assign GetNestedItem, ContinueNestedItem(Instance.Item(VBA.Mid$(mKey, 1, t - 1)), VBA.Mid$(mKey, 1 + t, Len(mKey)))
                Exit Do
            End If
        End If
    Loop Until (tEnd And cEnd)
Finally:
End Function

Private Function ContinueNestedItem(ByRef NestedItem As Variant, ByVal mKey As String) As Variant
    On Error GoTo Finally
    If IsObject(NestedItem) Then
        If TypeOf NestedItem Is Scripting.Dictionary Then
            Assign ContinueNestedItem, DictionaryEx.Create(NestedItem)(mKey)
        Else
            Assign ContinueNestedItem, NestedItem(mKey)
        End If
    Else
        Assign ContinueNestedItem, NestedItem(mKey)
    End If
Finally:
End Function

Private Sub Assign(ByRef Target, ByRef Value As Variant)
    If IsObject(Value) Then
        Set Target = Value
    Else
        Target = Value
    End If
End Sub

Private Sub CreateFromIterable(ByRef Target As Variant)
    Dim i As Long
    Set Instance = New Scripting.Dictionary
    
    If IsArray(Target(0)) Then
        For i = 0 To UBound(Target)
            Add Target(i)(0), Target(i)(1)
        Next i
    Else
        For i = 0 To UBound(Target) Step 2
            Add Target(i), Target(i + 1)
        Next i
    End If
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Set Instance = Nothing
    On Error GoTo 0
End Sub

Public Sub Dispose(): Set Instance = Nothing: End Sub


' --- TYPE CASTING ---

Private Function C2ArrayListEx(ByRef Value As Variant) As ArrayListEx: Set C2ArrayListEx = Value: End Function
Private Function C2ArrayList(ByRef Value As Variant) As ArrayList: Set C2ArrayList = Value: End Function


' --- ICollectionEx ---

Private Property Get ICollectionEx_Count() As Long: ICollectionEx_Count = Count: End Property
Private Property Get ICollectionEx_ColumnCount() As Long: ICollectionEx_ColumnCount = ColumnCount: End Property
Private Property Get ICollectionEx_Item(ByVal Index As Long, Optional ByVal ColumnIndex As Long = 0&) As Variant: ICollectionEx_Item = Row(Index)(ColumnIndex): End Property
Private Property Get ICollectionEx_Row(ByVal Index As Long) As Variant: ICollectionEx_Row = Row(Index): End Property
Private Function ICollectionEx_CreateBlank(ByVal RowsCount As Long, ByVal ColumnsCount As Long) As ICollectionEx: Set ICollectionEx_CreateBlank = CreateBlank(RowsCount, ColumnsCount): End Function
Private Function ICollectionEx_Create(Optional ByRef FromTarget As Variant) As ICollectionEx: Set ICollectionEx_Create = Create(FromTarget): End Function
Private Function ICollectionEx_Bind(Optional ByRef Target As Variant) As ICollectionEx: Set ICollectionEx_Bind = Bind(Target): End Function
Private Function ICollectionEx_Unbind() As ICollectionEx: Set ICollectionEx_Unbind = Unbind(): End Function
Private Sub ICollectionEx_Dispose(): Dispose: End Sub
Private Function ICollectionEx_Join(ByVal Target As ICollectionEx) As ICollectionEx: Set ICollectionEx_Join = Join(Target): End Function
Private Function ICollectionEx_AddRange(ByVal Target As ICollectionEx) As ICollectionEx: Set ICollectionEx_AddRange = AddRange(Target): End Function
Private Function ICollectionEx_GetRange(Optional ByVal Index As Variant, Optional ByVal GetCount As Variant, Optional ByVal ColumnIndexes As Variant) As ICollectionEx: Set ICollectionEx_GetRange = GetRange(Index, GetCount): End Function
Private Function ICollectionEx_ToCSV(Optional ByVal Delimiter As String = ",", Optional ByVal InLocalFormat As Boolean = False) As String: ICollectionEx_ToCSV = ToCSV(Delimiter, InLocalFormat): End Function
Private Function ICollectionEx_ToJSON() As String: ICollectionEx_ToJSON = ToJSON(): End Function
Private Function ICollectionEx_ToExcel() As String: ICollectionEx_ToExcel = ToExcel(): End Function
Private Function ICollectionEx_CopyToRange(ByRef RangeObject As Excel.Range, Optional ByVal ApplyUserLocale As Boolean = True, Optional ByVal WriteHeaders As Boolean = True) As ICollectionEx: Set ICollectionEx_CopyToRange = CopyToRange(RangeObject, ApplyUserLocale, WriteHeaders): End Function




