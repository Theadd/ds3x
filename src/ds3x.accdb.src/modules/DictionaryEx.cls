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

Public Instance As Scripting.Dictionary


'Gets the key count of the underlying Scripting.Dictionary.
Public Property Get Count() As Variant
    Count = Instance.Count
End Property

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
    Instance.Item(mKey) = mVal
End Property

Public Property Set Item(mKey As Variant, mVal As Variant)
    Set Instance.Item(mKey) = mVal
End Property

'Updates a key, replacing the old key with the new key.
Public Property Let Key(mKey As Variant, mNewKey As Variant)
    Instance.Key(mKey) = mNewKey
End Property


' --- PUBLIC METHODS ---

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

'Removes a key/item pair from the Scripting.Dictionary.
Public Function Remove(mKey As Variant) As DictionaryEx
    Instance.Remove mKey
    Set Remove = Me
End Function

'The RemoveAll method removes all key, item pairs from the Scripting.Dictionary.
Public Function RemoveAll() As DictionaryEx
    Instance.RemoveAll
    Set RemoveAll = Me
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
Public Function Entries() As Variant()
    Dim aX As ArrayListEx, dKey As Variant
    Set aX = ArrayListEx.Create()
    
    For Each dKey In Instance.Keys()
        aX.Add Array(dKey, Instance(dKey))
    Next dKey
    
    Entries = aX.ToArray()
End Function

'@Enumerator
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Dim enumerator As IUnknown, aX As ArrayListEx
    Set aX = ArrayListEx.Create(Entries())
    
    Set enumerator = aX.Instance.GetEnumerator(0, aX.Instance.Count)
    Set NewEnum = enumerator
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
    Set Unserialize.Instance = JSON.Parse(CStr(Serialized))
End Function

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

Private Function GetNestedItem(ByVal mKey As String) As Variant
    Dim t As Long, c As Long, c2 As Long, v As Variant, tEnd As Boolean, cEnd As Boolean
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
                        Assign v, Instance.Item(VBA.Mid$(mKey, 1, c - 1))
                        Assign v, v(c2)
                        If Len(mKey) = (c + Len(CStr(c2)) + 1) Then
                            Assign GetNestedItem, v
                        Else
                            Assign GetNestedItem, ContinueNestedItem(v, VBA.Mid$(mKey, 1 + c + Len(CStr(c2)) + 1, Len(mKey)))
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

Private Sub Class_Terminate()
    On Error Resume Next
    Set Instance = Nothing
End Sub
