VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DictionaryEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

Public Instance As Scripting.Dictionary


'Gets the key count of the underlying Scripting.Dictionary.
Public Property Get Count() As Variant
    Count = Instance.Count
End Property

'Wrapper property that returns the value for the specified key if it exists.
'If it doesn't exist, it returns the default rather than raise an error.
Public Property Get GetValue(mKey As Variant, Optional ByRef mDefaultVal As Variant) As Variant
    If Instance.Exists(mKey) Then
        If IsObject(Instance.Item(mKey)) Then
            Set GetValue = Instance.Item(mKey)
        Else
            GetValue = Instance.Item(mKey)
        End If
    Else
        If IsMissing(mDefaultVal) Then Exit Property

        If IsObject(mDefaultVal) Then
            Set GetValue = mDefaultVal
        Else
            GetValue = mDefaultVal
        End If
    End If
End Property

'Sets or returns an item for a specified key.
Public Property Get Item(mKey As Variant) As Variant
Attribute Item.VB_UserMemId = 0
    If Not Instance.Exists(mKey) Then Exit Property

    If IsObject(Instance.Item(mKey)) Then
        Set Item = Instance.Item(mKey)
    Else
        Item = Instance.Item(mKey)
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



'Public Function Create(Optional ByRef TargetDictionary As Scripting.Dictionary = Nothing) As DictionaryEx
'    Set Create = New DictionaryEx
'    If TargetDictionary Is Nothing Then
'        Dim t As Scripting.Dictionary
'        Set t = New Scripting.Dictionary
'        Create.Bind t
'    Else
'        Create.Bind TargetDictionary
'    End If
'End Function

Public Function Create(Optional ByRef DictionaryLike As Variant) As DictionaryEx
    Set Create = New DictionaryEx
    
    If Not IsMissing(DictionaryLike) Then
        If Not DictionaryLike Is Nothing Then
            Create.Bind DictionaryLike
            Exit Function
        End If
    End If
    
    Dim t As Scripting.Dictionary
    Set t = New Scripting.Dictionary
    Create.Bind t
End Function

'Public Function Bind(ByRef TargetDictionary As Scripting.Dictionary) As DictionaryEx
'    Set Instance = TargetDictionary: Set Bind = Me
'End Function

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
Public Function Keys()
    Keys = Instance.Keys
End Function

Public Function Clone() As DictionaryEx
    Dim t As New DictionaryEx
    Set t.Instance = ShallowCopyDictionary(Instance)
    
    Set Clone = t
'    Dim k As Variant
'    Set Clone = Create()
'
'    With Clone
'        For Each k In Me.Keys()
'            If IsObject(Instance.Item(k)) Then
'                Set .Item(k) = Instance.Item(k)
'            Else
'                .Item(k) = Instance.Item(k)
'            End If
'        Next k
'    End With
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




Private Function ShallowCopyDictionary(ByVal d As Scripting.Dictionary) As Scripting.Dictionary
    Dim s As New Scripting.Dictionary, k As Variant
    
    s.CompareMode = d.CompareMode
    For Each k In d.Keys
        s.Add k, d(k)
    Next k
    
    Set ShallowCopyDictionary = s
End Function

Private Sub Class_Terminate()
    On Error Resume Next
    Set Instance = Nothing
End Sub




