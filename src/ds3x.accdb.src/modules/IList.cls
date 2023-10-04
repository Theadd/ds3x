VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Interface
'@Folder("ds3x.Collections")
Option Compare Database
Option Explicit




Public Property Get Count() As Long
End Property

Public Property Get Version() As Long
End Property

Public Property Get Capacity() As Long
End Property

Public Property Let Capacity(ByVal Value As Long)
End Property

Public Property Get IsFixedSize() As Boolean
End Property

Public Property Get IsReadOnly() As Boolean
End Property

Public Property Get IsSynchronized() As Boolean
End Property

Public Property Get SyncRoot() As Object
End Property

Public Property Get Item(ByVal Index As Long) As Variant
End Property

Public Property Let Item(ByVal Index As Long, ByRef Value As Variant)
End Property

Public Property Set Item(ByVal Index As Long, ByRef Value As Variant)
End Property

Public Property Get Items() As Variant()
End Property


Public Function Add(ByRef Value As Variant) As Long
End Function

Public Sub Insert(ByVal Index As Long, ByRef Value As Variant)
End Sub

Public Sub AddRange(ByRef ArrayLike As Variant)
End Sub

Public Sub InsertRange(ByVal Index As Long, ByRef ArrayLike As Variant)
End Sub

Public Function GetRange(ByVal Index As Long, ByVal GetCount As Long) As IList
End Function

Public Sub Clear()
End Sub

Public Function Clone() As IList
End Function

Public Function Contains(ByRef Value As Variant) As Boolean
End Function

Public Sub CopyTo(Optional ByVal Index As Long = 0, Optional ByRef TargetArray As Variant, Optional ByVal TargetIndex As Long = 0, Optional ByVal GetCount As Variant)
End Sub

Public Function IndexOf(ByRef Value As Variant, Optional ByVal Index As Long = 0, Optional ByVal GetCount As Variant) As Long
End Function

Public Function LastIndexOf(ByRef Value As Variant, Optional ByVal Index As Long = 0, Optional ByVal GetCount As Variant) As Long
End Function

Public Function BinarySearch(ByVal Index As Long, ByVal GetCount As Long, ByRef Value As Variant, Optional ByRef Comparer As IComparer = Nothing) As Long
End Function

Public Sub Remove(ByRef Value As Variant)
End Sub

Public Sub RemoveAt(ByVal Index As Long)
End Sub

Public Sub RemoveRange(ByVal Index As Long, ByVal GetCount As Long)
End Sub

Public Sub Reverse(Optional ByVal Index As Long = 0, Optional ByVal GetCount As Variant)
End Sub

Public Sub SetRange(ByVal Index As Long, ByRef ArrayLike As Variant)
End Sub

Public Sub Sort(Optional ByVal Index As Long = 0, Optional ByVal GetCount As Variant, Optional ByRef Comparer As IComparer = Nothing)
End Sub

Public Function ToArray() As Variant()
End Function

Public Function GetAddressOfItemAt(ByVal Index As Long) As LongPtr
End Function
