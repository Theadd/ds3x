VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICollectionEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder "ds3x.Collections"
'@Exposed
'@Interface

Option Explicit

Public Property Get Count() As Long
End Property

Public Property Get ColumnCount() As Long
End Property

Public Property Get Item(ByVal Index As Long, Optional ByVal ColumnIndex As Long = 0) As Variant
End Property

Public Property Get Row(ByVal Index As Long) As Variant
End Property

Public Function CreateBlank(ByVal RowsCount As Long, ByVal ColumnsCount As Long) As ICollectionEx
End Function

Public Function Create(Optional ByRef FromTarget As Variant) As ICollectionEx
End Function

Public Function Bind(Optional ByRef Target As Variant) As ICollectionEx
End Function

Public Function Unbind() As ICollectionEx
End Function

Public Sub Dispose()
End Sub

Public Function Join(ByVal Target As ICollectionEx) As ICollectionEx
End Function

Public Function AddRange(ByVal Target As ICollectionEx) As ICollectionEx
End Function

Public Function GetRange(Optional ByVal Index As Variant, Optional ByVal GetCount As Variant, Optional ByVal ColumnIndexes As Variant) As ICollectionEx
End Function

Public Function ToCSV(Optional ByVal Delimiter As String = ",") As String
End Function

Public Function ToJSON() As String
End Function

Public Function ToExcel() As String
End Function

Public Function CopyToRange(ByRef RangeObject As Excel.Range, Optional ByVal ApplyUserLocale As Boolean = True, Optional ByVal WriteHeaders As Boolean = True) As ICollectionEx
End Function
