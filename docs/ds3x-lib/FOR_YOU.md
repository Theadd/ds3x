## Count()

### Array2dEx
```vb
Public Property Get Count() As Long
    On Error Resume Next
    Count = 1 + (UBound(pInstance, 1) - LBound(pInstance, 1))
    On Error GoTo 0
End Property
```
Returns the number of rows in the 2D array.

### ArrayListEx
```vb
Public Property Get Count() As Variant
Attribute Count.VB_Description = "Gets the count of elements in the underlying mscorlib.ArrayList."
    Count = Instance.Count
End Property
```
Gets the count of elements in the underlying mscorlib.ArrayList.

### ArraySliceGroup
```vb
Public Property Get Count() As Long
    On Error Resume Next
    Count = pGroups(0).Count
    On Error GoTo 0
End Property
```
Returns the number of elements in the first slice group.

### dsTable
```vb
Public Property Get Count() As Long
    On Error Resume Next
    Count = Records.Count
    On Error GoTo 0
End Property
```
Returns the number of records in the dsTable.

### DictionaryEx
```vb
Public Property Get Count() As Variant
    Count = Instance.Count
End Property
```
Gets the key count of the underlying Scripting.Dictionary.

### RecordsetEx
```vb
Public Property Get Count() As Long
Attribute Count.VB_Description = "Equivalente a Recordset.RecordCount"
    On Error Resume Next
    Count = Instance.RecordCount
    On Error GoTo 0
End Property
```
Equivalente a Recordset.RecordCount
