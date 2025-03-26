## **`ArraySliceGroup` Class** <sup><sub><sup> &nbsp; <kbd><code>__CHAINABLE__</code></kbd></sup></sub></sup>

_A **high-performance** collection of columnar data slices._

Implements __[ICollectionEx](./ICollectionEx.md)__.

---

### **Architecture & Benefits**
- **Immutability**: Operations return new instances to ensure thread safety and predictable state.
- **Efficient Memory Usage**: Slice-based operations avoid copying data by referencing underlying arrays.
- **Columnar Data Handling**: Groups `ArraySlice` instances as columns for multi-dimensional operations.
- **Integration**: Converts seamlessly to `Array2dEx`, `dsTable`, or Excel ranges via `CopyToRange`.
- **Performance**: Leverages pointer arithmetic for O(1) slice operations, ideal for large datasets.

---

### **Key Features**
- **Factory Methods**: `CreateBlank`, `CreateFrom` for initializing from arrays/objects.
- **Column Manipulation**: Add/insert/remove columns with `Add`, `Insert`, `RemoveAt`.
- **Range Operations**: `GetRange` for row/column subsetting, `Join` to merge groups.
- **Format Conversion**: `ToCSV`, `ToJSON`, and Excel integration via `CopyToRange`.

---

### **Usage Examples**
* **Create from scratch**:
```vb
Dim group As ArraySliceGroup
Set group = ArraySliceGroup.CreateBlank(100, 3) ' 100 rows × 3 columns
```

* **Convert Excel range to slices**:
```vb
Dim ws As xlSheetsEx
Set ws = xlSheetsEx.Create("Data")
Set group = ArraySliceGroup.Create(ws.UsedRange)
```

* **Perform column operations**:
```vb
' Add a new column from another slice group
Set group = group.Add(otherGroup.SliceAt(0))

' Extract rows 5-15 from columns 0 and 2
Set subset = group.GetRange(5, 10, Array(0, 2))
```

---

### **API Overview**
```vb
' Core Properties
Public Property Get Count() As Long
Public Property Get ColumnCount() As Long
Public Property Get Instance() As Array2dEx

' Factory Methods
Public Function CreateBlank(RowsCount As Long, ColumnsCount As Long) As ArraySliceGroup
Public Function Create(Optional ArrayLike As Variant) As ArraySliceGroup

' Column Operations
Public Function Add(Target As ArraySlice) As ArraySliceGroup
Public Function Insert(Target As ArraySlice, ColumnIndex As Long) As ArraySliceGroup
Public Function RemoveAt(ColumnIndex As Long) As ArraySliceGroup

' Range Selection
Public Function GetRange( _
    Optional Index As Variant, _
    Optional GetCount As Variant, _
    Optional ColumnIndexes As Variant _
) As ArraySliceGroup

' Data Conversion
Public Sub CopyToRange(RangeObject As Excel.Range, Optional ApplyUserLocale As Boolean)
Public Function ToCSV(Delimiter As String) As String
Public Function ToJSON() As String
```

---

<table width="100%"><caption>

### **`ArraySliceGroup` API**  
</caption>
<thead><tr><th colspan="2">PROPERTIES</th></tr></thead>
<tbody>

<tr><td align="left" valign="top">
```vb
Public Property Get Count() As Long
```
</td><td align="left" valign="top">
Total number of rows in the first column slice.
</td></tr>

<tr><td align="left" valign="top">
```vb
Public Property Get ColumnCount() As Long
```
</td><td align="left" valign="top">
Number of column slices in the group.
</td></tr>

<tr><td align="left" valign="top">
```vb
Public Property Get Instance() As Array2dEx
```
</td><td align="left" valign="top">
Converts the slice group to an `Array2dEx` instance for direct array access.
</td></tr>

</tbody>

<thead><tr><th colspan="2">FUNCTIONS</th></tr></thead>
<tbody>

<tr><td align="left" valign="top">
```vb
Public Function CreateBlank(RowsCount As Long, ColumnsCount As Long) As ArraySliceGroup
```
</td><td align="left" valign="top">
Creates a new group with empty slices of specified dimensions.
</td></tr>

<tr><td align="left" valign="top">
```vb
Public Function Create(Optional ArrayLike As Variant) As ArraySliceGroup
```
</td><td align="left" valign="top">
Initializes from `Array2dEx`, `Excel.Range`, or other array-like structures.
</td></tr>

<tr><td align="left" valign="top">
```vb
Public Function Add(Target As ArraySlice) As ArraySliceGroup
```
</td><td align="left" valign="top">
Adds a new column slice at the end of the group.
</td></tr>

<tr><td align="left" valign="top">
```vb
Public Function Insert(Target As ArraySlice, ColumnIndex As Long) As ArraySliceGroup
```
</td><td align="left" valign="top">
Inserts a column slice at the specified index.
</td></tr>

<tr><td align="left" valign="top">
```vb
Public Function GetRange( _
    Optional Index As Variant, _
    Optional GetCount As Variant, _
    Optional ColumnIndexes As Variant _
) As ArraySliceGroup
```
</td><td align="left" valign="top">
Creates a new group containing a subset of rows/columns.
<details><summary><code>PARAMETERS</code></summary><ul>
<li><kbd>Index</kbd> → Starting row index (default: 0)</li>
<li><kbd>GetCount</kbd> → Number of rows to include (default: all remaining)</li>
<li><kbd>ColumnIndexes</kbd> → Array of column indices to include (default: all)</li>
</ul></details>
</td></tr>

<tr><td align="left" valign="top">
```vb
Public Sub CopyToRange(RangeObject As Excel.Range, Optional ApplyUserLocale As Boolean)
```
</td><td align="left" valign="top">
Exports data to an Excel range with optional locale formatting.
</td></tr>

<tr><td align="left" valign="top">
```vb
Public Function ToCSV(Delimiter As String) As String
```
</td><td align="left" valign="top">
Exports data to CSV format with specified delimiter.
</td></tr>

</tbody>
</table>
