## **`dsTable` Class** <sup><sub><sup> &nbsp; <kbd><code>__IMMUTABLE__</code></kbd></sup></sub></sup>

_A **structured data table** implementation with header/record separation and Excel integration._

Implements __[ICollectionEx](./ICollectionEx.md)__.

---

- Maintains separate header definitions and record data with type inference
- Immutable operations return new instances preserving data integrity
- Direct conversion to/from Excel ranges with formatting retention
- Implements ICollectionEx interface for consistent data handling
- Supports advanced operations like recordset creation and markdown export

- `dsTable.Create()` supports creating instances from:
  - **Array2dEx**, **ArraySliceGroup**, or **2D arrays**
  - **ADODB.Recordset**/**RecordsetEx**
  - **ArrayListEx**
  - **Excel.Range**
  - **Scripting.Dictionary**/**DictionaryEx**
  - **JSON string** (parsed as a dictionary-like structure)
  - **Variant arrays** (1D or 2D)

---

### **Usage Examples**

* Create a table from an Excel range and export to CSV:

```vb
Dim ds As dsTable
Set ds = dsTable.Create(Worksheets("Data").UsedRange, AutoHeaders:=True)
ds.CopyToRange Worksheets("Output").Range("A1")
ds.ToCSV > "output.csv"
```

---

### **API Overview**

<table width="100%"><caption>

### **`dsTable` API**  
</caption>
<thead><tr><th colspan="2">PROPERTIES</th></tr></thead>
<tbody>


<tr><td align="left" valign="top">

```vb
Get Count() As Long
```
</td><td align="left" valign="top">
Number of data rows in the table
</td></tr>


<tr><td align="left" valign="top">

```vb
Get ColumnCount() As Long
```
</td><td align="left" valign="top">
Number of columns (matches header count)
</td></tr>


<tr><td align="left" valign="top">

```vb
Get Row(Index As Long) As Variant
```
</td><td align="left" valign="top">
Returns an array of values for the specified row index
</td></tr>


<tr><td align="left" valign="top">

```vb
Get Headers() As ICollectionEx
```
</td><td align="left" valign="top">
Column definitions including name, format, type, and size
</td></tr>


<tr><td align="left" valign="top">

```vb
Get Records() As ICollectionEx
```
</td><td align="left" valign="top">
Underlying data storage as ArraySliceGroup
</td></tr>
</tbody>



<thead><tr><th colspan="2">FUNCTIONS</th></tr></thead>
<tbody>


<tr><td align="left" valign="top">

```vb
Create(Optional TableLike As Variant, Optional AutoHeaders As Boolean = False) As dsTable
```
</td><td align="left" valign="top">
Creates a new instance from the following sources:
  - **Array2dEx**, **ArraySliceGroup**, or **2D arrays**
  - **ADODB.Recordset**/**RecordsetEx**
  - **ArrayListEx**
  - **Excel.Range**
  - **Scripting.Dictionary**/**DictionaryEx**
  - **JSON string** (parsed as a dictionary-like structure)
  - **Variant arrays** (1D or 2D)
  
AutoHeaders detects column names from the first row if enabled. If TableLike is omitted, returns an empty table.
</td></tr>

<tr><td align="left" valign="top">

```vb
CreateFromExcelRange(TargetRange As Excel.Range) As dsTable
```
</td><td align="left" valign="top">
Creates table instance from Excel range with automatic header detection
</td></tr>

<tr><td align="left" valign="top">

```vb
CreateNamedRecordset() As Recordset
```
</td><td align="left" valign="top">
Exports data to a named ADO recordset with schema based on headers
</td></tr>

<tr><td align="left" valign="top">

```vb
CreateIndexRecordset() As Recordset
```
</td><td align="left" valign="top">
Creates indexed ADO recordset with primary key based on first column
</td></tr>

<tr><td align="left" valign="top">

```vb
Bind(TableLike As Variant, Optional AutoHeaders As Boolean = False) As dsTable
```
</td><td align="left" valign="top">
Binds to existing data source preserving immutability
</td></tr>


<tr><td align="left" valign="top">

```vb
Unbind() As dsTable
```
</td><td align="left" valign="top">
Disassociates from current data source
</td></tr>


<tr><td align="left" valign="top">

```vb
SetHeaders(AllHeaders As Variant) As dsTable
```
</td><td align="left" valign="top">
Define columns with name/format/type info (immutable operation returns new instance)
</td></tr>


<tr><td align="left" valign="top">

```vb
Join(TargetTable As dsTable) As dsTable
```
</td><td align="left" valign="top">
Concatenates columns from another table (returns new instance)
</td></tr>


<tr><td align="left" valign="top">

```vb
AddRange(TargetTable As dsTable) As dsTable
```
</td><td align="left" valign="top">
Adds records from another table (returns new instance)
</td></tr>


<tr><td align="left" valign="top">

```vb
GetRange(Optional Index As Variant, Optional GetCount As Variant, Optional ColumnIndexes As Variant) As dsTable
```
</td><td align="left" valign="top">
Extract subset of data with optional row/column filtering
</td></tr>


<tr><td align="left" valign="top">

```vb
ToCSV(Delimiter As String = ",", InLocalFormat As Boolean = False) As String
```
</td><td align="left" valign="top">
Exports to CSV with optional delimiter and locale formatting
</td></tr>


<tr><td align="left" valign="top">

```vb
ToJSON() As String
```
</td><td align="left" valign="top">
JSON representation with schema and metadata
</td></tr>


<tr><td align="left" valign="top">

```vb
ToExcel() As String
```
</td><td align="left" valign="top">
Exports to tab-separated Excel-friendly format
</td></tr>


<tr><td align="left" valign="top">

```vb
ToMarkdownTable() As String
```
</td><td align="left" valign="top">
Generates markdown-formatted table with headers and data alignment
</td></tr>


<tr><td align="left" valign="top">

```vb
CopyToRange(RangeObject As Excel.Range, Optional ApplyUserLocale As Boolean = True, Optional WriteHeaders As Boolean = True) As dsTable
```
</td><td align="left" valign="top">
Exports to Excel with optional formatting and headers. Returns updated table instance
</td></tr>


</tbody>



<thead><tr><th colspan="2">STATIC</th></tr></thead>
<tbody>


<tr><td align="left" valign="top">

```vb
CreateBlank(RowsCount As Long, ColumnsCount As Long) As dsTable
```
</td><td align="left" valign="top">
Creates empty table with specified dimensions
</td></tr>


</tbody>

</table>
