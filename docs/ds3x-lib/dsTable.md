## **`dsTable` Class** <sup><sub><sup> &nbsp; <kbd><code>__IMMUTABLE__</code></kbd></sup></sub></sup>

_A **structured data table** implementation with header/record separation and Excel integration._

Implements __[ICollectionEx](./ICollectionEx.md)__.

---

- Maintains separate header definitions and record data
- Supports automatic type inference from Excel ranges
- Immutable operations return new instances
- Direct conversion to/from Excel ranges with formatting preservation
- Implements same ICollectionEx interface as other data types

---

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
Array of values for specified row
</td></tr>

<tr><td align="left" valign="top">

```vb
Get Headers() As ICollectionEx
```
</td><td align="left" valign="top">
Column definitions (name, format, type)
</td></tr>

<tr><td align="left" valign="top">

```vb
Get Records() As ICollectionEx
```
</td><td align="left" valign="top">
Underlying data storage (ArraySliceGroup)
</td></tr>

</tbody>

<thead><tr><th colspan="2">FUNCTIONS</th></tr></thead>
<tbody>

<tr><td align="left" valign="top">

```vb
Create(Optional TableLike, AutoHeaders) As dsTable
```
</td><td align="left" valign="top">
Construct from array/dictionary/recordset/range. AutoHeaders detects column names from first row.
</td></tr>

<tr><td align="left" valign="top">

```vb
SetHeaders(AllHeaders) As dsTable
```
</td><td align="left" valign="top">
Define columns with name/format/type info
</td></tr>

<tr><td align="left" valign="top">

```vb
CopyToRange(RangeObject, ApplyUserLocale, WriteHeaders)
```
</td><td align="left" valign="top">
Export to Excel with formatting and headers
</td></tr>

<tr><td align="left" valign="top">

```vb
ToCSV(Delimiter, InLocalFormat) As String
```
</td><td align="left" valign="top">
Export as CSV with header row
</td></tr>

<tr><td align="left" valign="top">

```vb
ToJSON() As String
```
</td><td align="left" valign="top">
JSON representation with schema
</td></tr>

</tbody>

<thead><tr><th colspan="2">STATIC</th></tr></thead>
<tbody>

<tr><td align="left" valign="top">

```vb
CreateBlank(Rows, Columns) As dsTable
```
</td><td align="left" valign="top">
Empty table with specified dimensions
</td></tr>

</tbody>
</table>
