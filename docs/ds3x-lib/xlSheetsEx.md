## **`xlSheetsEx` Class** <sup><sub><sup> &nbsp; <sup>__CHAINABLE__</sup></sup></sub></sup>

_A **lightweight** wrapper around `Excel.Worksheet` objects._

---

- `xlSheetsEx` instances provide extended operations for Excel worksheets and workbooks.
- All methods returning an `xlSheetsEx` are chainable.
- Publicly exposes the inner `Excel.Worksheet` instance for direct access.
- Supports workbook creation, sheet management, formatting, and advanced Excel operations.
- Implements safety features like automatic cleanup and workbook protection handling.
- Follows project patterns: Uses `Array2dEx` for Excel interactions, error handlers for COM ops.

---

### **Usage Examples**

* Create a new workbook with formatted table:

```vb
Dim ws As xlSheetsEx
Set ws = xlSheetsEx.Create("DataSheet")
    .FormatAsTable("SalesData", "TableStyleMedium8")
    .AutoSizeCells(ws.UsedRange)
    .FreezeHeaders(1)
    .SaveWorkbook("report.xlsx", True)
```

---

### **API Overview**

```vb
' Fields
Public Instance As Excel.Worksheet
' Properties
Property Get AllSheets() As ArrayListEx
Property Get WindowVisibility() As Boolean
Property Let WindowVisibility(ByVal ShouldBeVisible As Boolean)
Property Get Workbook() As Excel.Workbook
Property Get SheetIndex() As Long
Property Get SheetName() As String
Property Let SheetName(ByVal Value As String)
Property Get Cells() As Excel.Range
Property Get Columns() As Excel.Range
Property Get UsedRange() As Excel.Range
Property Get Protected() As Boolean
Property Let Protected(ByVal Value As Boolean)
Property Let DefaultSaveFormat(ByVal Value As Excel.XlFileFormat)
' Functions
Function GetSheet(ByVal SheetNameOrIndex As Variant) As xlSheetsEx
Function CreateFrom(ByRef Target As Object) As xlSheetsEx
Function Create(Optional ByVal WorksheetName As String, Optional ByVal TargetFile As Variant, Optional ByVal Options As DictionaryEx) As xlSheetsEx
Function AddSheet(Optional ByVal WorksheetName As String = "Sheet%1") As xlSheetsEx
Function SaveWorkbook(ByVal targetPath As String, Optional ByVal EnableMacros As Boolean = False) As xlSheetsEx
Function AutoSizeCells(Optional ByVal TargetRange As Excel.Range) As xlSheetsEx
Function FormatAsTable(Optional ByVal TableName As String, Optional ByVal TableStyle As String) As xlSheetsEx
Function FreezeHeaders(Optional ByVal HeaderRows As Long, Optional ByVal HeaderColumns As Long) As xlSheetsEx
Function AutoFormatCells(Optional ByVal TargetRange As Excel.Range) As xlSheetsEx
Function GetColumnsAutoNumberFormats(Optional ByVal TargetRange As Excel.Range) As Variant
```

<table width="100%"><caption>

### **`xlSheetsEx` API**  
</caption>
<thead><tr><th colspan="2">FIELDS</th></tr></thead>
<tbody>
<tr><td align="left" valign="top">

```vb
Instance As Excel.Worksheet
```
</td><td align="left" valign="top">
The underlying Excel worksheet object.
</td></tr>
</tbody>

<thead><tr><th colspan="2">PROPERTIES</th></tr></thead>
<tbody>
<tr><td align="left" valign="top">

```vb
Get AllSheets() As ArrayListEx
```
</td><td align="left" valign="top">
List of all managed worksheets in this workbook.
</td></tr>

<tr><td align="left" valign="top">

```vb
Get WindowVisibility() As Boolean
Let WindowVisibility(ByVal ShouldBeVisible As Boolean)
```
</td><td align="left" valign="top">
Controls Excel application visibility during operations.
</td></tr>

<tr><td align="left" valign="top">

```vb
Get Workbook() As Excel.Workbook
```
</td><td align="left" valign="top">
Reference to the parent workbook object.
</td></tr>

<tr><td align="left" valign="top">

```vb
Get SheetIndex() As Long
```
</td><td align="left" valign="top">
Worksheet's position index in the workbook.
</td></tr>

<tr><td align="left" valign="top">

```vb
Get SheetName() As String
Let SheetName(ByVal Value As String)
```
</td><td align="left" valign="top">
Get/set the worksheet's display name.
</td></tr>

<tr><td align="left" valign="top">

```vb
Get Cells() As Excel.Range
```
</td><td align="left" valign="top">
Entire worksheet's cells range.
</td></tr>

<tr><td align="left" valign="top">

```vb
Get UsedRange() As Excel.Range
```
</td><td align="left" valign="top">
Range containing all used cells.
</td></tr>

<tr><td align="left" valign="top">

```vb
Get Protected() As Boolean
Let Protected(ByVal Value As Boolean)
```
</td><td align="left" valign="top">
Control worksheet protection state.
</td></tr>

<tr><td align="left" valign="top">

```vb
Let DefaultSaveFormat(ByVal Value As Excel.XlFileFormat)
```
</td><td align="left" valign="top">
Set default workbook file format (e.g., .xlsx vs .xlsm).
</td></tr>
</tbody>

<thead><tr><th colspan="2">FUNCTIONS</th></tr></thead>
<tbody>
<tr><td align="left" valign="top">

```vb
GetSheet(ByVal SheetNameOrIndex As Variant) As xlSheetsEx
```
</td><td align="left" valign="top">
Retrieve worksheet by name/index.
</td></tr>

<tr><td align="left" valign="top">

```vb
CreateFrom(ByRef Target As Object) As xlSheetsEx
```
</td><td align="left" valign="top">
Initialize from existing worksheet/range.
</td></tr>

<tr><td align="left" valign="top">

```vb
Create(Optional ByVal WorksheetName As String, Optional ByVal TargetFile As Variant, Optional ByVal Options As DictionaryEx) As xlSheetsEx
```
</td><td align="left" valign="top">
Create new workbook/worksheet with options.
</td></tr>

<tr><td align="left" valign="top">

```vb
AddSheet(Optional ByVal WorksheetName As String = "Sheet%1") As xlSheetsEx
```
</td><td align="left" valign="top">
Add new worksheet to workbook.
</td></tr>

<tr><td align="left" valign="top">

```vb
SaveWorkbook(ByVal targetPath As String, Optional ByVal EnableMacros As Boolean = False) As xlSheetsEx
```
</td><td align="left" valign="top">
Save workbook with macro permissions option.
</td></tr>

<tr><td align="left" valign="top">

```vb
AutoSizeCells(Optional ByVal TargetRange As Excel.Range) As xlSheetsEx
```
</td><td align="left" valign="top">
Auto-adjust column widths based on content.
</td></tr>

<tr><td align="left" valign="top">

```vb
FormatAsTable(Optional ByVal TableName As String, Optional ByVal TableStyle As String) As xlSheetsEx
```
</td><td align="left" valign="top">
Convert range to Excel table with styling.
</td></tr>

<tr><td align="left" valign="top">

```vb
FreezeHeaders(Optional ByVal HeaderRows As Long, Optional ByVal HeaderColumns As Long) As xlSheetsEx
```
</td><td align="left" valign="top">
Freeze panes for headers.
</td></tr>

<tr><td align="left" valign="top">

```vb
AutoFormatCells(Optional ByVal TargetRange As Excel.Range) As xlSheetsEx
```
</td><td align="left" valign="top">
Apply intelligent formatting based on data types.
</td></tr>

<tr><td align="left" valign="top">

```vb
GetColumnsAutoNumberFormats(Optional ByVal TargetRange As Excel.Range) As Variant
```
</td><td align="left" valign="top">
Detect optimal number formats for columns.
</td></tr>
</tbody>
</table>
