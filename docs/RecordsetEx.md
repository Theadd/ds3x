## **`RecordsetEx` Class** <sup><sub><sup> &nbsp; <kbd><code>__CHAINABLE__</code></kbd></sup></sub></sup>

_A **lightweight** wrapper around `ADODB.Recordset` objects._

Implements __[ICollectionEx](./ICollectionEx.md)__.

---

- All methods returning an `RecordsetEx` are chainable.
- Publicly exposes the inner `ADODB.Recordset` instance for extensibility purposes with no code alterning needed.
- It's enumerable (for each support on rows) and with `.Item` as the default class member.
  - Where `.Item` can either return a shallow copy of this `RecordsetEx`, pointing at the provided numeric position or, directly return an `ADODB.Field` when a <code>String</code> value is supplied.
- `RecordsetEx.Create()` supports creating new instances by directly converting from:
  - **ADODB.Recordset**.
  - **ArrayList** and **ArrayListEx**.

---

### **API Overview**

```vb
' Fields
Public Instance As ADODB.Recordset
' Properties
Property Get Count() As Long
Property Get ColumnCount() As Long
Property Get Item(NumRowOrFieldName As Variant) As Variant
Property Get Fields() As ADODB.Fields
Property Get Field(FieldNameOrIndex As Variant) As ADODB.Field
Property Get Row(Index As Long) As Variant
Property Get ColumnNames() As Variant()
Property Get Items() As Variant()
' Functions
Function Create(Optional RecordsetLike As Variant) As RecordsetEx
Function Bind(TargetRecordset As ADODB.Recordset) As RecordsetEx
Function Unbind() As RecordsetEx
Function Filter(Optional QueryFilter As String) As RecordsetEx
Function MoveFirst() As RecordsetEx
Function MoveLast() As RecordsetEx
Function MoveNext() As RecordsetEx
Function MovePrevious() As RecordsetEx
Function Move(NumRecords As Long) As RecordsetEx
Function AddRange(Target As Varriant) As RecordsetEx
Function GetRange(Optional Index As Long, Optional GetCount As Long, Optional ColumnIndexes As Variant) As Array2dEx
Function CopyToRange(RangeObject As Excel.Range, Optional ApplyUserLocale As Boolean, Optional WriteHeaders As Boolean) As RecordsetEx
Function AsIterable(ParamArray QueryColumns() As Variant) As Variant()
Function AsIterableString(Delimiter As String, ParamArray QueryColumns() As Variant) As Variant()
Function AsIterableDictionary(ParamArray QueryColumns() As Variant) As Variant()
Function Requery() As RecordsetEx
Function IndexOf(SearchCriteria as String, Optional DefaultValue As Long) As Long
Function LastIndexOf(SearchCriteria as String, Optional DefaultValue As Long) As Long
Function Clone() As RecordsetEx
Function Duplicate() As RecordsetEx
Function Search(SearchCriteria As Variant, Optional ContinueBookmark As Variant, Optional SearchDirection As SearchDirectionEnum) As Variant
Function ToCSV(Optional Delimiter As String) As String
Function ToJSON() As String
Function ToExcel() As String
' Static
Function CreateBlank(RowsCount As Long, ColumnsCount As Long) As RecordsetEx
```


<table width="100%"><caption>

### **`RecordsetEx` API**  
</caption>
<thead><tr><th colspan="2">FIELDS</th></tr></thead>
<tbody>


<tr><td align="left" valign="top">

```vb
Instance As ADODB.Recordset
```
</td><td align="left" valign="top">
The <code>ADODB.Recordset</code> object wrapped by this <code>RecordsetEx</code> instance.
</td></tr>

</tbody>

</caption>
<thead><tr><th colspan="2">PROPERTIES</th></tr></thead>
<tbody>


<tr><td align="left" valign="top">

```vb
Get Count() As Long
```
</td><td align="left" valign="top">
Gets the <code>ADODB.Recordset</code>'s <code>.RecordCount</code>.
</td></tr>


<tr><td align="left" valign="top">

```vb
Get ColumnCount() As Long
```
</td><td align="left" valign="top">
Gets the <code>ADODB.Recordset</code>'s <code>.Fields.Count</code>.
</td></tr>


<tr><td align="left" valign="top">

```vb
Get Item(NumRowOrFieldName As Variant) As Variant
```
</td><td align="left" valign="top">
Gets a shallow copy (not even a <code>Recordset</code> clone) of this <code>RecordsetEx</code> which, at the moment you request any of it's values, and not until then, it will move the recordset's current record to the one it was related to (1-based).
<br/>But, if a <code>String</code> is provided, it will get the corresponding <code>ADODB.Field</code> that this <code>RecordsetEx</code> (or a shallow copy of it) is pointing to.
<br/>To access a field by it's <code>Index</code>, just supply it as <code>String</code> (e.g.: <code>"0"</code> instead of <code>0</code>).
</td></tr>


<tr><td align="left" valign="top">

```vb
Get Fields() As ADODB.Fields
```
</td><td align="left" valign="top">
Gets the <code>ADODB.Recordset</code>'s <code>.Fields</code> collection.
</td></tr>


<tr><td align="left" valign="top">

```vb
Get Field(FieldNameOrIndex As Variant) As ADODB.Field
```
</td><td align="left" valign="top">
Gets the <code>ADODB.Field</code> with the specified <code>FieldName</code> or <code>Index</code>.
<br/><em>Note: <code>.Value</code> is the default member of <code>ADODB.Field</code>.
</td></tr>


<tr><td align="left" valign="top">

```vb
Get Row(Index As Long) As Variant
```
</td><td align="left" valign="top">
Gets an <code>Array</code> containing all field values at the specified row <code>Index</code> (0-based). 
</td></tr>


<tr><td align="left" valign="top">

```vb
Get ColumnNames() As Variant()
```
</td><td align="left" valign="top">
Gets an <code>Array</code> containing the field names.
</td></tr>


<tr><td align="left" valign="top">

```vb
Get Items() As Variant()
```
</td><td align="left" valign="top">
Gets an iterable <code>Array</code> to be used in <code>For Each</code> loops.
</td></tr>


</tbody>


<thead><tr><th colspan="2">FUNCTIONS</th></tr></thead>
<tbody>


<tr><td align="left" valign="top">

```vb
Create(Optional RecordsetLike As Variant) As RecordsetEx
```
</td><td align="left" valign="top">
Returns a new <code>RecordsetEx</code> instance containing the <code>ADODB.Recordset</code> obtained by converting the provided <code>RecordsetLike</code> to an <code>ADODB.Recordset</code>.
</td></tr>


<tr><td align="left" valign="top">

```vb
Bind(TargetRecordset As ADODB.Recordset) As RecordsetEx
```
</td><td align="left" valign="top">
Instead of returning a new <code>RecordsetEx</code> instance, points this <code>RecordsetEx</code> instance to the specified <code>ADODB.Recordset</code>.
</td></tr>


<tr><td align="left" valign="top">

```vb
Unbind() As RecordsetEx
```
</td><td align="left" valign="top">
Removes the reference to the <code>ADODB.Recordset</code> instance.
</td></tr>


<tr><td align="left" valign="top">

```vb
Filter(Optional QueryFilter As String) As RecordsetEx
```
</td><td align="left" valign="top">
Chainable wrapper around <code>ADODB.Recordset</code>'s <code>Filter</code> property.
</td></tr>


<tr><td align="left" valign="top">

```vb
MoveFirst() As RecordsetEx
```
</td><td align="left" valign="top">
Chainable safe wrapper around <code>ADODB.Recordset</code>'s <code>MoveFirst</code>.
</td></tr>


<tr><td align="left" valign="top">

```vb
MoveLast() As RecordsetEx
```
</td><td align="left" valign="top">
Chainable safe wrapper around <code>ADODB.Recordset</code>'s <code>MoveLast</code>.
</td></tr>


<tr><td align="left" valign="top">

```vb
MoveNext() As RecordsetEx
```
</td><td align="left" valign="top">
Chainable safe wrapper around <code>ADODB.Recordset</code>'s <code>MoveNext</code>.
</td></tr>


<tr><td align="left" valign="top">

```vb
MovePrevious() As RecordsetEx
```
</td><td align="left" valign="top">
Chainable safe wrapper around <code>ADODB.Recordset</code>'s <code>MovePrevious</code>.
</td></tr>


<tr><td align="left" valign="top">

```vb
Move(NumRecords As Long) As RecordsetEx
```
</td><td align="left" valign="top">
Chainable safe wrapper around <code>ADODB.Recordset</code>'s <code>Move</code>.
</td></tr>


<tr><td align="left" valign="top">

```vb
AddRange(Target As Varriant) As RecordsetEx
```
</td><td align="left" valign="top">
Returns a new <code>RecordsetEx</code> instance containing the concatenation of all the records in this <code>ADODB.Recordset</code> with the records in <code>Target</code>'s <code>ADODB.Recordset</code>. The number of fields in both recordsets should be the same.
</td></tr>


<tr><td align="left" valign="top">

```vb
GetRange(Optional Index As Long, Optional GetCount As Long, Optional ColumnIndexes As Variant) As Array2dEx
```
</td><td align="left" valign="top">
Returns a new <code>Array2dEx</code> instance with the specified range of elements.
<br/><em>@see: <code>Array2dEx.GetRange</code></em>.
</td></tr>


<tr><td align="left" valign="top">

```vb
CopyToRange(RangeObject As Excel.Range, Optional ApplyUserLocale As Boolean, Optional WriteHeaders As Boolean) As RecordsetEx
```
</td><td align="left" valign="top">
Internally converts it to a <code>dsTable</code> and calls it's <code>CopyToRange</code> method instead of implementing it's own.
<br/><em>@see: <code>dsTable.CopyToRange</code>.</em>
</td></tr>


<tr><td align="left" valign="top">

```vb
AsIterable(ParamArray QueryColumns() As Variant) As Variant()
```
</td><td align="left" valign="top">
Returns the <code>ADODB.Recordset</code> as a <b>Jagged Array</b> <em>(Array of arrays)</em>.
</td></tr>


<tr><td align="left" valign="top">

```vb
AsIterableString(Delimiter As String, ParamArray QueryColumns() As Variant) As Variant()
```
</td><td align="left" valign="top">
Returns the <code>ADODB.Recordset</code> as a <code>String Array</code>, concatenating all values in each record with the specified <code>Delimiter</code>.
<details><summary><code>EXAMPLE</code></summary>

```vb
With RecordsetEx.Bind(pUsersRecordset).Filter("IsAdmin = 1")
    For Each cbItem In .AsIterableString(";", "Username", "FullName")
        Me.AdminUsersCombobox.AddItem cbItem
    Next cbItem
End With
```
</details>
</td></tr>


<tr><td align="left" valign="top">

```vb
AsIterableDictionary(ParamArray QueryColumns() As Variant) As Variant()
```
</td><td align="left" valign="top">
Returns the <code>ADODB.Recordset</code> as a <code>Scripting.Dictionary Array</code>.
</td></tr>


<tr><td align="left" valign="top">

```vb
Requery() As RecordsetEx
```
</td><td align="left" valign="top">
Better and safer alternative to <code>ADODB.Recordset</code>'s <code>Requery</code>.
</td></tr>


<tr><td align="left" valign="top">

```vb
IndexOf(SearchCriteria as String, Optional DefaultValue As Long) As Long
```
</td><td align="left" valign="top">
Returns the 1-based <code>Index</code> position of the first record matching <code>SearchCriteria</code> if found, otherwise, <code>DefaultValue</code> (which defaults to <code>-1</code>).
<details><summary><code>EXAMPLE</code></summary>
Select the first record matching <code>Id = 42</code> in a continuous form if any, do nothing otherwise.

```vb
Me.SelTop = RecordsetEx.Bind(Me.Recordset).IndexOf("Id = 42", Me.SelTop)
```
</details>
</td></tr>


<tr><td align="left" valign="top">

```vb
LastIndexOf(SearchCriteria as String, Optional DefaultValue As Long) As Long
```
</td><td align="left" valign="top">
Returns the 1-based <code>Index</code> position of the <b>last</b> record matching <code>SearchCriteria</code> if found, otherwise, <code>DefaultValue</code> (which defaults to <code>-1</code>).
</td></tr>


<tr><td align="left" valign="top">

```vb
Clone() As RecordsetEx
```
</td><td align="left" valign="top">
Returns a new <code>RecordsetEx</code> instance pointing to a <code>ADODB.Recordset</code>'s <code>Clone</code> of this recordset.
</td></tr>


<tr><td align="left" valign="top">

```vb
Duplicate() As RecordsetEx
```
</td><td align="left" valign="top">
Returns a new <code>RecordsetEx</code> containing a deep copy of this <code>ADODB.Recordset</code>.
</td></tr>


<tr><td align="left" valign="top">

```vb
Search(SearchCriteria As Variant, Optional ContinueBookmark As Variant, Optional SearchDirection As SearchDirectionEnum) As Variant
```
</td><td align="left" valign="top">
Returns the <code>ADODB.Recordset</code>'s <code>Bookmark</code> of the first record matching <code>SearchCriteria</code>.
<details><summary><code>PARAMETERS</code></summary><ul>
<li><kbd>SearchCriteria</kbd> → It can be a <code>String</code> or an <code>Array</code> in <code>SearchDefinition</code> format.
<br/><em>@see: <code>CollectionsLib.GenerateSearchDefinition</code></em>.</li>
<li><kbd>ContinueBookmark</kbd> → A <code>Bookmark</code> of the previous record to start searching for.
<br/>Default value: <code>-1</code>.</li>
<li><kbd>SearchDirection</kbd> → A <code>SearchDirectionEnum</code> value telling in which direction to search.
<br/>Default value: <code>adSearchForward</code></li>
</ul></details>
</td></tr>


<tr><td align="left" valign="top">

```vb
ToCSV(Optional Delimiter As String) As String
```
</td><td align="left" valign="top">
Returns a <code>String</code> representing this instance in <code>CSV</code>-style format.
<details><summary><code>PARAMETERS</code></summary><ul>
<li><kbd>Delimiter</kbd> → Character <code>String</code> used as delimiter between values within each row.<br/>Default value: <code>","</code></li>
</ul></details>
</td></tr>


<tr><td align="left" valign="top">

```vb
ToJSON() As String
```
</td><td align="left" valign="top">
Returns a <code>String</code> representing this instance in <code>JSON</code> format.
</td></tr>


<tr><td align="left" valign="top">

```vb
ToExcel() As String
```
</td><td align="left" valign="top">
Provides a simple way of direct copy-paste to an <code>Excel</code> document. <em>@see: <code>FileSystemLib.SystemClipboard</code>.</em>
<br/>Same as <code>.ToCSV(vbTab)</code>.
</td></tr>


</tbody>


<thead><tr><th colspan="2">PROCEDURES</th></tr></thead>
<tbody>


<tr><td align="left" valign="top">

```vb
Dispose()
```
</td><td align="left" valign="top">
Tells <code>RecordsetEx</code> that this instance won't be needed anymore and it can be safely disposed.
<br/>
<em>This shouldn't be necessary in most cases since all objects are automatically destroyed when there's no reference to them.</em>
</td></tr>


</tbody>


<thead><tr><th colspan="2">STATIC</th></tr></thead>
<tbody>


<tr><td align="left" valign="top">

```vb
CreateBlank(RowsCount As Long, ColumnsCount As Long) As RecordsetEx
```
</td><td align="left" valign="top">
Returns a new <code>RecordsetEx</code> instance with the specified number of rows and columns, containing <code>Empty</code> values.
</td></tr>


</tbody>

</table>




