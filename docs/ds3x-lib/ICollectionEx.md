## **`ICollectionEx` Interface**

A shared interface providing common members for other collection-like class modules.

Implemented by __[ArrayListEx](./ArrayListEx.md)__, __[Array2dEx](./Array2dEx.md)__, __[DictionaryEx](./DictionaryEx.md)__, __[RecordsetEx](./RecordsetEx.md)__ and __[dsTable](./dsTable.md)__.

---

<table width="100%"><caption>

### **`ICollectionEx` API**  
</caption>
<thead><tr><th colspan="2">PROPERTIES</th></tr></thead>
<tbody>


<tr><td align="left" valign="top">

```vb
Get Count() As Long
```
</td><td align="left" valign="top">
Gets the number of elements in a list-like collection or the number of rows in a table-like collection.
</td></tr>


<tr><td align="left" valign="top">

```vb
Get ColumnCount() As Long
```
</td><td align="left" valign="top">
Gets the number of columns in a table-like collection.
</td></tr>


<tr><td align="left" valign="top">

```vb
Get Row(Index As Long) As Variant
```
</td><td align="left" valign="top">
Gets an <code>Array</code> containing all the elements at the specified row <code>Index</code> in a table-like collection. 
</td></tr>


</tbody>



<thead><tr><th colspan="2">FUNCTIONS</th></tr></thead>
<tbody>


<tr><td align="left" valign="top">

```vb
CreateBlank(RowsCount As Long, ColumnsCount As Long) As ICollectionEx
```
</td><td align="left" valign="top">
Returns a new table-like <code>ICollectionEx</code> instance with the specified number of rows and columns, containing <code>Empty</code> values.
</td></tr>


<tr><td align="left" valign="top">

```vb
Create(Optional FromTarget As Variant) As ICollectionEx
```
</td><td align="left" valign="top">
When no parameter is provided, returns a new instance of the class module implementing this interface.
<br/>
Otherwise, returns a new instance composed of the elements obtained by converting the provided parameter to the object type of the class module implementing this interface.
</td></tr>


<tr><td align="left" valign="top">

```vb
Bind(Optional Target As Variant) As ICollectionEx
```
</td><td align="left" valign="top">
Instead of returning a new instance referencing a <code>Target</code> object, tells this instance to reference another object instead.
</td></tr>


<tr><td align="left" valign="top">

```vb
Unbind() As ICollectionEx
```
</td><td align="left" valign="top">
Dereferences any object being wrapped by this instance.
</td></tr>


<tr><td align="left" valign="top">

```vb
Join(Target As ICollectionEx) As ICollectionEx
```
</td><td align="left" valign="top">
In a table-like collection, concatenates all elements of another <code>Target</code> collection as additional columns into a new <code>ICollectionEx</code> instance.
</td></tr>


<tr><td align="left" valign="top">

```vb
AddRange(Target As ICollectionEx) As ICollectionEx
```
</td><td align="left" valign="top">
Appends the elements or rows of another <code>ICollectionEx</code> at the end of this one. On table-like collections, the number of columns in both collections should be the same.
</td></tr>


<tr><td align="left" valign="top">

```vb
GetRange(Optional Index As Long, Optional GetCount As Long, Optional ColumnIndexes As Variant) As ICollectionEx
```
</td><td align="left" valign="top">
Returns a new <code>ICollectionEx</code> instance which represents a subset of elements (in a list-like collection) or rows (in a table-like collection) in this instance.
<details><summary><code>PARAMETERS</code></summary><ul>
<li><kbd>Index</kbd> → Index of the first element or row to include in the subset.
<br/>Default value: <code>0</code></li>
<li><kbd>GetCount</kbd> → The number of elements or rows to include in the subset.
<br/>Default value: Count of remaining elements or rows, starting at <code>Index</code>.</li>
<li><kbd>ColumnIndexes</kbd> → An <code>Array</code> of column indices to include in the subset (must be a table-like collection), or <code>Empty</code>.
<br/>Default value: <code>Empty</code> (All columns)</li>
</ul></details>
</td></tr>


<tr><td align="left" valign="top">

```vb
ToCSV(Optional Delimiter As String) As String
```
</td><td align="left" valign="top">
Returns a <code>String</code> representing this instance in <code>CSV</code>-style format.
<details><summary><code>PARAMETERS</code></summary><ul>
<li><kbd>Delimiter</kbd> → Character <code>String</code> used as delimiter between values in a table-like collection.<br/>Default value: <code>","</code></li>
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


<tr><td align="left" valign="top">

```vb
CopyToRange(RangeObject As Excel.Range, Optional ApplyUserLocale As Boolean, Optional WriteHeaders As Boolean) As ICollectionEx
```
</td><td align="left" valign="top">
Copies all elements from this instance to the provided <code>Excel.Range</code> object.
<details><summary><code>PARAMETERS</code></summary><ul>
<li><kbd>RangeObject</kbd> → Target <code>Excel.Range</code> object.</li>
<li><kbd>ApplyUserLocale</kbd> → When <code>True</code>, copies the array of values directly to <code>Range.FormulaR1C1Local</code> instead of <code>Range.Value</code>.
<br/>Default value: <code>True</code></li>
<li><kbd>WriteHeaders</kbd> → When <code>True</code> and only if applicable, first copies a row of header names before copying all it's elements.
<br/>Default value: <code>True</code></li>
</ul></details>
</td></tr>

</tbody>


<thead><tr><th colspan="2">PROCEDURES</th></tr></thead>
<tbody>


<tr><td align="left" valign="top">

```vb
Dispose()
```
</td><td align="left" valign="top">
Tells the class module implementing this interface that this instance won't be needed anymore and the object(s) being wrapped by it can be safely disposed.
<br/>
<em>This shouldn't be necessary in most cases since all objects are automatically destroyed when there's nothing referencing them.</em>
</td></tr>


</tbody>

</table>




