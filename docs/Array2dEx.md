## **`Array2dEx` Class** <sup><sub><sup> &nbsp; <kbd><code>__IMMUTABLE__</code></kbd> &nbsp; <kbd><code>__CHAINABLE__</code></kbd></sup></sub></sup>

_A **lightweight** wrapper around `VBA`'s built-in `2D Array`._

Implements __[ICollectionEx](./ICollectionEx.md)__.

---

- `Array2dEx` instances can hold items of any type, like objects or arrays, not only value-types.
- `Array2dEx` instances are __immutable__.
  - Any call that could transform the array, returns a new instance instead.
  - It's inner `2D Array` is stored as a `Variant` instead of an actual array, forcing it to be a value-type so that direct manipulation has no effect.
- All methods returning an `Array2dEx` are chainable.
- Publicly exposes it's inner `2D Array` for extensibility purposes with no need to alter it's code.
- With `Array2dEx.Create()` supports creating new instances directly converting from:
  - Plain **2D Arrays** and **Jagged Arrays** (array of arrays).
  - **ArrayList** and **ArrayListEx**.
  - **Excel.Range**.
  - **ADODB.Recordset**, **RecordsetEx** and **ADODB.Fields**.

---

<table width="100%"><caption>

### **`Array2dEx` API**  
</caption>
<thead><tr><th colspan="2">FIELDS</th></tr></thead>
<tbody>


<tr><td align="left" valign="top">

```vb
Instance As Variant
```
</td><td align="left" valign="top">
The <code>2D Array</code> wrapped by this <code>Array2dEx</code> instance.<br/>
<code>VBA</code> <code>Array</code>s are always passed <em>by reference</em> but <code>Array2dEx</code> stores it as a <code>Variant</code> value instead, so direct manipulation of this array has no effect.
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
Gets the number of rows in this collection.
</td></tr>


<tr><td align="left" valign="top">

```vb
Get ColumnCount() As Long
```
</td><td align="left" valign="top">
Gets the number of columns in this collection.
</td></tr>


<tr><td align="left" valign="top">

```vb
Get Row(Index As Long) As Variant
```
</td><td align="left" valign="top">
Gets an <code>Array</code> containing all the values at the specified row <code>Index</code>. 
</td></tr>


</tbody>



<thead><tr><th colspan="2">FUNCTIONS</th></tr></thead>
<tbody>


<tr><td align="left" valign="top">

```vb
CreateBlank(RowsCount As Long, ColumnsCount As Long) As Array2dEx
```
</td><td align="left" valign="top">
Returns a new table-like <code>ICollectionEx</code> instance with the specified number of rows and columns, containing <code>Empty</code> values.
</td></tr>


<tr><td align="left" valign="top">

```vb
Create(Optional ArrayLike As Variant) As Array2dEx
```
</td><td align="left" valign="top">
When no parameter is provided, returns a new <code>Array2dEx</code> empty instance.
<br/>
Otherwise, returns a new instance containing the values obtained by converting the provided <code>ArrayLike</code> to a <code>2D Array</code>.
</td></tr>


<tr><td align="left" valign="top">

```vb
Join(Target As Array2dEx) As Array2dEx
```
</td><td align="left" valign="top">
Concatenates all elements of another <code>Array2dEx</code> instance as additional columns into a new <code>Array2dEx</code> instance.
</td></tr>


<tr><td align="left" valign="top">

```vb
GetRange(Optional Index As Long, Optional GetCount As Long, Optional ColumnIndexes As Variant) As Array2dEx
```
</td><td align="left" valign="top">
Returns a new <code>Array2dEx</code> instance which represents a subset of rows and/or columns from this instance.
<details><summary><code>PARAMETERS</code></summary><ul>
<li><kbd>Index</kbd> → Index of the first row to include in the subset.
<br/>Default value: <code>0</code></li>
<li><kbd>GetCount</kbd> → The number of rows to include in the subset.
<br/>Default value: Count of remaining rows when starting at <code>Index</code>.</li>
<li><kbd>ColumnIndexes</kbd> → An <code>Array</code> of column indices to include in the subset or <code>Empty</code>.
<br/>Default value: <code>Empty</code> (All columns)</li>
</ul></details>
</td></tr>


<tr><td align="left" valign="top">

```vb
GetRows(RowIndexes As Variant, Optional ModIndex As Long) As Array2dEx
```
</td><td align="left" valign="top">
Returns a new <code>Array2dEx</code> instance containing only those rows specified in the <code>RowIndexes</code> array.
<details><summary><code>PARAMETERS</code></summary><ul>
<li><kbd>RowIndexes</kbd> → An <code>Array</code> of row indices.</li>
<li><kbd>ModIndex</kbd> → A signed integer to shift values in the <code>RowIndexes</code> array.
<br/>Default value: <code>0</code></li>
</ul></details>
</td></tr>


<tr><td align="left" valign="top">

```vb
ToCSV(Optional Delimiter As String) As String
```
</td><td align="left" valign="top">
Returns a <code>String</code> representing this instance in <code>CSV</code>-style format.
<details><summary><code>PARAMETERS</code></summary><ul>
<li><kbd>Delimiter</kbd> → Character <code>String</code> used as delimiter between row values.<br/>Default value: <code>","</code></li>
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
CopyToRange(RangeObject As Excel.Range, Optional ApplyUserLocale As Boolean) As Array2dEx
```
</td><td align="left" valign="top">
Copies all elements from this instance to the provided <code>Excel.Range</code> object.
<details><summary><code>PARAMETERS</code></summary><ul>
<li><kbd>RangeObject</kbd> → Target <code>Excel.Range</code> object.</li>
<li><kbd>ApplyUserLocale</kbd> → When <code>True</code>, copies the array of values directly to <code>Range.FormulaR1C1Local</code> instead of <code>Range.Value</code>.
<br/>Default value: <code>True</code></li>
</ul></details>
</td></tr>


<tr><td align="left" valign="top">

```vb
Transpose() As String
```
</td><td align="left" valign="top">
Returns the transposed values in a new <code>Array2dEx</code> instance. Rows become columns, columns become rows.
</td></tr>


<tr><td align="left" valign="top">

```vb
Is2dArray(ArrayLike As Variant) As Boolean
```
</td><td align="left" valign="top">
Returns whether the provided <code>ArrayLike</code> is a plain <code>2D Array</code> or not.
</td></tr>


<tr><td align="left" valign="top">

```vb
Is1dArray(ArrayLike As Variant) As Boolean
```
</td><td align="left" valign="top">
Returns whether the provided <code>ArrayLike</code> is a plain <code>1D Array</code> or not.
</td></tr>


</tbody>


<thead><tr><th colspan="2">PROCEDURES</th></tr></thead>
<tbody>


<tr><td align="left" valign="top">

```vb
Dispose()
```
</td><td align="left" valign="top">
Tells <code>Array2dEx</code> that this instance won't be needed anymore and it can be safely disposed.
<br/>
<em>This shouldn't be necessary in most cases since all objects are automatically destroyed when there's nothing referencing them.</em>
</td></tr>


</tbody>

</table>




