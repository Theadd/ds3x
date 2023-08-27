## **`ArrayListEx` Class** <sup><sub><sup> &nbsp; <kbd><code>__CHAINABLE__</code></kbd></sup></sub></sup>

_A **lightweight** wrapper around [`ArrayList`](https://learn.microsoft.com/en-us/dotnet/api/system.collections.arraylist?view=netframework-3.5) objects._

Implements __[ICollectionEx](./ICollectionEx.md)__ <sup>(also implements [IList](https://learn.microsoft.com/en-us/dotnet/api/system.collections.ilist?view=netframework-3.5), [ICollection](https://learn.microsoft.com/en-us/dotnet/api/system.collections.icollection?view=netframework-3.5) and [IEnumerable](https://learn.microsoft.com/en-us/dotnet/api/system.collections.ienumerable?view=netframework-3.5) from `mscorlib`).</sup>

---

- `ArrayListEx` instances can hold elements of any type, like objects or arrays, not only value-types.
- All methods returning an `ArrayListEx` are chainable.
- Publicly exposes the inner `ArrayList` instance for extensibility purposes with no code alterning needed.
- `ArrayListEx.Create()` supports creating new instances by directly converting from:
  - Plain **1D Arrays**, **2D Arrays** and **Jagged Arrays** (array of arrays).
  - **Collection**.
  - **ArrayList**, **ArrayListEx**, **ICollection** or any other class implementing [ICollection](https://learn.microsoft.com/en-us/dotnet/api/system.collections.icollection?view=netframework-3.5).
  - **Array2dEx**.
  - **Excel.Range**.
  - **Scripting.Dictionary** and **DictionaryEx**.

---

<table width="100%"><caption>

### **`ArrayListEx` API**  
</caption>
<thead><tr><th colspan="2">FIELDS</th></tr></thead>
<tbody>


<tr><td align="left" valign="top">

```vb
Instance As ArrayList
```
</td><td align="left" valign="top">
The <code>ArrayList</code> object wrapped by this <code>ArrayListEx</code> instance.
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
Gets the number of elements in the <code>ArrayList</code>.
</td></tr>


<tr><td align="left" valign="top">

```vb
Get ColumnCount() As Long
```
</td><td align="left" valign="top">
In table-like instances, gets the number of columns (the count of elements within the first element, which is expected to be an array).
</td></tr>


<tr><td align="left" valign="top">

```vb
Get Item(Index As Variant) As Variant
Let Item(Index As Variant, Value As Variant)
Set Item(Index As Variant, Value As Variant)
```
</td><td align="left" valign="top">
Sets or gets the element at the specified <code>Index</code>. It's the <u>default class member</u>.
<br/>When used as getter, a negative <code>Index</code> is applied to <code>Count</code>, being <code>-1</code> the last element.
<br/>As as setter, assigning to a positive non-existing <code>Index</code> allocates empty elements up to that <code>Index</code>.
</td></tr>


<tr><td align="left" valign="top">

```vb
Get Row(Index As Long) As Variant
```
</td><td align="left" valign="top">
Gets an <code>Array</code> containing the values at the specified row <code>Index</code>. 
</td></tr>


</tbody>



<thead><tr><th colspan="2">FUNCTIONS</th></tr></thead>
<tbody>


<tr><td align="left" valign="top">

```vb
Create(Optional ArrayLike As Variant) As ArrayListEx
```
</td><td align="left" valign="top">
When no parameter is provided, returns a new <code>ArrayListEx</code> instance with an empty <code>ArrayList</code>.
<br/>
Otherwise, returns a new instance containing the values obtained by converting the provided <code>ArrayLike</code> to an <code>ArrayList</code>.
</td></tr>


<tr><td align="left" valign="top">

```vb
Add(Value As Variant) As ArrayListEx
```
</td><td align="left" valign="top">
Adds an element to the <code>ArrayList</code>.
<br/><em>@see also: <code>ArrayListEx.BinaryAdd</code>.</em>
</td></tr>


<tr><td align="left" valign="top">

```vb
AddRange(ArrayLike As Varriant) As ArrayListEx
```
</td><td align="left" valign="top">
Adds all elements found within the provided <code>ArrayLike</code> at the end of the <code>ArrayList</code>. <code>ArrayLike</code> can be any of type supported by <code>ArrayListEx.Create()</code>.<br/>
On table-like collections, the number of columns in both collections should be the same.
</td></tr>


<tr><td align="left" valign="top">

```vb
Insert(Index As Long, Value As Variant) As ArrayListEx
```
</td><td align="left" valign="top">
Inserts an item to the <code>ArrayList</code> instance at specified <code>Index</code>.
</td></tr>


<tr><td align="left" valign="top">

```vb
ToArray() As Variant
```
</td><td align="left" valign="top">
Returns the <code>ArrayList</code> as a plain <code>Array</code>.
</td></tr>


<tr><td align="left" valign="top">

```vb
Join(Target As Array2dEx) As ArrayListEx
```
</td><td align="left" valign="top">
Concatenates all elements of another <code>ArrayListEx</code> instance as additional columns into a new <code>ArrayListEx</code> instance.
</td></tr>


<tr><td align="left" valign="top">

```vb
GetRange(Optional Index As Long, Optional GetCount As Long, Optional ColumnIndexes As Variant) As ArrayListEx
```
</td><td align="left" valign="top">
Returns a new <code>ArrayListEx</code> instance which represents a subset of elements (in a list-like collection) or rows and/or columns (in a table-like collection) from this instance.
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
GetRows(RowIndexes As Variant, Optional ModIndex As Long) As ArrayListEx
```
</td><td align="left" valign="top">
Returns a new <code>ArrayListEx</code> instance containing only those rows specified in the <code>RowIndexes</code> array.
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


<tr><td align="left" valign="top">

```vb
CopyToRange(RangeObject As Excel.Range) As ArrayListEx
```
</td><td align="left" valign="top">
Copies all elements from the <code>ArrayList</code> to the specified <code>Excel.Range</code> object. The <code>Array2dEx.CopyToRange</code> implementation is <b>vastly more efficient</b> than this one.
<details><summary><code>PARAMETERS</code></summary><ul>
<li><kbd>RangeObject</kbd> → Target <code>Excel.Range</code> object.</li>
</ul></details>
</td></tr>


<tr><td align="left" valign="top">

```vb
Remove(Item As Variant) As ArrayListEx
```
</td><td align="left" valign="top">
Removes an element by value from the <code>ArrayList</code>.
</td></tr>


<tr><td align="left" valign="top">

```vb
RemoveAt(Index As Long) As ArrayListEx
```
</td><td align="left" valign="top">
Removes the element at <code>Index</code> from the <code>ArrayList</code>.
</td></tr>


<tr><td align="left" valign="top">

```vb
RemoveRange(Index As Long, RemoveCount As Long) As ArrayListEx
```
</td><td align="left" valign="top">
Removes a range of elements from the <code>ArrayList</code>.
</td></tr>


<tr><td align="left" valign="top">

```vb
RemoveFrom(ArrayLike As Variant) As ArrayListEx
```
</td><td align="left" valign="top">
Removes each element found in the <code>ArrayLike</code> collection from the <code>ArrayList</code>.
</td></tr>


<tr><td align="left" valign="top">

```vb
Clear() As ArrayListEx
```
</td><td align="left" valign="top">
Removes all elements from the <code>ArrayList</code>.
</td></tr>


<tr><td align="left" valign="top">

```vb
Unique() As ArrayListEx
```
</td><td align="left" valign="top">
Removes all duplicated elements from the <code>ArrayList</code>.
<br/>Warning: In large collections, calls to <code>.Unique()</code> can take several seconds. In those cases and when the nature of your data allows a sorted list of elements, you should make use of <code>Binary*</code> methods instead (see them below).
</td></tr>


<tr><td align="left" valign="top">

```vb
Contains(Value As Variant) As Boolean
```
</td><td align="left" valign="top">
Whether a <code>Value</code> is contained in the <code>ArrayList</code> or not, using a <b>STRICT</b> <u>equality comparer</u>, which compares by value and value type <kbd>(Double !== Long)<kbd>.
<br/><em>@see also: <code>ArrayListEx.IndexOf</code> and <code>ArrayListEx.BinarySearch</code></em>.
</td></tr>


<tr><td align="left" valign="top">

```vb
IndexOf(Value As Variant, Optional StartIndex As Long) As Long
```
</td><td align="left" valign="top">
Returns the <code>Index</code> of the first ocurrence of <code>Value</code> within the <code>ArrayList</code> starting at <code>StartIndex</code> (defaults to 0), using a <b>STRICT</b> <u>equality comparer</u>.
<br/><em>@see also: <code>ArrayListEx.BinarySearch</code></em>.
</td></tr>


<tr><td align="left" valign="top">

```vb
Reverse() As ArrayListEx
```
</td><td align="left" valign="top">
Reverses the order of the elements in the <code>ArrayList</code>.
</td></tr>


<tr><td align="left" valign="top">

```vb
Clone() As ArrayListEx
```
</td><td align="left" valign="top">
Returns a new <code>ArrayListEx</code> containing a shallow copy of the elements in this <code>ArrayList</code>.
</td></tr>


<tr><td align="left" valign="top">

```vb
Sort(Optional Comparer As IComparer) As ArrayListEx
```
</td><td align="left" valign="top">
Sorts the elements in this <code>ArrayList</code> using the <b>QuickSort</b> algorithm with the specified <code>Comparer</code>, if provided.
</td></tr>


<tr><td align="left" valign="top">

```vb
BinarySearch(Value As Variant, Optional StartIndex As Long, Optional SearchCount As Long, Optional Comparer As IComparer) As Long
```
</td><td align="left" valign="top">
Uses a binary search algorithm to search the elements in this sorted <code>ArrayList</code> using <code>Comparer</code>, if specified.
<br/>If found, returns the <code>Index</code> of the element. Otherwise, a negative number, which is the bitwise complement of the index of the next element that is larger than <code>Value</code> or, if there is no large element, the bitwise complement of <code>Count</code>.
</td></tr>


<tr><td align="left" valign="top">

```vb
BinaryAdd(Value As Variant, Optional StartIndex As Long, Optional Comparer As IComparer) As Long
```
</td><td align="left" valign="top">
Using the binary search algorithm, adds or inserts an element to the sorted <code>ArrayList</code> at it's corresponding <code>Index</code> position if it doesn't already exists.
<br/>Returning the <code>Index</code> position of the newly added element or the previously existing one.
<br/>To greatly increase performance in adding several elements to large collections, make sure that the values to add are sorted and, supply the return value of each call to this method to the <code>StartIndex</code> parameter of the next call.
</td></tr>


<tr><td align="left" valign="top">

```vb
BinaryToggle(Value As Variant, Optional StartIndex As Long, Optional Comparer As IComparer) As Long
```
</td><td align="left" valign="top">
Using the binary search algorithm, adds or inserts an element to the sorted <code>ArrayList</code> at it's corresponding <code>Index</code> position if it doesn't already exists, otherwise, removes it.
<br/>Returns the <code>Index</code> position of the newly added element or the position where the element was removed.
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


</tbody>


<thead><tr><th colspan="2">PROCEDURES</th></tr></thead>
<tbody>


<tr><td align="left" valign="top">

```vb
Dispose()
```
</td><td align="left" valign="top">
Tells <code>ArrayListEx</code> that this instance won't be needed anymore and it can be safely disposed.
<br/>
<em>This shouldn't be necessary in most cases since all objects are automatically destroyed when there's no reference to them.</em>
</td></tr>


</tbody>


<thead><tr><th colspan="2">STATIC</th></tr></thead>
<tbody>


<tr><td align="left" valign="top">

```vb
CreateBlank(RowsCount As Long, ColumnsCount As Long) As ArrayListEx
```
</td><td align="left" valign="top">
Returns a new <code>ArrayListEx</code> instance with the specified number of rows and columns, containing <code>Empty</code> values.
</td></tr>


<tr><td align="left" valign="top">

```vb
CountElementsIn(ArrayLike As Variant) As Long
```
</td><td align="left" valign="top">
Returns the number of elements or rows in an <code>ArrayLike</code> collection. It can be a plain <code>Array</code> or any object with <code>Count</code>.
</td></tr>


</tbody>

</table>




