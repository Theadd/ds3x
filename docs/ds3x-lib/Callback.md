## **`Callback` Class**

_A <kbd>`pass function as parameter`</kbd>-like feature on steroids with additional support for executing `javascript` code and `Filter`/`Map`/`Reduce` (`Where`/`Select`/`Aggregate` equivalents on `.NET`) calls on callback results._

---

### **API Overview**

```vb

```


<table width="100%"><caption>

### **`Callback` API**  
</caption>
<thead><tr><th colspan="2">EVENTS</th></tr></thead>
<tbody>


<tr><td align="left" valign="top">

```vb
Event ExecuteComplete(TargetCallable As Callback)
```
</td><td align="left" valign="top">
The <code>Event</code> raised after <code>Callback</code>'s call to <code>.Execute</code> or <code>.ExecuteOnArray</code> is completed.
</td></tr>

</tbody>

<thead><tr><th colspan="2">FIELDS</th></tr></thead>
<tbody>


<tr><td align="left" valign="top">

```vb
Result As Variant
```
</td><td align="left" valign="top">
A <code>Variant</code> containing the execution result of the <code>Callback</code>.
</td></tr>

</tbody>


<thead><tr><th colspan="2">FUNCTIONS</th></tr></thead>
<tbody>


<tr><td align="left" valign="top">

```vb
Create(Optional ObjectInstance As Object, Optional CallableName As String, Optional CallableType As VBA.VbCallType) As Callback
'Callback.Create(, "MillisecondsToTime")
```
</td><td align="left" valign="top">
Returns a new <code>Callback</code> instance pointing to the specified callable.
<details><summary><code>PARAMETERS</code></summary><ul>
<li><kbd>ObjectInstance</kbd> → It can be a <code>String</code> or an <code>Array</code> in <code>SearchDefinition</code> format.
<br/><em>@see: <code>CollectionsLib.GenerateSearchDefinition</code></em>.</li>
<li><kbd>CallableName</kbd> → A <code>Bookmark</code> of the previous record to start searching for.
<br/>Default value: <code>vbNullString</code>.</li>
<li><kbd>CallableType</kbd> → A <code>VBA.VbCallType</code> value specifying the callable type.
<br/>Default value: <code>VBA.VbCallType.VbMethod</code></li>
</ul></details>
</td></tr>


<tr><td align="left" valign="top">

```vb
Bind(Optional ObjectInstance As Object, Optional CallableName As String, Optional CallableType As VBA.VbCallType) As Callback
```
</td><td align="left" valign="top">
Returns a <code>String</code> representing the specified <code>Value</code> in <b>JSON</b> format.
</td></tr>


<tr><td align="left" valign="top">

```vb
Unbind() As Callback
```
</td><td align="left" valign="top">
Returns the parsed <b>JSON</b> <code>String</code> as a <code>Variant</code> value-type or object-type.
</td></tr>


<tr><td align="left" valign="top">

```vb
Execute(ParamArray vArgs() As Variant) As Callback
```
</td><td align="left" valign="top">
Returns the parsed <b>JSON</b> <code>String</code> as a <code>Variant</code> value-type or object-type.
</td></tr>


<tr><td align="left" valign="top">

```vb
ExecuteOnArray(vArgs() As Variant) As Callback
```
</td><td align="left" valign="top">
Returns the parsed <b>JSON</b> <code>String</code> as a <code>Variant</code> value-type or object-type.
</td></tr>


<tr><td align="left" valign="top">

```vb
Filter(PredicateFunction As String) As Callback
```
</td><td align="left" valign="top">
Just like <code>Array.Where()</code> in <b>.NET</b> or <code>Array.Filter()</code> in <b>JavaScript</b>.
<details><summary><code>EXAMPLE</code></summary>

```vb
Callback.Create()(Array(1, 3, 5, 2, 4, 6)).Filter("x => x < 5")   '.Result contains: [1, 3, 2, 4]
?JSON.Stringify(Callback.Create()(Array(1, 3, 5, 2, 4, 6)).Filter("x => x < 5").Result)
```
</details>
</td></tr>


</tbody>

</table>

