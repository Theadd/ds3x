## **`ScreenLib` Class** <sup><sub><sup> &nbsp; <kbd><code>__UTILITY__</code></kbd></sup></sub></sup>

_Twips/pixel conversion and UI management across multiple monitors._

---

### Key Features
- Twips/pixel conversion utilities
- Multi-monitor coordinate management
- Window positioning and Z-order control
- Cursor style management
- Form styling helpers

---

<table width="100%"><caption>

### **`ScreenLib` API**  
</caption>
<thead><tr><th colspan="2">SCREEN METRICS</th></tr></thead>
<tbody>

<tr><td align="left" valign="top">
```vb
TwipsPerPixelX (Property)
TwipsPerPixelY (Property)
```
</td><td align="left" valign="top">
Conversion factors between twips and pixels
</td></tr>

<tr><td align="left" valign="top">
```vb
GetCursorPosition() As POINTAPI
SetCursorPosition(p As POINTAPI)
```
</td><td align="left" valign="top">
Get/Set mouse position in twips
</td></tr>

</tbody>

<thead><tr><th colspan="2">WINDOW MANAGEMENT</th></tr></thead>
<tbody>

<tr><td align="left" valign="top">
```vb
WindowMoveTo(f, x, y)
WindowSizeTo(f, w, h)
```
</td><td align="left" valign="top">
Position/size Access forms
</td></tr>

<tr><td align="left" valign="top">
```vb
WindowAlwaysOnTop(f)
WindowSendToBack(f)
```
</td><td align="left" valign="top">
Control window Z-order
</td></tr>

</tbody>

<thead><tr><th colspan="2">CURSOR CONTROL</th></tr></thead>
<tbody>

<tr><td align="left" valign="top">
```vb
MouseMoveCursor (Property)
MouseHandCursor (Property)
```
</td><td align="left" valign="top">
Set cursor styles using boolean flags
</td></tr>

</tbody>

<thead><tr><th colspan="2">FORM STYLING</th></tr></thead>
<tbody>

<tr><td align="left" valign="top">
```vb
SetControlAsEnabled(TargetControl, Enable)
```
</td><td align="left" valign="top">
Toggle control enabled state with visual feedback
</td></tr>

</tbody>
</table>
