# __API DOCS__

```css
'
'             /* WindowMoveTo Worksheet, (pScrollModX - pScrollX), (pScrollModY - pScrollY) */
'
'                                         (18000)
'              ____________ * ScrollView.InsideWidth    _______ * ScrollView.OutOfBoundsScrollX
'             |____________|___________________________|_______|
'             |________________________________________|
'             |_____________________|                   * Viewport.ViewportContentFullWidth      -> All dsTable Columns Width
'                                    * Worksheet.MaxContentWidthLimit (30000)   -> Worksheet form' max width (~22 inches)
'                          :--------:
'                                    * .MaxFlexSpaceX (6000)
'                          :--::                                                                                           
'                              * .RSafeMarginX (2000 + 1000)                                                                                     
'       |_____________________|                                                                                                                 
'                              * Viewport.ScrollTo (4000)                                                                                           
'         ::--: 
'              * .LSafeMarginX (1000 + 2000)                                                                              
'          |_____________________|                                                                                                                            
'                                                                                                                                      
'                                [pScrollX] [pScrollModX]
'
'
```

### **ScrollView**

### **Viewport**

```vb
' ViewportContentFullWidth = (Total Columns in dsTable * Grid Cell Width)
Property Get ViewportContentFullWidth() As Long
```
````vb
Public Sub ScrollTo(X, Y)

Public Sub RecalculateViewportSizes()
````

 

 

 

### **Worksheet**

```vb
' Private Const pMaxAvailColumns As Long = 16
Property Get MaxAvailColumns() As Long
' Maximum available form width filled with table-related controls (almost 22 inches ~= 55.478cm ~= 31456 twips in my dev environment).
Property Get MaxContentWidthLimit() As Long
' Default Cell/Column width
Property Get GridCellSizeX() As Long
```


---

```css
'                                                                                                                                     '
'                                         0    5    10   15   20   25   30   35   40   45   50   55   60   65   70   75               '
'                                         |____|____|____|____|____|____|____|____|____|____|____|____|____|____|____|                '
'                                         ++++++++++++++++++++++++++++++++++++++                                                      '
'                                         +                                    +                                                      '
'                                         +                                    +                                                      '
'                                         ++++++++++++++++++++++++++++++++++++++                                                      '
'                                  |                                                                                                  '
'                                  | 0    5    10   15   20   25   30   35   40   45   50   55   60   65   70   75                    '
'                                   \|____|____|____|____|____|____|____|____|____|____|____|____|____|____|____|                     '
'                                                                                                                                     '
'    |                                                                              |                                                 '
'    | 0    5    10   15   20   25   30   35   40   45   50   55   60   65   70   75|                                                 '
'     \|____|____|____|____|____|____|____|____|____|____|____|____|____|____|____|/                                                  '
'                                                                                                                                     '
'                                                                                                                                     '
```

```css
'                                                                                                                                     '
'                                         0    5    10   15   20   25   30   35   40   45   50   55   60   65   70   75               '
'                                         |____|____|____|____|____|____|____|____|____|____|____|____|____|____|____|                '
'                                         +++++++++++++++++++++++++++++++++++++++++++++++++++++                                                      '
'                                         +                                                   +                                                      '
'                                         +                                                   +                                                      '
'                                         +++++++++++++++++++++++++++++++++++++++++++++++++++++                                                      '
'                                  |                                                                                                  '
'                                  | 0    5    10   15   20   25   30   35   40   45   50   55   60   65   70   75                    '
'                                   \|____|____|____|____|____|____|____|____|____|____|____|____|____|____|____|                     '
'                                                                                                                                     '
'                   |                                                                              |                                                 '
'                   | 0    5    10   15   20   25   30   35   40   45   50   55   60   65   70   75|                                                 '
'                    \|____|____|____|____|____|____|____|____|____|____|____|____|____|____|____|/                                                  '
'                                                                                                                                     '
'                                                                                                                                     '
```


