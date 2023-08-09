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

Private pCachedTracks As ArrayListEx
' The current ColumnsToLargeChangeTrack value used in CachedTracks, CachedTracks resets when this value changes
Private pTrackColumnSizesInCache As Long
Const NumPagesInLargeChangeRows As Long = 5
Const PageSize As Long = 10
Const PageCount As Long = 10


Private Type TViewportState
    ScrollPosX As Long
    ScrollPosY As Long
    ' Index of current visible track in Viewport
    TrackIndex As Long
    ' Index of current visible page in Viewport
    PageIndex As Long
    ' Index of the first visible column in **Table**
    FirstVisibleColumn As Long
    ' Index of the first visible row in **Table**
    FirstVisibleRow As Long
    ' The distance between the start of the first visible column to the viewport left edge (must be less than GridCellSizeX)
    FirstVisibleColumnPositionModX As Long
    ' Index of the first visible column relative to current visible **Track**
    FirstVisibleColumnInTrack As Long
    ' Index of the first visible row relative to current visible **Page**
    FirstVisibleRowInPage As Long
    ' Number of columns as the distance between track switching
    ColumnsToLargeChangeTrack As Long
    ' The distance from the current track left edge to the viewport left edge
    TrackPositionModX As Long
End Type

Private this As TViewportState
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


