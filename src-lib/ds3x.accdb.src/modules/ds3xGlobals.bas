Attribute VB_Name = "ds3xGlobals"
Option Compare Database
Option Explicit
Option Base 0


' --- ACCESS WINDOW HIDE / SHOW ---

Global Const SW_HIDE = 0
Global Const SW_SHOWNORMAL = 1
Global Const SW_SHOWMINIMIZED = 2
Global Const SW_SHOWMAXIMIZED = 3
Global Const SW_SHOW = 5
'Forces a top-level window onto the taskbar when the window is visible.
Public Const WS_EX_APPWINDOW As Long = &H40000

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type BOUNDS
    X As Long
    Y As Long
    W As Long
    h As Long
End Type


' --- Utility Functions ---

' USAGE: Printf("Name: %1, Age: %2", "John", 32) -> "Name: John, Age: 32"
Public Function Printf(ByVal mask As String, ParamArray Tokens() As Variant) As String
    Dim parts() As String: parts = Split(mask, "%")
    Dim i As Long, j As Long, isFound As Boolean, s As String
    
    For i = LBound(parts) + 1 To UBound(parts)
        If LenB(parts(i)) = 0 Then
            parts(i) = "%"
        Else
            isFound = False
            For j = UBound(Tokens) To LBound(Tokens) Step -1
                s = CStr(j + 1)
                If Left$(parts(i), Len(s)) = s Then
                    parts(i) = Tokens(j) & Right$(parts(i), Len(parts(i)) - Len(s))
                    isFound = True
                    Exit For
                End If
            Next j
            If Not isFound Then
                parts(i) = "%" & parts(i)
            End If
        End If
    Next i
    
    Printf = Join(parts, vbNullString)
End Function




