VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "dsComparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Option Base 0

Implements mscorlib.IComparer


Public Function Compare(ByVal X As Variant, ByVal Y As Variant) As Long: Compare = IComparer_Compare(X, Y): End Function

Private Function IComparer_Compare(ByVal X As Variant, ByVal Y As Variant) As Long
    On Error GoTo Fallback
    IComparer_Compare = CLng(Math.Sgn(X - Y))
ExitFunc:
    Exit Function
Fallback:
    On Error GoTo FallbackFail
    If Not IsObject(X) Then
        Select Case X
            Case Is < Y: IComparer_Compare = -1
            Case Is > Y: IComparer_Compare = 1
            Case Else: IComparer_Compare = 0
        End Select
    Else
        IComparer_Compare = X.CompareTo(Y)
    End If
FallbackFail:
    Resume ExitFunc
End Function




