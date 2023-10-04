VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder "ds3x.Collections"
Option Compare Database
Option Explicit

Private pItems() As Variant

Public Property Get Count() As Long: Count = 1 + UBound(pItems): End Property

'@DefaultMember
Public Property Get Items() As Variant()
    Items = pItems
End Property

Public Function Bind(ByRef TargetArray As Variant, Optional ByVal Index As Long = 0, Optional ByVal GetCount As Variant, Optional ByVal ArrayElementSize As Long = VARIANT_SIZE) As ListRange
    Dim sa As SAFEARRAY_1D
    If IsMissing(GetCount) Then GetCount = 1 + UBound(TargetArray) - Index
    With sa
        .cDims = 1
        .cbElements = ArrayElementSize
        .fFeatures = IIf(ArrayElementSize = VARIANT_SIZE, FADF_VARIANT Or FADF_EMBEDDED, FADF_EMBEDDED)
        .pvData = VarPtr(TargetArray(Index))
        .rgsabound0.cElements = GetCount
    End With
    MemLongPtr(VarPtrArr(pItems)) = VarPtr(sa)
    Set Bind = Me
End Function

Private Sub Class_Terminate()
    On Error Resume Next
    MemLongPtr(VarPtrArr(pItems)) = CLngPtr(0)
    On Error GoTo 0
End Sub
