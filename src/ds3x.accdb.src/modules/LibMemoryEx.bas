Attribute VB_Name = "LibMemoryEx"
'@Folder "ds3x.Libraries"
Option Compare Database
Option Explicit

Private Declare PtrSafe Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
Private Declare PtrSafe Function SafeArrayCopyData Lib "oleaut32" (ByRef psaSource As Any, ByRef psaTarget As Any) As Long

Private Declare PtrSafe Function CoTaskMemAlloc Lib "ole32.dll" (ByVal cb As Long) As LongPtr
Private Declare PtrSafe Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As LongPtr)

Public Const FADF_AUTO As Long = &H1            ' An array that is allocated on the stack.
Public Const FADF_VARIANT As Long = &H800       ' An array of VARIANTs.
Public Const FADF_EMBEDDED As Long = &H4        ' An array that is embedded in a structure.
Public Const FADF_FIXEDSIZE As Long = &H10      ' An array that may not be resized or reallocated.
Public Const FADF_HAVEVARTYPE As Long = &H80    ' An array that has a variant type. The variant type can be retrieved with SafeArrayGetVartype.

Public Const INT_SIZE As Long = 2

Public Enum MemoryMoveMode
    MemAllocCopyMemoryMode
    MemAllocMemCopyMode
    CopyMemoryMode
    MemCopyMode
End Enum

Public MemoryMovingMode As MemoryMoveMode


Public Function CreateMemoryCopy(ByRef TargetAddress As LongPtr, ByVal SourceAddress As LongPtr, ByVal ByteCount As Long) As Boolean
    TargetAddress = CoTaskMemAlloc(ByteCount)
    If TargetAddress <> 0 Then
        CopyMemory ByVal TargetAddress, ByVal SourceAddress, ByteCount
        CreateMemoryCopy = True
    End If
End Function

Public Sub FreeMemoryCopy(ByVal TargetAddress As LongPtr)
    If TargetAddress <> 0 Then CoTaskMemFree TargetAddress
End Sub

Public Sub ReassignArrayTo(ByRef Destination As Variant, ByRef Source As Variant)
    MemLongPtr(VarPtrArr(Destination)) = MemLongPtr(VarPtrArr(Source))
    MemLongPtr(VarPtrArr(Source)) = CLngPtr(0)
End Sub

Public Sub ZeroMemory(ByVal TargetAddress As LongPtr, ByVal ByteCount As Long)
    FillMemory ByVal TargetAddress, ByteCount, CByte(0)
End Sub

Public Sub VariantArrayClone(ByVal DestinationAddress As LongPtr, ByVal SourceAddress As LongPtr, ByVal GetCount As Long, Optional ByVal ArrayElementSize As Long = VARIANT_SIZE)
    Dim sASrc As SAFEARRAY_1D, sADst As SAFEARRAY_1D
    With sASrc
        .cDims = 1
        .cbElements = ArrayElementSize
        .fFeatures = IIf(ArrayElementSize = VARIANT_SIZE, FADF_VARIANT, 0)
        .pvData = SourceAddress
        .rgsabound0.cElements = GetCount
    End With
    With sADst
        .cDims = 1
        .cbElements = ArrayElementSize
        .fFeatures = IIf(ArrayElementSize = VARIANT_SIZE, FADF_VARIANT Or FADF_EMBEDDED, FADF_EMBEDDED)
        .pvData = DestinationAddress
        .rgsabound0.cElements = GetCount
    End With
    SafeArrayCopyData ByVal VarPtr(sASrc), ByVal VarPtr(sADst)
    With sASrc
        .pvData = CLngPtr(0)
        .rgsabound0.cElements = 0
    End With
    With sADst
        .cbElements = 2
        .fFeatures = FADF_EMBEDDED
        .pvData = CLngPtr(0)
        .rgsabound0.cElements = 0
    End With
End Sub

#If Win64 Then
    Public Function GetArrayDimsCount(ByRef TargetArray As Variant) As Long
        Dim ptr As LongPtr: ptr = ArrPtr(TargetArray)
        If ptr <> 0 Then GetArrayDimsCount = MemInt(ptr)
    End Function
#Else
    Public Function GetArrayDimsCount(ByRef arr As Variant) As Long
        Const MAX_DIMENSION As Long = 60 'VB limit
        Dim dimension As Long, tempBound As Long
        On Error GoTo FinalDimension
        For dimension = 1 To MAX_DIMENSION
            tempBound = LBound(arr, dimension)
        Next dimension
FinalDimension:
        GetArrayDimsCount = dimension - 1
    End Function
#End If
