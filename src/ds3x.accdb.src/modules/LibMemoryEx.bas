Attribute VB_Name = "LibMemoryEx"
'@Folder "ds3x.Libraries"
Option Compare Database
Option Private Module
Option Explicit


Private Declare PtrSafe Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
Private Declare PtrSafe Function SafeArrayCopyData Lib "oleaut32" (ByRef psaSource As Any, ByRef psaTarget As Any) As Long

Public Const FADF_AUTO As Long = &H1            ' An array that is allocated on the stack.
Public Const FADF_VARIANT As Long = &H800       ' An array of VARIANTs.
Public Const FADF_EMBEDDED As Long = &H4        ' An array that is embedded in a structure.

Public Const INT_SIZE As Long = 2


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
