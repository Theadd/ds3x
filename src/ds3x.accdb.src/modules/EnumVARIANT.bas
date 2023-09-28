Attribute VB_Name = "EnumVARIANT"
' Source: https://github.com/Kr00l/VBCCR - VB Common Controls Replacement Library
' MIT License
' Copyright (c) 2012-present Krool
'
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software
' and associated documentation files (the "Software"), to deal in the Software without restriction,
' including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense,
' and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so,
' subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial
' portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT
' LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
' IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
' WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
' SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
Option Compare Database
Option Explicit

#If Win64 Then
    Private Const NULL_PTR As LongPtr = 0
#Else
    Private Const NULL_PTR As Long = 0
#End If

Private Type TEnumVARIANT
    VTable As LongPtr
    RefCount As Long
    Enumerable As Object
    Index As Long
    Count As Long
End Type

Private Type IEnumGUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Declare PtrSafe Sub CoTaskMemFree Lib "ole32" (ByVal hMem As LongPtr)
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare PtrSafe Function CoTaskMemAlloc Lib "ole32" (ByVal cBytes As Long) As LongPtr
Private Declare PtrSafe Function VariantCopyToPtr Lib "oleaut32" Alias "VariantCopy" (ByVal pvargDest As LongPtr, ByRef pvargSrc As Variant) As Long

Private Const E_INVALIDARG As Long = &H80070057
Private Const E_NOTIMPL As Long = &H80004001
Private Const E_NOINTERFACE As Long = &H80004002
Private Const E_POINTER As Long = &H80004003
Private Const S_FALSE As Long = &H1
Private Const S_OK As Long = &H0

Private VTableEnumVARIANT(0 To 6) As LongPtr


Public Function GetNewEnum(ByVal This As Object, ByVal Upper As Long, ByVal Lower As Long) As IEnumVARIANT
    Dim data As TEnumVARIANT
    With data
        .VTable = GetVTableEnumVARIANT()
        .RefCount = 1
        Set .Enumerable = This
        .Index = Lower
        .Count = Upper
        Dim hMem As LongPtr
        hMem = CoTaskMemAlloc(LenB(data))
        If hMem <> NULL_PTR Then
            CopyMemory ByVal hMem, data, LenB(data)
            CopyMemory ByVal VarPtr(GetNewEnum), hMem, PTR_SIZE
            CopyMemory ByVal VarPtr(.Enumerable), NULL_PTR, PTR_SIZE
        End If
    End With
End Function

Private Function GetVTableEnumVARIANT() As LongPtr
    If VTableEnumVARIANT(0) = NULL_PTR Then
        VTableEnumVARIANT(0) = ProcPtr(AddressOf IEnumVARIANT_QueryInterface)
        VTableEnumVARIANT(1) = ProcPtr(AddressOf IEnumVARIANT_AddRef)
        VTableEnumVARIANT(2) = ProcPtr(AddressOf IEnumVARIANT_Release)
        VTableEnumVARIANT(3) = ProcPtr(AddressOf IEnumVARIANT_Next)
        VTableEnumVARIANT(4) = ProcPtr(AddressOf IEnumVARIANT_Skip)
        VTableEnumVARIANT(5) = ProcPtr(AddressOf IEnumVARIANT_Reset)
        VTableEnumVARIANT(6) = ProcPtr(AddressOf IEnumVARIANT_Clone)
    End If
    GetVTableEnumVARIANT = VarPtr(VTableEnumVARIANT(0))
End Function

Private Function IEnumVARIANT_QueryInterface(ByRef This As TEnumVARIANT, ByRef IID As IEnumGUID, ByRef pvObj As LongPtr) As Long
    If VarPtr(pvObj) = NULL_PTR Then
        IEnumVARIANT_QueryInterface = E_POINTER
        Exit Function
    End If
    ' IID_IEnumVARIANT = {00020404-0000-0000-C000-000000000046}
    If IID.Data1 = &H20404 And IID.Data2 = &H0 And IID.Data3 = &H0 Then
        If IID.Data4(0) = &HC0 And IID.Data4(1) = &H0 And IID.Data4(2) = &H0 And IID.Data4(3) = &H0 _
        And IID.Data4(4) = &H0 And IID.Data4(5) = &H0 And IID.Data4(6) = &H0 And IID.Data4(7) = &H46 Then
            pvObj = VarPtr(This)
            IEnumVARIANT_AddRef This
            IEnumVARIANT_QueryInterface = S_OK
        Else
            IEnumVARIANT_QueryInterface = E_NOINTERFACE
        End If
    Else
        IEnumVARIANT_QueryInterface = E_NOINTERFACE
    End If
End Function

Private Function IEnumVARIANT_AddRef(ByRef This As TEnumVARIANT) As Long
    This.RefCount = This.RefCount + 1
    IEnumVARIANT_AddRef = This.RefCount
End Function

Private Function IEnumVARIANT_Release(ByRef This As TEnumVARIANT) As Long
    This.RefCount = This.RefCount - 1
    IEnumVARIANT_Release = This.RefCount
    If IEnumVARIANT_Release = 0 Then
        Set This.Enumerable = Nothing
        CoTaskMemFree VarPtr(This)
    End If
End Function

Private Function IEnumVARIANT_Next(ByRef This As TEnumVARIANT, ByVal VntCount As Long, ByVal VntArrPtr As LongPtr, ByRef pcvFetched As Long) As Long
    If VntArrPtr = NULL_PTR Then
        IEnumVARIANT_Next = E_INVALIDARG
        Exit Function
    End If
    On Error GoTo CATCH_EXCEPTION
    Const VARIANT_CB As Long = 16
    Dim Fetched As Long
    With This
        Do Until .Index > .Count
            VariantCopyToPtr VntArrPtr, .Enumerable(.Index)
            .Index = .Index + 1
            Fetched = Fetched + 1
            If Fetched = VntCount Then Exit Do
            VntArrPtr = UnsignedAdd(VntArrPtr, VARIANT_CB)
        Loop
    End With
    If Fetched = VntCount Then
        IEnumVARIANT_Next = S_OK
    Else
        IEnumVARIANT_Next = S_FALSE
    End If
    If VarPtr(pcvFetched) <> NULL_PTR Then pcvFetched = Fetched
    Exit Function
CATCH_EXCEPTION:
    If VarPtr(pcvFetched) <> NULL_PTR Then pcvFetched = 0
    IEnumVARIANT_Next = E_NOTIMPL
End Function

Private Function IEnumVARIANT_Skip(ByRef This As TEnumVARIANT, ByVal VntCount As Long) As Long
    IEnumVARIANT_Skip = E_NOTIMPL
End Function

Private Function IEnumVARIANT_Reset(ByRef This As TEnumVARIANT) As Long
    IEnumVARIANT_Reset = E_NOTIMPL
End Function

Private Function IEnumVARIANT_Clone(ByRef This As TEnumVARIANT, ByRef ppEnum As IEnumVARIANT) As Long
    IEnumVARIANT_Clone = E_NOTIMPL
End Function

Private Function ProcPtr(ByVal Address As LongPtr) As LongPtr
    ProcPtr = Address
End Function

Private Function UnsignedAdd(ByVal Start As LongPtr, ByVal Incr As LongPtr) As LongPtr
    #If Win64 Then
        UnsignedAdd = ((Start Xor &H8000000000000000^) + Incr) Xor &H8000000000000000^
    #Else
        UnsignedAdd = ((Start Xor &H80000000) + Incr) Xor &H80000000
    #End If
End Function

