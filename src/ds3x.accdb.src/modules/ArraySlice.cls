VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArraySlice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
Option Compare Database
Option Explicit
Option Base 0


Private pSlices As Variant
Private pCount As Long

#If Win64 Then
    Const VARIANT_SIZE As Long = 24
#Else
    Const VARIANT_SIZE As Long = 16
#End If


Public Property Get Count() As Long: Count = pCount: End Property
Public Property Get ColumnCount() As Long: ColumnCount = 1: End Property
Public Property Get Row(ByVal Index As Long) As Variant: Row = Item(Index): End Property
Public Property Get Slices() As Variant: Slices = IIf(IsEmpty(pSlices), Array(), pSlices): End Property

Public Property Let Slices(ByRef Value As Variant)
    Dim v As Variant, r As Long
    If Not IsEmpty(pSlices) Then Err.Raise 70
    
    pSlices = Value
    For Each v In pSlices
        r = r + v(3)
    Next v
    pCount = r
End Property

'@DefaultMember
Public Property Get Item(ByVal Index As Long, Optional ByVal ColumnIndex As Long = 0) As Variant
Attribute Item.VB_UserMemId = 0
    Dim v As Variant, r As Long
    For Each v In pSlices
        If r + v(3) > Index Then
            Item = v(0).Item(v(2) + (Index - r), v(1))  ' + ColumnIndex
            Exit For
        End If
        r = r + v(3)
    Next v
End Property

Public Property Get Slice(Optional ByVal Index As Long = 0, Optional ByVal GetCount As Variant, Optional ByVal ColumnIndex As Long = 0) As ArraySlice
    Dim aX As New ArrayListEx, v As Variant, rMin As Long, rMax As Long, iMin As Long, iMax As Long, sMod As Long
    If IsMissing(GetCount) Then GetCount = pCount - Index
    iMin = Index: iMax = Index + (GetCount - 1)
    Set Slice = New ArraySlice

    For Each v In pSlices
        rMax = rMin + (v(3) - 1)
        
        If iMin <= rMax And iMax <= rMax Then
            '    #rMin                  #rMax
            ' ---|---------|------|-----|--------
            '              #iMin  #iMax
            sMod = iMin - rMin
            aX.Add Array(v(0), v(1), v(2) + sMod, v(3) - sMod - (rMax - iMax))
            Exit For
            
        ElseIf iMin <= rMax And iMax > rMax Then
            '    #rMin             #rMax
            ' ---|---------|-------|-------|----
            '              #iMin           #iMax
            sMod = iMin - rMin
            aX.Add Array(v(0), v(1), v(2) + sMod, v(3) - sMod)
        
        ElseIf iMin <= rMin And iMax <= rMax Then
            '           #rMin                  #rMax
            ' ---|------|-------------|--------|-----
            '    #iMin                #iMax
            aX.Add Array(v(0), v(1), v(2), v(3) - (rMax - iMax))
            Exit For
            
        ElseIf iMin <= rMin And iMax > rMax Then
            '           #rMin         #rMax
            ' ---|------|-------------|--------|-----
            '    #iMin                         #iMax
            aX.Add Array(v(0), v(1), v(2), v(3))
            
        End If
        rMin = rMin + v(3)
    Next v
    
    Slice.Slices = aX.ToArray()
End Property

Public Property Get Instance() As Variant()
    Dim t() As Variant, v As Variant, r As Long, i As Long, k() As Variant, a2dX As Array2dEx
    If pCount = 0 Then
        Instance = Array2dEx.CreateBlank(0, 0).Instance
        Exit Property
    End If
    ReDim t(0 To pCount - 1, 0 To 0)
    
    For Each v In pSlices
        Select Case True
            Case TypeOf v(0) Is Array2dEx
                Set a2dX = C2Array2dEx(v(0))
                MemoryLib.MemCopy VBA.VarPtr(t(r, 0)), VBA.VarPtr(a2dX.Instance(v(2), v(1))), VARIANT_SIZE * v(3)
            Case TypeOf v(0) Is RecordsetEx
                With C2RecordsetEx(v(0))
                    i = CLng(.Instance.AbsolutePosition)
                    .Instance.AbsolutePosition = v(2) + 1
                    k = .Instance.GetRows(v(3), adBookmarkCurrent, Array(v(1)))
                    MemoryLib.MemCopy VBA.VarPtr(t(r, 0)), VBA.VarPtr(k(0, 0)), VARIANT_SIZE * v(3)
                    .Instance.AbsolutePosition = i
                End With
            Case Else
                Err.Raise 17
        End Select
        r = r + v(3)
    Next v
    MemoryLib.FixVariantArray t, pCount, 2
    'Instance = CollectionsLib.GetArrayByRef(t)
    Instance = t
End Property






Public Function Create(Optional ByRef ArrayLike As Variant, Optional ByVal Index As Long = 0, Optional ByVal GetCount As Variant, Optional ByVal ColumnIndex As Long = 0) As ArraySlice
    If Not IsMissing(ArrayLike) Then
        If IsMissing(GetCount) Then GetCount = ArrayListEx.CountElementsIn(ArrayLike) - Index
        If IsObject(ArrayLike) Then
            If TypeOf ArrayLike Is RecordsetEx Then
                Set Create = New ArraySlice
                Create.Slices = ArrayListEx.Create(Array(C2RecordsetEx(ArrayLike).Clone(), ColumnIndex, Index, GetCount)).ToArray()
            ElseIf TypeOf ArrayLike Is ArraySliceGroup Then
                Set Create = C2ArraySliceGroup(ArrayLike).SliceAt(ColumnIndex).Slice(Index, GetCount)
            ElseIf TypeOf ArrayLike Is ArraySlice Then
                Set Create = C2ArraySlice(ArrayLike).Slice(Index, GetCount, ColumnIndex)
            Else
                Set Create = New ArraySlice
                Create.Slices = ArrayListEx.Create().Add(Array(C2Array2dEx(ArrayLike), ColumnIndex, Index, GetCount)).ToArray()
            End If
        Else
            Err.Raise 17, , "Not implemented."
        End If
    End If
End Function

Public Function Add(ByRef Target As ArraySlice) As ArraySlice
    Set Add = New ArraySlice
    Add.Slices = ArrayListEx.Create(Slices).AddRange(Target.Slices).ToArray()
End Function

' Returns a new ArraySlice containing all sub slices from AllRanges.
'
' @param AllRanges - An ArrayListEx containing all Array(Index, GetCount) elements.
Public Function SliceAllRanges(ByVal AllRanges As ArrayListEx) As ArraySlice
    Dim aX As New ArrayListEx, v As Variant, rMin As Long, rMax As Long, iMin As Long, iMax As Long, sMod As Long, vRange As Variant

    For Each vRange In AllRanges
        iMin = vRange(0)
        iMax = iMin + (vRange(1) - 1)
        
        For Each v In pSlices
            rMax = rMin + (v(3) - 1)
            
            If iMin <= rMax And iMax <= rMax Then
                '    #rMin                  #rMax
                ' ---|---------|------|-----|--------
                '              #iMin  #iMax
                sMod = iMin - rMin
                aX.Add Array(v(0), v(1), v(2) + sMod, v(3) - sMod - (rMax - iMax))
                Exit For
                
            ElseIf iMin <= rMax And iMax > rMax Then
                '    #rMin             #rMax
                ' ---|---------|-------|-------|----
                '              #iMin           #iMax
                sMod = iMin - rMin
                aX.Add Array(v(0), v(1), v(2) + sMod, v(3) - sMod)
            
            ElseIf iMin <= rMin And iMax <= rMax Then
                '           #rMin                  #rMax
                ' ---|------|-------------|--------|-----
                '    #iMin                #iMax
                aX.Add Array(v(0), v(1), v(2), v(3) - (rMax - iMax))
                Exit For
                
            ElseIf iMin <= rMin And iMax > rMax Then
                '           #rMin         #rMax
                ' ---|------|-------------|--------|-----
                '    #iMin                         #iMax
                aX.Add Array(v(0), v(1), v(2), v(3))
                
            End If
            rMin = rMin + v(3)
        Next v
        rMin = 0
    Next vRange
    
    Set SliceAllRanges = New ArraySlice
    SliceAllRanges.Slices = aX.ToArray()
End Function



Private Function C2ArraySlice(ByRef Value As Variant) As ArraySlice: Set C2ArraySlice = Value: End Function
Private Function C2Array2dEx(ByRef Value As Variant) As Array2dEx: Set C2Array2dEx = Value: End Function
Private Function C2RecordsetEx(ByRef Value As Variant) As RecordsetEx: Set C2RecordsetEx = Value: End Function
Private Function C2ArraySliceGroup(ByRef Value As Variant) As ArraySliceGroup: Set C2ArraySliceGroup = Value: End Function







