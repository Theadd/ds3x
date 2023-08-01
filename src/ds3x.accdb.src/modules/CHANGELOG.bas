Attribute VB_Name = "CHANGELOG"
Option Compare Database
Option Explicit
Option Base 0

'Public G_DSLIVEED As dsLiveEd
' --- XLS_DEV - FILES CHANGED ---

' CommonLib
'   @ MillisecondsToTime(ByVal Ms As Long) As String
'   @ SecondsToHMS(ByVal Value As Long) As String
'   @ XDaysAgo2ColumnName
'
'
'
'
'
'

Public Sub TEST_LoadFromRangeMemoryUsage()
    ShowAvailableVirtualMemory
    Dim xlS As xlSheetsEx, dsT As dsTable, dsT2 As dsTable, aX As ArrayListEx, bX As ArrayListEx, a2X As Array2dEx, b2X As Array2dEx
    Dim xlRange As Excel.Range, xlRange2 As Excel.Range, vA1 As Variant, vA2 As Variant, vA3 As Variant
    
    ShowAvailableVirtualMemory
    Set xlS = xlSheetsEx.Create(, "L:\7300\dsEx\docs\SAMPLE_INSYSAN_ANSI_FULL.csv")
    ShowAvailableVirtualMemory
    Set xlRange = xlS.UsedRange
    ' Set xlRange = xlS.UsedRange.Resize(29900)
    ' Set xlRange2 = xlS.UsedRange.Resize(8100)
    ShowAvailableVirtualMemory
    
    With xlRange
        Set xlRange2 = .Range(.Cells(2, 1), .Cells(.Rows.Count, .Columns.Count))
        ShowAvailableVirtualMemory
        Stop
        Set aX = ArrayListEx.Create(xlRange2)
        ShowAvailableVirtualMemory
        Stop
    End With
    
    Set xlRange = Nothing
    Set xlRange2 = Nothing
    xlS.Dispose
    Set xlS = Nothing
    ShowAvailableVirtualMemory
    Stop
    ArrayListEx.Dispose
    ArrayListEx.Unbind
    ShowAvailableVirtualMemory
    Stop
    Set aX = Nothing
    ShowAvailableVirtualMemory
    Stop
    
'    With xlRange
'        Set dsT = dsTable.Create(ArrayListEx.Create(.Range(.Cells(2, 1), .Cells(.Rows.Count, .Columns.Count))))
'        ShowAvailableVirtualMemory
'        dsT.SetHeaders ArrayListEx.Create(.Range(.Cells(1, 1), .Cells(1, .Columns.Count)))(0)
'    End With
'    ShowAvailableVirtualMemory
'
'    Stop
'
'    With xlRange2
'        Set dsT2 = dsTable.Create(ArrayListEx.Create(.Range(.Cells(2, 1), .Cells(.Rows.Count, .Columns.Count))))
'        ShowAvailableVirtualMemory
'        dsT2.SetHeaders ArrayListEx.Create(.Range(.Cells(1, 1), .Cells(1, .Columns.Count)))(0)
'    End With
'    ShowAvailableVirtualMemory
'
'    Stop
    
End Sub

Public Sub TEST_ArrayListExDispose()
    Dim aX As ArrayListEx, bX As ArrayListEx, cX As ArrayListEx, i As Long
    
    Set aX = ArrayListEx.Create()
    
    For i = 0 To 6
        aX.Add Array(i, i + 1, i + 2, "ID = " & CStr(i))
    Next i
    
    Debug.Print JSON.Stringify(aX)
    
    ' Set bX = aX.GetRange(0, 2, Array(0, 1, 2, 3))
    Set bX = aX.GetRange(0, 2)
    Debug.Print JSON.Stringify(aX)
    bX.Dispose
    Debug.Print JSON.Stringify(aX)
    Set bX = Nothing
    Debug.Print JSON.Stringify(aX)
    ' Set cX = aX.GetRange(2, 2, Array(0, 1, 2, 3))
'    Set cX = ArrayListEx.Create(aX.GetRange(2, 2))
'    cX.Clear
    Set bX = ArrayListEx.Create(aX.GetRange(2, 2))
    Debug.Print JSON.Stringify(aX)
    bX.Dispose
    Debug.Print JSON.Stringify(aX)
    Set bX = Nothing
    Debug.Print JSON.Stringify(aX)
End Sub

Public Sub TEST_OpenTextUTF8()
    Dim xlSheet As xlSheetsEx, dX As DictionaryEx

    Set dX = DictionaryEx.Create() _
                         .Add("UpdateLinks", False) _
                         .Add("ReadOnly", False) _
                         .Add("Local", True) _
                         .Add("UTF8", True) _
                         .Add("NoTextQualifier", True)
                         
    Set xlSheet = xlSheetsEx.Create(, "L:\7300\dsEx\docs\SAMPLE_INSYSAN.csv", dX)
    xlSheet.WindowVisibility = True
    xlSheet.Dispose
    Set xlSheet = Nothing
    Set dX = Nothing
End Sub



