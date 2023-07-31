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



