Attribute VB_Name = "CommonLib"
Option Compare Database
Option Explicit




Public Function KvSequenceToDictionary(ByVal KvSequence As Variant) As Scripting.Dictionary
    Dim i As Long, Entries As New Scripting.Dictionary
    
    For i = 0 To UBound(KvSequence) Step 2
        If IsObject(KvSequence(i + 1)) Then
            Set Entries(KvSequence(i)) = KvSequence(i + 1)
        Else
            Entries(KvSequence(i)) = KvSequence(i + 1)
        End If
    Next i
    
    Set KvSequenceToDictionary = Entries
End Function

Public Function AsKvPairs(ParamArray KvPairs() As Variant) As Scripting.Dictionary
    Dim i As Long, Entry As Scripting.Dictionary, Entries As New Scripting.Dictionary
    
    For i = 0 To UBound(KvPairs)
        Set Entry = KvPairs(i)
        If IsObject(Entry("VALUE")) Then
            Set Entries(Entry("KEY")) = Entry("VALUE")
        Else
            Entries(Entry("KEY")) = Entry("VALUE")
        End If
    Next i
    
    Set AsKvPairs = Entries
End Function

Public Function KvPair(ByRef mKey As String, ByRef mValue As Variant) As Scripting.Dictionary
    Dim Entry As New Scripting.Dictionary
    
    Entry("KEY") = mKey
    If IsObject(mValue) Then
        Set Entry("VALUE") = mValue
    Else
        Entry("VALUE") = mValue
    End If
    
    Set KvPair = Entry
End Function



Public Function GetControlText(ByRef TargetControl As Access.Control) As String
    On Error Resume Next
    GetControlText = TargetControl.Value
    GetControlText = TargetControl.Text
End Function



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''' LOG/REPORT/DEBUG/PROFILING UTILITIES '''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function LogMe(ByVal Message As String, Optional ByVal withTimestamp As Boolean = False) As String
    LogMe = Message
    
    If (DEBUG_MODE_ENABLED) Then
        Debug.Print IIf(withTimestamp, TimeInMS & " - " & Message, Message)
    End If
End Function

Public Function LogCall(ByRef mForm As Access.Form, ByVal mCallName As String, Optional ByVal mMessage As String = "")
    On Error Resume Next
    Debug.Print TimeInMS & " [" & mForm.Name & "] " & mCallName & " - " & mMessage
End Function

Public Function TimeInMS() As String
    TimeInMS = Strings.Format(Now, "HH:nn:ss") & "." & Strings.Right(Strings.Format(Timer, "#0.00"), 2)
End Function

Public Function DateTimeNow() As String
    DateTimeNow = Strings.Format(Now, "dd/mm/yyyy HH:nn:ss") & "." & Strings.Right(Strings.Format(Timer, "#0.000"), 3)
End Function

Public Function MillisecondsToTime(ByVal Ms As Long) As String
    Dim hrs As Long, mins, secs, m As Integer, t As String
    On Error GoTo Finally
    
    hrs = Fix(Ms / 3600000)
    mins = Fix(Ms / 60000) Mod 60
    secs = Fix((Ms Mod 60000) / 1000)
    
    If hrs >= 2 Or (hrs = 1 And mins > 39) Then
        t = t & " " & hrs & "h"
        If mins >= 5 Then t = t & " " & mins & "m"
    Else
        If (hrs = 1 And mins <= 39) Then mins = mins + 60
        
        If mins >= 1 Then
            t = t & " " & mins & "m"
            If mins < 15 Then t = t & " " & secs & "s"
        Else
            If secs >= 10 Then
                t = t & " " & secs & "s"
            Else
                m = IIf(secs >= 1, 10, 100)
                t = t & " " & (Fix((((secs * 1000) + (Ms Mod 1000)) / 1000) * m) / m) & "s"
            End If
        End If
    End If
    
    MillisecondsToTime = VBA.Mid(t, 2)
Finally:
End Function

Public Function SecondsToHMS(ByVal Value As Long) As String
    Dim hrs As Long, mins, secs, m As Integer, t As String
    On Error GoTo Finally
    
    hrs = Fix(Value / 3600)
    mins = Fix(Value / 60) Mod 60
    secs = Fix((Value Mod 60) / 1)
    
    If hrs >= 2 Or (hrs = 1 And mins > 39) Then
        t = t & " " & hrs & "h"
        If mins >= 5 Then t = t & " " & mins & "m"
    Else
        If (hrs = 1 And mins <= 39) Then mins = mins + 60
        
        If mins >= 1 Then
            t = t & " " & mins & "m"
            If mins < 15 Then t = t & " " & secs & "s"
        Else
'            If secs >= 10 Then
                t = t & " " & secs & "s"
'            Else
'                m = IIf(secs >= 1, 10, 100)
'                t = t & " " & (Fix((((secs * 1000) + (Value Mod 1000)) / 1000) * m) / m) & "s"
'            End If
        End If
    End If
    
    SecondsToHMS = VBA.Mid(t, 2)
Finally:
End Function

Public Function XDaysAgo2ColumnName(ByVal X_DAYS_AGO As String) As String
    If X_DAYS_AGO Like "*_DAYS_AGO" Then
        X_DAYS_AGO = Split(X_DAYS_AGO, "_", 2)(0)
        X_DAYS_AGO = UCase(VBA.Format$(DateAdd("d", 0 - Int(Val(X_DAYS_AGO)), Date), "dd MMM", vbMonday))
    End If
    
    XDaysAgo2ColumnName = X_DAYS_AGO
End Function

' Just an ALIAS
Public Function DateFromParts(ByVal dYear As Integer, ByVal dMonth As Integer, ByVal dDay As Integer) As Date
    DateFromParts = DateSerial(dYear, dMonth, dDay)
End Function

Public Function Printf(ByVal mask As String, ParamArray Tokens() As Variant) As String
'    Print Printf("Name: %1, Age: %2", "John", 32)
'    Name: John, Age: 32
    Dim parts() As String: parts = Split(mask, "%")
    Dim i As Long
    Dim j As Long
    Dim isFound As Boolean
    Dim s As String
    '
    'Always ignore first part - covers if mask started or not with %
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





'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''' MATH / NUMERIC UTILITIES ''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function MD5(ByRef Value As String) As String
  Dim Encoder As Object, enc As Object, t As String, pos As Long
  Dim textBytes() As Byte, bytes As Variant  ' Dim bytes() As Byte
  
  Set Encoder = CreateObject("System.Text.UTF8Encoding")
  Set enc = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
  textBytes = Encoder.GetBytes_4(Value)
  bytes = enc.ComputeHash_2((textBytes))
  For pos = 1 To LenB(bytes)
    t = t & LCase(Right("0" & Hex(AscB(MidB(bytes, pos, 1))), 2))
  Next pos
  
  MD5 = t
  Set enc = Nothing
  Set Encoder = Nothing
End Function

Public Function Max(X As Variant, y As Variant) As Variant
  Max = IIf(X > y, X, y)
End Function

Public Function Min(X As Variant, y As Variant) As Variant
   Min = IIf(X < y, X, y)
End Function



