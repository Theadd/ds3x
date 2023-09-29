VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CollectionsLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder("ds3x.Libraries")
Option Compare Database
'@PredeclaredId
Option Explicit

'
' Estructura de un array en formato SearchDefinition:
' SearchDefinition = Array(
'    PatternMatching As PatternMatchingType,
'    Tokens() As Variant
' )
' Donde Tokens() es un array de arrays, delimitando el primer nivel del array con el operador lógico AND y el
' segundo nivel con el operador lógico OR. Una posible representación sería: Tokens() = ANDs(ORs()).
'

Public Enum PatternMatchingType
    PartialMatch = 0
    LikePattern = 1
    ExactMatch = 2
End Enum


' USAGE: CollectionsLib.Tokenize("--task=..\Tasks\Entry.json --name=""Initial task"" --exec")   => ['--task=..\Tasks\Entry.json', '--name="Initial task"', '--exec']
Public Property Get Tokenize(ByVal SearchString As String, Optional ByVal Tokenizer As String = " ") As Variant: Tokenize = Split(TokenizeArgs(SearchString, Tokenizer), VBA.Chr$(0)): End Property
Public Property Get ParseToken(ByVal Target As String, Optional ByVal Splitter As String = "=") As Variant: ParseToken = ParseTokenizedArg(Target, Splitter): End Property

Public Property Get AsArray2dEx(ByRef Value As ICollectionEx) As Array2dEx
    Select Case True
        Case (TypeOf Value Is ArraySliceGroup): Set AsArray2dEx = C2ArraySliceGroup(Value).Instance
        Case (TypeOf Value Is Array2dEx): Set AsArray2dEx = Value
        Case (TypeOf Value Is ArrayListEx): Set AsArray2dEx = Array2dEx.Create(Value)
        Case Else: Err.Raise 17, , "Not implemented."
    End Select
End Property

Public Property Get AsArrayListEx(ByRef Value As ICollectionEx) As ArrayListEx
    Select Case True
        Case (TypeOf Value Is ArraySliceGroup): Set AsArrayListEx = ArrayListEx.Create(AsArray2dEx(Value))
        Case (TypeOf Value Is Array2dEx): Set AsArrayListEx = ArrayListEx.Create(Value)
        Case (TypeOf Value Is ArrayListEx): Set AsArrayListEx = Value
        Case Else: Err.Raise 17, , "Not implemented."
    End Select
End Property

Public Property Get AsArraySliceGroup(ByRef Value As ICollectionEx) As ArraySliceGroup
    Select Case True
        Case (TypeOf Value Is ArraySliceGroup): Set AsArraySliceGroup = Value
        Case (TypeOf Value Is Array2dEx): Set AsArraySliceGroup = ArraySliceGroup.Create(Value)
        Case (TypeOf Value Is ArrayListEx): Set AsArraySliceGroup = ArraySliceGroup.Create(AsArray2dEx(Value))
        Case Else: Err.Raise 17, , "Not implemented."
    End Select
End Property


' Devuelve un array estructurado en formato SearchDefinition como el descrito en los comentarios de la parte superior
' de este módulo, partiendo de una SearchString que puede estar compuesta por múltiples criterios de búsqueda y en
' distintos formatos.
'
' Se pueden usar los siguientes carácteres especiales:
'   ' ' (espacio)   -> Operador lógico AND.
'   '|'             -> Operador lógico OR.
'   '"'             -> 1. Interpreta el texto entre comillas dobles como un solo criterio de búsqueda.
'                         Es decir, 'A "B C"' se convertiria en: 'A' AND 'B C'.
'                      2. Pero también se puede utilizar para forzar un match exacto sobre un valor en lugar de parcial
'                         si estos engloban todo el contenido del texto.
'                         Es decir, '"XYZ"' solo hará match si un valor es exactamente 'XYZ', en este caso, no haria
'                         match si por ejemplo el valor es 'AXYZ'.
'
' Si el valor devuelto es utilizado juntamente con TryMatchSearchDefinitionIn, también acepta los carácteres especiales del 'Like':
'   '?'             -> Cualquier carácter.
'   '#'             -> Cualquier dígito (0-9).
'   '*'             -> Cero o más carácteres.
'   '[charlist]     -> Cualquier carácter en la lista.
'   '[!charlist]    -> Cualquier carácter que no esté en la lista.
'
Public Function GenerateSearchDefinition(ByVal SearchString As String) As Variant()
    Dim i As Long, k As Long, nGroups() As String, oGroups() As String, PatternMatching As Long, Tokens() As Variant

    nGroups = Split(TokenizeArgs(SearchString), Chr$(0))
    If UBound(nGroups) = -1 Then
        GenerateSearchDefinition = Array(PartialMatch, Array())
        Exit Function
    End If
    ReDim Tokens(UBound(nGroups))

    ' Retrieving Pattern Matching Type
    If SearchString Like "*[*?#[]*" Then
        PatternMatching = LikePattern
    Else
        If UBound(nGroups) = 0 And Left(SearchString, 1) = """" And Right(SearchString, 1) = """" Then
            PatternMatching = ExactMatch
        Else
            PatternMatching = PartialMatch
        End If
    End If

    For i = LBound(nGroups) To UBound(nGroups)
        oGroups = Split(nGroups(i), "|")
        For k = LBound(oGroups) To UBound(oGroups)
            Select Case PatternMatching
                Case LikePattern
                    oGroups(k) = Replace(oGroups(k), """", "*")
                    oGroups(k) = IIf((oGroups(k) Like "*[!*]*"), "*" & oGroups(k) & "*", "")
                Case ExactMatch, PartialMatch
                    oGroups(k) = Replace(oGroups(k), """", "")
            End Select
        Next k
        Tokens(i) = oGroups
    Next i

    GenerateSearchDefinition = Array(PatternMatching, Tokens)
End Function



' --- PUBLIC MATCHING FUNCTIONS ---

' Devuelve True si los criterios en la SearchDefinition coinciden con Items.
'
' Siendo Items una colección iterable como la propiedad .Fields de un Recordset o simplemente un Array de valores.
Public Function TryMatchSearchDefinitionIn(ByRef SearchDefinition As Variant, ByRef Items As Variant) As Boolean
    Dim sTerm As Variant, orTerm As Variant, isMatch As Boolean

    For Each sTerm In SearchDefinition(1)  '.Tokens
        isMatch = False
        For Each orTerm In sTerm
            If Len(orTerm) > 0 Then
                Select Case SearchDefinition(0)    '.PatternMatching
                    Case PatternMatchingType.PartialMatch
                        isMatch = TryPartialStringMatchIn(CStr(orTerm), Items)
                    Case PatternMatchingType.LikePattern
                        isMatch = TryLikePatternMatchIn(CStr(orTerm), Items)
                    Case PatternMatchingType.ExactMatch
                        isMatch = TryExactStringMatchIn(CStr(orTerm), Items)
                End Select
                If isMatch Then Exit For
            End If
        Next orTerm
        If Not isMatch Then Exit For
    Next sTerm

    TryMatchSearchDefinitionIn = isMatch
End Function

' Similar a TryMatchSearchDefinitionIn pero mediante una SearchCriteria en lugar de una SearchDefinition.
'
' Siendo SearchCriteria una string que delimita los distintos criterios de búsqueda mediante el carácter
' espacio " ", equivalente al operador lógico AND y mediante el caracter de barra vertical "|", equivalente
' al operador lógico OR.
Public Function TryMatchSearchCriteriaIn(ByRef SearchCriteria As Variant, ByRef Items As Variant) As Boolean
    Dim sTerm As Variant, orTerm As Variant, isMatch As Boolean

    For Each sTerm In Split(CStr(SearchCriteria), " ")
        If Len(sTerm) > 0 Then
            isMatch = False
            For Each orTerm In Split(sTerm, "|")
                If Len(orTerm) > 0 Then
                    If TryPartialStringMatchIn(CStr(orTerm), Items) Then
                        isMatch = True
                        Exit For
                    End If
                End If
            Next orTerm
            If Not isMatch Then Exit For
        End If
    Next sTerm

    TryMatchSearchCriteriaIn = isMatch
End Function


' --- PRIVATE SUBMATCHING FUNCTIONS ---

Private Function TryPartialStringMatchIn(ByRef sTerm As String, ByRef Items As Variant) As Boolean
    Dim Item As Variant

    For Each Item In Items
        If VBA.Mid(CStr(Nz(Item, "")), 5, 1) = "-" Then
            If InStr(1, AsCurrentLocale(CStr(Item)), sTerm, vbTextCompare) > 0 Then
                TryPartialStringMatchIn = True
                Exit For
            End If
        Else
            If InStr(1, CStr(Nz(Item, "")), sTerm, vbTextCompare) > 0 Then
                TryPartialStringMatchIn = True
                Exit For
            End If
        End If
    Next Item
End Function

Private Function TryLikePatternMatchIn(ByRef sTerm As String, ByRef Items As Variant) As Boolean
    Dim Item As Variant

    For Each Item In Items
        If VBA.Mid(CStr(Nz(Item, "")), 5, 1) = "-" Then
            If AsCurrentLocale(CStr(Item)) Like sTerm Then
                TryLikePatternMatchIn = True
                Exit For
            End If
        Else
            If CStr(Nz(Item, "")) Like sTerm Then
                TryLikePatternMatchIn = True
                Exit For
            End If
        End If
    Next Item
End Function

Private Function TryExactStringMatchIn(ByRef sTerm As String, ByRef Items As Variant) As Boolean
    Dim Item As Variant

    For Each Item In Items
        If VBA.Mid(CStr(Nz(Item, "")), 5, 1) = "-" Then
            If AsCurrentLocale(CStr(Item)) = sTerm Then
                TryExactStringMatchIn = True
                Exit For
            End If
        Else
            If CStr(Nz(Item, "")) = sTerm Then
                TryExactStringMatchIn = True
                Exit For
            End If
        End If
    Next Item
End Function


' --- UTILITY FUNCTIONS ---

Private Function AsCurrentLocale(ByRef Value As String) As String
    On Error GoTo Fallback

    If VBA.Mid(Value, 8, 1) = "-" Then
        AsCurrentLocale = DateValue(Value)
        Exit Function
    End If

Fallback:
    AsCurrentLocale = Value
End Function

' USAGE: sArgv() = Split(TokenizeArgs("one ""two twoB twoC"" three ""four fourB"" five"), Chr$(0))
Public Function TokenizeArgs(ByRef SearchString As String, Optional ByVal Tokenizer As String = " ") As String
   Dim sArgs As String, sChar As String, nCount As Long, bQuotes As Boolean

   For nCount = 1 To Len(SearchString)
      sChar = Mid$(SearchString, nCount, 1)
      If sChar = Chr$(34) Then
         bQuotes = Not bQuotes
      End If
      If sChar = Tokenizer Then
         If bQuotes Then
            sArgs = sArgs & sChar
         Else
            sArgs = sArgs & Chr$(0)
         End If
      Else
         sArgs = sArgs & sChar
      End If
   Next
   TokenizeArgs = sArgs
End Function

Public Function ParseTokenizedArg(ByVal Target As String, Optional ByVal Splitter As String = "=") As Variant
    Dim t(0 To 1) As Variant, r As Variant

    r = VBA.Split(Target, Splitter, 2)
    t(0) = r(0)
    If UBound(r) = 1 Then
        If (Left(r(1), 1) = """" Or Left(r(1), 1) = "'") And (Left(r(1), 1) = Right(r(1), 1)) Then
            t(1) = VBA.Mid$(r(1), 2, Len(r(1)) - 2)
        Else
            If r(1) = "true" Or r(1) = "false" Then
                t(1) = CBool(r(1))
            Else
                t(1) = r(1)
            End If
        End If
    Else
        t(1) = True
    End If

    ParseTokenizedArg = t
End Function

' Seamlessly supports both vbLf and vbCrLf as line delimiters, excluding those between double quotation marks.
' Returns whether all characters were successfully accounted for in their slices or there has been some leftover characters at the end,
' returning them as a single slice in the IncompleteSlice "out" parameter. It's not being automatically added to the list of slices to
' keep it buffer-friendly, allowing it to be called with string chunks instead of the whole one, assuming that slices might be beheaded
' and it will be attached with it's leftovers in the next call. Reaching the Limit of lines to slice, if specified, also returns true.
'
' Each single line-slice is just a plain array with two elements, Index and GetCount, respectively.
'
' @param SearchString - The string to be sliced into line slices.
' @param Slices - Where the slices are being dropped, byref.
' @param IncompleteSlice - The leftovers of the last slice which didn't properly terminate in a line ending character, or Empty.
' @param [Index] - The 1-based index of the character to start slicing lines from.
' @param [Limit] - The maximum limit of lines to slice, writting the leftovers as a single slice in the IncompleteSlice "out" parameter.
Public Function StringLineSlices(ByRef SearchString As String, ByRef Slices As ArrayListEx, ByRef IncompleteSlice As Variant, Optional ByVal Index As Long = 1, Optional ByVal Limit As Long = -1) As Boolean
    Dim i As Long, iMax As Long, bQuotes As Boolean, r As Long
    Const cQuote As String = """"
    iMax = Len(SearchString)
    IncompleteSlice = Empty
    r = Index

    For i = Index To iMax
        Select Case VBA.Mid$(SearchString, i, 1)
            Case cQuote: bQuotes = Not bQuotes
            Case vbLf
                If Not bQuotes Then
                    If VBA.Mid$(SearchString, i - 1, 1) = vbCr Then
                        Slices.Add Array(r, i - (r + 1))
                    Else
                        Slices.Add Array(r, i - r)
                    End If
                    r = i + 1
                    Limit = Limit - 1
                    If Limit = 0 Then
                        StringLineSlices = True
                        Exit For
                    End If
                End If
        End Select
    Next i
    If r < iMax Then
        IncompleteSlice = Array(r, iMax - r)
    Else
        StringLineSlices = True
    End If
End Function

Public Function StringLineSlicesV2(ByRef SearchString As String, ByRef Slices As ArrayListEx, ByRef IncompleteSlice As Variant, _
                                   Optional ByVal Index As Long = 1, Optional ByVal Limit As Long = -1, _
                                   Optional ByVal cQuote As String = """") As Boolean
    Dim i As Long, iLf As Long, iQt As Long, uStr As String, nQt As Long
    IncompleteSlice = Empty

    i = Index
    iLf = InStr(i, SearchString, vbLf, vbBinaryCompare)
    Do While (iLf > 0)
        uStr = VBA.Mid$(SearchString, i, iLf - i)
        iQt = InStr(1, uStr, cQuote, vbBinaryCompare)
        nQt = Sgn(iQt)
        Do
            Do While (iQt > 0)
                iQt = InStr(1 + iQt, uStr, cQuote, vbBinaryCompare)
                nQt = nQt + Sgn(iQt)
            Loop
            If (nQt Mod 2) = 1 Then
                iQt = iLf - i
                iLf = InStr(1 + i + iQt, SearchString, vbLf, vbBinaryCompare)
                If iLf = 0 Then GoTo BreakLoop
                uStr = VBA.Mid$(SearchString, i, iLf - i)
            End If
        Loop Until ((nQt Mod 2) = 0)
        If VBA.Mid$(uStr, Len(uStr)) = vbCr Then
            Slices.Add Array(i, iLf - (i + 1))
        Else
            Slices.Add Array(i, iLf - i)
        End If
        Limit = Limit - 1
        If Limit = 0 Then
            StringLineSlicesV2 = True
            Exit Do
        End If
        i = iLf + 1
        iLf = InStr(i, SearchString, vbLf, vbBinaryCompare)
    Loop
BreakLoop:
    If i < Len(SearchString) Then
        IncompleteSlice = Array(i, Len(SearchString) - i)
    Else
        StringLineSlicesV2 = True
    End If
End Function


' --- MISC ---

' USAGE: ArrayRange(5, 4) => [5, 6, 7, 8]
Public Function ArrayRange(ByVal Index As Long, ByVal GetCount As Long) As Long()
    Dim i As Long, t() As Long
    ReDim t(0 To GetCount - 1)
    For i = 0 To GetCount - 1
        t(i) = Index + i
    Next i
    MemoryLib.ReassignArrayTo ArrayRange, t
End Function

Public Function ColumnLetter(ByVal ColumnIndex As Long) As String
    Const sChars As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    Dim i As Long

    Do
        ColumnLetter = VBA.Mid$(sChars, (0 ^ i) + ((Int(ColumnIndex / CLng(CLng(26) ^ i)) - (1 - (0 ^ i))) Mod CLng(26)) + (1 - (0 ^ i)), 1) & ColumnLetter
        i = i + 1
    Loop Until ((26 ^ i) + (26 * (1 - (0 ^ (i - 1)))) > ColumnIndex)
End Function

Public Function ValueToString(ByRef Value As Variant) As String
    Select Case VarType(Value)
        Case vbString: ValueToString = """" & VBA.Replace(Value, """", """""") & """"
        Case vbNull: ValueToString = vbNullString
        Case vbDouble, vbDecimal, vbSingle, vbCurrency, vbInteger, vbByte, vbLong, &H14: ValueToString = VBA.Mid$(Str(Value), 2)
        Case vbDate
            Select Case Fix(Value)
                Case Value: ValueToString = VBA.Join(Array(Year(Value), Month(Value), Day(Value)), "-")
                Case 0: ValueToString = CStr(TimeValue(Value))
                Case Else: ValueToString = VBA.Join(Array(VBA.Join(Array(Year(Value), Month(Value), Day(Value)), "-"), CStr(TimeValue(Value))), " ")
            End Select
        Case vbBoolean: ValueToString = IIf(Value, "true", "false")
        Case Else: ValueToString = CStr(Value)
    End Select
End Function

Public Function ValueToLocalString(ByRef Value As Variant) As String
    Select Case VarType(Value)
        Case vbString: ValueToLocalString = """" & VBA.Replace(Value, """", """""") & """"
        Case vbNull: ValueToLocalString = vbNullString
        Case Else: ValueToLocalString = CStr(Value)
    End Select
End Function

Public Function ArrayItem(ByRef ArrayLike As Variant, ByVal Index As Long, Optional ByRef ValueIfNull As Variant) As Variant
    If Not IsMissing(ValueIfNull) Then Assign ArrayItem, ValueIfNull
    On Error Resume Next
    AssignWhenValid ArrayItem, ArrayLike(Index)
    On Error GoTo 0
End Function

Private Sub AssignWhenValid(ByRef Target, ByRef Value As Variant)
    Select Case True
        Case IsObject(Value): If Not Value Is Nothing Then Set Target = Value
        Case IsEmpty(Value), IsNull(Value), IsError(Value): ' Ignore invalid values
        Case Else: Target = Value
    End Select
End Sub

Private Sub Assign(ByRef Target, ByRef Value As Variant)
    If IsObject(Value) Then
        Set Target = Value
    Else
        Target = Value
    End If
End Sub

Private Function C2Array2dEx(ByRef Value As Variant) As Array2dEx: Set C2Array2dEx = Value: End Function
Private Function C2ArraySliceGroup(ByRef Value As Variant) As ArraySliceGroup: Set C2ArraySliceGroup = Value: End Function
