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
' @REQUIRES:
'   1. A reference to "Microsoft VBScript Regular Expressions 5.5"
'

Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)

Private Const VT_BYREF As Long = &H4000
#If Win64 Then
Private Const PTR_SIZE As Long = 8
#Else
Private Const PTR_SIZE As Long = 4
#End If

Private Const JSON_SERIALIZER As String = "'object'!=typeof JSON&&(JSON={}),function(){'use strict';var t,e=/[\\""\u0000-\u001f\u007f-\u009f\u00ad\u0600-\u0604\u070f\u17b4\u17b5\u200c-\u200f\u2028-\u202f\u2060-\u206f\ufeff\ufff0-\uffff]/g;function f(t){return t<10?'0'+t:t}function this_value(){return this.valueOf()}function quote(u){return e.lastIndex=0,e.test(u)?'""'+u.replace(e,(function(e){var u=t[e];return'string'==typeof u?u:'\\u'+('0000'+e.charCodeAt(0).toString(16)).slice(-4)}))+'""':'""'+u+'""'}function str(t,e){var u,o,n,r,i,s=e[t];switch(s&&'object'==typeof s&&'function'==typeof s.toJSON&&(s=s.toJSON(t)),typeof s){case'string':return quote(s);case'number':case'boolean':case'null':return String(s);case'object':if(!s)return'null';if(i=[],'[object Array]'===Object.prototype.toString.apply(s)){" & _
                                          "for(r=s.length,u=0;u<r;u+=1)i[u]=str(u,s)||'null';return 0===i.length?'[]':'['+i.join(',')+']'}for(o in s)Object.prototype.hasOwnProperty.call(s,o)&&(n=str(o,s))&&i.push(quote(o)+':'+n);return 0===i.length?'{}':'{'+i.join(',')+'}'}}'function'!=typeof Date.prototype.toJSON&&(Date.prototype.toJSON=function(){return isFinite(this.valueOf())?this.getUTCFullYear()+'-'+f(this.getUTCMonth()+1)+'-'+f(this.getUTCDate())+'T'+f(this.getUTCHours())+':'+f(this.getUTCMinutes())+':'+f(this.getUTCSeconds())+'Z':null},Boolean.prototype.toJSON=this_value,Number.prototype.toJSON=this_value,String.prototype.toJSON=this_value),'function'!=typeof JSON.serialize&&(t={'\b':'\\b','\t':'\\t','\n':'\\n','\f':'\\f','\r':'\\r','""':'\\""','\\':'\\\\'},JSON.serialize=function(t){return str('',{'':t})})}();"

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

Public Property Get JScriptCode(ByVal JScriptName As String) As String
    Select Case JScriptName
        Case "JSON.serialize": JScriptCode = JSON_SERIALIZER
    End Select
End Property



' --- ByRef Variant Array ---

' WARN: Provided arr must be declared with ending parenthesis or access will instantly crash.
'   @EXAMPLE:
'
'       Dim t() as Variant, t2 as Variant
'       ReDim t(0 To 10)
'       t(0) = "foo"
'       t2 = CollectionsLib.GetArrayByRef(t)
'       t2(1) = "bar"
'       Debug.Print JSON.Stringify(t)
'
Public Function GetArrayByRef(ByRef arr As Variant) As Variant
    If IsArray(arr) Then
        GetArrayByRef = VarPtrArr(arr)
        Dim vt As VbVarType: vt = VarType(arr) Or VT_BYREF
        CopyMemory GetArrayByRef, vt, 2
    Else
        Err.Raise 5, "GetArrayByRef", "Array required"
    End If
End Function

#If Win64 Then
Public Function VarPtrArr(ByRef arr As Variant) As LongLong
#Else
Public Function VarPtrArr(ByRef arr As Variant) As Long
#End If
    Const vtArrByRef As Long = vbArray + VT_BYREF
    Dim vt As VbVarType
    CopyMemory vt, arr, 2
    If (vt And vtArrByRef) = vtArrByRef Then
        Const pArrayOffset As Long = 8
        CopyMemory VarPtrArr, ByVal VarPtr(arr) + pArrayOffset, PTR_SIZE
    Else
        Err.Raise 5, "VarPtrArr", "Array required"
    End If
End Function

' ---


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


' --- JAVASCRIPT FUNCTIONS ---

' TODO: Retrieve function name of non-anonymous functions.
Public Function TryParseAsJScriptFunction(ByRef CodeString As String, Optional ByVal FunctionName As String = "fn", Optional ByRef JScriptFunctionDefinition As Variant) As Boolean
    Dim Item As Variant, fnArgs As String, fnBody As String, isValid As Boolean, bodyLines() As String
    Dim regEx As Object: Set regEx = CreateObject("VBSCript.RegExp")
    On Error GoTo Finally
    
    regEx.IgnoreCase = True
    regEx.Global = True
    regEx.Pattern = "^\s*\(*?([^\(]*?)\)?\s?=>(.*?);*\s*$"

    For Each Item In regEx.Execute(CodeString)
        fnArgs = Item.SubMatches.Item(0)
        fnBody = Item.SubMatches.Item(1)
        bodyLines = VBA.Split(fnBody, ";")
        bodyLines(UBound(bodyLines)) = " return " & Trim(bodyLines(UBound(bodyLines))) & ";"
        fnBody = Trim(VBA.Join(bodyLines, ";"))
        isValid = True
        Exit For
    Next Item
    
    If Not isValid Then
        regEx.Pattern = "^\s*(?:function)?.*?\((.*?)\)\s?\{(.*?);*\s*}\s*$"
        For Each Item In regEx.Execute(CodeString)
            fnArgs = Item.SubMatches.Item(0)
            fnBody = Item.SubMatches.Item(1)
            bodyLines = VBA.Split(fnBody, ";")
            bodyLines(UBound(bodyLines)) = " return " & Trim(Replace(bodyLines(UBound(bodyLines)), "return", "")) & ";"
            fnBody = Trim(VBA.Join(bodyLines, ";"))
            isValid = True
            Exit For
        Next Item
    End If

Finally:
    If isValid Then
        TryParseAsJScriptFunction = True
        If Not IsMissing(JScriptFunctionDefinition) Then
            JScriptFunctionDefinition = Array("function " & FunctionName & "(" & fnArgs & ") { " & fnBody & " }", FunctionName, fnArgs)
        End If
    End If
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

Public Function StringLineSlices(ByRef SearchString As String) As ArrayListEx
    Dim aX As ArrayListEx, i As Long, iMax As Long, bQuotes As Boolean, r As Long
    Const cQuote As String = """"
    Set aX = ArrayListEx.Create()
    iMax = Len(SearchString)
    r = 1
    
    For i = 1 To iMax
        Select Case VBA.Mid$(SearchString, i, 1)
            Case cQuote: bQuotes = Not bQuotes
            Case vbLf
                If Not bQuotes Then
                    If VBA.Mid$(SearchString, i - 1, 1) = vbCr Then
                        aX.Add Array(r, i - (r + 1))
                    Else
                        aX.Add Array(r, i - r)
                    End If
                    r = i + 1
                End If
        End Select
    Next i
    
    If r < iMax Then
        aX.Add Array(r, iMax - r)
    End If
    
    Set StringLineSlices = aX
End Function


' --- MISC ---

' USAGE: ArrayRange(5, 4) => [5, 6, 7, 8]
Public Function ArrayRange(ByVal Index As Long, ByVal GetCount As Long) As Variant
    Dim i As Long, t() As Long
    ReDim t(0 To GetCount - 1)
    For i = 0 To GetCount - 1
        t(i) = Index + i
    Next i
    ArrayRange = GetArrayByRef(t)
End Function

Public Function ColumnLetter(ByVal ColumnIndex As Long) As String
    Const sChars As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    Dim i As Long

    Do
        ColumnLetter = VBA.Mid$(sChars, (0 ^ i) + ((Int(ColumnIndex / CLng(CLng(26) ^ i)) - (1 - (0 ^ i))) Mod CLng(26)) + (1 - (0 ^ i)), 1) & ColumnLetter
        i = i + 1
    Loop Until ((26 ^ i) + (26 * (1 - (0 ^ (i - 1)))) > ColumnIndex)
End Function

Public Function ArrayItem(ByRef ArrayLike As Variant, ByVal Index As Long, Optional ByRef ValueIfNull As Variant) As Variant
    If Not IsMissing(ValueIfNull) Then Assign ArrayItem, ValueIfNull
    On Error Resume Next
    AssignWhenValid ArrayItem, ArrayLike(Index)
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

