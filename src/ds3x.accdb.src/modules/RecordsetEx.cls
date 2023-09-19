VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RecordsetEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "RecordsetEx es una clase ultra-ligera, encadenable (chainable) y accesible globalmente (PredeclaredId = True) que encapsula un objeto ADODB.Recordset."
'@Folder "ds3x.Collections"
Option Compare Database
Option Explicit

Implements ICollectionEx

' RecordsetEx Class Module
'
'   RecordsetEx es una clase ultra-ligera, encadenable (chainable) y accesible globalmente (PredeclaredId = True)
'   que encapsula un objeto ADODB.Recordset.
'
'   Todos los miembros que devuelven un RecordsetEx son encadenables, es decir, puede ser utilizada de la siguiente forma:
'       RecordsetEx.Clone.MoveFirst.Filter("ID <= 15").PrintAll.Filter("ID > 15").PrintAll
'
'   Al ser accesible globalmente (PredeclaredId = True) siempre hay una instancia globalmente visible
'   llamada RecordsetEx para ser reutilizada sin necesidad de declarar nuevas instancias, aunque siempre
'   puedes declararlas si lo deseas. Para reutilizar un RecordsetEx, simplemente vinculalo a otro
'   ADODB.Recordset mediante el método .Bind(TargetRecordset as ADODB.Recordset).
'
'   Cada instancia de RecordsetEx en memória, solo ocupa unos 13 o 17 bytes, según si se ejecuta
'   en una plataforma de 32-bits o de 64-bits, que corresponden a:
'       1. Una referéncia a un objeto ADODB.Recordset (4 bytes o 8 bytes según plataforma).
'       2. Un valor tipo Long (8 bytes).
'       3. Un valor tipo Boolean (1 byte).
'
'
' --- PUBLIC MEMBERS ---
' .Instance As ADODB.Recordset                  - La referencia al ADODB.Recordset vinculado.
'
' --- PUBLIC PROPERTY GETTERS ---
' .Count() As Long                              - Equivalente a Recordset.RecordCount
' .Columns() As Long                            - Equivalente a Recordset.Fields.Count
' .Item(ItemPosition As Long) As RecordsetEx    - Devuelve un RecordsetEx (Shallow Copy) apuntando a la fila especificada en ItemPosition.
' .Fields() As ADODB.Fields                     - Equivalente a Recordset.Fields
' .Field(FieldNameOrIndex As Variant) As ADODB.Field    - Equivalente a Recordset.Fields(FieldNameOrIndex)
' .ColumnNames() As Variant()                   - Devuelve un array de strings con los nombres de las columnas del Recordset.
' .Items() As Variant()                         - Devuelve un Array iterable que permite recorrer las filas del Recordset mediante For Each.
'
' --- PUBLIC FUNCTIONS ---
' .Create(Optional TargetRecordset As ADODB.Recordset = Nothing) As RecordsetEx         - Devuelve una nueva instancia de RecordsetEx vinculada al Recordset proporcionado.
' .Bind(TargetRecordset As ADODB.Recordset) As RecordsetEx   - Vincula el objeto Recordset proporcionado con este RecordsetEx.
' .Unbind() As RecordsetEx                                   - Desvincula el objeto Recordset.
' .Filter(Optional QueryFilter As String) As RecordsetEx     - Equivalente a la propiedad .Filter de Recordset pero como método encadenable.
' .MoveFirst() As RecordsetEx                                - Alternativa segura y encadenable al método .MoveFirst de Recordset.
' .MoveLast() As RecordsetEx                                 - Alternativa segura y encadenable al método .MoveLast de Recordset.
' .MoveNext() As RecordsetEx                                 - Alternativa segura y encadenable al método .MoveNext de Recordset.
' .MovePrevious() As RecordsetEx                             - Alternativa segura y encadenable al método .MovePrevious de Recordset.
' .Move(NumRecords As Long) As RecordsetEx                   - Alternativa segura y encadenable al método .Move de Recordset.
' .AsIterableDictionary(ParamArray QueryColumns()) As Variant()                         - Devuelve un Array() de Scripting.Dictionary (@SEE: RecordsetEx.AsIterableString).
' .AsIterable(ParamArray QueryColumns()) As Variant()        - Devuelve un Array() de Array()s (@SEE: RecordsetEx.AsIterableString).
' .AsIterableString(StringDelimiter As String, ParamArray QueryColumns()) As Variant()  - Devuelve un Array() de Strings, compuestas por la concatenación de los campos mediante el separador StringDelimiter.
' .Requery() As RecordsetEx                                  - Actualiza los datos del objeto Recordset volviendo a ejecutar la query que lo generó.
' .IndexOf(SearchCriteria As String) As Long                 - Devuelve la posición del primer registro en un Recordset que cumpla con SearchCritera.
' .LastIndexOf(SearchCriteria As String) As Long             - Devuelve la posición del último registro en un Recordset que cumpla con SearchCritera (@SEE: RecordsetEx.IndexOf).
' .Clone() As RecordsetEx                                    - Devuelve una nueva instancia de RecordsetEx vinculada a un Recordset.Clone() de esta Instance.
' .Duplicate() As RecordsetEx                                - Realiza y devuelve una copia entera del conjunto de datos, desvinculandolos del Recordset original.
' .Search(SearchCriteria As Variant, Optional ContinueBookmark As Variant, Optional SearchDirection As SearchDirectionEnum) As Variant   - Devuelve el Recordset.Bookmark del primer/siguiente registro que coincida con SearchCriteria.
' .ToString(Optional ColumnDelimiter As String, Optional LimitRows As Long) As String   - Similar al método Recordset.GetString().
' .PrintAll() As RecordsetEx                                 - Imprime el recordset en formato Markdown Table a la ventana Inmediato.
'
'
' @EXAMPLE: Ejemplo real para ejecutar diréctamente en consola (ventana Inmediato) sin necesidad de abrir ningún formulário.
'
'   RecordsetEx.Bind dbQuery.Create("SELECT ID, REPORT_TYPE, REPORT_CODE, REPORT_LEVEL, CONTENT_TYPE, REPORT_DATE, REPORT_STATUS FROM [dbo].[T_ATM_ISCS_APP_REPORTS]")
'
'   RecordsetEx.Filter("REPORT_DATE < '20/04/2023' AND REPORT_STATUS = 'RESUELTA'").PrintAll
'   ID  | REPORT_TYPE      | REPORT_CODE          | REPORT_LEVEL     | CONTENT_TYPE     | REPORT_DATE          | REPORT_STATUS
'   ------------------------------------------------------------------------------------------------------------------------------
'   1   | DATA_ERROR       | GMC:CODE_041         | CRITICAL         | JSON             | 11/04/2023 6:44:24   | RESUELTA
'   2   | DATA_ERROR       | DBF:CODE_043         | WARNING          | JSON             | 12/04/2023 23:30:30  | RESUELTA
'   3   | IMPORT_ERROR     | SQL_VEF:CODE_047     | WARNING          | JSON             | 14/04/2023 22:51:45  | RESUELTA
'   4   | DATA_ERROR       | GMC:CODE_041         | CRITICAL         | JSON             | 19/04/2023 0:17:37   | RESUELTA
'   5   | DATA_ERROR       | GMC:CODE_041         | CRITICAL         | JSON             | 19/04/2023 0:19:58   | RESUELTA
'
'
' @EXAMPLE: Recorrer todas las columnas de todas las filas de un recordset e imprimir sus valores.
'
'     Dim Row As Variant, i As Long
'
'     For Each Row In RecordsetEx.Items
'         For i = 0 To Row.Fields.Count - 1
'             Debug.Print Row.Field(i).Name & ": " & Nz(Row.Field(i))
'         Next i
'     Next Row
'
'
' @EXAMPLE: Acceder directamente al valor del campo "ID" de la fila 5 e imprimirlo en la ventana Inmediato.
'
'     Debug.Print RecordsetEx(5).Field("ID")
'

Public Instance As ADODB.Recordset
Private pShallowRecord As Long
Private pIsShallowCopy As Boolean


' --- PROPERTIES ---

' Equivalente a Recordset.RecordCount
Public Property Get Count() As Long
Attribute Count.VB_Description = "Equivalente a Recordset.RecordCount"
    On Error Resume Next: Count = Instance.RecordCount
End Property

' DEPRECATED
Public Property Get Columns() As Long
Attribute Columns.VB_Description = "Equivalente a Recordset.Fields.Count"
    On Error Resume Next: Columns = Instance.Fields.Count
End Property

Public Property Get ColumnCount() As Long
    On Error Resume Next: ColumnCount = Instance.Fields.Count
End Property

' Devuelve un RecordsetEx (Shallow Copy) que se desplazará automáticamente a la fila (record)
' indicada, pero solo en el momento en que se accede a sus respectivos campos o valores, ya que
' al ser una "Shallow Copy" comparte la misma instancia del Recordset con el RecordsetEx original
' y otras posibles copias.
'
' Como alternativa, puedes pasarle una String con el nombre de un campo o su posición (base 0)
' para obtener dicho campo (ADODB.Field) o diréctamente su valor.
'
' Incluso se puede combinar ambas posibilidades para acceder directamente a un valor concreto.
'
' Es decir, dado el siguiente código: Set rX = RecordsetEx.Create(...)
'   rX(4)       -> Devuelve un RecordsetEx (ShallowCopy), al acceder a sus valores, se moverá a la fila 4.
'   rX("ID")    -> Devuelve un ADODB.Field del valor ID del registro al que esté apuntando rX.Instance en ese instante.
'   rX(4)("ID") -> Se mueve a la fila 4 y devuelve el campo ID.
'   rX(1)("0")    -> Se mueve a la primera fila y devuelve el campo en primera posición.
'
'@DefaultMember
Public Property Get Item(ByVal NumRowOrFieldName As Variant) As Variant
Attribute Item.VB_UserMemId = 0
    If VarType(NumRowOrFieldName) = vbString Then
        Set Item = Field(NumRowOrFieldName)
    Else
        Set Item = CreateShallowCopy(NumRowOrFieldName)
    End If
End Property

Public Property Get Row(ByVal Index As Long) As Variant
    Dim aX As ArrayListEx, t As ADODB.Fields, i As Long
    
    Set aX = ArrayListEx.Create()
    Set t = CreateShallowCopy(Index + 1).Fields
    
    For i = 0 To ColumnCount - 1
        aX.Add t.Item(i).Value
    Next i
    
    Row = aX.ToArray()
End Property

' Equivalente a Recordset.Fields
Public Property Get Fields() As ADODB.Fields
Attribute Fields.VB_Description = "Equivalente a Recordset.Fields"
    MoveToShallowRecord
    Set Fields = Instance.Fields
End Property

' Equivalente a Recordset.Fields(FieldNameOrIndex)
Public Property Get Field(ByRef FieldNameOrIndex As Variant) As ADODB.Field
Attribute Field.VB_Description = "Equivalente a Recordset.Fields(FieldNameOrIndex)"
    MoveToShallowRecord
    On Error GoTo Fallback
    Set Field = Instance.Fields(FieldNameOrIndex)
    Exit Property
Fallback:
    If VarType(FieldNameOrIndex) = vbString And CStr(Val(FieldNameOrIndex)) = CStr(FieldNameOrIndex) Then
        Set Field = Instance.Fields(CLng(Val(FieldNameOrIndex)))
    End If
End Property

' Devuelve un array de strings con los nombres de las columnas del Recordset.
Public Property Get ColumnNames() As Variant()
Attribute ColumnNames.VB_Description = "Devuelve un array de strings con los nombres de las columnas del Recordset."
    Dim i As Long, sNames() As Variant
    ReDim sNames(Instance.Fields.Count - 1)
    
    For i = 0 To Instance.Fields.Count - 1
        sNames(i) = Instance.Fields(i).Name
    Next i

    ColumnNames = sNames
End Property

' Devuelve un Array iterable que permite recorrer las filas del Recordset mediante For Each.
Public Property Get Items() As Variant()
Attribute Items.VB_Description = "Devuelve un Array iterable que permite recorrer las filas del Recordset mediante For Each."
    Dim i As Long, sItems() As Variant
    ReDim sItems(Count - 1)

    For i = 0 To Count - 1
        Set sItems(i) = CreateShallowCopy(i + 1)
    Next i

    Items = sItems
End Property

Public Property Let ShallowRecord(ByVal Value As Long)
    pShallowRecord = Value
    pIsShallowCopy = True
End Property


' --- PUBLIC METHODS ---

Public Function CreateBlank(ByVal RowsCount As Long, ByVal ColumnsCount As Long) As RecordsetEx: Set CreateBlank = RecordsetEx.Create(CreateRecordset(RowsCount, 0, ColumnsCount)): End Function

' Devuelve una nueva instancia de RecordsetEx vinculada al Recordset proporcionado.
Public Function Create(Optional ByVal RecordsetLike As Variant) As RecordsetEx
    Set Create = New RecordsetEx
    Set Create.Instance = Nothing
    
    If Not IsMissing(RecordsetLike) Then
        If TypeOf RecordsetLike Is ADODB.Recordset Then
            Set Create.Instance = RecordsetLike
        ElseIf TypeOf RecordsetLike Is ArrayListEx Then
            Set Create.Instance = CreateFromArrayList(RecordsetLike.Instance)
        ElseIf TypeOf RecordsetLike Is ArrayList Then
            Set Create.Instance = CreateFromArrayList(C2ArrayList(RecordsetLike))
        ' TODO: Create from: Array2dEx, dsTable, JSON
        Else
            ' Dirty patch
            If TypeName(RecordsetLike) = "ArrayListEx" Then
                Set Create.Instance = CreateFromArrayList(RecordsetLike.Instance)
            End If
        End If
    End If
End Function

' Vincula el objeto Recordset proporcionado con este RecordsetEx
Public Function Bind(ByRef TargetRecordset As ADODB.Recordset) As RecordsetEx
Attribute Bind.VB_Description = "Vincula el objeto Recordset proporcionado con este RecordsetEx."
    Set Instance = Nothing: Set Instance = TargetRecordset: Set Bind = Me
End Function

' Desvincula el objeto Recordset
Public Function Unbind() As RecordsetEx
Attribute Unbind.VB_Description = "Desvincula el objeto Recordset."
    Set Instance = Nothing: Set Unbind = Me
End Function

' Equivalente a la propiedad .Filter de Recordset pero como método encadenable.
Public Function Filter(Optional ByVal QueryFilter As String = "") As RecordsetEx
Attribute Filter.VB_Description = "Equivalente a la propiedad .Filter de Recordset pero como método encadenable."
    Instance.Filter = QueryFilter: Set Filter = Me
End Function

' Alternativa segura y encadenable al método .MoveFirst de Recordset. (Una mejor definición sería: "Chainable safe wrapper around Recordset.MoveFirst")
Public Function MoveFirst() As RecordsetEx
Attribute MoveFirst.VB_Description = "Alternativa segura y encadenable al método .MoveFirst de Recordset."
    On Error Resume Next: Instance.MoveFirst: Set MoveFirst = Me
End Function

' Alternativa segura y encadenable al método .MoveLast de Recordset.
Public Function MoveLast() As RecordsetEx
Attribute MoveLast.VB_Description = "Alternativa segura y encadenable al método .MoveLast de Recordset."
    On Error Resume Next: Instance.MoveLast: Set MoveLast = Me
End Function

' Alternativa segura y encadenable al método .MoveNext de Recordset.
Public Function MoveNext() As RecordsetEx
Attribute MoveNext.VB_Description = "Alternativa segura y encadenable al método .MoveNext de Recordset."
    On Error Resume Next: Instance.MoveNext: Set MoveNext = Me
End Function

' Alternativa segura y encadenable al método .MovePrevious de Recordset.
Public Function MovePrevious() As RecordsetEx
Attribute MovePrevious.VB_Description = "Alternativa segura y encadenable al método .MovePrevious de Recordset."
    On Error Resume Next: Instance.MovePrevious: Set MovePrevious = Me
End Function

' Alternativa segura y encadenable al método .Move de Recordset.
Public Function Move(ByVal NumRecords As Long) As RecordsetEx
Attribute Move.VB_Description = "Alternativa segura y encadenable al método .Move de Recordset."
    On Error Resume Next: Instance.Move NumRecords: Set Move = Me
End Function

Public Function GetRange(Optional ByVal Index As Variant, Optional ByVal GetCount As Variant, Optional ByVal ColumnIndexes As Variant) As Array2dEx
    Dim i As Long, c As Long, t() As Variant, r() As Variant
    
    If IsMissing(Index) Or IsEmpty(Index) Then Index = 0
    If Index >= Me.Count Then
        Index = IIf(Me.Count - 1 < 0, 0, Me.Count - 1)
        GetCount = 0
    Else
        If IsMissing(GetCount) Or IsEmpty(GetCount) Then GetCount = Me.Count - Index
        If GetCount + Index > Me.Count Then GetCount = Me.Count - Index
    End If
    If IsMissing(ColumnIndexes) Or IsEmpty(ColumnIndexes) Then ColumnIndexes = CreateNumericIndexHeaders
    If GetCount <= 0 Then GoTo Finally
    ReDim t(0 To GetCount - 1, 0 To UBound(ColumnIndexes))
    
    For c = 0 To UBound(ColumnIndexes)
        If Index > 0 Then
            Instance.AbsolutePosition = Index + 1
        Else
            Me.MoveFirst
        End If
        r = Instance.GetRows(GetCount, adBookmarkCurrent, Array(ColumnIndexes(c)))
        If Instance.Fields(ColumnIndexes(c)).Type = adVarWChar Then
            For i = 0 To GetCount - 1
                t(i, c) = C2Date(r(0, i))
            Next i
        Else
            For i = 0 To GetCount - 1
                t(i, c) = r(0, i)
            Next i
        End If
    Next c
    
Finally:
    Set GetRange = New Array2dEx
    GetRange.Instance = CollectionsLib.GetArrayByRef(t)
End Function

' NOTICE: Returns a new RecordsetEx.
Public Function AddRange(ByVal Target As Variant) As RecordsetEx
    Dim dsT As dsTable, dsT2 As dsTable
    
    Set dsT = dsTable.Create(Me, True)
    Set dsT2 = dsTable.Create(Target, False).SetHeaders(Me)
    Set AddRange = RecordsetEx.Create(dsT.AddRange(dsT2).NamedRecordset)
End Function

Public Function CopyToRange(ByRef RangeObject As Excel.Range, Optional ByVal ApplyUserLocale As Boolean = True, Optional ByVal WriteHeaders As Boolean = True) As RecordsetEx
    Set CopyToRange = Me
    dsTable.Create(Me, True).CopyToRange RangeObject, ApplyUserLocale, WriteHeaders
End Function

' Devuelve un Array() de Scripting.Dictionary (@SEE: RecordsetEx.AsIterableString).
Public Function AsIterableDictionary(ParamArray QueryColumns() As Variant) As Variant()
Attribute AsIterableDictionary.VB_Description = "Devuelve un Array() de Scripting.Dictionary (@SEE: RecordsetEx.AsIterableString)."
    Dim sValues As Variant, sRows() As Variant, Entry As Scripting.Dictionary, sColumns() As Variant, i As Long, c As Long
    
    If Instance.RecordCount > 0 Then
        If UBound(QueryColumns) >= 0 Then
            sValues = Instance.GetRows(adGetRowsRest, adBookmarkFirst, QueryColumns)
            sColumns = QueryColumns
        Else
            sValues = Instance.GetRows(adGetRowsRest, adBookmarkFirst)
            sColumns = ColumnNames
        End If
        
        ReDim sRows(UBound(sValues, 2) - LBound(sValues, 2))
    
        For i = LBound(sValues, 2) To UBound(sValues, 2)
            Set Entry = New Scripting.Dictionary
            
            For c = 0 To UBound(sColumns)
                Entry(CStr(sColumns(c))) = sValues(c, i)
            Next c
    
            Set sRows(i) = Entry
        Next i
    Else
        sRows = Array()
    End If
    
    MoveFirst
    AsIterableDictionary = sRows
End Function

' Devuelve un Array() de Array()s (@SEE: RecordsetEx.AsIterableString).
Public Function AsIterable(ParamArray QueryColumns() As Variant) As Variant()
Attribute AsIterable.VB_Description = "Devuelve un Array() de Array()s (@SEE: RecordsetEx.AsIterableString)."
    Dim sValues As Variant, sRows() As Variant, Entry() As Variant, sColumns() As Variant, i As Long, c As Long
    
    If Instance.RecordCount > 0 Then
        If UBound(QueryColumns) >= 0 Then
            sValues = Instance.GetRows(adGetRowsRest, adBookmarkFirst, QueryColumns)
            sColumns = QueryColumns
        Else
            sValues = Instance.GetRows(adGetRowsRest, adBookmarkFirst)
            sColumns = ColumnNames
        End If
        
        ReDim sRows(UBound(sValues, 2) - LBound(sValues, 2))
        ReDim Entry(UBound(sColumns))
    
        For i = LBound(sValues, 2) To UBound(sValues, 2)
            For c = 0 To UBound(sColumns)
                Entry(c) = sValues(c, i)
            Next c
    
            sRows(i) = Entry
        Next i
    Else
        sRows = Array()
    End If
    
    MoveFirst
    AsIterable = sRows
End Function

' Devuelve un Array() de Strings, compuestas por la concatenación de los campos mediante el separador StringDelimiter.
'
' El parámetro QueryColumns() es opcional, por lo que si no se especifica ningún campo, se devuelven todos los campos del recordset.
'
' @EXAMPLE: Reconstruir un ComboBox con los usuarios cuyo perfil sea administrador.
'
'   Dim Item As Variant
'
'   ' Limpiar elementos existentes en ComboBox
'   Me.COMBO_ADMINS.RowSourceType = "Value List"
'   Me.COMBO_ADMINS.RowSource = vbNullString
'
'   ' Añadir elementos en ComboBox
'   For Each Item In RecordsetEx.Bind(pRecordsetUsuarios).Filter("IS_ADMIN = 1").AsIterableString(";", "USUARIO", "NOMBRE")
'       Me.COMBO_ADMINS.AddItem Item
'   Next Item
'
Public Function AsIterableString(ByVal StringDelimiter As String, ParamArray QueryColumns() As Variant) As Variant()
Attribute AsIterableString.VB_Description = "Devuelve un Array() de Strings, compuestas por la concatenación de los campos mediante el separador StringDelimiter."
    Dim sValues As Variant, sRows() As Variant, Entry As String, sColumns() As Variant, i As Long, c As Long
    
    If Instance.RecordCount > 0 Then
        If UBound(QueryColumns) >= 0 Then
            sValues = Instance.GetRows(adGetRowsRest, adBookmarkFirst, QueryColumns)
            sColumns = QueryColumns
        Else
            sValues = Instance.GetRows(adGetRowsRest, adBookmarkFirst)
            sColumns = ColumnNames
        End If
        
        ReDim sRows(UBound(sValues, 2) - LBound(sValues, 2))
    
        For i = LBound(sValues, 2) To UBound(sValues, 2)
            Entry = ""
            For c = 0 To UBound(sColumns)
                Entry = Entry & StringDelimiter & sValues(c, i)
            Next c
    
            sRows(i) = VBA.Mid(Entry, Len(StringDelimiter) + 1)
        Next i
    Else
        sRows = Array()
    End If
    
    MoveFirst
    AsIterableString = sRows
End Function

' Actualiza los datos del objeto Recordset volviendo a ejecutar la query que lo generó.
'
' NOTA: El método .Requery() de un objeto Recordset provoca error si el recordset se ha editado, solo
'       actualiza el cursor actual del recordset y además provoca que várias propiedades pasen a ser
'       read-only. Este método soluciona estos problemas.
Public Function Requery() As RecordsetEx
Attribute Requery.VB_Description = "Actualiza los datos del objeto Recordset volviendo a ejecutar la query que lo generó."
    On Error Resume Next
    Instance.Close
    Instance.Open
    Set Requery = Me
End Function

' Devuelve la posición del primer registro en un Recordset que cumpla con SearchCritera.
'
' SearchCritera acepta un formato similar al de un WHERE en SQL Server pero para una única condición,
' es decir, no permite el uso de funciones, AND(s), OR(s) o IN(s) pero se puede buscar por cualquier campo.
'
' @EXAMPLE: En un formulario continuo, selecciona el registro con ID = 42 si existe.
'
'   ' NOTA: Asignar un valor a .SelTop cambia el valor de .CurrentRecord, hace scroll del formulario para
'   '       mostrar el registro seleccionado y llama al evento OnCurrent() del formulario.
'
'   Me.SelTop = RecordsetEx.Bind(Me.Recordset).IndexOf("ID = 42", Me.SelTop)
'
Public Function IndexOf(ByVal SearchCriteria As String, Optional ByVal DefaultValue As Long = -1) As Long
Attribute IndexOf.VB_Description = "Devuelve la posición del primer registro en un Recordset que cumpla con SearchCritera."
    On Error GoTo Finally
    IndexOf = DefaultValue

    If Instance.RecordCount > 0 Then
        Instance.Find SearchCriteria, 0, adSearchForward, adBookmarkFirst
        If Instance.AbsolutePosition > 0 Then IndexOf = Instance.AbsolutePosition
    End If
    
Finally:
    MoveFirst
End Function

' Devuelve la posición del último registro en un Recordset que cumpla con SearchCritera (@SEE: RecordsetEx.IndexOf).
Public Function LastIndexOf(ByVal SearchCriteria As String, Optional ByVal DefaultValue As Long = -1) As Long
Attribute LastIndexOf.VB_Description = "Devuelve la posición del último registro en un Recordset que cumpla con SearchCritera (@SEE: RecordsetEx.IndexOf)."
    On Error GoTo Finally
    LastIndexOf = DefaultValue

    If Instance.RecordCount > 0 Then
        Instance.Find SearchCriteria, 0, adSearchBackward, adBookmarkLast
        If Instance.AbsolutePosition > 0 Then LastIndexOf = Instance.AbsolutePosition
    End If
    
Finally:
    MoveFirst
End Function

Public Function Join(ByVal Target As RecordsetEx) As RecordsetEx
    Set Join = RecordsetEx.Create()
    ' TODO: Implement
    Err.Raise 17, , "Not implemented."
End Function

' Devuelve una nueva instancia de RecordsetEx vinculada a un Recordset.Clone() de esta Instance.
Public Function Clone() As RecordsetEx
Attribute Clone.VB_Description = "Devuelve una nueva instancia de RecordsetEx vinculada a un Recordset.Clone() de esta Instance."
    Dim xClone As New RecordsetEx
    Set Clone = xClone.Bind(Instance.Clone)
    If VarType(Instance.Filter) = vbString Then Clone.Filter Instance.Filter
End Function

' Un Recordset.Clone() solo devuelve una nueva instancia de un objeto Recordset pero apuntando al mismo
' conjunto de datos que el Recordset original. Si el original se modifica, su clon también. Pero mediante
' RecordsetEx.Duplicate() se realiza una copia entera del conjunto de datos, desvinculandolos del Recordset
' original.
'
' NOTA: Si el Recordset no permite ediciones, como los devueltos por el conector a la BBDD, no hay razón para
'       hacer un duplicado entero de su conjunto de datos ya que estos nunca van a cambiar.
'
Public Function Duplicate() As RecordsetEx
Attribute Duplicate.VB_Description = "Realiza y devuelve una copia entera del conjunto de datos, desvinculandolos del Recordset original."
    Dim xClone As New RecordsetEx, rsClone As New ADODB.Recordset, rStream As New ADODB.Stream
    
    Instance.Save rStream
    rsClone.Open rStream
    
    rStream.Close
    Set rStream = Nothing
    
    Set Duplicate = xClone.Bind(rsClone)
    ' @SEE: You can populate the Fields collection before opening the Recordset by calling the Refresh method on the collection.
End Function

' Devuelve el Recordset.Bookmark de la primera coincidencia con SearchCriteria. Si se especifica un Bookmark en ContinueBookmark,
' devolverá el Recordset.Bookmark de la siguiente coincidencia con SearchCriteria partiendo del Bookmark especificado y según la
' dirección establecida en SearchDirection.
'
' NOTA: SearchCriteria puede ser tanto una string como un array en formato SearchDefinition (@SEE: GenerateSearchDefinition).
'
Public Function Search(ByRef SearchCriteria As Variant, Optional ByVal ContinueBookmark As Variant = -1, Optional ByVal SearchDirection As SearchDirectionEnum = adSearchForward) As Variant
Attribute Search.VB_Description = "Devuelve el Recordset.Bookmark del primer/siguiente registro que coincida con SearchCriteria."
    If Instance.RecordCount = 0 Then GoTo SearchFailedToMatch

    If ContinueBookmark <> -1 Then
        If Instance.Bookmark <> ContinueBookmark Then Instance.Bookmark = ContinueBookmark
        If Not TryMoveIn(SearchDirection) Then GoTo SearchFailedToMatch
    End If
    
    If ContinueBookmark = -1 Then
        MoveFirst
        If SearchDirection = adSearchBackward Then MoveLast
    End If
    
    Search = -1
    If VarType(SearchCriteria) = vbString Then
        Do
            If CollectionsLib.TryMatchSearchCriteriaIn(SearchCriteria, Instance.Fields) Then
                Search = Instance.Bookmark
                Exit Do
            End If
        Loop Until (Not TryMoveIn(SearchDirection))
    Else
        Do
            If CollectionsLib.TryMatchSearchDefinitionIn(SearchCriteria, Instance.Fields) Then
                Search = Instance.Bookmark
                Exit Do
            End If
        Loop Until (Not TryMoveIn(SearchDirection))
    End If
    
    Exit Function
SearchFailedToMatch:
    Search = -1
End Function


' --- .ToCSV(), .ToJSON(), .ToString() & .PrintAll() ---

Public Function ToCSV(Optional ByVal Delimiter As String = ",", Optional ByVal InLocalFormat As Boolean = False) As String
    ToCSV = dsTable.Create(Me, True).ToCSV(Delimiter, InLocalFormat)
End Function

Public Function ToJSON() As String: ToJSON = JSON.Stringify(Me, 2): End Function

Public Function ToExcel() As String: ToExcel = ToCSV(vbTab, True): End Function

Public Function ToString(Optional ByVal ColumnDelimiter As String = vbTab, Optional ByVal LimitRows As Long = -1) As String
Attribute ToString.VB_Description = "Similar al método Recordset.GetString()."
    On Error Resume Next
    ToString = Instance.GetString(adClipString, LimitRows, ColumnDelimiter)
End Function

' Imprime el recordset en formato Markdown Table a la ventana Inmediato.
'
' Al respetar las especificaciones del formato Markdown Table, permite mostrarlo en formato de tabla en un documento Markdown
' o incluso importar los datos como una tabla y poder trabajar con ellos en aplicaciones que soporten este formato, una de ellas
' es tableconvert.com, una herramienta online que te permite editarlas y descargarlas/convertirlas a otros formatos, como Excel,
' CSV, HTML o incluso en SQL, con sus scripts de CREATE TABLE e INSERT INTO, entre otros formatos.
Public Function PrintAll() As RecordsetEx
Attribute PrintAll.VB_Description = "Imprime el recordset en formato Markdown Table a la ventana Inmediato."
    On Error GoTo Finally
    Dim ColumnCount As Long, i As Long, k As Long, ColumnName As String, itemStr As String, itemLen As Long, accPosX As Long
    Dim Entry As Collection, Entries As New Collection, columnNamesCollection As New Collection, columnSizesDictionary As New Scripting.Dictionary

    MoveFirst
    With Instance
        ColumnCount = .Fields.Count

        For i = 0 To ColumnCount - 1
            ColumnName = CStr(.Fields(i).Name)
            columnNamesCollection.Add ColumnName
            columnSizesDictionary(i + 1) = (Len(ColumnName) + (4 - (Len(ColumnName) Mod 4)))
        Next i
        
        If Not .BOF And Not .EOF Then
            While (Not .EOF)
                Set Entry = New Collection
                
                For i = 0 To ColumnCount - 1
                    itemStr = CStr(Nz(.Fields(i), ""))
                    itemLen = (Len(itemStr) + (4 - (Len(itemStr) Mod 4)))
                    If CLng(columnSizesDictionary(i + 1)) < itemLen Then
                        columnSizesDictionary(i + 1) = itemLen
                    End If
                    Entry.Add itemStr
                Next i
                
                Entries.Add Entry
                .MoveNext
            Wend
            
            MoveFirst
        End If
        
        ' PRINTING HEADERS
        accPosX = 0
        For k = 1 To columnNamesCollection.Count
            If k <> columnNamesCollection.Count Then
                Debug.Print columnNamesCollection(k); Tab(accPosX + CLng(columnSizesDictionary(k))); " | ";
                accPosX = accPosX + CLng(columnSizesDictionary(k)) + 3
            Else
                Debug.Print columnNamesCollection(k); Tab(accPosX + CLng(columnSizesDictionary(k)));
                accPosX = accPosX + CLng(columnSizesDictionary(k))
            End If
        Next k
        Debug.Print ""; Tab(0);
        Debug.Print String(accPosX, "-")
        
        ' PRINTING TABLE ROWS
        For i = 1 To Entries.Count
            accPosX = 0
            Set Entry = Entries(i)
            For k = 1 To columnNamesCollection.Count
                If k <> columnNamesCollection.Count Then
                    Debug.Print Entry(k); Tab(accPosX + CLng(columnSizesDictionary(k))); " | ";
                    accPosX = accPosX + CLng(columnSizesDictionary(k)) + 3
                Else
                    Debug.Print Entry(k); Tab(accPosX + CLng(columnSizesDictionary(k)));
                    accPosX = accPosX + CLng(columnSizesDictionary(k))
                End If
            Next k
            Debug.Print ""; Tab(0);
        Next i

    End With
    Set PrintAll = Me
    Exit Function

Finally:
    Set PrintAll = Me
End Function


' --- PRIVATE METHODS ---

Private Sub Class_Initialize()
    pShallowRecord = 0
End Sub

Private Sub Class_Terminate()
    Dispose
End Sub

Public Sub Dispose(): Set Instance = Nothing: End Sub

Private Sub MoveToShallowRecord()
    If pIsShallowCopy Then
        If Instance.AbsolutePosition <= 0 Then MoveFirst
        Move pShallowRecord - Instance.AbsolutePosition
    End If
End Sub

Private Function CreateShallowCopy(ByVal AtRecord As Long) As RecordsetEx
    Set CreateShallowCopy = New RecordsetEx
    Set CreateShallowCopy.Instance = Instance
    CreateShallowCopy.ShallowRecord = AtRecord
End Function

Private Function TryMoveIn(MoveDirection As SearchDirectionEnum) As Boolean
    On Error GoTo MoveFail
    
    If MoveDirection = adSearchForward Then
        Instance.MoveNext
    Else
        Instance.MovePrevious
    End If
    
    TryMoveIn = (Instance.AbsolutePosition >= 0)
MoveFail:
End Function

Private Function CreateRecordset(ByVal RowsCount As Long, ByVal ColumnStartIndex As Long, ByVal ColumnsCount As Long) As ADODB.Recordset
    Dim rs As New ADODB.Recordset, iRow() As Variant, rValues() As Variant, i As Long
    
    ReDim iRow(0 To ColumnsCount - 1)
    ReDim rValues(0 To ColumnsCount - 1)
    
    With rs
        For i = LBound(iRow) To UBound(iRow)
            iRow(i) = CStr(ColumnStartIndex + i)
            .Fields.Append CStr(iRow(i)), adLongVarWChar, -1, adFldIsNullable
            rValues(i) = ""
        Next i
        .Open
        For i = 0 To RowsCount - 1
            .AddNew FieldList:=iRow, Values:=rValues
        Next i
        .MoveFirst
    End With
    
    Set CreateRecordset = rs
End Function

Private Function CreateFromArrayList(ByRef ArrayLike As ArrayList) As ADODB.Recordset
    Dim rs As New ADODB.Recordset, i As Long, c As Long, cMax As Long, cMin As Long, iRow As Variant
    Set CreateFromArrayList = rs
    If ArrayLike.Count = 0 Then Exit Function
    
    iRow = ArrayLike(0)
    cMax = UBound(iRow)
    cMin = LBound(iRow)
    
    With rs
        For c = cMin To cMax
            .Fields.Append CStr(iRow(c)), adLongVarWChar, -1, adFldIsNullable
        Next c
        
        .Open
        
        For i = 1 To ArrayLike.Count - 1
            .AddNew FieldList:=iRow, Values:=ArrayLike(i)
        Next i
    End With
    Set CreateFromArrayList = rs
End Function

Private Function CreateNumericIndexHeaders() As Variant
    Dim t As Variant, c As Long, cMax As Long
    cMax = Me.ColumnCount - 1
    ReDim t(0 To cMax)
    
    For c = 0 To cMax
        t(c) = c
    Next c
    
    CreateNumericIndexHeaders = t
End Function


' --- TYPE CASTING ---

Private Function C2Date(Optional ByVal Value As Variant) As Variant
    If IsMissing(Value) Then Exit Function
    On Error Resume Next
    C2Date = Value
    C2Date = DateValue(Value)
End Function

Private Function C2ArrayList(ByRef Value As Variant) As ArrayList: Set C2ArrayList = Value: End Function
Private Function C2Recordset(ByRef Value As Variant) As ADODB.Recordset: Set C2Recordset = Value: End Function


' --- ICollectionEx ---

Private Property Get ICollectionEx_Count() As Long: ICollectionEx_Count = Count: End Property
Private Property Get ICollectionEx_ColumnCount() As Long: ICollectionEx_ColumnCount = ColumnCount: End Property
Private Property Get ICollectionEx_Item(ByVal Index As Long, Optional ByVal ColumnIndex As Long = 0&) As Variant: ICollectionEx_Item = Item(Index)(CStr(ColumnIndex)): End Property
Private Property Get ICollectionEx_Row(ByVal Index As Long) As Variant: ICollectionEx_Row = Row(Index): End Property
Private Function ICollectionEx_CreateBlank(ByVal RowsCount As Long, ByVal ColumnsCount As Long) As ICollectionEx: Set ICollectionEx_CreateBlank = CreateBlank(RowsCount, ColumnsCount): End Function
Private Function ICollectionEx_Create(Optional ByRef FromTarget As Variant) As ICollectionEx: Set ICollectionEx_Create = Create(FromTarget): End Function
Private Function ICollectionEx_Bind(Optional ByRef Target As Variant) As ICollectionEx: Set ICollectionEx_Bind = Bind(C2Recordset(Target)): End Function
Private Function ICollectionEx_Unbind() As ICollectionEx: Set ICollectionEx_Unbind = Unbind(): End Function
Private Sub ICollectionEx_Dispose(): Dispose: End Sub
Private Function ICollectionEx_Join(ByVal Target As ICollectionEx) As ICollectionEx: Set ICollectionEx_Join = Join(Target): End Function
Private Function ICollectionEx_AddRange(ByVal Target As ICollectionEx) As ICollectionEx: Set ICollectionEx_AddRange = AddRange(Target): End Function
Private Function ICollectionEx_GetRange(Optional ByVal Index As Variant, Optional ByVal GetCount As Variant, Optional ByVal ColumnIndexes As Variant) As ICollectionEx: Set ICollectionEx_GetRange = GetRange(Index, GetCount, ColumnIndexes): End Function
Private Function ICollectionEx_ToCSV(Optional ByVal Delimiter As String = ",", Optional ByVal InLocalFormat As Boolean = False) As String: ICollectionEx_ToCSV = ToCSV(Delimiter, InLocalFormat): End Function
Private Function ICollectionEx_ToJSON() As String: ICollectionEx_ToJSON = ToJSON(): End Function
Private Function ICollectionEx_ToExcel() As String: ICollectionEx_ToExcel = ToExcel(): End Function
Private Function ICollectionEx_CopyToRange(ByRef RangeObject As Excel.Range, Optional ByVal ApplyUserLocale As Boolean = True, Optional ByVal WriteHeaders As Boolean = True) As ICollectionEx: Set ICollectionEx_CopyToRange = CopyToRange(RangeObject, ApplyUserLocale, WriteHeaders): End Function





