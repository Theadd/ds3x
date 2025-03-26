VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "dsQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder "ds3x.Collections.Extras"
' dbQuery Class Module

' --- PUBLIC PROPERTIES ---
' .LastError() As String                                                        - Contiene el Err.Description de la última llamada fallida a .TryQuery().
'
' --- PUBLIC METHODS ---
' .TryQuery(SqlQuery As String, Optional RS As ADODB.Recordset) As Boolean      - Ejecuta la query y devuelve si se produjo o no un error. Si se le pasa un objeto Recordset, este contendrá el resultado de la query.
' .Create(SqlQuery As String) As ADODB.Recordset                                - Ejecuta la consulta y devuelve un nuevo ADODB.Recordset VINCULADO a la conexión activa.
' .QueryAsync(SqlQuery As String)                                               - Envia la consulta al SQL Server para su ejecución y diréctamente devuelve el control a la aplicación.
' .Execute(SqlCommand As Variant, CommandParameters()) As ADODB.Recordset       - Permite ejecutar queries parametrizadas y devuelve un Recordset de un solo uso (adOpenForwardOnly).
'

Option Compare Database
Option Explicit

Private pConn As ADODB.Connection
Private pConnectionString As String
Private pLastError As String
Private pNoCloseOnTerminate As Boolean


' --- PROPERTIES ---

' Contiene el Err.Description de la última llamada fallida a TryQuery().
Public Property Get LastError() As String: LastError = pLastError: End Property

Public Property Get CloseOnTerminate() As Boolean: CloseOnTerminate = Not pNoCloseOnTerminate: End Property
Public Property Let CloseOnTerminate(ByVal Value As Boolean): pNoCloseOnTerminate = Not Value: End Property

Public Property Get Connection(Optional ByVal CursorLocation As ADODB.CursorLocationEnum = adUseServer) As ADODB.Connection
    If pConn Is Nothing Then OpenConnection CursorLocation
    Set Connection = pConn
End Property

Public Property Get ConnectionString() As String
    On Error GoTo Finally
    If pConnectionString = vbNullString Then
        If ObjPtr(Me) <> ObjPtr(dsQuery) Then
            pConnectionString = dsQuery.ConnectionString
        Else
            Dim s As String
            s = FileSystemLib.PathCombine(Application.CodeProject.Path, "package.json")
            If FileSystemLib.TryGetFileInAncestors(s, 3) Then
                If FileSystemLib.TryReadAllTextInFile(s, s) Then
                    pConnectionString = DictionaryEx.Create(s)("ds3x.dbQuery.ConnectionString")
                End If
            End If
        End If
    End If
Finally:
    ConnectionString = pConnectionString
End Property

Public Property Let ConnectionString(ByVal Value As String)
    On Error Resume Next
    pConnectionString = Value
    If Not pConn Is Nothing Then
        pConn.Close
        Set pConn = Nothing
    End If
    On Error GoTo 0
End Property

'@DefaultMember
Public Property Get Item(ByVal QueryString As String) As ADODB.Recordset
Attribute Item.VB_UserMemId = 0
    Dim rs As New ADODB.Recordset

    With rs
        Set .ActiveConnection = Me.Connection
        .Source = QueryString
        .CursorLocation = adUseClient
        .LockType = adLockReadOnly
        .CursorType = adOpenStatic
        .Open CursorType:=adOpenStatic, Options:=adCmdText
    End With

    Set Item = rs
End Property


' --- PUBLIC METHODS ---

' Ejecuta la query y devuelve si se produjo o no un error. Si se le pasa un objeto Recordset, este contendrá
' el resultado de la query y será DESVINCULADO de la conexión al SQL Server, por lo que los métodos del objeto
' Recordset que hagan uso del SQL Server no funcionarán, como .Requery() o .Filter, entre otros.
Public Function TryQuery(ByVal SqlQuery As String, Optional ByRef rs As ADODB.Recordset = Nothing) As Boolean
    On Error GoTo ErrorHandler

    If rs Is Nothing Then
        Me.Connection.Execute SqlQuery
    Else
        Set rs = New ADODB.Recordset

        With rs
            Set .ActiveConnection = Me.Connection
            .Source = SqlQuery
            .CursorLocation = adUseClient
            .LockType = adLockReadOnly
            .CursorType = adOpenStatic
            .Open
        End With

        Set rs.ActiveConnection = Nothing
    End If

    TryQuery = True
    Exit Function

ErrorHandler:
    pLastError = Err.Description
    TryQuery = False
End Function


Public Function Create(Optional ByRef ConnectionLike As Variant) As dsQuery
    Set Create = New dsQuery
    With Create
        If Not IsMissing(ConnectionLike) Then
            If IsObject(ConnectionLike) Then
                Dim c As ADODB.Connection
                Set c = ConnectionLike
                Create.ConnectionString = c.ConnectionString
            Else
                If VarType(ConnectionLike) = vbString Then
                    Create.ConnectionString = ConnectionLike
                End If
            End If
        Else
            Create.ConnectionString = ConnectionString
        End If
    End With
End Function

' Envia la consulta al SQL Server para su ejecución y diréctamente devuelve el control a la aplicación,
' sin esperar a que el SQL Server la haya ejecutado, teniendo o no éxito en ella, por lo que si se produce
' un error por parte del SQL Server en su ejecución, el error no será devuelto a la aplicación.
Public Sub QueryAsync(ByRef SqlQuery As String)
    Me.Connection.Execute SqlQuery, Options:=adAsyncExecute
End Sub

' Permite ejecutar queries parametrizadas y devuelve un Recordset de un solo uso (adOpenForwardOnly).
' Ideal para consultas simples. Permite reutilizar la consulta (ADODB.Command.CommandText) utilizando un
' conjunto distinto de parámetros, reemplazando alguno de los parámetros actuales (o todos).
'
' @EXAMPLE: Imprime el número de reportes con nivel WARNING que hay entre los 18 primeros registros.
'
'   Debug.Print dbQuery.Execute("SELECT COUNT(*) AS TOTAL FROM T_ATM_ISCS_APP_REPORTS WHERE ID <= ? AND REPORT_LEVEL = ?", 18, "WARNING")("TOTAL")
'
' Si en vez de imprimir directamente el valor consultado, asignas el recordset devuelto a una variable,
' te permitirá reutilizar la consulta modificando sus parámetros, volviendo a la consulta del ejemplo anterior:
'
'   Dim RS as ADODB.Recordset
'
'   Set RS = dbQuery.Execute("SELECT ... WHERE ID <= ? AND REPORT_LEVEL = ?", 18, "WARNING")
'   Debug.Print "Nº Warnings: " & RS("TOTAL")
'   Debug.Print "Nº Criticals: " & dbQuery.Execute(RS, , "CRITICAL")("TOTAL")
'   ' CommandParameters() también acepta un array de parámetros en lugar de un ParamArray.
'   Debug.Print "Nº Infos: " & dbQuery.Execute(RS, Array(12, "INFO"))("TOTAL")
'
Public Function Execute(ByRef SqlCommandOrRecordset As Variant, ParamArray CommandParameters() As Variant) As ADODB.Recordset
    Dim i As Long, cParams() As Variant

    If UBound(CommandParameters) = 0 And IsArray(CommandParameters(0)) Then
        cParams = CommandParameters(0)
    Else
        cParams = CommandParameters
    End If

    If IsObject(SqlCommandOrRecordset) Then
        Set SqlCommandOrRecordset.ActiveCommand.ActiveConnection = Me.Connection
        For i = 0 To UBound(cParams)
            If Not IsMissing(cParams(i)) Then
                If SqlCommandOrRecordset.ActiveCommand.Parameters(i).Type = adChar Then SqlCommandOrRecordset.ActiveCommand.Parameters(i).SIZE = Len(cParams(i))
                SqlCommandOrRecordset.ActiveCommand.Parameters(i).Value = cParams(i)
            End If
        Next i
        SqlCommandOrRecordset.Close
        SqlCommandOrRecordset.Open
        Set SqlCommandOrRecordset.ActiveCommand.ActiveConnection = Nothing
        Set Execute = SqlCommandOrRecordset
    Else
        With New ADODB.Command
            Set .ActiveConnection = Me.Connection
            .CommandText = SqlCommandOrRecordset
            Set Execute = .Execute(, cParams, adCmdText)
            Set .ActiveConnection = Nothing
        End With
    End If
End Function

Public Function CreateCommand(ByVal CommandText As String, Optional ByVal CommandType As ADODB.CommandTypeEnum = adCmdText, Optional ByVal CommandName As Variant) As ADODB.Command
    Set CreateCommand = New ADODB.Command

    With CreateCommand
        .Prepared = True
        .CommandType = CommandType
        .CommandTimeout = 30
        .ActiveConnection = Me.Connection
        .CommandText = CommandText
'        If Not IsMissing(CommandName) Then
'            .Name = CommandName
'        End If
    End With
End Function


' --- PRIVATE METHODS ---

Private Sub Class_Terminate()
    On Error Resume Next

    If Not pConn Is Nothing Then
        If Not pNoCloseOnTerminate Then pConn.Close
        Set pConn = Nothing
    End If
    On Error GoTo 0
End Sub

Private Sub OpenConnection(ByVal CursorLocation As ADODB.CursorLocationEnum)
    Set pConn = New ADODB.Connection
    With pConn
        .CursorLocation = CursorLocation
        .ConnectionString = ConnectionString
        .CommandTimeout = 240000   ' ~= 2d 18h
        .ConnectionTimeout = 10
        .Open
    End With
End Sub

Public Sub Dispose()
    On Error Resume Next
    If Not pConn Is Nothing Then
        pConn.Close
        Set pConn = Nothing
    End If
    On Error GoTo 0
End Sub
