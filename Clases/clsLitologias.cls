VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLitologias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pColLitologias          As New Collection
Private pDatos                      As clsDB

Public Property Set Datos(ByVal vNewValue As clsDB)
    Set pDatos = vNewValue
End Property

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = pColLitologias.[_NewEnum]
End Function

Public Property Get Count() As Long
    Count = pColLitologias.Count
End Property

Public Function colAgregar(ByVal objLitologia As clsLitologia) As Boolean
    
    On Error GoTo Error
    pColLitologias.Add objLitologia, objLitologia.Key
    colAgregar = True
    
Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsLitologias"
        pDatos.Error.Metodo = "colAgregar"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If

End Function

Public Function colQuitar(ByVal objLitologia As clsLitologia) As Boolean
    
    On Error GoTo Error
    pColLitologias.Remove objLitologia.Key
    colQuitar = True
    
Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsLitologias"
        pDatos.Error.Metodo = "colQuitar"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Function

Public Sub ConsultarLitologias()
    
    Dim StrSql                      As String
    Dim rs                          As Recordset
    Dim objLitologia              As clsLitologia
    
    On Error GoTo Error
    StrSql = "SELECT "
    StrSql = StrSql & "CodigoLitologia, "
    StrSql = StrSql & "Nombrelitologia "
    StrSql = StrSql & "From "
    StrSql = StrSql & "iga_Litologias "
    StrSql = StrSql & "order by "
    StrSql = StrSql & "NombreLitologia;"

    Set rs = pDatos.EjecutarSeleccion(StrSql)
    If Not rs Is Nothing Then
        If Not (rs.BOF And rs.EOF) Then
            Do While Not rs.EOF
                
                Set objLitologia = New clsLitologia
                objLitologia.CodigoLitologia = rs!CodigoLitologia
                objLitologia.Nombrelitologia = rs!Nombrelitologia
                
                Me.colAgregar objLitologia
                
                rs.MoveNext
            Loop
        End If
    End If
    
Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsLitologias"
        pDatos.Error.Metodo = "ConsultarLitologias"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Sub

Public Property Get colLitologia( _
                Optional ByVal CodigoLitologia As Long, _
                Optional ByVal Key As Variant _
                ) As clsLitologia
    
    On Error GoTo Error
    
    If IsMissing(Key) Then
        Key = CodigoLitologia & "ID"
    End If
    Set colLitologia = pColLitologias(Key)
    
Error:
    
    If Err.Number <> 0 Then
        If Err.Number <> 5 Then
            pDatos.Error.Objeto = "clsLitologias"
            pDatos.Error.Metodo = "colLitologia"
            pDatos.Error.Number = Err.Number
            pDatos.Error.Description = Err.Description
            pDatos.Error.Mostrar
        End If
        Err.Clear
    End If
    
End Property

Public Property Get dbMax() As Long
    
    Dim StrSql          As String
    Dim rs              As Recordset
    On Error GoTo Error
    
    StrSql = "Select max(CodigoLitologia) as maximo "
    StrSql = StrSql & "from "
    StrSql = StrSql & "iga_Litologias;"

    Set rs = pDatos.EjecutarSeleccion(StrSql)
    If Not rs Is Nothing Then
        If Not (rs.BOF And rs.EOF) Then
            If Not IsNull(rs!maximo) Then
                dbMax = rs!maximo
            Else
                dbMax = 0
            End If
        End If
    End If

Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsLitologias"
        pDatos.Error.Metodo = "dbMax"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Property

Public Property Get colMax() As Long
    
    Dim max As Long
    Dim objLitologia As clsLitologia
    
    On Error GoTo Error
    max = 0
    For Each objLitologia In pColLitologias
        If objLitologia.CodigoLitologia > max Then
            max = objLitologia.CodigoLitologia
        End If
    Next
    
    colMax = max
    
Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsLitologias"
        pDatos.Error.Metodo = "colMax"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Property

Public Function dbAgregar(ByRef objLitologia As clsLitologia) As Boolean
    
    Dim StrSql                      As String
    
    On Error GoTo Error
    
    StrSql = "Insert into iga_Litologias ("
    StrSql = StrSql & "CodigoLitologia, "
    StrSql = StrSql & "NombreLitologia "
    StrSql = StrSql & ") Values ("
    StrSql = StrSql & objLitologia.CodigoLitologia & " , '"
    StrSql = StrSql & objLitologia.Nombrelitologia & "' "
    StrSql = StrSql & ");"

    dbAgregar = pDatos.DbEjecutar(StrSql)
    
Error:
    
    If Err.Number <> 0 Then
        dbAgregar = False
        pDatos.Error.Objeto = "clsLitologias"
        pDatos.Error.Metodo = "dbAgregar"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Function

Public Function dbModificar(ByVal objLitologia As clsLitologia) As Boolean
    
    Dim StrSql                      As String
    
    On Error GoTo Error
    StrSql = "Update iga_Litologias Set "
    StrSql = StrSql & "NombreLitologia = '" & objLitologia.Nombrelitologia & "' "
    StrSql = StrSql & "Where "
    StrSql = StrSql & "CodigoLitologia = " & objLitologia.CodigoLitologia & ";"

    dbModificar = pDatos.DbEjecutar(StrSql)
    
Error:
    
    If Err.Number <> 0 Then
        dbModificar = False
        pDatos.Error.Objeto = "clsLitologias"
        pDatos.Error.Metodo = "dbModificar"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Function



Public Function dbEliminar(ByVal objLitologia As clsLitologia) As Boolean
    
    Dim StrSql                      As String
    
    On Error GoTo Error
    StrSql = "Delete " & pDatos.CaracterTodosLosCampos & " from iga_Litologias "
    StrSql = StrSql & "Where "
    StrSql = StrSql & "CodigoLitologia = " & objLitologia.CodigoLitologia & ";"

    dbEliminar = pDatos.DbEjecutar(StrSql)
    
Error:
    
    If Err.Number <> 0 Then
        dbEliminar = False
        pDatos.Error.Objeto = "clsLitologias"
        pDatos.Error.Metodo = "dbEliminar"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Function

