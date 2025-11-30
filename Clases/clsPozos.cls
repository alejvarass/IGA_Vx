VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPozos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pColPozos               As New Collection
Private pDatos                  As clsDB

Public Property Set Datos(ByVal vNewValue As clsDB)
    Set pDatos = vNewValue
End Property

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = pColPozos.[_NewEnum]
End Function

Public Property Get Count() As Long
    Count = pColPozos.Count
End Property

Public Function colAgregar(ByVal objPozo As clsPozo) As Boolean
    
    On Error GoTo Error
    pColPozos.Add objPozo, objPozo.Key
    colAgregar = True
    
Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsPozos"
        pDatos.Error.Metodo = "colAgregar"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If

End Function

Public Function colQuitar(ByVal objPozo As clsPozo) As Boolean
    
    On Error GoTo Error
    pColPozos.Remove objPozo.Key
    colQuitar = True
    
Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsPozos"
        pDatos.Error.Metodo = "colQuitar"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Function


Public Sub ConsultarPozos()
    
    Dim StrSql                      As String
    Dim rs                          As Recordset
    Dim objPozo                     As clsPozo
    
    On Error GoTo Error
    
    StrSql = "SELECT "
    StrSql = StrSql & "IDPozo, "
    StrSql = StrSql & "NombrePozo, "
    StrSql = StrSql & "DiametroPozo, "
    StrSql = StrSql & "Supervisor, "
    StrSql = StrSql & "CordenadaX, "
    StrSql = StrSql & "CordenadaY, "
    StrSql = StrSql & "CordenadaZ, "
    StrSql = StrSql & "ProfundidadInicioControlG, "
    StrSql = StrSql & "IdEquipo, "
    StrSql = StrSql & "Activo, "
    StrSql = StrSql & "ProfundidadActual, "
    StrSql = StrSql & "IdFormulaTVD , "
    StrSql = StrSql & "IdPozoWeb , "
    StrSql = StrSql & "IdCuenca "
    StrSql = StrSql & "FROM "
    StrSql = StrSql & "Pozos; "
    
    Set rs = pDatos.EjecutarSeleccion(StrSql)
    If Not rs Is Nothing Then
        Do Until rs.EOF
            Set objPozo = New clsPozo
            If Not IsNull(rs!CordenadaX) Then
                objPozo.CordenadaX = rs!CordenadaX
            End If
            If Not IsNull(rs!CordenadaY) Then
                objPozo.CordenadaY = rs!CordenadaY
            End If
            If Not IsNull(rs!CordenadaZ) Then
                objPozo.CordenadaZ = rs!CordenadaZ
            End If
            objPozo.IdPozo = rs!IdPozo
            If Not IsNull(rs!NombrePozo) Then
                objPozo.NombrePozo = rs!NombrePozo
            End If
            If Not IsNull(rs!ProfundidadInicioControlG) Then
                objPozo.ProfundidadInicioControlG = rs!ProfundidadInicioControlG
            End If
            If Not IsNull(rs!Supervisor) Then
                objPozo.Supervisor = rs!Supervisor
            End If
            objPozo.IdEquipo = rs!IdEquipo
            objPozo.Activo = True
            If Not IsNull(rs!ProfundidadActual) Then
                objPozo.ProfundidadActual = rs!ProfundidadActual
            End If
            objPozo.IdFormulaTVD = rs!IdFormulaTVD
            objPozo.IdPozoWeb = rs!IdPozoWeb
            objPozo.IdCuenca = rs!IdCuenca
            Me.colAgregar objPozo
            
            rs.MoveNext
        Loop
    End If
    
Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsPozos"
        pDatos.Error.Metodo = "ConsultarPozos"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Sub

Public Property Get dbPozo( _
    ByVal Activo As Boolean, _
    ByVal IdPozo As Long _
    ) As clsPozo
    
    Dim StrSql                      As String
    Dim rs                          As Recordset
    Dim objPozo                     As clsPozo
    
    On Error GoTo Error
    
    StrSql = "SELECT "
    StrSql = StrSql & "IDPozo, "
    StrSql = StrSql & "NombrePozo, "
    StrSql = StrSql & "DiametroPozo, "
    StrSql = StrSql & "Supervisor, "
    StrSql = StrSql & "CordenadaX, "
    StrSql = StrSql & "CordenadaY, "
    StrSql = StrSql & "CordenadaZ, "
    StrSql = StrSql & "ProfundidadInicioControlG, "
    StrSql = StrSql & "IdEquipo, "
    StrSql = StrSql & "Activo, "
    StrSql = StrSql & "ProfundidadActual, "
    StrSql = StrSql & "IdFormulaTVD , "
    StrSql = StrSql & "IdPozoWeb , "
    StrSql = StrSql & "IdCuenca "
    StrSql = StrSql & "FROM "
    StrSql = StrSql & "Pozos "
    StrSql = StrSql & "WHERE "
    
    If Activo Then
        StrSql = StrSql & "Activo = True ;"
    Else
        StrSql = StrSql & "IdPozo = " & objPozo.IdPozo & ";"
    End If
    
    Set rs = pDatos.EjecutarSeleccion(StrSql)
    If Not rs Is Nothing Then
        Do Until rs.EOF
            Set objPozo = New clsPozo
            If Not IsNull(rs!CordenadaX) Then
                objPozo.CordenadaX = rs!CordenadaX
            End If
            If Not IsNull(rs!CordenadaY) Then
                objPozo.CordenadaY = rs!CordenadaY
            End If
            If Not IsNull(rs!CordenadaZ) Then
                objPozo.CordenadaZ = rs!CordenadaZ
            End If
            objPozo.IdPozo = rs!IdPozo
            If Not IsNull(rs!NombrePozo) Then
                objPozo.NombrePozo = rs!NombrePozo
            End If
            If Not IsNull(rs!ProfundidadInicioControlG) Then
                objPozo.ProfundidadInicioControlG = rs!ProfundidadInicioControlG
            End If
            If Not IsNull(rs!Supervisor) Then
                objPozo.Supervisor = rs!Supervisor
            End If
            objPozo.IdEquipo = rs!IdEquipo
            objPozo.Activo = True
            If Not IsNull(rs!ProfundidadActual) Then
                objPozo.ProfundidadActual = rs!ProfundidadActual
            End If
            objPozo.IdFormulaTVD = rs!IdFormulaTVD
            objPozo.IdPozoWeb = rs!IdPozoWeb
            objPozo.IdCuenca = rs!IdCuenca
            
            rs.MoveNext
        Loop
    End If
    
    Set dbPozo = objPozo
    
Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsPozos"
        pDatos.Error.Metodo = "dbPozo"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Property

Public Property Get colPozo( _
                Optional ByVal IdPozo As Long, _
                Optional ByVal Key As Variant _
                ) As clsPozo
    
    On Error GoTo Error
    
    If IsMissing(Key) Then
        Key = IdPozo & "ID"
    End If
    Set colPozo = pColPozos(Key)
    
Error:
    
    If Err.Number <> 0 Then
        If Err.Number <> 5 Then
        pDatos.Error.Objeto = "clsPozos"
        pDatos.Error.Metodo = "colPozo"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        End If
        Err.Clear
    End If
    
End Property

Public Property Get dbMax( _
                Optional ByVal IdPozo As Long _
                ) As Long
    
    Dim StrSql          As String
    Dim rs              As Recordset
    On Error GoTo Error
    
    StrSql = "Select max(IdPozo) as maximo " & _
    "from " & _
    "Pozos;"

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
        pDatos.Error.Objeto = "clsPozos"
        pDatos.Error.Metodo = "dbMax"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Property

Public Property Get colMax( _
                Optional ByVal IdPozo As Long _
                ) As Long
    
    Dim max As Long
    Dim objPozo As clsPozo
    
    On Error GoTo Error
    max = 0
    For Each objPozo In pColPozos
        If objPozo.IdPozo > max Then
            max = objPozo.IdPozo
        End If
    Next
    
    colMax = max
    
Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsPozos"
        pDatos.Error.Metodo = "colMax"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Property

Public Function dbAgregar(ByVal objPozo As clsPozo) As Boolean
    
    Dim StrSql                          As String
    
    On Error GoTo Error
    
    StrSql = "Insert into Pozos ( "
    StrSql = StrSql & "IDPozo, "
    StrSql = StrSql & "DiametroPozo, "
    StrSql = StrSql & "NombrePozo, "
    StrSql = StrSql & "Supervisor, "
    StrSql = StrSql & "CordenadaX, "
    StrSql = StrSql & "CordenadaY, "
    StrSql = StrSql & "CordenadaZ, "
    StrSql = StrSql & "ProfundidadInicioControlG, "
    StrSql = StrSql & "IdEquipo, "
    StrSql = StrSql & "Activo, "
    StrSql = StrSql & "ProfundidadActual, "
    StrSql = StrSql & "IdFormulaTVD, "
    StrSql = StrSql & "IdPozoWeb , "
    StrSql = StrSql & "IdCuenca "
    StrSql = StrSql & " ) Values ( "
    StrSql = StrSql & objPozo.IdPozo & " , "
    StrSql = StrSql & objPozo.DiametroPozo & " , '"
    StrSql = StrSql & objPozo.NombrePozo & "' , '"
    StrSql = StrSql & objPozo.Supervisor & "' , "
    StrSql = StrSql & objPozo.CordenadaX & " , "
    StrSql = StrSql & objPozo.CordenadaY & " , "
    StrSql = StrSql & objPozo.CordenadaZ & " , "
    StrSql = StrSql & objPozo.ProfundidadInicioControlG & " , '"
    StrSql = StrSql & objPozo.IdEquipo & "' , "
    StrSql = StrSql & IIf(objPozo.Activo, "TRUE ", "FALSE ") & " , "
    StrSql = StrSql & objPozo.ProfundidadActual & " , "
    StrSql = StrSql & objPozo.IdFormulaTVD & " , "
    StrSql = StrSql & objPozo.IdPozoWeb & " , "
    StrSql = StrSql & objPozo.IdCuenca & " "
    StrSql = StrSql & ");"
        
    dbAgregar = pDatos.DbEjecutar(StrSql)
    
Error:
    
    If Err.Number <> 0 Then
        dbAgregar = False
        pDatos.Error.Objeto = "clsPozos"
        pDatos.Error.Metodo = "dbAgregar"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Function

Public Function dbModificar(ByVal objPozo As clsPozo) As Boolean
    
    Dim StrSql                          As String
    
    On Error GoTo Error
    
    StrSql = "UPDATE Pozos "
    StrSql = StrSql & "SET "
    StrSql = StrSql & "DiametroPozo = " & objPozo.DiametroPozo & " , "
    StrSql = StrSql & "NombrePozo = '" & objPozo.NombrePozo & "' , "
    StrSql = StrSql & "Supervisor = '" & objPozo.Supervisor & "' , "
    StrSql = StrSql & "CordenadaX = " & objPozo.CordenadaX & " , "
    StrSql = StrSql & "CordenadaY = " & objPozo.CordenadaY & " , "
    StrSql = StrSql & "CordenadaZ = " & objPozo.CordenadaZ & " , "
    StrSql = StrSql & "ProfundidadInicioControlG = " & objPozo.ProfundidadInicioControlG & " , "
    StrSql = StrSql & "IdEquipo = '" & objPozo.IdEquipo & "' , "
    StrSql = StrSql & "Activo = " & IIf(objPozo.Activo, "TRUE ", "FALSE ") & " , "
    StrSql = StrSql & "ProfundidadActual = " & objPozo.ProfundidadActual & " , "
    StrSql = StrSql & "IdFormulaTVD = " & objPozo.IdFormulaTVD & " , "
    StrSql = StrSql & "IdPozoWeb = " & objPozo.IdPozoWeb & " , "
    StrSql = StrSql & "IdCuenca = " & objPozo.IdCuenca & " "
    StrSql = StrSql & "WHERE "
    StrSql = StrSql & "IDPozo = " & objPozo.IdPozo & ";"

    dbModificar = pDatos.DbEjecutar(StrSql)
    
Error:
    
    If Err.Number <> 0 Then
        dbModificar = False
        pDatos.Error.Objeto = "clsPozos"
        pDatos.Error.Metodo = "dbModificar"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Function

Public Sub dbEliminar( _
                ByVal objPozo As clsPozo, _
                ByRef ok As Boolean _
                )
    
    Dim StrSql                          As String
    
    On Error GoTo Error
    
    StrSql = "Delete " & pDatos.CaracterTodosLosCampos & " from Pozos " & _
    "WHERE " & _
    "IDPozo = " & objPozo.IdPozo & ";"

    ok = pDatos.DbEjecutar(StrSql)
    
Error:
    
    If Err.Number <> 0 Then
        ok = False
        pDatos.Error.Objeto = "clsPozos"
        pDatos.Error.Metodo = "dbEliminar"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Sub
