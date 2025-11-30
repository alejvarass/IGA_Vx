VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPozoGasTiempos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pColPozoGasTiempos          As New Collection
Private pDatos                      As clsDB

Public Property Set Datos(ByVal vNewValue As clsDB)
    Set pDatos = vNewValue
End Property

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = pColPozoGasTiempos.[_NewEnum]
End Function

Public Property Get Count() As Long
    Count = pColPozoGasTiempos.Count
End Property

Public Function colAgregar(ByVal objPozoGasTiempo As clsPozoGasTiempo) As Boolean
    
    On Error GoTo Error
    pColPozoGasTiempos.Add objPozoGasTiempo, objPozoGasTiempo.Key
    colAgregar = True
    
Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsPozoGasTiempos"
        pDatos.Error.Metodo = "colAgregar"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If

End Function

Public Function colQuitar(ByVal objPozoGasTiempo As clsPozoGasTiempo) As Boolean
    
    On Error GoTo Error
    pColPozoGasTiempos.Remove objPozoGasTiempo.Key
    colQuitar = True
    
Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsPozoGasTiempos"
        pDatos.Error.Metodo = "colQuitar"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Function

Public Sub dbPozoGasesTiempo(ByVal idPozo As Long, ByVal FechaInicio As Date, ByVal FechaFin As Date)
    
    Dim StrSql                      As String
    Dim rs                          As Recordset
    Dim objPozoGasTiempo            As clsPozoGasTiempo
    
    On Error GoTo Error
    StrSql = "SELECT "
    StrSql = StrSql & "IDPozo, "
    StrSql = StrSql & "fecha, "
    StrSql = StrSql & "GasTotal, "
    StrSql = StrSql & "CO2, "
    StrSql = StrSql & "SH2, "
    StrSql = StrSql & "Comentario "
    StrSql = StrSql & "From "
    StrSql = StrSql & "iga_pozos_gases_tiempo "
    StrSql = StrSql & "Where "
    StrSql = StrSql & "idpozo = " & idPozo & " AND "
    StrSql = StrSql & "Fecha >=  " & pDatos.CaracterFecha & Format(FechaInicio, pDatos.FormatoFechaLarga) & pDatos.CaracterFecha & " AND "
    StrSql = StrSql & "Fecha <=  " & pDatos.CaracterFecha & Format(FechaFin, pDatos.FormatoFechaLarga) & pDatos.CaracterFecha & " "
    StrSql = StrSql & "order by "
    StrSql = StrSql & "fecha"

    Set rs = pDatos.EjecutarSeleccion(StrSql)
    If Not rs Is Nothing Then
        If Not (rs.BOF And rs.EOF) Then
            Do While Not rs.EOF
                
                Set objPozoGasTiempo = New clsPozoGasTiempo
                
                objPozoGasTiempo.idPozo = rs!idPozo
                objPozoGasTiempo.fecha = rs!fecha
                objPozoGasTiempo.GasTotal = rs!GasTotal
                objPozoGasTiempo.CO2 = rs!CO2
                objPozoGasTiempo.SH2 = rs!SH2
                If Not IsNull(rs!Comentario) Then
                    objPozoGasTiempo.Comentario = rs!Comentario
                End If
                Me.colAgregar objPozoGasTiempo
                
                rs.MoveNext
            Loop
        End If
    End If
    
Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsPozoGasTiempos"
        pDatos.Error.Metodo = "dbPozoGasesTiempo"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Sub

Public Property Get colPozoGasTiempo( _
                Optional ByVal idPozo As Long, _
                Optional ByVal fecha As Date, _
                Optional ByVal Key As Variant _
                ) As clsPozoGasTiempo
    
    On Error GoTo Error
    
    If IsMissing(Key) Then
        Key = idPozo & fecha & "ID"
    End If
    Set colPozoGasTiempo = pColPozoGasTiempos(Key)
    
Error:
    
    If Err.Number <> 0 Then
        If Err.Number <> 5 Then
            pDatos.Error.Objeto = "clsPozoGasTiempos"
            pDatos.Error.Metodo = "colPozoGasTiempo"
            pDatos.Error.Number = Err.Number
            pDatos.Error.Description = Err.Description
            pDatos.Error.Mostrar
        End If
        Err.Clear
    End If
    
End Property

Public Function dbAgregar(ByRef objPozoGasTiempo As clsPozoGasTiempo) As Boolean
    
    Dim StrSql                      As String
    
    On Error GoTo Error
    
    StrSql = "Insert into  iga_pozos_gases_tiempo ("
    StrSql = StrSql & "IDPozo, "
    StrSql = StrSql & "fecha, "
    StrSql = StrSql & "GasTotal, "
    StrSql = StrSql & "CO2, "
    StrSql = StrSql & "SH2, "
    StrSql = StrSql & "Comentario "
    StrSql = StrSql & " ) Values ( "
    StrSql = StrSql & objPozoGasTiempo.idPozo & ", "
    StrSql = StrSql & pDatos.CaracterFecha & Format(objPozoGasTiempo.fecha, pDatos.FormatoFechaLarga) & pDatos.CaracterFecha & ", "
    StrSql = StrSql & objPozoGasTiempo.GasTotal & ", "
    StrSql = StrSql & objPozoGasTiempo.CO2 & ", "
    StrSql = StrSql & objPozoGasTiempo.SH2 & ", '"
    StrSql = StrSql & objPozoGasTiempo.Comentario & "' "
    StrSql = StrSql & ");"


    dbAgregar = pDatos.DbEjecutar(StrSql)
    
Error:
    
    If Err.Number <> 0 Then
        dbAgregar = False
        pDatos.Error.Objeto = "clsPozoGasTiempos"
        pDatos.Error.Metodo = "dbAgregar"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Function

Public Function dbModificar(ByVal objPozoGasTiempo As clsPozoGasTiempo) As Boolean
    
    Dim StrSql                      As String
    
    On Error GoTo Error
    
    StrSql = "Update iga_pozos_gases_tiempo SET "
    StrSql = StrSql & "GasTotal = " & objPozoGasTiempo.GasTotal & ", "
    StrSql = StrSql & "CO2 = " & objPozoGasTiempo.CO2 & ", "
    StrSql = StrSql & "SH2 = " & objPozoGasTiempo.SH2 & ", "
    StrSql = StrSql & "Comentario = '" & objPozoGasTiempo.Comentario & "' "
    StrSql = StrSql & " Where "
    StrSql = StrSql & "IDPozo = " & objPozoGasTiempo.idPozo & " "
    StrSql = StrSql & "AND "
    StrSql = StrSql & "fecha = " & pDatos.CaracterFecha & Format(objPozoGasTiempo.fecha, pDatos.FormatoFechaLarga) & pDatos.CaracterFecha & ";"
    
    dbModificar = pDatos.DbEjecutar(StrSql)
    
Error:
    
    If Err.Number <> 0 Then
        dbModificar = False
        pDatos.Error.Objeto = "clsPozoGasTiempos"
        pDatos.Error.Metodo = "dbModificar"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Function



Public Function dbEliminar(ByVal objPozoGasTiempo As clsPozoGasTiempo) As Boolean
    
    Dim StrSql                      As String
    
    On Error GoTo Error
    
    StrSql = "delete " & pDatos.CaracterTodosLosCampos & " from iga_pozos_gases_tiempo "
    StrSql = StrSql & " Where "
    StrSql = StrSql & "IDPozo = " & objPozoGasTiempo.idPozo & " "
    StrSql = StrSql & "AND "
    StrSql = StrSql & "fecha = " & pDatos.CaracterFecha & Format(objPozoGasTiempo.fecha, pDatos.FormatoFechaLarga) & pDatos.CaracterFecha & ";"

    dbEliminar = pDatos.DbEjecutar(StrSql)
    
Error:
    
    If Err.Number <> 0 Then
        dbEliminar = False
        pDatos.Error.Objeto = "clsPozoGasTiempos"
        pDatos.Error.Metodo = "dbEliminar"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Function
