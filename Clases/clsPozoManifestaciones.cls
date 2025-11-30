VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPozoManifestaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pColPozoManifestaciones          As New Collection
Private pDatos                      As clsDB

Public Property Set Datos(ByVal vNewValue As clsDB)
    Set pDatos = vNewValue
End Property

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = pColPozoManifestaciones.[_NewEnum]
End Function

Public Property Get Count() As Long
    Count = pColPozoManifestaciones.Count
End Property

Public Function colAgregar(ByVal objPozoManifestacion As clsPozoManifestacion) As Boolean
    
    On Error GoTo Error
    pColPozoManifestaciones.Add objPozoManifestacion, objPozoManifestacion.Key
    colAgregar = True
    
Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsPozoManifestaciones"
        pDatos.Error.Metodo = "colAgregar"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If

End Function

Public Function colQuitar(ByVal objPozoManifestacion As clsPozoManifestacion) As Boolean
    
    On Error GoTo Error
    pColPozoManifestaciones.Remove objPozoManifestacion.Key
    colQuitar = True
    
Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsPozoManifestaciones"
        pDatos.Error.Metodo = "colQuitar"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Function


Public Function dbPozoManifestacion(ByVal IdPozo As Long, ByVal IdManifestacion As Long) As clsPozoManifestacion
    
    Dim StrSql                      As String
    Dim rs                          As Recordset
    Dim objPozoManifestacion              As clsPozoManifestacion
    
    On Error GoTo Error
    StrSql = "SELECT "
    StrSql = StrSql & "IdPozo, "
    StrSql = StrSql & "IdManifestacion, "
    StrSql = StrSql & "IntervaloDesde, "
    StrSql = StrSql & "IntervaloHasta, "
    StrSql = StrSql & "CodigoComposicion, "
    StrSql = StrSql & "CodigoPorosidadVisual, "
    StrSql = StrSql & "GasTotal, "
    StrSql = StrSql & "BackG, "
    StrSql = StrSql & "C1, "
    StrSql = StrSql & "C2, "
    StrSql = StrSql & "C3, "
    StrSql = StrSql & "IC4, "
    StrSql = StrSql & "NC4, "
    StrSql = StrSql & "IC5, "
    StrSql = StrSql & "NC5, "
    StrSql = StrSql & "CodigoRastrosCantidad, "
    StrSql = StrSql & "CodigoColor, "
    StrSql = StrSql & "CodigoTipoDeRastro, "
    StrSql = StrSql & "CodigoFD, "
    StrSql = StrSql & "CodigoFE "
    StrSql = StrSql & "From "
    StrSql = StrSql & "iga_Pozos_Manifestaciones "
    StrSql = StrSql & "WHERE "
    StrSql = StrSql & "IdPozo = " & IdPozo & " "
    StrSql = StrSql & "AND "
    StrSql = StrSql & "IdManifestacion = " & IdManifestacion & ""

    Set rs = pDatos.EjecutarSeleccion(StrSql)
    If Not rs Is Nothing Then
        If Not (rs.BOF And rs.EOF) Then
            Set objPozoManifestacion = New clsPozoManifestacion
            
            objPozoManifestacion.IdPozo = rs!IdPozo
            objPozoManifestacion.IdManifestacion = rs!IdManifestacion
            objPozoManifestacion.IntervaloDesde = rs!IntervaloDesde
            objPozoManifestacion.IntervaloHasta = rs!IntervaloHasta
            objPozoManifestacion.CodigoComposicion = rs!CodigoComposicion
            objPozoManifestacion.CodigoPorosidadVisual = rs!CodigoPorosidadVisual
            objPozoManifestacion.GasTotal = rs!GasTotal
            objPozoManifestacion.BackG = rs!BackG
            objPozoManifestacion.C1 = rs!C1
            objPozoManifestacion.C2 = rs!C2
            objPozoManifestacion.C3 = rs!C3
            objPozoManifestacion.IC4 = rs!IC4
            objPozoManifestacion.NC4 = rs!NC4
            objPozoManifestacion.IC5 = rs!IC5
            objPozoManifestacion.NC5 = rs!NC5
            objPozoManifestacion.CodigoRastrosCantidad = rs!CodigoRastrosCantidad
            objPozoManifestacion.CodigoColor = rs!CodigoColor
            objPozoManifestacion.CodigoTipoDeRastro = rs!CodigoTipoDeRastro
            objPozoManifestacion.CodigoFD = rs!CodigoFD
            objPozoManifestacion.CodigoFE = rs!CodigoFE
            
            Set dbPozoManifestacion = objPozoManifestacion
            
        End If
    End If
    
Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsPozoManifestaciones"
        pDatos.Error.Metodo = "dbPozoManifestacion"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Function

Public Sub ConsultarPozoManifestaciones()
    
    Dim StrSql                      As String
    Dim rs                          As Recordset
    Dim objPozoManifestacion        As clsPozoManifestacion
    
    On Error GoTo Error
    StrSql = "SELECT "
    StrSql = StrSql & "IdPozo, "
    StrSql = StrSql & "IdManifestacion, "
    StrSql = StrSql & "IntervaloDesde, "
    StrSql = StrSql & "IntervaloHasta, "
    StrSql = StrSql & "CodigoComposicion, "
    StrSql = StrSql & "CodigoPorosidadVisual, "
    StrSql = StrSql & "GasTotal, "
    StrSql = StrSql & "BackG, "
    StrSql = StrSql & "C1, "
    StrSql = StrSql & "C2, "
    StrSql = StrSql & "C3, "
    StrSql = StrSql & "IC4, "
    StrSql = StrSql & "NC4, "
    StrSql = StrSql & "IC5, "
    StrSql = StrSql & "NC5, "
    StrSql = StrSql & "CodigoRastrosCantidad, "
    StrSql = StrSql & "CodigoColor, "
    StrSql = StrSql & "CodigoTipoDeRastro, "
    StrSql = StrSql & "CodigoFD, "
    StrSql = StrSql & "CodigoFE "
    StrSql = StrSql & "From "
    StrSql = StrSql & "iga_Pozos_Manifestaciones;"

    Set rs = pDatos.EjecutarSeleccion(StrSql)
    If Not rs Is Nothing Then
        If Not (rs.BOF And rs.EOF) Then
            Do While Not rs.EOF
                
                Set objPozoManifestacion = New clsPozoManifestacion
                objPozoManifestacion.IdPozo = rs!IdPozo
                objPozoManifestacion.IdManifestacion = rs!IdManifestacion
                objPozoManifestacion.IntervaloDesde = rs!IntervaloDesde
                objPozoManifestacion.IntervaloHasta = rs!IntervaloHasta
                objPozoManifestacion.CodigoComposicion = rs!CodigoComposicion
                objPozoManifestacion.CodigoPorosidadVisual = rs!CodigoPorosidadVisual
                objPozoManifestacion.GasTotal = rs!GasTotal
                objPozoManifestacion.BackG = rs!BackG
                objPozoManifestacion.C1 = rs!C1
                objPozoManifestacion.C2 = rs!C2
                objPozoManifestacion.C3 = rs!C3
                objPozoManifestacion.IC4 = rs!IC4
                objPozoManifestacion.NC4 = rs!NC4
                objPozoManifestacion.IC5 = rs!IC5
                objPozoManifestacion.NC5 = rs!NC5
                objPozoManifestacion.CodigoRastrosCantidad = rs!CodigoRastrosCantidad
                objPozoManifestacion.CodigoColor = rs!CodigoColor
                objPozoManifestacion.CodigoTipoDeRastro = rs!CodigoTipoDeRastro
                objPozoManifestacion.CodigoFD = rs!CodigoFD
                objPozoManifestacion.CodigoFE = rs!CodigoFE
                
                Me.colAgregar objPozoManifestacion
                
                rs.MoveNext
            Loop
        End If
    End If
    
Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsPozoManifestaciones"
        pDatos.Error.Metodo = "ConsultarPozoManifestaciones"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Sub

Public Property Get colPozoManifestacion( _
                Optional ByVal IdPozo As Long, _
                Optional ByVal IdManifestacion As Long, _
                Optional ByVal Key As Variant _
                ) As clsPozoManifestacion
    
    On Error GoTo Error
    
    If IsMissing(Key) Then
        Key = IdPozo & IdManifestacion & "ID"
    End If
    Set colPozoManifestacion = pColPozoManifestaciones(Key)
    
Error:
    
    If Err.Number <> 0 Then
        If Err.Number <> 5 Then
            pDatos.Error.Objeto = "clsPozoManifestaciones"
            pDatos.Error.Metodo = "colPozoManifestacion"
            pDatos.Error.Number = Err.Number
            pDatos.Error.Description = Err.Description
            pDatos.Error.Mostrar
        End If
        Err.Clear
    End If
    
End Property

Public Property Get dbMax(ByVal IdPozo As Long) As Long
    
    Dim StrSql          As String
    Dim rs              As Recordset
    On Error GoTo Error
    
    StrSql = "Select max(IdManifestacion) as maximo "
    StrSql = StrSql & "from "
    StrSql = StrSql & "iga_Pozos_Manifestaciones "
    StrSql = StrSql & "Where "
    StrSql = StrSql & "IdPozo =  " & IdPozo & ";"

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
        pDatos.Error.Objeto = "clsPozoManifestaciones"
        pDatos.Error.Metodo = "dbMax"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Property

Public Property Get colMax() As Long
    
    Dim max As Long
    Dim objPozoManifestacion As clsPozoManifestacion
    
    On Error GoTo Error
    max = 0
    For Each objPozoManifestacion In pColPozoManifestaciones
        If objPozoManifestacion.IdManifestacion > max Then
            max = objPozoManifestacion.IdManifestacion
        End If
    Next
    
    colMax = max
    
Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsPozoManifestaciones"
        pDatos.Error.Metodo = "colMax"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Property

Public Function dbAgregar(ByRef objPozoManifestacion As clsPozoManifestacion) As Boolean
    
    Dim StrSql                      As String
    
    On Error GoTo Error
    
    StrSql = "Insert into iga_Pozos_Manifestaciones ("
    StrSql = StrSql & "IdPozo, "
    StrSql = StrSql & "IdManifestacion, "
    StrSql = StrSql & "IntervaloDesde, "
    StrSql = StrSql & "IntervaloHasta, "
    StrSql = StrSql & "CodigoComposicion, "
    StrSql = StrSql & "CodigoPorosidadVisual, "
    StrSql = StrSql & "GasTotal, "
    StrSql = StrSql & "BackG, "
    StrSql = StrSql & "C1, "
    StrSql = StrSql & "C2, "
    StrSql = StrSql & "C3, "
    StrSql = StrSql & "IC4, "
    StrSql = StrSql & "NC4, "
    StrSql = StrSql & "IC5, "
    StrSql = StrSql & "NC5, "
    StrSql = StrSql & "CodigoRastrosCantidad, "
    StrSql = StrSql & "CodigoColor, "
    StrSql = StrSql & "CodigoTipoDeRastro, "
    StrSql = StrSql & "CodigoFD, "
    StrSql = StrSql & "CodigoFE "
    StrSql = StrSql & ") Values ("
    StrSql = StrSql & objPozoManifestacion.IdPozo & ", "
    StrSql = StrSql & objPozoManifestacion.IdManifestacion & ", "
    StrSql = StrSql & objPozoManifestacion.IntervaloDesde & ", "
    StrSql = StrSql & objPozoManifestacion.IntervaloHasta & ", "
    StrSql = StrSql & objPozoManifestacion.CodigoComposicion & ", "
    StrSql = StrSql & objPozoManifestacion.CodigoPorosidadVisual & ", "
    StrSql = StrSql & objPozoManifestacion.GasTotal & ", "
    StrSql = StrSql & objPozoManifestacion.BackG & ", "
    StrSql = StrSql & objPozoManifestacion.C1 & ", "
    StrSql = StrSql & objPozoManifestacion.C2 & ", "
    StrSql = StrSql & objPozoManifestacion.C3 & ", "
    StrSql = StrSql & objPozoManifestacion.IC4 & ", "
    StrSql = StrSql & objPozoManifestacion.NC4 & ", "
    StrSql = StrSql & objPozoManifestacion.IC5 & ", "
    StrSql = StrSql & objPozoManifestacion.NC5 & ", "
    StrSql = StrSql & objPozoManifestacion.CodigoRastrosCantidad & ", "
    StrSql = StrSql & objPozoManifestacion.CodigoColor & ", "
    StrSql = StrSql & objPozoManifestacion.CodigoTipoDeRastro & ", "
    StrSql = StrSql & objPozoManifestacion.CodigoFD & ", "
    StrSql = StrSql & objPozoManifestacion.CodigoFE & " "
    StrSql = StrSql & ");"

    dbAgregar = pDatos.DbEjecutar(StrSql)
    
Error:
    
    If Err.Number <> 0 Then
        dbAgregar = False
        pDatos.Error.Objeto = "clsPozoManifestaciones"
        pDatos.Error.Metodo = "dbAgregar"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Function

Public Function dbModificar(ByVal objPozoManifestacion As clsPozoManifestacion) As Boolean
    
    Dim StrSql                      As String
    
    On Error GoTo Error
    StrSql = "Update iga_Pozos_Manifestaciones Set "
    StrSql = StrSql & "IntervaloDesde = " & objPozoManifestacion.IntervaloDesde & ", "
    StrSql = StrSql & "IntervaloHasta = " & objPozoManifestacion.IntervaloHasta & ", "
    StrSql = StrSql & "CodigoComposicion = " & objPozoManifestacion.CodigoComposicion & ", "
    StrSql = StrSql & "CodigoPorosidadVisual = " & objPozoManifestacion.CodigoPorosidadVisual & ", "
    StrSql = StrSql & "GasTotal = " & objPozoManifestacion.GasTotal & ", "
    StrSql = StrSql & "BackG = " & objPozoManifestacion.BackG & ", "
    StrSql = StrSql & "C1 = " & objPozoManifestacion.C1 & ", "
    StrSql = StrSql & "C2 = " & objPozoManifestacion.C2 & ", "
    StrSql = StrSql & "C3 = " & objPozoManifestacion.C3 & ", "
    StrSql = StrSql & "IC4 = " & objPozoManifestacion.IC4 & ", "
    StrSql = StrSql & "NC4 = " & objPozoManifestacion.NC4 & ", "
    StrSql = StrSql & "IC5 = " & objPozoManifestacion.IC5 & ", "
    StrSql = StrSql & "NC5 = " & objPozoManifestacion.NC5 & ", "
    StrSql = StrSql & "CodigoRastrosCantidad = " & objPozoManifestacion.CodigoRastrosCantidad & ", "
    StrSql = StrSql & "CodigoColor = " & objPozoManifestacion.CodigoColor & ", "
    StrSql = StrSql & "CodigoTipoDeRastro = " & objPozoManifestacion.CodigoTipoDeRastro & ", "
    StrSql = StrSql & "CodigoFD = " & objPozoManifestacion.CodigoFD & ", "
    StrSql = StrSql & "CodigoFE = " & objPozoManifestacion.CodigoFE & " "
    StrSql = StrSql & "Where "
    StrSql = StrSql & "IdPozo = " & objPozoManifestacion.IdPozo & " "
    StrSql = StrSql & "AND "
    StrSql = StrSql & "IdManifestacion = " & objPozoManifestacion.IdManifestacion & ";"

    dbModificar = pDatos.DbEjecutar(StrSql)
    
Error:
    
    If Err.Number <> 0 Then
        dbModificar = False
        pDatos.Error.Objeto = "clsPozoManifestaciones"
        pDatos.Error.Metodo = "dbModificar"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Function



Public Function dbEliminar(ByVal objPozoManifestacion As clsPozoManifestacion) As Boolean
    
    Dim StrSql                      As String
    
    On Error GoTo Error
    StrSql = "Delete " & pDatos.CaracterTodosLosCampos & " from iga_Pozos_Manifestaciones "
    StrSql = StrSql & "Where "
    StrSql = StrSql & "IdPozo = " & objPozoManifestacion.IdPozo & " "
    StrSql = StrSql & "AND "
    StrSql = StrSql & "IdManifestacion = " & objPozoManifestacion.IdManifestacion & ";"

    dbEliminar = pDatos.DbEjecutar(StrSql)
    
Error:
    
    If Err.Number <> 0 Then
        dbEliminar = False
        pDatos.Error.Objeto = "clsPozoManifestaciones"
        pDatos.Error.Metodo = "dbEliminar"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Function

