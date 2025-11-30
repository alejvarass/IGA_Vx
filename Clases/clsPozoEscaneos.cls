VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPozoEscaneos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public pColPozoEscaneos                 As New Collection
Private pDatos                          As clsDB

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = pColPozoEscaneos.[_NewEnum]
End Function
    
Public Property Set Datos(ByVal vNewValue As clsDB)
    Set pDatos = vNewValue
End Property

Public Property Get Count() As Long
    Count = pColPozoEscaneos.Count
End Property
    
Public Function colAgregar(ByVal objPozoEscaneo As clsPozoEscaneo) As Boolean
    
    On Error GoTo Error
    
    Dim Indice                      As Long
    Dim objPozoEscaneoEnCol         As clsPozoEscaneo
    
    On Error GoTo Error
    
    Indice = 0
    'Agrega a la colección de escaneos del pozo el escaneo ingresado por valor en orden
    'Decreciente según profundidad de pozo. El escaneo menos profundo está más al final de la colección.
    
    For Each objPozoEscaneoEnCol In pColPozoEscaneos
        
        Indice = Indice + 1
        
        If objPozoEscaneo.ProfundidadPozo > objPozoEscaneoEnCol.ProfundidadPozo Then
            Exit For
'        ElseIf indice = pColPozoEscaneos.Count Then
'            indice = indice + 1
        End If
        
    Next
    
    If Indice = 0 Or Indice = pColPozoEscaneos.Count Then
        pColPozoEscaneos.Add objPozoEscaneo, objPozoEscaneo.Key
    Else
        pColPozoEscaneos.Add objPozoEscaneo, objPozoEscaneo.Key, Indice
    End If
    
    colAgregar = True
    
Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsPozoEscaneos"
        pDatos.Error.Metodo = "colAgregar"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If

End Function

Public Function colQuitar(ByVal objPozoEscaneo As clsPozoEscaneo) As Boolean

On Error GoTo Error
    
    pColPozoEscaneos.Remove objPozoEscaneo.Key
    colQuitar = True
    
Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsPozoEscaneos"
        pDatos.Error.Metodo = "colQuitar"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If

End Function

Public Property Get ColPozoEscaneo(Optional ByVal IdPozo As Long, _
                                    Optional ByVal IdPozoEscaneo As Long, _
                                    Optional ByVal Key As Variant) As clsPozoEscaneo
                                    
    On Error GoTo Error
    
    Set ColPozoEscaneo = pColPozoEscaneos(IIf(IsMissing(Key), IdPozo & IdPozoEscaneo & "ID", Key))
    
Error:
    
    If Err.Number <> 0 Then
        If Err.Number <> 5 Then
            pDatos.Error.Objeto = "clsPozoEscaneos"
            pDatos.Error.Metodo = "colPozoEscaneo"
            pDatos.Error.Number = Err.Number
            pDatos.Error.Description = Err.Description
            pDatos.Error.Mostrar
            Err.Clear
        End If
    End If


End Property

Public Property Get ColPozoEscaneoPorIndice(ByVal Indice As Long) As clsPozoEscaneo
                                    
    On Error GoTo Error
    
    Set ColPozoEscaneoPorIndice = pColPozoEscaneos(Indice)
    
Error:
    
    If Err.Number <> 0 Then
        If Err.Number <> 5 Then
            pDatos.Error.Objeto = "clsPozoEscaneos"
            pDatos.Error.Metodo = "ColPozoEscaneoPorIndice"
            pDatos.Error.Number = Err.Number
            pDatos.Error.Description = Err.Description
            pDatos.Error.Mostrar
            Err.Clear
        End If
    End If


End Property


Public Function colMaxProfPozoEscaneo() As Double
        
    Dim objPozoEscaneo                      As clsPozoEscaneo
    Dim maxProf                             As Double
    
    On Error GoTo Error
    
    maxProf = 0
    For Each objPozoEscaneo In pColPozoEscaneos
        If maxProf < objPozoEscaneo.ProfundidadPozo Then
            'la colección está ordenada descendentemente
            maxProf = objPozoEscaneo.ProfundidadPozo
            Exit For
        End If
    Next
    colMaxProfPozoEscaneo = maxProf
Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsPozoEscaneos"
        pDatos.Error.Metodo = "ConsultarPozoEscaneos"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Function

Public Function dbMaxProfPozoEscaneo(ByVal IdPozo As Long) As Double
        
    Dim StrSql                              As String
    Dim rs                                  As Recordset
    Dim objPozoEscaneo                      As clsPozoEscaneo
    
    On Error GoTo Error
    
    StrSql = "SELECT " & _
    "max(PozosEscaneos.ProfundidadPozo) AS MayorProf " & _
    "FROM " & _
    "PozosEscaneos " & _
    "WHERE " & _
    "pozosescaneos.idpozo = " & IdPozo & " AND PozosEscaneos.IdReporte = 1 AND " & _
    "PozosEscaneos.DescartadoPorCorreccionProfundidad = False;"
    
    Set rs = pDatos.EjecutarSeleccion(StrSql)
    
    If Not (rs.BOF And rs.EOF) Then
        If Not IsNull(rs!MayorProf) Then
            dbMaxProfPozoEscaneo = rs!MayorProf
        Else
            dbMaxProfPozoEscaneo = 0
        End If
    Else
        dbMaxProfPozoEscaneo = 0
    End If
    
Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsPozoEscaneos"
        pDatos.Error.Metodo = "ConsultarPozoEscaneos"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Function

Public Function ConsultarPozoEscaneos()
        
    Dim StrSql                              As String
    Dim rs                                  As Recordset
    Dim objPozoEscaneo                      As clsPozoEscaneo
    
    On Error GoTo Error
    
    StrSql = "SELECT "
    StrSql = StrSql & "Pozos.IDPozo, "
    StrSql = StrSql & "PozosEscaneos.IdPozoEscaneo, "
    StrSql = StrSql & "PozosEscaneos.IdReporte, "
    StrSql = StrSql & "PozosEscaneos.CO2, "
    StrSql = StrSql & "PozosEscaneos.Crono, "
    StrSql = StrSql & "PozosEscaneos.GasTotal, "
    StrSql = StrSql & "PozosEscaneos.ROP, "
    StrSql = StrSql & "PozosEscaneos.SH2, "
    StrSql = StrSql & "PozosEscaneos.ProfundidadPozo "
    StrSql = StrSql & "FROM "
    StrSql = StrSql & "Pozos INNER JOIN PozosEscaneos "
    StrSql = StrSql & "ON Pozos.IDPozo = PozosEscaneos.IdPozo "
    StrSql = StrSql & "WHERE "
    StrSql = StrSql & "Pozos.Activo = True "
    StrSql = StrSql & "AND "
    StrSql = StrSql & "PozosEscaneos.IdReporte = 1 "
    StrSql = StrSql & "AND "
    StrSql = StrSql & "PozosEscaneos.DescartadoPorCorreccionProfundidad = False "
    StrSql = StrSql & "ORDER BY "
    StrSql = StrSql & "PozosEscaneos.ProfundidadPozo "
    StrSql = StrSql & "desc; "

    Set rs = pDatos.EjecutarSeleccion(StrSql)
    
    Do Until rs.EOF
        
        Set objPozoEscaneo = Me.ColPozoEscaneo(rs!IdPozo, rs!IdPozoEscaneo)
        If objPozoEscaneo Is Nothing Then
            
            Set objPozoEscaneo = New clsPozoEscaneo
            
            objPozoEscaneo.CO2 = rs!CO2
            objPozoEscaneo.Crono = rs!Crono
            objPozoEscaneo.GasTotal = rs!GasTotal
            objPozoEscaneo.IdPozo = rs!IdPozo
            objPozoEscaneo.IdPozoEscaneo = rs!IdPozoEscaneo
            objPozoEscaneo.IdReporte = rs!IdReporte
            objPozoEscaneo.ProfundidadPozo = rs!ProfundidadPozo
            objPozoEscaneo.ROP = rs!ROP
            objPozoEscaneo.SH2 = rs!SH2
        
            Me.colAgregar objPozoEscaneo
            
        End If
        
             
        rs.MoveNext
        
    Loop
    
Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsPozoEscaneos"
        pDatos.Error.Metodo = "ConsultarPozoEscaneos"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Function


Public Function ConsultarPozoEscaneosPorProfundidadPosteriores(ByVal Profundidad As Double) As clsPozoEscaneos
        
    Dim StrSql                              As String
    Dim rs                                  As Recordset
    Dim objPozoEscaneo                      As clsPozoEscaneo
    Dim objPozoEscaneos                     As New clsPozoEscaneos
    
    On Error GoTo Error
    
    StrSql = "Select "
    StrSql = StrSql & "Pozos.IDPozo, "
    StrSql = StrSql & "PozosEscaneos.IdPozoEscaneo, "
    StrSql = StrSql & "PozosEscaneos.IdReporte, "
    StrSql = StrSql & "PozosEscaneos.CO2, "
    StrSql = StrSql & "PozosEscaneos.Crono, "
    StrSql = StrSql & "PozosEscaneos.GasTotal, "
    StrSql = StrSql & "PozosEscaneos.ROP, "
    StrSql = StrSql & "PozosEscaneos.SH2, "
    StrSql = StrSql & "PozosEscaneos.ProfundidadPozo "
    StrSql = StrSql & "FROM "
    StrSql = StrSql & "Pozos INNER JOIN PozosEscaneos "
    StrSql = StrSql & "ON Pozos.IDPozo = PozosEscaneos.IdPozo "
    StrSql = StrSql & "WHERE "
    StrSql = StrSql & "Pozos.Activo = True "
    StrSql = StrSql & "AND "
    StrSql = StrSql & "PozosEscaneos.IdReporte = 1 "
    StrSql = StrSql & "AND "
    StrSql = StrSql & "PozosEscaneos.ProfundidadPozo > " & Profundidad & " "
    StrSql = StrSql & "ORDER BY "
    StrSql = StrSql & "PozosEscaneos.ProfundidadPozo;"

    Set rs = pDatos.EjecutarSeleccion(StrSql)
    
    Do Until rs.EOF
        
        Set objPozoEscaneo = Me.ColPozoEscaneo(rs!IdPozo, rs!IdPozoEscaneo)
        If objPozoEscaneo Is Nothing Then
            
            Set objPozoEscaneo = New clsPozoEscaneo
            
            objPozoEscaneo.CO2 = rs!CO2
            objPozoEscaneo.Crono = rs!Crono
            objPozoEscaneo.GasTotal = rs!GasTotal
            objPozoEscaneo.IdPozo = rs!IdPozo
            objPozoEscaneo.IdPozoEscaneo = rs!IdPozoEscaneo
            objPozoEscaneo.IdReporte = rs!IdReporte
            objPozoEscaneo.ProfundidadPozo = rs!ProfundidadPozo
            objPozoEscaneo.ROP = rs!ROP
            objPozoEscaneo.SH2 = rs!SH2
        
            objPozoEscaneos.colAgregar objPozoEscaneo
        End If
        rs.MoveNext
    Loop
    
    Set ConsultarPozoEscaneosPorProfundidadPosteriores = objPozoEscaneos
    
Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsPozoEscaneos"
        pDatos.Error.Metodo = "ConsultarPozoEscaneosPorProfundidadPosteriores"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Function

Public Function ConsultarPozoEscaneosPorProfundidadAnteriores(ByVal Profundidad As Double) As clsPozoEscaneos
        
    Dim StrSql                              As String
    Dim rs                                  As Recordset
    Dim objPozoEscaneo                      As clsPozoEscaneo
    Dim objPozoEscaneos                     As New clsPozoEscaneos
    
    On Error GoTo Error
    
    StrSql = "Select "
    StrSql = StrSql & "Pozos.IDPozo, "
    StrSql = StrSql & "PozosEscaneos.IdPozoEscaneo, "
    StrSql = StrSql & "PozosEscaneos.IdReporte, "
    StrSql = StrSql & "PozosEscaneos.CO2, "
    StrSql = StrSql & "PozosEscaneos.Crono, "
    StrSql = StrSql & "PozosEscaneos.GasTotal, "
    StrSql = StrSql & "PozosEscaneos.ROP, "
    StrSql = StrSql & "PozosEscaneos.SH2, "
    StrSql = StrSql & "PozosEscaneos.ProfundidadPozo "
    StrSql = StrSql & "FROM "
    StrSql = StrSql & "Pozos INNER JOIN PozosEscaneos "
    StrSql = StrSql & "ON Pozos.IDPozo = PozosEscaneos.IdPozo "
    StrSql = StrSql & "WHERE "
    StrSql = StrSql & "Pozos.Activo = True "
    StrSql = StrSql & "AND "
    StrSql = StrSql & "PozosEscaneos.IdReporte = 1 "
    StrSql = StrSql & "AND "
    StrSql = StrSql & "PozosEscaneos.ProfundidadPozo < " & Profundidad & " "
    StrSql = StrSql & "ORDER BY "
    StrSql = StrSql & "PozosEscaneos.ProfundidadPozo desc;"

    Set rs = pDatos.EjecutarSeleccion(StrSql)
    
    Do Until rs.EOF
        
        Set objPozoEscaneo = Me.ColPozoEscaneo(rs!IdPozo, rs!IdPozoEscaneo)
        If objPozoEscaneo Is Nothing Then
            
            Set objPozoEscaneo = New clsPozoEscaneo
            
            objPozoEscaneo.CO2 = rs!CO2
            objPozoEscaneo.Crono = rs!Crono
            objPozoEscaneo.GasTotal = rs!GasTotal
            objPozoEscaneo.IdPozo = rs!IdPozo
            objPozoEscaneo.IdPozoEscaneo = rs!IdPozoEscaneo
            objPozoEscaneo.IdReporte = rs!IdReporte
            objPozoEscaneo.ProfundidadPozo = rs!ProfundidadPozo
            objPozoEscaneo.ROP = rs!ROP
            objPozoEscaneo.SH2 = rs!SH2
        
            objPozoEscaneos.colAgregar objPozoEscaneo
        End If
        rs.MoveNext
    Loop
    
    Set ConsultarPozoEscaneosPorProfundidadAnteriores = objPozoEscaneos
    
Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsPozoEscaneos"
        pDatos.Error.Metodo = "ConsultarPozoEscaneosPorProfundidadAnteriores"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Function

Public Sub ConsultarGasesEscaneos(ByVal IdPozo As Long, ByVal Desde As Double, ByVal Hasta As Double)
        
    Dim StrSql                              As String
    Dim rs                                  As Recordset
    Dim objPozoEscaneo                      As clsPozoEscaneo
    
    On Error GoTo Error
    
    StrSql = "Select "
    StrSql = StrSql & "PozosEscaneos.IDPozo, "
    StrSql = StrSql & "PozosEscaneos.IdPozoEscaneo, "
    StrSql = StrSql & "PozosEscaneos.IdReporte, "
    StrSql = StrSql & "PozosEscaneos.CO2, "
    StrSql = StrSql & "PozosEscaneos.Crono, "
    StrSql = StrSql & "PozosEscaneos.GasTotal, "
    StrSql = StrSql & "PozosEscaneos.ROP, "
    StrSql = StrSql & "PozosEscaneos.SH2, "
    StrSql = StrSql & "PozosEscaneos.ProfundidadPozo "
    StrSql = StrSql & "FROM "
    StrSql = StrSql & "PozosEscaneos "
    StrSql = StrSql & "WHERE "
    StrSql = StrSql & "PozosEscaneos.idpozo = " & IdPozo & "  "
    StrSql = StrSql & "AND "
    StrSql = StrSql & "PozosEscaneos.IdReporte = 1 "
    StrSql = StrSql & "AND "
    StrSql = StrSql & "PozosEscaneos.DescartadoPorCorreccionProfundidad = False "
    StrSql = StrSql & "AND "
    StrSql = StrSql & "PozosEscaneos.ProfundidadPozo >= " & Desde & " "
    StrSql = StrSql & "AND "
    StrSql = StrSql & "PozosEscaneos.ProfundidadPozo <= " & Hasta & " "
    StrSql = StrSql & "ORDER BY "
    StrSql = StrSql & "PozosEscaneos.ProfundidadPozo;"

    Set rs = pDatos.EjecutarSeleccion(StrSql)
    
    Do Until rs.EOF
        
        Set objPozoEscaneo = Me.ColPozoEscaneo(rs!IdPozo, rs!IdPozoEscaneo)
        
        If Not objPozoEscaneo Is Nothing Then
            
            objPozoEscaneo.CO2 = rs!CO2
            objPozoEscaneo.Crono = rs!Crono
            objPozoEscaneo.GasTotal = rs!GasTotal
            objPozoEscaneo.ROP = rs!ROP
            objPozoEscaneo.SH2 = rs!SH2
        
        End If
        
        rs.MoveNext
    
    Loop
    
Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsPozoEscaneos"
        pDatos.Error.Metodo = "ConsultarGasesEscaneos"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Sub

Public Function ConsultarPozoEscaneosPorProfundidad(ByVal ProfundidadInicio As Double, ByVal ProfundidadFinal As Double) As clsPozoEscaneos
        
    Dim StrSql                              As String
    Dim rs                                  As Recordset
    Dim objPozoEscaneo                      As clsPozoEscaneo
    
    On Error GoTo Error
    
    StrSql = "Select "
    StrSql = StrSql & "Pozos.IDPozo, "
    StrSql = StrSql & "PozosEscaneos.IdPozoEscaneo, "
    StrSql = StrSql & "PozosEscaneos.IdReporte, "
    StrSql = StrSql & "PozosEscaneos.CO2, "
    StrSql = StrSql & "PozosEscaneos.Crono, "
    StrSql = StrSql & "PozosEscaneos.GasTotal, "
    StrSql = StrSql & "PozosEscaneos.ROP, "
    StrSql = StrSql & "PozosEscaneos.SH2, "
    StrSql = StrSql & "PozosEscaneos.ProfundidadPozo "
    StrSql = StrSql & "FROM "
    StrSql = StrSql & "Pozos INNER JOIN PozosEscaneos "
    StrSql = StrSql & "ON Pozos.IDPozo = PozosEscaneos.IdPozo "
    StrSql = StrSql & "WHERE "
    StrSql = StrSql & "Pozos.Activo = True "
    StrSql = StrSql & "AND "
    StrSql = StrSql & "PozosEscaneos.IdReporte = 1 "
    StrSql = StrSql & "AND "
    StrSql = StrSql & "PozosEscaneos.ProfundidadPozo >= " & ProfundidadInicio & " "
    StrSql = StrSql & "AND "
    StrSql = StrSql & "PozosEscaneos.ProfundidadPozo <= " & ProfundidadFinal & " "
    StrSql = StrSql & "ORDER BY "
    StrSql = StrSql & "PozosEscaneos.ProfundidadPozo;"

    Set rs = pDatos.EjecutarSeleccion(StrSql)
    
    Do Until rs.EOF
        
        Set objPozoEscaneo = Me.ColPozoEscaneo(rs!IdPozo, rs!IdPozoEscaneo)
        If objPozoEscaneo Is Nothing Then
            
            Set objPozoEscaneo = New clsPozoEscaneo
            
            objPozoEscaneo.CO2 = rs!CO2
            objPozoEscaneo.Crono = rs!Crono
            objPozoEscaneo.GasTotal = rs!GasTotal
            objPozoEscaneo.IdPozo = rs!IdPozo
            objPozoEscaneo.IdPozoEscaneo = rs!IdPozoEscaneo
            objPozoEscaneo.IdReporte = rs!IdReporte
            objPozoEscaneo.ProfundidadPozo = rs!ProfundidadPozo
            objPozoEscaneo.ROP = rs!ROP
            objPozoEscaneo.SH2 = rs!SH2
        
            Me.colAgregar objPozoEscaneo
        End If
        rs.MoveNext
    Loop
    
Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsPozoEscaneos"
        pDatos.Error.Metodo = "ConsultarPozoEscaneosPorProfundidad"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Function

Public Function dbModificar(ByVal objPozoEscaneo As clsPozoEscaneo) As Boolean
    
    Dim StrSql                      As String
    
    On Error GoTo Error
    StrSql = "Update PozosEscaneos Set "
    StrSql = StrSql & "CO2 = " & objPozoEscaneo.CO2 & ", "
    StrSql = StrSql & "Crono = " & objPozoEscaneo.Crono & ", "
    StrSql = StrSql & "GasTotal = " & objPozoEscaneo.GasTotal & ", "
    StrSql = StrSql & "ROP = " & objPozoEscaneo.ROP & ", "
    StrSql = StrSql & "SH2 = " & objPozoEscaneo.SH2 & " "
    StrSql = StrSql & "WHERE "
    StrSql = StrSql & "idPozo = " & objPozoEscaneo.IdPozo & " "
    StrSql = StrSql & "And "
    StrSql = StrSql & "IdPozoEscaneo = " & objPozoEscaneo.IdPozoEscaneo & ";"
    
    dbModificar = pDatos.DbEjecutar(StrSql)
    
Error:
    
    If Err.Number <> 0 Then
        dbModificar = False
        pDatos.Error.Objeto = "clsPozoEscaneos"
        pDatos.Error.Metodo = "dbModificar"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Function

Public Sub Clear()

    Set pColPozoEscaneos = Nothing
    Set pColPozoEscaneos = New Collection
    
End Sub
