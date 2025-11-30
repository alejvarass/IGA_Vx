VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPozoAnalisis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public pColPozoAnalisis                 As New Collection
Private pDatos                          As clsDB

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = pColPozoAnalisis.[_NewEnum]
End Function
    
Public Property Set Datos(ByVal vNewValue As clsDB)
    Set pDatos = vNewValue
End Property

Public Property Get Count() As Long
    Count = pColPozoAnalisis.Count
End Property
    
Public Function colAgregar(ByVal objPozoAnalis As clsPozoAnalis) As Boolean
    
    On Error GoTo Error
    
    pColPozoAnalisis.Add objPozoAnalis, objPozoAnalis.Key
    colAgregar = True
    
Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsPozoAnalisis"
        pDatos.Error.Metodo = "colAgregar"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If

End Function

Public Function colQuitar(ByVal objPozoAnalis As clsPozoAnalis) As Boolean

On Error GoTo Error
    
    pColPozoAnalisis.Remove objPozoAnalis.Key
    colQuitar = True
    
Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsPozoAnalisis"
        pDatos.Error.Metodo = "colQuitar"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If

End Function

Public Property Get ColPozoAnalisis() As clsPozoAnalisis

    ColPozoAnalisis = pColPozoAnalisis

End Property
                                                
                                                
Public Property Get ColPozoAnalis(Optional ByVal IdPozo As Long, _
                                    Optional ByVal NumeroAnalisis As Long, _
                                    Optional ByVal Key As Variant) As clsPozoAnalis
                                    
    On Error GoTo Error
    If IsMissing(Key) Then
        Key = IdPozo & NumeroAnalisis & "ID"
    End If
    Set ColPozoAnalis = pColPozoAnalisis(Key)
    
Error:
    
    If Err.Number <> 0 Then
        If Err.Number <> 5 Then
            pDatos.Error.Objeto = "clsPozoAnalisis"
            pDatos.Error.Metodo = "colPozoAnalis"
            pDatos.Error.Number = Err.Number
            pDatos.Error.Description = Err.Description
            pDatos.Error.Mostrar
        End If
        Err.Clear
        
    End If


End Property

Public Sub Clear()
    
    Set pColPozoAnalisis = Nothing
    Set pColPozoAnalisis = New Collection

End Sub
Public Function ConsultarPozoAnalisis( _
    IdPozo As Long, _
    Optional ByVal DesdeProf As Variant, _
    Optional ByVal HastaProf As Variant, _
    Optional ByVal Seleccionado As Variant, _
    Optional ByVal idTipoAnalisis As Variant, _
    Optional ByVal descendenteValorGas As Variant _
    )
        
    Dim StrSql                              As String
    Dim rs                                  As Recordset
    Dim objPozoAnalis                       As clsPozoAnalis
    Dim objPozoAnalisComponente             As ClsPozoAnalisComponente
    
    On Error GoTo Error
    
    StrSql = "SELECT "
    StrSql = StrSql & "iga_pozos_analisis.IdPozo , "
    StrSql = StrSql & "iga_pozos_analisis.NumeroAnalisis, "
    StrSql = StrSql & "iga_pozos_analisis.archivo, "
    StrSql = StrSql & "iga_pozos_analisis.Fecha, "
    StrSql = StrSql & "iga_pozos_analisis.Profundidad, "
    StrSql = StrSql & "iga_pozos_analisis.CO2, "
    StrSql = StrSql & "iga_pozos_analisis.SH2, "
    StrSql = StrSql & "iga_pozos_analisis.GasTotal, "
    StrSql = StrSql & "iga_pozos_analisis.GasTotalCromatografico, "
    StrSql = StrSql & "iga_pozos_analisis.CodigoTipoDeAnalisis, "
    StrSql = StrSql & "iga_pozos_analisis.EFast, ENormal, "
    StrSql = StrSql & "iga_pozos_analisis.Observaciones, "
    StrSql = StrSql & "iga_pozos_analisis.Seleccionado, "
    StrSql = StrSql & "iga_pozos_analisis.Bar2, "
    StrSql = StrSql & "iga_pozos_analisis.Bar3, "
    StrSql = StrSql & "iga_pozos_analisis.Bar4, "
    StrSql = StrSql & "iga_pozos_analisis.Bar5, "
    StrSql = StrSql & "iga_pozos_analisis.Cous1, "
    StrSql = StrSql & "iga_pozos_analisis.Cous2, "
    StrSql = StrSql & "iga_pozos_analisis.WH, "
    StrSql = StrSql & "iga_pozos_analisis.BH, "
    StrSql = StrSql & "iga_pozos_analisis.CH, "
    StrSql = StrSql & "iga_pozos_analisis.Geo1, "
    StrSql = StrSql & "iga_pozos_analisis.Geo2, "
    StrSql = StrSql & "iga_pozos_analisis.Geo3, "
    StrSql = StrSql & "iga_pozos_analisis.Geo4, "
    StrSql = StrSql & "iga_pozos_analisis.SnGeo, "
    StrSql = StrSql & "iga_pozos_analisis.C1, "
    StrSql = StrSql & "iga_pozos_analisis.C2, "
    StrSql = StrSql & "iga_pozos_analisis.C3, "
    StrSql = StrSql & "iga_pozos_analisis.C4, "
    StrSql = StrSql & "iga_pozos_analisis.C5, "
    StrSql = StrSql & "iga_pozos_analisis.IC4, "
    StrSql = StrSql & "iga_pozos_analisis.NC4, "
    StrSql = StrSql & "iga_pozos_analisis.IC5, "
    StrSql = StrSql & "iga_pozos_analisis.NC5, "
    StrSql = StrSql & "iga_pozos_analisis_componentes.NumeroComponente, "
    StrSql = StrSql & "iga_pozos_analisis_componentes.Component, "
    StrSql = StrSql & "iga_pozos_analisis_componentes.Retention, "
    StrSql = StrSql & "iga_pozos_analisis_componentes.Area, "
    StrSql = StrSql & "iga_pozos_analisis_componentes.Externo, "
    StrSql = StrSql & "iga_pozos_analisis_componentes.Units, "
    StrSql = StrSql & "iga_pozos_analisis_componentes.NormArea  "
    StrSql = StrSql & "FROM "
    StrSql = StrSql & "iga_pozos_analisis left join iga_pozos_analisis_componentes on "
    StrSql = StrSql & "iga_pozos_analisis.IdPozo = iga_pozos_analisis_componentes.IdPozo "
    StrSql = StrSql & "AND "
    StrSql = StrSql & "iga_pozos_analisis.NumeroAnalisis = iga_pozos_analisis_componentes.NumeroAnalisis "
    StrSql = StrSql & "WHERE "
    StrSql = StrSql & "   iga_pozos_analisis.IDPozo = " & IdPozo & " "
    If Not IsMissing(DesdeProf) Then
        StrSql = StrSql & " AND  iga_pozos_analisis.Profundidad >= " & DesdeProf
    End If
    If Not IsMissing(HastaProf) Then
        StrSql = StrSql & " AND  iga_pozos_analisis.Profundidad <= " & HastaProf & " "
    End If
    If Not IsMissing(Seleccionado) Then
        StrSql = StrSql & " AND  iga_pozos_analisis.seleccionado = " & IIf(Seleccionado, "TRUE", "FALSE") & " "
    End If
    If Not IsMissing(idTipoAnalisis) Then
        StrSql = StrSql & " AND  iga_pozos_analisis.CodigoTipoDeAnalisis = " & idTipoAnalisis & " "
    End If
    StrSql = StrSql & "ORDER BY "
    If IsMissing(descendenteValorGas) Then
        StrSql = StrSql & " NumeroAnalisis "
    Else
        StrSql = StrSql & " iga_pozos_analisis.GasTotalCromatografico DESC "
        StrSql = StrSql & "limit 1"
    End If
    StrSql = StrSql & ";"
    Set rs = pDatos.EjecutarSeleccion(StrSql)
    
    Do Until rs.EOF
        Set objPozoAnalis = Me.ColPozoAnalis(rs!IdPozo, rs!NumeroAnalisis)
        If objPozoAnalis Is Nothing Then
        
            Set objPozoAnalis = New clsPozoAnalis
            
            objPozoAnalis.archivo = rs!archivo
            objPozoAnalis.Bar2 = rs!Bar2
            objPozoAnalis.Bar3 = rs!Bar3
            objPozoAnalis.Bar4 = rs!Bar4
            objPozoAnalis.Bar5 = rs!Bar5
            objPozoAnalis.BH = rs!BH
            objPozoAnalis.C1 = rs!C1
            objPozoAnalis.C2 = rs!C2
            objPozoAnalis.C3 = rs!C3
            objPozoAnalis.C4 = rs!C4
            objPozoAnalis.C5 = rs!C5
            objPozoAnalis.CH = rs!CH
            objPozoAnalis.CO2 = rs!CO2
            objPozoAnalis.CodigoTipoDeAnalisis = rs!CodigoTipoDeAnalisis
            objPozoAnalis.Cous1 = rs!Cous1
            objPozoAnalis.Cous2 = rs!Cous2
            objPozoAnalis.EFast = rs!EFast
            objPozoAnalis.ENormal = rs!ENormal
            objPozoAnalis.fecha = rs!fecha
            objPozoAnalis.GasTotal = rs!GasTotal
            objPozoAnalis.GasTotalCromatografico = rs!GasTotalCromatografico
            objPozoAnalis.Geo1 = rs!Geo1
            objPozoAnalis.Geo2 = rs!Geo2
            objPozoAnalis.Geo3 = rs!Geo3
            objPozoAnalis.Geo4 = rs!Geo4
            objPozoAnalis.IC4 = rs!IC4
            objPozoAnalis.IC5 = rs!IC5
            objPozoAnalis.IdPozo = rs!IdPozo
            objPozoAnalis.NC4 = rs!NC4
            objPozoAnalis.NC5 = rs!NC5
            objPozoAnalis.NumeroAnalisis = rs!NumeroAnalisis
            If Not IsNull(rs!Observaciones) Then
                objPozoAnalis.Observaciones = rs!Observaciones
            End If
            objPozoAnalis.Profundidad = rs!Profundidad
            objPozoAnalis.Seleccionado = rs!Seleccionado
            objPozoAnalis.SH2 = rs!SH2
            objPozoAnalis.SnGeo = rs!SnGeo
            objPozoAnalis.WH = rs!WH
            
            Me.colAgregar objPozoAnalis
        End If
                                 
        If Not IsNull(rs!NumeroComponente) Then
            Set objPozoAnalisComponente = objPozoAnalis.PozoAnalisComponentes.colComponente(rs!IdPozo, rs!NumeroAnalisis, rs!NumeroComponente)
            If objPozoAnalisComponente Is Nothing Then
                
                Set objPozoAnalisComponente = New ClsPozoAnalisComponente
                
                objPozoAnalisComponente.Area = rs!Area
                objPozoAnalisComponente.Component = rs!Component
                objPozoAnalisComponente.Externo = rs!Externo
                objPozoAnalisComponente.IdPozo = rs!IdPozo
                objPozoAnalisComponente.NormArea = rs!NormArea
                objPozoAnalisComponente.NumeroAnalisis = rs!NumeroAnalisis
                objPozoAnalisComponente.NumeroComponente = rs!NumeroComponente
                objPozoAnalisComponente.Retention = rs!Retention
                objPozoAnalisComponente.Units = rs!Units
                
                objPozoAnalis.PozoAnalisComponentes.colAgregar objPozoAnalisComponente
                
            End If
        End If
             
        rs.MoveNext
        
    Loop
    
Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsPozoAnalisis"
        pDatos.Error.Metodo = "ConsultarPozoAnalisis"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Function

Public Function dbAgregar(ByRef objPozoAnalis As clsPozoAnalis) As Boolean
    
    Dim StrSql                      As String
    
    On Error GoTo Error
    
    StrSql = "INSERT INTO "
    StrSql = StrSql & "iga_pozos_analisis ("
    StrSql = StrSql & "idpozo, "
    StrSql = StrSql & "numeroanalisis, "
    StrSql = StrSql & "archivo, "
    StrSql = StrSql & "fecha, "
    StrSql = StrSql & "profundidad, "
    StrSql = StrSql & "co2, "
    StrSql = StrSql & "sh2, "
    StrSql = StrSql & "gastotal, "
    StrSql = StrSql & "gastotalcromatografico, "
    StrSql = StrSql & "codigotipodeanalisis, "
    StrSql = StrSql & "efast, "
    StrSql = StrSql & "enormal, "
    StrSql = StrSql & "observaciones, "
    StrSql = StrSql & "seleccionado, "
    StrSql = StrSql & "bar2, "
    StrSql = StrSql & "bar3, "
    StrSql = StrSql & "bar4, "
    StrSql = StrSql & "bar5, "
    StrSql = StrSql & "cous1, "
    StrSql = StrSql & "cous2, "
    StrSql = StrSql & "wh, "
    StrSql = StrSql & "bh, "
    StrSql = StrSql & "ch, "
    StrSql = StrSql & "geo1, "
    StrSql = StrSql & "geo2,"
    StrSql = StrSql & "geo3, "
    StrSql = StrSql & "geo4, "
    StrSql = StrSql & "sngeo, "
    StrSql = StrSql & "c1, "
    StrSql = StrSql & "c2, "
    StrSql = StrSql & "c3, "
    StrSql = StrSql & "c4, "
    StrSql = StrSql & "c5, "
    StrSql = StrSql & "ic4, "
    StrSql = StrSql & "nc4, "
    StrSql = StrSql & "ic5, "
    StrSql = StrSql & "nc5 "
    StrSql = StrSql & ") VALUES ("
    StrSql = StrSql & objPozoAnalis.IdPozo & ", "
    StrSql = StrSql & objPozoAnalis.NumeroAnalisis & ", '"
    StrSql = StrSql & objPozoAnalis.archivo & "', "
    StrSql = StrSql & IIf(objPozoAnalis.fecha <> 0, pDatos.CaracterFecha & Format(objPozoAnalis.fecha, pDatos.FormatoFechaLarga) & pDatos.CaracterFecha, "NULL") & ", "
    StrSql = StrSql & objPozoAnalis.Profundidad & ", "
    StrSql = StrSql & objPozoAnalis.CO2 & ", "
    StrSql = StrSql & objPozoAnalis.SH2 & ", "
    StrSql = StrSql & objPozoAnalis.GasTotal & ", "
    StrSql = StrSql & objPozoAnalis.GasTotalCromatografico & ", "
    StrSql = StrSql & objPozoAnalis.CodigoTipoDeAnalisis & ", "
    StrSql = StrSql & objPozoAnalis.EFast & ", "
    StrSql = StrSql & objPozoAnalis.ENormal & ", '"
    StrSql = StrSql & objPozoAnalis.Observaciones & "', "
    StrSql = StrSql & IIf(objPozoAnalis.Seleccionado, "TRUE", "FALSE") & ", "
    StrSql = StrSql & objPozoAnalis.Bar2 & ", "
    StrSql = StrSql & objPozoAnalis.Bar3 & ", "
    StrSql = StrSql & objPozoAnalis.Bar4 & ", "
    StrSql = StrSql & objPozoAnalis.Bar5 & ", "
    StrSql = StrSql & objPozoAnalis.Cous1 & ", "
    StrSql = StrSql & objPozoAnalis.Cous2 & ", "
    StrSql = StrSql & objPozoAnalis.WH & ", "
    StrSql = StrSql & objPozoAnalis.BH & ", "
    StrSql = StrSql & objPozoAnalis.CH & ", "
    StrSql = StrSql & objPozoAnalis.Geo1 & ", "
    StrSql = StrSql & objPozoAnalis.Geo2 & ", "
    StrSql = StrSql & objPozoAnalis.Geo3 & ", "
    StrSql = StrSql & objPozoAnalis.Geo4 & ", "
    StrSql = StrSql & objPozoAnalis.SnGeo & ", "
    StrSql = StrSql & objPozoAnalis.C1 & ", "
    StrSql = StrSql & objPozoAnalis.C2 & ", "
    StrSql = StrSql & objPozoAnalis.C3 & ", "
    StrSql = StrSql & objPozoAnalis.C4 & ", "
    StrSql = StrSql & objPozoAnalis.C5 & ", "
    StrSql = StrSql & objPozoAnalis.IC4 & ", "
    StrSql = StrSql & objPozoAnalis.NC4 & ", "
    StrSql = StrSql & objPozoAnalis.IC5 & ", "
    StrSql = StrSql & objPozoAnalis.NC5 & " "
    StrSql = StrSql & ");"

    dbAgregar = pDatos.DbEjecutar(StrSql)
    
Error:
    
    If Err.Number <> 0 Then
        dbAgregar = False
        pDatos.Error.Objeto = "clsPozoAnalisis"
        pDatos.Error.Metodo = "dbAgregar"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Function

Public Function dbModificar(ByVal objPozoAnalis As clsPozoAnalis) As Boolean
    
    Dim StrSql                      As String
    
    On Error GoTo Error
    StrSql = "Update iga_pozos_Analisis Set "
    StrSql = StrSql & " archivo = '" & objPozoAnalis.archivo & "', "
    StrSql = StrSql & " Fecha = " & IIf(objPozoAnalis.fecha <> 0, pDatos.CaracterFecha & Format(objPozoAnalis.fecha, pDatos.FormatoFechaLarga) & pDatos.CaracterFecha, Null) & ", "
    StrSql = StrSql & " Profundidad = " & objPozoAnalis.Profundidad & ", "
    StrSql = StrSql & " CO2 = " & objPozoAnalis.CO2 & ", "
    StrSql = StrSql & " SH2 = " & objPozoAnalis.SH2 & ", "
    StrSql = StrSql & " GasTotal = " & objPozoAnalis.GasTotal & ", "
    StrSql = StrSql & " GasTotalCromatografico = " & objPozoAnalis.GasTotalCromatografico & ", "
    StrSql = StrSql & " CodigoTipoDeAnalisis = " & objPozoAnalis.CodigoTipoDeAnalisis & ", "
    StrSql = StrSql & " EFast = " & objPozoAnalis.EFast & ", "
    StrSql = StrSql & " ENormal = " & objPozoAnalis.ENormal & ", "
    StrSql = StrSql & " Observaciones = '" & objPozoAnalis.Observaciones & "', "
    StrSql = StrSql & " Seleccionado = " & IIf(objPozoAnalis.Seleccionado, "TRUE", "FALSE") & ", "
    StrSql = StrSql & " Bar2 = " & objPozoAnalis.Bar2 & ", "
    StrSql = StrSql & " Bar3 = " & objPozoAnalis.Bar3 & ", "
    StrSql = StrSql & " Bar4 = " & objPozoAnalis.Bar4 & ", "
    StrSql = StrSql & " Bar5 = " & objPozoAnalis.Bar5 & ", "
    StrSql = StrSql & " Cous1 = " & objPozoAnalis.Cous1 & ", "
    StrSql = StrSql & " Cous2 = " & objPozoAnalis.Cous2 & ", "
    StrSql = StrSql & " WH = " & objPozoAnalis.WH & ", "
    StrSql = StrSql & " BH = " & objPozoAnalis.BH & ", "
    StrSql = StrSql & " CH = " & objPozoAnalis.CH & ", "
    StrSql = StrSql & " Geo1 = " & objPozoAnalis.Geo1 & ", "
    StrSql = StrSql & " Geo2 = " & objPozoAnalis.Geo2 & ", "
    StrSql = StrSql & " Geo3 = " & objPozoAnalis.Geo3 & ", "
    StrSql = StrSql & " Geo4 = " & objPozoAnalis.Geo4 & ", "
    StrSql = StrSql & " SnGeo = " & objPozoAnalis.SnGeo & ", "
    StrSql = StrSql & " C1 = " & objPozoAnalis.C1 & ", "
    StrSql = StrSql & " C2 = " & objPozoAnalis.C2 & ", "
    StrSql = StrSql & " C3 = " & objPozoAnalis.C3 & ", "
    StrSql = StrSql & " C4 = " & objPozoAnalis.C4 & ", "
    StrSql = StrSql & " C5 = " & objPozoAnalis.C5 & ", "
    StrSql = StrSql & " IC4 = " & objPozoAnalis.IC4 & ", "
    StrSql = StrSql & " NC4 = " & objPozoAnalis.NC4 & ", "
    StrSql = StrSql & " IC5 = " & objPozoAnalis.IC5 & ", "
    StrSql = StrSql & " NC5 = " & objPozoAnalis.NC5 & " "
    StrSql = StrSql & "WHERE "
    StrSql = StrSql & "idPozo = " & objPozoAnalis.IdPozo & " AND "
    StrSql = StrSql & "NumeroAnalisis = " & objPozoAnalis.NumeroAnalisis & ";"
    
    dbModificar = pDatos.DbEjecutar(StrSql)
    
Error:
    
    If Err.Number <> 0 Then
        dbModificar = False
        pDatos.Error.Objeto = "clsPozoAnalisis"
        pDatos.Error.Metodo = "dbModificar"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Function



Public Function dbEliminar(ByVal objPozoAnalis As clsPozoAnalis) As Boolean
    
    Dim StrSql                      As String
    
    On Error GoTo Error
    StrSql = "Delete " & pDatos.CaracterTodosLosCampos & " from iga_pozos_analisis "
    StrSql = StrSql & "WHERE "
    StrSql = StrSql & "idPozo = " & objPozoAnalis.IdPozo & " AND "
    StrSql = StrSql & "NumeroAnalisis = " & objPozoAnalis.NumeroAnalisis & ";"

    dbEliminar = pDatos.DbEjecutar(StrSql)
    
Error:
    
    If Err.Number <> 0 Then
        dbEliminar = False
        pDatos.Error.Objeto = "clsPozoAnalisis"
        pDatos.Error.Metodo = "dbEliminar"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Function


Public Property Get dbMax(ByVal IdPozo As Long) As Long
    
    Dim StrSql          As String
    Dim rs              As Recordset
    On Error GoTo Error
    
    StrSql = "Select "
    StrSql = StrSql & "max(NumeroAnalisis) as maximo "
    StrSql = StrSql & "from "
    StrSql = StrSql & "iga_pozos_analisis "
    StrSql = StrSql & "WHERE "
    StrSql = StrSql & "idpozo = " & IdPozo & ";"

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
        pDatos.Error.Objeto = "clsPozoAnalisis"
        pDatos.Error.Metodo = "dbMax"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Property


Public Property Get colMax() As Long
    
    Dim max As Long
    Dim objPozoAnalis As clsPozoAnalis
    
    On Error GoTo Error
    max = 0
    For Each objPozoAnalis In pColPozoAnalisis
        If objPozoAnalis.NumeroAnalisis > max Then
            max = objPozoAnalis.NumeroAnalisis
        End If
    Next
    
    colMax = max
    
Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsPozoAnalisis"
        pDatos.Error.Metodo = "colMax"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Property

