VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPozoAnalisComponentes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Public pColComponentes  As New Collection
Private pDatos          As clsDB

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = pColComponentes.[_NewEnum]
End Function
    
Public Property Set Datos(ByVal vNewValue As clsDB)
    Set pDatos = vNewValue
End Property

Public Property Get Count() As Long
    Count = pColComponentes.Count
End Property
    
Public Function colAgregar(ByVal objComponente As ClsPozoAnalisComponente) As Boolean
    On Error GoTo Error
    pColComponentes.Add objComponente, objComponente.Key
    colAgregar = True
    
Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsPozoAnalisComponentes"
        pDatos.Error.Metodo = "colAgregar"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If


End Function

Public Function colQuitar(ByVal objComponente As ClsPozoAnalisComponente) As Boolean

On Error GoTo Error
    
    pColComponentes.Remove objComponente.Key
    colQuitar = True
    
Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsPozoAnalisComponentes"
        pDatos.Error.Metodo = "colQuitar"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If



End Function
                                                
Public Property Get colComponente(Optional ByVal IdPozo As Long, _
                                  Optional ByVal NumeroAnalisis As Long, _
                                  Optional ByVal NumeroComponente As Long, _
                                  Optional ByVal Key As Variant) As ClsPozoAnalisComponente
    On Error GoTo Error
    If IsMissing(Key) Then
        Key = IdPozo & NumeroAnalisis & NumeroComponente & "ID"
    End If
    Set colComponente = pColComponentes(Key)
    
Error:
    
    If Err.Number <> 0 Then
        If Err.Number <> 5 Then
            pDatos.Error.Objeto = "clsPozoAnalisComponentes"
            pDatos.Error.Metodo = "colComponente"
            pDatos.Error.Number = Err.Number
            pDatos.Error.Description = Err.Description
            pDatos.Error.Mostrar
        End If
        Err.Clear
    End If


End Property

Public Sub SetearNormArea(ByRef objComponente As ClsPozoAnalisComponente, ByVal SumaArea As Double)
    
On Error GoTo Error

    If SumaArea <> 0 Then
        objComponente.NormArea = (objComponente.Externo / SumaArea) * 100
    End If

Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsPozoAnalisComponentes"
        pDatos.Error.Metodo = "SetearNormArea"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
End Sub

Public Property Get getSumaArea() As Double
    
    Dim SumaArea            As Double
    Dim objComponenteI      As ClsPozoAnalisComponente
    
On Error GoTo Error

    For Each objComponenteI In pColComponentes
        SumaArea = SumaArea + objComponenteI.Externo
    Next
    
    getSumaArea = SumaArea
    
Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsPozoAnalisComponentes"
        pDatos.Error.Metodo = "SetearNormArea"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
End Property
Public Property Get colComponentesPorPozoYAnalisis(ByVal IdPozo As Long, ByVal NumeroAnalisis As Long) As clsPozoAnalisComponentes
    
    Dim objComponentes As New clsPozoAnalisComponentes
    Dim objComponente As ClsPozoAnalisComponente
    
    On Error GoTo Error
    
    For Each objComponente In pColComponentes
        If objComponente.IdPozo = IdPozo And objComponente.NumeroAnalisis = NumeroAnalisis Then
            objComponentes.colAgregar objComponente
        End If
    Next
    Set colComponentesPorPozoYAnalisis = objComponentes
Error:
    
    If Err.Number <> 0 Then
        pDatos.Error.Objeto = "clsPozoAnalisComponentes"
        pDatos.Error.Metodo = "colComponentesPorPozoYAnalisis"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If

End Property

Public Function dbAgregar(ByRef objPozoAnalisComponente As ClsPozoAnalisComponente) As Boolean
    
    Dim StrSql                      As String
    
    On Error GoTo Error
    
    StrSql = "Insert into iga_pozos_analisis_componentes ("
    StrSql = StrSql & "idpozo, "
    StrSql = StrSql & "numeroanalisis, "
    StrSql = StrSql & "numerocomponente, "
    StrSql = StrSql & "component, "
    StrSql = StrSql & "retention, "
    StrSql = StrSql & "area, "
    StrSql = StrSql & "externo, "
    StrSql = StrSql & "units, "
    StrSql = StrSql & "normarea "
    StrSql = StrSql & ") Values ("
    StrSql = StrSql & objPozoAnalisComponente.IdPozo & ", "
    StrSql = StrSql & objPozoAnalisComponente.NumeroAnalisis & ", "
    StrSql = StrSql & objPozoAnalisComponente.NumeroComponente & ", '"
    StrSql = StrSql & objPozoAnalisComponente.Component & "', "
    StrSql = StrSql & objPozoAnalisComponente.Retention & ", "
    StrSql = StrSql & objPozoAnalisComponente.Area & ", "
    StrSql = StrSql & objPozoAnalisComponente.Externo & ", '"
    StrSql = StrSql & objPozoAnalisComponente.Units & "', "
    StrSql = StrSql & objPozoAnalisComponente.NormArea & " "
    StrSql = StrSql & ");"

    dbAgregar = pDatos.DbEjecutar(StrSql)
    
Error:
    
    If Err.Number <> 0 Then
        dbAgregar = False
        pDatos.Error.Objeto = "clsPozoAnalisComponentes"
        pDatos.Error.Metodo = "dbAgregar"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Function

Public Function dbModificar(ByVal objPozoAnalisComponente As ClsPozoAnalisComponente) As Boolean
    
    Dim StrSql                      As String
    
    On Error GoTo Error
    StrSql = "Update iga_pozos_analisis_componentes set "
    StrSql = StrSql & "component = '" & objPozoAnalisComponente.Component & "', "
    StrSql = StrSql & "retention = " & objPozoAnalisComponente.Retention & ", "
    StrSql = StrSql & "area = " & objPozoAnalisComponente.Area & ", "
    StrSql = StrSql & "externo = " & objPozoAnalisComponente.Externo & ", "
    StrSql = StrSql & "units = '" & objPozoAnalisComponente.Units & "', "
    StrSql = StrSql & "normarea = " & objPozoAnalisComponente.NormArea & " "
    StrSql = StrSql & "WHERE "
    StrSql = StrSql & "idpozo = " & objPozoAnalisComponente.IdPozo & " "
    StrSql = StrSql & "AND "
    StrSql = StrSql & "numeroanalisis = " & objPozoAnalisComponente.NumeroAnalisis & " "
    StrSql = StrSql & "AND "
    StrSql = StrSql & "numerocomponente =  " & objPozoAnalisComponente.NumeroComponente & ";"

    dbModificar = pDatos.DbEjecutar(StrSql)
    
Error:
    
    If Err.Number <> 0 Then
        dbModificar = False
        pDatos.Error.Objeto = "clsPozoAnalisComponentes"
        pDatos.Error.Metodo = "dbModificar"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Function


Public Function dbModificarComponente(ByVal objPozoAnalisComponente As ClsPozoAnalisComponente, ByVal ComponenteOld As Long) As Boolean
    
    Dim StrSql                      As String
    
    On Error GoTo Error
    StrSql = "Update iga_pozos_analisis_componentes set "
    StrSql = StrSql & "numerocomponente = " & objPozoAnalisComponente.NumeroComponente & " "
    StrSql = StrSql & "WHERE "
    StrSql = StrSql & "idpozo = " & objPozoAnalisComponente.IdPozo & " "
    StrSql = StrSql & "AND"
    StrSql = StrSql & "numeroanalisis = " & objPozoAnalisComponente.NumeroAnalisis & " "
    StrSql = StrSql & "AND"
    StrSql = StrSql & "numerocomponente =  " & ComponenteOld & ";"

    dbModificarComponente = pDatos.DbEjecutar(StrSql)
    
Error:
    
    If Err.Number <> 0 Then
        dbModificarComponente = False
        pDatos.Error.Objeto = "clsPozoAnalisComponentes"
        pDatos.Error.Metodo = "dbModificar"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Function



Public Function dbEliminar(ByVal objPozoAnalisComponente As ClsPozoAnalisComponente) As Boolean
    
    Dim StrSql                      As String
    
    On Error GoTo Error
    StrSql = "Delete " & pDatos.CaracterTodosLosCampos & " from iga_pozos_analisis_componentes "
    StrSql = StrSql & "WHERE "
    StrSql = StrSql & "idpozo = " & objPozoAnalisComponente.IdPozo & " "
    StrSql = StrSql & "AND "
    StrSql = StrSql & "numeroanalisis = " & objPozoAnalisComponente.NumeroAnalisis & " "
    StrSql = StrSql & "AND "
    StrSql = StrSql & "numerocomponente =  " & objPozoAnalisComponente.NumeroComponente & ";"
    
    dbEliminar = pDatos.DbEjecutar(StrSql)
    
Error:
    
    If Err.Number <> 0 Then
        dbEliminar = False
        pDatos.Error.Objeto = "clsPozoAnalisComponentes"
        pDatos.Error.Metodo = "dbEliminar"
        pDatos.Error.Number = Err.Number
        pDatos.Error.Description = Err.Description
        pDatos.Error.Mostrar
        Err.Clear
    End If
    
End Function

