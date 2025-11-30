VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private pConeccionDB                        As adodb.Connection
Private pRegistrosAfectados                 As Long
Private pError                              As New clsError
Private pCaracterFecha                      As String
Private pCaracterTodosLosCampos             As String
Private pFormatoFechaCorta                  As String
Private pFormatoFechaLarga                  As String
Private pPcDataCenter                       As String
Private pDSN                                As String
Private pGasTotalDesdeIga                   As Boolean
Private pCO2DesdeIga                        As Boolean
Private pSH2DesdeIga                        As Boolean
Private pMensajesWebHabilitados             As Boolean
Private pTimerMensajesWeb                   As Long
Private pMensajesLocalHabilitados           As Boolean
Private pTimerMensajesLocal                 As Long
Private pMensajesWanLanNotasHabilitados     As Boolean
Private pTimerEnvioNotasWanLan              As Long

Public Property Get Error() As clsError
    Set Error = pError
End Property

Public Property Get RegistrosAfectados() As Long
    RegistrosAfectados = pRegistrosAfectados
End Property
Public Sub Conectar()

On Error GoTo Error
    
    pConeccionDB.Open Me.DSN
    If Me.Conectada Then
        ConsultarParametrosGenerales
    End If
        
Error:
    If Err.Number <> 0 Then
        Me.Error.Objeto = "clsDb"
        Me.Error.Metodo = "Conectar"
        Me.Error.Number = Err.Number
        Me.Error.Description = Err.Description
        Me.Error.Mostrar
        Err.Clear
    End If

End Sub

Private Function ConsultarParametrosGenerales() As Boolean

    Dim StrSql                          As String
    Dim rs                              As New adodb.Recordset
    
    On Error GoTo Error
    
    StrSql = "SELECT "
    StrSql = StrSql & "IdSetupApp, "
    StrSql = StrSql & "CaracterFecha, "
    StrSql = StrSql & "CaracterTodosLosCampos, "
    StrSql = StrSql & "FormatoFechaCorta, "
    StrSql = StrSql & "FormatoFechaLarga, "
    StrSql = StrSql & "PCDataCenter, "
    StrSql = StrSql & "GasTotalDesdeIga, "
    StrSql = StrSql & "CO2DesdeIga, "
    StrSql = StrSql & "SH2DesdeIga, "
    StrSql = StrSql & "MensajesLocalHabilitados, "
    StrSql = StrSql & "MensajesWebHabilitados, "
    StrSql = StrSql & "TimerMensajesLocal, "
    StrSql = StrSql & "TimerMensajesWeb, "
    StrSql = StrSql & "MensajesWanLanNotasHabilitados , "
    StrSql = StrSql & "TimerEnvioNotasWanLan  "
    StrSql = StrSql & "FROM "
    StrSql = StrSql & "SetupApp "
    StrSql = StrSql & "WHERE "
    StrSql = StrSql & "IdSetupApp = 1;"

    Set rs = EjecutarSeleccion(StrSql)
    
    If Not rs Is Nothing Then
        If Not (rs.BOF And rs.EOF) Then
            
            pCaracterFecha = rs!CaracterFecha
            If Not IsNull(rs!CaracterTodosLosCampos) Then
                pCaracterTodosLosCampos = rs!CaracterTodosLosCampos
            End If
            pFormatoFechaCorta = rs!FormatoFechaCorta
            pFormatoFechaLarga = rs!FormatoFechaLarga
            pPcDataCenter = rs!PcDataCenter
            
            pCO2DesdeIga = rs!CO2DesdeIga
            pGasTotalDesdeIga = rs!GasTotalDesdeIga
            pSH2DesdeIga = rs!SH2DesdeIga
            
            pMensajesLocalHabilitados = rs!MensajesLocalHabilitados
            pMensajesWebHabilitados = rs!MensajesWebHabilitados
            pMensajesWanLanNotasHabilitados = rs!MensajesWanLanNotasHabilitados
            If Not IsNull(rs!TimerMensajesLocal) Then
                pTimerMensajesLocal = rs!TimerMensajesLocal
            End If
            If Not IsNull(rs!TimerMensajesWeb) Then
                pTimerMensajesWeb = rs!TimerMensajesWeb
            End If
            
            If Not IsNull(rs!TimerEnvioNotasWanLan) Then
                pTimerEnvioNotasWanLan = rs!TimerEnvioNotasWanLan
            End If
            
            ConsultarParametrosGenerales = True
            
        End If
    End If
    
Error:
    
    If Err.Number <> 0 Then
        ConsultarParametrosGenerales = False
        pError.Objeto = "clsdb"
        pError.Metodo = "ConsultarParametrosGenerales"
        pError.Number = Err.Number
        pError.Description = Err.Description
        pError.Mostrar
        Err.Clear
    End If
    
End Function

Public Function EjecutarSeleccion(ByVal StrSql As String) As Recordset

    On Error GoTo Error
            
    If Me.Conectada Then
        Set EjecutarSeleccion = pConeccionDB.Execute(StrSql)
    Else
        Reconectar
    End If
    
Error:
    If Err.Number <> 0 Then
        Me.Error.Objeto = "clsDb"
        Me.Error.Metodo = "EjecutarSeleccion"
        Me.Error.Number = Err.Number
        Me.Error.Description = Err.Description & "Cadena SQL = " & StrSql
        Me.Error.Mostrar
        Err.Clear
    End If

End Function

Public Function DbEjecutar(ByVal StrSql As String) As Boolean

On Error GoTo Error
            
    If Me.Conectada Then
        pConeccionDB.Execute StrSql, pRegistrosAfectados
        DbEjecutar = pRegistrosAfectados >= 1
    Else
        Reconectar
    End If
    
Error:
    If Err.Number <> 0 Then
        DbEjecutar = False
        pRegistrosAfectados = 0
        Me.Error.Objeto = "clsDb"
        Me.Error.Metodo = "DbEjecutar"
        Me.Error.Number = Err.Number
        Me.Error.Description = Err.Description & " - sql: " & StrSql
        Me.Error.Mostrar
        Err.Clear
    End If

End Function

Public Sub CommitTrans()

On Error GoTo Error
        
    If Me.Conectada Then
        pConeccionDB.CommitTrans
    Else
        Reconectar
    End If
    
    
Error:
    If Err.Number <> 0 Then
        Me.Error.Objeto = "clsDb"
        Me.Error.Metodo = "CommitTrans"
        Me.Error.Number = Err.Number
        Me.Error.Description = Err.Description
        Me.Error.Mostrar
        Err.Clear
    End If

End Sub

Public Sub RollBackTrans()

On Error GoTo Error
    
    If Me.Conectada Then
        pConeccionDB.RollBackTrans
    Else
        Reconectar
    End If
    
Error:
    If Err.Number <> 0 Then
        Me.Error.Objeto = "clsDb"
        Me.Error.Metodo = "RollBackTrans"
        Me.Error.Number = Err.Number
        Me.Error.Description = Err.Description
        Me.Error.Mostrar
        Err.Clear
    End If

End Sub

Public Sub BeginTrans()

On Error GoTo Error

    If Me.Conectada Then
        pConeccionDB.BeginTrans
    Else
        Reconectar
    End If

Error:
    If Err.Number <> 0 Then
        Me.Error.Objeto = "clsDb"
        Me.Error.Metodo = "BeginTrans"
        Me.Error.Number = Err.Number
        Me.Error.Description = Err.Description
        Me.Error.Mostrar
        Err.Clear
    End If

End Sub

Public Function Reconectar() As Boolean

On Error GoTo Error

    If pConeccionDB.State = 1 Then
        Reconectar = True
        Exit Function
    Else
        Set pConeccionDB = Nothing
        Set pConeccionDB = New adodb.Connection
    End If
    pConeccionDB.Open Me.DSN
    
    Reconectar = Me.Conectada
    
Error:
    If Err.Number <> 0 Then
        pError.Description = Err.Description & " - base de datos: "
        pError.Objeto = "clsdb"
        pError.Metodo = "Reconectar"
        pError.Number = Err.Number
        pError.Mostrar
        Err.Clear
    End If

End Function
Public Property Get CaracterFecha() As String
    CaracterFecha = pCaracterFecha
End Property

Public Property Let CaracterFecha(ByVal vNewValue As String)
    pCaracterFecha = vNewValue
End Property

Public Property Get CaracterTodosLosCampos() As String
    CaracterTodosLosCampos = pCaracterTodosLosCampos
End Property

Public Property Let CaracterTodosLosCampos(ByVal vNewValue As String)
    pCaracterTodosLosCampos = vNewValue
End Property

Public Property Get Conectada() As Boolean
    Conectada = pConeccionDB.State = adStateOpen
End Property

Private Sub Class_Initialize()
    Set pConeccionDB = New adodb.Connection
End Sub

Public Property Get FormatoFechaCorta() As String
    FormatoFechaCorta = pFormatoFechaCorta
End Property

Public Property Let FormatoFechaCorta(ByVal vNewValue As String)
    pFormatoFechaCorta = vNewValue
End Property

Public Property Get FormatoFechaLarga() As String
    FormatoFechaLarga = pFormatoFechaLarga
End Property

Public Property Let FormatoFechaLarga(ByVal vNewValue As String)
    pFormatoFechaLarga = vNewValue
End Property

Public Property Get PcDataCenter() As String
    PcDataCenter = pPcDataCenter
End Property

Public Property Let PcDataCenter(ByVal vNewValue As String)
    pPcDataCenter = vNewValue
End Property

Public Property Get DSN() As String
    DSN = pDSN
End Property

Public Property Let DSN(ByVal vNewValue As String)
    pDSN = vNewValue
End Property

Public Property Get GasTotalDesdeIga() As Boolean
    GasTotalDesdeIga = pGasTotalDesdeIga
End Property

Public Property Let GasTotalDesdeIga(ByVal vNewValue As Boolean)
    pGasTotalDesdeIga = vNewValue
End Property

Public Property Get CO2DesdeIga() As Boolean
    CO2DesdeIga = pCO2DesdeIga
End Property

Public Property Let CO2DesdeIga(ByVal vNewValue As Boolean)
    pCO2DesdeIga = vNewValue
End Property

Public Property Get SH2DesdeIga() As Boolean
    SH2DesdeIga = pSH2DesdeIga
End Property

Public Property Let SH2DesdeIga(ByVal vNewValue As Boolean)
    pSH2DesdeIga = vNewValue
End Property


Public Property Get MensajesWebHabilitados() As Boolean
    MensajesWebHabilitados = pMensajesWebHabilitados
End Property

Public Property Let MensajesWebHabilitados(ByVal vNewValue As Boolean)
    pMensajesWebHabilitados = vNewValue
End Property

Public Property Get TimerMensajesWeb() As Long
    TimerMensajesWeb = pTimerMensajesWeb
End Property

Public Property Let TimerMensajesWeb(ByVal vNewValue As Long)
    pTimerMensajesWeb = vNewValue
End Property

Public Property Get MensajesLocalHabilitados() As Boolean
    MensajesLocalHabilitados = pMensajesLocalHabilitados
End Property

Public Property Let MensajesLocalHabilitados(ByVal vNewValue As Boolean)
    pMensajesLocalHabilitados = vNewValue
End Property

Public Property Get TimerMensajesLocal() As Long
    TimerMensajesLocal = pTimerMensajesLocal
End Property

Public Property Let TimerMensajesLocal(ByVal vNewValue As Long)
    pTimerMensajesLocal = vNewValue
End Property

Public Property Get MensajesWanLanNotasHabilitados() As Boolean
    MensajesWanLanNotasHabilitados = pMensajesWanLanNotasHabilitados
End Property

Public Property Let MensajesWanLanNotasHabilitados(ByVal vNewValue As Boolean)
    pMensajesWanLanNotasHabilitados = vNewValue
End Property

Public Property Get TimerEnvioNotasWanLan() As Long
    TimerEnvioNotasWanLan = pTimerEnvioNotasWanLan
End Property

Public Property Let TimerEnvioNotasWanLan(ByVal vNewValue As Long)
    pTimerEnvioNotasWanLan = vNewValue
End Property

