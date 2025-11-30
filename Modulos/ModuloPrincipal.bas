Attribute VB_Name = "ModuloPrincipal"
Option Explicit

Public gRefrescarGrafico                                                As Boolean
Public gIndiceEscala                                                    As Long
Public gIndiceEscalaLog                                                 As Long
Public gCaminoArchivos                                                  As String

Public gDatos                                                           As New clsDB

Public ProximaProfundidad                                               As Double 'La profundidad de la proxima muestra en viaje
Public gIdPozoEscaneoProximo                                            As Long
Public ProfundidadAnalisis                                              As Double 'La profundidad de la muestra que se está analizando
Public gMetroCortadoActual                                              As Double

Public LlegoMuestra As Boolean 'Es verdadero cuando desaparece de la tabla MuestrasViajeCutting la profundidad = a ProximaProfundidad

Public gObjPozoActivo                                                   As clsPozo
Public gObjComponentes                                                  As New clsComponentes
Public gObjFactores                                                     As New clsFactores
Public gObjTiposDeAnalisis                                              As New clsTiposDeAnalisis
Public gObjAdminEventos                                                 As New clsAdminEventos

Public gObjTiposEquipoGas                                               As New clsTiposEquipoGas
Public gObjTiposEquipoCroma                                             As New clsTiposEquipoCroma

Public StartRun                                                         As Double
Public BackFlushAnalizing                                               As Double
Public EContinuoLocal                                                   As Double
Public ENormalLocal                                                     As Double
Public EFastLocal                                                       As Double
Public EContinuoControlador                                             As Double
Public ENormalControlador                                               As Double
Public EFastControlador                                                 As Double
Public PosFast                                                          As Double
Public PosNormal                                                        As Double
Public TempFID                                                          As Double
Public TempOVEN                                                         As Double
Public SamplePressure                                                   As Double
Public Signal                                                           As Double
Public SH2                                                              As Double
Public CO2                                                              As Double
Public GasTotal                                                         As Double
Public AlarmaVacio                                                      As Double
Public ProfundidadRetornoAnterior                                       As Double

Public EscalaAnteriorTHA                                                As Integer
Public HoraInicio                                                       As Double
Public CromaLanzada                                                     As Boolean
Public DeboTirarCroma                                                   As Boolean
Public DeboAgregarCartel                                                As Boolean
Public PosicionCartelProfundidad                                        As Long


Public Uno                                                              As Long
Public Diez                                                             As Long
Public Cien                                                             As Long


Public PuedoCerrarForm                                                  As Boolean

Public ComentarioGas                                                    As String
Public EstoyCargando                                                    As Boolean
Public gSpan                                                            As Double
Public gSpanYCrono                                                      As Double
Public gZonaMuerta                                                      As Double

Public gObjConfiguracion                                                As clsConfiguracion

''''**************************Variables Globales***********************
Public strData                                                          As Variant
Public strTemp                                                          As String
Public strlook                                                          As String
Public i                                                                As Integer
Public Pos                                                              As Integer
Public Valor                                                            As Long
Public Ticks                                                            As Integer
Public TicksAnterior                                                    As Integer
Public CommOK                                                           As Boolean
Public Hay                                                              As Boolean
Public CromaCorriendo                                                   As Boolean

''''**************************Modulo de Sonido*************************

Public Const SND_ASYNC = &H1     'modo asíncrono. La función retorna una vez iniciada la música (sonido en background).
Public Const SND_LOOP = &H8      'La música seguirá sonando repetidamente hasta
                                  'que la función sndPlaySound sea llamada de nuevo con un valor nulo para NombreWav (NULL).

Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" _
(ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

''''**************************Función de asignación********************
Public ProfundidadRetorno As Double
Public Estado As String
Public EstadoAnterior As String

'******************************Acciones de formularios************************
Public Enum euAccion
    euAccionAgregar = 0
    euAccionModificar = 1
    euAccionEliminar = 2
    euAccionSeleccionar = 3
    euAccionVer = 4
End Enum

Private Sub Main()
    
    Dim objPozos            As New clsPozos
    
    On Error GoTo Error
    
    gDatos.DSN = "dgsdbpost"
    gDatos.Error.Destination = "MSG"
    gDatos.Conectar
    
    gCaminoArchivos = GetSetting("Iga", "Caminos", "Archivos")

    gIndiceEscala = Val(GetSetting("Iga", "Config", "Escala"))

    If gIndiceEscala = 0 Then

        gIndiceEscala = 20

    End If

    gIndiceEscalaLog = Val(GetSetting("Iga", "Config", "EscalaLog"))

    If gIndiceEscalaLog = 0 Then

        gIndiceEscalaLog = 100

    End If

    

    gSpan = Val(GetSetting("Iga", "Config", "SpanTemporal"))

    If gSpan = 0 Then

        gSpan = 120

    End If

    If GetSetting("Iga", "Config", "SpanCrono") <> "" Then

        gSpanYCrono = GetSetting("Iga", "Config", "SpanCrono")

    Else

        gSpanYCrono = 2

    End If

    

    gZonaMuerta = Val(GetSetting("Iga", "Config", "ZonaMuerta"))
    
    If gDatos.Conectada Then
        If ConsultarConfiguracion() Then
            Set gObjComponentes.Datos = gDatos
            gObjComponentes.ConsultarComponentes
            
            Set gObjFactores.Datos = gDatos
            gObjFactores.ConsultarFactores
            
            Set gObjTiposDeAnalisis.Datos = gDatos
            gObjTiposDeAnalisis.ConsultarTiposDeAnalisiss
            
            Set gObjTiposEquipoCroma.Datos = gDatos
            gObjTiposEquipoCroma.ConsultarTiposEquipoCroma
            
            Set gObjTiposEquipoGas.Datos = gDatos
            gObjTiposEquipoGas.ConsultarTiposEquipoGas
            
            If gObjConfiguracion.ArchivoDeResultados = "" Then
                frmMsg.MostrarMsg "No se configuró el nombre del archivo de resultados. No se podrán cargar análisis", "Error", MdiPrincipal
            End If
            
            If gObjConfiguracion.ArchivoDeResultadosGenerales = "" Then
                frmMsg.MostrarMsg "No se configuró el nombre del archivo de resultados generales. No se podrán mantener un histórico de archivos de resultados", "Error", MdiPrincipal
            End If

            Set objPozos.Datos = gDatos
            Set gObjPozoActivo = objPozos.dbPozo(True, 1)
            If gObjPozoActivo Is Nothing Then
                frmMsg.MostrarMsg "No hay ningún pozo activo. No se podrán cargar análisis", "Error", MdiPrincipal
            End If
            
            frmSplash.Show vbModal
            
            MdiPrincipal.Show
        Else
            
            frmMsg.MostrarMsg "No se pudo obtener las rs de configuración del sistema", "Error", MdiPrincipal
            
        End If
    Else
        frmMsg.MostrarMsg "No es posible conectarse con la base de datos Iga", "Error", MdiPrincipal
    End If
    
Error:
    If Err.Number <> 0 Then
        frmMsg.MostrarMsg "Módulo: Main." & Chr(10) & Err.Number & " - " & Err.Description, "Error", MdiPrincipal
        Err.Clear
    End If
End Sub

Private Function ConsultarConfiguracion() As Boolean
    
    Dim StrSql                      As String
    Dim rs                          As Recordset
    Dim objPozos                    As New clsPozos

    On Error GoTo Error
    
    ConsultarConfiguracion = False
    
    StrSql = ""
    StrSql = StrSql & "SELECT "
    StrSql = StrSql & " ArchivoDeResultados, "
    StrSql = StrSql & " ArchivoDeResultadosGenerales, "
    StrSql = StrSql & " AnalisisAutomaticos, "
    StrSql = StrSql & " CambioEscalas, "
    StrSql = StrSql & " TomaProfundidad, "
    StrSql = StrSql & " DisparoCiclico, "
    StrSql = StrSql & " PuertoComm, "
    StrSql = StrSql & " TiempoAnalisis, "
    StrSql = StrSql & " NivelDisparoMetroMetro, "
    StrSql = StrSql & " IpRabit, "
    StrSql = StrSql & " IpDisparoCroma, "
    StrSql = StrSql & " IdTipoEquipoCroma, "
    StrSql = StrSql & " IdTipoEquipoGas "
    StrSql = StrSql & "FROM "
    StrSql = StrSql & " iga_configuracion "

    Set rs = gDatos.EjecutarSeleccion(StrSql)

    If Not rs.EOF Then

        Set gObjConfiguracion = New clsConfiguracion

        gObjConfiguracion.NivelDisparoMetroMetro = rs!NivelDisparoMetroMetro
        gObjConfiguracion.AnalisisAutomaticos = rs!AnalisisAutomaticos
        gObjConfiguracion.CambioEscalas = rs!CambioEscalas
        gObjConfiguracion.TomaProfundidad = rs!TomaProfundidad
        gObjConfiguracion.DisparoCiclico = rs!DisparoCiclico
        gObjConfiguracion.PuertoComm = rs!PuertoComm
        gObjConfiguracion.TiempoAnalisis = rs!TiempoAnalisis

        
        If Not IsNull(rs!ArchivoDeResultados) Then
            gObjConfiguracion.ArchivoDeResultados = rs!ArchivoDeResultados
        End If

        If Not IsNull(rs!ArchivoDeResultadosGenerales) Then
            gObjConfiguracion.ArchivoDeResultadosGenerales = rs!ArchivoDeResultadosGenerales
        End If
        gObjConfiguracion.AnalisisAutomaticos = rs!AnalisisAutomaticos
        
        If Not IsNull(rs!IpRabit) Then
            gObjConfiguracion.IpRabit = rs!IpRabit
        End If
        
        If Not IsNull(rs!IpDisparoCroma) Then
            gObjConfiguracion.IpDisparoCroma = rs!IpDisparoCroma
        End If
        
        gObjConfiguracion.IdTipoEquipoCroma = rs!IdTipoEquipoCroma
        gObjConfiguracion.IdTipoEquipoGas = rs!IdTipoEquipoGas
        
        
    End If

    '*************Aquisición de gases****************
    
    StrSql = ""
    StrSql = StrSql & "SELECT "
    StrSql = StrSql & " CO2Desdeiga, "
    StrSql = StrSql & " SH2Desdeiga, "
    StrSql = StrSql & " GasTotalDesdeiga "
    StrSql = StrSql & "FROM "
    StrSql = StrSql & " setupapp;"

    Set rs = gDatos.EjecutarSeleccion(StrSql)

    If Not rs.EOF Then

        gObjConfiguracion.CO2DesdeDataCenter = False
        gObjConfiguracion.SH2DesdeDataCenter = False
        gObjConfiguracion.GasTotalDesdeDataCenter = False

        If rs!CO2DesdeIga = "0" Then gObjConfiguracion.CO2DesdeDataCenter = True
        If rs!SH2DesdeIga = "0" Then gObjConfiguracion.SH2DesdeDataCenter = True
        If rs!GasTotalDesdeIga = "0" Then gObjConfiguracion.GasTotalDesdeDataCenter = True
        
        ConsultarConfiguracion = True
        
    End If
Error:

    If Err.Number <> 0 Then

        ConsultarConfiguracion = False

        frmMsg.MostrarMsg "Módulo: ConsultarConfiguracion " & Chr(10) & Err.Number & " - " & Err.Description, "Error", MdiPrincipal

        Err.Clear

    End If
    
End Function
