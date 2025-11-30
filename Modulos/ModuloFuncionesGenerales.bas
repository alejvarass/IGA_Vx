Attribute VB_Name = "ModuloFuncionesGenerales"
Option Explicit

Private pDatos                      As clsDB

Public Type MedidaDeGas
    Pkno As String
    Name As String
    RetTime As String
    Conc As String
    Units As String
    Area As String
End Type

Public Array_tokens()                    As String
Public Array_tokensHijos()               As String
Public Array_Data()                      As MedidaDeGas
    

Public Sub CentrarFormulario(Formulario As Form)
    
    Formulario.Left = (Screen.Width - Formulario.Width) / 2
    Formulario.Top = (Screen.Height - Formulario.Height) / 2 - 500
    
End Sub

Public Sub leer_log(ok As Boolean, ArchivoSeleccionado As String)
    
    Dim objPozoAnalis                               As New clsPozoAnalis
    Dim objPozoAnalisComponente                     As ClsPozoAnalisComponente
    Dim objComponente                               As clsComponente
    
    
    Dim CantidadTabs                                As Integer
    Dim LineaTexto                                  As String
    Dim Valor                                       As String
    Dim CantidadCaracteres                          As Integer

    Dim i                                           As Integer
    Dim J                                           As Integer
    Dim Contador                                    As Integer
    Dim fecha                                       As Date
    Dim archivo                                     As String
    Dim NumeroComponente                            As Integer
    Dim Component                                   As String
    Dim Retention                                   As Double
    Dim Area                                        As Double
    Dim Externo                                     As Long
    Dim Units                                       As String
    Dim NormArea                                    As Double
    Dim StrSql                                      As String
    Dim NumeroAnalisis                              As Long
    Dim MayorOrden                                  As Recordset
    Dim rChequear                                   As Recordset
    Dim GasTotalCromatografico                      As Integer
    Dim Profundidad                                 As Double
    Dim caracter                                    As String
    Dim SumaArea                                    As Double
        
    On Error GoTo Error
    
    If Not gObjPozoActivo Is Nothing Then
    
        ok = True
        Contador = 1
        Set gObjPozoActivo.PozoAnalisis.Datos = gDatos
        NumeroAnalisis = gObjPozoActivo.PozoAnalisis.dbMax(gObjPozoActivo.IdPozo) + 1
        
        ArchivoSeleccionado = gCaminoArchivos & gObjConfiguracion.ArchivoDeResultados
        CantidadTabs = 0
        Open ArchivoSeleccionado For Input As #1   ' Abre el archivo.
        
        Do While Not EOF(1) ' Repite el bucle hasta el final del archivo.
            Line Input #1, LineaTexto    ' Lee el carácter en la variable.
            Valor = ""
            CantidadCaracteres = Len(LineaTexto)
            For i = Contador To CantidadCaracteres + 1
                
                If i > CantidadCaracteres Then
                    caracter = Chr(9)
                Else
                    caracter = Mid(LineaTexto, i, 1)
                End If
                
                If Asc(caracter) <> 9 Then
                    If Asc(caracter) <> 34 Then
                        Valor = Valor & caracter
                    End If
                Else
                    Select Case CantidadTabs
                        Case Is = 0
                            archivo = Valor
                        Case Is = 1
                            fecha = Format(InvertirFecha(Valor), "DD/MM/YYYY")
                        Case Is = 2
                            
                            fecha = Format(fecha + CDate(Valor), "DD/MM/YYYY HH:MM:SS")
                            Profundidad = ProfundidadAnalisis
                            
                            objPozoAnalis.archivo = archivo
                            objPozoAnalis.EFast = EFastLocal
                            objPozoAnalis.ENormal = ENormalLocal
                            objPozoAnalis.CodigoTipoDeAnalisis = 1
                            objPozoAnalis.NumeroAnalisis = NumeroAnalisis
                            objPozoAnalis.IdPozo = gObjPozoActivo.IdPozo
                            objPozoAnalis.fecha = fecha
                            objPozoAnalis.Profundidad = ProfundidadAnalisis
                            objPozoAnalis.CO2 = CO2
                            objPozoAnalis.SH2 = SH2
                            objPozoAnalis.GasTotal = GasTotal
                            objPozoAnalis.Seleccionado = False
                        
                        Case Is = 3
                            NumeroComponente = Valor
                        Case Is = 4
                            Component = Valor
                        Case Is = 5
                            Area = Valor
                        Case Is = 6
                            Retention = Valor
                        Case Is = 7
                            Set objComponente = gObjComponentes.colComponente(NumeroComponente)
                            If objComponente Is Nothing Then
                                Externo = Valor
                            ElseIf objComponente.CodigoFactor = 1 Then
                                Externo = Valor * EFastLocal
                            ElseIf objComponente.CodigoFactor = 2 Then
                                Externo = Valor * ENormalLocal
                            Else
                                frmMsg.MostrarMsg "No se encontró el factor para el " & Component & ". Se utilizara el valor real leido", "Error", MdiPrincipal
                                Externo = Valor
                            End If
                        Case Is = 8
                            Units = Valor
                        Case Is = 9
                            NormArea = Valor
                            Set objPozoAnalisComponente = objPozoAnalis.PozoAnalisComponentes.colComponente(objPozoAnalis.IdPozo, objPozoAnalis.NumeroAnalisis, NumeroComponente)
                            If objPozoAnalisComponente Is Nothing Then
                                Set objPozoAnalisComponente = New ClsPozoAnalisComponente
                                objPozoAnalisComponente.NumeroAnalisis = NumeroAnalisis
                                objPozoAnalisComponente.IdPozo = gObjPozoActivo.IdPozo
                                objPozoAnalisComponente.NumeroComponente = NumeroComponente
                                objPozoAnalisComponente.Component = Component
                                objPozoAnalisComponente.Retention = Retention
                                objPozoAnalisComponente.Area = Area
                                objPozoAnalisComponente.Externo = Externo
                                objPozoAnalisComponente.Units = Units
                                objPozoAnalisComponente.NormArea = NormArea
                                
                                objPozoAnalis.PozoAnalisComponentes.colAgregar objPozoAnalisComponente
                            End If
                            
                            NumeroComponente = 0
                            Component = ""
                            Retention = 0
                            Area = 0
                            Externo = 0
                            Units = ""
                            NormArea = 0
    
                            CantidadTabs = 2
    
                    End Select
    
                    CantidadTabs = CantidadTabs + 1
                    Valor = ""
    
                End If
    
            Next
            Contador = 1
            CantidadTabs = 0
    
        Loop
    
        Close #1    ' Cierra el archivo.
        If gObjConfiguracion.ArchivoDeResultadosGenerales <> "" Then
            If ok Then
                Open gCaminoArchivos & gObjConfiguracion.ArchivoDeResultadosGenerales For Append As #2    ' Abre el archivo.
                Write #2, LineaTexto
                Close #2
            End If
        Else
            frmMsg.MostrarMsg "No se configuró el nombre del archivo de resultados generales. No se podrán mantener un histórico de archivos de resultados", "Error", MdiPrincipal
        End If
        If ok Then
            ActualizarGasTotalCromatografico objPozoAnalis, ok
            CalcularRelacionesCromatograficas objPozoAnalis, "AGREGAR", ok
            
            Set gObjPozoActivo.PozoAnalisis.Datos = gDatos
            ok = gObjPozoActivo.PozoAnalisis.dbAgregar(objPozoAnalis)
            If ok Then
                Set objPozoAnalis.PozoAnalisComponentes.Datos = gDatos
                SumaArea = objPozoAnalis.PozoAnalisComponentes.getSumaArea
                
                For Each objPozoAnalisComponente In objPozoAnalis.PozoAnalisComponentes
                    
                    objPozoAnalis.PozoAnalisComponentes.SetearNormArea objPozoAnalisComponente, SumaArea
                    
                    If Not objPozoAnalis.PozoAnalisComponentes.dbAgregar(objPozoAnalisComponente) Then
                        ok = False
                        Exit For
                    End If
                Next
                
                If ok Then
                    gObjPozoActivo.PozoAnalisis.colAgregar objPozoAnalis
                    If FormularioCargado("FrmAnalisis") Then
                        AgregarDatosTemporalAlList objPozoAnalis
                    End If
                End If
            End If
        End If
    End If

Error:
    If Err.Number <> 0 Then
        If Err.Number <> 53 Then
            gDatos.Error.Description = Err.Description
            gDatos.Error.Metodo = "leer_log"
            gDatos.Error.Objeto = "Func_Generales"
            gDatos.Error.Number = Err.Number
            gDatos.Error.Mostrar
        End If
        Err.Clear
    End If
End Sub

Public Function FormularioCargado(NombreFormulario As String) As Boolean
    
    Dim i As Integer
    Dim Encontrado As Boolean
    
    i = 0
    Encontrado = False
    
    Do Until i = Forms.Count Or Encontrado
        
        If Forms(i).Name = NombreFormulario Then
            
            Encontrado = True
            
        End If
        
        i = i + 1
        
    Loop
    
    If Encontrado Then
        
        FormularioCargado = True
        
    Else
        
        FormularioCargado = False
        
    End If
    
End Function

Private Sub AgregarDatosTemporalAlList(ByVal objPozoAnalis As clsPozoAnalis)
    
    Dim Item As ListItem

    On Error GoTo Error

    Set Item = FrmAnalisis.LvwAnalisisTemporal.ListItems.Add(, objPozoAnalis.NumeroAnalisis & "ID")

    Item.Tag = objPozoAnalis.CodigoTipoDeAnalisis
    Item.Text = Format(objPozoAnalis.NumeroAnalisis, "00000")
    Item.SubItems(1) = objPozoAnalis.NumeroAnalisis
    Item.SubItems(2) = Format(objPozoAnalis.fecha, "DD/MM/YYYY HH:MM:SS")
    Item.SubItems(3) = objPozoAnalis.Profundidad

    Item.SmallIcon = 1

Error:
    If Err.Number <> 0 Then
        frmMsg.MostrarMsg "Módulo: AgregarDatosTemporalAlList." & Chr(10) & Err.Number & " - " & Err.Description, "Error", MdiPrincipal
        Err.Clear
    End If
End Sub

Public Sub TextSelected()
    'Para que seleccione el texto de un TextBox.
    
    Dim i As Integer
    Dim Objeto As Object
    
    Set Objeto = Screen.ActiveControl
    
    If TypeName(Objeto) = "TextBox" Then
        
        i = Len(Objeto.Text)
        
        Objeto.SelStart = 0
        Objeto.SelLength = i
        
    End If
    
End Sub

Public Sub ActualizarGasTotalCromatografico(ByVal objPozoAnalis As clsPozoAnalis, ok As Boolean)
    
    Dim objPozoAnalisComponente                 As ClsPozoAnalisComponente
    Dim GasTotalCromatografico                  As Long

    ok = True
    GasTotalCromatografico = 0
    
''''    If gObjConfiguracion.ThaFid Then
        For Each objPozoAnalisComponente In objPozoAnalis.PozoAnalisComponentes
            Select Case objPozoAnalisComponente.NumeroComponente
                Case Is = 1
                    GasTotalCromatografico = GasTotalCromatografico + objPozoAnalisComponente.Externo
                Case Is = 2
                    GasTotalCromatografico = GasTotalCromatografico + (2 * objPozoAnalisComponente.Externo)
                Case Is = 3
                    GasTotalCromatografico = GasTotalCromatografico + (3 * objPozoAnalisComponente.Externo)
                Case Is = 4
                    GasTotalCromatografico = GasTotalCromatografico + (4 * objPozoAnalisComponente.Externo)
                Case Is = 5
                    GasTotalCromatografico = GasTotalCromatografico + (4 * objPozoAnalisComponente.Externo)
                Case Is = 6
                    GasTotalCromatografico = GasTotalCromatografico + (5 * objPozoAnalisComponente.Externo)
                Case Is = 7
                    GasTotalCromatografico = GasTotalCromatografico + (5 * objPozoAnalisComponente.Externo)
            End Select

        Next
''''    Else
''''        For Each objPozoAnalisComponente In objPozoAnalis.PozoAnalisComponentes
''''            Select Case objPozoAnalisComponente.NumeroComponente
''''                Case Is = 1
''''                    GasTotalCromatografico = GasTotalCromatografico + objPozoAnalisComponente.Externo
''''                Case Is = 2
''''                    GasTotalCromatografico = GasTotalCromatografico + (6.3 * objPozoAnalisComponente.Externo)
''''                Case Is = 3
''''                    GasTotalCromatografico = GasTotalCromatografico + (8.3 * objPozoAnalisComponente.Externo)
''''                Case Is = 4
''''                    GasTotalCromatografico = GasTotalCromatografico + (12.5 * objPozoAnalisComponente.Externo)
''''                Case Is = 5
''''                    GasTotalCromatografico = GasTotalCromatografico + (12.5 * objPozoAnalisComponente.Externo)
''''                Case Is = 6
''''                    GasTotalCromatografico = GasTotalCromatografico + (8.3 * objPozoAnalisComponente.Externo)
''''                Case Is = 7
''''                    GasTotalCromatografico = GasTotalCromatografico + (8.3 * objPozoAnalisComponente.Externo)
''''            End Select
''''
''''        Next
''''    End If
    objPozoAnalis.GasTotalCromatografico = GasTotalCromatografico

End Sub

Public Sub CalcularRelacionesCromatograficas(ByVal objPozoAnalis As clsPozoAnalis, Accion As String, ok As Boolean)
    
    Dim objPozoAnalisComponente                 As ClsPozoAnalisComponente
    
    Dim Metano As Long
    Dim Etano As Long
    Dim Propano As Long
    Dim IsoButano As Long
    Dim NormalButano As Long
    Dim IsoPentano As Long
    Dim NormalPentano As Long
    Dim StrSql As String
    Dim Componentes As Recordset
    Dim GasTotalCromatografico As Long
    Dim Bar2 As Double
    Dim Bar3 As Double
    Dim Bar4 As Double
    Dim Bar5 As Double
    Dim Cous1 As Double
    Dim Cous2 As Double
    Dim WH As Double
    Dim BH As Double
    Dim CH As Double
    Dim SnGeo As Double
    Dim Geo1 As Double
    Dim Geo2 As Double
    Dim Geo3 As Double
    Dim Geo4 As Double
    Dim C1 As Double
    Dim C2 As Double
    Dim C3 As Double
    Dim C4 As Double
    Dim C5 As Double
    Dim IC4 As Double
    Dim NC4 As Double
    Dim IC5 As Double
    Dim NC5 As Double
    Dim i As Integer

    ok = True

    GasTotalCromatografico = 0

    For Each objPozoAnalisComponente In objPozoAnalis.PozoAnalisComponentes
    
        Select Case objPozoAnalisComponente.NumeroComponente
            Case Is = 1
                Metano = objPozoAnalisComponente.Externo
            Case Is = 2
                Etano = objPozoAnalisComponente.Externo
            Case Is = 3
                Propano = objPozoAnalisComponente.Externo
            Case Is = 4
                IsoButano = objPozoAnalisComponente.Externo
            Case Is = 5
                NormalButano = objPozoAnalisComponente.Externo
            Case Is = 6
                IsoPentano = objPozoAnalisComponente.Externo
            Case Is = 7
                NormalPentano = objPozoAnalisComponente.Externo
        End Select
    Next
    If Etano <> 0 Then 'C2
        Bar2 = Format(Metano / Etano, "0.00")
    Else
        Bar2 = 0
    End If
    If Propano <> 0 Then 'C3
        Bar3 = Format(Metano / Propano, "0.00")
        Cous1 = Format((Etano / Propano) * 10, "0.00")
        CH = Format((IsoButano + NormalButano + IsoPentano + NormalPentano) / Propano, "0.00")
        Geo1 = Format((Metano + Etano) / Propano, "0.00")
    Else
        Bar3 = 0
        Cous1 = 0
        CH = 0
        Geo1 = 0
    End If
    If IsoButano + NormalButano <> 0 Then 'IC4 + NC4
        Bar4 = Format(Metano / (IsoButano + NormalButano), "0.00")
    Else
        Bar4 = 0
    End If
    If IsoPentano + NormalPentano <> 0 Then 'IC5 + NC5
        Bar5 = Format(Metano / (IsoPentano + NormalPentano), "0.00")
    Else
        Bar5 = 0
    End If
    If Etano + Propano + IsoButano + NormalButano + IsoPentano + NormalPentano <> 0 Then 'C2 + C3 + IC4 + NC4 + IC5 + NC5
        Cous2 = Format(Metano / (Etano + Propano + IsoButano + NormalButano + IsoPentano + NormalPentano), "0.00")
    Else
        Cous2 = 0
    End If
    If Metano + Etano + Propano + IsoButano + NormalButano + IsoPentano + NormalPentano <> 0 Then 'C1 + C2 + C3 + IC4 + NC4 + IC5 + NC5
        WH = Format(((Etano + Propano + IsoButano + NormalButano + IsoPentano + NormalPentano) / (Metano + Etano + Propano + IsoButano + NormalButano + IsoPentano + NormalPentano)) * 100, "0.00")
        SnGeo = Format(((Metano + Etano) / (Metano + Etano + Propano + IsoButano + NormalButano + IsoPentano + NormalPentano)) * 100, "0.00")
        Geo3 = Format(((Propano + IsoButano + NormalButano + IsoPentano + NormalPentano) / (Metano + Etano + Propano + IsoButano + NormalButano + IsoPentano + NormalPentano)) * 100, "0.00")
        Geo4 = Format(((IsoButano + NormalButano + IsoPentano + NormalPentano) / (Metano + Etano + Propano + IsoButano + NormalButano + IsoPentano + NormalPentano)) * 100, "0.00")
        C1 = Format((Metano / (Metano + Etano + Propano + IsoButano + NormalButano + IsoPentano + NormalPentano)) * 100, "0.00")
        C2 = Format((Etano / (Metano + Etano + Propano + IsoButano + NormalButano + IsoPentano + NormalPentano)) * 100, "0.00")
        C3 = Format((Propano / (Metano + Etano + Propano + IsoButano + NormalButano + IsoPentano + NormalPentano)) * 100, "0.00")
        C4 = Format(((IsoButano + NormalButano) / (Metano + Etano + Propano + IsoButano + NormalButano + IsoPentano + NormalPentano)) * 100, "0.00")
        C5 = Format(((IsoPentano + NormalPentano) / (Metano + Etano + Propano + IsoButano + NormalButano + IsoPentano + NormalPentano)) * 100, "0.00")
        IC4 = Format((IsoButano / (Metano + Etano + Propano + IsoButano + NormalButano + IsoPentano + NormalPentano)) * 100, "0.00")
        NC4 = Format((NormalButano / (Metano + Etano + Propano + IsoButano + NormalButano + IsoPentano + NormalPentano)) * 100, "0.00")
        IC5 = Format((IsoPentano / (Metano + Etano + Propano + IsoButano + NormalButano + IsoPentano + NormalPentano)) * 100, "0.00")
        NC5 = Format((NormalPentano / (Metano + Etano + Propano + IsoButano + NormalButano + IsoPentano + NormalPentano)) * 100, "0.00")
    Else
        WH = 0
        SnGeo = 0
        Geo3 = 0
        Geo4 = 0
        C1 = 0
        C2 = 0
        C3 = 0
        C4 = 0
        C5 = 0
        IC4 = 0
        NC4 = 0
        IC5 = 0
        NC5 = 0
    End If

    If Propano + IsoButano + NormalButano + IsoPentano + NormalPentano <> 0 Then 'C3 + IC4 + NC4 + IC5 + NC5
        BH = Format((Metano + Etano) / (Propano + IsoButano + NormalButano + IsoPentano + NormalPentano), "0.00")
    Else
        BH = 0
    End If
    If IsoButano + NormalButano + IsoPentano + NormalPentano <> 0 Then 'IC4 + NC4 + IC5 + NC5
        Geo2 = Format((Metano + Etano) / (IsoButano + NormalButano + IsoPentano + NormalPentano), "0.00")
    Else
        Geo2 = 0
    End If
    
    For Each objPozoAnalisComponente In objPozoAnalis.PozoAnalisComponentes
    
        Select Case objPozoAnalisComponente.NumeroComponente
            Case Is = 1
                objPozoAnalisComponente.NormArea = C1
            Case Is = 2
                objPozoAnalisComponente.NormArea = C2
            Case Is = 3
                objPozoAnalisComponente.NormArea = C3
            Case Is = 4
                objPozoAnalisComponente.NormArea = IC4
            Case Is = 5
                objPozoAnalisComponente.NormArea = NC4
            Case Is = 6
                objPozoAnalisComponente.NormArea = IC5
            Case Is = 7
                objPozoAnalisComponente.NormArea = NC5
        End Select
    Next
    
    objPozoAnalis.Bar2 = Bar2
    objPozoAnalis.Bar3 = Bar3
    objPozoAnalis.Bar4 = Bar4
    objPozoAnalis.Bar5 = Bar5
    objPozoAnalis.Cous1 = Cous1
    objPozoAnalis.Cous2 = Cous2
    objPozoAnalis.WH = WH
    objPozoAnalis.BH = BH
    objPozoAnalis.CH = CH
    objPozoAnalis.SnGeo = SnGeo
    objPozoAnalis.Geo1 = Geo1
    objPozoAnalis.Geo2 = Geo2
    objPozoAnalis.Geo3 = Geo3
    objPozoAnalis.Geo4 = Geo4
    objPozoAnalis.C1 = C1
    objPozoAnalis.C2 = C2
    objPozoAnalis.C3 = C3
    objPozoAnalis.C4 = C4
    objPozoAnalis.IC4 = IC4
    objPozoAnalis.NC4 = NC4
    objPozoAnalis.C5 = C5
    objPozoAnalis.IC5 = IC5
    objPozoAnalis.NC5 = NC5

End Sub


Public Function OcultarFormulario(NombreFormulario As String) As Boolean
    
    Dim i As Integer
    
    i = 0
    
    Do Until i = Forms.Count
        
        If Forms(i).Name <> NombreFormulario And Forms(i).Name <> "MdiPrincipal" Then
            
            Forms(i).Hide
            
        End If
        
        i = i + 1
        
    Loop
    
End Function

Public Function PosicionarSimpleCombo(Combo As ComboBox, Codigo As Variant) As Boolean
    
    Dim Encontrado As Boolean
    Dim i As Integer

    On Error GoTo Error

    PosicionarSimpleCombo = False

    i = 0
    Encontrado = False

    Do Until i = Combo.ListCount Or Encontrado

        If Combo.ItemData(i) = Codigo Then

            Combo.ListIndex = i

            Encontrado = True

            PosicionarSimpleCombo = True

        Else

            i = i + 1

        End If

    Loop

Error:

    If Err.Number <> 0 Then

        PosicionarSimpleCombo = False

        frmMsg.MostrarMsg "Módulo: PosicionarSimpleCombo." & Chr(10) & Err.Number & " - " & Err.Description, "Error", MdiPrincipal

        Err.Clear

    End If
    
End Function

Public Function ObtenerSpanEnDías(ByVal FechaDesde As Date, ByVal FechaHasta As Date) As Double
    ObtenerSpanEnDías = CDbl(DateDiff("s", FechaDesde, FechaHasta)) / 60# / 60# / 24#
End Function

Public Function InvertirFecha(ByVal fecha As String) As String
    Dim dia As String, mes As String, anio As String
    mes = Left(fecha, 2)
    dia = Mid(fecha, 4, 2)
    anio = Right(fecha, 4)
    InvertirFecha = dia & "/" & mes & "/" & anio
End Function

Public Sub leer_txt(ok As Boolean, ArchivoSeleccionado As String)
    
    Dim LineaTexto                          As String
    Dim ArrCadenas()                        As String
    Dim Valor                               As String
    Dim StrSql                              As String
    Dim NumeroAnalisis                      As Long
    Dim MayorOrden                          As Recordset
    Dim rChequear                           As Recordset
    Dim GasTotalCromatografico              As Integer
    Dim objPozoAnalis                       As New clsPozoAnalis
    Dim objPozoAnalisComponente             As ClsPozoAnalisComponente
    
    Dim fs                                  As New Scripting.FileSystemObject
    Dim archivo                             As TextStream
    Dim SumaArea                            As Double
    
    On Error GoTo Error
    
    If Not gObjPozoActivo Is Nothing Then
    
        ok = True
        Set gObjPozoActivo.PozoAnalisis.Datos = gDatos
        objPozoAnalis.IdPozo = gObjPozoActivo.IdPozo
        objPozoAnalis.NumeroAnalisis = gObjPozoActivo.PozoAnalisis.dbMax(gObjPozoActivo.IdPozo) + 1
        ArchivoSeleccionado = gCaminoArchivos & gObjConfiguracion.ArchivoDeResultados
        objPozoAnalis.archivo = " "
''        'Open ArchivoSeleccionado For Input As #1   ' Abre el archivo.
        
''        If EOF(1) Then
''            Debug.Print "Archivo vacío"
''        End If
        Set archivo = fs.OpenTextFile(ArchivoSeleccionado, ForReading)
        'archivo.WriteLine "Objeto = " & Me.Objeto & " - Método: " & Me.Metodo & "Error: " & Me.Number & " - " & Me.Description
        

        'Do While Not EOF(1) ' Repite el bucle hasta Encontrar la linea que tiene la fecha.
        Do While Not archivo.AtEndOfStream
            LineaTexto = archivo.ReadLine     ' Lee el carácter en la variable.
            If InStr(LineaTexto, ",") Then 'es la línea que contiene la fecha
                Valor = Right(LineaTexto, Len(LineaTexto) - InStr(LineaTexto, ","))
                objPozoAnalis.fecha = ExtraerFecha(Valor)
                Exit Do
            End If
        Loop
        objPozoAnalis.Profundidad = ProfundidadAnalisis
        objPozoAnalis.IdPozo = gObjPozoActivo.IdPozo
        objPozoAnalis.CodigoTipoDeAnalisis = 1
        objPozoAnalis.EFast = 1
        objPozoAnalis.ENormal = 1
        objPozoAnalis.GasTotal = GasTotal
        'Do While Not EOF(1) ' Repite el bucle hasta el final del archivo.
        Do While Not archivo.AtEndOfStream
            'Line Input #1, LineaTexto    ' Lee el carácter en la variable.
            LineaTexto = archivo.ReadLine
            If InStr(UCase(LineaTexto), "METANO") <> 0 Then
                objPozoAnalis.PozoAnalisComponentes.colAgregar ArmarComponenteDeCadena(LineaTexto, objPozoAnalis, 1)
            ElseIf InStr(UCase(LineaTexto), "ETANO") <> 0 Then
                objPozoAnalis.PozoAnalisComponentes.colAgregar ArmarComponenteDeCadena(LineaTexto, objPozoAnalis, 2)
            ElseIf InStr(UCase(LineaTexto), "PROPANO") <> 0 Then
                objPozoAnalis.PozoAnalisComponentes.colAgregar ArmarComponenteDeCadena(LineaTexto, objPozoAnalis, 3)
            ElseIf InStr(UCase(LineaTexto), "ISOBUTANO") <> 0 Then
                objPozoAnalis.PozoAnalisComponentes.colAgregar ArmarComponenteDeCadena(LineaTexto, objPozoAnalis, 4)
            ElseIf InStr(UCase(LineaTexto), "NORMALBUTANO") <> 0 Then
                objPozoAnalis.PozoAnalisComponentes.colAgregar ArmarComponenteDeCadena(LineaTexto, objPozoAnalis, 5)
            ElseIf InStr(UCase(LineaTexto), "ISOPENTANO") <> 0 Then
                objPozoAnalis.PozoAnalisComponentes.colAgregar ArmarComponenteDeCadena(LineaTexto, objPozoAnalis, 6)
            ElseIf InStr(UCase(LineaTexto), "NORMALPENTANO") <> 0 Then
                objPozoAnalis.PozoAnalisComponentes.colAgregar ArmarComponenteDeCadena(LineaTexto, objPozoAnalis, 7)
            ElseIf InStr(UCase(LineaTexto), "DIOXIDOCARBONO") <> 0 Then
                ArrCadenas = Split(LineaTexto, "|")
                objPozoAnalis.CO2 = Trim(Val(ArrCadenas(5)))
            ElseIf InStr(UCase(LineaTexto), "SULFUROHIDROGENO") <> 0 Then
                ArrCadenas = Split(LineaTexto, "|")
                objPozoAnalis.SH2 = Trim(Val(ArrCadenas(5)))
            End If
        Loop
''        Close #1    ' Cierra el archivo.
        archivo.Close
    
        ActualizarGasTotalCromatografico objPozoAnalis, ok
        CalcularRelacionesCromatograficas objPozoAnalis, "AGREGAR", ok
        
        Set gObjPozoActivo.PozoAnalisis.Datos = gDatos
        ok = gObjPozoActivo.PozoAnalisis.dbAgregar(objPozoAnalis)
        If ok Then
            Set objPozoAnalis.PozoAnalisComponentes.Datos = gDatos
            SumaArea = objPozoAnalis.PozoAnalisComponentes.getSumaArea
            
            For Each objPozoAnalisComponente In objPozoAnalis.PozoAnalisComponentes
                objPozoAnalis.PozoAnalisComponentes.SetearNormArea objPozoAnalisComponente, SumaArea
                If Not objPozoAnalis.PozoAnalisComponentes.dbAgregar(objPozoAnalisComponente) Then
                    ok = False
                    Exit For
                End If
            Next
        End If
        
        If ok Then
            gObjPozoActivo.PozoAnalisis.colAgregar objPozoAnalis
            If FormularioCargado("FrmAnalisis") Then
                AgregarDatosTemporalAlList objPozoAnalis
            End If
        End If
    
    End If
Error:

    If Err.Number <> 0 Then
        If Err.Number = 55 Then
            archivo.Close
        ElseIf Err.Number <> 53 Then
            gDatos.Error.Objeto = "Modulo Funciones Generales"
            gDatos.Error.Metodo = "leer_txt"
            gDatos.Error.Number = Err.Number
            gDatos.Error.Description = Err.Description
            gDatos.Error.Mostrar
        End If
        Err.Clear
    End If
    
End Sub

Public Sub ChequearPegar(UnaLineaTexto As String, UnaLectura As String)

Dim EstaEnLectura As Boolean
Dim EstaEnLista As Boolean

EstaEnLectura = False
EstaEnLista = False



EstaEnLectura = InStr(1, UCase(UnaLectura), "METANO") <> 0
EstaEnLista = InStr(1, UCase(UnaLineaTexto), "METANO") <> 0
    If EstaEnLectura And Not EstaEnLista Then
        UnaLineaTexto = UnaLineaTexto & UnaLectura & vbNewLine
       ' Debug.Print UnaLineaTexto
    End If

EstaEnLectura = InStr(1, UCase(UnaLectura), Chr(9) & "ETANO") <> 0
EstaEnLista = InStr(1, UCase(UnaLineaTexto), Chr(9) & "ETANO") <> 0
    If EstaEnLectura And Not EstaEnLista Then
        UnaLineaTexto = UnaLineaTexto & UnaLectura & vbNewLine
        'Debug.Print UnaLineaTexto
    End If

EstaEnLectura = InStr(1, UCase(UnaLectura), "PROPANO") <> 0
EstaEnLista = InStr(1, UCase(UnaLineaTexto), "PROPANO") <> 0
    If EstaEnLectura And Not EstaEnLista Then
        UnaLineaTexto = UnaLineaTexto & UnaLectura & vbNewLine
        'Debug.Print UnaLineaTexto
    End If


EstaEnLectura = InStr(1, UCase(UnaLectura), "ISOBUTANO") <> 0
EstaEnLista = InStr(1, UCase(UnaLineaTexto), "ISOBUTANO") <> 0
    If EstaEnLectura And Not EstaEnLista Then
        UnaLineaTexto = UnaLineaTexto & UnaLectura & vbNewLine
'        Debug.Print UnaLineaTexto
    End If

EstaEnLectura = InStr(1, UCase(UnaLectura), "NORMALBUTANO") <> 0
EstaEnLista = InStr(1, UCase(UnaLineaTexto), "NORMALBUTANO") <> 0
    If EstaEnLectura And Not EstaEnLista Then
        UnaLineaTexto = UnaLineaTexto & UnaLectura & vbNewLine
'        Debug.Print UnaLineaTexto
    End If

EstaEnLectura = InStr(1, UCase(UnaLectura), "ISOPENTANO") <> 0
EstaEnLista = InStr(1, UCase(UnaLineaTexto), "ISOPENTANO") <> 0
    If EstaEnLectura And Not EstaEnLista Then
        UnaLineaTexto = UnaLineaTexto & UnaLectura & vbNewLine
      '  Debug.Print UnaLineaTexto
    End If

EstaEnLectura = InStr(1, UCase(UnaLectura), "NORMALPENTANO") <> 0
EstaEnLista = InStr(1, UCase(UnaLineaTexto), "NORMALPENTANO") <> 0
    If EstaEnLectura And Not EstaEnLista Then
        UnaLineaTexto = UnaLineaTexto & UnaLectura & vbNewLine
        'Debug.Print UnaLineaTexto
    End If

EstaEnLectura = InStr(1, UCase(UnaLectura), "CO2") <> 0
EstaEnLista = InStr(1, UCase(UnaLineaTexto), "CO2") <> 0
    If EstaEnLectura And Not EstaEnLista Then
        UnaLineaTexto = UnaLineaTexto & UnaLectura & vbNewLine
        Debug.Print UnaLineaTexto
    End If

EstaEnLectura = InStr(1, UCase(UnaLectura), "DIOXIDODECARBONO") <> 0
EstaEnLista = InStr(1, UCase(UnaLineaTexto), "DIOXIDODECARBONO") <> 0
    If EstaEnLectura And Not EstaEnLista Then
        UnaLineaTexto = UnaLineaTexto & UnaLectura & vbNewLine
'        Debug.Print UnaLineaTexto
    End If
    


End Sub



Public Sub Leer_txt_Varian(ok As Boolean, ArchivoSeleccionado As String, ArchivoSeleccionadoDos As String)

Dim LineaTexto                          As String
    
    
    Dim Array_pico()                        As String
    Dim Array_fecha()                       As String
    Dim componente_pico                     As String
    
    Dim token                               As String
    Dim Valor                               As String
    Dim StrSql                              As String
    Dim NumeroAnalisis                      As Long
    
    Dim MayorOrden                          As Recordset
    Dim rChequear                           As Recordset
    
    Dim GasTotalCromatografico              As Integer
    Dim i                                   As Integer
    Dim objPozoAnalis                       As New clsPozoAnalis
    Dim objPozoAnalisComponente             As ClsPozoAnalisComponente
    
    Dim fs                                  As New Scripting.FileSystemObject
    Dim archivo                             As TextStream
    Dim SumaArea                            As Double
    Dim ascci_ant                           As Integer
    Dim toma_carac                          As Boolean
    Dim vieneConFecha                       As Boolean
    
    Dim stemp As String

    
    On Error GoTo Error
    
    If Not gObjPozoActivo Is Nothing Then
    
        ok = True
        Set gObjPozoActivo.PozoAnalisis.Datos = gDatos
        objPozoAnalis.IdPozo = gObjPozoActivo.IdPozo
        objPozoAnalis.NumeroAnalisis = gObjPozoActivo.PozoAnalisis.dbMax(gObjPozoActivo.IdPozo) + 1
        objPozoAnalis.archivo = " "
        objPozoAnalis.CodigoTipoDeAnalisis = 1
        
        LineaTexto = ""
            
        
        If ArchivoSeleccionado <> "" Then
            
            Set archivo = fs.OpenTextFile(ArchivoSeleccionado, ForReading)
            Do
                     ' Lee la linea del archivo, en teoría viene una sola linea.
                ChequearPegar LineaTexto, archivo.ReadLine
            Loop Until archivo.AtEndOfStream
            archivo.Close
        End If
        Debug.Print LineaTexto
        
        If ArchivoSeleccionadoDos <> "" Then
            Set archivo = fs.OpenTextFile(ArchivoSeleccionadoDos, ForReading)
            Do
                ChequearPegar LineaTexto, archivo.ReadLine
            
            Loop Until archivo.AtEndOfStream
            archivo.Close
        End If
        Debug.Print LineaTexto
        
        If ArchivoSeleccionado = "" And ArchivoSeleccionadoDos = "" Then GoTo Error
        
        Array_tokens = Split(LineaTexto, vbNewLine)

        'Prepara el obj para ir a base
        objPozoAnalis.Profundidad = ProfundidadAnalisis
        objPozoAnalis.IdPozo = gObjPozoActivo.IdPozo
        objPozoAnalis.EFast = 1
        objPozoAnalis.ENormal = 1
        objPozoAnalis.GasTotal = GasTotal
        
        objPozoAnalis.fecha = Now
        
        i = 0
        
        objPozoAnalis.PozoAnalisComponentes.colAgregar ArmarComponenteDeCadenaSkyCrome(objPozoAnalis.NumeroAnalisis, 0, "METANO", 0, 1, 0, "ppm")
        objPozoAnalis.PozoAnalisComponentes.colAgregar ArmarComponenteDeCadenaSkyCrome(objPozoAnalis.NumeroAnalisis, 0, "ETANO", 0, 2, 0, "ppm")
        objPozoAnalis.PozoAnalisComponentes.colAgregar ArmarComponenteDeCadenaSkyCrome(objPozoAnalis.NumeroAnalisis, 0, "PROPANO", 0, 3, 0, "ppm")
        objPozoAnalis.PozoAnalisComponentes.colAgregar ArmarComponenteDeCadenaSkyCrome(objPozoAnalis.NumeroAnalisis, 0, "ISOBUTANO", 0, 4, 0, "ppm")
        objPozoAnalis.PozoAnalisComponentes.colAgregar ArmarComponenteDeCadenaSkyCrome(objPozoAnalis.NumeroAnalisis, 0, "NORMALBUTANO", 0, 5, 0, "ppm")
        objPozoAnalis.PozoAnalisComponentes.colAgregar ArmarComponenteDeCadenaSkyCrome(objPozoAnalis.NumeroAnalisis, 0, "ISOPENTANO", 0, 6, 0, "ppm")
        objPozoAnalis.PozoAnalisComponentes.colAgregar ArmarComponenteDeCadenaSkyCrome(objPozoAnalis.NumeroAnalisis, 0, "NORMALPENTANO", 0, 7, 0, "ppm")
        
        Do Until i > UBound(Array_tokens) - 1
        
            Array_tokensHijos = Split(Array_tokens(i), vbTab)
            
            Set objPozoAnalisComponente = Nothing
                        
            If Trim(UCase(Array_tokensHijos(1))) = "METANO" Then
                Set objPozoAnalisComponente = objPozoAnalis.PozoAnalisComponentes.colComponente(objPozoAnalis.IdPozo, objPozoAnalis.NumeroAnalisis, 1)
            ElseIf Trim(UCase(Array_tokensHijos(1))) = "ETANO" Then
                Set objPozoAnalisComponente = objPozoAnalis.PozoAnalisComponentes.colComponente(objPozoAnalis.IdPozo, objPozoAnalis.NumeroAnalisis, 2)
            ElseIf Trim(UCase(Array_tokensHijos(1))) = "PROPANO" Then
                Set objPozoAnalisComponente = objPozoAnalis.PozoAnalisComponentes.colComponente(objPozoAnalis.IdPozo, objPozoAnalis.NumeroAnalisis, 3)
            ElseIf Trim(UCase(Array_tokensHijos(1))) = "ISOBUTANO" Then
                Set objPozoAnalisComponente = objPozoAnalis.PozoAnalisComponentes.colComponente(objPozoAnalis.IdPozo, objPozoAnalis.NumeroAnalisis, 4)
            ElseIf Trim(UCase(Array_tokensHijos(1))) = "NORMALBUTANO" Then
                Set objPozoAnalisComponente = objPozoAnalis.PozoAnalisComponentes.colComponente(objPozoAnalis.IdPozo, objPozoAnalis.NumeroAnalisis, 5)
            ElseIf Trim(UCase(Array_tokensHijos(1))) = "ISOPENTANO" Then
                Set objPozoAnalisComponente = objPozoAnalis.PozoAnalisComponentes.colComponente(objPozoAnalis.IdPozo, objPozoAnalis.NumeroAnalisis, 6)
            ElseIf Trim(UCase(Array_tokensHijos(1))) = "NORMALPENTANO" Then
                Set objPozoAnalisComponente = objPozoAnalis.PozoAnalisComponentes.colComponente(objPozoAnalis.IdPozo, objPozoAnalis.NumeroAnalisis, 7)
            End If
            
            If Not objPozoAnalisComponente Is Nothing Then
            
                objPozoAnalisComponente.Area = objPozoAnalisComponente.Area + Array_tokensHijos(5)
                objPozoAnalisComponente.Externo = objPozoAnalisComponente.Externo + Array_tokensHijos(3)
                objPozoAnalisComponente.Retention = objPozoAnalisComponente.Retention + Array_tokensHijos(2)
                
            End If
            
                            
            i = i + 1
        Loop
        
     
    
        ActualizarGasTotalCromatografico objPozoAnalis, ok
        CalcularRelacionesCromatograficas objPozoAnalis, "AGREGAR", ok
        
        Set gObjPozoActivo.PozoAnalisis.Datos = gDatos
        ok = gObjPozoActivo.PozoAnalisis.dbAgregar(objPozoAnalis)
        If ok Then
            Set objPozoAnalis.PozoAnalisComponentes.Datos = gDatos
            SumaArea = objPozoAnalis.PozoAnalisComponentes.getSumaArea
            For Each objPozoAnalisComponente In objPozoAnalis.PozoAnalisComponentes
                objPozoAnalis.PozoAnalisComponentes.SetearNormArea objPozoAnalisComponente, SumaArea
                If Not objPozoAnalis.PozoAnalisComponentes.dbAgregar(objPozoAnalisComponente) Then
                    ok = False
                    Exit For
                End If
            Next
        End If
        
        If ok Then
            gObjPozoActivo.PozoAnalisis.colAgregar objPozoAnalis
            If FormularioCargado("FrmAnalisis") Then
                AgregarDatosTemporalAlList objPozoAnalis
            End If
        End If
    
    End If
Error:

    If Err.Number <> 0 Then
        If Err.Number = 55 Then
            archivo.Close
        ElseIf Err.Number <> 53 Then
            gDatos.Error.Objeto = "Modulo Funciones Generales"
            gDatos.Error.Metodo = "Leer_txt_Varian"
            gDatos.Error.Number = Err.Number
            gDatos.Error.Description = Err.Description
            gDatos.Error.Mostrar
        End If
        Err.Clear
    End If
    


























End Sub

Public Sub Leer_txt_skycrhome(ok As Boolean, ArchivoSeleccionado As String)
    
    Dim LineaTexto                          As String
    Dim Array_tokens()                      As String
    Dim Array_pico()                        As String
    Dim Array_fecha()                       As String
    Dim componente_pico                     As String
    
    Dim token                               As String
    Dim Valor                               As String
    Dim StrSql                              As String
    Dim NumeroAnalisis                      As Long
    
    Dim MayorOrden                          As Recordset
    Dim rChequear                           As Recordset
    
    Dim GasTotalCromatografico              As Integer
    Dim i                                   As Integer
    Dim objPozoAnalis                       As New clsPozoAnalis
    Dim objPozoAnalisComponente             As ClsPozoAnalisComponente
    
    Dim fs                                  As New Scripting.FileSystemObject
    Dim archivo                             As TextStream
    Dim SumaArea                            As Double
    Dim ascci_ant                           As Integer
    Dim toma_carac                          As Boolean
    Dim vieneConFecha                       As Boolean
    
    Dim stemp As String

    
    On Error GoTo Error
    
    If Not gObjPozoActivo Is Nothing Then
    
        ok = True
        Set gObjPozoActivo.PozoAnalisis.Datos = gDatos
        objPozoAnalis.IdPozo = gObjPozoActivo.IdPozo
        objPozoAnalis.NumeroAnalisis = gObjPozoActivo.PozoAnalisis.dbMax(gObjPozoActivo.IdPozo) + 1
        
        vieneConFecha = InStr(1, ArchivoSeleccionado, "_")
        
        objPozoAnalis.archivo = " "
        
        Set archivo = fs.OpenTextFile(ArchivoSeleccionado, ForReading)
        LineaTexto = archivo.ReadLine     ' Lee la linea del archivo, en teoría viene una sola linea.
        stemp = ""
        ascci_ant = 0
        For i = 1 To Len(LineaTexto)
            
            If Mid(LineaTexto, i, 1) <> Chr(32) And Mid(LineaTexto, i, 1) <> Chr(13) Then
                Debug.Print Asc(Mid(LineaTexto, i, 1)) & " _ " & Mid(LineaTexto, i, 1)
                If Asc(Mid(LineaTexto, i, 1)) = 9 Then
                    If ascci_ant <> 9 Then
                        stemp = stemp & Mid(LineaTexto, i, 1)
                    End If
                Else
                    If ascci_ant = 109 Then
                        'es la terminación del componente ppm (letra m) y la actual no es un tab le agrego un tab
                        stemp = stemp & vbTab
                    End If
                    stemp = stemp & Mid(LineaTexto, i, 1)
                End If
                ascci_ant = Asc(Mid(LineaTexto, i, 1))
            End If
        Next
        LineaTexto = stemp
        
        'Array_tokens = Split(LineaTexto, Chr(32))
        Array_tokens = Split(LineaTexto, vbTab)
        
        objPozoAnalis.Profundidad = ProfundidadAnalisis
        objPozoAnalis.IdPozo = gObjPozoActivo.IdPozo
        objPozoAnalis.EFast = 1
        objPozoAnalis.ENormal = 1
        objPozoAnalis.GasTotal = GasTotal
        
        If Not vieneConFecha Then
            objPozoAnalis.fecha = Array_tokens(0) & " " & Array_tokens(1)
        Else
            objPozoAnalis.fecha = Now
        End If
        If vieneConFecha Then
            i = 1
        Else
            i = 2
        End If
        
        objPozoAnalis.PozoAnalisComponentes.colAgregar ArmarComponenteDeCadenaSkyCrome(objPozoAnalis.NumeroAnalisis, 0, "METANO", 0, 1, 0, "ppm")
        objPozoAnalis.PozoAnalisComponentes.colAgregar ArmarComponenteDeCadenaSkyCrome(objPozoAnalis.NumeroAnalisis, 0, "ETANO", 0, 2, 0, "ppm")
        objPozoAnalis.PozoAnalisComponentes.colAgregar ArmarComponenteDeCadenaSkyCrome(objPozoAnalis.NumeroAnalisis, 0, "PROPANO", 0, 3, 0, "ppm")
        objPozoAnalis.PozoAnalisComponentes.colAgregar ArmarComponenteDeCadenaSkyCrome(objPozoAnalis.NumeroAnalisis, 0, "ISOBUTANO", 0, 4, 0, "ppm")
        objPozoAnalis.PozoAnalisComponentes.colAgregar ArmarComponenteDeCadenaSkyCrome(objPozoAnalis.NumeroAnalisis, 0, "NORMALBUTANO", 0, 5, 0, "ppm")
        objPozoAnalis.PozoAnalisComponentes.colAgregar ArmarComponenteDeCadenaSkyCrome(objPozoAnalis.NumeroAnalisis, 0, "ISOPENTANO", 0, 6, 0, "ppm")
        objPozoAnalis.PozoAnalisComponentes.colAgregar ArmarComponenteDeCadenaSkyCrome(objPozoAnalis.NumeroAnalisis, 0, "NORMALPENTANO", 0, 7, 0, "ppm")
        
        Do Until i > UBound(Array_tokens)
        
            Set objPozoAnalisComponente = Nothing
            
            If UCase(Array_tokens(i + 3)) = "METANO" Then
                Set objPozoAnalisComponente = objPozoAnalis.PozoAnalisComponentes.colComponente(objPozoAnalis.IdPozo, objPozoAnalis.NumeroAnalisis, 1)
            ElseIf UCase(Array_tokens(i + 3)) = "ETANO" Then
                Set objPozoAnalisComponente = objPozoAnalis.PozoAnalisComponentes.colComponente(objPozoAnalis.IdPozo, objPozoAnalis.NumeroAnalisis, 2)
            ElseIf UCase(Array_tokens(i + 3)) = "PROPANO" Then
                Set objPozoAnalisComponente = objPozoAnalis.PozoAnalisComponentes.colComponente(objPozoAnalis.IdPozo, objPozoAnalis.NumeroAnalisis, 3)
            ElseIf UCase(Array_tokens(i + 3)) = "ISOBUTANO" Then
                Set objPozoAnalisComponente = objPozoAnalis.PozoAnalisComponentes.colComponente(objPozoAnalis.IdPozo, objPozoAnalis.NumeroAnalisis, 4)
            ElseIf UCase(Array_tokens(i + 3)) = "NORMALBUTANO" Then
                Set objPozoAnalisComponente = objPozoAnalis.PozoAnalisComponentes.colComponente(objPozoAnalis.IdPozo, objPozoAnalis.NumeroAnalisis, 5)
            ElseIf UCase(Array_tokens(i + 3)) = "ISOPENTANO" Then
                Set objPozoAnalisComponente = objPozoAnalis.PozoAnalisComponentes.colComponente(objPozoAnalis.IdPozo, objPozoAnalis.NumeroAnalisis, 6)
            ElseIf UCase(Array_tokens(i + 3)) = "NORMALPENTANO" Then
                Set objPozoAnalisComponente = objPozoAnalis.PozoAnalisComponentes.colComponente(objPozoAnalis.IdPozo, objPozoAnalis.NumeroAnalisis, 7)
            End If
            
            If Not objPozoAnalisComponente Is Nothing Then
            
                objPozoAnalisComponente.Area = objPozoAnalisComponente.Area + Array_tokens(i + 2)
                objPozoAnalisComponente.Externo = objPozoAnalisComponente.Externo + Array_tokens(i + 4)
                objPozoAnalisComponente.Retention = objPozoAnalisComponente.Retention + Array_tokens(i + 1)
                
            End If
            
                            
            i = i + 6
        Loop
        
'''''        If Array_tokens(i) = 1 Then
'''''            objPozoAnalis.PozoAnalisComponentes.colAgregar ArmarComponenteDeCadenaSkyCrome(objPozoAnalis.NumeroAnalisis, , , , , , Array_tokens(i + 5))
'''''        End If
'''''        i = i + 6
'''''        If Array_tokens(i) = 2 Then
'''''            objPozoAnalis.PozoAnalisComponentes.colAgregar ArmarComponenteDeCadenaSkyCrome(objPozoAnalis.NumeroAnalisis, Array_tokens(i + 2), Array_tokens(i + 3), Array_tokens(i + 4), 3, Array_tokens(i + 1), Array_tokens(i + 5))
'''''        End If
'''''        i = i + 6
'''''        If Array_tokens(i) = 3 Then
'''''            objPozoAnalis.PozoAnalisComponentes.colAgregar ArmarComponenteDeCadenaSkyCrome(objPozoAnalis.NumeroAnalisis, Array_tokens(i + 2), Array_tokens(i + 3), Array_tokens(i + 4), 4, Array_tokens(i + 1), Array_tokens(i + 5))
'''''        End If
'''''
'''''        i = i + 6
'''''        If Array_tokens(i) = 4 Then
'''''            objPozoAnalis.PozoAnalisComponentes.colAgregar ArmarComponenteDeCadenaSkyCrome(objPozoAnalis.NumeroAnalisis, Array_tokens(i + 2), Array_tokens(i + 3), Array_tokens(i + 4), 5, Array_tokens(i + 1), Array_tokens(i + 5))
'''''        End If
'''''        i = i + 6
'''''        If Array_tokens(i) = 5 Then
'''''            objPozoAnalis.PozoAnalisComponentes.colAgregar ArmarComponenteDeCadenaSkyCrome(objPozoAnalis.NumeroAnalisis, Array_tokens(i + 2), Array_tokens(i + 3), Array_tokens(i + 4), 6, Array_tokens(i + 1), Array_tokens(i + 5))
'''''        End If
'''''        i = i + 6
'''''        If Array_tokens(i) = 6 Then
'''''            objPozoAnalis.PozoAnalisComponentes.colAgregar ArmarComponenteDeCadenaSkyCrome(objPozoAnalis.NumeroAnalisis, Array_tokens(i + 2), Array_tokens(i + 3), Array_tokens(i + 4), 7, Array_tokens(i + 1), Array_tokens(i + 5))
'''''        End If

        archivo.Close
    
        ActualizarGasTotalCromatografico objPozoAnalis, ok
        CalcularRelacionesCromatograficas objPozoAnalis, "AGREGAR", ok
        
        Set gObjPozoActivo.PozoAnalisis.Datos = gDatos
        ok = gObjPozoActivo.PozoAnalisis.dbAgregar(objPozoAnalis)
        If ok Then
            Set objPozoAnalis.PozoAnalisComponentes.Datos = gDatos
            SumaArea = objPozoAnalis.PozoAnalisComponentes.getSumaArea
            For Each objPozoAnalisComponente In objPozoAnalis.PozoAnalisComponentes
                objPozoAnalis.PozoAnalisComponentes.SetearNormArea objPozoAnalisComponente, SumaArea
                If Not objPozoAnalis.PozoAnalisComponentes.dbAgregar(objPozoAnalisComponente) Then
                    ok = False
                    Exit For
                End If
            Next
        End If
        
        If ok Then
            gObjPozoActivo.PozoAnalisis.colAgregar objPozoAnalis
            If FormularioCargado("FrmAnalisis") Then
                AgregarDatosTemporalAlList objPozoAnalis
            End If
        End If
    
    End If
Error:

    If Err.Number <> 0 Then
        If Err.Number = 55 Then
            archivo.Close
        ElseIf Err.Number <> 53 Then
            gDatos.Error.Objeto = "Modulo Funciones Generales"
            gDatos.Error.Metodo = "Leer_txt_skycrhome"
            gDatos.Error.Number = Err.Number
            gDatos.Error.Description = Err.Description
            gDatos.Error.Mostrar
        End If
        Err.Clear
    End If
    
End Sub


Public Function ExtraerFecha(ByVal Valor As String) As Date
    Dim TempStr As String
    TempStr = Valor
    TempStr = Replace(TempStr, " de ", "/")
    TempStr = Replace(TempStr, "Enero", "01")
    TempStr = Replace(TempStr, "Febrero", "02")
    TempStr = Replace(TempStr, "Marzo", "03")
    TempStr = Replace(TempStr, "Abril", "04")
    TempStr = Replace(TempStr, "Mayo", "05")
    TempStr = Replace(TempStr, "Junio", "06")
    TempStr = Replace(TempStr, "Julio", "07")
    TempStr = Replace(TempStr, "Agosto", "08")
    TempStr = Replace(TempStr, "Setiembre", "09")
    TempStr = Replace(TempStr, "Octubre", "10")
    TempStr = Replace(TempStr, "Noviembre", "11")
    TempStr = Replace(TempStr, "Diciembre", "12")
    ExtraerFecha = CDate(TempStr)
    

End Function

Function ArmarComponenteDeCadena(ByVal LineaTexto As String, ByVal objAnalisis As clsPozoAnalis, ByVal NroComponente As Long) As ClsPozoAnalisComponente
    
    Dim objPozoAnalisComponente As New ClsPozoAnalisComponente
    
    Dim ArrCadenas() As String
    
        
    ArrCadenas = Split(LineaTexto, "|")
    Set objPozoAnalisComponente = New ClsPozoAnalisComponente
    objPozoAnalisComponente.Area = Trim(Val(ArrCadenas(3)))
    objPozoAnalisComponente.Component = Trim(ArrCadenas(6))
    On Error Resume Next
    objPozoAnalisComponente.Externo = Trim(Val(ArrCadenas(5))) 'por ahora
    If Err.Number = 6 Then 'hay desbordamiento por númeor demasiado grande
        objPozoAnalisComponente.Externo = 2 ^ 30
        Err.Clear
    End If
    On Error GoTo Error
    objPozoAnalisComponente.IdPozo = gObjPozoActivo.IdPozo
    objPozoAnalisComponente.NumeroAnalisis = objAnalisis.NumeroAnalisis
    objPozoAnalisComponente.NumeroComponente = NroComponente
    objPozoAnalisComponente.Retention = Trim(Val(ArrCadenas(1)))
    objPozoAnalisComponente.Units = "Ppm"
    
    Set ArmarComponenteDeCadena = objPozoAnalisComponente

Error:

    If Err.Number <> 0 Then
        gDatos.Error.Objeto = "Módulo Funciones Generales"
        gDatos.Error.Metodo = "ArmarComponenteDeCadena"
        gDatos.Error.Number = Err.Number
        gDatos.Error.Description = Err.Description
        gDatos.Error.Mostrar
        Err.Clear
    End If
    
End Function


Function ArmarComponenteDeCadenaSkyCrome( _
    ByVal NumeroAnalisis As Long, _
    ByVal Area As String, _
    ByVal Component As String, _
    ByVal Externo As String, _
    ByVal NumeroComponente As String, _
    ByVal Retention As String, _
    ByVal Units As String _
) As ClsPozoAnalisComponente
    
    Dim objPozoAnalisComponente As New ClsPozoAnalisComponente
    
    On Error GoTo Error
    
    Set objPozoAnalisComponente = New ClsPozoAnalisComponente
    
    objPozoAnalisComponente.Area = Area
    objPozoAnalisComponente.Component = Component
    objPozoAnalisComponente.Externo = Externo
    objPozoAnalisComponente.IdPozo = gObjPozoActivo.IdPozo
    objPozoAnalisComponente.NumeroAnalisis = NumeroAnalisis
    objPozoAnalisComponente.NumeroComponente = NumeroComponente
    objPozoAnalisComponente.Retention = Retention
    objPozoAnalisComponente.Units = Units
    
    Set ArmarComponenteDeCadenaSkyCrome = objPozoAnalisComponente

Error:

    If Err.Number <> 0 Then
        gDatos.Error.Objeto = "Módulo Funciones Generales"
        gDatos.Error.Metodo = "ArmarComponenteDeCadenaSkyCrome"
        gDatos.Error.Number = Err.Number
        gDatos.Error.Description = Err.Description
        gDatos.Error.Mostrar
        Err.Clear
    End If
    
End Function

Public Sub SoloNumeros(KeyAscii As Integer)

    Select Case KeyAscii
    
        Case 48 To 57   ' Permite los dígitos
        Case 8      ' Permite el carácter de retroceso
        Case Else
            KeyAscii = 0
            
    End Select

End Sub


Public Sub isDbl(KeyAscii As Integer, Nro As String)

    Select Case KeyAscii
    
        Case 48 To 57   ' Permite los dígitos
        Case 8          ' Permite el carácter de retroceso
        Case (44 Or 46) ' Permite la ,
            If InStr(1, CStr(Nro), ",") <> 0 Then
                KeyAscii = 0
                
            Else
                KeyAscii = 44
            End If

        Case Else
            KeyAscii = 0
            
    End Select
End Sub
