Attribute VB_Name = "Funcion"
Option Explicit

Public ProfundidadPozo As Double
Public ProfundidadTrepano As Double
Public ProfundidadRetorno As Double
Public TVD As Double
Public PesoGancho As Double
Public PesoAplicado As Double
Public MaximoPesoGancho As Double 'Variable inventada, dado que no viene en el paquete
Public Torque As Double
Public RPMMesa As Double
Public RPMFondo As Double
Public PresionBomba As Double
Public CaudalEntrada As Double
Public CaudalRetorno As Double
Public ExMBomba1 As Double
Public ExMBomba2 As Double
Public ExMBomba3 As Double
Public Retorno As Double
Public Bajada As Double
Public Pileta1 As Double
Public Pileta2 As Double
Public Pileta3 As Double
Public Pileta4 As Double
Public Pileta5 As Double
Public Pileta6 As Double
Public Pileta7 As Double
Public Pileta8 As Double
Public Pileta9 As Double
Public Pileta10 As Double
Public TotalNivelPiletas As Double
Public DiferenciaNivelPiletas As Double
Public InicioDiferencia As String
Public ActividadPozo As String

Public AvanceCarrera As Double
Public ROPCarrera As Double
Public HsMotor As Double
Public HsRotacion As Double
Public RevolucionesTotales As Double
Public ROP As Double
Public Cronometraje As Double
Public Penetrometro As Double
Public TripTank As Double
Public Estado As String
Public EstadoAnterior As String
Public ExpD As Double
Public KTF As Double
Public ECD As Double
Public Dc As Double
Public GradientedeFractura As Double
Public PresiónPoral As Double
Public Densidad As Double
Public GBF As Double
Public S As Double
Public Mu As Double
Public TorqueMáximo As Double
Public TorqueMínimo As Double
Public TorquePromedio As Double
Public PesoAplicadoMáximo As Double
Public PesoAplicadoMínimo As Double
Public PesoAplicadoPromedio As Double
Public RPMMáximo As Double
Public RPMMínimo As Double
Public RPMPromedio As Double
Public PresiónBombaMáximo As Double
Public PresiónBombaMínimo As Double
Public PresiónBombaPromedio As Double
Public LongitudTiroManiobra As Double
Public VelocidadTiroManiobra As Double
Public PesoPromedioManiobra As Double
Public PesoMáximoManiobra As Double
Public PesoTeóricoManiobra As Double
Public DespTeóricoAcumulado As Double
Public DespRealAcumulado As Double
Public DifDesplazamientosAcumulados As Double
Public TripTankManiobra As Double
Public SwabSurge As Double
Public TiempoManiobra As Double
Public DespTeóricoManiobra As Double
Public DespRealManiobra As Double
Public DifDesplazamientos As Double
Public LongHerramientaManiobra As Double
Public LongHerramientaAfueraManiobra As Double
Public ProfHerramientaManiobra As Double
Public NroTiro As Double
Public LongitudTiroEntubado As Double
Public TipoCañoentubado As Double
Public TiempoTiroEntubado As Double
Public PesoLodo100Entubado As Double
Public PesoLodo80Entubado As Double
Public PesoLodo60Entubado As Double
Public PesoRealEntubado As Double
Public PtajePesoEntubado As Double
Public Desplazamiento100Entubado As Double
Public Desplazamiento80Entubado As Double
Public Desplazamiento60Entubado As Double
Public DesplazamientoReal As Double
Public PtajeAcumuladoEntubado As Double
Public TripTankEntubado As Double
Public NroCaño As Double
Public TirosenelPozo As Double
Public TirosenelPeine As Double
Public ExMTotalBomba As Double
Public PresióndeCasing As Double

Public GasExplosivo1 As Double
Public GasExplosivo2 As Double
Public GasExplosivo3 As Double
Public GasExplosivo4 As Double
Public GasExplosivo5 As Double
Public EnergiaEspecificaMax As Double
Public EnergiaEspecificaProm As Double
Public EnergiaEspecificaMin As Double
Public EnergiaEspecificaCorregidaMax As Double
Public EnergiaEspecificaCorregidaProm As Double
Public EnergiaEspecificaCorregidaMin As Double
Public SlidingCoeficientMax As Double
Public SlidingCoeficientProm As Double
Public SlidingCoeficientMin As Double
Public SlidingCoeficientCorregidaMax As Double
Public SlidingCoeficientCorregidaProm As Double
Public SlidingCoeficientCorregidaMin As Double
Public MensajeLitología As String
Public MensajePerforación As String
Public MensajeVarios As String
Public NombrePozo As String
Public CalcimetriaA As Double
Public CalcimetriaB As Double

Public Sub Asignar(Unindice As Integer, Unstr As String)


Select Case Unindice   ' Evalúa el indice.
    Case 1
        ProfundidadPozo = Val(Unstr)
    Case 2
        ProfundidadTrepano = Val(Unstr)
    Case 3
        TVD = Val(Unstr)
    Case 4
        PesoGancho = Val(Unstr)
    Case 5
        PesoAplicado = Val(Unstr)
    Case 6
        Torque = Val(Unstr)
    Case 7
        RPMMesa = Val(Unstr)
    Case 8
        RPMFondo = Val(Unstr)
    Case 9
        PresionBomba = Val(Unstr)
    Case 10
        CaudalEntrada = Val(Unstr)
    Case 11
        CaudalRetorno = Val(Unstr)
    Case 12
        ExMBomba1 = Val(Unstr)
    Case 13
        ExMBomba2 = Val(Unstr)
    Case 14
        ExMBomba3 = Val(Unstr)
    Case 15
        Retorno = Val(Unstr)
    Case 16
        Bajada = Val(Unstr)
    Case 17
        Pileta1 = Val(Unstr)
    Case 18
        Pileta2 = Val(Unstr)
    Case 19
        Pileta3 = Val(Unstr)
    Case 20
        Pileta4 = Val(Unstr)
    Case 21
        Pileta5 = Val(Unstr)
    Case 22
        Pileta6 = Val(Unstr)
    Case 23
        Pileta7 = Val(Unstr)
    Case 24
        Pileta8 = Val(Unstr)
    Case 25
        Pileta9 = Val(Unstr)
    Case 26
        Pileta10 = Val(Unstr)
    Case 27
        TotalNivelPiletas = Val(Unstr)
    Case 28
        DiferenciaNivelPiletas = Val(Unstr)
    Case 29
        InicioDiferencia = Unstr
    Case 30
        GasTotal = Val(Unstr)
        
    Case 31
        AvanceCarrera = Val(Unstr)
    Case 32
        ROPCarrera = Val(Unstr)
    Case 33
        HsMotor = Val(Unstr)
    Case 34
        HsRotacion = Val(Unstr)
    Case 35
        RevolucionesTotales = Val(Unstr)
    Case 36
        ROP = Val(Unstr)
    Case 37
        Cronometraje = Val(Unstr)
    Case 38
        Penetrometro = Val(Unstr)
    Case 39
        TripTank = Val(Unstr)
    Case 40
        Estado = Unstr
    Case 41
        ExpD = Val(Unstr)
    Case 42
        KTF = Val(Unstr)
    Case 43
        ECD = Val(Unstr)
    Case 44
        Dc = Val(Unstr)
    Case 45
        GradientedeFractura = Val(Unstr)
    Case 46
        PresiónPoral = Val(Unstr)
    Case 47
        Densidad = Val(Unstr)
    Case 48
        GBF = Val(Unstr)
    Case 49
        S = Val(Unstr)
    Case 50
        Mu = Val(Unstr)
    Case 51
        TorqueMáximo = Val(Unstr)
    Case 52
        TorqueMínimo = Val(Unstr)
    Case 53
        TorquePromedio = Val(Unstr)
    Case 54
        PesoAplicadoMáximo = Val(Unstr)
    Case 55
        PesoAplicadoMínimo = Val(Unstr)
    Case 56
        PesoAplicadoPromedio = Val(Unstr)
    Case 57
        RPMMáximo = Val(Unstr)
    Case 58
        RPMMínimo = Val(Unstr)
    Case 59
        RPMPromedio = Val(Unstr)
    Case 60
        PresiónBombaMáximo = Val(Unstr)
    Case 61
        PresiónBombaMínimo = Val(Unstr)
    Case 62
        PresiónBombaPromedio = Val(Unstr)
    Case 63
        LongitudTiroManiobra = Val(Unstr)
    Case 64
        VelocidadTiroManiobra = Val(Unstr)
    Case 65
        PesoPromedioManiobra = Val(Unstr)
    Case 66
        PesoMáximoManiobra = Val(Unstr)
    Case 67
        PesoTeóricoManiobra = Val(Unstr)
    Case 68
        DespTeóricoAcumulado = Val(Unstr)
    Case 69
        
        DespRealAcumulado = Val(Unstr)
    Case 70
        DifDesplazamientosAcumulados = Val(Unstr)
    Case 71
        TripTankManiobra = Val(Unstr)
    Case 72
        SwabSurge = Val(Unstr)
    Case 73
        TiempoManiobra = Val(Unstr)
    Case 74
        DespTeóricoManiobra = Val(Unstr)
    Case 75
        DespRealManiobra = Val(Unstr)
    Case 76
        DifDesplazamientos = Val(Unstr)
    Case 77
        LongHerramientaManiobra = Val(Unstr)
    Case 78
        LongHerramientaAfueraManiobra = Val(Unstr)
    Case 79
        ProfHerramientaManiobra = Val(Unstr)
    Case 80
        NroTiro = Val(Unstr)
    Case 81
        LongitudTiroEntubado = Val(Unstr)
    Case 82
        TipoCañoentubado = Val(Unstr)
    Case 83
        TiempoTiroEntubado = Val(Unstr)
    Case 84
        PesoLodo100Entubado = Val(Unstr)
    Case 85
        PesoLodo80Entubado = Val(Unstr)
    Case 86
        PesoLodo60Entubado = Val(Unstr)
    Case 87
        PesoRealEntubado = Val(Unstr)
    Case 88
        PtajePesoEntubado = Val(Unstr)
    Case 89
        Desplazamiento100Entubado = Val(Unstr)
    Case 90
        Desplazamiento80Entubado = Val(Unstr)
    Case 91
        Desplazamiento60Entubado = Val(Unstr)
    Case 92
        DesplazamientoReal = Val(Unstr)
    Case 93
        PtajeAcumuladoEntubado = Val(Unstr)
    Case 94
        TripTankEntubado = Val(Unstr)
    Case 95
        NroCaño = Val(Unstr)
    Case 96
        TirosenelPozo = Val(Unstr)
    Case 97
        TirosenelPeine = Val(Unstr)
    Case 98
        ExMTotalBomba = Val(Unstr)
    Case 99
        PresióndeCasing = Val(Unstr)
    Case 100
    
       
            CO2 = Val(Unstr)
       
        
    Case 101
      
            SH2 = Val(Unstr)
      
        
'    Case 102
'        Sulfhídrico2 = Val(Unstr)
'    Case 103
'        Sulfhídrico3 = Val(Unstr)
'    Case 104
'        Sulfhídrico4 = Val(Unstr)
'    Case 105
'        Sulfhídrico5 = Val(Unstr)
    Case 106
        GasExplosivo1 = Val(Unstr)
    Case 107
        GasExplosivo2 = Val(Unstr)
    Case 108
        GasExplosivo3 = Val(Unstr)
    Case 109
        GasExplosivo4 = Val(Unstr)
    Case 110
        GasExplosivo5 = Val(Unstr)
    Case 111
        EnergiaEspecificaMax = Val(Unstr)
    Case 112
        EnergiaEspecificaProm = Val(Unstr)
    Case 113
        EnergiaEspecificaMin = Val(Unstr)
    Case 114
        EnergiaEspecificaCorregidaMax = Val(Unstr)
    Case 115
        EnergiaEspecificaCorregidaProm = Val(Unstr)
    Case 116
        EnergiaEspecificaCorregidaMin = Val(Unstr)
    Case 117
        SlidingCoeficientMax = Val(Unstr)
    Case 118
        SlidingCoeficientProm = Val(Unstr)
    Case 119
        SlidingCoeficientMin = Val(Unstr)
    Case 120
        SlidingCoeficientCorregidaMax = Val(Unstr)
    Case 121
        SlidingCoeficientCorregidaProm = Val(Unstr)
    Case 122
        SlidingCoeficientCorregidaMin = Val(Unstr)
    Case 123
        MensajeLitología = Unstr
    Case 124
        MensajePerforación = Unstr
    Case 125
        MensajeVarios = Unstr
    Case 126
        NombrePozo = Unstr
    Case 127
        CalcimetriaA = Val(Unstr)
    Case 128
        CalcimetriaB = Val(Unstr)

    Case 150
        ActividadPozo = Unstr
    
    Case 151
        ProfundidadRetorno = Val(Unstr)
        
        If ProfundidadRetorno <> 0 And ProfundidadRetorno <> ProfundidadRetornoAnterior And ProfundidadRetornoAnterior <> 0 Then
               
               DeboAgregarCartel = True
            Else
                DeboAgregarCartel = False
        End If
        
        ProfundidadRetornoAnterior = ProfundidadRetorno
    
    Case 152
        
        If Unstr = "LLEGO" And Not CromaLanzada Then
            
            DeboTirarCroma = True
            
        Else
          '  DeboTirarCroma = False
            
        End If
    
    Case 153
        If DeboTirarCroma Then
            ProfundidadAnalisis = ProfundidadRetorno
        End If
    
    Case 154
        
     '   If Unstr = "LANZAR" And Not CromaLanzada Then
      '
       '     DeboTirarCroma = True
            
       ' Else
          '  DeboTirarCroma = False
            
        'End If
    
    Case 155
        
        If Unstr = "CROMA CORRIENDO" Then
            
            CromaCorriendo = True
            
            
        Else
          CromaCorriendo = False
            
        End If
        
        
    Case Else   ' Otros valores.
'       Debug.Print "No está entre 1 y 10"
'       Debug.Print Unindice
    End Select


End Sub




