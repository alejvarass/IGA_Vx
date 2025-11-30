Attribute VB_Name = "ModuloSonido"
Option Explicit




    Public Const SND_ASYNC = &H1     'modo asíncrono. La función retorna una vez iniciada la música (sonido en background).
    Public Const SND_LOOP = &H8      'La música seguirá sonando repetidamente hasta
                                  'que la función sndPlaySound sea llamada de nuevo con un valor nulo para NombreWav (NULL).
        
    Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" _
    (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long


