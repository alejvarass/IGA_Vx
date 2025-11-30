Attribute VB_Name = "ModFunciones"
Option Explicit

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

