Attribute VB_Name = "modAutoResize"
Option Explicit

' Colecci?n para almacenar las dimensiones originales de todos los controles
Public g_ResizeCollection As New Collection
' Dimensiones de dise?o originales del formulario
Public g_FormDesignWidth As Single
Public g_FormDesignHeight As Single

' *** BANDERA CR?TICA PARA PREVENIR LA RECURSI?N C?CLICA Y EL DESBORDAMIENTO DE PILA ***
Public g_IsResizing As Boolean


'----------------------------------------------------------------
' 1. Inicializaci?n (Guarda las dimensiones de los controles anidados)
'----------------------------------------------------------------
Public Sub InitializeFormResize(frm As Form)
    
'     g_FormDesignWidth = MdiPrincipal.Width * 1.4
    g_FormDesignWidth = MdiPrincipal.ScaleWidth * 1.47
    
   
'    g_FormDesignHeight = MdiPrincipal.Height * 1.5
    g_FormDesignHeight = MdiPrincipal.ScaleHeight / 0.81
    
    On Error Resume Next
    Set g_ResizeCollection = New Collection
    On Error GoTo 0
    
    ' El formulario es el primer contenedor
    Call InitializeControls(frm, frm)
End Sub


'----------------------------------------------------------------
' 2. Procedimiento Recursivo para GUARDAR propiedades
'----------------------------------------------------------------
Public Sub InitializeControls(frmParent As Form, Container As Object)
    Dim ctrl As Control
    Dim clsOriginal As clsControlOriginal
    Dim sKey As String
    
    Dim sstab_ctrl_local As SSTab
    
    On Error Resume Next

    If TypeOf Container Is SSTab Then
        Set sstab_ctrl_local = Container
        
        ' Bandera de prevenci?n de recursi?n para evitar el stack overflow
        If sstab_ctrl_local.Tag = "INITIALIZING_TAB_LOOP" Then Exit Sub
        sstab_ctrl_local.Tag = "INITIALIZING_TAB_LOOP"
        
        Dim i As Long
        Dim OriginalTab As Long
        OriginalTab = sstab_ctrl_local.Tab
        
        ' 1. ITERAR TODAS LAS PESTA?AS (CR?TICO PARA EXPOSICI?N DE CONTROLES)
        For i = 0 To sstab_ctrl_local.Tabs - 1
            sstab_ctrl_local.Tab = i
            DoEvents ' VITAL: Fuerza a VB a mover los controles a la posici?n (Left >= 0)
            
            ' 2. ITERAR LOS CONTROLES DEL FORMULARIO PADRE Y FILTRAR POR CONTENEDOR
            For Each ctrl In frmParent.Controls
                If ctrl.Container Is sstab_ctrl_local Then
                    
                    ' 3. FILTRO CR?TICO: SOLO CONTROLES VISIBLES (Left >= 0)
                    If ctrl.Left >= 0 Then
                        
                        ' *** CORRECCI?N CR?TICA DE COLISI?N: Key = NombreSSTab_IndicePesta?a_NombreControl ***
                        sKey = sstab_ctrl_local.Name & "_" & CStr(i) & "_" & ctrl.Name
                        
                        g_ResizeCollection.Remove sKey
                        Set clsOriginal = New clsControlOriginal ' clsControlOriginal es necesario
                        
                        clsOriginal.LeftOriginal = ctrl.Left
                        clsOriginal.TopOriginal = ctrl.Top
                        clsOriginal.WidthOriginal = ctrl.Width
                        clsOriginal.HeightOriginal = ctrl.Height
                        
                        g_ResizeCollection.Add clsOriginal, sKey
                        
                        ' 4. RECURSIVIDAD PARA CONTENEDORES ANIDADOS (Frame/PictureBox) DENTRO DE LA PESTA?A
                        If TypeOf ctrl Is Frame Or TypeOf ctrl Is PictureBox Then
                            InitializeControls frmParent, ctrl
                        End If
                    End If
                End If
            Next ctrl
        Next i
        
        ' 5. RESTAURAR Y LIMPIAR
        sstab_ctrl_local.Tab = OriginalTab
        sstab_ctrl_local.Tag = ""
        DoEvents
        
    Else ' Container es el Formulario, un Frame o un PictureBox (No SSTab)
        Dim sContainerName As String
        If TypeOf Container Is Form Then
            sContainerName = Container.Name
        Else
            sContainerName = Container.Name
        End If

        For Each ctrl In Container.Controls
            ' Key: ContainerName_ControlName
            sKey = sContainerName & "_" & ctrl.Name
            
            g_ResizeCollection.Remove sKey
            
            Set clsOriginal = New clsControlOriginal
            
            clsOriginal.LeftOriginal = ctrl.Left
            clsOriginal.TopOriginal = ctrl.Top
            clsOriginal.WidthOriginal = ctrl.Width
            clsOriginal.HeightOriginal = ctrl.Height
            
            g_ResizeCollection.Add clsOriginal, sKey
            
            ' RECURSIVIDAD para SSTab anidado o contenedor anidado
            If TypeOf ctrl Is Frame Or TypeOf ctrl Is SSTab Or TypeOf ctrl Is PictureBox Then
                InitializeControls frmParent, ctrl
            End If
        Next ctrl
    End If
    
    On Error GoTo 0
End Sub


'----------------------------------------------------------------
' 3. Rutina principal para APLICAR el redimensionamiento
'----------------------------------------------------------------
Public Sub ResizeForm(frm As Form)
    
    If g_IsResizing Then Exit Sub
    g_IsResizing = True
    
    Dim ScaleX As Single
    Dim ScaleY As Single
    
    ScaleX = 1#
    ScaleY = 1#
    
    If g_FormDesignWidth > 0 Then ScaleX = frm.ScaleWidth / g_FormDesignWidth
    If g_FormDesignHeight > 0 Then ScaleY = frm.ScaleHeight / g_FormDesignHeight
  
    ResizeControls frm, frm, ScaleX, ScaleY
    
    g_IsResizing = False
End Sub


'----------------------------------------------------------------
' 4. Procedimiento Recursivo para APLICAR el escalado
'----------------------------------------------------------------
Public Sub ResizeControls(frmParent As Form, Container As Object, ScaleX As Single, ScaleY As Single)
    Dim ctrl As Control
    Dim clsOriginal As clsControlOriginal
    Dim sKey As String
    
    Dim sstab_ctrl_local As SSTab

    On Error Resume Next

    If TypeOf Container Is SSTab Then
        Set sstab_ctrl_local = Container

        ' Bandera de prevenci?n de recursi?n
        If sstab_ctrl_local.Tag = "RESIZING_TAB_LOOP" Then Exit Sub
        sstab_ctrl_local.Tag = "RESIZING_TAB_LOOP"
        
        Dim i As Long
        Dim OriginalTab As Long
        OriginalTab = sstab_ctrl_local.Tab
        
        ' 1. ITERAR TODAS LAS PESTA?AS (CR?TICO)
        For i = 0 To sstab_ctrl_local.Tabs - 1
            sstab_ctrl_local.Tab = i
            DoEvents ' VITAL
            
            ' 2. ITERAR LOS CONTROLES DEL FORMULARIO PADRE Y FILTRAR POR CONTENEDOR
            For Each ctrl In frmParent.Controls
                If ctrl.Container Is sstab_ctrl_local Then
                    
                    ' 3. FILTRO CR?TICO: SOLO CONTROLES VISIBLES (Left >= 0)
                    If ctrl.Left >= 0 Then
                        
                        ' *** CORRECCI?N CR?TICA DE COLISI?N: Key = NombreSSTab_IndicePesta?a_NombreControl ***
                        sKey = sstab_ctrl_local.Name & "_" & CStr(i) & "_" & ctrl.Name
                        
                        Set clsOriginal = g_ResizeCollection(sKey)
                        
                        If Not clsOriginal Is Nothing Then
                            ctrl.Left = clsOriginal.LeftOriginal * ScaleX
                            ctrl.Top = clsOriginal.TopOriginal * ScaleY
                            ctrl.Width = clsOriginal.WidthOriginal * ScaleX
                            ctrl.Height = clsOriginal.HeightOriginal * ScaleY
                            
                            ' Forzar Redibujado (ListView, Frame, etc.)
                            If TypeName(ctrl) = "ListView" Or TypeOf ctrl Is Frame Or TypeOf ctrl Is PictureBox Then
                                ctrl.Refresh
                            End If
                        End If
                        
                        ' 4. RECURSIVIDAD PARA CONTENEDORES ANIDADOS
                        If TypeOf ctrl Is Frame Or TypeOf ctrl Is PictureBox Then
                            ResizeControls frmParent, ctrl, ScaleX, ScaleY
                        End If
                    End If
                End If
            Next ctrl
            
        Next i
        
        ' 5. RESTAURAR Y LIMPIAR
        sstab_ctrl_local.Tab = OriginalTab
        sstab_ctrl_local.Tag = ""
        DoEvents
        
    Else ' Container es el Formulario, un Frame o un PictureBox (No SSTab)
        Dim sContainerName As String
        If TypeOf Container Is Form Then
            sContainerName = Container.Name
        Else
            sContainerName = Container.Name
        End If
        
        For Each ctrl In Container.Controls
            ' Key: ContainerName_ControlName
            sKey = sContainerName & "_" & ctrl.Name
            
            Set clsOriginal = g_ResizeCollection(sKey)
            
            If Not clsOriginal Is Nothing Then
                ctrl.Left = clsOriginal.LeftOriginal * ScaleX
                ctrl.Top = clsOriginal.TopOriginal * ScaleY
                ctrl.Width = clsOriginal.WidthOriginal * ScaleX
                ctrl.Height = clsOriginal.HeightOriginal * ScaleY
                
                ' Forzar Redibujado
                If TypeName(ctrl) = "ListView" Or TypeOf ctrl Is Frame Or TypeOf ctrl Is PictureBox Then
                    ctrl.Refresh
                End If
            End If
            
            ' RECURSIVIDAD para SSTab anidado o contenedor anidado
            If TypeOf ctrl Is Frame Or TypeOf ctrl Is SSTab Or TypeOf ctrl Is PictureBox Then
                ResizeControls frmParent, ctrl, ScaleX, ScaleY
            End If
        Next ctrl
    End If
    
    On Error GoTo 0
End Sub

