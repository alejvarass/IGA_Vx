VERSION 5.00
Begin VB.Form FrmPanelDeGas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gas monitoring System"
   ClientHeight    =   6615
   ClientLeft      =   270
   ClientTop       =   330
   ClientWidth     =   7785
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   7785
   WhatsThisHelp   =   -1  'True
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   9360
      Top             =   6360
   End
End
Attribute VB_Name = "FrmPanelDeGas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
