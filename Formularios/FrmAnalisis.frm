VERSION 5.00
Object = "{0A362340-2E5E-11D3-85BF-00105AC8B715}#1.0#0"; "isDigitalLibrary.ocx"
Object = "{DA259054-D93B-498C-8C10-DEBD83EF1357}#1.0#0"; "iPlotLibrary.ocx"
Object = "{C5412DA5-2E2F-11D3-85BF-00105AC8B715}#1.0#0"; "isAnalogLibrary.ocx"
Object = "{7D3C5781-B6D7-4C0E-9F2B-18DC0545EA5F}#1.0#0"; "IT_Hora.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form FrmAnalisis 
   BackColor       =   &H00E8E8E8&
   Caption         =   "Analisis"
   ClientHeight    =   9810
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11400
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmAnalisis.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9810
   ScaleWidth      =   11400
   WindowState     =   2  'Maximized
   Begin isAnalogLibrary.iAnalogOutputX DisplaySulfidrico 
      Height          =   675
      Left            =   6375
      TabIndex        =   190
      Top             =   450
      Width           =   1410
      Precision       =   0
      Value           =   0
      ValueMax        =   0
      ValueMin        =   0
      UnitsText       =   ""
      Alignment       =   1
      Color           =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   0
      Modified        =   0   'False
      SelLength       =   1
      SelStart        =   0
      SelText         =   "0"
      Text            =   "0"
      Object.Visible         =   -1  'True
      FontColor       =   65535
      BeepOnError     =   0   'False
      UndoOnError     =   -1  'True
      FilterStyle     =   0
      OptionSaveAllProperties=   0   'False
      AutoSelect      =   0   'False
      AutoSize        =   -1  'True
      Object.Width           =   94
      Object.Height          =   45
      ErrorActive     =   0   'False
      ErrorText       =   "Error"
      BeginProperty ErrorFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ErrorFontColor  =   16777215
      ErrorBackGroundColor=   255
      AlignmentMargin =   0
      AutoFrameRate   =   0   'False
      UpdateFrameRate =   60
      BorderStyle     =   2
      FontName        =   "Arial Narrow"
      FontSize        =   24
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontUnderline   =   0   'False
      FontStrikeOut   =   0   'False
      ErrorFontName   =   "MS Sans Serif"
      ErrorFontSize   =   8
      ErrorFontBold   =   -1  'True
      ErrorFontItalic =   0   'False
      ErrorFontUnderline=   0   'False
      ErrorFontStrikeOut=   0   'False
      OPCItemCount    =   0
   End
   Begin isAnalogLibrary.iAnalogOutputX DisplayDioxidoCarbono 
      Height          =   675
      Left            =   3255
      TabIndex        =   189
      Top             =   450
      Width           =   2235
      Precision       =   0
      Value           =   0
      ValueMax        =   0
      ValueMin        =   0
      UnitsText       =   ""
      Alignment       =   1
      Color           =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   0
      Modified        =   0   'False
      SelLength       =   1
      SelStart        =   0
      SelText         =   "0"
      Text            =   "0"
      Object.Visible         =   -1  'True
      FontColor       =   33023
      BeepOnError     =   0   'False
      UndoOnError     =   -1  'True
      FilterStyle     =   0
      OptionSaveAllProperties=   0   'False
      AutoSelect      =   0   'False
      AutoSize        =   -1  'True
      Object.Width           =   149
      Object.Height          =   45
      ErrorActive     =   0   'False
      ErrorText       =   "Error"
      BeginProperty ErrorFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ErrorFontColor  =   16777215
      ErrorBackGroundColor=   255
      AlignmentMargin =   0
      AutoFrameRate   =   0   'False
      UpdateFrameRate =   60
      BorderStyle     =   2
      FontName        =   "Arial Narrow"
      FontSize        =   24
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontUnderline   =   0   'False
      FontStrikeOut   =   0   'False
      ErrorFontName   =   "MS Sans Serif"
      ErrorFontSize   =   8
      ErrorFontBold   =   -1  'True
      ErrorFontItalic =   0   'False
      ErrorFontUnderline=   0   'False
      ErrorFontStrikeOut=   0   'False
      OPCItemCount    =   0
   End
   Begin isAnalogLibrary.iAnalogOutputX DisplayGasTotal 
      Height          =   675
      Left            =   75
      TabIndex        =   188
      Top             =   450
      Width           =   3000
      Precision       =   0
      Value           =   0
      ValueMax        =   0
      ValueMin        =   0
      UnitsText       =   ""
      Alignment       =   1
      Color           =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   0
      Modified        =   0   'False
      SelLength       =   1
      SelStart        =   0
      SelText         =   "0"
      Text            =   "0"
      Object.Visible         =   -1  'True
      FontColor       =   255
      BeepOnError     =   0   'False
      UndoOnError     =   -1  'True
      FilterStyle     =   0
      OptionSaveAllProperties=   0   'False
      AutoSelect      =   0   'False
      AutoSize        =   -1  'True
      Object.Width           =   200
      Object.Height          =   45
      ErrorActive     =   0   'False
      ErrorText       =   "Error"
      BeginProperty ErrorFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ErrorFontColor  =   16777215
      ErrorBackGroundColor=   255
      AlignmentMargin =   0
      AutoFrameRate   =   0   'False
      UpdateFrameRate =   60
      BorderStyle     =   2
      FontName        =   "Arial Narrow"
      FontSize        =   24
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontUnderline   =   0   'False
      FontStrikeOut   =   0   'False
      ErrorFontName   =   "MS Sans Serif"
      ErrorFontSize   =   8
      ErrorFontBold   =   -1  'True
      ErrorFontItalic =   0   'False
      ErrorFontUnderline=   0   'False
      ErrorFontStrikeOut=   0   'False
      OPCItemCount    =   0
   End
   Begin TabDlg.SSTab SSTab3 
      Height          =   7710
      Left            =   30
      TabIndex        =   6
      Top             =   2055
      Width           =   7785
      _ExtentX        =   13732
      _ExtentY        =   13600
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   15263976
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Analisis Cromatográfico"
      TabPicture(0)   =   "FrmAnalisis.frx":076A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Cronometraje - Gas Total - CO2"
      TabPicture(1)   =   "FrmAnalisis.frx":0786
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame14"
      Tab(1).Control(1)=   "LvwCronoGas"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame14 
         BackColor       =   &H00E8E8E8&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   -74880
         TabIndex        =   157
         Top             =   5985
         Width           =   7515
         Begin VB.Frame fraCmdRefresh 
            Height          =   495
            Left            =   120
            TabIndex        =   167
            Top             =   270
            Width           =   1335
            Begin VB.CommandButton cmdRefresh 
               BackColor       =   &H00E8E8E8&
               Height          =   810
               Left            =   -120
               Picture         =   "FrmAnalisis.frx":07A2
               Style           =   1  'Graphical
               TabIndex        =   168
               Top             =   -120
               Width           =   1575
            End
         End
         Begin VB.Frame fraCmdGuardar 
            Height          =   495
            Left            =   6000
            TabIndex        =   165
            Top             =   270
            Width           =   1335
            Begin VB.CommandButton CmdGuardar 
               BackColor       =   &H00E8E8E8&
               Height          =   810
               Left            =   -120
               Picture         =   "FrmAnalisis.frx":0BBB
               Style           =   1  'Graphical
               TabIndex        =   166
               Top             =   -120
               Width           =   1575
            End
         End
         Begin VB.TextBox TxtCO2 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   4980
            TabIndex        =   18
            Top             =   555
            Width           =   855
         End
         Begin VB.TextBox TxtGasTotal 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2415
            TabIndex        =   16
            Top             =   585
            Width           =   855
         End
         Begin VB.TextBox TxtROP 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4980
            TabIndex        =   17
            Top             =   195
            Width           =   855
         End
         Begin VB.TextBox TxtCrono 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2415
            TabIndex        =   15
            Top             =   195
            Width           =   855
         End
         Begin VB.Label Label67 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CO2"
            Height          =   210
            Left            =   4455
            TabIndex        =   161
            Top             =   660
            Width           =   315
         End
         Begin VB.Label Label66 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Gas Total"
            Height          =   210
            Left            =   1560
            TabIndex        =   160
            Top             =   660
            Width           =   690
         End
         Begin VB.Label Label65 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ROP"
            Height          =   210
            Left            =   4455
            TabIndex        =   159
            Top             =   270
            Width           =   315
         End
         Begin VB.Label Label64 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Crono"
            Height          =   210
            Left            =   1545
            TabIndex        =   158
            Top             =   255
            Width           =   435
         End
      End
      Begin MSComctlLib.ListView LvwCronoGas 
         Height          =   5835
         Left            =   -74880
         TabIndex        =   14
         Top             =   135
         Width           =   7530
         _ExtentX        =   13282
         _ExtentY        =   10292
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList2"
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E8E8E8&
         Caption         =   "Datos Temporales"
         Height          =   3495
         Left            =   60
         TabIndex        =   88
         Top             =   3600
         Width           =   7665
         Begin VB.CheckBox ChkMostrarTodosLosAnálisisTemporal 
            BackColor       =   &H00E8E8E8&
            Caption         =   "Mostrar todos los análisis"
            Height          =   210
            Left            =   45
            TabIndex        =   4
            Top             =   3225
            Width           =   2295
         End
         Begin TabDlg.SSTab SSTab1 
            Height          =   3210
            Left            =   3360
            TabIndex        =   3
            Top             =   210
            Width           =   4245
            _ExtentX        =   7488
            _ExtentY        =   5662
            _Version        =   393216
            Style           =   1
            TabHeight       =   520
            TabCaption(0)   =   "Elementos"
            TabPicture(0)   =   "FrmAnalisis.frx":0FB5
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "ArchivoTemporal"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "Label13"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "Label12"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "Label10"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "Label9"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "Label8"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "LvwTemporal"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "TxtGasTotalCromatograficoTemporal"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).Control(8)=   "TxtGasTotalTemporal"
            Tab(0).Control(8).Enabled=   0   'False
            Tab(0).Control(9)=   "TxtSH2Temporal"
            Tab(0).Control(9).Enabled=   0   'False
            Tab(0).Control(10)=   "TxtCO2Temporal"
            Tab(0).Control(10).Enabled=   0   'False
            Tab(0).Control(11)=   "TxtProfundidadTemporal"
            Tab(0).Control(11).Enabled=   0   'False
            Tab(0).Control(12)=   "fraCmdGuardarTemporal"
            Tab(0).Control(12).Enabled=   0   'False
            Tab(0).ControlCount=   13
            TabCaption(1)   =   "Rel Principales"
            TabPicture(1)   =   "FrmAnalisis.frx":0FD1
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Frame9"
            Tab(1).Control(1)=   "Frame11"
            Tab(1).ControlCount=   2
            TabCaption(2)   =   "Rel Secundarias"
            TabPicture(2)   =   "FrmAnalisis.frx":0FED
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "Frame10"
            Tab(2).Control(1)=   "Frame8"
            Tab(2).Control(2)=   "Frame7"
            Tab(2).ControlCount=   3
            Begin VB.Frame fraCmdGuardarTemporal 
               Height          =   375
               Left            =   3300
               TabIndex        =   169
               Top             =   2760
               Width           =   855
               Begin VB.CommandButton cmdGuardarTemporal 
                  Height          =   690
                  Left            =   -360
                  Picture         =   "FrmAnalisis.frx":1009
                  Style           =   1  'Graphical
                  TabIndex        =   170
                  ToolTipText     =   "Guardar"
                  Top             =   -120
                  Width           =   1575
               End
            End
            Begin VB.TextBox TxtProfundidadTemporal 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   450
               TabIndex        =   136
               Top             =   2070
               Width           =   675
            End
            Begin VB.TextBox TxtCO2Temporal 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   3330
               TabIndex        =   135
               Top             =   2055
               Width           =   855
            End
            Begin VB.TextBox TxtSH2Temporal 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   3720
               TabIndex        =   134
               Top             =   2415
               Width           =   465
            End
            Begin VB.TextBox TxtGasTotalTemporal 
               Alignment       =   1  'Right Justify
               ForeColor       =   &H000000FF&
               Height          =   315
               Left            =   1860
               TabIndex        =   133
               Top             =   2055
               Width           =   765
            End
            Begin VB.TextBox TxtGasTotalCromatograficoTemporal 
               Enabled         =   0   'False
               Height          =   315
               Left            =   1860
               TabIndex        =   132
               Top             =   2430
               Width           =   765
            End
            Begin VB.Frame Frame9 
               Caption         =   "Wtness/Balance/Character"
               Height          =   1035
               Left            =   -74940
               TabIndex        =   125
               Top             =   450
               Width           =   4110
               Begin VB.TextBox TxtCHTemporal 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   3150
                  TabIndex        =   128
                  Top             =   435
                  Width           =   645
               End
               Begin VB.TextBox TxtBHTemporal 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1740
                  TabIndex        =   127
                  Top             =   435
                  Width           =   645
               End
               Begin VB.TextBox TxtWHTemporal 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   495
                  TabIndex        =   126
                  Top             =   435
                  Width           =   645
               End
               Begin VB.Label Label39 
                  AutoSize        =   -1  'True
                  Caption         =   "CH"
                  Height          =   210
                  Left            =   2730
                  TabIndex        =   131
                  Top             =   495
                  Width           =   210
               End
               Begin VB.Label Label40 
                  AutoSize        =   -1  'True
                  Caption         =   "BH"
                  Height          =   210
                  Left            =   1350
                  TabIndex        =   130
                  Top             =   495
                  Width           =   210
               End
               Begin VB.Label Label41 
                  AutoSize        =   -1  'True
                  Caption         =   "WH"
                  Height          =   210
                  Left            =   120
                  TabIndex        =   129
                  Top             =   510
                  Width           =   255
               End
            End
            Begin VB.Frame Frame11 
               Caption         =   "Porcentajes"
               Height          =   1590
               Left            =   -74955
               TabIndex        =   114
               Top             =   1530
               Width           =   4140
               Begin VB.TextBox TxtC1Temporal 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   510
                  TabIndex        =   119
                  Top             =   450
                  Width           =   645
               End
               Begin VB.TextBox TxtC2Temporal 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1740
                  TabIndex        =   118
                  Top             =   450
                  Width           =   645
               End
               Begin VB.TextBox TxtC3Temporal 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   3180
                  TabIndex        =   117
                  Top             =   420
                  Width           =   645
               End
               Begin VB.TextBox TxtC4Temporal 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1140
                  TabIndex        =   116
                  Top             =   1050
                  Width           =   645
               End
               Begin VB.TextBox TxtC5Temporal 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   2475
                  TabIndex        =   115
                  Top             =   1050
                  Width           =   645
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  Caption         =   "C1"
                  Height          =   210
                  Left            =   105
                  TabIndex        =   124
                  Top             =   510
                  Width           =   195
               End
               Begin VB.Label Label14 
                  AutoSize        =   -1  'True
                  Caption         =   "C2"
                  Height          =   210
                  Left            =   1380
                  TabIndex        =   123
                  Top             =   510
                  Width           =   195
               End
               Begin VB.Label Label17 
                  AutoSize        =   -1  'True
                  Caption         =   "C3"
                  Height          =   210
                  Left            =   2790
                  TabIndex        =   122
                  Top             =   510
                  Width           =   195
               End
               Begin VB.Label Label20 
                  AutoSize        =   -1  'True
                  Caption         =   "C4"
                  Height          =   210
                  Left            =   840
                  TabIndex        =   121
                  Top             =   1110
                  Width           =   195
               End
               Begin VB.Label Label53 
                  AutoSize        =   -1  'True
                  Caption         =   "C5"
                  Height          =   210
                  Left            =   2160
                  TabIndex        =   120
                  Top             =   1110
                  Width           =   195
               End
            End
            Begin VB.Frame Frame10 
               Caption         =   "Others"
               Height          =   1590
               Left            =   -74940
               TabIndex        =   103
               Top             =   1530
               Width           =   4050
               Begin VB.TextBox TxtSnGeoTemporal 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1845
                  TabIndex        =   108
                  Top             =   210
                  Width           =   945
               End
               Begin VB.TextBox TxtGeo1Temporal 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1005
                  TabIndex        =   107
                  Top             =   660
                  Width           =   945
               End
               Begin VB.TextBox TxtGeo2Temporal 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   2880
                  TabIndex        =   106
                  Top             =   660
                  Width           =   945
               End
               Begin VB.TextBox TxtGeo3Temporal 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1005
                  TabIndex        =   105
                  Top             =   1170
                  Width           =   945
               End
               Begin VB.TextBox TxtGeo4Temporal 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   2880
                  TabIndex        =   104
                  Top             =   1170
                  Width           =   945
               End
               Begin VB.Label Label51 
                  AutoSize        =   -1  'True
                  Caption         =   "Sn Geo"
                  Height          =   210
                  Left            =   1215
                  TabIndex        =   113
                  Top             =   270
                  Width           =   540
               End
               Begin VB.Label Label50 
                  AutoSize        =   -1  'True
                  Caption         =   "Geo1"
                  Height          =   210
                  Left            =   495
                  TabIndex        =   112
                  Top             =   720
                  Width           =   390
               End
               Begin VB.Label Label49 
                  AutoSize        =   -1  'True
                  Caption         =   "Geo2"
                  Height          =   210
                  Left            =   2355
                  TabIndex        =   111
                  Top             =   720
                  Width           =   390
               End
               Begin VB.Label Label48 
                  AutoSize        =   -1  'True
                  Caption         =   "Geo3"
                  Height          =   210
                  Left            =   495
                  TabIndex        =   110
                  Top             =   1230
                  Width           =   390
               End
               Begin VB.Label Label47 
                  AutoSize        =   -1  'True
                  Caption         =   "Geo4"
                  Height          =   210
                  Left            =   2355
                  TabIndex        =   109
                  Top             =   1230
                  Width           =   390
               End
            End
            Begin VB.Frame Frame8 
               Caption         =   "Coustau"
               Height          =   1065
               Left            =   -72420
               TabIndex        =   98
               Top             =   465
               Width           =   1530
               Begin VB.TextBox TxtCous1Temporal 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   645
                  TabIndex        =   100
                  Top             =   255
                  Width           =   765
               End
               Begin VB.TextBox TxtCous2Temporal 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   630
                  TabIndex        =   99
                  Top             =   645
                  Width           =   765
               End
               Begin VB.Label Label35 
                  AutoSize        =   -1  'True
                  Caption         =   "Cous1"
                  Height          =   210
                  Left            =   135
                  TabIndex        =   102
                  Top             =   315
                  Width           =   465
               End
               Begin VB.Label Label34 
                  AutoSize        =   -1  'True
                  Caption         =   "Cous2"
                  Height          =   210
                  Left            =   105
                  TabIndex        =   101
                  Top             =   705
                  Width           =   465
               End
            End
            Begin VB.Frame Frame7 
               Caption         =   "Baroid (Pixler)"
               Height          =   1035
               Left            =   -74940
               TabIndex        =   89
               Top             =   480
               Width           =   2400
               Begin VB.TextBox TxtBar2Temporal 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   435
                  TabIndex        =   93
                  Top             =   255
                  Width           =   645
               End
               Begin VB.TextBox TxtBar3Temporal 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1545
                  TabIndex        =   92
                  Top             =   255
                  Width           =   645
               End
               Begin VB.TextBox TxtBar4Temporal 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   435
                  TabIndex        =   91
                  Top             =   630
                  Width           =   630
               End
               Begin VB.TextBox TxtBar5Temporal 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1545
                  TabIndex        =   90
                  Top             =   630
                  Width           =   645
               End
               Begin VB.Label Label31 
                  AutoSize        =   -1  'True
                  Caption         =   "Bar2"
                  Height          =   210
                  Left            =   45
                  TabIndex        =   97
                  Top             =   315
                  Width           =   345
               End
               Begin VB.Label Label30 
                  AutoSize        =   -1  'True
                  Caption         =   "Bar3"
                  Height          =   210
                  Left            =   1125
                  TabIndex        =   96
                  Top             =   315
                  Width           =   345
               End
               Begin VB.Label Label29 
                  AutoSize        =   -1  'True
                  Caption         =   "Bar4"
                  Height          =   210
                  Left            =   45
                  TabIndex        =   95
                  Top             =   705
                  Width           =   345
               End
               Begin VB.Label Label28 
                  AutoSize        =   -1  'True
                  Caption         =   "Bar5"
                  Height          =   210
                  Left            =   1125
                  TabIndex        =   94
                  Top             =   705
                  Width           =   345
               End
            End
            Begin MSComctlLib.ListView LvwTemporal 
               Height          =   1710
               Left            =   45
               TabIndex        =   137
               Top             =   330
               Width           =   4155
               _ExtentX        =   7329
               _ExtentY        =   3016
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   0   'False
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   0
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Prof"
               Height          =   210
               Left            =   135
               TabIndex        =   144
               Top             =   2130
               Width           =   300
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "CO2"
               Height          =   210
               Left            =   2925
               TabIndex        =   143
               Top             =   2100
               Width           =   315
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "SH2"
               Height          =   210
               Left            =   2940
               TabIndex        =   142
               Top             =   2460
               Width           =   300
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "THA"
               Height          =   210
               Left            =   1485
               TabIndex        =   141
               Top             =   2100
               Width           =   315
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               Caption         =   "Gas Total Crom."
               Height          =   210
               Left            =   705
               TabIndex        =   140
               Top             =   2475
               Width           =   1155
            End
            Begin VB.Label Label4 
               Caption         =   "Archivo Cromatografia Temporal.Add , ""Archivo"", ""Prof."", 1550"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   -90
               TabIndex        =   139
               Top             =   -270
               Width           =   3555
            End
            Begin VB.Label ArchivoTemporal 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   180
               TabIndex        =   138
               Top             =   2940
               Width           =   2955
            End
         End
         Begin MSComctlLib.ListView LvwAnalisisTemporal 
            Height          =   2985
            Left            =   30
            TabIndex        =   2
            Top             =   225
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   5265
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            _Version        =   393217
            SmallIcons      =   "ImageList1"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E8E8E8&
         Caption         =   "Datos Validados"
         Height          =   3495
         Left            =   60
         TabIndex        =   32
         Top             =   90
         Width           =   7635
         Begin MSComctlLib.ListView LvwAnalisisDefinitivo 
            Height          =   3210
            Left            =   15
            TabIndex        =   0
            Top             =   225
            Width           =   3285
            _ExtentX        =   5794
            _ExtentY        =   5662
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            _Version        =   393217
            SmallIcons      =   "ImageList1"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin TabDlg.SSTab SSTab2 
            Height          =   3210
            Left            =   3360
            TabIndex        =   1
            Top             =   210
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   5662
            _Version        =   393216
            Style           =   1
            TabHeight       =   520
            TabCaption(0)   =   "Elementos"
            TabPicture(0)   =   "FrmAnalisis.frx":1247
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "ArchivoDefinitivo"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "Label23"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "Label22"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "Label21"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "Label19"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "Label18"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "LvwDefinitivo"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "TxtProfundidadDefinitivo"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).Control(8)=   "TxtCO2Definitivo"
            Tab(0).Control(8).Enabled=   0   'False
            Tab(0).Control(9)=   "TxtSH2Definitivo"
            Tab(0).Control(9).Enabled=   0   'False
            Tab(0).Control(10)=   "TxtGasTotalDefinitivo"
            Tab(0).Control(10).Enabled=   0   'False
            Tab(0).Control(11)=   "TxtGasTotalCromatograficoDefinitivo"
            Tab(0).Control(11).Enabled=   0   'False
            Tab(0).ControlCount=   12
            TabCaption(1)   =   "Rel Principales"
            TabPicture(1)   =   "FrmAnalisis.frx":1263
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Frame5"
            Tab(1).Control(1)=   "Frame12"
            Tab(1).ControlCount=   2
            TabCaption(2)   =   "Rel Secundarias"
            TabPicture(2)   =   "FrmAnalisis.frx":127F
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "Frame6"
            Tab(2).Control(1)=   "Frame4"
            Tab(2).Control(2)=   "Frame3"
            Tab(2).ControlCount=   3
            Begin VB.TextBox TxtGasTotalCromatograficoDefinitivo 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   315
               Left            =   1770
               TabIndex        =   80
               Top             =   2580
               Width           =   735
            End
            Begin VB.TextBox TxtGasTotalDefinitivo 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   315
               Left            =   1770
               TabIndex        =   79
               Top             =   2190
               Width           =   735
            End
            Begin VB.TextBox TxtSH2Definitivo 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   315
               Left            =   3660
               TabIndex        =   78
               Top             =   2580
               Width           =   465
            End
            Begin VB.TextBox TxtCO2Definitivo 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   315
               Left            =   3360
               TabIndex        =   77
               Top             =   2190
               Width           =   765
            End
            Begin VB.TextBox TxtProfundidadDefinitivo 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   315
               Left            =   420
               TabIndex        =   76
               Top             =   2190
               Width           =   615
            End
            Begin VB.Frame Frame5 
               Caption         =   "Wtness/Balance/Character"
               Height          =   1035
               Left            =   -74955
               TabIndex        =   69
               Top             =   465
               Width           =   4050
               Begin VB.TextBox TxtWHDefinitivo 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   405
                  TabIndex        =   72
                  Top             =   435
                  Width           =   645
               End
               Begin VB.TextBox TxtBHDefinitivo 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1740
                  TabIndex        =   71
                  Top             =   435
                  Width           =   645
               End
               Begin VB.TextBox TxtCHDefinitivo 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   3060
                  TabIndex        =   70
                  Top             =   435
                  Width           =   705
               End
               Begin VB.Label Label36 
                  Caption         =   "WH"
                  Height          =   195
                  Left            =   75
                  TabIndex        =   75
                  Top             =   495
                  Width           =   285
               End
               Begin VB.Label Label37 
                  AutoSize        =   -1  'True
                  Caption         =   "BH"
                  Height          =   210
                  Left            =   1380
                  TabIndex        =   74
                  Top             =   495
                  Width           =   210
               End
               Begin VB.Label Label38 
                  AutoSize        =   -1  'True
                  Caption         =   "CH"
                  Height          =   210
                  Left            =   2700
                  TabIndex        =   73
                  Top             =   480
                  Width           =   210
               End
            End
            Begin VB.Frame Frame6 
               Caption         =   "Others"
               Height          =   1590
               Left            =   -74850
               TabIndex        =   58
               Top             =   1500
               Width           =   3900
               Begin VB.TextBox TxtGeo4Definitivo 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   2790
                  TabIndex        =   63
                  Top             =   1170
                  Width           =   945
               End
               Begin VB.TextBox TxtGeo3Definitivo 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   555
                  TabIndex        =   62
                  Top             =   1170
                  Width           =   945
               End
               Begin VB.TextBox TxtGeo2Definitivo 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   2790
                  TabIndex        =   61
                  Top             =   660
                  Width           =   945
               End
               Begin VB.TextBox TxtGeo1Definitivo 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   555
                  TabIndex        =   60
                  Top             =   660
                  Width           =   945
               End
               Begin VB.TextBox TxtSnGeoDefinitivo 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1755
                  TabIndex        =   59
                  Top             =   210
                  Width           =   945
               End
               Begin VB.Label Label46 
                  AutoSize        =   -1  'True
                  Caption         =   "Geo4"
                  Height          =   210
                  Left            =   2265
                  TabIndex        =   68
                  Top             =   1230
                  Width           =   390
               End
               Begin VB.Label Label45 
                  AutoSize        =   -1  'True
                  Caption         =   "Geo3"
                  Height          =   210
                  Left            =   45
                  TabIndex        =   67
                  Top             =   1230
                  Width           =   390
               End
               Begin VB.Label Label44 
                  AutoSize        =   -1  'True
                  Caption         =   "Geo2"
                  Height          =   210
                  Left            =   2265
                  TabIndex        =   66
                  Top             =   720
                  Width           =   390
               End
               Begin VB.Label Label43 
                  AutoSize        =   -1  'True
                  Caption         =   "Geo1"
                  Height          =   210
                  Left            =   45
                  TabIndex        =   65
                  Top             =   720
                  Width           =   390
               End
               Begin VB.Label Label42 
                  AutoSize        =   -1  'True
                  Caption         =   "Sn Geo"
                  Height          =   210
                  Left            =   1125
                  TabIndex        =   64
                  Top             =   270
                  Width           =   540
               End
            End
            Begin VB.Frame Frame4 
               Caption         =   "Coustau"
               Height          =   1065
               Left            =   -72450
               TabIndex        =   53
               Top             =   435
               Width           =   1500
               Begin VB.TextBox TxtCous2Definitivo 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   600
                  TabIndex        =   55
                  Top             =   645
                  Width           =   765
               End
               Begin VB.TextBox TxtCous1Definitivo 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   585
                  TabIndex        =   54
                  Top             =   255
                  Width           =   765
               End
               Begin VB.Label Label33 
                  AutoSize        =   -1  'True
                  Caption         =   "Cous2"
                  Height          =   210
                  Left            =   75
                  TabIndex        =   57
                  Top             =   705
                  Width           =   465
               End
               Begin VB.Label Label32 
                  AutoSize        =   -1  'True
                  Caption         =   "Cous1"
                  Height          =   210
                  Left            =   75
                  TabIndex        =   56
                  Top             =   315
                  Width           =   465
               End
            End
            Begin VB.Frame Frame3 
               Caption         =   "Baroid (Pixler)"
               Height          =   1035
               Left            =   -74865
               TabIndex        =   44
               Top             =   450
               Width           =   2280
               Begin VB.TextBox TxtBar5Definitivo 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1485
                  TabIndex        =   48
                  Top             =   630
                  Width           =   555
               End
               Begin VB.TextBox TxtBar4Definitivo 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   465
                  TabIndex        =   47
                  Top             =   630
                  Width           =   480
               End
               Begin VB.TextBox TxtBar3Definitivo 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1485
                  TabIndex        =   46
                  Top             =   255
                  Width           =   555
               End
               Begin VB.TextBox TxtBar2Definitivo 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   300
                  Left            =   465
                  TabIndex        =   45
                  Top             =   255
                  Width           =   495
               End
               Begin VB.Label Label27 
                  AutoSize        =   -1  'True
                  Caption         =   "Bar5"
                  Height          =   210
                  Left            =   1065
                  TabIndex        =   52
                  Top             =   705
                  Width           =   345
               End
               Begin VB.Label Label26 
                  AutoSize        =   -1  'True
                  Caption         =   "Bar4"
                  Height          =   210
                  Left            =   75
                  TabIndex        =   51
                  Top             =   705
                  Width           =   345
               End
               Begin VB.Label Label25 
                  AutoSize        =   -1  'True
                  Caption         =   "Bar3"
                  Height          =   210
                  Left            =   1065
                  TabIndex        =   50
                  Top             =   315
                  Width           =   345
               End
               Begin VB.Label Label24 
                  AutoSize        =   -1  'True
                  Caption         =   "Bar2"
                  Height          =   210
                  Left            =   75
                  TabIndex        =   49
                  Top             =   315
                  Width           =   345
               End
            End
            Begin VB.Frame Frame12 
               Caption         =   "Porcentajes"
               Height          =   1590
               Left            =   -74955
               TabIndex        =   33
               Top             =   1530
               Width           =   4080
               Begin VB.TextBox TxtC5Definitivo 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   2385
                  TabIndex        =   38
                  Top             =   1050
                  Width           =   645
               End
               Begin VB.TextBox TxtC4Definitivo 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1170
                  TabIndex        =   37
                  Top             =   1050
                  Width           =   645
               End
               Begin VB.TextBox TxtC3Definitivo 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   3030
                  TabIndex        =   36
                  Top             =   420
                  Width           =   645
               End
               Begin VB.TextBox TxtC2Definitivo 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1710
                  TabIndex        =   35
                  Top             =   450
                  Width           =   645
               End
               Begin VB.TextBox TxtC1Definitivo 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   420
                  TabIndex        =   34
                  Top             =   450
                  Width           =   645
               End
               Begin VB.Label Label54 
                  AutoSize        =   -1  'True
                  Caption         =   "C5"
                  Height          =   210
                  Left            =   2130
                  TabIndex        =   43
                  Top             =   1110
                  Width           =   195
               End
               Begin VB.Label Label55 
                  AutoSize        =   -1  'True
                  Caption         =   "C4"
                  Height          =   210
                  Left            =   870
                  TabIndex        =   42
                  Top             =   1110
                  Width           =   195
               End
               Begin VB.Label Label56 
                  AutoSize        =   -1  'True
                  Caption         =   "C3"
                  Height          =   210
                  Left            =   2760
                  TabIndex        =   41
                  Top             =   510
                  Width           =   195
               End
               Begin VB.Label Label57 
                  AutoSize        =   -1  'True
                  Caption         =   "C2"
                  Height          =   210
                  Left            =   1440
                  TabIndex        =   40
                  Top             =   510
                  Width           =   195
               End
               Begin VB.Label Label58 
                  AutoSize        =   -1  'True
                  Caption         =   "C1"
                  Height          =   210
                  Left            =   165
                  TabIndex        =   39
                  Top             =   510
                  Width           =   195
               End
            End
            Begin MSComctlLib.ListView LvwDefinitivo 
               Height          =   1830
               Left            =   45
               TabIndex        =   81
               Top             =   330
               Width           =   4125
               _ExtentX        =   7276
               _ExtentY        =   3228
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   0   'False
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   0
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               Caption         =   "Gas Total Crom"
               Height          =   210
               Left            =   555
               TabIndex        =   87
               Top             =   2655
               Width           =   1110
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               Caption         =   "THA"
               Height          =   210
               Left            =   1335
               TabIndex        =   86
               Top             =   2265
               Width           =   315
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               Caption         =   "SH2"
               Height          =   210
               Left            =   3000
               TabIndex        =   85
               Top             =   2625
               Width           =   300
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               Caption         =   "CO2"
               Height          =   210
               Left            =   2985
               TabIndex        =   84
               Top             =   2265
               Width           =   315
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               Caption         =   "Prof"
               Height          =   210
               Left            =   75
               TabIndex        =   83
               Top             =   2265
               Width           =   300
            End
            Begin VB.Label ArchivoDefinitivo 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   90
               TabIndex        =   82
               Top             =   2910
               Width           =   3795
            End
         End
      End
      Begin VB.Label Label60 
         Caption         =   "HAIRDRESSER"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   3750
         TabIndex        =   145
         Top             =   -1530
         Width           =   1755
      End
   End
   Begin VB.Timer Timer2 
      Left            =   7035
      Top             =   6045
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   11100
      Top             =   10620
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   540
      Top             =   11025
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAnalisis.frx":129B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAnalisis.frx":15B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAnalisis.frx":6DA7
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAnalisis.frx":72F9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame13 
      BackColor       =   &H00E8E8E8&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   60
      TabIndex        =   150
      Top             =   1020
      Width           =   7755
      Begin isDigitalLibrary.iSwitchSliderX SwTha 
         Height          =   735
         Left            =   570
         TabIndex        =   8
         Top             =   180
         Width           =   1845
         EndsMargin      =   10
         PointerIndicatorActiveColor=   255
         PointerIndicatorInactiveColor=   0
         Orientation     =   1
         OrientationLabels=   1
         PointerHeight   =   8
         PointerStyle    =   3
         PointerWidth    =   13
         TrackColor      =   16777215
         TrackStyle      =   0
         BackGroundColor =   15263976
         Position        =   0
         PositionLabels  =   "x 1, x 10, x 100"
         PositionLabelMargin=   5
         BeginProperty PositionLabelInactiveFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty PositionLabelActiveFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PositionIndicatorSize=   3
         PositionIndicatorMargin=   5
         PositionIndicatorColor=   4194368
         ShowPositionIndicators=   -1  'True
         ShowPositionLabels=   -1  'True
         PositionIndicatorBevelStyle=   2
         ShowFocusRect   =   -1  'True
         KeyArrowStepSize=   1
         KeyPageStepSize =   2
         BorderStyle     =   0
         Object.Visible         =   -1  'True
         Enabled         =   -1  'True
         PositionLabelActiveFontColor=   0
         PositionLabelInactiveFontColor=   -2147483640
         BackGroundPicture=   "FrmAnalisis.frx":740A
         PositionIndicatorStyle=   0
         Transparent     =   0   'False
         UpdateFrameRate =   60
         OptionSaveAllProperties=   0   'False
         AutoFrameRate   =   0   'False
         Object.Width           =   123
         Object.Height          =   49
         AutoCenter      =   0   'False
         OffsetX         =   0
         OffsetY         =   0
         PointerColor    =   -2147483633
         PointerHighLightColor=   -2147483628
         PointerBitmap   =   "FrmAnalisis.frx":7460
         PositionLabelActiveFontName=   "MS Sans Serif"
         PositionLabelActiveFontSize=   8
         PositionLabelActiveFontBold=   -1  'True
         PositionLabelActiveFontItalic=   0   'False
         PositionLabelActiveFontUnderline=   0   'False
         PositionLabelActiveFontStrikeOut=   0   'False
         PositionLabelInactiveFontName=   "MS Sans Serif"
         PositionLabelInactiveFontSize=   8
         PositionLabelInactiveFontBold=   0   'False
         PositionLabelInactiveFontItalic=   0   'False
         PositionLabelInactiveFontUnderline=   0   'False
         PositionLabelInactiveFontStrikeOut=   0   'False
         OPCItemCount    =   0
      End
      Begin isDigitalLibrary.iLedRoundX iLedFast 
         Height          =   225
         Left            =   5310
         TabIndex        =   154
         Top             =   540
         Width           =   225
         BackGroundColor =   -2147483633
         Active          =   -1  'True
         ActiveColor     =   255
         BevelStyle      =   2
         BorderStyle     =   0
         Object.Visible         =   -1  'True
         Enabled         =   -1  'True
         ShowReflection  =   -1  'True
         AutoInactiveColor=   -1  'True
         InactiveColor   =   21760
         Transparent     =   0   'False
         UpdateFrameRate =   60
         OptionSaveAllProperties=   0   'False
         AutoFrameRate   =   0   'False
         Object.Width           =   15
         Object.Height          =   15
         OPCItemCount    =   0
      End
      Begin isDigitalLibrary.iLedRoundX iLedTha 
         Height          =   225
         Left            =   150
         TabIndex        =   155
         Top             =   540
         Width           =   225
         BackGroundColor =   -2147483633
         Active          =   -1  'True
         ActiveColor     =   57088
         BevelStyle      =   2
         BorderStyle     =   0
         Object.Visible         =   -1  'True
         Enabled         =   -1  'True
         ShowReflection  =   -1  'True
         AutoInactiveColor=   -1  'True
         InactiveColor   =   21760
         Transparent     =   0   'False
         UpdateFrameRate =   60
         OptionSaveAllProperties=   0   'False
         AutoFrameRate   =   0   'False
         Object.Width           =   15
         Object.Height          =   15
         OPCItemCount    =   0
      End
      Begin isDigitalLibrary.iLedRoundX iLedNormal 
         Height          =   225
         Left            =   2820
         TabIndex        =   156
         Top             =   510
         Width           =   225
         BackGroundColor =   -2147483633
         Active          =   -1  'True
         ActiveColor     =   65535
         BevelStyle      =   2
         BorderStyle     =   0
         Object.Visible         =   -1  'True
         Enabled         =   -1  'True
         ShowReflection  =   -1  'True
         AutoInactiveColor=   -1  'True
         InactiveColor   =   4227200
         Transparent     =   0   'False
         UpdateFrameRate =   60
         OptionSaveAllProperties=   0   'False
         AutoFrameRate   =   0   'False
         Object.Width           =   15
         Object.Height          =   15
         OPCItemCount    =   0
      End
      Begin isDigitalLibrary.iSwitchSliderX SwNormal 
         Height          =   735
         Left            =   3240
         TabIndex        =   10
         Top             =   180
         Width           =   1845
         EndsMargin      =   10
         PointerIndicatorActiveColor=   255
         PointerIndicatorInactiveColor=   0
         Orientation     =   1
         OrientationLabels=   1
         PointerHeight   =   8
         PointerStyle    =   3
         PointerWidth    =   13
         TrackColor      =   16777215
         TrackStyle      =   0
         BackGroundColor =   15263976
         Position        =   0
         PositionLabels  =   "x 1, x 10, x 100"
         PositionLabelMargin=   5
         BeginProperty PositionLabelInactiveFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty PositionLabelActiveFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PositionIndicatorSize=   3
         PositionIndicatorMargin=   5
         PositionIndicatorColor=   4194368
         ShowPositionIndicators=   -1  'True
         ShowPositionLabels=   -1  'True
         PositionIndicatorBevelStyle=   2
         ShowFocusRect   =   -1  'True
         KeyArrowStepSize=   1
         KeyPageStepSize =   2
         BorderStyle     =   0
         Object.Visible         =   -1  'True
         Enabled         =   -1  'True
         PositionLabelActiveFontColor=   0
         PositionLabelInactiveFontColor=   -2147483640
         BackGroundPicture=   "FrmAnalisis.frx":74B6
         PositionIndicatorStyle=   0
         Transparent     =   0   'False
         UpdateFrameRate =   60
         OptionSaveAllProperties=   0   'False
         AutoFrameRate   =   0   'False
         Object.Width           =   123
         Object.Height          =   49
         AutoCenter      =   0   'False
         OffsetX         =   0
         OffsetY         =   0
         PointerColor    =   -2147483633
         PointerHighLightColor=   -2147483628
         PointerBitmap   =   "FrmAnalisis.frx":750C
         PositionLabelActiveFontName=   "MS Sans Serif"
         PositionLabelActiveFontSize=   8
         PositionLabelActiveFontBold=   -1  'True
         PositionLabelActiveFontItalic=   0   'False
         PositionLabelActiveFontUnderline=   0   'False
         PositionLabelActiveFontStrikeOut=   0   'False
         PositionLabelInactiveFontName=   "MS Sans Serif"
         PositionLabelInactiveFontSize=   8
         PositionLabelInactiveFontBold=   0   'False
         PositionLabelInactiveFontItalic=   0   'False
         PositionLabelInactiveFontUnderline=   0   'False
         PositionLabelInactiveFontStrikeOut=   0   'False
         OPCItemCount    =   0
      End
      Begin isDigitalLibrary.iSwitchSliderX SwFast 
         Height          =   735
         Left            =   5730
         TabIndex        =   12
         Top             =   180
         Width           =   1845
         EndsMargin      =   10
         PointerIndicatorActiveColor=   255
         PointerIndicatorInactiveColor=   0
         Orientation     =   1
         OrientationLabels=   1
         PointerHeight   =   8
         PointerStyle    =   3
         PointerWidth    =   13
         TrackColor      =   16777215
         TrackStyle      =   0
         BackGroundColor =   15263976
         Position        =   0
         PositionLabels  =   "x 1, x 10, x 100"
         PositionLabelMargin=   5
         BeginProperty PositionLabelInactiveFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty PositionLabelActiveFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PositionIndicatorSize=   3
         PositionIndicatorMargin=   5
         PositionIndicatorColor=   0
         ShowPositionIndicators=   -1  'True
         ShowPositionLabels=   -1  'True
         PositionIndicatorBevelStyle=   2
         ShowFocusRect   =   -1  'True
         KeyArrowStepSize=   1
         KeyPageStepSize =   2
         BorderStyle     =   0
         Object.Visible         =   -1  'True
         Enabled         =   -1  'True
         PositionLabelActiveFontColor=   0
         PositionLabelInactiveFontColor=   -2147483640
         BackGroundPicture=   "FrmAnalisis.frx":7562
         PositionIndicatorStyle=   0
         Transparent     =   0   'False
         UpdateFrameRate =   60
         OptionSaveAllProperties=   0   'False
         AutoFrameRate   =   0   'False
         Object.Width           =   123
         Object.Height          =   49
         AutoCenter      =   0   'False
         OffsetX         =   0
         OffsetY         =   0
         PointerColor    =   -2147483633
         PointerHighLightColor=   -2147483628
         PointerBitmap   =   "FrmAnalisis.frx":75B8
         PositionLabelActiveFontName=   "MS Sans Serif"
         PositionLabelActiveFontSize=   8
         PositionLabelActiveFontBold=   -1  'True
         PositionLabelActiveFontItalic=   0   'False
         PositionLabelActiveFontUnderline=   0   'False
         PositionLabelActiveFontStrikeOut=   0   'False
         PositionLabelInactiveFontName=   "MS Sans Serif"
         PositionLabelInactiveFontSize=   8
         PositionLabelInactiveFontBold=   0   'False
         PositionLabelInactiveFontItalic=   0   'False
         PositionLabelInactiveFontUnderline=   0   'False
         PositionLabelInactiveFontStrikeOut=   0   'False
         OPCItemCount    =   0
      End
      Begin VB.Label Label63 
         BackStyle       =   0  'Transparent
         Caption         =   "F A S T"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   5580
         TabIndex        =   153
         Top             =   150
         Width           =   105
      End
      Begin VB.Label Label62 
         BackStyle       =   0  'Transparent
         Caption         =   "N O R M"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   3090
         TabIndex        =   152
         Top             =   150
         Width           =   135
      End
      Begin VB.Label Label61 
         BackStyle       =   0  'Transparent
         Caption         =   "T H A"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   450
         TabIndex        =   151
         Top             =   360
         Width           =   105
      End
   End
   Begin TabDlg.SSTab SSTab4 
      Height          =   9915
      Left            =   7830
      TabIndex        =   24
      Top             =   -150
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   17489
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   15263976
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Tiempo"
      TabPicture(0)   =   "FrmAnalisis.frx":760E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblFecha"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblHora"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label70"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "iPlotGases"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "TxtComentario"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "CmdGuardarComentario"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Txtfecha"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Txthora"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "fraBtnPrint"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "fraCmdConsultar"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "fraCmdPlay"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "fraCmdStop"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "fraUpArrow"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "fraDownArrow"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "Profundidad"
      TabPicture(1)   =   "FrmAnalisis.frx":762A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraDownArrowProf"
      Tab(1).Control(1)=   "fraUpArrowProf"
      Tab(1).Control(2)=   "fraCmdActualizar"
      Tab(1).Control(3)=   "TxtProfundidadInicio"
      Tab(1).Control(4)=   "TxtSpan"
      Tab(1).Control(5)=   "iPlotMasterLog"
      Tab(1).Control(6)=   "Label16"
      Tab(1).Control(7)=   "Label15"
      Tab(1).ControlCount=   8
      Begin VB.Frame fraDownArrowProf 
         BackColor       =   &H00000000&
         Height          =   420
         Left            =   -68160
         TabIndex        =   185
         Top             =   7320
         Width           =   615
         Begin VB.CommandButton cmdDownArrowProf 
            BackColor       =   &H00000000&
            DisabledPicture =   "FrmAnalisis.frx":7646
            Height          =   690
            Left            =   -480
            Picture         =   "FrmAnalisis.frx":77F3
            Style           =   1  'Graphical
            TabIndex        =   186
            ToolTipText     =   "Down"
            Top             =   -120
            Width           =   1575
         End
      End
      Begin VB.Frame fraUpArrowProf 
         BackColor       =   &H00000000&
         Height          =   420
         Left            =   -68160
         TabIndex        =   183
         Top             =   240
         Width           =   615
         Begin VB.CommandButton cmdUpArrowProf 
            BackColor       =   &H00000000&
            DisabledPicture =   "FrmAnalisis.frx":799B
            Height          =   690
            Left            =   -480
            Picture         =   "FrmAnalisis.frx":7B45
            Style           =   1  'Graphical
            TabIndex        =   184
            ToolTipText     =   "Up"
            Top             =   -120
            Width           =   1575
         End
      End
      Begin VB.Frame fraDownArrow 
         BackColor       =   &H00000000&
         Height          =   420
         Left            =   6900
         TabIndex        =   181
         Top             =   6885
         Width           =   615
         Begin VB.CommandButton cmdDownArrow 
            BackColor       =   &H00000000&
            Height          =   690
            Left            =   -510
            Picture         =   "FrmAnalisis.frx":7CF7
            Style           =   1  'Graphical
            TabIndex        =   182
            ToolTipText     =   "Down"
            Top             =   -120
            Width           =   1575
         End
      End
      Begin VB.Frame fraUpArrow 
         BackColor       =   &H00000000&
         Height          =   420
         Left            =   6900
         TabIndex        =   180
         Top             =   285
         Width           =   615
         Begin VB.CommandButton cmdUpArrow 
            BackColor       =   &H00000000&
            Height          =   690
            Left            =   -450
            Picture         =   "FrmAnalisis.frx":7E9F
            Style           =   1  'Graphical
            TabIndex        =   187
            ToolTipText     =   "Up"
            Top             =   -120
            Width           =   1515
         End
      End
      Begin VB.Frame fraCmdStop 
         Height          =   375
         Left            =   600
         TabIndex        =   178
         Top             =   8320
         Width           =   375
         Begin VB.CommandButton cmdStop 
            DisabledPicture =   "FrmAnalisis.frx":8051
            Height          =   690
            Left            =   -600
            Picture         =   "FrmAnalisis.frx":81B5
            Style           =   1  'Graphical
            TabIndex        =   179
            ToolTipText     =   "Stop"
            Top             =   -120
            Width           =   1575
         End
      End
      Begin VB.Frame fraCmdPlay 
         Height          =   375
         Left            =   120
         TabIndex        =   176
         Top             =   8320
         Width           =   375
         Begin VB.CommandButton cmdPlay 
            DisabledPicture =   "FrmAnalisis.frx":82F4
            Height          =   690
            Left            =   -600
            Picture         =   "FrmAnalisis.frx":844F
            Style           =   1  'Graphical
            TabIndex        =   177
            ToolTipText     =   "Play"
            Top             =   -120
            Width           =   1575
         End
      End
      Begin VB.Frame fraCmdActualizar 
         Height          =   495
         Left            =   -70440
         TabIndex        =   175
         Top             =   8715
         Width           =   1455
         Begin VB.CommandButton cmdActualizar 
            DisabledPicture =   "FrmAnalisis.frx":85D4
            Height          =   810
            Left            =   -240
            Picture         =   "FrmAnalisis.frx":8A3C
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Actualizar"
            Top             =   -120
            Width           =   1935
         End
      End
      Begin VB.Frame fraCmdConsultar 
         Height          =   495
         Left            =   5760
         TabIndex        =   173
         Top             =   8280
         Width           =   1455
         Begin VB.CommandButton cmdConsultar 
            DisabledPicture =   "FrmAnalisis.frx":8E91
            Height          =   810
            Left            =   -120
            Picture         =   "FrmAnalisis.frx":92E8
            Style           =   1  'Graphical
            TabIndex        =   174
            ToolTipText     =   "Consultar"
            Top             =   -120
            Width           =   1815
         End
      End
      Begin VB.Frame fraBtnPrint 
         Height          =   495
         Left            =   1200
         TabIndex        =   171
         Top             =   8280
         Width           =   1095
         Begin VB.CommandButton BtnPrint 
            Height          =   810
            Left            =   -240
            Picture         =   "FrmAnalisis.frx":9737
            Style           =   1  'Graphical
            TabIndex        =   172
            ToolTipText     =   "Print"
            Top             =   -120
            Width           =   1575
         End
      End
      Begin ITHora.IT_Hora Txthora 
         Height          =   255
         Left            =   4920
         TabIndex        =   20
         Top             =   8385
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Text            =   ""
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.DTPicker Txtfecha 
         Height          =   255
         Left            =   3120
         TabIndex        =   19
         Top             =   8385
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   81264641
         CurrentDate     =   39197
      End
      Begin VB.CommandButton CmdGuardarComentario 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7095
         TabIndex        =   22
         Top             =   8850
         Width           =   300
      End
      Begin VB.TextBox TxtComentario 
         Height          =   315
         Left            =   960
         TabIndex        =   21
         Top             =   8835
         Width           =   6090
      End
      Begin VB.TextBox TxtProfundidadInicio 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   -74790
         TabIndex        =   25
         Top             =   8805
         Width           =   975
      End
      Begin VB.TextBox TxtSpan 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   -72300
         TabIndex        =   26
         Top             =   8805
         Width           =   975
      End
      Begin iPlotLibrary.iPlotX iPlotGases 
         Height          =   8130
         Left            =   105
         TabIndex        =   146
         Top             =   75
         Width           =   7485
         DataViewZHorz   =   2
         DataViewZVert   =   1
         XYAxesReverse   =   -1  'True
         OuterMarginLeft =   5
         OuterMarginTop  =   7
         OuterMarginRight=   30
         OuterMarginBottom=   5
         PrintOrientation=   1
         PrintMarginLeft =   1
         PrintMarginTop  =   1
         PrintMarginRight=   1
         PrintMarginBottom=   1
         PrintShowDialog =   -1  'True
         UpdateFrameRate =   60
         BackGroundColor =   0
         BorderStyle     =   2
         AutoFrameRate   =   -1  'True
         HintsShow       =   0   'False
         HintsPause      =   500
         HintsHidePause  =   2500
         TitleVisible    =   0   'False
         TitleText       =   "Untitled"
         TitleMargin     =   0.25
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TitleFontColor  =   16777215
         UserCanEditObjects=   -1  'True
         LogFileName     =   ""
         LogBufferSize   =   0
         OptionSaveAllProperties=   0   'False
         BeginProperty HintsFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HintsFontColor  =   -2147483640
         BeginProperty AnnotationDefaultFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AnnotationDefaultFontColor=   16777215
         AnnotationDefaultBrushStlye=   0
         AnnotationDefaultBrushColor=   16777215
         AnnotationDefaultPenStlye=   0
         AnnotationDefaultPenColor=   16777215
         AnnotationDefaultPenWidth=   1
         Object.Width           =   499
         Object.Height          =   542
         UserCanAddRemoveChannels=   0   'False
         Object.Visible         =   -1  'True
         Enabled         =   -1  'True
         EditorFormStyle =   0
         CopyToClipBoardFormat=   0
         PrintDocumentName=   "Untitled"
         PrinterName     =   ""
         ClipAnnotationsToAxes=   -1  'True
         BackGroundGradientEnabled=   0   'False
         BackGroundGradientDirection=   0
         BackGroundGradientStartColor=   8421440
         BackGroundGradientStopColor=   0
         DataFileColumnSeparator=   0
         DataFileFormat  =   0
         DataViewZHorz   =   2
         DataViewZVert   =   1
         ChannelCount    =   2
         XAxisCount      =   1
         YAxisCount      =   2
         ToolBarCount    =   1
         LegendCount     =   1
         DataViewCount   =   1
         DataCursorCount =   1
         LimitCount      =   1
         LabelCount      =   2
         TableCount      =   0
         TranslationCount=   0
         ToolBar0.Name   =   "Toolbar 1"
         ToolBar0.Visible=   0
         ToolBar0.Enabled=   0
         ToolBar0.Layer  =   100
         ToolBar0.PopupEnabled=   0
         ToolBar0.Horizontal=   1
         ToolBar0.ZOrder =   4
         ToolBar0.StartPercent=   0
         ToolBar0.StopPercent=   100
         ToolBar0.ShowResumeButton=   1
         ToolBar0.ShowPauseButton=   1
         ToolBar0.ShowAxesModeButtons=   0
         ToolBar0.ShowZoomInOutButtons=   0
         ToolBar0.ShowSelectButton=   0
         ToolBar0.ShowZoomBoxButton=   1
         ToolBar0.ShowCursorButton=   0
         ToolBar0.ShowEditButton=   0
         ToolBar0.ShowCopyButton=   0
         ToolBar0.ShowSaveButton=   0
         ToolBar0.ShowPrintButton=   1
         ToolBar0.ShowPreviewButton=   1
         ToolBar0.ZoomInOutFactor=   2
         ToolBar0.FlatBorder=   0
         ToolBar0.FlatButtons=   0
         ToolBar0.SmallButtons=   0
         Legend0.Name    =   "Legend 1"
         Legend0.Visible =   0
         Legend0.Enabled =   0
         Legend0.Layer   =   100
         Legend0.PopupEnabled=   0
         Legend0.Horizontal=   0
         Legend0.ZOrder  =   2
         Legend0.StartPercent=   0
         Legend0.StopPercent=   100
         Legend0.MarginLeft=   1
         Legend0.MarginTop=   1
         Legend0.MarginRight=   1
         Legend0.MarginBottom=   1
         Legend0.BackGroundColor=   8421504
         Legend0.BackGroundTransparent=   1
         Legend0.SelectedItemBackGroundColor=   65535
         Legend0.SelectedItemFont.Charset=   1
         Legend0.SelectedItemFont.Color=   0
         Legend0.SelectedItemFont.Height=   -11
         Legend0.SelectedItemFont.Name=   "MS Sans Serif"
         Legend0.SelectedItemFont.Pitch=   0
         Legend0.SelectedItemFont.Style=   0
         Legend0.ShowColumnLine=   0
         Legend0.ShowColumnMarker=   0
         Legend0.ShowColumnXAxisTitle=   0
         Legend0.ShowColumnYAxisTitle=   0
         Legend0.ShowColumnXValue=   0
         Legend0.ShowColumnYValue=   0
         Legend0.ShowColumnYMax=   0
         Legend0.ShowColumnYMin=   0
         Legend0.ShowColumnYMean=   0
         Legend0.CaptionColumnTitle=   "Title"
         Legend0.CaptionColumnXAxisTitle=   "X-Axis"
         Legend0.CaptionColumnYAxisTitle=   "Y-Axis"
         Legend0.CaptionColumnXValue=   "X"
         Legend0.CaptionColumnYValue=   "Y"
         Legend0.CaptionColumnYMax=   "Y-Max"
         Legend0.CaptionColumnYMin=   "Y-Min"
         Legend0.CaptionColumnYMean=   "Y-Mean"
         Legend0.Font.Charset=   1
         Legend0.Font.Color=   16777215
         Legend0.Font.Height=   -11
         Legend0.Font.Name=   "MS Sans Serif"
         Legend0.Font.Pitch=   0
         Legend0.Font.Style=   0
         Legend0.ColumnSpacing=   0.5
         Legend0.RowSpacing=   0.25
         Legend0.WrapColDesiredCount=   1
         Legend0.WrapColAutoCountEnabled=   0
         Legend0.WrapColAutoCountMax=   100
         Legend0.WrapColSpacingMin=   2
         Legend0.WrapColSpacingAuto=   1
         Legend0.WrapRowDesiredCount=   5
         Legend0.WrapRowAutoCountEnabled=   1
         Legend0.WrapRowAutoCountMax=   100
         Legend0.WrapRowSpacingMin=   0.25
         Legend0.WrapRowSpacingAuto=   0
         Legend0.ColumnTitlesVisible=   0
         Legend0.ColumnTitlesFont.Charset=   1
         Legend0.ColumnTitlesFont.Color=   16776960
         Legend0.ColumnTitlesFont.Height=   -11
         Legend0.ColumnTitlesFont.Name=   "MS Sans Serif"
         Legend0.ColumnTitlesFont.Pitch=   0
         Legend0.ColumnTitlesFont.Style=   1
         Legend0.ChannelNameMaxWidth=   0
         Legend0.ChannelNameColorStyle=   0
         XAxis0.Name     =   "Eje de Tiempo"
         XAxis0.Visible  =   1
         XAxis0.Enabled  =   0
         XAxis0.Layer    =   100
         XAxis0.PopupEnabled=   0
         XAxis0.Horizontal=   0
         XAxis0.ZOrder   =   0
         XAxis0.StartPercent=   0
         XAxis0.StopPercent=   100
         XAxis0.Min      =   0
         XAxis0.Span     =   0.0417
         XAxis0.DesiredStart=   0
         XAxis0.DesiredIncrement=   0
         XAxis0.ReverseScale=   0
         XAxis0.InnerMargin=   5
         XAxis0.OuterMargin=   5
         XAxis0.Title    =   "Profudidad"
         XAxis0.TitleMargin=   0.5
         XAxis0.TitleFont.Charset=   1
         XAxis0.TitleFont.Color=   16777215
         XAxis0.TitleFont.Height=   -13
         XAxis0.TitleFont.Name=   "Arial"
         XAxis0.TitleFont.Pitch=   0
         XAxis0.TitleFont.Style=   1
         XAxis0.TitleShow=   0
         XAxis0.TitleRotated=   0
         XAxis0.MajorLength=   7
         XAxis0.MinorLength=   3
         XAxis0.MinorCount=   1
         XAxis0.LabelsVisible=   1
         XAxis0.LabelsMargin=   0.25
         XAxis0.LabelsFont.Charset=   1
         XAxis0.LabelsFont.Color=   16777215
         XAxis0.LabelsFont.Height=   -11
         XAxis0.LabelsFont.Name=   "MS Sans Serif"
         XAxis0.LabelsFont.Pitch=   0
         XAxis0.LabelsFont.Style=   0
         XAxis0.LabelSeparation=   2
         XAxis0.LabelsRotation=   0
         XAxis0.LabelsPrecision=   3
         XAxis0.LabelsPrecisionStyle=   0
         XAxis0.LabelsFormatStyle=   3
         XAxis0.DateTimeFormat=   "hh:nn:ss"
         XAxis0.LabelsMinLength=   5
         XAxis0.LabelsMinLengthAutoAdjust=   0
         XAxis0.ScaleLineShow=   1
         XAxis0.ScaleLinesShow=   1
         XAxis0.ScaleLinesColor=   16777215
         XAxis0.StackingEndsMargin=   0.5
         XAxis0.ScaleType=   0
         XAxis0.TrackingEnabled=   1
         XAxis0.TrackingStyle=   3
         XAxis0.TrackingAlignFirstStyle=   2
         XAxis0.TrackingScrollCompressMax=   0
         XAxis0.CursorUseDefaultFormat=   1
         XAxis0.CursorFormatStyle=   0
         XAxis0.CursorDateTimeFormat=   "hh:nn:ss"
         XAxis0.CursorPrecisionStyle=   0
         XAxis0.CursorPrecision=   3
         XAxis0.CursorMinLength=   5
         XAxis0.CursorMinLengthAutoAdjust=   0
         XAxis0.LegendUseDefaultFormat=   1
         XAxis0.LegendFormatStyle=   0
         XAxis0.LegendDateTimeFormat=   "hh:nn:ss"
         XAxis0.LegendPrecisionStyle=   0
         XAxis0.LegendPrecision=   3
         XAxis0.LegendMinLength=   5
         XAxis0.LegendMinLengthAutoAdjust=   0
         XAxis0.CursorScaler=   1
         XAxis0.ScrollMinMaxEnabled=   0
         XAxis0.ScrollMax=   100
         XAxis0.ScrollMin=   0
         XAxis0.RestoreValuesOnResume=   1
         XAxis0.MasterUIInput=   1
         XAxis0.CartesianStyle=   0
         XAxis0.CartesianChildRefAxisName=   "<None>"
         XAxis0.CartesianChildRefValue=   0
         XAxis0.AlignRefAxisName=   "<None>"
         XAxis0.GridLinesVisible=   1
         XAxis0.ForceStacking=   0
         YAxis0.Name     =   "GasTotal"
         YAxis0.Visible  =   1
         YAxis0.Enabled  =   1
         YAxis0.Layer    =   100
         YAxis0.PopupEnabled=   0
         YAxis0.Horizontal=   1
         YAxis0.ZOrder   =   1
         YAxis0.StartPercent=   0
         YAxis0.StopPercent=   100
         YAxis0.Min      =   0
         YAxis0.Span     =   30000
         YAxis0.DesiredStart=   0
         YAxis0.DesiredIncrement=   0
         YAxis0.ReverseScale=   0
         YAxis0.InnerMargin=   5
         YAxis0.OuterMargin=   5
         YAxis0.Title    =   "GasTotal"
         YAxis0.TitleMargin=   0.25
         YAxis0.TitleFont.Charset=   1
         YAxis0.TitleFont.Color=   16777215
         YAxis0.TitleFont.Height=   -13
         YAxis0.TitleFont.Name=   "Arial"
         YAxis0.TitleFont.Pitch=   0
         YAxis0.TitleFont.Style=   1
         YAxis0.TitleShow=   1
         YAxis0.TitleRotated=   0
         YAxis0.MajorLength=   7
         YAxis0.MinorLength=   3
         YAxis0.MinorCount=   1
         YAxis0.LabelsVisible=   1
         YAxis0.LabelsMargin=   0.25
         YAxis0.LabelsFont.Charset=   1
         YAxis0.LabelsFont.Color=   16777215
         YAxis0.LabelsFont.Height=   -11
         YAxis0.LabelsFont.Name=   "MS Sans Serif"
         YAxis0.LabelsFont.Pitch=   0
         YAxis0.LabelsFont.Style=   0
         YAxis0.LabelSeparation=   2
         YAxis0.LabelsRotation=   0
         YAxis0.LabelsPrecision=   3
         YAxis0.LabelsPrecisionStyle=   0
         YAxis0.LabelsFormatStyle=   0
         YAxis0.DateTimeFormat=   "hh:nn:ss"
         YAxis0.LabelsMinLength=   5
         YAxis0.LabelsMinLengthAutoAdjust=   0
         YAxis0.ScaleLineShow=   1
         YAxis0.ScaleLinesShow=   1
         YAxis0.ScaleLinesColor=   16777215
         YAxis0.StackingEndsMargin=   0.5
         YAxis0.ScaleType=   0
         YAxis0.TrackingEnabled=   0
         YAxis0.TrackingStyle=   0
         YAxis0.TrackingAlignFirstStyle=   3
         YAxis0.TrackingScrollCompressMax=   0
         YAxis0.CursorUseDefaultFormat=   1
         YAxis0.CursorFormatStyle=   0
         YAxis0.CursorDateTimeFormat=   "hh:nn:ss"
         YAxis0.CursorPrecisionStyle=   0
         YAxis0.CursorPrecision=   3
         YAxis0.CursorMinLength=   5
         YAxis0.CursorMinLengthAutoAdjust=   0
         YAxis0.LegendUseDefaultFormat=   1
         YAxis0.LegendFormatStyle=   0
         YAxis0.LegendDateTimeFormat=   "hh:nn:ss"
         YAxis0.LegendPrecisionStyle=   0
         YAxis0.LegendPrecision=   3
         YAxis0.LegendMinLength=   5
         YAxis0.LegendMinLengthAutoAdjust=   0
         YAxis0.CursorScaler=   1
         YAxis0.ScrollMinMaxEnabled=   0
         YAxis0.ScrollMax=   100
         YAxis0.ScrollMin=   0
         YAxis0.RestoreValuesOnResume=   1
         YAxis0.MasterUIInput=   0
         YAxis0.CartesianStyle=   0
         YAxis0.CartesianChildRefAxisName=   "<None>"
         YAxis0.CartesianChildRefValue=   0
         YAxis0.AlignRefAxisName=   "<None>"
         YAxis0.GridLinesVisible=   1
         YAxis0.ForceStacking=   0
         YAxis1.Name     =   "CO2"
         YAxis1.Visible  =   0
         YAxis1.Enabled  =   0
         YAxis1.Layer    =   100
         YAxis1.PopupEnabled=   0
         YAxis1.Horizontal=   1
         YAxis1.ZOrder   =   0
         YAxis1.StartPercent=   0
         YAxis1.StopPercent=   100
         YAxis1.Min      =   0
         YAxis1.Span     =   30000
         YAxis1.DesiredStart=   0
         YAxis1.DesiredIncrement=   0
         YAxis1.ReverseScale=   0
         YAxis1.InnerMargin=   5
         YAxis1.OuterMargin=   5
         YAxis1.Title    =   "CO2"
         YAxis1.TitleMargin=   0.25
         YAxis1.TitleFont.Charset=   1
         YAxis1.TitleFont.Color=   16777215
         YAxis1.TitleFont.Height=   -13
         YAxis1.TitleFont.Name=   "Arial"
         YAxis1.TitleFont.Pitch=   0
         YAxis1.TitleFont.Style=   1
         YAxis1.TitleShow=   0
         YAxis1.TitleRotated=   0
         YAxis1.MajorLength=   7
         YAxis1.MinorLength=   3
         YAxis1.MinorCount=   1
         YAxis1.LabelsVisible=   1
         YAxis1.LabelsMargin=   0.25
         YAxis1.LabelsFont.Charset=   1
         YAxis1.LabelsFont.Color=   16777215
         YAxis1.LabelsFont.Height=   -11
         YAxis1.LabelsFont.Name=   "MS Sans Serif"
         YAxis1.LabelsFont.Pitch=   0
         YAxis1.LabelsFont.Style=   0
         YAxis1.LabelSeparation=   2
         YAxis1.LabelsRotation=   0
         YAxis1.LabelsPrecision=   3
         YAxis1.LabelsPrecisionStyle=   0
         YAxis1.LabelsFormatStyle=   0
         YAxis1.DateTimeFormat=   "hh:nn:ss"
         YAxis1.LabelsMinLength=   5
         YAxis1.LabelsMinLengthAutoAdjust=   0
         YAxis1.ScaleLineShow=   1
         YAxis1.ScaleLinesShow=   1
         YAxis1.ScaleLinesColor=   16777215
         YAxis1.StackingEndsMargin=   0.5
         YAxis1.ScaleType=   0
         YAxis1.TrackingEnabled=   1
         YAxis1.TrackingStyle=   3
         YAxis1.TrackingAlignFirstStyle=   3
         YAxis1.TrackingScrollCompressMax=   0
         YAxis1.CursorUseDefaultFormat=   1
         YAxis1.CursorFormatStyle=   0
         YAxis1.CursorDateTimeFormat=   "hh:nn:ss"
         YAxis1.CursorPrecisionStyle=   0
         YAxis1.CursorPrecision=   3
         YAxis1.CursorMinLength=   5
         YAxis1.CursorMinLengthAutoAdjust=   0
         YAxis1.LegendUseDefaultFormat=   1
         YAxis1.LegendFormatStyle=   0
         YAxis1.LegendDateTimeFormat=   "hh:nn:ss"
         YAxis1.LegendPrecisionStyle=   0
         YAxis1.LegendPrecision=   3
         YAxis1.LegendMinLength=   5
         YAxis1.LegendMinLengthAutoAdjust=   0
         YAxis1.CursorScaler=   1
         YAxis1.ScrollMinMaxEnabled=   0
         YAxis1.ScrollMax=   100
         YAxis1.ScrollMin=   0
         YAxis1.RestoreValuesOnResume=   1
         YAxis1.MasterUIInput=   0
         YAxis1.CartesianStyle=   0
         YAxis1.CartesianChildRefAxisName=   "<None>"
         YAxis1.CartesianChildRefValue=   0
         YAxis1.AlignRefAxisName=   "<None>"
         YAxis1.GridLinesVisible=   0
         YAxis1.ForceStacking=   0
         DataView0.Name  =   "Data View 1"
         DataView0.Visible=   1
         DataView0.Enabled=   1
         DataView0.Layer =   100
         DataView0.PopupEnabled=   0
         DataView0.Horizontal=   0
         DataView0.ZOrder=   0
         DataView0.StartPercent=   0
         DataView0.StopPercent=   100
         DataView0.Title =   ""
         DataView0.BackgroundTransparent=   1
         DataView0.BackgroundColor=   8421376
         DataView0.GridXAxisName=   "<All>"
         DataView0.GridYAxisName=   "<All>"
         DataView0.GridShow=   1
         DataView0.GridLineColor=   32768
         DataView0.GridLineShowLeft=   1
         DataView0.GridLineShowRight=   1
         DataView0.GridLineShowTop=   1
         DataView0.GridLineShowBottom=   1
         DataView0.GridLineShowXMajors=   1
         DataView0.GridLineShowXMinors=   1
         DataView0.GridLineShowYMajors=   1
         DataView0.GridLineShowYMinors=   1
         DataView0.GridLineMajorStyle=   0
         DataView0.GridLineMinorStyle=   0
         DataView0.GridLineXMajorCustom=   0
         DataView0.GridLineXMajorColor=   32768
         DataView0.GridLineXMajorWidth=   0
         DataView0.GridLineXMajorStyle=   0
         DataView0.GridLineXMinorCustom=   0
         DataView0.GridLineXMinorColor=   32768
         DataView0.GridLineXMinorWidth=   0
         DataView0.GridLineXMinorStyle=   0
         DataView0.GridLineYMajorCustom=   0
         DataView0.GridLineYMajorColor=   32768
         DataView0.GridLineYMajorWidth=   0
         DataView0.GridLineYMajorStyle=   0
         DataView0.GridLineYMinorCustom=   0
         DataView0.GridLineYMinorColor=   32768
         DataView0.GridLineYMinorWidth=   0
         DataView0.GridLineYMinorStyle=   0
         DataView0.AxesControlEnabled=   0
         DataView0.AxesControlMouseStyle=   2
         DataView0.AxesControlWheelStyle=   0
         Channel0.Name   =   "GasTotal"
         Channel0.Visible=   1
         Channel0.Enabled=   1
         Channel0.Layer  =   100
         Channel0.PopupEnabled=   0
         Channel0.TitleText=   "GasTotal"
         Channel0.VisibleInLegend=   1
         Channel0.RingBufferSize=   14400
         Channel0.TraceVisible=   1
         Channel0.Color  =   255
         Channel0.TraceLineStyle=   0
         Channel0.TraceLineWidth=   3
         Channel0.MarkersAllowIndividual=   0
         Channel0.MarkersPenUseChannelColor=   1
         Channel0.MarkersBrushUseChannelColor=   1
         Channel0.MarkersTurnOffLimit=   0
         Channel0.MarkersVisible=   0
         Channel0.MarkersSize=   3
         Channel0.MarkersStyle=   0
         Channel0.MarkersPenColor=   255
         Channel0.MarkersPenStyle=   0
         Channel0.MarkersPenWidth=   0
         Channel0.MarkersBrushColor=   255
         Channel0.MarkersBrushStyle=   0
         Channel0.MarkersFont.Charset=   1
         Channel0.MarkersFont.Color=   16777215
         Channel0.MarkersFont.Height=   -11
         Channel0.MarkersFont.Name=   "MS Sans Serif"
         Channel0.MarkersFont.Pitch=   0
         Channel0.MarkersFont.Style=   1
         Channel0.XAxisName=   "Eje de Tiempo"
         Channel0.YAxisName=   "GasTotal"
         Channel0.XAxisTrackingEnabled=   1
         Channel0.YAxisTrackingEnabled=   1
         Channel0.LogFileName=   ""
         Channel0.LogBufferSize=   0
         Channel0.DataStyle=   0
         Channel0.Tag    =   0
         Channel0.OPCGroupName=   ""
         Channel0.OPCComputerName=   "Local"
         Channel0.OPCServerName=   ""
         Channel0.OPCItemName=   ""
         Channel0.OPCUpdateRate=   500
         Channel0.OPCAutoConnect=   1
         Channel0.FastDrawEnabled=   1
         Channel0.InterpolationStyle=   0
         Channel0.FillEnabled=   0
         Channel0.FillReference=   0
         Channel0.FillStyle=   0
         Channel0.FillColor=   0
         Channel0.FillUseChannelColor=   1
         Channel0.DigitalEnabled=   0
         Channel0.DigitalReferenceStyle=   0
         Channel0.DigitalReferenceLow=   10
         Channel0.DigitalReferenceHigh=   90
         Channel0.HighLowStyle=   0
         Channel0.HighLowEnabled=   0
         Channel0.HighLowBarColor=   16776960
         Channel0.HighLowBarWidth=   0.5
         Channel0.HighLowOpenShow=   1
         Channel0.HighLowOpenColor=   65280
         Channel0.HighLowOpenWidth=   1
         Channel0.HighLowOpenHeight=   1
         Channel0.HighLowCloseShow=   1
         Channel0.HighLowCloseColor=   255
         Channel0.HighLowCloseWidth=   1
         Channel0.HighLowCloseHeight=   1
         Channel0.HighLowShadowColor=   8421504
         Channel0.HighLowBullishColor=   16777215
         Channel0.HighLowBearishColor=   8421504
         Channel0.BarEnabled=   0
         Channel0.BarPenUseChannelColor=   1
         Channel0.BarBrushUseChannelColor=   1
         Channel0.BarReference=   0
         Channel0.BarWidth=   5
         Channel0.BarPenColor=   255
         Channel0.BarPenWidth=   0
         Channel0.BarPenStyle=   0
         Channel0.BarBrushColor=   255
         Channel0.BarBrushStyle=   0
         Channel0.OPCXValueSource=   0
         Channel1.Name   =   "CO2"
         Channel1.Visible=   0
         Channel1.Enabled=   1
         Channel1.Layer  =   100
         Channel1.PopupEnabled=   0
         Channel1.TitleText=   "CO2"
         Channel1.VisibleInLegend=   0
         Channel1.RingBufferSize=   14400
         Channel1.TraceVisible=   1
         Channel1.Color  =   65535
         Channel1.TraceLineStyle=   0
         Channel1.TraceLineWidth=   2
         Channel1.MarkersAllowIndividual=   0
         Channel1.MarkersPenUseChannelColor=   1
         Channel1.MarkersBrushUseChannelColor=   1
         Channel1.MarkersTurnOffLimit=   0
         Channel1.MarkersVisible=   0
         Channel1.MarkersSize=   3
         Channel1.MarkersStyle=   0
         Channel1.MarkersPenColor=   255
         Channel1.MarkersPenStyle=   0
         Channel1.MarkersPenWidth=   0
         Channel1.MarkersBrushColor=   255
         Channel1.MarkersBrushStyle=   0
         Channel1.MarkersFont.Charset=   1
         Channel1.MarkersFont.Color=   16777215
         Channel1.MarkersFont.Height=   -11
         Channel1.MarkersFont.Name=   "MS Sans Serif"
         Channel1.MarkersFont.Pitch=   0
         Channel1.MarkersFont.Style=   1
         Channel1.XAxisName=   "Eje de Tiempo"
         Channel1.YAxisName=   "CO2"
         Channel1.XAxisTrackingEnabled=   1
         Channel1.YAxisTrackingEnabled=   1
         Channel1.LogFileName=   ""
         Channel1.LogBufferSize=   0
         Channel1.DataStyle=   0
         Channel1.Tag    =   0
         Channel1.OPCGroupName=   ""
         Channel1.OPCComputerName=   "Local"
         Channel1.OPCServerName=   ""
         Channel1.OPCItemName=   ""
         Channel1.OPCUpdateRate=   500
         Channel1.OPCAutoConnect=   1
         Channel1.FastDrawEnabled=   0
         Channel1.InterpolationStyle=   0
         Channel1.FillEnabled=   0
         Channel1.FillReference=   0
         Channel1.FillStyle=   0
         Channel1.FillColor=   0
         Channel1.FillUseChannelColor=   1
         Channel1.DigitalEnabled=   0
         Channel1.DigitalReferenceStyle=   0
         Channel1.DigitalReferenceLow=   10
         Channel1.DigitalReferenceHigh=   90
         Channel1.HighLowStyle=   0
         Channel1.HighLowEnabled=   0
         Channel1.HighLowBarColor=   16776960
         Channel1.HighLowBarWidth=   0.5
         Channel1.HighLowOpenShow=   1
         Channel1.HighLowOpenColor=   65280
         Channel1.HighLowOpenWidth=   1
         Channel1.HighLowOpenHeight=   1
         Channel1.HighLowCloseShow=   1
         Channel1.HighLowCloseColor=   255
         Channel1.HighLowCloseWidth=   1
         Channel1.HighLowCloseHeight=   1
         Channel1.HighLowShadowColor=   8421504
         Channel1.HighLowBullishColor=   16777215
         Channel1.HighLowBearishColor=   8421504
         Channel1.BarEnabled=   0
         Channel1.BarPenUseChannelColor=   1
         Channel1.BarBrushUseChannelColor=   1
         Channel1.BarReference=   0
         Channel1.BarWidth=   5
         Channel1.BarPenColor=   255
         Channel1.BarPenWidth=   0
         Channel1.BarPenStyle=   0
         Channel1.BarBrushColor=   255
         Channel1.BarBrushStyle=   0
         Channel1.OPCXValueSource=   0
         DataCursor0.Name=   "Cursor 1"
         DataCursor0.Visible=   0
         DataCursor0.Enabled=   1
         DataCursor0.Layer=   100
         DataCursor0.PopupEnabled=   0
         DataCursor0.ChannelName=   "GasTotal"
         DataCursor0.ChannelAllowAll=   1
         DataCursor0.ChannelShowAllInLegend=   1
         DataCursor0.Style=   0
         DataCursor0.Font.Charset=   1
         DataCursor0.Font.Color=   -2147483640
         DataCursor0.Font.Height=   -11
         DataCursor0.Font.Name=   "MS Sans Serif"
         DataCursor0.Font.Pitch=   0
         DataCursor0.Font.Style=   0
         DataCursor0.Color=   65535
         DataCursor0.UseChannelColor=   1
         DataCursor0.HintShow=   1
         DataCursor0.HintHideOnRelease=   0
         DataCursor0.HintOrientationSide=   0
         DataCursor0.HintPosition=   50
         DataCursor0.Pointer1Position=   50
         DataCursor0.Pointer2Position=   60
         DataCursor0.PointerPenWidth=   1
         DataCursor0.MenuUserCanChangeOptions=   1
         DataCursor0.MenuItemVisibleValueXY=   1
         DataCursor0.MenuItemVisibleValueX=   1
         DataCursor0.MenuItemVisibleValueY=   1
         DataCursor0.MenuItemVisibleDeltaX=   1
         DataCursor0.MenuItemVisibleDeltaY=   1
         DataCursor0.MenuItemVisibleInverseDeltaX=   1
         DataCursor0.MenuItemCaptionValueXY=   "Value X-Y"
         DataCursor0.MenuItemCaptionValueX=   "Value X"
         DataCursor0.MenuItemCaptionValueY=   "Value Y"
         DataCursor0.MenuItemCaptionDeltaX=   "Period"
         DataCursor0.MenuItemCaptionDeltaY=   "Peak-Peak"
         DataCursor0.MenuItemCaptionInverseDeltaX=   "Frequency"
         Limit0.Name     =   "NivelDisparo"
         Limit0.Visible  =   1
         Limit0.Enabled  =   1
         Limit0.Layer    =   100
         Limit0.PopupEnabled=   0
         Limit0.Color    =   255
         Limit0.LineStyle=   2
         Limit0.LineWidth=   1
         Limit0.FillStyle=   1
         Limit0.XAxisName=   "Eje de Tiempo"
         Limit0.YAxisName=   "GasTotal"
         Limit0.Style    =   1
         Limit0.Line1Position=   500
         Limit0.Line2Position=   50
         Limit0.UserCanMove=   0
         Label0.Name     =   "Title"
         Label0.Visible  =   0
         Label0.Enabled  =   1
         Label0.Layer    =   100
         Label0.PopupEnabled=   1
         Label0.Horizontal=   1
         Label0.ZOrder   =   3
         Label0.StartPercent=   0
         Label0.StopPercent=   100
         Label0.MarginLeft=   0
         Label0.MarginTop=   0
         Label0.MarginRight=   0
         Label0.MarginBottom=   0.25
         Label0.Caption  =   "Untitled"
         Label0.Alignment=   0
         Label0.Font.Charset=   1
         Label0.Font.Color=   16777215
         Label0.Font.Height=   -19
         Label0.Font.Name=   "Arial"
         Label0.Font.Pitch=   0
         Label0.Font.Style=   1
         Label1.Name     =   "Label 1"
         Label1.Visible  =   1
         Label1.Enabled  =   1
         Label1.Layer    =   100
         Label1.PopupEnabled=   1
         Label1.Horizontal=   0
         Label1.ZOrder   =   3
         Label1.StartPercent=   0
         Label1.StopPercent=   100
         Label1.MarginLeft=   0
         Label1.MarginTop=   0
         Label1.MarginRight=   0
         Label1.MarginBottom=   0
         Label1.Caption  =   ""
         Label1.Alignment=   0
         Label1.Font.Charset=   1
         Label1.Font.Color=   16777215
         Label1.Font.Height=   -32
         Label1.Font.Name=   "Arial"
         Label1.Font.Pitch=   0
         Label1.Font.Style=   1
      End
      Begin iPlotLibrary.iPlotX iPlotMasterLog 
         Height          =   8655
         Left            =   -74880
         Negotiate       =   -1  'True
         TabIndex        =   147
         Top             =   15
         Width           =   7815
         DataViewZHorz   =   1
         DataViewZVert   =   1
         XYAxesReverse   =   -1  'True
         OuterMarginLeft =   5
         OuterMarginTop  =   5
         OuterMarginRight=   60
         OuterMarginBottom=   5
         PrintOrientation=   1
         PrintMarginLeft =   1
         PrintMarginTop  =   1
         PrintMarginRight=   1
         PrintMarginBottom=   1
         PrintShowDialog =   -1  'True
         UpdateFrameRate =   60
         BackGroundColor =   0
         BorderStyle     =   2
         AutoFrameRate   =   -1  'True
         HintsShow       =   -1  'True
         HintsPause      =   500
         HintsHidePause  =   2500
         TitleVisible    =   0   'False
         TitleText       =   "Untitled"
         TitleMargin     =   0.25
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TitleFontColor  =   16777215
         UserCanEditObjects=   -1  'True
         LogFileName     =   ""
         LogBufferSize   =   0
         OptionSaveAllProperties=   0   'False
         BeginProperty HintsFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HintsFontColor  =   -2147483640
         BeginProperty AnnotationDefaultFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AnnotationDefaultFontColor=   16777215
         AnnotationDefaultBrushStlye=   0
         AnnotationDefaultBrushColor=   16777215
         AnnotationDefaultPenStlye=   0
         AnnotationDefaultPenColor=   16777215
         AnnotationDefaultPenWidth=   1
         Object.Width           =   0
         Object.Height          =   0
         UserCanAddRemoveChannels=   0   'False
         Object.Visible         =   -1  'True
         Enabled         =   -1  'True
         EditorFormStyle =   0
         CopyToClipBoardFormat=   0
         PrintDocumentName=   "Untitled"
         PrinterName     =   ""
         ClipAnnotationsToAxes=   -1  'True
         BackGroundGradientEnabled=   0   'False
         BackGroundGradientDirection=   0
         BackGroundGradientStartColor=   16711680
         BackGroundGradientStopColor=   0
         DataFileColumnSeparator=   0
         DataFileFormat  =   0
         DataViewZHorz   =   1
         DataViewZVert   =   1
         ChannelCount    =   10
         XAxisCount      =   1
         YAxisCount      =   2
         ToolBarCount    =   1
         LegendCount     =   1
         DataViewCount   =   1
         DataCursorCount =   1
         LimitCount      =   0
         LabelCount      =   1
         TableCount      =   0
         TranslationCount=   0
         ToolBar0.Name   =   "Toolbar 1"
         ToolBar0.Visible=   0
         ToolBar0.Enabled=   0
         ToolBar0.Layer  =   100
         ToolBar0.PopupEnabled=   0
         ToolBar0.Horizontal=   1
         ToolBar0.ZOrder =   3
         ToolBar0.StartPercent=   0
         ToolBar0.StopPercent=   100
         ToolBar0.ShowResumeButton=   0
         ToolBar0.ShowPauseButton=   0
         ToolBar0.ShowAxesModeButtons=   0
         ToolBar0.ShowZoomInOutButtons=   0
         ToolBar0.ShowSelectButton=   0
         ToolBar0.ShowZoomBoxButton=   0
         ToolBar0.ShowCursorButton=   0
         ToolBar0.ShowEditButton=   0
         ToolBar0.ShowCopyButton=   0
         ToolBar0.ShowSaveButton=   0
         ToolBar0.ShowPrintButton=   0
         ToolBar0.ShowPreviewButton=   0
         ToolBar0.ZoomInOutFactor=   2
         ToolBar0.FlatBorder=   0
         ToolBar0.FlatButtons=   0
         ToolBar0.SmallButtons=   0
         Legend0.Name    =   "Legend 1"
         Legend0.Visible =   0
         Legend0.Enabled =   0
         Legend0.Layer   =   100
         Legend0.PopupEnabled=   0
         Legend0.Horizontal=   0
         Legend0.ZOrder  =   2
         Legend0.StartPercent=   0
         Legend0.StopPercent=   100
         Legend0.MarginLeft=   1
         Legend0.MarginTop=   1
         Legend0.MarginRight=   1
         Legend0.MarginBottom=   1
         Legend0.BackGroundColor=   8421504
         Legend0.BackGroundTransparent=   1
         Legend0.SelectedItemBackGroundColor=   65535
         Legend0.SelectedItemFont.Charset=   1
         Legend0.SelectedItemFont.Color=   0
         Legend0.SelectedItemFont.Height=   -11
         Legend0.SelectedItemFont.Name=   "MS Sans Serif"
         Legend0.SelectedItemFont.Pitch=   0
         Legend0.SelectedItemFont.Style=   0
         Legend0.ShowColumnLine=   1
         Legend0.ShowColumnMarker=   0
         Legend0.ShowColumnXAxisTitle=   0
         Legend0.ShowColumnYAxisTitle=   0
         Legend0.ShowColumnXValue=   0
         Legend0.ShowColumnYValue=   0
         Legend0.ShowColumnYMax=   0
         Legend0.ShowColumnYMin=   0
         Legend0.ShowColumnYMean=   0
         Legend0.CaptionColumnTitle=   "Title"
         Legend0.CaptionColumnXAxisTitle=   "X-Axis"
         Legend0.CaptionColumnYAxisTitle=   "Y-Axis"
         Legend0.CaptionColumnXValue=   "X"
         Legend0.CaptionColumnYValue=   "Y"
         Legend0.CaptionColumnYMax=   "Y-Max"
         Legend0.CaptionColumnYMin=   "Y-Min"
         Legend0.CaptionColumnYMean=   "Y-Mean"
         Legend0.Font.Charset=   1
         Legend0.Font.Color=   16777215
         Legend0.Font.Height=   -11
         Legend0.Font.Name=   "MS Sans Serif"
         Legend0.Font.Pitch=   0
         Legend0.Font.Style=   0
         Legend0.ColumnSpacing=   0.5
         Legend0.RowSpacing=   0.25
         Legend0.WrapColDesiredCount=   1
         Legend0.WrapColAutoCountEnabled=   0
         Legend0.WrapColAutoCountMax=   100
         Legend0.WrapColSpacingMin=   2
         Legend0.WrapColSpacingAuto=   1
         Legend0.WrapRowDesiredCount=   5
         Legend0.WrapRowAutoCountEnabled=   1
         Legend0.WrapRowAutoCountMax=   100
         Legend0.WrapRowSpacingMin=   0.25
         Legend0.WrapRowSpacingAuto=   0
         Legend0.ColumnTitlesVisible=   0
         Legend0.ColumnTitlesFont.Charset=   1
         Legend0.ColumnTitlesFont.Color=   16776960
         Legend0.ColumnTitlesFont.Height=   -11
         Legend0.ColumnTitlesFont.Name=   "MS Sans Serif"
         Legend0.ColumnTitlesFont.Pitch=   0
         Legend0.ColumnTitlesFont.Style=   1
         Legend0.ChannelNameMaxWidth=   0
         Legend0.ChannelNameColorStyle=   0
         XAxis0.Name     =   "Profundidad"
         XAxis0.Visible  =   1
         XAxis0.Enabled  =   0
         XAxis0.Layer    =   100
         XAxis0.PopupEnabled=   0
         XAxis0.Horizontal=   0
         XAxis0.ZOrder   =   0
         XAxis0.StartPercent=   0
         XAxis0.StopPercent=   100
         XAxis0.Min      =   900
         XAxis0.Span     =   50
         XAxis0.DesiredStart=   0
         XAxis0.DesiredIncrement=   1
         XAxis0.ReverseScale=   0
         XAxis0.InnerMargin=   5
         XAxis0.OuterMargin=   5
         XAxis0.Title    =   "Profundidad"
         XAxis0.TitleMargin=   0.5
         XAxis0.TitleFont.Charset=   1
         XAxis0.TitleFont.Color=   16777215
         XAxis0.TitleFont.Height=   -13
         XAxis0.TitleFont.Name=   "Arial"
         XAxis0.TitleFont.Pitch=   0
         XAxis0.TitleFont.Style=   1
         XAxis0.TitleShow=   0
         XAxis0.TitleRotated=   0
         XAxis0.MajorLength=   7
         XAxis0.MinorLength=   3
         XAxis0.MinorCount=   1
         XAxis0.LabelsVisible=   1
         XAxis0.LabelsMargin=   0.25
         XAxis0.LabelsFont.Charset=   1
         XAxis0.LabelsFont.Color=   16777215
         XAxis0.LabelsFont.Height=   -11
         XAxis0.LabelsFont.Name=   "MS Sans Serif"
         XAxis0.LabelsFont.Pitch=   0
         XAxis0.LabelsFont.Style=   0
         XAxis0.LabelSeparation=   1
         XAxis0.LabelsRotation=   0
         XAxis0.LabelsPrecision=   0
         XAxis0.LabelsPrecisionStyle=   0
         XAxis0.LabelsFormatStyle=   0
         XAxis0.DateTimeFormat=   "hh:nn:ss"
         XAxis0.LabelsMinLength=   3
         XAxis0.LabelsMinLengthAutoAdjust=   1
         XAxis0.ScaleLineShow=   1
         XAxis0.ScaleLinesShow=   1
         XAxis0.ScaleLinesColor=   16777215
         XAxis0.StackingEndsMargin=   0.5
         XAxis0.ScaleType=   0
         XAxis0.TrackingEnabled=   1
         XAxis0.TrackingStyle=   3
         XAxis0.TrackingAlignFirstStyle=   2
         XAxis0.TrackingScrollCompressMax=   0
         XAxis0.CursorUseDefaultFormat=   1
         XAxis0.CursorFormatStyle=   0
         XAxis0.CursorDateTimeFormat=   "hh:nn:ss"
         XAxis0.CursorPrecisionStyle=   0
         XAxis0.CursorPrecision=   3
         XAxis0.CursorMinLength=   5
         XAxis0.CursorMinLengthAutoAdjust=   0
         XAxis0.LegendUseDefaultFormat=   1
         XAxis0.LegendFormatStyle=   0
         XAxis0.LegendDateTimeFormat=   "hh:nn:ss"
         XAxis0.LegendPrecisionStyle=   0
         XAxis0.LegendPrecision=   3
         XAxis0.LegendMinLength=   5
         XAxis0.LegendMinLengthAutoAdjust=   0
         XAxis0.CursorScaler=   1
         XAxis0.ScrollMinMaxEnabled=   0
         XAxis0.ScrollMax=   100
         XAxis0.ScrollMin=   0
         XAxis0.RestoreValuesOnResume=   1
         XAxis0.MasterUIInput=   0
         XAxis0.CartesianStyle=   0
         XAxis0.CartesianChildRefAxisName=   "<None>"
         XAxis0.CartesianChildRefValue=   0
         XAxis0.AlignRefAxisName=   "<None>"
         XAxis0.GridLinesVisible=   1
         XAxis0.ForceStacking=   0
         YAxis0.Name     =   "Crono"
         YAxis0.Visible  =   1
         YAxis0.Enabled  =   0
         YAxis0.Layer    =   100
         YAxis0.PopupEnabled=   0
         YAxis0.Horizontal=   1
         YAxis0.ZOrder   =   0
         YAxis0.StartPercent=   0
         YAxis0.StopPercent=   20
         YAxis0.Min      =   0
         YAxis0.Span     =   2
         YAxis0.DesiredStart=   0
         YAxis0.DesiredIncrement=   1
         YAxis0.ReverseScale=   0
         YAxis0.InnerMargin=   5
         YAxis0.OuterMargin=   5
         YAxis0.Title    =   "Crono"
         YAxis0.TitleMargin=   0.25
         YAxis0.TitleFont.Charset=   1
         YAxis0.TitleFont.Color=   16777215
         YAxis0.TitleFont.Height=   -13
         YAxis0.TitleFont.Name=   "Arial"
         YAxis0.TitleFont.Pitch=   0
         YAxis0.TitleFont.Style=   1
         YAxis0.TitleShow=   1
         YAxis0.TitleRotated=   0
         YAxis0.MajorLength=   7
         YAxis0.MinorLength=   3
         YAxis0.MinorCount=   1
         YAxis0.LabelsVisible=   1
         YAxis0.LabelsMargin=   0.25
         YAxis0.LabelsFont.Charset=   1
         YAxis0.LabelsFont.Color=   16777215
         YAxis0.LabelsFont.Height=   -8
         YAxis0.LabelsFont.Name=   "MS Sans Serif"
         YAxis0.LabelsFont.Pitch=   0
         YAxis0.LabelsFont.Style=   0
         YAxis0.LabelSeparation=   2
         YAxis0.LabelsRotation=   0
         YAxis0.LabelsPrecision=   0
         YAxis0.LabelsPrecisionStyle=   0
         YAxis0.LabelsFormatStyle=   0
         YAxis0.DateTimeFormat=   "hh:nn:ss"
         YAxis0.LabelsMinLength=   1
         YAxis0.LabelsMinLengthAutoAdjust=   0
         YAxis0.ScaleLineShow=   1
         YAxis0.ScaleLinesShow=   1
         YAxis0.ScaleLinesColor=   16777215
         YAxis0.StackingEndsMargin=   0.5
         YAxis0.ScaleType=   0
         YAxis0.TrackingEnabled=   0
         YAxis0.TrackingStyle=   0
         YAxis0.TrackingAlignFirstStyle=   3
         YAxis0.TrackingScrollCompressMax=   0
         YAxis0.CursorUseDefaultFormat=   1
         YAxis0.CursorFormatStyle=   0
         YAxis0.CursorDateTimeFormat=   "hh:nn:ss"
         YAxis0.CursorPrecisionStyle=   0
         YAxis0.CursorPrecision=   3
         YAxis0.CursorMinLength=   5
         YAxis0.CursorMinLengthAutoAdjust=   0
         YAxis0.LegendUseDefaultFormat=   1
         YAxis0.LegendFormatStyle=   0
         YAxis0.LegendDateTimeFormat=   "hh:nn:ss"
         YAxis0.LegendPrecisionStyle=   0
         YAxis0.LegendPrecision=   3
         YAxis0.LegendMinLength=   5
         YAxis0.LegendMinLengthAutoAdjust=   0
         YAxis0.CursorScaler=   1
         YAxis0.ScrollMinMaxEnabled=   0
         YAxis0.ScrollMax=   100
         YAxis0.ScrollMin=   0
         YAxis0.RestoreValuesOnResume=   1
         YAxis0.MasterUIInput=   0
         YAxis0.CartesianStyle=   0
         YAxis0.CartesianChildRefAxisName=   "<None>"
         YAxis0.CartesianChildRefValue=   0
         YAxis0.AlignRefAxisName=   "<None>"
         YAxis0.GridLinesVisible=   1
         YAxis0.ForceStacking=   0
         YAxis1.Name     =   "Gas-Cromatografia"
         YAxis1.Visible  =   1
         YAxis1.Enabled  =   1
         YAxis1.Layer    =   100
         YAxis1.PopupEnabled=   0
         YAxis1.Horizontal=   1
         YAxis1.ZOrder   =   0
         YAxis1.StartPercent=   20
         YAxis1.StopPercent=   100
         YAxis1.Min      =   10
         YAxis1.Span     =   500000
         YAxis1.DesiredStart=   1000
         YAxis1.DesiredIncrement=   10000
         YAxis1.ReverseScale=   0
         YAxis1.InnerMargin=   5
         YAxis1.OuterMargin=   5
         YAxis1.Title    =   "Gas - Cromatografía"
         YAxis1.TitleMargin=   0.25
         YAxis1.TitleFont.Charset=   1
         YAxis1.TitleFont.Color=   16777215
         YAxis1.TitleFont.Height=   -13
         YAxis1.TitleFont.Name=   "Arial"
         YAxis1.TitleFont.Pitch=   0
         YAxis1.TitleFont.Style=   1
         YAxis1.TitleShow=   1
         YAxis1.TitleRotated=   0
         YAxis1.MajorLength=   7
         YAxis1.MinorLength=   3
         YAxis1.MinorCount=   1
         YAxis1.LabelsVisible=   1
         YAxis1.LabelsMargin=   0.25
         YAxis1.LabelsFont.Charset=   1
         YAxis1.LabelsFont.Color=   16777215
         YAxis1.LabelsFont.Height=   -8
         YAxis1.LabelsFont.Name=   "MS Sans Serif"
         YAxis1.LabelsFont.Pitch=   0
         YAxis1.LabelsFont.Style=   0
         YAxis1.LabelSeparation=   2
         YAxis1.LabelsRotation=   0
         YAxis1.LabelsPrecision=   0
         YAxis1.LabelsPrecisionStyle=   0
         YAxis1.LabelsFormatStyle=   0
         YAxis1.DateTimeFormat=   "hh:nn:ss"
         YAxis1.LabelsMinLength=   1
         YAxis1.LabelsMinLengthAutoAdjust=   0
         YAxis1.ScaleLineShow=   1
         YAxis1.ScaleLinesShow=   1
         YAxis1.ScaleLinesColor=   16777215
         YAxis1.StackingEndsMargin=   0.5
         YAxis1.ScaleType=   1
         YAxis1.TrackingEnabled=   0
         YAxis1.TrackingStyle=   0
         YAxis1.TrackingAlignFirstStyle=   3
         YAxis1.TrackingScrollCompressMax=   0
         YAxis1.CursorUseDefaultFormat=   1
         YAxis1.CursorFormatStyle=   0
         YAxis1.CursorDateTimeFormat=   "hh:nn:ss"
         YAxis1.CursorPrecisionStyle=   0
         YAxis1.CursorPrecision=   3
         YAxis1.CursorMinLength=   5
         YAxis1.CursorMinLengthAutoAdjust=   0
         YAxis1.LegendUseDefaultFormat=   1
         YAxis1.LegendFormatStyle=   0
         YAxis1.LegendDateTimeFormat=   "hh:nn:ss"
         YAxis1.LegendPrecisionStyle=   0
         YAxis1.LegendPrecision=   3
         YAxis1.LegendMinLength=   5
         YAxis1.LegendMinLengthAutoAdjust=   0
         YAxis1.CursorScaler=   1
         YAxis1.ScrollMinMaxEnabled=   0
         YAxis1.ScrollMax=   100
         YAxis1.ScrollMin=   0
         YAxis1.RestoreValuesOnResume=   1
         YAxis1.MasterUIInput=   0
         YAxis1.CartesianStyle=   0
         YAxis1.CartesianChildRefAxisName=   "<None>"
         YAxis1.CartesianChildRefValue=   0
         YAxis1.AlignRefAxisName=   "<None>"
         YAxis1.GridLinesVisible=   1
         YAxis1.ForceStacking=   0
         DataView0.Name  =   "Data View 1"
         DataView0.Visible=   1
         DataView0.Enabled=   0
         DataView0.Layer =   100
         DataView0.PopupEnabled=   1
         DataView0.Horizontal=   0
         DataView0.ZOrder=   0
         DataView0.StartPercent=   0
         DataView0.StopPercent=   100
         DataView0.Title =   ""
         DataView0.BackgroundTransparent=   1
         DataView0.BackgroundColor=   8421376
         DataView0.GridXAxisName=   "<All>"
         DataView0.GridYAxisName=   "<All>"
         DataView0.GridShow=   1
         DataView0.GridLineColor=   32768
         DataView0.GridLineShowLeft=   1
         DataView0.GridLineShowRight=   1
         DataView0.GridLineShowTop=   1
         DataView0.GridLineShowBottom=   1
         DataView0.GridLineShowXMajors=   1
         DataView0.GridLineShowXMinors=   1
         DataView0.GridLineShowYMajors=   1
         DataView0.GridLineShowYMinors=   1
         DataView0.GridLineMajorStyle=   0
         DataView0.GridLineMinorStyle=   0
         DataView0.GridLineXMajorCustom=   0
         DataView0.GridLineXMajorColor=   32768
         DataView0.GridLineXMajorWidth=   0
         DataView0.GridLineXMajorStyle=   0
         DataView0.GridLineXMinorCustom=   0
         DataView0.GridLineXMinorColor=   32768
         DataView0.GridLineXMinorWidth=   0
         DataView0.GridLineXMinorStyle=   0
         DataView0.GridLineYMajorCustom=   0
         DataView0.GridLineYMajorColor=   32768
         DataView0.GridLineYMajorWidth=   0
         DataView0.GridLineYMajorStyle=   0
         DataView0.GridLineYMinorCustom=   0
         DataView0.GridLineYMinorColor=   32768
         DataView0.GridLineYMinorWidth=   0
         DataView0.GridLineYMinorStyle=   0
         DataView0.AxesControlEnabled=   0
         DataView0.AxesControlMouseStyle=   2
         DataView0.AxesControlWheelStyle=   0
         Channel0.Name   =   "Crono"
         Channel0.Visible=   1
         Channel0.Enabled=   1
         Channel0.Layer  =   100
         Channel0.PopupEnabled=   1
         Channel0.TitleText=   "Cronometraje"
         Channel0.VisibleInLegend=   1
         Channel0.RingBufferSize=   3600
         Channel0.TraceVisible=   1
         Channel0.Color  =   255
         Channel0.TraceLineStyle=   0
         Channel0.TraceLineWidth=   2
         Channel0.MarkersAllowIndividual=   0
         Channel0.MarkersPenUseChannelColor=   1
         Channel0.MarkersBrushUseChannelColor=   1
         Channel0.MarkersTurnOffLimit=   0
         Channel0.MarkersVisible=   0
         Channel0.MarkersSize=   3
         Channel0.MarkersStyle=   0
         Channel0.MarkersPenColor=   255
         Channel0.MarkersPenStyle=   0
         Channel0.MarkersPenWidth=   0
         Channel0.MarkersBrushColor=   255
         Channel0.MarkersBrushStyle=   0
         Channel0.MarkersFont.Charset=   1
         Channel0.MarkersFont.Color=   16777215
         Channel0.MarkersFont.Height=   -11
         Channel0.MarkersFont.Name=   "MS Sans Serif"
         Channel0.MarkersFont.Pitch=   0
         Channel0.MarkersFont.Style=   1
         Channel0.XAxisName=   "Profundidad"
         Channel0.YAxisName=   "Crono"
         Channel0.XAxisTrackingEnabled=   1
         Channel0.YAxisTrackingEnabled=   1
         Channel0.LogFileName=   ""
         Channel0.LogBufferSize=   0
         Channel0.DataStyle=   0
         Channel0.Tag    =   0
         Channel0.OPCGroupName=   ""
         Channel0.OPCComputerName=   "Local"
         Channel0.OPCServerName=   ""
         Channel0.OPCItemName=   ""
         Channel0.OPCUpdateRate=   500
         Channel0.OPCAutoConnect=   1
         Channel0.FastDrawEnabled=   1
         Channel0.InterpolationStyle=   4
         Channel0.FillEnabled=   0
         Channel0.FillReference=   0
         Channel0.FillStyle=   0
         Channel0.FillColor=   0
         Channel0.FillUseChannelColor=   1
         Channel0.DigitalEnabled=   0
         Channel0.DigitalReferenceStyle=   0
         Channel0.DigitalReferenceLow=   10
         Channel0.DigitalReferenceHigh=   90
         Channel0.HighLowStyle=   0
         Channel0.HighLowEnabled=   0
         Channel0.HighLowBarColor=   16776960
         Channel0.HighLowBarWidth=   0.5
         Channel0.HighLowOpenShow=   1
         Channel0.HighLowOpenColor=   65280
         Channel0.HighLowOpenWidth=   1
         Channel0.HighLowOpenHeight=   1
         Channel0.HighLowCloseShow=   1
         Channel0.HighLowCloseColor=   255
         Channel0.HighLowCloseWidth=   1
         Channel0.HighLowCloseHeight=   1
         Channel0.HighLowShadowColor=   8421504
         Channel0.HighLowBullishColor=   16777215
         Channel0.HighLowBearishColor=   8421504
         Channel0.BarEnabled=   0
         Channel0.BarPenUseChannelColor=   1
         Channel0.BarBrushUseChannelColor=   1
         Channel0.BarReference=   0
         Channel0.BarWidth=   5
         Channel0.BarPenColor=   255
         Channel0.BarPenWidth=   0
         Channel0.BarPenStyle=   0
         Channel0.BarBrushColor=   255
         Channel0.BarBrushStyle=   0
         Channel0.OPCXValueSource=   0
         Channel1.Name   =   "GasTotal"
         Channel1.Visible=   1
         Channel1.Enabled=   1
         Channel1.Layer  =   100
         Channel1.PopupEnabled=   0
         Channel1.TitleText=   "Gas Total"
         Channel1.VisibleInLegend=   1
         Channel1.RingBufferSize=   3600
         Channel1.TraceVisible=   1
         Channel1.Color  =   255
         Channel1.TraceLineStyle=   0
         Channel1.TraceLineWidth=   3
         Channel1.MarkersAllowIndividual=   0
         Channel1.MarkersPenUseChannelColor=   1
         Channel1.MarkersBrushUseChannelColor=   1
         Channel1.MarkersTurnOffLimit=   0
         Channel1.MarkersVisible=   0
         Channel1.MarkersSize=   3
         Channel1.MarkersStyle=   0
         Channel1.MarkersPenColor=   255
         Channel1.MarkersPenStyle=   0
         Channel1.MarkersPenWidth=   0
         Channel1.MarkersBrushColor=   255
         Channel1.MarkersBrushStyle=   0
         Channel1.MarkersFont.Charset=   1
         Channel1.MarkersFont.Color=   16777215
         Channel1.MarkersFont.Height=   -11
         Channel1.MarkersFont.Name=   "MS Sans Serif"
         Channel1.MarkersFont.Pitch=   0
         Channel1.MarkersFont.Style=   1
         Channel1.XAxisName=   "Profundidad"
         Channel1.YAxisName=   "Gas-Cromatografia"
         Channel1.XAxisTrackingEnabled=   1
         Channel1.YAxisTrackingEnabled=   1
         Channel1.LogFileName=   ""
         Channel1.LogBufferSize=   0
         Channel1.DataStyle=   0
         Channel1.Tag    =   0
         Channel1.OPCGroupName=   ""
         Channel1.OPCComputerName=   "Local"
         Channel1.OPCServerName=   ""
         Channel1.OPCItemName=   ""
         Channel1.OPCUpdateRate=   500
         Channel1.OPCAutoConnect=   1
         Channel1.FastDrawEnabled=   1
         Channel1.InterpolationStyle=   0
         Channel1.FillEnabled=   0
         Channel1.FillReference=   0
         Channel1.FillStyle=   0
         Channel1.FillColor=   0
         Channel1.FillUseChannelColor=   1
         Channel1.DigitalEnabled=   0
         Channel1.DigitalReferenceStyle=   0
         Channel1.DigitalReferenceLow=   10
         Channel1.DigitalReferenceHigh=   90
         Channel1.HighLowStyle=   0
         Channel1.HighLowEnabled=   0
         Channel1.HighLowBarColor=   16776960
         Channel1.HighLowBarWidth=   0.5
         Channel1.HighLowOpenShow=   1
         Channel1.HighLowOpenColor=   65280
         Channel1.HighLowOpenWidth=   1
         Channel1.HighLowOpenHeight=   1
         Channel1.HighLowCloseShow=   1
         Channel1.HighLowCloseColor=   255
         Channel1.HighLowCloseWidth=   1
         Channel1.HighLowCloseHeight=   1
         Channel1.HighLowShadowColor=   8421504
         Channel1.HighLowBullishColor=   16777215
         Channel1.HighLowBearishColor=   8421504
         Channel1.BarEnabled=   0
         Channel1.BarPenUseChannelColor=   1
         Channel1.BarBrushUseChannelColor=   1
         Channel1.BarReference=   0
         Channel1.BarWidth=   5
         Channel1.BarPenColor=   255
         Channel1.BarPenWidth=   0
         Channel1.BarPenStyle=   0
         Channel1.BarBrushColor=   255
         Channel1.BarBrushStyle=   0
         Channel1.OPCXValueSource=   0
         Channel2.Name   =   "CO2"
         Channel2.Visible=   1
         Channel2.Enabled=   1
         Channel2.Layer  =   100
         Channel2.PopupEnabled=   0
         Channel2.TitleText=   "Dioxido de Carbono"
         Channel2.VisibleInLegend=   1
         Channel2.RingBufferSize=   3600
         Channel2.TraceVisible=   1
         Channel2.Color  =   16711680
         Channel2.TraceLineStyle=   4
         Channel2.TraceLineWidth=   1
         Channel2.MarkersAllowIndividual=   1
         Channel2.MarkersPenUseChannelColor=   1
         Channel2.MarkersBrushUseChannelColor=   1
         Channel2.MarkersTurnOffLimit=   0
         Channel2.MarkersVisible=   1
         Channel2.MarkersSize=   1
         Channel2.MarkersStyle=   0
         Channel2.MarkersPenColor=   255
         Channel2.MarkersPenStyle=   0
         Channel2.MarkersPenWidth=   0
         Channel2.MarkersBrushColor=   255
         Channel2.MarkersBrushStyle=   0
         Channel2.MarkersFont.Charset=   1
         Channel2.MarkersFont.Color=   16777215
         Channel2.MarkersFont.Height=   -11
         Channel2.MarkersFont.Name=   "MS Sans Serif"
         Channel2.MarkersFont.Pitch=   0
         Channel2.MarkersFont.Style=   1
         Channel2.XAxisName=   "Profundidad"
         Channel2.YAxisName=   "Gas-Cromatografia"
         Channel2.XAxisTrackingEnabled=   1
         Channel2.YAxisTrackingEnabled=   1
         Channel2.LogFileName=   ""
         Channel2.LogBufferSize=   0
         Channel2.DataStyle=   0
         Channel2.Tag    =   0
         Channel2.OPCGroupName=   ""
         Channel2.OPCComputerName=   "Local"
         Channel2.OPCServerName=   ""
         Channel2.OPCItemName=   ""
         Channel2.OPCUpdateRate=   500
         Channel2.OPCAutoConnect=   1
         Channel2.FastDrawEnabled=   1
         Channel2.InterpolationStyle=   0
         Channel2.FillEnabled=   0
         Channel2.FillReference=   0
         Channel2.FillStyle=   0
         Channel2.FillColor=   0
         Channel2.FillUseChannelColor=   1
         Channel2.DigitalEnabled=   0
         Channel2.DigitalReferenceStyle=   0
         Channel2.DigitalReferenceLow=   10
         Channel2.DigitalReferenceHigh=   90
         Channel2.HighLowStyle=   0
         Channel2.HighLowEnabled=   0
         Channel2.HighLowBarColor=   16776960
         Channel2.HighLowBarWidth=   0.5
         Channel2.HighLowOpenShow=   1
         Channel2.HighLowOpenColor=   65280
         Channel2.HighLowOpenWidth=   1
         Channel2.HighLowOpenHeight=   1
         Channel2.HighLowCloseShow=   1
         Channel2.HighLowCloseColor=   255
         Channel2.HighLowCloseWidth=   1
         Channel2.HighLowCloseHeight=   1
         Channel2.HighLowShadowColor=   8421504
         Channel2.HighLowBullishColor=   16777215
         Channel2.HighLowBearishColor=   8421504
         Channel2.BarEnabled=   0
         Channel2.BarPenUseChannelColor=   1
         Channel2.BarBrushUseChannelColor=   1
         Channel2.BarReference=   0
         Channel2.BarWidth=   5
         Channel2.BarPenColor=   255
         Channel2.BarPenWidth=   0
         Channel2.BarPenStyle=   0
         Channel2.BarBrushColor=   255
         Channel2.BarBrushStyle=   0
         Channel2.OPCXValueSource=   0
         Channel3.Name   =   "Metano"
         Channel3.Visible=   1
         Channel3.Enabled=   1
         Channel3.Layer  =   100
         Channel3.PopupEnabled=   0
         Channel3.TitleText=   "Metano"
         Channel3.VisibleInLegend=   1
         Channel3.RingBufferSize=   3600
         Channel3.TraceVisible=   1
         Channel3.Color  =   65535
         Channel3.TraceLineStyle=   0
         Channel3.TraceLineWidth=   2
         Channel3.MarkersAllowIndividual=   0
         Channel3.MarkersPenUseChannelColor=   1
         Channel3.MarkersBrushUseChannelColor=   1
         Channel3.MarkersTurnOffLimit=   0
         Channel3.MarkersVisible=   0
         Channel3.MarkersSize=   3
         Channel3.MarkersStyle=   0
         Channel3.MarkersPenColor=   255
         Channel3.MarkersPenStyle=   0
         Channel3.MarkersPenWidth=   0
         Channel3.MarkersBrushColor=   255
         Channel3.MarkersBrushStyle=   0
         Channel3.MarkersFont.Charset=   1
         Channel3.MarkersFont.Color=   16777215
         Channel3.MarkersFont.Height=   -11
         Channel3.MarkersFont.Name=   "MS Sans Serif"
         Channel3.MarkersFont.Pitch=   0
         Channel3.MarkersFont.Style=   1
         Channel3.XAxisName=   "Profundidad"
         Channel3.YAxisName=   "Gas-Cromatografia"
         Channel3.XAxisTrackingEnabled=   1
         Channel3.YAxisTrackingEnabled=   1
         Channel3.LogFileName=   ""
         Channel3.LogBufferSize=   0
         Channel3.DataStyle=   0
         Channel3.Tag    =   0
         Channel3.OPCGroupName=   ""
         Channel3.OPCComputerName=   "Local"
         Channel3.OPCServerName=   ""
         Channel3.OPCItemName=   ""
         Channel3.OPCUpdateRate=   500
         Channel3.OPCAutoConnect=   1
         Channel3.FastDrawEnabled=   1
         Channel3.InterpolationStyle=   0
         Channel3.FillEnabled=   0
         Channel3.FillReference=   0
         Channel3.FillStyle=   0
         Channel3.FillColor=   0
         Channel3.FillUseChannelColor=   1
         Channel3.DigitalEnabled=   0
         Channel3.DigitalReferenceStyle=   0
         Channel3.DigitalReferenceLow=   10
         Channel3.DigitalReferenceHigh=   90
         Channel3.HighLowStyle=   0
         Channel3.HighLowEnabled=   0
         Channel3.HighLowBarColor=   16776960
         Channel3.HighLowBarWidth=   0.5
         Channel3.HighLowOpenShow=   1
         Channel3.HighLowOpenColor=   65280
         Channel3.HighLowOpenWidth=   1
         Channel3.HighLowOpenHeight=   1
         Channel3.HighLowCloseShow=   1
         Channel3.HighLowCloseColor=   255
         Channel3.HighLowCloseWidth=   1
         Channel3.HighLowCloseHeight=   1
         Channel3.HighLowShadowColor=   8421504
         Channel3.HighLowBullishColor=   16777215
         Channel3.HighLowBearishColor=   8421504
         Channel3.BarEnabled=   0
         Channel3.BarPenUseChannelColor=   1
         Channel3.BarBrushUseChannelColor=   1
         Channel3.BarReference=   0
         Channel3.BarWidth=   5
         Channel3.BarPenColor=   255
         Channel3.BarPenWidth=   0
         Channel3.BarPenStyle=   0
         Channel3.BarBrushColor=   255
         Channel3.BarBrushStyle=   0
         Channel3.OPCXValueSource=   0
         Channel4.Name   =   "Etano"
         Channel4.Visible=   1
         Channel4.Enabled=   1
         Channel4.Layer  =   100
         Channel4.PopupEnabled=   0
         Channel4.TitleText=   "Etano"
         Channel4.VisibleInLegend=   1
         Channel4.RingBufferSize=   3600
         Channel4.TraceVisible=   1
         Channel4.Color  =   33023
         Channel4.TraceLineStyle=   0
         Channel4.TraceLineWidth=   2
         Channel4.MarkersAllowIndividual=   0
         Channel4.MarkersPenUseChannelColor=   1
         Channel4.MarkersBrushUseChannelColor=   1
         Channel4.MarkersTurnOffLimit=   0
         Channel4.MarkersVisible=   0
         Channel4.MarkersSize=   3
         Channel4.MarkersStyle=   0
         Channel4.MarkersPenColor=   255
         Channel4.MarkersPenStyle=   0
         Channel4.MarkersPenWidth=   0
         Channel4.MarkersBrushColor=   255
         Channel4.MarkersBrushStyle=   0
         Channel4.MarkersFont.Charset=   1
         Channel4.MarkersFont.Color=   16777215
         Channel4.MarkersFont.Height=   -11
         Channel4.MarkersFont.Name=   "MS Sans Serif"
         Channel4.MarkersFont.Pitch=   0
         Channel4.MarkersFont.Style=   1
         Channel4.XAxisName=   "Profundidad"
         Channel4.YAxisName=   "Gas-Cromatografia"
         Channel4.XAxisTrackingEnabled=   1
         Channel4.YAxisTrackingEnabled=   1
         Channel4.LogFileName=   ""
         Channel4.LogBufferSize=   0
         Channel4.DataStyle=   0
         Channel4.Tag    =   0
         Channel4.OPCGroupName=   ""
         Channel4.OPCComputerName=   "Local"
         Channel4.OPCServerName=   ""
         Channel4.OPCItemName=   ""
         Channel4.OPCUpdateRate=   500
         Channel4.OPCAutoConnect=   1
         Channel4.FastDrawEnabled=   1
         Channel4.InterpolationStyle=   0
         Channel4.FillEnabled=   0
         Channel4.FillReference=   0
         Channel4.FillStyle=   0
         Channel4.FillColor=   0
         Channel4.FillUseChannelColor=   1
         Channel4.DigitalEnabled=   0
         Channel4.DigitalReferenceStyle=   0
         Channel4.DigitalReferenceLow=   10
         Channel4.DigitalReferenceHigh=   90
         Channel4.HighLowStyle=   0
         Channel4.HighLowEnabled=   0
         Channel4.HighLowBarColor=   16776960
         Channel4.HighLowBarWidth=   0.5
         Channel4.HighLowOpenShow=   1
         Channel4.HighLowOpenColor=   65280
         Channel4.HighLowOpenWidth=   1
         Channel4.HighLowOpenHeight=   1
         Channel4.HighLowCloseShow=   1
         Channel4.HighLowCloseColor=   255
         Channel4.HighLowCloseWidth=   1
         Channel4.HighLowCloseHeight=   1
         Channel4.HighLowShadowColor=   8421504
         Channel4.HighLowBullishColor=   16777215
         Channel4.HighLowBearishColor=   8421504
         Channel4.BarEnabled=   0
         Channel4.BarPenUseChannelColor=   1
         Channel4.BarBrushUseChannelColor=   1
         Channel4.BarReference=   0
         Channel4.BarWidth=   5
         Channel4.BarPenColor=   255
         Channel4.BarPenWidth=   0
         Channel4.BarPenStyle=   0
         Channel4.BarBrushColor=   255
         Channel4.BarBrushStyle=   0
         Channel4.OPCXValueSource=   0
         Channel5.Name   =   "Propano"
         Channel5.Visible=   1
         Channel5.Enabled=   1
         Channel5.Layer  =   100
         Channel5.PopupEnabled=   0
         Channel5.TitleText=   "Propano"
         Channel5.VisibleInLegend=   1
         Channel5.RingBufferSize=   3600
         Channel5.TraceVisible=   1
         Channel5.Color  =   16776960
         Channel5.TraceLineStyle=   0
         Channel5.TraceLineWidth=   2
         Channel5.MarkersAllowIndividual=   0
         Channel5.MarkersPenUseChannelColor=   1
         Channel5.MarkersBrushUseChannelColor=   1
         Channel5.MarkersTurnOffLimit=   0
         Channel5.MarkersVisible=   0
         Channel5.MarkersSize=   3
         Channel5.MarkersStyle=   0
         Channel5.MarkersPenColor=   255
         Channel5.MarkersPenStyle=   0
         Channel5.MarkersPenWidth=   0
         Channel5.MarkersBrushColor=   255
         Channel5.MarkersBrushStyle=   0
         Channel5.MarkersFont.Charset=   1
         Channel5.MarkersFont.Color=   16777215
         Channel5.MarkersFont.Height=   -11
         Channel5.MarkersFont.Name=   "MS Sans Serif"
         Channel5.MarkersFont.Pitch=   0
         Channel5.MarkersFont.Style=   1
         Channel5.XAxisName=   "Profundidad"
         Channel5.YAxisName=   "Gas-Cromatografia"
         Channel5.XAxisTrackingEnabled=   1
         Channel5.YAxisTrackingEnabled=   1
         Channel5.LogFileName=   ""
         Channel5.LogBufferSize=   0
         Channel5.DataStyle=   0
         Channel5.Tag    =   0
         Channel5.OPCGroupName=   ""
         Channel5.OPCComputerName=   "Local"
         Channel5.OPCServerName=   ""
         Channel5.OPCItemName=   ""
         Channel5.OPCUpdateRate=   500
         Channel5.OPCAutoConnect=   1
         Channel5.FastDrawEnabled=   1
         Channel5.InterpolationStyle=   0
         Channel5.FillEnabled=   0
         Channel5.FillReference=   0
         Channel5.FillStyle=   0
         Channel5.FillColor=   0
         Channel5.FillUseChannelColor=   1
         Channel5.DigitalEnabled=   0
         Channel5.DigitalReferenceStyle=   0
         Channel5.DigitalReferenceLow=   10
         Channel5.DigitalReferenceHigh=   90
         Channel5.HighLowStyle=   0
         Channel5.HighLowEnabled=   0
         Channel5.HighLowBarColor=   16776960
         Channel5.HighLowBarWidth=   0.5
         Channel5.HighLowOpenShow=   1
         Channel5.HighLowOpenColor=   65280
         Channel5.HighLowOpenWidth=   1
         Channel5.HighLowOpenHeight=   1
         Channel5.HighLowCloseShow=   1
         Channel5.HighLowCloseColor=   255
         Channel5.HighLowCloseWidth=   1
         Channel5.HighLowCloseHeight=   1
         Channel5.HighLowShadowColor=   8421504
         Channel5.HighLowBullishColor=   16777215
         Channel5.HighLowBearishColor=   8421504
         Channel5.BarEnabled=   0
         Channel5.BarPenUseChannelColor=   1
         Channel5.BarBrushUseChannelColor=   1
         Channel5.BarReference=   0
         Channel5.BarWidth=   5
         Channel5.BarPenColor=   255
         Channel5.BarPenWidth=   0
         Channel5.BarPenStyle=   0
         Channel5.BarBrushColor=   255
         Channel5.BarBrushStyle=   0
         Channel5.OPCXValueSource=   0
         Channel6.Name   =   "IsoButano"
         Channel6.Visible=   1
         Channel6.Enabled=   1
         Channel6.Layer  =   100
         Channel6.PopupEnabled=   0
         Channel6.TitleText=   "IsoButano"
         Channel6.VisibleInLegend=   1
         Channel6.RingBufferSize=   3600
         Channel6.TraceVisible=   1
         Channel6.Color  =   32768
         Channel6.TraceLineStyle=   0
         Channel6.TraceLineWidth=   2
         Channel6.MarkersAllowIndividual=   0
         Channel6.MarkersPenUseChannelColor=   1
         Channel6.MarkersBrushUseChannelColor=   1
         Channel6.MarkersTurnOffLimit=   0
         Channel6.MarkersVisible=   0
         Channel6.MarkersSize=   3
         Channel6.MarkersStyle=   0
         Channel6.MarkersPenColor=   255
         Channel6.MarkersPenStyle=   0
         Channel6.MarkersPenWidth=   0
         Channel6.MarkersBrushColor=   255
         Channel6.MarkersBrushStyle=   0
         Channel6.MarkersFont.Charset=   1
         Channel6.MarkersFont.Color=   16777215
         Channel6.MarkersFont.Height=   -11
         Channel6.MarkersFont.Name=   "MS Sans Serif"
         Channel6.MarkersFont.Pitch=   0
         Channel6.MarkersFont.Style=   1
         Channel6.XAxisName=   "Profundidad"
         Channel6.YAxisName=   "Gas-Cromatografia"
         Channel6.XAxisTrackingEnabled=   1
         Channel6.YAxisTrackingEnabled=   1
         Channel6.LogFileName=   ""
         Channel6.LogBufferSize=   0
         Channel6.DataStyle=   0
         Channel6.Tag    =   0
         Channel6.OPCGroupName=   ""
         Channel6.OPCComputerName=   "Local"
         Channel6.OPCServerName=   ""
         Channel6.OPCItemName=   ""
         Channel6.OPCUpdateRate=   500
         Channel6.OPCAutoConnect=   1
         Channel6.FastDrawEnabled=   1
         Channel6.InterpolationStyle=   0
         Channel6.FillEnabled=   0
         Channel6.FillReference=   0
         Channel6.FillStyle=   0
         Channel6.FillColor=   0
         Channel6.FillUseChannelColor=   1
         Channel6.DigitalEnabled=   0
         Channel6.DigitalReferenceStyle=   0
         Channel6.DigitalReferenceLow=   10
         Channel6.DigitalReferenceHigh=   90
         Channel6.HighLowStyle=   0
         Channel6.HighLowEnabled=   0
         Channel6.HighLowBarColor=   16776960
         Channel6.HighLowBarWidth=   0.5
         Channel6.HighLowOpenShow=   1
         Channel6.HighLowOpenColor=   65280
         Channel6.HighLowOpenWidth=   1
         Channel6.HighLowOpenHeight=   1
         Channel6.HighLowCloseShow=   1
         Channel6.HighLowCloseColor=   255
         Channel6.HighLowCloseWidth=   1
         Channel6.HighLowCloseHeight=   1
         Channel6.HighLowShadowColor=   8421504
         Channel6.HighLowBullishColor=   16777215
         Channel6.HighLowBearishColor=   8421504
         Channel6.BarEnabled=   0
         Channel6.BarPenUseChannelColor=   1
         Channel6.BarBrushUseChannelColor=   1
         Channel6.BarReference=   0
         Channel6.BarWidth=   5
         Channel6.BarPenColor=   255
         Channel6.BarPenWidth=   0
         Channel6.BarPenStyle=   0
         Channel6.BarBrushColor=   255
         Channel6.BarBrushStyle=   0
         Channel6.OPCXValueSource=   0
         Channel7.Name   =   "NormalButano"
         Channel7.Visible=   1
         Channel7.Enabled=   1
         Channel7.Layer  =   100
         Channel7.PopupEnabled=   0
         Channel7.TitleText=   "NormalButano"
         Channel7.VisibleInLegend=   1
         Channel7.RingBufferSize=   3600
         Channel7.TraceVisible=   1
         Channel7.Color  =   128
         Channel7.TraceLineStyle=   0
         Channel7.TraceLineWidth=   2
         Channel7.MarkersAllowIndividual=   0
         Channel7.MarkersPenUseChannelColor=   1
         Channel7.MarkersBrushUseChannelColor=   1
         Channel7.MarkersTurnOffLimit=   0
         Channel7.MarkersVisible=   0
         Channel7.MarkersSize=   3
         Channel7.MarkersStyle=   0
         Channel7.MarkersPenColor=   255
         Channel7.MarkersPenStyle=   0
         Channel7.MarkersPenWidth=   0
         Channel7.MarkersBrushColor=   255
         Channel7.MarkersBrushStyle=   0
         Channel7.MarkersFont.Charset=   1
         Channel7.MarkersFont.Color=   16777215
         Channel7.MarkersFont.Height=   -11
         Channel7.MarkersFont.Name=   "MS Sans Serif"
         Channel7.MarkersFont.Pitch=   0
         Channel7.MarkersFont.Style=   1
         Channel7.XAxisName=   "Profundidad"
         Channel7.YAxisName=   "Gas-Cromatografia"
         Channel7.XAxisTrackingEnabled=   1
         Channel7.YAxisTrackingEnabled=   1
         Channel7.LogFileName=   ""
         Channel7.LogBufferSize=   0
         Channel7.DataStyle=   0
         Channel7.Tag    =   0
         Channel7.OPCGroupName=   ""
         Channel7.OPCComputerName=   "Local"
         Channel7.OPCServerName=   ""
         Channel7.OPCItemName=   ""
         Channel7.OPCUpdateRate=   500
         Channel7.OPCAutoConnect=   1
         Channel7.FastDrawEnabled=   1
         Channel7.InterpolationStyle=   0
         Channel7.FillEnabled=   0
         Channel7.FillReference=   0
         Channel7.FillStyle=   0
         Channel7.FillColor=   0
         Channel7.FillUseChannelColor=   1
         Channel7.DigitalEnabled=   0
         Channel7.DigitalReferenceStyle=   0
         Channel7.DigitalReferenceLow=   10
         Channel7.DigitalReferenceHigh=   90
         Channel7.HighLowStyle=   0
         Channel7.HighLowEnabled=   0
         Channel7.HighLowBarColor=   16776960
         Channel7.HighLowBarWidth=   0.5
         Channel7.HighLowOpenShow=   1
         Channel7.HighLowOpenColor=   65280
         Channel7.HighLowOpenWidth=   1
         Channel7.HighLowOpenHeight=   1
         Channel7.HighLowCloseShow=   1
         Channel7.HighLowCloseColor=   255
         Channel7.HighLowCloseWidth=   1
         Channel7.HighLowCloseHeight=   1
         Channel7.HighLowShadowColor=   8421504
         Channel7.HighLowBullishColor=   16777215
         Channel7.HighLowBearishColor=   8421504
         Channel7.BarEnabled=   0
         Channel7.BarPenUseChannelColor=   1
         Channel7.BarBrushUseChannelColor=   1
         Channel7.BarReference=   0
         Channel7.BarWidth=   5
         Channel7.BarPenColor=   255
         Channel7.BarPenWidth=   0
         Channel7.BarPenStyle=   0
         Channel7.BarBrushColor=   255
         Channel7.BarBrushStyle=   0
         Channel7.OPCXValueSource=   0
         Channel8.Name   =   "IsoPentano"
         Channel8.Visible=   1
         Channel8.Enabled=   1
         Channel8.Layer  =   100
         Channel8.PopupEnabled=   0
         Channel8.TitleText=   "IsoPentano"
         Channel8.VisibleInLegend=   1
         Channel8.RingBufferSize=   3600
         Channel8.TraceVisible=   1
         Channel8.Color  =   16711808
         Channel8.TraceLineStyle=   0
         Channel8.TraceLineWidth=   2
         Channel8.MarkersAllowIndividual=   0
         Channel8.MarkersPenUseChannelColor=   1
         Channel8.MarkersBrushUseChannelColor=   1
         Channel8.MarkersTurnOffLimit=   0
         Channel8.MarkersVisible=   0
         Channel8.MarkersSize=   3
         Channel8.MarkersStyle=   0
         Channel8.MarkersPenColor=   255
         Channel8.MarkersPenStyle=   0
         Channel8.MarkersPenWidth=   0
         Channel8.MarkersBrushColor=   255
         Channel8.MarkersBrushStyle=   0
         Channel8.MarkersFont.Charset=   1
         Channel8.MarkersFont.Color=   16777215
         Channel8.MarkersFont.Height=   -11
         Channel8.MarkersFont.Name=   "MS Sans Serif"
         Channel8.MarkersFont.Pitch=   0
         Channel8.MarkersFont.Style=   1
         Channel8.XAxisName=   "Profundidad"
         Channel8.YAxisName=   "Gas-Cromatografia"
         Channel8.XAxisTrackingEnabled=   1
         Channel8.YAxisTrackingEnabled=   1
         Channel8.LogFileName=   ""
         Channel8.LogBufferSize=   0
         Channel8.DataStyle=   0
         Channel8.Tag    =   0
         Channel8.OPCGroupName=   ""
         Channel8.OPCComputerName=   "Local"
         Channel8.OPCServerName=   ""
         Channel8.OPCItemName=   ""
         Channel8.OPCUpdateRate=   500
         Channel8.OPCAutoConnect=   1
         Channel8.FastDrawEnabled=   1
         Channel8.InterpolationStyle=   0
         Channel8.FillEnabled=   0
         Channel8.FillReference=   0
         Channel8.FillStyle=   0
         Channel8.FillColor=   0
         Channel8.FillUseChannelColor=   1
         Channel8.DigitalEnabled=   0
         Channel8.DigitalReferenceStyle=   0
         Channel8.DigitalReferenceLow=   10
         Channel8.DigitalReferenceHigh=   90
         Channel8.HighLowStyle=   0
         Channel8.HighLowEnabled=   0
         Channel8.HighLowBarColor=   16776960
         Channel8.HighLowBarWidth=   0.5
         Channel8.HighLowOpenShow=   1
         Channel8.HighLowOpenColor=   65280
         Channel8.HighLowOpenWidth=   1
         Channel8.HighLowOpenHeight=   1
         Channel8.HighLowCloseShow=   1
         Channel8.HighLowCloseColor=   255
         Channel8.HighLowCloseWidth=   1
         Channel8.HighLowCloseHeight=   1
         Channel8.HighLowShadowColor=   8421504
         Channel8.HighLowBullishColor=   16777215
         Channel8.HighLowBearishColor=   8421504
         Channel8.BarEnabled=   0
         Channel8.BarPenUseChannelColor=   1
         Channel8.BarBrushUseChannelColor=   1
         Channel8.BarReference=   0
         Channel8.BarWidth=   5
         Channel8.BarPenColor=   255
         Channel8.BarPenWidth=   0
         Channel8.BarPenStyle=   0
         Channel8.BarBrushColor=   255
         Channel8.BarBrushStyle=   0
         Channel8.OPCXValueSource=   0
         Channel9.Name   =   "NormalPentano"
         Channel9.Visible=   1
         Channel9.Enabled=   1
         Channel9.Layer  =   100
         Channel9.PopupEnabled=   0
         Channel9.TitleText=   "NormalPentano"
         Channel9.VisibleInLegend=   1
         Channel9.RingBufferSize=   3600
         Channel9.TraceVisible=   1
         Channel9.Color  =   8421440
         Channel9.TraceLineStyle=   0
         Channel9.TraceLineWidth=   2
         Channel9.MarkersAllowIndividual=   0
         Channel9.MarkersPenUseChannelColor=   1
         Channel9.MarkersBrushUseChannelColor=   1
         Channel9.MarkersTurnOffLimit=   0
         Channel9.MarkersVisible=   0
         Channel9.MarkersSize=   3
         Channel9.MarkersStyle=   0
         Channel9.MarkersPenColor=   255
         Channel9.MarkersPenStyle=   0
         Channel9.MarkersPenWidth=   0
         Channel9.MarkersBrushColor=   255
         Channel9.MarkersBrushStyle=   0
         Channel9.MarkersFont.Charset=   1
         Channel9.MarkersFont.Color=   16777215
         Channel9.MarkersFont.Height=   -11
         Channel9.MarkersFont.Name=   "MS Sans Serif"
         Channel9.MarkersFont.Pitch=   0
         Channel9.MarkersFont.Style=   1
         Channel9.XAxisName=   "Profundidad"
         Channel9.YAxisName=   "Gas-Cromatografia"
         Channel9.XAxisTrackingEnabled=   1
         Channel9.YAxisTrackingEnabled=   1
         Channel9.LogFileName=   ""
         Channel9.LogBufferSize=   0
         Channel9.DataStyle=   0
         Channel9.Tag    =   0
         Channel9.OPCGroupName=   ""
         Channel9.OPCComputerName=   "Local"
         Channel9.OPCServerName=   ""
         Channel9.OPCItemName=   ""
         Channel9.OPCUpdateRate=   500
         Channel9.OPCAutoConnect=   1
         Channel9.FastDrawEnabled=   1
         Channel9.InterpolationStyle=   0
         Channel9.FillEnabled=   0
         Channel9.FillReference=   0
         Channel9.FillStyle=   0
         Channel9.FillColor=   0
         Channel9.FillUseChannelColor=   1
         Channel9.DigitalEnabled=   0
         Channel9.DigitalReferenceStyle=   0
         Channel9.DigitalReferenceLow=   10
         Channel9.DigitalReferenceHigh=   90
         Channel9.HighLowStyle=   0
         Channel9.HighLowEnabled=   0
         Channel9.HighLowBarColor=   16776960
         Channel9.HighLowBarWidth=   0.5
         Channel9.HighLowOpenShow=   1
         Channel9.HighLowOpenColor=   65280
         Channel9.HighLowOpenWidth=   1
         Channel9.HighLowOpenHeight=   1
         Channel9.HighLowCloseShow=   1
         Channel9.HighLowCloseColor=   255
         Channel9.HighLowCloseWidth=   1
         Channel9.HighLowCloseHeight=   1
         Channel9.HighLowShadowColor=   8421504
         Channel9.HighLowBullishColor=   16777215
         Channel9.HighLowBearishColor=   8421504
         Channel9.BarEnabled=   0
         Channel9.BarPenUseChannelColor=   1
         Channel9.BarBrushUseChannelColor=   1
         Channel9.BarReference=   0
         Channel9.BarWidth=   5
         Channel9.BarPenColor=   255
         Channel9.BarPenWidth=   0
         Channel9.BarPenStyle=   0
         Channel9.BarBrushColor=   255
         Channel9.BarBrushStyle=   0
         Channel9.OPCXValueSource=   0
         DataCursor0.Name=   "Cursor 1"
         DataCursor0.Visible=   0
         DataCursor0.Enabled=   1
         DataCursor0.Layer=   100
         DataCursor0.PopupEnabled=   0
         DataCursor0.ChannelName=   "Crono"
         DataCursor0.ChannelAllowAll=   1
         DataCursor0.ChannelShowAllInLegend=   1
         DataCursor0.Style=   0
         DataCursor0.Font.Charset=   1
         DataCursor0.Font.Color=   -2147483640
         DataCursor0.Font.Height=   -11
         DataCursor0.Font.Name=   "MS Sans Serif"
         DataCursor0.Font.Pitch=   0
         DataCursor0.Font.Style=   0
         DataCursor0.Color=   65535
         DataCursor0.UseChannelColor=   1
         DataCursor0.HintShow=   1
         DataCursor0.HintHideOnRelease=   0
         DataCursor0.HintOrientationSide=   0
         DataCursor0.HintPosition=   50
         DataCursor0.Pointer1Position=   50
         DataCursor0.Pointer2Position=   60
         DataCursor0.PointerPenWidth=   1
         DataCursor0.MenuUserCanChangeOptions=   1
         DataCursor0.MenuItemVisibleValueXY=   1
         DataCursor0.MenuItemVisibleValueX=   1
         DataCursor0.MenuItemVisibleValueY=   1
         DataCursor0.MenuItemVisibleDeltaX=   1
         DataCursor0.MenuItemVisibleDeltaY=   1
         DataCursor0.MenuItemVisibleInverseDeltaX=   1
         DataCursor0.MenuItemCaptionValueXY=   "Value X-Y"
         DataCursor0.MenuItemCaptionValueX=   "Value X"
         DataCursor0.MenuItemCaptionValueY=   "Value Y"
         DataCursor0.MenuItemCaptionDeltaX=   "Period"
         DataCursor0.MenuItemCaptionDeltaY=   "Peak-Peak"
         DataCursor0.MenuItemCaptionInverseDeltaX=   "Frequency"
         Label0.Name     =   "Title"
         Label0.Visible  =   0
         Label0.Enabled  =   1
         Label0.Layer    =   100
         Label0.PopupEnabled=   1
         Label0.Horizontal=   1
         Label0.ZOrder   =   2
         Label0.StartPercent=   0
         Label0.StopPercent=   100
         Label0.MarginLeft=   0
         Label0.MarginTop=   0
         Label0.MarginRight=   0
         Label0.MarginBottom=   0.25
         Label0.Caption  =   "Untitled"
         Label0.Alignment=   0
         Label0.Font.Charset=   1
         Label0.Font.Color=   16777215
         Label0.Font.Height=   -19
         Label0.Font.Name=   "Arial"
         Label0.Font.Pitch=   0
         Label0.Font.Style=   1
      End
      Begin VB.Label Label70 
         AutoSize        =   -1  'True
         Caption         =   "Comentario"
         Height          =   210
         Left            =   75
         TabIndex        =   164
         Top             =   8895
         Width           =   810
      End
      Begin VB.Label lblHora 
         AutoSize        =   -1  'True
         Caption         =   "Hora"
         Height          =   210
         Left            =   4515
         TabIndex        =   163
         Top             =   8415
         Width           =   345
      End
      Begin VB.Label lblFecha 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         Height          =   210
         Left            =   2595
         TabIndex        =   162
         Top             =   8415
         Width           =   450
      End
      Begin VB.Label Label16 
         Caption         =   "Profundidad Inicio"
         Height          =   255
         Left            =   -73770
         TabIndex        =   149
         Top             =   8865
         Width           =   1605
      End
      Begin VB.Label Label15 
         Caption         =   "Rango"
         Height          =   255
         Left            =   -71250
         TabIndex        =   148
         Top             =   8865
         Width           =   615
      End
   End
   Begin VB.TextBox TxtNumeroComponenteAnterior 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7245
      TabIndex        =   13
      Top             =   11715
      Width           =   960
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   -30
      Top             =   11010
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAnalisis.frx":9A57
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAnalisis.frx":9EAB
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAnalisis.frx":A2FF
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAnalisis.frx":A753
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAnalisis.frx":ABA7
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtPopupElegido 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2115
      TabIndex        =   5
      Top             =   11535
      Width           =   1380
   End
   Begin VB.Label Label52 
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6735
      TabIndex        =   31
      Top             =   135
      Width           =   255
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SH"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6345
      TabIndex        =   30
      Top             =   15
      Width           =   345
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3645
      TabIndex        =   29
      Top             =   120
      Width           =   165
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CO"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3255
      TabIndex        =   28
      Top             =   15
      Width           =   375
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GAS TOTAL"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   75
      TabIndex        =   23
      Top             =   15
      Width           =   1440
   End
   Begin VB.Label Label59 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Numero componente anterior"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   5175
      TabIndex        =   11
      Top             =   11760
      Width           =   2055
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Variables"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   4920
      TabIndex        =   9
      Top             =   11400
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Popup Elegido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   990
      TabIndex        =   7
      Top             =   11610
      Width           =   1035
   End
End
Attribute VB_Name = "FrmAnalisis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Private WithEvents pObjAdminEventos         As clsAdminEventos
Attribute pObjAdminEventos.VB_VarHelpID = -1
Private LoadOk                              As Boolean
Private BotonDerecho                        As Boolean
Private pMenuChequeado                      As Long
Private ContadorTicks                       As Long
Public ChkTiempoReal                        As Long

Private Sub BtnPrint_Click()

Dim CadenaId As String
Dim AuxHeight As Long
Dim AuxWidth As Long
Dim Indice As Long

    
    AuxWidth = iPlotGases.Width
    AuxHeight = iPlotGases.Height

    iPlotGases.Height = 16000
    iPlotGases.Width = 12000



iPlotGases.BeginUpdate

iPlotGases.PrintOrientation = poPortrait
'iPlotGases.PrintOrientation = poLandscape
iPlotGases.PrintShowDialog = False

iPlotGases.PrintMarginRight = 0
iPlotGases.PrintMarginLeft = 0
iPlotGases.PrintMarginTop = 0
iPlotGases.PrintMarginBottom = 0




'Set background colors for Chart and DataView areas to a light color or white
iPlotGases.BackGroundColor = vbWhite

For Indice = 0 To iPlotGases.YAxisCount - 1

    iPlotGases.YAxis(Indice).TitleFontColor = vbBlack
    iPlotGases.YAxis(Indice).ScaleLinesColor = vbBlack
    
Next

For Indice = 0 To iPlotGases.XAxisCount - 1

    iPlotGases.XAxis(Indice).LabelsFontColor = vbBlack
    
Next

For Indice = 0 To iPlotGases.YAxisCount - 1
    iPlotGases.YAxis(Indice).TitleFontColor = vbBlack
    iPlotGases.YAxis(Indice).ScaleLinesColor = vbBlack
    iPlotGases.YAxis(Indice).LabelsFontColor = vbBlack
Next



iPlotGases.Height = 16000


For Indice = 1 To iPlotGases.AnnotationCount
 
 iPlotGases.Annotation(Indice - 1).FontColor = vbBlack
 
Next


iPlotGases.PrintChart


iPlotGases.Width = AuxWidth
iPlotGases.Height = AuxHeight




'Set background colors back to their original settings
iPlotGases.BackGroundColor = vbBlack


For Indice = 0 To iPlotGases.XAxisCount - 1
    iPlotGases.XAxis(Indice).LabelsFontColor = vbWhite
Next


For Indice = 0 To iPlotGases.YAxisCount - 1
    iPlotGases.YAxis(Indice).TitleFontColor = vbWhite
    iPlotGases.YAxis(Indice).ScaleLinesColor = vbWhite
    iPlotGases.YAxis(Indice).LabelsFontColor = vbWhite
Next

For Indice = 1 To iPlotGases.AnnotationCount
    iPlotGases.Annotation(Indice - 1).FontColor = vbWhite
Next

For Indice = 1 To iPlotGases.LabelCount
    iPlotGases.Labels(Indice - 1).FontColor = vbWhite
Next


iPlotGases.EndUpdate


End Sub

Private Sub CmdConsultar_Click()
    
    Dim fecha                           As Date
    Dim FechaInicio                     As Date
    Dim FechaFin                        As Date
    Dim Indice                          As Long
    Dim Index                           As Long
    Dim objPozoGasTiempos               As New clsPozoGasTiempos
    Dim objPozoGasTiempo                As clsPozoGasTiempo
    
    
    If Txtfecha.Value <> 0 Then
        
        If Txthora.Text <> "" Then
            
            ChkTiempoReal = 0
            
            iPlotGases.ClearAllData
            iPlotGases.RemoveAllAnnotations
            
            FechaInicio = CDate(Txtfecha.Value & " " & Txthora.Text)
            FechaFin = DateAdd("n", gSpan, FechaInicio)
            iPlotGases.XAxis(0).Min = Month(FechaInicio) * 30 + Day(FechaInicio) + Hour(FechaInicio) / 24 + Minute(FechaInicio) / (24 * 60) + Second(FechaInicio) / (24# * 60# * 60#)
            iPlotGases.XAxis(0).Span = ObtenerSpanEnDías(FechaInicio, FechaFin)
            
            Set objPozoGasTiempos.Datos = gDatos
            objPozoGasTiempos.dbPozoGasesTiempo gObjPozoActivo.IdPozo, FechaInicio, FechaFin
            
            For Each objPozoGasTiempo In objPozoGasTiempos
                
                fecha = Month(objPozoGasTiempo.fecha) * 30 + Day(objPozoGasTiempo.fecha) + Hour(objPozoGasTiempo.fecha) / 24 + Minute(objPozoGasTiempo.fecha) / (24 * 60) + Second(objPozoGasTiempo.fecha) / (24# * 60# * 60#)
                iPlotGases.Channel(0).AddXY fecha, objPozoGasTiempo.GasTotal
                iPlotGases.Channel(1).AddXY fecha, objPozoGasTiempo.CO2
                'iPlotGases.Channel(2).AddXY Fecha, objPozoGasTiempo.SH2
                
                If objPozoGasTiempo.Comentario <> "" Then
                    
                     Index = iPlotGases.AddAnnotation
                     iPlotGases.Annotation(Index).Font.Size = 12
                     iPlotGases.Annotation(Index).Reference = iprtChannel
                     iPlotGases.Annotation(Index).ChannelName = iPlotGases.Channel(0).Name
                     iPlotGases.Annotation(Index).Y = 13000 'Center X Coordinate
                     iPlotGases.Annotation(Index).X = fecha 'Center Y Coordinate
                     iPlotGases.Annotation(Index).Style = ipasText 'Text Annotation
                     iPlotGases.Annotation(Index).FontColor = vbWhite 'White Font
                     iPlotGases.Annotation(Index).Text = objPozoGasTiempo.Comentario
                     iPlotGases.Annotation(Index).TextRotation = ira000   'Rotate up-side-down
                     
                End If
            Next
            
        End If
        
    End If
    FrmAnalisis.iPlotGases.Visible = False
    FrmAnalisis.iPlotGases.Visible = True

End Sub



Private Sub cmdDownArrowProf_Click()
    ScrollLeftProf
End Sub

Private Sub CmdGuardar_Click()
    
    Dim objPozoEscaneo                  As clsPozoEscaneo
    
    On Error GoTo Error
        
    Set objPozoEscaneo = gObjPozoActivo.PozoEscaneos.ColPozoEscaneo(, , LvwCronoGas.SelectedItem.Key)
    If Not objPozoEscaneo Is Nothing Then
    
        objPozoEscaneo.Crono = TxtCrono.Text
        objPozoEscaneo.ROP = TxtROP.Text
        objPozoEscaneo.GasTotal = TxtGasTotal.Text
        objPozoEscaneo.CO2 = TxtCO2.Text
        
        gDatos.BeginTrans
        
        If gObjPozoActivo.PozoEscaneos.dbModificar(objPozoEscaneo) Then
            gDatos.CommitTrans
            LvwCronoGas.SelectedItem.SubItems(1) = Format(objPozoEscaneo.Crono, "0.00")
            LvwCronoGas.SelectedItem.SubItems(2) = Format(objPozoEscaneo.ROP, "0.00")
            LvwCronoGas.SelectedItem.SubItems(3) = Format(objPozoEscaneo.GasTotal, "0")
            LvwCronoGas.SelectedItem.SubItems(4) = Format(objPozoEscaneo.CO2, "0")
        Else
            gDatos.RollBackTrans
        End If
        
    End If
Error:
    
    If Err.Number <> 0 Then
        
        frmMsg.MostrarMsg "Módulo: CmdGuardar_Click." & Chr(10) & Err.Number & " - " & Err.Description, "Error", MdiPrincipal
        
        Err.Clear
        
    End If
    
End Sub

Private Sub CmdGuardarComentario_Click()
    
    If TxtComentario.Text <> "" Then
        
        ComentarioGas = TxtComentario.Text
        
        TxtComentario.Text = ""
        
    End If
    
End Sub

Private Sub cmdDownArrow_Click()
    ScrollLeft
End Sub

Private Sub cmdNuevoPozo_Click()
End Sub

Private Sub cmdPlay_Click()
    
    Txtfecha.Enabled = False
    Txthora.Enabled = False
    
    lblFecha.Visible = False
    Txtfecha.Visible = False
    lblHora.Visible = False
    Txthora.Visible = False
    fraCmdConsultar.Visible = False
    cmdConsultar.Enabled = False
    cmdConsultar.Visible = False
    ChkTiempoReal = 1
    cmdPlay.Enabled = False
    cmdStop.Enabled = True
    cmdUpArrow.Visible = False
    fraUpArrow.Visible = False
    cmdDownArrow.Visible = False
    fraDownArrow.Visible = False
    LimpiarYConsultarDatosTiempo
    
End Sub

Private Sub cmdRefresh_Click()
    
    ConsultarPozoEscaneos 0

End Sub

Private Sub cmdUpArrow_Click()
    ScrollRight
End Sub

Private Sub cmdStop_Click()
    Dim fecha As Date
    fecha = DateAdd("n", -1 * gSpan, Now)
    lblFecha.Visible = True
    Txtfecha.Enabled = True
    lblHora.Visible = True
    Txthora.Enabled = True
    Txtfecha.Visible = True
    Txthora.Visible = True
    fraCmdConsultar.Visible = True
    cmdConsultar.Visible = True
    fraUpArrow.Visible = True
    cmdUpArrow.Visible = True
    fraDownArrow.Visible = True
    cmdDownArrow.Visible = True
    
    cmdConsultar.Enabled = True
    Txtfecha = CDate(Format(fecha, "dd/mm/yyyy"))
    Txthora.Text = Format(fecha, "hh:nn:ss")
    
    ChkTiempoReal = 0
    cmdPlay.Enabled = True
    cmdStop.Enabled = False
    CmdConsultar_Click

End Sub

Sub CambiarBotonesProf(ByVal Estado As Boolean)
    
    cmdActualizar.Enabled = Estado
    cmdUpArrowProf.Enabled = Estado
    cmdDownArrowProf.Enabled = Estado
    
    If Not Estado Then
        MousePointer = vbHourglass
    Else
        'MousePointer = vbArrow
    End If
    
    
End Sub

Private Sub cmdActualizar_Click()
    Dim i As Long
    
    CambiarBotonesProf False
    iPlotMasterLog.BeginUpdate
    If IsNumeric(TxtSpan) Then
        SaveSetting "Iga", "Config", "SpanProfundidad", TxtSpan.Text
    End If
    iPlotMasterLog.EndUpdate
    CambiarBotonesProf True
    
    pObjAdminEventos.ActualizarVistaProf
    
End Sub

Private Sub ChkMostrarTodosLosAnálisisTemporal_Click()
    
    Dim StrSql                  As String
    Dim Analisis                As Recordset
    Dim Item                    As ListItem
    Dim objPozoAnalisis         As New clsPozoAnalisis
    Dim objPozoAnalis           As clsPozoAnalis
    
    On Error GoTo Error
    If Not gObjPozoActivo Is Nothing Then
    
        LvwAnalisisTemporal.ListItems.Clear
        LimpiatDatosTemporales
        Set objPozoAnalisis.Datos = gDatos
        If ChkMostrarTodosLosAnálisisTemporal.Value <> 1 Then
            objPozoAnalisis.ConsultarPozoAnalisis gObjPozoActivo.IdPozo, , , False, 1
        Else
            objPozoAnalisis.ConsultarPozoAnalisis gObjPozoActivo.IdPozo, , , False
        End If
        
        For Each objPozoAnalis In objPozoAnalisis
            
            Set Item = LvwAnalisisTemporal.ListItems.Add(, objPozoAnalis.NumeroAnalisis & "ID")
            
            Item.Tag = objPozoAnalis.CodigoTipoDeAnalisis
            
            Item.Text = Format(objPozoAnalis.NumeroAnalisis, "00000")
            Item.SubItems(1) = objPozoAnalis.NumeroAnalisis
            Item.SubItems(2) = Format(objPozoAnalis.fecha, "DD/MM/YYYY HH:MM:SS")
            Item.SubItems(3) = objPozoAnalis.Profundidad
            
            Select Case objPozoAnalis.CodigoTipoDeAnalisis
                
                Case Is = 1
                    Item.SmallIcon = 1
                Case Is = 2
                    Item.SmallIcon = 2
                Case Is = 3
                    Item.SmallIcon = 3
                Case Is = 4
                    Item.SmallIcon = 4
                Case Is = 5
                    Item.SmallIcon = 5
            End Select
                
        Next
        If LvwAnalisisDefinitivo.ListItems.Count <> 0 Then
            
            LvwAnalisisDefinitivo.SelectedItem = LvwAnalisisDefinitivo.ListItems(1)
            LvwAnalisisDefinitivo.SelectedItem.Selected = True
            
        End If
        
        If LvwAnalisisTemporal.ListItems.Count <> 0 Then
            LvwAnalisisTemporal.SelectedItem = LvwAnalisisTemporal.ListItems(1)
            Set objPozoAnalis = gObjPozoActivo.PozoAnalisis.ColPozoAnalis(gObjPozoActivo.IdPozo, Val(LvwAnalisisTemporal.SelectedItem.Key))
            If Not objPozoAnalis Is Nothing Then
                BuscarComponentes objPozoAnalis, "TEMPORAL"
                BuscarDatosAnalisisTemporales objPozoAnalis
                BuscarRelacionesCromatograficasTemporal objPozoAnalis
            End If
        Else
            LimpiatDatosTemporales
        End If
    End If
Error:
    
    If Err.Number <> 0 Then
        
        frmMsg.MostrarMsg "Módulo: ChkMostrarTodosLosAnálisisTemporal." & Chr(10) & Err.Number & " - " & Err.Description, "Error", MdiPrincipal
        
        Err.Clear
        
    End If
    
End Sub

Private Sub CmdGuardarTemporal_Click()
    
    Dim NumeroAnalisis                  As Long
    Dim objPozoAnalis                   As clsPozoAnalis
    Dim StrSql                          As String
        
    If Not gObjPozoActivo Is Nothing Then
    
        NumeroAnalisis = Val(LvwAnalisisTemporal.SelectedItem.Key)
        Set objPozoAnalis = gObjPozoActivo.PozoAnalisis.ColPozoAnalis(gObjPozoActivo.IdPozo, NumeroAnalisis)
        If Not objPozoAnalis Is Nothing Then
            If TxtCO2Temporal.Text <> "" Then
                objPozoAnalis.CO2 = TxtCO2Temporal.Text
            Else
                objPozoAnalis.CO2 = 0
            End If
            If TxtSH2Temporal.Text <> "" Then
                objPozoAnalis.SH2 = TxtSH2Temporal.Text
            Else
                objPozoAnalis.SH2 = 0
            End If
            If TxtProfundidadTemporal.Text <> "" Then
                objPozoAnalis.Profundidad = TxtProfundidadTemporal.Text
            Else
                objPozoAnalis.Profundidad = 0
            End If
            objPozoAnalis.Observaciones = " "
            If TxtGasTotalTemporal.Text <> "" Then
                objPozoAnalis.GasTotal = TxtGasTotalTemporal.Text
            Else
                objPozoAnalis.GasTotal = 0
            End If
            
            gDatos.BeginTrans
            If gObjPozoActivo.PozoAnalisis.dbModificar(objPozoAnalis) Then
                gDatos.CommitTrans
                LvwAnalisisTemporal.SelectedItem.ListSubItems(3) = objPozoAnalis.Profundidad
            Else
                gDatos.RollBackTrans
            End If
                
        End If
    End If
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Public Sub LimpiarYConsultarDatosTiempo()
    
    Dim StrSql                          As String
    Dim FechaInicio                     As Date
    Dim FechaFin                        As Date
    Dim Datos                           As Recordset
    Dim fecha                           As Double
    Dim Index                           As Long
    Dim objPozoGasTiempo                   As clsPozoGasTiempo
    Dim objPozoGasTiempos                  As New clsPozoGasTiempos
        
    If Not gObjPozoActivo Is Nothing Then
        Set objPozoGasTiempos.Datos = gDatos
        gRefrescarGrafico = False
        If ChkTiempoReal = 1 Then
            
            EstoyCargando = True
            
            iPlotGases.ClearAllData
            iPlotGases.RemoveAllAnnotations
            
            FechaFin = Now()
            FechaInicio = DateAdd("n", -1 * gSpan, FechaFin)
            iPlotGases.XAxis(0).Span = ObtenerSpanEnDías(FechaInicio, DateAdd("n", gZonaMuerta, FechaFin))
            iPlotGases.XAxis(0).Min = Month(FechaInicio) * 30 + Day(FechaInicio) + Hour(FechaInicio) / 24 + Minute(FechaInicio) / (24 * 60) + Second(FechaInicio) / (24# * 60# * 60#)
            'here
            objPozoGasTiempos.dbPozoGasesTiempo gObjPozoActivo.IdPozo, FechaInicio, FechaFin
            
            For Each objPozoGasTiempo In objPozoGasTiempos
                
                fecha = Month(objPozoGasTiempo.fecha) * 30 + Day(objPozoGasTiempo.fecha) + Hour(objPozoGasTiempo.fecha) / 24 + Minute(objPozoGasTiempo.fecha) / (24 * 60) + Second(objPozoGasTiempo.fecha) / (24# * 60# * 60#)
                iPlotGases.Channel(0).AddXY fecha, objPozoGasTiempo.GasTotal
                iPlotGases.Channel(1).AddXY fecha, objPozoGasTiempo.CO2
                If objPozoGasTiempo.Comentario <> "" Then
                     Index = iPlotGases.AddAnnotation
                     iPlotGases.Annotation(Index).Font.Size = 12
                     iPlotGases.Annotation(Index).Reference = iprtChannel
                     iPlotGases.Annotation(Index).ChannelName = iPlotGases.Channel(0).Name
                     iPlotGases.Annotation(Index).Y = 13000 'Center X Coordinate
                     iPlotGases.Annotation(Index).X = fecha 'Center Y Coordinate
                     iPlotGases.Annotation(Index).Style = ipasText 'Text Annotation
                     iPlotGases.Annotation(Index).FontColor = vbWhite 'White Font
                     iPlotGases.Annotation(Index).Text = objPozoGasTiempo.Comentario
                     iPlotGases.Annotation(Index).TextRotation = ira000   'Rotate up-side-down
                End If
            Next
            
            EstoyCargando = False
            
        End If
        gRefrescarGrafico = False
    End If
End Sub

Private Sub Command4_Click()

    OrdenarPantallaImpresion iPlotGases

End Sub

Private Sub cmdUpArrowProf_Click()
    ScrollRightProf
End Sub
Public Sub configVistaEscalas()
        
    If gObjConfiguracion.IdTipoEquipoGas = 1 And gObjConfiguracion.IdTipoEquipoCroma = 1 Then
    
                SwTha.Visible = True
                iLedTha.Visible = SwTha.Visible = True
                Label61.Visible = SwTha.Visible = True

                SwNormal.Visible = True
                iLedFast.Visible = True
                Label62.Visible = True
    
                SwFast.Visible = True
                iLedNormal.Visible = True
                Label63.Visible = True
                
        'IMPROBABLE
        
        ElseIf gObjConfiguracion.IdTipoEquipoGas = 1 And gObjConfiguracion.IdTipoEquipoCroma = 2 Then
                SwTha.Visible = True
                iLedTha.Visible = True
                Label61.Visible = True

                SwNormal.Visible = True
                iLedFast.Visible = True
                Label62.Visible = True
    
                SwFast.Visible = True
                iLedNormal.Visible = True
                Label63.Visible = True
                
        
        ElseIf gObjConfiguracion.IdTipoEquipoGas = 1 And gObjConfiguracion.IdTipoEquipoCroma = 3 Then
                SwTha.Visible = True
                iLedTha.Visible = True
                Label61.Visible = True

                SwNormal.Visible = False
                iLedFast.Visible = False
                Label62.Visible = False
    
                SwFast.Visible = False
                iLedNormal.Visible = False
                Label63.Visible = False
                
                
                
        ElseIf gObjConfiguracion.IdTipoEquipoGas = 1 And gObjConfiguracion.IdTipoEquipoCroma = 4 Then
        
                SwTha.Visible = True
                iLedTha.Visible = True
                Label61.Visible = True

                SwNormal.Visible = False
                iLedFast.Visible = False
                Label62.Visible = False
    
                SwFast.Visible = False
                iLedNormal.Visible = False
                Label63.Visible = False
                
                
        ElseIf gObjConfiguracion.IdTipoEquipoGas = 1 And gObjConfiguracion.IdTipoEquipoCroma = 5 Then
                
                SwTha.Visible = True
                iLedTha.Visible = True
                Label61.Visible = True

                SwNormal.Visible = False
                iLedFast.Visible = False
                Label62.Visible = False
    
                SwFast.Visible = False
                iLedNormal.Visible = False
                Label63.Visible = False
                
    End If
                
                
                
                
    If gObjConfiguracion.IdTipoEquipoGas = 2 And gObjConfiguracion.IdTipoEquipoCroma = 1 Then
    
                SwTha.Visible = False
                iLedTha.Visible = False
                Label61.Visible = False

                SwNormal.Visible = True
                iLedFast.Visible = True
                Label62.Visible = True
    
                SwFast.Visible = True
                iLedNormal.Visible = True
                Label63.Visible = True
                 
         ElseIf gObjConfiguracion.IdTipoEquipoGas = 2 And gObjConfiguracion.IdTipoEquipoCroma = 2 Then
                 SwTha.Visible = False
                iLedTha.Visible = False
                Label61.Visible = False

                SwNormal.Visible = True
                iLedFast.Visible = True
                Label62.Visible = True
    
                SwFast.Visible = True
                iLedNormal.Visible = True
                Label63.Visible = True
         
         ElseIf gObjConfiguracion.IdTipoEquipoGas = 2 And gObjConfiguracion.IdTipoEquipoCroma = 3 Then
                 SwTha.Visible = False
                iLedTha.Visible = False
                Label61.Visible = False

                SwNormal.Visible = False
                iLedFast.Visible = False
                Label62.Visible = False
    
                SwFast.Visible = False
                iLedNormal.Visible = False
                Label63.Visible = False
                 
         ElseIf gObjConfiguracion.IdTipoEquipoGas = 2 And gObjConfiguracion.IdTipoEquipoCroma = 4 Then
                 SwTha.Visible = False
                iLedTha.Visible = False
                Label61.Visible = False

                SwNormal.Visible = False
                iLedFast.Visible = False
                Label62.Visible = False
    
                SwFast.Visible = False
                iLedNormal.Visible = False
                Label63.Visible = False
                
         ElseIf gObjConfiguracion.IdTipoEquipoGas = 2 And gObjConfiguracion.IdTipoEquipoCroma = 5 Then
                SwTha.Visible = False
                iLedTha.Visible = False
                Label61.Visible = False

                SwNormal.Visible = False
                iLedFast.Visible = False
                Label62.Visible = False
    
                SwFast.Visible = False
                iLedNormal.Visible = False
                Label63.Visible = False
    End If
    
    
    If gObjConfiguracion.IdTipoEquipoGas = 3 And gObjConfiguracion.IdTipoEquipoCroma = 1 Then
                SwTha.Visible = True
                iLedTha.Visible = True
                Label61.Visible = True

                SwNormal.Visible = True
                iLedFast.Visible = True
                Label62.Visible = True
    
                SwFast.Visible = True
                iLedNormal.Visible = True
                Label63.Visible = True
          
         ElseIf gObjConfiguracion.IdTipoEquipoGas = 3 And gObjConfiguracion.IdTipoEquipoCroma = 2 Then
                SwTha.Visible = True
                iLedTha.Visible = True
                Label61.Visible = True

                SwNormal.Visible = True
                iLedFast.Visible = True
                Label62.Visible = True
    
                SwFast.Visible = True
                iLedNormal.Visible = True
                Label63.Visible = True
                
         ElseIf gObjConfiguracion.IdTipoEquipoGas = 3 And gObjConfiguracion.IdTipoEquipoCroma = 3 Then
                SwTha.Visible = True
                iLedTha.Visible = True
                Label61.Visible = True

                SwNormal.Visible = False
                iLedFast.Visible = False
                Label62.Visible = False
    
                SwFast.Visible = False
                iLedNormal.Visible = False
                Label63.Visible = False
             
         ElseIf gObjConfiguracion.IdTipoEquipoGas = 3 And gObjConfiguracion.IdTipoEquipoCroma = 4 Then
                SwTha.Visible = True
                iLedTha.Visible = True
                Label61.Visible = True

                SwNormal.Visible = False
                iLedFast.Visible = False
                Label62.Visible = False
    
                SwFast.Visible = False
                iLedNormal.Visible = False
                Label63.Visible = False
                
         ElseIf gObjConfiguracion.IdTipoEquipoGas = 3 And gObjConfiguracion.IdTipoEquipoCroma = 5 Then
                SwTha.Visible = True
                iLedTha.Visible = True
                Label61.Visible = True

                SwNormal.Visible = False
                iLedFast.Visible = False
                Label62.Visible = False
    
                SwFast.Visible = False
                iLedNormal.Visible = False
                Label63.Visible = False
    End If
    
    

End Sub


Private Sub Form_Load()
 
    Dim Indice As Long

    On Error GoTo Error
    
    ' 1. Inicializar el redimensionamiento de componentes (esto SIEMPRE va primero)
    Call InitializeFormResize(Me)

        
    Set pObjAdminEventos = gObjAdminEventos
    SetearList
    
    ActualizoDisplayEscalas
    
    SetearRangoGraficacion
    BuscarAnalisis
    ConsultarPozoEscaneos 0
    
    configVistaEscalas
    
    iPlotMasterLog.YAxis(0).Span = gSpanYCrono
    iPlotGases.RemoveAllLabels
    iPlotGases.XAxis(0).TrackingEnabled = False
    CmdConsultar_Click
    cmdPlay_Click
    
    pObjAdminEventos.ActualizarVistaProf
    pObjAdminEventos.ActualizarVistaTiempo 3
    
    
    MdiPrincipal.TimerRefresh.Enabled = True
    
Error:
If Err.Number <> 0 Then
    frmMsg.MostrarMsg "No se encuentra el archivo de propiedades dle gráfico", "Error de archivo"
    Err.Clear
    End If
    
End Sub

Private Sub SetearList()
    
    With LvwAnalisisDefinitivo.ColumnHeaders
        .Clear
        
        .Add , "Nro", "Nro", 0
        .Add , "Nº", "Nº", 700
        .Add , "Fecha", "Fecha", 1730
        .Add , "Prof.", "Prof.", 550
        
    End With
    
    LvwAnalisisDefinitivo.Sorted = True
    
    With LvwAnalisisTemporal.ColumnHeaders
        .Clear
    
        .Add , "Nro", "Nro", 0
        .Add , "Nº", "Nº", 700
        .Add , "Fecha", "Fecha", 1730
        .Add , "Prof.", "Prof.", 550
        
    End With
    
    LvwAnalisisTemporal.Sorted = True
    
    With LvwDefinitivo.ColumnHeaders
        .Clear
        
        .Add , "Numero", "Nª", 300
        .Add , "Component", "Component", 1400
        .Add , "Retention", "R T  ", 0, lvwColumnRight '600
        .Add , "Area", "Area", 0, lvwColumnRight '900
        .Add , "Externo", "Externo", 850, lvwColumnRight
        .Add , "Units", "Units", 650, lvwColumnRight
        .Add , "NormArea", "  %  ", 600, lvwColumnRight
        
    End With
    
    LvwDefinitivo.Sorted = True
    
    With LvwTemporal.ColumnHeaders
        .Clear
        
        .Add , "Numero", "Nª", 300
        .Add , "Component", "Component", 1400
        .Add , "Retention", "R T  ", 0, lvwColumnRight ' 600
        .Add , "Area", "Area", 0, lvwColumnRight '900
        .Add , "Externo", "Externo", 850, lvwColumnRight
        .Add , "Units", "Units", 650, lvwColumnRight
        .Add , "NormArea", "  %  ", 600, lvwColumnRight
        
    End With
    
    LvwTemporal.Sorted = True
    LvwTemporal.SortKey = 0
    
    With LvwCronoGas.ColumnHeaders
        .Clear
        
        .Add , "Prof.", "Prof.", 1700
        .Add , "Crono", "Crono", 1400, 1
        .Add , "ROP", "ROP", 1400, 1
        .Add , "Gas Total", "Gas Total", 1400, 1
        .Add , "CO2", "CO2", 1320, 1
        .Add , "Orden", "Orden", 0
        .Add , "EstadoGas", "EstadoGas", 0
        
    End With
    
    LvwCronoGas.Sorted = True
    LvwCronoGas.SortOrder = lvwDescending
    LvwCronoGas.SortKey = 6
    
End Sub

Private Function BuscarComponentes(ByVal objPozoAnalis As clsPozoAnalis, List As String) As Boolean
    
    Dim Item                                    As ListItem
    Dim objPozoAnalisComponente                 As ClsPozoAnalisComponente
    Dim objComponente                           As clsComponente
    
    On Error GoTo Error
    
    BuscarComponentes = True
    
    
    If List = "TEMPORAL" Then
        LvwTemporal.ListItems.Clear
    Else
        LvwDefinitivo.ListItems.Clear
    End If

    For Each objPozoAnalisComponente In objPozoAnalis.PozoAnalisComponentes
    
        If List = "TEMPORAL" Then
            Set Item = LvwTemporal.ListItems.Add(, objPozoAnalisComponente.NumeroComponente & "ID")
        Else
            Set Item = LvwDefinitivo.ListItems.Add(, objPozoAnalisComponente.NumeroComponente & "ID")
        End If
            
        Item.Text = objPozoAnalisComponente.NumeroComponente
        Item.SubItems(1) = objPozoAnalisComponente.Component
        Item.SubItems(2) = Format(objPozoAnalisComponente.Retention, "0.00")
        Item.SubItems(3) = Format(objPozoAnalisComponente.Area, "0.0")
        Item.SubItems(4) = Format(objPozoAnalisComponente.Externo, "0")
        Item.SubItems(5) = objPozoAnalisComponente.Units
        Item.SubItems(6) = Format(objPozoAnalisComponente.NormArea, "0.0")
        
    Next
    
    If LvwTemporal.ListItems.Count <> 0 Then
        
        LvwTemporal.SelectedItem = LvwTemporal.ListItems(1)
        LvwTemporal.SelectedItem.Selected = True
        
    End If
    
    If LvwDefinitivo.ListItems.Count <> 0 Then
        
        LvwDefinitivo.SelectedItem = LvwDefinitivo.ListItems(1)
        LvwDefinitivo.SelectedItem.Selected = True
        
    End If
    
Error:
    
    If Err.Number <> 0 Then
        BuscarComponentes = False
        frmMsg.MostrarMsg "Módulo: BuscarComponentes." & Chr(10) & Err.Number & " - " & Err.Description, "Error", MdiPrincipal
        Err.Clear
    End If
    
End Function

Private Sub ListView1_BeforeLabelEdit(Cancel As Integer)

End Sub



Private Sub Form_Resize()
    Call modAutoResize.ResizeForm(Me)
End Sub

Private Sub iPlotGases_OnAfterCloseEditor()

On Error GoTo Error

iPlotGases.SavePropertiesToFile "c:\Registros\Setup\PropGas.txt"

Error:

    Err.Clear
    
End Sub

Private Sub iPlotGases_OnMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    If Button = vbRightButton And X > 120 Then
        PopupMenu MdiPrincipal.mnuGrafico
    End If
    
End Sub


Private Sub iPlotMasterLog_OnMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)
Dim FinEjeYCrono As Long
    
    FinEjeYCrono = iPlotMasterLog.OuterMarginLeft + _
                   iPlotMasterLog.XAxis(0).Width + _
                   iPlotMasterLog.XAxis(0).InnerMargin + _
                   iPlotMasterLog.XAxis(0).OuterMargin + _
                   iPlotMasterLog.XAxis(0).LabelsMargin + _
                   iPlotMasterLog.XAxis(0).MajorLength + _
                   iPlotMasterLog.XAxis(0).TitleMargin + _
                   iPlotMasterLog.YAxis(0).Height + 10
    
    
    If Button = vbRightButton Then
        If X > FinEjeYCrono Then
            PopupMenu MdiPrincipal.mnuEscalasLog
        Else
            frmSpanCrono.Show
            'iPlotMasterLog.ShowPropertyEditor
        End If
    End If
End Sub


Private Sub LvwAnalisisDefinitivo_Click()
    
    If BotonDerecho Then
        
        MdiPrincipal.mnuPopupDefinitivoModificar.Visible = False
        
        MdiPrincipal.mnuPopupDefinitivoBorrar.Enabled = LvwAnalisisDefinitivo.ListItems.Count <> 0
        
        Me.PopupMenu MdiPrincipal.mnuPopupDefinitivo
        
        If txtPopupElegido.Text = "BORRAR" Then
            
            AnalisisDefinitivoBorrar
            
        End If
        
        txtPopupElegido.Text = ""
        
    End If
    
End Sub

Private Sub LvwAnalisisDefinitivo_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    Dim NumeroAnalisis              As Long
    Dim objPozoAnalis               As clsPozoAnalis
    
    If LvwAnalisisDefinitivo.ListItems.Count <> 0 Then
        
        NumeroAnalisis = Val(LvwAnalisisDefinitivo.SelectedItem.Key)
        Set objPozoAnalis = gObjPozoActivo.PozoAnalisis.ColPozoAnalis(gObjPozoActivo.IdPozo, NumeroAnalisis)
        If Not objPozoAnalis Is Nothing Then
            If objPozoAnalis.Seleccionado Then
                ArchivoDefinitivo.caption = objPozoAnalis.archivo
            Else
                ArchivoTemporal.caption = objPozoAnalis.archivo
            End If
            BuscarComponentes objPozoAnalis, "DEFINITIVO"
            BuscarDatosAnalisisDefinitivos objPozoAnalis
            BuscarRelacionesCromatograficasDefinitivo objPozoAnalis
        End If
    End If
    
End Sub

Private Sub LvwAnalisisDefinitivo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    BotonDerecho = Button = vbRightButton
    
End Sub

Private Sub LvwAnalisisTemporal_Click()
    
    If BotonDerecho Then
        
        MdiPrincipal.mnuPopupTemporalAgregar.Visible = False
        MdiPrincipal.mnuPopupTemporalModificar.Visible = False
        MdiPrincipal.mnuPopupTemporalSeleccionar.Visible = True
        MdiPrincipal.mnuPopupTemporalSeparador1.Visible = True
        MdiPrincipal.mnuPopupTemporalCalibracion.Visible = True
        MdiPrincipal.mnuPopupTemporalGasDeYacimiento.Visible = True
        MdiPrincipal.mnuPopupTemporalCirculada.Visible = True
        
        MdiPrincipal.mnuPopupTemporalBorrar.Enabled = LvwAnalisisTemporal.ListItems.Count <> 0
        MdiPrincipal.mnuPopupTemporalSeleccionar.Enabled = LvwAnalisisTemporal.ListItems.Count <> 0
        MdiPrincipal.mnuPopupTemporalCalibracion.Enabled = LvwAnalisisTemporal.ListItems.Count <> 0
        MdiPrincipal.mnuPopupTemporalGasDeYacimiento.Enabled = LvwAnalisisTemporal.ListItems.Count <> 0
        MdiPrincipal.mnuPopupTemporalCirculada.Enabled = LvwAnalisisTemporal.ListItems.Count <> 0
        
        If LvwAnalisisTemporal.ListItems.Count <> 0 Then
            
            If LvwAnalisisTemporal.SelectedItem.SmallIcon = 4 Then
                
                MdiPrincipal.mnuPopupTemporalBorrar.caption = "Restaurar"
                
            Else
                
                MdiPrincipal.mnuPopupTemporalBorrar.caption = "Borrar"
                
            End If
            
            If LvwAnalisisTemporal.SelectedItem.SmallIcon = 2 Then
                
                MdiPrincipal.mnuPopupTemporalCalibracion.caption = "Analisis normal"
                
            Else
                
                MdiPrincipal.mnuPopupTemporalCalibracion.caption = "Calibracion"
                
            End If
            
            If LvwAnalisisTemporal.SelectedItem.SmallIcon = 3 Then
                
                MdiPrincipal.mnuPopupTemporalGasDeYacimiento.caption = "Analisis normal"
                
            Else
                
                MdiPrincipal.mnuPopupTemporalGasDeYacimiento.caption = "Gas de yacimiento"
                
            End If
            
            If LvwAnalisisTemporal.SelectedItem.SmallIcon = 5 Then
                
                MdiPrincipal.mnuPopupTemporalCirculada.caption = "Analisis normal"
                
            Else
                
                MdiPrincipal.mnuPopupTemporalCirculada.caption = "Circulada"
                
            End If
            
        End If
        
        Me.PopupMenu MdiPrincipal.mnuPopupTemporal
        
        If txtPopupElegido.Text = "SELECCIONAR" Then
            
            AnalisisTemporalSeleccionar
            
        ElseIf txtPopupElegido.Text = "BORRAR" Then
            
            AnalisisTemporalBorrar
            
        ElseIf txtPopupElegido.Text = "CALIBRACION" Then
            
            AnalisisCalibracion
            
        ElseIf txtPopupElegido.Text = "GAS DE YACIMIENTO" Then
            
            AnalisisGasDeYacimiento
            
        ElseIf txtPopupElegido.Text = "CIRCULADA" Then
            
            AnalisisCirculada
            
        End If
        
        txtPopupElegido.Text = ""
        
    End If
    
End Sub

Private Sub LvwAnalisisTemporal_ItemClick(ByVal Item As MSComctlLib.ListItem)

    Dim NumeroAnalisis                  As Long
    Dim objPozoAnalis                   As clsPozoAnalis
    
    If LvwAnalisisTemporal.ListItems.Count <> 0 Then
        
        NumeroAnalisis = Val(LvwAnalisisTemporal.SelectedItem.Key)
        Set objPozoAnalis = gObjPozoActivo.PozoAnalisis.ColPozoAnalis(gObjPozoActivo.IdPozo, NumeroAnalisis)
        If Not objPozoAnalis Is Nothing Then
            If objPozoAnalis.Seleccionado Then
                ArchivoTemporal.caption = objPozoAnalis.archivo
            Else
                ArchivoTemporal.caption = objPozoAnalis.archivo
            End If
            BuscarComponentes objPozoAnalis, "TEMPORAL"
            BuscarDatosAnalisisTemporales objPozoAnalis
            BuscarRelacionesCromatograficasTemporal objPozoAnalis
        End If
        
    End If
    
End Sub

Private Sub LvwAnalisisTemporal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    BotonDerecho = Button = vbRightButton
    
End Sub

Private Sub LvwCronoGas_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    If LvwCronoGas.ListItems.Count <> 0 Then
        
        TxtCrono.Text = Item.SubItems(1)
        TxtROP.Text = Item.SubItems(2)
        TxtGasTotal.Text = Item.SubItems(3)
        TxtCO2.Text = Item.SubItems(4)
        
    End If
    
End Sub

Private Sub LvwTemporal_AfterLabelEdit(Cancel As Integer, NewString As String)
    
    Dim objPozoAnalis                   As clsPozoAnalis
    Dim objPozoAnalisComponente         As ClsPozoAnalisComponente
    Dim objPozoAnalisComponenteNew      As ClsPozoAnalisComponente
    Dim objComponente                   As clsComponente
    
    Dim Numero As Integer
    Dim NumeroAnalisis As Long
    Dim NumeroComponenteAnterior As Long
    
    Dim ok As Boolean
    
    On Error GoTo Error
    
    NumeroAnalisis = LvwAnalisisTemporal.SelectedItem.Text
    NumeroComponenteAnterior = TxtNumeroComponenteAnterior.Text
    Set objPozoAnalis = gObjPozoActivo.PozoAnalisis.ColPozoAnalis(gObjPozoActivo.IdPozo, NumeroAnalisis)
    If Not objPozoAnalis Is Nothing Then
        If NewString <> "" Then
            
            If IsNumeric(NewString) Then
                Numero = NewString
                If Err.Number = 6 Then
                    'Numero mayor a un entero.
                    frmMsg.MostrarMsg "El número debe ser mayor que 0 (Cero) y menor que 32.767.", "Error", MdiPrincipal
                    Cancel = True
                    Err.Clear
                ElseIf Err.Number <> 0 Then
                    frmMsg.MostrarMsg Err.Number & " " & Err.Description, "Error", MdiPrincipal
                    Cancel = True
                    Err.Clear
                ElseIf Numero < 0 Then
                    frmMsg.MostrarMsg "El número debe ser mayor o igual que cero.", "Error", MdiPrincipal
                    Cancel = True
                Else
                    Set objPozoAnalisComponente = objPozoAnalis.PozoAnalisComponentes.colComponente(objPozoAnalis.IdPozo, objPozoAnalis.NumeroAnalisis, Numero)
                    If objPozoAnalisComponente Is Nothing Then
                        Set objPozoAnalisComponente = objPozoAnalis.PozoAnalisComponentes.colComponente(objPozoAnalis.IdPozo, objPozoAnalis.NumeroAnalisis, NumeroComponenteAnterior)
                        
                        Set objPozoAnalisComponenteNew = objPozoAnalisComponente.Duplicado
                        objPozoAnalisComponenteNew.NumeroComponente = Numero
                        
                        Set objComponente = gObjComponentes.colComponente(Numero)
                        NewString = Numero
                        LvwTemporal.SelectedItem.SubItems(1) = objComponente.NombreComponente
                        
                        gDatos.BeginTrans
                        If objPozoAnalis.PozoAnalisComponentes.dbModificarComponente(objPozoAnalisComponente, NumeroComponenteAnterior) Then
                            objPozoAnalis.PozoAnalisComponentes.colQuitar objPozoAnalisComponente
                            objPozoAnalis.PozoAnalisComponentes.colAgregar objPozoAnalisComponenteNew
                            CalcularRelacionesCromatograficas objPozoAnalis, "MODIFICAR", ok
                            ActualizarGasTotalCromatografico objPozoAnalis, ok
                            BuscarComponentes objPozoAnalis, "TEMPORAL"
                            BuscarRelacionesCromatograficasTemporal objPozoAnalis
                            BuscarDatosAnalisisTemporales objPozoAnalis
                            If gObjPozoActivo.PozoAnalisis.dbModificar(objPozoAnalis) Then
                                gDatos.CommitTrans
                            Else
                                gDatos.RollBackTrans
                                Cancel = True
                                frmMsg.MostrarMsg "No se pudo actualizar el gas total cromatográfico", "Error", MdiPrincipal
                            End If
                        Else
                            Cancel = True
                            gDatos.RollBackTrans
                            frmMsg.MostrarMsg "No se pudo modificar el componente", "Error", MdiPrincipal
                        End If
                    Else
                        frmMsg.MostrarMsg "Ya existe el componente " & Numero & " en el análisis", "Error", MdiPrincipal
                        Cancel = True
                    End If
                End If
            Else
                Cancel = True
                frmMsg.MostrarMsg "Debe ingresar un valor numérico.", "Error", MdiPrincipal
            End If
        Else
            Cancel = True
        End If
        LvwTemporal.SetFocus
    End If
Error:
    If Err.Number <> 0 Then
        Cancel = True
        frmMsg.MostrarMsg "Módulo: lvwArticulos_AfterLabelEdit." & Chr(10) & Err.Number & " - " & Err.Description, "Error", MdiPrincipal
        Err.Clear
    End If
End Sub

Private Sub LvwTemporal_BeforeLabelEdit(Cancel As Integer)
    
    TxtNumeroComponenteAnterior.Text = LvwTemporal.SelectedItem.Text
    
End Sub

Private Sub LvwTemporal_Click()
    
    If BotonDerecho Then
        
        MdiPrincipal.mnuPopupTemporalAgregar.Visible = True
        MdiPrincipal.mnuPopupTemporalModificar.Visible = True
        MdiPrincipal.mnuPopupTemporalSeleccionar.Visible = False
        MdiPrincipal.mnuPopupTemporalSeparador1.Visible = False
        MdiPrincipal.mnuPopupTemporalCalibracion.Visible = False
        MdiPrincipal.mnuPopupTemporalGasDeYacimiento.Visible = False
        MdiPrincipal.mnuPopupTemporalCirculada.Visible = False
        
        MdiPrincipal.mnuPopupTemporalAgregar.Enabled = LvwTemporal.ListItems.Count <> 0
        MdiPrincipal.mnuPopupTemporalBorrar.Enabled = LvwTemporal.ListItems.Count <> 0
        MdiPrincipal.mnuPopupTemporalModificar.Enabled = LvwTemporal.ListItems.Count <> 0
        
        Me.PopupMenu MdiPrincipal.mnuPopupTemporal
        
        If txtPopupElegido.Text = "AGREGAR" Then
            
            ComponenteTemporalAgregar
            
        ElseIf txtPopupElegido.Text = "MODIFICAR" Then
            
            ComponenteTemporalModificar
            
        ElseIf txtPopupElegido.Text = "BORRAR" Then
            
            ComponenteTemporalBorrar
            
        End If
        
        txtPopupElegido.Text = ""
        
    End If
    
End Sub

Private Sub LvwTemporal_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        
        LvwTemporal.StartLabelEdit
        
    End If
    
End Sub

Private Sub LvwTemporal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    BotonDerecho = Button = vbRightButton
    
End Sub

Public Sub AnalisisTemporalBorrar()
    
    Dim NumeroAnalisis                  As Long
    Dim objPozoAnalis                   As clsPozoAnalis
    Dim Respuesta                       As Integer
    Dim Accion                          As String
    
    NumeroAnalisis = Val(LvwAnalisisTemporal.SelectedItem.Key)
    
    If LvwAnalisisTemporal.SelectedItem.SmallIcon = 1 Then
        Respuesta = MsgBox("¿Está seguro que desea borrar el analisis seleccionado?", vbYesNo + vbQuestion)
        Accion = "MARCAR"
    ElseIf LvwAnalisisTemporal.SelectedItem.SmallIcon = 2 Then
        Respuesta = MsgBox("El análisis seleccionado es un análisis de calibración y se eliminará de forma permanente. ¿Está seguro que desea borrar el analisis seleccionado?", vbYesNo + vbQuestion)
        Accion = "ELIMINAR"
    ElseIf LvwAnalisisTemporal.SelectedItem.SmallIcon = 3 Then
        Respuesta = MsgBox("El análisis seleccionado es un análisis de gas de yacimiento y se eliminará de forma permanente. ¿Está seguro que desea borrar el analisis seleccionado?", vbYesNo + vbQuestion)
        Accion = "ELIMINAR"
    ElseIf LvwAnalisisTemporal.SelectedItem.SmallIcon = 4 Then
        Respuesta = MsgBox("¿Está seguro que desea restaurar el analisis seleccionado?", vbYesNo + vbQuestion)
        Accion = "DESMARCAR"
    ElseIf LvwAnalisisTemporal.SelectedItem.SmallIcon = 3 Then
        Respuesta = MsgBox("El análisis seleccionado es un análisis de circulada y se eliminará de forma permanente. ¿Está seguro que desea borrar el analisis seleccionado?", vbYesNo + vbQuestion)
        Accion = "ELIMINAR"
    End If
    
    If Respuesta = vbYes Then
        NumeroAnalisis = Val(LvwAnalisisTemporal.SelectedItem.Key)
        Set objPozoAnalis = gObjPozoActivo.PozoAnalisis.ColPozoAnalis(gObjPozoActivo.IdPozo, NumeroAnalisis)
        If Not objPozoAnalis Is Nothing Then
            Select Case Accion
                
                Case Is = "MARCAR"
                    
                    LvwAnalisisTemporal.SelectedItem.SmallIcon = 4
                    
                    objPozoAnalis.CodigoTipoDeAnalisis = 4
                    gDatos.BeginTrans
                    If gObjPozoActivo.PozoAnalisis.dbModificar(objPozoAnalis) Then
                        gDatos.CommitTrans
                        If ChkMostrarTodosLosAnálisisTemporal.Value = 0 Then
                            LvwAnalisisTemporal.ListItems.Remove LvwAnalisisTemporal.SelectedItem.Index
                            LvwAnalisisTemporal.SetFocus
                            If LvwAnalisisTemporal.ListItems.Count <> 0 Then
                            
                                LvwAnalisisTemporal.SelectedItem = LvwAnalisisTemporal.ListItems(1)
                                Set objPozoAnalis = gObjPozoActivo.PozoAnalisis.ColPozoAnalis(gObjPozoActivo.IdPozo, Val(LvwAnalisisTemporal.SelectedItem.Key))
                                If Not objPozoAnalis Is Nothing Then
                                    BuscarComponentes objPozoAnalis, "TEMPORAL"
                                    BuscarDatosAnalisisTemporales objPozoAnalis
                                    BuscarRelacionesCromatograficasTemporal objPozoAnalis
                                End If
                            Else
                                LimpiatDatosTemporales
                            End If
                        End If
                    Else
                        gDatos.CommitTrans
                    End If
                    
                Case Is = "DESMARCAR"
                    
                    LvwAnalisisTemporal.SelectedItem.SmallIcon = 1
                    
                    objPozoAnalis.CodigoTipoDeAnalisis = 1
                    gDatos.BeginTrans
                    If gObjPozoActivo.PozoAnalisis.dbModificar(objPozoAnalis) Then
                        gDatos.CommitTrans
                    Else
                        gDatos.RollBackTrans
                    End If
                    
                Case Is = "ELIMINAR"
                
                    gDatos.BeginTrans
                    
                    If gObjPozoActivo.PozoAnalisis.dbEliminar(objPozoAnalis) Then
                        gDatos.CommitTrans
                        gObjPozoActivo.PozoAnalisis.colQuitar objPozoAnalis
                        LvwAnalisisTemporal.ListItems.Remove LvwAnalisisTemporal.SelectedItem.Index
                        LimpiatDatosTemporales
                    Else
                        gDatos.RollBackTrans
                    End If
                    
            End Select
        End If
        
    Else
        LvwAnalisisTemporal.SetFocus
    End If
    
End Sub

Private Function AnalisisTemporalBorrarList() As Boolean
    
    On Error GoTo Error
    
    AnalisisTemporalBorrarList = True
    
    LvwAnalisisTemporal.ListItems.Remove LvwAnalisisTemporal.SelectedItem.Index
    
    If LvwAnalisisTemporal.ListItems.Count <> 0 Then
        
        LvwAnalisisTemporal.SelectedItem.Selected = True
        
    End If
    
    LvwTemporal.ListItems.Clear
    
Error:
    
    If Err.Number <> 0 Then
        
        AnalisisTemporalBorrarList = False
        
        frmMsg.MostrarMsg "Módulo: AnalisisTemporalBorrarList." & Chr(10) & Err.Number & " - " & Err.Description, "Error", MdiPrincipal
        
        Err.Clear
        
    End If
    
End Function

Public Sub AnalisisTemporalSeleccionar()
    
    Dim StrSql                      As String
    Dim Temporal                    As Recordset
    Dim Item                        As ListItem
    Dim NumeroAnalisis              As Long
    Dim objPozoAnalis               As clsPozoAnalis
    
    Dim Profundidad                 As Long
    
    If LvwAnalisisTemporal.SelectedItem.SmallIcon = 1 Then
        Profundidad = LvwAnalisisTemporal.SelectedItem.SubItems(3)
        If Profundidad <> 0 Then
            NumeroAnalisis = Val(LvwAnalisisTemporal.SelectedItem.Key)
            Set objPozoAnalis = gObjPozoActivo.PozoAnalisis.ColPozoAnalis(gObjPozoActivo.IdPozo, NumeroAnalisis)
            If Not objPozoAnalis Is Nothing Then
                Set Item = LvwAnalisisDefinitivo.ListItems.Add(, NumeroAnalisis & "ID")
                Item.Tag = NumeroAnalisis
                Item.Text = LvwAnalisisTemporal.SelectedItem.Text
                Item.SubItems(1) = LvwAnalisisTemporal.SelectedItem.SubItems(1)
                Item.SubItems(2) = LvwAnalisisTemporal.SelectedItem.SubItems(2)
                Item.SubItems(3) = LvwAnalisisTemporal.SelectedItem.SubItems(3)
                Item.SmallIcon = LvwAnalisisTemporal.SelectedItem.SmallIcon
                
                objPozoAnalis.Seleccionado = True
                gDatos.BeginTrans
                If gObjPozoActivo.PozoAnalisis.dbModificar(objPozoAnalis) Then
                    gDatos.CommitTrans
                    LvwAnalisisTemporal.ListItems.Remove LvwAnalisisTemporal.SelectedItem.Index
                    pObjAdminEventos.ActualizarVistaProf
                    If LvwAnalisisTemporal.ListItems.Count <> 0 Then
                        LvwAnalisisTemporal.SelectedItem = LvwAnalisisTemporal.ListItems(1)
                        Set objPozoAnalis = gObjPozoActivo.PozoAnalisis.ColPozoAnalis(gObjPozoActivo.IdPozo, Val(LvwAnalisisTemporal.SelectedItem.Key))
                        If Not objPozoAnalis Is Nothing Then
                            BuscarComponentes objPozoAnalis, "TEMPORAL"
                            BuscarDatosAnalisisTemporales objPozoAnalis
                            BuscarRelacionesCromatograficasTemporal objPozoAnalis
                            
                        End If
                    Else
                        LimpiatDatosTemporales
                    End If
                Else
                    gDatos.RollBackTrans
                End If
            End If
        Else
            frmMsg.MostrarMsg "Debe ingresar la profundidad antes de seleccionar el análisis", "Error", MdiPrincipal
            SSTab1.Tab = 0
            TxtProfundidadTemporal.SetFocus
        End If
    Else
        frmMsg.MostrarMsg "Solamente se pueden seleccionar análisis normales", "Error", MdiPrincipal
    End If
    
Error:
    
    If Err.Number <> 0 Then
        
        frmMsg.MostrarMsg "Módulo: AnalisisTemporalSeleccionar." & Chr(10) & Err.Number & " - " & Err.Description, "Error", MdiPrincipal
        
        Err.Clear
        
    End If
    
End Sub

Public Function BuscarAnalisis() As Boolean
    
    Dim StrSql                          As String
    Dim Analisis                        As Recordset
    Dim Item                            As ListItem
    Dim objPozoAnalis                   As clsPozoAnalis
    On Error GoTo Error
    
    BuscarAnalisis = True
    
    LvwAnalisisDefinitivo.ListItems.Clear
    LvwAnalisisTemporal.ListItems.Clear
    
    If Not gObjPozoActivo Is Nothing Then
        
        Set gObjPozoActivo.PozoAnalisis.Datos = gDatos
        gObjPozoActivo.PozoAnalisis.ConsultarPozoAnalisis gObjPozoActivo.IdPozo
        
        For Each objPozoAnalis In gObjPozoActivo.PozoAnalisis
            If objPozoAnalis.Seleccionado Then
                Set Item = LvwAnalisisDefinitivo.ListItems.Add(, objPozoAnalis.NumeroAnalisis & "ID")
                ArchivoDefinitivo.caption = objPozoAnalis.archivo
            Else
                Set Item = LvwAnalisisTemporal.ListItems.Add(, objPozoAnalis.NumeroAnalisis & "ID")
                ArchivoTemporal.caption = objPozoAnalis.archivo
            End If
            
            Item.Tag = objPozoAnalis.CodigoTipoDeAnalisis
            
            Item.Text = Format(objPozoAnalis.NumeroAnalisis, "00000")
            Item.SubItems(1) = objPozoAnalis.NumeroAnalisis
            Item.SubItems(2) = Format(objPozoAnalis.fecha, "DD/MM/YYYY HH:MM:SS")
            Item.SubItems(3) = objPozoAnalis.Profundidad
            
            
            Select Case objPozoAnalis.CodigoTipoDeAnalisis
                Case Is = 1
                    Item.SmallIcon = 1
                Case Is = 2
                    Item.SmallIcon = 2
                Case Is = 3
                    Item.SmallIcon = 3
                Case Is = 4
                    Item.SmallIcon = 4
                Case Is = 5
                    Item.SmallIcon = 5
            End Select
                
        Next
        If LvwAnalisisDefinitivo.ListItems.Count <> 0 Then
            
            LvwAnalisisDefinitivo.SelectedItem = LvwAnalisisDefinitivo.ListItems(1)
            LvwAnalisisDefinitivo.SelectedItem.Selected = True
            
        End If
        
        If LvwAnalisisTemporal.ListItems.Count <> 0 Then
            
            LvwAnalisisTemporal.SelectedItem = LvwAnalisisTemporal.ListItems(1)
            LvwAnalisisTemporal.SelectedItem.Selected = True
            
        End If
    End If
Error:
    
    If Err.Number <> 0 Then
        
        BuscarAnalisis = False
        
        frmMsg.MostrarMsg "Módulo: BuscarAnalisis." & Chr(10) & Err.Number & " - " & Err.Description, "Error", MdiPrincipal
        
        
        Err.Clear
        
    End If
    
End Function

Public Sub ComponenteTemporalBorrar()
    
    Dim NumeroAnalisis                          As Long
    Dim NumeroComponente                        As Long
    Dim objPozoAnalis                           As clsPozoAnalis
    Dim objPozoAnalisComponente                 As ClsPozoAnalisComponente
    Dim Respuesta                               As Integer
    Dim ok                                      As Boolean
    
    NumeroAnalisis = Val(LvwAnalisisTemporal.SelectedItem.Key)
    NumeroComponente = Val(LvwTemporal.SelectedItem.Key)
        
    Respuesta = MsgBox("Está seguro que desea borrar el Componente seleccionado", vbYesNo + vbQuestion)
        
    If Respuesta = vbYes Then
        If Not gObjPozoActivo Is Nothing Then
            Set objPozoAnalis = gObjPozoActivo.PozoAnalisis.ColPozoAnalis(gObjPozoActivo.IdPozo, NumeroAnalisis)
            If Not objPozoAnalis Is Nothing Then
                Set objPozoAnalisComponente = objPozoAnalis.PozoAnalisComponentes.colComponente(objPozoAnalis.IdPozo, objPozoAnalis.NumeroAnalisis, NumeroComponente)
                If Not objPozoAnalisComponente Is Nothing Then
                    gDatos.BeginTrans
                    Set objPozoAnalis.PozoAnalisComponentes.Datos = gDatos
                    If objPozoAnalis.PozoAnalisComponentes.dbEliminar(objPozoAnalisComponente) Then
                        objPozoAnalis.PozoAnalisComponentes.colQuitar objPozoAnalisComponente
                        If ComponenteTemporalBorrarList Then
                            CalcularRelacionesCromatograficas objPozoAnalis, "MODIFICAR", ok
                            If ok Then
                                ActualizarGasTotalCromatografico objPozoAnalis, ok
                                If ok Then
                                    BuscarRelacionesCromatograficasTemporal objPozoAnalis
                                    BuscarDatosAnalisisTemporales objPozoAnalis
                                    BuscarComponentes objPozoAnalis, "TEMPORAL"
                                    gDatos.CommitTrans
                                    LvwTemporal.SetFocus
                                Else
                                    gDatos.RollBackTrans
                                    frmMsg.MostrarMsg "No se pudo actualizar el gas total cromatográfico", "Error", MdiPrincipal
                                    
                                End If
                            Else
                                gDatos.RollBackTrans
                                frmMsg.MostrarMsg "No se pudieron actualizar las relaciones cromatográficas", "Error", MdiPrincipal
                            End If
                        Else
                            gDatos.RollBackTrans
                            Unload Me
                        End If
                    Else
                        gDatos.RollBackTrans
                        LvwTemporal.SetFocus
                    End If
                End If
            End If
        End If
    Else
        LvwTemporal.SetFocus
    End If
    
End Sub


Private Function ComponenteTemporalBorrarList() As Boolean
    
    On Error GoTo Error
    
    ComponenteTemporalBorrarList = True
    
    LvwTemporal.ListItems.Remove LvwTemporal.SelectedItem.Index
    
    If LvwTemporal.ListItems.Count <> 0 Then
        
        LvwTemporal.SelectedItem.Selected = True
        
    End If
    
Error:
    
    If Err.Number <> 0 Then
        
        ComponenteTemporalBorrarList = False
        
        frmMsg.MostrarMsg "Módulo: ComponenteTemporalBorrarList." & Chr(10) & Err.Number & " - " & Err.Description, "Error", MdiPrincipal
        
        Err.Clear
        
    End If
    
End Function

Public Sub ComponenteTemporalModificar()
    
    Dim Retention                       As Double
    Dim Area                            As Double
    Dim Externo                         As Long
    Dim Units                           As String
    Dim NormArea                        As Double
    Dim NumeroAnalisis                  As Long
    Dim objPozoAnalis                   As clsPozoAnalis
    Dim objPozoAnalisComponente         As ClsPozoAnalisComponente
    If Not gObjPozoActivo Is Nothing Then
        Set objPozoAnalis = gObjPozoActivo.PozoAnalisis.ColPozoAnalis(gObjPozoActivo.IdPozo, Val(LvwAnalisisTemporal.SelectedItem.Key))
        If Not objPozoAnalis Is Nothing Then
            Set objPozoAnalisComponente = objPozoAnalis.PozoAnalisComponentes.colComponente(objPozoAnalis.IdPozo, objPozoAnalis.NumeroAnalisis, Val(LvwTemporal.SelectedItem.Key))
            
            If Not objPozoAnalisComponente Is Nothing Then
            
                Set FrmDatosComponente.gObjPozoAnalis = objPozoAnalis
                Set FrmDatosComponente.gObjPozoAnalisComponente = objPozoAnalisComponente
                FrmDatosComponente.Accion = euAccionModificar
                FrmDatosComponente.Show vbModal
                If FormularioCargado("FrmDatosComponente") Then
                    Retention = FrmDatosComponente.TxtRetention.Text
                    Area = FrmDatosComponente.TxtArea.Text
                    Externo = FrmDatosComponente.TxtExternal.Text
                    Units = FrmDatosComponente.TxtUnits.Text
                    NormArea = FrmDatosComponente.TxtNormArea.Text
                    Unload FrmDatosComponente
                    BuscarRelacionesCromatograficasTemporal objPozoAnalis
                    BuscarDatosAnalisisTemporales objPozoAnalis
                    BuscarComponentes objPozoAnalis, "TEMPORAL"
                    LvwTemporal.SetFocus
                Else
                    LvwTemporal.SetFocus
                End If
            End If
            
        End If
    End If
End Sub

Public Sub AnalisisDefinitivoBorrar()
    
    Dim Item                                As ListItem
    Dim NumeroAnalisis                      As Long
    Dim fecha                               As Date
    Dim objPozoAnalis                       As clsPozoAnalis
    Dim objPozoAnalisComponente             As ClsPozoAnalisComponente
        
    NumeroAnalisis = Val(LvwAnalisisDefinitivo.SelectedItem.Key)
    Set objPozoAnalis = gObjPozoActivo.PozoAnalisis.ColPozoAnalis(gObjPozoActivo.IdPozo, NumeroAnalisis)
    If Not objPozoAnalis Is Nothing Then
        Set Item = LvwAnalisisTemporal.ListItems.Add(, NumeroAnalisis & "ID")
        Item.Tag = NumeroAnalisis
        Item.Text = LvwAnalisisDefinitivo.SelectedItem.Text
        Item.SubItems(1) = LvwAnalisisDefinitivo.SelectedItem.SubItems(1)
        Item.SubItems(2) = LvwAnalisisDefinitivo.SelectedItem.SubItems(2)
        Item.SubItems(3) = LvwAnalisisDefinitivo.SelectedItem.SubItems(3)
        Item.SmallIcon = LvwAnalisisDefinitivo.SelectedItem.SmallIcon
        objPozoAnalis.Seleccionado = False
        If gObjPozoActivo.PozoAnalisis.dbModificar(objPozoAnalis) Then
            LvwAnalisisDefinitivo.ListItems.Remove LvwAnalisisDefinitivo.SelectedItem.Index
            LvwAnalisisDefinitivo.SetFocus
            If LvwAnalisisDefinitivo.ListItems.Count <> 0 Then
                LvwAnalisisDefinitivo.SelectedItem = LvwAnalisisDefinitivo.ListItems(1)
                Set objPozoAnalis = gObjPozoActivo.PozoAnalisis.ColPozoAnalis(gObjPozoActivo.IdPozo, Val(LvwAnalisisDefinitivo.SelectedItem.Key))
                If Not objPozoAnalis Is Nothing Then
                    BuscarComponentes objPozoAnalis, "DEFINITIVO"
                    BuscarDatosAnalisisDefinitivos objPozoAnalis
                    BuscarRelacionesCromatograficasDefinitivo objPozoAnalis
                End If
            Else
                LimpiatDatosDefinitivos
            End If
        End If
    End If
Error:
    
    If Err.Number <> 0 Then
        
        frmMsg.MostrarMsg "Módulo: AnalisisTemporalSeleccionar." & Chr(10) & Err.Number & " - " & Err.Description, "Error", MdiPrincipal
        
        Err.Clear
        
    End If
    
End Sub

Private Function BuscarDatosAnalisisTemporales(ByVal objPozoAnalis As clsPozoAnalis) As Boolean
    
    Dim StrSql                          As String
    Dim Temporal                        As Recordset
    Dim Item                            As ListItem
    
    On Error GoTo Error
    
    TxtProfundidadTemporal.Text = objPozoAnalis.Profundidad
    TxtCO2Temporal.Text = objPozoAnalis.CO2
    TxtSH2Temporal.Text = objPozoAnalis.SH2
    TxtGasTotalTemporal.Text = objPozoAnalis.GasTotal
    TxtGasTotalCromatograficoTemporal.Text = objPozoAnalis.GasTotalCromatografico
    
    If LvwAnalisisTemporal.SelectedItem.SmallIcon = 1 Then
        TxtProfundidadTemporal.Enabled = True
    Else
        TxtProfundidadTemporal.Enabled = False
    End If
    
Error:
    
    If Err.Number <> 0 Then
        
        BuscarDatosAnalisisTemporales = False
        
        frmMsg.MostrarMsg "Módulo: BuscarComponentesTemporales." & Chr(10) & Err.Number & " - " & Err.Description, "Error", MdiPrincipal
        
        Err.Clear
        
    End If
    
End Function

Private Function BuscarDatosAnalisisDefinitivos(ByVal objPozoAnalis As clsPozoAnalis) As Boolean
    
    On Error GoTo Error
    
    BuscarDatosAnalisisDefinitivos = True
    
    TxtProfundidadDefinitivo.Text = objPozoAnalis.Profundidad
    TxtCO2Definitivo.Text = objPozoAnalis.CO2
    TxtSH2Definitivo.Text = objPozoAnalis.SH2
    TxtGasTotalDefinitivo.Text = objPozoAnalis.GasTotal
    TxtGasTotalCromatograficoDefinitivo.Text = objPozoAnalis.GasTotalCromatografico
            
    If Not IsNull(objPozoAnalis.Observaciones) Then
    
    End If
    
Error:
    
    If Err.Number <> 0 Then
        BuscarDatosAnalisisDefinitivos = False
        frmMsg.MostrarMsg "Módulo: BuscarComponentesDefinitivos." & Chr(10) & Err.Number & " - " & Err.Description, "Error", MdiPrincipal
        Err.Clear
    End If
    
End Function

Private Function BuscarRelacionesCromatograficasTemporal(ByVal objPozoAnalis As clsPozoAnalis) As Boolean
    
    Dim Item As ListItem
    
    On Error GoTo Error
    
    BuscarRelacionesCromatograficasTemporal = True
    
            
    TxtBar2Temporal.Text = objPozoAnalis.Bar2
    TxtBar3Temporal.Text = objPozoAnalis.Bar3
    TxtBar4Temporal.Text = objPozoAnalis.Bar4
    TxtBar5Temporal.Text = objPozoAnalis.Bar5
    TxtCous1Temporal.Text = objPozoAnalis.Cous1
    TxtCous2Temporal.Text = objPozoAnalis.Cous2
    TxtWHTemporal.Text = objPozoAnalis.WH
    TxtBHTemporal.Text = objPozoAnalis.BH
    TxtCHTemporal.Text = objPozoAnalis.CH
    TxtGeo1Temporal.Text = objPozoAnalis.Geo1
    TxtGeo2Temporal.Text = objPozoAnalis.Geo2
    TxtGeo3Temporal.Text = objPozoAnalis.Geo3
    TxtGeo4Temporal.Text = objPozoAnalis.Geo4
    TxtSnGeoTemporal.Text = objPozoAnalis.SnGeo
    TxtC1Temporal.Text = objPozoAnalis.C1
    TxtC2Temporal.Text = objPozoAnalis.C2
    TxtC3Temporal.Text = objPozoAnalis.C3
    TxtC4Temporal.Text = objPozoAnalis.C4
    TxtC5Temporal.Text = objPozoAnalis.C5
Error:
    
    If Err.Number <> 0 Then
        BuscarRelacionesCromatograficasTemporal = False
        frmMsg.MostrarMsg "Módulo: BuscarRelacionesCromatograficasTemporal." & Chr(10) & Err.Number & " - " & Err.Description, "Error", MdiPrincipal
        Err.Clear
    End If
    
End Function

Private Function BuscarRelacionesCromatograficasDefinitivo(ByVal objPozoAnalis As clsPozoAnalis) As Boolean
    
    Dim Item As ListItem
    
    On Error GoTo Error
    
    BuscarRelacionesCromatograficasDefinitivo = True
            
    TxtBar2Definitivo.Text = objPozoAnalis.Bar2
    TxtBar3Definitivo.Text = objPozoAnalis.Bar3
    TxtBar4Definitivo.Text = objPozoAnalis.Bar4
    TxtBar5Definitivo.Text = objPozoAnalis.Bar5
    TxtCous1Definitivo.Text = objPozoAnalis.Cous1
    TxtCous2Definitivo.Text = objPozoAnalis.Cous2
    TxtWHDefinitivo.Text = objPozoAnalis.WH
    TxtBHDefinitivo.Text = objPozoAnalis.BH
    TxtCHDefinitivo.Text = objPozoAnalis.CH
    TxtGeo1Definitivo.Text = objPozoAnalis.Geo1
    TxtGeo2Definitivo.Text = objPozoAnalis.Geo2
    TxtGeo3Definitivo.Text = objPozoAnalis.Geo3
    TxtGeo4Definitivo.Text = objPozoAnalis.Geo4
    TxtSnGeoDefinitivo.Text = objPozoAnalis.SnGeo
    TxtC1Definitivo.Text = objPozoAnalis.C1
    TxtC2Definitivo.Text = objPozoAnalis.C2
    TxtC3Definitivo.Text = objPozoAnalis.C3
    TxtC4Definitivo.Text = objPozoAnalis.C4
    TxtC5Definitivo.Text = objPozoAnalis.C5
    
Error:
    
    If Err.Number <> 0 Then
        BuscarRelacionesCromatograficasDefinitivo = False
        frmMsg.MostrarMsg "Módulo: BuscarRelacionesCromatograficasDefinitivo." & Chr(10) & Err.Number & " - " & Err.Description, "Error", MdiPrincipal
        Err.Clear
    End If
    
End Function

Public Sub AnalisisCalibracion()
    
    Dim NumeroAnalisis              As Long
    Dim objPozoAnalis               As clsPozoAnalis
    
    NumeroAnalisis = Val(LvwAnalisisTemporal.SelectedItem.Key)
    Set objPozoAnalis = gObjPozoActivo.PozoAnalisis.ColPozoAnalis(gObjPozoActivo.IdPozo, NumeroAnalisis)
    If Not objPozoAnalis Is Nothing Then
        If LvwAnalisisTemporal.SelectedItem.SmallIcon = 2 Then
            LvwAnalisisTemporal.SelectedItem.SmallIcon = 1
            objPozoAnalis.CodigoTipoDeAnalisis = 1
            gObjPozoActivo.PozoAnalisis.dbModificar objPozoAnalis
            TxtProfundidadTemporal.Enabled = True
        Else
            LvwAnalisisTemporal.SelectedItem.SmallIcon = 2
            objPozoAnalis.CodigoTipoDeAnalisis = 2
            objPozoAnalis.Profundidad = 0
            gObjPozoActivo.PozoAnalisis.dbModificar objPozoAnalis
            LvwAnalisisTemporal.SelectedItem.SubItems(3) = 0
            TxtProfundidadTemporal.Text = 0
            TxtProfundidadTemporal.Enabled = False
            
        End If
        
        If ChkMostrarTodosLosAnálisisTemporal.Value = 0 Then
            LvwAnalisisTemporal.ListItems.Remove LvwAnalisisTemporal.SelectedItem.Index
            LvwAnalisisTemporal.SetFocus
            If LvwAnalisisTemporal.ListItems.Count <> 0 Then
                
                LvwAnalisisTemporal.SelectedItem = LvwAnalisisTemporal.ListItems(1)
                Set objPozoAnalis = gObjPozoActivo.PozoAnalisis.ColPozoAnalis(gObjPozoActivo.IdPozo, Val(LvwAnalisisTemporal.SelectedItem.Key))
                If Not objPozoAnalis Is Nothing Then
                    BuscarComponentes objPozoAnalis, "TEMPORAL"
                    BuscarDatosAnalisisTemporales objPozoAnalis
                    BuscarRelacionesCromatograficasTemporal objPozoAnalis
                End If
            Else
                LimpiatDatosTemporales
            End If
            
        End If
    End If
End Sub

Public Sub LimpiatDatosTemporales()
    
    LvwTemporal.ListItems.Clear
    
    TxtBar2Temporal.Text = ""
    TxtBar3Temporal.Text = ""
    TxtBar4Temporal.Text = ""
    TxtBar5Temporal.Text = ""
    TxtCous1Temporal.Text = ""
    TxtCous2Temporal.Text = ""
    TxtWHTemporal.Text = ""
    TxtBHTemporal.Text = ""
    TxtCHTemporal.Text = ""
    TxtGeo1Temporal.Text = ""
    TxtGeo2Temporal.Text = ""
    TxtGeo3Temporal.Text = ""
    TxtGeo4Temporal.Text = ""
    TxtSnGeoTemporal.Text = ""
    TxtC1Temporal.Text = ""
    TxtC2Temporal.Text = ""
    TxtC3Temporal.Text = ""
    TxtC4Temporal.Text = ""
    TxtC5Temporal.Text = ""
    
    TxtProfundidadTemporal.Text = ""
    TxtCO2Temporal.Text = ""
    TxtSH2Temporal.Text = ""
    TxtGasTotalTemporal.Text = ""
    TxtGasTotalCromatograficoTemporal.Text = ""
    
End Sub

Public Sub LimpiatDatosDefinitivos()
    
    LvwDefinitivo.ListItems.Clear
    
    TxtBar2Definitivo.Text = ""
    TxtBar3Definitivo.Text = ""
    TxtBar4Definitivo.Text = ""
    TxtBar5Definitivo.Text = ""
    TxtCous1Definitivo.Text = ""
    TxtCous2Definitivo.Text = ""
    TxtWHDefinitivo.Text = ""
    TxtBHDefinitivo.Text = ""
    TxtCHDefinitivo.Text = ""
    TxtGeo1Definitivo.Text = ""
    TxtGeo2Definitivo.Text = ""
    TxtGeo3Definitivo.Text = ""
    TxtGeo4Definitivo.Text = ""
    TxtSnGeoDefinitivo.Text = ""
    TxtC1Definitivo.Text = ""
    TxtC2Definitivo.Text = ""
    TxtC3Definitivo.Text = ""
    TxtC4Definitivo.Text = ""
    TxtC5Definitivo.Text = ""
    
    TxtProfundidadDefinitivo.Text = ""
    TxtCO2Definitivo.Text = ""
    TxtSH2Definitivo.Text = ""
    TxtGasTotalDefinitivo.Text = ""
    TxtGasTotalCromatograficoDefinitivo.Text = ""
    
End Sub

Private Sub pObjAdminEventos_EvtActualizarValoresGases()
    
    DisplayDioxidoCarbono.Value = IIf(CO2 > 0, CO2, 0)
    DisplayGasTotal.Value = IIf(GasTotal > 0, GasTotal, 0)
    DisplaySulfidrico.Value = IIf(SH2 > 0, SH2, 0)

End Sub

Private Sub pObjAdminEventos_EvtActualizarVistaProf()

    Dim C1                                  As Double
    Dim C2                                  As Double
    Dim C3                                  As Double
    Dim IC4                                 As Double
    Dim NC4                                 As Double
    Dim IC5                                 As Double
    Dim NC5                                 As Double
    Dim objPozoEscaneo                      As clsPozoEscaneo
    Dim objPozoEscaneos                     As clsPozoEscaneos
    Dim profinicio                          As Double
    Dim proffinal                           As Double
    Dim Indice                              As Long
    
    Dim objPozoAnalis                       As New clsPozoAnalis
    Dim objPozoAnalisComponente             As ClsPozoAnalisComponente

    On Error GoTo Error

    Me.MousePointer = 13
    profinicio = CDbl(TxtProfundidadInicio.Text)
    proffinal = CDbl(TxtProfundidadInicio.Text) + CDbl(TxtSpan.Text)
    
    iPlotMasterLog.ClearAllData
    iPlotMasterLog.XAxis(0).Min = profinicio
    iPlotMasterLog.XAxis(0).Span = Val(TxtSpan)
    
    For Each objPozoAnalis In gObjPozoActivo.PozoAnalisis
        If objPozoAnalis.Seleccionado Then
            iPlotMasterLog.Channel(2).AddXY objPozoAnalis.Profundidad, objPozoAnalis.CO2
            C1 = 0
            C2 = 0
            C3 = 0
            IC4 = 0
            NC4 = 0
            IC5 = 0
            NC5 = 0
            For Each objPozoAnalisComponente In objPozoAnalis.PozoAnalisComponentes
                Select Case objPozoAnalisComponente.NumeroComponente
                    Case 1: C1 = objPozoAnalisComponente.Externo
                    Case 2: C2 = objPozoAnalisComponente.Externo
                    Case 3: C3 = objPozoAnalisComponente.Externo
                    Case 4: IC4 = objPozoAnalisComponente.Externo
                    Case 5: NC4 = objPozoAnalisComponente.Externo
                    Case 6: IC5 = objPozoAnalisComponente.Externo
                    Case 7: NC5 = objPozoAnalisComponente.Externo
                End Select
            Next
            iPlotMasterLog.Channel(3).AddXY objPozoAnalis.Profundidad, C1
            iPlotMasterLog.Channel(4).AddXY objPozoAnalis.Profundidad, C2
            iPlotMasterLog.Channel(5).AddXY objPozoAnalis.Profundidad, C3
            iPlotMasterLog.Channel(6).AddXY objPozoAnalis.Profundidad, IC4
            iPlotMasterLog.Channel(7).AddXY objPozoAnalis.Profundidad, NC4
            iPlotMasterLog.Channel(8).AddXY objPozoAnalis.Profundidad, IC5
            iPlotMasterLog.Channel(9).AddXY objPozoAnalis.Profundidad, NC5
        End If
    Next
    Indice = gObjPozoActivo.PozoEscaneos.Count
    Do Until Indice = 0
        Set objPozoEscaneo = gObjPozoActivo.PozoEscaneos.ColPozoEscaneoPorIndice(Indice)
        If Not objPozoEscaneo Is Nothing Then
            If objPozoEscaneo.ProfundidadPozo >= profinicio And objPozoEscaneo.ProfundidadPozo <= proffinal Then
                iPlotMasterLog.Channel(0).AddXY objPozoEscaneo.ProfundidadPozo, objPozoEscaneo.Crono
                iPlotMasterLog.Channel(1).AddXY objPozoEscaneo.ProfundidadPozo, objPozoEscaneo.GasTotal
            End If
        End If
        Indice = Indice - 1
    Loop
    
Error:
    If Err.Number <> 0 Then
        frmMsg.MostrarMsg "Módulo: ConsultarTodos los cronos ." & Chr(10) & "Ocurrió el error: " & Err.Number & " - " & Err.Description, "Error", MdiPrincipal
        Err.Clear
    End If
    Me.MousePointer = 0
    
End Sub

Private Sub pObjAdminEventos_EvtActualizarVistaTiempo(ByVal Intervalo As Long)
    Dim fecha As Double
    Dim Index As Long
    
    If ChkTiempoReal = 1 Then
        If (ChkTiempoReal = 1) And (Intervalo Mod 3 = 0) Then
            fecha = Month(Now()) * 30 + Day(Now()) + Hour(Now()) / 24 + Minute(Now()) / (24 * 60) + Second(Now()) / (24# * 60# * 60#)
            If Intervalo Mod 48 = 0 Or gRefrescarGrafico Then
                LimpiarYConsultarDatosTiempo
            Else
                iPlotGases.Channel(0).AddXY fecha, GasTotal
                iPlotGases.Channel(1).AddXY fecha, CO2
                
                iPlotGases.XAxis(0).Min = DateAdd("n", -1 * gSpan, fecha)
                Do While iPlotGases.Channel(0).DataX(0) < iPlotGases.XAxis(0).Min
                    iPlotGases.Channel(0).DeletePoints 1
                    iPlotGases.Channel(1).DeletePoints 1
                Loop
                iPlotGases.Visible = False
                iPlotGases.Visible = True
                If ComentarioGas <> "" Then
                     Index = iPlotGases.AddAnnotation
                     iPlotGases.Annotation(Index).Font.Size = 12
                     iPlotGases.Annotation(Index).Reference = iprtChannel
                     iPlotGases.Annotation(Index).ChannelName = iPlotGases.Channel(0).Name
                     iPlotGases.Annotation(Index).Y = 13000 'Center X Coordinate
                     iPlotGases.Annotation(Index).X = fecha 'Center Y Coordinate
                     iPlotGases.Annotation(Index).Style = ipasText 'Text Annotation
                     iPlotGases.Annotation(Index).FontColor = vbWhite 'White Font
                     iPlotGases.Annotation(Index).Text = ComentarioGas
                     iPlotGases.Annotation(Index).TextRotation = ira000   'Rotate up-side-down
                    If ComentarioGas <> "" Then
                        ComentarioGas = ""
                    End If
                End If
            End If
        End If
    End If

End Sub

Private Sub SwTha_OnMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

        Select Case SwTha.Position
            Case 0: EContinuoLocal = 1
            Case 1: EContinuoLocal = 10
            Case 2: EContinuoLocal = 100
        End Select
        
        If gObjConfiguracion.IdTipoEquipoGas = 1 Then
            MdiPrincipal.TasSerial1.ScanActive = False
            MdiPrincipal.TasSerial2.ScanActive = False
            MdiPrincipal.TasSerial1.AbortCommunication
            MdiPrincipal.TasSerial2.AbortCommunication
            MdiPrincipal.TasSerial1.Wait 1000
            MdiPrincipal.TasSerial1.ScanActive = True
            MdiPrincipal.TasSerial2.ScanActive = True
            MdiPrincipal.TasSerial2.Trigger
            MdiPrincipal.TimerCroma.Enabled3 = True
            
        ElseIf gObjConfiguracion.IdTipoEquipoGas = 3 Then
            MdiPrincipal.Salida_rabbit 1
        
        End If
        
End Sub

Private Sub SwTha_OnMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    SwTha.Tag = SwTha.Position

End Sub


Private Sub SwNormal_OnMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    SwNormal.Tag = SwNormal.Position

End Sub


Private Sub SwFast_OnMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

    SwFast.Tag = SwFast.Position

End Sub


Private Sub SwNormal_OnMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)
    
    
        Select Case SwNormal.Position
        
            Case 0: ENormalLocal = 1
                            
            Case 1: ENormalLocal = 10
            
            Case 2: ENormalLocal = 100
            
        End Select
        
        If gObjConfiguracion.IdTipoEquipoGas = 1 Then
            MdiPrincipal.TasSerial1.ScanActive = False
            MdiPrincipal.TasSerial2.ScanActive = False
        
            MdiPrincipal.TasSerial1.AbortCommunication
            MdiPrincipal.TasSerial2.AbortCommunication
        
            MdiPrincipal.TasSerial1.Wait 1000
        
            MdiPrincipal.TasSerial1.ScanActive = True
            MdiPrincipal.TasSerial2.ScanActive = True
        
        
            MdiPrincipal.TasSerial2.Trigger
            
            MdiPrincipal.TimerCroma.Enabled3 = True
        ElseIf gObjConfiguracion.IdTipoEquipoGas = 3 Then
            MdiPrincipal.Salida_rabbit 2
        End If


End Sub

Private Sub SwFast_OnMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)


    
        Select Case SwFast.Position
        
            Case 0: EFastLocal = 1
                            
            Case 1: EFastLocal = 10
            
            Case 2: EFastLocal = 100
            
        End Select
        
        If gObjConfiguracion.IdTipoEquipoGas = 1 Then
            MdiPrincipal.TasSerial1.ScanActive = False
            MdiPrincipal.TasSerial2.ScanActive = False
        
            MdiPrincipal.TasSerial1.AbortCommunication
            MdiPrincipal.TasSerial2.AbortCommunication
        
            MdiPrincipal.TasSerial1.Wait 1000
        
            MdiPrincipal.TasSerial1.ScanActive = True
            MdiPrincipal.TasSerial2.ScanActive = True
        
        
            MdiPrincipal.TasSerial2.Trigger
        
            MdiPrincipal.TimerCroma.Enabled3 = True
        ElseIf gObjConfiguracion.IdTipoEquipoCroma = 3 Then
            MdiPrincipal.Salida_rabbit 3
        End If

End Sub

Private Sub Timer1_Timer()
 If ContadorTicks Mod 15 = 0 Then
        OrdenarPantallaImpresion iPlotGases
        ContadorTicks = 1
    Else
        ContadorTicks = ContadorTicks + 1
    End If
End Sub

Public Sub OrdenarPantallaImpresion(UnGrafico As iPlotX)

Dim CadenaId As String
Dim AuxHeight As Long
Dim AuxWidth As Long
Dim Indice As Long
Dim CantidadAnotaciones As Long


On Error GoTo Error
    
    AuxWidth = UnGrafico.Width
    AuxHeight = UnGrafico.Height

    UnGrafico.Height = 16000
    UnGrafico.Width = 12000



UnGrafico.PrintOrientation = poPortrait
'UnGrafico.PrintOrientation = poLandscape
UnGrafico.PrintShowDialog = False

UnGrafico.PrintMarginRight = 0
UnGrafico.PrintMarginLeft = 0
UnGrafico.PrintMarginTop = 0
UnGrafico.PrintMarginBottom = 0



UnGrafico.BeginUpdate
'Set background colors for Chart and DataView areas to a light color or white
UnGrafico.BackGroundColor = vbWhite




For Indice = 0 To UnGrafico.YAxisCount - 1

    UnGrafico.YAxis(Indice).TitleFontColor = vbBlack
    UnGrafico.YAxis(Indice).ScaleLinesColor = vbBlack
    UnGrafico.YAxis(Indice).LabelsFontColor = vbBlack
    
Next

For Indice = 0 To UnGrafico.XAxisCount - 1

    UnGrafico.XAxis(Indice).LabelsFontColor = vbBlack
    
    
Next


For Indice = 0 To UnGrafico.AnnotationCount
    
    If Indice <> 0 Then
        UnGrafico.Annotation(Indice - 1).FontColor = vbBlack
        
    End If

Next





UnGrafico.Height = 14000
'UnGrafico.Annotation(0).FontColor = vbBlack

'UnGrafico.PrintChart


'CadenaId = UnGrafico.Name


CadenaId = "c:\Registros\Gas\"
    
CadenaId = CadenaId & Convertir(Now())

UnGrafico.SaveImageToJPEG CadenaId & ".jpg", 50, True

UnGrafico.Width = AuxWidth
UnGrafico.Height = AuxHeight


'Set background colors back to their original settings
UnGrafico.BackGroundColor = vbBlack

For Indice = 0 To UnGrafico.YAxisCount - 1

    UnGrafico.YAxis(Indice).TitleFontColor = vbWhite
    UnGrafico.YAxis(Indice).ScaleLinesColor = vbWhite
    UnGrafico.YAxis(Indice).LabelsFontColor = vbWhite
    
Next

For Indice = 0 To UnGrafico.XAxisCount - 1

    UnGrafico.XAxis(Indice).LabelsFontColor = vbWhite
    
Next

For Indice = 0 To UnGrafico.AnnotationCount
    If Indice <> 0 Then
        UnGrafico.Annotation(Indice - 1).FontColor = vbWhite
    End If
Next


UnGrafico.EndUpdate

Error:
        
        UnGrafico.Width = AuxWidth
        UnGrafico.Height = AuxHeight

        
        'Set background colors back to their original settings
        UnGrafico.BackGroundColor = vbBlack
        
        For Indice = 0 To UnGrafico.YAxisCount - 1
        
            UnGrafico.YAxis(Indice).TitleFontColor = vbWhite
            UnGrafico.YAxis(Indice).ScaleLinesColor = vbWhite
            UnGrafico.YAxis(Indice).LabelsFontColor = vbWhite
            
        Next
        
        For Indice = 0 To UnGrafico.XAxisCount - 1
        
            UnGrafico.XAxis(Indice).LabelsFontColor = vbWhite
            
        Next
        
        For Indice = 0 To UnGrafico.AnnotationCount
            If Indice <> 0 Then
                UnGrafico.Annotation(Indice - 1).FontColor = vbWhite
            End If
        Next


        Err.Clear
    

End Sub
Private Function Convertir(LaFecha As String) As String
    Dim Indice As Long
    Dim CadenaAux As String

    CadenaAux = ""
    
    For Indice = 1 To Len(LaFecha)
        If Mid(LaFecha, Indice, 1) = "/" Or _
            Mid(LaFecha, Indice, 1) = ":" Or _
                Mid(LaFecha, Indice, 1) = "." Or _
                    Mid(LaFecha, Indice, 1) = " " _
            Then
                CadenaAux = CadenaAux & "!"
            Else: CadenaAux = CadenaAux & Mid(LaFecha, Indice, 1)
        End If
    Next
    
    Convertir = CadenaAux

End Function



Private Sub TxtCO2_GotFocus()
    
    TextSelected
    
End Sub

Private Sub TxtCO2_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        
        SendKeys "{TAB}"
        
'    ElseIf KeyAscii = 46 Then
'
'        If InStr(1, TxtCO2.Text, ".") <> 0 Then
'
'            KeyAscii = 0
'
'        End If
        
    ElseIf (KeyAscii > 47 And KeyAscii < 58) Or (KeyAscii = 8) Then
        
        'Tecla permitida
        
    Else
        
        KeyAscii = 0
        
    End If
    
End Sub

Private Sub TxtCO2Definitivo_GotFocus()
    
    TextSelected
    
End Sub

Private Sub TxtCO2Definitivo_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        
        SendKeys "{TAB}"
        
    ElseIf (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8 Then
        
        'Tecla permitida
        
    Else
        
        KeyAscii = 0
        
    End If
    
End Sub

Private Sub TxtCO2Temporal_GotFocus()
    
    TextSelected
    
End Sub

Private Sub TxtCO2Temporal_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        
        SendKeys "{TAB}"
        
    ElseIf (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8 Then
        
        'Tecla permitida
        
    Else
        
        KeyAscii = 0
        
    End If
    
End Sub

Private Sub TxtCrono_GotFocus()
    
    TextSelected
    
End Sub

Private Sub TxtCrono_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        
        SendKeys "{TAB}"
        
    ElseIf KeyAscii = 46 Then
        
        If InStr(1, TxtCrono.Text, ".") <> 0 Then
            
            KeyAscii = 0
            
        End If
        
    ElseIf (KeyAscii > 47 And KeyAscii < 58) Or (KeyAscii = 8) Then
        
        'Tecla permitida
        
    Else
        
        KeyAscii = 0
        
    End If
    
End Sub

Private Sub TxtCrono_LostFocus()
    On Error Resume Next
    TxtROP = Format(60 / CDbl(TxtCrono), "0.00")
    
End Sub


Private Sub txtFecha_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        
        SendKeys "{TAB}"
        
    End If
End Sub

Private Sub txtHora_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub


Private Sub txtHora_LostFocus()
    
    If Txthora.Text <> "" Then
        
        If Not IsDate(Txthora.Text) Then
            
            frmMsg.MostrarMsg "Debe ingresar una Hora.", "Error", MdiPrincipal
            
            Txthora.SetFocus
            
        Else
            
            Txthora.Text = Format(Txthora.Text, "HH:MM:ss")
            
        End If
        
    End If

    If (Txtfecha.Value <> 0) And (Txthora.Text <> "") Then
        cmdConsultar.Enabled = True
    
        Else
            'CmdConsultar.Enabled = False
    
    End If
End Sub











Private Sub TxtGasTotal_GotFocus()
    
    TextSelected
    
End Sub

Private Sub TxtGasTotal_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        
        SendKeys "{TAB}"
        
'    ElseIf KeyAscii = 46 Then
'
'        If InStr(1, TxtGasTotal.Text, ".") <> 0 Then
'
'            KeyAscii = 0
'
'        End If
        
    ElseIf (KeyAscii > 47 And KeyAscii < 58) Or (KeyAscii = 8) Then
        
        'Tecla permitida
        
    Else
        
        KeyAscii = 0
        
    End If
    
End Sub

Private Sub TxtGasTotalDefinitivo_GotFocus()
    
    TextSelected
    
End Sub

Private Sub TxtGasTotalDefinitivo_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        
        SendKeys "{TAB}"
        
    ElseIf (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8 Then
        
        'Tecla permitida
        
    Else
        
        KeyAscii = 0
        
    End If
    
End Sub

Private Sub TxtGasTotalTemporal_GotFocus()
    
    TextSelected
    
End Sub

Private Sub TxtGasTotalTemporal_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        
        SendKeys "{TAB}"
        
    ElseIf (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8 Then
        
        'Tecla permitida
        
    Else
        
        KeyAscii = 0
        
    End If
    
End Sub

Private Sub TxtProfundidadDefinitivo_GotFocus()
    
    TextSelected
    
End Sub

Private Sub TxtProfundidadDefinitivo_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        
        SendKeys "{TAB}"
        
    ElseIf (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8 Then
        
        'Tecla permitida
        
    Else
        
        KeyAscii = 0
        
    End If
    
End Sub



Private Sub TxtProfundidadInicio_GotFocus()
    TextSelected
End Sub

Private Sub TxtProfundidadInicio_KeyPress(KeyAscii As Integer)
    SoloNumeros KeyAscii
End Sub

Private Sub TxtProfundidadInicio_LostFocus()
    If TxtProfundidadInicio.Text = 0 Then
        TxtProfundidadInicio.Text = 1
    End If
End Sub

Private Sub TxtProfundidadTemporal_GotFocus()
    
    TextSelected
    
End Sub

Private Sub TxtProfundidadTemporal_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        
        SendKeys "{TAB}"
        
    ElseIf (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8 Then
        
        'Tecla permitida
        
    Else
        
        KeyAscii = 0
        
    End If
    
End Sub

Private Sub TxtROP_GotFocus()
    
    TextSelected
    
End Sub

Private Sub TxtROP_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        
        SendKeys "{TAB}"
        
    ElseIf KeyAscii = 46 Then
        
        If InStr(1, TxtROP.Text, ".") <> 0 Then
            
            KeyAscii = 0
            
        End If
        
    ElseIf (KeyAscii > 47 And KeyAscii < 58) Or (KeyAscii = 8) Then
        
        'Tecla permitida
        
    Else
        
        KeyAscii = 0
        
    End If
    
End Sub

Private Sub TxtROP_LostFocus()
On Error Resume Next
    TxtCrono = Format(60 / CDbl(TxtROP), "0.00")
    
End Sub

Private Sub TxtSH2Definitivo_GotFocus()
    
    TextSelected
    
End Sub

Private Sub TxtSH2Definitivo_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        
        SendKeys "{TAB}"
        
    ElseIf (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8 Then
        
        'Tecla permitida
        
    Else
        
        KeyAscii = 0
        
    End If
    
End Sub

Private Sub TxtSH2Temporal_GotFocus()
    
    TextSelected
    
End Sub

Private Sub TxtSH2Temporal_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        
        SendKeys "{TAB}"
        
    ElseIf (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8 Then
        
        'Tecla permitida
        
    Else
        
        KeyAscii = 0
        
    End If
    
End Sub

Public Sub AnalisisGasDeYacimiento()
    
    Dim NumeroAnalisis                  As Long
    Dim objPozoAnalis                   As clsPozoAnalis
    
    Dim StrSql As String
    
    NumeroAnalisis = Val(LvwAnalisisTemporal.SelectedItem.Key)
    Set objPozoAnalis = gObjPozoActivo.PozoAnalisis.ColPozoAnalis(gObjPozoActivo.IdPozo, NumeroAnalisis)
    If Not objPozoAnalis Is Nothing Then
        If LvwAnalisisTemporal.SelectedItem.SmallIcon = 3 Then
            LvwAnalisisTemporal.SelectedItem.SmallIcon = 1
            objPozoAnalis.CodigoTipoDeAnalisis = 1
            gObjPozoActivo.PozoAnalisis.dbModificar objPozoAnalis
            TxtProfundidadTemporal.Enabled = True
        Else
            LvwAnalisisTemporal.SelectedItem.SmallIcon = 3
            objPozoAnalis.CodigoTipoDeAnalisis = 3
            objPozoAnalis.Profundidad = 0
            gObjPozoActivo.PozoAnalisis.dbModificar objPozoAnalis
            LvwAnalisisTemporal.SelectedItem.SubItems(3) = 0
            TxtProfundidadTemporal.Text = 0
            TxtProfundidadTemporal.Enabled = False
        End If
        If ChkMostrarTodosLosAnálisisTemporal.Value = 0 Then
            LvwAnalisisTemporal.ListItems.Remove LvwAnalisisTemporal.SelectedItem.Index
            LvwAnalisisTemporal.SetFocus
            If LvwAnalisisTemporal.ListItems.Count <> 0 Then
                LvwAnalisisTemporal.SelectedItem = LvwAnalisisTemporal.ListItems(1)
                Set objPozoAnalis = gObjPozoActivo.PozoAnalisis.ColPozoAnalis(gObjPozoActivo.IdPozo, Val(LvwAnalisisTemporal.SelectedItem.Key))
                If Not objPozoAnalis Is Nothing Then
                    BuscarComponentes objPozoAnalis, "TEMPORAL"
                    BuscarDatosAnalisisTemporales objPozoAnalis
                    BuscarRelacionesCromatograficasTemporal objPozoAnalis
                End If
            Else
                LimpiatDatosTemporales
            End If
        End If
    End If
End Sub

Public Sub ComponenteTemporalAgregar()
    
    Dim objPozoAnalis                   As clsPozoAnalis
    Dim ok                              As Boolean
    
    Set objPozoAnalis = gObjPozoActivo.PozoAnalisis.ColPozoAnalis(gObjPozoActivo.IdPozo, Val(LvwAnalisisTemporal.SelectedItem.Key))
    If Not objPozoAnalis Is Nothing Then
        Set FrmDatosComponente.gObjPozoAnalis = objPozoAnalis
        FrmDatosComponente.Accion = euAccionAgregar
        FrmDatosComponente.Show vbModal
    
        If FormularioCargado("FrmDatosComponente") Then
            Unload FrmDatosComponente
            gDatos.BeginTrans
            CalcularRelacionesCromatograficas objPozoAnalis, "MODIFICAR", ok
            If ok Then
                ActualizarGasTotalCromatografico objPozoAnalis, ok
                If ok Then
                    BuscarRelacionesCromatograficasTemporal objPozoAnalis
                    BuscarDatosAnalisisTemporales objPozoAnalis
                    BuscarComponentes objPozoAnalis, "TEMPORAL"
                    LvwTemporal.SetFocus
                    gDatos.CommitTrans
                Else
                    frmMsg.MostrarMsg "No se pudo actualizar el gas total cromatográfico", "Error", MdiPrincipal
                    gDatos.RollBackTrans
                End If
            Else
                frmMsg.MostrarMsg "No se pudieron actualizar las relaciones cromatográficas", "Error", MdiPrincipal
                gDatos.RollBackTrans
            End If
        Else
            LvwTemporal.SetFocus
        End If
    End If
    
End Sub

Public Sub AnalisisCirculada()
    
    Dim NumeroAnalisis              As Long
    Dim objPozoAnalis               As clsPozoAnalis
    
    NumeroAnalisis = Val(LvwAnalisisTemporal.SelectedItem.Key)
    Set objPozoAnalis = gObjPozoActivo.PozoAnalisis.ColPozoAnalis(gObjPozoActivo.IdPozo, NumeroAnalisis)
    If Not objPozoAnalis Is Nothing Then
        If LvwAnalisisTemporal.SelectedItem.SmallIcon = 5 Then
            LvwAnalisisTemporal.SelectedItem.SmallIcon = 1
            objPozoAnalis.CodigoTipoDeAnalisis = 1
            gObjPozoActivo.PozoAnalisis.dbModificar objPozoAnalis
            TxtProfundidadTemporal.Enabled = True
        Else
            
            LvwAnalisisTemporal.SelectedItem.SmallIcon = 5
            objPozoAnalis.CodigoTipoDeAnalisis = 5
            objPozoAnalis.Profundidad = 0
            gObjPozoActivo.PozoAnalisis.dbModificar objPozoAnalis
            LvwAnalisisTemporal.SelectedItem.SubItems(3) = 0
            TxtProfundidadTemporal.Text = 0
            TxtProfundidadTemporal.Enabled = False
            
        End If
        
        If ChkMostrarTodosLosAnálisisTemporal.Value = 0 Then
            LvwAnalisisTemporal.ListItems.Remove LvwAnalisisTemporal.SelectedItem.Index
            LvwAnalisisTemporal.SetFocus
            If LvwAnalisisTemporal.ListItems.Count <> 0 Then
                LvwAnalisisTemporal.SelectedItem = LvwAnalisisTemporal.ListItems(1)
                Set objPozoAnalis = gObjPozoActivo.PozoAnalisis.ColPozoAnalis(gObjPozoActivo.IdPozo, Val(LvwAnalisisTemporal.SelectedItem.Key))
                If Not objPozoAnalis Is Nothing Then
                    BuscarComponentes objPozoAnalis, "TEMPORAL"
                    BuscarDatosAnalisisTemporales objPozoAnalis
                    BuscarRelacionesCromatograficasTemporal objPozoAnalis
                End If
            Else
                LimpiatDatosTemporales
            End If
        End If
    End If
End Sub

Public Sub ActualizoDisplayEscalas()
    

    Select Case EContinuoControlador
            
            Case 1: iLedTha.ActiveColor = 57088   ' Verde
               '     SwTha.Position = 1
            
            Case 10: iLedTha.ActiveColor = 65535 ' Amarillo
              '      SwTha.Position = 2
            
            Case 100: iLedTha.ActiveColor = 255   ' Rojo
             '       SwTha.Position = 3
    End Select

    
    

     Select Case ENormalControlador

            Case 1: iLedNormal.ActiveColor = 57088 ' Verde
            'SwNormal.Position = 1
            
            Case 10: iLedNormal.ActiveColor = 65535 ' Amarillo
            'SwNormal.Position = 2
            
            Case 100: iLedNormal.ActiveColor = 255   ' Rojo
            'SwNormal.Position = 3
    End Select


    

    Select Case EFastControlador

            Case 1: iLedFast.ActiveColor = 57088 ' Verde
                   ' SwFast.Position = 1
            
            Case 10: iLedFast.ActiveColor = 65535 ' Amarillo
                   ' SwFast.Position = 2
            
            Case 100: iLedFast.ActiveColor = 255   ' Rojo
                   ' SwFast.Position = 3
    End Select


    Select Case EContinuoLocal
    
        Case 1: SwTha.Position = 0
                        
        Case 10: SwTha.Position = 1
        
        Case 100: SwTha.Position = 2
        
    End Select

    Select Case ENormalLocal
    
        Case 1: SwNormal.Position = 0
                        
        Case 10: SwNormal.Position = 1
        
        Case 100: SwNormal.Position = 2
        
    End Select

    Select Case EFastLocal
    
        Case 1: SwFast.Position = 0
                        
        Case 10: SwFast.Position = 1
        
        Case 100: SwFast.Position = 2
        
    End Select


End Sub


Public Function CargarMaximaProfundidad() As Double
    Set gObjPozoActivo.PozoEscaneos.Datos = gDatos
    CargarMaximaProfundidad = gObjPozoActivo.PozoEscaneos.dbMaxProfPozoEscaneo(gObjPozoActivo.IdPozo)
End Function
    

Public Sub SetearRangoGraficacion()
    
    TxtSpan = GetSetting("Iga", "Config", "SpanProfundidad")
    If Not IsNumeric(TxtSpan) Then
        TxtSpan = 50
    End If
    If CargarMaximaProfundidad > CDbl(TxtSpan) Then
        TxtProfundidadInicio = CargarMaximaProfundidad - CDbl(TxtSpan)
    Else
        TxtProfundidadInicio = 0
    End If
End Sub



Public Sub Interpolar()

    Dim objPozoEscaneo                      As clsPozoEscaneo
    Dim ProfundidadReferencia               As Double
    Dim IdPozoEscaneo                       As Long
    Dim ProfundidadAnterior                 As Double
    Dim UnGasAnterior                       As Double
    Dim ProfundidadPosterior                As Double
    Dim UnGasPosterior                      As Double
    Dim UnNuevoValorDeGas                   As Double

    On Error GoTo Error
    Me.MousePointer = 13

    For Each objPozoEscaneo In gObjPozoActivo.PozoEscaneos
        ProfundidadReferencia = objPozoEscaneo.ProfundidadPozo
        IdPozoEscaneo = objPozoEscaneo.IdPozoEscaneo
        If ObtenerAnterior(ProfundidadReferencia, ProfundidadAnterior, UnGasAnterior) _
            And ObtenerPosterior(ProfundidadReferencia, ProfundidadPosterior, UnGasPosterior) Then
            UnNuevoValorDeGas = UnGasAnterior + _
                ((UnGasPosterior - UnGasAnterior) / (ProfundidadPosterior - ProfundidadAnterior)) * _
                    (ProfundidadReferencia - ProfundidadAnterior)
            objPozoEscaneo.GasTotal = UnNuevoValorDeGas
            Set gObjPozoActivo.PozoEscaneos.Datos = gDatos
            gObjPozoActivo.PozoEscaneos.dbModificar objPozoEscaneo
            
        End If
            
    Next
    Me.MousePointer = 0

Error:
    If Err.Number <> 0 Then
        frmMsg.MostrarMsg "Módulo: ConsultarTodos los cronos ." & Chr(10) & "Ocurrió el error: " & Err.Number & " - " & Err.Description, "Error", MdiPrincipal
        Err.Clear
    End If
End Sub

Private Function ObtenerAnterior(UnaProfundidadReferencia As Double, _
                                 UnaProfundidad As Double, _
                                 UnGas As Double) As Boolean
    Dim objPozoEscaneo              As clsPozoEscaneo
    Dim objPozoEscaneos             As clsPozoEscaneos
    
On Error GoTo Error
    
    Set gObjPozoActivo.PozoEscaneos.Datos = gDatos
    objPozoEscaneos.ConsultarPozoEscaneosPorProfundidadAnteriores (UnaProfundidadReferencia)
    
    ObtenerAnterior = False
    For Each objPozoEscaneo In objPozoEscaneos
        ObtenerAnterior = True
        UnaProfundidad = objPozoEscaneo.ProfundidadPozo
        UnGas = objPozoEscaneo.GasTotal
        Exit For
    Next

Error:
    If Err.Number <> 0 Then
        frmMsg.MostrarMsg "Módulo: ConsultarTodos los cronos ." & Chr(10) & "Ocurrió el error: " & Err.Number & " - " & Err.Description, "Error", MdiPrincipal
        Err.Clear
    End If
End Function

Private Function ObtenerPosterior(UnaProfundidadReferencia As Double, _
                                 UnaProfundidad As Double, _
                                 UnGas As Double) As Boolean

    Dim objPozoEscaneo              As clsPozoEscaneo
    Dim objPozoEscaneos             As clsPozoEscaneos
    
On Error GoTo Error
    
    Set gObjPozoActivo.PozoEscaneos.Datos = gDatos
    objPozoEscaneos.ConsultarPozoEscaneosPorProfundidadPosteriores (UnaProfundidadReferencia)
    
    ObtenerPosterior = False
    For Each objPozoEscaneo In objPozoEscaneos
        ObtenerPosterior = True
        UnaProfundidad = objPozoEscaneo.ProfundidadPozo
        UnGas = objPozoEscaneo.GasTotal
        Exit For
    Next
Error:
    If Err.Number <> 0 Then
        frmMsg.MostrarMsg "Módulo: ConsultarTodos los cronos ." & Chr(10) & "Ocurrió el error: " & Err.Number & " - " & Err.Description, "Error", MdiPrincipal
        Err.Clear
    End If

End Function

Public Function ConsultarPozoEscaneos(ByVal a_partir_de As Double) As Boolean
    
    Dim objPozoEscaneo                  As clsPozoEscaneo
    Dim Item                            As ListItem
    Dim Indice                          As Long
    
    On Error GoTo Error
    
    ConsultarPozoEscaneos = True
    If a_partir_de = 0 Then
        LvwCronoGas.ListItems.Clear
        gObjPozoActivo.PozoEscaneos.Clear
        Set gObjPozoActivo.PozoEscaneos.Datos = gDatos
        gObjPozoActivo.PozoEscaneos.ConsultarPozoEscaneos
    End If
    Indice = gObjPozoActivo.PozoEscaneos.Count
    
    Do Until Indice < 1
        Set objPozoEscaneo = gObjPozoActivo.PozoEscaneos.ColPozoEscaneoPorIndice(Indice)
        If Not objPozoEscaneo Is Nothing Then
            If objPozoEscaneo.ProfundidadPozo > a_partir_de Then
                Set Item = LvwCronoGas.ListItems.Add(1)
                Item.SmallIcon = 1
                
                Item.Text = objPozoEscaneo.ProfundidadPozo
                Item.SubItems(1) = Format(objPozoEscaneo.Crono, "0.00")
                Item.SubItems(2) = Format(objPozoEscaneo.ROP, "0.00")
                Item.SubItems(3) = Format(objPozoEscaneo.GasTotal, "0")
                Item.SubItems(4) = Format(objPozoEscaneo.CO2, "0")
                Item.SubItems(5) = Format(objPozoEscaneo.ProfundidadPozo, "00000")
                
                Item.Key = objPozoEscaneo.Key
            End If
            
        End If
        Indice = Indice - 1
    Loop
    
    If LvwCronoGas.ListItems.Count <> 0 Then
        LvwCronoGas.ListItems(1).Selected = True
        TxtCrono.Text = LvwCronoGas.SelectedItem.SubItems(1)
        TxtROP.Text = LvwCronoGas.SelectedItem.SubItems(2)
        TxtGasTotal.Text = LvwCronoGas.SelectedItem.SubItems(3)
        TxtCO2.Text = LvwCronoGas.SelectedItem.SubItems(4)
        LvwCronoGas.ListItems(LvwCronoGas.SelectedItem.Index).EnsureVisible
    End If
    
Error:
    If Err.Number <> 0 Then
        ConsultarPozoEscaneos = False
        frmMsg.MostrarMsg "Módulo: ConsultarPozoEscaneos ." & Chr(10) & "Ocurrió el error: " & Err.Number & " - " & Err.Description, "Error", MdiPrincipal
        Err.Clear
    End If
End Function

Public Sub ActualizarGasesEscaneos(ByVal Desde As Double, ByVal Hasta As Double)
    
    Dim objPozoEscaneo                  As clsPozoEscaneo
    Dim Item                            As ListItem
    
    On Error GoTo Error
    
    For Each Item In LvwCronoGas.ListItems
        Set objPozoEscaneo = gObjPozoActivo.PozoEscaneos.ColPozoEscaneo(, , Item.Key)
        If Not objPozoEscaneo Is Nothing Then
            If objPozoEscaneo.ProfundidadPozo >= Desde And objPozoEscaneo.ProfundidadPozo <= Hasta Then
                Item.SubItems(3) = Format(objPozoEscaneo.GasTotal, "0")
                Item.SubItems(4) = Format(objPozoEscaneo.CO2, "0")
            End If
        End If
    Next

Error:
    If Err.Number <> 0 Then
        frmMsg.MostrarMsg "Módulo: ActualizarGasesEscaneos." & Chr(10) & "Ocurrió el error: " & Err.Number & " - " & Err.Description, "Error", MdiPrincipal
        Err.Clear
    End If
End Sub

Sub ScrollLeft()
    Dim fecha As Date
    fecha = CDate(Txtfecha & " " & Txthora.Text)
    fecha = DateAdd("n", -1 * gSpan, fecha)
    Txtfecha = CDate(Format(fecha, "dd/mm/yyyy"))
    Txthora.Text = Format(fecha, "hh:nn")
    CmdConsultar_Click
End Sub

Sub ScrollRight()
    Dim fecha As Date
    fecha = CDate(Txtfecha & " " & Txthora.Text)
    fecha = DateAdd("n", gSpan, fecha)
    Txtfecha = CDate(Format(fecha, "dd/mm/yyyy"))
    Txthora.Text = Format(fecha, "hh:nn")
    CmdConsultar_Click
End Sub

Sub ScrollRightProf()
    
    Dim Profu As Double
    If IsNumeric(TxtProfundidadInicio.Text) And IsNumeric(TxtSpan.Text) Then
        Profu = CDbl(TxtProfundidadInicio.Text)
        Profu = Profu + CDbl(TxtSpan.Text)
        TxtProfundidadInicio.Text = Profu
        cmdActualizar_Click
    End If
    
End Sub

Sub ScrollLeftProf()
    Dim Profu As Double
    If IsNumeric(TxtProfundidadInicio.Text) And IsNumeric(TxtSpan.Text) Then
        Profu = CDbl(TxtProfundidadInicio.Text)
        Profu = Profu - CDbl(TxtSpan.Text)
        If Profu < 0 Then Profu = 0
        TxtProfundidadInicio.Text = Profu
        cmdActualizar_Click
    End If
End Sub

Public Sub ReconsultarDatosTiempo()
    CmdConsultar_Click
End Sub

Private Sub TxtSpan_GotFocus()
    TextSelected
End Sub

Private Sub TxtSpan_KeyPress(KeyAscii As Integer)

    SoloNumeros KeyAscii

End Sub



Private Sub TxtSpan_LostFocus()

    If TxtSpan.Text = 0 Then
        TxtSpan.Text = 1
    End If
    
End Sub
