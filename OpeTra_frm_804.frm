VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_Car_PrePag_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   9150
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13320
   Icon            =   "OpeTra_frm_804.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9150
   ScaleWidth      =   13320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel SSPanel13 
      Height          =   9165
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   13300
      _Version        =   65536
      _ExtentX        =   23460
      _ExtentY        =   16166
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSPanel SSPanel1 
         Height          =   615
         Left            =   30
         TabIndex        =   11
         Top             =   30
         Width           =   13245
         _Version        =   65536
         _ExtentX        =   23363
         _ExtentY        =   1085
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Begin MSComDlg.CommonDialog dlg_Guarda 
            Left            =   10980
            Top             =   60
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   555
            Left            =   690
            TabIndex        =   12
            Top             =   30
            Width           =   4755
            _Version        =   65536
            _ExtentX        =   8387
            _ExtentY        =   979
            _StockProps     =   15
            Caption         =   "Carga de Cronogramas"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            Font3D          =   2
            Alignment       =   1
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   90
            Picture         =   "OpeTra_frm_804.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   30
         TabIndex        =   13
         Top             =   660
         Width           =   13245
         _Version        =   65536
         _ExtentX        =   23363
         _ExtentY        =   1138
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   660
            Picture         =   "OpeTra_frm_804.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   12630
            Picture         =   "OpeTra_frm_804.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ImpCro 
            Height          =   585
            Left            =   60
            Picture         =   "OpeTra_frm_804.frx":0B9A
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Consulta Cronograma de Pagos"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   2445
         Left            =   30
         TabIndex        =   14
         Top             =   1320
         Width           =   13245
         _Version        =   65536
         _ExtentX        =   23363
         _ExtentY        =   4313
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Begin TabDlg.SSTab SSTab1 
            Height          =   2220
            Index           =   1
            Left            =   120
            TabIndex        =   9
            Top             =   120
            Width           =   13095
            _ExtentX        =   23098
            _ExtentY        =   3916
            _Version        =   393216
            Tabs            =   2
            TabsPerRow      =   2
            TabHeight       =   520
            TabCaption(0)   =   "Datos del Crédito"
            TabPicture(0)   =   "OpeTra_frm_804.frx":0EA4
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "grd_Listad"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Datos del Prepago"
            TabPicture(1)   =   "OpeTra_frm_804.frx":0EC0
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "pnl_Tit_MtoTC"
            Tab(1).Control(1)=   "pnl_Tit_MtoTNC"
            Tab(1).Control(2)=   "pnl_Tit_MtoDep"
            Tab(1).Control(3)=   "pnl_Tit_FecEnv"
            Tab(1).Control(4)=   "pnl_Tit_MtoApli"
            Tab(1).Control(5)=   "pnl_Tit_FecPpg"
            Tab(1).Control(6)=   "pnl_Tit_TipPpg"
            Tab(1).Control(7)=   "pnl_Tit_ChkSel"
            Tab(1).Control(8)=   "grd_Listap"
            Tab(1).ControlCount=   9
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   1575
               Left            =   180
               TabIndex        =   113
               Top             =   480
               Width           =   12630
               _ExtentX        =   22278
               _ExtentY        =   2778
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
               Appearance      =   0
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listap 
               Height          =   1275
               Left            =   -74880
               TabIndex        =   114
               Top             =   780
               Width           =   12900
               _ExtentX        =   22754
               _ExtentY        =   2249
               _Version        =   393216
               Rows            =   21
               Cols            =   8
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
               Appearance      =   0
            End
            Begin Threed.SSPanel pnl_Tit_ChkSel 
               Height          =   285
               Left            =   -63720
               TabIndex        =   115
               Top             =   480
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "  Seleccionar"
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
               Alignment       =   1
               Begin VB.CheckBox chkSeleccionar 
                  BackColor       =   &H00004000&
                  Height          =   255
                  Left            =   1080
                  TabIndex        =   116
                  Top             =   0
                  Width           =   255
               End
            End
            Begin Threed.SSPanel pnl_Tit_TipPpg 
               Height          =   285
               Left            =   -74880
               TabIndex        =   117
               Top             =   480
               Width           =   2085
               _Version        =   65536
               _ExtentX        =   3678
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Tipo de Prepago"
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel pnl_Tit_FecPpg 
               Height          =   285
               Left            =   -72795
               TabIndex        =   118
               Top             =   480
               Width           =   1380
               _Version        =   65536
               _ExtentX        =   2434
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "F. Prepago"
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel pnl_Tit_MtoApli 
               Height          =   285
               Left            =   -69840
               TabIndex        =   119
               Top             =   480
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Mto. Aplicar"
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel pnl_Tit_FecEnv 
               Height          =   285
               Left            =   -65160
               TabIndex        =   120
               Top             =   480
               Width           =   1500
               _Version        =   65536
               _ExtentX        =   2646
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "F. Envío COFIDE"
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel pnl_Tit_MtoDep 
               Height          =   285
               Left            =   -71415
               TabIndex        =   121
               Top             =   480
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Mto. Depositado"
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel pnl_Tit_MtoTNC 
               Height          =   285
               Left            =   -68280
               TabIndex        =   122
               Top             =   480
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Mto. Aplicar TNC"
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel pnl_Tit_MtoTC 
               Height          =   285
               Left            =   -66720
               TabIndex        =   123
               Top             =   480
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Mto. Aplicar TC"
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   915
         Left            =   30
         TabIndex        =   15
         Top             =   3780
         Width           =   13245
         _Version        =   65536
         _ExtentX        =   23363
         _ExtentY        =   1614
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Begin VB.CommandButton cmd_Generar 
            Height          =   585
            Left            =   12390
            Picture         =   "OpeTra_frm_804.frx":0EDC
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Carga TC Cliente"
            Top             =   240
            Width           =   585
         End
         Begin VB.TextBox txt_NomArc 
            Height          =   315
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   1
            Text            =   "txt_NomArc"
            Top             =   480
            Width           =   9105
         End
         Begin VB.CommandButton cmd_BuscaArc 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   11190
            TabIndex        =   2
            ToolTipText     =   "Seleccionar archivo"
            Top             =   480
            Width           =   315
         End
         Begin VB.ComboBox cmb_TipCro 
            Height          =   315
            Left            =   1800
            TabIndex        =   0
            Text            =   "cmb_TipCro"
            Top             =   150
            Width           =   3165
         End
         Begin VB.CommandButton cmd_Import 
            Height          =   585
            Left            =   11790
            Picture         =   "OpeTra_frm_804.frx":11E6
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Importar archivo"
            Top             =   240
            Width           =   585
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Cronograma"
            Height          =   195
            Left            =   150
            TabIndex        =   100
            Top             =   210
            Width           =   1440
         End
         Begin VB.Label Label4 
            Caption         =   "Archivo a cargar:"
            Height          =   255
            Left            =   180
            TabIndex        =   16
            Top             =   510
            Width           =   1365
         End
      End
      Begin Threed.SSPanel SSPanel22 
         Height          =   4395
         Left            =   30
         TabIndex        =   17
         Top             =   4710
         Width           =   13245
         _Version        =   65536
         _ExtentX        =   23363
         _ExtentY        =   7752
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Begin TabDlg.SSTab tab_Cronog 
            Height          =   4245
            Left            =   60
            TabIndex        =   5
            Top             =   90
            Width           =   13110
            _ExtentX        =   23125
            _ExtentY        =   7488
            _Version        =   393216
            Style           =   1
            Tabs            =   4
            Tab             =   3
            TabsPerRow      =   4
            TabHeight       =   520
            TabCaption(0)   =   "FMV - Tramo No Concesional"
            TabPicture(0)   =   "OpeTra_frm_804.frx":1628
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "SSPanel14"
            Tab(0).Control(1)=   "SSPanel5"
            Tab(0).Control(2)=   "SSPanel12"
            Tab(0).Control(3)=   "SSPanel9"
            Tab(0).Control(4)=   "SSPanel6"
            Tab(0).Control(5)=   "grd_MViNCo_Listad"
            Tab(0).Control(6)=   "SSPanel10"
            Tab(0).Control(7)=   "SSPanel16"
            Tab(0).Control(8)=   "SSPanel11"
            Tab(0).ControlCount=   9
            TabCaption(1)   =   "FMV - Tramo Concesional"
            TabPicture(1)   =   "OpeTra_frm_804.frx":1644
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SSPanel15"
            Tab(1).Control(1)=   "SSPanel18"
            Tab(1).Control(2)=   "SSPanel41"
            Tab(1).Control(3)=   "SSPanel39"
            Tab(1).Control(4)=   "SSPanel38"
            Tab(1).Control(5)=   "SSPanel21"
            Tab(1).Control(6)=   "SSPanel20"
            Tab(1).Control(7)=   "SSPanel19"
            Tab(1).Control(8)=   "grd_MViCon_Listad"
            Tab(1).ControlCount=   9
            TabCaption(2)   =   "Cliente - Tramo Concesional"
            TabPicture(2)   =   "OpeTra_frm_804.frx":1660
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "grd_CliCon_Listad"
            Tab(2).Control(1)=   "SSPanel17"
            Tab(2).Control(2)=   "SSPanel42"
            Tab(2).Control(3)=   "SSPanel70"
            Tab(2).Control(4)=   "SSPanel63"
            Tab(2).Control(5)=   "SSPanel60"
            Tab(2).Control(6)=   "SSPanel58"
            Tab(2).Control(7)=   "SSPanel43"
            Tab(2).ControlCount=   8
            TabCaption(3)   =   "Cliente - Tramo No Concesional"
            TabPicture(3)   =   "OpeTra_frm_804.frx":167C
            Tab(3).ControlEnabled=   -1  'True
            Tab(3).Control(0)=   "SSPanel80"
            Tab(3).Control(0).Enabled=   0   'False
            Tab(3).Control(1)=   "SSPanel79"
            Tab(3).Control(1).Enabled=   0   'False
            Tab(3).Control(2)=   "SSPanel78"
            Tab(3).Control(2).Enabled=   0   'False
            Tab(3).Control(3)=   "grd_CliNCon_Listad"
            Tab(3).Control(3).Enabled=   0   'False
            Tab(3).Control(4)=   "SSPanel77"
            Tab(3).Control(4).Enabled=   0   'False
            Tab(3).Control(5)=   "SSPanel76"
            Tab(3).Control(5).Enabled=   0   'False
            Tab(3).Control(6)=   "SSPanel75"
            Tab(3).Control(6).Enabled=   0   'False
            Tab(3).Control(7)=   "SSPanel74"
            Tab(3).Control(7).Enabled=   0   'False
            Tab(3).Control(8)=   "SSPanel73"
            Tab(3).Control(8).Enabled=   0   'False
            Tab(3).Control(9)=   "SSPanel72"
            Tab(3).Control(9).Enabled=   0   'False
            Tab(3).Control(10)=   "SSPanel71"
            Tab(3).Control(10).Enabled=   0   'False
            Tab(3).ControlCount=   11
            Begin Threed.SSPanel SSPanel11 
               Height          =   285
               Left            =   -68310
               TabIndex        =   18
               Top             =   390
               Width           =   1500
               _Version        =   65536
               _ExtentX        =   2646
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Total Cuota"
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel16 
               Height          =   285
               Left            =   -69810
               TabIndex        =   19
               Top             =   390
               Width           =   1500
               _Version        =   65536
               _ExtentX        =   2646
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Comisión COFIDE"
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel10 
               Height          =   285
               Left            =   -72810
               TabIndex        =   20
               Top             =   390
               Width           =   1500
               _Version        =   65536
               _ExtentX        =   2646
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Capital"
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel23 
               Height          =   285
               Left            =   -67530
               TabIndex        =   21
               Top             =   360
               Width           =   2370
               _Version        =   65536
               _ExtentX        =   4180
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Total Cuota"
               ForeColor       =   16777215
               BackColor       =   32768
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel25 
               Height          =   285
               Left            =   -65190
               TabIndex        =   22
               Top             =   360
               Width           =   2370
               _Version        =   65536
               _ExtentX        =   4180
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Saldo Capital"
               ForeColor       =   16777215
               BackColor       =   32768
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel26 
               Height          =   285
               Left            =   -74940
               TabIndex        =   23
               Top             =   360
               Width           =   1185
               _Version        =   65536
               _ExtentX        =   2090
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Cuota"
               ForeColor       =   16777215
               BackColor       =   32768
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel27 
               Height          =   285
               Left            =   -73770
               TabIndex        =   24
               Top             =   360
               Width           =   1815
               _Version        =   65536
               _ExtentX        =   3201
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "F. Vcto"
               ForeColor       =   16777215
               BackColor       =   32768
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel28 
               Height          =   285
               Left            =   -71970
               TabIndex        =   25
               Top             =   360
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Capital"
               ForeColor       =   16777215
               BackColor       =   32768
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel29 
               Height          =   285
               Left            =   -70140
               TabIndex        =   26
               Top             =   360
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Interés"
               ForeColor       =   16777215
               BackColor       =   32768
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel31 
               Height          =   285
               Left            =   -66480
               TabIndex        =   27
               Top             =   360
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Total Cuota"
               ForeColor       =   16777215
               BackColor       =   32768
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel32 
               Height          =   285
               Left            =   -64650
               TabIndex        =   28
               Top             =   360
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Saldo Capital"
               ForeColor       =   16777215
               BackColor       =   32768
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel37 
               Height          =   285
               Left            =   -68310
               TabIndex        =   29
               Top             =   360
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Comisión"
               ForeColor       =   16777215
               BackColor       =   32768
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel40 
               Height          =   285
               Left            =   -74940
               TabIndex        =   30
               Top             =   360
               Width           =   1185
               _Version        =   65536
               _ExtentX        =   2090
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Cuota"
               ForeColor       =   16777215
               BackColor       =   32768
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel48 
               Height          =   285
               Left            =   -73770
               TabIndex        =   31
               Top             =   360
               Width           =   1815
               _Version        =   65536
               _ExtentX        =   3201
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "F. Vcto"
               ForeColor       =   16777215
               BackColor       =   32768
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel50 
               Height          =   285
               Left            =   -71970
               TabIndex        =   32
               Top             =   360
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Capital"
               ForeColor       =   16777215
               BackColor       =   32768
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel51 
               Height          =   285
               Left            =   -70140
               TabIndex        =   33
               Top             =   360
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Interés"
               ForeColor       =   16777215
               BackColor       =   32768
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel52 
               Height          =   285
               Left            =   -66480
               TabIndex        =   34
               Top             =   360
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Total Cuota"
               ForeColor       =   16777215
               BackColor       =   32768
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel53 
               Height          =   285
               Left            =   -64650
               TabIndex        =   35
               Top             =   360
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Saldo Capital"
               ForeColor       =   16777215
               BackColor       =   32768
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel57 
               Height          =   285
               Left            =   -68310
               TabIndex        =   36
               Top             =   360
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Comisión"
               ForeColor       =   16777215
               BackColor       =   32768
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel30 
               Height          =   285
               Left            =   -71760
               TabIndex        =   37
               Top             =   390
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Interés"
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel8 
               Height          =   285
               Left            =   -74940
               TabIndex        =   38
               Top             =   390
               Width           =   705
               _Version        =   65536
               _ExtentX        =   1244
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Cuota"
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel33 
               Height          =   285
               Left            =   -74250
               TabIndex        =   39
               Top             =   390
               Width           =   1425
               _Version        =   65536
               _ExtentX        =   2514
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "F. Vcto"
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel34 
               Height          =   285
               Left            =   -72840
               TabIndex        =   40
               Top             =   390
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Capital"
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel35 
               Height          =   285
               Left            =   -66360
               TabIndex        =   41
               Top             =   390
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Total Cuota"
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel36 
               Height          =   285
               Left            =   -65280
               TabIndex        =   42
               Top             =   390
               Width           =   1290
               _Version        =   65536
               _ExtentX        =   2275
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Saldo Capital"
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel59 
               Height          =   285
               Left            =   -70680
               TabIndex        =   43
               Top             =   390
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Seg. Prest."
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel61 
               Height          =   285
               Left            =   -69600
               TabIndex        =   44
               Top             =   390
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Seg. Vivienda"
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel62 
               Height          =   285
               Left            =   -68520
               TabIndex        =   45
               Top             =   390
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Portes"
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel24 
               Height          =   285
               Left            =   -67440
               TabIndex        =   46
               Top             =   390
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "C. COFIDE"
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel pnl_MViNCo_Comisi 
               Height          =   285
               Left            =   -67470
               TabIndex        =   47
               Top             =   6870
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
               ForeColor       =   16777215
               BackColor       =   192
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_CofNCo_TotCuo 
               Height          =   285
               Left            =   -67470
               TabIndex        =   48
               Top             =   6870
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
               ForeColor       =   16777215
               BackColor       =   192
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
               Alignment       =   4
            End
            Begin Threed.SSPanel SSPanel44 
               Height          =   285
               Left            =   -70950
               TabIndex        =   49
               Top             =   390
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Interés"
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel45 
               Height          =   285
               Left            =   -74940
               TabIndex        =   50
               Top             =   390
               Width           =   765
               _Version        =   65536
               _ExtentX        =   1349
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Cuota"
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel46 
               Height          =   285
               Left            =   -74190
               TabIndex        =   51
               Top             =   390
               Width           =   1515
               _Version        =   65536
               _ExtentX        =   2672
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "F. Vcto"
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel47 
               Height          =   285
               Left            =   -72690
               TabIndex        =   52
               Top             =   390
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Capital"
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel49 
               Height          =   285
               Left            =   -67470
               TabIndex        =   53
               Top             =   390
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Total Cuota"
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel54 
               Height          =   285
               Left            =   -65730
               TabIndex        =   54
               Top             =   390
               Width           =   1710
               _Version        =   65536
               _ExtentX        =   3016
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Saldo Capital"
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel55 
               Height          =   285
               Left            =   -69210
               TabIndex        =   55
               Top             =   390
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Comisión COFIDE"
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel pnl_CofNCo_Intere 
               Height          =   285
               Left            =   -70950
               TabIndex        =   56
               Top             =   6870
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
               ForeColor       =   16777215
               BackColor       =   192
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_CofNCo_Capita 
               Height          =   285
               Left            =   -72690
               TabIndex        =   57
               Top             =   6870
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
               ForeColor       =   16777215
               BackColor       =   192
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_CofNCo_Comisi 
               Height          =   285
               Left            =   -69210
               TabIndex        =   58
               Top             =   6870
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
               ForeColor       =   16777215
               BackColor       =   192
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
               Alignment       =   4
            End
            Begin Threed.SSPanel SSPanel56 
               Height          =   285
               Left            =   -70710
               TabIndex        =   59
               Top             =   390
               Width           =   1590
               _Version        =   65536
               _ExtentX        =   2805
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Interés"
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel64 
               Height          =   285
               Left            =   -74940
               TabIndex        =   60
               Top             =   390
               Width           =   915
               _Version        =   65536
               _ExtentX        =   1614
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Cuota"
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel65 
               Height          =   285
               Left            =   -74040
               TabIndex        =   61
               Top             =   390
               Width           =   1725
               _Version        =   65536
               _ExtentX        =   3043
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Fecha Vencimiento"
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel67 
               Height          =   285
               Left            =   -67710
               TabIndex        =   62
               Top             =   390
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Total Cuota"
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel68 
               Height          =   285
               Left            =   -65850
               TabIndex        =   63
               Top             =   390
               Width           =   1770
               _Version        =   65536
               _ExtentX        =   3122
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Saldo Capital"
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel pnl_Especial_Interes 
               Height          =   285
               Left            =   -70710
               TabIndex        =   64
               Top             =   6870
               Width           =   1590
               _Version        =   65536
               _ExtentX        =   2805
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
               ForeColor       =   16777215
               BackColor       =   192
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Especial_TotalCuota 
               Height          =   285
               Left            =   -67710
               TabIndex        =   65
               Top             =   6870
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
               ForeColor       =   16777215
               BackColor       =   192
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
               Alignment       =   4
            End
            Begin Threed.SSPanel SSPanel69 
               Height          =   285
               Left            =   -69120
               TabIndex        =   66
               Top             =   390
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Seguros"
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel66 
               Height          =   285
               Left            =   -72330
               TabIndex        =   67
               Top             =   390
               Width           =   1620
               _Version        =   65536
               _ExtentX        =   2857
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Capital"
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel pnl_Especial_Capital 
               Height          =   285
               Left            =   -72330
               TabIndex        =   68
               Top             =   6870
               Width           =   1620
               _Version        =   65536
               _ExtentX        =   2857
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
               ForeColor       =   16777215
               BackColor       =   192
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Especial_Seguros 
               Height          =   285
               Left            =   -69120
               TabIndex        =   69
               Top             =   6870
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
               ForeColor       =   16777215
               BackColor       =   192
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
               Alignment       =   4
            End
            Begin MSFlexGridLib.MSFlexGrid grd_MViNCo_Listad 
               Height          =   3465
               Left            =   -74970
               TabIndex        =   70
               Top             =   690
               Width           =   12945
               _ExtentX        =   22834
               _ExtentY        =   6112
               _Version        =   393216
               Rows            =   25
               Cols            =   8
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel SSPanel6 
               Height          =   285
               Left            =   -74940
               TabIndex        =   71
               Top             =   390
               Width           =   750
               _Version        =   65536
               _ExtentX        =   1323
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Cuota"
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel9 
               Height          =   285
               Left            =   -74190
               TabIndex        =   72
               Top             =   390
               Width           =   1400
               _Version        =   65536
               _ExtentX        =   2469
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "F. Vcto"
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel12 
               Height          =   285
               Left            =   -66810
               TabIndex        =   73
               Top             =   390
               Width           =   1500
               _Version        =   65536
               _ExtentX        =   2646
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Saldo Capital"
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin MSFlexGridLib.MSFlexGrid grd_MViCon_Listad 
               Height          =   3465
               Left            =   -74970
               TabIndex        =   74
               Top             =   690
               Width           =   12945
               _ExtentX        =   22834
               _ExtentY        =   6112
               _Version        =   393216
               Rows            =   25
               Cols            =   8
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel SSPanel19 
               Height          =   285
               Left            =   -74940
               TabIndex        =   75
               Top             =   390
               Width           =   750
               _Version        =   65536
               _ExtentX        =   1323
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Cuota"
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel20 
               Height          =   285
               Left            =   -74190
               TabIndex        =   76
               Top             =   390
               Width           =   1400
               _Version        =   65536
               _ExtentX        =   2469
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "F. Vcto"
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel21 
               Height          =   285
               Left            =   -72810
               TabIndex        =   77
               Top             =   390
               Width           =   1500
               _Version        =   65536
               _ExtentX        =   2646
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Capital"
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel38 
               Height          =   285
               Left            =   -68310
               TabIndex        =   78
               Top             =   390
               Width           =   1500
               _Version        =   65536
               _ExtentX        =   2646
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Total Cuota"
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel39 
               Height          =   285
               Left            =   -66810
               TabIndex        =   79
               Top             =   390
               Width           =   1500
               _Version        =   65536
               _ExtentX        =   2646
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Saldo Capital"
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel41 
               Height          =   285
               Left            =   -69810
               TabIndex        =   80
               Top             =   390
               Width           =   1500
               _Version        =   65536
               _ExtentX        =   2646
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Comisión COFIDE"
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel43 
               Height          =   285
               Left            =   -74940
               TabIndex        =   81
               Top             =   390
               Width           =   810
               _Version        =   65536
               _ExtentX        =   1429
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Cuota"
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel58 
               Height          =   285
               Left            =   -74130
               TabIndex        =   82
               Top             =   390
               Width           =   1590
               _Version        =   65536
               _ExtentX        =   2805
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "F. Vcto"
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel60 
               Height          =   285
               Left            =   -72540
               TabIndex        =   83
               Top             =   390
               Width           =   1800
               _Version        =   65536
               _ExtentX        =   3175
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Capital"
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel63 
               Height          =   285
               Left            =   -68940
               TabIndex        =   84
               Top             =   390
               Width           =   1800
               _Version        =   65536
               _ExtentX        =   3175
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Total Cuota"
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel70 
               Height          =   285
               Left            =   -67140
               TabIndex        =   85
               Top             =   390
               Width           =   1800
               _Version        =   65536
               _ExtentX        =   3175
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Saldo Capital"
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel18 
               Height          =   285
               Left            =   -71310
               TabIndex        =   86
               Top             =   390
               Width           =   1500
               _Version        =   65536
               _ExtentX        =   2646
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Interés"
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel5 
               Height          =   285
               Left            =   -71310
               TabIndex        =   87
               Top             =   390
               Width           =   1500
               _Version        =   65536
               _ExtentX        =   2646
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Interés"
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel42 
               Height          =   285
               Left            =   -70740
               TabIndex        =   88
               Top             =   390
               Width           =   1800
               _Version        =   65536
               _ExtentX        =   3175
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Interés"
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel14 
               Height          =   285
               Left            =   -65310
               TabIndex        =   97
               Top             =   390
               Width           =   1320
               _Version        =   65536
               _ExtentX        =   2328
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Indicador Carga"
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel15 
               Height          =   285
               Left            =   -65310
               TabIndex        =   98
               Top             =   390
               Width           =   1320
               _Version        =   65536
               _ExtentX        =   2328
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Indicador Carga"
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel17 
               Height          =   285
               Left            =   -65340
               TabIndex        =   99
               Top             =   390
               Width           =   1350
               _Version        =   65536
               _ExtentX        =   2381
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Indicador Carga"
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel71 
               Height          =   285
               Left            =   60
               TabIndex        =   101
               Top             =   390
               Width           =   810
               _Version        =   65536
               _ExtentX        =   1429
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Cuota"
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel72 
               Height          =   285
               Left            =   870
               TabIndex        =   102
               Top             =   390
               Width           =   1250
               _Version        =   65536
               _ExtentX        =   2205
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "F. Vcto"
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel73 
               Height          =   285
               Left            =   2100
               TabIndex        =   103
               Top             =   390
               Width           =   1250
               _Version        =   65536
               _ExtentX        =   2205
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Capital"
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel74 
               Height          =   285
               Left            =   8340
               TabIndex        =   104
               Top             =   390
               Width           =   1245
               _Version        =   65536
               _ExtentX        =   2205
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Total Cuota"
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel75 
               Height          =   285
               Left            =   9570
               TabIndex        =   105
               Top             =   390
               Width           =   1255
               _Version        =   65536
               _ExtentX        =   2214
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Saldo Capital"
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel76 
               Height          =   285
               Left            =   3350
               TabIndex        =   106
               Top             =   390
               Width           =   1250
               _Version        =   65536
               _ExtentX        =   2205
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Interés"
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel77 
               Height          =   285
               Left            =   10820
               TabIndex        =   107
               Top             =   390
               Width           =   1245
               _Version        =   65536
               _ExtentX        =   2205
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Indicador Carga"
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin MSFlexGridLib.MSFlexGrid grd_CliCon_Listad 
               Height          =   3465
               Left            =   -74970
               TabIndex        =   108
               Top             =   690
               Width           =   12945
               _ExtentX        =   22834
               _ExtentY        =   6112
               _Version        =   393216
               Rows            =   25
               Cols            =   7
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_CliNCon_Listad 
               Height          =   3465
               Left            =   30
               TabIndex        =   109
               Top             =   690
               Width           =   12945
               _ExtentX        =   22834
               _ExtentY        =   6112
               _Version        =   393216
               Rows            =   25
               Cols            =   10
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel SSPanel78 
               Height          =   285
               Left            =   4610
               TabIndex        =   110
               Top             =   390
               Width           =   1250
               _Version        =   65536
               _ExtentX        =   2205
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Seg.Prest."
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel79 
               Height          =   285
               Left            =   5870
               TabIndex        =   111
               Top             =   390
               Width           =   1245
               _Version        =   65536
               _ExtentX        =   2205
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Seg. Vivienda"
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin Threed.SSPanel SSPanel80 
               Height          =   285
               Left            =   7110
               TabIndex        =   112
               Top             =   390
               Width           =   1245
               _Version        =   65536
               _ExtentX        =   2205
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Portes"
               ForeColor       =   16777215
               BackColor       =   16384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               RoundedCorners  =   0   'False
               Outline         =   -1  'True
            End
            Begin VB.Label lbl_Totale 
               Alignment       =   1  'Right Justify
               Caption         =   "Totales ===> US$ "
               Height          =   315
               Index           =   5
               Left            =   -74610
               TabIndex        =   96
               Top             =   6870
               Width           =   1845
            End
            Begin VB.Label lbl_Totale 
               Alignment       =   1  'Right Justify
               Caption         =   "Totales ===> US$ "
               Height          =   315
               Index           =   4
               Left            =   -74610
               TabIndex        =   95
               Top             =   6870
               Width           =   1845
            End
            Begin VB.Label Label1 
               Caption         =   "Totales ==>"
               Height          =   285
               Left            =   -72930
               TabIndex        =   94
               Top             =   1470
               Width           =   945
            End
            Begin VB.Label Label14 
               Caption         =   "Totales ==>"
               Height          =   285
               Left            =   -72930
               TabIndex        =   93
               Top             =   1470
               Width           =   945
            End
            Begin VB.Label Label15 
               Caption         =   "Totales ==>"
               Height          =   285
               Left            =   -73230
               TabIndex        =   92
               Top             =   1470
               Width           =   945
            End
            Begin VB.Label lbl_Totale 
               Alignment       =   1  'Right Justify
               Caption         =   "Totales ===> US$ "
               Height          =   315
               Index           =   1
               Left            =   -74610
               TabIndex        =   91
               Top             =   6870
               Width           =   1845
            End
            Begin VB.Label lbl_Totale 
               Alignment       =   1  'Right Justify
               Caption         =   "Totales ===> US$ "
               Height          =   315
               Index           =   2
               Left            =   -74790
               TabIndex        =   90
               Top             =   6870
               Width           =   1845
            End
            Begin VB.Label lbl_Totale 
               Alignment       =   1  'Right Justify
               Caption         =   "Totales ===> US$ "
               Height          =   315
               Index           =   3
               Left            =   -74610
               TabIndex        =   89
               Top             =   6870
               Width           =   1845
            End
         End
      End
   End
End
Attribute VB_Name = "frm_Car_PrePag_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_str_OpeMVi     As String
Dim l_int_PerGra     As Integer
Dim l_int_CuoDbl     As Integer
Dim l_dbl_TasInt     As Double
Dim swtContador      As Integer

Private Sub chkSeleccionar_Click()
   Dim r_Fila As Integer
      
      If grd_Listap.Rows > 0 Then
         If chkSeleccionar.Value = 0 Then
            For r_Fila = 0 To grd_Listap.Rows - 1
                grd_Listap.TextMatrix(r_Fila, 7) = ""
            Next r_Fila
         End If
         If chkSeleccionar.Value = 1 Then
            For r_Fila = 0 To grd_Listap.Rows - 1
                grd_Listap.TextMatrix(r_Fila, 7) = "X"
            Next r_Fila
         End If
         Call gs_RefrescaGrid(grd_Listap)
      End If
End Sub

Private Sub cmd_BuscaArc_Click()
    dlg_Guarda.Filter = "Archivos Excel |*.xlsx;*.xls"
    dlg_Guarda.ShowOpen
    txt_NomArc.Text = UCase(dlg_Guarda.FileName)
    Exit Sub
End Sub

Private Sub cmd_Import_Click()
Dim r_Fila  As Integer
Dim r_int_ConSel As Integer
Dim r_int_Contad  As Integer
   'validaciones
   If cmb_TipCro.ListIndex = -1 Then
      MsgBox "Debe ingresar el tipo de cronograma a importar.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipCro)
      Exit Sub
   End If
   
   If Len(Trim(txt_NomArc.Text)) = 0 Then
      MsgBox "Debe ingresar la ubicación y nombre del archivo a importar.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NomArc)
      Exit Sub
   End If
      
   MsgBox "Antes de realizar el proceso de carga verifique los siguiente: " & vbCrLf & " - El archivo excel debe tener el formato del 2007. " & vbCrLf & " - La Columna B del archivo con formato 'dd/mm/aaaa'", vbInformation, modgen_g_str_NomPlt
   If MsgBox("¿Desea realizar la carga del archivo seleccionado?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   If InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Then 'moddat_g_str_CodPrd = "003" Then
      If fs_Carga_ArchivoCronogramasCME Then
          MsgBox "Proceso de carga de archivos finalizada satisfactoriamente.", vbInformation, modgen_g_str_NomPlt
      End If
      
   ElseIf moddat_g_str_CodPrd = "019" Then
      If fs_Carga_ArchivoCronogramaMICASAMAS Then
         MsgBox "Proceso de carga de archivos finalizada satisfactoriamente.", vbInformation, modgen_g_str_NomPlt
      End If
      
   Else
      swtContador = 0
      If cmb_TipCro.ListIndex = 1 Then
         If fs_Carga_ArchivoCronogramasFMV Then
            If fs_UbicaCuotas_CronogramasFMV Then
               If fs_Carga_ClienteConcesional Then
                  MsgBox "Proceso de carga de archivos finalizada satisfactoriamente.", vbInformation, modgen_g_str_NomPlt
               End If
           End If
         End If
     
         'solo ingresar si son aportes extraordinarios
         If l_int_CuoDbl <> 1 Then
            If swtContador > 1 Then
               For r_Fila = 0 To grd_MViNCo_Listad.Rows - 1
                  grd_MViNCo_Listad.TextMatrix(r_Fila, 7) = ""
               Next r_Fila
               For r_Fila = 0 To grd_MViCon_Listad.Rows - 1
                  grd_MViCon_Listad.TextMatrix(r_Fila, 7) = ""
               Next r_Fila
            End If
         End If
      
      ElseIf cmb_TipCro.ListIndex = 2 Then
         If fs_Carga_CofideNoConcesional Then
            If fs_UbicaCuotas_CronogramaFMVTNC Then
               MsgBox "Proceso de carga de archivos finalizada satisfactoriamente.", vbInformation, modgen_g_str_NomPlt
            End If
         End If
         
      Else
         If fs_Carga_ClienteNoConcesional Then
            MsgBox "Proceso de carga de archivos finalizada satisfactoriamente.", vbInformation, modgen_g_str_NomPlt
         End If
      End If
   End If
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Generar_Click()
   'Validación
   If cmb_TipCro.ListIndex = 1 Then
      If grd_MViNCo_Listad.Rows = 0 Or grd_MViCon_Listad.Rows = 0 Then
         MsgBox "Debe de haber registros en el FMV - Tramos No Concesional y Concesional.", vbInformation, modgen_g_str_NomPlt
         Exit Sub
      End If
      If Len(RTrim(LTrim(txt_NomArc.Text))) = 0 Then
          MsgBox "Debe ingresar la ubicación y nombre del archivo a importar.", vbExclamation, modgen_g_str_NomPlt
          Exit Sub
      End If
      
      Call gs_LimpiaGrid(grd_CliCon_Listad)
      
      If fs_Carga_ClienteConcesional Then
         MsgBox "Proceso de carga de archivos finalizada satisfactoriamente.", vbInformation, modgen_g_str_NomPlt
      End If
    
   Else
      If grd_CliNCon_Listad.Rows = 0 Then
         MsgBox "Debe de haber registros en el CLIENTE - Tramo No Concesional.", vbInformation, modgen_g_str_NomPlt
         Exit Sub
      End If
      If Len(RTrim(LTrim(txt_NomArc.Text))) = 0 Then
          MsgBox "Debe ingresar la ubicación y nombre del archivo a importar.", vbExclamation, modgen_g_str_NomPlt
          Exit Sub
      End If

      If fs_Carga_ClienteNoConcesional Then
         MsgBox "Proceso de carga de archivos finalizada satisfactoriamente.", vbInformation, modgen_g_str_NomPlt
      End If

   End If
End Sub

Private Sub cmd_ImpCro_Click()
   modmip_g_int_OrdAct = 1
   frm_Ges_CreHip_07.cmd_Cronog.Visible = False
   frm_Ges_CreHip_07.Show 1
End Sub
Private Sub cmd_Grabar_Click()

Dim r_int_NumPro        As Integer
Dim r_str_HorIni        As String
Dim r_int_NumRegCro1    As Integer
Dim r_int_NumRegCro2    As Integer
Dim r_int_NumRegCro3    As Integer
Dim r_int_NumRegCro4    As Integer
Dim r_int_Contad        As Integer
Dim r_int_ConSel        As Integer
Dim r_dbl_MtoCuo        As Double
Dim r_dbl_MtoPpg        As Double

   'valida datos
   If Not fs_Valida_Datos() Then
      Exit Sub
   End If
   
   If cmb_TipCro.ListIndex <> 0 Then
      
      If Me.grd_Listap.Rows > 0 Then 'valida selección
      
         r_int_ConSel = 0
         For r_int_Contad = 0 To grd_Listap.Rows - 1
            If grd_Listap.TextMatrix(r_int_Contad, 7) = "X" Then
               r_int_ConSel = r_int_ConSel + 1
               r_dbl_MtoPpg = r_dbl_MtoPpg + grd_Listap.TextMatrix(r_int_Contad, 2)
            End If
         Next r_int_Contad
         
         If r_int_ConSel = 0 Then
            MsgBox "No se han seleccionado Prepagos a Cargar.", vbInformation, modgen_g_str_NomPlt
            Me.SSTab1(1).Tab = 1
            Exit Sub
         End If
         If grd_MViNCo_Listad.TextMatrix(0, 7) = "X" Then
            MsgBox "No se puede seleccionar primera cuota.", vbInformation, modgen_g_str_NomPlt
            grd_MViNCo_Listad.TextMatrix(0, 7) = ""
            Screen.MousePointer = 0
            Exit Sub
         Else
            r_dbl_MtoCuo = fs_UbicaMonto_Prepago
            If r_dbl_MtoCuo > 0 Then
               If MsgBox("¿Desea cargar el prepago seleccionado de " & Format(r_dbl_MtoPpg, "###,##0.00") & ", siendo el prepago del cronograma de " & Format(r_dbl_MtoCuo, "###,##0.00") & " ?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
                  Screen.MousePointer = 0
                  Exit Sub
               End If
            Else
               If MsgBox("¿Desea cargar el prepago seleccionado?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
                  Screen.MousePointer = 0
                  Exit Sub
               End If
            End If
         End If
      End If
      
   End If
   
   Screen.MousePointer = 11
   r_str_HorIni = Format(Time, "hhmmss")
    
   If cmb_TipCro.ListIndex = 0 Then
       'confirma grabacion
      If MsgBox("¿Desea cargar a la base de datos la información del nuevo cronograma?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
         Screen.MousePointer = 0
         Exit Sub
      End If
      
      '*** Inicializa log
      moddat_g_int_CntErr = 0
      g_str_Parame = ""
      g_str_Parame = "USP_CRE_PROCROCAB ("
      g_str_Parame = g_str_Parame & Format(CDate(moddat_g_str_FecSis), "yyyymmdd") & ", "
      g_str_Parame = g_str_Parame & 0 & ", "
      g_str_Parame = g_str_Parame & 1 & ", "
      g_str_Parame = g_str_Parame & r_str_HorIni & ", "
      g_str_Parame = g_str_Parame & "'" & txt_NomArc.Text & "', "
      g_str_Parame = g_str_Parame & "'" & Dir(txt_NomArc.Text, vbArchive) & "', "
      g_str_Parame = g_str_Parame & 0 & ", "
      g_str_Parame = g_str_Parame & 0 & ", "
      g_str_Parame = g_str_Parame & 0 & ", "
      g_str_Parame = g_str_Parame & 0 & ", "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
      g_str_Parame = g_str_Parame & 1 & " ) "
      
      Do While (moddat_g_int_CntErr = 0)
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 1) Then
            If MsgBox("No se pudo completar el procedimiento. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
               moddat_g_int_CntErr = 1
               Screen.MousePointer = 0
               Exit Sub
            Else
               moddat_g_int_CntErr = 0
            End If
         Else
            moddat_g_int_CntErr = 1
         End If
      Loop
      
      g_rst_Princi.MoveFirst
      r_int_NumPro = g_rst_Princi!CORRELATIVO
      
      If fs_Actualiza_Cronograma_CLITNC(r_int_NumPro) Then
         Screen.MousePointer = 0
         MsgBox "Actualización realizada satisfactoriamente.", vbInformation, modgen_g_str_NomPlt
      End If
     
         
   Else
      'confirma grabación
      If MsgBox("¿Desea cargar a la base de datos la información de los nuevos cronogramas?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
         Exit Sub
      End If
      
      '*** Inicializa log
      moddat_g_int_CntErr = 0
      g_str_Parame = ""
      g_str_Parame = "USP_CRE_PROCROCAB ("
      g_str_Parame = g_str_Parame & Format(CDate(moddat_g_str_FecSis), "yyyymmdd") & ", "
      g_str_Parame = g_str_Parame & 0 & ", "
      g_str_Parame = g_str_Parame & 1 & ", "
      g_str_Parame = g_str_Parame & r_str_HorIni & ", "
      g_str_Parame = g_str_Parame & "'" & txt_NomArc.Text & "', "
      g_str_Parame = g_str_Parame & "'" & Dir(txt_NomArc.Text, vbArchive) & "', " 'Replace(Dir(txt_NomArc.Text, vbArchive), ",", "")
      g_str_Parame = g_str_Parame & 0 & ", "
      g_str_Parame = g_str_Parame & 0 & ", "
      g_str_Parame = g_str_Parame & 0 & ", "
      g_str_Parame = g_str_Parame & 0 & ", "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
      g_str_Parame = g_str_Parame & 1 & " ) "
      
      Do While (moddat_g_int_CntErr = 0)
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 1) Then
            If MsgBox("No se pudo completar el procedimiento. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
               moddat_g_int_CntErr = 1
               Screen.MousePointer = 0
               Exit Sub
            Else
               moddat_g_int_CntErr = 0
            End If
         Else
            moddat_g_int_CntErr = 1
         End If
      Loop
      
      g_rst_Princi.MoveFirst
      r_int_NumPro = g_rst_Princi!CORRELATIVO
      
      'proceso de grabacion
      If cmb_TipCro.ListIndex = 2 Then
         If fs_Actualiza_Cronograma_FMVTNC(r_int_NumPro) Then
            Screen.MousePointer = 0
            MsgBox "Actualización realizada satisfactoriamente.", vbInformation, modgen_g_str_NomPlt
         End If
      ElseIf Me.cmb_TipCro.ListIndex = 3 Then
         If fs_Actualiza_Cronograma_CME(r_int_NumPro) Then
            Screen.MousePointer = 0
            MsgBox "Actualización realizada satisfactoriamente.", vbInformation, modgen_g_str_NomPlt
         End If
      ElseIf fs_Actualiza_Cronograma_FMVTNC(r_int_NumPro) Then
         If fs_Actualiza_Cronograma_FMVTC(r_int_NumPro) Then
            If fs_Actualiza_Cronograma_CLITC(r_int_NumPro) Then
               Screen.MousePointer = 0
               MsgBox "Actualización realizada satisfactoriamente.", vbInformation, modgen_g_str_NomPlt
            End If
         End If
      End If
      
      For r_int_Contad = 0 To grd_Listap.Rows - 1
         If grd_Listap.TextMatrix(r_int_Contad, 7) = "X" Then
            'Actualiza el estado de la tabla CRE_PPGCAB
            g_str_Parame = ""
            g_str_Parame = g_str_Parame & " USP_ACTUALIZA_CRE_PPGCAB ("
            g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "', "
            g_str_Parame = g_str_Parame & "" & Format(grd_Listap.TextMatrix(r_int_Contad, 1), "yyyymmdd") & " , 5, 0, " 'g_rst_Princi!PPGCAB_FECPPG
            g_str_Parame = g_str_Parame & Format(CDate(moddat_g_str_FecSis), "yyyymmdd") & ", 0 , 0 ) "
         
            If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
               MsgBox "No se pudo completar la actualización del estado de los datos.", vbInformation, modgen_g_con_PltPar
               Exit Sub
            End If
         End If
      Next r_int_Contad
   End If
   
   'Cantidad de Filas por cada Cronograma que se deben guardar.
   For r_int_Contad = 0 To grd_CliNCon_Listad.Rows - 1
      If grd_CliNCon_Listad.TextMatrix(r_int_Contad, 9) = "X" Then
         r_int_NumRegCro1 = CInt(grd_CliNCon_Listad.Rows) - r_int_Contad
      End If
   Next r_int_Contad
   
   For r_int_Contad = 0 To grd_CliCon_Listad.Rows - 1
      If grd_CliCon_Listad.TextMatrix(r_int_Contad, 6) = "X" Then
         r_int_NumRegCro2 = CInt(grd_CliCon_Listad.Rows) - r_int_Contad
      End If
   Next r_int_Contad
   
   For r_int_Contad = 0 To grd_MViNCo_Listad.Rows - 1
      If grd_MViNCo_Listad.TextMatrix(r_int_Contad, 7) = "X" Then
         r_int_NumRegCro3 = CInt(grd_MViNCo_Listad.Rows) - r_int_Contad
      End If
   Next r_int_Contad
   
   For r_int_Contad = 0 To grd_MViCon_Listad.Rows - 1
      If grd_MViCon_Listad.TextMatrix(r_int_Contad, 7) = "X" Then
         r_int_NumRegCro4 = CInt(grd_MViCon_Listad.Rows) - r_int_Contad
      End If
   Next r_int_Contad
   
   '*** Finaliza log
   moddat_g_int_CntErr = 0
   g_str_Parame = "USP_CRE_PROCROCAB ("
   g_str_Parame = g_str_Parame & Format(CDate(moddat_g_str_FecSis), "yyyymmdd") & ", "
   g_str_Parame = g_str_Parame & r_int_NumPro & ", "
   g_str_Parame = g_str_Parame & 2 & ", "
   g_str_Parame = g_str_Parame & "'', "
   g_str_Parame = g_str_Parame & "'', "
   g_str_Parame = g_str_Parame & "'', "
   g_str_Parame = g_str_Parame & r_int_NumRegCro1 & ", "
   g_str_Parame = g_str_Parame & r_int_NumRegCro2 & ", "
   g_str_Parame = g_str_Parame & r_int_NumRegCro3 & ", "
   g_str_Parame = g_str_Parame & r_int_NumRegCro4 & ", "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
   g_str_Parame = g_str_Parame & 0 & " ) "
   
   Do While (moddat_g_int_CntErr = 0)
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 1) Then
         If MsgBox("No se pudo completar el procedimiento. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
            moddat_g_int_CntErr = 0
            
         Else
            moddat_g_int_CntErr = 0
         End If
      Else
         moddat_g_int_CntErr = 1
      End If
   Loop
   
   Screen.MousePointer = 0
   Unload Me
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   Me.cmb_TipCro.AddItem ("CRONOGRAMA 1")
   Me.cmb_TipCro.AddItem ("CRONOGRAMA 2, 3 y 4")
   Me.cmb_TipCro.AddItem ("CRONOGRAMA 3")
   Me.cmb_TipCro.AddItem ("CRONOGRAMA 5")
   
   Call fs_IniciaGrid
   Call fs_Limpiar
   Call modmip_gs_DatNumOpe(moddat_g_str_NumOpe, grd_Listad)
   Call fs_Buscar_DatosCredito
   Call fs_Buscar_DatosPrepago
   Call fs_Configura_CargaProd
   
   Call gs_CentraForm(Me)
   Call gs_SetFocus(txt_NomArc)
   Screen.MousePointer = 0
End Sub

Private Sub fs_IniciaGrid()
   'Datos del Credito
   grd_Listad.ColWidth(0) = 3000 '2800
   grd_Listad.ColWidth(1) = 8400
   grd_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
   grd_Listad.Rows = 0
   
'    'Datos del Prepago
'   grd_Listap.ColWidth(0) = 2325
'   grd_Listap.ColWidth(1) = 1860
'   grd_Listap.ColWidth(2) = 1860
'   grd_Listap.ColWidth(3) = 1860
'   grd_Listap.ColWidth(4) = 1860 'Seleccionar
'   grd_Listap.ColAlignment(1) = flexAlignCenterCenter
'   grd_Listap.ColAlignment(2) = flexAlignRightCenter
'   grd_Listap.ColAlignment(3) = flexAlignCenterCenter
'   grd_Listap.ColAlignment(4) = flexAlignCenterCenter
'   grd_Listap.Rows = 0
    'Datos del Prepago
   grd_Listap.ColWidth(0) = 2085 '2325
   grd_Listap.ColWidth(1) = 1380 '1860
   grd_Listap.ColWidth(2) = 1575 '1860
   grd_Listap.ColWidth(3) = 1575 '1860
   grd_Listap.ColWidth(4) = 1565 '1860 'Seleccionar
   grd_Listap.ColWidth(5) = 1565
   grd_Listap.ColWidth(6) = 1440
   grd_Listap.ColWidth(7) = 1440
   grd_Listap.ColAlignment(1) = flexAlignCenterCenter
   grd_Listap.ColAlignment(2) = flexAlignRightCenter
   grd_Listap.ColAlignment(3) = flexAlignRightCenter
   grd_Listap.ColAlignment(4) = flexAlignRightCenter
   grd_Listap.ColAlignment(5) = flexAlignRightCenter
   grd_Listap.ColAlignment(6) = flexAlignCenterCenter
   grd_Listap.ColAlignment(7) = flexAlignCenterCenter
   grd_Listap.Rows = 0
   
   'FMV No Concesional
   grd_MViNCo_Listad.ColWidth(0) = 750
   grd_MViNCo_Listad.ColWidth(1) = 1400
   grd_MViNCo_Listad.ColWidth(2) = 1500
   grd_MViNCo_Listad.ColWidth(3) = 1500
   grd_MViNCo_Listad.ColWidth(4) = 1500
   grd_MViNCo_Listad.ColWidth(5) = 1500
   grd_MViNCo_Listad.ColWidth(6) = 1500
   grd_MViNCo_Listad.ColWidth(7) = 1350
   grd_MViNCo_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_MViNCo_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_MViNCo_Listad.ColAlignment(2) = flexAlignRightCenter
   grd_MViNCo_Listad.ColAlignment(3) = flexAlignRightCenter
   grd_MViNCo_Listad.ColAlignment(4) = flexAlignRightCenter
   grd_MViNCo_Listad.ColAlignment(5) = flexAlignRightCenter
   grd_MViNCo_Listad.ColAlignment(6) = flexAlignRightCenter
   grd_MViNCo_Listad.ColAlignment(7) = flexAlignCenterCenter
   
   'FMV Concesional
   grd_MViCon_Listad.ColWidth(0) = 750
   grd_MViCon_Listad.ColWidth(1) = 1400
   grd_MViCon_Listad.ColWidth(2) = 1500
   grd_MViCon_Listad.ColWidth(3) = 1500
   grd_MViCon_Listad.ColWidth(4) = 1500
   grd_MViCon_Listad.ColWidth(5) = 1500
   grd_MViCon_Listad.ColWidth(6) = 1500
   grd_MViCon_Listad.ColWidth(7) = 1350
   grd_MViCon_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_MViCon_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_MViCon_Listad.ColAlignment(2) = flexAlignRightCenter
   grd_MViCon_Listad.ColAlignment(3) = flexAlignRightCenter
   grd_MViCon_Listad.ColAlignment(4) = flexAlignRightCenter
   grd_MViCon_Listad.ColAlignment(5) = flexAlignRightCenter
   grd_MViCon_Listad.ColAlignment(6) = flexAlignRightCenter
   grd_MViCon_Listad.ColAlignment(7) = flexAlignCenterCenter
   
   'Cliente Concesional
   grd_CliCon_Listad.ColWidth(0) = 810
   grd_CliCon_Listad.ColWidth(1) = 1560
   grd_CliCon_Listad.ColWidth(2) = 1800
   grd_CliCon_Listad.ColWidth(3) = 1800
   grd_CliCon_Listad.ColWidth(4) = 1800
   grd_CliCon_Listad.ColWidth(5) = 1800
   grd_CliCon_Listad.ColWidth(6) = 1350
   grd_CliCon_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_CliCon_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_CliCon_Listad.ColAlignment(2) = flexAlignRightCenter
   grd_CliCon_Listad.ColAlignment(3) = flexAlignRightCenter
   grd_CliCon_Listad.ColAlignment(4) = flexAlignRightCenter
   grd_CliCon_Listad.ColAlignment(5) = flexAlignRightCenter
   grd_CliCon_Listad.ColAlignment(6) = flexAlignCenterCenter
   
   'Cliente No Concesional
   grd_CliNCon_Listad.ColWidth(0) = 810
   grd_CliNCon_Listad.ColWidth(1) = 1250 '1560
   grd_CliNCon_Listad.ColWidth(2) = 1250 '1800
   grd_CliNCon_Listad.ColWidth(3) = 1250 '1800
   grd_CliNCon_Listad.ColWidth(4) = 1250 '1800
   grd_CliNCon_Listad.ColWidth(5) = 1250 '1800
   grd_CliNCon_Listad.ColWidth(6) = 1250 '1350
   grd_CliNCon_Listad.ColWidth(7) = 1230 '1350
   grd_CliNCon_Listad.ColWidth(8) = 1230 '1350
   grd_CliNCon_Listad.ColWidth(9) = 1250 '1350
   grd_CliNCon_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_CliNCon_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_CliNCon_Listad.ColAlignment(2) = flexAlignRightCenter
   grd_CliNCon_Listad.ColAlignment(3) = flexAlignRightCenter
   grd_CliNCon_Listad.ColAlignment(4) = flexAlignRightCenter
   grd_CliNCon_Listad.ColAlignment(5) = flexAlignRightCenter
   grd_CliNCon_Listad.ColAlignment(6) = flexAlignRightCenter
   grd_CliNCon_Listad.ColAlignment(7) = flexAlignRightCenter
   grd_CliNCon_Listad.ColAlignment(8) = flexAlignRightCenter
   grd_CliNCon_Listad.ColAlignment(9) = flexAlignCenterCenter
End Sub

Private Sub fs_Limpiar()
   Call gs_LimpiaGrid(grd_Listad)
   Call gs_LimpiaGrid(grd_Listap)
   Call gs_LimpiaGrid(grd_MViNCo_Listad)
   Call gs_LimpiaGrid(grd_MViCon_Listad)
   Call gs_LimpiaGrid(grd_CliCon_Listad)
   Call gs_LimpiaGrid(grd_CliNCon_Listad)
   
   txt_NomArc.Text = ""
   cmb_TipCro.Text = ""
   Call gs_SetFocus(cmb_TipCro)  'txt_NomArc
End Sub
Private Sub fs_Buscar_DatosPrepago()

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT PPGCAB_TIPPPG, PPGCAB_TIPPPGPAR, PPGCAB_FECPPG, PPGCAB_FECENV, HIPMAE_MONEDA, "
   g_str_Parame = g_str_Parame & "        PPGCAB_MTODEP, PPGCAB_MTOTOT, PPGCAB_MTOAPL, PPGCAB_APLTNC, PPGCAB_APLTC "
   g_str_Parame = g_str_Parame & "   FROM CRE_PPGCAB INNER JOIN CRE_HIPMAE ON HIPMAE_NUMOPE = PPGCAB_NUMOPE "
   g_str_Parame = g_str_Parame & "  WHERE PPGCAB_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   g_str_Parame = g_str_Parame & "    AND PPGCAB_FLGEST IN (1,2,3,4) "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   'Call fs_Activa(False)
   grd_Listap.Redraw = False
   
   g_rst_Princi.MoveFirst
   
   Do While Not g_rst_Princi.EOF
      grd_Listap.Rows = grd_Listap.Rows + 1
      grd_Listap.Row = grd_Listap.Rows - 1
      
      'tipo de prepago
      grd_Listap.Col = 0
      If g_rst_Princi!PPGCAB_TIPPPG = 1 Then
         If g_rst_Princi!PPGCAB_TIPPPGPAR = 1 Then
            grd_Listap.Text = "PARCIAL - RED MONTO"
         Else
            grd_Listap.Text = "PARCIAL - RED PLAZO"
         End If
      Else
        grd_Listap.Text = "TOTAL"
      End If
      
      'fecha del prepago (formateado)
      grd_Listap.Col = 1
      grd_Listap.Text = gf_FormatoFecha(CStr(g_rst_Princi!PPGCAB_FECPPG))
      
      'importe del prepago (formateado)
      grd_Listap.Col = 2
      If g_rst_Princi!PPGCAB_TIPPPG = 1 Then
         If g_rst_Princi!HIPMAE_MONEDA = 1 Then
            grd_Listap.Text = "S/.   " & Format(g_rst_Princi!PPGCAB_MTODEP, "###,###,###,##0.00")
         Else
            grd_Listap.Text = "US$   " & Format(g_rst_Princi!PPGCAB_MTODEP, "###,###,###,##0.00")
         End If
      Else
         If g_rst_Princi!HIPMAE_MONEDA = 1 Then
            grd_Listap.Text = "S/.   " & Format(g_rst_Princi!PPGCAB_MTOTOT, "###,###,###,##0.00")
         Else
            grd_Listap.Text = "US$   " & Format(g_rst_Princi!PPGCAB_MTOTOT, "###,###,###,##0.00")
         End If
      End If
      
      '''''
      grd_Listap.Col = 3
         If g_rst_Princi!HIPMAE_MONEDA = 1 Then
            grd_Listap.Text = "S/.   " & Format(g_rst_Princi!PPGCAB_MTOAPL, "###,###,###,##0.00")
         Else
            grd_Listap.Text = "US$   " & Format(g_rst_Princi!PPGCAB_MTOAPL, "###,###,###,##0.00")
         End If
      'grd_Listap.Text = Format(g_rst_Princi!PPGCAB_MTOAPL, "###,###,###,##0.00")
        
      grd_Listap.Col = 4
         If g_rst_Princi!HIPMAE_MONEDA = 1 Then
            grd_Listap.Text = "S/.   " & Format(g_rst_Princi!PPGCAB_APLTNC, "###,###,###,##0.00")
         Else
            grd_Listap.Text = "US$   " & Format(g_rst_Princi!PPGCAB_APLTNC, "###,###,###,##0.00")
         End If
      'grd_Listap.Text = Format(g_rst_Princi!PPGCAB_APLTNC, "###,###,###,##0.00")
      
      grd_Listap.Col = 5
         If g_rst_Princi!HIPMAE_MONEDA = 1 Then
            grd_Listap.Text = "S/.   " & Format(g_rst_Princi!PPGCAB_APLTC, "###,###,###,##0.00")
         Else
            grd_Listap.Text = "US$   " & Format(g_rst_Princi!PPGCAB_APLTC, "###,###,###,##0.00")
         End If
      'grd_Listap.Text = Format(g_rst_Princi!PPGCAB_APLTC, "###,###,###,##0.00")
      
      'fecha de envío a COFIDE
      grd_Listap.Col = 6
      If IsNull(g_rst_Princi!PPGCAB_FECENV) Then
         grd_Listap.Text = ""
      Else
         grd_Listap.Text = gf_FormatoFecha(CStr(g_rst_Princi!PPGCAB_FECENV))
      End If
      
      g_rst_Princi.MoveNext
   Loop
   
   grd_Listap.Redraw = True
   
   If grd_Listap.Rows > 0 Then
      grd_Listap.Enabled = True
   End If
   
   Call gs_UbiIniGrid(grd_Listap)
   Call gs_SetFocus(grd_Listap)
   
End Sub

Private Sub fs_Buscar_DatosCredito()
Dim r_str_CodPry     As String
Dim r_str_NomPry     As String
Dim r_str_CodBco     As String
   
   'Buscando Información del Crédito
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_HIPMAE "
   g_str_Parame = g_str_Parame & " WHERE HIPMAE_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   g_str_Parame = g_str_Parame & "   AND HIPMAE_SITUAC = 2 "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If

   g_rst_Princi.MoveFirst
   
   'Almacenando en Variables Globales
   moddat_g_int_TipDoc = g_rst_Princi!HIPMAE_TDOCLI
   moddat_g_str_NumDoc = Trim(g_rst_Princi!HIPMAE_NDOCLI)
   moddat_g_str_NumSol = Trim(g_rst_Princi!HIPMAE_NUMSOL)
   moddat_g_str_NumOpe = Trim(g_rst_Princi!HIPMAE_NUMOPE)
   
   'Obteniendo Nombre de Cliente
   moddat_g_str_NomCli = moddat_gf_Buscar_NomCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc)
   
   'Obteniendo Nombre y DOI de Cónyuge
   moddat_g_int_CygTDo = g_rst_Princi!HIPMAE_TDOCYG
   moddat_g_str_CygNDo = ""
   moddat_g_str_CygNom = ""
   
   If moddat_g_int_CygTDo > 0 Then
      moddat_g_str_CygNDo = Trim(g_rst_Princi!HIPMAE_NDOCYG & "")
      moddat_g_str_CygNom = moddat_gf_Buscar_NomCli(moddat_g_int_CygTDo, moddat_g_str_CygNDo)
   End If
   
   'Obteniendo Descripción de Producto
   moddat_g_str_CodPrd = Trim(g_rst_Princi!HIPMAE_CODPRD)
   moddat_g_str_NomPrd = moddat_gf_Consulta_Produc(Trim(g_rst_Princi!HIPMAE_CODPRD))
   moddat_g_str_CodSub = Trim(g_rst_Princi!HIPMAE_CODSUB)
   
   'Obeniendo Modalidad de Producto
   moddat_g_str_CodMod = Trim(g_rst_Princi!HIPMAE_CODMOD)
   moddat_g_str_DesMod = moddat_gf_Buscar_NomMod(Trim(g_rst_Princi!HIPMAE_CODPRD), moddat_g_str_CodMod)
   
   'Ejecutivo de Seguimiento
   moddat_g_str_CodEjeSeg = Trim(g_rst_Princi!HIPMAE_EJESEG & "")
   moddat_g_str_NomEjeSeg = moddat_gf_Buscar_NomEje(moddat_g_str_CodEjeSeg)
   
   'Consejero Hipotecario
   moddat_g_str_CodConHip = Trim(g_rst_Princi!HIPMAE_CONHIP & "")
   moddat_g_str_NomConHip = moddat_gf_Buscar_NomEje(moddat_g_str_CodConHip)
   
   'Moneda
   moddat_g_str_Moneda = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!HIPMAE_MONEDA))
   moddat_g_int_TipMon = g_rst_Princi!HIPMAE_MONEDA                           'Moneda Préstamo
   moddat_g_dbl_MtoPre = g_rst_Princi!HIPMAE_MTOPRE                           'Monto Préstamo
   moddat_g_int_CuoPen = g_rst_Princi!HIPMAE_CUOPEN                           'Cuotas Pendientes
   moddat_g_int_TotCuo = g_rst_Princi!HIPMAE_NUMCUO                           'Total de Cuotas
   moddat_g_dbl_SalCap = g_rst_Princi!HIPMAE_SALCAP                           'Saldo Capital
   moddat_g_str_FecApr = gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECDES))    'Fecha Desembolso
   l_dbl_TasInt = g_rst_Princi!HIPMAE_TASINT
   l_int_PerGra = g_rst_Princi!HIPMAE_PERGRA
   l_int_CuoDbl = g_rst_Princi!HIPMAE_CUOANO
   l_str_OpeMVi = ""
   If Not IsNull(Trim(g_rst_Princi!HIPMAE_OPEMVI)) Then
      l_str_OpeMVi = Trim(g_rst_Princi!HIPMAE_OPEMVI)
   End If
   
   'Situación de Crédito
   moddat_g_int_Situac = g_rst_Princi!HIPMAE_SITUAC
   moddat_g_str_Situac = moddat_gf_Consulta_ParDes("027", CStr(g_rst_Princi!HIPMAE_SITUAC))
   
   'Obteniendo Información del Inmueble
   Call moddat_gs_Consulta_DatInm(moddat_g_str_NumSol, moddat_g_str_Direcc, moddat_g_str_Distri, r_str_CodPry, r_str_NomPry, r_str_CodBco)
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Configura_CargaProd()
   Select Case moddat_g_str_CodPrd > 0
      Case InStr(moddat_g_str_AgrTMIC, moddat_g_str_CodPrd) '"002", "006", "011"
         MsgBox "La operación no tiene información que cargar.", vbExclamation, modgen_g_str_NomPlt
         cmd_BuscaArc.Enabled = False
         cmd_Import.Enabled = False
         cmd_Grabar.Enabled = False
         tab_Cronog.TabCaption(0) = ""
         tab_Cronog.TabCaption(1) = ""
         tab_Cronog.TabCaption(2) = ""
        ' tab_Cronog.TabCaption(3) = ""
         tab_Cronog.TabVisible(0) = False
         tab_Cronog.TabVisible(1) = False
         tab_Cronog.TabVisible(2) = False
        ' tab_Cronog.TabVisible(3) = False
         
      Case InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) Or InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd) '"003", "019", "021", "022", "023"
         cmd_BuscaArc.Enabled = True
         cmd_Import.Enabled = True
         cmd_Grabar.Enabled = True
         tab_Cronog.TabCaption(0) = "FMV - Cronograma"
         tab_Cronog.TabCaption(1) = ""
         tab_Cronog.TabCaption(2) = ""
         'tab_Cronog.TabCaption(3) = ""
         tab_Cronog.TabVisible(0) = True
         tab_Cronog.TabVisible(1) = False
         tab_Cronog.TabVisible(2) = False
         'tab_Cronog.TabVisible(3) = False
      
      Case Else
         cmd_BuscaArc.Enabled = True
         cmd_Import.Enabled = True
         cmd_Grabar.Enabled = True
         tab_Cronog.TabCaption(0) = "FMV - No Concesional"
         tab_Cronog.TabCaption(1) = "FMV - Concesional"
         tab_Cronog.TabCaption(2) = "CLIENTE - Concesional"
         'tab_Cronog.TabCaption(3) = "CLIENTE - No Concesional"
         
         tab_Cronog.TabVisible(0) = True
         tab_Cronog.TabVisible(1) = True
         tab_Cronog.TabVisible(2) = True
         'tab_Cronog.TabVisible(3) = False
   End Select
End Sub

Private Sub grd_CliNCon_Listad_DblClick()
   If grd_CliNCon_Listad.Rows > 0 Then
        grd_CliNCon_Listad.Col = 9
        If grd_CliNCon_Listad.Text = "X" Then
            grd_CliNCon_Listad.Text = ""
        Else
            For swtContador = 0 To grd_CliNCon_Listad.Rows - 1
                grd_CliNCon_Listad.TextMatrix(swtContador, 9) = ""
            Next swtContador
            grd_CliNCon_Listad.Text = "X"
        End If
        Call gs_RefrescaGrid(grd_CliNCon_Listad)
   End If
End Sub



Private Sub grd_Listap_Click()
   If grd_Listap.Rows > 0 Then
      If grd_Listap.TextMatrix(grd_Listap.Row, 7) = "X" Then
         grd_Listap.TextMatrix(grd_Listap.Row, 7) = ""
      Else
         grd_Listap.TextMatrix(grd_Listap.Row, 7) = "X"
      End If
      Call gs_RefrescaGrid(grd_Listap)
   End If
End Sub

'Private Sub grd_Listap_DblClick()
'  If grd_Listap.Rows > 0 Then
'      grd_Listap.Col = 4
'      If grd_Listap.Text = "X" Then
'         grd_Listap.Text = ""
'      Else
'         For swtContador = 0 To grd_Listap.Rows - 1
'            grd_Listap.TextMatrix(swtContador, 4) = ""
'         Next swtContador
'         grd_Listap.Text = "X"
'      End If
'      Call gs_RefrescaGrid(grd_Listap)
'   End If
'End Sub
Private Sub grd_MviCon_Listad_DblClick()
   If grd_MViCon_Listad.Rows > 0 Then
      grd_MViCon_Listad.Col = 7
      If grd_MViCon_Listad.Text = "X" Then
         grd_MViCon_Listad.Text = ""
      Else
         For swtContador = 0 To grd_MViCon_Listad.Rows - 1
            grd_MViCon_Listad.TextMatrix(swtContador, 7) = ""
         Next swtContador
         grd_MViCon_Listad.Text = "X"
      End If
      Call gs_RefrescaGrid(grd_MViCon_Listad)
   End If
End Sub

Private Sub grd_MViCon_Listad_SelChange()
   If grd_MViCon_Listad.Rows > 2 Then
      grd_MViCon_Listad.RowSel = grd_MViCon_Listad.Row
   End If
End Sub

Private Sub grd_MViNCo_Listad_DblClick()
   Dim Index As Integer
   
   If grd_MViNCo_Listad.Rows > 0 Then
        grd_MViNCo_Listad.Col = 7
        If grd_MViNCo_Listad.Text = "X" Then
            grd_MViNCo_Listad.Text = ""
        Else
            For swtContador = 0 To grd_MViNCo_Listad.Rows - 1
               grd_MViNCo_Listad.TextMatrix(swtContador, 7) = ""
            Next swtContador
            grd_MViNCo_Listad.Text = "X"
        End If
        Call gs_RefrescaGrid(grd_MViNCo_Listad)
   End If
End Sub

Private Sub grd_MViNCo_Listad_SelChange()
   If grd_MViNCo_Listad.Rows > 2 Then
      grd_MViNCo_Listad.RowSel = grd_MViNCo_Listad.Row
   End If
End Sub

Private Sub grd_CliCon_Listad_DblClick()
   If grd_CliCon_Listad.Rows > 0 Then
      grd_CliCon_Listad.Col = 6
      If grd_CliCon_Listad.Text = "X" Then
         grd_CliCon_Listad.Text = ""
      Else
         For swtContador = 0 To grd_CliCon_Listad.Rows - 1
            grd_CliCon_Listad.TextMatrix(swtContador, 6) = ""
         Next swtContador
         grd_CliCon_Listad.Text = "X"
      End If
      Call gs_RefrescaGrid(grd_CliCon_Listad)
   End If
End Sub
 
Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

Private Function fs_Carga_ArchivoCronogramasFMV() As Boolean
Dim r_obj_Excel         As Excel.Application
Dim r_int_FilExc        As Integer
Dim r_int_FilGrd        As Integer
Dim r_int_NumCuo        As Integer
Dim r_dbl_SumNoc        As Double
Dim r_dbl_SumCon        As Double
Dim r_dat_FecIni        As Date
Dim r_dat_FecFin        As Date
   
   fs_Carga_ArchivoCronogramasFMV = False
   Call gs_LimpiaGrid(grd_MViNCo_Listad)
   Call gs_LimpiaGrid(grd_MViCon_Listad)
   Call gs_LimpiaGrid(grd_CliCon_Listad)
   
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Open FileName:=txt_NomArc.Text
   
   'Valida y Carga Cronograma No Concesional FMV
   r_int_FilExc = 0
   r_int_FilGrd = 0
   r_dat_FecIni = CDate("01/01/2007")
   
   Do While Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 2).Value) <> ""
      If Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 2).Value) = "2" Then
         If r_int_FilExc >= grd_MViNCo_Listad.Rows Then
            grd_MViNCo_Listad.Rows = grd_MViNCo_Listad.Rows + 1
         End If
         
         'verifica número de operación
         If InStr(1, l_str_OpeMVi, Mid(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 1).Value), Len(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 1).Value)) - 4, 5), vbTextCompare) = 0 Then
            Call gs_LimpiaGrid(grd_MViNCo_Listad)
            MsgBox "No coincide el numero de operación MIVIVIENDA del sistema con el numero de contrato del archivo." & vbCrLf & "Favor verificar.", vbCritical, modgen_g_str_NomPlt
            GoTo Salir
         End If
         
         r_int_NumCuo = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 3).Value)
         
         If l_int_PerGra = 0 Then
            If Not IsDate(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 4).Value)) Then
               Call gs_LimpiaGrid(grd_MViNCo_Listad)
               MsgBox "Fecha Vencimiento invalida (FMV - TNC - Cuota: " & CStr(r_int_NumCuo) & ").", vbCritical, modgen_g_str_NomPlt
               GoTo Salir
            End If
            If Not Val(r_obj_Excel.Cells(r_int_FilExc + 2, 7).Value) > 0 Then
               Call gs_LimpiaGrid(grd_MViNCo_Listad)
               MsgBox "Capital debe ser mayor a cero (FMV - TNC - Cuota: " & CStr(r_int_NumCuo) & ").", vbCritical, modgen_g_str_NomPlt
               GoTo Salir
            End If
            If Not Val(r_obj_Excel.Cells(r_int_FilExc + 2, 8).Value) > 0 Then
               Call gs_LimpiaGrid(grd_MViNCo_Listad)
               MsgBox "Interes debe ser mayor a cero (FMV - TNC - Cuota: " & CStr(r_int_NumCuo) & ").", vbCritical, modgen_g_str_NomPlt
               GoTo Salir
            End If
            If Not Val(r_obj_Excel.Cells(r_int_FilExc + 2, 9).Value) > 0 Then
               Call gs_LimpiaGrid(grd_MViNCo_Listad)
               MsgBox "Comisión debe ser mayor a cero (FMV - TNC - Cuota: " & CStr(r_int_NumCuo) & ").", vbCritical, modgen_g_str_NomPlt
               GoTo Salir
            End If
         End If
         
         r_dbl_SumNoc = r_obj_Excel.Cells(r_int_FilExc + 2, 7).Value + r_obj_Excel.Cells(r_int_FilExc + 2, 8).Value + r_obj_Excel.Cells(r_int_FilExc + 2, 9).Value
         If Format(r_dbl_SumNoc, "###,###,##0.00") <> Format(r_obj_Excel.Cells(r_int_FilExc + 2, 10).Value, "###,###,##0.00") Then
            Call gs_LimpiaGrid(grd_MViNCo_Listad)
            MsgBox "Total Cuota no es igual a suma de campos Capital, Interes y Comision (FMV - TNC - Cuota: " & CStr(r_int_NumCuo) & ").", vbCritical, modgen_g_str_NomPlt
            GoTo Salir
         End If
         
         grd_MViNCo_Listad.TextMatrix(r_int_FilGrd, 0) = Format(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 3).Value), "000")
         grd_MViNCo_Listad.TextMatrix(r_int_FilGrd, 1) = Format(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 4).Value), "dd/mm/yyyy")
         grd_MViNCo_Listad.TextMatrix(r_int_FilGrd, 2) = Format(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 7).Value), "###,###,##0.00")
         grd_MViNCo_Listad.TextMatrix(r_int_FilGrd, 3) = Format(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 8).Value), "###,###,##0.00")
         grd_MViNCo_Listad.TextMatrix(r_int_FilGrd, 4) = Format(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 9).Value), "###,###,##0.00")
         grd_MViNCo_Listad.TextMatrix(r_int_FilGrd, 5) = Format(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 10).Value), "###,###,##0.00")
         grd_MViNCo_Listad.TextMatrix(r_int_FilGrd, 6) = Format(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 11).Value), "###,###,##0.00")
         '------------------------------------------------------------------------------------------
         If grd_MViNCo_Listad.Rows > 1 Then
            Dim swtCapitalOld As Double
            Dim swtCapitalNew As Double
         
            swtCapitalOld = CDbl(grd_MViNCo_Listad.TextMatrix(r_int_FilGrd - 1, 5))
            swtCapitalNew = CDbl(grd_MViNCo_Listad.TextMatrix(r_int_FilGrd, 5))
            If swtCapitalOld <> swtCapitalNew Then
               swtContador = swtContador + 1
             End If
         End If
         '------------------------------------------------------------------------------------------
         r_dat_FecFin = Format(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 4).Value), "dd/mm/yyyy")
         If r_dat_FecIni > r_dat_FecFin Then
            MsgBox "TNC: Fecha de vencimiento de la cuota " & Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 3).Value) & " es menor que la anterior.", vbCritical, modgen_g_str_NomPlt
            GoTo Salir
         End If
         r_dat_FecIni = r_dat_FecFin
         
         r_int_FilGrd = r_int_FilGrd + 1
      End If
      
      r_int_FilExc = r_int_FilExc + 1
   Loop
   
   If r_int_NumCuo = 0 Then
      MsgBox "El archivo seleccionado no tiene el formato adecuado.", vbCritical, modgen_g_str_NomPlt
      GoTo Salir
   End If
   
   'Valida y Carga Cronograma Concesional FMV
   r_int_FilExc = r_int_NumCuo
   r_int_FilGrd = 0
   r_dat_FecIni = CDate("01/01/2007")
   
   Dim cadena As String
   'Modificando
   Do While Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 2).Value) = ""
      r_int_FilExc = r_int_FilExc + 1
   Loop
   
   cadena = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 2).Value)
   If Len(RTrim(LTrim(cadena))) > 2 Then
      r_int_FilExc = r_int_FilExc + 1
   End If
   
   Do While Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 2).Value) = "1"
   
      If Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 2).Value) = "1" Then
         If r_int_FilGrd >= grd_MViCon_Listad.Rows Then
            grd_MViCon_Listad.Rows = grd_MViCon_Listad.Rows + 1
         End If
         
         r_int_NumCuo = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 3).Value)
         
         If Not IsDate(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 4).Value)) Then
            Call gs_LimpiaGrid(grd_MViNCo_Listad)
            Call gs_LimpiaGrid(grd_MViCon_Listad)
            MsgBox "Fecha Vencimiento invalida en FMV - TC - Cuota: " & CStr(r_int_NumCuo) & ").", vbCritical, modgen_g_str_NomPlt
            GoTo Salir
         End If
         If Not Val(r_obj_Excel.Cells(r_int_FilExc + 2, 7).Value) > 0 Then
            Call gs_LimpiaGrid(grd_MViNCo_Listad)
            Call gs_LimpiaGrid(grd_MViCon_Listad)
            MsgBox "Capital debe ser mayor a cero (FMV - TC - Cuota: " & CStr(r_int_NumCuo) & ").", vbCritical, modgen_g_str_NomPlt
            GoTo Salir
         End If
         If Not Val(r_obj_Excel.Cells(r_int_FilExc + 2, 8).Value) > 0 Then
            Call gs_LimpiaGrid(grd_MViNCo_Listad)
            Call gs_LimpiaGrid(grd_MViCon_Listad)
            MsgBox "Interes debe ser mayor a cero (FMV - TC - Cuota: " & CStr(r_int_NumCuo) & ").", vbCritical, modgen_g_str_NomPlt
            GoTo Salir
         End If
         If Not Val(r_obj_Excel.Cells(r_int_FilExc + 2, 9).Value) > 0 Then
            Call gs_LimpiaGrid(grd_MViNCo_Listad)
            Call gs_LimpiaGrid(grd_MViCon_Listad)
            MsgBox "Comision debe ser mayor a cero (FMV - TC - Cuota: " & CStr(r_int_NumCuo) & ").", vbCritical, modgen_g_str_NomPlt
            GoTo Salir
         End If
         
         r_dbl_SumCon = r_obj_Excel.Cells(r_int_FilExc + 2, 7).Value + r_obj_Excel.Cells(r_int_FilExc + 2, 8).Value + r_obj_Excel.Cells(r_int_FilExc + 2, 9).Value
         If Format(r_dbl_SumCon, "###,###,##0.00") <> Format(r_obj_Excel.Cells(r_int_FilExc + 2, 10).Value, "###,###,##0.00") Then
            Call gs_LimpiaGrid(grd_MViNCo_Listad)
            Call gs_LimpiaGrid(grd_MViCon_Listad)
            MsgBox "Total Cuota no es igual a suma de campos Capital, Interes y Comision (FMV - TC - Cuota: " & CStr(r_int_NumCuo) & ").", vbCritical, modgen_g_str_NomPlt
            GoTo Salir
         End If
         
         grd_MViCon_Listad.TextMatrix(r_int_FilGrd, 0) = Format(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 3).Value), "000")
         grd_MViCon_Listad.TextMatrix(r_int_FilGrd, 1) = Format(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 4).Value), "dd/mm/yyyy")
         grd_MViCon_Listad.TextMatrix(r_int_FilGrd, 2) = Format(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 7).Value), "###,###,##0.00")
         grd_MViCon_Listad.TextMatrix(r_int_FilGrd, 3) = Format(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 8).Value), "###,###,##0.00")
         grd_MViCon_Listad.TextMatrix(r_int_FilGrd, 4) = Format(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 9).Value), "###,###,##0.00")
         grd_MViCon_Listad.TextMatrix(r_int_FilGrd, 5) = Format(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 10).Value), "###,###,##0.00")
         grd_MViCon_Listad.TextMatrix(r_int_FilGrd, 6) = Format(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 11).Value), "###,###,##0.00")
         
         r_dat_FecFin = Format(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 4).Value), "dd/mm/yyyy")
         If r_dat_FecIni > r_dat_FecFin Then
            MsgBox "TC: Fecha de vencimiento de la cuota " & Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 3).Value) & " es menor que la anterior.", vbCritical, modgen_g_str_NomPlt
            GoTo Salir
         End If
         r_dat_FecIni = r_dat_FecFin
         
         r_int_FilGrd = r_int_FilGrd + 1
      End If
      
      r_int_FilExc = r_int_FilExc + 1
   Loop
   
   If r_int_NumCuo = 0 Then
      MsgBox "El archivo seleccionado no tiene el formato adecuado.", vbCritical, modgen_g_str_NomPlt
      GoTo Salir
   End If
   
   fs_Carga_ArchivoCronogramasFMV = True
   
Salir:
   r_obj_Excel.Quit
   Set r_obj_Excel = Nothing
End Function

Private Function fs_UbicaCuotas_CronogramasFMV() As Boolean
Dim r_int_FilGrd        As Integer
Dim r_int_TotReg        As Integer
Dim r_int_NumCuo        As Integer
Dim r_int_CuoCar        As Integer
Dim r_str_FecNco        As String
Dim r_str_FecCon        As String
Dim r_dbl_CuoTotOld     As Double
Dim r_dbl_CuoTotNew     As Double
   
   fs_UbicaCuotas_CronogramasFMV = False
   
   'Validaciones
   If grd_MViNCo_Listad.Rows = 0 Then
      Exit Function
   End If
   If grd_MViCon_Listad.Rows = 0 Then
      Exit Function
   End If
   
   'Seleccionar Cuota: FMV - TNC (Siempre que no tenga cuotas dobles)
   If l_int_CuoDbl = 1 Then
      r_dbl_CuoTotOld = 0
      r_dbl_CuoTotNew = 0
      r_int_TotReg = grd_MViNCo_Listad.Rows
      
      For r_int_FilGrd = 0 To r_int_TotReg - 1
         r_int_NumCuo = grd_MViNCo_Listad.TextMatrix(r_int_FilGrd, 0)
         r_dbl_CuoTotNew = grd_MViNCo_Listad.TextMatrix(r_int_FilGrd, 5)
         
         If r_dbl_CuoTotNew <> r_dbl_CuoTotOld Then
            If r_int_NumCuo <> r_int_TotReg Then
               r_int_CuoCar = r_int_NumCuo
               r_str_FecNco = grd_MViNCo_Listad.TextMatrix(r_int_FilGrd, 1)
            End If
         End If
         
         r_dbl_CuoTotOld = r_dbl_CuoTotNew
      Next r_int_FilGrd
      grd_MViNCo_Listad.TextMatrix(r_int_CuoCar - 1, 7) = "X"
   End If
   
   'Seleccionar Cuota: FMV - TC
   If Len(Trim(r_str_FecNco)) > 0 Then
      r_int_TotReg = grd_MViCon_Listad.Rows
      For r_int_FilGrd = 0 To r_int_TotReg - 1
         r_int_NumCuo = grd_MViCon_Listad.TextMatrix(r_int_FilGrd, 0)
         r_str_FecCon = grd_MViCon_Listad.TextMatrix(r_int_FilGrd, 1)
         
         If CDate(r_str_FecCon) > CDate(r_str_FecNco) Then
            grd_MViCon_Listad.TextMatrix(r_int_NumCuo - 1, 7) = "X"
            Exit For
         End If
      Next
   End If
   
   fs_UbicaCuotas_CronogramasFMV = True
End Function

Private Function fs_Carga_ClienteConcesional() As Boolean
Dim r_rst_ConCli        As ADODB.Recordset
Dim r_int_NumCuo        As Integer
Dim r_int_NumCuoCar     As Integer
Dim r_int_NumFilGrd     As Integer
Dim r_int_CuoTotCof     As Integer
Dim r_int_CuoCarCof     As Integer
Dim r_str_FecCarCof     As String
Dim r_str_FecCalIni     As String
Dim r_str_FecCalFin     As String
Dim r_dbl_MtoSalCap     As Double
Dim r_dbl_MtoCapCuo     As Double
Dim r_dbl_MtoIntCuo     As Double
   
   fs_Carga_ClienteConcesional = False
   r_int_CuoTotCof = grd_MViCon_Listad.Rows
   r_str_FecCarCof = ""
   r_str_FecCalIni = ""
   r_str_FecCalFin = ""
   r_dbl_MtoSalCap = 0
   r_int_NumCuo = 0
   
   'Obtiene fecha y numero de cuota del TC Cofide
   For r_int_NumFilGrd = 0 To r_int_CuoTotCof - 1
      If grd_MViCon_Listad.TextMatrix(r_int_NumFilGrd, 7) = "X" Then
         r_int_CuoCarCof = grd_MViCon_Listad.TextMatrix(r_int_NumFilGrd, 0)
         r_str_FecCarCof = grd_MViCon_Listad.TextMatrix(r_int_NumFilGrd, 1)
         r_int_NumCuoCar = fs_ObtieneCuota_CofideTC(moddat_g_str_NumOpe, Format(r_str_FecCarCof, "YYYYMMDD"))
         
         If r_int_NumFilGrd > 0 Then
            r_str_FecCalIni = grd_MViCon_Listad.TextMatrix(r_int_NumFilGrd - 1, 1)
            r_dbl_MtoSalCap = grd_MViCon_Listad.TextMatrix(r_int_NumFilGrd - 1, 6)
         Else
            r_str_FecCalIni = fs_ObtieneFechaInicio_TC(moddat_g_str_NumOpe)
            r_dbl_MtoSalCap = fs_ObtieneSaldoCapital_TC(moddat_g_str_NumOpe)
         End If
         Exit For
      End If
   Next
   
   If (r_int_CuoCarCof = 0) Or (r_str_FecCarCof = "") Then
      MsgBox "No se pudo determinar el numero de cuota a cargar del TC - Cofide.", vbExclamation, modgen_g_str_NomPlt
      Exit Function
   End If
   
   'Carga Cronograma del Cliente Concesional
   If r_int_NumCuoCar > 1 Then
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT HIPCUO_NUMCUO, HIPCUO_FECVCT, HIPCUO_CAPITA, HIPCUO_INTERE, HIPCUO_SALCAP "
      g_str_Parame = g_str_Parame & "  FROM CRE_HIPCUO "
      g_str_Parame = g_str_Parame & " WHERE HIPCUO_NUMOPE = '" & moddat_g_str_NumOpe & "' "
      g_str_Parame = g_str_Parame & "   AND HIPCUO_TIPCRO = 2 "
      g_str_Parame = g_str_Parame & " ORDER BY HIPCUO_NUMCUO "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         'Exit Function
      End If

      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         grd_CliCon_Listad.Redraw = False
         g_rst_Princi.MoveFirst
         
         Do While Not g_rst_Princi.EOF
            If r_int_NumCuoCar > g_rst_Princi!HIPCUO_NUMCUO Then
               grd_CliCon_Listad.Rows = grd_CliCon_Listad.Rows + 1
               grd_CliCon_Listad.Row = grd_CliCon_Listad.Rows - 1
               grd_CliCon_Listad.Col = 0
               grd_CliCon_Listad.Text = Format(g_rst_Princi!HIPCUO_NUMCUO, "000")
               grd_CliCon_Listad.Col = 1
               grd_CliCon_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECVCT))
               grd_CliCon_Listad.Col = 2
               grd_CliCon_Listad.Text = Format(g_rst_Princi!HIPCUO_CAPITA, "###,###,##0.00")
               grd_CliCon_Listad.Col = 3
               grd_CliCon_Listad.Text = Format(g_rst_Princi!HIPCUO_INTERE, "###,###,##0.00")
               grd_CliCon_Listad.Col = 4
               grd_CliCon_Listad.Text = Format(g_rst_Princi!HIPCUO_CAPITA + g_rst_Princi!HIPCUO_INTERE, "###,###,##0.00")
               grd_CliCon_Listad.Col = 5
               grd_CliCon_Listad.Text = Format(g_rst_Princi!HIPCUO_SALCAP, "###,###,##0.00")
               r_int_NumCuo = g_rst_Princi!HIPCUO_NUMCUO
            End If
            g_rst_Princi.MoveNext
         Loop
         
         g_rst_Princi.Close
         Set g_rst_Princi = Nothing
      End If
   End If
   
   'Carga Cronograma de Cofide Concesional (Para calculo de interes de la cuota)
   For r_int_NumFilGrd = 0 To r_int_CuoTotCof - 1
      If (grd_MViCon_Listad.TextMatrix(r_int_NumFilGrd, 0)) >= CInt(r_int_CuoCarCof) Then
         grd_CliCon_Listad.Rows = grd_CliCon_Listad.Rows + 1
         grd_CliCon_Listad.Row = grd_CliCon_Listad.Rows - 1
         r_int_NumCuo = r_int_NumCuo + 1
         
         'Numero de cuota
         grd_CliCon_Listad.Col = 0
         grd_CliCon_Listad.Text = Format(r_int_NumCuo, "000")
         
         'Fecha de vencimiento
         grd_CliCon_Listad.Col = 1
         r_str_FecCalFin = fs_ObtieneCuota_ClienteTC(moddat_g_str_NumOpe, r_int_NumCuo)
         
         If r_str_FecCalFin = "" Then
            MsgBox "N° de Cuotas del Cronograma COFIDE es mayor al N° de Cuotas del Cronograma Cliente ", vbInformation, modgen_g_str_NomPlt
            r_str_FecCalFin = Format(CDate(DateAdd("M", 6, grd_CliCon_Listad.TextMatrix(grd_CliCon_Listad.Row - 1, 1))), "dd/mm/yyyy")
            'r_str_FecCalFin = Format(CDate(DateAdd("M", 6, r_str_FecCalIni)), "dd/mm/yyyy")
            grd_CliCon_Listad.Text = r_str_FecCalFin
         Else
            grd_CliCon_Listad.Text = r_str_FecCalFin
         End If
         'Capital de la cuota
         grd_CliCon_Listad.Col = 2
         r_dbl_MtoCapCuo = grd_MViCon_Listad.TextMatrix(r_int_NumFilGrd, 2)
         grd_CliCon_Listad.Text = Format(r_dbl_MtoCapCuo, "###,###,##0.00")
         
         'Interes de la cuota
         grd_CliCon_Listad.Col = 3
         r_dbl_MtoIntCuo = r_dbl_MtoSalCap * (1 + (l_dbl_TasInt / 100)) ^ ((DateDiff("d", CDate(r_str_FecCalIni), CDate(r_str_FecCalFin))) / 360) - r_dbl_MtoSalCap
         grd_CliCon_Listad.Text = Format(r_dbl_MtoIntCuo, "###,###,##0.00")
         
         'Monto de la cuota
         grd_CliCon_Listad.Col = 4
         grd_CliCon_Listad.Text = Format(r_dbl_MtoCapCuo + r_dbl_MtoIntCuo, "###,###,##0.00")
         
         'Saldo capital
         grd_CliCon_Listad.Col = 5
         grd_CliCon_Listad.Text = grd_MViCon_Listad.TextMatrix(r_int_NumFilGrd, 6)
         
         'Marca
         If r_int_NumCuo = r_int_NumCuoCar Then
            grd_CliCon_Listad.Col = 6
            grd_CliCon_Listad.Text = "X"
         End If
         
         'Inicializa Variables
         r_str_FecCalIni = r_str_FecCalFin
         r_dbl_MtoSalCap = grd_MViCon_Listad.TextMatrix(r_int_NumFilGrd, 6)
      End If
   Next
   
   grd_CliCon_Listad.Redraw = True
   Call gs_UbiIniGrid(grd_CliCon_Listad)
   
   fs_Carga_ClienteConcesional = True
End Function
Private Function fs_Carga_ClienteNoConcesional() As Boolean
Dim r_obj_Excel         As Excel.Application
Dim r_int_FilExc        As Integer
Dim r_int_FilGrd        As Integer
Dim r_int_NumCuo        As Integer
Dim r_dbl_SumNoc        As Double
Dim r_dbl_SalCap        As Double
Dim r_dat_FecIni        As Date
Dim r_dat_FecFin        As Date
Dim r_int_Cont          As Integer
Dim r_lng_DifFec        As Long
Dim r_int_NumCuoX       As Integer

   fs_Carga_ClienteNoConcesional = False
   Call gs_LimpiaGrid(grd_CliNCon_Listad)

   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Open FileName:=txt_NomArc.Text
   
   'Valida y Carga Cronograma No Concesional CLIENTE
   r_int_FilExc = 0
   r_int_FilGrd = 0
   r_dat_FecIni = CDate("01/01/2007")
   
   Do While Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 2).Value) <> ""

         If r_int_FilExc >= grd_CliNCon_Listad.Rows Then
            grd_CliNCon_Listad.Rows = grd_CliNCon_Listad.Rows + 1
         End If
                
         r_int_NumCuo = IIf(IsNumeric(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 1).Value)), Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 1).Value), 0)
         
         If r_int_NumCuo = 0 Then
            MsgBox "El archivo seleccionado no tiene el formato adecuado.", vbCritical, modgen_g_str_NomPlt
            GoTo Salir
         End If
   
         If l_int_PerGra = 0 Then
            If Not IsDate(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 2).Value)) Then
               Call gs_LimpiaGrid(grd_CliNCon_Listad)
               MsgBox "Fecha Vencimiento inválida (CLIENTE-TNC - Cuota: " & CStr(r_int_NumCuo) & ").", vbCritical, modgen_g_str_NomPlt
               GoTo Salir
            End If
            'If Not Val(r_obj_Excel.Cells(r_int_FilExc + 2, 3).Value) > 0 Then
            '   Call gs_LimpiaGrid(grd_CliNCon_Listad)
            '   MsgBox "Capital debe ser mayor a cero (CLIENTE-TNC - Cuota: " & CStr(r_int_NumCuo) & ").", vbCritical, modgen_g_str_NomPlt
            '   GoTo Salir
            'End If
            If Not Val(r_obj_Excel.Cells(r_int_FilExc + 2, 4).Value) > 0 Then
               Call gs_LimpiaGrid(grd_CliNCon_Listad)
               MsgBox "Interes debe ser mayor a cero (CLIENTE-TNC - Cuota: " & CStr(r_int_NumCuo) & ").", vbCritical, modgen_g_str_NomPlt
               GoTo Salir
            End If
            'If Not Val(r_obj_Excel.Cells(r_int_FilExc + 2, 5).Value) > 0 Then
            '   Call gs_LimpiaGrid(grd_CliNCon_Listad)
            '   MsgBox "Seguro de Préstamo debe ser mayor a cero (CLIENTE-TNC - Cuota: " & CStr(r_int_NumCuo) & ").", vbCritical, modgen_g_str_NomPlt
            '   GoTo Salir
            'End If
            'If Not Val(r_obj_Excel.Cells(r_int_FilExc + 2, 6).Value) > 0 Then
            '   Call gs_LimpiaGrid(grd_CliNCon_Listad)
            '   MsgBox "Seguro de Vivienda debe ser mayor a cero (CLIENTE-TNC - Cuota: " & CStr(r_int_NumCuo) & ").", vbCritical, modgen_g_str_NomPlt
            '   GoTo Salir
            'End If
         End If
         
         r_dbl_SumNoc = r_obj_Excel.Cells(r_int_FilExc + 2, 3).Value + r_obj_Excel.Cells(r_int_FilExc + 2, 4).Value + r_obj_Excel.Cells(r_int_FilExc + 2, 5).Value + r_obj_Excel.Cells(r_int_FilExc + 2, 6).Value + r_obj_Excel.Cells(r_int_FilExc + 2, 7).Value
         If Format(r_dbl_SumNoc, "###,###,##0.00") <> Format(r_obj_Excel.Cells(r_int_FilExc + 2, 8).Value, "###,###,##0.00") Then
            Call gs_LimpiaGrid(grd_CliNCon_Listad)
            MsgBox "Total Cuota no es igual a Suma de Capital, Interes, Seguros y Portes (CLIENTE-TNC - Cuota: " & CStr(r_int_NumCuo) & ").", vbCritical, modgen_g_str_NomPlt
            GoTo Salir
         End If
         
         grd_CliNCon_Listad.TextMatrix(r_int_FilGrd, 0) = Format(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 1).Value), "000")
         grd_CliNCon_Listad.TextMatrix(r_int_FilGrd, 1) = Format(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 2).Value), "dd/mm/yyyy")
         grd_CliNCon_Listad.TextMatrix(r_int_FilGrd, 2) = Format(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 3).Value), "###,###,##0.00")
         grd_CliNCon_Listad.TextMatrix(r_int_FilGrd, 3) = Format(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 4).Value), "###,###,##0.00")
         grd_CliNCon_Listad.TextMatrix(r_int_FilGrd, 4) = Format(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 5).Value), "###,###,##0.00")
         grd_CliNCon_Listad.TextMatrix(r_int_FilGrd, 5) = Format(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 6).Value), "###,###,##0.00")
         grd_CliNCon_Listad.TextMatrix(r_int_FilGrd, 6) = Format(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 7).Value), "###,###,##0.00")
         grd_CliNCon_Listad.TextMatrix(r_int_FilGrd, 7) = Format(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 8).Value), "###,###,##0.00")
         grd_CliNCon_Listad.TextMatrix(r_int_FilGrd, 8) = Format(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 9).Value), "###,###,##0.00")
         
         If CDate(Format(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 2).Value), "dd/mm/yyyy")) > CDate(Format(Now, "dd/mm/yyyy")) And r_int_NumCuoX = 0 Then
            grd_CliNCon_Listad.TextMatrix(r_int_FilGrd, 9) = "X"
            r_int_NumCuoX = CInt(grd_CliNCon_Listad.TextMatrix(r_int_FilGrd, 0))
         End If

         r_dat_FecFin = Format(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 2).Value), "dd/mm/yyyy")
         If r_dat_FecIni > r_dat_FecFin Then
            MsgBox "CLIENTE-TNC: Fecha de vencimiento de la cuota " & Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 1).Value) & " es menor que la anterior.", vbCritical, modgen_g_str_NomPlt
            GoTo Salir
         ElseIf r_dat_FecIni = r_dat_FecFin Then
            MsgBox "CLIENTE-TNC: Fecha de vencimiento de la cuota " & Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 1).Value) & " es igual que la anterior.", vbCritical, modgen_g_str_NomPlt
            GoTo Salir
         End If
         
         'Compara si las fechas de vencimiento son mensuales
         'If l_int_PerGra = 0 Then
            If Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 1).Value) > 1 Then
            r_lng_DifFec = DateDiff("d", r_dat_FecIni, r_dat_FecFin)
               If r_lng_DifFec <> 31 And r_lng_DifFec <> 30 Then
                  If (r_lng_DifFec <> 29 Or r_lng_DifFec <> 28) And CInt(Trim(Mid(r_dat_FecIni, 4, 2))) <> 2 Then
                     If (r_lng_DifFec <> 29 Or r_lng_DifFec <> 28) And CInt(Trim(Mid(r_dat_FecFin, 4, 2))) <> 2 Then
                        MsgBox "CLIENTE-TNC: Fecha de vencimiento de la cuota " & Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 1).Value) & " debe ser mensual respecto a la anterior.", vbCritical, modgen_g_str_NomPlt
                        GoTo Salir
                     End If
                  End If
               End If
            End If
         'End If
         r_dat_FecIni = r_dat_FecFin
         
         r_int_FilGrd = r_int_FilGrd + 1
      
      r_int_FilExc = r_int_FilExc + 1
   Loop
   
   'Obtiene Saldo Capital a partir de la 1era fecha de vencimiento respecto al día de hoy
   r_dbl_SalCap = 0
   For r_int_Cont = r_int_NumCuoX To r_int_FilExc '+ 1
       r_dbl_SalCap = r_dbl_SalCap + Trim(grd_CliNCon_Listad.TextMatrix(r_int_Cont - 1, 2))
   Next r_int_Cont
   
   'Compara que los Saldos sean correctos
   r_int_FilExc = r_int_NumCuoX
   
   Do While Trim(r_obj_Excel.Cells(r_int_FilExc + 1, 3).Value) <> ""

       If CDbl(FormatNumber((CDbl(Trim(r_dbl_SalCap)) - CDbl(Trim(grd_CliNCon_Listad.TextMatrix(r_int_FilExc - 1, 2)))), 2)) - CDbl(FormatNumber(grd_CliNCon_Listad.TextMatrix(r_int_FilExc - 1, 8), 2)) <> 0 Then
         MsgBox "CLIENTE-TNC: El Saldo Capital de la cuota " & Trim(grd_CliNCon_Listad.TextMatrix(r_int_FilExc - 1, 0)) & " es incorrecto.", vbCritical, modgen_g_str_NomPlt
         Call gs_LimpiaGrid(grd_CliNCon_Listad)
         GoTo Salir
      End If
      r_dbl_SalCap = CDbl(grd_CliNCon_Listad.TextMatrix(r_int_FilExc - 1, 8))
      r_int_FilExc = r_int_FilExc + 1
   Loop
   
   If r_int_NumCuo = 0 Then
      MsgBox "El archivo seleccionado no tiene el formato adecuado.", vbCritical, modgen_g_str_NomPlt
      GoTo Salir
   End If
   
   fs_Carga_ClienteNoConcesional = True
   
Salir:
   r_obj_Excel.Quit
   Set r_obj_Excel = Nothing

End Function

Private Function fs_Carga_ArchivoCronogramasCME() As Boolean
Dim r_obj_Excel         As Excel.Application
Dim r_int_FilExc        As Integer
Dim r_int_FilGrd        As Integer
Dim r_int_NumCuo        As Integer
Dim r_dbl_SumNoc        As Double
Dim r_dbl_SumCon        As Double
Dim r_dat_FecIni        As Date
Dim r_dat_FecFin        As Date
Dim r_int_TotReg        As Integer
Dim r_int_CuoCar        As Integer
Dim r_str_FecNco        As String
Dim r_dbl_CuoTotOld     As Double
Dim r_dbl_CuoTotNew     As Double
   
   fs_Carga_ArchivoCronogramasCME = True
   Call gs_LimpiaGrid(grd_MViNCo_Listad)
   
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Open FileName:=txt_NomArc.Text
   
   'Valida y Carga Cronograma
   r_int_FilExc = 0
   r_int_FilGrd = 0
   r_dat_FecIni = CDate("01/01/2007")
   
   Do While Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 3).Value) <> ""
      If Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 2).Value) = "2" Then
         If r_int_FilExc >= grd_MViNCo_Listad.Rows Then
            grd_MViNCo_Listad.Rows = grd_MViNCo_Listad.Rows + 1
         End If
         
         'verifica numero de operacion
         If InStr(1, l_str_OpeMVi, Mid(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 1).Value), Len(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 1).Value)) - 4, 5), vbTextCompare) = 0 Then
            Call gs_LimpiaGrid(grd_MViNCo_Listad)
            MsgBox "No coincide el numero de operación MIVIVIENDA del sistema con el numero de contrato del archivo." & vbCrLf & "Favor verificar.", vbCritical, modgen_g_str_NomPlt
            GoTo Salir
         End If
         
         r_int_NumCuo = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 3).Value)
         
         If Not IsDate(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 4).Value)) Then
            Call gs_LimpiaGrid(grd_MViNCo_Listad)
            MsgBox "Fecha Vencimiento invalida (FMV - TNC - Cuota: " & CStr(r_int_NumCuo) & ").", vbCritical, modgen_g_str_NomPlt
            GoTo Salir
         End If
         If Not Val(r_obj_Excel.Cells(r_int_FilExc + 2, 7).Value) > 0 Then
            Call gs_LimpiaGrid(grd_MViNCo_Listad)
            MsgBox "Capital debe ser mayor a cero (FMV - TNC - Cuota: " & CStr(r_int_NumCuo) & ").", vbCritical, modgen_g_str_NomPlt
            GoTo Salir
         End If
         If Not Val(r_obj_Excel.Cells(r_int_FilExc + 2, 8).Value) > 0 Then
            Call gs_LimpiaGrid(grd_MViNCo_Listad)
            MsgBox "Interes debe ser mayor a cero (FMV - TNC - Cuota: " & CStr(r_int_NumCuo) & ").", vbCritical, modgen_g_str_NomPlt
            GoTo Salir
         End If
         If Not Val(r_obj_Excel.Cells(r_int_FilExc + 2, 9).Value) > 0 Then
            Call gs_LimpiaGrid(grd_MViNCo_Listad)
            MsgBox "Comisión debe ser mayor a cero (FMV - TNC - Cuota: " & CStr(r_int_NumCuo) & ").", vbCritical, modgen_g_str_NomPlt
            GoTo Salir
         End If
         
         r_dbl_SumNoc = r_obj_Excel.Cells(r_int_FilExc + 2, 7).Value + r_obj_Excel.Cells(r_int_FilExc + 2, 8).Value + r_obj_Excel.Cells(r_int_FilExc + 2, 9).Value
         If Format(r_dbl_SumNoc, "###,###,##0.00") <> Format(r_obj_Excel.Cells(r_int_FilExc + 2, 10).Value, "###,###,##0.00") Then
            Call gs_LimpiaGrid(grd_MViNCo_Listad)
            MsgBox "Total Cuota no es igual a suma de campos Capital, Interes y Comision (FMV - TNC - Cuota: " & CStr(r_int_NumCuo) & ").", vbCritical, modgen_g_str_NomPlt
            GoTo Salir
         End If
         
         grd_MViNCo_Listad.TextMatrix(r_int_FilGrd, 0) = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 3).Value)
         grd_MViNCo_Listad.TextMatrix(r_int_FilGrd, 1) = Format(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 4).Value), "dd/mm/yyyy")
         grd_MViNCo_Listad.TextMatrix(r_int_FilGrd, 2) = Format(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 7).Value), "###,###,##0.00")
         grd_MViNCo_Listad.TextMatrix(r_int_FilGrd, 3) = Format(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 8).Value), "###,###,##0.00")
         grd_MViNCo_Listad.TextMatrix(r_int_FilGrd, 4) = Format(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 9).Value), "###,###,##0.00")
         grd_MViNCo_Listad.TextMatrix(r_int_FilGrd, 5) = Format(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 10).Value), "###,###,##0.00")
         grd_MViNCo_Listad.TextMatrix(r_int_FilGrd, 6) = Format(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 11).Value), "###,###,##0.00")
         
         r_dat_FecFin = Format(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 4).Value), "dd/mm/yyyy")
         If r_dat_FecIni > r_dat_FecFin Then
            MsgBox "TNC: Fecha de vencimiento de la cuota " & Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 3).Value) & " es menor que la anterior.", vbCritical, modgen_g_str_NomPlt
            GoTo Salir
         End If
         
         r_dat_FecIni = r_dat_FecFin
         r_int_FilGrd = r_int_FilGrd + 1
      End If
      
      r_int_FilExc = r_int_FilExc + 1
   Loop
   
   If r_int_NumCuo = 0 Then
      MsgBox "El archivo seleccionado no tiene el formato adecuado.", vbCritical, modgen_g_str_NomPlt
      GoTo Salir
      fs_Carga_ArchivoCronogramasCME = False
   End If
     
   'Validaciones
   If grd_MViNCo_Listad.Rows = 0 Then
      Exit Function
   End If
   
   'Seleccionar Cuota: TNC
   r_dbl_CuoTotOld = 0
   r_dbl_CuoTotNew = 0
   r_int_TotReg = grd_MViNCo_Listad.Rows
   
   For r_int_FilGrd = 0 To r_int_TotReg - 1
      r_int_NumCuo = grd_MViNCo_Listad.TextMatrix(r_int_FilGrd, 0)
      r_dbl_CuoTotNew = grd_MViNCo_Listad.TextMatrix(r_int_FilGrd, 5)
      
      If r_dbl_CuoTotNew <> r_dbl_CuoTotOld Then
         If r_int_NumCuo <> r_int_TotReg Then
            r_int_CuoCar = r_int_NumCuo
            r_str_FecNco = grd_MViNCo_Listad.TextMatrix(r_int_FilGrd, 1)
         End If
      End If
      
      r_dbl_CuoTotOld = r_dbl_CuoTotNew
   Next r_int_FilGrd
   grd_MViNCo_Listad.TextMatrix(r_int_CuoCar - 1, 7) = "X"
   Exit Function
   
Salir:
   fs_Carga_ArchivoCronogramasCME = False
   r_obj_Excel.Quit
   Set r_obj_Excel = Nothing
End Function

Private Function fs_Carga_ArchivoCronogramaMICASAMAS() As Boolean
Dim r_obj_Excel         As Excel.Application
Dim r_int_FilExc        As Integer
Dim r_int_FilGrd        As Integer
Dim r_int_TotReg        As Integer
Dim r_int_NumCuo        As Integer
Dim r_dbl_SumNoc        As Double
Dim r_dat_FecIni        As Date
Dim r_dat_FecFin        As Date
Dim r_str_FecNco        As String
Dim r_int_CuoCar        As Integer
Dim r_dbl_CuoTotOld     As Double
Dim r_dbl_CuoTotNew     As Double

   fs_Carga_ArchivoCronogramaMICASAMAS = False
   Call gs_LimpiaGrid(grd_MViNCo_Listad)
   
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Open FileName:=txt_NomArc.Text
   
   'Valida y Carga Cronograma No Concesional FMV
   r_int_FilExc = 0
   r_int_FilGrd = 0
   r_dat_FecIni = CDate("01/01/2007")
   
   Do While Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 1).Value) <> ""
      If Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 2).Value) = "" Then
         If r_int_FilExc >= grd_MViNCo_Listad.Rows Then
            grd_MViNCo_Listad.Rows = grd_MViNCo_Listad.Rows + 1
         End If
         
         'verifica numero de operacion
         If InStr(1, l_str_OpeMVi, Mid(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 1).Value), Len(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 1).Value)) - 4, 5), vbTextCompare) = 0 Then
            Call gs_LimpiaGrid(grd_MViNCo_Listad)
            MsgBox "No coincide el numero de operación MIVIVIENDA del sistema con el numero de contrato del archivo." & vbCrLf & "Favor verificar.", vbCritical, modgen_g_str_NomPlt
            GoTo Salir
         End If
         
         r_int_NumCuo = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 3).Value)
         
         If Not IsDate(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 4).Value)) Then
            Call gs_LimpiaGrid(grd_MViNCo_Listad)
            MsgBox "Fecha Vencimiento invalida (FMV - TNC - Cuota: " & CStr(r_int_NumCuo) & ").", vbCritical, modgen_g_str_NomPlt
            GoTo Salir
         End If
         If Not Val(r_obj_Excel.Cells(r_int_FilExc + 2, 9).Value) > 0 Then
            Call gs_LimpiaGrid(grd_MViNCo_Listad)
            MsgBox "Comisión debe ser mayor a cero (FMV - TNC - Cuota: " & CStr(r_int_NumCuo) & ").", vbCritical, modgen_g_str_NomPlt
            GoTo Salir
         End If
         
         If l_int_PerGra = 0 Then
            If Not Val(r_obj_Excel.Cells(r_int_FilExc + 2, 7).Value) > 0 Then
               Call gs_LimpiaGrid(grd_MViNCo_Listad)
               MsgBox "Capital debe ser mayor a cero (FMV - TNC - Cuota: " & CStr(r_int_NumCuo) & ").", vbCritical, modgen_g_str_NomPlt
               GoTo Salir
            End If
            If Not Val(r_obj_Excel.Cells(r_int_FilExc + 2, 8).Value) > 0 Then
               Call gs_LimpiaGrid(grd_MViNCo_Listad)
               MsgBox "Interes debe ser mayor a cero (FMV - TNC - Cuota: " & CStr(r_int_NumCuo) & ").", vbCritical, modgen_g_str_NomPlt
               GoTo Salir
            End If
         End If
         
         r_dbl_SumNoc = r_obj_Excel.Cells(r_int_FilExc + 2, 7).Value + r_obj_Excel.Cells(r_int_FilExc + 2, 8).Value + r_obj_Excel.Cells(r_int_FilExc + 2, 9).Value
         If Format(r_dbl_SumNoc, "###,###,##0.00") <> Format(r_obj_Excel.Cells(r_int_FilExc + 2, 10).Value, "###,###,##0.00") Then
            Call gs_LimpiaGrid(grd_MViNCo_Listad)
            MsgBox "Total Cuota no es igual a suma de campos Capital, Interes y Comision (FMV - TNC - Cuota: " & CStr(r_int_NumCuo) & ").", vbCritical, modgen_g_str_NomPlt
            GoTo Salir
         End If
         
         grd_MViNCo_Listad.TextMatrix(r_int_FilGrd, 0) = Format(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 3).Value), "000")
         grd_MViNCo_Listad.TextMatrix(r_int_FilGrd, 1) = Format(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 4).Value), "dd/mm/yyyy")
         grd_MViNCo_Listad.TextMatrix(r_int_FilGrd, 2) = Format(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 7).Value), "###,###,##0.00")
         grd_MViNCo_Listad.TextMatrix(r_int_FilGrd, 3) = Format(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 8).Value), "###,###,##0.00")
         grd_MViNCo_Listad.TextMatrix(r_int_FilGrd, 4) = Format(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 9).Value), "###,###,##0.00")
         grd_MViNCo_Listad.TextMatrix(r_int_FilGrd, 5) = Format(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 10).Value), "###,###,##0.00")
         grd_MViNCo_Listad.TextMatrix(r_int_FilGrd, 6) = Format(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 11).Value), "###,###,##0.00")
         
         r_dat_FecFin = Format(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 4).Value), "dd/mm/yyyy")
         If r_dat_FecIni > r_dat_FecFin Then
            MsgBox "TNC: Fecha de vencimiento de la cuota " & Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 3).Value) & " es menor que la anterior.", vbCritical, modgen_g_str_NomPlt
            GoTo Salir
         End If
         r_dat_FecIni = r_dat_FecFin
         
         r_int_FilGrd = r_int_FilGrd + 1
      End If
      
      r_int_FilExc = r_int_FilExc + 1
   Loop
   
   If r_int_NumCuo = 0 Then
      MsgBox "El archivo seleccionado no tiene el formato adecuado.", vbCritical, modgen_g_str_NomPlt
      GoTo Salir
   End If
   
   'Ubica marcador a cargar
   If l_int_CuoDbl = 1 Then
      r_dbl_CuoTotOld = 0
      r_dbl_CuoTotNew = 0
      r_int_TotReg = grd_MViNCo_Listad.Rows
      
      For r_int_FilGrd = 0 To r_int_TotReg - 1
         r_int_NumCuo = grd_MViNCo_Listad.TextMatrix(r_int_FilGrd, 0)
         r_dbl_CuoTotNew = grd_MViNCo_Listad.TextMatrix(r_int_FilGrd, 5)
         
         If r_dbl_CuoTotNew <> r_dbl_CuoTotOld Then
            If r_int_NumCuo <> r_int_TotReg Then
               r_int_CuoCar = r_int_NumCuo
               r_str_FecNco = grd_MViNCo_Listad.TextMatrix(r_int_FilGrd, 1)
            End If
         End If
         
         r_dbl_CuoTotOld = r_dbl_CuoTotNew
      Next r_int_FilGrd
      grd_MViNCo_Listad.TextMatrix(r_int_CuoCar - 1, 7) = "X"
   End If
   
   fs_Carga_ArchivoCronogramaMICASAMAS = True
   
Salir:
   r_obj_Excel.Quit
   Set r_obj_Excel = Nothing
End Function

Private Function fs_ObtieneCuota_CofideTC(ByVal p_NumOpe As String, ByVal p_FecCuo As String) As Integer
Dim r_str_Parame     As String
Dim r_rst_Cuotas     As ADODB.Recordset
   
   fs_ObtieneCuota_CofideTC = 0
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & "SELECT HIPCUO_NUMCUO "
   r_str_Parame = r_str_Parame & "  FROM CRE_HIPCUO "
   r_str_Parame = r_str_Parame & " WHERE HIPCUO_NUMOPE = '" & p_NumOpe & "' "
   r_str_Parame = r_str_Parame & "   AND HIPCUO_TIPCRO = 4 "
   r_str_Parame = r_str_Parame & "   AND HIPCUO_FECVCT = " & p_FecCuo & " "
   
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Cuotas, 3) Then
      Exit Function
   End If
   
   If Not (r_rst_Cuotas.BOF And r_rst_Cuotas.EOF) Then
      r_rst_Cuotas.MoveFirst
      fs_ObtieneCuota_CofideTC = r_rst_Cuotas!HIPCUO_NUMCUO
   Else
      MsgBox "No se encontro numero de cuota del cronograma cofide", vbInformation, modgen_g_str_NomPlt
   End If
   
   r_rst_Cuotas.Close
   Set r_rst_Cuotas = Nothing
End Function

Private Function fs_ObtieneCuota_ClienteTC(ByVal p_NumOpe As String, ByVal p_NumCuo As Integer) As String
Dim r_str_Parame     As String
Dim r_rst_Cuotas     As ADODB.Recordset
   
   fs_ObtieneCuota_ClienteTC = ""
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & "SELECT HIPCUO_FECVCT "
   r_str_Parame = r_str_Parame & "  FROM CRE_HIPCUO "
   r_str_Parame = r_str_Parame & " WHERE HIPCUO_NUMOPE = '" & p_NumOpe & "' "
   r_str_Parame = r_str_Parame & "   AND HIPCUO_TIPCRO = 2 "
   r_str_Parame = r_str_Parame & "   AND HIPCUO_NUMCUO = " & p_NumCuo & " "
   
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Cuotas, 3) Then
      Exit Function
   End If
   
   If Not (r_rst_Cuotas.BOF And r_rst_Cuotas.EOF) Then
      r_rst_Cuotas.MoveFirst
      fs_ObtieneCuota_ClienteTC = gf_FormatoFecha(r_rst_Cuotas!HIPCUO_FECVCT)
   End If
   
   r_rst_Cuotas.Close
   Set r_rst_Cuotas = Nothing
End Function

Private Function fs_ObtieneFechaInicio_TC(ByVal p_NumOpe As String) As String
Dim r_str_Parame     As String
Dim r_rst_HipMae     As ADODB.Recordset
   
   fs_ObtieneFechaInicio_TC = ""
   
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & "SELECT EVACOF_FECDES "
   r_str_Parame = r_str_Parame & "  FROM TRA_EVACOF "
   r_str_Parame = r_str_Parame & " WHERE EVACOF_NUMSOL = (SELECT HIPMAE_NUMSOL FROM CRE_HIPMAE WHERE HIPMAE_NUMOPE = '" & p_NumOpe & "') "
   
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_HipMae, 3) Then
      Exit Function
   End If
   
   If Not (r_rst_HipMae.BOF And r_rst_HipMae.EOF) Then
      r_rst_HipMae.MoveFirst
      fs_ObtieneFechaInicio_TC = gf_FormatoFecha(r_rst_HipMae!EVACOF_FECDES)
   End If
   
   r_rst_HipMae.Close
   Set r_rst_HipMae = Nothing
End Function

Private Function fs_ObtieneSaldoCapital_TC(ByVal p_NumOpe As String) As Double
Dim r_str_Parame     As String
Dim r_rst_HipMae     As ADODB.Recordset
   
   fs_ObtieneSaldoCapital_TC = 0
   
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & "SELECT HIPMAE_IMPCON "
   r_str_Parame = r_str_Parame & "  FROM CRE_HIPMAE "
   r_str_Parame = r_str_Parame & " WHERE HIPMAE_NUMOPE = '" & p_NumOpe & "' "
   
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_HipMae, 3) Then
      Exit Function
   End If
   
   If Not (r_rst_HipMae.BOF And r_rst_HipMae.EOF) Then
      r_rst_HipMae.MoveFirst
      fs_ObtieneSaldoCapital_TC = CDbl(r_rst_HipMae!HIPMAE_IMPCON)
   End If
   
   r_rst_HipMae.Close
   Set r_rst_HipMae = Nothing
End Function

Private Function fs_Valida_Datos() As Boolean
Dim r_int_Contad     As Integer
Dim r_int_IndCar     As Integer

   fs_Valida_Datos = False

   If Len(Trim(txt_NomArc.Text)) = 0 Then
      MsgBox "Debe ingresar la ubicación y nombre del archivo a importar.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NomArc)
      Exit Function
   End If
   
   If cmb_TipCro.ListIndex = 0 Or cmb_TipCro.ListIndex = 3 Then
      'Valida indicador de carga Cliente TNC
      If cmb_TipCro.ListIndex = 0 Then
         r_int_IndCar = 0
         For r_int_Contad = 0 To grd_CliNCon_Listad.Rows - 1
            If grd_CliNCon_Listad.TextMatrix(r_int_Contad, 9) = "X" Then
               r_int_IndCar = r_int_IndCar + 1
            End If
         Next
         If r_int_IndCar <> 1 Then
            MsgBox "El cronograma 'Cliente Tramo No Concesional' tiene indicador de carga errado.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(txt_NomArc)
            Exit Function
         End If
      ElseIf cmb_TipCro.ListIndex = 3 Then
         r_int_IndCar = 0
         For r_int_Contad = 0 To grd_MViNCo_Listad.Rows - 1
            If grd_MViNCo_Listad.TextMatrix(r_int_Contad, 7) = "X" Then
               r_int_IndCar = r_int_IndCar + 1
            End If
         Next
         If r_int_IndCar <> 1 Then
            MsgBox "El cronograma CME, tiene indicador de carga errado.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(txt_NomArc)
            Exit Function
         End If
      End If
   Else
   
      'Valida indicador de carga FMV TNC
      r_int_IndCar = 0
      For r_int_Contad = 0 To grd_MViNCo_Listad.Rows - 1
         If grd_MViNCo_Listad.TextMatrix(r_int_Contad, 7) = "X" Then
            r_int_IndCar = r_int_IndCar + 1
         End If
      Next
      If r_int_IndCar <> 1 Then
         MsgBox "El cronograma 'FMV Tramo No Concesional' tiene indicador de carga errado.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NomArc)
         Exit Function
      End If
      
      If InStr(moddat_g_str_AgrTMIC, moddat_g_str_CodPrd) = 0 And InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd) = 0 Then 'moddat_g_str_CodPrd <> "002" And moddat_g_str_CodPrd <> "006" And moddat_g_str_CodPrd <> "011" And moddat_g_str_CodPrd <> "019" And moddat_g_str_CodPrd <> "021" And moddat_g_str_CodPrd <> "022" And moddat_g_str_CodPrd <> "023" Then
         'Valida indicador de carga FMV TC
         r_int_IndCar = 0
         For r_int_Contad = 0 To grd_MViCon_Listad.Rows - 1
            If grd_MViCon_Listad.TextMatrix(r_int_Contad, 7) = "X" Then
               r_int_IndCar = r_int_IndCar + 1
            End If
         Next
         If r_int_IndCar <> 1 Then
            MsgBox "El cronograma 'FMV Tramo Concesional' tiene indicador de carga errado.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(txt_NomArc)
            Exit Function
         End If
         
         'Valida indicador de carga Cliente TC
         r_int_IndCar = 0
         For r_int_Contad = 0 To grd_CliCon_Listad.Rows - 1
            If grd_CliCon_Listad.TextMatrix(r_int_Contad, 6) = "X" Then
               r_int_IndCar = r_int_IndCar + 1
            End If
         Next
         
         If r_int_IndCar <> 1 Then
            MsgBox "El cronograma 'Cliente Tramo Concesional' tiene indicador de carga errado.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(txt_NomArc)
            Exit Function
         End If
      End If
   End If
   fs_Valida_Datos = True
End Function

Private Function fs_Actualiza_Cronograma_FMVTNC(ByVal r_int_NumPro As Integer) As Boolean
Dim r_int_Contad     As Integer
Dim r_int_IndCar     As Integer
Dim r_str_Cadena     As String
Dim r_rst_Cuotas     As ADODB.Recordset
Dim r_int_NumCuo     As Integer
Dim r_str_FecVct     As String
Dim r_dbl_Capita     As Double
Dim r_dbl_Intere     As Double
Dim r_dbl_SalCap     As Double
Dim r_dbl_ComCof     As Double
Dim r_dbl_MtoCuo     As Double
Dim r_str_CadErr     As String
Dim r_str_Situac     As String 'Situación del registro (1:Con error, 2:Sin error)
Dim r_int_ConAux     As Integer

   fs_Actualiza_Cronograma_FMVTNC = False
   
   '*****************
   'Actualiza FMV TNC
   r_int_IndCar = 0
   For r_int_Contad = 0 To grd_MViNCo_Listad.Rows - 1
      
      If r_int_IndCar = 0 Then
         'ubica el indicador de carga
         If grd_MViNCo_Listad.TextMatrix(r_int_Contad, 7) = "X" Then
            'setea variables
            r_int_IndCar = 1
            r_int_NumCuo = 0
            r_str_FecVct = Mid(Trim(grd_MViNCo_Listad.TextMatrix(r_int_Contad, 1)), 7, 4) & Mid(Trim(grd_MViNCo_Listad.TextMatrix(r_int_Contad, 1)), 4, 2) & Mid(Trim(grd_MViNCo_Listad.TextMatrix(r_int_Contad, 1)), 1, 2)
            
            'obtiene numero de cuota
            r_str_Cadena = ""
            r_str_Cadena = r_str_Cadena & "SELECT HIPCUO_NUMCUO FROM CRE_HIPCUO "
            r_str_Cadena = r_str_Cadena & " WHERE HIPCUO_NUMOPE = '" & moddat_g_str_NumOpe & "' "
            r_str_Cadena = r_str_Cadena & "   AND HIPCUO_TIPCRO = 3 "
            r_str_Cadena = r_str_Cadena & "   AND HIPCUO_FECVCT = " & r_str_FecVct & " "
            
            If Not gf_EjecutaSQL(r_str_Cadena, r_rst_Cuotas, 3) Then
               MsgBox "No se pudo obtener la cuota a partir de la cual se reemplazara el cronograma FMV TNC.", vbExclamation, modgen_g_str_NomPlt
               Exit Function
            End If
            
            If Not (r_rst_Cuotas.BOF And r_rst_Cuotas.EOF) Then
               r_rst_Cuotas.MoveFirst
               r_int_NumCuo = r_rst_Cuotas!HIPCUO_NUMCUO
            End If
            
            r_rst_Cuotas.Close
            Set r_rst_Cuotas = Nothing
            
            If r_int_NumCuo = 0 Then
               If r_int_Contad = 0 Then
                  r_int_NumCuo = 1
               Else
                  MsgBox "Error, cuota no puede ser cero. Cronograma FMV TNC.", vbExclamation, modgen_g_str_NomPlt
                  Exit Function
               End If
            End If
            
            'elimina cuotas a reemplazar de la BD
            r_str_Cadena = ""
            r_str_Cadena = r_str_Cadena & "DELETE FROM CRE_HIPCUO "
            r_str_Cadena = r_str_Cadena & " WHERE HIPCUO_NUMOPE = '" & moddat_g_str_NumOpe & "' "
            r_str_Cadena = r_str_Cadena & "   AND HIPCUO_TIPCRO = 3 "
            r_str_Cadena = r_str_Cadena & "   AND HIPCUO_NUMCUO >= " & r_int_NumCuo & " "
            
            If Not gf_EjecutaSQL(r_str_Cadena, r_rst_Cuotas, 2) Then
               MsgBox "Error al eliminar las cuotas del cronograma FMV TNC.", vbExclamation, modgen_g_str_NomPlt
               Exit Function
            End If
            
            'carga variables e inserta cuota
            r_str_FecVct = grd_MViNCo_Listad.TextMatrix(r_int_Contad, 1)
            r_dbl_Capita = grd_MViNCo_Listad.TextMatrix(r_int_Contad, 2)
            r_dbl_Intere = grd_MViNCo_Listad.TextMatrix(r_int_Contad, 3)
            r_dbl_ComCof = grd_MViNCo_Listad.TextMatrix(r_int_Contad, 4)
            r_dbl_MtoCuo = grd_MViNCo_Listad.TextMatrix(r_int_Contad, 5)
            r_dbl_SalCap = grd_MViNCo_Listad.TextMatrix(r_int_Contad, 6)
            r_int_ConAux = 1
            
            If Not ff_Inserta_HipCuo(moddat_g_str_NumOpe, 3, r_int_NumCuo, r_str_FecVct, r_dbl_Capita, r_dbl_Intere, 0, 0, 0, r_dbl_SalCap, 0, 0, r_dbl_ComCof) Then
               r_str_Situac = 1
               r_str_CadErr = "No se pudo completar el procedimiento USP_CRE_HIPCUO_CREA."
               Exit For
            End If
            r_str_Situac = 2
            
            '*** Actualiza log
            moddat_g_int_CntErr = 0
            g_str_Parame = "USP_CRE_PROCRODET ("
            g_str_Parame = g_str_Parame & Format(CDate(moddat_g_str_FecSis), "yyyymmdd") & ", "
            g_str_Parame = g_str_Parame & r_int_NumPro & ", "
            g_str_Parame = g_str_Parame & 1 & ", "
            g_str_Parame = g_str_Parame & r_str_Situac & " , "
            g_str_Parame = g_str_Parame & 3 & " , "
            g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "' , "
            g_str_Parame = g_str_Parame & r_int_NumCuo & " , "
            g_str_Parame = g_str_Parame & Format(CDate(r_str_FecVct), "yyyymmdd") & " , "
            g_str_Parame = g_str_Parame & r_dbl_Capita & " , "
            g_str_Parame = g_str_Parame & r_dbl_Intere & " , "
            g_str_Parame = g_str_Parame & 0 & " , "
            g_str_Parame = g_str_Parame & 0 & " , "
            g_str_Parame = g_str_Parame & r_dbl_ComCof & " , "
            g_str_Parame = g_str_Parame & 0 & " , "
            g_str_Parame = g_str_Parame & r_dbl_MtoCuo & " , "
            g_str_Parame = g_str_Parame & r_dbl_SalCap & " , "
            g_str_Parame = g_str_Parame & "'" & r_str_CadErr & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
            g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "' ) "
            
            Do While (moddat_g_int_CntErr = 0)
               If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
                  If MsgBox("No se pudo completar el procedimiento. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
                     moddat_g_int_CntErr = 1
                  Else
                     moddat_g_int_CntErr = 0
                  End If
               Else
                  moddat_g_int_CntErr = 1
               End If
            Loop
            
         End If
      Else
         'carga variables e inserta cuota
         r_str_FecVct = grd_MViNCo_Listad.TextMatrix(r_int_Contad, 1)
         r_dbl_Capita = grd_MViNCo_Listad.TextMatrix(r_int_Contad, 2)
         r_dbl_Intere = grd_MViNCo_Listad.TextMatrix(r_int_Contad, 3)
         r_dbl_ComCof = grd_MViNCo_Listad.TextMatrix(r_int_Contad, 4)
         r_dbl_MtoCuo = grd_MViNCo_Listad.TextMatrix(r_int_Contad, 5)
         r_dbl_SalCap = grd_MViNCo_Listad.TextMatrix(r_int_Contad, 6)
         r_int_NumCuo = r_int_NumCuo + 1
         
         If Not ff_Inserta_HipCuo(moddat_g_str_NumOpe, 3, r_int_NumCuo, r_str_FecVct, r_dbl_Capita, r_dbl_Intere, 0, 0, 0, r_dbl_SalCap, 0, 0, r_dbl_ComCof) Then
            r_str_Situac = 1
            r_str_CadErr = "No se pudo completar el procedimiento USP_CRE_HIPCUO_CREA."
            Exit For
         End If
         r_str_Situac = 2
         '*** Actualiza log
         
            moddat_g_int_CntErr = 0
            r_int_ConAux = r_int_ConAux + 1
            g_str_Parame = "USP_CRE_PROCRODET ("
            g_str_Parame = g_str_Parame & Format(CDate(moddat_g_str_FecSis), "yyyymmdd") & ", "
            g_str_Parame = g_str_Parame & r_int_NumPro & ", "
            g_str_Parame = g_str_Parame & (r_int_ConAux) & ", "
            g_str_Parame = g_str_Parame & r_str_Situac & " , "
            g_str_Parame = g_str_Parame & 3 & " , "
            g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "' , "
            g_str_Parame = g_str_Parame & r_int_NumCuo & " , "
            g_str_Parame = g_str_Parame & Format(CDate(r_str_FecVct), "yyyymmdd") & " , "
            g_str_Parame = g_str_Parame & r_dbl_Capita & " , "
            g_str_Parame = g_str_Parame & r_dbl_Intere & " , "
            g_str_Parame = g_str_Parame & 0 & " , "
            g_str_Parame = g_str_Parame & 0 & " , "
            g_str_Parame = g_str_Parame & r_dbl_ComCof & " , "
            g_str_Parame = g_str_Parame & 0 & " , "
            g_str_Parame = g_str_Parame & r_dbl_MtoCuo & " , "
            g_str_Parame = g_str_Parame & r_dbl_SalCap & " , "
            g_str_Parame = g_str_Parame & "'" & r_str_CadErr & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
            g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "' ) "
            
            Do While (moddat_g_int_CntErr = 0)
               If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
                  If MsgBox("No se pudo completar el procedimiento. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
                     moddat_g_int_CntErr = 1
                  Else
                     moddat_g_int_CntErr = 0
                  End If
               Else
                  moddat_g_int_CntErr = 1
               End If
            Loop
            
      End If
   Next r_int_Contad
   
   If r_str_Situac = 1 Then
      '*** Actualiza log
            moddat_g_int_CntErr = 0
            g_str_Parame = "USP_CRE_PROCRODET ("
            g_str_Parame = g_str_Parame & Format(CDate(moddat_g_str_FecSis), "yyyymmdd") & ", "
            g_str_Parame = g_str_Parame & r_int_NumPro & ", "
            g_str_Parame = g_str_Parame & (r_int_ConAux + 1) & ", "
            g_str_Parame = g_str_Parame & r_str_Situac & " , "
            g_str_Parame = g_str_Parame & 3 & " , "
            g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "' , "
            g_str_Parame = g_str_Parame & r_int_NumCuo & " , "
            g_str_Parame = g_str_Parame & Format(CDate(r_str_FecVct), "yyyymmdd") & " , "
            g_str_Parame = g_str_Parame & r_dbl_Capita & " , "
            g_str_Parame = g_str_Parame & r_dbl_Intere & " , "
            g_str_Parame = g_str_Parame & 0 & " , "
            g_str_Parame = g_str_Parame & 0 & " , "
            g_str_Parame = g_str_Parame & r_dbl_ComCof & " , "
            g_str_Parame = g_str_Parame & 0 & " , "
            g_str_Parame = g_str_Parame & r_dbl_MtoCuo & " , "
            g_str_Parame = g_str_Parame & r_dbl_SalCap & " , "
            g_str_Parame = g_str_Parame & "'" & r_str_CadErr & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
            g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "' ) "
            
            Do While (moddat_g_int_CntErr = 0)
               If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
                  If MsgBox("No se pudo completar el procedimiento. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
                     moddat_g_int_CntErr = 1
                  Else
                     moddat_g_int_CntErr = 0
                  End If
               Else
                  moddat_g_int_CntErr = 1
               End If
            Loop
   End If
   fs_Actualiza_Cronograma_FMVTNC = True
End Function

Private Function fs_Actualiza_Cronograma_FMVTC(ByVal r_int_NumPro As Integer) As Boolean
Dim r_int_Contad     As Integer
Dim r_int_IndCar     As Integer
Dim r_str_Cadena     As String
Dim r_rst_Cuotas     As ADODB.Recordset
Dim r_int_NumCuo     As Integer
Dim r_str_FecVct     As String
Dim r_dbl_Capita     As Double
Dim r_dbl_Intere     As Double
Dim r_dbl_SalCap     As Double
Dim r_dbl_ComCof     As Double
Dim r_dbl_MtoCuo     As Double
Dim r_str_CadErr     As String
Dim r_str_Situac     As String 'Situación del registro (1:Con error, 2:Sin error)
Dim r_int_ConAux     As Integer

   fs_Actualiza_Cronograma_FMVTC = False
   
   '*****************
   'Actualiza FMV TNC
   r_int_IndCar = 0
   r_str_Situac = 0
   For r_int_Contad = 0 To grd_MViCon_Listad.Rows - 1
      
      If r_int_IndCar = 0 Then
         'ubica el indicador de carga
         If grd_MViCon_Listad.TextMatrix(r_int_Contad, 7) = "X" Then
            'setea variables
            r_int_IndCar = 1
            r_int_NumCuo = 0
            r_str_FecVct = Mid(Trim(grd_MViCon_Listad.TextMatrix(r_int_Contad, 1)), 7, 4) & Mid(Trim(grd_MViCon_Listad.TextMatrix(r_int_Contad, 1)), 4, 2) & Mid(Trim(grd_MViCon_Listad.TextMatrix(r_int_Contad, 1)), 1, 2)
            
            'obtiene numero de cuota
            r_str_Cadena = ""
            r_str_Cadena = r_str_Cadena & "SELECT HIPCUO_NUMCUO FROM CRE_HIPCUO "
            r_str_Cadena = r_str_Cadena & " WHERE HIPCUO_NUMOPE = '" & moddat_g_str_NumOpe & "' "
            r_str_Cadena = r_str_Cadena & "   AND HIPCUO_TIPCRO = 4 "
            r_str_Cadena = r_str_Cadena & "   AND HIPCUO_FECVCT = " & r_str_FecVct & " "
            
            If Not gf_EjecutaSQL(r_str_Cadena, r_rst_Cuotas, 3) Then
               MsgBox "No se pudo obtener la cuota a partir de la cual se reemplazara el cronograma FMV TC.", vbExclamation, modgen_g_str_NomPlt
               Exit Function
            End If
            
            If Not (r_rst_Cuotas.BOF And r_rst_Cuotas.EOF) Then
               r_rst_Cuotas.MoveFirst
               r_int_NumCuo = r_rst_Cuotas!HIPCUO_NUMCUO
            End If
            
            r_rst_Cuotas.Close
            Set r_rst_Cuotas = Nothing
            
            If r_int_NumCuo = 0 Then
               MsgBox "Error, cuota no puede ser cero. Cronograma FMV TC.", vbExclamation, modgen_g_str_NomPlt
               Exit Function
            End If
            
            'elimina cuotas a reemplazar de la BD
            r_str_Cadena = ""
            r_str_Cadena = r_str_Cadena & "DELETE FROM CRE_HIPCUO "
            r_str_Cadena = r_str_Cadena & " WHERE HIPCUO_NUMOPE = '" & moddat_g_str_NumOpe & "' "
            r_str_Cadena = r_str_Cadena & "   AND HIPCUO_TIPCRO = 4 "
            r_str_Cadena = r_str_Cadena & "   AND HIPCUO_NUMCUO >= " & r_int_NumCuo & " "
            
            If Not gf_EjecutaSQL(r_str_Cadena, r_rst_Cuotas, 2) Then
               MsgBox "Error al eliminar las cuotas del cronograma FMV TC.", vbExclamation, modgen_g_str_NomPlt
               Exit Function
            End If
            
            'carga variables e inserta cuota
            r_str_FecVct = grd_MViCon_Listad.TextMatrix(r_int_Contad, 1)
            r_dbl_Capita = grd_MViCon_Listad.TextMatrix(r_int_Contad, 2)
            r_dbl_Intere = grd_MViCon_Listad.TextMatrix(r_int_Contad, 3)
            r_dbl_ComCof = grd_MViCon_Listad.TextMatrix(r_int_Contad, 4)
            r_dbl_MtoCuo = grd_MViCon_Listad.TextMatrix(r_int_Contad, 5)
            r_dbl_SalCap = grd_MViCon_Listad.TextMatrix(r_int_Contad, 6)
            
            If Not ff_Inserta_HipCuo(moddat_g_str_NumOpe, 4, r_int_NumCuo, r_str_FecVct, r_dbl_Capita, r_dbl_Intere, 0, 0, 0, r_dbl_SalCap, 0, 0, r_dbl_ComCof) Then
               r_str_Situac = 1
               r_str_CadErr = "No se pudo completar el procedimiento USP_CRE_HIPCUO_CREA."
               Exit For
            End If
            r_str_Situac = 2
            '*** Actualiza log
            moddat_g_int_CntErr = 0
            r_int_ConAux = 1
            g_str_Parame = "USP_CRE_PROCRODET ("
            g_str_Parame = g_str_Parame & Format(CDate(moddat_g_str_FecSis), "yyyymmdd") & ", "
            g_str_Parame = g_str_Parame & r_int_NumPro & ", "
            g_str_Parame = g_str_Parame & r_int_ConAux & ", "
            g_str_Parame = g_str_Parame & r_str_Situac & " , "
            g_str_Parame = g_str_Parame & 4 & " , "
            g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "' , "
            g_str_Parame = g_str_Parame & r_int_NumCuo & " , "
            g_str_Parame = g_str_Parame & Format(CDate(r_str_FecVct), "yyyymmdd") & " , "
            g_str_Parame = g_str_Parame & r_dbl_Capita & " , "
            g_str_Parame = g_str_Parame & r_dbl_Intere & " , "
            g_str_Parame = g_str_Parame & 0 & " , "
            g_str_Parame = g_str_Parame & 0 & " , "
            g_str_Parame = g_str_Parame & r_dbl_ComCof & " , "
            g_str_Parame = g_str_Parame & 0 & " , "
            g_str_Parame = g_str_Parame & r_dbl_MtoCuo & " , "
            g_str_Parame = g_str_Parame & r_dbl_SalCap & " , "
            g_str_Parame = g_str_Parame & "'" & r_str_CadErr & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
            g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "' ) "
            
            Do While (moddat_g_int_CntErr = 0)
               If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
                  If MsgBox("No se pudo completar el procedimiento. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
                     moddat_g_int_CntErr = 1
                  Else
                     moddat_g_int_CntErr = 0
                  End If
               Else
                  moddat_g_int_CntErr = 1
               End If
            Loop
         End If
      Else
         'carga variables e inserta cuota
         r_str_FecVct = grd_MViCon_Listad.TextMatrix(r_int_Contad, 1)
         r_dbl_Capita = grd_MViCon_Listad.TextMatrix(r_int_Contad, 2)
         r_dbl_Intere = grd_MViCon_Listad.TextMatrix(r_int_Contad, 3)
         r_dbl_ComCof = grd_MViCon_Listad.TextMatrix(r_int_Contad, 4)
         r_dbl_MtoCuo = grd_MViCon_Listad.TextMatrix(r_int_Contad, 5)
         r_dbl_SalCap = grd_MViCon_Listad.TextMatrix(r_int_Contad, 6)
         r_int_NumCuo = r_int_NumCuo + 1
         
         If Not ff_Inserta_HipCuo(moddat_g_str_NumOpe, 4, r_int_NumCuo, r_str_FecVct, r_dbl_Capita, r_dbl_Intere, 0, 0, 0, r_dbl_SalCap, 0, 0, r_dbl_ComCof) Then
            r_str_Situac = 1
            r_str_CadErr = "No se pudo completar el procedimiento USP_CRE_HIPCUO_CREA."
            Exit For
         End If
         r_str_Situac = 2
         
         '*** Actualiza log
            moddat_g_int_CntErr = 0
            r_int_ConAux = r_int_ConAux + 1
            g_str_Parame = "USP_CRE_PROCRODET ("
            g_str_Parame = g_str_Parame & Format(CDate(moddat_g_str_FecSis), "yyyymmdd") & ", "
            g_str_Parame = g_str_Parame & r_int_NumPro & ", "
            g_str_Parame = g_str_Parame & r_int_ConAux & ", "
            g_str_Parame = g_str_Parame & r_str_Situac & " , "
            g_str_Parame = g_str_Parame & 4 & " , "
            g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "' , "
            g_str_Parame = g_str_Parame & r_int_NumCuo & " , "
            g_str_Parame = g_str_Parame & Format(CDate(r_str_FecVct), "yyyymmdd") & " , "
            g_str_Parame = g_str_Parame & r_dbl_Capita & " , "
            g_str_Parame = g_str_Parame & r_dbl_Intere & " , "
            g_str_Parame = g_str_Parame & 0 & " , "
            g_str_Parame = g_str_Parame & 0 & " , "
            g_str_Parame = g_str_Parame & r_dbl_ComCof & " , "
            g_str_Parame = g_str_Parame & 0 & " , "
            g_str_Parame = g_str_Parame & r_dbl_MtoCuo & " , "
            g_str_Parame = g_str_Parame & r_dbl_SalCap & " , "
            g_str_Parame = g_str_Parame & "'" & r_str_CadErr & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
            g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "' ) "
            
            Do While (moddat_g_int_CntErr = 0)
               If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
                  If MsgBox("No se pudo completar el procedimiento. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
                     moddat_g_int_CntErr = 1
                  Else
                     moddat_g_int_CntErr = 0
                  End If
               Else
                  moddat_g_int_CntErr = 1
               End If
            Loop
      End If
   Next r_int_Contad
   
   'Cuando se produce Error sale del For, y debe capturar este detalle
   If r_str_Situac = 1 Then
      '*** Actualiza log
      moddat_g_int_CntErr = 0
      r_int_ConAux = r_int_ConAux + 1
      g_str_Parame = "USP_CRE_PROCRODET ("
      g_str_Parame = g_str_Parame & Format(CDate(moddat_g_str_FecSis), "yyyymmdd") & ", "
      g_str_Parame = g_str_Parame & r_int_NumPro & ", "
      g_str_Parame = g_str_Parame & r_int_ConAux & ", "
      g_str_Parame = g_str_Parame & r_str_Situac & " , "
      g_str_Parame = g_str_Parame & 4 & " , "
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "' , "
      g_str_Parame = g_str_Parame & r_int_NumCuo & " , "
      g_str_Parame = g_str_Parame & Format(CDate(r_str_FecVct), "yyyymmdd") & " , "
      g_str_Parame = g_str_Parame & r_dbl_Capita & " , "
      g_str_Parame = g_str_Parame & r_dbl_Intere & " , "
      g_str_Parame = g_str_Parame & 0 & " , "
      g_str_Parame = g_str_Parame & 0 & " , "
      g_str_Parame = g_str_Parame & r_dbl_ComCof & " , "
      g_str_Parame = g_str_Parame & 0 & " , "
      g_str_Parame = g_str_Parame & r_dbl_MtoCuo & " , "
      g_str_Parame = g_str_Parame & r_dbl_SalCap & " , "
      g_str_Parame = g_str_Parame & "'" & r_str_CadErr & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "' ) "
      
      Do While (moddat_g_int_CntErr = 0)
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
            If MsgBox("No se pudo completar el procedimiento. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
               moddat_g_int_CntErr = 1
            Else
               moddat_g_int_CntErr = 0
            End If
         Else
            moddat_g_int_CntErr = 1
         End If
      Loop
   End If
   
   fs_Actualiza_Cronograma_FMVTC = True
End Function

Private Function fs_Actualiza_Cronograma_CME(ByVal r_int_NumPro As Integer) As Boolean
Dim r_int_Contad     As Integer
Dim r_int_IndCar     As Integer
Dim r_str_Cadena     As String
Dim r_rst_Cuotas     As ADODB.Recordset
Dim r_int_NumCuo     As Integer
Dim r_str_FecVct     As String
Dim r_dbl_Capita     As Double
Dim r_dbl_Intere     As Double
Dim r_dbl_SalCap     As Double
Dim r_dbl_ComCof     As Double
Dim r_dbl_MtoCuo     As Double
Dim r_str_CadErr     As String
Dim r_str_Situac     As String 'Situación del registro (1:Con error, 2:Sin error)
Dim r_int_ConAux     As Integer

   fs_Actualiza_Cronograma_CME = False
   
   '*****************
   'Actualiza FMV TNC
   r_int_IndCar = 0
   For r_int_Contad = 0 To grd_MViNCo_Listad.Rows - 1
      
      If r_int_IndCar = 0 Then
         'ubica el indicador de carga
         If grd_MViNCo_Listad.TextMatrix(r_int_Contad, 7) = "X" Then
            'setea variables
            r_int_IndCar = 1
            r_int_NumCuo = 0
            r_str_FecVct = Mid(Trim(grd_MViNCo_Listad.TextMatrix(r_int_Contad, 1)), 7, 4) & Mid(Trim(grd_MViNCo_Listad.TextMatrix(r_int_Contad, 1)), 4, 2) & Mid(Trim(grd_MViNCo_Listad.TextMatrix(r_int_Contad, 1)), 1, 2)
            
            'obtiene numero de cuota
            r_str_Cadena = ""
            r_str_Cadena = r_str_Cadena & "SELECT HIPCUO_NUMCUO FROM CRE_HIPCUO "
            r_str_Cadena = r_str_Cadena & " WHERE HIPCUO_NUMOPE = '" & moddat_g_str_NumOpe & "' "
            r_str_Cadena = r_str_Cadena & "   AND HIPCUO_TIPCRO = 5 "
            r_str_Cadena = r_str_Cadena & "   AND HIPCUO_FECVCT = " & r_str_FecVct & " "
            
            If Not gf_EjecutaSQL(r_str_Cadena, r_rst_Cuotas, 3) Then
               MsgBox "No se pudo obtener la cuota a partir de la cual se reemplazara el cronograma FMV TNC.", vbExclamation, modgen_g_str_NomPlt
               Exit Function
            End If
            
            If Not (r_rst_Cuotas.BOF And r_rst_Cuotas.EOF) Then
               r_rst_Cuotas.MoveFirst
               r_int_NumCuo = r_rst_Cuotas!HIPCUO_NUMCUO
            End If
            
            r_rst_Cuotas.Close
            Set r_rst_Cuotas = Nothing
            
            If r_int_NumCuo = 0 Then
               If r_int_Contad = 0 Then
                  r_int_NumCuo = 1
               Else
                  MsgBox "Error, cuota no puede ser cero. Cronograma CME TNC.", vbExclamation, modgen_g_str_NomPlt
                  Exit Function
               End If
            End If
            
            'elimina cuotas a reemplazar de la BD
            r_str_Cadena = ""
            r_str_Cadena = r_str_Cadena & "DELETE FROM CRE_HIPCUO "
            r_str_Cadena = r_str_Cadena & " WHERE HIPCUO_NUMOPE = '" & moddat_g_str_NumOpe & "' "
            r_str_Cadena = r_str_Cadena & "   AND HIPCUO_TIPCRO = 5 "
            r_str_Cadena = r_str_Cadena & "   AND HIPCUO_NUMCUO >= " & r_int_NumCuo & " "
            
            If Not gf_EjecutaSQL(r_str_Cadena, r_rst_Cuotas, 2) Then
               MsgBox "Error al eliminar las cuotas del cronograma CME TNC.", vbExclamation, modgen_g_str_NomPlt
               Exit Function
            End If
            
            'carga variables e inserta cuota
            r_str_FecVct = grd_MViNCo_Listad.TextMatrix(r_int_Contad, 1)
            r_dbl_Capita = grd_MViNCo_Listad.TextMatrix(r_int_Contad, 2)
            r_dbl_Intere = grd_MViNCo_Listad.TextMatrix(r_int_Contad, 3)
            r_dbl_ComCof = grd_MViNCo_Listad.TextMatrix(r_int_Contad, 4)
            r_dbl_MtoCuo = grd_MViNCo_Listad.TextMatrix(r_int_Contad, 5)
            r_dbl_SalCap = grd_MViNCo_Listad.TextMatrix(r_int_Contad, 6)
            r_int_ConAux = 1
            
            If Not ff_Inserta_HipCuo(moddat_g_str_NumOpe, 5, r_int_NumCuo, r_str_FecVct, r_dbl_Capita, r_dbl_Intere, 0, 0, 0, r_dbl_SalCap, 0, 0, r_dbl_ComCof) Then
               r_str_Situac = 1
               r_str_CadErr = "No se pudo completar el procedimiento USP_CRE_HIPCUO_CREA."
               Exit For
            End If
            r_str_Situac = 2
            
            '*** Actualiza log
            moddat_g_int_CntErr = 0
            g_str_Parame = "USP_CRE_PROCRODET ("
            g_str_Parame = g_str_Parame & Format(CDate(moddat_g_str_FecSis), "yyyymmdd") & ", "
            g_str_Parame = g_str_Parame & r_int_NumPro & ", "
            g_str_Parame = g_str_Parame & 1 & ", "
            g_str_Parame = g_str_Parame & r_str_Situac & " , "
            g_str_Parame = g_str_Parame & 5 & " , "
            g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "' , "
            g_str_Parame = g_str_Parame & r_int_NumCuo & " , "
            g_str_Parame = g_str_Parame & Format(CDate(r_str_FecVct), "yyyymmdd") & " , "
            g_str_Parame = g_str_Parame & r_dbl_Capita & " , "
            g_str_Parame = g_str_Parame & r_dbl_Intere & " , "
            g_str_Parame = g_str_Parame & 0 & " , "
            g_str_Parame = g_str_Parame & 0 & " , "
            g_str_Parame = g_str_Parame & r_dbl_ComCof & " , "
            g_str_Parame = g_str_Parame & 0 & " , "
            g_str_Parame = g_str_Parame & r_dbl_MtoCuo & " , "
            g_str_Parame = g_str_Parame & r_dbl_SalCap & " , "
            g_str_Parame = g_str_Parame & "'" & r_str_CadErr & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
            g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "' ) "
            
            Do While (moddat_g_int_CntErr = 0)
               If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
                  If MsgBox("No se pudo completar el procedimiento. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
                     moddat_g_int_CntErr = 1
                  Else
                     moddat_g_int_CntErr = 0
                  End If
               Else
                  moddat_g_int_CntErr = 1
               End If
            Loop
            
         End If
      Else
         'carga variables e inserta cuota
         r_str_FecVct = grd_MViNCo_Listad.TextMatrix(r_int_Contad, 1)
         r_dbl_Capita = grd_MViNCo_Listad.TextMatrix(r_int_Contad, 2)
         r_dbl_Intere = grd_MViNCo_Listad.TextMatrix(r_int_Contad, 3)
         r_dbl_ComCof = grd_MViNCo_Listad.TextMatrix(r_int_Contad, 4)
         r_dbl_MtoCuo = grd_MViNCo_Listad.TextMatrix(r_int_Contad, 5)
         r_dbl_SalCap = grd_MViNCo_Listad.TextMatrix(r_int_Contad, 6)
         r_int_NumCuo = r_int_NumCuo + 1
         
         If Not ff_Inserta_HipCuo(moddat_g_str_NumOpe, 5, r_int_NumCuo, r_str_FecVct, r_dbl_Capita, r_dbl_Intere, 0, 0, 0, r_dbl_SalCap, 0, 0, r_dbl_ComCof) Then
            r_str_Situac = 1
            r_str_CadErr = "No se pudo completar el procedimiento USP_CRE_HIPCUO_CREA."
            Exit For
         End If
         r_str_Situac = 2
         '*** Actualiza log
         
            moddat_g_int_CntErr = 0
            r_int_ConAux = r_int_ConAux + 1
            g_str_Parame = "USP_CRE_PROCRODET ("
            g_str_Parame = g_str_Parame & Format(CDate(moddat_g_str_FecSis), "yyyymmdd") & ", "
            g_str_Parame = g_str_Parame & r_int_NumPro & ", "
            g_str_Parame = g_str_Parame & (r_int_ConAux) & ", "
            g_str_Parame = g_str_Parame & r_str_Situac & " , "
            g_str_Parame = g_str_Parame & 5 & " , "
            g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "' , "
            g_str_Parame = g_str_Parame & r_int_NumCuo & " , "
            g_str_Parame = g_str_Parame & Format(CDate(r_str_FecVct), "yyyymmdd") & " , "
            g_str_Parame = g_str_Parame & r_dbl_Capita & " , "
            g_str_Parame = g_str_Parame & r_dbl_Intere & " , "
            g_str_Parame = g_str_Parame & 0 & " , "
            g_str_Parame = g_str_Parame & 0 & " , "
            g_str_Parame = g_str_Parame & r_dbl_ComCof & " , "
            g_str_Parame = g_str_Parame & 0 & " , "
            g_str_Parame = g_str_Parame & r_dbl_MtoCuo & " , "
            g_str_Parame = g_str_Parame & r_dbl_SalCap & " , "
            g_str_Parame = g_str_Parame & "'" & r_str_CadErr & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
            g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "' ) "
            
            Do While (moddat_g_int_CntErr = 0)
               If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
                  If MsgBox("No se pudo completar el procedimiento. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
                     moddat_g_int_CntErr = 1
                  Else
                     moddat_g_int_CntErr = 0
                  End If
               Else
                  moddat_g_int_CntErr = 1
               End If
            Loop
            
      End If
   Next r_int_Contad
   
   If r_str_Situac = 1 Then
      '*** Actualiza log
            moddat_g_int_CntErr = 0
            g_str_Parame = "USP_CRE_PROCRODET ("
            g_str_Parame = g_str_Parame & Format(CDate(moddat_g_str_FecSis), "yyyymmdd") & ", "
            g_str_Parame = g_str_Parame & r_int_NumPro & ", "
            g_str_Parame = g_str_Parame & (r_int_ConAux + 1) & ", "
            g_str_Parame = g_str_Parame & r_str_Situac & " , "
            g_str_Parame = g_str_Parame & 5 & " , "
            g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "' , "
            g_str_Parame = g_str_Parame & r_int_NumCuo & " , "
            g_str_Parame = g_str_Parame & Format(CDate(r_str_FecVct), "yyyymmdd") & " , "
            g_str_Parame = g_str_Parame & r_dbl_Capita & " , "
            g_str_Parame = g_str_Parame & r_dbl_Intere & " , "
            g_str_Parame = g_str_Parame & 0 & " , "
            g_str_Parame = g_str_Parame & 0 & " , "
            g_str_Parame = g_str_Parame & r_dbl_ComCof & " , "
            g_str_Parame = g_str_Parame & 0 & " , "
            g_str_Parame = g_str_Parame & r_dbl_MtoCuo & " , "
            g_str_Parame = g_str_Parame & r_dbl_SalCap & " , "
            g_str_Parame = g_str_Parame & "'" & r_str_CadErr & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
            g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "' ) "
            
            Do While (moddat_g_int_CntErr = 0)
               If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
                  If MsgBox("No se pudo completar el procedimiento. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
                     moddat_g_int_CntErr = 1
                  Else
                     moddat_g_int_CntErr = 0
                  End If
               Else
                  moddat_g_int_CntErr = 1
               End If
            Loop
   End If
   fs_Actualiza_Cronograma_CME = True
End Function

Private Function fs_Actualiza_Cronograma_CLITC(ByVal r_int_NumPro As Integer) As Boolean
Dim r_int_Contad     As Integer
Dim r_int_IndCar     As Integer
Dim r_str_Cadena     As String
Dim r_rst_Cuotas     As ADODB.Recordset
Dim r_int_NumCuo     As Integer
Dim r_str_FecVct     As String
Dim r_dbl_Capita     As Double
Dim r_dbl_Intere     As Double
Dim r_dbl_SalCap     As Double
Dim r_dbl_MtoCuo     As Double
Dim r_str_CadErr     As String
Dim r_str_Situac     As String 'Situación del registro (1:Con error, 2:Sin error)
Dim r_int_ConAux     As Integer

   fs_Actualiza_Cronograma_CLITC = True
   
   '*****************
   'Actualiza FMV TNC
   r_int_IndCar = 0
   r_str_Situac = 0
   For r_int_Contad = 0 To grd_CliCon_Listad.Rows - 1
      
      If r_int_IndCar = 0 Then
         'ubica el indicador de carga
         If grd_CliCon_Listad.TextMatrix(r_int_Contad, 6) = "X" Then
            'setea variables
            r_int_IndCar = 1
            r_int_NumCuo = 0
            r_str_FecVct = Mid(Trim(grd_CliCon_Listad.TextMatrix(r_int_Contad, 1)), 7, 4) & Mid(Trim(grd_CliCon_Listad.TextMatrix(r_int_Contad, 1)), 4, 2) & Mid(Trim(grd_CliCon_Listad.TextMatrix(r_int_Contad, 1)), 1, 2)
            
            'obtiene numero de cuota
            r_str_Cadena = ""
            r_str_Cadena = r_str_Cadena & "SELECT HIPCUO_NUMCUO FROM CRE_HIPCUO "
            r_str_Cadena = r_str_Cadena & " WHERE HIPCUO_NUMOPE = '" & moddat_g_str_NumOpe & "' "
            r_str_Cadena = r_str_Cadena & "   AND HIPCUO_TIPCRO = 2 "
            r_str_Cadena = r_str_Cadena & "   AND HIPCUO_FECVCT = " & r_str_FecVct & " "
            
            If Not gf_EjecutaSQL(r_str_Cadena, r_rst_Cuotas, 3) Then
               MsgBox "No se pudo obtener la cuota a partir de la cual se reemplazara el cronograma CLI TC.", vbExclamation, modgen_g_str_NomPlt
               Exit Function
            End If
            
            If Not (r_rst_Cuotas.BOF And r_rst_Cuotas.EOF) Then
               r_rst_Cuotas.MoveFirst
               r_int_NumCuo = r_rst_Cuotas!HIPCUO_NUMCUO
            End If
            
            r_rst_Cuotas.Close
            Set r_rst_Cuotas = Nothing
            
            If r_int_NumCuo = 0 Then
               MsgBox "Error, cuota no puede ser cero. Cronograma CLI TC.", vbExclamation, modgen_g_str_NomPlt
               Exit Function
            End If
            
            'elimina cuotas a reemplazar de la BD
            r_str_Cadena = ""
            r_str_Cadena = r_str_Cadena & "DELETE FROM CRE_HIPCUO "
            r_str_Cadena = r_str_Cadena & " WHERE HIPCUO_NUMOPE = '" & moddat_g_str_NumOpe & "' "
            r_str_Cadena = r_str_Cadena & "   AND HIPCUO_TIPCRO = 2 "
            r_str_Cadena = r_str_Cadena & "   AND HIPCUO_NUMCUO >= " & r_int_NumCuo & " "
            
            If Not gf_EjecutaSQL(r_str_Cadena, r_rst_Cuotas, 2) Then
               MsgBox "Error al eliminar las cuotas del cronograma CLI TC.", vbExclamation, modgen_g_str_NomPlt
               Exit Function
            End If
            
            'carga variables e inserta cuota
            r_str_FecVct = grd_CliCon_Listad.TextMatrix(r_int_Contad, 1)
            r_dbl_Capita = grd_CliCon_Listad.TextMatrix(r_int_Contad, 2)
            r_dbl_Intere = grd_CliCon_Listad.TextMatrix(r_int_Contad, 3)
            r_dbl_MtoCuo = grd_CliCon_Listad.TextMatrix(r_int_Contad, 4)
            r_dbl_SalCap = grd_CliCon_Listad.TextMatrix(r_int_Contad, 5)
            
            If Not ff_Inserta_HipCuo(moddat_g_str_NumOpe, 2, r_int_NumCuo, r_str_FecVct, r_dbl_Capita, r_dbl_Intere, 0, 0, 0, r_dbl_SalCap, 0, 0, 0) Then
               r_str_Situac = 1
               r_str_CadErr = "No se pudo completar el procedimiento USP_CRE_HIPCUO_CREA."
               Exit For
            End If
            r_str_Situac = 2
            '*** Actualiza log
            moddat_g_int_CntErr = 0
            r_int_ConAux = 1
            g_str_Parame = "USP_CRE_PROCRODET ("
            g_str_Parame = g_str_Parame & Format(CDate(moddat_g_str_FecSis), "yyyymmdd") & ", "
            g_str_Parame = g_str_Parame & r_int_NumPro & ", "
            g_str_Parame = g_str_Parame & r_int_ConAux & ", "
            g_str_Parame = g_str_Parame & r_str_Situac & " , "
            g_str_Parame = g_str_Parame & 2 & " , "
            g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "' , "
            g_str_Parame = g_str_Parame & r_int_NumCuo & " , "
            g_str_Parame = g_str_Parame & Format(CDate(r_str_FecVct), "yyyymmdd") & " , "
            g_str_Parame = g_str_Parame & r_dbl_Capita & " , "
            g_str_Parame = g_str_Parame & r_dbl_Intere & " , "
            g_str_Parame = g_str_Parame & 0 & " , "
            g_str_Parame = g_str_Parame & 0 & " , "
            g_str_Parame = g_str_Parame & 0 & " , "
            g_str_Parame = g_str_Parame & 0 & " , "
            g_str_Parame = g_str_Parame & r_dbl_MtoCuo & " , "
            g_str_Parame = g_str_Parame & r_dbl_SalCap & " , "
            g_str_Parame = g_str_Parame & "'" & r_str_CadErr & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
            g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "' ) "
            
            Do While (moddat_g_int_CntErr = 0)
               If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
                  If MsgBox("No se pudo completar el procedimiento. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
                     moddat_g_int_CntErr = 1
                  Else
                     moddat_g_int_CntErr = 0
                  End If
               Else
                  moddat_g_int_CntErr = 1
               End If
            Loop
         End If
      Else
         'carga variables e inserta cuota
         r_str_FecVct = grd_CliCon_Listad.TextMatrix(r_int_Contad, 1)
         r_dbl_Capita = grd_CliCon_Listad.TextMatrix(r_int_Contad, 2)
         r_dbl_Intere = grd_CliCon_Listad.TextMatrix(r_int_Contad, 3)
         r_dbl_MtoCuo = grd_CliCon_Listad.TextMatrix(r_int_Contad, 4)
         r_dbl_SalCap = grd_CliCon_Listad.TextMatrix(r_int_Contad, 5)
         r_int_NumCuo = r_int_NumCuo + 1
         
         If Not ff_Inserta_HipCuo(moddat_g_str_NumOpe, 2, r_int_NumCuo, r_str_FecVct, r_dbl_Capita, r_dbl_Intere, 0, 0, 0, r_dbl_SalCap, 0, 0, 0) Then
            r_str_Situac = 1
            r_str_CadErr = "No se pudo completar el procedimiento USP_CRE_HIPCUO_CREA."
            Exit For
         End If
         r_str_Situac = 2
         '*** Actualiza log
         moddat_g_int_CntErr = 0
         r_int_ConAux = r_int_ConAux + 1
         g_str_Parame = "USP_CRE_PROCRODET ("
         g_str_Parame = g_str_Parame & Format(CDate(moddat_g_str_FecSis), "yyyymmdd") & ", "
         g_str_Parame = g_str_Parame & r_int_NumPro & ", "
         g_str_Parame = g_str_Parame & r_int_ConAux & ", "
         g_str_Parame = g_str_Parame & r_str_Situac & " , "
         g_str_Parame = g_str_Parame & 2 & " , "
         g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "' , "
         g_str_Parame = g_str_Parame & r_int_NumCuo & " , "
         g_str_Parame = g_str_Parame & Format(CDate(r_str_FecVct), "yyyymmdd") & " , "
         g_str_Parame = g_str_Parame & r_dbl_Capita & " , "
         g_str_Parame = g_str_Parame & r_dbl_Intere & " , "
         g_str_Parame = g_str_Parame & 0 & " , "
         g_str_Parame = g_str_Parame & 0 & " , "
         g_str_Parame = g_str_Parame & 0 & " , "
         g_str_Parame = g_str_Parame & 0 & " , "
         g_str_Parame = g_str_Parame & r_dbl_MtoCuo & " , "
         g_str_Parame = g_str_Parame & r_dbl_SalCap & " , "
         g_str_Parame = g_str_Parame & "'" & r_str_CadErr & "', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
         g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "' ) "
         
         Do While (moddat_g_int_CntErr = 0)
            If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
               If MsgBox("No se pudo completar el procedimiento. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
                  moddat_g_int_CntErr = 1
               Else
                  moddat_g_int_CntErr = 0
               End If
            Else
               moddat_g_int_CntErr = 1
            End If
         Loop
      End If
   Next r_int_Contad
   
   If r_str_Situac = 1 Then
      '*** Actualiza log
      moddat_g_int_CntErr = 0
      r_int_ConAux = r_int_ConAux + 1
      g_str_Parame = "USP_CRE_PROCRODET ("
      g_str_Parame = g_str_Parame & Format(CDate(moddat_g_str_FecSis), "yyyymmdd") & ", "
      g_str_Parame = g_str_Parame & r_int_NumPro & ", "
      g_str_Parame = g_str_Parame & r_int_ConAux & ", "
      g_str_Parame = g_str_Parame & r_str_Situac & " , "
      g_str_Parame = g_str_Parame & 2 & " , "
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "' , "
      g_str_Parame = g_str_Parame & r_int_NumCuo & " , "
      g_str_Parame = g_str_Parame & Format(CDate(r_str_FecVct), "yyyymmdd") & " , "
      g_str_Parame = g_str_Parame & r_dbl_Capita & " , "
      g_str_Parame = g_str_Parame & r_dbl_Intere & " , "
      g_str_Parame = g_str_Parame & 0 & " , "
      g_str_Parame = g_str_Parame & 0 & " , "
      g_str_Parame = g_str_Parame & 0 & " , "
      g_str_Parame = g_str_Parame & 0 & " , "
      g_str_Parame = g_str_Parame & r_dbl_MtoCuo & " , "
      g_str_Parame = g_str_Parame & r_dbl_SalCap & " , "
      g_str_Parame = g_str_Parame & "'" & r_str_CadErr & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "' ) "
      
      Do While (moddat_g_int_CntErr = 0)
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
            If MsgBox("No se pudo completar el procedimiento. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
               moddat_g_int_CntErr = 1
            Else
               moddat_g_int_CntErr = 0
            End If
         Else
            moddat_g_int_CntErr = 1
         End If
      Loop
   End If
   fs_Actualiza_Cronograma_CLITC = True
End Function

Private Function fs_Actualiza_Cronograma_CLITNC(ByVal r_int_NumPro As Integer) As Boolean
Dim r_int_Contad     As Integer
Dim r_int_IndCar     As Integer
Dim r_str_Cadena     As String
Dim r_rst_Cuotas     As ADODB.Recordset
Dim r_int_NumCuo     As Integer
Dim r_str_FecVct     As String
Dim r_dbl_Capita     As Double
Dim r_dbl_Intere     As Double
Dim r_dbl_Portes     As Double
Dim r_dbl_SegPre     As Double
Dim r_dbl_SegViv     As Double
Dim r_dbl_MtoCuo     As Double
Dim r_dbl_SalCap     As Double
Dim r_str_CadErr     As String
Dim r_str_Situac     As String 'Situación del registro (1:Con error, 2:Sin error)
Dim r_int_ConAux     As Integer

   fs_Actualiza_Cronograma_CLITNC = True
   
   '*****************
   'Actualiza CLIENTE TNC
   r_int_IndCar = 0
   For r_int_Contad = 0 To grd_CliNCon_Listad.Rows - 1
      
      If r_int_IndCar = 0 Then
         'ubica el indicador de carga
         If grd_CliNCon_Listad.TextMatrix(r_int_Contad, 9) = "X" Then
            'setea variables
            r_int_IndCar = 1
            r_int_NumCuo = 0
            r_str_FecVct = Mid(Trim(grd_CliNCon_Listad.TextMatrix(r_int_Contad, 1)), 7, 4) & Mid(Trim(grd_CliNCon_Listad.TextMatrix(r_int_Contad, 1)), 4, 2) & Mid(Trim(grd_CliNCon_Listad.TextMatrix(r_int_Contad, 1)), 1, 2)
            
            'obtiene numero de cuota
            r_str_Cadena = ""
            r_str_Cadena = r_str_Cadena & "SELECT HIPCUO_NUMCUO FROM CRE_HIPCUO "
            r_str_Cadena = r_str_Cadena & " WHERE HIPCUO_NUMOPE = '" & moddat_g_str_NumOpe & "' "
            r_str_Cadena = r_str_Cadena & "   AND HIPCUO_TIPCRO = 1 "
            r_str_Cadena = r_str_Cadena & "   AND HIPCUO_FECVCT = " & r_str_FecVct & " "
            
            If Not gf_EjecutaSQL(r_str_Cadena, r_rst_Cuotas, 3) Then
               MsgBox "No se pudo obtener la cuota a partir de la cual se reemplazara el cronograma CLI TC.", vbExclamation, modgen_g_str_NomPlt
               Exit Function
            End If
            
            If Not (r_rst_Cuotas.BOF And r_rst_Cuotas.EOF) Then
               r_rst_Cuotas.MoveFirst
               r_int_NumCuo = r_rst_Cuotas!HIPCUO_NUMCUO
            End If
            
            r_rst_Cuotas.Close
            Set r_rst_Cuotas = Nothing
            
            If r_int_NumCuo = 0 Then
               MsgBox "Error, cuota no puede ser cero. Cronograma CLI TNC.", vbExclamation, modgen_g_str_NomPlt
               Exit Function
            End If
            
            'elimina cuotas a reemplazar de la BD
            r_str_Cadena = ""
            r_str_Cadena = r_str_Cadena & "DELETE FROM CRE_HIPCUO "
            r_str_Cadena = r_str_Cadena & " WHERE HIPCUO_NUMOPE = '" & moddat_g_str_NumOpe & "' "
            r_str_Cadena = r_str_Cadena & "   AND HIPCUO_TIPCRO = 1 "
            r_str_Cadena = r_str_Cadena & "   AND HIPCUO_NUMCUO >= " & r_int_NumCuo & " "
            
            If Not gf_EjecutaSQL(r_str_Cadena, r_rst_Cuotas, 2) Then
               MsgBox "Error al eliminar las cuotas del cronograma CLI TC.", vbExclamation, modgen_g_str_NomPlt
               Exit Function
            End If
            
            'carga variables e inserta cuota
            r_str_FecVct = grd_CliNCon_Listad.TextMatrix(r_int_Contad, 1)
            r_dbl_Capita = grd_CliNCon_Listad.TextMatrix(r_int_Contad, 2)
            r_dbl_Intere = grd_CliNCon_Listad.TextMatrix(r_int_Contad, 3)
            r_dbl_SegPre = grd_CliNCon_Listad.TextMatrix(r_int_Contad, 4)
            r_dbl_SegViv = grd_CliNCon_Listad.TextMatrix(r_int_Contad, 5)
            r_dbl_Portes = grd_CliNCon_Listad.TextMatrix(r_int_Contad, 6)
            r_dbl_MtoCuo = grd_CliNCon_Listad.TextMatrix(r_int_Contad, 7)
            r_dbl_SalCap = grd_CliNCon_Listad.TextMatrix(r_int_Contad, 8)
            
            If Not ff_Inserta_HipCuo(moddat_g_str_NumOpe, 1, r_int_NumCuo, r_str_FecVct, r_dbl_Capita, r_dbl_Intere, r_dbl_SegPre, r_dbl_SegViv, r_dbl_Portes, r_dbl_SalCap, 0, 0, 0) Then
               r_str_Situac = 1
               r_str_CadErr = "No se pudo completar el procedimiento USP_CRE_HIPCUO_CREA."
               Exit For
            End If
            r_str_Situac = 2
            '*** Actualiza log
            moddat_g_int_CntErr = 0
            r_int_ConAux = 1
            g_str_Parame = "USP_CRE_PROCRODET ("
            g_str_Parame = g_str_Parame & Format(CDate(moddat_g_str_FecSis), "yyyymmdd") & ", "
            g_str_Parame = g_str_Parame & r_int_NumPro & ", "
            g_str_Parame = g_str_Parame & r_int_ConAux & ", "
            g_str_Parame = g_str_Parame & r_str_Situac & " , "
            g_str_Parame = g_str_Parame & 1 & " , "                                             'Tipo de Cronograma
            g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "' , "
            g_str_Parame = g_str_Parame & r_int_NumCuo & " , "
            g_str_Parame = g_str_Parame & Format(CDate(r_str_FecVct), "yyyymmdd") & " , "
            g_str_Parame = g_str_Parame & r_dbl_Capita & " , "
            g_str_Parame = g_str_Parame & r_dbl_Intere & " , "
            g_str_Parame = g_str_Parame & r_dbl_SegPre & " , "
            g_str_Parame = g_str_Parame & r_dbl_SegViv & " , "
            g_str_Parame = g_str_Parame & 0 & " , "
            g_str_Parame = g_str_Parame & r_dbl_Portes & " , "
            g_str_Parame = g_str_Parame & r_dbl_MtoCuo & " , "
            g_str_Parame = g_str_Parame & r_dbl_SalCap & " , "
            g_str_Parame = g_str_Parame & "'" & r_str_CadErr & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
            g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "' ) "
            
            Do While (moddat_g_int_CntErr = 0)
               If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
                  If MsgBox("No se pudo completar el procedimiento. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
                     moddat_g_int_CntErr = 1
                  Else
                     moddat_g_int_CntErr = 0
                  End If
               Else
                  moddat_g_int_CntErr = 1
               End If
            Loop
         End If
      Else
         'carga variables e inserta cuota
         r_str_FecVct = grd_CliNCon_Listad.TextMatrix(r_int_Contad, 1)
         r_dbl_Capita = grd_CliNCon_Listad.TextMatrix(r_int_Contad, 2)
         r_dbl_Intere = grd_CliNCon_Listad.TextMatrix(r_int_Contad, 3)
         r_dbl_SegPre = grd_CliNCon_Listad.TextMatrix(r_int_Contad, 4)
         r_dbl_SegViv = grd_CliNCon_Listad.TextMatrix(r_int_Contad, 5)
         r_dbl_Portes = grd_CliNCon_Listad.TextMatrix(r_int_Contad, 6)
         r_dbl_MtoCuo = grd_CliNCon_Listad.TextMatrix(r_int_Contad, 7)
         r_dbl_SalCap = grd_CliNCon_Listad.TextMatrix(r_int_Contad, 8)
         
         r_int_NumCuo = r_int_NumCuo + 1
         
         If Not ff_Inserta_HipCuo(moddat_g_str_NumOpe, 1, r_int_NumCuo, r_str_FecVct, r_dbl_Capita, r_dbl_Intere, r_dbl_SegPre, r_dbl_SegViv, r_dbl_Portes, r_dbl_SalCap, 0, 0, 0) Then
            r_str_Situac = 1
            r_str_CadErr = "No se pudo completar el procedimiento USP_CRE_HIPCUO_CREA."
            Exit For
         End If
         
         r_str_Situac = 2
         '*** Actualiza log
            moddat_g_int_CntErr = 0
            r_int_ConAux = r_int_ConAux + 1
            g_str_Parame = "USP_CRE_PROCRODET ("
            g_str_Parame = g_str_Parame & Format(CDate(moddat_g_str_FecSis), "yyyymmdd") & ", "
            g_str_Parame = g_str_Parame & r_int_NumPro & ", "
            g_str_Parame = g_str_Parame & r_int_ConAux & ", "
            g_str_Parame = g_str_Parame & r_str_Situac & " , "
            g_str_Parame = g_str_Parame & 1 & " , "                                             'Tipo de Cronograma
            g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "' , "
            g_str_Parame = g_str_Parame & r_int_NumCuo & " , "
            g_str_Parame = g_str_Parame & Format(CDate(r_str_FecVct), "yyyymmdd") & " , "
            g_str_Parame = g_str_Parame & r_dbl_Capita & " , "
            g_str_Parame = g_str_Parame & r_dbl_Intere & " , "
            g_str_Parame = g_str_Parame & r_dbl_SegPre & " , "
            g_str_Parame = g_str_Parame & r_dbl_SegViv & " , "
            g_str_Parame = g_str_Parame & 0 & " , "
            g_str_Parame = g_str_Parame & r_dbl_Portes & " , "
            g_str_Parame = g_str_Parame & r_dbl_MtoCuo & " , "
            g_str_Parame = g_str_Parame & r_dbl_SalCap & " , "
            g_str_Parame = g_str_Parame & "'" & r_str_CadErr & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
            g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
            g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "' ) "
            
            Do While (moddat_g_int_CntErr = 0)
               If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
                  If MsgBox("No se pudo completar el procedimiento. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
                     moddat_g_int_CntErr = 1
                  Else
                     moddat_g_int_CntErr = 0
                  End If
               Else
                  moddat_g_int_CntErr = 1
               End If
            Loop
      End If
   Next r_int_Contad
   
   If r_str_Situac = 1 Then
      '*** Actualiza log
      moddat_g_int_CntErr = 0
      r_int_ConAux = r_int_ConAux + 1
      g_str_Parame = "USP_CRE_PROCRODET ("
      g_str_Parame = g_str_Parame & Format(CDate(moddat_g_str_FecSis), "yyyymmdd") & ", "
      g_str_Parame = g_str_Parame & r_int_NumPro & ", "
      g_str_Parame = g_str_Parame & r_int_ConAux & ", "
      g_str_Parame = g_str_Parame & r_str_Situac & " , "
      g_str_Parame = g_str_Parame & 1 & " , "                                             'Tipo de Cronograma
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "' , "
      g_str_Parame = g_str_Parame & r_int_NumCuo & " , "
      g_str_Parame = g_str_Parame & Format(CDate(r_str_FecVct), "yyyymmdd") & " , "
      g_str_Parame = g_str_Parame & r_dbl_Capita & " , "
      g_str_Parame = g_str_Parame & r_dbl_Intere & " , "
      g_str_Parame = g_str_Parame & r_dbl_SegPre & " , "
      g_str_Parame = g_str_Parame & r_dbl_SegViv & " , "
      g_str_Parame = g_str_Parame & 0 & " , "
      g_str_Parame = g_str_Parame & r_dbl_Portes & " , "
      g_str_Parame = g_str_Parame & r_dbl_MtoCuo & " , "
      g_str_Parame = g_str_Parame & r_dbl_SalCap & " , "
      g_str_Parame = g_str_Parame & "'" & r_str_CadErr & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "' ) "
      
      Do While (moddat_g_int_CntErr = 0)
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
            If MsgBox("No se pudo completar el procedimiento. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
               moddat_g_int_CntErr = 1
            Else
               moddat_g_int_CntErr = 0
            End If
         Else
            moddat_g_int_CntErr = 1
         End If
      Loop
   End If
   
   fs_Actualiza_Cronograma_CLITNC = True
End Function

Private Function ff_Inserta_HipCuo(ByVal p_NumOpe As String, ByVal p_TipCro As Integer, ByVal p_NumCuo As Integer, ByVal p_FecVct As String, ByVal p_Capita As Double, ByVal p_intere As Double, ByVal p_SegDes As Double, ByVal p_SegViv As Double, ByVal p_OtrGas As Double, ByVal p_SalCap As Double, ByVal p_ComCrc As Double, ByVal p_ComPbp As Double, ByVal p_ComCof As Double) As Integer
   ff_Inserta_HipCuo = False
   
   'Grabando Cabecera de Credito
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_CRE_HIPCUO_CREA ("
      g_str_Parame = g_str_Parame & "'" & p_NumOpe & "', "
      g_str_Parame = g_str_Parame & CStr(p_TipCro) & ", "
      g_str_Parame = g_str_Parame & CStr(p_NumCuo) & ", "
      g_str_Parame = g_str_Parame & Format(CDate(p_FecVct), "yyyymmdd") & ", "
      g_str_Parame = g_str_Parame & CStr(p_Capita) & ", "
      g_str_Parame = g_str_Parame & CStr(p_intere) & ", "
      g_str_Parame = g_str_Parame & CStr(p_SegDes) & ", "
      g_str_Parame = g_str_Parame & CStr(p_SegViv) & ", "
      g_str_Parame = g_str_Parame & CStr(p_OtrGas) & ", "
      g_str_Parame = g_str_Parame & CStr(p_SalCap) & ", "
      g_str_Parame = g_str_Parame & CStr(p_ComCrc) & ", "
      g_str_Parame = g_str_Parame & CStr(p_ComPbp) & ", "
      g_str_Parame = g_str_Parame & CStr(p_ComCof) & ", "
      
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                           'Código Usuario
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                           'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                            'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "                           'Código Sucursal
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_CRE_HIPCUO_CREA. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Function
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop

   ff_Inserta_HipCuo = True
End Function

Private Function fs_Carga_CofideNoConcesional() As Boolean
Dim r_obj_Excel         As Excel.Application
Dim r_int_FilExc        As Integer
Dim r_int_FilGrd        As Integer
Dim r_int_NumCuo        As Integer
Dim r_dbl_SumNoc        As Double
Dim r_dbl_SumCon        As Double
Dim r_dat_FecIni        As Date
Dim r_dat_FecFin        As Date
   
   fs_Carga_CofideNoConcesional = False
   Call gs_LimpiaGrid(grd_MViNCo_Listad)
   
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Open FileName:=txt_NomArc.Text
   
   'Valida y Carga Cronograma No Concesional FMV
   r_int_FilExc = 0
   r_int_FilGrd = 0
   r_dat_FecIni = CDate("01/01/2007")
   
   Do While Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 2).Value) <> ""
      If Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 2).Value) = "2" Then
         If r_int_FilExc >= grd_MViNCo_Listad.Rows Then
            grd_MViNCo_Listad.Rows = grd_MViNCo_Listad.Rows + 1
         End If
         
         'verifica número de operación
         If InStr(1, l_str_OpeMVi, Mid(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 1).Value), Len(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 1).Value)) - 4, 5), vbTextCompare) = 0 Then
            Call gs_LimpiaGrid(grd_MViNCo_Listad)
            MsgBox "No coincide el numero de operación MIVIVIENDA del sistema con el numero de contrato del archivo." & vbCrLf & "Favor verificar.", vbCritical, modgen_g_str_NomPlt
            GoTo Salir
         End If
         
         r_int_NumCuo = Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 3).Value)
         
         If l_int_PerGra = 0 Then
            If Not IsDate(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 4).Value)) Then
               Call gs_LimpiaGrid(grd_MViNCo_Listad)
               MsgBox "Fecha Vencimiento invalida (FMV - TNC - Cuota: " & CStr(r_int_NumCuo) & ").", vbCritical, modgen_g_str_NomPlt
               GoTo Salir
            End If
            If Not Val(r_obj_Excel.Cells(r_int_FilExc + 2, 7).Value) > 0 Then
               Call gs_LimpiaGrid(grd_MViNCo_Listad)
               MsgBox "Capital debe ser mayor a cero (FMV - TNC - Cuota: " & CStr(r_int_NumCuo) & ").", vbCritical, modgen_g_str_NomPlt
               GoTo Salir
            End If
            If Not Val(r_obj_Excel.Cells(r_int_FilExc + 2, 8).Value) > 0 Then
               Call gs_LimpiaGrid(grd_MViNCo_Listad)
               MsgBox "Interes debe ser mayor a cero (FMV - TNC - Cuota: " & CStr(r_int_NumCuo) & ").", vbCritical, modgen_g_str_NomPlt
               GoTo Salir
            End If
            'If Not Val(r_obj_Excel.Cells(r_int_FilExc + 2, 9).Value) > 0 Then
            '   Call gs_LimpiaGrid(grd_MViNCo_Listad)
            '   MsgBox "Comisión debe ser mayor a cero (FMV - TNC - Cuota: " & CStr(r_int_NumCuo) & ").", vbCritical, modgen_g_str_NomPlt
            '   GoTo Salir
            'End If
         End If
         
         r_dbl_SumNoc = r_obj_Excel.Cells(r_int_FilExc + 2, 7).Value + r_obj_Excel.Cells(r_int_FilExc + 2, 8).Value + r_obj_Excel.Cells(r_int_FilExc + 2, 9).Value
         If Format(r_dbl_SumNoc, "###,###,##0.00") <> Format(r_obj_Excel.Cells(r_int_FilExc + 2, 10).Value, "###,###,##0.00") Then
            Call gs_LimpiaGrid(grd_MViNCo_Listad)
            MsgBox "Total Cuota no es igual a suma de campos Capital, Interes y Comision (FMV - TNC - Cuota: " & CStr(r_int_NumCuo) & ").", vbCritical, modgen_g_str_NomPlt
            GoTo Salir
         End If
         
         grd_MViNCo_Listad.TextMatrix(r_int_FilGrd, 0) = Format(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 3).Value), "000")
         grd_MViNCo_Listad.TextMatrix(r_int_FilGrd, 1) = Format(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 4).Value), "dd/mm/yyyy")
         grd_MViNCo_Listad.TextMatrix(r_int_FilGrd, 2) = Format(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 7).Value), "###,###,##0.00")
         grd_MViNCo_Listad.TextMatrix(r_int_FilGrd, 3) = Format(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 8).Value), "###,###,##0.00")
         grd_MViNCo_Listad.TextMatrix(r_int_FilGrd, 4) = Format(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 9).Value), "###,###,##0.00")
         grd_MViNCo_Listad.TextMatrix(r_int_FilGrd, 5) = Format(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 10).Value), "###,###,##0.00")
         grd_MViNCo_Listad.TextMatrix(r_int_FilGrd, 6) = Format(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 11).Value), "###,###,##0.00")
         '------------------------------------------------------------------------------------------
         
         If grd_MViNCo_Listad.Rows > 1 Then
            Dim swtCapitalOld As Double
            Dim swtCapitalNew As Double
         
            swtCapitalOld = CDbl(grd_MViNCo_Listad.TextMatrix(r_int_FilGrd - 1, 5))
            swtCapitalNew = CDbl(grd_MViNCo_Listad.TextMatrix(r_int_FilGrd, 5))
            If swtCapitalOld <> swtCapitalNew Then
               swtContador = swtContador + 1
             End If
         End If
         '------------------------------------------------------------------------------------------
         r_dat_FecFin = Format(Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 4).Value), "dd/mm/yyyy")
         If r_dat_FecIni > r_dat_FecFin Then
            MsgBox "TNC: Fecha de vencimiento de la cuota " & Trim(r_obj_Excel.Cells(r_int_FilExc + 2, 3).Value) & " es menor que la anterior.", vbCritical, modgen_g_str_NomPlt
            GoTo Salir
         End If
         r_dat_FecIni = r_dat_FecFin
         
         r_int_FilGrd = r_int_FilGrd + 1
      End If
      
      r_int_FilExc = r_int_FilExc + 1
   Loop
   
   If r_int_NumCuo = 0 Then
      MsgBox "El archivo seleccionado no tiene el formato adecuado.", vbCritical, modgen_g_str_NomPlt
      GoTo Salir
   End If
   
   fs_Carga_CofideNoConcesional = True

Salir:
   r_obj_Excel.Quit
   Set r_obj_Excel = Nothing
   
End Function

Private Function fs_UbicaCuotas_CronogramaFMVTNC() As Boolean
Dim r_int_FilGrd        As Integer
Dim r_int_TotReg        As Integer
Dim r_int_NumCuo        As Integer
Dim r_int_CuoCar        As Integer
Dim r_str_FecNco        As String
Dim r_str_FecCon        As String
Dim r_dbl_CuoTotOld     As Double
Dim r_dbl_CuoTotNew     As Double

   fs_UbicaCuotas_CronogramaFMVTNC = False

   If grd_MViNCo_Listad.Rows = 0 Then
      Exit Function
   End If
                 
   'Seleccionar Cuota: FMV - TNC (Siempre que no tenga cuotas dobles)
   If l_int_CuoDbl = 1 Then
      r_dbl_CuoTotOld = 0
      r_dbl_CuoTotNew = 0
      r_int_TotReg = grd_MViNCo_Listad.Rows
      
      For r_int_FilGrd = 0 To r_int_TotReg - 1
         r_int_NumCuo = grd_MViNCo_Listad.TextMatrix(r_int_FilGrd, 0)
         r_dbl_CuoTotNew = grd_MViNCo_Listad.TextMatrix(r_int_FilGrd, 5)
         
         If r_dbl_CuoTotNew <> r_dbl_CuoTotOld Then
            If r_int_NumCuo <> r_int_TotReg Then
               r_int_CuoCar = r_int_NumCuo
               r_str_FecNco = grd_MViNCo_Listad.TextMatrix(r_int_FilGrd, 1)
            End If
         End If
         
         r_dbl_CuoTotOld = r_dbl_CuoTotNew
      Next r_int_FilGrd
      grd_MViNCo_Listad.TextMatrix(r_int_CuoCar - 1, 7) = "X"
   End If
   
   fs_UbicaCuotas_CronogramaFMVTNC = True
   
End Function

Private Sub cmb_TipCro_Click()
   Call gs_LimpiaGrid(grd_MViNCo_Listad)
   Call gs_LimpiaGrid(grd_MViCon_Listad)
   Call gs_LimpiaGrid(grd_CliCon_Listad)
   Call gs_LimpiaGrid(grd_CliNCon_Listad)
   txt_NomArc.Text = ""
   
   If cmb_TipCro.ListIndex = 0 Then
      cmd_Generar.Enabled = False
      tab_Cronog.TabCaption(0) = ""
      tab_Cronog.TabCaption(1) = ""
      tab_Cronog.TabCaption(2) = ""
      tab_Cronog.TabCaption(3) = "CLIENTE - Tramo No Concesional"
      tab_Cronog.TabVisible(0) = False
      tab_Cronog.TabVisible(1) = False
      tab_Cronog.TabVisible(2) = False
      tab_Cronog.TabVisible(3) = True
      tab_Cronog.Tab = 3
      
   ElseIf cmb_TipCro.ListIndex = 2 Then
      cmd_Generar.Enabled = False
      tab_Cronog.TabCaption(0) = "FMV - Tramo No Concesional"
      tab_Cronog.TabCaption(1) = ""
      tab_Cronog.TabCaption(2) = ""
      tab_Cronog.TabCaption(3) = ""
      tab_Cronog.TabVisible(0) = True
      tab_Cronog.TabVisible(1) = False
      tab_Cronog.TabVisible(2) = False
      tab_Cronog.TabVisible(3) = False
      tab_Cronog.Tab = 0
   
   ElseIf cmb_TipCro.ListIndex = 3 Then
      cmd_Generar.Enabled = False
      tab_Cronog.TabCaption(0) = "CME - Cronograma"
      tab_Cronog.TabCaption(1) = ""
      tab_Cronog.TabCaption(2) = ""
      tab_Cronog.TabCaption(3) = ""
      tab_Cronog.TabVisible(0) = True
      tab_Cronog.TabVisible(1) = False
      tab_Cronog.TabVisible(2) = False
      tab_Cronog.TabVisible(3) = False
      tab_Cronog.Tab = 0
      
   Else
      Call fs_Configura_CargaProd
      tab_Cronog.TabVisible(3) = False
      cmd_Generar.Enabled = True
   End If
End Sub

Private Function fs_UbicaMonto_Prepago() As Double
Dim r_int_FilGrd        As Integer
Dim r_int_TotReg        As Integer
Dim r_int_NumCuo        As Integer
Dim r_int_CuoCar        As Integer
Dim r_str_FecNco        As String
Dim r_str_FecCon        As String
Dim r_dbl_CuoTotOld     As Double
Dim r_dbl_CuoTotNew     As Double
  
   'Validaciones
   fs_UbicaMonto_Prepago = 0
   If grd_MViNCo_Listad.Rows = 0 Then
      Exit Function
   End If
   If grd_MViNCo_Listad.Rows = 0 Then
      Exit Function
   End If
   
   'Seleccionar Último Prepago
   If l_int_CuoDbl = 1 Then
      r_dbl_CuoTotOld = 0
      r_dbl_CuoTotNew = 0
      r_int_TotReg = grd_MViNCo_Listad.Rows
      
      For r_int_FilGrd = 0 To r_int_TotReg - 1
         r_int_NumCuo = grd_MViNCo_Listad.TextMatrix(r_int_FilGrd, 0)
         r_dbl_CuoTotNew = grd_MViNCo_Listad.TextMatrix(r_int_FilGrd, 5)
         
         If r_dbl_CuoTotNew <> r_dbl_CuoTotOld Then
            If r_int_NumCuo <> r_int_TotReg Then
               r_int_CuoCar = r_int_NumCuo
            End If
         End If
         r_dbl_CuoTotOld = r_dbl_CuoTotNew
      Next r_int_FilGrd
            
      If r_int_CuoCar > 1 Then
         fs_UbicaMonto_Prepago = grd_MViNCo_Listad.TextMatrix(r_int_CuoCar - 2, 5)
      End If
   End If
End Function

