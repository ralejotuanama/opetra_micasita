VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_Pos_ConCli_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form4"
   ClientHeight    =   10005
   ClientLeft      =   45
   ClientTop       =   1635
   ClientWidth     =   11370
   Icon            =   "OpeTra_frm_400.frx":0000
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10005
   ScaleWidth      =   11370
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   10005
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11370
      _Version        =   65536
      _ExtentX        =   20055
      _ExtentY        =   17648
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
      Begin Threed.SSPanel SSPanel6 
         Height          =   585
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   11295
         _Version        =   65536
         _ExtentX        =   19923
         _ExtentY        =   1032
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
         Begin Threed.SSPanel SSPanel7 
            Height          =   210
            Left            =   660
            TabIndex        =   2
            Top             =   240
            Width           =   3270
            _Version        =   65536
            _ExtentX        =   5768
            _ExtentY        =   370
            _StockProps     =   15
            Caption         =   "Posición Consolidada del Cliente"
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
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   10680
            Top             =   120
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            WindowTitle     =   "Presentación Preliminar"
            WindowControlBox=   -1  'True
            WindowMaxButton =   -1  'True
            WindowMinButton =   -1  'True
            WindowState     =   2
            PrintFileLinesPerPage=   60
            WindowShowPrintSetupBtn=   -1  'True
            WindowShowRefreshBtn=   -1  'True
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   60
            Picture         =   "OpeTra_frm_400.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel12 
         Height          =   645
         Left            =   30
         TabIndex        =   3
         Top             =   650
         Width           =   11295
         _Version        =   65536
         _ExtentX        =   19923
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
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   1830
            Picture         =   "OpeTra_frm_400.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Resumen de Crédito"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_EstCta 
            Height          =   585
            Left            =   1230
            Picture         =   "OpeTra_frm_400.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Imprimir Estado de Cuenta"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10680
            Picture         =   "OpeTra_frm_400.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_VerPag 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_400.frx":0EA4
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Consulta de Pagos del Cliente"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ImpCro 
            Height          =   585
            Left            =   630
            Picture         =   "OpeTra_frm_400.frx":11AE
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Consulta de Cronogramas de Pago"
            Top             =   30
            Width           =   585
         End
         Begin VB.Label Label1 
            Caption         =   "Nro. Doc. Id.:"
            Height          =   285
            Left            =   60
            TabIndex        =   9
            Top             =   1740
            Width           =   1065
         End
      End
      Begin Threed.SSPanel SSPanel20 
         Height          =   3015
         Left            =   30
         TabIndex        =   10
         Top             =   6930
         Width           =   11295
         _Version        =   65536
         _ExtentX        =   19923
         _ExtentY        =   5318
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
         Begin Threed.SSPanel pnl_Cuo_TotSal 
            Height          =   285
            Left            =   9450
            TabIndex        =   11
            Top             =   2640
            Width           =   1410
            _Version        =   65536
            _ExtentX        =   2487
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "0.00000 "
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
         Begin Threed.SSPanel pnl_Cuo_TotPag 
            Height          =   285
            Left            =   8025
            TabIndex        =   12
            Top             =   2640
            Width           =   1440
            _Version        =   65536
            _ExtentX        =   2540
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "0.00000 "
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
         Begin Threed.SSPanel pnl_Cuo_TotDeu 
            Height          =   285
            Left            =   6510
            TabIndex        =   13
            Top             =   2640
            Width           =   1530
            _Version        =   65536
            _ExtentX        =   2699
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "0.00000 "
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
         Begin Threed.SSPanel SSPanel11 
            Height          =   285
            Left            =   165
            TabIndex        =   14
            Top             =   300
            Width           =   855
            _Version        =   65536
            _ExtentX        =   1508
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
         Begin Threed.SSPanel SSPanel3 
            Height          =   285
            Left            =   1005
            TabIndex        =   15
            Top             =   300
            Width           =   1350
            _Version        =   65536
            _ExtentX        =   2372
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Vencim."
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
         Begin Threed.SSPanel SSPanel4 
            Height          =   285
            Left            =   6465
            TabIndex        =   16
            Top             =   300
            Width           =   1545
            _Version        =   65536
            _ExtentX        =   2734
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "T. Cuota"
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
            Left            =   2355
            TabIndex        =   17
            Top             =   300
            Width           =   1050
            _Version        =   65536
            _ExtentX        =   1843
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "D. Atraso"
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
            Left            =   7965
            TabIndex        =   18
            Top             =   300
            Width           =   1545
            _Version        =   65536
            _ExtentX        =   2734
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "T. Pagado"
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
            Left            =   5145
            TabIndex        =   19
            Top             =   300
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Ult. Pago"
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
            Left            =   9345
            TabIndex        =   20
            Top             =   300
            Width           =   1545
            _Version        =   65536
            _ExtentX        =   2734
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Saldo Deudor"
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
         Begin Threed.SSPanel SSPanel19 
            Height          =   285
            Left            =   3375
            TabIndex        =   21
            Top             =   300
            Width           =   1800
            _Version        =   65536
            _ExtentX        =   3175
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Situación"
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
         Begin MSFlexGridLib.MSFlexGrid grd_Cuotas 
            Height          =   2025
            Left            =   135
            TabIndex        =   24
            Top             =   570
            Width           =   11100
            _ExtentX        =   19579
            _ExtentY        =   3572
            _Version        =   393216
            Rows            =   8
            Cols            =   8
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin VB.Label lbl_Totale 
            Alignment       =   1  'Right Justify
            Caption         =   "Totales ==> US$ "
            Height          =   255
            Left            =   4740
            TabIndex        =   23
            Top             =   2670
            Width           =   1515
         End
         Begin VB.Label Label12 
            Caption         =   "Resumen de Cuotas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   160
            TabIndex        =   22
            Top             =   60
            Width           =   1875
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   5560
         Left            =   30
         TabIndex        =   25
         Top             =   1330
         Width           =   11295
         _Version        =   65536
         _ExtentX        =   19923
         _ExtentY        =   9807
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
            Height          =   4860
            Left            =   45
            TabIndex        =   26
            Top             =   630
            Width           =   11175
            _ExtentX        =   19711
            _ExtentY        =   8573
            _Version        =   393216
            Style           =   1
            Tabs            =   9
            Tab             =   7
            TabsPerRow      =   9
            TabHeight       =   520
            TabCaption(0)   =   "Resumen"
            TabPicture(0)   =   "OpeTra_frm_400.frx":14B8
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "SSPanel41"
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Titular"
            TabPicture(1)   =   "OpeTra_frm_400.frx":14D4
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SSPanel14"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Cónyuge"
            TabPicture(2)   =   "OpeTra_frm_400.frx":14F0
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "SSPanel33"
            Tab(2).ControlCount=   1
            TabCaption(3)   =   "Inmueble"
            TabPicture(3)   =   "OpeTra_frm_400.frx":150C
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "SSPanel13"
            Tab(3).ControlCount=   1
            TabCaption(4)   =   "Datos Crédito"
            TabPicture(4)   =   "OpeTra_frm_400.frx":1528
            Tab(4).ControlEnabled=   0   'False
            Tab(4).Control(0)=   "SSPanel16"
            Tab(4).ControlCount=   1
            TabCaption(5)   =   "Excepciones y Condiciones"
            TabPicture(5)   =   "OpeTra_frm_400.frx":1544
            Tab(5).ControlEnabled=   0   'False
            Tab(5).Control(0)=   "SSPanel21"
            Tab(5).ControlCount=   1
            TabCaption(6)   =   "Datos Garantía"
            TabPicture(6)   =   "OpeTra_frm_400.frx":1560
            Tab(6).ControlEnabled=   0   'False
            Tab(6).Control(0)=   "SSPanel15"
            Tab(6).ControlCount=   1
            TabCaption(7)   =   "Datos RCC"
            TabPicture(7)   =   "OpeTra_frm_400.frx":157C
            Tab(7).ControlEnabled=   -1  'True
            Tab(7).Control(0)=   "SSPanel18"
            Tab(7).Control(0).Enabled=   0   'False
            Tab(7).ControlCount=   1
            TabCaption(8)   =   "Clasificación Cliente "
            TabPicture(8)   =   "OpeTra_frm_400.frx":1598
            Tab(8).ControlEnabled=   0   'False
            Tab(8).Control(0)=   "SSPanel17"
            Tab(8).ControlCount=   1
            Begin Threed.SSPanel SSPanel41 
               Height          =   4425
               Left            =   -74940
               TabIndex        =   32
               Top             =   360
               Width           =   11055
               _Version        =   65536
               _ExtentX        =   19500
               _ExtentY        =   7805
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
               Begin MSFlexGridLib.MSFlexGrid grd_Listad 
                  Height          =   4380
                  Index           =   1
                  Left            =   30
                  TabIndex        =   33
                  Top             =   30
                  Width           =   10950
                  _ExtentX        =   19315
                  _ExtentY        =   7726
                  _Version        =   393216
                  Rows            =   21
                  FixedRows       =   0
                  FixedCols       =   0
                  BackColorSel    =   32768
                  FocusRect       =   0
                  ScrollBars      =   2
                  SelectionMode   =   1
               End
            End
            Begin Threed.SSPanel SSPanel17 
               Height          =   4425
               Left            =   -74970
               TabIndex        =   34
               Top             =   390
               Width           =   11055
               _Version        =   65536
               _ExtentX        =   19500
               _ExtentY        =   7805
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
               Begin Threed.SSPanel pnl_Tit_Produc 
                  Height          =   285
                  Left            =   60
                  TabIndex        =   35
                  Top             =   60
                  Width           =   1510
                  _Version        =   65536
                  _ExtentX        =   2663
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Año"
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
               Begin Threed.SSPanel pnl_Tit_NumSol 
                  Height          =   285
                  Left            =   1550
                  TabIndex        =   36
                  Top             =   60
                  Width           =   2240
                  _Version        =   65536
                  _ExtentX        =   3951
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Mes"
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
               Begin Threed.SSPanel pnl_Tit_DocIde 
                  Height          =   285
                  Left            =   3700
                  TabIndex        =   37
                  Top             =   60
                  Width           =   3660
                  _Version        =   65536
                  _ExtentX        =   6456
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Clasificación Interna"
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
               Begin Threed.SSPanel pnl_Tit_FecSol 
                  Height          =   285
                  Left            =   7155
                  TabIndex        =   38
                  Top             =   60
                  Width           =   3720
                  _Version        =   65536
                  _ExtentX        =   6562
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Clasificación Alineada"
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
               Begin MSFlexGridLib.MSFlexGrid grd_Listad_his 
                  Height          =   4030
                  Left            =   45
                  TabIndex        =   39
                  Top             =   360
                  Width           =   10950
                  _ExtentX        =   19315
                  _ExtentY        =   7117
                  _Version        =   393216
                  Rows            =   15
                  Cols            =   5
                  FixedRows       =   0
                  FixedCols       =   0
                  BackColorSel    =   32768
                  FocusRect       =   0
                  ScrollBars      =   2
                  SelectionMode   =   1
               End
            End
            Begin Threed.SSPanel SSPanel18 
               Height          =   4440
               Left            =   30
               TabIndex        =   40
               Top             =   360
               Width           =   11055
               _Version        =   65536
               _ExtentX        =   19500
               _ExtentY        =   7832
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
               Begin VB.CommandButton cmd_Export 
                  Height          =   500
                  Left            =   10410
                  Picture         =   "OpeTra_frm_400.frx":15B4
                  Style           =   1  'Graphical
                  TabIndex        =   48
                  ToolTipText     =   "Exportar Excel"
                  Top             =   50
                  Width           =   585
               End
               Begin Threed.SSPanel pnl_Total7 
                  Height          =   285
                  Left            =   9240
                  TabIndex        =   41
                  Top             =   2110
                  Width           =   1215
                  _Version        =   65536
                  _ExtentX        =   2143
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "0.00000 "
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
               Begin Threed.SSPanel pnl_Total6 
                  Height          =   285
                  Left            =   7935
                  TabIndex        =   42
                  Top             =   2110
                  Width           =   1305
                  _Version        =   65536
                  _ExtentX        =   2302
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "0.00000 "
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
               Begin Threed.SSPanel pnl_Total1 
                  Height          =   285
                  Left            =   1410
                  TabIndex        =   43
                  Top             =   2110
                  Width           =   1305
                  _Version        =   65536
                  _ExtentX        =   2302
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "0.00000 "
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
               Begin Threed.SSPanel pnl_Total2 
                  Height          =   285
                  Left            =   2715
                  TabIndex        =   44
                  Top             =   2110
                  Width           =   1305
                  _Version        =   65536
                  _ExtentX        =   2302
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "0.00000 "
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
               Begin Threed.SSPanel pnl_Total3 
                  Height          =   285
                  Left            =   4020
                  TabIndex        =   45
                  Top             =   2110
                  Width           =   1305
                  _Version        =   65536
                  _ExtentX        =   2302
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "0.00000 "
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
               Begin Threed.SSPanel pnl_Total4 
                  Height          =   285
                  Left            =   5325
                  TabIndex        =   46
                  Top             =   2110
                  Width           =   1305
                  _Version        =   65536
                  _ExtentX        =   2302
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "0.00000 "
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
               Begin Threed.SSPanel pnl_Total5 
                  Height          =   285
                  Left            =   6630
                  TabIndex        =   47
                  Top             =   2110
                  Width           =   1305
                  _Version        =   65536
                  _ExtentX        =   2302
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "0.00000 "
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
               Begin MSFlexGridLib.MSFlexGrid grd_Listad_rcc2 
                  Height          =   1500
                  Left            =   45
                  TabIndex        =   49
                  Top             =   2940
                  Width           =   10965
                  _ExtentX        =   19341
                  _ExtentY        =   2646
                  _Version        =   393216
                  Rows            =   12
                  Cols            =   10
                  FixedRows       =   0
                  FixedCols       =   0
                  BackColorSel    =   32768
                  FocusRect       =   0
                  ScrollBars      =   2
                  SelectionMode   =   1
               End
               Begin Threed.SSPanel SSPanel23 
                  Height          =   285
                  Left            =   7290
                  TabIndex        =   50
                  Top             =   2685
                  Width           =   1020
                  _Version        =   65536
                  _ExtentX        =   1799
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Monto (S/.)"
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
                  Left            =   2430
                  TabIndex        =   51
                  Top             =   2685
                  Width           =   1410
                  _Version        =   65536
                  _ExtentX        =   2487
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Clasificacion"
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
               Begin Threed.SSPanel SSPanel25 
                  Height          =   285
                  Left            =   8310
                  TabIndex        =   52
                  Top             =   2685
                  Width           =   1035
                  _Version        =   65536
                  _ExtentX        =   1817
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Monto (US$)"
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
               Begin Threed.SSPanel SSPanel28 
                  Height          =   285
                  Left            =   3840
                  TabIndex        =   53
                  Top             =   2685
                  Width           =   3450
                  _Version        =   65536
                  _ExtentX        =   6085
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Tipo Deuda"
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
               Begin MSFlexGridLib.MSFlexGrid grd_Listad_rcc1 
                  Height          =   1545
                  Left            =   45
                  TabIndex        =   54
                  Top             =   570
                  Width           =   10965
                  _ExtentX        =   19341
                  _ExtentY        =   2725
                  _Version        =   393216
                  Rows            =   1
                  Cols            =   9
                  FixedRows       =   0
                  FixedCols       =   0
                  BackColorSel    =   32768
                  FocusRect       =   0
                  ScrollBars      =   2
               End
               Begin Threed.SSPanel pnl_Periodo1 
                  Height          =   285
                  Left            =   1410
                  TabIndex        =   55
                  Top             =   315
                  Width           =   1305
                  _Version        =   65536
                  _ExtentX        =   2311
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Periodo 1"
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
               Begin Threed.SSPanel pnl_Periodo5 
                  Height          =   285
                  Left            =   6630
                  TabIndex        =   56
                  Top             =   315
                  Width           =   1305
                  _Version        =   65536
                  _ExtentX        =   2293
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Periodo 5"
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
               Begin Threed.SSPanel pnl_Periodo2 
                  Height          =   285
                  Left            =   2715
                  TabIndex        =   57
                  Top             =   315
                  Width           =   1305
                  _Version        =   65536
                  _ExtentX        =   2293
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Periodo 2"
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
               Begin Threed.SSPanel pnl_Periodo6 
                  Height          =   285
                  Left            =   7935
                  TabIndex        =   58
                  Top             =   315
                  Width           =   1305
                  _Version        =   65536
                  _ExtentX        =   2293
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Periodo 6"
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
               Begin Threed.SSPanel pnl_Periodo3 
                  Height          =   285
                  Left            =   4020
                  TabIndex        =   59
                  Top             =   315
                  Width           =   1305
                  _Version        =   65536
                  _ExtentX        =   2293
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Periodo 3"
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
               Begin Threed.SSPanel pnl_Tipo 
                  Height          =   285
                  Left            =   60
                  TabIndex        =   60
                  Top             =   315
                  Width           =   1350
                  _Version        =   65536
                  _ExtentX        =   2381
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Calif. \ Periodo"
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
               Begin Threed.SSPanel SSPanel22 
                  Height          =   285
                  Left            =   60
                  TabIndex        =   61
                  Top             =   2685
                  Width           =   2370
                  _Version        =   65536
                  _ExtentX        =   4180
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Nombre Empresa"
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
               Begin Threed.SSPanel pnl_Periodo4 
                  Height          =   285
                  Left            =   5325
                  TabIndex        =   62
                  Top             =   315
                  Width           =   1305
                  _Version        =   65536
                  _ExtentX        =   2293
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Periodo 4"
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
               Begin Threed.SSPanel SSPanel31 
                  Height          =   285
                  Left            =   9240
                  TabIndex        =   63
                  Top             =   315
                  Width           =   1125
                  _Version        =   65536
                  _ExtentX        =   1984
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "%"
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
               Begin Threed.SSPanel SSPanel32 
                  Height          =   285
                  Left            =   9345
                  TabIndex        =   64
                  Top             =   2685
                  Width           =   1035
                  _Version        =   65536
                  _ExtentX        =   1826
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Total"
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
               Begin VB.Label pnl_Detalle 
                  AutoSize        =   -1  'True
                  Caption         =   "Detalle"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   75
                  TabIndex        =   68
                  Top             =   2475
                  Width           =   615
               End
               Begin VB.Label Label6 
                  AutoSize        =   -1  'True
                  Caption         =   "Resumen"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   75
                  TabIndex        =   67
                  Top             =   70
                  Width           =   795
               End
               Begin VB.Label Label7 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "Totales ==> S/."
                  Height          =   195
                  Left            =   75
                  TabIndex        =   66
                  Top             =   2160
                  Width           =   1110
               End
               Begin VB.Label lbl_endeudado 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "SOBRE ENDEUDADO"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   7290
                  TabIndex        =   65
                  Top             =   75
                  Width           =   3075
               End
            End
            Begin Threed.SSPanel SSPanel15 
               Height          =   4425
               Left            =   -74940
               TabIndex        =   69
               Top             =   360
               Width           =   11055
               _Version        =   65536
               _ExtentX        =   19500
               _ExtentY        =   7805
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
               Begin MSFlexGridLib.MSFlexGrid grd_Listad_gar 
                  Height          =   4380
                  Left            =   45
                  TabIndex        =   70
                  Top             =   30
                  Width           =   10950
                  _ExtentX        =   19315
                  _ExtentY        =   7726
                  _Version        =   393216
                  Rows            =   8
                  FixedRows       =   0
                  FixedCols       =   0
                  BackColorSel    =   32768
                  FocusRect       =   0
                  ScrollBars      =   2
                  SelectionMode   =   1
               End
            End
            Begin Threed.SSPanel SSPanel21 
               Height          =   4425
               Left            =   -74970
               TabIndex        =   71
               Top             =   360
               Width           =   11055
               _Version        =   65536
               _ExtentX        =   19500
               _ExtentY        =   7805
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
               Begin VB.TextBox txt_LevCon 
                  Height          =   645
                  Left            =   1320
                  MaxLength       =   2000
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   72
                  Top             =   3720
                  Width           =   9680
               End
               Begin MSFlexGridLib.MSFlexGrid grd_LisExc 
                  Height          =   855
                  Left            =   45
                  TabIndex        =   73
                  Top             =   670
                  Width           =   10990
                  _ExtentX        =   19394
                  _ExtentY        =   1508
                  _Version        =   393216
                  Rows            =   21
                  Cols            =   6
                  FixedRows       =   0
                  FixedCols       =   0
                  BackColorSel    =   32768
                  FocusRect       =   0
                  ScrollBars      =   2
                  SelectionMode   =   1
               End
               Begin Threed.SSPanel SSPanel26 
                  Height          =   285
                  Left            =   60
                  TabIndex        =   74
                  Top             =   380
                  Width           =   1185
                  _Version        =   65536
                  _ExtentX        =   2090
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "F. Excepción"
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
               Begin Threed.SSPanel SSPanel27 
                  Height          =   285
                  Left            =   5610
                  TabIndex        =   75
                  Top             =   380
                  Width           =   5360
                  _Version        =   65536
                  _ExtentX        =   9454
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Descripción Excepción"
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
               Begin Threed.SSPanel SSPanel30 
                  Height          =   285
                  Left            =   1170
                  TabIndex        =   76
                  Top             =   380
                  Width           =   1185
                  _Version        =   65536
                  _ExtentX        =   2090
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "H. Excepción"
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
                  Left            =   2340
                  TabIndex        =   77
                  Top             =   380
                  Width           =   3285
                  _Version        =   65536
                  _ExtentX        =   5794
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Instancia"
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
               Begin MSFlexGridLib.MSFlexGrid grd_LisCon 
                  Height          =   975
                  Left            =   45
                  TabIndex        =   78
                  Top             =   2685
                  Width           =   10995
                  _ExtentX        =   19394
                  _ExtentY        =   1720
                  _Version        =   393216
                  Rows            =   21
                  Cols            =   4
                  FixedRows       =   0
                  FixedCols       =   0
                  BackColorSel    =   32768
                  FocusRect       =   0
                  ScrollBars      =   2
                  SelectionMode   =   1
               End
               Begin Threed.SSPanel SSPanel38 
                  Height          =   285
                  Left            =   60
                  TabIndex        =   79
                  Top             =   2380
                  Width           =   2745
                  _Version        =   65536
                  _ExtentX        =   4842
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Instancia"
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
               Begin Threed.SSPanel SSPanel40 
                  Height          =   285
                  Left            =   2680
                  TabIndex        =   80
                  Top             =   2380
                  Width           =   6615
                  _Version        =   65536
                  _ExtentX        =   11668
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Condiciones de Aprobación"
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
               Begin Threed.SSPanel pnl_TipAut 
                  Height          =   315
                  Left            =   1380
                  TabIndex        =   81
                  Top             =   1590
                  Width           =   9550
                  _Version        =   65536
                  _ExtentX        =   16845
                  _ExtentY        =   556
                  _StockProps     =   15
                  Caption         =   "INGRESO A INSTANCIA"
                  ForeColor       =   32768
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
                  Font3D          =   2
                  Alignment       =   1
               End
               Begin Threed.SSPanel SSPanel39 
                  Height          =   285
                  Left            =   9290
                  TabIndex        =   82
                  Top             =   2380
                  Width           =   1670
                  _Version        =   65536
                  _ExtentX        =   2928
                  _ExtentY        =   503
                  _StockProps     =   15
                  Caption         =   "Situación"
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
               Begin VB.Label Label5 
                  Caption         =   "Aprobación Condicionada"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   120
                  TabIndex        =   86
                  Top             =   2080
                  Width           =   2235
               End
               Begin VB.Label Label4 
                  Caption         =   "Excepciones Aplicadas"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   120
                  TabIndex        =   85
                  Top             =   75
                  Width           =   2355
               End
               Begin VB.Label Label3 
                  Caption         =   "Autorizado por:"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   84
                  Top             =   1620
                  Width           =   1095
               End
               Begin VB.Label Label2 
                  Caption         =   "Levantamiento de Condiciones:"
                  Height          =   495
                  Left            =   120
                  TabIndex        =   83
                  Top             =   3765
                  Width           =   1215
               End
            End
            Begin Threed.SSPanel SSPanel16 
               Height          =   4425
               Left            =   -74940
               TabIndex        =   87
               Top             =   360
               Width           =   11055
               _Version        =   65536
               _ExtentX        =   19500
               _ExtentY        =   7805
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
               Begin MSFlexGridLib.MSFlexGrid grd_Listad 
                  Height          =   4380
                  Index           =   2
                  Left            =   45
                  TabIndex        =   88
                  Top             =   30
                  Width           =   10950
                  _ExtentX        =   19315
                  _ExtentY        =   7726
                  _Version        =   393216
                  Rows            =   21
                  FixedRows       =   0
                  FixedCols       =   0
                  BackColorSel    =   32768
                  FocusRect       =   0
                  ScrollBars      =   2
                  SelectionMode   =   1
               End
            End
            Begin Threed.SSPanel SSPanel13 
               Height          =   4425
               Left            =   -74940
               TabIndex        =   89
               Top             =   360
               Width           =   11055
               _Version        =   65536
               _ExtentX        =   19500
               _ExtentY        =   7805
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
               Begin MSFlexGridLib.MSFlexGrid grd_Listad_inm 
                  Height          =   4380
                  Left            =   45
                  TabIndex        =   90
                  Top             =   30
                  Width           =   10950
                  _ExtentX        =   19315
                  _ExtentY        =   7726
                  _Version        =   393216
                  Rows            =   21
                  FixedRows       =   0
                  FixedCols       =   0
                  BackColorSel    =   32768
                  FocusRect       =   0
                  ScrollBars      =   2
                  SelectionMode   =   1
               End
            End
            Begin Threed.SSPanel SSPanel14 
               Height          =   4425
               Left            =   -74940
               TabIndex        =   91
               Top             =   360
               Width           =   11055
               _Version        =   65536
               _ExtentX        =   19500
               _ExtentY        =   7805
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
               Begin MSFlexGridLib.MSFlexGrid grd_Listad 
                  Height          =   4380
                  Index           =   0
                  Left            =   45
                  TabIndex        =   92
                  Top             =   30
                  Width           =   10950
                  _ExtentX        =   19315
                  _ExtentY        =   7726
                  _Version        =   393216
                  Rows            =   21
                  FixedRows       =   0
                  FixedCols       =   0
                  BackColorSel    =   32768
                  FocusRect       =   0
                  ScrollBars      =   2
                  SelectionMode   =   1
               End
            End
            Begin Threed.SSPanel SSPanel33 
               Height          =   4425
               Left            =   -74940
               TabIndex        =   93
               Top             =   360
               Width           =   11055
               _Version        =   65536
               _ExtentX        =   19500
               _ExtentY        =   7805
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
               Begin MSFlexGridLib.MSFlexGrid grd_Listad 
                  Height          =   4380
                  Index           =   3
                  Left            =   45
                  TabIndex        =   94
                  Top             =   30
                  Width           =   10950
                  _ExtentX        =   19315
                  _ExtentY        =   7726
                  _Version        =   393216
                  Rows            =   21
                  FixedRows       =   0
                  FixedCols       =   0
                  BackColorSel    =   32768
                  FocusRect       =   0
                  ScrollBars      =   2
                  SelectionMode   =   1
               End
            End
         End
         Begin Threed.SSPanel SSPanel29 
            Height          =   495
            Left            =   60
            TabIndex        =   27
            Top             =   60
            Width           =   11175
            _Version        =   65536
            _ExtentX        =   19711
            _ExtentY        =   873
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
            Begin Threed.SSPanel pnl_NumOpe 
               Height          =   315
               Left            =   1560
               TabIndex        =   28
               Top             =   90
               Width           =   2055
               _Version        =   65536
               _ExtentX        =   3625
               _ExtentY        =   556
               _StockProps     =   15
               ForeColor       =   32768
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   1
               Font3D          =   2
               Alignment       =   1
            End
            Begin Threed.SSPanel pnl_NomCli 
               Height          =   315
               Left            =   5460
               TabIndex        =   29
               Top             =   90
               Width           =   5595
               _Version        =   65536
               _ExtentX        =   9869
               _ExtentY        =   556
               _StockProps     =   15
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   1
            End
            Begin VB.Label Label8 
               Caption         =   "Nombre del Cliente:"
               Height          =   315
               Left            =   3930
               TabIndex        =   31
               Top             =   120
               Width           =   1395
            End
            Begin VB.Label Label9 
               Caption         =   "Nro. Operación:"
               Height          =   315
               Left            =   90
               TabIndex        =   30
               Top             =   120
               Width           =   1245
            End
         End
      End
   End
End
Attribute VB_Name = "frm_Pos_ConCli_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_str_PerMes As String
Dim l_str_PerAno As String
Dim l_int_TipGar As Integer
Dim l_int_MonGar As Integer
Dim l_dbl_MtoGar As Double
Dim l_str_CodSbs As String

Private Sub Form_Load()
   Dim r_arr_Mtz()      As moddat_g_tpo_DatCom
   
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   pnl_NumOpe.Caption = ""
   pnl_NomCli.Caption = ""
   
   'Limpia grids
   Call fs_IniciaGrid
   
   Call modmip_gs_DatNumOpe(moddat_g_str_NumOpe, grd_Listad(1)) 'fs_Buscar
   pnl_NumOpe.Caption = gf_Formato_NumOpe(moddat_g_str_NumOpe)
   pnl_NomCli.Caption = CStr(moddat_g_int_TipDoc) & " - " & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   
   'Buscar Información del Cliente
   moddat_g_int_CygTDo = 0
   moddat_g_str_CygNDo = ""
   
   l_str_CodSbs = ""
   Call fs_DatCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc)
   Call modmip_gs_DatCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc, grd_Listad(0), 0) 'Buscar Información del Cliente
   Call modmip_gs_DatCli(moddat_g_int_CygTDo, moddat_g_str_CygNDo, grd_Listad(3), 1) 'Buscar Información del Cónyuge
   
  'Buscar Datos del Inmueble*****************
   Call modmip_gs_DatInm(grd_Listad_inm, True)
   Call fs_DatHip
   Call fs_HistCli
   'Call fs_DatCre 'modmip_gs_DatNumOpe(moddat_g_str_NumOpe, grd_Listad(2))
      
   Call modmip_gs_DatCre(grd_Listad(2), r_arr_Mtz)
   Call fs_GenRcc
   Call fs_Buscar_LisExc
   Call fs_Buscar_LisCon
   Call fs_Buscar_Cuotas
   
   Call grd_Listad_rcc1_SelChange

   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Buscar2()
   Dim r_str_CodPry     As String
   Dim r_str_NomPry     As String
   Dim r_str_CodBco     As String
   
   'Buscando Información del Crédito
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_HIPMAE "
   g_str_Parame = g_str_Parame & " WHERE HIPMAE_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   g_str_Parame = g_str_Parame & "   AND (HIPMAE_SITUAC = 2 OR HIPMAE_SITUAC = 3 OR HIPMAE_SITUAC = 6 OR HIPMAE_SITUAC = 7 OR HIPMAE_SITUAC = 9)"
   
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
   moddat_g_str_NumDoc = CStr(Trim(g_rst_Princi!HIPMAE_NDOCLI))
   moddat_g_str_NumSol = Trim(g_rst_Princi!hipmae_numsol)
   moddat_g_str_NumOpe = Trim(g_rst_Princi!HIPMAE_NUMOPE)
   moddat_g_str_CodSub = Trim(g_rst_Princi!HIPMAE_CODSUB)
   moddat_g_str_CodPrd = Trim(g_rst_Princi!HIPMAE_CODPRD)
   moddat_g_int_TipMon = g_rst_Princi!HIPMAE_MONEDA
   l_int_TipGar = g_rst_Princi!HIPMAE_TIPGAR
   l_int_MonGar = g_rst_Princi!HIPMAE_MONGAR
   l_dbl_MtoGar = g_rst_Princi!HIPMAE_MTOGAR
   
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
End Sub

Private Sub fs_Buscar()
   Dim r_str_CodPry     As String
   Dim r_str_NomPry     As String
   Dim r_str_CodBco     As String
   
   'Buscando Información del Crédito
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPMAE "
   g_str_Parame = g_str_Parame & " INNER JOIN CRE_SOLMAE ON SOLMAE_NUMERO = HIPMAE_NUMSOL "
   g_str_Parame = g_str_Parame & " WHERE HIPMAE_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   g_str_Parame = g_str_Parame & "   AND (HIPMAE_SITUAC = 2 OR HIPMAE_SITUAC = 3 OR HIPMAE_SITUAC = 6 OR HIPMAE_SITUAC = 7 OR HIPMAE_SITUAC = 9)"
   
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
   moddat_g_str_NumSol = Trim(g_rst_Princi!hipmae_numsol)
   moddat_g_str_NumOpe = Trim(g_rst_Princi!HIPMAE_NUMOPE)
   l_int_TipGar = g_rst_Princi!HIPMAE_TIPGAR
   l_int_MonGar = g_rst_Princi!HIPMAE_MONGAR
   l_dbl_MtoGar = g_rst_Princi!HIPMAE_MTOGAR
   
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
   moddat_g_int_TipMon = g_rst_Princi!HIPMAE_MONEDA
   
   moddat_g_dbl_MtoPre = g_rst_Princi!HIPMAE_MTOPRE                  'Monto Préstamo
   moddat_g_int_CuoPen = g_rst_Princi!HIPMAE_CUOPEN                  'Cuotas Pendientes
   moddat_g_int_TotCuo = g_rst_Princi!HIPMAE_NUMCUO                  'Total de Cuotas
   moddat_g_dbl_SalCap = g_rst_Princi!HIPMAE_SALCAP                  'Saldo Capital
   moddat_g_str_FecApr = gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECDES))
   
   'Situación de Crédito
   moddat_g_int_Situac = g_rst_Princi!HIPMAE_SITUAC
   moddat_g_str_Situac = moddat_gf_Consulta_ParDes("027", CStr(g_rst_Princi!HIPMAE_SITUAC))
   
   'Obteniendo Información del Inmueble
   Call moddat_gs_Consulta_DatInm(moddat_g_str_NumSol, moddat_g_str_Direcc, moddat_g_str_Distri, r_str_CodPry, r_str_NomPry, r_str_CodBco)
  
   'Cargando en Grid
   grd_Listad(1).Rows = grd_Listad(1).Rows + 1
   grd_Listad(1).Row = grd_Listad(1).Rows - 1
   grd_Listad(1).Col = 0
   grd_Listad(1).CellFontBold = True
   grd_Listad(1).Text = "Número de Operación"
   
   grd_Listad(1).Col = 1
   grd_Listad(1).CellFontBold = True
   grd_Listad(1).Text = gf_Formato_NumOpe(g_rst_Princi!HIPMAE_NUMOPE)
   
   grd_Listad(1).Rows = grd_Listad(1).Rows + 1
   grd_Listad(1).Row = grd_Listad(1).Rows - 1
   grd_Listad(1).Col = 0
   grd_Listad(1).CellFontBold = True
   grd_Listad(1).Text = "Situación"
   
   grd_Listad(1).Col = 1
   grd_Listad(1).CellFontBold = True
   If moddat_g_int_Situac = 6 Then
      grd_Listad(1).Text = moddat_g_str_Situac & "    -    FECHA : " & gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECCAN))
   Else
      grd_Listad(1).Text = moddat_g_str_Situac
   End If
   
   grd_Listad(1).Rows = grd_Listad(1).Rows + 1
   grd_Listad(1).Row = grd_Listad(1).Rows - 1
   grd_Listad(1).Col = 0
   grd_Listad(1).CellFontBold = True
   grd_Listad(1).Text = "Cliente"
   
   grd_Listad(1).Col = 1
   grd_Listad(1).CellFontBold = True
   grd_Listad(1).Text = CStr(g_rst_Princi!HIPMAE_TDOCLI) & " - " & Trim(g_rst_Princi!HIPMAE_NDOCLI) & " / " & moddat_g_str_NomCli
   
   If g_rst_Princi!HIPMAE_TDOCYG > 0 Then
      grd_Listad(1).Rows = grd_Listad(1).Rows + 1
      grd_Listad(1).Row = grd_Listad(1).Rows - 1
      grd_Listad(1).Col = 0
      grd_Listad(1).Text = "Cónyuge"
      
      grd_Listad(1).Col = 1
      grd_Listad(1).Text = CStr(g_rst_Princi!HIPMAE_TDOCYG) & " - " & Trim(g_rst_Princi!HIPMAE_NDOCYG) & " / " & moddat_g_str_CygNom
   End If
   
   grd_Listad(1).Rows = grd_Listad(1).Rows + 1
   grd_Listad(1).Row = grd_Listad(1).Rows - 1
   grd_Listad(1).Col = 0
   grd_Listad(1).Text = "Producto"
   
   grd_Listad(1).Col = 1
   grd_Listad(1).Text = moddat_g_str_NomPrd & " / " & moddat_gf_Consulta_SubPrd(g_rst_Princi!HIPMAE_CODPRD, g_rst_Princi!HIPMAE_CODSUB)
   
   grd_Listad(1).Rows = grd_Listad(1).Rows + 1
   grd_Listad(1).Row = grd_Listad(1).Rows - 1
   grd_Listad(1).Col = 0
   grd_Listad(1).Text = "Moneda Préstamo"
   
   grd_Listad(1).Col = 1
   grd_Listad(1).Text = moddat_g_str_Moneda
   
   grd_Listad(1).Rows = grd_Listad(1).Rows + 2
   grd_Listad(1).Row = grd_Listad(1).Rows - 1
   grd_Listad(1).Col = 0
   grd_Listad(1).Text = "Primera Vivienda"
   
   grd_Listad(1).Col = 1
   grd_Listad(1).Text = moddat_gf_Consulta_ParDes("214", g_rst_Princi!HIPMAE_PRIVIV)
   
   grd_Listad(1).Rows = grd_Listad(1).Rows + 1
   grd_Listad(1).Row = grd_Listad(1).Rows - 1
   grd_Listad(1).Col = 0
   grd_Listad(1).Text = "Modalidad"
   
   grd_Listad(1).Col = 1
   grd_Listad(1).Text = moddat_g_str_DesMod
   
   grd_Listad(1).Rows = grd_Listad(1).Rows + 1
   grd_Listad(1).Row = grd_Listad(1).Rows - 1
   grd_Listad(1).Col = 0
   grd_Listad(1).Text = "Dirección Inmueble"
   
   grd_Listad(1).Col = 1
   grd_Listad(1).Text = moddat_g_str_Direcc
   
   grd_Listad(1).Rows = grd_Listad(1).Rows + 1
   grd_Listad(1).Row = grd_Listad(1).Rows - 1
   grd_Listad(1).Col = 0
   grd_Listad(1).Text = "Distrito"
   
   grd_Listad(1).Col = 1
   grd_Listad(1).Text = moddat_g_str_Distri
   
   If g_rst_Princi!HIPMAE_PRYMCS = 1 Or (g_rst_Princi!HIPMAE_PRYMCS = 2 And CInt(g_rst_Princi!HIPMAE_CODMOD) = 2 Or CInt(g_rst_Princi!HIPMAE_CODMOD) = 3) Then
      grd_Listad(1).Rows = grd_Listad(1).Rows + 1
      grd_Listad(1).Row = grd_Listad(1).Rows - 1
      grd_Listad(1).Col = 0
      grd_Listad(1).Text = "Proyecto Inmobiliario"
      
      grd_Listad(1).Col = 1
      grd_Listad(1).Text = moddat_gf_Consulta_NomPry(g_rst_Princi!HIPMAE_PRYINM & "")
      
      If g_rst_Princi!HIPMAE_PRYMCS = 2 Then
         grd_Listad(1).Text = grd_Listad(1).Text & " (" & moddat_gf_Consulta_ParDes("513", r_str_CodBco) & ")"
      End If
   End If
   
   If moddat_g_int_TipMon = 1 Then
      grd_Listad(1).Rows = grd_Listad(1).Rows + 2
      grd_Listad(1).Row = grd_Listad(1).Rows - 1
      grd_Listad(1).Col = 0
      grd_Listad(1).Text = "Valor Compra Venta"
      
      grd_Listad(1).Col = 1
      grd_Listad(1).CellFontName = "Lucida Console"
      grd_Listad(1).CellFontSize = 8
      grd_Listad(1).Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_CVTSOL, 12, 2)
   
      grd_Listad(1).Rows = grd_Listad(1).Rows + 1
      grd_Listad(1).Row = grd_Listad(1).Rows - 1
      grd_Listad(1).Col = 0
      grd_Listad(1).Text = "Aporte Propio"
      
      grd_Listad(1).Col = 1
      grd_Listad(1).CellFontName = "Lucida Console"
      grd_Listad(1).CellFontSize = 8
      
      'If moddat_g_str_CodPrd = "021" Or moddat_g_str_CodPrd = "022" Or moddat_g_str_CodPrd = "023" Then
      If InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd) > 0 And moddat_g_str_CodPrd <> "019" Then
         grd_Listad(1).Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_APOSOL, 12, 2) & "  (INCLUYE BBP " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & Format(g_rst_Princi!SOLMAE_FMVBBP, "##,###,##0.00") & ") "
      Else
         grd_Listad(1).Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_APOSOL, 12, 2)
      End If
   Else
      grd_Listad(1).Rows = grd_Listad(1).Rows + 2
      grd_Listad(1).Row = grd_Listad(1).Rows - 1
      grd_Listad(1).Col = 0
      grd_Listad(1).Text = "Valor Compra Venta"
      
      grd_Listad(1).Col = 1
      grd_Listad(1).CellFontName = "Lucida Console"
      grd_Listad(1).CellFontSize = 8
      grd_Listad(1).Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_CVTDOL, 12, 2)
   
      grd_Listad(1).Rows = grd_Listad(1).Rows + 1
      grd_Listad(1).Row = grd_Listad(1).Rows - 1
      grd_Listad(1).Col = 0
      grd_Listad(1).Text = "Aporte Propio"
      
      grd_Listad(1).Col = 1
      grd_Listad(1).CellFontName = "Lucida Console"
      grd_Listad(1).CellFontSize = 8
      grd_Listad(1).Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_APODOL, 12, 2)
   End If
   
   grd_Listad(1).Rows = grd_Listad(1).Rows + 1
   grd_Listad(1).Row = grd_Listad(1).Rows - 1
   grd_Listad(1).Col = 0
   grd_Listad(1).Text = "Monto Desembolsado"
   
   grd_Listad(1).Col = 1
   grd_Listad(1).CellFontName = "Lucida Console"
   grd_Listad(1).CellFontSize = 8
   grd_Listad(1).Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_IMPDES, 12, 2)
   
   grd_Listad(1).Rows = grd_Listad(1).Rows + 2
   grd_Listad(1).Row = grd_Listad(1).Rows - 1
   grd_Listad(1).Col = 0
   grd_Listad(1).Text = "Monto Préstamo"
   
   grd_Listad(1).Col = 1
   grd_Listad(1).CellFontName = "Lucida Console"
   grd_Listad(1).CellFontSize = 8
   grd_Listad(1).Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_MTOPRE, 12, 2)
   
   grd_Listad(1).Rows = grd_Listad(1).Rows + 1
   grd_Listad(1).Row = grd_Listad(1).Rows - 1
   grd_Listad(1).Col = 0
   grd_Listad(1).Text = "Interés Capitalizado"
   
   grd_Listad(1).Col = 1
   grd_Listad(1).CellFontName = "Lucida Console"
   grd_Listad(1).CellFontSize = 8
   grd_Listad(1).Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_INTCAP, 12, 2)
   
   grd_Listad(1).Rows = grd_Listad(1).Rows + 1
   grd_Listad(1).Row = grd_Listad(1).Rows - 1
   grd_Listad(1).Col = 0
   grd_Listad(1).Text = "Total Préstamo"
   
   grd_Listad(1).Col = 1
   grd_Listad(1).CellFontName = "Lucida Console"
   grd_Listad(1).CellFontSize = 8
   grd_Listad(1).Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_TOTPRE, 12, 2)
   
   grd_Listad(1).Rows = grd_Listad(1).Rows + 2
   grd_Listad(1).Row = grd_Listad(1).Rows - 1
   grd_Listad(1).Col = 0
   grd_Listad(1).Text = "Fecha Activación"
   
   grd_Listad(1).Col = 1
   grd_Listad(1).Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECACT))
   
   grd_Listad(1).Rows = grd_Listad(1).Rows + 1
   grd_Listad(1).Row = grd_Listad(1).Rows - 1
   grd_Listad(1).Col = 0
   grd_Listad(1).Text = "Fecha Desembolso"
   
   grd_Listad(1).Col = 1
   grd_Listad(1).Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECDES))
   
   If g_rst_Princi!HIPMAE_FECESC > 0 Then
      grd_Listad(1).Rows = grd_Listad(1).Rows + 1
      grd_Listad(1).Row = grd_Listad(1).Rows - 1
      grd_Listad(1).Col = 0
      grd_Listad(1).Text = "Fecha Firma EE.PP"
      
      grd_Listad(1).Col = 1
      grd_Listad(1).Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECESC))
   End If
   
   If moddat_g_str_CodPrd <> "002" Then
      grd_Listad(1).Rows = grd_Listad(1).Rows + 2
      grd_Listad(1).Row = grd_Listad(1).Rows - 1
      grd_Listad(1).Col = 0
      
      Select Case moddat_g_str_CodPrd > 0
         Case InStr(moddat_g_str_AgrCRC, moddat_g_str_CodPrd): grd_Listad(1).Text = "Nro. Operación Mivivienda"   '"001"
         Case InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd): grd_Listad(1).Text = "Nro. Operación COFIDE"       '"003"
         Case InStr(moddat_g_str_AgrTFMV, moddat_g_str_CodPrd): grd_Listad(1).Text = "Nro. Operación COFIDE"      '"004", "007", "009", "010", "013", "014", "015", "016", "017", "018", "019", "020", "021", "022", "023"
      End Select
      
      grd_Listad(1).Col = 1
      grd_Listad(1).Text = Trim(g_rst_Princi!HIPMAE_OPEMVI & "")
      
      If InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Then 'moddat_g_str_CodPrd = "003" Then
         grd_Listad(1).Rows = grd_Listad(1).Rows + 1
         grd_Listad(1).Row = grd_Listad(1).Rows - 1
         grd_Listad(1).Col = 0
         grd_Listad(1).Text = "Nro. Operación Mivivienda"
         
         grd_Listad(1).Col = 1
         grd_Listad(1).Text = Trim(g_rst_Princi!HIPMAE_OPEMV1 & "")
      End If
      
      grd_Listad(1).Rows = grd_Listad(1).Rows + 1
      grd_Listad(1).Row = grd_Listad(1).Rows - 1
      grd_Listad(1).Col = 0
      grd_Listad(1).Text = "Monto Préstamo (Tramo No Conces.)"
      
      grd_Listad(1).Col = 1
      grd_Listad(1).CellFontName = "Lucida Console"
      grd_Listad(1).CellFontSize = 8
      grd_Listad(1).Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_IMPNCO, 12, 2)
   
      grd_Listad(1).Rows = grd_Listad(1).Rows + 1
      grd_Listad(1).Row = grd_Listad(1).Rows - 1
      grd_Listad(1).Col = 0
      grd_Listad(1).Text = "Monto Préstamo (Tramo Conces.)"
      
      grd_Listad(1).Col = 1
      grd_Listad(1).CellFontName = "Lucida Console"
      grd_Listad(1).CellFontSize = 8
      grd_Listad(1).Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_IMPCON, 12, 2)
      
      If InStr(moddat_g_str_AgrCRC, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Then 'moddat_g_str_CodPrd = "001" Or moddat_g_str_CodPrd = "003" Then
         grd_Listad(1).Rows = grd_Listad(1).Rows + 1
         grd_Listad(1).Row = grd_Listad(1).Rows - 1
         grd_Listad(1).Col = 0
         grd_Listad(1).Text = "Tasa de Interés Mivivienda"
      
         grd_Listad(1).Col = 1
         grd_Listad(1).Text = Format(g_rst_Princi!HIPMAE_TASMVI, "##0.00") & " %"
      End If
      
      If InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_AgrTFMV, moddat_g_str_CodPrd) > 0 Then 'moddat_g_str_CodPrd = "004" Or moddat_g_str_CodPrd = "003" Or moddat_g_str_CodPrd = "007" Or moddat_g_str_CodPrd = "009" Or moddat_g_str_CodPrd = "010" Or moddat_g_str_CodPrd = "013" Or moddat_g_str_CodPrd = "014" Or moddat_g_str_CodPrd = "015" Or moddat_g_str_CodPrd = "016" Or moddat_g_str_CodPrd = "017" Or moddat_g_str_CodPrd = "018" Or moddat_g_str_CodPrd = "019" Or moddat_g_str_CodPrd = "020" Or moddat_g_str_CodPrd = "021" Or moddat_g_str_CodPrd = "022" Or moddat_g_str_CodPrd = "023" Then
         grd_Listad(1).Rows = grd_Listad(1).Rows + 1
         grd_Listad(1).Row = grd_Listad(1).Rows - 1
         grd_Listad(1).Col = 0
         grd_Listad(1).Text = "Tasa de Interés COFIDE"
      
         grd_Listad(1).Col = 1
         grd_Listad(1).Text = Format(g_rst_Princi!HIPMAE_TASCOF, "##0.00") & " %"
         
         grd_Listad(1).Rows = grd_Listad(1).Rows + 1
         grd_Listad(1).Row = grd_Listad(1).Rows - 1
         grd_Listad(1).Col = 0
         grd_Listad(1).Text = "Tasa de Comisión COFIDE"
         
         grd_Listad(1).Col = 1
         grd_Listad(1).Text = Format(g_rst_Princi!HIPMAE_COMCOF, "##0.00") & " %"
      End If
   End If
   
   grd_Listad(1).Rows = grd_Listad(1).Rows + 2
   grd_Listad(1).Row = grd_Listad(1).Rows - 1
   grd_Listad(1).Col = 0
   grd_Listad(1).Text = "Plazo"
   
   grd_Listad(1).Col = 1
   grd_Listad(1).Text = CStr(g_rst_Princi!HIPMAE_PLAANO) & " Años"
   
   grd_Listad(1).Rows = grd_Listad(1).Rows + 1
   grd_Listad(1).Row = grd_Listad(1).Rows - 1
   grd_Listad(1).Col = 0
   grd_Listad(1).Text = "Tasa de Interés"
   
   grd_Listad(1).Col = 1
   grd_Listad(1).Text = Format(g_rst_Princi!HIPMAE_TASINT, "##0.00") & " %"
   
   grd_Listad(1).Rows = grd_Listad(1).Rows + 1
   grd_Listad(1).Row = grd_Listad(1).Rows - 1
   grd_Listad(1).Col = 0
   grd_Listad(1).Text = "Nro. de Cuotas"
   
   grd_Listad(1).Col = 1
   grd_Listad(1).Text = CStr(g_rst_Princi!HIPMAE_NUMCUO)
   
   grd_Listad(1).Rows = grd_Listad(1).Rows + 1
   grd_Listad(1).Row = grd_Listad(1).Rows - 1
   grd_Listad(1).Col = 0
   grd_Listad(1).Text = "Período de Gracia"
   
   grd_Listad(1).Col = 1
   grd_Listad(1).Text = CStr(g_rst_Princi!HIPMAE_PERGRA) & " Meses"
   
   grd_Listad(1).Rows = grd_Listad(1).Rows + 1
   grd_Listad(1).Row = grd_Listad(1).Rows - 1
   grd_Listad(1).Col = 0
   grd_Listad(1).Text = "Cuotas Extraordinarias"
   
   grd_Listad(1).Col = 1
   grd_Listad(1).Text = moddat_gf_Consulta_ParDes("277", CStr(g_rst_Princi!HIPMAE_CUOANO))
   
   grd_Listad(1).Rows = grd_Listad(1).Rows + 1
   grd_Listad(1).Row = grd_Listad(1).Rows - 1
   grd_Listad(1).Col = 0
   grd_Listad(1).Text = "Compañía de Seguros"
   
   grd_Listad(1).Col = 1
   grd_Listad(1).Text = moddat_gf_Consulta_ComSeg(g_rst_Princi!HIPMAE_SEGPRE & "")
   
   grd_Listad(1).Rows = grd_Listad(1).Rows + 1
   grd_Listad(1).Row = grd_Listad(1).Rows - 1
   grd_Listad(1).Col = 0
   grd_Listad(1).Text = "Tipo de Seguro Desg."
   
   grd_Listad(1).Col = 1
   grd_Listad(1).Text = moddat_gf_Consulta_TipSeg(g_rst_Princi!HIPMAE_SEGPRE, g_rst_Princi!HIPMAE_TIPSEG)
   
   grd_Listad(1).Rows = grd_Listad(1).Rows + 2
   grd_Listad(1).Row = grd_Listad(1).Rows - 1
   grd_Listad(1).Col = 0
   grd_Listad(1).Text = "Tipo Garantía"
   
   grd_Listad(1).Col = 1
   grd_Listad(1).Text = moddat_gf_Consulta_ParDes("241", CStr(g_rst_Princi!HIPMAE_TIPGAR))
   
   grd_Listad(1).Rows = grd_Listad(1).Rows + 1
   grd_Listad(1).Row = grd_Listad(1).Rows - 1
   grd_Listad(1).Col = 0
   grd_Listad(1).Text = "Monto Garantía"
   
   grd_Listad(1).Col = 1
   grd_Listad(1).CellFontName = "Lucida Console"
   grd_Listad(1).CellFontSize = 8
   If g_rst_Princi!HIPMAE_MONGAR = 0 Then
      grd_Listad(1).Text = gf_FormatoNumero(g_rst_Princi!HIPMAE_MTOGAR, 12, 2)
   Else
      grd_Listad(1).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!HIPMAE_MONGAR)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_MTOGAR, 12, 2)
   End If
   
   grd_Listad(1).Rows = grd_Listad(1).Rows + 2
   grd_Listad(1).Row = grd_Listad(1).Rows - 1
   grd_Listad(1).Col = 0
   grd_Listad(1).Text = "Saldo Capital"
   
   grd_Listad(1).Col = 1
   grd_Listad(1).CellFontName = "Lucida Console"
   grd_Listad(1).CellFontSize = 8
   grd_Listad(1).Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_SALCAP + g_rst_Princi!HIPMAE_SALCON, 12, 2)
   
   grd_Listad(1).Rows = grd_Listad(1).Rows + 1
   grd_Listad(1).Row = grd_Listad(1).Rows - 1
   grd_Listad(1).Col = 0
   grd_Listad(1).Text = "Cuotas Pendientes de Pago"
   
   grd_Listad(1).Col = 1
   grd_Listad(1).Text = CStr(g_rst_Princi!HIPMAE_CUOPEN)
   
   grd_Listad(1).Rows = grd_Listad(1).Rows + 1
   grd_Listad(1).Row = grd_Listad(1).Rows - 1
   grd_Listad(1).Col = 0
   grd_Listad(1).Text = "Días de Atraso"
   
   grd_Listad(1).Col = 1
   grd_Listad(1).Text = CStr(g_rst_Princi!HIPMAE_DIAMOR) & " Días"
   
   grd_Listad(1).Rows = grd_Listad(1).Rows + 2
   grd_Listad(1).Row = grd_Listad(1).Rows - 1
   grd_Listad(1).Col = 0
   grd_Listad(1).Text = "Saldo Capital (Tramo No Conces.)"
   
   grd_Listad(1).Col = 1
   grd_Listad(1).CellFontName = "Lucida Console"
   grd_Listad(1).CellFontSize = 8
   grd_Listad(1).Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_SALCAP, 12, 2)
   
   grd_Listad(1).Rows = grd_Listad(1).Rows + 1
   grd_Listad(1).Row = grd_Listad(1).Rows - 1
   grd_Listad(1).Col = 0
   grd_Listad(1).Text = "Saldo Capital (Tramo Conces.)"
   
   grd_Listad(1).Col = 1
   grd_Listad(1).CellFontName = "Lucida Console"
   grd_Listad(1).CellFontSize = 8
   grd_Listad(1).Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_SALCON, 12, 2)
   
   grd_Listad(1).Rows = grd_Listad(1).Rows + 2
   grd_Listad(1).Row = grd_Listad(1).Rows - 1
   grd_Listad(1).Col = 0
   grd_Listad(1).Text = "Consejero Hipotecario"
   
   grd_Listad(1).Col = 1
   grd_Listad(1).Text = moddat_g_str_NomConHip
   
   grd_Listad(1).Rows = grd_Listad(1).Rows + 1
   grd_Listad(1).Row = grd_Listad(1).Rows - 1
   grd_Listad(1).Col = 0
   grd_Listad(1).Text = "Ejecutivo de Seguimiento"
   
   grd_Listad(1).Col = 1
   grd_Listad(1).Text = moddat_g_str_NomEjeSeg
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call gs_UbiIniGrid(grd_Listad(1))
End Sub
Private Sub fs_Buscar_ant()
   Dim r_str_CodPry     As String
   Dim r_str_NomPry     As String
   Dim r_str_CodBco     As String
   
   'Buscando Información del Crédito
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPMAE "
   g_str_Parame = g_str_Parame & " INNER JOIN CRE_SOLMAE ON SOLMAE_NUMERO = HIPMAE_NUMSOL "
   g_str_Parame = g_str_Parame & " WHERE HIPMAE_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   g_str_Parame = g_str_Parame & "   AND (HIPMAE_SITUAC = 2 OR HIPMAE_SITUAC = 3 OR HIPMAE_SITUAC = 6 OR HIPMAE_SITUAC = 7 OR HIPMAE_SITUAC = 9)"
   
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
   moddat_g_str_NumSol = Trim(g_rst_Princi!hipmae_numsol)
   moddat_g_str_NumOpe = Trim(g_rst_Princi!HIPMAE_NUMOPE)
   l_int_TipGar = g_rst_Princi!HIPMAE_TIPGAR
   l_int_MonGar = g_rst_Princi!HIPMAE_MONGAR
   l_dbl_MtoGar = g_rst_Princi!HIPMAE_MTOGAR
   
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
   moddat_g_int_TipMon = g_rst_Princi!HIPMAE_MONEDA
   
   moddat_g_dbl_MtoPre = g_rst_Princi!HIPMAE_MTOPRE                  'Monto Préstamo
   moddat_g_int_CuoPen = g_rst_Princi!HIPMAE_CUOPEN                  'Cuotas Pendientes
   moddat_g_int_TotCuo = g_rst_Princi!HIPMAE_NUMCUO                  'Total de Cuotas
   moddat_g_dbl_SalCap = g_rst_Princi!HIPMAE_SALCAP                  'Saldo Capital
   moddat_g_str_FecApr = gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECDES))
   
   'Situación de Crédito
   moddat_g_int_Situac = g_rst_Princi!HIPMAE_SITUAC
   moddat_g_str_Situac = moddat_gf_Consulta_ParDes("027", CStr(g_rst_Princi!HIPMAE_SITUAC))
   
   'Obteniendo Información del Inmueble
   Call moddat_gs_Consulta_DatInm(moddat_g_str_NumSol, moddat_g_str_Direcc, moddat_g_str_Distri, r_str_CodPry, r_str_NomPry, r_str_CodBco)
  
   'Cargando en Grid
   grd_Listad(1).Rows = grd_Listad(1).Rows + 1
   grd_Listad(1).Row = grd_Listad(1).Rows - 1
   grd_Listad(1).Col = 0
   grd_Listad(1).CellFontBold = True
   grd_Listad(1).Text = "Número de Operación"
   
   grd_Listad(1).Col = 1
   grd_Listad(1).CellFontBold = True
   grd_Listad(1).Text = gf_Formato_NumOpe(g_rst_Princi!HIPMAE_NUMOPE)
   
   grd_Listad(1).Rows = grd_Listad(1).Rows + 1
   grd_Listad(1).Row = grd_Listad(1).Rows - 1
   grd_Listad(1).Col = 0
   grd_Listad(1).CellFontBold = True
   grd_Listad(1).Text = "Situación"
   
   grd_Listad(1).Col = 1
   grd_Listad(1).CellFontBold = True
   If moddat_g_int_Situac = 6 Then
      grd_Listad(1).Text = moddat_g_str_Situac & "    -    FECHA : " & gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECCAN))
   Else
      grd_Listad(1).Text = moddat_g_str_Situac
   End If
   
   grd_Listad(1).Rows = grd_Listad(1).Rows + 1
   grd_Listad(1).Row = grd_Listad(1).Rows - 1
   grd_Listad(1).Col = 0
   grd_Listad(1).CellFontBold = True
   grd_Listad(1).Text = "Cliente"
   
   grd_Listad(1).Col = 1
   grd_Listad(1).CellFontBold = True
   grd_Listad(1).Text = CStr(g_rst_Princi!HIPMAE_TDOCLI) & " - " & Trim(g_rst_Princi!HIPMAE_NDOCLI) & " / " & moddat_g_str_NomCli
   
   If g_rst_Princi!HIPMAE_TDOCYG > 0 Then
      grd_Listad(1).Rows = grd_Listad(1).Rows + 1
      grd_Listad(1).Row = grd_Listad(1).Rows - 1
      grd_Listad(1).Col = 0
      grd_Listad(1).Text = "Cónyuge"
      
      grd_Listad(1).Col = 1
      grd_Listad(1).Text = CStr(g_rst_Princi!HIPMAE_TDOCYG) & " - " & Trim(g_rst_Princi!HIPMAE_NDOCYG) & " / " & moddat_g_str_CygNom
   End If
   
   grd_Listad(1).Rows = grd_Listad(1).Rows + 1
   grd_Listad(1).Row = grd_Listad(1).Rows - 1
   grd_Listad(1).Col = 0
   grd_Listad(1).Text = "Producto"
   
   grd_Listad(1).Col = 1
   grd_Listad(1).Text = moddat_g_str_NomPrd & " / " & moddat_gf_Consulta_SubPrd(g_rst_Princi!HIPMAE_CODPRD, g_rst_Princi!HIPMAE_CODSUB)
   
   grd_Listad(1).Rows = grd_Listad(1).Rows + 1
   grd_Listad(1).Row = grd_Listad(1).Rows - 1
   grd_Listad(1).Col = 0
   grd_Listad(1).Text = "Moneda Préstamo"
   
   grd_Listad(1).Col = 1
   grd_Listad(1).Text = moddat_g_str_Moneda
   
   grd_Listad(1).Rows = grd_Listad(1).Rows + 2
   grd_Listad(1).Row = grd_Listad(1).Rows - 1
   grd_Listad(1).Col = 0
   grd_Listad(1).Text = "Primera Vivienda"
   
   grd_Listad(1).Col = 1
   grd_Listad(1).Text = moddat_gf_Consulta_ParDes("214", g_rst_Princi!HIPMAE_PRIVIV)
   
   grd_Listad(1).Rows = grd_Listad(1).Rows + 1
   grd_Listad(1).Row = grd_Listad(1).Rows - 1
   grd_Listad(1).Col = 0
   grd_Listad(1).Text = "Modalidad"
   
   grd_Listad(1).Col = 1
   grd_Listad(1).Text = moddat_g_str_DesMod
   
   grd_Listad(1).Rows = grd_Listad(1).Rows + 1
   grd_Listad(1).Row = grd_Listad(1).Rows - 1
   grd_Listad(1).Col = 0
   grd_Listad(1).Text = "Dirección Inmueble"
   
   grd_Listad(1).Col = 1
   grd_Listad(1).Text = moddat_g_str_Direcc
   
   grd_Listad(1).Rows = grd_Listad(1).Rows + 1
   grd_Listad(1).Row = grd_Listad(1).Rows - 1
   grd_Listad(1).Col = 0
   grd_Listad(1).Text = "Distrito"
   
   grd_Listad(1).Col = 1
   grd_Listad(1).Text = moddat_g_str_Distri
   
   If g_rst_Princi!HIPMAE_PRYMCS = 1 Or (g_rst_Princi!HIPMAE_PRYMCS = 2 And CInt(g_rst_Princi!HIPMAE_CODMOD) = 2 Or CInt(g_rst_Princi!HIPMAE_CODMOD) = 3) Then
      grd_Listad(1).Rows = grd_Listad(1).Rows + 1
      grd_Listad(1).Row = grd_Listad(1).Rows - 1
      grd_Listad(1).Col = 0
      grd_Listad(1).Text = "Proyecto Inmobiliario"
      
      grd_Listad(1).Col = 1
      grd_Listad(1).Text = moddat_gf_Consulta_NomPry(g_rst_Princi!HIPMAE_PRYINM & "")
      
      If g_rst_Princi!HIPMAE_PRYMCS = 2 Then
         grd_Listad(1).Text = grd_Listad(1).Text & " (" & moddat_gf_Consulta_ParDes("513", r_str_CodBco) & ")"
      End If
   End If
   
   If moddat_g_int_TipMon = 1 Then
      grd_Listad(1).Rows = grd_Listad(1).Rows + 2
      grd_Listad(1).Row = grd_Listad(1).Rows - 1
      grd_Listad(1).Col = 0
      grd_Listad(1).Text = "Valor Compra Venta"
      
      grd_Listad(1).Col = 1
      grd_Listad(1).CellFontName = "Lucida Console"
      grd_Listad(1).CellFontSize = 8
      grd_Listad(1).Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_CVTSOL, 12, 2)
   
      grd_Listad(1).Rows = grd_Listad(1).Rows + 1
      grd_Listad(1).Row = grd_Listad(1).Rows - 1
      grd_Listad(1).Col = 0
      grd_Listad(1).Text = "Aporte Propio"
      
      grd_Listad(1).Col = 1
      grd_Listad(1).CellFontName = "Lucida Console"
      grd_Listad(1).CellFontSize = 8
      
      'If moddat_g_str_CodPrd = "021" Or moddat_g_str_CodPrd = "022" Or moddat_g_str_CodPrd = "023" Then
      If InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd) > 0 And moddat_g_str_CodPrd <> "019" Then
         grd_Listad(1).Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_APOSOL, 12, 2) & "  (INCLUYE BBP " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & Format(g_rst_Princi!SOLMAE_FMVBBP, "##,###,##0.00") & ") "
      Else
         grd_Listad(1).Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_APOSOL, 12, 2)
      End If
   Else
      grd_Listad(1).Rows = grd_Listad(1).Rows + 2
      grd_Listad(1).Row = grd_Listad(1).Rows - 1
      grd_Listad(1).Col = 0
      grd_Listad(1).Text = "Valor Compra Venta"
      
      grd_Listad(1).Col = 1
      grd_Listad(1).CellFontName = "Lucida Console"
      grd_Listad(1).CellFontSize = 8
      grd_Listad(1).Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_CVTDOL, 12, 2)
   
      grd_Listad(1).Rows = grd_Listad(1).Rows + 1
      grd_Listad(1).Row = grd_Listad(1).Rows - 1
      grd_Listad(1).Col = 0
      grd_Listad(1).Text = "Aporte Propio"
      
      grd_Listad(1).Col = 1
      grd_Listad(1).CellFontName = "Lucida Console"
      grd_Listad(1).CellFontSize = 8
      grd_Listad(1).Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_APODOL, 12, 2)
   End If
   
   grd_Listad(1).Rows = grd_Listad(1).Rows + 1
   grd_Listad(1).Row = grd_Listad(1).Rows - 1
   grd_Listad(1).Col = 0
   grd_Listad(1).Text = "Monto Desembolsado"
   
   grd_Listad(1).Col = 1
   grd_Listad(1).CellFontName = "Lucida Console"
   grd_Listad(1).CellFontSize = 8
   grd_Listad(1).Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_IMPDES, 12, 2)
   
   grd_Listad(1).Rows = grd_Listad(1).Rows + 2
   grd_Listad(1).Row = grd_Listad(1).Rows - 1
   grd_Listad(1).Col = 0
   grd_Listad(1).Text = "Monto Préstamo"
   
   grd_Listad(1).Col = 1
   grd_Listad(1).CellFontName = "Lucida Console"
   grd_Listad(1).CellFontSize = 8
   grd_Listad(1).Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_MTOPRE, 12, 2)
   
   grd_Listad(1).Rows = grd_Listad(1).Rows + 1
   grd_Listad(1).Row = grd_Listad(1).Rows - 1
   grd_Listad(1).Col = 0
   grd_Listad(1).Text = "Interés Capitalizado"
   
   grd_Listad(1).Col = 1
   grd_Listad(1).CellFontName = "Lucida Console"
   grd_Listad(1).CellFontSize = 8
   grd_Listad(1).Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_INTCAP, 12, 2)
   
   grd_Listad(1).Rows = grd_Listad(1).Rows + 1
   grd_Listad(1).Row = grd_Listad(1).Rows - 1
   grd_Listad(1).Col = 0
   grd_Listad(1).Text = "Total Préstamo"
   
   grd_Listad(1).Col = 1
   grd_Listad(1).CellFontName = "Lucida Console"
   grd_Listad(1).CellFontSize = 8
   grd_Listad(1).Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_TOTPRE, 12, 2)
   
   grd_Listad(1).Rows = grd_Listad(1).Rows + 2
   grd_Listad(1).Row = grd_Listad(1).Rows - 1
   grd_Listad(1).Col = 0
   grd_Listad(1).Text = "Fecha Activación"
   
   grd_Listad(1).Col = 1
   grd_Listad(1).Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECACT))
   
   grd_Listad(1).Rows = grd_Listad(1).Rows + 1
   grd_Listad(1).Row = grd_Listad(1).Rows - 1
   grd_Listad(1).Col = 0
   grd_Listad(1).Text = "Fecha Desembolso"
   
   grd_Listad(1).Col = 1
   grd_Listad(1).Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECDES))
   
   If g_rst_Princi!HIPMAE_FECESC > 0 Then
      grd_Listad(1).Rows = grd_Listad(1).Rows + 1
      grd_Listad(1).Row = grd_Listad(1).Rows - 1
      grd_Listad(1).Col = 0
      grd_Listad(1).Text = "Fecha Firma EE.PP"
      
      grd_Listad(1).Col = 1
      grd_Listad(1).Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECESC))
   End If
   
   If moddat_g_str_CodPrd <> "002" Then
      grd_Listad(1).Rows = grd_Listad(1).Rows + 2
      grd_Listad(1).Row = grd_Listad(1).Rows - 1
      grd_Listad(1).Col = 0
      
      Select Case moddat_g_str_CodPrd > 0
         Case InStr(moddat_g_str_AgrCRC, moddat_g_str_CodPrd): grd_Listad(1).Text = "Nro. Operación Mivivienda"   '"001"
         Case InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd): grd_Listad(1).Text = "Nro. Operación COFIDE"       '"003"
         Case InStr(moddat_g_str_AgrTFMV, moddat_g_str_CodPrd): grd_Listad(1).Text = "Nro. Operación COFIDE"      '"004", "007", "009", "010", "013", "014", "015", "016", "017", "018", "019", "020", "021", "022", "023"
      End Select
      
      grd_Listad(1).Col = 1
      grd_Listad(1).Text = Trim(g_rst_Princi!HIPMAE_OPEMVI & "")
      
      If InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Then 'moddat_g_str_CodPrd = "003" Then
         grd_Listad(1).Rows = grd_Listad(1).Rows + 1
         grd_Listad(1).Row = grd_Listad(1).Rows - 1
         grd_Listad(1).Col = 0
         grd_Listad(1).Text = "Nro. Operación Mivivienda"
         
         grd_Listad(1).Col = 1
         grd_Listad(1).Text = Trim(g_rst_Princi!HIPMAE_OPEMV1 & "")
      End If
      
      grd_Listad(1).Rows = grd_Listad(1).Rows + 1
      grd_Listad(1).Row = grd_Listad(1).Rows - 1
      grd_Listad(1).Col = 0
      grd_Listad(1).Text = "Monto Préstamo (Tramo No Conces.)"
      
      grd_Listad(1).Col = 1
      grd_Listad(1).CellFontName = "Lucida Console"
      grd_Listad(1).CellFontSize = 8
      grd_Listad(1).Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_IMPNCO, 12, 2)
   
      grd_Listad(1).Rows = grd_Listad(1).Rows + 1
      grd_Listad(1).Row = grd_Listad(1).Rows - 1
      grd_Listad(1).Col = 0
      grd_Listad(1).Text = "Monto Préstamo (Tramo Conces.)"
      
      grd_Listad(1).Col = 1
      grd_Listad(1).CellFontName = "Lucida Console"
      grd_Listad(1).CellFontSize = 8
      grd_Listad(1).Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_IMPCON, 12, 2)
      
      If InStr(moddat_g_str_AgrCRC, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Then 'moddat_g_str_CodPrd = "001" Or moddat_g_str_CodPrd = "003" Then
         grd_Listad(1).Rows = grd_Listad(1).Rows + 1
         grd_Listad(1).Row = grd_Listad(1).Rows - 1
         grd_Listad(1).Col = 0
         grd_Listad(1).Text = "Tasa de Interés Mivivienda"
      
         grd_Listad(1).Col = 1
         grd_Listad(1).Text = Format(g_rst_Princi!HIPMAE_TASMVI, "##0.00") & " %"
      End If
      
      If InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_AgrTFMV, moddat_g_str_CodPrd) > 0 Then 'moddat_g_str_CodPrd = "004" Or moddat_g_str_CodPrd = "003" Or moddat_g_str_CodPrd = "007" Or moddat_g_str_CodPrd = "009" Or moddat_g_str_CodPrd = "010" Or moddat_g_str_CodPrd = "013" Or moddat_g_str_CodPrd = "014" Or moddat_g_str_CodPrd = "015" Or moddat_g_str_CodPrd = "016" Or moddat_g_str_CodPrd = "017" Or moddat_g_str_CodPrd = "018" Or moddat_g_str_CodPrd = "019" Or moddat_g_str_CodPrd = "020" Or moddat_g_str_CodPrd = "021" Or moddat_g_str_CodPrd = "022" Or moddat_g_str_CodPrd = "023" Then
         grd_Listad(1).Rows = grd_Listad(1).Rows + 1
         grd_Listad(1).Row = grd_Listad(1).Rows - 1
         grd_Listad(1).Col = 0
         grd_Listad(1).Text = "Tasa de Interés COFIDE"
      
         grd_Listad(1).Col = 1
         grd_Listad(1).Text = Format(g_rst_Princi!HIPMAE_TASCOF, "##0.00") & " %"
         
         grd_Listad(1).Rows = grd_Listad(1).Rows + 1
         grd_Listad(1).Row = grd_Listad(1).Rows - 1
         grd_Listad(1).Col = 0
         grd_Listad(1).Text = "Tasa de Comisión COFIDE"
         
         grd_Listad(1).Col = 1
         grd_Listad(1).Text = Format(g_rst_Princi!HIPMAE_COMCOF, "##0.00") & " %"
      End If
   End If
   
   grd_Listad(1).Rows = grd_Listad(1).Rows + 2
   grd_Listad(1).Row = grd_Listad(1).Rows - 1
   grd_Listad(1).Col = 0
   grd_Listad(1).Text = "Plazo"
   
   grd_Listad(1).Col = 1
   grd_Listad(1).Text = CStr(g_rst_Princi!HIPMAE_PLAANO) & " Años"
   
   grd_Listad(1).Rows = grd_Listad(1).Rows + 1
   grd_Listad(1).Row = grd_Listad(1).Rows - 1
   grd_Listad(1).Col = 0
   grd_Listad(1).Text = "Tasa de Interés"
   
   grd_Listad(1).Col = 1
   grd_Listad(1).Text = Format(g_rst_Princi!HIPMAE_TASINT, "##0.00") & " %"
   
   grd_Listad(1).Rows = grd_Listad(1).Rows + 1
   grd_Listad(1).Row = grd_Listad(1).Rows - 1
   grd_Listad(1).Col = 0
   grd_Listad(1).Text = "Nro. de Cuotas"
   
   grd_Listad(1).Col = 1
   grd_Listad(1).Text = CStr(g_rst_Princi!HIPMAE_NUMCUO)
   
   grd_Listad(1).Rows = grd_Listad(1).Rows + 1
   grd_Listad(1).Row = grd_Listad(1).Rows - 1
   grd_Listad(1).Col = 0
   grd_Listad(1).Text = "Período de Gracia"
   
   grd_Listad(1).Col = 1
   grd_Listad(1).Text = CStr(g_rst_Princi!HIPMAE_PERGRA) & " Meses"
   
   grd_Listad(1).Rows = grd_Listad(1).Rows + 1
   grd_Listad(1).Row = grd_Listad(1).Rows - 1
   grd_Listad(1).Col = 0
   grd_Listad(1).Text = "Cuotas Extraordinarias"
   
   grd_Listad(1).Col = 1
   grd_Listad(1).Text = moddat_gf_Consulta_ParDes("277", CStr(g_rst_Princi!HIPMAE_CUOANO))
   
   grd_Listad(1).Rows = grd_Listad(1).Rows + 1
   grd_Listad(1).Row = grd_Listad(1).Rows - 1
   grd_Listad(1).Col = 0
   grd_Listad(1).Text = "Compañía de Seguros"
   
   grd_Listad(1).Col = 1
   grd_Listad(1).Text = moddat_gf_Consulta_ComSeg(g_rst_Princi!HIPMAE_SEGPRE & "")
   
   grd_Listad(1).Rows = grd_Listad(1).Rows + 1
   grd_Listad(1).Row = grd_Listad(1).Rows - 1
   grd_Listad(1).Col = 0
   grd_Listad(1).Text = "Tipo de Seguro Desg."
   
   grd_Listad(1).Col = 1
   grd_Listad(1).Text = moddat_gf_Consulta_TipSeg(g_rst_Princi!HIPMAE_SEGPRE, g_rst_Princi!HIPMAE_TIPSEG)
   
   grd_Listad(1).Rows = grd_Listad(1).Rows + 2
   grd_Listad(1).Row = grd_Listad(1).Rows - 1
   grd_Listad(1).Col = 0
   grd_Listad(1).Text = "Tipo Garantía"
   
   grd_Listad(1).Col = 1
   grd_Listad(1).Text = moddat_gf_Consulta_ParDes("241", CStr(g_rst_Princi!HIPMAE_TIPGAR))
   
   grd_Listad(1).Rows = grd_Listad(1).Rows + 1
   grd_Listad(1).Row = grd_Listad(1).Rows - 1
   grd_Listad(1).Col = 0
   grd_Listad(1).Text = "Monto Garantía"
   
   grd_Listad(1).Col = 1
   grd_Listad(1).CellFontName = "Lucida Console"
   grd_Listad(1).CellFontSize = 8
   If g_rst_Princi!HIPMAE_MONGAR = 0 Then
      grd_Listad(1).Text = gf_FormatoNumero(g_rst_Princi!HIPMAE_MTOGAR, 12, 2)
   Else
      grd_Listad(1).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!HIPMAE_MONGAR)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_MTOGAR, 12, 2)
   End If
   
   grd_Listad(1).Rows = grd_Listad(1).Rows + 2
   grd_Listad(1).Row = grd_Listad(1).Rows - 1
   grd_Listad(1).Col = 0
   grd_Listad(1).Text = "Saldo Capital"
   
   grd_Listad(1).Col = 1
   grd_Listad(1).CellFontName = "Lucida Console"
   grd_Listad(1).CellFontSize = 8
   grd_Listad(1).Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_SALCAP + g_rst_Princi!HIPMAE_SALCON, 12, 2)
   
   grd_Listad(1).Rows = grd_Listad(1).Rows + 1
   grd_Listad(1).Row = grd_Listad(1).Rows - 1
   grd_Listad(1).Col = 0
   grd_Listad(1).Text = "Cuotas Pendientes de Pago"
   
   grd_Listad(1).Col = 1
   grd_Listad(1).Text = CStr(g_rst_Princi!HIPMAE_CUOPEN)
   
   grd_Listad(1).Rows = grd_Listad(1).Rows + 1
   grd_Listad(1).Row = grd_Listad(1).Rows - 1
   grd_Listad(1).Col = 0
   grd_Listad(1).Text = "Días de Atraso"
   
   grd_Listad(1).Col = 1
   grd_Listad(1).Text = CStr(g_rst_Princi!HIPMAE_DIAMOR) & " Días"
   
   grd_Listad(1).Rows = grd_Listad(1).Rows + 2
   grd_Listad(1).Row = grd_Listad(1).Rows - 1
   grd_Listad(1).Col = 0
   grd_Listad(1).Text = "Saldo Capital (Tramo No Conces.)"
   
   grd_Listad(1).Col = 1
   grd_Listad(1).CellFontName = "Lucida Console"
   grd_Listad(1).CellFontSize = 8
   grd_Listad(1).Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_SALCAP, 12, 2)
   
   grd_Listad(1).Rows = grd_Listad(1).Rows + 1
   grd_Listad(1).Row = grd_Listad(1).Rows - 1
   grd_Listad(1).Col = 0
   grd_Listad(1).Text = "Saldo Capital (Tramo Conces.)"
   
   grd_Listad(1).Col = 1
   grd_Listad(1).CellFontName = "Lucida Console"
   grd_Listad(1).CellFontSize = 8
   grd_Listad(1).Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_SALCON, 12, 2)
   
   grd_Listad(1).Rows = grd_Listad(1).Rows + 2
   grd_Listad(1).Row = grd_Listad(1).Rows - 1
   grd_Listad(1).Col = 0
   grd_Listad(1).Text = "Consejero Hipotecario"
   
   grd_Listad(1).Col = 1
   grd_Listad(1).Text = moddat_g_str_NomConHip
   
   grd_Listad(1).Rows = grd_Listad(1).Rows + 1
   grd_Listad(1).Row = grd_Listad(1).Rows - 1
   grd_Listad(1).Col = 0
   grd_Listad(1).Text = "Ejecutivo de Seguimiento"
   
   grd_Listad(1).Col = 1
   grd_Listad(1).Text = moddat_g_str_NomEjeSeg
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call gs_UbiIniGrid(grd_Listad(1))
End Sub

Private Sub fs_IniciaGrid()
Dim r_int_Contad     As Integer

   'Grid de Resumen
   grd_Listad(1).ColWidth(0) = 3000
   grd_Listad(1).ColWidth(1) = 7940
   grd_Listad(1).ColAlignment(0) = flexAlignLeftCenter
   grd_Listad(1).ColAlignment(1) = flexAlignLeftCenter
   grd_Listad(1).Rows = 0

   'Grid de Cliente
   grd_Listad(0).ColWidth(0) = 3000
   grd_Listad(0).ColWidth(1) = 7940
   grd_Listad(0).ColAlignment(0) = flexAlignLeftCenter
   grd_Listad(0).ColAlignment(1) = flexAlignLeftCenter
   grd_Listad(0).Rows = 0
   
   'Grid de Cónyuge
   grd_Listad(3).ColWidth(0) = 3000
   grd_Listad(3).ColWidth(1) = 7940
   grd_Listad(3).ColAlignment(0) = flexAlignLeftCenter
   grd_Listad(3).ColAlignment(1) = flexAlignLeftCenter
   grd_Listad(3).Rows = 0
   
   'Grid Datos Inmueble
   grd_Listad_inm.ColWidth(0) = 3000
   grd_Listad_inm.ColWidth(1) = 7940
   grd_Listad_inm.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad_inm.ColAlignment(1) = flexAlignLeftCenter
   'Call gs_LimpiaGrid(grd_Listad)
   
   'Grid Datos Garantia
   grd_Listad_gar.ColWidth(0) = 3000
   grd_Listad_gar.ColWidth(1) = 7940
   grd_Listad_gar.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad_gar.ColAlignment(1) = flexAlignLeftCenter
   
   'Inicializando Grid de Datos del Cliente
   grd_Listad_his.ColWidth(0) = 0
   grd_Listad_his.ColWidth(1) = 1480
   grd_Listad_his.ColWidth(2) = 2220
   grd_Listad_his.ColWidth(3) = 3590
   grd_Listad_his.ColWidth(4) = 3280
   grd_Listad_his.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad_his.ColAlignment(2) = flexAlignCenterCenter
   grd_Listad_his.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad_his.ColAlignment(4) = flexAlignCenterCenter
   
   'Datos del Crédito
   grd_Listad(2).ColWidth(0) = 3000:   grd_Listad(2).ColAlignment(0) = flexAlignLeftCenter
   grd_Listad(2).ColWidth(1) = 7940:   grd_Listad(2).ColAlignment(1) = flexAlignLeftCenter
   
   'Inicializando RCC del Cliente
   grd_Listad_rcc1.ColWidth(0) = 0
   grd_Listad_rcc1.ColWidth(1) = 1330
   grd_Listad_rcc1.ColWidth(2) = 1300
   grd_Listad_rcc1.ColWidth(3) = 1300
   grd_Listad_rcc1.ColWidth(4) = 1300
   grd_Listad_rcc1.ColWidth(5) = 1300
   grd_Listad_rcc1.ColWidth(6) = 1300
   grd_Listad_rcc1.ColWidth(7) = 1310
   grd_Listad_rcc1.ColWidth(8) = 1200
   grd_Listad_rcc1.ColAlignment(1) = flexAlignLeftCenter
   
   grd_Listad_rcc2.ColWidth(0) = 0
   grd_Listad_rcc2.ColWidth(1) = 0
   grd_Listad_rcc2.ColWidth(2) = 2360
   grd_Listad_rcc2.ColWidth(3) = 1380
   grd_Listad_rcc2.ColWidth(4) = 3470
   grd_Listad_rcc2.ColWidth(5) = 0
   grd_Listad_rcc2.ColWidth(6) = 1020
   grd_Listad_rcc2.ColWidth(7) = 1030
   grd_Listad_rcc2.ColWidth(8) = 1030
   grd_Listad_rcc2.ColWidth(9) = 0
   grd_Listad_rcc2.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad_rcc2.ColAlignment(3) = flexAlignCenterCenter
   
   'Lista de Excepciones
   grd_LisExc.ColWidth(0) = 1160
   grd_LisExc.ColWidth(1) = 1130
   grd_LisExc.ColWidth(2) = 3250
   grd_LisExc.ColWidth(3) = 5320
   grd_LisExc.ColWidth(4) = 0
   grd_LisExc.ColWidth(5) = 0
   
   grd_LisExc.ColAlignment(0) = flexAlignCenterCenter
   grd_LisExc.ColAlignment(1) = flexAlignCenterCenter
   grd_LisExc.ColAlignment(2) = flexAlignLeftCenter
   grd_LisExc.ColAlignment(3) = flexAlignLeftCenter

   pnl_TipAut.Caption = ""

   'Listado de Aprobaciones Condicionadas
   grd_LisCon.ColWidth(0) = 2715
   grd_LisCon.ColWidth(1) = 6500
   grd_LisCon.ColWidth(2) = 1615
   grd_LisCon.ColWidth(3) = 0
   
   grd_LisCon.ColAlignment(0) = flexAlignLeftCenter
   grd_LisCon.ColAlignment(1) = flexAlignLeftCenter
   grd_LisCon.ColAlignment(2) = flexAlignLeftCenter
   
   'Inicializando Grid de Cuotas
   grd_Cuotas.ColWidth(0) = 830
   grd_Cuotas.ColWidth(1) = 1365
   grd_Cuotas.ColWidth(2) = 1065
   grd_Cuotas.ColWidth(3) = 1750
   grd_Cuotas.ColWidth(4) = 1315
   grd_Cuotas.ColWidth(5) = 1470
   grd_Cuotas.ColWidth(6) = 1470
   grd_Cuotas.ColWidth(7) = 1460
   grd_Cuotas.ColAlignment(0) = flexAlignCenterCenter
   grd_Cuotas.ColAlignment(1) = flexAlignCenterCenter
   grd_Cuotas.ColAlignment(2) = flexAlignCenterCenter
   grd_Cuotas.ColAlignment(3) = flexAlignCenterCenter
   grd_Cuotas.ColAlignment(4) = flexAlignCenterCenter
   grd_Cuotas.ColAlignment(5) = flexAlignRightCenter
   grd_Cuotas.ColAlignment(6) = flexAlignRightCenter
   grd_Cuotas.ColAlignment(7) = flexAlignRightCenter
End Sub

Private Sub fs_DatCli(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String)
   g_str_Parame = "SELECT DATGEN_CODSBS FROM CLI_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_TIPDOC = " & CStr(p_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_NUMDOC = '" & p_NumDoc & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      l_str_CodSbs = IIf(IsNull(g_rst_Princi!DATGEN_CODSBS), "", Trim(g_rst_Princi!DATGEN_CODSBS))
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_DatHip()
   Dim r_dbl_TotHip     As Double

   g_str_Parame = "SELECT * FROM CRE_HIPGAR WHERE "
   g_str_Parame = g_str_Parame & "HIPGAR_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   g_str_Parame = g_str_Parame & "ORDER BY HIPGAR_BIEGAR ASC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   grd_Listad_gar.Rows = 0
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      'En caso este registrada la hipoteca
      g_rst_Princi.MoveFirst
      
      r_dbl_TotHip = 0
      Do While Not g_rst_Princi.EOF
         If Not IsNull(g_rst_Princi!HIPGAR_MTOHIP) Then
            r_dbl_TotHip = r_dbl_TotHip + g_rst_Princi!HIPGAR_MTOHIP
         End If
         g_rst_Princi.MoveNext
      Loop
   
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         If g_rst_Princi!HIPGAR_BIEGAR = 1 Then
            grd_Listad_gar.Rows = grd_Listad_gar.Rows + 1
            grd_Listad_gar.Row = grd_Listad_gar.Rows - 1
            grd_Listad_gar.Col = 0
            grd_Listad_gar.Text = "Sede Registral"
            
            grd_Listad_gar.Col = 1
            grd_Listad_gar.Text = moddat_gf_Consulta_ParDes("511", CStr(g_rst_Princi!HIPGAR_SEDREG & ""))
         
            grd_Listad_gar.Rows = grd_Listad_gar.Rows + 1
            grd_Listad_gar.Row = grd_Listad_gar.Rows - 1
            grd_Listad_gar.Col = 0
            grd_Listad_gar.Text = "Moneda Hipoteca"
            
            grd_Listad_gar.Col = 1
            grd_Listad_gar.Text = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!HIPGAR_TIPMON))
         
            grd_Listad_gar.Rows = grd_Listad_gar.Rows + 1
            grd_Listad_gar.Row = grd_Listad_gar.Rows - 1
            grd_Listad_gar.Col = 0
            grd_Listad_gar.Text = "Total Hipoteca"
            
            grd_Listad_gar.Col = 1
            grd_Listad_gar.CellFontName = "Lucida Console"
            grd_Listad_gar.CellFontSize = 8
            grd_Listad_gar.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(r_dbl_TotHip, 12, 2)
         End If
         
         grd_Listad_gar.Rows = grd_Listad_gar.Rows + 2
         grd_Listad_gar.Row = grd_Listad_gar.Rows - 1
         grd_Listad_gar.Col = 0
         grd_Listad_gar.Text = "Bien en Garantía"
         
         grd_Listad_gar.Col = 1
         grd_Listad_gar.Text = moddat_gf_Consulta_ParDes("030", CStr(g_rst_Princi!HIPGAR_BIEGAR))
         
         grd_Listad_gar.Rows = grd_Listad_gar.Rows + 1
         grd_Listad_gar.Row = grd_Listad_gar.Rows - 1
         grd_Listad_gar.Col = 0
         grd_Listad_gar.Text = "Fecha Presentación"
         
         If Not IsNull(g_rst_Princi!HIPGAR_FECINS) Then
            grd_Listad_gar.Col = 1
            grd_Listad_gar.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPGAR_FECINS))
         End If
         
         grd_Listad_gar.Rows = grd_Listad_gar.Rows + 1
         grd_Listad_gar.Row = grd_Listad_gar.Rows - 1
         grd_Listad_gar.Col = 0
         grd_Listad_gar.Text = "Nro. Presentación"
         
         grd_Listad_gar.Col = 1
         grd_Listad_gar.Text = Trim(g_rst_Princi!HIPGAR_NUMINS & "")
         
         grd_Listad_gar.Rows = grd_Listad_gar.Rows + 1
         grd_Listad_gar.Row = grd_Listad_gar.Rows - 1
         grd_Listad_gar.Col = 0
         grd_Listad_gar.Text = "Fecha Inscripción"
         
         If Not IsNull(g_rst_Princi!HIPGAR_FECCON) Then
            grd_Listad_gar.Col = 1
            grd_Listad_gar.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPGAR_FECCON))
         End If
      
         grd_Listad_gar.Rows = grd_Listad_gar.Rows + 1
         grd_Listad_gar.Row = grd_Listad_gar.Rows - 1
         grd_Listad_gar.Col = 0
         grd_Listad_gar.Text = "Doc. Registral (Inmueble)"
      
         If Not IsNull(g_rst_Princi!HIPGAR_TDOREG) Then
            grd_Listad_gar.Col = 1
            grd_Listad_gar.Text = moddat_gf_Consulta_ParDes("026", g_rst_Princi!HIPGAR_TDOREG)
         
            Select Case g_rst_Princi!HIPGAR_TDOREG
               Case 1, 2
                  grd_Listad_gar.Text = grd_Listad_gar.Text & " NRO. " & Trim(g_rst_Princi!HIPGAR_PARFIC & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!HIPGAR_NUMASI & "")
                  
               Case 3
                  grd_Listad_gar.Text = grd_Listad_gar.Text & " (" & Trim(g_rst_Princi!HIPGAR_NUMTOM & "") & " / " & Trim(g_rst_Princi!HIPGAR_NUMFOJ & "") & " / " & Trim(g_rst_Princi!HIPGAR_NUMLIB & "") & ")"
            End Select
         End If
         
         grd_Listad_gar.Rows = grd_Listad_gar.Rows + 1
         grd_Listad_gar.Row = grd_Listad_gar.Rows - 1
         grd_Listad_gar.Col = 0
         grd_Listad_gar.Text = "Monto Hipotecado"
         
         If Not IsNull(g_rst_Princi!HIPGAR_MTOHIP) Then
            grd_Listad_gar.Col = 1
            grd_Listad_gar.CellFontName = "Lucida Console"
            grd_Listad_gar.CellFontSize = 8
            grd_Listad_gar.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPGAR_MTOHIP, 12, 2)
         End If
         
         DoEvents
         g_rst_Princi.MoveNext
      Loop
   Else
      'En caso no este registrada la hipoteca
      grd_Listad_gar.Rows = grd_Listad_gar.Rows + 1
      grd_Listad_gar.Row = grd_Listad_gar.Rows - 1
      grd_Listad_gar.Col = 0
      grd_Listad_gar.Text = "Tipo de Garantía"
      
      grd_Listad_gar.Col = 1
      grd_Listad_gar.Text = moddat_gf_Consulta_ParDes("241", CStr(l_int_TipGar & ""))
      
      grd_Listad_gar.Rows = grd_Listad_gar.Rows + 1
      grd_Listad_gar.Row = grd_Listad_gar.Rows - 1
      grd_Listad_gar.Col = 0
      grd_Listad_gar.Text = "Moneda de la Garantía"
      
      grd_Listad_gar.Col = 1
      If l_int_MonGar > 0 Then
         grd_Listad_gar.Text = moddat_gf_Consulta_ParDes("204", CStr(l_int_MonGar & ""))
      Else
         grd_Listad_gar.Text = ""
      End If
      
      grd_Listad_gar.Rows = grd_Listad_gar.Rows + 1
      grd_Listad_gar.Row = grd_Listad_gar.Rows - 1
      grd_Listad_gar.Col = 0
      grd_Listad_gar.Text = "Monto de la Garantía"
      
      grd_Listad_gar.Col = 1
      grd_Listad_gar.Text = Format(l_dbl_MtoGar, "###,###,##0.00")
      
      If l_int_TipGar = 2 Then
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & "SELECT EVALEG_FECBLQ_INM "
         g_str_Parame = g_str_Parame & "  FROM TRA_EVALEG "
         g_str_Parame = g_str_Parame & " WHERE EVALEG_NUMSOL = '" & moddat_g_str_NumSol & "' "
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
            Exit Sub
         End If
         
         If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
            grd_Listad_gar.Rows = grd_Listad_gar.Rows + 1
            grd_Listad_gar.Row = grd_Listad_gar.Rows - 1
            grd_Listad_gar.Col = 0
            grd_Listad_gar.Text = "Fecha de Bloqueo"
            
            grd_Listad_gar.Col = 1
            grd_Listad_gar.Text = gf_FormatoFecha(CStr(g_rst_Princi!EVALEG_FECBLQ_INM))
         End If
      End If
      
      If l_int_TipGar = 4 Then
         g_str_Parame = "SELECT HIPDES_TIPDES, HIPDES_CHECGO, HIPDES_BANCGO, HIPDES_NUMFIA, HIPDES_BANFIA FROM CRE_HIPDES WHERE "
         g_str_Parame = g_str_Parame & "HIPDES_NUMOPE = '" & moddat_g_str_NumOpe & "' "
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
            Exit Sub
         End If
         
         If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
            If Len(Trim(g_rst_Princi!HIPDES_NUMFIA & "")) > 0 Then
               grd_Listad_gar.Rows = grd_Listad_gar.Rows + 1
               grd_Listad_gar.Row = grd_Listad_gar.Rows - 1
               grd_Listad_gar.Col = 0
               grd_Listad_gar.Text = "Banco Emisor"
               
               grd_Listad_gar.Col = 1
               grd_Listad_gar.Text = moddat_gf_Consulta_ParDes("505", g_rst_Princi!HIPDES_BANFIA)
            End If
         End If
      End If
   End If
   
   Call gs_UbiIniGrid(grd_Listad_gar)
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_HistCli()
   'Buscando Información del Crédito
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT  HIPCIE_PERANO "
   g_str_Parame = g_str_Parame & ", HIPCIE_PERMES "
   g_str_Parame = g_str_Parame & ", TRIM(TO_CHAR(TO_DATE('2012/'||TRIM(HIPCIE_PERMES)|| '/01', 'yyyy/mm/dd'), 'MONTH','NLS_DATE_LANGUAGE=SPANISH' ) ) AS MES "
   g_str_Parame = g_str_Parame & ", TRIM(C.TIPCLA_DESCRI) AS CLACLI "
   g_str_Parame = g_str_Parame & ", TRIM(A.TIPCLA_DESCRI) AS CLAALI "
   g_str_Parame = g_str_Parame & ", HIPCIE_CLACLI, HIPCIE_CLAALI, HIPCIE_CLAPRV "
   g_str_Parame = g_str_Parame & "FROM CRE_HIPCIE "
   g_str_Parame = g_str_Parame & "INNER JOIN CTB_TIPCLA C ON (C.TIPCLA_CODIGO = HIPCIE_CLACLI) "
   g_str_Parame = g_str_Parame & "INNER JOIN CTB_TIPCLA A ON (A.TIPCLA_CODIGO = HIPCIE_CLAPRV) "
   g_str_Parame = g_str_Parame & "WHERE HIPCIE_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   g_str_Parame = g_str_Parame & "AND C.TIPCLA_TIPCRE = '13' "
   g_str_Parame = g_str_Parame & "AND A.TIPCLA_TIPCRE = '13' "
   g_str_Parame = g_str_Parame & "ORDER BY HIPCIE_PERANO DESC, HIPCIE_PERMES DESC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   grd_Listad_his.Rows = 0
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "El cliente no cuenta con una clasificación.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         grd_Listad_his.Rows = grd_Listad_his.Rows + 1
         grd_Listad_his.Row = grd_Listad_his.Rows - 1
                         
         grd_Listad_his.Col = 1
         grd_Listad_his.Text = CStr(g_rst_Princi!HIPCIE_PERANO)
           
         grd_Listad_his.Col = 2
         grd_Listad_his.Text = CStr(g_rst_Princi!mes)
         
         grd_Listad_his.Col = 3
         grd_Listad_his.Text = CStr(g_rst_Princi!CLACLI)
         
         grd_Listad_his.Col = 4
         grd_Listad_his.Text = CStr(g_rst_Princi!CLAALI)
         g_rst_Princi.MoveNext
      Loop
      
      Call gs_UbiIniGrid(grd_Listad_his)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_DatCre()
   'Buscando Información del Crédito
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPMAE "
   g_str_Parame = g_str_Parame & " INNER JOIN CRE_SOLMAE ON SOLMAE_NUMERO = HIPMAE_NUMSOL "
   g_str_Parame = g_str_Parame & " WHERE HIPMAE_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If

   g_rst_Princi.MoveFirst
   grd_Listad(2).Redraw = False
   
   'Cargando en Grid
   grd_Listad(2).Rows = 0
   grd_Listad(2).Rows = grd_Listad(2).Rows + 1
   grd_Listad(2).Row = grd_Listad(2).Rows - 1
   grd_Listad(2).Col = 0
   grd_Listad(2).Text = "Primera Vivienda"
   
   grd_Listad(2).Col = 1
   grd_Listad(2).Text = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!HIPMAE_PRIVIV))
   
   grd_Listad(2).Rows = grd_Listad(2).Rows + 1
   grd_Listad(2).Row = grd_Listad(2).Rows - 1
   grd_Listad(2).Col = 0
   grd_Listad(2).Text = "Moneda Préstamo"
   
   grd_Listad(2).Col = 1
   grd_Listad(2).Text = moddat_gf_Consulta_ParDes("204", CStr(moddat_g_int_TipMon))
   
   If moddat_g_int_TipMon = 1 Then
      grd_Listad(2).Rows = grd_Listad(2).Rows + 1
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).Text = "Valor Compra Venta"
      
      grd_Listad(2).Col = 1
      grd_Listad(2).CellFontName = "Lucida Console"
      grd_Listad(2).CellFontSize = 8
      grd_Listad(2).Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_CVTSOL, 12, 2)
   
      grd_Listad(2).Rows = grd_Listad(2).Rows + 1
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).Text = "Aporte Propio"
      
      grd_Listad(2).Col = 1
      grd_Listad(2).CellFontName = "Lucida Console"
      grd_Listad(2).CellFontSize = 8
      If g_rst_Princi!HIPMAE_CODPRD = "021" Or g_rst_Princi!HIPMAE_CODPRD = "022" Or g_rst_Princi!HIPMAE_CODPRD = "023" Then
         grd_Listad(2).Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_APOSOL, 12, 2) & "  (INCLUYE BBP " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & Format(g_rst_Princi!SOLMAE_FMVBBP, "##,###,##0.00") & ") "
      Else
         grd_Listad(2).Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_APOSOL, 12, 2)
      End If
   Else
      grd_Listad(2).Rows = grd_Listad(2).Rows + 1
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).Text = "Valor Compra Venta"
      
      grd_Listad(2).Col = 1
      grd_Listad(2).CellFontName = "Lucida Console"
      grd_Listad(2).CellFontSize = 8
      grd_Listad(2).Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_CVTDOL, 12, 2)
   
      grd_Listad(2).Rows = grd_Listad(2).Rows + 1
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).Text = "Aporte Propio"
      
      grd_Listad(2).Col = 1
      grd_Listad(2).CellFontName = "Lucida Console"
      grd_Listad(2).CellFontSize = 8
      grd_Listad(2).Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_APODOL, 12, 2)
   End If
   
   grd_Listad(2).Rows = grd_Listad(2).Rows + 1
   grd_Listad(2).Row = grd_Listad(2).Rows - 1
   grd_Listad(2).Col = 0
   grd_Listad(2).Text = "Monto Préstamo"
   
   grd_Listad(2).Col = 1
   grd_Listad(2).CellFontName = "Lucida Console"
   grd_Listad(2).CellFontSize = 8
   grd_Listad(2).Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_MTOPRE, 12, 2)
   
   If g_rst_Princi!HIPMAE_FECESC > 0 Then
      grd_Listad(2).Rows = grd_Listad(2).Rows + 2
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).Text = "Fecha Firma EE.PP"
      
      grd_Listad(2).Col = 1
      grd_Listad(2).Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECESC))
   End If
   
   grd_Listad(2).Rows = grd_Listad(2).Rows + 2
   grd_Listad(2).Row = grd_Listad(2).Rows - 1
   grd_Listad(2).Col = 0
   grd_Listad(2).Text = "Plazo"
   
   grd_Listad(2).Col = 1
   grd_Listad(2).Text = CStr(g_rst_Princi!HIPMAE_PLAANO) & " Años"
   
   grd_Listad(2).Rows = grd_Listad(2).Rows + 1
   grd_Listad(2).Row = grd_Listad(2).Rows - 1
   grd_Listad(2).Col = 0
   grd_Listad(2).Text = "Tasa de Interés"
   
   grd_Listad(2).Col = 1
   grd_Listad(2).Text = Format(g_rst_Princi!HIPMAE_TASINT, "##0.00") & " %"
   
   grd_Listad(2).Rows = grd_Listad(2).Rows + 1
   grd_Listad(2).Row = grd_Listad(2).Rows - 1
   grd_Listad(2).Col = 0
   grd_Listad(2).Text = "Nro. de Cuotas"
   
   grd_Listad(2).Col = 1
   grd_Listad(2).Text = CStr(g_rst_Princi!HIPMAE_NUMCUO)
   
   grd_Listad(2).Rows = grd_Listad(2).Rows + 1
   grd_Listad(2).Row = grd_Listad(2).Rows - 1
   grd_Listad(2).Col = 0
   grd_Listad(2).Text = "Período de Gracia"
   
   grd_Listad(2).Col = 1
   grd_Listad(2).Text = CStr(g_rst_Princi!HIPMAE_PERGRA) & " Meses"
   
   grd_Listad(2).Rows = grd_Listad(2).Rows + 1
   grd_Listad(2).Row = grd_Listad(2).Rows - 1
   grd_Listad(2).Col = 0
   grd_Listad(2).Text = "Cuotas Extraordinarias"
   
   grd_Listad(2).Col = 1
   grd_Listad(2).Text = moddat_gf_Consulta_ParDes("277", CStr(g_rst_Princi!HIPMAE_CUOANO))
   
   grd_Listad(2).Rows = grd_Listad(2).Rows + 1
   grd_Listad(2).Row = grd_Listad(2).Rows - 1
   grd_Listad(2).Col = 0
   grd_Listad(2).Text = "Día de Pago"
   
   grd_Listad(2).Col = 1
   grd_Listad(2).Text = CStr(g_rst_Princi!HIPMAE_DIAPAG)
   
   grd_Listad(2).Rows = grd_Listad(2).Rows + 1
   grd_Listad(2).Row = grd_Listad(2).Rows - 1
   grd_Listad(2).Col = 0
   grd_Listad(2).Text = "Interes Capitalizado"
   
   grd_Listad(2).Col = 1
   grd_Listad(2).Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_INTCAP, 12, 2)
   
   grd_Listad(2).Rows = grd_Listad(2).Rows + 1
   grd_Listad(2).Row = grd_Listad(2).Rows - 1
   grd_Listad(2).Col = 0
   grd_Listad(2).Text = "Compañía de Seguros"
   
   grd_Listad(2).Col = 1
   grd_Listad(2).Text = moddat_gf_Consulta_ComSeg(g_rst_Princi!HIPMAE_SEGPRE & "")
   
   grd_Listad(2).Rows = grd_Listad(2).Rows + 1
   grd_Listad(2).Row = grd_Listad(2).Rows - 1
   grd_Listad(2).Col = 0
   grd_Listad(2).Text = "Tipo de Seguro Desg."
   
   grd_Listad(2).Col = 1
   grd_Listad(2).Text = moddat_gf_Consulta_TipSeg(g_rst_Princi!HIPMAE_SEGPRE, g_rst_Princi!HIPMAE_TIPSEG)
   
   grd_Listad(2).Rows = grd_Listad(2).Rows + 2
   grd_Listad(2).Row = grd_Listad(2).Rows - 1
   grd_Listad(2).Col = 0
   grd_Listad(2).Text = "Consejero Hipotecario"
   
   grd_Listad(2).Col = 1
   grd_Listad(2).Text = moddat_gf_Buscar_NomEje(g_rst_Princi!HIPMAE_CONHIP)
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   grd_Listad(2).Redraw = True
   Call gs_UbiIniGrid(grd_Listad(2))
End Sub

Private Sub fs_GenRcc()
Dim r_str_auxfch As String
Dim r_int_Filaux  As Integer
Dim r_str_Cadena  As String
Dim r_int_filCol  As Integer
Dim r_dbl_importe As Double
Dim r_str_PerMes  As String
Dim r_str_PerAno  As String
Dim r_str_perio1  As String
Dim r_str_perio2  As String
Dim r_str_perio3  As String
Dim r_str_perio4  As String
Dim r_str_perio5  As String
Dim r_str_perio6  As String

    g_str_Parame = ""
    g_str_Parame = g_str_Parame & "SELECT * "
    g_str_Parame = g_str_Parame & "  FROM (SELECT DISTINCT RCCCAB_PERANO, RCCCAB_PERMES "
    g_str_Parame = g_str_Parame & "          FROM CLI_RCCCAB "
    g_str_Parame = g_str_Parame & "         ORDER BY RCCCAB_PERANO DESC, RCCCAB_PERMES DESC) "
    g_str_Parame = g_str_Parame & " WHERE ROWNUM < 2 "
    g_str_Parame = g_str_Parame & " ORDER BY RCCCAB_PERANO DESC "
      
    If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
    End If
   
    r_str_PerMes = g_rst_Princi!RCCCAB_PERMES
    r_str_PerAno = g_rst_Princi!RCCCAB_PERANO
    r_str_Cadena = "01/" & r_str_PerMes & "/" & r_str_PerAno
    
    pnl_Periodo1.Caption = ""
    pnl_Periodo2.Caption = ""
    pnl_Periodo3.Caption = ""
    pnl_Periodo4.Caption = ""
    pnl_Periodo5.Caption = ""
    pnl_Periodo6.Caption = ""
    
    r_str_perio1 = r_str_Cadena
    r_str_perio2 = DateAdd("m", -1, CDate(r_str_Cadena))
    r_str_perio3 = DateAdd("m", -1, CDate(r_str_perio2))
    r_str_perio4 = DateAdd("m", -1, CDate(r_str_perio3))
    r_str_perio5 = DateAdd("m", -1, CDate(r_str_perio4))
    r_str_perio6 = DateAdd("m", -1, CDate(r_str_perio5))
    
    pnl_Periodo1.Caption = Year(r_str_perio6) & "-" & Format(Month(r_str_perio6), "00")
    pnl_Periodo2.Caption = Year(r_str_perio5) & "-" & Format(Month(r_str_perio5), "00")
    pnl_Periodo3.Caption = Year(r_str_perio4) & "-" & Format(Month(r_str_perio4), "00")
    pnl_Periodo4.Caption = Year(r_str_perio3) & "-" & Format(Month(r_str_perio3), "00")
    pnl_Periodo5.Caption = Year(r_str_perio2) & "-" & Format(Month(r_str_perio2), "00")
    pnl_Periodo6.Caption = Year(r_str_perio1) & "-" & Format(Month(r_str_perio1), "00")

    g_str_Parame = ""
    g_str_Parame = g_str_Parame & "SELECT HIPCAB_TIPDOC  , HIPCAB_DOCIDE, HIPCAB_PERMES, HIPCAB_PERANO, "
    g_str_Parame = g_str_Parame & "       HIPCAB_CODSBS  , HIPCAB_NUMEMP, HIPCAB_DEUNOR, HIPCAB_DEUCPP, "
    g_str_Parame = g_str_Parame & "       HIPCAB_DEUDEF  , HIPCAB_DEUDUD, HIPCAB_DEUPER "
    g_str_Parame = g_str_Parame & "  From RCC_HIPCAB "
    g_str_Parame = g_str_Parame & " WHERE ((HIPCAB_PERANO = '" & Left(pnl_Periodo1.Caption, 4) & "' AND HIPCAB_PERMES = '" & Right(pnl_Periodo1.Caption, 2) & "') OR"
    g_str_Parame = g_str_Parame & "        (HIPCAB_PERANO = '" & Left(pnl_Periodo2.Caption, 4) & "' AND HIPCAB_PERMES = '" & Right(pnl_Periodo2.Caption, 2) & "') OR"
    g_str_Parame = g_str_Parame & "        (HIPCAB_PERANO = '" & Left(pnl_Periodo3.Caption, 4) & "' AND HIPCAB_PERMES = '" & Right(pnl_Periodo3.Caption, 2) & "') OR"
    g_str_Parame = g_str_Parame & "        (HIPCAB_PERANO = '" & Left(pnl_Periodo4.Caption, 4) & "' AND HIPCAB_PERMES = '" & Right(pnl_Periodo4.Caption, 2) & "') OR"
    g_str_Parame = g_str_Parame & "        (HIPCAB_PERANO = '" & Left(pnl_Periodo5.Caption, 4) & "' AND HIPCAB_PERMES = '" & Right(pnl_Periodo5.Caption, 2) & "') OR"
    g_str_Parame = g_str_Parame & "        (HIPCAB_PERANO = '" & Left(pnl_Periodo6.Caption, 4) & "' AND HIPCAB_PERMES = '" & Right(pnl_Periodo6.Caption, 2) & "')) "
    g_str_Parame = g_str_Parame & "   AND HIPCAB_CODSBS = '" & l_str_CodSbs & "' "
    g_str_Parame = g_str_Parame & " order by HIPCAB_PERANO asc, HIPCAB_PERMES DESC"
   
    If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
    End If
   
    grd_Listad_rcc1.Rows = 6
    grd_Listad_rcc1.TextMatrix(0, 1) = "Num. Empresas"
    grd_Listad_rcc1.TextMatrix(1, 1) = "D. Normal"
    grd_Listad_rcc1.TextMatrix(2, 1) = "D. CPP"
    grd_Listad_rcc1.TextMatrix(3, 1) = "D. Deficiente"
    grd_Listad_rcc1.TextMatrix(4, 1) = "D. Dudoso"
    grd_Listad_rcc1.TextMatrix(5, 1) = "D. Perdida"
       
    For r_int_Filaux = 1 To grd_Listad_rcc1.Rows - 1
        grd_Listad_rcc1.TextMatrix(r_int_Filaux, 2) = "0.00"
        grd_Listad_rcc1.TextMatrix(r_int_Filaux, 3) = "0.00"
        grd_Listad_rcc1.TextMatrix(r_int_Filaux, 4) = "0.00"
        grd_Listad_rcc1.TextMatrix(r_int_Filaux, 5) = "0.00"
        grd_Listad_rcc1.TextMatrix(r_int_Filaux, 6) = "0.00"
        grd_Listad_rcc1.TextMatrix(r_int_Filaux, 7) = "0.00"
        grd_Listad_rcc1.TextMatrix(r_int_Filaux, 8) = "0.00"
    Next

    grd_Listad_rcc2.Rows = 0

    r_int_filCol = 0
    r_int_Filaux = 1
    If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
       g_rst_Princi.MoveFirst
       Do While Not g_rst_Princi.EOF
          r_str_Cadena = Trim(g_rst_Princi!HIPCAB_PERANO) & "-" & Format(Trim(g_rst_Princi!HIPCAB_PERMES), "00")
          If pnl_Periodo1.Caption = r_str_Cadena Then       '(r_int_Filaux = 1) Then
              r_int_filCol = 2 '7
          ElseIf pnl_Periodo2.Caption = r_str_Cadena Then   '(r_int_Filaux = 2) Then
              r_int_filCol = 3 '6
          ElseIf pnl_Periodo3.Caption = r_str_Cadena Then   '(r_int_Filaux = 3) Then
              r_int_filCol = 4 '5
          ElseIf pnl_Periodo4.Caption = r_str_Cadena Then   '(r_int_Filaux = 4) Then
              r_int_filCol = 5 '4
          ElseIf pnl_Periodo5.Caption = r_str_Cadena Then   '(r_int_Filaux = 5) Then
              r_int_filCol = 6 '3
          ElseIf pnl_Periodo6.Caption = r_str_Cadena Then   '(r_int_Filaux = 6) Then
              r_int_filCol = 7 '2
          End If
          
          grd_Listad_rcc1.TextMatrix(0, r_int_filCol) = Trim(g_rst_Princi!HIPCAB_NUMEMP)
          
          If (g_rst_Princi!HIPCAB_DEUNOR > 0) Then
              grd_Listad_rcc1.TextMatrix(1, r_int_filCol) = Format(g_rst_Princi!HIPCAB_DEUNOR, "###,###,##0.00")
              Else
              grd_Listad_rcc1.TextMatrix(1, r_int_filCol) = "0.00"
          End If
          If (g_rst_Princi!HIPCAB_DEUCPP > 0) Then
              grd_Listad_rcc1.TextMatrix(2, r_int_filCol) = Format(g_rst_Princi!HIPCAB_DEUCPP, "###,###,##0.00")
              Else
              grd_Listad_rcc1.TextMatrix(2, r_int_filCol) = "0.00"
          End If
          If (g_rst_Princi!HIPCAB_DEUDEF > 0) Then
              grd_Listad_rcc1.TextMatrix(3, r_int_filCol) = Format(g_rst_Princi!HIPCAB_DEUDEF, "###,###,##0.00")
              Else
              grd_Listad_rcc1.TextMatrix(3, r_int_filCol) = "0.00"
          End If
          If (g_rst_Princi!HIPCAB_DEUDUD > 0) Then
              grd_Listad_rcc1.TextMatrix(4, r_int_filCol) = Format(g_rst_Princi!HIPCAB_DEUDUD, "###,###,##0.00")
              Else
              grd_Listad_rcc1.TextMatrix(4, r_int_filCol) = "0.00"
          End If
          If (g_rst_Princi!HIPCAB_DEUPER > 0) Then
              grd_Listad_rcc1.TextMatrix(5, r_int_filCol) = Format(g_rst_Princi!HIPCAB_DEUPER, "###,###,##0.00")
              Else
              grd_Listad_rcc1.TextMatrix(5, r_int_filCol) = "0.00"
          End If
          
          g_rst_Princi.MoveNext
          r_int_Filaux = r_int_Filaux + 1
          DoEvents
       Loop
      
       'totales del resumen
       For r_int_Filaux = 1 To grd_Listad_rcc1.Rows - 1
           pnl_Total1.Caption = CStr(CDbl(pnl_Total1.Caption) + CDbl(grd_Listad_rcc1.TextMatrix(r_int_Filaux, 2)))
           pnl_Total2.Caption = CStr(CDbl(pnl_Total2.Caption) + CDbl(grd_Listad_rcc1.TextMatrix(r_int_Filaux, 3)))
           pnl_Total3.Caption = CStr(CDbl(pnl_Total3.Caption) + CDbl(grd_Listad_rcc1.TextMatrix(r_int_Filaux, 4)))
           pnl_Total4.Caption = CStr(CDbl(pnl_Total4.Caption) + CDbl(grd_Listad_rcc1.TextMatrix(r_int_Filaux, 5)))
           pnl_Total5.Caption = CStr(CDbl(pnl_Total5.Caption) + CDbl(grd_Listad_rcc1.TextMatrix(r_int_Filaux, 6)))
           pnl_Total6.Caption = CStr(CDbl(pnl_Total6.Caption) + CDbl(grd_Listad_rcc1.TextMatrix(r_int_Filaux, 7)))
       Next
       
       'porcentaje
       Dim r_dbl_NumMay As Double
       Dim r_dbl_NumFil As Double
       r_dbl_NumMay = 0
       r_dbl_NumFil = 0
       r_dbl_importe = 0
       For r_int_Filaux = 1 To grd_Listad_rcc1.Rows - 1
           If CDbl(pnl_Total6.Caption) = 0 Then
               r_dbl_importe = 0
           Else
               r_dbl_importe = (CDbl(grd_Listad_rcc1.TextMatrix(r_int_Filaux, 7)) * 100) / CDbl(pnl_Total6.Caption)
               r_dbl_importe = Round(r_dbl_importe, 2)
           End If
           If (r_dbl_importe >= r_dbl_NumMay) Then
               r_dbl_NumFil = r_int_Filaux 'r_dbl_importe
               r_dbl_NumMay = r_dbl_importe
           End If
           grd_Listad_rcc1.TextMatrix(r_int_Filaux, 8) = Format(r_dbl_importe, "###,###,##0.00")
           pnl_Total7.Caption = CStr(CDbl(pnl_Total7.Caption) + r_dbl_importe)
       Next
       
       'ajuste
       r_dbl_importe = 0
       r_dbl_importe = Round(100 - CDbl(pnl_Total7.Caption), 2)
       If (r_dbl_importe <> CDbl(0)) Then
           grd_Listad_rcc1.TextMatrix(r_dbl_NumFil, 8) = r_dbl_importe + grd_Listad_rcc1.TextMatrix(r_dbl_NumFil, 8)
           pnl_Total7.Caption = CStr(CDbl(pnl_Total7.Caption) + r_dbl_importe)
       End If
       
       pnl_Total1.Caption = Format(CDbl(pnl_Total1.Caption), "###,###,##0.00") & " "
       pnl_Total2.Caption = Format(CDbl(pnl_Total2.Caption), "###,###,##0.00") & " "
       pnl_Total3.Caption = Format(CDbl(pnl_Total3.Caption), "###,###,##0.00") & " "
       pnl_Total4.Caption = Format(CDbl(pnl_Total4.Caption), "###,###,##0.00") & " "
       pnl_Total5.Caption = Format(CDbl(pnl_Total5.Caption), "###,###,##0.00") & " "
       pnl_Total6.Caption = Format(CDbl(pnl_Total6.Caption), "###,###,##0.00") & " "
       pnl_Total7.Caption = Format(CDbl(pnl_Total7.Caption), "###,###,##0.00") & " "
      
       grd_Listad_rcc1.Row = 0
       grd_Listad_rcc1.Col = 0
     
       g_str_Parame = ""
       g_str_Parame = g_str_Parame & "SELECT HIPDET_TIPDOC , HIPDET_DOCIDE  , HIPDET_PERMES , HIPDET_PERANO , HIPDET_TIPDEU , HIPDET_DIAATR, "
       g_str_Parame = g_str_Parame & "       HIPDET_CLASIF , HIPDET_MTOSOL, HIPDET_MTODOL, HIPDET_TIPMON, HIPDET_CODEMP "
       g_str_Parame = g_str_Parame & "  From RCC_HIPDET "
       g_str_Parame = g_str_Parame & " WHERE ((HIPDET_PERANO = '" & Left(pnl_Periodo1.Caption, 4) & "' AND HIPDET_PERMES = '" & Right(pnl_Periodo1.Caption, 2) & "') OR "
       g_str_Parame = g_str_Parame & "        (HIPDET_PERANO = '" & Left(pnl_Periodo2.Caption, 4) & "' AND HIPDET_PERMES = '" & Right(pnl_Periodo2.Caption, 2) & "') OR "
       g_str_Parame = g_str_Parame & "        (HIPDET_PERANO = '" & Left(pnl_Periodo3.Caption, 4) & "' AND HIPDET_PERMES = '" & Right(pnl_Periodo3.Caption, 2) & "') OR "
       g_str_Parame = g_str_Parame & "        (HIPDET_PERANO = '" & Left(pnl_Periodo4.Caption, 4) & "' AND HIPDET_PERMES = '" & Right(pnl_Periodo4.Caption, 2) & "') OR "
       g_str_Parame = g_str_Parame & "        (HIPDET_PERANO = '" & Left(pnl_Periodo5.Caption, 4) & "' AND HIPDET_PERMES = '" & Right(pnl_Periodo5.Caption, 2) & "') OR "
       g_str_Parame = g_str_Parame & "        (HIPDET_PERANO = '" & Left(pnl_Periodo6.Caption, 4) & "' AND HIPDET_PERMES = '" & Right(pnl_Periodo6.Caption, 2) & "')) "
       g_str_Parame = g_str_Parame & "   AND HIPDET_TIPDOC = '" & CStr(moddat_g_int_TipDoc) & "' "
       g_str_Parame = g_str_Parame & "   AND HIPDET_DOCIDE = '" & moddat_g_str_NumDoc & "' "
       g_str_Parame = g_str_Parame & "   AND HIPDET_CODEMP <> 240 "
       g_str_Parame = g_str_Parame & " UNION "
       g_str_Parame = g_str_Parame & "SELECT HIPDET_TIPDOC , HIPDET_DOCIDE  , HIPDET_PERMES , HIPDET_PERANO , HIPDET_TIPDEU , 0 AS HIPDET_DIAATR, "
       g_str_Parame = g_str_Parame & "       HIPDET_CLASIF , Sum(HIPDET_MTOSOL), Sum(HIPDET_MTODOL), HIPDET_TIPMON, HIPDET_CODEMP "
       g_str_Parame = g_str_Parame & "  From RCC_HIPDET "
       g_str_Parame = g_str_Parame & " WHERE ((HIPDET_PERANO = '" & Left(pnl_Periodo1.Caption, 4) & "' AND HIPDET_PERMES = '" & Right(pnl_Periodo1.Caption, 2) & "') OR "
       g_str_Parame = g_str_Parame & "        (HIPDET_PERANO = '" & Left(pnl_Periodo2.Caption, 4) & "' AND HIPDET_PERMES = '" & Right(pnl_Periodo2.Caption, 2) & "') OR "
       g_str_Parame = g_str_Parame & "        (HIPDET_PERANO = '" & Left(pnl_Periodo3.Caption, 4) & "' AND HIPDET_PERMES = '" & Right(pnl_Periodo3.Caption, 2) & "') OR "
       g_str_Parame = g_str_Parame & "        (HIPDET_PERANO = '" & Left(pnl_Periodo4.Caption, 4) & "' AND HIPDET_PERMES = '" & Right(pnl_Periodo4.Caption, 2) & "') OR "
       g_str_Parame = g_str_Parame & "        (HIPDET_PERANO = '" & Left(pnl_Periodo5.Caption, 4) & "' AND HIPDET_PERMES = '" & Right(pnl_Periodo5.Caption, 2) & "') OR "
       g_str_Parame = g_str_Parame & "        (HIPDET_PERANO = '" & Left(pnl_Periodo6.Caption, 4) & "' AND HIPDET_PERMES = '" & Right(pnl_Periodo6.Caption, 2) & "')) "
       g_str_Parame = g_str_Parame & "   AND HIPDET_TIPDOC =  '" & CStr(moddat_g_int_TipDoc) & "' "
       g_str_Parame = g_str_Parame & "   AND HIPDET_DOCIDE = '" & moddat_g_str_NumDoc & "' "
       g_str_Parame = g_str_Parame & "   AND HIPDET_CODEMP = 240 "
       g_str_Parame = g_str_Parame & " GROUP BY HIPDET_TIPDOC, HIPDET_DOCIDE, HIPDET_PERMES, HIPDET_PERANO, "
       g_str_Parame = g_str_Parame & "          HIPDET_TIPDEU, HIPDET_CLASIF, HIPDET_TIPMON, HIPDET_CODEMP "
    End If
      
    g_rst_Princi.Close
    Set g_rst_Princi = Nothing
   
    If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
    End If
   
    grd_Listad_rcc2.Rows = 0
   
    If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
       g_rst_Princi.MoveFirst
       Do While Not g_rst_Princi.EOF
          'Buscando datos de la Garantía en Registro de Hipotecas
          grd_Listad_rcc2.Rows = grd_Listad_rcc2.Rows + 1
          grd_Listad_rcc2.Row = grd_Listad_rcc2.Rows - 1

          grd_Listad_rcc2.Col = 1
          grd_Listad_rcc2.Text = Trim(g_rst_Princi!HIPDET_PERANO) & "-" & Format(Trim(g_rst_Princi!HIPDET_PERMES), "00")

          grd_Listad_rcc2.Col = 2
          grd_Listad_rcc2.Text = gf_Buscar_NomEmp(Trim(g_rst_Princi!HIPDET_CODEMP))
          'grd_Listad_rcc2.Text = Trim(g_rst_Princi!HIPDET_TIPDEU)

          grd_Listad_rcc2.Col = 3
          If IsNull(g_rst_Princi!HIPDET_CLASIF) Then
             grd_Listad_rcc2.Text = ""
          Else
             grd_Listad_rcc2.Text = Trim(g_rst_Princi!HIPDET_CLASIF)
          End If

          grd_Listad_rcc2.Col = 4
          grd_Listad_rcc2.Text = Trim(g_rst_Princi!HIPDET_TIPDEU)

          grd_Listad_rcc2.Col = 5
          grd_Listad_rcc2.Text = Trim(g_rst_Princi!HIPDET_TIPMON)

          grd_Listad_rcc2.Col = 6
          grd_Listad_rcc2.Text = Format(g_rst_Princi!HIPDET_MTOSOL, "###,###,##0.00")

          grd_Listad_rcc2.Col = 7
          grd_Listad_rcc2.Text = Format(g_rst_Princi!HIPDET_MTODOL, "###,###,##0.00")
          
          r_dbl_importe = CDbl(IIf(IsNull(g_rst_Princi!HIPDET_MTOSOL) = True, "0.00", g_rst_Princi!HIPDET_MTOSOL)) + _
                          CDbl(IIf(IsNull(g_rst_Princi!HIPDET_MTODOL) = True, "0.00", g_rst_Princi!HIPDET_MTODOL))
          grd_Listad_rcc2.Col = 8
          grd_Listad_rcc2.Text = Format(r_dbl_importe, "###,###,##0.00")
          
          grd_Listad_rcc2.Col = 9
          grd_Listad_rcc2.Text = g_rst_Princi!HIPDET_DIAATR
          
          g_rst_Princi.MoveNext
          DoEvents
       Loop

       grd_Listad_rcc2.Row = 0
       grd_Listad_rcc2.Col = 0
    End If
    
    'Valida sobre endeudamiento del titular y del conyuge
    If moddat_gf_Consulta_SobreEndeudamiento(moddat_g_int_TipDoc, moddat_g_str_NumDoc, r_str_PerMes, r_str_PerAno) = "1" Then
       lbl_endeudado.Caption = "SOBRE ENDEUDADO  SI - TIT"
       If Len(Trim(CStr(moddat_g_int_CygTDo))) > 0 And Len(Trim(moddat_g_str_CygNDo)) > 0 Then
          If moddat_gf_Consulta_SobreEndeudamiento(moddat_g_int_CygTDo, moddat_g_str_CygNDo, r_str_PerMes, r_str_PerAno) = "1" Then
             lbl_endeudado.Caption = "SOBRE ENDEUDADO  SI - AMB"
          End If
       End If
    Else
       If Len(Trim(CStr(moddat_g_int_CygTDo))) > 0 And Len(Trim(moddat_g_str_CygNDo)) > 0 Then
          If moddat_gf_Consulta_SobreEndeudamiento(moddat_g_int_CygTDo, moddat_g_str_CygNDo, r_str_PerMes, r_str_PerAno) = "1" Then
             lbl_endeudado.Caption = "SOBRE ENDEUDADO  SI - CYG"
          Else
             lbl_endeudado.Caption = "SOBRE ENDEUDADO  NO"
          End If
       Else
          lbl_endeudado.Caption = "SOBRE ENDEUDADO  NO"
       End If
    End If
                   
    grd_Listad_rcc2.Col = 2
    grd_Listad_rcc2.Sort = 7
                   
    g_rst_Princi.Close
    Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Buscar_LisExc()
   Dim r_str_FecOcu  As String
   
   grd_LisExc.Rows = 0
   
   g_str_Parame = modgen_gf_Buscar_Excepc
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
     g_rst_Princi.Close
     Set g_rst_Princi = Nothing
     Exit Sub
   End If
   
   grd_LisExc.Redraw = False
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      grd_LisExc.Rows = grd_LisExc.Rows + 1
      grd_LisExc.Row = grd_LisExc.Rows - 1
      
      'Fecha de Excepción
      grd_LisExc.Col = 0
      grd_LisExc.Text = gf_FormatoFecha(CStr(g_rst_Princi!SEGFECCRE))
      
      'Hora de Excepción
      grd_LisExc.Col = 1
      grd_LisExc.Text = gf_FormatoHora(Format(g_rst_Princi!SEGHORCRE, "000000"))
      
      'Instancia
      grd_LisExc.Col = 2
      grd_LisExc.Text = moddat_gf_Consulta_ParDes("002", CStr(g_rst_Princi!SEGEXC_CODINS))
      
      'Descripción Excepción
      grd_LisExc.Col = 3
      grd_LisExc.Text = Trim(g_rst_Princi!SEGEXC_DESCRI & "")
      
      'Tipo Autorización
      grd_LisExc.Col = 4
      grd_LisExc.Text = moddat_gf_Consulta_ParDes("243", CStr(g_rst_Princi!SEGEXC_TIPAUT))
      
      'Motivo de Excepción
      grd_LisExc.Col = 5
      grd_LisExc.Text = Trim(g_rst_Princi!PARDES_DESCRI)
      
      g_rst_Princi.MoveNext
   Loop
   
   grd_LisExc.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Call gs_UbiIniGrid(grd_LisExc)
   grd_LisExc.Row = 0
   grd_LisExc.Col = 0
   
   grd_LisExc.RowSel = 0
   grd_LisExc.ColSel = grd_LisExc.Cols - 1
   
   Call grd_LisExc_Click
End Sub

Private Sub grd_Cuotas_DblClick()
   If grd_Cuotas.Rows = 0 Then
      Exit Sub
   End If
   
   moddat_g_int_NumCuo = 0
   moddat_g_int_TipAut = 0
   
   grd_Cuotas.Col = 0
   moddat_g_int_NumCuo = CInt(grd_Cuotas)
   
   grd_Cuotas.Col = 3
   If CStr(grd_Cuotas) = "PAGADA" Then
      moddat_g_int_TipAut = 1
   End If
   
   Call gs_RefrescaGrid(grd_Cuotas)
   
   frm_Ges_CreHip_03.Show 1
End Sub

Private Sub grd_LisCon_SelChange()
If grd_LisCon.Rows > 2 Then
      grd_LisCon.RowSel = grd_LisCon.Row
   End If
   
   Call grd_LisCon_Click
End Sub

Private Sub grd_LisExc_Click()
   If grd_LisExc.Rows > 0 Then
      grd_LisExc.Col = 4
      pnl_TipAut.Caption = grd_LisExc.Text
        grd_LisExc.Col = 0
        grd_LisExc.ColSel = grd_LisExc.Cols - 1
        grd_LisExc.RowSel = grd_LisExc.Row
   Else
      pnl_TipAut.Caption = ""
   End If
End Sub

Private Sub fs_Buscar_LisCon()
   grd_LisCon.Rows = 0
   
   g_str_Parame = "SELECT * FROM TRA_SEGCON WHERE "
   g_str_Parame = g_str_Parame & "SEGCON_NUMSOL = '" & moddat_g_str_NumSol & "' "
   g_str_Parame = g_str_Parame & "ORDER BY SEGCON_SITUAC ASC, SEGCON_CODINS DESC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   grd_LisCon.Redraw = False
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      grd_LisCon.Rows = grd_LisCon.Rows + 1
      grd_LisCon.Row = grd_LisCon.Rows - 1
      
      'Instancia
      grd_LisCon.Col = 0
      grd_LisCon.Text = moddat_gf_Consulta_ParDes("002", CStr(g_rst_Princi!SEGCON_CODINS))
      
      'Descripción Condiciones
      grd_LisCon.Col = 1
      grd_LisCon.Text = Trim(g_rst_Princi!SEGCON_OBSCON & "")
      
      'Situación
      grd_LisCon.Col = 2
      grd_LisCon.Text = moddat_gf_Consulta_ParDes("244", CStr(g_rst_Princi!SEGCON_SITUAC))
      
      'Descripción Levantamiento Condiciones
      grd_LisCon.Col = 3
      grd_LisCon.Text = Trim(g_rst_Princi!SEGCON_OBSLEV & "")
      
      g_rst_Princi.MoveNext
   Loop
   
   grd_LisCon.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   grd_LisCon.Row = 0
   grd_LisCon.Col = 0
   
   grd_LisCon.RowSel = 0
   grd_LisCon.ColSel = grd_LisCon.Cols - 1
   
   Call grd_LisCon_Click
End Sub

Private Sub grd_LisCon_Click()
   If grd_LisCon.Rows > 0 Then
      grd_LisCon.Col = 3
      txt_LevCon.Text = grd_LisCon.Text
      
      grd_LisCon.Col = 0
      grd_LisCon.ColSel = grd_LisCon.Cols - 1
      grd_LisCon.RowSel = grd_LisCon.Row
   End If
End Sub

Private Sub fs_Buscar_Cuotas()
   Dim r_dbl_Pag_TotCuo    As Double
   Dim r_dbl_Pag_Capita    As Double
   Dim r_dbl_Pag_Intere    As Double
   Dim r_dbl_Pag_SegDes    As Double
   Dim r_dbl_Pag_SegViv    As Double
   Dim r_dbl_Pag_OtrCar    As Double
   Dim r_dbl_Pag_IntMor    As Double
   Dim r_dbl_Pag_IntCom    As Double
   Dim r_dbl_Pag_GasCob    As Double
   Dim r_dbl_Pag_OtrGas    As Double
   Dim r_dbl_Deu_TotCuo    As Double
   Dim r_dbl_Deu_Capita    As Double
   Dim r_dbl_Deu_Intere    As Double
   Dim r_dbl_Deu_SegDes    As Double
   Dim r_dbl_Deu_SegViv    As Double
   Dim r_dbl_Deu_OtrCar    As Double
   Dim r_dbl_Deu_IntMor    As Double
   Dim r_dbl_Deu_IntCom    As Double
   Dim r_dbl_Deu_GasCob    As Double
   Dim r_dbl_Deu_OtrGas    As Double
   Dim r_dbl_Sal_TotCuo    As Double
   Dim r_dbl_Sal_Capita    As Double
   Dim r_dbl_Sal_Intere    As Double
   Dim r_dbl_Sal_SegDes    As Double
   Dim r_dbl_Sal_SegViv    As Double
   Dim r_dbl_Sal_OtrCar    As Double
   Dim r_dbl_Sal_IntMor    As Double
   Dim r_dbl_Sal_IntCom    As Double
   Dim r_dbl_Sal_GasCob    As Double
   Dim r_dbl_Sal_OtrGas    As Double
   Dim r_dbl_Gen_TotDeu    As Double
   Dim r_dbl_Gen_TotPag    As Double
   Dim r_dbl_Gen_TotSal    As Double
   
   Dim r_int_FilPag        As Integer
   
   r_dbl_Gen_TotDeu = 0
   r_dbl_Gen_TotPag = 0
   r_dbl_Gen_TotSal = 0
      
   'Cuotas Vencidas
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT HIPCUO_CAPITA, HIPCUO_CAPBBP, HIPCUO_INTERE, HIPCUO_INTBBP, "
   g_str_Parame = g_str_Parame & "        HIPCUO_DESORG, HIPCUO_VIVORG, HIPCUO_OTRORG, HIPCUO_INTMOR, "
   g_str_Parame = g_str_Parame & "        HIPCUO_INTCOM, HIPCUO_GASCOB, HIPCUO_OTRGAS, HIPCUO_CAPPAG, "
   g_str_Parame = g_str_Parame & "        HIPCUO_CBPPAG, HIPCUO_INTPAG, HIPCUO_IBPPAG, HIPCUO_DESPAG, "
   g_str_Parame = g_str_Parame & "        HIPCUO_VIVPAG, HIPCUO_OTRPAG, HIPCUO_ICOPAG, HIPCUO_IMOPAG, "
   g_str_Parame = g_str_Parame & "        HIPCUO_GCOPAG, HIPCUO_OTGPAG, HIPCUO_FECPAG, HIPCUO_FECVCT, "
   g_str_Parame = g_str_Parame & "        HIPCUO_NUMCUO, HIPCUO_SITUAC, HIPMAE_FECCAN "
   g_str_Parame = g_str_Parame & "   FROM CRE_HIPCUO A INNER JOIN CRE_HIPMAE B ON B.HIPMAE_NUMOPE = A.HIPCUO_NUMOPE "
   g_str_Parame = g_str_Parame & "  WHERE A.HIPCUO_NUMOPE = '" & moddat_g_str_NumOpe & "'  "
   g_str_Parame = g_str_Parame & "    AND A.HIPCUO_TIPCRO = 1 "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   grd_Cuotas.Rows = 0
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_Cuotas.Redraw = False
      g_rst_Princi.MoveFirst
            
      Do While Not g_rst_Princi.EOF
         'A Pagar
         r_dbl_Deu_Capita = CDbl(Format(g_rst_Princi!HIPCUO_CAPITA + g_rst_Princi!HIPCUO_CAPBBP, "###,###,##0.00"))
         r_dbl_Deu_Intere = CDbl(Format(g_rst_Princi!HIPCUO_INTERE + g_rst_Princi!HIPCUO_INTBBP, "###,###,##0.00"))
         r_dbl_Deu_SegDes = CDbl(Format(g_rst_Princi!HIPCUO_DESORG, "###,###,##0.00"))
         r_dbl_Deu_SegViv = CDbl(Format(g_rst_Princi!HIPCUO_VIVORG, "###,###,##0.00"))
         r_dbl_Deu_OtrCar = CDbl(Format(g_rst_Princi!HIPCUO_OTRORG, "###,###,##0.00"))
         r_dbl_Deu_IntMor = CDbl(Format(g_rst_Princi!HIPCUO_INTMOR, "###,###,##0.00"))
         r_dbl_Deu_IntCom = CDbl(Format(g_rst_Princi!HIPCUO_INTCOM, "###,###,##0.00"))
         r_dbl_Deu_GasCob = CDbl(Format(g_rst_Princi!HIPCUO_GASCOB, "###,###,##0.00"))
         r_dbl_Deu_OtrGas = CDbl(Format(g_rst_Princi!HIPCUO_OTRGAS, "###,###,##0.00"))
         
         r_dbl_Deu_TotCuo = 0
         r_dbl_Deu_TotCuo = r_dbl_Deu_TotCuo + r_dbl_Deu_Capita
         r_dbl_Deu_TotCuo = r_dbl_Deu_TotCuo + r_dbl_Deu_Intere
         r_dbl_Deu_TotCuo = r_dbl_Deu_TotCuo + r_dbl_Deu_SegDes
         r_dbl_Deu_TotCuo = r_dbl_Deu_TotCuo + r_dbl_Deu_SegViv
         r_dbl_Deu_TotCuo = r_dbl_Deu_TotCuo + r_dbl_Deu_OtrCar
         r_dbl_Deu_TotCuo = r_dbl_Deu_TotCuo + r_dbl_Deu_IntMor
         r_dbl_Deu_TotCuo = r_dbl_Deu_TotCuo + r_dbl_Deu_IntCom
         r_dbl_Deu_TotCuo = r_dbl_Deu_TotCuo + r_dbl_Deu_GasCob
         r_dbl_Deu_TotCuo = r_dbl_Deu_TotCuo + r_dbl_Deu_OtrGas
         
         'Pagado
         r_dbl_Pag_Capita = CDbl(Format(g_rst_Princi!HIPCUO_CAPPAG + g_rst_Princi!HIPCUO_CBPPAG, "###,###,##0.00"))
         r_dbl_Pag_Intere = CDbl(Format(g_rst_Princi!HIPCUO_INTPAG + g_rst_Princi!HIPCUO_IBPPAG, "###,###,##0.00"))
         r_dbl_Pag_SegDes = CDbl(Format(g_rst_Princi!HIPCUO_DESPAG, "###,###,##0.00"))
         r_dbl_Pag_SegViv = CDbl(Format(g_rst_Princi!HIPCUO_VIVPAG, "###,###,##0.00"))
         r_dbl_Pag_OtrCar = CDbl(Format(g_rst_Princi!HIPCUO_OTRPAG, "###,###,##0.00"))
         r_dbl_Pag_IntCom = CDbl(Format(g_rst_Princi!HIPCUO_ICOPAG, "###,###,##0.00"))
         r_dbl_Pag_IntMor = CDbl(Format(g_rst_Princi!HIPCUO_IMOPAG, "###,###,##0.00"))
         r_dbl_Pag_GasCob = CDbl(Format(g_rst_Princi!HIPCUO_GCOPAG, "###,###,##0.00"))
         r_dbl_Pag_OtrGas = CDbl(Format(g_rst_Princi!HIPCUO_OTGPAG, "###,###,##0.00"))
         
         r_dbl_Pag_TotCuo = 0
         r_dbl_Pag_TotCuo = r_dbl_Pag_TotCuo + r_dbl_Pag_Capita
         r_dbl_Pag_TotCuo = r_dbl_Pag_TotCuo + r_dbl_Pag_Intere
         r_dbl_Pag_TotCuo = r_dbl_Pag_TotCuo + r_dbl_Pag_SegDes
         r_dbl_Pag_TotCuo = r_dbl_Pag_TotCuo + r_dbl_Pag_SegViv
         r_dbl_Pag_TotCuo = r_dbl_Pag_TotCuo + r_dbl_Pag_OtrCar
         r_dbl_Pag_TotCuo = r_dbl_Pag_TotCuo + r_dbl_Pag_IntCom
         r_dbl_Pag_TotCuo = r_dbl_Pag_TotCuo + r_dbl_Pag_IntMor
         r_dbl_Pag_TotCuo = r_dbl_Pag_TotCuo + r_dbl_Pag_GasCob
         r_dbl_Pag_TotCuo = r_dbl_Pag_TotCuo + r_dbl_Pag_OtrGas
         
         'Saldo Pago
         r_dbl_Sal_Capita = r_dbl_Deu_Capita - r_dbl_Pag_Capita
         r_dbl_Sal_Intere = r_dbl_Deu_Intere - r_dbl_Pag_Intere
         r_dbl_Sal_IntCom = r_dbl_Deu_IntCom - r_dbl_Pag_IntCom
         r_dbl_Sal_IntMor = r_dbl_Deu_IntMor - r_dbl_Pag_IntMor
         r_dbl_Sal_GasCob = r_dbl_Deu_GasCob - r_dbl_Pag_GasCob
         r_dbl_Sal_OtrGas = r_dbl_Deu_OtrGas - r_dbl_Pag_OtrGas
         r_dbl_Sal_SegDes = r_dbl_Deu_SegDes - r_dbl_Pag_SegDes
         r_dbl_Sal_SegViv = r_dbl_Deu_SegViv - r_dbl_Pag_SegViv
         r_dbl_Sal_OtrCar = r_dbl_Deu_OtrCar - r_dbl_Pag_OtrCar
         
         'Total Cuota
         r_dbl_Sal_TotCuo = 0
         r_dbl_Sal_TotCuo = r_dbl_Sal_TotCuo + r_dbl_Sal_Capita
         r_dbl_Sal_TotCuo = r_dbl_Sal_TotCuo + r_dbl_Sal_Intere
         r_dbl_Sal_TotCuo = r_dbl_Sal_TotCuo + r_dbl_Sal_SegDes
         r_dbl_Sal_TotCuo = r_dbl_Sal_TotCuo + r_dbl_Sal_SegViv
         r_dbl_Sal_TotCuo = r_dbl_Sal_TotCuo + r_dbl_Sal_OtrCar
         r_dbl_Sal_TotCuo = r_dbl_Sal_TotCuo + r_dbl_Sal_IntCom
         r_dbl_Sal_TotCuo = r_dbl_Sal_TotCuo + r_dbl_Sal_IntMor
         r_dbl_Sal_TotCuo = r_dbl_Sal_TotCuo + r_dbl_Sal_GasCob
         r_dbl_Sal_TotCuo = r_dbl_Sal_TotCuo + r_dbl_Sal_OtrGas
         
         grd_Cuotas.Rows = grd_Cuotas.Rows + 1
         grd_Cuotas.Row = grd_Cuotas.Rows - 1
         
         grd_Cuotas.Col = 0
         grd_Cuotas.Text = Format(g_rst_Princi!HIPCUO_NUMCUO, "000")
         
         grd_Cuotas.Col = 1
         grd_Cuotas.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECVCT))
         
         'Si Situación es No-Pagado
         If g_rst_Princi!HIPCUO_SITUAC = 2 Then
            If moddat_g_int_Situac <> 6 And moddat_g_int_Situac <> 9 Then
               If CDate(gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECVCT))) < CDate(moddat_g_str_FecSis) Then
                  grd_Cuotas.Col = 2
                  grd_Cuotas.Text = CStr(CInt(CDate(moddat_g_str_FecSis) - CDate(gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECVCT)))))
                  
                  grd_Cuotas.Col = 3
                  grd_Cuotas.Text = "VENCIDA"
               Else
                  grd_Cuotas.Col = 2
                  grd_Cuotas.Text = "-"
                  
                  grd_Cuotas.Col = 3
                  grd_Cuotas.Text = "POR VENCER"
               End If
            ElseIf moddat_g_int_Situac = 6 Then
               If CDate(gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECVCT))) <= CDate(gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECCAN))) Then
                  grd_Cuotas.Col = 2
                  'grd_Cuotas.Text = CStr(CInt(CDate(gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECCAN))) - CDate(gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECVCT)))))
                  grd_Cuotas.Text = CStr(CInt(CDate(moddat_g_str_FecSis) - CDate(gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECVCT)))))
                  
                  grd_Cuotas.Col = 3
                  grd_Cuotas.Text = "VENCIDA"
               End If
            End If
         Else
            If CInt(CDate(gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECPAG))) - CDate(gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECVCT)))) > 0 Then
               grd_Cuotas.Col = 2
               grd_Cuotas.Text = CStr(CInt(CDate(gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECPAG))) - CDate(gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECVCT)))))
            Else
               grd_Cuotas.Col = 2
               grd_Cuotas.Text = "-"
            End If
            
            grd_Cuotas.Col = 3
            grd_Cuotas.Text = "PAGADA"
            r_int_FilPag = grd_Cuotas.Row
         End If
         
         If g_rst_Princi!HIPCUO_FECPAG > 0 Then
            grd_Cuotas.Col = 4
            grd_Cuotas.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECPAG))
         End If
      
         'Valor Cuota
         grd_Cuotas.Col = 5
         grd_Cuotas.Text = Format(r_dbl_Deu_TotCuo, "###,###,##0.00")
                   
         'Importe Pagado
         grd_Cuotas.Col = 6
         grd_Cuotas.Text = Format(r_dbl_Pag_TotCuo, "###,###,##0.00")
      
         'Saldo
         grd_Cuotas.Col = 7
         grd_Cuotas.Text = Format(r_dbl_Sal_TotCuo, "###,###,##0.00")
      
         'Sumando Totales
         r_dbl_Gen_TotDeu = r_dbl_Gen_TotDeu + r_dbl_Deu_TotCuo
         r_dbl_Gen_TotPag = r_dbl_Gen_TotPag + r_dbl_Pag_TotCuo
         r_dbl_Gen_TotSal = r_dbl_Gen_TotSal + r_dbl_Sal_TotCuo
      
         g_rst_Princi.MoveNext
      Loop
      
      pnl_Cuo_TotDeu.Caption = Format(r_dbl_Gen_TotDeu, "###,###,##0.00") & " "
      pnl_Cuo_TotPag.Caption = Format(r_dbl_Gen_TotPag, "###,###,##0.00") & " "
      pnl_Cuo_TotSal.Caption = Format(r_dbl_Gen_TotSal, "###,###,##0.00") & " "
      
      grd_Cuotas.Redraw = True
      Call gs_UbiIniGrid(grd_Cuotas)
   End If
   
   lbl_Totale.Caption = "Totales ===> " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " "
   
   'Coloca el cursor en la última cuota pagada
   With grd_Cuotas
      .SelectionMode = flexSelectionByRow
      If .Rows > 1 Then
      .Row = r_int_FilPag
      .TopRow = r_int_FilPag
      .RowSel = r_int_FilPag
      .Col = 0
      .ColSel = .Cols - 1
      End If
   End With
  
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Public Function gf_Buscar_NomEmp(ByVal p_CodEmp As Integer) As String
   gf_Buscar_NomEmp = ""
   
   g_str_Parame = "SELECT * FROM CTB_EMPSUP WHERE EMPSUP_CODIGO = " & p_CodEmp & " "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If

   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      gf_Buscar_NomEmp = Trim(g_rst_Listas!EMPSUP_NOMBRE)
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Public Function gf_Buscar_TipCla(ByVal p_CodCla As Integer, ByVal p_CodCre As Integer) As String
   gf_Buscar_TipCla = ""
   
   g_str_Parame = "SELECT * FROM CTB_TIPCLA WHERE TIPCLA_TIPCRE = " & p_CodCre & " AND TIPCLA_CODIGO = " & p_CodCla
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If

   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      gf_Buscar_TipCla = Trim(g_rst_Listas!TIPCLA_DESCRI)
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Public Function gf_Buscar_TipCre(ByVal p_CodCre As Integer) As String
   gf_Buscar_TipCre = ""
   g_str_Parame = "SELECT * FROM CTB_TIPCRE WHERE TIPCRE_CODIGO = " & p_CodCre
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Function
   End If

   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      gf_Buscar_TipCre = Trim(g_rst_Listas!TIPCRE_DESCRI)
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Function

Private Sub cmd_EstCta_Click()
   frm_Ges_CreHip_10.Show 1
End Sub

Private Sub cmd_ImpCro_Click()
   modmip_g_int_OrdAct = 1
   frm_Ges_CreHip_07.Show 1
End Sub

Private Sub cmd_Imprim_Click()
   If MsgBox("¿Está seguro de imprimir el Resumen de Crédito?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   'Borrando Spool (Cabecera)
   g_str_Parame = "DELETE FROM RPT_ECTACB WHERE "
   g_str_Parame = g_str_Parame & "ECTACB_CODTER = '" & modgen_g_str_NombPC & "' AND "
   g_str_Parame = g_str_Parame & "ECTACB_NOMRPT = 'OPE_ESTCTA_02.RPT' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
       Exit Sub
   End If
   
   'Borrando Spool (Detalle)
   g_str_Parame = "DELETE FROM RPT_ECTADT WHERE "
   g_str_Parame = g_str_Parame & "ECTADT_CODTER = '" & modgen_g_str_NombPC & "' AND "
   g_str_Parame = g_str_Parame & "ECTADT_NOMRPT = 'OPE_ESTCTA_02.RPT' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
       Exit Sub
   End If
   
   Call opecaj_gs_ResCuo(moddat_g_str_NumOpe)

   'Se conecta al crystal report
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
   
   'Se envia las tablas correspondientes en el orden que fueron utilizadas
   crp_Imprim.DataFiles(0) = "RPT_ECTACB"
   crp_Imprim.DataFiles(1) = "RPT_ECTADT"
      
   'Se pone la llamada del nombre del reporte y se escoge donde se destinara el reporte
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_ESTCTA_02.RPT"
   crp_Imprim.SelectionFormula = "{RPT_ECTADT.ECTADT_NOMRPT} = 'OPE_ESTCTA_02.RPT' AND "
   crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{RPT_ECTADT.ECTADT_CODTER} = '" & modgen_g_str_NombPC & "' "
   crp_Imprim.WindowShowPrintSetupBtn = True
   crp_Imprim.Action = 1
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub cmd_VerPag_Click()
   frm_Ges_CreHip_05.Show 1
End Sub

Private Sub grd_LisExc_SelChange()
   If grd_LisExc.Rows > 2 Then
      grd_LisExc.RowSel = grd_LisExc.Row
   End If
   Call grd_LisExc_Click
End Sub

Private Sub grd_Listad_rcc1_SelChange()
Dim r_str_auxfecha As String
Dim r_int_NroFil   As Integer

    pnl_Detalle.Caption = "Detalle"
    If grd_Listad_rcc1.Rows = 0 Then
       Exit Sub
    End If

    If (grd_Listad_rcc1.Col = 2) Then
        r_str_auxfecha = Trim(pnl_Periodo1.Caption)
    ElseIf (grd_Listad_rcc1.Col = 3) Then
        r_str_auxfecha = Trim(pnl_Periodo2.Caption)
    ElseIf (grd_Listad_rcc1.Col = 4) Then
        r_str_auxfecha = Trim(pnl_Periodo3.Caption)
    ElseIf (grd_Listad_rcc1.Col = 5) Then
        r_str_auxfecha = Trim(pnl_Periodo4.Caption)
    ElseIf (grd_Listad_rcc1.Col = 6) Then
        r_str_auxfecha = Trim(pnl_Periodo5.Caption)
    ElseIf (grd_Listad_rcc1.Col = 7) Then
        r_str_auxfecha = Trim(pnl_Periodo6.Caption)
    End If
    
    pnl_Detalle.Caption = "Detalle Periodo : " & r_str_auxfecha

    For r_int_NroFil = 0 To grd_Listad_rcc2.Rows - 1
       grd_Listad_rcc2.RowHeight(r_int_NroFil) = 0
       If (grd_Listad_rcc2.TextMatrix(r_int_NroFil, 1) = r_str_auxfecha) Then
          grd_Listad_rcc2.RowHeight(r_int_NroFil) = 240
       End If
    Next
    
    'Call gs_RefrescaGrid(grd_Listad_rcc1)
End Sub

Private Sub txt_LevCon_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub cmd_Export_Click()
   'Confirmacion
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
      
   Call fs_GenExc

End Sub

Private Sub fs_GenExc()
Dim r_obj_Excel      As Excel.Application
Dim r_int_FilGrd     As Integer
Dim r_int_FilExl     As Integer
Dim r_int_filCol     As Integer
Dim r_int_fildet     As Integer
Dim r_str_Cadena     As String
Dim r_int_VarAux     As Integer
Dim r_bol_Estado As Boolean
       
   Screen.MousePointer = 11
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
        r_int_VarAux = 2
       .Range("B" & r_int_VarAux) = "REPORTE CONSOLIDADO CREDITICIO"
       .Range("B" & r_int_VarAux & ":I" & r_int_VarAux).Font.Underline = True
       .Range("B" & r_int_VarAux & ":I" & r_int_VarAux).Font.Bold = True
       .Range("B" & r_int_VarAux & ":I" & r_int_VarAux).Font.Size = 8
       .Range("B" & r_int_VarAux & ":I" & r_int_VarAux).Merge
       .Range("B" & r_int_VarAux).HorizontalAlignment = xlHAlignCenter
            
       .Range("A1:J85").Font.Name = "Arial"
       .Range("A3:J85").Font.Size = 8
       '.Rows("1:100").RowHeight = 11.25
      
       .Columns("J").HorizontalAlignment = xlHAlignCenter
       .Columns("A").ColumnWidth = 5
       .Columns("B").ColumnWidth = 14
       .Columns("C").ColumnWidth = 12
       .Columns("D").ColumnWidth = 14
       .Columns("E").ColumnWidth = 14
       .Columns("F").ColumnWidth = 14
       .Columns("G").ColumnWidth = 13
       .Columns("H").ColumnWidth = 13
       .Columns("I").ColumnWidth = 13
       .Columns("J").ColumnWidth = 12
       'Bordes de las celdas
        r_int_VarAux = 8
       .Range("B" & r_int_VarAux).Borders(xlEdgeLeft).LineStyle = xlContinuous
       .Range("I" & r_int_VarAux).Borders(xlEdgeRight).LineStyle = xlContinuous
       .Range("B" & r_int_VarAux & ":I" & r_int_VarAux).Borders(xlEdgeTop).LineStyle = xlContinuous
       .Range("B" & r_int_VarAux & ":I" & r_int_VarAux).Borders(xlEdgeBottom).LineStyle = xlContinuous
       '.Range("B45:F45").Borders(xlEdgeTop).Weight = xlThin
       
       .Range("B" & r_int_VarAux & ":J" & r_int_VarAux).HorizontalAlignment = xlHAlignCenter
       .Range("B9:B14").HorizontalAlignment = xlHAlignLeft
       .Range("C9:J15").HorizontalAlignment = xlHAlignRight
       .Range("B15").HorizontalAlignment = xlHAlignCenter
       .Cells(4, r_int_VarAux).HorizontalAlignment = xlHAlignCenter
              
       .Range("H4:I4").Merge
       '.Range("B8:C8").Merge
       '.Range("B9:C9").Merge
       '.Range("B10:C10").Merge
       '.Range("B11:C11").Merge
       '.Range("B12:C12").Merge
       '.Range("B13:C13").Merge
       '.Range("B14:C14").Merge
       '.Range("B15:C15").Merge
      
       .Cells(4, 2) = "NRO DE OPERACION:"
       .Cells(5, 2) = "CLIENTE:"
       .Cells(4, 4) = Trim(pnl_NumOpe.Caption)
       .Cells(5, 4) = Trim(pnl_NomCli.Caption)
       .Cells(4, 8) = Trim(lbl_endeudado.Caption)
      
        r_int_VarAux = 7
       .Cells(r_int_VarAux, 2) = "RESUMEN"
       .Cells(r_int_VarAux, 2).Font.Bold = True
        r_int_VarAux = 8
       .Cells(r_int_VarAux, 2) = "CLASIF. \ PERIODO"
       .Cells(r_int_VarAux, 3) = Trim(pnl_Periodo1.Caption)
       .Cells(r_int_VarAux, 4) = Trim(pnl_Periodo2.Caption)
       .Cells(r_int_VarAux, 5) = Trim(pnl_Periodo3.Caption)
       .Cells(r_int_VarAux, 6) = Trim(pnl_Periodo4.Caption)
       .Cells(r_int_VarAux, 7) = Trim(pnl_Periodo5.Caption)
       .Cells(r_int_VarAux, 8) = Trim(pnl_Periodo6.Caption)
       .Cells(r_int_VarAux, 9) = "%"
               
       r_int_FilExl = 9
       r_int_filCol = 3
       For r_int_FilGrd = 0 To grd_Listad_rcc1.Rows - 1
           .Cells(r_int_FilExl, 2) = grd_Listad_rcc1.TextMatrix(r_int_FilGrd, 1)
           .Cells(r_int_FilExl, 3) = grd_Listad_rcc1.TextMatrix(r_int_FilGrd, 2)
           .Cells(r_int_FilExl, 4) = grd_Listad_rcc1.TextMatrix(r_int_FilGrd, 3)
           .Cells(r_int_FilExl, 5) = grd_Listad_rcc1.TextMatrix(r_int_FilGrd, 4)
           .Cells(r_int_FilExl, 6) = grd_Listad_rcc1.TextMatrix(r_int_FilGrd, 5)
           .Cells(r_int_FilExl, 7) = grd_Listad_rcc1.TextMatrix(r_int_FilGrd, 6)
           .Cells(r_int_FilExl, 8) = grd_Listad_rcc1.TextMatrix(r_int_FilGrd, 7)
           .Cells(r_int_FilExl, 9) = grd_Listad_rcc1.TextMatrix(r_int_FilGrd, 8)
          
           r_int_FilExl = r_int_FilExl + 1
           r_int_filCol = r_int_filCol + 1
       Next
              
       .Range("C10:J15").NumberFormat = "###,###,##0.00"
       r_int_VarAux = 15
       .Cells(r_int_VarAux, 3) = pnl_Total1.Caption
       .Cells(r_int_VarAux, 4) = pnl_Total2.Caption
       .Cells(r_int_VarAux, 5) = pnl_Total3.Caption
       .Cells(r_int_VarAux, 6) = pnl_Total4.Caption
       .Cells(r_int_VarAux, 7) = pnl_Total5.Caption
       .Cells(r_int_VarAux, 8) = pnl_Total6.Caption
       .Cells(r_int_VarAux, 9) = pnl_Total7.Caption
       .Range("B" & r_int_VarAux & ":J" & r_int_VarAux).Font.Bold = True
       .Cells(r_int_VarAux, 2) = "TOTAL"
       r_int_VarAux = 17
       .Cells(r_int_VarAux, 2) = "DETALLE"
       .Cells(r_int_VarAux, 2).Font.Bold = True
       r_int_VarAux = 18
       .Cells(r_int_VarAux, 2) = "NOMBRE EMPRESA"
       .Cells(r_int_VarAux, 4) = "CLASIFICACION"
       .Cells(r_int_VarAux, 5) = "TIPO DEUDA"
       .Cells(r_int_VarAux, 7) = "MONTO (S/.)"
       .Cells(r_int_VarAux, 8) = "MONTO (US$)"
       .Cells(r_int_VarAux, 9) = "TOTAL (S/.)"
       '.Cells(r_int_VarAux, 10) = "DIAS ATRASO"
      
       'Bordes de las celdas
       .Range("B" & r_int_VarAux).Borders(xlEdgeLeft).LineStyle = xlContinuous
       .Range("I" & r_int_VarAux).Borders(xlEdgeRight).LineStyle = xlContinuous
       .Range("B" & r_int_VarAux & ":I" & r_int_VarAux).Borders(xlEdgeTop).LineStyle = xlContinuous
       .Range("B" & r_int_VarAux & ":I" & r_int_VarAux).Borders(xlEdgeBottom).LineStyle = xlContinuous
       
       .Range("B" & r_int_VarAux & ":C" & r_int_VarAux).Merge
       .Range("E" & r_int_VarAux & ":F" & r_int_VarAux).Merge
       .Range("B" & r_int_VarAux & ":I" & r_int_VarAux).HorizontalAlignment = xlHAlignCenter
      
       r_str_Cadena = ""
       r_int_FilExl = 19
       r_int_fildet = 0
       r_bol_Estado = False
       
       For r_int_FilGrd = 1 To 6
           Select Case r_int_FilGrd
                  Case 1: r_str_Cadena = pnl_Periodo1.Caption
                  Case 2: r_str_Cadena = pnl_Periodo2.Caption
                  Case 3: r_str_Cadena = pnl_Periodo3.Caption
                  Case 4: r_str_Cadena = pnl_Periodo4.Caption
                  Case 5: r_str_Cadena = pnl_Periodo5.Caption
                  Case 6: r_str_Cadena = pnl_Periodo6.Caption
           End Select
                      
           If (Len(Trim(r_str_Cadena)) > 3) Then
               If (r_bol_Estado = False) Then
                   .Cells(19, 2) = "Detalle del Periodo : " & r_str_Cadena
                   r_int_FilExl = r_int_FilExl + 1
                   .Range("B19:C19").Merge
                   .Range("B19:C19").Font.Bold = True
               Else
                   .Cells(r_int_FilExl, 2) = "Detalle del Periodo : " & r_str_Cadena
                   .Range("B" & r_int_FilExl & ":C" & r_int_FilExl).Merge
                   .Range("B" & r_int_FilExl & ":C" & r_int_FilExl).Font.Bold = True
                   r_int_FilExl = r_int_FilExl + 1
               End If
   
               For r_int_fildet = 0 To grd_Listad_rcc2.Rows - 1
                   If (Trim(grd_Listad_rcc2.TextMatrix(r_int_fildet, 1)) = r_str_Cadena) Then
                       .Cells(r_int_FilExl, 2) = "     " & Trim(grd_Listad_rcc2.TextMatrix(r_int_fildet, 2))
                       .Cells(r_int_FilExl, 4) = Trim(grd_Listad_rcc2.TextMatrix(r_int_fildet, 3)) '"Clasificación"
                       .Cells(r_int_FilExl, 5) = Trim(grd_Listad_rcc2.TextMatrix(r_int_fildet, 4)) '"Tipo Deuda
                       .Cells(r_int_FilExl, 7) = Trim(grd_Listad_rcc2.TextMatrix(r_int_fildet, 6))
                       .Cells(r_int_FilExl, 8) = Trim(grd_Listad_rcc2.TextMatrix(r_int_fildet, 7))
                       .Cells(r_int_FilExl, 7).NumberFormat = "###,###,##0.00"
                       .Cells(r_int_FilExl, 8).NumberFormat = "###,###,##0.00"
                       .Cells(r_int_FilExl, 9) = Trim(grd_Listad_rcc2.TextMatrix(r_int_fildet, 8))
                       .Cells(r_int_FilExl, 9).NumberFormat = "###,###,##0.00"
                       '.Cells(r_int_filexl, 10) = Trim(grd_Listad_rcc2.TextMatrix(r_int_fildet, 9))
                       
                       r_int_FilExl = r_int_FilExl + 1
                       r_bol_Estado = True
                   End If
               Next
               r_int_FilExl = r_int_FilExl + 1
           End If
           
           If (r_bol_Estado = False) Then
               r_int_FilExl = 19
           End If
       Next
       
       If (r_bol_Estado = False) Then
           .Cells(19, 2) = ""
       End If
   End With

   Screen.MousePointer = 0
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
   
End Sub


