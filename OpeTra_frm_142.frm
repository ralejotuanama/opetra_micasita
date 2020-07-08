VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_Ges_CreHip_07 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   9555
   ClientLeft      =   3255
   ClientTop       =   1830
   ClientWidth     =   11700
   Icon            =   "OpeTra_frm_142.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9555
   ScaleWidth      =   11700
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Height          =   585
      Left            =   3480
      Picture         =   "OpeTra_frm_142.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   124
      ToolTipText     =   "Exportar a Excel"
      Top             =   840
      Width           =   585
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   9555
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   11715
      _Version        =   65536
      _ExtentX        =   20664
      _ExtentY        =   16854
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
      Begin Threed.SSPanel SSPanel39 
         Height          =   645
         Left            =   0
         TabIndex        =   9
         Top             =   720
         Width           =   11655
         _Version        =   65536
         _ExtentX        =   20558
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
         Begin VB.CommandButton Command1 
            Height          =   585
            Left            =   2520
            Picture         =   "OpeTra_frm_142.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   123
            ToolTipText     =   "Exportar a Excel"
            Top             =   120
            Width           =   585
         End
         Begin VB.CommandButton cmd_Cronog 
            Height          =   585
            Left            =   1230
            Picture         =   "OpeTra_frm_142.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   120
            ToolTipText     =   "Actualizar Cronograma"
            Top             =   30
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   630
            Picture         =   "OpeTra_frm_142.frx":0EEA
            Style           =   1  'Graphical
            TabIndex        =   105
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   0
            Picture         =   "OpeTra_frm_142.frx":11F4
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Imprimir Cronograma"
            Top             =   0
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   11040
            Picture         =   "OpeTra_frm_142.frx":1636
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   10
         Top             =   30
         Width           =   11655
         _Version        =   65536
         _ExtentX        =   20558
         _ExtentY        =   1191
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
            Height          =   315
            Left            =   690
            TabIndex        =   11
            Top             =   30
            Width           =   5685
            _Version        =   65536
            _ExtentX        =   10028
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Gestión de Crédito Hipotecario"
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
         Begin Threed.SSPanel SSPanel8 
            Height          =   315
            Left            =   690
            TabIndex        =   12
            Top             =   330
            Width           =   5505
            _Version        =   65536
            _ExtentX        =   9710
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Cronogramas de Pago"
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
            Left            =   11040
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
            Picture         =   "OpeTra_frm_142.frx":1A78
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   30
         TabIndex        =   13
         Top             =   1380
         Width           =   11655
         _Version        =   65536
         _ExtentX        =   20558
         _ExtentY        =   1349
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
            TabIndex        =   14
            Top             =   60
            Width           =   2535
            _Version        =   65536
            _ExtentX        =   4471
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
            Left            =   1560
            TabIndex        =   15
            Top             =   390
            Width           =   9855
            _Version        =   65536
            _ExtentX        =   17383
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "1-07521154 / IKEHARA PUNK MIGUEL ANGEL"
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
         Begin Threed.SSPanel pnl_Periodo2 
            Height          =   315
            Left            =   8880
            TabIndex        =   121
            Top             =   60
            Width           =   2535
            _Version        =   65536
            _ExtentX        =   4471
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
         End
         Begin VB.Label pnl_Periodo1 
            Caption         =   "Período:"
            Height          =   315
            Left            =   8040
            TabIndex        =   122
            Top             =   90
            Width           =   885
         End
         Begin VB.Label Label5 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   90
            TabIndex        =   17
            Top             =   390
            Width           =   1395
         End
         Begin VB.Label Label12 
            Caption         =   "Nro. Operación:"
            Height          =   315
            Left            =   90
            TabIndex        =   16
            Top             =   90
            Width           =   1395
         End
      End
      Begin Threed.SSPanel SSPanel22 
         Height          =   7365
         Left            =   30
         TabIndex        =   18
         Top             =   2160
         Width           =   11655
         _Version        =   65536
         _ExtentX        =   20558
         _ExtentY        =   12991
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
            Height          =   7245
            Left            =   120
            TabIndex        =   0
            Top             =   60
            Width           =   11400
            _ExtentX        =   20108
            _ExtentY        =   12779
            _Version        =   393216
            Style           =   1
            Tabs            =   6
            TabsPerRow      =   6
            TabHeight       =   520
            TabCaption(0)   =   "Cliente - Tramo No Concesional"
            TabPicture(0)   =   "OpeTra_frm_142.frx":1D82
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "lbl_Totale(0)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "pnl_CliNCo_TotCuo"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "SSPanel63"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "SSPanel60"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "SSPanel58"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "SSPanel43"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "SSPanel42"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "SSPanel41"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).Control(8)=   "SSPanel38"
            Tab(0).Control(8).Enabled=   0   'False
            Tab(0).Control(9)=   "SSPanel5"
            Tab(0).Control(9).Enabled=   0   'False
            Tab(0).Control(10)=   "SSPanel4"
            Tab(0).Control(10).Enabled=   0   'False
            Tab(0).Control(11)=   "pnl_CliNCo_OtrCar"
            Tab(0).Control(11).Enabled=   0   'False
            Tab(0).Control(12)=   "grd_CliNCo_Listad"
            Tab(0).Control(12).Enabled=   0   'False
            Tab(0).Control(13)=   "pnl_CliNCo_Intere"
            Tab(0).Control(13).Enabled=   0   'False
            Tab(0).Control(14)=   "pnl_CliNCo_SegPre"
            Tab(0).Control(14).Enabled=   0   'False
            Tab(0).Control(15)=   "pnl_CliNCo_SegViv"
            Tab(0).Control(15).Enabled=   0   'False
            Tab(0).Control(16)=   "pnl_CliNCo_Capita"
            Tab(0).Control(16).Enabled=   0   'False
            Tab(0).ControlCount=   17
            TabCaption(1)   =   "Cliente - Tramo Concesional"
            TabPicture(1)   =   "OpeTra_frm_142.frx":1D9E
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "lbl_Totale(1)"
            Tab(1).Control(1)=   "pnl_CliCon_Capita"
            Tab(1).Control(2)=   "pnl_CliCon_Intere"
            Tab(1).Control(3)=   "SSPanel21"
            Tab(1).Control(4)=   "SSPanel13"
            Tab(1).Control(5)=   "SSPanel12"
            Tab(1).Control(6)=   "SSPanel11"
            Tab(1).Control(7)=   "SSPanel10"
            Tab(1).Control(8)=   "grd_CliCon_Listad"
            Tab(1).Control(9)=   "SSPanel9"
            Tab(1).Control(10)=   "pnl_CliCon_TotCuo"
            Tab(1).ControlCount=   11
            TabCaption(2)   =   "FMV - Tramo No Concesional"
            TabPicture(2)   =   "OpeTra_frm_142.frx":1DBA
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "lbl_Totale(2)"
            Tab(2).Control(1)=   "pnl_MViNCo_Comisi"
            Tab(2).Control(2)=   "SSPanel3"
            Tab(2).Control(3)=   "SSPanel62"
            Tab(2).Control(4)=   "SSPanel61"
            Tab(2).Control(5)=   "SSPanel59"
            Tab(2).Control(6)=   "SSPanel36"
            Tab(2).Control(7)=   "SSPanel35"
            Tab(2).Control(8)=   "SSPanel34"
            Tab(2).Control(9)=   "SSPanel33"
            Tab(2).Control(10)=   "SSPanel2"
            Tab(2).Control(11)=   "SSPanel30"
            Tab(2).Control(12)=   "pnl_MViNCo_TotCuo"
            Tab(2).Control(13)=   "pnl_MViNCo_OtrCar"
            Tab(2).Control(14)=   "grd_MViNCo_Listad"
            Tab(2).Control(15)=   "pnl_MViNCo_Intere"
            Tab(2).Control(16)=   "pnl_MViNCo_SegPre"
            Tab(2).Control(17)=   "pnl_MViNCo_SegViv"
            Tab(2).Control(18)=   "pnl_MViNCo_Capita"
            Tab(2).ControlCount=   19
            TabCaption(3)   =   "FMV - Tramo Concesional"
            TabPicture(3)   =   "OpeTra_frm_142.frx":1DD6
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "lbl_Totale(3)"
            Tab(3).Control(1)=   "pnl_MViCon_Comisi"
            Tab(3).Control(2)=   "pnl_MViCon_Capita"
            Tab(3).Control(3)=   "pnl_MViCon_Intere"
            Tab(3).Control(4)=   "SSPanel20"
            Tab(3).Control(5)=   "SSPanel19"
            Tab(3).Control(6)=   "SSPanel18"
            Tab(3).Control(7)=   "SSPanel17"
            Tab(3).Control(8)=   "SSPanel16"
            Tab(3).Control(9)=   "SSPanel15"
            Tab(3).Control(10)=   "grd_MviCon_Listad"
            Tab(3).Control(11)=   "SSPanel14"
            Tab(3).Control(12)=   "pnl_MViCon_TotCuo"
            Tab(3).ControlCount=   13
            TabCaption(4)   =   "Cofide"
            TabPicture(4)   =   "OpeTra_frm_142.frx":1DF2
            Tab(4).ControlEnabled=   0   'False
            Tab(4).Control(0)=   "lbl_Totale(4)"
            Tab(4).Control(1)=   "pnl_CofNCo_Comisi"
            Tab(4).Control(2)=   "pnl_CofNCo_Capita"
            Tab(4).Control(3)=   "pnl_CofNCo_Intere"
            Tab(4).Control(4)=   "SSPanel55"
            Tab(4).Control(5)=   "SSPanel54"
            Tab(4).Control(6)=   "SSPanel49"
            Tab(4).Control(7)=   "SSPanel47"
            Tab(4).Control(8)=   "SSPanel46"
            Tab(4).Control(9)=   "SSPanel45"
            Tab(4).Control(10)=   "grd_CofNCo_Listad"
            Tab(4).Control(11)=   "SSPanel44"
            Tab(4).Control(12)=   "pnl_CofNCo_TotCuo"
            Tab(4).ControlCount=   13
            TabCaption(5)   =   "Cronograma Especial"
            TabPicture(5)   =   "OpeTra_frm_142.frx":1E0E
            Tab(5).ControlEnabled=   0   'False
            Tab(5).Control(0)=   "lbl_Totale(5)"
            Tab(5).Control(1)=   "pnl_Especial_Seguros"
            Tab(5).Control(2)=   "pnl_Especial_Capital"
            Tab(5).Control(3)=   "SSPanel66"
            Tab(5).Control(4)=   "SSPanel69"
            Tab(5).Control(5)=   "pnl_Especial_TotalCuota"
            Tab(5).Control(6)=   "pnl_Especial_Interes"
            Tab(5).Control(7)=   "SSPanel68"
            Tab(5).Control(8)=   "SSPanel67"
            Tab(5).Control(9)=   "SSPanel65"
            Tab(5).Control(10)=   "SSPanel64"
            Tab(5).Control(11)=   "grd_Especial_Cli"
            Tab(5).Control(12)=   "SSPanel56"
            Tab(5).ControlCount=   13
            Begin Threed.SSPanel pnl_MViCon_TotCuo 
               Height          =   285
               Left            =   -67470
               TabIndex        =   19
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
            Begin Threed.SSPanel pnl_CliNCo_Capita 
               Height          =   285
               Left            =   2280
               TabIndex        =   20
               Top             =   6870
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
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
            Begin Threed.SSPanel pnl_CliNCo_SegViv 
               Height          =   285
               Left            =   5790
               TabIndex        =   21
               Top             =   6870
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
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
            Begin Threed.SSPanel pnl_CliNCo_SegPre 
               Height          =   285
               Left            =   4620
               TabIndex        =   22
               Top             =   6870
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
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
            Begin Threed.SSPanel pnl_CliNCo_Intere 
               Height          =   285
               Left            =   3450
               TabIndex        =   23
               Top             =   6870
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
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
            Begin MSFlexGridLib.MSFlexGrid grd_CliNCo_Listad 
               Height          =   6135
               Left            =   30
               TabIndex        =   1
               Top             =   690
               Width           =   11265
               _ExtentX        =   19870
               _ExtentY        =   10821
               _Version        =   393216
               Rows            =   25
               Cols            =   9
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel SSPanel23 
               Height          =   285
               Left            =   -67530
               TabIndex        =   24
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
               TabIndex        =   25
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
               TabIndex        =   26
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
               TabIndex        =   27
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
               TabIndex        =   28
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
               TabIndex        =   29
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
               TabIndex        =   30
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
               TabIndex        =   31
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
               TabIndex        =   32
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
               TabIndex        =   33
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
               TabIndex        =   34
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
               TabIndex        =   35
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
               TabIndex        =   36
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
               TabIndex        =   37
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
               TabIndex        =   38
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
               TabIndex        =   39
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
            Begin Threed.SSPanel pnl_CliNCo_OtrCar 
               Height          =   285
               Left            =   6990
               TabIndex        =   40
               Top             =   6870
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
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
            Begin Threed.SSPanel SSPanel14 
               Height          =   285
               Left            =   -70950
               TabIndex        =   41
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
            Begin MSFlexGridLib.MSFlexGrid grd_MviCon_Listad 
               Height          =   6135
               Left            =   -74970
               TabIndex        =   4
               Top             =   690
               Width           =   11265
               _ExtentX        =   19870
               _ExtentY        =   10821
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
            Begin Threed.SSPanel SSPanel15 
               Height          =   285
               Left            =   -74940
               TabIndex        =   42
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
            Begin Threed.SSPanel SSPanel16 
               Height          =   285
               Left            =   -74190
               TabIndex        =   43
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
            Begin Threed.SSPanel SSPanel17 
               Height          =   285
               Left            =   -72690
               TabIndex        =   44
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
            Begin Threed.SSPanel SSPanel18 
               Height          =   285
               Left            =   -67470
               TabIndex        =   45
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
            Begin Threed.SSPanel SSPanel19 
               Height          =   285
               Left            =   -65730
               TabIndex        =   46
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
            Begin Threed.SSPanel SSPanel20 
               Height          =   285
               Left            =   -69210
               TabIndex        =   47
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
            Begin Threed.SSPanel pnl_MViCon_Intere 
               Height          =   285
               Left            =   -70950
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
            Begin Threed.SSPanel pnl_MViCon_Capita 
               Height          =   285
               Left            =   -72690
               TabIndex        =   49
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
            Begin Threed.SSPanel pnl_MViCon_Comisi 
               Height          =   285
               Left            =   -69210
               TabIndex        =   50
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
            Begin Threed.SSPanel pnl_CliCon_TotCuo 
               Height          =   285
               Left            =   -68370
               TabIndex        =   51
               Top             =   6870
               Width           =   2175
               _Version        =   65536
               _ExtentX        =   3828
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
            Begin Threed.SSPanel SSPanel9 
               Height          =   285
               Left            =   -70530
               TabIndex        =   52
               Top             =   390
               Width           =   2175
               _Version        =   65536
               _ExtentX        =   3836
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Interes"
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
               Height          =   6135
               Left            =   -74970
               TabIndex        =   2
               Top             =   690
               Width           =   11265
               _ExtentX        =   19870
               _ExtentY        =   10821
               _Version        =   393216
               Rows            =   25
               Cols            =   6
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel SSPanel10 
               Height          =   285
               Left            =   -74940
               TabIndex        =   53
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
            Begin Threed.SSPanel SSPanel11 
               Height          =   285
               Left            =   -74190
               TabIndex        =   54
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
            Begin Threed.SSPanel SSPanel12 
               Height          =   285
               Left            =   -72690
               TabIndex        =   55
               Top             =   390
               Width           =   2175
               _Version        =   65536
               _ExtentX        =   3828
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
            Begin Threed.SSPanel SSPanel13 
               Height          =   285
               Left            =   -68370
               TabIndex        =   56
               Top             =   390
               Width           =   2175
               _Version        =   65536
               _ExtentX        =   3836
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
            Begin Threed.SSPanel SSPanel21 
               Height          =   285
               Left            =   -66210
               TabIndex        =   57
               Top             =   390
               Width           =   2205
               _Version        =   65536
               _ExtentX        =   3889
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
            Begin Threed.SSPanel pnl_CliCon_Intere 
               Height          =   285
               Left            =   -70530
               TabIndex        =   58
               Top             =   6870
               Width           =   2175
               _Version        =   65536
               _ExtentX        =   3828
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
            Begin Threed.SSPanel pnl_CliCon_Capita 
               Height          =   285
               Left            =   -72690
               TabIndex        =   59
               Top             =   6870
               Width           =   2175
               _Version        =   65536
               _ExtentX        =   3828
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
            Begin Threed.SSPanel pnl_MViNCo_Capita 
               Height          =   285
               Left            =   -72840
               TabIndex        =   60
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
            Begin Threed.SSPanel pnl_MViNCo_SegViv 
               Height          =   285
               Left            =   -69630
               TabIndex        =   61
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
            Begin Threed.SSPanel pnl_MViNCo_SegPre 
               Height          =   285
               Left            =   -70710
               TabIndex        =   62
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
            Begin Threed.SSPanel pnl_MViNCo_Intere 
               Height          =   285
               Left            =   -71790
               TabIndex        =   63
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
            Begin MSFlexGridLib.MSFlexGrid grd_MViNCo_Listad 
               Height          =   6135
               Left            =   -74970
               TabIndex        =   3
               Top             =   690
               Width           =   11265
               _ExtentX        =   19870
               _ExtentY        =   10821
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
            Begin Threed.SSPanel pnl_MViNCo_OtrCar 
               Height          =   285
               Left            =   -68550
               TabIndex        =   64
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
            Begin Threed.SSPanel pnl_MViNCo_TotCuo 
               Height          =   285
               Left            =   -66390
               TabIndex        =   65
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
            Begin Threed.SSPanel SSPanel4 
               Height          =   285
               Left            =   3450
               TabIndex        =   66
               Top             =   390
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
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
               Left            =   60
               TabIndex        =   67
               Top             =   390
               Width           =   795
               _Version        =   65536
               _ExtentX        =   1402
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
            Begin Threed.SSPanel SSPanel38 
               Height          =   285
               Left            =   840
               TabIndex        =   68
               Top             =   390
               Width           =   1455
               _Version        =   65536
               _ExtentX        =   2566
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
            Begin Threed.SSPanel SSPanel41 
               Height          =   285
               Left            =   2280
               TabIndex        =   69
               Top             =   390
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
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
            Begin Threed.SSPanel SSPanel42 
               Height          =   285
               Left            =   8130
               TabIndex        =   70
               Top             =   390
               Width           =   1320
               _Version        =   65536
               _ExtentX        =   2328
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
            Begin Threed.SSPanel SSPanel43 
               Height          =   285
               Left            =   9450
               TabIndex        =   71
               Top             =   390
               Width           =   1560
               _Version        =   65536
               _ExtentX        =   2752
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
            Begin Threed.SSPanel SSPanel58 
               Height          =   285
               Left            =   4620
               TabIndex        =   72
               Top             =   390
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
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
            Begin Threed.SSPanel SSPanel60 
               Height          =   285
               Left            =   5790
               TabIndex        =   73
               Top             =   390
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
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
            Begin Threed.SSPanel SSPanel63 
               Height          =   285
               Left            =   6960
               TabIndex        =   74
               Top             =   390
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
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
            Begin Threed.SSPanel SSPanel30 
               Height          =   285
               Left            =   -71760
               TabIndex        =   75
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
            Begin Threed.SSPanel SSPanel2 
               Height          =   285
               Left            =   -74940
               TabIndex        =   76
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
               TabIndex        =   77
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
               TabIndex        =   78
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
               TabIndex        =   79
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
               TabIndex        =   80
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
               TabIndex        =   81
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
               TabIndex        =   82
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
               TabIndex        =   83
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
            Begin Threed.SSPanel SSPanel3 
               Height          =   285
               Left            =   -67440
               TabIndex        =   84
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
               TabIndex        =   85
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
               TabIndex        =   86
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
               TabIndex        =   87
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
            Begin MSFlexGridLib.MSFlexGrid grd_CofNCo_Listad 
               Height          =   6135
               Left            =   -74970
               TabIndex        =   5
               Top             =   690
               Width           =   11265
               _ExtentX        =   19870
               _ExtentY        =   10821
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
            Begin Threed.SSPanel SSPanel45 
               Height          =   285
               Left            =   -74940
               TabIndex        =   88
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
               TabIndex        =   89
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
               TabIndex        =   90
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
               TabIndex        =   91
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
               TabIndex        =   92
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
               TabIndex        =   93
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
               TabIndex        =   94
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
               TabIndex        =   95
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
               TabIndex        =   96
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
            Begin Threed.SSPanel pnl_CliNCo_TotCuo 
               Height          =   285
               Left            =   8190
               TabIndex        =   106
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
            Begin Threed.SSPanel SSPanel56 
               Height          =   285
               Left            =   -70710
               TabIndex        =   107
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
            Begin MSFlexGridLib.MSFlexGrid grd_Especial_Cli 
               Height          =   6135
               Left            =   -74970
               TabIndex        =   108
               Top             =   690
               Width           =   11265
               _ExtentX        =   19870
               _ExtentY        =   10821
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
            Begin Threed.SSPanel SSPanel64 
               Height          =   285
               Left            =   -74940
               TabIndex        =   109
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
               TabIndex        =   110
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
               TabIndex        =   111
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
               TabIndex        =   112
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
               TabIndex        =   114
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
               TabIndex        =   115
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
               TabIndex        =   116
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
               TabIndex        =   117
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
               TabIndex        =   118
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
               TabIndex        =   119
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
            Begin VB.Label lbl_Totale 
               Alignment       =   1  'Right Justify
               Caption         =   "Totales ===> US$ "
               Height          =   315
               Index           =   5
               Left            =   -74610
               TabIndex        =   113
               Top             =   6870
               Width           =   1845
            End
            Begin VB.Label lbl_Totale 
               Alignment       =   1  'Right Justify
               Caption         =   "Totales ===> US$ "
               Height          =   315
               Index           =   4
               Left            =   -74610
               TabIndex        =   104
               Top             =   6870
               Width           =   1845
            End
            Begin VB.Label Label4 
               Caption         =   "Totales ==>"
               Height          =   285
               Left            =   -72930
               TabIndex        =   103
               Top             =   1470
               Width           =   945
            End
            Begin VB.Label Label14 
               Caption         =   "Totales ==>"
               Height          =   285
               Left            =   -72930
               TabIndex        =   102
               Top             =   1470
               Width           =   945
            End
            Begin VB.Label Label15 
               Caption         =   "Totales ==>"
               Height          =   285
               Left            =   -73230
               TabIndex        =   101
               Top             =   1470
               Width           =   945
            End
            Begin VB.Label lbl_Totale 
               Alignment       =   1  'Right Justify
               Caption         =   "Totales ===> US$ "
               Height          =   255
               Index           =   0
               Left            =   390
               TabIndex        =   100
               Top             =   6870
               Width           =   1845
            End
            Begin VB.Label lbl_Totale 
               Alignment       =   1  'Right Justify
               Caption         =   "Totales ===> US$ "
               Height          =   315
               Index           =   1
               Left            =   -74610
               TabIndex        =   99
               Top             =   6870
               Width           =   1845
            End
            Begin VB.Label lbl_Totale 
               Alignment       =   1  'Right Justify
               Caption         =   "Totales ===> US$ "
               Height          =   315
               Index           =   2
               Left            =   -74790
               TabIndex        =   98
               Top             =   6870
               Width           =   1845
            End
            Begin VB.Label lbl_Totale 
               Alignment       =   1  'Right Justify
               Caption         =   "Totales ===> US$ "
               Height          =   315
               Index           =   3
               Left            =   -74610
               TabIndex        =   97
               Top             =   6870
               Width           =   1845
            End
         End
      End
   End
End
Attribute VB_Name = "frm_Ges_CreHip_07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public l_str_PerMes        As String
Public l_str_PerAno        As String
Private l_int_FlgCro       As Integer

Private Sub cmd_Imprim_Click()
   If tab_Cronog.Tab = 2 And (InStr(moddat_g_str_AgrCRC, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0) Then   '"001" "003"
      MsgBox "Estos cronogramas no se imprimen, ya que son de uso interno con Mivivienda.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de imprimir el Cronograma?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
   crp_Imprim.DataFiles(0) = "CRE_HIPCUO"
   crp_Imprim.DataFiles(1) = "CRE_HIPMAE"
   crp_Imprim.DataFiles(2) = "CLI_DATGEN"
   crp_Imprim.DataFiles(3) = "CRE_PRODUC"
   crp_Imprim.DataFiles(4) = ""
   
   If tab_Cronog.Tab = 0 Then
      'Cliente No Concesional
      '********* 22102020   inicio rat
      Dim a, b, m, n As String
      
'      b = ""
'
'b = b & "41869714','41916917','41922234','41922641','41937011','41953349','41971140','41987027','42007137','42020411','42020923','42031556','42032619','42039361','42050691','42052792','42058707','42074623','42081430','42082753','42100020','42104060','42114346','42126512','42129959','42129984','42137568','42138265','42139496','42147430','42152597','42184982',"
' b = b & "'42188273','42204842','42226012','42241840','42250537','42270630','42273570','42275937','42281239','42283537','42285984','42287053','42297607','42309064','42316810','42334029','42345033','42368075','42371824','42376073','42383293','42427647','42449521','42465611','42484439','42503038','42507550','42535568','42565350','42592776','42599422','42599531',"
'  b = b & "'42621225','42625536','42655331','42672824','42674130','42683641','42684211','42702479','42715691','42725511','42743215','42748199','42773317','42821688','42823291','42829990','42855221','42856256','42858791','42861790','42866117','42874016','42876121','42883470','42899060','42899851','42903868','42912782','42932592','42949920','42960161','42961333',"
' b = b & "'42979451','42995218','43021411','43026010','43035307','43053243','43055563','43058109','43066347','43071750','43079155','43085083','43087580','43134726','43140916','43141441','43156961','43157372','43167504','43176785','43193212','43197200','43210266','43214595','43221389','43232468','43241394','43242062','43247494','43258497','43263883','43264356',"
' b = b & "'43264654','43273789','43301432','43304408','43305928','43320458','43330419','43347140','43354202','43354208','43357738','43362009','43362493','43372269','43417957','43431308','43464893','43466267','43495397','43498317','43522692','43538305','43562246','43575203','43591925','43593559','43617629','43628041','43628478','43637712','43638770','43642298',"
' b = b & "'43651693','43654803','43657896','43663740','43675850','43682494','43689525','43713116','43714077','43756555','43778770','43802053','43823332','43842046','43852327','43888685','43900536','43935342','43940717','43942096','43969015','43969021','43976211','43976793','43983419','43997886','44003244','44004673','44019126','44029471','44030252','44035414',"
' b = b & "'44057613','44094310','44097700','44112461','44137197','44144548','44148011','44157342','44173167','44175445','44175747','44177969','44200206','44201570','44204812','44204919','44223240','44224192','44237029','44239414','44254213','44268226','44273008','44292101','44292226','44310293','44319200','44336258','44346473','44354937','44355148','44360465',"
' b = b & "'44385608','44386510','44386782','44393669','44407975','44418967','44427207','44427588','44433888','44453494','44466096','44467591','44471511','44478700','44496511','44502474','44512380','44524792','44525490','44534548','44554720','44561617','44582158','44583235','44587732','44605949','44608241','44617428','44621941','44622602','44633213','44635641',"
' b = b & "'44645104','44662600','44694382','44711024','44728065','44745660','44784897','44791782','44805359','44810892','44844300','44846486','44857854','44894152','44916441','44922279','44932636','44937146','44950620','44961447','45003898','45011735','45017428','45027796','45030889','45041722','45051032','45059360','45064041','45078537','45086920','45096079',"
' b = b & "'45110164','45125420','45130663','45136368','45137679','45137840','45190937','45198501','45198521','45219245','45220422','45230983','45235593','45256495','45257629','45260005','45281933','45312992','45326514','45338660','45345585','45346920','45359551','45364741','45365361','45376557','45387271','45387618','45433118','45437719','45442414','45443325',"
' b = b & "'45443399','45447812','45448315','45462974','45473282','45478546','45488923','45500044','45508297','45512741','45516968','45517711','45524772','45526553','45531340','45532995','45538272','45538648','45549854','45551889','45558577','45568544','45578971','45583456','45591021','45600320','45617405','45617954','45626140','45628808','45629814','45642145',"
' b = b & "'45685265','45694715','45712810','45730553','45738244','45741578','45763758','45782327','45782503','45786720','45786980','45787925','45796095','45797613','45799147','45799258','45800877','45801517','45813554','45813934','45821313','45836999','45842299','45845682','45876055','45878497','45879302','45891952','45894913','45895769','45905394','45915191',"
' b = b & "'45935486','45948968','45957321','45962257','45963086','45976367','45990439','46004516','46012796','46016032','46042221','46047977','46066585','46082044','46097825','46100145','46104845','46118216','46121377','46121761','46122794','46146097','46166238','46176561','46181705','46183828','46184003','46191684','46194314','46208446','46212061','46223091',"
' b = b & "'46230025','46238757','46244174','46245186','46266849','46280909','46282380','46314611','46317933','46325428','46328296','46332277','46339719','46361780','46362912','46364474','46366640','46370221','46374761','46383530','46391492','46420623','46421729','46424834','46431380','46433155','46436652','46441255','46451260','46460518','46466506','46469313',"
' b = b & "'46471898','46487986','46489323','46497130','46502132','46523855','46580781','46588452','46590386','46592130','46592655','46595343','46596096','46606058','46608670','46609598','46617854','46621496','46623760','46640979','46655878','46663365','46665615','46666260','46666608','46666636','46666687','46677518','46686374','46687089','46687843','46695297',"
' b = b & "'46700715','46704901','46709934','46710125','46722861','46727145','46740372','46741752','46744513','46766541','46772094','46773973','46785680','46790003','46792239','46806617','46808350','46815250','46818400','46819109','46822746','46833183','46853522','46858129','46864431','46879474','46883203','46925640','46931153','46944224','46954976','46957396',"
' b = b & "'46958596','46959921','46964881','46976180','46995541','47004066','47016614','47025284','47026255','47031228','47034080','47062436','47065355','47092153','47097878','47105035','47109128','47117729','47122978','47134399','47140226','47142905','47149127','47153896','47159124','47164869','47165216','47168705','47170792','47170916','47179386','47185748',"
' b = b & "'47190216','47190813','47199052','47243654','47244153','47245939','47263796','47286734','47287008','47304043','47341935','47347695','47367892','47367946','47388978','47397502','47401620','47408378','47420142','47429085','47445163','47449310','47453648','47454175','47461248','47483337','47485202','47492011','47504659','47506285','47529764','47541104',"
' b = b & "'47541163','47578384','47582867','47583034','47587446','47596603','47604780','47605293','47617657','47673028','47691837','47747840','47749955','47757585','47761253','47817583','47824701','47829759','47834368','47840686','47851263','47856510','47859230','47898253','47899566','47902540','47939399','47958735','47999937','48006321','48010700','48021346',"
' b = b & "'48040014','48079535','48118871','48149672','48182515','48222515','48269873','48281787','48286293','48296280','48329890','48341882','48367985','48398359','48401103','48419238','48446867','48485520','48486093','48491520','48506938','48516711','48585861','48667405','48972620','48973029','48975696','60132806','60914123','60927967','70000672','70001169',"
' b = b & "'70005251','70051283','70055386','70057990','70067082','70069002','70085559','70095600','70123522','70131028','70145427','70157124','70163051','70172995','70179974','70183153','70188740','70192197','70193030','70195257','70197781','70215985','70223408','70229074','70268136','70270416','70270429','70299664','70309678','70314187','70314752','70336266',"
' b = b & "'70338043','70356880','70427782','70438603','70466285','70472985','70477680','70481666','70492225','70504451','70508278','70545437','70548294','70568410','70582212','70651212','70690329','70745326','70750865','70776941','70788718','70840784','70904530','70992591','70992595','70999220','71003130','71007763','71008282','71067330','71069843','71081291',"
' b = b & "'71241786','71261782','71285577','71328514','71417495','71418639','71442219','71447921','71454012','71490985','71495505','71507531','71542026','71542303','71560076','71618895','71648732','71693448','71696720','71716332','71718406','71726174','71770048','71821086','71921376','71936784','71992209','72003571','72145932','72163758','72198108','72208634',"
' b = b & "'72219393','72246206','72261152','72276885','72386014','72401589','72406949','72450089','72519001','72519551','72528422','72529305','72533158','72617878','72644589','72647391','72651035','72655528','72662639','72675989','72682586','72683023','72683816','72698991','72716207','72751460','72781092','72864122','72866423','72869158','72883896','72907883',"
' b = b & "'72910946','72928443','72982258','73030417','73046294','73071856','73116138','73142432','73182198','73211078','73237160','73257848','73273452','73276114','73305564','73329216','73473857','73520982','73648996','73695946','73702185','73741015','73755868','73777557','73808678','73869526','73884477','73940701','73947444','73992812','74050320','74066531',"
' b = b & "'74092963','74171539','74238065','74268333','74309769','74478962','74487200','74571678','74608731','74754116','74843135','74939179','74977073','74987618','74993842','75105654','75210940','75270586','75332225','75533034','75552736','75594652','75731514','75922092','76209461','76211352','76232693','76257131','76265040','76463782','76536292','76770587',"
' b = b & "'76960244','76976636','76991929','77034562','77065368','77092430','77236386','77506138','77682420','80247667','80256264','80257725','80346619','80347372','80493602','80638500','80671219','000409653''000767975','000923586','001539111','001556905','001716233','002069448','01780568"
          
'  m = ""
'
' m = m & "41406494','41633337', '00767975', '41610849','41438214','41629660','41398885','41334082', '41730404', '41491947','41467461','41339636','41489510','41401872','41399210',   '41693178',    '41534611',    '41368131',    '41476944',"
'm = m & "'41468552','41476414','41641414','41341216','41533268','41393336','41715512', '41583547', '41555036',  '41273120', '41466721', '41273644',"
'm = m & "'41309034', '41279583', '41705121','00409653','41688472', '41258211', '41508382', '41674452', '41388456', '41367362',"
'm = m & "'41301129', '41273108', '41683886', '41295417', '41459201','41528704','41403782', '41269541','41438888', '41703555', '41608967', '41729212"
     
     
'      a = "00113951','00118465"
     ' a = ""
'a = a & "00113951','00118465','00241532','00253086','00373794','00505509','00807482','00837874','00873584','00931107','00949386','00950494','00954508','00964816','00965789','01003924','01061914','01065255','01115416','01117242','01121487','01127522','01127565','01134817','01146497','01157203','01159272','01160668','01187069','01209607','01483908','02606960"
' a = a & "'02659587','02669562','02766514','02797348','02816492','02819491','02830609','02858057','02862467','02880443','03102972','03208780','03631527','03643639','03686552','03692158','03886816','05376058','05380629','05407080','05612494','05620568','05644946','06005052','06018823','06048016','06146489','06253887','06268463','06309139','06460743','06460748',"
' a = a & "'06650061','06675905','06708188','06709982','06744273','06775018','06783975','06813045','06885583','06925963','06934086','06936939','06955031','06955790','07178016','07200211','07244266','07292282','07339816','07376401','07465847','07478368','07499595','07525799','07529560','07554115','07574288','07642372','07675091','07706280','07758684','07814449',"
' a = a & "'07817450','07875956','07966125','08010489','08045456','08054305','08089043','08099102','08134509','08268077','08274952','08302105','08427026','08442208','08507623','08544626','08594091','08619301','08664834','08684908','08686770','08695611','08749402','08869544','08869657','09063584','09108222','09289662','09332489','09402406','09426107','09444451',"
' a = a & "'09454186','09458795','09478329','09544592','09576967','09594112','09604336','09605567','09619874','09624819','09668315','09673949','09675647','09731364','09733432','09743374','09797091','09823738','09885722','09886623','09893185','09898963','09926589','09930052','09935171','09936119','09963466','09983474','09985118','09986396','09988739','10072588',"
' a = a & "'10090459','10092674','10092770','10106732','10129084','10140202','10151980','10160359','10167870','10182209','10187341','10193359','10195503','10201971','10205976','10215535','10218179','10220469','10330250','10357009','10374883','10380719','10385560','10396705','10397271','10397292','10404836','10416772','10420076','10453624','10457434','10480073',"
' a = a & "'10515725','10567653','10657011','10664194','10686636','10713677','10736348','10743263','10744009','10747183','10763258','10764940','10771634','10797953','10810522','10812187','10881206','10886468','15283739','15414250','15585232','15627471','15714203','15761374','15764576','15842482','15847543','16022941','16125774','16483601','16610548','16657127',"
' a = a & "'16670875','16708853','16727300','16755464','16763165','16768149','16792526','16796128','17446034','17452468','17530453','17611151','17814274','17840258','17903415','17930793','18029505','18095464','18098262','18106846','18131558','18134872','18161756','18180643','18198768','18224401','18861681','18863706','18901146','19021115','19248300','19251714',"
' a = a & "'19429707','19561727','19814745','19870317','19916768','19923133','20027406','20045072','20097965','20551675','20587087','20902336','21144494','21271863','21405755','21415484','21424136','21426941','21446843','21449627','21453688','21458074','21462310','21463051','21464122','21487821','21489739','21501825','21520508','21520834','21522301','21522709',"
' a = a & "'21523385','21523402','21525533','21528413','21529219','21534042','21535025','21544870','21551283','21552401','21552899','21556125','21560074','21562968','21568148','21568528','21569288','21569770','21574645','21878106','22085840','22092618','22095112','22269869','22283139','22486551','23974670','24889563','25414019','25588300','25681697','25709318',"
'a = a & "'25718786','25731559','25745476','25762422','25780757','25818969','25819815','26693473','26949910','27060465','27143435','27383318','27689985','28312410','28313932','28803242','29408811','29470531','29534693','29572596','29623205','30497582','30500666','31628695','31650919','31883278','32489588','32608461','32610669','32917352','32945435','33340731',"
'a = a & "'33342629','33562770','33649079','33734182','40004489','40005619','40008000','40040638','40043568','40051412','40085791','40100099','40104867','40111042','40115973','40118043','40173108','40175089','40178488','40210780','40231433','40240426','40257541','40258028','40278864','40279013','40281748','40283117','40310141','40311201','40314774','40318488',"
' a = a & "'40337375','40339298','40348930','40371079','40383229','40392593','40416359','40463065','40470133','40483430','40488691','40496068','40498785','40511272','40520556','40541153','40557565','40577659','40584021','40604709','40609490','40616979','40624614','40628222','40629194','40630611','40636274','40643516','40648438','40652472','40671275','40686968',"
' a = a & "'40693222','40694743','40710820','40718245','40720849','40748904','40749909','40750514','40752120','40759170','40785713','40802522','40804782','40807782','40811180','40822875','40857370','40876758','40891834','40895155','40900853','40905992','40919021','40923064','40974640','40983317','40983949','40989524','41000530','41000842','41002278','41006375',"
' a = a & "'41016219','41034131','41042709','41057650','41075647','41077129','41087990','41114556','41133699','41173662','41184010','41188604','41192496','41198804','41200535','41204096','41212904','41216391','41217156','41226204','41253812','41737850','41739445','41745802','41745902','41748891','41769971','41787981','41815528','41821529','41826321','41866041"
      
      
      
n = ""
n = n & "46445475', '43199183', '43388140', '72739849', '46668595', '21523419', '09432557', '43571897', '71460452', '70306790',"
n = n & "'21568235', '71204058', '48383881', '21557865', '22306243', '71207284', '21525542', '22183474', '42874644', '21577756',"
n = n & "'45911173', '47005853', '21535843', '44879401', '41559046', '42302315', '21551664', '47525170', '28804760', '43638408',"
n = n & "'42914949', '42449521', '22081055', '43315296', '45079891', '42709595', '70432458', '40891267', '40771195', '42876419', '21465840',"
n = n & "'45999212', '46146884', '22181638', '44904828', '09545186', '21569770', '71886339', '40730593', '21562968', '70336266',"
n = n & "'21522709', '70750865', '72751460', '45364741', '47749955', '71495505', '74582271', '46440456', '43752288', '44502474"
      
      
      
       g_str_Parame = ""
       g_str_Parame = g_str_Parame & "select hipmae_numope as operacion, trim(hipmae_ndocli) as documento from cre_hipmae"
'      g_str_Parame = g_str_Parame & "       HIPCUO_DESORG, HIPCUO_VIVORG, HIPCUO_OTRORG, HIPCUO_SALCAP, HIPCUO_COMCOF  "
'      g_str_Parame = g_str_Parame & "  FROM CRE_HIPCUO "
       g_str_Parame = g_str_Parame & "  where hipmae_ndocli in ('" & n & "') "
       MsgBox (g_str_Parame)
'      g_str_Parame = g_str_Parame & "   AND HIPCUO_TIPCRO = 6 "
'      g_str_Parame = g_str_Parame & " ORDER BY HIPCUO_NUMCUO "
      '********* 22102020   fin rat
       If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
       End If
       Dim r_int_ConVer As Integer
       g_rst_Princi.MoveFirst
       r_int_ConVer = 1
       Do While Not g_rst_Princi.EOF
    'r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 1
    'r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = gf_Formato_NumOpe(g_rst_Princi!HIPCUO_NUMOPE)
    'r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = CStr(g_rst_Princi!HIPCUO_NUMCUO)
    'r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = "" & gf_FormatoFecha(g_rst_Princi!HIPCUO_FECVCT)
    'r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = IIf(CStr(g_rst_Princi!HIPCUO_SITUAC) = 2, "POR VENCER", "PAGADA")
    'r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = gf_FormatoNumero(g_rst_Princi!HIPCUO_CAPITA, 12, 2)
    'r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = gf_FormatoNumero(g_rst_Princi!HIPCUO_INTERE, 12, 2)
    'r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = gf_FormatoNumero(g_rst_Princi!HIPCUO_DESORG, 12, 2)
    'r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = gf_FormatoNumero(g_rst_Princi!HIPCUO_VIVORG, 12, 2)
    'r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = gf_FormatoNumero(g_rst_Princi!HIPCUO_OTRORG, 14, 4)
    'r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = gf_FormatoNumero(g_rst_Princi!HIPCUO_CAPITA + g_rst_Princi!HIPCUO_INTERE + g_rst_Princi!HIPCUO_DESORG + g_rst_Princi!HIPCUO_VIVORG + g_rst_Princi!HIPCUO_OTRORG, 14, 4)
    'r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = gf_FormatoNumero(g_rst_Princi!HIPCUO_SALCAP, 14, 4)
'        MsgBox (g_rst_Princi!OPERACION)

'    crp_Imprim.SelectionFormula = "{CRE_HIPCUO.HIPCUO_NUMOPE} = '" & moddat_g_str_NumOpe & "' AND {CRE_HIPCUO.HIPCUO_TIPCRO} = 1 "
'    If moddat_g_str_CodPrd = "002" Then
'    crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_CROPAG_11.RPT"
'    Else
'    crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_CROPAG_12.RPT"
'    End If

       fs_GenExcNuevo (g_rst_Princi!OPERACION)

       r_int_ConVer = r_int_ConVer + 1
       g_rst_Princi.MoveNext
       DoEvents
     Loop

    g_rst_Princi.Close
    Set g_rst_Princi = Nothing
      
       
        '********* 22102020   fin rat
        
   ElseIf tab_Cronog.Tab = 1 Then
      'Cliente Concesional
      crp_Imprim.SelectionFormula = "{CRE_HIPCUO.HIPCUO_NUMOPE} = '" & moddat_g_str_NumOpe & "' AND {CRE_HIPCUO.HIPCUO_TIPCRO} = 2 "
      crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_CROPAG_13.RPT"
      
   ElseIf tab_Cronog.Tab = 2 Then
      'AdeuDado No Concesional
      If InStr(moddat_g_str_AgrMIHG, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_Agr2FMV, moddat_g_str_CodPrd) > 0 Then '"004" "007" "009" "010" "012" "013" "014" "015" "016" "017" "018"
         crp_Imprim.DataFiles(4) = "TRA_EVACOF"
      End If
      crp_Imprim.SelectionFormula = "{CRE_HIPCUO.HIPCUO_NUMOPE} = '" & moddat_g_str_NumOpe & "' AND {CRE_HIPCUO.HIPCUO_TIPCRO} = 3 "
      
      If InStr(moddat_g_str_AgrTFMV, moddat_g_str_CodPrd) > 0 Then    '"004" "007" "009" "010" "012" "013" "014" "015" "016" "017" "018" "019" "021" "022" "023"
         crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_CROPAG_16.RPT"
      End If
   
   ElseIf tab_Cronog.Tab = 3 Then
      'Adeudado Concesional
      If InStr(moddat_g_str_AgrMIHG, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_Agr2FMV, moddat_g_str_CodPrd) > 0 Then   '"004" "007" "009" "010" "012" "013" "014" "015" "016" "017" "018"
         crp_Imprim.DataFiles(4) = "TRA_EVACOF"
      End If
   
      crp_Imprim.SelectionFormula = "{CRE_HIPCUO.HIPCUO_NUMOPE} = '" & moddat_g_str_NumOpe & "' AND {CRE_HIPCUO.HIPCUO_TIPCRO} = 4 "
      If InStr(moddat_g_str_AgrCRC, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Then     '"001" "003"
         crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_CROPAG_14.RPT"
      ElseIf InStr(moddat_g_str_AgrMIHG, moddat_g_str_CodPrd) Or InStr(moddat_g_str_Agr2FMV, moddat_g_str_CodPrd) > 0 Then   '"004" "007" "009" "010" "012" "013" "014" "015" "016" "017" "018"
         crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_CROPAG_15.RPT"
      End If
   
   ElseIf tab_Cronog.Tab = 4 Then
      'Cofide
      crp_Imprim.DataFiles(4) = "TRA_EVACOF"
      crp_Imprim.SelectionFormula = "{CRE_HIPCUO.HIPCUO_NUMOPE} = '" & moddat_g_str_NumOpe & "' AND {CRE_HIPCUO.HIPCUO_TIPCRO} = 5 "
      crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_CROPAG_17.RPT"
   End If
   
   crp_Imprim.WindowShowPrintSetupBtn = True
'   crp_Imprim.Action = 1
End Sub

Private Sub cmd_ExpExc_Click()
   If tab_Cronog.Tab = 2 And (InStr(moddat_g_str_AgrCRC, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0) Then '"001" "003"
      MsgBox "Estos cronogramas no se imprimen, ya que son de uso interno con Mivivienda.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de exportar el Cronograma?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Select Case CInt(tab_Cronog.Tab)
'      Case 0: Call fs_Exc_Tramo1
'      Case 1: Call fs_Exc_Tramo2
'      Case 2: Call fs_Exc_Tramo3
'      Case 3: Call fs_Exc_Tramo4
      Case 0: Call fs_Exc_Tramo1_Nuevo
      Case 1: Call fs_Exc_Tramo2_Nuevo
      Case 2: Call fs_Exc_Tramo3_Nuevo
      Case 3: Call fs_Exc_Tramo4_Nuevo
      Case 4: Call fs_Exc_Tramo5
   End Select

'Call fs_Exc_TramoPruebas

End Sub

Private Sub cmd_Cronog_Click()
   frm_Con_Cuadre_03.l_str_TipCro = tab_Cronog.Tab + 1
   Select Case tab_Cronog.Tab
      Case 0: frm_Con_Cuadre_03.l_str_NumCuo = grd_CliNCo_Listad.TextMatrix(grd_CliNCo_Listad.Row, 0)
      Case 1: frm_Con_Cuadre_03.l_str_NumCuo = grd_CliCon_Listad.TextMatrix(grd_CliCon_Listad.Row, 0)
      Case 2: frm_Con_Cuadre_03.l_str_NumCuo = grd_MViNCo_Listad.TextMatrix(grd_MViNCo_Listad.Row, 0)
      Case 3: frm_Con_Cuadre_03.l_str_NumCuo = grd_MviCon_Listad.TextMatrix(grd_MviCon_Listad.Row, 0)
      Case 4: frm_Con_Cuadre_03.l_str_NumCuo = grd_CofNCo_Listad.TextMatrix(grd_CofNCo_Listad.Row, 0)
      Case 5: frm_Con_Cuadre_03.l_str_NumCuo = grd_Especial_Cli.TextMatrix(grd_Especial_Cli.Row, 0)
   End Select
   
   frm_Con_Cuadre_03.Show 1
   
   Call fs_Carga_Cro_CliNCo
   Call fs_Carga_Cro_CliCon
   Call fs_Carga_Cro_MViNCo
   Call fs_Carga_Cro_MViCon
   Call fs_Carga_Cro_CofNCo
   Call fs_Carga_Cro_Especial
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Command1_Click()

Call fs_GenExc2

End Sub

Private Sub Command2_Click()
 fs_GenWord
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   pnl_NumOpe.Caption = ""
   pnl_NomCli.Caption = ""
   pnl_NumOpe.Caption = gf_Formato_NumOpe(moddat_g_str_NumOpe)
   pnl_NomCli.Caption = CStr(moddat_g_int_TipDoc) & " - " & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   
   'Determina de donde se obtiene la informacion
   l_int_FlgCro = 1
   pnl_Periodo1.Visible = False
   pnl_Periodo2.Visible = False
   If modmip_g_int_OrdAct = 2 Then
      If CInt(l_str_PerMes) > 0 And CInt(l_str_PerAno) > 2008 Then
         l_int_FlgCro = 2
         pnl_Periodo1.Visible = True
         pnl_Periodo2.Visible = True
         pnl_Periodo2.Caption = Format(CInt(l_str_PerMes), "00") & " / " & Format(CInt(l_str_PerAno), "0000")
      End If
   End If
   
   Call fs_Inicia
   Call fs_Carga_Cro_CliNCo
   tab_Cronog.TabVisible(5) = False
   
   If InStr(moddat_g_str_AgrTMIC, moddat_g_str_CodPrd) > 0 Then 'moddat_g_str_CodPrd = "002" Or moddat_g_str_CodPrd = "006" Or moddat_g_str_CodPrd = "011" Then
      tab_Cronog.TabCaption(0) = "Cliente"
      tab_Cronog.TabVisible(1) = False
      tab_Cronog.TabVisible(2) = False
      tab_Cronog.TabVisible(3) = False
      tab_Cronog.TabVisible(4) = False
      
      If moddat_g_str_CodPrd = "002" Then
         lbl_Totale(0).Caption = "Totales ===> US$ "
      Else
         lbl_Totale(0).Caption = "Totales ===> S/. "
      End If
      If InStr(moddat_g_str_Agr2MIC, moddat_g_str_CodPrd) > 0 Then 'moddat_g_str_CodPrd = "006" Then
         tab_Cronog.TabCaption(1) = "Cliente Concesional"
         tab_Cronog.TabVisible(1) = True
         Call fs_Carga_Cro_CliCon
      End If
      
   ElseIf InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd) > 0 Then 'moddat_g_str_CodPrd = "019" Or moddat_g_str_CodPrd = "020" Or moddat_g_str_CodPrd = "021" Or moddat_g_str_CodPrd = "022" Or moddat_g_str_CodPrd = "023" Then
      tab_Cronog.TabCaption(0) = "Cliente"
      tab_Cronog.TabCaption(2) = "Cofide"
      tab_Cronog.TabVisible(1) = False
      tab_Cronog.TabVisible(2) = True
      tab_Cronog.TabVisible(3) = False
      tab_Cronog.TabVisible(4) = False
      lbl_Totale(0).Caption = "Totales ===> S/. "
      lbl_Totale(1).Caption = "Totales ===> S/. "
      lbl_Totale(2).Caption = "Totales ===> S/. "
      lbl_Totale(3).Caption = "Totales ===> S/. "
      
      Call fs_Carga_Cro_MViNCo
   
   Else
      tab_Cronog.TabCaption(0) = "Cliente No Concesional"
      tab_Cronog.TabCaption(1) = "Cliente Concesional"
      
      If InStr(moddat_g_str_AgrCRC, moddat_g_str_CodPrd) > 0 Then 'moddat_g_str_CodPrd = "001" Then
         tab_Cronog.TabCaption(2) = "FMV No Concesional"
         tab_Cronog.TabCaption(3) = "FMV Concesional"
         tab_Cronog.TabVisible(4) = False
         lbl_Totale(0).Caption = "Totales ===> US$ "
         lbl_Totale(1).Caption = "Totales ===> US$ "
         lbl_Totale(2).Caption = "Totales ===> US$ "
         lbl_Totale(3).Caption = "Totales ===> US$ "
         
      ElseIf InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Then 'moddat_g_str_CodPrd = "003" Then
         tab_Cronog.TabCaption(2) = "FMV No Concesional"
         tab_Cronog.TabCaption(3) = "FMV Concesional"
         tab_Cronog.TabCaption(4) = "Cofide"
         tab_Cronog.TabVisible(4) = True
         lbl_Totale(0).Caption = "Totales ===> S/. "
         lbl_Totale(1).Caption = "Totales ===> S/. "
         lbl_Totale(2).Caption = "Totales ===> S/. "
         lbl_Totale(3).Caption = "Totales ===> S/. "
         lbl_Totale(4).Caption = "Totales ===> S/. "
      
      ElseIf InStr(moddat_g_str_AgrMIHG, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_Agr2FMV, moddat_g_str_CodPrd) > 0 Then 'moddat_g_str_CodPrd = "004" Or moddat_g_str_CodPrd = "007" Or moddat_g_str_CodPrd = "009" Or moddat_g_str_CodPrd = "010" Or moddat_g_str_CodPrd = "013" Or moddat_g_str_CodPrd = "014" Or moddat_g_str_CodPrd = "015" Or moddat_g_str_CodPrd = "016" Or moddat_g_str_CodPrd = "017" Or moddat_g_str_CodPrd = "018" Then
         tab_Cronog.TabCaption(2) = "Cofide No Concesional"
         tab_Cronog.TabCaption(3) = "Cofide Concesional"
         tab_Cronog.TabVisible(4) = False
         lbl_Totale(0).Caption = "Totales ===> S/. "
         lbl_Totale(1).Caption = "Totales ===> S/. "
         lbl_Totale(2).Caption = "Totales ===> S/. "
         lbl_Totale(3).Caption = "Totales ===> S/. "
         
      End If
      
      Call fs_Carga_Cro_CliCon
      Call fs_Carga_Cro_MViNCo
      Call fs_Carga_Cro_MViCon
      
      If InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Then 'moddat_g_str_CodPrd = "003" Then
         Call fs_Carga_Cro_CofNCo
            
         If moddat_g_str_NumOpe = "0030800036" Then
            tab_Cronog.TabVisible(5) = True
            Call fs_Carga_Cro_Especial
         End If
      End If
   End If
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   'Cliente No Concesional
   grd_CliNCo_Listad.Cols = 10
   grd_CliNCo_Listad.ColWidth(0) = 795
   grd_CliNCo_Listad.ColWidth(1) = 1425
   grd_CliNCo_Listad.ColWidth(2) = 1180
   grd_CliNCo_Listad.ColWidth(3) = 1170
   grd_CliNCo_Listad.ColWidth(4) = 1160
   grd_CliNCo_Listad.ColWidth(5) = 1160
   grd_CliNCo_Listad.ColWidth(6) = 1160
   grd_CliNCo_Listad.ColWidth(7) = 1320
   grd_CliNCo_Listad.ColWidth(8) = 1560
   grd_CliNCo_Listad.ColWidth(9) = 0
   grd_CliNCo_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_CliNCo_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_CliNCo_Listad.ColAlignment(2) = flexAlignRightCenter
   grd_CliNCo_Listad.ColAlignment(3) = flexAlignRightCenter
   grd_CliNCo_Listad.ColAlignment(4) = flexAlignRightCenter
   grd_CliNCo_Listad.ColAlignment(5) = flexAlignRightCenter
   grd_CliNCo_Listad.ColAlignment(6) = flexAlignRightCenter
   grd_CliNCo_Listad.ColAlignment(7) = flexAlignRightCenter
   grd_CliNCo_Listad.ColAlignment(8) = flexAlignRightCenter
   grd_CliNCo_Listad.ColAlignment(9) = flexAlignLeftCenter

   'Cliente Concesional
   grd_CliCon_Listad.Cols = 7
   grd_CliCon_Listad.ColWidth(0) = 770
   grd_CliCon_Listad.ColWidth(1) = 1485
   grd_CliCon_Listad.ColWidth(2) = 2170
   grd_CliCon_Listad.ColWidth(3) = 2160
   grd_CliCon_Listad.ColWidth(4) = 2170
   grd_CliCon_Listad.ColWidth(5) = 2160
   grd_CliCon_Listad.ColWidth(6) = 0
   grd_CliCon_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_CliCon_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_CliCon_Listad.ColAlignment(2) = flexAlignRightCenter
   grd_CliCon_Listad.ColAlignment(3) = flexAlignRightCenter
   grd_CliCon_Listad.ColAlignment(4) = flexAlignRightCenter
   grd_CliCon_Listad.ColAlignment(5) = flexAlignRightCenter
   grd_CliCon_Listad.ColAlignment(6) = flexAlignLeftCenter
   
   'Mivivienda No Concesional
   grd_MViNCo_Listad.Cols = 11
   grd_MViNCo_Listad.ColWidth(0) = 695
   grd_MViNCo_Listad.ColWidth(1) = 1415
   grd_MViNCo_Listad.ColWidth(2) = 1070
   grd_MViNCo_Listad.ColWidth(3) = 1070
   grd_MViNCo_Listad.ColWidth(4) = 1080
   grd_MViNCo_Listad.ColWidth(5) = 1080
   grd_MViNCo_Listad.ColWidth(6) = 1080
   grd_MViNCo_Listad.ColWidth(7) = 1080
   grd_MViNCo_Listad.ColWidth(8) = 1080
   grd_MViNCo_Listad.ColWidth(9) = 1290
   grd_MViNCo_Listad.ColWidth(10) = 0
   grd_MViNCo_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_MViNCo_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_MViNCo_Listad.ColAlignment(2) = flexAlignRightCenter
   grd_MViNCo_Listad.ColAlignment(3) = flexAlignRightCenter
   grd_MViNCo_Listad.ColAlignment(4) = flexAlignRightCenter
   grd_MViNCo_Listad.ColAlignment(5) = flexAlignRightCenter
   grd_MViNCo_Listad.ColAlignment(6) = flexAlignRightCenter
   grd_MViNCo_Listad.ColAlignment(7) = flexAlignRightCenter
   grd_MViNCo_Listad.ColAlignment(8) = flexAlignRightCenter
   grd_MViNCo_Listad.ColAlignment(9) = flexAlignRightCenter
   grd_MViNCo_Listad.ColAlignment(10) = flexAlignLeftCenter

   'Mivivienda Concesional
   grd_MviCon_Listad.Cols = 8
   grd_MviCon_Listad.ColWidth(0) = 770
   grd_MviCon_Listad.ColWidth(1) = 1485
   grd_MviCon_Listad.ColWidth(2) = 1730
   grd_MviCon_Listad.ColWidth(3) = 1740
   grd_MviCon_Listad.ColWidth(4) = 1740
   grd_MviCon_Listad.ColWidth(5) = 1740
   grd_MviCon_Listad.ColWidth(6) = 1730
   grd_MviCon_Listad.ColWidth(7) = 0
   grd_MviCon_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_MviCon_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_MviCon_Listad.ColAlignment(2) = flexAlignRightCenter
   grd_MviCon_Listad.ColAlignment(3) = flexAlignRightCenter
   grd_MviCon_Listad.ColAlignment(4) = flexAlignRightCenter
   grd_MviCon_Listad.ColAlignment(5) = flexAlignRightCenter
   grd_MviCon_Listad.ColAlignment(6) = flexAlignRightCenter
   grd_MviCon_Listad.ColAlignment(7) = flexAlignLeftCenter
   
   'Cofide No Concesional
   grd_CofNCo_Listad.ColWidth(0) = 770
   grd_CofNCo_Listad.ColWidth(1) = 1485
   grd_CofNCo_Listad.ColWidth(2) = 1730
   grd_CofNCo_Listad.ColWidth(3) = 1740
   grd_CofNCo_Listad.ColWidth(4) = 1740
   grd_CofNCo_Listad.ColWidth(5) = 1740
   grd_CofNCo_Listad.ColWidth(6) = 1730
   grd_CofNCo_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_CofNCo_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_CofNCo_Listad.ColAlignment(2) = flexAlignRightCenter
   grd_CofNCo_Listad.ColAlignment(3) = flexAlignRightCenter
   grd_CofNCo_Listad.ColAlignment(4) = flexAlignRightCenter
   grd_CofNCo_Listad.ColAlignment(5) = flexAlignRightCenter
   grd_CofNCo_Listad.ColAlignment(6) = flexAlignRightCenter
   
   'Especial (ABAD LOLI)
   grd_Especial_Cli.ColWidth(0) = 900
   grd_Especial_Cli.ColWidth(1) = 1740
   grd_Especial_Cli.ColWidth(2) = 1590
   grd_Especial_Cli.ColWidth(3) = 1590
   grd_Especial_Cli.ColWidth(4) = 1410
   grd_Especial_Cli.ColWidth(5) = 1860
   grd_Especial_Cli.ColWidth(6) = 1760
   grd_Especial_Cli.ColAlignment(0) = flexAlignCenterCenter
   grd_Especial_Cli.ColAlignment(1) = flexAlignCenterCenter
   grd_Especial_Cli.ColAlignment(2) = flexAlignRightCenter
   grd_Especial_Cli.ColAlignment(3) = flexAlignRightCenter
   grd_Especial_Cli.ColAlignment(4) = flexAlignRightCenter
   grd_Especial_Cli.ColAlignment(5) = flexAlignRightCenter
   grd_Especial_Cli.ColAlignment(6) = flexAlignRightCenter
End Sub

Private Sub fs_Carga_Cro_CliNCo()
Dim r_dbl_Capita     As Double
Dim r_dbl_Intere     As Double
Dim r_dbl_SegDes     As Double
Dim r_dbl_SegViv     As Double
Dim r_dbl_OtrCar     As Double
Dim r_dbl_ImpCuo     As Double
Dim r_dbl_TotCuo     As Double
      
   Call gs_LimpiaGrid(grd_CliNCo_Listad)
   r_dbl_Capita = 0
   r_dbl_Intere = 0
   r_dbl_SegDes = 0
   r_dbl_SegViv = 0
   r_dbl_OtrCar = 0
   r_dbl_TotCuo = 0
   
   If l_int_FlgCro = 1 Then
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT HIPCUO_NUMCUO, HIPCUO_FECVCT, HIPCUO_CAPITA, HIPCUO_INTERE,  "
      g_str_Parame = g_str_Parame & "       HIPCUO_DESORG, HIPCUO_VIVORG, HIPCUO_OTRORG, HIPCUO_SALCAP, HIPCUO_SITUAC  "
      g_str_Parame = g_str_Parame & "  FROM CRE_HIPCUO "
      g_str_Parame = g_str_Parame & " WHERE HIPCUO_NUMOPE = '" & moddat_g_str_NumOpe & "' "
      g_str_Parame = g_str_Parame & "   AND HIPCUO_TIPCRO = 1 "
      g_str_Parame = g_str_Parame & " ORDER BY HIPCUO_NUMCUO "
   Else
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT CUOCIE_NUMCUO AS HIPCUO_NUMCUO, CUOCIE_FECVCT AS HIPCUO_FECVCT, CUOCIE_CAPITA AS HIPCUO_CAPITA, "
      g_str_Parame = g_str_Parame & "       CUOCIE_INTERE AS HIPCUO_INTERE, CUOCIE_DESORG AS HIPCUO_DESORG, CUOCIE_VIVORG AS HIPCUO_VIVORG, "
      g_str_Parame = g_str_Parame & "       CUOCIE_OTRORG AS HIPCUO_OTRORG, CUOCIE_SALCAP AS HIPCUO_SALCAP, CUOCIE_SITUAC AS HIPCUO_SITUAC "
      g_str_Parame = g_str_Parame & "  FROM CRE_CUOCIE "
      g_str_Parame = g_str_Parame & " WHERE CUOCIE_PERMES = " & l_str_PerMes & " "
      g_str_Parame = g_str_Parame & "   AND CUOCIE_PERANO = " & l_str_PerAno & " "
      g_str_Parame = g_str_Parame & "   AND CUOCIE_NUMOPE = '" & moddat_g_str_NumOpe & "' "
      g_str_Parame = g_str_Parame & "   AND CUOCIE_TIPCRO = 1 "
      g_str_Parame = g_str_Parame & " ORDER BY CUOCIE_NUMCUO "
   End If
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_CliNCo_Listad.Redraw = False
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         grd_CliNCo_Listad.Rows = grd_CliNCo_Listad.Rows + 1
         grd_CliNCo_Listad.Row = grd_CliNCo_Listad.Rows - 1
         r_dbl_ImpCuo = 0
         
         grd_CliNCo_Listad.Col = 0
         grd_CliNCo_Listad.Text = Format(g_rst_Princi!HIPCUO_NUMCUO, "000")
      
         grd_CliNCo_Listad.Col = 1
         grd_CliNCo_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECVCT))
         
         grd_CliNCo_Listad.Col = 2
         grd_CliNCo_Listad.Text = Format(g_rst_Princi!HIPCUO_CAPITA, "###,###,##0.00")
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(grd_CliNCo_Listad.Text)
         
         grd_CliNCo_Listad.Col = 3
         grd_CliNCo_Listad.Text = Format(g_rst_Princi!HIPCUO_INTERE, "###,###,##0.00")
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(grd_CliNCo_Listad.Text)
         
         grd_CliNCo_Listad.Col = 4
         grd_CliNCo_Listad.Text = Format(g_rst_Princi!HIPCUO_DESORG, "###,###,##0.00")
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(grd_CliNCo_Listad.Text)
         
         grd_CliNCo_Listad.Col = 5
         grd_CliNCo_Listad.Text = Format(g_rst_Princi!HIPCUO_VIVORG, "###,###,##0.00")
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(grd_CliNCo_Listad.Text)
         
         grd_CliNCo_Listad.Col = 6
         grd_CliNCo_Listad.Text = Format(g_rst_Princi!HIPCUO_OTRORG, "###,###,##0.00")
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(grd_CliNCo_Listad.Text)
         
         grd_CliNCo_Listad.Col = 7
         grd_CliNCo_Listad.Text = Format(r_dbl_ImpCuo, "###,###,##0.00")
                  
         grd_CliNCo_Listad.Col = 8
         grd_CliNCo_Listad.Text = Format(g_rst_Princi!HIPCUO_SALCAP, "###,###,##0.00")
         
         grd_CliNCo_Listad.Col = 9
         grd_CliNCo_Listad.Text = IIf(CStr(g_rst_Princi!HIPCUO_SITUAC) = 2, "POR VENCER", "PAGADA")

         r_dbl_Capita = r_dbl_Capita + CDbl(Format(g_rst_Princi!HIPCUO_CAPITA, "###,###,##0.00"))
         r_dbl_Intere = r_dbl_Intere + CDbl(Format(g_rst_Princi!HIPCUO_INTERE, "###,###,##0.00"))
         r_dbl_SegDes = r_dbl_SegDes + CDbl(Format(g_rst_Princi!HIPCUO_DESORG, "###,###,##0.00"))
         r_dbl_SegViv = r_dbl_SegViv + CDbl(Format(g_rst_Princi!HIPCUO_VIVORG, "###,###,##0.00"))
         r_dbl_OtrCar = r_dbl_OtrCar + CDbl(Format(g_rst_Princi!HIPCUO_OTRORG, "###,###,##0.00"))
         r_dbl_TotCuo = r_dbl_TotCuo + CDbl(Format(r_dbl_ImpCuo, "###,###,##0.00"))
         g_rst_Princi.MoveNext
      Loop
      
      grd_CliNCo_Listad.Redraw = True
      Call gs_UbiIniGrid(grd_CliNCo_Listad)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   pnl_CliNCo_Capita.Caption = Format(r_dbl_Capita, "###,###,##0.00") & " "
   pnl_CliNCo_Intere.Caption = Format(r_dbl_Intere, "###,###,##0.00") & " "
   pnl_CliNCo_SegPre.Caption = Format(r_dbl_SegDes, "###,###,##0.00") & " "
   pnl_CliNCo_SegViv.Caption = Format(r_dbl_SegViv, "###,###,##0.00") & " "
   pnl_CliNCo_OtrCar.Caption = Format(r_dbl_OtrCar, "###,###,##0.00") & " "
   pnl_CliNCo_TotCuo.Caption = Format(r_dbl_TotCuo, "###,###,##0.00") & " "
End Sub

Private Sub fs_Carga_Cro_CliCon()
Dim r_dbl_Capita     As Double
Dim r_dbl_Intere     As Double
Dim r_dbl_ImpCuo     As Double
Dim r_dbl_TotCuo     As Double
   
   Call gs_LimpiaGrid(grd_CliCon_Listad)
   r_dbl_Capita = 0
   r_dbl_Intere = 0
   r_dbl_TotCuo = 0
   
   If l_int_FlgCro = 1 Then
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT HIPCUO_NUMCUO, HIPCUO_FECVCT, HIPCUO_CAPITA, HIPCUO_INTERE,  "
      g_str_Parame = g_str_Parame & "       HIPCUO_DESORG, HIPCUO_VIVORG, HIPCUO_OTRORG, HIPCUO_SALCAP, HIPCUO_SITUAC "
      g_str_Parame = g_str_Parame & "  FROM CRE_HIPCUO "
      g_str_Parame = g_str_Parame & " WHERE HIPCUO_NUMOPE = '" & moddat_g_str_NumOpe & "' "
      g_str_Parame = g_str_Parame & "   AND HIPCUO_TIPCRO = 2 "
      g_str_Parame = g_str_Parame & " ORDER BY HIPCUO_NUMCUO "
   Else
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT CUOCIE_NUMCUO AS HIPCUO_NUMCUO, CUOCIE_FECVCT AS HIPCUO_FECVCT, CUOCIE_CAPITA AS HIPCUO_CAPITA, "
      g_str_Parame = g_str_Parame & "       CUOCIE_INTERE AS HIPCUO_INTERE, CUOCIE_DESORG AS HIPCUO_DESORG, CUOCIE_VIVORG AS HIPCUO_VIVORG, "
      g_str_Parame = g_str_Parame & "       CUOCIE_OTRORG AS HIPCUO_OTRORG, CUOCIE_SALCAP AS HIPCUO_SALCAP, CUOCIE_SITUAC AS HIPCUO_SITUAC "
      g_str_Parame = g_str_Parame & "  FROM CRE_CUOCIE "
      g_str_Parame = g_str_Parame & " WHERE CUOCIE_PERMES = " & l_str_PerMes & " "
      g_str_Parame = g_str_Parame & "   AND CUOCIE_PERANO = " & l_str_PerAno & " "
      g_str_Parame = g_str_Parame & "   AND CUOCIE_NUMOPE = '" & moddat_g_str_NumOpe & "' "
      g_str_Parame = g_str_Parame & "   AND CUOCIE_TIPCRO = 2 "
      g_str_Parame = g_str_Parame & " ORDER BY CUOCIE_NUMCUO "
   End If
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_CliCon_Listad.Redraw = False
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         grd_CliCon_Listad.Rows = grd_CliCon_Listad.Rows + 1
         grd_CliCon_Listad.Row = grd_CliCon_Listad.Rows - 1
         r_dbl_ImpCuo = 0
         
         grd_CliCon_Listad.Col = 0
         grd_CliCon_Listad.Text = Format(g_rst_Princi!HIPCUO_NUMCUO, "000")
      
         grd_CliCon_Listad.Col = 1
         grd_CliCon_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECVCT))
         
         grd_CliCon_Listad.Col = 2
         grd_CliCon_Listad.Text = Format(g_rst_Princi!HIPCUO_CAPITA, "###,###,##0.00")
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(grd_CliCon_Listad.Text)
         
         grd_CliCon_Listad.Col = 3
         grd_CliCon_Listad.Text = Format(g_rst_Princi!HIPCUO_INTERE, "###,###,##0.00")
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(grd_CliCon_Listad.Text)
         
         grd_CliCon_Listad.Col = 4
         grd_CliCon_Listad.Text = Format(r_dbl_ImpCuo, "###,###,##0.00")
         
         grd_CliCon_Listad.Col = 5
         grd_CliCon_Listad.Text = Format(g_rst_Princi!HIPCUO_SALCAP, "###,###,##0.00")

         grd_CliCon_Listad.Col = 6
         grd_CliCon_Listad.Text = IIf(CStr(g_rst_Princi!HIPCUO_SITUAC) = 2, "POR VENCER", "PAGADA")
         
         r_dbl_Capita = r_dbl_Capita + CDbl(Format(g_rst_Princi!HIPCUO_CAPITA, "###,###,##0.00"))
         r_dbl_Intere = r_dbl_Intere + CDbl(Format(g_rst_Princi!HIPCUO_INTERE, "###,###,##0.00"))
         r_dbl_TotCuo = r_dbl_TotCuo + CDbl(Format(r_dbl_ImpCuo, "###,###,##0.00"))
         g_rst_Princi.MoveNext
      Loop
      
      grd_CliCon_Listad.Redraw = True
      Call gs_UbiIniGrid(grd_CliCon_Listad)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   pnl_CliCon_Capita.Caption = Format(r_dbl_Capita, "###,###,##0.00") & " "
   pnl_CliCon_Intere.Caption = Format(r_dbl_Intere, "###,###,##0.00") & " "
   pnl_CliCon_TotCuo.Caption = Format(r_dbl_TotCuo, "###,###,##0.00") & " "
End Sub

Private Sub fs_Carga_Cro_MViNCo()
Dim r_dbl_Capita     As Double
Dim r_dbl_Intere     As Double
Dim r_dbl_SegDes     As Double
Dim r_dbl_SegViv     As Double
Dim r_dbl_OtrCar     As Double
Dim r_dbl_Comisi     As Double
Dim r_dbl_ImpCuo     As Double
Dim r_dbl_TotCuo     As Double
   
   Call gs_LimpiaGrid(grd_MViNCo_Listad)
   r_dbl_Capita = 0
   r_dbl_Intere = 0
   r_dbl_SegDes = 0
   r_dbl_SegViv = 0
   r_dbl_OtrCar = 0
   r_dbl_Comisi = 0
   r_dbl_TotCuo = 0
   
   If l_int_FlgCro = 1 Then
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT HIPCUO_NUMCUO, HIPCUO_FECVCT, HIPCUO_CAPITA, HIPCUO_INTERE, HIPCUO_DESORG,  "
      g_str_Parame = g_str_Parame & "       HIPCUO_VIVORG, HIPCUO_OTRORG, HIPCUO_SALCAP, HIPCUO_COMCOF, HIPCUO_SITUAC  "
      g_str_Parame = g_str_Parame & "  FROM CRE_HIPCUO "
      g_str_Parame = g_str_Parame & " WHERE HIPCUO_NUMOPE = '" & moddat_g_str_NumOpe & "' "
      g_str_Parame = g_str_Parame & "   AND HIPCUO_TIPCRO = 3 "
      g_str_Parame = g_str_Parame & " ORDER BY HIPCUO_NUMCUO "
   Else
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT CUOCIE_NUMCUO AS HIPCUO_NUMCUO, CUOCIE_FECVCT AS HIPCUO_FECVCT, CUOCIE_CAPITA AS HIPCUO_CAPITA, "
      g_str_Parame = g_str_Parame & "       CUOCIE_INTERE AS HIPCUO_INTERE, CUOCIE_DESORG AS HIPCUO_DESORG, CUOCIE_VIVORG AS HIPCUO_VIVORG, "
      g_str_Parame = g_str_Parame & "       CUOCIE_OTRORG AS HIPCUO_OTRORG, CUOCIE_SALCAP AS HIPCUO_SALCAP, CUOCIE_COMCOF AS HIPCUO_COMCOF, "
      g_str_Parame = g_str_Parame & "       CUOCIE_SITUAC AS HIPCUO_SITUAC  "
      g_str_Parame = g_str_Parame & "  FROM CRE_CUOCIE "
      g_str_Parame = g_str_Parame & " WHERE CUOCIE_PERMES = " & l_str_PerMes & " "
      g_str_Parame = g_str_Parame & "   AND CUOCIE_PERANO = " & l_str_PerAno & " "
      g_str_Parame = g_str_Parame & "   AND CUOCIE_NUMOPE = '" & moddat_g_str_NumOpe & "' "
      g_str_Parame = g_str_Parame & "   AND CUOCIE_TIPCRO = 3 "
      g_str_Parame = g_str_Parame & " ORDER BY CUOCIE_NUMCUO "
   End If
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_MViNCo_Listad.Redraw = False
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         grd_MViNCo_Listad.Rows = grd_MViNCo_Listad.Rows + 1
         grd_MViNCo_Listad.Row = grd_MViNCo_Listad.Rows - 1
         r_dbl_ImpCuo = 0
         
         grd_MViNCo_Listad.Col = 0
         grd_MViNCo_Listad.Text = Format(g_rst_Princi!HIPCUO_NUMCUO, "000")
         
         grd_MViNCo_Listad.Col = 1
         grd_MViNCo_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECVCT))
         
         grd_MViNCo_Listad.Col = 2
         grd_MViNCo_Listad.Text = Format(g_rst_Princi!HIPCUO_CAPITA, "###,###,##0.00")
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(grd_MViNCo_Listad.Text)
         
         grd_MViNCo_Listad.Col = 3
         grd_MViNCo_Listad.Text = Format(g_rst_Princi!HIPCUO_INTERE, "###,###,##0.00")
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(grd_MViNCo_Listad.Text)
         
         grd_MViNCo_Listad.Col = 4
         grd_MViNCo_Listad.Text = Format(g_rst_Princi!HIPCUO_DESORG, "###,###,##0.00")
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(grd_MViNCo_Listad.Text)
         
         grd_MViNCo_Listad.Col = 5
         grd_MViNCo_Listad.Text = Format(g_rst_Princi!HIPCUO_VIVORG, "###,###,##0.00")
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(grd_MViNCo_Listad.Text)
         
         grd_MViNCo_Listad.Col = 6
         grd_MViNCo_Listad.Text = Format(g_rst_Princi!HIPCUO_OTRORG, "###,###,##0.00")
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(grd_MViNCo_Listad.Text)
         
         grd_MViNCo_Listad.Col = 7
         grd_MViNCo_Listad.Text = Format(g_rst_Princi!HIPCUO_COMCOF, "###,###,##0.00")
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(grd_MViNCo_Listad.Text)
         
         grd_MViNCo_Listad.Col = 8
         grd_MViNCo_Listad.Text = Format(r_dbl_ImpCuo, "###,###,##0.00")
         
         grd_MViNCo_Listad.Col = 9
         grd_MViNCo_Listad.Text = Format(g_rst_Princi!HIPCUO_SALCAP, "###,###,##0.00")
         
         grd_MViNCo_Listad.Col = 10
         grd_MViNCo_Listad.Text = IIf(CStr(g_rst_Princi!HIPCUO_SITUAC) = 2, "POR VENCER", "PAGADA")
         
         r_dbl_Capita = r_dbl_Capita + CDbl(Format(g_rst_Princi!HIPCUO_CAPITA, "###,###,##0.00"))
         r_dbl_Intere = r_dbl_Intere + CDbl(Format(g_rst_Princi!HIPCUO_INTERE, "###,###,##0.00"))
         r_dbl_SegDes = r_dbl_SegDes + CDbl(Format(g_rst_Princi!HIPCUO_DESORG, "###,###,##0.00"))
         r_dbl_SegViv = r_dbl_SegViv + CDbl(Format(g_rst_Princi!HIPCUO_VIVORG, "###,###,##0.00"))
         r_dbl_OtrCar = r_dbl_OtrCar + CDbl(Format(g_rst_Princi!HIPCUO_OTRORG, "###,###,##0.00"))
         r_dbl_Comisi = r_dbl_Comisi + CDbl(Format(g_rst_Princi!HIPCUO_COMCOF, "###,###,##0.00"))
         r_dbl_TotCuo = r_dbl_TotCuo + CDbl(Format(r_dbl_ImpCuo, "###,###,##0.00"))
         g_rst_Princi.MoveNext
      Loop
      
      grd_MViNCo_Listad.Redraw = True
      Call gs_UbiIniGrid(grd_MViNCo_Listad)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   pnl_MViNCo_Capita.Caption = Format(r_dbl_Capita, "###,###,##0.00") & " "
   pnl_MViNCo_Intere.Caption = Format(r_dbl_Intere, "###,###,##0.00") & " "
   pnl_MViNCo_SegPre.Caption = Format(r_dbl_SegDes, "###,###,##0.00") & " "
   pnl_MViNCo_SegViv.Caption = Format(r_dbl_SegViv, "###,###,##0.00") & " "
   pnl_MViNCo_OtrCar.Caption = Format(r_dbl_OtrCar, "###,###,##0.00") & " "
   pnl_MViNCo_Comisi.Caption = Format(r_dbl_Comisi, "###,###,##0.00") & " "
   pnl_MViNCo_TotCuo.Caption = Format(r_dbl_TotCuo, "###,###,##0.00") & " "
End Sub

Private Sub fs_Carga_Cro_MViCon()
Dim r_dbl_Capita     As Double
Dim r_dbl_Intere     As Double
Dim r_dbl_Comisi     As Double
Dim r_dbl_ImpCuo     As Double
Dim r_dbl_TotCuo     As Double
   
   Call gs_LimpiaGrid(grd_MviCon_Listad)
   r_dbl_Capita = 0
   r_dbl_Intere = 0
   r_dbl_Comisi = 0
   r_dbl_TotCuo = 0
   
   If l_int_FlgCro = 1 Then
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT HIPCUO_NUMCUO, HIPCUO_FECVCT, HIPCUO_CAPITA, HIPCUO_INTERE, HIPCUO_DESORG, "
      g_str_Parame = g_str_Parame & "       HIPCUO_VIVORG, HIPCUO_OTRORG, HIPCUO_SALCAP, HIPCUO_COMCOF, HIPCUO_SITUAC  "
      g_str_Parame = g_str_Parame & "  FROM CRE_HIPCUO "
      g_str_Parame = g_str_Parame & " WHERE HIPCUO_NUMOPE = '" & moddat_g_str_NumOpe & "' "
      g_str_Parame = g_str_Parame & "   AND HIPCUO_TIPCRO = 4 "
      g_str_Parame = g_str_Parame & " ORDER BY HIPCUO_NUMCUO "
   Else
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT CUOCIE_NUMCUO AS HIPCUO_NUMCUO, CUOCIE_FECVCT AS HIPCUO_FECVCT, CUOCIE_CAPITA AS HIPCUO_CAPITA, "
      g_str_Parame = g_str_Parame & "       CUOCIE_INTERE AS HIPCUO_INTERE, CUOCIE_DESORG AS HIPCUO_DESORG, CUOCIE_VIVORG AS HIPCUO_VIVORG, "
      g_str_Parame = g_str_Parame & "       CUOCIE_OTRORG AS HIPCUO_OTRORG, CUOCIE_SALCAP AS HIPCUO_SALCAP, CUOCIE_COMCOF AS HIPCUO_COMCOF, "
      g_str_Parame = g_str_Parame & "       CUOCIE_SITUAC AS HIPCUO_SITUAC "
      g_str_Parame = g_str_Parame & "  FROM CRE_CUOCIE "
      g_str_Parame = g_str_Parame & " WHERE CUOCIE_PERMES = " & l_str_PerMes & " "
      g_str_Parame = g_str_Parame & "   AND CUOCIE_PERANO = " & l_str_PerAno & " "
      g_str_Parame = g_str_Parame & "   AND CUOCIE_NUMOPE = '" & moddat_g_str_NumOpe & "' "
      g_str_Parame = g_str_Parame & "   AND CUOCIE_TIPCRO = 4 "
      g_str_Parame = g_str_Parame & " ORDER BY CUOCIE_NUMCUO "
   End If
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_MviCon_Listad.Redraw = False
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         grd_MviCon_Listad.Rows = grd_MviCon_Listad.Rows + 1
         grd_MviCon_Listad.Row = grd_MviCon_Listad.Rows - 1
         r_dbl_ImpCuo = 0
         
         grd_MviCon_Listad.Col = 0
         grd_MviCon_Listad.Text = Format(g_rst_Princi!HIPCUO_NUMCUO, "000")
      
         grd_MviCon_Listad.Col = 1
         grd_MviCon_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECVCT))
         
         grd_MviCon_Listad.Col = 2
         grd_MviCon_Listad.Text = Format(g_rst_Princi!HIPCUO_CAPITA, "###,###,##0.00")
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(grd_MviCon_Listad.Text)
         
         grd_MviCon_Listad.Col = 3
         grd_MviCon_Listad.Text = Format(g_rst_Princi!HIPCUO_INTERE, "###,###,##0.00")
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(grd_MviCon_Listad.Text)
         
         grd_MviCon_Listad.Col = 4
         grd_MviCon_Listad.Text = Format(g_rst_Princi!HIPCUO_COMCOF, "###,###,##0.00")
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(grd_MviCon_Listad.Text)
         
         grd_MviCon_Listad.Col = 5
         grd_MviCon_Listad.Text = Format(r_dbl_ImpCuo, "###,###,##0.00")
         
         grd_MviCon_Listad.Col = 6
         grd_MviCon_Listad.Text = Format(g_rst_Princi!HIPCUO_SALCAP, "###,###,##0.00")

         grd_MviCon_Listad.Col = 7
         grd_MviCon_Listad.Text = IIf(CStr(g_rst_Princi!HIPCUO_SITUAC) = 2, "POR VENCER", "PAGADA")
         
         r_dbl_Capita = r_dbl_Capita + CDbl(Format(g_rst_Princi!HIPCUO_CAPITA, "###,###,##0.00"))
         r_dbl_Intere = r_dbl_Intere + CDbl(Format(g_rst_Princi!HIPCUO_INTERE, "###,###,##0.00"))
         r_dbl_Comisi = r_dbl_Comisi + CDbl(Format(g_rst_Princi!HIPCUO_COMCOF, "###,###,##0.00"))
         r_dbl_TotCuo = r_dbl_TotCuo + CDbl(Format(r_dbl_ImpCuo, "###,###,##0.00"))
         g_rst_Princi.MoveNext
      Loop
      
      grd_MviCon_Listad.Redraw = True
      Call gs_UbiIniGrid(grd_MviCon_Listad)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   pnl_MViCon_Capita.Caption = Format(r_dbl_Capita, "###,###,##0.00") & " "
   pnl_MViCon_Intere.Caption = Format(r_dbl_Intere, "###,###,##0.00") & " "
   pnl_MViCon_Comisi.Caption = Format(r_dbl_Comisi, "###,###,##0.00") & " "
   pnl_MViCon_TotCuo.Caption = Format(r_dbl_TotCuo, "###,###,##0.00") & " "
End Sub

Private Sub fs_Carga_Cro_CofNCo()
Dim r_dbl_Capita     As Double
Dim r_dbl_Intere     As Double
Dim r_dbl_Comisi     As Double
Dim r_dbl_ImpCuo     As Double
Dim r_dbl_TotCuo     As Double
   
   Call gs_LimpiaGrid(grd_CofNCo_Listad)
   r_dbl_Capita = 0
   r_dbl_Intere = 0
   r_dbl_Comisi = 0
   r_dbl_TotCuo = 0
   
   If l_int_FlgCro = 1 Then
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT HIPCUO_NUMCUO, HIPCUO_FECVCT, HIPCUO_CAPITA, HIPCUO_INTERE,  "
      g_str_Parame = g_str_Parame & "       HIPCUO_DESORG, HIPCUO_VIVORG, HIPCUO_OTRORG, HIPCUO_SALCAP, HIPCUO_COMCOF  "
      g_str_Parame = g_str_Parame & "  FROM CRE_HIPCUO "
      g_str_Parame = g_str_Parame & " WHERE HIPCUO_NUMOPE = '" & moddat_g_str_NumOpe & "' "
      g_str_Parame = g_str_Parame & "   AND HIPCUO_TIPCRO = 5 "
      g_str_Parame = g_str_Parame & " ORDER BY HIPCUO_NUMCUO "
   Else
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT CUOCIE_NUMCUO AS HIPCUO_NUMCUO, CUOCIE_FECVCT AS HIPCUO_FECVCT, CUOCIE_CAPITA AS HIPCUO_CAPITA, "
      g_str_Parame = g_str_Parame & "       CUOCIE_INTERE AS HIPCUO_INTERE, CUOCIE_DESORG AS HIPCUO_DESORG, CUOCIE_VIVORG AS HIPCUO_VIVORG, "
      g_str_Parame = g_str_Parame & "       CUOCIE_OTRORG AS HIPCUO_OTRORG, CUOCIE_SALCAP AS HIPCUO_SALCAP, CUOCIE_COMCOF AS HIPCUO_COMCOF "
      g_str_Parame = g_str_Parame & "  FROM CRE_CUOCIE "
      g_str_Parame = g_str_Parame & " WHERE CUOCIE_PERMES = " & l_str_PerMes & " "
      g_str_Parame = g_str_Parame & "   AND CUOCIE_PERANO = " & l_str_PerAno & " "
      g_str_Parame = g_str_Parame & "   AND CUOCIE_NUMOPE = '" & moddat_g_str_NumOpe & "' "
      g_str_Parame = g_str_Parame & "   AND CUOCIE_TIPCRO = 5 "
      g_str_Parame = g_str_Parame & " ORDER BY CUOCIE_NUMCUO "
   End If
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_CofNCo_Listad.Redraw = False
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         grd_CofNCo_Listad.Rows = grd_CofNCo_Listad.Rows + 1
         grd_CofNCo_Listad.Row = grd_CofNCo_Listad.Rows - 1
         r_dbl_ImpCuo = 0
         
         grd_CofNCo_Listad.Col = 0
         grd_CofNCo_Listad.Text = Format(g_rst_Princi!HIPCUO_NUMCUO, "000")
      
         grd_CofNCo_Listad.Col = 1
         grd_CofNCo_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECVCT))
         
         grd_CofNCo_Listad.Col = 2
         grd_CofNCo_Listad.Text = Format(g_rst_Princi!HIPCUO_CAPITA, "###,###,##0.00")
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(grd_CofNCo_Listad.Text)
         
         grd_CofNCo_Listad.Col = 3
         grd_CofNCo_Listad.Text = Format(g_rst_Princi!HIPCUO_INTERE, "###,###,##0.00")
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(grd_CofNCo_Listad.Text)
         
         grd_CofNCo_Listad.Col = 4
         grd_CofNCo_Listad.Text = Format(g_rst_Princi!HIPCUO_COMCOF, "###,###,##0.00")
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(grd_CofNCo_Listad.Text)
         
         grd_CofNCo_Listad.Col = 5
         grd_CofNCo_Listad.Text = Format(r_dbl_ImpCuo, "###,###,##0.00")
         
         grd_CofNCo_Listad.Col = 6
         grd_CofNCo_Listad.Text = Format(g_rst_Princi!HIPCUO_SALCAP, "###,###,##0.00")

         r_dbl_Capita = r_dbl_Capita + CDbl(Format(g_rst_Princi!HIPCUO_CAPITA, "###,###,##0.00"))
         r_dbl_Intere = r_dbl_Intere + CDbl(Format(g_rst_Princi!HIPCUO_INTERE, "###,###,##0.00"))
         r_dbl_Comisi = r_dbl_Comisi + CDbl(Format(g_rst_Princi!HIPCUO_COMCOF, "###,###,##0.00"))
         r_dbl_TotCuo = r_dbl_TotCuo + CDbl(Format(r_dbl_ImpCuo, "###,###,##0.00"))
         g_rst_Princi.MoveNext
      Loop
      
      grd_CofNCo_Listad.Redraw = True
      Call gs_UbiIniGrid(grd_CofNCo_Listad)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   pnl_CofNCo_Capita.Caption = Format(r_dbl_Capita, "###,###,##0.00") & " "
   pnl_CofNCo_Intere.Caption = Format(r_dbl_Intere, "###,###,##0.00") & " "
   pnl_CofNCo_Comisi.Caption = Format(r_dbl_Comisi, "###,###,##0.00") & " "
   pnl_CofNCo_TotCuo.Caption = Format(r_dbl_TotCuo, "###,###,##0.00") & " "
End Sub

Private Sub fs_Carga_Cro_Especial()
Dim r_dbl_Capita     As Double
Dim r_dbl_Intere     As Double
Dim r_dbl_Seguro     As Double
Dim r_dbl_ImpCuo     As Double
Dim r_dbl_TotCuo     As Double
   
   Call gs_LimpiaGrid(grd_Especial_Cli)
   r_dbl_Capita = 0
   r_dbl_Intere = 0
   r_dbl_Seguro = 0
   r_dbl_TotCuo = 0
   
   If l_int_FlgCro = 1 Then
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT HIPCUO_NUMCUO, HIPCUO_FECVCT, HIPCUO_CAPITA, HIPCUO_INTERE,  "
      g_str_Parame = g_str_Parame & "       HIPCUO_DESORG, HIPCUO_VIVORG, HIPCUO_OTRORG, HIPCUO_SALCAP, HIPCUO_COMCOF  "
      g_str_Parame = g_str_Parame & "  FROM CRE_HIPCUO "
      g_str_Parame = g_str_Parame & " WHERE HIPCUO_NUMOPE = '" & moddat_g_str_NumOpe & "' "
      g_str_Parame = g_str_Parame & "   AND HIPCUO_TIPCRO = 6 "
      g_str_Parame = g_str_Parame & " ORDER BY HIPCUO_NUMCUO "
   Else
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT CUOCIE_NUMCUO AS HIPCUO_NUMCUO, CUOCIE_FECVCT AS HIPCUO_FECVCT, CUOCIE_CAPITA AS HIPCUO_CAPITA, "
      g_str_Parame = g_str_Parame & "       CUOCIE_INTERE AS HIPCUO_INTERE, CUOCIE_DESORG AS HIPCUO_DESORG, CUOCIE_VIVORG AS HIPCUO_VIVORG, "
      g_str_Parame = g_str_Parame & "       CUOCIE_OTRORG AS HIPCUO_OTRORG, CUOCIE_SALCAP AS HIPCUO_SALCAP, CUOCIE_COMCOF AS HIPCUO_COMCOF "
      g_str_Parame = g_str_Parame & "  FROM CRE_CUOCIE "
      g_str_Parame = g_str_Parame & " WHERE CUOCIE_PERMES = " & l_str_PerMes & " "
      g_str_Parame = g_str_Parame & "   AND CUOCIE_PERANO = " & l_str_PerAno & " "
      g_str_Parame = g_str_Parame & "   AND CUOCIE_NUMOPE = '" & moddat_g_str_NumOpe & "' "
      g_str_Parame = g_str_Parame & "   AND CUOCIE_TIPCRO = 6 "
      g_str_Parame = g_str_Parame & " ORDER BY CUOCIE_NUMCUO "
   End If
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_Especial_Cli.Redraw = False
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         grd_Especial_Cli.Rows = grd_Especial_Cli.Rows + 1
         grd_Especial_Cli.Row = grd_Especial_Cli.Rows - 1
         r_dbl_ImpCuo = 0
         
         grd_Especial_Cli.Col = 0
         grd_Especial_Cli.Text = Format(g_rst_Princi!HIPCUO_NUMCUO, "000")
      
         grd_Especial_Cli.Col = 1
         grd_Especial_Cli.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECVCT))
         
         grd_Especial_Cli.Col = 2
         grd_Especial_Cli.Text = Format(g_rst_Princi!HIPCUO_CAPITA, "###,###,##0.00")
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(grd_Especial_Cli.Text)
         
         grd_Especial_Cli.Col = 3
         grd_Especial_Cli.Text = Format(g_rst_Princi!HIPCUO_INTERE, "###,###,##0.00")
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(grd_Especial_Cli.Text)
         
         grd_Especial_Cli.Col = 4
         grd_Especial_Cli.Text = Format(g_rst_Princi!HIPCUO_DESORG, "###,###,##0.00")
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(grd_Especial_Cli.Text)
         
         grd_Especial_Cli.Col = 5
         grd_Especial_Cli.Text = Format(r_dbl_ImpCuo, "###,###,##0.00")
         
         grd_Especial_Cli.Col = 6
         grd_Especial_Cli.Text = Format(g_rst_Princi!HIPCUO_SALCAP, "###,###,##0.00")

         r_dbl_Capita = r_dbl_Capita + CDbl(Format(g_rst_Princi!HIPCUO_CAPITA, "###,###,##0.00"))
         r_dbl_Intere = r_dbl_Intere + CDbl(Format(g_rst_Princi!HIPCUO_INTERE, "###,###,##0.00"))
         r_dbl_Seguro = r_dbl_Seguro + CDbl(Format(g_rst_Princi!HIPCUO_DESORG, "###,###,##0.00"))
         r_dbl_TotCuo = r_dbl_TotCuo + CDbl(Format(r_dbl_ImpCuo, "###,###,##0.00"))
         g_rst_Princi.MoveNext
      Loop
      
      grd_Especial_Cli.Redraw = True
      Call gs_UbiIniGrid(grd_Especial_Cli)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   pnl_Especial_Capital.Caption = Format(r_dbl_Capita, "###,###,##0.00") & " "
   pnl_Especial_Interes.Caption = Format(r_dbl_Intere, "###,###,##0.00") & " "
   pnl_Especial_Seguros.Caption = Format(r_dbl_Seguro, "###,###,##0.00") & " "
   pnl_Especial_TotalCuota.Caption = Format(r_dbl_TotCuo, "###,###,##0.00") & " "
End Sub


' ******* 22012020 INICIO

Private Sub fs_Exc_TramoPruebas()
Dim r_obj_Excel      As Excel.Application
Dim r_int_ConVer     As Integer
Dim r_str_EmpSeg     As String

'Dim a, b As String
'a = "0030900002 , 0040800082"
'b = "0040800082','0030900002,'0040800084"
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_HIPCUO C"
'     g_str_Parame = g_str_Parame & " WHERE HIPCUO_NUMOPE IN ( '" & b & "')  "
     g_str_Parame = g_str_Parame & "WHERE  C.HIPCUO_TIPCRO = 1 AND rownum <= 50 "
     MsgBox (g_str_Parame)
'   g_str_Parame = g_str_Parame & " WHERE HIPCUO_NUMOPE = '" & moddat_g_str_NumOpe & "' "
'   g_str_Parame = g_str_Parame & "   AND HIPCUO_TIPCRO IN 1 "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Cells(1, 1) = "ITEM"
      .Cells(1, 2) = "NRO. OPERACION"
      .Cells(1, 3) = "NRO. CUOTA"
      .Cells(1, 4) = "FEC. VEN."
      .Cells(1, 5) = "SITUACION"
      .Cells(1, 6) = "CAPITAL"
      .Cells(1, 7) = "INTERES"
      .Cells(1, 8) = "SEG. PREST."
      .Cells(1, 9) = "SEG. VIVIENDA"
      .Cells(1, 10) = "PORTES"
      .Cells(1, 11) = "MTO. CUOTA"
      .Cells(1, 12) = "SALDO CAPITAL"

      .Range(.Cells(1, 1), .Cells(1, 12)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 12)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(1, 1), .Cells(1, 12)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(1, 1), .Cells(1, 12)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(1, 1), .Cells(1, 12)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(1, 1), .Cells(1, 12)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(1, 1), .Cells(1, 12)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(1, 1), .Cells(1, 12)).Borders(xlInsideVertical).LineStyle = xlContinuous
       
      .Columns("A").ColumnWidth = 5
      .Columns("B").ColumnWidth = 16
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 12
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 10
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("D").NumberFormat = "@"
      .Columns("E").ColumnWidth = 12
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").ColumnWidth = 9
      '.Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("F").NumberFormat = "###,###,##0.00"
      .Columns("G").ColumnWidth = 8
      '.Columns("G").HorizontalAlignment = xlHAlignCenter
      .Columns("G").NumberFormat = "###,###,##0.00"
      .Columns("H").ColumnWidth = 11
      '.Columns("H").HorizontalAlignment = xlHAlignCenter
      .Columns("H").NumberFormat = "###,###,##0.00"
      .Columns("I").ColumnWidth = 14
      '.Columns("I").HorizontalAlignment = xlHAlignCenter
      .Columns("I").NumberFormat = "###,###,##0.00"
      .Columns("J").ColumnWidth = 7
      '.Columns("J").HorizontalAlignment = xlHAlignCenter
      .Columns("J").NumberFormat = "###,###,##0.00"
      .Columns("K").ColumnWidth = 12
      '.Columns("K").HorizontalAlignment = xlHAlignCenter
      .Columns("K").NumberFormat = "###,###,##0.00"
      .Columns("L").ColumnWidth = 14
      '.Columns("L").HorizontalAlignment = xlHAlignCenter
      .Columns("L").NumberFormat = "###,###,##0.00"
      .Columns("M").ColumnWidth = 10
   End With
   
   g_rst_Princi.MoveFirst
   r_int_ConVer = 2
   
   Do While Not g_rst_Princi.EOF
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 1
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = gf_Formato_NumOpe(g_rst_Princi!HIPCUO_NUMOPE)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = CStr(g_rst_Princi!HIPCUO_NUMCUO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = "" & gf_FormatoFecha(g_rst_Princi!HIPCUO_FECVCT)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = IIf(CStr(g_rst_Princi!HIPCUO_SITUAC) = 2, "POR VENCER", "PAGADA")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = gf_FormatoNumero(g_rst_Princi!HIPCUO_CAPITA, 12, 2)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = gf_FormatoNumero(g_rst_Princi!HIPCUO_INTERE, 12, 2)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = gf_FormatoNumero(g_rst_Princi!HIPCUO_DESORG, 12, 2)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = gf_FormatoNumero(g_rst_Princi!HIPCUO_VIVORG, 12, 2)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = gf_FormatoNumero(g_rst_Princi!HIPCUO_OTRORG, 14, 4)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = gf_FormatoNumero(g_rst_Princi!HIPCUO_CAPITA + g_rst_Princi!HIPCUO_INTERE + g_rst_Princi!HIPCUO_DESORG + g_rst_Princi!HIPCUO_VIVORG + g_rst_Princi!HIPCUO_OTRORG, 14, 4)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = gf_FormatoNumero(g_rst_Princi!HIPCUO_SALCAP, 14, 4)
              
      r_int_ConVer = r_int_ConVer + 1
      g_rst_Princi.MoveNext
      DoEvents
   Loop
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Screen.MousePointer = 0
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

' ******* 22012020 FIN


Private Sub fs_Exc_Tramo1()
Dim r_obj_Excel      As Excel.Application
Dim r_int_ConVer     As Integer
Dim r_str_EmpSeg     As String
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_HIPCUO "
   g_str_Parame = g_str_Parame & " WHERE HIPCUO_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_TIPCRO = 1 "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Cells(1, 1) = "ITEM"
      .Cells(1, 2) = "NRO. OPERACION"
      .Cells(1, 3) = "NRO. CUOTA"
      .Cells(1, 4) = "FEC. VEN."
      .Cells(1, 5) = "SITUACION"
      .Cells(1, 6) = "CAPITAL"
      .Cells(1, 7) = "INTERES"
      .Cells(1, 8) = "SEG. PREST."
      .Cells(1, 9) = "SEG. VIVIENDA"
      .Cells(1, 10) = "PORTES"
      .Cells(1, 11) = "MTO. CUOTA"
      .Cells(1, 12) = "SALDO CAPITAL"

      .Range(.Cells(1, 1), .Cells(1, 12)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 12)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(1, 1), .Cells(1, 12)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(1, 1), .Cells(1, 12)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(1, 1), .Cells(1, 12)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(1, 1), .Cells(1, 12)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(1, 1), .Cells(1, 12)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(1, 1), .Cells(1, 12)).Borders(xlInsideVertical).LineStyle = xlContinuous
       
      .Columns("A").ColumnWidth = 5
      .Columns("B").ColumnWidth = 16
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 12
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 10
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("D").NumberFormat = "@"
      .Columns("E").ColumnWidth = 12
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").ColumnWidth = 9
      '.Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("F").NumberFormat = "###,###,##0.00"
      .Columns("G").ColumnWidth = 8
      '.Columns("G").HorizontalAlignment = xlHAlignCenter
      .Columns("G").NumberFormat = "###,###,##0.00"
      .Columns("H").ColumnWidth = 11
      '.Columns("H").HorizontalAlignment = xlHAlignCenter
      .Columns("H").NumberFormat = "###,###,##0.00"
      .Columns("I").ColumnWidth = 14
      '.Columns("I").HorizontalAlignment = xlHAlignCenter
      .Columns("I").NumberFormat = "###,###,##0.00"
      .Columns("J").ColumnWidth = 7
      '.Columns("J").HorizontalAlignment = xlHAlignCenter
      .Columns("J").NumberFormat = "###,###,##0.00"
      .Columns("K").ColumnWidth = 12
      '.Columns("K").HorizontalAlignment = xlHAlignCenter
      .Columns("K").NumberFormat = "###,###,##0.00"
      .Columns("L").ColumnWidth = 14
      '.Columns("L").HorizontalAlignment = xlHAlignCenter
      .Columns("L").NumberFormat = "###,###,##0.00"
      .Columns("M").ColumnWidth = 10
   End With
   
   g_rst_Princi.MoveFirst
   r_int_ConVer = 2
   
   Do While Not g_rst_Princi.EOF
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 1
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = gf_Formato_NumOpe(g_rst_Princi!HIPCUO_NUMOPE)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = CStr(g_rst_Princi!HIPCUO_NUMCUO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = "" & gf_FormatoFecha(g_rst_Princi!HIPCUO_FECVCT)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = IIf(CStr(g_rst_Princi!HIPCUO_SITUAC) = 2, "POR VENCER", "PAGADA")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = gf_FormatoNumero(g_rst_Princi!HIPCUO_CAPITA, 12, 2)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = gf_FormatoNumero(g_rst_Princi!HIPCUO_INTERE, 12, 2)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = gf_FormatoNumero(g_rst_Princi!HIPCUO_DESORG, 12, 2)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = gf_FormatoNumero(g_rst_Princi!HIPCUO_VIVORG, 12, 2)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = gf_FormatoNumero(g_rst_Princi!HIPCUO_OTRORG, 14, 4)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = gf_FormatoNumero(g_rst_Princi!HIPCUO_CAPITA + g_rst_Princi!HIPCUO_INTERE + g_rst_Princi!HIPCUO_DESORG + g_rst_Princi!HIPCUO_VIVORG + g_rst_Princi!HIPCUO_OTRORG, 14, 4)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = gf_FormatoNumero(g_rst_Princi!HIPCUO_SALCAP, 14, 4)
              
      r_int_ConVer = r_int_ConVer + 1
      g_rst_Princi.MoveNext
      DoEvents
   Loop
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Screen.MousePointer = 0
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub fs_Exc_Tramo1_Nuevo()
Dim r_obj_Excel      As Excel.Application
Dim r_int_ConVer     As Integer
Dim r_int_Filas      As Integer
   
   Screen.MousePointer = 11
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Cells(1, 1) = "ITEM"
      .Cells(1, 2) = "NRO. OPERACION"
      .Cells(1, 3) = "NRO. CUOTA"
      .Cells(1, 4) = "FEC. VEN."
      .Cells(1, 5) = "SITUACION"
      .Cells(1, 6) = "CAPITAL"
      .Cells(1, 7) = "INTERES"
      .Cells(1, 8) = "SEG. PREST."
      .Cells(1, 9) = "SEG. VIVIENDA"
      .Cells(1, 10) = "PORTES"
      .Cells(1, 11) = "MTO. CUOTA"
      .Cells(1, 12) = "SALDO CAPITAL"
      
      .Range(.Cells(1, 1), .Cells(1, 12)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 12)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(1, 1), .Cells(1, 12)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(1, 1), .Cells(1, 12)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(1, 1), .Cells(1, 12)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(1, 1), .Cells(1, 12)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(1, 1), .Cells(1, 12)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(1, 1), .Cells(1, 12)).Borders(xlInsideVertical).LineStyle = xlContinuous
      
'      .Range(.Cells(1, 1), .Cells(1, 12)).HorizontalAlignment = xlHAlignCenter
      .Columns("A").ColumnWidth = 5
      .Columns("B").ColumnWidth = 16
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 12
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 10
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("D").NumberFormat = "@"
      .Columns("E").ColumnWidth = 12
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").ColumnWidth = 9
      .Columns("F").NumberFormat = "###,###,##0.00"
      .Columns("G").ColumnWidth = 8
      .Columns("G").NumberFormat = "###,###,##0.00"
      .Columns("H").ColumnWidth = 11
      .Columns("H").NumberFormat = "###,###,##0.00"
      .Columns("I").ColumnWidth = 14
      .Columns("I").NumberFormat = "###,###,##0.00"
      .Columns("J").ColumnWidth = 7
      .Columns("J").NumberFormat = "###,###,##0.00"
      .Columns("K").ColumnWidth = 12
      .Columns("K").NumberFormat = "###,###,##0.00"
      .Columns("L").ColumnWidth = 14
      .Columns("L").NumberFormat = "###,###,##0.00"
      .Columns("M").ColumnWidth = 10
   End With
      
   r_int_ConVer = 2
   r_int_Filas = 0

   Do While r_int_Filas < grd_CliNCo_Listad.Rows
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 1
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = pnl_NumOpe.Caption
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = r_int_Filas + 1
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = grd_CliNCo_Listad.TextMatrix(r_int_Filas, 1)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = grd_CliNCo_Listad.TextMatrix(r_int_Filas, 9)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = gf_FormatoNumero(grd_CliNCo_Listad.TextMatrix(r_int_Filas, 2), 12, 2)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = gf_FormatoNumero(grd_CliNCo_Listad.TextMatrix(r_int_Filas, 3), 12, 2)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = gf_FormatoNumero(grd_CliNCo_Listad.TextMatrix(r_int_Filas, 4), 12, 2)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = gf_FormatoNumero(grd_CliNCo_Listad.TextMatrix(r_int_Filas, 5), 12, 2)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = gf_FormatoNumero(grd_CliNCo_Listad.TextMatrix(r_int_Filas, 6), 14, 4)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = gf_FormatoNumero(CDbl(grd_CliNCo_Listad.TextMatrix(r_int_Filas, 2)) + CDbl(grd_CliNCo_Listad.TextMatrix(r_int_Filas, 3)) + CDbl(grd_CliNCo_Listad.TextMatrix(r_int_Filas, 4)) + CDbl(grd_CliNCo_Listad.TextMatrix(r_int_Filas, 5)) + CDbl(grd_CliNCo_Listad.TextMatrix(r_int_Filas, 6)), 14, 4)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = gf_FormatoNumero(grd_CliNCo_Listad.TextMatrix(r_int_Filas, 8), 14, 4)
              
      r_int_ConVer = r_int_ConVer + 1
      r_int_Filas = r_int_Filas + 1
      DoEvents
   Loop
     
   Screen.MousePointer = 0
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub fs_Exc_Tramo2()
Dim r_obj_Excel      As Excel.Application
Dim r_int_ConVer     As Integer
Dim r_str_EmpSeg     As String
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_HIPCUO "
   g_str_Parame = g_str_Parame & " WHERE HIPCUO_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_TIPCRO = 2 "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Cells(1, 1) = "ITEM"
      .Cells(1, 2) = "NRO. OPERACION"
      .Cells(1, 3) = "NRO. CUOTA"
      .Cells(1, 4) = "FEC. VEN."
      .Cells(1, 5) = "SITUACION"
      .Cells(1, 6) = "CAPITAL"
      .Cells(1, 7) = "INTERES"
      .Cells(1, 8) = "MTO. CUOTA"
      .Cells(1, 9) = "SALDO CAPITAL"
       
      .Range(.Cells(1, 1), .Cells(1, 9)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 9)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(1, 1), .Cells(1, 9)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(1, 1), .Cells(1, 9)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(1, 1), .Cells(1, 9)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(1, 1), .Cells(1, 9)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(1, 1), .Cells(1, 9)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(1, 1), .Cells(1, 9)).Borders(xlInsideVertical).LineStyle = xlContinuous
       
      .Columns("A").ColumnWidth = 5
      .Columns("B").ColumnWidth = 16
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 12
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 10
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("D").NumberFormat = "@"
      .Columns("E").ColumnWidth = 12
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").ColumnWidth = 9
      '.Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("F").NumberFormat = "###,###,##0.00"
      .Columns("G").ColumnWidth = 8
      '.Columns("G").HorizontalAlignment = xlHAlignCenter
      .Columns("G").NumberFormat = "###,###,##0.00"
      .Columns("H").ColumnWidth = 12
      '.Columns("H").HorizontalAlignment = xlHAlignCenter
      .Columns("H").NumberFormat = "###,###,##0.00"
      .Columns("I").ColumnWidth = 14
      '.Columns("I").HorizontalAlignment = xlHAlignCenter
      .Columns("I").NumberFormat = "###,###,##0.00"
   End With
   
   g_rst_Princi.MoveFirst
   r_int_ConVer = 2
   
   Do While Not g_rst_Princi.EOF
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 1
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = gf_Formato_NumOpe(g_rst_Princi!HIPCUO_NUMOPE)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = CStr(g_rst_Princi!HIPCUO_NUMCUO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = "" & gf_FormatoFecha(g_rst_Princi!HIPCUO_FECVCT)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = IIf(CStr(g_rst_Princi!HIPCUO_SITUAC) = 2, "POR VENCER", "PAGADA")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = gf_FormatoNumero(g_rst_Princi!HIPCUO_CAPITA, 12, 2)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = gf_FormatoNumero(g_rst_Princi!HIPCUO_INTERE, 12, 2)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = gf_FormatoNumero(g_rst_Princi!HIPCUO_CAPITA + g_rst_Princi!HIPCUO_INTERE, 14, 4)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = gf_FormatoNumero(g_rst_Princi!HIPCUO_SALCAP, 14, 4)
              
      r_int_ConVer = r_int_ConVer + 1
      g_rst_Princi.MoveNext
      DoEvents
   Loop
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   Screen.MousePointer = 0
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub fs_Exc_Tramo2_Nuevo()
Dim r_obj_Excel      As Excel.Application
Dim r_int_ConVer     As Integer
Dim r_int_Filas      As Integer
      
   Screen.MousePointer = 11
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Cells(1, 1) = "ITEM"
      .Cells(1, 2) = "NRO. OPERACION"
      .Cells(1, 3) = "NRO. CUOTA"
      .Cells(1, 4) = "FEC. VEN."
      .Cells(1, 5) = "SITUACION"
      .Cells(1, 6) = "CAPITAL"
      .Cells(1, 7) = "INTERES"
      .Cells(1, 8) = "MTO. CUOTA"
      .Cells(1, 9) = "SALDO CAPITAL"
       
      .Range(.Cells(1, 1), .Cells(1, 9)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 9)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(1, 1), .Cells(1, 9)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(1, 1), .Cells(1, 9)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(1, 1), .Cells(1, 9)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(1, 1), .Cells(1, 9)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(1, 1), .Cells(1, 9)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(1, 1), .Cells(1, 9)).Borders(xlInsideVertical).LineStyle = xlContinuous
       
      .Columns("A").ColumnWidth = 5
      .Columns("B").ColumnWidth = 16
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 12
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 10
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("D").NumberFormat = "@"
      .Columns("E").ColumnWidth = 12
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").ColumnWidth = 9
      .Columns("F").NumberFormat = "###,###,##0.00"
      .Columns("G").ColumnWidth = 8
      .Columns("G").NumberFormat = "###,###,##0.00"
      .Columns("H").ColumnWidth = 12
      .Columns("H").NumberFormat = "###,###,##0.00"
      .Columns("I").ColumnWidth = 14
      .Columns("I").NumberFormat = "###,###,##0.00"
   End With
   
   r_int_ConVer = 2
   r_int_Filas = 0
   
   Do While r_int_Filas < grd_CliCon_Listad.Rows
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 1
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = pnl_NumOpe.Caption
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = r_int_Filas + 1
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = grd_CliCon_Listad.TextMatrix(r_int_Filas, 1)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = grd_CliCon_Listad.TextMatrix(r_int_Filas, 6)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = gf_FormatoNumero(grd_CliCon_Listad.TextMatrix(r_int_Filas, 2), 12, 2)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = gf_FormatoNumero(grd_CliCon_Listad.TextMatrix(r_int_Filas, 3), 12, 2)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = gf_FormatoNumero(CDbl(grd_CliCon_Listad.TextMatrix(r_int_Filas, 2)) + CDbl(grd_CliCon_Listad.TextMatrix(r_int_Filas, 3)), 14, 4)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = gf_FormatoNumero(grd_CliCon_Listad.TextMatrix(r_int_Filas, 5), 14, 4)
              
      r_int_ConVer = r_int_ConVer + 1
      r_int_Filas = r_int_Filas + 1
      DoEvents
   Loop
   
   Screen.MousePointer = 0
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub fs_Exc_Tramo3()
Dim r_obj_Excel      As Excel.Application
Dim r_int_ConVer     As Integer
Dim r_str_EmpSeg     As String
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_HIPCUO "
   g_str_Parame = g_str_Parame & " WHERE HIPCUO_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_TIPCRO = 3 "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Cells(1, 1) = "ITEM"
      .Cells(1, 2) = "NRO. OPERACION"
      .Cells(1, 3) = "NRO. CUOTA"
      .Cells(1, 4) = "FEC. VEN."
      .Cells(1, 5) = "SITUACION"
      .Cells(1, 6) = "CAPITAL"
      .Cells(1, 7) = "INTERES"
      .Cells(1, 8) = "SEG. PREST."
      .Cells(1, 9) = "SEG. VIVIENDA"
      .Cells(1, 10) = "PORTES"
      .Cells(1, 11) = "COM. COFIDE"
      .Cells(1, 12) = "MTO. CUOTA"
      .Cells(1, 13) = "SALDO CAPITAL"
      .Range(.Cells(1, 1), .Cells(1, 13)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 13)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(1, 1), .Cells(1, 13)).Interior.Color = RGB(146, 208, 80)
      
      .Range(.Cells(1, 1), .Cells(1, 13)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(1, 1), .Cells(1, 13)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(1, 1), .Cells(1, 13)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(1, 1), .Cells(1, 13)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(1, 1), .Cells(1, 13)).Borders(xlInsideVertical).LineStyle = xlContinuous
       
      .Columns("A").ColumnWidth = 5
      .Columns("B").ColumnWidth = 16
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 12
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 10
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("D").NumberFormat = "@"
      .Columns("E").ColumnWidth = 12
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").ColumnWidth = 9
      '.Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("F").NumberFormat = "###,###,##0.00"
      .Columns("G").ColumnWidth = 8
      '.Columns("G").HorizontalAlignment = xlHAlignCenter
      .Columns("G").NumberFormat = "###,###,##0.00"
      .Columns("H").ColumnWidth = 11
      '.Columns("H").HorizontalAlignment = xlHAlignCenter
      .Columns("H").NumberFormat = "###,###,##0.00"
      .Columns("I").ColumnWidth = 14
      '.Columns("I").HorizontalAlignment = xlHAlignCenter
      .Columns("I").NumberFormat = "###,###,##0.00"
      .Columns("J").ColumnWidth = 7
      '.Columns("J").HorizontalAlignment = xlHAlignCenter
      .Columns("J").NumberFormat = "###,###,##0.00"
      .Columns("K").ColumnWidth = 12
      '.Columns("K").HorizontalAlignment = xlHAlignCenter
      .Columns("K").NumberFormat = "###,###,##0.00"
      .Columns("L").ColumnWidth = 14
      '.Columns("L").HorizontalAlignment = xlHAlignCenter
      .Columns("L").NumberFormat = "###,###,##0.00"
      .Columns("M").ColumnWidth = 14
      '.Columns("M").HorizontalAlignment = xlHAlignCenter
      .Columns("M").NumberFormat = "###,###,##0.00"
   End With
   
   g_rst_Princi.MoveFirst
   r_int_ConVer = 2
   
   Do While Not g_rst_Princi.EOF
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 1
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = gf_Formato_NumOpe(g_rst_Princi!HIPCUO_NUMOPE)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = CStr(g_rst_Princi!HIPCUO_NUMCUO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = "" & gf_FormatoFecha(g_rst_Princi!HIPCUO_FECVCT)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = IIf(CStr(g_rst_Princi!HIPCUO_SITUAC) = 2, "POR VENCER", "PAGADA")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = gf_FormatoNumero(g_rst_Princi!HIPCUO_CAPITA, 12, 2)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = gf_FormatoNumero(g_rst_Princi!HIPCUO_INTERE, 12, 2)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = gf_FormatoNumero(g_rst_Princi!HIPCUO_DESORG, 12, 2)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = gf_FormatoNumero(g_rst_Princi!HIPCUO_VIVORG, 12, 2)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = gf_FormatoNumero(g_rst_Princi!HIPCUO_OTRORG, 14, 4)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = gf_FormatoNumero(g_rst_Princi!HIPCUO_COMCOF, 14, 4)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = gf_FormatoNumero(g_rst_Princi!HIPCUO_CAPITA + g_rst_Princi!HIPCUO_INTERE + g_rst_Princi!HIPCUO_DESORG + g_rst_Princi!HIPCUO_VIVORG + g_rst_Princi!HIPCUO_OTRORG + g_rst_Princi!HIPCUO_COMCOF, 14, 4)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = gf_FormatoNumero(g_rst_Princi!HIPCUO_SALCAP, 14, 4)
              
      r_int_ConVer = r_int_ConVer + 1
      g_rst_Princi.MoveNext
      DoEvents
   Loop
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   Screen.MousePointer = 0
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub


Private Sub fs_Exc_Tramo3_Nuevo()
Dim r_obj_Excel      As Excel.Application
Dim r_int_ConVer     As Integer
Dim r_int_Filas      As Integer
  
   Screen.MousePointer = 11
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Cells(1, 1) = "ITEM"
      .Cells(1, 2) = "NRO. OPERACION"
      .Cells(1, 3) = "NRO. CUOTA"
      .Cells(1, 4) = "FEC. VEN."
      .Cells(1, 5) = "SITUACION"
      .Cells(1, 6) = "CAPITAL"
      .Cells(1, 7) = "INTERES"
      .Cells(1, 8) = "SEG. PREST."
      .Cells(1, 9) = "SEG. VIVIENDA"
      .Cells(1, 10) = "PORTES"
      .Cells(1, 11) = "COM. COFIDE"
      .Cells(1, 12) = "MTO. CUOTA"
      .Cells(1, 13) = "SALDO CAPITAL"
      .Range(.Cells(1, 1), .Cells(1, 13)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 13)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(1, 1), .Cells(1, 13)).Interior.Color = RGB(146, 208, 80)
      
      .Range(.Cells(1, 1), .Cells(1, 13)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(1, 1), .Cells(1, 13)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(1, 1), .Cells(1, 13)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(1, 1), .Cells(1, 13)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(1, 1), .Cells(1, 13)).Borders(xlInsideVertical).LineStyle = xlContinuous
       
      .Columns("A").ColumnWidth = 5
      .Columns("B").ColumnWidth = 16
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 12
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 10
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("D").NumberFormat = "@"
      .Columns("E").ColumnWidth = 12
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").ColumnWidth = 9
      .Columns("F").NumberFormat = "###,###,##0.00"
      .Columns("G").ColumnWidth = 8
      .Columns("G").NumberFormat = "###,###,##0.00"
      .Columns("H").ColumnWidth = 11
      .Columns("H").NumberFormat = "###,###,##0.00"
      .Columns("I").ColumnWidth = 14
      .Columns("I").NumberFormat = "###,###,##0.00"
      .Columns("J").ColumnWidth = 7
      .Columns("J").NumberFormat = "###,###,##0.00"
      .Columns("K").ColumnWidth = 12
      .Columns("K").NumberFormat = "###,###,##0.00"
      .Columns("L").ColumnWidth = 14
      .Columns("L").NumberFormat = "###,###,##0.00"
      .Columns("M").ColumnWidth = 14
      .Columns("M").NumberFormat = "###,###,##0.00"
   End With
   
   r_int_ConVer = 2
   r_int_Filas = 0
   
   Do While r_int_Filas < grd_MViNCo_Listad.Rows
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 1
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = pnl_NumOpe.Caption
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = r_int_Filas + 1
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = grd_MViNCo_Listad.TextMatrix(r_int_Filas, 1)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = grd_MViNCo_Listad.TextMatrix(r_int_Filas, 10)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = gf_FormatoNumero(grd_MViNCo_Listad.TextMatrix(r_int_Filas, 2), 12, 2)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = gf_FormatoNumero(grd_MViNCo_Listad.TextMatrix(r_int_Filas, 3), 12, 2)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = gf_FormatoNumero(grd_MViNCo_Listad.TextMatrix(r_int_Filas, 4), 12, 2)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = gf_FormatoNumero(grd_MViNCo_Listad.TextMatrix(r_int_Filas, 5), 12, 2)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = gf_FormatoNumero(grd_MViNCo_Listad.TextMatrix(r_int_Filas, 6), 14, 4)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = gf_FormatoNumero(grd_MViNCo_Listad.TextMatrix(r_int_Filas, 7), 14, 4)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = gf_FormatoNumero(CDbl(grd_MViNCo_Listad.TextMatrix(r_int_Filas, 2)) + CDbl(grd_MViNCo_Listad.TextMatrix(r_int_Filas, 3)) + CDbl(grd_MViNCo_Listad.TextMatrix(r_int_Filas, 4)) + CDbl(grd_MViNCo_Listad.TextMatrix(r_int_Filas, 5)) + CDbl(grd_MViNCo_Listad.TextMatrix(r_int_Filas, 6)) + CDbl(grd_MViNCo_Listad.TextMatrix(r_int_Filas, 7)), 14, 4)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = gf_FormatoNumero(grd_MViNCo_Listad.TextMatrix(r_int_Filas, 9), 14, 4)
              
      r_int_ConVer = r_int_ConVer + 1
      r_int_Filas = r_int_Filas + 1
      DoEvents
   Loop
      
   Screen.MousePointer = 0
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub


Private Sub fs_Exc_Tramo4()
Dim r_obj_Excel      As Excel.Application
Dim r_int_ConVer     As Integer
Dim r_str_EmpSeg     As String
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_HIPCUO "
   g_str_Parame = g_str_Parame & " WHERE HIPCUO_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_TIPCRO = 4 "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Cells(1, 1) = "ITEM"
      .Cells(1, 2) = "NRO. OPERACION"
      .Cells(1, 3) = "NRO. CUOTA"
      .Cells(1, 4) = "FEC. VEN."
      .Cells(1, 5) = "SITUACION"
      .Cells(1, 6) = "CAPITAL"
      .Cells(1, 7) = "INTERES"
      .Cells(1, 8) = "COM. COFIDE"
      .Cells(1, 9) = "MTO. CUOTA"
      .Cells(1, 10) = "SALDO CAPITAL"
       
      .Range(.Cells(1, 1), .Cells(1, 10)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 10)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(1, 1), .Cells(1, 10)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(1, 1), .Cells(1, 10)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(1, 1), .Cells(1, 10)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(1, 1), .Cells(1, 10)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(1, 1), .Cells(1, 10)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(1, 1), .Cells(1, 10)).Borders(xlInsideVertical).LineStyle = xlContinuous
       
      .Columns("A").ColumnWidth = 5
      .Columns("B").ColumnWidth = 16
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 12
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 10
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("D").NumberFormat = "@"
      .Columns("E").ColumnWidth = 12
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").ColumnWidth = 9
      '.Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("F").NumberFormat = "###,###,##0.00"
      .Columns("G").ColumnWidth = 8
      '.Columns("G").HorizontalAlignment = xlHAlignCenter
      .Columns("G").NumberFormat = "###,###,##0.00"
      .Columns("H").ColumnWidth = 12
      '.Columns("H").HorizontalAlignment = xlHAlignCenter
      .Columns("H").NumberFormat = "###,###,##0.00"
      .Columns("I").ColumnWidth = 14
      '.Columns("I").HorizontalAlignment = xlHAlignCenter
      .Columns("I").NumberFormat = "###,###,##0.00"
      .Columns("J").ColumnWidth = 14
      '.Columns("J").HorizontalAlignment = xlHAlignCenter
      .Columns("J").NumberFormat = "###,###,##0.00"
   End With
   
   g_rst_Princi.MoveFirst
   r_int_ConVer = 2
   
   Do While Not g_rst_Princi.EOF
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 1
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = gf_Formato_NumOpe(g_rst_Princi!HIPCUO_NUMOPE)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = CStr(g_rst_Princi!HIPCUO_NUMCUO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = "" & gf_FormatoFecha(g_rst_Princi!HIPCUO_FECVCT)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = IIf(CStr(g_rst_Princi!HIPCUO_SITUAC) = 2, "POR VENCER", "PAGADA")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = gf_FormatoNumero(g_rst_Princi!HIPCUO_CAPITA, 12, 2)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = gf_FormatoNumero(g_rst_Princi!HIPCUO_INTERE, 12, 2)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = gf_FormatoNumero(g_rst_Princi!HIPCUO_COMCOF, 12, 2)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = gf_FormatoNumero(g_rst_Princi!HIPCUO_CAPITA + g_rst_Princi!HIPCUO_INTERE + g_rst_Princi!HIPCUO_COMCOF, 14, 4)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = gf_FormatoNumero(g_rst_Princi!HIPCUO_SALCAP, 14, 4)
              
      r_int_ConVer = r_int_ConVer + 1
      g_rst_Princi.MoveNext
      DoEvents
   Loop
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   Screen.MousePointer = 0
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub


Private Sub fs_Exc_Tramo4_Nuevo()
Dim r_obj_Excel      As Excel.Application
Dim r_int_ConVer     As Integer
Dim r_int_Filas      As Integer

   Screen.MousePointer = 11
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Cells(1, 1) = "ITEM"
      .Cells(1, 2) = "NRO. OPERACION"
      .Cells(1, 3) = "NRO. CUOTA"
      .Cells(1, 4) = "FEC. VEN."
      .Cells(1, 5) = "SITUACION"
      .Cells(1, 6) = "CAPITAL"
      .Cells(1, 7) = "INTERES"
      .Cells(1, 8) = "COM. COFIDE"
      .Cells(1, 9) = "MTO. CUOTA"
      .Cells(1, 10) = "SALDO CAPITAL"
       
      .Range(.Cells(1, 1), .Cells(1, 10)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 10)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(1, 1), .Cells(1, 10)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(1, 1), .Cells(1, 10)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(1, 1), .Cells(1, 10)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(1, 1), .Cells(1, 10)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(1, 1), .Cells(1, 10)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(1, 1), .Cells(1, 10)).Borders(xlInsideVertical).LineStyle = xlContinuous
       
      .Columns("A").ColumnWidth = 5
      .Columns("B").ColumnWidth = 16
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 12
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 10
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("D").NumberFormat = "@"
      .Columns("E").ColumnWidth = 12
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").ColumnWidth = 9
      .Columns("F").NumberFormat = "###,###,##0.00"
      .Columns("G").ColumnWidth = 8
      .Columns("G").NumberFormat = "###,###,##0.00"
      .Columns("H").ColumnWidth = 12
      .Columns("H").NumberFormat = "###,###,##0.00"
      .Columns("I").ColumnWidth = 14
      .Columns("I").NumberFormat = "###,###,##0.00"
      .Columns("J").ColumnWidth = 14
      .Columns("J").NumberFormat = "###,###,##0.00"
   End With
   
   r_int_ConVer = 2
   r_int_Filas = 0
   
   Do While r_int_Filas < grd_MviCon_Listad.Rows
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 1
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = pnl_NumOpe.Caption
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = r_int_Filas + 1
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = grd_MviCon_Listad.TextMatrix(r_int_Filas, 1)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = grd_MviCon_Listad.TextMatrix(r_int_Filas, 7)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = gf_FormatoNumero(grd_MviCon_Listad.TextMatrix(r_int_Filas, 2), 12, 2)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = gf_FormatoNumero(grd_MviCon_Listad.TextMatrix(r_int_Filas, 3), 12, 2)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = gf_FormatoNumero(grd_MviCon_Listad.TextMatrix(r_int_Filas, 4), 12, 2)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = gf_FormatoNumero(CDbl(grd_MviCon_Listad.TextMatrix(r_int_Filas, 2)) + CDbl(grd_MviCon_Listad.TextMatrix(r_int_Filas, 3)) + CDbl(grd_MviCon_Listad.TextMatrix(r_int_Filas, 4)), 14, 4)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = gf_FormatoNumero(grd_MviCon_Listad.TextMatrix(r_int_Filas, 6), 14, 4)
              
      r_int_ConVer = r_int_ConVer + 1
      r_int_Filas = r_int_Filas + 1
      DoEvents
   Loop
   
   Screen.MousePointer = 0
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub fs_Exc_Tramo5()
Dim r_obj_Excel      As Excel.Application
Dim r_int_ConVer     As Integer
Dim r_str_EmpSeg     As String
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_HIPCUO "
   g_str_Parame = g_str_Parame & " WHERE HIPCUO_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_TIPCRO = 5 "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Cells(1, 1) = "ITEM"
      .Cells(1, 2) = "NRO. OPERACION"
      .Cells(1, 3) = "NRO. CUOTA"
      .Cells(1, 4) = "FEC. VEN."
      .Cells(1, 5) = "SITUACION"
      .Cells(1, 6) = "CAPITAL"
      .Cells(1, 7) = "INTERES"
      .Cells(1, 8) = "COM. COFIDE"
      .Cells(1, 9) = "MTO. CUOTA"
      .Cells(1, 10) = "SALDO CAPITAL"
       
      .Range(.Cells(1, 1), .Cells(1, 10)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 10)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(1, 1), .Cells(1, 10)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(1, 1), .Cells(1, 10)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(1, 1), .Cells(1, 10)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(1, 1), .Cells(1, 10)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(1, 1), .Cells(1, 10)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(1, 1), .Cells(1, 10)).Borders(xlInsideVertical).LineStyle = xlContinuous
       
      .Columns("A").ColumnWidth = 5
      .Columns("B").ColumnWidth = 16
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 12
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 10
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("D").NumberFormat = "@"
      .Columns("E").ColumnWidth = 12
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").ColumnWidth = 9
      '.Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("F").NumberFormat = "###,###,##0.00"
      .Columns("G").ColumnWidth = 8
      '.Columns("G").HorizontalAlignment = xlHAlignCenter
      .Columns("G").NumberFormat = "###,###,##0.00"
      .Columns("H").ColumnWidth = 12
      '.Columns("H").HorizontalAlignment = xlHAlignCenter
      .Columns("H").NumberFormat = "###,###,##0.00"
      .Columns("I").ColumnWidth = 14
      '.Columns("I").HorizontalAlignment = xlHAlignCenter
      .Columns("I").NumberFormat = "###,###,##0.00"
      .Columns("J").ColumnWidth = 14
      '.Columns("J").HorizontalAlignment = xlHAlignCenter
      .Columns("J").NumberFormat = "###,###,##0.00"
   End With
   
   g_rst_Princi.MoveFirst
   r_int_ConVer = 2
   
   Do While Not g_rst_Princi.EOF
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 1
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = gf_Formato_NumOpe(g_rst_Princi!HIPCUO_NUMOPE)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = CStr(g_rst_Princi!HIPCUO_NUMCUO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = "" & gf_FormatoFecha(g_rst_Princi!HIPCUO_FECVCT)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = IIf(CStr(g_rst_Princi!HIPCUO_SITUAC) = 2, "POR VENCER", "PAGADA")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = gf_FormatoNumero(g_rst_Princi!HIPCUO_CAPITA, 12, 2)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = gf_FormatoNumero(g_rst_Princi!HIPCUO_INTERE, 12, 2)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = gf_FormatoNumero(g_rst_Princi!HIPCUO_COMCOF, 12, 2)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = gf_FormatoNumero(g_rst_Princi!HIPCUO_CAPITA + g_rst_Princi!HIPCUO_INTERE + g_rst_Princi!HIPCUO_COMCOF, 14, 4)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = gf_FormatoNumero(g_rst_Princi!HIPCUO_SALCAP, 14, 4)
      
      r_int_ConVer = r_int_ConVer + 1
      g_rst_Princi.MoveNext
      DoEvents
   Loop
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   Screen.MousePointer = 0
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub tab_Cronog_DblClick()
   cmd_Cronog_Click
End Sub

Private Sub grd_CliCon_Listad_DblClick()
   If cmd_Cronog.Visible = True Then cmd_Cronog_Click
End Sub

Private Sub grd_CliCon_Listad_SelChange()
   If grd_CliCon_Listad.Rows > 2 Then
      grd_CliCon_Listad.RowSel = grd_CliCon_Listad.Row
   End If
End Sub

Private Sub grd_CliNCo_Listad_DblClick()
   If cmd_Cronog.Visible = True Then cmd_Cronog_Click
End Sub

Private Sub grd_CliNCo_Listad_SelChange()
   If grd_CliNCo_Listad.Rows > 2 Then
      grd_CliNCo_Listad.RowSel = grd_CliNCo_Listad.Row
   End If
End Sub

Private Sub grd_CofNCo_Listad_DblClick()
   If cmd_Cronog.Visible = True Then cmd_Cronog_Click
End Sub

Private Sub grd_CofNCo_Listad_SelChange()
   If grd_CofNCo_Listad.Rows > 2 Then
      grd_CofNCo_Listad.RowSel = grd_CofNCo_Listad.Row
   End If
End Sub

Private Sub grd_Especial_Cli_DblClick()
   If cmd_Cronog.Visible = True Then cmd_Cronog_Click
End Sub

Private Sub grd_Especial_Cli_SelChange()
   If grd_Especial_Cli.Rows > 2 Then
      grd_Especial_Cli.RowSel = grd_Especial_Cli.Row
   End If
End Sub

Private Sub grd_MviCon_Listad_DblClick()
   If cmd_Cronog.Visible = True Then cmd_Cronog_Click
End Sub

Private Sub grd_MViNCo_Listad_DblClick()
   If cmd_Cronog.Visible = True Then cmd_Cronog_Click
End Sub

Private Sub grd_MViNCo_Listad_SelChange()
   If grd_MViNCo_Listad.Rows > 2 Then
      grd_MViNCo_Listad.RowSel = grd_MViNCo_Listad.Row
   End If
End Sub

Private Sub grd_MViCon_Listad_SelChange()
   If grd_MviCon_Listad.Rows > 2 Then
      grd_MviCon_Listad.RowSel = grd_MviCon_Listad.Row
   End If
End Sub


'**** RAT 22012010 INICIO

Private Function fs_GenExcNuevo(var As String) As String
Dim r_rst_Princi      As ADODB.Recordset
Dim r_rst_Prindet      As ADODB.Recordset
Dim r_obj_Excel       As Excel.Application
Dim r_int_NumFil      As Integer
Dim r_str_Parame      As String

      r_str_Parame = ""
      r_str_Parame = r_str_Parame & "SELECT CRE_HIPCUO.HIPCUO_NUMOPE, CRE_HIPCUO.HIPCUO_NUMCUO, HIPCUO_FECVCT, CRE_HIPCUO.HIPCUO_CAPITA, CRE_HIPCUO.HIPCUO_INTERE,"
      r_str_Parame = r_str_Parame & "CRE_HIPCUO.HIPCUO_DESORG, CRE_HIPCUO.HIPCUO_VIVORG, CRE_HIPCUO.HIPCUO_OTRORG, CRE_HIPCUO.HIPCUO_SALCAP, CRE_HIPMAE.HIPMAE_TDOCLI,"
      r_str_Parame = r_str_Parame & "CAST(CRE_HIPMAE.HIPMAE_NDOCLI AS VARCHAR(30)) AS HIPMAE_NDOCLI , CRE_HIPMAE.HIPMAE_PLAANO,"
      r_str_Parame = r_str_Parame & "CRE_HIPMAE.HIPMAE_CUOANO,CRE_HIPMAE.HIPMAE_PERGRA, CRE_HIPMAE.HIPMAE_NUMCUO, CRE_HIPMAE.HIPMAE_FECDES, CRE_HIPMAE.HIPMAE_MONEDA,"
      r_str_Parame = r_str_Parame & "CRE_HIPMAE.HIPMAE_MTOPRE, CRE_HIPMAE.HIPMAE_INTCAP, CRE_HIPMAE.HIPMAE_TASINT, CRE_HIPMAE.HIPMAE_COSEFE, CLI_DATGEN.DATGEN_APEPAT,"
      r_str_Parame = r_str_Parame & "CLI_DATGEN.DATGEN_APEMAT , CLI_DATGEN.DATGEN_APECAS, CLI_DATGEN.DATGEN_NOMBRE, CRE_PRODUC.PRODUC_DESCRI"
      r_str_Parame = r_str_Parame & " From "
      r_str_Parame = r_str_Parame & "CRE_HIPCUO CRE_HIPCUO,"
      r_str_Parame = r_str_Parame & "CRE_HIPMAE CRE_HIPMAE,"
      r_str_Parame = r_str_Parame & "CLI_DATGEN CLI_DATGEN,"
      r_str_Parame = r_str_Parame & "CRE_PRODUC CRE_PRODUC"
      r_str_Parame = r_str_Parame & " Where "
      r_str_Parame = r_str_Parame & "CRE_HIPCUO.HIPCUO_NUMOPE = CRE_HIPMAE.HIPMAE_NUMOPE AND "
      r_str_Parame = r_str_Parame & "CRE_HIPMAE.HIPMAE_TDOCLI = CLI_DATGEN.DATGEN_TIPDOC AND "
      r_str_Parame = r_str_Parame & "CRE_HIPMAE.HIPMAE_NDOCLI = CLI_DATGEN.DATGEN_NUMDOC AND "
      r_str_Parame = r_str_Parame & "CRE_HIPMAE.HIPMAE_CODPRD = CRE_PRODUC.Produc_Codigo AND "
        r_str_Parame = r_str_Parame & "CRE_HIPCUO.HIPCUO_TIPCRO = 1 AND "
      r_str_Parame = r_str_Parame & "CRE_HIPCUO.HIPCUO_NUMOPE = '" & var & "'"
   
   
'   HIPCUO_TIPCRO
   
'MsgBox (r_str_Parame)

   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Princi, 3) Then
      Screen.MousePointer = 0
      MsgBox "Error al ejecutar el Procedimiento.", vbCritical, modgen_g_str_NomPlt
      Exit Function
   End If
   
   If r_rst_Princi.BOF And r_rst_Princi.EOF Then
      r_rst_Princi.Close
      Set r_rst_Princi = Nothing
      Screen.MousePointer = 0
      Exit Function
   End If
   
   r_rst_Princi.MoveFirst

   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
   
      'MARGENES
'      .PageSetup.LeftMargin = Application.CentimetersToPoints(1.5)
'      .PageSetup.RightMargin = Application.CentimetersToPoints(0.4)
'      .PageSetup.TopMargin = Application.CentimetersToPoints(1)
'      .PageSetup.BottomMargin = Application.CentimetersToPoints(1)
      
      .Columns("A").ColumnWidth = 10
      .Columns("B").ColumnWidth = 10
       .Columns("C").ColumnWidth = 8
      .Columns("D").ColumnWidth = 8
      .Columns("E").ColumnWidth = 8
      .Columns("F").ColumnWidth = 10
      .Columns("G").ColumnWidth = 10
      .Columns("H").ColumnWidth = 8
      
      .Columns("I").ColumnWidth = 8
      .Range(.Cells(1, 1), .Cells(600, 12)).Font.Name = "Arial (Western)"
      .Range(.Cells(1, 1), .Cells(600, 12)).Font.Size = 8
      .Range(.Cells(1, 1), .Cells(600, 12)).RowHeight = 14
      
      .Pictures.Insert(g_str_RutLog & "\" & "image001.gif").Select
      
      .Cells(7, 1) = "CRONOGRAMA DE PAGOS"
      .Range(.Cells(7, 1), .Cells(7, 9)).Merge
      .Range(.Cells(7, 1), .Cells(7, 9)).Font.Bold = True
      .Range(.Cells(7, 1), .Cells(7, 9)).Font.Underline = True
      .Range(.Cells(7, 1), .Cells(7, 9)).HorizontalAlignment = xlHAlignCenter

      .Rows(1).RowHeight = 1
      .Rows(8).RowHeight = 9
      .Rows(9).RowHeight = 5
      .Range(.Cells(9, 1), .Cells(9, 9)).Borders(xlEdgeTop).LineStyle = xlContinuous
      '.CELDA(FILA, COLUMNA)
      .Cells(2, 7) = "Nombre Reporte:"
      .Cells(3, 7) = "Fecha Emisión:"
      .Cells(4, 7) = "Hora Emisión:"
      .Cells(5, 7) = "Página:"
      
      .Cells(2, 9) = "PAGOS"
          .Cells(2, 9).HorizontalAlignment = xlHAlignCenter
      .Cells(3, 9) = Format(date, "dd/mm/yyyy")
      .Cells(3, 9).HorizontalAlignment = xlHAlignCenter
      .Cells(4, 9) = Format(Time, "hh:mm:ss")
      .Cells(4, 9).HorizontalAlignment = xlHAlignCenter
      .Cells(5, 9) = "1"
      .Cells(5, 9).HorizontalAlignment = xlHAlignCenter
'      .Range(.Cells(2, 9), .Cells(5, 9)).HorizontalAlignment = xlHAlignRight
      
'      .Cells(10, 1) = "Nro Operación:"
'      .Cells(10, 2) = r_rst_Princi!HIPCUO_NUMOPE
'      .Range(.Cells(10, 1), .Cells(10, 10)).Font.Bold = True



       .Range("A10:B10").Merge
      .Range("A10") = "Nro Operación:"
       .Range("C10:E10").Merge
       .Range("C10") = r_rst_Princi!HIPCUO_NUMOPE
      
       .Range("A11:B11").Merge
      .Range("A11") = "Cliente:"
       .Range("C11:E11").Merge
       .Range("C11") = Trim(r_rst_Princi!DATGEN_APEPAT) & " " & Trim(r_rst_Princi!DATGEN_APEMAT) & " " & Trim(r_rst_Princi!DATGEN_NOMBRE)
      
       .Range("A12:B12").Merge
      .Range("A12") = "Producto:"
       .Range("C12:F12").Merge
       .Range("C12") = Trim(r_rst_Princi!PRODUC_DESCRI)
      
       .Range("A13:B13").Merge
      .Range("A13") = "Moneda:"
       .Range("C13:E13").Merge
       .Range("C13") = IIf(r_rst_Princi!HIPMAE_MONEDA = 1, "SOLES", "DOLARES AMERICANOS")
       
       
      .Range("A14:B14").Merge
      .Range("A14") = "Plazo:"
       .Range("C14:E14").Merge
       .Range("C14") = Trim(r_rst_Princi!HIPMAE_PLAANO & " Años")
       
       
       .Range("A15:B15").Merge
      .Range("A15") = "Período Gracia:"
       .Range("C15:E15").Merge
       .Range("C15") = Trim(r_rst_Princi!HIPMAE_PERGRA & " Mes(es)")
       
        Dim a As String
             If r_rst_Princi!HIPMAE_CUOANO = 1 Then
               a = "NO"
             ElseIf r_rst_Princi!HIPMAE_CUOANO = 2 Then
               a = "JULIO"
             ElseIf r_rst_Princi!HIPMAE_CUOANO = 3 Then
               a = "DICIEMBRE"
             ElseIf r_rst_Princi!HIPMAE_CUOANO = 4 Then
               a = "JULIO Y DICIEMBRE"
             Else
               a = ""
             End If
       
       .Range("A16:B16").Merge
      .Range("A16") = "Cuotas Extraord:"
       .Range("C16:E16").Merge
       .Range("C16") = a
       
         .Range("A17:B17").Merge
      .Range("A17") = "Fecha Desembolso:"
       .Range("C17:E17").Merge
       .Range("C17") = gf_FormatoFecha(r_rst_Princi!HIPMAE_FECDES)
       
         .Range("A18:B18").Merge
      .Range("A18") = "Tasa de Interés:"
       .Range("C18:E18").Merge
       .Range("C18") = ((r_rst_Princi!HIPMAE_TASINT * 100) & " %")
       .Range("C18").NumberFormat = "###,###,##0.00"
        .Range("C18").HorizontalAlignment = xlHAlignLeft
       

      .Range("F14:G14").Merge
      .Range("F14") = "Código Cliente:"
       .Range("H14:I14").Merge
       .Range("H14") = "'" & r_rst_Princi!HIPMAE_NDOCLI
       
       
      .Range("H14").Font.Bold = True
      .Range("H14").Font.Underline = True
'       .Range("H15").NumberFormat = "@"
       
        .Range("F15:G15").Merge
       .Range("F15") = "Monto Préstamo:"
       .Range("H15:I15").Merge
       
'        "'" & Format(r_rst_Prindet!HIPCUO_CAPITA, "###,###,##0.00")
        .Range("H15") = "S/. " & Format(Trim(r_rst_Princi!HIPMAE_MTOPRE), "###,###,##0.00")
'        .Range("H15").HorizontalAlignment = xlHAlignCenter
'       .Range("H15") = "S/. " & Trim(r_rst_Princi!HIPMAE_MTOPRE)
'       .Range("H15").NumberFormat = "###,###,##0.00"
       
          .Range("F16:G16").Merge
          .Range("F16") = "Nro Cuotas:"
          .Range("H16:I16").Merge
          .Range("H16") = Trim(r_rst_Princi!HIPMAE_NUMCUO)
          
          .Range("H16").HorizontalAlignment = xlHAlignLeft
          
          
           .Range("F17:G17").Merge
          .Range("F17") = "Intereses Capitalizados:"
          .Range("H17:I17").Merge
          .Range("H17") = Trim(r_rst_Princi!HIPMAE_INTCAP)
          .Range("H17").NumberFormat = "###,###,##0.00"
           .Range("H17").HorizontalAlignment = xlHAlignLeft
          
          .Range("F18:G18").Merge
          .Range("F18") = "Monto Préstamo TNC:"
          .Range("H18:I18").Merge
          .Range("H18") = "S/. " & Trim(r_rst_Princi!HIPMAE_COSEFE)
       
          .Range(.Cells(20, 1), .Cells(20, 9)).Borders(xlEdgeTop).LineStyle = xlContinuous
      '.CELDA(FILA, COLUMNA)
      
          .Cells(21, 1) = "Cuota"
          .Cells(21, 2) = "F.Vcto"
          .Cells(21, 3) = "Capital"
          .Cells(21, 4) = "Interés"
          .Cells(21, 5) = "S.Desg."
          .Cells(21, 6) = "S.Inm."
          .Cells(21, 7) = "Portes"
          .Cells(21, 8) = "T.Cuota"
          .Cells(21, 9) = "S.Capital"
      
          .Range(.Cells(21, 1), .Cells(21, 9)).HorizontalAlignment = xlHAlignCenter
      
      r_str_Parame = ""
      r_str_Parame = r_str_Parame & "SELECT CRE_HIPCUO.HIPCUO_NUMOPE, CRE_HIPCUO.HIPCUO_NUMCUO, HIPCUO_FECVCT, CRE_HIPCUO.HIPCUO_CAPITA, CRE_HIPCUO.HIPCUO_INTERE,"
      r_str_Parame = r_str_Parame & "CRE_HIPCUO.HIPCUO_DESORG, CRE_HIPCUO.HIPCUO_VIVORG, CRE_HIPCUO.HIPCUO_OTRORG, CRE_HIPCUO.HIPCUO_SALCAP, CRE_HIPMAE.HIPMAE_TDOCLI,"
      r_str_Parame = r_str_Parame & "CAST(CRE_HIPMAE.HIPMAE_NDOCLI AS VARCHAR(30)) AS HIPMAE_NDOCLI , CRE_HIPMAE.HIPMAE_PLAANO,"
      r_str_Parame = r_str_Parame & "CRE_HIPMAE.HIPMAE_CUOANO,CRE_HIPMAE.HIPMAE_PERGRA, CRE_HIPMAE.HIPMAE_NUMCUO, CRE_HIPMAE.HIPMAE_FECDES, CRE_HIPMAE.HIPMAE_MONEDA,"
      r_str_Parame = r_str_Parame & "CRE_HIPMAE.HIPMAE_MTOPRE, CRE_HIPMAE.HIPMAE_INTCAP, CRE_HIPMAE.HIPMAE_TASINT, CRE_HIPMAE.HIPMAE_COSEFE, CLI_DATGEN.DATGEN_APEPAT,"
      r_str_Parame = r_str_Parame & "CLI_DATGEN.DATGEN_APEMAT , CLI_DATGEN.DATGEN_APECAS, CLI_DATGEN.DATGEN_NOMBRE, CRE_PRODUC.PRODUC_DESCRI"
      r_str_Parame = r_str_Parame & " From "
      r_str_Parame = r_str_Parame & "CRE_HIPCUO CRE_HIPCUO,"
      r_str_Parame = r_str_Parame & "CRE_HIPMAE CRE_HIPMAE,"
      r_str_Parame = r_str_Parame & "CLI_DATGEN CLI_DATGEN,"
      r_str_Parame = r_str_Parame & "CRE_PRODUC CRE_PRODUC"
      r_str_Parame = r_str_Parame & " Where "
      r_str_Parame = r_str_Parame & "CRE_HIPCUO.HIPCUO_NUMOPE = CRE_HIPMAE.HIPMAE_NUMOPE AND "
      r_str_Parame = r_str_Parame & "CRE_HIPMAE.HIPMAE_TDOCLI = CLI_DATGEN.DATGEN_TIPDOC AND "
      r_str_Parame = r_str_Parame & "CRE_HIPMAE.HIPMAE_NDOCLI = CLI_DATGEN.DATGEN_NUMDOC AND "
      r_str_Parame = r_str_Parame & "CRE_HIPMAE.HIPMAE_CODPRD = CRE_PRODUC.Produc_Codigo AND "
      r_str_Parame = r_str_Parame & "CRE_HIPCUO.HIPCUO_TIPCRO = 1 AND "
      r_str_Parame = r_str_Parame & "CRE_HIPCUO.HIPCUO_NUMOPE = '" & r_rst_Princi!HIPCUO_NUMOPE & "'"

'
'      MsgBox (r_str_Parame)
      
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Prindet, 3) Then
      Screen.MousePointer = 0
      MsgBox "Error al ejecutar el Procedimiento.", vbCritical, modgen_g_str_NomPlt
      Exit Function
   End If
   
   If r_rst_Prindet.BOF And r_rst_Prindet.EOF Then
      r_rst_Prindet.Close
      Set r_rst_Prindet = Nothing
      Screen.MousePointer = 0
      Exit Function
   End If
      
      Dim r_int_NroFil, r_int_corre As Integer
      Dim sum1, sum2, sum3, sum4, sum5, sum6 As Double
       r_int_NroFil = 22
       r_int_corre = 1
      sum1 = 0
      sum2 = 0
      sum3 = 0
      sum4 = 0
      sum5 = 0
      sum6 = 0
      
    
      r_rst_Prindet.MoveFirst
      
     Do While Not r_rst_Prindet.EOF


             .Cells(r_int_NroFil, 1) = r_rst_Prindet!HIPCUO_NUMCUO
             .Cells(r_int_NroFil, 1).HorizontalAlignment = xlHAlignLeft
             .Cells(r_int_NroFil, 2) = "" & gf_FormatoFecha(r_rst_Prindet!HIPCUO_FECVCT)
'            .Cells(r_int_NroFil, 3) = r_rst_Prindet!HIPCUO_CAPITA
'            .Cells(r_int_NroFil, 3).NumberFormat = "###,###,##0.00"
             .Cells(r_int_NroFil, 2).HorizontalAlignment = xlHAlignLeft
             .Cells(r_int_NroFil, 3) = "'" & Format(r_rst_Prindet!HIPCUO_CAPITA, "###,###,##0.00")
             .Cells(r_int_NroFil, 3).HorizontalAlignment = xlHAlignLeft
'            .Cells(r_int_NroFil, 4) = r_rst_Prindet!HIPCUO_INTERE
'            .Cells(r_int_NroFil, 4).NumberFormat = "###,###,##0.00"
             .Cells(r_int_NroFil, 4) = "'" & Format(r_rst_Prindet!HIPCUO_INTERE, "###,###,##0.00")
             .Cells(r_int_NroFil, 4).HorizontalAlignment = xlHAlignLeft
            .Cells(r_int_NroFil, 5) = r_rst_Prindet!HIPCUO_DESORG
            .Cells(r_int_NroFil, 5).NumberFormat = "###,###,##0.00"
            .Cells(r_int_NroFil, 5).HorizontalAlignment = xlHAlignLeft
            .Cells(r_int_NroFil, 6) = r_rst_Prindet!HIPCUO_VIVORG
            .Cells(r_int_NroFil, 6).NumberFormat = "###,###,##0.00"
            .Cells(r_int_NroFil, 6).HorizontalAlignment = xlHAlignLeft
            .Cells(r_int_NroFil, 7) = r_rst_Prindet!HIPCUO_OTRORG
            .Cells(r_int_NroFil, 7).NumberFormat = "###,###,##0.00"
            .Cells(r_int_NroFil, 7).HorizontalAlignment = xlHAlignLeft
            .Cells(r_int_NroFil, 8) = r_rst_Prindet!HIPCUO_CAPITA
            .Cells(r_int_NroFil, 8).NumberFormat = "###,###,##0.00"
            .Cells(r_int_NroFil, 8).HorizontalAlignment = xlHAlignLeft
            .Cells(r_int_NroFil, 9) = r_rst_Prindet!HIPCUO_SALCAP
            .Cells(r_int_NroFil, 9).NumberFormat = "###,###,##0.00"
            .Cells(r_int_NroFil, 9).HorizontalAlignment = xlHAlignLeft
            .Range(.Cells(r_int_NroFil, 1), .Cells(r_int_NroFil, 9)).HorizontalAlignment = xlHAlignCenter
            
             sum1 = sum1 + CDbl(.Cells(r_int_NroFil, 3))
             sum2 = sum2 + CDbl(.Cells(r_int_NroFil, 4))
             sum3 = sum3 + CDbl(.Cells(r_int_NroFil, 5))
             sum4 = sum4 + CDbl(.Cells(r_int_NroFil, 6))
             sum5 = sum5 + CDbl(.Cells(r_int_NroFil, 7))
             sum6 = sum6 + CDbl(.Cells(r_int_NroFil, 8))
                  
       r_int_corre = r_int_corre + 1
       r_int_NroFil = r_int_NroFil + 1
       

       r_rst_Prindet.MoveNext
       DoEvents
       

      Loop
      
      .Cells(r_int_NroFil, 2) = "TOTALES"
      .Cells(r_int_NroFil, 2).HorizontalAlignment = xlHAlignLeft
      
      
      .Cells(r_int_NroFil, 3) = sum1
      .Cells(r_int_NroFil, 3).NumberFormat = "###,###,##0.00"
      .Cells(r_int_NroFil, 3).HorizontalAlignment = xlHAlignLeft
      
      
      .Cells(r_int_NroFil, 4) = sum2
      .Cells(r_int_NroFil, 4).NumberFormat = "###,###,##0.00"
      .Cells(r_int_NroFil, 4).HorizontalAlignment = xlHAlignLeft
      
      
      .Cells(r_int_NroFil, 5) = sum3
      .Cells(r_int_NroFil, 5).NumberFormat = "###,###,##0.00"
      .Cells(r_int_NroFil, 5).HorizontalAlignment = xlHAlignLeft
      
      
      .Cells(r_int_NroFil, 6) = sum4
      .Cells(r_int_NroFil, 6).NumberFormat = "###,###,##0.00"
      .Cells(r_int_NroFil, 6).HorizontalAlignment = xlHAlignCenter
      
      
      .Cells(r_int_NroFil, 7) = sum5
      .Cells(r_int_NroFil, 7).NumberFormat = "###,###,##0.00"
      .Cells(r_int_NroFil, 7).HorizontalAlignment = xlHAlignCenter
      
      
      .Cells(r_int_NroFil, 8) = sum6
      .Cells(r_int_NroFil, 8).NumberFormat = "###,###,##0.00"
      .Cells(r_int_NroFil, 9).HorizontalAlignment = xlHAlignLeft
      

   End With
   
      
   
  
    fs_GenExcNuevo = ""
    '   fs_GenExcNuevo = "_49_9_" & Format(date, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".PDF"

      fs_GenExcNuevo = r_rst_Princi!HIPCUO_NUMOPE & ".PDF"
   
    r_obj_Excel.ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, FileName:="C:/PDF100/" & fs_GenExcNuevo, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    r_obj_Excel.ActiveWorkbook.Close SaveChanges:=False
   
   
    r_obj_Excel.Application.Quit
    Set r_obj_Excel = Nothing
   
    r_rst_Princi.Close
    Set r_rst_Princi = Nothing
    Screen.MousePointer = 0
   
End Function







Private Function fs_GenExc2() As String

Dim r_obj_Excel       As Excel.Application
Dim r_int_NumFil      As Integer
Dim r_str_Parame      As String



ActiveSheet.Pictures.Insert(g_str_RutLog & "\" & "imagen003.png").Select ' Insertamos la imagen en la hpja
'UserForm1.Image1.Picture = LoadPicture(g_str_RutLog & "\" & "imagen003.png") ' cargamos la imagen en el formulario


'   HIPCUO_TIPCRO
   
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
   
      'MARGENES
'      .PageSetup.LeftMargin = Application.CentimetersToPoints(1.5)
'      .PageSetup.RightMargin = Application.CentimetersToPoints(0.4)
'      .PageSetup.TopMargin = Application.CentimetersToPoints(1)
'      .PageSetup.BottomMargin = Application.CentimetersToPoints(1)
      
      .Columns("A").ColumnWidth = 10
      .Columns("B").ColumnWidth = 10
      .Columns("C").ColumnWidth = 8
      .Columns("D").ColumnWidth = 8
      .Columns("E").ColumnWidth = 8
      .Columns("F").ColumnWidth = 10
      .Columns("G").ColumnWidth = 10
      .Columns("H").ColumnWidth = 8
      
      
     
      
      .Columns("I").ColumnWidth = 8
      .Range(.Cells(1, 1), .Cells(600, 12)).Font.Name = "Arial (Western)"
      .Range(.Cells(1, 1), .Cells(600, 12)).Font.Size = 8
'      .Range(.Cells(1, 1), .Cells(600, 12)).RowHeight = 12
      
      .Pictures.Insert(g_str_RutLog & "\" & "imagen003.png").Select
      
      .Cells(6, 1) = "Estado de  Situación de la Cuenta"
      .Range(.Cells(6, 1), .Cells(6, 9)).Merge
'      .Range(.Cells(6, 1), .Cells(6, 9)).Font.Bold = True
      .Range(.Cells(6, 1), .Cells(6, 9)).Font.Size = 12
'      .Range(.Cells(7, 1), .Cells(7, 9)).Font.Underline = True
      .Range(.Cells(6, 1), .Cells(6, 9)).HorizontalAlignment = xlHAlignLeft
      
       r_obj_Excel.Visible = True
       
       .Cells(3, 2) = "Pruebas"
       .Cells(6, 1) = "Av. Rivera Navarrete 645(San Isidro)"
       .Range(.Cells(6, 1), .Cells(6, 6)).Merge
       .Range(.Cells(6, 1), .Cells(6, 6)).Font.Size = 12
       .Range(.Cells(6, 1), .Cells(6, 6)).HorizontalAlignment = xlHAlignLeft
      
      
       .Cells(6, 7) = "Banca Telefónica Lima: (01)221-8899"
       .Range(.Cells(6, 7), .Cells(6, 10)).Merge
       .Range(.Cells(6, 7), .Cells(6, 10)).Font.Size = 10
       .Range(.Cells(6, 7), .Cells(6, 10)).HorizontalAlignment = xlHAlignLeft
      
      
      .Cells(7, 1) = "Ruc: 20511904162"
      .Range(.Cells(7, 1), .Cells(7, 6)).Merge
      .Range(.Cells(7, 1), .Cells(7, 6)).Font.Size = 12
      .Range(.Cells(7, 1), .Cells(7, 6)).HorizontalAlignment = xlHAlignLeft
      
      .Cells(7, 7) = "www.micasita.com.pe"
      .Range(.Cells(7, 7), .Cells(7, 10)).Merge
      .Range(.Cells(7, 7), .Cells(7, 10)).Font.Size = 12
      .Range(.Cells(7, 7), .Cells(7, 10)).HorizontalAlignment = xlHAlignLeft
      
      
      .Cells(9, 1) = "INFORMACIÓN DEL CLIENTE"
      .Range(.Cells(9, 1), .Cells(9, 5)).Merge
      .Range(.Cells(9, 1), .Cells(9, 5)).Font.Size = 10
      .Range(.Cells(9, 1), .Cells(9, 5)).HorizontalAlignment = xlHAlignCenter
      
      
      .Cells(9, 6) = "INFORMACIÓN DEL CRÉDITO"
      .Range(.Cells(9, 6), .Cells(9, 10)).Merge
      .Range(.Cells(9, 6), .Cells(9, 10)).Font.Size = 10
      .Range(.Cells(9, 6), .Cells(9, 10)).HorizontalAlignment = xlHAlignCenter
      
      r_obj_Excel.Visible = True
      
      
       .Range(.Cells(10, 1), .Cells(10, 5)).Borders(xlEdgeTop).LineStyle = xlContinuous
       .Range(.Cells(10, 1), .Cells(18, 1)).Borders(xlEdgeLeft).LineStyle = xlContinuous
       .Range(.Cells(10, 5), .Cells(18, 5)).Borders(xlEdgeRight).LineStyle = xlContinuous
       .Range(.Cells(18, 1), .Cells(18, 5)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      
      
       .Range(.Cells(10, 6), .Cells(10, 9)).Borders(xlEdgeTop).LineStyle = xlContinuous
       .Range(.Cells(10, 6), .Cells(18, 6)).Borders(xlEdgeLeft).LineStyle = xlContinuous
       .Range(.Cells(10, 9), .Cells(18, 9)).Borders(xlEdgeRight).LineStyle = xlContinuous
       .Range(.Cells(18, 6), .Cells(18, 9)).Borders(xlEdgeBottom).LineStyle = xlContinuous
       
       
       
       r_obj_Excel.Visible = True
       
       .Cells(25, 2) = "<table><tr><td>1</td><td>2</td></tr></table>"
       
       
      
      
      
'      .Cells(20, 6).Pictures.Insert(g_str_RutLog & "\" & "img005.png").Select
      
      
'      .Rows(1).RowHeight = 1
'      .Rows(8).RowHeight = 9
'      .Rows(9).RowHeight = 5
'      .Range(.Cells(9, 1), .Cells(9, 9)).Borders(xlEdgeTop).LineStyle = xlContinuous
      '.CELDA(FILA, COLUMNA)
'      .Cells(2, 7) = "Nombre Reporte:"
'      .Cells(3, 7) = "Fecha Emisión:"
'      .Cells(4, 7) = "Hora Emisión:"
'      .Cells(5, 7) = "Página:"
      
'      .Cells(2, 9) = "PAGOS"
'          .Cells(2, 9).HorizontalAlignment = xlHAlignCenter
'      .Cells(3, 9) = Format(date, "dd/mm/yyyy")
'      .Cells(3, 9).HorizontalAlignment = xlHAlignCenter
'      .Cells(4, 9) = Format(Time, "hh:mm:ss")
'      .Cells(4, 9).HorizontalAlignment = xlHAlignCenter
'      .Cells(5, 9) = "1"
'      .Cells(5, 9).HorizontalAlignment = xlHAlignCenter
'      .Range(.Cells(2, 9), .Cells(5, 9)).HorizontalAlignment = xlHAlignRight
      
'      .Cells(10, 1) = "Nro Operación:"
'      .Cells(10, 2) = r_rst_Princi!HIPCUO_NUMOPE
'      .Range(.Cells(10, 1), .Cells(10, 10)).Font.Bold = True



'       .Range("A10:B10").Merge
'      .Range("A10") = "Nro Operación:"
'       .Range("C10:E10").Merge
'       .Range("C10") = r_rst_Princi!HIPCUO_NUMOPE
      
'       .Range("A11:B11").Merge
'      .Range("A11") = "Cliente:"
'       .Range("C11:E11").Merge
'       .Range("C11") = Trim(r_rst_Princi!DATGEN_APEPAT) & " " & Trim(r_rst_Princi!DATGEN_APEMAT) & " " & Trim(r_rst_Princi!DATGEN_NOMBRE)
'
'       .Range("A12:B12").Merge
'      .Range("A12") = "Producto:"
'       .Range("C12:F12").Merge
'       .Range("C12") = Trim(r_rst_Princi!PRODUC_DESCRI)
'
'       .Range("A13:B13").Merge
'      .Range("A13") = "Moneda:"
'       .Range("C13:E13").Merge
'       .Range("C13") = IIf(r_rst_Princi!HIPMAE_MONEDA = 1, "SOLES", "DOLARES AMERICANOS")
'
       
'      .Range("A14:B14").Merge
'      .Range("A14") = "Plazo:"
'       .Range("C14:E14").Merge
'       .Range("C14") = Trim(r_rst_Princi!HIPMAE_PLAANO & " Años")
       
       
   
       

      
     
      

   End With
   
      
   
  
  fs_GenExc2 = ""
'   fs_GenExcNuevo = "_49_9_" & Format(date, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".PDF"

      fs_GenExc2 = "R.PDF"
   
   r_obj_Excel.ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, FileName:="C:/PDF110/" & fs_GenExc2, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
   r_obj_Excel.ActiveWorkbook.Close SaveChanges:=False
   
   
   r_obj_Excel.Application.Quit
   Set r_obj_Excel = Nothing
   

   Screen.MousePointer = 0
   
End Function



Private Sub fs_GenWord()

      Dim objWord
      Dim objDoc, objNewDoc
      Dim objRange1, objRange2
      Dim n As String
      Dim r_rst_Princinuevo As ADODB.Recordset
      Dim r_str_Paramenuevo As String
 
      Set objWord = CreateObject("Word.Application")
      Set objDoc = objWord.Documents.Open("C:\pruebas5.docx")
   
      objWord.Visible = True
        
      n = ""
      n = n & "46445475', '43199183', '43388140"
         
      r_str_Paramenuevo = ""
      r_str_Paramenuevo = r_str_Paramenuevo & "select hipmae_numope as operacion, trim(hipmae_ndocli) as documento from cre_hipmae"

      r_str_Paramenuevo = r_str_Paramenuevo & "  where hipmae_ndocli in ('" & n & "') "
      ' MsgBox (r_str_Paramenuevo)

        If Not gf_EjecutaSQL(r_str_Paramenuevo, r_rst_Princinuevo, 3) Then
          Exit Sub
        End If
      Dim r_int_ConVer As Integer
      r_rst_Princinuevo.MoveFirst
      r_int_ConVer = 1
      Do While Not r_rst_Princinuevo.EOF
      
             Set objRange1 = objDoc.Bookmarks("nombre").Range
             Set objRange2 = objDoc.Bookmarks("apellido").Range
             objRange1.InsertAfter (r_rst_Princinuevo!OPERACION)
             objRange2.InsertAfter (r_rst_Princinuevo!DOCUMENTO)
             'objDoc.SaveAs "C:\" & r_rst_Princinuevo!OPERACION & ".docx", FileFormat:=wdFormatDocumentDefault
      
       r_int_ConVer = r_int_ConVer + 1
       r_rst_Princinuevo.MoveNext
       DoEvents
     Loop

    r_rst_Princinuevo.Close
    Set r_rst_Princinuevo = Nothing
    
    
End Sub


