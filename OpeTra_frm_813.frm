VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frm_Pro_AsgSegDes_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   9375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12045
   Icon            =   "OpeTra_frm_813.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9375
   ScaleWidth      =   12045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel13 
      Height          =   9375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12045
      _Version        =   65536
      _ExtentX        =   21246
      _ExtentY        =   16536
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
         TabIndex        =   1
         Top             =   30
         Width           =   11955
         _Version        =   65536
         _ExtentX        =   21087
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
         Begin Threed.SSPanel SSPanel7 
            Height          =   555
            Left            =   690
            TabIndex        =   2
            Top             =   30
            Width           =   4755
            _Version        =   65536
            _ExtentX        =   8387
            _ExtentY        =   979
            _StockProps     =   15
            Caption         =   "Endoso del Seguro de Desgravamen"
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
            Picture         =   "OpeTra_frm_813.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   30
         TabIndex        =   3
         Top             =   660
         Width           =   11955
         _Version        =   65536
         _ExtentX        =   21087
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
            Left            =   30
            Picture         =   "OpeTra_frm_813.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   11325
            Picture         =   "OpeTra_frm_813.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   615
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   2685
         Left            =   30
         TabIndex        =   6
         Top             =   1320
         Width           =   11955
         _Version        =   65536
         _ExtentX        =   21087
         _ExtentY        =   4736
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
            Height          =   2295
            Left            =   60
            TabIndex        =   7
            Top             =   330
            Width           =   11730
            _ExtentX        =   20690
            _ExtentY        =   4048
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
         Begin VB.Label Label2 
            Caption         =   "Datos del Crédito"
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
            TabIndex        =   8
            Top             =   60
            Width           =   1875
         End
      End
      Begin Threed.SSPanel SSPanel22 
         Height          =   5295
         Left            =   30
         TabIndex        =   9
         Top             =   4050
         Width           =   11955
         _Version        =   65536
         _ExtentX        =   21087
         _ExtentY        =   9340
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
            Height          =   5085
            Left            =   90
            TabIndex        =   10
            Top             =   150
            Width           =   11805
            _ExtentX        =   20823
            _ExtentY        =   8969
            _Version        =   393216
            Style           =   1
            Tabs            =   2
            TabsPerRow      =   4
            TabHeight       =   520
            TabCaption(0)   =   "Cliente - TNC - Cronograma Vigente"
            TabPicture(0)   =   "OpeTra_frm_813.frx":0B9A
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SSPanel74"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "SSPanel80"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "SSPanel73"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "SSPanel79"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "SSPanel78"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "grd_CliNCon_Listad"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "SSPanel76"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "SSPanel75"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).Control(8)=   "SSPanel72"
            Tab(0).Control(8).Enabled=   0   'False
            Tab(0).Control(9)=   "SSPanel71"
            Tab(0).Control(9).Enabled=   0   'False
            Tab(0).ControlCount=   10
            TabCaption(1)   =   "Cliente - TNC - Cronograma Regenerado"
            TabPicture(1)   =   "OpeTra_frm_813.frx":0BB6
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SSPanel15"
            Tab(1).Control(1)=   "SSPanel14"
            Tab(1).Control(2)=   "SSPanel12"
            Tab(1).Control(3)=   "SSPanel11"
            Tab(1).Control(4)=   "SSPanel10"
            Tab(1).Control(5)=   "SSPanel9"
            Tab(1).Control(6)=   "SSPanel6"
            Tab(1).Control(7)=   "SSPanel5"
            Tab(1).Control(8)=   "SSPanel3"
            Tab(1).Control(9)=   "grd_CliNConR_Listad"
            Tab(1).ControlCount=   10
            Begin Threed.SSPanel SSPanel23 
               Height          =   285
               Left            =   -67530
               TabIndex        =   11
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
               TabIndex        =   12
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
               TabIndex        =   13
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
               TabIndex        =   14
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
               TabIndex        =   15
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
               TabIndex        =   16
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
               TabIndex        =   17
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
               TabIndex        =   18
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
               TabIndex        =   19
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
               TabIndex        =   20
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
               TabIndex        =   21
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
               TabIndex        =   22
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
               TabIndex        =   23
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
               TabIndex        =   24
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
               TabIndex        =   25
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
               TabIndex        =   26
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
               TabIndex        =   27
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
               TabIndex        =   28
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
               TabIndex        =   29
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
               TabIndex        =   30
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
               TabIndex        =   31
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
               TabIndex        =   32
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
               TabIndex        =   33
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
               TabIndex        =   34
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
               TabIndex        =   35
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
               TabIndex        =   36
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
               TabIndex        =   37
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
               TabIndex        =   38
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
               TabIndex        =   39
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
               TabIndex        =   40
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
               TabIndex        =   41
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
               TabIndex        =   42
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
               TabIndex        =   43
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
               TabIndex        =   44
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
               TabIndex        =   45
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
               TabIndex        =   46
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
               TabIndex        =   47
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
            Begin Threed.SSPanel SSPanel56 
               Height          =   285
               Left            =   -70710
               TabIndex        =   49
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
               TabIndex        =   50
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
               TabIndex        =   51
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
               TabIndex        =   52
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
               TabIndex        =   53
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
               TabIndex        =   54
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
               TabIndex        =   55
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
               TabIndex        =   56
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
               TabIndex        =   57
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
               TabIndex        =   58
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
               TabIndex        =   59
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
            Begin Threed.SSPanel SSPanel71 
               Height          =   285
               Left            =   135
               TabIndex        =   68
               Top             =   390
               Width           =   840
               _Version        =   65536
               _ExtentX        =   1482
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
               Left            =   960
               TabIndex        =   69
               Top             =   390
               Width           =   1245
               _Version        =   65536
               _ExtentX        =   2196
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
            Begin Threed.SSPanel SSPanel75 
               Height          =   285
               Left            =   9645
               TabIndex        =   70
               Top             =   390
               Width           =   1230
               _Version        =   65536
               _ExtentX        =   2170
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
               Left            =   3450
               TabIndex        =   71
               Top             =   390
               Width           =   1245
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
            Begin MSFlexGridLib.MSFlexGrid grd_CliNCon_Listad 
               Height          =   4305
               Left            =   120
               TabIndex        =   72
               Top             =   705
               Width           =   11280
               _ExtentX        =   19897
               _ExtentY        =   7594
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
            Begin Threed.SSPanel SSPanel78 
               Height          =   285
               Left            =   4680
               TabIndex        =   73
               Top             =   390
               Width           =   1275
               _Version        =   65536
               _ExtentX        =   2249
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
               Left            =   5940
               TabIndex        =   74
               Top             =   390
               Width           =   1275
               _Version        =   65536
               _ExtentX        =   2249
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
            Begin MSFlexGridLib.MSFlexGrid grd_CliNConR_Listad 
               Height          =   4305
               Left            =   -74880
               TabIndex        =   75
               Top             =   705
               Width           =   11280
               _ExtentX        =   19897
               _ExtentY        =   7594
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
            Begin Threed.SSPanel SSPanel73 
               Height          =   285
               Left            =   2190
               TabIndex        =   76
               Top             =   390
               Width           =   1275
               _Version        =   65536
               _ExtentX        =   2249
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
            Begin Threed.SSPanel SSPanel80 
               Height          =   285
               Left            =   7200
               TabIndex        =   77
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
            Begin Threed.SSPanel SSPanel74 
               Height          =   285
               Left            =   8430
               TabIndex        =   78
               Top             =   390
               Width           =   1230
               _Version        =   65536
               _ExtentX        =   2170
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
            Begin Threed.SSPanel SSPanel3 
               Height          =   285
               Left            =   -74865
               TabIndex        =   79
               Top             =   390
               Width           =   840
               _Version        =   65536
               _ExtentX        =   1482
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
            Begin Threed.SSPanel SSPanel5 
               Height          =   285
               Left            =   -74040
               TabIndex        =   80
               Top             =   390
               Width           =   1245
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
            Begin Threed.SSPanel SSPanel6 
               Height          =   285
               Left            =   -65355
               TabIndex        =   81
               Top             =   390
               Width           =   1230
               _Version        =   65536
               _ExtentX        =   2170
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
            Begin Threed.SSPanel SSPanel9 
               Height          =   285
               Left            =   -71550
               TabIndex        =   82
               Top             =   390
               Width           =   1245
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
            Begin Threed.SSPanel SSPanel10 
               Height          =   285
               Left            =   -70320
               TabIndex        =   83
               Top             =   390
               Width           =   1275
               _Version        =   65536
               _ExtentX        =   2249
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
            Begin Threed.SSPanel SSPanel11 
               Height          =   285
               Left            =   -69060
               TabIndex        =   84
               Top             =   390
               Width           =   1275
               _Version        =   65536
               _ExtentX        =   2249
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
            Begin Threed.SSPanel SSPanel12 
               Height          =   285
               Left            =   -72810
               TabIndex        =   85
               Top             =   390
               Width           =   1275
               _Version        =   65536
               _ExtentX        =   2249
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
            Begin Threed.SSPanel SSPanel14 
               Height          =   285
               Left            =   -67800
               TabIndex        =   86
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
            Begin Threed.SSPanel SSPanel15 
               Height          =   285
               Left            =   -66570
               TabIndex        =   87
               Top             =   390
               Width           =   1230
               _Version        =   65536
               _ExtentX        =   2170
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
            Begin VB.Label lbl_Totale 
               Alignment       =   1  'Right Justify
               Caption         =   "Totales ===> US$ "
               Height          =   315
               Index           =   5
               Left            =   -74610
               TabIndex        =   67
               Top             =   6870
               Width           =   1845
            End
            Begin VB.Label lbl_Totale 
               Alignment       =   1  'Right Justify
               Caption         =   "Totales ===> US$ "
               Height          =   315
               Index           =   4
               Left            =   -74610
               TabIndex        =   66
               Top             =   6870
               Width           =   1845
            End
            Begin VB.Label Label1 
               Caption         =   "Totales ==>"
               Height          =   285
               Left            =   -72930
               TabIndex        =   65
               Top             =   1470
               Width           =   945
            End
            Begin VB.Label Label14 
               Caption         =   "Totales ==>"
               Height          =   285
               Left            =   -72930
               TabIndex        =   64
               Top             =   1470
               Width           =   945
            End
            Begin VB.Label Label15 
               Caption         =   "Totales ==>"
               Height          =   285
               Left            =   -73230
               TabIndex        =   63
               Top             =   1470
               Width           =   945
            End
            Begin VB.Label lbl_Totale 
               Alignment       =   1  'Right Justify
               Caption         =   "Totales ===> US$ "
               Height          =   315
               Index           =   1
               Left            =   -74610
               TabIndex        =   62
               Top             =   6870
               Width           =   1845
            End
            Begin VB.Label lbl_Totale 
               Alignment       =   1  'Right Justify
               Caption         =   "Totales ===> US$ "
               Height          =   315
               Index           =   2
               Left            =   -74790
               TabIndex        =   61
               Top             =   6870
               Width           =   1845
            End
            Begin VB.Label lbl_Totale 
               Alignment       =   1  'Right Justify
               Caption         =   "Totales ===> US$ "
               Height          =   315
               Index           =   3
               Left            =   -74610
               TabIndex        =   60
               Top             =   6870
               Width           =   1845
            End
         End
      End
   End
End
Attribute VB_Name = "frm_Pro_AsgSegDes_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim r_int_NumCuo        As Integer
Dim r_dbl_MtoAse        As Double
Dim l_Arr_TNC_Cli()     As String
Dim l_Arr_TC_Cli()      As String
Dim l_Arr_TNC_Cof()     As String
Dim l_Arr_TC_Cof()      As String
Dim l_int_PagCuo        As Integer

Private Sub cmd_Grabar_Click()
Dim r_int_Contad        As Integer
Dim r_int_ConSel        As Integer
  
   'Confirma
   If MsgBox("¿Está seguro de Actualizar?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   r_int_NumCuo = 0

   If fs_Actualiza_Cronograma_CLITNC Then
      MsgBox "Actualización realizada satisfactoriamente.", vbInformation, modgen_g_str_NomPlt
   Else
      MsgBox "Error en Asignación de Seguro de Desgravamen.", vbExclamation, modgen_g_str_NomPlt
   End If
   
   Call gs_LimpiaGrid(grd_CliNCon_Listad)
   Call gs_LimpiaGrid(grd_CliNConR_Listad)
   Screen.MousePointer = 0
   Unload Me
End Sub

Private Function fs_Actualiza_Cronograma_CLITNC() As Boolean
Dim r_int_Contad     As Integer
Dim r_str_Cadena     As String
Dim r_rst_Cuotas     As ADODB.Recordset
Dim r_str_FecVct     As String
Dim r_dbl_Capita     As Double
Dim r_dbl_Intere     As Double
Dim r_dbl_Portes     As Double
Dim r_dbl_SegPre     As Double
Dim r_dbl_SegViv     As Double
Dim r_dbl_MtoCuo     As Double
Dim r_dbl_SalCap     As Double
Dim r_str_CadErr     As String
Dim r_int_ConAux     As Integer

   fs_Actualiza_Cronograma_CLITNC = True
   
   '*****************
   'Actualiza CLIENTE TNC
   r_int_NumCuo = 0
   r_str_FecVct = Mid(Trim(grd_CliNConR_Listad.TextMatrix(r_int_Contad, 1)), 7, 4) & Mid(Trim(grd_CliNConR_Listad.TextMatrix(r_int_Contad, 1)), 4, 2) & Mid(Trim(grd_CliNCon_Listad.TextMatrix(r_int_Contad, 1)), 1, 2)
   
   'obtiene numero de cuota
   r_str_Cadena = ""
   r_str_Cadena = r_str_Cadena & "SELECT HIPCUO_NUMCUO FROM CRE_HIPCUO "
   r_str_Cadena = r_str_Cadena & " WHERE HIPCUO_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   r_str_Cadena = r_str_Cadena & "   AND HIPCUO_TIPCRO = 1 "
   r_str_Cadena = r_str_Cadena & "   AND HIPCUO_FECVCT = " & r_str_FecVct & " "
   
   If Not gf_EjecutaSQL(r_str_Cadena, r_rst_Cuotas, 3) Then
      MsgBox "No se pudo obtener la cuota a partir de la cual se reemplazará el cronograma CLIENTE TNC.", vbExclamation, modgen_g_str_NomPlt
      Exit Function
   End If
   
   If Not (r_rst_Cuotas.BOF And r_rst_Cuotas.EOF) Then
      r_rst_Cuotas.MoveFirst
      r_int_NumCuo = r_rst_Cuotas!HIPCUO_NUMCUO
      r_int_ConAux = r_int_NumCuo
   End If
   
   r_rst_Cuotas.Close
   Set r_rst_Cuotas = Nothing
   
   If r_int_NumCuo = 0 Then
      MsgBox "Error, cuota no puede ser cero. Cronograma CLIENTE TNC.", vbExclamation, modgen_g_str_NomPlt
      Exit Function
   End If
   
   'elimina cuotas a reemplazar de la BD
   r_str_Cadena = ""
   r_str_Cadena = r_str_Cadena & "DELETE FROM CRE_HIPCUO "
   r_str_Cadena = r_str_Cadena & " WHERE HIPCUO_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   r_str_Cadena = r_str_Cadena & "   AND HIPCUO_TIPCRO = 1 "
   r_str_Cadena = r_str_Cadena & "   AND HIPCUO_NUMCUO >= " & r_int_NumCuo & " "
   
   If Not gf_EjecutaSQL(r_str_Cadena, r_rst_Cuotas, 2) Then
      MsgBox "Error al eliminar las cuotas del cronograma CLIENTE TNC.", vbExclamation, modgen_g_str_NomPlt
      Exit Function
   End If
   
   For r_int_Contad = 0 To grd_CliNConR_Listad.Rows - 1
      'carga variables e inserta cuota
      r_int_NumCuo = grd_CliNConR_Listad.TextMatrix(r_int_Contad, 0)
      r_str_FecVct = grd_CliNConR_Listad.TextMatrix(r_int_Contad, 1)
      r_dbl_Capita = grd_CliNConR_Listad.TextMatrix(r_int_Contad, 2)
      r_dbl_Intere = grd_CliNConR_Listad.TextMatrix(r_int_Contad, 3)
      r_dbl_SegPre = grd_CliNConR_Listad.TextMatrix(r_int_Contad, 4)
      r_dbl_SegViv = grd_CliNConR_Listad.TextMatrix(r_int_Contad, 5)
      r_dbl_Portes = grd_CliNConR_Listad.TextMatrix(r_int_Contad, 6)
      r_dbl_MtoCuo = grd_CliNConR_Listad.TextMatrix(r_int_Contad, 7)
      r_dbl_SalCap = grd_CliNConR_Listad.TextMatrix(r_int_Contad, 8)
      
      If Not ff_Inserta_HipCuo(moddat_g_str_NumOpe, 1, r_int_NumCuo, r_str_FecVct, r_dbl_Capita, r_dbl_Intere, r_dbl_SegPre, r_dbl_SegViv, r_dbl_Portes, r_dbl_SalCap, 0, 0, 0) Then
         Exit Function 'Exit For
      End If
   Next r_int_Contad
   
   fs_Actualiza_Cronograma_CLITNC = True
   
   '****REGISTRAR LOG
   If (fs_Actualiza_Cronograma_CLITNC = True) Then
       g_str_Parame = ""
       g_str_Parame = g_str_Parame & "INSERT INTO CRE_SEGINM ("
       g_str_Parame = g_str_Parame & "SEGINM_NUMOPE, "
       g_str_Parame = g_str_Parame & "SEGINM_TIPCAR, "
       g_str_Parame = g_str_Parame & "SEGINM_FECCAR, "
       g_str_Parame = g_str_Parame & "SEGINM_HORCAR, "
       g_str_Parame = g_str_Parame & "SEGINM_MTOSEG, "
       g_str_Parame = g_str_Parame & "SEGINM_CUOCAR, "
       g_str_Parame = g_str_Parame & "SEGUSUCRE, "
       g_str_Parame = g_str_Parame & "SEGFECCRE, "
       g_str_Parame = g_str_Parame & "SEGHORCRE, "
       g_str_Parame = g_str_Parame & "SEGPLTCRE, "
       g_str_Parame = g_str_Parame & "SEGTERCRE, "
       g_str_Parame = g_str_Parame & "SEGSUCCRE) "
       g_str_Parame = g_str_Parame & "VALUES ( "
       g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "' , "
       g_str_Parame = g_str_Parame & 4 & " , "
       g_str_Parame = g_str_Parame & Format(CDate(moddat_g_str_FecSis), "YYYYMMDD") & " , "
       g_str_Parame = g_str_Parame & Format(Time, "HHMMSS") & " , "
       g_str_Parame = g_str_Parame & 0 & " , "
       g_str_Parame = g_str_Parame & CInt(r_int_ConAux) & " , "
       g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "' ,"
       g_str_Parame = g_str_Parame & " " & Format(CDate(moddat_g_str_FecSis), "YYYYMMDD") & ", "
       g_str_Parame = g_str_Parame & " " & Format(Time, "HHMMSS") & ", "
       g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
       g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "' ,"
       g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "')"
                                      
       If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
          Exit Function
       End If
   End If
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

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_IniciaGrid
   Call fs_Limpiar
   'Buscando Información del Crédito
   Call modmip_gs_DatNumOpe(moddat_g_str_NumOpe, grd_Listad) 'fs_Buscar_DatosCredito
   Call fs_Cargar_Cron01
   Call fs_Cargar_Cron01_Reg
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_IniciaGrid()
   'Datos del Credito
   grd_Listad.ColWidth(0) = 2800
   grd_Listad.ColWidth(1) = 8400
   grd_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
   grd_Listad.Rows = 0
   
    'Cliente No Concesional
   grd_CliNCon_Listad.ColWidth(0) = 810
   grd_CliNCon_Listad.ColWidth(1) = 1250
   grd_CliNCon_Listad.ColWidth(2) = 1250
   grd_CliNCon_Listad.ColWidth(3) = 1250
   grd_CliNCon_Listad.ColWidth(4) = 1250
   grd_CliNCon_Listad.ColWidth(5) = 1250
   grd_CliNCon_Listad.ColWidth(6) = 1250
   grd_CliNCon_Listad.ColWidth(7) = 1230
   grd_CliNCon_Listad.ColWidth(8) = 1230
   grd_CliNCon_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_CliNCon_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_CliNCon_Listad.ColAlignment(2) = flexAlignRightCenter
   grd_CliNCon_Listad.ColAlignment(3) = flexAlignRightCenter
   grd_CliNCon_Listad.ColAlignment(4) = flexAlignRightCenter
   grd_CliNCon_Listad.ColAlignment(5) = flexAlignRightCenter
   grd_CliNCon_Listad.ColAlignment(6) = flexAlignRightCenter
   grd_CliNCon_Listad.ColAlignment(7) = flexAlignRightCenter
   grd_CliNCon_Listad.ColAlignment(8) = flexAlignRightCenter
   
   'Cliente No Concesional Regenerado
   grd_CliNConR_Listad.ColWidth(0) = 810
   grd_CliNConR_Listad.ColWidth(1) = 1250
   grd_CliNConR_Listad.ColWidth(2) = 1250
   grd_CliNConR_Listad.ColWidth(3) = 1250
   grd_CliNConR_Listad.ColWidth(4) = 1250
   grd_CliNConR_Listad.ColWidth(5) = 1250
   grd_CliNConR_Listad.ColWidth(6) = 1250
   grd_CliNConR_Listad.ColWidth(7) = 1230
   grd_CliNConR_Listad.ColWidth(8) = 1230
   grd_CliNConR_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_CliNConR_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_CliNConR_Listad.ColAlignment(2) = flexAlignRightCenter
   grd_CliNConR_Listad.ColAlignment(3) = flexAlignRightCenter
   grd_CliNConR_Listad.ColAlignment(4) = flexAlignRightCenter
   grd_CliNConR_Listad.ColAlignment(5) = flexAlignRightCenter
   grd_CliNConR_Listad.ColAlignment(6) = flexAlignRightCenter
   grd_CliNConR_Listad.ColAlignment(7) = flexAlignRightCenter
   grd_CliNConR_Listad.ColAlignment(8) = flexAlignRightCenter
End Sub

Private Sub fs_Limpiar()
   Call gs_LimpiaGrid(grd_Listad)
   Call gs_LimpiaGrid(grd_CliNCon_Listad)
   Call gs_LimpiaGrid(grd_CliNConR_Listad)
End Sub

Private Sub fs_Cargar_Cron01()
   r_int_NumCuo = 0
   'Buscando Información del Cronograma Tipo 1
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT HIPCUO_NUMCUO, HIPCUO_FECVCT, HIPCUO_CAPITA, HIPCUO_INTERE, "
   g_str_Parame = g_str_Parame & "        HIPCUO_DESORG, HIPCUO_VIVORG, HIPCUO_OTRORG, HIPCUO_SALCAP, HIPCUO_SITUAC "
   g_str_Parame = g_str_Parame & "   FROM CRE_HIPCUO "
   g_str_Parame = g_str_Parame & "  WHERE HIPCUO_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   g_str_Parame = g_str_Parame & "    AND HIPCUO_TIPCRO = 1 "
   g_str_Parame = g_str_Parame & "  ORDER BY HIPCUO_NUMCUO "
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
  
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_CliNCon_Listad.Redraw = False
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         grd_CliNCon_Listad.Rows = grd_CliNCon_Listad.Rows + 1
         grd_CliNCon_Listad.Row = grd_CliNCon_Listad.Rows - 1
         grd_CliNCon_Listad.Col = 0
         grd_CliNCon_Listad.Text = Format(g_rst_Princi!HIPCUO_NUMCUO, "000")
         grd_CliNCon_Listad.Col = 1
         grd_CliNCon_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECVCT))
         grd_CliNCon_Listad.Col = 2
         grd_CliNCon_Listad.Text = Format(g_rst_Princi!HIPCUO_CAPITA, "###,###,##0.00")
         grd_CliNCon_Listad.Col = 3
         grd_CliNCon_Listad.Text = Format(g_rst_Princi!HIPCUO_INTERE, "###,###,##0.00")
         grd_CliNCon_Listad.Col = 4
         grd_CliNCon_Listad.Text = Format(g_rst_Princi!HIPCUO_DESORG, "###,###,##0.00")
         grd_CliNCon_Listad.Col = 5
         grd_CliNCon_Listad.Text = Format(g_rst_Princi!HIPCUO_VIVORG, "###,###,##0.00")
         grd_CliNCon_Listad.Col = 6
         grd_CliNCon_Listad.Text = Format(g_rst_Princi!HIPCUO_OTRORG, "###,###,##0.00")
         grd_CliNCon_Listad.Col = 7
         grd_CliNCon_Listad.Text = Format(g_rst_Princi!HIPCUO_CAPITA + g_rst_Princi!HIPCUO_INTERE + g_rst_Princi!HIPCUO_DESORG + g_rst_Princi!HIPCUO_VIVORG + g_rst_Princi!HIPCUO_OTRORG, "###,###,##0.00")
         grd_CliNCon_Listad.Col = 8
         grd_CliNCon_Listad.Text = Format(g_rst_Princi!HIPCUO_SALCAP, "###,###,##0.00")
         
         g_rst_Princi.MoveNext
      Loop
      
      Call gs_UbiIniGrid(grd_CliNCon_Listad)
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      grd_CliNCon_Listad.Redraw = True
   End If
End Sub

Private Sub fs_Cargar_Cron01_Reg()
Dim int_Produc    As Integer
Dim int_CuoDbl    As Integer
Dim dbl_ValInm    As Double
Dim dbl_CuoIni    As Double
Dim dbl_MtoCon    As Double
Dim dbl_MtoTas    As Double
Dim int_PlaPre    As Integer
Dim dbl_TasInt    As Double
Dim dbl_TasCof    As Double
Dim dbl_ComCof    As Double
Dim str_FecDes    As String
Dim dat_FecDes    As Date
Dim int_DiaVct    As Integer
Dim int_PerGra    As Integer
Dim str_PriVct    As String
Dim dbl_Portes    As Double
Dim dbl_SegViv    As Double
Dim int_TipSDe    As Integer
Dim dbl_SegDes    As Double
           
Dim obj_Cronog    As Object

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT HIPMAE_NUMOPE, HIPMAE_CUOANO, HIPMAE_MONEDA, HIPMAE_CVTSOL, HIPMAE_CVTDOL, HIPMAE_APOSOL, "
   g_str_Parame = g_str_Parame & "        HIPMAE_APODOL, HIPMAE_NUMCUO, HIPMAE_TASINT, HIPMAE_FOIPRE, HIPMAE_FOIVIV, HIPMAE_COMCOF, "
   g_str_Parame = g_str_Parame & "        HIPMAE_TASCOF, HIPMAE_DIAPAG, HIPMAE_PRXVCT, HIPMAE_OTRIMP, HIPMAE_TIPSEG, HIPMAE_SALCON, "
   g_str_Parame = g_str_Parame & "        HIPMAE_SALCAP,HIPMAE_FECDES, HIPMAE_CUOPAG "
   g_str_Parame = g_str_Parame & "   FROM CRE_HIPMAE "
   g_str_Parame = g_str_Parame & "  WHERE HIPMAE_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
     
      int_Produc = 1
      int_CuoDbl = g_rst_Princi!HIPMAE_CUOANO
      
      If g_rst_Princi!HIPMAE_MONEDA = 1 Then
         dbl_ValInm = g_rst_Princi!HIPMAE_CVTSOL + g_rst_Princi!HIPMAE_SALCON
         dbl_CuoIni = g_rst_Princi!HIPMAE_CVTSOL - g_rst_Princi!HIPMAE_SALCAP
      Else
         dbl_ValInm = g_rst_Princi!HIPMAE_CVTDOL + g_rst_Princi!HIPMAE_SALCON
         dbl_CuoIni = g_rst_Princi!HIPMAE_CVTDOL - g_rst_Princi!HIPMAE_SALCAP
      End If
      
      dbl_MtoCon = g_rst_Princi!HIPMAE_SALCON
      Call fs_Buscar_Tasacion
      dbl_MtoTas = r_dbl_MtoAse
      int_PlaPre = g_rst_Princi!HIPMAE_NUMCUO - g_rst_Princi!HIPMAE_CUOPAG
      dbl_TasInt = g_rst_Princi!HIPMAE_TASINT
      dbl_TasCof = g_rst_Princi!HIPMAE_TASCOF
      dbl_ComCof = g_rst_Princi!HIPMAE_COMCOF
      str_FecDes = gf_FormatoFecha(fs_Obtiene_FechaPago(g_rst_Princi!HIPMAE_NUMOPE, 1, g_rst_Princi!HIPMAE_FECDES))
      dat_FecDes = CDate(Format(str_FecDes, "dd/mm/yyyy")) '
      int_DiaVct = g_rst_Princi!HIPMAE_DIAPAG
      int_PerGra = 0
      str_PriVct = gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_PRXVCT))
      dbl_Portes = CDbl(g_rst_Princi!HIPMAE_OTRIMP)
      dbl_SegViv = CDbl(g_rst_Princi!HIPMAE_FOIVIV)
      int_TipSDe = CInt(g_rst_Princi!HIPMAE_TIPSEG) - 10
      dbl_SegDes = 0
      l_int_PagCuo = g_rst_Princi!HIPMAE_CUOPAG
      
      'Calculando cronogramas
      Set obj_Cronog = CreateObject("ComCronograma.ClsCronograma")
      Call obj_Cronog.Listar(l_Arr_TNC_Cli(), l_Arr_TC_Cli(), l_Arr_TNC_Cof(), l_Arr_TC_Cof(), int_Produc, int_CuoDbl, dbl_ValInm, dbl_CuoIni, dbl_MtoCon, dbl_MtoTas, int_PlaPre, dbl_TasInt, dbl_TasCof, dbl_ComCof, dat_FecDes, 0, int_DiaVct, str_PriVct, int_PerGra, dbl_Portes, dbl_SegViv, int_TipSDe, dbl_SegDes)
      
      'Mostrando Cronograma 1
      Call fs_Muestra_Cron01
   End If
End Sub

Private Sub fs_Muestra_Cron01()
Dim r_dbl_Cuo_Capita    As Double
Dim r_dbl_Cuo_Intere    As Double
Dim r_dbl_Cuo_SegPre    As Double
Dim r_dbl_Cuo_SegViv    As Double
Dim r_dbl_Cuo_Portes    As Double
Dim r_dbl_Cuo_TotCuo    As Double
Dim r_dbl_Tot_Capita    As Double
Dim r_dbl_Tot_Intere    As Double
Dim r_dbl_Tot_SegPre    As Double
Dim r_dbl_Tot_SegViv    As Double
Dim r_dbl_Tot_Portes    As Double
Dim r_dbl_Tot_TotCuo    As Double
Dim r_int_Contad        As Integer

   grd_CliNConR_Listad.Redraw = False
   Call gs_LimpiaGrid(grd_CliNConR_Listad)
   r_dbl_Tot_Capita = 0
   r_dbl_Tot_Intere = 0
   r_dbl_Tot_SegPre = 0
   r_dbl_Tot_SegViv = 0
   r_dbl_Tot_Portes = 0
   r_dbl_Tot_TotCuo = 0
   
   For r_int_Contad = 1 To UBound(l_Arr_TNC_Cli)
      grd_CliNConR_Listad.Rows = grd_CliNConR_Listad.Rows + 1
      grd_CliNConR_Listad.Row = grd_CliNConR_Listad.Rows - 1
      
      r_dbl_Cuo_Capita = CDbl(Format(l_Arr_TNC_Cli(r_int_Contad, 4), "###,##0.00"))
      r_dbl_Cuo_Intere = CDbl(Format(l_Arr_TNC_Cli(r_int_Contad, 5), "###,##0.00"))
      r_dbl_Cuo_SegPre = CDbl(Format(l_Arr_TNC_Cli(r_int_Contad, 6), "###,##0.00"))
      r_dbl_Cuo_SegViv = CDbl(Format(l_Arr_TNC_Cli(r_int_Contad, 7), "###,##0.00"))
      r_dbl_Cuo_Portes = CDbl(Format(l_Arr_TNC_Cli(r_int_Contad, 8), "###,##0.00"))
      r_dbl_Cuo_TotCuo = CDbl(Format(l_Arr_TNC_Cli(r_int_Contad, 9), "###,##0.00"))
      r_dbl_Tot_Capita = r_dbl_Tot_Capita + r_dbl_Cuo_Capita
      r_dbl_Tot_Intere = r_dbl_Tot_Intere + r_dbl_Cuo_Intere
      r_dbl_Tot_SegPre = r_dbl_Tot_SegPre + r_dbl_Cuo_SegPre
      r_dbl_Tot_SegViv = r_dbl_Tot_SegViv + r_dbl_Cuo_SegViv
      r_dbl_Tot_Portes = r_dbl_Tot_Portes + r_dbl_Cuo_Portes
      r_dbl_Tot_TotCuo = r_dbl_Tot_TotCuo + r_dbl_Cuo_TotCuo
      
      grd_CliNConR_Listad.Col = 0
      grd_CliNConR_Listad.Text = Format(r_int_Contad + l_int_PagCuo, "000")
      
      grd_CliNConR_Listad.Col = 1
      grd_CliNConR_Listad.Text = Format(l_Arr_TNC_Cli(r_int_Contad, 2), "dd/mm/yyyy")
      
      grd_CliNConR_Listad.Col = 2
      grd_CliNConR_Listad.Text = Format(r_dbl_Cuo_Capita, "###,##0.00")
      
      grd_CliNConR_Listad.Col = 3
      grd_CliNConR_Listad.Text = Format(r_dbl_Cuo_Intere, "###,##0.00")
      
      grd_CliNConR_Listad.Col = 4
      grd_CliNConR_Listad.Text = Format(r_dbl_Cuo_SegPre, "###,##0.00")
      
      grd_CliNConR_Listad.Col = 5
      grd_CliNConR_Listad.Text = Format(r_dbl_Cuo_SegViv, "###,##0.00")
      
      grd_CliNConR_Listad.Col = 6
      grd_CliNConR_Listad.Text = Format(r_dbl_Cuo_Portes, "###,##0.00")
      
      grd_CliNConR_Listad.Col = 7
      grd_CliNConR_Listad.Text = Format(r_dbl_Cuo_TotCuo, "###,##0.00")
      
      grd_CliNConR_Listad.Col = 8
      grd_CliNConR_Listad.Text = Format(l_Arr_TNC_Cli(r_int_Contad, 10), "###,##0.00")
   Next r_int_Contad
   
   grd_CliNConR_Listad.Redraw = True
   Call gs_UbiIniGrid(grd_CliNConR_Listad)
End Sub

Private Sub fs_Buscar_Tasacion()
Dim r_rst_Temp    As Recordset
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT EVATAS_TIPMON, EVATAS_SUMASE_INM, EVATAS_SUMASE_ES1, EVATAS_SUMASE_ES2, EVATAS_SUMASE_DEP "
   g_str_Parame = g_str_Parame & "  FROM TRA_EVATAS "
   g_str_Parame = g_str_Parame & " WHERE EVATAS_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Temp, 3) Then
       Exit Sub
   End If
   
   If Not (r_rst_Temp.BOF And r_rst_Temp.EOF) Then
      r_rst_Temp.MoveFirst
      r_dbl_MtoAse = gf_FormatoNumero(r_rst_Temp!EVATAS_SUMASE_INM + r_rst_Temp!EVATAS_SUMASE_ES1 + r_rst_Temp!EVATAS_SUMASE_ES2 + r_rst_Temp!EVATAS_SUMASE_DEP, 12, 2) & " "
   End If
   
   r_rst_Temp.Close
   Set r_rst_Temp = Nothing
End Sub

Private Function fs_Obtiene_FechaPago(ByVal p_NumOpe As String, ByVal p_TipCro As Integer, ByVal p_FecDes As String) As String
Dim r_rst_Temp    As Recordset
   fs_Obtiene_FechaPago = p_FecDes
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT HIPCUO_FECVCT "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPCUO "
   g_str_Parame = g_str_Parame & " WHERE HIPCUO_NUMOPE = '" & p_NumOpe & "' "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_TIPCRO = " & p_TipCro & " "
   g_str_Parame = g_str_Parame & "ORDER BY HIPCUO_FECVCT DESC"
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Temp, 3) Then
       Exit Function
   End If
   
   If Not (r_rst_Temp.BOF And r_rst_Temp.EOF) Then
      r_rst_Temp.MoveFirst
      fs_Obtiene_FechaPago = r_rst_Temp!HIPCUO_FECVCT
   End If
   
   r_rst_Temp.Close
   Set r_rst_Temp = Nothing
End Function

Private Sub fs_Buscar_DatosCredito_ant()
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
   
   'Situación de Crédito
   moddat_g_int_Situac = g_rst_Princi!HIPMAE_SITUAC
   moddat_g_str_Situac = moddat_gf_Consulta_ParDes("027", CStr(g_rst_Princi!HIPMAE_SITUAC))
   
   'Obteniendo Información del Inmueble
   Call moddat_gs_Consulta_DatInm(moddat_g_str_NumSol, moddat_g_str_Direcc, moddat_g_str_Distri, r_str_CodPry, r_str_NomPry, r_str_CodBco)
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub


