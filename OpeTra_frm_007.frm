VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frm_Des_CreHip_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   10545
   ClientLeft      =   1200
   ClientTop       =   405
   ClientWidth     =   12810
   Icon            =   "OpeTra_frm_007.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10545
   ScaleWidth      =   12810
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   10545
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   12825
      _Version        =   65536
      _ExtentX        =   22622
      _ExtentY        =   18600
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
      Begin Threed.SSPanel SSPanel3 
         Height          =   2205
         Left            =   30
         TabIndex        =   75
         Top             =   2730
         Width           =   12735
         _Version        =   65536
         _ExtentX        =   22463
         _ExtentY        =   3889
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
         Begin TabDlg.SSTab tab_Princi 
            Height          =   2085
            Left            =   60
            TabIndex        =   78
            Top             =   60
            Width           =   12585
            _ExtentX        =   22199
            _ExtentY        =   3678
            _Version        =   393216
            Style           =   1
            Tabs            =   4
            TabsPerRow      =   4
            TabHeight       =   520
            TabCaption(0)   =   "Datos Crediticios"
            TabPicture(0)   =   "OpeTra_frm_007.frx":000C
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "lbl_NomGlo(25)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "lbl_NomGlo(26)"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "lbl_NomGlo(27)"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "lbl_NomGlo(29)"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "lbl_NomGlo(30)"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "lbl_NomGlo(31)"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "lbl_NomGlo(32)"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "lbl_NomGlo(34)"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).Control(8)=   "lbl_NomGlo(21)"
            Tab(0).Control(8).Enabled=   0   'False
            Tab(0).Control(9)=   "pnl_Cre_MtoPre"
            Tab(0).Control(9).Enabled=   0   'False
            Tab(0).Control(10)=   "pnl_Cre_PerGra"
            Tab(0).Control(10).Enabled=   0   'False
            Tab(0).Control(11)=   "pnl_Cre_NumCuo"
            Tab(0).Control(11).Enabled=   0   'False
            Tab(0).Control(12)=   "pnl_Cre_MtoMPr"
            Tab(0).Control(12).Enabled=   0   'False
            Tab(0).Control(13)=   "pnl_Cre_MtoSol"
            Tab(0).Control(13).Enabled=   0   'False
            Tab(0).Control(14)=   "pnl_Cre_MtoDol"
            Tab(0).Control(14).Enabled=   0   'False
            Tab(0).Control(15)=   "pnl_Cre_ApoPro"
            Tab(0).Control(15).Enabled=   0   'False
            Tab(0).Control(16)=   "pnl_Cre_TipMon"
            Tab(0).Control(16).Enabled=   0   'False
            Tab(0).Control(17)=   "pnl_Cre_ComVta"
            Tab(0).Control(17).Enabled=   0   'False
            Tab(0).ControlCount=   18
            TabCaption(1)   =   "Datos Inmueble"
            TabPicture(1)   =   "OpeTra_frm_007.frx":0028
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "lbl_NomGlo(219)"
            Tab(1).Control(1)=   "lbl_NomGlo(189)"
            Tab(1).Control(2)=   "pnl_Inm_Propie"
            Tab(1).Control(3)=   "pnl_Inm_Direcc"
            Tab(1).ControlCount=   4
            TabCaption(2)   =   "Datos Legales"
            TabPicture(2)   =   "OpeTra_frm_007.frx":0044
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "lbl_NomGlo(107)"
            Tab(2).Control(1)=   "txt_Leg_InfLeg"
            Tab(2).ControlCount=   2
            TabCaption(3)   =   "Autorización Desembolso"
            TabPicture(3)   =   "OpeTra_frm_007.frx":0060
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "lbl_NomGlo(216)"
            Tab(3).Control(1)=   "lbl_NomGlo(217)"
            Tab(3).Control(2)=   "lbl_NomGlo(218)"
            Tab(3).Control(3)=   "lbl_NomGlo(220)"
            Tab(3).Control(4)=   "pnl_Aut_FecDes"
            Tab(3).Control(5)=   "pnl_Aut_BonoBP"
            Tab(3).Control(6)=   "pnl_Aut_FueFin"
            Tab(3).Control(7)=   "txt_Aut_Observ"
            Tab(3).ControlCount=   8
            Begin VB.TextBox txt_Aut_Observ 
               Height          =   645
               Left            =   -73200
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   101
               Text            =   "OpeTra_frm_007.frx":007C
               Top             =   1380
               Width           =   10755
            End
            Begin VB.TextBox txt_Leg_InfLeg 
               Height          =   1635
               Left            =   -73200
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   99
               Text            =   "OpeTra_frm_007.frx":0080
               Top             =   390
               Width           =   10755
            End
            Begin Threed.SSPanel pnl_Cre_ComVta 
               Height          =   315
               Left            =   1800
               TabIndex        =   79
               Top             =   720
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "20,000.00 "
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Cre_TipMon 
               Height          =   315
               Left            =   1800
               TabIndex        =   80
               Top             =   390
               Width           =   3855
               _Version        =   65536
               _ExtentX        =   6800
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "NUEVOS SOLES"
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
            Begin Threed.SSPanel pnl_Cre_ApoPro 
               Height          =   315
               Left            =   6030
               TabIndex        =   81
               Top             =   720
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "20,000.00 "
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Cre_MtoDol 
               Height          =   315
               Left            =   1800
               TabIndex        =   85
               Top             =   1050
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "20,000.00 "
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Cre_MtoSol 
               Height          =   315
               Left            =   6030
               TabIndex        =   86
               Top             =   1050
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "20,000.00 "
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Cre_MtoMPr 
               Height          =   315
               Left            =   10170
               TabIndex        =   87
               Top             =   1050
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "20,000.00 "
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Cre_NumCuo 
               Height          =   315
               Left            =   1800
               TabIndex        =   91
               Top             =   1380
               Width           =   735
               _Version        =   65536
               _ExtentX        =   1296
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "20,000.00 "
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Cre_PerGra 
               Height          =   315
               Left            =   6030
               TabIndex        =   93
               Top             =   1380
               Width           =   675
               _Version        =   65536
               _ExtentX        =   1191
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "20,000.00 "
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
               Alignment       =   4
            End
            Begin Threed.SSPanel pnl_Inm_Direcc 
               Height          =   615
               Left            =   -73200
               TabIndex        =   95
               Top             =   390
               Width           =   10695
               _Version        =   65536
               _ExtentX        =   18865
               _ExtentY        =   1085
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
               Alignment       =   0
            End
            Begin Threed.SSPanel pnl_Inm_Propie 
               Height          =   885
               Left            =   -73200
               TabIndex        =   97
               Top             =   1020
               Width           =   10695
               _Version        =   65536
               _ExtentX        =   18865
               _ExtentY        =   1561
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
               Alignment       =   0
            End
            Begin Threed.SSPanel pnl_Aut_FueFin 
               Height          =   315
               Left            =   -73200
               TabIndex        =   102
               Top             =   390
               Width           =   4455
               _Version        =   65536
               _ExtentX        =   7858
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "IKEHARA PUNK MIGUEL ANGEL"
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
            Begin Threed.SSPanel pnl_Aut_BonoBP 
               Height          =   315
               Left            =   -73200
               TabIndex        =   103
               Top             =   720
               Width           =   1485
               _Version        =   65536
               _ExtentX        =   2619
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "IKEHARA PUNK MIGUEL ANGEL"
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
            Begin Threed.SSPanel pnl_Aut_FecDes 
               Height          =   315
               Left            =   -73200
               TabIndex        =   104
               Top             =   1050
               Width           =   1485
               _Version        =   65536
               _ExtentX        =   2619
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "IKEHARA PUNK MIGUEL ANGEL"
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
            Begin Threed.SSPanel pnl_Cre_MtoPre 
               Height          =   315
               Left            =   10170
               TabIndex        =   109
               Top             =   390
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "20,000.00 "
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
               Alignment       =   4
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Monto Préstamo:"
               Height          =   315
               Index           =   21
               Left            =   8490
               TabIndex        =   110
               Top             =   390
               Width           =   1635
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Comentarios de Autorización:"
               Height          =   555
               Index           =   220
               Left            =   -74880
               TabIndex        =   108
               Top             =   1380
               Width           =   1305
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Fuente Financ.:"
               Height          =   285
               Index           =   218
               Left            =   -74880
               TabIndex        =   107
               Top             =   390
               Width           =   1515
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Flag Bono Buen Pag.:"
               Height          =   285
               Index           =   217
               Left            =   -74880
               TabIndex        =   106
               Top             =   720
               Width           =   1575
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Fecha Valor Desemb.:"
               Height          =   285
               Index           =   216
               Left            =   -74880
               TabIndex        =   105
               Top             =   1050
               Width           =   1575
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Informe Legal:"
               Height          =   315
               Index           =   107
               Left            =   -74880
               TabIndex        =   100
               Top             =   390
               Width           =   1545
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Propietario:"
               Height          =   285
               Index           =   189
               Left            =   -74880
               TabIndex        =   98
               Top             =   1020
               Width           =   1305
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Dirección:"
               Height          =   285
               Index           =   219
               Left            =   -74880
               TabIndex        =   96
               Top             =   390
               Width           =   1305
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Período de Gracia:"
               Height          =   315
               Index           =   34
               Left            =   4350
               TabIndex        =   94
               Top             =   1380
               Width           =   1545
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Número de Cuotas:"
               Height          =   315
               Index           =   32
               Left            =   120
               TabIndex        =   92
               Top             =   1380
               Width           =   1545
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Monto Aprob. US$:"
               Height          =   315
               Index           =   31
               Left            =   120
               TabIndex        =   90
               Top             =   1050
               Width           =   1635
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Monto Aprob. MPr.:"
               Height          =   315
               Index           =   30
               Left            =   8490
               TabIndex        =   89
               Top             =   1050
               Width           =   1635
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Monto Aprob. S/.:"
               Height          =   315
               Index           =   29
               Left            =   4350
               TabIndex        =   88
               Top             =   1050
               Width           =   1635
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Aporte Propio US$:"
               Height          =   315
               Index           =   27
               Left            =   4350
               TabIndex        =   84
               Top             =   720
               Width           =   1635
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "V. Compra-Venta US$:"
               Height          =   315
               Index           =   26
               Left            =   120
               TabIndex        =   83
               Top             =   720
               Width           =   1635
            End
            Begin VB.Label lbl_NomGlo 
               Caption         =   "Moneda Préstamo:"
               Height          =   315
               Index           =   25
               Left            =   120
               TabIndex        =   82
               Top             =   390
               Width           =   1545
            End
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   765
         Left            =   30
         TabIndex        =   72
         Top             =   9720
         Width           =   12735
         _Version        =   65536
         _ExtentX        =   22463
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
         Begin VB.CommandButton cmd_Cancel 
            Height          =   675
            Left            =   12000
            Picture         =   "OpeTra_frm_007.frx":0084
            Style           =   1  'Graphical
            TabIndex        =   74
            ToolTipText     =   "Cancelar"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   675
            Left            =   11310
            Picture         =   "OpeTra_frm_007.frx":038E
            Style           =   1  'Graphical
            TabIndex        =   73
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel12 
         Height          =   2835
         Left            =   30
         TabIndex        =   42
         Top             =   6840
         Width           =   12735
         _Version        =   65536
         _ExtentX        =   22463
         _ExtentY        =   5001
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
         Begin VB.ComboBox cmb_CtaCar 
            Height          =   315
            Left            =   6150
            Style           =   2  'Dropdown List
            TabIndex        =   64
            Top             =   1050
            Width           =   2385
         End
         Begin VB.TextBox txt_Observ 
            Height          =   495
            Left            =   1620
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   20
            Text            =   "OpeTra_frm_007.frx":07D0
            Top             =   2280
            Width           =   11055
         End
         Begin VB.TextBox txt_Garant 
            Height          =   555
            Left            =   1620
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   19
            Text            =   "OpeTra_frm_007.frx":07D6
            Top             =   1710
            Width           =   11055
         End
         Begin VB.TextBox txt_TraAbo 
            Height          =   315
            Left            =   10290
            MaxLength       =   25
            TabIndex        =   18
            Text            =   "Text1"
            Top             =   1380
            Width           =   2385
         End
         Begin VB.ComboBox cmb_BanAbo 
            Height          =   315
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   1380
            Width           =   2775
         End
         Begin VB.TextBox txt_CtaAbo 
            Height          =   315
            Left            =   6150
            MaxLength       =   25
            TabIndex        =   17
            Text            =   "Text1"
            Top             =   1380
            Width           =   2385
         End
         Begin VB.TextBox txt_CheCar 
            Height          =   315
            Left            =   10290
            MaxLength       =   25
            TabIndex        =   15
            Text            =   "Text1"
            Top             =   1050
            Width           =   2385
         End
         Begin VB.ComboBox cmb_BanCar 
            Height          =   315
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   1050
            Width           =   2775
         End
         Begin VB.ComboBox cmb_TipDes 
            Height          =   315
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   720
            Width           =   2775
         End
         Begin VB.ComboBox cmb_MonDes 
            Height          =   315
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   60
            Width           =   2775
         End
         Begin EditLib.fpDoubleSingle ipp_MtoDes 
            Height          =   315
            Left            =   1620
            TabIndex        =   12
            Top             =   390
            Width           =   1635
            _Version        =   196608
            _ExtentX        =   2893
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            ThreeDInsideStyle=   1
            ThreeDInsideHighlightColor=   -2147483637
            ThreeDInsideShadowColor=   -2147483642
            ThreeDInsideWidth=   1
            ThreeDOutsideStyle=   1
            ThreeDOutsideHighlightColor=   -2147483628
            ThreeDOutsideShadowColor=   -2147483632
            ThreeDOutsideWidth=   1
            ThreeDFrameWidth=   0
            BorderStyle     =   0
            BorderColor     =   -2147483642
            BorderWidth     =   1
            ButtonDisable   =   0   'False
            ButtonHide      =   0   'False
            ButtonIncrement =   1
            ButtonMin       =   0
            ButtonMax       =   100
            ButtonStyle     =   0
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   2
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            Text            =   "0.00"
            DecimalPlaces   =   2
            DecimalPoint    =   "."
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "9000000000"
            MinValue        =   "-9000000000"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ","
            UseSeparator    =   -1  'True
            IncInt          =   1
            IncDec          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin Threed.SSPanel pnl_DesMPr 
            Height          =   315
            Left            =   10290
            TabIndex        =   56
            Top             =   390
            Width           =   1785
            _Version        =   65536
            _ExtentX        =   3149
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_ImpITF 
            Height          =   315
            Left            =   6150
            TabIndex        =   58
            Top             =   390
            Visible         =   0   'False
            Width           =   1785
            _Version        =   65536
            _ExtentX        =   3149
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_TCaMPr 
            Height          =   315
            Left            =   10290
            TabIndex        =   60
            Top             =   60
            Width           =   1785
            _Version        =   65536
            _ExtentX        =   3149
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_TCaDol 
            Height          =   315
            Left            =   6150
            TabIndex        =   61
            Top             =   60
            Width           =   1785
            _Version        =   65536
            _ExtentX        =   3149
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
            Alignment       =   4
         End
         Begin VB.Label lbl_NomGlo 
            Caption         =   "T/C Mon. Prest.:"
            Height          =   285
            Index           =   14
            Left            =   8790
            TabIndex        =   63
            Top             =   60
            Width           =   1455
         End
         Begin VB.Label lbl_NomGlo 
            Caption         =   "T/C US$:"
            Height          =   285
            Index           =   13
            Left            =   4620
            TabIndex        =   62
            Top             =   60
            Width           =   1275
         End
         Begin VB.Label lbl_NomGlo 
            Caption         =   "Importe ITF:"
            Height          =   285
            Index           =   12
            Left            =   4620
            TabIndex        =   59
            Top             =   390
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.Label lbl_NomGlo 
            Caption         =   "Imp. Desemb. MPr.:"
            Height          =   285
            Index           =   15
            Left            =   8790
            TabIndex        =   57
            Top             =   390
            Width           =   1455
         End
         Begin VB.Label lbl_NomGlo 
            Caption         =   "Observaciones:"
            Height          =   285
            Index           =   10
            Left            =   90
            TabIndex        =   55
            Top             =   2280
            Width           =   1335
         End
         Begin VB.Label lbl_NomGlo 
            Caption         =   "Garantía:"
            Height          =   285
            Index           =   9
            Left            =   90
            TabIndex        =   52
            Top             =   1710
            Width           =   1335
         End
         Begin VB.Label lbl_NomGlo 
            Caption         =   "Nro. Transferencia:"
            Height          =   285
            Index           =   17
            Left            =   8790
            TabIndex        =   51
            Top             =   1380
            Width           =   1485
         End
         Begin VB.Label lbl_NomGlo 
            Caption         =   "Banco de Abono:"
            Height          =   315
            Index           =   8
            Left            =   90
            TabIndex        =   50
            Top             =   1380
            Width           =   1455
         End
         Begin VB.Label Label10 
            Caption         =   "Nro. Cuenta Abono:"
            Height          =   285
            Left            =   4620
            TabIndex        =   49
            Top             =   1380
            Width           =   1485
         End
         Begin VB.Label lbl_NomGlo 
            Caption         =   "Nro. Cheque Cargo:"
            Height          =   285
            Index           =   16
            Left            =   8790
            TabIndex        =   48
            Top             =   1050
            Width           =   1485
         End
         Begin VB.Label lbl_NomGlo 
            Caption         =   "Banco de Cargo:"
            Height          =   315
            Index           =   7
            Left            =   90
            TabIndex        =   47
            Top             =   1050
            Width           =   1455
         End
         Begin VB.Label lbl_NomGlo 
            Caption         =   "Nro. Cuenta Cargo:"
            Height          =   285
            Index           =   11
            Left            =   4620
            TabIndex        =   46
            Top             =   1050
            Width           =   1485
         End
         Begin VB.Label lbl_NomGlo 
            Caption         =   "Tipo de Desemb.:"
            Height          =   315
            Index           =   6
            Left            =   90
            TabIndex        =   45
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label lbl_NomGlo 
            Caption         =   "Moneda Desemb.:"
            Height          =   315
            Index           =   4
            Left            =   90
            TabIndex        =   44
            Top             =   60
            Width           =   1455
         End
         Begin VB.Label lbl_NomGlo 
            Caption         =   "Importe Desemb.:"
            Height          =   285
            Index           =   5
            Left            =   90
            TabIndex        =   43
            Top             =   390
            Width           =   1395
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   1815
         Left            =   30
         TabIndex        =   31
         Top             =   4980
         Width           =   12735
         _Version        =   65536
         _ExtentX        =   22463
         _ExtentY        =   3201
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
         Begin VB.CommandButton cmd_NueDes 
            Height          =   675
            Left            =   1440
            Picture         =   "OpeTra_frm_007.frx":07DC
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Nuevo Desembolso"
            Top             =   1080
            Width           =   675
         End
         Begin VB.CommandButton cmd_ImpLiq 
            Height          =   675
            Left            =   750
            Picture         =   "OpeTra_frm_007.frx":0AE6
            Style           =   1  'Graphical
            TabIndex        =   111
            ToolTipText     =   "Liquidación de Desembolso"
            Top             =   1080
            Width           =   675
         End
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   735
            Left            =   30
            TabIndex        =   7
            Top             =   330
            Width           =   12675
            _ExtentX        =   22357
            _ExtentY        =   1296
            _Version        =   393216
            Rows            =   21
            Cols            =   8
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin VB.CommandButton cmd_ImpCro 
            Height          =   675
            Left            =   60
            Picture         =   "OpeTra_frm_007.frx":0DF0
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Imprimir Cronogramas"
            Top             =   1080
            Width           =   675
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   675
            Left            =   2130
            Picture         =   "OpeTra_frm_007.frx":10FA
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Comprobante de Desembolso"
            Top             =   1080
            Width           =   675
         End
         Begin Threed.SSPanel pnl_TotDes 
            Height          =   315
            Left            =   10590
            TabIndex        =   32
            Top             =   1440
            Width           =   1785
            _Version        =   65536
            _ExtentX        =   3149
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   16777215
            BackColor       =   192
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
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
            Left            =   60
            TabIndex        =   33
            Top             =   60
            Width           =   1185
            _Version        =   65536
            _ExtentX        =   2090
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro. Desemb."
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
         Begin Threed.SSPanel pnl_SalPen 
            Height          =   315
            Left            =   10590
            TabIndex        =   34
            Top             =   1110
            Width           =   1785
            _Version        =   65536
            _ExtentX        =   3149
            _ExtentY        =   556
            _StockProps     =   15
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
            Left            =   1230
            TabIndex        =   35
            Top             =   60
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Desemb."
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
         Begin Threed.SSPanel SSPanel10 
            Height          =   285
            Left            =   2310
            TabIndex        =   36
            Top             =   60
            Width           =   2805
            _Version        =   65536
            _ExtentX        =   4948
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Tipo Desembolso"
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
         Begin Threed.SSPanel SSPanel18 
            Height          =   285
            Left            =   5100
            TabIndex        =   37
            Top             =   60
            Width           =   1575
            _Version        =   65536
            _ExtentX        =   2778
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Moneda Desemb."
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
         Begin Threed.SSPanel SSPanel8 
            Height          =   285
            Index           =   0
            Left            =   6660
            TabIndex        =   40
            Top             =   60
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Imp. Desemb."
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
         Begin Threed.SSPanel SSPanel8 
            Height          =   285
            Index           =   1
            Left            =   10590
            TabIndex        =   41
            Top             =   60
            Width           =   1785
            _Version        =   65536
            _ExtentX        =   3149
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Imp. Desemb. (M.Pr)"
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
         Begin Threed.SSPanel SSPanel8 
            Height          =   285
            Index           =   2
            Left            =   7980
            TabIndex        =   53
            Top             =   60
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "ITF"
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
         Begin Threed.SSPanel SSPanel8 
            Height          =   285
            Index           =   3
            Left            =   9300
            TabIndex        =   54
            Top             =   60
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Neto Desemb."
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
         Begin VB.Label lbl_NomGlo 
            Caption         =   "Saldo Pend. Desemb.:"
            Height          =   285
            Index           =   18
            Left            =   8910
            TabIndex        =   39
            Top             =   1110
            Width           =   1635
         End
         Begin VB.Label lbl_NomGlo 
            Caption         =   "Total Desemb."
            Height          =   285
            Index           =   19
            Left            =   8910
            TabIndex        =   38
            Top             =   1410
            Width           =   1215
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   22
         Top             =   30
         Width           =   12735
         _Version        =   65536
         _ExtentX        =   22463
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
            Height          =   495
            Left            =   630
            TabIndex        =   23
            Top             =   60
            Width           =   4905
            _Version        =   65536
            _ExtentX        =   8652
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Desembolso de Crédito Hipotecario"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
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
         Begin Threed.SSPanel pnl_Client 
            Height          =   405
            Left            =   4920
            TabIndex        =   24
            Top             =   120
            Width           =   7755
            _Version        =   65536
            _ExtentX        =   13679
            _ExtentY        =   714
            _StockProps     =   15
            Caption         =   "DNI - 07521154 / IKEHARA PUNK MIGUEL ANGEL "
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
            BevelOuter      =   0
            Font3D          =   2
            Alignment       =   4
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   60
            Picture         =   "OpeTra_frm_007.frx":1404
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel20 
         Height          =   795
         Left            =   30
         TabIndex        =   25
         Top             =   750
         Width           =   12735
         _Version        =   65536
         _ExtentX        =   22463
         _ExtentY        =   1402
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
         Begin VB.ComboBox cmb_TipDoc 
            Height          =   315
            Left            =   6210
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   60
            Width           =   2775
         End
         Begin VB.TextBox txt_NumDoc 
            Height          =   315
            Left            =   6210
            MaxLength       =   12
            TabIndex        =   2
            Text            =   "Text1"
            Top             =   390
            Width           =   2775
         End
         Begin VB.ComboBox cmb_TipBus 
            Height          =   315
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   2775
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   12000
            Picture         =   "OpeTra_frm_007.frx":170E
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Salir"
            Top             =   60
            Width           =   675
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   675
            Left            =   11280
            Picture         =   "OpeTra_frm_007.frx":1B50
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Limpiar Datos"
            Top             =   60
            Width           =   675
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   675
            Left            =   10560
            Picture         =   "OpeTra_frm_007.frx":1E5A
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Buscar Registros"
            Top             =   60
            Width           =   675
         End
         Begin MSMask.MaskEdBox msk_NumOpe 
            Height          =   315
            Left            =   1620
            TabIndex        =   3
            Top             =   390
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   12
            Mask            =   "###-##-#####"
            PromptChar      =   " "
         End
         Begin VB.Label lbl_NomGlo 
            Caption         =   "Nro. Operación:"
            Height          =   285
            Index           =   1
            Left            =   90
            TabIndex        =   30
            Top             =   390
            Width           =   1335
         End
         Begin VB.Label lbl_NomGlo 
            Caption         =   "Tipo Doc. Ident.:"
            Height          =   315
            Index           =   2
            Left            =   4830
            TabIndex        =   29
            Top             =   60
            Width           =   1395
         End
         Begin VB.Label lbl_NomGlo 
            Caption         =   "Nro. Doc. Ident.:"
            Height          =   285
            Index           =   3
            Left            =   4830
            TabIndex        =   28
            Top             =   390
            Width           =   1335
         End
         Begin VB.Label lbl_NomGlo 
            Caption         =   "Tipo de Búsqueda:"
            Height          =   315
            Index           =   0
            Left            =   90
            TabIndex        =   27
            Top             =   60
            Width           =   1455
         End
         Begin VB.Label Label20 
            Caption         =   "Nro. Doc. Id.:"
            Height          =   285
            Left            =   60
            TabIndex        =   26
            Top             =   1740
            Width           =   1065
         End
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   1095
         Index           =   4
         Left            =   30
         TabIndex        =   65
         Top             =   1590
         Width           =   12735
         _Version        =   65536
         _ExtentX        =   22463
         _ExtentY        =   1931
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
            Left            =   1620
            TabIndex        =   66
            Top             =   60
            Width           =   1725
            _Version        =   65536
            _ExtentX        =   3043
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "001-001-04-0001"
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
         Begin Threed.SSPanel pnl_Modali 
            Height          =   315
            Left            =   1620
            TabIndex        =   67
            Top             =   720
            Width           =   3825
            _Version        =   65536
            _ExtentX        =   6747
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "BIEN TERMINADO"
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
         Begin Threed.SSPanel pnl_Produc 
            Height          =   315
            Left            =   1620
            TabIndex        =   68
            Top             =   390
            Width           =   3825
            _Version        =   65536
            _ExtentX        =   6747
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "CREDITO HIPOTECARIO - MIVIVIENDA"
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
         Begin Threed.SSPanel pnl_NumSol 
            Height          =   315
            Left            =   8760
            TabIndex        =   76
            Top             =   60
            Width           =   1725
            _Version        =   65536
            _ExtentX        =   3043
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "001-001-04-0001"
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
         Begin VB.Label lbl_NomGlo 
            Caption         =   "Nro. Solicitud"
            Height          =   315
            Index           =   20
            Left            =   7110
            TabIndex        =   77
            Top             =   60
            Width           =   1335
         End
         Begin VB.Label lbl_NomGlo 
            Caption         =   "Nro. Operación:"
            Height          =   315
            Index           =   184
            Left            =   60
            TabIndex        =   71
            Top             =   60
            Width           =   1335
         End
         Begin VB.Label lbl_NomGlo 
            Caption         =   "Modalidad:"
            Height          =   315
            Index           =   187
            Left            =   60
            TabIndex        =   70
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label lbl_NomGlo 
            Caption         =   "Producto:"
            Height          =   315
            Index           =   188
            Left            =   60
            TabIndex        =   69
            Top             =   390
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "frm_Des_CreHip_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_BanCar()      As moddat_tpo_Genera
Dim l_arr_BanAbo()      As moddat_tpo_Genera
Dim l_arr_CtaCar()      As moddat_tpo_Genera
Dim l_dbl_MtoPre        As Double
Dim l_dbl_MtoDes        As Double
Dim l_dbl_TCaDol        As Double
Dim l_dbl_TCaMPr        As Double
Dim l_dbl_PorITF        As Double

Private Sub cmb_BanAbo_Click()
   Call gs_SetFocus(txt_CtaAbo)
End Sub

Private Sub cmb_BanAbo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_BanAbo_Click
   End If
End Sub

Private Sub cmb_BanCar_Click()
   If cmb_TipDes.ListIndex > -1 Then
      If cmb_TipDes.ItemData(cmb_TipDes.ListIndex) = 1 Then
         Call gs_SetFocus(txt_CheCar)
      Else
         Screen.MousePointer = 11
         Call moddat_gs_Carga_CtaBan(l_arr_BanCar(cmb_BanCar.ListIndex + 1).Genera_Codigo, cmb_CtaCar, l_arr_CtaCar)
         Screen.MousePointer = 0
         Call gs_SetFocus(cmb_CtaCar)
      End If
   End If
End Sub

Private Sub cmb_BanCar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_BanCar_Click
   End If
End Sub

Private Sub cmb_CtaCar_Click()
   If cmb_TipDes.ListIndex > -1 Then
      If cmb_TipDes.ItemData(cmb_TipDes.ListIndex) = 1 Then
         Call gs_SetFocus(txt_Garant)
      Else
         Call gs_SetFocus(cmb_BanAbo)
      End If
   End If
End Sub

Private Sub cmb_CtaCar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_CtaCar_Click
   End If
End Sub

Private Sub cmb_MonDes_Click()
   Call gs_SetFocus(ipp_MtoDes)
   
   pnl_DesMPr.Caption = "0.00 "
   If CDbl(ipp_MtoDes.Text) > 0 Then
      If cmb_MonDes.ListIndex > -1 Then
         If cmb_MonDes.ItemData(cmb_MonDes.ListIndex) = 1 Then
            'pnl_ImpITF.Caption = gf_Truncar_Numero(CDbl(ipp_MtoDes.Text) * (l_dbl_PorITF / 100), 2) & " "
            'pnl_DesMPr.Caption = Format((CDbl(ipp_MtoDes.Text) - CDbl(pnl_ImpITF.Caption)) / l_dbl_TCaMPr, "###,###,##0.00") & " "
            
            pnl_DesMPr.Caption = Format(CDbl(ipp_MtoDes.Text) / l_dbl_TCaMPr, "###,###,##0.00") & " "
         ElseIf cmb_MonDes.ItemData(cmb_MonDes.ListIndex) = 2 Then
            'pnl_ImpITF.Caption = gf_Truncar_Numero(CDbl(ipp_MtoDes.Text) * (l_dbl_PorITF / 100), 2) & " "
            'pnl_DesMPr.Caption = Format((CDbl(ipp_MtoDes.Text) - CDbl(pnl_ImpITF.Caption)) * l_dbl_TCaDol / l_dbl_TCaMPr, "###,###,##0.00") & " "
            
            pnl_DesMPr.Caption = Format(CDbl(ipp_MtoDes.Text) * l_dbl_TCaDol / l_dbl_TCaMPr, "###,###,##0.00") & " "
         End If
      End If
   End If
End Sub

Private Sub cmb_MonDes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_MonDes_Click
   End If
End Sub

Private Sub cmb_TipBus_Click()
   If cmb_TipBus.ListIndex > -1 Then
      If cmb_TipBus.ItemData(cmb_TipBus.ListIndex) = 1 Then
         cmb_TipDoc.Enabled = True
         txt_NumDoc.Enabled = True
         msk_NumOpe.Enabled = False
         
         msk_NumOpe.Mask = ""
         msk_NumOpe.Text = ""
         msk_NumOpe.Mask = "###-##-#####"
         
         Call gs_SetFocus(cmb_TipDoc)
      Else
         cmb_TipDoc.Enabled = False
         txt_NumDoc.Enabled = False
         msk_NumOpe.Enabled = True
         
         cmb_TipDoc.ListIndex = -1
         txt_NumDoc.Text = ""
         
         Call gs_SetFocus(msk_NumOpe)
      End If
   Else
      cmb_TipDoc.Enabled = False
      txt_NumDoc.Enabled = False
      
      msk_NumOpe.Enabled = False
   
      cmb_TipDoc.ListIndex = -1
      txt_NumDoc.Text = ""
      msk_NumOpe.Mask = ""
      msk_NumOpe.Text = ""
      msk_NumOpe.Mask = "###-##-#####"
   End If
End Sub

Private Sub cmb_TipBus_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipBus_Click
   End If
End Sub

Private Sub cmb_TipDes_Click()
   Call gs_SetFocus(cmb_BanCar)
   
   If cmb_TipDes.ListIndex > -1 Then
      If cmb_TipDes.ItemData(cmb_TipDes.ListIndex) = 1 Then
         txt_CheCar.Enabled = True
         
         cmb_CtaCar.Enabled = False
         cmb_BanAbo.Enabled = False
         txt_CtaAbo.Enabled = False
         txt_TraAbo.Enabled = False
         
         cmb_CtaCar.Clear
         cmb_BanAbo.ListIndex = -1
         txt_CtaAbo.Text = ""
         txt_TraAbo.Text = ""
      Else
         txt_CheCar.Enabled = False
         
         cmb_CtaCar.Enabled = True
         cmb_BanAbo.Enabled = True
         txt_CtaAbo.Enabled = True
         txt_TraAbo.Enabled = True
         
         txt_CheCar.Text = ""
      End If
   End If
End Sub

Private Sub cmb_TipDes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipDes_Click
   End If
End Sub

Private Sub cmb_TipDoc_Click()
   If cmb_TipDoc.ListIndex > -1 Then
      Select Case cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
         Case 1:  txt_NumDoc.MaxLength = 8
         Case 2:  txt_NumDoc.MaxLength = 12
         Case 3:  txt_NumDoc.MaxLength = 12
      End Select
   End If
   Call gs_SetFocus(txt_NumDoc)
End Sub

Private Sub cmb_TipDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipDoc_Click
   End If
End Sub

Private Sub cmd_Buscar_Click()
   If cmb_TipBus.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Búsqueda.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipBus)
      Exit Sub
   End If
   
   If cmb_TipBus.ItemData(cmb_TipBus.ListIndex) = 1 Then
      If cmb_TipDoc.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Tipo de Documento de Identidad.4", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_TipDoc)
         Exit Sub
      End If
      
      If Len(Trim(txt_NumDoc.Text)) = 0 Then
         MsgBox "Debe ingresar el Número de Documento de Identidad.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NumDoc)
         Exit Sub
      End If
      
      If cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) = 1 Then
         txt_NumDoc.Text = Format(txt_NumDoc.Text, "00000000")
      End If
      
      moddat_g_int_TipDoc = cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
      moddat_g_str_TipDoc = cmb_TipDoc.Text
      moddat_g_str_NumDoc = txt_NumDoc.Text
   Else
      If Len(Trim(msk_NumOpe.Text)) < 10 Then
         MsgBox "Debe ingresar el Número de Operación.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(msk_NumOpe)
         Exit Sub
      End If
      
      moddat_g_str_NumOpe = msk_NumOpe.Text
   End If
   
   If cmb_TipBus.ItemData(cmb_TipBus.ListIndex) = 1 Then
      g_str_Parame = "SELECT * FROM CRE_HIPMAE WHERE "
      g_str_Parame = g_str_Parame & "HIPMAE_TDOCLI = " & CStr(moddat_g_int_TipDoc) & " AND "
      g_str_Parame = g_str_Parame & "HIPMAE_NDOCLI = '" & moddat_g_str_NumDoc & "' AND "
      g_str_Parame = g_str_Parame & "HIPMAE_SITDES <> 3 AND "
      g_str_Parame = g_str_Parame & "(HIPMAE_SITUAC = 1 OR "
      g_str_Parame = g_str_Parame & "HIPMAE_SITUAC = 2) "
   Else
      g_str_Parame = "SELECT * FROM CRE_HIPMAE WHERE "
      g_str_Parame = g_str_Parame & "HIPMAE_NUMOPE = '" & moddat_g_str_NumOpe & "' AND "
      g_str_Parame = g_str_Parame & "HIPMAE_SITDES <> 3 AND "
      g_str_Parame = g_str_Parame & "(HIPMAE_SITUAC = 1 OR "
      g_str_Parame = g_str_Parame & "HIPMAE_SITUAC = 2) "
   End If
         
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Call cmd_Limpia_Click
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      MsgBox "No existe ninguna Operación para desembolsar para la Búsqueda deseada. ", vbExclamation, modgen_g_str_NomPlt
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Call cmd_Limpia_Click
      Exit Sub
   End If

   Call fs_Buscar_DatGen

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   
   Call fs_ActivaItem(False)
   Call fs_Activa(False)

   Call fs_Buscar_DatLeg
   Call fs_Buscar_DatInm
   Call fs_Buscar_DatAut
   
   Call fs_Buscar_Desemb
End Sub

Private Sub fs_Buscar_DatLeg()
   g_str_Parame = "SELECT * FROM TRA_EVALEG WHERE "
   g_str_Parame = g_str_Parame & "EVALEG_NUMSOL = '" & moddat_g_str_NumSol & "'"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   txt_Leg_InfLeg.Text = Trim(g_rst_Princi!EVALEG_INFLEG)
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Buscar_DatInm()
   Dim r_str_TipVia  As String
   Dim r_str_TipZon  As String
   Dim r_str_Depart  As String
   Dim r_str_Provin  As String
   Dim r_str_Distri  As String
   
   g_str_Parame = "SELECT * FROM CRE_SOLINM WHERE "
   g_str_Parame = g_str_Parame & "SOLINM_NUMSOL = '" & moddat_g_str_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "SOLINM_SITUAC = 1"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   r_str_TipVia = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!SOLINM_TIPVIA))
   r_str_TipZon = moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!SOLINM_TIPZON))

   pnl_Inm_Direcc.Caption = r_str_TipVia & " " & Trim(g_rst_Princi!SOLINM_NOMVIA) & " " & Trim(g_rst_Princi!SOLINM_NUMERO)
   
   If Len(Trim(Trim(g_rst_Princi!SOLINM_INTDPT))) > 0 Then
      pnl_Inm_Direcc.Caption = pnl_Inm_Direcc.Caption & " (" & Trim(g_rst_Princi!SOLINM_INTDPT) & ")"
   End If
   
   If Len(Trim(Trim(g_rst_Princi!SOLINM_NOMZON))) > 0 Then
      pnl_Inm_Direcc.Caption = pnl_Inm_Direcc.Caption & " - " & r_str_TipZon & " " & Trim(g_rst_Princi!SOLINM_NOMZON) & Chr(13) & Chr(10)
   Else
      pnl_Inm_Direcc.Caption = pnl_Inm_Direcc.Caption & Chr(13) & Chr(10)
   End If
   
   'Departamento
   r_str_Depart = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO, 2) & "0000")
   
   'Provincia
   r_str_Provin = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO, 4) & "00")
   
   'Distrito
   r_str_Distri = moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!SOLINM_UBIGEO))
   
   pnl_Inm_Direcc.Caption = pnl_Inm_Direcc.Caption & r_str_Distri & " - " & r_str_Provin & " - " & r_str_Depart
   
   pnl_Inm_Propie.Caption = moddat_gf_Consulta_ParDes("221", CStr(g_rst_Princi!SOLINM_TIPPER)) & Chr(13) & Chr(10)
   
   If g_rst_Princi!SOLINM_TIPPER = 2 Then
      'Persona Jurídica
      pnl_Inm_Propie.Caption = pnl_Inm_Propie & moddat_gf_Consulta_ParDes("203", CStr(g_rst_Princi!SOLINM_PROTDO)) & "-" & Trim(g_rst_Princi!SOLINM_PRONDO) & " / " & Trim(g_rst_Princi!SOLINM_PRORZS) & Chr(13) & Chr(10)
      
      r_str_TipVia = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!SOLINM_PROTVI))
      r_str_TipZon = moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!SOLINM_PROTZO))
   
      pnl_Inm_Propie.Caption = pnl_Inm_Propie & r_str_TipVia & " " & Trim(g_rst_Princi!SOLINM_PRONVI) & " " & Trim(g_rst_Princi!SOLINM_PRONUM)
      
      If Len(Trim(Trim(g_rst_Princi!SOLINM_PROINT))) > 0 Then
         pnl_Inm_Propie.Caption = pnl_Inm_Propie & " (" & Trim(g_rst_Princi!SOLINM_PROINT) & ")"
      End If
      
      If Len(Trim(Trim(g_rst_Princi!SOLINM_PRONZO))) > 0 Then
         pnl_Inm_Propie.Caption = pnl_Inm_Propie & " - " & r_str_TipZon & " " & Trim(g_rst_Princi!SOLINM_PRONZO) & Chr(13) & Chr(10)
      Else
         pnl_Inm_Propie.Caption = pnl_Inm_Propie & Chr(13) & Chr(10)
      End If
      
      r_str_Depart = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_PROUBI, 2) & "0000")
      r_str_Provin = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_PROUBI, 4) & "00")
      r_str_Distri = moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!SOLINM_PROUBI))
      
      pnl_Inm_Propie.Caption = pnl_Inm_Propie & r_str_Distri & " - " & r_str_Provin & " - " & r_str_Depart & Chr(13) & Chr(10)
      
      pnl_Inm_Propie.Caption = pnl_Inm_Propie & Trim(g_rst_Princi!SOLINM_PROAPP) & " " & Trim(g_rst_Princi!SOLINM_PROAPM) & " " & Trim(g_rst_Princi!SOLINM_PRONOM)
   Else
      'Persona Natural
      pnl_Inm_Propie.Caption = pnl_Inm_Propie & moddat_gf_Consulta_ParDes("203", CStr(g_rst_Princi!SOLINM_PROTDO)) & "-" & Trim(g_rst_Princi!SOLINM_PRONDO) & " / " & Trim(g_rst_Princi!SOLINM_PROAPP) & " " & Trim(g_rst_Princi!SOLINM_PROAPM) & " " & Trim(g_rst_Princi!SOLINM_PRONOM) & Chr(13) & Chr(10)
      
      If g_rst_Princi!SOLINM_CYGTDO > 0 Then
         pnl_Inm_Propie.Caption = pnl_Inm_Propie & moddat_gf_Consulta_ParDes("203", CStr(g_rst_Princi!SOLINM_CYGTDO)) & "-" & Trim(g_rst_Princi!SOLINM_CYGNDO) & " / " & Trim(g_rst_Princi!SOLINM_CYGAPP) & " " & Trim(g_rst_Princi!SOLINM_CYGAPM) & " " & Trim(g_rst_Princi!SOLINM_CYGNOM)
      End If
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Buscar_DatAut()
   g_str_Parame = "SELECT * FROM TRA_AUTDES WHERE "
   g_str_Parame = g_str_Parame & "AUTDES_NUMSOL = '" & moddat_g_str_NumSol & "'"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   pnl_Aut_FueFin.Caption = moddat_gf_Consulta_ParDes("502", g_rst_Princi!AUTDES_FUEFIN)
   pnl_Aut_BonoBP.Caption = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!AUTDES_BONOBP))
   pnl_Aut_FecDes.Caption = gf_FormatoFecha(CStr(g_rst_Princi!AUTDES_FECDES))
   txt_Aut_Observ.Text = Trim(g_rst_Princi!AUTDES_OBSERV & "")
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub cmd_Cancel_Click()
   Call fs_LimpiaItem
   Call fs_ActivaItem(False)
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub cmd_Grabar_Click()
   Dim r_str_Operac     As String
   
   If cmb_MonDes.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Moneda del Desembolso.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_MonDes)
      Exit Sub
   End If
   
   If CDbl(ipp_MtoDes.Text) = 0 Then
      MsgBox "Debe ingresar el Monto del Desembolso.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_MtoDes)
      Exit Sub
   End If
   
   If cmb_TipDes.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Tipo del Desembolso.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipDes)
      Exit Sub
   End If
   
   If cmb_BanCar.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Banco de Cargo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_BanCar)
      Exit Sub
   End If
   
   If cmb_TipDes.ItemData(cmb_TipDes.ListIndex) = 1 Then
      'Cheque de Gerencia
      
      r_str_Operac = moddat_gf_Consulta_Operac(moddat_g_str_CodPrd, "22")
      
      If Len(Trim(txt_CheCar.Text)) = 0 Then
         MsgBox "Debe ingresar el Número de Cheque de Cargo.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_CheCar)
         Exit Sub
      End If
   Else
      r_str_Operac = moddat_gf_Consulta_Operac(moddat_g_str_CodPrd, "21")
      
      If cmb_CtaCar.ListIndex = -1 Then
         MsgBox "Debe seleccionar la Cuenta de Cargo.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_CtaCar)
         Exit Sub
      End If
   
      If cmb_BanAbo.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Banco de Abono.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_BanAbo)
         Exit Sub
      End If
   
      If Len(Trim(txt_CtaAbo.Text)) = 0 Then
         MsgBox "Debe ingresar el Número de Cuenta de Abono.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_CtaAbo)
         Exit Sub
      End If
   
      If Len(Trim(txt_TraAbo.Text)) = 0 Then
         MsgBox "Debe ingresar el Número de Transferencia.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_TraAbo)
         Exit Sub
      End If
   End If
   
   If Len(Trim(txt_Garant.Text)) = 0 Then
      MsgBox "Debe ingresar la Descripción de la Garantía.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Garant)
      Exit Sub
   End If
   
   If CDbl(pnl_DesMPr.Caption) > CDbl(pnl_SalPen.Caption) Then
      MsgBox "El importe del desembolso no puede exceder el Saldo Pendiente.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_MtoDes)
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   r_str_Operac = CStr(moddat_g_int_TipMon) & Right(r_str_Operac, 5)
   
   'Actualizando Solicitud de Crédito
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_CRE_HIPDES ("
      
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "', "
      
      If cmb_CtaCar.ListIndex > -1 Then
         g_str_Parame = g_str_Parame & "'" & l_arr_CtaCar(cmb_CtaCar.ListIndex + 1).Genera_Codigo & "', "
      Else
         g_str_Parame = g_str_Parame & "'', "
      End If
      
      g_str_Parame = g_str_Parame & "'" & txt_CheCar.Text & "', "
      g_str_Parame = g_str_Parame & "'" & l_arr_BanCar(cmb_BanCar.ListIndex + 1).Genera_Codigo & "', "
      g_str_Parame = g_str_Parame & "'" & txt_CtaAbo.Text & "', "
      
      If cmb_BanAbo.ListIndex > -1 Then
         g_str_Parame = g_str_Parame & "'" & l_arr_BanCar(cmb_BanCar.ListIndex + 1).Genera_Codigo & "', "
      Else
         g_str_Parame = g_str_Parame & "'', "
      End If
      
      g_str_Parame = g_str_Parame & "'" & txt_TraAbo.Text & "', "
      g_str_Parame = g_str_Parame & CStr(cmb_TipDes.ItemData(cmb_TipDes.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & CStr(moddat_g_int_TipMon) & ", "
      'g_str_Parame = g_str_Parame & CStr(cmb_MonDes.ItemData(cmb_MonDes.ListIndex)) & ", "
      
      'g_str_Parame = g_str_Parame & CStr(CDbl(ipp_MtoDes.Text)) & ", "
      'g_str_Parame = g_str_Parame & CStr(CDbl(pnl_ImpITF.Caption)) & ", "
      'g_str_Parame = g_str_Parame & CStr(CDbl(pnl_DesMPr.Caption)) & ", "
      
      g_str_Parame = g_str_Parame & CStr(CDbl(ipp_MtoDes.Text)) & ", "
      g_str_Parame = g_str_Parame & CStr(0) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_DesMPr.Caption)) & ", "
      
      If cmb_MonDes.ItemData(cmb_MonDes.ListIndex) = 1 Then
         g_str_Parame = g_str_Parame & CStr(CDbl(ipp_MtoDes.Text) - CDbl(pnl_ImpITF.Caption)) & ", "
         g_str_Parame = g_str_Parame & Format((CDbl(ipp_MtoDes.Text) - CDbl(pnl_ImpITF.Caption)) / l_dbl_TCaDol, "########0.00") & ", "
      Else
         g_str_Parame = g_str_Parame & Format((CDbl(ipp_MtoDes.Text) - CDbl(pnl_ImpITF.Caption)) * l_dbl_TCaDol, "########0.00") & ", "
         g_str_Parame = g_str_Parame & CStr(CDbl(ipp_MtoDes.Text) - CDbl(pnl_ImpITF.Caption)) & ", "
      End If
      
      g_str_Parame = g_str_Parame & CStr(l_dbl_TCaDol) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_TCaMPr) & ", "
      g_str_Parame = g_str_Parame & "'" & txt_Garant.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_Observ.Text & "', "
      g_str_Parame = g_str_Parame & "'" & r_str_Operac & "', "
      
      
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                           'Código Usuario
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                           'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                            'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "                           'Código Sucursal
      g_str_Parame = g_str_Parame & "1)"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_CRE_HIPDES. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   Call fs_Buscar_Desemb
   Call cmd_Cancel_Click
End Sub

Private Sub cmd_ImpCro_Click()
   frm_Des_CreHip_01.Show 1
End Sub

Private Sub cmd_ImpLiq_Click()
   Dim r_str_Direcc     As String
   Dim r_str_Distri     As String
   
   
   'Obteniendo Información de la Operación
   g_str_Parame = "SELECT * FROM CRE_HIPMAE WHERE "
   g_str_Parame = g_str_Parame & "HIPMAE_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If

   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   
   'Obteniendo Información del Inmueble
   Call moddat_gs_Consulta_DatInm(moddat_g_str_NumSol, r_str_Direcc, r_str_Distri)

   'Inicializando Arreglo de Impresiones
   ReDim g_arr_Imprim(0)

   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp(Space(98) & "Fecha: " & moddat_g_str_FecSis)
   Call gs_LinImp(Space(98) & "Hora:  " & Space(2) & Format(Time, "hh:mm:ss"))
   Call gs_LinImp("")
   Call gs_LinImp(Space(36) & "LIQUIDACION DE DESEMBOLSO - CREDITO HIPOTECARIO")
   Call gs_LinImp(Space(36) & "-----------------------------------------------")
   Call gs_LinImp("")
   
   Call gs_LinImp(Space(5) & "Nro. de Operación     : " & Mid(moddat_g_str_NumOpe, 1, 3) & "-" & Mid(moddat_g_str_NumOpe, 4, 2) & "-" & Mid(moddat_g_str_NumOpe, 6, 5))
   
   Call gs_LinImp(Space(5) & "Docum. Ident. Cliente : " & CStr(g_rst_Princi!HIPMAE_TDOCLI) & "-" & Trim(g_rst_Princi!HIPMAE_NDOCLI))
   Call gs_LinImp(Space(5) & "Nombre Cliente        : " & moddat_gf_Buscar_NomCli(g_rst_Princi!HIPMAE_TDOCLI, Trim(g_rst_Princi!HIPMAE_NDOCLI)))
   Call gs_LinImp(Space(5) & String(110, "-"))
   Call gs_LinImp(Space(5) & "Dirección Inmueble    : " & r_str_Direcc)
   Call gs_LinImp(Space(5) & Space(24) & r_str_Distri)
   Call gs_LinImp(Space(5) & String(110, "-"))
   
   Call gs_LinImp(Space(5) & "Producto de Crédito   : " & moddat_g_str_NomPrd)
   Call gs_LinImp(Space(5) & "Modalidad de Crédito  : " & moddat_g_str_DesMod)
   Call gs_LinImp(Space(5) & String(110, "-"))
   Call gs_LinImp(Space(5) & "Moneda de Préstamo    : " & moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!HIPMAE_MONEDA)))
   Call gs_LinImp(Space(5) & "Total Préstamo        : " & gf_FormatoNumero(g_rst_Princi!HIPMAE_PREMPR, 15))
   Call gs_LinImp(Space(5) & "Bono Buen Pagador     : " & moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!HIPMAE_FLGBBP)))
   Call gs_LinImp(Space(5) & "Tramo No Concesional  : " & gf_FormatoNumero(g_rst_Princi!HIPMAE_IMPNCO, 15))
   Call gs_LinImp(Space(5) & "Tramo Concesional     : " & gf_FormatoNumero(g_rst_Princi!HIPMAE_IMPCON, 15))
   
   Call gs_LinImp(Space(5) & String(110, "-"))
   Call gs_LinImp(Space(5) & "Fecha Desembolso      : " & gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECAPR)))
   Call gs_LinImp(Space(5) & "Nro. Cuotas           : " & Format(g_rst_Princi!HIPMAE_NUMCUO, "000"))
   Call gs_LinImp(Space(5) & "Cuotas Extradordin.   : " & Mid(moddat_gf_Consulta_ParDes("223", g_rst_Princi!HIPMAE_CUOANO) & Space(20), 1, 18))
   Call gs_LinImp(Space(5) & "Período de Gracia     : " & Mid(Format(g_rst_Princi!HIPMAE_PERGRA, "#0") & Space(18), 1, 18))
   Call gs_LinImp(Space(5) & String(110, "-"))
   Call gs_LinImp(Space(5) & "Nro. de Solicitud     : " & Mid(moddat_g_str_NumSol, 1, 3) & "-" & Mid(moddat_g_str_NumSol, 4, 3) & "-" & Mid(moddat_g_str_NumSol, 7, 2) & "-" & Mid(moddat_g_str_NumSol, 9, 4))
   Call gs_LinImp(Space(5) & "Tipo de Seguro        : " & moddat_gf_Consulta_TipSeg(g_rst_Princi!HIPMAE_SEGPRE, g_rst_Princi!HIPMAE_TIPSEG))
   Call gs_LinImp(Space(5) & String(110, "-"))
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp(Space(5) & String(30, "-") & Space(50) & String(30, "-"))
   Call gs_LinImp(Space(5) & Space(11) & "CLIENTE" & Space(12) & Space(50) & Space(11) & "MICASITA")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp(Space(5) & "San Isidro, ____________ de ____________________________ del 20______")
   
   Screen.MousePointer = 0
   
   frm_Imprim_01.Show 1
End Sub

Private Sub cmd_Imprim_Click()
   Dim r_str_NumDes     As String
   Dim r_str_Direcc     As String
   Dim r_str_Distri     As String
   
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If

   grd_Listad.Col = 0
   r_str_NumDes = CStr(CInt(grd_Listad.Text))
   
   Call gs_RefrescaGrid(grd_Listad)

   'Obteniendo Información del Inmueble
   Call moddat_gs_Consulta_DatInm(moddat_g_str_NumSol, r_str_Direcc, r_str_Distri)

   'Inicializando Arreglo de Impresiones
   ReDim g_arr_Imprim(0)

   g_str_Parame = "SELECT * FROM CRE_HIPDES WHERE "
   g_str_Parame = g_str_Parame & "HIPDES_NUMOPE = '" & moddat_g_str_NumOpe & "' AND "
   g_str_Parame = g_str_Parame & "HIPDES_NUMDES = " & r_str_NumDes & " "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp(Space(98) & "Fecha: " & moddat_g_str_FecSis)
   Call gs_LinImp(Space(98) & "Hora:  " & Space(2) & Format(Time, "hh:mm:ss"))
   Call gs_LinImp("")
   Call gs_LinImp(Space(36) & "LIQUIDACION DE DESEMBOLSO - TRASPASO DE FONDOS")
   Call gs_LinImp(Space(36) & "----------------------------------------------")
   Call gs_LinImp("")
   
   Call gs_LinImp(Space(5) & "Nro. de Operación     : " & Mid(moddat_g_str_NumOpe, 1, 3) & "-" & Mid(moddat_g_str_NumOpe, 4, 2) & "-" & Mid(moddat_g_str_NumOpe, 6, 5))
   
   Call gs_LinImp(Space(5) & "Docum. Ident. Cliente : " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc)
   Call gs_LinImp(Space(5) & "Nombre Cliente        : " & moddat_gf_Buscar_NomCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc))
   Call gs_LinImp(Space(5) & String(110, "-"))
   Call gs_LinImp(Space(5) & "Dirección Inmueble    : " & r_str_Direcc)
   Call gs_LinImp(Space(5) & Space(24) & r_str_Distri)
   Call gs_LinImp(Space(5) & String(110, "-"))
   
   Call gs_LinImp(Space(5) & "Producto de Crédito   : " & moddat_g_str_NomPrd)
   Call gs_LinImp(Space(5) & "Modalidad de Crédito  : " & moddat_g_str_DesMod)
   Call gs_LinImp(Space(5) & String(110, "-"))
   
   Call gs_LinImp(Space(5) & "Número de Desembolso  : " & Format(g_rst_Princi!HIPDES_NUMDES, "000"))
   Call gs_LinImp(Space(5) & "Fecha de Desembolso   : " & gf_FormatoFecha(CStr(g_rst_Princi!HIPDES_FECDES)))
   Call gs_LinImp(Space(5) & "Tipo de Desembolso    : " & moddat_gf_Consulta_ParDes("226", g_rst_Princi!HIPDES_TIPDES))
   
   If g_rst_Princi!HIPDES_TIPDES = 1 Then
      Call gs_LinImp(Space(5) & "Banco Cheque Emitido  : " & moddat_gf_Consulta_ParDes("505", g_rst_Princi!HIPDES_BANCGO))
      Call gs_LinImp(Space(5) & "Cheque de Gerencia    : " & g_rst_Princi!HIPDES_CHECGO)
   Else
      Call gs_LinImp(Space(5) & "Banco Abono Transfer. : " & moddat_gf_Consulta_ParDes("505", g_rst_Princi!HIPDES_BANABO))
      Call gs_LinImp(Space(5) & "Nro. de Cuenta Transf.: " & g_rst_Princi!HIPDES_CTAABO)
      Call gs_LinImp(Space(5) & "Nro. de Transferencia : " & g_rst_Princi!HIPDES_NUMTRA)
   End If
   
   Call gs_LinImp(Space(5) & String(110, "-"))
   
   If g_rst_Princi!HIPDES_TCADOL = 0 Then
      Call gs_LinImp(Space(5) & "Tipo de Moneda        : " & moddat_gf_Consulta_ParDes("204", 1))
   ElseIf g_rst_Princi!HIPDES_TCADOL = g_rst_Princi!HIPDES_TCAMPR And g_rst_Princi!HIPDES_TCADOL > 0 And moddat_g_int_TipMon <> 3 Then
      Call gs_LinImp(Space(5) & "Tipo de Moneda        : " & moddat_gf_Consulta_ParDes("204", g_rst_Princi!HIPDES_TIPMON))
   Else
      Call gs_LinImp(Space(5) & "Tipo de Moneda        : " & moddat_gf_Consulta_ParDes("204", 1))
   End If
   
   Call gs_LinImp(Space(5) & "Importe Desembolsado  : " & gf_FormatoNumero(g_rst_Princi!HIPDES_IMPORT - g_rst_Princi!HIPDES_IMPITF, 15))
   
   
   Call gs_LinImp(Space(5) & String(110, "-"))
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp(Space(5) & String(30, "-"))
   Call gs_LinImp(Space(5) & Space(11) & "MICASITA")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp(Space(5) & String(30, "-"))
   Call gs_LinImp(Space(5) & "Vendedor/Constructor : ")
   Call gs_LinImp(Space(5) & "Documento Identidad  : ")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp(Space(5) & String(30, "-"))
   Call gs_LinImp(Space(5) & "Vendedor/Constructor : ")
   Call gs_LinImp(Space(5) & "Documento Identidad  : ")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp(Space(5) & "San Isidro, ____________ de ____________________________ del 20______")
   
   'Voucher para Contabilidad
   Call gs_LinImp("SP")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp(Space(98) & "Fecha: " & moddat_g_str_FecSis)
   Call gs_LinImp(Space(98) & "Hora:  " & Space(2) & Format(Time, "hh:mm:ss"))
   Call gs_LinImp("")
   Call gs_LinImp(Space(40) & "COMPROBANTE DE LIQUIDACION DE DESEMBOLSO")
   Call gs_LinImp(Space(40) & "----------------------------------------")
   Call gs_LinImp("")
   
   Call gs_LinImp(Space(5) & "Nro. de Operación     : " & Mid(moddat_g_str_NumOpe, 1, 3) & "-" & Mid(moddat_g_str_NumOpe, 4, 2) & "-" & Mid(moddat_g_str_NumOpe, 6, 5))
   
   Call gs_LinImp(Space(5) & "Docum. Ident. Cliente : " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc)
   Call gs_LinImp(Space(5) & "Nombre Cliente        : " & moddat_gf_Buscar_NomCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc))
   Call gs_LinImp(Space(5) & String(110, "-"))
   Call gs_LinImp(Space(5) & "Dirección Inmueble    : " & r_str_Direcc)
   Call gs_LinImp(Space(5) & Space(24) & r_str_Distri)
   Call gs_LinImp(Space(5) & String(110, "-"))
   
   Call gs_LinImp(Space(5) & "Producto de Crédito   : " & moddat_g_str_NomPrd)
   Call gs_LinImp(Space(5) & "Modalidad de Crédito  : " & moddat_g_str_DesMod)
   Call gs_LinImp(Space(5) & "Moneda de Préstamo    : " & moddat_gf_Consulta_ParDes("204", CStr(moddat_g_int_TipMon)))
   Call gs_LinImp(Space(5) & "Monto Desembolsado    : " & gf_FormatoNumero(l_dbl_MtoPre, 15))
   Call gs_LinImp(Space(5) & String(110, "-"))
   
   Call gs_LinImp(Space(5) & "Número de Desembolso  : " & Format(g_rst_Princi!HIPDES_NUMDES, "000"))
   Call gs_LinImp(Space(5) & "Fecha de Desembolso   : " & gf_FormatoFecha(CStr(g_rst_Princi!HIPDES_FECDES)))
   Call gs_LinImp(Space(5) & "Tipo de Desembolso    : " & moddat_gf_Consulta_ParDes("226", g_rst_Princi!HIPDES_TIPDES))
   
   If g_rst_Princi!HIPDES_TIPDES = 1 Then
      Call gs_LinImp(Space(5) & "Banco Cheque Emitido  : " & moddat_gf_Consulta_ParDes("505", g_rst_Princi!HIPDES_BANCGO))
      Call gs_LinImp(Space(5) & "Cheque de Gerencia    : " & g_rst_Princi!HIPDES_CHECGO)
   Else
      Call gs_LinImp(Space(5) & "Banco Abono Transfer. : " & moddat_gf_Consulta_ParDes("505", g_rst_Princi!HIPDES_BANABO))
      Call gs_LinImp(Space(5) & "Nro. de Cuenta Transf.: " & g_rst_Princi!HIPDES_CTAABO)
      Call gs_LinImp(Space(5) & "Nro. de Transferencia : " & g_rst_Princi!HIPDES_NUMTRA)
   End If
   
   Call gs_LinImp(Space(5) & String(110, "-"))
   
   If g_rst_Princi!HIPDES_TCADOL = 0 Then
      Call gs_LinImp(Space(5) & "Tipo de Moneda        : " & moddat_gf_Consulta_ParDes("204", 1))
   ElseIf g_rst_Princi!HIPDES_TCADOL = g_rst_Princi!HIPDES_TCAMPR And g_rst_Princi!HIPDES_TCADOL > 0 And moddat_g_int_TipMon <> 3 Then
      Call gs_LinImp(Space(5) & "Tipo de Moneda        : " & moddat_gf_Consulta_ParDes("204", g_rst_Princi!HIPDES_TIPMON))
   Else
      Call gs_LinImp(Space(5) & "Tipo de Moneda        : " & moddat_gf_Consulta_ParDes("204", 1))
   End If
   
   Call gs_LinImp(Space(5) & "Importe Desembolsado  : " & gf_FormatoNumero(g_rst_Princi!HIPDES_IMPORT - g_rst_Princi!HIPDES_IMPITF, 15))
   Call gs_LinImp(Space(5) & String(110, "-"))
   Call gs_LinImp(Space(5) & "Tipo Cambio (M. Prst.): " & gf_FormatoTipCam(g_rst_Princi!HIPDES_TCAMPR, 15, 6))
   Call gs_LinImp(Space(5) & "Tipo Cambio (US$)     : " & gf_FormatoTipCam(g_rst_Princi!HIPDES_TCADOL, 15, 6))
   Call gs_LinImp(Space(5) & "Importe en M. Prest.  : " & gf_FormatoNumero(g_rst_Princi!HIPDES_DESMPR, 15))
   
   
   Call gs_LinImp(Space(5) & String(110, "-"))
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp(Space(5) & String(30, "-"))
   Call gs_LinImp(Space(5) & Space(10) & "OPERACIONES")
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Screen.MousePointer = 0
   
   frm_Imprim_01.Show 1
End Sub

Private Sub cmd_Limpia_Click()
   Call fs_Limpia
   Call gs_SetFocus(cmb_TipBus)
End Sub

Private Sub cmd_NueDes_Click()
   l_dbl_TCaDol = moddat_gf_Obtiene_TipCam(1, 2)
   l_dbl_TCaMPr = moddat_gf_Obtiene_TipCam(1, moddat_g_int_TipMon)
   
   pnl_TCaDol.Caption = Format(l_dbl_TCaDol, "###,###,##0.000000") & " "
   pnl_TCaMPr.Caption = Format(l_dbl_TCaMPr, "###,###,##0.000000") & " "
   
   If l_dbl_TCaMPr = 0 Then
      MsgBox "No se ha encuentra registrado el Tipo de Cambio (MONEDA DEL PRESTAMO).", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If l_dbl_TCaDol = 0 Then
      MsgBox "No se ha encuentra registrado el Tipo de Cambio (DOLARES AMERICANOS).", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   Call fs_ActivaItem(True)
   Call gs_SetFocus(cmb_MonDes)
   
   cmb_CtaCar.Enabled = False
   txt_CheCar.Enabled = False
   cmb_BanAbo.Enabled = False
   txt_CtaAbo.Enabled = False
   txt_TraAbo.Enabled = False
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub ipp_MtoDes_Change()
   pnl_DesMPr.Caption = "0.00 "
   
   If CDbl(ipp_MtoDes.Text) > 0 Then
      If cmb_MonDes.ListIndex > -1 Then
         If cmb_MonDes.ItemData(cmb_MonDes.ListIndex) = 1 Then
            'pnl_ImpITF.Caption = gf_Truncar_Numero(CDbl(ipp_MtoDes.Text) * (l_dbl_PorITF / 100), 2) & " "
            'pnl_DesMPr.Caption = Format((CDbl(ipp_MtoDes.Text) - CDbl(pnl_ImpITF.Caption)) / l_dbl_TCaMPr, "###,###,##0.00") & " "
            
            pnl_DesMPr.Caption = Format(CDbl(ipp_MtoDes.Text) / l_dbl_TCaMPr, "###,###,##0.00") & " "
         ElseIf cmb_MonDes.ItemData(cmb_MonDes.ListIndex) = 2 Then
            'pnl_ImpITF.Caption = gf_Truncar_Numero(CDbl(ipp_MtoDes.Text) * (l_dbl_PorITF / 100), 2) & " "
            'pnl_DesMPr.Caption = Format((CDbl(ipp_MtoDes.Text) - CDbl(pnl_ImpITF.Caption)) * l_dbl_TCaDol / l_dbl_TCaMPr, "###,###,##0.00") & " "
            
            pnl_DesMPr.Caption = Format(CDbl(ipp_MtoDes.Text) * l_dbl_TCaDol / l_dbl_TCaMPr, "###,###,##0.00") & " "
         End If
      End If
   End If
End Sub

Private Sub ipp_MtoDes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipDes)
   End If
End Sub

Private Sub msk_NumOpe_GotFocus()
   Call gs_SelecTodo(msk_NumOpe)
End Sub

Private Sub msk_NumOpe_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Buscar)
   End If
End Sub

Private Sub txt_AutDes_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txt_CheCar_GotFocus()
   Call gs_SelecTodo(txt_CheCar)
End Sub

Private Sub txt_CheCar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Garant)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "-")
   End If
End Sub

Private Sub txt_CtaAbo_GotFocus()
   Call gs_SelecTodo(txt_CtaAbo)
End Sub

Private Sub txt_CtaAbo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_TraAbo)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "-")
   End If
End Sub

Private Sub txt_CtaCar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmb_TipDes.ListIndex > -1 Then
         If cmb_TipDes.ItemData(cmb_TipDes.ListIndex) = 1 Then
            Call gs_SetFocus(txt_Garant)
         Else
            Call gs_SetFocus(cmb_BanAbo)
         End If
      End If
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_InfLeg_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txt_NumDoc_GotFocus()
   Call gs_SelecTodo(txt_NumDoc)
End Sub

Private Sub txt_NumDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Buscar)
   Else
      If cmb_TipDoc.ListIndex > -1 Then
         Select Case cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
            Case 1:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
            Case 2:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-")
            Case 3:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-")
         End Select
      Else
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call cmd_Limpia_Click
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call modsis_gs_Carga_TipBus_1(cmb_TipBus)
   Call moddat_gs_Carga_TipDocIde(cmb_TipDoc, 1)
   
   Call moddat_gs_Carga_LisIte(cmb_BanCar, l_arr_BanCar, 1, "505")
   Call moddat_gs_Carga_LisIte(cmb_BanAbo, l_arr_BanAbo, 1, "505")
   Call moddat_gs_Carga_TipMon(cmb_MonDes, 1)

   Call moddat_gs_Carga_LisIte_Combo(cmb_TipDes, 1, "226")

   grd_Listad.ColWidth(0) = 1175
   grd_Listad.ColWidth(1) = 1085
   grd_Listad.ColWidth(2) = 2795
   grd_Listad.ColWidth(3) = 1565
   grd_Listad.ColWidth(4) = 1325
   grd_Listad.ColWidth(5) = 1325
   grd_Listad.ColWidth(6) = 1325
   grd_Listad.ColWidth(7) = 1775
   
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignLeftCenter
   grd_Listad.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad.ColAlignment(4) = flexAlignRightCenter
   grd_Listad.ColAlignment(5) = flexAlignRightCenter
   grd_Listad.ColAlignment(6) = flexAlignRightCenter
   grd_Listad.ColAlignment(7) = flexAlignRightCenter

   'Obteniendo ITF
   Call moddat_gs_FecSis
   l_dbl_PorITF = opecaj_gf_Consulta_ITF(Format(CDate(moddat_g_str_FecSis), "yyyymmdd"), 1)
End Sub

Private Sub fs_Limpia()
   Call fs_ActivaItem(False)
   Call fs_Activa(True)
   
   cmb_TipBus.ListIndex = -1
   cmb_TipDoc.Enabled = False
   txt_NumDoc.Enabled = False
   msk_NumOpe.Enabled = False

   msk_NumOpe.Mask = ""
   msk_NumOpe.Text = ""
   msk_NumOpe.Mask = "###-##-#####"
   
   txt_NumDoc.Text = ""
   
   pnl_Client.Caption = ""
   
   Call gs_LimpiaGrid(grd_Listad)
   
   pnl_SalPen.Caption = "0.00 "
   pnl_TotDes.Caption = "0.00 "
   pnl_ImpITF.Caption = "0.00 "
   
   
   pnl_NumOpe.Caption = ""
   pnl_Produc.Caption = ""
   pnl_Modali.Caption = ""
   pnl_NumSol.Caption = ""
   
   pnl_Cre_TipMon.Caption = ""
   pnl_Cre_ComVta.Caption = "0.00 "
   pnl_Cre_ApoPro.Caption = "0.00 "
   pnl_Cre_MtoPre.Caption = "0.00 "
   pnl_Cre_MtoDol.Caption = "0.00 "
   pnl_Cre_MtoSol.Caption = "0.00 "
   pnl_Cre_MtoMPr.Caption = "0.00 "
   pnl_Cre_NumCuo.Caption = "0 "
   pnl_Cre_PerGra.Caption = "0 "
   
   pnl_Inm_Direcc.Caption = ""
   pnl_Inm_Propie.Caption = ""
   
   txt_Leg_InfLeg.Text = ""
   
   pnl_Aut_FueFin.Caption = ""
   pnl_Aut_BonoBP.Caption = ""
   pnl_Aut_FecDes.Caption = ""
   txt_Aut_Observ.Text = ""
   
   Call fs_LimpiaItem
End Sub

Private Sub fs_Activa(ByVal p_Habilita As Integer)
   cmb_TipBus.Enabled = p_Habilita
   cmb_TipDoc.Enabled = p_Habilita
   txt_NumDoc.Enabled = p_Habilita
   msk_NumOpe.Enabled = p_Habilita
   cmd_Buscar.Enabled = p_Habilita
   
   grd_Listad.Enabled = Not p_Habilita
   
   tab_Princi.Enabled = Not p_Habilita
   cmd_ImpCro.Enabled = Not p_Habilita
   cmd_ImpLiq.Enabled = Not p_Habilita
   cmd_NueDes.Enabled = Not p_Habilita
   cmd_Imprim.Enabled = Not p_Habilita
End Sub

Private Sub fs_ActivaItem(ByVal p_Habilita As Integer)
   grd_Listad.Enabled = Not p_Habilita
   cmd_ImpCro.Enabled = Not p_Habilita
   cmd_ImpLiq.Enabled = Not p_Habilita
   cmd_NueDes.Enabled = Not p_Habilita
   cmd_Imprim.Enabled = Not p_Habilita

   cmb_MonDes.Enabled = p_Habilita
   ipp_MtoDes.Enabled = p_Habilita
   cmb_TipDes.Enabled = p_Habilita
   cmb_BanCar.Enabled = p_Habilita
   cmb_CtaCar.Enabled = p_Habilita
   txt_CheCar.Enabled = p_Habilita
   cmb_BanAbo.Enabled = p_Habilita
   txt_CtaAbo.Enabled = p_Habilita
   txt_TraAbo.Enabled = p_Habilita
   txt_Garant.Enabled = p_Habilita
   txt_Observ.Enabled = p_Habilita
   
   cmd_Grabar.Enabled = p_Habilita
   cmd_Cancel.Enabled = p_Habilita
End Sub

Private Sub fs_LimpiaItem()
   cmb_MonDes.ListIndex = -1
   ipp_MtoDes.Value = 0
   cmb_TipDes.ListIndex = -1
   cmb_BanCar.ListIndex = -1
   cmb_CtaCar.Clear
   txt_CheCar.Text = ""
   cmb_BanAbo.ListIndex = -1
   txt_CtaAbo.Text = ""
   txt_TraAbo.Text = ""
   txt_Garant.Text = ""
   txt_Observ.Text = ""
   
   pnl_DesMPr.Caption = "0.00 "
   pnl_TCaDol.Caption = "0.00 "
   pnl_TCaMPr.Caption = "0.00 "
End Sub

Private Sub fs_Buscar_DatGen()
   g_rst_Princi.MoveFirst
   
   moddat_g_int_TipDoc = g_rst_Princi!HIPMAE_TDOCLI
   moddat_g_str_NumDoc = Trim(g_rst_Princi!HIPMAE_NDOCLI)
   moddat_g_str_NumSol = Trim(g_rst_Princi!HIPMAE_NUMSOL)
   moddat_g_str_NumOpe = Trim(g_rst_Princi!HIPMAE_NUMOPE)
   
   'Obteniendo Nombre de Cliente
   moddat_g_str_NomCli = moddat_gf_Buscar_NomCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc)
   
   'Obteniendo Nombre y DOI de Cónyuge
   moddat_g_int_CygTDo = g_rst_Princi!HIPMAE_TDOCYG
   moddat_g_str_CygNDo = Trim(g_rst_Princi!HIPMAE_NDOCYG & "")
   moddat_g_str_CygNom = moddat_gf_Buscar_NomCli(moddat_g_int_CygTDo, moddat_g_str_CygNDo)
   
   'Obteniendo Descripción de Producto
   moddat_g_str_CodPrd = Trim(g_rst_Princi!HIPMAE_CODPRD)
   moddat_g_str_NomPrd = moddat_gf_Consulta_Produc(Trim(g_rst_Princi!HIPMAE_CODPRD))

   'Obeniendo Modalidad de Producto
   moddat_g_str_CodMod = Trim(g_rst_Princi!HIPMAE_CODMOD)
   moddat_g_str_DesMod = moddat_gf_Buscar_NomMod(Trim(g_rst_Princi!HIPMAE_CODPRD), moddat_g_str_CodMod)

   'Moneda
   moddat_g_str_Moneda = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!HIPMAE_MONEDA))
   moddat_g_int_TipMon = g_rst_Princi!HIPMAE_MONEDA
   
   'Fecha de Activación
   moddat_g_str_FecApr = gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECAPR))
   
   'Obteniendo Monto de Préstamo y Total Desembolsado
   l_dbl_MtoPre = g_rst_Princi!HIPMAE_MTOPRE
   l_dbl_MtoDes = g_rst_Princi!HIPMAE_IMPDES
   
   pnl_NumOpe.Caption = Mid(moddat_g_str_NumOpe, 1, 3) & "-" & Mid(moddat_g_str_NumOpe, 4, 2) & "-" & Mid(moddat_g_str_NumOpe, 6, 5)
   pnl_NumSol.Caption = Mid(moddat_g_str_NumSol, 1, 3) & "-" & Mid(moddat_g_str_NumSol, 4, 3) & "-" & Mid(moddat_g_str_NumSol, 7, 2) & "-" & Mid(moddat_g_str_NumSol, 9, 4)
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   pnl_Produc.Caption = moddat_g_str_NomPrd
   pnl_Modali.Caption = moddat_g_str_DesMod
   
   pnl_Cre_TipMon.Caption = moddat_g_str_Moneda
   pnl_Cre_ComVta.Caption = Format(g_rst_Princi!HIPMAE_COMVTA, "###,###,##0.00") & " "
   pnl_Cre_ApoPro.Caption = Format(g_rst_Princi!HIPMAE_APOPRO, "###,###,##0.00") & " "
   pnl_Cre_MtoDol.Caption = Format(g_rst_Princi!HIPMAE_PREDOL, "###,###,##0.00") & " "
   pnl_Cre_MtoSol.Caption = Format(g_rst_Princi!HIPMAE_PRESOL, "###,###,##0.00") & " "
   pnl_Cre_MtoMPr.Caption = Format(g_rst_Princi!HIPMAE_PREMPR, "###,###,##0.00") & " "
   pnl_Cre_MtoPre.Caption = Format(g_rst_Princi!HIPMAE_MTOPRE, "###,###,##0.00") & " "
   pnl_Cre_NumCuo.Caption = Format(g_rst_Princi!HIPMAE_NUMCUO, "###0") & " "
   pnl_Cre_PerGra.Caption = Format(g_rst_Princi!HIPMAE_PERGRA, "###0") & " "
   
   pnl_SalPen.Caption = Format(l_dbl_MtoPre - l_dbl_MtoDes, "###,###,##0.00") & " "
   pnl_TotDes.Caption = Format(l_dbl_MtoDes, "###,###,##0.00") & " "
End Sub

Private Sub txt_TraAbo_GotFocus()
   Call gs_SelecTodo(txt_TraAbo)
End Sub

Private Sub txt_TraAbo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Garant)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_Garant_GotFocus()
   Call gs_SelecTodo(txt_Garant)
End Sub

Private Sub txt_Garant_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Observ)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_., ;:()/&%$·!ª@#=?¿+*" & Chr(10))
   End If
End Sub

Private Sub txt_Observ_GotFocus()
   Call gs_SelecTodo(txt_Observ)
End Sub

Private Sub txt_Observ_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_., ;:()/&%$·!ª@#=?¿+*" & Chr(10))
   End If
End Sub

Private Sub fs_Buscar_Desemb()
   Call gs_LimpiaGrid(grd_Listad)

   g_str_Parame = "SELECT * FROM CRE_HIPDES WHERE "
   g_str_Parame = g_str_Parame & "HIPDES_NUMOPE = '" & moddat_g_str_NumOpe & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Exit Sub
   End If
   
   grd_Listad.Redraw = False
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      
      grd_Listad.Col = 0
      grd_Listad.Text = Format(g_rst_Princi!HIPDES_NUMDES, "000")
   
      grd_Listad.Col = 1
      grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPDES_FECDES))
   
      grd_Listad.Col = 2
      grd_Listad.Text = moddat_gf_Consulta_ParDes("226", g_rst_Princi!HIPDES_TIPDES)
   
      grd_Listad.Col = 3
      grd_Listad.Text = moddat_gf_Consulta_ParDes("204", g_rst_Princi!HIPDES_TIPMON)
   
      grd_Listad.Col = 4
      grd_Listad.Text = Format(g_rst_Princi!HIPDES_IMPORT, "###,###,##0.00")
   
      grd_Listad.Col = 5
      grd_Listad.Text = Format(g_rst_Princi!HIPDES_IMPITF, "###,###,##0.00")
   
      grd_Listad.Col = 6
      grd_Listad.Text = Format(g_rst_Princi!HIPDES_IMPORT - g_rst_Princi!HIPDES_IMPITF, "###,###,##0.00")
   
      grd_Listad.Col = 7
      grd_Listad.Text = Format(g_rst_Princi!HIPDES_DESMPR, "###,###,##0.00")
   
      g_rst_Princi.MoveNext
   Loop
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   grd_Listad.Redraw = True
   Call gs_UbiIniGrid(grd_Listad)
   
   
   'Obteniendo Nuevos Saldos
   g_str_Parame = "SELECT * FROM CRE_HIPMAE WHERE "
   g_str_Parame = g_str_Parame & "HIPMAE_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   l_dbl_MtoPre = g_rst_Princi!HIPMAE_MTOPRE
   l_dbl_MtoDes = g_rst_Princi!HIPMAE_IMPDES
   
   pnl_SalPen.Caption = Format(l_dbl_MtoPre - l_dbl_MtoDes, "###,###,##0.00") & " "
   pnl_TotDes.Caption = Format(l_dbl_MtoDes, "###,###,##0.00") & " "
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub


