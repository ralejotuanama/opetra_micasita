VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm_Des_CreHip_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   6570
   ClientLeft      =   990
   ClientTop       =   2310
   ClientWidth     =   12810
   Icon            =   "OpeTra_frm_002.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   12810
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   6555
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12825
      _Version        =   65536
      _ExtentX        =   22622
      _ExtentY        =   11562
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
      Begin Threed.SSPanel SSPanel2 
         Height          =   765
         Left            =   30
         TabIndex        =   1
         Top             =   5730
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
         Begin VB.CommandButton cmd_Imprim 
            Height          =   675
            Left            =   11310
            Picture         =   "OpeTra_frm_002.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Imprimir Cronograma"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   12000
            Picture         =   "OpeTra_frm_002.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   4
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
            TabIndex        =   5
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
            TabIndex        =   6
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
            Picture         =   "OpeTra_frm_002.frx":0890
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   1095
         Index           =   4
         Left            =   30
         TabIndex        =   7
         Top             =   750
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
            TabIndex        =   8
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
            TabIndex        =   9
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
            TabIndex        =   10
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
            TabIndex        =   11
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
            Caption         =   "Producto:"
            Height          =   315
            Index           =   188
            Left            =   60
            TabIndex        =   15
            Top             =   390
            Width           =   1335
         End
         Begin VB.Label lbl_NomGlo 
            Caption         =   "Modalidad:"
            Height          =   315
            Index           =   187
            Left            =   60
            TabIndex        =   14
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label lbl_NomGlo 
            Caption         =   "Nro. Operación:"
            Height          =   315
            Index           =   184
            Left            =   60
            TabIndex        =   13
            Top             =   60
            Width           =   1335
         End
         Begin VB.Label lbl_NomGlo 
            Caption         =   "Nro. Solicitud"
            Height          =   315
            Index           =   20
            Left            =   7110
            TabIndex        =   12
            Top             =   60
            Width           =   1335
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   3795
         Left            =   30
         TabIndex        =   16
         Top             =   1890
         Width           =   12735
         _Version        =   65536
         _ExtentX        =   22463
         _ExtentY        =   6694
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
            Height          =   3675
            Left            =   60
            TabIndex        =   17
            Top             =   60
            Width           =   12585
            _ExtentX        =   22199
            _ExtentY        =   6482
            _Version        =   393216
            Style           =   1
            Tabs            =   4
            TabsPerRow      =   4
            TabHeight       =   520
            TabCaption(0)   =   "Cliente - Tramo No Concesional"
            TabPicture(0)   =   "OpeTra_frm_002.frx":0B9A
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Label3"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "pnl_CliNCo_TotCuo"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "pnl_CliNCo_OtrCar"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "SSPanel62"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "SSPanel61"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "SSPanel59"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "SSPanel14"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "SSPanel13"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).Control(8)=   "SSPanel12"
            Tab(0).Control(8).Enabled=   0   'False
            Tab(0).Control(9)=   "SSPanel11"
            Tab(0).Control(9).Enabled=   0   'False
            Tab(0).Control(10)=   "SSPanel10"
            Tab(0).Control(10).Enabled=   0   'False
            Tab(0).Control(11)=   "SSPanel9"
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
            TabPicture(1)   =   "OpeTra_frm_002.frx":0BB6
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "grd_CliCon_Listad"
            Tab(1).Control(1)=   "SSPanel23"
            Tab(1).Control(2)=   "SSPanel19"
            Tab(1).Control(3)=   "SSPanel21"
            Tab(1).Control(4)=   "SSPanel22"
            Tab(1).Control(5)=   "SSPanel24"
            Tab(1).Control(6)=   "SSPanel25"
            Tab(1).Control(7)=   "pnl_CliCon_TotCuo"
            Tab(1).Control(8)=   "pnl_CliCon_Intere"
            Tab(1).Control(9)=   "pnl_CliCon_Capita"
            Tab(1).Control(10)=   "Label13"
            Tab(1).ControlCount=   11
            TabCaption(2)   =   "Cofide - Tramo No Concesional"
            TabPicture(2)   =   "OpeTra_frm_002.frx":0BD2
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "grd_CofNCo_Listad"
            Tab(2).Control(1)=   "SSPanel26"
            Tab(2).Control(2)=   "SSPanel27"
            Tab(2).Control(3)=   "SSPanel28"
            Tab(2).Control(4)=   "SSPanel29"
            Tab(2).Control(5)=   "SSPanel31"
            Tab(2).Control(6)=   "SSPanel32"
            Tab(2).Control(7)=   "pnl_CofNCo_Comisi"
            Tab(2).Control(8)=   "pnl_CofNCo_Intere"
            Tab(2).Control(9)=   "pnl_CofNCo_Capita"
            Tab(2).Control(10)=   "SSPanel37"
            Tab(2).Control(11)=   "pnl_CofNCo_TotCuo"
            Tab(2).Control(12)=   "Label1"
            Tab(2).ControlCount=   13
            TabCaption(3)   =   "Cofide - Tramo Concesional"
            TabPicture(3)   =   "OpeTra_frm_002.frx":0BEE
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "Label2"
            Tab(3).Control(1)=   "pnl_CofCon_TotCuo"
            Tab(3).Control(2)=   "SSPanel57"
            Tab(3).Control(3)=   "pnl_CofCon_Capita"
            Tab(3).Control(4)=   "pnl_CofCon_Intere"
            Tab(3).Control(5)=   "pnl_CofCon_Comisi"
            Tab(3).Control(6)=   "SSPanel53"
            Tab(3).Control(7)=   "SSPanel52"
            Tab(3).Control(8)=   "SSPanel51"
            Tab(3).Control(9)=   "SSPanel50"
            Tab(3).Control(10)=   "SSPanel48"
            Tab(3).Control(11)=   "SSPanel40"
            Tab(3).Control(12)=   "grd_CofCon_Listad"
            Tab(3).ControlCount=   13
            Begin Threed.SSPanel pnl_CliNCo_Capita 
               Height          =   285
               Left            =   2550
               TabIndex        =   18
               Top             =   3300
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "0.00 "
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
               Left            =   6690
               TabIndex        =   19
               Top             =   3300
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "0.00 "
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
               Left            =   5310
               TabIndex        =   20
               Top             =   3300
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "0.00 "
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
               Left            =   3930
               TabIndex        =   21
               Top             =   3300
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "0.00 "
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
               Height          =   2625
               Left            =   30
               TabIndex        =   22
               Top             =   660
               Width           =   12465
               _ExtentX        =   21987
               _ExtentY        =   4630
               _Version        =   393216
               Rows            =   21
               Cols            =   9
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_CofCon_Listad 
               Height          =   2625
               Left            =   -74970
               TabIndex        =   23
               Top             =   660
               Width           =   12465
               _ExtentX        =   21987
               _ExtentY        =   4630
               _Version        =   393216
               Rows            =   21
               Cols            =   7
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_CofNCo_Listad 
               Height          =   2625
               Left            =   -74970
               TabIndex        =   24
               Top             =   660
               Width           =   12465
               _ExtentX        =   21987
               _ExtentY        =   4630
               _Version        =   393216
               Rows            =   21
               Cols            =   7
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_CliCon_Listad 
               Height          =   2625
               Left            =   -74970
               TabIndex        =   25
               Top             =   660
               Width           =   12465
               _ExtentX        =   21987
               _ExtentY        =   4630
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
            Begin Threed.SSPanel SSPanel23 
               Height          =   285
               Left            =   -69870
               TabIndex        =   26
               Top             =   360
               Width           =   2370
               _Version        =   65536
               _ExtentX        =   4180
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
            Begin Threed.SSPanel SSPanel19 
               Height          =   285
               Left            =   -74940
               TabIndex        =   27
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
            Begin Threed.SSPanel SSPanel21 
               Height          =   285
               Left            =   -73770
               TabIndex        =   28
               Top             =   360
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
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
            Begin Threed.SSPanel SSPanel22 
               Height          =   285
               Left            =   -72210
               TabIndex        =   29
               Top             =   360
               Width           =   2370
               _Version        =   65536
               _ExtentX        =   4180
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
            Begin Threed.SSPanel SSPanel24 
               Height          =   285
               Left            =   -67530
               TabIndex        =   30
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
               TabIndex        =   31
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
            Begin Threed.SSPanel pnl_CliCon_TotCuo 
               Height          =   285
               Left            =   -67530
               TabIndex        =   32
               Top             =   3300
               Width           =   2370
               _Version        =   65536
               _ExtentX        =   4180
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "0.00 "
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
            Begin Threed.SSPanel pnl_CliCon_Intere 
               Height          =   285
               Left            =   -69870
               TabIndex        =   33
               Top             =   3300
               Width           =   2370
               _Version        =   65536
               _ExtentX        =   4180
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "0.00 "
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
               Left            =   -72210
               TabIndex        =   34
               Top             =   3300
               Width           =   2370
               _Version        =   65536
               _ExtentX        =   4180
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "0.00 "
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
            Begin Threed.SSPanel SSPanel26 
               Height          =   285
               Left            =   -74940
               TabIndex        =   35
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
               TabIndex        =   36
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
               TabIndex        =   37
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
               TabIndex        =   38
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
               TabIndex        =   39
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
               TabIndex        =   40
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
            Begin Threed.SSPanel pnl_CofNCo_Comisi 
               Height          =   285
               Left            =   -68310
               TabIndex        =   41
               Top             =   3300
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "0.00 "
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
            Begin Threed.SSPanel pnl_CofNCo_Intere 
               Height          =   285
               Left            =   -70140
               TabIndex        =   42
               Top             =   3300
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "0.00 "
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
               Left            =   -71970
               TabIndex        =   43
               Top             =   3300
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "0.00 "
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
            Begin Threed.SSPanel SSPanel37 
               Height          =   285
               Left            =   -68310
               TabIndex        =   44
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
            Begin Threed.SSPanel pnl_CofNCo_TotCuo 
               Height          =   285
               Left            =   -66480
               TabIndex        =   45
               Top             =   3300
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "0.00 "
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
            Begin Threed.SSPanel SSPanel40 
               Height          =   285
               Left            =   -74940
               TabIndex        =   46
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
               TabIndex        =   47
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
               TabIndex        =   48
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
               TabIndex        =   49
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
               TabIndex        =   50
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
               TabIndex        =   51
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
            Begin Threed.SSPanel pnl_CofCon_Comisi 
               Height          =   285
               Left            =   -68310
               TabIndex        =   52
               Top             =   3300
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "0.00 "
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
            Begin Threed.SSPanel pnl_CofCon_Intere 
               Height          =   285
               Left            =   -70140
               TabIndex        =   53
               Top             =   3300
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "0.00 "
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
            Begin Threed.SSPanel pnl_CofCon_Capita 
               Height          =   285
               Left            =   -71970
               TabIndex        =   54
               Top             =   3300
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "0.00 "
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
            Begin Threed.SSPanel SSPanel57 
               Height          =   285
               Left            =   -68310
               TabIndex        =   55
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
            Begin Threed.SSPanel pnl_CofCon_TotCuo 
               Height          =   285
               Left            =   -66480
               TabIndex        =   56
               Top             =   3300
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "0.00 "
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
               Left            =   60
               TabIndex        =   57
               Top             =   360
               Width           =   1125
               _Version        =   65536
               _ExtentX        =   1984
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
            Begin Threed.SSPanel SSPanel10 
               Height          =   285
               Left            =   1170
               TabIndex        =   58
               Top             =   360
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
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
            Begin Threed.SSPanel SSPanel11 
               Height          =   285
               Left            =   2550
               TabIndex        =   59
               Top             =   360
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
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
            Begin Threed.SSPanel SSPanel12 
               Height          =   285
               Left            =   3930
               TabIndex        =   60
               Top             =   360
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
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
            Begin Threed.SSPanel SSPanel13 
               Height          =   285
               Left            =   9450
               TabIndex        =   61
               Top             =   360
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
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
            Begin Threed.SSPanel SSPanel14 
               Height          =   285
               Left            =   10830
               TabIndex        =   62
               Top             =   360
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
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
            Begin Threed.SSPanel SSPanel59 
               Height          =   285
               Left            =   5310
               TabIndex        =   63
               Top             =   360
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Seg. Prest."
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
            Begin Threed.SSPanel SSPanel61 
               Height          =   285
               Left            =   6690
               TabIndex        =   64
               Top             =   360
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Seg. Vivienda"
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
            Begin Threed.SSPanel SSPanel62 
               Height          =   285
               Left            =   8070
               TabIndex        =   65
               Top             =   360
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Otros Cargos"
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
               Left            =   8070
               TabIndex        =   66
               Top             =   3300
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "0.00 "
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
               Left            =   9450
               TabIndex        =   67
               Top             =   3300
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "0.00 "
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
            Begin VB.Label Label13 
               Caption         =   "Totales ==>"
               Height          =   285
               Left            =   -73230
               TabIndex        =   71
               Top             =   3300
               Width           =   945
            End
            Begin VB.Label Label1 
               Caption         =   "Totales ==>"
               Height          =   285
               Left            =   -72930
               TabIndex        =   70
               Top             =   3300
               Width           =   945
            End
            Begin VB.Label Label2 
               Caption         =   "Totales ==>"
               Height          =   285
               Left            =   -72930
               TabIndex        =   69
               Top             =   3300
               Width           =   945
            End
            Begin VB.Label Label3 
               Caption         =   "Totales ==>"
               Height          =   285
               Left            =   1470
               TabIndex        =   68
               Top             =   3300
               Width           =   945
            End
         End
      End
   End
End
Attribute VB_Name = "frm_Des_CreHip_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Imprim_Click()
   Select Case tab_Cronog.Tab
      Case 0
         Call fs_ImpCro_CliNCo(moddat_g_str_NumOpe)
         
      Case 1
         If grd_CliCon_Listad.Rows = 0 Then
            Exit Sub
         End If
         Call fs_ImpCro_CliCon(moddat_g_str_NumOpe)
         
      Case 2
         If grd_CofNCo_Listad.Rows = 0 Then
            Exit Sub
         End If
         Call fs_ImpCro_CofNCo(moddat_g_str_NumOpe)
         
      Case 3
         If grd_CofCon_Listad.Rows = 0 Then
            Exit Sub
         End If
         Call fs_ImpCro_CofCon(moddat_g_str_NumOpe)

   End Select
   
   frm_Imprim_01.Show 1
End Sub

Private Sub fs_ImpCro_CliNCo(ByVal p_NumOpe As String)
   Dim r_str_Linea      As String
   
   Dim r_dbl_Capita     As Double
   Dim r_dbl_Intere     As Double
   Dim r_dbl_SegDes     As Double
   Dim r_dbl_SegViv     As Double
   Dim r_dbl_OtrCar     As Double
   Dim r_dbl_ImpCuo     As Double
   Dim r_dbl_TotCuo     As Double
   
   Dim r_int_NumLin     As Integer
   
   
   r_dbl_Capita = 0
   r_dbl_Intere = 0
   r_dbl_SegDes = 0
   r_dbl_SegViv = 0
   r_dbl_OtrCar = 0
   r_dbl_TotCuo = 0
   
   'Obteniendo Información de la Operación
   g_str_Parame = "SELECT * FROM CRE_HIPMAE WHERE "
   g_str_Parame = g_str_Parame & "HIPMAE_NUMOPE = '" & p_NumOpe & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If

   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      Exit Sub
   End If
   
   'Inicializando Arreglo de Impresiones
   ReDim g_arr_Imprim(0)

   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp(Space(98) & "Fecha: " & moddat_g_str_FecSis)
   Call gs_LinImp(Space(98) & "Hora:  " & Space(2) & Format(Time, "hh:mm:ss"))
   Call gs_LinImp("")
   Call gs_LinImp(Space(38) & "CRONOGRAMA DE PAGOS - CREDITO HIPOTECARIO")
   Call gs_LinImp(Space(38) & "-----------------------------------------")
   Call gs_LinImp("")
   
   Call gs_LinImp(Space(5) & "Nro. de Operación     : " & Mid(p_NumOpe, 1, 3) & "-" & Mid(p_NumOpe, 4, 2) & "-" & Mid(p_NumOpe, 6, 5))
   Call gs_LinImp(Space(5) & "Docum. Ident. Cliente : " & CStr(g_rst_Princi!HIPMAE_TDOCLI) & "-" & Trim(g_rst_Princi!HIPMAE_NDOCLI))
   Call gs_LinImp(Space(5) & "Nombre Cliente        : " & moddat_gf_Buscar_NomCli(g_rst_Princi!HIPMAE_TDOCLI, Trim(g_rst_Princi!HIPMAE_NDOCLI)))
   Call gs_LinImp("")
   
   Call gs_LinImp(Space(5) & "Moneda de Préstamo    : " & moddat_gf_Consulta_Pardes("204", CStr(g_rst_Princi!HIPMAE_MONEDA)))
   Call gs_LinImp(Space(5) & "Total Préstamo        : " & gf_FormatoNumero(g_rst_Princi!HIPMAE_PREMPR, 15))
   Call gs_LinImp(Space(5) & "Fecha Desembolso      : " & gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECAPR)))
   
   '1er Bloque hasta 42 Caracteres / 2do Bloque Hasta 35 /3er Bloque Hasta 30
   
   r_str_Linea = ""
   r_str_Linea = r_str_Linea & "Nro. Cuotas           : " & Format(g_rst_Princi!HIPMAE_NUMCUO, "000") & Space(15)
   r_str_Linea = r_str_Linea & "Cuotas Extradord. : " & Mid(moddat_gf_Consulta_Pardes("223", g_rst_Princi!HIPMAE_CUOANO) & Space(20), 1, 18)
   r_str_Linea = r_str_Linea & "Período de Gracia : " & Mid(Format(g_rst_Princi!HIPMAE_PERGRA, "#0") & Space(18), 1, 18)
   
   Call gs_LinImp(Space(5) & r_str_Linea)
   
   r_str_Linea = ""
   r_str_Linea = r_str_Linea & "Tasa de Interes       : " & gf_FormatoNumero(g_rst_Princi!HIPMAE_TASINT, 6, 0) & "%"
   
   Call gs_LinImp(Space(5) & r_str_Linea)
   
   Call gs_LinImp(Space(5) & "Tipo de Cronograma    : " & moddat_gf_Consulta_Pardes("028", CStr(1)))
   Call gs_LinImp(Space(5) & "Monto Préstamo Tramo  : " & gf_FormatoNumero(g_rst_Princi!HIPMAE_IMPNCO, 15))
   Call gs_LinImp("")
   
   
   
   Call gs_LinImp(Space(5) & String(110, "-"))
   
   r_str_Linea = ""
   r_str_Linea = r_str_Linea & "Cuota" & Space(2)
   r_str_Linea = r_str_Linea & "F. Vencimiento" & Space(2)
   r_str_Linea = r_str_Linea & "  Capital " & Space(2)
   r_str_Linea = r_str_Linea & "  Interes " & Space(2)
   r_str_Linea = r_str_Linea & "Seg.Prest." & Space(2)
   r_str_Linea = r_str_Linea & "Seg.Vivie." & Space(2)
   r_str_Linea = r_str_Linea & "Otr.Cargos" & Space(2)
   r_str_Linea = r_str_Linea & "Total Cuota" & Space(2)
   r_str_Linea = r_str_Linea & "Saldo Capital" & Space(2)
   
   Call gs_LinImp(Space(5) & r_str_Linea)
   Call gs_LinImp(Space(5) & String(110, "-"))
   
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   r_int_NumLin = 22
   
   g_str_Parame = "SELECT * FROM CRE_HIPCUO WHERE "
   g_str_Parame = g_str_Parame & "HIPCUO_NUMOPE = '" & p_NumOpe & "' AND "
   g_str_Parame = g_str_Parame & "HIPCUO_TIPCRO = 1"
   g_str_Parame = g_str_Parame & "ORDER BY HIPCUO_NUMCUO ASC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If

   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst

      Do While Not g_rst_Princi.EOF
         If r_int_NumLin = 90 Then
            Call gs_LinImp("SP")
            r_int_NumLin = 1
         End If
         If r_int_NumLin = 1 Then
            Call gs_LinImp("")
            Call gs_LinImp("")
            
            Call gs_LinImp(Space(5) & String(110, "-"))
            
            r_str_Linea = ""
            r_str_Linea = r_str_Linea & "Cuota" & Space(2)
            r_str_Linea = r_str_Linea & "F. Vencimiento" & Space(2)
            r_str_Linea = r_str_Linea & "  Capital " & Space(2)
            r_str_Linea = r_str_Linea & "  Interes " & Space(2)
            r_str_Linea = r_str_Linea & "Seg.Prest." & Space(2)
            r_str_Linea = r_str_Linea & "Seg.Vivie." & Space(2)
            r_str_Linea = r_str_Linea & "Otr.Cargos" & Space(2)
            r_str_Linea = r_str_Linea & "Total Cuota" & Space(2)
            r_str_Linea = r_str_Linea & "Saldo Capital" & Space(2)
            
            Call gs_LinImp(Space(5) & r_str_Linea)
            Call gs_LinImp(Space(5) & String(110, "-"))
            
            r_int_NumLin = 5
         End If
         
         r_dbl_ImpCuo = 0
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(Format(g_rst_Princi!HIPCUO_CAPITA, "###,###,##0.00"))
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(Format(g_rst_Princi!HIPCUO_INTERE, "###,###,##0.00"))
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(Format(g_rst_Princi!HIPCUO_DESORG, "###,###,##0.00"))
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(Format(g_rst_Princi!HIPCUO_VIVORG, "###,###,##0.00"))
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(Format(g_rst_Princi!HIPCUO_OTRORG, "###,###,##0.00"))

         r_dbl_Capita = r_dbl_Capita + CDbl(Format(g_rst_Princi!HIPCUO_CAPITA, "###,###,##0.00"))
         r_dbl_Intere = r_dbl_Intere + CDbl(Format(g_rst_Princi!HIPCUO_INTERE, "###,###,##0.00"))
         r_dbl_SegDes = r_dbl_SegDes + CDbl(Format(g_rst_Princi!HIPCUO_DESORG, "###,###,##0.00"))
         r_dbl_SegViv = r_dbl_SegViv + CDbl(Format(g_rst_Princi!HIPCUO_VIVORG, "###,###,##0.00"))
         r_dbl_OtrCar = r_dbl_OtrCar + CDbl(Format(g_rst_Princi!HIPCUO_OTRORG, "###,###,##0.00"))
         r_dbl_TotCuo = r_dbl_TotCuo + r_dbl_ImpCuo

         r_str_Linea = ""
         r_str_Linea = r_str_Linea & Space(1) & Format(g_rst_Princi!HIPCUO_NUMCUO, "000") & Space(1)
         r_str_Linea = r_str_Linea & Space(2)
         
         r_str_Linea = r_str_Linea & Space(2) & gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECVCT)) & Space(2)
         r_str_Linea = r_str_Linea & Space(2)
         
         r_str_Linea = r_str_Linea & gf_FormatoNumero(g_rst_Princi!HIPCUO_CAPITA, 10) & Space(2)
         r_str_Linea = r_str_Linea & gf_FormatoNumero(g_rst_Princi!HIPCUO_INTERE, 10) & Space(2)
         r_str_Linea = r_str_Linea & gf_FormatoNumero(g_rst_Princi!HIPCUO_DESORG, 10) & Space(2)
         r_str_Linea = r_str_Linea & gf_FormatoNumero(g_rst_Princi!HIPCUO_VIVORG, 10) & Space(2)
         r_str_Linea = r_str_Linea & gf_FormatoNumero(g_rst_Princi!HIPCUO_OTRORG, 10) & Space(2)
         r_str_Linea = r_str_Linea & gf_FormatoNumero(r_dbl_ImpCuo, 10) & Space(2)
         r_str_Linea = r_str_Linea & gf_FormatoNumero(g_rst_Princi!HIPCUO_SALCAP, 12) & Space(2)
         
         Call gs_LinImp(Space(5) & r_str_Linea)

         r_int_NumLin = r_int_NumLin + 1
         
         g_rst_Princi.MoveNext
      Loop
   End If

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   
   Call gs_LinImp(Space(5) & String(110, "-"))
   
   r_str_Linea = ""
   r_str_Linea = r_str_Linea & "TOTAL GENERAL" & Space(10)
   r_str_Linea = r_str_Linea & gf_FormatoNumero(r_dbl_Capita, 10) & Space(2)
   r_str_Linea = r_str_Linea & gf_FormatoNumero(r_dbl_Intere, 10) & Space(2)
   r_str_Linea = r_str_Linea & gf_FormatoNumero(r_dbl_SegDes, 10) & Space(2)
   r_str_Linea = r_str_Linea & gf_FormatoNumero(r_dbl_SegViv, 10) & Space(2)
   r_str_Linea = r_str_Linea & gf_FormatoNumero(r_dbl_OtrCar, 10) & Space(2)
   r_str_Linea = r_str_Linea & gf_FormatoNumero(r_dbl_TotCuo, 10) & Space(2)
   
   Call gs_LinImp(Space(5) & r_str_Linea)
   Call gs_LinImp(Space(5) & String(110, "-"))
End Sub

Private Sub fs_ImpCro_CliCon(ByVal p_NumOpe As String)
   Dim r_str_Linea      As String
   
   Dim r_dbl_Capita     As Double
   Dim r_dbl_Intere     As Double
   Dim r_dbl_SegDes     As Double
   Dim r_dbl_SegViv     As Double
   Dim r_dbl_OtrCar     As Double
   Dim r_dbl_ImpCuo     As Double
   Dim r_dbl_TotCuo     As Double
   
   Dim r_int_NumLin     As Integer
   
   
   r_dbl_Capita = 0
   r_dbl_Intere = 0
   r_dbl_SegDes = 0
   r_dbl_SegViv = 0
   r_dbl_OtrCar = 0
   r_dbl_TotCuo = 0
   
   'Obteniendo Información de la Operación
   g_str_Parame = "SELECT * FROM CRE_HIPMAE WHERE "
   g_str_Parame = g_str_Parame & "HIPMAE_NUMOPE = '" & p_NumOpe & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If

   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      Exit Sub
   End If
   
   'Inicializando Arreglo de Impresiones
   ReDim g_arr_Imprim(0)

   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp(Space(98) & "Fecha: " & moddat_g_str_FecSis)
   Call gs_LinImp(Space(98) & "Hora:  " & Space(2) & Format(Time, "hh:mm:ss"))
   Call gs_LinImp("")
   Call gs_LinImp(Space(38) & "CRONOGRAMA DE PAGOS - CREDITO HIPOTECARIO")
   Call gs_LinImp(Space(38) & "-----------------------------------------")
   Call gs_LinImp("")
   
   Call gs_LinImp(Space(5) & "Nro. de Operación     : " & Mid(p_NumOpe, 1, 3) & "-" & Mid(p_NumOpe, 4, 2) & "-" & Mid(p_NumOpe, 6, 5))
   Call gs_LinImp(Space(5) & "Docum. Ident. Cliente : " & CStr(g_rst_Princi!HIPMAE_TDOCLI) & "-" & Trim(g_rst_Princi!HIPMAE_NDOCLI))
   Call gs_LinImp(Space(5) & "Nombre Cliente        : " & moddat_gf_Buscar_NomCli(g_rst_Princi!HIPMAE_TDOCLI, Trim(g_rst_Princi!HIPMAE_NDOCLI)))
   Call gs_LinImp("")
   
   Call gs_LinImp(Space(5) & "Moneda de Préstamo    : " & moddat_gf_Consulta_Pardes("204", CStr(g_rst_Princi!HIPMAE_MONEDA)))
   Call gs_LinImp(Space(5) & "Total Préstamo        : " & gf_FormatoNumero(g_rst_Princi!HIPMAE_PREMPR, 15))
   Call gs_LinImp(Space(5) & "Fecha Desembolso      : " & gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECAPR)))
   
   '1er Bloque hasta 42 Caracteres / 2do Bloque Hasta 35 /3er Bloque Hasta 30
   
   r_str_Linea = ""
   r_str_Linea = r_str_Linea & "Nro. Cuotas           : " & Format(g_rst_Princi!HIPMAE_NUMCUO, "000") & Space(15)
   r_str_Linea = r_str_Linea & "Cuotas Extradord. : " & Mid(moddat_gf_Consulta_Pardes("223", g_rst_Princi!HIPMAE_CUOANO) & Space(20), 1, 18)
   r_str_Linea = r_str_Linea & "Período de Gracia : " & Mid(Format(g_rst_Princi!HIPMAE_PERGRA, "#0") & Space(18), 1, 18)
   
   Call gs_LinImp(Space(5) & r_str_Linea)
   
   r_str_Linea = ""
   r_str_Linea = r_str_Linea & "Tasa de Interes       : " & gf_FormatoNumero(g_rst_Princi!HIPMAE_TASINT, 6, 0) & "%"
   
   Call gs_LinImp(Space(5) & r_str_Linea)
   
   Call gs_LinImp(Space(5) & "Tipo de Cronograma    : " & moddat_gf_Consulta_Pardes("028", CStr(2)))
   Call gs_LinImp(Space(5) & "Monto Préstamo Tramo  : " & gf_FormatoNumero(g_rst_Princi!HIPMAE_IMPCON, 15))
   Call gs_LinImp("")
   
   
   
   Call gs_LinImp(Space(5) & String(110, "-"))
   
   r_str_Linea = ""
   r_str_Linea = r_str_Linea & "Cuota" & Space(2)
   r_str_Linea = r_str_Linea & "F. Vencimiento" & Space(2)
   r_str_Linea = r_str_Linea & "    Capital    " & Space(2)
   r_str_Linea = r_str_Linea & "    Interes    " & Space(2)
   r_str_Linea = r_str_Linea & "  Total Cuota  " & Space(2)
   r_str_Linea = r_str_Linea & " Saldo Capital " & Space(2)
   
   Call gs_LinImp(Space(5) & r_str_Linea)
   Call gs_LinImp(Space(5) & String(110, "-"))
   
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   r_int_NumLin = 22
   
   g_str_Parame = "SELECT * FROM CRE_HIPCUO WHERE "
   g_str_Parame = g_str_Parame & "HIPCUO_NUMOPE = '" & p_NumOpe & "' AND "
   g_str_Parame = g_str_Parame & "HIPCUO_TIPCRO = 2"
   g_str_Parame = g_str_Parame & "ORDER BY HIPCUO_NUMCUO ASC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If

   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst

      Do While Not g_rst_Princi.EOF
         If r_int_NumLin = 90 Then
            Call gs_LinImp("SP")
            r_int_NumLin = 1
         End If
         If r_int_NumLin = 1 Then
            Call gs_LinImp("")
            Call gs_LinImp("")
            
            Call gs_LinImp(Space(5) & String(110, "-"))
            
            r_str_Linea = ""
            r_str_Linea = r_str_Linea & "Cuota" & Space(2)
            r_str_Linea = r_str_Linea & "F. Vencimiento" & Space(2)
            r_str_Linea = r_str_Linea & "    Capital    " & Space(2)
            r_str_Linea = r_str_Linea & "    Interes    " & Space(2)
            r_str_Linea = r_str_Linea & "  Total Cuota  " & Space(2)
            r_str_Linea = r_str_Linea & " Saldo Capital " & Space(2)
            
            Call gs_LinImp(Space(5) & r_str_Linea)
            Call gs_LinImp(Space(5) & String(110, "-"))
            
            r_int_NumLin = 5
         End If
         
         r_dbl_ImpCuo = 0
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(Format(g_rst_Princi!HIPCUO_CAPITA, "###,###,##0.00"))
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(Format(g_rst_Princi!HIPCUO_INTERE, "###,###,##0.00"))

         r_dbl_Capita = r_dbl_Capita + CDbl(Format(g_rst_Princi!HIPCUO_CAPITA, "###,###,##0.00"))
         r_dbl_Intere = r_dbl_Intere + CDbl(Format(g_rst_Princi!HIPCUO_INTERE, "###,###,##0.00"))
         r_dbl_TotCuo = r_dbl_TotCuo + r_dbl_ImpCuo

         r_str_Linea = ""
         r_str_Linea = r_str_Linea & Space(1) & Format(g_rst_Princi!HIPCUO_NUMCUO, "000") & Space(1)
         r_str_Linea = r_str_Linea & Space(2)
         
         r_str_Linea = r_str_Linea & Space(2) & gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECVCT)) & Space(2)
         r_str_Linea = r_str_Linea & Space(2)
         
         r_str_Linea = r_str_Linea & gf_FormatoNumero(g_rst_Princi!HIPCUO_CAPITA, 15) & Space(2)
         r_str_Linea = r_str_Linea & gf_FormatoNumero(g_rst_Princi!HIPCUO_INTERE, 15) & Space(2)
         r_str_Linea = r_str_Linea & gf_FormatoNumero(r_dbl_ImpCuo, 15) & Space(2)
         r_str_Linea = r_str_Linea & gf_FormatoNumero(g_rst_Princi!HIPCUO_SALCAP, 15) & Space(2)
         
         Call gs_LinImp(Space(5) & r_str_Linea)

         r_int_NumLin = r_int_NumLin + 1
         
         g_rst_Princi.MoveNext
      Loop
   End If

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   
   Call gs_LinImp(Space(5) & String(110, "-"))
   
   r_str_Linea = ""
   r_str_Linea = r_str_Linea & "TOTAL GENERAL" & Space(10)
   r_str_Linea = r_str_Linea & gf_FormatoNumero(r_dbl_Capita, 15) & Space(2)
   r_str_Linea = r_str_Linea & gf_FormatoNumero(r_dbl_Intere, 15) & Space(2)
   r_str_Linea = r_str_Linea & gf_FormatoNumero(r_dbl_TotCuo, 15) & Space(2)
   
   Call gs_LinImp(Space(5) & r_str_Linea)
   Call gs_LinImp(Space(5) & String(110, "-"))
End Sub

Private Sub fs_ImpCro_CofNCo(ByVal p_NumOpe As String)
   Dim r_str_Linea      As String
   
   Dim r_dbl_Capita     As Double
   Dim r_dbl_Intere     As Double
   Dim r_dbl_Comisi     As Double
   Dim r_dbl_ImpCuo     As Double
   Dim r_dbl_TotCuo     As Double
   
   Dim r_int_NumLin     As Integer
   
   
   r_dbl_Capita = 0
   r_dbl_Intere = 0
   r_dbl_Comisi = 0
   r_dbl_TotCuo = 0
   
   'Obteniendo Información de la Operación
   g_str_Parame = "SELECT * FROM CRE_HIPMAE WHERE "
   g_str_Parame = g_str_Parame & "HIPMAE_NUMOPE = '" & p_NumOpe & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If

   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      Exit Sub
   End If
   
   'Inicializando Arreglo de Impresiones
   ReDim g_arr_Imprim(0)

   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp(Space(98) & "Fecha: " & moddat_g_str_FecSis)
   Call gs_LinImp(Space(98) & "Hora:  " & Space(2) & Format(Time, "hh:mm:ss"))
   Call gs_LinImp("")
   Call gs_LinImp(Space(38) & "CRONOGRAMA DE PAGOS - CREDITO HIPOTECARIO")
   Call gs_LinImp(Space(38) & "-----------------------------------------")
   Call gs_LinImp("")
   
   Call gs_LinImp(Space(5) & "Nro. de Operación     : " & Mid(p_NumOpe, 1, 3) & "-" & Mid(p_NumOpe, 4, 2) & "-" & Mid(p_NumOpe, 6, 5))
   Call gs_LinImp(Space(5) & "Docum. Ident. Cliente : " & CStr(g_rst_Princi!HIPMAE_TDOCLI) & "-" & Trim(g_rst_Princi!HIPMAE_NDOCLI))
   Call gs_LinImp(Space(5) & "Nombre Cliente        : " & moddat_gf_Buscar_NomCli(g_rst_Princi!HIPMAE_TDOCLI, Trim(g_rst_Princi!HIPMAE_NDOCLI)))
   Call gs_LinImp("")
   
   Call gs_LinImp(Space(5) & "Moneda de Préstamo    : " & moddat_gf_Consulta_Pardes("204", CStr(g_rst_Princi!HIPMAE_MONEDA)))
   Call gs_LinImp(Space(5) & "Total Préstamo        : " & gf_FormatoNumero(g_rst_Princi!HIPMAE_PREMPR, 15))
   Call gs_LinImp(Space(5) & "Fecha Desembolso      : " & gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECAPR)))
   
   '1er Bloque hasta 18 Caracteres / 2do Bloque Hasta 18 /3er Bloque Hasta 18
   
   r_str_Linea = ""
   r_str_Linea = r_str_Linea & "Nro. Cuotas           : " & Format(g_rst_Princi!HIPMAE_NUMCUO, "000") & Space(15)
   r_str_Linea = r_str_Linea & "Cuotas Extradord. : " & Mid(moddat_gf_Consulta_Pardes("223", g_rst_Princi!HIPMAE_CUOANO) & Space(20), 1, 18)
   r_str_Linea = r_str_Linea & "Período de Gracia : " & Mid(Format(g_rst_Princi!HIPMAE_PERGRA, "#0") & Space(18), 1, 18)
   
   Call gs_LinImp(Space(5) & r_str_Linea)
   
   r_str_Linea = ""
   r_str_Linea = r_str_Linea & "Tasa de Interes       : " & Mid(gf_FormatoNumero(g_rst_Princi!HIPMAE_TASCOF, 6, 0) & "%" & Space(18), 1, 18)
   r_str_Linea = r_str_Linea & "Comisión          : " & gf_FormatoNumero(g_rst_Princi!HIPMAE_COMCOF, 6, 0) & "%"
   
   Call gs_LinImp(Space(5) & r_str_Linea)
   
   Call gs_LinImp(Space(5) & "Tipo de Cronograma    : " & moddat_gf_Consulta_Pardes("028", CStr(3)))
   Call gs_LinImp(Space(5) & "Monto Préstamo Tramo  : " & gf_FormatoNumero(g_rst_Princi!HIPMAE_IMPNCO, 15))
   Call gs_LinImp("")
   
   
   Call gs_LinImp(Space(5) & String(110, "-"))
   
   r_str_Linea = ""
   r_str_Linea = r_str_Linea & "Cuota" & Space(2)
   r_str_Linea = r_str_Linea & "F. Vencimiento" & Space(2)
   r_str_Linea = r_str_Linea & "    Capital    " & Space(2)
   r_str_Linea = r_str_Linea & "    Interes    " & Space(2)
   r_str_Linea = r_str_Linea & "    Comisión   " & Space(2)
   r_str_Linea = r_str_Linea & "  Total Cuota  " & Space(2)
   r_str_Linea = r_str_Linea & " Saldo Capital " & Space(2)
   
   Call gs_LinImp(Space(5) & r_str_Linea)
   Call gs_LinImp(Space(5) & String(110, "-"))
   
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   r_int_NumLin = 22
   
   g_str_Parame = "SELECT * FROM CRE_HIPCUO WHERE "
   g_str_Parame = g_str_Parame & "HIPCUO_NUMOPE = '" & p_NumOpe & "' AND "
   g_str_Parame = g_str_Parame & "HIPCUO_TIPCRO = 3"
   g_str_Parame = g_str_Parame & "ORDER BY HIPCUO_NUMCUO ASC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If

   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst

      Do While Not g_rst_Princi.EOF
         If r_int_NumLin = 90 Then
            Call gs_LinImp("SP")
            r_int_NumLin = 1
         End If
         If r_int_NumLin = 1 Then
            Call gs_LinImp("")
            Call gs_LinImp("")
            
            Call gs_LinImp(Space(5) & String(110, "-"))
            
            r_str_Linea = ""
            r_str_Linea = r_str_Linea & "Cuota" & Space(2)
            r_str_Linea = r_str_Linea & "F. Vencimiento" & Space(2)
            r_str_Linea = r_str_Linea & "    Capital    " & Space(2)
            r_str_Linea = r_str_Linea & "    Interes    " & Space(2)
            r_str_Linea = r_str_Linea & "    Comisión   " & Space(2)
            r_str_Linea = r_str_Linea & "  Total Cuota  " & Space(2)
            r_str_Linea = r_str_Linea & " Saldo Capital " & Space(2)
            
            Call gs_LinImp(Space(5) & r_str_Linea)
            Call gs_LinImp(Space(5) & String(110, "-"))
            
            r_int_NumLin = 5
         End If
         
         r_dbl_ImpCuo = 0
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(Format(g_rst_Princi!HIPCUO_CAPITA, "###,###,##0.00"))
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(Format(g_rst_Princi!HIPCUO_INTERE, "###,###,##0.00"))
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(Format(g_rst_Princi!HIPCUO_COMISI, "###,###,##0.00"))

         r_dbl_Capita = r_dbl_Capita + CDbl(Format(g_rst_Princi!HIPCUO_CAPITA, "###,###,##0.00"))
         r_dbl_Intere = r_dbl_Intere + CDbl(Format(g_rst_Princi!HIPCUO_INTERE, "###,###,##0.00"))
         r_dbl_Comisi = r_dbl_Comisi + CDbl(Format(g_rst_Princi!HIPCUO_COMISI, "###,###,##0.00"))
         r_dbl_TotCuo = r_dbl_TotCuo + r_dbl_ImpCuo

         r_str_Linea = ""
         r_str_Linea = r_str_Linea & Space(1) & Format(g_rst_Princi!HIPCUO_NUMCUO, "000") & Space(1)
         r_str_Linea = r_str_Linea & Space(2)
         
         r_str_Linea = r_str_Linea & Space(2) & gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECVCT)) & Space(2)
         r_str_Linea = r_str_Linea & Space(2)
         
         r_str_Linea = r_str_Linea & gf_FormatoNumero(g_rst_Princi!HIPCUO_CAPITA, 15) & Space(2)
         r_str_Linea = r_str_Linea & gf_FormatoNumero(g_rst_Princi!HIPCUO_INTERE, 15) & Space(2)
         r_str_Linea = r_str_Linea & gf_FormatoNumero(g_rst_Princi!HIPCUO_COMISI, 15) & Space(2)
         r_str_Linea = r_str_Linea & gf_FormatoNumero(r_dbl_ImpCuo, 15) & Space(2)
         r_str_Linea = r_str_Linea & gf_FormatoNumero(g_rst_Princi!HIPCUO_SALCAP, 15) & Space(2)
         
         Call gs_LinImp(Space(5) & r_str_Linea)

         r_int_NumLin = r_int_NumLin + 1
         
         g_rst_Princi.MoveNext
      Loop
   End If

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   
   Call gs_LinImp(Space(5) & String(110, "-"))
   
   r_str_Linea = ""
   r_str_Linea = r_str_Linea & "TOTAL GENERAL" & Space(10)
   r_str_Linea = r_str_Linea & gf_FormatoNumero(r_dbl_Capita, 15) & Space(2)
   r_str_Linea = r_str_Linea & gf_FormatoNumero(r_dbl_Intere, 15) & Space(2)
   r_str_Linea = r_str_Linea & gf_FormatoNumero(r_dbl_Comisi, 15) & Space(2)
   r_str_Linea = r_str_Linea & gf_FormatoNumero(r_dbl_TotCuo, 15) & Space(2)
   
   Call gs_LinImp(Space(5) & r_str_Linea)
   Call gs_LinImp(Space(5) & String(110, "-"))
End Sub

Private Sub fs_ImpCro_CofCon(ByVal p_NumOpe As String)
   Dim r_str_Linea      As String
   
   Dim r_dbl_Capita     As Double
   Dim r_dbl_Intere     As Double
   Dim r_dbl_Comisi     As Double
   Dim r_dbl_ImpCuo     As Double
   Dim r_dbl_TotCuo     As Double
   
   Dim r_int_NumLin     As Integer
   
   
   r_dbl_Capita = 0
   r_dbl_Intere = 0
   r_dbl_Comisi = 0
   r_dbl_TotCuo = 0
   
   'Obteniendo Información de la Operación
   g_str_Parame = "SELECT * FROM CRE_HIPMAE WHERE "
   g_str_Parame = g_str_Parame & "HIPMAE_NUMOPE = '" & p_NumOpe & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If

   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      Exit Sub
   End If
   
   'Inicializando Arreglo de Impresiones
   ReDim g_arr_Imprim(0)

   Call gs_LinImp("")
   Call gs_LinImp("")
   Call gs_LinImp(Space(98) & "Fecha: " & moddat_g_str_FecSis)
   Call gs_LinImp(Space(98) & "Hora:  " & Space(2) & Format(Time, "hh:mm:ss"))
   Call gs_LinImp("")
   Call gs_LinImp(Space(38) & "CRONOGRAMA DE PAGOS - CREDITO HIPOTECARIO")
   Call gs_LinImp(Space(38) & "-----------------------------------------")
   Call gs_LinImp("")
   
   Call gs_LinImp(Space(5) & "Nro. de Operación     : " & Mid(p_NumOpe, 1, 3) & "-" & Mid(p_NumOpe, 4, 2) & "-" & Mid(p_NumOpe, 6, 5))
   Call gs_LinImp(Space(5) & "Docum. Ident. Cliente : " & CStr(g_rst_Princi!HIPMAE_TDOCLI) & "-" & Trim(g_rst_Princi!HIPMAE_NDOCLI))
   Call gs_LinImp(Space(5) & "Nombre Cliente        : " & moddat_gf_Buscar_NomCli(g_rst_Princi!HIPMAE_TDOCLI, Trim(g_rst_Princi!HIPMAE_NDOCLI)))
   Call gs_LinImp("")
   
   Call gs_LinImp(Space(5) & "Moneda de Préstamo    : " & moddat_gf_Consulta_Pardes("204", CStr(g_rst_Princi!HIPMAE_MONEDA)))
   Call gs_LinImp(Space(5) & "Total Préstamo        : " & gf_FormatoNumero(g_rst_Princi!HIPMAE_PREMPR, 15))
   Call gs_LinImp(Space(5) & "Fecha Desembolso      : " & gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECAPR)))
   
   '1er Bloque hasta 18 Caracteres / 2do Bloque Hasta 18 /3er Bloque Hasta 18
   
   r_str_Linea = ""
   r_str_Linea = r_str_Linea & "Nro. Cuotas           : " & Format(g_rst_Princi!HIPMAE_NUMCUO, "000") & Space(15)
   r_str_Linea = r_str_Linea & "Cuotas Extradord. : " & Mid(moddat_gf_Consulta_Pardes("223", g_rst_Princi!HIPMAE_CUOANO) & Space(20), 1, 18)
   r_str_Linea = r_str_Linea & "Período de Gracia : " & Mid(Format(g_rst_Princi!HIPMAE_PERGRA, "#0") & Space(18), 1, 18)
   
   Call gs_LinImp(Space(5) & r_str_Linea)
   
   r_str_Linea = ""
   r_str_Linea = r_str_Linea & "Tasa de Interes       : " & Mid(gf_FormatoNumero(g_rst_Princi!HIPMAE_TASCOF, 6, 0) & "%" & Space(18), 1, 18)
   r_str_Linea = r_str_Linea & "Comisión          : " & gf_FormatoNumero(g_rst_Princi!HIPMAE_COMCOF, 6, 0) & "%"
   
   Call gs_LinImp(Space(5) & r_str_Linea)
   
   Call gs_LinImp(Space(5) & "Tipo de Cronograma    : " & moddat_gf_Consulta_Pardes("028", CStr(4)))
   Call gs_LinImp(Space(5) & "Monto Préstamo Tramo  : " & gf_FormatoNumero(g_rst_Princi!HIPMAE_IMPNCO, 15))
   Call gs_LinImp("")
   
   
   Call gs_LinImp(Space(5) & String(110, "-"))
   
   r_str_Linea = ""
   r_str_Linea = r_str_Linea & "Cuota" & Space(2)
   r_str_Linea = r_str_Linea & "F. Vencimiento" & Space(2)
   r_str_Linea = r_str_Linea & "    Capital    " & Space(2)
   r_str_Linea = r_str_Linea & "    Interes    " & Space(2)
   r_str_Linea = r_str_Linea & "    Comisión   " & Space(2)
   r_str_Linea = r_str_Linea & "  Total Cuota  " & Space(2)
   r_str_Linea = r_str_Linea & " Saldo Capital " & Space(2)
   
   Call gs_LinImp(Space(5) & r_str_Linea)
   Call gs_LinImp(Space(5) & String(110, "-"))
   
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   r_int_NumLin = 22
   
   g_str_Parame = "SELECT * FROM CRE_HIPCUO WHERE "
   g_str_Parame = g_str_Parame & "HIPCUO_NUMOPE = '" & p_NumOpe & "' AND "
   g_str_Parame = g_str_Parame & "HIPCUO_TIPCRO = 4"
   g_str_Parame = g_str_Parame & "ORDER BY HIPCUO_NUMCUO ASC"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If

   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst

      Do While Not g_rst_Princi.EOF
         If r_int_NumLin = 90 Then
            Call gs_LinImp("SP")
            r_int_NumLin = 1
         End If
         If r_int_NumLin = 1 Then
            Call gs_LinImp("")
            Call gs_LinImp("")
            
            Call gs_LinImp(Space(5) & String(110, "-"))
            
            r_str_Linea = ""
            r_str_Linea = r_str_Linea & "Cuota" & Space(2)
            r_str_Linea = r_str_Linea & "F. Vencimiento" & Space(2)
            r_str_Linea = r_str_Linea & "    Capital    " & Space(2)
            r_str_Linea = r_str_Linea & "    Interes    " & Space(2)
            r_str_Linea = r_str_Linea & "    Comisión   " & Space(2)
            r_str_Linea = r_str_Linea & "  Total Cuota  " & Space(2)
            r_str_Linea = r_str_Linea & " Saldo Capital " & Space(2)
            
            Call gs_LinImp(Space(5) & r_str_Linea)
            Call gs_LinImp(Space(5) & String(110, "-"))
            
            r_int_NumLin = 5
         End If
         
         r_dbl_ImpCuo = 0
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(Format(g_rst_Princi!HIPCUO_CAPITA, "###,###,##0.00"))
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(Format(g_rst_Princi!HIPCUO_INTERE, "###,###,##0.00"))
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(Format(g_rst_Princi!HIPCUO_COMISI, "###,###,##0.00"))

         r_dbl_Capita = r_dbl_Capita + CDbl(Format(g_rst_Princi!HIPCUO_CAPITA, "###,###,##0.00"))
         r_dbl_Intere = r_dbl_Intere + CDbl(Format(g_rst_Princi!HIPCUO_INTERE, "###,###,##0.00"))
         r_dbl_Comisi = r_dbl_Comisi + CDbl(Format(g_rst_Princi!HIPCUO_COMISI, "###,###,##0.00"))
         r_dbl_TotCuo = r_dbl_TotCuo + r_dbl_ImpCuo

         r_str_Linea = ""
         r_str_Linea = r_str_Linea & Space(1) & Format(g_rst_Princi!HIPCUO_NUMCUO, "000") & Space(1)
         r_str_Linea = r_str_Linea & Space(2)
         
         r_str_Linea = r_str_Linea & Space(2) & gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECVCT)) & Space(2)
         r_str_Linea = r_str_Linea & Space(2)
         
         r_str_Linea = r_str_Linea & gf_FormatoNumero(g_rst_Princi!HIPCUO_CAPITA, 15) & Space(2)
         r_str_Linea = r_str_Linea & gf_FormatoNumero(g_rst_Princi!HIPCUO_INTERE, 15) & Space(2)
         r_str_Linea = r_str_Linea & gf_FormatoNumero(g_rst_Princi!HIPCUO_COMISI, 15) & Space(2)
         r_str_Linea = r_str_Linea & gf_FormatoNumero(r_dbl_ImpCuo, 15) & Space(2)
         r_str_Linea = r_str_Linea & gf_FormatoNumero(g_rst_Princi!HIPCUO_SALCAP, 15) & Space(2)
         
         Call gs_LinImp(Space(5) & r_str_Linea)

         r_int_NumLin = r_int_NumLin + 1
         
         g_rst_Princi.MoveNext
      Loop
   End If

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   
   Call gs_LinImp(Space(5) & String(110, "-"))
   
   r_str_Linea = ""
   r_str_Linea = r_str_Linea & "TOTAL GENERAL" & Space(10)
   r_str_Linea = r_str_Linea & gf_FormatoNumero(r_dbl_Capita, 15) & Space(2)
   r_str_Linea = r_str_Linea & gf_FormatoNumero(r_dbl_Intere, 15) & Space(2)
   r_str_Linea = r_str_Linea & gf_FormatoNumero(r_dbl_Comisi, 15) & Space(2)
   r_str_Linea = r_str_Linea & gf_FormatoNumero(r_dbl_TotCuo, 15) & Space(2)
   
   Call gs_LinImp(Space(5) & r_str_Linea)
   Call gs_LinImp(Space(5) & String(110, "-"))
End Sub


Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Call fs_Inicia
   
   Me.Caption = modgen_g_str_NomPlt

   pnl_NumOpe.Caption = Mid(moddat_g_str_NumOpe, 1, 3) & "-" & Mid(moddat_g_str_NumOpe, 4, 2) & "-" & Mid(moddat_g_str_NumOpe, 6, 5)
   pnl_NumSol.Caption = Mid(moddat_g_str_NumSol, 1, 3) & "-" & Mid(moddat_g_str_NumSol, 4, 3) & "-" & Mid(moddat_g_str_NumSol, 7, 2) & "-" & Mid(moddat_g_str_NumSol, 9, 4)
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   pnl_Produc.Caption = moddat_g_str_NomPrd
   pnl_Modali.Caption = moddat_g_str_DesMod

   Call fs_Carga_Cro_CliNCo
   Call fs_Carga_Cro_CliCon
   Call fs_Carga_Cro_CofNCo
   Call fs_Carga_Cro_CofCon

   Call gs_CentraForm(Me)

   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   'Cliente No Concesional
   grd_CliNCo_Listad.ColWidth(0) = 1110
   grd_CliNCo_Listad.ColWidth(1) = 1380
   grd_CliNCo_Listad.ColWidth(2) = 1395
   grd_CliNCo_Listad.ColWidth(3) = 1390
   grd_CliNCo_Listad.ColWidth(4) = 1385
   grd_CliNCo_Listad.ColWidth(5) = 1380
   grd_CliNCo_Listad.ColWidth(6) = 1370
   grd_CliNCo_Listad.ColWidth(7) = 1375
   grd_CliNCo_Listad.ColWidth(8) = 1380
   
   grd_CliNCo_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_CliNCo_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_CliNCo_Listad.ColAlignment(2) = flexAlignRightCenter
   grd_CliNCo_Listad.ColAlignment(3) = flexAlignRightCenter
   grd_CliNCo_Listad.ColAlignment(4) = flexAlignRightCenter
   grd_CliNCo_Listad.ColAlignment(5) = flexAlignRightCenter
   grd_CliNCo_Listad.ColAlignment(6) = flexAlignRightCenter
   grd_CliNCo_Listad.ColAlignment(7) = flexAlignRightCenter
   grd_CliNCo_Listad.ColAlignment(8) = flexAlignRightCenter
   
   'Cliente Concesional
   grd_CliCon_Listad.ColWidth(0) = 1170
   grd_CliCon_Listad.ColWidth(1) = 1560
   grd_CliCon_Listad.ColWidth(2) = 2355
   grd_CliCon_Listad.ColWidth(3) = 2355
   grd_CliCon_Listad.ColWidth(4) = 2355
   grd_CliCon_Listad.ColWidth(5) = 2355
   
   grd_CliCon_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_CliCon_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_CliCon_Listad.ColAlignment(2) = flexAlignRightCenter
   grd_CliCon_Listad.ColAlignment(3) = flexAlignRightCenter
   grd_CliCon_Listad.ColAlignment(4) = flexAlignRightCenter
   grd_CliCon_Listad.ColAlignment(5) = flexAlignRightCenter
   
   'Cofide No Concesional
   grd_CofNCo_Listad.ColWidth(0) = 1170
   grd_CofNCo_Listad.ColWidth(1) = 1800
   grd_CofNCo_Listad.ColWidth(2) = 1835
   grd_CofNCo_Listad.ColWidth(3) = 1835
   grd_CofNCo_Listad.ColWidth(4) = 1835
   grd_CofNCo_Listad.ColWidth(5) = 1840
   grd_CofNCo_Listad.ColWidth(6) = 1835
   
   grd_CofNCo_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_CofNCo_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_CofNCo_Listad.ColAlignment(2) = flexAlignRightCenter
   grd_CofNCo_Listad.ColAlignment(3) = flexAlignRightCenter
   grd_CofNCo_Listad.ColAlignment(4) = flexAlignRightCenter
   grd_CofNCo_Listad.ColAlignment(5) = flexAlignRightCenter
   grd_CofNCo_Listad.ColAlignment(6) = flexAlignRightCenter
   
   'Cofide Concesional
   grd_CofCon_Listad.ColWidth(0) = 1170
   grd_CofCon_Listad.ColWidth(1) = 1800
   grd_CofCon_Listad.ColWidth(2) = 1835
   grd_CofCon_Listad.ColWidth(3) = 1835
   grd_CofCon_Listad.ColWidth(4) = 1835
   grd_CofCon_Listad.ColWidth(5) = 1840
   grd_CofCon_Listad.ColWidth(6) = 1835
   
   grd_CofCon_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_CofCon_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_CofCon_Listad.ColAlignment(2) = flexAlignRightCenter
   grd_CofCon_Listad.ColAlignment(3) = flexAlignRightCenter
   grd_CofCon_Listad.ColAlignment(4) = flexAlignRightCenter
   grd_CofCon_Listad.ColAlignment(5) = flexAlignRightCenter
   grd_CofCon_Listad.ColAlignment(6) = flexAlignRightCenter
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

   g_str_Parame = "SELECT * FROM CRE_HIPCUO WHERE "
   g_str_Parame = g_str_Parame & "HIPCUO_NUMOPE = '" & moddat_g_str_NumOpe & "' AND "
   g_str_Parame = g_str_Parame & "HIPCUO_TIPCRO = 1 "
   
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

   g_str_Parame = "SELECT * FROM CRE_HIPCUO WHERE "
   g_str_Parame = g_str_Parame & "HIPCUO_NUMOPE = '" & moddat_g_str_NumOpe & "' AND "
   g_str_Parame = g_str_Parame & "HIPCUO_TIPCRO = 2 "
   
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

   g_str_Parame = "SELECT * FROM CRE_HIPCUO WHERE "
   g_str_Parame = g_str_Parame & "HIPCUO_NUMOPE = '" & moddat_g_str_NumOpe & "' AND "
   g_str_Parame = g_str_Parame & "HIPCUO_TIPCRO = 3 "
   
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
         grd_CofNCo_Listad.Text = Format(g_rst_Princi!HIPCUO_COMISI, "###,###,##0.00")
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(grd_CofNCo_Listad.Text)
         
         grd_CofNCo_Listad.Col = 5
         grd_CofNCo_Listad.Text = Format(r_dbl_ImpCuo, "###,###,##0.00")
         
         grd_CofNCo_Listad.Col = 6
         grd_CofNCo_Listad.Text = Format(g_rst_Princi!HIPCUO_SALCAP, "###,###,##0.00")

         r_dbl_Capita = r_dbl_Capita + CDbl(Format(g_rst_Princi!HIPCUO_CAPITA, "###,###,##0.00"))
         r_dbl_Intere = r_dbl_Intere + CDbl(Format(g_rst_Princi!HIPCUO_INTERE, "###,###,##0.00"))
         r_dbl_Comisi = r_dbl_Comisi + CDbl(Format(g_rst_Princi!HIPCUO_COMISI, "###,###,##0.00"))
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

Private Sub fs_Carga_Cro_CofCon()
   Dim r_dbl_Capita     As Double
   Dim r_dbl_Intere     As Double
   Dim r_dbl_Comisi     As Double
   Dim r_dbl_ImpCuo     As Double
   Dim r_dbl_TotCuo     As Double
   
   Call gs_LimpiaGrid(grd_CofCon_Listad)

   r_dbl_Capita = 0
   r_dbl_Intere = 0
   r_dbl_Comisi = 0
   r_dbl_TotCuo = 0

   g_str_Parame = "SELECT * FROM CRE_HIPCUO WHERE "
   g_str_Parame = g_str_Parame & "HIPCUO_NUMOPE = '" & moddat_g_str_NumOpe & "' AND "
   g_str_Parame = g_str_Parame & "HIPCUO_TIPCRO = 4 "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_CofCon_Listad.Redraw = False
      
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         grd_CofCon_Listad.Rows = grd_CofCon_Listad.Rows + 1
         grd_CofCon_Listad.Row = grd_CofCon_Listad.Rows - 1
         
         r_dbl_ImpCuo = 0
         
         grd_CofCon_Listad.Col = 0
         grd_CofCon_Listad.Text = Format(g_rst_Princi!HIPCUO_NUMCUO, "000")
      
         grd_CofCon_Listad.Col = 1
         grd_CofCon_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECVCT))
         
         grd_CofCon_Listad.Col = 2
         grd_CofCon_Listad.Text = Format(g_rst_Princi!HIPCUO_CAPITA, "###,###,##0.00")
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(grd_CofCon_Listad.Text)
         
         grd_CofCon_Listad.Col = 3
         grd_CofCon_Listad.Text = Format(g_rst_Princi!HIPCUO_INTERE, "###,###,##0.00")
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(grd_CofCon_Listad.Text)
         
         grd_CofCon_Listad.Col = 4
         grd_CofCon_Listad.Text = Format(g_rst_Princi!HIPCUO_COMISI, "###,###,##0.00")
         r_dbl_ImpCuo = r_dbl_ImpCuo + CDbl(grd_CofCon_Listad.Text)
         
         grd_CofCon_Listad.Col = 5
         grd_CofCon_Listad.Text = Format(r_dbl_ImpCuo, "###,###,##0.00")
         
         grd_CofCon_Listad.Col = 6
         grd_CofCon_Listad.Text = Format(g_rst_Princi!HIPCUO_SALCAP, "###,###,##0.00")

         r_dbl_Capita = r_dbl_Capita + CDbl(Format(g_rst_Princi!HIPCUO_CAPITA, "###,###,##0.00"))
         r_dbl_Intere = r_dbl_Intere + CDbl(Format(g_rst_Princi!HIPCUO_INTERE, "###,###,##0.00"))
         r_dbl_Comisi = r_dbl_Comisi + CDbl(Format(g_rst_Princi!HIPCUO_COMISI, "###,###,##0.00"))
         r_dbl_TotCuo = r_dbl_TotCuo + CDbl(Format(r_dbl_ImpCuo, "###,###,##0.00"))
            
         g_rst_Princi.MoveNext
      Loop
      
      grd_CofCon_Listad.Redraw = True
      
      Call gs_UbiIniGrid(grd_CofCon_Listad)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   pnl_CofCon_Capita.Caption = Format(r_dbl_Capita, "###,###,##0.00") & " "
   pnl_CofCon_Intere.Caption = Format(r_dbl_Intere, "###,###,##0.00") & " "
   pnl_CofCon_Comisi.Caption = Format(r_dbl_Comisi, "###,###,##0.00") & " "
   pnl_CofCon_TotCuo.Caption = Format(r_dbl_TotCuo, "###,###,##0.00") & " "
End Sub



