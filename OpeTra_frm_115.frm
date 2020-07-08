VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frm_ActOpe_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   9945
   ClientLeft      =   1920
   ClientTop       =   840
   ClientWidth     =   11610
   Icon            =   "OpeTra_frm_115.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9945
   ScaleWidth      =   11610
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   9945
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11640
      _Version        =   65536
      _ExtentX        =   20532
      _ExtentY        =   17542
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
         Height          =   675
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
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
            Height          =   285
            Left            =   630
            TabIndex        =   2
            Top             =   30
            Width           =   6345
            _Version        =   65536
            _ExtentX        =   11192
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Créditos Hipotecarios"
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
         Begin Threed.SSPanel SSPanel15 
            Height          =   285
            Left            =   630
            TabIndex        =   3
            Top             =   330
            Width           =   6345
            _Version        =   65536
            _ExtentX        =   11192
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Activación de Crédito Hipotecario"
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
         Begin MSMAPI.MAPIMessages mps_Mensaj 
            Left            =   10920
            Top             =   30
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            AddressEditFieldCount=   1
            AddressModifiable=   0   'False
            AddressResolveUI=   0   'False
            FetchSorted     =   0   'False
            FetchUnreadOnly =   0   'False
         End
         Begin MSMAPI.MAPISession mps_Sesion 
            Left            =   10350
            Top             =   30
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            DownloadMail    =   -1  'True
            LogonUI         =   -1  'True
            NewSession      =   0   'False
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   60
            Picture         =   "OpeTra_frm_115.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   30
         TabIndex        =   4
         Top             =   1440
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
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
         Begin Threed.SSPanel pnl_NumSol 
            Height          =   315
            Left            =   1440
            TabIndex        =   5
            Top             =   60
            Width           =   2535
            _Version        =   65536
            _ExtentX        =   4471
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
         End
         Begin Threed.SSPanel pnl_FecIng 
            Height          =   315
            Left            =   10050
            TabIndex        =   50
            Top             =   60
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2514
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
         End
         Begin Threed.SSPanel pnl_Client 
            Height          =   315
            Left            =   1440
            TabIndex        =   52
            Top             =   390
            Width           =   10035
            _Version        =   65536
            _ExtentX        =   17701
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "1-07521154 / IKEHARA PUNK MIGUEL ANGEL"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
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
         Begin VB.Label Label20 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   60
            TabIndex        =   53
            Top             =   390
            Width           =   1125
         End
         Begin VB.Label Label3 
            Caption         =   "F. Ingreso Solicitud:"
            Height          =   315
            Left            =   8400
            TabIndex        =   51
            Top             =   60
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Nro. Solicitud"
            Height          =   315
            Left            =   60
            TabIndex        =   6
            Top             =   60
            Width           =   1335
         End
      End
      Begin Threed.SSPanel SSPanel39 
         Height          =   645
         Left            =   30
         TabIndex        =   7
         Top             =   750
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
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
         Begin VB.CommandButton cmd_Aprueb 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_115.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Aprobar Solicitud"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10920
            Picture         =   "OpeTra_frm_115.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   7635
         Left            =   30
         TabIndex        =   10
         Top             =   2250
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   13467
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
            Height          =   7515
            Left            =   60
            TabIndex        =   11
            Top             =   60
            Width           =   11385
            _ExtentX        =   20082
            _ExtentY        =   13256
            _Version        =   393216
            Style           =   1
            Tabs            =   12
            TabsPerRow      =   6
            TabHeight       =   520
            TabCaption(0)   =   "Datos del Cliente"
            TabPicture(0)   =   "OpeTra_frm_115.frx":0A62
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "grd_Listad(0)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Datos del Cónyuge"
            TabPicture(1)   =   "OpeTra_frm_115.frx":0A7E
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "grd_Listad(1)"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Datos del Patrimonio"
            TabPicture(2)   =   "OpeTra_frm_115.frx":0A9A
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "grd_Listad(4)"
            Tab(2).ControlCount=   1
            TabCaption(3)   =   "Referencias Personales"
            TabPicture(3)   =   "OpeTra_frm_115.frx":0AB6
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "grd_Listad(3)"
            Tab(3).ControlCount=   1
            TabCaption(4)   =   "Datos del Inmueble"
            TabPicture(4)   =   "OpeTra_frm_115.frx":0AD2
            Tab(4).ControlEnabled=   0   'False
            Tab(4).Control(0)=   "grd_Listad(2)"
            Tab(4).ControlCount=   1
            TabCaption(5)   =   "Datos del Crédito"
            TabPicture(5)   =   "OpeTra_frm_115.frx":0AEE
            Tab(5).ControlEnabled=   0   'False
            Tab(5).Control(0)=   "grd_Listad(5)"
            Tab(5).Control(1)=   "grd_Listad(10)"
            Tab(5).Control(2)=   "SSPanel13"
            Tab(5).Control(3)=   "Label10"
            Tab(5).ControlCount=   4
            TabCaption(6)   =   "Gastos Administ."
            TabPicture(6)   =   "OpeTra_frm_115.frx":0B0A
            Tab(6).ControlEnabled=   0   'False
            Tab(6).Control(0)=   "Label8"
            Tab(6).Control(0).Enabled=   0   'False
            Tab(6).Control(1)=   "SSPanel9"
            Tab(6).Control(1).Enabled=   0   'False
            Tab(6).Control(2)=   "SSPanel12"
            Tab(6).Control(2).Enabled=   0   'False
            Tab(6).Control(3)=   "SSPanel10"
            Tab(6).Control(3).Enabled=   0   'False
            Tab(6).Control(4)=   "SSPanel8"
            Tab(6).Control(4).Enabled=   0   'False
            Tab(6).Control(5)=   "SSPanel11"
            Tab(6).Control(5).Enabled=   0   'False
            Tab(6).Control(6)=   "pnl_TotGas"
            Tab(6).Control(6).Enabled=   0   'False
            Tab(6).Control(7)=   "grd_GasAdm"
            Tab(6).Control(7).Enabled=   0   'False
            Tab(6).ControlCount=   8
            TabCaption(7)   =   "Evaluación Crediticia"
            TabPicture(7)   =   "OpeTra_frm_115.frx":0B26
            Tab(7).ControlEnabled=   0   'False
            Tab(7).Control(0)=   "grd_Listad(6)"
            Tab(7).ControlCount=   1
            TabCaption(8)   =   "Tasación del Inmueble"
            TabPicture(8)   =   "OpeTra_frm_115.frx":0B42
            Tab(8).ControlEnabled=   0   'False
            Tab(8).Control(0)=   "grd_Listad(7)"
            Tab(8).Control(1)=   "grd_Listad(11)"
            Tab(8).Control(2)=   "SSPanel14"
            Tab(8).Control(3)=   "Label11"
            Tab(8).ControlCount=   4
            TabCaption(9)   =   "Evaluación de Seguros"
            TabPicture(9)   =   "OpeTra_frm_115.frx":0B5E
            Tab(9).ControlEnabled=   0   'False
            Tab(9).Control(0)=   "txt_ObsSeg"
            Tab(9).Control(1)=   "grd_Listad(8)"
            Tab(9).Control(2)=   "SSPanel5"
            Tab(9).Control(3)=   "Label7"
            Tab(9).ControlCount=   4
            TabCaption(10)  =   "Evaluación Legal"
            TabPicture(10)  =   "OpeTra_frm_115.frx":0B7A
            Tab(10).ControlEnabled=   0   'False
            Tab(10).Control(0)=   "txt_ComCre"
            Tab(10).Control(1)=   "txt_InfLeg"
            Tab(10).Control(2)=   "SSPanel3"
            Tab(10).Control(3)=   "grd_Listad(9)"
            Tab(10).Control(4)=   "SSPanel4"
            Tab(10).Control(5)=   "Label5"
            Tab(10).Control(6)=   "Label4"
            Tab(10).Control(7)=   "Label9"
            Tab(10).ControlCount=   8
            TabCaption(11)  =   "Mivivienda / Cofide"
            TabPicture(11)  =   "OpeTra_frm_115.frx":0B96
            Tab(11).ControlEnabled=   0   'False
            Tab(11).Control(0)=   "Label12"
            Tab(11).Control(0).Enabled=   0   'False
            Tab(11).Control(1)=   "SSPanel17"
            Tab(11).Control(1).Enabled=   0   'False
            Tab(11).Control(2)=   "grd_Listad(12)"
            Tab(11).Control(2).Enabled=   0   'False
            Tab(11).Control(3)=   "txt_ObsMVi"
            Tab(11).Control(3).Enabled=   0   'False
            Tab(11).ControlCount=   4
            Begin VB.TextBox txt_ComCre 
               Height          =   1185
               Left            =   -74940
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   15
               Text            =   "OpeTra_frm_115.frx":0BB2
               Top             =   3960
               Width           =   11085
            End
            Begin VB.TextBox txt_InfLeg 
               Height          =   2535
               Left            =   -74940
               MaxLength       =   8000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   14
               Text            =   "OpeTra_frm_115.frx":0BB6
               Top             =   960
               Width           =   11085
            End
            Begin VB.TextBox txt_ObsSeg 
               Height          =   1995
               Left            =   -74940
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   13
               Text            =   "OpeTra_frm_115.frx":0BBA
               Top             =   5430
               Width           =   11085
            End
            Begin VB.TextBox txt_ObsMVi 
               Height          =   1155
               Left            =   -74940
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   12
               Text            =   "OpeTra_frm_115.frx":0BBE
               Top             =   6270
               Width           =   11085
            End
            Begin Threed.SSPanel SSPanel3 
               Height          =   60
               Left            =   -74970
               TabIndex        =   16
               Top             =   5220
               Width           =   11145
               _Version        =   65536
               _ExtentX        =   19659
               _ExtentY        =   106
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
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   6765
               Index           =   0
               Left            =   60
               TabIndex        =   17
               Top             =   690
               Width           =   11085
               _ExtentX        =   19553
               _ExtentY        =   11933
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   6765
               Index           =   2
               Left            =   -74940
               TabIndex        =   18
               Top             =   690
               Width           =   11085
               _ExtentX        =   19553
               _ExtentY        =   11933
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   6765
               Index           =   3
               Left            =   -74940
               TabIndex        =   19
               Top             =   690
               Width           =   11085
               _ExtentX        =   19553
               _ExtentY        =   11933
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   6765
               Index           =   4
               Left            =   -74940
               TabIndex        =   20
               Top             =   690
               Width           =   11085
               _ExtentX        =   19553
               _ExtentY        =   11933
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   6765
               Index           =   1
               Left            =   -74940
               TabIndex        =   21
               Top             =   690
               Width           =   11085
               _ExtentX        =   19553
               _ExtentY        =   11933
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   4485
               Index           =   5
               Left            =   -74940
               TabIndex        =   22
               Top             =   690
               Width           =   11085
               _ExtentX        =   19553
               _ExtentY        =   7911
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   6765
               Index           =   6
               Left            =   -74940
               TabIndex        =   23
               Top             =   690
               Width           =   11085
               _ExtentX        =   19553
               _ExtentY        =   11933
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   4485
               Index           =   7
               Left            =   -74940
               TabIndex        =   24
               Top             =   690
               Width           =   11085
               _ExtentX        =   19553
               _ExtentY        =   7911
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   4275
               Index           =   8
               Left            =   -74940
               TabIndex        =   25
               Top             =   690
               Width           =   11085
               _ExtentX        =   19553
               _ExtentY        =   7541
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   1815
               Index           =   9
               Left            =   -74940
               TabIndex        =   26
               Top             =   5610
               Width           =   11115
               _ExtentX        =   19606
               _ExtentY        =   3201
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel SSPanel4 
               Height          =   60
               Left            =   -74970
               TabIndex        =   27
               Top             =   3570
               Width           =   11145
               _Version        =   65536
               _ExtentX        =   19659
               _ExtentY        =   106
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
            End
            Begin Threed.SSPanel SSPanel5 
               Height          =   60
               Left            =   -74970
               TabIndex        =   28
               Top             =   5010
               Width           =   11145
               _Version        =   65536
               _ExtentX        =   19659
               _ExtentY        =   106
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
            End
            Begin MSFlexGridLib.MSFlexGrid grd_GasAdm 
               Height          =   6045
               Left            =   -74970
               TabIndex        =   29
               Top             =   990
               Width           =   11115
               _ExtentX        =   19606
               _ExtentY        =   10663
               _Version        =   393216
               Rows            =   21
               Cols            =   5
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel pnl_TotGas 
               Height          =   315
               Left            =   -65100
               TabIndex        =   30
               Top             =   7080
               Width           =   1215
               _Version        =   65536
               _ExtentX        =   2143
               _ExtentY        =   556
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
               Outline         =   -1  'True
               Alignment       =   4
            End
            Begin Threed.SSPanel SSPanel11 
               Height          =   285
               Left            =   -74940
               TabIndex        =   31
               Top             =   690
               Width           =   3975
               _Version        =   65536
               _ExtentX        =   7011
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Concepto"
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
               Left            =   -71010
               TabIndex        =   32
               Top             =   690
               Width           =   2385
               _Version        =   65536
               _ExtentX        =   4207
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Tipo de Moneda"
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
               Left            =   -68670
               TabIndex        =   33
               Top             =   690
               Width           =   1215
               _Version        =   65536
               _ExtentX        =   2143
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Importe"
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
               Left            =   -67470
               TabIndex        =   34
               Top             =   690
               Width           =   1965
               _Version        =   65536
               _ExtentX        =   3466
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
            Begin Threed.SSPanel SSPanel9 
               Height          =   285
               Left            =   -65520
               TabIndex        =   35
               Top             =   690
               Width           =   1365
               _Version        =   65536
               _ExtentX        =   2408
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Fecha Pago"
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
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   1845
               Index           =   10
               Left            =   -74940
               TabIndex        =   36
               Top             =   5610
               Width           =   11085
               _ExtentX        =   19553
               _ExtentY        =   3254
               _Version        =   393216
               Rows            =   21
               Cols            =   1
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel SSPanel13 
               Height          =   60
               Left            =   -74970
               TabIndex        =   37
               Top             =   5220
               Width           =   11145
               _Version        =   65536
               _ExtentX        =   19659
               _ExtentY        =   106
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
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   1845
               Index           =   11
               Left            =   -74940
               TabIndex        =   38
               Top             =   5610
               Width           =   11085
               _ExtentX        =   19553
               _ExtentY        =   3254
               _Version        =   393216
               Rows            =   21
               Cols            =   1
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel SSPanel14 
               Height          =   60
               Left            =   -74940
               TabIndex        =   39
               Top             =   5220
               Width           =   11145
               _Version        =   65536
               _ExtentX        =   19659
               _ExtentY        =   106
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
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   5025
               Index           =   12
               Left            =   -74940
               TabIndex        =   40
               Top             =   690
               Width           =   11085
               _ExtentX        =   19553
               _ExtentY        =   8864
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel SSPanel17 
               Height          =   60
               Left            =   -74940
               TabIndex        =   41
               Top             =   5790
               Width           =   11145
               _Version        =   65536
               _ExtentX        =   19659
               _ExtentY        =   106
               _StockProps     =   15
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.26
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   1
            End
            Begin VB.Label Label11 
               Caption         =   "Documentos Recibidos"
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
               Left            =   -74940
               TabIndex        =   49
               Top             =   5340
               Width           =   2805
            End
            Begin VB.Label Label10 
               Caption         =   "Documentos Recibidos"
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
               Left            =   -74940
               TabIndex        =   48
               Top             =   5340
               Width           =   2805
            End
            Begin VB.Label Label8 
               Caption         =   "Total de Gastos:"
               Height          =   315
               Left            =   -66480
               TabIndex        =   47
               Top             =   7080
               Width           =   1335
            End
            Begin VB.Label Label7 
               Caption         =   "Observaciones"
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
               Left            =   -74940
               TabIndex        =   46
               Top             =   5130
               Width           =   2805
            End
            Begin VB.Label Label5 
               Caption         =   "Contratos y Bloqueo Registral"
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
               Left            =   -74940
               TabIndex        =   45
               Top             =   5340
               Width           =   2805
            End
            Begin VB.Label Label4 
               Caption         =   "Comité de Créditos"
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
               Left            =   -74940
               TabIndex        =   44
               Top             =   3690
               Width           =   2805
            End
            Begin VB.Label Label9 
               Caption         =   "Informe Legal"
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
               Left            =   -74940
               TabIndex        =   43
               Top             =   690
               Width           =   2805
            End
            Begin VB.Label Label12 
               Caption         =   "Observaciones"
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
               Left            =   -74940
               TabIndex        =   42
               Top             =   5910
               Width           =   2805
            End
         End
      End
   End
End
Attribute VB_Name = "frm_ActOpe_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_str_UbiGeo_Inm    As String
Dim l_int_PryMCs_Inm    As Integer
Dim l_str_PryCod_Inm    As String
Dim l_int_PlaAno_Cre    As Integer
Dim l_int_CuoExt_Cre    As Integer
Dim l_int_PerGra_Cre    As Integer
Dim l_int_DiaPag_Cre    As Integer
Dim l_dbl_IntCap_Cre    As Double
Dim l_str_FecCom_Leg    As String
Dim l_dbl_TCaSBS_Leg    As Double
Dim l_dbl_CVtDol_Cre    As Double
Dim l_dbl_ApoDol_Cre    As Double
Dim l_dbl_CVtSol_Cre    As Double
Dim l_dbl_ApoSol_Cre    As Double
Dim l_dbl_MtoPre_Cre    As Double
Dim l_dbl_TasInt_Cre    As Double
Dim l_str_ESgDes_Seg    As String
Dim l_int_TipSeg_Seg    As Integer
Dim l_int_AplDes_Seg    As Integer
Dim l_dbl_FoIDes_Seg    As Double
Dim l_int_AplViv_Seg    As Integer
Dim l_dbl_FoIViv_Seg    As Double
Dim l_int_PriViv_Inm    As Integer
Dim l_str_OpeMVi        As String
Dim l_str_OpeMv1        As String
Dim l_str_FecCof        As String

Private Sub cmd_Aprueb_Click()
   Dim r_str_NumOpe     As String
   Dim r_int_CodCla_Prd As Integer
   Dim r_int_IndITF_Prd As Integer
   Dim r_dbl_Portes_Prd As Double

   If MsgBox("¿Está seguro de generar la Operación?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   'Obteniendo Información de Producto
   r_int_CodCla_Prd = 0
   r_int_IndITF_Prd = 0
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM CRE_PRODUC "
   g_str_Parame = g_str_Parame & " WHERE PRODUC_CODIGO = '" & moddat_g_str_CodPrd & "'"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      r_int_CodCla_Prd = g_rst_Princi!PRODUC_CODCLA
      r_int_IndITF_Prd = g_rst_Princi!PRODUC_INDITF
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Obteniendo Portes
   r_dbl_Portes_Prd = 0
   If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), moddat_g_str_CodPrd, moddat_g_str_CodSub, "002", "401") Then
      r_dbl_Portes_Prd = moddat_g_arr_Genera(1).Genera_Cantid
   End If
   
   'Generando Operación
   r_str_NumOpe = ff_Genera_NumOpe()
   
   'Grabando Cabecera de Credito
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   'Validación de los datos del inmueble
   If fs_Valida_Datos_Credito(moddat_g_str_NumSol, CStr(moddat_g_int_TipMon), IIf(CStr(moddat_g_int_TipMon) = 1, l_dbl_CVtSol_Cre, l_dbl_CVtSol_Cre), IIf(CStr(moddat_g_int_TipMon) = 1, l_dbl_ApoSol_Cre, l_dbl_ApoDol_Cre), l_dbl_MtoPre_Cre) Then

      Do While moddat_g_int_FlgGOK = False
         g_str_Parame = "USP_CRE_HIPMAE_CREA ("
         g_str_Parame = g_str_Parame & "'" & r_str_NumOpe & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumSol & "', "
         g_str_Parame = g_str_Parame & CStr(moddat_g_int_TipDoc) & ", "
         g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumDoc & "', "
         g_str_Parame = g_str_Parame & CStr(moddat_g_int_CygTDo) & ", "
         g_str_Parame = g_str_Parame & "'" & moddat_g_str_CygNDo & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_g_str_CodPrd & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_g_str_CodSub & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_g_str_CodMod & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_g_str_CodConHip & "', "
         g_str_Parame = g_str_Parame & "'" & moddat_g_str_CodEjeSeg & "', "
         g_str_Parame = g_str_Parame & "'" & l_str_UbiGeo_Inm & "', "
         g_str_Parame = g_str_Parame & CStr(l_int_PryMCs_Inm) & ", "
         g_str_Parame = g_str_Parame & "'" & l_str_PryCod_Inm & "', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
         g_str_Parame = g_str_Parame & CStr(r_int_CodCla_Prd) & ", "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "1, "
         g_str_Parame = g_str_Parame & "1, "
         g_str_Parame = g_str_Parame & CStr(l_int_PlaAno_Cre) & ", "
         g_str_Parame = g_str_Parame & CStr(l_int_CuoExt_Cre) & ", "
         g_str_Parame = g_str_Parame & CStr(l_int_PerGra_Cre) & ", "
         g_str_Parame = g_str_Parame & CStr(l_int_DiaPag_Cre) & ", "
         g_str_Parame = g_str_Parame & CStr(l_int_PlaAno_Cre * 12) & ", "
         g_str_Parame = g_str_Parame & Format(CDate(l_str_FecCom_Leg), "yyyymmdd") & ", "
         
         'Fecha Activación Crédito (Si Producto es con Recursos Mivivienda Fecha de Activación = Fecha de Desembolso COFIDE)
         If InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_AgrTFMV, moddat_g_str_CodPrd) > 0 Then 'moddat_g_str_CodPrd = "003" Or moddat_g_str_CodPrd = "004" Or moddat_g_str_CodPrd = "007" Or moddat_g_str_CodPrd = "009" Or moddat_g_str_CodPrd = "010" Or moddat_g_str_CodPrd = "013" Or moddat_g_str_CodPrd = "014" Or moddat_g_str_CodPrd = "015" Or moddat_g_str_CodPrd = "016" Or moddat_g_str_CodPrd = "017" Or moddat_g_str_CodPrd = "018" Or moddat_g_str_CodPrd = "019" Or moddat_g_str_CodPrd = "020" Or moddat_g_str_CodPrd = "021" Or moddat_g_str_CodPrd = "022" Or moddat_g_str_CodPrd = "023" Then
            g_str_Parame = g_str_Parame & Format(CDate(l_str_FecCof), "yyyymmdd") & ", "
         Else
            g_str_Parame = g_str_Parame & Format(date, "yyyymmdd") & ", "
         End If
         
         g_str_Parame = g_str_Parame & "1, "
         g_str_Parame = g_str_Parame & CStr(moddat_g_int_TipMon) & ", "
         g_str_Parame = g_str_Parame & CStr(l_dbl_CVtDol_Cre) & ", "
         g_str_Parame = g_str_Parame & CStr(l_dbl_ApoDol_Cre) & ", "
         g_str_Parame = g_str_Parame & CStr(l_dbl_CVtSol_Cre) & ", "
         g_str_Parame = g_str_Parame & CStr(l_dbl_ApoSol_Cre) & ", "
         g_str_Parame = g_str_Parame & CStr(l_dbl_MtoPre_Cre) & ", "
         g_str_Parame = g_str_Parame & CStr(l_dbl_IntCap_Cre) & ", "
         g_str_Parame = g_str_Parame & CStr(l_dbl_MtoPre_Cre) & ", "
         g_str_Parame = g_str_Parame & CStr(l_dbl_TasInt_Cre) & ", "
         g_str_Parame = g_str_Parame & "'" & l_str_ESgDes_Seg & "', "
         g_str_Parame = g_str_Parame & CStr(l_int_TipSeg_Seg) & ", "
         g_str_Parame = g_str_Parame & "'" & l_str_ESgDes_Seg & "', "
         g_str_Parame = g_str_Parame & CStr(l_int_AplDes_Seg) & ", "
         g_str_Parame = g_str_Parame & CStr(l_dbl_FoIDes_Seg) & ", "
         g_str_Parame = g_str_Parame & CStr(l_int_AplViv_Seg) & ", "
         g_str_Parame = g_str_Parame & CStr(l_dbl_FoIViv_Seg) & ", "
         g_str_Parame = g_str_Parame & CStr(r_dbl_Portes_Prd) & ", "
         g_str_Parame = g_str_Parame & "'" & l_str_OpeMVi & "', "          'Operacióm Mivivienda
         g_str_Parame = g_str_Parame & CStr(l_dbl_MtoPre_Cre) & ", "       'Monto Préstamo Mivivienda
         
         If InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd) > 0 Then 'moddat_g_str_CodPrd = "002" Or moddat_g_str_CodPrd = "011" Or moddat_g_str_CodPrd = "019" Or moddat_g_str_CodPrd = "020" Or moddat_g_str_CodPrd = "021" Or moddat_g_str_CodPrd = "022" Or moddat_g_str_CodPrd = "023" Then
            g_str_Parame = g_str_Parame & "0, "                            'Flag de Bono Buen Pagador
         Else
            g_str_Parame = g_str_Parame & "1, "                            'Flag de Bono Buen Pagador
         End If
         
         g_str_Parame = g_str_Parame & "0, "                               'Tasa de Interés Mivivienda
         g_str_Parame = g_str_Parame & "0, "                               'Tasa de Comisión COFIDE
         g_str_Parame = g_str_Parame & "0, "                               'Tasa de Comisión CRC
         g_str_Parame = g_str_Parame & "0, "                               'Tasa de Comisión PBP
         g_str_Parame = g_str_Parame & "0, "                               'Importe Tramo No Concesional
         g_str_Parame = g_str_Parame & "0, "                               'Importe Tramo Concesional
         g_str_Parame = g_str_Parame & CStr(r_int_IndITF_Prd) & ", "
         g_str_Parame = g_str_Parame & CStr(l_dbl_TCaSBS_Leg) & ", "
         g_str_Parame = g_str_Parame & "'" & l_str_OpeMv1 & "', "          'Operacióm Mivivienda CME
         
         'Datos de Auditoria
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "   'Código Usuario
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "   'Nombre Terminal
         g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "    'Nombre Ejecutable
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "   'Código Sucursal
         
         If InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd) > 0 Then 'moddat_g_str_CodPrd = "002" Or moddat_g_str_CodPrd = "011" Then
            g_str_Parame = g_str_Parame & "'" & "999999" & "', "           'Linea de Garantia
         Else
            g_str_Parame = g_str_Parame & "'" & "000001" & "', "           'Linea de Garantia
         End If
         
         g_str_Parame = g_str_Parame & "'" & 1 & "', "                     'situacion Actual
         g_str_Parame = g_str_Parame & "'" & 13 & "', "                    'Nueva Clasificacion SBS
         g_str_Parame = g_str_Parame & CStr(l_int_PriViv_Inm) & ")"        'Primera Vivienda
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
            moddat_g_int_CntErr = moddat_g_int_CntErr + 1
         Else
            moddat_g_int_FlgGOK = True
         End If
      
         If moddat_g_int_CntErr = 6 Then
            If MsgBox("No se pudo completar el procedimiento. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
               Exit Sub
            Else
               moddat_g_int_CntErr = 0
            End If
         End If
      Loop
      
      'Creando Nueva Ocurrencia en Detalle de Seguimiento
      If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, 81, 81, 0, "", 0, 0) Then
         Exit Sub
      End If
      
      'Grabando en Seguimiento de Desembolso
      moddat_g_int_FlgGOK = False
      moddat_g_int_CntErr = 0
      
      Do While moddat_g_int_FlgGOK = False
         Screen.MousePointer = 11
         Call moddat_gs_FecSis
         
         g_str_Parame = "USP_TRA_SEGDES_FECGEN ("
         g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumSol & "', "
         g_str_Parame = g_str_Parame & Format(date, "yyyymmdd") & ", "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
         g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
            moddat_g_int_CntErr = moddat_g_int_CntErr + 1
         Else
            moddat_g_int_FlgGOK = True
         End If
         
         If moddat_g_int_CntErr = 6 Then
            If MsgBox("No se pudo completar la grabación de los datos. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
               Exit Sub
            Else
               moddat_g_int_CntErr = 0
            End If
         End If
         
         Screen.MousePointer = 0
      Loop
      
      'Enviando Correo Electrónico
      'modgen_g_str_Mail_Asunto = "GENERACION DE OPERACION CREDITICIA (" & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & " - " & Format(Time, "hh:mm:ss") & ")"
      'modgen_g_str_Mail_Mensaj = ""
      'modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NUMERO DE SOLICITUD : " & pnl_NumSol.Caption & Chr(13)
      'modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "ID CLIENTE          : " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & Chr(13)
      'modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NOMBRE CLIENTE      : " & moddat_g_str_NomCli & Chr(13)
      'modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & Chr(13)
      'Call fs_Envia_CorreoEle(modgen_g_str_Mail_Asunto, modgen_g_str_Mail_Mensaj, moddat_g_str_CodEjeSeg, "", "", 0, False, False, False)
   Else
      MsgBox "Los datos del Crédito no son correctos", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   MsgBox "Se generó la Operación Crediticia. El Número de Operacion es el " & Left(r_str_NumOpe, 3) & "-" & Mid(r_str_NumOpe, 4, 2) & "-" & Right(r_str_NumOpe, 5) & ".", vbInformation, modgen_g_str_NomPlt
   moddat_g_int_FlgAct = 2
   Unload Me
End Sub
Function fs_Valida_Datos_Credito(ByVal p_NumSol As String, ByVal p_TipMon As Integer, ByVal p_ValVta As Double, ByVal p_ApoPro As Double, ByVal p_MtoPre As Double) As Boolean
   fs_Valida_Datos_Credito = False
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT SOLMAE_COMVTA_DOL, SOLMAE_COMVTA_SOL, SOLMAE_APOPRO_DOL, SOLMAE_APOPRO_SOL, SOLMAE_MTOPRE_MPR "
   g_str_Parame = g_str_Parame & "   FROM CRE_SOLMAE "
   g_str_Parame = g_str_Parame & "  WHERE SOLMAE_NUMERO = '" & p_NumSol & "' "
   g_str_Parame = g_str_Parame & "    AND SOLMAE_MTOPRE_MPR = " & p_MtoPre & ""
   If p_TipMon = 1 Then
      g_str_Parame = g_str_Parame & "    AND SOLMAE_COMVTA_SOL = " & p_ValVta & ""
      g_str_Parame = g_str_Parame & "    AND SOLMAE_APOPRO_SOL = " & p_ApoPro & ""
   Else
      g_str_Parame = g_str_Parame & "    AND SOLMAE_COMVTA_DOL = " & p_ValVta & ""
      g_str_Parame = g_str_Parame & "    AND SOLMAE_APOPRO_DOL = " & p_ApoPro & ""
   End If
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      fs_Valida_Datos_Credito = True
   End If
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Function
Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
Dim r_arr_Mtz()      As moddat_g_tpo_DatCom

   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   pnl_NumSol.Caption = gf_Formato_NumSol(moddat_g_str_NumSol)
   pnl_FecIng.Caption = moddat_g_str_FecIng
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
    
   Call fs_Inicia
   moddat_g_int_CygTDo = 0
   moddat_g_str_CygNDo = ""
   
   'Buscar Información de la Solicitud
   Call modmip_gs_DatCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc, grd_Listad(0), 0)      'Buscar Información del Cliente
   Call modmip_gs_DatCli(moddat_g_int_CygTDo, moddat_g_str_CygNDo, grd_Listad(1), 1)      'Buscar Información del Cónyuge
   
   Call fs_DatPat        'Datos del Patrimonio
   Call fs_DatRef        'Referencias Personales
   
   Call modmip_gs_DatInm(grd_Listad(2), True) 'Buscar datos del Inmueble
   Call fs_DatInm_Aux    'Datos del auxiliares
      
   'Call fs_DatCre        'Datos del Crédito
   Call modmip_gs_DatCre(grd_Listad(5), r_arr_Mtz)
   moddat_g_int_TipMon = r_arr_Mtz(0).DatCom_TipMon
   moddat_g_str_CodMod = r_arr_Mtz(0).DatCom_CodMod
   l_int_PriViv_Inm = r_arr_Mtz(0).DatCom_PriViv
   l_dbl_TasInt_Cre = r_arr_Mtz(0).DatCom_TasInt
   l_dbl_CVtDol_Cre = r_arr_Mtz(0).DatCom_ComVta_Dol
   l_dbl_ApoDol_Cre = r_arr_Mtz(0).DatCom_ApoPro_Dol
   l_dbl_CVtSol_Cre = r_arr_Mtz(0).DatCom_ComVta_Sol
   l_dbl_ApoSol_Cre = r_arr_Mtz(0).DatCom_ApoPro_Sol
   If moddat_g_int_TipMon = 1 Then
      l_dbl_MtoPre_Cre = r_arr_Mtz(0).DatCom_MtoPre_Sol
   Else
      l_dbl_MtoPre_Cre = r_arr_Mtz(0).DatCom_MtoPre_Dol
   End If
   l_int_PerGra_Cre = r_arr_Mtz(0).DatCom_PerGra
   l_int_PlaAno_Cre = r_arr_Mtz(0).DatCom_PlaAno
   l_int_CuoExt_Cre = r_arr_Mtz(0).DatCom_CuoExt
   l_int_DiaPag_Cre = r_arr_Mtz(0).DatCom_DiaPag
   l_dbl_IntCap_Cre = r_arr_Mtz(0).DatCom_IntGra
   moddat_g_str_CodEjeSeg = r_arr_Mtz(0).DatCom_EjeSeg
   moddat_g_str_CodConHip = r_arr_Mtz(0).DatCom_ConHip
   
   Call fs_SolDoc        'Documentos Recibidos
   Call fs_SolDoc_Inm    'Documentos Recibidos del Inmueble
   Call fs_GasAdm        'Gastos Administrativos
   Call fs_EvaCre        'Evaluación Crediticia
   Call modmip_gs_EvaTas(grd_Listad(7))                                                      'Call fs_DatTas     'Tasación
   Call modmip_gs_EvaSeg(grd_Listad(8))                                                      'Seguros
   Call fs_DatSeg
   Call modmip_gs_Buscar_EvaLeg(grd_Listad(9), grd_Listad(9), txt_InfLeg, txt_ComCre)        'Legal
   Call fs_DatLeg
   Call modmip_gs_TraMVi(grd_Listad(12), txt_ObsMVi)                                         'Mivivienda
   Call fs_DatMVi
   Call modmip_gs_Buscar_TraCof(grd_Listad(12), txt_ObsMVi)                                  'COFIDE
   Call fs_DatCof

   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
Dim r_int_Contad     As Integer
   
   grd_GasAdm.ColWidth(0) = 3960
   grd_GasAdm.ColWidth(1) = 2340
   grd_GasAdm.ColWidth(2) = 1170
   grd_GasAdm.ColWidth(3) = 1955
   grd_GasAdm.ColWidth(4) = 1365
   grd_GasAdm.ColAlignment(0) = flexAlignLeftCenter
   grd_GasAdm.ColAlignment(1) = flexAlignCenterCenter
   grd_GasAdm.ColAlignment(2) = flexAlignRightCenter
   grd_GasAdm.ColAlignment(3) = flexAlignCenterCenter
   grd_GasAdm.ColAlignment(4) = flexAlignCenterCenter
   
   txt_ObsSeg.Text = ""
   txt_InfLeg.Text = ""
   txt_ComCre.Text = ""
   txt_ObsMVi.Text = ""
   
   'Inicializando Grid de Cliente y de Cónyuge
   For r_int_Contad = 0 To 9
      grd_Listad(r_int_Contad).ColWidth(0) = 3000
      grd_Listad(r_int_Contad).ColWidth(1) = 7940
      grd_Listad(r_int_Contad).ColAlignment(0) = flexAlignLeftCenter
      grd_Listad(r_int_Contad).ColAlignment(1) = flexAlignLeftCenter
      Call gs_LimpiaGrid(grd_Listad(r_int_Contad))
   Next r_int_Contad
   
   grd_Listad(10).ColWidth(0) = 10940
   grd_Listad(10).ColAlignment(0) = flexAlignLeftCenter
   Call gs_LimpiaGrid(grd_Listad(10))

   grd_Listad(11).ColWidth(0) = 10940
   grd_Listad(11).ColAlignment(0) = flexAlignLeftCenter
   Call gs_LimpiaGrid(grd_Listad(11))

   'Grid Mivivienda
   grd_Listad(12).ColWidth(0) = 3000
   grd_Listad(12).ColWidth(1) = 7940
   grd_Listad(12).ColAlignment(0) = flexAlignLeftCenter
   grd_Listad(12).ColAlignment(1) = flexAlignLeftCenter
   Call gs_LimpiaGrid(grd_Listad(12))
End Sub

Private Sub fs_DatPat()
   Dim r_int_Contad     As Integer
   
   Call gs_LimpiaGrid(grd_Listad(4))
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM CRE_SOLMAE "
   g_str_Parame = g_str_Parame & " WHERE SOLMAE_NUMERO = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   grd_Listad(4).Redraw = False
   g_rst_Princi.MoveFirst
   
   If g_rst_Princi!SOLMAE_REGIMB = 1 Then
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(4).Text = "INMUEBLES"
      
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT * "
      g_str_Parame = g_str_Parame & "  FROM CRE_SOLINB "
      g_str_Parame = g_str_Parame & " WHERE SOLINB_NUMSOL = '" & moddat_g_str_NumSol & "' ORDER BY SOLINB_NUMITE ASC"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
          Exit Sub
      End If
      
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         g_rst_Genera.MoveFirst
         
         r_int_Contad = 1
         Do While Not g_rst_Genera.EOF
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = "Tipo Inmueble (" & Format(r_int_Contad, "00") & ")"
            
            grd_Listad(4).Col = 1
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = moddat_gf_Consulta_ParDes("216", CStr(g_rst_Genera!SOLINB_TIPINM))
            
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = "Fecha de Adquisición (" & Format(r_int_Contad, "00") & ")"
            
            grd_Listad(4).Col = 1
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = gf_FormatoFecha(CStr(g_rst_Genera!SOLINB_FECADQ))
            
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = "Importe Valorizado (" & Format(r_int_Contad, "00") & ")"
            
            grd_Listad(4).Col = 1
            grd_Listad(4).CellFontName = "Lucida Console"
            grd_Listad(4).CellFontSize = 8
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = gf_FormatoNumero(g_rst_Genera!SOLINB_IMPVAL, 12, 2)
            
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = "Dirección (" & Format(r_int_Contad, "00") & ")"
            
            grd_Listad(4).Col = 1
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = Trim(g_rst_Genera!SOLINB_DIRECC & "")
            
            g_rst_Genera.MoveNext
            r_int_Contad = r_int_Contad + 1
            
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
         Loop
      End If
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   Else
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(4).Text = "INMUEBLES"
      
      grd_Listad(4).Col = 1
      grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(4).Text = "NO REGISTRA"
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
   End If
   
   If g_rst_Princi!SOLMAE_REGTAR = 1 Then
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(4).Text = "TARJETAS DE CREDITO"
      
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT * "
      g_str_Parame = g_str_Parame & "  FROM CRE_SOLTRJ "
      g_str_Parame = g_str_Parame & " WHERE SOLTRJ_NUMSOL = '" & moddat_g_str_NumSol & "' "
      g_str_Parame = g_str_Parame & " ORDER BY SOLTRJ_NUMITE ASC"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
         Exit Sub
      End If
   
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         g_rst_Genera.MoveFirst
   
         r_int_Contad = 1
         Do While Not g_rst_Genera.EOF
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(4).Text = "Institución Financiera (" & Format(r_int_Contad, "00") & ")"
   
            grd_Listad(4).Col = 1
            grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(4).Text = moddat_gf_Consulta_ParDes("505", g_rst_Genera!SOLTRJ_CODINS)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(4).Text = "Tipo de Tarjeta (" & Format(r_int_Contad, "00") & ")"
      
            grd_Listad(4).Col = 1
            grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(4).Text = moddat_gf_Consulta_ParDes("506", g_rst_Genera!SOLTRJ_TIPTRJ)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(4).Text = "Número de Tarjeta (" & Format(r_int_Contad, "00") & ")"
      
            grd_Listad(4).Col = 1
            grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(4).Text = Trim(g_rst_Genera!SOLTRJ_NUMTRJ & "")
   
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(4).Text = "Moneda (" & Format(r_int_Contad, "00") & ")"
      
            grd_Listad(4).Col = 1
            grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(4).Text = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Genera!SOLTRJ_TIPMON))
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(4).Text = "Saldo Actual (" & Format(r_int_Contad, "00") & ")"
      
            grd_Listad(4).Col = 1
            grd_Listad(4).CellFontName = "Lucida Console"
            grd_Listad(4).CellFontSize = 8
            grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(4).Text = gf_FormatoNumero(g_rst_Genera!SOLTRJ_SALACT, 12, 2)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(4).Text = "Línea Crédito (" & Format(r_int_Contad, "00") & ")"
      
            grd_Listad(4).Col = 1
            grd_Listad(4).CellFontName = "Lucida Console"
            grd_Listad(4).CellFontSize = 8
            grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(4).Text = gf_FormatoNumero(g_rst_Genera!SOLTRJ_LIMCRD, 12, 2)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(4).Text = "Pago Mínimo (" & Format(r_int_Contad, "00") & ")"
      
            grd_Listad(4).Col = 1
            grd_Listad(4).CellFontName = "Lucida Console"
            grd_Listad(4).CellFontSize = 8
            grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(4).Text = gf_FormatoNumero(g_rst_Genera!SOLTRJ_PAGMIN, 12, 2)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      
            g_rst_Genera.MoveNext
            r_int_Contad = r_int_Contad + 1
         Loop
      End If
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   Else
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(4).Text = "TARJETAS DE CREDITO"
      
      grd_Listad(4).Col = 1
      grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(4).Text = "NO REGISTRA"
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
   End If
   
   If g_rst_Princi!SOLMAE_REGDEU = 1 Then
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(4).Text = "DEUDAS"
      
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT * "
      g_str_Parame = g_str_Parame & "  FROM CRE_SOLDEU "
      g_str_Parame = g_str_Parame & " WHERE SOLDEU_NUMSOL = '" & moddat_g_str_NumSol & "' "
      g_str_Parame = g_str_Parame & " ORDER BY SOLDEU_NUMITE ASC"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
          Exit Sub
      End If
      
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         g_rst_Genera.MoveFirst
         
         r_int_Contad = 1
         Do While Not g_rst_Genera.EOF
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = "Institución Financiera (" & Format(r_int_Contad, "00") & ")"
            
            grd_Listad(4).Col = 1
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = moddat_gf_Consulta_ParDes("505", g_rst_Genera!SOLDEU_CODINS)
            
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = "Número de Operación (" & Format(r_int_Contad, "00") & ")"
            
            grd_Listad(4).Col = 1
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = Trim(g_rst_Genera!SOLDEU_NUMOPE & "")
            
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = "Moneda (" & Format(r_int_Contad, "00") & ")"
            
            grd_Listad(4).Col = 1
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Genera!SOLDEU_TIPMON))
            
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = "Monto del Préstamo (" & Format(r_int_Contad, "00") & ")"
            
            grd_Listad(4).Col = 1
            grd_Listad(4).CellFontName = "Lucida Console"
            grd_Listad(4).CellFontSize = 8
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = gf_FormatoNumero(g_rst_Genera!SOLDEU_MTOOTO, 12, 2)
            
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = "Saldo por Pagar (" & Format(r_int_Contad, "00") & ")"
            
            grd_Listad(4).Col = 1
            grd_Listad(4).CellFontName = "Lucida Console"
            grd_Listad(4).CellFontSize = 8
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = gf_FormatoNumero(g_rst_Genera!SOLDEU_SALPAG, 12, 2)
            
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = "Cuota Mensual (" & Format(r_int_Contad, "00") & ")"
      
            grd_Listad(4).Col = 1
            grd_Listad(4).CellFontName = "Lucida Console"
            grd_Listad(4).CellFontSize = 8
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = gf_FormatoNumero(g_rst_Genera!SOLDEU_CUOMEN, 12, 2)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = "Meses x Pagar (" & Format(r_int_Contad, "00") & ")"
      
            grd_Listad(4).Col = 1
            grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(4).Text = CStr(g_rst_Genera!SOLDEU_PLAMEN)
      
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      
            g_rst_Genera.MoveNext
            r_int_Contad = r_int_Contad + 1
         Loop
      End If
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   Else
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(4).Text = "DEUDAS"
      
      grd_Listad(4).Col = 1
      grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(4).Text = "NO REGISTRA"
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
   End If
   
   If g_rst_Princi!SOLMAE_REGGAS = 1 Then
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(4).Text = "GASTOS MENSUALES"
      
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT * "
      g_str_Parame = g_str_Parame & "  FROM CRE_SOLEYM "
      g_str_Parame = g_str_Parame & " WHERE SOLEYM_NUMSOL = '" & moddat_g_str_NumSol & "' "
      g_str_Parame = g_str_Parame & " ORDER BY SOLEYM_NUMITE ASC"

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         g_rst_Genera.MoveFirst
   
         r_int_Contad = 1
         Do While Not g_rst_Genera.EOF
            grd_Listad(4).Rows = grd_Listad(4).Rows + 1
            grd_Listad(4).Row = grd_Listad(4).Rows - 1
            grd_Listad(4).Col = 0
            grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(4).Text = moddat_gf_Consulta_ParDes("220", g_rst_Genera!SOLEYM_CODEYM)
      
            grd_Listad(4).Col = 1
            grd_Listad(4).CellFontName = "Lucida Console"
            grd_Listad(4).CellFontSize = 8
            grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(4).Text = gf_FormatoNumero(g_rst_Genera!SOLEYM_IMPORT, 12, 2)
      
            g_rst_Genera.MoveNext
            r_int_Contad = r_int_Contad + 1
         Loop
      End If
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   Else
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(4).Text = "GASTOS MENSUALES"
      
      grd_Listad(4).Col = 1
      grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(4).Text = "NO REGISTRA"
   End If
   
   grd_Listad(4).Redraw = True
   Call gs_UbiIniGrid(grd_Listad(4))
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_DatRef()
   Dim r_var_ColTxt

   r_var_ColTxt = modgen_g_con_ColNeg
   Call gs_LimpiaGrid(grd_Listad(3))
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM CRE_SOLREF "
   g_str_Parame = g_str_Parame & " WHERE SOLREF_NUMSOL = '" & moddat_g_str_NumSol & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_Listad(3).Redraw = False
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         grd_Listad(3).Rows = grd_Listad(3).Rows + 1
         grd_Listad(3).Row = grd_Listad(3).Rows - 1
         grd_Listad(3).Col = 0
         grd_Listad(3).CellForeColor = r_var_ColTxt
         grd_Listad(3).Text = "Tipo de Referencia"
            
         grd_Listad(3).Col = 1
         grd_Listad(3).CellForeColor = r_var_ColTxt
         grd_Listad(3).Text = moddat_gf_Consulta_ParDes("010", CStr(g_rst_Princi!SOLREF_TIPREF))
      
         grd_Listad(3).Rows = grd_Listad(3).Rows + 1
         grd_Listad(3).Row = grd_Listad(3).Rows - 1
         grd_Listad(3).Col = 0
         grd_Listad(3).CellForeColor = r_var_ColTxt
         grd_Listad(3).Text = "Tipo de Parentesco"
         
         grd_Listad(3).Col = 1
         grd_Listad(3).CellForeColor = r_var_ColTxt
         
         If g_rst_Princi!SOLREF_TIPREF = 3 Then
            grd_Listad(3).Text = moddat_gf_Consulta_ParDes("271", CStr(g_rst_Princi!SOLREF_TIPPAR))
         ElseIf g_rst_Princi!SOLREF_TIPREF = 1 Then
            grd_Listad(3).Text = moddat_gf_Consulta_ParDes("212", CStr(g_rst_Princi!SOLREF_TIPPAR))
         Else
            grd_Listad(3).Text = moddat_gf_Consulta_ParDes("213", CStr(g_rst_Princi!SOLREF_TIPPAR))
         End If
      
         grd_Listad(3).Rows = grd_Listad(3).Rows + 1
         grd_Listad(3).Row = grd_Listad(3).Rows - 1
         grd_Listad(3).Col = 0
         grd_Listad(3).CellForeColor = r_var_ColTxt
         grd_Listad(3).Text = "Apellidos y Nombres"
   
         grd_Listad(3).Col = 1
         grd_Listad(3).CellForeColor = r_var_ColTxt
         grd_Listad(3).Text = Trim(g_rst_Princi!SOLREF_APEPAT & "") & " " & Trim(g_rst_Princi!SOLREF_APEMAT & "") & " " & Trim(g_rst_Princi!SOLREF_NOMBRE & "")
      
         grd_Listad(3).Rows = grd_Listad(3).Rows + 1
         grd_Listad(3).Row = grd_Listad(3).Rows - 1
         grd_Listad(3).Col = 0
         grd_Listad(3).CellForeColor = r_var_ColTxt
         grd_Listad(3).Text = "Teléfono"

         grd_Listad(3).Col = 1
         grd_Listad(3).CellForeColor = r_var_ColTxt
         grd_Listad(3).Text = Trim(g_rst_Princi!SOLREF_TELEFO & "")
      
         grd_Listad(3).Rows = grd_Listad(3).Rows + 1
         grd_Listad(3).Row = grd_Listad(3).Rows - 1
         grd_Listad(3).Col = 0
         grd_Listad(3).CellForeColor = r_var_ColTxt
         grd_Listad(3).Text = "Celular"
   
         grd_Listad(3).Col = 1
         grd_Listad(3).CellForeColor = r_var_ColTxt
         grd_Listad(3).Text = Trim(g_rst_Princi!SOLREF_CELULA & "")
   
         g_rst_Princi.MoveNext
         
         If r_var_ColTxt = modgen_g_con_ColNeg Then
            r_var_ColTxt = modgen_g_con_ColAzu
         Else
            r_var_ColTxt = modgen_g_con_ColNeg
         End If
      Loop
      
      grd_Listad(3).Redraw = True
      Call gs_UbiIniGrid(grd_Listad(3))
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_DatInm_Aux()
   g_str_Parame = "SELECT * FROM CRE_SOLINM WHERE "
   g_str_Parame = g_str_Parame & "SOLINM_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      l_str_UbiGeo_Inm = g_rst_Princi!SOLINM_UBIGEO
      l_int_PryMCs_Inm = g_rst_Princi!SOLINM_PRYMCS
      l_str_PryCod_Inm = Trim(g_rst_Princi!SOLINM_PRYCOD & "")
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_DatCre()
   Call gs_LimpiaGrid(grd_Listad(5))
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM CRE_SOLMAE "
   g_str_Parame = g_str_Parame & " WHERE SOLMAE_NUMERO = '" & moddat_g_str_NumSol & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   grd_Listad(5).Redraw = False
   g_rst_Princi.MoveFirst
   
   moddat_g_int_TipMon = g_rst_Princi!SOLMAE_TIPMON
   moddat_g_str_CodMod = Format(CInt(CStr(g_rst_Princi!SOLMAE_CODMOD)), "00")
   
   grd_Listad(5).Rows = grd_Listad(5).Rows + 1
   grd_Listad(5).Row = grd_Listad(5).Rows - 1
   grd_Listad(5).Col = 0
   grd_Listad(5).Text = "Sub-Producto"

   grd_Listad(5).Col = 1
   grd_Listad(5).Text = moddat_gf_Consulta_SubPrd(g_rst_Princi!SOLMAE_CODPRD, g_rst_Princi!SOLMAE_CODSUB)
   
   grd_Listad(5).Rows = grd_Listad(5).Rows + 1
   grd_Listad(5).Row = grd_Listad(5).Rows - 1
   grd_Listad(5).Col = 0
   grd_Listad(5).Text = "Tipo de Evaluación"

   grd_Listad(5).Col = 1
   grd_Listad(5).Text = moddat_gf_Consulta_ParDes("038", CStr(g_rst_Princi!SOLMAE_TIPEVA))
   
   grd_Listad(5).Rows = grd_Listad(5).Rows + 1
   grd_Listad(5).Row = grd_Listad(5).Rows - 1
   grd_Listad(5).Col = 0
   grd_Listad(5).Text = "Primera Vivienda"

   grd_Listad(5).Col = 1
   grd_Listad(5).Text = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!SOLMAE_PRIVIV))
   l_int_PriViv_Inm = g_rst_Princi!SOLMAE_PRIVIV
   
   grd_Listad(5).Rows = grd_Listad(5).Rows + 1
   grd_Listad(5).Row = grd_Listad(5).Rows - 1
   grd_Listad(5).Col = 0
   grd_Listad(5).Text = "Moneda del Préstamo"

   grd_Listad(5).Col = 1
   grd_Listad(5).Text = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!SOLMAE_TIPMON))
   
   grd_Listad(5).Rows = grd_Listad(5).Rows + 1
   grd_Listad(5).Row = grd_Listad(5).Rows - 1
   grd_Listad(5).Col = 0
   grd_Listad(5).Text = "Fecha de Solicitud"

   grd_Listad(5).Col = 1
   grd_Listad(5).Text = gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL))
   
   grd_Listad(5).Rows = grd_Listad(5).Rows + 1
   grd_Listad(5).Row = grd_Listad(5).Rows - 1
   grd_Listad(5).Col = 0
   grd_Listad(5).Text = "Tasa de Interés"
   
   grd_Listad(5).Col = 1
   grd_Listad(5).Text = CStr(g_rst_Princi!SOLMAE_TASINT) & "%"
   
   If g_rst_Princi!SOLMAE_TIPMON = 2 Then
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Valor de Compra Venta"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).CellFontName = "Lucida Console"
      grd_Listad(5).CellFontSize = 8
      grd_Listad(5).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_COMVTA_DOL, 12, 2)
   
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Aporte Propio"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).CellFontName = "Lucida Console"
      grd_Listad(5).CellFontSize = 8
      grd_Listad(5).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_APOPRO_DOL, 12, 2)
   
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Monto Préstamo"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).CellFontName = "Lucida Console"
      grd_Listad(5).CellFontSize = 8
      grd_Listad(5).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_MTOPRE_DOL, 12, 2)
   Else
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Valor de Compra Venta"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).CellFontName = "Lucida Console"
      grd_Listad(5).CellFontSize = 8
      grd_Listad(5).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_COMVTA_SOL, 12, 2)
   
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Aporte Propio"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).CellFontName = "Lucida Console"
      grd_Listad(5).CellFontSize = 8
      If g_rst_Princi!SOLMAE_CODPRD = "021" Or g_rst_Princi!SOLMAE_CODPRD = "022" Or g_rst_Princi!SOLMAE_CODPRD = "023" Then
         grd_Listad(5).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_APOPRO_SOL, 12, 2) & "  (INCLUYE BBP " & moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & Format(g_rst_Princi!SOLMAE_FMVBBP, "##,###,##0.00") & ")"
      Else
         grd_Listad(5).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_APOPRO_SOL, 12, 2)
      End If
   
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Monto Préstamo"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).CellFontName = "Lucida Console"
      grd_Listad(5).CellFontSize = 8
      grd_Listad(5).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_MTOPRE_SOL, 12, 2)
   End If

   grd_Listad(5).Rows = grd_Listad(5).Rows + 1
   grd_Listad(5).Row = grd_Listad(5).Rows - 1
   grd_Listad(5).Col = 0
   grd_Listad(5).Text = "Plazo (Años)"

   grd_Listad(5).Col = 1
   grd_Listad(5).Text = CStr(g_rst_Princi!SOLMAE_PLAANO)

   grd_Listad(5).Rows = grd_Listad(5).Rows + 1
   grd_Listad(5).Row = grd_Listad(5).Rows - 1
   grd_Listad(5).Col = 0
   grd_Listad(5).Text = "Período de Gracia (Meses)"

   grd_Listad(5).Col = 1
   grd_Listad(5).Text = CStr(g_rst_Princi!SOLMAE_PERGRA)

   grd_Listad(5).Rows = grd_Listad(5).Rows + 1
   grd_Listad(5).Row = grd_Listad(5).Rows - 1
   grd_Listad(5).Col = 0
   grd_Listad(5).Text = "Interés Capitalizado"

   grd_Listad(5).Col = 1
   grd_Listad(5).CellFontName = "Lucida Console"
   grd_Listad(5).CellFontSize = 8
   grd_Listad(5).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_INTGRA, 12, 2)
   
   grd_Listad(5).Rows = grd_Listad(5).Rows + 1
   grd_Listad(5).Row = grd_Listad(5).Rows - 1
   grd_Listad(5).Col = 0
   grd_Listad(5).Text = "Cuotas Extraordinarias"

   grd_Listad(5).Col = 1
   grd_Listad(5).Text = moddat_gf_Consulta_ParDes("277", CStr(g_rst_Princi!SOLMAE_CUOEXT))
   
   grd_Listad(5).Rows = grd_Listad(5).Rows + 1
   grd_Listad(5).Row = grd_Listad(5).Rows - 1
   grd_Listad(5).Col = 0
   grd_Listad(5).Text = "Compañía de Seguros"

   grd_Listad(5).Col = 1
   grd_Listad(5).Text = moddat_gf_Consulta_ComSeg(g_rst_Princi!SOLMAE_ESGDES & "")
   
   grd_Listad(5).Rows = grd_Listad(5).Rows + 1
   grd_Listad(5).Row = grd_Listad(5).Rows - 1
   grd_Listad(5).Col = 0
   grd_Listad(5).Text = "Tipo de Seguro Desgravamen"

   grd_Listad(5).Col = 1
   grd_Listad(5).Text = moddat_gf_Consulta_TipSeg(g_rst_Princi!SOLMAE_ESGDES, g_rst_Princi!SOLMAE_TIPSEG)
   
   grd_Listad(5).Rows = grd_Listad(5).Rows + 1
   grd_Listad(5).Row = grd_Listad(5).Rows - 1
   grd_Listad(5).Col = 0
   grd_Listad(5).Text = "Día de Pago"

   grd_Listad(5).Col = 1
   grd_Listad(5).Text = Format(g_rst_Princi!SOLMAE_DIAPAG, "00")
   
   If g_rst_Princi!SOLMAE_TIPEVA = 2 Then
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Institución Financiera de Ahorro"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).Text = moddat_gf_Consulta_ParDes("505", g_rst_Princi!SOLMAE_INSFIN)
      
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Monto Mínimo de Ahorro Mensual"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).CellFontName = "Lucida Console"
      grd_Listad(5).CellFontSize = 8
      grd_Listad(5).Text = moddat_gf_Consulta_ParDes("229", g_rst_Princi!SOLMAE_MONAHO) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_MTOAHO, 12, 2)
   
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Meses Ahorrados"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).Text = CStr(g_rst_Princi!SOLMAE_MESAHO)
   End If
   
   grd_Listad(5).Rows = grd_Listad(5).Rows + 1
   grd_Listad(5).Row = grd_Listad(5).Rows - 1
   grd_Listad(5).Col = 0
   grd_Listad(5).Text = "Observaciones"

   grd_Listad(5).Col = 1
   grd_Listad(5).Text = Trim(g_rst_Princi!SOLMAE_OBSERV & "")
   
   grd_Listad(5).Rows = grd_Listad(5).Rows + 1
   grd_Listad(5).Row = grd_Listad(5).Rows - 1
   grd_Listad(5).Col = 0
   grd_Listad(5).Text = "Consejero Hipotecario"

   grd_Listad(5).Col = 1
   grd_Listad(5).Text = moddat_gf_Buscar_NomEje(g_rst_Princi!SOLMAE_CONHIP)
   
   grd_Listad(5).Rows = grd_Listad(5).Rows + 1
   grd_Listad(5).Row = grd_Listad(5).Rows - 1
   grd_Listad(5).Col = 0
   grd_Listad(5).Text = "Ejecutivo de Seguimiento"

   grd_Listad(5).Col = 1
   grd_Listad(5).Text = moddat_gf_Buscar_NomEje(g_rst_Princi!SOLMAE_EJESEG)
   
   l_dbl_TasInt_Cre = g_rst_Princi!SOLMAE_TASINT
   l_dbl_CVtDol_Cre = g_rst_Princi!SOLMAE_COMVTA_DOL
   l_dbl_ApoDol_Cre = g_rst_Princi!SOLMAE_APOPRO_DOL
   l_dbl_CVtSol_Cre = g_rst_Princi!SOLMAE_COMVTA_SOL
   l_dbl_ApoSol_Cre = g_rst_Princi!SOLMAE_APOPRO_SOL
   
   If moddat_g_int_TipMon = 1 Then
      l_dbl_MtoPre_Cre = g_rst_Princi!SOLMAE_MTOPRE_SOL
   Else
      l_dbl_MtoPre_Cre = g_rst_Princi!SOLMAE_MTOPRE_DOL
   End If
   
   l_int_PerGra_Cre = g_rst_Princi!SOLMAE_PERGRA
   l_int_PlaAno_Cre = g_rst_Princi!SOLMAE_PLAANO
   l_int_CuoExt_Cre = g_rst_Princi!SOLMAE_CUOEXT
   l_int_DiaPag_Cre = g_rst_Princi!SOLMAE_DIAPAG
   l_dbl_IntCap_Cre = g_rst_Princi!SOLMAE_INTGRA
   
   moddat_g_str_CodEjeSeg = Trim(g_rst_Princi!SOLMAE_EJESEG & "")
   moddat_g_str_CodConHip = Trim(g_rst_Princi!SOLMAE_CONHIP & "")
   
   grd_Listad(5).Redraw = True
   Call gs_UbiIniGrid(grd_Listad(5))
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_SolDoc()
   Call gs_LimpiaGrid(grd_Listad(10))
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM CRE_SOLDOC "
   g_str_Parame = g_str_Parame & " WHERE SOLDOC_NUMSOL = '" & moddat_g_str_NumSol & "' "
   g_str_Parame = g_str_Parame & "   AND (SOLDOC_TIPDOC = 1 OR SOLDOC_TIPDOC = 2)"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   grd_Listad(10).Redraw = False
   Do While Not g_rst_Princi.EOF
      grd_Listad(10).Rows = grd_Listad(10).Rows + 1
      grd_Listad(10).Row = grd_Listad(10).Rows - 1
      grd_Listad(10).Col = 0
      
      If g_rst_Princi!SOLDOC_TIPDOC = 1 Then
         'Buscar en Parámetros por Producto
         If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), g_rst_Princi!SOLDOC_CODPRD, g_rst_Princi!SOLDOC_CODSUB, g_rst_Princi!SOLDOC_CODGRP, g_rst_Princi!SOLDOC_CODITE) Then
            grd_Listad(10).Text = moddat_g_arr_Genera(1).Genera_Nombre
         End If
      Else
         'Buscar en Parámetros por Actividad Económica
         If moddat_gf_Consulta_ParAct(moddat_g_arr_Genera(), g_rst_Princi!SOLDOC_CODPRD, g_rst_Princi!SOLDOC_CODSUB, CStr(g_rst_Princi!SOLDOC_CODACT), g_rst_Princi!SOLDOC_CODGRP, g_rst_Princi!SOLDOC_CODITE) Then
            grd_Listad(10).Text = moddat_g_arr_Genera(1).Genera_Nombre
         End If
      End If
      
      g_rst_Princi.MoveNext
   Loop
   
   grd_Listad(10).Redraw = True
   Call gs_UbiIniGrid(grd_Listad(10))
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_SolDoc_Inm()
   Call gs_LimpiaGrid(grd_Listad(11))
   
   'Mostrar Todos los Documentos Recibidos
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM CRE_SOLDOC "
   g_str_Parame = g_str_Parame & " WHERE SOLDOC_NUMSOL = '" & moddat_g_str_NumSol & "' "
   g_str_Parame = g_str_Parame & "   AND SOLDOC_TIPDOC = 3 "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   grd_Listad(11).Redraw = False
   Do While Not g_rst_Princi.EOF
      grd_Listad(11).Rows = grd_Listad(11).Rows + 1
      grd_Listad(11).Row = grd_Listad(11).Rows - 1
      grd_Listad(11).Col = 0
      
      If g_rst_Princi!SOLDOC_TIPDOC = 3 Then
         'Buscar en Parámetros por Producto
         If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), g_rst_Princi!SOLDOC_CODPRD, g_rst_Princi!SOLDOC_CODSUB, g_rst_Princi!SOLDOC_CODGRP, g_rst_Princi!SOLDOC_CODITE) Then
            grd_Listad(11).Text = moddat_g_arr_Genera(1).Genera_Nombre
         End If
      Else
         'Buscar en Parámetros por Actividad Económica
         If moddat_gf_Consulta_ParAct(moddat_g_arr_Genera(), g_rst_Princi!SOLDOC_CODPRD, g_rst_Princi!SOLDOC_CODSUB, CStr(g_rst_Princi!SOLDOC_CODACT), g_rst_Princi!SOLDOC_CODGRP, g_rst_Princi!SOLDOC_CODITE) Then
            grd_Listad(11).Text = moddat_g_arr_Genera(1).Genera_Nombre
         End If
      End If
      
      g_rst_Princi.MoveNext
   Loop
   
   grd_Listad(11).Redraw = True
   Call gs_UbiIniGrid(grd_Listad(11))
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_DatTas()
   Call gs_LimpiaGrid(grd_Listad(7))
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM TRA_EVATAS "
   g_str_Parame = g_str_Parame & " WHERE EVATAS_NUMSOL = '" & moddat_g_str_NumSol & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      grd_Listad(7).Rows = grd_Listad(7).Rows + 1
      grd_Listad(7).Row = grd_Listad(7).Rows - 1
      grd_Listad(7).Col = 0
      grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(7).Text = "Empresa Peritaje"
      
      grd_Listad(7).Col = 1
      grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(7).Text = moddat_gf_Consulta_ParDes("507", g_rst_Princi!EVATAS_CODEMP)
      
      grd_Listad(7).Rows = grd_Listad(7).Rows + 1
      grd_Listad(7).Row = grd_Listad(7).Rows - 1
      grd_Listad(7).Col = 0
      grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(7).Text = "Nombre Perito"
      
      grd_Listad(7).Col = 1
      grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(7).Text = Trim(g_rst_Princi!EVATAS_NOMPER & "")
      
      grd_Listad(7).Rows = grd_Listad(7).Rows + 1
      grd_Listad(7).Row = grd_Listad(7).Rows - 1
      grd_Listad(7).Col = 0
      grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(7).Text = "Código REPEV SBS"
      
      grd_Listad(7).Col = 1
      grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(7).Text = Trim(g_rst_Princi!EVATAS_CODPER & "")
      
      grd_Listad(7).Rows = grd_Listad(7).Rows + 1
      grd_Listad(7).Row = grd_Listad(7).Rows - 1
      grd_Listad(7).Col = 0
      grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(7).Text = "Nro. de Informe"
      
      grd_Listad(7).Col = 1
      grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(7).Text = Trim(g_rst_Princi!EVATAS_NUMINF & "")
      
      grd_Listad(7).Rows = grd_Listad(7).Rows + 1
      grd_Listad(7).Row = grd_Listad(7).Rows - 1
      grd_Listad(7).Col = 0
      grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(7).Text = "Fecha Evaluación"
      
      grd_Listad(7).Col = 1
      grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(7).Text = gf_FormatoFecha(CStr(g_rst_Princi!EVATAS_FECEVA))
      
      grd_Listad(7).Rows = grd_Listad(7).Rows + 1
      grd_Listad(7).Row = grd_Listad(7).Rows - 1
      grd_Listad(7).Col = 0
      grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(7).Text = "Año de Construcción"
      
      grd_Listad(7).Col = 1
      grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(7).Text = CStr(g_rst_Princi!EVATAS_ANOCON)
      
      grd_Listad(7).Rows = grd_Listad(7).Rows + 1
      grd_Listad(7).Row = grd_Listad(7).Rows - 1
      grd_Listad(7).Col = 0
      grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(7).Text = "Nro. de Pisos"
      
      grd_Listad(7).Col = 1
      grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(7).Text = CStr(g_rst_Princi!EVATAS_NUMPIS)
      
      grd_Listad(7).Rows = grd_Listad(7).Rows + 1
      grd_Listad(7).Row = grd_Listad(7).Rows - 1
      grd_Listad(7).Col = 0
      grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(7).Text = "Nro. de Sótanos"
      
      grd_Listad(7).Col = 1
      grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(7).Text = CStr(g_rst_Princi!EVATAS_NUMSOT)
      
      grd_Listad(7).Rows = grd_Listad(7).Rows + 1
      grd_Listad(7).Row = grd_Listad(7).Rows - 1
      grd_Listad(7).Col = 0
      grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(7).Text = "Tipo de Inmueble"
      
      grd_Listad(7).Col = 1
      grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(7).Text = moddat_gf_Consulta_ParDes("221", CStr(g_rst_Princi!EVATAS_TIPINM))
      
      grd_Listad(7).Rows = grd_Listad(7).Rows + 1
      grd_Listad(7).Row = grd_Listad(7).Rows - 1
      grd_Listad(7).Col = 0
      grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(7).Text = "Uso de Inmueble"
      
      grd_Listad(7).Col = 1
      grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(7).Text = moddat_gf_Consulta_ParDes("222", CStr(g_rst_Princi!EVATAS_USOINM))
      
      grd_Listad(7).Rows = grd_Listad(7).Rows + 1
      grd_Listad(7).Row = grd_Listad(7).Rows - 1
      grd_Listad(7).Col = 0
      grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(7).Text = "Material de Construcción"
      
      grd_Listad(7).Col = 1
      grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(7).Text = moddat_gf_Consulta_ParDes("223", CStr(g_rst_Princi!EVATAS_MATCON))
      
      'Total
      grd_Listad(7).Rows = grd_Listad(7).Rows + 1
      grd_Listad(7).Row = grd_Listad(7).Rows - 1
      grd_Listad(7).Col = 0
      grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(7).Text = "Area Terreno (Total)"
      
      grd_Listad(7).Col = 1
      grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(7).CellFontName = "Lucida Console"
      grd_Listad(7).CellFontSize = 8
      grd_Listad(7).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARETER_INM + g_rst_Princi!EVATAS_ARETER_ES1 + g_rst_Princi!EVATAS_ARETER_ES2 + g_rst_Princi!EVATAS_ARETER_DEP, 12, 2) & " m2"
      
      grd_Listad(7).Rows = grd_Listad(7).Rows + 1
      grd_Listad(7).Row = grd_Listad(7).Rows - 1
      grd_Listad(7).Col = 0
      grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(7).Text = "Area Construida (Total)"
      
      grd_Listad(7).Col = 1
      grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(7).CellFontName = "Lucida Console"
      grd_Listad(7).CellFontSize = 8
      grd_Listad(7).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARECON_INM + g_rst_Princi!EVATAS_ARECON_ES1 + g_rst_Princi!EVATAS_ARECON_ES2 + g_rst_Princi!EVATAS_ARECON_DEP, 12, 2) & " m2"
      
      grd_Listad(7).Rows = grd_Listad(7).Rows + 1
      grd_Listad(7).Row = grd_Listad(7).Rows - 1
      grd_Listad(7).Col = 0
      grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(7).Text = "Suma Asegurada (Total)"
      
      grd_Listad(7).Col = 1
      grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(7).CellFontName = "Lucida Console"
      grd_Listad(7).CellFontSize = 8
      grd_Listad(7).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_SUMASE_INM + g_rst_Princi!EVATAS_SUMASE_ES1 + g_rst_Princi!EVATAS_SUMASE_ES2 + g_rst_Princi!EVATAS_SUMASE_DEP, 12, 2)
      
      grd_Listad(7).Rows = grd_Listad(7).Rows + 1
      grd_Listad(7).Row = grd_Listad(7).Rows - 1
      grd_Listad(7).Col = 0
      grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(7).Text = "Valor Comercial (Total)"
      
      grd_Listad(7).Col = 1
      grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(7).CellFontName = "Lucida Console"
      grd_Listad(7).CellFontSize = 8
      grd_Listad(7).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALCOM_INM + g_rst_Princi!EVATAS_VALCOM_ES1 + g_rst_Princi!EVATAS_VALCOM_ES2 + g_rst_Princi!EVATAS_VALCOM_DEP, 12, 2)
      
      grd_Listad(7).Rows = grd_Listad(7).Rows + 1
      grd_Listad(7).Row = grd_Listad(7).Rows - 1
      grd_Listad(7).Col = 0
      grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(7).Text = "Valor Realización (Total)"
      
      grd_Listad(7).Col = 1
      grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(7).CellFontName = "Lucida Console"
      grd_Listad(7).CellFontSize = 8
      grd_Listad(7).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALREA_INM + g_rst_Princi!EVATAS_VALREA_ES1 + g_rst_Princi!EVATAS_VALREA_ES2 + g_rst_Princi!EVATAS_VALREA_DEP, 12, 2)
      
      grd_Listad(7).Rows = grd_Listad(7).Rows + 1
      grd_Listad(7).Row = grd_Listad(7).Rows - 1
      grd_Listad(7).Col = 0
      grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(7).Text = "Valor Terreno (Total)"
      
      grd_Listad(7).Col = 1
      grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(7).CellFontName = "Lucida Console"
      grd_Listad(7).CellFontSize = 8
      grd_Listad(7).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALTER_INM + g_rst_Princi!EVATAS_VALTER_ES1 + g_rst_Princi!EVATAS_VALTER_ES2 + g_rst_Princi!EVATAS_VALTER_DEP, 12, 2)
      
      grd_Listad(7).Rows = grd_Listad(7).Rows + 1
      grd_Listad(7).Row = grd_Listad(7).Rows - 1
      grd_Listad(7).Col = 0
      grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(7).Text = "Valor Edificación (Total)"
      
      grd_Listad(7).Col = 1
      grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(7).CellFontName = "Lucida Console"
      grd_Listad(7).CellFontSize = 8
      grd_Listad(7).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALEDI_INM + g_rst_Princi!EVATAS_VALEDI_ES1 + g_rst_Princi!EVATAS_VALEDI_ES2 + g_rst_Princi!EVATAS_VALEDI_DEP, 12, 2)
   
      grd_Listad(7).Rows = grd_Listad(7).Rows + 1
      grd_Listad(7).Row = grd_Listad(7).Rows - 1
      grd_Listad(7).Col = 0
      grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(7).Text = "Valor Areas Comunes (Total)"
      
      grd_Listad(7).Col = 1
      grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(7).CellFontName = "Lucida Console"
      grd_Listad(7).CellFontSize = 8
      grd_Listad(7).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALACO_INM + g_rst_Princi!EVATAS_VALACO_ES1 + g_rst_Princi!EVATAS_VALACO_ES2 + g_rst_Princi!EVATAS_VALACO_DEP, 12, 2)
   
      'Inmueble
      grd_Listad(7).Rows = grd_Listad(7).Rows + 2
      grd_Listad(7).Row = grd_Listad(7).Rows - 1
      grd_Listad(7).Col = 0
      grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(7).Text = "Area Terreno (Inmueble)"
      
      grd_Listad(7).Col = 1
      grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(7).CellFontName = "Lucida Console"
      grd_Listad(7).CellFontSize = 8
      grd_Listad(7).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARETER_INM, 12, 2) & " m2"
      
      grd_Listad(7).Rows = grd_Listad(7).Rows + 1
      grd_Listad(7).Row = grd_Listad(7).Rows - 1
      grd_Listad(7).Col = 0
      grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(7).Text = "Area Construida (Inmueble)"
      
      grd_Listad(7).Col = 1
      grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(7).CellFontName = "Lucida Console"
      grd_Listad(7).CellFontSize = 8
      grd_Listad(7).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARECON_INM, 12, 2) & " m2"
      
      grd_Listad(7).Rows = grd_Listad(7).Rows + 1
      grd_Listad(7).Row = grd_Listad(7).Rows - 1
      grd_Listad(7).Col = 0
      grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(7).Text = "Suma Asegurada (Inmueble)"
      
      grd_Listad(7).Col = 1
      grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(7).CellFontName = "Lucida Console"
      grd_Listad(7).CellFontSize = 8
      grd_Listad(7).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_SUMASE_INM, 12, 2)
      
      grd_Listad(7).Rows = grd_Listad(7).Rows + 1
      grd_Listad(7).Row = grd_Listad(7).Rows - 1
      grd_Listad(7).Col = 0
      grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(7).Text = "Valor Comercial (Inmueble)"
      
      grd_Listad(7).Col = 1
      grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(7).CellFontName = "Lucida Console"
      grd_Listad(7).CellFontSize = 8
      grd_Listad(7).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALCOM_INM, 12, 2)
      
      grd_Listad(7).Rows = grd_Listad(7).Rows + 1
      grd_Listad(7).Row = grd_Listad(7).Rows - 1
      grd_Listad(7).Col = 0
      grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(7).Text = "Valor Realización (Inmueble)"
      
      grd_Listad(7).Col = 1
      grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(7).CellFontName = "Lucida Console"
      grd_Listad(7).CellFontSize = 8
      grd_Listad(7).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALREA_INM, 12, 2)
      
      grd_Listad(7).Rows = grd_Listad(7).Rows + 1
      grd_Listad(7).Row = grd_Listad(7).Rows - 1
      grd_Listad(7).Col = 0
      grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(7).Text = "Valor Terreno (Inmueble)"
      
      grd_Listad(7).Col = 1
      grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(7).CellFontName = "Lucida Console"
      grd_Listad(7).CellFontSize = 8
      grd_Listad(7).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALTER_INM, 12, 2)
      
      grd_Listad(7).Rows = grd_Listad(7).Rows + 1
      grd_Listad(7).Row = grd_Listad(7).Rows - 1
      grd_Listad(7).Col = 0
      grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(7).Text = "Valor Edificación (Inmueble)"
      
      grd_Listad(7).Col = 1
      grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(7).CellFontName = "Lucida Console"
      grd_Listad(7).CellFontSize = 8
      grd_Listad(7).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALEDI_INM, 12, 2)
   
      grd_Listad(7).Rows = grd_Listad(7).Rows + 1
      grd_Listad(7).Row = grd_Listad(7).Rows - 1
      grd_Listad(7).Col = 0
      grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(7).Text = "Valor Areas Comunes (Inmueble)"
      
      grd_Listad(7).Col = 1
      grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(7).CellFontName = "Lucida Console"
      grd_Listad(7).CellFontSize = 8
      grd_Listad(7).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALACO_INM, 12, 2)
   
      'Estacionamiento 1
      If g_rst_Princi!EVATAS_FLGEST_ES1 = 1 Then
         grd_Listad(7).Rows = grd_Listad(7).Rows + 2
         grd_Listad(7).Row = grd_Listad(7).Rows - 1
         grd_Listad(7).Col = 0
         grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(7).Text = "Area Terreno (Estac. 1)"
         
         grd_Listad(7).Col = 1
         grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(7).CellFontName = "Lucida Console"
         grd_Listad(7).CellFontSize = 8
         grd_Listad(7).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARETER_ES1, 12, 2) & " m2"
         
         grd_Listad(7).Rows = grd_Listad(7).Rows + 1
         grd_Listad(7).Row = grd_Listad(7).Rows - 1
         grd_Listad(7).Col = 0
         grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(7).Text = "Area Construida (Estac. 1)"
         
         grd_Listad(7).Col = 1
         grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(7).CellFontName = "Lucida Console"
         grd_Listad(7).CellFontSize = 8
         grd_Listad(7).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARECON_ES1, 12, 2) & " m2"
         
         grd_Listad(7).Rows = grd_Listad(7).Rows + 1
         grd_Listad(7).Row = grd_Listad(7).Rows - 1
         grd_Listad(7).Col = 0
         grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(7).Text = "Suma Asegurada (Estac. 1)"
         
         grd_Listad(7).Col = 1
         grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(7).CellFontName = "Lucida Console"
         grd_Listad(7).CellFontSize = 8
         grd_Listad(7).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_SUMASE_ES1, 12, 2)
         
         grd_Listad(7).Rows = grd_Listad(7).Rows + 1
         grd_Listad(7).Row = grd_Listad(7).Rows - 1
         grd_Listad(7).Col = 0
         grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(7).Text = "Valor Comercial (Estac. 1)"
         
         grd_Listad(7).Col = 1
         grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(7).CellFontName = "Lucida Console"
         grd_Listad(7).CellFontSize = 8
         grd_Listad(7).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALCOM_ES1, 12, 2)
         
         grd_Listad(7).Rows = grd_Listad(7).Rows + 1
         grd_Listad(7).Row = grd_Listad(7).Rows - 1
         grd_Listad(7).Col = 0
         grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(7).Text = "Valor Realización (Estac. 1)"
         
         grd_Listad(7).Col = 1
         grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(7).CellFontName = "Lucida Console"
         grd_Listad(7).CellFontSize = 8
         grd_Listad(7).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALREA_ES1, 12, 2)
         
         grd_Listad(7).Rows = grd_Listad(7).Rows + 1
         grd_Listad(7).Row = grd_Listad(7).Rows - 1
         grd_Listad(7).Col = 0
         grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(7).Text = "Valor Terreno (Estac. 1)"
         
         grd_Listad(7).Col = 1
         grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(7).CellFontName = "Lucida Console"
         grd_Listad(7).CellFontSize = 8
         grd_Listad(7).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALTER_ES1, 12, 2)
         
         grd_Listad(7).Rows = grd_Listad(7).Rows + 1
         grd_Listad(7).Row = grd_Listad(7).Rows - 1
         grd_Listad(7).Col = 0
         grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(7).Text = "Valor Edificación (Estac. 1)"
         
         grd_Listad(7).Col = 1
         grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(7).CellFontName = "Lucida Console"
         grd_Listad(7).CellFontSize = 8
         grd_Listad(7).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALEDI_ES1, 12, 2)
      
         grd_Listad(7).Rows = grd_Listad(7).Rows + 1
         grd_Listad(7).Row = grd_Listad(7).Rows - 1
         grd_Listad(7).Col = 0
         grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(7).Text = "Valor Areas Comunes (Estac. 1)"
         
         grd_Listad(7).Col = 1
         grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(7).CellFontName = "Lucida Console"
         grd_Listad(7).CellFontSize = 8
         grd_Listad(7).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALACO_ES1, 12, 2)
      End If
   
      If g_rst_Princi!EVATAS_FLGEST_ES2 = 1 Then
         grd_Listad(7).Rows = grd_Listad(7).Rows + 2
         grd_Listad(7).Row = grd_Listad(7).Rows - 1
         grd_Listad(7).Col = 0
         grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(7).Text = "Area Terreno (Estac. 2)"
         
         grd_Listad(7).Col = 1
         grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(7).CellFontName = "Lucida Console"
         grd_Listad(7).CellFontSize = 8
         grd_Listad(7).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARETER_ES2, 12, 2) & " m2"
         
         grd_Listad(7).Rows = grd_Listad(7).Rows + 1
         grd_Listad(7).Row = grd_Listad(7).Rows - 1
         grd_Listad(7).Col = 0
         grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(7).Text = "Area Construida (Estac. 2)"
         
         grd_Listad(7).Col = 1
         grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(7).CellFontName = "Lucida Console"
         grd_Listad(7).CellFontSize = 8
         grd_Listad(7).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARECON_ES2, 12, 2) & " m2"
         
         grd_Listad(7).Rows = grd_Listad(7).Rows + 1
         grd_Listad(7).Row = grd_Listad(7).Rows - 1
         grd_Listad(7).Col = 0
         grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(7).Text = "Suma Asegurada (Estac. 2)"
         
         grd_Listad(7).Col = 1
         grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(7).CellFontName = "Lucida Console"
         grd_Listad(7).CellFontSize = 8
         grd_Listad(7).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_SUMASE_ES2, 12, 2)
         
         grd_Listad(7).Rows = grd_Listad(7).Rows + 1
         grd_Listad(7).Row = grd_Listad(7).Rows - 1
         grd_Listad(7).Col = 0
         grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(7).Text = "Valor Comercial (Estac. 2)"
         
         grd_Listad(7).Col = 1
         grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(7).CellFontName = "Lucida Console"
         grd_Listad(7).CellFontSize = 8
         grd_Listad(7).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALCOM_ES2, 12, 2)
         
         grd_Listad(7).Rows = grd_Listad(7).Rows + 1
         grd_Listad(7).Row = grd_Listad(7).Rows - 1
         grd_Listad(7).Col = 0
         grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(7).Text = "Valor Realización (Estac. 2)"
         
         grd_Listad(7).Col = 1
         grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(7).CellFontName = "Lucida Console"
         grd_Listad(7).CellFontSize = 8
         grd_Listad(7).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALREA_ES2, 12, 2)
         
         grd_Listad(7).Rows = grd_Listad(7).Rows + 1
         grd_Listad(7).Row = grd_Listad(7).Rows - 1
         grd_Listad(7).Col = 0
         grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(7).Text = "Valor Terreno (Estac. 2)"
         
         grd_Listad(7).Col = 1
         grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(7).CellFontName = "Lucida Console"
         grd_Listad(7).CellFontSize = 8
         grd_Listad(7).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALTER_ES2, 12, 2)
         
         grd_Listad(7).Rows = grd_Listad(7).Rows + 1
         grd_Listad(7).Row = grd_Listad(7).Rows - 1
         grd_Listad(7).Col = 0
         grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(7).Text = "Valor Edificación (Estac. 2)"
         
         grd_Listad(7).Col = 1
         grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(7).CellFontName = "Lucida Console"
         grd_Listad(7).CellFontSize = 8
         grd_Listad(7).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALEDI_ES2, 12, 2)
      
         grd_Listad(7).Rows = grd_Listad(7).Rows + 1
         grd_Listad(7).Row = grd_Listad(7).Rows - 1
         grd_Listad(7).Col = 0
         grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(7).Text = "Valor Areas Comunes (Estac. 2)"
         
         grd_Listad(7).Col = 1
         grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(7).CellFontName = "Lucida Console"
         grd_Listad(7).CellFontSize = 8
         grd_Listad(7).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALACO_ES2, 12, 2)
      End If
   
      If g_rst_Princi!EVATAS_FLGEST_DEP = 1 Then
         grd_Listad(7).Rows = grd_Listad(7).Rows + 2
         grd_Listad(7).Row = grd_Listad(7).Rows - 1
         grd_Listad(7).Col = 0
         grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(7).Text = "Area Terreno (Depósito)"
         
         grd_Listad(7).Col = 1
         grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(7).CellFontName = "Lucida Console"
         grd_Listad(7).CellFontSize = 8
         grd_Listad(7).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARETER_DEP, 12, 2) & " m2"
         
         grd_Listad(7).Rows = grd_Listad(7).Rows + 1
         grd_Listad(7).Row = grd_Listad(7).Rows - 1
         grd_Listad(7).Col = 0
         grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(7).Text = "Area Construida (Depósito)"
         
         grd_Listad(7).Col = 1
         grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(7).CellFontName = "Lucida Console"
         grd_Listad(7).CellFontSize = 8
         grd_Listad(7).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARECON_DEP, 12, 2) & " m2"
         
         grd_Listad(7).Rows = grd_Listad(7).Rows + 1
         grd_Listad(7).Row = grd_Listad(7).Rows - 1
         grd_Listad(7).Col = 0
         grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(7).Text = "Suma Asegurada (Depósito)"
         
         grd_Listad(7).Col = 1
         grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(7).CellFontName = "Lucida Console"
         grd_Listad(7).CellFontSize = 8
         grd_Listad(7).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_SUMASE_DEP, 12, 2)
         
         grd_Listad(7).Rows = grd_Listad(7).Rows + 1
         grd_Listad(7).Row = grd_Listad(7).Rows - 1
         grd_Listad(7).Col = 0
         grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(7).Text = "Valor Comercial (Depósito)"
         
         grd_Listad(7).Col = 1
         grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(7).CellFontName = "Lucida Console"
         grd_Listad(7).CellFontSize = 8
         grd_Listad(7).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALCOM_DEP, 12, 2)
         
         grd_Listad(7).Rows = grd_Listad(7).Rows + 1
         grd_Listad(7).Row = grd_Listad(7).Rows - 1
         grd_Listad(7).Col = 0
         grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(7).Text = "Valor Realización (Depósito)"
         
         grd_Listad(7).Col = 1
         grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(7).CellFontName = "Lucida Console"
         grd_Listad(7).CellFontSize = 8
         grd_Listad(7).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALREA_DEP, 12, 2)
         
         grd_Listad(7).Rows = grd_Listad(7).Rows + 1
         grd_Listad(7).Row = grd_Listad(7).Rows - 1
         grd_Listad(7).Col = 0
         grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(7).Text = "Valor Terreno (Depósito)"
         
         grd_Listad(7).Col = 1
         grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(7).CellFontName = "Lucida Console"
         grd_Listad(7).CellFontSize = 8
         grd_Listad(7).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALTER_DEP, 12, 2)
         
         grd_Listad(7).Rows = grd_Listad(7).Rows + 1
         grd_Listad(7).Row = grd_Listad(7).Rows - 1
         grd_Listad(7).Col = 0
         grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(7).Text = "Valor Edificación (Depósito)"
         
         grd_Listad(7).Col = 1
         grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(7).CellFontName = "Lucida Console"
         grd_Listad(7).CellFontSize = 8
         grd_Listad(7).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALEDI_DEP, 12, 2)
      
         grd_Listad(7).Rows = grd_Listad(7).Rows + 1
         grd_Listad(7).Row = grd_Listad(7).Rows - 1
         grd_Listad(7).Col = 0
         grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(7).Text = "Valor Areas Comunes (Depósito)"
         
         grd_Listad(7).Col = 1
         grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(7).CellFontName = "Lucida Console"
         grd_Listad(7).CellFontSize = 8
         grd_Listad(7).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALACO_DEP, 12, 2)
      End If
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_DatSeg()
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT EVASEG_ESGDES, EVASEG_TIPSEG, EVASEG_TIPDES, EVASEG_FOIDES, EVASEG_TIPVIV, EVASEG_FOIVIV  "
   g_str_Parame = g_str_Parame & "  FROM TRA_EVASEG "
   g_str_Parame = g_str_Parame & " WHERE EVASEG_NUMSOL = '" & moddat_g_str_NumSol & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
     
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
              
      l_str_ESgDes_Seg = Trim(g_rst_Princi!EVASEG_ESGDES & "")
      l_int_TipSeg_Seg = g_rst_Princi!EVASEG_TIPSEG
      l_int_AplDes_Seg = g_rst_Princi!EVASEG_TIPDES
      l_dbl_FoIDes_Seg = g_rst_Princi!EVASEG_FOIDES
      
      l_int_AplViv_Seg = g_rst_Princi!EVASEG_TIPVIV
      l_dbl_FoIViv_Seg = g_rst_Princi!EVASEG_FOIVIV
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub
Private Sub fs_DatSeg_old()
   Call gs_LimpiaGrid(grd_Listad(8))
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM TRA_EVASEG "
   g_str_Parame = g_str_Parame & " WHERE EVASEG_NUMSOL = '" & moddat_g_str_NumSol & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM TRA_POLIZA "
   g_str_Parame = g_str_Parame & " WHERE POLIZA_NUMSOL = '" & moddat_g_str_NumSol & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      g_rst_Genera.MoveFirst
      
      grd_Listad(8).Rows = grd_Listad(8).Rows + 1
      grd_Listad(8).Row = grd_Listad(8).Rows - 1
      grd_Listad(8).Col = 0
      grd_Listad(8).Text = "Empresa de Seguros"
      
      grd_Listad(8).Col = 1
      grd_Listad(8).Text = moddat_gf_Consulta_ComSeg(g_rst_Princi!EVASEG_ESGDES & "")
   
      grd_Listad(8).Rows = grd_Listad(8).Rows + 2
      grd_Listad(8).Row = grd_Listad(8).Rows - 1
      grd_Listad(8).Col = 0
      grd_Listad(8).Text = "Tipo de Seguro Desgravamen"

      grd_Listad(8).Col = 1
      grd_Listad(8).Text = moddat_gf_Consulta_TipSeg(g_rst_Princi!EVASEG_ESGDES, g_rst_Princi!EVASEG_TIPSEG)
      
      grd_Listad(8).Rows = grd_Listad(8).Rows + 1
      grd_Listad(8).Row = grd_Listad(8).Rows - 1
      grd_Listad(8).Col = 0
      grd_Listad(8).Text = "Fecha Evaluación (Seg. Desgravamen)"
      
      grd_Listad(8).Col = 1
      grd_Listad(8).Text = gf_FormatoFecha(CStr(g_rst_Princi!EVASEG_EVADES))
      
      grd_Listad(8).Rows = grd_Listad(8).Rows + 1
      grd_Listad(8).Row = grd_Listad(8).Rows - 1
      grd_Listad(8).Col = 0
      grd_Listad(8).Text = "Tipo de Valor (Seg. Desgravamen)"
      
      grd_Listad(8).Col = 1
      grd_Listad(8).Text = moddat_gf_Consulta_ParDes("227", CStr(g_rst_Princi!EVASEG_TIPDES))
      
      grd_Listad(8).Rows = grd_Listad(8).Rows + 1
      grd_Listad(8).Row = grd_Listad(8).Rows - 1
      grd_Listad(8).Col = 0
      grd_Listad(8).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(8).Text = "Valor a Aplicar"
      
      grd_Listad(8).Col = 1
      grd_Listad(8).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(8).Text = Format(g_rst_Princi!EVASEG_FOIDES, "###,###,##0.000000")
      
      
      grd_Listad(8).Rows = grd_Listad(8).Rows + 1
      grd_Listad(8).Row = grd_Listad(8).Rows - 1
      grd_Listad(8).Col = 0
      grd_Listad(8).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(8).Text = "Fecha Emisión Póliza"
      
      grd_Listad(8).Col = 1
      grd_Listad(8).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(8).Text = gf_FormatoFecha(CStr(g_rst_Genera!POLIZA_FEMDES))
      
      grd_Listad(8).Rows = grd_Listad(8).Rows + 1
      grd_Listad(8).Row = grd_Listad(8).Rows - 1
      grd_Listad(8).Col = 0
      grd_Listad(8).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(8).Text = "Número de Póliza"
      
      grd_Listad(8).Col = 1
      grd_Listad(8).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(8).Text = Trim(g_rst_Genera!POLIZA_NUMDES & "") & IIf(Len(Trim(g_rst_Genera!POLIZA_NUMCYG & "")) > 0, " / " & Trim(g_rst_Genera!POLIZA_NUMCYG & ""), "")
      
      grd_Listad(8).Rows = grd_Listad(8).Rows + 2
      grd_Listad(8).Row = grd_Listad(8).Rows - 1
      grd_Listad(8).Col = 0
      grd_Listad(8).Text = "Fecha Evaluación (Seg. Inmueble)"
      
      grd_Listad(8).Col = 1
      grd_Listad(8).Text = gf_FormatoFecha(CStr(g_rst_Princi!EVASEG_EVAVIV))
      
      grd_Listad(8).Rows = grd_Listad(8).Rows + 1
      grd_Listad(8).Row = grd_Listad(8).Rows - 1
      grd_Listad(8).Col = 0
      grd_Listad(8).Text = "Tipo de Valor (Seg. Inmueble)"
      
      grd_Listad(8).Col = 1
      grd_Listad(8).Text = moddat_gf_Consulta_ParDes("227", CStr(g_rst_Princi!EVASEG_TIPVIV))
      
      grd_Listad(8).Rows = grd_Listad(8).Rows + 1
      grd_Listad(8).Row = grd_Listad(8).Rows - 1
      grd_Listad(8).Col = 0
      grd_Listad(8).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(8).Text = "Valor a Aplicar"
      
      grd_Listad(8).Col = 1
      grd_Listad(8).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(8).Text = Format(g_rst_Princi!EVASEG_FOIVIV, "###,###,##0.000000")
      
      grd_Listad(8).Rows = grd_Listad(8).Rows + 1
      grd_Listad(8).Row = grd_Listad(8).Rows - 1
      grd_Listad(8).Col = 0
      grd_Listad(8).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(8).Text = "Fecha Emisión Póliza"
      
      grd_Listad(8).Col = 1
      grd_Listad(8).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(8).Text = gf_FormatoFecha(CStr(g_rst_Genera!POLIZA_FEMVIV))
      
      grd_Listad(8).Rows = grd_Listad(8).Rows + 1
      grd_Listad(8).Row = grd_Listad(8).Rows - 1
      grd_Listad(8).Col = 0
      grd_Listad(8).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(8).Text = "Número de Póliza"
      
      grd_Listad(8).Col = 1
      grd_Listad(8).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(8).Text = Trim(g_rst_Genera!POLIZA_NUMVIV & "")
      
      txt_ObsSeg.Text = Trim(g_rst_Princi!EVASEG_OBSERV & "")
      
      l_str_ESgDes_Seg = Trim(g_rst_Princi!EVASEG_ESGDES & "")
      l_int_TipSeg_Seg = g_rst_Princi!EVASEG_TIPSEG
      l_int_AplDes_Seg = g_rst_Princi!EVASEG_TIPDES
      l_dbl_FoIDes_Seg = g_rst_Princi!EVASEG_FOIDES
      
      l_int_AplViv_Seg = g_rst_Princi!EVASEG_TIPVIV
      l_dbl_FoIViv_Seg = g_rst_Princi!EVASEG_FOIVIV
      
      Call gs_UbiIniGrid(grd_Listad(8))
   End If
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub
Private Sub grd_Listad_SelChange(Index As Integer)
   If grd_Listad(Index).Rows > 2 Then
      grd_Listad(Index).RowSel = grd_Listad(Index).Row
   End If
End Sub

Private Sub txt_ComCre_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txt_InfLeg_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txt_ObsMVi_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txt_ObsSeg_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub fs_DatLeg()
Dim r_rst_DatCre  As ADODB.Recordset
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT EVALEG_FECCOM, EVALEG_FECCVT, EVALEG_TCASBS"
   g_str_Parame = g_str_Parame & "  FROM TRA_EVALEG "
   g_str_Parame = g_str_Parame & " WHERE EVALEG_NUMSOL = '" & moddat_g_str_NumSol & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      If g_rst_Princi!EVALEG_FECCOM > 0 Then
         l_str_FecCom_Leg = gf_FormatoFecha(CStr(g_rst_Princi!EVALEG_FECCOM))
      Else
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & "SELECT EVACRE_FECCOM "
         g_str_Parame = g_str_Parame & "  FROM TRA_EVACRE "
         g_str_Parame = g_str_Parame & " WHERE EVACRE_NUMSOL = '" & moddat_g_str_NumSol & "' "
         
         If Not gf_EjecutaSQL(g_str_Parame, r_rst_DatCre, 3) Then
             Exit Sub
         End If
         
         If Not (r_rst_DatCre.BOF And r_rst_DatCre.EOF) Then
            If r_rst_DatCre!EVACRE_FECCOM > 0 Then
               l_str_FecCom_Leg = gf_FormatoFecha(CStr(r_rst_DatCre!EVACRE_FECCOM))
            End If
         End If
         
         r_rst_DatCre.Close
         Set r_rst_DatCre = Nothing
      End If
      
      If g_rst_Princi!EVALEG_FECCVT > 0 Then
        
         l_dbl_TCaSBS_Leg = 0
         
         If Not IsNull(g_rst_Princi!EVALEG_TCASBS) Then
            If g_rst_Princi!EVALEG_TCASBS > 0 Then
               l_dbl_TCaSBS_Leg = g_rst_Princi!EVALEG_TCASBS
            End If
         End If
      End If
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub
Private Sub fs_DatLeg_old()
Dim r_rst_DatCre  As ADODB.Recordset

   Call gs_LimpiaGrid(grd_Listad(9))
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM TRA_EVALEG "
   g_str_Parame = g_str_Parame & " WHERE EVALEG_NUMSOL = '" & moddat_g_str_NumSol & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      txt_InfLeg.Text = Trim(g_rst_Princi!EVALEG_INFLG1 & "") & Trim(g_rst_Princi!EVALEG_INFLG2 & "") & Trim(g_rst_Princi!EVALEG_INFLG3 & "") & Trim(g_rst_Princi!EVALEG_INFLG4 & "")
      
      If g_rst_Princi!EVALEG_FECCOM > 0 Then
         l_str_FecCom_Leg = gf_FormatoFecha(CStr(g_rst_Princi!EVALEG_FECCOM))
         txt_ComCre.Text = "Fecha de Comité de Créditos: " & gf_FormatoFecha(CStr(g_rst_Princi!EVALEG_FECCOM)) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Trim(g_rst_Princi!EVALEG_OBSCOM & "")
      Else
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & "SELECT * "
         g_str_Parame = g_str_Parame & "  FROM TRA_EVACRE "
         g_str_Parame = g_str_Parame & " WHERE EVACRE_NUMSOL = '" & moddat_g_str_NumSol & "' "
         
         If Not gf_EjecutaSQL(g_str_Parame, r_rst_DatCre, 3) Then
             Exit Sub
         End If
         
         If Not (r_rst_DatCre.BOF And r_rst_DatCre.EOF) Then
            If r_rst_DatCre!EVACRE_FECCOM > 0 Then
               l_str_FecCom_Leg = gf_FormatoFecha(CStr(r_rst_DatCre!EVACRE_FECCOM))
               txt_ComCre.Text = "Fecha del Comité de Créditos: " & gf_FormatoFecha(CStr(r_rst_DatCre!EVACRE_FECCOM)) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "Numero de Acta del Comité de Créditos: " & Trim(r_rst_DatCre!EVACRE_NROACT & "")
            End If
         End If
         
         r_rst_DatCre.Close
         Set r_rst_DatCre = Nothing
      End If
      
      If g_rst_Princi!EVALEG_FECCVT > 0 Then
         grd_Listad(9).Rows = grd_Listad(9).Rows + 1
         grd_Listad(9).Row = grd_Listad(9).Rows - 1
         grd_Listad(9).Col = 0
         grd_Listad(9).Text = "Fecha Firma Contrato Compra Venta"
         
         grd_Listad(9).Col = 1
         grd_Listad(9).Text = gf_FormatoFecha(CStr(g_rst_Princi!EVALEG_FECCVT))
         
         l_dbl_TCaSBS_Leg = 0
         
         If Not IsNull(g_rst_Princi!EVALEG_TCASBS) Then
            If g_rst_Princi!EVALEG_TCASBS > 0 Then
               grd_Listad(9).Rows = grd_Listad(9).Rows + 1
               grd_Listad(9).Row = grd_Listad(9).Rows - 1
               grd_Listad(9).Col = 0
               grd_Listad(9).Text = "Tipo de Cambio SBS"
               
               grd_Listad(9).Col = 1
               grd_Listad(9).CellFontName = "Lucida Console"
               grd_Listad(9).CellFontSize = 8
               grd_Listad(9).Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVALEG_TCASBS, 14, 4)
               
               l_dbl_TCaSBS_Leg = g_rst_Princi!EVALEG_TCASBS
            End If
         End If
      
         If g_rst_Princi!EVALEG_TCACVT > 0 Then
            grd_Listad(9).Rows = grd_Listad(9).Rows + 1
            grd_Listad(9).Row = grd_Listad(9).Rows - 1
            grd_Listad(9).Col = 0
            grd_Listad(9).Text = "Tipo de Cambio aplicado"
            
            grd_Listad(9).Col = 1
            grd_Listad(9).CellFontName = "Lucida Console"
            grd_Listad(9).CellFontSize = 8
            grd_Listad(9).Text = "S/. " & gf_FormatoNumero(g_rst_Princi!EVALEG_TCACVT, 14, 4)
         End If
      End If
      
      If g_rst_Princi!EVALEG_FIRCON > 0 Then
         grd_Listad(9).Rows = grd_Listad(9).Rows + 1
         grd_Listad(9).Row = grd_Listad(9).Rows - 1
         grd_Listad(9).Col = 0
         grd_Listad(9).Text = "Fecha Firma Contrato"
         
         grd_Listad(9).Col = 1
         grd_Listad(9).Text = gf_FormatoFecha(CStr(g_rst_Princi!EVALEG_FIRCON))
      
         grd_Listad(9).Rows = grd_Listad(9).Rows + 1
         grd_Listad(9).Row = grd_Listad(9).Rows - 1
         grd_Listad(9).Col = 0
         grd_Listad(9).Text = "Notaria"
         
         grd_Listad(9).Col = 1
         grd_Listad(9).Text = moddat_gf_Consulta_ParDes("509", g_rst_Princi!EVALEG_CODNOT)
      
         grd_Listad(9).Rows = grd_Listad(9).Rows + 1
         grd_Listad(9).Row = grd_Listad(9).Rows - 1
         grd_Listad(9).Col = 0
         grd_Listad(9).Text = "Representante Legal 1"
         
         grd_Listad(9).Col = 1
         grd_Listad(9).Text = moddat_gf_Consulta_ParDes("512", g_rst_Princi!EVALEG_REPLG1)
      
         grd_Listad(9).Rows = grd_Listad(9).Rows + 1
         grd_Listad(9).Row = grd_Listad(9).Rows - 1
         grd_Listad(9).Col = 0
         grd_Listad(9).Text = "Representante Legal 2"
         
         grd_Listad(9).Col = 1
         grd_Listad(9).Text = moddat_gf_Consulta_ParDes("512", g_rst_Princi!EVALEG_REPLG2)
         
         grd_Listad(9).Rows = grd_Listad(9).Rows + 1
         grd_Listad(9).Row = grd_Listad(9).Rows - 1
         grd_Listad(9).Col = 0
         grd_Listad(9).Text = "Monto Hipoteca"
         
         grd_Listad(9).Col = 1
         grd_Listad(9).CellFontName = "Lucida Console"
         grd_Listad(9).CellFontSize = 8
         grd_Listad(9).Text = moddat_gf_Consulta_ParDes("229", g_rst_Princi!EVALEG_MONHIP) & " " & gf_FormatoNumero(g_rst_Princi!EVALEG_MTOHIP, 12, 2)
         
         Call gs_UbiIniGrid(grd_Listad(9))
      End If
      
      If g_rst_Princi!EVALEG_FECBLQ_INM > 0 Then
         grd_Listad(9).Rows = grd_Listad(9).Rows + 2
         grd_Listad(9).Row = grd_Listad(9).Rows - 1
         grd_Listad(9).Col = 0
         grd_Listad(9).Text = "Sede Registral"
         
         grd_Listad(9).Col = 1
         grd_Listad(9).Text = moddat_gf_Consulta_ParDes("511", CStr(g_rst_Princi!EVALEG_SEDREG & ""))
         
         grd_Listad(9).Rows = grd_Listad(9).Rows + 1
         grd_Listad(9).Row = grd_Listad(9).Rows - 1
         grd_Listad(9).Col = 0
         grd_Listad(9).Text = "Fecha Bloqueo (Inmueble)"
         
         grd_Listad(9).Col = 1
         grd_Listad(9).Text = gf_FormatoFecha(CStr(g_rst_Princi!EVALEG_FECBLQ_INM))
                  
         grd_Listad(9).Rows = grd_Listad(9).Rows + 1
         grd_Listad(9).Row = grd_Listad(9).Rows - 1
         grd_Listad(9).Col = 0
         grd_Listad(9).Text = "Doc. Registral (Inmueble)"
         
         grd_Listad(9).Col = 1
         grd_Listad(9).Text = moddat_gf_Consulta_ParDes("026", g_rst_Princi!EVALEG_TIPDOC_INM)
                  
         Select Case g_rst_Princi!EVALEG_TIPDOC_INM
            Case 1: grd_Listad(9).Text = grd_Listad(9).Text & " NRO. " & Trim(g_rst_Princi!EVALEG_NUMPAR_INM & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!EVALEG_NUMAPA_INM & "")
            Case 2: grd_Listad(9).Text = grd_Listad(9).Text & " NRO. " & Trim(g_rst_Princi!EVALEG_NUMFIC_INM & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!EVALEG_NUMAFI_INM & "")
            Case 3: grd_Listad(9).Text = grd_Listad(9).Text & " (" & Trim(g_rst_Princi!EVALEG_NUMTOM_INM & "") & " / " & Trim(g_rst_Princi!EVALEG_NUMFOJ_INM & "") & " / " & Trim(g_rst_Princi!EVALEG_NUMLIB_INM & "") & ")"
         End Select
         
         If g_rst_Princi!EVALEG_FLGEST_ES1 = 1 Then
            grd_Listad(9).Rows = grd_Listad(9).Rows + 2
            grd_Listad(9).Row = grd_Listad(9).Rows - 1
            grd_Listad(9).Col = 0
            grd_Listad(9).Text = "Fecha Bloqueo (Estac. 1)"
            
            grd_Listad(9).Col = 1
            grd_Listad(9).Text = gf_FormatoFecha(CStr(g_rst_Princi!EVALEG_FECBLQ_ES1))
                       
            grd_Listad(9).Rows = grd_Listad(9).Rows + 1
            grd_Listad(9).Row = grd_Listad(9).Rows - 1
            grd_Listad(9).Col = 0
            grd_Listad(9).Text = "Doc. Registral (Estac. 1)"
            
            grd_Listad(9).Col = 1
            grd_Listad(9).Text = moddat_gf_Consulta_ParDes("026", g_rst_Princi!EVALEG_TIPDOC_ES1)
            
            Select Case g_rst_Princi!EVALEG_TIPDOC_ES1
               Case 1: grd_Listad(9).Text = grd_Listad(9).Text & " NRO. " & Trim(g_rst_Princi!EVALEG_NUMPAR_ES1 & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!EVALEG_NUMAPA_ES1 & "")
               Case 2: grd_Listad(9).Text = grd_Listad(9).Text & " NRO. " & Trim(g_rst_Princi!EVALEG_NUMFIC_ES1 & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!EVALEG_NUMAFI_ES1 & "")
               Case 3: grd_Listad(9).Text = grd_Listad(9).Text & " (" & Trim(g_rst_Princi!EVALEG_NUMTOM_ES1 & "") & " / " & Trim(g_rst_Princi!EVALEG_NUMFOJ_ES1 & "") & " / " & Trim(g_rst_Princi!EVALEG_NUMLIB_ES1 & "") & ")"
            End Select
         End If
         
         If g_rst_Princi!EVALEG_FLGEST_ES2 = 1 Then
            grd_Listad(9).Rows = grd_Listad(9).Rows + 2
            grd_Listad(9).Row = grd_Listad(9).Rows - 1
            grd_Listad(9).Col = 0
            grd_Listad(9).Text = "Fecha Bloqueo (Estac. 2)"
            
            grd_Listad(9).Col = 1
            grd_Listad(9).Text = gf_FormatoFecha(CStr(g_rst_Princi!EVALEG_FECBLQ_ES2))
                        
            grd_Listad(9).Rows = grd_Listad(9).Rows + 1
            grd_Listad(9).Row = grd_Listad(9).Rows - 1
            grd_Listad(9).Col = 0
            grd_Listad(9).Text = "Doc. Registral (Estac. 2)"
            
            grd_Listad(9).Col = 1
            grd_Listad(9).Text = moddat_gf_Consulta_ParDes("026", g_rst_Princi!EVALEG_TIPDOC_ES2)
            
            Select Case g_rst_Princi!EVALEG_TIPDOC_ES2
               Case 1: grd_Listad(9).Text = grd_Listad(9).Text & " NRO. " & Trim(g_rst_Princi!EVALEG_NUMPAR_ES2 & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!EVALEG_NUMAPA_ES2 & "")
               Case 2: grd_Listad(9).Text = grd_Listad(9).Text & " NRO. " & Trim(g_rst_Princi!EVALEG_NUMFIC_ES2 & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!EVALEG_NUMAFI_ES2 & "")
               Case 3: grd_Listad(9).Text = grd_Listad(9).Text & " (" & Trim(g_rst_Princi!EVALEG_NUMTOM_ES2 & "") & " / " & Trim(g_rst_Princi!EVALEG_NUMFOJ_ES2 & "") & " / " & Trim(g_rst_Princi!EVALEG_NUMLIB_ES2 & "") & ")"
            End Select
         End If
         
         If g_rst_Princi!EVALEG_FLGEST_DEP = 1 Then
            grd_Listad(9).Rows = grd_Listad(9).Rows + 2
            grd_Listad(9).Row = grd_Listad(9).Rows - 1
            grd_Listad(9).Col = 0
            grd_Listad(9).Text = "Fecha Bloqueo (Depósito)"
            
            grd_Listad(9).Col = 1
            grd_Listad(9).Text = gf_FormatoFecha(CStr(g_rst_Princi!EVALEG_FECBLQ_DEP))
                        
            grd_Listad(9).Rows = grd_Listad(9).Rows + 1
            grd_Listad(9).Row = grd_Listad(9).Rows - 1
            grd_Listad(9).Col = 0
            grd_Listad(9).Text = "Doc. Registral (Depósito)"
            
            grd_Listad(9).Col = 1
            grd_Listad(9).Text = moddat_gf_Consulta_ParDes("026", g_rst_Princi!EVALEG_TIPDOC_DEP)
                        
            Select Case g_rst_Princi!EVALEG_TIPDOC_DEP
               Case 1: grd_Listad(9).Text = grd_Listad(9).Text & " NRO. " & Trim(g_rst_Princi!EVALEG_NUMPAR_DEP & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!EVALEG_NUMAPA_DEP & "")
               Case 2: grd_Listad(9).Text = grd_Listad(9).Text & " NRO. " & Trim(g_rst_Princi!EVALEG_NUMFIC_DEP & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!EVALEG_NUMAFI_DEP & "")
               Case 3: grd_Listad(9).Text = grd_Listad(9).Text & " (" & Trim(g_rst_Princi!EVALEG_NUMTOM_DEP & "") & " / " & Trim(g_rst_Princi!EVALEG_NUMFOJ_DEP & "") & " / " & Trim(g_rst_Princi!EVALEG_NUMLIB_DEP & "") & ")"
            End Select
         End If
         
         Call gs_UbiIniGrid(grd_Listad(9))
      End If
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub
Private Sub fs_GasAdm()
   Dim r_dbl_Import  As Double
   
   r_dbl_Import = 0
   
   Call gs_LimpiaGrid(grd_GasAdm)
   pnl_TotGas.Caption = "0.00 "
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM TRA_GASADM "
   g_str_Parame = g_str_Parame & " WHERE GASADM_NUMSOL = '" & moddat_g_str_NumSol & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   grd_GasAdm.Redraw = False
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      grd_GasAdm.Rows = grd_GasAdm.Rows + 1
      grd_GasAdm.Row = grd_GasAdm.Rows - 1
      
      'Buscando Descripción de Gastos Administrativos
      grd_GasAdm.Col = 0
      If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), moddat_g_str_CodPrd, moddat_g_str_CodSub, "007", Format(g_rst_Princi!GASADM_CODGAS, "00") & Format(g_rst_Princi!GASADM_TIPMON, "0")) Then
         grd_GasAdm.Text = Trim(moddat_g_arr_Genera(1).Genera_Nombre)
      End If
      
      'Tipo de Moneda
      grd_GasAdm.Col = 1
      grd_GasAdm.Text = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!GASADM_TIPMON))
      
      'Importe
      grd_GasAdm.Col = 2
      grd_GasAdm.CellFontName = "Lucida Console"
      grd_GasAdm.CellFontSize = 8
      grd_GasAdm.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!GASADM_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!GASADM_IMPORT, 12, 2)
      
      r_dbl_Import = r_dbl_Import + g_rst_Princi!GASADM_IMPORT
      
      'Situación
      grd_GasAdm.Col = 3
      grd_GasAdm.Text = moddat_gf_Consulta_ParDes("001", CStr(g_rst_Princi!GASADM_SITUAC))
      
      'Fecha de Pago
      grd_GasAdm.Col = 4
      If g_rst_Princi!GASADM_PAGFEC > 0 Then
         grd_GasAdm.Text = gf_FormatoFecha(CStr(g_rst_Princi!GASADM_PAGFEC))
      End If
      
      g_rst_Princi.MoveNext
   Loop
   
   pnl_TotGas.Caption = Format(r_dbl_Import, "###,###,##0.00") & " "
   grd_GasAdm.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call gs_UbiIniGrid(grd_GasAdm)
End Sub

Private Sub fs_EvaCre()
   Call gs_LimpiaGrid(grd_Listad(6))
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM CRE_SOLMAE "
   g_str_Parame = g_str_Parame & " WHERE SOLMAE_NUMERO = '" & moddat_g_str_NumSol & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      grd_Listad(6).Rows = grd_Listad(6).Rows + 1
      grd_Listad(6).Row = grd_Listad(6).Rows - 1
      grd_Listad(6).Col = 0
      grd_Listad(6).Text = "Cuota Aceptada por Cliente"
      
      grd_Listad(6).Col = 1
      grd_Listad(6).CellFontName = "Lucida Console"
      grd_Listad(6).CellFontSize = 8
      grd_Listad(6).Text = "S/. " & gf_FormatoNumero(g_rst_Princi!SOLMAE_CUOAPR_SOL, 12, 2)
      
      If moddat_g_int_TipMon <> 1 Then
         grd_Listad(6).Rows = grd_Listad(6).Rows + 1
         grd_Listad(6).Row = grd_Listad(6).Rows - 1
         grd_Listad(6).Col = 0
         grd_Listad(6).Text = "Cuota Aceptada por Cliente"
         
         grd_Listad(6).Col = 1
         grd_Listad(6).CellFontName = "Lucida Console"
         grd_Listad(6).CellFontSize = 8
         grd_Listad(6).Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_CUOAPR_MPR, 12, 2)
         
         grd_Listad(6).Rows = grd_Listad(6).Rows + 1
      End If
      
      grd_Listad(6).Rows = grd_Listad(6).Rows + 1
      grd_Listad(6).Row = grd_Listad(6).Rows - 1
      grd_Listad(6).Col = 0
      grd_Listad(6).Text = "Cuota Aprobada"
      
      grd_Listad(6).Col = 1
      grd_Listad(6).CellFontName = "Lucida Console"
      grd_Listad(6).CellFontSize = 8
      grd_Listad(6).Text = "S/. " & gf_FormatoNumero(g_rst_Princi!SOLMAE_CUOMEN_SOL, 12, 2)
      
      If moddat_g_int_TipMon <> 1 Then
         grd_Listad(6).Rows = grd_Listad(6).Rows + 1
         grd_Listad(6).Row = grd_Listad(6).Rows - 1
         grd_Listad(6).Col = 0
         grd_Listad(6).Text = "Cuota Aprobada"
         
         grd_Listad(6).Col = 1
         grd_Listad(6).CellFontName = "Lucida Console"
         grd_Listad(6).CellFontSize = 8
         grd_Listad(6).Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_CUOMEN_MPR, 12, 2)
         
         grd_Listad(6).Rows = grd_Listad(6).Rows + 1
      End If
   
      grd_Listad(6).Rows = grd_Listad(6).Rows + 1
      grd_Listad(6).Row = grd_Listad(6).Rows - 1
      grd_Listad(6).Col = 0
      grd_Listad(6).Text = "Total Ingreso Líquido Neto"
      
      grd_Listad(6).Col = 1
      grd_Listad(6).CellFontName = "Lucida Console"
      grd_Listad(6).CellFontSize = 8
      grd_Listad(6).Text = "S/. " & gf_FormatoNumero(g_rst_Princi!SOLMAE_INGNET, 12, 2)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   Call gs_UbiIniGrid(grd_Listad(6))
End Sub

Private Function ff_Genera_NumOpe() As String
   Dim r_lng_NumOpe     As Long
   Dim r_str_NumOpe     As String
   
   ff_Genera_NumOpe = ""
   
   'Obteniendo Número de Solicitud
   Call moddat_gs_FecSis
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM CRE_FOLIOS "
   g_str_Parame = g_str_Parame & " WHERE FOLIOS_TIPFOL = 2 "
   g_str_Parame = g_str_Parame & "   AND FOLIOS_CODPRD = '" & moddat_g_str_CodPrd & "' "
   g_str_Parame = g_str_Parame & "   AND FOLIOS_CODSUC = '000' "
   g_str_Parame = g_str_Parame & "   AND FOLIOS_PERANO = " & Right(Format(Year(CDate(moddat_g_str_FecSis)), "0000"), 2)

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Function
   End If

   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      r_lng_NumOpe = 1
   Else
      r_lng_NumOpe = g_rst_Genera!FOLIOS_NUMERO + 1
   End If

   r_str_NumOpe = moddat_g_str_CodPrd & Right(Format(Year(CDate(moddat_g_str_FecSis)), "0000"), 2) & Format(r_lng_NumOpe, "00000")
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      'Actualizando Correlativo
      g_str_Parame = "USP_CRE_FOLIOS ("
      g_str_Parame = g_str_Parame & "2, "
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_CodPrd & "', "
      g_str_Parame = g_str_Parame & "'000', "
      g_str_Parame = g_str_Parame & Right(Format(Year(CDate(moddat_g_str_FecSis)), "0000"), 2) & ", "
      g_str_Parame = g_str_Parame & CStr(r_lng_NumOpe) & ", "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
      g_str_Parame = g_str_Parame & "1, "
      If r_lng_NumOpe = 1 Then
         g_str_Parame = g_str_Parame & "1) "
      Else
         g_str_Parame = g_str_Parame & "2) "
      End If
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_CRE_FOLIOS. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Function
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
   
   ff_Genera_NumOpe = r_str_NumOpe
End Function
Private Sub fs_DatMVi()
   If InStr(moddat_g_str_AgrCRC, moddat_g_str_CodPrd) = 0 Then   'moddat_g_str_CodPrd <> "001"
      Exit Sub
   End If
   
   l_str_OpeMVi = ""
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT EVAMVI_CODMVI "
   g_str_Parame = g_str_Parame & "  FROM TRA_EVAMVI "
   g_str_Parame = g_str_Parame & " WHERE EVAMVI_NUMSOL = '" & moddat_g_str_NumSol & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      l_str_OpeMVi = Trim(g_rst_Princi!EVAMVI_CODMVI & "")
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub
Private Sub fs_DatMVi_old()
   If InStr(moddat_g_str_AgrCRC, moddat_g_str_CodPrd) = 0 Then   'moddat_g_str_CodPrd <> "001"
      Exit Sub
   End If
   
   l_str_OpeMVi = ""
   txt_ObsMVi.Text = ""
   Call gs_LimpiaGrid(grd_Listad(12))
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM TRA_EVAMVI "
   g_str_Parame = g_str_Parame & " WHERE EVAMVI_NUMSOL = '" & moddat_g_str_NumSol & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      l_str_OpeMVi = Trim(g_rst_Princi!EVAMVI_CODMVI & "")
      
      grd_Listad(12).Rows = grd_Listad(12).Rows + 1
      grd_Listad(12).Row = grd_Listad(12).Rows - 1
      grd_Listad(12).Col = 0
      grd_Listad(12).Text = "Fecha Envío"
      
      grd_Listad(12).Col = 1
      grd_Listad(12).Text = gf_FormatoFecha(CStr(g_rst_Princi!EVAMVI_FECENV))
   
      If g_rst_Princi!EVAMVI_FECREC > 0 Then
         grd_Listad(12).Rows = grd_Listad(12).Rows + 1
         grd_Listad(12).Row = grd_Listad(12).Rows - 1
         grd_Listad(12).Col = 0
         grd_Listad(12).Text = "Fecha de Recepción"
   
         grd_Listad(12).Col = 1
         grd_Listad(12).Text = gf_FormatoFecha(CStr(g_rst_Princi!EVAMVI_FECREC))
      
         grd_Listad(12).Rows = grd_Listad(12).Rows + 1
         grd_Listad(12).Row = grd_Listad(12).Rows - 1
         grd_Listad(12).Col = 0
         grd_Listad(12).Text = "Nro. Expediente Mivivienda"
   
         grd_Listad(12).Col = 1
         grd_Listad(12).Text = Trim(g_rst_Princi!EVAMVI_CODMVI & "")
      
         txt_ObsMVi.Text = Trim(g_rst_Princi!EVAMVI_OBSERV & "")
      End If
      
      Call gs_UbiIniGrid(grd_Listad(12))
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_DatCof()
   If Not (InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_AgrTFMV, moddat_g_str_CodPrd) > 0) Then  '(moddat_g_str_CodPrd = "004" Or moddat_g_str_CodPrd = "003" Or moddat_g_str_CodPrd = "007" Or moddat_g_str_CodPrd = "009" Or moddat_g_str_CodPrd = "010" Or moddat_g_str_CodPrd = "013" Or moddat_g_str_CodPrd = "014" Or moddat_g_str_CodPrd = "015" Or moddat_g_str_CodPrd = "016" Or moddat_g_str_CodPrd = "017" Or moddat_g_str_CodPrd = "018" Or moddat_g_str_CodPrd = "019" Or moddat_g_str_CodPrd = "020" Or moddat_g_str_CodPrd = "021" Or moddat_g_str_CodPrd = "022" Or moddat_g_str_CodPrd = "023") Then
      Exit Sub
   End If
   
   l_str_OpeMVi = ""
   l_str_OpeMv1 = ""
   l_str_FecCof = ""
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT EVACOF_CODMV1, EVACOF_CODMVI, EVACOF_FECREC, EVACOF_FECDES "
   g_str_Parame = g_str_Parame & "  FROM TRA_EVACOF "
   g_str_Parame = g_str_Parame & " WHERE EVACOF_NUMSOL = '" & moddat_g_str_NumSol & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      If InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Then
         l_str_OpeMVi = Trim(g_rst_Princi!EVACOF_CODMV1 & "")
         l_str_OpeMv1 = Trim(g_rst_Princi!EVACOF_CODMVI & "")
      ElseIf InStr(moddat_g_str_AgrTFMV, moddat_g_str_CodPrd) > 0 Then
         l_str_OpeMVi = Trim(g_rst_Princi!EVACOF_CODMVI & "")
         l_str_OpeMv1 = Trim(g_rst_Princi!EVACOF_CODMV1 & "")
      End If

      If g_rst_Princi!EVACOF_FECREC > 0 Then
         l_str_FecCof = gf_FormatoFecha(CStr(g_rst_Princi!EVACOF_FECDES))
      End If
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub
Private Sub fs_DatCof_old()
   If Not (InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_AgrTFMV, moddat_g_str_CodPrd) > 0) Then  '(moddat_g_str_CodPrd = "004" Or moddat_g_str_CodPrd = "003" Or moddat_g_str_CodPrd = "007" Or moddat_g_str_CodPrd = "009" Or moddat_g_str_CodPrd = "010" Or moddat_g_str_CodPrd = "013" Or moddat_g_str_CodPrd = "014" Or moddat_g_str_CodPrd = "015" Or moddat_g_str_CodPrd = "016" Or moddat_g_str_CodPrd = "017" Or moddat_g_str_CodPrd = "018" Or moddat_g_str_CodPrd = "019" Or moddat_g_str_CodPrd = "020" Or moddat_g_str_CodPrd = "021" Or moddat_g_str_CodPrd = "022" Or moddat_g_str_CodPrd = "023") Then
      Exit Sub
   End If
   
   l_str_OpeMVi = ""
   l_str_OpeMv1 = ""
   l_str_FecCof = ""
   txt_ObsMVi.Text = ""
   Call gs_LimpiaGrid(grd_Listad(12))
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM TRA_EVACOF "
   g_str_Parame = g_str_Parame & " WHERE EVACOF_NUMSOL = '" & moddat_g_str_NumSol & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      If InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Then 'moddat_g_str_CodPrd = "003" Then
         l_str_OpeMVi = Trim(g_rst_Princi!EVACOF_CODMV1 & "")
         l_str_OpeMv1 = Trim(g_rst_Princi!EVACOF_CODMVI & "")
      ElseIf InStr(moddat_g_str_AgrTFMV, moddat_g_str_CodPrd) > 0 Then 'moddat_g_str_CodPrd = "004" Or moddat_g_str_CodPrd = "007" Or moddat_g_str_CodPrd = "009" Or moddat_g_str_CodPrd = "010" Or moddat_g_str_CodPrd = "013" Or moddat_g_str_CodPrd = "014" Or moddat_g_str_CodPrd = "015" Or moddat_g_str_CodPrd = "016" Or moddat_g_str_CodPrd = "017" Or moddat_g_str_CodPrd = "018" Or moddat_g_str_CodPrd = "019" Or moddat_g_str_CodPrd = "020" Or moddat_g_str_CodPrd = "021" Or moddat_g_str_CodPrd = "022" Or moddat_g_str_CodPrd = "023" Then
         l_str_OpeMVi = Trim(g_rst_Princi!EVACOF_CODMVI & "")
         l_str_OpeMv1 = Trim(g_rst_Princi!EVACOF_CODMV1 & "")
      End If
      
      grd_Listad(12).Rows = grd_Listad(12).Rows + 1
      grd_Listad(12).Row = grd_Listad(12).Rows - 1
      grd_Listad(12).Col = 0
      grd_Listad(12).Text = "Fecha Envío"
      
      grd_Listad(12).Col = 1
      grd_Listad(12).Text = gf_FormatoFecha(CStr(g_rst_Princi!EVACOF_FECENV))
   
      If g_rst_Princi!EVACOF_FECREC > 0 Then
         If InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Then 'moddat_g_str_CodPrd = "003" Then
            grd_Listad(12).Rows = grd_Listad(12).Rows + 1
            grd_Listad(12).Row = grd_Listad(12).Rows - 1
            grd_Listad(12).Col = 0
            grd_Listad(12).Text = "Nro. Operación Mivivienda"
      
            grd_Listad(12).Col = 1
            grd_Listad(12).Text = Trim(g_rst_Princi!EVACOF_CODMV1 & "")
            
            grd_Listad(12).Rows = grd_Listad(12).Rows + 1
            grd_Listad(12).Row = grd_Listad(12).Rows - 1
            grd_Listad(12).Col = 0
            grd_Listad(12).Text = "Fecha Aprobación Mivivienda"
      
            grd_Listad(12).Col = 1
            grd_Listad(12).Text = gf_FormatoFecha(CStr(g_rst_Princi!EVACOF_APRMVI))
         End If
         
         grd_Listad(12).Rows = grd_Listad(12).Rows + 1
         grd_Listad(12).Row = grd_Listad(12).Rows - 1
         grd_Listad(12).Col = 0
         grd_Listad(12).Text = "Nro. Carta COFIDE"
   
         grd_Listad(12).Col = 1
         grd_Listad(12).Text = Trim(g_rst_Princi!EVACOF_NUMCAR & "")
         
         grd_Listad(12).Rows = grd_Listad(12).Rows + 1
         grd_Listad(12).Row = grd_Listad(12).Rows - 1
         grd_Listad(12).Col = 0
         grd_Listad(12).Text = "Fecha Recepción Carta COFIDE"
   
         grd_Listad(12).Col = 1
         grd_Listad(12).Text = gf_FormatoFecha(CStr(g_rst_Princi!EVACOF_FECREC))
      
         grd_Listad(12).Rows = grd_Listad(12).Rows + 1
         grd_Listad(12).Row = grd_Listad(12).Rows - 1
         grd_Listad(12).Col = 0
         grd_Listad(12).Text = "Nro. Operación COFIDE"
   
         grd_Listad(12).Col = 1
         grd_Listad(12).Text = Trim(g_rst_Princi!EVACOF_CODMVI & "")
         
         grd_Listad(12).Rows = grd_Listad(12).Rows + 1
         grd_Listad(12).Row = grd_Listad(12).Rows - 1
         grd_Listad(12).Col = 0
         grd_Listad(12).Text = "Fecha Desembolso COFIDE"
   
         grd_Listad(12).Col = 1
         grd_Listad(12).Text = gf_FormatoFecha(CStr(g_rst_Princi!EVACOF_FECDES))
         
         l_str_FecCof = gf_FormatoFecha(CStr(g_rst_Princi!EVACOF_FECDES))
         
         grd_Listad(12).Rows = grd_Listad(12).Rows + 1
         grd_Listad(12).Row = grd_Listad(12).Rows - 1
         grd_Listad(12).Col = 0
         grd_Listad(12).Text = "Importe Desembolsado"
   
         grd_Listad(12).Col = 1
         grd_Listad(12).CellFontName = "Lucida Console"
         grd_Listad(12).CellFontSize = 8
         grd_Listad(12).Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!EVACOF_MTODES, 12, 2)
         
         txt_ObsMVi.Text = Trim(g_rst_Princi!EVACOF_OBSERV & "")
      End If
      
      Call gs_UbiIniGrid(grd_Listad(12))
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub
