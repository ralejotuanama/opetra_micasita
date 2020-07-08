VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_Desemb_22 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   10140
   ClientLeft      =   6675
   ClientTop       =   2400
   ClientWidth     =   11610
   Icon            =   "OpeTra_frm_116.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10140
   ScaleWidth      =   11610
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   10125
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   11640
      _Version        =   65536
      _ExtentX        =   20532
      _ExtentY        =   17859
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
         Height          =   825
         Left            =   30
         TabIndex        =   11
         Top             =   9240
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   1455
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
         Begin VB.TextBox txt_ObsDes 
            Height          =   705
            Left            =   60
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   12
            Text            =   "OpeTra_frm_116.frx":000C
            Top             =   60
            Width           =   11445
         End
      End
      Begin Threed.SSPanel SSPanel15 
         Height          =   2085
         Left            =   30
         TabIndex        =   13
         Top             =   7110
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   3678
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
         Begin MSFlexGridLib.MSFlexGrid grd_LisEva 
            Height          =   1995
            Left            =   60
            TabIndex        =   14
            Top             =   60
            Width           =   11445
            _ExtentX        =   20188
            _ExtentY        =   3519
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
      Begin Threed.SSPanel SSPanel2 
         Height          =   4815
         Left            =   30
         TabIndex        =   15
         Top             =   2250
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   8493
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
            Height          =   4695
            Left            =   60
            TabIndex        =   16
            Top             =   60
            Width           =   11385
            _ExtentX        =   20082
            _ExtentY        =   8281
            _Version        =   393216
            Style           =   1
            Tabs            =   10
            Tab             =   9
            TabsPerRow      =   10
            TabHeight       =   520
            TabCaption(0)   =   "Cliente"
            TabPicture(0)   =   "OpeTra_frm_116.frx":0010
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "grd_Listad(0)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Cónyuge"
            TabPicture(1)   =   "OpeTra_frm_116.frx":002C
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "grd_Listad(8)"
            Tab(1).Control(0).Enabled=   0   'False
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Inmueble"
            TabPicture(2)   =   "OpeTra_frm_116.frx":0048
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "grd_Listad(1)"
            Tab(2).Control(0).Enabled=   0   'False
            Tab(2).ControlCount=   1
            TabCaption(3)   =   "Crédito"
            TabPicture(3)   =   "OpeTra_frm_116.frx":0064
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "grd_Listad(2)"
            Tab(3).Control(0).Enabled=   0   'False
            Tab(3).ControlCount=   1
            TabCaption(4)   =   "Ev. Crediticia"
            TabPicture(4)   =   "OpeTra_frm_116.frx":0080
            Tab(4).ControlEnabled=   0   'False
            Tab(4).Control(0)=   "grd_Listad(3)"
            Tab(4).Control(0).Enabled=   0   'False
            Tab(4).ControlCount=   1
            TabCaption(5)   =   "Tasación"
            TabPicture(5)   =   "OpeTra_frm_116.frx":009C
            Tab(5).ControlEnabled=   0   'False
            Tab(5).Control(0)=   "grd_Listad(4)"
            Tab(5).Control(0).Enabled=   0   'False
            Tab(5).ControlCount=   1
            TabCaption(6)   =   "Seguros"
            TabPicture(6)   =   "OpeTra_frm_116.frx":00B8
            Tab(6).ControlEnabled=   0   'False
            Tab(6).Control(0)=   "Label5"
            Tab(6).Control(0).Enabled=   0   'False
            Tab(6).Control(1)=   "Label4"
            Tab(6).Control(1).Enabled=   0   'False
            Tab(6).Control(2)=   "grd_Listad(5)"
            Tab(6).Control(2).Enabled=   0   'False
            Tab(6).Control(3)=   "txt_ObsSeg"
            Tab(6).Control(3).Enabled=   0   'False
            Tab(6).ControlCount=   4
            TabCaption(7)   =   "Informe Legal"
            TabPicture(7)   =   "OpeTra_frm_116.frx":00D4
            Tab(7).ControlEnabled=   0   'False
            Tab(7).Control(0)=   "txt_InfLeg"
            Tab(7).Control(0).Enabled=   0   'False
            Tab(7).ControlCount=   1
            TabCaption(8)   =   "Ev. Legal"
            TabPicture(8)   =   "OpeTra_frm_116.frx":00F0
            Tab(8).ControlEnabled=   0   'False
            Tab(8).Control(0)=   "Label3"
            Tab(8).Control(0).Enabled=   0   'False
            Tab(8).Control(1)=   "Label2"
            Tab(8).Control(1).Enabled=   0   'False
            Tab(8).Control(2)=   "grd_Listad(6)"
            Tab(8).Control(2).Enabled=   0   'False
            Tab(8).Control(3)=   "txt_ComCre"
            Tab(8).Control(3).Enabled=   0   'False
            Tab(8).ControlCount=   4
            TabCaption(9)   =   "Mivivienda / Cofide"
            TabPicture(9)   =   "OpeTra_frm_116.frx":010C
            Tab(9).ControlEnabled=   -1  'True
            Tab(9).Control(0)=   "Label7"
            Tab(9).Control(0).Enabled=   0   'False
            Tab(9).Control(1)=   "Label6"
            Tab(9).Control(1).Enabled=   0   'False
            Tab(9).Control(2)=   "grd_Listad(7)"
            Tab(9).Control(2).Enabled=   0   'False
            Tab(9).Control(3)=   "txt_ObsMVi"
            Tab(9).Control(3).Enabled=   0   'False
            Tab(9).ControlCount=   4
            Begin VB.TextBox txt_ObsSeg 
               Height          =   1065
               Left            =   -74940
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   40
               Text            =   "OpeTra_frm_116.frx":0128
               Top             =   3540
               Width           =   11235
            End
            Begin VB.TextBox txt_InfLeg 
               Height          =   4215
               Left            =   -74940
               MaxLength       =   8000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   39
               Text            =   "OpeTra_frm_116.frx":012C
               Top             =   390
               Width           =   11235
            End
            Begin VB.TextBox txt_ComCre 
               Height          =   705
               Left            =   -74940
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   35
               Text            =   "OpeTra_frm_116.frx":0130
               Top             =   660
               Width           =   11235
            End
            Begin VB.TextBox txt_ObsMVi 
               Height          =   1155
               Left            =   90
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   31
               Text            =   "OpeTra_frm_116.frx":0134
               Top             =   3480
               Width           =   11235
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   4215
               Index           =   0
               Left            =   -74940
               TabIndex        =   17
               Top             =   390
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   7435
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
               Height          =   2535
               Index           =   7
               Left            =   90
               TabIndex        =   32
               Top             =   630
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   4471
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
               Height          =   2955
               Index           =   6
               Left            =   -74940
               TabIndex        =   36
               Top             =   1680
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   5212
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
               Height          =   2595
               Index           =   5
               Left            =   -74940
               TabIndex        =   41
               Top             =   630
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   4577
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
               Height          =   4215
               Index           =   4
               Left            =   -74940
               TabIndex        =   44
               Top             =   420
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   7435
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
               Height          =   4215
               Index           =   3
               Left            =   -74940
               TabIndex        =   45
               Top             =   420
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   7435
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
               Height          =   4215
               Index           =   2
               Left            =   -74940
               TabIndex        =   46
               Top             =   420
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   7435
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
               Height          =   4215
               Index           =   1
               Left            =   -74940
               TabIndex        =   47
               Top             =   420
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   7435
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
               Height          =   4215
               Index           =   8
               Left            =   -74940
               TabIndex        =   48
               Top             =   420
               Width           =   11235
               _ExtentX        =   19817
               _ExtentY        =   7435
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin VB.Label Label4 
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
               TabIndex        =   43
               Top             =   3300
               Width           =   3495
            End
            Begin VB.Label Label5 
               Caption         =   "Datos de la Evaluación"
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
               Top             =   390
               Width           =   3495
            End
            Begin VB.Label Label2 
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
               TabIndex        =   38
               Top             =   420
               Width           =   3495
            End
            Begin VB.Label Label3 
               Caption         =   "Datos de la Evaluación"
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
               TabIndex        =   37
               Top             =   1440
               Width           =   3495
            End
            Begin VB.Label Label6 
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
               Left            =   90
               TabIndex        =   34
               Top             =   3240
               Width           =   3495
            End
            Begin VB.Label Label7 
               Caption         =   "Datos de la Evaluación"
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
               Left            =   90
               TabIndex        =   33
               Top             =   390
               Width           =   3495
            End
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   18
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
         Begin MSMAPI.MAPIMessages mps_Mensaj 
            Left            =   9300
            Top             =   60
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
            Left            =   8730
            Top             =   60
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            DownloadMail    =   -1  'True
            LogonUI         =   -1  'True
            NewSession      =   0   'False
         End
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   8130
            Top             =   90
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
         Begin MSComDlg.CommonDialog dlg_Guarda 
            Left            =   10140
            Top             =   90
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            CancelError     =   -1  'True
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   315
            Left            =   690
            TabIndex        =   19
            Top             =   30
            Width           =   6495
            _Version        =   65536
            _ExtentX        =   11456
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Créditos Hipotecarios"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
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
         Begin Threed.SSPanel SSPanel68 
            Height          =   315
            Left            =   690
            TabIndex        =   20
            Top             =   330
            Width           =   6495
            _Version        =   65536
            _ExtentX        =   11456
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Desembolso - Datos de Operación"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
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
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_116.frx":0138
            Stretch         =   -1  'True
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   30
         TabIndex        =   21
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
         Begin Threed.SSPanel pnl_Produc 
            Height          =   315
            Left            =   1440
            TabIndex        =   22
            Top             =   60
            Width           =   10065
            _Version        =   65536
            _ExtentX        =   17754
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "CREDITO HIPOTECARIO MICASITA"
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
         Begin Threed.SSPanel pnl_NumSol 
            Height          =   315
            Left            =   8580
            TabIndex        =   23
            Top             =   390
            Width           =   2925
            _Version        =   65536
            _ExtentX        =   5159
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
         Begin Threed.SSPanel pnl_NumOpe 
            Height          =   315
            Left            =   1440
            TabIndex        =   24
            Top             =   390
            Width           =   2925
            _Version        =   65536
            _ExtentX        =   5159
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
         Begin VB.Label Label12 
            Caption         =   "Nro. Operación:"
            Height          =   315
            Left            =   60
            TabIndex        =   27
            Top             =   390
            Width           =   1245
         End
         Begin VB.Label Label21 
            Caption         =   "Producto:"
            Height          =   315
            Left            =   60
            TabIndex        =   26
            Top             =   60
            Width           =   1275
         End
         Begin VB.Label Label1 
            Caption         =   "Nro. Solicitud:"
            Height          =   315
            Left            =   7200
            TabIndex        =   25
            Top             =   390
            Width           =   1335
         End
      End
      Begin Threed.SSPanel SSPanel39 
         Height          =   645
         Left            =   30
         TabIndex        =   28
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
         Begin VB.CommandButton cmd_Tasaci 
            Height          =   585
            Left            =   5400
            Picture         =   "OpeTra_frm_116.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   30
            ToolTipText     =   "Modificar Informe de Tasación"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_DatInm 
            Height          =   585
            Left            =   6000
            Picture         =   "OpeTra_frm_116.frx":0D0C
            Style           =   1  'Graphical
            TabIndex        =   29
            ToolTipText     =   "Modificar Datos del Inmueble"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_EnvLeg 
            Height          =   585
            Left            =   4830
            Picture         =   "OpeTra_frm_116.frx":15D6
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Envío a Legal"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Export 
            Height          =   585
            Left            =   1230
            Picture         =   "OpeTra_frm_116.frx":18E0
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Exportación de Cronogramas para Mivivienda"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   630
            Picture         =   "OpeTra_frm_116.frx":1BEA
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Imprimir Formatos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Desemb 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_116.frx":202C
            Style           =   1  'Graphical
            TabIndex        =   0
            ToolTipText     =   "Generar Desembolso"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10920
            Picture         =   "OpeTra_frm_116.frx":2336
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_PolSeg 
            Height          =   585
            Left            =   1830
            Picture         =   "OpeTra_frm_116.frx":2778
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Regularizar Póliza de Seguros"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ChqGer 
            Height          =   585
            Left            =   2430
            Picture         =   "OpeTra_frm_116.frx":2A82
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Regularizar Cheque de Gerencia"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_CarFia 
            Height          =   585
            Left            =   3030
            Picture         =   "OpeTra_frm_116.frx":2D8C
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Regularizar Carta Fianza"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_CerPar 
            Height          =   585
            Left            =   3630
            Picture         =   "OpeTra_frm_116.frx":3096
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Regularizar Certificado de Participación"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_DiaPag 
            Height          =   585
            Left            =   4230
            Picture         =   "OpeTra_frm_116.frx":33A0
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Cambiar Día de Pago"
            Top             =   30
            Width           =   585
         End
      End
   End
End
Attribute VB_Name = "frm_Desemb_22"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_str_MonBlq     As String
Dim l_dbl_ImpTas     As Double
Dim l_dbl_ImpNot     As Double
Dim l_dbl_ImpEst     As Double
Dim l_dbl_ImpEva     As Double
Dim l_dbl_ImpAdm     As Double
Dim l_dbl_ImpRed     As Double
Dim l_dbl_ImpBlq     As Double
Dim l_int_ChqReg     As Integer
Dim l_int_PolReg     As Integer
Dim l_int_FiaReg     As Integer
Dim l_int_CerReg     As Integer
Dim l_int_FlgCVt     As Integer
Dim l_int_MonCvt     As Integer

Private Sub cmd_CarFia_Click()
   If l_int_FiaReg <> 1 Then
      MsgBox "No hay Carta Fianza pendiente de regularización.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   moddat_g_int_FlgAct = 1
   frm_Desemb_15.Show 1
   
   If moddat_g_int_FlgAct = 2 Then
      Screen.MousePointer = 11
      Call fs_Carga_DatEva
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmd_CerPar_Click()
   If l_int_CerReg <> 1 Then
      MsgBox "No hay Certificado de Participación pendiente de regularización.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   moddat_g_int_FlgAct = 1
   frm_Desemb_16.Show 1
   
   If moddat_g_int_FlgAct = 2 Then
      Screen.MousePointer = 11
      Call fs_Carga_DatEva
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmd_ChqGer_Click()
   If l_int_ChqReg <> 1 Then
      MsgBox "No hay Cheque de Gerencia pendiente de regularización.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   moddat_g_int_FlgAct = 1
   frm_Desemb_13.Show 1
   
   If moddat_g_int_FlgAct = 2 Then
      Screen.MousePointer = 11
      Call fs_Carga_DatEva
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmd_DatInm_Click()
    moddat_g_int_FlgGrb = 2
    moddat_g_int_FlgAct = 1
   
    'frm_Tra_EvaTas_05.Show 1
    frm_Seg_SolHip_54.Show 1
   
    If moddat_g_int_FlgAct = 2 Then
        Screen.MousePointer = 11
        Call gs_LimpiaGrid(grd_Listad(1))
        Call modmip_gs_DatInm(grd_Listad(1), True)
        Screen.MousePointer = 0
    End If
End Sub

Private Sub cmd_Desemb_Click()
   If grd_LisEva.Rows > 0 Then
      MsgBox "Ya registro la información del Desembolso.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   moddat_g_int_FlgAct = 1
   frm_Desemb_11.Show 1
   
   If moddat_g_int_FlgAct = 2 Then
      Screen.MousePointer = 11
      Call fs_Carga_DatEva
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmd_DiaPag_Click()
   If grd_LisEva.Rows > 0 Then
      MsgBox "Ya registro la información del Desembolso.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   moddat_g_int_FlgAct = 1
   frm_Desemb_17.Show 1
   
   If moddat_g_int_FlgAct = 2 Then
      Screen.MousePointer = 11
      Call fs_DatCre
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmd_EnvLeg_Click()
   'Verificando si existe registro de envío a Legal
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM TRA_SEGDES "
   g_str_Parame = g_str_Parame & " WHERE SEGDES_NUMSOL = '" & moddat_g_str_NumSol & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      If g_rst_Princi!SEGDES_FECLEG <> 0 Then
         MsgBox "Ya registro el envío a Legal. ", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
      moddat_g_int_FlgGrb = 2
   Else
      moddat_g_int_FlgGrb = 1
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   frm_Desemb_18.Show 1
End Sub

Private Sub cmd_Export_Click()
   frm_Desemb_19.Show 1
End Sub

Private Sub cmd_Imprim_Click()
   If grd_LisEva.Rows = 0 Then
      MsgBox "No ha registrado la información del Desembolso.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   moddat_g_int_InsAct = l_int_FlgCVt
   moddat_g_int_TipRep = 0
   
   'frm_Desemb_06_.Show 1
   frm_Desemb_06.Show 1
   
'   If moddat_g_int_TipRep > 0 Then
'      Select Case moddat_g_int_TipRep
'         Case 1:     Call fs_LiqDes
'         Case 2:     Call fs_HojRes
'         Case 3:     Call fs_Cronog_MiCasita
'         Case 4:     Call fs_Cronog_Mivivienda_NCoCli
'         Case 5:     Call fs_Cronog_Mivivienda_ConCli
'         Case 7, 9:  Call fs_Cronog_Mivivienda_ConMVi
'         Case 8:     Call fs_Cronog_Mivivienda_NCoMVi
'         Case 10:    Call fs_Cronog_Mivivienda_NCoCof
'         Case 11:    Call fs_ComPag
'         Case 12
'            If l_int_FlgCVt = 0 Then
'               MsgBox "No se puede emitir este formato, porque no se realizón ninguna operación.", vbInformation, modgen_g_str_NomPlt
'               Exit Sub
'            End If
'            If l_int_MonCvt = moddat_g_int_TipMon Then
'               MsgBox "No se puede emitir este formato, porque la Moneda de Compra-Venta es igual a la Moneda de Préstamo.", vbInformation, modgen_g_str_NomPlt
'               Exit Sub
'            End If
'            Call fs_LiqTipoCambio
'      End Select
'   End If
End Sub

Private Sub cmd_PolSeg_Click()
   If l_int_PolReg <> 1 Then
      MsgBox "No hay Póliza de Seguro pendiente de regularización.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   moddat_g_int_FlgAct = 1
   frm_Desemb_14.Show 1
   
   If moddat_g_int_FlgAct = 2 Then
      Screen.MousePointer = 11
      Call fs_DatSeg
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub cmd_Tasaci_Click()
    If moddat_g_int_CodIns < 51 Then
        MsgBox "No se puede modificar el Informe de Tasación.", vbExclamation, modgen_g_str_NomPlt
        Exit Sub
    End If

    moddat_g_int_FlgAct = 1
    frm_ModSol_06.Show 1
   
    If moddat_g_int_FlgAct = 2 Then
        Screen.MousePointer = 11
        Call modmip_gs_EvaTas(grd_Listad(4))                'Call fs_DatTas              'Refresca Datos de Tasación
        Screen.MousePointer = 0
    End If
End Sub

Private Sub fs_Buscar_LisOcu()
   Dim r_str_FecOcu  As String
   
'   Call gs_LimpiaGrid(grd_LisOcu)
'
'   g_str_Parame = ""
'   g_str_Parame = g_str_Parame & "SELECT * "
'   g_str_Parame = g_str_Parame & "  FROM TRA_SEGDET "
'   g_str_Parame = g_str_Parame & " WHERE SEGDET_NUMSOL = '" & moddat_g_str_NumSol & "' "
'   g_str_Parame = g_str_Parame & " ORDER BY SEGFECCRE DESC, SEGHORCRE DESC, SEGDET_CODINS DESC "
'
'   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
'       Exit Sub
'   End If
'
'   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
'     g_rst_Princi.Close
'     Set g_rst_Princi = Nothing
'     Exit Sub
'   End If
'
'   grd_LisOcu.Redraw = False
'   g_rst_Princi.MoveFirst
'   Do While Not g_rst_Princi.EOF
'      grd_LisOcu.Rows = grd_LisOcu.Rows + 1
'      grd_LisOcu.Row = grd_LisOcu.Rows - 1
'
'      'Fecha de Ocurrencia
'      grd_LisOcu.Col = 0
'      grd_LisOcu.Text = gf_FormatoFecha(CStr(g_rst_Princi!SEGFECCRE))
'
'      'Hora de Ocurrencia
'      grd_LisOcu.Col = 1
'      grd_LisOcu.Text = gf_FormatoHora(Format(g_rst_Princi!SEGHORCRE, "000000"))
'
'      'Instancia de Evaluación
'      grd_LisOcu.Col = 2
'      grd_LisOcu.Text = moddat_gf_Consulta_ParDes("002", CStr(g_rst_Princi!SEGDET_CODINS))
'
'      'Descripción Ocurrencia
'      grd_LisOcu.Col = 3
'      grd_LisOcu.Text = moddat_gf_Consulta_ParDes("004", Format(g_rst_Princi!SEGDET_CODOCU, "000000"))
'
'      If g_rst_Princi!SEGFECACT > 0 Then
'         r_str_FecOcu = gf_FormatoFecha(CStr(g_rst_Princi!SEGFECACT))
'         grd_LisOcu.Text = grd_LisOcu.Text & " (DESCARGO EFECTUADO - " & r_str_FecOcu
'         grd_LisOcu.Text = grd_LisOcu.Text & " / " & gf_FormatoHora(Format(g_rst_Princi!SEGHORACT, "000000")) & ")"
'      End If
'
'      grd_LisOcu.Col = 4
'      grd_LisOcu.Text = Trim(g_rst_Princi!SEGDET_OBSERV & "")
'
'      grd_LisOcu.Col = 5
'      grd_LisOcu.Text = Trim(g_rst_Princi!SEGDET_OBSDES & "")
'
'      g_rst_Princi.MoveNext
'   Loop
'
'   grd_LisOcu.Redraw = True
'
'   g_rst_Princi.Close
'   Set g_rst_Princi = Nothing
'
'   Call gs_UbiIniGrid(grd_LisOcu)
'   Call grd_LisOcu_Click
End Sub
Private Sub fs_Buscar_LisExc()
Dim r_str_FecOcu  As String
   
'   Call gs_LimpiaGrid(grd_LisExc)
'
'   g_str_Parame = modgen_gf_Buscar_Excepc
'
'   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
'       Exit Sub
'   End If
'
'   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
'     g_rst_Princi.Close
'     Set g_rst_Princi = Nothing
'     Exit Sub
'   End If
'
'   grd_LisExc.Redraw = False
'   g_rst_Princi.MoveFirst
'   Do While Not g_rst_Princi.EOF
'      grd_LisExc.Rows = grd_LisExc.Rows + 1
'      grd_LisExc.Row = grd_LisExc.Rows - 1
'
'      'Fecha de Excepción
'      grd_LisExc.Col = 0
'      grd_LisExc.Text = gf_FormatoFecha(CStr(g_rst_Princi!SEGFECCRE))
'
'      'Hora de Excepción
'      grd_LisExc.Col = 1
'      grd_LisExc.Text = gf_FormatoHora(Format(g_rst_Princi!SEGHORCRE, "000000"))
'
'      'Instancia
'      grd_LisExc.Col = 2
'      grd_LisExc.Text = moddat_gf_Consulta_ParDes("002", CStr(g_rst_Princi!SEGEXC_CODINS))
'
'      'Descripción Excepción
'      grd_LisExc.Col = 3
'      grd_LisExc.Text = Trim(g_rst_Princi!SEGEXC_DESCRI & "")
'
'      'Tipo Autorización
'      grd_LisExc.Col = 4
'      grd_LisExc.Text = moddat_gf_Consulta_ParDes("243", CStr(g_rst_Princi!SEGEXC_TIPAUT))
'
'      'Motivo de Excepción
'      grd_LisExc.Col = 5
'      grd_LisExc.Text = Trim(g_rst_Princi!PARDES_DESCRI)
'
'      g_rst_Princi.MoveNext
'   Loop
'
'   grd_LisExc.Redraw = True
'
'   g_rst_Princi.Close
'   Set g_rst_Princi = Nothing
'
'   Call gs_UbiIniGrid(grd_LisExc)
'   Call grd_LisExc_Click
End Sub

Private Sub Form_Load()
Dim r_arr_Mtz()      As moddat_g_tpo_DatCom
   
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   pnl_NumSol.Caption = Mid(moddat_g_str_NumSol, 1, 3) & "-" & Mid(moddat_g_str_NumSol, 4, 3) & "-" & Mid(moddat_g_str_NumSol, 7, 2) & "-" & Mid(moddat_g_str_NumSol, 9, 4)
   pnl_NumOpe.Caption = Mid(moddat_g_str_NumOpe, 1, 3) & "-" & Mid(moddat_g_str_NumOpe, 4, 2) & "-" & Mid(moddat_g_str_NumOpe, 6, 5)
   pnl_Produc.Caption = moddat_g_str_NomPrd
   
   Call fs_Inicia

   moddat_g_int_CygTDo = 0
   moddat_g_str_CygNDo = ""
   Call modmip_gs_DatCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc, grd_Listad(0), 0)      'Buscar Información del Cliente
   Call modmip_gs_DatCli(moddat_g_int_CygTDo, moddat_g_str_CygNDo, grd_Listad(8), 1)      'Buscar Información del Cónyuge
   Call modmip_gs_DatInm(grd_Listad(1), True)
   
   'Call fs_DatCre
   Call modmip_gs_DatCre(grd_Listad(2), r_arr_Mtz)
   moddat_g_str_CodEjeSeg = r_arr_Mtz(0).DatCom_EjeSeg
   moddat_g_str_CodConHip = r_arr_Mtz(0).DatCom_ConHip
   
   Call fs_EvaCre
   Call modmip_gs_EvaTas(grd_Listad(4))                                                   'Call fs_DatTas
   Call modmip_gs_EvaSeg(grd_Listad(5))
   Call fs_DatSeg
   Call modmip_gs_Buscar_EvaLeg(grd_Listad(6), grd_Listad(6), txt_InfLeg, txt_ComCre)
   Call fs_DatLeg
   Call modmip_gs_TraMVi(grd_Listad(7), txt_ObsMVi)                                       'Call fs_DatMVi
   Call modmip_gs_Buscar_TraCof(grd_Listad(7), txt_ObsMVi)                                'Call fs_DatCof
   Call fs_GasAdm
   Call fs_Carga_DatEva
   Call gs_CentraForm(Me)
   
   If Not (InStr(moddat_g_str_AgrCRC, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0) Then  '(moddat_g_str_CodPrd = "001" Or moddat_g_str_CodPrd = "003") Then
      cmd_Export.Enabled = False
   End If
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   'Datos del Cliente
   grd_Listad(0).ColWidth(0) = 3000:   grd_Listad(0).ColAlignment(0) = flexAlignLeftCenter
   grd_Listad(0).ColWidth(1) = 7940:   grd_Listad(0).ColAlignment(1) = flexAlignLeftCenter
   Call gs_LimpiaGrid(grd_Listad(0))
   
   'Datos del Cónyuge
   grd_Listad(8).ColWidth(0) = 3000:   grd_Listad(8).ColAlignment(0) = flexAlignLeftCenter
   grd_Listad(8).ColWidth(1) = 7940:   grd_Listad(8).ColAlignment(1) = flexAlignLeftCenter
   Call gs_LimpiaGrid(grd_Listad(8))

   'Datos del Inmueble
   grd_Listad(1).ColWidth(0) = 3000:   grd_Listad(1).ColAlignment(0) = flexAlignLeftCenter
   grd_Listad(1).ColWidth(1) = 7940:   grd_Listad(1).ColAlignment(1) = flexAlignLeftCenter
   Call gs_LimpiaGrid(grd_Listad(1))

   'Datos del Crédito
   grd_Listad(2).ColWidth(0) = 3000:   grd_Listad(2).ColAlignment(0) = flexAlignLeftCenter
   grd_Listad(2).ColWidth(1) = 7940:   grd_Listad(2).ColAlignment(1) = flexAlignLeftCenter
   Call gs_LimpiaGrid(grd_Listad(2))

   'Evaluación Crediticia
   grd_Listad(3).ColWidth(0) = 3000:   grd_Listad(3).ColAlignment(0) = flexAlignLeftCenter
   grd_Listad(3).ColWidth(1) = 7940:   grd_Listad(3).ColAlignment(1) = flexAlignLeftCenter
   Call gs_LimpiaGrid(grd_Listad(3))

   'Tasación del Inmueble
   grd_Listad(4).ColWidth(0) = 3000:   grd_Listad(4).ColAlignment(0) = flexAlignLeftCenter
   grd_Listad(4).ColWidth(1) = 7940:   grd_Listad(4).ColAlignment(1) = flexAlignLeftCenter
   Call gs_LimpiaGrid(grd_Listad(4))

   'Evaluación de Seguros
   grd_Listad(5).ColWidth(0) = 3000:   grd_Listad(5).ColAlignment(0) = flexAlignLeftCenter
   grd_Listad(5).ColWidth(1) = 7940:   grd_Listad(5).ColAlignment(1) = flexAlignLeftCenter
   Call gs_LimpiaGrid(grd_Listad(5))

   'Evaluación Legal
   grd_Listad(6).ColWidth(0) = 3000:   grd_Listad(6).ColAlignment(0) = flexAlignLeftCenter
   grd_Listad(6).ColWidth(1) = 7940:   grd_Listad(6).ColAlignment(1) = flexAlignLeftCenter
   Call gs_LimpiaGrid(grd_Listad(6))

   'Trámites Mivivienda / Cofide
   grd_Listad(7).ColWidth(0) = 3000:   grd_Listad(7).ColAlignment(0) = flexAlignLeftCenter
   grd_Listad(7).ColWidth(1) = 7940:   grd_Listad(7).ColAlignment(1) = flexAlignLeftCenter
   Call gs_LimpiaGrid(grd_Listad(7))

   'Grid de Evaluación
   grd_LisEva.ColWidth(0) = 3200
   grd_LisEva.ColWidth(1) = 7940
   grd_LisEva.ColAlignment(0) = flexAlignLeftCenter
   grd_LisEva.ColAlignment(1) = flexAlignLeftCenter
   
   txt_ObsSeg.Text = ""
   txt_InfLeg.Text = ""
   txt_ComCre.Text = ""
   txt_ObsMVi.Text = ""
   l_int_ChqReg = 0
   l_int_PolReg = 0
   l_int_FiaReg = 0
   l_int_CerReg = 0
End Sub

Private Sub grd_LisEva_SelChange()
   If grd_LisEva.Rows > 2 Then
      grd_LisEva.RowSel = grd_LisEva.Row
   End If
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

Private Sub txt_ObsDes_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txt_ObsMVi_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub txt_ObsSeg_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub

Private Sub fs_Carga_DatEva()
   Call gs_LimpiaGrid(grd_LisEva)
   txt_ObsDes.Text = ""
   l_int_FlgCVt = 0

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPDES "
   g_str_Parame = g_str_Parame & " WHERE HIPDES_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      grd_LisEva.Rows = grd_LisEva.Rows + 1
      grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0
      grd_LisEva.Text = "Fecha de Desembolso"
      
      grd_LisEva.Col = 1
      grd_LisEva.Text = gf_FormatoFecha(g_rst_Princi!HIPDES_FECDES)
      
      grd_LisEva.Rows = grd_LisEva.Rows + 1
      grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0
      grd_LisEva.Text = "Tipo de Desembolso"
      
      grd_LisEva.Col = 1
      grd_LisEva.Text = "CONTRA " & moddat_gf_Consulta_ParDes("241", g_rst_Princi!HIPDES_TIPGAR)
      
      If g_rst_Princi!HIPDES_TIPGAR = 2 Or g_rst_Princi!HIPDES_TIPGAR = 4 Or g_rst_Princi!HIPDES_TIPGAR = 5 Or g_rst_Princi!HIPDES_TIPGAR = 3 Then
         grd_LisEva.Rows = grd_LisEva.Rows + 1
         grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0
         grd_LisEva.Text = "Forma de Desembolso"
         
         grd_LisEva.Col = 1
         grd_LisEva.Text = moddat_gf_Consulta_ParDes("226", g_rst_Princi!HIPDES_TIPDES)
      End If
      
      If g_rst_Princi!HIPDES_TIPDES = 1 Then
         If Len(Trim(g_rst_Princi!HIPDES_CHECGO & "")) > 0 Then
            grd_LisEva.Rows = grd_LisEva.Rows + 1
            grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0
            grd_LisEva.Text = "Nro. de Cheque"
            
            grd_LisEva.Col = 1
            grd_LisEva.Text = Trim(g_rst_Princi!HIPDES_CHECGO & "")
            
            grd_LisEva.Rows = grd_LisEva.Rows + 1
            grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0
            grd_LisEva.Text = "Banco Emisor (Cuenta)"
            
            grd_LisEva.Col = 1
            grd_LisEva.Text = moddat_gf_Consulta_ParDes("516", g_rst_Princi!HIPDES_BANCGO & "") & " (" & Trim(g_rst_Princi!HIPDES_CTACGO & "") & ")"
         Else
            grd_LisEva.Rows = grd_LisEva.Rows + 1
            grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0
            grd_LisEva.Text = "Nro. de Cheque"
            
            grd_LisEva.Col = 1
            grd_LisEva.Text = "CHEQUE NO EMITIDO"
            
            l_int_ChqReg = 1
         End If
      End If
      
      grd_LisEva.Rows = grd_LisEva.Rows + 1
      grd_LisEva.Row = grd_LisEva.Rows - 1
      grd_LisEva.Col = 0
      grd_LisEva.Text = "Importe Desembolsado"
      
      grd_LisEva.Col = 1
      grd_LisEva.CellFontName = "Lucida Console"
      grd_LisEva.CellFontSize = 8
      grd_LisEva.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPDES_DESMPR, 12, 2)
      
      If g_rst_Princi!HIPDES_TIPGAR = 4 Then
         If Len(Trim(g_rst_Princi!HIPDES_NUMFIA & "")) > 0 Then
            grd_LisEva.Rows = grd_LisEva.Rows + 2
            grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0
            grd_LisEva.Text = "Nro. Carta Fianza"
            
            grd_LisEva.Col = 1
            grd_LisEva.Text = Trim(g_rst_Princi!HIPDES_NUMFIA & "")
            
            grd_LisEva.Rows = grd_LisEva.Rows + 1
            grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0
            grd_LisEva.Text = "Banco Emisor "
            
            grd_LisEva.Col = 1
            grd_LisEva.Text = moddat_gf_Consulta_ParDes("505", g_rst_Princi!HIPDES_BANFIA)
         
            grd_LisEva.Rows = grd_LisEva.Rows + 1
            grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0
            grd_LisEva.Text = "Fecha Emisión"
            
            grd_LisEva.Col = 1
            grd_LisEva.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPDES_EMIFIA))
         
            grd_LisEva.Rows = grd_LisEva.Rows + 1
            grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0
            grd_LisEva.Text = "Fecha Vencimiento"
            
            grd_LisEva.Col = 1
            grd_LisEva.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPDES_VCTFIA))
            
            grd_LisEva.Rows = grd_LisEva.Rows + 1
            grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0
            grd_LisEva.Text = "Importe Carta Fianza"
            
            grd_LisEva.Col = 1
            grd_LisEva.CellFontName = "Lucida Console"
            grd_LisEva.CellFontSize = 8
            grd_LisEva.Text = moddat_gf_Consulta_ParDes("229", g_rst_Princi!HIPDES_MONFIA) & " " & gf_FormatoNumero(g_rst_Princi!HIPDES_IMPFIA, 12, 2)
         Else
            grd_LisEva.Rows = grd_LisEva.Rows + 2
            grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0
            grd_LisEva.Text = "Nro. Carta Fianza"
            
            grd_LisEva.Col = 1
            grd_LisEva.Text = "CARTA FIANZA NO RECIBIDA"
            
            l_int_FiaReg = 1
         End If
      End If
      
      If g_rst_Princi!HIPDES_TIPGAR = 5 Then
         If Len(Trim(g_rst_Princi!HIPDES_DOCGAR & "")) > 0 Then
            grd_LisEva.Rows = grd_LisEva.Rows + 2
            grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0
            grd_LisEva.Text = "Nro. Certificado de Participación"
            
            grd_LisEva.Col = 1
            grd_LisEva.Text = Trim(g_rst_Princi!HIPDES_DOCGAR & "")
            
            grd_LisEva.Rows = grd_LisEva.Rows + 1
            grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0
            grd_LisEva.Text = "Banco Emisor "
            
            grd_LisEva.Col = 1
            grd_LisEva.Text = moddat_gf_Consulta_ParDes("505", g_rst_Princi!HIPDES_BCOGAR)
            
            grd_LisEva.Rows = grd_LisEva.Rows + 1
            grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0
            grd_LisEva.Text = "Importe Certificado"
            
            grd_LisEva.Col = 1
            grd_LisEva.CellFontName = "Lucida Console"
            grd_LisEva.CellFontSize = 8
            grd_LisEva.Text = moddat_gf_Consulta_ParDes("229", g_rst_Princi!HIPDES_MONGAR) & " " & gf_FormatoNumero(g_rst_Princi!HIPDES_MTOGAR, 12, 2)
         Else
            grd_LisEva.Rows = grd_LisEva.Rows + 2
            grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0
            grd_LisEva.Text = "Nro. Certificado de Participación"
            
            grd_LisEva.Col = 1
            grd_LisEva.Text = "CERTIFICADO NO RECIBIDO"
            
            l_int_CerReg = 1
         End If
      End If
            
      Call gs_UbiIniGrid(grd_LisEva)
      txt_ObsDes.Text = Trim(g_rst_Princi!HIPDES_OBSERV & "")
      
      If Not IsNull(g_rst_Princi!HIPDES_MONCVT) Then
         l_int_FlgCVt = 1
      End If
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_DatCre()
   Call gs_LimpiaGrid(grd_Listad(2))
   
   'Buscando Información del Crédito
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT HIPMAE_PRIVIV, HIPMAE_CVTSOL, HIPMAE_APOSOL, HIPMAE_CVTDOL, HIPMAE_APODOL, HIPMAE_MTOPRE, HIPMAE_FECESC, "
   g_str_Parame = g_str_Parame & "       HIPMAE_FECESC, HIPMAE_PLAANO, HIPMAE_TASINT, HIPMAE_NUMCUO, HIPMAE_PERGRA, HIPMAE_CUOANO, HIPMAE_DIAPAG, "
   g_str_Parame = g_str_Parame & "       HIPMAE_INTCAP, HIPMAE_SEGPRE, HIPMAE_SEGPRE, HIPMAE_TIPSEG, HIPMAE_CONHIP, HIPMAE_CODPRD, SOLMAE_FMVBBP  "
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
         grd_Listad(2).Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_APOSOL, 12, 2) & "  (INCLUYE BBP " & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & Format(g_rst_Princi!SOLMAE_FMVBBP, "##,###,##0.00") & ")"
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

Private Sub fs_EvaCre()
   Call gs_LimpiaGrid(grd_Listad(3))
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM CRE_SOLMAE "
   g_str_Parame = g_str_Parame & " WHERE SOLMAE_NUMERO = '" & moddat_g_str_NumSol & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      grd_Listad(3).Rows = grd_Listad(3).Rows + 1
      grd_Listad(3).Row = grd_Listad(3).Rows - 1
      grd_Listad(3).Col = 0
      grd_Listad(3).Text = "Cuota Aceptada por Cliente"
      
      grd_Listad(3).Col = 1
      grd_Listad(3).CellFontName = "Lucida Console"
      grd_Listad(3).CellFontSize = 8
      grd_Listad(3).Text = "S/. " & gf_FormatoNumero(g_rst_Princi!SOLMAE_CUOAPR_SOL, 12, 2)
      
      If moddat_g_int_TipMon <> 1 Then
         grd_Listad(3).Rows = grd_Listad(3).Rows + 1
         grd_Listad(3).Row = grd_Listad(3).Rows - 1
         grd_Listad(3).Col = 0
         grd_Listad(3).Text = "Cuota Aceptada por Cliente (M. Prest.)"
         
         grd_Listad(3).Col = 1
         grd_Listad(3).CellFontName = "Lucida Console"
         grd_Listad(3).CellFontSize = 8
         grd_Listad(3).Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_CUOAPR_MPR, 12, 2)
         
         grd_Listad(3).Rows = grd_Listad(3).Rows + 1
      End If
      
      grd_Listad(3).Rows = grd_Listad(3).Rows + 1
      grd_Listad(3).Row = grd_Listad(3).Rows - 1
      grd_Listad(3).Col = 0
      grd_Listad(3).Text = "Cuota Aprobada"
      
      grd_Listad(3).Col = 1
      grd_Listad(3).CellFontName = "Lucida Console"
      grd_Listad(3).CellFontSize = 8
      grd_Listad(3).Text = "S/. " & gf_FormatoNumero(g_rst_Princi!SOLMAE_CUOMEN_SOL, 12, 2)
      
      If moddat_g_int_TipMon <> 1 Then
         grd_Listad(3).Rows = grd_Listad(3).Rows + 1
         grd_Listad(3).Row = grd_Listad(3).Rows - 1
         grd_Listad(3).Col = 0
         grd_Listad(3).Text = "Cuota Aprobada (M. Prest.)"
         
         grd_Listad(3).Col = 1
         grd_Listad(3).CellFontName = "Lucida Console"
         grd_Listad(3).CellFontSize = 8
         grd_Listad(3).Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_CUOMEN_MPR, 12, 2)
         
         grd_Listad(3).Rows = grd_Listad(3).Rows + 1
      End If
   
      grd_Listad(3).Rows = grd_Listad(3).Rows + 1
      grd_Listad(3).Row = grd_Listad(3).Rows - 1
      grd_Listad(3).Col = 0
      grd_Listad(3).Text = "Total Ingreso Líquido"
      
      grd_Listad(3).Col = 1
      grd_Listad(3).CellFontName = "Lucida Console"
      grd_Listad(3).CellFontSize = 8
      grd_Listad(3).Text = "S/. " & gf_FormatoNumero(g_rst_Princi!SOLMAE_INGNET, 12, 2)
   End If
   
   moddat_g_str_CodEjeSeg = Trim(g_rst_Princi!SOLMAE_EJESEG & "")
   moddat_g_str_CodConHip = Trim(g_rst_Princi!SOLMAE_CONHIP & "")
   
    moddat_g_str_CodPrd = Trim(g_rst_Princi!SOLMAE_CODPRD)
    moddat_g_str_CodSub = Trim(g_rst_Princi!SOLMAE_CODSUB)
    moddat_g_str_FecIng = g_rst_Princi!SOLMAE_FECSOL
    moddat_g_int_CodIns = g_rst_Princi!SOLMAE_CODINS
    moddat_g_int_TipEva = g_rst_Princi!SOLMAE_TIPEVA
    moddat_g_dbl_TasInt = g_rst_Princi!SOLMAE_TASINT
    moddat_g_int_TipMon = g_rst_Princi!SOLMAE_TIPMON
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   Call gs_UbiIniGrid(grd_Listad(3))
End Sub

Private Sub fs_DatTas()
   Call gs_LimpiaGrid(grd_Listad(4))
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM TRA_EVATAS "
   g_str_Parame = g_str_Parame & " WHERE EVATAS_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(4).Text = "Empresa Peritaje"
      
      grd_Listad(4).Col = 1
      grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(4).Text = moddat_gf_Consulta_ParDes("507", g_rst_Princi!EVATAS_CODEMP)
      
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(4).Text = "Nombre Perito"
      
      grd_Listad(4).Col = 1
      grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(4).Text = Trim(g_rst_Princi!EVATAS_NOMPER & "")
      
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(4).Text = "Código REPEV SBS"
      
      grd_Listad(4).Col = 1
      grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(4).Text = Trim(g_rst_Princi!EVATAS_CODPER & "")
      
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(4).Text = "Nro. de Informe"
      
      grd_Listad(4).Col = 1
      grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(4).Text = Trim(g_rst_Princi!EVATAS_NUMINF & "")
      
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(4).Text = "Fecha Evaluación"
      
      grd_Listad(4).Col = 1
      grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(4).Text = gf_FormatoFecha(CStr(g_rst_Princi!EVATAS_FECEVA))
      
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(4).Text = "Año de Construcción"
      
      grd_Listad(4).Col = 1
      grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(4).Text = CStr(g_rst_Princi!EVATAS_ANOCON)
      
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(4).Text = "Nro. de Pisos"
      
      grd_Listad(4).Col = 1
      grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(4).Text = CStr(g_rst_Princi!EVATAS_NUMPIS)
      
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(4).Text = "Nro. de Sótanos"
      
      grd_Listad(4).Col = 1
      grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(4).Text = CStr(g_rst_Princi!EVATAS_NUMSOT)
      
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(4).Text = "Tipo de Inmueble"
      
      grd_Listad(4).Col = 1
      grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(4).Text = moddat_gf_Consulta_ParDes("221", CStr(g_rst_Princi!EVATAS_TIPINM))
      
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(4).Text = "Uso de Inmueble"
      
      grd_Listad(4).Col = 1
      grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(4).Text = moddat_gf_Consulta_ParDes("222", CStr(g_rst_Princi!EVATAS_USOINM))
      
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(4).Text = "Material de Construcción"
      
      grd_Listad(4).Col = 1
      grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(4).Text = moddat_gf_Consulta_ParDes("223", CStr(g_rst_Princi!EVATAS_MATCON))
      
      'Total
      grd_Listad(4).Rows = grd_Listad(4).Rows + 2
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(4).Text = "Area Terreno (Total)"
      
      grd_Listad(4).Col = 1
      grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(4).CellFontName = "Lucida Console"
      grd_Listad(4).CellFontSize = 8
      grd_Listad(4).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARETER_INM + g_rst_Princi!EVATAS_ARETER_ES1 + g_rst_Princi!EVATAS_ARETER_ES2 + g_rst_Princi!EVATAS_ARETER_DEP, 12, 2) & " m2"
      
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(4).Text = "Area Construida (Total)"
      
      grd_Listad(4).Col = 1
      grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(4).CellFontName = "Lucida Console"
      grd_Listad(4).CellFontSize = 8
      grd_Listad(4).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARECON_INM + g_rst_Princi!EVATAS_ARECON_ES1 + g_rst_Princi!EVATAS_ARECON_ES2 + g_rst_Princi!EVATAS_ARECON_DEP, 12, 2) & " m2"
      
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(4).Text = "Suma Asegurada (Total)"
      
      grd_Listad(4).Col = 1
      grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(4).CellFontName = "Lucida Console"
      grd_Listad(4).CellFontSize = 8
      grd_Listad(4).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_SUMASE_INM + g_rst_Princi!EVATAS_SUMASE_ES1 + g_rst_Princi!EVATAS_SUMASE_ES2 + g_rst_Princi!EVATAS_SUMASE_DEP, 12, 2)
      
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(4).Text = "Valor Comercial (Total)"
      
      grd_Listad(4).Col = 1
      grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(4).CellFontName = "Lucida Console"
      grd_Listad(4).CellFontSize = 8
      grd_Listad(4).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALCOM_INM + g_rst_Princi!EVATAS_VALCOM_ES1 + g_rst_Princi!EVATAS_VALCOM_ES2 + g_rst_Princi!EVATAS_VALCOM_DEP, 12, 2)
      
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(4).Text = "Valor Realización (Total)"
      
      grd_Listad(4).Col = 1
      grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(4).CellFontName = "Lucida Console"
      grd_Listad(4).CellFontSize = 8
      grd_Listad(4).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALREA_INM + g_rst_Princi!EVATAS_VALREA_ES1 + g_rst_Princi!EVATAS_VALREA_ES2 + g_rst_Princi!EVATAS_VALREA_DEP, 12, 2)
      
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(4).Text = "Valor Terreno (Total)"
      
      grd_Listad(4).Col = 1
      grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(4).CellFontName = "Lucida Console"
      grd_Listad(4).CellFontSize = 8
      grd_Listad(4).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALTER_INM + g_rst_Princi!EVATAS_VALTER_ES1 + g_rst_Princi!EVATAS_VALTER_ES2 + g_rst_Princi!EVATAS_VALTER_DEP, 12, 2)
      
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(4).Text = "Valor Edificación (Total)"
      
      grd_Listad(4).Col = 1
      grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(4).CellFontName = "Lucida Console"
      grd_Listad(4).CellFontSize = 8
      grd_Listad(4).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALEDI_INM + g_rst_Princi!EVATAS_VALEDI_ES1 + g_rst_Princi!EVATAS_VALEDI_ES2 + g_rst_Princi!EVATAS_VALEDI_DEP, 12, 2)
   
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(4).Text = "Valor Areas Comunes (Total)"
      
      grd_Listad(4).Col = 1
      grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(4).CellFontName = "Lucida Console"
      grd_Listad(4).CellFontSize = 8
      grd_Listad(4).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALACO_INM + g_rst_Princi!EVATAS_VALACO_ES1 + g_rst_Princi!EVATAS_VALACO_ES2 + g_rst_Princi!EVATAS_VALACO_DEP, 12, 2)
   
      'Inmueble
      grd_Listad(4).Rows = grd_Listad(4).Rows + 2
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(4).Text = "Area Terreno (Inmueble)"
      
      grd_Listad(4).Col = 1
      grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(4).CellFontName = "Lucida Console"
      grd_Listad(4).CellFontSize = 8
      grd_Listad(4).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARETER_INM, 12, 2) & " m2"
      
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(4).Text = "Area Construida (Inmueble)"
      
      grd_Listad(4).Col = 1
      grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(4).CellFontName = "Lucida Console"
      grd_Listad(4).CellFontSize = 8
      grd_Listad(4).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARECON_INM, 12, 2) & " m2"
      
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(4).Text = "Suma Asegurada (Inmueble)"
      
      grd_Listad(4).Col = 1
      grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(4).CellFontName = "Lucida Console"
      grd_Listad(4).CellFontSize = 8
      grd_Listad(4).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_SUMASE_INM, 12, 2)
      
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(4).Text = "Valor Comercial (Inmueble)"
      
      grd_Listad(4).Col = 1
      grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(4).CellFontName = "Lucida Console"
      grd_Listad(4).CellFontSize = 8
      grd_Listad(4).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALCOM_INM, 12, 2)
      
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(4).Text = "Valor Realización (Inmueble)"
      
      grd_Listad(4).Col = 1
      grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(4).CellFontName = "Lucida Console"
      grd_Listad(4).CellFontSize = 8
      grd_Listad(4).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALREA_INM, 12, 2)
      
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(4).Text = "Valor Terreno (Inmueble)"
      
      grd_Listad(4).Col = 1
      grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(4).CellFontName = "Lucida Console"
      grd_Listad(4).CellFontSize = 8
      grd_Listad(4).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALTER_INM, 12, 2)
      
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(4).Text = "Valor Edificación (Inmueble)"
      
      grd_Listad(4).Col = 1
      grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(4).CellFontName = "Lucida Console"
      grd_Listad(4).CellFontSize = 8
      grd_Listad(4).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALEDI_INM, 12, 2)
   
      grd_Listad(4).Rows = grd_Listad(4).Rows + 1
      grd_Listad(4).Row = grd_Listad(4).Rows - 1
      grd_Listad(4).Col = 0
      grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(4).Text = "Valor Areas Comunes (Inmueble)"
      
      grd_Listad(4).Col = 1
      grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(4).CellFontName = "Lucida Console"
      grd_Listad(4).CellFontSize = 8
      grd_Listad(4).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALACO_INM, 12, 2)
   
      'Estacionamiento 1
      If g_rst_Princi!EVATAS_FLGEST_ES1 = 1 Then
         grd_Listad(4).Rows = grd_Listad(4).Rows + 2
         grd_Listad(4).Row = grd_Listad(4).Rows - 1
         grd_Listad(4).Col = 0
         grd_Listad(4).CellForeColor = modgen_g_con_ColRoj
         grd_Listad(4).Text = "Area Terreno (Estac. 1)"
         
         grd_Listad(4).Col = 1
         grd_Listad(4).CellForeColor = modgen_g_con_ColRoj
         grd_Listad(4).CellFontName = "Lucida Console"
         grd_Listad(4).CellFontSize = 8
         grd_Listad(4).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARETER_ES1, 12, 2) & " m2"
         
         grd_Listad(4).Rows = grd_Listad(4).Rows + 1
         grd_Listad(4).Row = grd_Listad(4).Rows - 1
         grd_Listad(4).Col = 0
         grd_Listad(4).CellForeColor = modgen_g_con_ColRoj
         grd_Listad(4).Text = "Area Construida (Estac. 1)"
         
         grd_Listad(4).Col = 1
         grd_Listad(4).CellForeColor = modgen_g_con_ColRoj
         grd_Listad(4).CellFontName = "Lucida Console"
         grd_Listad(4).CellFontSize = 8
         grd_Listad(4).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARECON_ES1, 12, 2) & " m2"
         
         grd_Listad(4).Rows = grd_Listad(4).Rows + 1
         grd_Listad(4).Row = grd_Listad(4).Rows - 1
         grd_Listad(4).Col = 0
         grd_Listad(4).CellForeColor = modgen_g_con_ColRoj
         grd_Listad(4).Text = "Suma Asegurada (Estac. 1)"
         
         grd_Listad(4).Col = 1
         grd_Listad(4).CellForeColor = modgen_g_con_ColRoj
         grd_Listad(4).CellFontName = "Lucida Console"
         grd_Listad(4).CellFontSize = 8
         grd_Listad(4).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_SUMASE_ES1, 12, 2)
         
         grd_Listad(4).Rows = grd_Listad(4).Rows + 1
         grd_Listad(4).Row = grd_Listad(4).Rows - 1
         grd_Listad(4).Col = 0
         grd_Listad(4).CellForeColor = modgen_g_con_ColRoj
         grd_Listad(4).Text = "Valor Comercial (Estac. 1)"
         
         grd_Listad(4).Col = 1
         grd_Listad(4).CellForeColor = modgen_g_con_ColRoj
         grd_Listad(4).CellFontName = "Lucida Console"
         grd_Listad(4).CellFontSize = 8
         grd_Listad(4).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALCOM_ES1, 12, 2)
         
         grd_Listad(4).Rows = grd_Listad(4).Rows + 1
         grd_Listad(4).Row = grd_Listad(4).Rows - 1
         grd_Listad(4).Col = 0
         grd_Listad(4).CellForeColor = modgen_g_con_ColRoj
         grd_Listad(4).Text = "Valor Realización (Estac. 1)"
         
         grd_Listad(4).Col = 1
         grd_Listad(4).CellForeColor = modgen_g_con_ColRoj
         grd_Listad(4).CellFontName = "Lucida Console"
         grd_Listad(4).CellFontSize = 8
         grd_Listad(4).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALREA_ES1, 12, 2)
         
         grd_Listad(4).Rows = grd_Listad(4).Rows + 1
         grd_Listad(4).Row = grd_Listad(4).Rows - 1
         grd_Listad(4).Col = 0
         grd_Listad(4).CellForeColor = modgen_g_con_ColRoj
         grd_Listad(4).Text = "Valor Terreno (Estac. 1)"
         
         grd_Listad(4).Col = 1
         grd_Listad(4).CellForeColor = modgen_g_con_ColRoj
         grd_Listad(4).CellFontName = "Lucida Console"
         grd_Listad(4).CellFontSize = 8
         grd_Listad(4).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALTER_ES1, 12, 2)
         
         grd_Listad(4).Rows = grd_Listad(4).Rows + 1
         grd_Listad(4).Row = grd_Listad(4).Rows - 1
         grd_Listad(4).Col = 0
         grd_Listad(4).CellForeColor = modgen_g_con_ColRoj
         grd_Listad(4).Text = "Valor Edificación (Estac. 1)"
         
         grd_Listad(4).Col = 1
         grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(4).CellFontName = "Lucida Console"
         grd_Listad(4).CellFontSize = 8
         grd_Listad(4).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALEDI_ES1, 12, 2)
      
         grd_Listad(4).Rows = grd_Listad(4).Rows + 1
         grd_Listad(4).Row = grd_Listad(4).Rows - 1
         grd_Listad(4).Col = 0
         grd_Listad(4).CellForeColor = modgen_g_con_ColRoj
         grd_Listad(4).Text = "Valor Areas Comunes (Estac. 1)"
         
         grd_Listad(4).Col = 1
         grd_Listad(4).CellForeColor = modgen_g_con_ColRoj
         grd_Listad(4).CellFontName = "Lucida Console"
         grd_Listad(4).CellFontSize = 8
         grd_Listad(4).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALACO_ES1, 12, 2)
      End If
   
      If g_rst_Princi!EVATAS_FLGEST_ES2 = 1 Then
         grd_Listad(4).Rows = grd_Listad(4).Rows + 2
         grd_Listad(4).Row = grd_Listad(4).Rows - 1
         grd_Listad(4).Col = 0
         grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(4).Text = "Area Terreno (Estac. 2)"
         
         grd_Listad(4).Col = 1
         grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(4).CellFontName = "Lucida Console"
         grd_Listad(4).CellFontSize = 8
         grd_Listad(4).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARETER_ES2, 12, 2)
         
         grd_Listad(4).Rows = grd_Listad(4).Rows + 1
         grd_Listad(4).Row = grd_Listad(4).Rows - 1
         grd_Listad(4).Col = 0
         grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(4).Text = "Area Construida (Estac. 2)"
         
         grd_Listad(4).Col = 1
         grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(4).CellFontName = "Lucida Console"
         grd_Listad(4).CellFontSize = 8
         grd_Listad(4).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARECON_ES2, 12, 2)
         
         grd_Listad(4).Rows = grd_Listad(4).Rows + 1
         grd_Listad(4).Row = grd_Listad(4).Rows - 1
         grd_Listad(4).Col = 0
         grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(4).Text = "Suma Asegurada (Estac. 2)"
         
         grd_Listad(4).Col = 1
         grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(4).CellFontName = "Lucida Console"
         grd_Listad(4).CellFontSize = 8
         grd_Listad(4).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_SUMASE_ES2, 12, 2)
         
         grd_Listad(4).Rows = grd_Listad(4).Rows + 1
         grd_Listad(4).Row = grd_Listad(4).Rows - 1
         grd_Listad(4).Col = 0
         grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(4).Text = "Valor Comercial (Estac. 2)"
         
         grd_Listad(4).Col = 1
         grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(4).CellFontName = "Lucida Console"
         grd_Listad(4).CellFontSize = 8
         grd_Listad(4).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALCOM_ES2, 12, 2)
         
         grd_Listad(4).Rows = grd_Listad(4).Rows + 1
         grd_Listad(4).Row = grd_Listad(4).Rows - 1
         grd_Listad(4).Col = 0
         grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(4).Text = "Valor Realización (Estac. 2)"
         
         grd_Listad(4).Col = 1
         grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(4).CellFontName = "Lucida Console"
         grd_Listad(4).CellFontSize = 8
         grd_Listad(4).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALREA_ES2, 12, 2)
         
         grd_Listad(4).Rows = grd_Listad(4).Rows + 1
         grd_Listad(4).Row = grd_Listad(4).Rows - 1
         grd_Listad(4).Col = 0
         grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(4).Text = "Valor Terreno (Estac. 2)"
         
         grd_Listad(4).Col = 1
         grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(4).CellFontName = "Lucida Console"
         grd_Listad(4).CellFontSize = 8
         grd_Listad(4).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALTER_ES2, 12, 2)
         
         grd_Listad(4).Rows = grd_Listad(4).Rows + 1
         grd_Listad(4).Row = grd_Listad(4).Rows - 1
         grd_Listad(4).Col = 0
         grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(4).Text = "Valor Edificación (Estac. 2)"
         
         grd_Listad(4).Col = 1
         grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(4).CellFontName = "Lucida Console"
         grd_Listad(4).CellFontSize = 8
         grd_Listad(4).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALEDI_ES2, 12, 2)
      
         grd_Listad(4).Rows = grd_Listad(4).Rows + 1
         grd_Listad(4).Row = grd_Listad(4).Rows - 1
         grd_Listad(4).Col = 0
         grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(4).Text = "Valor Areas Comunes (Estac. 2)"
         
         grd_Listad(4).Col = 1
         grd_Listad(4).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(4).CellFontName = "Lucida Console"
         grd_Listad(4).CellFontSize = 8
         grd_Listad(4).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALACO_ES2, 12, 2)
      End If
   
      If g_rst_Princi!EVATAS_FLGEST_DEP = 1 Then
         grd_Listad(4).Rows = grd_Listad(4).Rows + 2
         grd_Listad(4).Row = grd_Listad(4).Rows - 1
         grd_Listad(4).Col = 0
         grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(4).Text = "Area Terreno (Depósito)"
         
         grd_Listad(4).Col = 1
         grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(4).CellFontName = "Lucida Console"
         grd_Listad(4).CellFontSize = 8
         grd_Listad(4).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARETER_DEP, 12, 2)
         
         grd_Listad(4).Rows = grd_Listad(4).Rows + 1
         grd_Listad(4).Row = grd_Listad(4).Rows - 1
         grd_Listad(4).Col = 0
         grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(4).Text = "Area Construida (Depósito)"
         
         grd_Listad(4).Col = 1
         grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(4).CellFontName = "Lucida Console"
         grd_Listad(4).CellFontSize = 8
         grd_Listad(4).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARECON_DEP, 12, 2)
         
         grd_Listad(4).Rows = grd_Listad(4).Rows + 1
         grd_Listad(4).Row = grd_Listad(4).Rows - 1
         grd_Listad(4).Col = 0
         grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(4).Text = "Suma Asegurada (Depósito)"
         
         grd_Listad(4).Col = 1
         grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(4).CellFontName = "Lucida Console"
         grd_Listad(4).CellFontSize = 8
         grd_Listad(4).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_SUMASE_DEP, 12, 2)
         
         grd_Listad(4).Rows = grd_Listad(4).Rows + 1
         grd_Listad(4).Row = grd_Listad(4).Rows - 1
         grd_Listad(4).Col = 0
         grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(4).Text = "Valor Comercial (Depósito)"
         
         grd_Listad(4).Col = 1
         grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(4).CellFontName = "Lucida Console"
         grd_Listad(4).CellFontSize = 8
         grd_Listad(4).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALCOM_DEP, 12, 2)
         
         grd_Listad(4).Rows = grd_Listad(4).Rows + 1
         grd_Listad(4).Row = grd_Listad(4).Rows - 1
         grd_Listad(4).Col = 0
         grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(4).Text = "Valor Realización (Depósito)"
         
         grd_Listad(4).Col = 1
         grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(4).CellFontName = "Lucida Console"
         grd_Listad(4).CellFontSize = 8
         grd_Listad(4).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALREA_DEP, 12, 2)
         
         grd_Listad(4).Rows = grd_Listad(4).Rows + 1
         grd_Listad(4).Row = grd_Listad(4).Rows - 1
         grd_Listad(4).Col = 0
         grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(4).Text = "Valor Terreno (Depósito)"
         
         grd_Listad(4).Col = 1
         grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(4).CellFontName = "Lucida Console"
         grd_Listad(4).CellFontSize = 8
         grd_Listad(4).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALTER_DEP, 12, 2)
         
         grd_Listad(4).Rows = grd_Listad(4).Rows + 1
         grd_Listad(4).Row = grd_Listad(4).Rows - 1
         grd_Listad(4).Col = 0
         grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(4).Text = "Valor Edificación (Depósito)"
         
         grd_Listad(4).Col = 1
         grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(4).CellFontName = "Lucida Console"
         grd_Listad(4).CellFontSize = 8
         grd_Listad(4).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALEDI_DEP, 12, 2)
      
         grd_Listad(4).Rows = grd_Listad(4).Rows + 1
         grd_Listad(4).Row = grd_Listad(4).Rows - 1
         grd_Listad(4).Col = 0
         grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(4).Text = "Valor Areas Comunes (Depósito)"
         
         grd_Listad(4).Col = 1
         grd_Listad(4).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(4).CellFontName = "Lucida Console"
         grd_Listad(4).CellFontSize = 8
         grd_Listad(4).Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!EVATAS_TIPMON)) & " " & gf_FormatoNumero(g_rst_Princi!EVATAS_VALACO_DEP, 12, 2)
      End If
      
      Call gs_UbiIniGrid(grd_Listad(4))
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub
Private Sub fs_DatSeg()
      
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT POLIZA_NUMVIV "
   g_str_Parame = g_str_Parame & "  FROM TRA_POLIZA "
   g_str_Parame = g_str_Parame & " WHERE POLIZA_NUMSOL = '" & moddat_g_str_NumSol & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      g_rst_Genera.MoveFirst
      If Len(Trim(g_rst_Genera!POLIZA_NUMVIV & "")) = 0 Then
         l_int_PolReg = 1
      End If
   End If
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub
Private Sub fs_DatSeg_old()
   Call gs_LimpiaGrid(grd_Listad(5))
   
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
      
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Empresa de Seguros"
      
      grd_Listad(5).Col = 1
      grd_Listad(5).Text = moddat_gf_Consulta_ComSeg(g_rst_Princi!EVASEG_ESGDES & "")
   
      grd_Listad(5).Rows = grd_Listad(5).Rows + 2
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Tipo de Seguro Desgravamen"

      grd_Listad(5).Col = 1
      grd_Listad(5).Text = moddat_gf_Consulta_TipSeg(g_rst_Princi!EVASEG_ESGDES, g_rst_Princi!EVASEG_TIPSEG)
      
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Fecha Evaluación (Seg. Desgravamen)"
      
      grd_Listad(5).Col = 1
      grd_Listad(5).Text = gf_FormatoFecha(CStr(g_rst_Princi!EVASEG_EVADES))
      
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Tipo de Valor (Seg. Desgravamen)"
      
      grd_Listad(5).Col = 1
      grd_Listad(5).Text = moddat_gf_Consulta_ParDes("227", CStr(g_rst_Princi!EVASEG_TIPDES))
      
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Valor a Aplicar"
      
      grd_Listad(5).Col = 1
      grd_Listad(5).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(5).Text = Format(g_rst_Princi!EVASEG_FOIDES, "###,###,##0.000000")
      
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(5).Text = "Fecha Emisión Póliza"
      
      grd_Listad(5).Col = 1
      grd_Listad(5).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(5).Text = gf_FormatoFecha(CStr(g_rst_Genera!POLIZA_FEMDES))
      
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Número de Póliza"
      
      grd_Listad(5).Col = 1
      grd_Listad(5).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(5).Text = Trim(g_rst_Genera!POLIZA_NUMDES & "") & IIf(Len(Trim(g_rst_Genera!POLIZA_NUMCYG & "")) > 0, " / " & Trim(g_rst_Genera!POLIZA_NUMCYG & ""), "")
      
      grd_Listad(5).Rows = grd_Listad(5).Rows + 2
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Fecha Evaluación (Seg. Inmueble)"
      
      grd_Listad(5).Col = 1
      grd_Listad(5).Text = gf_FormatoFecha(CStr(g_rst_Princi!EVASEG_EVAVIV))
      
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Tipo de Valor (Seg. Inmueble)"
      
      grd_Listad(5).Col = 1
      grd_Listad(5).Text = moddat_gf_Consulta_ParDes("227", CStr(g_rst_Princi!EVASEG_TIPVIV))
      
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Valor a Aplicar"
      
      grd_Listad(5).Col = 1
      grd_Listad(5).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(5).Text = Format(g_rst_Princi!EVASEG_FOIVIV, "###,###,##0.000000")
      
      If Len(Trim(g_rst_Genera!POLIZA_NUMVIV & "")) > 0 Then
         grd_Listad(5).Rows = grd_Listad(5).Rows + 1
         grd_Listad(5).Row = grd_Listad(5).Rows - 1
         grd_Listad(5).Col = 0
         grd_Listad(5).Text = "Fecha Emisión Póliza"
         
         grd_Listad(5).Col = 1
         grd_Listad(5).Text = gf_FormatoFecha(CStr(g_rst_Genera!POLIZA_FEMVIV))
         
         grd_Listad(5).Rows = grd_Listad(5).Rows + 1
         grd_Listad(5).Row = grd_Listad(5).Rows - 1
         grd_Listad(5).Col = 0
         grd_Listad(5).Text = "Número de Póliza"
         
         grd_Listad(5).Col = 1
         grd_Listad(5).Text = Trim(g_rst_Genera!POLIZA_NUMVIV & "")
      Else
         grd_Listad(5).Rows = grd_Listad(5).Rows + 1
         grd_Listad(5).Row = grd_Listad(5).Rows - 1
         grd_Listad(5).Col = 0
         grd_Listad(5).Text = "Número de Póliza"
         
         grd_Listad(5).Col = 1
         grd_Listad(5).Text = "POLIZA NO RECIBIDA"
         
         l_int_PolReg = 1
      End If
      
      txt_ObsSeg.Text = Trim(g_rst_Princi!EVASEG_OBSERV & "")
      Call gs_UbiIniGrid(grd_Listad(5))
   End If
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_DatLeg()
   
   l_int_MonCvt = 0
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT EVALEG_MONCVT "
   g_str_Parame = g_str_Parame & "  FROM TRA_EVALEG "
   g_str_Parame = g_str_Parame & " WHERE EVALEG_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
            
      If Not IsNull(g_rst_Princi!EVALEG_MONCVT) Then
         l_int_MonCvt = g_rst_Princi!EVALEG_MONCVT
      End If
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub
Private Sub fs_DatLeg_old()
   Call gs_LimpiaGrid(grd_Listad(6))
   l_int_MonCvt = 0
   
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
      txt_ComCre.Text = "Fecha de Comité de Créditos: " & gf_FormatoFecha(CStr(g_rst_Princi!EVALEG_FECCOM)) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Trim(g_rst_Princi!EVALEG_OBSCOM & "")
      
      If g_rst_Princi!EVALEG_FECCVT > 0 Then
         grd_Listad(6).Rows = grd_Listad(6).Rows + 1
         grd_Listad(6).Row = grd_Listad(6).Rows - 1
         grd_Listad(6).Col = 0
         grd_Listad(6).Text = "Fecha Firma Contrato Compra Venta"
         
         grd_Listad(6).Col = 1
         grd_Listad(6).Text = gf_FormatoFecha(CStr(g_rst_Princi!EVALEG_FECCVT))
         
         If Not IsNull(g_rst_Princi!EVALEG_TCASBS) Then
            If g_rst_Princi!EVALEG_TCASBS > 0 Then
               grd_Listad(6).Rows = grd_Listad(6).Rows + 1
               grd_Listad(6).Row = grd_Listad(6).Rows - 1
               grd_Listad(6).Col = 0
               grd_Listad(6).Text = "Tipo de Cambio SBS"
               
               grd_Listad(6).Col = 1
               grd_Listad(6).Text = Format(g_rst_Princi!EVALEG_TCASBS, "###,##0.0000")
            End If
         End If
      
         If g_rst_Princi!EVALEG_TCACVT > 0 Then
            grd_Listad(6).Rows = grd_Listad(6).Rows + 1
            grd_Listad(6).Row = grd_Listad(6).Rows - 1
            grd_Listad(6).Col = 0
            grd_Listad(6).Text = "Tipo de Cambio aplicado"
            
            grd_Listad(6).Col = 1
            grd_Listad(6).Text = Format(g_rst_Princi!EVALEG_TCACVT, "###,##0.0000")
         End If
      End If
      
      If Not IsNull(g_rst_Princi!EVALEG_MONCVT) Then
         grd_Listad(6).Rows = grd_Listad(6).Rows + 2
         grd_Listad(6).Row = grd_Listad(6).Rows - 1
         grd_Listad(6).Col = 0
         grd_Listad(6).Text = "Moneda Compra-Venta"
         
         grd_Listad(6).Col = 1
         grd_Listad(6).Text = moddat_gf_Consulta_ParDes("204", g_rst_Princi!EVALEG_MONCVT)
         
         grd_Listad(6).Rows = grd_Listad(6).Rows + 1
         grd_Listad(6).Row = grd_Listad(6).Rows - 1
         grd_Listad(6).Col = 0
         grd_Listad(6).Text = "Valor Compra-Venta"
         
         grd_Listad(6).Col = 1
         grd_Listad(6).CellFontName = "Lucida Console"
         grd_Listad(6).CellFontSize = 8
         grd_Listad(6).Text = moddat_gf_Consulta_ParDes("229", g_rst_Princi!EVALEG_MONCVT) & " " & gf_FormatoNumero(g_rst_Princi!EVALEG_COMVTA, 12, 2)
      
         grd_Listad(6).Rows = grd_Listad(6).Rows + 1
         grd_Listad(6).Row = grd_Listad(6).Rows - 1
         grd_Listad(6).Col = 0
         grd_Listad(6).Text = "Aporte Propio"
         
         grd_Listad(6).Col = 1
         grd_Listad(6).CellFontName = "Lucida Console"
         grd_Listad(6).CellFontSize = 8
         grd_Listad(6).Text = moddat_gf_Consulta_ParDes("229", g_rst_Princi!EVALEG_MONCVT) & " " & gf_FormatoNumero(g_rst_Princi!EVALEG_APOPRO, 12, 2)
      
         grd_Listad(6).Rows = grd_Listad(6).Rows + 1
         grd_Listad(6).Row = grd_Listad(6).Rows - 1
         grd_Listad(6).Col = 0
         grd_Listad(6).Text = "Monto Préstamo"
         
         grd_Listad(6).Col = 1
         grd_Listad(6).CellFontName = "Lucida Console"
         grd_Listad(6).CellFontSize = 8
         grd_Listad(6).Text = moddat_gf_Consulta_ParDes("229", g_rst_Princi!EVALEG_MONCVT) & " " & gf_FormatoNumero(g_rst_Princi!EVALEG_MTOPRE, 12, 2)
      End If
      
      If grd_Listad(6).Rows = 0 Then
         grd_Listad(6).Rows = grd_Listad(6).Rows + 1
      Else
         grd_Listad(6).Rows = grd_Listad(6).Rows + 2
      End If
      
      grd_Listad(6).Row = grd_Listad(6).Rows - 1
      grd_Listad(6).Col = 0
      grd_Listad(6).Text = "Fecha Firma Contrato (Crédito)"
      
      grd_Listad(6).Col = 1
      grd_Listad(6).Text = gf_FormatoFecha(CStr(g_rst_Princi!EVALEG_FIRCON))
   
      grd_Listad(6).Rows = grd_Listad(6).Rows + 1
      grd_Listad(6).Row = grd_Listad(6).Rows - 1
      grd_Listad(6).Col = 0
      grd_Listad(6).Text = "Notaria"
      
      grd_Listad(6).Col = 1
      grd_Listad(6).Text = moddat_gf_Consulta_ParDes("509", g_rst_Princi!EVALEG_CODNOT & "")
   
      grd_Listad(6).Rows = grd_Listad(6).Rows + 1
      grd_Listad(6).Row = grd_Listad(6).Rows - 1
      grd_Listad(6).Col = 0
      grd_Listad(6).Text = "Representante Legal 1"
      
      grd_Listad(6).Col = 1
      grd_Listad(6).Text = moddat_gf_Consulta_ParDes("512", g_rst_Princi!EVALEG_REPLG1 & "")
   
      grd_Listad(6).Rows = grd_Listad(6).Rows + 1
      grd_Listad(6).Row = grd_Listad(6).Rows - 1
      grd_Listad(6).Col = 0
      grd_Listad(6).Text = "Representante Legal 2"
      
      grd_Listad(6).Col = 1
      grd_Listad(6).Text = moddat_gf_Consulta_ParDes("512", g_rst_Princi!EVALEG_REPLG2 & "")
      
      grd_Listad(6).Rows = grd_Listad(6).Rows + 1
      grd_Listad(6).Row = grd_Listad(6).Rows - 1
      grd_Listad(6).Col = 0
      grd_Listad(6).Text = "Monto Hipoteca "
      
      grd_Listad(6).Col = 1
      grd_Listad(6).CellFontName = "Lucida Console"
      grd_Listad(6).CellFontSize = 8
      grd_Listad(6).Text = moddat_gf_Consulta_ParDes("229", g_rst_Princi!EVALEG_MONHIP) & " " & gf_FormatoNumero(g_rst_Princi!EVALEG_MTOHIP, 12, 2)
      
      If g_rst_Princi!EVALEG_FECBLQ_INM > 0 Then
         grd_Listad(6).Rows = grd_Listad(6).Rows + 2
         grd_Listad(6).Row = grd_Listad(6).Rows - 1
         grd_Listad(6).Col = 0
         grd_Listad(6).Text = "Bloqueo Registral Inscrito"
         
         grd_Listad(6).Col = 1
         grd_Listad(6).Text = "SI"
      
         grd_Listad(6).Rows = grd_Listad(6).Rows + 1
         grd_Listad(6).Row = grd_Listad(6).Rows - 1
         grd_Listad(6).Col = 0
         grd_Listad(6).Text = "Sede Registral"
         
         grd_Listad(6).Col = 1
         grd_Listad(6).Text = moddat_gf_Consulta_ParDes("511", CStr(g_rst_Princi!EVALEG_SEDREG & ""))
         
         grd_Listad(6).Rows = grd_Listad(6).Rows + 2
         grd_Listad(6).Row = grd_Listad(6).Rows - 1
         grd_Listad(6).Col = 0
         grd_Listad(6).Text = "Fecha Bloqueo (Inmueble)"
         
         grd_Listad(6).Col = 1
         grd_Listad(6).Text = gf_FormatoFecha(CStr(g_rst_Princi!EVALEG_FECBLQ_INM))
                  
         grd_Listad(6).Rows = grd_Listad(6).Rows + 1
         grd_Listad(6).Row = grd_Listad(6).Rows - 1
         grd_Listad(6).Col = 0
         grd_Listad(6).Text = "Doc. Registral (Inmueble)"
         
         grd_Listad(6).Col = 1
         grd_Listad(6).Text = moddat_gf_Consulta_ParDes("026", g_rst_Princi!EVALEG_TIPDOC_INM)
                  
         Select Case g_rst_Princi!EVALEG_TIPDOC_INM
            Case 1: grd_Listad(6).Text = grd_Listad(6).Text & " NRO. " & Trim(g_rst_Princi!EVALEG_NUMPAR_INM & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!EVALEG_NUMAPA_INM & "")
            Case 2: grd_Listad(6).Text = grd_Listad(6).Text & " NRO. " & Trim(g_rst_Princi!EVALEG_NUMFIC_INM & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!EVALEG_NUMAFI_INM & "")
            Case 3: grd_Listad(6).Text = grd_Listad(6).Text & " (" & Trim(g_rst_Princi!EVALEG_NUMTOM_INM & "") & " / " & Trim(g_rst_Princi!EVALEG_NUMFOJ_INM & "") & " / " & Trim(g_rst_Princi!EVALEG_NUMLIB_INM & "") & ")"
         End Select
         
         If g_rst_Princi!EVALEG_FLGEST_ES1 = 1 Then
            grd_Listad(6).Rows = grd_Listad(6).Rows + 2
            grd_Listad(6).Row = grd_Listad(6).Rows - 1
            grd_Listad(6).Col = 0
            grd_Listad(6).Text = "Fecha Bloqueo (Estac. 1)"
            
            grd_Listad(6).Col = 1
            grd_Listad(6).Text = gf_FormatoFecha(CStr(g_rst_Princi!EVALEG_FECBLQ_ES1))
                       
            grd_Listad(6).Rows = grd_Listad(6).Rows + 1
            grd_Listad(6).Row = grd_Listad(6).Rows - 1
            grd_Listad(6).Col = 0
            grd_Listad(6).Text = "Doc. Registral (Estac. 1)"
            
            grd_Listad(6).Col = 1
            grd_Listad(6).Text = moddat_gf_Consulta_ParDes("026", g_rst_Princi!EVALEG_TIPDOC_ES1)
            
            Select Case g_rst_Princi!EVALEG_TIPDOC_ES1
               Case 1: grd_Listad(6).Text = grd_Listad(6).Text & " NRO. " & Trim(g_rst_Princi!EVALEG_NUMPAR_ES1 & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!EVALEG_NUMAPA_ES1 & "")
               Case 2: grd_Listad(6).Text = grd_Listad(6).Text & " NRO. " & Trim(g_rst_Princi!EVALEG_NUMFIC_ES1 & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!EVALEG_NUMAFI_ES1 & "")
               Case 3: grd_Listad(6).Text = grd_Listad(6).Text & " (" & Trim(g_rst_Princi!EVALEG_NUMTOM_ES1 & "") & " / " & Trim(g_rst_Princi!EVALEG_NUMFOJ_ES1 & "") & " / " & Trim(g_rst_Princi!EVALEG_NUMLIB_ES1 & "") & ")"
            End Select
         End If
         
         If g_rst_Princi!EVALEG_FLGEST_ES2 = 1 Then
            grd_Listad(6).Rows = grd_Listad(6).Rows + 2
            grd_Listad(6).Row = grd_Listad(6).Rows - 1
            grd_Listad(6).Col = 0
            grd_Listad(6).Text = "Fecha Bloqueo (Estac. 2)"
            
            grd_Listad(6).Col = 1
            grd_Listad(6).Text = gf_FormatoFecha(CStr(g_rst_Princi!EVALEG_FECBLQ_ES2))
                        
            grd_Listad(6).Rows = grd_Listad(6).Rows + 1
            grd_Listad(6).Row = grd_Listad(6).Rows - 1
            grd_Listad(6).Col = 0
            grd_Listad(6).Text = "Doc. Registral (Estac. 2)"
            
            grd_Listad(6).Col = 1
            grd_Listad(6).Text = moddat_gf_Consulta_ParDes("026", g_rst_Princi!EVALEG_TIPDOC_ES2)
            
            Select Case g_rst_Princi!EVALEG_TIPDOC_ES2
               Case 1: grd_Listad(6).Text = grd_Listad(6).Text & " NRO. " & Trim(g_rst_Princi!EVALEG_NUMPAR_ES2 & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!EVALEG_NUMAPA_ES2 & "")
               Case 2: grd_Listad(6).Text = grd_Listad(6).Text & " NRO. " & Trim(g_rst_Princi!EVALEG_NUMFIC_ES2 & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!EVALEG_NUMAFI_ES2 & "")
               Case 3: grd_Listad(6).Text = grd_Listad(6).Text & " (" & Trim(g_rst_Princi!EVALEG_NUMTOM_ES2 & "") & " / " & Trim(g_rst_Princi!EVALEG_NUMFOJ_ES2 & "") & " / " & Trim(g_rst_Princi!EVALEG_NUMLIB_ES2 & "") & ")"
            End Select
         End If
         
         If g_rst_Princi!EVALEG_FLGEST_DEP = 1 Then
            grd_Listad(6).Rows = grd_Listad(6).Rows + 2
            grd_Listad(6).Row = grd_Listad(6).Rows - 1
            grd_Listad(6).Col = 0
            grd_Listad(6).Text = "Fecha Bloqueo (Depósito)"
            
            grd_Listad(6).Col = 1
            grd_Listad(6).Text = gf_FormatoFecha(CStr(g_rst_Princi!EVALEG_FECBLQ_DEP))
                        
            grd_Listad(6).Rows = grd_Listad(6).Rows + 1
            grd_Listad(6).Row = grd_Listad(6).Rows - 1
            grd_Listad(6).Col = 0
            grd_Listad(6).Text = "Doc. Registral (Depósito)"
            
            grd_Listad(6).Col = 1
            grd_Listad(6).Text = moddat_gf_Consulta_ParDes("026", g_rst_Princi!EVALEG_TIPDOC_DEP)
                        
            Select Case g_rst_Princi!EVALEG_TIPDOC_DEP
               Case 1: grd_Listad(6).Text = grd_Listad(6).Text & " NRO. " & Trim(g_rst_Princi!EVALEG_NUMPAR_DEP & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!EVALEG_NUMAPA_DEP & "")
               Case 2: grd_Listad(6).Text = grd_Listad(6).Text & " NRO. " & Trim(g_rst_Princi!EVALEG_NUMFIC_DEP & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!EVALEG_NUMAFI_DEP & "")
               Case 3: grd_Listad(6).Text = grd_Listad(6).Text & " (" & Trim(g_rst_Princi!EVALEG_NUMTOM_DEP & "") & " / " & Trim(g_rst_Princi!EVALEG_NUMFOJ_DEP & "") & " / " & Trim(g_rst_Princi!EVALEG_NUMLIB_DEP & "") & ")"
            End Select
         End If
      End If
      
      If Not IsNull(g_rst_Princi!EVALEG_MONCVT) Then
         l_int_MonCvt = g_rst_Princi!EVALEG_MONCVT
      End If
      
      Call gs_UbiIniGrid(grd_Listad(6))
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub
Private Sub fs_DatMVi()
   If InStr(moddat_g_str_AgrCRC, moddat_g_str_CodPrd) = 0 Then 'moddat_g_str_CodPrd <> "001" Then
      Exit Sub
   End If

   txt_ObsMVi.Text = ""
   Call gs_LimpiaGrid(grd_Listad(7))
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM TRA_EVAMVI "
   g_str_Parame = g_str_Parame & " WHERE EVAMVI_NUMSOL = '" & moddat_g_str_NumSol & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      grd_Listad(7).Rows = grd_Listad(7).Rows + 1
      grd_Listad(7).Row = grd_Listad(7).Rows - 1
      grd_Listad(7).Col = 0
      grd_Listad(7).Text = "Fecha Envío"
      
      grd_Listad(7).Col = 1
      grd_Listad(7).Text = gf_FormatoFecha(CStr(g_rst_Princi!EVAMVI_FECENV))
   
      If g_rst_Princi!EVAMVI_FECREC > 0 Then
         grd_Listad(7).Rows = grd_Listad(7).Rows + 1
         grd_Listad(7).Row = grd_Listad(7).Rows - 1
         grd_Listad(7).Col = 0
         grd_Listad(7).Text = "Fecha de Recepción"
   
         grd_Listad(7).Col = 1
         grd_Listad(7).Text = gf_FormatoFecha(CStr(g_rst_Princi!EVAMVI_FECREC))
      
         grd_Listad(7).Rows = grd_Listad(7).Rows + 1
         grd_Listad(7).Row = grd_Listad(7).Rows - 1
         grd_Listad(7).Col = 0
         grd_Listad(7).Text = "Nro. Expediente Mivivienda"
   
         grd_Listad(7).Col = 1
         grd_Listad(7).Text = Trim(g_rst_Princi!EVAMVI_CODMVI & "")
      
         txt_ObsMVi.Text = Trim(g_rst_Princi!EVAMVI_OBSERV & "")
      End If
      
      Call gs_UbiIniGrid(grd_Listad(7))
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_DatCof()
   If Not (InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_AgrTFMV, moddat_g_str_CodPrd) > 0) Then  'moddat_g_str_CodPrd = "004" Or moddat_g_str_CodPrd = "003" Or moddat_g_str_CodPrd = "007" Or moddat_g_str_CodPrd = "009" Or moddat_g_str_CodPrd = "010" Or moddat_g_str_CodPrd = "013" Or moddat_g_str_CodPrd = "014" Or moddat_g_str_CodPrd = "015" Or moddat_g_str_CodPrd = "016" Or moddat_g_str_CodPrd = "017" Or moddat_g_str_CodPrd = "018" Or moddat_g_str_CodPrd = "019" Or moddat_g_str_CodPrd = "020" Or moddat_g_str_CodPrd = "021" Or moddat_g_str_CodPrd = "022" Or moddat_g_str_CodPrd = "023") Then
      Exit Sub
   End If
   
   txt_ObsMVi.Text = ""
   Call gs_LimpiaGrid(grd_Listad(7))
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM TRA_EVACOF "
   g_str_Parame = g_str_Parame & " WHERE EVACOF_NUMSOL = '" & moddat_g_str_NumSol & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      grd_Listad(7).Rows = grd_Listad(7).Rows + 1
      grd_Listad(7).Row = grd_Listad(7).Rows - 1
      grd_Listad(7).Col = 0
      grd_Listad(7).Text = "Fecha Envío"
      
      grd_Listad(7).Col = 1
      grd_Listad(7).Text = gf_FormatoFecha(CStr(g_rst_Princi!EVACOF_FECENV))
   
      If g_rst_Princi!EVACOF_FECREC > 0 Then
         If InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Then 'moddat_g_str_CodPrd = "003" Then
            grd_Listad(7).Rows = grd_Listad(7).Rows + 1
            grd_Listad(7).Row = grd_Listad(7).Rows - 1
            grd_Listad(7).Col = 0
            grd_Listad(7).Text = "Nro. Operación Mivivienda"
            
            grd_Listad(7).Col = 1
            grd_Listad(7).Text = Trim(g_rst_Princi!EVACOF_CODMV1 & "")
            
            grd_Listad(7).Rows = grd_Listad(7).Rows + 1
            grd_Listad(7).Row = grd_Listad(7).Rows - 1
            grd_Listad(7).Col = 0
            grd_Listad(7).Text = "Fecha Aprobación Mivivienda"
            
            grd_Listad(7).Col = 1
            grd_Listad(7).Text = gf_FormatoFecha(CStr(g_rst_Princi!EVACOF_APRMVI))
         End If
         
         grd_Listad(7).Rows = grd_Listad(7).Rows + 1
         grd_Listad(7).Row = grd_Listad(7).Rows - 1
         grd_Listad(7).Col = 0
         grd_Listad(7).Text = "Nro. Carta COFIDE"
         
         grd_Listad(7).Col = 1
         grd_Listad(7).Text = Trim(g_rst_Princi!EVACOF_NUMCAR & "")
         
         grd_Listad(7).Rows = grd_Listad(7).Rows + 1
         grd_Listad(7).Row = grd_Listad(7).Rows - 1
         grd_Listad(7).Col = 0
         grd_Listad(7).Text = "Fecha Recepción Carta COFIDE"
         
         grd_Listad(7).Col = 1
         grd_Listad(7).Text = gf_FormatoFecha(CStr(g_rst_Princi!EVACOF_FECREC))
         
         grd_Listad(7).Rows = grd_Listad(7).Rows + 1
         grd_Listad(7).Row = grd_Listad(7).Rows - 1
         grd_Listad(7).Col = 0
         grd_Listad(7).Text = "Nro. Operación COFIDE"
         
         grd_Listad(7).Col = 1
         grd_Listad(7).Text = Trim(g_rst_Princi!EVACOF_CODMVI & "")
         
         grd_Listad(7).Rows = grd_Listad(7).Rows + 1
         grd_Listad(7).Row = grd_Listad(7).Rows - 1
         grd_Listad(7).Col = 0
         grd_Listad(7).Text = "Fecha Desembolso COFIDE"
         
         grd_Listad(7).Col = 1
         grd_Listad(7).Text = gf_FormatoFecha(CStr(g_rst_Princi!EVACOF_FECDES))
         
         grd_Listad(7).Rows = grd_Listad(7).Rows + 1
         grd_Listad(7).Row = grd_Listad(7).Rows - 1
         grd_Listad(7).Col = 0
         grd_Listad(7).Text = "Importe Desembolsado"
         
         grd_Listad(7).Col = 1
         grd_Listad(7).CellFontName = "Lucida Console"
         grd_Listad(7).CellFontSize = 8
         grd_Listad(7).Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!EVACOF_MTODES, 12, 2)
         
         txt_ObsMVi.Text = Trim(g_rst_Princi!EVACOF_OBSERV & "")
      End If
      
      Call gs_UbiIniGrid(grd_Listad(7))
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_GasAdm()
   'Inicializando Variables para Hoja Resumen
   l_dbl_ImpTas = 0
   l_dbl_ImpNot = 0
   l_dbl_ImpEst = 0
   l_dbl_ImpEva = 0
   l_dbl_ImpAdm = 0
   l_dbl_ImpRed = 0
   l_dbl_ImpBlq = 0
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM TRA_GASADM "
   g_str_Parame = g_str_Parame & " WHERE GASADM_NUMSOL = '" & moddat_g_str_NumSol & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         Select Case g_rst_Princi!GASADM_CODGAS
            Case 11: l_dbl_ImpTas = g_rst_Princi!GASADM_IMPORT
            Case 12: l_dbl_ImpNot = g_rst_Princi!GASADM_IMPORT
            Case 14: l_dbl_ImpEst = g_rst_Princi!GASADM_IMPORT
            Case 15, 22: l_dbl_ImpEva = g_rst_Princi!GASADM_IMPORT
            Case 16: l_dbl_ImpBlq = g_rst_Princi!GASADM_IMPORT
            Case 20: l_dbl_ImpAdm = g_rst_Princi!GASADM_IMPORT
            Case 21: l_dbl_ImpRed = g_rst_Princi!GASADM_IMPORT
         End Select
         
         g_rst_Princi.MoveNext
         DoEvents
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Genera_Archivos(ByVal p_NomFil As String)
Dim r_int_Contad     As Integer
Dim r_int_NumFil     As Integer
Dim r_str_NomCab     As String
Dim r_str_NomDet     As String
Dim r_dbl_PorNCo     As Double
Dim r_dbl_PorCon     As Double
Dim r_dbl_TotPre     As Double
Dim r_dbl_TotCuo     As Double
Dim r_str_OpeMVi     As String
Dim r_int_PerGra     As Integer
Dim r_dbl_SalCap     As Double
Dim r_dbl_MtoPre     As Double
Dim r_dbl_MtoGra     As Double
   
   For r_int_Contad = Len(p_NomFil) To 1 Step -1
      If Mid(p_NomFil, r_int_Contad, 1) = "\" Then
         Exit For
      End If
   Next r_int_Contad

   r_str_NomCab = Mid(p_NomFil, 1, r_int_Contad) & "C" & Format(date, "yymmdd") & ".064"
   r_str_NomDet = Mid(p_NomFil, 1, r_int_Contad) & "D" & Format(date, "yymmdd") & ".064"

   r_int_NumFil = FreeFile
   Open r_str_NomCab For Output As r_int_NumFil
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPMAE "
   g_str_Parame = g_str_Parame & " WHERE HIPMAE_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      r_int_PerGra = g_rst_Princi!HIPMAE_PERGRA
      r_dbl_MtoPre = g_rst_Princi!HIPMAE_MTOMVI
      r_dbl_MtoGra = g_rst_Princi!HIPMAE_IMPCON + g_rst_Princi!HIPMAE_IMPNCO
      r_dbl_PorNCo = g_rst_Princi!HIPMAE_IMPNCO / (g_rst_Princi!HIPMAE_MTOMVI + g_rst_Princi!HIPMAE_INTCAP) * 100
      r_dbl_PorNCo = CDbl(Format(r_dbl_PorNCo, "######0.0000"))
      r_dbl_PorCon = g_rst_Princi!HIPMAE_IMPCON / (g_rst_Princi!HIPMAE_MTOMVI + g_rst_Princi!HIPMAE_INTCAP) * 100
      r_dbl_PorCon = CDbl(Format(r_dbl_PorCon, "######0.0000"))
      r_dbl_TotPre = g_rst_Princi!HIPMAE_IMPCON + g_rst_Princi!HIPMAE_IMPNCO
      r_str_OpeMVi = Trim(g_rst_Princi!HIPMAE_OPEMVI)
      
      Print #r_int_NumFil, "1" & " " & _
                           Format(CDate(gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECDES))), "yyyymmdd") & " " & _
                           Space(16 - Len(Trim(g_rst_Princi!HIPMAE_OPEMVI))) & Trim(g_rst_Princi!HIPMAE_OPEMVI) & " " & _
                           Space(20 - Len(Trim(g_rst_Princi!HIPMAE_NUMOPE))) & Trim(g_rst_Princi!HIPMAE_NUMOPE) & " " & _
                           CStr(g_rst_Princi!HIPMAE_TDOCLI) & " " & _
                           Space(12 - Len(Trim(g_rst_Princi!HIPMAE_NDOCLI))) & Trim(g_rst_Princi!HIPMAE_NDOCLI) & " " & _
                           Mid(moddat_gf_Consulta_ParDes("237", CStr(g_rst_Princi!HIPMAE_MONEDA)) & Space(1), 1, 1) & " " & _
                           Space(12 - Len(gf_ComaDecimal(Format(g_rst_Princi!HIPMAE_MTOMVI, "########0.00"), 2))) & gf_ComaDecimal(Format(g_rst_Princi!HIPMAE_MTOMVI, "########0.00"), 2) & " " & _
                           Space(12 - Len(gf_ComaDecimal(Format(g_rst_Princi!HIPMAE_IMPNCO, "########0.00"), 2))) & gf_ComaDecimal(Format(g_rst_Princi!HIPMAE_IMPNCO, "########0.00"), 2) & " " & _
                           Space(12 - Len(gf_ComaDecimal(Format(g_rst_Princi!HIPMAE_IMPCON, "########0.00"), 2))) & gf_ComaDecimal(Format(g_rst_Princi!HIPMAE_IMPCON, "########0.00"), 2) & " " & _
                           Space(12 - Len(gf_ComaDecimal(CStr(r_dbl_PorNCo), 4))) & gf_ComaDecimal(CStr(r_dbl_PorNCo), 4) & " " & _
                           Space(12 - Len(gf_ComaDecimal(CStr(r_dbl_PorCon), 4))) & gf_ComaDecimal(CStr(r_dbl_PorCon), 4) & " " & _
                           Format(g_rst_Princi!HIPMAE_NUMCUO + g_rst_Princi!HIPMAE_PERGRA, "000") & " " & _
                           Space(3 - Len(CStr(g_rst_Princi!HIPMAE_PERGRA))) & CStr(g_rst_Princi!HIPMAE_PERGRA) & " " & _
                           Space(12 - Len(gf_ComaDecimal(CStr(r_dbl_TotPre), 2))) & gf_ComaDecimal(CStr(r_dbl_TotPre), 2)

   End If
   
   'Cerrando Archivo Cabecera
   Close #r_int_NumFil

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Creando Detalle
   r_int_NumFil = FreeFile
   Open r_str_NomDet For Output As r_int_NumFil
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPCUO "
   g_str_Parame = g_str_Parame & " WHERE HIPCUO_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_TIPCRO = 3 "
   g_str_Parame = g_str_Parame & " ORDER BY HIPCUO_NUMCUO ASC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         r_dbl_SalCap = g_rst_Princi!HIPCUO_SALCAP
         
         If r_int_PerGra > 0 Then
            If g_rst_Princi!HIPCUO_NUMCUO < r_int_PerGra Then
               r_dbl_SalCap = r_dbl_MtoPre
            ElseIf g_rst_Princi!HIPCUO_NUMCUO = r_int_PerGra Then
               r_dbl_SalCap = r_dbl_MtoGra
            End If
         End If
         
         r_dbl_TotCuo = g_rst_Princi!HIPCUO_CAPITA + g_rst_Princi!HIPCUO_INTERE + g_rst_Princi!HIPCUO_DESORG + g_rst_Princi!HIPCUO_VIVORG + g_rst_Princi!HIPCUO_OTRORG
      
         Print #r_int_NumFil, Space(16 - Len(Trim(r_str_OpeMVi))) & Trim(r_str_OpeMVi) & " " & _
                              Space(3 - Len(CStr(g_rst_Princi!HIPCUO_NUMCUO))) & CStr(g_rst_Princi!HIPCUO_NUMCUO) & " " & _
                              Format(CDate(gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECVCT))), "yyyymmdd") & " " & _
                              Space(12 - Len(gf_ComaDecimal(Format(g_rst_Princi!HIPCUO_CAPITA, "########0.00"), 2))) & gf_ComaDecimal(Format(g_rst_Princi!HIPCUO_CAPITA, "########0.00"), 2) & " " & _
                              Space(12 - Len(gf_ComaDecimal(Format(g_rst_Princi!HIPCUO_INTERE, "########0.00"), 2))) & gf_ComaDecimal(Format(g_rst_Princi!HIPCUO_INTERE, "########0.00"), 2) & " " & _
                              Space(12 - Len(gf_ComaDecimal(Format(g_rst_Princi!HIPCUO_DESORG, "########0.00"), 2))) & gf_ComaDecimal(Format(g_rst_Princi!HIPCUO_DESORG, "########0.00"), 2) & " " & _
                              Space(12 - Len(gf_ComaDecimal(Format(g_rst_Princi!HIPCUO_VIVORG, "########0.00"), 2))) & gf_ComaDecimal(Format(g_rst_Princi!HIPCUO_VIVORG, "########0.00"), 2) & " " & _
                              Space(12 - Len(gf_ComaDecimal(Format(g_rst_Princi!HIPCUO_OTRORG, "########0.00"), 2))) & gf_ComaDecimal(Format(g_rst_Princi!HIPCUO_OTRORG, "########0.00"), 2) & " " & _
                              Space(12 - Len(gf_ComaDecimal(Format(r_dbl_TotCuo, "########0.00"), 2))) & gf_ComaDecimal(Format(r_dbl_TotCuo, "########0.00"), 2) & " " & _
                              Space(12 - Len(gf_ComaDecimal(Format(r_dbl_SalCap, "########0.00"), 2))) & gf_ComaDecimal(Format(r_dbl_SalCap, "########0.00"), 2)

         DoEvents
         g_rst_Princi.MoveNext
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   'Cerrando Archivo Detalle
   Close #r_int_NumFil
End Sub

'Private Sub fs_LiqDes()
'Dim r_rst_Genera     As ADODB.Recordset
'Dim r_str_Direcc     As String
'Dim r_str_Distri     As String
'Dim r_str_Modali     As String
'
'   Screen.MousePointer = 11
'
'   g_str_Parame = ""
'   g_str_Parame = g_str_Parame & "DELETE FROM RPT_FORDES "
'   g_str_Parame = g_str_Parame & " WHERE FORDES_NUMOPE = '" & moddat_g_str_NumOpe & "' "
'
'   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
'      Exit Sub
'   End If
'
'   'Leyendo Tabla de Créditos
'   g_str_Parame = ""
'   g_str_Parame = g_str_Parame & "SELECT * "
'   g_str_Parame = g_str_Parame & "  FROM CRE_HIPMAE "
'   g_str_Parame = g_str_Parame & " WHERE HIPMAE_NUMOPE = '" & moddat_g_str_NumOpe & "' "
'
'   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
'      Exit Sub
'   End If
'
'   g_rst_Princi.MoveFirst
'
'   'Leyendo Tabla de Desembolso
'   g_str_Parame = ""
'   g_str_Parame = g_str_Parame & "SELECT * "
'   g_str_Parame = g_str_Parame & "  FROM CRE_HIPDES "
'   g_str_Parame = g_str_Parame & " WHERE HIPDES_NUMOPE = '" & moddat_g_str_NumOpe & "' "
'
'   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Genera, 3) Then
'      Exit Sub
'   End If
'
'   g_rst_Princi.MoveFirst
'
'   'Para obtener Modalidad
'   r_str_Modali = ""
'   If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), g_rst_Princi!HIPMAE_CODPRD, g_rst_Princi!HIPMAE_CODSUB, "003", Format(CInt(CStr(g_rst_Princi!HIPMAE_CODMOD)), "000")) Then
'      r_str_Modali = moddat_g_arr_Genera(1).Genera_Nombre
'   End If
'
'   'Para obtener Dirección de Inmueble
'   Call moddat_gs_Consulta_DatInm(g_rst_Princi!hipmae_numsol, r_str_Direcc, r_str_Distri)
'
'   'Insertando Registro
'   g_str_Parame = ""
'   g_str_Parame = g_str_Parame & "INSERT INTO RPT_FORDES("
'   g_str_Parame = g_str_Parame & "FORDES_NUMOPE, "
'   g_str_Parame = g_str_Parame & "FORDES_NUMSOL, "
'   g_str_Parame = g_str_Parame & "FORDES_MODALI, "
'   g_str_Parame = g_str_Parame & "FORDES_DIRINM, "
'   g_str_Parame = g_str_Parame & "FORDES_DSTINM, "
'   g_str_Parame = g_str_Parame & "FORDES_EMPSEG, "
'   g_str_Parame = g_str_Parame & "FORDES_TIPSEG, "
'   g_str_Parame = g_str_Parame & "FORDES_BANDES, "
'   g_str_Parame = g_str_Parame & "FORDES_NUMCTA) "
'
'   g_str_Parame = g_str_Parame & "VALUES ("
'   g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "', "
'   g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumSol & "', "
'   g_str_Parame = g_str_Parame & "'" & r_str_Modali & "', "
'   g_str_Parame = g_str_Parame & "'" & r_str_Direcc & "', "
'   g_str_Parame = g_str_Parame & "'" & r_str_Distri & "', "
'   g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ComSeg(g_rst_Princi!HIPMAE_SEGPRE & "") & "', "
'   g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_TipSeg(g_rst_Princi!HIPMAE_SEGPRE, g_rst_Princi!HIPMAE_TIPSEG) & "', "
'   g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("516", r_rst_Genera!HIPDES_BANCGO & "") & "', "
'   g_str_Parame = g_str_Parame & "'" & Trim(r_rst_Genera!HIPDES_CTACGO & "") & "' )"
'
'
'   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
'      Exit Sub
'   End If
'
'   r_rst_Genera.Close
'   Set r_rst_Genera = Nothing
'
'   g_rst_Princi.Close
'   Set g_rst_Princi = Nothing
'
'   Screen.MousePointer = 0
'
'   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
'   crp_Imprim.DataFiles(0) = "CRE_HIPMAE"
'   crp_Imprim.DataFiles(1) = "CRE_HIPDES"
'   crp_Imprim.DataFiles(2) = "CLI_DATGEN"
'   crp_Imprim.DataFiles(3) = "RPT_FORDES"
'   crp_Imprim.DataFiles(4) = "CRE_PRODUC"
'   crp_Imprim.DataFiles(5) = ""
'   crp_Imprim.DataFiles(6) = ""
'
'   crp_Imprim.SelectionFormula = "{RPT_FORDES.FORDES_NUMOPE} = '" & moddat_g_str_NumOpe & "' "
'   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_LIQDES_11.RPT"
'   crp_Imprim.Action = 1
'End Sub
'
'Private Sub fs_HojRes()
'   Dim r_rst_Genera     As ADODB.Recordset
'   Dim r_str_Direcc     As String
'   Dim r_str_Distri     As String
'   Dim r_str_Modali     As String
'   Dim r_dbl_PorITF     As Double
'   Dim r_dbl_IntMor     As Double
'   Dim r_dbl_PrePag     As Double
'   Dim r_dbl_LevHip     As Double
'   Dim r_dbl_CamTas     As Double
'   Dim r_dbl_CobJud     As Double
'   Dim r_dbl_CanMVi     As Double
'   Dim r_dbl_CobDi1     As Double
'   Dim r_dbl_CobIm1     As Double
'   Dim r_dbl_CobDi2     As Double
'   Dim r_dbl_CobIm2     As Double
'   Dim r_dbl_CobDi3     As Double
'   Dim r_dbl_CobIm3     As Double
'   Dim r_int_Indice     As Integer
'
'   Screen.MousePointer = 11
'
'   g_str_Parame = ""
'   g_str_Parame = g_str_Parame & "DELETE FROM RPT_FORDES "
'   g_str_Parame = g_str_Parame & " WHERE FORDES_NUMOPE = '" & moddat_g_str_NumOpe & "' "
'
'   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
'      Exit Sub
'   End If
'
'   'Leyendo Tabla de Créditos
'   g_str_Parame = ""
'   g_str_Parame = g_str_Parame & "SELECT * "
'   g_str_Parame = g_str_Parame & "  FROM CRE_HIPMAE "
'   g_str_Parame = g_str_Parame & " WHERE HIPMAE_NUMOPE = '" & moddat_g_str_NumOpe & "' "
'
'   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
'      Exit Sub
'   End If
'
'   g_rst_Princi.MoveFirst
'
'   'Leyendo Tabla de Desembolso
'   g_str_Parame = ""
'   g_str_Parame = g_str_Parame & "SELECT * "
'   g_str_Parame = g_str_Parame & "  FROM CRE_HIPDES "
'   g_str_Parame = g_str_Parame & " WHERE HIPDES_NUMOPE = '" & moddat_g_str_NumOpe & "' "
'
'   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Genera, 3) Then
'      Exit Sub
'   End If
'
'   g_rst_Princi.MoveFirst
'
'   'Para obtener Modalidad
'   r_str_Modali = ""
'   If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), g_rst_Princi!HIPMAE_CODPRD, g_rst_Princi!HIPMAE_CODSUB, "003", Format(CInt(CStr(g_rst_Princi!HIPMAE_CODMOD)), "000")) Then
'      r_str_Modali = moddat_g_arr_Genera(1).Genera_Nombre
'   End If
'
'   'Para obtener Interes Moratorio
'   r_dbl_IntMor = 0
'   If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), moddat_g_str_CodPrd, moddat_g_str_CodSub, "002", "201") Then
'      r_dbl_IntMor = moddat_g_arr_Genera(1).Genera_Cantid
'   End If
'
'   'Para obtener ITF
'   r_dbl_PorITF = opecaj_gf_Consulta_ITF(Format(CDate(moddat_g_str_FecSis), "yyyymmdd"), 1)
'
'   'Para obtener Dirección de Inmueble
'   Call moddat_gs_Consulta_DatInm(g_rst_Princi!hipmae_numsol, r_str_Direcc, r_str_Distri)
'
'   'Otras Comisiones - Prepagos
'   r_dbl_PrePag = 0
'   If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), g_rst_Princi!HIPMAE_CODPRD, g_rst_Princi!HIPMAE_CODSUB, "002", "501") Then
'      r_dbl_PrePag = moddat_g_arr_Genera(1).Genera_Cantid
'   End If
'
'   'Otras Comisiones - Levantamiento de Hipoteca
'   r_dbl_LevHip = 0
'   'If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), g_rst_Princi!HIPMAE_CODPRD, g_rst_Princi!HIPMAE_CODSUB, "002", "701") Then
'   '   r_dbl_LevHip = moddat_g_arr_Genera(1).Genera_Cantid
'   'End If
'
'   'Otras Comisiones - Cambio de Fecha, Tasa de Interes, Moneda o Cuota
'   r_dbl_CamTas = 0
'   If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), g_rst_Princi!HIPMAE_CODPRD, g_rst_Princi!HIPMAE_CODSUB, "002", "702") Then
'      r_dbl_CamTas = moddat_g_arr_Genera(1).Genera_Cantid
'   End If
'
'   'Otras Comisiones - Cobranza Judicial
'   r_dbl_CobJud = 0
'   If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), g_rst_Princi!HIPMAE_CODPRD, g_rst_Princi!HIPMAE_CODSUB, "002", "703") Then
'      r_dbl_CobJud = moddat_g_arr_Genera(1).Genera_Cantid
'   End If
'
'   'Otras Comisiones - Caducidad del Servicio MiVivienda
'   r_dbl_CanMVi = 0
'   If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), g_rst_Princi!HIPMAE_CODPRD, g_rst_Princi!HIPMAE_CODSUB, "002", "901") Then
'      r_dbl_CanMVi = moddat_g_arr_Genera(1).Genera_Cantid
'   End If
'
'   'Gastos de Cobranzas
'   r_dbl_CobDi1 = 0
'   r_dbl_CobIm1 = 0
'   r_dbl_CobDi2 = 0
'   r_dbl_CobIm2 = 0
'   r_dbl_CobDi3 = 0
'   r_dbl_CobIm3 = 0
'
'   g_str_Parame = ""
'   g_str_Parame = g_str_Parame & "SELECT * "
'   g_str_Parame = g_str_Parame & "  FROM OPE_GASCOB "
'   g_str_Parame = g_str_Parame & " WHERE GASCOB_CODPRD = '" & g_rst_Princi!HIPMAE_CODPRD & "' "
'   g_str_Parame = g_str_Parame & "   AND GASCOB_CODSUB = '" & g_rst_Princi!HIPMAE_CODSUB & "' "
'   g_str_Parame = g_str_Parame & "   AND GASCOB_IMPORT > 0 "
'   g_str_Parame = g_str_Parame & " ORDER BY GASCOB_DIAINI ASC "
'
'   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
'       Exit Sub
'   End If
'
'   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
'      g_rst_Genera.MoveFirst
'
'      r_int_Indice = 1
'      Do While Not g_rst_Genera.EOF
'         If r_int_Indice = 1 Then
'            r_dbl_CobDi1 = g_rst_Genera!GASCOB_DIAINI
'            r_dbl_CobIm1 = g_rst_Genera!GASCOB_IMPORT
'         ElseIf r_int_Indice = 2 Then
'            r_dbl_CobDi2 = g_rst_Genera!GASCOB_DIAINI
'            r_dbl_CobIm2 = g_rst_Genera!GASCOB_IMPORT
'         ElseIf r_int_Indice = 3 Then
'            r_dbl_CobDi3 = g_rst_Genera!GASCOB_DIAINI
'            r_dbl_CobIm3 = g_rst_Genera!GASCOB_IMPORT
'         End If
'
'         r_int_Indice = r_int_Indice + 1
'         g_rst_Genera.MoveNext
'         DoEvents
'      Loop
'
'      g_rst_Genera.Close
'      Set g_rst_Genera = Nothing
'   End If
'
'   'Insertando Registro
'   g_str_Parame = ""
'   g_str_Parame = g_str_Parame & "INSERT INTO RPT_FORDES("
'   g_str_Parame = g_str_Parame & "FORDES_NUMOPE, "
'   g_str_Parame = g_str_Parame & "FORDES_NUMSOL, "
'   g_str_Parame = g_str_Parame & "FORDES_MODALI, "
'   g_str_Parame = g_str_Parame & "FORDES_DIRINM, "
'   g_str_Parame = g_str_Parame & "FORDES_DSTINM, "
'   g_str_Parame = g_str_Parame & "FORDES_EMPSEG, "
'   g_str_Parame = g_str_Parame & "FORDES_TIPSEG, "
'   g_str_Parame = g_str_Parame & "FORDES_BANDES, "
'   g_str_Parame = g_str_Parame & "FORDES_NUMCTA,"
'   g_str_Parame = g_str_Parame & "FORDES_TASMOR, "
'   g_str_Parame = g_str_Parame & "FORDES_PORITF, "
'   g_str_Parame = g_str_Parame & "FORDES_GASTAS, "
'   g_str_Parame = g_str_Parame & "FORDES_GASNOT, "
'   g_str_Parame = g_str_Parame & "FORDES_ESTTIT, "
'   g_str_Parame = g_str_Parame & "FORDES_EVACRE, "
'   g_str_Parame = g_str_Parame & "FORDES_ADMTAS, "
'   g_str_Parame = g_str_Parame & "FORDES_REDCON, "
'   g_str_Parame = g_str_Parame & "FORDES_BLQREG, "
'   g_str_Parame = g_str_Parame & "FORDES_PREPAG, "
'   g_str_Parame = g_str_Parame & "FORDES_LEVHIP, "
'   g_str_Parame = g_str_Parame & "FORDES_CAMTAS, "
'   g_str_Parame = g_str_Parame & "FORDES_COBJUD, "
'   g_str_Parame = g_str_Parame & "FORDES_COBDI1, "
'   g_str_Parame = g_str_Parame & "FORDES_COBDI2, "
'   g_str_Parame = g_str_Parame & "FORDES_COBDI3, "
'   g_str_Parame = g_str_Parame & "FORDES_COBIM1, "
'   g_str_Parame = g_str_Parame & "FORDES_COBIM2, "
'   g_str_Parame = g_str_Parame & "FORDES_COBIM3) "
'
'   g_str_Parame = g_str_Parame & "VALUES ("
'   g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "', "
'   g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumSol & "', "
'   g_str_Parame = g_str_Parame & "'" & r_str_Modali & "', "
'   g_str_Parame = g_str_Parame & "'" & r_str_Direcc & "', "
'   g_str_Parame = g_str_Parame & "'" & r_str_Distri & "', "
'   g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ComSeg(g_rst_Princi!HIPMAE_SEGPRE & "") & "', "
'   g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_TipSeg(g_rst_Princi!HIPMAE_SEGPRE, g_rst_Princi!HIPMAE_TIPSEG) & "', "
'   g_str_Parame = g_str_Parame & "'" & moddat_gf_Consulta_ParDes("516", r_rst_Genera!HIPDES_BANCGO & "") & "', "
'   g_str_Parame = g_str_Parame & "'" & Trim(r_rst_Genera!HIPDES_CTACGO & "") & "', "
'   g_str_Parame = g_str_Parame & CStr(r_dbl_IntMor) & ", "
'   g_str_Parame = g_str_Parame & CStr(r_dbl_PorITF) & ", "
'   g_str_Parame = g_str_Parame & CStr(l_dbl_ImpTas) & ", "
'   g_str_Parame = g_str_Parame & CStr(l_dbl_ImpNot) & ", "
'   g_str_Parame = g_str_Parame & CStr(l_dbl_ImpEst) & ", "
'   g_str_Parame = g_str_Parame & CStr(l_dbl_ImpEva) & ", "
'   g_str_Parame = g_str_Parame & CStr(l_dbl_ImpAdm) & ", "
'   g_str_Parame = g_str_Parame & CStr(l_dbl_ImpRed) & ", "
'   g_str_Parame = g_str_Parame & CStr(l_dbl_ImpBlq) & ", "
'   g_str_Parame = g_str_Parame & CStr(r_dbl_PrePag) & ", "
'   g_str_Parame = g_str_Parame & CStr(r_dbl_LevHip) & ", "
'   g_str_Parame = g_str_Parame & CStr(r_dbl_CamTas) & ", "
'   g_str_Parame = g_str_Parame & CStr(r_dbl_CobJud) & ", "
'   g_str_Parame = g_str_Parame & CStr(r_dbl_CobDi1) & ", "
'   g_str_Parame = g_str_Parame & CStr(r_dbl_CobDi2) & ", "
'   g_str_Parame = g_str_Parame & CStr(r_dbl_CobDi3) & ", "
'   g_str_Parame = g_str_Parame & CStr(r_dbl_CobIm1) & ", "
'   g_str_Parame = g_str_Parame & CStr(r_dbl_CobIm2) & ", "
'   g_str_Parame = g_str_Parame & CStr(r_dbl_CobIm3) & ") "
'
'   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
'      Exit Sub
'   End If
'
'   r_rst_Genera.Close
'   Set r_rst_Genera = Nothing
'
'   g_rst_Princi.Close
'   Set g_rst_Princi = Nothing
'
'   Screen.MousePointer = 0
'
'   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
'   crp_Imprim.DataFiles(0) = "CRE_HIPMAE"
'   crp_Imprim.DataFiles(1) = "CRE_HIPDES"
'   crp_Imprim.DataFiles(2) = "CLI_DATGEN"
'   crp_Imprim.DataFiles(3) = "RPT_FORDES"
'   crp_Imprim.DataFiles(4) = "CRE_PRODUC"
'   crp_Imprim.DataFiles(5) = "TRA_EVALEG"
'   crp_Imprim.DataFiles(6) = "TRA_POLIZA"
'   crp_Imprim.SelectionFormula = "{RPT_FORDES.FORDES_NUMOPE} = '" & moddat_g_str_NumOpe & "' "
'
'   If moddat_g_str_CodPrd = "002" Or moddat_g_str_CodPrd = "011" Or moddat_g_str_CodPrd = "019" Then
'      crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_HOJRES_11.RPT"
'   ElseIf moddat_g_str_CodPrd = "003" Then
'      crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_HOJRES_12.RPT"
'   ElseIf moddat_g_str_CodPrd = "004" Then
'      crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_HOJRES_13.RPT"
'   ElseIf moddat_g_str_CodPrd = "007" Or moddat_g_str_CodPrd = "009" Or moddat_g_str_CodPrd = "010" Or moddat_g_str_CodPrd = "013" Or moddat_g_str_CodPrd = "014" Or moddat_g_str_CodPrd = "015" Or moddat_g_str_CodPrd = "016" Or moddat_g_str_CodPrd = "017" Or moddat_g_str_CodPrd = "018" Then
'      crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_HOJRES_14.RPT"
'   End If
'
'   crp_Imprim.Action = 1
'End Sub
'
'Private Sub fs_Cronog_MiCasita()
'   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
'   crp_Imprim.DataFiles(0) = "CRE_HIPCUO"
'   crp_Imprim.DataFiles(1) = "CRE_HIPMAE"
'   crp_Imprim.DataFiles(2) = "CLI_DATGEN"
'   crp_Imprim.DataFiles(3) = "CRE_PRODUC"
'   crp_Imprim.DataFiles(4) = ""
'   crp_Imprim.DataFiles(5) = ""
'   crp_Imprim.DataFiles(6) = ""
'
'   crp_Imprim.SelectionFormula = "{CRE_HIPCUO.HIPCUO_NUMOPE} = '" & moddat_g_str_NumOpe & "' AND {CRE_HIPCUO.HIPCUO_TIPCRO} = 1 "
'   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_CROPAG_11.RPT"
'   crp_Imprim.Action = 1
'End Sub
'
'Private Sub fs_Cronog_Mivivienda_NCoCli()
'   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
'   crp_Imprim.DataFiles(0) = "CRE_HIPCUO"
'   crp_Imprim.DataFiles(1) = "CRE_HIPMAE"
'   crp_Imprim.DataFiles(2) = "CLI_DATGEN"
'   crp_Imprim.DataFiles(3) = "CRE_PRODUC"
'   crp_Imprim.DataFiles(4) = ""
'   crp_Imprim.DataFiles(5) = ""
'   crp_Imprim.DataFiles(6) = ""
'
'   crp_Imprim.SelectionFormula = "{CRE_HIPCUO.HIPCUO_NUMOPE} = '" & moddat_g_str_NumOpe & "' AND {CRE_HIPCUO.HIPCUO_TIPCRO} = 1 "
'   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_CROPAG_12.RPT"
'   crp_Imprim.Action = 1
'End Sub
'
'Private Sub fs_Cronog_Mivivienda_ConCli()
'   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
'   crp_Imprim.DataFiles(0) = "CRE_HIPCUO"
'   crp_Imprim.DataFiles(1) = "CRE_HIPMAE"
'   crp_Imprim.DataFiles(2) = "CLI_DATGEN"
'   crp_Imprim.DataFiles(3) = "CRE_PRODUC"
'   crp_Imprim.DataFiles(4) = ""
'   crp_Imprim.DataFiles(5) = ""
'   crp_Imprim.DataFiles(6) = ""
'
'   crp_Imprim.SelectionFormula = "{CRE_HIPCUO.HIPCUO_NUMOPE} = '" & moddat_g_str_NumOpe & "' AND {CRE_HIPCUO.HIPCUO_TIPCRO} = 2 "
'   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_CROPAG_13.RPT"
'   crp_Imprim.Action = 1
'End Sub
'
'Private Sub fs_Cronog_Mivivienda_ConMVi()
'   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
'   crp_Imprim.DataFiles(0) = "CRE_HIPCUO"
'   crp_Imprim.DataFiles(1) = "CRE_HIPMAE"
'   crp_Imprim.DataFiles(2) = "CLI_DATGEN"
'   crp_Imprim.DataFiles(3) = "CRE_PRODUC"
'
'   If moddat_g_str_CodPrd = "004" Or moddat_g_str_CodPrd = "007" Or moddat_g_str_CodPrd = "009" Or moddat_g_str_CodPrd = "010" Or moddat_g_str_CodPrd = "013" Or moddat_g_str_CodPrd = "014" Or moddat_g_str_CodPrd = "015" Or moddat_g_str_CodPrd = "016" Or moddat_g_str_CodPrd = "017" Or moddat_g_str_CodPrd = "018" Then
'      crp_Imprim.DataFiles(4) = "TRA_EVACOF"
'   Else
'      crp_Imprim.DataFiles(4) = ""
'   End If
'
'   crp_Imprim.DataFiles(5) = ""
'   crp_Imprim.DataFiles(6) = ""
'   crp_Imprim.SelectionFormula = "{CRE_HIPCUO.HIPCUO_NUMOPE} = '" & moddat_g_str_NumOpe & "' AND {CRE_HIPCUO.HIPCUO_TIPCRO} = 4 "
'
'   If moddat_g_str_CodPrd = "001" Or moddat_g_str_CodPrd = "003" Then
'      crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_CROPAG_14.RPT"
'   ElseIf moddat_g_str_CodPrd = "004" Or moddat_g_str_CodPrd = "007" Or moddat_g_str_CodPrd = "009" Or moddat_g_str_CodPrd = "010" Or moddat_g_str_CodPrd = "013" Or moddat_g_str_CodPrd = "014" Or moddat_g_str_CodPrd = "015" Or moddat_g_str_CodPrd = "016" Or moddat_g_str_CodPrd = "017" Or moddat_g_str_CodPrd = "018" Then
'      crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_CROPAG_15.RPT"
'   End If
'
'   crp_Imprim.Action = 1
'End Sub
'
'Private Sub fs_Cronog_Mivivienda_NCoMVi()
'   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
'   crp_Imprim.DataFiles(0) = "CRE_HIPCUO"
'   crp_Imprim.DataFiles(1) = "CRE_HIPMAE"
'   crp_Imprim.DataFiles(2) = "CLI_DATGEN"
'   crp_Imprim.DataFiles(3) = "CRE_PRODUC"
'
'   If moddat_g_str_CodPrd = "004" Or moddat_g_str_CodPrd = "007" Or moddat_g_str_CodPrd = "009" Or moddat_g_str_CodPrd = "010" Or moddat_g_str_CodPrd = "013" Or moddat_g_str_CodPrd = "014" Or moddat_g_str_CodPrd = "015" Or moddat_g_str_CodPrd = "016" Or moddat_g_str_CodPrd = "017" Or moddat_g_str_CodPrd = "018" Or moddat_g_str_CodPrd = "019" Then
'      crp_Imprim.DataFiles(4) = "TRA_EVACOF"
'   Else
'      crp_Imprim.DataFiles(4) = ""
'   End If
'
'   crp_Imprim.DataFiles(5) = ""
'   crp_Imprim.DataFiles(6) = ""
'   crp_Imprim.SelectionFormula = "{CRE_HIPCUO.HIPCUO_NUMOPE} = '" & moddat_g_str_NumOpe & "' AND {CRE_HIPCUO.HIPCUO_TIPCRO} = 3 "
'
'   If moddat_g_str_CodPrd = "004" Or moddat_g_str_CodPrd = "007" Or moddat_g_str_CodPrd = "009" Or moddat_g_str_CodPrd = "010" Or moddat_g_str_CodPrd = "013" Or moddat_g_str_CodPrd = "014" Or moddat_g_str_CodPrd = "015" Or moddat_g_str_CodPrd = "016" Or moddat_g_str_CodPrd = "017" Or moddat_g_str_CodPrd = "018" Or moddat_g_str_CodPrd = "019" Then
'      crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_CROPAG_16.RPT"
'   End If
'
'   crp_Imprim.Action = 1
'End Sub
'
'Private Sub fs_Cronog_Mivivienda_NCoCof()
'   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
'   crp_Imprim.DataFiles(0) = "CRE_HIPCUO"
'   crp_Imprim.DataFiles(1) = "CRE_HIPMAE"
'   crp_Imprim.DataFiles(2) = "CLI_DATGEN"
'   crp_Imprim.DataFiles(3) = "CRE_PRODUC"
'   crp_Imprim.DataFiles(4) = "TRA_EVACOF"
'   crp_Imprim.DataFiles(5) = ""
'   crp_Imprim.DataFiles(6) = ""
'
'   crp_Imprim.SelectionFormula = "{CRE_HIPCUO.HIPCUO_NUMOPE} = '" & moddat_g_str_NumOpe & "' AND {CRE_HIPCUO.HIPCUO_TIPCRO} = 5 "
'   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_CROPAG_17.RPT"
'   crp_Imprim.Action = 1
'End Sub
'
'Private Sub fs_ComPag()
'Dim r_str_CodSuc     As String
'Dim r_str_NumMov     As String
'Dim r_str_FecMov     As String
'
'   'Buscando en OPE_CAJMOV
'   g_str_Parame = ""
'   g_str_Parame = g_str_Parame & "SELECT * "
'   g_str_Parame = g_str_Parame & "  FROM OPE_CAJMOV "
'   g_str_Parame = g_str_Parame & " WHERE CAJMOV_TIPMOV = 1103 "
'   g_str_Parame = g_str_Parame & "   AND CAJMOV_NUMOPE = '" & moddat_g_str_NumOpe & "' "
'
'   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
'       Exit Sub
'   End If
'
'   g_rst_Princi.MoveFirst
'
'   r_str_CodSuc = Trim(g_rst_Princi!CAJMOV_SUCMOV)
'   r_str_NumMov = CStr(g_rst_Princi!CAJMOV_NUMMOV)
'   r_str_FecMov = CStr(g_rst_Princi!CAJMOV_FECMOV)
'
'   g_rst_Princi.Close
'   Set g_rst_Princi = Nothing
'
'   'Borrar Spool de PC (Cabecera)
'   g_str_Parame = ""
'   g_str_Parame = g_str_Parame & "DELETE FROM RPT_COMPGC "
'   g_str_Parame = g_str_Parame & " WHERE COMPGC_CODTER = '" & modgen_g_str_NombPC & "' "
'
'   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
'       Exit Sub
'   End If
'
'   'Borrar Spool de PC (Detalle)
'   g_str_Parame = ""
'   g_str_Parame = g_str_Parame & "DELETE FROM RPT_COMPGD "
'   g_str_Parame = g_str_Parame & " WHERE COMPGD_CODTER = '" & modgen_g_str_NombPC & "' "
'
'   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
'       Exit Sub
'   End If
'
'   Screen.MousePointer = 11
'   Call opecaj_gs_ComPago(r_str_CodSuc, r_str_NumMov, r_str_FecMov, 1, 1)
'   Screen.MousePointer = 0
'
'   'Se conecta al crystal report
'   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
'   crp_Imprim.DataFiles(0) = "RPT_COMPGC"
'   crp_Imprim.DataFiles(1) = "RPT_COMPGD"
'   crp_Imprim.DataFiles(2) = ""
'   crp_Imprim.DataFiles(3) = ""
'   crp_Imprim.DataFiles(4) = ""
'   crp_Imprim.DataFiles(5) = ""
'   crp_Imprim.DataFiles(6) = ""
'
'   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_COMPAG_01.RPT"
'   crp_Imprim.SelectionFormula = "{RPT_COMPGC.COMPGC_CODTER} = '" & modgen_g_str_NombPC & "'"
'   crp_Imprim.Destination = crptToWindow
'   crp_Imprim.Action = 1
'End Sub
'
'Private Sub fs_LiqTipoCambio()
'   'Se conecta al crystal report
'   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
'   crp_Imprim.DataFiles(0) = "CRE_HIPDES"
'   crp_Imprim.DataFiles(1) = "CRE_HIPMAE"
'   crp_Imprim.DataFiles(2) = "CLI_DATGEN"
'   crp_Imprim.DataFiles(3) = ""
'   crp_Imprim.DataFiles(4) = ""
'   crp_Imprim.DataFiles(5) = ""
'   crp_Imprim.DataFiles(6) = ""
'
'   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_LIQDES_12.RPT"
'   crp_Imprim.SelectionFormula = "{CRE_HIPDES.HIPDES_NUMOPE} = '" & moddat_g_str_NumOpe & "'"
'   crp_Imprim.Destination = crptToWindow
'   crp_Imprim.Action = 1
'End Sub
