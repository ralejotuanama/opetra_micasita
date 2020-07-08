VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "msmapi32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_Desemb_12 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   10140
   ClientLeft      =   2355
   ClientTop       =   1680
   ClientWidth     =   11640
   Icon            =   "OpeTra_frm_072.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10140
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   10125
      Left            =   0
      TabIndex        =   0
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
         TabIndex        =   19
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
            TabIndex        =   20
            Text            =   "OpeTra_frm_072.frx":000C
            Top             =   60
            Width           =   11445
         End
      End
      Begin Threed.SSPanel SSPanel15 
         Height          =   2085
         Left            =   30
         TabIndex        =   1
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
            TabIndex        =   2
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
         TabIndex        =   3
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
            TabIndex        =   4
            Top             =   60
            Width           =   11385
            _ExtentX        =   20082
            _ExtentY        =   8281
            _Version        =   393216
            Style           =   1
            Tabs            =   9
            TabsPerRow      =   9
            TabHeight       =   520
            TabCaption(0)   =   "Cliente"
            TabPicture(0)   =   "OpeTra_frm_072.frx":0010
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "grd_Listad(0)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Inmueble"
            TabPicture(1)   =   "OpeTra_frm_072.frx":002C
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "grd_Listad(1)"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Crédito"
            TabPicture(2)   =   "OpeTra_frm_072.frx":0048
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "grd_Listad(2)"
            Tab(2).ControlCount=   1
            TabCaption(3)   =   "Ev. Crediticia"
            TabPicture(3)   =   "OpeTra_frm_072.frx":0064
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "grd_Listad(3)"
            Tab(3).ControlCount=   1
            TabCaption(4)   =   "Tasación"
            TabPicture(4)   =   "OpeTra_frm_072.frx":0080
            Tab(4).ControlEnabled=   0   'False
            Tab(4).Control(0)=   "grd_Listad(4)"
            Tab(4).ControlCount=   1
            TabCaption(5)   =   "Seguros"
            TabPicture(5)   =   "OpeTra_frm_072.frx":009C
            Tab(5).ControlEnabled=   0   'False
            Tab(5).Control(0)=   "Label5"
            Tab(5).Control(1)=   "Label4"
            Tab(5).Control(2)=   "grd_Listad(5)"
            Tab(5).Control(3)=   "txt_ObsSeg"
            Tab(5).ControlCount=   4
            TabCaption(6)   =   "Informe Legal"
            TabPicture(6)   =   "OpeTra_frm_072.frx":00B8
            Tab(6).ControlEnabled=   0   'False
            Tab(6).Control(0)=   "txt_InfLeg"
            Tab(6).ControlCount=   1
            TabCaption(7)   =   "Ev. Legal"
            TabPicture(7)   =   "OpeTra_frm_072.frx":00D4
            Tab(7).ControlEnabled=   0   'False
            Tab(7).Control(0)=   "Label3"
            Tab(7).Control(1)=   "Label2"
            Tab(7).Control(2)=   "grd_Listad(6)"
            Tab(7).Control(3)=   "txt_ComCre"
            Tab(7).ControlCount=   4
            TabCaption(8)   =   "Mivivienda / Cofide"
            TabPicture(8)   =   "OpeTra_frm_072.frx":00F0
            Tab(8).ControlEnabled=   0   'False
            Tab(8).Control(0)=   "txt_ObsMVi"
            Tab(8).Control(1)=   "grd_Listad(7)"
            Tab(8).Control(2)=   "Label6"
            Tab(8).Control(3)=   "Label7"
            Tab(8).ControlCount=   4
            Begin VB.TextBox txt_ObsMVi 
               Height          =   1155
               Left            =   -74940
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   38
               Text            =   "OpeTra_frm_072.frx":010C
               Top             =   3480
               Width           =   11235
            End
            Begin VB.TextBox txt_ComCre 
               Height          =   705
               Left            =   -74940
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   34
               Text            =   "OpeTra_frm_072.frx":0110
               Top             =   630
               Width           =   11235
            End
            Begin VB.TextBox txt_InfLeg 
               Height          =   4215
               Left            =   -74940
               MaxLength       =   8000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   33
               Text            =   "OpeTra_frm_072.frx":0114
               Top             =   390
               Width           =   11235
            End
            Begin VB.TextBox txt_ObsSeg 
               Height          =   1065
               Left            =   -74940
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   29
               Text            =   "OpeTra_frm_072.frx":0118
               Top             =   3540
               Width           =   11235
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   4215
               Index           =   0
               Left            =   60
               TabIndex        =   5
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
               Height          =   4215
               Index           =   1
               Left            =   -74940
               TabIndex        =   25
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
               Height          =   4215
               Index           =   2
               Left            =   -74940
               TabIndex        =   26
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
               Height          =   4215
               Index           =   3
               Left            =   -74940
               TabIndex        =   27
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
               Height          =   4215
               Index           =   4
               Left            =   -74940
               TabIndex        =   28
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
               Height          =   2595
               Index           =   5
               Left            =   -74940
               TabIndex        =   30
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
               Height          =   2955
               Index           =   6
               Left            =   -74940
               TabIndex        =   35
               Top             =   1650
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
               Height          =   2535
               Index           =   7
               Left            =   -74940
               TabIndex        =   39
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
               Left            =   -74940
               TabIndex        =   41
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
               Left            =   -74940
               TabIndex        =   40
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
               TabIndex        =   37
               Top             =   390
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
               TabIndex        =   36
               Top             =   1410
               Width           =   3495
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
               TabIndex        =   32
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
               TabIndex        =   31
               Top             =   390
               Width           =   3495
            End
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   6
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
            TabIndex        =   42
            Top             =   30
            Width           =   6495
            _Version        =   65536
            _ExtentX        =   11456
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Créditos Hipotecarios - Desembolso"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
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
            TabIndex        =   43
            Top             =   330
            Width           =   6495
            _Version        =   65536
            _ExtentX        =   11456
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Datos de Operación"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
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
            Picture         =   "OpeTra_frm_072.frx":011C
            Stretch         =   -1  'True
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   30
         TabIndex        =   7
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
            TabIndex        =   8
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
         Begin Threed.SSPanel pnl_NumSol 
            Height          =   315
            Left            =   8580
            TabIndex        =   9
            Top             =   390
            Width           =   2925
            _Version        =   65536
            _ExtentX        =   5159
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
         Begin Threed.SSPanel pnl_NumOpe 
            Height          =   315
            Left            =   1440
            TabIndex        =   10
            Top             =   390
            Width           =   2925
            _Version        =   65536
            _ExtentX        =   5159
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
         Begin VB.Label Label1 
            Caption         =   "Nro. Solicitud:"
            Height          =   315
            Left            =   7200
            TabIndex        =   13
            Top             =   390
            Width           =   1335
         End
         Begin VB.Label Label21 
            Caption         =   "Producto:"
            Height          =   315
            Left            =   60
            TabIndex        =   12
            Top             =   60
            Width           =   1275
         End
         Begin VB.Label Label12 
            Caption         =   "Nro. Operación:"
            Height          =   315
            Left            =   60
            TabIndex        =   11
            Top             =   390
            Width           =   1245
         End
      End
      Begin Threed.SSPanel SSPanel39 
         Height          =   645
         Left            =   30
         TabIndex        =   14
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
         Begin VB.CommandButton cmd_DiaPag 
            Height          =   585
            Left            =   4230
            Picture         =   "OpeTra_frm_072.frx":0426
            Style           =   1  'Graphical
            TabIndex        =   44
            ToolTipText     =   "Cambiar Día de Pago"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_CerPar 
            Height          =   585
            Left            =   3630
            Picture         =   "OpeTra_frm_072.frx":0730
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Regularizar Certificado de Participación"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_CarFia 
            Height          =   585
            Left            =   3030
            Picture         =   "OpeTra_frm_072.frx":0A3A
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Regularizar Carta Fianza"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ChqGer 
            Height          =   585
            Left            =   2430
            Picture         =   "OpeTra_frm_072.frx":0D44
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Regularizar Cheque de Gerencia"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_PolSeg 
            Height          =   585
            Left            =   1830
            Picture         =   "OpeTra_frm_072.frx":104E
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Regularizar Póliza de Seguros"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10920
            Picture         =   "OpeTra_frm_072.frx":1358
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Desemb 
            Height          =   585
            Left            =   1230
            Picture         =   "OpeTra_frm_072.frx":179A
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Generar Desembolso"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_072.frx":1AA4
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Imprimir Formatos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Export 
            Height          =   585
            Left            =   630
            Picture         =   "OpeTra_frm_072.frx":1EE6
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Generar Cronogramas para Mivivienda"
            Top             =   30
            Width           =   585
         End
      End
   End
End
Attribute VB_Name = "frm_Desemb_12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_str_MonTas     As String
Dim l_str_MonNot     As String
Dim l_str_MonEst     As String
Dim l_str_MonEva     As String
Dim l_str_MonAdm     As String
Dim l_str_MonRed     As String
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

Private Sub cmd_Export_Click()
  On Error GoTo cmd_ArcTxt_Error
   
   If grd_LisEva.Rows = 0 Then
      MsgBox "No ha registrado la información del Desembolso.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de generar los archivos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If

   dlg_Guarda.Filter = "Todos los archivos (*.*)|*.*"
   dlg_Guarda.ShowSave
   
   Screen.MousePointer = 11
   Call fs_Genera_Archivos(dlg_Guarda.FileName)
   Screen.MousePointer = 0
   
   MsgBox "Archivo Generado correctamente.", vbInformation, modgen_g_str_NomPlt
   
   Exit Sub
   
   
cmd_ArcTxt_Error:
   If Err.Number <> 32755 Then
      MsgBox Err.Description & " Número: " & CStr(Err.Number), vbCritical, modgen_g_str_NomPlt
   End If
   
   Exit Sub

End Sub

Private Sub cmd_Imprim_Click()
   If grd_LisEva.Rows = 0 Then
      MsgBox "No ha registrado la información del Desembolso.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   moddat_g_int_InsAct = 51
   moddat_g_int_TipRep = 0
   
   frm_Desemb_06.Show 1
   
   If moddat_g_int_TipRep > 0 Then
      Select Case moddat_g_int_TipRep
         Case 1:     Call fs_LiqDes
         Case 2:     Call fs_HojRes
         Case 3:     Call fs_Cronog_MiCasita
         Case 4:     Call fs_Cronog_Mivivienda_NCoCli
         Case 5:     Call fs_Cronog_Mivivienda_ConCli
         Case 7, 9:  Call fs_Cronog_Mivivienda_ConMVi
         Case 8:     Call fs_Cronog_Mivivienda_NCoMVi
         Case 10:    Call fs_Cronog_Mivivienda_NCoCof
      End Select
   End If
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

Private Sub Command3_Click()

End Sub

Private Sub Command4_Click()

End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt

   pnl_NumSol.Caption = Mid(moddat_g_str_NumSol, 1, 3) & "-" & Mid(moddat_g_str_NumSol, 4, 3) & "-" & Mid(moddat_g_str_NumSol, 7, 2) & "-" & Mid(moddat_g_str_NumSol, 9, 4)
   pnl_NumOpe.Caption = Mid(moddat_g_str_NumOpe, 1, 3) & "-" & Mid(moddat_g_str_NumOpe, 4, 2) & "-" & Mid(moddat_g_str_NumOpe, 6, 5)
   pnl_Produc.Caption = moddat_g_str_NomPrd
   
   Call fs_Inicia

   moddat_g_int_CygTDo = 0
   moddat_g_str_CygNDo = ""
   
   Call fs_DatCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc, 0)
   Call fs_DatCli(moddat_g_int_CygTDo, moddat_g_str_CygNDo, 1)
   
   Call fs_DatInm
   Call fs_DatCre
   Call fs_EvaCre
   Call fs_DatTas
   Call fs_DatSeg
   Call fs_DatLeg
   Call fs_DatMVi
   Call fs_DatCof
   Call fs_GasAdm

   Call fs_Carga_DatEva
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   'Datos del Cliente
   grd_Listad(0).ColWidth(0) = 3000:   grd_Listad(0).ColAlignment(0) = flexAlignLeftCenter
   grd_Listad(0).ColWidth(1) = 7940:   grd_Listad(0).ColAlignment(1) = flexAlignLeftCenter
   
   Call gs_LimpiaGrid(grd_Listad(0))

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
   
   g_str_Parame = "SELECT * FROM CRE_HIPDES WHERE "
   g_str_Parame = g_str_Parame & "HIPDES_NUMOPE = '" & moddat_g_str_NumOpe & "' "

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
            grd_LisEva.Text = moddat_gf_Consulta_ParDes("505", g_rst_Princi!HIPDES_BANCGO & "") & " (" & Trim(g_rst_Princi!HIPDES_CTACGO & "") & ")"
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
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_DatCli(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String, ByVal p_Indice As Integer)
   Dim r_str_TipCli     As String
   
   r_str_TipCli = ""

   g_str_Parame = "SELECT * FROM CLI_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_TIPDOC = " & CStr(p_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_NUMDOC = '" & p_NumDoc & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_Listad(0).Redraw = False
      
      If p_Indice = 1 Then
         r_str_TipCli = " (Cónyuge)"
         grd_Listad(0).Rows = grd_Listad(0).Rows + 1
      End If
      
      g_rst_Princi.MoveFirst
      
      grd_Listad(0).Rows = grd_Listad(0).Rows + 1
      grd_Listad(0).Row = grd_Listad(0).Rows - 1
      grd_Listad(0).Col = 0
      grd_Listad(0).Text = "Documento de Identidad" & r_str_TipCli
      
      grd_Listad(0).Col = 1
      grd_Listad(0).Text = moddat_gf_Consulta_ParDes("203", CStr(g_rst_Princi!DatGen_TipDoc)) & " - " & Trim(g_rst_Princi!DatGen_NumDoc & "")
   
      grd_Listad(0).Rows = grd_Listad(0).Rows + 1
      grd_Listad(0).Row = grd_Listad(0).Rows - 1
      grd_Listad(0).Col = 0
      grd_Listad(0).Text = "Apellidos y Nombres" & r_str_TipCli
      
      grd_Listad(0).Col = 1
      grd_Listad(0).Text = Trim(g_rst_Princi!DATGEN_APEPAT) & " " & Trim(g_rst_Princi!DATGEN_APEMAT) & IIf(Len(Trim(g_rst_Princi!DatGen_ApeCas)) > 0, " DE " & Trim(g_rst_Princi!DatGen_ApeCas), "") & " " & Trim(g_rst_Princi!DATGEN_NOMBRE)
      
      If p_Indice = 0 Then
         grd_Listad(0).Rows = grd_Listad(0).Rows + 1
         grd_Listad(0).Row = grd_Listad(0).Rows - 1
         grd_Listad(0).Col = 0
         grd_Listad(0).Text = "Estado Civil"
         
         grd_Listad(0).Col = 1
         grd_Listad(0).Text = moddat_gf_Consulta_ParDes("205", CStr(g_rst_Princi!DATGEN_ESTCIV)) & IIf(g_rst_Princi!DATGEN_ESTCIV = 2, " / " & moddat_gf_Consulta_ParDes("206", g_rst_Princi!DATGEN_REGCYG), "")
         
         If g_rst_Princi!DATGEN_ESTCIV = 2 Or g_rst_Princi!DATGEN_ESTCIV = 5 Then
            moddat_g_int_CygTDo = g_rst_Princi!DATGEN_CYGTDO
            moddat_g_str_CygNDo = Trim(g_rst_Princi!DATGEN_CYGNDO & "")
         End If
      End If

      grd_Listad(0).Rows = grd_Listad(0).Rows + 1
      grd_Listad(0).Row = grd_Listad(0).Rows - 1
      grd_Listad(0).Col = 0
      grd_Listad(0).Text = "Celular" & r_str_TipCli
      
      grd_Listad(0).Col = 1
      grd_Listad(0).Text = Trim(g_rst_Princi!DATGEN_NUMCEL & "")
      
      If p_Indice = 0 Then
         grd_Listad(0).Rows = grd_Listad(0).Rows + 1
         grd_Listad(0).Row = grd_Listad(0).Rows - 1
         grd_Listad(0).Col = 0
         grd_Listad(0).Text = "Domicilio"
         
         grd_Listad(0).Col = 1
         grd_Listad(0).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!DatGen_TipVia)) & _
                                     " " & Trim(g_rst_Princi!DatGen_NomVia) & " " & Trim(g_rst_Princi!DatGen_Numero) & _
                                     IIf(Len(Trim(g_rst_Princi!DatGen_IntDpt)) > 0, " (" & Trim(g_rst_Princi!DatGen_IntDpt) & ")", "") & _
                                     IIf(Len(Trim(g_rst_Princi!DatGen_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!DatGen_TipZon)) & " " & Trim(g_rst_Princi!DatGen_NomZon), "")
         
         grd_Listad(0).Rows = grd_Listad(0).Rows + 1
         grd_Listad(0).Row = grd_Listad(0).Rows - 1
         grd_Listad(0).Col = 0
         grd_Listad(0).Text = "Referencia"
   
         grd_Listad(0).Col = 1
         grd_Listad(0).Text = Trim(g_rst_Princi!DatGen_Refere & "")
         
         grd_Listad(0).Rows = grd_Listad(0).Rows + 1
         grd_Listad(0).Row = grd_Listad(0).Rows - 1
         grd_Listad(0).Col = 0
         grd_Listad(0).Text = "Departamento / Provincia / Distrito"
   
         grd_Listad(0).Col = 1
         grd_Listad(0).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!DatGen_Ubigeo, 2) & "0000") & _
                                     " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!DatGen_Ubigeo, 4) & "00") & _
                                     " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!DatGen_Ubigeo))
      
         grd_Listad(0).Rows = grd_Listad(0).Rows + 1
         grd_Listad(0).Row = grd_Listad(0).Rows - 1
         grd_Listad(0).Col = 0
         grd_Listad(0).Text = "Teléfono Domicilio"
   
         grd_Listad(0).Col = 1
         grd_Listad(0).Text = Trim(g_rst_Princi!DatGen_Telefo & "")
      End If
      
      grd_Listad(0).Redraw = True
      Call gs_UbiIniGrid(grd_Listad(0))
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_DatInm()
   Call gs_LimpiaGrid(grd_Listad(1))
   
   g_str_Parame = "SELECT * FROM CRE_SOLINM WHERE "
   g_str_Parame = g_str_Parame & "SOLINM_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_Listad(1).Redraw = False
      
      g_rst_Princi.MoveFirst
      
      grd_Listad(1).Rows = grd_Listad(1).Rows + 1
      grd_Listad(1).Row = grd_Listad(1).Rows - 1
      grd_Listad(1).Col = 0
      grd_Listad(1).Text = "Modalidad"
      
      If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), moddat_g_str_CodPrd, moddat_g_str_CodSub, "003", Format(CInt(CStr(g_rst_Princi!SOLINM_CODMOD)), "000")) Then
         grd_Listad(1).Col = 1
         grd_Listad(1).Text = moddat_g_arr_Genera(1).Genera_Nombre
      End If
      
      grd_Listad(1).Rows = grd_Listad(1).Rows + 1
      grd_Listad(1).Row = grd_Listad(1).Rows - 1
      grd_Listad(1).Col = 0
      grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(1).Text = "Tipo de Inmueble"
         
      grd_Listad(1).Col = 1
      grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(1).Text = moddat_gf_Consulta_ParDes("217", CStr(g_rst_Princi!SOLINM_TIPINM))
      
      grd_Listad(1).Rows = grd_Listad(1).Rows + 1
      grd_Listad(1).Row = grd_Listad(1).Rows - 1
      grd_Listad(1).Col = 0
      grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(1).Text = "Dirección"
      
      grd_Listad(1).Col = 1
      grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(1).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!SOLINM_TIPVIA)) & _
                        " " & Trim(g_rst_Princi!SOLINM_NOMVIA) & " " & Trim(g_rst_Princi!SOLINM_NUMVIA) & _
                        IIf(Len(Trim(g_rst_Princi!SOLINM_INTDPT)) > 0, " (" & Trim(g_rst_Princi!SOLINM_INTDPT) & ")", "") & _
                        IIf(Len(Trim(g_rst_Princi!SOLINM_NOMZON)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!SOLINM_TIPZON)) & " " & Trim(g_rst_Princi!SOLINM_NOMZON), "")
      
      grd_Listad(1).Rows = grd_Listad(1).Rows + 1
      grd_Listad(1).Row = grd_Listad(1).Rows - 1
      grd_Listad(1).Col = 0
      grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(1).Text = "Referencia"

      grd_Listad(1).Col = 1
      grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(1).Text = Trim(g_rst_Princi!SOLINM_REFERE & "")
      
      grd_Listad(1).Rows = grd_Listad(1).Rows + 1
      grd_Listad(1).Row = grd_Listad(1).Rows - 1
      grd_Listad(1).Col = 0
      grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(1).Text = "Estacionamiento"

      grd_Listad(1).Col = 1
      grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(1).Text = Trim(g_rst_Princi!SOLINM_ESTACI & "")
      
      grd_Listad(1).Rows = grd_Listad(1).Rows + 1
      grd_Listad(1).Row = grd_Listad(1).Rows - 1
      grd_Listad(1).Col = 0
      grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(1).Text = "Departamento / Provincia / Distrito"

      grd_Listad(1).Col = 1
      grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(1).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO, 2) & "0000") & _
                        " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO, 4) & "00") & _
                        " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!SOLINM_UBIGEO))
      
      grd_Listad(1).Rows = grd_Listad(1).Rows + 2
      grd_Listad(1).Row = grd_Listad(1).Rows - 1
      grd_Listad(1).Col = 0
      grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(1).Text = "Proyecto miCasita"

      grd_Listad(1).Col = 1
      grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(1).Text = moddat_gf_Consulta_ParDes("214", g_rst_Princi!SOLINM_PRYMCS)
      
      If g_rst_Princi!SOLINM_TABPRY = 2 Then
         If Not IsNull(g_rst_Princi!SOLINM_PRYBCO) Then
            grd_Listad(1).Rows = grd_Listad(1).Rows + 1
            grd_Listad(1).Row = grd_Listad(1).Rows - 1
            grd_Listad(1).Col = 0
            grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(1).Text = "Proyecto anclado en Otra IFI"
      
            grd_Listad(1).Col = 1
            grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(1).Text = moddat_gf_Consulta_ParDes("513", g_rst_Princi!SOLINM_PRYBCO)
         End If
         
         If Len(Trim(g_rst_Princi!SOLINM_PRYCOD)) > 0 Then
            grd_Listad(1).Rows = grd_Listad(1).Rows + 1
            grd_Listad(1).Row = grd_Listad(1).Rows - 1
            grd_Listad(1).Col = 0
            grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(1).Text = "Nombre Proyecto"
   
            grd_Listad(1).Col = 1
            grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(1).Text = moddat_gf_Consulta_NomPry(g_rst_Princi!SOLINM_PRYCOD)
         Else
            If Len(Trim(g_rst_Princi!SOLINM_PRYNOM)) > 0 Then
               grd_Listad(1).Rows = grd_Listad(1).Rows + 1
               grd_Listad(1).Row = grd_Listad(1).Rows - 1
               grd_Listad(1).Col = 0
               grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
               grd_Listad(1).Text = "Nombre Proyecto"
   
               grd_Listad(1).Col = 1
               grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
               grd_Listad(1).Text = Trim(g_rst_Princi!SOLINM_PRYNOM & "")
            End If
         End If
      
         grd_Listad(1).Rows = grd_Listad(1).Rows + 2
         grd_Listad(1).Row = grd_Listad(1).Rows - 1
         grd_Listad(1).Col = 0
         grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(1).Text = "Propietario / Promotor"
   
         grd_Listad(1).Col = 1
         grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(1).Text = moddat_gf_Consulta_ParDes("218", g_rst_Princi!SOLINM_FLGPRO)
         
         grd_Listad(1).Rows = grd_Listad(1).Rows + 1
         grd_Listad(1).Row = grd_Listad(1).Rows - 1
         grd_Listad(1).Col = 0
         grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(1).Text = "Docum. Identidad Propietario/Promotor"
   
         grd_Listad(1).Col = 1
         grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(1).Text = moddat_gf_Consulta_ParDes("236", CStr(g_rst_Princi!SOLINM_TIPDOC_PRO)) & " - " & Trim(g_rst_Princi!SOLINM_NUMDOC_PRO & "")
         
         grd_Listad(1).Rows = grd_Listad(1).Rows + 1
         grd_Listad(1).Row = grd_Listad(1).Rows - 1
         grd_Listad(1).Col = 0
         grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(1).Text = "Nombre o Razón Social"
   
         grd_Listad(1).Col = 1
         grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(1).Text = Trim(g_rst_Princi!SOLINM_RAZSOC_PRO & "")
         
         grd_Listad(1).Rows = grd_Listad(1).Rows + 1
         grd_Listad(1).Row = grd_Listad(1).Rows - 1
         grd_Listad(1).Col = 0
         grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(1).Text = "Dirección"
         
         grd_Listad(1).Col = 1
         grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(1).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!SOLINM_TIPVIA_PRO)) & _
                           " " & Trim(g_rst_Princi!SOLINM_NOMVIA_PRO) & " " & Trim(g_rst_Princi!SOLINM_NUMVIA_PRO) & _
                           IIf(Len(Trim(g_rst_Princi!SOLINM_INTDPT_PRO)) > 0, " (" & Trim(g_rst_Princi!SOLINM_INTDPT_PRO) & ")", "") & _
                           IIf(Len(Trim(g_rst_Princi!SOLINM_NOMZON_PRO)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!SOLINM_TIPZON_PRO)) & " " & Trim(g_rst_Princi!SOLINM_NOMZON_PRO), "")
         
         grd_Listad(1).Rows = grd_Listad(1).Rows + 1
         grd_Listad(1).Row = grd_Listad(1).Rows - 1
         grd_Listad(1).Col = 0
         grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(1).Text = "Referencia"
   
         grd_Listad(1).Col = 1
         grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(1).Text = Trim(g_rst_Princi!SOLINM_REFERE_PRO & "")
         
         grd_Listad(1).Rows = grd_Listad(1).Rows + 1
         grd_Listad(1).Row = grd_Listad(1).Rows - 1
         grd_Listad(1).Col = 0
         grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(1).Text = "Departamento / Provincia / Distrito"
   
         grd_Listad(1).Col = 1
         grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(1).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO_PRO, 2) & "0000") & _
                           " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO_PRO, 4) & "00") & _
                           " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!SOLINM_UBIGEO_PRO))
         
         grd_Listad(1).Rows = grd_Listad(1).Rows + 1
         grd_Listad(1).Row = grd_Listad(1).Rows - 1
         grd_Listad(1).Col = 0
         grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(1).Text = "Teléfono"
   
         grd_Listad(1).Col = 1
         grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(1).Text = Trim(g_rst_Princi!SOLINM_TELEFO_PRO & "")
         
         If g_rst_Princi!SOLINM_FLGCON = 1 Then
            grd_Listad(1).Rows = grd_Listad(1).Rows + 2
            grd_Listad(1).Row = grd_Listad(1).Rows - 1
            grd_Listad(1).Col = 0
            grd_Listad(1).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(1).Text = "Docum. Identidad Constructor"
      
            grd_Listad(1).Col = 1
            grd_Listad(1).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(1).Text = moddat_gf_Consulta_ParDes("236", CStr(g_rst_Princi!SOLINM_TIPDOC_CON)) & " - " & Trim(g_rst_Princi!SOLINM_NUMDOC_CON & "")
            
            grd_Listad(1).Rows = grd_Listad(1).Rows + 1
            grd_Listad(1).Row = grd_Listad(1).Rows - 1
            grd_Listad(1).Col = 0
            grd_Listad(1).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(1).Text = "Nombre o Razón Social"
      
            grd_Listad(1).Col = 1
            grd_Listad(1).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(1).Text = Trim(g_rst_Princi!SOLINM_RAZSOC_CON & "")
            
            grd_Listad(1).Rows = grd_Listad(1).Rows + 1
            grd_Listad(1).Row = grd_Listad(1).Rows - 1
            grd_Listad(1).Col = 0
            grd_Listad(1).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(1).Text = "Dirección"
            
            grd_Listad(1).Col = 1
            grd_Listad(1).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(1).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!SOLINM_TIPVIA_CON)) & _
                              " " & Trim(g_rst_Princi!SOLINM_NOMVIA_CON) & " " & Trim(g_rst_Princi!SOLINM_NUMVIA_CON) & _
                              IIf(Len(Trim(g_rst_Princi!SOLINM_INTDPT_CON)) > 0, " (" & Trim(g_rst_Princi!SOLINM_INTDPT_CON) & ")", "") & _
                              IIf(Len(Trim(g_rst_Princi!SOLINM_NOMZON_CON)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!SOLINM_TIPZON_CON)) & " " & Trim(g_rst_Princi!SOLINM_NOMZON_CON), "")
            
            grd_Listad(1).Rows = grd_Listad(1).Rows + 1
            grd_Listad(1).Row = grd_Listad(1).Rows - 1
            grd_Listad(1).Col = 0
            grd_Listad(1).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(1).Text = "Referencia"
      
            grd_Listad(1).Col = 1
            grd_Listad(1).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(1).Text = Trim(g_rst_Princi!SOLINM_REFERE_CON & "")
            
            grd_Listad(1).Rows = grd_Listad(1).Rows + 1
            grd_Listad(1).Row = grd_Listad(1).Rows - 1
            grd_Listad(1).Col = 0
            grd_Listad(1).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(1).Text = "Departamento / Provincia / Distrito"
      
            grd_Listad(1).Col = 1
            grd_Listad(1).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(1).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO_CON, 2) & "0000") & _
                              " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO_CON, 4) & "00") & _
                              " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!SOLINM_UBIGEO_CON))
            
            grd_Listad(1).Rows = grd_Listad(1).Rows + 1
            grd_Listad(1).Row = grd_Listad(1).Rows - 1
            grd_Listad(1).Col = 0
            grd_Listad(1).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(1).Text = "Teléfono"
      
            grd_Listad(1).Col = 1
            grd_Listad(1).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(1).Text = Trim(g_rst_Princi!SOLINM_TELEFO_CON & "")
         End If
      Else
         If Len(Trim(g_rst_Princi!SOLINM_PRYCOD & "")) > 0 Then
            If g_rst_Princi!SOLINM_PRYMCS = 1 Then
               grd_Listad(1).Rows = grd_Listad(1).Rows + 1
               grd_Listad(1).Row = grd_Listad(1).Rows - 1
               grd_Listad(1).Col = 0
               grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
               grd_Listad(1).Text = "Proyecto Vinculado"
            Else
               grd_Listad(1).Rows = grd_Listad(1).Rows + 1
               grd_Listad(1).Row = grd_Listad(1).Rows - 1
               grd_Listad(1).Col = 0
               grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
               grd_Listad(1).Text = "Entidad Financiera"
         
               grd_Listad(1).Col = 1
               grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
               grd_Listad(1).Text = moddat_gf_Consulta_ParDes("513", g_rst_Princi!SOLINM_PRYBCO)
               
               grd_Listad(1).Rows = grd_Listad(1).Rows + 1
               grd_Listad(1).Row = grd_Listad(1).Rows - 1
               grd_Listad(1).Col = 0
               grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
               grd_Listad(1).Text = "Proyecto No Vinculado"
            End If
         
            grd_Listad(1).Col = 1
            grd_Listad(1).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(1).Text = moddat_gf_Consulta_NomPry(g_rst_Princi!SOLINM_PRYCOD)
         End If
         
         If CInt(g_rst_Princi!SOLINM_CODMOD) = 1 Or CInt(g_rst_Princi!SOLINM_CODMOD) = 4 Then
            grd_Listad(1).Rows = grd_Listad(1).Rows + 2
            grd_Listad(1).Row = grd_Listad(1).Rows - 1
            grd_Listad(1).Col = 0
            grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(1).Text = "Docum. Identidad Propietario"
      
            grd_Listad(1).Col = 1
            grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(1).Text = moddat_gf_Consulta_ParDes("236", CStr(g_rst_Princi!SOLINM_TIPDOC_PRO)) & " - " & Trim(g_rst_Princi!SOLINM_NUMDOC_PRO & "")
            
            grd_Listad(1).Rows = grd_Listad(1).Rows + 1
            grd_Listad(1).Row = grd_Listad(1).Rows - 1
            grd_Listad(1).Col = 0
            grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(1).Text = "Nombre o Razón Social"
      
            grd_Listad(1).Col = 1
            grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(1).Text = Trim(g_rst_Princi!SOLINM_RAZSOC_PRO & "")
            
            grd_Listad(1).Rows = grd_Listad(1).Rows + 1
            grd_Listad(1).Row = grd_Listad(1).Rows - 1
            grd_Listad(1).Col = 0
            grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(1).Text = "Dirección"
            
            grd_Listad(1).Col = 1
            grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(1).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!SOLINM_TIPVIA_PRO)) & _
                              " " & Trim(g_rst_Princi!SOLINM_NOMVIA_PRO) & " " & Trim(g_rst_Princi!SOLINM_NUMVIA_PRO) & _
                              IIf(Len(Trim(g_rst_Princi!SOLINM_INTDPT_PRO)) > 0, " (" & Trim(g_rst_Princi!SOLINM_INTDPT_PRO) & ")", "") & _
                              IIf(Len(Trim(g_rst_Princi!SOLINM_NOMZON_PRO)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!SOLINM_TIPZON_PRO)) & " " & Trim(g_rst_Princi!SOLINM_NOMZON_PRO), "")
            
            grd_Listad(1).Rows = grd_Listad(1).Rows + 1
            grd_Listad(1).Row = grd_Listad(1).Rows - 1
            grd_Listad(1).Col = 0
            grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(1).Text = "Referencia"
      
            grd_Listad(1).Col = 1
            grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(1).Text = Trim(g_rst_Princi!SOLINM_REFERE_PRO & "")
            
            grd_Listad(1).Rows = grd_Listad(1).Rows + 1
            grd_Listad(1).Row = grd_Listad(1).Rows - 1
            grd_Listad(1).Col = 0
            grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(1).Text = "Departamento / Provincia / Distrito"
      
            grd_Listad(1).Col = 1
            grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(1).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO_PRO, 2) & "0000") & _
                              " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO_PRO, 4) & "00") & _
                              " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!SOLINM_UBIGEO_PRO))
            
            grd_Listad(1).Rows = grd_Listad(1).Rows + 1
            grd_Listad(1).Row = grd_Listad(1).Rows - 1
            grd_Listad(1).Col = 0
            grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(1).Text = "Teléfono"
      
            grd_Listad(1).Col = 1
            grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(1).Text = Trim(g_rst_Princi!SOLINM_TELEFO_PRO & "")
         Else
            'Promotor
            grd_Listad(1).Rows = grd_Listad(1).Rows + 2
            grd_Listad(1).Row = grd_Listad(1).Rows - 1
            grd_Listad(1).Col = 0
            grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(1).Text = "Doc. Ident. Promotor"
            
            grd_Listad(1).Col = 1
            grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(1).Text = CStr(g_rst_Princi!SOLINM_TIPDOC_PRO) & "-" & Trim(g_rst_Princi!SOLINM_NUMDOC_PRO)
            
            grd_Listad(1).Rows = grd_Listad(1).Rows + 1
            grd_Listad(1).Row = grd_Listad(1).Rows - 1
            grd_Listad(1).Col = 0
            grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(1).Text = "Razón Social Promotor"
            
            grd_Listad(1).Col = 1
            grd_Listad(1).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(1).Text = moddat_gf_Consulta_RazSoc(g_rst_Princi!SOLINM_TIPDOC_PRO, g_rst_Princi!SOLINM_NUMDOC_PRO)
            
            'Constructor
            grd_Listad(1).Rows = grd_Listad(1).Rows + 2
            grd_Listad(1).Row = grd_Listad(1).Rows - 1
            grd_Listad(1).Col = 0
            grd_Listad(1).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(1).Text = "Doc. Ident. Constructor"
            
            grd_Listad(1).Col = 1
            grd_Listad(1).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(1).Text = CStr(g_rst_Princi!SOLINM_TIPDOC_CON) & "-" & Trim(g_rst_Princi!SOLINM_NUMDOC_CON)
            
            grd_Listad(1).Rows = grd_Listad(1).Rows + 1
            grd_Listad(1).Row = grd_Listad(1).Rows - 1
            grd_Listad(1).Col = 0
            grd_Listad(1).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(1).Text = "Razón Social Constructor"
            
            grd_Listad(1).Col = 1
            grd_Listad(1).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(1).Text = moddat_gf_Consulta_RazSoc(g_rst_Princi!SOLINM_TIPDOC_CON, g_rst_Princi!SOLINM_NUMDOC_CON)
         End If
      End If
      
      grd_Listad(1).Redraw = True
      
      Call gs_UbiIniGrid(grd_Listad(1))
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_DatCre()
   'Buscando Información del Crédito
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
   
   'Cargando en Grid
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
      grd_Listad(2).Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_APOSOL, 12, 2)
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
   
   Call gs_UbiIniGrid(grd_Listad(2))
End Sub

Private Sub fs_EvaCre()
   Call gs_LimpiaGrid(grd_Listad(3))
   
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = '" & moddat_g_str_NumSol & "' "

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
      
      grd_Listad(3).Rows = grd_Listad(3).Rows + 1
      grd_Listad(3).Row = grd_Listad(3).Rows - 1
      grd_Listad(3).Col = 0
      grd_Listad(3).Text = "Cuota Aceptada por Cliente (M. Prest.)"
      
      grd_Listad(3).Col = 1
      grd_Listad(3).CellFontName = "Lucida Console"
      grd_Listad(3).CellFontSize = 8
      grd_Listad(3).Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_CUOAPR_MPR, 12, 2)
      
      If moddat_g_int_TipMon <> 1 Then
         grd_Listad(3).Rows = grd_Listad(3).Rows + 1
         grd_Listad(3).Row = grd_Listad(3).Rows - 1
         grd_Listad(3).Col = 0
         grd_Listad(3).Text = "Tipo de Cambio"
         
         grd_Listad(3).Col = 1
         grd_Listad(3).CellFontName = "Lucida Console"
         grd_Listad(3).CellFontSize = 8
         grd_Listad(3).Text = "S/. " & gf_FormatoNumero(g_rst_Princi!SOLMAE_TCAMPR_APR, 12, 4)
      End If
      
      grd_Listad(3).Rows = grd_Listad(3).Rows + 2
      grd_Listad(3).Row = grd_Listad(3).Rows - 1
      grd_Listad(3).Col = 0
      grd_Listad(3).Text = "Cuota Aprobada"
      
      grd_Listad(3).Col = 1
      grd_Listad(3).CellFontName = "Lucida Console"
      grd_Listad(3).CellFontSize = 8
      grd_Listad(3).Text = "S/. " & gf_FormatoNumero(g_rst_Princi!SOLMAE_CUOMEN_SOL, 12, 2)
      
      grd_Listad(3).Rows = grd_Listad(3).Rows + 1
      grd_Listad(3).Row = grd_Listad(3).Rows - 1
      grd_Listad(3).Col = 0
      grd_Listad(3).Text = "Cuota Aprobada (M. Prest.)"
      
      grd_Listad(3).Col = 1
      grd_Listad(3).CellFontName = "Lucida Console"
      grd_Listad(3).CellFontSize = 8
      grd_Listad(3).Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!SOLMAE_CUOMEN_MPR, 12, 2)
   
      If moddat_g_int_TipMon <> 1 Then
         grd_Listad(3).Rows = grd_Listad(3).Rows + 1
         grd_Listad(3).Row = grd_Listad(3).Rows - 1
         grd_Listad(3).Col = 0
         grd_Listad(3).Text = "Tipo de Cambio"
         
         grd_Listad(3).Col = 1
         grd_Listad(3).CellFontName = "Lucida Console"
         grd_Listad(3).CellFontSize = 8
         grd_Listad(3).Text = "S/. " & gf_FormatoNumero(g_rst_Princi!SOLMAE_TIPCAM, 12, 4)
      End If
   
      grd_Listad(3).Rows = grd_Listad(3).Rows + 2
      grd_Listad(3).Row = grd_Listad(3).Rows - 1
      grd_Listad(3).Col = 0
      grd_Listad(3).Text = "Total Ingreso Líquido"
      
      grd_Listad(3).Col = 1
      grd_Listad(3).CellFontName = "Lucida Console"
      grd_Listad(3).CellFontSize = 8
      grd_Listad(3).Text = "S/. " & gf_FormatoNumero(g_rst_Princi!SOLMAE_INGNET, 12, 2)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   Call gs_UbiIniGrid(grd_Listad(3))
End Sub

Private Sub fs_DatTas()
   Call gs_LimpiaGrid(grd_Listad(4))
   
   g_str_Parame = "SELECT * FROM TRA_EVATAS WHERE "
   g_str_Parame = g_str_Parame & "EVATAS_NUMSOL = '" & moddat_g_str_NumSol & "' "

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
   Call gs_LimpiaGrid(grd_Listad(5))
   
   g_str_Parame = "SELECT * FROM TRA_EVASEG WHERE "
   g_str_Parame = g_str_Parame & "EVASEG_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   g_str_Parame = "SELECT * FROM TRA_POLIZA WHERE "
   g_str_Parame = g_str_Parame & "POLIZA_NUMSOL = '" & moddat_g_str_NumSol & "' "

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
   Call gs_LimpiaGrid(grd_Listad(6))

   g_str_Parame = "SELECT * FROM TRA_EVALEG WHERE "
   g_str_Parame = g_str_Parame & "EVALEG_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      txt_InfLeg.Text = Trim(g_rst_Princi!EVALEG_INFLG1 & "") & Trim(g_rst_Princi!EVALEG_INFLG2 & "") & Trim(g_rst_Princi!EVALEG_INFLG3 & "") & Trim(g_rst_Princi!EVALEG_INFLG4 & "")
      txt_ComCre.Text = "Fecha de Comité de Créditos: " & gf_FormatoFecha(CStr(g_rst_Princi!EVALEG_FECCOM)) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Trim(g_rst_Princi!EVALEG_OBSCOM & "")
      
      grd_Listad(6).Rows = grd_Listad(6).Rows + 1
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
      grd_Listad(6).Text = moddat_gf_Consulta_ParDes("509", g_rst_Princi!EVALEG_CODNOT)
   
      grd_Listad(6).Rows = grd_Listad(6).Rows + 1
      grd_Listad(6).Row = grd_Listad(6).Rows - 1
      grd_Listad(6).Col = 0
      grd_Listad(6).Text = "Representante Legal 1"
      
      grd_Listad(6).Col = 1
      grd_Listad(6).Text = moddat_gf_Consulta_ParDes("512", g_rst_Princi!EVALEG_REPLG1)
   
      grd_Listad(6).Rows = grd_Listad(6).Rows + 1
      grd_Listad(6).Row = grd_Listad(6).Rows - 1
      grd_Listad(6).Col = 0
      grd_Listad(6).Text = "Representante Legal 2"
      
      grd_Listad(6).Col = 1
      grd_Listad(6).Text = moddat_gf_Consulta_ParDes("512", g_rst_Princi!EVALEG_REPLG2)
      
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
            Case 1
               grd_Listad(6).Text = grd_Listad(6).Text & " NRO. " & Trim(g_rst_Princi!EVALEG_NUMPAR_INM & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!EVALEG_NUMAPA_INM & "")
               
            Case 2
               grd_Listad(6).Text = grd_Listad(6).Text & " NRO. " & Trim(g_rst_Princi!EVALEG_NUMFIC_INM & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!EVALEG_NUMAFI_INM & "")
               
            Case 3
               grd_Listad(6).Text = grd_Listad(6).Text & " (" & Trim(g_rst_Princi!EVALEG_NUMTOM_INM & "") & " / " & Trim(g_rst_Princi!EVALEG_NUMFOJ_INM & "") & " / " & Trim(g_rst_Princi!EVALEG_NUMLIB_INM & "") & ")"
               
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
               Case 1
                  grd_Listad(6).Text = grd_Listad(6).Text & " NRO. " & Trim(g_rst_Princi!EVALEG_NUMPAR_ES1 & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!EVALEG_NUMAPA_ES1 & "")
                  
               Case 2
                  grd_Listad(6).Text = grd_Listad(6).Text & " NRO. " & Trim(g_rst_Princi!EVALEG_NUMFIC_ES1 & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!EVALEG_NUMAFI_ES1 & "")
                  
               Case 3
                  grd_Listad(6).Text = grd_Listad(6).Text & " (" & Trim(g_rst_Princi!EVALEG_NUMTOM_ES1 & "") & " / " & Trim(g_rst_Princi!EVALEG_NUMFOJ_ES1 & "") & " / " & Trim(g_rst_Princi!EVALEG_NUMLIB_ES1 & "") & ")"
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
               Case 1
                  grd_Listad(6).Text = grd_Listad(6).Text & " NRO. " & Trim(g_rst_Princi!EVALEG_NUMPAR_ES2 & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!EVALEG_NUMAPA_ES2 & "")
                  
               Case 2
                  grd_Listad(6).Text = grd_Listad(6).Text & " NRO. " & Trim(g_rst_Princi!EVALEG_NUMFIC_ES2 & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!EVALEG_NUMAFI_ES2 & "")
                  
               Case 3
                  grd_Listad(6).Text = grd_Listad(6).Text & " (" & Trim(g_rst_Princi!EVALEG_NUMTOM_ES2 & "") & " / " & Trim(g_rst_Princi!EVALEG_NUMFOJ_ES2 & "") & " / " & Trim(g_rst_Princi!EVALEG_NUMLIB_ES2 & "") & ")"
                  
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
               Case 1
                  grd_Listad(6).Text = grd_Listad(6).Text & " NRO. " & Trim(g_rst_Princi!EVALEG_NUMPAR_DEP & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!EVALEG_NUMAPA_DEP & "")
                  
               Case 2
                  grd_Listad(6).Text = grd_Listad(6).Text & " NRO. " & Trim(g_rst_Princi!EVALEG_NUMFIC_DEP & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!EVALEG_NUMAFI_DEP & "")
                  
               Case 3
                  grd_Listad(6).Text = grd_Listad(6).Text & " (" & Trim(g_rst_Princi!EVALEG_NUMTOM_DEP & "") & " / " & Trim(g_rst_Princi!EVALEG_NUMFOJ_DEP & "") & " / " & Trim(g_rst_Princi!EVALEG_NUMLIB_DEP & "") & ")"
            End Select
         End If
      End If
      
      Call gs_UbiIniGrid(grd_Listad(6))
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_DatMVi()
   If moddat_g_str_CodPrd <> "001" Then
      Exit Sub
   End If

   txt_ObsMVi.Text = ""
   Call gs_LimpiaGrid(grd_Listad(7))
   
   g_str_Parame = "SELECT * FROM TRA_EVAMVI WHERE "
   g_str_Parame = g_str_Parame & "EVAMVI_NUMSOL = '" & moddat_g_str_NumSol & "' "

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
   If Not (moddat_g_str_CodPrd = "004" Or moddat_g_str_CodPrd = "003") Then
      Exit Sub
   End If
   
   txt_ObsMVi.Text = ""
   Call gs_LimpiaGrid(grd_Listad(7))
   
   g_str_Parame = "SELECT * FROM TRA_EVACOF WHERE "
   g_str_Parame = g_str_Parame & "EVACOF_NUMSOL = '" & moddat_g_str_NumSol & "' "

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
         If moddat_g_str_CodPrd = "003" Then
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

   r_str_NomCab = Mid(p_NomFil, 1, r_int_Contad) & "C" & Format(Date, "yymmdd") & ".240"
   r_str_NomDet = Mid(p_NomFil, 1, r_int_Contad) & "D" & Format(Date, "yymmdd") & ".240"

   r_int_NumFil = FreeFile
   Open r_str_NomCab For Output As r_int_NumFil

   g_str_Parame = "SELECT * FROM CRE_HIPMAE WHERE "
   g_str_Parame = g_str_Parame & "HIPMAE_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   
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

   g_str_Parame = "SELECT * FROM CRE_HIPCUO WHERE "
   g_str_Parame = g_str_Parame & "HIPCUO_NUMOPE = '" & moddat_g_str_NumOpe & "' AND "
   g_str_Parame = g_str_Parame & "HIPCUO_TIPCRO = 3 "
   g_str_Parame = g_str_Parame & "ORDER BY HIPCUO_NUMCUO ASC"
   
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

Private Sub fs_LiqDes()
   Dim r_str_Direcc     As String
   Dim r_str_Distri     As String

   Screen.MousePointer = 11
   
   'Borrando Spool Local
   moddat_g_bdt_Report.Execute "DELETE FROM RPT_CROCAB"
   moddat_g_bdt_Report.Execute "DELETE FROM RPT_CROCUO"
   DoEvents
   
   'Grabando Cabecera
   g_str_Parame = "SELECT * FROM CRE_HIPMAE WHERE "
   g_str_Parame = g_str_Parame & "HIPMAE_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      'Grabando en DAO
      moddat_g_str_CadDAO = "SELECT * FROM RPT_CROCAB WHERE CROCAB_NUMOPE = ' '"
      Set moddat_g_rst_RecDAO = moddat_g_bdt_Report.OpenRecordset(moddat_g_str_CadDAO, dbOpenDynaset)
      
      moddat_g_rst_RecDAO.AddNew
      
      moddat_g_rst_RecDAO("CROCAB_NUMOPE") = Mid(moddat_g_str_NumOpe, 1, 3) & "-" & Mid(moddat_g_str_NumOpe, 4, 2) & "-" & Mid(moddat_g_str_NumOpe, 6, 5)
      moddat_g_rst_RecDAO("CROCAB_DOCIDE") = CStr(g_rst_Princi!HIPMAE_TDOCLI) & "-" & Trim(g_rst_Princi!HIPMAE_NDOCLI & "")
      moddat_g_rst_RecDAO("CROCAB_NOMCLI") = moddat_gf_Buscar_NomCli(CStr(g_rst_Princi!HIPMAE_TDOCLI), Trim(g_rst_Princi!HIPMAE_NDOCLI))
      moddat_g_rst_RecDAO("CROCAB_MTOPRE") = g_rst_Princi!HIPMAE_MTOPRE
      moddat_g_rst_RecDAO("CROCAB_PLAPRE") = g_rst_Princi!HIPMAE_PLAANO
      moddat_g_rst_RecDAO("CROCAB_NUMCUO") = g_rst_Princi!HIPMAE_NUMCUO
      moddat_g_rst_RecDAO("CROCAB_PERGRA") = g_rst_Princi!HIPMAE_PERGRA
      moddat_g_rst_RecDAO("CROCAB_CUOEXT") = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!HIPMAE_CUOANO))
      moddat_g_rst_RecDAO("CROCAB_PRODUC") = moddat_gf_Consulta_Produc(g_rst_Princi!HIPMAE_CODPRD)
                           
      moddat_g_rst_RecDAO("CROCAB_MODALI") = ""
      If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), g_rst_Princi!HIPMAE_CODPRD, g_rst_Princi!HIPMAE_CODSUB, "003", Format(CInt(CStr(g_rst_Princi!HIPMAE_CODMOD)), "000")) Then
         moddat_g_rst_RecDAO("CROCAB_MODALI") = moddat_g_arr_Genera(1).Genera_Nombre
      End If
      
      moddat_g_rst_RecDAO("CROCAB_MONEDA") = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!HIPMAE_MONEDA))
      moddat_g_rst_RecDAO("CROCAB_TASINT") = g_rst_Princi!HIPMAE_TASINT
      
      Call moddat_gs_Consulta_DatInm(g_rst_Princi!HIPMAE_NUMSOL, r_str_Direcc, r_str_Distri)
      
      moddat_g_rst_RecDAO("CROCAB_DIRINM") = r_str_Direcc
      moddat_g_rst_RecDAO("CROCAB_DIRUBI") = r_str_Distri
      
      moddat_g_rst_RecDAO("CROCAB_NUMSOL") = Mid(g_rst_Princi!HIPMAE_NUMSOL, 1, 3) & "-" & Mid(g_rst_Princi!HIPMAE_NUMSOL, 4, 3) & "-" & Mid(g_rst_Princi!HIPMAE_NUMSOL, 7, 2) & "-" & Mid(g_rst_Princi!HIPMAE_NUMSOL, 9, 4)
      
      moddat_g_rst_RecDAO("CROCAB_EMPSEG") = moddat_gf_Consulta_ComSeg(g_rst_Princi!HIPMAE_SEGPRE & "")
      moddat_g_rst_RecDAO("CROCAB_TIPSEG") = moddat_gf_Consulta_TipSeg(g_rst_Princi!HIPMAE_SEGPRE, g_rst_Princi!HIPMAE_TIPSEG)
      
      moddat_g_rst_RecDAO("CROCAB_TSGDES") = g_rst_Princi!HIPMAE_FOIPRE
      moddat_g_rst_RecDAO("CROCAB_TSGVIV") = g_rst_Princi!HIPMAE_FOIVIV
      
      moddat_g_rst_RecDAO("CROCAB_FECDES") = gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECDES))
      
      
      'Obteniendo Información del Desembolso
      g_str_Parame = "SELECT * FROM CRE_HIPDES WHERE "
      g_str_Parame = g_str_Parame & "HIPDES_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
          Exit Sub
      End If
      
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         g_rst_Genera.MoveFirst
      
         moddat_g_rst_RecDAO("CROCAB_NUMCHQ") = Trim(g_rst_Genera!HIPDES_CHECGO & "")
         moddat_g_rst_RecDAO("CROCAB_NOMBCO") = moddat_gf_Consulta_ParDes("516", g_rst_Genera!HIPDES_BANCGO & "") & " (" & Trim(g_rst_Genera!HIPDES_CTACGO & "") & ")"
         moddat_g_rst_RecDAO("CROCAB_MONDES") = moddat_gf_Consulta_ParDes("204", g_rst_Genera!HIPDES_TIPMON)
         moddat_g_rst_RecDAO("CROCAB_MTODES") = g_rst_Genera!HIPDES_IMPORT
      
         If g_rst_Genera!HIPDES_FLGFIA = 1 Then
            If Len(Trim(g_rst_Genera!HIPDES_BANFIA & "")) > 0 Then
               moddat_g_rst_RecDAO("CROCAB_BANFIA") = moddat_gf_Consulta_ParDes("505", g_rst_Genera!HIPDES_BANFIA & "")
               moddat_g_rst_RecDAO("CROCAB_NUMFIA") = Trim(g_rst_Genera!HIPDES_NUMFIA & "")
               moddat_g_rst_RecDAO("CROCAB_EMIFIA") = gf_FormatoFecha(CStr(g_rst_Genera!HIPDES_EMIFIA))
               moddat_g_rst_RecDAO("CROCAB_VCTFIA") = gf_FormatoFecha(CStr(g_rst_Genera!HIPDES_VCTFIA))
               moddat_g_rst_RecDAO("CROCAB_MONFIA") = moddat_gf_Consulta_ParDes("204", g_rst_Genera!HIPDES_MONFIA)
               moddat_g_rst_RecDAO("CROCAB_MTOFIA") = g_rst_Genera!HIPDES_IMPFIA
            End If
         End If
      
         moddat_g_rst_RecDAO("CROCAB_OBSERV") = Trim(g_rst_Genera!HIPDES_OBSERV & "") & " "
      End If
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      
      moddat_g_rst_RecDAO.Update
      moddat_g_rst_RecDAO.Close
      
      DoEvents
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Screen.MousePointer = 0
   
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_LIQDES_01.RPT"
   crp_Imprim.Action = 1
End Sub

Private Sub fs_HojRes()
   Dim r_str_Direcc     As String
   Dim r_str_Distri     As String
   Dim r_int_Indice     As String
   Dim r_int_Modali     As Integer

   Screen.MousePointer = 11
   
   'Borrando Spool Local
   moddat_g_bdt_Report.Execute "DELETE FROM RPT_CROCAB"
   moddat_g_bdt_Report.Execute "DELETE FROM RPT_CROCUO"
   DoEvents
   
   'Grabando Cabecera
   g_str_Parame = "SELECT * FROM CRE_HIPMAE WHERE "
   g_str_Parame = g_str_Parame & "HIPMAE_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      'Grabando en DAO
      moddat_g_str_CadDAO = "SELECT * FROM RPT_CROCAB WHERE CROCAB_NUMOPE = ' '"
      Set moddat_g_rst_RecDAO = moddat_g_bdt_Report.OpenRecordset(moddat_g_str_CadDAO, dbOpenDynaset)
      
      moddat_g_rst_RecDAO.AddNew
      
      moddat_g_rst_RecDAO("CROCAB_NUMOPE") = Mid(moddat_g_str_NumOpe, 1, 3) & "-" & Mid(moddat_g_str_NumOpe, 4, 2) & "-" & Mid(moddat_g_str_NumOpe, 6, 5)
      moddat_g_rst_RecDAO("CROCAB_DOCIDE") = CStr(g_rst_Princi!HIPMAE_TDOCLI) & "-" & Trim(g_rst_Princi!HIPMAE_NDOCLI & "")
      moddat_g_rst_RecDAO("CROCAB_NOMCLI") = moddat_gf_Buscar_NomCli(CStr(g_rst_Princi!HIPMAE_TDOCLI), Trim(g_rst_Princi!HIPMAE_NDOCLI))
      
      If g_rst_Princi!HIPMAE_TDOCYG > 0 Then
         moddat_g_rst_RecDAO("CROCAB_DOICYG") = CStr(g_rst_Princi!HIPMAE_TDOCYG) & "-" & Trim(g_rst_Princi!HIPMAE_NDOCYG & "")
         moddat_g_rst_RecDAO("CROCAB_NOMCYG") = moddat_gf_Buscar_NomCli(CStr(g_rst_Princi!HIPMAE_TDOCYG), Trim(g_rst_Princi!HIPMAE_NDOCYG))
      End If
      
      moddat_g_rst_RecDAO("CROCAB_MTOPRE") = g_rst_Princi!HIPMAE_MTOPRE
      moddat_g_rst_RecDAO("CROCAB_PLAPRE") = g_rst_Princi!HIPMAE_PLAANO
      moddat_g_rst_RecDAO("CROCAB_NUMCUO") = g_rst_Princi!HIPMAE_NUMCUO
      moddat_g_rst_RecDAO("CROCAB_PERGRA") = g_rst_Princi!HIPMAE_PERGRA
      moddat_g_rst_RecDAO("CROCAB_CUOEXT") = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!HIPMAE_CUOANO))
      moddat_g_rst_RecDAO("CROCAB_PRODUC") = moddat_gf_Consulta_Produc(g_rst_Princi!HIPMAE_CODPRD)
                           
      moddat_g_rst_RecDAO("CROCAB_MODALI") = ""
      If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), g_rst_Princi!HIPMAE_CODPRD, g_rst_Princi!HIPMAE_CODSUB, "003", Format(CInt(CStr(g_rst_Princi!HIPMAE_CODMOD)), "000")) Then
         moddat_g_rst_RecDAO("CROCAB_MODALI") = moddat_g_arr_Genera(1).Genera_Nombre
         
         r_int_Modali = CInt(g_rst_Princi!HIPMAE_CODMOD)
      End If
      
      moddat_g_rst_RecDAO("CROCAB_MONEDA") = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!HIPMAE_MONEDA))
      moddat_g_rst_RecDAO("CROCAB_SIMMON") = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!HIPMAE_MONEDA))
      
      If g_rst_Princi!HIPMAE_MONEDA = 1 Then
         moddat_g_rst_RecDAO("CROCAB_VALVIV") = g_rst_Princi!HIPMAE_CVTSOL
         moddat_g_rst_RecDAO("CROCAB_APOPRO") = g_rst_Princi!HIPMAE_APOSOL
      Else
         moddat_g_rst_RecDAO("CROCAB_VALVIV") = g_rst_Princi!HIPMAE_CVTDOL
         moddat_g_rst_RecDAO("CROCAB_APOPRO") = g_rst_Princi!HIPMAE_APODOL
      End If
      
      DoEvents
      moddat_g_rst_RecDAO("CROCAB_TASINT") = g_rst_Princi!HIPMAE_TASINT
      
      Call moddat_gs_Consulta_DatInm(g_rst_Princi!HIPMAE_NUMSOL, r_str_Direcc, r_str_Distri)
      
      moddat_g_rst_RecDAO("CROCAB_DIRINM") = r_str_Direcc
      moddat_g_rst_RecDAO("CROCAB_DIRUBI") = r_str_Distri
      
      moddat_g_rst_RecDAO("CROCAB_NUMSOL") = Mid(g_rst_Princi!HIPMAE_NUMSOL, 1, 3) & "-" & Mid(g_rst_Princi!HIPMAE_NUMSOL, 4, 3) & "-" & Mid(g_rst_Princi!HIPMAE_NUMSOL, 7, 2) & "-" & Mid(g_rst_Princi!HIPMAE_NUMSOL, 9, 4)
      
      moddat_g_rst_RecDAO("CROCAB_EMPSEG") = moddat_gf_Consulta_ComSeg(g_rst_Princi!HIPMAE_SEGPRE & "")
      moddat_g_rst_RecDAO("CROCAB_TIPSEG") = moddat_gf_Consulta_TipSeg(g_rst_Princi!HIPMAE_SEGPRE, g_rst_Princi!HIPMAE_TIPSEG)
      
      moddat_g_rst_RecDAO("CROCAB_TSGDES") = g_rst_Princi!HIPMAE_FOIPRE
      moddat_g_rst_RecDAO("CROCAB_TSGVIV") = g_rst_Princi!HIPMAE_FOIVIV
      
      moddat_g_rst_RecDAO("CROCAB_FECDES") = gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECDES))
      
      DoEvents
      
      'Datos de la Hipoteca en Evaluación Legal
      g_str_Parame = "SELECT * FROM TRA_EVALEG WHERE "
      g_str_Parame = g_str_Parame & "EVALEG_NUMSOL = '" & g_rst_Princi!HIPMAE_NUMSOL & "' "

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
         Exit Sub
      End If
   
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         g_rst_Princi.MoveFirst
      
         moddat_g_rst_RecDAO("CROCAB_MONGAR") = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Genera!EVALEG_MONHIP))
         moddat_g_rst_RecDAO("CROCAB_SIMGAR") = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Genera!EVALEG_MONHIP))
         moddat_g_rst_RecDAO("CROCAB_MTOGAR") = g_rst_Genera!EVALEG_MTOHIP
      End If
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      
      'Interes Compensatorio
      moddat_g_rst_RecDAO("CROCAB_INTCOM") = 0
      If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), moddat_g_str_CodPrd, moddat_g_str_CodSub, "002", "301") Then
         moddat_g_rst_RecDAO("CROCAB_INTCOM") = moddat_g_arr_Genera(1).Genera_Cantid
      End If
      
      'Interes Moratorio
      moddat_g_rst_RecDAO("CROCAB_INTMOR") = 0
      If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), moddat_g_str_CodPrd, moddat_g_str_CodSub, "002", "201") Then
         moddat_g_rst_RecDAO("CROCAB_INTMOR") = moddat_g_arr_Genera(1).Genera_Cantid
      End If
      
      'ITF
      moddat_g_rst_RecDAO("CROCAB_PORITF") = opecaj_gf_Consulta_ITF(Format(CDate(moddat_g_str_FecSis), "yyyymmdd"), 1)
      
      DoEvents
      
      'Pólizas de Seguro
      g_str_Parame = "SELECT * FROM TRA_POLIZA WHERE "
      g_str_Parame = g_str_Parame & "POLIZA_NUMSOL = '" & g_rst_Princi!HIPMAE_NUMSOL & "' "

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
          Exit Sub
      End If
      
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         g_rst_Genera.MoveFirst
      
         moddat_g_rst_RecDAO("CROCAB_POLDES") = Trim(g_rst_Genera!POLIZA_NUMDES & "") & IIf(Len(Trim(g_rst_Genera!POLIZA_NUMCYG & "")) > 0, " / " & Trim(g_rst_Genera!POLIZA_NUMCYG & ""), "")
         moddat_g_rst_RecDAO("CROCAB_POLVIV") = Trim(g_rst_Genera!POLIZA_NUMVIV & "")
      End If
   
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      
      'Gastos de Cierre
      moddat_g_rst_RecDAO("CROCAB_MONTAS") = l_str_MonTas
      moddat_g_rst_RecDAO("CROCAB_IMPTAS") = l_dbl_ImpTas
      moddat_g_rst_RecDAO("CROCAB_MONNOT") = l_str_MonNot
      moddat_g_rst_RecDAO("CROCAB_IMPNOT") = l_dbl_ImpNot
      moddat_g_rst_RecDAO("CROCAB_MONEST") = l_str_MonEst
      moddat_g_rst_RecDAO("CROCAB_IMPEST") = l_dbl_ImpEst
      moddat_g_rst_RecDAO("CROCAB_MONEVA") = l_str_MonEva
      moddat_g_rst_RecDAO("CROCAB_IMPEVA") = l_dbl_ImpEva
      moddat_g_rst_RecDAO("CROCAB_MONADM") = l_str_MonAdm
      moddat_g_rst_RecDAO("CROCAB_IMPADM") = l_dbl_ImpAdm
      moddat_g_rst_RecDAO("CROCAB_MONRED") = l_str_MonRed
      moddat_g_rst_RecDAO("CROCAB_IMPRED") = l_dbl_ImpRed
      moddat_g_rst_RecDAO("CROCAB_MONBLQ") = l_str_MonBlq
      moddat_g_rst_RecDAO("CROCAB_IMPBLQ") = l_dbl_ImpBlq
      
      'Otras Comisiones - Prepagos
      moddat_g_rst_RecDAO("CROCAB_COMPRE") = 0
      If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), g_rst_Princi!HIPMAE_CODPRD, g_rst_Princi!HIPMAE_CODSUB, "002", "501") Then
         moddat_g_rst_RecDAO("CROCAB_COMPRE") = moddat_g_arr_Genera(1).Genera_Cantid
      End If
      
      'Otras Comisiones - Levantamiento de Hipoteca
      moddat_g_rst_RecDAO("CROCAB_LEVHIP") = 0
      If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), g_rst_Princi!HIPMAE_CODPRD, g_rst_Princi!HIPMAE_CODSUB, "002", "701") Then
         moddat_g_rst_RecDAO("CROCAB_LEVHIP") = moddat_g_arr_Genera(1).Genera_Cantid
      End If
      
      'Otras Comisiones - Cambio de Fecha, Tasa de Interes, Moneda o Cuota
      moddat_g_rst_RecDAO("CROCAB_CAMFEC") = 0
      If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), g_rst_Princi!HIPMAE_CODPRD, g_rst_Princi!HIPMAE_CODSUB, "002", "702") Then
         moddat_g_rst_RecDAO("CROCAB_CAMFEC") = moddat_g_arr_Genera(1).Genera_Cantid
      End If
      
      'Otras Comisiones - Portes
      moddat_g_rst_RecDAO("CROCAB_PORTES") = 0
      If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), g_rst_Princi!HIPMAE_CODPRD, g_rst_Princi!HIPMAE_CODSUB, "002", "401") Then
         moddat_g_rst_RecDAO("CROCAB_PORTES") = moddat_g_arr_Genera(1).Genera_Cantid
      End If
      
      'Otras Comisiones - Cobranza Judicial
      moddat_g_rst_RecDAO("CROCAB_COBJUD") = 0
      If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), g_rst_Princi!HIPMAE_CODPRD, g_rst_Princi!HIPMAE_CODSUB, "002", "703") Then
         moddat_g_rst_RecDAO("CROCAB_COBJUD") = moddat_g_arr_Genera(1).Genera_Cantid
      End If
      
      'Otras Comisiones - Caducidad del Servicio MiVivienda
      moddat_g_rst_RecDAO("CROCAB_CANCRE") = 0
      If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), g_rst_Princi!HIPMAE_CODPRD, g_rst_Princi!HIPMAE_CODSUB, "002", "901") Then
         moddat_g_rst_RecDAO("CROCAB_CANCRE") = moddat_g_arr_Genera(1).Genera_Cantid
      End If
      
      'Gastos de Cobranzas
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT * FROM OPE_GASCOB WHERE "
      g_str_Parame = g_str_Parame & "GASCOB_CODPRD = '" & g_rst_Princi!HIPMAE_CODPRD & "' AND "
      g_str_Parame = g_str_Parame & "GASCOB_CODSUB = '" & g_rst_Princi!HIPMAE_CODSUB & "' AND "
      g_str_Parame = g_str_Parame & "GASCOB_IMPORT > 0 "
      g_str_Parame = g_str_Parame & "ORDER BY GASCOB_DIAINI ASC"

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
          Exit Sub
      End If
   
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         g_rst_Genera.MoveFirst
         
         r_int_Indice = 1
         Do While Not g_rst_Genera.EOF
            moddat_g_rst_RecDAO("CROCAB_GASDI" & Format(r_int_Indice)) = g_rst_Genera!GASCOB_DIAINI
            moddat_g_rst_RecDAO("CROCAB_GASDF" & Format(r_int_Indice)) = g_rst_Genera!GASCOB_DIAFIN
            moddat_g_rst_RecDAO("CROCAB_GASIM" & Format(r_int_Indice)) = g_rst_Genera!GASCOB_IMPORT
            
            r_int_Indice = r_int_Indice + 1
            
            DoEvents
            g_rst_Genera.MoveNext
         Loop
   
         g_rst_Genera.Close
         Set g_rst_Genera = Nothing
      End If
      
      moddat_g_rst_RecDAO("CROCAB_OBSERV") = " "
      moddat_g_rst_RecDAO.Update
      moddat_g_rst_RecDAO.Close
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Screen.MousePointer = 0
   
   Select Case moddat_g_str_CodPrd
      Case "002"
         If r_int_Modali = 1 Then
            crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_HOJRES_01.RPT"
         Else
            crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_HOJRES_02.RPT"
         End If
      Case "001"
         'If r_int_Modali <> 1 Then
            crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_HOJRES_04.RPT"
         'End If
         
      Case "004"
         crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_HOJRES_06.RPT"
   End Select
   
   crp_Imprim.Action = 1
End Sub

Private Sub fs_Cronog_MiCasita()
   Screen.MousePointer = 11
   
   'Borrando Spool Local
   moddat_g_bdt_Report.Execute "DELETE FROM RPT_CROCAB"
   moddat_g_bdt_Report.Execute "DELETE FROM RPT_CROCUO"
   DoEvents
   
   'Grabando Cabecera
   g_str_Parame = "SELECT * FROM CRE_HIPMAE WHERE "
   g_str_Parame = g_str_Parame & "HIPMAE_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      'Grabando en DAO
      moddat_g_str_CadDAO = "SELECT * FROM RPT_CROCAB WHERE CROCAB_NUMOPE = ' '"
      Set moddat_g_rst_RecDAO = moddat_g_bdt_Report.OpenRecordset(moddat_g_str_CadDAO, dbOpenDynaset)
      
      moddat_g_rst_RecDAO.AddNew
      
      moddat_g_rst_RecDAO("CROCAB_NUMOPE") = Mid(moddat_g_str_NumOpe, 1, 3) & "-" & Mid(moddat_g_str_NumOpe, 4, 2) & "-" & Mid(moddat_g_str_NumOpe, 6, 5)
      moddat_g_rst_RecDAO("CROCAB_DOCIDE") = CStr(g_rst_Princi!HIPMAE_TDOCLI) & "-" & Trim(g_rst_Princi!HIPMAE_NDOCLI & "")
      moddat_g_rst_RecDAO("CROCAB_NOMCLI") = moddat_gf_Buscar_NomCli(CStr(g_rst_Princi!HIPMAE_TDOCLI), Trim(g_rst_Princi!HIPMAE_NDOCLI))
      moddat_g_rst_RecDAO("CROCAB_MTOPRE") = g_rst_Princi!HIPMAE_MTOPRE
      moddat_g_rst_RecDAO("CROCAB_PLAPRE") = g_rst_Princi!HIPMAE_PLAANO
      moddat_g_rst_RecDAO("CROCAB_NUMCUO") = g_rst_Princi!HIPMAE_NUMCUO
      moddat_g_rst_RecDAO("CROCAB_PERGRA") = g_rst_Princi!HIPMAE_PERGRA
      moddat_g_rst_RecDAO("CROCAB_CUOEXT") = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!HIPMAE_CUOANO))
      moddat_g_rst_RecDAO("CROCAB_PRODUC") = moddat_gf_Consulta_Produc(g_rst_Princi!HIPMAE_CODPRD)
                           
      If g_rst_Princi!HIPMAE_PERGRA > 0 Then
         moddat_g_rst_RecDAO("CROCAB_INTGRA") = g_rst_Princi!HIPMAE_INTCAP
      End If
                           
      moddat_g_rst_RecDAO("CROCAB_MODALI") = ""
      If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), g_rst_Princi!HIPMAE_CODPRD, g_rst_Princi!HIPMAE_CODSUB, "003", Format(CInt(CStr(g_rst_Princi!HIPMAE_CODMOD)), "000")) Then
         moddat_g_rst_RecDAO("CROCAB_MODALI") = moddat_g_arr_Genera(1).Genera_Nombre
      End If
      
      moddat_g_rst_RecDAO("CROCAB_MONEDA") = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!HIPMAE_MONEDA))
      moddat_g_rst_RecDAO("CROCAB_TASINT") = g_rst_Princi!HIPMAE_TASINT
      moddat_g_rst_RecDAO("CROCAB_FECDES") = gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECDES))
      
      moddat_g_rst_RecDAO.Update
      moddat_g_rst_RecDAO.Close
      
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   
   'Grabando Detalle
   g_str_Parame = "SELECT * FROM CRE_HIPCUO WHERE "
   g_str_Parame = g_str_Parame & "HIPCUO_NUMOPE = '" & moddat_g_str_NumOpe & "' AND "
   g_str_Parame = g_str_Parame & "HIPCUO_TIPCRO = 1 "
   g_str_Parame = g_str_Parame & "ORDER BY HIPCUO_NUMCUO ASC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         'Grabando en DAO
         moddat_g_str_CadDAO = "SELECT * FROM RPT_CROCUO WHERE CROCUO_NUMOPE = ' '"
         Set moddat_g_rst_RecDAO = moddat_g_bdt_Report.OpenRecordset(moddat_g_str_CadDAO, dbOpenDynaset)
         
         moddat_g_rst_RecDAO.AddNew
         
         moddat_g_rst_RecDAO("CROCUO_NUMOPE") = Mid(moddat_g_str_NumOpe, 1, 3) & "-" & Mid(moddat_g_str_NumOpe, 4, 2) & "-" & Mid(moddat_g_str_NumOpe, 6, 5)
         moddat_g_rst_RecDAO("CROCUO_NUMCUO") = Format(g_rst_Princi!HIPCUO_NUMCUO, "000")
         moddat_g_rst_RecDAO("CROCUO_FECVCT") = gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECVCT))
         moddat_g_rst_RecDAO("CROCUO_CAPITA") = g_rst_Princi!HIPCUO_CAPITA
         moddat_g_rst_RecDAO("CROCUO_INTERE") = g_rst_Princi!HIPCUO_INTERE
         moddat_g_rst_RecDAO("CROCUO_SEGPRE") = g_rst_Princi!HIPCUO_DESORG
         moddat_g_rst_RecDAO("CROCUO_SEGVIV") = g_rst_Princi!HIPCUO_VIVORG
         moddat_g_rst_RecDAO("CROCUO_PORTES") = g_rst_Princi!HIPCUO_OTRORG
         moddat_g_rst_RecDAO("CROCUO_TOTCUO") = g_rst_Princi!HIPCUO_CAPITA + g_rst_Princi!HIPCUO_INTERE + g_rst_Princi!HIPCUO_DESORG + g_rst_Princi!HIPCUO_VIVORG + g_rst_Princi!HIPCUO_OTRORG
         moddat_g_rst_RecDAO("CROCUO_SALCAP") = g_rst_Princi!HIPCUO_SALCAP
                              
         moddat_g_rst_RecDAO.Update
         moddat_g_rst_RecDAO.Close
      
         DoEvents
         g_rst_Princi.MoveNext
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Screen.MousePointer = 0
   
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_CROPAG_01.RPT"
   crp_Imprim.Action = 1
End Sub

Private Sub fs_Cronog_Mivivienda_NCoCli()
   Screen.MousePointer = 11
   
   'Borrando Spool Local
   moddat_g_bdt_Report.Execute "DELETE FROM RPT_CROCAB"
   moddat_g_bdt_Report.Execute "DELETE FROM RPT_CROCUO"
   DoEvents
   
   'Grabando Cabecera
   g_str_Parame = "SELECT * FROM CRE_HIPMAE WHERE "
   g_str_Parame = g_str_Parame & "HIPMAE_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      'Grabando en DAO
      moddat_g_str_CadDAO = "SELECT * FROM RPT_CROCAB WHERE CROCAB_NUMOPE = ' '"
      Set moddat_g_rst_RecDAO = moddat_g_bdt_Report.OpenRecordset(moddat_g_str_CadDAO, dbOpenDynaset)
      
      moddat_g_rst_RecDAO.AddNew
      
      moddat_g_rst_RecDAO("CROCAB_NUMOPE") = Mid(moddat_g_str_NumOpe, 1, 3) & "-" & Mid(moddat_g_str_NumOpe, 4, 2) & "-" & Mid(moddat_g_str_NumOpe, 6, 5)
      moddat_g_rst_RecDAO("CROCAB_DOCIDE") = CStr(g_rst_Princi!HIPMAE_TDOCLI) & "-" & Trim(g_rst_Princi!HIPMAE_NDOCLI & "")
      moddat_g_rst_RecDAO("CROCAB_NOMCLI") = moddat_gf_Buscar_NomCli(CStr(g_rst_Princi!HIPMAE_TDOCLI), Trim(g_rst_Princi!HIPMAE_NDOCLI))
      moddat_g_rst_RecDAO("CROCAB_MTOPRE") = g_rst_Princi!HIPMAE_MTOPRE
      moddat_g_rst_RecDAO("CROCAB_PLAPRE") = g_rst_Princi!HIPMAE_PLAANO
      moddat_g_rst_RecDAO("CROCAB_NUMCUO") = g_rst_Princi!HIPMAE_NUMCUO
      moddat_g_rst_RecDAO("CROCAB_PERGRA") = g_rst_Princi!HIPMAE_PERGRA
      moddat_g_rst_RecDAO("CROCAB_CUOEXT") = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!HIPMAE_CUOANO))
      moddat_g_rst_RecDAO("CROCAB_PRODUC") = moddat_gf_Consulta_Produc(g_rst_Princi!HIPMAE_CODPRD)
                           
      If g_rst_Princi!HIPMAE_PERGRA > 0 Then
         moddat_g_rst_RecDAO("CROCAB_INTGRA") = g_rst_Princi!HIPMAE_INTCAP
      Else
         moddat_g_rst_RecDAO("CROCAB_INTGRA") = 0
      End If
            
      moddat_g_rst_RecDAO("CROCAB_MODALI") = ""
      If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), g_rst_Princi!HIPMAE_CODPRD, g_rst_Princi!HIPMAE_CODSUB, "003", Format(CInt(CStr(g_rst_Princi!HIPMAE_CODMOD)), "000")) Then
         moddat_g_rst_RecDAO("CROCAB_MODALI") = moddat_g_arr_Genera(1).Genera_Nombre
      End If
      
      moddat_g_rst_RecDAO("CROCAB_MONEDA") = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!HIPMAE_MONEDA))
      moddat_g_rst_RecDAO("CROCAB_TASINT") = g_rst_Princi!HIPMAE_TASINT
      moddat_g_rst_RecDAO("CROCAB_FECDES") = gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECDES))
      
      moddat_g_rst_RecDAO("CROCAB_MTONCO") = g_rst_Princi!HIPMAE_IMPNCO
      
      moddat_g_rst_RecDAO.Update
      moddat_g_rst_RecDAO.Close
      
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   
   'Grabando Detalle
   g_str_Parame = "SELECT * FROM CRE_HIPCUO WHERE "
   g_str_Parame = g_str_Parame & "HIPCUO_NUMOPE = '" & moddat_g_str_NumOpe & "' AND "
   g_str_Parame = g_str_Parame & "HIPCUO_TIPCRO = 1 "
   g_str_Parame = g_str_Parame & "ORDER BY HIPCUO_NUMCUO ASC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         'Grabando en DAO
         moddat_g_str_CadDAO = "SELECT * FROM RPT_CROCUO WHERE CROCUO_NUMOPE = ' '"
         Set moddat_g_rst_RecDAO = moddat_g_bdt_Report.OpenRecordset(moddat_g_str_CadDAO, dbOpenDynaset)
         
         moddat_g_rst_RecDAO.AddNew
         
         moddat_g_rst_RecDAO("CROCUO_NUMOPE") = Mid(moddat_g_str_NumOpe, 1, 3) & "-" & Mid(moddat_g_str_NumOpe, 4, 2) & "-" & Mid(moddat_g_str_NumOpe, 6, 5)
         moddat_g_rst_RecDAO("CROCUO_NUMCUO") = Format(g_rst_Princi!HIPCUO_NUMCUO, "000")
         moddat_g_rst_RecDAO("CROCUO_FECVCT") = gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECVCT))
         moddat_g_rst_RecDAO("CROCUO_CAPITA") = g_rst_Princi!HIPCUO_CAPITA
         moddat_g_rst_RecDAO("CROCUO_INTERE") = g_rst_Princi!HIPCUO_INTERE
         moddat_g_rst_RecDAO("CROCUO_SEGPRE") = g_rst_Princi!HIPCUO_DESORG
         moddat_g_rst_RecDAO("CROCUO_SEGVIV") = g_rst_Princi!HIPCUO_VIVORG
         moddat_g_rst_RecDAO("CROCUO_PORTES") = g_rst_Princi!HIPCUO_OTRORG
         moddat_g_rst_RecDAO("CROCUO_TOTCUO") = g_rst_Princi!HIPCUO_CAPITA + g_rst_Princi!HIPCUO_INTERE + g_rst_Princi!HIPCUO_DESORG + g_rst_Princi!HIPCUO_VIVORG + g_rst_Princi!HIPCUO_OTRORG
         moddat_g_rst_RecDAO("CROCUO_SALCAP") = g_rst_Princi!HIPCUO_SALCAP
                              
         moddat_g_rst_RecDAO.Update
         moddat_g_rst_RecDAO.Close
      
         DoEvents
         g_rst_Princi.MoveNext
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Screen.MousePointer = 0
   
   If moddat_g_str_CodPrd = "001" Then
      crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_CROPAG_02.RPT"
   ElseIf moddat_g_str_CodPrd = "004" Or moddat_g_str_CodPrd = "003" Then
      crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_CROPAG_08.RPT"
   End If
   
   crp_Imprim.Action = 1
End Sub

Private Sub fs_Cronog_Mivivienda_ConCli()
   Screen.MousePointer = 11
   
   'Borrando Spool Local
   moddat_g_bdt_Report.Execute "DELETE FROM RPT_CROCAB"
   moddat_g_bdt_Report.Execute "DELETE FROM RPT_CROCUO"
   DoEvents
   
   'Grabando Cabecera
   g_str_Parame = "SELECT * FROM CRE_HIPMAE WHERE "
   g_str_Parame = g_str_Parame & "HIPMAE_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      'Grabando en DAO
      moddat_g_str_CadDAO = "SELECT * FROM RPT_CROCAB WHERE CROCAB_NUMOPE = ' '"
      Set moddat_g_rst_RecDAO = moddat_g_bdt_Report.OpenRecordset(moddat_g_str_CadDAO, dbOpenDynaset)
      
      moddat_g_rst_RecDAO.AddNew
      
      moddat_g_rst_RecDAO("CROCAB_NUMOPE") = Mid(moddat_g_str_NumOpe, 1, 3) & "-" & Mid(moddat_g_str_NumOpe, 4, 2) & "-" & Mid(moddat_g_str_NumOpe, 6, 5)
      moddat_g_rst_RecDAO("CROCAB_DOCIDE") = CStr(g_rst_Princi!HIPMAE_TDOCLI) & "-" & Trim(g_rst_Princi!HIPMAE_NDOCLI & "")
      moddat_g_rst_RecDAO("CROCAB_NOMCLI") = moddat_gf_Buscar_NomCli(CStr(g_rst_Princi!HIPMAE_TDOCLI), Trim(g_rst_Princi!HIPMAE_NDOCLI))
      moddat_g_rst_RecDAO("CROCAB_MTOPRE") = g_rst_Princi!HIPMAE_MTOPRE
      moddat_g_rst_RecDAO("CROCAB_PLAPRE") = g_rst_Princi!HIPMAE_PLAANO
      moddat_g_rst_RecDAO("CROCAB_NUMCUO") = g_rst_Princi!HIPMAE_NUMCUO
      moddat_g_rst_RecDAO("CROCAB_PERGRA") = g_rst_Princi!HIPMAE_PERGRA
      moddat_g_rst_RecDAO("CROCAB_CUOEXT") = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!HIPMAE_CUOANO))
      moddat_g_rst_RecDAO("CROCAB_PRODUC") = moddat_gf_Consulta_Produc(g_rst_Princi!HIPMAE_CODPRD)
                           
      If g_rst_Princi!HIPMAE_PERGRA > 0 Then
         moddat_g_rst_RecDAO("CROCAB_INTGRA") = g_rst_Princi!HIPMAE_INTCAP
      Else
         moddat_g_rst_RecDAO("CROCAB_INTGRA") = 0
      End If
            
      moddat_g_rst_RecDAO("CROCAB_MODALI") = ""
      If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), g_rst_Princi!HIPMAE_CODPRD, g_rst_Princi!HIPMAE_CODSUB, "003", Format(CInt(CStr(g_rst_Princi!HIPMAE_CODMOD)), "000")) Then
         moddat_g_rst_RecDAO("CROCAB_MODALI") = moddat_g_arr_Genera(1).Genera_Nombre
      End If
      
      moddat_g_rst_RecDAO("CROCAB_MONEDA") = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!HIPMAE_MONEDA))
      moddat_g_rst_RecDAO("CROCAB_TASINT") = g_rst_Princi!HIPMAE_TASINT
      moddat_g_rst_RecDAO("CROCAB_FECDES") = gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECDES))
      
      moddat_g_rst_RecDAO("CROCAB_MTOCON") = g_rst_Princi!HIPMAE_IMPCON
      
      moddat_g_rst_RecDAO.Update
      moddat_g_rst_RecDAO.Close
      
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   
   'Grabando Detalle
   g_str_Parame = "SELECT * FROM CRE_HIPCUO WHERE "
   g_str_Parame = g_str_Parame & "HIPCUO_NUMOPE = '" & moddat_g_str_NumOpe & "' AND "
   g_str_Parame = g_str_Parame & "HIPCUO_TIPCRO = 2 "
   g_str_Parame = g_str_Parame & "ORDER BY HIPCUO_NUMCUO ASC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         'Grabando en DAO
         moddat_g_str_CadDAO = "SELECT * FROM RPT_CROCUO WHERE CROCUO_NUMOPE = ' '"
         Set moddat_g_rst_RecDAO = moddat_g_bdt_Report.OpenRecordset(moddat_g_str_CadDAO, dbOpenDynaset)
         
         moddat_g_rst_RecDAO.AddNew
         
         moddat_g_rst_RecDAO("CROCUO_NUMOPE") = Mid(moddat_g_str_NumOpe, 1, 3) & "-" & Mid(moddat_g_str_NumOpe, 4, 2) & "-" & Mid(moddat_g_str_NumOpe, 6, 5)
         moddat_g_rst_RecDAO("CROCUO_NUMCUO") = Format(g_rst_Princi!HIPCUO_NUMCUO, "000")
         moddat_g_rst_RecDAO("CROCUO_FECVCT") = gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECVCT))
         moddat_g_rst_RecDAO("CROCUO_CAPITA") = g_rst_Princi!HIPCUO_CAPITA
         moddat_g_rst_RecDAO("CROCUO_INTERE") = g_rst_Princi!HIPCUO_INTERE
         moddat_g_rst_RecDAO("CROCUO_TOTCUO") = g_rst_Princi!HIPCUO_CAPITA + g_rst_Princi!HIPCUO_INTERE
         moddat_g_rst_RecDAO("CROCUO_SALCAP") = g_rst_Princi!HIPCUO_SALCAP
                              
         moddat_g_rst_RecDAO.Update
         moddat_g_rst_RecDAO.Close
      
         DoEvents
         g_rst_Princi.MoveNext
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Screen.MousePointer = 0
   
   If moddat_g_str_CodPrd = "001" Or moddat_g_str_CodPrd = "004" Or moddat_g_str_CodPrd = "003" Then
      crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_CROPAG_03.RPT"
   End If
   
   crp_Imprim.Action = 1
End Sub

Private Sub fs_Cronog_Mivivienda_ConMVi()
   Screen.MousePointer = 11
   
   'Borrando Spool Local
   moddat_g_bdt_Report.Execute "DELETE FROM RPT_CROCAB"
   moddat_g_bdt_Report.Execute "DELETE FROM RPT_CROCUO"
   DoEvents
   
   'Grabando Cabecera
   g_str_Parame = "SELECT * FROM CRE_HIPMAE WHERE "
   g_str_Parame = g_str_Parame & "HIPMAE_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      'Grabando en DAO
      moddat_g_str_CadDAO = "SELECT * FROM RPT_CROCAB WHERE CROCAB_NUMOPE = ' '"
      Set moddat_g_rst_RecDAO = moddat_g_bdt_Report.OpenRecordset(moddat_g_str_CadDAO, dbOpenDynaset)
      
      moddat_g_rst_RecDAO.AddNew
      
      moddat_g_rst_RecDAO("CROCAB_NUMOPE") = Mid(moddat_g_str_NumOpe, 1, 3) & "-" & Mid(moddat_g_str_NumOpe, 4, 2) & "-" & Mid(moddat_g_str_NumOpe, 6, 5)
      moddat_g_rst_RecDAO("CROCAB_DOCIDE") = CStr(g_rst_Princi!HIPMAE_TDOCLI) & "-" & Trim(g_rst_Princi!HIPMAE_NDOCLI & "")
      moddat_g_rst_RecDAO("CROCAB_NOMCLI") = moddat_gf_Buscar_NomCli(CStr(g_rst_Princi!HIPMAE_TDOCLI), Trim(g_rst_Princi!HIPMAE_NDOCLI))
      moddat_g_rst_RecDAO("CROCAB_MTOPRE") = g_rst_Princi!HIPMAE_MTOPRE
      moddat_g_rst_RecDAO("CROCAB_PLAPRE") = g_rst_Princi!HIPMAE_PLAANO
      moddat_g_rst_RecDAO("CROCAB_NUMCUO") = g_rst_Princi!HIPMAE_NUMCUO
      moddat_g_rst_RecDAO("CROCAB_PERGRA") = g_rst_Princi!HIPMAE_PERGRA
      moddat_g_rst_RecDAO("CROCAB_CUOEXT") = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!HIPMAE_CUOANO))
      moddat_g_rst_RecDAO("CROCAB_PRODUC") = moddat_gf_Consulta_Produc(g_rst_Princi!HIPMAE_CODPRD)
                           
      If g_rst_Princi!HIPMAE_PERGRA > 0 Then
         moddat_g_rst_RecDAO("CROCAB_INTGRA") = g_rst_Princi!HIPMAE_INTCAP
      Else
         moddat_g_rst_RecDAO("CROCAB_INTGRA") = 0
      End If
            
      moddat_g_rst_RecDAO("CROCAB_MODALI") = ""
      If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), g_rst_Princi!HIPMAE_CODPRD, g_rst_Princi!HIPMAE_CODSUB, "003", Format(CInt(CStr(g_rst_Princi!HIPMAE_CODMOD)), "000")) Then
         moddat_g_rst_RecDAO("CROCAB_MODALI") = moddat_g_arr_Genera(1).Genera_Nombre
      End If
      
      moddat_g_rst_RecDAO("CROCAB_MONEDA") = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!HIPMAE_MONEDA))
      moddat_g_rst_RecDAO("CROCAB_TASINT") = g_rst_Princi!HIPMAE_TASMVI
      moddat_g_rst_RecDAO("CROCAB_FECDES") = gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECDES))
      
      moddat_g_rst_RecDAO("CROCAB_MTOCON") = g_rst_Princi!HIPMAE_IMPCON
      
      moddat_g_rst_RecDAO.Update
      moddat_g_rst_RecDAO.Close
      
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   
   'Grabando Detalle
   g_str_Parame = "SELECT * FROM CRE_HIPCUO WHERE "
   g_str_Parame = g_str_Parame & "HIPCUO_NUMOPE = '" & moddat_g_str_NumOpe & "' AND "
   g_str_Parame = g_str_Parame & "HIPCUO_TIPCRO = 4 "
   g_str_Parame = g_str_Parame & "ORDER BY HIPCUO_NUMCUO ASC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         'Grabando en DAO
         moddat_g_str_CadDAO = "SELECT * FROM RPT_CROCUO WHERE CROCUO_NUMOPE = ' '"
         Set moddat_g_rst_RecDAO = moddat_g_bdt_Report.OpenRecordset(moddat_g_str_CadDAO, dbOpenDynaset)
         
         moddat_g_rst_RecDAO.AddNew
         
         moddat_g_rst_RecDAO("CROCUO_NUMOPE") = Mid(moddat_g_str_NumOpe, 1, 3) & "-" & Mid(moddat_g_str_NumOpe, 4, 2) & "-" & Mid(moddat_g_str_NumOpe, 6, 5)
         moddat_g_rst_RecDAO("CROCUO_NUMCUO") = Format(g_rst_Princi!HIPCUO_NUMCUO, "000")
         moddat_g_rst_RecDAO("CROCUO_FECVCT") = gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECVCT))
         moddat_g_rst_RecDAO("CROCUO_CAPITA") = g_rst_Princi!HIPCUO_CAPITA
         moddat_g_rst_RecDAO("CROCUO_INTERE") = g_rst_Princi!HIPCUO_INTERE
         moddat_g_rst_RecDAO("CROCUO_COMCOF") = g_rst_Princi!HIPCUO_COMCOF
         moddat_g_rst_RecDAO("CROCUO_TOTCUO") = g_rst_Princi!HIPCUO_CAPITA + g_rst_Princi!HIPCUO_INTERE + g_rst_Princi!HIPCUO_COMCOF
         moddat_g_rst_RecDAO("CROCUO_SALCAP") = g_rst_Princi!HIPCUO_SALCAP
                              
         moddat_g_rst_RecDAO.Update
         moddat_g_rst_RecDAO.Close
      
         DoEvents
         g_rst_Princi.MoveNext
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Screen.MousePointer = 0
   
   Select Case moddat_g_str_CodPrd
      Case "001": crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_CROPAG_04.RPT"
      Case "003": crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_CROPAG_04.RPT"
      Case "004": crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_CROPAG_05.RPT"
   End Select
   
   crp_Imprim.Action = 1
End Sub

Private Sub fs_Cronog_Mivivienda_NCoMVi()
   Screen.MousePointer = 11
   
   'Borrando Spool Local
   moddat_g_bdt_Report.Execute "DELETE FROM RPT_CROCAB"
   moddat_g_bdt_Report.Execute "DELETE FROM RPT_CROCUO"
   DoEvents
   
   'Grabando Cabecera
   g_str_Parame = "SELECT * FROM CRE_HIPMAE WHERE "
   g_str_Parame = g_str_Parame & "HIPMAE_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      'Grabando en DAO
      moddat_g_str_CadDAO = "SELECT * FROM RPT_CROCAB WHERE CROCAB_NUMOPE = ' '"
      Set moddat_g_rst_RecDAO = moddat_g_bdt_Report.OpenRecordset(moddat_g_str_CadDAO, dbOpenDynaset)
      
      moddat_g_rst_RecDAO.AddNew
      
      moddat_g_rst_RecDAO("CROCAB_NUMOPE") = Mid(moddat_g_str_NumOpe, 1, 3) & "-" & Mid(moddat_g_str_NumOpe, 4, 2) & "-" & Mid(moddat_g_str_NumOpe, 6, 5)
      moddat_g_rst_RecDAO("CROCAB_DOCIDE") = CStr(g_rst_Princi!HIPMAE_TDOCLI) & "-" & Trim(g_rst_Princi!HIPMAE_NDOCLI & "")
      moddat_g_rst_RecDAO("CROCAB_NOMCLI") = moddat_gf_Buscar_NomCli(CStr(g_rst_Princi!HIPMAE_TDOCLI), Trim(g_rst_Princi!HIPMAE_NDOCLI))
      moddat_g_rst_RecDAO("CROCAB_MTOPRE") = g_rst_Princi!HIPMAE_MTOPRE
      moddat_g_rst_RecDAO("CROCAB_PLAPRE") = g_rst_Princi!HIPMAE_PLAANO
      moddat_g_rst_RecDAO("CROCAB_NUMCUO") = g_rst_Princi!HIPMAE_NUMCUO
      moddat_g_rst_RecDAO("CROCAB_PERGRA") = g_rst_Princi!HIPMAE_PERGRA
      moddat_g_rst_RecDAO("CROCAB_CUOEXT") = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!HIPMAE_CUOANO))
      moddat_g_rst_RecDAO("CROCAB_PRODUC") = moddat_gf_Consulta_Produc(g_rst_Princi!HIPMAE_CODPRD)
                           
      If g_rst_Princi!HIPMAE_PERGRA > 0 Then
         moddat_g_rst_RecDAO("CROCAB_INTGRA") = g_rst_Princi!HIPMAE_INTCAP
      Else
         moddat_g_rst_RecDAO("CROCAB_INTGRA") = 0
      End If
            
      moddat_g_rst_RecDAO("CROCAB_MODALI") = ""
      If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), g_rst_Princi!HIPMAE_CODPRD, g_rst_Princi!HIPMAE_CODSUB, "003", Format(CInt(CStr(g_rst_Princi!HIPMAE_CODMOD)), "000")) Then
         moddat_g_rst_RecDAO("CROCAB_MODALI") = moddat_g_arr_Genera(1).Genera_Nombre
      End If
      
      moddat_g_rst_RecDAO("CROCAB_MONEDA") = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!HIPMAE_MONEDA))
      moddat_g_rst_RecDAO("CROCAB_TASINT") = g_rst_Princi!HIPMAE_TASMVI
      moddat_g_rst_RecDAO("CROCAB_FECDES") = gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECDES))
      
      moddat_g_rst_RecDAO("CROCAB_MTOCON") = g_rst_Princi!HIPMAE_IMPNCO
      
      moddat_g_rst_RecDAO.Update
      moddat_g_rst_RecDAO.Close
      
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   
   'Grabando Detalle
   g_str_Parame = "SELECT * FROM CRE_HIPCUO WHERE "
   g_str_Parame = g_str_Parame & "HIPCUO_NUMOPE = '" & moddat_g_str_NumOpe & "' AND "
   g_str_Parame = g_str_Parame & "HIPCUO_TIPCRO = 3 "
   g_str_Parame = g_str_Parame & "ORDER BY HIPCUO_NUMCUO ASC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         'Grabando en DAO
         moddat_g_str_CadDAO = "SELECT * FROM RPT_CROCUO WHERE CROCUO_NUMOPE = ' '"
         Set moddat_g_rst_RecDAO = moddat_g_bdt_Report.OpenRecordset(moddat_g_str_CadDAO, dbOpenDynaset)
         
         moddat_g_rst_RecDAO.AddNew
         
         moddat_g_rst_RecDAO("CROCUO_NUMOPE") = Mid(moddat_g_str_NumOpe, 1, 3) & "-" & Mid(moddat_g_str_NumOpe, 4, 2) & "-" & Mid(moddat_g_str_NumOpe, 6, 5)
         moddat_g_rst_RecDAO("CROCUO_NUMCUO") = Format(g_rst_Princi!HIPCUO_NUMCUO, "000")
         moddat_g_rst_RecDAO("CROCUO_FECVCT") = gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECVCT))
         moddat_g_rst_RecDAO("CROCUO_CAPITA") = g_rst_Princi!HIPCUO_CAPITA
         moddat_g_rst_RecDAO("CROCUO_INTERE") = g_rst_Princi!HIPCUO_INTERE
         moddat_g_rst_RecDAO("CROCUO_COMCOF") = g_rst_Princi!HIPCUO_COMCOF
         moddat_g_rst_RecDAO("CROCUO_TOTCUO") = g_rst_Princi!HIPCUO_CAPITA + g_rst_Princi!HIPCUO_INTERE + g_rst_Princi!HIPCUO_COMCOF
         moddat_g_rst_RecDAO("CROCUO_SALCAP") = g_rst_Princi!HIPCUO_SALCAP
                              
         moddat_g_rst_RecDAO.Update
         moddat_g_rst_RecDAO.Close
      
         DoEvents
         g_rst_Princi.MoveNext
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Screen.MousePointer = 0
   
   Select Case moddat_g_str_CodPrd
      Case "004": crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_CROPAG_06.RPT"
   End Select
   
   crp_Imprim.Action = 1
End Sub

Private Sub fs_Cronog_Mivivienda_NCoCof()
   Screen.MousePointer = 11
   
   'Borrando Spool Local
   moddat_g_bdt_Report.Execute "DELETE FROM RPT_CROCAB"
   moddat_g_bdt_Report.Execute "DELETE FROM RPT_CROCUO"
   DoEvents
   
   'Grabando Cabecera
   g_str_Parame = "SELECT * FROM CRE_HIPMAE WHERE "
   g_str_Parame = g_str_Parame & "HIPMAE_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      'Grabando en DAO
      moddat_g_str_CadDAO = "SELECT * FROM RPT_CROCAB WHERE CROCAB_NUMOPE = ' '"
      Set moddat_g_rst_RecDAO = moddat_g_bdt_Report.OpenRecordset(moddat_g_str_CadDAO, dbOpenDynaset)
      
      moddat_g_rst_RecDAO.AddNew
      
      moddat_g_rst_RecDAO("CROCAB_NUMOPE") = Mid(moddat_g_str_NumOpe, 1, 3) & "-" & Mid(moddat_g_str_NumOpe, 4, 2) & "-" & Mid(moddat_g_str_NumOpe, 6, 5)
      moddat_g_rst_RecDAO("CROCAB_DOCIDE") = CStr(g_rst_Princi!HIPMAE_TDOCLI) & "-" & Trim(g_rst_Princi!HIPMAE_NDOCLI & "")
      moddat_g_rst_RecDAO("CROCAB_NOMCLI") = moddat_gf_Buscar_NomCli(CStr(g_rst_Princi!HIPMAE_TDOCLI), Trim(g_rst_Princi!HIPMAE_NDOCLI))
      moddat_g_rst_RecDAO("CROCAB_MTOPRE") = g_rst_Princi!HIPMAE_MTOPRE
      moddat_g_rst_RecDAO("CROCAB_PLAPRE") = g_rst_Princi!HIPMAE_PLAANO
      moddat_g_rst_RecDAO("CROCAB_NUMCUO") = g_rst_Princi!HIPMAE_NUMCUO
      moddat_g_rst_RecDAO("CROCAB_PERGRA") = g_rst_Princi!HIPMAE_PERGRA
      moddat_g_rst_RecDAO("CROCAB_CUOEXT") = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!HIPMAE_CUOANO))
      moddat_g_rst_RecDAO("CROCAB_PRODUC") = moddat_gf_Consulta_Produc(g_rst_Princi!HIPMAE_CODPRD)
                           
      If g_rst_Princi!HIPMAE_PERGRA > 0 Then
         moddat_g_rst_RecDAO("CROCAB_INTGRA") = g_rst_Princi!HIPMAE_INTCAP
      Else
         moddat_g_rst_RecDAO("CROCAB_INTGRA") = 0
      End If
            
      moddat_g_rst_RecDAO("CROCAB_MODALI") = ""
      If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), g_rst_Princi!HIPMAE_CODPRD, g_rst_Princi!HIPMAE_CODSUB, "003", Format(CInt(CStr(g_rst_Princi!HIPMAE_CODMOD)), "000")) Then
         moddat_g_rst_RecDAO("CROCAB_MODALI") = moddat_g_arr_Genera(1).Genera_Nombre
      End If
      
      moddat_g_rst_RecDAO("CROCAB_MONEDA") = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!HIPMAE_MONEDA))
      moddat_g_rst_RecDAO("CROCAB_TASINT") = g_rst_Princi!HIPMAE_TASMVI
      moddat_g_rst_RecDAO("CROCAB_FECDES") = gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECDES))
      
      
      moddat_g_rst_RecDAO.Update
      moddat_g_rst_RecDAO.Close
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   
   'Grabando Detalle
   g_str_Parame = "SELECT * FROM CRE_HIPCUO WHERE "
   g_str_Parame = g_str_Parame & "HIPCUO_NUMOPE = '" & moddat_g_str_NumOpe & "' AND "
   g_str_Parame = g_str_Parame & "HIPCUO_TIPCRO = 5 "
   g_str_Parame = g_str_Parame & "ORDER BY HIPCUO_NUMCUO ASC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         'Grabando en DAO
         moddat_g_str_CadDAO = "SELECT * FROM RPT_CROCUO WHERE CROCUO_NUMOPE = ' '"
         Set moddat_g_rst_RecDAO = moddat_g_bdt_Report.OpenRecordset(moddat_g_str_CadDAO, dbOpenDynaset)
         
         moddat_g_rst_RecDAO.AddNew
         
         moddat_g_rst_RecDAO("CROCUO_NUMOPE") = Mid(moddat_g_str_NumOpe, 1, 3) & "-" & Mid(moddat_g_str_NumOpe, 4, 2) & "-" & Mid(moddat_g_str_NumOpe, 6, 5)
         moddat_g_rst_RecDAO("CROCUO_NUMCUO") = Format(g_rst_Princi!HIPCUO_NUMCUO, "000")
         moddat_g_rst_RecDAO("CROCUO_FECVCT") = gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECVCT))
         moddat_g_rst_RecDAO("CROCUO_CAPITA") = g_rst_Princi!HIPCUO_CAPITA
         moddat_g_rst_RecDAO("CROCUO_INTERE") = g_rst_Princi!HIPCUO_INTERE
         moddat_g_rst_RecDAO("CROCUO_COMCOF") = g_rst_Princi!HIPCUO_COMCOF
         moddat_g_rst_RecDAO("CROCUO_TOTCUO") = g_rst_Princi!HIPCUO_CAPITA + g_rst_Princi!HIPCUO_INTERE + g_rst_Princi!HIPCUO_COMCOF
         moddat_g_rst_RecDAO("CROCUO_SALCAP") = g_rst_Princi!HIPCUO_SALCAP
                              
         moddat_g_rst_RecDAO.Update
         moddat_g_rst_RecDAO.Close
      
         DoEvents
         g_rst_Princi.MoveNext
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Screen.MousePointer = 0
   
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_CROPAG_07.RPT"
   
   crp_Imprim.Action = 1
End Sub

Private Sub fs_GasAdm()
   'Inicializando Variables para Hoja Resumen
   l_str_MonTas = ""
   l_str_MonNot = ""
   l_str_MonEst = ""
   l_str_MonEva = ""
   l_str_MonAdm = ""
   l_str_MonRed = ""
   l_str_MonBlq = ""
   
   l_dbl_ImpTas = 0
   l_dbl_ImpNot = 0
   l_dbl_ImpEst = 0
   l_dbl_ImpEva = 0
   l_dbl_ImpAdm = 0
   l_dbl_ImpRed = 0
   l_dbl_ImpBlq = 0

   
   g_str_Parame = "SELECT * FROM TRA_GASADM WHERE "
   g_str_Parame = g_str_Parame & "GASADM_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         Select Case g_rst_Princi!GASADM_CODGAS
            Case 11
               l_str_MonTas = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!GASADM_TIPMON))
               l_dbl_ImpTas = g_rst_Princi!GASADM_IMPORT
               
            Case 12
               l_str_MonNot = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!GASADM_TIPMON))
               l_dbl_ImpNot = g_rst_Princi!GASADM_IMPORT
               
            Case 14
               l_str_MonEst = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!GASADM_TIPMON))
               l_dbl_ImpEst = g_rst_Princi!GASADM_IMPORT
               
            Case 15
               l_str_MonEva = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!GASADM_TIPMON))
               l_dbl_ImpEva = g_rst_Princi!GASADM_IMPORT
               
            Case 16
               l_str_MonAdm = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!GASADM_TIPMON))
               l_dbl_ImpAdm = g_rst_Princi!GASADM_IMPORT
               
            Case 17
               l_str_MonRed = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!GASADM_TIPMON))
               l_dbl_ImpRed = g_rst_Princi!GASADM_IMPORT
               
            Case 18
               l_str_MonBlq = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!GASADM_TIPMON))
               l_dbl_ImpBlq = g_rst_Princi!GASADM_IMPORT
         End Select
         
         g_rst_Princi.MoveNext
         DoEvents
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub



