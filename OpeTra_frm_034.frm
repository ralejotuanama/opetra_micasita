VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "msmapi32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_Desemb_04 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form4"
   ClientHeight    =   10530
   ClientLeft      =   3195
   ClientTop       =   570
   ClientWidth     =   11640
   Icon            =   "OpeTra_frm_034.frx":0000
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10530
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   10515
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11640
      _Version        =   65536
      _ExtentX        =   20532
      _ExtentY        =   18547
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
      Begin Threed.SSPanel SSPanel15 
         Height          =   2595
         Left            =   30
         TabIndex        =   52
         Top             =   7050
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   4577
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
            Height          =   465
            Left            =   90
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   57
            Text            =   "OpeTra_frm_034.frx":000C
            Top             =   2070
            Width           =   11355
         End
         Begin MSFlexGridLib.MSFlexGrid grd_LisEva 
            Height          =   1305
            Left            =   60
            TabIndex        =   53
            Top             =   300
            Width           =   11445
            _ExtentX        =   20188
            _ExtentY        =   2302
            _Version        =   393216
            Rows            =   21
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   49152
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel16 
            Height          =   60
            Left            =   60
            TabIndex        =   58
            Top             =   1650
            Width           =   11385
            _Version        =   65536
            _ExtentX        =   20082
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
         Begin VB.Label Label3 
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
            TabIndex        =   59
            Top             =   1770
            Width           =   2805
         End
         Begin VB.Label Label14 
            Caption         =   "Resumen de Evaluación"
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
            TabIndex        =   54
            Top             =   60
            Width           =   2805
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   5115
         Left            =   30
         TabIndex        =   1
         Top             =   1890
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   9022
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
            Height          =   4755
            Left            =   60
            TabIndex        =   2
            Top             =   300
            Width           =   11385
            _ExtentX        =   20082
            _ExtentY        =   8387
            _Version        =   393216
            Style           =   1
            Tabs            =   12
            TabsPerRow      =   6
            TabHeight       =   520
            TabCaption(0)   =   "Datos del Cliente"
            TabPicture(0)   =   "OpeTra_frm_034.frx":0010
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "grd_Listad(0)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Datos del Cónyuge"
            TabPicture(1)   =   "OpeTra_frm_034.frx":002C
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "grd_Listad(1)"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Datos del Patrimonio"
            TabPicture(2)   =   "OpeTra_frm_034.frx":0048
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "grd_Listad(4)"
            Tab(2).ControlCount=   1
            TabCaption(3)   =   "Referencias Personales"
            TabPicture(3)   =   "OpeTra_frm_034.frx":0064
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "grd_Listad(3)"
            Tab(3).ControlCount=   1
            TabCaption(4)   =   "Datos del Inmueble"
            TabPicture(4)   =   "OpeTra_frm_034.frx":0080
            Tab(4).ControlEnabled=   0   'False
            Tab(4).Control(0)=   "grd_Listad(2)"
            Tab(4).ControlCount=   1
            TabCaption(5)   =   "Datos del Crédito"
            TabPicture(5)   =   "OpeTra_frm_034.frx":009C
            Tab(5).ControlEnabled=   0   'False
            Tab(5).Control(0)=   "grd_Listad(5)"
            Tab(5).Control(1)=   "grd_Listad(10)"
            Tab(5).Control(2)=   "SSPanel13"
            Tab(5).Control(3)=   "Label10"
            Tab(5).ControlCount=   4
            TabCaption(6)   =   "Gastos Administ."
            TabPicture(6)   =   "OpeTra_frm_034.frx":00B8
            Tab(6).ControlEnabled=   0   'False
            Tab(6).Control(0)=   "grd_GasAdm"
            Tab(6).Control(1)=   "pnl_TotGas"
            Tab(6).Control(2)=   "SSPanel11"
            Tab(6).Control(3)=   "SSPanel8"
            Tab(6).Control(4)=   "SSPanel10"
            Tab(6).Control(5)=   "SSPanel12"
            Tab(6).Control(6)=   "SSPanel9"
            Tab(6).Control(7)=   "Label8"
            Tab(6).ControlCount=   8
            TabCaption(7)   =   "Evaluación Crediticia"
            TabPicture(7)   =   "OpeTra_frm_034.frx":00D4
            Tab(7).ControlEnabled=   0   'False
            Tab(7).Control(0)=   "grd_Listad(6)"
            Tab(7).ControlCount=   1
            TabCaption(8)   =   "Tasación del Inmueble"
            TabPicture(8)   =   "OpeTra_frm_034.frx":00F0
            Tab(8).ControlEnabled=   0   'False
            Tab(8).Control(0)=   "grd_Listad(7)"
            Tab(8).Control(1)=   "grd_Listad(11)"
            Tab(8).Control(2)=   "SSPanel14"
            Tab(8).Control(3)=   "Label11"
            Tab(8).ControlCount=   4
            TabCaption(9)   =   "Evaluación de Seguros"
            TabPicture(9)   =   "OpeTra_frm_034.frx":010C
            Tab(9).ControlEnabled=   0   'False
            Tab(9).Control(0)=   "txt_ObsSeg"
            Tab(9).Control(1)=   "grd_Listad(8)"
            Tab(9).Control(2)=   "SSPanel5"
            Tab(9).Control(3)=   "Label7"
            Tab(9).ControlCount=   4
            TabCaption(10)  =   "Evaluación Legal"
            TabPicture(10)  =   "OpeTra_frm_034.frx":0128
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
            TabPicture(11)  =   "OpeTra_frm_034.frx":0144
            Tab(11).ControlEnabled=   0   'False
            Tab(11).Control(0)=   "txt_ObsMVi"
            Tab(11).Control(1)=   "grd_Listad(12)"
            Tab(11).Control(2)=   "SSPanel17"
            Tab(11).Control(3)=   "Label6"
            Tab(11).ControlCount=   4
            Begin VB.TextBox txt_ObsMVi 
               Height          =   1155
               Left            =   -74970
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   60
               Text            =   "OpeTra_frm_034.frx":0160
               Top             =   3420
               Width           =   11085
            End
            Begin VB.TextBox txt_ComCre 
               Height          =   615
               Left            =   -74910
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   5
               Text            =   "OpeTra_frm_034.frx":0164
               Top             =   2910
               Width           =   11085
            End
            Begin VB.TextBox txt_InfLeg 
               Height          =   1515
               Left            =   -74940
               MaxLength       =   8000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   4
               Text            =   "OpeTra_frm_034.frx":0168
               Top             =   960
               Width           =   11085
            End
            Begin VB.TextBox txt_ObsSeg 
               Height          =   675
               Left            =   -74910
               MaxLength       =   2000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   3
               Text            =   "OpeTra_frm_034.frx":016C
               Top             =   3990
               Width           =   11085
            End
            Begin Threed.SSPanel SSPanel3 
               Height          =   60
               Left            =   -74970
               TabIndex        =   6
               Top             =   3600
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
               Height          =   4005
               Index           =   0
               Left            =   60
               TabIndex        =   7
               Top             =   690
               Width           =   11085
               _ExtentX        =   19553
               _ExtentY        =   7064
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   49152
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   4005
               Index           =   2
               Left            =   -74940
               TabIndex        =   8
               Top             =   690
               Width           =   11085
               _ExtentX        =   19553
               _ExtentY        =   7064
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   49152
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   4005
               Index           =   3
               Left            =   -74940
               TabIndex        =   9
               Top             =   690
               Width           =   11085
               _ExtentX        =   19553
               _ExtentY        =   7064
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   49152
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   4005
               Index           =   4
               Left            =   -74940
               TabIndex        =   10
               Top             =   690
               Width           =   11085
               _ExtentX        =   19553
               _ExtentY        =   7064
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   49152
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
               Height          =   4005
               Index           =   1
               Left            =   -74940
               TabIndex        =   11
               Top             =   690
               Width           =   11085
               _ExtentX        =   19553
               _ExtentY        =   7064
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   49152
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   2235
               Index           =   5
               Left            =   -74940
               TabIndex        =   12
               Top             =   690
               Width           =   11085
               _ExtentX        =   19553
               _ExtentY        =   3942
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   49152
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   4005
               Index           =   6
               Left            =   -74940
               TabIndex        =   13
               Top             =   690
               Width           =   11085
               _ExtentX        =   19553
               _ExtentY        =   7064
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   49152
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   2145
               Index           =   7
               Left            =   -74940
               TabIndex        =   14
               Top             =   690
               Width           =   11085
               _ExtentX        =   19553
               _ExtentY        =   3784
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   49152
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   2805
               Index           =   8
               Left            =   -74940
               TabIndex        =   15
               Top             =   690
               Width           =   11085
               _ExtentX        =   19553
               _ExtentY        =   4948
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   49152
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   705
               Index           =   9
               Left            =   -74940
               TabIndex        =   16
               Top             =   3990
               Width           =   11085
               _ExtentX        =   19553
               _ExtentY        =   1244
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   49152
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel SSPanel4 
               Height          =   60
               Left            =   -74940
               TabIndex        =   17
               Top             =   2520
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
               Left            =   -74940
               TabIndex        =   18
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
            Begin MSFlexGridLib.MSFlexGrid grd_GasAdm 
               Height          =   3225
               Left            =   -74970
               TabIndex        =   19
               Top             =   990
               Width           =   11115
               _ExtentX        =   19606
               _ExtentY        =   5689
               _Version        =   393216
               Rows            =   21
               Cols            =   5
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   49152
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel pnl_TotGas 
               Height          =   315
               Left            =   -65250
               TabIndex        =   20
               Top             =   4290
               Width           =   1215
               _Version        =   65536
               _ExtentX        =   2143
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
            Begin Threed.SSPanel SSPanel11 
               Height          =   285
               Left            =   -74940
               TabIndex        =   21
               Top             =   690
               Width           =   3975
               _Version        =   65536
               _ExtentX        =   7011
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Concepto"
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
               Left            =   -71010
               TabIndex        =   22
               Top             =   690
               Width           =   2385
               _Version        =   65536
               _ExtentX        =   4207
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Tipo de Moneda"
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
               Left            =   -68670
               TabIndex        =   23
               Top             =   690
               Width           =   1215
               _Version        =   65536
               _ExtentX        =   2143
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Importe"
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
               Left            =   -67470
               TabIndex        =   24
               Top             =   690
               Width           =   1965
               _Version        =   65536
               _ExtentX        =   3466
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Situación"
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
            Begin Threed.SSPanel SSPanel9 
               Height          =   285
               Left            =   -65520
               TabIndex        =   25
               Top             =   690
               Width           =   1365
               _Version        =   65536
               _ExtentX        =   2408
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Fecha Pago"
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
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   1215
               Index           =   10
               Left            =   -74940
               TabIndex        =   26
               Top             =   3330
               Width           =   11085
               _ExtentX        =   19553
               _ExtentY        =   2143
               _Version        =   393216
               Rows            =   21
               Cols            =   1
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   49152
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel SSPanel13 
               Height          =   60
               Left            =   -74970
               TabIndex        =   27
               Top             =   2940
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
               Height          =   1365
               Index           =   11
               Left            =   -74940
               TabIndex        =   28
               Top             =   3300
               Width           =   11085
               _ExtentX        =   19553
               _ExtentY        =   2408
               _Version        =   393216
               Rows            =   21
               Cols            =   1
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   49152
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel SSPanel14 
               Height          =   60
               Left            =   -74940
               TabIndex        =   29
               Top             =   2910
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
               Height          =   2175
               Index           =   12
               Left            =   -74940
               TabIndex        =   61
               Top             =   690
               Width           =   11085
               _ExtentX        =   19553
               _ExtentY        =   3836
               _Version        =   393216
               Rows            =   21
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   49152
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel SSPanel17 
               Height          =   60
               Left            =   -75000
               TabIndex        =   62
               Top             =   2940
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
               Left            =   -74970
               TabIndex        =   63
               Top             =   3060
               Width           =   2805
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
               TabIndex        =   36
               Top             =   3030
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
               TabIndex        =   35
               Top             =   3060
               Width           =   2805
            End
            Begin VB.Label Label8 
               Caption         =   "Total de Gastos:"
               Height          =   315
               Left            =   -66630
               TabIndex        =   34
               Top             =   4290
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
               Left            =   -74910
               TabIndex        =   33
               Top             =   3690
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
               TabIndex        =   32
               Top             =   3720
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
               Left            =   -74910
               TabIndex        =   31
               Top             =   2640
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
               TabIndex        =   30
               Top             =   690
               Width           =   2805
            End
         End
         Begin VB.Label Label2 
            Caption         =   "Información de la Solicitud de Crédito"
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
            TabIndex        =   37
            Top             =   60
            Width           =   4035
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   38
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
            Height          =   495
            Left            =   630
            TabIndex        =   39
            Top             =   60
            Width           =   8265
            _Version        =   65536
            _ExtentX        =   14579
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Desembolso de Créditos Hipotecarios"
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
         Begin VB.Image Image1 
            Height          =   480
            Left            =   60
            Picture         =   "OpeTra_frm_034.frx":0170
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   1095
         Left            =   30
         TabIndex        =   40
         Top             =   750
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
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
         Begin Threed.SSPanel pnl_Produc 
            Height          =   315
            Left            =   1440
            TabIndex        =   41
            Top             =   60
            Width           =   4485
            _Version        =   65536
            _ExtentX        =   7911
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
            Left            =   1440
            TabIndex        =   42
            Top             =   390
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
         Begin Threed.SSPanel pnl_NumOpe 
            Height          =   315
            Left            =   1440
            TabIndex        =   48
            Top             =   720
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
         Begin Threed.SSPanel pnl_GenOpe 
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
         Begin VB.Label Label13 
            Caption         =   "F. Generac. Operac.:"
            Height          =   315
            Left            =   8400
            TabIndex        =   51
            Top             =   60
            Width           =   1545
         End
         Begin VB.Label Label12 
            Caption         =   "Nro. Operación:"
            Height          =   315
            Left            =   60
            TabIndex        =   49
            Top             =   720
            Width           =   1245
         End
         Begin VB.Label Label21 
            Caption         =   "Producto:"
            Height          =   315
            Left            =   60
            TabIndex        =   44
            Top             =   60
            Width           =   1275
         End
         Begin VB.Label Label1 
            Caption         =   "Nro. Solicitud:"
            Height          =   315
            Left            =   60
            TabIndex        =   43
            Top             =   390
            Width           =   1335
         End
      End
      Begin Threed.SSPanel SSPanel39 
         Height          =   765
         Left            =   30
         TabIndex        =   45
         Top             =   9690
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
         Begin VB.CommandButton cmd_Export 
            Height          =   675
            Left            =   720
            Picture         =   "OpeTra_frm_034.frx":047A
            Style           =   1  'Graphical
            TabIndex        =   56
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   675
            Left            =   30
            Picture         =   "OpeTra_frm_034.frx":0784
            Style           =   1  'Graphical
            TabIndex        =   55
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Desemb 
            Height          =   675
            Left            =   10140
            Picture         =   "OpeTra_frm_034.frx":0BC6
            Style           =   1  'Graphical
            TabIndex        =   47
            ToolTipText     =   "Aprobar Solicitud"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   10830
            Picture         =   "OpeTra_frm_034.frx":0ED0
            Style           =   1  'Graphical
            TabIndex        =   46
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   675
         End
         Begin MSMAPI.MAPIMessages mps_Mensaj 
            Left            =   4110
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
            Left            =   3540
            Top             =   60
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            DownloadMail    =   -1  'True
            LogonUI         =   -1  'True
            NewSession      =   0   'False
         End
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   2940
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
            Left            =   4950
            Top             =   90
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
      End
   End
End
Attribute VB_Name = "frm_Desemb_04"
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
'   On Error GoTo cmd_ArcTxt_Error
   
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
   
   
cmd_ArcTxt_Error:
   'Exit Sub


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

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt

   pnl_NumSol.Caption = Mid(moddat_g_str_NumSol, 1, 3) & "-" & Mid(moddat_g_str_NumSol, 4, 3) & "-" & Mid(moddat_g_str_NumSol, 7, 2) & "-" & Mid(moddat_g_str_NumSol, 9, 4)
   pnl_NumOpe.Caption = Mid(moddat_g_str_NumOpe, 1, 3) & "-" & Mid(moddat_g_str_NumOpe, 4, 2) & "-" & Mid(moddat_g_str_NumOpe, 6, 5)
   pnl_Produc.Caption = moddat_g_str_NomPrd
   pnl_GenOpe.Caption = moddat_g_str_FecIng
   
   Call fs_Inicia
   
   'Buscar Información de la Solicitud
   moddat_g_int_CygTDo = 0
   moddat_g_str_CygNDo = ""
   
   Call fs_DatCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc, 0)
   Call fs_DatCli(moddat_g_int_CygTDo, moddat_g_str_CygNDo, 1)
   Call fs_DatPat        'Datos del Patrimonio
   Call fs_DatRef        'Referencias Personales
   Call fs_DatInm        'Datos del Inmueble
   Call fs_DatCre        'Datos del Crédito
   Call fs_SolDoc        'Documentos Recibidos
   Call fs_SolDoc_Inm    'Documentos Recibidos del Inmueble
   Call fs_GasAdm        'Gastos Administrativos
   Call fs_EvaCre        'Evaluación Crediticia
   Call fs_DatTas        'Tasación
   Call fs_DatSeg        'Seguros
   Call fs_DatLeg        'Legal
   Call fs_DatMVi        'Mivivienda
   Call fs_DatCof        'COFIDE

   Call fs_Carga_DatEva
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

   'Grid de Evaluación
   grd_LisEva.ColWidth(0) = 3200
   grd_LisEva.ColWidth(1) = 7940

   grd_LisEva.ColAlignment(0) = flexAlignLeftCenter
   grd_LisEva.ColAlignment(1) = flexAlignLeftCenter
End Sub

Private Sub fs_DatCli(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String, ByVal p_Indice As Integer)
   g_str_Parame = "SELECT * FROM CLI_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_TIPDOC = " & CStr(p_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_NUMDOC = '" & p_NumDoc & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_Listad(p_Indice).Redraw = False
      
      g_rst_Princi.MoveFirst
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).Text = "Documento de Identidad"
      
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("203", CStr(g_rst_Princi!DatGen_TipDoc)) & " - " & Trim(g_rst_Princi!DatGen_NumDoc & "")
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).Text = "Apellidos y Nombres"
      
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!DatGen_ApePat) & " " & Trim(g_rst_Princi!DatGen_ApeMat) & IIf(Len(Trim(g_rst_Princi!DatGen_ApeCas)) > 0, " DE " & Trim(g_rst_Princi!DatGen_ApeCas), "") & " " & Trim(g_rst_Princi!DatGen_Nombre)
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).Text = "Documento Adicional de Identidad"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!DatGen_FLGDOA)) & IIf(g_rst_Princi!DatGen_FLGDOA = 1, " ( " & moddat_gf_Consulta_ParDes("203", CStr(g_rst_Princi!DatGen_TIPDOA)) & " - " & Trim(g_rst_Princi!DatGen_NUMDOA) & ")", "")
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).Text = "Sexo"
      
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("207", CStr(g_rst_Princi!DatGen_CodSex))
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).Text = "Fecha de Nacimiento"
      
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).Text = gf_FormatoFecha(CStr(g_rst_Princi!DATGEN_NACFEC))
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).Text = "Nacionalidad"
      
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("500", Trim(g_rst_Princi!DATGEN_NACPAI))
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).Text = "Lugar de Nacimiento"
   
      If Trim(g_rst_Princi!DATGEN_NACPAI) = "004028" Then
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!DATGEN_NACLUG, 2) & "0000") & " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!DATGEN_NACLUG, 4) & "00") & " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!DATGEN_NACLUG))
      Else
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).Text = "<< NO REGISTRADO >>"
      End If
      
      If p_Indice = 0 Then
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).Text = "Estado Civil"
         
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("205", CStr(g_rst_Princi!DATGEN_ESTCIV)) & IIf(g_rst_Princi!DATGEN_ESTCIV = 2, " / " & moddat_gf_Consulta_ParDes("206", g_rst_Princi!DatGen_RegCyg), "")
         
         If g_rst_Princi!DATGEN_ESTCIV = 2 Or g_rst_Princi!DATGEN_ESTCIV = 5 Then
            moddat_g_int_CygTDo = g_rst_Princi!DATGEN_CYGTDO
            moddat_g_str_CygNDo = Trim(g_rst_Princi!DATGEN_CYGNDO & "")
         End If
      End If
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).Text = "Nivel de Estudios"
      
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("209", CStr(g_rst_Princi!DatGen_NivEst))

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).Text = "Profesión"
      
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("501", CStr(g_rst_Princi!DatGen_Profes))

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).Text = "Celular"
      
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!DATGEN_NUMCEL & "")
      
      If p_Indice = 0 Then
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).Text = "Nro. Dependientes Económicos"
         
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).Text = CStr(g_rst_Princi!DatGen_DepEco)
      
      
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).Text = "Edades"
         
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).Text = IIf(g_rst_Princi!DatGen_EDAD01 > 0, CStr(g_rst_Princi!DatGen_EDAD01), "") & _
                                     IIf(g_rst_Princi!DatGen_EDAD02 > 0, " - " & CStr(g_rst_Princi!DatGen_EDAD02), "") & _
                                     IIf(g_rst_Princi!DatGen_EDAD03 > 0, " - " & CStr(g_rst_Princi!DatGen_EDAD03), "") & _
                                     IIf(g_rst_Princi!DatGen_EDAD04 > 0, " - " & CStr(g_rst_Princi!DatGen_EDAD04), "") & _
                                     IIf(g_rst_Princi!DatGen_EDAD05 > 0, " - " & CStr(g_rst_Princi!DatGen_EDAD05), "")
      End If
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).Text = "E-mail"
      
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!DatGen_DirEle & "")
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).Text = "Autorización Envío"
      
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!DATGEN_AUTENV))
      
      If p_Indice = 0 Then
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).Text = "Domicilio"
         
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!DatGen_TipVia)) & _
                                     " " & Trim(g_rst_Princi!DatGen_NomVia) & " " & Trim(g_rst_Princi!DatGen_Numero) & _
                                     IIf(Len(Trim(g_rst_Princi!DatGen_IntDpt)) > 0, " (" & Trim(g_rst_Princi!DatGen_IntDpt) & ")", "") & _
                                     IIf(Len(Trim(g_rst_Princi!DatGen_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!DatGen_TipZon)) & " " & Trim(g_rst_Princi!DatGen_NomZon), "")
         
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).Text = "Referencia"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).Text = Trim(g_rst_Princi!DatGen_Refere & "")
         
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).Text = "Departamento / Provincia / Distrito"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!DatGen_Ubigeo, 2) & "0000") & _
                                     " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!DatGen_Ubigeo, 4) & "00") & _
                                     " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!DatGen_Ubigeo))
      
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).Text = "Teléfono Domicilio"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).Text = Trim(g_rst_Princi!DatGen_Telefo & "")
      End If
      
      grd_Listad(p_Indice).Redraw = True
      Call gs_UbiIniGrid(grd_Listad(p_Indice))
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Actividad Económica Principal
   Call fs_ActEco(p_TipDoc, p_NumDoc, 1, p_Indice)
   Call fs_ActEco(p_TipDoc, p_NumDoc, 2, p_Indice)
End Sub

Private Sub fs_ActEco(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String, ByVal p_OrdAct As Integer, ByVal p_Indice As Integer)
   Dim r_var_ColTxt
   
   
   If p_OrdAct = 1 Then
      r_var_ColTxt = modgen_g_con_ColAzu
   Else
      r_var_ColTxt = modgen_g_con_ColRoj
   End If

   g_str_Parame = "SELECT * FROM CLI_ACTECO WHERE "
   g_str_Parame = g_str_Parame & "ACTECO_CLITDO = " & CStr(p_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "ACTECO_CLINDO = '" & p_NumDoc & "' AND "
   g_str_Parame = g_str_Parame & "ACTECO_ORDACT = " & CStr(p_OrdAct)

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_Listad(p_Indice).Redraw = False
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = r_var_ColTxt
      grd_Listad(p_Indice).Text = "Ocupación " & Left(moddat_gf_Consulta_ParDes("007", CStr(g_rst_Princi!ActEco_OrdAct)), 1) & Mid(LCase(moddat_gf_Consulta_ParDes("007", CStr(g_rst_Princi!ActEco_OrdAct))), 2)
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = r_var_ColTxt
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("008", g_rst_Princi!ActEco_CodAct)
      
      Select Case g_rst_Princi!ActEco_CodAct
         Case 11: Call fs_ActEco_Dep(p_Indice, r_var_ColTxt)
         Case 21: Call fs_ActEco_Ind(p_Indice, r_var_ColTxt)
         Case 31: Call fs_ActEco_Com(p_Indice, r_var_ColTxt)
         Case 41: Call fs_ActEco_Acc(p_Indice, r_var_ColTxt)
         Case 51: Call fs_ActEco_Ren(p_Indice, r_var_ColTxt)
         Case 61: Call fs_ActEco_Otr(p_Indice, r_var_ColTxt)
      End Select
      
      grd_Listad(p_Indice).Redraw = True
      Call gs_UbiIniGrid(grd_Listad(p_Indice))
   End If

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_ActEco_Dep(ByVal p_Indice As Integer, ByVal p_ColTxt)
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Documento Identidad Empleador"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("203", g_rst_Princi!ActEco_Dep_TipDoc) & " - " & Trim(g_rst_Princi!ActEco_Dep_NumDoc & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Situación como Trabajador"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("235", g_rst_Princi!ActEco_Dep_SitTra)

   'Buscar en Maestro de Empresas
   g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(g_rst_Princi!ActEco_Dep_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & g_rst_Princi!ActEco_Dep_NumDoc & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Razón Social"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_RazSoc & "")
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Nombre Comercial"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_NomCom & "")
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "CIIU"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = g_rst_Princi!ActEco_Dep_CodCiu & " - " & moddat_gf_Consulta_ParDes("102", CStr(g_rst_Princi!ActEco_Dep_CodCiu))
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Teléfono RR.HH"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_TeleRH & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Dep_AnexRH & "")) > 0, " ANEXO: " & Trim(g_rst_Princi!ActEco_Dep_AnexRH & ""), "")
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Dirección"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!ActEco_Dep_TipVia)) & _
                                  " " & Trim(g_rst_Princi!ActEco_Dep_NomVia) & " " & Trim(g_rst_Princi!ActEco_Dep_NumVia) & _
                                  IIf(Len(Trim(g_rst_Princi!ActEco_Dep_IntDpt)) > 0, " (" & Trim(g_rst_Princi!ActEco_Dep_IntDpt) & ")", "") & _
                                  IIf(Len(Trim(g_rst_Princi!ActEco_Dep_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!ActEco_Dep_TipZon)) & " " & Trim(g_rst_Princi!ActEco_Dep_NomZon), "")

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Referencia"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_Refere & "")

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Departamento / Provincia / Distrito"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Dep_UbiGeo, 2) & "0000") & _
                                  " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Dep_UbiGeo, 4) & "00") & _
                                  " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!ActEco_Dep_UbiGeo))
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Teléfono(s)"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_Telef1 & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Dep_Telef2 & "")) > 0, " - " & Trim(g_rst_Princi!ActEco_Dep_Telef2 & ""), "")
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Fax"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_NumFax & "")
   Else
      g_rst_Genera.MoveFirst

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Razón Social"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_RAZSOC & "")

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Razón Social"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_NOMCOM & "")

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "CIIU"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = g_rst_Genera!DATGEN_CODCIU & " - " & moddat_gf_Consulta_ParDes("102", CStr(g_rst_Genera!DATGEN_CODCIU))

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Teléfono RR.HH"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_TELERH & "") & IIf(Len(Trim(g_rst_Genera!DATGEN_TELERH & "")) > 0, " ANEXO: " & Trim(g_rst_Genera!DATGEN_ANEXRH & ""), "")

      If g_rst_Princi!ActEco_Dep_TipOfi = 1 Then
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Dirección"
      
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Genera!DatGen_TipVia)) & _
                                     " " & Trim(g_rst_Genera!DatGen_NomVia) & " " & Trim(g_rst_Genera!DatGen_numVia) & _
                                     IIf(Len(Trim(g_rst_Genera!DatGen_IntDpt)) > 0, " (" & Trim(g_rst_Genera!DatGen_IntDpt) & ")", "") & _
                                     IIf(Len(Trim(g_rst_Genera!DatGen_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Genera!DatGen_TipZon)) & " " & Trim(g_rst_Genera!DatGen_NomZon), "")

         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Referencia"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DatGen_Refere & "")

         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Departamento / Provincia / Distrito"

         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Genera!DatGen_Ubigeo, 2) & "0000") & _
                                     " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Genera!DatGen_Ubigeo, 4) & "00") & _
                                     " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Genera!DatGen_Ubigeo))
      
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Teléfono(s)"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_TELEF1 & "") & IIf(Len(Trim(g_rst_Genera!DATGEN_TELEF2 & "")) > 0, " - " & Trim(g_rst_Genera!DATGEN_TELEF2 & ""), "")
      
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Fax"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DatGen_NUMFAX & "")
      Else
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Dirección"
      
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!ActEco_Dep_TipVia)) & _
                                     " " & Trim(g_rst_Princi!ActEco_Dep_NomVia) & " " & Trim(g_rst_Princi!ActEco_Dep_NumVia) & _
                                     IIf(Len(Trim(g_rst_Princi!ActEco_Dep_IntDpt)) > 0, " (" & Trim(g_rst_Princi!ActEco_Dep_IntDpt) & ")", "") & _
                                     IIf(Len(Trim(g_rst_Princi!ActEco_Dep_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!ActEco_Dep_TipZon)) & " " & Trim(g_rst_Princi!ActEco_Dep_NomZon), "")
   
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Referencia"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_Refere & "")
   
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Departamento / Provincia / Distrito"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Dep_UbiGeo, 2) & "0000") & _
                                     " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Dep_UbiGeo, 4) & "00") & _
                                     " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!ActEco_Dep_UbiGeo))
         
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Teléfono(s)"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_Telef1 & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Dep_Telef2 & "")) > 0, " - " & Trim(g_rst_Princi!ActEco_Dep_Telef2 & ""), "")
         
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Fax"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_NumFax & "")
      End If
   End If

   g_rst_Genera.Close
   Set g_rst_Genera = Nothing

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Ingreso Neto"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoNumero(g_rst_Princi!ActEco_Dep_IngNet, 15, 2)

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Frecuencia de Haberes"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("210", CStr(g_rst_Princi!ActEco_Dep_FreHab))

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Fecha de Ingreso"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Dep_FecIng))

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Cargo"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = IIf(g_rst_Princi!ActEco_Dep_CodCar = "999999", Trim(g_rst_Princi!ActEco_Dep_NomCar & ""), moddat_gf_Consulta_ParDes("503", g_rst_Princi!ActEco_Dep_CodCar))

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Area"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_NomAre & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Anexo"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_NumAnx & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Teléfono Directo"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_TelDir & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Celular"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_Celula & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "E-mail"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_DirEle & "")

   If g_rst_Princi!ActEco_Dep_TraAnt = 1 Then
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Documento Identidad Empleador Anterior"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("203", g_rst_Princi!ActEco_Dep_TipDoc_Ant) & " - " & Trim(g_rst_Princi!ActEco_Dep_NumDoc_Ant & "")
      
      'Buscar en Maestro de Empresas
      g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
      g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(g_rst_Princi!ActEco_Dep_TipDoc_Ant) & " AND "
      g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & g_rst_Princi!ActEco_Dep_NumDoc_Ant & "' "

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
         Exit Sub
      End If
      
      If g_rst_Genera.BOF And g_rst_Genera.EOF Then
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Razón Social (Empleador Anterior)"
      
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_RazSoc_Ant & "")
   
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Nombre Comercial (Empleador Anterior)"
      
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_NomCom_Ant & "")
      
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Teléfono(s) (Empleador Anterior)"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Dep_Telef1_Ant & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Dep_Telef2_Ant & "")) > 0, " - " & Trim(g_rst_Princi!ActEco_Dep_Telef2_Ant & ""), "")
      Else
         g_rst_Genera.MoveFirst

         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Razón Social"
      
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_RAZSOC & "")

         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Razón Social"
      
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_NOMCOM & "")

         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Teléfono(s)"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_TELEF1 & "") & IIf(Len(Trim(g_rst_Genera!DATGEN_TELEF2 & "")) > 0, " - " & Trim(g_rst_Genera!DATGEN_TELEF2 & ""), "")
      End If

      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Fecha de Ingreso (Empleador Anterior)"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Dep_FecIng_Ant))
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Fecha de Cese (Empleador Anterior)"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Dep_FecCes_Ant))
   End If
End Sub

Private Sub fs_ActEco_Ind(ByVal p_Indice As Integer, ByVal p_ColTxt)
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Documento Identidad"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("203", g_rst_Princi!ActEco_Ind_TipDoc) & " - " & Trim(g_rst_Princi!ActEco_Ind_NumDoc & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Dirección"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!ActEco_Ind_TipVia)) & _
                               " " & Trim(g_rst_Princi!ActEco_Ind_NomVia) & " " & Trim(g_rst_Princi!ActEco_Ind_NumVia) & _
                               IIf(Len(Trim(g_rst_Princi!ActEco_Ind_IntDpt)) > 0, " (" & Trim(g_rst_Princi!ActEco_Ind_IntDpt) & ")", "") & _
                               IIf(Len(Trim(g_rst_Princi!ActEco_Ind_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!ActEco_Ind_TipZon)) & " " & Trim(g_rst_Princi!ActEco_Ind_NomZon), "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Referencia"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Ind_Refere & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Indartamento / Provincia / Distrito"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Ind_UbiGeo, 2) & "0000") & _
                               " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Ind_UbiGeo, 4) & "00") & _
                               " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!ActEco_Ind_UbiGeo))
   
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Teléfono(s)"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Ind_Telef1 & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Ind_Telef2 & "")) > 0, " - " & Trim(g_rst_Princi!ActEco_Ind_Telef2 & ""), "")
   
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Fax"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Ind_NumFax & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "CIIU"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = g_rst_Princi!ActEco_Ind_CodCiu & " - " & moddat_gf_Consulta_ParDes("102", CStr(g_rst_Princi!ActEco_Ind_CodCiu))

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Ingreso Neto"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoNumero(g_rst_Princi!ActEco_Ind_IngNet, 15, 2)
   
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Fecha de Inicio de Actividades"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Ind_IniAct))
   
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Contrato de Locación de Servicios"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!ActEco_Ind_ConLoc))
   
   If g_rst_Princi!ActEco_Ind_ConLoc = 1 Then
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Documento Identidad Empleador"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("203", g_rst_Princi!ActEco_Ind_TipDoc_Emp) & " - " & Trim(g_rst_Princi!ActEco_Ind_NumDoc_Emp & "")
      
      'Buscar en Maestro de Empresas
      g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
      g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(g_rst_Princi!ActEco_Ind_TipDoc_Emp) & " AND "
      g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & g_rst_Princi!ActEco_Ind_NumDoc_Emp & "' "

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
         Exit Sub
      End If
      
      If g_rst_Genera.BOF And g_rst_Genera.EOF Then
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Razón Social"
      
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Ind_RazSoc_Emp & "")
   
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Nombre Comercial"
      
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Ind_NomCom_Emp & "")
      
         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Teléfono(s)"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Ind_Telef1_Emp & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Ind_Telef2_Emp & "")) > 0, " - " & Trim(g_rst_Princi!ActEco_Ind_Telef2_Emp & ""), "")
      Else
         g_rst_Genera.MoveFirst

         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Razón Social"
      
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_RAZSOC & "")

         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Razón Social"
      
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_NOMCOM & "")

         grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
         grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
         grd_Listad(p_Indice).Col = 0
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = "Teléfono(s)"
   
         grd_Listad(p_Indice).Col = 1
         grd_Listad(p_Indice).CellForeColor = p_ColTxt
         grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_TELEF1 & "") & IIf(Len(Trim(g_rst_Genera!DATGEN_TELEF2 & "")) > 0, " - " & Trim(g_rst_Genera!DATGEN_TELEF2 & ""), "")
      End If

      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Fecha de Ingreso"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Ind_FecIng_Emp))
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Cargo"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = IIf(g_rst_Princi!ActEco_Ind_CodCar = "999999", Trim(g_rst_Princi!ActEco_Ind_NomCar & ""), moddat_gf_Consulta_ParDes("503", g_rst_Princi!ActEco_Ind_CodCar))
   End If
End Sub

Private Sub fs_ActEco_Com(ByVal p_Indice As Integer, ByVal p_ColTxt)
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Documento Identidad"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("203", g_rst_Princi!ActEco_Com_TipDoc) & " - " & Trim(g_rst_Princi!ActEco_Com_NumDoc & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Razón Social"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Com_RazSoc & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Nombre Comercial"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Com_NomCom & "")
   
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Dirección"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!ActEco_Com_TipVia)) & _
                               " " & Trim(g_rst_Princi!ActEco_Com_NomVia) & " " & Trim(g_rst_Princi!ActEco_Com_NumVia) & _
                               IIf(Len(Trim(g_rst_Princi!ActEco_Com_IntDpt)) > 0, " (" & Trim(g_rst_Princi!ActEco_Com_IntDpt) & ")", "") & _
                               IIf(Len(Trim(g_rst_Princi!ActEco_Com_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!ActEco_Com_TipZon)) & " " & Trim(g_rst_Princi!ActEco_Com_NomZon), "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Referencia"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Com_Refere & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Departamento / Provincia / Distrito"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Com_UbiGeo, 2) & "0000") & _
                               " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Com_UbiGeo, 4) & "00") & _
                               " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!ActEco_Com_UbiGeo))
   
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Teléfono(s)"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Com_Telef1 & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Com_Telef2 & "")) > 0, " - " & Trim(g_rst_Princi!ActEco_Com_Telef2 & ""), "")
   
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Fax"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Com_NumFax & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "CIIU"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = g_rst_Princi!ActEco_Com_CodCiu & " - " & moddat_gf_Consulta_ParDes("102", CStr(g_rst_Princi!ActEco_Com_CodCiu))

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Giro Comercial"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_GirCom(g_rst_Princi!ActEco_Com_GirCom)

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Ingreso Neto"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoNumero(g_rst_Princi!ActEco_Com_IngNet, 15, 2)

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Ventas Mensuales"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoNumero(g_rst_Princi!ActEco_Com_VtaMen, 15, 2)

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Fecha de Inicio de Operaciones"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Com_IniOpe))
   
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Cargo"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = IIf(g_rst_Princi!ActEco_Com_CodCar = "999999", Trim(g_rst_Princi!ActEco_Com_NomCar & ""), moddat_gf_Consulta_ParDes("503", g_rst_Princi!ActEco_Com_CodCar))
   
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Régimen Tributario"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("215", CStr(g_rst_Princi!ActEco_Com_RegTri))

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Porcentaje Participación"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoNumero(g_rst_Princi!ActEco_Com_PorPar, 7, 2)

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Tipo de Local Comercial"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("208", CStr(g_rst_Princi!ActEco_Com_TipLoc))
   
   If g_rst_Princi!ActEco_Com_TipLoc = 2 Then
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Alquiler Mensual"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = gf_FormatoNumero(g_rst_Princi!ActEco_Com_AlqMen, 15, 2)
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Nombre Arrendador"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Com_NomArr & "")
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Teléfono Arrendador"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Com_TelArr & "")
   End If
End Sub

Private Sub fs_ActEco_Acc(ByVal p_Indice As Integer, ByVal p_ColTxt)
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Documento Identidad"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("203", g_rst_Princi!ActEco_Acc_TipDoc) & " - " & Trim(g_rst_Princi!ActEco_Acc_NumDoc & "")

   'Buscar en Maestro de Empresas
   g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(g_rst_Princi!ActEco_Acc_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & g_rst_Princi!ActEco_Acc_NumDoc & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Razón Social"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Acc_RazSoc & "")
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Nombre Comercial"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Acc_NomCom & "")
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "CIIU"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = g_rst_Princi!ActEco_Acc_CodCiu & " - " & moddat_gf_Consulta_ParDes("102", CStr(g_rst_Princi!ActEco_Acc_CodCiu))
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Dirección"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!ActEco_Acc_TipVia)) & _
                                  " " & Trim(g_rst_Princi!ActEco_Acc_NomVia) & " " & Trim(g_rst_Princi!ActEco_Acc_NumVia) & _
                                  IIf(Len(Trim(g_rst_Princi!ActEco_Acc_IntDpt)) > 0, " (" & Trim(g_rst_Princi!ActEco_Acc_IntDpt) & ")", "") & _
                                  IIf(Len(Trim(g_rst_Princi!ActEco_Acc_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!ActEco_Acc_TipZon)) & " " & Trim(g_rst_Princi!ActEco_Acc_NomZon), "")

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Referencia"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Acc_Refere & "")

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Departamento / Provincia / Distrito"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Acc_UbiGeo, 2) & "0000") & _
                                  " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!ActEco_Acc_UbiGeo, 4) & "00") & _
                                  " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!ActEco_Acc_UbiGeo))
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Teléfono(s)"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Acc_Telef1 & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Acc_Telef2 & "")) > 0, " - " & Trim(g_rst_Princi!ActEco_Dep_Telef2 & ""), "")
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Fax"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Acc_NumFax & "")
   Else
      g_rst_Genera.MoveFirst

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Razón Social"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_RAZSOC & "")

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Razón Social"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_NOMCOM & "")

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "CIIU"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = g_rst_Genera!DATGEN_CODCIU & " - " & moddat_gf_Consulta_ParDes("102", CStr(g_rst_Genera!DATGEN_CODCIU))

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Dirección"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Genera!DatGen_TipVia)) & _
                                  " " & Trim(g_rst_Genera!DatGen_NomVia) & " " & Trim(g_rst_Genera!DatGen_numVia) & _
                                  IIf(Len(Trim(g_rst_Genera!DatGen_IntDpt)) > 0, " (" & Trim(g_rst_Genera!DatGen_IntDpt) & ")", "") & _
                                  IIf(Len(Trim(g_rst_Genera!DatGen_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Genera!DatGen_TipZon)) & " " & Trim(g_rst_Genera!DatGen_NomZon), "")

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Referencia"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DatGen_Refere & "")

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Departamento / Provincia / Distrito"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Genera!DatGen_Ubigeo, 2) & "0000") & _
                                  " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Genera!DatGen_Ubigeo, 4) & "00") & _
                                  " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Genera!DatGen_Ubigeo))
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Teléfono(s)"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DATGEN_TELEF1 & "") & IIf(Len(Trim(g_rst_Genera!DATGEN_TELEF2 & "")) > 0, " - " & Trim(g_rst_Genera!DATGEN_TELEF2 & ""), "")
   
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Fax"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Genera!DatGen_NUMFAX & "")
   End If

   g_rst_Genera.Close
   Set g_rst_Genera = Nothing

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Ingreso Neto"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoNumero(g_rst_Princi!ActEco_Acc_IngNet, 15, 2)

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Fecha de Antigüedad"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Acc_FecAnt))

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Porcentaje Participación"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoNumero(g_rst_Princi!ActEco_Acc_PorPar, 7, 2)
End Sub

Private Sub fs_ActEco_Ren(ByVal p_Indice As Integer, ByVal p_ColTxt)
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Ingreso Neto"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoNumero(g_rst_Princi!ActEco_Ren_IngNet, 15, 2)

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Dirección de Propiedad 01"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Ren_Direc1 & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Nombre de Arrendatario"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Ren_NomAr1 & "")
   
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Fecha de Inicio de Alquiler"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Ren_IniAl1))
   
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Teléfono(s)"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Ren_Tele11 & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Ren_Tele21 & "")) > 0, " - " & Trim(g_rst_Princi!ActEco_Ren_Tele21 & ""), "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Alquiler Mensual"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoNumero(g_rst_Princi!ActEco_Ren_AlqMe1, 15, 2)
   
   If g_rst_Princi!ActEco_Ren_SegPro = 1 Then
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Dirección de Propiedad 02"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Ren_Direc2 & "")

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Nombre de Arrendatario"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Ren_NomAr2 & "")
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Fecha de Inicio de Alquiler"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Ren_IniAl2))
      
      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Teléfono(s)"

      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Ren_Tele12 & "") & IIf(Len(Trim(g_rst_Princi!ActEco_Ren_Tele22 & "")) > 0, " - " & Trim(g_rst_Princi!ActEco_Ren_Tele22 & ""), "")

      grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
      grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
      grd_Listad(p_Indice).Col = 0
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = "Alquiler Mensual"
   
      grd_Listad(p_Indice).Col = 1
      grd_Listad(p_Indice).CellForeColor = p_ColTxt
      grd_Listad(p_Indice).Text = gf_FormatoNumero(g_rst_Princi!ActEco_Ren_AlqMe2, 15, 2)
   End If
End Sub

Private Sub fs_ActEco_Otr(ByVal p_Indice As Integer, ByVal p_ColTxt)
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Ingreso Neto"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = gf_FormatoNumero(g_rst_Princi!ActEco_Otr_IngNet, 15, 2)

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Actividad"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Otr_Activi & "")

   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "CIIU"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = g_rst_Princi!ActEco_Otr_CodCiu & " - " & moddat_gf_Consulta_ParDes("102", CStr(g_rst_Princi!ActEco_Otr_CodCiu))
   
   grd_Listad(p_Indice).Rows = grd_Listad(p_Indice).Rows + 1
   grd_Listad(p_Indice).Row = grd_Listad(p_Indice).Rows - 1
   grd_Listad(p_Indice).Col = 0
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = "Observaciones"

   grd_Listad(p_Indice).Col = 1
   grd_Listad(p_Indice).CellForeColor = p_ColTxt
   grd_Listad(p_Indice).Text = Trim(g_rst_Princi!ActEco_Otr_Observ & "")
End Sub

Private Sub fs_DatPat()
   Dim r_int_Contad     As Integer
   
   Call gs_LimpiaGrid(grd_Listad(4))
   
   'Mostrar Todos los Documentos Recibidos
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = '" & moddat_g_str_NumSol & "' "

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
      
      g_str_Parame = "SELECT * FROM CRE_SOLINB WHERE "
      g_str_Parame = g_str_Parame & "SOLINB_NUMSOL = '" & moddat_g_str_NumSol & "' ORDER BY SOLINB_NUMITE ASC"

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
      
      g_str_Parame = "SELECT * FROM CRE_SOLTRJ WHERE "
      g_str_Parame = g_str_Parame & "SOLTRJ_NUMSOL = '" & moddat_g_str_NumSol & "' ORDER BY SOLTRJ_NUMITE ASC"

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
      
      g_str_Parame = "SELECT * FROM CRE_SOLDEU WHERE "
      g_str_Parame = g_str_Parame & "SOLDEU_NUMSOL = '" & moddat_g_str_NumSol & "' ORDER BY SOLDEU_NUMITE ASC"

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
      
      g_str_Parame = "SELECT * FROM CRE_SOLEYM WHERE "
      g_str_Parame = g_str_Parame & "SOLEYM_NUMSOL = '" & moddat_g_str_NumSol & "' ORDER BY SOLEYM_NUMITE ASC"

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

   g_str_Parame = "SELECT * FROM CRE_SOLREF WHERE "
   g_str_Parame = g_str_Parame & "SOLREF_NUMSOL = '" & moddat_g_str_NumSol & "' "

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
         
         If g_rst_Princi!SOLREF_TIPREF = 1 Then
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

Private Sub fs_DatInm()
   g_str_Parame = "SELECT * FROM CRE_SOLINM WHERE "
   g_str_Parame = g_str_Parame & "SOLINM_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_Listad(2).Redraw = False
      
      g_rst_Princi.MoveFirst
      
      grd_Listad(2).Rows = grd_Listad(2).Rows + 1
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).Text = "Modalidad"
      
      If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), moddat_g_str_CodPrd, moddat_g_str_CodSub, "003", Format(CInt(CStr(g_rst_Princi!SOLINM_CODMOD)), "000")) Then
         grd_Listad(2).Col = 1
         grd_Listad(2).Text = moddat_g_arr_Genera(1).Genera_Nombre
      End If
      
      grd_Listad(2).Rows = grd_Listad(2).Rows + 1
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(2).Text = "Tipo de Inmueble"
         
      grd_Listad(2).Col = 1
      grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(2).Text = moddat_gf_Consulta_ParDes("217", CStr(g_rst_Princi!SOLINM_TIPINM))
      
      grd_Listad(2).Rows = grd_Listad(2).Rows + 1
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(2).Text = "Dirección"
      
      grd_Listad(2).Col = 1
      grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(2).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!SOLINM_TIPVIA)) & _
                        " " & Trim(g_rst_Princi!SOLINM_NOMVIA) & " " & Trim(g_rst_Princi!SOLINM_NUMVIA) & _
                        IIf(Len(Trim(g_rst_Princi!SOLINM_INTDPT)) > 0, " (" & Trim(g_rst_Princi!SOLINM_INTDPT) & ")", "") & _
                        IIf(Len(Trim(g_rst_Princi!SOLINM_NOMZON)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!SOLINM_TIPZON)) & " " & Trim(g_rst_Princi!SOLINM_NOMZON), "")
      
      grd_Listad(2).Rows = grd_Listad(2).Rows + 1
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(2).Text = "Referencia"

      grd_Listad(2).Col = 1
      grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(2).Text = Trim(g_rst_Princi!SOLINM_REFERE & "")
      
      grd_Listad(2).Rows = grd_Listad(2).Rows + 1
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(2).Text = "Estacionamiento"

      grd_Listad(2).Col = 1
      grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(2).Text = Trim(g_rst_Princi!SOLINM_ESTACI & "")
      
      grd_Listad(2).Rows = grd_Listad(2).Rows + 1
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(2).Text = "Departamento / Provincia / Distrito"

      grd_Listad(2).Col = 1
      grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(2).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO, 2) & "0000") & _
                        " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO, 4) & "00") & _
                        " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!SOLINM_UBIGEO))
      
      grd_Listad(2).Rows = grd_Listad(2).Rows + 2
      grd_Listad(2).Row = grd_Listad(2).Rows - 1
      grd_Listad(2).Col = 0
      grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(2).Text = "Proyecto miCasita"

      grd_Listad(2).Col = 1
      grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(2).Text = moddat_gf_Consulta_ParDes("214", g_rst_Princi!SOLINM_PRYMCS)
      
      If g_rst_Princi!SOLINM_TABPRY = 2 Then
         If Not IsNull(g_rst_Princi!SOLINM_PRYBCO) Then
            grd_Listad(2).Rows = grd_Listad(2).Rows + 1
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(2).Text = "Proyecto anclado en Otra IFI"
      
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(2).Text = moddat_gf_Consulta_ParDes("513", g_rst_Princi!SOLINM_PRYBCO)
         End If
         
         If Len(Trim(g_rst_Princi!SOLINM_PRYCOD)) > 0 Then
            grd_Listad(2).Rows = grd_Listad(2).Rows + 1
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(2).Text = "Nombre Proyecto"
   
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(2).Text = moddat_gf_Consulta_NomPry(g_rst_Princi!SOLINM_PRYCOD)
         Else
            If Len(Trim(g_rst_Princi!SOLINM_PRYNOM)) > 0 Then
               grd_Listad(2).Rows = grd_Listad(2).Rows + 1
               grd_Listad(2).Row = grd_Listad(2).Rows - 1
               grd_Listad(2).Col = 0
               grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
               grd_Listad(2).Text = "Nombre Proyecto"
   
               grd_Listad(2).Col = 1
               grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
               grd_Listad(2).Text = Trim(g_rst_Princi!SOLINM_PRYNOM & "")
            End If
         End If
      
         grd_Listad(2).Rows = grd_Listad(2).Rows + 2
         grd_Listad(2).Row = grd_Listad(2).Rows - 1
         grd_Listad(2).Col = 0
         grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(2).Text = "Propietario / Promotor"
   
         grd_Listad(2).Col = 1
         grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(2).Text = moddat_gf_Consulta_ParDes("218", g_rst_Princi!SOLINM_FLGPRO)
         
         grd_Listad(2).Rows = grd_Listad(2).Rows + 1
         grd_Listad(2).Row = grd_Listad(2).Rows - 1
         grd_Listad(2).Col = 0
         grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(2).Text = "Docum. Identidad Propietario/Promotor"
   
         grd_Listad(2).Col = 1
         grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(2).Text = moddat_gf_Consulta_ParDes("236", CStr(g_rst_Princi!SOLINM_TIPDOC_PRO)) & " - " & Trim(g_rst_Princi!SOLINM_NUMDOC_PRO & "")
         
         grd_Listad(2).Rows = grd_Listad(2).Rows + 1
         grd_Listad(2).Row = grd_Listad(2).Rows - 1
         grd_Listad(2).Col = 0
         grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(2).Text = "Nombre o Razón Social"
   
         grd_Listad(2).Col = 1
         grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(2).Text = Trim(g_rst_Princi!SOLINM_RAZSOC_PRO & "")
         
         grd_Listad(2).Rows = grd_Listad(2).Rows + 1
         grd_Listad(2).Row = grd_Listad(2).Rows - 1
         grd_Listad(2).Col = 0
         grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(2).Text = "Dirección"
         
         grd_Listad(2).Col = 1
         grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(2).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!SOLINM_TIPVIA_PRO)) & _
                           " " & Trim(g_rst_Princi!SOLINM_NOMVIA_PRO) & " " & Trim(g_rst_Princi!SOLINM_NUMVIA_PRO) & _
                           IIf(Len(Trim(g_rst_Princi!SOLINM_INTDPT_PRO)) > 0, " (" & Trim(g_rst_Princi!SOLINM_INTDPT_PRO) & ")", "") & _
                           IIf(Len(Trim(g_rst_Princi!SOLINM_NOMZON_PRO)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!SOLINM_TIPZON_PRO)) & " " & Trim(g_rst_Princi!SOLINM_NOMZON_PRO), "")
         
         grd_Listad(2).Rows = grd_Listad(2).Rows + 1
         grd_Listad(2).Row = grd_Listad(2).Rows - 1
         grd_Listad(2).Col = 0
         grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(2).Text = "Referencia"
   
         grd_Listad(2).Col = 1
         grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(2).Text = Trim(g_rst_Princi!SOLINM_REFERE_PRO & "")
         
         grd_Listad(2).Rows = grd_Listad(2).Rows + 1
         grd_Listad(2).Row = grd_Listad(2).Rows - 1
         grd_Listad(2).Col = 0
         grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(2).Text = "Departamento / Provincia / Distrito"
   
         grd_Listad(2).Col = 1
         grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(2).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO_PRO, 2) & "0000") & _
                           " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO_PRO, 4) & "00") & _
                           " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!SOLINM_UBIGEO_PRO))
         
         grd_Listad(2).Rows = grd_Listad(2).Rows + 1
         grd_Listad(2).Row = grd_Listad(2).Rows - 1
         grd_Listad(2).Col = 0
         grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(2).Text = "Teléfono"
   
         grd_Listad(2).Col = 1
         grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(2).Text = Trim(g_rst_Princi!SOLINM_TELEFO_PRO & "")
         
         If g_rst_Princi!SOLINM_FLGCON = 1 Then
            grd_Listad(2).Rows = grd_Listad(2).Rows + 2
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(2).Text = "Docum. Identidad Constructor"
      
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(2).Text = moddat_gf_Consulta_ParDes("236", CStr(g_rst_Princi!SOLINM_TIPDOC_CON)) & " - " & Trim(g_rst_Princi!SOLINM_NUMDOC_CON & "")
            
            grd_Listad(2).Rows = grd_Listad(2).Rows + 1
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(2).Text = "Nombre o Razón Social"
      
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(2).Text = Trim(g_rst_Princi!SOLINM_RAZSOC_CON & "")
            
            grd_Listad(2).Rows = grd_Listad(2).Rows + 1
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(2).Text = "Dirección"
            
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(2).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!SOLINM_TIPVIA_CON)) & _
                              " " & Trim(g_rst_Princi!SOLINM_NOMVIA_CON) & " " & Trim(g_rst_Princi!SOLINM_NUMVIA_CON) & _
                              IIf(Len(Trim(g_rst_Princi!SOLINM_INTDPT_CON)) > 0, " (" & Trim(g_rst_Princi!SOLINM_INTDPT_CON) & ")", "") & _
                              IIf(Len(Trim(g_rst_Princi!SOLINM_NOMZON_CON)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!SOLINM_TIPZON_CON)) & " " & Trim(g_rst_Princi!SOLINM_NOMZON_CON), "")
            
            grd_Listad(2).Rows = grd_Listad(2).Rows + 1
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(2).Text = "Referencia"
      
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(2).Text = Trim(g_rst_Princi!SOLINM_REFERE_CON & "")
            
            grd_Listad(2).Rows = grd_Listad(2).Rows + 1
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(2).Text = "Departamento / Provincia / Distrito"
      
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(2).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO_CON, 2) & "0000") & _
                              " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO_CON, 4) & "00") & _
                              " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!SOLINM_UBIGEO_CON))
            
            grd_Listad(2).Rows = grd_Listad(2).Rows + 1
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(2).Text = "Teléfono"
      
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(2).Text = Trim(g_rst_Princi!SOLINM_TELEFO_CON & "")
         End If
      Else
         If Len(Trim(g_rst_Princi!SOLINM_PRYCOD & "")) > 0 Then
            If g_rst_Princi!SOLINM_PRYMCS = 1 Then
               grd_Listad(2).Rows = grd_Listad(2).Rows + 1
               grd_Listad(2).Row = grd_Listad(2).Rows - 1
               grd_Listad(2).Col = 0
               grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
               grd_Listad(2).Text = "Proyecto Vinculado"
            Else
               grd_Listad(2).Rows = grd_Listad(2).Rows + 1
               grd_Listad(2).Row = grd_Listad(2).Rows - 1
               grd_Listad(2).Col = 0
               grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
               grd_Listad(2).Text = "Entidad Financiera"
         
               grd_Listad(2).Col = 1
               grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
               grd_Listad(2).Text = moddat_gf_Consulta_ParDes("513", g_rst_Princi!SOLINM_PRYBCO)
               
               grd_Listad(2).Rows = grd_Listad(2).Rows + 1
               grd_Listad(2).Row = grd_Listad(2).Rows - 1
               grd_Listad(2).Col = 0
               grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
               grd_Listad(2).Text = "Proyecto No Vinculado"
            End If
         
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColNeg
            grd_Listad(2).Text = moddat_gf_Consulta_NomPry(g_rst_Princi!SOLINM_PRYCOD)
         End If
         
         If CInt(g_rst_Princi!SOLINM_CODMOD) = 1 Or CInt(g_rst_Princi!SOLINM_CODMOD) = 4 Then
            grd_Listad(2).Rows = grd_Listad(2).Rows + 2
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(2).Text = "Docum. Identidad Propietario"
      
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(2).Text = moddat_gf_Consulta_ParDes("236", CStr(g_rst_Princi!SOLINM_TIPDOC_PRO)) & " - " & Trim(g_rst_Princi!SOLINM_NUMDOC_PRO & "")
            
            grd_Listad(2).Rows = grd_Listad(2).Rows + 1
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(2).Text = "Nombre o Razón Social"
      
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(2).Text = Trim(g_rst_Princi!SOLINM_RAZSOC_PRO & "")
            
            grd_Listad(2).Rows = grd_Listad(2).Rows + 1
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(2).Text = "Dirección"
            
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(2).Text = moddat_gf_Consulta_ParDes("201", CStr(g_rst_Princi!SOLINM_TIPVIA_PRO)) & _
                              " " & Trim(g_rst_Princi!SOLINM_NOMVIA_PRO) & " " & Trim(g_rst_Princi!SOLINM_NUMVIA_PRO) & _
                              IIf(Len(Trim(g_rst_Princi!SOLINM_INTDPT_PRO)) > 0, " (" & Trim(g_rst_Princi!SOLINM_INTDPT_PRO) & ")", "") & _
                              IIf(Len(Trim(g_rst_Princi!SOLINM_NOMZON_PRO)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_Princi!SOLINM_TIPZON_PRO)) & " " & Trim(g_rst_Princi!SOLINM_NOMZON_PRO), "")
            
            grd_Listad(2).Rows = grd_Listad(2).Rows + 1
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(2).Text = "Referencia"
      
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(2).Text = Trim(g_rst_Princi!SOLINM_REFERE_PRO & "")
            
            grd_Listad(2).Rows = grd_Listad(2).Rows + 1
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(2).Text = "Departamento / Provincia / Distrito"
      
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(2).Text = moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO_PRO, 2) & "0000") & _
                              " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_Princi!SOLINM_UBIGEO_PRO, 4) & "00") & _
                              " - " & moddat_gf_Consulta_ParDes("101", Trim(g_rst_Princi!SOLINM_UBIGEO_PRO))
            
            grd_Listad(2).Rows = grd_Listad(2).Rows + 1
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(2).Text = "Teléfono"
      
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(2).Text = Trim(g_rst_Princi!SOLINM_TELEFO_PRO & "")
         Else
            'Promotor
            grd_Listad(2).Rows = grd_Listad(2).Rows + 2
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(2).Text = "Doc. Ident. Promotor"
            
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(2).Text = CStr(g_rst_Princi!SOLINM_TIPDOC_PRO) & "-" & Trim(g_rst_Princi!SOLINM_NUMDOC_PRO)
            
            grd_Listad(2).Rows = grd_Listad(2).Rows + 1
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(2).Text = "Razón Social Promotor"
            
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColAzu
            grd_Listad(2).Text = moddat_gf_Consulta_RazSoc(g_rst_Princi!SOLINM_TIPDOC_PRO, g_rst_Princi!SOLINM_NUMDOC_PRO)
            
            'Constructor
            grd_Listad(2).Rows = grd_Listad(2).Rows + 2
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(2).Text = "Doc. Ident. Constructor"
            
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(2).Text = CStr(g_rst_Princi!SOLINM_TIPDOC_CON) & "-" & Trim(g_rst_Princi!SOLINM_NUMDOC_CON)
            
            grd_Listad(2).Rows = grd_Listad(2).Rows + 1
            grd_Listad(2).Row = grd_Listad(2).Rows - 1
            grd_Listad(2).Col = 0
            grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(2).Text = "Razón Social Constructor"
            
            grd_Listad(2).Col = 1
            grd_Listad(2).CellForeColor = modgen_g_con_ColRoj
            grd_Listad(2).Text = moddat_gf_Consulta_RazSoc(g_rst_Princi!SOLINM_TIPDOC_CON, g_rst_Princi!SOLINM_NUMDOC_CON)
         End If
      End If
      
      grd_Listad(2).Redraw = True
      
      Call gs_UbiIniGrid(grd_Listad(2))
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_DatCre()
   Call gs_LimpiaGrid(grd_Listad(5))
   
   'Mostrar Todos los Documentos Recibidos
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = '" & moddat_g_str_NumSol & "' "

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
   
   grd_Listad(5).Rows = grd_Listad(5).Rows + 1
   grd_Listad(5).Row = grd_Listad(5).Rows - 1
   grd_Listad(5).Col = 0
   grd_Listad(5).Text = "Valor de Compra Venta (US$)"

   grd_Listad(5).Col = 1
   grd_Listad(5).CellFontName = "Lucida Console"
   grd_Listad(5).CellFontSize = 8
   grd_Listad(5).Text = gf_FormatoNumero(g_rst_Princi!SOLMAE_COMVTA_DOL, 12, 2)

   grd_Listad(5).Rows = grd_Listad(5).Rows + 1
   grd_Listad(5).Row = grd_Listad(5).Rows - 1
   grd_Listad(5).Col = 0
   grd_Listad(5).Text = "Aporte Propio (US$)"

   grd_Listad(5).Col = 1
   grd_Listad(5).CellFontName = "Lucida Console"
   grd_Listad(5).CellFontSize = 8
   grd_Listad(5).Text = gf_FormatoNumero(g_rst_Princi!SOLMAE_APOPRO_DOL, 12, 2)

   grd_Listad(5).Rows = grd_Listad(5).Rows + 1
   grd_Listad(5).Row = grd_Listad(5).Rows - 1
   grd_Listad(5).Col = 0
   grd_Listad(5).Text = "Monto Préstamo (US$)"

   grd_Listad(5).Col = 1
   grd_Listad(5).CellFontName = "Lucida Console"
   grd_Listad(5).CellFontSize = 8
   grd_Listad(5).Text = gf_FormatoNumero(g_rst_Princi!SOLMAE_MTOPRE_DOL, 12, 2)

   grd_Listad(5).Rows = grd_Listad(5).Rows + 1
   grd_Listad(5).Row = grd_Listad(5).Rows - 1
   grd_Listad(5).Col = 0
   grd_Listad(5).Text = "Valor de Compra Venta (S/.)"

   grd_Listad(5).Col = 1
   grd_Listad(5).CellFontName = "Lucida Console"
   grd_Listad(5).CellFontSize = 8
   grd_Listad(5).Text = gf_FormatoNumero(g_rst_Princi!SOLMAE_COMVTA_SOL, 12, 2)

   grd_Listad(5).Rows = grd_Listad(5).Rows + 1
   grd_Listad(5).Row = grd_Listad(5).Rows - 1
   grd_Listad(5).Col = 0
   grd_Listad(5).Text = "Aporte Propio (S/.)"

   grd_Listad(5).Col = 1
   grd_Listad(5).CellFontName = "Lucida Console"
   grd_Listad(5).CellFontSize = 8
   grd_Listad(5).Text = gf_FormatoNumero(g_rst_Princi!SOLMAE_APOPRO_SOL, 12, 2)

   grd_Listad(5).Rows = grd_Listad(5).Rows + 1
   grd_Listad(5).Row = grd_Listad(5).Rows - 1
   grd_Listad(5).Col = 0
   grd_Listad(5).Text = "Monto Préstamo (S/.)"

   grd_Listad(5).Col = 1
   grd_Listad(5).CellFontName = "Lucida Console"
   grd_Listad(5).CellFontSize = 8
   grd_Listad(5).Text = gf_FormatoNumero(g_rst_Princi!SOLMAE_MTOPRE_SOL, 12, 2)

   grd_Listad(5).Rows = grd_Listad(5).Rows + 1
   grd_Listad(5).Row = grd_Listad(5).Rows - 1
   grd_Listad(5).Col = 0
   grd_Listad(5).Text = "Tipo de Cambio Referencial"

   grd_Listad(5).Col = 1
   grd_Listad(5).CellFontName = "Lucida Console"
   grd_Listad(5).CellFontSize = 8
   grd_Listad(5).Text = gf_FormatoNumero(g_rst_Princi!SOLMAE_MTOPRE_SOL / g_rst_Princi!SOLMAE_MTOPRE_DOL, 12, 4)

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
   grd_Listad(5).Text = gf_FormatoNumero(g_rst_Princi!SOLMAE_INTGRA, 12, 2)

   grd_Listad(5).Rows = grd_Listad(5).Rows + 1
   grd_Listad(5).Row = grd_Listad(5).Rows - 1
   grd_Listad(5).Col = 0
   grd_Listad(5).Text = "Cuotas Extraordinarias"

   grd_Listad(5).Col = 1
   grd_Listad(5).Text = moddat_gf_Consulta_ParDes("214", CStr(g_rst_Princi!SOLMAE_CUOEXT))
   
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
      grd_Listad(5).Text = "Moneda de Ahorro"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).Text = moddat_gf_Consulta_ParDes("204", g_rst_Princi!SOLMAE_MONAHO)
      
      grd_Listad(5).Rows = grd_Listad(5).Rows + 1
      grd_Listad(5).Row = grd_Listad(5).Rows - 1
      grd_Listad(5).Col = 0
      grd_Listad(5).Text = "Monto Mínimo de Ahorro Mensual"
   
      grd_Listad(5).Col = 1
      grd_Listad(5).CellFontName = "Lucida Console"
      grd_Listad(5).CellFontSize = 8
      grd_Listad(5).Text = gf_FormatoNumero(g_rst_Princi!SOLMAE_MTOAHO, 12, 2)
   
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
   
   moddat_g_str_CodEjeSeg = Trim(g_rst_Princi!SOLMAE_EJESEG & "")
   moddat_g_str_CodConHip = Trim(g_rst_Princi!SOLMAE_CONHIP & "")
   
   
   grd_Listad(5).Redraw = True
   Call gs_UbiIniGrid(grd_Listad(5))
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_SolDoc()
   Call gs_LimpiaGrid(grd_Listad(10))
   
   'Mostrar Todos los Documentos Recibidos
   g_str_Parame = "SELECT * FROM CRE_SOLDOC WHERE "
   g_str_Parame = g_str_Parame & "SOLDOC_NUMSOL = '" & moddat_g_str_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "(SOLDOC_TIPDOC = 1 OR SOLDOC_TIPDOC = 2)"

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
   g_str_Parame = "SELECT * FROM CRE_SOLDOC WHERE "
   g_str_Parame = g_str_Parame & "SOLDOC_NUMSOL = '" & moddat_g_str_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "SOLDOC_TIPDOC = 3 "

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

Private Sub fs_Envia_CorEle(ByVal p_Asunto As String, ByVal p_Mensaje As String)
   Dim r_str_Cadena     As String
   
   'Destinatarios de Correo
   ReDim moddat_g_arr_Genera(0)
   
   'Usuario de Seguimiento
   ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
   moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = moddat_gf_Buscar_DirEle_Codigo(moddat_g_str_CodEjeSeg)
   
   'Jefe de Seguimiento
   r_str_Cadena = moddat_gf_Buscar_DirEle_TipUsu(130)
   If Not moddat_gf_Verifica_DirEle(moddat_g_arr_Genera, r_str_Cadena) Then
      ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
      moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = r_str_Cadena
   End If
   
   'Jefe de Ventas
   r_str_Cadena = moddat_gf_Buscar_DirEle_TipUsu(120)
   If Not moddat_gf_Verifica_DirEle(moddat_g_arr_Genera, r_str_Cadena) Then
      ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
      moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = r_str_Cadena
   End If
   
   'Director Comercial
   r_str_Cadena = moddat_gf_Buscar_DirEle_TipUsu(100)
   If Not moddat_gf_Verifica_DirEle(moddat_g_arr_Genera, r_str_Cadena) Then
      ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
      moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = r_str_Cadena
   End If
   
   'Jefe de Operaciones
   r_str_Cadena = moddat_gf_Buscar_DirEle_TipUsu(220)
   If Not moddat_gf_Verifica_DirEle(moddat_g_arr_Genera, r_str_Cadena) Then
      ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
      moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = r_str_Cadena
   End If
   
   'Director de Producción
   r_str_Cadena = moddat_gf_Buscar_DirEle_TipUsu(200)
   If Not moddat_gf_Verifica_DirEle(moddat_g_arr_Genera, r_str_Cadena) Then
      ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
      moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = r_str_Cadena
   End If
   
   Call moddat_gs_EnvCor(mps_Sesion, mps_Mensaj, moddat_g_arr_Genera, modgen_g_str_Mail_Asunto, modgen_g_str_Mail_Mensaj)
End Sub

Private Sub fs_DatTas()
   Call gs_LimpiaGrid(grd_Listad(7))
   
   g_str_Parame = "SELECT * FROM TRA_EVATAS WHERE "
   g_str_Parame = g_str_Parame & "EVATAS_NUMSOL = '" & moddat_g_str_NumSol & "' "

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
      
      grd_Listad(7).Rows = grd_Listad(7).Rows + 1
      grd_Listad(7).Row = grd_Listad(7).Rows - 1
      grd_Listad(7).Col = 0
      grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(7).Text = "Tipo de Moneda"
      
      grd_Listad(7).Col = 1
      grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(7).Text = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!EVATAS_TIPMON))
      
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
      grd_Listad(7).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARETER_INM + g_rst_Princi!EVATAS_ARETER_ES1 + g_rst_Princi!EVATAS_ARETER_ES2 + g_rst_Princi!EVATAS_ARETER_DEP, 12, 2)
      
      grd_Listad(7).Rows = grd_Listad(7).Rows + 1
      grd_Listad(7).Row = grd_Listad(7).Rows - 1
      grd_Listad(7).Col = 0
      grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(7).Text = "Area Construida (Total)"
      
      grd_Listad(7).Col = 1
      grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(7).CellFontName = "Lucida Console"
      grd_Listad(7).CellFontSize = 8
      grd_Listad(7).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARECON_INM + g_rst_Princi!EVATAS_ARECON_ES1 + g_rst_Princi!EVATAS_ARECON_ES2 + g_rst_Princi!EVATAS_ARECON_DEP, 12, 2)
      
      grd_Listad(7).Rows = grd_Listad(7).Rows + 1
      grd_Listad(7).Row = grd_Listad(7).Rows - 1
      grd_Listad(7).Col = 0
      grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(7).Text = "Suma Asegurada (Total)"
      
      grd_Listad(7).Col = 1
      grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(7).CellFontName = "Lucida Console"
      grd_Listad(7).CellFontSize = 8
      grd_Listad(7).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_SUMASE_INM + g_rst_Princi!EVATAS_SUMASE_ES1 + g_rst_Princi!EVATAS_SUMASE_ES2 + g_rst_Princi!EVATAS_SUMASE_DEP, 12, 2)
      
      grd_Listad(7).Rows = grd_Listad(7).Rows + 1
      grd_Listad(7).Row = grd_Listad(7).Rows - 1
      grd_Listad(7).Col = 0
      grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(7).Text = "Valor Comercial (Total)"
      
      grd_Listad(7).Col = 1
      grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(7).CellFontName = "Lucida Console"
      grd_Listad(7).CellFontSize = 8
      grd_Listad(7).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_VALCOM_INM + g_rst_Princi!EVATAS_VALCOM_ES1 + g_rst_Princi!EVATAS_VALCOM_ES2 + g_rst_Princi!EVATAS_VALCOM_DEP, 12, 2)
      
      grd_Listad(7).Rows = grd_Listad(7).Rows + 1
      grd_Listad(7).Row = grd_Listad(7).Rows - 1
      grd_Listad(7).Col = 0
      grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(7).Text = "Valor Realización (Total)"
      
      grd_Listad(7).Col = 1
      grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(7).CellFontName = "Lucida Console"
      grd_Listad(7).CellFontSize = 8
      grd_Listad(7).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_VALREA_INM + g_rst_Princi!EVATAS_VALREA_ES1 + g_rst_Princi!EVATAS_VALREA_ES2 + g_rst_Princi!EVATAS_VALREA_DEP, 12, 2)
      
      grd_Listad(7).Rows = grd_Listad(7).Rows + 1
      grd_Listad(7).Row = grd_Listad(7).Rows - 1
      grd_Listad(7).Col = 0
      grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(7).Text = "Valor Terreno (Total)"
      
      grd_Listad(7).Col = 1
      grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(7).CellFontName = "Lucida Console"
      grd_Listad(7).CellFontSize = 8
      grd_Listad(7).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_VALTER_INM + g_rst_Princi!EVATAS_VALTER_ES1 + g_rst_Princi!EVATAS_VALTER_ES2 + g_rst_Princi!EVATAS_VALTER_DEP, 12, 2)
      
      grd_Listad(7).Rows = grd_Listad(7).Rows + 1
      grd_Listad(7).Row = grd_Listad(7).Rows - 1
      grd_Listad(7).Col = 0
      grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(7).Text = "Valor Edificación (Total)"
      
      grd_Listad(7).Col = 1
      grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(7).CellFontName = "Lucida Console"
      grd_Listad(7).CellFontSize = 8
      grd_Listad(7).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_VALEDI_INM + g_rst_Princi!EVATAS_VALEDI_ES1 + g_rst_Princi!EVATAS_VALEDI_ES2 + g_rst_Princi!EVATAS_VALEDI_DEP, 12, 2)
   
      grd_Listad(7).Rows = grd_Listad(7).Rows + 1
      grd_Listad(7).Row = grd_Listad(7).Rows - 1
      grd_Listad(7).Col = 0
      grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(7).Text = "Valor Areas Comunes (Total)"
      
      grd_Listad(7).Col = 1
      grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
      grd_Listad(7).CellFontName = "Lucida Console"
      grd_Listad(7).CellFontSize = 8
      grd_Listad(7).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_VALACO_INM + g_rst_Princi!EVATAS_VALACO_ES1 + g_rst_Princi!EVATAS_VALACO_ES2 + g_rst_Princi!EVATAS_VALACO_DEP, 12, 2)
   
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
      grd_Listad(7).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARETER_INM, 12, 2)
      
      grd_Listad(7).Rows = grd_Listad(7).Rows + 1
      grd_Listad(7).Row = grd_Listad(7).Rows - 1
      grd_Listad(7).Col = 0
      grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(7).Text = "Area Construida (Inmueble)"
      
      grd_Listad(7).Col = 1
      grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(7).CellFontName = "Lucida Console"
      grd_Listad(7).CellFontSize = 8
      grd_Listad(7).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARECON_INM, 12, 2)
      
      grd_Listad(7).Rows = grd_Listad(7).Rows + 1
      grd_Listad(7).Row = grd_Listad(7).Rows - 1
      grd_Listad(7).Col = 0
      grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(7).Text = "Suma Asegurada (Inmueble)"
      
      grd_Listad(7).Col = 1
      grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(7).CellFontName = "Lucida Console"
      grd_Listad(7).CellFontSize = 8
      grd_Listad(7).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_SUMASE_INM, 12, 2)
      
      grd_Listad(7).Rows = grd_Listad(7).Rows + 1
      grd_Listad(7).Row = grd_Listad(7).Rows - 1
      grd_Listad(7).Col = 0
      grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(7).Text = "Valor Comercial (Inmueble)"
      
      grd_Listad(7).Col = 1
      grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(7).CellFontName = "Lucida Console"
      grd_Listad(7).CellFontSize = 8
      grd_Listad(7).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_VALCOM_INM, 12, 2)
      
      grd_Listad(7).Rows = grd_Listad(7).Rows + 1
      grd_Listad(7).Row = grd_Listad(7).Rows - 1
      grd_Listad(7).Col = 0
      grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(7).Text = "Valor Realización (Inmueble)"
      
      grd_Listad(7).Col = 1
      grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(7).CellFontName = "Lucida Console"
      grd_Listad(7).CellFontSize = 8
      grd_Listad(7).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_VALREA_INM, 12, 2)
      
      grd_Listad(7).Rows = grd_Listad(7).Rows + 1
      grd_Listad(7).Row = grd_Listad(7).Rows - 1
      grd_Listad(7).Col = 0
      grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(7).Text = "Valor Terreno (Inmueble)"
      
      grd_Listad(7).Col = 1
      grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(7).CellFontName = "Lucida Console"
      grd_Listad(7).CellFontSize = 8
      grd_Listad(7).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_VALTER_INM, 12, 2)
      
      grd_Listad(7).Rows = grd_Listad(7).Rows + 1
      grd_Listad(7).Row = grd_Listad(7).Rows - 1
      grd_Listad(7).Col = 0
      grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(7).Text = "Valor Edificación (Inmueble)"
      
      grd_Listad(7).Col = 1
      grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(7).CellFontName = "Lucida Console"
      grd_Listad(7).CellFontSize = 8
      grd_Listad(7).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_VALEDI_INM, 12, 2)
   
      grd_Listad(7).Rows = grd_Listad(7).Rows + 1
      grd_Listad(7).Row = grd_Listad(7).Rows - 1
      grd_Listad(7).Col = 0
      grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(7).Text = "Valor Areas Comunes (Inmueble)"
      
      grd_Listad(7).Col = 1
      grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
      grd_Listad(7).CellFontName = "Lucida Console"
      grd_Listad(7).CellFontSize = 8
      grd_Listad(7).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_VALACO_INM, 12, 2)
   
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
         grd_Listad(7).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARETER_ES1, 12, 2)
         
         grd_Listad(7).Rows = grd_Listad(7).Rows + 1
         grd_Listad(7).Row = grd_Listad(7).Rows - 1
         grd_Listad(7).Col = 0
         grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(7).Text = "Area Construida (Estac. 1)"
         
         grd_Listad(7).Col = 1
         grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(7).CellFontName = "Lucida Console"
         grd_Listad(7).CellFontSize = 8
         grd_Listad(7).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARECON_ES1, 12, 2)
         
         grd_Listad(7).Rows = grd_Listad(7).Rows + 1
         grd_Listad(7).Row = grd_Listad(7).Rows - 1
         grd_Listad(7).Col = 0
         grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(7).Text = "Suma Asegurada (Estac. 1)"
         
         grd_Listad(7).Col = 1
         grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(7).CellFontName = "Lucida Console"
         grd_Listad(7).CellFontSize = 8
         grd_Listad(7).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_SUMASE_ES1, 12, 2)
         
         grd_Listad(7).Rows = grd_Listad(7).Rows + 1
         grd_Listad(7).Row = grd_Listad(7).Rows - 1
         grd_Listad(7).Col = 0
         grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(7).Text = "Valor Comercial (Estac. 1)"
         
         grd_Listad(7).Col = 1
         grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(7).CellFontName = "Lucida Console"
         grd_Listad(7).CellFontSize = 8
         grd_Listad(7).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_VALCOM_ES1, 12, 2)
         
         grd_Listad(7).Rows = grd_Listad(7).Rows + 1
         grd_Listad(7).Row = grd_Listad(7).Rows - 1
         grd_Listad(7).Col = 0
         grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(7).Text = "Valor Realización (Estac. 1)"
         
         grd_Listad(7).Col = 1
         grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(7).CellFontName = "Lucida Console"
         grd_Listad(7).CellFontSize = 8
         grd_Listad(7).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_VALREA_ES1, 12, 2)
         
         grd_Listad(7).Rows = grd_Listad(7).Rows + 1
         grd_Listad(7).Row = grd_Listad(7).Rows - 1
         grd_Listad(7).Col = 0
         grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(7).Text = "Valor Terreno (Estac. 1)"
         
         grd_Listad(7).Col = 1
         grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(7).CellFontName = "Lucida Console"
         grd_Listad(7).CellFontSize = 8
         grd_Listad(7).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_VALTER_ES1, 12, 2)
         
         grd_Listad(7).Rows = grd_Listad(7).Rows + 1
         grd_Listad(7).Row = grd_Listad(7).Rows - 1
         grd_Listad(7).Col = 0
         grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(7).Text = "Valor Edificación (Estac. 1)"
         
         grd_Listad(7).Col = 1
         grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(7).CellFontName = "Lucida Console"
         grd_Listad(7).CellFontSize = 8
         grd_Listad(7).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_VALEDI_ES1, 12, 2)
      
         grd_Listad(7).Rows = grd_Listad(7).Rows + 1
         grd_Listad(7).Row = grd_Listad(7).Rows - 1
         grd_Listad(7).Col = 0
         grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(7).Text = "Valor Areas Comunes (Estac. 1)"
         
         grd_Listad(7).Col = 1
         grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(7).CellFontName = "Lucida Console"
         grd_Listad(7).CellFontSize = 8
         grd_Listad(7).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_VALACO_ES1, 12, 2)
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
         grd_Listad(7).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARETER_ES2, 12, 2)
         
         grd_Listad(7).Rows = grd_Listad(7).Rows + 1
         grd_Listad(7).Row = grd_Listad(7).Rows - 1
         grd_Listad(7).Col = 0
         grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(7).Text = "Area Construida (Estac. 2)"
         
         grd_Listad(7).Col = 1
         grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(7).CellFontName = "Lucida Console"
         grd_Listad(7).CellFontSize = 8
         grd_Listad(7).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARECON_ES2, 12, 2)
         
         grd_Listad(7).Rows = grd_Listad(7).Rows + 1
         grd_Listad(7).Row = grd_Listad(7).Rows - 1
         grd_Listad(7).Col = 0
         grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(7).Text = "Suma Asegurada (Estac. 2)"
         
         grd_Listad(7).Col = 1
         grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(7).CellFontName = "Lucida Console"
         grd_Listad(7).CellFontSize = 8
         grd_Listad(7).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_SUMASE_ES2, 12, 2)
         
         grd_Listad(7).Rows = grd_Listad(7).Rows + 1
         grd_Listad(7).Row = grd_Listad(7).Rows - 1
         grd_Listad(7).Col = 0
         grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(7).Text = "Valor Comercial (Estac. 2)"
         
         grd_Listad(7).Col = 1
         grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(7).CellFontName = "Lucida Console"
         grd_Listad(7).CellFontSize = 8
         grd_Listad(7).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_VALCOM_ES2, 12, 2)
         
         grd_Listad(7).Rows = grd_Listad(7).Rows + 1
         grd_Listad(7).Row = grd_Listad(7).Rows - 1
         grd_Listad(7).Col = 0
         grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(7).Text = "Valor Realización (Estac. 2)"
         
         grd_Listad(7).Col = 1
         grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(7).CellFontName = "Lucida Console"
         grd_Listad(7).CellFontSize = 8
         grd_Listad(7).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_VALREA_ES2, 12, 2)
         
         grd_Listad(7).Rows = grd_Listad(7).Rows + 1
         grd_Listad(7).Row = grd_Listad(7).Rows - 1
         grd_Listad(7).Col = 0
         grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(7).Text = "Valor Terreno (Estac. 2)"
         
         grd_Listad(7).Col = 1
         grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(7).CellFontName = "Lucida Console"
         grd_Listad(7).CellFontSize = 8
         grd_Listad(7).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_VALTER_ES2, 12, 2)
         
         grd_Listad(7).Rows = grd_Listad(7).Rows + 1
         grd_Listad(7).Row = grd_Listad(7).Rows - 1
         grd_Listad(7).Col = 0
         grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(7).Text = "Valor Edificación (Estac. 2)"
         
         grd_Listad(7).Col = 1
         grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(7).CellFontName = "Lucida Console"
         grd_Listad(7).CellFontSize = 8
         grd_Listad(7).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_VALEDI_ES2, 12, 2)
      
         grd_Listad(7).Rows = grd_Listad(7).Rows + 1
         grd_Listad(7).Row = grd_Listad(7).Rows - 1
         grd_Listad(7).Col = 0
         grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(7).Text = "Valor Areas Comunes (Estac. 2)"
         
         grd_Listad(7).Col = 1
         grd_Listad(7).CellForeColor = modgen_g_con_ColNeg
         grd_Listad(7).CellFontName = "Lucida Console"
         grd_Listad(7).CellFontSize = 8
         grd_Listad(7).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_VALACO_ES2, 12, 2)
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
         grd_Listad(7).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARETER_DEP, 12, 2)
         
         grd_Listad(7).Rows = grd_Listad(7).Rows + 1
         grd_Listad(7).Row = grd_Listad(7).Rows - 1
         grd_Listad(7).Col = 0
         grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(7).Text = "Area Construida (Depósito)"
         
         grd_Listad(7).Col = 1
         grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(7).CellFontName = "Lucida Console"
         grd_Listad(7).CellFontSize = 8
         grd_Listad(7).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_ARECON_DEP, 12, 2)
         
         grd_Listad(7).Rows = grd_Listad(7).Rows + 1
         grd_Listad(7).Row = grd_Listad(7).Rows - 1
         grd_Listad(7).Col = 0
         grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(7).Text = "Suma Asegurada (Depósito)"
         
         grd_Listad(7).Col = 1
         grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(7).CellFontName = "Lucida Console"
         grd_Listad(7).CellFontSize = 8
         grd_Listad(7).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_SUMASE_DEP, 12, 2)
         
         grd_Listad(7).Rows = grd_Listad(7).Rows + 1
         grd_Listad(7).Row = grd_Listad(7).Rows - 1
         grd_Listad(7).Col = 0
         grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(7).Text = "Valor Comercial (Depósito)"
         
         grd_Listad(7).Col = 1
         grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(7).CellFontName = "Lucida Console"
         grd_Listad(7).CellFontSize = 8
         grd_Listad(7).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_VALCOM_DEP, 12, 2)
         
         grd_Listad(7).Rows = grd_Listad(7).Rows + 1
         grd_Listad(7).Row = grd_Listad(7).Rows - 1
         grd_Listad(7).Col = 0
         grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(7).Text = "Valor Realización (Depósito)"
         
         grd_Listad(7).Col = 1
         grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(7).CellFontName = "Lucida Console"
         grd_Listad(7).CellFontSize = 8
         grd_Listad(7).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_VALREA_DEP, 12, 2)
         
         grd_Listad(7).Rows = grd_Listad(7).Rows + 1
         grd_Listad(7).Row = grd_Listad(7).Rows - 1
         grd_Listad(7).Col = 0
         grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(7).Text = "Valor Terreno (Depósito)"
         
         grd_Listad(7).Col = 1
         grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(7).CellFontName = "Lucida Console"
         grd_Listad(7).CellFontSize = 8
         grd_Listad(7).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_VALTER_DEP, 12, 2)
         
         grd_Listad(7).Rows = grd_Listad(7).Rows + 1
         grd_Listad(7).Row = grd_Listad(7).Rows - 1
         grd_Listad(7).Col = 0
         grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(7).Text = "Valor Edificación (Depósito)"
         
         grd_Listad(7).Col = 1
         grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(7).CellFontName = "Lucida Console"
         grd_Listad(7).CellFontSize = 8
         grd_Listad(7).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_VALEDI_DEP, 12, 2)
      
         grd_Listad(7).Rows = grd_Listad(7).Rows + 1
         grd_Listad(7).Row = grd_Listad(7).Rows - 1
         grd_Listad(7).Col = 0
         grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(7).Text = "Valor Areas Comunes (Depósito)"
         
         grd_Listad(7).Col = 1
         grd_Listad(7).CellForeColor = modgen_g_con_ColAzu
         grd_Listad(7).CellFontName = "Lucida Console"
         grd_Listad(7).CellFontSize = 8
         grd_Listad(7).Text = gf_FormatoNumero(g_rst_Princi!EVATAS_VALACO_DEP, 12, 2)
      End If
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_DatSeg()
   Call gs_LimpiaGrid(grd_Listad(8))
   
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
      
      Call gs_UbiIniGrid(grd_Listad(8))
   End If
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub grd_Listad_SelChange(index As Integer)
   If grd_Listad(index).Rows > 2 Then
      grd_Listad(index).RowSel = grd_Listad(index).Row
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

Private Sub fs_DatLeg()
   Call gs_LimpiaGrid(grd_Listad(9))

   g_str_Parame = "SELECT * FROM TRA_EVALEG WHERE "
   g_str_Parame = g_str_Parame & "EVALEG_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      txt_InfLeg.Text = Trim(g_rst_Princi!EVALEG_INFLG1 & "") & Trim(g_rst_Princi!EVALEG_INFLG2 & "") & Trim(g_rst_Princi!EVALEG_INFLG3 & "") & Trim(g_rst_Princi!EVALEG_INFLG4 & "")
      
      If g_rst_Princi!EVALEG_FECCOM > 0 Then
         txt_ComCre.Text = "Fecha de Comité de Créditos: " & gf_FormatoFecha(CStr(g_rst_Princi!EVALEG_FECCOM)) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Trim(g_rst_Princi!EVALEG_OBSCOM & "")
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
         
         Call gs_UbiIniGrid(grd_Listad(9))
      
         If Not IsNull(g_rst_Princi!EVALEG_TCASBS) Then
            grd_Listad(9).Rows = grd_Listad(9).Rows + 1
            grd_Listad(9).Row = grd_Listad(9).Rows - 1
            grd_Listad(9).Col = 0
            grd_Listad(9).Text = "Tipo de Cambio SBS"
            
            grd_Listad(9).Col = 1
            grd_Listad(9).Text = Format(g_rst_Princi!EVALEG_TCASBS, "###,##0.0000")
            
            
            grd_Listad(9).Rows = grd_Listad(9).Rows + 1
            grd_Listad(9).Row = grd_Listad(9).Rows - 1
            grd_Listad(9).Col = 0
            grd_Listad(9).Text = "Moneda Hipoteca"
            
            grd_Listad(9).Col = 1
            grd_Listad(9).Text = moddat_gf_Consulta_ParDes("204", g_rst_Princi!EVALEG_MONHIP)
            
            grd_Listad(9).Rows = grd_Listad(9).Rows + 1
            grd_Listad(9).Row = grd_Listad(9).Rows - 1
            grd_Listad(9).Col = 0
            grd_Listad(9).Text = "Monto Hipoteca"
            
            grd_Listad(9).Col = 1
            grd_Listad(9).Text = Format(g_rst_Princi!EVALEG_MTOHIP, "###,##0.0000")
            
            Call gs_UbiIniGrid(grd_Listad(9))
         End If
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
            Case 1
               grd_Listad(9).Text = grd_Listad(9).Text & " NRO. " & Trim(g_rst_Princi!EVALEG_NUMPAR_INM & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!EVALEG_NUMAPA_INM & "")
               
            Case 2
               grd_Listad(9).Text = grd_Listad(9).Text & " NRO. " & Trim(g_rst_Princi!EVALEG_NUMFIC_INM & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!EVALEG_NUMAFI_INM & "")
               
            Case 3
               grd_Listad(9).Text = grd_Listad(9).Text & " (" & Trim(g_rst_Princi!EVALEG_NUMTOM_INM & "") & " / " & Trim(g_rst_Princi!EVALEG_NUMFOJ_INM & "") & " / " & Trim(g_rst_Princi!EVALEG_NUMLIB_INM & "") & ")"
               
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
               Case 1
                  grd_Listad(9).Text = grd_Listad(9).Text & " NRO. " & Trim(g_rst_Princi!EVALEG_NUMPAR_ES1 & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!EVALEG_NUMAPA_ES1 & "")
                  
               Case 2
                  grd_Listad(9).Text = grd_Listad(9).Text & " NRO. " & Trim(g_rst_Princi!EVALEG_NUMFIC_ES1 & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!EVALEG_NUMAFI_ES1 & "")
                  
               Case 3
                  grd_Listad(9).Text = grd_Listad(9).Text & " (" & Trim(g_rst_Princi!EVALEG_NUMTOM_ES1 & "") & " / " & Trim(g_rst_Princi!EVALEG_NUMFOJ_ES1 & "") & " / " & Trim(g_rst_Princi!EVALEG_NUMLIB_ES1 & "") & ")"
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
               Case 1
                  grd_Listad(9).Text = grd_Listad(9).Text & " NRO. " & Trim(g_rst_Princi!EVALEG_NUMPAR_ES2 & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!EVALEG_NUMAPA_ES2 & "")
                  
               Case 2
                  grd_Listad(9).Text = grd_Listad(9).Text & " NRO. " & Trim(g_rst_Princi!EVALEG_NUMFIC_ES2 & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!EVALEG_NUMAFI_ES2 & "")
                  
               Case 3
                  grd_Listad(9).Text = grd_Listad(9).Text & " (" & Trim(g_rst_Princi!EVALEG_NUMTOM_ES2 & "") & " / " & Trim(g_rst_Princi!EVALEG_NUMFOJ_ES2 & "") & " / " & Trim(g_rst_Princi!EVALEG_NUMLIB_ES2 & "") & ")"
                  
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
               Case 1
                  grd_Listad(9).Text = grd_Listad(9).Text & " NRO. " & Trim(g_rst_Princi!EVALEG_NUMPAR_DEP & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!EVALEG_NUMAPA_DEP & "")
                  
               Case 2
                  grd_Listad(9).Text = grd_Listad(9).Text & " NRO. " & Trim(g_rst_Princi!EVALEG_NUMFIC_DEP & "") & " - ASIENTO NRO. " & Trim(g_rst_Princi!EVALEG_NUMAFI_DEP & "")
                  
               Case 3
                  grd_Listad(9).Text = grd_Listad(9).Text & " (" & Trim(g_rst_Princi!EVALEG_NUMTOM_DEP & "") & " / " & Trim(g_rst_Princi!EVALEG_NUMFOJ_DEP & "") & " / " & Trim(g_rst_Princi!EVALEG_NUMLIB_DEP & "") & ")"
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
   
   r_dbl_Import = 0
   
   Call gs_LimpiaGrid(grd_GasAdm)
   pnl_TotGas.Caption = "0.00 "
   
   g_str_Parame = "SELECT * FROM TRA_GASADM WHERE "
   g_str_Parame = g_str_Parame & "GASADM_NUMSOL = '" & moddat_g_str_NumSol & "' "

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
      grd_GasAdm.Text = Format(g_rst_Princi!GASADM_IMPORT, "###,###,##0.00")
      
      r_dbl_Import = r_dbl_Import + g_rst_Princi!GASADM_IMPORT
      
      'Situación
      grd_GasAdm.Col = 3
      grd_GasAdm.Text = moddat_gf_Consulta_ParDes("001", CStr(g_rst_Princi!GASADM_SITUAC))
      
      'Fecha de Pago
      grd_GasAdm.Col = 4
      If g_rst_Princi!GASADM_PAGFEC > 0 Then
         grd_GasAdm.Text = gf_FormatoFecha(CStr(g_rst_Princi!GASADM_PAGFEC))
      End If
      
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
   Loop
   
   pnl_TotGas.Caption = Format(r_dbl_Import, "###,###,##0.00") & " "
   
   grd_GasAdm.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call gs_UbiIniGrid(grd_GasAdm)
End Sub

Private Sub fs_EvaCre()
   Call gs_LimpiaGrid(grd_Listad(6))
   
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      grd_Listad(6).Rows = grd_Listad(6).Rows + 1
      grd_Listad(6).Row = grd_Listad(6).Rows - 1
      grd_Listad(6).Col = 0
      grd_Listad(6).Text = "Cuota Aceptada por Cliente (S/.)"
      
      grd_Listad(6).Col = 1
      grd_Listad(6).CellFontName = "Lucida Console"
      grd_Listad(6).CellFontSize = 8
      grd_Listad(6).Text = gf_FormatoNumero(g_rst_Princi!SOLMAE_CUOAPR_SOL, 12, 2)
      
      grd_Listad(6).Rows = grd_Listad(6).Rows + 1
      grd_Listad(6).Row = grd_Listad(6).Rows - 1
      grd_Listad(6).Col = 0
      grd_Listad(6).Text = "Cuota Aceptada por Cliente (" & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & ")"
      
      grd_Listad(6).Col = 1
      grd_Listad(6).CellFontName = "Lucida Console"
      grd_Listad(6).CellFontSize = 8
      grd_Listad(6).Text = gf_FormatoNumero(g_rst_Princi!SOLMAE_CUOAPR_MPR, 12, 2)
      
      If moddat_g_int_TipMon <> 1 Then
         grd_Listad(6).Rows = grd_Listad(6).Rows + 1
         grd_Listad(6).Row = grd_Listad(6).Rows - 1
         grd_Listad(6).Col = 0
         grd_Listad(6).Text = "Tipo de Cambio de Aceptación (" & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & ")"
         
         grd_Listad(6).Col = 1
         grd_Listad(6).CellFontName = "Lucida Console"
         grd_Listad(6).CellFontSize = 8
         grd_Listad(6).Text = gf_FormatoNumero(g_rst_Princi!SOLMAE_TCAMPR_APR, 12, 4)
      End If
      
      grd_Listad(6).Rows = grd_Listad(6).Rows + 2
      grd_Listad(6).Row = grd_Listad(6).Rows - 1
      grd_Listad(6).Col = 0
      grd_Listad(6).Text = "Cuota Aprobada (S/.)"
      
      grd_Listad(6).Col = 1
      grd_Listad(6).CellFontName = "Lucida Console"
      grd_Listad(6).CellFontSize = 8
      grd_Listad(6).Text = gf_FormatoNumero(g_rst_Princi!SOLMAE_CUOMEN_SOL, 12, 2)
      
      grd_Listad(6).Rows = grd_Listad(6).Rows + 1
      grd_Listad(6).Row = grd_Listad(6).Rows - 1
      grd_Listad(6).Col = 0
      grd_Listad(6).Text = "Cuota Aprobada (" & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & ")"
      
      grd_Listad(6).Col = 1
      grd_Listad(6).CellFontName = "Lucida Console"
      grd_Listad(6).CellFontSize = 8
      grd_Listad(6).Text = gf_FormatoNumero(g_rst_Princi!SOLMAE_CUOMEN_MPR, 12, 2)
   
      If moddat_g_int_TipMon <> 1 Then
         grd_Listad(6).Rows = grd_Listad(6).Rows + 1
         grd_Listad(6).Row = grd_Listad(6).Rows - 1
         grd_Listad(6).Col = 0
         grd_Listad(6).Text = "Tipo de Cambio de Aprobación (" & moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & ")"
         
         grd_Listad(6).Col = 1
         grd_Listad(6).CellFontName = "Lucida Console"
         grd_Listad(6).CellFontSize = 8
         grd_Listad(6).Text = gf_FormatoNumero(g_rst_Princi!SOLMAE_TIPCAM, 12, 4)
      End If
   
      grd_Listad(6).Rows = grd_Listad(6).Rows + 2
      grd_Listad(6).Row = grd_Listad(6).Rows - 1
      grd_Listad(6).Col = 0
      grd_Listad(6).Text = "Total Ingreso Líquido S/. "
      
      grd_Listad(6).Col = 1
      grd_Listad(6).CellFontName = "Lucida Console"
      grd_Listad(6).CellFontSize = 8
      grd_Listad(6).Text = gf_FormatoNumero(g_rst_Princi!SOLMAE_INGNET, 12, 2)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   Call gs_UbiIniGrid(grd_Listad(6))
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
      grd_LisEva.Text = moddat_gf_Consulta_ParDes("241", g_rst_Princi!HIPDES_TIPGAR)
      
      If g_rst_Princi!HIPDES_TIPGAR = 2 Or g_rst_Princi!HIPDES_TIPGAR = 4 Or g_rst_Princi!HIPDES_TIPGAR = 5 Or g_rst_Princi!HIPDES_TIPGAR = 3 Then
         grd_LisEva.Rows = grd_LisEva.Rows + 1
         grd_LisEva.Row = grd_LisEva.Rows - 1
         grd_LisEva.Col = 0
         grd_LisEva.Text = "Forma de Desembolso"
         
         grd_LisEva.Col = 1
         grd_LisEva.Text = moddat_gf_Consulta_ParDes("226", g_rst_Princi!HIPDES_TIPDES)
      End If
      
      If g_rst_Princi!HIPDES_TIPDES = 1 Then
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
            grd_LisEva.Rows = grd_LisEva.Rows + 1
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
            grd_LisEva.Rows = grd_LisEva.Rows + 1
            grd_LisEva.Row = grd_LisEva.Rows - 1
            grd_LisEva.Col = 0
            grd_LisEva.Text = "Nro. Carta Fianza"
            
            grd_LisEva.Col = 1
            grd_LisEva.Text = "CARTA FIANZA NO RECIBIDA"
         End If
      End If
      Call gs_UbiIniGrid(grd_LisEva)
      
      txt_ObsDes.Text = Trim(g_rst_Princi!HIPDES_OBSERV & "")
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
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
            moddat_g_rst_RecDAO("CROCAB_BANFIA") = moddat_gf_Consulta_ParDes("505", g_rst_Genera!HIPDES_BANFIA)
            moddat_g_rst_RecDAO("CROCAB_NUMFIA") = Trim(g_rst_Genera!HIPDES_NUMFIA & "")
            moddat_g_rst_RecDAO("CROCAB_EMIFIA") = gf_FormatoFecha(CStr(g_rst_Genera!HIPDES_EMIFIA))
            moddat_g_rst_RecDAO("CROCAB_VCTFIA") = gf_FormatoFecha(CStr(g_rst_Genera!HIPDES_VCTFIA))
            moddat_g_rst_RecDAO("CROCAB_MONFIA") = moddat_gf_Consulta_ParDes("204", g_rst_Genera!HIPDES_MONFIA)
            moddat_g_rst_RecDAO("CROCAB_MTOFIA") = g_rst_Genera!HIPDES_IMPFIA
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

Private Sub fs_DatMVi()
   If moddat_g_str_CodPrd <> "001" Then
      Exit Sub
   End If

   txt_ObsMVi.Text = ""
   Call gs_LimpiaGrid(grd_Listad(12))
   
   g_str_Parame = "SELECT * FROM TRA_EVAMVI WHERE "
   g_str_Parame = g_str_Parame & "EVAMVI_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
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
   If Not (moddat_g_str_CodPrd = "004" Or moddat_g_str_CodPrd = "003") Then
      Exit Sub
   End If
   
   txt_ObsMVi.Text = ""
   Call gs_LimpiaGrid(grd_Listad(12))
   
   g_str_Parame = "SELECT * FROM TRA_EVACOF WHERE "
   g_str_Parame = g_str_Parame & "EVACOF_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      grd_Listad(12).Rows = grd_Listad(12).Rows + 1
      grd_Listad(12).Row = grd_Listad(12).Rows - 1
      grd_Listad(12).Col = 0
      grd_Listad(12).Text = "Fecha Envío"
      
      grd_Listad(12).Col = 1
      grd_Listad(12).Text = gf_FormatoFecha(CStr(g_rst_Princi!EVACOF_FECENV))
   
      If g_rst_Princi!EVACOF_FECREC > 0 Then
         If moddat_g_str_CodPrd = "003" Then
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

