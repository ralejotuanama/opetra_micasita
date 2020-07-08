VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frm_Ges_TecPro_05 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   10065
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14625
   Icon            =   "OpeTra_frm_824.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10065
   ScaleWidth      =   14625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   10095
      Left            =   0
      TabIndex        =   71
      Top             =   0
      Width           =   14655
      _Version        =   65536
      _ExtentX        =   25850
      _ExtentY        =   17806
      _StockProps     =   15
      BackColor       =   14215660
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
         Height          =   10035
         Left            =   0
         TabIndex        =   72
         Top             =   30
         Width           =   14625
         _Version        =   65536
         _ExtentX        =   25797
         _ExtentY        =   17701
         _StockProps     =   15
         BackColor       =   14215660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin Threed.SSPanel SSPanel8 
            Height          =   675
            Left            =   60
            TabIndex        =   73
            Top             =   750
            Width           =   14505
            _Version        =   65536
            _ExtentX        =   25585
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
            Begin VB.CommandButton cmd_Salida 
               Height          =   585
               Left            =   13890
               Picture         =   "OpeTra_frm_824.frx":000C
               Style           =   1  'Graphical
               TabIndex        =   69
               ToolTipText     =   "Salir"
               Top             =   30
               Width           =   585
            End
            Begin VB.CommandButton cmd_Cancel 
               Height          =   585
               Left            =   13320
               Picture         =   "OpeTra_frm_824.frx":044E
               Style           =   1  'Graphical
               TabIndex        =   68
               ToolTipText     =   "Cancelar"
               Top             =   30
               Width           =   585
            End
            Begin VB.CommandButton cmd_ExpExc 
               Height          =   585
               Left            =   1800
               Picture         =   "OpeTra_frm_824.frx":0758
               Style           =   1  'Graphical
               TabIndex        =   67
               ToolTipText     =   "Exportar a Excel"
               Top             =   30
               Width           =   585
            End
            Begin VB.CommandButton cmd_Grabar 
               Height          =   585
               Left            =   12750
               Picture         =   "OpeTra_frm_824.frx":0A62
               Style           =   1  'Graphical
               TabIndex        =   63
               ToolTipText     =   "Grabar Datos"
               Top             =   30
               Width           =   585
            End
            Begin VB.CommandButton cmd_Borrar 
               Height          =   585
               Left            =   1230
               Picture         =   "OpeTra_frm_824.frx":0EA4
               Style           =   1  'Graphical
               TabIndex        =   66
               ToolTipText     =   "Borrar Registro"
               Top             =   30
               Width           =   585
            End
            Begin VB.CommandButton cmd_Editar 
               Height          =   585
               Left            =   660
               Picture         =   "OpeTra_frm_824.frx":11AE
               Style           =   1  'Graphical
               TabIndex        =   65
               ToolTipText     =   "Modificar Registro"
               Top             =   30
               Width           =   585
            End
            Begin VB.CommandButton cmd_Agrega 
               Height          =   585
               Left            =   90
               Picture         =   "OpeTra_frm_824.frx":14B8
               Style           =   1  'Graphical
               TabIndex        =   64
               ToolTipText     =   "Nuevo Registro"
               Top             =   30
               Width           =   585
            End
         End
         Begin Threed.SSPanel SSPanel6 
            Height          =   675
            Left            =   60
            TabIndex        =   74
            Top             =   30
            Width           =   14505
            _Version        =   65536
            _ExtentX        =   25585
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
               Left            =   630
               TabIndex        =   75
               Top             =   30
               Width           =   4215
               _Version        =   65536
               _ExtentX        =   7435
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "Gestión de Crédito Hipotecario"
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
            Begin Threed.SSPanel SSPanel15 
               Height          =   315
               Left            =   630
               TabIndex        =   76
               Top             =   330
               Width           =   4215
               _Version        =   65536
               _ExtentX        =   7435
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "Techo Propio - Registro de Garantías"
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
               Height          =   480
               Left            =   60
               Picture         =   "OpeTra_frm_824.frx":17C2
               Top             =   60
               Width           =   480
            End
         End
         Begin Threed.SSPanel SSPanel9 
            Height          =   1155
            Left            =   60
            TabIndex        =   77
            Top             =   1470
            Width           =   14505
            _Version        =   65536
            _ExtentX        =   25585
            _ExtentY        =   2037
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
            Begin Threed.SSPanel pnl_RazSoc 
               Height          =   315
               Left            =   1620
               TabIndex        =   78
               Top             =   450
               Width           =   5625
               _Version        =   65536
               _ExtentX        =   9922
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
               Alignment       =   1
            End
            Begin Threed.SSPanel pnl_TipDoc 
               Height          =   315
               Left            =   1620
               TabIndex        =   79
               Top             =   120
               Width           =   5625
               _Version        =   65536
               _ExtentX        =   9922
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
               Alignment       =   1
            End
            Begin Threed.SSPanel pnl_NroDoc 
               Height          =   315
               Left            =   9360
               TabIndex        =   80
               Top             =   120
               Width           =   2895
               _Version        =   65536
               _ExtentX        =   5106
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
               Alignment       =   1
            End
            Begin Threed.SSPanel pnl_TipEmp 
               Height          =   315
               Left            =   1620
               TabIndex        =   81
               Top             =   780
               Width           =   2895
               _Version        =   65536
               _ExtentX        =   5106
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
               Alignment       =   1
            End
            Begin Threed.SSPanel pnl_NumRef 
               Height          =   315
               Left            =   9360
               TabIndex        =   86
               Top             =   450
               Visible         =   0   'False
               Width           =   2895
               _Version        =   65536
               _ExtentX        =   5106
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
               Alignment       =   1
            End
            Begin VB.Label Label3 
               Caption         =   "Nro. Referencia:"
               Height          =   255
               Left            =   7770
               TabIndex        =   87
               Top             =   480
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label lbl_TipDoc 
               Caption         =   "Tipo Documento:"
               Height          =   255
               Left            =   120
               TabIndex        =   85
               Top             =   150
               Width           =   1335
            End
            Begin VB.Label lbl_NumDoc 
               Caption         =   "Nro. Documento:"
               Height          =   225
               Left            =   7770
               TabIndex        =   84
               Top             =   150
               Width           =   1335
            End
            Begin VB.Label lbl_RazSoc 
               Caption         =   "Razón Social:"
               Height          =   255
               Left            =   120
               TabIndex        =   83
               Top             =   480
               Width           =   1335
            End
            Begin VB.Label lbl_TipEmp 
               Caption         =   "Tipo Empresa:"
               Height          =   255
               Left            =   120
               TabIndex        =   82
               Top             =   810
               Width           =   1335
            End
         End
         Begin Threed.SSPanel SSPanel4 
            Height          =   2475
            Left            =   60
            TabIndex        =   88
            Top             =   2670
            Width           =   14505
            _Version        =   65536
            _ExtentX        =   25585
            _ExtentY        =   4366
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
               Height          =   1935
               Left            =   90
               TabIndex        =   70
               Top             =   450
               Width           =   14310
               _ExtentX        =   25241
               _ExtentY        =   3413
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
            Begin Threed.SSPanel pnl_Tit_TipGar 
               Height          =   285
               Left            =   90
               TabIndex        =   89
               Top             =   150
               Width           =   2565
               _Version        =   65536
               _ExtentX        =   4524
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Tipo Garantía"
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
            Begin Threed.SSPanel pnl_Tit_NumRef 
               Height          =   285
               Left            =   8070
               TabIndex        =   90
               Top             =   150
               Width           =   4170
               _Version        =   65536
               _ExtentX        =   7355
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Cartas Fianza"
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
            Begin Threed.SSPanel pnl_Tit_FecEmi 
               Height          =   285
               Left            =   2640
               TabIndex        =   91
               Top             =   150
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Fecha Emisión"
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
            Begin Threed.SSPanel pnl_Tit_TipMon 
               Height          =   285
               Left            =   4020
               TabIndex        =   92
               Top             =   150
               Width           =   2250
               _Version        =   65536
               _ExtentX        =   3969
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Moneda"
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
            Begin Threed.SSPanel pnl_Tit_MtoGar 
               Height          =   285
               Left            =   6240
               TabIndex        =   93
               Top             =   150
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Mto. Garantía"
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
               Left            =   12210
               TabIndex        =   94
               Top             =   150
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Monto CF"
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
         Begin Threed.SSPanel SSPanel10 
            Height          =   3165
            Left            =   60
            TabIndex        =   95
            Top             =   6270
            Width           =   14505
            _Version        =   65536
            _ExtentX        =   25585
            _ExtentY        =   5583
            _StockProps     =   15
            BackColor       =   14215660
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
            Begin TabDlg.SSTab tab_Genera 
               Height          =   3075
               Left            =   90
               TabIndex        =   96
               Top             =   30
               Width           =   14325
               _ExtentX        =   25268
               _ExtentY        =   5424
               _Version        =   393216
               Style           =   1
               Tabs            =   6
               TabsPerRow      =   7
               TabHeight       =   520
               TabCaption(0)   =   "Garantía Líquida"
               TabPicture(0)   =   "OpeTra_frm_824.frx":1ACC
               Tab(0).ControlEnabled=   -1  'True
               Tab(0).Control(0)=   "Label2"
               Tab(0).Control(0).Enabled=   0   'False
               Tab(0).Control(1)=   "Label5"
               Tab(0).Control(1).Enabled=   0   'False
               Tab(0).Control(2)=   "Label4"
               Tab(0).Control(2).Enabled=   0   'False
               Tab(0).Control(3)=   "Label16"
               Tab(0).Control(3).Enabled=   0   'False
               Tab(0).Control(4)=   "pnl_NumCFi"
               Tab(0).Control(4).Enabled=   0   'False
               Tab(0).Control(5)=   "ipp_FecEmi"
               Tab(0).Control(5).Enabled=   0   'False
               Tab(0).Control(6)=   "ipp_ImpGar"
               Tab(0).Control(6).Enabled=   0   'False
               Tab(0).Control(7)=   "cmd_CfiAso"
               Tab(0).Control(7).Enabled=   0   'False
               Tab(0).Control(8)=   "cmb_NumRef"
               Tab(0).Control(8).Enabled=   0   'False
               Tab(0).Control(9)=   "cmb_Moneda"
               Tab(0).Control(9).Enabled=   0   'False
               Tab(0).ControlCount=   10
               TabCaption(1)   =   "Inmueble"
               TabPicture(1)   =   "OpeTra_frm_824.frx":1AE8
               Tab(1).ControlEnabled=   0   'False
               Tab(1).Control(0)=   "txt_NumFic_Inm"
               Tab(1).Control(1)=   "txt_NumAFi_Inm"
               Tab(1).Control(2)=   "txt_NumPre_Inm"
               Tab(1).Control(3)=   "txt_NumAPa_Inm"
               Tab(1).Control(4)=   "txt_NumPar_Inm"
               Tab(1).Control(5)=   "cmb_TipDoc_Inm"
               Tab(1).Control(6)=   "cmb_Moneda_Inm"
               Tab(1).Control(7)=   "ipp_FecPre_Inm"
               Tab(1).Control(8)=   "ipp_FecIns_Inm"
               Tab(1).Control(9)=   "ipp_MtoHip_Inm"
               Tab(1).Control(10)=   "Label7"
               Tab(1).Control(11)=   "Label6"
               Tab(1).Control(12)=   "Label55"
               Tab(1).Control(13)=   "Label47"
               Tab(1).Control(14)=   "Label46"
               Tab(1).Control(15)=   "Label15"
               Tab(1).Control(16)=   "Label8"
               Tab(1).Control(17)=   "Label9"
               Tab(1).Control(18)=   "Label10"
               Tab(1).Control(19)=   "Label22"
               Tab(1).ControlCount=   20
               TabCaption(2)   =   "Estacionamiento 1"
               TabPicture(2)   =   "OpeTra_frm_824.frx":1B04
               Tab(2).ControlEnabled=   0   'False
               Tab(2).Control(0)=   "Label24"
               Tab(2).Control(1)=   "Label14"
               Tab(2).Control(2)=   "Label20"
               Tab(2).Control(3)=   "Label19"
               Tab(2).Control(4)=   "Label18"
               Tab(2).Control(5)=   "Label17"
               Tab(2).Control(6)=   "Label23"
               Tab(2).Control(7)=   "Label48"
               Tab(2).Control(8)=   "Label49"
               Tab(2).Control(9)=   "Label56"
               Tab(2).Control(10)=   "Label11"
               Tab(2).Control(11)=   "ipp_MtoHip_Es1"
               Tab(2).Control(12)=   "ipp_FecIns_Es1"
               Tab(2).Control(13)=   "ipp_FecPre_Es1"
               Tab(2).Control(14)=   "cmb_Moneda_Es1"
               Tab(2).Control(15)=   "cmb_FlgEst_Es1"
               Tab(2).Control(16)=   "cmb_TipDoc_Es1"
               Tab(2).Control(17)=   "txt_NumPar_Es1"
               Tab(2).Control(18)=   "txt_NumAPa_Es1"
               Tab(2).Control(19)=   "txt_NumFic_Es1"
               Tab(2).Control(20)=   "txt_NumAFi_Es1"
               Tab(2).Control(21)=   "txt_NumPre_Es1"
               Tab(2).ControlCount=   22
               TabCaption(3)   =   "Estacionamiento 2"
               TabPicture(3)   =   "OpeTra_frm_824.frx":1B20
               Tab(3).ControlEnabled=   0   'False
               Tab(3).Control(0)=   "Label37"
               Tab(3).Control(1)=   "Label38"
               Tab(3).Control(2)=   "Label39"
               Tab(3).Control(3)=   "Label63"
               Tab(3).Control(4)=   "Label64"
               Tab(3).Control(5)=   "Label65"
               Tab(3).Control(6)=   "Label66"
               Tab(3).Control(7)=   "Label67"
               Tab(3).Control(8)=   "Label68"
               Tab(3).Control(9)=   "Label69"
               Tab(3).Control(10)=   "Label70"
               Tab(3).Control(11)=   "ipp_MtoHip_Es2"
               Tab(3).Control(12)=   "ipp_FecIns_Es2"
               Tab(3).Control(13)=   "ipp_FecPre_Es2"
               Tab(3).Control(14)=   "txt_NumPre_Es2"
               Tab(3).Control(15)=   "txt_NumAFi_Es2"
               Tab(3).Control(16)=   "txt_NumFic_Es2"
               Tab(3).Control(17)=   "txt_NumAPa_Es2"
               Tab(3).Control(18)=   "txt_NumPar_Es2"
               Tab(3).Control(19)=   "cmb_TipDoc_Es2"
               Tab(3).Control(20)=   "cmb_FlgEst_Es2"
               Tab(3).Control(21)=   "cmb_Moneda_Es2"
               Tab(3).ControlCount=   22
               TabCaption(4)   =   "Depósito 1"
               TabPicture(4)   =   "OpeTra_frm_824.frx":1B3C
               Tab(4).ControlEnabled=   0   'False
               Tab(4).Control(0)=   "Label58"
               Tab(4).Control(1)=   "Label53"
               Tab(4).Control(2)=   "Label52"
               Tab(4).Control(3)=   "Label51"
               Tab(4).Control(4)=   "Label44"
               Tab(4).Control(5)=   "Label43"
               Tab(4).Control(6)=   "Label42"
               Tab(4).Control(7)=   "Label41"
               Tab(4).Control(8)=   "Label40"
               Tab(4).Control(9)=   "Label50"
               Tab(4).Control(10)=   "Label25"
               Tab(4).Control(11)=   "ipp_MtoHip_Dep1"
               Tab(4).Control(12)=   "ipp_FecIns_Dep1"
               Tab(4).Control(13)=   "ipp_FecPre_Dep1"
               Tab(4).Control(14)=   "txt_NumPre_Dep1"
               Tab(4).Control(15)=   "txt_NumAFi_Dep1"
               Tab(4).Control(16)=   "txt_NumFic_Dep1"
               Tab(4).Control(17)=   "txt_NumAPa_Dep1"
               Tab(4).Control(18)=   "txt_NumPar_Dep1"
               Tab(4).Control(19)=   "cmb_TipDoc_Dep1"
               Tab(4).Control(20)=   "cmb_FlgEst_Dep1"
               Tab(4).Control(21)=   "cmb_Moneda_Dep1"
               Tab(4).ControlCount=   22
               TabCaption(5)   =   "Depósito 2"
               TabPicture(5)   =   "OpeTra_frm_824.frx":1B58
               Tab(5).ControlEnabled=   0   'False
               Tab(5).Control(0)=   "Label27"
               Tab(5).Control(1)=   "Label28"
               Tab(5).Control(2)=   "Label71"
               Tab(5).Control(3)=   "Label72"
               Tab(5).Control(4)=   "Label73"
               Tab(5).Control(5)=   "Label74"
               Tab(5).Control(6)=   "Label75"
               Tab(5).Control(7)=   "Label76"
               Tab(5).Control(8)=   "Label77"
               Tab(5).Control(9)=   "Label78"
               Tab(5).Control(10)=   "Label79"
               Tab(5).Control(11)=   "ipp_MtoHip_Dep2"
               Tab(5).Control(12)=   "ipp_FecIns_Dep2"
               Tab(5).Control(13)=   "ipp_FecPre_Dep2"
               Tab(5).Control(14)=   "txt_NumPre_Dep2"
               Tab(5).Control(15)=   "txt_NumAFi_Dep2"
               Tab(5).Control(16)=   "txt_NumFic_Dep2"
               Tab(5).Control(17)=   "txt_NumAPa_Dep2"
               Tab(5).Control(18)=   "txt_NumPar_Dep2"
               Tab(5).Control(19)=   "cmb_TipDoc_Dep2"
               Tab(5).Control(20)=   "cmb_FlgEst_Dep2"
               Tab(5).Control(21)=   "cmb_Moneda_Dep2"
               Tab(5).ControlCount=   22
               Begin VB.ComboBox cmb_Moneda_Dep2 
                  Height          =   315
                  Left            =   -69510
                  Style           =   2  'Dropdown List
                  TabIndex        =   62
                  Top             =   2700
                  Width           =   2715
               End
               Begin VB.ComboBox cmb_FlgEst_Dep2 
                  Height          =   315
                  Left            =   -73140
                  Style           =   2  'Dropdown List
                  TabIndex        =   52
                  Top             =   390
                  Width           =   975
               End
               Begin VB.ComboBox cmb_TipDoc_Dep2 
                  Height          =   315
                  Left            =   -73140
                  Style           =   2  'Dropdown List
                  TabIndex        =   56
                  Top             =   1710
                  Width           =   3825
               End
               Begin VB.TextBox txt_NumPar_Dep2 
                  Height          =   315
                  Left            =   -73140
                  MaxLength       =   12
                  TabIndex        =   57
                  Top             =   2040
                  Width           =   1425
               End
               Begin VB.TextBox txt_NumAPa_Dep2 
                  Height          =   315
                  Left            =   -69510
                  MaxLength       =   12
                  TabIndex        =   58
                  Top             =   2040
                  Width           =   1425
               End
               Begin VB.TextBox txt_NumFic_Dep2 
                  Height          =   315
                  Left            =   -73140
                  MaxLength       =   12
                  TabIndex        =   59
                  Top             =   2370
                  Width           =   1425
               End
               Begin VB.TextBox txt_NumAFi_Dep2 
                  Height          =   315
                  Left            =   -69510
                  MaxLength       =   12
                  TabIndex        =   60
                  Top             =   2370
                  Width           =   1425
               End
               Begin VB.TextBox txt_NumPre_Dep2 
                  Height          =   315
                  Left            =   -73140
                  MaxLength       =   12
                  TabIndex        =   54
                  Top             =   1050
                  Width           =   1425
               End
               Begin VB.ComboBox cmb_Moneda_Es2 
                  Height          =   315
                  Left            =   -69510
                  Style           =   2  'Dropdown List
                  TabIndex        =   40
                  Top             =   2700
                  Width           =   2715
               End
               Begin VB.ComboBox cmb_FlgEst_Es2 
                  Height          =   315
                  Left            =   -73140
                  Style           =   2  'Dropdown List
                  TabIndex        =   30
                  Top             =   390
                  Width           =   975
               End
               Begin VB.ComboBox cmb_TipDoc_Es2 
                  Height          =   315
                  Left            =   -73140
                  Style           =   2  'Dropdown List
                  TabIndex        =   34
                  Top             =   1710
                  Width           =   3825
               End
               Begin VB.TextBox txt_NumPar_Es2 
                  Height          =   315
                  Left            =   -73140
                  MaxLength       =   12
                  TabIndex        =   35
                  Top             =   2040
                  Width           =   1425
               End
               Begin VB.TextBox txt_NumAPa_Es2 
                  Height          =   315
                  Left            =   -69510
                  MaxLength       =   12
                  TabIndex        =   36
                  Top             =   2040
                  Width           =   1425
               End
               Begin VB.TextBox txt_NumFic_Es2 
                  Height          =   315
                  Left            =   -73140
                  MaxLength       =   12
                  TabIndex        =   37
                  Top             =   2370
                  Width           =   1425
               End
               Begin VB.TextBox txt_NumAFi_Es2 
                  Height          =   315
                  Left            =   -69510
                  MaxLength       =   12
                  TabIndex        =   38
                  Top             =   2370
                  Width           =   1425
               End
               Begin VB.TextBox txt_NumPre_Es2 
                  Height          =   315
                  Left            =   -73140
                  MaxLength       =   12
                  TabIndex        =   32
                  Top             =   1050
                  Width           =   1425
               End
               Begin VB.ComboBox cmb_Moneda_Dep1 
                  Height          =   315
                  Left            =   -69510
                  Style           =   2  'Dropdown List
                  TabIndex        =   51
                  Top             =   2700
                  Width           =   2715
               End
               Begin VB.ComboBox cmb_FlgEst_Dep1 
                  Height          =   315
                  Left            =   -73140
                  Style           =   2  'Dropdown List
                  TabIndex        =   41
                  Top             =   390
                  Width           =   975
               End
               Begin VB.ComboBox cmb_TipDoc_Dep1 
                  Height          =   315
                  Left            =   -73140
                  Style           =   2  'Dropdown List
                  TabIndex        =   45
                  Top             =   1710
                  Width           =   3825
               End
               Begin VB.TextBox txt_NumPar_Dep1 
                  Height          =   315
                  Left            =   -73140
                  MaxLength       =   12
                  TabIndex        =   46
                  Top             =   2040
                  Width           =   1425
               End
               Begin VB.TextBox txt_NumAPa_Dep1 
                  Height          =   315
                  Left            =   -69510
                  MaxLength       =   12
                  TabIndex        =   47
                  Top             =   2040
                  Width           =   1425
               End
               Begin VB.TextBox txt_NumFic_Dep1 
                  Height          =   315
                  Left            =   -73140
                  MaxLength       =   12
                  TabIndex        =   48
                  Top             =   2370
                  Width           =   1425
               End
               Begin VB.TextBox txt_NumAFi_Dep1 
                  Height          =   315
                  Left            =   -69510
                  MaxLength       =   12
                  TabIndex        =   49
                  Top             =   2370
                  Width           =   1425
               End
               Begin VB.TextBox txt_NumPre_Dep1 
                  Height          =   315
                  Left            =   -73140
                  MaxLength       =   12
                  TabIndex        =   43
                  Top             =   1050
                  Width           =   1425
               End
               Begin VB.TextBox txt_NumPre_Es1 
                  Height          =   315
                  Left            =   -73140
                  MaxLength       =   12
                  TabIndex        =   21
                  Top             =   1050
                  Width           =   1425
               End
               Begin VB.TextBox txt_NumAFi_Es1 
                  Height          =   315
                  Left            =   -69510
                  MaxLength       =   12
                  TabIndex        =   27
                  Top             =   2370
                  Width           =   1425
               End
               Begin VB.TextBox txt_NumFic_Es1 
                  Height          =   315
                  Left            =   -73140
                  MaxLength       =   12
                  TabIndex        =   26
                  Top             =   2370
                  Width           =   1425
               End
               Begin VB.TextBox txt_NumAPa_Es1 
                  Height          =   315
                  Left            =   -69510
                  MaxLength       =   12
                  TabIndex        =   25
                  Top             =   2040
                  Width           =   1425
               End
               Begin VB.TextBox txt_NumPar_Es1 
                  Height          =   315
                  Left            =   -73140
                  MaxLength       =   12
                  TabIndex        =   24
                  Top             =   2040
                  Width           =   1425
               End
               Begin VB.ComboBox cmb_TipDoc_Es1 
                  Height          =   315
                  Left            =   -73140
                  Style           =   2  'Dropdown List
                  TabIndex        =   23
                  Top             =   1710
                  Width           =   3825
               End
               Begin VB.TextBox txt_NumFic_Inm 
                  Height          =   315
                  Left            =   -73140
                  MaxLength       =   12
                  TabIndex        =   15
                  Top             =   2250
                  Width           =   1425
               End
               Begin VB.TextBox txt_NumAFi_Inm 
                  Height          =   315
                  Left            =   -69510
                  MaxLength       =   12
                  TabIndex        =   16
                  Top             =   2250
                  Width           =   1425
               End
               Begin VB.TextBox txt_NumPre_Inm 
                  Height          =   315
                  Left            =   -73140
                  MaxLength       =   12
                  TabIndex        =   10
                  Top             =   810
                  Width           =   1425
               End
               Begin VB.TextBox txt_NumAPa_Inm 
                  Height          =   315
                  Left            =   -69510
                  MaxLength       =   12
                  TabIndex        =   14
                  Top             =   1890
                  Width           =   1425
               End
               Begin VB.TextBox txt_NumPar_Inm 
                  Height          =   315
                  Left            =   -73140
                  MaxLength       =   12
                  TabIndex        =   13
                  Top             =   1890
                  Width           =   1425
               End
               Begin VB.ComboBox cmb_TipDoc_Inm 
                  Height          =   315
                  Left            =   -73140
                  Style           =   2  'Dropdown List
                  TabIndex        =   12
                  Top             =   1530
                  Width           =   3825
               End
               Begin VB.ComboBox cmb_FlgEst_Es1 
                  Height          =   315
                  Left            =   -73140
                  Style           =   2  'Dropdown List
                  TabIndex        =   19
                  Top             =   390
                  Width           =   975
               End
               Begin VB.ComboBox cmb_Moneda_Inm 
                  Height          =   315
                  Left            =   -69510
                  Style           =   2  'Dropdown List
                  TabIndex        =   18
                  Top             =   2610
                  Width           =   2715
               End
               Begin VB.ComboBox cmb_Moneda_Es1 
                  Height          =   315
                  Left            =   -69510
                  Style           =   2  'Dropdown List
                  TabIndex        =   29
                  Top             =   2700
                  Width           =   2715
               End
               Begin VB.ComboBox cmb_Moneda 
                  Height          =   315
                  Left            =   1920
                  Style           =   2  'Dropdown List
                  TabIndex        =   6
                  Top             =   1110
                  Width           =   2715
               End
               Begin VB.ComboBox cmb_NumRef 
                  Height          =   315
                  Left            =   1920
                  Style           =   2  'Dropdown List
                  TabIndex        =   8
                  Top             =   1830
                  Width           =   2715
               End
               Begin VB.CommandButton cmd_CfiAso 
                  Caption         =   "..."
                  Height          =   285
                  Left            =   4740
                  TabIndex        =   97
                  Top             =   1860
                  Width           =   375
               End
               Begin EditLib.fpDateTime ipp_FecPre_Es1 
                  Height          =   315
                  Left            =   -73140
                  TabIndex        =   20
                  Top             =   720
                  Width           =   1425
                  _Version        =   196608
                  _ExtentX        =   2514
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
                  ThreeDInsideHighlightColor=   -2147483633
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
                  ButtonStyle     =   3
                  ButtonWidth     =   0
                  ButtonWrap      =   -1  'True
                  ButtonDefaultAction=   -1  'True
                  ThreeDText      =   0
                  ThreeDTextHighlightColor=   -2147483633
                  ThreeDTextShadowColor=   -2147483632
                  ThreeDTextOffset=   1
                  AlignTextH      =   0
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
                  Text            =   "28/09/2004"
                  DateCalcMethod  =   0
                  DateTimeFormat  =   0
                  UserDefinedFormat=   ""
                  DateMax         =   "00000000"
                  DateMin         =   "00000000"
                  TimeMax         =   "000000"
                  TimeMin         =   "000000"
                  TimeString1159  =   ""
                  TimeString2359  =   ""
                  DateDefault     =   "00000000"
                  TimeDefault     =   "000000"
                  TimeStyle       =   0
                  BorderGrayAreaColor=   -2147483637
                  ThreeDOnFocusInvert=   0   'False
                  ThreeDFrameColor=   -2147483633
                  Appearance      =   2
                  BorderDropShadow=   0
                  BorderDropShadowColor=   -2147483632
                  BorderDropShadowWidth=   3
                  PopUpType       =   0
                  DateCalcY2KSplit=   60
                  CaretPosition   =   0
                  IncYear         =   1
                  IncMonth        =   1
                  IncDay          =   1
                  IncHour         =   1
                  IncMinute       =   1
                  IncSecond       =   1
                  ButtonColor     =   -2147483633
                  AutoMenu        =   0   'False
                  StartMonth      =   4
                  ButtonAlign     =   0
                  BoundDataType   =   0
                  OLEDropMode     =   0
                  OLEDragMode     =   0
               End
               Begin EditLib.fpDateTime ipp_FecIns_Es1 
                  Height          =   315
                  Left            =   -73140
                  TabIndex        =   22
                  Top             =   1380
                  Width           =   1425
                  _Version        =   196608
                  _ExtentX        =   2514
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
                  ThreeDInsideHighlightColor=   -2147483633
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
                  ButtonStyle     =   3
                  ButtonWidth     =   0
                  ButtonWrap      =   -1  'True
                  ButtonDefaultAction=   -1  'True
                  ThreeDText      =   0
                  ThreeDTextHighlightColor=   -2147483633
                  ThreeDTextShadowColor=   -2147483632
                  ThreeDTextOffset=   1
                  AlignTextH      =   0
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
                  Text            =   "28/09/2004"
                  DateCalcMethod  =   0
                  DateTimeFormat  =   0
                  UserDefinedFormat=   ""
                  DateMax         =   "00000000"
                  DateMin         =   "00000000"
                  TimeMax         =   "000000"
                  TimeMin         =   "000000"
                  TimeString1159  =   ""
                  TimeString2359  =   ""
                  DateDefault     =   "00000000"
                  TimeDefault     =   "000000"
                  TimeStyle       =   0
                  BorderGrayAreaColor=   -2147483637
                  ThreeDOnFocusInvert=   0   'False
                  ThreeDFrameColor=   -2147483633
                  Appearance      =   2
                  BorderDropShadow=   0
                  BorderDropShadowColor=   -2147483632
                  BorderDropShadowWidth=   3
                  PopUpType       =   0
                  DateCalcY2KSplit=   60
                  CaretPosition   =   0
                  IncYear         =   1
                  IncMonth        =   1
                  IncDay          =   1
                  IncHour         =   1
                  IncMinute       =   1
                  IncSecond       =   1
                  ButtonColor     =   -2147483633
                  AutoMenu        =   0   'False
                  StartMonth      =   4
                  ButtonAlign     =   0
                  BoundDataType   =   0
                  OLEDropMode     =   0
                  OLEDragMode     =   0
               End
               Begin EditLib.fpDoubleSingle ipp_MtoHip_Es1 
                  Height          =   315
                  Left            =   -73140
                  TabIndex        =   28
                  Top             =   2700
                  Width           =   1425
                  _Version        =   196608
                  _ExtentX        =   2514
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
               Begin EditLib.fpDateTime ipp_FecPre_Inm 
                  Height          =   315
                  Left            =   -73140
                  TabIndex        =   9
                  Top             =   450
                  Width           =   1425
                  _Version        =   196608
                  _ExtentX        =   2514
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
                  ThreeDInsideHighlightColor=   -2147483633
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
                  ButtonStyle     =   3
                  ButtonWidth     =   0
                  ButtonWrap      =   -1  'True
                  ButtonDefaultAction=   -1  'True
                  ThreeDText      =   0
                  ThreeDTextHighlightColor=   -2147483633
                  ThreeDTextShadowColor=   -2147483632
                  ThreeDTextOffset=   1
                  AlignTextH      =   0
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
                  Text            =   "28/09/2004"
                  DateCalcMethod  =   0
                  DateTimeFormat  =   0
                  UserDefinedFormat=   ""
                  DateMax         =   "00000000"
                  DateMin         =   "00000000"
                  TimeMax         =   "000000"
                  TimeMin         =   "000000"
                  TimeString1159  =   ""
                  TimeString2359  =   ""
                  DateDefault     =   "00000000"
                  TimeDefault     =   "000000"
                  TimeStyle       =   0
                  BorderGrayAreaColor=   -2147483637
                  ThreeDOnFocusInvert=   0   'False
                  ThreeDFrameColor=   -2147483633
                  Appearance      =   2
                  BorderDropShadow=   0
                  BorderDropShadowColor=   -2147483632
                  BorderDropShadowWidth=   3
                  PopUpType       =   0
                  DateCalcY2KSplit=   60
                  CaretPosition   =   0
                  IncYear         =   1
                  IncMonth        =   1
                  IncDay          =   1
                  IncHour         =   1
                  IncMinute       =   1
                  IncSecond       =   1
                  ButtonColor     =   -2147483633
                  AutoMenu        =   0   'False
                  StartMonth      =   4
                  ButtonAlign     =   0
                  BoundDataType   =   0
                  OLEDropMode     =   0
                  OLEDragMode     =   0
               End
               Begin EditLib.fpDateTime ipp_FecIns_Inm 
                  Height          =   315
                  Left            =   -73140
                  TabIndex        =   11
                  Top             =   1170
                  Width           =   1425
                  _Version        =   196608
                  _ExtentX        =   2514
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
                  ThreeDInsideHighlightColor=   -2147483633
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
                  ButtonStyle     =   3
                  ButtonWidth     =   0
                  ButtonWrap      =   -1  'True
                  ButtonDefaultAction=   -1  'True
                  ThreeDText      =   0
                  ThreeDTextHighlightColor=   -2147483633
                  ThreeDTextShadowColor=   -2147483632
                  ThreeDTextOffset=   1
                  AlignTextH      =   0
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
                  Text            =   "28/09/2004"
                  DateCalcMethod  =   0
                  DateTimeFormat  =   0
                  UserDefinedFormat=   ""
                  DateMax         =   "00000000"
                  DateMin         =   "00000000"
                  TimeMax         =   "000000"
                  TimeMin         =   "000000"
                  TimeString1159  =   ""
                  TimeString2359  =   ""
                  DateDefault     =   "00000000"
                  TimeDefault     =   "000000"
                  TimeStyle       =   0
                  BorderGrayAreaColor=   -2147483637
                  ThreeDOnFocusInvert=   0   'False
                  ThreeDFrameColor=   -2147483633
                  Appearance      =   2
                  BorderDropShadow=   0
                  BorderDropShadowColor=   -2147483632
                  BorderDropShadowWidth=   3
                  PopUpType       =   0
                  DateCalcY2KSplit=   60
                  CaretPosition   =   0
                  IncYear         =   1
                  IncMonth        =   1
                  IncDay          =   1
                  IncHour         =   1
                  IncMinute       =   1
                  IncSecond       =   1
                  ButtonColor     =   -2147483633
                  AutoMenu        =   0   'False
                  StartMonth      =   4
                  ButtonAlign     =   0
                  BoundDataType   =   0
                  OLEDropMode     =   0
                  OLEDragMode     =   0
               End
               Begin EditLib.fpDoubleSingle ipp_MtoHip_Inm 
                  Height          =   315
                  Left            =   -73140
                  TabIndex        =   17
                  Top             =   2610
                  Width           =   1425
                  _Version        =   196608
                  _ExtentX        =   2514
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
               Begin EditLib.fpDoubleSingle ipp_ImpGar 
                  Height          =   315
                  Left            =   1920
                  TabIndex        =   7
                  Top             =   1470
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
               Begin EditLib.fpDateTime ipp_FecEmi 
                  Height          =   315
                  Left            =   1920
                  TabIndex        =   5
                  Top             =   750
                  Width           =   1635
                  _Version        =   196608
                  _ExtentX        =   2884
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
                  ButtonStyle     =   3
                  ButtonWidth     =   0
                  ButtonWrap      =   -1  'True
                  ButtonDefaultAction=   -1  'True
                  ThreeDText      =   0
                  ThreeDTextHighlightColor=   -2147483637
                  ThreeDTextShadowColor=   -2147483632
                  ThreeDTextOffset=   1
                  AlignTextH      =   0
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
                  Text            =   "28/09/2004"
                  DateCalcMethod  =   0
                  DateTimeFormat  =   0
                  UserDefinedFormat=   ""
                  DateMax         =   "00000000"
                  DateMin         =   "00000000"
                  TimeMax         =   "000000"
                  TimeMin         =   "000000"
                  TimeString1159  =   ""
                  TimeString2359  =   ""
                  DateDefault     =   "00000000"
                  TimeDefault     =   "000000"
                  TimeStyle       =   0
                  BorderGrayAreaColor=   -2147483637
                  ThreeDOnFocusInvert=   0   'False
                  ThreeDFrameColor=   -2147483637
                  Appearance      =   2
                  BorderDropShadow=   0
                  BorderDropShadowColor=   -2147483632
                  BorderDropShadowWidth=   3
                  PopUpType       =   0
                  DateCalcY2KSplit=   60
                  CaretPosition   =   0
                  IncYear         =   1
                  IncMonth        =   1
                  IncDay          =   1
                  IncHour         =   1
                  IncMinute       =   1
                  IncSecond       =   1
                  ButtonColor     =   -2147483637
                  AutoMenu        =   0   'False
                  StartMonth      =   4
                  ButtonAlign     =   0
                  BoundDataType   =   0
                  OLEDropMode     =   0
                  OLEDragMode     =   0
               End
               Begin Threed.SSPanel pnl_NumCFi 
                  Height          =   315
                  Left            =   5250
                  TabIndex        =   98
                  Top             =   1860
                  Width           =   6645
                  _Version        =   65536
                  _ExtentX        =   11721
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
                  Alignment       =   1
               End
               Begin EditLib.fpDateTime ipp_FecPre_Dep1 
                  Height          =   315
                  Left            =   -73140
                  TabIndex        =   42
                  Top             =   720
                  Width           =   1425
                  _Version        =   196608
                  _ExtentX        =   2514
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
                  ThreeDInsideHighlightColor=   -2147483633
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
                  ButtonStyle     =   3
                  ButtonWidth     =   0
                  ButtonWrap      =   -1  'True
                  ButtonDefaultAction=   -1  'True
                  ThreeDText      =   0
                  ThreeDTextHighlightColor=   -2147483633
                  ThreeDTextShadowColor=   -2147483632
                  ThreeDTextOffset=   1
                  AlignTextH      =   0
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
                  Text            =   "28/09/2004"
                  DateCalcMethod  =   0
                  DateTimeFormat  =   0
                  UserDefinedFormat=   ""
                  DateMax         =   "00000000"
                  DateMin         =   "00000000"
                  TimeMax         =   "000000"
                  TimeMin         =   "000000"
                  TimeString1159  =   ""
                  TimeString2359  =   ""
                  DateDefault     =   "00000000"
                  TimeDefault     =   "000000"
                  TimeStyle       =   0
                  BorderGrayAreaColor=   -2147483637
                  ThreeDOnFocusInvert=   0   'False
                  ThreeDFrameColor=   -2147483633
                  Appearance      =   2
                  BorderDropShadow=   0
                  BorderDropShadowColor=   -2147483632
                  BorderDropShadowWidth=   3
                  PopUpType       =   0
                  DateCalcY2KSplit=   60
                  CaretPosition   =   0
                  IncYear         =   1
                  IncMonth        =   1
                  IncDay          =   1
                  IncHour         =   1
                  IncMinute       =   1
                  IncSecond       =   1
                  ButtonColor     =   -2147483633
                  AutoMenu        =   0   'False
                  StartMonth      =   4
                  ButtonAlign     =   0
                  BoundDataType   =   0
                  OLEDropMode     =   0
                  OLEDragMode     =   0
               End
               Begin EditLib.fpDateTime ipp_FecIns_Dep1 
                  Height          =   315
                  Left            =   -73140
                  TabIndex        =   44
                  Top             =   1380
                  Width           =   1425
                  _Version        =   196608
                  _ExtentX        =   2514
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
                  ThreeDInsideHighlightColor=   -2147483633
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
                  ButtonStyle     =   3
                  ButtonWidth     =   0
                  ButtonWrap      =   -1  'True
                  ButtonDefaultAction=   -1  'True
                  ThreeDText      =   0
                  ThreeDTextHighlightColor=   -2147483633
                  ThreeDTextShadowColor=   -2147483632
                  ThreeDTextOffset=   1
                  AlignTextH      =   0
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
                  Text            =   "28/09/2004"
                  DateCalcMethod  =   0
                  DateTimeFormat  =   0
                  UserDefinedFormat=   ""
                  DateMax         =   "00000000"
                  DateMin         =   "00000000"
                  TimeMax         =   "000000"
                  TimeMin         =   "000000"
                  TimeString1159  =   ""
                  TimeString2359  =   ""
                  DateDefault     =   "00000000"
                  TimeDefault     =   "000000"
                  TimeStyle       =   0
                  BorderGrayAreaColor=   -2147483637
                  ThreeDOnFocusInvert=   0   'False
                  ThreeDFrameColor=   -2147483633
                  Appearance      =   2
                  BorderDropShadow=   0
                  BorderDropShadowColor=   -2147483632
                  BorderDropShadowWidth=   3
                  PopUpType       =   0
                  DateCalcY2KSplit=   60
                  CaretPosition   =   0
                  IncYear         =   1
                  IncMonth        =   1
                  IncDay          =   1
                  IncHour         =   1
                  IncMinute       =   1
                  IncSecond       =   1
                  ButtonColor     =   -2147483633
                  AutoMenu        =   0   'False
                  StartMonth      =   4
                  ButtonAlign     =   0
                  BoundDataType   =   0
                  OLEDropMode     =   0
                  OLEDragMode     =   0
               End
               Begin EditLib.fpDoubleSingle ipp_MtoHip_Dep1 
                  Height          =   315
                  Left            =   -73140
                  TabIndex        =   50
                  Top             =   2700
                  Width           =   1425
                  _Version        =   196608
                  _ExtentX        =   2514
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
               Begin EditLib.fpDateTime ipp_FecPre_Es2 
                  Height          =   315
                  Left            =   -73140
                  TabIndex        =   31
                  Top             =   720
                  Width           =   1425
                  _Version        =   196608
                  _ExtentX        =   2514
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
                  ThreeDInsideHighlightColor=   -2147483633
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
                  ButtonStyle     =   3
                  ButtonWidth     =   0
                  ButtonWrap      =   -1  'True
                  ButtonDefaultAction=   -1  'True
                  ThreeDText      =   0
                  ThreeDTextHighlightColor=   -2147483633
                  ThreeDTextShadowColor=   -2147483632
                  ThreeDTextOffset=   1
                  AlignTextH      =   0
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
                  Text            =   "28/09/2004"
                  DateCalcMethod  =   0
                  DateTimeFormat  =   0
                  UserDefinedFormat=   ""
                  DateMax         =   "00000000"
                  DateMin         =   "00000000"
                  TimeMax         =   "000000"
                  TimeMin         =   "000000"
                  TimeString1159  =   ""
                  TimeString2359  =   ""
                  DateDefault     =   "00000000"
                  TimeDefault     =   "000000"
                  TimeStyle       =   0
                  BorderGrayAreaColor=   -2147483637
                  ThreeDOnFocusInvert=   0   'False
                  ThreeDFrameColor=   -2147483633
                  Appearance      =   2
                  BorderDropShadow=   0
                  BorderDropShadowColor=   -2147483632
                  BorderDropShadowWidth=   3
                  PopUpType       =   0
                  DateCalcY2KSplit=   60
                  CaretPosition   =   0
                  IncYear         =   1
                  IncMonth        =   1
                  IncDay          =   1
                  IncHour         =   1
                  IncMinute       =   1
                  IncSecond       =   1
                  ButtonColor     =   -2147483633
                  AutoMenu        =   0   'False
                  StartMonth      =   4
                  ButtonAlign     =   0
                  BoundDataType   =   0
                  OLEDropMode     =   0
                  OLEDragMode     =   0
               End
               Begin EditLib.fpDateTime ipp_FecIns_Es2 
                  Height          =   315
                  Left            =   -73140
                  TabIndex        =   33
                  Top             =   1380
                  Width           =   1425
                  _Version        =   196608
                  _ExtentX        =   2514
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
                  ThreeDInsideHighlightColor=   -2147483633
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
                  ButtonStyle     =   3
                  ButtonWidth     =   0
                  ButtonWrap      =   -1  'True
                  ButtonDefaultAction=   -1  'True
                  ThreeDText      =   0
                  ThreeDTextHighlightColor=   -2147483633
                  ThreeDTextShadowColor=   -2147483632
                  ThreeDTextOffset=   1
                  AlignTextH      =   0
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
                  Text            =   "28/09/2004"
                  DateCalcMethod  =   0
                  DateTimeFormat  =   0
                  UserDefinedFormat=   ""
                  DateMax         =   "00000000"
                  DateMin         =   "00000000"
                  TimeMax         =   "000000"
                  TimeMin         =   "000000"
                  TimeString1159  =   ""
                  TimeString2359  =   ""
                  DateDefault     =   "00000000"
                  TimeDefault     =   "000000"
                  TimeStyle       =   0
                  BorderGrayAreaColor=   -2147483637
                  ThreeDOnFocusInvert=   0   'False
                  ThreeDFrameColor=   -2147483633
                  Appearance      =   2
                  BorderDropShadow=   0
                  BorderDropShadowColor=   -2147483632
                  BorderDropShadowWidth=   3
                  PopUpType       =   0
                  DateCalcY2KSplit=   60
                  CaretPosition   =   0
                  IncYear         =   1
                  IncMonth        =   1
                  IncDay          =   1
                  IncHour         =   1
                  IncMinute       =   1
                  IncSecond       =   1
                  ButtonColor     =   -2147483633
                  AutoMenu        =   0   'False
                  StartMonth      =   4
                  ButtonAlign     =   0
                  BoundDataType   =   0
                  OLEDropMode     =   0
                  OLEDragMode     =   0
               End
               Begin EditLib.fpDoubleSingle ipp_MtoHip_Es2 
                  Height          =   315
                  Left            =   -73140
                  TabIndex        =   39
                  Top             =   2700
                  Width           =   1425
                  _Version        =   196608
                  _ExtentX        =   2514
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
               Begin EditLib.fpDateTime ipp_FecPre_Dep2 
                  Height          =   315
                  Left            =   -73140
                  TabIndex        =   53
                  Top             =   720
                  Width           =   1425
                  _Version        =   196608
                  _ExtentX        =   2514
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
                  ThreeDInsideHighlightColor=   -2147483633
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
                  ButtonStyle     =   3
                  ButtonWidth     =   0
                  ButtonWrap      =   -1  'True
                  ButtonDefaultAction=   -1  'True
                  ThreeDText      =   0
                  ThreeDTextHighlightColor=   -2147483633
                  ThreeDTextShadowColor=   -2147483632
                  ThreeDTextOffset=   1
                  AlignTextH      =   0
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
                  Text            =   "28/09/2004"
                  DateCalcMethod  =   0
                  DateTimeFormat  =   0
                  UserDefinedFormat=   ""
                  DateMax         =   "00000000"
                  DateMin         =   "00000000"
                  TimeMax         =   "000000"
                  TimeMin         =   "000000"
                  TimeString1159  =   ""
                  TimeString2359  =   ""
                  DateDefault     =   "00000000"
                  TimeDefault     =   "000000"
                  TimeStyle       =   0
                  BorderGrayAreaColor=   -2147483637
                  ThreeDOnFocusInvert=   0   'False
                  ThreeDFrameColor=   -2147483633
                  Appearance      =   2
                  BorderDropShadow=   0
                  BorderDropShadowColor=   -2147483632
                  BorderDropShadowWidth=   3
                  PopUpType       =   0
                  DateCalcY2KSplit=   60
                  CaretPosition   =   0
                  IncYear         =   1
                  IncMonth        =   1
                  IncDay          =   1
                  IncHour         =   1
                  IncMinute       =   1
                  IncSecond       =   1
                  ButtonColor     =   -2147483633
                  AutoMenu        =   0   'False
                  StartMonth      =   4
                  ButtonAlign     =   0
                  BoundDataType   =   0
                  OLEDropMode     =   0
                  OLEDragMode     =   0
               End
               Begin EditLib.fpDateTime ipp_FecIns_Dep2 
                  Height          =   315
                  Left            =   -73140
                  TabIndex        =   55
                  Top             =   1380
                  Width           =   1425
                  _Version        =   196608
                  _ExtentX        =   2514
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
                  ThreeDInsideHighlightColor=   -2147483633
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
                  ButtonStyle     =   3
                  ButtonWidth     =   0
                  ButtonWrap      =   -1  'True
                  ButtonDefaultAction=   -1  'True
                  ThreeDText      =   0
                  ThreeDTextHighlightColor=   -2147483633
                  ThreeDTextShadowColor=   -2147483632
                  ThreeDTextOffset=   1
                  AlignTextH      =   0
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
                  Text            =   "28/09/2004"
                  DateCalcMethod  =   0
                  DateTimeFormat  =   0
                  UserDefinedFormat=   ""
                  DateMax         =   "00000000"
                  DateMin         =   "00000000"
                  TimeMax         =   "000000"
                  TimeMin         =   "000000"
                  TimeString1159  =   ""
                  TimeString2359  =   ""
                  DateDefault     =   "00000000"
                  TimeDefault     =   "000000"
                  TimeStyle       =   0
                  BorderGrayAreaColor=   -2147483637
                  ThreeDOnFocusInvert=   0   'False
                  ThreeDFrameColor=   -2147483633
                  Appearance      =   2
                  BorderDropShadow=   0
                  BorderDropShadowColor=   -2147483632
                  BorderDropShadowWidth=   3
                  PopUpType       =   0
                  DateCalcY2KSplit=   60
                  CaretPosition   =   0
                  IncYear         =   1
                  IncMonth        =   1
                  IncDay          =   1
                  IncHour         =   1
                  IncMinute       =   1
                  IncSecond       =   1
                  ButtonColor     =   -2147483633
                  AutoMenu        =   0   'False
                  StartMonth      =   4
                  ButtonAlign     =   0
                  BoundDataType   =   0
                  OLEDropMode     =   0
                  OLEDragMode     =   0
               End
               Begin EditLib.fpDoubleSingle ipp_MtoHip_Dep2 
                  Height          =   315
                  Left            =   -73140
                  TabIndex        =   61
                  Top             =   2700
                  Width           =   1425
                  _Version        =   196608
                  _ExtentX        =   2514
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
               Begin VB.Label Label79 
                  Caption         =   "Moneda:"
                  Height          =   195
                  Left            =   -70410
                  TabIndex        =   160
                  Top             =   2760
                  Width           =   795
               End
               Begin VB.Label Label78 
                  Caption         =   "Depósito 2:"
                  Height          =   195
                  Left            =   -74880
                  TabIndex        =   159
                  Top             =   450
                  Width           =   1365
               End
               Begin VB.Label Label77 
                  Caption         =   "Tipo Doc. Registral:"
                  Height          =   225
                  Left            =   -74880
                  TabIndex        =   158
                  Top             =   1755
                  Width           =   1485
               End
               Begin VB.Label Label76 
                  Caption         =   "Partida Electrónica:"
                  Height          =   225
                  Left            =   -74880
                  TabIndex        =   157
                  Top             =   2085
                  Width           =   1485
               End
               Begin VB.Label Label75 
                  Caption         =   "Asiento:"
                  Height          =   225
                  Left            =   -70410
                  TabIndex        =   156
                  Top             =   2085
                  Width           =   765
               End
               Begin VB.Label Label74 
                  Caption         =   "Ficha Registral:"
                  Height          =   225
                  Left            =   -74880
                  TabIndex        =   155
                  Top             =   2415
                  Width           =   1485
               End
               Begin VB.Label Label73 
                  Caption         =   "Asiento:"
                  Height          =   225
                  Left            =   -70410
                  TabIndex        =   154
                  Top             =   2415
                  Width           =   765
               End
               Begin VB.Label Label72 
                  Caption         =   "Fecha Presentación:"
                  Height          =   225
                  Left            =   -74880
                  TabIndex        =   153
                  Top             =   765
                  Width           =   1575
               End
               Begin VB.Label Label71 
                  Caption         =   "Nro. Presentación:"
                  Height          =   225
                  Left            =   -74880
                  TabIndex        =   152
                  Top             =   1095
                  Width           =   1485
               End
               Begin VB.Label Label28 
                  Caption         =   "Fecha Inscripción:"
                  Height          =   225
                  Left            =   -74880
                  TabIndex        =   151
                  Top             =   1425
                  Width           =   1425
               End
               Begin VB.Label Label27 
                  Caption         =   "Monto Hipoteca:"
                  Height          =   225
                  Left            =   -74880
                  TabIndex        =   150
                  Top             =   2745
                  Width           =   1305
               End
               Begin VB.Label Label70 
                  Caption         =   "Moneda:"
                  Height          =   195
                  Left            =   -70410
                  TabIndex        =   149
                  Top             =   2760
                  Width           =   795
               End
               Begin VB.Label Label69 
                  Caption         =   "Estacionamiento 2:"
                  Height          =   225
                  Left            =   -74880
                  TabIndex        =   148
                  Top             =   435
                  Width           =   1365
               End
               Begin VB.Label Label68 
                  Caption         =   "Tipo Doc. Registral:"
                  Height          =   225
                  Left            =   -74880
                  TabIndex        =   147
                  Top             =   1755
                  Width           =   1485
               End
               Begin VB.Label Label67 
                  Caption         =   "Partida Electrónica:"
                  Height          =   225
                  Left            =   -74880
                  TabIndex        =   146
                  Top             =   2085
                  Width           =   1485
               End
               Begin VB.Label Label66 
                  Caption         =   "Asiento:"
                  Height          =   225
                  Left            =   -70410
                  TabIndex        =   145
                  Top             =   2085
                  Width           =   765
               End
               Begin VB.Label Label65 
                  Caption         =   "Ficha Registral:"
                  Height          =   225
                  Left            =   -74880
                  TabIndex        =   144
                  Top             =   2415
                  Width           =   1485
               End
               Begin VB.Label Label64 
                  Caption         =   "Fecha Presentación:"
                  Height          =   225
                  Left            =   -74880
                  TabIndex        =   143
                  Top             =   765
                  Width           =   1575
               End
               Begin VB.Label Label63 
                  Caption         =   "Nro. Presentación:"
                  Height          =   225
                  Left            =   -74880
                  TabIndex        =   142
                  Top             =   1095
                  Width           =   1485
               End
               Begin VB.Label Label39 
                  Caption         =   "Fecha Inscripción:"
                  Height          =   225
                  Left            =   -74880
                  TabIndex        =   141
                  Top             =   1425
                  Width           =   1425
               End
               Begin VB.Label Label38 
                  Caption         =   "Monto Hipoteca:"
                  Height          =   225
                  Left            =   -74880
                  TabIndex        =   140
                  Top             =   2745
                  Width           =   1305
               End
               Begin VB.Label Label37 
                  Caption         =   "Asiento:"
                  Height          =   225
                  Left            =   -70410
                  TabIndex        =   139
                  Top             =   2415
                  Width           =   765
               End
               Begin VB.Label Label25 
                  Caption         =   "Moneda:"
                  Height          =   195
                  Left            =   -70410
                  TabIndex        =   138
                  Top             =   2760
                  Width           =   795
               End
               Begin VB.Label Label50 
                  Caption         =   "Depósito 1:"
                  Height          =   195
                  Left            =   -74880
                  TabIndex        =   137
                  Top             =   450
                  Width           =   1365
               End
               Begin VB.Label Label40 
                  Caption         =   "Tipo Doc. Registral:"
                  Height          =   225
                  Left            =   -74880
                  TabIndex        =   136
                  Top             =   1755
                  Width           =   1485
               End
               Begin VB.Label Label41 
                  Caption         =   "Partida Electrónica:"
                  Height          =   225
                  Left            =   -74880
                  TabIndex        =   135
                  Top             =   2085
                  Width           =   1485
               End
               Begin VB.Label Label42 
                  Caption         =   "Asiento:"
                  Height          =   225
                  Left            =   -70410
                  TabIndex        =   134
                  Top             =   2085
                  Width           =   765
               End
               Begin VB.Label Label43 
                  Caption         =   "Ficha Registral:"
                  Height          =   225
                  Left            =   -74880
                  TabIndex        =   133
                  Top             =   2415
                  Width           =   1485
               End
               Begin VB.Label Label44 
                  Caption         =   "Asiento:"
                  Height          =   225
                  Left            =   -70410
                  TabIndex        =   132
                  Top             =   2415
                  Width           =   765
               End
               Begin VB.Label Label51 
                  Caption         =   "Fecha Presentación:"
                  Height          =   225
                  Left            =   -74880
                  TabIndex        =   131
                  Top             =   765
                  Width           =   1575
               End
               Begin VB.Label Label52 
                  Caption         =   "Nro. Presentación:"
                  Height          =   225
                  Left            =   -74880
                  TabIndex        =   130
                  Top             =   1095
                  Width           =   1485
               End
               Begin VB.Label Label53 
                  Caption         =   "Fecha Inscripción:"
                  Height          =   225
                  Left            =   -74880
                  TabIndex        =   129
                  Top             =   1425
                  Width           =   1425
               End
               Begin VB.Label Label58 
                  Caption         =   "Monto Hipoteca:"
                  Height          =   225
                  Left            =   -74880
                  TabIndex        =   128
                  Top             =   2745
                  Width           =   1305
               End
               Begin VB.Label Label54 
                  Caption         =   "Nro. Tomo:"
                  Height          =   285
                  Left            =   -74910
                  TabIndex        =   127
                  Top             =   2700
                  Width           =   1485
               End
               Begin VB.Label Label21 
                  Caption         =   "Nro. Libro:"
                  Height          =   285
                  Left            =   -66960
                  TabIndex        =   126
                  Top             =   2700
                  Width           =   795
               End
               Begin VB.Label Label13 
                  Caption         =   "Asiento:"
                  Height          =   285
                  Left            =   -70440
                  TabIndex        =   125
                  Top             =   2370
                  Width           =   765
               End
               Begin VB.Label Label12 
                  Caption         =   "Estacionamiento 1:"
                  Height          =   315
                  Left            =   -74910
                  TabIndex        =   124
                  Top             =   390
                  Width           =   1365
               End
               Begin VB.Label Label11 
                  Caption         =   "Asiento:"
                  Height          =   225
                  Left            =   -70410
                  TabIndex        =   123
                  Top             =   2415
                  Width           =   765
               End
               Begin VB.Label Label56 
                  Caption         =   "Monto Hipoteca:"
                  Height          =   225
                  Left            =   -74880
                  TabIndex        =   122
                  Top             =   2745
                  Width           =   1305
               End
               Begin VB.Label Label49 
                  Caption         =   "Fecha Inscripción:"
                  Height          =   225
                  Left            =   -74880
                  TabIndex        =   121
                  Top             =   1425
                  Width           =   1425
               End
               Begin VB.Label Label48 
                  Caption         =   "Nro. Presentación:"
                  Height          =   225
                  Left            =   -74880
                  TabIndex        =   120
                  Top             =   1095
                  Width           =   1485
               End
               Begin VB.Label Label23 
                  Caption         =   "Fecha Presentación:"
                  Height          =   225
                  Left            =   -74880
                  TabIndex        =   119
                  Top             =   765
                  Width           =   1575
               End
               Begin VB.Label Label17 
                  Caption         =   "Ficha Registral:"
                  Height          =   225
                  Left            =   -74880
                  TabIndex        =   118
                  Top             =   2415
                  Width           =   1485
               End
               Begin VB.Label Label18 
                  Caption         =   "Asiento:"
                  Height          =   225
                  Left            =   -70410
                  TabIndex        =   117
                  Top             =   2085
                  Width           =   765
               End
               Begin VB.Label Label19 
                  Caption         =   "Partida Electrónica:"
                  Height          =   225
                  Left            =   -74880
                  TabIndex        =   116
                  Top             =   2085
                  Width           =   1485
               End
               Begin VB.Label Label20 
                  Caption         =   "Tipo Doc. Registral:"
                  Height          =   225
                  Left            =   -74880
                  TabIndex        =   115
                  Top             =   1755
                  Width           =   1485
               End
               Begin VB.Label Label7 
                  Caption         =   "Ficha Registral:"
                  Height          =   255
                  Left            =   -74880
                  TabIndex        =   114
                  Top             =   2280
                  Width           =   1485
               End
               Begin VB.Label Label6 
                  Caption         =   "Asiento:"
                  Height          =   285
                  Left            =   -70410
                  TabIndex        =   113
                  Top             =   2280
                  Width           =   765
               End
               Begin VB.Label Label55 
                  Caption         =   "Monto Hipoteca:"
                  Height          =   255
                  Left            =   -74880
                  TabIndex        =   112
                  Top             =   2640
                  Width           =   1305
               End
               Begin VB.Label Label47 
                  Caption         =   "Fecha Inscripción:"
                  Height          =   225
                  Left            =   -74880
                  TabIndex        =   111
                  Top             =   1215
                  Width           =   1425
               End
               Begin VB.Label Label46 
                  Caption         =   "Nro. Presentación:"
                  Height          =   255
                  Left            =   -74880
                  TabIndex        =   110
                  Top             =   840
                  Width           =   1485
               End
               Begin VB.Label Label15 
                  Caption         =   "Fecha Presentación:"
                  Height          =   255
                  Left            =   -74880
                  TabIndex        =   109
                  Top             =   480
                  Width           =   1695
               End
               Begin VB.Label Label8 
                  Caption         =   "Asiento:"
                  Height          =   285
                  Left            =   -70410
                  TabIndex        =   108
                  Top             =   1905
                  Width           =   765
               End
               Begin VB.Label Label9 
                  Caption         =   "Partida Electrónica:"
                  Height          =   255
                  Left            =   -74880
                  TabIndex        =   107
                  Top             =   1920
                  Width           =   1485
               End
               Begin VB.Label Label10 
                  Caption         =   "Tipo Doc. Registral:"
                  Height          =   255
                  Left            =   -74880
                  TabIndex        =   106
                  Top             =   1560
                  Width           =   1485
               End
               Begin VB.Label Label14 
                  Caption         =   "Estacionamiento 1:"
                  Height          =   225
                  Left            =   -74880
                  TabIndex        =   105
                  Top             =   435
                  Width           =   1365
               End
               Begin VB.Label Label22 
                  Caption         =   "Moneda:"
                  Height          =   225
                  Left            =   -70410
                  TabIndex        =   104
                  Top             =   2640
                  Width           =   795
               End
               Begin VB.Label Label24 
                  Caption         =   "Moneda:"
                  Height          =   195
                  Left            =   -70410
                  TabIndex        =   103
                  Top             =   2760
                  Width           =   795
               End
               Begin VB.Label Label16 
                  Caption         =   "Fecha Emisión:"
                  Height          =   255
                  Left            =   150
                  TabIndex        =   102
                  Top             =   780
                  Width           =   1245
               End
               Begin VB.Label Label4 
                  Caption         =   "Moneda:"
                  Height          =   225
                  Left            =   150
                  TabIndex        =   101
                  Top             =   1155
                  Width           =   1545
               End
               Begin VB.Label Label5 
                  Caption         =   "Importe:"
                  Height          =   225
                  Left            =   150
                  TabIndex        =   100
                  Top             =   1515
                  Width           =   1545
               End
               Begin VB.Label Label2 
                  Caption         =   "N° Carta Fianzas:"
                  Height          =   255
                  Left            =   150
                  TabIndex        =   99
                  Top             =   1860
                  Width           =   1545
               End
            End
         End
         Begin Threed.SSPanel SSPanel12 
            Height          =   495
            Left            =   60
            TabIndex        =   161
            Top             =   9480
            Width           =   14505
            _Version        =   65536
            _ExtentX        =   25585
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
            Begin Threed.SSPanel pnl_TotHip 
               Height          =   315
               Left            =   1950
               TabIndex        =   162
               Top             =   90
               Width           =   1425
               _Version        =   65536
               _ExtentX        =   2514
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "0.00 "
               ForeColor       =   32768
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
               Font3D          =   2
               Alignment       =   4
            End
            Begin VB.Label Label60 
               Caption         =   "Total Hipoteca:"
               Height          =   195
               Left            =   150
               TabIndex        =   163
               Top             =   150
               Width           =   1485
            End
         End
         Begin Threed.SSPanel SSPanel13 
            Height          =   1035
            Left            =   60
            TabIndex        =   164
            Top             =   5190
            Width           =   14505
            _Version        =   65536
            _ExtentX        =   25585
            _ExtentY        =   1826
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
            Begin VB.ComboBox cmb_FecEva 
               Height          =   315
               Left            =   12480
               Style           =   2  'Dropdown List
               TabIndex        =   4
               Top             =   660
               Width           =   1875
            End
            Begin VB.ComboBox cmb_NumInf 
               Height          =   315
               Left            =   9210
               Style           =   2  'Dropdown List
               TabIndex        =   3
               Top             =   660
               Width           =   1875
            End
            Begin VB.ComboBox cmb_EmpPer 
               Height          =   315
               Left            =   1860
               Style           =   2  'Dropdown List
               TabIndex        =   2
               Top             =   660
               Width           =   5025
            End
            Begin VB.ComboBox Cmb_TipGar 
               Height          =   315
               Left            =   1860
               Style           =   2  'Dropdown List
               TabIndex        =   0
               Top             =   60
               Width           =   5025
            End
            Begin VB.ComboBox cmb_SedReg 
               Height          =   315
               Left            =   9210
               Style           =   2  'Dropdown List
               TabIndex        =   1
               Top             =   60
               Width           =   5145
            End
            Begin VB.Label Label31 
               AutoSize        =   -1  'True
               Caption         =   "F. Tasación:"
               Height          =   195
               Left            =   11400
               TabIndex        =   170
               Top             =   720
               Width           =   855
            End
            Begin VB.Label Label30 
               AutoSize        =   -1  'True
               Caption         =   "Número de Informe:"
               Height          =   195
               Left            =   7620
               TabIndex        =   169
               Top             =   720
               Width           =   1395
            End
            Begin VB.Label Label29 
               AutoSize        =   -1  'True
               Caption         =   "Tasación"
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
               Left            =   120
               TabIndex        =   168
               Top             =   420
               Width           =   795
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               Caption         =   "Empresa Peritaje:"
               Height          =   195
               Left            =   120
               TabIndex        =   167
               Top             =   720
               Width           =   1230
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Tipo de Garantía:"
               Height          =   195
               Left            =   120
               TabIndex        =   166
               Top             =   120
               Width           =   1260
            End
            Begin VB.Label Label45 
               AutoSize        =   -1  'True
               Caption         =   "Sede Registral:"
               Height          =   195
               Left            =   7590
               TabIndex        =   165
               Top             =   120
               Width           =   1080
            End
         End
      End
   End
End
Attribute VB_Name = "frm_Ges_TecPro_05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_int_NumIte     As Integer
Dim l_arr_EmpPer()   As moddat_tpo_Genera
'Dim l_dbl_MtoHip     As Double

Private Sub cmb_EmpPer_Click()
  If cmb_EmpPer.ListIndex > -1 Then
      Screen.MousePointer = 11
      cmb_NumInf.Clear
      cmb_FecEva.Clear
      Call fs_Buscar_NumInf_FecTas(cmb_NumInf, cmb_FecEva, Format(cmb_EmpPer.ItemData(cmb_EmpPer.ListIndex), "000000"), 0)
      Call gs_SetFocus(cmb_NumInf)
      Screen.MousePointer = 0
  End If
End Sub

Private Sub cmb_EmpPer_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_EmpPer_Click
   End If
End Sub

Private Sub cmb_FecEva_Click()
   tab_Genera.Tab = 1
   Call gs_SetFocus(ipp_FecPre_Inm)
End Sub

Private Sub cmb_FecEva_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_FecEva_Click
   End If
End Sub

Private Sub cmb_FlgEst_Dep1_Click()
   If cmb_FlgEst_Dep1.ListIndex > -1 Then
      If cmb_FlgEst_Dep1.ItemData(cmb_FlgEst_Dep1.ListIndex) = 1 Then
         ipp_FecPre_Dep1.Enabled = True
         txt_NumPre_Dep1.Enabled = True
         ipp_FecIns_Dep1.Enabled = True
         cmb_TipDoc_Dep1.Enabled = True
         txt_NumPar_Dep1.Enabled = True
         txt_NumAPa_Dep1.Enabled = True
         txt_NumFic_Dep1.Enabled = True
         txt_NumAFi_Dep1.Enabled = True
'         txt_NumTom_Dep1.Enabled = True
'         txt_NumFoj_Dep1.Enabled = True
'         txt_NumLib_Dep1.Enabled = True
         cmb_Moneda_Dep1.Enabled = True
         ipp_MtoHip_Dep1.Enabled = True
         
         Call gs_SetFocus(ipp_FecPre_Dep1)
      Else
         ipp_FecPre_Dep1.Text = Format(date, "dd/mm/yyyy")
         txt_NumPre_Dep1.Text = ""
         ipp_FecIns_Dep1.Text = Format(date, "dd/mm/yyyy")
         cmb_TipDoc_Dep1.ListIndex = -1
         txt_NumPar_Dep1.Text = ""
         txt_NumAPa_Dep1.Text = ""
         txt_NumFic_Dep1.Text = ""
         txt_NumAFi_Dep1.Text = ""
'         txt_NumTom_Dep1.Text = ""
'         txt_NumFoj_Dep1.Text = ""
'         txt_NumLib_Dep1.Text = ""
         cmb_Moneda_Dep1.ListIndex = -1
         ipp_MtoHip_Dep1.Value = 0
         
         
         ipp_FecPre_Dep1.Enabled = False
         txt_NumPre_Dep1.Enabled = False
         ipp_FecIns_Dep1.Enabled = False
         cmb_TipDoc_Dep1.Enabled = False
         txt_NumPar_Dep1.Enabled = False
         txt_NumAPa_Dep1.Enabled = False
         txt_NumFic_Dep1.Enabled = False
         txt_NumAFi_Dep1.Enabled = False
'         txt_NumTom_Dep1.Enabled = False
'         txt_NumFoj_Dep1.Enabled = False
'         txt_NumLib_Dep1.Enabled = False
         cmb_Moneda_Dep1.Enabled = False
         ipp_MtoHip_Dep1.Enabled = False
         
         tab_Genera.Tab = 5
         Call gs_SetFocus(cmd_Grabar)
      End If
   End If
End Sub

Private Sub cmb_FlgEst_Dep1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_FlgEst_Dep1_Click
   End If
End Sub



Private Sub cmb_FlgEst_Dep2_Click()
 If cmb_FlgEst_Dep2.ListIndex > -1 Then
      If cmb_FlgEst_Dep2.ItemData(cmb_FlgEst_Dep2.ListIndex) = 1 Then
         ipp_FecPre_Dep2.Enabled = True
         txt_NumPre_Dep2.Enabled = True
         ipp_FecIns_Dep2.Enabled = True
         cmb_TipDoc_Dep2.Enabled = True
         txt_NumPar_Dep2.Enabled = True
         txt_NumAPa_Dep2.Enabled = True
         txt_NumFic_Dep2.Enabled = True
         txt_NumAFi_Dep2.Enabled = True
'         txt_NumTom_Dep2.Enabled = True
'         txt_NumFoj_Dep2.Enabled = True
'         txt_NumLib_Dep2.Enabled = True
         cmb_Moneda_Dep2.Enabled = True
         ipp_MtoHip_Dep2.Enabled = True
         
         Call gs_SetFocus(ipp_FecPre_Dep2)
      Else
         ipp_FecPre_Dep2.Text = Format(date, "dd/mm/yyyy")
         txt_NumPre_Dep2.Text = ""
         ipp_FecIns_Dep2.Text = Format(date, "dd/mm/yyyy")
         cmb_TipDoc_Dep2.ListIndex = -1
         txt_NumPar_Dep2.Text = ""
         txt_NumAPa_Dep2.Text = ""
         txt_NumFic_Dep2.Text = ""
         txt_NumAFi_Dep2.Text = ""
'         txt_NumTom_Dep2.Text = ""
'         txt_NumFoj_Dep2.Text = ""
'         txt_NumLib_Dep2.Text = ""
         cmb_Moneda_Dep2.ListIndex = -1
         ipp_MtoHip_Dep2.Value = 0
         
         
         ipp_FecPre_Dep2.Enabled = False
         txt_NumPre_Dep2.Enabled = False
         ipp_FecIns_Dep2.Enabled = False
         cmb_TipDoc_Dep2.Enabled = False
         txt_NumPar_Dep2.Enabled = False
         txt_NumAPa_Dep2.Enabled = False
         txt_NumFic_Dep2.Enabled = False
         txt_NumAFi_Dep2.Enabled = False
'         txt_NumTom_Dep2.Enabled = False
'         txt_NumFoj_Dep2.Enabled = False
'         txt_NumLib_Dep2.Enabled = False
         cmb_Moneda_Dep2.Enabled = False
         ipp_MtoHip_Dep2.Enabled = False
         
         tab_Genera.Tab = 1
         Call gs_SetFocus(cmd_Grabar)
      End If
   End If
End Sub

Private Sub cmb_FlgEst_Dep2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_FlgEst_Dep2_Click
   End If
End Sub

Private Sub cmb_FlgEst_Es1_Click()
   If cmb_FlgEst_Es1.ListIndex > -1 Then
      If cmb_FlgEst_Es1.ItemData(cmb_FlgEst_Es1.ListIndex) = 1 Then
         ipp_FecPre_Es1.Enabled = True
         txt_NumPre_Es1.Enabled = True
         ipp_FecIns_Es1.Enabled = True
         cmb_TipDoc_Es1.Enabled = True
         txt_NumPar_Es1.Enabled = True
         txt_NumAPa_Es1.Enabled = True
         txt_NumFic_Es1.Enabled = True
         txt_NumAFi_Es1.Enabled = True
         ipp_MtoHip_Es1.Enabled = True
         cmb_Moneda_Es1.Enabled = True
         
         Call gs_SetFocus(ipp_FecPre_Es1)
      Else
         ipp_FecPre_Es1.Text = Format(date, "dd/mm/yyyy")
         txt_NumPre_Es1.Text = ""
         ipp_FecIns_Es1.Text = Format(date, "dd/mm/yyyy")
         cmb_TipDoc_Es1.ListIndex = -1
         txt_NumPar_Es1.Text = ""
         txt_NumAPa_Es1.Text = ""
         txt_NumFic_Es1.Text = ""
         txt_NumAFi_Es1.Text = ""
         ipp_MtoHip_Es1.Value = 0
         cmb_Moneda_Es1.ListIndex = -1
         
         ipp_FecPre_Es1.Enabled = False
         txt_NumPre_Es1.Enabled = False
         ipp_FecIns_Es1.Enabled = False
         cmb_TipDoc_Es1.Enabled = False
         txt_NumPar_Es1.Enabled = False
         txt_NumAPa_Es1.Enabled = False
         txt_NumFic_Es1.Enabled = False
         txt_NumAFi_Es1.Enabled = False
         ipp_MtoHip_Es1.Enabled = False
         cmb_Moneda_Es1.Enabled = False
         
         tab_Genera.Tab = 3
         Call gs_SetFocus(cmb_FlgEst_Dep1)
      End If
   End If
End Sub

Private Sub cmb_FlgEst_Es1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_FlgEst_Es1_Click
   End If
End Sub

Private Sub cmb_FlgEst_Es2_Click()
 If cmb_FlgEst_Es2.ListIndex > -1 Then
      If cmb_FlgEst_Es2.ItemData(cmb_FlgEst_Es2.ListIndex) = 1 Then
         ipp_FecPre_Es2.Enabled = True
         txt_NumPre_Es2.Enabled = True
         ipp_FecIns_Es2.Enabled = True
         cmb_TipDoc_Es2.Enabled = True
         txt_NumPar_Es2.Enabled = True
         txt_NumAPa_Es2.Enabled = True
         txt_NumFic_Es2.Enabled = True
         txt_NumAFi_Es2.Enabled = True
         ipp_MtoHip_Es2.Enabled = True
         cmb_Moneda_Es2.Enabled = True
         
         Call gs_SetFocus(ipp_FecPre_Es2)
      Else
         ipp_FecPre_Es2.Text = Format(date, "dd/mm/yyyy")
         txt_NumPre_Es2.Text = ""
         ipp_FecIns_Es2.Text = Format(date, "dd/mm/yyyy")
         cmb_TipDoc_Es2.ListIndex = -1
         txt_NumPar_Es2.Text = ""
         txt_NumAPa_Es2.Text = ""
         txt_NumFic_Es2.Text = ""
         txt_NumAFi_Es2.Text = ""
         ipp_MtoHip_Es2.Value = 0
         cmb_Moneda_Es2.ListIndex = -1
         
         ipp_FecPre_Es2.Enabled = False
         txt_NumPre_Es2.Enabled = False
         ipp_FecIns_Es2.Enabled = False
         cmb_TipDoc_Es2.Enabled = False
         txt_NumPar_Es2.Enabled = False
         txt_NumAPa_Es2.Enabled = False
         txt_NumFic_Es2.Enabled = False
         txt_NumAFi_Es2.Enabled = False
         ipp_MtoHip_Es2.Enabled = False
         cmb_Moneda_Es2.Enabled = False
         
         tab_Genera.Tab = 4
         Call gs_SetFocus(cmb_FlgEst_Dep1)
      End If
   End If
End Sub

Private Sub cmb_FlgEst_Es2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_FlgEst_Es2_Click
   End If
End Sub

Private Sub cmb_Moneda_Click()
   Call gs_SetFocus(ipp_ImpGar)
End Sub


Private Sub cmb_Moneda_Dep1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      tab_Genera.Tab = 5
      Call gs_SetFocus(cmb_FlgEst_Dep2)
   End If
End Sub



Private Sub cmb_Moneda_Dep2_KeyPress(KeyAscii As Integer)
   Call gs_SetFocus(cmd_Grabar)
End Sub

Private Sub cmb_Moneda_Es1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      tab_Genera.Tab = 3
      Call gs_SetFocus(cmb_FlgEst_Es2)
   End If
End Sub

Private Sub cmb_Moneda_Es2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      tab_Genera.Tab = 4
      Call gs_SetFocus(cmb_FlgEst_Dep1)
   End If
End Sub

Private Sub cmb_Moneda_Inm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      tab_Genera.Tab = 2
      Call gs_SetFocus(cmb_FlgEst_Es1)
   End If
End Sub

Private Sub cmb_Moneda_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ImpGar)
   End If
End Sub

Private Sub cmb_NumInf_Click()
  If cmb_NumInf.ListIndex > -1 Then
      Screen.MousePointer = 11
      cmb_FecEva.Clear
      Call fs_Buscar_NumInf_FecTas(cmb_NumInf, cmb_FecEva, Format(cmb_EmpPer.ItemData(cmb_EmpPer.ListIndex), "000000"), 1)
      Call gs_SetFocus(cmb_FecEva)
      Screen.MousePointer = 0
  End If
End Sub

Private Sub cmb_NumInf_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_NumInf_Click
   End If
End Sub

Private Sub cmb_NumRef_Click()
   If cmb_NumRef.ListIndex <> -1 Then
      If moddat_g_int_FlgAct_2 = 1 Then 'moddat_g_int_FlgGrb_2
         pnl_NumCFi.Caption = Trim(cmb_NumRef.Text)
      Else
         pnl_NumCFi.Caption = pnl_NumCFi.Caption & "|" & Trim(cmb_NumRef.Text)
      End If
   End If
   Call gs_SetFocus(cmd_CfiAso)
End Sub

Private Sub cmb_NumRef_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     Call gs_SetFocus(cmd_Grabar)
   End If
End Sub

Private Sub cmb_SedReg_Click()
   Call gs_SetFocus(cmb_EmpPer)
End Sub

Private Sub cmb_SedReg_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_SedReg_Click
   End If
End Sub

Private Sub cmb_TipDoc_Dep1_Click()
   If cmb_TipDoc_Dep1.ListIndex > -1 Then
      Select Case cmb_TipDoc_Dep1.ItemData(cmb_TipDoc_Dep1.ListIndex)
         Case 1
            txt_NumPar_Dep1.Enabled = True
            txt_NumAPa_Dep1.Enabled = True
            
            Call gs_SetFocus(txt_NumPar_Dep1)
            
            txt_NumFic_Dep1.Enabled = False
            txt_NumAFi_Dep1.Enabled = False
            txt_NumFic_Dep1.Text = ""
            txt_NumAFi_Dep1.Text = ""
            
'            txt_NumTom_Es1.Enabled = False
'            txt_NumFoj_Es1.Enabled = False
'            txt_NumLib_Es1.Enabled = False
'            txt_NumTom_Es1.Text = ""
'            txt_NumFoj_Es1.Text = ""
'            txt_NumLib_Es1.Text = ""
            
         Case 2
            txt_NumPar_Dep1.Enabled = False
            txt_NumAPa_Dep1.Enabled = False
            txt_NumPar_Dep1.Text = ""
            txt_NumAPa_Dep1.Text = ""
            
            
            txt_NumFic_Dep1.Enabled = True
            txt_NumAFi_Dep1.Enabled = True
            
            Call gs_SetFocus(txt_NumFic_Dep1)
            
'            txt_NumTom_Es1.Enabled = False
'            txt_NumFoj_Es1.Enabled = False
'            txt_NumLib_Es1.Enabled = False
'            txt_NumTom_Es1.Text = ""
'            txt_NumFoj_Es1.Text = ""
'            txt_NumLib_Es1.Text = ""
            
         Case 3
            txt_NumPar_Dep1.Enabled = False
            txt_NumAPa_Dep1.Enabled = False
            txt_NumPar_Dep1.Text = ""
            txt_NumAPa_Dep1.Text = ""
            
            txt_NumFic_Dep1.Enabled = False
            txt_NumAFi_Dep1.Enabled = False
            txt_NumFic_Dep1.Text = ""
            txt_NumAFi_Dep1.Text = ""
            
'            txt_NumTom_Es1.Enabled = True
'            txt_NumFoj_Es1.Enabled = True
'            txt_NumLib_Es1.Enabled = True
            
            'Call gs_SetFocus(txt_NumTom_Es1)
      End Select
   End If
End Sub

Private Sub cmb_TipDoc_Dep1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipDoc_Dep1_Click
   End If
End Sub

Private Sub cmb_TipDoc_Dep2_Click()
   If cmb_TipDoc_Dep2.ListIndex > -1 Then
      Select Case cmb_TipDoc_Dep2.ItemData(cmb_TipDoc_Dep2.ListIndex)
         Case 1
            txt_NumPar_Dep2.Enabled = True
            txt_NumAPa_Dep2.Enabled = True
            
            Call gs_SetFocus(txt_NumPar_Dep2)
            
            txt_NumFic_Dep2.Enabled = False
            txt_NumAFi_Dep2.Enabled = False
            txt_NumFic_Dep2.Text = ""
            txt_NumAFi_Dep2.Text = ""
            
'            txt_NumTom_Es1.Enabled = False
'            txt_NumFoj_Es1.Enabled = False
'            txt_NumLib_Es1.Enabled = False
'            txt_NumTom_Es1.Text = ""
'            txt_NumFoj_Es1.Text = ""
'            txt_NumLib_Es1.Text = ""
            
         Case 2
            txt_NumPar_Dep2.Enabled = False
            txt_NumAPa_Dep2.Enabled = False
            txt_NumPar_Dep2.Text = ""
            txt_NumAPa_Dep2.Text = ""
            
            
            txt_NumFic_Dep2.Enabled = True
            txt_NumAFi_Dep2.Enabled = True
            
            Call gs_SetFocus(txt_NumFic_Dep2)
            
'            txt_NumTom_Es1.Enabled = False
'            txt_NumFoj_Es1.Enabled = False
'            txt_NumLib_Es1.Enabled = False
'            txt_NumTom_Es1.Text = ""
'            txt_NumFoj_Es1.Text = ""
'            txt_NumLib_Es1.Text = ""
            
         Case 3
            txt_NumPar_Dep2.Enabled = False
            txt_NumAPa_Dep2.Enabled = False
            txt_NumPar_Dep2.Text = ""
            txt_NumAPa_Dep2.Text = ""
            
            txt_NumFic_Dep2.Enabled = False
            txt_NumAFi_Dep2.Enabled = False
            txt_NumFic_Dep2.Text = ""
            txt_NumAFi_Dep2.Text = ""
            
'            txt_NumTom_Es1.Enabled = True
'            txt_NumFoj_Es1.Enabled = True
'            txt_NumLib_Es1.Enabled = True
            
            'Call gs_SetFocus(txt_NumTom_Es1)
      End Select
   End If
End Sub

Private Sub cmb_TipDoc_Dep2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipDoc_Dep2_Click
   End If
End Sub

Private Sub cmb_TipDoc_Es1_Click()

   If cmb_TipDoc_Es1.ListIndex > -1 Then
      Select Case cmb_TipDoc_Es1.ItemData(cmb_TipDoc_Es1.ListIndex)
         Case 1
            txt_NumPar_Es1.Enabled = True
            txt_NumAPa_Es1.Enabled = True
            
            Call gs_SetFocus(txt_NumPar_Es1)
            
            txt_NumFic_Es1.Enabled = False
            txt_NumAFi_Es1.Enabled = False
            txt_NumFic_Es1.Text = ""
            txt_NumAFi_Es1.Text = ""
            
'            txt_NumTom_Es1.Enabled = False
'            txt_NumFoj_Es1.Enabled = False
'            txt_NumLib_Es1.Enabled = False
'            txt_NumTom_Es1.Text = ""
'            txt_NumFoj_Es1.Text = ""
'            txt_NumLib_Es1.Text = ""
            
         Case 2
            txt_NumPar_Es1.Enabled = False
            txt_NumAPa_Es1.Enabled = False
            txt_NumPar_Es1.Text = ""
            txt_NumAPa_Es1.Text = ""
            
            
            txt_NumFic_Es1.Enabled = True
            txt_NumAFi_Es1.Enabled = True
            
            Call gs_SetFocus(txt_NumFic_Es1)
            
'            txt_NumTom_Es1.Enabled = False
'            txt_NumFoj_Es1.Enabled = False
'            txt_NumLib_Es1.Enabled = False
'            txt_NumTom_Es1.Text = ""
'            txt_NumFoj_Es1.Text = ""
'            txt_NumLib_Es1.Text = ""
            
         Case 3
            txt_NumPar_Es1.Enabled = False
            txt_NumAPa_Es1.Enabled = False
            txt_NumPar_Es1.Text = ""
            txt_NumAPa_Es1.Text = ""
            
            txt_NumFic_Es1.Enabled = False
            txt_NumAFi_Es1.Enabled = False
            txt_NumFic_Es1.Text = ""
            txt_NumAFi_Es1.Text = ""
            
'            txt_NumTom_Es1.Enabled = True
'            txt_NumFoj_Es1.Enabled = True
'            txt_NumLib_Es1.Enabled = True
            
            'Call gs_SetFocus(txt_NumTom_Es1)
      End Select
   End If
End Sub

Private Sub cmb_TipDoc_Es1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipDoc_Es1_Click
   End If
End Sub

Private Sub cmb_TipDoc_Es2_Click()

   If cmb_TipDoc_Es2.ListIndex > -1 Then
      Select Case cmb_TipDoc_Es2.ItemData(cmb_TipDoc_Es2.ListIndex)
         Case 1
            txt_NumPar_Es2.Enabled = True
            txt_NumAPa_Es2.Enabled = True
            
            Call gs_SetFocus(txt_NumPar_Es2)
            
            txt_NumFic_Es2.Enabled = False
            txt_NumAFi_Es2.Enabled = False
            txt_NumFic_Es2.Text = ""
            txt_NumAFi_Es2.Text = ""
            
'            txt_NumTom_Es2.Enabled = False
'            txt_NumFoj_Es2.Enabled = False
'            txt_NumLib_Es2.Enabled = False
'            txt_NumTom_Es2.Text = ""
'            txt_NumFoj_Es2.Text = ""
'            txt_NumLib_Es2.Text = ""
            
         Case 2
            txt_NumPar_Es2.Enabled = False
            txt_NumAPa_Es2.Enabled = False
            txt_NumPar_Es2.Text = ""
            txt_NumAPa_Es2.Text = ""
            
            
            txt_NumFic_Es2.Enabled = True
            txt_NumAFi_Es2.Enabled = True
            
            Call gs_SetFocus(txt_NumFic_Es2)
            
'            txt_NumTom_Es2.Enabled = False
'            txt_NumFoj_Es2.Enabled = False
'            txt_NumLib_Es2.Enabled = False
'            txt_NumTom_Es2.Text = ""
'            txt_NumFoj_Es2.Text = ""
'            txt_NumLib_Es2.Text = ""
            
         Case 3
            txt_NumPar_Es2.Enabled = False
            txt_NumAPa_Es2.Enabled = False
            txt_NumPar_Es2.Text = ""
            txt_NumAPa_Es2.Text = ""
            
            txt_NumFic_Es2.Enabled = False
            txt_NumAFi_Es2.Enabled = False
            txt_NumFic_Es2.Text = ""
            txt_NumAFi_Es2.Text = ""
            
'            txt_NumTom_Es2.Enabled = True
'            txt_NumFoj_Es2.Enabled = True
'            txt_NumLib_Es2.Enabled = True
            
            'Call gs_SetFocus(txt_NumTom_Es2)
      End Select
   End If
End Sub

Private Sub cmb_TipDoc_Es2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipDoc_Es2_Click
   End If
End Sub

Private Sub cmb_TipDoc_Inm_Click()

If cmb_TipDoc_Inm.ListIndex > -1 Then
      Select Case cmb_TipDoc_Inm.ItemData(cmb_TipDoc_Inm.ListIndex)
         Case 1
            txt_NumPar_Inm.Enabled = True
            txt_NumAPa_Inm.Enabled = True
            
            Call gs_SetFocus(txt_NumPar_Inm)
            
            txt_NumFic_Inm.Enabled = False
            txt_NumAFi_Inm.Enabled = False
            txt_NumFic_Inm.Text = ""
            txt_NumAFi_Inm.Text = ""
            
'            txt_NumTom_Inm.Enabled = False
'            txt_NumFoj_Inm.Enabled = False
'            txt_NumLib_Inm.Enabled = False
'            txt_NumTom_Inm.Text = ""
'            txt_NumFoj_Inm.Text = ""
'            txt_NumLib_Inm.Text = ""
            
         Case 2
            txt_NumPar_Inm.Enabled = False
            txt_NumAPa_Inm.Enabled = False
            txt_NumPar_Inm.Text = ""
            txt_NumAPa_Inm.Text = ""
            
            
            txt_NumFic_Inm.Enabled = True
            txt_NumAFi_Inm.Enabled = True
            
            Call gs_SetFocus(txt_NumFic_Inm)
            
'            txt_NumTom_Inm.Enabled = False
'            txt_NumFoj_Inm.Enabled = False
'            txt_NumLib_Inm.Enabled = False
'            txt_NumTom_Inm.Text = ""
'            txt_NumFoj_Inm.Text = ""
'            txt_NumLib_Inm.Text = ""
            
         Case 3
            txt_NumPar_Inm.Enabled = False
            txt_NumAPa_Inm.Enabled = False
            txt_NumPar_Inm.Text = ""
            txt_NumAPa_Inm.Text = ""
            
            txt_NumFic_Inm.Enabled = False
            txt_NumAFi_Inm.Enabled = False
            txt_NumFic_Inm.Text = ""
            txt_NumAFi_Inm.Text = ""
            
'            txt_NumTom_Inm.Enabled = True
'            txt_NumFoj_Inm.Enabled = True
'            txt_NumLib_Inm.Enabled = True
'
            'Call gs_SetFocus(txt_NumTom_Inm)
      End Select
   End If
End Sub

Private Sub cmb_TipDoc_Inm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipDoc_Inm_Click
   End If
End Sub

Private Sub Cmb_TipGar_Click()
   If Cmb_TipGar.ListIndex = -1 Then Exit Sub
   If Cmb_TipGar.ItemData(Cmb_TipGar.ListIndex) = 1 Then
      tab_Genera.Tab = 0
      tab_Genera.TabEnabled(0) = True
      tab_Genera.TabEnabled(1) = False
      tab_Genera.TabEnabled(2) = False
      tab_Genera.TabEnabled(3) = False
      tab_Genera.TabEnabled(4) = False
      tab_Genera.TabEnabled(5) = False
      cmb_SedReg.Enabled = False
      cmb_EmpPer.Enabled = False
      cmb_NumInf.Enabled = False
      cmb_FecEva.Enabled = False
      Call gs_SetFocus(ipp_FecEmi)
   Else
      tab_Genera.Tab = 1
      tab_Genera.TabEnabled(0) = False
      tab_Genera.TabEnabled(1) = True
      tab_Genera.TabEnabled(2) = True
      tab_Genera.TabEnabled(3) = True
      tab_Genera.TabEnabled(4) = True
      tab_Genera.TabEnabled(5) = True
      cmb_SedReg.Enabled = True
      cmb_EmpPer.Enabled = True
      cmb_NumInf.Enabled = True
      cmb_FecEva.Enabled = True
      Call gs_SetFocus(cmb_SedReg)
   End If
End Sub

Private Sub cmb_TipGar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      'Call gs_SetFocus(txt_NumOpe)
      Call gs_SetFocus(ipp_FecEmi)
   End If
End Sub

Private Sub cmd_Agrega_Click()
   'moddat_g_int_FlgGrb_2 = 1       'Insertar
   moddat_g_int_FlgAct_2 = 1
   Call fs_Activa(True)
   Call gs_SetFocus(Cmb_TipGar)
   cmd_Editar.Enabled = False
   cmd_Borrar.Enabled = False
End Sub

Private Sub cmd_Borrar_Click()
   
   If fs_Validar_MovGar = True Then
      MsgBox "Verifique que el registro no tenga movimientos en el módulo de Gestión.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(grd_Listad)
      Exit Sub
   End If
   
   If MsgBox("Recuerde eliminar el Asiento Contable que se registró, ¿Está seguro de eliminar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      Screen.MousePointer = 11
   
      'Grabando Información de Carta Fianza
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & " USP_TPR_MAEGAR_ELIMINA ("
      g_str_Parame = g_str_Parame & CStr(grd_Listad.TextMatrix(grd_Listad.Row, 5)) & ", "
      g_str_Parame = g_str_Parame & CStr(moddat_g_int_TipDoc) & ", "
      g_str_Parame = g_str_Parame & "'" & CStr(moddat_g_str_NumDoc) & "') "
                        
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If
      
      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar la eliminación de los datos. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
      
      Screen.MousePointer = 0
   Loop
   
   'Actualiza la Grilla
   Call fs_Buscar
   Call fs_Activa(False)
   Call fs_ConsultaNumRef(moddat_g_int_TipDoc, moddat_g_str_NumDoc)
   Call frm_Ges_TecPro_01.fs_Buscar
End Sub
Private Function fs_Validar_MovGar() As Boolean
   fs_Validar_MovGar = False

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT COUNT(*) AS CONTADOR "
   g_str_Parame = g_str_Parame & "    FROM TPR_MAERDE "
   g_str_Parame = g_str_Parame & "   WHERE MAERDE_NUMREF = '" & CStr(Trim(Replace(pnl_NumRef.Caption, "-", ""))) & "' "
   g_str_Parame = g_str_Parame & "     AND MAERDE_CODIGO = 6 "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
     Exit Function
   End If
   
   If g_rst_GenAux.BOF And g_rst_GenAux.EOF Then
      g_rst_GenAux.Close
      Set g_rst_GenAux = Nothing
      Exit Function
   End If
     
   If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
      If (g_rst_GenAux!CONTADOR) > 0 Then
         fs_Validar_MovGar = True
      End If
   End If
End Function
Private Sub cmd_Cancel_Click()
   Call fs_Activa(False)
   Call fs_Limpia
   tab_Genera.Tab = 1
End Sub

Private Sub cmd_CfiAso_Click()
   cmb_NumRef.ListIndex = -1
   cmb_NumRef.Enabled = False
   If moddat_g_int_FlgAct_2 = 1 Then 'moddat_g_int_FlgGrb_2 = 1
      pnl_NumCFi.Caption = ""
   End If
   frm_Ges_TecPro_14.Show 1
End Sub

Private Sub cmd_Editar_Click()
'   'moddat_g_int_FlgGrb_2 = 2       'Actualiza
   moddat_g_int_FlgAct_2 = 2
   Call fs_Activa(True)
   
   If grd_Listad.Row = -1 Then Exit Sub
   
   Cmb_TipGar.Text = grd_Listad.TextMatrix(grd_Listad.Row, 0)
   l_int_NumIte = grd_Listad.TextMatrix(grd_Listad.Row, 5)
    
   If Cmb_TipGar.ItemData(Cmb_TipGar.ListIndex) = 1 Then
      ipp_FecEmi.Text = Format(CStr(grd_Listad.TextMatrix(grd_Listad.Row, 1)), "dd/mm/yyyy")
      cmb_Moneda.Text = grd_Listad.TextMatrix(grd_Listad.Row, 2)
      ipp_ImpGar.Value = Format(grd_Listad.TextMatrix(grd_Listad.Row, 3), "###,###,###,##0.00")
      pnl_NumCFi.Caption = grd_Listad.TextMatrix(grd_Listad.Row, 4)
   Else
      
      
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & " SELECT MAEGAR_NUMITE    , MAEGAR_TIPDOC    , MAEGAR_NUMDOC    , MAEGAR_CODTAS    , MAEGAR_TIPGAR    , MAEGAR_SEDREG    , TRIM(MAEGAR_NUMREF) NUMREF,  "
      g_str_Parame = g_str_Parame & "        MAEGAR_FECPRE_INM, MAEGAR_NUMPRE_INM, MAEGAR_FECINS_INM, MAEGAR_TDOREG_INM, MAEGAR_PARFIC_INM, MAEGAR_NUMASI_INM, MAEGAR_TIPMON_INM, NVL(MAEGAR_MTOGAR_INM,0) MTOGAR_INM, "
      g_str_Parame = g_str_Parame & "        MAEGAR_FLGEST_ES1, MAEGAR_FECPRE_ES1, MAEGAR_NUMPRE_ES1, MAEGAR_FECINS_ES1, MAEGAR_TDOREG_ES1, MAEGAR_PARFIC_ES1, MAEGAR_NUMASI_ES1, MAEGAR_TIPMON_ES1, NVL(MAEGAR_MTOGAR_ES1,0) MTOGAR_ES1 , "
      g_str_Parame = g_str_Parame & "        MAEGAR_FLGEST_ES2, MAEGAR_FECPRE_ES2, MAEGAR_NUMPRE_ES2, MAEGAR_FECINS_ES2, MAEGAR_TDOREG_ES2, MAEGAR_PARFIC_ES2, MAEGAR_NUMASI_ES2, MAEGAR_TIPMON_ES2, NVL(MAEGAR_MTOGAR_ES2,0) MTOGAR_ES2 , "
      g_str_Parame = g_str_Parame & "        MAEGAR_FLGEST_DE1, MAEGAR_FECPRE_DE1, MAEGAR_NUMPRE_DE1, MAEGAR_FECINS_DE1, MAEGAR_TDOREG_DE1, MAEGAR_PARFIC_DE1, MAEGAR_NUMASI_DE1, MAEGAR_TIPMON_DE1, NVL(MAEGAR_MTOGAR_DE1,0) MTOGAR_DE1 , "
      g_str_Parame = g_str_Parame & "        MAEGAR_FLGEST_DE2, MAEGAR_FECPRE_DE2, MAEGAR_NUMPRE_DE2, MAEGAR_FECINS_DE2, MAEGAR_TDOREG_DE2, MAEGAR_PARFIC_DE2, MAEGAR_NUMASI_DE2, MAEGAR_TIPMON_DE2, NVL(MAEGAR_MTOGAR_DE2,0) MTOGAR_DE2 , "
      g_str_Parame = g_str_Parame & "        MAEGAR_SITUAC    , MAEGAR_OBSERV    , MAEGAR_NROCNT    , MAEGAR_FECLIB    , EVATAS_CODEMP    , EVATAS_NUMINF    , EVATAS_FECEVA "
      g_str_Parame = g_str_Parame & "   FROM TPR_MAEGAR "
      g_str_Parame = g_str_Parame & "        LEFT JOIN TPR_EVATAS ON EVATAS_TIPDOC = MAEGAR_TIPDOC AND EVATAS_NUMDOC = MAEGAR_NUMDOC AND EVATAS_CODTAS = MAEGAR_CODTAS "
      g_str_Parame = g_str_Parame & "  WHERE MAEGAR_NUMITE = " & CStr(l_int_NumIte) & " "
      g_str_Parame = g_str_Parame & "    AND MAEGAR_TIPDOC = " & CStr(moddat_g_int_TipDoc) & " "
      g_str_Parame = g_str_Parame & "    AND MAEGAR_NUMDOC = '" & CStr(moddat_g_str_NumDoc) & "' "
      g_str_Parame = g_str_Parame & "    AND MAEGAR_SITUAC = 1 "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
         Exit Sub
      End If
           
      If g_rst_Genera.BOF And g_rst_Genera.EOF Then
         g_rst_Genera.Close
         Set g_rst_Genera = Nothing
         Exit Sub
      End If
      
      If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
         g_rst_Genera.MoveFirst
         
         Do While Not g_rst_Genera.EOF
            If Not IsNull(g_rst_Genera!MAEGAR_SEDREG) Then
               cmb_SedReg.Text = moddat_gf_Consulta_ParDes("511", CStr(g_rst_Genera!MAEGAR_SEDREG))
            End If
            
            If Not IsNull(g_rst_Genera!EVATAS_CODEMP) Then
               cmb_EmpPer.Text = moddat_gf_Consulta_ParDes("507", CStr(g_rst_Genera!EVATAS_CODEMP))
               Call gs_BuscarCombo_Item(cmb_EmpPer, CStr(g_rst_Genera!EVATAS_CODEMP))
            End If
            
            If Not IsNull(g_rst_Genera!EVATAS_NUMINF) Then
               Call gs_BuscarCombo(cmb_NumInf, CStr(Trim(g_rst_Genera!EVATAS_NUMINF)))
            End If
            
            If Not IsNull(g_rst_Genera!EVATAS_FECEVA) Then
               cmb_FecEva.Text = Format(gf_FormatoFecha(CStr(g_rst_Genera!EVATAS_FECEVA)), "dd/mm/yyyy")
            End If
            
            If Not IsNull(g_rst_Genera!MAEGAR_FECPRE_INM) Then
               ipp_FecPre_Inm.Text = Format(gf_FormatoFecha(CStr(g_rst_Genera!MAEGAR_FECPRE_INM)), "dd/mm/yyyy")
            End If
            
            If Not IsNull(g_rst_Genera!MAEGAR_NUMPRE_INM) Then
               txt_NumPre_Inm.Text = Trim(g_rst_Genera!MAEGAR_NUMPRE_INM)
            End If
            
            If Not IsNull(g_rst_Genera!MAEGAR_FECINS_INM) Then
               ipp_FecIns_Inm.Text = Format(gf_FormatoFecha(CStr(g_rst_Genera!MAEGAR_FECINS_INM)), "dd/mm/yyyy")
            End If
            
            If Not IsNull(g_rst_Genera!MAEGAR_TDOREG_INM) Then
               Call gs_BuscarCombo_Item(cmb_TipDoc_Inm, CStr(g_rst_Genera!MAEGAR_TDOREG_INM))
            
               If cmb_TipDoc_Inm.ItemData(cmb_TipDoc_Inm.ListIndex) = 1 Then
                  If Not IsNull(Trim(g_rst_Genera!MAEGAR_PARFIC_INM)) Then txt_NumPar_Inm.Text = Trim(g_rst_Genera!MAEGAR_PARFIC_INM)
                  If Not IsNull(Trim(g_rst_Genera!MAEGAR_NUMASI_INM)) Then txt_NumAPa_Inm.Text = Trim(g_rst_Genera!MAEGAR_NUMASI_INM)
               Else
                  If Not IsNull(Trim(g_rst_Genera!MAEGAR_PARFIC_INM)) Then txt_NumFic_Inm.Text = Trim(g_rst_Genera!MAEGAR_PARFIC_INM)
                  If Not IsNull(Trim(g_rst_Genera!MAEGAR_NUMASI_INM)) Then txt_NumAFi_Inm.Text = Trim(g_rst_Genera!MAEGAR_NUMASI_INM)
               End If
            End If
            ipp_MtoHip_Inm.Text = Format(g_rst_Genera!MTOGAR_INM, "###,###,###,##0.00")
            Call gs_BuscarCombo_Item(cmb_Moneda_Inm, CStr(g_rst_Genera!MAEGAR_TIPMON_INM))
            
            If g_rst_Genera!MAEGAR_FLGEST_ES1 = 1 Then
               cmb_FlgEst_Es1.ListIndex = 0
               If Not IsNull(g_rst_Genera!MAEGAR_FECPRE_ES1) Then
                  ipp_FecPre_Es1.Text = Format(gf_FormatoFecha(CStr(g_rst_Genera!MAEGAR_FECPRE_ES1)), "dd/mm/yyyy")
               End If
               If Not IsNull(g_rst_Genera!MAEGAR_NUMPRE_ES1) Then
                  txt_NumPre_Es1.Text = Trim(g_rst_Genera!MAEGAR_NUMPRE_ES1)
               End If
               If Not IsNull(g_rst_Genera!MAEGAR_FECINS_ES1) Then
                  ipp_FecIns_Es1.Text = Format(gf_FormatoFecha(CStr(g_rst_Genera!MAEGAR_FECINS_ES1)), "dd/mm/yyyy")
               End If
               Call gs_BuscarCombo_Item(cmb_TipDoc_Es1, CStr(g_rst_Genera!MAEGAR_TDOREG_ES1))
               
               If cmb_TipDoc_Es1.ItemData(cmb_TipDoc_Es1.ListIndex) = 1 Then
                  If Not IsNull(g_rst_Genera!MAEGAR_PARFIC_ES1) Then
                     txt_NumPar_Es1.Text = Trim(g_rst_Genera!MAEGAR_PARFIC_ES1)
                  End If
                  If Not IsNull(g_rst_Genera!MAEGAR_NUMASI_ES1) Then
                     txt_NumAPa_Es1.Text = Trim(g_rst_Genera!MAEGAR_NUMASI_ES1)
                  End If
               Else
                  If Not IsNull(g_rst_Genera!MAEGAR_PARFIC_ES1) Then txt_NumFic_Es1.Text = Trim(g_rst_Genera!MAEGAR_PARFIC_ES1)
                  If Not IsNull(g_rst_Genera!MAEGAR_NUMASI_ES1) Then txt_NumAFi_Es1.Text = Trim(g_rst_Genera!MAEGAR_NUMASI_ES1)
               End If
               
               ipp_MtoHip_Es1.Text = Format(g_rst_Genera!MTOGAR_ES1, "###,###,###,##0.00")
               Call gs_BuscarCombo_Item(cmb_Moneda_Es1, CStr(g_rst_Genera!MAEGAR_TIPMON_ES1))
            Else
               cmb_FlgEst_Es1.ListIndex = 1
            End If
            
            If g_rst_Genera!MAEGAR_FLGEST_ES2 = 1 Then
               cmb_FlgEst_Es2.ListIndex = 0
               ipp_FecPre_Es2.Text = Format(gf_FormatoFecha(CStr(g_rst_Genera!MAEGAR_FECPRE_ES2)), "dd/mm/yyyy")
               txt_NumPre_Es2.Text = Trim(g_rst_Genera!MAEGAR_NUMPRE_ES2)
               ipp_FecIns_Es2.Text = Format(gf_FormatoFecha(CStr(g_rst_Genera!MAEGAR_FECINS_ES2)), "dd/mm/yyyy")
               
               Call gs_BuscarCombo_Item(cmb_TipDoc_Es2, CStr(g_rst_Genera!MAEGAR_TDOREG_ES2))
               
               If cmb_TipDoc_Es2.ItemData(cmb_TipDoc_Es2.ListIndex) = 1 Then
                  txt_NumPar_Es2.Text = Trim(g_rst_Genera!MAEGAR_PARFIC_ES2)
                  txt_NumAPa_Es2.Text = Trim(g_rst_Genera!MAEGAR_NUMASI_ES2)
               Else
                  txt_NumFic_Es2.Text = Trim(g_rst_Genera!MAEGAR_PARFIC_ES2)
                  txt_NumAFi_Es2.Text = Trim(g_rst_Genera!MAEGAR_NUMASI_ES2)
               End If
               
               ipp_MtoHip_Es2.Text = Format(g_rst_Genera!MTOGAR_ES2, "###,###,###,##0.00")
               Call gs_BuscarCombo_Item(cmb_Moneda_Es2, CStr(g_rst_Genera!MAEGAR_TIPMON_ES2))
            Else
               cmb_FlgEst_Es2.ListIndex = 1
            End If
            
            If g_rst_Genera!MAEGAR_FLGEST_DE1 = 1 Then
               cmb_FlgEst_Dep1.ListIndex = 0
               If Not IsNull(g_rst_Genera!MAEGAR_FECPRE_DE1) Then
                  ipp_FecPre_Dep1.Text = Format(gf_FormatoFecha(CStr(g_rst_Genera!MAEGAR_FECPRE_DE1)), "dd/mm/yyyy")
               End If
               If Not IsNull(g_rst_Genera!MAEGAR_NUMPRE_DE1) Then
                  txt_NumPre_Dep1.Text = Trim(g_rst_Genera!MAEGAR_NUMPRE_DE1)
               End If
               If Not IsNull(g_rst_Genera!MAEGAR_FECINS_DE1) Then
                  ipp_FecIns_Dep1.Text = Format(gf_FormatoFecha(CStr(g_rst_Genera!MAEGAR_FECINS_DE1)), "dd/mm/yyyy")
               End If
               
               Call gs_BuscarCombo_Item(cmb_TipDoc_Dep1, CStr(g_rst_Genera!MAEGAR_TDOREG_DE1))
               
               If cmb_TipDoc_Dep1.ItemData(cmb_TipDoc_Dep1.ListIndex) = 1 Then
                  If Not IsNull(g_rst_Genera!MAEGAR_PARFIC_DE1) Then
                     txt_NumPar_Dep1.Text = Trim(g_rst_Genera!MAEGAR_PARFIC_DE1)
                  End If
                  If Not IsNull(g_rst_Genera!MAEGAR_NUMASI_DE1) Then
                     txt_NumAPa_Dep1.Text = Trim(g_rst_Genera!MAEGAR_NUMASI_DE1)
                  End If
               Else
                  txt_NumFic_Dep1.Text = Trim(g_rst_Genera!MAEGAR_PARFIC_DE1)
                  txt_NumAFi_Dep1.Text = Trim(g_rst_Genera!MAEGAR_NUMASI_DE1)
               End If
               
               ipp_MtoHip_Dep1.Text = Format(g_rst_Genera!MTOGAR_DE1, "###,###,###,##0.00")
               Call gs_BuscarCombo_Item(cmb_Moneda_Dep1, CStr(g_rst_Genera!MAEGAR_TIPMON_DE1))
            Else
               cmb_FlgEst_Dep1.ListIndex = 1
            End If
            
            If g_rst_Genera!MAEGAR_FLGEST_DE2 = 1 Then
               cmb_FlgEst_Dep2.ListIndex = 0
               ipp_FecPre_Dep2.Text = Format(gf_FormatoFecha(CStr(g_rst_Genera!MAEGAR_FECPRE_DE2)), "dd/mm/yyyy")
               txt_NumPre_Dep2.Text = Trim(g_rst_Genera!MAEGAR_NUMPRE_DE2)
               ipp_FecIns_Dep2.Text = Format(gf_FormatoFecha(CStr(g_rst_Genera!MAEGAR_FECINS_DE2)), "dd/mm/yyyy")
               
               Call gs_BuscarCombo_Item(cmb_TipDoc_Dep2, CStr(g_rst_Genera!MAEGAR_TDOREG_DE2))
               
               If cmb_TipDoc_Dep2.ItemData(cmb_TipDoc_Dep2.ListIndex) = 1 Then
                  txt_NumPar_Dep2.Text = Trim(g_rst_Genera!MAEGAR_PARFIC_DE2)
                  txt_NumAPa_Dep2.Text = Trim(g_rst_Genera!MAEGAR_NUMASI_DE2)
               Else
                  txt_NumFic_Dep2.Text = Trim(g_rst_Genera!MAEGAR_PARFIC_DE2)
                  txt_NumAFi_Dep2.Text = Trim(g_rst_Genera!MAEGAR_NUMASI_DE2)
               End If
               
               ipp_MtoHip_Dep2.Text = Format(g_rst_Genera!MTOGAR_DE2, "###,###,###,##0.00")
               Call gs_BuscarCombo_Item(cmb_Moneda_Dep2, CStr(g_rst_Genera!MAEGAR_TIPMON_DE2))
            Else
               cmb_FlgEst_Dep2.ListIndex = 1
            End If
            
            g_rst_Genera.MoveNext
         Loop
      End If
   End If
End Sub

Private Sub cmd_ExpExc_Click()
    'Confirmacion
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   Screen.MousePointer = 11
   Call fs_GenExc
   Screen.MousePointer = 0
End Sub
Private Sub fs_GenExc()
Dim r_obj_Excel         As Excel.Application
Dim r_str_ParAux        As String
Dim r_str_FecRpt        As String
Dim r_int_Contad        As Integer
Dim r_int_NroFil        As Integer
Dim r_int_NoFlLi        As Integer
   
    r_int_NroFil = 8
    
    Set r_obj_Excel = New Excel.Application
    r_obj_Excel.SheetsInNewWorkbook = 1
    r_obj_Excel.Workbooks.Add

    With r_obj_Excel.ActiveSheet
        .Cells(2, 2) = "REPORTE DE GARANTIAS"
        .Range(.Cells(2, 2), .Cells(2, 6)).Merge
        .Range(.Cells(2, 2), .Cells(2, 6)).Font.Bold = True
        .Range(.Cells(2, 2), .Cells(2, 6)).HorizontalAlignment = xlHAlignCenter
        .Range(.Cells(2, 2), .Cells(2, 6)).Font.Size = 14

        .Cells(4, 2) = "TIPO DE DOCUMENTO"
        .Cells(4, 3) = Trim(pnl_TipDoc.Caption)
        .Cells(5, 2) = "NRO. DOCUMENTO"
        .Cells(5, 3) = "'" & Trim(pnl_NroDoc.Caption)
        .Cells(6, 2) = "RAZÓN SOCIAL"
        .Cells(6, 3) = Trim(pnl_RazSoc.Caption)
        .Range(.Cells(3, 2), .Cells(6, 2)).Font.Bold = True
        
        .Cells(r_int_NroFil, 2) = "TIPO GARANTIA"
        .Range(.Cells(r_int_NroFil, 2), .Cells(r_int_NroFil + 1, 2)).Merge
        .Cells(r_int_NroFil, 3) = "FECHA EMISION"
        .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil + 1, 3)).Merge
        .Cells(r_int_NroFil, 4) = "MONEDA"
        .Range(.Cells(r_int_NroFil, 4), .Cells(r_int_NroFil + 1, 4)).Merge
        .Cells(r_int_NroFil, 5) = "MONTO GARANTIA"
        .Range(.Cells(r_int_NroFil, 5), .Cells(r_int_NroFil + 1, 5)).Merge
        .Cells(r_int_NroFil, 6) = "CARTAS FIANZAS ASOCIADAS"
        .Range(.Cells(r_int_NroFil, 6), .Cells(r_int_NroFil + 1, 6)).Merge
        .Cells(r_int_NroFil, 7) = "MONTO CARTA FIANZA"
        .Range(.Cells(r_int_NroFil, 7), .Cells(r_int_NroFil + 1, 7)).Merge
        
        .Range(.Cells(r_int_NroFil, 2), .Cells(r_int_NroFil + 1, 7)).Interior.Color = RGB(146, 208, 80)
        .Range(.Cells(r_int_NroFil, 2), .Cells(r_int_NroFil + 1, 7)).Font.Bold = True
        .Range(.Cells(r_int_NroFil, 2), .Cells(r_int_NroFil + 1, 7)).HorizontalAlignment = xlHAlignCenter
        
        .Columns("A").ColumnWidth = 1
        .Columns("B").ColumnWidth = 25.5
        .Columns("C").ColumnWidth = 15
        .Columns("D").ColumnWidth = 15
        .Columns("D").HorizontalAlignment = xlHAlignCenter
        .Columns("E").ColumnWidth = 15
        .Columns("E").NumberFormat = "###,###,###,##0.00"
        .Columns("E").HorizontalAlignment = xlHAlignRight
        .Columns("F").ColumnWidth = 25.5
        .Columns("F").HorizontalAlignment = xlHAlignCenter
        .Columns("G").ColumnWidth = 15
        .Columns("G").NumberFormat = "###,###,###,##0.00"
        .Columns("G").HorizontalAlignment = xlHAlignRight
        
        With .Range(.Cells(8, 2), .Cells(9, 7))
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
        End With
        
        .Range(.Cells(1, 1), .Cells(99, 99)).Font.Name = "Calibri"
        .Range(.Cells(3, 1), .Cells(99, 99)).Font.Size = 11
         
        r_int_NroFil = r_int_NroFil + 2
         
        For r_int_NoFlLi = 0 To grd_Listad.Rows - 1

            .Cells(r_int_NroFil, 2) = "'" & grd_Listad.TextMatrix(r_int_NoFlLi, 0)
            .Cells(r_int_NroFil, 3) = grd_Listad.TextMatrix(r_int_NoFlLi, 1)
            .Cells(r_int_NroFil, 4) = grd_Listad.TextMatrix(r_int_NoFlLi, 2)
            .Cells(r_int_NroFil, 5) = grd_Listad.TextMatrix(r_int_NoFlLi, 3)
            .Cells(r_int_NroFil, 6) = "'" & grd_Listad.TextMatrix(r_int_NoFlLi, 4)
            .Cells(r_int_NroFil, 7) = grd_Listad.TextMatrix(r_int_NoFlLi, 6)
            
            r_int_NroFil = r_int_NroFil + 1
        Next r_int_NoFlLi
        
        With .Range(.Cells(10, 2), .Cells(r_int_NroFil, 3))
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
        End With
    
        With .Range(.Cells(8, 2), .Cells(r_int_NroFil - 1, 7))
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlInsideVertical).LineStyle = xlContinuous
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        End With
   End With
   
   r_obj_Excel.Visible = True
End Sub
Private Sub cmd_Grabar_Click()
Dim r_int_NumIte     As Integer
Dim r_str_MsjGrb     As String
Dim r_int_CodTas     As Integer

   r_int_NumIte = 0
   
   If Cmb_TipGar.ListIndex = -1 Then
      MsgBox "Debe seleccionar Tipo de Garantía.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(Cmb_TipGar)
      Exit Sub
   End If
      
   If Cmb_TipGar.ItemData(Cmb_TipGar.ListIndex) = 1 Then
   
      If CDate(ipp_FecEmi.Text) > date Then
         MsgBox "Debe ingresar una Fecha de Emisión válida.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_FecEmi)
         Exit Sub
      End If
      
      If cmb_Moneda.ListIndex = -1 Then
         MsgBox "Debe seleccionar Moneda.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_Moneda)
         Exit Sub
      End If
    
      If ipp_ImpGar.Value = 0 Then
         MsgBox "Debe ingresar Importe de Garantía.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_ImpGar)
         Exit Sub
      End If
      
      'Valida que el Monto Total de garantías sea menor e igual al Monto de Carta Fianza
   '   If fs_Validar_MtoGar = False Then
   '      MsgBox "El monto ingresado excede al Total de la Línea Asignada.", vbExclamation, modgen_g_str_NomPlt
   '      Call gs_SetFocus(ipp_ImpGar)
   '      Exit Sub
   '   End If
      
      'Valida que el Monto de Garantía no sea mayor al monto ya pagado
   '   If fs_Validar_MtoGarPag = False Then
   '      MsgBox "El Importe ingresado es menor al Monto Pagado de Garantía.", vbExclamation, modgen_g_str_NomPlt
   '      Call gs_SetFocus(ipp_ImpGar)
   '      Exit Sub
   '   End If
      
'      Valida que el Monto de Cartas Fianza no sea menor ni mayor al monto de Garantía
      If pnl_NumCFi.Caption <> "" Then
         If fs_Buscar_Cartas_NoCliente = True Then
            If fs_Validar_MtoCfiGar = False Then
               MsgBox "El Importe CSO de No Cliente, no es igual al Monto de Garantía.", vbExclamation, modgen_g_str_NomPlt
               'MsgBox "El Importe de Cartas Fianza es mayor al Monto de Garantía.", vbExclamation, modgen_g_str_NomPlt
               Call gs_SetFocus(ipp_ImpGar)
               Exit Sub
            End If
         End If
      End If
      
      If CStr(pnl_NumCFi.Caption) = "" Then
         r_str_MsjGrb = "¿Está seguro de grabar los datos sin asociar ninguna Carta Fianza?"
      Else
         r_str_MsjGrb = "¿Está seguro de grabar los datos?"
      End If
      
   Else
      
      If cmb_SedReg.ListIndex = -1 Then
         MsgBox "Debe seleccionar la Sede Registral.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_SedReg)
         Exit Sub
      End If
      '--Datos de Tasación
      If cmb_EmpPer.ListIndex = -1 Then
         MsgBox "Debe seleccionar Empresa de Peritaje.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_EmpPer)
         Exit Sub
      End If
      If cmb_NumInf.ListIndex = -1 Then
         MsgBox "Debe seleccionar Número de Informe.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_NumInf)
         Exit Sub
      End If
      If cmb_FecEva.ListIndex = -1 Then
         MsgBox "Debe seleccionar Fecha de Tasación.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_FecEva)
         Exit Sub
      End If
      '--
      If CDate(ipp_FecPre_Inm.Text) > date Then
         MsgBox "La Fecha de Presentación no puede ser mayor a la fecha actual.", vbExclamation, modgen_g_str_NomPlt
         tab_Genera.Tab = 1
         Call gs_SetFocus(ipp_FecPre_Inm)
         Exit Sub
      End If
'      If Len(Trim(txt_NumPre_Inm.Text)) = 0 Then
'         MsgBox "Debe ingresar el Número de Presentación de la Hipoteca para Inmueble.", vbExclamation, modgen_g_str_NomPlt
'         tab_Genera.Tab = 1
'         Call gs_SetFocus(txt_NumPre_Inm)
'         Exit Sub
'      End If
'      If CDate(ipp_FecIns_Inm.Text) > date Then
'         MsgBox "La Fecha de Inscripción no puede ser mayor a la fecha actual.", vbExclamation, modgen_g_str_NomPlt
'         tab_Genera.Tab = 1
'         Call gs_SetFocus(ipp_FecIns_Inm)
'         Exit Sub
'      End If
      If CDate(ipp_FecIns_Inm.Text) < CDate(ipp_FecPre_Inm.Text) Then
         MsgBox "La Fecha de Inscripción no puede ser menor a la Fecha de Presentación.", vbExclamation, modgen_g_str_NomPlt
         tab_Genera.Tab = 1
         Call gs_SetFocus(ipp_FecIns_Inm)
         Exit Sub
      End If
'      If cmb_TipDoc_Inm.ListIndex = -1 Then
'         MsgBox "Debe seleccionar el Tipo de Documento Registral para Inmueble.", vbExclamation, modgen_g_str_NomPlt
'         tab_Genera.Tab = 1
'         Call gs_SetFocus(cmb_TipDoc_Inm)
'         Exit Sub
'      End If
'
'      Select Case cmb_TipDoc_Inm.ItemData(cmb_TipDoc_Inm.ListIndex)
'         Case 1
'            If Len(Trim(txt_NumPar_Inm.Text)) = 0 Then
'               MsgBox "Debe ingresar el Nro. de Partida Electrónica para Inmueble.", vbExclamation, modgen_g_str_NomPlt
'               tab_Genera.Tab = 1
'               Call gs_SetFocus(txt_NumPar_Inm)
'               Exit Sub
'            End If
'            If Len(Trim(txt_NumAPa_Inm.Text)) = 0 Then
'               MsgBox "Debe ingresar el Nro. de Asiento en la Partida Electrónica para Inmueble.", vbExclamation, modgen_g_str_NomPlt
'               tab_Genera.Tab = 1
'               Call gs_SetFocus(txt_NumAPa_Inm)
'               Exit Sub
'            End If
'
'         Case 2
'            If Len(Trim(txt_NumFic_Inm.Text)) = 0 Then
'               MsgBox "Debe ingresar el Nro. de Ficha Registral para Inmueble.", vbExclamation, modgen_g_str_NomPlt
'               tab_Genera.Tab = 1
'               Call gs_SetFocus(txt_NumFic_Inm)
'               Exit Sub
'            End If
'            If Len(Trim(txt_NumAFi_Inm.Text)) = 0 Then
'               MsgBox "Debe ingresar el Nro. de Asiento en la Ficha Registral para Inmueble.", vbExclamation, modgen_g_str_NomPlt
'               tab_Genera.Tab = 1
'               Call gs_SetFocus(txt_NumAFi_Inm)
'               Exit Sub
'            End If
'
''         Case 3
''            If Len(Trim(txt_NumTom_Inm.Text)) = 0 Then
''               MsgBox "Debe ingresar el Nro. de Tomo.", vbExclamation, modgen_g_str_NomPlt
''               tab_Genera.Tab = 0
''               Call gs_SetFocus(txt_NumTom_Inm)
''               Exit Sub
''            End If
''            If Len(Trim(txt_NumFoj_Inm.Text)) = 0 Then
''               MsgBox "Debe ingresar el Nro. de Foja.", vbExclamation, modgen_g_str_NomPlt
''               tab_Genera.Tab = 0
''               Call gs_SetFocus(txt_NumFoj_Inm)
''               Exit Sub
''            End If
''            If Len(Trim(txt_NumLib_Inm.Text)) = 0 Then
''               MsgBox "Debe ingresar el Nro. de Libro.", vbExclamation, modgen_g_str_NomPlt
''               tab_Genera.Tab = 0
''               Call gs_SetFocus(txt_NumLib_Inm)
''               Exit Sub
''            End If
'      End Select
   
      If ipp_MtoHip_Inm.Value = 0 Then
         tab_Genera.Tab = 1
         MsgBox "Debe ingresar el Monto Hipotecado para Inmueble.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_MtoHip_Inm)
         Exit Sub
      End If
      
'      If ipp_ValRea_Inm.Value = 0 Then
'         tab_Genera.Tab = 1
'         MsgBox "Debe ingresar el Valor de Realización del Inmueble.", vbExclamation, modgen_g_str_NomPlt
'         Call gs_SetFocus(ipp_ValRea_Inm)
'         Exit Sub
'      End If
      
      If cmb_Moneda_Inm.ListIndex = -1 Then
         tab_Genera.Tab = 1
         MsgBox "Debe ingresar Moneda para Inmueble.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_Moneda_Inm)
         Exit Sub
      End If
      
      If cmb_FlgEst_Es1.ListIndex = -1 Then
         MsgBox "Debe seleccionar si hay Hipoteca para Estacionamiento 1.", vbExclamation, modgen_g_str_NomPlt
         tab_Genera.Tab = 2
         Call gs_SetFocus(cmb_FlgEst_Es1)
         Exit Sub
      End If
      
      If cmb_FlgEst_Es1.ItemData(cmb_FlgEst_Es1.ListIndex) = 1 Then
         If CDate(ipp_FecPre_Es1.Text) > date Then
            MsgBox "La Fecha de Presentación no puede ser mayor a la fecha actual.", vbExclamation, modgen_g_str_NomPlt
            tab_Genera.Tab = 2
            Call gs_SetFocus(ipp_FecPre_Es1)
            Exit Sub
         End If
'         If Len(Trim(txt_NumPre_Es1.Text)) = 0 Then
'            MsgBox "Debe ingresar el Número de Presentación para Estacionamiento 1.", vbExclamation, modgen_g_str_NomPlt
'            tab_Genera.Tab = 2
'            Call gs_SetFocus(txt_NumPre_Es1)
'            Exit Sub
'         End If
         If CDate(ipp_FecIns_Es1.Text) > date Then
            MsgBox "La Fecha de Inscripción no puede ser mayor a la fecha actual.", vbExclamation, modgen_g_str_NomPlt
            tab_Genera.Tab = 2
            Call gs_SetFocus(ipp_FecIns_Es1)
            Exit Sub
         End If
         If CDate(ipp_FecIns_Es1.Text) < CDate(ipp_FecPre_Es1.Text) Then
            MsgBox "La Fecha de Inscripción no puede ser menor a la Fecha de Presentación.", vbExclamation, modgen_g_str_NomPlt
            tab_Genera.Tab = 2
            Call gs_SetFocus(ipp_FecIns_Es1)
            Exit Sub
         End If
'         If cmb_TipDoc_Es1.ListIndex = -1 Then
'            MsgBox "Debe seleccionar el Tipo de Documento Registral para Estacionamiento 1.", vbExclamation, modgen_g_str_NomPlt
'            tab_Genera.Tab = 2
'            Call gs_SetFocus(cmb_TipDoc_Es1)
'            Exit Sub
'         End If
'
'         Select Case cmb_TipDoc_Es1.ItemData(cmb_TipDoc_Es1.ListIndex)
'            Case 1
'               If Len(Trim(txt_NumPar_Es1.Text)) = 0 Then
'                  MsgBox "Debe ingresar el Nro. de Partida Electrónica para Estacionamiento 1.", vbExclamation, modgen_g_str_NomPlt
'                  tab_Genera.Tab = 2
'                  Call gs_SetFocus(txt_NumPar_Es1)
'                  Exit Sub
'               End If
'               If Len(Trim(txt_NumAPa_Es1.Text)) = 0 Then
'                  MsgBox "Debe ingresar el Nro. de Asiento en la Partida Electrónica para Estacionamiento 1.", vbExclamation, modgen_g_str_NomPlt
'                  tab_Genera.Tab = 2
'                  Call gs_SetFocus(txt_NumAPa_Es1)
'                  Exit Sub
'               End If
'
'            Case 2
'               If Len(Trim(txt_NumFic_Es1.Text)) = 0 Then
'                  MsgBox "Debe ingresar el Nro. de Ficha Registral para Estacionamiento 1.", vbExclamation, modgen_g_str_NomPlt
'                  tab_Genera.Tab = 2
'                  Call gs_SetFocus(txt_NumFic_Es1)
'                  Exit Sub
'               End If
'               If Len(Trim(txt_NumAFi_Es1.Text)) = 0 Then
'                  MsgBox "Debe ingresar el Nro. de Asiento en la Ficha Registral para Estacionamiento 1.", vbExclamation, modgen_g_str_NomPlt
'                  tab_Genera.Tab = 2
'                  Call gs_SetFocus(txt_NumAFi_Es1)
'                  Exit Sub
'               End If
               
'            Case 3
'               If Len(Trim(txt_NumTom_Es1.Text)) = 0 Then
'                  MsgBox "Debe ingresar el Nro. de Tomo.", vbExclamation, modgen_g_str_NomPlt
'                  tab_Genera.Tab = 1
'                  Call gs_SetFocus(txt_NumTom_Es1)
'                  Exit Sub
'               End If
'               If Len(Trim(txt_NumFoj_Es1.Text)) = 0 Then
'                  MsgBox "Debe ingresar el Nro. de Foja.", vbExclamation, modgen_g_str_NomPlt
'                  tab_Genera.Tab = 1
'                  Call gs_SetFocus(txt_NumFoj_Es1)
'                  Exit Sub
'               End If
'               If Len(Trim(txt_NumLib_Es1.Text)) = 0 Then
'                  MsgBox "Debe ingresar el Nro. de Libro.", vbExclamation, modgen_g_str_NomPlt
'                  tab_Genera.Tab = 1
'                  Call gs_SetFocus(txt_NumLib_Es1)
'                  Exit Sub
'               End If
'         End Select
         
         If ipp_MtoHip_Es1.Value = 0 Then
            tab_Genera.Tab = 2
            MsgBox "Debe ingresar el Monto Hipotecado para Estacionamiento 1.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_MtoHip_Es1)
            Exit Sub
         End If
'         If ipp_ValRea_Es1.Value = 0 Then
'            tab_Genera.Tab = 2
'            MsgBox "Debe ingresar el Valor de Realización del Estacionamiento1.", vbExclamation, modgen_g_str_NomPlt
'            Call gs_SetFocus(ipp_ValRea_Es1)
'            Exit Sub
'         End If
         If cmb_Moneda_Es1.ListIndex = -1 Then
            tab_Genera.Tab = 2
            MsgBox "Debe ingresar Moneda para Estacionamiento 1.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmb_Moneda_Es1)
            Exit Sub
         End If
      End If
      
      '---
      
      If cmb_FlgEst_Es2.ListIndex = -1 Then
         MsgBox "Debe seleccionar si hay Hipoteca para Estacionamiento 2.", vbExclamation, modgen_g_str_NomPlt
         tab_Genera.Tab = 3
         Call gs_SetFocus(cmb_FlgEst_Es2)
         Exit Sub
      End If
      
      If cmb_FlgEst_Es2.ItemData(cmb_FlgEst_Es2.ListIndex) = 1 Then
         If CDate(ipp_FecPre_Es2.Text) > date Then
            MsgBox "La Fecha de Presentación no puede ser mayor a la fecha actual.", vbExclamation, modgen_g_str_NomPlt
            tab_Genera.Tab = 3
            Call gs_SetFocus(ipp_FecPre_Es2)
            Exit Sub
         End If
         If Len(Trim(txt_NumPre_Es2.Text)) = 0 Then
            MsgBox "Debe ingresar el Número de Presentación para Estacionamiento 2.", vbExclamation, modgen_g_str_NomPlt
            tab_Genera.Tab = 3
            Call gs_SetFocus(txt_NumPre_Es2)
            Exit Sub
         End If
         If CDate(ipp_FecIns_Es2.Text) > date Then
            MsgBox "La Fecha de Inscripción no puede ser mayor a la fecha actual.", vbExclamation, modgen_g_str_NomPlt
            tab_Genera.Tab = 3
            Call gs_SetFocus(ipp_FecIns_Es2)
            Exit Sub
         End If
         If CDate(ipp_FecIns_Es2.Text) < CDate(ipp_FecPre_Es2.Text) Then
            MsgBox "La Fecha de Inscripción no puede ser menor a la Fecha de Presentación.", vbExclamation, modgen_g_str_NomPlt
            tab_Genera.Tab = 3
            Call gs_SetFocus(ipp_FecIns_Es2)
            Exit Sub
         End If
         If cmb_TipDoc_Es2.ListIndex = -1 Then
            MsgBox "Debe seleccionar el Tipo de Documento Registral para Estacionamiento 2.", vbExclamation, modgen_g_str_NomPlt
            tab_Genera.Tab = 3
            Call gs_SetFocus(cmb_TipDoc_Es2)
            Exit Sub
         End If
         
'         Select Case cmb_TipDoc_Es2.ItemData(cmb_TipDoc_Es2.ListIndex)
'            Case 1
'               If Len(Trim(txt_NumPar_Es2.Text)) = 0 Then
'                  MsgBox "Debe ingresar el Nro. de Partida Electrónica para Estacionamiento 2.", vbExclamation, modgen_g_str_NomPlt
'                  tab_Genera.Tab = 3
'                  Call gs_SetFocus(txt_NumPar_Es2)
'                  Exit Sub
'               End If
'               If Len(Trim(txt_NumAPa_Es2.Text)) = 0 Then
'                  MsgBox "Debe ingresar el Nro. de Asiento en la Partida Electrónica para Estacionamiento 2.", vbExclamation, modgen_g_str_NomPlt
'                  tab_Genera.Tab = 3
'                  Call gs_SetFocus(txt_NumAPa_Es2)
'                  Exit Sub
'               End If
'
'            Case 2
'               If Len(Trim(txt_NumFic_Es2.Text)) = 0 Then
'                  MsgBox "Debe ingresar el Nro. de Ficha Registral para Estacionamiento 2.", vbExclamation, modgen_g_str_NomPlt
'                  tab_Genera.Tab = 3
'                  Call gs_SetFocus(txt_NumFic_Es2)
'                  Exit Sub
'               End If
'               If Len(Trim(txt_NumAFi_Es2.Text)) = 0 Then
'                  MsgBox "Debe ingresar el Nro. de Asiento en la Ficha Registral para Estacionamiento 2.", vbExclamation, modgen_g_str_NomPlt
'                  tab_Genera.Tab = 3
'                  Call gs_SetFocus(txt_NumAFi_Es2)
'                  Exit Sub
'               End If
               
'            Case 3
'               If Len(Trim(txt_NumTom_Es2.Text)) = 0 Then
'                  MsgBox "Debe ingresar el Nro. de Tomo.", vbExclamation, modgen_g_str_NomPlt
'                  tab_Genera.Tab = 3
'                  Call gs_SetFocus(txt_NumTom_Es2)
'                  Exit Sub
'               End If
'               If Len(Trim(txt_NumFoj_Es2.Text)) = 0 Then
'                  MsgBox "Debe ingresar el Nro. de Foja.", vbExclamation, modgen_g_str_NomPlt
'                  tab_Genera.Tab = 3
'                  Call gs_SetFocus(txt_NumFoj_Es2)
'                  Exit Sub
'               End If
'               If Len(Trim(txt_NumLib_Es2.Text)) = 0 Then
'                  MsgBox "Debe ingresar el Nro. de Libro.", vbExclamation, modgen_g_str_NomPlt
'                  tab_Genera.Tab = 3
'                  Call gs_SetFocus(txt_NumLib_Es2)
'                  Exit Sub
'               End If
'         End Select
         
         If ipp_MtoHip_Es2.Value = 0 Then
            tab_Genera.Tab = 3
            MsgBox "Debe ingresar el Monto Hipotecado para Estacionamiento 2.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_MtoHip_Es2)
            Exit Sub
         End If
      End If
      '---

      If cmb_FlgEst_Dep1.ListIndex = -1 Then
         MsgBox "Debe seleccionar si hay Hipoteca para Depósito 1.", vbExclamation, modgen_g_str_NomPlt
         tab_Genera.Tab = 4
         Call gs_SetFocus(cmb_FlgEst_Dep1)
         Exit Sub
      End If
      
      If cmb_FlgEst_Dep1.ItemData(cmb_FlgEst_Dep1.ListIndex) = 1 Then
         If CDate(ipp_FecPre_Dep1.Text) > date Then
            MsgBox "La Fecha de Presentación no puede ser mayor a la fecha actual.", vbExclamation, modgen_g_str_NomPlt
            tab_Genera.Tab = 4
            Call gs_SetFocus(ipp_FecPre_Dep1)
            Exit Sub
         End If
         If Len(Trim(txt_NumPre_Dep1.Text)) = 0 Then
            MsgBox "Debe ingresar el Número de Presentación para Depósito 1.", vbExclamation, modgen_g_str_NomPlt
            tab_Genera.Tab = 4
            Call gs_SetFocus(txt_NumPre_Dep1)
            Exit Sub
         End If
         If CDate(ipp_FecIns_Dep1.Text) > date Then
            MsgBox "La Fecha de Inscripción no puede ser mayor a la fecha actual.", vbExclamation, modgen_g_str_NomPlt
            tab_Genera.Tab = 4
            Call gs_SetFocus(ipp_FecIns_Dep1)
            Exit Sub
         End If
         If CDate(ipp_FecIns_Dep1.Text) < CDate(ipp_FecPre_Dep1.Text) Then
            MsgBox "La Fecha de Inscripción no puede ser menor a la Fecha de Presentación.", vbExclamation, modgen_g_str_NomPlt
            tab_Genera.Tab = 4
            Call gs_SetFocus(ipp_FecIns_Dep1)
            Exit Sub
         End If
         If cmb_TipDoc_Dep1.ListIndex = -1 Then
            MsgBox "Debe seleccionar el Tipo de Documento Registral para Depósito 1.", vbExclamation, modgen_g_str_NomPlt
            tab_Genera.Tab = 4
            Call gs_SetFocus(cmb_TipDoc_Dep1)
            Exit Sub
         End If
         
'         Select Case cmb_TipDoc_Dep1.ItemData(cmb_TipDoc_Dep1.ListIndex)
'            Case 1
'               If Len(Trim(txt_NumPar_Dep1.Text)) = 0 Then
'                  MsgBox "Debe ingresar el Nro. de Partida Electrónica para Depósito 1.", vbExclamation, modgen_g_str_NomPlt
'                  tab_Genera.Tab = 4
'                  Call gs_SetFocus(txt_NumPar_Dep1)
'                  Exit Sub
'               End If
'               If Len(Trim(txt_NumAPa_Dep1.Text)) = 0 Then
'                  MsgBox "Debe ingresar el Nro. de Asiento en la Partida Electrónica para Depósito 1.", vbExclamation, modgen_g_str_NomPlt
'                  tab_Genera.Tab = 4
'                  Call gs_SetFocus(txt_NumAPa_Dep1)
'                  Exit Sub
'               End If
'
'            Case 2
'               If Len(Trim(txt_NumFic_Dep1.Text)) = 0 Then
'                  MsgBox "Debe ingresar el Nro. de Ficha Registral para Depósito 1.", vbExclamation, modgen_g_str_NomPlt
'                  tab_Genera.Tab = 4
'                  Call gs_SetFocus(txt_NumFic_Dep1)
'                  Exit Sub
'               End If
'               If Len(Trim(txt_NumAFi_Dep1.Text)) = 0 Then
'                  MsgBox "Debe ingresar el Nro. de Asiento en la Ficha Registral para Depósito 1.", vbExclamation, modgen_g_str_NomPlt
'                  tab_Genera.Tab = 4
'                  Call gs_SetFocus(txt_NumAFi_Dep1)
'                  Exit Sub
'               End If

'            Case 3
'               If Len(Trim(txt_NumTom_Dep1.Text)) = 0 Then
'                  MsgBox "Debe ingresar el Nro. de Tomo.", vbExclamation, modgen_g_str_NomPlt
'                  tab_Genera.Tab = 4
'                  Call gs_SetFocus(txt_NumTom_Dep1)
'                  Exit Sub
'               End If
'               If Len(Trim(txt_NumFoj_Dep1.Text)) = 0 Then
'                  MsgBox "Debe ingresar el Nro. de Foja.", vbExclamation, modgen_g_str_NomPlt
'                  tab_Genera.Tab = 4
'                  Call gs_SetFocus(txt_NumFoj_Dep1)
'                  Exit Sub
'               End If
'               If Len(Trim(txt_NumLib_Dep1.Text)) = 0 Then
'                  MsgBox "Debe ingresar el Nro. de Libro.", vbExclamation, modgen_g_str_NomPlt
'                  tab_Genera.Tab = 4
'                  Call gs_SetFocus(txt_NumLib_Dep1)
'                  Exit Sub
'               End If
'         End Select
         
         If ipp_MtoHip_Dep1.Value = 0 Then
            tab_Genera.Tab = 4
            MsgBox "Debe ingresar el Monto Hipotecado para Depósito 1.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_MtoHip_Dep1)
            Exit Sub
         End If
'         If ipp_ValRea_Dep1.Value = 0 Then
'            tab_Genera.Tab = 4
'            MsgBox "Debe ingresar el Valor de Realización de Depósito1.", vbExclamation, modgen_g_str_NomPlt
'            Call gs_SetFocus(ipp_ValRea_Dep1)
'            Exit Sub
'         End If
         If cmb_Moneda_Dep1.ListIndex = -1 Then
            tab_Genera.Tab = 4
            MsgBox "Debe ingresar Moneda de Hipoteca para Depósito 1.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmb_Moneda_Dep1)
            Exit Sub
         End If
      End If
      
      If cmb_FlgEst_Dep2.ListIndex = -1 Then
         MsgBox "Debe seleccionar si hay Hipoteca para Depósito 2.", vbExclamation, modgen_g_str_NomPlt
         tab_Genera.Tab = 5
         Call gs_SetFocus(cmb_FlgEst_Dep2)
         Exit Sub
      End If
      
      If cmb_FlgEst_Dep2.ItemData(cmb_FlgEst_Dep2.ListIndex) = 1 Then
         If CDate(ipp_FecPre_Dep2.Text) > date Then
            MsgBox "La Fecha de Presentación no puede ser mayor a la fecha actual.", vbExclamation, modgen_g_str_NomPlt
            tab_Genera.Tab = 5
            Call gs_SetFocus(ipp_FecPre_Dep2)
            Exit Sub
         End If
         If Len(Trim(txt_NumPre_Dep2.Text)) = 0 Then
            MsgBox "Debe ingresar el Número de Presentación de la Hipoteca para Depósito 2.", vbExclamation, modgen_g_str_NomPlt
            tab_Genera.Tab = 5
            Call gs_SetFocus(txt_NumPre_Dep2)
            Exit Sub
         End If
         If CDate(ipp_FecIns_Dep2.Text) > date Then
            MsgBox "La Fecha de Inscripción no puede ser mayor a la fecha actual.", vbExclamation, modgen_g_str_NomPlt
            tab_Genera.Tab = 5
            Call gs_SetFocus(ipp_FecIns_Dep2)
            Exit Sub
         End If
         If CDate(ipp_FecIns_Dep2.Text) < CDate(ipp_FecPre_Dep2.Text) Then
            MsgBox "La Fecha de Inscripción no puede ser menor a la Fecha de Presentación.", vbExclamation, modgen_g_str_NomPlt
            tab_Genera.Tab = 5
            Call gs_SetFocus(ipp_FecIns_Dep2)
            Exit Sub
         End If
         If cmb_TipDoc_Dep2.ListIndex = -1 Then
            MsgBox "Debe seleccionar el Tipo de Documento Registral para Depósito 2.", vbExclamation, modgen_g_str_NomPlt
            tab_Genera.Tab = 5
            Call gs_SetFocus(cmb_TipDoc_Dep2)
            Exit Sub
         End If
         
'         Select Case cmb_TipDoc_Dep2.ItemData(cmb_TipDoc_Dep2.ListIndex)
'            Case 1
'               If Len(Trim(txt_NumPar_Dep2.Text)) = 0 Then
'                  MsgBox "Debe ingresar el Nro. de Partida Electrónica para Depósito 2.", vbExclamation, modgen_g_str_NomPlt
'                  tab_Genera.Tab = 5
'                  Call gs_SetFocus(txt_NumPar_Dep2)
'                  Exit Sub
'               End If
'               If Len(Trim(txt_NumAPa_Dep2.Text)) = 0 Then
'                  MsgBox "Debe ingresar el Nro. de Asiento en la Partida Electrónica para Depósito 2.", vbExclamation, modgen_g_str_NomPlt
'                  tab_Genera.Tab = 5
'                  Call gs_SetFocus(txt_NumAPa_Dep2)
'                  Exit Sub
'               End If
'
'            Case 2
'               If Len(Trim(txt_NumFic_Dep2.Text)) = 0 Then
'                  MsgBox "Debe ingresar el Nro. de Ficha Registral para Depósito 2.", vbExclamation, modgen_g_str_NomPlt
'                  tab_Genera.Tab = 5
'                  Call gs_SetFocus(txt_NumFic_Dep2)
'                  Exit Sub
'               End If
'               If Len(Trim(txt_NumAFi_Dep2.Text)) = 0 Then
'                  MsgBox "Debe ingresar el Nro. de Asiento en la Ficha Registral para Depósito 2.", vbExclamation, modgen_g_str_NomPlt
'                  tab_Genera.Tab = 5
'                  Call gs_SetFocus(txt_NumAFi_Dep2)
'                  Exit Sub
'               End If

'            Case 3
'               If Len(Trim(txt_NumTom_Dep2.Text)) = 0 Then
'                  MsgBox "Debe ingresar el Nro. de Tomo.", vbExclamation, modgen_g_str_NomPlt
'                  tab_Genera.Tab = 5
'                  Call gs_SetFocus(txt_NumTom_Dep2)
'                  Exit Sub
'               End If
'               If Len(Trim(txt_NumFoj_Dep2.Text)) = 0 Then
'                  MsgBox "Debe ingresar el Nro. de Foja.", vbExclamation, modgen_g_str_NomPlt
'                  tab_Genera.Tab = 5
'                  Call gs_SetFocus(txt_NumFoj_Dep2)
'                  Exit Sub
'               End If
'               If Len(Trim(txt_NumLib_Dep2.Text)) = 0 Then
'                  MsgBox "Debe ingresar el Nro. de Libro.", vbExclamation, modgen_g_str_NomPlt
'                  tab_Genera.Tab = 5
'                  Call gs_SetFocus(txt_NumLib_Dep2)
'                  Exit Sub
'               End If
'         End Select
         
         If ipp_MtoHip_Dep2.Value = 0 Then
            tab_Genera.Tab = 5
            MsgBox "Debe ingresar el Monto Hipotecado para Depósito 2.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_MtoHip_Dep2)
            Exit Sub
         End If
'         If ipp_ValRea_Dep2.Value = 0 Then
'            tab_Genera.Tab = 3
'            MsgBox "Debe ingresar el Valor de Realización de Depósito 2.", vbExclamation, modgen_g_str_NomPlt
'            Call gs_SetFocus(ipp_ValRea_Dep2)
'            Exit Sub
'         End If
         If cmb_Moneda_Dep2.ListIndex = -1 Then
            tab_Genera.Tab = 5
            MsgBox "Debe ingresar Moneda de Hipoteca para Depósito 2.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmb_Moneda_Dep2)
            Exit Sub
         End If
      End If
      
      r_str_MsjGrb = "¿Está seguro de grabar los datos?"
      
'      'Validando Total Hipoteca contra Monto puesto en Informe Legal
'      If CDbl(pnl_TotHip.Caption) <> l_dbl_MtoHip Then
'         MsgBox "El Total de la Hipoteca no coincide con el Monto Hipoteca del Informe Legal y el Contrato de Crédito.", vbExclamation, modgen_g_str_NomPlt
'         Exit Sub
'      End If
   End If
   
   'Validando que no haya sobre exposición con el ingreso de la nueva garantía
   If moddat_gf_Consulta_ExposicionGlobal(Mid(pnl_TipDoc.Caption, 1, 1), pnl_NroDoc.Caption, IIf(moddat_g_int_FlgGrb = 4, 0, moddat_g_dbl_TotGar), Cmb_TipGar.ItemData(Cmb_TipGar.ListIndex), IIf(tab_Genera.Tab = 0, ipp_ImpGar.Value, CDbl(pnl_TotHip.Caption))) = True Then   ' moddat_g_str_CodMes, moddat_g_str_CodAno,
      If moddat_g_str_DesObs = "" Then
         MsgBox "SobreExposición, ingresar Garantía Hipotecaria o Líquida.", vbExclamation, modgen_g_str_NomPlt
      Else
         MsgBox "SobreExposición, el monto de la garantía debe ser como mínimo " & moddat_g_str_DesObs, vbExclamation, modgen_g_str_NomPlt  'Hipotecaria o Líquida debe ingresar un monto mayor de la Garantía."
      End If
      If MsgBox("Desea ingresar la garantía? ", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
         If tab_Genera.Tab = 0 Then
            Call gs_SetFocus(ipp_ImpGar)
         Else
            Call gs_SetFocus(ipp_MtoHip_Inm)
         End If
         Exit Sub
      End If
   End If
   
   'Validando que las hipotecas creadas no superen el % de límite de exposición
   'Gar. Líquida
'   If cmb_TipGar.ItemData(cmb_TipGar.ListIndex) = 1 Then
'      If fs_Valida_Garantia(CDbl(ipp_ImpGar.Value)) = False Then
'         MsgBox "La Garantía Líquida, supera el 30% del Patrimonio Efectivo.", vbExclamation, modgen_g_str_NomPlt
'         Call gs_SetFocus(ipp_ImpGar)
'         Exit Sub
'      End If
'   Else
'   'Gar. Hipotecaria
'      If fs_Valida_Garantia(CDbl(CDbl(ipp_MtoHip_Inm.Value) + CDbl(ipp_MtoHip_Es1.Value) + CDbl(ipp_MtoHip_Es2.Value) + CDbl(ipp_MtoHip_Dep1.Value) + CDbl(ipp_MtoHip_Dep2.Value))) = False Then
'         MsgBox "La Garantía Líquida, supera el 15% del Patrimonio Efectivo.", vbExclamation, modgen_g_str_NomPlt
'         If tab_Genera.Tab = 0 Then
'            Call gs_SetFocus(ipp_ImpGar)
'         Else
'            Call gs_SetFocus(ipp_MtoHip_Inm)
'         End If
'         Exit Sub
'      End If
'   End If
   
   If MsgBox(r_str_MsjGrb, vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   If Cmb_TipGar.ItemData(Cmb_TipGar.ListIndex) = 1 Then
   
      Do While moddat_g_int_FlgGOK = False
         Screen.MousePointer = 11
         Call moddat_gs_FecSis
   
         If moddat_g_int_FlgAct_2 = 1 Then  'moddat_g_int_FlgGrb_2
            r_int_NumIte = fs_GeneraNumIte
         Else
            r_int_NumIte = l_int_NumIte
         End If
         
         g_str_Parame = "USP_TPR_MAEGAR ("
         g_str_Parame = g_str_Parame & CStr(r_int_NumIte) & ", "
         g_str_Parame = g_str_Parame & moddat_g_int_TipDoc & ", "
         g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumDoc & "', "
         If pnl_NumCFi.Caption = "" Then
            g_str_Parame = g_str_Parame & "' ', "
         Else
            g_str_Parame = g_str_Parame & "'" & CStr(Trim(Replace(pnl_NumCFi.Caption, "-", ""))) & "', "
         End If
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & CStr(Cmb_TipGar.ItemData(Cmb_TipGar.ListIndex)) & ", "
         g_str_Parame = g_str_Parame & "'" & Format(ipp_FecEmi.Value, "yyyymmdd") & "', "
         g_str_Parame = g_str_Parame & CStr(cmb_Moneda.ItemData(cmb_Moneda.ListIndex)) & ", "
         g_str_Parame = g_str_Parame & CStr(ipp_ImpGar.Value) & ", "
         g_str_Parame = g_str_Parame & CStr(moddat_g_int_FlgAct_2) & ", " 'moddat_g_int_FlgGrb_2
   
          'Datos de Auditoria
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
   
   Else
      Do While moddat_g_int_FlgGOK = False
       
         Screen.MousePointer = 11
'         Call moddat_gs_FecSis
         
         If moddat_g_int_FlgAct_2 = 1 Then 'moddat_g_int_FlgGrb_2
            r_int_NumIte = fs_GeneraNumIte
         Else
            r_int_NumIte = l_int_NumIte
         End If
         
         'Obtener Código de Tasación
         r_int_CodTas = fs_ObtenerCodTas
         '--
         g_str_Parame = "usp_tpr_maegar_crea ("
         g_str_Parame = g_str_Parame & CStr(r_int_NumIte) & ", "
         g_str_Parame = g_str_Parame & moddat_g_int_TipDoc & ", "
         g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumDoc & "', "
         g_str_Parame = g_str_Parame & CStr(r_int_CodTas) & ", "
         g_str_Parame = g_str_Parame & CStr(Cmb_TipGar.ItemData(Cmb_TipGar.ListIndex)) & ", "
         g_str_Parame = g_str_Parame & "'" & Format(cmb_SedReg.ItemData(cmb_SedReg.ListIndex), "0000") & "', "
         g_str_Parame = g_str_Parame & "' ', "
         'g_str_Parame = g_str_Parame & "1, "
         
         'Inmueble
         g_str_Parame = g_str_Parame & Format(CDate(ipp_FecPre_Inm.Text), "yyyymmdd") & ", "
         g_str_Parame = g_str_Parame & "'" & txt_NumPre_Inm.Text & "', "
         g_str_Parame = g_str_Parame & Format(CDate(ipp_FecIns_Inm.Text), "yyyymmdd") & ", "
         
         If cmb_TipDoc_Inm.ListIndex = -1 Then
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            
         Else
            g_str_Parame = g_str_Parame & CStr(cmb_TipDoc_Inm.ItemData(cmb_TipDoc_Inm.ListIndex)) & ", "
            If cmb_TipDoc_Inm.ItemData(cmb_TipDoc_Inm.ListIndex) = 1 Or cmb_TipDoc_Inm.ItemData(cmb_TipDoc_Inm.ListIndex) = 2 Then
               If cmb_TipDoc_Inm.ItemData(cmb_TipDoc_Inm.ListIndex) = 1 Then
                  g_str_Parame = g_str_Parame & "'" & txt_NumPar_Inm.Text & "', "
                  g_str_Parame = g_str_Parame & "'" & txt_NumAPa_Inm.Text & "', "
               Else
                  g_str_Parame = g_str_Parame & "'" & txt_NumFic_Inm.Text & "', "
                  g_str_Parame = g_str_Parame & "'" & txt_NumAFi_Inm.Text & "', "
               End If
            Else
               g_str_Parame = g_str_Parame & "'', "
               g_str_Parame = g_str_Parame & "'', "
            End If
         End If
         
         g_str_Parame = g_str_Parame & CStr(CDbl(ipp_MtoHip_Inm.Text)) & ", "
         g_str_Parame = g_str_Parame & CStr(cmb_Moneda_Inm.ItemData(cmb_Moneda_Inm.ListIndex)) & ", "
         
         'Estacionamiento 1
         g_str_Parame = g_str_Parame & cmb_FlgEst_Es1.ItemData(cmb_FlgEst_Es1.ListIndex) & ", "
         
         If cmb_FlgEst_Es1.ItemData(cmb_FlgEst_Es1.ListIndex) = 1 Then
            g_str_Parame = g_str_Parame & Format(CDate(ipp_FecPre_Es1.Text), "yyyymmdd") & ", "
            g_str_Parame = g_str_Parame & "'" & txt_NumPre_Es1.Text & "', "
            g_str_Parame = g_str_Parame & Format(CDate(ipp_FecIns_Es1.Text), "yyyymmdd") & ", "
            g_str_Parame = g_str_Parame & CStr(cmb_TipDoc_Es1.ItemData(cmb_TipDoc_Es1.ListIndex)) & ", "
            
            If cmb_TipDoc_Es1.ItemData(cmb_TipDoc_Es1.ListIndex) = 1 Or cmb_TipDoc_Es1.ItemData(cmb_TipDoc_Es1.ListIndex) = 2 Then
               If cmb_TipDoc_Es1.ItemData(cmb_TipDoc_Es1.ListIndex) = 1 Then
                  g_str_Parame = g_str_Parame & "'" & txt_NumPar_Es1.Text & "', "
                  g_str_Parame = g_str_Parame & "'" & txt_NumAPa_Es1.Text & "', "
               Else
                  g_str_Parame = g_str_Parame & "'" & txt_NumFic_Es1.Text & "', "
                  g_str_Parame = g_str_Parame & "'" & txt_NumAFi_Es1.Text & "', "
               End If
            Else
               g_str_Parame = g_str_Parame & "'', "
               g_str_Parame = g_str_Parame & "'', "
            End If
            g_str_Parame = g_str_Parame & CStr(CDbl(ipp_MtoHip_Es1.Text)) & ", "
            g_str_Parame = g_str_Parame & CStr(cmb_Moneda_Es1.ItemData(cmb_Moneda_Es1.ListIndex)) & ", "
         Else
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
         End If
         
         'Estacionamiento 2
         g_str_Parame = g_str_Parame & cmb_FlgEst_Es2.ItemData(cmb_FlgEst_Es2.ListIndex) & ", "
         
         If cmb_FlgEst_Es2.ItemData(cmb_FlgEst_Es2.ListIndex) = 1 Then
            g_str_Parame = g_str_Parame & Format(CDate(ipp_FecPre_Es2.Text), "yyyymmdd") & ", "
            g_str_Parame = g_str_Parame & "'" & txt_NumPre_Es2.Text & "', "
            g_str_Parame = g_str_Parame & Format(CDate(ipp_FecIns_Es2.Text), "yyyymmdd") & ", "
            g_str_Parame = g_str_Parame & CStr(cmb_TipDoc_Es2.ItemData(cmb_TipDoc_Es2.ListIndex)) & ", "
            
            If cmb_TipDoc_Es2.ItemData(cmb_TipDoc_Es2.ListIndex) = 1 Or cmb_TipDoc_Es2.ItemData(cmb_TipDoc_Es2.ListIndex) = 2 Then
               If cmb_TipDoc_Es2.ItemData(cmb_TipDoc_Es2.ListIndex) = 1 Then
                  g_str_Parame = g_str_Parame & "'" & txt_NumPar_Es2.Text & "', "
                  g_str_Parame = g_str_Parame & "'" & txt_NumAPa_Es2.Text & "', "
               Else
                  g_str_Parame = g_str_Parame & "'" & txt_NumFic_Es2.Text & "', "
                  g_str_Parame = g_str_Parame & "'" & txt_NumAFi_Es2.Text & "', "
               End If
            Else
               g_str_Parame = g_str_Parame & "'', "
               g_str_Parame = g_str_Parame & "'', "
            End If
            g_str_Parame = g_str_Parame & CStr(CDbl(ipp_MtoHip_Es2.Text)) & ", "
            g_str_Parame = g_str_Parame & CStr(cmb_Moneda_Es2.ItemData(cmb_Moneda_Es2.ListIndex)) & ", "
            
         Else
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
         End If
         
         'Depósito1
         g_str_Parame = g_str_Parame & cmb_FlgEst_Dep1.ItemData(cmb_FlgEst_Dep1.ListIndex) & ", "
         
         If cmb_FlgEst_Dep1.ItemData(cmb_FlgEst_Dep1.ListIndex) = 1 Then
            g_str_Parame = g_str_Parame & Format(CDate(ipp_FecPre_Dep1.Text), "yyyymmdd") & ", "
            g_str_Parame = g_str_Parame & "'" & txt_NumPre_Dep1.Text & "', "
            g_str_Parame = g_str_Parame & Format(CDate(ipp_FecIns_Dep1.Text), "yyyymmdd") & ", "
            g_str_Parame = g_str_Parame & CStr(cmb_TipDoc_Dep1.ItemData(cmb_TipDoc_Dep1.ListIndex)) & ", "
            
            If cmb_TipDoc_Dep1.ItemData(cmb_TipDoc_Dep1.ListIndex) = 1 Or cmb_TipDoc_Dep1.ItemData(cmb_TipDoc_Dep1.ListIndex) = 2 Then
               If cmb_TipDoc_Dep1.ItemData(cmb_TipDoc_Dep1.ListIndex) = 1 Then
                  g_str_Parame = g_str_Parame & "'" & txt_NumPar_Dep1.Text & "', "
                  g_str_Parame = g_str_Parame & "'" & txt_NumAPa_Dep1.Text & "', "
               Else
                  g_str_Parame = g_str_Parame & "'" & txt_NumFic_Dep1.Text & "', "
                  g_str_Parame = g_str_Parame & "'" & txt_NumAFi_Dep1.Text & "', "
               End If
            Else
               g_str_Parame = g_str_Parame & "'', "
               g_str_Parame = g_str_Parame & "'', "
            End If
            g_str_Parame = g_str_Parame & CStr(CDbl(ipp_MtoHip_Dep1.Text)) & ", "
            g_str_Parame = g_str_Parame & CStr(cmb_Moneda_Dep1.ItemData(cmb_Moneda_Dep1.ListIndex)) & ", "
           
         Else
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
         End If
         
         'Depósito2
         g_str_Parame = g_str_Parame & cmb_FlgEst_Dep2.ItemData(cmb_FlgEst_Dep2.ListIndex) & ", "
         
         If cmb_FlgEst_Dep2.ItemData(cmb_FlgEst_Dep2.ListIndex) = 1 Then
            g_str_Parame = g_str_Parame & Format(CDate(ipp_FecPre_Dep2.Text), "yyyymmdd") & ", "
            g_str_Parame = g_str_Parame & "'" & txt_NumPre_Dep2.Text & "', "
            g_str_Parame = g_str_Parame & Format(CDate(ipp_FecIns_Dep2.Text), "yyyymmdd") & ", "
            g_str_Parame = g_str_Parame & CStr(cmb_TipDoc_Dep2.ItemData(cmb_TipDoc_Dep2.ListIndex)) & ", "
            
            If cmb_TipDoc_Dep2.ItemData(cmb_TipDoc_Dep2.ListIndex) = 1 Or cmb_TipDoc_Dep2.ItemData(cmb_TipDoc_Dep2.ListIndex) = 2 Then
               If cmb_TipDoc_Dep2.ItemData(cmb_TipDoc_Dep2.ListIndex) = 1 Then
                  g_str_Parame = g_str_Parame & "'" & txt_NumPar_Dep2.Text & "', "
                  g_str_Parame = g_str_Parame & "'" & txt_NumAPa_Dep2.Text & "', "
               Else
                  g_str_Parame = g_str_Parame & "'" & txt_NumFic_Dep2.Text & "', "
                  g_str_Parame = g_str_Parame & "'" & txt_NumAFi_Dep2.Text & "', "
               End If
            Else
               g_str_Parame = g_str_Parame & "'', "
               g_str_Parame = g_str_Parame & "'', "
            End If
            g_str_Parame = g_str_Parame & CStr(CDbl(ipp_MtoHip_Dep2.Text)) & ", "
            g_str_Parame = g_str_Parame & CStr(cmb_Moneda_Dep2.ItemData(cmb_Moneda_Dep2.ListIndex)) & ", "
            
         Else
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
         End If
        
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & CStr(moddat_g_int_FlgAct_2) & ", " 'moddat_g_int_FlgGrb_2
         
         'Datos de Auditoria
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
            
   End If
   
   'Generar Asientos automáticos
'   If moddat_g_int_FlgGOK = True And moddat_g_int_FlgGrb_2 = 1 Then
'      Call fs_GeneraAsiento(Trim(Replace(pnl_NumRef.Caption, "-", "")), CStr(moddat_g_int_TipDoc), CStr(moddat_g_str_NumDoc), CStr(moddat_g_str_NomCli), "151719010104", "251419010110", CDbl(ipp_ImpGar.Value), r_int_NumIte)
'   End If
         
   MsgBox "Actualización realizada satisfactoriamente.", vbInformation, modgen_g_con_PltPar
   Call fs_Buscar
   Call fs_Limpia
   Call fs_Activa(False)
   Call gs_SetFocus(cmd_Agrega)
   Call fs_ConsultaNumRef(moddat_g_int_TipDoc, moddat_g_str_NumDoc)
   Call frm_Ges_TecPro_01.fs_Buscar

'   Unload Me
End Sub
Private Function fs_Valida_Garantia(ByVal p_MtoGar As Double) As Boolean
Dim r_dbl_ValExSG  As Double
Dim r_dbl_ValExGH  As Double
Dim r_dbl_ValExGL  As Double
Dim r_dbl_ValToGa  As Double
Dim r_dbl_ValToCa  As Double
   
   fs_Valida_Garantia = False
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "    SELECT  "
   g_str_Parame = g_str_Parame & "           ( SELECT NVL(SUM(NVL(MAEGAR_MTOGAR_INM,0) + NVL(MAEGAR_MTOGAR_ES1,0) + NVL(MAEGAR_MTOGAR_ES2,0) + NVL(MAEGAR_MTOGAR_DE1,0) + NVL(MAEGAR_MTOGAR_DE2,0)),0)  "
   g_str_Parame = g_str_Parame & "               FROM TPR_MAEGAR "
   g_str_Parame = g_str_Parame & "              WHERE MAEGAR_TIPDOC = '" & moddat_g_int_TipDoc & "' "
   g_str_Parame = g_str_Parame & "                AND MAEGAR_NUMDOC = '" & moddat_g_str_NumDoc & "' "
   g_str_Parame = g_str_Parame & "                AND MAEGAR_TIPGAR = 1 " '" & Cmb_TipGar.ItemData(Cmb_TipGar.ListIndex) & ""
   g_str_Parame = g_str_Parame & "                AND MAEGAR_SITUAC = 1) AS GARANTIA_LIQUIDA, "
   
   g_str_Parame = g_str_Parame & "           ( SELECT NVL(SUM(NVL(MAEGAR_MTOGAR_INM,0) + NVL(MAEGAR_MTOGAR_ES1,0) + NVL(MAEGAR_MTOGAR_ES2,0) + NVL(MAEGAR_MTOGAR_DE1,0) + NVL(MAEGAR_MTOGAR_DE2,0)),0)  "
   g_str_Parame = g_str_Parame & "               FROM TPR_MAEGAR "
   g_str_Parame = g_str_Parame & "              WHERE MAEGAR_TIPDOC = '" & moddat_g_int_TipDoc & "' "
   g_str_Parame = g_str_Parame & "                AND MAEGAR_NUMDOC = '" & moddat_g_str_NumDoc & "' "
   g_str_Parame = g_str_Parame & "                AND MAEGAR_TIPGAR = 2 "
   g_str_Parame = g_str_Parame & "                AND MAEGAR_SITUAC = 1) AS GARANTIA_HIPOTECARIA, "
   
   g_str_Parame = g_str_Parame & "          ( SELECT NVL(SUM(MAECFI_GARFIA), 0) "
   g_str_Parame = g_str_Parame & "              FROM TPR_MAECFI "
   g_str_Parame = g_str_Parame & "             WHERE MAECFI_TIPDOC = " & Mid(pnl_TipDoc.Caption, 1, 1) & " "
   g_str_Parame = g_str_Parame & "               AND MAECFI_NUMDOC = '" & pnl_NroDoc.Caption & "'"
   g_str_Parame = g_str_Parame & "               AND MAECFI_CODPRD <> '008' "
   g_str_Parame = g_str_Parame & "               AND MAECFI_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                                                            ) + "
   
   g_str_Parame = g_str_Parame & "          ( SELECT NVL(SUM(MAECFI_IMPFIA), 0) "
   g_str_Parame = g_str_Parame & "              FROM TPR_MAECFI "
   g_str_Parame = g_str_Parame & "             WHERE MAECFI_TIPDOC = " & Mid(pnl_TipDoc.Caption, 1, 1) & " "
   g_str_Parame = g_str_Parame & "               AND MAECFI_NUMDOC = '" & pnl_NroDoc.Caption & "' "
   g_str_Parame = g_str_Parame & "               AND MAECFI_CODPRD = '008' "
   g_str_Parame = g_str_Parame & "               AND MAECFI_SITUAC = 1 )"
   g_str_Parame = g_str_Parame & "           AS MONTO_GARANTIZADO  "
   g_str_Parame = g_str_Parame & "      FROM DUAL"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Function
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      Exit Function
   End If
   
   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      g_rst_Genera.MoveFirst
      
      If moddat_g_int_FlgGrb = 4 Then
         moddat_g_dbl_TotGar = 0
      End If
      
      'Total de Garantías
      r_dbl_ValToGa = CDbl(CDbl(g_rst_Genera!GARANTIA_LIQUIDA) + CDbl(g_rst_Genera!GARANTIA_HIPOTECARIA) + CDbl(p_MtoGar))
      
      'Total de Créditos Directos e Indirectos
      r_dbl_ValToCa = CDbl(CDbl(g_rst_Genera!MONTO_GARANTIZADO) + CDbl(moddat_g_dbl_TotGar))
      
      'Cuando no hay Sobre Exposición
      If CDbl(CDbl(moddat_g_dbl_ValNv1) + CDbl(r_dbl_ValToGa)) >= CDbl(r_dbl_ValToCa) Then
        
         'Gar. Líquida
         If Cmb_TipGar.ItemData(Cmb_TipGar.ListIndex) = 1 Then
            If Round(CDbl(p_MtoGar) + CDbl(g_rst_Genera!GARANTIA_LIQUIDA) + CDbl(r_dbl_ValToCa), 2) <= Round(moddat_g_dbl_ValNv2, 2) Then '- moddat_g_dbl_ValSGa
               fs_Valida_Garantia = True
            End If
            
         'Gar. Hipotecaria
         Else
            If Round(CDbl(p_MtoGar) + CDbl(g_rst_Genera!GARANTIA_HIPOTECARIA) + CDbl(r_dbl_ValToCa), 2) <= Round(moddat_g_dbl_ValNv3, 2) Then '- moddat_g_dbl_ValSGa
               fs_Valida_Garantia = True
            End If
         End If
      End If
   End If
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   
End Function
Private Function fs_Validar_MtoGar() As Boolean
   fs_Validar_MtoGar = False
      
'   g_str_Parame = ""
'   g_str_Parame = g_str_Parame & "  SELECT MAECFI_IMPFIA CARTA_FIANZA, NVL(GARANTIA,0) GARANTIA "
'   g_str_Parame = g_str_Parame & "    FROM TPR_MAECFI "
'   g_str_Parame = g_str_Parame & "         LEFT JOIN (SELECT MAEGAR_NUMREF, NVL(SUM(MAEGAR_MTOGAR),0) GARANTIA "
'   g_str_Parame = g_str_Parame & "                       FROM TPR_MAEGAR "
'   g_str_Parame = g_str_Parame & "                      GROUP BY MAEGAR_NUMREF) B ON MAECFI_NUMREF = MAEGAR_NUMREF "
'   g_str_Parame = g_str_Parame & "   WHERE MAECFI_NUMREF =  '" & CStr(Trim(Replace(cmb_NumRef.Text, "-", ""))) & "' "

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT MAEETE_LINASI LINEA_ASIGNADA, NVL(GARANTIA,0) GARANTIA "
   g_str_Parame = g_str_Parame & "    FROM TPR_MAEETE "
   g_str_Parame = g_str_Parame & "         LEFT JOIN (SELECT MAEGAR_TIPDOC, MAEGAR_NUMDOC, NVL(SUM(NVL(MAEGAR_MTOGAR_INM,0) + NVL(MAEGAR_MTOGAR_ES1,0) + NVL(MAEGAR_MTOGAR_ES2,0) + NVL(MAEGAR_MTOGAR_DE1,0) + NVL(MAEGAR_MTOGAR_DE2,0)),0) GARANTIA "
   g_str_Parame = g_str_Parame & "                       FROM TPR_MAEGAR "
   g_str_Parame = g_str_Parame & "                      GROUP BY MAEGAR_TIPDOC, MAEGAR_NUMDOC ) B ON MAEETE_TIPDOC = MAEGAR_TIPDOC AND MAEETE_NUMDOC = MAEGAR_NUMDOC "
   g_str_Parame = g_str_Parame & "   WHERE MAEETE_TIPDOC =  '" & moddat_g_int_TipDoc & "' "
   g_str_Parame = g_str_Parame & "     AND MAEETE_NUMDOC =  '" & moddat_g_str_NumDoc & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
     Exit Function
   End If
   
   If g_rst_GenAux.BOF And g_rst_GenAux.EOF Then
      g_rst_GenAux.Close
      Set g_rst_GenAux = Nothing
      Exit Function
   End If
     
   If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
      If moddat_g_int_FlgAct_2 = 2 Then 'moddat_g_int_FlgGrb_2
         If CDbl(g_rst_GenAux!LINEA_ASIGNADA) >= CDbl(ipp_ImpGar.Value) Then
            fs_Validar_MtoGar = True
         End If
      Else
         If CDbl(g_rst_GenAux!LINEA_ASIGNADA) >= CDbl(g_rst_GenAux!GARANTIA) + CDbl(ipp_ImpGar.Value) Then
            fs_Validar_MtoGar = True
         End If
      End If
   End If
End Function
Private Function fs_Buscar_Cartas_NoCliente() As Boolean
   Dim r_str_Cadena  As String
   Dim r_str_CadAux  As String
   Dim r_str_CadRef  As String
   
   r_str_Cadena = CStr(Trim(Replace(pnl_NumCFi.Caption, "-", "")))
      
   fs_Buscar_Cartas_NoCliente = False
   
   r_str_CadAux = Trim(r_str_Cadena)
   
   If Len(r_str_CadAux) > 1 Then
      While InStr(r_str_CadAux, "|")
         r_str_CadRef = Trim(Mid(r_str_CadAux, 1, InStr(r_str_CadAux, "|") - 1))
         r_str_CadRef = fs_Obtener_NumRef(r_str_CadRef)
         r_str_CadAux = Trim(Mid(r_str_CadAux, InStr(r_str_CadAux, "|") + 1))
      Wend
            
      r_str_CadRef = r_str_CadRef & "|" & fs_Obtener_NumRef(r_str_CadAux)
      If InStr(r_str_CadRef, "|") = 1 Then
         r_str_CadRef = Replace(r_str_CadRef, "|", "")
      End If
      r_str_CadRef = Replace(r_str_CadRef, "|", "' , '")
      
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "  SELECT MAECFI_NOCLIE "
      g_str_Parame = g_str_Parame & "    FROM TPR_MAECFI "
      g_str_Parame = g_str_Parame & "   WHERE MAECFI_NUMREF IN ('" & r_str_CadRef & "')  " ' =  '" & CStr(Trim(Replace(pnl_NumCFi.Caption, "-", ""))) & "' "
   
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
        Exit Function
      End If
      
      If g_rst_GenAux.BOF And g_rst_GenAux.EOF Then
         g_rst_GenAux.Close
         Set g_rst_GenAux = Nothing
         Exit Function
      End If
      g_rst_GenAux.MoveFirst
      Do While Not g_rst_GenAux.EOF
         If g_rst_GenAux!MAECFI_NOCLIE = 1 Then
            fs_Buscar_Cartas_NoCliente = True
         Else
            fs_Buscar_Cartas_NoCliente = False
            Exit Function
         End If
         g_rst_GenAux.MoveNext
      Loop
        
   End If
End Function

Private Function fs_Validar_MtoCfiGar() As Boolean
   Dim r_str_Cadena  As String
   Dim r_str_CadAux  As String
   Dim r_str_CadRef  As String
   
   r_str_Cadena = CStr(Trim(Replace(pnl_NumCFi.Caption, "-", "")))
   'r_str_Cadena = Replace(r_str_Cadena, "|", "' , '")
   
   fs_Validar_MtoCfiGar = False
   
   r_str_CadAux = Trim(r_str_Cadena)
   If Len(r_str_CadAux) > 1 Then
      While InStr(r_str_CadAux, "|")
         r_str_CadRef = Trim(Mid(r_str_CadAux, 1, InStr(r_str_CadAux, "|") - 1))
         r_str_CadRef = fs_Obtener_NumRef(r_str_CadRef)
         r_str_CadAux = Trim(Mid(r_str_CadAux, InStr(r_str_CadAux, "|") + 1))
      Wend
            
      r_str_CadRef = r_str_CadRef & "|" & fs_Obtener_NumRef(r_str_CadAux)
      If InStr(r_str_CadRef, "|") = 1 Then
         r_str_CadRef = Replace(r_str_CadRef, "|", "")
      End If
      r_str_CadRef = Replace(r_str_CadRef, "|", "' , '")
   End If
      
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT NVL(SUM(MAECFI_IMPFIA),0) MTOCFI "
   g_str_Parame = g_str_Parame & "    FROM TPR_MAECFI "
   g_str_Parame = g_str_Parame & "   WHERE MAECFI_NUMREF IN  ('" & r_str_CadRef & "') AND MAECFI_NOCLIE = 1 " ' =  '" & CStr(Trim(Replace(pnl_NumCFi.Caption, "-", ""))) & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
     Exit Function
   End If
   
   If g_rst_GenAux.BOF And g_rst_GenAux.EOF Then
      g_rst_GenAux.Close
      Set g_rst_GenAux = Nothing
      Exit Function
   End If
     
   If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
      If CDbl(g_rst_GenAux!MTOCFI) = CDbl(ipp_ImpGar.Value) Then 'If CDbl(g_rst_GenAux!MTOCFI) <= CDbl(ipp_ImpGar.Value) Then
         fs_Validar_MtoCfiGar = True
      End If
   End If
End Function
Private Function fs_Validar_MtoGarPag() As Boolean
Dim r_dbl_ValMod  As Double
Dim r_int_Contad  As Integer
Dim r_dbl_MtoTot  As Double

   fs_Validar_MtoGarPag = False
   If grd_Listad.Rows > 0 Then
      r_dbl_ValMod = CDbl(grd_Listad.TextMatrix(grd_Listad.Row, 4))
   End If
   For r_int_Contad = 0 To grd_Listad.Rows - 1
      r_dbl_MtoTot = r_dbl_MtoTot + CDbl(grd_Listad.TextMatrix(r_int_Contad, 4))
   Next r_int_Contad
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT MAEGAR_NUMREF, SUM(NVL(MAEGAR_MTOGAR_INM,0) + NVL(MAEGAR_MTOGAR_ES1,0) + NVL(MAEGAR_MTOGAR_ES2,0) + NVL(MAEGAR_MTOGAR_DE1,0) + NVL(MAEGAR_MTOGAR_DE2,0)) GARANTIA ,"
   g_str_Parame = g_str_Parame & "         NVL((SELECT NVL(SUM(NVL(MAERDE_IMPORT,0)),0) "
   g_str_Parame = g_str_Parame & "                FROM TPR_MAERDE "
   g_str_Parame = g_str_Parame & "               WHERE MAERDE_CODIGO = 6 "
   g_str_Parame = g_str_Parame & "                 AND MAERDE_NUMREF = MAEGAR_NUMREF "
   g_str_Parame = g_str_Parame & "               GROUP BY MAERDE_NUMREF),0) GARANTIA_PAGADO "
   g_str_Parame = g_str_Parame & "    FROM TPR_MAEGAR "
   g_str_Parame = g_str_Parame & "   WHERE MAEGAR_NUMREF = '" & CStr(Trim(Replace(cmb_NumRef.Text, "-", ""))) & "' "
   g_str_Parame = g_str_Parame & "   GROUP BY MAEGAR_NUMREF "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
     Exit Function
   End If
   
   If g_rst_GenAux.BOF And g_rst_GenAux.EOF Then
      fs_Validar_MtoGarPag = True
      g_rst_GenAux.Close
      Set g_rst_GenAux = Nothing
      Exit Function
   End If
     
   If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
      If moddat_g_int_FlgAct_2 = 2 Then  'moddat_g_int_FlgGrb_2
          If CDbl(g_rst_GenAux!GARANTIA_PAGADO) <= CDbl(r_dbl_MtoTot - r_dbl_ValMod) + CDbl(ipp_ImpGar.Value) Then
            fs_Validar_MtoGarPag = True
         End If
      Else
         If CDbl(g_rst_GenAux!GARANTIA_PAGADO) <= CDbl(g_rst_GenAux!GARANTIA_PAGADO) + CDbl(ipp_ImpGar.Value) Then
            fs_Validar_MtoGarPag = True
         End If
      End If
   End If
End Function
Private Function fs_GeneraNumIte() As Integer
Dim r_str_Parame     As String

   fs_GeneraNumIte = 0
   
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & " SELECT NVL(MAX(MAEGAR_NUMITE),0) NUMITE FROM TPR_MAEGAR "
   r_str_Parame = r_str_Parame & "  WHERE MAEGAR_TIPDOC =  " & moddat_g_int_TipDoc & ""
   r_str_Parame = r_str_Parame & "    AND MAEGAR_NUMDOC =  '" & CStr(moddat_g_str_NumDoc) & "'"
   
   If Not gf_EjecutaSQL(r_str_Parame, g_rst_GenAux, 3) Then
       Exit Function
   End If
   
   If g_rst_GenAux.BOF And g_rst_GenAux.EOF Then
      g_rst_GenAux.Close
      Set g_rst_GenAux = Nothing
      Exit Function
   End If
     
   If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
      fs_GeneraNumIte = g_rst_GenAux!NUMITE + 1
   End If
End Function
Private Function fs_ObtenerCodTas() As Integer
Dim r_str_Parame     As String

   fs_ObtenerCodTas = 0
   
   If cmb_EmpPer.ListIndex > -1 And cmb_NumInf.ListIndex > -1 And cmb_FecEva.ListIndex > -1 Then
      r_str_Parame = ""
      r_str_Parame = r_str_Parame & " SELECT NVL(EVATAS_CODTAS,0) CODTAS FROM TPR_EVATAS "
      r_str_Parame = r_str_Parame & "  WHERE EVATAS_TIPDOC =  " & moddat_g_int_TipDoc & ""
      r_str_Parame = r_str_Parame & "    AND EVATAS_NUMDOC =  '" & CStr(moddat_g_str_NumDoc) & "'"
      r_str_Parame = r_str_Parame & "    AND EVATAS_CODEMP =  '" & Format(CStr(cmb_EmpPer.ItemData(cmb_EmpPer.ListIndex)), "000000") & "'"
      r_str_Parame = r_str_Parame & "    AND EVATAS_NUMINF =  '" & CStr(cmb_NumInf.Text) & "'"
      r_str_Parame = r_str_Parame & "    AND EVATAS_FECEVA =  '" & CStr(cmb_FecEva.ItemData(cmb_FecEva.ListIndex)) & "'"
      
      If Not gf_EjecutaSQL(r_str_Parame, g_rst_GenAux, 3) Then
          Exit Function
      End If
      
      If g_rst_GenAux.BOF And g_rst_GenAux.EOF Then
         g_rst_GenAux.Close
         Set g_rst_GenAux = Nothing
         Exit Function
      End If
        
      If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
         fs_ObtenerCodTas = g_rst_GenAux!CODTAS
      End If
   End If
End Function
Private Sub fs_GeneraAsiento(ByVal p_NumRef As String, ByVal p_TipDoc As String, ByVal p_NumDoc As String, ByVal p_RazSoc As String, ByVal p_CtaDeb As String, ByVal p_CtaHab As String, ByVal p_Importe As Double, ByVal p_NumIte As Integer)
Dim r_arr_LogPro()      As modprc_g_tpo_LogPro
Dim r_int_NumIte        As Integer
Dim r_str_AsiGen        As String
Dim r_str_Origen        As String
Dim r_str_TipNot        As String
Dim r_int_NumLib        As Integer
Dim r_int_NumAsi        As Integer
Dim r_str_FecCon        As String
Dim r_str_FecReg        As String
Dim r_str_CtaCtb        As String
Dim r_str_DebHab        As String
Dim r_str_Glosa         As String
Dim r_dbl_MtoSol        As Double
Dim r_dbl_MtoDol        As Double
Dim r_dbl_importe       As Double
Dim r_dbl_TipCam        As Double
Dim r_int_ConAux        As Integer
Dim r_str_NroCnt        As String

   ReDim r_arr_LogPro(0)
   ReDim r_arr_LogPro(1)
   r_arr_LogPro(1).LogPro_CodPro = "CTBP1013"
   r_arr_LogPro(1).LogPro_FInEje = Format(date, "yyyymmdd")
   r_arr_LogPro(1).LogPro_HInEje = Format(Time, "hhmmss")
   r_arr_LogPro(1).LogPro_NumErr = 0
   r_str_Origen = "LM"
   r_str_TipNot = "D"
   r_int_NumLib = 6
   r_str_AsiGen = ""
   r_int_NumAsi = 0
   r_int_NumIte = 0
      

   'Obteniendo Tipo de Cambio del día
   r_dbl_TipCam = moddat_gf_ObtieneTipCamDia(2, 2, Format(CDate(date), "yyyymmdd"), 2)
     
   'Obteniendo el Número de Asiento
   r_int_NumAsi = modprc_ff_NumAsi(r_arr_LogPro, CInt(moddat_g_str_CodAno), CInt(moddat_g_str_CodMes), r_str_Origen, r_int_NumLib)
   r_str_AsiGen = CStr(r_int_NumAsi)
   r_str_FecCon = CDate(ipp_FecEmi.Text)
   r_str_FecReg = moddat_g_str_FecSis
   
   r_str_Glosa = "GARANTIA - CF"

   'Insertar en CABECERA
   Call modprc_fs_Inserta_CabAsi(r_arr_LogPro, r_str_Origen, CInt(moddat_g_str_CodAno), CInt(moddat_g_str_CodMes), r_int_NumLib, r_int_NumAsi, Format(1, "000"), r_dbl_TipCam, r_str_TipNot, Trim(r_str_Glosa), r_str_FecCon, "1")

   '*************************************************
   'GENERACION DE ASIENTOS CONTABLES DE GARANTIA
   '*************************************************
   For r_int_ConAux = 1 To 2
      
       r_dbl_importe = p_Importe

       If r_int_ConAux = 1 Then r_str_DebHab = "D": r_str_CtaCtb = p_CtaDeb Else r_str_DebHab = "H": r_str_CtaCtb = p_CtaHab
       
       r_str_Glosa = "CF" & Trim(p_NumRef) & "/" & Trim(p_NumDoc) & "/" & Trim(p_RazSoc)
       r_str_Glosa = Trim(Mid(r_str_Glosa, 1, 60))
       
       If (r_dbl_importe > 0) Then
           r_int_NumIte = r_int_NumIte + 1
           r_dbl_MtoSol = Format(r_dbl_importe, "###,###,##0.00")
           r_dbl_MtoDol = Format(0, "###,###,##0.00")
           
           Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, CInt(moddat_g_str_CodAno), CInt(moddat_g_str_CodMes), r_int_NumLib, r_int_NumAsi, r_int_NumIte, r_str_CtaCtb, CDate(r_str_FecCon), r_str_Glosa, r_str_DebHab, r_dbl_MtoSol, r_dbl_MtoDol, 1, CDate(r_str_FecCon))
           r_dbl_importe = 0
       End If
   Next r_int_ConAux
   
   r_str_NroCnt = r_str_Origen & "/" & moddat_g_str_CodAno & "/" & Format(moddat_g_str_CodMes, "00") & "/" & r_int_NumLib & "/" & r_int_NumAsi
   
   'Grabando en TPR_MAERDE, año/mes/nro_libro/nro_asiento
   modprc_g_str_CadEje = ""
   modprc_g_str_CadEje = modprc_g_str_CadEje & "UPDATE TPR_MAEGAR "
   modprc_g_str_CadEje = modprc_g_str_CadEje & "   SET MAEGAR_NROCNT = '" & CStr(r_str_NroCnt) & "'"
   modprc_g_str_CadEje = modprc_g_str_CadEje & " WHERE MAEGAR_NUMITE = " & p_NumIte & ""
   modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND MAEGAR_TIPDOC = " & p_TipDoc & ""
   modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND MAEGAR_NUMDOC = '" & p_NumDoc & "'"
   
   If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Grabar, 2) Then
      Exit Sub
   End If
   
End Sub
Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call gs_CentraForm(Me)
   Call fs_Inicia
   Call fs_Limpia
   Call fs_Buscar
'   Call fs_Buscar_DatEva
   Call fs_Activa(False)
     
   Screen.MousePointer = 0
End Sub
Private Sub fs_Inicia()
   'empresa tasación
   Call fs_Buscar_EmpPer(cmb_EmpPer)
  ' Call fs_Buscar_NumInf(cmb_NumInf, cmb_FecEva)

   'Moneda
   Call moddat_gs_Carga_LisIte_Combo(cmb_Moneda, 1, "204")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Moneda_Inm, 1, "204")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Moneda_Es1, 1, "204")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Moneda_Es2, 1, "204")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Moneda_Dep1, 1, "204")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Moneda_Dep2, 1, "204")
   
   'Tipo de Garantía
   Call moddat_gs_Carga_LisIte_Combo(Cmb_TipGar, 1, "527")
   
'   'Año y mes
   Call moddat_gf_ConsultaPerMesActivo("000001", 1, moddat_g_str_FecIni, moddat_g_str_FecFin, moddat_g_str_CodMes, moddat_g_str_CodAno)
   
   'Referencias
   Call fs_ConsultaNumRef(moddat_g_int_TipDoc, moddat_g_str_NumDoc)
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_SedReg, 1, "511")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipDoc_Inm, 1, "026")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipDoc_Es1, 1, "026")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipDoc_Es2, 1, "026")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipDoc_Dep1, 1, "026")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipDoc_Dep2, 1, "026")
   Call moddat_gs_Carga_LisIte_Combo(cmb_FlgEst_Es1, 1, "214")
   Call moddat_gs_Carga_LisIte_Combo(cmb_FlgEst_Es2, 1, "214")
   Call moddat_gs_Carga_LisIte_Combo(cmb_FlgEst_Dep1, 1, "214")
   Call moddat_gs_Carga_LisIte_Combo(cmb_FlgEst_Dep2, 1, "214")

   pnl_TipDoc.Caption = moddat_gf_Consulta_ParDes("118", moddat_g_int_TipDoc)
   pnl_NroDoc.Caption = moddat_g_str_NumDoc
   pnl_RazSoc.Caption = moddat_g_str_NomCli
   pnl_TipEmp.Caption = moddat_g_str_Descri
   'pnl_NumRef.Caption = fs_Formato_NumRef(moddat_g_str_NumFia)
   
   
   'Inicializando Rejilla
   grd_Listad.ColWidth(0) = 2550          ' TIPO DE GARANTIA
   grd_Listad.ColWidth(1) = 1410          ' FECHA DE INSCRIPCION O EMISION
   grd_Listad.ColWidth(2) = 2200          ' MONEDA
   grd_Listad.ColWidth(3) = 1830          ' MONTO GARANTIA
   grd_Listad.ColWidth(4) = 4150          ' CARTAS FIANZA ASOCIADAS
   grd_Listad.ColWidth(5) = 0             ' NUMITE DE GARANTIAS
   grd_Listad.ColWidth(6) = 1830          ' MONTO DE CARTAS FIANZAS
   grd_Listad.ColWidth(7) = 0             ' CODIGO DE TASACION
   
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignCenterCenter
   grd_Listad.ColAlignment(3) = flexAlignRightCenter
   grd_Listad.ColAlignment(4) = flexAlignLeftCenter
   grd_Listad.ColAlignment(6) = flexAlignRightCenter
      
   ipp_FecPre_Inm.Text = Format(date, "dd/mm/yyyy")
   txt_NumPre_Inm.Text = ""
   ipp_FecIns_Inm.Text = Format(date, "dd/mm/yyyy")
   txt_NumPar_Inm.Text = ""
   txt_NumAPa_Inm.Text = ""
   txt_NumFic_Inm.Text = ""
   txt_NumAFi_Inm.Text = ""
   ipp_MtoHip_Inm.Value = 0
   
   ipp_FecPre_Es1.Text = Format(date, "dd/mm/yyyy")
   txt_NumPre_Es1.Text = ""
   ipp_FecIns_Es1.Text = Format(date, "dd/mm/yyyy")
   cmb_TipDoc_Es1.ListIndex = -1
   txt_NumPar_Es1.Text = ""
   txt_NumAPa_Es1.Text = ""
   txt_NumFic_Es1.Text = ""
   txt_NumAFi_Es1.Text = ""
   ipp_MtoHip_Es1.Value = 0
   
   ipp_FecPre_Es2.Text = Format(date, "dd/mm/yyyy")
   txt_NumPre_Es2.Text = ""
   ipp_FecIns_Es2.Text = Format(date, "dd/mm/yyyy")
   cmb_TipDoc_Es2.ListIndex = -1
   txt_NumPar_Es2.Text = ""
   txt_NumAPa_Es2.Text = ""
   txt_NumFic_Es2.Text = ""
   txt_NumAFi_Es2.Text = ""
   ipp_MtoHip_Es2.Value = 0
   
   ipp_FecPre_Dep1.Text = Format(date, "dd/mm/yyyy")
   txt_NumPre_Dep1.Text = ""
   ipp_FecIns_Dep1.Text = Format(date, "dd/mm/yyyy")
   cmb_TipDoc_Dep1.ListIndex = -1
   txt_NumPar_Dep1.Text = ""
   txt_NumAPa_Dep1.Text = ""
   txt_NumFic_Dep1.Text = ""
   txt_NumAFi_Dep1.Text = ""
   ipp_MtoHip_Dep1.Value = 0
'   ipp_ValRea_Dep1.Value = 0
   
   ipp_FecPre_Dep2.Text = Format(date, "dd/mm/yyyy")
   txt_NumPre_Dep2.Text = ""
   ipp_FecIns_Dep2.Text = Format(date, "dd/mm/yyyy")
   cmb_TipDoc_Dep2.ListIndex = -1
   txt_NumPar_Dep2.Text = ""
   txt_NumAPa_Dep2.Text = ""
   txt_NumFic_Dep2.Text = ""
   txt_NumAFi_Dep2.Text = ""
   ipp_MtoHip_Dep2.Value = 0

   pnl_TotHip.Caption = "0.00 "
   
End Sub
Private Sub fs_ConsultaNumRef(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String)

Dim r_str_Cadena  As String
Dim r_str_CadAux  As String
Dim r_str_CadRef  As String

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "   SELECT MAEGAR_NUMREF AS NUMREF FROM TPR_MAEGAR "
   g_str_Parame = g_str_Parame & "    WHERE MAEGAR_TIPDOC = " & CStr(p_TipDoc) & " "
   g_str_Parame = g_str_Parame & "      AND MAEGAR_NUMDOC = '" & CStr(p_NumDoc) & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
     Exit Sub
   End If
   
   If g_rst_GenAux.BOF And g_rst_GenAux.EOF Then
      g_rst_GenAux.Close
      Set g_rst_GenAux = Nothing
      GoTo Ingresar
   End If
   
   If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
      g_rst_GenAux.MoveFirst
      r_str_Cadena = Trim(g_rst_GenAux!NUMREF)
      
      'Obtiene los números de referencias actuales
      r_str_CadAux = Trim(r_str_Cadena)
      If Len(r_str_CadAux) > 1 Then
         While InStr(r_str_CadAux, "|")
            r_str_CadRef = Trim(Mid(r_str_CadAux, 1, InStr(r_str_CadAux, "|") - 1))
            r_str_CadRef = fs_Obtener_NumRef(r_str_CadRef)
            r_str_CadAux = Trim(Mid(r_str_CadAux, InStr(r_str_CadAux, "|") + 1))
         Wend
         
         r_str_CadRef = r_str_CadRef & "|" & fs_Obtener_NumRef(r_str_CadAux)
         If InStr(r_str_CadRef, "|") = 1 Then
            r_str_CadRef = Replace(r_str_CadRef, "|", "")
         End If
         r_str_CadRef = Replace(r_str_CadRef, "|", "' , '")
      End If
   End If
   g_rst_GenAux.Close
   Set g_rst_GenAux = Nothing
      
     
Ingresar:

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT MAECFI_NUMREF, MAECFI_NUMANT"
   g_str_Parame = g_str_Parame & "   FROM TPR_MAECFI "
   g_str_Parame = g_str_Parame & "  WHERE MAECFI_TIPDOC = " & CStr(p_TipDoc) & " "
   g_str_Parame = g_str_Parame & "    AND MAECFI_NUMDOC = '" & CStr(p_NumDoc) & "' "
   
'   If r_str_CadRef <> "" Then
'      g_str_Parame = g_str_Parame & "    AND MAECFI_NUMREF NOT IN ('" & r_str_CadRef & "')" 'r_str_Cadena
'   End If
   
   g_str_Parame = g_str_Parame & "    AND MAECFI_SITUAC = 1 "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   cmb_NumRef.Clear
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
     g_rst_Princi.Close
     Set g_rst_Princi = Nothing
     Exit Sub
   End If
     
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         If Not IsNull(g_rst_Princi!MAECFI_NUMANT) Then
            cmb_NumRef.AddItem gf_Formato_NumRef(g_rst_Princi!MAECFI_NUMANT, Mid(g_rst_Princi!MAECFI_NUMANT, 1, 1))
         Else
            If Mid(g_rst_Princi!MAECFI_NUMREF, 1, 3) = "008" Then
               cmb_NumRef.AddItem gf_Formato_NumRef(g_rst_Princi!MAECFI_NUMREF, 1)
            Else
               cmb_NumRef.AddItem gf_Formato_NumRef(g_rst_Princi!MAECFI_NUMREF, Mid(g_rst_Princi!MAECFI_NUMREF, 1, 1))
            End If
         End If
         g_rst_Princi.MoveNext
      Loop
   End If
End Sub
Private Function fs_Obtener_NumRef(ByVal p_NumRef As String) As String ', ByVal p_Tipo As Integer
   fs_Obtener_NumRef = ""
   p_NumRef = Format(p_NumRef, "0000000000")
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT MAECFI_NUMREF, MAECFI_NUMANT "
   g_str_Parame = g_str_Parame & "   FROM TPR_MAECFI "
   g_str_Parame = g_str_Parame & "  WHERE MAECFI_SITUAC = 1 "
   
   'If p_Tipo = 0 Then
      g_str_Parame = g_str_Parame & "     AND MAECFI_NUMANT = '" & p_NumRef & "'"
   'End If
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
       Exit Function
   End If
   
   If g_rst_Listas.BOF And g_rst_Listas.EOF Then
     g_rst_Listas.Close
     Set g_rst_Listas = Nothing
     fs_Obtener_NumRef = p_NumRef
     Exit Function
   End If
   
   fs_Obtener_NumRef = g_rst_Listas!MAECFI_NUMREF
End Function

'Private Function fs_Formato_NumRef(ByVal p_Numref As String) As String
'   p_Numref = Format(p_Numref, "0000000000")
'   'fs_Formato_NumRef = Left(p_Numref, 4) & "-" & Mid(p_Numref, 5, 2) & "-" & Right(p_Numref, 4)
'   fs_Formato_NumRef = Mid(p_Numref, 1, 1) & Mid(p_Numref, 2, 2) & "-" & Mid(p_Numref, 4, 2) & "-" & Right(p_Numref, 5)
'End Function
Private Sub fs_Limpia()
 
   cmb_Moneda.ListIndex = -1
   Cmb_TipGar.ListIndex = -1
   ipp_ImpGar.Value = Format(0, "###,###,###,##0.00")
   ipp_FecEmi.Text = Format(date, "dd/mm/yyyy")
   cmb_NumRef.ListIndex = -1
   pnl_NumCFi.Caption = ""
   
   ipp_FecPre_Inm.Text = Format(date, "dd/mm/yyyy")
   txt_NumPre_Inm.Text = ""
   ipp_FecIns_Inm.Text = Format(date, "dd/mm/yyyy")
   cmb_TipDoc_Inm.ListIndex = -1
   txt_NumPar_Inm.Text = ""
   txt_NumAPa_Inm.Text = ""
   txt_NumFic_Inm.Text = ""
   txt_NumAFi_Inm.Text = ""
   ipp_MtoHip_Inm.Value = 0
   cmb_Moneda_Inm.ListIndex = -1
   
   cmb_FlgEst_Es1.ListIndex = -1
   ipp_FecPre_Es1.Text = Format(date, "dd/mm/yyyy")
   txt_NumPre_Es1.Text = ""
   ipp_FecIns_Es1.Text = Format(date, "dd/mm/yyyy")
   cmb_TipDoc_Es1.ListIndex = -1
   txt_NumPar_Es1.Text = ""
   txt_NumAPa_Es1.Text = ""
   txt_NumFic_Es1.Text = ""
   txt_NumAFi_Es1.Text = ""
   ipp_MtoHip_Es1.Value = 0
   cmb_Moneda_Es1.ListIndex = -1
   
   cmb_FlgEst_Es2.ListIndex = -1
   ipp_FecPre_Es2.Text = Format(date, "dd/mm/yyyy")
   txt_NumPre_Es2.Text = ""
   ipp_FecIns_Es2.Text = Format(date, "dd/mm/yyyy")
   cmb_TipDoc_Es2.ListIndex = -1
   txt_NumPar_Es2.Text = ""
   txt_NumAPa_Es2.Text = ""
   txt_NumFic_Es2.Text = ""
   txt_NumAFi_Es2.Text = ""
   ipp_MtoHip_Es2.Value = 0
   cmb_Moneda_Es2.ListIndex = -1
   
   cmb_FlgEst_Dep1.ListIndex = -1
   ipp_FecPre_Dep1.Text = Format(date, "dd/mm/yyyy")
   txt_NumPre_Dep1.Text = ""
   ipp_FecIns_Dep1.Text = Format(date, "dd/mm/yyyy")
   cmb_TipDoc_Dep1.ListIndex = -1
   txt_NumPar_Dep1.Text = ""
   txt_NumAPa_Dep1.Text = ""
   txt_NumFic_Dep1.Text = ""
   txt_NumAFi_Dep1.Text = ""
   ipp_MtoHip_Dep1.Value = 0
   cmb_Moneda_Dep1.ListIndex = -1
   
   cmb_FlgEst_Dep2.ListIndex = -1
   ipp_FecPre_Dep2.Text = Format(date, "dd/mm/yyyy")
   txt_NumPre_Dep2.Text = ""
   ipp_FecIns_Dep2.Text = Format(date, "dd/mm/yyyy")
   cmb_TipDoc_Dep2.ListIndex = -1
   txt_NumPar_Dep2.Text = ""
   txt_NumAPa_Dep2.Text = ""
   txt_NumFic_Dep2.Text = ""
   txt_NumAFi_Dep2.Text = ""
   ipp_MtoHip_Dep2.Value = 0
   cmb_Moneda_Dep2.ListIndex = -1
   
   cmb_SedReg.ListIndex = -1
   cmb_EmpPer.ListIndex = -1
   cmb_NumInf.ListIndex = -1
   cmb_FecEva.ListIndex = -1
   pnl_TotHip.Caption = "0.00 "
   
   moddat_g_str_DesObs = ""
End Sub
Private Sub fs_Buscar_DatEva()

'   moddat_g_int_FlgGrb = 1
   'Datos de la tasación
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT * FROM TPR_EVATAS "
   g_str_Parame = g_str_Parame & "  WHERE EVATAS_TIPDOC =  " & moddat_g_int_TipDoc & ""
   g_str_Parame = g_str_Parame & "    AND EVATAS_NUMDOC =  '" & CStr(moddat_g_str_NumDoc) & "'"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      cmb_EmpPer.ListIndex = gf_Busca_Arregl(l_arr_EmpPer, g_rst_Princi!EVATAS_CODEMP) - 1
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

End Sub
Private Sub fs_Buscar_EmpPer(p_Combo As ComboBox)

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT EVATAS_CODEMP, TRIM(PARDES_DESCRI) AS EMPPER FROM TPR_EVATAS "
   g_str_Parame = g_str_Parame & "  INNER JOIN MNT_PARDES ON PARDES_CODGRP = 507 AND PARDES_CODITE = EVATAS_CODEMP"
   g_str_Parame = g_str_Parame & "  WHERE EVATAS_TIPDOC =  " & moddat_g_int_TipDoc & ""
   g_str_Parame = g_str_Parame & "    AND EVATAS_NUMDOC =  '" & CStr(moddat_g_str_NumDoc) & "'"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      p_Combo.AddItem Trim$(g_rst_Genera!EMPPER)
      p_Combo.ItemData(p_Combo.NewIndex) = CLng(g_rst_Genera!EVATAS_CODEMP)
   End If
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing

End Sub
Private Sub fs_Buscar_NumInf_FecTas(p_Combo As ComboBox, p_Combo1 As ComboBox, p_CodEmp As String, p_Codigo As Integer)

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT EVATAS_NUMINF, EVATAS_FECEVA FROM TPR_EVATAS "
   g_str_Parame = g_str_Parame & "  WHERE EVATAS_TIPDOC =  " & moddat_g_int_TipDoc & ""
   g_str_Parame = g_str_Parame & "    AND EVATAS_NUMDOC =  '" & CStr(moddat_g_str_NumDoc) & "'"
   g_str_Parame = g_str_Parame & "    AND EVATAS_CODEMP = '" & p_CodEmp & "'"
   
   If p_Codigo = 1 Then
      g_str_Parame = g_str_Parame & "    AND EVATAS_NUMINF = '" & CStr(cmb_NumInf.Text) & "'"
   End If
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
      g_rst_Listas.MoveFirst
      Do While Not g_rst_Listas.EOF
         If p_Codigo = 0 Then
            p_Combo.AddItem Trim$(g_rst_Listas!EVATAS_NUMINF)
         ElseIf p_Codigo = 1 Then
            p_Combo1.AddItem Format(gf_FormatoFecha(CStr(g_rst_Listas!EVATAS_FECEVA)), "dd/mm/yyyy")
            p_Combo1.ItemData(p_Combo1.NewIndex) = CLng(g_rst_Listas!EVATAS_FECEVA)
         End If
         g_rst_Listas.MoveNext
      Loop
   End If
   
   g_rst_Listas.Close
   Set g_rst_Listas = Nothing
End Sub
Private Sub fs_Buscar()

Dim r_str_Cadena  As String
Dim r_str_CadAux  As String
Dim r_str_CadRef  As String
Dim r_dbl_MtoCfi  As Double

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT MAEGAR_NUMITE     , MAEGAR_TIPDOC    , MAEGAR_NUMDOC    , MAEGAR_CODTAS    , MAEGAR_TIPGAR    , MAEGAR_FECPRE_INM    , TRIM(MAEGAR_NUMREF) NUMREF, MAEGAR_TIPMON_INM, "
   g_str_Parame = g_str_Parame & "        (NVL(MAEGAR_MTOGAR_INM,0) + NVL(MAEGAR_MTOGAR_ES1,0) + NVL(MAEGAR_MTOGAR_ES2,0) + NVL(MAEGAR_MTOGAR_DE1,0) + NVL(MAEGAR_MTOGAR_DE2,0)) AS MTOGAR "
   g_str_Parame = g_str_Parame & "   FROM TPR_MAEGAR "
   g_str_Parame = g_str_Parame & "  WHERE MAEGAR_TIPDOC = " & CStr(moddat_g_int_TipDoc) & ""
   g_str_Parame = g_str_Parame & "    AND MAEGAR_NUMDOC = '" & CStr(moddat_g_str_NumDoc) & "' "
   g_str_Parame = g_str_Parame & "    AND MAEGAR_SITUAC = 1 "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   grd_Listad.Redraw = False
   
   Call gs_LimpiaGrid(grd_Listad)
    
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
     g_rst_Princi.Close
     Set g_rst_Princi = Nothing
     grd_Listad.Redraw = True
     Exit Sub
   End If
     
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
     
      Do While Not g_rst_Princi.EOF
         
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         
         r_str_Cadena = IIf(IsNull(Trim(g_rst_Princi!NUMREF)), "", Trim(g_rst_Princi!NUMREF))
         
         r_str_CadAux = Trim(r_str_Cadena)
         
         While InStr(r_str_CadAux, "|") > 0
      
            r_str_CadRef = Trim(Mid(r_str_CadAux, 1, InStr(r_str_CadAux, "|") - 1))
            r_str_CadRef = fs_Obtener_NumRef(r_str_CadRef)
            
            g_str_Parame = ""
            g_str_Parame = g_str_Parame & "   SELECT SUM(B.MAECFI_IMPFIA) MONTO_CFI "
            g_str_Parame = g_str_Parame & "     FROM TPR_MAECFI B "
            g_str_Parame = g_str_Parame & "    WHERE MAECFI_NUMREF = '" & r_str_CadRef & "' "
         
            If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
               Exit Sub
            End If
            
            If g_rst_Genera.BOF And g_rst_Genera.EOF Then
              g_rst_Genera.Close
              Set g_rst_Genera = Nothing
              Exit Sub
            End If
            
            r_dbl_MtoCfi = r_dbl_MtoCfi + IIf(IsNull(g_rst_Genera!MONTO_CFI), 0, g_rst_Genera!MONTO_CFI)
            'r_str_Cadena = Trim(Mid(r_str_Cadena, InStr(r_str_Cadena, r_str_CadAux) + Len(r_str_CadAux) + 1))
            r_str_CadAux = Trim(Mid(r_str_CadAux, InStr(r_str_CadAux, "|") + 1))
            r_str_CadRef = ""
         Wend
      
         If r_str_CadAux <> "" Then
            r_str_CadRef = fs_Obtener_NumRef(r_str_CadAux)
         
            g_str_Parame = ""
            g_str_Parame = g_str_Parame & "   SELECT SUM(B.MAECFI_IMPFIA) MONTO_CFI "
            g_str_Parame = g_str_Parame & "     FROM TPR_MAECFI B "
            g_str_Parame = g_str_Parame & "    WHERE MAECFI_NUMREF = '" & r_str_CadRef & "' "
         
            If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
               Exit Sub
            End If
            
            If g_rst_Genera.BOF And g_rst_Genera.EOF Then
              g_rst_Genera.Close
              Set g_rst_Genera = Nothing
              Exit Sub
            End If
            
            r_dbl_MtoCfi = r_dbl_MtoCfi + IIf(IsNull(g_rst_Genera!MONTO_CFI), 0, g_rst_Genera!MONTO_CFI)
         End If
         
         grd_Listad.Col = 0
         grd_Listad.Text = moddat_gf_Consulta_ParDes("527", g_rst_Princi!MAEGAR_TIPGAR)
                 
         grd_Listad.Col = 1
         If Not IsNull(g_rst_Princi!MAEGAR_FECPRE_INM) Then
            grd_Listad.Text = Format(gf_FormatoFecha(CStr(g_rst_Princi!MAEGAR_FECPRE_INM)), "dd/mm/yyyy")
         Else
            grd_Listad.Text = ""
         End If
         grd_Listad.Col = 2
         grd_Listad.Text = moddat_gf_Consulta_ParDes("204", g_rst_Princi!MAEGAR_TIPMON_INM)
         
         grd_Listad.Col = 3
         grd_Listad.Text = Format(CStr(g_rst_Princi!MTOGAR), "###,###,###,##0.00") ' MAEGAR_MTOGAR_INM
         
         grd_Listad.Col = 4
         If IsNull(Trim(g_rst_Princi!NUMREF)) Then
            grd_Listad.Text = ""
         Else
            grd_Listad.Text = CStr(Trim(g_rst_Princi!NUMREF)) 'fs_Formato_NumRef
         End If
         
         grd_Listad.Col = 5
         grd_Listad.Text = CStr(g_rst_Princi!MAEGAR_NUMITE)
         
         grd_Listad.Col = 6
         If IsNull(Trim(r_dbl_MtoCfi)) Then
            grd_Listad.Text = ""
         Else
            grd_Listad.Text = Format(CStr(r_dbl_MtoCfi), "###,###,###,##0.00")
         End If
         r_dbl_MtoCfi = 0
         
                  
         grd_Listad.Col = 7
         If Not IsNull(g_rst_Princi!MAEGAR_CODTAS) Then
            grd_Listad.Text = CStr(g_rst_Princi!MAEGAR_CODTAS) 'MAEGAR_SEDREG
         Else
            grd_Listad.Text = ""
         End If
         
         g_rst_Princi.MoveNext
      Loop
   End If
   
   grd_Listad.Redraw = True
   Call gs_UbiIniGrid(grd_Listad)
   
End Sub
Private Sub fs_Activa(ByVal p_Activa As Integer)
 
   cmb_NumRef.Enabled = p_Activa
   cmb_Moneda.Enabled = p_Activa
   Cmb_TipGar.Enabled = p_Activa
   ipp_ImpGar.Enabled = p_Activa
   ipp_FecEmi.Enabled = p_Activa
   cmd_CfiAso.Enabled = p_Activa
   pnl_NumCFi.Enabled = p_Activa
   
   cmb_SedReg.Enabled = p_Activa
   cmb_EmpPer.Enabled = p_Activa
   cmb_NumInf.Enabled = p_Activa
   cmb_FecEva.Enabled = p_Activa
   
   ipp_FecPre_Inm.Enabled = p_Activa
   txt_NumPre_Inm.Enabled = p_Activa
   ipp_FecIns_Inm.Enabled = p_Activa
   cmb_TipDoc_Inm.Enabled = p_Activa
   txt_NumPar_Inm.Enabled = p_Activa
   txt_NumAPa_Inm.Enabled = p_Activa
   txt_NumFic_Inm.Enabled = p_Activa
   txt_NumAFi_Inm.Enabled = p_Activa
   ipp_MtoHip_Inm.Enabled = p_Activa
   cmb_Moneda_Inm.Enabled = p_Activa
   
   cmb_FlgEst_Es1.Enabled = p_Activa
   ipp_FecPre_Es1.Enabled = p_Activa
   txt_NumPre_Es1.Enabled = p_Activa
   ipp_FecIns_Es1.Enabled = p_Activa
   cmb_TipDoc_Es1.Enabled = p_Activa
   txt_NumPar_Es1.Enabled = p_Activa
   txt_NumAPa_Es1.Enabled = p_Activa
   txt_NumFic_Es1.Enabled = p_Activa
   txt_NumAFi_Es1.Enabled = p_Activa
   ipp_MtoHip_Es1.Enabled = p_Activa
   cmb_Moneda_Es1.Enabled = p_Activa
   
   cmb_FlgEst_Es2.Enabled = p_Activa
   ipp_FecPre_Es2.Enabled = p_Activa
   txt_NumPre_Es2.Enabled = p_Activa
   ipp_FecIns_Es2.Enabled = p_Activa
   cmb_TipDoc_Es2.Enabled = p_Activa
   txt_NumPar_Es2.Enabled = p_Activa
   txt_NumAPa_Es2.Enabled = p_Activa
   txt_NumFic_Es2.Enabled = p_Activa
   txt_NumAFi_Es2.Enabled = p_Activa
   ipp_MtoHip_Es2.Enabled = p_Activa
   cmb_Moneda_Es2.Enabled = p_Activa
   
   cmb_FlgEst_Dep1.Enabled = p_Activa
   ipp_FecPre_Dep1.Enabled = p_Activa
   txt_NumPre_Dep1.Enabled = p_Activa
   ipp_FecIns_Dep1.Enabled = p_Activa
   cmb_TipDoc_Dep1.Enabled = p_Activa
   txt_NumPar_Dep1.Enabled = p_Activa
   txt_NumAPa_Dep1.Enabled = p_Activa
   txt_NumFic_Dep1.Enabled = p_Activa
   txt_NumAFi_Dep1.Enabled = p_Activa
   ipp_MtoHip_Dep1.Enabled = p_Activa
   cmb_Moneda_Dep1.Enabled = p_Activa

   cmb_FlgEst_Dep2.Enabled = p_Activa
   ipp_FecPre_Dep2.Enabled = p_Activa
   txt_NumPre_Dep2.Enabled = p_Activa
   ipp_FecIns_Dep2.Enabled = p_Activa
   cmb_TipDoc_Dep2.Enabled = p_Activa
   txt_NumPar_Dep2.Enabled = p_Activa
   txt_NumAPa_Dep2.Enabled = p_Activa
   txt_NumFic_Dep2.Enabled = p_Activa
   txt_NumAFi_Dep2.Enabled = p_Activa
   ipp_MtoHip_Dep2.Enabled = p_Activa
   cmb_Moneda_Dep2.Enabled = p_Activa
      
   cmd_Agrega.Enabled = Not p_Activa
   
   If Me.grd_Listad.Row < 0 Then
      cmd_Editar.Enabled = p_Activa
      cmd_Borrar.Enabled = p_Activa
      cmd_ExpExc.Enabled = p_Activa
   Else
      cmd_Editar.Enabled = Not p_Activa
      cmd_Borrar.Enabled = Not p_Activa
      cmd_ExpExc.Enabled = Not p_Activa
   End If
   cmd_Grabar.Enabled = p_Activa
   cmd_Cancel.Enabled = p_Activa
   
End Sub

Private Sub grd_Listad_Click()
   Call fs_Limpia
   Call cmd_Editar_Click
   Call fs_Activa(False)
End Sub

Private Sub ipp_FecEmi_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Moneda)
   End If
End Sub

Private Sub ipp_FecIns_Dep1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipDoc_Dep1)
   End If
End Sub

Private Sub ipp_FecIns_Dep2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipDoc_Dep2)
   End If
End Sub

Private Sub ipp_FecIns_Es1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipDoc_Es1)
   End If
End Sub

Private Sub ipp_FecIns_Es2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipDoc_Es2)
   End If
End Sub

Private Sub ipp_FecIns_Inm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipDoc_Inm)
   End If
End Sub

Private Sub ipp_FecPre_Dep1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NumPre_Dep1)
   End If
End Sub

Private Sub ipp_FecPre_Dep2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NumPre_Dep2)
   End If
End Sub

Private Sub ipp_FecPre_Es1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NumPre_Es1)
   End If
End Sub

Private Sub ipp_FecPre_Es2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NumPre_Es2)
   End If
End Sub

Private Sub ipp_FecPre_Inm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NumPre_Inm)
   End If
End Sub

Private Sub ipp_ImpGar_Change()
   pnl_TotHip.Caption = Format(CDbl(ipp_ImpGar.Text), "###,###,##0.00") & " "
End Sub

Private Sub ipp_ImpGar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
'      If cmb_NumRef.ListIndex = -1 Then
'         Call gs_SetFocus(cmd_Grabar)
'      Else
'         Call gs_SetFocus(cmb_NumRef)
'      End If
      Call gs_SetFocus(cmb_NumRef)
   End If
End Sub

'Private Sub Txt_NumRef_KeyPress(KeyAscii As Integer)
'   If KeyAscii = 13 Then
'      Call gs_SetFocus(ipp_FecEmi)
'   End If
'End Sub

'Private Sub txt_NumOpe_KeyPress(KeyAscii As Integer)
'   If KeyAscii = 13 Then
'      Call gs_SetFocus(ipp_FecEmi)
'   End If
'End Sub
Private Sub fs_CalcularTotHip()
   pnl_TotHip.Caption = Format(CDbl(ipp_MtoHip_Inm.Text) + CDbl(ipp_MtoHip_Es1.Text) + CDbl(ipp_MtoHip_Dep1.Text), "###,###,###,##0.00")
End Sub

Private Sub ipp_MtoHip_Dep1_Change()
  Call ipp_MtoHip_Inm_Change
End Sub

Private Sub ipp_MtoHip_Dep1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Moneda_Dep1)
   End If
End Sub

Private Sub ipp_MtoHip_Dep2_Change()
   Call ipp_MtoHip_Inm_Change
End Sub

Private Sub ipp_MtoHip_Dep2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Moneda_Dep2)
   End If
End Sub

Private Sub ipp_MtoHip_Es1_Change()
   Call ipp_MtoHip_Inm_Change
End Sub

Private Sub ipp_MtoHip_Es1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Moneda_Es1)
   End If
End Sub

Private Sub ipp_MtoHip_Es2_Change()
   Call ipp_MtoHip_Inm_Change
End Sub

Private Sub ipp_MtoHip_Es2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Moneda_Es2)
   End If
End Sub

Private Sub ipp_MtoHip_Inm_Change()
   pnl_TotHip.Caption = Format(CDbl(ipp_MtoHip_Inm.Text) + CDbl(ipp_MtoHip_Es1.Text) + CDbl(ipp_MtoHip_Es2.Text) + CDbl(ipp_MtoHip_Dep1.Text) + CDbl(ipp_MtoHip_Dep2.Text), "###,###,##0.00") & " "
End Sub

Private Sub ipp_MtoHip_Inm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Moneda_Inm)
   End If
End Sub

'Private Sub ipp_ValRea_Dep1_KeyPress(KeyAscii As Integer)
'   If KeyAscii = 13 Then
'      Call gs_SetFocus(cmb_Moneda_Dep1)
'   End If
'End Sub

'Private Sub ipp_ValRea_Es1_KeyPress(KeyAscii As Integer)
'   If KeyAscii = 13 Then
'      Call gs_SetFocus(cmb_Moneda_Es1)
'   End If
'End Sub

'Private Sub ipp_ValRea_Inm_KeyPress(KeyAscii As Integer)
'   If KeyAscii = 13 Then
'      Call gs_SetFocus(cmb_Moneda_Inm)
'   End If
'End Sub

Private Sub txt_NumAFi_Dep1_GotFocus()
   Call gs_SelecTodo(txt_NumAFi_Dep1)
End Sub

Private Sub txt_NumAFi_Dep1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_MtoHip_Dep1)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_")
   End If
End Sub

Private Sub txt_NumAFi_Dep2_Change()
   Call gs_SelecTodo(txt_NumAFi_Dep2)
End Sub

Private Sub txt_NumAFi_Dep2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_MtoHip_Dep2)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_")
   End If
End Sub

Private Sub txt_NumAFi_Es1_GotFocus()
   Call gs_SelecTodo(txt_NumAFi_Es1)
End Sub

Private Sub txt_NumAFi_Es1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_MtoHip_Es1)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_")
   End If
End Sub

Private Sub txt_NumAFi_Es2_GotFocus()
   Call gs_SelecTodo(txt_NumAFi_Es2)
End Sub

Private Sub txt_NumAFi_Es2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_MtoHip_Es2)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_")
   End If
End Sub

Private Sub txt_NumAFi_Inm_GotFocus()
   Call gs_SelecTodo(txt_NumAFi_Inm)
End Sub

Private Sub txt_NumAFi_Inm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_MtoHip_Inm)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_")
   End If
End Sub

Private Sub txt_NumAPa_Dep1_GotFocus()
   Call gs_SelecTodo(txt_NumAPa_Dep1)
End Sub

Private Sub txt_NumAPa_Dep1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_MtoHip_Dep1)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_")
   End If
End Sub

Private Sub txt_NumAPa_Dep2_GotFocus()
   Call gs_SelecTodo(txt_NumAPa_Dep2)
End Sub

Private Sub txt_NumAPa_Dep2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_MtoHip_Dep2)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_")
   End If
End Sub

Private Sub txt_NumAPa_Es1_GotFocus()
   Call gs_SelecTodo(txt_NumAPa_Es1)
End Sub

Private Sub txt_NumAPa_Es1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_MtoHip_Es1)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_")
   End If
End Sub

Private Sub txt_NumAPa_Es2_GotFocus()
   Call gs_SelecTodo(txt_NumAPa_Es2)
End Sub

Private Sub txt_NumAPa_Es2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_MtoHip_Es2)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_")
   End If
End Sub

Private Sub txt_NumAPa_Inm_GotFocus()
   Call gs_SelecTodo(txt_NumAPa_Inm)
End Sub

Private Sub txt_NumAPa_Inm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_MtoHip_Inm)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_")
   End If
End Sub

Private Sub txt_NumFic_Dep1_GotFocus()
   Call gs_SelecTodo(txt_NumFic_Dep1)
End Sub

Private Sub txt_NumFic_Dep1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NumAFi_Dep1)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_")
   End If
End Sub

Private Sub txt_NumFic_Dep2_GotFocus()
   Call gs_SelecTodo(txt_NumFic_Dep2)
End Sub

Private Sub txt_NumFic_Dep2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NumAFi_Dep2)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_")
   End If
End Sub

Private Sub txt_NumFic_Es1_GotFocus()
   Call gs_SelecTodo(txt_NumFic_Es1)
End Sub

Private Sub txt_NumFic_Es1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NumAFi_Es1)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_")
   End If
End Sub

Private Sub txt_NumFic_Es2_GotFocus()
   Call gs_SelecTodo(txt_NumFic_Es2)
End Sub

Private Sub txt_NumFic_Es2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NumAFi_Es2)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_")
   End If
End Sub

Private Sub txt_NumFic_Inm_GotFocus()
   Call gs_SelecTodo(txt_NumFic_Inm)
End Sub

Private Sub txt_NumFic_Inm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NumAFi_Inm)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_")
   End If
End Sub


Private Sub txt_NumPar_Dep1_GotFocus()
   Call gs_SelecTodo(txt_NumPar_Dep1)
End Sub

Private Sub txt_NumPar_Dep1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NumAPa_Dep1)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_")
   End If
End Sub

Private Sub txt_NumPar_Dep2_GotFocus()
   Call gs_SelecTodo(txt_NumPar_Dep2)
End Sub

Private Sub txt_NumPar_Dep2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NumAPa_Dep2)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_")
   End If
End Sub

Private Sub txt_NumPar_Es1_GotFocus()
   Call gs_SelecTodo(txt_NumPar_Es1)
End Sub

Private Sub txt_NumPar_Es1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NumAPa_Es1)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_")
   End If
End Sub

Private Sub txt_NumPar_Es2_GotFocus()
   Call gs_SelecTodo(txt_NumPar_Es2)
End Sub

Private Sub txt_NumPar_Es2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NumAPa_Es2)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_")
   End If
End Sub

Private Sub txt_NumPar_Inm_GotFocus()
   Call gs_SelecTodo(txt_NumPar_Inm)
End Sub

Private Sub txt_NumPar_Inm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NumAPa_Inm)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_")
   End If
End Sub

Private Sub txt_NumPre_Dep1_GotFocus()
 Call gs_SelecTodo(txt_NumPre_Dep1)
End Sub

Private Sub txt_NumPre_Dep1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecIns_Dep1)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_")
   End If
End Sub
Private Sub txt_NumPre_Dep2_GotFocus()
   Call gs_SelecTodo(txt_NumPre_Dep2)
End Sub
Private Sub txt_NumPre_Dep2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecIns_Dep2)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_")
   End If
End Sub

Private Sub txt_NumPre_Es1_GotFocus()
   Call gs_SelecTodo(txt_NumPre_Es1)
End Sub

Private Sub txt_NumPre_Es1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecIns_Es1)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_")
   End If
End Sub
Private Sub txt_NumPre_Es2_GotFocus()
   Call gs_SelecTodo(txt_NumPre_Es2)
End Sub

Private Sub txt_NumPre_Es2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecIns_Es2)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_")
   End If
End Sub

Private Sub txt_NumPre_Inm_GotFocus()
   Call gs_SelecTodo(txt_NumPre_Inm)
End Sub

Private Sub txt_NumPre_Inm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecIns_Inm)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_")
   End If
End Sub
