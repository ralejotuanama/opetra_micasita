VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm_Ges_TecPro_03 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   9270
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   17130
   Icon            =   "OpeTra_frm_821.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9270
   ScaleWidth      =   17130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   9255
      Left            =   30
      TabIndex        =   10
      Top             =   30
      Width           =   17115
      _Version        =   65536
      _ExtentX        =   30189
      _ExtentY        =   16325
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
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   11
         Top             =   30
         Width           =   17025
         _Version        =   65536
         _ExtentX        =   30030
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
            TabIndex        =   12
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
            TabIndex        =   13
            Top             =   330
            Width           =   4215
            _Version        =   65536
            _ExtentX        =   7435
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Techo Propio - Cartas Fianza y Adendas"
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
            Picture         =   "OpeTra_frm_821.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   1155
         Left            =   30
         TabIndex        =   14
         Top             =   1470
         Width           =   17025
         _Version        =   65536
         _ExtentX        =   30030
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
         Begin VB.ComboBox cmb_EstFia 
            Height          =   315
            Left            =   11550
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   750
            Width           =   2955
         End
         Begin EditLib.fpDateTime ipp_FecIni 
            Height          =   315
            Left            =   11550
            TabIndex        =   0
            Top             =   420
            Width           =   1365
            _Version        =   196608
            _ExtentX        =   2408
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
         Begin EditLib.fpDateTime ipp_FecFin 
            Height          =   315
            Left            =   13140
            TabIndex        =   1
            Top             =   420
            Width           =   1365
            _Version        =   196608
            _ExtentX        =   2408
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
         Begin Threed.SSPanel pnl_RazSoc 
            Height          =   315
            Left            =   1620
            TabIndex        =   15
            Top             =   420
            Width           =   6615
            _Version        =   65536
            _ExtentX        =   11668
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
            TabIndex        =   23
            Top             =   60
            Width           =   6615
            _Version        =   65536
            _ExtentX        =   11668
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
            Height          =   285
            Left            =   11550
            TabIndex        =   24
            Top             =   60
            Width           =   2925
            _Version        =   65536
            _ExtentX        =   5159
            _ExtentY        =   503
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
         Begin Threed.SSPanel pnl_TipEmp 
            Height          =   315
            Left            =   1620
            TabIndex        =   25
            Top             =   780
            Width           =   6615
            _Version        =   65536
            _ExtentX        =   11668
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
         Begin VB.Label Label1 
            Caption         =   "Estado:"
            Height          =   255
            Left            =   9690
            TabIndex        =   22
            Top             =   780
            Width           =   675
         End
         Begin VB.Label lbl_TipDoc 
            Caption         =   "Tipo Documento:"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label Label6 
            Caption         =   "Fecha Emisión:"
            Height          =   285
            Left            =   9690
            TabIndex        =   19
            Top             =   465
            Width           =   1245
         End
         Begin VB.Label lbl_NumDoc 
            Caption         =   "Nro. Documento:"
            Height          =   225
            Left            =   9690
            TabIndex        =   18
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label lbl_RazSoc 
            Caption         =   "Razón Social:"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   465
            Width           =   1335
         End
         Begin VB.Label lbl_TipEmp 
            Caption         =   "Tipo Empresa:"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   780
            Width           =   1335
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   675
         Left            =   30
         TabIndex        =   21
         Top             =   750
         Width           =   17025
         _Version        =   65536
         _ExtentX        =   30030
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
         Begin VB.CommandButton cmd_ExpLiq 
            Height          =   585
            Left            =   6030
            Picture         =   "OpeTra_frm_821.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   30
            ToolTipText     =   "Imprimir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   5430
            Picture         =   "OpeTra_frm_821.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   29
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_DatBen 
            Height          =   585
            Left            =   4830
            Picture         =   "OpeTra_frm_821.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   28
            ToolTipText     =   "Beneficiarios"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Gestion 
            Height          =   585
            Left            =   4230
            Picture         =   "OpeTra_frm_821.frx":0EA4
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Gestionar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Histor 
            Height          =   585
            Left            =   3630
            Picture         =   "OpeTra_frm_821.frx":11AE
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Histórico"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Renova 
            Height          =   585
            Left            =   3030
            Picture         =   "OpeTra_frm_821.frx":15F0
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Renovar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Borrar 
            Height          =   585
            Left            =   2430
            Picture         =   "OpeTra_frm_821.frx":1A32
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Borrar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Editar 
            Height          =   585
            Left            =   1830
            Picture         =   "OpeTra_frm_821.frx":1D3C
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Modificar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Agrega 
            Height          =   585
            Left            =   1230
            Picture         =   "OpeTra_frm_821.frx":2046
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Nuevo"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   630
            Picture         =   "OpeTra_frm_821.frx":2350
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Limpiar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_821.frx":265A
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Buscar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   16410
            Picture         =   "OpeTra_frm_821.frx":2964
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   6525
         Left            =   30
         TabIndex        =   31
         Top             =   2660
         Width           =   17025
         _Version        =   65536
         _ExtentX        =   30030
         _ExtentY        =   11509
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
            Height          =   6465
            Left            =   30
            TabIndex        =   32
            Top             =   60
            Width           =   16995
            _ExtentX        =   29977
            _ExtentY        =   11404
            _Version        =   393216
            Tabs            =   2
            TabsPerRow      =   2
            TabHeight       =   520
            TabCaption(0)   =   "CREDITOS INDIRECTOS"
            TabPicture(0)   =   "OpeTra_frm_821.frx":2DA6
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SSPanel3"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "pnl_Modali"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "pnl_Garant"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "pnl_Valor"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "pnl_ImpDes"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "pnl_ImpFre"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "pnl_SalCom"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "pnl_SalDes"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).Control(8)=   "pnl_DesDes"
            Tab(0).Control(8).Enabled=   0   'False
            Tab(0).Control(9)=   "pnl_PagDes"
            Tab(0).Control(9).Enabled=   0   'False
            Tab(0).Control(10)=   "pnl_DesFre"
            Tab(0).Control(10).Enabled=   0   'False
            Tab(0).Control(11)=   "pnl_DesCom"
            Tab(0).Control(11).Enabled=   0   'False
            Tab(0).Control(12)=   "pnl_EstFia"
            Tab(0).Control(12).Enabled=   0   'False
            Tab(0).Control(13)=   "pnl_SalFre"
            Tab(0).Control(13).Enabled=   0   'False
            Tab(0).Control(14)=   "pnl_RecFre"
            Tab(0).Control(14).Enabled=   0   'False
            Tab(0).Control(15)=   "pnl_PagCom"
            Tab(0).Control(15).Enabled=   0   'False
            Tab(0).Control(16)=   "pnl_NroCar"
            Tab(0).Control(16).Enabled=   0   'False
            Tab(0).Control(17)=   "pnl_ImpCom"
            Tab(0).Control(17).Enabled=   0   'False
            Tab(0).Control(18)=   "pnl_FecVto"
            Tab(0).Control(18).Enabled=   0   'False
            Tab(0).Control(19)=   "pnl_FecEmi"
            Tab(0).Control(19).Enabled=   0   'False
            Tab(0).Control(20)=   "grd_Listad"
            Tab(0).Control(20).Enabled=   0   'False
            Tab(0).ControlCount=   21
            TabCaption(1)   =   "CREDITOS DIRECTOS"
            TabPicture(1)   =   "OpeTra_frm_821.frx":2DC2
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "pnl_ImpCom_Dir"
            Tab(1).Control(1)=   "pnl_DesCom_Dir"
            Tab(1).Control(2)=   "pnl_DesDes_Dir"
            Tab(1).Control(3)=   "pnl_TasMor"
            Tab(1).Control(4)=   "pnl_Modali_Dir"
            Tab(1).Control(5)=   "pnl_TasInt"
            Tab(1).Control(6)=   "pnl_MtoPre"
            Tab(1).Control(7)=   "pnl_ImpDes_Dir"
            Tab(1).Control(8)=   "pnl_SalCom_Dir"
            Tab(1).Control(9)=   "pnl_SalDes_Dir"
            Tab(1).Control(10)=   "pnl_PagDes_Dir"
            Tab(1).Control(11)=   "pnl_Estado"
            Tab(1).Control(12)=   "pnl_PagCom_Dir"
            Tab(1).Control(13)=   "pnl_NroCre"
            Tab(1).Control(14)=   "pnl_FecVto_Dir"
            Tab(1).Control(15)=   "pnl_FecEmi_Dir"
            Tab(1).Control(16)=   "grd_Listad_Dir"
            Tab(1).ControlCount=   17
            Begin MSFlexGridLib.MSFlexGrid grd_Listad 
               Height          =   5565
               Left            =   30
               TabIndex        =   33
               Top             =   870
               Width           =   16905
               _ExtentX        =   29819
               _ExtentY        =   9816
               _Version        =   393216
               Rows            =   30
               Cols            =   22
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel pnl_FecEmi 
               Height          =   555
               Left            =   2280
               TabIndex        =   34
               Top             =   330
               Width           =   1005
               _Version        =   65536
               _ExtentX        =   1773
               _ExtentY        =   979
               _StockProps     =   15
               Caption         =   "  Fecha   Emisión"
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
            Begin Threed.SSPanel pnl_FecVto 
               Height          =   555
               Left            =   3270
               TabIndex        =   35
               Top             =   330
               Width           =   1005
               _Version        =   65536
               _ExtentX        =   1773
               _ExtentY        =   979
               _StockProps     =   15
               Caption         =   "    Fecha      Vcto."
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
            Begin Threed.SSPanel pnl_ImpCom 
               Height          =   285
               Left            =   6300
               TabIndex        =   36
               Top             =   600
               Width           =   1050
               _Version        =   65536
               _ExtentX        =   1852
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
            Begin Threed.SSPanel pnl_NroCar 
               Height          =   555
               Left            =   1170
               TabIndex        =   37
               Top             =   330
               Width           =   1125
               _Version        =   65536
               _ExtentX        =   1984
               _ExtentY        =   979
               _StockProps     =   15
               Caption         =   "Número"
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
            Begin Threed.SSPanel pnl_PagCom 
               Height          =   285
               Left            =   7320
               TabIndex        =   38
               Top             =   600
               Width           =   1050
               _Version        =   65536
               _ExtentX        =   1852
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Pagado"
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
            Begin Threed.SSPanel pnl_RecFre 
               Height          =   285
               Left            =   10380
               TabIndex        =   39
               Top             =   600
               Width           =   1050
               _Version        =   65536
               _ExtentX        =   1852
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Recibido"
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
            Begin Threed.SSPanel pnl_SalFre 
               Height          =   285
               Left            =   11400
               TabIndex        =   40
               Top             =   600
               Width           =   1050
               _Version        =   65536
               _ExtentX        =   1852
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Saldo"
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
            Begin Threed.SSPanel pnl_EstFia 
               Height          =   555
               Left            =   15480
               TabIndex        =   41
               Top             =   330
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
               _ExtentY        =   979
               _StockProps     =   15
               Caption         =   "Estado"
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
            Begin Threed.SSPanel pnl_DesCom 
               Height          =   285
               Left            =   6300
               TabIndex        =   42
               Top             =   330
               Width           =   3090
               _Version        =   65536
               _ExtentX        =   5450
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "COMISIONES"
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
            Begin Threed.SSPanel pnl_DesFre 
               Height          =   285
               Left            =   9360
               TabIndex        =   43
               Top             =   330
               Width           =   3090
               _Version        =   65536
               _ExtentX        =   5450
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "FONDOS RECIBIDOS"
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
            Begin Threed.SSPanel pnl_PagDes 
               Height          =   285
               Left            =   13440
               TabIndex        =   44
               Top             =   600
               Width           =   1050
               _Version        =   65536
               _ExtentX        =   1852
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Pagado"
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
            Begin Threed.SSPanel pnl_DesDes 
               Height          =   285
               Left            =   12420
               TabIndex        =   45
               Top             =   330
               Width           =   3090
               _Version        =   65536
               _ExtentX        =   5450
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "DESEMBOLSOS - ET"
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
            Begin Threed.SSPanel pnl_SalDes 
               Height          =   285
               Left            =   14460
               TabIndex        =   46
               Top             =   600
               Width           =   1050
               _Version        =   65536
               _ExtentX        =   1852
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Saldo"
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
            Begin Threed.SSPanel pnl_SalCom 
               Height          =   285
               Left            =   8340
               TabIndex        =   47
               Top             =   600
               Width           =   1050
               _Version        =   65536
               _ExtentX        =   1852
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Saldo"
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
            Begin Threed.SSPanel pnl_ImpFre 
               Height          =   285
               Left            =   9360
               TabIndex        =   48
               Top             =   600
               Width           =   1050
               _Version        =   65536
               _ExtentX        =   1852
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
            Begin Threed.SSPanel pnl_ImpDes 
               Height          =   285
               Left            =   12420
               TabIndex        =   49
               Top             =   600
               Width           =   1050
               _Version        =   65536
               _ExtentX        =   1852
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
            Begin Threed.SSPanel pnl_Valor 
               Height          =   555
               Left            =   4260
               TabIndex        =   50
               Top             =   330
               Width           =   1020
               _Version        =   65536
               _ExtentX        =   1799
               _ExtentY        =   979
               _StockProps     =   15
               Caption         =   "Valor"
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
            Begin Threed.SSPanel pnl_Garant 
               Height          =   555
               Left            =   5250
               TabIndex        =   51
               Top             =   330
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
               _ExtentY        =   979
               _StockProps     =   15
               Caption         =   "Garantizado"
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
            Begin Threed.SSPanel pnl_Modali 
               Height          =   555
               Left            =   60
               TabIndex        =   52
               Top             =   330
               Width           =   465
               _Version        =   65536
               _ExtentX        =   820
               _ExtentY        =   979
               _StockProps     =   15
               Caption         =   "Tipo"
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
               Height          =   555
               Left            =   510
               TabIndex        =   53
               Top             =   330
               Width           =   675
               _Version        =   65536
               _ExtentX        =   1191
               _ExtentY        =   979
               _StockProps     =   15
               Caption         =   "N° FMV"
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
            Begin MSFlexGridLib.MSFlexGrid grd_Listad_Dir 
               Height          =   5565
               Left            =   -74970
               TabIndex        =   54
               Top             =   870
               Width           =   16905
               _ExtentX        =   29819
               _ExtentY        =   9816
               _Version        =   393216
               Rows            =   30
               Cols            =   17
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   32768
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel pnl_FecEmi_Dir 
               Height          =   555
               Left            =   -73260
               TabIndex        =   55
               Top             =   330
               Width           =   1065
               _Version        =   65536
               _ExtentX        =   1879
               _ExtentY        =   979
               _StockProps     =   15
               Caption         =   "  Fecha   Emisión"
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
            Begin Threed.SSPanel pnl_FecVto_Dir 
               Height          =   555
               Left            =   -72210
               TabIndex        =   56
               Top             =   330
               Width           =   1065
               _Version        =   65536
               _ExtentX        =   1879
               _ExtentY        =   979
               _StockProps     =   15
               Caption         =   "    Fecha      Vcto."
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
            Begin Threed.SSPanel pnl_NroCre 
               Height          =   555
               Left            =   -74430
               TabIndex        =   57
               Top             =   330
               Width           =   1185
               _Version        =   65536
               _ExtentX        =   2090
               _ExtentY        =   979
               _StockProps     =   15
               Caption         =   "Número"
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
            Begin Threed.SSPanel pnl_PagCom_Dir 
               Height          =   285
               Left            =   -66390
               TabIndex        =   58
               Top             =   600
               Width           =   1350
               _Version        =   65536
               _ExtentX        =   2381
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Pagado"
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
            Begin Threed.SSPanel pnl_Estado 
               Height          =   555
               Left            =   -59790
               TabIndex        =   59
               Top             =   330
               Width           =   1380
               _Version        =   65536
               _ExtentX        =   2434
               _ExtentY        =   979
               _StockProps     =   15
               Caption         =   "Estado"
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
            Begin Threed.SSPanel pnl_PagDes_Dir 
               Height          =   285
               Left            =   -62430
               TabIndex        =   60
               Top             =   600
               Width           =   1350
               _Version        =   65536
               _ExtentX        =   2381
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Pagado"
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
            Begin Threed.SSPanel pnl_SalDes_Dir 
               Height          =   285
               Left            =   -61110
               TabIndex        =   61
               Top             =   600
               Width           =   1350
               _Version        =   65536
               _ExtentX        =   2381
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Saldo"
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
            Begin Threed.SSPanel pnl_SalCom_Dir 
               Height          =   315
               Left            =   -65070
               TabIndex        =   62
               Top             =   600
               Width           =   1350
               _Version        =   65536
               _ExtentX        =   2381
               _ExtentY        =   556
               _StockProps     =   15
               Caption         =   "Saldo"
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
            Begin Threed.SSPanel pnl_ImpDes_Dir 
               Height          =   285
               Left            =   -63750
               TabIndex        =   63
               Top             =   600
               Width           =   1350
               _Version        =   65536
               _ExtentX        =   2381
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
            Begin Threed.SSPanel pnl_MtoPre 
               Height          =   555
               Left            =   -69060
               TabIndex        =   64
               Top             =   330
               Width           =   1380
               _Version        =   65536
               _ExtentX        =   2434
               _ExtentY        =   979
               _StockProps     =   15
               Caption         =   "Monto Préstamo"
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
            Begin Threed.SSPanel pnl_TasInt 
               Height          =   555
               Left            =   -71160
               TabIndex        =   65
               Top             =   330
               Width           =   1065
               _Version        =   65536
               _ExtentX        =   1879
               _ExtentY        =   979
               _StockProps     =   15
               Caption         =   "  Tasa    Interés"
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
            Begin Threed.SSPanel pnl_Modali_Dir 
               Height          =   555
               Left            =   -74940
               TabIndex        =   66
               Top             =   330
               Width           =   525
               _Version        =   65536
               _ExtentX        =   926
               _ExtentY        =   979
               _StockProps     =   15
               Caption         =   "Tipo"
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
            Begin Threed.SSPanel pnl_TasMor 
               Height          =   555
               Left            =   -70110
               TabIndex        =   67
               Top             =   330
               Width           =   1065
               _Version        =   65536
               _ExtentX        =   1879
               _ExtentY        =   979
               _StockProps     =   15
               Caption         =   " Tasa   Moratoria"
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
            Begin Threed.SSPanel pnl_DesDes_Dir 
               Height          =   285
               Left            =   -63750
               TabIndex        =   68
               Top             =   330
               Width           =   3990
               _Version        =   65536
               _ExtentX        =   7038
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "DESEMBOLSOS"
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
            Begin Threed.SSPanel pnl_DesCom_Dir 
               Height          =   285
               Left            =   -67710
               TabIndex        =   69
               Top             =   330
               Width           =   3990
               _Version        =   65536
               _ExtentX        =   7038
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "COMISIONES"
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
            Begin Threed.SSPanel pnl_ImpCom_Dir 
               Height          =   285
               Left            =   -67710
               TabIndex        =   70
               Top             =   600
               Width           =   1350
               _Version        =   65536
               _ExtentX        =   2381
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
         End
      End
   End
   Begin VB.Menu MnuPopUp 
      Caption         =   "MnuPopUp"
      Visible         =   0   'False
      Begin VB.Menu smnu 
         Caption         =   "Imprimir Liquidación"
         Index           =   0
      End
      Begin VB.Menu smnu 
         Caption         =   "Imprimir Carta"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frm_Ges_TecPro_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmb_EstFia_Click()
   Call gs_SetFocus(cmd_Buscar)
End Sub

Private Sub cmb_EstFia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Buscar)
   End If
End Sub

Private Sub cmd_Agrega_Click()
   moddat_g_int_FlgGrb_1 = 1       'Insertar
   frm_Ges_TecPro_04.Show 1
End Sub

Private Sub cmd_Borrar_Click()
   moddat_g_int_FlgGrb_1 = 3
   
   If moddat_g_str_TipCre = 1 And moddat_g_int_TipPan = 1 Then
      If Trim(grd_Listad.TextMatrix(grd_Listad.Row, 16)) <> "VIGENTE" Then
         If Trim(grd_Listad.TextMatrix(grd_Listad.Row, 0)) = "AD" Then
            MsgBox "La Adenda no se encuentra Vigente.", vbExclamation, modgen_g_str_NomPlt
         ElseIf Trim(grd_Listad.TextMatrix(grd_Listad.Row, 0)) = "CSO" Then
            MsgBox "La Carta Seriedad Oferta no se encuentra Vigente.", vbExclamation, modgen_g_str_NomPlt
         Else
            MsgBox "La Carta Fianza no se encuentra Vigente.", vbExclamation, modgen_g_str_NomPlt
         End If
         Call gs_SetFocus(cmb_EstFia)
         Exit Sub
      End If
   ElseIf moddat_g_str_TipCre = 2 And moddat_g_int_TipPan = 0 Then
       If Trim(grd_Listad_Dir.TextMatrix(grd_Listad_Dir.Row, 13)) <> "VIGENTE" Then
         If Trim(grd_Listad_Dir.TextMatrix(grd_Listad_Dir.Row, 0)) = "LC" Then
            MsgBox "La Línea de Crédito no se encuentra Vigente.", vbExclamation, modgen_g_str_NomPlt
         ElseIf Trim(grd_Listad_Dir.TextMatrix(grd_Listad_Dir.Row, 0)) = "CP" Then
            MsgBox "El Crédito Puntual no se encuentra Vigente.", vbExclamation, modgen_g_str_NomPlt
         End If
         Call gs_SetFocus(cmb_EstFia)
         Exit Sub
      End If
   Else
      Exit Sub
   End If
   
   If fs_Validar_MovCFia = True Then
      MsgBox "Verifique que el registro no tenga movimientos en los módulos de Gestión y Garantía.", vbExclamation, modgen_g_str_NomPlt
      If moddat_g_str_TipCre = 1 And moddat_g_int_TipPan = 1 Then
         Call gs_SetFocus(grd_Listad)
      ElseIf moddat_g_str_TipCre = 2 And moddat_g_int_TipPan = 0 Then
         Call gs_SetFocus(grd_Listad_Dir)
      End If
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
      
      If moddat_g_str_TipCre = 1 And moddat_g_int_TipPan = 1 Then
         g_str_Parame = g_str_Parame & " USP_TPR_MAECFI_ELIMINA ("
         g_str_Parame = g_str_Parame & "'" & CStr(Replace(Trim(grd_Listad.TextMatrix(grd_Listad.Row, 2)), "-", "")) & "', "
         g_str_Parame = g_str_Parame & "'" & Format(CStr(grd_Listad.TextMatrix(grd_Listad.Row, 3)), "yyyymmdd") & "', "
      ElseIf moddat_g_str_TipCre = 2 And moddat_g_int_TipPan = 0 Then
         g_str_Parame = g_str_Parame & " USP_TPR_MAECFI_ELIMINA ("
         g_str_Parame = g_str_Parame & "'" & CStr(Replace(Trim(grd_Listad_Dir.TextMatrix(grd_Listad_Dir.Row, 1)), "-", "")) & "', "
         g_str_Parame = g_str_Parame & "'" & Format(CStr(grd_Listad_Dir.TextMatrix(grd_Listad_Dir.Row, 2)), "yyyymmdd") & "', "
      End If
      If g_str_Parame <> "" Then
         g_str_Parame = g_str_Parame & CStr(moddat_g_int_TipDoc) & ", "
         g_str_Parame = g_str_Parame & "'" & CStr(moddat_g_str_NumDoc) & "') " ', "
                     
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
      End If
      Screen.MousePointer = 0
   Loop
   
   'Actualiza la Grilla
   Call fs_Buscar_Creditos_Indirectos
   Call fs_Buscar_Creditos_Directos
   Call frm_Ges_TecPro_01.fs_Buscar
   
   If Me.grd_Listad.Rows = 0 And Me.grd_Listad_Dir.Rows = 0 Then
      Call fs_Activa(True)
   End If
End Sub

Private Function fs_Validar_MovCFia() As Boolean
   fs_Validar_MovCFia = False
   
   If moddat_g_str_TipCre = 1 And moddat_g_int_TipPan = 1 Then If grd_Listad.Rows < 0 Then Exit Function
   If moddat_g_str_TipCre = 2 And moddat_g_int_TipPan = 0 Then If grd_Listad_Dir.Rows < 0 Then Exit Function
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT ( "
   g_str_Parame = g_str_Parame & "         SELECT COUNT(*) "
   g_str_Parame = g_str_Parame & "           FROM TPR_MAEGAR "
   
   If moddat_g_str_TipCre = 1 And moddat_g_int_TipPan = 1 Then
      g_str_Parame = g_str_Parame & "          WHERE MAEGAR_NUMREF = '" & CStr(Trim(grd_Listad.TextMatrix(grd_Listad.Row, 2))) & "' ) CONGAR , "
   Else
      g_str_Parame = g_str_Parame & "          WHERE MAEGAR_NUMREF = '" & CStr(Trim(grd_Listad_Dir.TextMatrix(grd_Listad_Dir.Row, 2))) & "' ) CONGAR , "
   End If
   g_str_Parame = g_str_Parame & "        (SELECT COUNT(*) "
   g_str_Parame = g_str_Parame & "           FROM TPR_MAERDE A "
   
   If moddat_g_str_TipCre = 1 And moddat_g_int_TipPan = 1 Then
      g_str_Parame = g_str_Parame & "          WHERE MAERDE_NUMREF = '" & CStr(Trim(grd_Listad.TextMatrix(grd_Listad.Row, 2))) & "') CONREG "
   Else
      g_str_Parame = g_str_Parame & "          WHERE MAERDE_NUMREF = '" & CStr(Trim(grd_Listad_Dir.TextMatrix(grd_Listad_Dir.Row, 2))) & "') CONREG "
   End If
   
   g_str_Parame = g_str_Parame & "   FROM DUAL"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
     Exit Function
   End If
   
   If g_rst_GenAux.BOF And g_rst_GenAux.EOF Then
      g_rst_GenAux.Close
      Set g_rst_GenAux = Nothing
      Exit Function
   End If
     
   If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
      If (g_rst_GenAux!CONGAR) > 0 Or (g_rst_GenAux!CONREG) > 0 Then
         fs_Validar_MovCFia = True
      End If
   End If
End Function

Private Sub cmd_Buscar_Click()
   Call fs_Buscar_Creditos_Indirectos
   Call fs_Buscar_Creditos_Directos
End Sub

Private Sub cmd_DatBen_Click()
   moddat_g_str_DesObs = ""
   moddat_g_str_DesIte = ""
   moddat_g_str_CodPrd = ""
   moddat_g_str_CodSub = ""
   
   moddat_g_int_FlgGrb_1 = 8
   If moddat_g_str_TipCre = 1 And moddat_g_int_TipPan = 1 Then
      If grd_Listad.Rows > 0 Then
         moddat_g_str_NumFia = Trim(Replace(grd_Listad.TextMatrix(grd_Listad.Row, 2), "-", ""))
         moddat_g_str_DesIte = Trim(grd_Listad.TextMatrix(grd_Listad.Row, 17))
         moddat_g_str_DesObs = Trim(grd_Listad.TextMatrix(grd_Listad.Row, 16))
         moddat_g_str_CodPrd = Trim(grd_Listad.TextMatrix(grd_Listad.Row, 18))
         moddat_g_str_CodSub = Trim(grd_Listad.TextMatrix(grd_Listad.Row, 19))
         moddat_g_str_CodMod = Trim(grd_Listad.TextMatrix(grd_Listad.Row, 20))
         If Trim(grd_Listad.TextMatrix(grd_Listad.Row, 21)) <> "" Then
            moddat_g_int_TipRep = Trim(grd_Listad.TextMatrix(grd_Listad.Row, 21))
         End If
   
         frm_Ges_TecPro_12.Show 1
      End If
   End If
End Sub

Private Sub cmd_Editar_Click()
   moddat_g_int_FlgGrb_1 = 2    'Actualiza
   If fs_Validar = True Then
      frm_Ges_TecPro_04.Show 1
   Else
      If moddat_g_str_TipCre = 1 And moddat_g_int_TipPan = 1 Then
         If grd_Listad.Rows > 0 Then
            If Trim(grd_Listad.TextMatrix(grd_Listad.Row, 0)) = "AD" Then
               MsgBox "La Adenda no se encuentra Vigente.", vbExclamation, modgen_g_str_NomPlt
            ElseIf Trim(grd_Listad.TextMatrix(grd_Listad.Row, 0)) = "CSO" Then
               MsgBox "La Carta Seriedad Oferta no se encuentra Vigente.", vbExclamation, modgen_g_str_NomPlt
            Else
               MsgBox "La Carta Fianza no se encuentra Vigente.", vbExclamation, modgen_g_str_NomPlt
            End If
         End If
      ElseIf moddat_g_str_TipCre = 2 And moddat_g_int_TipPan = 0 Then
         If grd_Listad_Dir.Rows > 0 Then
            If Trim(grd_Listad_Dir.TextMatrix(grd_Listad_Dir.Row, 0)) = "LC" Then
               MsgBox "La Línea de Crédito no se encuentra Vigente.", vbExclamation, modgen_g_str_NomPlt
            ElseIf Trim(grd_Listad_Dir.TextMatrix(grd_Listad_Dir.Row, 0)) = "CP" Then
               MsgBox "El Crédito Puntual no se encuentra Vigente.", vbExclamation, modgen_g_str_NomPlt
            End If
         End If
      End If
      Call gs_SetFocus(cmb_EstFia)
      Exit Sub
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

Private Sub fs_GenWrd_ConSitPro_MejViv(ByVal p_RstDat As ADODB.Recordset)
Dim r_obj_Word          As Word.Application
Dim r_str_Modali        As String
Dim r_str_ParEnt        As String
Dim r_str_ParDec        As String
Dim r_str_MtoLtr        As String

   If Not (p_RstDat.BOF And p_RstDat.EOF) Then
      
      If p_RstDat!MAECFI_CODMOD = "006" Then
         r_str_Modali = "CSP Individual-Construcción en Sitio Propio"
      ElseIf p_RstDat!MAECFI_CODMOD = "007" Then
         r_str_Modali = "MV-Mejoramiento de Vivienda"
      Else
         r_str_Modali = "XXXX"
      End If
      
      r_str_ParEnt = fs_NroEnLetras(CLng(Mid(Format(p_RstDat!MAECFI_GARFIA, "###,##0.00"), 1, InStr(Format(p_RstDat!MAECFI_GARFIA, "###,##0.00"), ".") - 1)))
      r_str_ParDec = Right(Format(p_RstDat!MAECFI_GARFIA, "###,##0.00"), 2)
      r_str_MtoLtr = r_str_ParEnt & "con " & r_str_ParDec & "/100 " & moddat_gf_Consulta_ParDes("204", CStr(p_RstDat!MAECFI_MONFIA))
 
      Set r_obj_Word = CreateObject("Word.Application")
      
      With r_obj_Word
         .Application.Documents.Add (moddat_g_str_RutCFi & "\" & moddat_g_str_NomCF1)
      End With
              
      With r_obj_Word.Selection
          .Font.Name = "Arial"
          .Font.Size = 10
'            .ParagraphFormat.LineSpacing = LinesToPoints(1)
      
          .WholeStory
          With .ParagraphFormat
'                .LeftIndent = CentimetersToPoints(0)
'                .RightIndent = CentimetersToPoints(0)
              .SpaceBefore = 0
              .SpaceBeforeAuto = False
              .SpaceAfter = 0
              .SpaceAfterAuto = False
              .LineSpacingRule = wdLineSpaceSingle
              .WidowControl = True
              .KeepWithNext = False
              .KeepTogether = False
              .PageBreakBefore = False
              .NoLineNumber = False
              .Hyphenation = True
'                .FirstLineIndent = CentimetersToPoints(0)
              .OutlineLevel = wdOutlineLevelBodyText
              .CharacterUnitLeftIndent = 0
              .CharacterUnitRightIndent = 0
              .CharacterUnitFirstLineIndent = 0
              .LineUnitBefore = 0
              .LineUnitAfter = 0
              .MirrorIndents = False
              .TextboxTightWrap = wdTightNone
          End With
      
          .Font.Bold = True
          .TypeText "CARTA FIANZA DE FIEL CUMPLIMIENTO"
          .ParagraphFormat.Alignment = wdAlignParagraphCenter
          .TypeParagraph
          
          .Font.Size = 9
          
          If InStr(r_str_Modali, "-") = 0 Then
            If Not IsNull(p_RstDat!RECURSO) Then
               .TypeText "(Garantiza el desembolso del " & IIf(p_RstDat!RECURSO = "BONO", "BFH", p_RstDat!RECURSO) & " de los beneficiarios)"
            Else
               .TypeText "(Garantiza el desembolso del XXX de los beneficiarios)"
            End If
          Else
            If Not IsNull(p_RstDat!RECURSO) Then
                .TypeText "(Garantiza el desembolso del " & IIf(p_RstDat!RECURSO = "BONO", "BFH", p_RstDat!RECURSO) & " de los beneficiarios - " & Mid(r_str_Modali, 1, InStr(r_str_Modali, "-") - 1) & ")"
            Else
               .TypeText "(Garantiza el desembolso del XXX de los beneficiarios)"
            End If
          End If
          
          .ParagraphFormat.Alignment = wdAlignParagraphCenter
          .TypeParagraph
          .TypeParagraph
          
          .Font.Size = 10
          .Font.Bold = False
          .TypeText "San Isidro, " & Day(gf_FormatoFecha(CStr(p_RstDat!MAECFI_EMIFIA))) & " de " & fs_NomMes(Month(gf_FormatoFecha(CStr(p_RstDat!MAECFI_EMIFIA)))) & " del " & Year(gf_FormatoFecha(CStr(p_RstDat!MAECFI_EMIFIA)))
          .ParagraphFormat.Alignment = wdAlignParagraphRight
          .TypeParagraph
          .TypeParagraph
          
          .Font.Bold = True
          .TypeText "Carta Fianza N°: " & gf_Formato_NumRef(CStr(Trim(p_RstDat!MAECFI_NUMREF)), Mid(p_RstDat!MAECFI_NUMREF, 1, 1))
          .ParagraphFormat.Alignment = wdAlignParagraphLeft
          .TypeParagraph
          .TypeParagraph

          .Font.Bold = False
          .TypeText "Señores"
          .ParagraphFormat.Alignment = wdAlignParagraphLeft
          .TypeParagraph
          
          .Font.Bold = True
          .TypeText "FONDO MIVIVIENDA S.A."
          .ParagraphFormat.Alignment = wdAlignParagraphLeft
          .TypeParagraph
          
          .Font.Bold = False
          .Font.Underline = True
          .TypeText "Presente"
          .Font.Underline = False
          .TypeText ". -"
          .ParagraphFormat.Alignment = wdAlignParagraphLeft
          .TypeParagraph
          .TypeParagraph
          
          .Font.Underline = False
          .TypeText "De nuestra consideración:"
          .ParagraphFormat.Alignment = wdAlignParagraphLeft
          .TypeParagraph
          .TypeParagraph
          
          .ParagraphFormat.Alignment = wdAlignParagraphJustify
          .TypeText "A solicitud de nuestro cliente "
          .Font.Bold = True
          .TypeText p_RstDat!MAEPRV_RAZSOC
          .Font.Bold = False
          .TypeText ", otorgamos fianza irrevocable, solidaria, incondicionada, sin beneficio de excusión y de realización automática a simple solicitud y a favor del Fondo MIVIVIENDA S.A, por la suma de S/ " & Format(p_RstDat!MAECFI_GARFIA, "###,###,###,##0.00")
          .TypeText " (" & r_str_MtoLtr & ")"
          .TypeText " para garantizar las obligaciones derivadas de su participación en el Programa Techo Propio en la Modalidad de " & Mid(r_str_Modali, InStr(r_str_Modali, "-") + 1) & ", y las obligaciones adquiridas en cada uno de los Contratos de Construcción de un total de "
          .TypeText p_RstDat!CANT_BEN
          .TypeText " viviendas que se encuentra inscritas en el Registro de Proyectos de Vivienda del Programa Techo Propio del Ministerio de Vivienda, Construcción y Saneamiento, contratos que están suscritos con los beneficiarios del Bono Familiar Habitacional (BFH), cuyos nombres, apellidos, documento oficial de identidad, código de registro de su vivienda (proyecto), y valor del "
          If Not IsNull(p_RstDat!RECURSO) Then
             .TypeText p_RstDat!RECURSO
          End If
          .TypeText ", figuran en el anexo adjunto que forma parte del presente documento."
          .TypeParagraph
          .TypeParagraph
          
          .TypeText "La presente fianza rige desde el " & Day(gf_FormatoFecha(CStr(p_RstDat!MAECFI_EMIFIA))) & " de " & fs_NomMes(Month(gf_FormatoFecha(CStr(p_RstDat!MAECFI_EMIFIA)))) & " de " & Year(gf_FormatoFecha(CStr(p_RstDat!MAECFI_EMIFIA)))
          .TypeText " hasta el " & Day(gf_FormatoFecha(CStr(p_RstDat!MAECFI_VTOFIA))) & " de " & fs_NomMes(Month(gf_FormatoFecha(CStr(p_RstDat!MAECFI_VTOFIA)))) & " del " & Year(gf_FormatoFecha(CStr(p_RstDat!MAECFI_VTOFIA))) & "."
          .TypeParagraph
          .TypeParagraph
          
          .TypeText "La falta de renovación oportuna, antes de la fecha de vencimiento de esta garantía, es causal de ejecución de la misma."
          .TypeParagraph
          .TypeParagraph
          
          .TypeText "Queda expresamente entendido por nosotros que esta fianza será ejecutada por la entidad a cuyo favor se emite, de conformidad con lo dispuesto en el artículo 1898° del Código Civil vigente."
          .TypeParagraph
          .TypeParagraph
          
          .TypeText "Para honrar la presente Fianza bastará el simple requerimiento efectuado por ustedes mediante conducto notarial y toda demora de nuestra parte para honrarla devengará un interés a la Libor de 180 días más "
          .TypeText p_RstDat!MAECFI_TASFIA & "%"
          .TypeText ". La tasa Libor de 180 días será la de cierre al día anterior al pago, debiendo devengarse los intereses a partir del quinto día hábil luego de la fecha en que sea exigido el honramiento de esta Fianza."
          .TypeParagraph
          .TypeParagraph
          
          .TypeText "Atentamente,"
          .ParagraphFormat.Alignment = wdAlignParagraphJustify
          
          .TypeParagraph
          .TypeParagraph
          .TypeParagraph
          .TypeParagraph
          .TypeParagraph
          .TypeParagraph
          .TypeParagraph
          .TypeParagraph
                      
          .Font.Bold = True
'            .ParagraphFormat.LineSpacing = LinesToPoints(1.5)
          .TypeText "__________________________                        "
          .TypeText "_________________________________________"
          .TypeParagraph
          .TypeParagraph
          
          .TypeText "  Julio Rodríguez Nakagawa"
          .TypeText "                                                  "
          .TypeText "Carlos Pacheco Caycho"
          .TypeParagraph
          .TypeParagraph
          .TypeText "GERENTE DE PRODUCCION"
          .TypeText "                           "
          .TypeText "GERENTE DE ADMINISTRACION Y TESORERIA"
          
          .InsertNewPage
          .Font.Bold = True
          .TypeText "Anexo Carta Fianza N° " & gf_Formato_NumRef(CStr(Trim(p_RstDat!MAECFI_NUMREF)), Mid(p_RstDat!MAECFI_NUMREF, 1, 1))
          .ParagraphFormat.Alignment = wdAlignParagraphCenter
          .TypeParagraph
          .TypeParagraph
          .Font.Size = 9
          
          'Tabla
          If p_RstDat!CANT_BEN > 0 Then
              Call fs_docCrearTabla(8, CInt(p_RstDat!CANT_BEN), r_obj_Word, Trim(p_RstDat!MAECFI_NUMREF), CInt(p_RstDat!MAECFI_TIPREC))
              Call fs_docCombinarCeldas(5, r_obj_Word)
          End If
          
      'Guarda el documento
      '   objWD.ActiveDocument.SaveAs FileName:="mydoc.doc"
      End With
      r_obj_Word.Visible = True
      Set r_obj_Word = Nothing
    End If
End Sub

Private Sub fs_GenWrd_AdqViv(ByVal p_RstDat As ADODB.Recordset)
Dim r_obj_Word          As Word.Application
Dim r_str_NomPry        As String
Dim r_str_NOMPRO        As String
Dim r_str_TipPry        As String
Dim r_str_MtoLtr        As String
Dim r_str_ParEnt        As String
Dim r_str_ParDec        As String

   If Not (p_RstDat.BOF And p_RstDat.EOF) Then
   
'      r_str_ParEnt = gf_Convertir_NumLet(CLng(Mid(Format(p_RstDat!MAECFI_GARFIA, "###,##0.00"), 1, InStr(Format(p_RstDat!MAECFI_GARFIA, "###,##0.00"), ".") - 1)))
      r_str_ParEnt = fs_NroEnLetras(CLng(Mid(Format(p_RstDat!MAECFI_GARFIA, "###,##0.00"), 1, InStr(Format(p_RstDat!MAECFI_GARFIA, "###,##0.00"), ".") - 1)))
      r_str_ParDec = Right(Format(p_RstDat!MAECFI_GARFIA, "###,##0.00"), 2)
      r_str_MtoLtr = r_str_ParEnt & "con " & r_str_ParDec & "/100 " & moddat_gf_Consulta_ParDes("204", CStr(p_RstDat!MAECFI_MONFIA))
      
      If Not IsNull(g_rst_Princi!MAECFI_NOMPRY) Then
         Call modmip_gs_Consulta_NomPry(CStr(g_rst_Princi!MAECFI_NOMPRY), r_str_NomPry, r_str_NOMPRO, r_str_TipPry)
      End If
      
      Set r_obj_Word = CreateObject("Word.Application")
      
      With r_obj_Word
         .Application.Documents.Add (moddat_g_str_RutCFi & "\" & moddat_g_str_NomCF1)
      End With
        
      With r_obj_Word.Selection
          .Font.Name = "Arial"
          .Font.Size = 10
      '            .ParagraphFormat.LineSpacing = LinesToPoints(1)
          
          .WholeStory
          With .ParagraphFormat
      '                .LeftIndent = CentimetersToPoints(0)
      '                .RightIndent = CentimetersToPoints(0)
              .SpaceBefore = 0
              .SpaceBeforeAuto = False
              .SpaceAfter = 0
              .SpaceAfterAuto = False
              .LineSpacingRule = wdLineSpaceSingle
              .WidowControl = True
              .KeepWithNext = False
              .KeepTogether = False
              .PageBreakBefore = False
              .NoLineNumber = False
              .Hyphenation = True
      '                .FirstLineIndent = CentimetersToPoints(0)
              .OutlineLevel = wdOutlineLevelBodyText
              .CharacterUnitLeftIndent = 0
              .CharacterUnitRightIndent = 0
              .CharacterUnitFirstLineIndent = 0
              .LineUnitBefore = 0
              .LineUnitAfter = 0
              .MirrorIndents = False
              .TextboxTightWrap = wdTightNone
          End With
      
          .Font.Bold = True
          .TypeText "CARTA FIANZA DE FIEL CUMPLIMIENTO"
          .ParagraphFormat.Alignment = wdAlignParagraphCenter
          .TypeParagraph
          .Font.Size = 9
          
          If Not IsNull(p_RstDat!RECURSO) Then
            .TypeText "(Garantiza el desembolso del " & IIf(p_RstDat!RECURSO = "BONO", "BFH", p_RstDat!RECURSO) & " de los beneficiarios)"
          Else
            .TypeText "(Garantiza el desembolso del XXX de los beneficiarios)"
          End If
                   
          .ParagraphFormat.Alignment = wdAlignParagraphCenter
          .TypeParagraph
          .TypeParagraph
          
          .Font.Size = 10
          .Font.Bold = False
          .TypeText "San Isidro, " & Day(gf_FormatoFecha(CStr(p_RstDat!MAECFI_EMIFIA))) & " de " & fs_NomMes(Month(gf_FormatoFecha(CStr(p_RstDat!MAECFI_EMIFIA)))) & " del " & Year(gf_FormatoFecha(CStr(p_RstDat!MAECFI_EMIFIA)))
          .ParagraphFormat.Alignment = wdAlignParagraphRight
          .TypeParagraph
          .TypeParagraph
                      
          .Font.Bold = True
          .TypeText "Carta Fianza N°: " & gf_Formato_NumRef(CStr(Trim(p_RstDat!MAECFI_NUMREF)), Mid(p_RstDat!MAECFI_NUMREF, 1, 1))
          .ParagraphFormat.Alignment = wdAlignParagraphLeft
          .TypeParagraph
          .TypeParagraph
          
          .Font.Bold = False
          .TypeText "Señores"
          .ParagraphFormat.Alignment = wdAlignParagraphLeft
          .TypeParagraph
          .Font.Bold = True
          .TypeText "FONDO MIVIVIENDA S.A."
          .ParagraphFormat.Alignment = wdAlignParagraphLeft
          .TypeParagraph
          .Font.Bold = False
          .Font.Underline = True
          .TypeText "Presente"
          .Font.Underline = False
          .TypeText ". -"
          .ParagraphFormat.Alignment = wdAlignParagraphLeft
          .TypeParagraph
          .TypeParagraph
          .Font.Underline = False
          .TypeText "De nuestra consideración:"
          .ParagraphFormat.Alignment = wdAlignParagraphLeft
          .TypeParagraph
          .TypeParagraph
          .ParagraphFormat.Alignment = wdAlignParagraphJustify
          .TypeText "A solicitud de nuestro cliente "
          .Font.Bold = True
          .TypeText p_RstDat!MAEPRV_RAZSOC
          .Font.Bold = False
          .TypeText ", otorgamos fianza irrevocable, solidaria, incondicionada, sin beneficio de excusión y de realización automática a simple solicitud y a favor del Fondo MIVIVIENDA S.A, por la suma de S/ " & Format(p_RstDat!MAECFI_GARFIA, "###,###,###,##0.00")
          .TypeText " (" & r_str_MtoLtr & ")"
          .TypeText " para garantizar las obligaciones del cliente derivadas de su participación en el Proyecto Techo Propio, y las obligaciones asumidas por el cliente en las minutas de compraventa de "
          .TypeText p_RstDat!CANT_BEN
          .TypeText " viviendas del Proyecto Inmobiliario denominado "
          .Font.Bold = True
          .TypeText r_str_NomPry
          .Font.Bold = False
          .TypeText " con"
          .Font.Bold = True
          .TypeText " código de Proyecto N° "
          If Not IsNull(p_RstDat!MAECFI_CODPRY) Then
              .TypeText CStr(Trim(p_RstDat!MAECFI_CODPRY))
          End If
          .Font.Bold = False
          .TypeText " inscrito en el Registro de Proyectos de Vivienda del Proyecto Techo Propio del Ministerio de Vivienda, Construcción y Saneamiento, que han sido adquiridas por los beneficiarios del Bono Familiar Habitacional (BFH), cuyos nombres, apellidos, documento oficial de identidad, y valor del "
          If Not IsNull(p_RstDat!RECURSO) Then
             .TypeText p_RstDat!RECURSO
          End If
          .TypeText ", figuran en el anexo adjunto que forma parte del presente documento."
          .TypeParagraph
          .TypeParagraph
          .TypeText "La presente fianza rige desde el " & Day(gf_FormatoFecha(CStr(p_RstDat!MAECFI_EMIFIA))) & " de " & fs_NomMes(Month(gf_FormatoFecha(CStr(p_RstDat!MAECFI_EMIFIA)))) & " de " & Year(gf_FormatoFecha(CStr(p_RstDat!MAECFI_EMIFIA)))
          .TypeText " hasta el " & Day(gf_FormatoFecha(CStr(p_RstDat!MAECFI_VTOFIA))) & " de " & fs_NomMes(Month(gf_FormatoFecha(CStr(p_RstDat!MAECFI_VTOFIA)))) & " del " & Year(gf_FormatoFecha(CStr(p_RstDat!MAECFI_VTOFIA))) & "."
          .TypeParagraph
          .TypeParagraph
          .TypeText "La falta de renovación oportuna, antes de la fecha de vencimiento de esta garantía, es causal de ejecución de la misma."
          .TypeParagraph
          .TypeParagraph
          .TypeText "Queda expresamente entendido por nosotros que esta fianza será ejecutada por la entidad a cuyo favor se emite, de conformidad con lo dispuesto en el artículo 1898° del Código Civil vigente."
          .TypeParagraph
          .TypeParagraph
          .TypeText "Para honrar la presente Fianza bastará el simple requerimiento efectuado por ustedes mediante conducto notarial y toda demora de nuestra parte para honrarla devengará un interés a la Libor de 180 días más "
          .TypeText p_RstDat!MAECFI_TASFIA & "%"
          .TypeText ". La tasa Libor de 180 días será la de cierre al día anterior al pago, debiendo devengarse los intereses a partir del quinto día hábil luego de la fecha en que sea exigido el honramiento de esta Fianza."
          .TypeParagraph
          .TypeParagraph
          .TypeText "Atentamente,"
          .ParagraphFormat.Alignment = wdAlignParagraphJustify
          
          .TypeParagraph
          .TypeParagraph
          .TypeParagraph
          .TypeParagraph
          .TypeParagraph
          .TypeParagraph
          .TypeParagraph
          .TypeParagraph
      '           .TypeParagraph
      '           .TypeParagraph
                       
         .Font.Bold = True
'            .ParagraphFormat.LineSpacing = LinesToPoints(1.5)
          .TypeText "__________________________                        "
          .TypeText "_________________________________________"
          .TypeParagraph
          .TypeParagraph
          
          .TypeText "  Julio Rodríguez Nakagawa"
          .TypeText "                                                  "
          .TypeText "Carlos Pacheco Caycho"
          .TypeParagraph
          .TypeParagraph
          .TypeText "GERENTE DE PRODUCCION"
          .TypeText "                           "
          .TypeText "GERENTE DE ADMINISTRACION Y TESORERIA"
         
          .InsertNewPage
          .Font.Bold = True
          .TypeText "Anexo Carta Fianza N° " & gf_Formato_NumRef(CStr(Trim(p_RstDat!MAECFI_NUMREF)), Mid(p_RstDat!MAECFI_NUMREF, 1, 1))
          .ParagraphFormat.Alignment = wdAlignParagraphCenter
          .TypeParagraph
          .TypeParagraph
          .TypeParagraph
          .Font.Size = 9
         
          'Tabla
          If p_RstDat!CANT_BEN > 0 Then
              Call fs_docCrearTabla(7, CInt(p_RstDat!CANT_BEN), r_obj_Word, Trim(p_RstDat!MAECFI_NUMREF), CInt(p_RstDat!MAECFI_TIPREC))
              Call fs_docCombinarCeldas(4, r_obj_Word)
          End If
            
       End With
       
       r_obj_Word.Visible = True
      Set r_obj_Word = Nothing
    End If
End Sub

Private Sub fs_docCrearTabla(numcols As Integer, numrows As Integer, p_objwrd As Word.Application, p_NumRef As String, p_TipRec As Integer)
Dim r_int_NumCol    As Integer
Dim r_int_NumFil    As Integer
Dim r_dbl_ImpTot    As Double
Dim r_str_PorGar    As String
Dim r_str_CodPrd    As String

    g_str_Parame = ""
    g_str_Parame = g_str_Parame & "  SELECT DATBEN_CODIGO , DATBEN_TIPDOC , DATBEN_NUMDOC , DATBEN_APEPAT , DATBEN_APEMAT , DATBEN_APECAS , "
    g_str_Parame = g_str_Parame & "         DATBEN_NOMBRE , DATBEN_IMPBFH , DATBEN_CODPRY , TRIM(PARDES_DESCRI) AS PORGAR , MAECFI_CODPRD "
    g_str_Parame = g_str_Parame & "    FROM TPR_DATBEN "
    g_str_Parame = g_str_Parame & "         INNER JOIN TPR_MAECFI ON MAECFI_NUMREF = DATBEN_NUMREF "
    g_str_Parame = g_str_Parame & "          LEFT JOIN MNT_PARDES ON PARDES_CODGRP = 535 AND PARDES_CODITE = MAECFI_PORGAR "
    g_str_Parame = g_str_Parame & "   WHERE DATBEN_NUMREF = '" & p_NumRef & "' "
    g_str_Parame = g_str_Parame & "     AND DATBEN_TIPREC = " & p_TipRec & ""
    g_str_Parame = g_str_Parame & "   ORDER BY DATBEN_CODIGO "
    
    If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
        Exit Sub
    End If
    
    If g_rst_Genera.BOF And g_rst_Genera.EOF Then
        g_rst_Genera.Close
        Set g_rst_Genera = Nothing
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    
    If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
        g_rst_Genera.MoveFirst
        r_str_PorGar = Replace(g_rst_Genera!PORGAR, "%", "")
        r_str_CodPrd = g_rst_Genera!MAECFI_CODPRD
        
        p_objwrd.Selection.Tables.Add Range:=p_objwrd.Selection.Range, numrows:=numrows + 1, NumColumns:=numcols
            
        With p_objwrd.Selection.Tables(p_objwrd.Selection.Tables.Count)
            
            If .Style <> "Tabla con cuadrícula" Then
                .Style = "Tabla con cuadrícula"
            End If
            .Columns(1).Width = 25
            .Columns(2).Width = 45
            .Columns(3).Width = 55
            .Columns(4).Width = 75
            .Columns(5).Width = 75
            .Columns(6).Width = 95
            
            If numcols = 7 Then
               .Columns(7).Width = 95
            Else
               .Columns(7).Width = 100
               .Columns(8).Width = 75
            End If
            
            'CABECERA
            With p_objwrd.Selection
            
                .Tables(1).Rows(1).Select               'Seleccionamos la 1° fila
                .Font.Bold = True                       'Le ponemos negritas
                .Tables(1).Cell(1, 1).Select            'nos posicionamos en la celda 1,1
                .ParagraphFormat.Alignment = wdAlignParagraphCenter
                .Cells.VerticalAlignment = wdCellAlignVerticalCenter
                .TypeText "N°"                          'Escribimos
                
                .Tables(1).Cell(1, 2).Select            'nos posicionamos en la celda 1,2
                .ParagraphFormat.Alignment = wdAlignParagraphCenter
                .Cells.VerticalAlignment = wdCellAlignVerticalCenter
                .TypeText "Tipo Doc."                   'Escribimos
                
                .Tables(1).Cell(1, 3).Select            'nos posicionamos en la celda 1,3
                .ParagraphFormat.Alignment = wdAlignParagraphCenter
                .Cells.VerticalAlignment = wdCellAlignVerticalCenter
                .TypeText "DNI"                         'Escribimos
                
                .Tables(1).Cell(1, 4).Select        'nos posicionamos en la celda 1,4
                .ParagraphFormat.Alignment = wdAlignParagraphCenter
                .Cells.VerticalAlignment = wdCellAlignVerticalCenter
                .TypeText "Apellido Paterno"        'Escribimos
                
                .Tables(1).Cell(1, 5).Select            'nos posicionamos en la celda 1,5
                .ParagraphFormat.Alignment = wdAlignParagraphCenter
                .Cells.VerticalAlignment = wdCellAlignVerticalCenter
                .TypeText "Apellido Materno"            'Escribimos
                
                .Tables(1).Cell(1, 6).Select            'nos posicionamos en la celda 1,6
                .ParagraphFormat.Alignment = wdAlignParagraphCenter
                .Cells.VerticalAlignment = wdCellAlignVerticalCenter
                .TypeText "Nombres"                     'Escribimos
                
                If numcols = 7 Then
                    .Tables(1).Cell(1, 7).Select        'nos posicionamos en la celda 1,8
                    .ParagraphFormat.Alignment = wdAlignParagraphCenter
                    .Cells.VerticalAlignment = wdCellAlignVerticalCenter
                    
                    If p_TipRec = 1 Then
                        If r_str_CodPrd = "026" Then
                           .TypeText "Importe del BFH(S/.)"        'Escribimos
                        ElseIf r_str_CodPrd = "027" Then
                           .TypeText "Importe del BPVV(S/.)"       'Escribimos
                        Else
                           .TypeText "Importe del BONO(S/.)"       'Escribimos
                        End If
                    ElseIf p_TipRec = 2 Then
                        .TypeText "Importe del AHORRO(S/.)"    'Escribimos
                    End If
                Else
                    .Tables(1).Cell(1, 7).Select        'nos posicionamos en la celda 1,7
                    .ParagraphFormat.Alignment = wdAlignParagraphCenter
                    .Cells.VerticalAlignment = wdCellAlignVerticalCenter
                    .TypeText "Código Proyecto"
                
                    .Tables(1).Cell(1, 8).Select        'nos posicionamos en la celda 1,8
                    .ParagraphFormat.Alignment = wdAlignParagraphCenter
                    .Cells.VerticalAlignment = wdCellAlignVerticalCenter
                    If p_TipRec = 1 Then
                        If r_str_CodPrd = "026" Then
                           .TypeText "Importe del BFH(S/.)"       'Escribimos
                        ElseIf r_str_CodPrd = "027" Then
                           .TypeText "Importe del BPVV(S/.)"       'Escribimos
                        Else
                           .TypeText "Importe del BONO(S/.)"       'Escribimos
                        End If
                    ElseIf p_TipRec = 2 Then
                        .TypeText "Importe del AHORRO(S/.)"    'Escribimos
                    End If
                End If
                .MoveEnd
'                   .MoveDown                           'salimos de la tabla

            End With
            
            Do While Not g_rst_Genera.EOF
                For r_int_NumFil = 2 To numrows + 1
                    For r_int_NumCol = 1 To numcols
                        p_objwrd.Selection.Tables(1).Rows(r_int_NumFil).Select
                        p_objwrd.Selection.Font.Bold = False
                        p_objwrd.Selection.Tables(1).Cell(r_int_NumFil, r_int_NumCol).Select
                        
                        If r_int_NumCol = 1 Then
                            p_objwrd.Selection.TypeText CInt(r_int_NumFil - 1) 'g_rst_Genera!DATBEN_CODIGO
                        ElseIf r_int_NumCol = 2 Then
                            p_objwrd.Selection.TypeText moddat_gf_Consulta_ParDes("270", g_rst_Genera!DATBEN_TIPDOC)
                        
                        ElseIf r_int_NumCol = 3 Then
                            p_objwrd.Selection.TypeText Trim(g_rst_Genera!DATBEN_NUMDOC)
                        
                        ElseIf r_int_NumCol = 4 Then
                            p_objwrd.Selection.Tables(1).Cell(r_int_NumFil, 4).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
                            p_objwrd.Selection.TypeText Trim(g_rst_Genera!DATBEN_APEPAT)
                        
                        ElseIf r_int_NumCol = 5 Then
                            p_objwrd.Selection.Tables(1).Cell(r_int_NumFil, 5).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
                            p_objwrd.Selection.TypeText Trim(g_rst_Genera!DATBEN_APEMAT)
                        
                        ElseIf r_int_NumCol = 6 Then
                            p_objwrd.Selection.Tables(1).Cell(r_int_NumFil, 6).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
                            p_objwrd.Selection.TypeText Trim(g_rst_Genera!DATBEN_NOMBRE)
                        End If
                        
                        If numcols = 7 Then
                            If r_int_NumCol = 7 Then
                                p_objwrd.Selection.Tables(1).Cell(r_int_NumFil, 7).Range.ParagraphFormat.Alignment = wdAlignParagraphRight
                                p_objwrd.Selection.TypeText "S/ " & Format(CDbl(g_rst_Genera!DATBEN_IMPBFH), "###,###,###,##0.00")
                                r_dbl_ImpTot = r_dbl_ImpTot + CDbl(g_rst_Genera!DATBEN_IMPBFH)
                            End If
                        Else
                            If r_int_NumCol = 7 Then
                                p_objwrd.Selection.Tables(1).Cell(r_int_NumFil, 7).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
                                p_objwrd.Selection.TypeText Trim(g_rst_Genera!DATBEN_CODPRY)
                                
                            ElseIf r_int_NumCol = 8 Then
                                p_objwrd.Selection.Tables(1).Cell(r_int_NumFil, 8).Range.ParagraphFormat.Alignment = wdAlignParagraphRight
                                p_objwrd.Selection.TypeText "S/ " & Format(CDbl(g_rst_Genera!DATBEN_IMPBFH), "###,###,###,##0.00")
                                r_dbl_ImpTot = r_dbl_ImpTot + CDbl(g_rst_Genera!DATBEN_IMPBFH)
                            End If
                        End If
                    Next r_int_NumCol
                    g_rst_Genera.MoveNext
                Next r_int_NumFil
            Loop
            
            'TOTAL
            .Rows.Add
            p_objwrd.Selection.Tables(1).Rows(.Rows.Count).Select
            p_objwrd.Selection.Font.Bold = True
            p_objwrd.Selection.Tables(1).Cell(.Rows.Count, numcols - 1).Select
            p_objwrd.Selection.Tables(1).Cell(.Rows.Count, numcols - 1).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
            p_objwrd.Selection.TypeText "TOTAL"
            p_objwrd.Selection.Tables(1).Cell(.Rows.Count, numcols).Select
            p_objwrd.Selection.Tables(1).Cell(.Rows.Count, numcols).Range.ParagraphFormat.Alignment = wdAlignParagraphRight
            p_objwrd.Selection.TypeText "S/ " & Format(CDbl(r_dbl_ImpTot), "###,###,###,##0.00")
            
            '% garantizado, 05% o 10%
            .Rows.Add
            p_objwrd.Selection.Tables(1).Rows(.Rows.Count).Select
            p_objwrd.Selection.Font.Bold = True
            p_objwrd.Selection.Tables(1).Cell(.Rows.Count, numcols - 1).Select
            p_objwrd.Selection.Tables(1).Cell(.Rows.Count, numcols - 1).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
            
            g_rst_Genera.MoveFirst
            
            p_objwrd.Selection.TypeText r_str_PorGar & "%"
            p_objwrd.Selection.Tables(1).Cell(.Rows.Count, numcols).Select
            p_objwrd.Selection.Font.Bold = False
            p_objwrd.Selection.Tables(1).Cell(.Rows.Count, numcols).Range.ParagraphFormat.Alignment = wdAlignParagraphRight
            p_objwrd.Selection.TypeText "S/ " & Format(CDbl(r_dbl_ImpTot * (CDbl(r_str_PorGar)) / 100), "###,###,###,##0.00")
            
            'TOTAL BFH 100% + %garantizado
            .Rows.Add
            p_objwrd.Selection.Tables(1).Rows(.Rows.Count).Select
            p_objwrd.Selection.Font.Bold = True
            p_objwrd.Selection.Tables(1).Cell(.Rows.Count, numcols - 1).Select
            p_objwrd.Selection.Tables(1).Cell(.Rows.Count, numcols - 1).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
            
            If r_str_CodPrd = "026" Then
               p_objwrd.Selection.TypeText "TOTAL " & IIf(p_TipRec = 2, "AHORRO", "BFH") & " " & 100 + CDbl(r_str_PorGar) & "%"
            ElseIf r_str_CodPrd = "027" Then
               p_objwrd.Selection.TypeText "TOTAL " & IIf(p_TipRec = 2, "AHORRO", "BPVV") & " " & 100 + CDbl(r_str_PorGar) & "%"
            Else
               p_objwrd.Selection.TypeText "TOTAL " & IIf(p_TipRec = 2, "AHORRO", "BONO") & " " & 100 + CDbl(r_str_PorGar) & "%"
            End If
            p_objwrd.Selection.Tables(1).Cell(.Rows.Count, numcols).Select
            p_objwrd.Selection.Tables(1).Cell(.Rows.Count, numcols).Range.ParagraphFormat.Alignment = wdAlignParagraphRight
            p_objwrd.Selection.TypeText "S/ " & Format(CDbl(r_dbl_ImpTot) * (1 + (CDbl(r_str_PorGar) / 100)), "###,###,###,##0.00")
        End With
        
    End If
End Sub

Private Sub fs_docCombinarCeldas(numceldas As Integer, p_objwrd As Word.Application)
    
    With p_objwrd.Selection
    
        .Tables(1).Cell(.Tables(1).Rows.Count - 2, 1).Select
        .MoveRight Unit:=wdCharacter, Count:=numceldas, Extend:=wdExtend
        .Cells.Merge
    
        .Tables(1).Cell(.Tables(1).Rows.Count - 1, 1).Select
        .MoveRight Unit:=wdCharacter, Count:=numceldas, Extend:=wdExtend
        .Cells.Merge
    
        .Tables(1).Cell(.Tables(1).Rows.Count, 1).Select
        .MoveRight Unit:=wdCharacter, Count:=numceldas, Extend:=wdExtend
        .Cells.Merge
    
        .Tables(1).Cell(.Tables(1).Rows.Count - 2, 1).Select
        .MoveDown Unit:=wdLine, Count:=2, Extend:=wdExtend
        .Cells.Merge
        
        .Tables(1).Cell(.Tables(1).Rows.Count - 2, 1).Select
        With .Cells
            .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
            .Borders(wdBorderRight).LineStyle = wdLineStyleSingle
            .Borders(wdBorderTop).LineStyle = wdLineStyleSingle
            .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
            .Borders(wdBorderHorizontal).LineStyle = wdLineStyleSingle
            .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
            .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
            .Borders.Shadow = False
        End With

        .Tables(1).Rows.Select
        .Rows.Alignment = wdAlignRowCenter
        
        .MoveEnd
        .MoveDown
  
    End With
End Sub

Private Sub fs_GenWrd_CSO(ByVal p_RstDat As ADODB.Recordset)
Dim r_obj_Word          As Word.Application
Dim r_str_MtoLtr        As String
Dim r_str_ParEnt        As String
Dim r_str_ParDec        As String
Dim r_str_Modali        As String

   If Not (p_RstDat.BOF And p_RstDat.EOF) Then
    
'      r_str_ParEnt = gf_Convertir_NumLet(CLng(Mid(Format(p_RstDat!MAECFI_GARFIA, "###,##0.00"), 1, InStr(Format(p_RstDat!MAECFI_GARFIA, "###,##0.00"), ".") - 1)))
      r_str_ParEnt = fs_NroEnLetras(CLng(Mid(Format(p_RstDat!MAECFI_GARFIA, "###,##0.00"), 1, InStr(Format(p_RstDat!MAECFI_GARFIA, "###,##0.00"), ".") - 1)))
      r_str_ParDec = Right(Format(p_RstDat!MAECFI_GARFIA, "###,##0.00"), 2)
      r_str_MtoLtr = r_str_ParEnt & "con " & r_str_ParDec & "/100 " & moddat_gf_Consulta_ParDes("204", CStr(p_RstDat!MAECFI_MONFIA))
 
      If p_RstDat!MAECFI_CODPRD = "026" Then
         If p_RstDat!MAECFI_CODSUB = "001" Then
            r_str_Modali = "Adquisición de Vivienda Nueva"
         ElseIf p_RstDat!MAECFI_CODSUB = "002" Then
            r_str_Modali = "Construcción en Sitio Propio"
         ElseIf p_RstDat!MAECFI_CODSUB = "003" Then
            r_str_Modali = "Mejoramiento de Vivienda"
         End If
      ElseIf p_RstDat!MAECFI_CODPRD = "027" Then
         r_str_Modali = "Bono de Reforzamiento Estructural"
      End If
      
      Set r_obj_Word = CreateObject("Word.Application")
       
      With r_obj_Word
         .Application.Documents.Add (moddat_g_str_RutCFi & "\" & moddat_g_str_NomCF1)
      End With
        
      With r_obj_Word.Selection
          .Font.Name = "Arial"
          .Font.Size = 10
'            .ParagraphFormat.LineSpacing = LinesToPoints(1)
          
          .WholeStory
          With .ParagraphFormat
'                .LeftIndent = CentimetersToPoints(0)
'                .RightIndent = CentimetersToPoints(0)
              .SpaceBefore = 0
              .SpaceBeforeAuto = False
              .SpaceAfter = 0
              .SpaceAfterAuto = False
              .LineSpacingRule = wdLineSpaceSingle
              .WidowControl = True
              .KeepWithNext = False
              .KeepTogether = False
              .PageBreakBefore = False
              .NoLineNumber = False
              .Hyphenation = True
'                .FirstLineIndent = CentimetersToPoints(0)
              .OutlineLevel = wdOutlineLevelBodyText
              .CharacterUnitLeftIndent = 0
              .CharacterUnitRightIndent = 0
              .CharacterUnitFirstLineIndent = 0
              .LineUnitBefore = 0
              .LineUnitAfter = 0
              .MirrorIndents = False
              .TextboxTightWrap = wdTightNone
          End With
      
          .Font.Bold = True
          .TypeText "CARTA FIANZA BANCARIA DE "
          .ParagraphFormat.Alignment = wdAlignParagraphCenter
          .TypeParagraph
          .TypeText "GARANTÍA DE SERIEDAD DE LA OFERTA"
          .ParagraphFormat.Alignment = wdAlignParagraphCenter
          .TypeParagraph
          .TypeParagraph

          .Font.Size = 10
          .Font.Bold = False
          .TypeText "San Isidro, " & Day(gf_FormatoFecha(CStr(p_RstDat!MAECFI_EMIFIA))) & " de " & fs_NomMes(Month(gf_FormatoFecha(CStr(p_RstDat!MAECFI_EMIFIA)))) & " del " & Year(gf_FormatoFecha(CStr(p_RstDat!MAECFI_EMIFIA)))
          .ParagraphFormat.Alignment = wdAlignParagraphRight
          .TypeParagraph
          .TypeParagraph
          
          .Font.Bold = True
          .TypeText "Carta Fianza N°: " & gf_Formato_NumRef(CStr(Trim(p_RstDat!MAECFI_NUMREF)), Mid(p_RstDat!MAECFI_NUMREF, 1, 1))
          .ParagraphFormat.Alignment = wdAlignParagraphLeft
          .TypeParagraph
          .TypeParagraph
          
          .Font.Bold = False
          .TypeText "Señores"
          .ParagraphFormat.Alignment = wdAlignParagraphLeft
          .TypeParagraph
          
          .Font.Bold = True
          .TypeText "FONDO MIVIVIENDA S.A."
          .ParagraphFormat.Alignment = wdAlignParagraphLeft
          .TypeParagraph
          .Font.Bold = False
          .Font.Underline = True
          .TypeText "Presente"
          .Font.Underline = False
          .TypeText ". -"
          .ParagraphFormat.Alignment = wdAlignParagraphLeft
          .TypeParagraph
          .TypeParagraph
          
          .Font.Underline = False
          .TypeText "De nuestra consideración:"
          .ParagraphFormat.Alignment = wdAlignParagraphLeft
          .TypeParagraph
          .TypeParagraph
          
          .ParagraphFormat.Alignment = wdAlignParagraphJustify
          .TypeText "Por la presente y a solicitud de nuestro cliente, "
          
          .Font.Bold = True
          .TypeText p_RstDat!MAEPRV_RAZSOC
          
          .Font.Bold = False
          .TypeText ", otorgamos fianza irrevocable, solidaria, incondicionada y de realización automática, sin beneficio de excusión, ni división, hasta por la suma de S/ " & Format(p_RstDat!MAECFI_GARFIA, "###,###,###,##0.00")
          .TypeText " (" & r_str_MtoLtr & ")"
          .TypeText " a favor del Fondo MIVIVIENDA S.A., para garantizar la seriedad de la oferta, equivalente al 2.5% del valor total de la obra del proyecto presentado por nuestro cliente antes mencionado, en el marco de lo establecido en el Reglamento Operativo para acceder al Bono Familiar Habitacional - BFH, para la modalidad de Aplicación de "
          .TypeText r_str_Modali
          .TypeText " aprobado mediante Resolución Ministerial N°236-2018-VIVIENDA y sus modificatorias."
          .TypeParagraph
          .TypeParagraph
          
          .TypeText "Queda expresamente entendido por nosotros que esta fianza será ejecutada por la entidad a cuyo favor se emite, de conformidad con lo dispuesto en el artículo 1898° del Código Civil Vigente"
          .TypeParagraph
          .TypeParagraph
          
          .TypeText "La falta de renovación oportuna, antes de la fecha de vencimiento de esta garantía, es causal de ejecución de la misma."
          .TypeParagraph
          .TypeParagraph
          
          .TypeText "La ejecución de la presente Fianza, se realizará ante un incumplimiento de las obligaciones de la ET antes mencionada, salvo que dicho incumplimiento se deba a que el GFE no cumple con hacer el depósito del ahorro en la cuenta recaudadora del FMV, según lo estipulado en la Resolución Ministerial N°236-2018-VIVIENDA y sus modificatorias; por lo que "
          .TypeText "bastará el requerimiento de pago efectuado por ustedes mediante conducto notarial a nuestro domicilio  Av. Rivera Navarrete 645 -  San Isidro, y toda demora de nuestra parte para honrarla devengará un interés a la Tasa Libor de 90 días más 3%"
          .TypeText ". La Tasa Libor de 90 días será la del cierre al día anterior de pago, debiendo devengarse los intereses, sin necesidad de intimación, a partir del quinto día hábil luego de la fecha en que sea exigido el honramiento de esta Fianza."
          .TypeParagraph
          .TypeParagraph
          
          .TypeText "El plazo de vigencia de esta Fianza se iniciará el  " & Day(gf_FormatoFecha(CStr(p_RstDat!MAECFI_EMIFIA))) & " de " & fs_NomMes(Month(gf_FormatoFecha(CStr(p_RstDat!MAECFI_EMIFIA)))) & " de " & Year(gf_FormatoFecha(CStr(p_RstDat!MAECFI_EMIFIA)))
          .TypeText " hasta el " & Day(gf_FormatoFecha(CStr(p_RstDat!MAECFI_VTOFIA))) & " de " & fs_NomMes(Month(gf_FormatoFecha(CStr(p_RstDat!MAECFI_VTOFIA)))) & " del " & Year(gf_FormatoFecha(CStr(p_RstDat!MAECFI_VTOFIA))) & "."
          .TypeParagraph
          .TypeParagraph
          
          .TypeText "Atentamente,"
          .ParagraphFormat.Alignment = wdAlignParagraphJustify
          
          .TypeParagraph
          .TypeParagraph
          .TypeParagraph
          .TypeParagraph
          .TypeParagraph
          .TypeParagraph
          .TypeParagraph
'          .TypeParagraph
'          .TypeParagraph
'          .TypeParagraph
                      
          .Font.Bold = True
'            .ParagraphFormat.LineSpacing = LinesToPoints(1.5)
           .TypeText "__________________________                        "
          .TypeText "_________________________________________"
          .TypeParagraph
          .TypeParagraph
          
          .TypeText "  Julio Rodríguez Nakagawa"
          .TypeText "                                                  "
          .TypeText "Carlos Pacheco Caycho"
          .TypeParagraph
          .TypeParagraph
          .TypeText "GERENTE DE PRODUCCION"
          .TypeText "                           "
          .TypeText "GERENTE DE ADMINISTRACION Y TESORERIA"
          
      End With
        
      r_obj_Word.Visible = True
      Set r_obj_Word = Nothing
    End If
End Sub

Private Sub fs_GenWrd_Reforzamiento_Estructural(ByVal p_RstDat As ADODB.Recordset)
Dim r_obj_Word          As Word.Application
Dim r_str_ParEnt        As String
Dim r_str_ParDec        As String
Dim r_str_MtoLtr        As String

   If Not (p_RstDat.BOF And p_RstDat.EOF) Then
    
'      r_str_ParEnt = gf_Convertir_NumLet(CLng(Mid(Format(p_RstDat!MAECFI_GARFIA, "###,##0.00"), 1, InStr(Format(p_RstDat!MAECFI_GARFIA, "###,##0.00"), ".") - 1)))
      r_str_ParEnt = fs_NroEnLetras(CLng(Mid(Format(p_RstDat!MAECFI_GARFIA, "###,##0.00"), 1, InStr(Format(p_RstDat!MAECFI_GARFIA, "###,##0.00"), ".") - 1)))
      r_str_ParDec = Right(Format(p_RstDat!MAECFI_GARFIA, "###,##0.00"), 2)
      r_str_MtoLtr = r_str_ParEnt & "con " & r_str_ParDec & "/100 " & moddat_gf_Consulta_ParDes("204", CStr(p_RstDat!MAECFI_MONFIA))
       
      Set r_obj_Word = CreateObject("Word.Application")
      
      With r_obj_Word
         .Application.Documents.Add (moddat_g_str_RutCFi & "\" & moddat_g_str_NomCF1)
      End With
        
      With r_obj_Word.Selection
          .Font.Name = "Arial"
          .Font.Size = 10
'            .ParagraphFormat.LineSpacing = LinesToPoints(1)
          
          .WholeStory
          With .ParagraphFormat
'                .LeftIndent = CentimetersToPoints(0)
'                .RightIndent = CentimetersToPoints(0)
              .SpaceBefore = 0
              .SpaceBeforeAuto = False
              .SpaceAfter = 0
              .SpaceAfterAuto = False
              .LineSpacingRule = wdLineSpaceSingle
              .WidowControl = True
              .KeepWithNext = False
              .KeepTogether = False
              .PageBreakBefore = False
              .NoLineNumber = False
              .Hyphenation = True
'                .FirstLineIndent = CentimetersToPoints(0)
              .OutlineLevel = wdOutlineLevelBodyText
              .CharacterUnitLeftIndent = 0
              .CharacterUnitRightIndent = 0
              .CharacterUnitFirstLineIndent = 0
              .LineUnitBefore = 0
              .LineUnitAfter = 0
              .MirrorIndents = False
              .TextboxTightWrap = wdTightNone
          End With
      
          .Font.Bold = True
          .TypeText "MODELO DE FIANZA DE FIEL CUMPLIMIENTO"
          .ParagraphFormat.Alignment = wdAlignParagraphCenter
          .TypeParagraph
          .Font.Size = 9
          .TypeText "(Garantiza el desembolso del Bono de Reforzamiento Estructural)"
          .ParagraphFormat.Alignment = wdAlignParagraphCenter
          .TypeParagraph
          .TypeParagraph
          .TypeParagraph

          .Font.Size = 10
          .Font.Bold = False
          .TypeText "San Isidro, " & Day(gf_FormatoFecha(CStr(p_RstDat!MAECFI_EMIFIA))) & " de " & fs_NomMes(Month(gf_FormatoFecha(CStr(p_RstDat!MAECFI_EMIFIA)))) & " del " & Year(gf_FormatoFecha(CStr(p_RstDat!MAECFI_EMIFIA)))
          .ParagraphFormat.Alignment = wdAlignParagraphRight
          .TypeParagraph
          .TypeParagraph
          
          .Font.Bold = True
          .TypeText "Carta Fianza N°: " & gf_Formato_NumRef(CStr(Trim(p_RstDat!MAECFI_NUMREF)), Mid(p_RstDat!MAECFI_NUMREF, 1, 1))
          .ParagraphFormat.Alignment = wdAlignParagraphLeft
          .TypeParagraph
          .TypeParagraph
          
          .Font.Bold = False
          .TypeText "Señores"
          .ParagraphFormat.Alignment = wdAlignParagraphLeft
          .TypeParagraph
          
          .Font.Bold = True
          .TypeText "FONDO MIVIVIENDA S.A."
          .ParagraphFormat.Alignment = wdAlignParagraphLeft
          .TypeParagraph
          .Font.Bold = False
          .Font.Underline = True
          .TypeText "Presente"
          .Font.Underline = False
          .TypeText ". -"
          .ParagraphFormat.Alignment = wdAlignParagraphLeft
          .TypeParagraph
          .TypeParagraph
          
          .Font.Underline = False
          .TypeText "De nuestra consideración:"
          .ParagraphFormat.Alignment = wdAlignParagraphLeft
          .TypeParagraph
          .TypeParagraph
          
          .ParagraphFormat.Alignment = wdAlignParagraphJustify
          .TypeText "A solicitud de nuestro cliente "
          
          .Font.Bold = True
          .TypeText p_RstDat!MAEPRV_RAZSOC
          
          .Font.Bold = False
          .TypeText ", otorgamos fianza irrevocable, solidaria, incondicionada, sin beneficio de excusión y de realización automática a simple solicitud y a favor del Fondo MIVIVIENDA S.A, por la suma de S/ " & Format(p_RstDat!MAECFI_GARFIA, "###,###,###,##0.00")
          .TypeText " (" & r_str_MtoLtr & ")"
          .TypeText " para garantizar la conclusión de la obra de "
          .TypeText p_RstDat!CANT_BEN
          .TypeText " proyectos de Reforzamiento de Vivienda conforme al Expediente Técnico de Intervención aprobado por el Ministerio de Vivienda, Construcción y Saneamiento a cargo de la Entidad Técnica "
          .Font.Bold = True
          .TypeText p_RstDat!MAEPRV_RAZSOC
          .Font.Bold = False
          .TypeText ", con Código N° " & CStr(Trim(p_RstDat!MAECFI_CODETE)) & ", según el anexo que forma parte integrante del presente documento."
          .TypeParagraph
          .TypeParagraph
          
          .TypeText "La presente fianza rige desde el " & Day(gf_FormatoFecha(CStr(p_RstDat!MAECFI_EMIFIA))) & " de " & fs_NomMes(Month(gf_FormatoFecha(CStr(p_RstDat!MAECFI_EMIFIA)))) & " de " & Year(gf_FormatoFecha(CStr(p_RstDat!MAECFI_EMIFIA)))
          .TypeText " hasta el " & Day(gf_FormatoFecha(CStr(p_RstDat!MAECFI_VTOFIA))) & " de " & fs_NomMes(Month(gf_FormatoFecha(CStr(p_RstDat!MAECFI_VTOFIA)))) & " del " & Year(gf_FormatoFecha(CStr(p_RstDat!MAECFI_VTOFIA))) & "."
          .TypeParagraph
          .TypeParagraph
          
          .TypeText "La falta de renovación oportuna, antes de la fecha de vencimiento de esta garantía, es causal de ejecución de la misma."
          .TypeParagraph
          .TypeParagraph
          
          .TypeText "Queda expresamente entendido por nosotros que esta fianza será ejecutada por la entidad a cuyo favor se emite, de conformidad con lo dispuesto en el artículo 1898° del Código Civil vigente."
          .TypeParagraph
          .TypeParagraph
          
          .TypeText "Para honrar la presente Fianza bastará el simple requerimiento efectuado por ustedes mediante conducto notarial y toda demora de nuestra parte para honrarla devengará un interés a la Libor de 90 días más "
          .TypeText p_RstDat!MAECFI_TASFIA & "%"
          .TypeText ". La tasa Libor de 90 días será la de cierre al día anterior al pago, debiendo devengarse los intereses a partir del quinto día hábil luego de la fecha en que sea exigido el honramiento de esta Fianza."
          .TypeParagraph
          .TypeParagraph
          
          .TypeText "Atentamente,"
          .ParagraphFormat.Alignment = wdAlignParagraphJustify
          
          .TypeParagraph
          .TypeParagraph
          .TypeParagraph
          .TypeParagraph
          .TypeParagraph
          .TypeParagraph
          .TypeParagraph
          .TypeParagraph
'          .TypeParagraph
'          .TypeParagraph
                      
          .Font.Bold = True
'            .ParagraphFormat.LineSpacing = LinesToPoints(1.5)
          .TypeText "__________________________                        "
          .TypeText "_________________________________________"
          .TypeParagraph
          .TypeParagraph
          
          .TypeText "  Julio Rodríguez Nakagawa"
          .TypeText "                                                  "
          .TypeText "Carlos Pacheco Caycho"
          .TypeParagraph
          .TypeParagraph
          .TypeText "GERENTE DE PRODUCCION"
          .TypeText "                           "
          .TypeText "GERENTE DE ADMINISTRACION Y TESORERIA"
          
          
          .InsertNewPage
          .Font.Bold = True
          .TypeText "Anexo Carta Fianza N° " & gf_Formato_NumRef(CStr(Trim(p_RstDat!MAECFI_NUMREF)), Mid(p_RstDat!MAECFI_NUMREF, 1, 1))
          .ParagraphFormat.Alignment = wdAlignParagraphCenter
          .TypeParagraph
          .TypeParagraph
          .Font.Size = 9
          
          'Tabla
          If p_RstDat!CANT_BEN > 0 Then
              Call fs_docCrearTabla(8, CInt(p_RstDat!CANT_BEN), r_obj_Word, Trim(p_RstDat!MAECFI_NUMREF), CInt(p_RstDat!MAECFI_TIPREC))
              Call fs_docCombinarCeldas(5, r_obj_Word)
          End If
          
      End With
        
      r_obj_Word.Visible = True
      Set r_obj_Word = Nothing
    End If
End Sub

Private Sub fs_GenWrd_Renovacion(ByVal p_RstDat As ADODB.Recordset)
Dim r_obj_Word          As Word.Application
Dim r_str_ParEnt        As String
Dim r_str_ParDec        As String
Dim r_str_MtoLtr        As String
Dim r_str_Modali        As String
   
   r_str_ParEnt = ""
   r_str_ParDec = ""
   r_str_MtoLtr = ""
   
   If Not (p_RstDat.BOF And p_RstDat.EOF) Then
    
      If p_RstDat!MAECFI_NUMREN > 0 Then
        
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & "  SELECT MAECFI_NUMREF, MAEPRV_RAZSOC, MAECFI_CODPRD, MAECFI_CODSUB, MAECFI_CODMOD, MAECFI_EMIFIA, MAECFI_VTOFIA, MAECFI_CODPRY, MAECFI_CODETE, MAECFI_SITUAC, MAECFI_MONFIA, "
         g_str_Parame = g_str_Parame & "         MAECFI_GARFIA, MAECFI_TASFIA, MAECFI_NUMREN, MAECFI_REFANT, CASE WHEN MAECFI_TIPREC = 1 THEN 'BFH' ELSE 'AHORRO' END AS RECURSO, NVL(DATBEN_CODIGO,0) AS CANT_BEN "
         g_str_Parame = g_str_Parame & "    FROM TPR_MAECFI "
         g_str_Parame = g_str_Parame & "          INNER JOIN CNTBL_MAEPRV ON MAECFI_TIPDOC = MAEPRV_TIPDOC AND MAECFI_NUMDOC = MAEPRV_NUMDOC"
         g_str_Parame = g_str_Parame & "           LEFT JOIN TPR_DATBEN ON DATBEN_NUMREF = MAECFI_NUMREF"
         g_str_Parame = g_str_Parame & "   WHERE MAECFI_NUMREF = '" & p_RstDat!MAECFI_REFANT & "' "
         g_str_Parame = g_str_Parame & "     AND MAECFI_SITUAC = 4 "
           
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
             Exit Sub
         End If
           
         If g_rst_GenAux.BOF And g_rst_GenAux.EOF Then
            g_rst_GenAux.Close
            Set g_rst_GenAux = Nothing
            Exit Sub
         End If
                       
         If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
               
            If g_rst_GenAux!MAECFI_CODPRD = "026" And g_rst_GenAux!MAECFI_CODSUB = "002" And g_rst_GenAux!MAECFI_CODMOD = "006" Then
               r_str_Modali = "CSP Individual-Construcción en Sitio Propio"
            ElseIf g_rst_GenAux!MAECFI_CODPRD = "026" And g_rst_GenAux!MAECFI_CODSUB = "003" And g_rst_GenAux!MAECFI_CODMOD = "007" Then
               r_str_Modali = "MV-Mejoramiento de Vivienda"
            End If
        
            Set r_obj_Word = CreateObject("Word.Application")
                
            With r_obj_Word
               .Application.Documents.Add (moddat_g_str_RutCFi & "\" & moddat_g_str_NomCF1)
            End With
                 
            With r_obj_Word.Selection
               .Font.Name = "Arial"
               .Font.Size = 10
'                    .ParagraphFormat.LineSpacing = LinesToPoints(1)
                
               .WholeStory
               With .ParagraphFormat
'                         .LeftIndent = CentimetersToPoints(0)
'                         .RightIndent = CentimetersToPoints(0)
                    .SpaceBefore = 0
                    .SpaceBeforeAuto = False
                    .SpaceAfter = 0
                    .SpaceAfterAuto = False
                    .LineSpacingRule = wdLineSpaceSingle
                    .WidowControl = True
                    .KeepWithNext = False
                    .KeepTogether = False
                    .PageBreakBefore = False
                    .NoLineNumber = False
                    .Hyphenation = True
'                         .FirstLineIndent = CentimetersToPoints(0)
                    .OutlineLevel = wdOutlineLevelBodyText
                    .CharacterUnitLeftIndent = 0
                    .CharacterUnitRightIndent = 0
                    .CharacterUnitFirstLineIndent = 0
                    .LineUnitBefore = 0
                    .LineUnitAfter = 0
                    .MirrorIndents = False
                    .TextboxTightWrap = wdTightNone
               End With
               
               If g_rst_GenAux!MAECFI_CODMOD = "008" Then
                  .Font.Bold = True
                  .TypeText "CARTA FIANZA BANCARIA DE "
                  .ParagraphFormat.Alignment = wdAlignParagraphCenter
                  .TypeParagraph
                  .TypeText "GARANTÍA DE SERIEDAD DE LA OFERTA"
                  .ParagraphFormat.Alignment = wdAlignParagraphCenter
                  .TypeParagraph
                  .Font.Size = 9
               Else
                  .Font.Bold = True
                  .TypeText "CARTA FIANZA DE FIEL CUMPLIMIENTO"
                  .ParagraphFormat.Alignment = wdAlignParagraphCenter
                  .TypeParagraph
                  .Font.Size = 9
               
                  If r_str_Modali <> "" Then
                     If Not IsNull(p_RstDat!RECURSO) Then
                       .TypeText "(Garantiza el desembolso del " & IIf(p_RstDat!RECURSO = "BONO", "BFH", p_RstDat!RECURSO) & " de los beneficiarios - " & Mid(r_str_Modali, 1, InStr(r_str_Modali, "-") - 1) & ")"
                     Else
                       .TypeText "(Garantiza el desembolso del XXX de los beneficiarios)"
                     End If
                  Else
                     If Not IsNull(p_RstDat!RECURSO) Then
                       .TypeText "(Garantiza el desembolso del " & IIf(p_RstDat!RECURSO = "BONO", "BFH", p_RstDat!RECURSO) & " de los beneficiarios)"
                     Else
                       .TypeText "(Garantiza el desembolso del XXX de los beneficiarios)"
                     End If
                  End If
               End If
               .ParagraphFormat.Alignment = wdAlignParagraphCenter
               .TypeParagraph
               .TypeParagraph
               .TypeParagraph
               .TypeParagraph
               
               .Font.Bold = False
               .TypeText "San Isidro, " & Day(gf_FormatoFecha(CStr(p_RstDat!MAECFI_EMIFIA))) & " de " & fs_NomMes(Month(gf_FormatoFecha(CStr(p_RstDat!MAECFI_EMIFIA)))) & " del " & Year(gf_FormatoFecha(CStr(p_RstDat!MAECFI_EMIFIA)))
               .ParagraphFormat.Alignment = wdAlignParagraphRight
               .TypeParagraph
               .TypeParagraph
               
               .TypeText "Señores"
               .ParagraphFormat.Alignment = wdAlignParagraphLeft
               .TypeParagraph
               
               .Font.Bold = True
               .TypeText "FONDO MIVIVIENDA S.A."
               .ParagraphFormat.Alignment = wdAlignParagraphLeft
               .TypeParagraph
               .Font.Bold = False
               .Font.Underline = True
               .TypeText "Presente"
               .Font.Underline = False
               .TypeText ". -"
               .ParagraphFormat.Alignment = wdAlignParagraphLeft
               .TypeParagraph
               .TypeParagraph
               
               .Font.Size = 10
               .Font.Bold = True
               .TypeText "Carta Fianza N°: " & gf_Formato_NumRef(CStr(Trim(p_RstDat!MAECFI_NUMREF)), Mid(p_RstDat!MAECFI_NUMREF, 1, 1))
               .ParagraphFormat.Alignment = wdAlignParagraphCenter
               .TypeParagraph
               .TypeParagraph
               .TypeParagraph
               
               .Font.Bold = False
               .Font.Underline = False
'               .ParagraphFormat.Alignment = wdAlignParagraphJustify
               .TypeText "REF: Nueva Carta Fianza N° " & gf_Formato_NumRef(CStr(Trim(p_RstDat!MAECFI_REFANT)), Mid(p_RstDat!MAECFI_REFANT, 1, 1))
               .TypeText " del "
               .TypeText gf_FormatoFecha(CStr(g_rst_GenAux!MAECFI_EMIFIA)) & " por la suma de S/ " & Format(g_rst_GenAux!MAECFI_GARFIA, "###,###,###,##0.00")
               
'               r_str_ParEnt = gf_Convertir_NumLet(CLng(Mid(Format(g_rst_GenAux!MAECFI_GARFIA, "###,##0.00"), 1, InStr(Format(g_rst_GenAux!MAECFI_GARFIA, "###,##0.00"), ".") - 1)))
               r_str_ParEnt = fs_NroEnLetras(CLng(Mid(Format(g_rst_GenAux!MAECFI_GARFIA, "###,##0.00"), 1, InStr(Format(g_rst_GenAux!MAECFI_GARFIA, "###,##0.00"), ".") - 1)))
               r_str_ParDec = Right(Format(g_rst_GenAux!MAECFI_GARFIA, "###,##0.00"), 2)
               r_str_MtoLtr = r_str_ParEnt & "con " & r_str_ParDec & "/100 " & moddat_gf_Consulta_ParDes("204", CStr(g_rst_GenAux!MAECFI_MONFIA))
      
               .TypeText " (" & r_str_MtoLtr & ")" & " con vencimiento " & gf_FormatoFecha(CStr(g_rst_GenAux!MAECFI_VTOFIA))
               .TypeText " a favor de ustedes y a/c de:"
               .ParagraphFormat.Alignment = wdAlignParagraphJustify
               .TypeParagraph
               .TypeParagraph
               
               .Font.Bold = True
               .ParagraphFormat.Alignment = wdAlignParagraphCenter
               .TypeText p_RstDat!MAEPRV_RAZSOC
               .TypeParagraph
               .TypeText "_______________________________________________________________________________"
               .TypeParagraph
               .TypeParagraph
               
               .Font.Bold = False
               .ParagraphFormat.Alignment = wdAlignParagraphJustify
               .TypeText "Estimados señores:"
               .TypeParagraph
               .TypeParagraph
               
               .TypeText "Sírvanse tomar nota de que, a solicitud de nuestros garantizados hemos procedido a:"
               .TypeParagraph
               .TypeParagraph
'                   MsgBox .ChildShapeRange.AutoShapeType
               '.InsertSymbol 9644, "Arial", True
               r_str_ParEnt = ""
               r_str_ParDec = ""
               r_str_MtoLtr = ""
               
'               r_str_ParEnt = gf_Convertir_NumLet(CLng(Mid(Format(p_RstDat!MAECFI_GARFIA, "###,##0.00"), 1, InStr(Format(p_RstDat!MAECFI_GARFIA, "###,##0.00"), ".") - 1)))
               r_str_ParEnt = fs_NroEnLetras(CLng(Mid(Format(p_RstDat!MAECFI_GARFIA, "###,##0.00"), 1, InStr(Format(p_RstDat!MAECFI_GARFIA, "###,##0.00"), ".") - 1)))
               r_str_ParDec = Right(Format(p_RstDat!MAECFI_GARFIA, "###,##0.00"), 2)
               r_str_MtoLtr = r_str_ParEnt & "con " & r_str_ParDec & "/100 " & moddat_gf_Consulta_ParDes("204", CStr(p_RstDat!MAECFI_MONFIA))
               
               .TypeText "- " & "Renovar el monto de la Carta Fianza a S/ " & Format(p_RstDat!MAECFI_GARFIA, "###,###,###,##0.00")
               .TypeText " (" & r_str_MtoLtr & ")"
               .TypeParagraph
               '.InsertSymbol 9644, "Arial", True
               .TypeText "- " & "La nueva fecha de vencimiento es el "
               .TypeText gf_FormatoFecha(CStr(p_RstDat!MAECFI_VTOFIA))
               .TypeParagraph
               .TypeParagraph
               
               .TypeText "Manteniéndose vigente todas los demás términos y condiciones de la misma manera."
               .TypeParagraph
               .TypeParagraph
               
               .TypeText "Esta modificación estará vigente a partir de la recepción de la presente por parte de ustedes."
               .TypeParagraph
               .TypeParagraph
               .TypeParagraph
               
               .TypeText "Atentamente,"
               .ParagraphFormat.Alignment = wdAlignParagraphJustify
                
               .TypeParagraph
               .TypeParagraph
               .TypeParagraph
               .TypeParagraph
               .TypeParagraph
               .TypeParagraph
               .TypeParagraph
               .TypeParagraph
               .TypeParagraph
               .TypeParagraph
                
               .Font.Bold = True
'                    .ParagraphFormat.LineSpacing = LinesToPoints(1.5)
               .TypeText "__________________________                        "
               .TypeText "_________________________________________"
               .TypeParagraph
               .TypeParagraph
               
               .TypeText "  Julio Rodríguez Nakagawa"
               .TypeText "                                                  "
               .TypeText "Carlos Pacheco Caycho"
               .TypeParagraph
               .TypeParagraph
               .TypeText "GERENTE DE PRODUCCION"
               .TypeText "                           "
               .TypeText "GERENTE DE ADMINISTRACION Y TESORERIA"
                 '.Collapse Direction:=wdCollapseEnd ActiveDocument.Footnotes.Add Range:=Selection.Range , _ Text:="The Willow Tree, (Lone Creek Press, 1996)."
             End With
         End If
      End If
      r_obj_Word.Visible = True
      Set r_obj_Word = Nothing
    End If
End Sub

Private Sub cmd_ExpLiq_Click()
   If moddat_g_str_TipCre = 1 And moddat_g_int_TipPan = 1 Then
      Me.PopupMenu MnuPopUp
   End If
End Sub

Private Sub cmd_Histor_Click()
   'moddat_g_int_FlgGrb_1 = 7
   If grd_Listad.Row < 0 Then Exit Sub
   
   If moddat_g_str_TipCre = 1 And moddat_g_int_TipPan = 1 Then
      If Trim(grd_Listad.TextMatrix(grd_Listad.Row, 16)) = "VIGENTE" Or Trim(grd_Listad.TextMatrix(grd_Listad.Row, 16)) = "CANCELADO" Then 'fs_Validar = True Then
         
         moddat_g_str_DesIte = ""
         moddat_g_str_CodPrd = ""
   
         If grd_Listad.Rows > 0 Then
            moddat_g_str_NumFia = Trim(Replace(grd_Listad.TextMatrix(grd_Listad.Row, 2), "-", ""))
            moddat_g_str_DesIte = Trim(grd_Listad.TextMatrix(grd_Listad.Row, 17))
            moddat_g_str_DesObs = Trim(grd_Listad.TextMatrix(grd_Listad.Row, 16))
            moddat_g_str_CodPrd = Trim(grd_Listad.TextMatrix(grd_Listad.Row, 18))
            
            frm_Ges_TecPro_11.Show 1
         End If
      Else
         If grd_Listad.Rows > 0 Then
            If Trim(grd_Listad.TextMatrix(grd_Listad.Row, 0)) = "AD" Then
               MsgBox "La Adenda no se encuentra Vigente.", vbExclamation, modgen_g_str_NomPlt
            ElseIf Trim(grd_Listad.TextMatrix(grd_Listad.Row, 0)) = "CSO" Then
               MsgBox "La Carta Seriedad ferta no se encuentra Vigente.", vbExclamation, modgen_g_str_NomPlt
            Else
               MsgBox "La Carta Fianza no se encuentra Vigente.", vbExclamation, modgen_g_str_NomPlt
            End If
         End If
         Call gs_SetFocus(cmb_EstFia)
         Exit Sub
      End If
   Else
      MsgBox "Este Crédito no tiene histórico de renovación", vbExclamation, modgen_g_str_NomPlt
   End If
   
End Sub

Private Sub cmd_Limpia_Click()
   cmb_EstFia.ListIndex = -1
   Call gs_SetFocus(ipp_FecIni)
   cmb_EstFia.Enabled = True
   ipp_FecIni.Enabled = True
   ipp_FecFin.Enabled = True
End Sub

Private Sub cmd_Renova_Click()
   If grd_Listad.Rows = 0 And grd_Listad_Dir.Rows = 0 Then
      Exit Sub
   End If
   If moddat_g_str_TipCre = 1 And moddat_g_int_TipPan = 1 Then
      If CDbl(grd_Listad.TextMatrix(grd_Listad.Row, 9)) <> 0 Then 'And moddat_g_int_TipRec = 2
         MsgBox "No se puede renovar, la comisión tiene montos pendientes de pago.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
      
      moddat_g_int_FlgGrb_1 = 6
      If fs_Validar = True Then
         frm_Ges_TecPro_04.Show 1
      Else
         If grd_Listad.Rows > 0 Then
            If Trim(grd_Listad.TextMatrix(grd_Listad.Row, 0)) = "AD" Then
               MsgBox "La Adenda no se encuentra Vigente.", vbExclamation, modgen_g_str_NomPlt
            ElseIf Trim(grd_Listad.TextMatrix(grd_Listad.Row, 0)) = "CSO" Then
               MsgBox "La Carta Seriedad Oferta no se encuentra Vigente.", vbExclamation, modgen_g_str_NomPlt
            Else
               MsgBox "La Carta Fianza no se encuentra Vigente.", vbExclamation, modgen_g_str_NomPlt
            End If
         End If
         Call gs_SetFocus(cmb_EstFia)
         Exit Sub
      End If
   Else
      MsgBox "Este Crédito no se puede renovar", vbExclamation, modgen_g_str_NomPlt
   End If
End Sub

Private Sub grd_Listad_DblClick()
   moddat_g_str_TipCre = 1
   cmd_Editar_Click
End Sub

Private Sub cmd_Gestion_Click()
   moddat_g_str_DesObs = ""
   moddat_g_str_DesIte = ""
   moddat_g_str_NumFia = ""
   moddat_g_int_FlgGrb_1 = 5
   
   If moddat_g_str_TipCre = "" Then Exit Sub
   
   If moddat_g_str_TipCre = 1 And moddat_g_int_TipPan = 1 Then
      If grd_Listad.Rows > 0 Then
         moddat_g_str_NumFia = Trim(Replace(grd_Listad.TextMatrix(grd_Listad.Row, 2), "-", ""))
         moddat_g_str_DesIte = Trim(grd_Listad.TextMatrix(grd_Listad.Row, 17))
         moddat_g_str_DesObs = Trim(grd_Listad.TextMatrix(grd_Listad.Row, 16))
         
         moddat_g_str_CodPrd = Trim(grd_Listad.TextMatrix(grd_Listad.Row, 18))
         moddat_g_str_CodSub = Trim(grd_Listad.TextMatrix(grd_Listad.Row, 19))
         moddat_g_str_CodMod = Trim(grd_Listad.TextMatrix(grd_Listad.Row, 20))
         
      End If
   ElseIf moddat_g_str_TipCre = 2 And moddat_g_int_TipPan = 0 Then
      If grd_Listad_Dir.Rows > 0 Then
         moddat_g_str_NumFia = Trim(Replace(grd_Listad_Dir.TextMatrix(grd_Listad_Dir.Row, 1), "-", ""))
         moddat_g_str_DesIte = Trim(Replace(grd_Listad_Dir.TextMatrix(grd_Listad_Dir.Row, 1), "-", ""))
         moddat_g_str_DesObs = Trim(grd_Listad_Dir.TextMatrix(grd_Listad_Dir.Row, 13))
         
         moddat_g_str_CodPrd = Trim(grd_Listad_Dir.TextMatrix(grd_Listad_Dir.Row, 14))
         moddat_g_str_CodSub = Trim(grd_Listad_Dir.TextMatrix(grd_Listad_Dir.Row, 15))
         moddat_g_str_CodMod = Trim(grd_Listad_Dir.TextMatrix(grd_Listad_Dir.Row, 16))
      End If
   End If
   If moddat_g_str_NumFia <> "" Then
      frm_Ges_TecPro_06.Show 1
   End If
End Sub

Private Function fs_Validar() As Boolean
   fs_Validar = False

   If moddat_g_str_TipCre = 1 And moddat_g_int_TipPan = 1 Then
      If grd_Listad.Rows > 0 Then
         moddat_g_str_NumFia = Trim(Replace(grd_Listad.TextMatrix(grd_Listad.Row, 2), "-", ""))
         moddat_g_str_DesIte = Trim(grd_Listad.TextMatrix(grd_Listad.Row, 17))
         If moddat_g_str_NumFia <> "" Then
            fs_Validar = True
         Else
            fs_Validar = False
         End If
         If Trim(grd_Listad.TextMatrix(grd_Listad.Row, 16)) <> "VIGENTE" Then
            fs_Validar = False
         End If
      End If
   ElseIf moddat_g_str_TipCre = 2 And moddat_g_int_TipPan = 0 Then
      If grd_Listad_Dir.Rows > 0 Then
         moddat_g_str_NumFia = Trim(Replace(grd_Listad_Dir.TextMatrix(grd_Listad_Dir.Row, 1), "-", ""))
         moddat_g_str_DesIte = Trim(Replace(grd_Listad_Dir.TextMatrix(grd_Listad_Dir.Row, 1), "-", ""))
         If moddat_g_str_NumFia <> "" Then
            fs_Validar = True
         Else
            fs_Validar = False
         End If
         If Trim(grd_Listad_Dir.TextMatrix(grd_Listad_Dir.Row, 13)) <> "VIGENTE" Then
            fs_Validar = False
         End If
      End If
   End If
End Function

Private Sub cmd_Salida_Click()
   frm_Ges_TecPro_01.fs_Buscar
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Buscar_Creditos_Indirectos
   Call fs_Buscar_Creditos_Directos
   
   If moddat_g_int_NumCuo = 0 Then
      Call fs_Activa(True)
      cmb_EstFia.Enabled = True
      ipp_FecIni.Enabled = True
      ipp_FecFin.Enabled = True
      'Call fs_Limpia
   Else
      Call fs_Activa(False)
      cmd_Agrega.Enabled = True
      cmd_Buscar.Enabled = True
      cmd_Limpia.Enabled = True
   End If
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
Dim r_str_FecAux As String
   
   'Estado
   Call moddat_gs_Carga_LisIte_Combo(cmb_EstFia, 1, "529")
   cmb_EstFia.ListIndex = 0
   
   'Año y Mes Activo
   r_str_FecAux = ""
   moddat_g_str_FecIni = ""
   moddat_g_str_FecFin = ""
   moddat_g_str_CodAno = 0
   moddat_g_str_CodMes = 0
   moddat_g_str_DesObs = ""
   moddat_g_str_DesIte = ""
   moddat_g_str_NumFia = ""
   moddat_g_int_TipPan = 1
   moddat_g_str_TipCre = 0
   
'   ipp_FecIni.Text = Format(date, "dd/mm/yyyy")
'   ipp_FecFin.Text = Format(date, "dd/mm/yyyy")
'   ipp_FecIni.DateMin = Format(CDate(DateAdd("M", -8, ipp_FecIni.Text)), "yyyymmdd") '-6
'   ipp_FecIni.Text = Format(CDate(DateAdd("M", -8, ipp_FecIni.Text)), "DD/MM/YYYY") '-6
'   ipp_FecFin.DateMax = Format(date, "yyyymmdd")
    
   Call moddat_gf_ConsultaPerMesActivo("000001", 1, moddat_g_str_FecIni, moddat_g_str_FecFin, moddat_g_str_CodMes, moddat_g_str_CodAno)
   r_str_FecAux = DateAdd("m", 1, "01/" & Format(moddat_g_str_CodMes, "00") & "/" & moddat_g_str_CodAno)
   moddat_g_str_FecFin = DateAdd("d", -1, r_str_FecAux)
     
'      ipp_FecIni.DateMin = Format(CDate(DateAdd("M", -12, moddat_g_str_FecIni)), "yyyymmdd")
   ipp_FecIni.Text = Format(CDate(DateAdd("M", -24, moddat_g_str_FecIni)), "DD/MM/YYYY")
   'ipp_FecFin.DateMax = Format(moddat_g_str_FecFin, "yyyymmdd")
   ipp_FecFin.Text = moddat_g_str_FecFin
   
   'Créditos Indirectos
   grd_Listad.ColWidth(0) = 450
   grd_Listad.ColWidth(1) = 660     'NUMERO FMV
   grd_Listad.ColWidth(2) = 1125    'NUMERO CF
   grd_Listad.ColWidth(3) = 990     'FECHA EMISION
   grd_Listad.ColWidth(4) = 990     'FECHA VENCIMIENTO
   grd_Listad.ColWidth(5) = 1020    'IMPORTE CF
   grd_Listad.ColWidth(6) = 1020    'GARANTIZADO
   grd_Listad.ColWidth(7) = 1020    'COMISION IMPORTE
   grd_Listad.ColWidth(8) = 1020    'COMISION PAGADO
   grd_Listad.ColWidth(9) = 1020    'COMISION SALDO
   grd_Listad.ColWidth(10) = 1020   'FONDOS IMPORTE
   grd_Listad.ColWidth(11) = 1020   'FONDOS RECIBIDO
   grd_Listad.ColWidth(12) = 1020   'FONDOS SALDO
   grd_Listad.ColWidth(13) = 1020   'DESEMBOLSO IMPORTE
   grd_Listad.ColWidth(14) = 1020   'DESEMBOLSO PAGADO
   grd_Listad.ColWidth(15) = 1020   'DESEMBOLSO SALDO
   grd_Listad.ColWidth(16) = 1200   'SITUACION
   grd_Listad.ColWidth(17) = 0      'NUMERO DE REFERENCIA
   grd_Listad.ColWidth(18) = 0      'PRODUCTO
   grd_Listad.ColWidth(19) = 0      'SUB_PRODUCTO
   grd_Listad.ColWidth(20) = 0      'MODALIDAD
   grd_Listad.ColWidth(21) = 0      'TIPO RECURSO
   
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignCenterCenter
   grd_Listad.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad.ColAlignment(4) = flexAlignCenterCenter
   grd_Listad.ColAlignment(5) = flexAlignRightCenter
   grd_Listad.ColAlignment(6) = flexAlignRightCenter
   grd_Listad.ColAlignment(7) = flexAlignRightCenter
   grd_Listad.ColAlignment(8) = flexAlignRightCenter
   grd_Listad.ColAlignment(9) = flexAlignRightCenter
   grd_Listad.ColAlignment(10) = flexAlignRightCenter
   grd_Listad.ColAlignment(11) = flexAlignRightCenter
   grd_Listad.ColAlignment(12) = flexAlignRightCenter
   grd_Listad.ColAlignment(13) = flexAlignRightCenter
   grd_Listad.ColAlignment(14) = flexAlignRightCenter
   grd_Listad.ColAlignment(15) = flexAlignRightCenter
   grd_Listad.ColAlignment(16) = flexAlignCenterCenter
   Call gs_LimpiaGrid(grd_Listad)
   
   'Créditos Directos
   grd_Listad_Dir.ColWidth(0) = 520
   grd_Listad_Dir.ColWidth(1) = 1155    'NUMERO CF
   grd_Listad_Dir.ColWidth(2) = 1050    'FECHA EMISION
   grd_Listad_Dir.ColWidth(3) = 1050    'FECHA VENCIMIENTO
   grd_Listad_Dir.ColWidth(4) = 1050    'TASA INTERES
   grd_Listad_Dir.ColWidth(5) = 1050    'TASA MORATORIA
   grd_Listad_Dir.ColWidth(6) = 1350    'IMPORTE CF
   grd_Listad_Dir.ColWidth(7) = 1320    'COMISION IMPORTE
   grd_Listad_Dir.ColWidth(8) = 1320    'COMISION PAGADO
   grd_Listad_Dir.ColWidth(9) = 1320    'COMISION SALDO
   grd_Listad_Dir.ColWidth(10) = 1320   'DESEMBOLSO IMPORTE
   grd_Listad_Dir.ColWidth(11) = 1320   'DESEMBOLSO PAGADO
   grd_Listad_Dir.ColWidth(12) = 1320   'DESEMBOLSO SALDO
   grd_Listad_Dir.ColWidth(13) = 1350   'SITUACION
   
   grd_Listad_Dir.ColWidth(14) = 0      'PRODUCTO
   grd_Listad_Dir.ColWidth(15) = 0      'SUBPRODUCTO
   grd_Listad_Dir.ColWidth(16) = 0      'MODALIDAD
      
   grd_Listad_Dir.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad_Dir.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad_Dir.ColAlignment(2) = flexAlignCenterCenter
   grd_Listad_Dir.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad_Dir.ColAlignment(4) = flexAlignCenterCenter
   grd_Listad_Dir.ColAlignment(5) = flexAlignCenterCenter
   grd_Listad_Dir.ColAlignment(6) = flexAlignRightCenter
   grd_Listad_Dir.ColAlignment(7) = flexAlignRightCenter
   grd_Listad_Dir.ColAlignment(8) = flexAlignRightCenter
   grd_Listad_Dir.ColAlignment(9) = flexAlignRightCenter
   grd_Listad_Dir.ColAlignment(10) = flexAlignRightCenter
   grd_Listad_Dir.ColAlignment(11) = flexAlignRightCenter
   grd_Listad_Dir.ColAlignment(12) = flexAlignRightCenter
   grd_Listad_Dir.ColAlignment(13) = flexAlignCenterCenter

   Call gs_LimpiaGrid(grd_Listad_Dir)
   
   pnl_TipDoc.Caption = moddat_gf_Consulta_ParDes("118", moddat_g_int_TipDoc)
   pnl_NroDoc.Caption = moddat_g_str_NumDoc
   pnl_RazSoc.Caption = Trim(moddat_g_str_NomCli)
   pnl_TipEmp.Caption = moddat_g_str_Descri
End Sub

Private Sub fs_Limpia()
   Call gs_LimpiaGrid(grd_Listad)
End Sub

Public Sub fs_Buscar_Creditos_Directos()

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "   SELECT REFERENCIA           , FECHA_EMISION     , FECHA_VENCIMIENTO                                           , VALOR_CARTA_FIANZA  , "
   g_str_Parame = g_str_Parame & "          IMPORTE_COMISION     , PAGADO_COMISION   , (IMPORTE_COMISION - PAGADO_COMISION) SALDO_COMISION         , EXTORNO_COMISION    , "
   g_str_Parame = g_str_Parame & "          IMPORTE_DESEMBOLSADO , PAGADO_DESEMBOLSO , (IMPORTE_DESEMBOLSADO - PAGADO_DESEMBOLSO) SALDO_DESEMBOLSO ,                       "
   g_str_Parame = g_str_Parame & "          IMPORTE_GARANTIA     , PAGADO_GARANTIA   , (IMPORTE_GARANTIA - PAGADO_GARANTIA) SALDO_GARANTIA         ,                       "
   g_str_Parame = g_str_Parame & "          SITUACION            , MODALIDAD         , TASA_INTERES                                                , TASA_MORATORIA      , "
   g_str_Parame = g_str_Parame & "          PRODUCTO             , SUB_PRODUCTO "
   g_str_Parame = g_str_Parame & "    FROM( "
   g_str_Parame = g_str_Parame & "          SELECT A.MAECFI_NUMREF                                                      REFERENCIA,          "
   g_str_Parame = g_str_Parame & "                 A.MAECFI_EMIFIA                                                      FECHA_EMISION,       "
   g_str_Parame = g_str_Parame & "                 A.MAECFI_VTOFIA                                                      FECHA_VENCIMIENTO,   "
   g_str_Parame = g_str_Parame & "                 A.MAECFI_PORTEA                                                      TASA_INTERES,        "
   g_str_Parame = g_str_Parame & "                 A.MAECFI_TASFIA                                                      TASA_MORATORIA,      "
   g_str_Parame = g_str_Parame & "                 A.MAECFI_CODPRD                                                      PRODUCTO,            "
   g_str_Parame = g_str_Parame & "                 A.MAECFI_CODSUB                                                      SUB_PRODUCTO,        "
   g_str_Parame = g_str_Parame & "                 A.MAECFI_IMPFIA                                                      VALOR_CARTA_FIANZA,  "
   g_str_Parame = g_str_Parame & "                 MAECFI_COMFIA                                                        IMPORTE_COMISION,    "
   
   g_str_Parame = g_str_Parame & "                 NVL((SELECT NVL(SUM(MAERDE_IMPORT),0) "
   g_str_Parame = g_str_Parame & "                        FROM TPR_MAERDE B "
   g_str_Parame = g_str_Parame & "                       WHERE B.MAERDE_NUMREF = A.MAECFI_NUMREF AND (MAERDE_CODIGO = 1 OR MAERDE_CODIGO = 2 OR MAERDE_CODIGO = 13) "
   g_str_Parame = g_str_Parame & "                         AND B.MAERDE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                       GROUP BY MAERDE_NUMREF),0)                                     PAGADO_COMISION, "
   
   g_str_Parame = g_str_Parame & "                 NVL((SELECT NVL(SUM(MAERDE_IMPORT),0) "
   g_str_Parame = g_str_Parame & "                        FROM TPR_MAERDE B "
   g_str_Parame = g_str_Parame & "                       WHERE B.MAERDE_NUMREF = A.MAECFI_NUMREF AND MAERDE_CODIGO = 13 "
   g_str_Parame = g_str_Parame & "                         AND B.MAERDE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                       GROUP BY MAERDE_NUMREF),0)                                     COMISION_DEPOSITO, "
   
   g_str_Parame = g_str_Parame & "                 NVL((SELECT NVL(SUM(MAERDE_IMPORT),0) "
   g_str_Parame = g_str_Parame & "                        FROM TPR_MAERDE B "
   g_str_Parame = g_str_Parame & "                       WHERE B.MAERDE_NUMREF = A.MAECFI_NUMREF AND MAERDE_CODIGO = 15 "
   g_str_Parame = g_str_Parame & "                         AND B.MAERDE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                       GROUP BY MAERDE_NUMREF),0)                                     EXTORNO_COMISION,"
       
   'SI HA SIDO RENOVADA Y ESTÁ VIGENTE
   g_str_Parame = g_str_Parame & "                 CASE WHEN A.MAECFI_NUMREN > 0 THEN "
   
   g_str_Parame = g_str_Parame & "                              NVL((SELECT NVL(SUM(MAERDE_IMPORT),0) "
   g_str_Parame = g_str_Parame & "                                     FROM TPR_MAERDE B "
   g_str_Parame = g_str_Parame & "                                    WHERE B.MAERDE_NUMREF = A.MAECFI_NUMREF AND (MAERDE_CODIGO = 2 OR MAERDE_CODIGO = 10 OR MAERDE_CODIGO = 3 OR MAERDE_CODIGO = 11 OR MAERDE_CODIGO = 12 OR MAERDE_CODIGO = 14 OR MAERDE_CODIGO = 17 OR MAERDE_CODIGO = 19)  "
   g_str_Parame = g_str_Parame & "                                      AND B.MAERDE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                                    GROUP BY MAERDE_NUMREF),0) "
   g_str_Parame = g_str_Parame & "                 ELSE                   "
   g_str_Parame = g_str_Parame & "                              NVL((SELECT NVL(SUM(MAERDE_IMPORT),0) "
   g_str_Parame = g_str_Parame & "                                     FROM TPR_MAERDE B "
   g_str_Parame = g_str_Parame & "                                    WHERE B.MAERDE_NUMREF = A.MAECFI_NUMREF AND (MAERDE_CODIGO = 2 OR MAERDE_CODIGO = 10 OR MAERDE_CODIGO = 3 OR MAERDE_CODIGO = 11 OR MAERDE_CODIGO = 12 OR MAERDE_CODIGO = 14 OR MAERDE_CODIGO = 19 )  "
   g_str_Parame = g_str_Parame & "                                      AND B.MAERDE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                                    GROUP BY MAERDE_NUMREF),0)                                     "
   g_str_Parame = g_str_Parame & "                 END                                                                  IMPORTE_DESEMBOLSADO, "
   
   g_str_Parame = g_str_Parame & "                 NVL((SELECT NVL(SUM(MAERDE_IMPORT),0)"
   g_str_Parame = g_str_Parame & "                        FROM TPR_MAERDE B"
   g_str_Parame = g_str_Parame & "                       WHERE B.MAERDE_NUMREF = A.MAECFI_NUMREF AND (MAERDE_CODIGO = 1 Or MAERDE_CODIGO = 2 Or MAERDE_CODIGO = 4 Or MAERDE_CODIGO = 5 OR MAERDE_CODIGO = 10 OR MAERDE_CODIGO = 6 OR MAERDE_CODIGO = 14 OR MAERDE_CODIGO = 18)"
   g_str_Parame = g_str_Parame & "                         AND B.MAERDE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                       GROUP BY MAERDE_NUMREF),0)                                     PAGADO_DESEMBOLSO,"
   
   g_str_Parame = g_str_Parame & "                 NVL((SELECT NVL(SUM(NVL(MAEGAR_MTOGAR_INM,0) + NVL(MAEGAR_MTOGAR_ES1,0) + NVL(MAEGAR_MTOGAR_ES2,0) + NVL(MAEGAR_MTOGAR_DE1,0) + NVL(MAEGAR_MTOGAR_DE2,0)),0) "
   g_str_Parame = g_str_Parame & "                        FROM TPR_MAEGAR "
   g_str_Parame = g_str_Parame & "                       WHERE MAEGAR_NUMREF = A.MAECFI_NUMREF "
   g_str_Parame = g_str_Parame & "                       GROUP BY MAEGAR_NUMREF),0)                                     IMPORTE_GARANTIA, "
   
   g_str_Parame = g_str_Parame & "                 NVL((SELECT NVL(SUM(MAERDE_IMPORT),0) "
   g_str_Parame = g_str_Parame & "                        FROM TPR_MAERDE B "
   g_str_Parame = g_str_Parame & "                       WHERE B.MAERDE_NUMREF = A.MAECFI_NUMREF AND MAERDE_CODIGO = 6 "
   g_str_Parame = g_str_Parame & "                         AND B.MAERDE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                       GROUP BY MAERDE_NUMREF),0)                                     PAGADO_GARANTIA, "
     
   g_str_Parame = g_str_Parame & "                 TRIM(PARDES_DESCRI)                                                  SITUACION, "
   g_str_Parame = g_str_Parame & "                 MAECFI_CODMOD                                                        MODALIDAD "
   
   g_str_Parame = g_str_Parame & "            FROM TPR_MAECFI A "
   g_str_Parame = g_str_Parame & "                 INNER JOIN CNTBL_MAEPRV ON MAEPRV_TIPDOC = A.MAECFI_TIPDOC AND MAEPRV_NUMDOC = MAECFI_NUMDOC "
   g_str_Parame = g_str_Parame & "                 INNER JOIN MNT_PARDES   ON PARDES_CODGRP = '529' AND PARDES_CODITE = A.MAECFI_SITUAC "
   g_str_Parame = g_str_Parame & "           WHERE A.MAECFI_TIPDOC = " & CStr(moddat_g_int_TipDoc) & "  AND A.MAECFI_NUMDOC = '" & CStr(moddat_g_str_NumDoc) & "' "
   g_str_Parame = g_str_Parame & "             AND A.MAECFI_CODPRD = '008' "
   
   If Me.cmb_EstFia.ListIndex <> -1 Then
        g_str_Parame = g_str_Parame & "        AND A.MAECFI_SITUAC = " & cmb_EstFia.ItemData(cmb_EstFia.ListIndex) & " "
   End If
    
   g_str_Parame = g_str_Parame & "        ) "
   
   g_str_Parame = g_str_Parame & "     WHERE FECHA_EMISION >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " "
   g_str_Parame = g_str_Parame & "       AND FECHA_EMISION <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & " "
   g_str_Parame = g_str_Parame & "     ORDER BY REFERENCIA"

            
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   grd_Listad_Dir.Redraw = False
   Call gs_LimpiaGrid(grd_Listad_Dir)
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      grd_Listad_Dir.Redraw = True
      
      If grd_Listad.Rows = 0 And grd_Listad_Dir.Rows = 0 Then
         Call fs_Activa(True)
      End If
      
      cmd_Buscar.Enabled = True
      cmd_Limpia.Enabled = True
      cmd_Agrega.Enabled = True
      Exit Sub
   End If
   
'   Call fs_Obtiene_Cabecera
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      cmd_Histor.Enabled = True
      cmd_Gestion.Enabled = True
      cmd_DatBen.Enabled = True
      cmd_ExpExc.Enabled = True
      cmd_ExpLiq.Enabled = True
      
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
      
         grd_Listad_Dir.Rows = grd_Listad_Dir.Rows + 1
         grd_Listad_Dir.Row = grd_Listad_Dir.Rows - 1
         grd_Listad_Dir.BackColor = &H80000005
         grd_Listad_Dir.ForeColor = &H80000008
         
         grd_Listad_Dir.Col = 0
         grd_Listad_Dir.Text = fs_Obtener_Tipo(CStr(Trim(g_rst_Princi!REFERENCIA)), "008", Trim(g_rst_Princi!MODALIDAD))
                 
         grd_Listad_Dir.Col = 1
         grd_Listad_Dir.Text = gf_Formato_NumRef(CStr(Trim(g_rst_Princi!REFERENCIA)), 1)
        
         grd_Listad_Dir.Col = 2
         grd_Listad_Dir.Text = Format(gf_FormatoFecha(CStr(g_rst_Princi!FECHA_EMISION)), "dd/mm/yyyy")
         
         grd_Listad_Dir.Col = 3
         grd_Listad_Dir.Text = Format(gf_FormatoFecha(CStr(g_rst_Princi!FECHA_VENCIMIENTO)), "dd/mm/yyyy")
         
         grd_Listad_Dir.Col = 4
         grd_Listad_Dir.Text = Format(CStr(g_rst_Princi!TASA_INTERES))
         
         grd_Listad_Dir.Col = 5
         grd_Listad_Dir.Text = Format(CStr(g_rst_Princi!TASA_MORATORIA))
         
         grd_Listad_Dir.Col = 6
         grd_Listad_Dir.Text = Format(CStr(g_rst_Princi!VALOR_CARTA_FIANZA), "###,###,###,##0.00")
         
         grd_Listad_Dir.Col = 7
         grd_Listad_Dir.Text = Format(CStr(g_rst_Princi!IMPORTE_COMISION), "###,###,###,##0.00")
         
         grd_Listad_Dir.Col = 8
         grd_Listad_Dir.Text = Format(CStr(g_rst_Princi!PAGADO_COMISION), "###,###,###,##0.00")
         
         grd_Listad_Dir.Col = 9
         grd_Listad_Dir.Text = Format(CStr(g_rst_Princi!SALDO_COMISION), "###,###,###,##0.00")
                 
         grd_Listad_Dir.Col = 10
         grd_Listad_Dir.Text = Format(CStr(g_rst_Princi!IMPORTE_DESEMBOLSADO), "###,###,###,##0.00")
         
         grd_Listad_Dir.Col = 11
         grd_Listad_Dir.Text = Format(CStr(g_rst_Princi!PAGADO_DESEMBOLSO), "###,###,###,##0.00") '- g_rst_Princi!DEVOLUCION_GARANTIA
         
         grd_Listad_Dir.Col = 12
         grd_Listad_Dir.Text = Format(CStr(g_rst_Princi!SALDO_DESEMBOLSO), "###,###,###,##0.00") '+ g_rst_Princi!DEVOLUCION_GARANTIA
         
         grd_Listad_Dir.Col = 13
         grd_Listad_Dir.Text = CStr(g_rst_Princi!SITUACION)
         
         grd_Listad_Dir.Col = 14
         grd_Listad_Dir.Text = CStr(g_rst_Princi!PRODUCTO)
         
         grd_Listad_Dir.Col = 15
         If IsNull(g_rst_Princi!SUB_PRODUCTO) Then
            grd_Listad_Dir.Text = ""
         Else
            grd_Listad_Dir.Text = CStr(g_rst_Princi!SUB_PRODUCTO)
         End If
         
         grd_Listad_Dir.Col = 16
         grd_Listad_Dir.Text = CStr(g_rst_Princi!MODALIDAD)
                  
         g_rst_Princi.MoveNext
      Loop
   End If
'   grd_Listad_DIR.FixedRows = 2
   grd_Listad_Dir.Redraw = True
   Call gs_UbiIniGrid(grd_Listad_Dir)
End Sub

Public Sub fs_Buscar_Creditos_Indirectos()

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "   SELECT REFERENCIA           , FECHA_EMISION     , FECHA_VENCIMIENTO    , VALOR_CARTA_FIANZA                   , GARANTIZADO         , "
   g_str_Parame = g_str_Parame & "          IMPORTE_COMISION     , PAGADO_COMISION   , (IMPORTE_COMISION - PAGADO_COMISION) SALDO_COMISION         , EXTORNO_COMISION    , "
   g_str_Parame = g_str_Parame & "          IMPORTE_FONDOS       , RECIBIDO_FONDOS   , (IMPORTE_FONDOS - RECIBIDO_FONDOS) SALDO_FONDOS             ,                       "
   g_str_Parame = g_str_Parame & "          IMPORTE_DESEMBOLSADO , PAGADO_DESEMBOLSO , (IMPORTE_DESEMBOLSADO - PAGADO_DESEMBOLSO) SALDO_DESEMBOLSO , PRODUCTO            , "
   g_str_Parame = g_str_Parame & "          IMPORTE_GARANTIA     , PAGADO_GARANTIA   , (IMPORTE_GARANTIA - PAGADO_GARANTIA) SALDO_GARANTIA         , RETENCION_GARANTIA  , DEVOLUCION_GARANTIA,"
   g_str_Parame = g_str_Parame & "          SITUACION            , MODALIDAD         , REFERENCIA_ANTERIOR                                         , REF_FMV             , SUB_PRODUCTO , "
   g_str_Parame = g_str_Parame & "          TIPO_RECURSO         , DEVOLUCION_FONDO_CLIENTE "
   g_str_Parame = g_str_Parame & "    FROM( "
   g_str_Parame = g_str_Parame & "          SELECT A.MAECFI_NUMREF                                                      REFERENCIA,          "
   g_str_Parame = g_str_Parame & "                 A.MAECFI_CODPRD                                                      PRODUCTO,            "
   g_str_Parame = g_str_Parame & "                 A.MAECFI_CODSUB                                                      SUB_PRODUCTO,        "
   g_str_Parame = g_str_Parame & "                 A.MAECFI_NUMADE                                                      REF_FMV,             "
   g_str_Parame = g_str_Parame & "                 A.MAECFI_EMIFIA                                                      FECHA_EMISION,       "
   g_str_Parame = g_str_Parame & "                 A.MAECFI_VTOFIA                                                      FECHA_VENCIMIENTO,   "
   g_str_Parame = g_str_Parame & "                 A.MAECFI_IMPFIA                                                      VALOR_CARTA_FIANZA,  "
   g_str_Parame = g_str_Parame & "                 A.MAECFI_GARFIA                                                      GARANTIZADO,         "
   g_str_Parame = g_str_Parame & "                 A.MAECFI_NUMANT                                                      REFERENCIA_ANTERIOR, "
   g_str_Parame = g_str_Parame & "                 A.MAECFI_COMFIA                                                      IMPORTE_COMISION,    "
   g_str_Parame = g_str_Parame & "                 A.MAECFI_TIPREC                                                      TIPO_RECURSO,        "
   
   g_str_Parame = g_str_Parame & "                 NVL((SELECT NVL(SUM(MAERDE_IMPORT),0) "
   g_str_Parame = g_str_Parame & "                        FROM TPR_MAERDE B "
   g_str_Parame = g_str_Parame & "                       WHERE B.MAERDE_NUMREF = A.MAECFI_NUMREF AND (MAERDE_CODIGO = 1 OR MAERDE_CODIGO = 2 OR MAERDE_CODIGO = 13) "
   g_str_Parame = g_str_Parame & "                         AND B.MAERDE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                       GROUP BY MAERDE_NUMREF),0)                                     PAGADO_COMISION, "
   
   g_str_Parame = g_str_Parame & "                 NVL((SELECT NVL(SUM(MAERDE_IMPORT),0) "
   g_str_Parame = g_str_Parame & "                        FROM TPR_MAERDE B "
   g_str_Parame = g_str_Parame & "                       WHERE B.MAERDE_NUMREF = A.MAECFI_NUMREF AND MAERDE_CODIGO = 13 "
   g_str_Parame = g_str_Parame & "                         AND B.MAERDE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                       GROUP BY MAERDE_NUMREF),0)                                     COMISION_DEPOSITO, "
   
   g_str_Parame = g_str_Parame & "                 NVL((SELECT NVL(SUM(MAERDE_IMPORT),0) "
   g_str_Parame = g_str_Parame & "                        FROM TPR_MAERDE B "
   g_str_Parame = g_str_Parame & "                       WHERE B.MAERDE_NUMREF = A.MAECFI_NUMREF AND MAERDE_CODIGO = 15 "
   g_str_Parame = g_str_Parame & "                         AND B.MAERDE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                       GROUP BY MAERDE_NUMREF),0)                                     EXTORNO_COMISION,"
   
   'SI HA SIDO RENOVADA Y ESTÁ VIGENTE
   g_str_Parame = g_str_Parame & "                 CASE WHEN A.MAECFI_NUMREN > 0 THEN "
   
   g_str_Parame = g_str_Parame & "                              NVL(( SELECT NVL(SUM(MAERDE_IMPORT),0) "
   g_str_Parame = g_str_Parame & "                                      FROM TPR_MAERDE B "
   g_str_Parame = g_str_Parame & "                                     WHERE B.MAERDE_NUMREF = A.MAECFI_NUMREF AND MAERDE_CODIGO = 16   "
   g_str_Parame = g_str_Parame & "                                       AND B.MAERDE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                                     GROUP BY MAERDE_NUMREF),0)                                       "
   
   g_str_Parame = g_str_Parame & "                 ELSE "
   g_str_Parame = g_str_Parame & "                              A.MAECFI_IMPFIA                                                         "
   g_str_Parame = g_str_Parame & "                 END                                                                  IMPORTE_FONDOS, "
   
   g_str_Parame = g_str_Parame & "                 NVL((SELECT NVL(SUM(MAERDE_IMPORT),0) "
   g_str_Parame = g_str_Parame & "                        FROM TPR_MAERDE B "
   g_str_Parame = g_str_Parame & "                       WHERE B.MAERDE_NUMREF = A.MAECFI_NUMREF AND (MAERDE_CODIGO = 2 OR MAERDE_CODIGO = 10 OR MAERDE_CODIGO = 3 OR MAERDE_CODIGO = 11 OR MAERDE_CODIGO = 12 OR MAERDE_CODIGO = 14 OR MAERDE_CODIGO = 19 OR MAERDE_CODIGO = 20)  "
   g_str_Parame = g_str_Parame & "                         AND B.MAERDE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                       GROUP BY MAERDE_NUMREF),0)                                     RECIBIDO_FONDOS,"
   
   'SI HA SIDO RENOVADA Y ESTÁ VIGENTE
   g_str_Parame = g_str_Parame & "                 CASE WHEN A.MAECFI_NUMREN > 0 THEN "
   
   g_str_Parame = g_str_Parame & "                              NVL((SELECT NVL(SUM(MAERDE_IMPORT),0) "
   g_str_Parame = g_str_Parame & "                                     FROM TPR_MAERDE B "
   g_str_Parame = g_str_Parame & "                                    WHERE B.MAERDE_NUMREF = A.MAECFI_NUMREF AND (MAERDE_CODIGO = 2 OR MAERDE_CODIGO = 10 OR MAERDE_CODIGO = 3 OR MAERDE_CODIGO = 11 OR MAERDE_CODIGO = 12 OR MAERDE_CODIGO = 14 OR MAERDE_CODIGO = 17 OR MAERDE_CODIGO = 19 OR MAERDE_CODIGO = 20)  "
   g_str_Parame = g_str_Parame & "                                      AND B.MAERDE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                                    GROUP BY MAERDE_NUMREF),0) "
   g_str_Parame = g_str_Parame & "                 ELSE                   "
   g_str_Parame = g_str_Parame & "                              NVL((SELECT NVL(SUM(MAERDE_IMPORT),0) "
   g_str_Parame = g_str_Parame & "                                     FROM TPR_MAERDE B "
   g_str_Parame = g_str_Parame & "                                    WHERE B.MAERDE_NUMREF = A.MAECFI_NUMREF AND (MAERDE_CODIGO = 2 OR MAERDE_CODIGO = 10 OR MAERDE_CODIGO = 3 OR MAERDE_CODIGO = 11 OR MAERDE_CODIGO = 12 OR MAERDE_CODIGO = 14 OR MAERDE_CODIGO = 19 OR MAERDE_CODIGO = 20 )  "
   g_str_Parame = g_str_Parame & "                                      AND B.MAERDE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                                    GROUP BY MAERDE_NUMREF),0)                                     "
   g_str_Parame = g_str_Parame & "                 END                                                                  IMPORTE_DESEMBOLSADO, "
   
   g_str_Parame = g_str_Parame & "                 NVL((SELECT NVL(SUM(MAERDE_IMPORT),0)"
   g_str_Parame = g_str_Parame & "                        FROM TPR_MAERDE B"
   g_str_Parame = g_str_Parame & "                       WHERE B.MAERDE_NUMREF = A.MAECFI_NUMREF AND (MAERDE_CODIGO = 1 Or MAERDE_CODIGO = 2 Or MAERDE_CODIGO = 4 Or MAERDE_CODIGO = 5 OR MAERDE_CODIGO = 10 OR MAERDE_CODIGO = 6 OR MAERDE_CODIGO = 14 OR MAERDE_CODIGO = 18 OR MAERDE_CODIGO = 21 )"
   g_str_Parame = g_str_Parame & "                         AND B.MAERDE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                       GROUP BY MAERDE_NUMREF),0)                                     PAGADO_DESEMBOLSO,"
   
   g_str_Parame = g_str_Parame & "                 NVL((SELECT NVL(SUM(NVL(MAEGAR_MTOGAR_INM,0) + NVL(MAEGAR_MTOGAR_ES1,0) + NVL(MAEGAR_MTOGAR_ES2,0) + NVL(MAEGAR_MTOGAR_DE1,0) + NVL(MAEGAR_MTOGAR_DE2,0)),0) "
   g_str_Parame = g_str_Parame & "                        FROM TPR_MAEGAR "
   g_str_Parame = g_str_Parame & "                       WHERE MAEGAR_NUMREF = A.MAECFI_NUMREF "
   g_str_Parame = g_str_Parame & "                       GROUP BY MAEGAR_NUMREF),0)                                     IMPORTE_GARANTIA, "
   
   g_str_Parame = g_str_Parame & "                 NVL((SELECT NVL(SUM(MAERDE_IMPORT),0) "
   g_str_Parame = g_str_Parame & "                        FROM TPR_MAERDE B "
   g_str_Parame = g_str_Parame & "                       WHERE B.MAERDE_NUMREF = A.MAECFI_NUMREF AND MAERDE_CODIGO = 6 "
   g_str_Parame = g_str_Parame & "                         AND B.MAERDE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                       GROUP BY MAERDE_NUMREF),0)                                     PAGADO_GARANTIA, "
   
   g_str_Parame = g_str_Parame & "                 NVL((SELECT NVL(SUM(MAERDE_IMPORT),0) "
   g_str_Parame = g_str_Parame & "                        FROM TPR_MAERDE B "
   g_str_Parame = g_str_Parame & "                       WHERE B.MAERDE_NUMREF = A.MAECFI_NUMREF AND MAERDE_CODIGO = 6  "
   g_str_Parame = g_str_Parame & "                         AND B.MAERDE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                       GROUP BY MAERDE_NUMREF),0)                                     RETENCION_GARANTIA, "
   
   g_str_Parame = g_str_Parame & "                 NVL((SELECT NVL(SUM(MAERDE_IMPORT),0) "
   g_str_Parame = g_str_Parame & "                        FROM TPR_MAERDE B "
   g_str_Parame = g_str_Parame & "                       WHERE B.MAERDE_NUMREF = A.MAECFI_NUMREF AND MAERDE_CODIGO = 7  "
   g_str_Parame = g_str_Parame & "                         AND B.MAERDE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                       GROUP BY MAERDE_NUMREF),0)                                     DEVOLUCION_GARANTIA, "
   
   g_str_Parame = g_str_Parame & "                 NVL((SELECT NVL(SUM(MAERDE_IMPORT),0) "
   g_str_Parame = g_str_Parame & "                        FROM TPR_MAERDE B "
   g_str_Parame = g_str_Parame & "                       WHERE B.MAERDE_NUMREF = A.MAECFI_NUMREF AND MAERDE_CODIGO = 22 "
   g_str_Parame = g_str_Parame & "                         AND B.MAERDE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                       GROUP BY MAERDE_NUMREF),0)                                     DEVOLUCION_FONDO_CLIENTE, "
   
   g_str_Parame = g_str_Parame & "                 TRIM(PARDES_DESCRI)                                                  SITUACION, "
   g_str_Parame = g_str_Parame & "                 MAECFI_CODMOD                                                        MODALIDAD "
   
   g_str_Parame = g_str_Parame & "            FROM TPR_MAECFI A "
   g_str_Parame = g_str_Parame & "                 INNER JOIN CNTBL_MAEPRV ON MAEPRV_TIPDOC = A.MAECFI_TIPDOC AND MAEPRV_NUMDOC = MAECFI_NUMDOC "
   g_str_Parame = g_str_Parame & "                 INNER JOIN MNT_PARDES   ON PARDES_CODGRP = '529' AND PARDES_CODITE = A.MAECFI_SITUAC "
   g_str_Parame = g_str_Parame & "           WHERE A.MAECFI_TIPDOC = " & CStr(moddat_g_int_TipDoc) & "  AND A.MAECFI_NUMDOC = '" & CStr(moddat_g_str_NumDoc) & "' "
   g_str_Parame = g_str_Parame & "             AND A.MAECFI_CODPRD IN ('026', '027') "
   
   If Me.cmb_EstFia.ListIndex <> -1 Then
        g_str_Parame = g_str_Parame & "        AND A.MAECFI_SITUAC = " & cmb_EstFia.ItemData(cmb_EstFia.ListIndex) & " "
   End If
    
   g_str_Parame = g_str_Parame & "        ) "
   
   g_str_Parame = g_str_Parame & "     WHERE FECHA_EMISION >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " "
   g_str_Parame = g_str_Parame & "       AND FECHA_EMISION <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & " "
   g_str_Parame = g_str_Parame & "     ORDER BY REFERENCIA"
            
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   grd_Listad.Redraw = False
   Call gs_LimpiaGrid(grd_Listad)
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      grd_Listad.Redraw = True
      
      If grd_Listad.Rows = 0 And grd_Listad_Dir.Rows = 0 Then
         Call fs_Activa(True)
      End If
      
      cmd_Buscar.Enabled = True
      cmd_Limpia.Enabled = True
      cmd_Agrega.Enabled = True
      Exit Sub
   End If
   
'   Call fs_Obtiene_Cabecera
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      cmd_Histor.Enabled = True
      cmd_Gestion.Enabled = True
      cmd_DatBen.Enabled = True
      cmd_ExpExc.Enabled = True
      cmd_ExpLiq.Enabled = True
      
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         grd_Listad.BackColor = &H80000005
         grd_Listad.ForeColor = &H80000008
         
         grd_Listad.Col = 0
         If Not IsNull(g_rst_Princi!REFERENCIA_ANTERIOR) Then
            grd_Listad.Text = fs_Obtener_Tipo(CStr(Trim(g_rst_Princi!REFERENCIA_ANTERIOR)), g_rst_Princi!PRODUCTO)
         Else
            grd_Listad.Text = fs_Obtener_Tipo(CStr(Trim(g_rst_Princi!REFERENCIA)), g_rst_Princi!PRODUCTO)
         End If
         
         grd_Listad.Col = 1
         If Not IsNull(g_rst_Princi!REF_FMV) Then
            grd_Listad.Text = Trim(g_rst_Princi!REF_FMV)
         Else
            grd_Listad.Text = ""
         End If
         
         grd_Listad.Col = 2
         If Not IsNull(g_rst_Princi!REFERENCIA_ANTERIOR) Then
            grd_Listad.Text = gf_Formato_NumRef(CStr(Trim(g_rst_Princi!REFERENCIA_ANTERIOR)), Mid(g_rst_Princi!REFERENCIA_ANTERIOR, 1, 1))
         Else
            grd_Listad.Text = gf_Formato_NumRef(CStr(Trim(g_rst_Princi!REFERENCIA)), Mid(g_rst_Princi!REFERENCIA, 1, 1))
         End If
         
         grd_Listad.Col = 3
         grd_Listad.Text = Format(gf_FormatoFecha(CStr(g_rst_Princi!FECHA_EMISION)), "dd/mm/yyyy")
         
         grd_Listad.Col = 4
         grd_Listad.Text = Format(gf_FormatoFecha(CStr(g_rst_Princi!FECHA_VENCIMIENTO)), "dd/mm/yyyy")
         
         grd_Listad.Col = 5
         grd_Listad.Text = Format(CStr(g_rst_Princi!VALOR_CARTA_FIANZA), "###,###,###,##0.00")
         
         grd_Listad.Col = 6
         grd_Listad.Text = Format(CStr(g_rst_Princi!GARANTIZADO), "###,###,###,##0.00")
      
         grd_Listad.Col = 7
         grd_Listad.Text = Format(CStr(g_rst_Princi!IMPORTE_COMISION), "###,###,###,##0.00")
         
         grd_Listad.Col = 8
         grd_Listad.Text = Format(CStr(g_rst_Princi!PAGADO_COMISION), "###,###,###,##0.00")
         
         grd_Listad.Col = 9
         grd_Listad.Text = Format(CStr(g_rst_Princi!SALDO_COMISION), "###,###,###,##0.00")
         
         grd_Listad.Col = 10
         grd_Listad.Text = Format(CStr(g_rst_Princi!IMPORTE_FONDOS), "###,###,###,##0.00")
         
         grd_Listad.Col = 11
         grd_Listad.Text = Format(CStr(g_rst_Princi!RECIBIDO_FONDOS), "###,###,###,##0.00")
         
         grd_Listad.Col = 12
         grd_Listad.Text = IIf(g_rst_Princi!SALDO_FONDOS < 0, 0, Format(CStr(g_rst_Princi!SALDO_FONDOS), "###,###,###,##0.00"))
         
         grd_Listad.Col = 13
         grd_Listad.Text = Format(CStr(g_rst_Princi!IMPORTE_DESEMBOLSADO), "###,###,###,##0.00")
         
         grd_Listad.Col = 14
         grd_Listad.Text = Format(CStr(g_rst_Princi!PAGADO_DESEMBOLSO - g_rst_Princi!DEVOLUCION_GARANTIA - g_rst_Princi!DEVOLUCION_FONDO_CLIENTE), "###,###,###,##0.00")
         
         grd_Listad.Col = 15
         grd_Listad.Text = Format(CStr(g_rst_Princi!SALDO_DESEMBOLSO + g_rst_Princi!DEVOLUCION_GARANTIA + g_rst_Princi!DEVOLUCION_FONDO_CLIENTE), "###,###,###,##0.00")
         
         grd_Listad.Col = 16
         grd_Listad.Text = CStr(g_rst_Princi!SITUACION)
         
         grd_Listad.Col = 17
         grd_Listad.Text = CStr(g_rst_Princi!REFERENCIA)
         
         grd_Listad.Col = 18
         grd_Listad.Text = CStr(g_rst_Princi!PRODUCTO)
         
         grd_Listad.Col = 19
         If IsNull(g_rst_Princi!SUB_PRODUCTO) Then
            grd_Listad.Text = ""
         Else
            grd_Listad.Text = CStr(g_rst_Princi!SUB_PRODUCTO)
         End If
         
         grd_Listad.Col = 20
         grd_Listad.Text = CStr(g_rst_Princi!MODALIDAD)
         
         grd_Listad.Col = 21
         If IsNull(g_rst_Princi!TIPO_RECURSO) Then
            grd_Listad.Text = ""
         Else
            grd_Listad.Text = CStr(g_rst_Princi!TIPO_RECURSO)
         End If
         
         g_rst_Princi.MoveNext
      Loop
   End If
   grd_Listad.Redraw = True
   Call gs_UbiIniGrid(grd_Listad)
End Sub

Public Sub fs_Buscar_old()
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "   SELECT REFERENCIA           , FECHA_EMISION     , FECHA_VENCIMIENTO    , VALOR_CARTA_FIANZA                   , GARANTIZADO         , "
   g_str_Parame = g_str_Parame & "          IMPORTE_COMISION     , PAGADO_COMISION   , (IMPORTE_COMISION - PAGADO_COMISION) SALDO_COMISION         , EXTORNO_COMISION    , "
   g_str_Parame = g_str_Parame & "          IMPORTE_FONDOS       , RECIBIDO_FONDOS   , (IMPORTE_FONDOS - RECIBIDO_FONDOS) SALDO_FONDOS             ,                       "
   g_str_Parame = g_str_Parame & "          IMPORTE_DESEMBOLSADO , PAGADO_DESEMBOLSO , (IMPORTE_DESEMBOLSADO - PAGADO_DESEMBOLSO) SALDO_DESEMBOLSO , PRODUCTO            , "
   g_str_Parame = g_str_Parame & "          IMPORTE_GARANTIA     , PAGADO_GARANTIA   , (IMPORTE_GARANTIA - PAGADO_GARANTIA) SALDO_GARANTIA         , RETENCION_GARANTIA  , DEVOLUCION_GARANTIA,"
   g_str_Parame = g_str_Parame & "          SITUACION            , MODALIDAD         , REFERENCIA_ANTERIOR                                         , REF_FMV             , SUB_PRODUCTO , "
   g_str_Parame = g_str_Parame & "          TIPO_RECURSO "
   g_str_Parame = g_str_Parame & "    FROM( "
   g_str_Parame = g_str_Parame & "          SELECT A.MAECFI_NUMREF                                                      REFERENCIA,          "
   g_str_Parame = g_str_Parame & "                 A.MAECFI_CODPRD                                                      PRODUCTO,            "
   g_str_Parame = g_str_Parame & "                 A.MAECFI_CODSUB                                                      SUB_PRODUCTO,        "
   g_str_Parame = g_str_Parame & "                 A.MAECFI_NUMADE                                                      REF_FMV,             "
   g_str_Parame = g_str_Parame & "                 A.MAECFI_EMIFIA                                                      FECHA_EMISION,       "
   g_str_Parame = g_str_Parame & "                 A.MAECFI_VTOFIA                                                      FECHA_VENCIMIENTO,   "
   g_str_Parame = g_str_Parame & "                 A.MAECFI_IMPFIA                                                      VALOR_CARTA_FIANZA,  "
   g_str_Parame = g_str_Parame & "                 A.MAECFI_GARFIA                                                      GARANTIZADO,         "
   g_str_Parame = g_str_Parame & "                 A.MAECFI_NUMANT                                                      REFERENCIA_ANTERIOR, "
   g_str_Parame = g_str_Parame & "                 A.MAECFI_COMFIA                                                      IMPORTE_COMISION,    "
   g_str_Parame = g_str_Parame & "                 A.MAECFI_TIPREC                                                      TIPO_RECURSO,        "
   
   g_str_Parame = g_str_Parame & "                 NVL((SELECT NVL(SUM(MAERDE_IMPORT),0) "
   g_str_Parame = g_str_Parame & "                        FROM TPR_MAERDE B "
   g_str_Parame = g_str_Parame & "                       WHERE B.MAERDE_NUMREF = A.MAECFI_NUMREF AND (MAERDE_CODIGO = 1 OR MAERDE_CODIGO = 2 OR MAERDE_CODIGO = 13) "
   g_str_Parame = g_str_Parame & "                         AND B.MAERDE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                       GROUP BY MAERDE_NUMREF),0)                                     PAGADO_COMISION, "
   
   g_str_Parame = g_str_Parame & "                 NVL((SELECT NVL(SUM(MAERDE_IMPORT),0) "
   g_str_Parame = g_str_Parame & "                        FROM TPR_MAERDE B "
   g_str_Parame = g_str_Parame & "                       WHERE B.MAERDE_NUMREF = A.MAECFI_NUMREF AND MAERDE_CODIGO = 13 "
   g_str_Parame = g_str_Parame & "                         AND B.MAERDE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                       GROUP BY MAERDE_NUMREF),0)                                     COMISION_DEPOSITO, "
   
   g_str_Parame = g_str_Parame & "                 NVL((SELECT NVL(SUM(MAERDE_IMPORT),0) "
   g_str_Parame = g_str_Parame & "                        FROM TPR_MAERDE B "
   g_str_Parame = g_str_Parame & "                       WHERE B.MAERDE_NUMREF = A.MAECFI_NUMREF AND MAERDE_CODIGO = 15 "
   g_str_Parame = g_str_Parame & "                         AND B.MAERDE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                       GROUP BY MAERDE_NUMREF),0)                                     EXTORNO_COMISION,"
   
   'SI HA SIDO RENOVADA Y ESTÁ VIGENTE
   g_str_Parame = g_str_Parame & "                 CASE WHEN A.MAECFI_NUMREN > 0 THEN "
   
   g_str_Parame = g_str_Parame & "                              NVL(( SELECT NVL(SUM(MAERDE_IMPORT),0) "
   g_str_Parame = g_str_Parame & "                                      FROM TPR_MAERDE B "
   g_str_Parame = g_str_Parame & "                                     WHERE B.MAERDE_NUMREF = A.MAECFI_NUMREF AND MAERDE_CODIGO = 16   "
   g_str_Parame = g_str_Parame & "                                       AND B.MAERDE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                                     GROUP BY MAERDE_NUMREF),0)                                       "
   
   g_str_Parame = g_str_Parame & "                 ELSE "
   g_str_Parame = g_str_Parame & "                              A.MAECFI_IMPFIA                                                         "
   g_str_Parame = g_str_Parame & "                 END                                                                  IMPORTE_FONDOS, "
   
   g_str_Parame = g_str_Parame & "                 NVL((SELECT NVL(SUM(MAERDE_IMPORT),0) "
   g_str_Parame = g_str_Parame & "                        FROM TPR_MAERDE B "
   g_str_Parame = g_str_Parame & "                       WHERE B.MAERDE_NUMREF = A.MAECFI_NUMREF AND (MAERDE_CODIGO = 2 OR MAERDE_CODIGO = 10 OR MAERDE_CODIGO = 3 OR MAERDE_CODIGO = 11 OR MAERDE_CODIGO = 12 OR MAERDE_CODIGO = 14 OR MAERDE_CODIGO = 19)  "
   g_str_Parame = g_str_Parame & "                         AND B.MAERDE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                       GROUP BY MAERDE_NUMREF),0)                                     RECIBIDO_FONDOS,"
   
   'SI HA SIDO RENOVADA Y ESTÁ VIGENTE
   g_str_Parame = g_str_Parame & "                 CASE WHEN A.MAECFI_NUMREN > 0 THEN "
   
   g_str_Parame = g_str_Parame & "                              NVL((SELECT NVL(SUM(MAERDE_IMPORT),0) "
   g_str_Parame = g_str_Parame & "                                     FROM TPR_MAERDE B "
   g_str_Parame = g_str_Parame & "                                    WHERE B.MAERDE_NUMREF = A.MAECFI_NUMREF AND (MAERDE_CODIGO = 2 OR MAERDE_CODIGO = 10 OR MAERDE_CODIGO = 3 OR MAERDE_CODIGO = 11 OR MAERDE_CODIGO = 12 OR MAERDE_CODIGO = 14 OR MAERDE_CODIGO = 17 OR MAERDE_CODIGO = 19)  "
   g_str_Parame = g_str_Parame & "                                      AND B.MAERDE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                                    GROUP BY MAERDE_NUMREF),0) "
   g_str_Parame = g_str_Parame & "                 ELSE                   "
   g_str_Parame = g_str_Parame & "                              NVL((SELECT NVL(SUM(MAERDE_IMPORT),0) "
   g_str_Parame = g_str_Parame & "                                     FROM TPR_MAERDE B "
   g_str_Parame = g_str_Parame & "                                    WHERE B.MAERDE_NUMREF = A.MAECFI_NUMREF AND (MAERDE_CODIGO = 2 OR MAERDE_CODIGO = 10 OR MAERDE_CODIGO = 3 OR MAERDE_CODIGO = 11 OR MAERDE_CODIGO = 12 OR MAERDE_CODIGO = 14 OR MAERDE_CODIGO = 19 )  "
   g_str_Parame = g_str_Parame & "                                      AND B.MAERDE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                                    GROUP BY MAERDE_NUMREF),0)                                     "
   g_str_Parame = g_str_Parame & "                 END                                                                  IMPORTE_DESEMBOLSADO, "
   
   g_str_Parame = g_str_Parame & "                 NVL((SELECT NVL(SUM(MAERDE_IMPORT),0)"
   g_str_Parame = g_str_Parame & "                        FROM TPR_MAERDE B"
   g_str_Parame = g_str_Parame & "                       WHERE B.MAERDE_NUMREF = A.MAECFI_NUMREF AND (MAERDE_CODIGO = 1 Or MAERDE_CODIGO = 2 Or MAERDE_CODIGO = 4 Or MAERDE_CODIGO = 5 OR MAERDE_CODIGO = 10 OR MAERDE_CODIGO = 6 OR MAERDE_CODIGO = 14 OR MAERDE_CODIGO = 18)"
   g_str_Parame = g_str_Parame & "                         AND B.MAERDE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                       GROUP BY MAERDE_NUMREF),0)                                     PAGADO_DESEMBOLSO,"
   
   g_str_Parame = g_str_Parame & "                 NVL((SELECT NVL(SUM(NVL(MAEGAR_MTOGAR_INM,0) + NVL(MAEGAR_MTOGAR_ES1,0) + NVL(MAEGAR_MTOGAR_ES2,0) + NVL(MAEGAR_MTOGAR_DE1,0) + NVL(MAEGAR_MTOGAR_DE2,0)),0) "
   g_str_Parame = g_str_Parame & "                        FROM TPR_MAEGAR "
   g_str_Parame = g_str_Parame & "                       WHERE MAEGAR_NUMREF = A.MAECFI_NUMREF "
   g_str_Parame = g_str_Parame & "                       GROUP BY MAEGAR_NUMREF),0)                                     IMPORTE_GARANTIA, "
   
   g_str_Parame = g_str_Parame & "                 NVL((SELECT NVL(SUM(MAERDE_IMPORT),0) "
   g_str_Parame = g_str_Parame & "                        FROM TPR_MAERDE B "
   g_str_Parame = g_str_Parame & "                       WHERE B.MAERDE_NUMREF = A.MAECFI_NUMREF AND MAERDE_CODIGO = 6 "
   g_str_Parame = g_str_Parame & "                         AND B.MAERDE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                       GROUP BY MAERDE_NUMREF),0)                                     PAGADO_GARANTIA, "
   
   g_str_Parame = g_str_Parame & "                 NVL((SELECT NVL(SUM(MAERDE_IMPORT),0) "
   g_str_Parame = g_str_Parame & "                        FROM TPR_MAERDE B "
   g_str_Parame = g_str_Parame & "                       WHERE B.MAERDE_NUMREF = A.MAECFI_NUMREF AND MAERDE_CODIGO = 6 "
   g_str_Parame = g_str_Parame & "                         AND B.MAERDE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                       GROUP BY MAERDE_NUMREF),0)                                     RETENCION_GARANTIA, "
   
   g_str_Parame = g_str_Parame & "                 NVL((SELECT NVL(SUM(MAERDE_IMPORT),0) "
   g_str_Parame = g_str_Parame & "                        FROM TPR_MAERDE B "
   g_str_Parame = g_str_Parame & "                       WHERE B.MAERDE_NUMREF = A.MAECFI_NUMREF AND MAERDE_CODIGO = 7 "
   g_str_Parame = g_str_Parame & "                         AND B.MAERDE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                       GROUP BY MAERDE_NUMREF),0)                                     DEVOLUCION_GARANTIA, "
   
   g_str_Parame = g_str_Parame & "                 TRIM(PARDES_DESCRI)                                                  SITUACION, "
   g_str_Parame = g_str_Parame & "                 MAECFI_CODMOD                                                        MODALIDAD "
   
   g_str_Parame = g_str_Parame & "            FROM TPR_MAECFI A "
   g_str_Parame = g_str_Parame & "                 INNER JOIN CNTBL_MAEPRV ON MAEPRV_TIPDOC = A.MAECFI_TIPDOC AND MAEPRV_NUMDOC = MAECFI_NUMDOC "
   g_str_Parame = g_str_Parame & "                 INNER JOIN MNT_PARDES   ON PARDES_CODGRP = '529' AND PARDES_CODITE = A.MAECFI_SITUAC "
   g_str_Parame = g_str_Parame & "           WHERE A.MAECFI_TIPDOC = " & CStr(moddat_g_int_TipDoc) & "  AND A.MAECFI_NUMDOC = '" & CStr(moddat_g_str_NumDoc) & "' "
   
   If Me.cmb_EstFia.ListIndex <> -1 Then
        g_str_Parame = g_str_Parame & "        AND A.MAECFI_SITUAC = " & cmb_EstFia.ItemData(cmb_EstFia.ListIndex) & " "
   End If
    
   g_str_Parame = g_str_Parame & "        ) "
   
   g_str_Parame = g_str_Parame & "     WHERE FECHA_EMISION >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " "
   g_str_Parame = g_str_Parame & "       AND FECHA_EMISION <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & " "
   g_str_Parame = g_str_Parame & "     ORDER BY REFERENCIA"
            
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   grd_Listad.Redraw = False
   Call gs_LimpiaGrid(grd_Listad)
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      grd_Listad.Redraw = True
      Call fs_Activa(True)
      cmd_Buscar.Enabled = True
      cmd_Limpia.Enabled = True
      cmd_Agrega.Enabled = True
      Exit Sub
   End If
   
'   Call fs_Obtiene_Cabecera
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      cmd_Histor.Enabled = True
      cmd_Gestion.Enabled = True
      cmd_DatBen.Enabled = True
      cmd_ExpExc.Enabled = True
      cmd_ExpLiq.Enabled = True
      
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         grd_Listad.BackColor = &H80000005
         grd_Listad.ForeColor = &H80000008
         
         grd_Listad.Col = 0
         If Not IsNull(g_rst_Princi!REFERENCIA_ANTERIOR) Then
            grd_Listad.Text = fs_Obtener_Tipo(CStr(Trim(g_rst_Princi!REFERENCIA_ANTERIOR)), g_rst_Princi!PRODUCTO)
         Else
            grd_Listad.Text = fs_Obtener_Tipo(CStr(Trim(g_rst_Princi!REFERENCIA)), g_rst_Princi!PRODUCTO)
         End If
         
         grd_Listad.Col = 1
         If Not IsNull(g_rst_Princi!REF_FMV) Then
            grd_Listad.Text = Trim(g_rst_Princi!REF_FMV)
         Else
            grd_Listad.Text = ""
         End If
         
         grd_Listad.Col = 2
         If Not IsNull(g_rst_Princi!REFERENCIA_ANTERIOR) Then
            grd_Listad.Text = gf_Formato_NumRef(CStr(Trim(g_rst_Princi!REFERENCIA_ANTERIOR)), Mid(g_rst_Princi!REFERENCIA_ANTERIOR, 1, 1))
         Else
            grd_Listad.Text = gf_Formato_NumRef(CStr(Trim(g_rst_Princi!REFERENCIA)), Mid(g_rst_Princi!REFERENCIA, 1, 1))
         End If
         
         grd_Listad.Col = 3
         grd_Listad.Text = Format(gf_FormatoFecha(CStr(g_rst_Princi!FECHA_EMISION)), "dd/mm/yyyy")
         
         grd_Listad.Col = 4
         grd_Listad.Text = Format(gf_FormatoFecha(CStr(g_rst_Princi!FECHA_VENCIMIENTO)), "dd/mm/yyyy")
         
         grd_Listad.Col = 5
         grd_Listad.Text = Format(CStr(g_rst_Princi!VALOR_CARTA_FIANZA), "###,###,###,##0.00")
         
         grd_Listad.Col = 6
         grd_Listad.Text = Format(CStr(g_rst_Princi!GARANTIZADO), "###,###,###,##0.00")
      
         grd_Listad.Col = 7
         grd_Listad.Text = Format(CStr(g_rst_Princi!IMPORTE_COMISION), "###,###,###,##0.00")
         
         grd_Listad.Col = 8
         grd_Listad.Text = Format(CStr(g_rst_Princi!PAGADO_COMISION), "###,###,###,##0.00")
         
         grd_Listad.Col = 9
         grd_Listad.Text = Format(CStr(g_rst_Princi!SALDO_COMISION), "###,###,###,##0.00")
         
         grd_Listad.Col = 10
         grd_Listad.Text = Format(CStr(g_rst_Princi!IMPORTE_FONDOS), "###,###,###,##0.00")
         
         grd_Listad.Col = 11
         grd_Listad.Text = Format(CStr(g_rst_Princi!RECIBIDO_FONDOS), "###,###,###,##0.00")
         
         grd_Listad.Col = 12
         grd_Listad.Text = IIf(g_rst_Princi!SALDO_FONDOS < 0, 0, Format(CStr(g_rst_Princi!SALDO_FONDOS), "###,###,###,##0.00"))
         
         grd_Listad.Col = 13
         grd_Listad.Text = Format(CStr(g_rst_Princi!IMPORTE_DESEMBOLSADO), "###,###,###,##0.00")
         
         grd_Listad.Col = 14
         grd_Listad.Text = Format(CStr(g_rst_Princi!PAGADO_DESEMBOLSO - g_rst_Princi!DEVOLUCION_GARANTIA), "###,###,###,##0.00")
         
         grd_Listad.Col = 15
         grd_Listad.Text = Format(CStr(g_rst_Princi!SALDO_DESEMBOLSO + g_rst_Princi!DEVOLUCION_GARANTIA), "###,###,###,##0.00")
         
         grd_Listad.Col = 16
         grd_Listad.Text = CStr(g_rst_Princi!SITUACION)
         
         grd_Listad.Col = 17
         grd_Listad.Text = CStr(g_rst_Princi!REFERENCIA)
         
         grd_Listad.Col = 18
         grd_Listad.Text = CStr(g_rst_Princi!PRODUCTO)
         
         grd_Listad.Col = 19
         If IsNull(g_rst_Princi!SUB_PRODUCTO) Then
            grd_Listad.Text = ""
         Else
            grd_Listad.Text = CStr(g_rst_Princi!SUB_PRODUCTO)
         End If
         
         grd_Listad.Col = 20
         grd_Listad.Text = CStr(g_rst_Princi!MODALIDAD)
         
         grd_Listad.Col = 21
         If IsNull(g_rst_Princi!TIPO_RECURSO) Then
            grd_Listad.Text = ""
         Else
            grd_Listad.Text = CStr(g_rst_Princi!TIPO_RECURSO)
         End If
         
         g_rst_Princi.MoveNext
      Loop
   End If
'   grd_Listad.FixedRows = 2
   grd_Listad.Redraw = True
   Call gs_UbiIniGrid(grd_Listad)
End Sub

Public Sub fs_Buscar_ant()
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT MAECFI_NUMREF, MAECFI_EMIFIA, MAECFI_VTOFIA, MAECFI_IMPFIA, MAECFI_GARFIA, "
   g_str_Parame = g_str_Parame & "        IMPORTE_GARANTIA, PAGADO_GARANTIA, (IMPORTE_GARANTIA - PAGADO_GARANTIA) SALDO_GARANTIA, "
   g_str_Parame = g_str_Parame & "        IMPORTE_COMISION, PAGADO_COMISION,(IMPORTE_COMISION - PAGADO_COMISION) SALDO_COMISION, "
   g_str_Parame = g_str_Parame & "        MAECFI_IMPFIA IMPORTE_FONDOS, RECIBIDO_FONDOS," 'CASE WHEN COMISION_DEPOSITO >0 THEN RECIBIDO_FONDOS - COMISION_DEPOSITO ELSE RECIBIDO_FONDOS END AS RECIBIDO_FONDOS, "
   g_str_Parame = g_str_Parame & "        (MAECFI_IMPFIA - RECIBIDO_FONDOS) SALDO_FONDOS, FONDOS AS IMPORTE_DESEMBOLSADO, "
   g_str_Parame = g_str_Parame & "        PAGADO_DESEMBOLSO , " 'CASE WHEN COMISION_DEPOSITO > 0 THEN PAGADO_DESEMBOLSO - COMISION_DEPOSITO ELSE PAGADO_DESEMBOLSO END AS PAGADO_DESEMBOLSO , "
   g_str_Parame = g_str_Parame & "        DEVOLUCION_GARANTIA , MAECFI_NUMANT, "
   g_str_Parame = g_str_Parame & "        (FONDOS - PAGADO_DESEMBOLSO) SALDO_DESEMBOLSO, SITUACION, MODALIDAD, RETENCION_GARANTIA, EXTORNO_COMISION "
   g_str_Parame = g_str_Parame & "  FROM( "
   g_str_Parame = g_str_Parame & "        SELECT A.MAECFI_NUMREF, A.MAECFI_EMIFIA, A.MAECFI_VTOFIA, A.MAECFI_IMPFIA, A.MAECFI_GARFIA, MAECFI_NUMANT, "
   g_str_Parame = g_str_Parame & "               NVL((SELECT NVL(SUM(NVL(MAEGAR_MTOGAR_INM,0) + NVL(MAEGAR_MTOGAR_ES1,0) + NVL(MAEGAR_MTOGAR_ES2,0) + NVL(MAEGAR_MTOGAR_DE1,0) + NVL(MAEGAR_MTOGAR_DE2,0)),0) "
   g_str_Parame = g_str_Parame & "                      FROM TPR_MAEGAR "
   g_str_Parame = g_str_Parame & "                     WHERE MAEGAR_NUMREF = A.MAECFI_NUMREF "
   g_str_Parame = g_str_Parame & "                     GROUP BY MAEGAR_NUMREF),0) IMPORTE_GARANTIA, "
   g_str_Parame = g_str_Parame & "               NVL((SELECT NVL(SUM(MAERDE_IMPORT),0) "
   g_str_Parame = g_str_Parame & "                      FROM TPR_MAERDE B "
   g_str_Parame = g_str_Parame & "                     WHERE B.MAERDE_NUMREF = A.MAECFI_NUMREF AND MAERDE_CODIGO = 6 "
   g_str_Parame = g_str_Parame & "                       AND B.MAERDE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                     GROUP BY MAERDE_NUMREF),0) PAGADO_GARANTIA, "
   
'   g_str_Parame = g_str_Parame & "               CASE WHEN (NVL((SELECT NVL(SUM(MAECFI_COMFIA),0) "
'   g_str_Parame = g_str_Parame & "                                FROM TPR_MAECFI B"
'   g_str_Parame = g_str_Parame & "                               Where B.MAECFI_REFORI = A.MAECFI_REFORI"
'   g_str_Parame = g_str_Parame & "                               GROUP BY MAECFI_REFORI),0)) = 0 THEN MAECFI_COMFIA"
'   g_str_Parame = g_str_Parame & "               ELSE "
'   g_str_Parame = g_str_Parame & "                   NVL((SELECT NVL(SUM(MAECFI_COMFIA),0)"
'   g_str_Parame = g_str_Parame & "                          FROM TPR_MAECFI B"
'   g_str_Parame = g_str_Parame & "                         Where B.MAECFI_REFORI = A.MAECFI_REFORI"
'   g_str_Parame = g_str_Parame & "                         GROUP BY MAECFI_REFORI),0)"
'   g_str_Parame = g_str_Parame & "               END AS IMPORTE_COMISION,"
   
   g_str_Parame = g_str_Parame & "               MAECFI_COMFIA AS IMPORTE_COMISION, "
   
   g_str_Parame = g_str_Parame & "               NVL((SELECT NVL(SUM(MAERDE_IMPORT),0) "
   g_str_Parame = g_str_Parame & "                      FROM TPR_MAERDE B "
   g_str_Parame = g_str_Parame & "                     WHERE B.MAERDE_NUMREF = A.MAECFI_NUMREF AND (MAERDE_CODIGO = 1 OR MAERDE_CODIGO = 2 OR MAERDE_CODIGO = 13) "
   g_str_Parame = g_str_Parame & "                       AND B.MAERDE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                     GROUP BY MAERDE_NUMREF),0) PAGADO_COMISION, "
   
   g_str_Parame = g_str_Parame & "               NVL((SELECT NVL(SUM(MAERDE_IMPORT),0) "
   g_str_Parame = g_str_Parame & "                      FROM TPR_MAERDE B "
   g_str_Parame = g_str_Parame & "                     WHERE B.MAERDE_NUMREF = A.MAECFI_NUMREF AND (MAERDE_CODIGO = 2 OR MAERDE_CODIGO = 10 OR MAERDE_CODIGO = 3 OR MAERDE_CODIGO = 11 OR MAERDE_CODIGO = 12 OR MAERDE_CODIGO = 14 ) " 'OR MAERDE_CODIGO = 15
   g_str_Parame = g_str_Parame & "                       AND B.MAERDE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                     GROUP BY MAERDE_NUMREF),0) FONDOS, " 'FONDOS = IMPORTE_DESEMBOLSADO

   g_str_Parame = g_str_Parame & "               NVL((SELECT NVL(SUM(MAERDE_IMPORT),0) "
   g_str_Parame = g_str_Parame & "                      FROM TPR_MAERDE B "
   g_str_Parame = g_str_Parame & "                     WHERE B.MAERDE_NUMREF = A.MAECFI_NUMREF AND (MAERDE_CODIGO = 2 OR MAERDE_CODIGO = 10 OR MAERDE_CODIGO = 3 OR MAERDE_CODIGO = 11 OR MAERDE_CODIGO = 12 OR MAERDE_CODIGO = 14 ) " 'OR MAERDE_CODIGO = 15
   g_str_Parame = g_str_Parame & "                       AND B.MAERDE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                     GROUP BY MAERDE_NUMREF),0) RECIBIDO_FONDOS, "
   
   g_str_Parame = g_str_Parame & "               NVL((SELECT NVL(SUM(MAERDE_IMPORT),0)"
   g_str_Parame = g_str_Parame & "                      FROM TPR_MAERDE B"
   g_str_Parame = g_str_Parame & "                     WHERE B.MAERDE_NUMREF = A.MAECFI_NUMREF AND (MAERDE_CODIGO = 1 Or MAERDE_CODIGO = 2 Or MAERDE_CODIGO = 4 Or MAERDE_CODIGO = 5 OR MAERDE_CODIGO = 10 OR MAERDE_CODIGO = 6)"
   g_str_Parame = g_str_Parame & "                       AND B.MAERDE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                     GROUP BY MAERDE_NUMREF),0) PAGADO_DESEMBOLSO,"
   
   g_str_Parame = g_str_Parame & "               NVL((SELECT NVL(SUM(MAERDE_IMPORT),0) "
   g_str_Parame = g_str_Parame & "                      FROM TPR_MAERDE B "
   g_str_Parame = g_str_Parame & "                     WHERE B.MAERDE_NUMREF = A.MAECFI_NUMREF AND MAERDE_CODIGO = 6 "
   g_str_Parame = g_str_Parame & "                       AND B.MAERDE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                     GROUP BY MAERDE_NUMREF),0) RETENCION_GARANTIA, "
   
   g_str_Parame = g_str_Parame & "               NVL((SELECT NVL(SUM(MAERDE_IMPORT),0) "
   g_str_Parame = g_str_Parame & "                      FROM TPR_MAERDE B "
   g_str_Parame = g_str_Parame & "                     WHERE B.MAERDE_NUMREF = A.MAECFI_NUMREF AND MAERDE_CODIGO = 7 "
   g_str_Parame = g_str_Parame & "                       AND B.MAERDE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                     GROUP BY MAERDE_NUMREF),0) DEVOLUCION_GARANTIA, "
   g_str_Parame = g_str_Parame & "               TRIM(PARDES_DESCRI) SITUACION, MAECFI_CODMOD MODALIDAD, "
   
   g_str_Parame = g_str_Parame & "               NVL((SELECT NVL(SUM(MAERDE_IMPORT),0) "
   g_str_Parame = g_str_Parame & "                      FROM TPR_MAERDE B "
   g_str_Parame = g_str_Parame & "                     WHERE B.MAERDE_NUMREF = A.MAECFI_NUMREF AND MAERDE_CODIGO = 13 "
   g_str_Parame = g_str_Parame & "                       AND B.MAERDE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                     GROUP BY MAERDE_NUMREF),0) COMISION_DEPOSITO, "
   
   g_str_Parame = g_str_Parame & "               NVL((SELECT NVL(SUM(MAERDE_IMPORT),0) "
   g_str_Parame = g_str_Parame & "                      FROM TPR_MAERDE B "
   g_str_Parame = g_str_Parame & "                     WHERE B.MAERDE_NUMREF = A.MAECFI_NUMREF AND MAERDE_CODIGO = 15 "
   g_str_Parame = g_str_Parame & "                       AND B.MAERDE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                     GROUP BY MAERDE_NUMREF),0) EXTORNO_COMISION "
   
   g_str_Parame = g_str_Parame & "          FROM TPR_MAECFI A "
   g_str_Parame = g_str_Parame & "               INNER JOIN CNTBL_MAEPRV ON MAEPRV_TIPDOC = MAECFI_TIPDOC AND MAEPRV_NUMDOC = MAECFI_NUMDOC "
   g_str_Parame = g_str_Parame & "               INNER JOIN MNT_PARDES ON PARDES_CODGRP = '529' AND PARDES_CODITE = MAECFI_SITUAC "
   g_str_Parame = g_str_Parame & "         WHERE MAECFI_TIPDOC = " & CStr(moddat_g_int_TipDoc) & "  AND MAECFI_NUMDOC = '" & CStr(moddat_g_str_NumDoc) & "' "
   
   If Me.cmb_EstFia.ListIndex <> -1 Then
      g_str_Parame = g_str_Parame & "        AND MAECFI_SITUAC = " & cmb_EstFia.ItemData(cmb_EstFia.ListIndex) & " "
   End If
   
   g_str_Parame = g_str_Parame & "      ) "
   
   g_str_Parame = g_str_Parame & "   WHERE MAECFI_EMIFIA >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " "
   g_str_Parame = g_str_Parame & "     AND MAECFI_EMIFIA <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & " "
   
   g_str_Parame = g_str_Parame & "   ORDER BY MAECFI_NUMREF "
            
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   grd_Listad.Redraw = False
   Call gs_LimpiaGrid(grd_Listad)
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      grd_Listad.Redraw = True
      
      Call fs_Activa(True)
      cmd_Buscar.Enabled = True
      cmd_Limpia.Enabled = True
      cmd_Agrega.Enabled = True
      Exit Sub
   End If

   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      cmd_Gestion.Enabled = True
      cmd_DatBen.Enabled = True
      cmd_ExpExc.Enabled = True
      
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         grd_Listad.BackColor = &H80000005
         grd_Listad.ForeColor = &H80000008
         
         grd_Listad.Col = 0
         If Not IsNull(g_rst_Princi!MAECFI_NUMANT) Then
            grd_Listad.Text = fs_Obtener_Tipo(CStr(Trim(g_rst_Princi!MAECFI_NUMANT)))
         Else
            grd_Listad.Text = fs_Obtener_Tipo(CStr(Trim(g_rst_Princi!MAECFI_NUMREF)))
         End If
                 
         grd_Listad.Col = 1
         If Not IsNull(g_rst_Princi!MAECFI_NUMANT) Then
            grd_Listad.Text = gf_Formato_NumRef(CStr(Trim(g_rst_Princi!MAECFI_NUMANT)), Mid(g_rst_Princi!MAECFI_NUMANT, 1, 1))
         Else
            grd_Listad.Text = gf_Formato_NumRef(CStr(Trim(g_rst_Princi!MAECFI_NUMREF)), Mid(g_rst_Princi!MAECFI_NUMREF, 1, 1))
         End If
         
         grd_Listad.Col = 2
         grd_Listad.Text = Format(gf_FormatoFecha(CStr(g_rst_Princi!MAECFI_EMIFIA)), "dd/mm/yyyy")
         
         grd_Listad.Col = 3
         grd_Listad.Text = Format(gf_FormatoFecha(CStr(g_rst_Princi!MAECFI_VTOFIA)), "dd/mm/yyyy")
         
         grd_Listad.Col = 4
         grd_Listad.Text = Format(CStr(g_rst_Princi!MAECFI_IMPFIA), "###,###,###,##0.00")
         
         grd_Listad.Col = 5
         grd_Listad.Text = Format(CStr(g_rst_Princi!MAECFI_GARFIA), "###,###,###,##0.00")
      
         grd_Listad.Col = 6
         grd_Listad.Text = Format(CStr(g_rst_Princi!IMPORTE_COMISION), "###,###,###,##0.00")
         
         grd_Listad.Col = 7
         grd_Listad.Text = Format(CStr(g_rst_Princi!PAGADO_COMISION), "###,###,###,##0.00")
         
         grd_Listad.Col = 8
         grd_Listad.Text = Format(CStr(g_rst_Princi!SALDO_COMISION), "###,###,###,##0.00")
         
         grd_Listad.Col = 9
         grd_Listad.Text = Format(CStr(g_rst_Princi!IMPORTE_FONDOS), "###,###,###,##0.00")
         
         grd_Listad.Col = 10
         grd_Listad.Text = Format(CStr(g_rst_Princi!RECIBIDO_FONDOS), "###,###,###,##0.00")
         
         grd_Listad.Col = 11
         grd_Listad.Text = IIf(g_rst_Princi!SALDO_FONDOS < 0, 0, Format(CStr(g_rst_Princi!SALDO_FONDOS), "###,###,###,##0.00"))
         
         grd_Listad.Col = 12
         grd_Listad.Text = Format(CStr(g_rst_Princi!IMPORTE_DESEMBOLSADO), "###,###,###,##0.00")
         
         grd_Listad.Col = 13
         grd_Listad.Text = Format(CStr(g_rst_Princi!PAGADO_DESEMBOLSO - g_rst_Princi!DEVOLUCION_GARANTIA), "###,###,###,##0.00")
         
         grd_Listad.Col = 14
         grd_Listad.Text = Format(CStr(g_rst_Princi!SALDO_DESEMBOLSO + g_rst_Princi!DEVOLUCION_GARANTIA), "###,###,###,##0.00") '- g_rst_Princi!RETENCION_GARANTIA
         ' + g_rst_Princi!EXTORNO_COMISION
         
         grd_Listad.Col = 15
         grd_Listad.Text = CStr(g_rst_Princi!SITUACION)
         
         grd_Listad.Col = 16
         grd_Listad.Text = CStr(g_rst_Princi!MAECFI_NUMREF)
         
         g_rst_Princi.MoveNext
      Loop
   End If
   
   grd_Listad.Redraw = True
   Call gs_UbiIniGrid(grd_Listad)
End Sub

Private Function fs_Obtener_Tipo(ByVal p_NumRef As String, Optional ByVal p_CodPrd As String, Optional p_CodMod As String) As String
   p_NumRef = Format(p_NumRef, "0000000000")
   p_NumRef = Mid(p_NumRef, 1, 1)
   
   If p_CodPrd = "026" Or p_CodPrd = "027" Then
      If p_NumRef <= 1 Then
         fs_Obtener_Tipo = "CF"
      ElseIf p_NumRef = 2 Then
         fs_Obtener_Tipo = "AD"
      Else
         fs_Obtener_Tipo = "CSO"
      End If
   Else
      If p_CodMod = "001" Then
         fs_Obtener_Tipo = "LC"
      ElseIf p_CodMod = "002" Then
         fs_Obtener_Tipo = "CP"
      End If
   End If
End Function

Private Sub fs_GenExc()
Dim r_obj_Excel         As Excel.Application
Dim r_str_ParAux        As String
Dim r_str_FecRpt        As String
Dim r_int_Contad        As Integer
Dim r_int_NroFil        As Integer
Dim r_int_NoFlLi        As Integer
   
   r_int_NroFil = 8
   Set r_obj_Excel = New Excel.Application
   
   If grd_Listad_Dir.Rows > 0 Then
      r_obj_Excel.SheetsInNewWorkbook = 2
   Else
      r_obj_Excel.SheetsInNewWorkbook = 1
   End If
   r_obj_Excel.Workbooks.Add
   
   'CREDITOS INDIRECTOS
   r_obj_Excel.Sheets(1).Name = "CREDITOS INDIRECTOS"
   
   With r_obj_Excel.Sheets(1)
      .Cells(1, 2) = "REPORTE DE CARTAS FIANZA Y ADENDAS"
      .Range(.Cells(1, 2), .Cells(1, 18)).Merge
      .Range(.Cells(1, 2), .Cells(1, 18)).Font.Bold = True
      .Range(.Cells(1, 2), .Cells(1, 18)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(1, 2), .Cells(1, 18)).Font.Size = 14
        
      .Cells(3, 2) = "FECHA DE EMISIÓN"
      .Cells(3, 3) = "'" & Format(CDate(ipp_FecIni.Text), "dd/mm/yyyy")
      .Cells(3, 4) = " AL "
      .Cells(3, 5) = "'" & Format(CDate(ipp_FecFin.Text), "dd/mm/yyyy")
      .Cells(3, 4).Font.Bold = True
      .Range(.Cells(3, 3), .Cells(3, 5)).HorizontalAlignment = xlHAlignCenter
      
      .Cells(4, 2) = "TIPO DE DOCUMENTO"
      .Cells(4, 3) = Trim(pnl_TipDoc.Caption)
      .Cells(5, 2) = "NRO. DOCUMENTO"
      .Cells(5, 3) = "'" & Trim(pnl_NroDoc.Caption)
      .Cells(6, 2) = "RAZÓN SOCIAL"
      .Cells(6, 3) = Trim(pnl_RazSoc.Caption)
      .Range(.Cells(3, 2), .Cells(6, 2)).Font.Bold = True
        
      .Cells(r_int_NroFil, 2) = "TIPO"
      .Range(.Cells(r_int_NroFil, 2), .Cells(r_int_NroFil + 1, 2)).Merge
      .Cells(r_int_NroFil, 3) = "N° FMV"
      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil + 1, 3)).Merge
      .Cells(r_int_NroFil, 4) = "NÚMERO" ' CF
      .Range(.Cells(r_int_NroFil, 4), .Cells(r_int_NroFil + 1, 4)).Merge
      .Cells(r_int_NroFil, 5) = "FECHA EMISIÓN"
      .Range(.Cells(r_int_NroFil, 5), .Cells(r_int_NroFil + 1, 5)).Merge
      .Cells(r_int_NroFil, 6) = "FECHA     VCTO."
      .Range(.Cells(r_int_NroFil, 6), .Cells(r_int_NroFil + 1, 6)).Merge
      .Cells(r_int_NroFil, 7) = "VALOR" 'CARTA FIANZA"
      .Range(.Cells(r_int_NroFil, 7), .Cells(r_int_NroFil + 1, 7)).Merge
      .Cells(r_int_NroFil, 8) = "GARANTIZADO"
      .Range(.Cells(r_int_NroFil, 8), .Cells(r_int_NroFil + 1, 8)).Merge
      .Cells(r_int_NroFil, 9) = "COMISIONES"
      .Range(.Cells(r_int_NroFil, 9), .Cells(r_int_NroFil, 11)).Merge
      .Cells(r_int_NroFil + 1, 9) = "IMPORTE"
      .Cells(r_int_NroFil + 1, 10) = "PAGADO"
      .Cells(r_int_NroFil + 1, 11) = "SALDO"
      .Cells(r_int_NroFil, 12) = "FONDOS RECIBIDOS"
      .Range(.Cells(r_int_NroFil, 12), .Cells(r_int_NroFil, 14)).Merge
      .Cells(r_int_NroFil + 1, 12) = "IMPORTE"
      .Cells(r_int_NroFil + 1, 13) = "RECIBIDO"
      .Cells(r_int_NroFil + 1, 14) = "SALDO"
      .Cells(r_int_NroFil, 15) = "DESEMBOLSOS - ET"
      .Range(.Cells(r_int_NroFil, 15), .Cells(r_int_NroFil, 17)).Merge
      .Cells(r_int_NroFil + 1, 15) = "IMPORTE"
      .Cells(r_int_NroFil + 1, 16) = "PAGADO"
      .Cells(r_int_NroFil + 1, 17) = "SALDO"
      .Cells(r_int_NroFil, 18) = "ESTADO"
      .Range(.Cells(r_int_NroFil, 18), .Cells(r_int_NroFil + 1, 18)).Merge
      
      .Range(.Cells(r_int_NroFil, 2), .Cells(r_int_NroFil + 1, 18)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(r_int_NroFil, 2), .Cells(r_int_NroFil + 1, 18)).Font.Bold = True
      .Range(.Cells(r_int_NroFil, 2), .Cells(r_int_NroFil + 1, 18)).HorizontalAlignment = xlHAlignCenter
      
      .Columns("A").ColumnWidth = 1
      .Columns("B").ColumnWidth = 19.5
      .Columns("C").ColumnWidth = 13 '19.5
      .Columns("D").ColumnWidth = 13
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").ColumnWidth = 13
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").ColumnWidth = 13
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("G").ColumnWidth = 13.5
      .Columns("G").NumberFormat = "###,###,###,##0.00"
      .Columns("G").HorizontalAlignment = xlHAlignRight
      .Columns("H").ColumnWidth = 13.5
      .Columns("H").NumberFormat = "###,###,###,##0.00"
      .Columns("H").HorizontalAlignment = xlHAlignRight
      .Columns("I").ColumnWidth = 13.5
      .Columns("I").NumberFormat = "###,###,###,##0.00"
      .Columns("I").HorizontalAlignment = xlHAlignRight
      .Columns("J").ColumnWidth = 13.5
      .Columns("J").NumberFormat = "###,###,###,##0.00"
      .Columns("J").HorizontalAlignment = xlHAlignRight
      .Columns("K").ColumnWidth = 13.5
      .Columns("K").NumberFormat = "###,###,###,##0.00"
      .Columns("K").HorizontalAlignment = xlHAlignRight
      .Columns("L").ColumnWidth = 13.5
      .Columns("L").NumberFormat = "###,###,###,##0.00"
      .Columns("L").HorizontalAlignment = xlHAlignRight
      .Columns("M").ColumnWidth = 13.5
      .Columns("M").NumberFormat = "###,###,###,##0.00"
      .Columns("M").HorizontalAlignment = xlHAlignRight
      .Columns("N").ColumnWidth = 13.5
      .Columns("N").NumberFormat = "###,###,###,##0.00"
      .Columns("N").HorizontalAlignment = xlHAlignRight
      .Columns("O").ColumnWidth = 13.5
      .Columns("O").NumberFormat = "###,###,###,##0.00"
      .Columns("O").HorizontalAlignment = xlHAlignRight
      .Columns("P").ColumnWidth = 13.5
      .Columns("P").NumberFormat = "###,###,###,##0.00"
      .Columns("P").HorizontalAlignment = xlHAlignRight
      .Columns("Q").ColumnWidth = 13.5
      .Columns("Q").NumberFormat = "###,###,###,##0.00"
      .Columns("Q").HorizontalAlignment = xlHAlignRight
      .Columns("R").ColumnWidth = 13.5
      .Columns("R").HorizontalAlignment = xlHAlignCenter
        
      With .Range(.Cells(8, 2), .Cells(9, 18))
          .HorizontalAlignment = xlCenter
          .VerticalAlignment = xlCenter
          .WrapText = True
      End With
      
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Name = "Calibri"
      .Range(.Cells(2, 1), .Cells(99, 99)).Font.Size = 11
       
      r_int_NroFil = r_int_NroFil + 2
      For r_int_NoFlLi = 0 To grd_Listad.Rows - 1
          .Cells(r_int_NroFil, 2) = "'" & grd_Listad.TextMatrix(r_int_NoFlLi, 0)
          .Cells(r_int_NroFil, 3) = "'" & grd_Listad.TextMatrix(r_int_NoFlLi, 1)
          .Cells(r_int_NroFil, 4) = "'" & grd_Listad.TextMatrix(r_int_NoFlLi, 2)
          .Cells(r_int_NroFil, 5) = "'" & grd_Listad.TextMatrix(r_int_NoFlLi, 3)
          .Cells(r_int_NroFil, 6) = "'" & grd_Listad.TextMatrix(r_int_NoFlLi, 4)
          .Cells(r_int_NroFil, 7) = grd_Listad.TextMatrix(r_int_NoFlLi, 5)
          .Cells(r_int_NroFil, 8) = grd_Listad.TextMatrix(r_int_NoFlLi, 6)
          .Cells(r_int_NroFil, 9) = grd_Listad.TextMatrix(r_int_NoFlLi, 7)
          .Cells(r_int_NroFil, 10) = grd_Listad.TextMatrix(r_int_NoFlLi, 8)
          .Cells(r_int_NroFil, 11) = grd_Listad.TextMatrix(r_int_NoFlLi, 9)
          .Cells(r_int_NroFil, 12) = grd_Listad.TextMatrix(r_int_NoFlLi, 10)
          .Cells(r_int_NroFil, 13) = grd_Listad.TextMatrix(r_int_NoFlLi, 11)
          .Cells(r_int_NroFil, 14) = grd_Listad.TextMatrix(r_int_NoFlLi, 12)
          .Cells(r_int_NroFil, 15) = grd_Listad.TextMatrix(r_int_NoFlLi, 13)
          .Cells(r_int_NroFil, 16) = grd_Listad.TextMatrix(r_int_NoFlLi, 14)
          .Cells(r_int_NroFil, 17) = grd_Listad.TextMatrix(r_int_NoFlLi, 15)
          .Cells(r_int_NroFil, 18) = grd_Listad.TextMatrix(r_int_NoFlLi, 16)
         
          r_int_NroFil = r_int_NroFil + 1
      Next r_int_NoFlLi
      
      With .Range(.Cells(10, 2), .Cells(r_int_NroFil, 3))
          .HorizontalAlignment = xlCenter
          .VerticalAlignment = xlCenter
          .WrapText = True
      End With
      With .Range(.Cells(1, 2), .Cells(1, 18))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous
          .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous
          .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous
          .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous
          .Borders(xlEdgeRight).Weight = xlMedium
      End With
      With .Range(.Cells(8, 2), .Cells(r_int_NroFil - 1, 8))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous
          .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous
          .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous
          .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous
          .Borders(xlEdgeRight).Weight = xlMedium
          .Borders(xlInsideVertical).LineStyle = xlContinuous
          .Borders(xlInsideHorizontal).LineStyle = xlContinuous
      End With
      With .Range(.Cells(8, 9), .Cells(r_int_NroFil - 1, 11))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous
          .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous
          .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous
          .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous
          .Borders(xlEdgeRight).Weight = xlMedium
          .Borders(xlInsideVertical).LineStyle = xlContinuous
          .Borders(xlInsideHorizontal).LineStyle = xlContinuous
      End With
      With .Range(.Cells(8, 12), .Cells(r_int_NroFil - 1, 14))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous
          .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous
          .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous
          .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous
          .Borders(xlEdgeRight).Weight = xlMedium
          .Borders(xlInsideVertical).LineStyle = xlContinuous
          .Borders(xlInsideHorizontal).LineStyle = xlContinuous
      End With
      With .Range(.Cells(8, 15), .Cells(r_int_NroFil - 1, 17))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous
          .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous
          .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous
          .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous
          .Borders(xlEdgeRight).Weight = xlMedium
          .Borders(xlInsideVertical).LineStyle = xlContinuous
          .Borders(xlInsideHorizontal).LineStyle = xlContinuous
      End With

      With .Range(.Cells(8, 18), .Cells(r_int_NroFil - 1, 18))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous
          .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous
          .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous
          .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous
          .Borders(xlEdgeRight).Weight = xlMedium
          .Borders(xlInsideVertical).LineStyle = xlContinuous
          .Borders(xlInsideHorizontal).LineStyle = xlContinuous
      End With
      With .Range(.Cells(8, 2), .Cells(9, 18))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous
          .Borders(xlEdgeLeft).Weight = xlMedium
          .Borders(xlEdgeTop).LineStyle = xlContinuous
          .Borders(xlEdgeTop).Weight = xlMedium
          .Borders(xlEdgeBottom).LineStyle = xlContinuous
          .Borders(xlEdgeBottom).Weight = xlMedium
          .Borders(xlEdgeRight).LineStyle = xlContinuous
          .Borders(xlEdgeRight).Weight = xlMedium
          .Borders(xlInsideHorizontal).LineStyle = xlContinuous
          .Borders(xlInsideHorizontal).Weight = xlThin
      End With
   End With
   
   r_int_NroFil = 8
   'CREDITOS DIRECTOS
   If grd_Listad_Dir.Rows > 0 Then
      r_obj_Excel.Sheets(2).Name = "CREDITOS DIRECTOS"
       
      With r_obj_Excel.Sheets(2)
         .Cells(1, 2) = "REPORTE DE CREDITOS DIRECTOS"
         .Range(.Cells(1, 2), .Cells(1, 15)).Merge
         .Range(.Cells(1, 2), .Cells(1, 15)).Font.Bold = True
         .Range(.Cells(1, 2), .Cells(1, 15)).HorizontalAlignment = xlHAlignCenter
         .Range(.Cells(1, 2), .Cells(1, 15)).Font.Size = 14
         
         .Cells(3, 2) = "FECHA DE EMISIÓN"
         .Cells(3, 3) = "'" & Format(CDate(ipp_FecIni.Text), "dd/mm/yyyy")
         .Cells(3, 4) = " AL "
         .Cells(3, 5) = "'" & Format(CDate(ipp_FecFin.Text), "dd/mm/yyyy")
         .Cells(3, 4).Font.Bold = True
         .Range(.Cells(3, 3), .Cells(3, 5)).HorizontalAlignment = xlHAlignCenter
         
         .Cells(4, 2) = "TIPO DE DOCUMENTO"
         .Cells(4, 3) = Trim(pnl_TipDoc.Caption)
         .Cells(5, 2) = "NRO. DOCUMENTO"
         .Cells(5, 3) = "'" & Trim(pnl_NroDoc.Caption)
         .Cells(6, 2) = "RAZÓN SOCIAL"
         .Cells(6, 3) = Trim(pnl_RazSoc.Caption)
         .Range(.Cells(3, 2), .Cells(6, 2)).Font.Bold = True
         
         .Cells(r_int_NroFil, 2) = "TIPO"
         .Range(.Cells(r_int_NroFil, 2), .Cells(r_int_NroFil + 1, 2)).Merge
         .Cells(r_int_NroFil, 3) = "NÚMERO"
         .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil + 1, 3)).Merge
         .Cells(r_int_NroFil, 4) = "FECHA EMISIÓN"
         .Range(.Cells(r_int_NroFil, 4), .Cells(r_int_NroFil + 1, 4)).Merge
         .Cells(r_int_NroFil, 5) = "FECHA     VCTO."
         .Range(.Cells(r_int_NroFil, 5), .Cells(r_int_NroFil + 1, 5)).Merge
         .Cells(r_int_NroFil, 6) = "TASA INTERES"
         .Range(.Cells(r_int_NroFil, 6), .Cells(r_int_NroFil + 1, 6)).Merge
         .Cells(r_int_NroFil, 7) = "TASA MORATORIA"
         .Range(.Cells(r_int_NroFil, 7), .Cells(r_int_NroFil + 1, 7)).Merge
         .Cells(r_int_NroFil, 8) = "MONTO PRESTAMO"
         .Range(.Cells(r_int_NroFil, 8), .Cells(r_int_NroFil + 1, 8)).Merge
         .Cells(r_int_NroFil, 9) = "COMISIONES"
         .Range(.Cells(r_int_NroFil, 9), .Cells(r_int_NroFil, 11)).Merge
         .Cells(r_int_NroFil + 1, 9) = "IMPORTE"
         .Cells(r_int_NroFil + 1, 10) = "PAGADO"
         .Cells(r_int_NroFil + 1, 11) = "SALDO"
         .Cells(r_int_NroFil, 12) = "DESEMBOLSOS - ET"
         .Range(.Cells(r_int_NroFil, 12), .Cells(r_int_NroFil, 14)).Merge
         .Cells(r_int_NroFil + 1, 12) = "IMPORTE"
         .Cells(r_int_NroFil + 1, 13) = "PAGADO"
         .Cells(r_int_NroFil + 1, 14) = "SALDO"
         .Cells(r_int_NroFil, 15) = "ESTADO"
         .Range(.Cells(r_int_NroFil, 15), .Cells(r_int_NroFil + 1, 15)).Merge
         
         .Range(.Cells(r_int_NroFil, 2), .Cells(r_int_NroFil + 1, 15)).Interior.Color = RGB(146, 208, 80)
         .Range(.Cells(r_int_NroFil, 2), .Cells(r_int_NroFil + 1, 15)).Font.Bold = True
         .Range(.Cells(r_int_NroFil, 2), .Cells(r_int_NroFil + 1, 15)).HorizontalAlignment = xlHAlignCenter
           
         .Columns("A").ColumnWidth = 1
         .Columns("B").ColumnWidth = 19.5
         .Columns("C").ColumnWidth = 13 '19.5
         .Columns("D").ColumnWidth = 13
         .Columns("D").HorizontalAlignment = xlHAlignCenter
         .Columns("E").ColumnWidth = 13
         .Columns("E").HorizontalAlignment = xlHAlignCenter
         .Columns("F").ColumnWidth = 13
         .Columns("F").NumberFormat = "###,###,###,##0.00"
         .Columns("F").HorizontalAlignment = xlHAlignCenter
         .Columns("G").ColumnWidth = 13.5
         .Columns("G").NumberFormat = "###,###,###,##0.00"
         .Columns("G").HorizontalAlignment = xlHAlignRight
         .Columns("H").ColumnWidth = 13.5
         .Columns("H").NumberFormat = "###,###,###,##0.00"
         .Columns("H").HorizontalAlignment = xlHAlignRight
         .Columns("I").ColumnWidth = 13.5
         .Columns("I").NumberFormat = "###,###,###,##0.00"
         .Columns("I").HorizontalAlignment = xlHAlignRight
         .Columns("J").ColumnWidth = 13.5
         .Columns("J").NumberFormat = "###,###,###,##0.00"
         .Columns("J").HorizontalAlignment = xlHAlignRight
         .Columns("K").ColumnWidth = 13.5
         .Columns("K").NumberFormat = "###,###,###,##0.00"
         .Columns("K").HorizontalAlignment = xlHAlignRight
         .Columns("L").ColumnWidth = 13.5
         .Columns("L").NumberFormat = "###,###,###,##0.00"
         .Columns("L").HorizontalAlignment = xlHAlignRight
         .Columns("M").ColumnWidth = 13.5
         .Columns("M").NumberFormat = "###,###,###,##0.00"
         .Columns("M").HorizontalAlignment = xlHAlignRight
         .Columns("N").ColumnWidth = 13.5
         .Columns("N").NumberFormat = "###,###,###,##0.00"
         .Columns("N").HorizontalAlignment = xlHAlignRight
         .Columns("O").ColumnWidth = 13.5
         .Columns("O").HorizontalAlignment = xlHAlignCenter
         
         With .Range(.Cells(8, 2), .Cells(9, 15))
             .HorizontalAlignment = xlCenter
             .VerticalAlignment = xlCenter
             .WrapText = True
         End With
         
         .Range(.Cells(1, 1), .Cells(99, 99)).Font.Name = "Calibri"
         .Range(.Cells(2, 1), .Cells(99, 99)).Font.Size = 11
          
         r_int_NroFil = r_int_NroFil + 2
         For r_int_NoFlLi = 0 To grd_Listad_Dir.Rows - 1
             .Cells(r_int_NroFil, 2) = "'" & grd_Listad_Dir.TextMatrix(r_int_NoFlLi, 0)
             .Cells(r_int_NroFil, 3) = "'" & grd_Listad_Dir.TextMatrix(r_int_NoFlLi, 1)
             .Cells(r_int_NroFil, 4) = "'" & grd_Listad_Dir.TextMatrix(r_int_NoFlLi, 2)
             .Cells(r_int_NroFil, 5) = "'" & grd_Listad_Dir.TextMatrix(r_int_NoFlLi, 3)
             .Cells(r_int_NroFil, 6) = grd_Listad_Dir.TextMatrix(r_int_NoFlLi, 4)
             .Cells(r_int_NroFil, 7) = grd_Listad_Dir.TextMatrix(r_int_NoFlLi, 5)
             .Cells(r_int_NroFil, 8) = grd_Listad_Dir.TextMatrix(r_int_NoFlLi, 6)
             .Cells(r_int_NroFil, 9) = grd_Listad_Dir.TextMatrix(r_int_NoFlLi, 7)
             .Cells(r_int_NroFil, 10) = grd_Listad_Dir.TextMatrix(r_int_NoFlLi, 8)
             .Cells(r_int_NroFil, 11) = grd_Listad_Dir.TextMatrix(r_int_NoFlLi, 9)
             .Cells(r_int_NroFil, 12) = grd_Listad_Dir.TextMatrix(r_int_NoFlLi, 10)
             .Cells(r_int_NroFil, 13) = grd_Listad_Dir.TextMatrix(r_int_NoFlLi, 11)
             .Cells(r_int_NroFil, 14) = grd_Listad_Dir.TextMatrix(r_int_NoFlLi, 12)
             .Cells(r_int_NroFil, 15) = grd_Listad_Dir.TextMatrix(r_int_NoFlLi, 13)
            
             r_int_NroFil = r_int_NroFil + 1
         Next r_int_NoFlLi
         
         With .Range(.Cells(10, 2), .Cells(r_int_NroFil, 3))
             .HorizontalAlignment = xlCenter
             .VerticalAlignment = xlCenter
             .WrapText = True
         End With
         With .Range(.Cells(1, 2), .Cells(1, 15))
             .Borders(xlEdgeLeft).LineStyle = xlContinuous
             .Borders(xlEdgeLeft).Weight = xlMedium
             .Borders(xlEdgeTop).LineStyle = xlContinuous
             .Borders(xlEdgeTop).Weight = xlMedium
             .Borders(xlEdgeBottom).LineStyle = xlContinuous
             .Borders(xlEdgeBottom).Weight = xlMedium
             .Borders(xlEdgeRight).LineStyle = xlContinuous
             .Borders(xlEdgeRight).Weight = xlMedium
         End With
         With .Range(.Cells(8, 2), .Cells(r_int_NroFil - 1, 15))
             .Borders(xlEdgeLeft).LineStyle = xlContinuous
             .Borders(xlEdgeTop).LineStyle = xlContinuous
             .Borders(xlEdgeBottom).LineStyle = xlContinuous
             .Borders(xlEdgeRight).LineStyle = xlContinuous
             .Borders(xlInsideVertical).LineStyle = xlContinuous
             .Borders(xlInsideHorizontal).LineStyle = xlContinuous
         End With
         With .Range(.Cells(8, 2), .Cells(r_int_NroFil - 1, 8))
             .Borders(xlEdgeLeft).LineStyle = xlContinuous
             .Borders(xlEdgeLeft).Weight = xlMedium
             .Borders(xlEdgeTop).LineStyle = xlContinuous
             .Borders(xlEdgeTop).Weight = xlMedium
             .Borders(xlEdgeBottom).LineStyle = xlContinuous
             .Borders(xlEdgeBottom).Weight = xlMedium
             .Borders(xlEdgeRight).LineStyle = xlContinuous
             .Borders(xlEdgeRight).Weight = xlMedium
             .Borders(xlInsideVertical).LineStyle = xlContinuous
             .Borders(xlInsideHorizontal).LineStyle = xlContinuous
         End With
         With .Range(.Cells(8, 9), .Cells(r_int_NroFil - 1, 11))
             .Borders(xlEdgeLeft).LineStyle = xlContinuous
             .Borders(xlEdgeLeft).Weight = xlMedium
             .Borders(xlEdgeTop).LineStyle = xlContinuous
             .Borders(xlEdgeTop).Weight = xlMedium
             .Borders(xlEdgeBottom).LineStyle = xlContinuous
             .Borders(xlEdgeBottom).Weight = xlMedium
             .Borders(xlEdgeRight).LineStyle = xlContinuous
             .Borders(xlEdgeRight).Weight = xlMedium
             .Borders(xlInsideVertical).LineStyle = xlContinuous
             .Borders(xlInsideHorizontal).LineStyle = xlContinuous
         End With
         With .Range(.Cells(8, 12), .Cells(r_int_NroFil - 1, 14))
             .Borders(xlEdgeLeft).LineStyle = xlContinuous
             .Borders(xlEdgeLeft).Weight = xlMedium
             .Borders(xlEdgeTop).LineStyle = xlContinuous
             .Borders(xlEdgeTop).Weight = xlMedium
             .Borders(xlEdgeBottom).LineStyle = xlContinuous
             .Borders(xlEdgeBottom).Weight = xlMedium
             .Borders(xlEdgeRight).LineStyle = xlContinuous
             .Borders(xlEdgeRight).Weight = xlMedium
             .Borders(xlInsideVertical).LineStyle = xlContinuous
             .Borders(xlInsideHorizontal).LineStyle = xlContinuous
         End With
   
         With .Range(.Cells(8, 15), .Cells(r_int_NroFil - 1, 15))
             .Borders(xlEdgeLeft).LineStyle = xlContinuous
             .Borders(xlEdgeLeft).Weight = xlMedium
             .Borders(xlEdgeTop).LineStyle = xlContinuous
             .Borders(xlEdgeTop).Weight = xlMedium
             .Borders(xlEdgeBottom).LineStyle = xlContinuous
             .Borders(xlEdgeBottom).Weight = xlMedium
             .Borders(xlEdgeRight).LineStyle = xlContinuous
             .Borders(xlEdgeRight).Weight = xlMedium
             .Borders(xlInsideVertical).LineStyle = xlContinuous
             .Borders(xlInsideHorizontal).LineStyle = xlContinuous
         End With
         
         With .Range(.Cells(8, 2), .Cells(9, 15))
             .Borders(xlEdgeLeft).LineStyle = xlContinuous
             .Borders(xlEdgeLeft).Weight = xlMedium
             .Borders(xlEdgeTop).LineStyle = xlContinuous
             .Borders(xlEdgeTop).Weight = xlMedium
             .Borders(xlEdgeBottom).LineStyle = xlContinuous
             .Borders(xlEdgeBottom).Weight = xlMedium
             .Borders(xlEdgeRight).LineStyle = xlContinuous
             .Borders(xlEdgeRight).Weight = xlMedium
             .Borders(xlInsideHorizontal).LineStyle = xlContinuous
             .Borders(xlInsideHorizontal).Weight = xlThin
         End With
      End With
   End If
   r_obj_Excel.Visible = True
End Sub
Private Sub fs_GenLiqCarFia()
Dim r_obj_Excel         As Excel.Application
Dim r_str_CorLiq        As String
Dim r_str_NumRef        As String
Dim r_str_NumAux        As String

Dim l_Mar_Izq           As Double
Dim l_Mar_Der           As Double
Dim l_Mar_Sup           As Double
Dim l_Mar_Inf           As Double

    'MARGENES DE IMPRESIÓN
   l_Mar_Izq = 1
   l_Mar_Der = 0.8
   l_Mar_Sup = 1.9
   l_Mar_Inf = 1.9
   
   r_str_NumRef = grd_Listad.TextMatrix(grd_Listad.Row, 2)
   
   If Mid(r_str_NumRef, 1, 1) = 0 Then
      r_str_NumAux = "1" & Mid(r_str_NumRef, 11) & Mid(r_str_NumRef, 6, 2) & Right("00000" & Mid(r_str_NumRef, 1, 4), 5)
      r_str_NumAux = Replace(r_str_NumAux, "-", "")
   Else
      r_str_NumAux = Replace(r_str_NumRef, "-", "")
   End If
   
   'Obtiene el correlativo de liquidación
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT (CASE WHEN (SELECT MAECFI_CORLIQ FROM TPR_MAECFI WHERE MAECFI_NUMREF = '" & r_str_NumAux & "') IS NULL THEN "
   g_str_Parame = g_str_Parame & "            (SELECT MAX(NVL(SUBSTR(MAECFI_CORLIQ,5),0)) FROM TPR_MAECFI) "
   g_str_Parame = g_str_Parame & "         ELSE "
   g_str_Parame = g_str_Parame & "            (SELECT MAECFI_CORLIQ FROM TPR_MAECFI WHERE MAECFI_NUMREF = '" & r_str_NumAux & "') "
   g_str_Parame = g_str_Parame & "          END) AS Correlativo "
   g_str_Parame = g_str_Parame & "   FROM DUAL "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
      Exit Sub
   End If
      
   If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
      If g_rst_GenAux!CORRELATIVO = 0 Or Len(g_rst_GenAux!CORRELATIVO) = 4 Then
         r_str_CorLiq = Year(Now) & "-" & Right("0000" & CInt(g_rst_GenAux!CORRELATIVO) + 1, 4)
      Else
         r_str_CorLiq = Mid(g_rst_GenAux!CORRELATIVO, 1, 4) & "-" & Mid(g_rst_GenAux!CORRELATIVO, 5, 4)
      End If
   End If
   
   g_rst_GenAux.Close
   Set g_rst_GenAux = Nothing
   
      
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT MAEETE_DIRREP, MAEETE_UBIGEO, MAECFI_TASFIA, MAECFI_EMIFIA, MAECFI_VTOFIA , MAECFI_GARFIA, MAECFI_CORLIQ "
   g_str_Parame = g_str_Parame & "    FROM TPR_MAEETE "
   g_str_Parame = g_str_Parame & "         INNER JOIN TPR_MAECFI ON MAECFI_TIPDOC = MAEETE_TIPDOC AND MAECFI_NUMDOC = MAEETE_NUMDOC "
   g_str_Parame = g_str_Parame & "         AND MAECFI_NUMREF = '" & r_str_NumAux & "' "
   g_str_Parame = g_str_Parame & "   WHERE MAEETE_TIPDOC = " & Mid(pnl_TipDoc.Caption, 1, 1) & " "
   g_str_Parame = g_str_Parame & "     AND MAEETE_NUMDOC = '" & CStr(pnl_NroDoc.Caption) & "'"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
      Exit Sub
   End If
   
   If g_rst_GenAux.BOF And g_rst_GenAux.EOF Then
      g_rst_GenAux.Close
      Set g_rst_GenAux = Nothing
      Exit Sub
   End If
   
   If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
      
      Set r_obj_Excel = New Excel.Application
      r_obj_Excel.SheetsInNewWorkbook = 1
      r_obj_Excel.Workbooks.Add
      
      With r_obj_Excel.ActiveSheet.PageSetup
      
         If .Orientation = xlPortrait Then
            .Orientation = xlPortrait
         Else
            .Orientation = xlLandscape
         End If
         
         'Configuración de márgenes:
         .LeftMargin = Application.CentimetersToPoints(l_Mar_Izq) '
         .RightMargin = Application.CentimetersToPoints(l_Mar_Der)
         .TopMargin = Application.CentimetersToPoints(l_Mar_Sup)
         .BottomMargin = Application.CentimetersToPoints(l_Mar_Inf)
            
         'AJUSTE DE ESCALA
         .Zoom = 70
         
         '.CenterHorizontally = True
         .CenterVertically = True
         .PrintGridlines = False
         .PrintArea = ""
      End With
   
      With r_obj_Excel.ActiveSheet
         .Columns("A:A").Range(.Columns("A:A"), .Columns("A:A").End(xlToRight)).ColumnWidth = 3.43
         
         With .Cells.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
            .PatternTintAndShade = 0
         End With
         
         .Cells(4, 2) = "LIQUIDACION DE CARTA FIANZA"
         .Range(.Cells(4, 2), .Cells(4, 30)).Merge
         .Range(.Cells(4, 2), .Cells(4, 30)).Font.Bold = True
         .Range(.Cells(4, 2), .Cells(4, 30)).HorizontalAlignment = xlHAlignCenter
         .Range(.Cells(4, 2), .Cells(4, 30)).Font.Size = 11
         
         .Cells(6, 2) = "EDYPYME MICASITA S.A."
         .Cells(6, 31) = "'" & Format(CDate(Now), "dd/mm/yyyy")
         .Range(.Cells(6, 30), .Cells(6, 31)).HorizontalAlignment = xlHAlignRight
         .Cells(7, 2) = "RUC 20511904162"
         .Cells(7, 31) = r_str_CorLiq
         .Range(.Cells(7, 30), .Cells(7, 31)).HorizontalAlignment = xlHAlignRight
         .Cells(8, 2) = "'" & "---------------------------------------------------------------------------------------------------------------"
            
         .Cells(9, 2) = "Cliente:"
         .Cells(9, 7) = Trim(pnl_RazSoc.Caption)
         .Cells(9, 23) = "Forma de Cobro:"
         .Cells(9, 31) = "S/."
         .Range(.Cells(9, 30), .Cells(9, 31)).HorizontalAlignment = xlHAlignRight
         
         .Cells(10, 2) = "Dirección:"
         .Cells(10, 7) = "'" & IIf(IsNull(g_rst_GenAux!MAEETE_DIRREP), "", Trim(g_rst_GenAux!MAEETE_DIRREP))
         .Cells(10, 23) = "DEPOSITO EN CUENTA"
         .Cells(10, 31) = "'" & grd_Listad.TextMatrix(grd_Listad.Row, 8)
         .Range(.Cells(10, 30), .Cells(10, 31)).HorizontalAlignment = xlHAlignRight
         
         If Not IsNull(g_rst_GenAux!MAEETE_UBIGEO) Then
            .Cells(11, 7) = "'" & moddat_gf_Consulta_ParDes("101", Trim(g_rst_GenAux!MAEETE_UBIGEO)) & " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_GenAux!MAEETE_UBIGEO, 4) & "00") & " - " & moddat_gf_Consulta_ParDes("101", Left(g_rst_GenAux!MAEETE_UBIGEO, 2) & "0000")
         End If
         
         .Cells(11, 18) = "RUC"
         .Cells(11, 19) = "'" & Trim(pnl_NroDoc.Caption)
         .Cells(11, 23) = "'" & "0011-0661-01-00040896"
         
         .Cells(12, 2) = "'" & "---------------------------------------------------------------------------------------------------------------"
         
         .Cells(13, 2) = "Referencia:"
         .Cells(13, 7) = "Emisión Carta Fianza"
         .Cells(14, 2) = "Número:"
         .Cells(14, 7) = r_str_NumRef
         .Cells(14, 12) = "'" & "Tasa Operación:"
         .Cells(14, 17) = Format(g_rst_GenAux!MAECFI_TASFIA, "##0.0000")
         .Cells(14, 17).NumberFormat = "0.00"
         .Cells(14, 23) = "Fecha Vencimiento:"
         .Cells(14, 31) = "'" & Format(gf_FormatoFecha(CStr(g_rst_GenAux!MAECFI_VTOFIA)), "dd/mm/yyyy")
         .Range(.Cells(14, 30), .Cells(14, 31)).HorizontalAlignment = xlHAlignRight
         .Cells(15, 2) = "Frecuencia Pago:"
         .Cells(15, 7) = "A LA EMISIÓN"
         
         .Cells(17, 2) = "Importe:"
         .Cells(17, 13) = "'" & Format(CStr(g_rst_GenAux!MAECFI_GARFIA), "###,###,###,##0.00")
         .Cells(17, 16) = "SOL"
         .Cells(18, 2) = "'" & "---------------------------------------------------------"
         
         .Cells(19, 2) = "C O N C E P T O"
         .Range(.Cells(19, 2), .Cells(19, 7)).Merge
         .Range(.Cells(19, 2), .Cells(19, 7)).HorizontalAlignment = xlHAlignCenter
         
         .Cells(19, 12) = "I M P O R T E S/."
         .Range(.Cells(19, 12), .Cells(19, 16)).Merge
         .Range(.Cells(19, 12), .Cells(19, 16)).HorizontalAlignment = xlHAlignCenter
         
         .Cells(20, 2) = "COMISION"
         .Cells(20, 9) = DateDiff("d", Format(gf_FormatoFecha(CStr(g_rst_GenAux!MAECFI_EMIFIA)), "dd/mm/yyyy"), Format(gf_FormatoFecha(CStr(g_rst_GenAux!MAECFI_VTOFIA)), "dd/mm/yyyy")) & "d"
         .Range(.Cells(20, 9), .Cells(20, 10)).Merge
         .Range(.Cells(20, 9), .Cells(20, 10)).HorizontalAlignment = xlHAlignCenter
         .Cells(20, 16).FormulaR1C1 = "=+R[-10]C[15]"
         .Range(.Cells(20, 16), .Cells(20, 16)).HorizontalAlignment = xlHAlignRight
         
         .Cells(22, 5) = "TOTAL A COBRAR …………:"
         .Cells(22, 11) = "S/."
         .Cells(22, 16).FormulaR1C1 = "=+R[-2]C"
         .Range(.Cells(22, 16), .Cells(22, 16)).HorizontalAlignment = xlHAlignRight
         
         .Cells(23, 2) = "'" & "---------------------------------------------------------"
         .Cells(23, 20) = "'" & "--------------"
         .Cells(23, 27) = "'" & "--------------"
         
         .Cells(24, 2) = "Beneficiario    FONDO MIVIVIENDA S.A."
         .Cells(24, 21) = "VoBo"
         .Range(.Cells(24, 21), .Cells(24, 22)).Merge
         .Cells(24, 28) = "Cliente"
         .Range(.Cells(24, 28), .Cells(24, 29)).Merge
                
         .Rows("8:8").RowHeight = 8.25
         .Rows("12:12").RowHeight = 8.25
         .Rows("18:18").RowHeight = 8.25
         .Rows("23:23").RowHeight = 8.25
         .Range(.Cells(1, 1), .Cells(99, 99)).Font.Name = "Consolas" '"Calibri"
         .Range(.Cells(2, 1), .Cells(99, 99)).Font.Size = 11
         
         With .Range(.Cells(19, 2), .Cells(19, 7))
            With .Borders(xlEdgeBottom)
               .LineStyle = xlContinuous
               .ColorIndex = 0
               .TintAndShade = 0
               .Weight = xlThin
            End With
         End With
    
         With .Range(.Cells(19, 12), .Cells(19, 16))
            With .Borders(xlEdgeBottom)
               .LineStyle = xlContinuous
               .ColorIndex = 0
               .TintAndShade = 0
               .Weight = xlThin
            End With
         End With
         
         With .Range(.Cells(22, 12), .Cells(22, 16))
            With .Borders(xlEdgeTop)
               .LineStyle = xlContinuous
               .ColorIndex = 0
               .TintAndShade = 0
               .Weight = xlThin
            End With
         End With

         'COPIA DE LIQUIDACION
         .Cells(44, 2) = "LIQUIDACION DE CARTA FIANZA"
         .Range(.Cells(44, 2), .Cells(44, 30)).Merge
         .Range(.Cells(44, 2), .Cells(44, 30)).Font.Bold = True
         .Range(.Cells(44, 2), .Cells(44, 30)).HorizontalAlignment = xlHAlignCenter
         .Range(.Cells(44, 2), .Cells(44, 30)).Font.Size = 11
         
         .Cells(46, 2) = "EDYPYME MICASITA S.A."
         .Cells(46, 31).FormulaR1C1 = "=+R[-40]C"
         .Range(.Cells(46, 30), .Cells(46, 31)).HorizontalAlignment = xlHAlignRight
         .Cells(47, 2) = "RUC 20511904162"
         .Cells(47, 31).FormulaR1C1 = "=+R[-40]C"
         .Range(.Cells(47, 30), .Cells(47, 31)).HorizontalAlignment = xlHAlignRight
         .Cells(48, 2) = "'" & "---------------------------------------------------------------------------------------------------------------"
            
         .Cells(49, 2) = "Cliente:"
         .Cells(49, 7).FormulaR1C1 = "=+R[-40]C"
         .Cells(49, 23) = "Forma de Cobro:"
         .Cells(49, 31) = "=+R[-40]C"
         .Range(.Cells(49, 30), .Cells(49, 31)).HorizontalAlignment = xlHAlignRight
         
         .Cells(50, 2) = "Dirección:"
         .Cells(50, 7).FormulaR1C1 = "=+R[-40]C"
         .Cells(50, 23) = "DEPOSITO EN CUENTA"
         .Cells(50, 31).FormulaR1C1 = "=+R[-40]C"
         .Range(.Cells(50, 30), .Cells(50, 31)).HorizontalAlignment = xlHAlignRight
         
         .Cells(51, 7).FormulaR1C1 = "=+R[-40]C"
         .Cells(51, 18) = "RUC"
         .Cells(51, 19).FormulaR1C1 = "=+R[-40]C"
         .Cells(51, 23) = "'" & "0011-0661-01-00040896"
         
         .Cells(52, 2) = "'" & "---------------------------------------------------------------------------------------------------------------"
         
         .Cells(53, 2) = "Referencia:"
         .Cells(53, 7) = "Emisión Carta Fianza"
         .Cells(54, 2) = "Número:"
         .Cells(54, 7).FormulaR1C1 = "=+R[-40]C"
         .Cells(54, 12) = "'" & "Tasa Operación:"
         .Cells(54, 17).FormulaR1C1 = "=+R[-40]C"
         .Cells(54, 23) = "Fecha Vencimiento:"
         .Cells(54, 31).FormulaR1C1 = "=+R[-40]C"
         .Range(.Cells(54, 30), .Cells(54, 31)).HorizontalAlignment = xlHAlignRight
         .Cells(55, 2) = "Frecuencia Pago:"
         .Cells(55, 7) = "A LA EMISIÓN"
         
         .Cells(57, 2) = "Importe:"
         .Cells(57, 13).FormulaR1C1 = "=+R[-40]C"
         .Cells(57, 16) = "SOL"
         .Cells(58, 2) = "'" & "---------------------------------------------------------"
         
         .Cells(59, 2) = "C O N C E P T O"
         .Range(.Cells(59, 2), .Cells(59, 7)).Merge
         .Range(.Cells(59, 2), .Cells(59, 7)).HorizontalAlignment = xlHAlignCenter
         
         .Cells(59, 12) = "I M P O R T E S/."
         .Range(.Cells(59, 12), .Cells(59, 16)).Merge
         .Range(.Cells(59, 12), .Cells(59, 16)).HorizontalAlignment = xlHAlignCenter
         
         .Cells(60, 2) = "COMISION"
         .Cells(60, 9).FormulaR1C1 = "=+R[-40]C"
         .Range(.Cells(60, 9), .Cells(60, 10)).Merge
         .Range(.Cells(60, 9), .Cells(60, 10)).HorizontalAlignment = xlHAlignCenter
         .Cells(60, 16).FormulaR1C1 = "=+R[-10]C[15]"
         .Range(.Cells(60, 16), .Cells(60, 16)).HorizontalAlignment = xlHAlignRight
         
         .Cells(62, 5) = "TOTAL A COBRAR …………:"
         .Cells(62, 11) = "S/."
         .Cells(62, 16).FormulaR1C1 = "=+R[-2]C"
         .Range(.Cells(62, 16), .Cells(62, 16)).HorizontalAlignment = xlHAlignRight
         
         .Cells(63, 2) = "'" & "---------------------------------------------------------"
         .Cells(63, 20) = "'" & "--------------"
         .Cells(63, 27) = "'" & "--------------"
         
         .Cells(64, 2) = "Beneficiario    FONDO MIVIVIENDA S.A."
         .Cells(64, 21) = "VoBo"
         .Range(.Cells(64, 21), .Cells(64, 22)).Merge
         .Cells(64, 28) = "Cliente"
         .Range(.Cells(64, 28), .Cells(64, 29)).Merge
                
         .Rows("48:48").RowHeight = 8.25
         .Rows("52:52").RowHeight = 8.25
         .Rows("58:58").RowHeight = 8.25
         .Rows("63:63").RowHeight = 8.25
         .Range(.Cells(1, 1), .Cells(99, 99)).Font.Name = "Consolas" '"Calibri"
         .Range(.Cells(2, 1), .Cells(99, 99)).Font.Size = 11
         
         With .Range(.Cells(59, 2), .Cells(59, 7))
            With .Borders(xlEdgeBottom)
               .LineStyle = xlContinuous
               .ColorIndex = 0
               .TintAndShade = 0
               .Weight = xlThin
            End With
         End With
         
         With .Range(.Cells(59, 12), .Cells(59, 16))
            With .Borders(xlEdgeBottom)
               .LineStyle = xlContinuous
               .ColorIndex = 0
               .TintAndShade = 0
               .Weight = xlThin
            End With
         End With
         
         With .Range(.Cells(62, 12), .Cells(62, 16))
            With .Borders(xlEdgeTop)
               .LineStyle = xlContinuous
               .ColorIndex = 0
               .TintAndShade = 0
               .Weight = xlThin
            End With
         End With
      End With
   End If
   
   
   'Actualiza correlativo
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "UPDATE TPR_MAECFI "
   g_str_Parame = g_str_Parame & "   SET MAECFI_CORLIQ = '" & Replace(r_str_CorLiq, "-", "") & "' "
   g_str_Parame = g_str_Parame & " WHERE MAECFI_NUMREF = '" & r_str_NumAux & "' "
   g_str_Parame = g_str_Parame & "   AND MAECFI_CORLIQ IS NULL "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 2) Then
      MsgBox "Error al actualzar en TPR_MAECFI (Referencia) : " & r_str_NumRef, vbInformation, modgen_g_str_NomPlt
   End If
         
   r_obj_Excel.Visible = True
End Sub

Private Sub fs_Obtiene_Cabecera()
Dim r_int_ConFil As Integer
Dim r_int_ConCol As Integer
   
   grd_Listad.Redraw = False
   Call gs_LimpiaGrid(grd_Listad)
   
   'Primera Linea
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Row = 0:   grd_Listad.Text = ""
   grd_Listad.Col = 0:   grd_Listad.Text = "Tipo"
   grd_Listad.Col = 1:   grd_Listad.Text = "N° FMV"
   grd_Listad.Col = 2:   grd_Listad.Text = "Número"
   grd_Listad.Col = 3:   grd_Listad.Text = "F. Emisión"
   grd_Listad.Col = 4:   grd_Listad.Text = "Fec. Vcto."
   grd_Listad.Col = 5:   grd_Listad.Text = "Valor       "
   grd_Listad.Col = 6:   grd_Listad.Text = "Garantizado "
   grd_Listad.Col = 7:   grd_Listad.Text = "COMISIONES                            "
   grd_Listad.Col = 8:   grd_Listad.Text = "COMISIONES                            "
   grd_Listad.Col = 9:   grd_Listad.Text = "COMISIONES                            "
   grd_Listad.Col = 10:  grd_Listad.Text = "FONDOS RECIBIDOS                      "
   grd_Listad.Col = 11:  grd_Listad.Text = "FONDOS RECIBIDOS                      "
   grd_Listad.Col = 12:  grd_Listad.Text = "FONDOS RECIBIDOS                      "
   grd_Listad.Col = 13:  grd_Listad.Text = "DESEMBOLSOS - ET                      "
   grd_Listad.Col = 14:  grd_Listad.Text = "DESEMBOLSOS - ET                      "
   grd_Listad.Col = 15:  grd_Listad.Text = "DESEMBOLSOS - ET                      "
   grd_Listad.Col = 16:  grd_Listad.Text = "ESTADO"

   'Segunda linea
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0:   grd_Listad.Text = "Tipo"
   grd_Listad.Col = 1:   grd_Listad.Text = "N° FMV"
   grd_Listad.Col = 2:   grd_Listad.Text = "Número"
   grd_Listad.Col = 3:   grd_Listad.Text = "F. Emisión"
   grd_Listad.Col = 4:   grd_Listad.Text = "Fec. Vcto."
   grd_Listad.Col = 5:   grd_Listad.Text = "Valor       "
   grd_Listad.Col = 6:   grd_Listad.Text = "Garantizado "
   grd_Listad.Col = 7:   grd_Listad.Text = "Importe      "
   grd_Listad.Col = 8:   grd_Listad.Text = "Pagado       "
   grd_Listad.Col = 9:   grd_Listad.Text = "Saldo        "
   grd_Listad.Col = 10:  grd_Listad.Text = "Importe      "
   grd_Listad.Col = 11:  grd_Listad.Text = "Pagado       "
   grd_Listad.Col = 12:  grd_Listad.Text = "Saldo        "
   grd_Listad.Col = 13:  grd_Listad.Text = "Importe      "
   grd_Listad.Col = 14:  grd_Listad.Text = "Recibido     "
   grd_Listad.Col = 15:  grd_Listad.Text = "Saldo        "
   grd_Listad.Col = 16:  grd_Listad.Text = "ESTADO"

   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1

   With grd_Listad
      .MergeCells = flexMergeFree
      .MergeCol(0) = True
      .MergeCol(1) = True
      .MergeCol(2) = True
      .MergeCol(3) = True
      .MergeCol(4) = True
      .MergeCol(5) = True
      .MergeCol(6) = True
      .MergeCol(16) = True
      .MergeRow(0) = True
      .FixedRows = 2
   End With

   grd_Listad.Rows = grd_Listad.Rows - 1

   For r_int_ConFil = 0 To grd_Listad.Rows - 1
      For r_int_ConCol = 0 To grd_Listad.Cols - 1
         grd_Listad.Col = r_int_ConCol
         grd_Listad.Row = r_int_ConFil
         grd_Listad.CellBackColor = &H4000&
         grd_Listad.ForeColorFixed = &HFFFFFF
      Next r_int_ConCol
   Next r_int_ConFil
   grd_Listad.Redraw = True
End Sub

Public Sub fs_Activa(ByVal p_Activa As Integer)
   ipp_FecIni.Enabled = Not p_Activa
   ipp_FecFin.Enabled = Not p_Activa
   cmd_Buscar.Enabled = p_Activa 'Not
   cmd_Limpia.Enabled = p_Activa 'Not
   cmb_EstFia.Enabled = Not p_Activa

   cmd_Gestion.Enabled = Not p_Activa
   cmd_DatBen.Enabled = Not p_Activa
   cmd_ExpExc.Enabled = Not p_Activa
   cmd_ExpLiq.Enabled = Not p_Activa
      
   cmd_Agrega.Enabled = p_Activa
   cmd_Editar.Enabled = Not p_Activa
   cmd_Borrar.Enabled = Not p_Activa
   'grd_Listad.Enabled = Not p_Activa
   
   cmd_Renova.Enabled = Not p_Activa
   cmd_Histor.Enabled = Not p_Activa
End Sub

Private Sub grd_Listad_Dir_DblClick()
   moddat_g_str_TipCre = 2
   cmd_Editar_Click
End Sub

Private Sub grd_Listad_Dir_LeaveCell()
'   moddat_g_str_TipCre = 2
'   moddat_g_int_TipPan = SSTab1.Tab
End Sub

Private Sub grd_Listad_Dir_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   moddat_g_str_TipCre = 2
End Sub

Private Sub grd_Listad_LeaveCell()
'   moddat_g_str_TipCre = 1
'   moddat_g_int_TipPan = SSTab1.Tab
End Sub

Private Sub grd_Listad_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   moddat_g_str_TipCre = 1
End Sub

Private Sub ipp_FecFin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_EstFia)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub ipp_FecIni_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecFin)
   Else
      KeyAscii = 0
   End If
End Sub

Function fs_NomMes(mes As Integer) As String
    Select Case mes
        Case 1
            fs_NomMes = "enero"
        Case 2
            fs_NomMes = "febrero"
        Case 3
            fs_NomMes = "marzo"
        Case 4
            fs_NomMes = "abril"
        Case 5
            fs_NomMes = "mayo"
        Case 6
            fs_NomMes = "junio"
        Case 7
            fs_NomMes = "julio"
        Case 8
            fs_NomMes = "agosto"
        Case 9
            fs_NomMes = "setiembre"
        Case 10
            fs_NomMes = "octubre"
        Case 11
            fs_NomMes = "noviembre"
        Case 12
            fs_NomMes = "diciembre"
    End Select
End Function

Private Sub smnu_Click(Index As Integer)
Dim r_str_NumRef        As String

    Select Case Index
        Case 0:
            If grd_Listad.TextMatrix(grd_Listad.Row, 8) = 0 Then
                MsgBox "Para imprimir Liquidación, la Comisión debe estar pagada.", vbExclamation, modgen_g_str_NomPlt
                Call gs_SetFocus(cmd_ExpLiq)
                Exit Sub
            End If
            
            'Confirmacion
            If MsgBox("¿Está seguro de imprimir los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
                Exit Sub
            End If
            
            Screen.MousePointer = 11
            Call fs_GenLiqCarFia
            Screen.MousePointer = 0
            
        Case 1:
            
            If Trim(grd_Listad.TextMatrix(grd_Listad.Row, 0)) = "AD" Then
               MsgBox "No se permite imprimir Adenda.", vbExclamation, modgen_g_str_NomPlt
               Exit Sub
            End If
            'Confirmacion
            If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
              Exit Sub
            End If
            
            r_str_NumRef = grd_Listad.TextMatrix(grd_Listad.Row, 2)
            
            g_str_Parame = ""
            g_str_Parame = g_str_Parame & "  SELECT MAECFI_NUMREF, MAEPRV_RAZSOC, MAECFI_CODPRD, MAECFI_CODSUB, MAECFI_CODMOD, MAECFI_EMIFIA, MAECFI_VTOFIA, MAECFI_CODPRY, MAECFI_CODETE, MAECFI_SITUAC, MAECFI_NOMPRY, "
            g_str_Parame = g_str_Parame & "         MAECFI_GARFIA, MAECFI_TASFIA, MAECFI_NUMREN, MAECFI_REFANT, MAECFI_TIPREC, MAECFI_MONFIA, " 'CASE WHEN MAECFI_TIPREC = 1 THEN 'BFH' ELSE 'AHORRO' END AS RECURSO, NVL(DATBEN_CODIGO,0) AS CANT_BEN "
            g_str_Parame = g_str_Parame & "         CASE WHEN MAECFI_TIPREC = 1 THEN 'BFH' "
            g_str_Parame = g_str_Parame & "              WHEN MAECFI_TIPREC = 2 THEN 'AHORRO' "
            g_str_Parame = g_str_Parame & "          END AS RECURSO, "
            g_str_Parame = g_str_Parame & "          (SELECT COUNT(*) "
            g_str_Parame = g_str_Parame & "             FROM TPR_DATBEN"
            g_str_Parame = g_str_Parame & "            WHERE DATBEN_NUMREF = MAECFI_NUMREF "
            g_str_Parame = g_str_Parame & "              AND DATBEN_TIPREC = MAECFI_TIPREC "
            g_str_Parame = g_str_Parame & "           ) AS CANT_BEN"
            g_str_Parame = g_str_Parame & "    FROM TPR_MAECFI "
            g_str_Parame = g_str_Parame & "          INNER JOIN CNTBL_MAEPRV ON MAECFI_TIPDOC = MAEPRV_TIPDOC AND MAECFI_NUMDOC = MAEPRV_NUMDOC"
            'g_str_Parame = g_str_Parame & "           LEFT JOIN TPR_DATBEN ON DATBEN_NUMREF = MAECFI_NUMREF AND DATBEN_TIPREC = MAECFI_TIPREC "
            g_str_Parame = g_str_Parame & "   WHERE MAECFI_NUMREF = '" & Replace(r_str_NumRef, "-", "") & "' "
            g_str_Parame = g_str_Parame & "     AND MAECFI_SITUAC = 1 "
              
            If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
                Exit Sub
            End If
              
            If g_rst_Princi.BOF And g_rst_Princi.EOF Then
               g_rst_Princi.Close
               Set g_rst_Princi = Nothing
               Exit Sub
            End If
            
            Screen.MousePointer = 11
            
            If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
                'VIGENTE SIN RENOVACION
                If g_rst_Princi!MAECFI_NUMREN = 0 Then
                    If g_rst_Princi!MAECFI_CODPRD = "026" And (g_rst_Princi!MAECFI_CODSUB = "002" Or g_rst_Princi!MAECFI_CODSUB = "003") And (g_rst_Princi!MAECFI_CODMOD = "004" Or g_rst_Princi!MAECFI_CODMOD = "006" Or g_rst_Princi!MAECFI_CODMOD = "007") Then
                        Call fs_GenWrd_ConSitPro_MejViv(g_rst_Princi)
                    ElseIf g_rst_Princi!MAECFI_CODPRD = "026" And g_rst_Princi!MAECFI_CODSUB = "001" And (g_rst_Princi!MAECFI_CODMOD = "004") Then 'Or g_rst_Princi!MAECFI_CODMOD = "005"
                        Call fs_GenWrd_AdqViv(g_rst_Princi)
                    ElseIf g_rst_Princi!MAECFI_CODPRD = "026" And (g_rst_Princi!MAECFI_CODSUB = "001" Or g_rst_Princi!MAECFI_CODSUB = "002" Or g_rst_Princi!MAECFI_CODSUB = "003") And (g_rst_Princi!MAECFI_CODMOD = "008") Then
                        Call fs_GenWrd_CSO(g_rst_Princi)
                    ElseIf g_rst_Princi!MAECFI_CODPRD = "027" And g_rst_Princi!MAECFI_CODSUB = "004" And (g_rst_Princi!MAECFI_CODMOD = "008") Then
                        Call fs_GenWrd_CSO(g_rst_Princi)
                    ElseIf g_rst_Princi!MAECFI_CODPRD = "027" And g_rst_Princi!MAECFI_CODSUB = "004" And (g_rst_Princi!MAECFI_CODMOD = "002") Then
                        Call fs_GenWrd_Reforzamiento_Estructural(g_rst_Princi)
                    End If
                'RENOVADO
                Else
                    Call fs_GenWrd_Renovacion(g_rst_Princi)
                End If
            End If
            Screen.MousePointer = 0
    End Select
End Sub

Private Function fs_NroEnLetras(numero As String) As String
    Dim b, paso As Integer
    Dim expresion, entero, deci, flag As String
        
    flag = "N"
    For paso = 1 To Len(numero)
        If Mid(numero, paso, 1) = "." Then
            flag = "S"
        Else
            If flag = "N" Then
                entero = entero + Mid(numero, paso, 1)   'Extae la parte entera del numero
            Else
                deci = deci + Mid(numero, paso, 1)       'Extrae la parte decimal del numero
            End If
        End If
    Next paso
    
    If Len(deci) = 1 Then
        deci = deci & "0"
    End If
    
    flag = "N"
    If Val(numero) >= -999999999 And Val(numero) <= 999999999 Then 'si el numero esta dentro de 0 a 999.999.999
        For paso = Len(entero) To 1 Step -1
            b = Len(entero) - (paso - 1)
            Select Case paso
            Case 3, 6, 9
                Select Case Mid(entero, b, 1)
                    Case "1"
                        If Mid(entero, b + 1, 1) = "0" And Mid(entero, b + 2, 1) = "0" Then
                            expresion = expresion & "Cien "
                        Else
                            expresion = expresion & "Ciento "
                        End If
                    Case "2"
                        expresion = expresion & "Doscientos "
                    Case "3"
                        expresion = expresion & "Trescientos "
                    Case "4"
                        expresion = expresion & "Cuatrocientos "
                    Case "5"
                        expresion = expresion & "Quinientos "
                    Case "6"
                        expresion = expresion & "Seiscientos "
                    Case "7"
                        expresion = expresion & "Setecientos "
                    Case "8"
                        expresion = expresion & "Ochocientos "
                    Case "9"
                        expresion = expresion & "Novecientos "
                End Select
                
            Case 2, 5, 8
                Select Case Mid(entero, b, 1)
   
                    Case "1"
                        If Mid(entero, b + 1, 1) = "0" Then
                            flag = "S"
                            expresion = expresion & "Diez "
                        End If
                        If Mid(entero, b + 1, 1) = "1" Then
                            flag = "S"
                            expresion = expresion & "Once "
                        End If
                        If Mid(entero, b + 1, 1) = "2" Then
                            flag = "S"
                            expresion = expresion & "Doce "
                        End If
                        If Mid(entero, b + 1, 1) = "3" Then
                            flag = "S"
                            expresion = expresion & "Trece "
                        End If
                        If Mid(entero, b + 1, 1) = "4" Then
                            flag = "S"
                            expresion = expresion & "Catorce "
                        End If
                        If Mid(entero, b + 1, 1) = "5" Then
                            flag = "S"
                            expresion = expresion & "Quince "
                        End If
                        If Mid(entero, b + 1, 1) > "5" Then
                            flag = "N"
                            expresion = expresion & "Dieci"
                        End If
                
                    Case "2"
                        If Mid(entero, b + 1, 1) = "0" Then
                            expresion = expresion & "Veinte "
                            flag = "S"
                        Else
                            expresion = expresion & "Veinti"
                            flag = "N"
                        End If
                    
                    Case "3"
                        If Mid(entero, b + 1, 1) = "0" Then
                            expresion = expresion & "Treinta "
                            flag = "S"
                        Else
                            expresion = expresion & "Treinta y "
                            flag = "N"
                        End If
                
                    Case "4"
                        If Mid(entero, b + 1, 1) = "0" Then
                            expresion = expresion & "Cuarenta "
                            flag = "S"
                        Else
                            expresion = expresion & "Cuarenta y "
                            flag = "N"
                        End If
                
                    Case "5"
                        If Mid(entero, b + 1, 1) = "0" Then
                            expresion = expresion & "Cincuenta "
                            flag = "S"
                        Else
                            expresion = expresion & "Cincuenta y "
                            flag = "N"
                        End If
                
                    Case "6"
                        If Mid(entero, b + 1, 1) = "0" Then
                            expresion = expresion & "Sesenta "
                            flag = "S"
                        Else
                            expresion = expresion & "Sesenta y "
                            flag = "N"
                        End If
                
                    Case "7"
                        If Mid(entero, b + 1, 1) = "0" Then
                            expresion = expresion & "Setenta "
                            flag = "S"
                        Else
                            expresion = expresion & "Setenta y "
                            flag = "N"
                        End If
                
                    Case "8"
                        If Mid(entero, b + 1, 1) = "0" Then
                            expresion = expresion & "Ochenta "
                            flag = "S"
                        Else
                            expresion = expresion & "Ochenta y "
                            flag = "N"
                        End If
                
                    Case "9"
                        If Mid(entero, b + 1, 1) = "0" Then
                            expresion = expresion & "Noventa "
                            flag = "S"
                        Else
                            expresion = expresion & "Noventa y "
                            flag = "N"
                        End If
                End Select
                
            Case 1, 4, 7
                Select Case Mid(entero, b, 1)
                     
                    Case "1"
                        If flag = "N" Then
                            If paso = 1 Then
                                expresion = expresion & "uno "
                            Else
                                expresion = expresion & "Un "
                            End If
                        End If
                    Case "2"
                        If flag = "N" Then
                            expresion = expresion & "dos "
                        End If
                    Case "3"
                        If flag = "N" Then
                            expresion = expresion & "tres "
                        End If
                    Case "4"
                        If flag = "N" Then
                            expresion = expresion & "cuatro "
                        End If
                    Case "5"
                        If flag = "N" Then
                            expresion = expresion & "cinco "
                        End If
                    Case "6"
                        If flag = "N" Then
                            expresion = expresion & "seis "
                        End If
                    Case "7"
                        If flag = "N" Then
                            expresion = expresion & "siete "
                        End If
                    Case "8"
                        If flag = "N" Then
                            expresion = expresion & "ocho "
                        Else
                            expresion = expresion & "ocho "
                        End If
                    Case "9"
                        If flag = "N" Then
                            expresion = expresion & "nueve "
                        End If
                End Select
            End Select
            If paso = 4 Then
                If Mid(entero, 6, 1) <> "0" Or Mid(entero, 5, 1) <> "0" Or Mid(entero, 4, 1) <> "0" Or _
                  (Mid(entero, 6, 1) = "0" And Mid(entero, 5, 1) = "0" And Mid(entero, 4, 1) = "0" And _
                   Len(entero) <= 6) Then
                    expresion = expresion & "Mil "
                End If
            End If
            If paso = 7 Then
                If Len(entero) = 7 And Mid(entero, 1, 1) = "1" Then
                    expresion = expresion & "Millón "
                Else
                    expresion = expresion & "Millones "
                End If
            End If
        Next paso
        
        If deci <> "" Then
            If Mid(entero, 1, 1) = "-" Then 'si el numero es negativo
                fs_NroEnLetras = "menos " & expresion '& "con " & deci & "/100"
            Else
                fs_NroEnLetras = expresion '& "con " & deci & "/100"
            End If
        Else
            If Mid(entero, 1, 1) = "-" Then 'si el numero es negativo
                fs_NroEnLetras = "menos " & expresion
            Else
                fs_NroEnLetras = expresion '& "con 00/100"
            End If
        End If
    Else 'si el numero a convertir esta fuera del rango superior e inferior
        fs_NroEnLetras = ""
    End If
End Function

Private Sub SSTab1_Click(PreviousTab As Integer)
   moddat_g_int_TipPan = PreviousTab
End Sub

