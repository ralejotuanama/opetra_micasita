VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_Caj_CiePag_03 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8685
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11175
   Icon            =   "OpeTra_frm_820.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8685
   ScaleWidth      =   11175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   8655
      Left            =   30
      TabIndex        =   5
      Top             =   30
      Width           =   11145
      _Version        =   65536
      _ExtentX        =   19659
      _ExtentY        =   15266
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
         Height          =   645
         Left            =   30
         TabIndex        =   13
         Top             =   750
         Width           =   11070
         _Version        =   65536
         _ExtentX        =   19526
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
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10470
            Picture         =   "OpeTra_frm_820.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_BusCli 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_820.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Buscar Cliente"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   600
            Picture         =   "OpeTra_frm_820.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Limpiar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Borrar 
            Height          =   585
            Left            =   1170
            Picture         =   "OpeTra_frm_820.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Borrar Registro"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Valida 
            Height          =   585
            Left            =   1740
            Picture         =   "OpeTra_frm_820.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Validar Solicitud"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   2310
            Picture         =   "OpeTra_frm_820.frx":1076
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Grabar 
            Enabled         =   0   'False
            Height          =   585
            Left            =   2880
            Picture         =   "OpeTra_frm_820.frx":1380
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   14
         Top             =   30
         Width           =   11070
         _Version        =   65536
         _ExtentX        =   19526
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
            Height          =   255
            Left            =   720
            TabIndex        =   15
            Top             =   120
            Width           =   4995
            _Version        =   65536
            _ExtentX        =   8819
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "Operaciones Financieras"
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
         Begin Threed.SSPanel SSPanel4 
            Height          =   255
            Left            =   720
            TabIndex        =   16
            Top             =   390
            Width           =   6435
            _Version        =   65536
            _ExtentX        =   11351
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "Créditos Hipotecarios - Pago Proveedores de Gastos de Cierre"
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
            Left            =   120
            Picture         =   "OpeTra_frm_820.frx":17C2
            Top             =   120
            Width           =   480
         End
      End
      Begin Threed.SSPanel pnl_SolEva 
         Height          =   5265
         Left            =   30
         TabIndex        =   17
         Top             =   3330
         Width           =   11070
         _Version        =   65536
         _ExtentX        =   19526
         _ExtentY        =   9287
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
         Begin Threed.SSPanel pnl_Tit_DocIde 
            Height          =   285
            Left            =   60
            TabIndex        =   18
            Top             =   60
            Width           =   1695
            _Version        =   65536
            _ExtentX        =   2990
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "DNI"
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
         Begin Threed.SSPanel pnl_Tit_NomCli 
            Height          =   285
            Left            =   1695
            TabIndex        =   19
            Top             =   60
            Width           =   4125
            _Version        =   65536
            _ExtentX        =   7276
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "CLIENTE"
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
         Begin Threed.SSPanel pnl_Tit_Saldo 
            Height          =   285
            Left            =   5820
            TabIndex        =   20
            Top             =   60
            Width           =   1605
            _Version        =   65536
            _ExtentX        =   2831
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "SALDO"
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
         Begin Threed.SSPanel pnl_Tit_Pago 
            Height          =   285
            Left            =   7425
            TabIndex        =   21
            Top             =   60
            Width           =   1605
            _Version        =   65536
            _ExtentX        =   2831
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "PAGO"
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
            Begin VB.CheckBox chkSeleccionar 
               BackColor       =   &H00004000&
               Height          =   255
               Left            =   1320
               TabIndex        =   22
               Top             =   10
               Width           =   255
            End
         End
         Begin Threed.SSPanel pnl_TotSal 
            Height          =   285
            Left            =   5820
            TabIndex        =   23
            Top             =   4800
            Width           =   1605
            _Version        =   65536
            _ExtentX        =   2831
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
         Begin EditLib.fpDoubleSingle ipp_MtoPag 
            Height          =   315
            Left            =   120
            TabIndex        =   24
            Top             =   4800
            Visible         =   0   'False
            Width           =   1815
            _Version        =   196608
            _ExtentX        =   3201
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
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   4305
            Left            =   0
            TabIndex        =   4
            Top             =   360
            Width           =   11025
            _ExtentX        =   19447
            _ExtentY        =   7594
            _Version        =   393216
            Rows            =   6
            Cols            =   6
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   2
            ScrollBars      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSPanel pnl_Tit_Ajuste 
            Height          =   285
            Left            =   9030
            TabIndex        =   25
            Top             =   60
            Width           =   1605
            _Version        =   65536
            _ExtentX        =   2831
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "AJUSTE"
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
         Begin Threed.SSPanel pnl_TotPag 
            Height          =   285
            Left            =   7410
            TabIndex        =   26
            Top             =   4800
            Width           =   1605
            _Version        =   65536
            _ExtentX        =   2831
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
         Begin Threed.SSPanel pnl_TotAju 
            Height          =   285
            Left            =   9000
            TabIndex        =   27
            Top             =   4800
            Width           =   1605
            _Version        =   65536
            _ExtentX        =   2831
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
         Begin VB.Label Label4 
            Caption         =   "Total ==>"
            Height          =   285
            Left            =   4740
            TabIndex        =   28
            Top             =   4800
            Width           =   1215
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   1845
         Left            =   30
         TabIndex        =   29
         Top             =   1440
         Width           =   11070
         _Version        =   65536
         _ExtentX        =   19526
         _ExtentY        =   3254
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
         Begin VB.ComboBox cmb_GasAdm 
            Height          =   315
            Left            =   2460
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   5415
         End
         Begin VB.TextBox txt_NumDoc 
            Height          =   315
            Left            =   2460
            MaxLength       =   25
            TabIndex        =   3
            Top             =   1050
            Width           =   4875
         End
         Begin VB.ComboBox cmb_TipDoc 
            Height          =   315
            Left            =   2460
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   720
            Width           =   4875
         End
         Begin EditLib.fpDateTime ipp_FecPag 
            Height          =   315
            Left            =   2460
            TabIndex        =   1
            Top             =   405
            Width           =   1335
            _Version        =   196608
            _ExtentX        =   2355
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
            Text            =   "24/01/2014"
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
            Left            =   2460
            TabIndex        =   30
            Top             =   1410
            Width           =   5415
            _Version        =   65536
            _ExtentX        =   9551
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
         Begin VB.Label Label8 
            Caption         =   "Operación"
            Height          =   285
            Left            =   60
            TabIndex        =   35
            Top             =   75
            Width           =   915
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha:"
            Height          =   285
            Left            =   60
            TabIndex        =   34
            Top             =   415
            Width           =   915
         End
         Begin VB.Label Label2 
            Caption         =   "Nro. Documento Proveedor:"
            Height          =   285
            Left            =   60
            TabIndex        =   33
            Top             =   1065
            Width           =   1995
         End
         Begin VB.Label Label3 
            Caption         =   "Razón Social:"
            Height          =   255
            Left            =   60
            TabIndex        =   32
            Top             =   1410
            Width           =   1155
         End
         Begin VB.Label lbl_TipDoc 
            Caption         =   "Tipo Documento Proveedor:"
            Height          =   255
            Left            =   60
            TabIndex        =   31
            Top             =   750
            Width           =   2265
         End
      End
   End
End
Attribute VB_Name = "frm_Caj_CiePag_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_int_CodGas     As Integer
Dim l_int_NumFila    As Integer
Dim l_str_FilDel     As String

Private Sub chkSeleccionar_Click()
Dim r_str_MtoVal  As String
Dim r_int_Contad  As Integer

   If grd_Listad.Rows > 0 Then
      If chkSeleccionar.Value = 1 Then
         r_str_MtoVal = InputBox("Ingrese Monto a Pagar:", Me.cmb_GasAdm.Text)
         If CStr(r_str_MtoVal) <> "" Then
            For r_int_Contad = 0 To grd_Listad.Rows - 1
               grd_Listad.TextMatrix(r_int_Contad, 3) = Format(Val(r_str_MtoVal), "###,###,##0.00")
            Next r_int_Contad
         End If
      Else
          For r_int_Contad = 0 To grd_Listad.Rows - 1
              grd_Listad.TextMatrix(r_int_Contad, 3) = Format(Val(0), "###,###,##0.00")
              grd_Listad.TextMatrix(r_int_Contad, 4) = Format(Val(0), "###,###,##0.00")
          Next r_int_Contad
      End If
   End If
   Call fs_Total_Pago
   Call fs_Total_Ajuste
End Sub

Private Sub cmb_GasAdm_Click()
   If cmb_GasAdm.ListIndex <> -1 Then
      l_int_CodGas = cmb_GasAdm.ItemData(cmb_GasAdm.ListIndex)
   End If
   Call gs_SetFocus(ipp_FecPag)
End Sub

Private Sub cmb_GasAdm_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(ipp_FecPag)
    End If
End Sub

Private Sub cmb_TipDoc_Click()
   Call gs_SetFocus(txt_NumDoc)
End Sub

Private Sub cmb_TipDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NumDoc)
   End If
End Sub

Private Sub cmd_Borrar_Click()
   
   If l_str_FilDel <> "" Then
      
      'Confirma
      If MsgBox("¿Está seguro de Eliminar la solicitud seleccionada?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
         Exit Sub
      End If
   
      If grd_Listad.Rows = 1 Then
         Call gs_LimpiaGrid(grd_Listad)
         Call fs_Limpia
         fs_Activa (False)
      Else
         If Val(l_str_FilDel) = grd_Listad.Rows Then
            l_str_FilDel = Empty
            Call fs_Total_Saldo
            Call fs_Total_Pago
            Call fs_Total_Ajuste
            Exit Sub
         End If
         If Val(l_str_FilDel) = grd_Listad.Row And l_str_FilDel = 0 Then
            grd_Listad.RemoveItem (Val(l_str_FilDel))
            l_str_FilDel = Empty
            Call fs_Total_Saldo
            Call fs_Total_Pago
            Call fs_Total_Ajuste
            Exit Sub
         Else
            grd_Listad.RemoveItem (Val(l_str_FilDel))
         End If
      End If
      l_str_FilDel = Empty
      Call fs_Total_Saldo
      Call fs_Total_Pago
      Call fs_Total_Ajuste
   Else
       MsgBox "Debe seleccionar la solicitud a eliminar.", vbExclamation, modgen_g_str_NomPlt
   End If
End Sub

Private Sub cmd_BusCli_Click()
   If Me.cmb_GasAdm.ListIndex = -1 Then
      MsgBox "Debe seleccionar Gasto de Cierre ", vbExclamation, modgen_g_con_OpeTra
      Call gs_SetFocus(cmb_GasAdm)
      Exit Sub
   End If
   
   moddat_g_int_FlgCre = 4
   frm_Caj_SolHip_01.Show 1
End Sub

Private Sub cmd_ExpExc_Click()
Dim r_int_Contad     As Integer

   If Me.cmb_GasAdm.ListIndex = -1 Then
       MsgBox "Debe seleccionar Operación", modgen_g_str_NomPlt
       Call gs_SetFocus(cmb_GasAdm)
   End If
   If Len(Trim(txt_NumDoc.Text)) = 0 Then
       MsgBox "Debe ingresar el número de documento del proveedor", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(txt_NumDoc)
       Exit Sub
   End If
   If Len(Trim(Me.pnl_RazSoc.Caption)) = 0 Then
       MsgBox "Debe ingresar un documento válido", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(cmb_GasAdm)
       Exit Sub
   End If
   If CDate(ipp_FecPag.Text) > date Then
      MsgBox "Debe ingresar una fecha correcta.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecPag)
      Exit Sub
   End If
   If CDbl(pnl_TotPag.Caption) = 0 Then
      MsgBox "Existen importes ceros.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(chkSeleccionar)
      Exit Sub
   End If
   If fs_ValidarCeldaVacias = True Then
      MsgBox "No deben existir Celdas Vacías.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(grd_Listad)
      Exit Sub
   End If
   
   r_int_Contad = fs_ValidarCeldaCero
   If grd_Listad.Row >= 0 And (r_int_Contad > 0) Then
      MsgBox "El Monto Pagado no puede ser cero.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(grd_Listad)
      Exit Sub
   End If
   If fs_ValidarSaldoMenor = True Then
      MsgBox "Saldo Menor al Monto Pagado, es necesario Ajuste.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_Valida)
      Exit Sub
   End If
   If grd_Listad.Rows = 0 Then
      MsgBox "No existen datos a exportar.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_ExpExc)
      Exit Sub
   End If
   
   'Confirmacion
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
      
   Call fs_GenExc
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
        .Cells(1, 2) = "REPORTE DE PAGO PROVEEDORES - GASTOS DE CIERRE"
        .Range(.Cells(1, 2), .Cells(1, 6)).Merge
        .Range(.Cells(1, 2), .Cells(1, 6)).Font.Bold = True
        .Range(.Cells(1, 2), .Cells(1, 6)).HorizontalAlignment = xlHAlignCenter
        
        .Cells(3, 2) = "OPERACIÓN"
        .Cells(3, 3) = Trim(cmb_GasAdm.Text)
        .Cells(4, 2) = "FECHA"
        .Cells(4, 3) = "'" & Format(CDate(ipp_FecPag.Text), "dd/mm/yyyy")
        .Cells(5, 2) = "NRO. DOCUMENTO"
        .Cells(5, 3) = "'" & Trim(txt_NumDoc.Text)
        .Cells(6, 2) = "RAZÓN SOCIAL"
        .Cells(6, 3) = Trim(pnl_RazSoc.Caption)
        .Range(.Cells(3, 2), .Cells(6, 2)).Font.Bold = True
        
        .Cells(r_int_NroFil, 2) = "DNI"
        .Cells(r_int_NroFil, 3) = "CLIENTE"
        .Cells(r_int_NroFil, 4) = "SALDO"
        .Cells(r_int_NroFil, 5) = "PAGO"
        .Cells(r_int_NroFil, 6) = "AJUSTE"
                
        .Range(.Cells(r_int_NroFil, 2), .Cells(r_int_NroFil, 6)).Interior.Color = RGB(146, 208, 80)
        .Range(.Cells(r_int_NroFil, 2), .Cells(r_int_NroFil, 6)).Font.Bold = True
        .Range(.Cells(r_int_NroFil, 2), .Cells(r_int_NroFil, 6)).HorizontalAlignment = xlHAlignCenter
        
        .Columns("A").ColumnWidth = 1
        .Columns("B").ColumnWidth = 17.5
        .Columns("C").ColumnWidth = 45
        .Columns("D").ColumnWidth = 13.5
        .Columns("D").HorizontalAlignment = xlHAlignRight
        .Columns("D").NumberFormat = "###,###,###,##0.00"
        .Columns("E").ColumnWidth = 13.5
        .Columns("E").NumberFormat = "###,###,###,##0.00"
        .Columns("E").HorizontalAlignment = xlHAlignRight
        .Columns("F").ColumnWidth = 13.5
        .Columns("F").NumberFormat = "###,###,###,##0.00"
        .Columns("F").HorizontalAlignment = xlHAlignRight
            
        .Range(.Cells(1, 1), .Cells(99, 99)).Font.Name = "Calibri"
        .Range(.Cells(1, 1), .Cells(99, 99)).Font.Size = 11
         
        r_int_NroFil = r_int_NroFil + 1
        For r_int_NoFlLi = 0 To grd_Listad.Rows - 1
            .Cells(r_int_NroFil, 2) = grd_Listad.TextMatrix(r_int_NoFlLi, 0)
            .Cells(r_int_NroFil, 3) = grd_Listad.TextMatrix(r_int_NoFlLi, 1)
            .Cells(r_int_NroFil, 4) = grd_Listad.TextMatrix(r_int_NoFlLi, 2)
            .Cells(r_int_NroFil, 5) = grd_Listad.TextMatrix(r_int_NoFlLi, 3)
            .Cells(r_int_NroFil, 6) = grd_Listad.TextMatrix(r_int_NoFlLi, 4)
            
            r_int_NroFil = r_int_NroFil + 1
        Next r_int_NoFlLi
        
        .Range(.Cells(r_int_NroFil, 4), .Cells(r_int_NroFil, 6)).FormulaR1C1 = "=SUM(R[-" & r_int_NroFil - 9 & "]C:R[-1]C)"
        .Range(.Cells(r_int_NroFil, 4), .Cells(r_int_NroFil, 6)).Font.Bold = True
   End With
   
   r_obj_Excel.Visible = True
End Sub

Private Sub fs_InsertaCompensacion(ByVal p_CodOpe As String, ByVal p_DatCta As String)
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
           
   Do While moddat_g_int_FlgGOK = False
      Screen.MousePointer = 11
      
      g_str_Parame = "USP_CNTBL_COMAUT ("
      g_str_Parame = g_str_Parame & "'" & p_CodOpe & "', "
      g_str_Parame = g_str_Parame & Format(CDate(ipp_FecPag.Value), "yyyymmdd") & ", "
      g_str_Parame = g_str_Parame & cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) & ", "
      g_str_Parame = g_str_Parame & "'" & txt_NumDoc.Text & "', "
      g_str_Parame = g_str_Parame & 1 & ", "                                           'Tipo de Moneda
      g_str_Parame = g_str_Parame & CDbl(pnl_TotPag.Caption) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_CodIte & "', "                  'Código del Banco - Proveedor
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_DesIte & "', "                  'Cuenta Corriente - Proveedor
      g_str_Parame = g_str_Parame & "'251419010109', "
      g_str_Parame = g_str_Parame & "'" & p_DatCta & "', "
      
      g_str_Parame = g_str_Parame & "'PAGO MASIVO', "
      g_str_Parame = g_str_Parame & "1, "
      
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "')"
           
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
End Sub

Private Function fs_ValidarCeldaVacias() As Boolean
Dim r_int_ContCol    As Integer
Dim r_int_ContFil   As Integer

   fs_ValidarCeldaVacias = False
   With grd_Listad
      For r_int_ContCol = 1 To .Cols - 1
         .Col = r_int_ContCol
         For r_int_ContFil = 0 To .Rows - 1
            .Row = r_int_ContFil
            If .Text = "" Then
               fs_ValidarCeldaVacias = True
            End If
         Next r_int_ContFil
      Next r_int_ContCol
   End With
End Function

Private Function fs_ValidarSaldoMenor() As Boolean
Dim r_int_ContFil    As Integer
Dim r_dbl_MtoSal     As Double
Dim r_dbl_MtoPag     As Double
Dim r_dbl_MtoAju     As Double

   fs_ValidarSaldoMenor = False
   With grd_Listad
      For r_int_ContFil = 0 To .Rows - 1
         r_dbl_MtoSal = .TextMatrix(r_int_ContFil, 2)
         r_dbl_MtoPag = .TextMatrix(r_int_ContFil, 3)
         r_dbl_MtoAju = .TextMatrix(r_int_ContFil, 4)
         
         If CDbl(r_dbl_MtoSal + r_dbl_MtoAju) < CDbl(r_dbl_MtoPag) Then
            If CDbl(r_dbl_MtoAju) = r_dbl_MtoPag - r_dbl_MtoSal Then
               fs_ValidarSaldoMenor = False
            Else
               fs_ValidarSaldoMenor = True
            End If
         End If
      Next r_int_ContFil
   End With
End Function

Private Function fs_ValidarCeldaCero() As Integer
Dim r_int_ContCol    As Integer
Dim r_int_ContFil   As Integer

   fs_ValidarCeldaCero = 0
   With grd_Listad
      If .Row >= 0 Then
         For r_int_ContFil = 0 To .Rows - 1
            If .TextMatrix(r_int_ContFil, 3) = 0 Then
               fs_ValidarCeldaCero = fs_ValidarCeldaCero + 1
            End If
         Next r_int_ContFil
      End If
   End With
End Function

Private Function fs_ValidarOperacion() As String
Dim r_int_ContFil    As Integer
Dim r_str_DetMsj     As String
Dim r_dbl_MtoSal     As Double
Dim r_dbl_MtoPag     As Double
Dim r_dbl_MtoAju     As Double

   fs_ValidarOperacion = ""
   With grd_Listad
      For r_int_ContFil = 0 To .Rows - 1
        'Verificando que el Gasto Administrativo seleccionado no se haya pagado
         'g_str_Parame = ""
         'g_str_Parame = g_str_Parame & "SELECT GASADM_NUMSOL "
         'g_str_Parame = g_str_Parame & "  FROM TRA_GASADM "
         'g_str_Parame = g_str_Parame & " WHERE GASADM_NUMSOL = '" & .TextMatrix(r_int_ContFil, 5) & "' "
         'g_str_Parame = g_str_Parame & "   AND GASADM_CODGAS = '" & l_int_CodGas & "' "
         'g_str_Parame = g_str_Parame & "   AND GASADM_IMPORT > GASADM_MTOPAGPRV"
         
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & " SELECT GASADM_NUMSOL, SUM(NVL(GASADM_IMPORT,0)) AS GASADM_IMPORT, SUM(NVL(GASADM_MTOPAGPRV,0)) AS GASADM_MTOPAGPRV  "
         g_str_Parame = g_str_Parame & "   FROM TRA_GASADM  "
         g_str_Parame = g_str_Parame & "  WHERE GASADM_NUMSOL = '" & .TextMatrix(r_int_ContFil, 5) & "' "
         g_str_Parame = g_str_Parame & "  GROUP BY GASADM_NUMSOL  "
         g_str_Parame = g_str_Parame & " HAVING Sum(NVL(GASADM_IMPORT, 0)) > Sum(NVL(GASADM_MTOPAGPRV, 0))  "
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
            r_str_DetMsj = " - " & Mid(.TextMatrix(r_int_ContFil, 0), 3) & " " & Trim(.TextMatrix(r_int_ContFil, 1))
            fs_ValidarOperacion = fs_ValidarOperacion & " " & r_str_DetMsj
         End If
      
         If (Not (g_rst_Genera.BOF And g_rst_Genera.EOF)) Then
            fs_ValidarOperacion = r_str_DetMsj
         Else
            r_str_DetMsj = " - " & Mid(.TextMatrix(r_int_ContFil, 0), 3) & " " & Trim(.TextMatrix(r_int_ContFil, 1))
            fs_ValidarOperacion = Trim(fs_ValidarOperacion & vbCrLf & r_str_DetMsj)
         End If
      Next r_int_ContFil
      
      'fs_ValidarOperacion = Replace(fs_ValidarOperacion, "  ", "-")
   End With
End Function

Private Sub cmd_Grabar_Click()
Dim r_int_Contad  As Integer
Dim r_str_DetMsj  As String
Dim r_dbl_MtoAju  As Double

   '*** validaciones
   If cmb_GasAdm.ListIndex = -1 Then
       MsgBox "Debe seleccionar Operación", modgen_g_str_NomPlt
       Call gs_SetFocus(cmb_GasAdm)
   End If
   If cmb_TipDoc.ListIndex = -1 Then
       MsgBox "Debe seleccionar el Tipo de Documento del Proveedor", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(cmb_TipDoc)
       Exit Sub
   End If
   If Len(Trim(txt_NumDoc.Text)) = 0 Then
       MsgBox "Debe ingresar el Número de Documento del Proveedor", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(txt_NumDoc)
       Exit Sub
   End If
   If Len(Trim(Me.pnl_RazSoc.Caption)) = 0 Then
       MsgBox "Debe ingresar un documento válido", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(txt_NumDoc)
       Exit Sub
   End If
   If CDate(ipp_FecPag.Text) > date Then
      MsgBox "Debe ingresar una fecha correcta.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecPag)
      Exit Sub
   End If
   If CDbl(pnl_TotPag.Caption) = 0 Then
      MsgBox "Existen importes ceros.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(chkSeleccionar)
      Exit Sub
   End If
   If fs_ValidarCeldaVacias = True Then
      MsgBox "No deben existir Celdas Vacías.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(grd_Listad)
      Exit Sub
   End If
   
   r_int_Contad = fs_ValidarCeldaCero
   If grd_Listad.Row >= 0 And (r_int_Contad > 0) Then
      MsgBox "El Monto Pagado no puede ser cero.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(grd_Listad)
      Exit Sub
   End If
   If fs_ValidarSaldoMenor = True Then
      MsgBox "Saldo Menor al Monto Pagado, es necesario Ajuste.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmd_Valida)
      Exit Sub
   End If
   
   r_str_DetMsj = fs_ValidarOperacion
   
   If r_str_DetMsj <> "" Then
      MsgBox "El gasto de cierre ya está pagado, para los cliente(s): " & vbCrLf & r_str_DetMsj & "", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(grd_Listad)
      Exit Sub
   End If
   
   '*** confirmacion
   If MsgBox("¿Está seguro de registrar los pagos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   '*** grabacion
   Screen.MousePointer = 11
   Dim r_str_Resul   As String
   r_str_Resul = ""
   'Actualiza la tabla TRA_GASADM
   For r_int_Contad = 0 To grd_Listad.Rows - 1
       If grd_Listad.TextMatrix(r_int_Contad, 4) > 0 Then
         'If Not fs_PagPrv_GasAdm(grd_Listad.TextMatrix(r_int_Contad, 5), l_int_CodGas, Format(CDate(ipp_FecPag.Text), "yyyymmdd"), txt_NumDoc, _
         '                        IIf(grd_Listad.TextMatrix(r_int_Contad, 4) > 0, grd_Listad.TextMatrix(r_int_Contad, 2), grd_Listad.TextMatrix(r_int_Contad, 3))) Then
         '   Exit Sub
         'End If
         'Ingresa el ajuste
         If Not fs_PagPrv_GasAdm_Ajuste(grd_Listad.TextMatrix(r_int_Contad, 5), 24, 1, grd_Listad.TextMatrix(r_int_Contad, 4), moddat_g_str_FecSis, _
                                         moddat_g_str_FecSis, moddat_g_str_FecSis, Trim(txt_NumDoc.Text), grd_Listad.TextMatrix(r_int_Contad, 4)) Then
            Screen.MousePointer = 0
            Exit Sub
         End If
         
         If grd_Listad.TextMatrix(r_int_Contad, 4) > 0 Then
           r_dbl_MtoAju = grd_Listad.TextMatrix(r_int_Contad, 2)
         Else
           r_dbl_MtoAju = grd_Listad.TextMatrix(r_int_Contad, 3)
         End If
         If Not fs_PagPrv_GasAdm_PagMgv(grd_Listad.TextMatrix(r_int_Contad, 5), l_int_CodGas, Format(CDate(ipp_FecPag.Text), "yyyymmdd"), txt_NumDoc, _
                                       grd_Listad.TextMatrix(r_int_Contad, 2), r_dbl_MtoAju, r_str_Resul) Then
           'grd_Listad.TextMatrix(r_int_Contad, 3)
           Screen.MousePointer = 0
           Exit Sub
         End If
         
         If Trim(r_str_Resul) <> "" Then
           MsgBox "El cliente " & grd_Listad.TextMatrix(r_int_Contad, 0) & ", su proceso se detuvo al realizar la distribución por la diferencia de saldos.", vbExclamation, modgen_g_str_NomPlt
           Screen.MousePointer = 0
           Exit Sub
         End If

       Else
          'If Not fs_PagPrv_GasAdm(grd_Listad.TextMatrix(r_int_Contad, 5), l_int_CodGas, Format(CDate(ipp_FecPag.Text), "yyyymmdd"), txt_NumDoc, _
          '                        IIf(grd_Listad.TextMatrix(r_int_Contad, 4) > 0, grd_Listad.TextMatrix(r_int_Contad, 2), grd_Listad.TextMatrix(r_int_Contad, 3))) Then
          '   Exit Sub
          If Not fs_PagPrv_GasAdm_PagMgv(grd_Listad.TextMatrix(r_int_Contad, 5), l_int_CodGas, Format(CDate(ipp_FecPag.Text), "yyyymmdd"), txt_NumDoc, _
                                         grd_Listad.TextMatrix(r_int_Contad, 2), grd_Listad.TextMatrix(r_int_Contad, 3), r_str_Resul) Then
             Screen.MousePointer = 0
             Exit Sub
          End If
          If Trim(r_str_Resul) <> "" Then
             MsgBox "El cliente " & grd_Listad.TextMatrix(r_int_Contad, 0) & ", su proceso se detuvo al realizar la distribución por la diferencia de saldos.", vbExclamation, modgen_g_str_NomPlt
             Screen.MousePointer = 0
             Exit Sub
          End If
       End If
   Next r_int_Contad
   
   'Se añade el asiento según corresponda
   Call fs_GeneraAsiento(l_int_CodGas, Trim(txt_NumDoc.Text))
   Call cmd_Limpia_Click
   Call fs_Limpia
   MsgBox "Los datos se grabaron correctamente.", vbInformation, modgen_g_str_NomPlt
   Unload Me
   
   Screen.MousePointer = 11
   frm_Caj_CiePag_01.fs_Buscar
   frm_Caj_CiePag_01.fs_Habilitado (True)
   Screen.MousePointer = 0
End Sub

Private Function fs_PagPrv_GasAdm_Ajuste(ByVal p_NumSol As String, ByVal p_CodGas As Integer, ByVal p_TipMon As Integer, ByVal p_MtoPag As Double, ByVal p_FecAsig As String, ByVal p_FecPag As String, ByVal p_FecPagPrv As String, ByVal p_NumDoc As String, ByVal p_MtoPagPrv As Double) As Integer
   fs_PagPrv_GasAdm_Ajuste = False
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
      
   'Se añade Ajuste en tra_gasadm
   Do While moddat_g_int_FlgGOK = False
      Screen.MousePointer = 11
      Call moddat_gs_FecSis
      
      g_str_Parame = "USP_TRA_GASADM_AJUSTE ( "
      g_str_Parame = g_str_Parame & "'" & p_NumSol & "', "
      g_str_Parame = g_str_Parame & CStr(p_CodGas) & ", "
      g_str_Parame = g_str_Parame & CStr(p_TipMon) & ", "
      g_str_Parame = g_str_Parame & CDbl(p_MtoPag) & ", "
      g_str_Parame = g_str_Parame & Format(CDate(p_FecAsig), "yyyymmdd") & ", "
      g_str_Parame = g_str_Parame & p_FecPag & ", "
      g_str_Parame = g_str_Parame & "1, "
      g_str_Parame = g_str_Parame & Format(CDate(p_FecPagPrv), "yyyymmdd") & ", "
      g_str_Parame = g_str_Parame & "'" & p_NumDoc & "', "
      g_str_Parame = g_str_Parame & CDbl(p_MtoPagPrv) & ", "
      
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
      g_str_Parame = g_str_Parame & CStr(1) & ", "
      g_str_Parame = g_str_Parame & "1)"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If
      
      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar la grabación de los datos. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Function
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
      
      Screen.MousePointer = 0
   Loop
   
   fs_PagPrv_GasAdm_Ajuste = True
End Function

Private Function fs_PagPrv_GasAdm(ByVal p_NumSol As String, ByVal p_CodGas As Integer, ByVal p_FecPag As String, ByVal p_NumDoc As String, ByVal p_MtoPag As Double) As Integer
   fs_PagPrv_GasAdm = False
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0

   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_TRA_GASADM_PAGOPRV_2 ("
      g_str_Parame = g_str_Parame & "'" & p_NumSol & "', "
      g_str_Parame = g_str_Parame & CStr(p_CodGas) & ", "
      g_str_Parame = g_str_Parame & p_FecPag & ", "
      g_str_Parame = g_str_Parame & "'" & CStr(p_NumDoc) & "', "
      g_str_Parame = g_str_Parame & p_MtoPag & ", "

      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                              'Código Usuario
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                              'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                               'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "                              'Código Sucursal
      g_str_Parame = g_str_Parame & "1)"

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_TRA_GASADM_PAGOPRV_2. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Function
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop

   fs_PagPrv_GasAdm = True
End Function

Private Function fs_PagPrv_GasAdm_PagMgv(p_NumSol As String, p_CodGas As Integer, p_FecPag As String, p_NumDoc As String, p_Saldo As Double, p_TotPag As Double, ByRef p_Resul As String) As Boolean
   
   fs_PagPrv_GasAdm_PagMgv = False
   p_Resul = ""
  
   g_str_Parame = "USP_TRA_GASADM_PAGMSV ("
   g_str_Parame = g_str_Parame & "'" & p_NumSol & "', "        'GASADM_NUMSOL
   g_str_Parame = g_str_Parame & CStr(p_CodGas) & ", "         'GASADM_CODGAS
   g_str_Parame = g_str_Parame & p_FecPag & ", "               'GASADM_FECPAGPRV
   g_str_Parame = g_str_Parame & "'" & CStr(p_NumDoc) & "', "  'GASADM_NUMDOCPRV
   g_str_Parame = g_str_Parame & p_TotPag & ", "
   
   'Datos de Auditoria
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                              'Código Usuario
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                              'Nombre Terminal
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                               'Nombre Ejecutable
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "                              'Código Sucursal

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      fs_PagPrv_GasAdm_PagMgv = False
      p_Resul = ""
      MsgBox "No se pudo completar el procedimiento USP_TRA_GASADM_PAGMSV, en el nro solicitud:" & p_NumSol, vbExclamation, modgen_g_str_NomPlt
   Else
      fs_PagPrv_GasAdm_PagMgv = True
      If g_rst_Princi!RESUL = 0 Then
         p_Resul = p_NumSol
      End If
   End If
End Function

Private Sub fs_GeneraAsiento(ByVal p_CodGas As Integer, ByVal p_NumDoc As String)

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
Dim r_dbl_ValImp        As Double
Dim r_int_PerMes        As Integer
Dim r_int_PerAno        As Integer
Dim r_dbl_TipCam        As Double
Dim r_int_ConAux        As Integer
Dim r_str_NroCnt        As String
Dim r_str_CodOpe        As String
Dim r_dbl_TotHab        As Double
Dim r_int_Contad        As Integer
Dim r_str_NumSol        As String

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
      
   'Obteniendo Nro. de Asiento (único)
   If grd_Listad.Rows > 0 Then
      
      'Obteniendo Tipo de Cambio del día
      r_dbl_TipCam = moddat_gf_ObtieneTipCamDia(2, 2, Format(CDate(date), "yyyymmdd"), 2)
        
      'Obteniendo el Número de Asiento
      r_int_NumAsi = modprc_ff_NumAsi(r_arr_LogPro, CInt(moddat_g_str_CodAno), CInt(moddat_g_str_CodMes), r_str_Origen, r_int_NumLib)
      r_str_AsiGen = CStr(r_int_NumAsi)
      r_str_FecCon = CDate(ipp_FecPag.Text)
      r_str_FecReg = moddat_g_str_FecSis
      
      'Glosa Cabecera
      r_str_Glosa = "GASTOS DE CIERRE"
   
      'Insertar en CABECERA
      Call modprc_fs_Inserta_CabAsi(r_arr_LogPro, r_str_Origen, CInt(moddat_g_str_CodAno), CInt(moddat_g_str_CodMes), r_int_NumLib, r_int_NumAsi, Format(1, "000"), r_dbl_TipCam, r_str_TipNot, Trim(r_str_Glosa), r_str_FecCon, "1")
   End If
   
   '*****************************************************
   'GENERACION DE ASIENTOS CONTABLES DE GASTOS DE CIERRE
   '*****************************************************
   For r_int_Contad = 0 To grd_Listad.Rows - 1
   
      If grd_Listad.TextMatrix(r_int_Contad, 4) > 0 Then
         For r_int_ConAux = 1 To 2
            r_dbl_ValImp = grd_Listad.TextMatrix(r_int_Contad, 4)                   'Importe del Ajuste
            
            If r_int_ConAux = 1 Then r_str_DebHab = "D": r_str_CtaCtb = "451301290110" Else r_str_DebHab = "H": r_str_CtaCtb = "251419010114" '291807010112
            
            r_str_Glosa = Mid(grd_Listad.TextMatrix(r_int_Contad, 0), 3) & "/" & Me.txt_NumDoc.Text & "/" & "AJUSTE GASTO DE CIERRE"
             
            If (r_dbl_ValImp > 0) Then
                r_int_NumIte = r_int_NumIte + 1
                r_dbl_MtoSol = Format(r_dbl_ValImp, "###,###,##0.00")
                r_dbl_MtoDol = Format(0, "###,###,##0.00")
                
                Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, CInt(moddat_g_str_CodAno), CInt(moddat_g_str_CodMes), r_int_NumLib, r_int_NumAsi, r_int_NumIte, r_str_CtaCtb, CDate(r_str_FecCon), r_str_Glosa, r_str_DebHab, r_dbl_MtoSol, r_dbl_MtoDol, 1, CDate(r_str_FecCon))
                r_dbl_ValImp = 0
            End If
         Next r_int_ConAux
         
         r_dbl_ValImp = grd_Listad.TextMatrix(r_int_Contad, 3)
         r_str_DebHab = "D"
         r_str_CtaCtb = "251419010114" '291807010112
         
         r_str_Glosa = Mid(grd_Listad.TextMatrix(r_int_Contad, 0), 3) & "/" & p_NumDoc & "/" & "GASTO DE CIERRE"
          
         If (r_dbl_ValImp > 0) Then
             r_int_NumIte = r_int_NumIte + 1
             r_dbl_MtoSol = Format(r_dbl_ValImp, "###,###,##0.00")
             r_dbl_MtoDol = Format(0, "###,###,##0.00")
             r_dbl_TotHab = r_dbl_TotHab + r_dbl_ValImp
             
             Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, CInt(moddat_g_str_CodAno), CInt(moddat_g_str_CodMes), r_int_NumLib, r_int_NumAsi, r_int_NumIte, r_str_CtaCtb, CDate(r_str_FecCon), r_str_Glosa, r_str_DebHab, r_dbl_MtoSol, r_dbl_MtoDol, 1, CDate(r_str_FecCon))
             r_dbl_ValImp = 0
         End If
      Else
            r_dbl_ValImp = grd_Listad.TextMatrix(r_int_Contad, 3)
            r_str_DebHab = "D"
            r_str_CtaCtb = "251419010114" '291807010112
            
            r_str_Glosa = Mid(grd_Listad.TextMatrix(r_int_Contad, 0), 3) & "/" & p_NumDoc & "/" & "GASTO DE CIERRE"
             
            If (r_dbl_ValImp > 0) Then
                r_int_NumIte = r_int_NumIte + 1
                r_dbl_MtoSol = Format(r_dbl_ValImp, "###,###,##0.00")
                r_dbl_MtoDol = Format(0, "###,###,##0.00")
                r_dbl_TotHab = r_dbl_TotHab + r_dbl_ValImp
                
                Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, CInt(moddat_g_str_CodAno), CInt(moddat_g_str_CodMes), r_int_NumLib, r_int_NumAsi, r_int_NumIte, r_str_CtaCtb, CDate(r_str_FecCon), r_str_Glosa, r_str_DebHab, r_dbl_MtoSol, r_dbl_MtoDol, 1, CDate(r_str_FecCon))
                r_dbl_ValImp = 0
            End If
      End If
   Next r_int_Contad
   
   'Añade en el mismo asiento del Haber en un solo registro
   r_str_DebHab = "H"
   r_str_CtaCtb = "251419010109"
   r_str_Glosa = p_NumDoc & "/" & "GASTO DE CIERRE"
   
   If (r_dbl_TotHab > 0) Then
       r_int_NumIte = r_int_NumIte + 1
       r_dbl_MtoSol = Format(r_dbl_TotHab, "###,###,##0.00")
       r_dbl_MtoDol = Format(0, "###,###,##0.00")
       
       Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, CInt(moddat_g_str_CodAno), CInt(moddat_g_str_CodMes), r_int_NumLib, r_int_NumAsi, r_int_NumIte, r_str_CtaCtb, CDate(r_str_FecCon), r_str_Glosa, r_str_DebHab, r_dbl_MtoSol, r_dbl_MtoDol, 1, CDate(r_str_FecCon))
       r_dbl_TotHab = 0
   End If
            
   r_str_NroCnt = r_str_Origen & "/" & moddat_g_str_CodAno & "/" & Format(moddat_g_str_CodMes, "00") & "/" & r_int_NumLib & "/" & r_int_NumAsi
   r_str_CodOpe = modmip_gf_Genera_CodGen(3, 3)
   
   'Actualizando en tra_gasadm, GASADM_NROCNT = Origen/año/mes/nro_libro/nro_asiento
   'Actualizando en tra_gasadm, GASADM_CODOPE = r_str_CodOpe (Código de Operación)
   For r_int_Contad = 0 To grd_Listad.Rows - 1
   
      r_str_NumSol = grd_Listad.TextMatrix(r_int_Contad, 5)
      
      modprc_g_str_CadEje = ""
      modprc_g_str_CadEje = modprc_g_str_CadEje & "UPDATE TRA_GASADM "
      modprc_g_str_CadEje = modprc_g_str_CadEje & "   SET GASADM_NROCNT = '" & CStr(r_str_NroCnt) & "', "
      modprc_g_str_CadEje = modprc_g_str_CadEje & "       GASADM_CODOPE = " & r_str_CodOpe & ""
      modprc_g_str_CadEje = modprc_g_str_CadEje & " WHERE GASADM_NUMSOL = '" & r_str_NumSol & "'"
      modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND GASADM_CODGAS = '" & p_CodGas & "'"
         
      If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Grabar, 2) Then
         Exit Sub
      End If
      
      If grd_Listad.TextMatrix(r_int_Contad, 4) > 0 Then 'Para el Ajuste, solo se actualiza el campo GASADM_NROCNT con el asiento
         modprc_g_str_CadEje = ""
         modprc_g_str_CadEje = modprc_g_str_CadEje & "UPDATE TRA_GASADM "
         modprc_g_str_CadEje = modprc_g_str_CadEje & "   SET GASADM_NROCNT = '" & CStr(r_str_NroCnt) & "' "
         modprc_g_str_CadEje = modprc_g_str_CadEje & " WHERE GASADM_NUMSOL = '" & r_str_NumSol & "'"
         modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND GASADM_CODGAS = " & 24 & ""
            
         If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Grabar, 2) Then
            Exit Sub
         End If
      End If
   Next r_int_Contad
   
   'Agregando en la tabla CNTBL_COMAUT - PARA LAS APROBACIONES
   Call fs_InsertaCompensacion(r_str_CodOpe, CStr(r_str_NroCnt))
End Sub

Private Sub cmd_Limpia_Click()
   Call gs_LimpiaGrid(grd_Listad)
   Call fs_Activa(False)
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Public Sub gs_IngOper(ByVal p_NumOpe As String)
Dim r_int_Contad     As Integer
Dim r_int_RepNum     As Integer

   r_int_RepNum = 0

   For r_int_Contad = 0 To grd_Listad.Rows - 1
      If grd_Listad.TextMatrix(r_int_Contad, 5) = p_NumOpe Then
         r_int_RepNum = r_int_RepNum + 1
      End If
   Next r_int_Contad
   
   If r_int_RepNum > 0 Then
      MsgBox "El Cliente ya se encuentra ingresado. ", vbExclamation, modgen_g_con_OpeTra
      Exit Sub
   Else
                                                                                                                     
   If Me.cmb_GasAdm.ListIndex = -1 Then
      MsgBox "Debe seleccionar Gasto de Cierre ", vbExclamation, modgen_g_con_OpeTra
      Call gs_SetFocus(cmb_GasAdm)
      Exit Sub
   End If
            
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & " SELECT GASADM_NUMSOL, TRIM(DATGEN_TIPDOC) || '-' || TRIM(DATGEN_NUMDOC) AS DOCUMENTO ,  " 'GASADM_PAGFEC,
      g_str_Parame = g_str_Parame & "        (TRIM(E.DATGEN_APEPAT) || ' ' || TRIM(E.DATGEN_APEMAT) || ' ' ||  TRIM(E.DATGEN_NOMBRE)) AS CLIENTE, "
      g_str_Parame = g_str_Parame & "        PAGO_CLIENTE, PAGO_PROVEEDOR, SALDO"
      g_str_Parame = g_str_Parame & "   FROM (SELECT GASADM_NUMSOL, NVL(SUM(GASADM_PAGIMP),0) AS PAGO_CLIENTE, NVL(SUM(GASADM_MTOPAGPRV),0) PAGO_PROVEEDOR, "
      g_str_Parame = g_str_Parame & "                NVL(NVL(SUM(GASADM_PAGIMP), 0) - NVL(SUM(GASADM_MTOPAGPRV), 0), 0) AS SALDO " ', GASADM_PAGFEC
      g_str_Parame = g_str_Parame & "           FROM TRA_GASADM "
      g_str_Parame = g_str_Parame & "          WHERE  GASADM_CODGAS <> 13 " 'SUBSTR(GASADM_NUMSOL,1,3) NOT IN ('001','003') AND
      g_str_Parame = g_str_Parame & "          GROUP BY GASADM_NUMSOL) A " ', GASADM_PAGFEC
      g_str_Parame = g_str_Parame & "          INNER JOIN CRE_SOLMAE B ON B.SOLMAE_NUMERO = A.GASADM_NUMSOL AND SOLMAE_SITUAC IN (1,2,3) "
      g_str_Parame = g_str_Parame & "          INNER JOIN CLI_DATGEN E ON E.DATGEN_TIPDOC = B.SOLMAE_TITTDO AND E.DATGEN_NUMDOC = B.SOLMAE_TITNDO "
      'g_str_Parame = g_str_Parame & " WHERE SALDO > 0 AND GASADM_NUMSOL = '" & p_NumOpe & "' "
      g_str_Parame = g_str_Parame & "  WHERE GASADM_NUMSOL = '" & p_NumOpe & "' "
      g_str_Parame = g_str_Parame & "  ORDER BY GASADM_NUMSOL, DOCUMENTO "
   
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      'grd_Listad.Redraw = False
      
      If g_rst_Princi.BOF And g_rst_Princi.EOF Then
         MsgBox "Ésta Operación ya ha sido cancelada. ", vbExclamation, modgen_g_con_OpeTra
         g_rst_Princi.Close
         Set g_rst_Princi = Nothing
         Exit Sub
      Else
         'Verificando que el Gastos Administrativos seleccionado no se haya pagado
'         g_str_Parame = ""
'         g_str_Parame = g_str_Parame & "SELECT GASADM_NUMSOL, GASADM_MTOPAGPRV, GASADM_IMPORT  "
'         g_str_Parame = g_str_Parame & "  FROM TRA_GASADM "
'         g_str_Parame = g_str_Parame & " WHERE GASADM_NUMSOL = '" & g_rst_Princi!GASADM_NUMSOL & "' "
'         g_str_Parame = g_str_Parame & "   AND GASADM_CODGAS = '" & l_int_CodGas & "' "
'         g_str_Parame = g_str_Parame & "   AND GASADM_IMPORT > GASADM_MTOPAGPRV  "
'
'         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
'            cmd_Grabar.Enabled = False
'            Exit Sub
'         End If
'
'         If (Not (g_rst_Genera.BOF And g_rst_Genera.EOF)) Then
            grd_Listad.Rows = grd_Listad.Rows + 1
            grd_Listad.Row = grd_Listad.Rows - 1
            
            grd_Listad.Col = 0
            grd_Listad.Text = g_rst_Princi!DOCUMENTO
            
            grd_Listad.Col = 1
            grd_Listad.Text = g_rst_Princi!CLIENTE
         
            grd_Listad.Col = 2
            grd_Listad.Text = Format(g_rst_Princi!SALDO, "###,###,##0.00")
            
            grd_Listad.Col = 3
            grd_Listad.Text = Format(0, "###,###,##0.00")
              
            grd_Listad.Col = 4
            grd_Listad.Text = Format(0, "###,###,##0.00")
            
            grd_Listad.Col = 5
            grd_Listad.Text = p_NumOpe
            
            'grd_Listad.Col = 6
            'grd_Listad.Text = g_rst_Princi!GASADM_PAGFEC
            
            'grd_Listad.Col = 7
            'grd_Listad.Text = g_rst_Genera!GASADM_MTOPAGPRV
            '
            'grd_Listad.Col = 8
            'grd_Listad.Text = g_rst_Genera!GASADM_IMPORT
            
            Call gs_RefrescaGrid(grd_Listad)
'         Else
'            MsgBox "La Solicitud ya está pagada o no ha sido registrada, para el Gasto de Cierre seleccionado.", vbExclamation, modgen_g_str_NomPlt
'         End If
         
         'cierra recordset de gasto no pagado
'         g_rst_Genera.Close
'         Set g_rst_Genera = Nothing
      End If
   
      If grd_Listad.Rows > 0 Then
         fs_Activa (True)
         cmd_BusCli.Enabled = True
         Call gs_SetFocus(cmd_BusCli)

          If ipp_MtoPag.Visible Then
             grd_Listad.Text = ipp_MtoPag.Value
             ipp_MtoPag.Visible = False
          End If
      End If
      
      Call fs_Total_Saldo
   End If
End Sub

Private Sub cmd_Valida_Click()
Dim r_int_Contad  As Integer
Dim r_dbl_MtoAju  As Double

   For r_int_Contad = 0 To grd_Listad.Rows - 1
      If CDbl(grd_Listad.TextMatrix(r_int_Contad, 3)) > 0 Then
          If CDbl(grd_Listad.TextMatrix(r_int_Contad, 2)) < CDbl(grd_Listad.TextMatrix(r_int_Contad, 3)) Then
            grd_Listad.TextMatrix(r_int_Contad, 4) = Format(CDbl(CDbl(grd_Listad.TextMatrix(r_int_Contad, 3)) - CDbl(grd_Listad.TextMatrix(r_int_Contad, 2))), "###,###,###,##0.00")
          End If
      Else
         MsgBox "Existen Pagos con importe cero", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(chkSeleccionar)
         Exit Sub
      End If
   Next

  Call fs_Total_Ajuste
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Activa(False)
   Call cmd_Limpia_Click
   Call gs_CentraForm(Me)
  
   Screen.MousePointer = 0
End Sub

Private Sub fs_Activa(ByVal estado As Boolean)
   cmd_BusCli.Enabled = Not estado
   cmd_Borrar.Enabled = estado
   cmd_Limpia.Enabled = estado
   cmd_Grabar.Enabled = estado
   cmd_ExpExc.Enabled = estado
   cmd_Valida.Enabled = estado
End Sub

Private Sub fs_Inicia()
   grd_Listad.ColWidth(0) = 1700 '1660 DNI
   grd_Listad.ColWidth(1) = 4100 'CLIENTE
   grd_Listad.ColWidth(2) = 1610 'SALDO
   grd_Listad.ColWidth(3) = 1610 'PAGO
   grd_Listad.ColWidth(4) = 1610 'AJUSTE
   grd_Listad.ColWidth(5) = 0    'SOLICITUD
   'grd_Listad.ColWidth(6) = 0    'FECHA PAGO
   
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
   grd_Listad.ColAlignment(2) = flexAlignRightCenter
   grd_Listad.ColAlignment(3) = flexAlignRightCenter
   grd_Listad.ColAlignment(4) = flexAlignRightCenter
   
   'Gastos administrativos
   Call gs_Carga_ParSubPrd_Combo(cmb_GasAdm, "007")
    
   'Tipo de Documento
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipDoc, 1, "118")
   
   ipp_FecPag.Text = Format(CDate(date), "dd/mm/yyyy")
   
   'Tamaño de la celda de grd_DetAsi
   grd_Listad.RowHeightMin = ipp_MtoPag.Height
   
   'la fuente utilizada para mostrar texto
   ipp_MtoPag.FontName = grd_Listad.FontName
  
   'el tamaño de la fuente que se va a utilizar
   ipp_MtoPag.FontSize = grd_Listad.FontSize
  
   ipp_MtoPag.Visible = False
   
   'sin borde
   ipp_MtoPag.BorderStyle = vbBSNone
   
   Call moddat_gf_ConsultaPerMesActivo("000001", 1, moddat_g_str_FecIni, moddat_g_str_FecFin, moddat_g_str_CodMes, moddat_g_str_CodAno)
End Sub

Private Sub fs_Limpia()
   chkSeleccionar.Value = False
   pnl_TotSal.Caption = Format(0, "###,###,###,##0.00") & " "
   pnl_TotPag.Caption = Format(0, "###,###,###,##0.00") & " "
   pnl_TotAju.Caption = Format(0, "###,###,###,##0.00") & " "
   cmb_GasAdm.ListIndex = -1
   txt_NumDoc.Text = Empty
   pnl_RazSoc.Caption = Empty
   ipp_FecPag.Text = Format(CDate(date), "dd/mm/yyyy")
End Sub

Private Sub gs_Carga_ParSubPrd_Combo(p_Combo As ComboBox, ByVal p_CodGrp As String)
   p_Combo.Clear
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT SUBSTR(PARPRD_CODITE,1,2) AS CODGAS, PARPRD_DESCRI "
   g_str_Parame = g_str_Parame & "   FROM CRE_PARPRD "
   g_str_Parame = g_str_Parame & "  WHERE PARPRD_CODGRP = '" & p_CodGrp & "' "
   g_str_Parame = g_str_Parame & "    AND PARPRD_CODITE <> '000' "
   g_str_Parame = g_str_Parame & "    AND SUBSTR(PARPRD_CODITE,3,1) = '1' "
   g_str_Parame = g_str_Parame & "    AND SUBSTR(PARPRD_CODITE,1,2) NOT IN ('24','25','26') "
   g_str_Parame = g_str_Parame & "    AND PARPRD_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "  GROUP BY PARPRD_CODITE, PARPRD_DESCRI "
   g_str_Parame = g_str_Parame & "  ORDER BY PARPRD_CODITE ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Sub
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
     g_rst_Genera.Close
     Set g_rst_Genera = Nothing
     
     Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   Do While Not g_rst_Genera.EOF
      p_Combo.AddItem Trim$(g_rst_Genera!PARPRD_DESCRI)
      p_Combo.ItemData(p_Combo.NewIndex) = CInt(g_rst_Genera!CODGAS)
      
      g_rst_Genera.MoveNext
   Loop
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub

Private Sub grd_Listad_Click()
    Call fs_IniciaEdicion
End Sub

Private Sub fs_IniciaEdicion()
   'empieza la edicion del grid verificando de que control se hará uso (Txt, cmb, date y ipp)
   If grd_Listad.Col = 3 Then
      l_int_NumFila = grd_Listad.Row
      fs_GridEditFp Asc(" ")
   End If
End Sub

'------------- PARA PERMITIR LA EDICION DEL GRID cuando se selecciona ipp_MtoCta -----------
Private Sub fs_GridEditFp(ByVal KeyAscii As Integer)
    'posiciona el ipp encima de la celda
    If grd_Listad.Col = 3 Then
        ipp_MtoPag.Left = grd_Listad.CellLeft + grd_Listad.Left
        ipp_MtoPag.Top = grd_Listad.CellTop + grd_Listad.Top
        ipp_MtoPag.Width = grd_Listad.CellWidth
        ipp_MtoPag.Height = grd_Listad.CellHeight
        ipp_MtoPag.Visible = True
        ipp_MtoPag.Enabled = True
        ipp_MtoPag.SetFocus
    End If
    Select Case KeyAscii
        Case 0 To Asc(" ")                               'para cualquier caracter extraño que se quiera introducir
            ipp_MtoPag.Value = grd_Listad.Text
            ipp_MtoPag.SelStart = Len(ipp_MtoPag.Text)   'donde se ubica el punto inicial del txt
        Case Else
            ipp_MtoPag.Value = Chr(KeyAscii)
            ipp_MtoPag.SelStart = 1                      'coloca el cursor despues del valor valido
    End Select
End Sub

'------------- PARA PERMITIR FINALIZAR LA EDICIÓN DEL GRID cuando se selecciona ipp_MtoCta -----------
Private Sub fs_EndEditIpp(ByVal r_int_FilSel As Integer, ByVal r_int_ColSel As Integer)
   If ipp_MtoPag.Visible Then
     grd_Listad.Col = 3
     If grd_Listad.Text <> Empty Then
         grd_Listad.TextMatrix(r_int_FilSel, r_int_ColSel) = Format(ipp_MtoPag.Value, "###,###,###,##0.00")
         grd_Listad.SetFocus
         ipp_MtoPag.Visible = False
         Call fs_Total_Pago
     End If
   End If
End Sub

Private Sub ipp_FecReg_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NumDoc)
    End If
End Sub

Private Sub grd_Listad_GotFocus()
    If ipp_MtoPag.Visible Then
      grd_Listad.Text = ipp_MtoPag.Value
      ipp_MtoPag.Visible = False
   End If
End Sub

Private Sub grd_Listad_KeyPress(KeyAscii As Integer)
    Select Case grd_Listad.Col
      Case 3
          fs_GridEditFp KeyAscii
   End Select
End Sub

Private Sub grd_Listad_LeaveCell()
   If ipp_MtoPag.Visible Then
      grd_Listad.Text = ipp_MtoPag.Value
      ipp_MtoPag.Visible = False
   End If
End Sub

Private Sub grd_Listad_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   l_str_FilDel = CStr(grd_Listad.Row)
End Sub

Private Sub ipp_FecPag_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipDoc)
    End If
End Sub
 
Private Sub ipp_MtoPag_GotFocus()
   ipp_MtoPag.BackColor = modgen_g_con_ColAma
End Sub

Private Sub ipp_MtoPag_KeyDown(KeyCode As Integer, Shift As Integer)
Dim r_int_FilSel As Integer
Dim r_int_ColSel As Integer

   r_int_FilSel = grd_Listad.Row
   r_int_ColSel = grd_Listad.Col

   Select Case KeyCode
      ' keycode conjunto de constantes que se presionan ejm.f1, f2 ,space
       Case vbKeyEscape
           'salgo del Dtp sin cambiar su valor
           ipp_MtoPag.Visible = False
           grd_Listad.SetFocus
       Case vbKeyReturn
           'Finalizo la captura o entrada de datos
           Call fs_EndEditIpp(r_int_FilSel, r_int_ColSel)
       Case vbKeyDown
           ' Me muevo una fila hacia abajo
           grd_Listad.SetFocus
           DoEvents
           If grd_Listad.Row < grd_Listad.Rows - 1 Then
               grd_Listad.Row = grd_Listad.Row + 1
           End If
       Case vbKeyUp
           'Me muevo una fila hacia arriba
           grd_Listad.SetFocus
           DoEvents
           If grd_Listad.Row > grd_Listad.FixedRows Then
               grd_Listad.Row = grd_Listad.Row - 1
           End If
   End Select
End Sub

Private Sub ipp_MtoPag_KeyPress(KeyAscii As Integer)
   If (KeyAscii = vbKeyReturn) Or (KeyAscii = vbKeyEscape) Then
      KeyAscii = 0
   End If
   ' KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & Chr(22))
End Sub

Private Sub ipp_MtoPag_LostFocus()
   ipp_MtoPag.BackColor = modgen_g_con_ColAma 'l_var_ColAnt
   Call fs_Total_Pago
End Sub

Private Sub fs_Total_Saldo()
Dim r_int_Contad     As Integer
Dim r_dbl_MtoSal     As Double
   
   'Total de Saldos
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   r_dbl_MtoSal = 0
   
   For r_int_Contad = 0 To grd_Listad.Rows - 1
      grd_Listad.Row = r_int_Contad
      grd_Listad.Col = 2:  r_dbl_MtoSal = r_dbl_MtoSal + CDbl(grd_Listad.Text)
   Next r_int_Contad
        
   grd_Listad.Redraw = True
   pnl_TotSal.Caption = Format(r_dbl_MtoSal, "###,###,###,##0.00") & " "
End Sub

Private Sub fs_Total_Pago()
Dim r_int_Contad     As Integer
Dim r_dbl_MtoPag     As Double
   
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   r_dbl_MtoPag = 0
   For r_int_Contad = 0 To grd_Listad.Rows - 1
      grd_Listad.Row = r_int_Contad
      grd_Listad.Col = 3:  r_dbl_MtoPag = r_dbl_MtoPag + CDbl(grd_Listad.Text)
   Next r_int_Contad
     
   Me.pnl_TotPag.Caption = Format(r_dbl_MtoPag, "###,###,###,##0.00") & " "
End Sub

Private Sub fs_Total_Ajuste()
Dim r_int_Contad     As Integer
Dim r_dbl_MtoAju     As Double

   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   r_dbl_MtoAju = 0
   For r_int_Contad = 0 To grd_Listad.Rows - 1
      grd_Listad.Row = r_int_Contad
      grd_Listad.Col = 4: r_dbl_MtoAju = r_dbl_MtoAju + CDbl(grd_Listad.Text)
   Next r_int_Contad
   pnl_TotAju.Caption = Format(r_dbl_MtoAju, "###,###,###,##0.00") & " "
End Sub

Private Sub pnl_Tit_NomCli_Click()
   If Len(Trim(pnl_Tit_NomCli.Tag)) = 0 Or pnl_Tit_NomCli.Tag = "D" Then
      pnl_Tit_NomCli.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 1, "C")
   Else
      pnl_Tit_NomCli.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 1, "C-")
   End If
End Sub

Private Sub pnl_Tit_DocIde_Click()
   If Len(Trim(pnl_Tit_DocIde.Tag)) = 0 Or pnl_Tit_DocIde.Tag = "D" Then
      pnl_Tit_DocIde.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 2, "C")
   Else
      pnl_Tit_DocIde.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 2, "C-")
   End If
End Sub

Public Sub fs_BuscarProv()
   pnl_RazSoc.Caption = ""
   If cmb_TipDoc.ListIndex = -1 Then
      Exit Sub
   End If
   If Trim(txt_NumDoc.Text) = "" Then
      Exit Sub
   End If
    
    g_str_Parame = ""
    g_str_Parame = g_str_Parame & " SELECT MAEPRV_TIPDOC, MAEPRV_NUMDOC, MAEPRV_RAZSOC, MAEPRV_CODBNC_MN1, MAEPRV_CTACRR_MN1, MAEPRV_NROCCI_MN1 "
    g_str_Parame = g_str_Parame & "   FROM CNTBL_MAEPRV A  "
    g_str_Parame = g_str_Parame & "       INNER JOIN MNT_PARDES B ON A.MAEPRV_TIPCNT = B.PARDES_CODITE AND B.PARDES_CODGRP = 119  "
    g_str_Parame = g_str_Parame & "       INNER JOIN MNT_PARDES C ON A.MAEPRV_CONDIC = C.PARDES_CODITE AND C.PARDES_CODGRP = 120  "
    g_str_Parame = g_str_Parame & "       INNER JOIN MNT_PARDES D ON A.MAEPRV_TIPPER = D.PARDES_CODITE AND D.PARDES_CODGRP = 127  "
    g_str_Parame = g_str_Parame & "  WHERE MAEPRV_SITUAC = 1  "
    If Len(Trim(txt_NumDoc.Text)) > 0 Then
       g_str_Parame = g_str_Parame & "   AND MAEPRV_TIPDOC = " & cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) & " "
       g_str_Parame = g_str_Parame & "   AND MAEPRV_NUMDOC = '" & Trim(txt_NumDoc.Text) & "' "
    End If
    g_str_Parame = g_str_Parame & "  ORDER BY MAEPRV_TIPDOC, MAEPRV_RAZSOC ASC  "
    
    If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
    End If
    
    If g_rst_Princi.BOF And g_rst_Princi.EOF Then
       g_rst_Princi.Close
       Set g_rst_Princi = Nothing
       Screen.MousePointer = 0
       Exit Sub
    End If
    
    g_rst_Princi.MoveFirst
    If Not g_rst_Princi.EOF Then
       If Trim(g_rst_Princi!MAEPRV_CODBNC_MN1) = 0 Then
          If MsgBox("El proveedor seleccionado no tiene ninguna cuenta corriente asignada, favor consulte con contabilidad." & vbCrLf & "Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
             txt_NumDoc.Text = ""
             pnl_RazSoc.Caption = ""
             Exit Sub
          End If
       End If
       
       pnl_RazSoc.Caption = Trim(g_rst_Princi!MAEPRV_RAZSOC & "")
       If Trim(g_rst_Princi!MAEPRV_CODBNC_MN1) <> 0 Then
          moddat_g_str_CodIte = Trim(g_rst_Princi!MAEPRV_CODBNC_MN1)
          If CInt(moddat_g_str_CodIte) = 11 Then                         'SI ES BBVA, ENVIAR CTACTE
            moddat_g_str_DesIte = Trim(g_rst_Princi!MAEPRV_CTACRR_MN1)
          Else
            moddat_g_str_DesIte = Trim(g_rst_Princi!MAEPRV_NROCCI_MN1)
          End If
       End If
    End If
    
    g_rst_Princi.Close
    Set g_rst_Princi = Nothing
    Screen.MousePointer = 0
End Sub

Private Sub txt_NumDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        'Call fs_BuscarProv
        Call txt_NumDoc_LostFocus
        Call gs_SetFocus(cmd_BusCli)
    Else
        KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
    End If
End Sub

Private Sub txt_NumDoc_LostFocus()
   Call fs_BuscarProv
End Sub
