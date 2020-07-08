VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_Caj_CiePag_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form4"
   ClientHeight    =   9255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9825
   Icon            =   "OpeTra_frm_806.frx":0000
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9255
   ScaleWidth      =   9825
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel2 
      Height          =   9255
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   9825
      _Version        =   65536
      _ExtentX        =   17330
      _ExtentY        =   16325
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
      Begin Threed.SSPanel SSPanel5 
         Height          =   1875
         Left            =   30
         TabIndex        =   9
         Top             =   7335
         Width           =   9735
         _Version        =   65536
         _ExtentX        =   17171
         _ExtentY        =   3307
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
         Begin VB.TextBox txt_NumDoc 
            Height          =   315
            Left            =   2730
            MaxLength       =   25
            TabIndex        =   3
            Top             =   1140
            Width           =   2775
         End
         Begin EditLib.fpDateTime ipp_FecPag 
            Height          =   315
            Left            =   2730
            TabIndex        =   2
            Top             =   840
            Width           =   1515
            _Version        =   196608
            _ExtentX        =   2672
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
         Begin EditLib.fpDoubleSingle ipp_MtoPag 
            Height          =   315
            Left            =   2730
            TabIndex        =   4
            Top             =   1470
            Width           =   1485
            _Version        =   196608
            _ExtentX        =   2619
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
            MaxValue        =   "900000"
            MinValue        =   "0"
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
         Begin Threed.SSPanel pnl_Concepto 
            Height          =   315
            Left            =   2730
            TabIndex        =   32
            Top             =   150
            Width           =   6585
            _Version        =   65536
            _ExtentX        =   11615
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
         Begin Threed.SSPanel pnl_Importe 
            Height          =   315
            Left            =   2730
            TabIndex        =   33
            Top             =   495
            Width           =   1485
            _Version        =   65536
            _ExtentX        =   2619
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
            Alignment       =   4
         End
         Begin VB.Label Label5 
            Caption         =   "Importe Pagado por el Cliente"
            Height          =   315
            Left            =   120
            TabIndex        =   35
            Top             =   525
            Width           =   2595
         End
         Begin VB.Label Label4 
            Caption         =   "Concepto de Gasto"
            Height          =   315
            Left            =   120
            TabIndex        =   34
            Top             =   180
            Width           =   2595
         End
         Begin VB.Label Label11 
            Caption         =   "Monto Pagado al Proveedor"
            Height          =   285
            Left            =   120
            TabIndex        =   12
            Top             =   1515
            Width           =   2595
         End
         Begin VB.Label Label10 
            Caption         =   "Nro Operacion:"
            Height          =   315
            Left            =   120
            TabIndex        =   11
            Top             =   1200
            Width           =   2595
         End
         Begin VB.Label Label6 
            Caption         =   "Fecha de Pago al Proveedor:"
            Height          =   315
            Left            =   120
            TabIndex        =   10
            Top             =   870
            Width           =   2595
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   13
         Top             =   30
         Width           =   9735
         _Version        =   65536
         _ExtentX        =   17171
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
            Left            =   600
            TabIndex        =   14
            Top             =   60
            Width           =   5000
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
            Left            =   600
            TabIndex        =   15
            Top             =   330
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
            Left            =   60
            Picture         =   "OpeTra_frm_806.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   1875
         Left            =   30
         TabIndex        =   16
         Top             =   1440
         Width           =   9735
         _Version        =   65536
         _ExtentX        =   17171
         _ExtentY        =   3307
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
            Left            =   1620
            TabIndex        =   17
            Top             =   120
            Width           =   1905
            _Version        =   65536
            _ExtentX        =   3360
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
         Begin Threed.SSPanel pnl_NomCli 
            Height          =   315
            Left            =   1620
            TabIndex        =   18
            Top             =   840
            Width           =   6100
            _Version        =   65536
            _ExtentX        =   10760
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
         Begin Threed.SSPanel pnl_Moneda 
            Height          =   315
            Left            =   4980
            TabIndex        =   19
            Top             =   1200
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
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
         Begin Threed.SSPanel pnl_DocIde 
            Height          =   315
            Left            =   1620
            TabIndex        =   20
            Top             =   480
            Width           =   1905
            _Version        =   65536
            _ExtentX        =   3360
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
         Begin Threed.SSPanel pnl_FechaPago 
            Height          =   315
            Left            =   1620
            TabIndex        =   36
            Top             =   1560
            Width           =   1905
            _Version        =   65536
            _ExtentX        =   3360
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
         Begin Threed.SSPanel pnl_ImpSal 
            Height          =   315
            Left            =   1620
            TabIndex        =   45
            Top             =   1200
            Width           =   1905
            _Version        =   65536
            _ExtentX        =   3360
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_Estado 
            Height          =   315
            Left            =   4950
            TabIndex        =   48
            Top             =   120
            Width           =   2745
            _Version        =   65536
            _ExtentX        =   4842
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
         Begin VB.Label Label14 
            Caption         =   "Estado:"
            Height          =   315
            Left            =   4050
            TabIndex        =   49
            Top             =   150
            Width           =   915
         End
         Begin VB.Label Label13 
            Caption         =   "Saldo:"
            Height          =   315
            Left            =   120
            TabIndex        =   46
            Top             =   1230
            Width           =   1275
         End
         Begin VB.Label Label9 
            Caption         =   "F. Pago Cliente:"
            Height          =   225
            Left            =   120
            TabIndex        =   37
            Top             =   1590
            Width           =   1395
         End
         Begin VB.Label Label3 
            Caption         =   "Nombre Cliente:"
            Height          =   315
            Left            =   120
            TabIndex        =   24
            Top             =   870
            Width           =   1155
         End
         Begin VB.Label Label2 
            Caption         =   "Nro. de Solicitud:"
            Height          =   315
            Left            =   120
            TabIndex        =   23
            Top             =   150
            Width           =   1395
         End
         Begin VB.Label Label1 
            Caption         =   "Moneda:"
            Height          =   315
            Left            =   4050
            TabIndex        =   22
            Top             =   1230
            Width           =   915
         End
         Begin VB.Label Label12 
            Caption         =   "DOI Cliente:"
            Height          =   315
            Left            =   120
            TabIndex        =   21
            Top             =   510
            Width           =   1395
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   2400
         Left            =   30
         TabIndex        =   25
         Top             =   3360
         Width           =   9735
         _Version        =   65536
         _ExtentX        =   17171
         _ExtentY        =   4233
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
            Height          =   2025
            Left            =   0
            TabIndex        =   0
            Top             =   330
            Width           =   9705
            _ExtentX        =   17119
            _ExtentY        =   3572
            _Version        =   393216
            Rows            =   8
            Cols            =   6
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel8 
            Height          =   285
            Left            =   5490
            TabIndex        =   29
            Top             =   60
            Width           =   1245
            _Version        =   65536
            _ExtentX        =   2196
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Fec. Pago"
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
            Left            =   4335
            TabIndex        =   27
            Top             =   60
            Width           =   1170
            _Version        =   65536
            _ExtentX        =   2064
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
         Begin Threed.SSPanel SSPanel13 
            Height          =   285
            Left            =   8175
            TabIndex        =   30
            Top             =   60
            Width           =   1230
            _Version        =   65536
            _ExtentX        =   2170
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Monto Pag."
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
            Left            =   6735
            TabIndex        =   31
            Top             =   60
            Width           =   1440
            _Version        =   65536
            _ExtentX        =   2540
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "N° Documento"
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
            Left            =   30
            TabIndex        =   26
            Top             =   60
            Width           =   4320
            _Version        =   65536
            _ExtentX        =   7620
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
      End
      Begin Threed.SSPanel SSPanel12 
         Height          =   645
         Left            =   30
         TabIndex        =   28
         Top             =   750
         Width           =   9735
         _Version        =   65536
         _ExtentX        =   17171
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
         Begin VB.CommandButton cmd_Borrar 
            Height          =   585
            Left            =   1230
            Picture         =   "OpeTra_frm_806.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   47
            ToolTipText     =   "Borrar Registro"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_NueIte 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_806.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "Nuevo Registro"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Editar 
            Height          =   585
            Left            =   630
            Picture         =   "OpeTra_frm_806.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Editar Registro"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   9135
            Picture         =   "OpeTra_frm_806.frx":0C34
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Salida"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Cancel 
            Height          =   585
            Left            =   2385
            Picture         =   "OpeTra_frm_806.frx":1076
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Cancelar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   1800
            Picture         =   "OpeTra_frm_806.frx":1380
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel16 
         Height          =   1500
         Left            =   30
         TabIndex        =   39
         Top             =   5805
         Width           =   9735
         _Version        =   65536
         _ExtentX        =   17171
         _ExtentY        =   2646
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
         Begin VB.TextBox txt_NumOper 
            Height          =   315
            Left            =   2730
            MaxLength       =   25
            TabIndex        =   50
            Top             =   750
            Width           =   2775
         End
         Begin VB.ComboBox cmb_GasAdm 
            Height          =   315
            Left            =   2730
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   90
            Width           =   3855
         End
         Begin EditLib.fpDoubleSingle ipp_Import 
            Height          =   315
            Left            =   2730
            TabIndex        =   41
            Top             =   420
            Width           =   975
            _Version        =   196608
            _ExtentX        =   1720
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
         Begin Threed.SSPanel pnl_MonGas 
            Height          =   315
            Left            =   3750
            TabIndex        =   42
            Top             =   420
            Width           =   2835
            _Version        =   65536
            _ExtentX        =   5001
            _ExtentY        =   556
            _StockProps     =   15
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
            Alignment       =   1
         End
         Begin EditLib.fpDateTime ipp_FecGasCie 
            Height          =   315
            Left            =   2730
            TabIndex        =   52
            Top             =   1080
            Width           =   1515
            _Version        =   196608
            _ExtentX        =   2672
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
         Begin VB.Label Label16 
            Caption         =   "Fecha de Pago al Proveedor:"
            Height          =   315
            Left            =   120
            TabIndex        =   53
            Top             =   1110
            Width           =   2595
         End
         Begin VB.Label Label15 
            Caption         =   "Nro Operacion:"
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   780
            Width           =   2055
         End
         Begin VB.Label Label8 
            Caption         =   "Importe:"
            Height          =   285
            Left            =   120
            TabIndex        =   44
            Top             =   450
            Width           =   1485
         End
         Begin VB.Label Label7 
            Caption         =   "Concepto de Gasto:"
            Height          =   285
            Left            =   120
            TabIndex        =   43
            Top             =   90
            Width           =   1485
         End
      End
   End
End
Attribute VB_Name = "frm_Caj_CiePag_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_int_CodGas     As Integer
Dim l_arr_GasAdm()   As moddat_tpo_Genera
Dim l_bol_estado     As Boolean

Private Sub cmb_GasAdm_Click()
    Call gs_SetFocus(ipp_Import)
End Sub

Private Sub cmb_GasAdm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_GasAdm_Click
   End If
End Sub

Private Sub cmd_Borrar_Click()
Dim r_int_Contad As Integer

   r_int_Contad = grd_Listad.Row

   If CStr(grd_Listad.TextMatrix(r_int_Contad, 2)) <> 24 Then
      MsgBox "No se puede eliminar el Gasto de Cierre Seleccionado.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(grd_Listad)
      Exit Sub
   End If
   
 'Confirma
   MsgBox "Recuerde eliminar el Asiento Contable que se registró.", vbExclamation, modgen_g_str_NomPlt
    
   If MsgBox("¿Está seguro de Eliminar el Gasto de Cierre seleccionado?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   moddat_g_int_FlgGOK = False
   
   If grd_Listad.Rows = 1 Then
      Call gs_LimpiaGrid(grd_Listad)
   Else
      
      If r_int_Contad = grd_Listad.Rows Then
         Exit Sub
      End If
      If r_int_Contad = grd_Listad.Row And r_int_Contad = 0 Then
         Call gs_LimpiaGrid(grd_Listad)
         Exit Sub
      End If
      
      'Elimina el Ajuste de Gasto de Cierre
      Do While moddat_g_int_FlgGOK = False
         g_str_Parame = "USP_TRA_GASADM_PAGOPRV_2 ("
         g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumSol & "', "
         g_str_Parame = g_str_Parame & CStr(grd_Listad.TextMatrix(r_int_Contad, 2)) & ", "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'', "
         
         'Datos de Auditoria
         g_str_Parame = g_str_Parame & "'', "                              'Código Usuario
         g_str_Parame = g_str_Parame & "'', "                              'Nombre Terminal
         g_str_Parame = g_str_Parame & "'', "                              'Nombre Ejecutable
         g_str_Parame = g_str_Parame & "'', "                              'Código Sucursal
         g_str_Parame = g_str_Parame & "2)"
            
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
            moddat_g_int_CntErr = moddat_g_int_CntErr + 1
         Else
            moddat_g_int_FlgGOK = True
         End If
   
         If moddat_g_int_CntErr = 6 Then
            If MsgBox("No se pudo completar el procedimiento USP_TRA_GASADM_PAGOPRV_2. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
               Exit Sub
            Else
               moddat_g_int_CntErr = 0
            End If
         End If
      Loop
      grd_Listad.RemoveItem (r_int_Contad)
    End If
End Sub

Private Sub cmd_Cancel_Click()
   Call fs_LimpiaItem
   Call fs_Habilitar(False)
   grd_Listad.SetFocus
End Sub

Private Sub cmd_Editar_Click()
   Call grd_Listad_DblClick
End Sub

Private Sub cmd_Grabar_Click()
Dim l_int_TipPag     As Integer

   '*** validaciones para ingreso de nuevo gasto de cierre
   If moddat_g_int_FlgGrb = 1 Then
        If cmb_GasAdm.ListIndex = -1 Then
           MsgBox "Debe seleccionar el Gasto Administrativo.", vbExclamation, modgen_g_str_NomPlt
           Call gs_SetFocus(cmb_GasAdm)
           Exit Sub
        End If
        If ipp_Import.Value = 0 Then
           MsgBox "Debe ingresar el Importe del Gasto.", vbExclamation, modgen_g_str_NomPlt
           Call gs_SetFocus(ipp_Import)
           Exit Sub
        End If
        If Trim(txt_NumOper.Text) = "" Then
           MsgBox "Debe de ingresar un numero de operación.", vbExclamation, modgen_g_str_NomPlt
           Call gs_SetFocus(txt_NumOper)
           Exit Sub
        End If
      If CDate(ipp_FecGasCie.Text) > date Then
           MsgBox "Debe ingresar una fecha correcta.", vbExclamation, modgen_g_str_NomPlt
           Call gs_SetFocus(ipp_FecGasCie)
           Exit Sub
        End If
        If cmb_GasAdm.ItemData(cmb_GasAdm.ListIndex) = 26 Then
            If CDbl(ipp_Import.Value) <> CDbl(pnl_ImpSal.Caption) Then
               MsgBox "El Monto no es igual al Saldo.", vbExclamation, modgen_g_str_NomPlt
               Call gs_SetFocus(ipp_Import)
               Exit Sub
            End If
        End If
        'Validar que el Gasto no este ingresado si es Agregar
        If moddat_g_int_FlgGrb = 1 Then
           g_str_Parame = "SELECT * FROM TRA_GASADM WHERE "
           g_str_Parame = g_str_Parame & " GASADM_NUMSOL = '" & moddat_g_str_NumSol & "' AND "
           g_str_Parame = g_str_Parame & " GASADM_CODGAS = " & cmb_GasAdm.ItemData(cmb_GasAdm.ListIndex)  'Left(l_arr_GasAdm(cmb_GasAdm.ListIndex + 1).Genera_Codigo, 2)
        
           If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
              Exit Sub
           End If
           
           If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
              g_rst_Genera.Close
              Set g_rst_Genera = Nothing
              MsgBox "El Gasto Administrativo ya fue ingresado.", vbExclamation, modgen_g_str_NomPlt
             
              Call gs_SetFocus(cmb_GasAdm)
              Exit Sub
           End If
           
           g_rst_Genera.Close
           Set g_rst_Genera = Nothing
        End If
   Else
        If CDate(ipp_FecPag.Text) > date Then
           MsgBox "Debe ingresar una fecha correcta.", vbExclamation, modgen_g_str_NomPlt
           Call gs_SetFocus(ipp_FecPag)
           Exit Sub
        End If
        If l_int_CodGas <> 19 Then
           If Len(Trim(txt_NumDoc.Text)) = 0 Then
              MsgBox "Debe ingresar el Número de Documento.", vbExclamation, modgen_g_str_NomPlt
              Call gs_SetFocus(txt_NumDoc)
              Exit Sub
           End If
        End If
        If ipp_MtoPag.Value = 0 Then
           MsgBox "Debe ingresar el Monto Pagado.", vbExclamation, modgen_g_str_NomPlt
           Call gs_SetFocus(ipp_MtoPag)
           Exit Sub
        End If
        If CDbl(ipp_MtoPag.Value) > CDbl(pnl_ImpSal.Caption) Then
           MsgBox "El Monto Pagado es mayor al Saldo.", vbExclamation, modgen_g_str_NomPlt
           Call gs_SetFocus(ipp_MtoPag)
           Exit Sub
        End If
   End If

   '*** confirmacion
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
     
   '*** grabacion
   Screen.MousePointer = 11
   
   If moddat_g_int_FlgGrb = 1 Then
        moddat_g_int_FlgGOK = False
        moddat_g_int_CntErr = 0
        
        Do While moddat_g_int_FlgGOK = False
           Screen.MousePointer = 11
           Call moddat_gs_FecSis
         
           If cmb_GasAdm.ItemData(cmb_GasAdm.ListIndex) = 25 Then                      'REINGRESO
               g_str_Parame = "USP_TRA_GASADM_AJUSTE ("
               g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumSol & "', "
               g_str_Parame = g_str_Parame & cmb_GasAdm.ItemData(cmb_GasAdm.ListIndex) & ", "
               g_str_Parame = g_str_Parame & 1 & ", "
               g_str_Parame = g_str_Parame & CStr(CDbl(ipp_Import.Text)) & ", "
               g_str_Parame = g_str_Parame & Format(CDate(moddat_g_str_FecSis), "yyyymmdd") & ", "
               g_str_Parame = g_str_Parame & Format(CDate(ipp_FecGasCie.Value), "yyyymmdd") & ", " 'pnl_FechaPago.Caption
               g_str_Parame = g_str_Parame & "1, "
               g_str_Parame = g_str_Parame & "'', "
               g_str_Parame = g_str_Parame & "'" & Trim(txt_NumOper.Text) & "', "
               g_str_Parame = g_str_Parame & "'', "
               g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
               g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
               g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
               g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
               g_str_Parame = g_str_Parame & CStr(moddat_g_int_FlgGrb) & ", "
               g_str_Parame = g_str_Parame & "1)"
               
           ElseIf cmb_GasAdm.ItemData(cmb_GasAdm.ListIndex) = 26 Then                  'DEVOLUCIÓN
               g_str_Parame = "USP_TRA_GASADM_AJUSTE ("
               g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumSol & "', "
               g_str_Parame = g_str_Parame & cmb_GasAdm.ItemData(cmb_GasAdm.ListIndex) & ", "
               g_str_Parame = g_str_Parame & 1 & ", "
               g_str_Parame = g_str_Parame & CStr(CDbl(0)) & ", "                                  'ipp_Import.Text
               g_str_Parame = g_str_Parame & "'', "
               g_str_Parame = g_str_Parame & Format(CDate(ipp_FecGasCie.Value), "yyyymmdd") & ", " 'pnl_FechaPago.Caption
               g_str_Parame = g_str_Parame & "1, "
               g_str_Parame = g_str_Parame & Format(CDate(moddat_g_str_FecSis), "yyyymmdd") & ", "
               g_str_Parame = g_str_Parame & "'" & Trim(txt_NumOper.Text) & "', "
               g_str_Parame = g_str_Parame & CStr(CDbl(ipp_Import.Text)) & ", "
               g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
               g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
               g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
               g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
               g_str_Parame = g_str_Parame & CStr(moddat_g_int_FlgGrb) & ", "
               g_str_Parame = g_str_Parame & "1)"
               
           ElseIf cmb_GasAdm.ItemData(cmb_GasAdm.ListIndex) = 24 Then                  'AJUSTE DE GASTOS DE CIERRE
           
               g_str_Parame = "USP_TRA_GASADM_AJUSTE ("
               g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumSol & "', " 'NRO_SOLI
               g_str_Parame = g_str_Parame & cmb_GasAdm.ItemData(cmb_GasAdm.ListIndex) & ", "
               g_str_Parame = g_str_Parame & 1 & ", " 'COD_MON
               g_str_Parame = g_str_Parame & CStr(CDbl(0)) & ", "
               g_str_Parame = g_str_Parame & "null, "
               g_str_Parame = g_str_Parame & Format(CDate(ipp_FecGasCie.Value), "yyyymmdd") & ", " 'ipp_FecPag.Text
               g_str_Parame = g_str_Parame & "1, "
               g_str_Parame = g_str_Parame & Format(CDate(moddat_g_str_FecSis), "yyyymmdd") & ", "
               g_str_Parame = g_str_Parame & "'" & Trim(txt_NumOper.Text) & "', "
               g_str_Parame = g_str_Parame & CStr(CDbl(ipp_Import.Text)) & ", "
               g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
               g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
               g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
               g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
               g_str_Parame = g_str_Parame & CStr(moddat_g_int_FlgGrb) & ", "
               g_str_Parame = g_str_Parame & "1)"
               
           End If
           
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
        
        If moddat_g_int_FlgGOK = True Then
            If cmb_GasAdm.ItemData(cmb_GasAdm.ListIndex) = 25 Then               'REINGRESO
               Call fs_GeneraAsiento(moddat_g_str_NumSol, cmb_GasAdm.ItemData(cmb_GasAdm.ListIndex), "111301060201", "251419010114", 1, ipp_Import.Text, cmb_GasAdm.Text) '291807010112
            ElseIf cmb_GasAdm.ItemData(cmb_GasAdm.ListIndex) = 26 Then           'DEVOLUCION
               Call fs_GeneraAsiento(moddat_g_str_NumSol, cmb_GasAdm.ItemData(cmb_GasAdm.ListIndex), "251419010114", "251419010109", 1, ipp_Import.Text, cmb_GasAdm.Text) '291807010112
            ElseIf cmb_GasAdm.ItemData(cmb_GasAdm.ListIndex) = 24 Then           'AJUSTE DE GASTOS DE CIERRE
               Call fs_GeneraAsiento_Ajuste(moddat_g_str_NumSol, cmb_GasAdm.ItemData(cmb_GasAdm.ListIndex))
            End If
        End If
        
        Call fs_LimpiaItem
        Call fs_ActivaItem(True)
        moddat_g_int_FlgGrb = 2
   Else
        If Not fs_PagPrv_GasAdm(moddat_g_str_NumSol, l_int_CodGas, Format(CDate(ipp_FecPag.Text), "yyyymmdd"), txt_NumDoc, ipp_MtoPag) Then
           Exit Sub
        Else
            Call fs_GeneraAsiento(moddat_g_str_NumSol, l_int_CodGas, "251419010114", "251419010109", 2, ipp_MtoPag.Text, cmb_GasAdm.Text) '291807010112
        End If
   End If
   
   grd_Listad.Redraw = True
   Call fs_Buscar
   Call fs_Habilitar(False)
   frm_Caj_CiePag_01.fs_Buscar
   grd_Listad.SetFocus
   Screen.MousePointer = 0
End Sub

Private Sub fs_GeneraAsiento(ByVal p_NumSol As String, ByVal p_CodGas As Integer, ByVal p_CtaDeb As String, ByVal p_CtaHab As String, ByVal p_Tipo As Integer, ByVal p_Monto As Double, ByVal p_Glosa As String)
Dim r_arr_LogPro()      As modprc_g_tpo_LogPro
Dim r_int_Contad        As Integer
Dim r_int_NumIte        As Integer
Dim r_str_AsiGen        As String
Dim r_str_Origen        As String
Dim r_str_TipNot        As String
Dim r_int_NumLib        As Integer
Dim r_int_NumAsi        As Integer
Dim r_dbl_TipSbs        As Double
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
Dim r_str_CodOpe        As String

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
   r_int_NumAsi = 0 'Inicializa variables
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
      
      If p_Tipo = 1 Then r_str_Glosa = p_Glosa Else r_str_Glosa = "GASTOS DE CIERRE"
   
      'Insertar en CABECERA
      Call modprc_fs_Inserta_CabAsi(r_arr_LogPro, r_str_Origen, CInt(moddat_g_str_CodAno), CInt(moddat_g_str_CodMes), r_int_NumLib, r_int_NumAsi, Format(1, "000"), r_dbl_TipCam, r_str_TipNot, Trim(r_str_Glosa), r_str_FecCon, "1")
   End If
   
      '*************************************************
      'GENERACION DE ASIENTOS CONTABLES DE GASTOS DE CIERRE
      '*************************************************
      
      If p_Monto > 0 Then
          For r_int_ConAux = 1 To 2
              r_dbl_importe = p_Monto
              If r_int_ConAux = 1 Then r_str_DebHab = "D": r_str_CtaCtb = p_CtaDeb Else r_str_DebHab = "H": r_str_CtaCtb = p_CtaHab
              
              If moddat_g_int_FlgGrb = 1 Then
                 r_str_Glosa = pnl_DocIde.Caption & "/" & txt_NumOper.Text & "/" & IIf(p_Tipo = 1, p_Glosa, "GASTO DE CIERRE")
              Else
                 r_str_Glosa = pnl_DocIde.Caption & "/" & txt_NumDoc.Text & "/" & IIf(p_Tipo = 1, p_Glosa, "GASTO DE CIERRE")
              End If
              
              If (r_dbl_importe > 0) Then
                  r_int_NumIte = r_int_NumIte + 1
                  r_dbl_MtoSol = Format(r_dbl_importe, "###,###,##0.00")
                  r_dbl_MtoDol = Format(0, "###,###,##0.00")
                  
                  Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, CInt(moddat_g_str_CodAno), CInt(moddat_g_str_CodMes), r_int_NumLib, r_int_NumAsi, r_int_NumIte, r_str_CtaCtb, CDate(r_str_FecCon), r_str_Glosa, r_str_DebHab, r_dbl_MtoSol, r_dbl_MtoDol, 1, CDate(r_str_FecCon))
                  r_dbl_importe = 0
              End If
          Next r_int_ConAux
          
          r_str_NroCnt = r_str_Origen & "/" & moddat_g_str_CodAno & "/" & Format(moddat_g_str_CodMes, "00") & "/" & r_int_NumLib & "/" & r_int_NumAsi
          '----Implementacion 15-05-2017
          If cmb_GasAdm.ListIndex <> -1 Then
             If cmb_GasAdm.ItemData(cmb_GasAdm.ListIndex) = 26 Then 'DEVOLUCION
                r_str_CodOpe = modmip_gf_Genera_CodGen(3, 3)
                  
                modprc_g_str_CadEje = ""
                modprc_g_str_CadEje = modprc_g_str_CadEje & "UPDATE TRA_GASADM "
                modprc_g_str_CadEje = modprc_g_str_CadEje & "   SET GASADM_NROCNT = '" & CStr(r_str_NroCnt) & "', "
                modprc_g_str_CadEje = modprc_g_str_CadEje & "       GASADM_CODOPE = " & r_str_CodOpe & ""
                modprc_g_str_CadEje = modprc_g_str_CadEje & " WHERE GASADM_NUMSOL = '" & p_NumSol & "'"
                modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND GASADM_CODGAS = '" & p_CodGas & "'"
                   
                If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Grabar, 2) Then
                   Exit Sub
                End If
             End If
          End If
          '----Implementacion fin
          
          'Grabando en tra_gasadm, año/mes/nro_libro/nro_asiento
          modprc_g_str_CadEje = ""
          modprc_g_str_CadEje = modprc_g_str_CadEje & "UPDATE TRA_GASADM "
          modprc_g_str_CadEje = modprc_g_str_CadEje & "   SET GASADM_NROCNT = '" & CStr(r_str_NroCnt) & "'"
          modprc_g_str_CadEje = modprc_g_str_CadEje & " WHERE GASADM_NUMSOL = '" & p_NumSol & "'"
          modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND GASADM_CODGAS = '" & p_CodGas & "'"
            
          If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Grabar, 2) Then
             Exit Sub
          End If
        
      End If
         
   If cmb_GasAdm.ListIndex <> -1 Then
      If cmb_GasAdm.ItemData(cmb_GasAdm.ListIndex) = 26 Then '----DEVOLUCION
         'Agregando en la tabla CNTBL_COMAUT - PARA LAS APROBACIONES
         Call fs_InsertaCompensacion(r_str_CodOpe, CStr(r_str_NroCnt))
      End If
   End If
End Sub

Private Sub fs_GeneraAsiento_Ajuste(ByVal p_NumSol As String, ByVal p_CodGas As Integer)
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
   
    For r_int_ConAux = 1 To 2
        r_dbl_ValImp = ipp_Import.Text  'Importe del Ajuste
            
        If r_int_ConAux = 1 Then r_str_DebHab = "D": r_str_CtaCtb = "251419010114" Else r_str_DebHab = "H": r_str_CtaCtb = "451301290110" '291807010112
            
           r_str_Glosa = Mid(pnl_DocIde.Caption, 3) & "/" & Trim(txt_NumOper.Text) & "/" & "AJUSTE GASTO DE CIERRE"
             
           If (r_dbl_ValImp > 0) Then
                r_int_NumIte = r_int_NumIte + 1
                r_dbl_MtoSol = Format(r_dbl_ValImp, "###,###,##0.00")
                r_dbl_MtoDol = Format(0, "###,###,##0.00")
                
                Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, CInt(moddat_g_str_CodAno), CInt(moddat_g_str_CodMes), r_int_NumLib, r_int_NumAsi, r_int_NumIte, r_str_CtaCtb, CDate(r_str_FecCon), r_str_Glosa, r_str_DebHab, r_dbl_MtoSol, r_dbl_MtoDol, 1, CDate(r_str_FecCon))
                r_dbl_ValImp = 0
           End If
    Next r_int_ConAux
    
    r_str_NroCnt = r_str_Origen & "/" & moddat_g_str_CodAno & "/" & Format(moddat_g_str_CodMes, "00") & "/" & r_int_NumLib & "/" & r_int_NumAsi
    'Grabando en tra_gasadm, año/mes/nro_libro/nro_asiento
    modprc_g_str_CadEje = ""
    modprc_g_str_CadEje = modprc_g_str_CadEje & "UPDATE TRA_GASADM "
    modprc_g_str_CadEje = modprc_g_str_CadEje & "   SET GASADM_NROCNT = '" & CStr(r_str_NroCnt) & "'"
    modprc_g_str_CadEje = modprc_g_str_CadEje & " WHERE GASADM_NUMSOL = '" & p_NumSol & "'"
    modprc_g_str_CadEje = modprc_g_str_CadEje & "   AND GASADM_CODGAS = '" & p_CodGas & "'"
     
    If Not gf_EjecutaSQL(modprc_g_str_CadEje, modprc_g_rst_Grabar, 2) Then
       Exit Sub
    End If
          
    'Agregando en la tabla CNTBL_COMAUT - PARA LAS APROBACIONES
    'Call fs_InsertaCompensacion(r_str_CodOpe, CStr(r_str_NroCnt))
End Sub

Private Sub fs_InsertaCompensacion(ByVal p_CodOpe As String, ByVal p_DatCta As String)
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
           
   Do While moddat_g_int_FlgGOK = False
      Screen.MousePointer = 11
      
      g_str_Parame = "USP_CNTBL_COMAUT ("
      g_str_Parame = g_str_Parame & "'" & p_CodOpe & "', "
      g_str_Parame = g_str_Parame & Format(CDate(ipp_FecPag.Text), "yyyymmdd") & ", "
      g_str_Parame = g_str_Parame & CStr(moddat_g_int_TipDoc) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumDoc & "', "
      g_str_Parame = g_str_Parame & 1 & ", "                                           'Tipo de Moneda
      g_str_Parame = g_str_Parame & CDbl(ipp_Import.Text) & ", "
      g_str_Parame = g_str_Parame & " NULL, "                  'Código del Banco - Proveedor
      g_str_Parame = g_str_Parame & "'', "                  'Cuenta Corriente - Proveedor
      g_str_Parame = g_str_Parame & "'251419010109', "
      g_str_Parame = g_str_Parame & "'" & p_DatCta & "', "
      g_str_Parame = g_str_Parame & "'PAGO POR CLIENTE', "
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

Private Sub cmd_NueIte_Click()
   moddat_g_int_FlgGrb = 1
   Call fs_ActivaItem(True)
   Call gs_SetFocus(cmb_GasAdm)
End Sub
Private Sub fs_ActivaItem(ByVal p_Habilita As Integer)
   cmb_GasAdm.Enabled = p_Habilita
   ipp_Import.Enabled = p_Habilita
   txt_NumOper.Enabled = p_Habilita
   ipp_FecGasCie.Enabled = p_Habilita
   
   cmd_Grabar.Enabled = p_Habilita
   cmd_Cancel.Enabled = p_Habilita
   
   cmd_NueIte.Enabled = Not p_Habilita
   cmd_Borrar.Enabled = Not p_Habilita
   cmd_Editar.Enabled = Not p_Habilita
End Sub
Private Sub fs_LimpiaItem()
   cmb_GasAdm.ListIndex = -1
   ipp_Import.Value = 0
   txt_NumOper.Text = ""
   
   pnl_MonGas.Caption = ""
End Sub
Private Sub cmd_Salida_Click()
   Unload Me
End Sub
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
Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Buscar
   Call fs_Habilitar(False)
      
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   grd_Listad.ColWidth(0) = 4300    'Concepto
   grd_Listad.ColWidth(1) = 1150    'Importe
   grd_Listad.ColWidth(2) = 0       'Codigo gasto
   grd_Listad.ColWidth(3) = 1250    'Fecha de Pago
   grd_Listad.ColWidth(4) = 1450    'Numero de documento
   grd_Listad.ColWidth(5) = 1200    'Monto Pagado
   
   grd_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad.ColAlignment(1) = flexAlignRightCenter
   grd_Listad.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad.ColAlignment(4) = flexAlignCenterCenter
   grd_Listad.ColAlignment(5) = flexAlignRightCenter
     
   moddat_g_str_FecIni = ""
   moddat_g_str_FecFin = ""
   moddat_g_str_CodAno = 0
   moddat_g_str_CodMes = 0
 
   Call moddat_gf_ConsultaPerMesActivo("000001", 1, moddat_g_str_FecIni, moddat_g_str_FecFin, moddat_g_str_CodMes, moddat_g_str_CodAno)
   
   Call gs_Carga_ParSubPrd_Combo(cmb_GasAdm, "007")
    
   ipp_FecPag.Text = Format(CDate(date), "dd/mm/yyyy")
   ipp_FecPag.DateMin = Format(CDate(moddat_g_str_FecIni), "yyyymmdd")
   ipp_FecPag.DateMax = Format(CDate(moddat_g_str_FecFin), "yyyymmdd")
   
   ipp_FecGasCie.Text = Format(CDate(date), "dd/mm/yyyy")
   ipp_FecGasCie.DateMin = Format(CDate(moddat_g_str_FecIni), "yyyymmdd")
   ipp_FecGasCie.DateMax = Format(CDate(moddat_g_str_FecFin), "yyyymmdd")
End Sub
Private Sub gs_Carga_ParSubPrd_Combo(p_Combo As ComboBox, ByVal p_CodGrp As String)
   p_Combo.Clear
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT SUBSTR(PARPRD_CODITE,1,2) AS CODGAS, PARPRD_DESCRI "
   g_str_Parame = g_str_Parame & "   FROM CRE_PARPRD "
   g_str_Parame = g_str_Parame & "  WHERE PARPRD_CODGRP = '" & p_CodGrp & "' "
   g_str_Parame = g_str_Parame & "    AND PARPRD_CODITE IN ('251','261','241') "
   g_str_Parame = g_str_Parame & "    AND SUBSTR(PARPRD_CODITE,3,1) = '1' "
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
Private Sub fs_Buscar()
   
   pnl_NumSol.Caption = gf_Formato_NumSol(moddat_g_str_NumSol)
   pnl_DocIde.Caption = CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc
   pnl_NomCli.Caption = moddat_g_str_NomCli
   pnl_Moneda.Caption = moddat_gf_Consulta_ParDes("204", CStr(moddat_g_int_TipMon))
   pnl_FechaPago.Caption = moddat_g_str_FecIng
   pnl_ImpSal.Caption = Format(moddat_g_dbl_IngDec, "###,###,###,##0.00") & " "
   pnl_Estado.Caption = moddat_gf_Consulta_ParDes("020", CStr(moddat_g_int_Situac))
   
   Call gs_LimpiaGrid(grd_Listad)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT A.GASADM_NUMSOL, A.GASADM_CODGAS, D.PARPRD_DESCRI, A.GASADM_IMPORT, "
   g_str_Parame = g_str_Parame & "       A.GASADM_FECPAGPRV, A.GASADM_TIPPAGPRV, A.GASADM_NUMDOCPRV,  "
   g_str_Parame = g_str_Parame & "       C.PARDES_DESCRI, A.GASADM_MTOPAGPRV, A.GASADM_FECENTPRV "
   g_str_Parame = g_str_Parame & "  FROM TRA_GASADM A "
   g_str_Parame = g_str_Parame & " INNER JOIN CRE_SOLMAE B ON B.SOLMAE_NUMERO = A.GASADM_NUMSOL "
   g_str_Parame = g_str_Parame & " INNER JOIN CRE_PARPRD D ON D.PARPRD_CODPRD = B.SOLMAE_CODPRD "
   g_str_Parame = g_str_Parame & "                        AND D.PARPRD_CODSUB = B.SOLMAE_CODSUB AND D.PARPRD_CODGRP = '007'"
   g_str_Parame = g_str_Parame & "                        AND D.PARPRD_CODITE = GASADM_CODGAS||GASADM_TIPMON"
   g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES C ON C.PARDES_CODGRP = 371 AND C.PARDES_CODITE = A.GASADM_TIPPAGPRV"
   g_str_Parame = g_str_Parame & " WHERE A.GASADM_NUMSOL = '" & moddat_g_str_NumSol & "'"
   g_str_Parame = g_str_Parame & " ORDER BY A.GASADM_CODGAS "
    
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_Listad.Redraw = False
      
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         
         grd_Listad.Col = 0
         grd_Listad.Text = Trim(g_rst_Princi!PARPRD_DESCRI)
         
         grd_Listad.Col = 1
         grd_Listad.Text = Format(g_rst_Princi!GASADM_IMPORT, "###,###,##0.00")
         
         grd_Listad.Col = 2
         grd_Listad.Text = g_rst_Princi!GASADM_CODGAS
         
         grd_Listad.Col = 3
         If Not IsNull(g_rst_Princi!GASADM_FECPAGPRV) Then
            grd_Listad.Text = gf_FormatoFecha(g_rst_Princi!GASADM_FECPAGPRV)
         Else
            grd_Listad.Text = ""
         End If

         grd_Listad.Col = 4
         If Not IsNull(g_rst_Princi!GASADM_NUMDOCPRV) Then
            grd_Listad.Text = CStr(g_rst_Princi!GASADM_NUMDOCPRV)
         Else
            grd_Listad.Text = ""
         End If
         
         grd_Listad.Col = 5
         If Not IsNull(g_rst_Princi!GASADM_MTOPAGPRV) Then
            grd_Listad.Text = Format(g_rst_Princi!GASADM_MTOPAGPRV, "###,###,##0.00")
         Else
            grd_Listad.Text = ""
         End If
                  
         g_rst_Princi.MoveNext
      Loop
         
      grd_Listad.Redraw = True
      Call gs_UbiIniGrid(grd_Listad)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Habilitar(ByVal p_Habilitado As Boolean)
   
   pnl_Concepto.Caption = ""
   pnl_Importe.Caption = ""
   ipp_FecPag.Text = date
   txt_NumDoc.Text = ""
   ipp_MtoPag.Text = "0.00"
   
   cmd_NueIte.Enabled = Not p_Habilitado
   cmd_Cancel.Enabled = p_Habilitado
   cmd_Grabar.Enabled = p_Habilitado
   cmd_Editar.Enabled = Not p_Habilitado
   cmd_Borrar.Enabled = Not p_Habilitado
   grd_Listad.Enabled = Not p_Habilitado
   ipp_FecPag.Enabled = p_Habilitado
   txt_NumDoc.Enabled = p_Habilitado
   ipp_MtoPag.Enabled = p_Habilitado
   
   cmb_GasAdm.Enabled = p_Habilitado
   ipp_Import.Enabled = p_Habilitado
   txt_NumOper.Enabled = p_Habilitado
   ipp_FecGasCie.Enabled = p_Habilitado
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 1 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

Private Sub grd_Listad_Click()
   Call grd_Listad_SelChange
End Sub

Private Sub grd_Listad_DblClick()
Dim r_dbl_importe   As Double
Dim r_dbl_ImpPag     As Double

moddat_g_int_FlgGrb = 2
   r_dbl_importe = CDbl(grd_Listad.TextMatrix(grd_Listad.Row, 1))
   
   If Me.grd_Listad.TextMatrix(grd_Listad.Row, 5) = "" Then
      r_dbl_ImpPag = 0
   Else
      r_dbl_ImpPag = CDbl(Me.grd_Listad.TextMatrix(grd_Listad.Row, 5))
   End If
   
   If CDbl(r_dbl_importe) - CDbl(r_dbl_ImpPag) = 0 Then
      MsgBox "Gasto no puede modificarse.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
        
   If grd_Listad.Rows > 0 Then
      grd_Listad.Col = 2
      l_int_CodGas = grd_Listad.Text
      
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT A.GASADM_CODGAS, D.PARPRD_DESCRI, A.GASADM_IMPORT, "
      g_str_Parame = g_str_Parame & "       A.GASADM_FECPAGPRV, A.GASADM_TIPPAGPRV, C.PARDES_DESCRI, "
      g_str_Parame = g_str_Parame & "       A.GASADM_NUMDOCPRV, A.GASADM_MTOPAGPRV, A.GASADM_FECENTPRV "
      g_str_Parame = g_str_Parame & "  FROM TRA_GASADM A "
      g_str_Parame = g_str_Parame & " INNER JOIN CRE_SOLMAE B ON B.SOLMAE_NUMERO = A.GASADM_NUMSOL "
      g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES C ON C.PARDES_CODGRP = 371 AND C.PARDES_CODITE = A.GASADM_TIPPAGPRV "
      g_str_Parame = g_str_Parame & " INNER JOIN CRE_PARPRD D ON D.PARPRD_CODPRD = B.SOLMAE_CODPRD AND D.PARPRD_CODSUB = B.SOLMAE_CODSUB "
      g_str_Parame = g_str_Parame & "                        AND D.PARPRD_CODGRP = '007' AND D.PARPRD_CODITE = GASADM_CODGAS||GASADM_TIPMON "
      g_str_Parame = g_str_Parame & " WHERE GASADM_NUMSOL = '" & moddat_g_str_NumSol & "' "
      g_str_Parame = g_str_Parame & "   AND GASADM_CODGAS = '" & l_int_CodGas & "' "
        
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         g_rst_Princi.MoveFirst
         Call fs_Habilitar(True)
         
         grd_Listad.Col = 0
         pnl_Concepto.Caption = Trim(grd_Listad.Text)
         
         grd_Listad.Col = 1
         pnl_Importe.Caption = grd_Listad.Text & " "
         
         If Not IsNull(g_rst_Princi!GASADM_FECPAGPRV) Then
            ipp_FecPag.Text = gf_FormatoFecha(g_rst_Princi!GASADM_FECPAGPRV)
         Else
            ipp_FecPag.Text = date
         End If
                 
         If Not IsNull(g_rst_Princi!GASADM_NUMDOCPRV) Then
            txt_NumDoc.Text = CStr(g_rst_Princi!GASADM_NUMDOCPRV)
            txt_NumOper.Text = CStr(g_rst_Princi!GASADM_NUMDOCPRV)
         Else
            txt_NumDoc.Text = ""
            txt_NumOper.Text = ""
         End If
         
         If Not IsNull(g_rst_Princi!GASADM_MTOPAGPRV) Then
            ipp_MtoPag.Text = Format(g_rst_Princi!GASADM_MTOPAGPRV, "###,###,##0.00")
         Else
            ipp_MtoPag.Text = "0.00"
         End If
                 
         Call gs_UbicaGrid(grd_Listad, grd_Listad.Row)
         cmb_GasAdm.Enabled = False
         ipp_Import.Enabled = False
         txt_NumOper.Enabled = False
         ipp_FecGasCie.Enabled = False
         ipp_FecPag.SetFocus
      End If
   End If
End Sub

Private Sub ipp_FecPag_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NumDoc)
   End If
End Sub
Private Sub ipp_FecPag_InvalidData(NextWnd As Long)
    If CDate(ipp_FecPag.Text) < CDate(moddat_g_str_FecIni) Then
      ipp_FecPag.Text = moddat_g_str_FecIni
   ElseIf CDate(ipp_FecPag.Text) > CDate(moddat_g_str_FecFin) Then
      ipp_FecPag.Text = moddat_g_str_FecFin
   End If
End Sub
Private Sub cmb_TipPag_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NumDoc)
   End If
End Sub

Private Sub ipp_Import_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NumOper)
   End If
End Sub

Private Sub txt_NumDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_MtoPag)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "- ()?¿)(/&%$·#@_.,;:")
   End If
End Sub

Private Sub ipp_MtoPag_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   End If
End Sub

Private Sub ipp_FecEnt_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   End If
End Sub

Private Sub txt_NumOper_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
    Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "- ()?¿)(/&%$·#@_.,;:")
   End If
End Sub
