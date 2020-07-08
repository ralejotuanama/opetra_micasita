VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_Ges_CreHip_15 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   5565
   ClientLeft      =   3255
   ClientTop       =   3165
   ClientWidth     =   13230
   Icon            =   "OpeTra_frm_150.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   13230
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   5565
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   13245
      _Version        =   65536
      _ExtentX        =   23363
      _ExtentY        =   9816
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
      Begin Threed.SSPanel SSPanel39 
         Height          =   645
         Left            =   30
         TabIndex        =   9
         Top             =   750
         Width           =   13155
         _Version        =   65536
         _ExtentX        =   23204
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
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_150.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   30
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   12540
            Picture         =   "OpeTra_frm_150.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   11
         Top             =   30
         Width           =   13155
         _Version        =   65536
         _ExtentX        =   23204
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
            Left            =   720
            TabIndex        =   12
            Top             =   30
            Width           =   5685
            _Version        =   65536
            _ExtentX        =   10028
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Gestión de Crédito Hipotecario"
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
            Height          =   315
            Left            =   720
            TabIndex        =   13
            Top             =   330
            Width           =   5505
            _Version        =   65536
            _ExtentX        =   9710
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Exnoración de Cargos por Cobranza Morosa"
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
            Picture         =   "OpeTra_frm_150.frx":0890
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   30
         TabIndex        =   14
         Top             =   1440
         Width           =   13155
         _Version        =   65536
         _ExtentX        =   23204
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
         Begin Threed.SSPanel pnl_NumOpe 
            Height          =   315
            Left            =   2070
            TabIndex        =   15
            Top             =   60
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
         Begin Threed.SSPanel pnl_NomCli 
            Height          =   315
            Left            =   2070
            TabIndex        =   16
            Top             =   390
            Width           =   11025
            _Version        =   65536
            _ExtentX        =   19447
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "1-07521154 / IKEHARA PUNK MIGUEL ANGEL"
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
         Begin VB.Label Label12 
            Caption         =   "Nro. Operación:"
            Height          =   315
            Left            =   60
            TabIndex        =   18
            Top             =   60
            Width           =   1245
         End
         Begin VB.Label Label5 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   60
            TabIndex        =   17
            Top             =   390
            Width           =   1395
         End
      End
      Begin Threed.SSPanel SSPanel14 
         Height          =   1515
         Left            =   30
         TabIndex        =   19
         Top             =   2250
         Width           =   13155
         _Version        =   65536
         _ExtentX        =   23204
         _ExtentY        =   2672
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
         Begin MSFlexGridLib.MSFlexGrid grd_InfCuo 
            Height          =   1395
            Left            =   60
            TabIndex        =   21
            Top             =   60
            Width           =   13035
            _ExtentX        =   22992
            _ExtentY        =   2461
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
         Height          =   825
         Left            =   30
         TabIndex        =   20
         Top             =   4680
         Width           =   13155
         _Version        =   65536
         _ExtentX        =   23204
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
         Begin EditLib.fpDoubleSingle ipp_IntMor 
            Height          =   315
            Left            =   2070
            TabIndex        =   2
            Top             =   90
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
            MaxValue        =   "9999"
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
         Begin EditLib.fpDoubleSingle ipp_IntCom 
            Height          =   315
            Left            =   6750
            TabIndex        =   3
            Top             =   90
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
         Begin EditLib.fpDoubleSingle ipp_GasCob 
            Height          =   315
            Left            =   11430
            TabIndex        =   4
            Top             =   90
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
         Begin EditLib.fpDoubleSingle ipp_OtrGas 
            Height          =   315
            Left            =   2070
            TabIndex        =   5
            Top             =   420
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
         Begin EditLib.fpDoubleSingle ipp_CapPBP 
            Height          =   315
            Left            =   6750
            TabIndex        =   6
            Top             =   420
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
         Begin EditLib.fpDoubleSingle ipp_IntPBP 
            Height          =   315
            Left            =   11430
            TabIndex        =   7
            Top             =   420
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
         Begin VB.Label lbl_SimMon 
            Alignment       =   1  'Right Justify
            Caption         =   "US$"
            Height          =   285
            Index           =   5
            Left            =   10830
            TabIndex        =   34
            Top             =   420
            Width           =   465
         End
         Begin VB.Label Label13 
            Caption         =   "Interés PBP:"
            Height          =   285
            Left            =   9060
            TabIndex        =   33
            Top             =   420
            Width           =   1695
         End
         Begin VB.Label lbl_SimMon 
            Alignment       =   1  'Right Justify
            Caption         =   "US$"
            Height          =   285
            Index           =   3
            Left            =   6150
            TabIndex        =   32
            Top             =   420
            Width           =   465
         End
         Begin VB.Label Label10 
            Caption         =   "Capital PBP:"
            Height          =   285
            Left            =   4380
            TabIndex        =   31
            Top             =   420
            Width           =   1695
         End
         Begin VB.Label lbl_SimMon 
            Alignment       =   1  'Right Justify
            Caption         =   "US$"
            Height          =   285
            Index           =   1
            Left            =   1470
            TabIndex        =   29
            Top             =   420
            Width           =   465
         End
         Begin VB.Label Label8 
            Caption         =   "Otros Gastos:"
            Height          =   285
            Left            =   90
            TabIndex        =   28
            Top             =   420
            Width           =   1695
         End
         Begin VB.Label lbl_SimMon 
            Alignment       =   1  'Right Justify
            Caption         =   "US$"
            Height          =   285
            Index           =   4
            Left            =   10830
            TabIndex        =   27
            Top             =   90
            Width           =   465
         End
         Begin VB.Label Label6 
            Caption         =   "Gastos de Cobranza:"
            Height          =   285
            Left            =   9060
            TabIndex        =   26
            Top             =   90
            Width           =   1695
         End
         Begin VB.Label lbl_SimMon 
            Alignment       =   1  'Right Justify
            Caption         =   "US$"
            Height          =   285
            Index           =   2
            Left            =   6150
            TabIndex        =   25
            Top             =   90
            Width           =   465
         End
         Begin VB.Label Label3 
            Caption         =   "Interés Compensatorio:"
            Height          =   285
            Left            =   4380
            TabIndex        =   24
            Top             =   90
            Width           =   1695
         End
         Begin VB.Label lbl_SimMon 
            Alignment       =   1  'Right Justify
            Caption         =   "US$"
            Height          =   285
            Index           =   0
            Left            =   1470
            TabIndex        =   23
            Top             =   90
            Width           =   465
         End
         Begin VB.Label Label41 
            Caption         =   "Interés Moratorio:"
            Height          =   285
            Left            =   90
            TabIndex        =   22
            Top             =   90
            Width           =   1365
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   825
         Left            =   30
         TabIndex        =   35
         Top             =   3810
         Width           =   13155
         _Version        =   65536
         _ExtentX        =   23204
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
         Begin VB.ComboBox cmb_NivAut 
            Height          =   315
            Left            =   2070
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   420
            Width           =   11025
         End
         Begin VB.ComboBox cmb_MotExo 
            Height          =   315
            Left            =   2070
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   11025
         End
         Begin VB.Label Label26 
            Caption         =   "Motivo Exoneración:"
            Height          =   285
            Left            =   60
            TabIndex        =   37
            Top             =   60
            Width           =   1515
         End
         Begin VB.Label Label20 
            Caption         =   "Autorizado por:"
            Height          =   285
            Left            =   60
            TabIndex        =   36
            Top             =   420
            Width           =   1305
         End
      End
   End
End
Attribute VB_Name = "frm_Ges_CreHip_15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_dbl_IntMor     As Double
Dim l_dbl_IntCom     As Double
Dim l_dbl_GasCob     As Double
Dim l_dbl_OtrGas     As Double
Dim l_dbl_CapPBP     As Double
Dim l_dbl_IntPBP     As Double

Private Sub cmb_MotExo_Click()
   Call gs_SetFocus(cmb_NivAut)
End Sub

Private Sub cmb_MotExo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_MotExo_Click
   End If
End Sub

Private Sub cmb_NivAut_Click()
   Call gs_SetFocus(ipp_IntMor)
End Sub

Private Sub cmb_NivAut_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_NivAut_Click
   End If
End Sub

Private Sub cmd_Grabar_Click()
   Dim r_dbl_TotOrg     As Double
   Dim r_dbl_TotNue     As Double
   Dim r_int_NumExo     As Integer

   If cmb_MotExo.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Motivo de Exoneración.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_MotExo)
      Exit Sub
   End If
   
   If cmb_NivAut.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Nivel de Autorización.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_NivAut)
      Exit Sub
   End If
   
   r_dbl_TotOrg = l_dbl_IntMor + l_dbl_IntCom + l_dbl_GasCob + l_dbl_OtrGas + l_dbl_CapPBP + l_dbl_IntPBP
   r_dbl_TotNue = CDbl(ipp_IntMor.Value) + CDbl(ipp_IntCom.Value) + CDbl(ipp_GasCob.Value) + CDbl(ipp_OtrGas.Value) + CDbl(ipp_CapPBP.Value) + CDbl(ipp_IntPBP.Value)
   
   If r_dbl_TotNue > r_dbl_TotOrg Then
      If MsgBox("El monto de Exoneración es mayor al monto anterior. ¿Está seguro de que son los datos correctos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
         Exit Sub
      End If
   End If
   
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   'Obteniendo Información de Cuota
   g_str_Parame = "SELECT MAX(HIPEXO_NUMEXO) AS NUMERO FROM CRE_HIPEXO WHERE "
   g_str_Parame = g_str_Parame & "HIPEXO_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If IsNull(g_rst_Princi!NUMERO) Then
      r_int_NumExo = 1
   Else
      r_int_NumExo = g_rst_Princi!NUMERO + 1
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      Screen.MousePointer = 11
      Call moddat_gs_FecSis
      
      g_str_Parame = "USP_CRE_HIPEXO ("
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "', "
      g_str_Parame = g_str_Parame & CStr(r_int_NumExo) & ", "
      g_str_Parame = g_str_Parame & CStr(cmb_MotExo.ItemData(cmb_MotExo.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & CStr(cmb_NivAut.ItemData(cmb_NivAut.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & CStr(moddat_g_int_NumCuo) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_IntMor) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_IntCom) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_GasCob) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_OtrGas) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_CapPBP) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_IntPBP) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_IntMor.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_IntCom.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_GasCob.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_OtrGas.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_CapPBP.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_IntPBP.Value) & ", "
      
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_int_FlgGrb) & ") "
      
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

   moddat_g_int_FlgAct = 2
   Unload Me
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   pnl_NumOpe.Caption = ""
   pnl_NomCli.Caption = ""
   pnl_NumOpe.Caption = gf_Formato_NumOpe(moddat_g_str_NumOpe)
   pnl_NomCli.Caption = CStr(moddat_g_int_TipDoc) & " - " & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   
   Call fs_Inicia
   Call fs_Buscar_DatCuo
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   grd_InfCuo.ColWidth(0) = 2650
   grd_InfCuo.ColWidth(1) = 10000
   grd_InfCuo.ColAlignment(0) = flexAlignLeftCenter
   grd_InfCuo.ColAlignment(1) = flexAlignLeftCenter

   Call moddat_gs_Carga_LisIte_Combo(cmb_MotExo, 1, "252")
   Call moddat_gs_Carga_LisIte_Combo(cmb_NivAut, 1, "253")

   If modgen_g_int_TipUsu = 18000 Then
      ipp_IntMor.Enabled = True
      ipp_IntCom.Enabled = True
      ipp_GasCob.Enabled = True
      ipp_OtrGas.Enabled = True
   Else
      ipp_IntMor.Enabled = False
      ipp_IntCom.Enabled = False
      ipp_GasCob.Enabled = False
      ipp_OtrGas.Enabled = False
   End If
End Sub

Private Sub fs_Buscar_DatCuo()
   Dim r_int_Contad     As Integer
   
   Call gs_LimpiaGrid(grd_InfCuo)
   
   'Obteniendo Información de Cuota
   g_str_Parame = "SELECT * FROM CRE_HIPCUO WHERE "
   g_str_Parame = g_str_Parame & "HIPCUO_NUMOPE = '" & moddat_g_str_NumOpe & "' AND "
   g_str_Parame = g_str_Parame & "HIPCUO_NUMCUO = " & CStr(moddat_g_int_NumCuo) & " AND "
   g_str_Parame = g_str_Parame & "HIPCUO_TIPCRO = 1 "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      grd_InfCuo.Rows = grd_InfCuo.Rows + 1
      grd_InfCuo.Row = grd_InfCuo.Rows - 1
      grd_InfCuo.Col = 0
      grd_InfCuo.Text = "Número de Cuota"
      
      grd_InfCuo.Col = 1
      grd_InfCuo.Text = CStr(moddat_g_int_NumCuo)
      
      grd_InfCuo.Rows = grd_InfCuo.Rows + 1
      grd_InfCuo.Row = grd_InfCuo.Rows - 1
      grd_InfCuo.Col = 0
      grd_InfCuo.Text = "Fecha de Vencimiento"
      
      grd_InfCuo.Col = 1
      grd_InfCuo.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECVCT))
      
      'Si Situación es No-Pagado
      If g_rst_Princi!HIPCUO_SITUAC = 2 Then
         If CDate(gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECVCT))) < CDate(moddat_g_str_FecSis) Then
            grd_InfCuo.Rows = grd_InfCuo.Rows + 1
            grd_InfCuo.Row = grd_InfCuo.Rows - 1
            grd_InfCuo.Col = 0
            grd_InfCuo.Text = "Situación"
            
            grd_InfCuo.Col = 1
            grd_InfCuo.Text = "VENCIDA"
            
            grd_InfCuo.Rows = grd_InfCuo.Rows + 1
            grd_InfCuo.Row = grd_InfCuo.Rows - 1
            grd_InfCuo.Col = 0
            grd_InfCuo.Text = "Días de Atraso"
            
            grd_InfCuo.Col = 1
            grd_InfCuo.Text = CStr(CInt(CDate(moddat_g_str_FecSis) - CDate(gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECVCT)))))
         Else
            grd_InfCuo.Rows = grd_InfCuo.Rows + 1
            grd_InfCuo.Row = grd_InfCuo.Rows - 1
            grd_InfCuo.Col = 0
            grd_InfCuo.Text = "Situación"
            
            grd_InfCuo.Col = 1
            grd_InfCuo.Text = "POR VENCER"
            
            grd_InfCuo.Rows = grd_InfCuo.Rows + 1
            grd_InfCuo.Row = grd_InfCuo.Rows - 1
            grd_InfCuo.Col = 0
            grd_InfCuo.Text = "Días de Atraso"
            
            grd_InfCuo.Col = 1
            grd_InfCuo.Text = "0"
         End If
      End If
      
      For r_int_Contad = 0 To 5
         lbl_SimMon(r_int_Contad).Caption = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon))
      Next r_int_Contad
      
      ipp_IntMor.Value = g_rst_Princi!HIPCUO_INTMOR
      ipp_IntCom.Value = g_rst_Princi!HIPCUO_INTCOM
      ipp_GasCob.Value = g_rst_Princi!HIPCUO_GASCOB
      ipp_OtrGas.Value = g_rst_Princi!HIPCUO_OTRGAS
      ipp_CapPBP.Value = g_rst_Princi!HIPCUO_CAPBBP
      ipp_IntPBP.Value = g_rst_Princi!HIPCUO_INTBBP
      l_dbl_IntMor = g_rst_Princi!HIPCUO_INTMOR
      l_dbl_IntCom = g_rst_Princi!HIPCUO_INTCOM
      l_dbl_GasCob = g_rst_Princi!HIPCUO_GASCOB
      l_dbl_OtrGas = g_rst_Princi!HIPCUO_OTRGAS
      l_dbl_CapPBP = g_rst_Princi!HIPCUO_CAPBBP
      l_dbl_IntPBP = g_rst_Princi!HIPCUO_INTBBP
      
      Call gs_UbiIniGrid(grd_InfCuo)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub ipp_IntMor_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_IntCom)
   End If
End Sub

Private Sub ipp_IntCom_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_GasCob)
   End If
End Sub

Private Sub ipp_GasCob_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_OtrGas)
   End If
End Sub

Private Sub ipp_OtrGas_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_CapPBP)
   End If
End Sub

Private Sub ipp_CapPBP_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_IntPBP)
   End If
End Sub

Private Sub ipp_IntPBP_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   End If
End Sub
