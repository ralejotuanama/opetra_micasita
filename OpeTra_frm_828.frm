VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_Con_Cuadre_04 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11265
   Icon            =   "OpeTra_frm_828.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   11265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel111 
      Height          =   7815
      Left            =   -90
      TabIndex        =   23
      Top             =   0
      Width           =   11415
      _Version        =   65536
      _ExtentX        =   20135
      _ExtentY        =   13785
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
      Begin Threed.SSPanel SSPanel4 
         Height          =   2235
         Left            =   150
         TabIndex        =   24
         Top             =   1470
         Width           =   11160
         _Version        =   65536
         _ExtentX        =   19685
         _ExtentY        =   3942
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
            Height          =   1815
            Left            =   60
            TabIndex        =   25
            Top             =   330
            Width           =   11050
            _ExtentX        =   19500
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
         Begin VB.Label Label2 
            Caption         =   "Datos del Crédito"
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
            TabIndex        =   26
            Top             =   90
            Width           =   1875
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   675
         Left            =   150
         TabIndex        =   27
         Top             =   60
         Width           =   11160
         _Version        =   65536
         _ExtentX        =   19685
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
            Height          =   600
            Left            =   690
            TabIndex        =   28
            Top             =   30
            Width           =   4905
            _Version        =   65536
            _ExtentX        =   8652
            _ExtentY        =   1058
            _StockProps     =   15
            Caption         =   "Cuadre de Operaciones Adjudicadas y Recuperados"
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
            Left            =   90
            Picture         =   "OpeTra_frm_828.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   645
         Left            =   150
         TabIndex        =   29
         Top             =   780
         Width           =   11160
         _Version        =   65536
         _ExtentX        =   19685
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
         Begin VB.CommandButton cmd_ImpCro 
            Enabled         =   0   'False
            Height          =   585
            Left            =   8730
            Picture         =   "OpeTra_frm_828.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Consulta Cronograma de Pagos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_VerPag 
            Enabled         =   0   'False
            Height          =   585
            Left            =   8130
            Picture         =   "OpeTra_frm_828.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Consulta Pagos del Cliente"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Grabar 
            Enabled         =   0   'False
            Height          =   585
            Left            =   9330
            Picture         =   "OpeTra_frm_828.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10530
            Picture         =   "OpeTra_frm_828.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   9930
            Picture         =   "OpeTra_frm_828.frx":11AE
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Limpiar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.ComboBox cmb_PerMes 
            Height          =   315
            Left            =   3300
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   180
            Width           =   1725
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   585
            Left            =   7530
            Picture         =   "OpeTra_frm_828.frx":14B8
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Buscar Operación"
            Top             =   30
            Width           =   585
         End
         Begin MSMask.MaskEdBox msk_NumOpe 
            Height          =   315
            Left            =   1320
            TabIndex        =   0
            Top             =   180
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   12
            Mask            =   "###-##-#####"
            PromptChar      =   " "
         End
         Begin EditLib.fpDoubleSingle ipp_PerAno 
            Height          =   315
            Left            =   5640
            TabIndex        =   2
            Top             =   180
            Width           =   825
            _Version        =   196608
            _ExtentX        =   1455
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
            ButtonStyle     =   1
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
            Text            =   "0"
            DecimalPlaces   =   0
            DecimalPoint    =   "."
            FixedPoint      =   0   'False
            LeadZero        =   0
            MaxValue        =   "9999"
            MinValue        =   "1900"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ","
            UseSeparator    =   0   'False
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
         Begin VB.Label lbl_Numero 
            Caption         =   "Nro. Operación:"
            Height          =   285
            Left            =   120
            TabIndex        =   32
            Top             =   210
            Width           =   1275
         End
         Begin VB.Label Label50 
            Caption         =   "Mes:"
            Height          =   315
            Left            =   2880
            TabIndex        =   31
            Top             =   210
            Width           =   555
         End
         Begin VB.Label Label51 
            Caption         =   "Año:"
            Height          =   255
            Left            =   5250
            TabIndex        =   30
            Top             =   240
            Width           =   405
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   1905
         Left            =   150
         TabIndex        =   33
         Top             =   3750
         Width           =   11160
         _Version        =   65536
         _ExtentX        =   19685
         _ExtentY        =   3360
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Begin VB.ComboBox cmb_Estado 
            Height          =   315
            Left            =   2400
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   1460
            Width           =   2475
         End
         Begin VB.ComboBox cmb_TipOrigen 
            Height          =   315
            Left            =   2400
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   450
            Width           =   2475
         End
         Begin EditLib.fpDoubleSingle ipp_IntDeu 
            Height          =   315
            Left            =   7980
            TabIndex        =   14
            Top             =   795
            Width           =   1320
            _Version        =   196608
            _ExtentX        =   2328
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
         Begin EditLib.fpDoubleSingle ipp_CapDeu 
            Height          =   315
            Left            =   7980
            TabIndex        =   13
            Top             =   450
            Width           =   1320
            _Version        =   196608
            _ExtentX        =   2328
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
         Begin Threed.SSPanel pnl_TotDeuda 
            Height          =   315
            Left            =   7980
            TabIndex        =   15
            Top             =   1140
            Width           =   1320
            _Version        =   65536
            _ExtentX        =   2328
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
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
            RoundedCorners  =   0   'False
            Font3D          =   2
            Alignment       =   4
         End
         Begin EditLib.fpDateTime ipp_FecAdj 
            Height          =   315
            Left            =   2400
            TabIndex        =   10
            Top             =   795
            Width           =   1320
            _Version        =   196608
            _ExtentX        =   2328
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
         Begin EditLib.fpDateTime ipp_FecVcto 
            Height          =   315
            Left            =   2400
            TabIndex        =   11
            Top             =   1130
            Width           =   1320
            _Version        =   196608
            _ExtentX        =   2328
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
            AllowNull       =   -1  'True
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
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Estado:"
            Height          =   195
            Left            =   120
            TabIndex        =   57
            Top             =   1500
            Width           =   540
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "="
            Height          =   195
            Left            =   9390
            TabIndex        =   55
            Top             =   1170
            Width           =   90
         End
         Begin VB.Line Line6 
            X1              =   5130
            X2              =   5130
            Y1              =   450
            Y2              =   1750
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "+"
            Height          =   195
            Left            =   9390
            TabIndex        =   50
            Top             =   480
            Width           =   90
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Vcto. Prorroga(F):"
            Height          =   195
            Left            =   120
            TabIndex        =   49
            Top             =   1170
            Width           =   1740
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Valor Total que Cancela (D):"
            Height          =   195
            Left            =   5430
            TabIndex        =   39
            Top             =   1170
            Width           =   2010
         End
         Begin VB.Label Label29 
            Caption         =   "Adjudicación, Dación en Pago o Recuperación"
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
            Left            =   120
            TabIndex        =   38
            Top             =   120
            UseMnemonic     =   0   'False
            Width           =   4485
         End
         Begin VB.Label Label26 
            Caption         =   "Capital Deuda que Cancela (E):"
            Height          =   195
            Left            =   5430
            TabIndex        =   37
            Top             =   510
            Width           =   2310
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Origen (B):"
            Height          =   195
            Left            =   120
            TabIndex        =   36
            Top             =   510
            Width           =   1110
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Origen:"
            Height          =   195
            Left            =   120
            TabIndex        =   35
            Top             =   840
            Width           =   1005
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Interés y Otros que Cancela (E):"
            Height          =   195
            Left            =   5430
            TabIndex        =   34
            Top             =   840
            Width           =   2250
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   2025
         Left            =   150
         TabIndex        =   40
         Top             =   5700
         Width           =   11160
         _Version        =   65536
         _ExtentX        =   19685
         _ExtentY        =   3572
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Begin EditLib.fpDoubleSingle ipp_PrvCon 
            Height          =   315
            Left            =   2400
            TabIndex        =   17
            Top             =   795
            Width           =   1320
            _Version        =   196608
            _ExtentX        =   2328
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
         Begin Threed.SSPanel pnl_PrvTot 
            Height          =   315
            Left            =   2400
            TabIndex        =   19
            Top             =   1485
            Width           =   1320
            _Version        =   65536
            _ExtentX        =   2328
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
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
            RoundedCorners  =   0   'False
            Font3D          =   2
            Alignment       =   4
         End
         Begin EditLib.fpDateTime ipp_FecRea 
            Height          =   315
            Left            =   7980
            TabIndex        =   20
            Top             =   450
            Width           =   1320
            _Version        =   196608
            _ExtentX        =   2328
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
            AllowNull       =   -1  'True
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
         Begin EditLib.fpDoubleSingle ipp_PrvIni 
            Height          =   315
            Left            =   2400
            TabIndex        =   16
            Top             =   450
            Width           =   1320
            _Version        =   196608
            _ExtentX        =   2328
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
         Begin EditLib.fpDoubleSingle ipp_PrvDsv 
            Height          =   315
            Left            =   2400
            TabIndex        =   18
            Top             =   1140
            Width           =   1320
            _Version        =   196608
            _ExtentX        =   2328
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
         Begin EditLib.fpDoubleSingle ipp_ValRea 
            Height          =   315
            Left            =   7980
            TabIndex        =   21
            Top             =   795
            Width           =   1320
            _Version        =   196608
            _ExtentX        =   2328
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
         Begin EditLib.fpDoubleSingle ipp_ValLib 
            Height          =   315
            Left            =   7980
            TabIndex        =   22
            Top             =   1140
            Width           =   1320
            _Version        =   196608
            _ExtentX        =   2328
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
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "="
            Height          =   195
            Left            =   3780
            TabIndex        =   56
            Top             =   1500
            Width           =   90
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "+ (Inicial)"
            Height          =   195
            Left            =   3780
            TabIndex        =   54
            Top             =   510
            Width           =   630
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "+ (Posterior)"
            Height          =   195
            Left            =   3780
            TabIndex        =   53
            Top             =   1170
            Width           =   840
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "+ (Posterior)"
            Height          =   195
            Left            =   3780
            TabIndex        =   52
            Top             =   840
            Width           =   840
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "(En Libros)"
            Height          =   195
            Left            =   9390
            TabIndex        =   51
            Top             =   1200
            Width           =   750
         End
         Begin VB.Line Line1 
            X1              =   5160
            X2              =   5160
            Y1              =   450
            Y2              =   1800
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Valor Neto Realización (J):"
            Height          =   195
            Left            =   5400
            TabIndex        =   48
            Top             =   840
            Width           =   1875
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Valor Neto a la Fecha Reporte (K):"
            Height          =   195
            Left            =   5400
            TabIndex        =   47
            Top             =   1170
            Width           =   2445
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Realización (J):"
            Height          =   195
            Left            =   5400
            TabIndex        =   46
            Top             =   510
            Width           =   1575
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Total Provisiones:"
            Height          =   195
            Left            =   120
            TabIndex        =   45
            Top             =   1500
            Width           =   1260
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Prov. Desvalorización (I):"
            Height          =   195
            Left            =   120
            TabIndex        =   44
            Top             =   1170
            Width           =   1770
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Provisión (G):"
            Height          =   195
            Left            =   120
            TabIndex        =   43
            Top             =   510
            Width           =   945
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Prov. Constituida Mensual (H):"
            Height          =   195
            Left            =   120
            TabIndex        =   42
            Top             =   840
            Width           =   2145
         End
         Begin VB.Label Label3 
            Caption         =   "Provisión por Bienes"
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
            Left            =   120
            TabIndex        =   41
            Top             =   120
            UseMnemonic     =   0   'False
            Width           =   2835
         End
      End
   End
End
Attribute VB_Name = "frm_Con_Cuadre_04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   moddat_g_int_FlgGrb = 0
   
   Call fs_Inicia
   Call fs_Limpiar
   Call fs_Validar_Botones(False)
   Call gs_LimpiaGrid(grd_Listad)
   
   Call gs_CentraForm(Me)
   Call gs_SetFocus(msk_NumOpe)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
    'Inicializando Grid de Datos del Crédito
    grd_Listad.ColWidth(0) = 2900
    grd_Listad.ColWidth(1) = 8150
    grd_Listad.ColAlignment(0) = flexAlignLeftCenter
    grd_Listad.ColAlignment(1) = flexAlignLeftCenter
    
    Call moddat_gs_Carga_LisIte_Combo(cmb_PerMes, 1, "033")
    Call moddat_gs_Carga_LisIte_Combo(cmb_Estado, 1, "244")
    
    cmb_TipOrigen.AddItem "ADJUDICACION"
    cmb_TipOrigen.ItemData(cmb_TipOrigen.NewIndex) = CLng(1)
    cmb_TipOrigen.AddItem "DACION EN PAGO"
    cmb_TipOrigen.ItemData(cmb_TipOrigen.NewIndex) = CLng(2)
    cmb_TipOrigen.AddItem "BIENES RECUPERADOS"
    cmb_TipOrigen.ItemData(cmb_TipOrigen.NewIndex) = CLng(3)
End Sub

Private Sub fs_Limpiar()
   msk_NumOpe.Text = ""
   cmb_PerMes.ListIndex = -1
   ipp_PerAno.Text = Year(date)
   Call fs_Limpiar_DatCre
End Sub

Private Sub fs_Limpiar_DatCre()
   cmb_TipOrigen.ListIndex = -1
   ipp_FecAdj.Text = date
   ipp_FecVcto.Text = ""
   ipp_CapDeu.Text = "0.00"
   ipp_IntDeu.Text = "0.00"
   pnl_TotDeuda.Caption = "0.00" & " "
   ipp_PrvIni.Text = "0.00"
   ipp_PrvCon.Text = "0.00"
   ipp_PrvDsv.Text = "0.00"
   pnl_PrvTot.Caption = "0.00" & " "
   ipp_FecRea.Text = ""
   ipp_ValRea.Text = "0.00"
   ipp_ValLib.Text = ""
   cmb_Estado.ListIndex = -1
End Sub

Private Sub cmd_Buscar_Click()
Dim r_str_FecAct As String
   modctb_int_PerAno = 0
   modctb_int_PerMes = 0
   
   Call moddat_gf_ConsultaPerMesActivo("000001", 1, modctb_str_FecIni, modctb_str_FecFin, modctb_int_PerMes, modctb_int_PerAno)

   If Len(Trim(msk_NumOpe.Text)) < 10 Then
       MsgBox "Debe ingresar el Número de Operación.", vbExclamation, modgen_g_con_OpeTra
       msk_NumOpe.Text = ""
       msk_NumOpe.Mask = "###-##-#####"
       Call gs_SetFocus(msk_NumOpe)
       Exit Sub
   End If
   If cmb_PerMes.ListIndex = -1 Then
       MsgBox "Debe seleccionar el Mes.", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(cmb_PerMes)
       Exit Sub
   End If
   If ipp_PerAno.Text < 2010 Then
       MsgBox "Debe ingresar el año correcto.", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(ipp_PerAno)
       Exit Sub
   End If
    
   moddat_g_str_NumOpe = msk_NumOpe.Text
   Screen.MousePointer = 11
   Me.Enabled = False
   moddat_g_int_FlgGrb = 0
        
   Call fs_Buscar_Credito
              
   Me.Enabled = True
   Screen.MousePointer = 0
    
   If moddat_g_int_CntErr = 2 Then
      msk_NumOpe.Text = ""
      msk_NumOpe.Mask = "###-##-#####"
      Call gs_SetFocus(msk_NumOpe)
   Else
      Call gs_SetFocus(cmb_TipOrigen)
   End If
End Sub

Private Sub fs_Buscar_Credito()
Dim r_str_CodPry     As String
Dim r_str_NomPry     As String
Dim r_str_CodBco     As String
Dim r_str_FecAct     As String
    
   moddat_g_int_CntErr = 1
   Call gs_LimpiaGrid(grd_Listad)
   Call fs_Limpiar_DatCre
   
   'Buscando Información del Crédito
   g_str_Parame = " "
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_HIPMAE  "
   g_str_Parame = g_str_Parame & " WHERE HIPMAE_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   g_str_Parame = g_str_Parame & "   AND HIPMAE_SITUAC IN (2,6,9) "
    
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      moddat_g_int_CntErr = 2
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      MsgBox "No existe ninguna Operación registrada con ese Número. ", vbExclamation, modgen_g_con_OpeTra
      Call gs_LimpiaGrid(grd_Listad)
      Call fs_Limpiar
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      moddat_g_int_CntErr = 2
      Exit Sub
   End If
   
   If g_rst_Princi!HIPMAE_SITUAC = 6 Then
      MsgBox "Operación se encuentra transferida.", vbExclamation, modgen_g_con_OpeTra
   End If
   If g_rst_Princi!HIPMAE_SITUAC = 9 Then
      MsgBox "Operación se encuentra cancelada.", vbExclamation, modgen_g_con_OpeTra
   End If
   
   r_str_FecAct = "01/" & Format(modctb_int_PerMes, "00") & "/" & modctb_int_PerAno
   r_str_FecAct = DateAdd("d", -1, r_str_FecAct)
   modctb_int_PerMes = Format(r_str_FecAct, "mm")
   modctb_int_PerAno = Format(r_str_FecAct, "yyyy")
   r_str_FecAct = Format(r_str_FecAct, "yyyymm")
   
   If r_str_FecAct <> ipp_PerAno.Text & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") Then
      Call fs_Validar_Botones(False)
   Else
      Call fs_Validar_Botones(True)
   End If
   
   'cmd_Grabar.Enabled = True
   If Trim(g_rst_Princi!HIPMAE_SITADJ & "") = "" Then  'CANCELADO
      MsgBox "Operación no se encuentra en situación de Adjudicada.", vbExclamation, modgen_g_con_OpeTra
      Call fs_Validar_Botones(False)
   End If
   If g_rst_Princi!HIPMAE_SITADJ = 0 Then  'CANCELADO
      MsgBox "Operación no se encuentra en situación de Adjudicada.", vbExclamation, modgen_g_con_OpeTra
      Call fs_Validar_Botones(False)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
         
   'Información del Crédito
   Call modmip_gs_DatNumOpe(moddat_g_str_NumOpe, grd_Listad)
   
   'Datos del Padron
   Call fs_Buscar_DatCred
End Sub

Private Sub fs_Validar_Botones(ByVal r_bol_FlagEn As Boolean)
   msk_NumOpe.Enabled = Not r_bol_FlagEn
   cmb_PerMes.Enabled = Not r_bol_FlagEn
   ipp_PerAno.Enabled = Not r_bol_FlagEn
   cmd_Buscar.Enabled = Not r_bol_FlagEn
   cmd_VerPag.Enabled = r_bol_FlagEn
   cmd_ImpCro.Enabled = r_bol_FlagEn
   cmd_Grabar.Enabled = r_bol_FlagEn
    
   cmb_TipOrigen.Enabled = r_bol_FlagEn
   cmb_Estado.Enabled = r_bol_FlagEn
   
   ipp_FecAdj.Enabled = r_bol_FlagEn
   ipp_FecVcto.Enabled = r_bol_FlagEn
   ipp_CapDeu.Enabled = r_bol_FlagEn
   ipp_IntDeu.Enabled = r_bol_FlagEn
   ipp_PrvIni.Enabled = r_bol_FlagEn
   ipp_PrvCon.Enabled = r_bol_FlagEn
   ipp_PrvDsv.Enabled = r_bol_FlagEn
   ipp_FecRea.Enabled = r_bol_FlagEn
   ipp_ValRea.Enabled = r_bol_FlagEn
   ipp_ValLib.Enabled = r_bol_FlagEn
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub cmd_Limpia_Click()
    Call gs_LimpiaGrid(grd_Listad)
    Call fs_Validar_Botones(False)
    Call fs_Limpiar
    Call gs_SetFocus(msk_NumOpe)
End Sub

Private Sub cmd_ImpCro_Click()
   modmip_g_int_OrdAct = 2
   frm_Ges_CreHip_07.l_str_PerAno = CStr(ipp_PerAno.Text)
   frm_Ges_CreHip_07.l_str_PerMes = CStr(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   frm_Ges_CreHip_07.Show 1
End Sub

Private Sub cmd_VerPag_Click()
   frm_Ges_CreHip_05.Show 1
End Sub

Private Sub fs_Buscar_DatCred()
Dim r_int_PerAct   As Integer

   'DATOS DE OPERACION
   g_str_Parame = " "
   g_str_Parame = g_str_Parame & " SELECT BIEADJ_TIPORI, BIEADJ_FECADJ, BIEADJ_FECVTO, BIEADJ_CAPDEU, "
   g_str_Parame = g_str_Parame & "        BIEADJ_INTDEU, BIEADJ_PRVINI, BIEADJ_PRVCON, BIEADJ_PRVDSV, "
   g_str_Parame = g_str_Parame & "        BIEADJ_FECREA, BIEADJ_MTOREA, BIEADJ_MTOLIB, BIEADJ_SITUAC,  "
   g_str_Parame = g_str_Parame & "        BIEADJ_PERMES, BIEADJ_PERANO  "
   g_str_Parame = g_str_Parame & "   FROM CRE_BIEADJ "
   g_str_Parame = g_str_Parame & "  WHERE BIEADJ_PERMES = " & CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   g_str_Parame = g_str_Parame & "    AND BIEADJ_PERANO = " & ipp_PerAno.Text
   g_str_Parame = g_str_Parame & "    AND BIEADJ_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
      Exit Sub
   End If
   
   If g_rst_GenAux.BOF And g_rst_GenAux.EOF Then
      moddat_g_int_FlgGrb = 1 'insert
      g_rst_GenAux.Close
      Set g_rst_GenAux = Nothing
      Exit Sub
   End If
      
   'If g_rst_GenAux!BIEADJ_SITUAC = 2 Then
   '   Call fs_Validar_Botones(False)
   'End If
   If modctb_int_PerAno & Format(modctb_int_PerMes, "00") <> g_rst_GenAux!BIEADJ_PERANO & Format(g_rst_GenAux!BIEADJ_PERMES, "00") Then
      Call fs_Validar_Botones(False)
   End If
   
   g_rst_GenAux.MoveFirst
   moddat_g_int_FlgGrb = 2 'editar
   
   Call gs_BuscarCombo_Item(cmb_TipOrigen, g_rst_GenAux!BIEADJ_TIPORI)
      
   If IsNull(g_rst_GenAux!BIEADJ_FECADJ) Then
      ipp_FecAdj.Text = date
   Else
      ipp_FecAdj.Text = gf_FormatoFecha(g_rst_GenAux!BIEADJ_FECADJ)
   End If
   If IsNull(g_rst_GenAux!BIEADJ_FECVTO) Then
      ipp_FecVcto.Text = ""
   Else
      ipp_FecVcto.Text = gf_FormatoFecha(g_rst_GenAux!BIEADJ_FECVTO)
   End If
   
   Call gs_BuscarCombo_Item(cmb_Estado, g_rst_GenAux!BIEADJ_SITUAC)
   
   If IsNull(g_rst_GenAux!BIEADJ_CAPDEU) Then
      ipp_CapDeu.Text = Format(0, "###,###,#00.00") & " "
   Else
      ipp_CapDeu.Text = Format(CDbl(g_rst_GenAux!BIEADJ_CAPDEU), "###,###,#00.00") & " "
   End If
   
   If IsNull(g_rst_GenAux!BIEADJ_INTDEU) Then
      ipp_IntDeu.Text = Format(0, "###,###,#00.00") & " "
   Else
      ipp_IntDeu.Text = Format(CDbl(g_rst_GenAux!BIEADJ_INTDEU), "###,###,#00.00") & " "
   End If
   
   'CALCULO SUMA TOTAL
   If IsNull(g_rst_GenAux!BIEADJ_PRVINI) Then
      ipp_PrvIni.Text = Format(0, "###,###,#00.00") & " "
   Else
      ipp_PrvIni.Text = Format(CDbl(g_rst_GenAux!BIEADJ_PRVINI), "###,###,#00.00") & " "
   End If
   
   If IsNull(g_rst_GenAux!BIEADJ_PRVCON) Then
      ipp_PrvCon.Text = Format(0, "###,###,#00.00") & " "
   Else
      ipp_PrvCon.Text = Format(CDbl(g_rst_GenAux!BIEADJ_PRVCON), "###,###,#00.00") & " "
   End If
   
   If IsNull(g_rst_GenAux!BIEADJ_PRVDSV) Then
      ipp_PrvDsv.Text = Format(0, "###,###,#00.00") & " "
   Else
      ipp_PrvDsv.Text = Format(CDbl(g_rst_GenAux!BIEADJ_PRVDSV), "###,###,#00.00") & " "
   End If
   
   'CALCULO SUMA TOTAL
   If IsNull(g_rst_GenAux!BIEADJ_FECREA) Then
      ipp_FecRea.Text = ""
   Else
      ipp_FecRea.Text = gf_FormatoFecha(g_rst_GenAux!BIEADJ_FECREA)
   End If
   
   If IsNull(g_rst_GenAux!BIEADJ_MTOREA) Then
      ipp_ValRea.Text = Format(0, "###,###,#00.00") & " "
   Else
      ipp_ValRea.Text = Format(CDbl(g_rst_GenAux!BIEADJ_MTOREA), "###,###,#00.00") & " "
   End If
   
   If IsNull(g_rst_GenAux!BIEADJ_MTOLIB) Then
      ipp_ValLib.Text = Format(0, "###,###,#00.00") & " "
   Else
      ipp_ValLib.Text = Format(CDbl(g_rst_GenAux!BIEADJ_MTOLIB), "###,###,#00.00") & " "
   End If
         
   Call fs_CalculaTotales
   
   g_rst_GenAux.Close
   Set g_rst_GenAux = Nothing
End Sub

Private Sub cmd_Grabar_Click()
   If cmb_TipOrigen.ListIndex = -1 Then
      MsgBox "Tiene que seleccionar un tipo de origen.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipOrigen)
      Exit Sub
   End If
   If Trim(ipp_FecAdj.Text & "") = "" Then
      MsgBox "Tiene que ingresar un fecha de origen.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecAdj)
      Exit Sub
   End If
   If CDbl(ipp_CapDeu.Text) = 0 Then
      MsgBox "Tiene que ingresar el monto que se cancelo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_CapDeu)
      Exit Sub
   End If
   If cmb_Estado.ListIndex = -1 Then
      MsgBox "Tiene que seleccionar un estado.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Estado)
      Exit Sub
   End If
   
   If MsgBox("¿Esta seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Me.Enabled = False
   Call fs_Grabar
   Call cmd_Limpia_Click
   Me.Enabled = True
   Screen.MousePointer = 0
End Sub

Private Sub fs_Grabar()
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " USP_CRE_BIEADJ ( "
   g_str_Parame = g_str_Parame & CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)) & ", "
   g_str_Parame = g_str_Parame & ipp_PerAno & ", "
   g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "', "
   g_str_Parame = g_str_Parame & cmb_TipOrigen.ItemData(cmb_TipOrigen.ListIndex) & ", "
   g_str_Parame = g_str_Parame & Format(ipp_FecAdj.Text, "yyyymmdd") & ", "
   If Trim(ipp_FecVcto.Text) = "" Then
      g_str_Parame = g_str_Parame & "null, "
   Else
      g_str_Parame = g_str_Parame & Format(ipp_FecVcto.Text, "yyyymmdd") & ", "
   End If
   g_str_Parame = g_str_Parame & CDbl(ipp_CapDeu.Text) & ", "
   g_str_Parame = g_str_Parame & CDbl(ipp_IntDeu.Text) & ", "
   g_str_Parame = g_str_Parame & CDbl(ipp_PrvIni.Text) & ", "
   g_str_Parame = g_str_Parame & CDbl(ipp_PrvCon.Text) & ", "
   g_str_Parame = g_str_Parame & CDbl(ipp_PrvDsv.Text) & ", "
   If Trim(ipp_FecRea.Text) = "" Then
      g_str_Parame = g_str_Parame & "null, "
   Else
      g_str_Parame = g_str_Parame & Format(ipp_FecRea.Text, "yyyymmdd") & ", "
   End If
   g_str_Parame = g_str_Parame & CDbl(ipp_ValRea.Text) & ", "
   g_str_Parame = g_str_Parame & CDbl(ipp_ValLib.Text) & ", "
   g_str_Parame = g_str_Parame & cmb_Estado.ItemData(cmb_Estado.ListIndex) & ", " 'BIEADJ_SITUAC
   'Datos de Auditoria
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                  'Código Usuario
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                  'Nombre Terminal
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                   'Nombre Ejecutable
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "                  'Código Sucursal
   g_str_Parame = g_str_Parame & CStr(moddat_g_int_FlgGrb) & ") "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      MsgBox "No se pudo completar la grabación de los datos.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   Else
       'Call Grabar_Auditoria
      If (g_rst_Genera!RESUL = 1) Then
          MsgBox "Los datos se grabaron correctamente.", vbInformation, modgen_g_str_NomPlt
      ElseIf (g_rst_Genera!RESUL = 2) Then
          MsgBox "Los datos se modificaron correctamente.", vbInformation, modgen_g_str_NomPlt
      End If
   End If
End Sub

'Controles
Private Sub msk_NumOpe_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(cmb_PerMes)
    End If
End Sub

Private Sub cmb_PerMes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(ipp_PerAno)
    End If
End Sub

Private Sub ipp_PerAno_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(cmd_Buscar)
    End If
End Sub

Private Sub cmb_TipOrigen_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
       Call gs_SetFocus(ipp_FecAdj)
   End If
End Sub

Private Sub ipp_FecAdj_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
       Call gs_SetFocus(ipp_FecVcto)
   End If
End Sub

Private Sub ipp_FecVcto_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmb_Estado.Enabled = False Then
         Call gs_SetFocus(ipp_CapDeu)
      Else
         Call gs_SetFocus(cmb_Estado)
      End If
   End If
End Sub

Private Sub cmb_Estado_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
       Call gs_SetFocus(ipp_CapDeu)
   End If
End Sub

Private Sub ipp_CapDeu_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
       Call gs_SetFocus(ipp_IntDeu)
   End If
End Sub

Private Sub ipp_CapDeu_LostFocus()
   Call fs_CalculaTotales
End Sub

Private Sub ipp_IntDeu_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
       Call gs_SetFocus(ipp_PrvIni)
   End If
End Sub

Private Sub ipp_IntDeu_LostFocus()
   Call fs_CalculaTotales
End Sub

Private Sub ipp_PrvIni_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
       Call gs_SetFocus(ipp_PrvCon)
   End If
End Sub

Private Sub ipp_PrvIni_LostFocus()
   Call fs_CalculaTotales
End Sub

Private Sub ipp_PrvCon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
       Call gs_SetFocus(ipp_PrvDsv)
   End If
End Sub

Private Sub ipp_PrvCon_LostFocus()
   Call fs_CalculaTotales
End Sub

Private Sub ipp_PrvDsv_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
       Call gs_SetFocus(ipp_FecRea)
   End If
End Sub

Private Sub ipp_PrvDsv_LostFocus()
   Call fs_CalculaTotales
End Sub

Private Sub ipp_FecRea_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
       Call gs_SetFocus(ipp_ValRea)
   End If
End Sub

Private Sub ipp_ValRea_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
       Call gs_SetFocus(ipp_ValLib)
   End If
End Sub

Private Sub ipp_ValLib_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
       Call gs_SetFocus(cmd_Grabar)
   End If
End Sub

Private Sub fs_CalculaTotales()
   pnl_TotDeuda.Caption = Format(CDbl(ipp_IntDeu.Text) + CDbl(ipp_CapDeu.Text), "###,###,#0.00") & " "
   pnl_PrvTot.Caption = Format(CDbl(ipp_PrvIni.Text) + CDbl(ipp_PrvCon.Text) + CDbl(ipp_PrvDsv.Text), "###,###,#0.00") & " "
   ipp_ValRea.Value = CDbl(pnl_TotDeuda.Caption) - CDbl(pnl_PrvTot.Caption)
   ipp_ValLib.Value = CDbl(pnl_TotDeuda.Caption) - CDbl(pnl_PrvTot.Caption)
End Sub
