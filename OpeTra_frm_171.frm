VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_MntCli_61 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6360
   ClientLeft      =   2175
   ClientTop       =   2400
   ClientWidth     =   11640
   Icon            =   "OpeTra_frm_171.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   6375
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   11685
      _Version        =   65536
      _ExtentX        =   20611
      _ExtentY        =   11245
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
         Height          =   2085
         Left            =   30
         TabIndex        =   31
         Top             =   3750
         Width           =   11595
         _Version        =   65536
         _ExtentX        =   20452
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
         Begin VB.ComboBox cmb_SegPro 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   60
            Width           =   765
         End
         Begin VB.TextBox txt_NomArr_2 
            Height          =   315
            Left            =   2010
            MaxLength       =   250
            TabIndex        =   8
            Text            =   "Text1"
            Top             =   720
            Width           =   9525
         End
         Begin VB.TextBox txt_Telef2_2 
            Height          =   315
            Left            =   3660
            MaxLength       =   25
            TabIndex        =   11
            Text            =   "Text1"
            Top             =   1380
            Width           =   1640
         End
         Begin VB.TextBox txt_Direcc_2 
            Height          =   315
            Left            =   2010
            MaxLength       =   250
            TabIndex        =   7
            Text            =   "Text1"
            Top             =   390
            Width           =   9525
         End
         Begin VB.TextBox txt_Telef1_2 
            Height          =   315
            Left            =   2010
            MaxLength       =   25
            TabIndex        =   10
            Text            =   "Text1"
            Top             =   1380
            Width           =   1640
         End
         Begin EditLib.fpDateTime ipp_IniAlq_2 
            Height          =   315
            Left            =   2010
            TabIndex        =   9
            Top             =   1050
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
         Begin EditLib.fpDoubleSingle ipp_AlqMen_2 
            Height          =   315
            Left            =   2010
            TabIndex        =   12
            Top             =   1710
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
         Begin VB.Label Label11 
            Caption         =   "2da Propiedad:"
            Height          =   285
            Left            =   90
            TabIndex        =   37
            Top             =   60
            Width           =   1785
         End
         Begin VB.Label lbl_General 
            Caption         =   "Nombre Arredantario:"
            Height          =   285
            Index           =   1
            Left            =   90
            TabIndex        =   36
            Top             =   720
            Width           =   1785
         End
         Begin VB.Label lbl_General 
            Caption         =   "Teléfono (s):"
            Height          =   285
            Index           =   2
            Left            =   90
            TabIndex        =   35
            Top             =   1380
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Dirección Propiedad:"
            Height          =   285
            Index           =   3
            Left            =   90
            TabIndex        =   34
            Top             =   390
            Width           =   1695
         End
         Begin VB.Label lbl_General 
            Caption         =   "Fecha Inicio Alquiler:"
            Height          =   315
            Index           =   4
            Left            =   90
            TabIndex        =   33
            Top             =   1050
            Width           =   1845
         End
         Begin VB.Label lbl_General 
            Caption         =   "Alquiler Mensual:"
            Height          =   285
            Index           =   5
            Left            =   90
            TabIndex        =   32
            Top             =   1710
            Width           =   1755
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   1755
         Left            =   30
         TabIndex        =   18
         Top             =   1950
         Width           =   11595
         _Version        =   65536
         _ExtentX        =   20452
         _ExtentY        =   3096
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
         Begin VB.TextBox txt_NomArr_1 
            Height          =   315
            Left            =   2010
            MaxLength       =   250
            TabIndex        =   1
            Text            =   "Text1"
            Top             =   390
            Width           =   9525
         End
         Begin VB.TextBox txt_Telef2_1 
            Height          =   315
            Left            =   3660
            MaxLength       =   25
            TabIndex        =   4
            Text            =   "Text1"
            Top             =   1050
            Width           =   1640
         End
         Begin VB.TextBox txt_Direcc_1 
            Height          =   315
            Left            =   2010
            MaxLength       =   250
            TabIndex        =   0
            Text            =   "Text1"
            Top             =   60
            Width           =   9525
         End
         Begin VB.TextBox txt_Telef1_1 
            Height          =   315
            Left            =   2010
            MaxLength       =   25
            TabIndex        =   3
            Text            =   "Text1"
            Top             =   1050
            Width           =   1640
         End
         Begin EditLib.fpDateTime ipp_IniAlq_1 
            Height          =   315
            Left            =   2010
            TabIndex        =   2
            Top             =   720
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
         Begin EditLib.fpDoubleSingle ipp_AlqMen_1 
            Height          =   315
            Left            =   2010
            TabIndex        =   5
            Top             =   1380
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
         Begin VB.Label lbl_General 
            Caption         =   "Nombre Arredantario:"
            Height          =   285
            Index           =   49
            Left            =   90
            TabIndex        =   23
            Top             =   390
            Width           =   1785
         End
         Begin VB.Label lbl_General 
            Caption         =   "Teléfono (s):"
            Height          =   285
            Index           =   46
            Left            =   90
            TabIndex        =   22
            Top             =   1050
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Dirección Propiedad:"
            Height          =   285
            Index           =   37
            Left            =   90
            TabIndex        =   21
            Top             =   60
            Width           =   1695
         End
         Begin VB.Label lbl_General 
            Caption         =   "Fecha Inicio Alquiler:"
            Height          =   315
            Index           =   58
            Left            =   90
            TabIndex        =   20
            Top             =   720
            Width           =   1845
         End
         Begin VB.Label lbl_General 
            Caption         =   "Alquiler Mensual:"
            Height          =   285
            Index           =   0
            Left            =   90
            TabIndex        =   19
            Top             =   1380
            Width           =   1755
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   24
         Top             =   30
         Width           =   11595
         _Version        =   65536
         _ExtentX        =   20452
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
         Begin Threed.SSPanel SSPanel16 
            Height          =   255
            Left            =   600
            TabIndex        =   29
            Top             =   60
            Width           =   6165
            _Version        =   65536
            _ExtentX        =   10874
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "Mantenimiento de Clientes"
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
         Begin Threed.SSPanel SSPanel7 
            Height          =   255
            Left            =   600
            TabIndex        =   30
            Top             =   330
            Width           =   6165
            _Version        =   65536
            _ExtentX        =   10874
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "Actividades Económicas - Rentista"
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
            Picture         =   "OpeTra_frm_171.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   435
         Left            =   30
         TabIndex        =   25
         Top             =   1470
         Width           =   11595
         _Version        =   65536
         _ExtentX        =   20452
         _ExtentY        =   767
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
         Begin Threed.SSPanel pnl_NomCli 
            Height          =   315
            Left            =   2010
            TabIndex        =   26
            Top             =   60
            Width           =   9525
            _Version        =   65536
            _ExtentX        =   16801
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "1-07522154 / IKEHARA PUNK MIGUEL ANGEL (1-07521154 / IKEHARA PUNK MIGUEL ANGEL)"
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
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   90
            TabIndex        =   27
            Top             =   60
            Width           =   1725
         End
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   675
         Left            =   30
         TabIndex        =   28
         Top             =   750
         Width           =   11595
         _Version        =   65536
         _ExtentX        =   20452
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
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_171.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10980
            Picture         =   "OpeTra_frm_171.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   435
         Left            =   30
         TabIndex        =   38
         Top             =   5880
         Width           =   11595
         _Version        =   65536
         _ExtentX        =   20452
         _ExtentY        =   767
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
         Begin VB.ComboBox cmb_MonIng 
            Height          =   315
            Left            =   1980
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   60
            Width           =   3315
         End
         Begin VB.CommandButton cmd_BusEmp_Ant 
            Caption         =   "..."
            Height          =   315
            Left            =   10620
            TabIndex        =   39
            ToolTipText     =   "Obtener Dirección de Domicilio"
            Top             =   6600
            Width           =   435
         End
         Begin EditLib.fpDoubleSingle ipp_IngNet 
            Height          =   315
            Left            =   8190
            TabIndex        =   14
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
         Begin Threed.SSPanel pnl_FlgEmp_Ant 
            Height          =   315
            Left            =   11100
            TabIndex        =   40
            Top             =   6600
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "NR"
            ForeColor       =   255
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   0
         End
         Begin VB.Label lbl_General 
            Caption         =   "Moneda de Ingresos:"
            Height          =   285
            Index           =   7
            Left            =   60
            TabIndex        =   42
            Top             =   60
            Width           =   1665
         End
         Begin VB.Label lbl_General 
            Caption         =   "Ingreso Declarado:"
            Height          =   285
            Index           =   61
            Left            =   6180
            TabIndex        =   41
            Top             =   90
            Width           =   1755
         End
      End
   End
End
Attribute VB_Name = "frm_MntCli_61"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmb_MonIng_Click()
   Call gs_SetFocus(ipp_IngNet)
End Sub

Private Sub cmb_MonIng_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_MonIng_Click
   End If
End Sub

Private Sub cmb_SegPro_Click()
   If cmb_SegPro.ListIndex = -1 Then
      txt_Direcc_2.Text = ""
      txt_NomArr_2.Text = ""
      ipp_IniAlq_2.Text = Format(date, "dd/mm/yyyy")
      txt_Telef1_2.Text = ""
      txt_Telef2_2.Text = ""
      ipp_AlqMen_2.Value = 0
      
      txt_Direcc_2.Enabled = False
      txt_NomArr_2.Enabled = False
      ipp_IniAlq_2.Enabled = False
      txt_Telef1_2.Enabled = False
      txt_Telef2_2.Enabled = False
      ipp_AlqMen_2.Enabled = False
      
      Call gs_SetFocus(cmb_MonIng)
   Else
      If cmb_SegPro.ItemData(cmb_SegPro.ListIndex) = 1 Then
         txt_Direcc_2.Enabled = True
         txt_NomArr_2.Enabled = True
         ipp_IniAlq_2.Enabled = True
         txt_Telef1_2.Enabled = True
         txt_Telef2_2.Enabled = True
         ipp_AlqMen_2.Enabled = True
      
         Call gs_SetFocus(txt_Direcc_2)
      Else
         txt_Direcc_2.Text = ""
         txt_NomArr_2.Text = ""
         ipp_IniAlq_2.Text = Format(date, "dd/mm/yyyy")
         txt_Telef1_2.Text = ""
         txt_Telef2_2.Text = ""
         ipp_AlqMen_2.Value = 0
         
         txt_Direcc_2.Enabled = False
         txt_NomArr_2.Enabled = False
         ipp_IniAlq_2.Enabled = False
         txt_Telef1_2.Enabled = False
         txt_Telef2_2.Enabled = False
         ipp_AlqMen_2.Enabled = False
         
         Call gs_SetFocus(cmb_MonIng)
      End If
   End If
End Sub

Private Sub cmb_SegPro_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_SegPro_Click
   End If
End Sub

Private Sub cmd_Grabar_Click()
   If Len(Trim(txt_Direcc_1.Text)) = 0 Then
      MsgBox "Debe ingresar la Dirección de la Propiedad.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Direcc_1)
      Exit Sub
   End If

   If Len(Trim(txt_NomArr_1.Text)) = 0 Then
      MsgBox "Debe ingresar el Nombre del Arrendatario.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NomArr_1)
      Exit Sub
   End If

   If CDate(ipp_IniAlq_1.Text) > date Then
      MsgBox "La Fecha de Inicio de Alquiler no puede ser mayor a la fecha actual.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_IniAlq_1)
      Exit Sub
   End If

   If Len(Trim(txt_Telef1_1.Text)) = 0 Then
      MsgBox "Debe ingresar el Teléfono del Arrendatario.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Telef1_1)
      Exit Sub
   End If

   If ipp_AlqMen_1.Value = 0 Then
      MsgBox "El Alquiler Mensual no puede ser igual a cero.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_AlqMen_1)
      Exit Sub
   End If

   If cmb_SegPro.ListIndex = -1 Then
      MsgBox "Debe seleccionar si el Cliente presenta Segunda Propiedad.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_SegPro)
      Exit Sub
   End If
   
   If cmb_SegPro.ItemData(cmb_SegPro.ListIndex) = 1 Then
      If Len(Trim(txt_Direcc_2.Text)) = 0 Then
         MsgBox "Debe ingresar la Dirección de la Propiedad.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_Direcc_2)
         Exit Sub
      End If
   
      If Len(Trim(txt_NomArr_2.Text)) = 0 Then
         MsgBox "Debe ingresar el Nombre del Arrendatario.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NomArr_2)
         Exit Sub
      End If
   
      If CDate(ipp_IniAlq_2.Text) > date Then
         MsgBox "La Fecha de Inicio de Alquiler no puede ser mayor a la fecha actual.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_IniAlq_2)
         Exit Sub
      End If
   
      If Len(Trim(txt_Telef1_2.Text)) = 0 Then
         MsgBox "Debe ingresar el Teléfono del Arrendatario.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_Telef1_2)
         Exit Sub
      End If
   
      If ipp_AlqMen_2.Value = 0 Then
         MsgBox "El Alquiler Mensual no puede ser igual a cero.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_AlqMen_2)
         Exit Sub
      End If
   End If

   If cmb_MonIng.ListIndex = -1 Then
      MsgBox "Debe ingresar la Moneda de Ingresos.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_MonIng)
      Exit Sub
   End If
   
   If ipp_IngNet.Value = 0 Then
      MsgBox "El Ingreso Declarado no puede ser igual a cero.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_IngNet)
      Exit Sub
   End If

   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   
   If modmip_g_int_FlgGrb_1 = 2 Then
      'Borrar Actividad Económica
      g_str_Parame = "DELETE FROM CLI_ACTECO WHERE "
      
      If modmip_g_int_TipCli = 1 Then
         g_str_Parame = g_str_Parame & "ACTECO_CLITDO = " & CStr(moddat_g_int_TipDoc) & " AND "
         g_str_Parame = g_str_Parame & "ACTECO_CLINDO = '" & CStr(moddat_g_str_NumDoc) & "' AND "
      Else
         g_str_Parame = g_str_Parame & "ACTECO_CLITDO = " & CStr(moddat_g_int_CygTDo) & " AND "
         g_str_Parame = g_str_Parame & "ACTECO_CLINDO = '" & CStr(moddat_g_str_CygNDo) & "' AND "
      End If
      
      g_str_Parame = g_str_Parame & "ACTECO_ORDACT = " & CStr(modmip_g_int_OrdAct) & " "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
          Exit Sub
      End If
   End If
   
   'Insertando Actividad Económica
   g_str_Parame = "USP_CLI_ACTECO_AGREGA ("
   
   If modmip_g_int_TipCli = 1 Then
      g_str_Parame = g_str_Parame & CStr(moddat_g_int_TipDoc) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumDoc & "', "
   End If
   
   g_str_Parame = g_str_Parame & CStr(modmip_g_int_OrdAct) & ", "
   g_str_Parame = g_str_Parame & "51, "                                                      'Código Actividad Económica (Rentista)
   
   'Dependiente
   g_str_Parame = g_str_Parame & "0, "                                                       'Tipo DOI Empleador
   g_str_Parame = g_str_Parame & "'', "                                                      'Número DOI Empleador
   g_str_Parame = g_str_Parame & "'', "                                                      'Razón Social
   g_str_Parame = g_str_Parame & "'', "                                                      'Nombre Comercial
   g_str_Parame = g_str_Parame & "0, "                                                       'Tipo Oficina
   g_str_Parame = g_str_Parame & "0, "                                                       'Situación trabajador
   g_str_Parame = g_str_Parame & "0, "                                                       'Tipo de Via
   g_str_Parame = g_str_Parame & "'', "                                                      'Nombre de Vía
   g_str_Parame = g_str_Parame & "'', "                                                      'Número de Vía
   g_str_Parame = g_str_Parame & "'', "                                                      'Interior / Dpto.
   g_str_Parame = g_str_Parame & "0, "                                                       'Tipo de Zona
   g_str_Parame = g_str_Parame & "'', "                                                      'Nombre de Zona
   g_str_Parame = g_str_Parame & "'', "                                                      'Ubicación Geográfica
   g_str_Parame = g_str_Parame & "'', "                                                      'Referencia
   g_str_Parame = g_str_Parame & "'', "                                                      'Teléfono 1
   g_str_Parame = g_str_Parame & "'', "                                                      'Teléfono 2
   g_str_Parame = g_str_Parame & "'', "                                                      'Fax
   g_str_Parame = g_str_Parame & "0, "                                                       'Código CIIU
   g_str_Parame = g_str_Parame & "'', "                                                      'Telefono RR.HH
   g_str_Parame = g_str_Parame & "'', "                                                      'Anexo RR.HH
   g_str_Parame = g_str_Parame & "0, "                                                       'Ingreso Neto
   g_str_Parame = g_str_Parame & "0, "                                                       'Frecuencia de Haberes
   g_str_Parame = g_str_Parame & "0, "                                                       'Fecha de Ingreso
   g_str_Parame = g_str_Parame & "'', "                                                      'Código de Cargo
   g_str_Parame = g_str_Parame & "'', "                                                      'Nombre de Cargo
   g_str_Parame = g_str_Parame & "'', "                                                      'Nombre de Area
   g_str_Parame = g_str_Parame & "'', "                                                      'Número de Anexo
   g_str_Parame = g_str_Parame & "'', "                                                      'Teléfono Directo
   g_str_Parame = g_str_Parame & "'', "                                                      'Celular del Trabajo
   g_str_Parame = g_str_Parame & "'', "                                                      'E-mail del Trabajo
   g_str_Parame = g_str_Parame & "2, "                                                       'Flag de Trabajo Anterior
   g_str_Parame = g_str_Parame & "0, "                                                       'Tipo DOI Empleador Anterior
   g_str_Parame = g_str_Parame & "'', "                                                      'Número DOI Empleador Anterior
   g_str_Parame = g_str_Parame & "'', "                                                      'Razón Social Empleador Anterior
   g_str_Parame = g_str_Parame & "'', "                                                      'Nombre Comercial Empleador Anterior
   g_str_Parame = g_str_Parame & "'', "                                                      'Teléfono 1 Empleador Anterior
   g_str_Parame = g_str_Parame & "'', "                                                      'Teléfono 2 Empleador Anterior
   g_str_Parame = g_str_Parame & "0, "                                                       'Fecha Ingreso Empleador Anterior
   g_str_Parame = g_str_Parame & "0, "                                                       'Fecha Cese Empleador Anterior
   
   'Independiente
   g_str_Parame = g_str_Parame & "0, "                                                    'Tipo DOI
   g_str_Parame = g_str_Parame & "'', "                                                   'Número DOI
   g_str_Parame = g_str_Parame & "0, "                                                    'Tipo Vía
   g_str_Parame = g_str_Parame & "'', "                                                   'Nombre Vía
   g_str_Parame = g_str_Parame & "'', "                                                   'Número Vía
   g_str_Parame = g_str_Parame & "'', "                                                   'Interior
   g_str_Parame = g_str_Parame & "0, "                                                    'Tipo Zona
   g_str_Parame = g_str_Parame & "'', "                                                   'Nombre zona
   g_str_Parame = g_str_Parame & "'', "                                                   'Ubicación Geográfica
   g_str_Parame = g_str_Parame & "'', "                                                   'Referencia
   g_str_Parame = g_str_Parame & "'', "                                                   'Teléfono 1
   g_str_Parame = g_str_Parame & "'', "                                                   'Teléfono 2
   g_str_Parame = g_str_Parame & "'', "                                                   'Fax
   g_str_Parame = g_str_Parame & "0, "                                                    'CIIU
   g_str_Parame = g_str_Parame & "0, "                                                    'Ingreso Neto
   g_str_Parame = g_str_Parame & "0, "                                                    'Inicio de Actividad
   g_str_Parame = g_str_Parame & "0, "                                                    'Contrato Locación
   g_str_Parame = g_str_Parame & "0, "                                                    'Tipo DOI Empleador
   g_str_Parame = g_str_Parame & "'', "                                                   'Número DOI Empleador
   g_str_Parame = g_str_Parame & "'', "                                                   'Razón Social Empleador
   g_str_Parame = g_str_Parame & "'', "                                                   'Nombre Comercial Empleador
   g_str_Parame = g_str_Parame & "'', "                                                   'Teléfono 1 Empleador
   g_str_Parame = g_str_Parame & "'', "                                                   'Teléfono 2 Empleador
   g_str_Parame = g_str_Parame & "0, "                                                    'Fecha Ingreso Empleador
   g_str_Parame = g_str_Parame & "'', "                                                   'Código Cargo
   g_str_Parame = g_str_Parame & "'', "                                                   'Nombre Cargo
   
   'Comerciante
   g_str_Parame = g_str_Parame & "0, "                                                    'Tipo DOI
   g_str_Parame = g_str_Parame & "'', "                                                   'Número DOI
   g_str_Parame = g_str_Parame & "'', "                                                   'Razón Social Empleador
   g_str_Parame = g_str_Parame & "'', "                                                   'Nombre Comercial Empleador
   g_str_Parame = g_str_Parame & "0, "                                                    'Tipo Vía
   g_str_Parame = g_str_Parame & "'', "                                                   'Nombre Vía
   g_str_Parame = g_str_Parame & "'', "                                                   'Número Vía
   g_str_Parame = g_str_Parame & "'', "                                                   'Interior
   g_str_Parame = g_str_Parame & "0, "                                                    'Tipo Zona
   g_str_Parame = g_str_Parame & "'', "                                                   'Nombre zona
   g_str_Parame = g_str_Parame & "'', "                                                   'Ubicación Geográfica
   g_str_Parame = g_str_Parame & "'', "                                                   'Referencia
   g_str_Parame = g_str_Parame & "'', "                                                   'Teléfono 1
   g_str_Parame = g_str_Parame & "'', "                                                   'Teléfono 2
   g_str_Parame = g_str_Parame & "'', "                                                   'Fax
   g_str_Parame = g_str_Parame & "0, "                                                    'CIIU
   g_str_Parame = g_str_Parame & "'', "                                                   'Giro Comercial
   g_str_Parame = g_str_Parame & "0, "                                                    'Ingreso Neto
   g_str_Parame = g_str_Parame & "0, "                                                    'Ventas Mensuales
   g_str_Parame = g_str_Parame & "0, "                                                    'Inicio de Operaciones
   g_str_Parame = g_str_Parame & "'', "                                                   'Código Cargo
   g_str_Parame = g_str_Parame & "'', "                                                   'Nombre Cargo
   g_str_Parame = g_str_Parame & "0, "                                                    'Régimen Tributario
   g_str_Parame = g_str_Parame & "0, "                                                    'Porcentaje Participación
   g_str_Parame = g_str_Parame & "0, "                                                    'Tipo Local
   g_str_Parame = g_str_Parame & "0, "                                                    'Alquiler Mensual
   g_str_Parame = g_str_Parame & "'', "                                                   'Nombre Arrendador
   g_str_Parame = g_str_Parame & "'', "                                                   'Teléfono Arrendador
   
   'Accionista
   g_str_Parame = g_str_Parame & "0, "                                                    'Tipo DOI
   g_str_Parame = g_str_Parame & "'', "                                                   'Número DOI
   g_str_Parame = g_str_Parame & "'', "                                                   'Razón Social Empleador
   g_str_Parame = g_str_Parame & "'', "                                                   'Nombre Comercial Empleador
   g_str_Parame = g_str_Parame & "0, "                                                    'Tipo Vía
   g_str_Parame = g_str_Parame & "'', "                                                   'Nombre Vía
   g_str_Parame = g_str_Parame & "'', "                                                   'Número Vía
   g_str_Parame = g_str_Parame & "'', "                                                   'Interior
   g_str_Parame = g_str_Parame & "0, "                                                    'Tipo Zona
   g_str_Parame = g_str_Parame & "'', "                                                   'Nombre zona
   g_str_Parame = g_str_Parame & "'', "                                                   'Ubicación Geográfica
   g_str_Parame = g_str_Parame & "'', "                                                   'Referencia
   g_str_Parame = g_str_Parame & "'', "                                                   'Teléfono 1
   g_str_Parame = g_str_Parame & "'', "                                                   'Teléfono 2
   g_str_Parame = g_str_Parame & "'', "                                                   'Fax
   g_str_Parame = g_str_Parame & "0, "                                                    'CIIU
   g_str_Parame = g_str_Parame & "0, "                                                    'Ingreso Neto
   g_str_Parame = g_str_Parame & "0, "                                                    'Porcentaje Participación
   g_str_Parame = g_str_Parame & "0, "                                                    'Fecha Antigüedad
   
   'Rentista
   g_str_Parame = g_str_Parame & CStr(ipp_IngNet.Value) & ","                             'Ingreso Neto
   g_str_Parame = g_str_Parame & "'" & txt_Direcc_1.Text & "', "                          'Dirección 1
   g_str_Parame = g_str_Parame & "'" & txt_NomArr_1.Text & "', "                          'Nombre 1
   g_str_Parame = g_str_Parame & Format(CDate(ipp_IniAlq_1.Text), "yyyymmdd") & ", "           'Inicio Alquiler 1
   g_str_Parame = g_str_Parame & "'" & txt_Telef1_1.Text & "', "                          'Teléfono 1 - 1
   g_str_Parame = g_str_Parame & "'" & txt_Telef2_1.Text & "', "                          'Teléfono 2 - 1
   g_str_Parame = g_str_Parame & CStr(ipp_AlqMen_1.Value) & ","                           'Monto Alquiler 1
   g_str_Parame = g_str_Parame & CStr(cmb_SegPro.ItemData(cmb_SegPro.ListIndex)) & ", "   'Tipo DOI
   
   If cmb_SegPro.ItemData(cmb_SegPro.ListIndex) = 1 Then
      g_str_Parame = g_str_Parame & "'" & txt_Direcc_2.Text & "', "                          'Dirección 2
      g_str_Parame = g_str_Parame & "'" & txt_NomArr_2.Text & "', "                          'Nombre 2
      g_str_Parame = g_str_Parame & Format(CDate(ipp_IniAlq_2.Text), "yyyymmdd") & ", "           'Inicio Alquiler 2
      g_str_Parame = g_str_Parame & "'" & txt_Telef1_2.Text & "', "                          'Teléfono 1 - 2
      g_str_Parame = g_str_Parame & "'" & txt_Telef2_2.Text & "', "                          'Teléfono 2 - 2
      g_str_Parame = g_str_Parame & CStr(ipp_AlqMen_2.Value) & ","                           'Monto Alquiler 2
   Else
      g_str_Parame = g_str_Parame & "'', "                                                   'Dirección 2
      g_str_Parame = g_str_Parame & "'', "                                                   'Nombre 2
      g_str_Parame = g_str_Parame & "0, "                                                    'Inicio Alquiler 2
      g_str_Parame = g_str_Parame & "'', "                                                   'Teléfono 1 - 2
      g_str_Parame = g_str_Parame & "'', "                                                   'Teléfono 2 - 2
      g_str_Parame = g_str_Parame & "0, "                                                    'Monto Alquiler 2
   End If
   
   'Otros
   g_str_Parame = g_str_Parame & "0, "                                                    'Ingreso Neto
   g_str_Parame = g_str_Parame & "'', "                                                   'Actividad
   g_str_Parame = g_str_Parame & "0, "                                                    'CIIU
   g_str_Parame = g_str_Parame & "'', "                                                   'Observaciones
   
   'Dependiente
   g_str_Parame = g_str_Parame & "0, "                                                    'Moneda Ingresos
   g_str_Parame = g_str_Parame & "'', "                                                   'Dirección
   g_str_Parame = g_str_Parame & "'', "                                                   'Ciudad
   g_str_Parame = g_str_Parame & "'', "                                                   'Código Postal
   
   'Independiente
   g_str_Parame = g_str_Parame & "0, "                                                    'Moneda Ingresos
   g_str_Parame = g_str_Parame & "'', "                                                   'Dirección
   g_str_Parame = g_str_Parame & "'', "                                                   'Ciudad
   g_str_Parame = g_str_Parame & "'', "                                                   'Código Postal
   
   'Comerciante
   g_str_Parame = g_str_Parame & "0, "                                                    'Moneda Ingresos
   g_str_Parame = g_str_Parame & "'', "                                                   'Dirección
   g_str_Parame = g_str_Parame & "'', "                                                   'Ciudad
   g_str_Parame = g_str_Parame & "'', "                                                   'Código Postal
   
   'Accionista
   g_str_Parame = g_str_Parame & "0, "                                                    'Moneda Ingresos
   
   'Rentista
   g_str_Parame = g_str_Parame & CStr(cmb_MonIng.ItemData(cmb_MonIng.ListIndex)) & ", "   'Moneda Ingresos
   
   'Otros
   g_str_Parame = g_str_Parame & "0, "                                                    'Moneda Ingresos
   
   'Dependiente
   g_str_Parame = g_str_Parame & "0, "    'TIPO IDENTIFICACION(ACTECO_IND_TIPIDE)
   
   'Independiente
   g_str_Parame = g_str_Parame & "0, "   'TIPO IDENTIFICACION(ACTECO_DEP_TIPIDE)
   
   'Datos de Auditoria
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
      MsgBox "Error al ejecutar el Procedimiento USP_CLI_ACTECO_AGREGA.", vbCritical, modgen_g_str_NomPlt
      
      Exit Sub
   End If
   
   If modmip_g_int_OrdAct = 1 Then
      'Actualizar en Maestro de Clientes
      g_str_Parame = "USP_CLI_DATGEN_ACTECOPRI ("
      
      If modmip_g_int_TipCli = 1 Then
         g_str_Parame = g_str_Parame & CStr(moddat_g_int_TipDoc) & ", "
         g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumDoc & "', "
      Else
         g_str_Parame = g_str_Parame & "ACTECO_CLITDO = " & CStr(moddat_g_int_CygTDo) & " AND "
         g_str_Parame = g_str_Parame & "ACTECO_CLINDO = '" & CStr(moddat_g_str_CygNDo) & "' AND "
      End If
      
      g_str_Parame = g_str_Parame & "51, "
      g_str_Parame = g_str_Parame & "9999, "
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "'', "
      
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "
         
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         MsgBox "Error al ejecutar el Procedimiento USP_CLI_ACTECO_AGREGA.", vbCritical, modgen_g_str_NomPlt
         
         Exit Sub
      End If
   End If
   
   modmip_g_int_FlgAct_1 = 2
   moddat_g_int_FlgAct = 2
   
   Screen.MousePointer = 0
   
   Unload Me
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   If modmip_g_int_TipCli = 1 Then
      pnl_NomCli.Caption = CStr(moddat_g_int_TipDoc) & " - " & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   Else
      pnl_NomCli.Caption = CStr(moddat_g_int_TipDoc) & " - " & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli & " (" & CStr(moddat_g_int_CygTDo) & " - " & moddat_g_str_CygNDo & " / " & moddat_g_str_CygNom & ")"
   End If
   
   Call fs_Inicio
   Call fs_Limpia
   
   If modmip_g_int_FlgGrb_1 = 2 Then
      g_str_Parame = "SELECT * FROM CLI_ACTECO WHERE "
      
      If modmip_g_int_TipCli = 1 Then
         g_str_Parame = g_str_Parame & "ACTECO_CLITDO = " & CStr(moddat_g_int_TipDoc) & " AND "
         g_str_Parame = g_str_Parame & "ACTECO_CLINDO = '" & CStr(moddat_g_str_NumDoc) & "' AND "
      Else
         g_str_Parame = g_str_Parame & "ACTECO_CLITDO = " & CStr(moddat_g_int_CygTDo) & " AND "
         g_str_Parame = g_str_Parame & "ACTECO_CLINDO = '" & CStr(moddat_g_str_CygNDo) & "' AND "
      End If
      
      g_str_Parame = g_str_Parame & "ACTECO_ORDACT = " & CStr(modmip_g_int_OrdAct) & " "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      txt_Direcc_1.Text = Trim(g_rst_Princi!ActEco_Ren_Direc1 & "")
      txt_NomArr_1.Text = Trim(g_rst_Princi!ActEco_Ren_NomAr1 & "")
      ipp_IniAlq_1.Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Ren_IniAl1))
      txt_Telef1_1.Text = Trim(g_rst_Princi!ActEco_Ren_Tele11 & "")
      txt_Telef2_1.Text = Trim(g_rst_Princi!ActEco_Ren_Tele21 & "")
      ipp_AlqMen_1.Value = g_rst_Princi!ActEco_Ren_AlqMe1
      
      Call gs_BuscarCombo_Item(cmb_SegPro, g_rst_Princi!ActEco_Ren_SegPro)
      
      If cmb_SegPro.ItemData(cmb_SegPro.ListIndex) = 1 Then
         txt_Direcc_2.Text = Trim(g_rst_Princi!ActEco_Ren_Direc2 & "")
         txt_NomArr_2.Text = Trim(g_rst_Princi!ActEco_Ren_NomAr2 & "")
         ipp_IniAlq_2.Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Ren_IniAl2))
         txt_Telef1_2.Text = Trim(g_rst_Princi!ActEco_Ren_Tele12 & "")
         txt_Telef2_2.Text = Trim(g_rst_Princi!ActEco_Ren_Tele22 & "")
         ipp_AlqMen_2.Value = g_rst_Princi!ActEco_Ren_AlqMe2
      End If
      
      Call gs_BuscarCombo_Item(cmb_MonIng, g_rst_Princi!ActEco_ren_MonIng)
      ipp_IngNet.Value = g_rst_Princi!ActEco_Ren_IngNet
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   End If
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub ipp_AlqMen_1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_SegPro)
   End If
End Sub

Private Sub ipp_AlqMen_2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_MonIng)
   End If
End Sub

Private Sub ipp_IngNet_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   End If
End Sub

Private Sub ipp_IniAlq_1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Telef1_1)
   End If
End Sub

Private Sub ipp_IniAlq_2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Telef1_2)
   End If
End Sub

Private Sub txt_Direcc_1_GotFocus()
   Call gs_SelecTodo(txt_Direcc_1)
End Sub

Private Sub txt_Direcc_1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NomArr_1)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "-_ ,.;:º#()/")
   End If
End Sub

Private Sub txt_NomArr_1_GotFocus()
   Call gs_SelecTodo(txt_NomArr_1)
End Sub

Private Sub txt_NomArr_1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_IniAlq_1)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "-_ ,.;:º#()/")
   End If
End Sub

Private Sub txt_Telef1_1_GotFocus()
   Call gs_SelecTodo(txt_Telef1_1)
End Sub

Private Sub txt_Telef1_1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Telef2_1)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "-")
   End If
End Sub

Private Sub txt_Telef2_1_GotFocus()
   Call gs_SelecTodo(txt_Telef2_1)
End Sub

Private Sub txt_Telef2_1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_AlqMen_1)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "-")
   End If
End Sub

Private Sub txt_Direcc_2_GotFocus()
   Call gs_SelecTodo(txt_Direcc_2)
End Sub

Private Sub txt_Direcc_2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NomArr_2)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "-_ ,.;:º#()/")
   End If
End Sub

Private Sub txt_NomArr_2_GotFocus()
   Call gs_SelecTodo(txt_NomArr_2)
End Sub

Private Sub txt_NomArr_2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_IniAlq_2)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "-_ ,.;:º#()/")
   End If
End Sub

Private Sub txt_Telef1_2_GotFocus()
   Call gs_SelecTodo(txt_Telef1_2)
End Sub

Private Sub txt_Telef1_2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Telef2_2)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "-")
   End If
End Sub

Private Sub txt_Telef2_2_GotFocus()
   Call gs_SelecTodo(txt_Telef2_2)
End Sub

Private Sub txt_Telef2_2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_AlqMen_2)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "-")
   End If
End Sub

Private Sub fs_Inicio()
   Call moddat_gs_Carga_LisIte_Combo(cmb_MonIng, 1, "113")
   Call moddat_gs_Carga_LisIte_Combo(cmb_SegPro, 1, "214")
End Sub

Private Sub fs_Limpia()
   cmb_MonIng.ListIndex = -1
   ipp_IngNet.Value = 0
   
   txt_Direcc_1.Text = ""
   txt_NomArr_1.Text = ""
   ipp_IniAlq_1.Text = Format(date, "dd/mm/yyyy")
   txt_Telef1_1.Text = ""
   txt_Telef2_1.Text = ""
   ipp_AlqMen_1.Value = 0
   
   cmb_SegPro.ListIndex = -1
   
   txt_Direcc_2.Text = ""
   txt_NomArr_2.Text = ""
   ipp_IniAlq_2.Text = Format(date, "dd/mm/yyyy")
   txt_Telef1_2.Text = ""
   txt_Telef2_2.Text = ""
   ipp_AlqMen_2.Value = 0
   
   txt_Direcc_2.Enabled = False
   txt_NomArr_2.Enabled = False
   ipp_IniAlq_2.Enabled = False
   txt_Telef1_2.Enabled = False
   txt_Telef2_2.Enabled = False
   ipp_AlqMen_2.Enabled = False
End Sub


