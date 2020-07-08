VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_Ges_TecPro_04 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12795
   Icon            =   "OpeTra_frm_823.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   12795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel2 
      Height          =   6495
      Left            =   30
      TabIndex        =   25
      Top             =   30
      Width           =   12765
      _Version        =   65536
      _ExtentX        =   22516
      _ExtentY        =   11456
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
      Begin Threed.SSPanel SSPanel3 
         Height          =   2805
         Left            =   30
         TabIndex        =   26
         Top             =   3630
         Width           =   12675
         _Version        =   65536
         _ExtentX        =   22357
         _ExtentY        =   4948
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
         Begin VB.ComboBox cmb_TipLin 
            Height          =   315
            Left            =   10200
            Style           =   2  'Dropdown List
            TabIndex        =   68
            Top             =   2640
            Visible         =   0   'False
            Width           =   2175
         End
         Begin Threed.SSCheck chk_NesCli 
            Height          =   255
            Left            =   8520
            TabIndex        =   4
            Top             =   150
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   "No Cliente"
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
         Begin VB.ComboBox cmb_PorGar 
            Height          =   315
            ItemData        =   "OpeTra_frm_823.frx":000C
            Left            =   10200
            List            =   "OpeTra_frm_823.frx":000E
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   1290
            Width           =   735
         End
         Begin VB.ComboBox cmb_NomPry 
            Height          =   315
            Left            =   1710
            TabIndex        =   21
            Text            =   "cmb_NomPry"
            Top             =   2400
            Width           =   6495
         End
         Begin VB.TextBox txt_CodEte 
            Height          =   315
            Left            =   10200
            MaxLength       =   16
            TabIndex        =   7
            Top             =   510
            Width           =   2175
         End
         Begin VB.TextBox txt_ParReg 
            Height          =   315
            Left            =   10200
            MaxLength       =   10
            TabIndex        =   20
            Top             =   2040
            Width           =   2175
         End
         Begin VB.TextBox txt_CodPry 
            Height          =   315
            Left            =   6000
            MaxLength       =   10
            TabIndex        =   19
            Top             =   2040
            Width           =   2175
         End
         Begin VB.TextBox txt_NumAde 
            Height          =   315
            Left            =   1710
            MaxLength       =   10
            TabIndex        =   18
            Top             =   2040
            Width           =   2175
         End
         Begin VB.ComboBox cmb_TipRen 
            Height          =   315
            Left            =   10200
            Style           =   2  'Dropdown List
            TabIndex        =   55
            Top             =   120
            Width           =   2175
         End
         Begin VB.ComboBox cmb_Modali 
            Height          =   315
            Left            =   1710
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   120
            Width           =   6495
         End
         Begin VB.TextBox txt_NumRef 
            Height          =   315
            Left            =   1710
            MaxLength       =   10
            TabIndex        =   5
            Top             =   510
            Width           =   2175
         End
         Begin VB.ComboBox cmb_Moneda 
            Height          =   315
            Left            =   1710
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   1290
            Width           =   2175
         End
         Begin EditLib.fpDateTime ipp_FecVct 
            Height          =   315
            Left            =   10200
            TabIndex        =   10
            Top             =   900
            Width           =   2175
            _Version        =   196608
            _ExtentX        =   3836
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
         Begin EditLib.fpDoubleSingle ipp_ImpCar 
            Height          =   315
            Left            =   6000
            TabIndex        =   12
            Top             =   1290
            Width           =   2175
            _Version        =   196608
            _ExtentX        =   3836
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
         Begin EditLib.fpDoubleSingle ipp_TEACom 
            Height          =   315
            Left            =   1710
            TabIndex        =   15
            Top             =   1680
            Width           =   2175
            _Version        =   196608
            _ExtentX        =   3836
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
         Begin Threed.SSPanel pnl_ValGar 
            Height          =   315
            Left            =   10950
            TabIndex        =   14
            Top             =   1290
            Width           =   1420
            _Version        =   65536
            _ExtentX        =   2505
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00  "
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
         Begin Threed.SSPanel pnl_ValCom 
            Height          =   315
            Left            =   6000
            TabIndex        =   16
            Top             =   1680
            Width           =   2175
            _Version        =   65536
            _ExtentX        =   3836
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00  "
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
         Begin Threed.SSPanel pnl_ValMin 
            Height          =   315
            Left            =   10200
            TabIndex        =   17
            Top             =   1680
            Width           =   2175
            _Version        =   65536
            _ExtentX        =   3836
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00  "
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
         Begin EditLib.fpLongInteger ipp_PlaCar 
            Height          =   315
            HelpContextID   =   3
            Left            =   6000
            TabIndex        =   9
            Top             =   900
            Width           =   2175
            _Version        =   196608
            _ExtentX        =   3836
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
            MaxValue        =   "360"
            MinValue        =   "77"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
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
         Begin EditLib.fpDoubleSingle ipp_PorRet 
            Height          =   315
            Left            =   6000
            TabIndex        =   6
            Top             =   510
            Width           =   2175
            _Version        =   196608
            _ExtentX        =   3836
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
            Left            =   1710
            TabIndex        =   8
            Top             =   900
            Width           =   2175
            _Version        =   196608
            _ExtentX        =   3836
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
         Begin EditLib.fpDoubleSingle ipp_PorTEA 
            Height          =   315
            Left            =   6150
            TabIndex        =   63
            Top             =   510
            Visible         =   0   'False
            Width           =   2175
            _Version        =   196608
            _ExtentX        =   3836
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
         Begin EditLib.fpDoubleSingle ipp_TasMor 
            Height          =   315
            Left            =   10200
            TabIndex        =   22
            Top             =   2400
            Visible         =   0   'False
            Width           =   2175
            _Version        =   196608
            _ExtentX        =   3836
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
         Begin EditLib.fpDoubleSingle ipp_ValCom 
            Height          =   315
            Left            =   6120
            TabIndex        =   65
            Top             =   1680
            Visible         =   0   'False
            Width           =   2175
            _Version        =   196608
            _ExtentX        =   3836
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
         Begin VB.Label lbl_TipLin 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Línea:"
            Height          =   195
            Left            =   8520
            TabIndex        =   69
            Top             =   2715
            Visible         =   0   'False
            Width           =   825
         End
         Begin VB.Label lbl_TasMor 
            AutoSize        =   -1  'True
            Caption         =   "Tasa Int. Moratorio:"
            Height          =   195
            Left            =   8520
            TabIndex        =   64
            Top             =   2460
            Visible         =   0   'False
            Width           =   1380
         End
         Begin VB.Label lbl_NomPry 
            AutoSize        =   -1  'True
            Caption         =   "Nombre Proyecto :"
            Height          =   195
            Left            =   90
            TabIndex        =   62
            Top             =   2460
            Width           =   1320
         End
         Begin VB.Label Lbl_CodEte 
            AutoSize        =   -1  'True
            Caption         =   "Código de ET.:"
            Height          =   195
            Left            =   8520
            TabIndex        =   60
            Top             =   540
            Width           =   1065
         End
         Begin VB.Label lbl_ParReg 
            Caption         =   "Partida Registral:"
            Height          =   225
            Left            =   8520
            TabIndex        =   59
            Top             =   2085
            Width           =   1545
         End
         Begin VB.Label lbl_Codpry 
            Caption         =   "Código de Proyecto:"
            Height          =   225
            Left            =   4290
            TabIndex        =   58
            Top             =   2085
            Width           =   1545
         End
         Begin VB.Label lbl_NumAde 
            Caption         =   "Número de Adenda:"
            Height          =   225
            Left            =   90
            TabIndex        =   57
            Top             =   2085
            Width           =   1545
         End
         Begin VB.Label lbl_TipRen 
            Caption         =   "Tipo de Renovación:"
            Height          =   255
            Left            =   8520
            TabIndex        =   56
            Top             =   150
            Width           =   1545
         End
         Begin VB.Label lbl_Modali 
            Caption         =   "Modalidad:"
            Height          =   255
            Left            =   90
            TabIndex        =   54
            Top             =   150
            Width           =   1275
         End
         Begin VB.Label lbl_PorRet 
            Caption         =   "Porc. Retención:"
            Height          =   225
            Left            =   4290
            TabIndex        =   37
            Top             =   540
            Width           =   1635
         End
         Begin VB.Label Label16 
            Caption         =   "Fecha Emisión:"
            Height          =   255
            Left            =   90
            TabIndex        =   36
            Top             =   930
            Width           =   1245
         End
         Begin VB.Label lbl_ValMin 
            Caption         =   "Valor Mínimo 3M:"
            Height          =   315
            Left            =   8520
            TabIndex        =   35
            Top             =   1725
            Width           =   1545
         End
         Begin VB.Label lbl_ValCom 
            Caption         =   "Valor Comisión:"
            Height          =   315
            Left            =   4290
            TabIndex        =   34
            Top             =   1725
            Width           =   1545
         End
         Begin VB.Label lbl_TeaCom 
            Caption         =   "TEA% Comisión:"
            Height          =   225
            Left            =   90
            TabIndex        =   33
            Top             =   1725
            Width           =   1545
         End
         Begin VB.Label Label2 
            Caption         =   "Nro. Referencia:"
            Height          =   255
            Left            =   90
            TabIndex        =   32
            Top             =   540
            Width           =   1245
         End
         Begin VB.Label Label8 
            Caption         =   "Plazo (Días):"
            Height          =   315
            Left            =   4290
            TabIndex        =   31
            Top             =   930
            Width           =   1245
         End
         Begin VB.Label Label3 
            Caption         =   "F. Vencimiento:"
            Height          =   315
            Left            =   8520
            TabIndex        =   30
            Top             =   930
            Width           =   1545
         End
         Begin VB.Label Label4 
            Caption         =   "Moneda:"
            Height          =   225
            Left            =   90
            TabIndex        =   29
            Top             =   1335
            Width           =   1545
         End
         Begin VB.Label lbl_NomVal 
            Caption         =   "Valor:"
            Height          =   315
            Left            =   4290
            TabIndex        =   28
            Top             =   1335
            Width           =   1545
         End
         Begin VB.Label lbl_ValGar 
            Caption         =   "Garantizado:"
            Height          =   315
            Left            =   8520
            TabIndex        =   27
            Top             =   1335
            Width           =   1545
         End
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   675
         Left            =   30
         TabIndex        =   38
         Top             =   750
         Width           =   12675
         _Version        =   65536
         _ExtentX        =   22357
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
         Begin VB.CommandButton cmd_NueGar 
            Height          =   585
            Left            =   600
            Picture         =   "OpeTra_frm_823.frx":0010
            Style           =   1  'Graphical
            TabIndex        =   67
            ToolTipText     =   "Garantía"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_NueTas 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_823.frx":031A
            Style           =   1  'Graphical
            TabIndex        =   66
            ToolTipText     =   "Registrar Informe de Tasación"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   12060
            Picture         =   "OpeTra_frm_823.frx":0BE4
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   11490
            Picture         =   "OpeTra_frm_823.frx":1026
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel11 
         Height          =   945
         Left            =   30
         TabIndex        =   39
         Top             =   2640
         Width           =   12675
         _Version        =   65536
         _ExtentX        =   22357
         _ExtentY        =   1667
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
         Begin VB.ComboBox cmb_TipRec 
            Height          =   315
            Left            =   10200
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   480
            Width           =   2175
         End
         Begin VB.ComboBox cmb_SubPrd 
            Height          =   315
            Left            =   1710
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   510
            Width           =   6495
         End
         Begin VB.ComboBox cmb_Produc 
            Height          =   315
            Left            =   1710
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   120
            Width           =   10660
         End
         Begin VB.Label Lbl_TipRec 
            AutoSize        =   -1  'True
            Caption         =   "Recurso:"
            Height          =   195
            Left            =   8520
            TabIndex        =   61
            Top             =   555
            Width           =   645
         End
         Begin VB.Label Label6 
            Caption         =   "Sub-Producto:"
            Height          =   225
            Left            =   120
            TabIndex        =   53
            Top             =   555
            Width           =   1395
         End
         Begin VB.Label Label1 
            Caption         =   "Producto:"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   165
            Width           =   1275
         End
         Begin VB.Line Line1 
            X1              =   3900
            X2              =   3930
            Y1              =   4710
            Y2              =   4740
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   41
         Top             =   30
         Width           =   12675
         _Version        =   65536
         _ExtentX        =   22357
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
            TabIndex        =   42
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
         Begin Threed.SSPanel pnl_Descri 
            Height          =   315
            Left            =   630
            TabIndex        =   43
            Top             =   330
            Width           =   4215
            _Version        =   65536
            _ExtentX        =   7435
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Techo Propio - Registro"
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
            Picture         =   "OpeTra_frm_823.frx":1468
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   1125
         Left            =   30
         TabIndex        =   44
         Top             =   1470
         Width           =   12675
         _Version        =   65536
         _ExtentX        =   22357
         _ExtentY        =   1984
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
            Left            =   1710
            TabIndex        =   45
            Top             =   450
            Width           =   5625
            _Version        =   65536
            _ExtentX        =   9922
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
         Begin Threed.SSPanel pnl_TipDoc 
            Height          =   315
            Left            =   1710
            TabIndex        =   46
            Top             =   120
            Width           =   5625
            _Version        =   65536
            _ExtentX        =   9922
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
         Begin Threed.SSPanel pnl_NroDoc 
            Height          =   315
            Left            =   9390
            TabIndex        =   47
            Top             =   120
            Width           =   2895
            _Version        =   65536
            _ExtentX        =   5106
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
         Begin Threed.SSPanel pnl_TipEmp 
            Height          =   315
            Left            =   1710
            TabIndex        =   48
            Top             =   780
            Width           =   2895
            _Version        =   65536
            _ExtentX        =   5106
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
         Begin VB.Label lbl_TipDoc 
            Caption         =   "Tipo Documento:"
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label lbl_NumDoc 
            Caption         =   "Nro. Documento:"
            Height          =   225
            Left            =   7770
            TabIndex        =   51
            Top             =   135
            Width           =   1335
         End
         Begin VB.Label lbl_RazSoc 
            Caption         =   "Razón Social:"
            Height          =   255
            Left            =   150
            TabIndex        =   50
            Top             =   435
            Width           =   1035
         End
         Begin VB.Label lbl_TipEmp 
            Caption         =   "Tipo Empresa:"
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   780
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "frm_Ges_TecPro_04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim r_dbl_ValFia        As Double
Dim l_arr_Produc()      As moddat_tpo_Genera
Dim l_arr_SubPrd()      As moddat_tpo_Genera
Dim l_arr_Modali()      As moddat_tpo_Genera
Dim l_arr_Proyec()      As moddat_tpo_Genera
Dim l_arr_TipRec()      As moddat_tpo_Genera
Dim l_arr_TipLin()      As moddat_tpo_Genera
Dim l_str_FecVct        As String
Dim l_str_VctCal        As String
Dim l_str_NumRef        As String
Dim l_dbl_ValGar        As Double
Dim l_dbl_ValImp        As Double
Dim l_dbl_MtoGar        As Double
Dim l_str_FVeRen        As String
Dim l_str_FEmRen        As String

Private Sub chk_NesCli_Click(Value As Integer)
   If Value = -1 Then
         ipp_PlaCar.MinValue = 0 '31 '90
         ipp_PlaCar.MaxValue = 150 '120 '90
      If moddat_g_int_FlgGrb_1 <> 2 Then
         ipp_PlaCar.Value = 90
         ipp_TEACom.Value = 9
      End If
   Else
      'Plazos (días)
      If moddat_g_int_FlgGrb_1 = 6 Then
         ipp_PlaCar.MinValue = 0 '31 '45
         ipp_PlaCar.MaxValue = 180
      
      Else
         If moddat_g_int_FlgGrb_1 = 1 Then
            ipp_PlaCar.MinValue = 45 '90
            ipp_PlaCar.MaxValue = 360
         Else
            'Verificar si ya ha tenido una renovación
            If fs_Validar_Renovacion(moddat_g_str_DesIte) Then 'moddat_g_str_NumFia
               ipp_PlaCar.MinValue = 0 '31 '45
               ipp_PlaCar.MaxValue = 180
            Else
               ipp_PlaCar.MinValue = 45 '90
               ipp_PlaCar.MaxValue = 360
            End If
         End If
      End If
   End If
   Call fs_Calcular
   Call gs_SetFocus(ipp_PorRet)
End Sub

Private Sub chk_NesCli_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If moddat_g_str_CodMod = "008" Then
         Call gs_SetFocus(ipp_PorRet)
      End If
   End If
End Sub

Private Sub cmb_Modali_Click()
   moddat_g_str_CodMod = ""
   
'   'Plazos (días)
'   If moddat_g_int_FlgGrb_1 = 6 Then
'      ipp_PlaCar.MinValue = 31
'      ipp_PlaCar.MaxValue = 180
'
'   Else
'      If moddat_g_int_FlgGrb_1 = 1 Then
'         ipp_PlaCar.MinValue = 45
'         ipp_PlaCar.MaxValue = 360
'      Else
'         'Verificar si ya ha tenido una renovación
'         If fs_Validar_Renovacion(moddat_g_str_DesIte) Then
'            ipp_PlaCar.MinValue = 31
'            ipp_PlaCar.MaxValue = 180
'         Else
'            ipp_PlaCar.MinValue = 45
'            ipp_PlaCar.MaxValue = 360
'         End If
'      End If
'   End If
   
   If cmb_Modali.ListIndex > -1 Then
      moddat_g_str_CodMod = l_arr_Modali(cmb_Modali.ListIndex + 1).Genera_Codigo
      
      If moddat_g_str_CodPrd = "026" Or moddat_g_str_CodPrd = "027" Then
         
         Call fs_estadocontroles(True)
         
         If moddat_g_str_CodPrd = "026" And moddat_g_str_CodSub = "001" And (moddat_g_str_CodMod = "004" Or moddat_g_str_CodMod = "005") Then
            txt_ParReg.Enabled = True
            txt_CodPry.Enabled = True
            txt_NumAde.Enabled = True
            cmb_NomPry.Enabled = True
         Else
            txt_ParReg.Enabled = False
            txt_CodPry.Enabled = False
            txt_NumAde.Enabled = False
            cmb_NomPry.Enabled = False
         End If
         
         If moddat_g_str_CodMod = "008" Then
            chk_NesCli.Visible = True
            Call gs_SetFocus(chk_NesCli)
         Else
            chk_NesCli.Visible = False
            chk_NesCli.Value = 0
            Call gs_SetFocus(ipp_PorRet)
         End If
         
         ipp_TasMor.Visible = False
         lbl_TasMor.Visible = False
         ipp_ValCom.Visible = False
'         lbl_NomPry.Visible = True
         ipp_PorTEA.Visible = False
         lbl_PorRet.Caption = "Porc. Retención:"
         lbl_NomVal.Caption = "Valor"
         
         lbl_NomPry.Left = 90
         lbl_NomPry.Top = 2460

         cmb_NomPry.Left = 1710
         cmb_NomPry.Top = 2400
         
      ElseIf moddat_g_str_CodPrd = "008" Then
      
         Call gs_SetFocus(ipp_PorTEA)
         Call fs_estadocontroles(False)
         
         cmb_NomPry.Enabled = True
         ipp_TasMor.Visible = True
         lbl_TasMor.Visible = True
         ipp_ValCom.Visible = True
         lbl_ValCom.Visible = True
         lbl_NomPry.Visible = True
         cmb_NomPry.Visible = True
         ipp_PorTEA.Visible = True
         lbl_PorRet.Visible = True
         lbl_PorRet.Caption = "% TEA"
         lbl_NomVal.Caption = "Monto Préstamo"
         
         If moddat_g_str_CodMod = "001" Then
            cmb_TipLin.Visible = True
            lbl_TipLin.Visible = True
            lbl_TipLin.Top = lbl_TipRen.Top
            cmb_TipLin.Top = cmb_TipRen.Top
         Else
            cmb_TipLin.Visible = False
            lbl_TipLin.Visible = False
         End If
         
         ipp_ValCom.Left = 6000
         ipp_ValCom.Top = 1680
         
         lbl_TasMor.Left = 90
         lbl_TasMor.Top = 1680

         ipp_TasMor.Left = 1710
         ipp_TasMor.Top = 1680
         
         lbl_NomPry.Left = 90
         lbl_NomPry.Top = 2085

         cmb_NomPry.Left = 1710
         cmb_NomPry.Top = 2040
         
         ipp_PorTEA.Left = 6000
                  
      End If
   End If
End Sub
Private Sub fs_estadocontroles(ByVal p_Estado As Boolean)
   If moddat_g_int_FlgGrb_1 = 6 Then
      lbl_TipRen.Visible = p_Estado
      cmb_TipRen.Visible = p_Estado
   End If
'   lbl_PorRet.Visible = p_Estado
   ipp_PorRet.Visible = p_Estado
   lbl_ValGar.Visible = p_Estado
   pnl_ValGar.Visible = p_Estado
   cmb_PorGar.Visible = p_Estado
   lbl_TeaCom.Visible = p_Estado
   ipp_TEACom.Visible = p_Estado
   lbl_ValCom.Visible = p_Estado
   pnl_ValCom.Visible = p_Estado
   lbl_ValMin.Visible = p_Estado
   pnl_ValMin.Visible = p_Estado
   lbl_NumAde.Visible = p_Estado
   txt_NumAde.Visible = p_Estado
   lbl_Codpry.Visible = p_Estado
   txt_CodPry.Visible = p_Estado
   lbl_ParReg.Visible = p_Estado
   txt_ParReg.Visible = p_Estado
   txt_CodEte.Visible = p_Estado
   Lbl_CodEte.Visible = p_Estado
   
   chk_NesCli.Visible = p_Estado
   
'   lbl_NomPry.Visible = p_Estado
'   cmb_NomPry.Visible = p_Estado
End Sub
Private Sub cmb_Modali_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Modali_Click
   End If
End Sub

Private Sub cmb_NomPry_Click()
    Call gs_SetFocus(cmd_Grabar)
End Sub

Private Sub cmb_NomPry_GotFocus()
   Call SendMessage(cmb_NomPry.hwnd, CB_SHOWDROPDOWN, 1, 0&)
End Sub

Private Sub cmb_NomPry_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_NomPry_Click
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   End If
End Sub

Private Sub cmb_NomPry_LostFocus()
   Call SendMessage(cmb_NomPry.hwnd, CB_SHOWDROPDOWN, 0, 0&)
End Sub

Private Sub cmb_PorGar_Click()
   'Call fs_Calcular
   If moddat_g_str_CodPrd <> "008" Then
      Call gs_SetFocus(ipp_TEACom)
   Else
      Call gs_SetFocus(ipp_TasMor)
   End If
End Sub

Private Sub cmb_PorGar_GotFocus()
   Call fs_Calcular
End Sub

Private Sub cmb_PorGar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_PorGar_Click
   End If
End Sub

Private Sub cmb_Produc_Click()
   moddat_g_str_CodPrd = ""
   moddat_g_str_CodSub = ""
   moddat_g_str_CodMod = ""
   
   If cmb_Produc.ListIndex > -1 Then
      Screen.MousePointer = 11
      moddat_g_str_CodPrd = l_arr_Produc(cmb_Produc.ListIndex + 1).Genera_Codigo
      
      'Sub-Producto
      cmb_SubPrd.Clear
      cmb_Modali.Clear
      ReDim l_arr_SubPrd(0)
         
      If moddat_g_str_CodPrd = "026" Then
         
         ReDim Preserve l_arr_SubPrd(UBound(l_arr_SubPrd) + 1)
         l_arr_SubPrd(UBound(l_arr_SubPrd)).Genera_Codigo = Trim$("001")
         l_arr_SubPrd(UBound(l_arr_SubPrd)).Genera_Nombre = Trim$("AVN - ADQUISICION VIVIENDA NUEVA")
         cmb_SubPrd.AddItem Trim$("AVN - AQUISICION VIVIENDA NUEVA")
         
         ReDim Preserve l_arr_SubPrd(UBound(l_arr_SubPrd) + 1)
         l_arr_SubPrd(UBound(l_arr_SubPrd)).Genera_Codigo = Trim$("002")
         l_arr_SubPrd(UBound(l_arr_SubPrd)).Genera_Nombre = Trim$("CSP - CONSTRUCCION SITIO PROPIO")
         cmb_SubPrd.AddItem Trim$("CSP - CONSTRUCCION SITIO PROPIO")
         
         ReDim Preserve l_arr_SubPrd(UBound(l_arr_SubPrd) + 1)
         l_arr_SubPrd(UBound(l_arr_SubPrd)).Genera_Codigo = Trim$("003")
         l_arr_SubPrd(UBound(l_arr_SubPrd)).Genera_Nombre = Trim$("MV - MEJORAMIENTO DE VIVIENDA")
         cmb_SubPrd.AddItem Trim$("MV - MEJORAMIENTO DE VIVIENDA")

      ElseIf moddat_g_str_CodPrd = "027" Then
      
         ReDim Preserve l_arr_SubPrd(UBound(l_arr_SubPrd) + 1)
         l_arr_SubPrd(UBound(l_arr_SubPrd)).Genera_Codigo = Trim$("004")
         l_arr_SubPrd(UBound(l_arr_SubPrd)).Genera_Nombre = Trim$("BPVV - BONO DE REFORZAMIENO ESTRUCTURAL")
         cmb_SubPrd.AddItem Trim$("BPVV - BONO DE REFORZAMIENO ESTRUCTURAL")
      
      ElseIf moddat_g_str_CodPrd = "008" Then
      
         ReDim Preserve l_arr_SubPrd(UBound(l_arr_SubPrd) + 1)
         l_arr_SubPrd(UBound(l_arr_SubPrd)).Genera_Codigo = Trim$("008")
         l_arr_SubPrd(UBound(l_arr_SubPrd)).Genera_Nombre = Trim$("CREDITO AL CONSTRUCTOR")
         cmb_SubPrd.AddItem Trim$("CREDITO AL CONSTRUCTOR")
      End If
      
      Call gs_SetFocus(cmb_SubPrd)
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmb_Produc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Produc_Click
   End If
End Sub

Private Sub cmb_SubPrd_Click()
   moddat_g_str_CodSub = ""
   
   If cmb_SubPrd.ListIndex > -1 Then
      moddat_g_str_CodSub = l_arr_SubPrd(cmb_SubPrd.ListIndex + 1).Genera_Codigo
        
      'Modalidad
      cmb_Modali.Clear
      ReDim l_arr_Modali(0)
      
      If moddat_g_str_CodPrd = "026" Then
      
'         ReDim Preserve l_arr_Modali(UBound(l_arr_Modali) + 1)
'         l_arr_Modali(UBound(l_arr_Modali)).Genera_Codigo = Trim$("001")
'         l_arr_Modali(UBound(l_arr_Modali)).Genera_Nombre = Trim$("BONO DE EMERGENCIA")
'         cmb_Modali.AddItem Trim$("BONO DE EMERGENCIA")
'
'         ReDim Preserve l_arr_Modali(UBound(l_arr_Modali) + 1)
'         l_arr_Modali(UBound(l_arr_Modali)).Genera_Codigo = Trim$("003")
'         l_arr_Modali(UBound(l_arr_Modali)).Genera_Nombre = Trim$("BONO NORMAL")
'         cmb_Modali.AddItem Trim$("BONO NORMAL")
         If moddat_g_str_CodSub = "001" Then
            ReDim Preserve l_arr_Modali(UBound(l_arr_Modali) + 1)
            l_arr_Modali(UBound(l_arr_Modali)).Genera_Codigo = Trim$("004")
            l_arr_Modali(UBound(l_arr_Modali)).Genera_Nombre = Trim$("AVN - ADQUISICION VIVIENDA NUEVA - CF")
            cmb_Modali.AddItem Trim$("AVN - ADQUISICION VIVIENDA NUEVA - CF")
         
            ReDim Preserve l_arr_Modali(UBound(l_arr_Modali) + 1)
            l_arr_Modali(UBound(l_arr_Modali)).Genera_Codigo = Trim$("005")
            l_arr_Modali(UBound(l_arr_Modali)).Genera_Nombre = Trim$("AVN - ADQUISICION VIVIENDA NUEVA - AD")
            cmb_Modali.AddItem Trim$("AVN - ADQUISICION VIVIENDA NUEVA - AD")
         
            'CSO
            ReDim Preserve l_arr_Modali(UBound(l_arr_Modali) + 1)
            l_arr_Modali(UBound(l_arr_Modali)).Genera_Codigo = Trim$("008")
            l_arr_Modali(UBound(l_arr_Modali)).Genera_Nombre = Trim$("CSO - CARTA DE SERIEDAD DE OFERTA")
            cmb_Modali.AddItem Trim$("CSO - CARTA DE SERIEDAD DE OFERTA")
           
         ElseIf moddat_g_str_CodSub = "002" Then
            ReDim Preserve l_arr_Modali(UBound(l_arr_Modali) + 1)
            l_arr_Modali(UBound(l_arr_Modali)).Genera_Codigo = Trim$("006")
            l_arr_Modali(UBound(l_arr_Modali)).Genera_Nombre = Trim$("CSP - CONSTRUCCION SITIO PROPIO")
            cmb_Modali.AddItem Trim$("CSP - CONSTRUCCION SITIO PROPIO")
            
            'CSO
            ReDim Preserve l_arr_Modali(UBound(l_arr_Modali) + 1)
            l_arr_Modali(UBound(l_arr_Modali)).Genera_Codigo = Trim$("008")
            l_arr_Modali(UBound(l_arr_Modali)).Genera_Nombre = Trim$("CSO - CARTA DE SERIEDAD DE OFERTA")
            cmb_Modali.AddItem Trim$("CSO - CARTA DE SERIEDAD DE OFERTA")
            
         ElseIf moddat_g_str_CodSub = "003" Then
            ReDim Preserve l_arr_Modali(UBound(l_arr_Modali) + 1)
            l_arr_Modali(UBound(l_arr_Modali)).Genera_Codigo = Trim$("007")
            l_arr_Modali(UBound(l_arr_Modali)).Genera_Nombre = Trim$("CMV - MEJORAMIENTO DE VIVIENDA")
            cmb_Modali.AddItem Trim$("CMV - MEJORAMIENTO DE VIVIENDA")
            
            'CSO
            ReDim Preserve l_arr_Modali(UBound(l_arr_Modali) + 1)
            l_arr_Modali(UBound(l_arr_Modali)).Genera_Codigo = Trim$("008")
            l_arr_Modali(UBound(l_arr_Modali)).Genera_Nombre = Trim$("CSO - CARTA DE SERIEDAD DE OFERTA")
            cmb_Modali.AddItem Trim$("CSO - CARTA DE SERIEDAD DE OFERTA")
         End If
         
      ElseIf moddat_g_str_CodPrd = "027" Then 'And moddat_g_str_CodSub = ""
         
         ReDim Preserve l_arr_Modali(UBound(l_arr_Modali) + 1)
         l_arr_Modali(UBound(l_arr_Modali)).Genera_Codigo = Trim$("002")
         l_arr_Modali(UBound(l_arr_Modali)).Genera_Nombre = Trim$("BPVV - BONO DE REFORZAMIENO ESTRUCTURAL")
         cmb_Modali.AddItem Trim$("BPVV - BONO DE REFORZAMIENO ESTRUCTURAL")
         
         ReDim Preserve l_arr_Modali(UBound(l_arr_Modali) + 1)
         l_arr_Modali(UBound(l_arr_Modali)).Genera_Codigo = Trim$("008")
         l_arr_Modali(UBound(l_arr_Modali)).Genera_Nombre = Trim$("CSO - CARTA DE SERIEDAD DE OFERTA")
         cmb_Modali.AddItem Trim$("CSO - CARTA DE SERIEDAD DE OFERTA")
           
      ElseIf moddat_g_str_CodPrd = "008" Then
      
         If moddat_g_str_CodSub = "008" Then
        
            ReDim Preserve l_arr_Modali(UBound(l_arr_Modali) + 1)
            l_arr_Modali(UBound(l_arr_Modali)).Genera_Codigo = Trim$("001")
            l_arr_Modali(UBound(l_arr_Modali)).Genera_Nombre = Trim$("LINEA DE CREDITO")
            cmb_Modali.AddItem Trim$("LINEA DE CREDITO")
            
            ReDim Preserve l_arr_Modali(UBound(l_arr_Modali) + 1)
            l_arr_Modali(UBound(l_arr_Modali)).Genera_Codigo = Trim$("002")
            l_arr_Modali(UBound(l_arr_Modali)).Genera_Nombre = Trim$("CREDITO PUNTUAL")
            cmb_Modali.AddItem Trim$("CREDITO PUNTUAL")
         End If
      End If
         
      If moddat_g_str_CodPrd <> "008" Then
         Lbl_TipRec.Visible = True
         cmb_TipRec.Visible = True
         cmb_TipRec.ListIndex = -1
         Call gs_SetFocus(cmb_TipRec)
         cmb_TipLin.Visible = False
         lbl_TipLin.Visible = False
      Else
         Lbl_TipRec.Visible = False
         cmb_TipRec.Visible = False
         lbl_TipLin.Top = Lbl_TipRec.Top
         cmb_TipLin.Top = cmb_TipRec.Top
         cmb_TipLin.Visible = False
         lbl_TipLin.Visible = False
         Call gs_SetFocus(cmb_Modali)
      End If
    End If
End Sub

Private Sub cmb_SubPrd_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_SubPrd_Click
   End If
End Sub

Private Sub cmb_TipLin_Click()
   Call gs_SetFocus(ipp_PorTEA)
End Sub

Private Sub cmb_TipLin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipLin_Click
   End If
End Sub

Private Sub cmb_TipRec_Click()
    Call gs_SetFocus(cmb_Modali)
End Sub

Private Sub cmb_TipRec_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipRec_Click
   End If
End Sub

Private Sub cmb_TipRen_Click()

moddat_g_int_TipRec = cmb_TipRen.ItemData(cmb_TipRen.ListIndex)

   If moddat_g_int_FlgGrb_1 = 6 Then
     If cmb_TipRen.ItemData(cmb_TipRen.ListIndex) = 1 Then
         ipp_FecEmi.Enabled = True
         ipp_PlaCar.Enabled = True
         ipp_FecVct.Enabled = True
         cmd_Grabar.Enabled = True
         ipp_ImpCar.Enabled = True
         ipp_TEACom.Enabled = True
         cmb_PorGar.Enabled = True
        
         ipp_FecEmi.Text = Format(gf_FormatoFecha(CStr(l_str_FVeRen)), "dd/mm/yyyy")
         ipp_FecVct.Text = Format(CDate(DateAdd("D", ipp_PlaCar.Text, ipp_FecEmi.Text)), "DD/MM/YYYY")
         
         Call gs_SetFocus(ipp_FecEmi)
         
     ElseIf cmb_TipRen.ItemData(cmb_TipRen.ListIndex) = 2 Then
         ipp_FecEmi.Enabled = False
         ipp_PlaCar.Enabled = False
         ipp_FecVct.Enabled = False
         cmd_Grabar.Enabled = True
         ipp_ImpCar.Enabled = True
         ipp_TEACom.Enabled = True
         cmb_PorGar.Enabled = True
        
         ipp_FecEmi.Text = Format(date, "dd/mm/yyyy")
         ipp_FecVct.Text = Format(gf_FormatoFecha(CStr(l_str_FVeRen)), "dd/mm/yyyy")
         ipp_PlaCar.Value = DateDiff("d", ipp_FecEmi.Text, ipp_FecVct.Text)
'         ipp_FecEmi.Text = Format(gf_FormatoFecha(CStr(l_str_FecVct)), "dd/mm/yyyy")
'         ipp_FecVct.Text = Format(ipp_FecEmi.Value, "DD/MM/YYYY")
         Call fs_Calcular
         
         Call gs_SetFocus(ipp_ImpCar)
     End If
   End If
   If ipp_PorRet.Text = "0.00%" Then
      ipp_PorRet.Enabled = True
   End If
   
   
End Sub

Private Sub cmb_TipRen_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipRen_Click
   End If
End Sub

Private Sub cmd_Grabar_Click()
Dim r_str_CtaCon     As String
Dim r_dbl_ImpTot     As Double
Dim r_dbl_ImpSal     As Double
Dim r_dbl_ImpPag     As Double
Dim r_str_CtaDeb     As String
Dim r_str_CtaHab     As String
Dim r_int_NumIte     As Integer
Dim r_dbl_SalFon     As Double
Dim r_dbl_SalDes     As Double
   
   'Valida ingreso de informacion
   If moddat_g_str_CodPrd <> "008" Then
      If cmb_Produc.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Producto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_Produc)
         Exit Sub
      End If
      If cmb_SubPrd.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Sub-Producto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_SubPrd)
         Exit Sub
      End If
      If moddat_g_int_FlgGrb_1 <> 6 Then
         If moddat_g_str_CodMod <> "005" And moddat_g_str_CodMod <> "008" Then
           If cmb_TipRec.ListIndex = -1 Then
              MsgBox "Debe seleccionar el Tipo de Recurso.", vbExclamation, modgen_g_str_NomPlt
              Call gs_SetFocus(cmb_TipRec)
              Exit Sub
           End If
         End If
      End If
      '   If Len(Trim(txt_NumRef.Text)) = 0 Then
      '      MsgBox "Debe ingresar el Número de Referencia.", vbExclamation, modgen_g_str_NomPlt
      '      Call gs_SetFocus(txt_NumRef)
      '      Exit Sub
      '   End If
      '   If Trim(Year(CDate(ipp_FecEmi.Text))) <> Trim(Year(date)) And Year(CDate(ipp_FecEmi.Text)) + 1 <> Year(date) Then
      '      MsgBox "Debe ingresar una Fecha de Emisión válida.", vbExclamation, modgen_g_str_NomPlt
      '      Call gs_SetFocus(ipp_FecEmi)
      '      Exit Sub
      '   End If
      If (moddat_g_str_CodPrd = "026" Or moddat_g_str_CodPrd = "027") And moddat_g_str_CodSub <> "001" And moddat_g_str_CodMod <> "008" Then
         If Len(Trim(txt_CodEte.Text)) = 0 Then
            MsgBox "Debe ingresar Código de ETE.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(txt_CodEte)
            Exit Sub
         End If
      End If
      If moddat_g_int_FlgGrb_1 <> 6 Then
         If (Format(ipp_FecEmi.Text, "yyyymmdd") < Format(moddat_g_str_FecIni, "yyyymmdd") Or _
             Format(ipp_FecEmi.Text, "yyyymmdd") > Format(moddat_g_str_FecFin, "yyyymmdd")) Then
             If moddat_g_str_CodMod = "005" Then
               MsgBox "No es posible registrar la Adenda en un periodo errado.", vbExclamation, modgen_g_str_NomPlt
             ElseIf moddat_g_str_CodMod = "008" Then
               MsgBox "No es posible registrar la Carta Seriedad de Oferta en un periodo errado.", vbExclamation, modgen_g_str_NomPlt
             Else
               MsgBox "No es posible registrar la Carta Fianza en un periodo errado.", vbExclamation, modgen_g_str_NomPlt
             End If
             Call gs_SetFocus(ipp_FecEmi)
             Exit Sub
         End If
      Else
         'Verifica que se haya liberado la retención de garantía líquida antes de renovar
         If moddat_g_int_FlgGrb_1 = 6 Then '
            If fs_Validar_Retencion_Garantia = False Then
               MsgBox "Se debe liberar la Retención de Garantía Líquida.", vbExclamation, modgen_g_str_NomPlt
               Call gs_SetFocus(cmd_Salida)
               Exit Sub
            End If
         End If
      End If
      If ipp_PlaCar.Value = 0 Then
         MsgBox "Debe seleccionar plazo.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_PlaCar)
         Exit Sub
      End If
      If moddat_g_int_FlgGrb_1 <> 6 Then
         If Not (CInt(ipp_PlaCar.Text) >= ipp_PlaCar.MinValue And CInt(ipp_PlaCar.Text) <= ipp_PlaCar.MaxValue) Then
            MsgBox "El Plazo está fuera del rango permitido (Entre " & CStr(ipp_PlaCar.MinValue) & " y " & CStr(ipp_PlaCar.MaxValue) & " días).", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_PlaCar)
            Exit Sub
         End If
      Else
         If Not (CInt(ipp_PlaCar.Text) >= ipp_PlaCar.MinValue And CInt(ipp_PlaCar.Text) <= ipp_PlaCar.MaxValue) Then
            MsgBox "El Plazo de renovación está fuera del rango permitido (Entre " & CStr(ipp_PlaCar.MinValue) & " y " & CStr(ipp_PlaCar.MaxValue) & " días).", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_PlaCar)
            Exit Sub
         End If
      End If
      If moddat_g_int_FlgGrb_1 <> 6 Then
         If Len((l_str_FecVct)) > 0 Then
            If Format(ipp_FecEmi.Text, "yyyymmdd") > l_str_FecVct Then
               MsgBox "Fecha de Emisión sobrepasa el periodo de vigencia (" & Format(gf_FormatoFecha(CStr(l_str_FecVct)), "dd/mm/yyyy") & ") de Entidad Técnica.", vbExclamation, modgen_g_str_NomPlt
               Call gs_SetFocus(ipp_FecVct)
               Exit Sub
            End If
         Else
            MsgBox "Debe ingresar Periodo de Vigencia de Entidad Técnica.", vbExclamation, modgen_g_str_NomPlt
            Exit Sub
         End If
      End If
      
      If cmb_Moneda.ListIndex = -1 Then
         MsgBox "Debe seleccionar Moneda.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_Moneda)
         Exit Sub
      End If
      
      If ipp_ImpCar.Value = 0 Then
         MsgBox "Debe ingresar Importe de Carta Fianza.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_ImpCar)
         Exit Sub
      End If
      
      If cmb_PorGar.ListIndex = -1 Then
         If moddat_g_str_CodPrd <> "008" Then
            If moddat_g_str_CodMod <> "008" Then          'CSO no se selecciona el porcentaje
               MsgBox "Debe seleccionar porcentaje de Garantizado.", vbExclamation, modgen_g_str_NomPlt
               Call gs_SetFocus(cmb_PorGar)
               Exit Sub
            End If
         End If
      End If
      
      If pnl_ValGar.Caption = 0 Then
         MsgBox "Debe ingresar Importe Garantizado.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_ImpCar)
         Exit Sub
      End If
      
      If moddat_g_int_TipRec <> 2 Then
         If CDbl(Replace(ipp_TEACom.Value, "%", "")) = 0 Then
            MsgBox "Debe ingresar Porcentaje de Comisión.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_TEACom)
            Exit Sub
         End If
      End If
      
      If moddat_g_int_TipRec <> 2 Then
         If pnl_ValCom.Caption = 0 Then
            MsgBox "El Importe de Comisión no puede ser cero.", vbExclamation, modgen_g_str_NomPlt
            Exit Sub
         End If
           
         If pnl_ValMin.Caption = 0 Then
            MsgBox "El Importe Mínimo no puede ser cero.", vbExclamation, modgen_g_str_NomPlt
            Exit Sub
         End If
      End If
      If moddat_g_int_FlgGrb_1 <> 6 Then
         If Len(txt_CodPry.Text) = 0 Then
             If moddat_g_str_CodPrd = "026" And moddat_g_str_CodSub = "001" And moddat_g_str_CodMod <> "008" Then
                 MsgBox "Debe ingresar Código de Proyecto.", vbExclamation, modgen_g_str_NomPlt
                 Call gs_SetFocus(txt_CodPry)
                 Exit Sub
             End If
         End If
         If cmb_NomPry.ListIndex = -1 Then
             If moddat_g_str_CodPrd = "026" And moddat_g_str_CodSub = "001" And moddat_g_str_CodMod <> "008" And moddat_g_str_CodMod <> "005" Then
                 MsgBox "Debe seleccionar Proyecto.", vbExclamation, modgen_g_str_NomPlt
                 Call gs_SetFocus(cmb_NomPry)
                 Exit Sub
             End If
         End If
      End If
      If moddat_g_str_CodMod <> "008" Then
         If CDbl(Replace(ipp_PorRet.Value, "%", "")) = 0 Then
            MsgBox "Debe ingresar Porcentaje de Retención.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_PorRet)
            Exit Sub
         End If
      End If
      
      'Valida que el Monto Total de Cartas Fianzas, Adendas, CSO y Cred. Directos sea menor e igual a la Línea Asignada Total
      If fs_Validar_Mto_CarFia(CDbl(pnl_ValGar.Caption), 0) = False Then
         If moddat_g_str_CodMod = "005" Then
           MsgBox "El valor ingresado de Adenda excede Línea Asignada.", vbExclamation, modgen_g_str_NomPlt
         ElseIf moddat_g_str_CodMod = "008" Then
           MsgBox "El valor ingresado de Carta Seriedad de Oferta, excede Línea Asignada.", vbExclamation, modgen_g_str_NomPlt
         Else
           MsgBox "El valor ingresado de Carta Fianza, excede Línea Asignada.", vbExclamation, modgen_g_str_NomPlt
         End If
         Call gs_SetFocus(ipp_ImpCar)
         Exit Sub
      End If
      
      'Valida que el Monto Total de Cartas Fianzas, Adendas, CSO sea menor e igual a la Línea Asignada de Créditos Indirectos
      If fs_Validar_Mto_CarFia(CDbl(pnl_ValGar.Caption), 1) = False Then
         MsgBox "El valor ingresado de Carta Fianza, excede Línea Asignada de Créditos Indirectos.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_ImpCar)
         Exit Sub
      End If
      
'      'Valida que el Monto Total de Créditos Directos sea menor e igual a la Línea Asignada de Créditos Directos
'      If fs_Validar_Mto_CarFia(CDbl(ipp_ImpCar.Value), 2) = False Then
'         MsgBox "El valor ingresado de Carta Fianza, excede Línea Asignada de Créditos Directos.", vbExclamation, modgen_g_str_NomPlt
'         Call gs_SetFocus(ipp_ImpCar)
'         Exit Sub
'      End If
      
      'MODIFICACIÓN
      If moddat_g_int_FlgGrb_1 = 2 Then
         'Valida que el Monto de las Garantías no sea mayor al monto de Cartas Fianza
         If fs_Validar_Mto_CFiGar = False Then
            MsgBox "El Valor Carta Fianza, es menor al Monto de Garantía.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_ImpCar)
            Exit Sub
         End If
      
         'Valida que el Monto de las Comisiones no sea mayor al monto ya pagado
         If fs_Validar_Mto_ComPag = False Then
            MsgBox "El Valor Comisión es menor a las Comisiones pagadas.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_ImpCar)
            Exit Sub
         End If
      
         'Valida que el Monto de los Fondos Recibidos no sea mayor al monto ya pagado
         If fs_Validar_Mto_FRePag = False Then
            MsgBox "El Valor de Carta Fianza es menor a los Fondos Recibidos pagados.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_ImpCar)
            Exit Sub
         End If
      End If
      
'      'Valida Nivel de Endeudamiento
'      If moddat_g_str_CodPrd <> "008" Then
'         If moddat_g_int_FlgGrb_1 = 1 Then
'            l_dbl_MtoGar = CDbl(pnl_ValGar.Caption)
'         Else
'            l_dbl_MtoGar = CDbl(pnl_ValGar.Caption) - l_dbl_ValGar
'         End If
'      Else
'         If moddat_g_int_FlgGrb_1 = 1 Then
'            l_dbl_MtoGar = CDbl(ipp_ImpCar.Value)
'         Else
'            l_dbl_MtoGar = CDbl(ipp_ImpCar.Value) - l_dbl_ValImp
'         End If
'      End If
      
'      If l_dbl_MtoGar <> 0 Then
'         If moddat_gf_Consulta_NivelEndeudamiento(Mid(pnl_TipDoc.Caption, 1, 1), pnl_NroDoc.Caption, moddat_g_str_CodMes, moddat_g_str_CodAno, l_dbl_MtoGar) = True Then
'            MsgBox "El Valor de Carta Fianza sobrepasa el nivel de endeudamiento permitido según norma, en Créditos Comerciales y/o Cartas Fianza.", vbExclamation, modgen_g_str_NomPlt
'            Call gs_SetFocus(ipp_ImpCar)
'            Exit Sub
'         End If
'      End If
   
   Else 'CREDITOS DIRECTOS
   
      If cmb_Produc.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Producto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_Produc)
         Exit Sub
      End If
      If cmb_SubPrd.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Sub-Producto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_SubPrd)
         Exit Sub
      End If
      
      If cmb_Modali.ListIndex = -1 Then
         MsgBox "Debe seleccionar Modalidad.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_Modali)
         Exit Sub
      End If
      
      If moddat_g_str_CodMod = "001" Then
         If cmb_TipLin.ListIndex = -1 Then
            MsgBox "Debe seleccionar Tipo de Línea.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmb_TipLin)
            Exit Sub
         End If
      End If
      If CDbl(Replace(ipp_PorTEA.Value, "%", "")) = 0 Then
         MsgBox "Debe ingresar TEA.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_PorTEA)
         Exit Sub
      End If
      
      If (Format(ipp_FecEmi.Text, "yyyymmdd") < Format(moddat_g_str_FecIni, "yyyymmdd") Or Format(ipp_FecEmi.Text, "yyyymmdd") > Format(moddat_g_str_FecFin, "yyyymmdd")) Then
         MsgBox "No es posible registrar el crédito en un periodo errado.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_FecEmi)
         Exit Sub
      End If
      
      If ipp_PlaCar.Value = 0 Then
         MsgBox "Debe seleccionar plazo.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_PlaCar)
         Exit Sub
      End If
      
      If cmb_Moneda.ListIndex = -1 Then
         MsgBox "Debe seleccionar Moneda.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_Moneda)
         Exit Sub
      End If
      
      If ipp_ImpCar.Value = 0 Then
         MsgBox "Debe ingresar Monto del Préstamo.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_ImpCar)
         Exit Sub
      End If
      
      'Valida que el Monto del Préstamo
'      If fs_Validar_Mto_CarFia(CDbl(ipp_ImpCar.Value)) = False Then
'         If moddat_g_str_CodMod = "001" Then
'            MsgBox "El valor ingresado de Línea de Crédito excede Línea Asignada.", vbExclamation, modgen_g_str_NomPlt
''         ElseIf moddat_g_str_CodMod = "002" Then
''            MsgBox "El valor ingresado del Crédito Puntual, excede Línea Asignada.", vbExclamation, modgen_g_str_NomPlt
'            Call gs_SetFocus(ipp_ImpCar)
'            Exit Sub
'         End If
'      End If

      'Valida que el Monto Total de Cartas Fianzas, Adendas, CSO y Cred. Directos sea menor e igual a la Línea Asignada Total
      If fs_Validar_Mto_CarFia(CDbl(ipp_ImpCar.Value), 0) = False Then
         If moddat_g_str_CodMod = "005" Then
           MsgBox "El valor ingresado de Adenda excede Línea Asignada.", vbExclamation, modgen_g_str_NomPlt
         ElseIf moddat_g_str_CodMod = "008" Then
           MsgBox "El valor ingresado de Carta Seriedad de Oferta, excede Línea Asignada.", vbExclamation, modgen_g_str_NomPlt
         Else
           MsgBox "El valor ingresado de Carta Fianza, excede Línea Asignada.", vbExclamation, modgen_g_str_NomPlt
         End If
         Call gs_SetFocus(ipp_ImpCar)
         Exit Sub
      End If
      
      'Valida que el Monto Total de Créditos Directos sea menor e igual a la Línea Asignada de Créditos Directos
      If fs_Validar_Mto_CarFia(CDbl(ipp_ImpCar.Value), 2) = False Then
         MsgBox "Monto de Préstamo ingresado, excede Línea Asignada de Créditos Directos.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_ImpCar)
         Exit Sub
      End If
      
       'Valida que el Monto Total de Créditos Directos Revolventes sea menor e igual a la Línea Asignada de Créditos Directos Revolventes
      If cmb_TipLin.ListIndex <> -1 Then
         If fs_Validar_CreDir_TipLin(CDbl(ipp_ImpCar.Value), CStr(l_arr_TipLin(cmb_TipLin.ListIndex + 1).Genera_Codigo)) = False Then
            MsgBox "Monto de préstamo ingresado, excede Línea Asignada de Créditos Directos " & IIf(CStr(l_arr_TipLin(cmb_TipLin.ListIndex + 1).Genera_Codigo) = 1, "no revolventes", "revolventes"), vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_ImpCar)
            Exit Sub
         End If
      End If
      
      If CDbl(Replace(ipp_TasMor.Value, "%", "")) = 0 Then
         MsgBox "Debe ingresar Tasa Moratoria.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_TasMor)
         Exit Sub
      End If
      
      If ipp_ValCom.Value = 0 Then
         If MsgBox("¿El valor de la Comisión es cero?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
            Call gs_SetFocus(ipp_ValCom)
            Exit Sub
         End If
         'MsgBox "Debe ingresar Valor de Comisión.", vbExclamation, modgen_g_str_NomPlt
         '
      End If
      
      If cmb_NomPry.ListIndex = -1 Then
         MsgBox "No se asignó Proyecto.", vbInformation, modgen_g_str_NomPlt
         'MsgBox "Debe seleccionar Proyecto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_NomPry)
         'Exit Sub
      End If
   End If
   
   'Valida Nivel de Endeudamiento
   If moddat_g_str_CodPrd <> "008" Then
      If moddat_g_int_FlgGrb_1 = 1 Then
         l_dbl_MtoGar = CDbl(pnl_ValGar.Caption)
      Else
         l_dbl_MtoGar = CDbl(pnl_ValGar.Caption) - l_dbl_ValGar
      End If
   Else
      If moddat_g_int_FlgGrb_1 = 1 Then
         l_dbl_MtoGar = CDbl(ipp_ImpCar.Value)
      Else
         l_dbl_MtoGar = CDbl(ipp_ImpCar.Value) - l_dbl_ValImp
      End If
   End If
      
   If l_dbl_MtoGar <> 0 Then
      moddat_g_dbl_TotGar = CDbl(l_dbl_MtoGar)
      'Sobreexposición
      If moddat_gf_Consulta_ExposicionGlobal(Mid(pnl_TipDoc.Caption, 1, 1), pnl_NroDoc.Caption, l_dbl_MtoGar, 0, 0) = True Then ' moddat_g_str_CodMes, moddat_g_str_CodAno,
         MsgBox "Debe ingresar Garantía Hipotecaria o Líquida.", vbExclamation, modgen_g_str_NomPlt
         cmd_NueTas.Enabled = True
         cmd_NueGar.Enabled = True
         Call gs_SetFocus(cmd_NueTas)
         Exit Sub
      End If
    End If
      
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      
      Screen.MousePointer = 11
      Call moddat_gs_FecSis
      
      'Nuevo - Renovación
      If moddat_g_int_FlgGrb_1 = 1 Or moddat_g_int_FlgGrb_1 = 6 Then
          l_str_NumRef = fs_GeneraNumRef
      Else
          l_str_NumRef = moddat_g_str_DesIte 'txt_NumRef.Text
      End If
      
      'Se Obtiene el Saldo de la Referencia antes de ser renovada
      If moddat_g_int_FlgGrb_1 = 6 Then
          Call fs_Obtener_SalRen(Trim(txt_NumRef.Text), CStr(moddat_g_int_TipDoc), CStr(moddat_g_str_NumDoc), r_dbl_SalFon, r_dbl_SalDes)
      End If
        
      'Grabando Información de Carta Fianza
      g_str_Parame = "USP_TPR_MAECFI ("
      g_str_Parame = g_str_Parame & "'" & CStr(moddat_g_str_CodPrd) & "', "
      g_str_Parame = g_str_Parame & "'" & CStr(moddat_g_str_CodSub) & "', "
      g_str_Parame = g_str_Parame & "'" & CStr(moddat_g_str_CodMod) & "', "
      g_str_Parame = g_str_Parame & "'" & CStr(l_str_NumRef) & "', "
      g_str_Parame = g_str_Parame & "'" & Format(ipp_FecEmi.Text, "yyyymmdd") & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_int_TipDoc) & ", "
      g_str_Parame = g_str_Parame & "'" & CStr(moddat_g_str_NumDoc) & "', "
      g_str_Parame = g_str_Parame & CStr(ipp_PlaCar.Value) & ", "
      g_str_Parame = g_str_Parame & "'" & Format(ipp_FecVct.Text, "yyyymmdd") & "', "
      g_str_Parame = g_str_Parame & CStr(cmb_Moneda.ItemData(cmb_Moneda.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ImpCar.Value) & ", "
     
      If moddat_g_str_CodPrd <> "008" Then
         g_str_Parame = g_str_Parame & CStr(CDbl(pnl_ValGar.Caption)) & ", "
         g_str_Parame = g_str_Parame & CStr(ipp_TEACom.Value) & ", "
         g_str_Parame = g_str_Parame & CStr(CDbl(pnl_ValCom.Caption)) & ", "
         g_str_Parame = g_str_Parame & CStr(CDbl(pnl_ValMin.Caption)) & ", "
         g_str_Parame = g_str_Parame & CStr(Replace(ipp_PorRet.Value, "%", "")) & ", "
      Else
         g_str_Parame = g_str_Parame & CStr(CDbl(0)) & ", "
         g_str_Parame = g_str_Parame & CStr(Replace(ipp_TasMor.Value, "%", "")) & ", "
         g_str_Parame = g_str_Parame & CStr(CDbl(ipp_ValCom.Value)) & ", "
         g_str_Parame = g_str_Parame & CStr(0) & ", "
         g_str_Parame = g_str_Parame & CStr(0) & ", "
      End If
        
      If moddat_g_int_FlgGrb_1 <> 6 Then
          g_str_Parame = g_str_Parame & "'', "
      Else
          g_str_Parame = g_str_Parame & "'" & CStr(moddat_g_str_DesIte) & "', "                       'moddat_g_str_NumFia
      End If
      g_str_Parame = g_str_Parame & "'', "                                                            'Fecha de Cancelación
      g_str_Parame = g_str_Parame & "'" & CStr(CStr(txt_NumAde.Text)) & "', "
      
      If moddat_g_str_CodPrd <> "008" Then
         g_str_Parame = g_str_Parame & "'" & CStr(CStr(txt_CodPry.Text)) & "', "
         g_str_Parame = g_str_Parame & "'" & CStr(CStr(txt_ParReg.Text)) & "', "
         g_str_Parame = g_str_Parame & "'" & CStr(CStr(txt_CodEte.Text)) & "', "
         
         If l_arr_TipRec(cmb_TipRec.ListIndex + 1).Genera_Codigo <> "" Then
             g_str_Parame = g_str_Parame & CStr(l_arr_TipRec(cmb_TipRec.ListIndex + 1).Genera_Codigo) & ", "
         Else
             g_str_Parame = g_str_Parame & 0 & ", "
         End If
      Else
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & 0 & ", "
      End If
      
      If l_arr_Proyec(cmb_NomPry.ListIndex + 1).Genera_Codigo <> "" Then
          g_str_Parame = g_str_Parame & "'" & CStr(CStr(l_arr_Proyec(cmb_NomPry.ListIndex + 1).Genera_Codigo)) & "', "
      Else
          g_str_Parame = g_str_Parame & 0 & ", "
      End If
      
      If moddat_g_str_CodPrd <> "008" Then
         g_str_Parame = g_str_Parame & CStr(CDbl(0)) & ", "
      Else
         g_str_Parame = g_str_Parame & CStr(Replace(ipp_PorTEA.Value, "%", "")) & ", "
      End If
      
      If moddat_g_str_CodPrd <> "008" And moddat_g_str_CodMod <> "008" Then
         g_str_Parame = g_str_Parame & CStr(cmb_PorGar.ItemData(cmb_PorGar.ListIndex)) & ", "
      Else
         g_str_Parame = g_str_Parame & CStr(0) & ", "
      End If
      
      If chk_NesCli.Value = False Then
         g_str_Parame = g_str_Parame & CStr(0) & ", "
      Else
         g_str_Parame = g_str_Parame & CStr(1) & ", "
      End If
      
      'Tipo de Línea
      If moddat_g_str_CodPrd <> "008" Then 'And moddat_g_str_CodMod <> "008"
         g_str_Parame = g_str_Parame & CStr(0) & ", "
      Else
         If l_arr_TipLin(cmb_TipLin.ListIndex + 1).Genera_Codigo <> "" Then
             g_str_Parame = g_str_Parame & CStr(l_arr_TipLin(cmb_TipLin.ListIndex + 1).Genera_Codigo) & ", "
         Else
             g_str_Parame = g_str_Parame & 0 & ", "
         End If
      End If
      
      'Línea de Crédito
      If moddat_g_str_CodPrd <> "008" Then   'And moddat_g_str_CodMod <> "008"
         g_str_Parame = g_str_Parame & CStr(0) & ", "
      ElseIf moddat_g_str_CodPrd = "008" And moddat_g_str_CodMod = "002" Then
         g_str_Parame = g_str_Parame & CStr(0) & ", "
      Else
         If moddat_g_int_FlgGrb_1 = 1 Then
            If CStr(cmb_TipLin.ItemData(cmb_TipLin.ListIndex)) = 1 Then       'NO REVOLVENTE
               g_str_Parame = g_str_Parame & fs_Calcula_LinCre(CStr(moddat_g_int_TipDoc), CStr(moddat_g_str_NumDoc), ipp_ImpCar.Value) & ", "
               
            ElseIf CStr(cmb_TipLin.ItemData(cmb_TipLin.ListIndex)) = 2 Then   'REVOLVENTE
               g_str_Parame = g_str_Parame & fs_Calcula_LinCre(CStr(moddat_g_int_TipDoc), CStr(moddat_g_str_NumDoc), ipp_ImpCar.Value) & ", "
               
            Else
               g_str_Parame = g_str_Parame & CStr(0) & ", "
            End If
         Else
            g_str_Parame = g_str_Parame & CStr(0) & ", "
         End If
      End If
      
      g_str_Parame = g_str_Parame & CStr(moddat_g_int_FlgGrb_1) & ", "
      
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
   
   'Asociar numref renovada y desasociar numref anterior
   If moddat_g_int_FlgGrb_1 = 6 Then
      Call fs_DesAso_NumRef_Renova(CStr(Replace(moddat_g_str_NumFia, "-", "")), l_str_NumRef, moddat_g_int_TipDoc, CStr(moddat_g_str_NumDoc))
   End If
   
   'Generar Asientos automáticos para Cartas Fianza
   If moddat_g_int_FlgGOK = True Then   'And moddat_g_str_CodMod <> "005"
      
      If moddat_g_str_CodPrd <> "008" Then      'CREDITOS INDIRECTOS
      
         If moddat_g_int_TipCli = 1 Then        'MICRO
            r_str_CtaCon = "721201010103"
         ElseIf moddat_g_int_TipCli = 2 Then    'PEQUEÑA
            r_str_CtaCon = "721201010102"
         ElseIf moddat_g_int_TipCli = 3 Then    'MEDIANA
            r_str_CtaCon = "721201010101"
         ElseIf moddat_g_int_TipCli = 4 Then    'GRANDE
            r_str_CtaCon = "721201010101"
         End If
         
         If r_str_CtaCon <> "" Then
            If moddat_g_int_FlgGrb_1 = 1 Then                     'NUEVA CF - AD
               If moddat_g_str_CodMod <> "005" Then
                  Call fs_GeneraAsiento(Trim(txt_NumRef.Text), CStr(moddat_g_str_NumDoc), CStr(moddat_g_str_NomCli), "711201010101", r_str_CtaCon, CDbl(pnl_ValGar.Caption))
                  Call fs_GeneraAsiento(Trim(txt_NumRef.Text), CStr(moddat_g_str_NumDoc), CStr(moddat_g_str_NomCli), "151719010114", "291102070101", CDbl(pnl_ValCom.Caption))  '"151719010104"
               Else
                  If CDbl(pnl_ValCom.Caption) > 0 Then            'COMISION - AD
                     Call fs_GeneraAsiento(Trim(txt_NumRef.Text), CStr(moddat_g_str_NumDoc), CStr(moddat_g_str_NomCli), "151719010114", "291102070101", CDbl(pnl_ValCom.Caption))
                  End If
               End If
            ElseIf moddat_g_int_FlgGrb_1 = 2 Then                 'MODIFICACION DE CF - AD
               If moddat_g_str_CodMod <> "005" Then
                  'AUMENTO DE CF
                  If CDbl(ipp_ImpCar.Value) > r_dbl_ValFia Then
   '                  Call fs_GeneraAsiento(Trim(txt_NumRef.Text), CStr(moddat_g_str_NumDoc), CStr(moddat_g_str_NomCli), "711201010101", r_str_CtaCon, CDbl(pnl_ValGar.Caption))
   '                  Call fs_GeneraAsiento(Trim(txt_NumRef.Text), CStr(moddat_g_str_NumDoc), CStr(moddat_g_str_NomCli), "151719010112", "251419010110", CDbl(ipp_ImpCar.Value))
   '
                  'DISMINUCION DE CF
                  ElseIf CDbl(ipp_ImpCar.Value) < r_dbl_ValFia Then
   '                  Call fs_GeneraAsiento(Trim(txt_NumRef.Text), CStr(moddat_g_str_NumDoc), CStr(moddat_g_str_NomCli), r_str_CtaCon, "711201010101", CDbl(pnl_ValGar.Caption))
   '                  Call fs_GeneraAsiento(Trim(txt_NumRef.Text), CStr(moddat_g_str_NumDoc), CStr(moddat_g_str_NomCli), "251419010110", "151719010112", CDbl(ipp_ImpCar.Value))
                  End If
               End If
            ElseIf moddat_g_int_FlgGrb_1 = 6 Then                 'RENOVACION DE CF - AD
               If moddat_g_str_CodMod <> "005" Then
                  If CDbl(ipp_ImpCar.Value) > r_dbl_ValFia Then   'AUMENTO DE LINEA - CF
                     Call fs_GeneraAsiento(Trim(txt_NumRef.Text), CStr(moddat_g_str_NumDoc), CStr(moddat_g_str_NomCli), "711201010101", r_str_CtaCon, CDbl(pnl_ValGar.Caption))
                  End If
                  If CDbl(pnl_ValCom.Caption) > 0 Then            'COMISION DE RENOVACIÓN - CF
                     Call fs_GeneraAsiento(Trim(txt_NumRef.Text), CStr(moddat_g_str_NumDoc), CStr(moddat_g_str_NomCli), "151719010114", "291102070101", CDbl(pnl_ValCom.Caption))  '"151719010104"
                  End If
               Else
                  If CDbl(pnl_ValCom.Caption) > 0 Then            'COMISION DE RENOVACIÓN - AD
                     Call fs_GeneraAsiento(Trim(txt_NumRef.Text), CStr(moddat_g_str_NumDoc), CStr(moddat_g_str_NomCli), "151719010114", "291102070101", CDbl(pnl_ValCom.Caption))
                  End If
               End If
            End If
         End If
      
      Else  'CREDITOS DIRECTOS
      
      End If
   End If
   
   'Renovación - Ingreso de Saldos
   If moddat_g_int_FlgGrb_1 = 6 Then
      'Call fs_Obtener_SalRen(Trim(txt_NumRef.Text), l_str_NumRef, moddat_g_str_FecSis, CStr(cmb_Moneda.ItemData(cmb_Moneda.ListIndex)), CStr(moddat_g_int_TipDoc), CStr(moddat_g_str_NumDoc))
      If r_dbl_SalFon > 0 Then
         Call fs_Ingresar_MaeRde(1, 16, l_str_NumRef, Format(moddat_g_str_FecSis, "yyyymmdd"), CStr(cmb_Moneda.ItemData(cmb_Moneda.ListIndex)), r_dbl_SalFon, "SALDO DE FONDO RECIBIDO DE CF " & Trim(txt_NumRef.Text), 0, "", "", 1, "", 0, "")
      End If
      If r_dbl_SalDes > 0 Then
         Call fs_Ingresar_MaeRde(2, 17, l_str_NumRef, Format(moddat_g_str_FecSis, "yyyymmdd"), CStr(cmb_Moneda.ItemData(cmb_Moneda.ListIndex)), r_dbl_SalDes, "SALDO A DESEMBOLSAR DE CF " & Trim(txt_NumRef.Text), 0, "", "", 1, "", 0, "")
      End If
      'Verifica si tiene Depósito Garantía Cliente
      Call fs_Verificar_DepGarCli(Trim(txt_NumRef.Text))
   End If
   
   MsgBox "Actualización realizada satisfactoriamente.", vbInformation, modgen_g_con_PltPar
   
   'Actualiza la Carta Fianza en la Grilla
   frm_Ges_TecPro_03.fs_Buscar_Creditos_Indirectos
   frm_Ges_TecPro_03.fs_Buscar_Creditos_Directos
   frm_Ges_TecPro_03.fs_Activa (False)
   frm_Ges_TecPro_03.cmd_Agrega.Enabled = True
   frm_Ges_TecPro_03.cmd_Buscar.Enabled = True
   frm_Ges_TecPro_03.cmd_Limpia.Enabled = True
   frm_Ges_TecPro_01.fs_Buscar
   Call fs_Limpia
   Unload Me
End Sub
Private Function fs_DesAso_NumRef_Renova(ByVal p_NumRef As String, ByVal p_RefRen As String, ByVal p_TipDoc As Integer, ByVal p_NumDoc As String) As Integer
   
   fs_DesAso_NumRef_Renova = False
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
      
   Do While moddat_g_int_FlgGOK = False
      Screen.MousePointer = 11
               
      'Grabando Información en Maestro de Garantías
      g_str_Parame = "USP_TPR_MAEGAR_NUMREF_RENUEVA ("
      g_str_Parame = g_str_Parame & "'" & CStr(p_NumRef) & "', "
      g_str_Parame = g_str_Parame & "'" & CStr(p_RefRen) & "', "
      g_str_Parame = g_str_Parame & CStr(p_TipDoc) & ", "
      g_str_Parame = g_str_Parame & "'" & CStr(p_NumDoc) & "', "

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
            Exit Function

         Else
            moddat_g_int_CntErr = 0
         End If
      End If

      Screen.MousePointer = 0
   Loop
   
   fs_DesAso_NumRef_Renova = True
End Function
Private Function fs_GeneraNumIte() As Integer
   fs_GeneraNumIte = 0
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT NVL(MAX(MAERDE_NUMITE),0) NUMITE FROM TPR_MAERDE WHERE MAERDE_NUMREF =  '" & CStr(Replace(moddat_g_str_DesIte, "-", "")) & "'" 'pnl_NumRef.Caption
    
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
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
Private Function fs_Calcula_LinCre(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String, ByVal p_ImpCar As Double) As Double
   
   fs_Calcula_LinCre = 0
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "    SELECT CASE WHEN (SELECT COUNT(*)"
   g_str_Parame = g_str_Parame & "                        FROM TPR_MAECFI"
   g_str_Parame = g_str_Parame & "                       WHERE MAECFI_TIPDOC = " & p_TipDoc & " "
   g_str_Parame = g_str_Parame & "                         AND MAECFI_NUMDOC = '" & p_NumDoc & "' "
   g_str_Parame = g_str_Parame & "                         AND MAECFI_CODPRD = '008' "
   g_str_Parame = g_str_Parame & "                         AND MAECFI_SITUAC =  1 "
   g_str_Parame = g_str_Parame & "                      ) > 0 THEN "
   g_str_Parame = g_str_Parame & "                     (SELECT MAECFI_LINCRE"
   g_str_Parame = g_str_Parame & "                        FROM TPR_MAECFI"
   g_str_Parame = g_str_Parame & "                       WHERE MAECFI_TIPDOC = " & p_TipDoc & " "
   g_str_Parame = g_str_Parame & "                         AND MAECFI_NUMDOC = '" & p_NumDoc & "' "
   g_str_Parame = g_str_Parame & "                         AND MAECFI_CODPRD =  '008' "
   g_str_Parame = g_str_Parame & "                         AND MAECFI_SITUAC =  1 "
   g_str_Parame = g_str_Parame & "                         AND MAECFI_EMIFIA = (SELECT MAX(MAECFI_EMIFIA)"
   g_str_Parame = g_str_Parame & "                                                FROM TPR_MAECFI"
   g_str_Parame = g_str_Parame & "                                               WHERE MAECFI_TIPDOC = " & p_TipDoc & " "
   g_str_Parame = g_str_Parame & "                                                 AND MAECFI_NUMDOC = '" & p_NumDoc & "' "
   g_str_Parame = g_str_Parame & "                                                 AND MAECFI_CODPRD =  '008' "
   g_str_Parame = g_str_Parame & "                                                 AND MAECFI_SITUAC =  1 )"
   g_str_Parame = g_str_Parame & "                      )"
   g_str_Parame = g_str_Parame & "           ELSE "
   g_str_Parame = g_str_Parame & "                  (SELECT MAEETE_LINASI_DIR "
   g_str_Parame = g_str_Parame & "                     FROM TPR_MAEETE "
   g_str_Parame = g_str_Parame & "                    WHERE MAEETE_TIPDOC = " & p_TipDoc & " "
   g_str_Parame = g_str_Parame & "                      AND MAEETE_NUMDOC = '" & p_NumDoc & "' "
   g_str_Parame = g_str_Parame & "                      AND MAEETE_SITUAC =  1 ) "
   g_str_Parame = g_str_Parame & "           END AS MONTO "
   g_str_Parame = g_str_Parame & "      FROM DUAL "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
       Exit Function
   End If

   If g_rst_GenAux.BOF And g_rst_GenAux.EOF Then
      g_rst_GenAux.Close
      Set g_rst_GenAux = Nothing
      Exit Function
   End If
   
   If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
      fs_Calcula_LinCre = CDbl(g_rst_GenAux!MONTO) - CDbl(p_ImpCar)
   End If

End Function

Private Sub fs_Obtener_SalRen(ByVal p_Numref_Ori As String, ByVal p_TipDoc As Integer, ByVal p_NumDoc As String, ByRef p_SalFon As Double, ByRef p_SalDes As Double)
'ByVal p_Numref_Ren As String, ByVal p_FecOpe As String, ByVal p_TipMon As Integer,

   p_SalFon = 0
   p_SalDes = 0
     
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "   SELECT REFERENCIA           , IMPORTE_FONDOS       , RECIBIDO_FONDOS   , (IMPORTE_FONDOS - RECIBIDO_FONDOS) SALDO_FONDOS             , "
   g_str_Parame = g_str_Parame & "          IMPORTE_DESEMBOLSADO , PAGADO_DESEMBOLSO , (IMPORTE_DESEMBOLSADO - PAGADO_DESEMBOLSO) SALDO_DESEMBOLSO "
   g_str_Parame = g_str_Parame & "    FROM( "
   g_str_Parame = g_str_Parame & "          SELECT A.MAECFI_NUMREF                                                      REFERENCIA, "
   
   g_str_Parame = g_str_Parame & "                 CASE WHEN A.MAECFI_NUMREN > 0 THEN"
   g_str_Parame = g_str_Parame & "                              NVL(( SELECT NVL(SUM(MAERDE_IMPORT),0) "
   g_str_Parame = g_str_Parame & "                                      FROM TPR_MAERDE B "
   g_str_Parame = g_str_Parame & "                                     WHERE B.MAERDE_NUMREF = A.MAECFI_NUMREF AND MAERDE_CODIGO = 16 "
   g_str_Parame = g_str_Parame & "                                       AND B.MAERDE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                                     GROUP BY MAERDE_NUMREF),0)                                      "
   g_str_Parame = g_str_Parame & "                 ELSE"
   g_str_Parame = g_str_Parame & "                              A.MAECFI_IMPFIA                                                      "
   g_str_Parame = g_str_Parame & "                 END                                                                  IMPORTE_FONDOS, "
   
   g_str_Parame = g_str_Parame & "                 NVL((SELECT NVL(SUM(MAERDE_IMPORT),0) "
   g_str_Parame = g_str_Parame & "                        FROM TPR_MAERDE B "
   g_str_Parame = g_str_Parame & "                       WHERE B.MAERDE_NUMREF = A.MAECFI_NUMREF AND (MAERDE_CODIGO = 2 OR MAERDE_CODIGO = 10 OR MAERDE_CODIGO = 3 OR MAERDE_CODIGO = 11 OR MAERDE_CODIGO = 12 OR MAERDE_CODIGO = 14 OR MAERDE_CODIGO = 19)  "
   g_str_Parame = g_str_Parame & "                         AND B.MAERDE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                       GROUP BY MAERDE_NUMREF),0)                                     RECIBIDO_FONDOS,"
   
   g_str_Parame = g_str_Parame & "                 CASE WHEN A.MAECFI_NUMREN > 0 THEN"
   g_str_Parame = g_str_Parame & "                              NVL((SELECT NVL(SUM(MAERDE_IMPORT),0) "
   g_str_Parame = g_str_Parame & "                                     FROM TPR_MAERDE B "
   g_str_Parame = g_str_Parame & "                                    WHERE B.MAERDE_NUMREF = A.MAECFI_NUMREF AND (MAERDE_CODIGO = 2 OR MAERDE_CODIGO = 10 OR MAERDE_CODIGO = 3 OR MAERDE_CODIGO = 11 OR MAERDE_CODIGO = 12 OR MAERDE_CODIGO = 14 OR MAERDE_CODIGO = 17 OR MAERDE_CODIGO = 19)  "
   g_str_Parame = g_str_Parame & "                                      AND B.MAERDE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                                    GROUP BY MAERDE_NUMREF),0) "
   g_str_Parame = g_str_Parame & "                 ELSE                   "
   g_str_Parame = g_str_Parame & "                              NVL((SELECT NVL(SUM(MAERDE_IMPORT),0) "
   g_str_Parame = g_str_Parame & "                                     FROM TPR_MAERDE B "
   g_str_Parame = g_str_Parame & "                                    WHERE B.MAERDE_NUMREF = A.MAECFI_NUMREF AND (MAERDE_CODIGO = 2 OR MAERDE_CODIGO = 10 OR MAERDE_CODIGO = 3 OR MAERDE_CODIGO = 11 OR MAERDE_CODIGO = 12 OR MAERDE_CODIGO = 14 OR MAERDE_CODIGO = 7 OR MAERDE_CODIGO = 19)  "
   g_str_Parame = g_str_Parame & "                                      AND B.MAERDE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                                    GROUP BY MAERDE_NUMREF),0)                                     "
   g_str_Parame = g_str_Parame & "                 END                                                                  IMPORTE_DESEMBOLSADO,"
   
   g_str_Parame = g_str_Parame & "                 NVL((SELECT NVL(SUM(MAERDE_IMPORT),0)"
   g_str_Parame = g_str_Parame & "                        FROM TPR_MAERDE B"
   g_str_Parame = g_str_Parame & "                       WHERE B.MAERDE_NUMREF = A.MAECFI_NUMREF AND (MAERDE_CODIGO = 1 Or MAERDE_CODIGO = 2 Or MAERDE_CODIGO = 4 Or MAERDE_CODIGO = 5 OR MAERDE_CODIGO = 6 OR MAERDE_CODIGO = 18)" 'OR MAERDE_CODIGO = 10
   g_str_Parame = g_str_Parame & "                         AND B.MAERDE_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                       GROUP BY MAERDE_NUMREF),0)                                     PAGADO_DESEMBOLSO"
   
   g_str_Parame = g_str_Parame & "            FROM TPR_MAECFI A "
   g_str_Parame = g_str_Parame & "                 INNER JOIN MNT_PARDES   ON PARDES_CODGRP = '529' AND PARDES_CODITE = MAECFI_SITUAC "
   g_str_Parame = g_str_Parame & "           WHERE A.MAECFI_NUMREF = '" & CStr(p_Numref_Ori) & "' "
   g_str_Parame = g_str_Parame & "             AND A.MAECFI_TIPDOC = " & p_TipDoc & ""
   g_str_Parame = g_str_Parame & "             AND A.MAECFI_NUMDOC = '" & CStr(p_NumDoc) & "' "
   g_str_Parame = g_str_Parame & "        ) "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If

   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      If moddat_g_int_TipRec = 2 Then
         p_SalFon = 0
      Else
         p_SalFon = CDbl(g_rst_Princi!SALDO_FONDOS)
      End If
      p_SalDes = CDbl(g_rst_Princi!SALDO_DESEMBOLSO)
   End If
   
'  Call fs_Ingresar_MaeRde(1, 16, p_Numref_Ren, Format(p_FecOpe, "yyyymmdd"), p_TipMon, r_dbl_ImpOpe, "SALDO DE FONDO RECIBIDO DE CF " & p_Numref_Ori, 0, "", "", 1, "", 0, "")
'  Call fs_Ingresar_MaeRde(2, 17, p_Numref_Ren, Format(p_FecOpe, "yyyymmdd"), p_TipMon, r_dbl_ImpSal, "SALDO A DESEMBOLSAR DE CF " & p_Numref_Ori, 0, "", "", 1, "", 0, "")
End Sub

Private Sub fs_Ingresar_Movimientos(ByVal p_NumRef As String)
Dim r_int_NumIte  As Integer
Dim r_str_OpeRef  As String
Dim r_str_CtaCte  As String
Dim r_str_NumCCI  As String
Dim r_str_TdoPrv  As String
Dim r_str_NdoPrv  As String
Dim r_str_NumMov  As String

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT MAERDE_NUMITE, MAERDE_CODIGO, MAERDE_NUMREF, MAERDE_FECASG, MAERDE_TIPMON"
   g_str_Parame = g_str_Parame & "         MAERDE_IMPORT, MAERDE_OPEREF, MAERDE_CODBAN, MAERDE_CTACTE, MAERDE_NUMCCI,"
   g_str_Parame = g_str_Parame & "         MAERDE_SITUAC, MAERDE_TDOPRV, MAERDE_NDOPRV, MAERDE_NUMMOV "
   g_str_Parame = g_str_Parame & "    FROM TPR_MAERDE "
   g_str_Parame = g_str_Parame & "   WHERE MAERDE_NUMREF = '" & CStr(p_NumRef) & "' "
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If

   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      r_int_NumIte = fs_GeneraNumIte
      If IsNull(g_rst_Princi!MAERDE_OPEREF) Then
         r_str_OpeRef = ""
      End If
      If IsNull(g_rst_Princi!MAERDE_CTACTE) Then
         r_str_CtaCte = ""
      End If
      If IsNull(g_rst_Princi!MAERDE_NUMCCI) Then
         r_str_NumCCI = ""
      End If
      If IsNull(g_rst_Princi!MAERDE_TDOPRV) Then
         r_str_TdoPrv = ""
      End If
      If IsNull(g_rst_Princi!MAERDE_NDOPRV) Then
         r_str_NdoPrv = ""
      End If
      If IsNull(g_rst_Princi!MAERDE_NUMMOV) Then
         r_str_NumMov = ""
      End If
      If fs_Ingresar_MaeRde(r_int_NumIte, g_rst_Princi!MAERDE_CODIGO, l_str_NumRef, g_rst_Princi!MAERDE_FECASG, g_rst_Princi!MAERDE_TIPMON, _
                            g_rst_Princi!MAERDE_IMPORT, r_str_OpeRef, g_rst_Princi!MAERDE_CODBAN, r_str_CtaCte, r_str_NumCCI, 1, r_str_NumMov, r_str_TdoPrv, r_str_NdoPrv) Then
         g_rst_Princi.MoveNext
      End If
   Loop
End Sub

Private Sub fs_Ingresar_Garantia(ByVal p_NumRef As String)
Dim r_int_NumIte  As Integer

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT MAEGAR_NUMREF, MAEGAR_NUMOPE, MAEGAR_TIPGAR, MAEGAR_FECINS, MAEGAR_TIPMON, MAEGAR_SITUAC, MAEGAR_NROCNT, "
   g_str_Parame = g_str_Parame & "         (NVL(MAEGAR_MTOGAR_INM,0) + NVL(MAEGAR_MTOGAR_ES1,0) + NVL(MAEGAR_MTOGAR_ES2,0) + NVL(MAEGAR_MTOGAR_DE1,0) + NVL(MAEGAR_MTOGAR_DE2,0)) AS MAEGAR_MTOGAR "
   g_str_Parame = g_str_Parame & "    FROM TPR_MAEGAR "
   g_str_Parame = g_str_Parame & "   WHERE MAEGAR_NUMREF = '" & CStr(p_NumRef) & "' "
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If

   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      r_int_NumIte = fs_GeneraNumIte
      
      If fs_Ingresar_MaeGar(r_int_NumIte, l_str_NumRef, g_rst_Princi!MAEGAR_TIPGAR, g_rst_Princi!MAEGAR_FECINS, g_rst_Princi!MAEGAR_TIPMON, g_rst_Princi!MAEGAR_MTOGAR, 1) Then
         g_rst_Princi.MoveNext
      End If
   Loop
End Sub

Function fs_Buscar_Saldo_Comision(ByRef p_ImpTot As Double, ByRef p_ImpSal As Double, ByRef p_ImpPag As Double)  'As Double
   p_ImpTot = 0
   p_ImpSal = 0
 
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT A.MAECFI_COMFIA IMPORTE_COMISION, NVL(SUM(NVL(B.MAERDE_IMPORT,0)),0) IMPORTE_PAGADO "
   g_str_Parame = g_str_Parame & "    FROM TPR_MAECFI A "
   g_str_Parame = g_str_Parame & "         LEFT JOIN (SELECT B.MAERDE_NUMREF , B.MAERDE_NUMITE, B.MAERDE_IMPORT "
   g_str_Parame = g_str_Parame & "                      FROM TPR_MAERDE B "
   g_str_Parame = g_str_Parame & "                     WHERE (MAERDE_CODIGO = 1 OR MAERDE_CODIGO = 2))B "
   g_str_Parame = g_str_Parame & "                        ON B.MAERDE_NUMREF = A.MAECFI_NUMREF "
   g_str_Parame = g_str_Parame & "   WHERE A.MAECFI_NUMREF = '" & CStr(l_str_NumRef) & "' "
   g_str_Parame = g_str_Parame & "   GROUP BY MAECFI_COMFIA "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Function
   End If

   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Function
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      p_ImpTot = Format(CDbl(g_rst_Princi!IMPORTE_COMISION), "###,###,###,##0.00")
      p_ImpPag = Format(CDbl(g_rst_Princi!IMPORTE_PAGADO), "###,###,###,##0.00")
      p_ImpSal = Format(CDbl(g_rst_Princi!IMPORTE_COMISION) - CDbl(g_rst_Princi!IMPORTE_PAGADO), "###,###,###,##0.00")
   End If
End Function

Private Sub fs_Buscar_Saldo_Desembolso(ByRef p_ImpTot As Double, ByRef p_ImpSal As Double)
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT NVL(SUM(A.MAERDE_IMPORT), 0) IMPORTE_RECIBIDO, NVL((IMPORTE_PAGADO),0) IMPORTE_PAGADO " 'SUM
   g_str_Parame = g_str_Parame & "    FROM TPR_MAERDE A "
   g_str_Parame = g_str_Parame & "         LEFT JOIN (SELECT B.MAERDE_NUMREF, NVL(SUM(NVL(B.MAERDE_IMPORT,0)),0) IMPORTE_PAGADO "
   g_str_Parame = g_str_Parame & "                      FROM TPR_MAERDE B "
   g_str_Parame = g_str_Parame & "                     WHERE (MAERDE_CODIGO = 4 OR MAERDE_CODIGO = 5 ) "
   g_str_Parame = g_str_Parame & "                     GROUP BY B.MAERDE_NUMREF ) B "
   g_str_Parame = g_str_Parame & "                        ON B.MAERDE_NUMREF = A.MAERDE_NUMREF "
   g_str_Parame = g_str_Parame & "   WHERE A.MAERDE_NUMREF = '" & CStr(txt_NumRef.Text) & "' "
   'g_str_Parame = g_str_Parame & "     AND A.MAERDE_CODIGO = 3 "
   g_str_Parame = g_str_Parame & "     AND A.MAERDE_CODIGO IN (3,19) "
   g_str_Parame = g_str_Parame & "   GROUP BY IMPORTE_PAGADO"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If

   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      p_ImpTot = Format(CDbl(g_rst_Princi!IMPORTE_RECIBIDO), "###,###,###,##0.00")
      p_ImpSal = CDbl(g_rst_Princi!IMPORTE_RECIBIDO) - CDbl(g_rst_Princi!IMPORTE_PAGADO)
   End If
End Sub

Private Function fs_GeneraNumRef() As String
Dim r_int_NumRef  As Integer
Dim r_int_TipMod  As Integer

   fs_GeneraNumRef = ""
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "      SELECT MAX(SUBSTR(MAECFI_NUMREF,6,5)) NUMREF "
   g_str_Parame = g_str_Parame & "        FROM TPR_MAECFI "
   
   If moddat_g_str_CodPrd = "026" Or moddat_g_str_CodPrd = "027" Then
      g_str_Parame = g_str_Parame & "       WHERE MAECFI_CODPRD IN ('026','027')"
   Else
      g_str_Parame = g_str_Parame & "       WHERE MAECFI_CODPRD = '" & moddat_g_str_CodPrd & "'"
   End If
   
   If moddat_g_str_CodPrd = "026" Or moddat_g_str_CodPrd = "027" Then
      If moddat_g_str_CodMod = "005" Then
         g_str_Parame = g_str_Parame & "   AND SUBSTR(MAECFI_NUMREF,1,1) = 2 "
      ElseIf moddat_g_str_CodMod = "008" Then
         g_str_Parame = g_str_Parame & "   AND SUBSTR(MAECFI_NUMREF,1,1) = 3 "
      Else
         g_str_Parame = g_str_Parame & "   AND SUBSTR(MAECFI_NUMREF,1,1) = 1 "
      End If
   End If
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
       Exit Function
   End If
   
   If g_rst_GenAux.BOF And g_rst_GenAux.EOF Then
      g_rst_GenAux.Close
      Set g_rst_GenAux = Nothing
      Exit Function
   End If
     
   If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
      r_int_NumRef = IIf(IsNull(g_rst_GenAux!NUMREF), 0, g_rst_GenAux!NUMREF) + 1
   End If
   If moddat_g_str_CodPrd = "026" Or moddat_g_str_CodPrd = "027" Then
      If moddat_g_str_CodMod = "005" Then
         r_int_TipMod = 2
      ElseIf moddat_g_str_CodMod = "008" Then
         r_int_TipMod = 3
      Else
        r_int_TipMod = 1
      End If
      fs_GeneraNumRef = r_int_TipMod & Right("00" & Year(Now), 2) & Right("00" & Month(Now), 2) & Right("00000" & r_int_NumRef, 5)
   Else
      fs_GeneraNumRef = moddat_g_str_CodPrd & Right("00" & Year(Now), 2) & Right("00000" & r_int_NumRef, 5)
   End If
End Function

Private Function fs_GeneraNumRef_OLD() As String
Dim r_int_NumRef  As Integer
Dim r_int_TipMod  As Integer

   fs_GeneraNumRef_OLD = ""
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT MAX(SUBSTR(MAECFI_NUMREF,6,5)) NUMREF "
   g_str_Parame = g_str_Parame & "   FROM TPR_MAECFI "
   g_str_Parame = g_str_Parame & "  WHERE MAECFI_CODPRD = '" & IIf(moddat_g_str_CodPrd = "027", "026", moddat_g_str_CodPrd) & "' "
   If moddat_g_str_CodPrd = "026" Or moddat_g_str_CodPrd = "027" Then
      If moddat_g_str_CodMod = "005" Then
         g_str_Parame = g_str_Parame & "  AND MAECFI_CODMOD = '" & moddat_g_str_CodMod & "' "
      ElseIf moddat_g_str_CodMod = "008" Then
      
      
      End If
   End If
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
       Exit Function
   End If
   
   If g_rst_GenAux.BOF And g_rst_GenAux.EOF Then
      g_rst_GenAux.Close
      Set g_rst_GenAux = Nothing
      Exit Function
   End If
     
   If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
      r_int_NumRef = IIf(IsNull(g_rst_GenAux!NUMREF), 0, g_rst_GenAux!NUMREF) + 1
   End If
   If moddat_g_str_CodPrd = "026" Or moddat_g_str_CodPrd = "027" Then
      If moddat_g_str_CodMod = "005" Then
         r_int_TipMod = 2
      ElseIf moddat_g_str_CodMod = "008" Then
         r_int_TipMod = 3
      Else
         r_int_TipMod = 1
      End If
      fs_GeneraNumRef_OLD = r_int_TipMod & Right("00" & Year(Now), 2) & Right("00" & Month(Now), 2) & Right("00000" & r_int_NumRef, 5)
   Else
      fs_GeneraNumRef_OLD = moddat_g_str_CodPrd & Right("00" & Year(Now), 2) & Right("00000" & r_int_NumRef, 5)
   End If
End Function

Private Sub fs_Verificar_DepGarCli(ByVal p_NumRef As String)

Dim r_int_NumIte  As Integer

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT ("
   g_str_Parame = g_str_Parame & "         SELECT COUNT(MAERDE_CODIGO)  "
   g_str_Parame = g_str_Parame & "           FROM TPR_MAERDE  "
   g_str_Parame = g_str_Parame & "          WHERE MAERDE_NUMREF = '" & p_NumRef & "' "
   g_str_Parame = g_str_Parame & "            AND MAERDE_CODIGO = 20) AS CANT_20, "
   g_str_Parame = g_str_Parame & "        (SELECT COUNT(MAERDE_CODIGO)  "
   g_str_Parame = g_str_Parame & "           FROM TPR_MAERDE  "
   g_str_Parame = g_str_Parame & "          WHERE MAERDE_NUMREF = '" & p_NumRef & "' "
   g_str_Parame = g_str_Parame & "            AND MAERDE_CODIGO = 21) AS CANT_21 "
   g_str_Parame = g_str_Parame & "   FROM DUAL "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
       Exit Sub
   End If
   
   If g_rst_GenAux.BOF And g_rst_GenAux.EOF Then
      g_rst_GenAux.Close
      Set g_rst_GenAux = Nothing
      Exit Sub
   End If
     
   If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
      If g_rst_GenAux!CANT_20 <> g_rst_GenAux!CANT_21 Then
      
         If CInt(g_rst_GenAux!CANT_20) - CInt(g_rst_GenAux!CANT_21) = 1 Then
            g_str_Parame = ""
            g_str_Parame = g_str_Parame & "   SELECT MAERDE_CODIGO, MAERDE_FECASG, MAERDE_TIPMON, MAERDE_IMPORT, MAERDE_OPEREF, MAERDE_NUMMOV  "
            g_str_Parame = g_str_Parame & "     FROM TPR_MAERDE  "
            g_str_Parame = g_str_Parame & "    WHERE MAERDE_NUMREF = '" & p_NumRef & "' "
            g_str_Parame = g_str_Parame & "      AND MAERDE_CODIGO = 20 "
            g_str_Parame = g_str_Parame & "      AND MAERDE_NUMITE = ( SELECT MAX(MAERDE_NUMITE) FROM TPR_MAERDE WHERE MAERDE_NUMREF = " & p_NumRef & " AND MAERDE_CODIGO = 20 ) "
            
            If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
                Exit Sub
            End If
            
            If g_rst_Listas.BOF And g_rst_Listas.EOF Then
               g_rst_Listas.Close
               Set g_rst_Listas = Nothing
               Exit Sub
            End If
            
            r_int_NumIte = fs_GeneraNumIte

            Call fs_Ingresar_MaeRde(r_int_NumIte, g_rst_Listas!MAERDE_CODIGO, l_str_NumRef, g_rst_Listas!MAERDE_FECASG, CStr(g_rst_Listas!MAERDE_TIPMON), g_rst_Listas!MAERDE_IMPORT, Mid(Trim("DEPÓSITO GARANTÍA CLIENTE DE " & Trim(txt_NumRef.Text) & " - " & Trim(g_rst_Listas!MAERDE_OPEREF)), 1, 100), 0, "", "", 1, g_rst_Listas!MAERDE_NUMMOV, 0, "")
            
            g_rst_Listas.Close
            Set g_rst_Listas = Nothing
         End If
      End If
      g_rst_GenAux.Close
      Set g_rst_GenAux = Nothing
   End If
End Sub

Private Function fs_Ingresar_MaeRde(ByVal p_NumIte As Integer, ByVal p_TipOper As Integer, ByVal p_NumRef As String, _
                               ByVal p_FecOpe As String, ByVal p_TipMon As Integer, ByVal p_ImpDes As Double, _
                               ByVal p_Refer As String, ByVal p_CodBan As Integer, ByVal p_CtaBan As String, _
                               ByVal p_CCIBan As String, ByVal p_Flag As Integer, ByVal p_NumMov As String, _
                               Optional ByVal p_TipDoc As String, Optional ByVal p_NumDoc As String) As Integer

   fs_Ingresar_MaeRde = False
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
      
   Do While moddat_g_int_FlgGOK = False
      Screen.MousePointer = 11
'      Call moddat_gs_FecSis
      
      g_str_Parame = "USP_TPR_MAERDE ("
      g_str_Parame = g_str_Parame & CStr(p_NumIte) & ", "
      g_str_Parame = g_str_Parame & CStr(p_TipOper) & ", "
      g_str_Parame = g_str_Parame & "'" & CStr(Replace(p_NumRef, "-", "")) & "', "
      g_str_Parame = g_str_Parame & "'" & p_FecOpe & "', "
      g_str_Parame = g_str_Parame & CStr(p_TipMon) & ", "
      g_str_Parame = g_str_Parame & CDbl(p_ImpDes) & ", "
      g_str_Parame = g_str_Parame & "'" & Trim(p_Refer) & "', "
      
      If p_TipDoc = "" Then
         g_str_Parame = g_str_Parame & "null, "
      Else
         g_str_Parame = g_str_Parame & CStr(p_TipDoc) & ", "
      End If
      If p_NumDoc = Empty Then
         g_str_Parame = g_str_Parame & "null, "
      Else
         g_str_Parame = g_str_Parame & CStr(p_NumDoc) & ", "
      End If
      
      g_str_Parame = g_str_Parame & CStr(p_CodBan) & ", "
      If InStr(p_CtaBan, "-") > 0 Then
         g_str_Parame = g_str_Parame & "'" & Trim(Mid(p_CtaBan, 1, InStr(p_CtaBan, "-") - 1)) & "', "
      Else
         g_str_Parame = g_str_Parame & "'" & Trim(p_CtaBan) & "', "
      End If
      g_str_Parame = g_str_Parame & "'" & Trim(p_CCIBan) & "', "
      g_str_Parame = g_str_Parame & "'', "
      
      g_str_Parame = g_str_Parame & CStr(p_Flag) & ", "
      g_str_Parame = g_str_Parame & "'" & CStr(p_NumMov) & "', "
      
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
            Exit Function
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
      
      Screen.MousePointer = 0
   Loop
   
   fs_Ingresar_MaeRde = True
End Function

Private Function fs_Ingresar_MaeGar(ByVal p_NumIte As Integer, ByVal p_NumRef As String, ByVal p_TipGar As Integer, ByVal p_FecEmi As String, ByVal p_TipMon As Integer, ByVal p_ImpGar As Double, ByVal p_Flag As Integer) As Integer
   fs_Ingresar_MaeGar = False
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
      
   Do While moddat_g_int_FlgGOK = False
      Screen.MousePointer = 11
      
      g_str_Parame = "USP_TPR_MAEGAR ("
      g_str_Parame = g_str_Parame & CStr(p_NumIte) & ", "
      g_str_Parame = g_str_Parame & "'" & CStr(p_NumRef) & "', "
      g_str_Parame = g_str_Parame & CStr(p_TipGar) & ", "
      g_str_Parame = g_str_Parame & "'" & p_FecEmi & "', "
      g_str_Parame = g_str_Parame & CStr(p_TipMon) & ", "
      g_str_Parame = g_str_Parame & CStr(p_ImpGar) & ", "
      g_str_Parame = g_str_Parame & CStr(p_Flag) & ", "
      
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
            Exit Function
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
      
      Screen.MousePointer = 0
   Loop
   
   fs_Ingresar_MaeGar = True
End Function

Private Function fs_Validar_Mto_CarFia(ByVal p_Valor As Double, ByVal p_Tipo As Integer) As Boolean
   fs_Validar_Mto_CarFia = False
        
   g_str_Parame = ""
   
   If p_Tipo = 0 Then
      g_str_Parame = g_str_Parame & "   SELECT (NVL(MAEETE_LINASI_IND,0) + NVL(MAEETE_LINASI_DIR,0)) AS LINASI, "
      g_str_Parame = g_str_Parame & "           NVL(SUM(CASE WHEN MAECFI_CODPRD <> '008' THEN MAECFI_GARFIA "
      g_str_Parame = g_str_Parame & "                   ELSE CASE WHEN MAECFI_CODMOD <> '002' THEN MAECFI_IMPFIA ELSE 0 END "
      g_str_Parame = g_str_Parame & "                    END),0) AS CARTA_FIANZA"
      g_str_Parame = g_str_Parame & "     FROM TPR_MAEETE"
      g_str_Parame = g_str_Parame & "          LEFT JOIN TPR_MAECFI ON MAEETE_TIPDOC = MAECFI_TIPDOC AND MAEETE_NUMDOC = MAECFI_NUMDOC AND MAECFI_SITUAC = 1 "
      g_str_Parame = g_str_Parame & "    WHERE MAEETE_TIPDOC = " & CStr(moddat_g_int_TipDoc) & " "
      g_str_Parame = g_str_Parame & "      AND MAEETE_NUMDOC = '" & CStr(moddat_g_str_NumDoc) & "'"
      g_str_Parame = g_str_Parame & "    GROUP BY MAECFI_TIPDOC, MAECFI_NUMDOC, MAEETE_LINASI_IND, MAEETE_LINASI_DIR "
   
   ElseIf p_Tipo = 1 Then
      g_str_Parame = g_str_Parame & "   SELECT (NVL(MAEETE_LINASI_IND,0)) AS LINASI, "
      g_str_Parame = g_str_Parame & "           NVL(SUM(CASE WHEN MAECFI_CODPRD <> '008' THEN MAECFI_GARFIA "
      g_str_Parame = g_str_Parame & "                   ELSE 0 "
      g_str_Parame = g_str_Parame & "                    END),0) AS CARTA_FIANZA"
      g_str_Parame = g_str_Parame & "     FROM TPR_MAEETE"
      g_str_Parame = g_str_Parame & "          LEFT JOIN TPR_MAECFI ON MAEETE_TIPDOC = MAECFI_TIPDOC AND MAEETE_NUMDOC = MAECFI_NUMDOC AND MAECFI_SITUAC = 1 "
      g_str_Parame = g_str_Parame & "    WHERE MAEETE_TIPDOC = " & CStr(moddat_g_int_TipDoc) & " "
      g_str_Parame = g_str_Parame & "      AND MAEETE_NUMDOC = '" & CStr(moddat_g_str_NumDoc) & "'"
      g_str_Parame = g_str_Parame & "    GROUP BY MAECFI_TIPDOC, MAECFI_NUMDOC, MAEETE_LINASI_IND"
   
   ElseIf p_Tipo = 2 Then
      g_str_Parame = g_str_Parame & "   SELECT (NVL(MAEETE_LINASI_DIR,0)) AS LINASI, "
      g_str_Parame = g_str_Parame & "           NVL(SUM(CASE WHEN MAECFI_CODPRD = '008' AND MAECFI_CODMOD <> '002' THEN MAECFI_IMPFIA "
      g_str_Parame = g_str_Parame & "                   ELSE 0 "
      g_str_Parame = g_str_Parame & "                    END),0) AS CARTA_FIANZA"
      g_str_Parame = g_str_Parame & "     FROM TPR_MAEETE"
      g_str_Parame = g_str_Parame & "          LEFT JOIN TPR_MAECFI ON MAEETE_TIPDOC = MAECFI_TIPDOC AND MAEETE_NUMDOC = MAECFI_NUMDOC AND MAECFI_SITUAC = 1 "
      g_str_Parame = g_str_Parame & "    WHERE MAEETE_TIPDOC = " & CStr(moddat_g_int_TipDoc) & " "
      g_str_Parame = g_str_Parame & "      AND MAEETE_NUMDOC = '" & CStr(moddat_g_str_NumDoc) & "'"
      g_str_Parame = g_str_Parame & "    GROUP BY MAECFI_TIPDOC, MAECFI_NUMDOC, MAEETE_LINASI_DIR "
   End If
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
      Exit Function
   End If
   
   If g_rst_GenAux.BOF And g_rst_GenAux.EOF Then
      g_rst_GenAux.Close
      Set g_rst_GenAux = Nothing
      Exit Function
   End If
     
   If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
      If moddat_g_int_FlgGrb_1 = 1 Then
         If CDbl(g_rst_GenAux!CARTA_FIANZA) + CDbl(Trim(p_Valor)) <= CDbl(g_rst_GenAux!LINASI) Then
            fs_Validar_Mto_CarFia = True
         End If
      Else
         If CDbl(g_rst_GenAux!CARTA_FIANZA) = CDbl(Trim(p_Valor)) Then
            If CDbl(g_rst_GenAux!CARTA_FIANZA) <= CDbl(g_rst_GenAux!LINASI) Then
               fs_Validar_Mto_CarFia = True
            End If
         Else
            If CDbl(g_rst_GenAux!CARTA_FIANZA) - l_dbl_ValGar + CDbl(Trim(p_Valor)) <= CDbl(g_rst_GenAux!LINASI) Then
               fs_Validar_Mto_CarFia = True
            Else
               fs_Validar_Mto_CarFia = False
            End If
         End If
      End If
   End If
End Function
Private Function fs_Validar_CreDir_TipLin(ByVal p_Valor As Double, ByVal p_TipLin As Integer) As Boolean
   fs_Validar_CreDir_TipLin = False
        
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "   SELECT " & IIf(p_TipLin = 1, "NVL(MAEETE_LINNRE_DIR, 0)", "NVL(MAEETE_LINREV_DIR, 0)") & " AS LINASI, "
   g_str_Parame = g_str_Parame & "           NVL(SUM(CASE WHEN MAECFI_CODPRD <> '008' THEN MAECFI_GARFIA "
   g_str_Parame = g_str_Parame & "                   ELSE CASE WHEN MAECFI_CODMOD <> '002' THEN MAECFI_IMPFIA ELSE 0 END "
   g_str_Parame = g_str_Parame & "                    END),0) AS CARTA_FIANZA"
   g_str_Parame = g_str_Parame & "     FROM TPR_MAEETE"
   g_str_Parame = g_str_Parame & "          LEFT JOIN TPR_MAECFI ON MAEETE_TIPDOC = MAECFI_TIPDOC AND MAEETE_NUMDOC = MAECFI_NUMDOC  "
   g_str_Parame = g_str_Parame & "                AND MAECFI_SITUAC = 1 AND MAECFI_TIPLIN = " & p_TipLin & " "
   g_str_Parame = g_str_Parame & "    WHERE MAEETE_TIPDOC = " & CStr(moddat_g_int_TipDoc) & " "
   g_str_Parame = g_str_Parame & "      AND MAEETE_NUMDOC = '" & CStr(moddat_g_str_NumDoc) & "'"
   g_str_Parame = g_str_Parame & "    GROUP BY MAECFI_TIPDOC, MAECFI_NUMDOC, MAEETE_LINREV_DIR, MAEETE_LINNRE_DIR "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
      Exit Function
   End If
   
   If g_rst_GenAux.BOF And g_rst_GenAux.EOF Then
      g_rst_GenAux.Close
      Set g_rst_GenAux = Nothing
      Exit Function
   End If
     
   If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
      If moddat_g_int_FlgGrb_1 = 1 Then
         If CDbl(g_rst_GenAux!CARTA_FIANZA) + CDbl(Trim(p_Valor)) <= CDbl(g_rst_GenAux!LINASI) Then
            fs_Validar_CreDir_TipLin = True
         End If
      Else
         If CDbl(g_rst_GenAux!CARTA_FIANZA) = CDbl(Trim(p_Valor)) Then
            If CDbl(g_rst_GenAux!CARTA_FIANZA) <= CDbl(g_rst_GenAux!LINASI) Then
               fs_Validar_CreDir_TipLin = True
            End If
         Else
            If CDbl(g_rst_GenAux!CARTA_FIANZA) - l_dbl_ValGar + CDbl(Trim(p_Valor)) <= CDbl(g_rst_GenAux!LINASI) Then
               fs_Validar_CreDir_TipLin = True
            Else
               fs_Validar_CreDir_TipLin = False
            End If
         End If
      End If
   End If
End Function

Private Function fs_Validar_Retencion_Garantia() As Boolean
   
   fs_Validar_Retencion_Garantia = False
        
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "   SELECT NVL(SUM(A.MAERDE_IMPORT),0) RETENCION , NVL(B.DEVOLUCION,0) AS DEVOLUCION  "
   g_str_Parame = g_str_Parame & "     FROM TPR_MAERDE A "
   g_str_Parame = g_str_Parame & "          LEFT JOIN (SELECT B.MAERDE_NUMREF, NVL(SUM(B.MAERDE_IMPORT),0) AS DEVOLUCION "
   g_str_Parame = g_str_Parame & "                       FROM TPR_MAERDE B "
   g_str_Parame = g_str_Parame & "                      WHERE B.MAERDE_CODIGO = 7 "
   g_str_Parame = g_str_Parame & "                      GROUP BY B.MAERDE_NUMREF) B ON "
   g_str_Parame = g_str_Parame & "                      A.MAERDE_NUMREF = B.MAERDE_NUMREF "
   g_str_Parame = g_str_Parame & "    WHERE A.MAERDE_NUMREF = " & CStr(moddat_g_str_DesIte) & " "
   g_str_Parame = g_str_Parame & "      AND A.MAERDE_CODIGO = 6 "
   g_str_Parame = g_str_Parame & "    GROUP BY B.DEVOLUCION "
          
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
     Exit Function
   End If
   
   If g_rst_GenAux.BOF And g_rst_GenAux.EOF Then
      g_rst_GenAux.Close
      Set g_rst_GenAux = Nothing
      fs_Validar_Retencion_Garantia = True
      Exit Function
   End If
     
   If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
      If CDbl(g_rst_GenAux!RETENCION) - CDbl(g_rst_GenAux!DEVOLUCION) = 0 Then
         fs_Validar_Retencion_Garantia = True
      Else
         fs_Validar_Retencion_Garantia = False
      End If
   End If
End Function

Private Function fs_Validar_Mto_CFiGar() As Boolean
   fs_Validar_Mto_CFiGar = False
        
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "   SELECT NVL(MAECFI_GARFIA,0) AS  CARTA_FIANZA," 'MAECFI_IMPFIA
   g_str_Parame = g_str_Parame & "          (SELECT NVL(SUM(NVL(MAEGAR_MTOGAR_INM,0) + NVL(MAEGAR_MTOGAR_ES1,0) + NVL(MAEGAR_MTOGAR_ES2,0) + NVL(MAEGAR_MTOGAR_DE1,0) + NVL(MAEGAR_MTOGAR_DE2,0)),0)"
   g_str_Parame = g_str_Parame & "             FROM TPR_MAEGAR"
   g_str_Parame = g_str_Parame & "            WHERE MAECFI_NUMREF = MAEGAR_NUMREF ) GARANTIA"
   g_str_Parame = g_str_Parame & "     FROM TPR_MAECFI"
   g_str_Parame = g_str_Parame & "    WHERE MAECFI_NUMREF = '" & CStr(moddat_g_str_DesIte) & "'" 'moddat_g_str_NumFia
    
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
      Exit Function
   End If
   
   If g_rst_GenAux.BOF And g_rst_GenAux.EOF Then
      g_rst_GenAux.Close
      Set g_rst_GenAux = Nothing
      Exit Function
   End If
     
   If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
      If CDbl(ipp_ImpCar.Value) >= CDbl(g_rst_GenAux!GARANTIA) Then
         fs_Validar_Mto_CFiGar = True
      End If
   End If
End Function

Private Function fs_Validar_Mto_ComPag() As Boolean
   fs_Validar_Mto_ComPag = False
      
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT MAECFI_NUMREF, NVL(COMISION_PAGADO,0) COMISION_PAGADO "
   g_str_Parame = g_str_Parame & "    FROM TPR_MAECFI"
   g_str_Parame = g_str_Parame & "           LEFT JOIN (SELECT MAERDE_NUMREF, NVL(SUM(NVL(MAERDE_IMPORT,0)),0) COMISION_PAGADO "
   g_str_Parame = g_str_Parame & "                        FROM TPR_MAERDE "
   g_str_Parame = g_str_Parame & "                       WHERE MAERDE_CODIGO = 1 OR MAERDE_CODIGO = 2 "
   g_str_Parame = g_str_Parame & "                       GROUP BY MAERDE_NUMREF) B ON MAERDE_NUMREF = MAECFI_NUMREF"
   g_str_Parame = g_str_Parame & "   WHERE MAECFI_NUMREF =  '" & CStr(moddat_g_str_DesIte) & "' " 'moddat_g_str_NumFia
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
     Exit Function
   End If
   
   If g_rst_GenAux.BOF And g_rst_GenAux.EOF Then
      fs_Validar_Mto_ComPag = True
      g_rst_GenAux.Close
      Set g_rst_GenAux = Nothing
      Exit Function
   End If
     
   If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
      If CDbl(g_rst_GenAux!COMISION_PAGADO) <= CDbl(pnl_ValCom.Caption) Then
         fs_Validar_Mto_ComPag = True
      End If
   End If
End Function

Private Function fs_Validar_Mto_FRePag() As Boolean
   fs_Validar_Mto_FRePag = False
      
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT MAECFI_NUMREF, NVL(MAECFI_GARFIA,0) FONDOS , "
   g_str_Parame = g_str_Parame & "         NVL((SELECT NVL(SUM(NVL(MAERDE_IMPORT,0)),0) "
   g_str_Parame = g_str_Parame & "                FROM TPR_MAERDE "
   g_str_Parame = g_str_Parame & "               WHERE MAERDE_CODIGO IN (3,19) "
   g_str_Parame = g_str_Parame & "                 AND MAERDE_NUMREF = MAECFI_NUMREF "
   g_str_Parame = g_str_Parame & "               GROUP BY MAERDE_NUMREF),0) FONDOS_PAGADO "
   g_str_Parame = g_str_Parame & "    FROM TPR_MAECFI "
   g_str_Parame = g_str_Parame & "   WHERE MAECFI_NUMREF =  '" & CStr(moddat_g_str_DesIte) & "' " 'moddat_g_str_NumFia
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
     Exit Function
   End If
   
   If g_rst_GenAux.BOF And g_rst_GenAux.EOF Then
      fs_Validar_Mto_FRePag = True
      g_rst_GenAux.Close
      Set g_rst_GenAux = Nothing
      Exit Function
   End If
     
   If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
      If CDbl(g_rst_GenAux!FONDOS_PAGADO) <= CDbl(ipp_ImpCar.Value) Then
         fs_Validar_Mto_FRePag = True
      End If
   End If
End Function

Private Sub fs_GeneraAsiento(ByVal p_NumRef As String, ByVal p_NumDoc As String, ByVal p_RazSoc As String, ByVal p_CtaDeb As String, ByVal p_CtaHab As String, ByVal p_Importe As Double)
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
   
   If Year(r_str_FecCon) <> CInt(moddat_g_str_CodAno) Or Month(r_str_FecCon) <> CInt(moddat_g_str_CodMes) Then
      r_str_FecCon = moddat_g_str_FecSis
   End If
   
   If moddat_g_str_CodMod = "005" Then
      r_str_Glosa = "ADENDA"
   ElseIf moddat_g_str_CodMod = "008" Then
      r_str_Glosa = "CARTA SERIEDAD OFERTA"
   Else
      r_str_Glosa = "CARTA FIANZA"
   End If

   'Insertar en CABECERA
   Call modprc_fs_Inserta_CabAsi(r_arr_LogPro, r_str_Origen, CInt(moddat_g_str_CodAno), CInt(moddat_g_str_CodMes), r_int_NumLib, r_int_NumAsi, Format(1, "000"), r_dbl_TipCam, r_str_TipNot, Trim(r_str_Glosa), r_str_FecCon, "1")

   '*************************************************
   'GENERACION DE ASIENTOS CONTABLES DE CARTA FIANZA
   '*************************************************
   For r_int_ConAux = 1 To 2
      
       r_dbl_importe = p_Importe

       If r_int_ConAux = 1 Then r_str_DebHab = "D": r_str_CtaCtb = p_CtaDeb Else r_str_DebHab = "H": r_str_CtaCtb = p_CtaHab
       
       r_str_Glosa = IIf(moddat_g_str_CodMod = "005", "AD", IIf(moddat_g_str_CodMod = "008", "CSO", "CF")) & Trim(p_NumRef) & "/" & Trim(p_NumDoc) & "/" & Trim(p_RazSoc)
       r_str_Glosa = Trim(Mid(r_str_Glosa, 1, 60))
        
       If (r_dbl_importe > 0) Then
           r_int_NumIte = r_int_NumIte + 1
           r_dbl_MtoSol = Format(r_dbl_importe, "###,###,##0.00")
           r_dbl_MtoDol = Format(0, "###,###,##0.00")
           
           Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, CInt(moddat_g_str_CodAno), CInt(moddat_g_str_CodMes), r_int_NumLib, r_int_NumAsi, r_int_NumIte, r_str_CtaCtb, CDate(r_str_FecCon), r_str_Glosa, r_str_DebHab, r_dbl_MtoSol, r_dbl_MtoDol, 1, CDate(r_str_FecCon))
           r_dbl_importe = 0
       End If
   Next r_int_ConAux
End Sub

Private Sub cmd_NueGar_Click()
'   moddat_g_int_FlgGrb_1 = 4
   'If fs_Validar = True Then
      frm_Ges_TecPro_05.Show 1
   'End If
End Sub

Private Sub cmd_NueTas_Click()
'   moddat_g_int_FlgGrb_1 = 5
   'If fs_Validar = True Then
      frm_Ges_TecPro_15.Show 1
   'End If
End Sub

Private Sub cmd_Salida_Click()
   Call fs_Limpia
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   pnl_Descri.Caption = "Techo Propio - Registro"
   
   Call gs_CentraForm(Me)
   Call fs_Limpia
   Call fs_Inicia
   If moddat_g_int_FlgGrb_1 = 2 Or moddat_g_int_FlgGrb_1 = 6 Then
'      If Mid(moddat_g_str_NumFia, 1, 1) = 2 Then 'If moddat_g_str_CodMod = "005" Then
'         If moddat_g_int_FlgGrb_1 = 2 Then pnl_Descri.Caption = "Adenda - Actualizar"
'         If moddat_g_int_FlgGrb_1 = 6 Then pnl_Descri.Caption = "Adenda - Renovación"
'      Else
'         If moddat_g_int_FlgGrb_1 = 2 Then pnl_Descri.Caption = "Carta Fianza - Actualizar"
'         If moddat_g_int_FlgGrb_1 = 6 Then pnl_Descri.Caption = "Carta Fianza - Renovación"
'      End If
      If moddat_g_int_FlgGrb_1 = 2 Then pnl_Descri.Caption = "Techo Propio - Modificación"
      If moddat_g_int_FlgGrb_1 = 6 Then pnl_Descri.Caption = "Techo Propio - Renovación"
      
      If moddat_g_str_TipCre = 1 And moddat_g_int_TipPan = 1 Then
         Call fs_Buscar_Credito_Indirecto
      ElseIf moddat_g_str_TipCre = 2 And moddat_g_int_TipPan = 0 Then
         Call fs_Buscar_Credito_Directo
      End If
      
   End If
   If moddat_g_int_FlgGrb_1 = 1 Then
      Call fs_Activa(True)
   End If
  ' Call gs_SetFocus(ipp_FecEmi) 'txt_NumRef
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   'Proyecto
   Call moddat_gs_Carga_Proyec(cmb_NomPry, l_arr_Proyec)
   
   'Producto
   cmb_Produc.Clear
   ReDim l_arr_Produc(0)
   
   'Producto
   ReDim Preserve l_arr_Produc(UBound(l_arr_Produc) + 1)
   cmb_Produc.AddItem Trim$("CREDITO INDIRECTO - TECHO PROPIO")
   l_arr_Produc(UBound(l_arr_Produc)).Genera_Codigo = Trim$("026")
   l_arr_Produc(UBound(l_arr_Produc)).Genera_Nombre = Trim$("CREDITO INDIRECTO - TECHO PROPIO")
   
   ReDim Preserve l_arr_Produc(UBound(l_arr_Produc) + 1)
   cmb_Produc.AddItem Trim$("BONO DE REFORZAMIENO ESTUCTURAL")
   l_arr_Produc(UBound(l_arr_Produc)).Genera_Codigo = Trim$("027")
   l_arr_Produc(UBound(l_arr_Produc)).Genera_Nombre = Trim$("BPVV - BONO DE REFORZAMIENO ESTUCTURAL")
   
   ReDim Preserve l_arr_Produc(UBound(l_arr_Produc) + 1)
   cmb_Produc.AddItem Trim$("CREDITO DIRECTO")
   l_arr_Produc(UBound(l_arr_Produc)).Genera_Codigo = Trim$("008")
   l_arr_Produc(UBound(l_arr_Produc)).Genera_Nombre = Trim$("CREDITO DIRECTO")
   
'''   'Sub-Producto
'''   cmb_SubPrd.Clear
'''   ReDim l_arr_SubPrd(0)
'''
'''   ReDim Preserve l_arr_SubPrd(UBound(l_arr_SubPrd) + 1)
'''   l_arr_SubPrd(UBound(l_arr_SubPrd)).Genera_Codigo = Trim$("001")
'''   l_arr_SubPrd(UBound(l_arr_SubPrd)).Genera_Nombre = Trim$("AVN - ADQUISICION VIVIENDA NUEVA")
'''   cmb_SubPrd.AddItem Trim$("AVN - AQUISICION VIVIENDA NUEVA")
'''
'''   ReDim Preserve l_arr_SubPrd(UBound(l_arr_SubPrd) + 1)
'''   l_arr_SubPrd(UBound(l_arr_SubPrd)).Genera_Codigo = Trim$("002")
'''   l_arr_SubPrd(UBound(l_arr_SubPrd)).Genera_Nombre = Trim$("CSP - CONSTRUCCION SITIO PROPIO")
'''   cmb_SubPrd.AddItem Trim$("CSP - CONSTRUCCION SITIO PROPIO")
'''
'''   ReDim Preserve l_arr_SubPrd(UBound(l_arr_SubPrd) + 1)
'''   l_arr_SubPrd(UBound(l_arr_SubPrd)).Genera_Codigo = Trim$("003")
'''   l_arr_SubPrd(UBound(l_arr_SubPrd)).Genera_Nombre = Trim$("MV - MEJORAMIENTO DE VIVIENDAO")
'''   cmb_SubPrd.AddItem Trim$("CMV - MEJORAMIENTO DE VIVIENDA")
'''
'''   ReDim Preserve l_arr_SubPrd(UBound(l_arr_SubPrd) + 1)
'''   l_arr_SubPrd(UBound(l_arr_SubPrd)).Genera_Codigo = Trim$("004")
'''   l_arr_SubPrd(UBound(l_arr_SubPrd)).Genera_Nombre = Trim$("BPV - BONO DE REFORZAMIENO ESTUCTURAL")
'''   cmb_SubPrd.AddItem Trim$("BPV - BONO DE REFORZAMIENO ESTUCTURAL")
'''
'''   'Modalidad
'''   cmb_Modali.Clear
'''   ReDim l_arr_Modali(0)
'''
'''   ReDim Preserve l_arr_Modali(UBound(l_arr_Modali) + 1)
'''   l_arr_Modali(UBound(l_arr_Modali)).Genera_Codigo = Trim$("001")
'''   l_arr_Modali(UBound(l_arr_Modali)).Genera_Nombre = Trim$("BONO DE EMERGENCIA")
'''   cmb_Modali.AddItem Trim$("BONO DE EMERGENCIA")
'''
'''   ReDim Preserve l_arr_Modali(UBound(l_arr_Modali) + 1)
'''   l_arr_Modali(UBound(l_arr_Modali)).Genera_Codigo = Trim$("003")
'''   l_arr_Modali(UBound(l_arr_Modali)).Genera_Nombre = Trim$("BONO NORMAL")
'''   cmb_Modali.AddItem Trim$("BONO NORMAL")
'''
'''   ReDim Preserve l_arr_Modali(UBound(l_arr_Modali) + 1)
'''   l_arr_Modali(UBound(l_arr_Modali)).Genera_Codigo = Trim$("004")
'''   l_arr_Modali(UBound(l_arr_Modali)).Genera_Nombre = Trim$("ADQUISICION VIVIENDA NUEVA - CF")
'''   cmb_Modali.AddItem Trim$("ADQUISICION VIVIENDA NUEVA - CF")
'''
'''   ReDim Preserve l_arr_Modali(UBound(l_arr_Modali) + 1)
'''   l_arr_Modali(UBound(l_arr_Modali)).Genera_Codigo = Trim$("005")
'''   l_arr_Modali(UBound(l_arr_Modali)).Genera_Nombre = Trim$("ADQUISICION VIVIENDA NUEVA - AD")
'''   cmb_Modali.AddItem Trim$("ADQUISICION VIVIENDA NUEVA - AD")
'''
'''   ReDim Preserve l_arr_Modali(UBound(l_arr_Modali) + 1)
'''   l_arr_Modali(UBound(l_arr_Modali)).Genera_Codigo = Trim$("006")
'''   l_arr_Modali(UBound(l_arr_Modali)).Genera_Nombre = Trim$("CONSTRUCCION SITIO PROPIO")
'''   cmb_Modali.AddItem Trim$("CONSTRUCCION SITIO PROPIO")
'''
'''   ReDim Preserve l_arr_Modali(UBound(l_arr_Modali) + 1)
'''   l_arr_Modali(UBound(l_arr_Modali)).Genera_Codigo = Trim$("007")
'''   l_arr_Modali(UBound(l_arr_Modali)).Genera_Nombre = Trim$("MEJORAMIENTO DE VIVIENDA")
'''   cmb_Modali.AddItem Trim$("MEJORAMIENTO DE VIVIENDA")
'''
'''   'CSO
'''   ReDim Preserve l_arr_Modali(UBound(l_arr_Modali) + 1)
'''   l_arr_Modali(UBound(l_arr_Modali)).Genera_Codigo = Trim$("008")
'''   l_arr_Modali(UBound(l_arr_Modali)).Genera_Nombre = Trim$("CARTA DE SERIEDAD DE OFERTA")
'''   cmb_Modali.AddItem Trim$("CARTA DE SERIEDAD DE OFERTA")
'''
'''   ReDim Preserve l_arr_Modali(UBound(l_arr_Modali) + 1)
'''   l_arr_Modali(UBound(l_arr_Modali)).Genera_Codigo = Trim$("002")
'''   l_arr_Modali(UBound(l_arr_Modali)).Genera_Nombre = Trim$("BPV - BONO DE REFORZAMIENO ESTUCTURAL")
'''   cmb_Modali.AddItem Trim$("BPV - BONO DE REFORZAMIENO ESTUCTURAL")
   
   'Tipo de Renovación
   cmb_TipRen.AddItem Trim$("POR PLAZO")
   cmb_TipRen.ItemData(cmb_TipRen.NewIndex) = CInt(1)
   cmb_TipRen.AddItem Trim$("POR MONTO")
   cmb_TipRen.ItemData(cmb_TipRen.NewIndex) = CInt(2)
   
   
   'Porcentaje para Garantizado
   Call moddat_gs_Carga_LisIte_Combo(cmb_PorGar, 1, "535")
   
'   cmb_Porcen.AddItem Trim$("5")
'   cmb_Porcen.ItemData(cmb_Porcen.NewIndex) = CInt(5)
'   cmb_Porcen.AddItem Trim$("10")
'   cmb_Porcen.ItemData(cmb_Porcen.NewIndex) = CInt(10)
'   cmb_Porcen.ListIndex = 0
   
   'Recurso
   cmb_TipRec.Clear
   ReDim l_arr_TipRec(0)
   
   ReDim Preserve l_arr_TipRec(UBound(l_arr_TipRec) + 1)
   cmb_TipRec.AddItem Trim$("BONO")
   l_arr_TipRec(UBound(l_arr_TipRec)).Genera_Codigo = CInt(1)
   l_arr_TipRec(UBound(l_arr_TipRec)).Genera_Nombre = Trim$("BONO")
   
   ReDim Preserve l_arr_TipRec(UBound(l_arr_TipRec) + 1)
   cmb_TipRec.AddItem Trim$("AHORRO")
   l_arr_TipRec(UBound(l_arr_TipRec)).Genera_Codigo = CInt(2)
   l_arr_TipRec(UBound(l_arr_TipRec)).Genera_Nombre = Trim$("AHORRO")
   
'   ReDim Preserve l_arr_TipRec(UBound(l_arr_TipRec) + 1)
'   cmb_TipRec.AddItem Trim$("ABONO/AHORRO")
'   l_arr_TipRec(UBound(l_arr_TipRec)).Genera_Codigo = CInt(3)
'   l_arr_TipRec(UBound(l_arr_TipRec)).Genera_Nombre = Trim$("ABONO/AHORRO")
   
'   cmb_TipRec.AddItem Trim$("BONO")
'   cmb_TipRec.ItemData(cmb_TipRec.NewIndex) = CInt(1)
'   cmb_TipRec.AddItem Trim$("AHORRO")
'   cmb_TipRec.ItemData(cmb_TipRec.NewIndex) = CInt(2)
'   cmb_TipRec.AddItem Trim$("ABONO/AHORRO")
'   cmb_TipRec.ItemData(cmb_TipRec.NewIndex) = CInt(3)
   
   'Moneda
   Call moddat_gs_Carga_LisIte_Combo(cmb_Moneda, 1, "204")
   
   'Porcentaje de Retención
   Call fs_Parametros 'Porcentaje de Retención y Fecha de Vencimiento de ETE
   
   'Plazos (días)
   If moddat_g_int_FlgGrb_1 = 6 Then
      ipp_PlaCar.MinValue = 0 '31 '45
      ipp_PlaCar.MaxValue = 180
      
   Else
      If moddat_g_int_FlgGrb_1 = 1 Then
         ipp_PlaCar.MinValue = 45 '90
         ipp_PlaCar.MaxValue = 360
      Else
         'Verificar si ya ha tenido una renovación
         If fs_Validar_Renovacion(moddat_g_str_DesIte) Then 'moddat_g_str_NumFia
            ipp_PlaCar.MinValue = 0 '31 '45
            ipp_PlaCar.MaxValue = 180
         Else
            ipp_PlaCar.MinValue = 45 '90
            ipp_PlaCar.MaxValue = 360
         End If
      End If
   End If
            
   'Tipo de Línea
   cmb_TipLin.Clear
   ReDim l_arr_TipLin(0)
   
   ReDim Preserve l_arr_TipLin(UBound(l_arr_TipLin) + 1)
   cmb_TipLin.AddItem Trim$("NO REVOLVENTE")
   l_arr_TipLin(UBound(l_arr_TipLin)).Genera_Codigo = CInt(1)
   l_arr_TipLin(UBound(l_arr_TipLin)).Genera_Nombre = Trim$("NO REVOLVENTE")
   
   ReDim Preserve l_arr_TipLin(UBound(l_arr_TipLin) + 1)
   cmb_TipLin.AddItem Trim$("REVOLVENTE")
   l_arr_TipLin(UBound(l_arr_TipLin)).Genera_Codigo = CInt(2)
   l_arr_TipLin(UBound(l_arr_TipLin)).Genera_Nombre = Trim$("REVOLVENTE")
   
'
'   cmb_TipLin.AddItem Trim$("NO REVOLVENTE")
'   cmb_TipLin.ItemData(cmb_TipLin.NewIndex) = CInt(1)
'   cmb_TipLin.AddItem Trim$("REVOLVENTE")
'   cmb_TipLin.ItemData(cmb_TipLin.NewIndex) = CInt(2)
   
   pnl_TipDoc.Caption = moddat_gf_Consulta_ParDes("118", moddat_g_int_TipDoc)
   pnl_NroDoc.Caption = moddat_g_str_NumDoc
   pnl_RazSoc.Caption = moddat_g_str_NomCli
   pnl_TipEmp.Caption = moddat_g_str_Descri
   
'   'Año y mes
'   Call moddat_gf_ConsultaPerMesActivo("000001", 1, moddat_g_str_FecIni, moddat_g_str_FecFin, moddat_g_str_CodMes, moddat_g_str_CodAno)
End Sub

Private Function fs_Validar_Renovacion(ByVal p_NumRef As String) As Boolean
   fs_Validar_Renovacion = False
  
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT COUNT(*) CONTADOR "
   g_str_Parame = g_str_Parame & "   FROM TPR_MAECFI A "
   g_str_Parame = g_str_Parame & "  WHERE MAECFI_REFORI = (SELECT MAECFI_REFORI"
   g_str_Parame = g_str_Parame & "                           FROM TPR_MAECFI"
   g_str_Parame = g_str_Parame & "                          WHERE MAECFI_NUMREF = '" & CStr(p_NumRef) & "' "
   g_str_Parame = g_str_Parame & "                            AND MAECFI_TIPDOC = " & CStr(moddat_g_int_TipDoc) & " "
   g_str_Parame = g_str_Parame & "                            AND MAECFI_NUMDOC = '" & CStr(moddat_g_str_NumDoc) & "' ) "
   g_str_Parame = g_str_Parame & "    AND MAECFI_TIPDOC = " & CStr(moddat_g_int_TipDoc) & " "
   g_str_Parame = g_str_Parame & "    AND MAECFI_NUMDOC = '" & CStr(moddat_g_str_NumDoc) & "' "
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Function
   End If
       
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
     g_rst_Princi.Close
     Set g_rst_Princi = Nothing
     Exit Function
   End If
   
   g_rst_Princi.MoveFirst
   
   If CInt(g_rst_Princi!CONTADOR) > 1 Then
      fs_Validar_Renovacion = True
   End If
End Function

Private Sub fs_Parametros()
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT MAEETE_PORRET, MAEETE_FECVCT, MAECFI_PORRET, MAEETE_PORTEA "
   g_str_Parame = g_str_Parame & "   FROM TPR_MAEETE LEFT JOIN TPR_MAECFI ON MAEETE_TIPDOC = MAECFI_TIPDOC AND MAEETE_NUMDOC = MAECFI_NUMDOC "
   g_str_Parame = g_str_Parame & "  WHERE MAEETE_TIPDOC = " & CStr(moddat_g_int_TipDoc) & " "
   g_str_Parame = g_str_Parame & "    AND MAEETE_NUMDOC = '" & CStr(moddat_g_str_NumDoc) & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
       
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   If Not IsNull(g_rst_Princi!MAECFI_PORRET) And g_rst_Princi!MAECFI_PORRET > 0 Then
      ipp_PorRet.Value = g_rst_Princi!MAECFI_PORRET
   Print
      If Not IsNull(g_rst_Princi!MAEETE_PORRET) Then
         ipp_PorRet.Value = g_rst_Princi!MAEETE_PORRET
      End If
   End If
   If Not IsNull(g_rst_Princi!MAEETE_PORTEA) And g_rst_Princi!MAEETE_PORTEA > 0 Then
      ipp_PorTEA.Value = g_rst_Princi!MAEETE_PORTEA
   Else
      If Not IsNull(g_rst_Princi!MAEETE_PORTEA) Then
         ipp_PorTEA.Value = g_rst_Princi!MAEETE_PORTEA
      End If
   End If
   If Not IsNull(g_rst_Princi!MAEETE_FECVCT) Then
      l_str_FecVct = g_rst_Princi!MAEETE_FECVCT
   End If
End Sub

Private Sub fs_Limpia()
   cmb_Produc.ListIndex = -1
   cmb_SubPrd.ListIndex = -1
   cmb_Moneda.ListIndex = -1
   cmb_TipRec.ListIndex = -1
   cmb_NomPry.ListIndex = -1
   ipp_PlaCar.Value = 0
   txt_NumRef.Text = ""
   txt_CodEte.Text = ""
   txt_ParReg.Text = ""
   txt_CodPry.Text = ""
   txt_NumAde.Text = ""
   ipp_FecEmi.Text = Format(date, "dd/mm/yyyy")
   ipp_FecVct.Text = Format(date, "dd/mm/yyyy")
   ipp_ImpCar.Value = Format(0, "###,###,###,##0.00")
   pnl_ValGar.Caption = Format(0, "###,###,###,##0.00") & "  "
   pnl_ValCom.Caption = Format(0, "###,###,###,##0.00") & "  "
   pnl_ValMin.Caption = Format(0, "###,###,###,##0.00") & "  "
   ipp_TEACom.Value = "0.00%"
   ipp_PorRet.Value = "0.00%"
   cmb_PorGar.ListIndex = -1
   chk_NesCli.Value = False
   l_str_FecVct = ""
   moddat_g_dbl_TotGar = 0
   moddat_g_str_DesObs = ""
   '   l_dbl_MtoCFi = 0
End Sub
Private Sub fs_Buscar_Credito_Directo()
Dim x As String
Dim r_str_NomPry As String
Dim r_str_NOMPRO As String
Dim r_str_TipPry As String

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT MAECFI_NUMREF, MAECFI_EMIFIA, MAECFI_PLZFIA, MAECFI_VTOFIA, MAECFI_MONFIA, MAECFI_CODPRD, MAECFI_CODSUB, "
   g_str_Parame = g_str_Parame & "        MAECFI_CODMOD, MAECFI_IMPFIA, MAECFI_GARFIA, MAECFI_TASFIA, MAECFI_COMFIA, MAEETE_FECVCT, "
   g_str_Parame = g_str_Parame & "        MAECFI_CODPRY, MAECFI_NUMADE, MAECFI_NUMANT, MAECFI_PORTEA, MAECFI_NOMPRY, MAECFI_TIPLIN "
   g_str_Parame = g_str_Parame & "   FROM TPR_MAECFI INNER JOIN TPR_MAEETE ON MAEETE_TIPDOC = MAECFI_TIPDOC AND MAEETE_NUMDOC = MAECFI_NUMDOC "
   g_str_Parame = g_str_Parame & "  WHERE MAECFI_NUMREF = '" & CStr(moddat_g_str_NumFia) & "' "
   g_str_Parame = g_str_Parame & "    AND MAECFI_TIPDOC = " & CStr(moddat_g_int_TipDoc) & " "
   g_str_Parame = g_str_Parame & "    AND MAECFI_NUMDOC = '" & CStr(moddat_g_str_NumDoc) & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
       
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
     g_rst_Princi.Close
     Set g_rst_Princi = Nothing
     Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   If Not IsNull((g_rst_Princi!MAECFI_CODPRD)) Then
      moddat_g_str_CodPrd = CStr(g_rst_Princi!MAECFI_CODPRD)
      cmb_Produc.ListIndex = gf_Busca_Arregl(l_arr_Produc, Trim(CStr(g_rst_Princi!MAECFI_CODPRD) & "")) - 1
   End If
   If Not IsNull((g_rst_Princi!MAECFI_CODSUB)) Then
      moddat_g_str_CodSub = CStr(g_rst_Princi!MAECFI_CODSUB)
      cmb_SubPrd.ListIndex = gf_Busca_Arregl(l_arr_SubPrd, Trim(CStr(g_rst_Princi!MAECFI_CODSUB) & "")) - 1
   End If
   If Not IsNull((g_rst_Princi!MAECFI_CODMOD)) Then
      moddat_g_str_CodMod = CStr(g_rst_Princi!MAECFI_CODMOD)
      cmb_Modali.ListIndex = gf_Busca_Arregl(l_arr_Modali, Trim(CStr(g_rst_Princi!MAECFI_CODMOD) & "")) - 1
   End If
   
   If Not IsNull(g_rst_Princi!MAECFI_NUMANT) Then
      txt_NumRef.Text = CStr(Trim(g_rst_Princi!MAECFI_NUMANT))
   Else
      txt_NumRef.Text = CStr(Trim(g_rst_Princi!MAECFI_NUMREF))
   End If
   ipp_FecEmi.Text = Format(gf_FormatoFecha(CStr(g_rst_Princi!MAECFI_EMIFIA)), "dd/mm/yyyy")
      
   ipp_FecVct.Text = Format(gf_FormatoFecha(CStr(g_rst_Princi!MAECFI_VTOFIA)), "dd/mm/yyyy")
   
   l_str_FEmRen = CStr(g_rst_Princi!MAECFI_EMIFIA)
   l_str_FVeRen = CStr(g_rst_Princi!MAECFI_VTOFIA)
      
   cmb_Moneda.Text = moddat_gf_Consulta_ParDes("204", g_rst_Princi!MAECFI_MONFIA)
   
   ipp_ImpCar.Value = Format(g_rst_Princi!MAECFI_IMPFIA, "###,###,###,##0.00")
   r_dbl_ValFia = Format(g_rst_Princi!MAECFI_IMPFIA, "###,###,###,##0.00")
   
   ipp_TasMor.Value = g_rst_Princi!MAECFI_TASFIA
   ipp_ValCom.Value = Format(g_rst_Princi!MAECFI_COMFIA, "###,###,###,##0.00") & "  "
  
   If Not IsNull(g_rst_Princi!MAECFI_PORTEA) Then
      ipp_PorTEA.Value = g_rst_Princi!MAECFI_PORTEA
   End If
   
   If Not IsNull(g_rst_Princi!MAECFI_NOMPRY) Then
      Call modmip_gs_Consulta_NomPry(CStr(g_rst_Princi!MAECFI_NOMPRY), r_str_NomPry, r_str_NOMPRO, r_str_TipPry)
      If r_str_NomPry <> "" Then
         cmb_NomPry.Text = r_str_NomPry
      End If
   End If
   
   If Not IsNull((g_rst_Princi!MAECFI_TIPLIN)) Then
      cmb_TipLin.ListIndex = gf_Busca_Arregl(l_arr_TipLin, Trim(CStr(g_rst_Princi!MAECFI_TIPLIN) & "")) - 1
   End If
   
   If moddat_g_int_FlgGrb_1 = 6 Then
      
      ipp_PlaCar.Text = 31 '45
      ipp_FecEmi.Text = Format(gf_FormatoFecha(CStr(g_rst_Princi!MAECFI_VTOFIA)), "dd/mm/yyyy")
      ipp_FecVct.Text = Format(CDate(DateAdd("D", ipp_PlaCar.Text, Me.ipp_FecEmi.Value)), "DD/MM/YYYY")
      Call fs_Calcular
      Call fs_Activa(False)
   Else
      ipp_PlaCar.Text = g_rst_Princi!MAECFI_PLZFIA
      txt_NumRef.Enabled = False
      lbl_TipRen.Visible = False
      cmb_TipRen.Visible = False
      cmd_NueGar.Enabled = False
      cmd_NueTas.Enabled = False
   End If
End Sub
Private Sub fs_Buscar_Credito_Indirecto()
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT MAECFI_NUMREF, MAECFI_EMIFIA, MAECFI_PLZFIA, MAECFI_VTOFIA, MAECFI_MONFIA, MAECFI_CODPRD, MAECFI_CODSUB, "
   g_str_Parame = g_str_Parame & "        MAECFI_CODMOD, MAECFI_IMPFIA, MAECFI_GARFIA, MAECFI_TASFIA, MAECFI_COMFIA, MAECFI_MINFIA, MAECFI_PARREG, "
   g_str_Parame = g_str_Parame & "        CASE WHEN MAECFI_PORRET <> MAEETE_PORRET THEN MAECFI_PORRET ELSE MAEETE_PORRET END AS PORC_RETENCION, "
   g_str_Parame = g_str_Parame & "        MAEETE_FECVCT, MAECFI_CODPRY, MAECFI_NUMADE, MAECFI_NUMANT, MAECFI_CODETE, MAECFI_TIPREC, MAECFI_NOMPRY, "
   g_str_Parame = g_str_Parame & "        MAECFI_PORGAR, MAECFI_NOCLIE "
   g_str_Parame = g_str_Parame & "   FROM TPR_MAECFI INNER JOIN TPR_MAEETE ON MAEETE_TIPDOC = MAECFI_TIPDOC AND MAEETE_NUMDOC = MAECFI_NUMDOC "
   g_str_Parame = g_str_Parame & "  WHERE MAECFI_NUMREF= '" & CStr(moddat_g_str_DesIte) & "' "
   g_str_Parame = g_str_Parame & "    AND MAECFI_TIPDOC = " & CStr(moddat_g_int_TipDoc) & " "
   g_str_Parame = g_str_Parame & "    AND MAECFI_NUMDOC = '" & CStr(moddat_g_str_NumDoc) & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
       
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
     g_rst_Princi.Close
     Set g_rst_Princi = Nothing
     Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   
   If Not IsNull((g_rst_Princi!MAECFI_CODPRD)) Then
      moddat_g_str_CodPrd = CStr(g_rst_Princi!MAECFI_CODPRD)
      cmb_Produc.ListIndex = gf_Busca_Arregl(l_arr_Produc, Trim(CStr(g_rst_Princi!MAECFI_CODPRD) & "")) - 1
   End If
   If Not IsNull((g_rst_Princi!MAECFI_CODSUB)) Then
      moddat_g_str_CodSub = CStr(g_rst_Princi!MAECFI_CODSUB)
      cmb_SubPrd.ListIndex = gf_Busca_Arregl(l_arr_SubPrd, Trim(CStr(g_rst_Princi!MAECFI_CODSUB) & "")) - 1
   End If
   If Not IsNull((g_rst_Princi!MAECFI_CODMOD)) Then
      moddat_g_str_CodMod = CStr(g_rst_Princi!MAECFI_CODMOD)
      cmb_Modali.ListIndex = gf_Busca_Arregl(l_arr_Modali, Trim(CStr(g_rst_Princi!MAECFI_CODMOD) & "")) - 1
   End If
   
   If Not IsNull(g_rst_Princi!MAECFI_NUMANT) Then
      txt_NumRef.Text = CStr(Trim(g_rst_Princi!MAECFI_NUMANT))
   Else
      txt_NumRef.Text = CStr(Trim(g_rst_Princi!MAECFI_NUMREF))
   End If
   ipp_FecEmi.Text = Format(gf_FormatoFecha(CStr(g_rst_Princi!MAECFI_EMIFIA)), "dd/mm/yyyy")
   ipp_PlaCar.Text = CStr(g_rst_Princi!MAECFI_PLZFIA)
   ipp_FecVct.Text = Format(gf_FormatoFecha(CStr(g_rst_Princi!MAECFI_VTOFIA)), "dd/mm/yyyy")
   cmb_Moneda.Text = moddat_gf_Consulta_ParDes("204", g_rst_Princi!MAECFI_MONFIA)
   
   l_str_FEmRen = CStr(g_rst_Princi!MAECFI_EMIFIA)
   l_str_FVeRen = CStr(g_rst_Princi!MAECFI_VTOFIA)
   
   ipp_ImpCar.Value = Format(g_rst_Princi!MAECFI_IMPFIA, "###,###,###,##0.00")
   l_dbl_ValImp = g_rst_Princi!MAECFI_IMPFIA
   r_dbl_ValFia = Format(g_rst_Princi!MAECFI_IMPFIA, "###,###,###,##0.00")
   pnl_ValGar.Caption = Format(g_rst_Princi!MAECFI_GARFIA, "###,###,###,##0.00") & "  "
   l_dbl_ValGar = g_rst_Princi!MAECFI_GARFIA
   ipp_TEACom.Value = g_rst_Princi!MAECFI_TASFIA
   pnl_ValCom.Caption = Format(g_rst_Princi!MAECFI_COMFIA, "###,###,###,##0.00") & "  "
   pnl_ValMin.Caption = Format(g_rst_Princi!MAECFI_MINFIA, "###,###,###,##0.00") & "  "
   
   If Not IsNull(g_rst_Princi!PORC_RETENCION) Then
      ipp_PorRet.Value = g_rst_Princi!PORC_RETENCION
   End If
   
   If Not IsNull(g_rst_Princi!MAECFI_PARREG) Then
      txt_ParReg.Text = g_rst_Princi!MAECFI_PARREG
   End If
   
   If Not IsNull(g_rst_Princi!MAECFI_CODETE) Then
      txt_CodEte.Text = g_rst_Princi!MAECFI_CODETE
   End If
   
   If Not IsNull(g_rst_Princi!MAECFI_CODPRY) Then
      txt_CodPry.Text = g_rst_Princi!MAECFI_CODPRY
   End If
   
   If Not IsNull(g_rst_Princi!MAECFI_NUMADE) Then
      txt_NumAde.Text = g_rst_Princi!MAECFI_NUMADE
   End If
   
   If Not IsNull((g_rst_Princi!MAECFI_TIPREC)) Then
      moddat_g_str_CodGrp = CStr(g_rst_Princi!MAECFI_TIPREC)
      cmb_TipRec.ListIndex = gf_Busca_Arregl(l_arr_TipRec, Trim(CStr(g_rst_Princi!MAECFI_TIPREC) & "")) - 1
   End If
   
   If Not IsNull((g_rst_Princi!MAECFI_NOMPRY)) Then
      moddat_g_str_NomCom = CStr(g_rst_Princi!MAECFI_NOMPRY)
      cmb_NomPry.ListIndex = gf_Busca_Arregl(l_arr_Proyec, Trim(CStr(g_rst_Princi!MAECFI_NOMPRY) & "")) - 1
   End If
   
   If Not IsNull(g_rst_Princi!MAECFI_PORGAR) And g_rst_Princi!MAECFI_PORGAR <> 0 Then
      cmb_PorGar.Text = moddat_gf_Consulta_ParDes("535", g_rst_Princi!MAECFI_PORGAR)
   End If
   
   If Not IsNull(g_rst_Princi!MAECFI_NOCLIE) Then
      chk_NesCli.Value = g_rst_Princi!MAECFI_NOCLIE
   End If
   
   If moddat_g_int_FlgGrb_1 = 6 Then
      
      ipp_PlaCar.Text = 31 '45
      ipp_FecEmi.Text = Format(gf_FormatoFecha(CStr(g_rst_Princi!MAECFI_VTOFIA)), "dd/mm/yyyy") 'Format(CDate(DateAdd("D", 1, Format(gf_FormatoFecha(CStr(g_rst_Princi!MAECFI_VTOFIA)), "dd/mm/yyyy"))), "DD/MM/YYYY") '
      ipp_FecVct.Text = Format(CDate(DateAdd("D", ipp_PlaCar.Text, Me.ipp_FecEmi.Value)), "DD/MM/YYYY")
      Call fs_Calcular
      Call fs_Activa(False)
   Else
      ipp_PlaCar.Text = g_rst_Princi!MAECFI_PLZFIA
      txt_NumRef.Enabled = False
      lbl_TipRen.Visible = False
      cmb_TipRen.Visible = False
   End If
End Sub

Private Sub fs_Activa(ByVal p_Activa As Integer)
      cmb_Produc.Enabled = p_Activa
      cmb_SubPrd.Enabled = p_Activa
      cmb_Modali.Enabled = p_Activa
      ipp_FecEmi.Enabled = p_Activa
      ipp_PlaCar.Enabled = p_Activa
      ipp_FecVct.Enabled = p_Activa
      ipp_ImpCar.Enabled = p_Activa
      pnl_ValGar.Enabled = p_Activa
      ipp_TEACom.Enabled = p_Activa
      pnl_ValCom.Enabled = p_Activa
      pnl_ValMin.Enabled = p_Activa
      ipp_PorRet.Enabled = p_Activa
      txt_NumRef.Enabled = Not p_Activa
      txt_ParReg.Enabled = Not p_Activa
      txt_CodPry.Enabled = Not p_Activa
      txt_NumAde.Enabled = Not p_Activa
      txt_CodEte.Enabled = p_Activa
      cmb_NomPry.Enabled = p_Activa
      cmb_Moneda.Enabled = p_Activa
      cmb_TipRec.Enabled = p_Activa
      cmd_Grabar.Enabled = p_Activa
      'cmd_Cancel.Enabled = p_Activa
      cmb_TipRen.Visible = False
      lbl_TipRen.Visible = False
      cmb_PorGar.Enabled = p_Activa
      chk_NesCli.Visible = False
      cmb_TipLin.Visible = False
      lbl_TipLin.Visible = False
      cmd_NueTas.Enabled = False
      cmd_NueGar.Enabled = False
   
   If moddat_g_int_FlgGrb_1 = 6 Then
      txt_NumRef.Enabled = p_Activa
      cmb_TipRen.Visible = True
      lbl_TipRen.Visible = True
      txt_ParReg.Enabled = p_Activa
      txt_CodPry.Enabled = p_Activa
      txt_NumAde.Enabled = p_Activa
   End If
End Sub
Private Sub ipp_FecEmi_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
       Call gs_SetFocus(ipp_PlaCar)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub ipp_FecEmi_LostFocus()
   Call fs_Calcular
End Sub
Private Sub ipp_FecVct_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If moddat_g_int_FlgGrb_1 = 6 Then
         If moddat_g_int_TipRec = 1 Then
            Call gs_SetFocus(ipp_ImpCar)
         Else
            Call gs_SetFocus(cmd_Grabar) 'ipp_ImpCar
         End If
      Else
         Call gs_SetFocus(cmb_Moneda)
      End If
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub ipp_FecVct_LostFocus()
Dim r_lng_PlzCar  As Long
   
   If Format(ipp_FecVct.Text, "yyyymmdd") <> l_str_VctCal Then
      r_lng_PlzCar = DateDiff("D", ipp_FecEmi.Value, ipp_FecVct.Value)
      ipp_PlaCar.Text = r_lng_PlzCar
      Call gs_SetFocus(cmb_Moneda)
   End If
End Sub

Private Sub ipp_ImpCar_Change()
   pnl_ValGar.Caption = "0.00 "
   pnl_ValCom.Caption = "0.00 "
   pnl_ValMin.Caption = "0.00 "
End Sub

Private Sub ipp_ImpCar_GotFocus()
   Call fs_Calcular
End Sub

Private Sub ipp_ImpCar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If moddat_g_str_CodPrd = "008" Then
         Call gs_SetFocus(ipp_TasMor)
      Else
         Call gs_SetFocus(cmb_PorGar)
      End If
   End If
End Sub

Private Sub ipp_ImpCar_LostFocus()
   Call fs_Calcular
End Sub

Private Sub fs_Calcular()
Dim r_dbl_TEACom  As Double
Dim r_int_PlaCar   As Integer
Dim r_dbl_PorGar     As Double
Dim r_dbl_ValCom     As Double
   
   If ipp_ImpCar.Value < 0 Then
      pnl_ValGar.Caption = "0.00 "
      pnl_ValCom.Caption = "0.00 "
      pnl_ValMin.Caption = "0.00 "
   Else
      If moddat_g_int_TipRec = 2 Then
         r_dbl_TEACom = 0
      Else
         r_dbl_TEACom = CDbl(Replace(ipp_TEACom.Value, "%", "")) / 100
      End If
      r_int_PlaCar = ipp_PlaCar.Value
      
      If moddat_g_str_CodPrd <> "008" Then
         If moddat_g_str_CodMod = "008" Then
            pnl_ValGar.Caption = Format(CDbl(ipp_ImpCar), "##,###,##0.00") & "  "
         Else
            If cmb_PorGar.ListIndex <> -1 Then
               r_dbl_PorGar = Replace(cmb_PorGar.Text, "%", "")
               pnl_ValGar.Caption = Format(CDbl(ipp_ImpCar * (100 + r_dbl_PorGar) / 100), "##,###,##0.00") & "  "  'cmb_Porcen.ItemData(cmb_Porcen.ListIndex)
            End If
         End If
         
         r_dbl_ValCom = Format(CDbl(pnl_ValGar.Caption) * (CDbl(r_dbl_TEACom) / 360) * r_int_PlaCar, "##,###,##0.00") & "  " '12 (mes)
         
         If r_dbl_ValCom >= CInt(150) Then
            pnl_ValCom.Caption = Format(CDbl(r_dbl_ValCom), "##,###,##0.00") & "  "  'Format(CDbl(pnl_ValGar.Caption) * (CDbl(r_dbl_TEACom) / 360) * r_int_PlaCar, "##,###,##0.00") & "  " '12 (mes)
         Else
            If moddat_g_int_TipRec = 2 Then
               pnl_ValCom.Caption = Format(CDbl(r_dbl_ValCom), "##,###,##0.00") & "  "
               ipp_TEACom.Value = r_dbl_TEACom
            Else
               pnl_ValCom.Caption = Format(CDbl(150), "##,###,##0.00") & "  "
            End If
         End If
                  
         pnl_ValMin.Caption = Format(pnl_ValGar.Caption * (CDbl(r_dbl_TEACom) / 12 * 3), "##,###,##0.00") & "  "
      End If
      
   End If
   ipp_FecVct.Text = Format(CDate(DateAdd("D", r_int_PlaCar, Me.ipp_FecEmi.Value)), "DD/MM/YYYY")
   l_str_VctCal = Format(ipp_FecVct.Text, "yyyymmdd")
End Sub

Private Sub ipp_PlaCar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
       Call gs_SetFocus(ipp_FecVct)
   End If
End Sub

Private Sub ipp_PlaCar_LostFocus()
   Call fs_Calcular
   ipp_FecVct.Text = Format(CDate(DateAdd("D", ipp_PlaCar.Text, Me.ipp_FecEmi.Value)), "DD/MM/YYYY")
End Sub

Private Sub ipp_PorRet_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
       Call gs_SetFocus(txt_CodEte)
   End If
End Sub

Private Sub ipp_PorRet_LostFocus()
   Call fs_Calcular
End Sub

Private Sub ipp_PorTEA_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecEmi)
   End If
End Sub

Private Sub ipp_TasMor_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValCom)
   End If
End Sub

Private Sub ipp_TEACom_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If moddat_g_int_FlgGrb_1 = 6 Then
            Call gs_SetFocus(cmd_Grabar)
        Else
            If moddat_g_str_CodPrd = "026" And moddat_g_str_CodSub = "001" And moddat_g_str_CodMod = "005" Then
                Call gs_SetFocus(txt_NumAde)
            ElseIf moddat_g_str_CodPrd = "026" And moddat_g_str_CodSub = "001" And moddat_g_str_CodMod <> "008" Then
                Call gs_SetFocus(txt_CodPry)
            Else
                Call gs_SetFocus(cmd_Grabar)
            End If
        End If
    End If
End Sub

Private Sub ipp_TEACom_LostFocus()
   Call fs_Calcular
End Sub

Private Sub ipp_ValCom_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_NomPry)
   End If
End Sub

Private Sub txt_CodEte_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecEmi)
   Else
       KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "() ,.-_:;#@$=%&/+*\" & Chr(22))
   End If
End Sub

Private Sub txt_CodPry_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_ParReg)
   End If
End Sub

Private Sub txt_NumAde_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_CodPry)
   Else
       KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "() ,.-_:;#@$=%&/+*\" & Chr(22))
   End If
End Sub

Private Sub Txt_NumRef_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Call gs_SetFocus(ipp_PorRet)
    Else
       KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
    End If
End Sub

Private Sub cmb_Moneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ImpCar)
   End If
End Sub

Private Sub txt_ParReg_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If moddat_g_str_CodPrd = "026" And moddat_g_str_CodSub = "001" Then
            Call gs_SetFocus(cmb_NomPry)
        Else
            Call gs_SetFocus(cmd_Grabar)
        End If
    End If
End Sub
