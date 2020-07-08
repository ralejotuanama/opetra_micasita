VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_Con_PreSeg_03 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11370
   Icon            =   "OpeTra_frm_404.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   11370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   6525
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   11400
      _Version        =   65536
      _ExtentX        =   20108
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
      Begin Threed.SSPanel SSPanel9 
         Height          =   2070
         Left            =   60
         TabIndex        =   21
         Top             =   2595
         Width           =   11235
         _Version        =   65536
         _ExtentX        =   19817
         _ExtentY        =   3651
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
         Begin EditLib.fpDoubleSingle ipp_Capital 
            Height          =   315
            Left            =   1950
            TabIndex        =   0
            Top             =   300
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
         Begin EditLib.fpDoubleSingle ipp_Interes 
            Height          =   315
            Left            =   5520
            TabIndex        =   1
            Top             =   300
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
         Begin EditLib.fpDoubleSingle ipp_ComComp 
            Height          =   315
            Left            =   9120
            TabIndex        =   2
            Top             =   300
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
         Begin EditLib.fpDoubleSingle ipp_ComOtr 
            Height          =   315
            Left            =   1950
            TabIndex        =   3
            Top             =   630
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
         Begin EditLib.fpDoubleSingle ipp_MtoTelex 
            Height          =   315
            Left            =   5520
            TabIndex        =   4
            Top             =   630
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
         Begin EditLib.fpDoubleSingle ipp_MtoPort 
            Height          =   315
            Left            =   9120
            TabIndex        =   5
            Top             =   630
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
         Begin Threed.SSPanel pnl_TotPag 
            Height          =   315
            Left            =   9120
            TabIndex        =   19
            Top             =   1620
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
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
         Begin EditLib.fpDoubleSingle ipp_MtoMora 
            Height          =   315
            Left            =   1950
            TabIndex        =   6
            Top             =   960
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
         Begin EditLib.fpDoubleSingle ipp_GasDiv 
            Height          =   315
            Left            =   5520
            TabIndex        =   7
            Top             =   960
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
         Begin EditLib.fpDoubleSingle ipp_ComCof 
            Height          =   315
            Left            =   9120
            TabIndex        =   8
            Top             =   960
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
         Begin EditLib.fpDoubleSingle ipp_ComExt 
            Height          =   315
            Left            =   1950
            TabIndex        =   9
            Top             =   1290
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
         Begin EditLib.fpDoubleSingle ipp_ComRenov 
            Height          =   315
            Left            =   5520
            TabIndex        =   10
            Top             =   1290
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
         Begin EditLib.fpDoubleSingle ipp_DevPBP 
            Height          =   315
            Left            =   1950
            TabIndex        =   11
            Top             =   1620
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
         Begin EditLib.fpDoubleSingle ipp_IntLeg 
            Height          =   315
            Left            =   5520
            TabIndex        =   12
            Top             =   1620
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
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Devolución PBP:"
            Height          =   195
            Left            =   150
            TabIndex        =   60
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Interés Legal:"
            Height          =   195
            Left            =   3780
            TabIndex        =   59
            Top             =   1680
            Width           =   960
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Cobranza"
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
            Left            =   150
            TabIndex        =   50
            Top             =   60
            Width           =   810
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Total Cobranza:"
            Height          =   195
            Left            =   7380
            TabIndex        =   46
            Top             =   1680
            Width           =   1125
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Comisión Renovación:"
            Height          =   195
            Left            =   3780
            TabIndex        =   45
            Top             =   1350
            Width           =   1590
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Comisión Com. Exterior:"
            Height          =   195
            Left            =   150
            TabIndex        =   44
            Top             =   1350
            Width           =   1650
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Comisión COFIDE:"
            Height          =   195
            Left            =   7380
            TabIndex        =   43
            Top             =   1020
            Width           =   1305
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Gastos Diversos:"
            Height          =   195
            Left            =   3780
            TabIndex        =   42
            Top             =   1020
            Width           =   1200
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Mora:"
            Height          =   195
            Left            =   150
            TabIndex        =   41
            Top             =   1020
            Width           =   405
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Portes:"
            Height          =   195
            Left            =   7380
            TabIndex        =   27
            Top             =   690
            Width           =   495
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Capital:"
            Height          =   195
            Left            =   150
            TabIndex        =   26
            Top             =   360
            Width           =   525
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Interés:"
            Height          =   195
            Left            =   3780
            TabIndex        =   25
            Top             =   360
            Width           =   525
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Comisión Compromiso:"
            Height          =   195
            Left            =   7380
            TabIndex        =   24
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Comisión Otra:"
            Height          =   195
            Left            =   150
            TabIndex        =   23
            Top             =   690
            Width           =   1020
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Telex:"
            Height          =   195
            Left            =   3780
            TabIndex        =   22
            Top             =   690
            Width           =   435
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   645
         Left            =   60
         TabIndex        =   28
         Top             =   780
         Width           =   11235
         _Version        =   65536
         _ExtentX        =   19817
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
            Picture         =   "OpeTra_frm_404.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10620
            Picture         =   "OpeTra_frm_404.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   675
         Left            =   60
         TabIndex        =   29
         Top             =   60
         Width           =   11235
         _Version        =   65536
         _ExtentX        =   19817
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
            TabIndex        =   31
            Top             =   90
            Width           =   6615
            _Version        =   65536
            _ExtentX        =   11668
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Créditos Hipotecarios"
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
            Height          =   315
            Left            =   630
            TabIndex        =   32
            Top             =   330
            Width           =   6765
            _Version        =   65536
            _ExtentX        =   11933
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Recepción de Cronograma Pasivo"
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
            Picture         =   "OpeTra_frm_404.frx":0890
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   1080
         Left            =   60
         TabIndex        =   30
         Top             =   1470
         Width           =   11235
         _Version        =   65536
         _ExtentX        =   19817
         _ExtentY        =   1905
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
            Left            =   1320
            TabIndex        =   33
            Top             =   300
            Width           =   1485
            _Version        =   65536
            _ExtentX        =   2619
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
            Left            =   1320
            TabIndex        =   34
            Top             =   630
            Width           =   6915
            _Version        =   65536
            _ExtentX        =   12197
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
         Begin Threed.SSPanel pnl_FecPpg 
            Height          =   315
            Left            =   9690
            TabIndex        =   37
            Top             =   630
            Width           =   1305
            _Version        =   65536
            _ExtentX        =   2302
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
         Begin Threed.SSPanel pnl_ImpPpg 
            Height          =   315
            Left            =   9690
            TabIndex        =   39
            Top             =   300
            Width           =   1305
            _Version        =   65536
            _ExtentX        =   2302
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
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_TipPpg 
            Height          =   315
            Left            =   4080
            TabIndex        =   47
            Top             =   300
            Width           =   2265
            _Version        =   65536
            _ExtentX        =   3995
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
         Begin Threed.SSPanel pnl_Moneda 
            Height          =   315
            Left            =   7230
            TabIndex        =   57
            Top             =   300
            Width           =   1005
            _Version        =   65536
            _ExtentX        =   1773
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
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Moneda:"
            Height          =   195
            Left            =   6540
            TabIndex        =   58
            Top             =   360
            Width           =   630
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Datos"
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
            Left            =   150
            TabIndex        =   51
            Top             =   60
            Width           =   510
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Prepago:"
            Height          =   195
            Left            =   3030
            TabIndex        =   48
            Top             =   360
            Width           =   1005
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Importe PPG.:"
            Height          =   195
            Left            =   8520
            TabIndex        =   40
            Top             =   360
            Width           =   990
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Prepago:"
            Height          =   195
            Left            =   8520
            TabIndex        =   38
            Top             =   690
            Width           =   1140
         End
         Begin VB.Label Label51 
            AutoSize        =   -1  'True
            Caption         =   "Nro. Operación:"
            Height          =   195
            Left            =   150
            TabIndex        =   36
            Top             =   360
            Width           =   1125
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            Caption         =   "Cliente:"
            Height          =   195
            Left            =   150
            TabIndex        =   35
            Top             =   690
            Width           =   525
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   1755
         Left            =   60
         TabIndex        =   49
         Top             =   4710
         Width           =   11235
         _Version        =   65536
         _ExtentX        =   19817
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
         Begin VB.ComboBox cmb_Proveedor 
            Height          =   315
            Left            =   1950
            TabIndex        =   14
            Top             =   630
            Width           =   7140
         End
         Begin VB.ComboBox cmb_Banco 
            Height          =   315
            Left            =   1950
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   960
            Width           =   4400
         End
         Begin VB.ComboBox cmb_CtaCte 
            Height          =   315
            Left            =   1950
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   1290
            Width           =   4400
         End
         Begin VB.ComboBox cmb_TipDoc 
            Height          =   315
            Left            =   1950
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   300
            Width           =   4400
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "Proveedor:"
            Height          =   195
            Left            =   150
            TabIndex        =   56
            Top             =   690
            Width           =   780
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            Caption         =   "Banco:"
            Height          =   195
            Left            =   150
            TabIndex        =   55
            Top             =   1020
            Width           =   510
         End
         Begin VB.Label lbl_Cuenta 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta:"
            Height          =   195
            Left            =   150
            TabIndex        =   54
            Top             =   1350
            Width           =   555
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Documento:"
            Height          =   195
            Left            =   150
            TabIndex        =   53
            Top             =   360
            Width           =   1230
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Proveedor"
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
            Left            =   150
            TabIndex        =   52
            Top             =   60
            Width           =   885
         End
      End
   End
End
Attribute VB_Name = "frm_Con_PreSeg_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_arr_CtaCteSol()   As moddat_tpo_Genera
Dim l_arr_CtaCteDol()   As moddat_tpo_Genera
Dim l_arr_MaePrv()      As moddat_tpo_Genera
Dim l_int_CodMon        As Integer
Dim l_int_Contar        As Integer

Private Sub cmb_Banco_Click()
Dim r_str_Cadena  As String
   
   cmb_CtaCte.Clear
   r_str_Cadena = ""
   lbl_Cuenta.Caption = "Cuenta:"
   
   'If (cmb_Moneda.ListIndex = -1) Then
   '    Exit Sub
   'End If
   
   If l_int_CodMon = 1 Then
       For l_int_Contar = 1 To UBound(l_arr_CtaCteSol)
           r_str_Cadena = ""
           If (Trim(cmb_Banco.ItemData(cmb_Banco.ListIndex)) = Trim(l_arr_CtaCteSol(l_int_Contar).Genera_Codigo)) Then
               If (Trim(cmb_Banco.ItemData(cmb_Banco.ListIndex)) = 11) Then 'Banco continental
                   r_str_Cadena = Trim(l_arr_CtaCteSol(l_int_Contar).Genera_Prefij)
                   lbl_Cuenta.Caption = "Cuenta Corriente:"
               Else
                   r_str_Cadena = Trim(l_arr_CtaCteSol(l_int_Contar).Genera_Refere)
                   lbl_Cuenta.Caption = "CCI:"
               End If
           End If
           If (Len(Trim(r_str_Cadena)) > 0) Then
               cmb_CtaCte.AddItem Trim(r_str_Cadena)
           End If
       Next
   End If
   
   If l_int_CodMon = 2 Then
       For l_int_Contar = 1 To UBound(l_arr_CtaCteDol)
           r_str_Cadena = ""
           If (Trim(cmb_Banco.ItemData(cmb_Banco.ListIndex)) = Trim(l_arr_CtaCteDol(l_int_Contar).Genera_Codigo)) Then
               If (Trim(cmb_Banco.ItemData(cmb_Banco.ListIndex)) = 11) Then 'Banco continental
                   r_str_Cadena = Trim(l_arr_CtaCteDol(l_int_Contar).Genera_Prefij)
                   lbl_Cuenta.Caption = "Cuenta Corriente:"
               Else
                   r_str_Cadena = Trim(l_arr_CtaCteDol(l_int_Contar).Genera_Refere)
                   lbl_Cuenta.Caption = "CCI:"
               End If
           End If
           If (Len(Trim(r_str_Cadena)) > 0) Then
               cmb_CtaCte.AddItem Trim(r_str_Cadena)
           End If
       Next
   End If
End Sub

Private Sub cmb_Proveedor_Click()
   Call fs_Buscar_prov
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Limpia
   
   'consultar
   If moddat_g_int_FlgGrb = 0 Then
      Call fs_Cargar
   End If
   
   cmb_TipDoc.Enabled = False
   cmb_Proveedor.Enabled = False
   cmb_Banco.Enabled = False
   cmb_CtaCte.Enabled = False
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
Dim r_str_CadAux As String
Dim r_int_Contad As Integer

   Call moddat_gs_FecSis
   
   pnl_NumOpe.Caption = ""
   pnl_NumOpe.Tag = ""
   pnl_NomCli.Caption = ""
   pnl_FecPpg.Caption = ""
   pnl_TipPpg.Caption = ""
   cmb_TipDoc.ListIndex = -1
   cmb_Proveedor.Text = ""
   cmb_Banco.Clear
   cmb_CtaCte.Clear
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipDoc, 1, "118")
   
   'estado inssertar
   If moddat_g_int_FlgGrb = 1 Or (moddat_g_int_FlgGrb = 0 And Trim(moddat_g_str_Codigo) = "") Then
      r_int_Contad = frm_Con_PreSeg_01.grd_Listad.Row
      pnl_NumOpe.Caption = frm_Con_PreSeg_01.grd_Listad.TextMatrix(r_int_Contad, 0)    'FORMATEADA
      pnl_NumOpe.Tag = frm_Con_PreSeg_01.grd_Listad.TextMatrix(r_int_Contad, 9)        'SIN FORMATEADA
      pnl_NomCli.Caption = frm_Con_PreSeg_01.grd_Listad.TextMatrix(r_int_Contad, 1) & " / " & frm_Con_PreSeg_01.grd_Listad.TextMatrix(r_int_Contad, 2)
      pnl_FecPpg.Caption = frm_Con_PreSeg_01.grd_Listad.TextMatrix(r_int_Contad, 4)
      pnl_ImpPpg.Caption = Format(frm_Con_PreSeg_01.grd_Listad.TextMatrix(r_int_Contad, 17), "###,###,###,##0.00") & " "
      pnl_TipPpg.Caption = Trim(frm_Con_PreSeg_01.grd_Listad.TextMatrix(r_int_Contad, 3))
      l_int_CodMon = frm_Con_PreSeg_01.grd_Listad.TextMatrix(r_int_Contad, 15)
      If l_int_CodMon = 1 Then
         pnl_Moneda.Caption = "SOLES"
      Else
         pnl_Moneda.Caption = "DOLARES"
      End If
            
      Call gs_BuscarCombo_Item(cmb_TipDoc, 6)
      cmb_Proveedor.ListIndex = fs_ComboIndex(cmb_Proveedor, "20100116392", 0)
      If cmb_Banco.ListCount > 0 Then
         cmb_Banco.ListIndex = 0
      End If
      If cmb_CtaCte.ListCount > 0 Then
      cmb_CtaCte.ListIndex = 0
      End If
   End If
   
   modctb_str_FecIni = ""
   modctb_str_FecFin = ""
   modctb_int_PerAno = 0
   modctb_int_PerMes = 0
   r_str_CadAux = ""
   
   Call moddat_gf_ConsultaPerMesActivo("000001", 1, modctb_str_FecIni, modctb_str_FecFin, modctb_int_PerMes, modctb_int_PerAno)
   r_str_CadAux = DateAdd("m", 1, "01/" & Format(modctb_int_PerMes, "00") & "/" & modctb_int_PerAno)
   modctb_str_FecFin = DateAdd("d", -1, r_str_CadAux)
   modctb_str_FecIni = DateAdd("m", -1, modctb_str_FecFin)
   modctb_str_FecIni = "01/" & Format(Month(modctb_str_FecIni), "00") & "/" & Year(modctb_str_FecIni)
End Sub

Private Sub fs_Limpia()
   ipp_Capital.Text = "0.00"
   ipp_Interes.Text = "0.00"
   ipp_ComComp.Text = "0.00"
   ipp_ComOtr.Text = "0.00"
   ipp_MtoTelex.Text = "0.00"
   ipp_MtoPort.Text = "0.00"
   ipp_MtoMora.Text = "0.00"
   ipp_GasDiv.Text = "0.00"
   ipp_ComCof.Text = "0.00"
   ipp_ComExt.Text = "0.00"
   ipp_ComRenov.Text = "0.00"
   ipp_DevPBP.Text = "0.00"
   ipp_IntLeg.Text = "0.00"
   pnl_TotPag.Caption = "0.00" & " "
   
   Call gs_SetFocus(ipp_Capital)
End Sub

Private Function fs_ComboIndex(p_Combo As ComboBox, cadena As String, p_Tipo As Integer) As Integer
Dim r_int_Contad As Integer

   fs_ComboIndex = -1
   For r_int_Contad = 0 To p_Combo.ListCount - 1
       If Trim(cadena) = Trim(Mid(p_Combo.List(r_int_Contad), 1, InStr(Trim(p_Combo.List(r_int_Contad)), "-") - 1)) Then
          fs_ComboIndex = r_int_Contad
          Exit For
       End If
   Next
End Function

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub cmd_Grabar_Click()
Dim r_int_Contad  As Integer
Dim r_dbl_TipCam  As Double
Dim r_lng_FecPpg  As Long
Dim r_bol_Estado  As Boolean

   If CDbl(pnl_TotPag.Caption) <= 0 Then
      MsgBox "La Suma total tiene que ser mayor a cero.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_Capital)
      Exit Sub
   End If
   If Trim(pnl_NumOpe.Caption) = "" Or Trim(pnl_NumOpe.Tag) = "" Then
      MsgBox "Tiene que estar definido un número de operación.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_Capital)
      Exit Sub
   End If
   If Trim(pnl_FecPpg.Caption) = "" Then
      MsgBox "Tiene que estar definida la fecha de prepago", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_Capital)
      Exit Sub
   End If
   If cmb_TipDoc.ListIndex = -1 Then
      MsgBox "Tiene que seleccionar un tipo de documento.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipDoc)
      Exit Sub
   End If
   If Len(Trim(cmb_Proveedor.Text)) = 0 Then
       MsgBox "Tiene que ingresar un proveedor.", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(cmb_Proveedor)
       Exit Sub
   Else
       If (fs_ValNumDoc() = False) Then
           Exit Sub
       Else
           r_bol_Estado = False
           If InStr(1, Trim(cmb_Proveedor.Text), "-") > 0 Then
              For l_int_Contar = 1 To UBound(l_arr_MaePrv)
                  If Trim(Mid(cmb_Proveedor.Text, 1, InStr(Trim(cmb_Proveedor.Text), "-") - 1)) = Trim(l_arr_MaePrv(l_int_Contar).Genera_Codigo) Then
                     r_bol_Estado = True
                     Exit For
                  End If
              Next
           End If
           If r_bol_Estado = False Then
              MsgBox "El Proveedor no se encuentra en la lista.", vbExclamation, modgen_g_str_NomPlt
              Call gs_SetFocus(cmb_Proveedor)
              Exit Sub
           End If
       End If
   End If
   If cmb_Banco.ListIndex = -1 Then
      MsgBox "Tiene que seleccionar un banco.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Banco)
      Exit Sub
   End If
   If cmb_CtaCte.ListIndex = -1 Then
      MsgBox "Tiene que seleccionar un nro cuenta.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_CtaCte)
      Exit Sub
   End If
   
   r_dbl_TipCam = moddat_gf_ObtieneTipCamDia(2, 2, Format(moddat_g_str_FecSis, "yyyymmdd"), 1)
   If r_dbl_TipCam = 0 Then
      MsgBox "Tiene que ingresar el tipo de cambio SBS del día.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_Capital)
      Exit Sub
   End If
   
   'Valida que las solicitudes seleccionadas solo sean --> Debe tener estado de Enviado a COFIDE (Est = 2) para que pase a Abono a COFIDE
   r_int_Contad = frm_Con_PreSeg_01.grd_Listad.Row 'Fila seleccionada
   Screen.MousePointer = 11
   If Not frm_Con_PreSeg_01.fs_Valida_EstPrepago(frm_Con_PreSeg_01.grd_Listad.TextMatrix(r_int_Contad, 13), 3) Then
      If Trim(frm_Con_PreSeg_01.grd_Listad.TextMatrix(r_int_Contad, 13)) <> "RECEPCION DE CALENDARIO" Then
         MsgBox "La solicitud " & frm_Con_PreSeg_01.grd_Listad.TextMatrix(r_int_Contad, 0) & " no se puede Recepcionar, porque no se encuentra en instancia ABONO A COFIDE.", vbInformation, modgen_g_str_NomPlt
      Else
         MsgBox "La solicitud " & frm_Con_PreSeg_01.grd_Listad.TextMatrix(r_int_Contad, 0) & " no se puede Recepcionar, porque ya se encuentra en instancia RECEPCION DE CALENDARIO.", vbInformation, modgen_g_str_NomPlt
      End If
      Screen.MousePointer = 0
      Exit Sub
   End If
   Screen.MousePointer = 0
   
   r_lng_FecPpg = Format(moddat_g_str_FecSis, "yyyymmdd")
   If (r_lng_FecPpg < Format(modctb_str_FecIni, "yyyymmdd") Or r_lng_FecPpg > Format(modctb_str_FecFin, "yyyymmdd")) Then
       MsgBox "Intenta registrar un documento en un periodo errado.", vbExclamation, modgen_g_str_NomPlt
       Call gs_SetFocus(ipp_Capital)
       Exit Sub
   End If
      
   If MsgBox("¿Esta seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_Grabar
   Screen.MousePointer = 0
End Sub

Private Sub fs_Grabar()
Dim r_str_Cadena   As String
Dim r_int_AsiGen   As Integer
Dim r_rst_Genera   As ADODB.Recordset
Dim r_int_Contad   As Integer
Dim r_lng_FecAct   As String
Dim r_str_TipCli   As String
Dim r_str_CodGen As String

   r_str_CodGen = ""
   r_str_CodGen = modmip_gf_Genera_CodGen(3, 10)

   If Len(Trim(r_str_CodGen)) = 0 Then
      MsgBox "No se genero el código automatico del folio.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If

   r_str_Cadena = ""
   r_str_Cadena = r_str_Cadena & "USP_CRE_PPGPASCTB("
   r_str_Cadena = r_str_Cadena & "'" & pnl_NumOpe.Tag & "', "
   r_str_Cadena = r_str_Cadena & Format(pnl_FecPpg.Caption, "yyyymmdd") & ", "
   r_str_Cadena = r_str_Cadena & r_str_CodGen & ", "
   If Trim(pnl_TipPpg.Caption) = "TOTAL" Then
      r_str_Cadena = r_str_Cadena & "5, "
   Else
      r_str_Cadena = r_str_Cadena & "4, "
   End If
   r_str_Cadena = r_str_Cadena & CDbl(ipp_Capital.Text) & ", "
   r_str_Cadena = r_str_Cadena & CDbl(ipp_Interes.Text) & ", "
   r_str_Cadena = r_str_Cadena & CDbl(ipp_ComComp.Text) & ", "
   r_str_Cadena = r_str_Cadena & CDbl(ipp_ComOtr.Text) & ", "
   r_str_Cadena = r_str_Cadena & CDbl(ipp_MtoTelex.Text) & ", "
   r_str_Cadena = r_str_Cadena & CDbl(ipp_MtoPort.Text) & ", "
   r_str_Cadena = r_str_Cadena & CDbl(ipp_MtoMora.Text) & ", "
   r_str_Cadena = r_str_Cadena & CDbl(ipp_GasDiv.Text) & ", "
   r_str_Cadena = r_str_Cadena & CDbl(ipp_ComCof.Text) & ", "
   r_str_Cadena = r_str_Cadena & CDbl(ipp_ComExt.Text) & ", "
   r_str_Cadena = r_str_Cadena & CDbl(ipp_ComRenov.Text) & ", "
   r_str_Cadena = r_str_Cadena & CDbl(ipp_DevPBP.Text) & ", "
   r_str_Cadena = r_str_Cadena & CDbl(ipp_IntLeg.Text) & ", "
   r_str_Cadena = r_str_Cadena & cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) & ", " 'CTAPAG_TIPDOC
   r_str_Cadena = r_str_Cadena & "'" & fs_NumDoc(cmb_Proveedor.Text) & "', "      'CTAPAG_NUMDOC
   r_str_Cadena = r_str_Cadena & cmb_Banco.ItemData(cmb_Banco.ListIndex) & ", "   'CTAPAG_CODBCO
   r_str_Cadena = r_str_Cadena & "'" & Trim(cmb_CtaCte.Text) & "', "              'CTAPAG_CTACRR
   'Datos de Auditoria
   r_str_Cadena = r_str_Cadena & "'" & modgen_g_str_CodUsu & "', "  'Código Usuario
   r_str_Cadena = r_str_Cadena & "'" & modgen_g_str_NombPC & "', "  'Nombre Terminal
   r_str_Cadena = r_str_Cadena & "'" & UCase(App.EXEName) & "', "   'Nombre Ejecutable
   r_str_Cadena = r_str_Cadena & "'" & modgen_g_str_CodSuc & "') "  'Código Sucursal
   'r_str_Cadena = r_str_Cadena & "1) "
  
   If Not gf_EjecutaSQL(r_str_Cadena, r_rst_Genera, 3) Then
      r_rst_Genera.Close
      Set r_rst_Genera = Nothing
      MsgBox "No se pudo insertar el registro.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If (r_rst_Genera!RESUL = 1) Then
       'generar asiento
       Call fs_GeneraAsiento(pnl_NumOpe.Tag, r_str_CodGen, r_int_AsiGen, l_int_CodMon)
       Call frm_Con_PreSeg_01.fs_Buscar
       
       MsgBox "Los datos se grabaron correctamente." & vbCrLf & "El asiento generado es: " & CStr(r_int_AsiGen), vbInformation, modgen_g_str_NomPlt
       Unload Me
   End If
   
   r_rst_Genera.Close
   Set r_rst_Genera = Nothing
End Sub

Private Sub fs_RecepCronog()
Dim r_int_Contad        As Integer
Dim r_lng_FecAct        As String
Dim r_str_TipCli        As String
      
   'Actualiza el estado del prepago
   r_lng_FecAct = Format(pnl_FecPpg.Caption, "yyyymmdd")

   g_str_Parame = "USP_ACTUALIZA_CRE_PPGCAB ("
   g_str_Parame = g_str_Parame & "'" & Trim(pnl_NumOpe.Tag) & "', "
   
   If Trim(pnl_TipPpg.Caption) = "TOTAL" Then
      g_str_Parame = g_str_Parame & "" & r_lng_FecAct & " , 5, 0, 0, 0, "
   Else
      g_str_Parame = g_str_Parame & "" & r_lng_FecAct & " , 4, 0, 0, 0, "
   End If
   g_str_Parame = g_str_Parame & Format(CDate(Now), "yyyymmdd") & ") "
 
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
      MsgBox "No se pudo completar la actualización del estado de los datos.", vbInformation, modgen_g_con_PltPar
      Exit Sub
   End If
End Sub

Private Sub fs_Cargar()
Dim r_str_Parame    As String
Dim r_rst_Genera    As ADODB.Recordset
Dim r_dbl_Import    As Double
         
   If moddat_g_int_FlgGrb = 0 And Trim(moddat_g_str_Codigo) = "" Then
      r_str_Parame = ""
      r_str_Parame = r_str_Parame & " SELECT PPGPAS_MTOCAP, PPGPAS_MTOINT, PPGPAS_COMCMP, PPGPAS_COMOTR, "
      r_str_Parame = r_str_Parame & "        PPGPAS_MTOTEL, PPGPAS_MTOPOR, PPGPAS_MTOMOR, PPGPAS_GASDIV, "
      r_str_Parame = r_str_Parame & "        PPGPAS_COMCOF, PPGPAS_COMEXT, PPGPAS_COMRNV, "
      r_str_Parame = r_str_Parame & "        PPGPAS_TIPDOC, PPGPAS_NUMDOC, PPGPAS_CODBNC, PPGPAS_NUMCTA "
      r_str_Parame = r_str_Parame & "   FROM CRE_PPGPASCTB A "
      r_str_Parame = r_str_Parame & "  WHERE A.PPGPAS_NUMOPE = '" & Trim(pnl_NumOpe.Tag) & "'"
      r_str_Parame = r_str_Parame & "    AND A.PPGPAS_FECPPG = " & Format(pnl_FecPpg.Caption, "yyyymmdd")
      r_str_Parame = r_str_Parame & "    AND A.PPGPAS_SITUAC = 1 "
   Else
      r_str_Parame = ""
      r_str_Parame = r_str_Parame & " SELECT PPGPAS_CODREG, A.PPGPAS_NUMOPE, A.PPGPAS_FECPPG, B.HIPMAE_MONEDA, C.PPGCAB_TIPPPG,  "
      r_str_Parame = r_str_Parame & "        TRIM(D.DATGEN_APEPAT)||' '||TRIM(D.DATGEN_APEMAT)||' '||TRIM(D.DATGEN_NOMBRE) AS CLIENTE,  "
      r_str_Parame = r_str_Parame & "        B.HIPMAE_TDOCLI, B.HIPMAE_NDOCLI, PPGPAS_MTOCAP, PPGPAS_MTOINT, PPGPAS_COMCMP, PPGPAS_COMOTR,  "
      r_str_Parame = r_str_Parame & "        PPGPAS_MTOTEL, PPGPAS_MTOPOR, PPGPAS_MTOMOR, PPGPAS_GASDIV, PPGPAS_COMCOF, "
      r_str_Parame = r_str_Parame & "        PPGPAS_COMEXT, PPGPAS_COMRNV, PPGPAS_TIPDOC, PPGPAS_NUMDOC, PPGPAS_CODBNC, PPGPAS_NUMCTA "
      r_str_Parame = r_str_Parame & "   FROM CRE_PPGPASCTB A  "
      r_str_Parame = r_str_Parame & "  INNER JOIN CRE_HIPMAE B ON B.HIPMAE_NUMOPE = A.PPGPAS_NUMOPE  "
      r_str_Parame = r_str_Parame & "  INNER JOIN CRE_PPGCAB C ON C.PPGCAB_NUMOPE = A.PPGPAS_NUMOPE AND C.PPGCAB_FECPPG = A.PPGPAS_FECPPG  "
      r_str_Parame = r_str_Parame & "  INNER JOIN CLI_DATGEN D ON D.DATGEN_TIPDOC = B.HIPMAE_TDOCLI AND D.DATGEN_NUMDOC = B.HIPMAE_NDOCLI  "
      r_str_Parame = r_str_Parame & "  WHERE A.PPGPAS_CODREG = " & CLng(moddat_g_str_Codigo)
   End If
   
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Genera, 3) Then
      r_rst_Genera.Close
      Set r_rst_Genera = Nothing
      Exit Sub
   End If
   
   If r_rst_Genera.BOF And r_rst_Genera.EOF Then
      r_rst_Genera.Close
      Set r_rst_Genera = Nothing
      Exit Sub
   End If
      
   If moddat_g_int_FlgGrb = 0 And Trim(moddat_g_str_Codigo) <> "" Then
      pnl_NumOpe.Caption = gf_Formato_NumOpe(r_rst_Genera!PPGPAS_NUMOPE)  'FORMATEADO
      pnl_NumOpe.Tag = r_rst_Genera!PPGPAS_NUMOPE                         'SIN FORMATO
      pnl_NomCli.Caption = r_rst_Genera!HIPMAE_TDOCLI & "-" & r_rst_Genera!HIPMAE_NDOCLI & " / " & r_rst_Genera!CLIENTE
      pnl_FecPpg.Caption = r_rst_Genera!PPGPAS_FECPPG
      r_dbl_Import = r_rst_Genera!PPGPAS_MTOCAP + r_rst_Genera!PPGPAS_MTOINT + r_rst_Genera!PPGPAS_COMCMP + r_rst_Genera!PPGPAS_COMOTR + r_rst_Genera!PPGPAS_MTOTEL + _
                     r_rst_Genera!PPGPAS_MTOPOR + r_rst_Genera!PPGPAS_MTOMOR + r_rst_Genera!PPGPAS_GASDIV + r_rst_Genera!PPGPAS_COMCOF + r_rst_Genera!PPGPAS_COMEXT + _
                     r_rst_Genera!PPGPAS_COMRNV + r_rst_Genera!PPGPAS_DEVPBP + r_rst_Genera!PPGPAS_INTLEG
      pnl_ImpPpg.Caption = Format(r_dbl_Import, "###,###,###,##0.00") & " "
            
      l_int_CodMon = r_rst_Genera!HIPMAE_MONEDA
      If r_rst_Genera!PPGCAB_TIPPPG = 1 Then
         If r_rst_Genera!PPGCAB_TIPPPGPAR = 1 Then
            pnl_TipPpg.Caption = "PARCIAL - RED MONTO"
         Else
            pnl_TipPpg.Caption = "PARCIAL - RED PLAZO"
         End If
      Else
         pnl_TipPpg.Caption = "TOTAL"
      End If
      If r_rst_Genera!HIPMAE_MONEDA = 1 Then
         pnl_Moneda.Caption = "SOLES"
      Else
         pnl_Moneda.Caption = "DOLARES"
      End If
   End If
   
   cmd_Grabar.Visible = False
   
   ipp_Capital.Text = r_rst_Genera!PPGPAS_MTOCAP
   ipp_Interes.Text = r_rst_Genera!PPGPAS_MTOINT
   ipp_ComComp.Text = r_rst_Genera!PPGPAS_COMCMP
   ipp_ComOtr.Text = r_rst_Genera!PPGPAS_COMOTR
   ipp_MtoTelex.Text = r_rst_Genera!PPGPAS_MTOTEL
   ipp_MtoPort.Text = r_rst_Genera!PPGPAS_MTOPOR
   ipp_MtoMora.Text = r_rst_Genera!PPGPAS_MTOMOR
   ipp_GasDiv.Text = r_rst_Genera!PPGPAS_GASDIV
   ipp_ComCof.Text = r_rst_Genera!PPGPAS_COMCOF
   ipp_ComExt.Text = r_rst_Genera!PPGPAS_COMEXT
   ipp_ComRenov.Text = r_rst_Genera!PPGPAS_COMRNV
   ipp_DevPBP.Text = r_rst_Genera!PPGPAS_DEVPBP
   ipp_IntLeg.Text = r_rst_Genera!PPGPAS_INTLEG
   Call gs_BuscarCombo_Item(cmb_TipDoc, r_rst_Genera!PPGPAS_TIPDOC)
   cmb_Proveedor.ListIndex = fs_ComboIndex(cmb_Proveedor, r_rst_Genera!PPGPAS_NUMDOC & "", 0)
   Call gs_BuscarCombo_Item(cmb_Banco, r_rst_Genera!PPGPAS_CODBNC)
   Call gs_BuscarCombo_Text(cmb_CtaCte, r_rst_Genera!PPGPAS_NUMCTA, -1)
   
   Call ipp_ComRenov_LostFocus
   
   ipp_Capital.Enabled = False
   ipp_Interes.Enabled = False
   ipp_ComComp.Enabled = False
   ipp_ComOtr.Enabled = False
   ipp_MtoTelex.Enabled = False
   ipp_MtoPort.Enabled = False
   ipp_MtoMora.Enabled = False
   ipp_GasDiv.Enabled = False
   ipp_ComCof.Enabled = False
   ipp_ComExt.Enabled = False
   ipp_ComRenov.Enabled = False
   ipp_DevPBP.Enabled = False
   ipp_IntLeg.Enabled = False
   
   r_rst_Genera.Close
   Set r_rst_Genera = Nothing
End Sub

Private Sub fs_GeneraAsiento(ByVal p_NumOpe As String, ByVal p_Codigo As Long, ByRef p_AsiGen As Integer, p_CodMon As Integer)
Dim r_str_Cadena    As String
Dim r_str_CodPrd    As String
Dim r_str_CtaHab    As String
Dim r_str_CtaDeb    As String
Dim r_rst_Genera    As ADODB.Recordset
Dim r_rst_GenAux    As ADODB.Recordset
Dim r_arr_Matriz()  As modprc_g_tpo_Matriz
Dim r_arr_LogPro()  As modprc_g_tpo_LogPro
Dim r_int_NumIte    As Integer
Dim r_str_Origen    As String
Dim r_str_TipNot    As String
Dim r_int_NumLib    As String
Dim r_str_AsiGen    As String
Dim r_int_NumAsi    As Integer
Dim r_int_AuxCon    As Integer
Dim r_dbl_ImpSol    As Double
Dim r_dbl_ImpDol    As Double
Dim r_dbl_Import    As Double
Dim r_dbl_TipCam    As Double
Dim r_int_PerAno    As Integer
Dim r_int_PerMes    As Integer
Dim r_str_FechaL    As String
Dim r_str_CadAux    As String

   ReDim r_arr_LogPro(0)
   ReDim r_arr_LogPro(1)
   r_arr_LogPro(1).LogPro_CodPro = "CTBP1090"
   r_arr_LogPro(1).LogPro_FInEje = Format(date, "yyyymmdd")
   r_arr_LogPro(1).LogPro_HInEje = Format(Time, "hhmmss")
   r_arr_LogPro(1).LogPro_NumErr = 0
   
   r_str_Origen = "LM"
   r_str_TipNot = "O"
   r_int_NumLib = 6
   r_str_AsiGen = ""
   p_AsiGen = 0
   
   r_int_PerAno = Format(moddat_g_str_FecSis, "yyyy")  'Format(pnl_FecPpg.Caption, "yyyy")
   r_int_PerMes = Format(moddat_g_str_FecSis, "mm")    'Format(pnl_FecPpg.Caption, "mm")
   r_str_FechaL = moddat_g_str_FecSis                  'pnl_FecPpg.Caption

   r_str_CodPrd = Mid(Trim(p_NumOpe), 1, 3)
   r_dbl_TipCam = moddat_gf_ObtieneTipCamDia(2, 2, Format(r_str_FechaL, "yyyymmdd"), 1) 'Format(moddat_g_str_FecSis, "yyyymmdd"), 1)
   
   'query todos los productos agrupados
   Call moddat_gf_Cargar_AgrPrd
   ReDim r_arr_Matriz(0)
   
   If InStr(moddat_g_str_AgrCME, r_str_CodPrd) Or InStr(moddat_g_str_AgrMIHG, r_str_CodPrd) Or InStr(moddat_g_str_AgrTFMV, r_str_CodPrd) Then
      If (CDbl(ipp_Capital.Text) + CDbl(ipp_Interes.Text) + CDbl(ipp_ComComp.Text) + CDbl(ipp_ComOtr.Text) + CDbl(ipp_MtoTelex.Text) + CDbl(ipp_MtoPort.Text) + CDbl(ipp_MtoMora.Text) + _
          CDbl(ipp_GasDiv.Text) + CDbl(ipp_ComCof.Text) + CDbl(ipp_ComExt.Text) + CDbl(ipp_ComRenov.Text) + CDbl(ipp_DevPBP.Text) + CDbl(ipp_IntLeg.Text)) > 0 Then
          
          If CDbl(ipp_Capital.Text) > 0 Then
             ReDim Preserve r_arr_Matriz(UBound(r_arr_Matriz) + 1)
             r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = CDbl(ipp_Capital.Text)
             r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_DesNot = "PPG - " & p_NumOpe & " - CAP. PASIVO"
             r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_FlagDH = "D"
             
             If InStr(moddat_g_str_AgrCME, r_str_CodPrd) Then
                r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_CtaCtb = IIf(p_CodMon = 1, "261202010102", "262202010102")
             ElseIf InStr(moddat_g_str_AgrMIHG, r_str_CodPrd) Then
                r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_CtaCtb = IIf(p_CodMon = 1, "261202010101", "262202010101")
             ElseIf InStr(moddat_g_str_AgrTFMV, r_str_CodPrd) Then
                r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_CtaCtb = IIf(p_CodMon = 1, "261202010103", "262202010103")
             End If
          End If
             
          If CDbl(ipp_Interes.Text) > 0 Then
             ReDim Preserve r_arr_Matriz(UBound(r_arr_Matriz) + 1)
             r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = CDbl(ipp_Interes.Text)
             r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_DesNot = "PPG - " & p_NumOpe & " - INT. PASIVO"
             r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_FlagDH = "D"
             
             If InStr(moddat_g_str_AgrCME, r_str_CodPrd) Then
                r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_CtaCtb = IIf(p_CodMon = 1, "411402020102", "412402020102")
             ElseIf InStr(moddat_g_str_AgrMIHG, r_str_CodPrd) Then
                r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_CtaCtb = IIf(p_CodMon = 1, "411402020101", "412402020101")
             ElseIf InStr(moddat_g_str_AgrTFMV, r_str_CodPrd) Then
                r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_CtaCtb = IIf(p_CodMon = 1, "411402020103", "412402020103")
             End If
          End If
          If CDbl(ipp_ComComp.Text) > 0 Then
             ReDim Preserve r_arr_Matriz(UBound(r_arr_Matriz) + 1)
             r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = CDbl(ipp_ComComp.Text)
             r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_DesNot = "PPG - " & p_NumOpe & " - COM.COMP. PASIVO"
             r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_FlagDH = "D"
             
             If InStr(moddat_g_str_AgrCME, r_str_CodPrd) Then
                r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_CtaCtb = IIf(p_CodMon = 1, "411703020104", "412703020104")
             ElseIf InStr(moddat_g_str_AgrMIHG, r_str_CodPrd) Then
                r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_CtaCtb = IIf(p_CodMon = 1, "411703020102", "412703020102")
             ElseIf InStr(moddat_g_str_AgrTFMV, r_str_CodPrd) Then
                r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_CtaCtb = IIf(p_CodMon = 1, "411703020106", "412703020106")
             End If
          End If
          If CDbl(ipp_ComOtr.Text) > 0 Then
             ReDim Preserve r_arr_Matriz(UBound(r_arr_Matriz) + 1)
             r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = CDbl(ipp_ComOtr.Text)
             r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_DesNot = "PPG - " & p_NumOpe & " - COM.OTRAS. PASIVO"
             r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_FlagDH = "D"
             
             If InStr(moddat_g_str_AgrCME, r_str_CodPrd) Then
                r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_CtaCtb = IIf(p_CodMon = 1, "411703020104", "412703020104")
             ElseIf InStr(moddat_g_str_AgrMIHG, r_str_CodPrd) Then
                r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_CtaCtb = IIf(p_CodMon = 1, "411703020102", "412703020102")
             ElseIf InStr(moddat_g_str_AgrTFMV, r_str_CodPrd) Then
                r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_CtaCtb = IIf(p_CodMon = 1, "411703020106", "412703020106")
             End If
          End If
          If CDbl(ipp_MtoTelex.Text) > 0 Then
             ReDim Preserve r_arr_Matriz(UBound(r_arr_Matriz) + 1)
             r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = CDbl(ipp_MtoTelex.Text)
             r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_DesNot = "PPG - " & p_NumOpe & " - TELEX. PASIVO"
             r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_FlagDH = "D"
             
             If InStr(moddat_g_str_AgrCME, r_str_CodPrd) Then
                r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_CtaCtb = IIf(p_CodMon = 1, "411703020104", "412703020104")
             ElseIf InStr(moddat_g_str_AgrMIHG, r_str_CodPrd) Then
                r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_CtaCtb = IIf(p_CodMon = 1, "411703020102", "412703020102")
             ElseIf InStr(moddat_g_str_AgrTFMV, r_str_CodPrd) Then
                r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_CtaCtb = IIf(p_CodMon = 1, "411703020106", "412703020106")
             End If
          End If
          If CDbl(ipp_MtoPort.Text) > 0 Then
             ReDim Preserve r_arr_Matriz(UBound(r_arr_Matriz) + 1)
             r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = CDbl(ipp_MtoPort.Text)
             r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_DesNot = "PPG - " & p_NumOpe & " - PORTES. PASIVO"
             r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_FlagDH = "D"
             
             If InStr(moddat_g_str_AgrCME, r_str_CodPrd) Then
                r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_CtaCtb = IIf(p_CodMon = 1, "411703020104", "412703020104")
             ElseIf InStr(moddat_g_str_AgrMIHG, r_str_CodPrd) Then
                r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_CtaCtb = IIf(p_CodMon = 1, "411703020102", "412703020102")
             ElseIf InStr(moddat_g_str_AgrTFMV, r_str_CodPrd) Then
                r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_CtaCtb = IIf(p_CodMon = 1, "411703020106", "412703020106")
             End If
          End If
          If CDbl(ipp_MtoMora.Text) > 0 Then
             ReDim Preserve r_arr_Matriz(UBound(r_arr_Matriz) + 1)
             r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = CDbl(ipp_MtoMora.Text)
             r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_DesNot = "PPG - " & p_NumOpe & " - MORA. PASIVO"
             r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_FlagDH = "D"
             
             If InStr(moddat_g_str_AgrCME, r_str_CodPrd) Then
                r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_CtaCtb = IIf(p_CodMon = 1, "411703020104", "412703020104")
             ElseIf InStr(moddat_g_str_AgrMIHG, r_str_CodPrd) Then
                r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_CtaCtb = IIf(p_CodMon = 1, "411703020102", "412703020102")
             ElseIf InStr(moddat_g_str_AgrTFMV, r_str_CodPrd) Then
                r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_CtaCtb = IIf(p_CodMon = 1, "411703020106", "412703020106")
             End If
          End If
          If CDbl(ipp_GasDiv.Text) > 0 Then
             ReDim Preserve r_arr_Matriz(UBound(r_arr_Matriz) + 1)
             r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = CDbl(ipp_GasDiv.Text)
             r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_DesNot = "PPG - " & p_NumOpe & " - GASTO.DIV. PASIVO"
             r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_FlagDH = "D"
             
             If InStr(moddat_g_str_AgrCME, r_str_CodPrd) Then
                r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_CtaCtb = IIf(p_CodMon = 1, "411703020104", "412703020104")
             ElseIf InStr(moddat_g_str_AgrMIHG, r_str_CodPrd) Then
                r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_CtaCtb = IIf(p_CodMon = 1, "411703020102", "412703020102")
             ElseIf InStr(moddat_g_str_AgrTFMV, r_str_CodPrd) Then
                r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_CtaCtb = IIf(p_CodMon = 1, "411703020106", "412703020106")
             End If
          End If
          If CDbl(ipp_ComCof.Text) > 0 Then
             ReDim Preserve r_arr_Matriz(UBound(r_arr_Matriz) + 1)
             r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = CDbl(ipp_ComCof.Text)
             r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_DesNot = "PPG - " & p_NumOpe & " - COM.COF. PASIVO"
             r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_FlagDH = "D"
             
             If InStr(moddat_g_str_AgrCME, r_str_CodPrd) Then
                r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_CtaCtb = IIf(p_CodMon = 1, "411703020104", "412703020104")
             ElseIf InStr(moddat_g_str_AgrMIHG, r_str_CodPrd) Then
                r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_CtaCtb = IIf(p_CodMon = 1, "411703020102", "412703020102")
             ElseIf InStr(moddat_g_str_AgrTFMV, r_str_CodPrd) Then
                r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_CtaCtb = IIf(p_CodMon = 1, "411703020106", "412703020106")
             End If
          End If
          If CDbl(ipp_ComExt.Text) > 0 Then
             ReDim Preserve r_arr_Matriz(UBound(r_arr_Matriz) + 1)
             r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = CDbl(ipp_ComExt.Text)
             r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_DesNot = "PPG - " & p_NumOpe & " - COM.EXT. PASIVO"
             r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_FlagDH = "D"
             
             If InStr(moddat_g_str_AgrCME, r_str_CodPrd) Then
                r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_CtaCtb = IIf(p_CodMon = 1, "411703020104", "412703020104")
             ElseIf InStr(moddat_g_str_AgrMIHG, r_str_CodPrd) Then
                r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_CtaCtb = IIf(p_CodMon = 1, "411703020102", "412703020102")
             ElseIf InStr(moddat_g_str_AgrTFMV, r_str_CodPrd) Then
                r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_CtaCtb = IIf(p_CodMon = 1, "411703020106", "412703020106")
             End If
          End If
          If CDbl(ipp_ComRenov.Text) > 0 Then
             ReDim Preserve r_arr_Matriz(UBound(r_arr_Matriz) + 1)
             r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = CDbl(ipp_ComRenov.Text)
             r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_CtaCtb = IIf(p_CodMon = 1, "411402020102", "412402020102")
             r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_DesNot = "PPG - " & p_NumOpe & " - COM.RENV. PASIVO"
             r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_FlagDH = "D"
             
             If InStr(moddat_g_str_AgrCME, r_str_CodPrd) Then
                r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_CtaCtb = IIf(p_CodMon = 1, "411703020104", "412703020104")
             ElseIf InStr(moddat_g_str_AgrMIHG, r_str_CodPrd) Then
                 r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_CtaCtb = IIf(p_CodMon = 1, "411703020102", "412703020102")
             ElseIf InStr(moddat_g_str_AgrTFMV, r_str_CodPrd) Then
                r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_CtaCtb = IIf(p_CodMon = 1, "411703020106", "412703020106")
             End If
          End If
          
          ReDim Preserve r_arr_Matriz(UBound(r_arr_Matriz) + 1)
          r_dbl_Import = 0
          r_dbl_Import = CDbl(ipp_Capital.Text) + CDbl(ipp_Interes.Text) + CDbl(ipp_ComComp.Text) + _
                         CDbl(ipp_ComOtr.Text) + CDbl(ipp_MtoTelex.Text) + CDbl(ipp_MtoPort.Text) + _
                         CDbl(ipp_MtoMora.Text) + CDbl(ipp_GasDiv.Text) + CDbl(ipp_ComCof.Text) + _
                         CDbl(ipp_ComExt.Text) + CDbl(ipp_ComRenov.Text)
          r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_import = r_dbl_Import
                                                             
          r_str_CtaHab = ""
          r_str_CtaHab = IIf(p_CodMon = 1, "251419010111", "252419010111")
          r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_CtaCtb = r_str_CtaHab
          r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_DesNot = "PPG - " & p_NumOpe & " - PROVISION PAGO"
          r_arr_Matriz(UBound(r_arr_Matriz)).Matriz_FlagDH = "H"
                       
         'Obteniendo Nro. de Asiento
         r_int_NumAsi = modprc_ff_NumAsi(r_arr_LogPro, r_int_PerAno, r_int_PerMes, r_str_Origen, r_int_NumLib)
         p_AsiGen = r_int_NumAsi
         
         'Insertar en CABECERA
         Call modprc_fs_Inserta_CabAsi(r_arr_LogPro, r_str_Origen, r_int_PerAno, r_int_PerMes, r_int_NumLib, r_int_NumAsi, Format(1, "000"), _
                                       r_dbl_TipCam, r_str_TipNot, "PPG CUOTA PASIVO COFIDE FMV - " & p_NumOpe, r_str_FechaL, "1")
                           
         r_int_NumIte = 1
         For r_int_AuxCon = 1 To UBound(r_arr_Matriz)
             r_dbl_ImpSol = 0
             r_dbl_ImpDol = 0
             If r_arr_Matriz(r_int_AuxCon).Matriz_import > 0 Then
                If p_CodMon = 1 Then
                   r_dbl_ImpSol = r_arr_Matriz(r_int_AuxCon).Matriz_import
                   r_dbl_ImpDol = CDbl(Format(r_arr_Matriz(r_int_AuxCon).Matriz_import / r_dbl_TipCam, "#######0.00"))
                ElseIf p_CodMon = 2 Then
                   r_dbl_ImpSol = CDbl(Format(r_arr_Matriz(r_int_AuxCon).Matriz_import * r_dbl_TipCam, "########0.00"))
                   r_dbl_ImpDol = r_arr_Matriz(r_int_AuxCon).Matriz_import
                End If
                
                Call modprc_fs_Inserta_DetAsi_PagVar(r_arr_LogPro, r_str_Origen, r_int_PerAno, r_int_PerMes, r_int_NumLib, r_int_NumAsi, r_int_NumIte, _
                                                     r_arr_Matriz(r_int_AuxCon).Matriz_CtaCtb, CDate(r_str_FechaL), UCase(r_arr_Matriz(r_int_AuxCon).Matriz_DesNot), _
                                                     r_arr_Matriz(r_int_AuxCon).Matriz_FlagDH, r_dbl_ImpSol, r_dbl_ImpDol, 1, CDate(r_str_FechaL))
                r_int_NumIte = r_int_NumIte + 1
             Else
                r_int_NumIte = r_int_NumIte
             End If
         Next
         
         'Actualiza flag de contabilizacion
         r_str_CadAux = ""
         r_str_CadAux = r_str_Origen & "/" & r_int_PerAno & "/" & Format(r_int_PerMes, "00") & "/" & Format(r_int_NumLib, "00") & "/" & r_int_NumAsi
         
         r_str_Cadena = ""
         r_str_Cadena = r_str_Cadena & "UPDATE CRE_PPGPASCTB "
         r_str_Cadena = r_str_Cadena & "   SET PPGPAS_FECCTB = " & Format(moddat_g_str_FecSis, "yyyymmdd") & ", "
         r_str_Cadena = r_str_Cadena & "       PPGPAS_DATCTB = '" & r_str_CadAux & "' "
         r_str_Cadena = r_str_Cadena & " WHERE PPGPAS_NUMOPE = '" & p_NumOpe & "' "
         r_str_Cadena = r_str_Cadena & "   AND PPGPAS_FECPPG = " & Format(pnl_FecPpg.Caption, "yyyymmdd")
         r_str_Cadena = r_str_Cadena & "   AND PPGPAS_CODREG = " & p_Codigo
         If Not gf_EjecutaSQL(r_str_Cadena, r_rst_Genera, 2) Then
            Exit Sub
         End If
      
         'Enviar a la tabla de autorizaciones
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & " USP_CNTBL_COMAUT ( "
         g_str_Parame = g_str_Parame & " " & CLng(p_Codigo) & ", " 'COMAUT_CODOPE
         g_str_Parame = g_str_Parame & " " & Format(r_str_FechaL, "yyyymmdd") & ", " 'COMAUT_FECOPE
         g_str_Parame = g_str_Parame & cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) & ", " 'COMAUT_TIPDOC
         g_str_Parame = g_str_Parame & " '" & fs_NumDoc(cmb_Proveedor.Text) & "', "    'COMAUT_NUMDOC
         g_str_Parame = g_str_Parame & " " & p_CodMon & ", " 'COMAUT_CODMON
         g_str_Parame = g_str_Parame & " " & CDbl(r_dbl_Import) & ", " 'COMAUT_IMPPAG
         g_str_Parame = g_str_Parame & cmb_Banco.ItemData(cmb_Banco.ListIndex) & ", "  'COMAUT_CODBNC
         g_str_Parame = g_str_Parame & "'" & Trim(cmb_CtaCte.Text) & "', "  'COMAUT_CTACRR
         g_str_Parame = g_str_Parame & " '" & r_str_CtaHab & "', "  'COMAUT_CTACTB
         g_str_Parame = g_str_Parame & " '" & r_str_CadAux & "',  " 'COMAUT_DATCTB
         g_str_Parame = g_str_Parame & " 'PPG " & IIf(Trim(pnl_TipPpg.Caption) = "TOTAL", "TOTAL", "PARCIAL") & "',  " 'COMAUT_DESCRIPCION Trim(pnl_TipPpg.Caption) = "TOTAL"
         g_str_Parame = g_str_Parame & " 1,  " 'COMAUT_TIPOPE
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', " 'SEGUSUCRE
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', " 'SEGPLTCRE
         g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "  'SEGTERCRE
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') " 'SEGSUCCRE

         If Not gf_EjecutaSQL(g_str_Parame, r_rst_Genera, 2) Then
            Exit Sub
         End If
         
      End If
   End If
   
End Sub

Private Sub fs_Calcular_Total()
Dim r_dbl_Total As Double
    
   r_dbl_Total = CDbl(ipp_Capital.Text) + CDbl(ipp_Interes.Text) + CDbl(ipp_ComComp.Text) + CDbl(ipp_ComOtr.Text) + CDbl(ipp_MtoTelex.Text) + CDbl(ipp_MtoPort.Text) + CDbl(ipp_MtoMora.Text) + _
                 CDbl(ipp_GasDiv.Text) + CDbl(ipp_ComCof.Text) + CDbl(ipp_ComExt.Text) + CDbl(ipp_ComRenov.Text) + CDbl(ipp_DevPBP.Text) + CDbl(ipp_IntLeg.Text)
   pnl_TotPag.Caption = Format(r_dbl_Total, "###,##0.00") & " "
End Sub

Private Sub cmb_TipDoc_Click()
   Call fs_CargarPrv
End Sub

Private Sub fs_CargarPrv()
   ReDim l_arr_CtaCteSol(0)
   ReDim l_arr_CtaCteDol(0)
   ReDim l_arr_MaePrv(0)
   cmb_Proveedor.Clear
   cmb_Proveedor.Text = ""
   cmb_Banco.Clear
   cmb_CtaCte.Clear
   
   If (cmb_TipDoc.ListIndex = -1) Then
       Exit Sub
   End If
    
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT A.MAEPRV_TIPDOC, A.MAEPRV_NUMDOC, A.MAEPRV_RAZSOC "
   g_str_Parame = g_str_Parame & "   FROM CNTBL_MAEPRV A "
   g_str_Parame = g_str_Parame & "  WHERE A.MAEPRV_TIPDOC = " & cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
   If moddat_g_int_FlgGrb = 1 Then 'INSERT
      g_str_Parame = g_str_Parame & " AND A.MAEPRV_SITUAC = 1 "
   End If
   g_str_Parame = g_str_Parame & "  ORDER BY A.MAEPRV_RAZSOC ASC "
      
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
      cmb_Proveedor.AddItem Trim(g_rst_Genera!maeprv_numdoc & "") & " - " & Trim(g_rst_Genera!MAEPRV_RAZSOC & "")
      
      ReDim Preserve l_arr_MaePrv(UBound(l_arr_MaePrv) + 1)
      l_arr_MaePrv(UBound(l_arr_MaePrv)).Genera_Codigo = Trim(g_rst_Genera!maeprv_numdoc & "")
      l_arr_MaePrv(UBound(l_arr_MaePrv)).Genera_Nombre = Trim(g_rst_Genera!MAEPRV_RAZSOC & "")
      
      g_rst_Genera.MoveNext
   Loop
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
End Sub

Private Sub fs_Buscar_prov()
Dim r_str_NumDoc As String

   ReDim l_arr_CtaCteSol(0)
   ReDim l_arr_CtaCteDol(0)
   cmb_Banco.Clear
   cmb_CtaCte.Clear
   r_str_NumDoc = ""
   
   If (moddat_g_int_FlgGrb = 1) Then
       If cmb_TipDoc.ListIndex = -1 Then
          MsgBox "Debe seleccionar el tipo de documento de identidad.", vbExclamation, modgen_g_str_NomPlt
          Call gs_SetFocus(cmb_TipDoc)
          Exit Sub
       End If
       If cmb_Proveedor.ListIndex = -1 Then
          Exit Sub
       End If
      
       If (fs_ValNumDoc() = False) Then
           Exit Sub
       End If
   End If
   
   r_str_NumDoc = fs_NumDoc(cmb_Proveedor.Text)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT A.MAEPRV_TIPDOC, A.MAEPRV_NUMDOC, A.MAEPRV_RAZSOC, "
   g_str_Parame = g_str_Parame & "        A.MAEPRV_CODBNC_MN1, A.MAEPRV_CTACRR_MN1, "
   g_str_Parame = g_str_Parame & "        A.MAEPRV_NROCCI_MN1, A.MAEPRV_CODBNC_MN2, A.MAEPRV_CTACRR_MN2, "
   g_str_Parame = g_str_Parame & "        A.MAEPRV_NROCCI_MN2, A.MAEPRV_CODBNC_MN3, A.MAEPRV_CTACRR_MN3, "
   g_str_Parame = g_str_Parame & "        A.MAEPRV_NROCCI_MN3, A.MAEPRV_CODBNC_DL1, A.MAEPRV_CTACRR_DL1, "
   g_str_Parame = g_str_Parame & "        A.MAEPRV_NROCCI_DL1, A.MAEPRV_CODBNC_DL2, A.MAEPRV_CTACRR_DL2, "
   g_str_Parame = g_str_Parame & "        A.MAEPRV_NROCCI_DL2, A.MAEPRV_CODBNC_DL3, A.MAEPRV_CTACRR_DL3, "
   g_str_Parame = g_str_Parame & "        A.MAEPRV_NROCCI_DL3, A.MAEPRV_CONDIC "
   g_str_Parame = g_str_Parame & "   FROM CNTBL_MAEPRV A "
   'If (moddat_g_int_FlgGrb = 1 Or moddat_g_int_FlgGrb = 2) Then
       g_str_Parame = g_str_Parame & "  WHERE A.MAEPRV_TIPDOC = " & cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
       g_str_Parame = g_str_Parame & "    AND TRIM(A.MAEPRV_NUMDOC) = '" & Trim(r_str_NumDoc) & "' "
   'End If
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
      Exit Sub
   End If
   
   If g_rst_GenAux.BOF And g_rst_GenAux.EOF Then
      MsgBox "No se ha encontrado el proveedor.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Proveedor)
      g_rst_GenAux.Close
      Set g_rst_GenAux = Nothing
      Exit Sub
   End If
   
   If (moddat_g_int_FlgGrb = 1) Then
       If (g_rst_GenAux!MAEPRV_CONDIC = 2) Then
          MsgBox "El proveedor se encuentra en condición de NO HABIDO, revisar sunat.", vbExclamation, modgen_g_str_NomPlt
          g_rst_GenAux.Close
          Set g_rst_GenAux = Nothing
          Exit Sub
       End If
       'Call gs_SetFocus(txt_Descrip)
   End If
      
   ReDim l_arr_CtaCteSol(0)
   ReDim l_arr_CtaCteDol(0)

   If (g_rst_GenAux!MAEPRV_CODBNC_MN1 <> 0) Then
       ReDim Preserve l_arr_CtaCteSol(UBound(l_arr_CtaCteSol) + 1)
       l_arr_CtaCteSol(UBound(l_arr_CtaCteSol)).Genera_Codigo = Trim(g_rst_GenAux!MAEPRV_CODBNC_MN1)
       l_arr_CtaCteSol(UBound(l_arr_CtaCteSol)).Genera_Nombre = Trim(moddat_gf_Consulta_ParDes("122", Format(g_rst_GenAux!MAEPRV_CODBNC_MN1, "000000")))
       l_arr_CtaCteSol(UBound(l_arr_CtaCteSol)).Genera_Prefij = Trim(g_rst_GenAux!MAEPRV_CTACRR_MN1 & "")
       l_arr_CtaCteSol(UBound(l_arr_CtaCteSol)).Genera_Refere = Trim(g_rst_GenAux!MAEPRV_NROCCI_MN1 & "")
       l_arr_CtaCteSol(UBound(l_arr_CtaCteSol)).Genera_TipMon = 1
   End If
   If (g_rst_GenAux!MAEPRV_CODBNC_MN2 <> 0) Then
       ReDim Preserve l_arr_CtaCteSol(UBound(l_arr_CtaCteSol) + 1)
       l_arr_CtaCteSol(UBound(l_arr_CtaCteSol)).Genera_Codigo = Trim(g_rst_GenAux!MAEPRV_CODBNC_MN2)
       l_arr_CtaCteSol(UBound(l_arr_CtaCteSol)).Genera_Nombre = Trim(moddat_gf_Consulta_ParDes("122", Format(g_rst_GenAux!MAEPRV_CODBNC_MN2, "000000")))
       l_arr_CtaCteSol(UBound(l_arr_CtaCteSol)).Genera_Prefij = Trim(g_rst_GenAux!MAEPRV_CTACRR_MN2 & "")
       l_arr_CtaCteSol(UBound(l_arr_CtaCteSol)).Genera_Refere = Trim(g_rst_GenAux!MAEPRV_NROCCI_MN2 & "")
       l_arr_CtaCteSol(UBound(l_arr_CtaCteSol)).Genera_TipMon = 1
   End If
   If (g_rst_GenAux!MAEPRV_CODBNC_MN3 <> 0) Then
       ReDim Preserve l_arr_CtaCteSol(UBound(l_arr_CtaCteSol) + 1)
       l_arr_CtaCteSol(UBound(l_arr_CtaCteSol)).Genera_Codigo = Trim(g_rst_GenAux!MAEPRV_CODBNC_MN3)
       l_arr_CtaCteSol(UBound(l_arr_CtaCteSol)).Genera_Nombre = Trim(moddat_gf_Consulta_ParDes("122", Format(g_rst_GenAux!MAEPRV_CODBNC_MN3, "000000")))
       l_arr_CtaCteSol(UBound(l_arr_CtaCteSol)).Genera_Prefij = Trim(g_rst_GenAux!MAEPRV_CTACRR_MN3 & "")
       l_arr_CtaCteSol(UBound(l_arr_CtaCteSol)).Genera_Refere = Trim(g_rst_GenAux!MAEPRV_NROCCI_MN3 & "")
       l_arr_CtaCteSol(UBound(l_arr_CtaCteSol)).Genera_TipMon = 1
   End If
   
   If (g_rst_GenAux!MAEPRV_CODBNC_DL1 <> 0) Then
       ReDim Preserve l_arr_CtaCteDol(UBound(l_arr_CtaCteDol) + 1)
       l_arr_CtaCteDol(UBound(l_arr_CtaCteDol)).Genera_Codigo = Trim(g_rst_GenAux!MAEPRV_CODBNC_DL1)
       l_arr_CtaCteDol(UBound(l_arr_CtaCteDol)).Genera_Nombre = Trim(moddat_gf_Consulta_ParDes("122", Format(g_rst_GenAux!MAEPRV_CODBNC_DL1, "000000")))
       l_arr_CtaCteDol(UBound(l_arr_CtaCteDol)).Genera_Prefij = Trim(g_rst_GenAux!MAEPRV_CTACRR_DL1 & "")
       l_arr_CtaCteDol(UBound(l_arr_CtaCteDol)).Genera_Refere = Trim(g_rst_GenAux!MAEPRV_NROCCI_DL1 & "")
       l_arr_CtaCteDol(UBound(l_arr_CtaCteDol)).Genera_TipMon = 2
   End If
   If (g_rst_GenAux!MAEPRV_CODBNC_DL2 <> 0) Then
       ReDim Preserve l_arr_CtaCteDol(UBound(l_arr_CtaCteDol) + 1)
       l_arr_CtaCteDol(UBound(l_arr_CtaCteDol)).Genera_Codigo = Trim(g_rst_GenAux!MAEPRV_CODBNC_DL2)
       l_arr_CtaCteDol(UBound(l_arr_CtaCteDol)).Genera_Nombre = Trim(moddat_gf_Consulta_ParDes("122", Format(g_rst_GenAux!MAEPRV_CODBNC_DL2, "000000")))
       l_arr_CtaCteDol(UBound(l_arr_CtaCteDol)).Genera_Prefij = Trim(g_rst_GenAux!MAEPRV_CTACRR_DL2 & "")
       l_arr_CtaCteDol(UBound(l_arr_CtaCteDol)).Genera_Refere = Trim(g_rst_GenAux!MAEPRV_NROCCI_DL2 & "")
       l_arr_CtaCteDol(UBound(l_arr_CtaCteDol)).Genera_TipMon = 2
   End If
   If (g_rst_GenAux!MAEPRV_CODBNC_DL3 <> 0) Then
       ReDim Preserve l_arr_CtaCteDol(UBound(l_arr_CtaCteDol) + 1)
       l_arr_CtaCteDol(UBound(l_arr_CtaCteDol)).Genera_Codigo = Trim(g_rst_GenAux!MAEPRV_CODBNC_DL3)
       l_arr_CtaCteDol(UBound(l_arr_CtaCteDol)).Genera_Nombre = Trim(moddat_gf_Consulta_ParDes("122", Format(g_rst_GenAux!MAEPRV_CODBNC_DL3, "000000")))
       l_arr_CtaCteDol(UBound(l_arr_CtaCteDol)).Genera_Prefij = Trim(g_rst_GenAux!MAEPRV_CTACRR_DL3 & "")
       l_arr_CtaCteDol(UBound(l_arr_CtaCteDol)).Genera_Refere = Trim(g_rst_GenAux!MAEPRV_NROCCI_DL3 & "")
       l_arr_CtaCteDol(UBound(l_arr_CtaCteDol)).Genera_TipMon = 2
   End If
   
   Call fs_CargarBancos
   
   g_rst_GenAux.Close
   Set g_rst_GenAux = Nothing
End Sub

Private Sub fs_CargarBancos()
Dim r_bol_Estado   As Boolean
Dim r_int_File     As Integer
   cmb_Banco.Clear
   cmb_CtaCte.Clear
   
   'If (cmb_Moneda.ListIndex = -1) Then
   '    Exit Sub
   'End If
   
   'soles
   'If (cmb_Moneda.ListIndex = 0) Then
   If (l_int_CodMon = 1) Then
       For l_int_Contar = 1 To UBound(l_arr_CtaCteSol)
           r_bol_Estado = True
           For r_int_File = 0 To cmb_Banco.ListCount - 1
               If (Trim(cmb_Banco.ItemData(r_int_File)) = Trim(l_arr_CtaCteSol(l_int_Contar).Genera_Codigo)) Then
                   r_bol_Estado = False
                   Exit For
               End If
           Next
           If (r_bol_Estado = True) Then
               cmb_Banco.AddItem Trim(l_arr_CtaCteSol(l_int_Contar).Genera_Nombre)
               cmb_Banco.ItemData(cmb_Banco.NewIndex) = Trim(l_arr_CtaCteSol(l_int_Contar).Genera_Codigo)
           End If
       Next
   End If
   'dolares
   If l_int_CodMon = 2 Then
       For l_int_Contar = 1 To UBound(l_arr_CtaCteDol)
           r_bol_Estado = True
           For r_int_File = 0 To cmb_Banco.ListCount - 1
               If (Trim(cmb_Banco.ItemData(r_int_File)) = Trim(l_arr_CtaCteDol(l_int_Contar).Genera_Codigo)) Then
                   r_bol_Estado = False
                   Exit For
               End If
           Next
           If (r_bol_Estado = True) Then
               cmb_Banco.AddItem Trim(l_arr_CtaCteDol(l_int_Contar).Genera_Nombre)
               cmb_Banco.ItemData(cmb_Banco.NewIndex) = Trim(l_arr_CtaCteDol(l_int_Contar).Genera_Codigo)
           End If
       Next
   End If
End Sub

Private Function fs_ValNumDoc() As Boolean
Dim r_str_NumDoc As String
   fs_ValNumDoc = True
   r_str_NumDoc = ""

   r_str_NumDoc = fs_NumDoc(cmb_Proveedor.Text)
   If (cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) = 1) Then 'DNI - 8
       If Len(Trim(r_str_NumDoc)) <> 8 Then
          MsgBox "El documento de identidad es de 8 digitos.", vbExclamation, modgen_g_str_NomPlt
          Call gs_SetFocus(cmb_Proveedor)
          fs_ValNumDoc = False
       End If
   ElseIf (cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) = 6) Then 'RUC - 11
       If Not gf_Valida_RUC(Trim(r_str_NumDoc), Mid(Trim(r_str_NumDoc), Len(Trim(r_str_NumDoc)), 1)) Then
          MsgBox "El Número de RUC no es valido.", vbExclamation, modgen_g_str_NomPlt
          Call gs_SetFocus(cmb_Proveedor)
          fs_ValNumDoc = False
       End If
   Else 'OTROS
       If Len(Trim(cmb_Proveedor.Text)) = 0 Then
          MsgBox "Debe ingresar un numero de documento.", vbExclamation, modgen_g_str_NomPlt
          Call gs_SetFocus(cmb_Proveedor)
          fs_ValNumDoc = False
       End If
   End If
End Function

Private Function fs_NumDoc(p_Cadena As String) As String
   fs_NumDoc = ""
   If (cmb_TipDoc.ListIndex > -1) Then
      If (cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) = 1) Then
          fs_NumDoc = Mid(p_Cadena, 1, 8)
      ElseIf (cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) = 6) Then
          fs_NumDoc = Mid(p_Cadena, 1, 11)
      Else
          If p_Cadena <> "" Then
             fs_NumDoc = Trim(Mid(p_Cadena, 1, InStr(Trim(p_Cadena), "-") - 1))
          End If
      End If
   End If
End Function

Private Sub ipp_Capital_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Interes)
   End If
End Sub

Private Sub ipp_Capital_LostFocus()
   Call fs_Calcular_Total
End Sub

Private Sub ipp_Interes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ComComp)
   End If
End Sub

Private Sub ipp_Interes_LostFocus()
   Call fs_Calcular_Total
End Sub

Private Sub ipp_ComComp_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ComOtr)
   End If
End Sub

Private Sub ipp_ComComp_LostFocus()
   Call fs_Calcular_Total
End Sub

Private Sub ipp_ComOtr_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_MtoTelex)
   End If
End Sub

Private Sub ipp_ComOtr_LostFocus()
   Call fs_Calcular_Total
End Sub

Private Sub ipp_MtoTelex_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_MtoPort)
   End If
End Sub

Private Sub ipp_MtoTelex_LostFocus()
   Call fs_Calcular_Total
End Sub

Private Sub ipp_MtoPort_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_MtoMora)
   End If
End Sub

Private Sub ipp_MtoPort_LostFocus()
   Call fs_Calcular_Total
End Sub

Private Sub ipp_MtoMora_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_GasDiv)
   End If
End Sub

Private Sub ipp_MtoMora_LostFocus()
   Call fs_Calcular_Total
End Sub

Private Sub ipp_GasDiv_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ComCof)
   End If
End Sub

Private Sub ipp_GasDiv_LostFocus()
   Call fs_Calcular_Total
End Sub

Private Sub ipp_ComCof_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ComExt)
   End If
End Sub

Private Sub ipp_ComCof_LostFocus()
   Call fs_Calcular_Total
End Sub

Private Sub ipp_ComExt_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ComRenov)
   End If
End Sub

Private Sub ipp_ComExt_LostFocus()
   Call fs_Calcular_Total
End Sub

Private Sub ipp_ComRenov_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_DevPBP)
   End If
End Sub

Private Sub ipp_ComRenov_LostFocus()
   Call fs_Calcular_Total
End Sub

Private Sub ipp_DevPBP_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_IntLeg)
   End If
End Sub

Private Sub ipp_DevPBP_LostFocus()
   Call fs_Calcular_Total
End Sub

Private Sub ipp_IntLeg_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   End If
End Sub

Private Sub ipp_IntLeg_LostFocus()
   Call fs_Calcular_Total
End Sub

