VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "msmapi32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_Desemb_05 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form4"
   ClientHeight    =   10410
   ClientLeft      =   2760
   ClientTop       =   675
   ClientWidth     =   11640
   Icon            =   "OpeTra_frm_036.frx":0000
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10410
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   10395
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   11640
      _Version        =   65536
      _ExtentX        =   20532
      _ExtentY        =   18336
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
      Begin Threed.SSPanel SSPanel3 
         Height          =   4395
         Left            =   30
         TabIndex        =   67
         Top             =   5130
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   7752
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
         Begin VB.TextBox txt_Observ 
            Height          =   705
            Left            =   1590
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   12
            Text            =   "OpeTra_frm_036.frx":000C
            Top             =   3630
            Width           =   9885
         End
         Begin VB.ComboBox cmb_MonFia 
            Height          =   315
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   3180
            Width           =   3225
         End
         Begin VB.ComboBox cmb_BanFia 
            Height          =   315
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   2520
            Width           =   3225
         End
         Begin VB.TextBox txt_NumFia 
            Height          =   315
            Left            =   7590
            MaxLength       =   25
            TabIndex        =   7
            Text            =   "Text1"
            Top             =   2520
            Width           =   3225
         End
         Begin VB.ComboBox cmb_FlgFia 
            Height          =   315
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   2190
            Width           =   915
         End
         Begin Threed.SSPanel SSPanel4 
            Height          =   60
            Left            =   60
            TabIndex        =   77
            Top             =   2070
            Width           =   11445
            _Version        =   65536
            _ExtentX        =   20188
            _ExtentY        =   106
            _StockProps     =   15
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
         End
         Begin VB.ComboBox cmb_MonDes 
            Height          =   315
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   3225
         End
         Begin VB.ComboBox cmb_BanChq 
            Height          =   315
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   1710
            Width           =   3225
         End
         Begin VB.TextBox txt_NumChq 
            Height          =   315
            Left            =   1590
            MaxLength       =   25
            TabIndex        =   2
            Text            =   "Text1"
            Top             =   1380
            Width           =   3225
         End
         Begin VB.ComboBox cmb_CtaChq 
            Height          =   315
            Left            =   7590
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   1710
            Width           =   3225
         End
         Begin EditLib.fpDoubleSingle ipp_MtoDes 
            Height          =   315
            Left            =   1590
            TabIndex        =   1
            Top             =   390
            Width           =   1575
            _Version        =   196608
            _ExtentX        =   2778
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
         Begin Threed.SSPanel pnl_TCaMPr 
            Height          =   315
            Left            =   7590
            TabIndex        =   73
            Top             =   390
            Width           =   1785
            _Version        =   65536
            _ExtentX        =   3149
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "9,999.9999 "
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
         Begin Threed.SSPanel pnl_TCaDol 
            Height          =   315
            Left            =   7590
            TabIndex        =   74
            Top             =   60
            Width           =   1785
            _Version        =   65536
            _ExtentX        =   3149
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "9,999.9999 "
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
         Begin EditLib.fpDateTime ipp_FVcFia 
            Height          =   315
            Left            =   7590
            TabIndex        =   9
            Top             =   2850
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
         Begin EditLib.fpDateTime ipp_FEmFia 
            Height          =   315
            Left            =   1590
            TabIndex        =   8
            Top             =   2850
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
         Begin EditLib.fpDoubleSingle ipp_MtoFia 
            Height          =   315
            Left            =   7590
            TabIndex        =   11
            Top             =   3180
            Width           =   1455
            _Version        =   196608
            _ExtentX        =   2566
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
         Begin Threed.SSPanel SSPanel5 
            Height          =   60
            Left            =   60
            TabIndex        =   85
            Top             =   3540
            Width           =   11445
            _Version        =   65536
            _ExtentX        =   20188
            _ExtentY        =   106
            _StockProps     =   15
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
         End
         Begin Threed.SSPanel pnl_DesMPr 
            Height          =   315
            Left            =   3240
            TabIndex        =   87
            Top             =   390
            Width           =   1575
            _Version        =   65536
            _ExtentX        =   2778
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "999,999.99 "
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
         Begin Threed.SSPanel pnl_NomPry 
            Height          =   315
            Left            =   1590
            TabIndex        =   88
            Top             =   720
            Width           =   9855
            _Version        =   65536
            _ExtentX        =   17383
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "999,999.99 "
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
         Begin Threed.SSPanel pnl_ProCon 
            Height          =   315
            Left            =   1590
            TabIndex        =   90
            Top             =   1050
            Width           =   9855
            _Version        =   65536
            _ExtentX        =   17383
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "999,999.99 "
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
         Begin VB.Label lbl_NomGlo 
            Caption         =   "Promotor / Constr.:"
            Height          =   285
            Index           =   8
            Left            =   60
            TabIndex        =   91
            Top             =   1050
            Width           =   1395
         End
         Begin VB.Label lbl_NomGlo 
            Caption         =   "Proyecto miCasita:"
            Height          =   285
            Index           =   6
            Left            =   60
            TabIndex        =   89
            Top             =   720
            Width           =   1395
         End
         Begin VB.Label lbl_NomGlo 
            Caption         =   "Observaciones:"
            Height          =   285
            Index           =   10
            Left            =   60
            TabIndex        =   86
            Top             =   3630
            Width           =   1335
         End
         Begin VB.Label lbl_NomGlo 
            Caption         =   "Moneda Fianza:"
            Height          =   315
            Index           =   3
            Left            =   60
            TabIndex        =   84
            Top             =   3180
            Width           =   1455
         End
         Begin VB.Label lbl_NomGlo 
            Caption         =   "Banco:"
            Height          =   315
            Index           =   2
            Left            =   60
            TabIndex        =   83
            Top             =   2520
            Width           =   1365
         End
         Begin VB.Label lbl_NomGlo 
            Caption         =   "Monto Fianza:"
            Height          =   285
            Index           =   1
            Left            =   5820
            TabIndex        =   82
            Top             =   3180
            Width           =   1395
         End
         Begin VB.Label Label7 
            Caption         =   "Fecha Emisión:"
            Height          =   315
            Left            =   60
            TabIndex        =   81
            Top             =   2850
            Width           =   1425
         End
         Begin VB.Label Label5 
            Caption         =   "Fecha Vcto.:"
            Height          =   315
            Left            =   5820
            TabIndex        =   80
            Top             =   2850
            Width           =   1425
         End
         Begin VB.Label Label8 
            Caption         =   "Nro. Carta Fianza:"
            Height          =   285
            Left            =   5820
            TabIndex        =   79
            Top             =   2520
            Width           =   1425
         End
         Begin VB.Label lbl_NomGlo 
            Caption         =   "Contra Carta Fianza:"
            Height          =   315
            Index           =   0
            Left            =   60
            TabIndex        =   78
            Top             =   2190
            Width           =   1485
         End
         Begin VB.Label lbl_NomGlo 
            Caption         =   "Tipo Cambio US$:"
            Height          =   285
            Index           =   13
            Left            =   5820
            TabIndex        =   76
            Top             =   60
            Width           =   1635
         End
         Begin VB.Label lbl_NomGlo 
            Caption         =   "Tipo Cambio M. Prest.:"
            Height          =   285
            Index           =   14
            Left            =   5820
            TabIndex        =   75
            Top             =   390
            Width           =   1665
         End
         Begin VB.Label lbl_NomGlo 
            Caption         =   "Importe Desemb.:"
            Height          =   285
            Index           =   5
            Left            =   60
            TabIndex        =   72
            Top             =   390
            Width           =   1395
         End
         Begin VB.Label lbl_NomGlo 
            Caption         =   "Moneda Desemb.:"
            Height          =   315
            Index           =   4
            Left            =   60
            TabIndex        =   71
            Top             =   60
            Width           =   1455
         End
         Begin VB.Label lbl_NomGlo 
            Caption         =   "Nro. Cuenta:"
            Height          =   285
            Index           =   11
            Left            =   5820
            TabIndex        =   70
            Top             =   1710
            Width           =   1485
         End
         Begin VB.Label lbl_NomGlo 
            Caption         =   "Banco:"
            Height          =   315
            Index           =   7
            Left            =   60
            TabIndex        =   69
            Top             =   1710
            Width           =   1365
         End
         Begin VB.Label lbl_NomGlo 
            Caption         =   "Nro. Cheque:"
            Height          =   285
            Index           =   16
            Left            =   60
            TabIndex        =   68
            Top             =   1380
            Width           =   1395
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   16
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
            TabIndex        =   17
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
            Picture         =   "OpeTra_frm_036.frx":0012
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   1425
         Left            =   30
         TabIndex        =   18
         Top             =   750
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   2514
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
            TabIndex        =   19
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
         Begin Threed.SSPanel pnl_NumSol 
            Height          =   315
            Left            =   1440
            TabIndex        =   20
            Top             =   390
            Width           =   2535
            _Version        =   65536
            _ExtentX        =   4471
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
         Begin Threed.SSPanel pnl_NumOpe 
            Height          =   315
            Left            =   1440
            TabIndex        =   21
            Top             =   720
            Width           =   2535
            _Version        =   65536
            _ExtentX        =   4471
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
         Begin Threed.SSPanel pnl_GenOpe 
            Height          =   315
            Left            =   10050
            TabIndex        =   22
            Top             =   60
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2514
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
         Begin Threed.SSPanel pnl_Client 
            Height          =   315
            Left            =   1440
            TabIndex        =   28
            Top             =   1050
            Width           =   10035
            _Version        =   65536
            _ExtentX        =   17701
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "1-07521154 / IKEHARA PUNK MIGUEL ANGEL"
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
         Begin VB.Label Label2 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   60
            TabIndex        =   29
            Top             =   1050
            Width           =   1125
         End
         Begin VB.Label Label1 
            Caption         =   "Nro. Solicitud:"
            Height          =   315
            Left            =   60
            TabIndex        =   26
            Top             =   390
            Width           =   1335
         End
         Begin VB.Label Label21 
            Caption         =   "Producto:"
            Height          =   315
            Left            =   60
            TabIndex        =   25
            Top             =   60
            Width           =   1275
         End
         Begin VB.Label Label12 
            Caption         =   "Nro. Operación:"
            Height          =   315
            Left            =   60
            TabIndex        =   24
            Top             =   720
            Width           =   1245
         End
         Begin VB.Label Label13 
            Caption         =   "F. Generac. Operac.:"
            Height          =   315
            Left            =   8400
            TabIndex        =   23
            Top             =   60
            Width           =   1545
         End
      End
      Begin Threed.SSPanel SSPanel39 
         Height          =   765
         Left            =   30
         TabIndex        =   27
         Top             =   9570
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
         Begin VB.CommandButton cmd_Recalc 
            Height          =   675
            Left            =   5130
            Picture         =   "OpeTra_frm_036.frx":031C
            Style           =   1  'Graphical
            TabIndex        =   146
            ToolTipText     =   "Aprobar Solicitud"
            Top             =   30
            Width           =   675
         End
         Begin VB.ComboBox cmb_TipCal 
            Height          =   315
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   144
            Top             =   180
            Width           =   3525
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   10830
            Picture         =   "OpeTra_frm_036.frx":0626
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   675
            Left            =   10140
            Picture         =   "OpeTra_frm_036.frx":0A68
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Aprobar Solicitud"
            Top             =   30
            Width           =   675
         End
         Begin MSMAPI.MAPIMessages mps_Mensaj 
            Left            =   8460
            Top             =   150
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
            Left            =   7890
            Top             =   150
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            DownloadMail    =   -1  'True
            LogonUI         =   -1  'True
            NewSession      =   0   'False
         End
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   7290
            Top             =   180
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
         Begin VB.Label lbl_NomGlo 
            Caption         =   "Cronog. NC COFIDE Mihogar:"
            Height          =   465
            Index           =   9
            Left            =   60
            TabIndex        =   145
            Top             =   180
            Width           =   1335
         End
      End
      Begin Threed.SSPanel SSPanel22 
         Height          =   2865
         Left            =   30
         TabIndex        =   30
         Top             =   2220
         Width           =   11535
         _Version        =   65536
         _ExtentX        =   20346
         _ExtentY        =   5054
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
         Begin TabDlg.SSTab tab_Cronog 
            Height          =   2775
            Left            =   60
            TabIndex        =   31
            Top             =   60
            Width           =   11385
            _ExtentX        =   20082
            _ExtentY        =   4895
            _Version        =   393216
            Style           =   1
            Tabs            =   5
            TabsPerRow      =   5
            TabHeight       =   520
            TabCaption(0)   =   "Cronograma - Cliente TNC"
            TabPicture(0)   =   "OpeTra_frm_036.frx":0EAA
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "pnl_CliNCo_TotCuo"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "pnl_CliNCo_OtrCar"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "SSPanel62"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "SSPanel61"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "SSPanel59"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "SSPanel36"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "SSPanel35"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "SSPanel34"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).Control(8)=   "SSPanel33"
            Tab(0).Control(8).Enabled=   0   'False
            Tab(0).Control(9)=   "SSPanel2"
            Tab(0).Control(9).Enabled=   0   'False
            Tab(0).Control(10)=   "grd_CliNCo_Listad"
            Tab(0).Control(10).Enabled=   0   'False
            Tab(0).Control(11)=   "pnl_CliNCo_Intere"
            Tab(0).Control(11).Enabled=   0   'False
            Tab(0).Control(12)=   "pnl_CliNCo_SegPre"
            Tab(0).Control(12).Enabled=   0   'False
            Tab(0).Control(13)=   "pnl_CliNCo_SegViv"
            Tab(0).Control(13).Enabled=   0   'False
            Tab(0).Control(14)=   "pnl_CliNCo_Capita"
            Tab(0).Control(14).Enabled=   0   'False
            Tab(0).Control(15)=   "SSPanel30"
            Tab(0).Control(15).Enabled=   0   'False
            Tab(0).ControlCount=   16
            TabCaption(1)   =   "Cliente - Tramo Concesional"
            TabPicture(1)   =   "OpeTra_frm_036.frx":0EC6
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "pnl_CliCon_Capita"
            Tab(1).Control(1)=   "pnl_CliCon_Intere"
            Tab(1).Control(2)=   "SSPanel21"
            Tab(1).Control(3)=   "SSPanel13"
            Tab(1).Control(4)=   "SSPanel12"
            Tab(1).Control(5)=   "SSPanel11"
            Tab(1).Control(6)=   "SSPanel10"
            Tab(1).Control(7)=   "grd_CliCon_Listad"
            Tab(1).Control(8)=   "SSPanel9"
            Tab(1).Control(9)=   "pnl_CliCon_TotCuo"
            Tab(1).ControlCount=   10
            TabCaption(2)   =   "Mivivienda - Tramo No Concesional"
            TabPicture(2)   =   "OpeTra_frm_036.frx":0EE2
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "pnl_MViNCo_Comisi"
            Tab(2).Control(1)=   "pnl_MViNCo_TotCuo"
            Tab(2).Control(2)=   "pnl_MViNCo_OtrCar"
            Tab(2).Control(3)=   "pnl_MViNCo_Intere"
            Tab(2).Control(4)=   "pnl_MViNCo_SegPre"
            Tab(2).Control(5)=   "pnl_MViNCo_SegViv"
            Tab(2).Control(6)=   "pnl_MViNCo_Capita"
            Tab(2).Control(7)=   "SSPanel49"
            Tab(2).Control(8)=   "SSPanel47"
            Tab(2).Control(9)=   "SSPanel46"
            Tab(2).Control(10)=   "SSPanel45"
            Tab(2).Control(11)=   "SSPanel44"
            Tab(2).Control(12)=   "SSPanel43"
            Tab(2).Control(13)=   "SSPanel42"
            Tab(2).Control(14)=   "SSPanel41"
            Tab(2).Control(15)=   "SSPanel38"
            Tab(2).Control(16)=   "SSPanel8"
            Tab(2).Control(17)=   "grd_MViNCo_Listad"
            Tab(2).ControlCount=   18
            TabCaption(3)   =   "Mivivienda - Tramo Concesional"
            TabPicture(3)   =   "OpeTra_frm_036.frx":0EFE
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "pnl_MViCon_TotCuo"
            Tab(3).Control(1)=   "SSPanel14"
            Tab(3).Control(2)=   "grd_MviCon_Listad"
            Tab(3).Control(3)=   "SSPanel15"
            Tab(3).Control(4)=   "SSPanel16"
            Tab(3).Control(5)=   "SSPanel17"
            Tab(3).Control(6)=   "SSPanel18"
            Tab(3).Control(7)=   "SSPanel19"
            Tab(3).Control(8)=   "SSPanel20"
            Tab(3).Control(9)=   "pnl_MViCon_Intere"
            Tab(3).Control(10)=   "pnl_MViCon_Capita"
            Tab(3).Control(11)=   "pnl_MViCon_Comisi"
            Tab(3).ControlCount=   12
            TabCaption(4)   =   "Cofide"
            TabPicture(4)   =   "OpeTra_frm_036.frx":0F1A
            Tab(4).ControlEnabled=   0   'False
            Tab(4).Control(0)=   "pnl_CofNCo_TotCuo"
            Tab(4).Control(1)=   "SSPanel55"
            Tab(4).Control(2)=   "grd_CofNCo_Listad"
            Tab(4).Control(3)=   "SSPanel56"
            Tab(4).Control(4)=   "SSPanel58"
            Tab(4).Control(5)=   "SSPanel60"
            Tab(4).Control(6)=   "SSPanel63"
            Tab(4).Control(7)=   "SSPanel64"
            Tab(4).Control(8)=   "SSPanel65"
            Tab(4).Control(9)=   "pnl_CofNCo_Intere"
            Tab(4).Control(10)=   "pnl_CofNCo_Capita"
            Tab(4).Control(11)=   "pnl_CofNCo_Comisi"
            Tab(4).ControlCount=   12
            Begin Threed.SSPanel pnl_MViCon_TotCuo 
               Height          =   285
               Left            =   -67470
               TabIndex        =   102
               Top             =   2370
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
            Begin Threed.SSPanel SSPanel30 
               Height          =   285
               Left            =   3450
               TabIndex        =   32
               Top             =   360
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Interés"
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
            Begin Threed.SSPanel pnl_CliNCo_Capita 
               Height          =   285
               Left            =   2280
               TabIndex        =   33
               Top             =   2370
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
            Begin Threed.SSPanel pnl_CliNCo_SegViv 
               Height          =   285
               Left            =   5790
               TabIndex        =   34
               Top             =   2370
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
            Begin Threed.SSPanel pnl_CliNCo_SegPre 
               Height          =   285
               Left            =   4620
               TabIndex        =   35
               Top             =   2370
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
            Begin Threed.SSPanel pnl_CliNCo_Intere 
               Height          =   285
               Left            =   3450
               TabIndex        =   36
               Top             =   2370
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
            Begin MSFlexGridLib.MSFlexGrid grd_CliNCo_Listad 
               Height          =   1695
               Left            =   30
               TabIndex        =   37
               Top             =   660
               Width           =   11265
               _ExtentX        =   19870
               _ExtentY        =   2990
               _Version        =   393216
               Rows            =   21
               Cols            =   9
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   49152
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel SSPanel23 
               Height          =   285
               Left            =   -67530
               TabIndex        =   38
               Top             =   360
               Width           =   2370
               _Version        =   65536
               _ExtentX        =   4180
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Total Cuota"
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
            Begin Threed.SSPanel SSPanel25 
               Height          =   285
               Left            =   -65190
               TabIndex        =   39
               Top             =   360
               Width           =   2370
               _Version        =   65536
               _ExtentX        =   4180
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Saldo Capital"
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
            Begin Threed.SSPanel SSPanel26 
               Height          =   285
               Left            =   -74940
               TabIndex        =   40
               Top             =   360
               Width           =   1185
               _Version        =   65536
               _ExtentX        =   2090
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Cuota"
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
            Begin Threed.SSPanel SSPanel27 
               Height          =   285
               Left            =   -73770
               TabIndex        =   41
               Top             =   360
               Width           =   1815
               _Version        =   65536
               _ExtentX        =   3201
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "F. Vcto"
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
            Begin Threed.SSPanel SSPanel28 
               Height          =   285
               Left            =   -71970
               TabIndex        =   42
               Top             =   360
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Capital"
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
            Begin Threed.SSPanel SSPanel29 
               Height          =   285
               Left            =   -70140
               TabIndex        =   43
               Top             =   360
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Interés"
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
            Begin Threed.SSPanel SSPanel31 
               Height          =   285
               Left            =   -66480
               TabIndex        =   44
               Top             =   360
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Total Cuota"
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
            Begin Threed.SSPanel SSPanel32 
               Height          =   285
               Left            =   -64650
               TabIndex        =   45
               Top             =   360
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Saldo Capital"
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
            Begin Threed.SSPanel SSPanel37 
               Height          =   285
               Left            =   -68310
               TabIndex        =   46
               Top             =   360
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Comisión"
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
            Begin Threed.SSPanel SSPanel40 
               Height          =   285
               Left            =   -74940
               TabIndex        =   47
               Top             =   360
               Width           =   1185
               _Version        =   65536
               _ExtentX        =   2090
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Cuota"
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
            Begin Threed.SSPanel SSPanel48 
               Height          =   285
               Left            =   -73770
               TabIndex        =   48
               Top             =   360
               Width           =   1815
               _Version        =   65536
               _ExtentX        =   3201
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "F. Vcto"
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
            Begin Threed.SSPanel SSPanel50 
               Height          =   285
               Left            =   -71970
               TabIndex        =   49
               Top             =   360
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Capital"
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
            Begin Threed.SSPanel SSPanel51 
               Height          =   285
               Left            =   -70140
               TabIndex        =   50
               Top             =   360
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Interés"
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
            Begin Threed.SSPanel SSPanel52 
               Height          =   285
               Left            =   -66480
               TabIndex        =   51
               Top             =   360
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Total Cuota"
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
            Begin Threed.SSPanel SSPanel53 
               Height          =   285
               Left            =   -64650
               TabIndex        =   52
               Top             =   360
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Saldo Capital"
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
            Begin Threed.SSPanel SSPanel57 
               Height          =   285
               Left            =   -68310
               TabIndex        =   53
               Top             =   360
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Comisión"
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
            Begin Threed.SSPanel SSPanel2 
               Height          =   285
               Left            =   60
               TabIndex        =   54
               Top             =   360
               Width           =   795
               _Version        =   65536
               _ExtentX        =   1402
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Cuota"
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
            Begin Threed.SSPanel SSPanel33 
               Height          =   285
               Left            =   840
               TabIndex        =   55
               Top             =   360
               Width           =   1455
               _Version        =   65536
               _ExtentX        =   2566
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "F. Vcto"
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
            Begin Threed.SSPanel SSPanel34 
               Height          =   285
               Left            =   2280
               TabIndex        =   56
               Top             =   360
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Capital"
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
            Begin Threed.SSPanel SSPanel35 
               Height          =   285
               Left            =   8130
               TabIndex        =   57
               Top             =   360
               Width           =   1320
               _Version        =   65536
               _ExtentX        =   2328
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Total Cuota"
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
            Begin Threed.SSPanel SSPanel36 
               Height          =   285
               Left            =   9420
               TabIndex        =   58
               Top             =   360
               Width           =   1560
               _Version        =   65536
               _ExtentX        =   2752
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Saldo Capital"
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
            Begin Threed.SSPanel SSPanel59 
               Height          =   285
               Left            =   4620
               TabIndex        =   59
               Top             =   360
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Seg. Prest."
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
            Begin Threed.SSPanel SSPanel61 
               Height          =   285
               Left            =   5790
               TabIndex        =   60
               Top             =   360
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Seg. Vivienda"
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
            Begin Threed.SSPanel SSPanel62 
               Height          =   285
               Left            =   6960
               TabIndex        =   61
               Top             =   360
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Portes"
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
            Begin Threed.SSPanel pnl_CliNCo_OtrCar 
               Height          =   285
               Left            =   6960
               TabIndex        =   62
               Top             =   2370
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
            Begin Threed.SSPanel pnl_CliNCo_TotCuo 
               Height          =   285
               Left            =   8130
               TabIndex        =   63
               Top             =   2370
               Width           =   1320
               _Version        =   65536
               _ExtentX        =   2328
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
            Begin Threed.SSPanel SSPanel14 
               Height          =   285
               Left            =   -70950
               TabIndex        =   92
               Top             =   360
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Interés"
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
            Begin MSFlexGridLib.MSFlexGrid grd_MviCon_Listad 
               Height          =   1695
               Left            =   -74970
               TabIndex        =   93
               Top             =   660
               Width           =   11265
               _ExtentX        =   19870
               _ExtentY        =   2990
               _Version        =   393216
               Rows            =   21
               Cols            =   7
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   49152
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel SSPanel15 
               Height          =   285
               Left            =   -74940
               TabIndex        =   94
               Top             =   360
               Width           =   765
               _Version        =   65536
               _ExtentX        =   1349
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Cuota"
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
            Begin Threed.SSPanel SSPanel16 
               Height          =   285
               Left            =   -74190
               TabIndex        =   95
               Top             =   360
               Width           =   1515
               _Version        =   65536
               _ExtentX        =   2672
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "F. Vcto"
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
            Begin Threed.SSPanel SSPanel17 
               Height          =   285
               Left            =   -72690
               TabIndex        =   96
               Top             =   360
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Capital"
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
            Begin Threed.SSPanel SSPanel18 
               Height          =   285
               Left            =   -67470
               TabIndex        =   97
               Top             =   360
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Total Cuota"
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
            Begin Threed.SSPanel SSPanel19 
               Height          =   285
               Left            =   -65730
               TabIndex        =   98
               Top             =   360
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Saldo Capital"
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
            Begin Threed.SSPanel SSPanel20 
               Height          =   285
               Left            =   -69210
               TabIndex        =   99
               Top             =   360
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Comisión COFIDE"
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
            Begin Threed.SSPanel pnl_MViCon_Intere 
               Height          =   285
               Left            =   -70950
               TabIndex        =   100
               Top             =   2370
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
            Begin Threed.SSPanel pnl_MViCon_Capita 
               Height          =   285
               Left            =   -72690
               TabIndex        =   101
               Top             =   2370
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
            Begin Threed.SSPanel pnl_MViCon_Comisi 
               Height          =   285
               Left            =   -69210
               TabIndex        =   103
               Top             =   2370
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
            Begin Threed.SSPanel pnl_CliCon_TotCuo 
               Height          =   285
               Left            =   -68370
               TabIndex        =   104
               Top             =   2370
               Width           =   2170
               _Version        =   65536
               _ExtentX        =   3828
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
            Begin Threed.SSPanel SSPanel9 
               Height          =   285
               Left            =   -70530
               TabIndex        =   105
               Top             =   360
               Width           =   2175
               _Version        =   65536
               _ExtentX        =   3828
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Interes"
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
            Begin MSFlexGridLib.MSFlexGrid grd_CliCon_Listad 
               Height          =   1695
               Left            =   -74970
               TabIndex        =   106
               Top             =   660
               Width           =   11265
               _ExtentX        =   19870
               _ExtentY        =   2990
               _Version        =   393216
               Rows            =   21
               Cols            =   6
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   49152
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel SSPanel10 
               Height          =   285
               Left            =   -74940
               TabIndex        =   107
               Top             =   360
               Width           =   765
               _Version        =   65536
               _ExtentX        =   1349
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Cuota"
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
            Begin Threed.SSPanel SSPanel11 
               Height          =   285
               Left            =   -74190
               TabIndex        =   108
               Top             =   360
               Width           =   1515
               _Version        =   65536
               _ExtentX        =   2672
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "F. Vcto"
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
               Left            =   -72690
               TabIndex        =   109
               Top             =   360
               Width           =   2170
               _Version        =   65536
               _ExtentX        =   3828
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Capital"
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
            Begin Threed.SSPanel SSPanel13 
               Height          =   285
               Left            =   -68370
               TabIndex        =   110
               Top             =   360
               Width           =   2175
               _Version        =   65536
               _ExtentX        =   3828
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Total Cuota"
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
            Begin Threed.SSPanel SSPanel21 
               Height          =   285
               Left            =   -66210
               TabIndex        =   111
               Top             =   360
               Width           =   2235
               _Version        =   65536
               _ExtentX        =   3942
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Saldo Capital"
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
            Begin Threed.SSPanel pnl_CliCon_Intere 
               Height          =   285
               Left            =   -70530
               TabIndex        =   112
               Top             =   2370
               Width           =   2170
               _Version        =   65536
               _ExtentX        =   3828
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
            Begin Threed.SSPanel pnl_CliCon_Capita 
               Height          =   285
               Left            =   -72690
               TabIndex        =   113
               Top             =   2370
               Width           =   2170
               _Version        =   65536
               _ExtentX        =   3828
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
            Begin MSFlexGridLib.MSFlexGrid grd_MViNCo_Listad 
               Height          =   1695
               Left            =   -74970
               TabIndex        =   114
               Top             =   660
               Width           =   11265
               _ExtentX        =   19870
               _ExtentY        =   2990
               _Version        =   393216
               Rows            =   21
               Cols            =   10
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   49152
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel SSPanel8 
               Height          =   285
               Left            =   -71790
               TabIndex        =   115
               Top             =   360
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Interés"
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
            Begin Threed.SSPanel SSPanel38 
               Height          =   285
               Left            =   -74940
               TabIndex        =   116
               Top             =   360
               Width           =   705
               _Version        =   65536
               _ExtentX        =   1244
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Cuota"
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
            Begin Threed.SSPanel SSPanel41 
               Height          =   285
               Left            =   -74250
               TabIndex        =   117
               Top             =   360
               Width           =   1425
               _Version        =   65536
               _ExtentX        =   2514
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "F. Vcto"
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
            Begin Threed.SSPanel SSPanel42 
               Height          =   285
               Left            =   -72840
               TabIndex        =   118
               Top             =   360
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Capital"
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
            Begin Threed.SSPanel SSPanel43 
               Height          =   285
               Left            =   -66390
               TabIndex        =   119
               Top             =   360
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Total Cuota"
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
            Begin Threed.SSPanel SSPanel44 
               Height          =   285
               Left            =   -65310
               TabIndex        =   120
               Top             =   360
               Width           =   1290
               _Version        =   65536
               _ExtentX        =   2275
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Saldo Capital"
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
            Begin Threed.SSPanel SSPanel45 
               Height          =   285
               Left            =   -70710
               TabIndex        =   121
               Top             =   360
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Seg. Prest."
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
            Begin Threed.SSPanel SSPanel46 
               Height          =   285
               Left            =   -69630
               TabIndex        =   122
               Top             =   360
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Seg. Vivienda"
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
            Begin Threed.SSPanel SSPanel47 
               Height          =   285
               Left            =   -68550
               TabIndex        =   123
               Top             =   360
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Portes"
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
            Begin Threed.SSPanel SSPanel49 
               Height          =   285
               Left            =   -67470
               TabIndex        =   124
               Top             =   360
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "C. COFIDE"
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
            Begin Threed.SSPanel pnl_MViNCo_Capita 
               Height          =   285
               Left            =   -72840
               TabIndex        =   125
               Top             =   2370
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
            Begin Threed.SSPanel pnl_MViNCo_SegViv 
               Height          =   285
               Left            =   -69630
               TabIndex        =   126
               Top             =   2370
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
            Begin Threed.SSPanel pnl_MViNCo_SegPre 
               Height          =   285
               Left            =   -70710
               TabIndex        =   127
               Top             =   2370
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
            Begin Threed.SSPanel pnl_MViNCo_Intere 
               Height          =   285
               Left            =   -71790
               TabIndex        =   128
               Top             =   2370
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
            Begin Threed.SSPanel pnl_MViNCo_OtrCar 
               Height          =   285
               Left            =   -68550
               TabIndex        =   129
               Top             =   2370
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
            Begin Threed.SSPanel pnl_MViNCo_TotCuo 
               Height          =   285
               Left            =   -66390
               TabIndex        =   130
               Top             =   2370
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
            Begin Threed.SSPanel pnl_MViNCo_Comisi 
               Height          =   285
               Left            =   -67470
               TabIndex        =   131
               Top             =   2370
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
            Begin Threed.SSPanel pnl_CofNCo_TotCuo 
               Height          =   285
               Left            =   -67500
               TabIndex        =   132
               Top             =   2370
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
            Begin Threed.SSPanel SSPanel55 
               Height          =   285
               Left            =   -70980
               TabIndex        =   133
               Top             =   360
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Interés"
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
            Begin MSFlexGridLib.MSFlexGrid grd_CofNCo_Listad 
               Height          =   1695
               Left            =   -74970
               TabIndex        =   134
               Top             =   660
               Width           =   11265
               _ExtentX        =   19870
               _ExtentY        =   2990
               _Version        =   393216
               Rows            =   21
               Cols            =   7
               FixedRows       =   0
               FixedCols       =   0
               BackColorSel    =   49152
               FocusRect       =   0
               ScrollBars      =   2
               SelectionMode   =   1
            End
            Begin Threed.SSPanel SSPanel56 
               Height          =   285
               Left            =   -74940
               TabIndex        =   135
               Top             =   360
               Width           =   765
               _Version        =   65536
               _ExtentX        =   1349
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Cuota"
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
            Begin Threed.SSPanel SSPanel58 
               Height          =   285
               Left            =   -74220
               TabIndex        =   136
               Top             =   360
               Width           =   1515
               _Version        =   65536
               _ExtentX        =   2672
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "F. Vcto"
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
            Begin Threed.SSPanel SSPanel60 
               Height          =   285
               Left            =   -72720
               TabIndex        =   137
               Top             =   360
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Capital"
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
            Begin Threed.SSPanel SSPanel63 
               Height          =   285
               Left            =   -67500
               TabIndex        =   138
               Top             =   360
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Total Cuota"
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
            Begin Threed.SSPanel SSPanel64 
               Height          =   285
               Left            =   -65760
               TabIndex        =   139
               Top             =   360
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Saldo Capital"
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
            Begin Threed.SSPanel SSPanel65 
               Height          =   285
               Left            =   -69240
               TabIndex        =   140
               Top             =   360
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "Comisión COFIDE"
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
            Begin Threed.SSPanel pnl_CofNCo_Intere 
               Height          =   285
               Left            =   -70980
               TabIndex        =   141
               Top             =   2370
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
            Begin Threed.SSPanel pnl_CofNCo_Capita 
               Height          =   285
               Left            =   -72720
               TabIndex        =   142
               Top             =   2370
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
            Begin Threed.SSPanel pnl_CofNCo_Comisi 
               Height          =   285
               Left            =   -69240
               TabIndex        =   143
               Top             =   2370
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "9,999,999.99 "
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
               Caption         =   "Totales ==>"
               Height          =   285
               Left            =   -72930
               TabIndex        =   66
               Top             =   1470
               Width           =   945
            End
            Begin VB.Label Label14 
               Caption         =   "Totales ==>"
               Height          =   285
               Left            =   -72930
               TabIndex        =   65
               Top             =   1470
               Width           =   945
            End
            Begin VB.Label Label15 
               Caption         =   "Totales ==>"
               Height          =   285
               Left            =   -73230
               TabIndex        =   64
               Top             =   1470
               Width           =   945
            End
         End
      End
   End
End
Attribute VB_Name = "frm_Desemb_05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_CliNCo()         As modcal_g_est_CuoCli
Dim l_arr_CliCon()         As modcal_g_est_CuoCli
Dim l_arr_MViCon()         As modcal_g_est_CuoCli
Dim l_arr_MViNCo()         As modcal_g_est_CuoCli
Dim l_arr_CofNCo()         As modcal_g_est_CuoCli
Dim l_dbl_TipCam_Dol       As Double
Dim l_dbl_TipCam_MPr       As Double
Dim l_arr_BanFia()         As moddat_tpo_Genera
Dim l_arr_BanChq()         As moddat_tpo_Genera
Dim l_arr_CtaChq()         As moddat_tpo_Genera
Dim l_dbl_MtoPre           As Double
Dim l_dbl_IntCap           As Double
Dim l_dbl_MtoCon           As Double
Dim l_dbl_MtoNCo           As Double
Dim l_int_FlgPry           As Integer
Dim l_int_PlaAno           As Integer
Dim l_dbl_ComCRC           As Double
Dim l_dbl_ComPBP           As Double
Dim l_dbl_ComCof           As Double
Dim l_dbl_TasMVi           As Double
Dim l_dbl_TasCof           As Double
Dim l_str_DesCof           As String
Dim l_int_NumCuo           As Integer
Dim l_dbl_TopCon           As Double
Dim l_int_PerGra           As Integer

Private Sub cmb_BanChq_Click()
   Call gs_SetFocus(cmb_CtaChq)
   
   If cmb_BanChq.ListIndex > -1 Then
      Screen.MousePointer = 11
      Call moddat_gs_Carga_CtaBan(l_arr_BanChq(cmb_BanChq.ListIndex + 1).Genera_Codigo, cmb_CtaChq, l_arr_CtaChq)
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmb_BanChq_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_BanChq_Click
   End If
End Sub

Private Sub cmb_BanFia_Click()
   Call gs_SetFocus(txt_NumFia)
End Sub

Private Sub cmb_BanFia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_BanFia_Click
   End If
End Sub

Private Sub cmb_CtaChq_Click()
   Call gs_SetFocus(cmb_FlgFia)
End Sub

Private Sub cmb_CtaChq_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_CtaChq_Click
   End If
End Sub

Private Sub cmb_FlgFia_Click()
   If cmb_FlgFia.ListIndex > -1 Then
      If cmb_FlgFia.ItemData(cmb_FlgFia.ListIndex) = 1 Then
         cmb_BanFia.Enabled = True
         txt_NumFia.Enabled = True
         ipp_FEmFia.Enabled = True
         ipp_FVcFia.Enabled = True
         cmb_MonFia.Enabled = True
         ipp_MtoFia.Enabled = True
         
         Call gs_SetFocus(cmb_BanFia)
      Else
         cmb_BanFia.ListIndex = -1
         txt_NumFia.Text = ""
         ipp_FEmFia.Text = Format(Date, "dd/mm/yyyy")
         ipp_FVcFia.Text = Format(Date, "dd/mm/yyyy")
         cmb_MonFia.ListIndex = -1
         ipp_MtoFia.Value = 0
      
         cmb_BanFia.Enabled = False
         txt_NumFia.Enabled = False
         ipp_FEmFia.Enabled = False
         ipp_FVcFia.Enabled = False
         cmb_MonFia.Enabled = False
         ipp_MtoFia.Enabled = False
         
         Call gs_SetFocus(txt_Observ)
      End If
   End If
End Sub

Private Sub cmb_FlgFia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_FlgFia_Click
   End If
End Sub

Private Sub cmb_MonDes_Click()
   Call gs_SetFocus(ipp_MtoDes)
   Call ipp_MtoDes_Change
End Sub

Private Sub cmb_MonDes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_MonDes_Click
   End If
End Sub

Private Sub cmb_MonFia_Click()
   Call gs_SetFocus(ipp_MtoFia)
End Sub

Private Sub cmb_MonFia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_MonFia_Click
   End If
End Sub

Private Sub cmd_Grabar_Click()
   If cmb_MonDes.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Moneda de Desembolso.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_MonDes)
      Exit Sub
   End If
   
   If ipp_MtoDes.Value = 0 Then
      MsgBox "Debe ingresar el Monto del Desembolso.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_MtoDes)
      Exit Sub
   End If
   
   If l_int_FlgPry = 2 Then
      If Len(Trim(txt_NumChq.Text)) = 0 Then
         MsgBox "Debe ingresar el Número de Cheque.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NumChq)
         Exit Sub
      End If
      
      If cmb_BanChq.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Banco del Cheque.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_BanChq)
         Exit Sub
      End If
   
      If cmb_CtaChq.ListIndex = -1 Then
         MsgBox "Debe seleccionar la Cuenta del Cheque.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_CtaChq)
         Exit Sub
      End If
   
      If cmb_FlgFia.ListIndex = -1 Then
         MsgBox "Debe seleccionar si el Desembolso es contra Carta Fianza.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_FlgFia)
         Exit Sub
      End If
   
      If cmb_FlgFia.ItemData(cmb_FlgFia.ListIndex) = 1 Then
         If cmb_BanFia.ListIndex = -1 Then
            MsgBox "Debe seleccionar el Banco de la Fianza.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmb_BanFia)
            Exit Sub
         End If
      
         If Len(Trim(txt_NumFia.Text)) = 0 Then
            MsgBox "Debe ingresar el Número de Carta Fianza.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(txt_NumFia)
            Exit Sub
         End If
         
         If CDate(ipp_FEmFia.Text) > Date Then
            MsgBox "La Fecha de Emisión de la Carta Fianza no puede ser mayor a la Fecha Actual.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_FEmFia)
            Exit Sub
         End If
         
         If CDate(ipp_FVcFia.Text) < Date Then
            MsgBox "La Fecha de Vencimiento de la Carta Fianza no puede ser menor a la Fecha Actual.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_FVcFia)
            Exit Sub
         End If
         
         If CDate(ipp_FVcFia.Text) < CDate(ipp_FEmFia.Text) Then
            MsgBox "La Fecha de Vencimiento de la Carta Fianza no puede ser menor a la Fecha de Emisión.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_FVcFia)
            Exit Sub
         End If
      
         If cmb_MonFia.ListIndex = -1 Then
            MsgBox "Debe seleccionar la Moneda de la Carta Fianza.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(cmb_MonFia)
            Exit Sub
         End If
         
         If ipp_MtoFia.Value = 0 Then
            MsgBox "Debe seleccionar el Monto de la Carta Fianza.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_MtoFia)
            Exit Sub
         End If
      End If
   End If
   
   If l_dbl_MtoPre <> CDbl(pnl_DesMPr.Caption) Then
      MsgBox "El Monto Desembolsado no coincide con el Monto del Préstamo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_MtoDes)
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If

   Screen.MousePointer = 11

   'Grabar Desembolso
   Call fs_Grabar_Desemb
   DoEvents
   
   'Grabar Cronograma
   Call fs_Grabar_Cronog
   
   'Actualizar Maestro de Créditos
   Call fs_Grabar_Credit
   
   'Grabar Información del Cliente y de Actividad Económica
   Call moddat_gs_Inicia_ActEco(1, 1)
   Call moddat_gs_Inicia_ActEco(1, 2)
   
   DoEvents
   Call moddat_gs_Inicia_ActEco(2, 1)
   Call moddat_gs_Inicia_ActEco(2, 2)
   
   Call fs_Grabar_DatCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc)
   
   DoEvents
   Call fs_Cargar_ActEco_Tit(moddat_g_int_TipDoc, moddat_g_str_NumDoc, 1)
   Call fs_Cargar_ActEco_Tit(moddat_g_int_TipDoc, moddat_g_str_NumDoc, 2)
   Call fs_Grabar_ActEco_Tit(1)

   DoEvents
   If moddat_g_arr_ActEco_Tit(2).ActEco_TipAct > 0 Then
      Call fs_Grabar_ActEco_Tit(2)
   End If
   
   If moddat_g_int_CygTDo > 0 Then
      DoEvents
      Call fs_Grabar_DatCli(moddat_g_int_CygTDo, moddat_g_str_CygNDo)
      
      DoEvents
      Call fs_Cargar_ActEco_Cyg(moddat_g_int_CygTDo, moddat_g_str_CygNDo, 1)
      Call fs_Cargar_ActEco_Cyg(moddat_g_int_CygTDo, moddat_g_str_CygNDo, 2)
   
      If moddat_g_arr_ActEco_Cyg(1).ActEco_TipAct > 0 Then
         Call fs_Grabar_ActEco_Cyg(1)
      End If
   
      If moddat_g_arr_ActEco_Cyg(2).ActEco_TipAct > 0 Then
         Call fs_Grabar_ActEco_Cyg(2)
      End If
   End If
   
   'Creando Nueva Ocurrencia en Detalle de Seguimiento
   If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, 81, 82, 0, "", 0, 0) Then
      Exit Sub
   End If
   
   'Enviando Correo Electrónico
   modgen_g_str_Mail_Asunto = "DESEMBOLSO DE CREDITO (" & Format(CDate(moddat_g_str_FecSis), "dd/mm/yyyy") & " - " & Format(Time, "hh:mm:ss") & ")"
   
   modgen_g_str_Mail_Mensaj = ""
   modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NUMERO DE SOLICITUD : " & pnl_NumSol.Caption & Chr(13)
   modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NUMERO DE OPERACION : " & pnl_NumOpe.Caption & Chr(13)
   modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "ID CLIENTE          : " & CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & Chr(13)
   modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "NOMBRE CLIENTE      : " & moddat_g_str_NomCli & Chr(13)
   modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & Chr(13)
   
   Call fs_Envia_CorEle(modgen_g_str_Mail_Asunto, modgen_g_str_Mail_Mensaj)
   
   Screen.MousePointer = 0
   
   MsgBox "Se desembolso el crédito.", vbInformation, modgen_g_str_NomPlt
   
   moddat_g_int_FlgAct = 2
   
   Unload Me
End Sub

Private Sub cmd_Recalc_Click()
   Dim r_dbl_Capita  As Double
   Dim r_dbl_Intere  As Double
   Dim r_dbl_Comisi  As Double
   Dim r_dbl_SegPre  As Double
   Dim r_dbl_SegViv  As Double
   Dim r_dbl_OtrCar  As Double
   Dim r_dbl_TotCuo  As Double
   Dim r_int_Contad  As Integer
   
   Screen.MousePointer = 11

   If moddat_g_str_CodPrd = "004" Then
      If cmb_TipCal.ListIndex > -1 Then
         Select Case cmb_TipCal.ItemData(cmb_TipCal.ListIndex)
            Case 1:  Call gs_Cronog_Mihogar_NCCof(l_arr_MViNCo(), l_dbl_MtoPre, l_dbl_TopCon, l_int_NumCuo, l_int_PerGra, l_dbl_TasCof, l_dbl_ComCof, l_str_DesCof)
            Case 2:  Call gs_Cronog_Mihogar_NCCof_01(l_arr_MViNCo(), l_dbl_MtoPre, l_dbl_TopCon, l_int_NumCuo, l_int_PerGra, l_dbl_TasCof, l_dbl_ComCof, l_str_DesCof)
            Case 3:  Call gs_Cronog_Mihogar_NCCof_02(l_arr_MViNCo(), l_dbl_MtoPre, l_dbl_TopCon, l_int_NumCuo, l_int_PerGra, l_dbl_TasCof, l_dbl_ComCof, l_str_DesCof)
            Case 4:  Call gs_Cronog_Mihogar_NCCof_03(l_arr_MViNCo(), l_dbl_MtoPre, l_dbl_TopCon, l_int_NumCuo, l_int_PerGra, l_dbl_TasCof, l_dbl_ComCof, l_str_DesCof)
         End Select
      End If
      
      Call gs_LimpiaGrid(grd_MViNCo_Listad)
      
      pnl_MViNCo_Capita.Caption = "0.00 "
      pnl_MViNCo_Intere.Caption = "0.00 "
      pnl_MViNCo_SegPre.Caption = "0.00 "
      pnl_MViNCo_SegViv.Caption = "0.00 "
      pnl_MViNCo_OtrCar.Caption = "0.00 "
      pnl_MViNCo_TotCuo.Caption = "0.00 "
      
      'Mivivienda No Concesional
      r_dbl_Capita = 0
      r_dbl_Intere = 0
      r_dbl_SegPre = 0
      r_dbl_SegViv = 0
      r_dbl_OtrCar = 0
      r_dbl_Comisi = 0
      r_dbl_TotCuo = 0
      
      grd_MViNCo_Listad.Redraw = False
      For r_int_Contad = 1 To UBound(l_arr_MViNCo)
         grd_MViNCo_Listad.Rows = grd_MViNCo_Listad.Rows + 1
         grd_MViNCo_Listad.Row = grd_MViNCo_Listad.Rows - 1
         
         'Número de Cuota
         grd_MViNCo_Listad.Col = 0
         grd_MViNCo_Listad.Text = Format(r_int_Contad, "000")
      
         'Fecha de Vencimiento
         grd_MViNCo_Listad.Col = 1
         grd_MViNCo_Listad.Text = l_arr_MViNCo(r_int_Contad).CuoCli_FecVct
      
         'Capital
         grd_MViNCo_Listad.Col = 2
         grd_MViNCo_Listad.Text = Format(l_arr_MViNCo(r_int_Contad).CuoCli_Capita, "###,###,##0.00")
         r_dbl_Capita = r_dbl_Capita + CDbl(grd_MViNCo_Listad)
         
         'Interes
         grd_MViNCo_Listad.Col = 3
         grd_MViNCo_Listad.Text = Format(l_arr_MViNCo(r_int_Contad).CuoCli_Intere, "###,###,##0.00")
         r_dbl_Intere = r_dbl_Intere + CDbl(grd_MViNCo_Listad)
      
         'Seguro Desgravamen
         grd_MViNCo_Listad.Col = 4
         grd_MViNCo_Listad.Text = Format(l_arr_MViNCo(r_int_Contad).CuoCli_SegPre, "###,###,##0.00")
         r_dbl_SegPre = r_dbl_SegPre + CDbl(grd_MViNCo_Listad)
      
         'Seguro Vivienda
         grd_MViNCo_Listad.Col = 5
         grd_MViNCo_Listad.Text = Format(l_arr_MViNCo(r_int_Contad).CuoCli_SegViv, "###,###,##0.00")
         r_dbl_SegViv = r_dbl_SegViv + CDbl(grd_MViNCo_Listad)
      
         'Otros Cargos
         grd_MViNCo_Listad.Col = 6
         grd_MViNCo_Listad.Text = Format(l_arr_MViNCo(r_int_Contad).CuoCli_Portes, "###,###,##0.00")
         r_dbl_OtrCar = r_dbl_OtrCar + CDbl(grd_MViNCo_Listad)
      
         'Comisión COFIDE
         grd_MViNCo_Listad.Col = 7
         grd_MViNCo_Listad.Text = Format(l_arr_MViNCo(r_int_Contad).CuoCli_Comisi, "###,###,##0.00")
         r_dbl_Comisi = r_dbl_Comisi + CDbl(grd_MViNCo_Listad)
      
         'Valor Cuota
         grd_MViNCo_Listad.Col = 8
         grd_MViNCo_Listad.Text = Format(l_arr_MViNCo(r_int_Contad).CuoCli_ValCuo, "###,###,##0.00")
         r_dbl_TotCuo = r_dbl_TotCuo + CDbl(grd_MViNCo_Listad)
      
         'Saldo Capital
         grd_MViNCo_Listad.Col = 9
         grd_MViNCo_Listad.Text = Format(l_arr_MViNCo(r_int_Contad).CuoCli_SalCap, "###,###,##0.00")
      Next r_int_Contad
      
      grd_MViNCo_Listad.Redraw = True
      
      pnl_MViNCo_Capita.Caption = Format(r_dbl_Capita, "###,###,##0.00") & " "
      pnl_MViNCo_Intere.Caption = Format(r_dbl_Intere, "###,###,##0.00") & " "
      pnl_MViNCo_SegPre.Caption = Format(r_dbl_SegPre, "###,###,##0.00") & " "
      pnl_MViNCo_SegViv.Caption = Format(r_dbl_SegViv, "###,###,##0.00") & " "
      pnl_MViNCo_OtrCar.Caption = Format(r_dbl_OtrCar, "###,###,##0.00") & " "
      pnl_MViNCo_Comisi.Caption = Format(r_dbl_Comisi, "###,###,##0.00") & " "
      pnl_MViNCo_TotCuo.Caption = Format(r_dbl_TotCuo, "###,###,##0.00") & " "
      
      If grd_MViNCo_Listad.Rows > 0 Then
         Call gs_UbiIniGrid(grd_MViNCo_Listad)
      End If
   End If
   
   If moddat_g_str_CodPrd = "003" Then
      If cmb_TipCal.ListIndex > -1 Then
         Select Case cmb_TipCal.ItemData(cmb_TipCal.ListIndex)
            Case 1:  Call gs_Cronog_CME_NCCof(l_arr_CofNCo(), l_dbl_MtoPre, 0, l_int_NumCuo, l_int_PerGra, l_dbl_TasCof, l_dbl_ComCof, l_str_DesCof)
            Case 2:  Call gs_Cronog_CME_NCCof_1(l_arr_CofNCo(), l_dbl_MtoPre, 0, l_int_NumCuo, l_int_PerGra, l_dbl_TasCof, l_dbl_ComCof, l_str_DesCof)
         End Select
      End If
      
      Call gs_LimpiaGrid(grd_CofNCo_Listad)
      
      pnl_CofNCo_Capita.Caption = "0.00 "
      pnl_CofNCo_Intere.Caption = "0.00 "
      pnl_CofNCo_Comisi.Caption = "0.00 "
      pnl_CofNCo_TotCuo.Caption = "0.00 "
      
      'Cofide No Concesional
      r_dbl_Capita = 0
      r_dbl_Intere = 0
      r_dbl_Comisi = 0
      r_dbl_TotCuo = 0
      
      grd_CofNCo_Listad.Redraw = False
      For r_int_Contad = 1 To UBound(l_arr_CofNCo)
         grd_CofNCo_Listad.Rows = grd_CofNCo_Listad.Rows + 1
         grd_CofNCo_Listad.Row = grd_CofNCo_Listad.Rows - 1
         
         'Número de Cuota
         grd_CofNCo_Listad.Col = 0
         grd_CofNCo_Listad.Text = Format(r_int_Contad, "000")
      
         'Fecha de Vencimiento
         grd_CofNCo_Listad.Col = 1
         grd_CofNCo_Listad.Text = l_arr_CofNCo(r_int_Contad).CuoCli_FecVct
      
         'Capital
         grd_CofNCo_Listad.Col = 2
         grd_CofNCo_Listad.Text = Format(l_arr_CofNCo(r_int_Contad).CuoCli_Capita, "###,###,##0.00")
         r_dbl_Capita = r_dbl_Capita + CDbl(grd_CofNCo_Listad)
         
         'Interes
         grd_CofNCo_Listad.Col = 3
         grd_CofNCo_Listad.Text = Format(l_arr_CofNCo(r_int_Contad).CuoCli_Intere, "###,###,##0.00")
         r_dbl_Intere = r_dbl_Intere + CDbl(grd_CofNCo_Listad)
      
         'Comisión
         grd_CofNCo_Listad.Col = 4
         grd_CofNCo_Listad.Text = Format(l_arr_CofNCo(r_int_Contad).CuoCli_Comisi, "###,###,##0.00")
         r_dbl_Comisi = r_dbl_Comisi + CDbl(grd_CofNCo_Listad)
      
         'Valor Cuota
         grd_CofNCo_Listad.Col = 5
         grd_CofNCo_Listad.Text = Format(l_arr_CofNCo(r_int_Contad).CuoCli_ValCuo, "###,###,##0.00")
         r_dbl_TotCuo = r_dbl_TotCuo + CDbl(grd_CofNCo_Listad)
      
         'Saldo Capital
         grd_CofNCo_Listad.Col = 6
         grd_CofNCo_Listad.Text = Format(l_arr_CofNCo(r_int_Contad).CuoCli_SalCap, "###,###,##0.00")
      Next r_int_Contad
      
      grd_CofNCo_Listad.Redraw = True
      
      pnl_CofNCo_Capita.Caption = Format(r_dbl_Capita, "###,###,##0.00") & " "
      pnl_CofNCo_Intere.Caption = Format(r_dbl_Intere, "###,###,##0.00") & " "
      pnl_CofNCo_Comisi.Caption = Format(r_dbl_Comisi, "###,###,##0.00") & " "
      pnl_CofNCo_TotCuo.Caption = Format(r_dbl_TotCuo, "###,###,##0.00") & " "
      
      If grd_CofNCo_Listad.Rows > 0 Then
         Call gs_UbiIniGrid(grd_CofNCo_Listad)
      End If
   End If
   
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Dim r_dbl_TasInt  As Double
   Dim r_int_CuoExt  As Integer
   Dim r_dbl_FoIDes  As Double
   Dim r_int_AplViv  As Integer
   Dim r_dbl_FoIViv  As Double
   Dim r_dbl_Portes  As Double
   Dim r_int_DiaPag  As Integer
   Dim r_dbl_MtoViv  As Double
   Dim r_dbl_PorCon  As Double
   Dim r_dbl_TCaSbs  As Double
   
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt

   pnl_NumSol.Caption = Mid(moddat_g_str_NumSol, 1, 3) & "-" & Mid(moddat_g_str_NumSol, 4, 3) & "-" & Mid(moddat_g_str_NumSol, 7, 2) & "-" & Mid(moddat_g_str_NumSol, 9, 4)
   pnl_NumOpe.Caption = Mid(moddat_g_str_NumOpe, 1, 3) & "-" & Mid(moddat_g_str_NumOpe, 4, 2) & "-" & Mid(moddat_g_str_NumOpe, 6, 5)
   pnl_Produc.Caption = moddat_g_str_NomPrd
   pnl_GenOpe.Caption = moddat_g_str_FecIng
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & " - " & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   
   Call fs_Inicia
   Call fs_Limpia
   
   'Obteniendo Tipo de Cambio
   l_dbl_TipCam_Dol = 0
   l_dbl_TipCam_Dol = moddat_gf_Obtiene_TipCam(1, 2)
   pnl_TCaDol.Caption = Format(l_dbl_TipCam_Dol, "###,##0.0000") & " "
   
   l_dbl_TipCam_MPr = 0
   If moddat_g_int_TipMon <> 1 Then
      l_dbl_TipCam_MPr = moddat_gf_Obtiene_TipCam(1, moddat_g_int_TipMon)
      pnl_TCaMPr.Caption = Format(l_dbl_TipCam_MPr, "###,##0.0000") & " "
   End If
   
   'Obteniendo Datos del Préstamo
   l_dbl_MtoPre = 0
   l_dbl_IntCap = 0
   l_dbl_MtoNCo = 0
   l_dbl_MtoCon = 0
   l_int_PlaAno = 0
   r_dbl_TasInt = 0
   l_int_NumCuo = 0
   r_int_CuoExt = 0
   r_dbl_FoIDes = 0
   r_int_AplViv = 0
   r_dbl_FoIViv = 0
   r_dbl_Portes = 0
   r_int_DiaPag = 0
   l_int_PerGra = 0
   r_dbl_TCaSbs = 0
   
   g_str_Parame = "SELECT * FROM CRE_HIPMAE WHERE "
   g_str_Parame = g_str_Parame & "HIPMAE_NUMOPE = '" & moddat_g_str_NumOpe & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      l_dbl_MtoPre = g_rst_Princi!HIPMAE_MTOPRE
      r_dbl_TasInt = g_rst_Princi!HIPMAE_TASINT
      l_int_PlaAno = g_rst_Princi!HIPMAE_PLAANO
      l_int_NumCuo = g_rst_Princi!HIPMAE_NUMCUO
      r_int_CuoExt = g_rst_Princi!HIPMAE_CUOANO
      r_dbl_FoIDes = g_rst_Princi!HIPMAE_FOIPRE
      r_int_AplViv = g_rst_Princi!HIPMAE_APLVIV
      r_dbl_FoIViv = g_rst_Princi!HIPMAE_FOIVIV
      r_dbl_Portes = g_rst_Princi!HIPMAE_OTRIMP
      r_int_DiaPag = g_rst_Princi!HIPMAE_DIAPAG
      l_int_PerGra = g_rst_Princi!HIPMAE_PERGRA
      r_dbl_TCaSbs = g_rst_Princi!HIPMAE_TCASBS
      
      l_int_FlgPry = g_rst_Princi!HIPMAE_PRYMCS
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Fecha de Desembolso COFIDE
   If moddat_g_str_CodPrd = "004" Or moddat_g_str_CodPrd = "003" Then
      g_str_Parame = "SELECT * FROM TRA_EVACOF WHERE "
      g_str_Parame = g_str_Parame & "EVACOF_NUMSOL = '" & moddat_g_str_NumSol & "' "
   
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
          Exit Sub
      End If
      
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         g_rst_Princi.MoveFirst
         
         l_str_DesCof = gf_FormatoFecha(CStr(g_rst_Princi!EVACOF_FECDES))
      End If
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   End If
   
   'Obteniendo Parámetros Mivivienda
   l_dbl_ComCRC = 0
   l_dbl_ComPBP = 0
   l_dbl_TasCof = 0
   l_dbl_ComCof = 0
   l_dbl_TasMVi = 0

   'Comisiones y Tasas Mivivienda / COFIDE
   '1 - Comisión CRC
   '2 - Comisión PBP
   '3 - Tasa de Interés TC
   '4 - Comisión COFIDE
   '5 - Tasa COFIDE

   If moddat_g_str_CodPrd = "001" Then
      l_dbl_ComCRC = moddat_gf_ComMVi(moddat_g_str_CodPrd, 1, moddat_g_int_TipMon, l_int_PlaAno)
      l_dbl_ComPBP = moddat_gf_ComMVi(moddat_g_str_CodPrd, 2, moddat_g_int_TipMon, l_int_PlaAno)
      l_dbl_TasMVi = moddat_gf_ComMVi(moddat_g_str_CodPrd, 3, moddat_g_int_TipMon, l_int_PlaAno)
   ElseIf moddat_g_str_CodPrd = "003" Then
      l_dbl_ComCRC = moddat_gf_ComMVi(moddat_g_str_CodPrd, 1, moddat_g_int_TipMon, l_int_PlaAno)
      l_dbl_ComPBP = moddat_gf_ComMVi(moddat_g_str_CodPrd, 2, moddat_g_int_TipMon, l_int_PlaAno)
      l_dbl_TasMVi = moddat_gf_ComMVi(moddat_g_str_CodPrd, 3, moddat_g_int_TipMon, l_int_PlaAno)
      l_dbl_ComCof = moddat_gf_ComMVi(moddat_g_str_CodPrd, 4, moddat_g_int_TipMon, l_int_PlaAno)
      l_dbl_TasCof = moddat_gf_ComMVi(moddat_g_str_CodPrd, 5, moddat_g_int_TipMon, l_int_PlaAno)
   ElseIf moddat_g_str_CodPrd = "004" Then
      l_dbl_ComCof = moddat_gf_ComMVi(moddat_g_str_CodPrd, 4, moddat_g_int_TipMon, l_int_PlaAno)
      l_dbl_TasCof = moddat_gf_ComMVi(moddat_g_str_CodPrd, 5, moddat_g_int_TipMon, l_int_PlaAno)
   End If
   
   'Obteniendo Valor de Inmueble
   r_dbl_MtoViv = 0
   g_str_Parame = "SELECT * FROM TRA_EVATAS WHERE "
   g_str_Parame = g_str_Parame & "EVATAS_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      r_dbl_MtoViv = g_rst_Princi!EVATAS_SUMASE_INM + g_rst_Princi!EVATAS_SUMASE_ES1 + g_rst_Princi!EVATAS_SUMASE_ES2 + g_rst_Princi!EVATAS_SUMASE_DEP
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Generando Cronogramas de Pago
   Select Case moddat_g_str_CodPrd
      Case "001"
         tab_Cronog.TabCaption(0) = "Cliente - No Concesional"
         tab_Cronog.TabCaption(1) = "Cliente - Concesional"
         tab_Cronog.TabCaption(2) = "Mivivienda - No Concesional"
         tab_Cronog.TabCaption(3) = "Mivivienda - Concesional"
         
      Case "002"
         tab_Cronog.TabCaption(0) = "Cliente"
   
      Case "003"
         tab_Cronog.TabCaption(0) = "Cliente - No Concesional"
         tab_Cronog.TabCaption(1) = "Cliente - Concesional"
         tab_Cronog.TabCaption(2) = "Mivivienda - No Concesional"
         tab_Cronog.TabCaption(3) = "Mivivienda - Concesional"
         tab_Cronog.TabCaption(4) = "Cofide"
         
      Case "004"
         tab_Cronog.TabCaption(0) = "Cliente - No Concesional"
         tab_Cronog.TabCaption(1) = "Cliente - Concesional"
         tab_Cronog.TabCaption(2) = "Cofide - No Concesional"
         tab_Cronog.TabCaption(3) = "Cofide - Concesional"
   End Select
   
   If moddat_g_str_CodPrd = "002" Then
      tab_Cronog.TabVisible(1) = False
      tab_Cronog.TabVisible(2) = False
      tab_Cronog.TabVisible(3) = False
      tab_Cronog.TabVisible(4) = False
      
      'Calculando Cronograma de Pagos
      Call gs_Cronog_MiCasita(l_arr_CliNCo(), r_dbl_MtoViv, l_dbl_MtoPre, l_int_NumCuo, r_int_CuoExt, r_dbl_TasInt, r_dbl_FoIDes, r_int_AplViv, r_dbl_FoIViv, r_dbl_Portes, Format(Date, "dd/mm/yyyy"), r_int_DiaPag, l_int_PerGra, l_dbl_IntCap, 1)
   Else
      r_dbl_PorCon = 0
      l_dbl_TopCon = 0

      If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "012") Then
         l_dbl_TopCon = moddat_g_arr_Genera(1).Genera_Cantid
      End If
   
      If moddat_g_str_CodPrd = "001" Then
         If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "011") Then
            r_dbl_PorCon = moddat_g_arr_Genera(1).Genera_Cantid
         End If
      
         tab_Cronog.TabVisible(1) = True
         tab_Cronog.TabVisible(2) = True
         tab_Cronog.TabVisible(3) = True
         tab_Cronog.TabVisible(4) = False
      
         Call gs_Cronog_CRCPBP_NC(l_arr_CliNCo(), l_dbl_MtoPre, r_dbl_PorCon, l_dbl_TopCon, r_dbl_TCaSbs, r_dbl_MtoViv, l_int_NumCuo, l_int_PerGra, r_dbl_TasInt, r_dbl_FoIDes, r_int_AplViv, r_dbl_FoIViv, r_dbl_Portes, Format(Date, "dd/mm/yyyy"), r_int_DiaPag, l_dbl_MtoNCo, l_dbl_MtoCon, l_dbl_IntCap, 1)
         Call gs_Cronog_CRCPBP_NCMVI(l_arr_MViNCo(), l_dbl_MtoPre, r_dbl_PorCon, l_dbl_TopCon, r_dbl_TCaSbs, r_dbl_MtoViv, l_int_NumCuo, l_int_PerGra, r_dbl_TasInt, r_dbl_FoIDes, r_int_AplViv, r_dbl_FoIViv, r_dbl_Portes, Format(Date, "dd/mm/yyyy"), r_int_DiaPag, l_dbl_MtoNCo, l_dbl_MtoCon, l_dbl_IntCap, 1)
         Call gs_Cronog_CRCPBP_ConMVi(l_arr_MViCon(), l_dbl_MtoCon, l_int_NumCuo / 6, l_int_PerGra, l_dbl_TasMVi, Format(Date, "dd/mm/yyyy"), r_int_DiaPag)
         Call gs_Cronog_CRCPBP_ConCli(l_arr_MViCon(), l_arr_CliCon(), l_int_NumCuo / 6, r_dbl_TasInt)
         
      ElseIf moddat_g_str_CodPrd = "003" Then
         If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "011") Then
            r_dbl_PorCon = moddat_g_arr_Genera(1).Genera_Cantid
         End If
      
         tab_Cronog.TabVisible(1) = True
         tab_Cronog.TabVisible(2) = True
         tab_Cronog.TabVisible(3) = True
         tab_Cronog.TabVisible(4) = True
      
         Call gs_Cronog_CME_NC(l_arr_CliNCo(), l_dbl_MtoPre, r_dbl_PorCon, l_dbl_TopCon, r_dbl_MtoViv, l_int_NumCuo, l_int_PerGra, r_dbl_TasInt, r_dbl_FoIDes, r_int_AplViv, r_dbl_FoIViv, r_dbl_Portes, Format(Date, "dd/mm/yyyy"), r_int_DiaPag, l_dbl_MtoNCo, l_dbl_MtoCon, l_dbl_IntCap, 1)
         Call gs_Cronog_CME_NCMVI(l_arr_MViNCo(), l_dbl_MtoPre, r_dbl_PorCon, l_dbl_TopCon, r_dbl_MtoViv, l_int_NumCuo, l_int_PerGra, r_dbl_TasInt, r_dbl_FoIDes, r_int_AplViv, r_dbl_FoIViv, r_dbl_Portes, Format(Date, "dd/mm/yyyy"), r_int_DiaPag, l_dbl_MtoNCo, l_dbl_MtoCon, l_dbl_IntCap, 1)
         Call gs_Cronog_CME_ConMVi(l_arr_MViCon(), l_dbl_MtoCon, l_int_NumCuo / 6, l_int_PerGra, l_dbl_TasMVi, Format(Date, "dd/mm/yyyy"), r_int_DiaPag)
         Call gs_Cronog_CME_ConCli(l_arr_MViCon(), l_arr_CliCon(), l_int_NumCuo / 6, r_dbl_TasInt)
         Call gs_Cronog_CME_NCCof(l_arr_CofNCo(), l_dbl_MtoPre, 0, l_int_NumCuo, l_int_PerGra, l_dbl_TasCof, l_dbl_ComCof, l_str_DesCof)
         
      ElseIf moddat_g_str_CodPrd = "004" Then
         tab_Cronog.TabVisible(1) = True
         tab_Cronog.TabVisible(2) = True
         tab_Cronog.TabVisible(3) = True
         tab_Cronog.TabVisible(4) = False
      
         Call gs_Cronog_Mihogar_NC(l_arr_CliNCo(), l_dbl_MtoPre, l_dbl_TopCon, r_dbl_MtoViv, l_int_NumCuo, l_int_PerGra, r_dbl_TasInt, r_dbl_FoIDes, r_int_AplViv, r_dbl_FoIViv, r_dbl_Portes, Format(Date, "dd/mm/yyyy"), r_int_DiaPag, l_dbl_MtoNCo, l_dbl_MtoCon, l_dbl_IntCap, 1)
         Call gs_Cronog_Mihogar_NCCof(l_arr_MViNCo(), l_dbl_MtoPre, l_dbl_TopCon, l_int_NumCuo, l_int_PerGra, l_dbl_TasCof, l_dbl_ComCof, l_str_DesCof)
         
         Call gs_Cronog_Mihogar_ConCof(l_arr_MViCon(), l_dbl_TopCon, l_int_NumCuo / 6, l_int_PerGra / 6, l_dbl_TasCof, l_dbl_ComCof, l_str_DesCof)
         Call gs_Cronog_Mihogar_Con(l_arr_CliCon(), l_arr_MViCon(), l_int_NumCuo / 6, l_int_PerGra / 6, r_dbl_TasInt, Format(Date, "dd/mm/yyyy"), r_int_DiaPag)
      End If
   End If
   
   Call fs_GenCro

   If l_int_FlgPry = 1 Then
      txt_NumChq.Enabled = False
      cmb_BanChq.Enabled = False
      cmb_CtaChq.Enabled = False
      cmb_FlgFia.Enabled = False
      cmb_BanFia.Enabled = False
      txt_NumFia.Enabled = False
      ipp_FEmFia.Enabled = False
      ipp_FVcFia.Enabled = False
      cmb_MonFia.Enabled = False
      ipp_MtoFia.Enabled = False
   End If

   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_LisIte(cmb_BanChq, l_arr_BanChq, 1, "516")
   Call moddat_gs_Carga_LisIte(cmb_BanFia, l_arr_BanFia, 1, "505")

   Call moddat_gs_Carga_LisIte_Combo(cmb_MonDes, 1, "204")
   Call moddat_gs_Carga_LisIte_Combo(cmb_MonFia, 1, "204")
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_FlgFia, 1, "214")

   cmb_TipCal.Clear
   
   cmb_TipCal.AddItem "Forma 1"
   cmb_TipCal.ItemData(cmb_TipCal.NewIndex) = 1
   
   cmb_TipCal.AddItem "Forma 2"
   cmb_TipCal.ItemData(cmb_TipCal.NewIndex) = 2

   cmb_TipCal.AddItem "Forma 3"
   cmb_TipCal.ItemData(cmb_TipCal.NewIndex) = 3

   cmb_TipCal.AddItem "Forma 4"
   cmb_TipCal.ItemData(cmb_TipCal.NewIndex) = 4

   'Cliente No Concesional
   grd_CliNCo_Listad.ColWidth(0) = 795
   grd_CliNCo_Listad.ColWidth(1) = 1425
   grd_CliNCo_Listad.ColWidth(2) = 1180
   grd_CliNCo_Listad.ColWidth(3) = 1170
   grd_CliNCo_Listad.ColWidth(4) = 1160
   grd_CliNCo_Listad.ColWidth(5) = 1160
   grd_CliNCo_Listad.ColWidth(6) = 1160
   grd_CliNCo_Listad.ColWidth(7) = 1320
   grd_CliNCo_Listad.ColWidth(8) = 1560
   
   grd_CliNCo_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_CliNCo_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_CliNCo_Listad.ColAlignment(2) = flexAlignRightCenter
   grd_CliNCo_Listad.ColAlignment(3) = flexAlignRightCenter
   grd_CliNCo_Listad.ColAlignment(4) = flexAlignRightCenter
   grd_CliNCo_Listad.ColAlignment(5) = flexAlignRightCenter
   grd_CliNCo_Listad.ColAlignment(6) = flexAlignRightCenter
   grd_CliNCo_Listad.ColAlignment(7) = flexAlignRightCenter
   grd_CliNCo_Listad.ColAlignment(8) = flexAlignRightCenter

   'Mivivienda No Concesional
   grd_MViNCo_Listad.ColWidth(0) = 695
   grd_MViNCo_Listad.ColWidth(1) = 1415
   grd_MViNCo_Listad.ColWidth(2) = 1070
   grd_MViNCo_Listad.ColWidth(3) = 1070
   grd_MViNCo_Listad.ColWidth(4) = 1080
   grd_MViNCo_Listad.ColWidth(5) = 1080
   grd_MViNCo_Listad.ColWidth(6) = 1080
   grd_MViNCo_Listad.ColWidth(7) = 1080
   grd_MViNCo_Listad.ColWidth(8) = 1080
   grd_MViNCo_Listad.ColWidth(9) = 1290
   
   grd_MViNCo_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_MViNCo_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_MViNCo_Listad.ColAlignment(2) = flexAlignRightCenter
   grd_MViNCo_Listad.ColAlignment(3) = flexAlignRightCenter
   grd_MViNCo_Listad.ColAlignment(4) = flexAlignRightCenter
   grd_MViNCo_Listad.ColAlignment(5) = flexAlignRightCenter
   grd_MViNCo_Listad.ColAlignment(6) = flexAlignRightCenter
   grd_MViNCo_Listad.ColAlignment(7) = flexAlignRightCenter
   grd_MViNCo_Listad.ColAlignment(8) = flexAlignRightCenter
   grd_MViNCo_Listad.ColAlignment(9) = flexAlignRightCenter

   'Mivivienda Concesional
   grd_MviCon_Listad.ColWidth(0) = 770
   grd_MviCon_Listad.ColWidth(1) = 1485
   grd_MviCon_Listad.ColWidth(2) = 1730
   grd_MviCon_Listad.ColWidth(3) = 1740
   grd_MviCon_Listad.ColWidth(4) = 1740
   grd_MviCon_Listad.ColWidth(5) = 1740
   grd_MviCon_Listad.ColWidth(6) = 1740
   
   grd_MviCon_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_MviCon_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_MviCon_Listad.ColAlignment(2) = flexAlignRightCenter
   grd_MviCon_Listad.ColAlignment(3) = flexAlignRightCenter
   grd_MviCon_Listad.ColAlignment(4) = flexAlignRightCenter
   grd_MviCon_Listad.ColAlignment(5) = flexAlignRightCenter
   grd_MviCon_Listad.ColAlignment(6) = flexAlignRightCenter

   'Cliente Concesional
   grd_CliCon_Listad.ColWidth(0) = 770
   grd_CliCon_Listad.ColWidth(1) = 1485
   grd_CliCon_Listad.ColWidth(2) = 2170
   grd_CliCon_Listad.ColWidth(3) = 2160
   grd_CliCon_Listad.ColWidth(4) = 2170
   grd_CliCon_Listad.ColWidth(5) = 2170
   
   grd_CliCon_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_CliCon_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_CliCon_Listad.ColAlignment(2) = flexAlignRightCenter
   grd_CliCon_Listad.ColAlignment(3) = flexAlignRightCenter
   grd_CliCon_Listad.ColAlignment(4) = flexAlignRightCenter
   grd_CliCon_Listad.ColAlignment(5) = flexAlignRightCenter

   'Cofide No Concesional
   grd_CofNCo_Listad.ColWidth(0) = 770
   grd_CofNCo_Listad.ColWidth(1) = 1485
   grd_CofNCo_Listad.ColWidth(2) = 1730
   grd_CofNCo_Listad.ColWidth(3) = 1740
   grd_CofNCo_Listad.ColWidth(4) = 1740
   grd_CofNCo_Listad.ColWidth(5) = 1740
   grd_CofNCo_Listad.ColWidth(6) = 1740
   
   grd_CofNCo_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_CofNCo_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_CofNCo_Listad.ColAlignment(2) = flexAlignRightCenter
   grd_CofNCo_Listad.ColAlignment(3) = flexAlignRightCenter
   grd_CofNCo_Listad.ColAlignment(4) = flexAlignRightCenter
   grd_CofNCo_Listad.ColAlignment(5) = flexAlignRightCenter
   grd_CofNCo_Listad.ColAlignment(6) = flexAlignRightCenter
End Sub

Private Sub fs_Limpia()
   Call gs_LimpiaGrid(grd_CliNCo_Listad)
   Call gs_LimpiaGrid(grd_CliCon_Listad)
   Call gs_LimpiaGrid(grd_MViNCo_Listad)
   Call gs_LimpiaGrid(grd_MviCon_Listad)
   Call gs_LimpiaGrid(grd_CofNCo_Listad)
   
   pnl_CliNCo_Capita.Caption = "0.00 "
   pnl_CliNCo_Intere.Caption = "0.00 "
   pnl_CliNCo_SegPre.Caption = "0.00 "
   pnl_CliNCo_SegViv.Caption = "0.00 "
   pnl_CliNCo_OtrCar.Caption = "0.00 "
   pnl_CliNCo_TotCuo.Caption = "0.00 "
   
   pnl_MViNCo_Capita.Caption = "0.00 "
   pnl_MViNCo_Intere.Caption = "0.00 "
   pnl_MViNCo_SegPre.Caption = "0.00 "
   pnl_MViNCo_SegViv.Caption = "0.00 "
   pnl_MViNCo_OtrCar.Caption = "0.00 "
   pnl_MViNCo_TotCuo.Caption = "0.00 "
   
   pnl_CofNCo_Capita.Caption = "0.00 "
   pnl_CofNCo_Intere.Caption = "0.00 "
   pnl_CofNCo_Comisi.Caption = "0.00 "
   pnl_CofNCo_TotCuo.Caption = "0.00 "
   
   pnl_CliCon_Capita.Caption = "0.00 "
   pnl_CliCon_Intere.Caption = "0.00 "
   pnl_CliCon_TotCuo.Caption = "0.00 "
   
   pnl_MViCon_Capita.Caption = "0.00 "
   pnl_MViCon_Intere.Caption = "0.00 "
   pnl_MViCon_Comisi.Caption = "0.00 "
   pnl_MViCon_TotCuo.Caption = "0.00 "
   
   pnl_DesMPr.Caption = "0.00 "
   pnl_TCaDol.Caption = "0.0000 "
   pnl_TCaMPr.Caption = "0.0000 "
   
   cmb_MonDes.ListIndex = -1
   ipp_MtoDes.Value = 0
   txt_NumChq.Text = ""
   cmb_BanChq.ListIndex = -1
   cmb_CtaChq.Clear
   
   cmb_FlgFia.ListIndex = -1
   
   cmb_BanFia.ListIndex = -1
   txt_NumFia.Text = ""
   ipp_FEmFia.Text = Format(Date, "dd/mm/yyyy")
   ipp_FVcFia.Text = Format(Date, "dd/mm/yyyy")
   cmb_MonFia.ListIndex = -1
   ipp_MtoFia.Value = 0
   cmb_BanFia.Enabled = False
   txt_NumFia.Enabled = False
   ipp_FEmFia.Enabled = False
   ipp_FVcFia.Enabled = False
   cmb_MonFia.Enabled = False
   ipp_MtoFia.Enabled = False
   
   txt_Observ.Text = ""
End Sub

Private Sub grd_CliCon_Listad_SelChange()
   If grd_CliCon_Listad.Rows > 2 Then
      grd_CliCon_Listad.RowSel = grd_CliCon_Listad.Row
   End If
End Sub

Private Sub grd_CliNCo_Listad_SelChange()
   If grd_CliNCo_Listad.Rows > 2 Then
      grd_CliNCo_Listad.RowSel = grd_CliNCo_Listad.Row
   End If
End Sub

Private Sub grd_CofNCo_Listad_SelChange()
   If grd_CofNCo_Listad.Rows > 2 Then
      grd_CofNCo_Listad.RowSel = grd_CofNCo_Listad.Row
   End If
End Sub

Private Sub grd_MviCon_Listad_SelChange()
   If grd_MviCon_Listad.Rows > 2 Then
      grd_MviCon_Listad.RowSel = grd_MviCon_Listad.Row
   End If
End Sub

Private Sub grd_MViNCo_Listad_SelChange()
   If grd_MViNCo_Listad.Rows > 2 Then
      grd_MViNCo_Listad.RowSel = grd_MViNCo_Listad.Row
   End If
End Sub

Private Sub ipp_FEmFia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FVcFia)
   End If
End Sub

Private Sub ipp_FVcFia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_MonFia)
   End If
End Sub

Private Sub ipp_MtoDes_Change()
   If moddat_g_int_TipMon = 1 Then
      If cmb_MonDes.ListIndex > -1 Then
         If cmb_MonDes.ItemData(cmb_MonDes.ListIndex) = 1 Then
            pnl_DesMPr.Caption = Format(ipp_MtoDes.Value, "###,###,##0.00") & " "
         ElseIf cmb_MonDes.ItemData(cmb_MonDes.ListIndex) = 2 Then
            pnl_DesMPr.Caption = Format(ipp_MtoDes.Value * l_dbl_TipCam_Dol, "###,###,##0.00") & " "
         End If
      End If
   Else
      If cmb_MonDes.ListIndex > -1 Then
         If cmb_MonDes.ItemData(cmb_MonDes.ListIndex) = 1 Then
            pnl_DesMPr.Caption = Format(ipp_MtoDes.Value / l_dbl_TipCam_MPr, "###,###,##0.00") & " "
         ElseIf cmb_MonDes.ItemData(cmb_MonDes.ListIndex) = 2 Then
            pnl_DesMPr.Caption = Format(ipp_MtoDes.Value * l_dbl_TipCam_Dol / l_dbl_TipCam_MPr, "###,###,##0.00") & " "
         End If
      End If
   End If
End Sub

Private Sub ipp_MtoDes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NumChq)
   End If
End Sub

Private Sub ipp_MtoFia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Observ)
   End If
End Sub

Private Sub txt_NumChq_GotFocus()
   Call gs_SelecTodo(txt_NumChq)
End Sub

Private Sub txt_NumChq_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_BanChq)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "-")
   End If
End Sub

Private Sub txt_NumFia_GotFocus()
   Call gs_SelecTodo(txt_NumFia)
End Sub

Private Sub txt_NumFia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FEmFia)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "._-")
   End If
End Sub

Private Sub txt_Observ_GotFocus()
   Call gs_SelecTodo(txt_Observ)
End Sub

Private Sub txt_Observ_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_., ;:()/&%$·!ª@#=?¿+*" & Chr(10))
   End If
End Sub

Private Sub fs_GenCro()
   Dim r_dbl_Capita  As Double
   Dim r_dbl_Intere  As Double
   Dim r_dbl_Comisi  As Double
   Dim r_dbl_SegPre  As Double
   Dim r_dbl_SegViv  As Double
   Dim r_dbl_OtrCar  As Double
   Dim r_dbl_TotCuo  As Double
   Dim r_int_Contad  As Integer
   
   'Generando Cronogramas de Pago
   
   'Cliente No Concesional
   r_dbl_Capita = 0
   r_dbl_Intere = 0
   r_dbl_SegPre = 0
   r_dbl_SegViv = 0
   r_dbl_OtrCar = 0
   r_dbl_TotCuo = 0
   
   grd_CliNCo_Listad.Redraw = False
   For r_int_Contad = 1 To UBound(l_arr_CliNCo)
      grd_CliNCo_Listad.Rows = grd_CliNCo_Listad.Rows + 1
      grd_CliNCo_Listad.Row = grd_CliNCo_Listad.Rows - 1
      
      'Número de Cuota
      grd_CliNCo_Listad.Col = 0
      grd_CliNCo_Listad.Text = Format(r_int_Contad, "000")
   
      'Fecha de Vencimiento
      grd_CliNCo_Listad.Col = 1
      grd_CliNCo_Listad.Text = l_arr_CliNCo(r_int_Contad).CuoCli_FecVct
   
      'Capital
      grd_CliNCo_Listad.Col = 2
      grd_CliNCo_Listad.Text = Format(l_arr_CliNCo(r_int_Contad).CuoCli_Capita, "###,###,##0.00")
      r_dbl_Capita = r_dbl_Capita + CDbl(grd_CliNCo_Listad)
      
      'Interes
      grd_CliNCo_Listad.Col = 3
      grd_CliNCo_Listad.Text = Format(l_arr_CliNCo(r_int_Contad).CuoCli_Intere, "###,###,##0.00")
      r_dbl_Intere = r_dbl_Intere + CDbl(grd_CliNCo_Listad)
   
      'Seguro Desgravamen
      grd_CliNCo_Listad.Col = 4
      grd_CliNCo_Listad.Text = Format(l_arr_CliNCo(r_int_Contad).CuoCli_SegPre, "###,###,##0.00")
      r_dbl_SegPre = r_dbl_SegPre + CDbl(grd_CliNCo_Listad)
   
      'Seguro Vivienda
      grd_CliNCo_Listad.Col = 5
      grd_CliNCo_Listad.Text = Format(l_arr_CliNCo(r_int_Contad).CuoCli_SegViv, "###,###,##0.00")
      r_dbl_SegViv = r_dbl_SegViv + CDbl(grd_CliNCo_Listad)
   
      'Otros Cargos
      grd_CliNCo_Listad.Col = 6
      grd_CliNCo_Listad.Text = Format(l_arr_CliNCo(r_int_Contad).CuoCli_Portes, "###,###,##0.00")
      r_dbl_OtrCar = r_dbl_OtrCar + CDbl(grd_CliNCo_Listad)
   
      'Valor Cuota
      grd_CliNCo_Listad.Col = 7
      grd_CliNCo_Listad.Text = Format(l_arr_CliNCo(r_int_Contad).CuoCli_ValCuo, "###,###,##0.00")
      r_dbl_TotCuo = r_dbl_TotCuo + CDbl(grd_CliNCo_Listad)
   
      'Saldo Capital
      grd_CliNCo_Listad.Col = 8
      grd_CliNCo_Listad.Text = Format(l_arr_CliNCo(r_int_Contad).CuoCli_SalCap, "###,###,##0.00")
   Next r_int_Contad
   
   grd_CliNCo_Listad.Redraw = True
   
   pnl_CliNCo_Capita.Caption = Format(r_dbl_Capita, "###,###,##0.00") & " "
   pnl_CliNCo_Intere.Caption = Format(r_dbl_Intere, "###,###,##0.00") & " "
   pnl_CliNCo_SegPre.Caption = Format(r_dbl_SegPre, "###,###,##0.00") & " "
   pnl_CliNCo_SegViv.Caption = Format(r_dbl_SegViv, "###,###,##0.00") & " "
   pnl_CliNCo_OtrCar.Caption = Format(r_dbl_OtrCar, "###,###,##0.00") & " "
   pnl_CliNCo_TotCuo.Caption = Format(r_dbl_TotCuo, "###,###,##0.00") & " "
   
   If grd_CliNCo_Listad.Rows > 0 Then
      Call gs_UbiIniGrid(grd_CliNCo_Listad)
   End If


   If moddat_g_str_CodPrd <> "002" Then
      'Mivivienda No Concesional
      r_dbl_Capita = 0
      r_dbl_Intere = 0
      r_dbl_SegPre = 0
      r_dbl_SegViv = 0
      r_dbl_OtrCar = 0
      r_dbl_Comisi = 0
      r_dbl_TotCuo = 0
      
      grd_MViNCo_Listad.Redraw = False
      For r_int_Contad = 1 To UBound(l_arr_MViNCo)
         grd_MViNCo_Listad.Rows = grd_MViNCo_Listad.Rows + 1
         grd_MViNCo_Listad.Row = grd_MViNCo_Listad.Rows - 1
         
         'Número de Cuota
         grd_MViNCo_Listad.Col = 0
         grd_MViNCo_Listad.Text = Format(r_int_Contad, "000")
      
         'Fecha de Vencimiento
         grd_MViNCo_Listad.Col = 1
         grd_MViNCo_Listad.Text = l_arr_MViNCo(r_int_Contad).CuoCli_FecVct
      
         'Capital
         grd_MViNCo_Listad.Col = 2
         grd_MViNCo_Listad.Text = Format(l_arr_MViNCo(r_int_Contad).CuoCli_Capita, "###,###,##0.00")
         r_dbl_Capita = r_dbl_Capita + CDbl(grd_MViNCo_Listad)
         
         'Interes
         grd_MViNCo_Listad.Col = 3
         grd_MViNCo_Listad.Text = Format(l_arr_MViNCo(r_int_Contad).CuoCli_Intere, "###,###,##0.00")
         r_dbl_Intere = r_dbl_Intere + CDbl(grd_MViNCo_Listad)
      
         'Seguro Desgravamen
         grd_MViNCo_Listad.Col = 4
         grd_MViNCo_Listad.Text = Format(l_arr_MViNCo(r_int_Contad).CuoCli_SegPre, "###,###,##0.00")
         r_dbl_SegPre = r_dbl_SegPre + CDbl(grd_MViNCo_Listad)
      
         'Seguro Vivienda
         grd_MViNCo_Listad.Col = 5
         grd_MViNCo_Listad.Text = Format(l_arr_MViNCo(r_int_Contad).CuoCli_SegViv, "###,###,##0.00")
         r_dbl_SegViv = r_dbl_SegViv + CDbl(grd_MViNCo_Listad)
      
         'Otros Cargos
         grd_MViNCo_Listad.Col = 6
         grd_MViNCo_Listad.Text = Format(l_arr_MViNCo(r_int_Contad).CuoCli_Portes, "###,###,##0.00")
         r_dbl_OtrCar = r_dbl_OtrCar + CDbl(grd_MViNCo_Listad)
      
         'Comisión COFIDE
         grd_MViNCo_Listad.Col = 7
         grd_MViNCo_Listad.Text = Format(l_arr_MViNCo(r_int_Contad).CuoCli_Comisi, "###,###,##0.00")
         r_dbl_Comisi = r_dbl_Comisi + CDbl(grd_MViNCo_Listad)
      
         'Valor Cuota
         grd_MViNCo_Listad.Col = 8
         grd_MViNCo_Listad.Text = Format(l_arr_MViNCo(r_int_Contad).CuoCli_ValCuo, "###,###,##0.00")
         r_dbl_TotCuo = r_dbl_TotCuo + CDbl(grd_MViNCo_Listad)
      
         'Saldo Capital
         grd_MViNCo_Listad.Col = 9
         grd_MViNCo_Listad.Text = Format(l_arr_MViNCo(r_int_Contad).CuoCli_SalCap, "###,###,##0.00")
      Next r_int_Contad
      
      grd_MViNCo_Listad.Redraw = True
      
      pnl_MViNCo_Capita.Caption = Format(r_dbl_Capita, "###,###,##0.00") & " "
      pnl_MViNCo_Intere.Caption = Format(r_dbl_Intere, "###,###,##0.00") & " "
      pnl_MViNCo_SegPre.Caption = Format(r_dbl_SegPre, "###,###,##0.00") & " "
      pnl_MViNCo_SegViv.Caption = Format(r_dbl_SegViv, "###,###,##0.00") & " "
      pnl_MViNCo_OtrCar.Caption = Format(r_dbl_OtrCar, "###,###,##0.00") & " "
      pnl_MViNCo_Comisi.Caption = Format(r_dbl_Comisi, "###,###,##0.00") & " "
      pnl_MViNCo_TotCuo.Caption = Format(r_dbl_TotCuo, "###,###,##0.00") & " "
      
      If grd_MViNCo_Listad.Rows > 0 Then
         Call gs_UbiIniGrid(grd_MViNCo_Listad)
      End If
      
   
      'Mivivienda Concesional
      r_dbl_Capita = 0
      r_dbl_Intere = 0
      r_dbl_Comisi = 0
      r_dbl_TotCuo = 0
      
      grd_MviCon_Listad.Redraw = False
      For r_int_Contad = 1 To UBound(l_arr_MViCon)
         grd_MviCon_Listad.Rows = grd_MviCon_Listad.Rows + 1
         grd_MviCon_Listad.Row = grd_MviCon_Listad.Rows - 1
         
         'Número de Cuota
         grd_MviCon_Listad.Col = 0
         grd_MviCon_Listad.Text = Format(r_int_Contad, "000")
      
         'Fecha de Vencimiento
         grd_MviCon_Listad.Col = 1
         grd_MviCon_Listad.Text = l_arr_MViCon(r_int_Contad).CuoCli_FecVct
      
         'Capital
         grd_MviCon_Listad.Col = 2
         grd_MviCon_Listad.Text = Format(l_arr_MViCon(r_int_Contad).CuoCli_Capita, "###,###,##0.00")
         r_dbl_Capita = r_dbl_Capita + CDbl(grd_MviCon_Listad)
         
         'Interes
         grd_MviCon_Listad.Col = 3
         grd_MviCon_Listad.Text = Format(l_arr_MViCon(r_int_Contad).CuoCli_Intere, "###,###,##0.00")
         r_dbl_Intere = r_dbl_Intere + CDbl(grd_MviCon_Listad)
      
         'Comisión
         grd_MviCon_Listad.Col = 4
         grd_MviCon_Listad.Text = Format(l_arr_MViCon(r_int_Contad).CuoCli_Comisi, "###,###,##0.00")
         r_dbl_Comisi = r_dbl_Comisi + CDbl(grd_MviCon_Listad)
      
         'Valor Cuota
         grd_MviCon_Listad.Col = 5
         grd_MviCon_Listad.Text = Format(l_arr_MViCon(r_int_Contad).CuoCli_ValCuo, "###,###,##0.00")
         r_dbl_TotCuo = r_dbl_TotCuo + CDbl(grd_MviCon_Listad)
      
         'Saldo Capital
         grd_MviCon_Listad.Col = 6
         grd_MviCon_Listad.Text = Format(l_arr_MViCon(r_int_Contad).CuoCli_SalCap, "###,###,##0.00")
      Next r_int_Contad
      
      grd_MviCon_Listad.Redraw = True
      
      pnl_MViCon_Capita.Caption = Format(r_dbl_Capita, "###,###,##0.00") & " "
      pnl_MViCon_Intere.Caption = Format(r_dbl_Intere, "###,###,##0.00") & " "
      pnl_MViCon_Comisi.Caption = Format(r_dbl_Comisi, "###,###,##0.00") & " "
      pnl_MViCon_TotCuo.Caption = Format(r_dbl_TotCuo, "###,###,##0.00") & " "
      
      If grd_MviCon_Listad.Rows > 0 Then
         Call gs_UbiIniGrid(grd_MviCon_Listad)
      End If
      
      'Cliente Concesional
      r_dbl_Capita = 0
      r_dbl_Intere = 0
      r_dbl_Comisi = 0
      r_dbl_TotCuo = 0
      
      grd_CliCon_Listad.Redraw = False
      For r_int_Contad = 1 To UBound(l_arr_CliCon)
         grd_CliCon_Listad.Rows = grd_CliCon_Listad.Rows + 1
         grd_CliCon_Listad.Row = grd_CliCon_Listad.Rows - 1
         
         'Número de Cuota
         grd_CliCon_Listad.Col = 0
         grd_CliCon_Listad.Text = Format(r_int_Contad, "000")
      
         'Fecha de Vencimiento
         grd_CliCon_Listad.Col = 1
         grd_CliCon_Listad.Text = l_arr_CliCon(r_int_Contad).CuoCli_FecVct
      
         'Capital
         grd_CliCon_Listad.Col = 2
         grd_CliCon_Listad.Text = Format(l_arr_CliCon(r_int_Contad).CuoCli_Capita, "###,###,##0.00")
         r_dbl_Capita = r_dbl_Capita + CDbl(grd_CliCon_Listad)
         
         'Interes
         grd_CliCon_Listad.Col = 3
         grd_CliCon_Listad.Text = Format(l_arr_CliCon(r_int_Contad).CuoCli_Intere, "###,###,##0.00")
         r_dbl_Intere = r_dbl_Intere + CDbl(grd_CliCon_Listad)
      
         'Valor Cuota
         grd_CliCon_Listad.Col = 4
         grd_CliCon_Listad.Text = Format(l_arr_CliCon(r_int_Contad).CuoCli_ValCuo, "###,###,##0.00")
         r_dbl_TotCuo = r_dbl_TotCuo + CDbl(grd_CliCon_Listad)
      
         'Saldo Capital
         grd_CliCon_Listad.Col = 5
         grd_CliCon_Listad.Text = Format(l_arr_CliCon(r_int_Contad).CuoCli_SalCap, "###,###,##0.00")
      Next r_int_Contad
      
      grd_CliCon_Listad.Redraw = True
      
      pnl_CliCon_Capita.Caption = Format(r_dbl_Capita, "###,###,##0.00") & " "
      pnl_CliCon_Intere.Caption = Format(r_dbl_Intere, "###,###,##0.00") & " "
      pnl_CliCon_TotCuo.Caption = Format(r_dbl_TotCuo, "###,###,##0.00") & " "
      
      If grd_CliCon_Listad.Rows > 0 Then
         Call gs_UbiIniGrid(grd_CliCon_Listad)
      End If
   
      If moddat_g_str_CodPrd = "003" Then
         'Cofide No Concesional
         r_dbl_Capita = 0
         r_dbl_Intere = 0
         r_dbl_Comisi = 0
         r_dbl_TotCuo = 0
         
         grd_CofNCo_Listad.Redraw = False
         For r_int_Contad = 1 To UBound(l_arr_CofNCo)
            grd_CofNCo_Listad.Rows = grd_CofNCo_Listad.Rows + 1
            grd_CofNCo_Listad.Row = grd_CofNCo_Listad.Rows - 1
            
            'Número de Cuota
            grd_CofNCo_Listad.Col = 0
            grd_CofNCo_Listad.Text = Format(r_int_Contad, "000")
         
            'Fecha de Vencimiento
            grd_CofNCo_Listad.Col = 1
            grd_CofNCo_Listad.Text = l_arr_CofNCo(r_int_Contad).CuoCli_FecVct
         
            'Capital
            grd_CofNCo_Listad.Col = 2
            grd_CofNCo_Listad.Text = Format(l_arr_CofNCo(r_int_Contad).CuoCli_Capita, "###,###,##0.00")
            r_dbl_Capita = r_dbl_Capita + CDbl(grd_CofNCo_Listad)
            
            'Interes
            grd_CofNCo_Listad.Col = 3
            grd_CofNCo_Listad.Text = Format(l_arr_CofNCo(r_int_Contad).CuoCli_Intere, "###,###,##0.00")
            r_dbl_Intere = r_dbl_Intere + CDbl(grd_CofNCo_Listad)
         
            'Comisión
            grd_CofNCo_Listad.Col = 4
            grd_CofNCo_Listad.Text = Format(l_arr_CofNCo(r_int_Contad).CuoCli_Comisi, "###,###,##0.00")
            r_dbl_Comisi = r_dbl_Comisi + CDbl(grd_CofNCo_Listad)
         
            'Valor Cuota
            grd_CofNCo_Listad.Col = 5
            grd_CofNCo_Listad.Text = Format(l_arr_CofNCo(r_int_Contad).CuoCli_ValCuo, "###,###,##0.00")
            r_dbl_TotCuo = r_dbl_TotCuo + CDbl(grd_CofNCo_Listad)
         
            'Saldo Capital
            grd_CofNCo_Listad.Col = 6
            grd_CofNCo_Listad.Text = Format(l_arr_CofNCo(r_int_Contad).CuoCli_SalCap, "###,###,##0.00")
         Next r_int_Contad
         
         grd_CofNCo_Listad.Redraw = True
         
         pnl_CofNCo_Capita.Caption = Format(r_dbl_Capita, "###,###,##0.00") & " "
         pnl_CofNCo_Intere.Caption = Format(r_dbl_Intere, "###,###,##0.00") & " "
         pnl_CofNCo_Comisi.Caption = Format(r_dbl_Comisi, "###,###,##0.00") & " "
         pnl_CofNCo_TotCuo.Caption = Format(r_dbl_TotCuo, "###,###,##0.00") & " "
         
         If grd_CofNCo_Listad.Rows > 0 Then
            Call gs_UbiIniGrid(grd_CofNCo_Listad)
         End If
      End If
   End If
End Sub

Private Sub fs_Envia_CorEle(ByVal p_Asunto As String, ByVal p_Mensaje As String)
   Dim r_str_Cadena     As String
   
   'Destinatarios de Correo
   ReDim moddat_g_arr_Genera(0)
   
   'Consejero Hipotecario
   ReDim Preserve moddat_g_arr_Genera(UBound(moddat_g_arr_Genera) + 1)
   moddat_g_arr_Genera(UBound(moddat_g_arr_Genera)).Genera_Codigo = moddat_gf_Buscar_DirEle_Codigo(moddat_g_str_CodConHip)
   
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

Private Sub fs_Grabar_Desemb()
   Dim r_dbl_PorITF     As Double
   Dim r_dbl_ImpITF     As Double
   Dim r_str_Operac     As String
   
   If l_int_FlgPry = 2 Then
      r_str_Operac = moddat_gf_Consulta_Operac(moddat_g_str_CodPrd, "22")
   Else
      r_str_Operac = moddat_gf_Consulta_Operac(moddat_g_str_CodPrd, "21")
   End If
   
   r_str_Operac = CStr(moddat_g_int_TipMon) & Right(r_str_Operac, 5)
   
   r_dbl_PorITF = opecaj_gf_Consulta_ITF(Format(CDate(moddat_g_str_FecSis), "yyyymmdd"), 1)
   r_dbl_ImpITF = CDbl(gf_Truncar_Numero(CDbl(ipp_MtoDes.Text) * (r_dbl_PorITF / 100), 2))

   'Grabando Cabecera de Credito
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_CRE_HIPDES_OPEIND ("
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumSol & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "', "
      
      If l_int_FlgPry = 2 Then
         g_str_Parame = g_str_Parame & "1, "
         g_str_Parame = g_str_Parame & "'" & txt_NumChq.Text & "', "
         g_str_Parame = g_str_Parame & "'" & l_arr_CtaChq(cmb_CtaChq.ListIndex + 1).Genera_Codigo & "', "
         g_str_Parame = g_str_Parame & "'" & l_arr_BanChq(cmb_BanChq.ListIndex + 1).Genera_Codigo & "', "
         g_str_Parame = g_str_Parame & CStr(cmb_MonDes.ItemData(cmb_MonDes.ListIndex)) & ", "
         g_str_Parame = g_str_Parame & CStr(CDbl(ipp_MtoDes.Text)) & ", "
         g_str_Parame = g_str_Parame & CStr(r_dbl_ImpITF) & ", "
         g_str_Parame = g_str_Parame & CStr(CDbl(pnl_DesMPr.Caption)) & ", "
         g_str_Parame = g_str_Parame & CStr(CDbl(pnl_TCaDol.Caption)) & ", "
         g_str_Parame = g_str_Parame & CStr(CDbl(pnl_TCaMPr.Caption)) & ", "
         g_str_Parame = g_str_Parame & CStr(cmb_FlgFia.ItemData(cmb_FlgFia.ListIndex)) & ", "
         
         If cmb_FlgFia.ItemData(cmb_FlgFia.ListIndex) = 1 Then
            g_str_Parame = g_str_Parame & "'" & l_arr_BanFia(cmb_BanFia.ListIndex + 1).Genera_Codigo & "', "
            g_str_Parame = g_str_Parame & "'" & txt_NumFia.Text & "', "
            g_str_Parame = g_str_Parame & Format(CDate(ipp_FEmFia.Text), "yyyymmdd") & ", "
            g_str_Parame = g_str_Parame & Format(CDate(ipp_FVcFia.Text), "yyyymmdd") & ", "
            g_str_Parame = g_str_Parame & CStr(cmb_MonFia.ItemData(cmb_MonFia.ListIndex)) & ", "
            g_str_Parame = g_str_Parame & CStr(CDbl(ipp_MtoFia.Text)) & ", "
         Else
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "'', "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
            g_str_Parame = g_str_Parame & "0, "
         End If
      Else
         g_str_Parame = g_str_Parame & "2, "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & CStr(cmb_MonDes.ItemData(cmb_MonDes.ListIndex)) & ", "
         g_str_Parame = g_str_Parame & CStr(CDbl(ipp_MtoDes.Text)) & ", "
         g_str_Parame = g_str_Parame & CStr(r_dbl_ImpITF) & ", "
         g_str_Parame = g_str_Parame & CStr(CDbl(pnl_DesMPr.Caption)) & ", "
         g_str_Parame = g_str_Parame & CStr(CDbl(pnl_TCaDol.Caption)) & ", "
         g_str_Parame = g_str_Parame & CStr(CDbl(pnl_TCaMPr.Caption)) & ", "
         g_str_Parame = g_str_Parame & "2, "
         
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "0, "
      End If
      
      g_str_Parame = g_str_Parame & "'" & txt_Observ.Text & "', "
      g_str_Parame = g_str_Parame & "'" & r_str_Operac & "', "
      g_str_Parame = g_str_Parame & "'" & CStr(moddat_g_int_TipDoc) & moddat_g_str_NumDoc & "', "
      
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                           'Código Usuario
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                           'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                            'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "                           'Código Sucursal
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
End Sub

Private Sub fs_Grabar_Cronog()
   Dim r_int_Contad  As Integer
   Dim r_int_NumCuo  As Integer
   Dim r_str_FecVct  As String
   Dim r_dbl_Capita  As Double
   Dim r_dbl_Intere  As Double
   Dim r_dbl_ComCof  As Double
   Dim r_dbl_SegDes  As Double
   Dim r_dbl_SegViv  As Double
   Dim r_dbl_OtrCar  As Double
   Dim r_dbl_SalCap  As Double
   
   'Grabando Cronograma Cliente No Concesional
   grd_CliNCo_Listad.Redraw = False
   For r_int_Contad = 0 To grd_CliNCo_Listad.Rows - 1
      grd_CliNCo_Listad.Row = r_int_Contad
   
      grd_CliNCo_Listad.Col = 0:          r_int_NumCuo = CInt(grd_CliNCo_Listad.Text)
      grd_CliNCo_Listad.Col = 1:          r_str_FecVct = grd_CliNCo_Listad.Text
      grd_CliNCo_Listad.Col = 2:          r_dbl_Capita = CDbl(grd_CliNCo_Listad.Text)
      grd_CliNCo_Listad.Col = 3:          r_dbl_Intere = CDbl(grd_CliNCo_Listad.Text)
      grd_CliNCo_Listad.Col = 4:          r_dbl_SegDes = CDbl(grd_CliNCo_Listad.Text)
      grd_CliNCo_Listad.Col = 5:          r_dbl_SegViv = CDbl(grd_CliNCo_Listad.Text)
      grd_CliNCo_Listad.Col = 6:          r_dbl_OtrCar = CDbl(grd_CliNCo_Listad.Text)
      grd_CliNCo_Listad.Col = 8:          r_dbl_SalCap = CDbl(grd_CliNCo_Listad.Text)
      
      DoEvents
      If Not ff_Inserta_HipCuo(moddat_g_str_NumOpe, 1, r_int_NumCuo, r_str_FecVct, r_dbl_Capita, r_dbl_Intere, r_dbl_SegDes, r_dbl_SegViv, r_dbl_OtrCar, r_dbl_SalCap, 0, 0, 0) Then
         Exit Sub
      End If
   Next r_int_Contad
   
   grd_CliNCo_Listad.Redraw = True
   'Call gs_UbiIniGrid(grd_CliNCo_Listad)
   
   If moddat_g_str_CodPrd = "001" Or moddat_g_str_CodPrd = "004" Or moddat_g_str_CodPrd = "003" Then
      'Grabando Cronograma Mivivienda No Concesional
      grd_MViNCo_Listad.Redraw = False
      For r_int_Contad = 0 To grd_MViNCo_Listad.Rows - 1
         grd_MViNCo_Listad.Row = r_int_Contad
      
         grd_MViNCo_Listad.Col = 0:          r_int_NumCuo = CInt(grd_MViNCo_Listad.Text)
         grd_MViNCo_Listad.Col = 1:          r_str_FecVct = grd_MViNCo_Listad.Text
         grd_MViNCo_Listad.Col = 2:          r_dbl_Capita = CDbl(grd_MViNCo_Listad.Text)
         grd_MViNCo_Listad.Col = 3:          r_dbl_Intere = CDbl(grd_MViNCo_Listad.Text)
         grd_MViNCo_Listad.Col = 4:          r_dbl_SegDes = CDbl(grd_MViNCo_Listad.Text)
         grd_MViNCo_Listad.Col = 5:          r_dbl_SegViv = CDbl(grd_MViNCo_Listad.Text)
         grd_MViNCo_Listad.Col = 6:          r_dbl_OtrCar = CDbl(grd_MViNCo_Listad.Text)
         grd_MViNCo_Listad.Col = 7:          r_dbl_ComCof = CDbl(grd_MViNCo_Listad.Text)
         grd_MViNCo_Listad.Col = 9:          r_dbl_SalCap = CDbl(grd_MViNCo_Listad.Text)
         
         If r_dbl_Capita = 0 Then
            r_dbl_Intere = 0
            r_dbl_SegDes = 0
            r_dbl_SegViv = 0
         End If
         
         DoEvents
         If Not ff_Inserta_HipCuo(moddat_g_str_NumOpe, 3, r_int_NumCuo, r_str_FecVct, r_dbl_Capita, r_dbl_Intere, r_dbl_SegDes, r_dbl_SegViv, r_dbl_OtrCar, r_dbl_SalCap, 0, 0, r_dbl_ComCof) Then
            Exit Sub
         End If
      Next r_int_Contad
      
      grd_MViNCo_Listad.Redraw = True
   
      'Cliente Concesional
      grd_CliCon_Listad.Redraw = False
      For r_int_Contad = 0 To grd_CliCon_Listad.Rows - 1
         grd_CliCon_Listad.Row = r_int_Contad
      
         grd_CliCon_Listad.Col = 0:          r_int_NumCuo = CInt(grd_CliCon_Listad.Text)
         grd_CliCon_Listad.Col = 1:          r_str_FecVct = grd_CliCon_Listad.Text
         grd_CliCon_Listad.Col = 2:          r_dbl_Capita = CDbl(grd_CliCon_Listad.Text)
         grd_CliCon_Listad.Col = 3:          r_dbl_Intere = CDbl(grd_CliCon_Listad.Text)
         grd_CliCon_Listad.Col = 5:          r_dbl_SalCap = CDbl(grd_CliCon_Listad.Text)
         
         DoEvents
         If Not ff_Inserta_HipCuo(moddat_g_str_NumOpe, 2, r_int_NumCuo, r_str_FecVct, r_dbl_Capita, r_dbl_Intere, 0, 0, 0, r_dbl_SalCap, 0, 0, 0) Then
            Exit Sub
         End If
         
      Next r_int_Contad
      
      grd_CliCon_Listad.Redraw = True
      'Call gs_UbiIniGrid(grd_CliCon_Listad)
   
      'Mivivienda Concesional
      grd_MviCon_Listad.Redraw = False
      For r_int_Contad = 0 To grd_MviCon_Listad.Rows - 1
         grd_MviCon_Listad.Row = r_int_Contad
      
         grd_MviCon_Listad.Col = 0:          r_int_NumCuo = CInt(grd_MviCon_Listad.Text)
         grd_MviCon_Listad.Col = 1:          r_str_FecVct = grd_MviCon_Listad.Text
         grd_MviCon_Listad.Col = 2:          r_dbl_Capita = CDbl(grd_MviCon_Listad.Text)
         grd_MviCon_Listad.Col = 3:          r_dbl_Intere = CDbl(grd_MviCon_Listad.Text)
         grd_MviCon_Listad.Col = 4:          r_dbl_ComCof = CDbl(grd_MviCon_Listad.Text)
         grd_MviCon_Listad.Col = 6:          r_dbl_SalCap = CDbl(grd_MviCon_Listad.Text)
         
         DoEvents
         If Not ff_Inserta_HipCuo(moddat_g_str_NumOpe, 4, r_int_NumCuo, r_str_FecVct, r_dbl_Capita, r_dbl_Intere, 0, 0, 0, r_dbl_SalCap, 0, 0, r_dbl_ComCof) Then
            Exit Sub
         End If
         
      Next r_int_Contad
      
      grd_MviCon_Listad.Redraw = True
      'Call gs_UbiIniGrid(grd_MViCon_Listad)
   
      If moddat_g_str_CodPrd = "003" Then
         'Cofide No Concesional
         grd_CofNCo_Listad.Redraw = False
         For r_int_Contad = 0 To grd_CofNCo_Listad.Rows - 1
            grd_CofNCo_Listad.Row = r_int_Contad
         
            grd_CofNCo_Listad.Col = 0:          r_int_NumCuo = CInt(grd_CofNCo_Listad.Text)
            grd_CofNCo_Listad.Col = 1:          r_str_FecVct = grd_CofNCo_Listad.Text
            grd_CofNCo_Listad.Col = 2:          r_dbl_Capita = CDbl(grd_CofNCo_Listad.Text)
            grd_CofNCo_Listad.Col = 3:          r_dbl_Intere = CDbl(grd_CofNCo_Listad.Text)
            grd_CofNCo_Listad.Col = 4:          r_dbl_ComCof = CDbl(grd_CofNCo_Listad.Text)
            grd_CofNCo_Listad.Col = 6:          r_dbl_SalCap = CDbl(grd_CofNCo_Listad.Text)
            
            DoEvents
            If Not ff_Inserta_HipCuo(moddat_g_str_NumOpe, 5, r_int_NumCuo, r_str_FecVct, r_dbl_Capita, r_dbl_Intere, 0, 0, 0, r_dbl_SalCap, 0, 0, r_dbl_ComCof) Then
               Exit Sub
            End If
            
         Next r_int_Contad
         
         grd_CofNCo_Listad.Redraw = True
      End If
   End If
End Sub

Private Sub fs_Grabar_Credit()
   Dim r_str_PriVct     As String
   Dim r_str_UltVct     As String
   Dim r_dbl_CuoFij     As Double


   grd_CliNCo_Listad.Row = 0
   grd_CliNCo_Listad.Col = 1
   r_str_PriVct = grd_CliNCo_Listad.Text
   
   grd_CliNCo_Listad.Row = grd_CliNCo_Listad.Rows - 1
   grd_CliNCo_Listad.Col = 1
   r_str_UltVct = grd_CliNCo_Listad.Text
   
   grd_CliNCo_Listad.Row = 1
   grd_CliNCo_Listad.Col = 7
   r_dbl_CuoFij = CDbl(grd_CliNCo_Listad.Text)
   
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_CRE_HIPMAE_GENCRO ("
      
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "', "
      g_str_Parame = g_str_Parame & Format(CDate(r_str_PriVct), "yyyymmdd") & ", "
      g_str_Parame = g_str_Parame & Format(CDate(r_str_UltVct), "yyyymmdd") & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_IntCap) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_MtoPre + l_dbl_IntCap) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_CuoFij) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_CliNCo_Capita.Caption)) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_MtoNCo) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_MtoNCo) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_MtoCon) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_MtoCon) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_TasMVi) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_ComCRC) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_ComPBP) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_ComCof) & ", "
      g_str_Parame = g_str_Parame & CStr(l_dbl_TasCof) & ", "
      
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                           'Código Usuario
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                           'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                            'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "                           'Código Sucursal
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
End Sub

Private Function ff_Inserta_HipCuo(ByVal p_NumOpe As String, ByVal p_TipCro As Integer, ByVal p_NumCuo As Integer, ByVal p_FecVct As String, ByVal p_Capita As Double, ByVal p_Intere As Double, ByVal p_SegDes As Double, ByVal p_SegViv As Double, ByVal p_OtrGas As Double, ByVal p_SalCap As Double, ByVal p_ComCrc As Double, ByVal p_ComPbp As Double, ByVal p_ComCof As Double) As Integer
   ff_Inserta_HipCuo = False
   
   'Grabando Cabecera de Credito
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_CRE_HIPCUO_CREA ("
      
      g_str_Parame = g_str_Parame & "'" & p_NumOpe & "', "
      g_str_Parame = g_str_Parame & CStr(p_TipCro) & ", "
      g_str_Parame = g_str_Parame & CStr(p_NumCuo) & ", "
      g_str_Parame = g_str_Parame & Format(CDate(p_FecVct), "yyyymmdd") & ", "
      g_str_Parame = g_str_Parame & CStr(p_Capita) & ", "
      g_str_Parame = g_str_Parame & CStr(p_Intere) & ", "
      g_str_Parame = g_str_Parame & CStr(p_SegDes) & ", "
      g_str_Parame = g_str_Parame & CStr(p_SegViv) & ", "
      g_str_Parame = g_str_Parame & CStr(p_OtrGas) & ", "
      g_str_Parame = g_str_Parame & CStr(p_SalCap) & ", "
      g_str_Parame = g_str_Parame & CStr(p_ComCrc) & ", "
      g_str_Parame = g_str_Parame & CStr(p_ComPbp) & ", "
      g_str_Parame = g_str_Parame & CStr(p_ComCof) & ", "
      
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                           'Código Usuario
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                           'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                            'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "                           'Código Sucursal
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_CRE_HIPCUO_CREA. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Function
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop

   ff_Inserta_HipCuo = True
End Function

Private Sub fs_Grabar_DatCli(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String)
   Dim r_int_TipCli As Integer
   Dim r_int_FlgDoA As Integer
   Dim r_int_TipDoA As Integer
   Dim r_int_EstCiv As Integer
   Dim r_int_RegCyg As Integer
   Dim r_int_NivEst As Integer
   Dim r_int_CodSex As Integer
   Dim r_int_DepEco As Integer
   Dim r_int_Edad01 As Integer
   Dim r_int_Edad02 As Integer
   Dim r_int_Edad03 As Integer
   Dim r_int_Edad04 As Integer
   Dim r_int_Edad05 As Integer
   Dim r_int_TipVia As Integer
   Dim r_int_TipZon As Integer
   Dim r_int_AutEnv As Integer
   Dim r_int_CarDom As Integer
   Dim r_int_AnoDom As Integer
   Dim r_int_ActEco As Integer
   Dim r_int_Ocupac As Integer
   Dim r_int_CodCiu As Integer
   Dim r_int_TDoTri As Integer
   Dim r_int_CygTDo As Integer

   Dim r_str_NumDoA As String
   Dim r_str_ApePat As String
   Dim r_str_ApeMat As String
   Dim r_str_ApeCas As String
   Dim r_str_Nombre As String
   Dim r_str_Profes As String
   Dim r_str_NacFec As String
   Dim r_str_NacPai As String
   Dim r_str_NacLug As String
   Dim r_str_NomVia As String
   Dim r_str_Numero As String
   Dim r_str_IntDpt As String
   Dim r_str_NomZon As String
   Dim r_str_UbiGeo As String
   Dim r_str_Refere As String
   Dim r_str_NumCel As String
   Dim r_str_Telefo As String
   Dim r_str_DirEle As String
   Dim r_str_ClaSbs As String
   Dim r_str_ClasMC As String
   Dim r_str_Reside As String
   Dim r_str_FlgAcc As String
   Dim r_str_RelLab As String
   Dim r_str_NDoTri As String
   Dim r_str_CygNDo As String
   Dim r_str_CodSbs As String
   
   g_str_Parame = "SELECT * FROM CLI_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_TIPDOC = " & CStr(p_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_NUMDOC = '" & p_NumDoc & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If

   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      r_int_TipCli = g_rst_Princi!DATGEN_TIPCLI
      r_int_FlgDoA = g_rst_Princi!DatGen_FLGDOA
      r_int_TipDoA = g_rst_Princi!DatGen_TIPDOA
      r_str_NumDoA = Trim(g_rst_Princi!DatGen_NUMDOA & "")
      
      r_str_ApePat = Trim(g_rst_Princi!DatGen_ApePat & "")
      r_str_ApeMat = Trim(g_rst_Princi!DatGen_ApeMat & "")
      r_str_ApeCas = Trim(g_rst_Princi!DatGen_ApeCas & "")
      r_str_Nombre = Trim(g_rst_Princi!DatGen_Nombre & "")
      
      r_int_EstCiv = g_rst_Princi!DATGEN_ESTCIV
      r_int_RegCyg = g_rst_Princi!DatGen_RegCyg
      r_int_NivEst = g_rst_Princi!DatGen_NivEst
      r_str_Profes = g_rst_Princi!DatGen_Profes
      r_int_CodSex = g_rst_Princi!DatGen_CodSex
      r_str_NacFec = CStr(g_rst_Princi!DATGEN_NACFEC)
      r_str_NacPai = g_rst_Princi!DATGEN_NACPAI
      r_str_NacLug = g_rst_Princi!DATGEN_NACLUG
      
      r_int_DepEco = g_rst_Princi!DatGen_DepEco
      r_int_Edad01 = g_rst_Princi!DatGen_EDAD01
      r_int_Edad02 = g_rst_Princi!DatGen_EDAD02
      r_int_Edad03 = g_rst_Princi!DatGen_EDAD03
      r_int_Edad04 = g_rst_Princi!DatGen_EDAD04
      r_int_Edad05 = g_rst_Princi!DatGen_EDAD05
      
      r_int_TipVia = g_rst_Princi!DatGen_TipVia
      r_str_NomVia = Trim(g_rst_Princi!DatGen_NomVia & "")
      r_str_Numero = Trim(g_rst_Princi!DatGen_Numero & "")
      r_str_IntDpt = Trim(g_rst_Princi!DatGen_IntDpt & "")
      r_int_TipZon = g_rst_Princi!DatGen_TipZon
      r_str_NomZon = Trim(g_rst_Princi!DatGen_NomZon & "")
      r_str_UbiGeo = g_rst_Princi!DatGen_Ubigeo
      r_str_Refere = Trim(g_rst_Princi!DatGen_Refere & "")
      
      r_str_NumCel = Trim(g_rst_Princi!DATGEN_NUMCEL & "")
      r_str_Telefo = Trim(g_rst_Princi!DatGen_Telefo & "")
      r_str_DirEle = Trim(g_rst_Princi!DatGen_DirEle & "")
      r_int_AutEnv = g_rst_Princi!DATGEN_AUTENV
      
      r_int_CarDom = g_rst_Princi!DATGEN_CARDOM
      r_int_AnoDom = g_rst_Princi!DatGen_ANODOM
      
      r_str_ClaSbs = Trim(g_rst_Princi!DATGEN_CLASBS & "")
      r_str_ClasMC = Trim(g_rst_Princi!DATGEN_CLASMC & "")
      
      r_str_Reside = Trim(g_rst_Princi!DATGEN_RESIDE & "")
      r_str_FlgAcc = Trim(g_rst_Princi!DATGEN_FLGACC & "")
      r_str_RelLab = Trim(g_rst_Princi!DATGEN_RELLAB & "")
      r_int_ActEco = g_rst_Princi!DATGEN_ACTECO
      r_int_Ocupac = g_rst_Princi!DATGEN_OCUPAC
      r_int_CodCiu = g_rst_Princi!DATGEN_CODCIU
      r_int_TDoTri = g_rst_Princi!DATGEN_TDOTRI
      r_str_NDoTri = Trim(g_rst_Princi!DATGEN_NDOTRI & "")
      r_int_CygTDo = g_rst_Princi!DATGEN_CYGTDO
      r_str_CygNDo = Trim(g_rst_Princi!DATGEN_CYGNDO & "")
      r_str_CodSbs = Trim(g_rst_Princi!DatGen_CodSbs & "")
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   'Grabando Datos
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      g_str_Parame = "USP_TRA_CLIGEN ("
      
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumSol & "', "
      g_str_Parame = g_str_Parame & CStr(r_int_TipCli) & ", "
      g_str_Parame = g_str_Parame & CStr(p_TipDoc) & ", "
      g_str_Parame = g_str_Parame & "'" & p_NumDoc & "', "
      g_str_Parame = g_str_Parame & CStr(r_int_FlgDoA) & ", "
      g_str_Parame = g_str_Parame & CStr(r_int_TipDoA) & ", "
      g_str_Parame = g_str_Parame & "'" & r_str_NumDoA & "', "
      g_str_Parame = g_str_Parame & "'" & r_str_ApePat & "', "
      g_str_Parame = g_str_Parame & "'" & r_str_ApeMat & "', "
      g_str_Parame = g_str_Parame & "'" & r_str_ApeCas & "', "
      g_str_Parame = g_str_Parame & "'" & r_str_Nombre & "', "
      g_str_Parame = g_str_Parame & CStr(r_int_EstCiv) & ", "
      g_str_Parame = g_str_Parame & CStr(r_int_RegCyg) & ", "
      g_str_Parame = g_str_Parame & CStr(r_int_NivEst) & ", "
      g_str_Parame = g_str_Parame & "'" & r_str_Profes & "', "
      g_str_Parame = g_str_Parame & CStr(r_int_CodSex) & ", "
      g_str_Parame = g_str_Parame & r_str_NacFec & ", "
      g_str_Parame = g_str_Parame & r_str_NacLug & ", "
      g_str_Parame = g_str_Parame & r_str_NacPai & ", "
      g_str_Parame = g_str_Parame & CStr(r_int_DepEco) & ", "
      g_str_Parame = g_str_Parame & CStr(r_int_Edad01) & ", "
      g_str_Parame = g_str_Parame & CStr(r_int_Edad02) & ", "
      g_str_Parame = g_str_Parame & CStr(r_int_Edad03) & ", "
      g_str_Parame = g_str_Parame & CStr(r_int_Edad04) & ", "
      g_str_Parame = g_str_Parame & CStr(r_int_Edad05) & ", "
      g_str_Parame = g_str_Parame & CStr(r_int_TipVia) & ", "
      g_str_Parame = g_str_Parame & "'" & r_str_NomVia & "', "
      g_str_Parame = g_str_Parame & "'" & r_str_Numero & "', "
      g_str_Parame = g_str_Parame & "'" & r_str_IntDpt & "', "
      g_str_Parame = g_str_Parame & CStr(r_int_TipZon) & ", "
      g_str_Parame = g_str_Parame & "'" & r_str_NomZon & "', "
      g_str_Parame = g_str_Parame & "'" & r_str_Refere & "', "
      g_str_Parame = g_str_Parame & "'" & r_str_UbiGeo & "', "
      g_str_Parame = g_str_Parame & "'" & r_str_NumCel & "', "
      g_str_Parame = g_str_Parame & "'" & r_str_Telefo & "', "
      g_str_Parame = g_str_Parame & "'" & r_str_DirEle & "', "
      g_str_Parame = g_str_Parame & CStr(r_int_AutEnv) & ", "
      g_str_Parame = g_str_Parame & CStr(r_int_CarDom) & ", "
      g_str_Parame = g_str_Parame & CStr(r_int_AnoDom) & ", "
      g_str_Parame = g_str_Parame & "'" & r_str_ClaSbs & "', "
      g_str_Parame = g_str_Parame & "'" & r_str_ClasMC & "', "
      g_str_Parame = g_str_Parame & "'" & r_str_Reside & "', "
      g_str_Parame = g_str_Parame & "'" & r_str_FlgAcc & "', "
      g_str_Parame = g_str_Parame & "'" & r_str_RelLab & "', "
      g_str_Parame = g_str_Parame & CStr(r_int_ActEco) & ", "
      g_str_Parame = g_str_Parame & CStr(r_int_Ocupac) & ", "
      g_str_Parame = g_str_Parame & CStr(r_int_CodCiu) & ", "
      g_str_Parame = g_str_Parame & CStr(r_int_TDoTri) & ", "
      g_str_Parame = g_str_Parame & "'" & r_str_NDoTri & "', "
      g_str_Parame = g_str_Parame & CStr(r_int_CygTDo) & ", "
      g_str_Parame = g_str_Parame & "'" & r_str_CygNDo & "', "
      
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "                           'Código Usuario
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "                           'Nombre Terminal
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "                            'Nombre Ejecutable
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "                           'Código Sucursal
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If

      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar el procedimiento USP_TRA_CLIGEN. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
   Loop
End Sub

Private Sub fs_Cargar_ActEco_Tit(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String, ByVal p_Indice As Integer)
   g_str_Parame = "SELECT * FROM CLI_ACTECO WHERE "
   g_str_Parame = g_str_Parame & "ACTECO_CLITDO = " & CStr(p_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "ACTECO_CLINDO = '" & p_NumDoc & "' AND "
   g_str_Parame = g_str_Parame & "ACTECO_ORDACT = " & CStr(p_Indice) & " "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If

   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_OrdAct = p_Indice
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_TipAct = g_rst_Princi!ActEco_CodAct
   
      'Dependiente
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_SitTra = g_rst_Princi!ActEco_Dep_SitTra
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_TipDoc = g_rst_Princi!ActEco_Dep_TipDoc
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NumDoc = Trim(g_rst_Princi!ActEco_Dep_NumDoc & "")
      
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_TipOfi = g_rst_Princi!ActEco_Dep_TipOfi
      
      'Buscar si empresa ya esta registrada
      g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
      g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_TipDoc) & " AND "
      g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NumDoc & "' "

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
         Exit Sub
      End If
      
      If g_rst_Genera.BOF And g_rst_Genera.EOF Then
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_FlgEmp = "9"
      
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_RazSoc = Trim(g_rst_Princi!ActEco_Dep_RazSoc & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NomCom = Trim(g_rst_Princi!ActEco_Dep_NomCom & "")
      
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_CodCiu = g_rst_Princi!ActEco_Dep_CodCiu
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_TeleRH = Trim(g_rst_Princi!ActEco_Dep_TeleRH & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_AnexRH = Trim(g_rst_Princi!ActEco_Dep_AnexRH & "")
      
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_TipVia = g_rst_Princi!ActEco_Dep_TipVia
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NomVia = Trim(g_rst_Princi!ActEco_Dep_NomVia & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NumVia = Trim(g_rst_Princi!ActEco_Dep_NumVia & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_IntDpt = Trim(g_rst_Princi!ActEco_Dep_IntDpt & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_TipZon = g_rst_Princi!ActEco_Dep_TipZon
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NomZon = Trim(g_rst_Princi!ActEco_Dep_NomZon & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_Refere = Trim(g_rst_Princi!ActEco_Dep_Refere & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_UbiGeo = Trim(g_rst_Princi!ActEco_Dep_UbiGeo & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_Telef1 = Trim(g_rst_Princi!ActEco_Dep_Telef1 & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_Telef2 = Trim(g_rst_Princi!ActEco_Dep_Telef2 & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NumFax = Trim(g_rst_Princi!ActEco_Dep_NumFax & "")
      Else
         g_rst_Genera.MoveFirst
      
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_FlgEmp = CStr(g_rst_Genera!DATGEN_CLASIF)
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_RazSoc = Trim(g_rst_Genera!DATGEN_RAZSOC & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NomCom = Trim(g_rst_Genera!DATGEN_NOMCOM & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_CodCiu = g_rst_Genera!DATGEN_CODCIU
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_TeleRH = Trim(g_rst_Genera!DATGEN_TELERH & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_AnexRH = Trim(g_rst_Genera!DATGEN_ANEXRH & "")
      
         If moddat_g_arr_ActEco_Tit(moddat_g_int_OrdAct).ActEco_Dep_TipOfi = 1 Then
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_TipVia = g_rst_Genera!DatGen_TipVia
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NomVia = Trim(g_rst_Genera!DatGen_NomVia & "")
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NumVia = Trim(g_rst_Genera!DatGen_numVia & "")
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_IntDpt = Trim(g_rst_Genera!DatGen_IntDpt & "")
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_TipZon = g_rst_Genera!DatGen_TipZon
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NomZon = Trim(g_rst_Genera!DatGen_NomZon & "")
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_Refere = Trim(g_rst_Genera!DatGen_Refere & "")
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_UbiGeo = Trim(g_rst_Genera!DatGen_Ubigeo & "")
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_Telef1 = Trim(g_rst_Genera!DATGEN_TELEF1 & "")
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_Telef2 = Trim(g_rst_Genera!DATGEN_TELEF2 & "")
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NumFax = Trim(g_rst_Genera!DatGen_NUMFAX & "")
         Else
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_TipVia = g_rst_Princi!ActEco_Dep_TipVia
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NomVia = Trim(g_rst_Princi!ActEco_Dep_NomVia & "")
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NumVia = Trim(g_rst_Princi!ActEco_Dep_NumVia & "")
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_IntDpt = Trim(g_rst_Princi!ActEco_Dep_IntDpt & "")
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_TipZon = g_rst_Princi!ActEco_Dep_TipZon
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NomZon = Trim(g_rst_Princi!ActEco_Dep_NomZon & "")
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_Refere = Trim(g_rst_Princi!ActEco_Dep_Refere & "")
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_UbiGeo = Trim(g_rst_Princi!ActEco_Dep_UbiGeo & "")
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_Telef1 = Trim(g_rst_Princi!ActEco_Dep_Telef1 & "")
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_Telef2 = Trim(g_rst_Princi!ActEco_Dep_Telef2 & "")
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NumFax = Trim(g_rst_Princi!ActEco_Dep_NumFax & "")
         End If
      End If
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_IngNet = g_rst_Princi!ActEco_Dep_IngNet
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_FreHab = g_rst_Princi!ActEco_Dep_FreHab
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_FecIng = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Dep_FecIng))
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_CodCar = Trim(g_rst_Princi!ActEco_Dep_CodCar & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NomCar = Trim(g_rst_Princi!ActEco_Dep_NomCar & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NomAre = Trim(g_rst_Princi!ActEco_Dep_NomAre & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NumAnx = Trim(g_rst_Princi!ActEco_Dep_NumAnx & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_TelDir = Trim(g_rst_Princi!ActEco_Dep_TelDir & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_Celula = Trim(g_rst_Princi!ActEco_Dep_Celula & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_DirEle = Trim(g_rst_Princi!ActEco_Dep_DirEle & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_TraAnt = g_rst_Princi!ActEco_Dep_TraAnt
      
      If moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_TraAnt = 1 Then
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_TipDoc_Ant = g_rst_Princi!ActEco_Dep_TipDoc_Ant
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NumDoc_Ant = g_rst_Princi!ActEco_Dep_NumDoc_Ant
         
         'Buscar si empresa ya esta registrada
         g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
         g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_TipDoc_Ant) & " AND "
         g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NumDoc_Ant & "' "
   
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
            Exit Sub
         End If
         
         If g_rst_Genera.BOF And g_rst_Genera.EOF Then
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_FlgEmp_Ant = "9"
         
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_RazSoc_Ant = Trim(g_rst_Princi!ActEco_Dep_RazSoc_Ant & "")
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NomCom_Ant = Trim(g_rst_Princi!ActEco_Dep_NomCom_Ant & "")
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_Telef1_Ant = Trim(g_rst_Princi!ActEco_Dep_Telef1_Ant & "")
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_Telef2_Ant = Trim(g_rst_Princi!ActEco_Dep_Telef2_Ant & "")
         Else
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_FlgEmp_Ant = CStr(g_rst_Genera!DATGEN_CLASIF)
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_RazSoc_Ant = Trim(g_rst_Genera!DATGEN_RAZSOC & "")
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NomCom_Ant = Trim(g_rst_Genera!DATGEN_NOMCOM & "")
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_Telef1_Ant = Trim(g_rst_Genera!DATGEN_TELEF1 & "")
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_Telef2_Ant = Trim(g_rst_Genera!DATGEN_TELEF2 & "")
         End If
         
         g_rst_Genera.Close
         Set g_rst_Genera = Nothing
         
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_FecIng_Ant = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Dep_FecIng_Ant))
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_FecCes_Ant = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Dep_FecCes_Ant))
      End If
      
      'Independiente
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_TipDoc = g_rst_Princi!ActEco_Ind_TipDoc
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_NumDoc = Trim(g_rst_Princi!ActEco_Ind_NumDoc & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_TipVia = g_rst_Princi!ActEco_Ind_TipVia
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_NomVia = Trim(g_rst_Princi!ActEco_Ind_NomVia & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_NumVia = Trim(g_rst_Princi!ActEco_Ind_NumVia & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_IntDpt = Trim(g_rst_Princi!ActEco_Ind_IntDpt & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_TipZon = g_rst_Princi!ActEco_Ind_TipZon
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_NomZon = Trim(g_rst_Princi!ActEco_Ind_NomZon & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_UbiGeo = Trim(g_rst_Princi!ActEco_Ind_UbiGeo & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_Refere = Trim(g_rst_Princi!ActEco_Ind_Refere & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_Telef1 = Trim(g_rst_Princi!ActEco_Ind_Telef1 & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_Telef2 = Trim(g_rst_Princi!ActEco_Ind_Telef2 & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_NumFax = Trim(g_rst_Princi!ActEco_Ind_NumFax & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_CodCiu = g_rst_Princi!ActEco_Ind_CodCiu
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_IngNet = g_rst_Princi!ActEco_Ind_IngNet
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_IniAct = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Ind_IniAct))
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_ConLoc = g_rst_Princi!ActEco_Ind_ConLoc
      
      If moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_ConLoc = 1 Then
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_TipDoc_Emp = g_rst_Princi!ActEco_Ind_TipDoc_Emp
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_NumDoc_Emp = g_rst_Princi!ActEco_Ind_NumDoc_Emp
         
         'Buscar si empresa ya esta registrada
         g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
         g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_TipDoc_Emp) & " AND "
         g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_NumDoc_Emp & "' "
   
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
            Exit Sub
         End If
         
         If g_rst_Genera.BOF And g_rst_Genera.EOF Then
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_FlgEmp = "9"
         
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_RazSoc_Emp = Trim(g_rst_Princi!ActEco_Ind_RazSoc_Emp & "")
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_NomCom_Emp = Trim(g_rst_Princi!ActEco_Ind_NomCom_Emp & "")
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_Telef1_Emp = Trim(g_rst_Princi!ActEco_Ind_Telef1_Emp & "")
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_Telef2_Emp = Trim(g_rst_Princi!ActEco_Ind_Telef2_Emp & "")
         Else
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_FlgEmp = CStr(g_rst_Genera!DATGEN_CLASIF)
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_RazSoc_Emp = Trim(g_rst_Genera!DATGEN_RAZSOC & "")
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_NomCom_Emp = Trim(g_rst_Genera!DATGEN_NOMCOM & "")
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_Telef1_Emp = Trim(g_rst_Genera!DATGEN_TELEF1 & "")
            moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_Telef2_Emp = Trim(g_rst_Genera!DATGEN_TELEF2 & "")
         End If
         
         g_rst_Genera.Close
         Set g_rst_Genera = Nothing
         
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_FecIng_Emp = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Ind_FecIng_Emp))
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_CodCar = Trim(g_rst_Princi!ActEco_Ind_CodCar & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_NomCar = Trim(g_rst_Princi!ActEco_Ind_NomCar & "")
      End If
         
      'Comerciante
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_TipDoc = g_rst_Princi!ActEco_Com_TipDoc
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_NumDoc = Trim(g_rst_Princi!ActEco_Com_NumDoc & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_RazSoc = Trim(g_rst_Princi!ActEco_Com_RazSoc & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_NomCom = Trim(g_rst_Princi!ActEco_Com_NomCom & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_TipVia = g_rst_Princi!ActEco_Com_TipVia
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_NomVia = Trim(g_rst_Princi!ActEco_Com_NomVia & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_NumVia = Trim(g_rst_Princi!ActEco_Com_NumVia & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_IntDpt = Trim(g_rst_Princi!ActEco_Com_IntDpt & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_TipZon = g_rst_Princi!ActEco_Com_TipZon
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_NomZon = Trim(g_rst_Princi!ActEco_Com_NomZon & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_UbiGeo = Trim(g_rst_Princi!ActEco_Com_UbiGeo & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_Refere = Trim(g_rst_Princi!ActEco_Com_Refere & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_Telef1 = Trim(g_rst_Princi!ActEco_Com_Telef1 & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_Telef2 = Trim(g_rst_Princi!ActEco_Com_Telef2 & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_NumFax = Trim(g_rst_Princi!ActEco_Com_NumFax & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_CodCiu = g_rst_Princi!ActEco_Com_CodCiu
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_GirCom = Trim(g_rst_Princi!ActEco_Com_GirCom & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_IngNet = g_rst_Princi!ActEco_Com_IngNet
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_VtaMen = g_rst_Princi!ActEco_Com_VtaMen
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_IniOpe = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Com_IniOpe))
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_CodCar = Trim(g_rst_Princi!ActEco_Com_CodCar & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_NomCar = Trim(g_rst_Princi!ActEco_Com_NomCar & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_RegTri = g_rst_Princi!ActEco_Com_RegTri
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_PorPar = g_rst_Princi!ActEco_Com_PorPar
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_TipLoc = g_rst_Princi!ActEco_Com_TipLoc
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_AlqMen = g_rst_Princi!ActEco_Com_AlqMen
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_NomArr = Trim(g_rst_Princi!ActEco_Com_NomArr & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_TelArr = Trim(g_rst_Princi!ActEco_Com_TelArr & "")
      
      'Accionista
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_TipDoc = g_rst_Princi!ActEco_Acc_TipDoc
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_NumDoc = Trim(g_rst_Princi!ActEco_Com_NumDoc & "")
      
      'Buscar si empresa ya esta registrada
      g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
      g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_TipDoc) & " AND "
      g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_NumDoc & "' "

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
         Exit Sub
      End If
      
      If g_rst_Genera.BOF And g_rst_Genera.EOF Then
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_FlgEmp = "9"
      
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_RazSoc = Trim(g_rst_Princi!ActEco_Acc_RazSoc & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_NomCom = Trim(g_rst_Princi!ActEco_Acc_NomCom & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_TipVia = g_rst_Princi!ActEco_Acc_TipVia
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_NomVia = Trim(g_rst_Princi!ActEco_Acc_NomVia & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_NumVia = Trim(g_rst_Princi!ActEco_Acc_NumVia & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_IntDpt = Trim(g_rst_Princi!ActEco_Acc_IntDpt & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_TipZon = g_rst_Princi!ActEco_Acc_TipZon
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_NomZon = Trim(g_rst_Princi!ActEco_Acc_NomZon & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_UbiGeo = Trim(g_rst_Princi!ActEco_Acc_UbiGeo & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_Refere = Trim(g_rst_Princi!ActEco_Acc_Refere & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_Telef1 = Trim(g_rst_Princi!ActEco_Acc_Telef1 & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_Telef2 = Trim(g_rst_Princi!ActEco_Acc_Telef2 & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_NumFax = Trim(g_rst_Princi!ActEco_Acc_NumFax & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_CodCiu = g_rst_Princi!ActEco_Acc_CodCiu
      Else
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_FlgEmp = CStr(g_rst_Genera!DATGEN_CLASIF)
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_RazSoc = Trim(g_rst_Genera!DATGEN_RAZSOC & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_NomCom = Trim(g_rst_Genera!DATGEN_NOMCOM & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_TipVia = g_rst_Genera!DatGen_TipVia
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_NomVia = Trim(g_rst_Genera!DatGen_NomVia & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_NumVia = Trim(g_rst_Genera!DatGen_numVia & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_IntDpt = Trim(g_rst_Genera!DatGen_IntDpt & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_TipZon = g_rst_Genera!DatGen_TipZon
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_NomZon = Trim(g_rst_Genera!DatGen_NomZon & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_Refere = Trim(g_rst_Genera!DatGen_Refere & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_UbiGeo = Trim(g_rst_Genera!DatGen_Ubigeo & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_Telef1 = Trim(g_rst_Genera!DATGEN_TELEF1 & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_Telef2 = Trim(g_rst_Genera!DATGEN_TELEF2 & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_NumFax = Trim(g_rst_Genera!DatGen_NUMFAX & "")
         moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_CodCiu = g_rst_Genera!DATGEN_CODCIU
      End If
         
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_IngNet = g_rst_Princi!ActEco_Acc_IngNet
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_PorPar = g_rst_Princi!ActEco_Acc_PorPar
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_FecAnt = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Acc_FecAnt))
      
      'Rentista
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_IngNet = g_rst_Princi!ActEco_Ren_IngNet
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_Direc1 = Trim(g_rst_Princi!ActEco_Ren_Direc1 & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_NomAr1 = Trim(g_rst_Princi!ActEco_Ren_NomAr1 & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_IniAl1 = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Ren_IniAl1))
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_Tele11 = Trim(g_rst_Princi!ActEco_Ren_Tele11 & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_Tele21 = Trim(g_rst_Princi!ActEco_Ren_Tele21 & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_AlqMe1 = g_rst_Princi!ActEco_Ren_AlqMe1
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_SegPro = g_rst_Princi!ActEco_Ren_SegPro
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_Direc2 = Trim(g_rst_Princi!ActEco_Ren_Direc2 & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_NomAr2 = Trim(g_rst_Princi!ActEco_Ren_NomAr2 & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_IniAl2 = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Ren_IniAl2))
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_Tele12 = Trim(g_rst_Princi!ActEco_Ren_Tele12 & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_Tele22 = Trim(g_rst_Princi!ActEco_Ren_Tele22 & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_AlqMe2 = g_rst_Princi!ActEco_Ren_AlqMe2
      
      'Otros
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Otr_IngNet = g_rst_Princi!ActEco_Otr_IngNet
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Otr_Activi = Trim(g_rst_Princi!ActEco_Otr_Activi & "")
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Otr_CodCiu = g_rst_Princi!ActEco_Otr_CodCiu
      moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Otr_Observ = Trim(g_rst_Princi!ActEco_Otr_Observ & "")
   
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   End If
End Sub

Private Sub fs_Cargar_ActEco_Cyg(ByVal p_TipDoc As Integer, ByVal p_NumDoc As String, ByVal p_Indice As Integer)
   g_str_Parame = "SELECT * FROM CLI_ACTECO WHERE "
   g_str_Parame = g_str_Parame & "ACTECO_CLITDO = " & CStr(p_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "ACTECO_CLINDO = '" & p_NumDoc & "' AND "
   g_str_Parame = g_str_Parame & "ACTECO_ORDACT = " & CStr(p_Indice) & " "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If

   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_OrdAct = p_Indice
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_TipAct = g_rst_Princi!ActEco_CodAct
   
      'Dependiente
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_SitTra = g_rst_Princi!ActEco_Dep_SitTra
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_TipDoc = g_rst_Princi!ActEco_Dep_TipDoc
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NumDoc = Trim(g_rst_Princi!ActEco_Dep_NumDoc & "")
      
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_TipOfi = g_rst_Princi!ActEco_Dep_TipOfi
      
      'Buscar si empresa ya esta registrada
      g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
      g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_TipDoc) & " AND "
      g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NumDoc & "' "

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
         Exit Sub
      End If
      
      If g_rst_Genera.BOF And g_rst_Genera.EOF Then
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_FlgEmp = "9"
      
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_RazSoc = Trim(g_rst_Princi!ActEco_Dep_RazSoc & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NomCom = Trim(g_rst_Princi!ActEco_Dep_NomCom & "")
      
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_CodCiu = g_rst_Princi!ActEco_Dep_CodCiu
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_TeleRH = Trim(g_rst_Princi!ActEco_Dep_TeleRH & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_AnexRH = Trim(g_rst_Princi!ActEco_Dep_AnexRH & "")
      
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_TipVia = g_rst_Princi!ActEco_Dep_TipVia
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NomVia = Trim(g_rst_Princi!ActEco_Dep_NomVia & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NumVia = Trim(g_rst_Princi!ActEco_Dep_NumVia & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_IntDpt = Trim(g_rst_Princi!ActEco_Dep_IntDpt & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_TipZon = g_rst_Princi!ActEco_Dep_TipZon
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NomZon = Trim(g_rst_Princi!ActEco_Dep_NomZon & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_Refere = Trim(g_rst_Princi!ActEco_Dep_Refere & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_UbiGeo = Trim(g_rst_Princi!ActEco_Dep_UbiGeo & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_Telef1 = Trim(g_rst_Princi!ActEco_Dep_Telef1 & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_Telef2 = Trim(g_rst_Princi!ActEco_Dep_Telef2 & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NumFax = Trim(g_rst_Princi!ActEco_Dep_NumFax & "")
      Else
         g_rst_Genera.MoveFirst
      
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_FlgEmp = CStr(g_rst_Genera!DATGEN_CLASIF)
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_RazSoc = Trim(g_rst_Genera!DATGEN_RAZSOC & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NomCom = Trim(g_rst_Genera!DATGEN_NOMCOM & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_CodCiu = g_rst_Genera!DATGEN_CODCIU
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_TeleRH = Trim(g_rst_Genera!DATGEN_TELERH & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_AnexRH = Trim(g_rst_Genera!DATGEN_ANEXRH & "")
      
         If moddat_g_arr_ActEco_Cyg(moddat_g_int_OrdAct).ActEco_Dep_TipOfi = 1 Then
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_TipVia = g_rst_Genera!DatGen_TipVia
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NomVia = Trim(g_rst_Genera!DatGen_NomVia & "")
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NumVia = Trim(g_rst_Genera!DatGen_numVia & "")
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_IntDpt = Trim(g_rst_Genera!DatGen_IntDpt & "")
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_TipZon = g_rst_Genera!DatGen_TipZon
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NomZon = Trim(g_rst_Genera!DatGen_NomZon & "")
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_Refere = Trim(g_rst_Genera!DatGen_Refere & "")
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_UbiGeo = Trim(g_rst_Genera!DatGen_Ubigeo & "")
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_Telef1 = Trim(g_rst_Genera!DATGEN_TELEF1 & "")
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_Telef2 = Trim(g_rst_Genera!DATGEN_TELEF2 & "")
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NumFax = Trim(g_rst_Genera!DatGen_NUMFAX & "")
         Else
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_TipVia = g_rst_Princi!ActEco_Dep_TipVia
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NomVia = Trim(g_rst_Princi!ActEco_Dep_NomVia & "")
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NumVia = Trim(g_rst_Princi!ActEco_Dep_NumVia & "")
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_IntDpt = Trim(g_rst_Princi!ActEco_Dep_IntDpt & "")
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_TipZon = g_rst_Princi!ActEco_Dep_TipZon
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NomZon = Trim(g_rst_Princi!ActEco_Dep_NomZon & "")
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_Refere = Trim(g_rst_Princi!ActEco_Dep_Refere & "")
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_UbiGeo = Trim(g_rst_Princi!ActEco_Dep_UbiGeo & "")
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_Telef1 = Trim(g_rst_Princi!ActEco_Dep_Telef1 & "")
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_Telef2 = Trim(g_rst_Princi!ActEco_Dep_Telef2 & "")
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NumFax = Trim(g_rst_Princi!ActEco_Dep_NumFax & "")
         End If
      End If
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_IngNet = g_rst_Princi!ActEco_Dep_IngNet
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_FreHab = g_rst_Princi!ActEco_Dep_FreHab
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_FecIng = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Dep_FecIng))
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_CodCar = Trim(g_rst_Princi!ActEco_Dep_CodCar & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NomCar = Trim(g_rst_Princi!ActEco_Dep_NomCar & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NomAre = Trim(g_rst_Princi!ActEco_Dep_NomAre & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NumAnx = Trim(g_rst_Princi!ActEco_Dep_NumAnx & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_TelDir = Trim(g_rst_Princi!ActEco_Dep_TelDir & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_Celula = Trim(g_rst_Princi!ActEco_Dep_Celula & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_DirEle = Trim(g_rst_Princi!ActEco_Dep_DirEle & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_TraAnt = g_rst_Princi!ActEco_Dep_TraAnt
      
      If moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_TraAnt = 1 Then
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_TipDoc_Ant = g_rst_Princi!ActEco_Dep_TipDoc_Ant
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NumDoc_Ant = g_rst_Princi!ActEco_Dep_NumDoc_Ant
         
         'Buscar si empresa ya esta registrada
         g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
         g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_TipDoc_Ant) & " AND "
         g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NumDoc_Ant & "' "
   
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
            Exit Sub
         End If
         
         If g_rst_Genera.BOF And g_rst_Genera.EOF Then
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_FlgEmp_Ant = "9"
         
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_RazSoc_Ant = Trim(g_rst_Princi!ActEco_Dep_RazSoc_Ant & "")
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NomCom_Ant = Trim(g_rst_Princi!ActEco_Dep_NomCom_Ant & "")
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_Telef1_Ant = Trim(g_rst_Princi!ActEco_Dep_Telef1_Ant & "")
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_Telef2_Ant = Trim(g_rst_Princi!ActEco_Dep_Telef2_Ant & "")
         Else
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_FlgEmp_Ant = CStr(g_rst_Genera!DATGEN_CLASIF)
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_RazSoc_Ant = Trim(g_rst_Genera!DATGEN_RAZSOC & "")
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NomCom_Ant = Trim(g_rst_Genera!DATGEN_NOMCOM & "")
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_Telef1_Ant = Trim(g_rst_Genera!DATGEN_TELEF1 & "")
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_Telef2_Ant = Trim(g_rst_Genera!DATGEN_TELEF2 & "")
         End If
         
         g_rst_Genera.Close
         Set g_rst_Genera = Nothing
         
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_FecIng_Ant = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Dep_FecIng_Ant))
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_FecCes_Ant = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Dep_FecCes_Ant))
      End If
      
      'Independiente
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_TipDoc = g_rst_Princi!ActEco_Ind_TipDoc
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_NumDoc = Trim(g_rst_Princi!ActEco_Ind_NumDoc & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_TipVia = g_rst_Princi!ActEco_Ind_TipVia
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_NomVia = Trim(g_rst_Princi!ActEco_Ind_NomVia & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_NumVia = Trim(g_rst_Princi!ActEco_Ind_NumVia & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_IntDpt = Trim(g_rst_Princi!ActEco_Ind_IntDpt & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_TipZon = g_rst_Princi!ActEco_Ind_TipZon
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_NomZon = Trim(g_rst_Princi!ActEco_Ind_NomZon & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_UbiGeo = Trim(g_rst_Princi!ActEco_Ind_UbiGeo & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_Refere = Trim(g_rst_Princi!ActEco_Ind_Refere & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_Telef1 = Trim(g_rst_Princi!ActEco_Ind_Telef1 & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_Telef2 = Trim(g_rst_Princi!ActEco_Ind_Telef2 & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_NumFax = Trim(g_rst_Princi!ActEco_Ind_NumFax & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_CodCiu = g_rst_Princi!ActEco_Ind_CodCiu
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_IngNet = g_rst_Princi!ActEco_Ind_IngNet
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_IniAct = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Ind_IniAct))
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_ConLoc = g_rst_Princi!ActEco_Ind_ConLoc
      
      If moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_ConLoc = 1 Then
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_TipDoc_Emp = g_rst_Princi!ActEco_Ind_TipDoc_Emp
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_NumDoc_Emp = g_rst_Princi!ActEco_Ind_NumDoc_Emp
         
         'Buscar si empresa ya esta registrada
         g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
         g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_TipDoc_Emp) & " AND "
         g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_NumDoc_Emp & "' "
   
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
            Exit Sub
         End If
         
         If g_rst_Genera.BOF And g_rst_Genera.EOF Then
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_FlgEmp = "9"
         
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_RazSoc_Emp = Trim(g_rst_Princi!ActEco_Ind_RazSoc_Emp & "")
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_NomCom_Emp = Trim(g_rst_Princi!ActEco_Ind_NomCom_Emp & "")
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_Telef1_Emp = Trim(g_rst_Princi!ActEco_Ind_Telef1_Emp & "")
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_Telef2_Emp = Trim(g_rst_Princi!ActEco_Ind_Telef2_Emp & "")
         Else
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_FlgEmp = CStr(g_rst_Genera!DATGEN_CLASIF)
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_RazSoc_Emp = Trim(g_rst_Genera!DATGEN_RAZSOC & "")
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_NomCom_Emp = Trim(g_rst_Genera!DATGEN_NOMCOM & "")
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_Telef1_Emp = Trim(g_rst_Genera!DATGEN_TELEF1 & "")
            moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_Telef2_Emp = Trim(g_rst_Genera!DATGEN_TELEF2 & "")
         End If
         
         g_rst_Genera.Close
         Set g_rst_Genera = Nothing
         
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_FecIng_Emp = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Ind_FecIng_Emp))
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_CodCar = Trim(g_rst_Princi!ActEco_Ind_CodCar & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_NomCar = Trim(g_rst_Princi!ActEco_Ind_NomCar & "")
      End If
         
      'Comerciante
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_TipDoc = g_rst_Princi!ActEco_Com_TipDoc
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_NumDoc = Trim(g_rst_Princi!ActEco_Com_NumDoc & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_RazSoc = Trim(g_rst_Princi!ActEco_Com_RazSoc & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_NomCom = Trim(g_rst_Princi!ActEco_Com_NomCom & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_TipVia = g_rst_Princi!ActEco_Com_TipVia
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_NomVia = Trim(g_rst_Princi!ActEco_Com_NomVia & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_NumVia = Trim(g_rst_Princi!ActEco_Com_NumVia & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_IntDpt = Trim(g_rst_Princi!ActEco_Com_IntDpt & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_TipZon = g_rst_Princi!ActEco_Com_TipZon
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_NomZon = Trim(g_rst_Princi!ActEco_Com_NomZon & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_UbiGeo = Trim(g_rst_Princi!ActEco_Com_UbiGeo & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_Refere = Trim(g_rst_Princi!ActEco_Com_Refere & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_Telef1 = Trim(g_rst_Princi!ActEco_Com_Telef1 & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_Telef2 = Trim(g_rst_Princi!ActEco_Com_Telef2 & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_NumFax = Trim(g_rst_Princi!ActEco_Com_NumFax & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_CodCiu = g_rst_Princi!ActEco_Com_CodCiu
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_GirCom = Trim(g_rst_Princi!ActEco_Com_GirCom & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_IngNet = g_rst_Princi!ActEco_Com_IngNet
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_VtaMen = g_rst_Princi!ActEco_Com_VtaMen
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_IniOpe = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Com_IniOpe))
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_CodCar = Trim(g_rst_Princi!ActEco_Com_CodCar & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_NomCar = Trim(g_rst_Princi!ActEco_Com_NomCar & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_RegTri = g_rst_Princi!ActEco_Com_RegTri
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_PorPar = g_rst_Princi!ActEco_Com_PorPar
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_TipLoc = g_rst_Princi!ActEco_Com_TipLoc
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_AlqMen = g_rst_Princi!ActEco_Com_AlqMen
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_NomArr = Trim(g_rst_Princi!ActEco_Com_NomArr & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_TelArr = Trim(g_rst_Princi!ActEco_Com_TelArr & "")
      
      'Accionista
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_TipDoc = g_rst_Princi!ActEco_Acc_TipDoc
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_NumDoc = Trim(g_rst_Princi!ActEco_Acc_NumDoc & "")
      
      'Buscar si empresa ya esta registrada
      g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
      g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_TipDoc) & " AND "
      g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_NumDoc & "' "

      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
         Exit Sub
      End If
      
      If g_rst_Genera.BOF And g_rst_Genera.EOF Then
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_FlgEmp = "9"
      
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_RazSoc = Trim(g_rst_Princi!ActEco_Acc_RazSoc & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_NomCom = Trim(g_rst_Princi!ActEco_Acc_NomCom & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_TipVia = g_rst_Princi!ActEco_Acc_TipVia
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_NomVia = Trim(g_rst_Princi!ActEco_Acc_NomVia & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_NumVia = Trim(g_rst_Princi!ActEco_Acc_NumVia & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_IntDpt = Trim(g_rst_Princi!ActEco_Acc_IntDpt & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_TipZon = g_rst_Princi!ActEco_Acc_TipZon
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_NomZon = Trim(g_rst_Princi!ActEco_Acc_NomZon & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_UbiGeo = Trim(g_rst_Princi!ActEco_Acc_UbiGeo & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_Refere = Trim(g_rst_Princi!ActEco_Acc_Refere & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_Telef1 = Trim(g_rst_Princi!ActEco_Acc_Telef1 & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_Telef2 = Trim(g_rst_Princi!ActEco_Acc_Telef2 & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_NumFax = Trim(g_rst_Princi!ActEco_Acc_NumFax & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_CodCiu = g_rst_Princi!ActEco_Acc_CodCiu
      Else
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_FlgEmp = CStr(g_rst_Genera!DATGEN_CLASIF)
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_RazSoc = Trim(g_rst_Genera!DATGEN_RAZSOC & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_NomCom = Trim(g_rst_Genera!DATGEN_NOMCOM & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_TipVia = g_rst_Genera!DatGen_TipVia
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_NomVia = Trim(g_rst_Genera!DatGen_NomVia & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_NumVia = Trim(g_rst_Genera!DatGen_numVia & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_IntDpt = Trim(g_rst_Genera!DatGen_IntDpt & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_TipZon = g_rst_Genera!DatGen_TipZon
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_NomZon = Trim(g_rst_Genera!DatGen_NomZon & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_Refere = Trim(g_rst_Genera!DatGen_Refere & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_UbiGeo = Trim(g_rst_Genera!DatGen_Ubigeo & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_Telef1 = Trim(g_rst_Genera!DATGEN_TELEF1 & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_Telef2 = Trim(g_rst_Genera!DATGEN_TELEF2 & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_NumFax = Trim(g_rst_Genera!DatGen_NUMFAX & "")
         moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_CodCiu = g_rst_Genera!DATGEN_CODCIU
      End If
         
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_IngNet = g_rst_Princi!ActEco_Acc_IngNet
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_PorPar = g_rst_Princi!ActEco_Acc_PorPar
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_FecAnt = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Acc_FecAnt))
      
      'Rentista
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_IngNet = g_rst_Princi!ActEco_Ren_IngNet
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_Direc1 = Trim(g_rst_Princi!ActEco_Ren_Direc1 & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_NomAr1 = Trim(g_rst_Princi!ActEco_Ren_NomAr1 & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_IniAl1 = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Ren_IniAl1))
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_Tele11 = Trim(g_rst_Princi!ActEco_Ren_Tele11 & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_Tele21 = Trim(g_rst_Princi!ActEco_Ren_Tele21 & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_AlqMe1 = g_rst_Princi!ActEco_Ren_AlqMe1
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_SegPro = g_rst_Princi!ActEco_Ren_SegPro
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_Direc2 = Trim(g_rst_Princi!ActEco_Ren_Direc2 & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_NomAr2 = Trim(g_rst_Princi!ActEco_Ren_NomAr2 & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_IniAl2 = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Ren_IniAl2))
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_Tele12 = Trim(g_rst_Princi!ActEco_Ren_Tele12 & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_Tele22 = Trim(g_rst_Princi!ActEco_Ren_Tele22 & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_AlqMe2 = g_rst_Princi!ActEco_Ren_AlqMe2
   
      'Otros
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Otr_IngNet = g_rst_Princi!ActEco_Otr_IngNet
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Otr_Activi = Trim(g_rst_Princi!ActEco_Otr_Activi & "")
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Otr_CodCiu = g_rst_Princi!ActEco_Otr_CodCiu
      moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Otr_Observ = Trim(g_rst_Princi!ActEco_Otr_Observ & "")
   
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   End If
End Sub

Private Sub fs_Grabar_ActEco_Tit(ByVal p_Indice As Integer)
   'Grabando Información de Actividad Económica
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      Screen.MousePointer = 11
      
      g_str_Parame = "USP_TRA_CLIACT ("
      
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumSol & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_int_TipDoc) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumDoc & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_OrdAct) & ", "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_TipAct) & ", "
      
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_TipDoc) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NumDoc & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_RazSoc & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NomCom & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_TipOfi) & ", "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_SitTra) & ", "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_TipVia) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NomVia & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NumVia & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_IntDpt & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_TipZon) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NomZon & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_UbiGeo & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_Refere & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_Telef1 & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_Telef2 & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NumFax & "', "
      g_str_Parame = g_str_Parame & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_CodCiu & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_TeleRH & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_AnexRH & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_IngNet) & ", "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_FreHab) & ", "
      
      If Len(Trim(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_FecIng)) > 0 Then
         g_str_Parame = g_str_Parame & Format(CDate(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_FecIng), "yyyymmdd") & ", "
      Else
         g_str_Parame = g_str_Parame & "0, "
      End If
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_CodCar & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NomCar & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NomAre & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NumAnx & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_TelDir & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_Celula & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_DirEle & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_TraAnt) & ", "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_TipDoc_Ant) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NumDoc_Ant & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_RazSoc_Ant & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_NomCom_Ant & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_Telef1_Ant & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_Telef2_Ant & "', "
      
      If Len(Trim(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_FecIng_Ant)) > 0 Then
         g_str_Parame = g_str_Parame & Format(CDate(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_FecIng_Ant), "yyyymmdd") & ", "
      Else
         g_str_Parame = g_str_Parame & "0, "
      End If
      
      If Len(Trim(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_FecCes_Ant)) > 0 Then
         g_str_Parame = g_str_Parame & Format(CDate(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Dep_FecCes_Ant), "yyyymmdd") & ", "
      Else
         g_str_Parame = g_str_Parame & "0, "
      End If
      
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_TipDoc) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_NumDoc & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_TipVia) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_NomVia & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_NumVia & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_IntDpt & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_TipZon) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_NomZon & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_UbiGeo & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_Refere & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_Telef1 & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_Telef2 & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_NumFax & "', "
      g_str_Parame = g_str_Parame & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_CodCiu & ", "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_IngNet) & ", "
      
      If Len(Trim(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_IniAct)) > 0 Then
         g_str_Parame = g_str_Parame & Format(CDate(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_IniAct), "yyyymmdd") & ", "
      Else
         g_str_Parame = g_str_Parame & "0, "
      End If
      
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_ConLoc) & ", "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_TipDoc_Emp) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_NumDoc_Emp & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_RazSoc_Emp & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_NomCom_Emp & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_Telef1_Emp & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_Telef2_Emp & "', "
      
      If Len(Trim(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_FecIng_Emp)) > 0 Then
         g_str_Parame = g_str_Parame & Format(CDate(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_FecIng_Emp), "yyyymmdd") & ", "
      Else
         g_str_Parame = g_str_Parame & "0, "
      End If
      
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_CodCar & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ind_NomCar & "', "
      
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_TipDoc) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_NumDoc & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_RazSoc & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_NomCom & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_TipVia) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_NomVia & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_NumVia & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_IntDpt & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_TipZon) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_NomZon & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_UbiGeo & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_Refere & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_Telef1 & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_Telef2 & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_NumFax & "', "
      g_str_Parame = g_str_Parame & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_CodCiu & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_GirCom & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_IngNet) & ", "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_VtaMen) & ", "
      
      If Len(Trim(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_IniOpe)) > 0 Then
         g_str_Parame = g_str_Parame & Format(CDate(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_IniOpe), "yyyymmdd") & ", "
      Else
         g_str_Parame = g_str_Parame & "0, "
      End If
      
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_CodCar & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_NomCar & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_RegTri) & ", "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_PorPar) & ", "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_TipLoc) & ", "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_AlqMen) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_NomArr & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Com_TelArr & "', "
      
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_TipDoc) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_NumDoc & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_RazSoc & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_NomCom & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_TipVia) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_NomVia & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_NumVia & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_IntDpt & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_TipZon) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_NomZon & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_UbiGeo & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_Refere & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_Telef1 & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_Telef2 & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_NumFax & "', "
      g_str_Parame = g_str_Parame & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_CodCiu & ", "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_IngNet) & ", "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_PorPar) & ", "
      
      If Len(Trim(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_FecAnt)) > 0 Then
         g_str_Parame = g_str_Parame & Format(CDate(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Acc_FecAnt), "yyyymmdd") & ", "
      Else
         g_str_Parame = g_str_Parame & "0, "
      End If
      
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_IngNet) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_Direc1 & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_NomAr1 & "', "
      
      If Len(Trim(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_IniAl1)) > 0 Then
         g_str_Parame = g_str_Parame & Format(CDate(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_IniAl1), "yyyymmdd") & ", "
      Else
         g_str_Parame = g_str_Parame & "0, "
      End If
      
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_Tele11 & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_Tele21 & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_AlqMe1) & ", "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_SegPro) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_Direc2 & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_NomAr2 & "', "
      
      If Len(Trim(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_IniAl2)) > 0 Then
         g_str_Parame = g_str_Parame & Format(CDate(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_IniAl2), "yyyymmdd") & ", "
      Else
         g_str_Parame = g_str_Parame & "0, "
      End If
      
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_Tele12 & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_Tele22 & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Ren_AlqMe2) & ", "
      
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Otr_IngNet) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Otr_Activi & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Otr_CodCiu) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Tit(p_Indice).ActEco_Otr_Observ & "', "
      
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If
      
      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar la grabación de los datos. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
      
      Screen.MousePointer = 0
   Loop
End Sub

Private Sub fs_Grabar_ActEco_Cyg(ByVal p_Indice As Integer)
   'Grabando Información de Actividad Económica
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      Screen.MousePointer = 11
      
      g_str_Parame = "USP_TRA_CLIACT ("
      
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumSol & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_int_CygTDo) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_CygNDo & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_OrdAct) & ", "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_TipAct) & ", "
      
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_TipDoc) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NumDoc & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_RazSoc & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NomCom & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_TipOfi) & ", "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_SitTra) & ", "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_TipVia) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NomVia & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NumVia & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_IntDpt & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_TipZon) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NomZon & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_UbiGeo & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_Refere & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_Telef1 & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_Telef2 & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NumFax & "', "
      g_str_Parame = g_str_Parame & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_CodCiu & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_TeleRH & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_AnexRH & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_IngNet) & ", "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_FreHab) & ", "
      
      If Len(Trim(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_FecIng)) > 0 Then
         g_str_Parame = g_str_Parame & Format(CDate(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_FecIng), "yyyymmdd") & ", "
      Else
         g_str_Parame = g_str_Parame & "0, "
      End If
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_CodCar & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NomCar & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NomAre & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NumAnx & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_TelDir & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_Celula & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_DirEle & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_TraAnt) & ", "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_TipDoc_Ant) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NumDoc_Ant & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_RazSoc_Ant & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_NomCom_Ant & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_Telef1_Ant & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_Telef2_Ant & "', "
      
      If Len(Trim(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_FecIng_Ant)) > 0 Then
         g_str_Parame = g_str_Parame & Format(CDate(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_FecIng_Ant), "yyyymmdd") & ", "
      Else
         g_str_Parame = g_str_Parame & "0, "
      End If
      
      If Len(Trim(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_FecCes_Ant)) > 0 Then
         g_str_Parame = g_str_Parame & Format(CDate(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Dep_FecCes_Ant), "yyyymmdd") & ", "
      Else
         g_str_Parame = g_str_Parame & "0, "
      End If
      
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_TipDoc) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_NumDoc & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_TipVia) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_NomVia & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_NumVia & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_IntDpt & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_TipZon) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_NomZon & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_UbiGeo & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_Refere & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_Telef1 & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_Telef2 & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_NumFax & "', "
      g_str_Parame = g_str_Parame & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_CodCiu & ", "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_IngNet) & ", "
      
      If Len(Trim(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_IniAct)) > 0 Then
         g_str_Parame = g_str_Parame & Format(CDate(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_IniAct), "yyyymmdd") & ", "
      Else
         g_str_Parame = g_str_Parame & "0, "
      End If
      
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_ConLoc) & ", "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_TipDoc_Emp) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_NumDoc_Emp & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_RazSoc_Emp & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_NomCom_Emp & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_Telef1_Emp & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_Telef2_Emp & "', "
      
      If Len(Trim(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_FecIng_Emp)) > 0 Then
         g_str_Parame = g_str_Parame & Format(CDate(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_FecIng_Emp), "yyyymmdd") & ", "
      Else
         g_str_Parame = g_str_Parame & "0, "
      End If
      
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_CodCar & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ind_NomCar & "', "
      
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_TipDoc) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_NumDoc & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_RazSoc & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_NomCom & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_TipVia) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_NomVia & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_NumVia & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_IntDpt & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_TipZon) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_NomZon & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_UbiGeo & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_Refere & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_Telef1 & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_Telef2 & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_NumFax & "', "
      g_str_Parame = g_str_Parame & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_CodCiu & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_GirCom & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_IngNet) & ", "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_VtaMen) & ", "
      
      If Len(Trim(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_IniOpe)) > 0 Then
         g_str_Parame = g_str_Parame & Format(CDate(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_IniOpe), "yyyymmdd") & ", "
      Else
         g_str_Parame = g_str_Parame & "0, "
      End If
      
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_CodCar & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_NomCar & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_RegTri) & ", "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_PorPar) & ", "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_TipLoc) & ", "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_AlqMen) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_NomArr & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Com_TelArr & "', "
      
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_TipDoc) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_NumDoc & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_RazSoc & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_NomCom & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_TipVia) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_NomVia & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_NumVia & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_IntDpt & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_TipZon) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_NomZon & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_UbiGeo & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_Refere & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_Telef1 & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_Telef2 & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_NumFax & "', "
      g_str_Parame = g_str_Parame & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_CodCiu & ", "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_IngNet) & ", "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_PorPar) & ", "
      
      If Len(Trim(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_FecAnt)) > 0 Then
         g_str_Parame = g_str_Parame & Format(CDate(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Acc_FecAnt), "yyyymmdd") & ", "
      Else
         g_str_Parame = g_str_Parame & "0, "
      End If
      
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_IngNet) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_Direc1 & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_NomAr1 & "', "
      
      If Len(Trim(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_IniAl1)) > 0 Then
         g_str_Parame = g_str_Parame & Format(CDate(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_IniAl1), "yyyymmdd") & ", "
      Else
         g_str_Parame = g_str_Parame & "0, "
      End If
      
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_Tele11 & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_Tele21 & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_AlqMe1) & ", "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_SegPro) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_Direc2 & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_NomAr2 & "', "
      
      If Len(Trim(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_IniAl2)) > 0 Then
         g_str_Parame = g_str_Parame & Format(CDate(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_IniAl2), "yyyymmdd") & ", "
      Else
         g_str_Parame = g_str_Parame & "0, "
      End If
      
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_Tele12 & "', "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_Tele22 & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Ren_AlqMe2) & ", "
      
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Otr_IngNet) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Otr_Activi & "', "
      g_str_Parame = g_str_Parame & CStr(moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Otr_CodCiu) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_arr_ActEco_Cyg(p_Indice).ActEco_Otr_Observ & "', "
      
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If
      
      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar la grabación de los datos. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
      
      Screen.MousePointer = 0
   Loop
End Sub


