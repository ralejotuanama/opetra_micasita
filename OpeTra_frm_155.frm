VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frm_ModSol_06 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   9180
   ClientLeft      =   2910
   ClientTop       =   1185
   ClientWidth     =   11250
   Icon            =   "OpeTra_frm_155.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9180
   ScaleWidth      =   11250
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   9180
      Left            =   0
      TabIndex        =   50
      Top             =   0
      Width           =   11250
      _Version        =   65536
      _ExtentX        =   19844
      _ExtentY        =   16192
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
      Begin Threed.SSPanel SSPanel8 
         Height          =   1545
         Left            =   30
         TabIndex        =   51
         Top             =   7590
         Width           =   11175
         _Version        =   65536
         _ExtentX        =   19711
         _ExtentY        =   2725
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
         Begin Threed.SSPanel pnl_AreTer 
            Height          =   315
            Left            =   1860
            TabIndex        =   52
            Top             =   390
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_AreCon 
            Height          =   315
            Left            =   6030
            TabIndex        =   53
            Top             =   390
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel SSPanel16 
            Height          =   60
            Left            =   30
            TabIndex        =   54
            Top             =   750
            Width           =   11115
            _Version        =   65536
            _ExtentX        =   19606
            _ExtentY        =   106
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
         End
         Begin Threed.SSPanel pnl_SumAse 
            Height          =   315
            Left            =   1860
            TabIndex        =   55
            Top             =   840
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
            Left            =   6030
            TabIndex        =   56
            Top             =   840
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_ValRea 
            Height          =   315
            Left            =   9780
            TabIndex        =   57
            Top             =   840
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_ValTer 
            Height          =   315
            Left            =   1860
            TabIndex        =   58
            Top             =   1170
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_ValEdi 
            Height          =   315
            Left            =   6030
            TabIndex        =   59
            Top             =   1170
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_ValACo 
            Height          =   315
            Left            =   9780
            TabIndex        =   60
            Top             =   1170
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
            Alignment       =   4
         End
         Begin VB.Label Label56 
            Caption         =   "Valor Areas Comunes:"
            Height          =   315
            Left            =   7920
            TabIndex        =   69
            Top             =   1170
            Width           =   1815
         End
         Begin VB.Label Label55 
            Caption         =   "Valor Edificación:"
            Height          =   315
            Left            =   4170
            TabIndex        =   68
            Top             =   1170
            Width           =   1485
         End
         Begin VB.Label Label54 
            Caption         =   "Valor Terreno:"
            Height          =   315
            Left            =   90
            TabIndex        =   67
            Top             =   1170
            Width           =   1485
         End
         Begin VB.Label Label53 
            Caption         =   "Valor Realización:"
            Height          =   315
            Left            =   7920
            TabIndex        =   66
            Top             =   840
            Width           =   1485
         End
         Begin VB.Label Label52 
            Caption         =   "Valor Comercial:"
            Height          =   315
            Left            =   4170
            TabIndex        =   65
            Top             =   840
            Width           =   1485
         End
         Begin VB.Label Label51 
            Caption         =   "Suma Asegurada:"
            Height          =   315
            Left            =   90
            TabIndex        =   64
            Top             =   840
            Width           =   1485
         End
         Begin VB.Label Label50 
            Caption         =   "Area Construcción:"
            Height          =   315
            Left            =   4170
            TabIndex        =   63
            Top             =   390
            Width           =   1485
         End
         Begin VB.Label Label49 
            Caption         =   "Totales"
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
            TabIndex        =   62
            Top             =   60
            Width           =   1065
         End
         Begin VB.Label Label11 
            Caption         =   "Area Terreno:"
            Height          =   315
            Left            =   90
            TabIndex        =   61
            Top             =   390
            Width           =   1185
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   2745
         Left            =   30
         TabIndex        =   70
         Top             =   2250
         Width           =   11175
         _Version        =   65536
         _ExtentX        =   19711
         _ExtentY        =   4842
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
         Begin VB.ComboBox cmb_MatCon 
            Height          =   315
            Left            =   1860
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   2040
            Width           =   9255
         End
         Begin VB.ComboBox cmb_UsoInm 
            Height          =   315
            Left            =   8340
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   1710
            Width           =   2775
         End
         Begin VB.ComboBox cmb_TipInm 
            Height          =   315
            Left            =   1860
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   1710
            Width           =   3315
         End
         Begin VB.ComboBox cmb_TipMon 
            Height          =   315
            Left            =   1860
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   2370
            Width           =   3315
         End
         Begin VB.TextBox txt_NumInf 
            Height          =   315
            Left            =   1860
            MaxLength       =   25
            TabIndex        =   2
            Text            =   "Text1"
            Top             =   720
            Width           =   1635
         End
         Begin VB.ComboBox cmb_EmpPer 
            Height          =   315
            Left            =   1860
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   9255
         End
         Begin VB.ComboBox cmb_PerTas 
            Height          =   315
            Left            =   1860
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   390
            Width           =   9255
         End
         Begin EditLib.fpDateTime ipp_FecEva 
            Height          =   315
            Left            =   8340
            TabIndex        =   3
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
         Begin EditLib.fpDoubleSingle ipp_TipCam 
            Height          =   315
            Left            =   8340
            TabIndex        =   11
            Top             =   2370
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
            Text            =   "0.0000"
            DecimalPlaces   =   4
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
         Begin EditLib.fpDoubleSingle ipp_AnoCon 
            Height          =   315
            Left            =   1860
            TabIndex        =   4
            Top             =   1050
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
         Begin EditLib.fpDoubleSingle ipp_NumPis 
            Height          =   315
            Left            =   1860
            TabIndex        =   5
            Top             =   1380
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
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "99"
            MinValue        =   "1"
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
         Begin EditLib.fpDoubleSingle ipp_NumSot 
            Height          =   315
            Left            =   8340
            TabIndex        =   6
            Top             =   1380
            Width           =   765
            _Version        =   196608
            _ExtentX        =   1349
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
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "99"
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
         Begin VB.Label Label62 
            Caption         =   "Material Construcción:"
            Height          =   315
            Left            =   60
            TabIndex        =   82
            Top             =   2040
            Width           =   1635
         End
         Begin VB.Label Label61 
            Caption         =   "Uso Inmueble:"
            Height          =   315
            Left            =   6780
            TabIndex        =   81
            Top             =   1710
            Width           =   1065
         End
         Begin VB.Label Label60 
            Caption         =   "Tipo Inmueble:"
            Height          =   315
            Left            =   60
            TabIndex        =   80
            Top             =   1710
            Width           =   1065
         End
         Begin VB.Label Label59 
            Caption         =   "Nro. Sótanos:"
            Height          =   285
            Left            =   6780
            TabIndex        =   79
            Top             =   1380
            Width           =   1485
         End
         Begin VB.Label Label58 
            Caption         =   "Nro. Pisos:"
            Height          =   285
            Left            =   60
            TabIndex        =   78
            Top             =   1380
            Width           =   1485
         End
         Begin VB.Label Label57 
            Caption         =   "Año Construcción:"
            Height          =   285
            Left            =   60
            TabIndex        =   77
            Top             =   1050
            Width           =   1485
         End
         Begin VB.Label Label14 
            Caption         =   "Tipo de Cambio:"
            Height          =   315
            Left            =   6780
            TabIndex        =   76
            Top             =   2370
            Width           =   1365
         End
         Begin VB.Label Label13 
            Caption         =   "Moneda:"
            Height          =   315
            Left            =   60
            TabIndex        =   75
            Top             =   2370
            Width           =   1065
         End
         Begin VB.Label Label12 
            Caption         =   "F. Evaluación:"
            Height          =   285
            Left            =   6780
            TabIndex        =   74
            Top             =   720
            Width           =   1485
         End
         Begin VB.Label Label8 
            Caption         =   "Número Informe:"
            Height          =   285
            Left            =   60
            TabIndex        =   73
            Top             =   720
            Width           =   1485
         End
         Begin VB.Label Label5 
            Caption         =   "Empresa Peritaje:"
            Height          =   285
            Left            =   60
            TabIndex        =   72
            Top             =   60
            Width           =   1395
         End
         Begin VB.Label Label4 
            Caption         =   "Perito Tasador:"
            Height          =   285
            Left            =   60
            TabIndex        =   71
            Top             =   390
            Width           =   1395
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   83
         Top             =   30
         Width           =   11175
         _Version        =   65536
         _ExtentX        =   19711
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
         Begin Threed.SSPanel pnl_TitPri 
            Height          =   315
            Left            =   630
            TabIndex        =   136
            Top             =   60
            Width           =   8565
            _Version        =   65536
            _ExtentX        =   15108
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Modificación de Solicitud de Crédito Hipotecario"
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
         Begin Threed.SSPanel pnl_TitSec 
            Height          =   315
            Left            =   630
            TabIndex        =   137
            Top             =   300
            Width           =   8565
            _Version        =   65536
            _ExtentX        =   15108
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Modificación de Tasación"
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
            Picture         =   "OpeTra_frm_155.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   30
         TabIndex        =   84
         Top             =   1440
         Width           =   11175
         _Version        =   65536
         _ExtentX        =   19711
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
         Begin Threed.SSPanel pnl_NumSol 
            Height          =   315
            Left            =   1860
            TabIndex        =   85
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
         End
         Begin Threed.SSPanel pnl_FecIng 
            Height          =   315
            Left            =   9690
            TabIndex        =   86
            Top             =   60
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2514
            _ExtentY        =   556
            _StockProps     =   15
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
         Begin Threed.SSPanel pnl_Client 
            Height          =   315
            Left            =   1860
            TabIndex        =   87
            Top             =   390
            Width           =   9255
            _Version        =   65536
            _ExtentX        =   16325
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "1-07521154 / IKEHARA PUNK MIGUEL ANGEL"
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
         Begin VB.Label Label2 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   60
            TabIndex        =   90
            Top             =   390
            Width           =   1125
         End
         Begin VB.Label Label3 
            Caption         =   "F. Ingreso Solicitud:"
            Height          =   315
            Left            =   8040
            TabIndex        =   89
            Top             =   60
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Nro. Solicitud"
            Height          =   315
            Left            =   60
            TabIndex        =   88
            Top             =   60
            Width           =   1335
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   91
         Top             =   750
         Width           =   11175
         _Version        =   65536
         _ExtentX        =   19711
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
         Begin VB.CommandButton cmd_Export 
            Height          =   585
            Left            =   600
            Picture         =   "OpeTra_frm_155.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   48
            ToolTipText     =   "Exportar datos a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10560
            Picture         =   "OpeTra_frm_155.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   49
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_155.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   47
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   2505
         Left            =   30
         TabIndex        =   92
         Top             =   5040
         Width           =   11175
         _Version        =   65536
         _ExtentX        =   19711
         _ExtentY        =   4419
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
         Begin TabDlg.SSTab tab_Genera 
            Height          =   2385
            Left            =   60
            TabIndex        =   93
            Top             =   60
            Width           =   11055
            _ExtentX        =   19500
            _ExtentY        =   4207
            _Version        =   393216
            Style           =   1
            Tabs            =   4
            Tab             =   3
            TabsPerRow      =   4
            TabHeight       =   520
            TabCaption(0)   =   "Inmueble"
            TabPicture(0)   =   "OpeTra_frm_155.frx":0EA4
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "Label22"
            Tab(0).Control(1)=   "Label23"
            Tab(0).Control(2)=   "Label15"
            Tab(0).Control(3)=   "Label16"
            Tab(0).Control(4)=   "Label17"
            Tab(0).Control(5)=   "Label18"
            Tab(0).Control(6)=   "Label19"
            Tab(0).Control(7)=   "Label20"
            Tab(0).Control(8)=   "Label6"
            Tab(0).Control(9)=   "ipp_ValRea_Inm"
            Tab(0).Control(10)=   "ipp_ValCom_Inm"
            Tab(0).Control(11)=   "ipp_ValACo_Inm"
            Tab(0).Control(12)=   "ipp_ValEdi_Inm"
            Tab(0).Control(13)=   "ipp_ValTer_Inm"
            Tab(0).Control(14)=   "ipp_SumAse_Inm"
            Tab(0).Control(15)=   "ipp_AreCon_Inm"
            Tab(0).Control(16)=   "ipp_AreTer_Inm"
            Tab(0).Control(17)=   "SSPanel5"
            Tab(0).Control(18)=   "cmb_TieAzo"
            Tab(0).ControlCount=   19
            TabCaption(1)   =   "Estacionamiento 1"
            TabPicture(1)   =   "OpeTra_frm_155.frx":0EC0
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Label9"
            Tab(1).Control(1)=   "Label10"
            Tab(1).Control(2)=   "Label24"
            Tab(1).Control(3)=   "Label25"
            Tab(1).Control(4)=   "Label26"
            Tab(1).Control(5)=   "Label27"
            Tab(1).Control(6)=   "Label28"
            Tab(1).Control(7)=   "Label29"
            Tab(1).Control(8)=   "Label30"
            Tab(1).Control(9)=   "SSPanel10"
            Tab(1).Control(10)=   "ipp_ValRea_Es1"
            Tab(1).Control(11)=   "ipp_ValCom_Es1"
            Tab(1).Control(12)=   "ipp_ValACo_Es1"
            Tab(1).Control(13)=   "ipp_ValEdi_Es1"
            Tab(1).Control(14)=   "ipp_ValTer_Es1"
            Tab(1).Control(15)=   "ipp_SumAse_Es1"
            Tab(1).Control(16)=   "ipp_AreCon_Es1"
            Tab(1).Control(17)=   "ipp_AreTer_Es1"
            Tab(1).Control(18)=   "SSPanel9"
            Tab(1).Control(19)=   "cmb_FlgEst_Es1"
            Tab(1).ControlCount=   20
            TabCaption(2)   =   "Estacionamiento 2"
            TabPicture(2)   =   "OpeTra_frm_155.frx":0EDC
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "cmb_FlgEst_Es2"
            Tab(2).Control(1)=   "SSPanel11"
            Tab(2).Control(2)=   "ipp_AreTer_Es2"
            Tab(2).Control(3)=   "ipp_AreCon_Es2"
            Tab(2).Control(4)=   "ipp_SumAse_Es2"
            Tab(2).Control(5)=   "ipp_ValTer_Es2"
            Tab(2).Control(6)=   "ipp_ValEdi_Es2"
            Tab(2).Control(7)=   "ipp_ValACo_Es2"
            Tab(2).Control(8)=   "ipp_ValCom_Es2"
            Tab(2).Control(9)=   "ipp_ValRea_Es2"
            Tab(2).Control(10)=   "SSPanel12"
            Tab(2).Control(11)=   "Label39"
            Tab(2).Control(12)=   "Label38"
            Tab(2).Control(13)=   "Label37"
            Tab(2).Control(14)=   "Label36"
            Tab(2).Control(15)=   "Label35"
            Tab(2).Control(16)=   "Label34"
            Tab(2).Control(17)=   "Label33"
            Tab(2).Control(18)=   "Label32"
            Tab(2).Control(19)=   "Label31"
            Tab(2).ControlCount=   20
            TabCaption(3)   =   "Depósito"
            TabPicture(3)   =   "OpeTra_frm_155.frx":0EF8
            Tab(3).ControlEnabled=   -1  'True
            Tab(3).Control(0)=   "Label40"
            Tab(3).Control(0).Enabled=   0   'False
            Tab(3).Control(1)=   "Label41"
            Tab(3).Control(1).Enabled=   0   'False
            Tab(3).Control(2)=   "Label42"
            Tab(3).Control(2).Enabled=   0   'False
            Tab(3).Control(3)=   "Label43"
            Tab(3).Control(3).Enabled=   0   'False
            Tab(3).Control(4)=   "Label44"
            Tab(3).Control(4).Enabled=   0   'False
            Tab(3).Control(5)=   "Label45"
            Tab(3).Control(5).Enabled=   0   'False
            Tab(3).Control(6)=   "Label46"
            Tab(3).Control(6).Enabled=   0   'False
            Tab(3).Control(7)=   "Label47"
            Tab(3).Control(7).Enabled=   0   'False
            Tab(3).Control(8)=   "Label48"
            Tab(3).Control(8).Enabled=   0   'False
            Tab(3).Control(9)=   "SSPanel14"
            Tab(3).Control(9).Enabled=   0   'False
            Tab(3).Control(10)=   "ipp_ValRea_Dep"
            Tab(3).Control(10).Enabled=   0   'False
            Tab(3).Control(11)=   "ipp_ValCom_Dep"
            Tab(3).Control(11).Enabled=   0   'False
            Tab(3).Control(12)=   "ipp_ValACo_Dep"
            Tab(3).Control(12).Enabled=   0   'False
            Tab(3).Control(13)=   "ipp_ValEdi_Dep"
            Tab(3).Control(13).Enabled=   0   'False
            Tab(3).Control(14)=   "ipp_ValTer_Dep"
            Tab(3).Control(14).Enabled=   0   'False
            Tab(3).Control(15)=   "ipp_SumAse_Dep"
            Tab(3).Control(15).Enabled=   0   'False
            Tab(3).Control(16)=   "ipp_AreCon_Dep"
            Tab(3).Control(16).Enabled=   0   'False
            Tab(3).Control(17)=   "ipp_AreTer_Dep"
            Tab(3).Control(17).Enabled=   0   'False
            Tab(3).Control(18)=   "SSPanel13"
            Tab(3).Control(18).Enabled=   0   'False
            Tab(3).Control(19)=   "cmb_FlgEst_Dep"
            Tab(3).Control(19).Enabled=   0   'False
            Tab(3).ControlCount=   20
            Begin VB.ComboBox cmb_TieAzo 
               Height          =   315
               Left            =   -69480
               Style           =   2  'Dropdown List
               TabIndex        =   138
               Top             =   390
               Width           =   1335
            End
            Begin VB.ComboBox cmb_FlgEst_Dep 
               Height          =   315
               Left            =   1860
               Style           =   2  'Dropdown List
               TabIndex        =   38
               Top             =   390
               Width           =   975
            End
            Begin VB.ComboBox cmb_FlgEst_Es2 
               Height          =   315
               Left            =   -73140
               Style           =   2  'Dropdown List
               TabIndex        =   29
               Top             =   390
               Width           =   975
            End
            Begin VB.ComboBox cmb_FlgEst_Es1 
               Height          =   315
               Left            =   -73140
               Style           =   2  'Dropdown List
               TabIndex        =   20
               Top             =   390
               Width           =   975
            End
            Begin Threed.SSPanel SSPanel5 
               Height          =   60
               Left            =   -74970
               TabIndex        =   94
               Top             =   1080
               Width           =   10965
               _Version        =   65536
               _ExtentX        =   19341
               _ExtentY        =   106
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
            End
            Begin EditLib.fpDoubleSingle ipp_AreTer_Inm 
               Height          =   315
               Left            =   -73140
               TabIndex        =   12
               Top             =   390
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
            Begin EditLib.fpDoubleSingle ipp_AreCon_Inm 
               Height          =   315
               Left            =   -73140
               TabIndex        =   13
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
            Begin EditLib.fpDoubleSingle ipp_SumAse_Inm 
               Height          =   315
               Left            =   -73140
               TabIndex        =   14
               Top             =   1200
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
            Begin EditLib.fpDoubleSingle ipp_ValTer_Inm 
               Height          =   315
               Left            =   -73140
               TabIndex        =   17
               Top             =   1530
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
            Begin EditLib.fpDoubleSingle ipp_ValEdi_Inm 
               Height          =   315
               Left            =   -69480
               TabIndex        =   18
               Top             =   1530
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
            Begin EditLib.fpDoubleSingle ipp_ValACo_Inm 
               Height          =   315
               Left            =   -65550
               TabIndex        =   19
               Top             =   1530
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
            Begin EditLib.fpDoubleSingle ipp_ValCom_Inm 
               Height          =   315
               Left            =   -69480
               TabIndex        =   15
               Top             =   1200
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
            Begin EditLib.fpDoubleSingle ipp_ValRea_Inm 
               Height          =   315
               Left            =   -65550
               TabIndex        =   16
               Top             =   1200
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
            Begin Threed.SSPanel SSPanel9 
               Height          =   60
               Left            =   -74970
               TabIndex        =   95
               Top             =   750
               Width           =   10965
               _Version        =   65536
               _ExtentX        =   19341
               _ExtentY        =   106
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
            End
            Begin EditLib.fpDoubleSingle ipp_AreTer_Es1 
               Height          =   315
               Left            =   -73140
               TabIndex        =   21
               Top             =   870
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
            Begin EditLib.fpDoubleSingle ipp_AreCon_Es1 
               Height          =   315
               Left            =   -73140
               TabIndex        =   22
               Top             =   1200
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
            Begin EditLib.fpDoubleSingle ipp_SumAse_Es1 
               Height          =   315
               Left            =   -73140
               TabIndex        =   23
               Top             =   1680
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
            Begin EditLib.fpDoubleSingle ipp_ValTer_Es1 
               Height          =   315
               Left            =   -73140
               TabIndex        =   26
               Top             =   2010
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
            Begin EditLib.fpDoubleSingle ipp_ValEdi_Es1 
               Height          =   315
               Left            =   -69480
               TabIndex        =   27
               Top             =   2010
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
            Begin EditLib.fpDoubleSingle ipp_ValACo_Es1 
               Height          =   315
               Left            =   -65550
               TabIndex        =   28
               Top             =   2010
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
            Begin EditLib.fpDoubleSingle ipp_ValCom_Es1 
               Height          =   315
               Left            =   -69480
               TabIndex        =   24
               Top             =   1680
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
            Begin EditLib.fpDoubleSingle ipp_ValRea_Es1 
               Height          =   315
               Left            =   -65550
               TabIndex        =   25
               Top             =   1680
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
            Begin Threed.SSPanel SSPanel10 
               Height          =   60
               Left            =   -74970
               TabIndex        =   96
               Top             =   1560
               Width           =   10965
               _Version        =   65536
               _ExtentX        =   19341
               _ExtentY        =   106
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
            End
            Begin Threed.SSPanel SSPanel11 
               Height          =   60
               Left            =   -74970
               TabIndex        =   97
               Top             =   750
               Width           =   10965
               _Version        =   65536
               _ExtentX        =   19341
               _ExtentY        =   106
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
            End
            Begin EditLib.fpDoubleSingle ipp_AreTer_Es2 
               Height          =   315
               Left            =   -73140
               TabIndex        =   30
               Top             =   870
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
            Begin EditLib.fpDoubleSingle ipp_AreCon_Es2 
               Height          =   315
               Left            =   -73140
               TabIndex        =   31
               Top             =   1200
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
            Begin EditLib.fpDoubleSingle ipp_SumAse_Es2 
               Height          =   315
               Left            =   -73140
               TabIndex        =   32
               Top             =   1680
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
            Begin EditLib.fpDoubleSingle ipp_ValTer_Es2 
               Height          =   315
               Left            =   -73140
               TabIndex        =   35
               Top             =   2010
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
            Begin EditLib.fpDoubleSingle ipp_ValEdi_Es2 
               Height          =   315
               Left            =   -69480
               TabIndex        =   36
               Top             =   2010
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
            Begin EditLib.fpDoubleSingle ipp_ValACo_Es2 
               Height          =   315
               Left            =   -65550
               TabIndex        =   37
               Top             =   2010
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
            Begin EditLib.fpDoubleSingle ipp_ValCom_Es2 
               Height          =   315
               Left            =   -69480
               TabIndex        =   33
               Top             =   1680
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
            Begin EditLib.fpDoubleSingle ipp_ValRea_Es2 
               Height          =   315
               Left            =   -65550
               TabIndex        =   34
               Top             =   1680
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
            Begin Threed.SSPanel SSPanel12 
               Height          =   60
               Left            =   -74970
               TabIndex        =   98
               Top             =   1560
               Width           =   10965
               _Version        =   65536
               _ExtentX        =   19341
               _ExtentY        =   106
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
            End
            Begin Threed.SSPanel SSPanel13 
               Height          =   60
               Left            =   30
               TabIndex        =   99
               Top             =   750
               Width           =   10965
               _Version        =   65536
               _ExtentX        =   19341
               _ExtentY        =   106
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
            End
            Begin EditLib.fpDoubleSingle ipp_AreTer_Dep 
               Height          =   315
               Left            =   1860
               TabIndex        =   39
               Top             =   870
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
            Begin EditLib.fpDoubleSingle ipp_AreCon_Dep 
               Height          =   315
               Left            =   1860
               TabIndex        =   40
               Top             =   1200
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
            Begin EditLib.fpDoubleSingle ipp_SumAse_Dep 
               Height          =   315
               Left            =   1860
               TabIndex        =   41
               Top             =   1680
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
            Begin EditLib.fpDoubleSingle ipp_ValTer_Dep 
               Height          =   315
               Left            =   1860
               TabIndex        =   44
               Top             =   2010
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
            Begin EditLib.fpDoubleSingle ipp_ValEdi_Dep 
               Height          =   315
               Left            =   5520
               TabIndex        =   45
               Top             =   2010
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
            Begin EditLib.fpDoubleSingle ipp_ValACo_Dep 
               Height          =   315
               Left            =   9450
               TabIndex        =   46
               Top             =   2010
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
            Begin EditLib.fpDoubleSingle ipp_ValCom_Dep 
               Height          =   315
               Left            =   5520
               TabIndex        =   42
               Top             =   1680
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
            Begin EditLib.fpDoubleSingle ipp_ValRea_Dep 
               Height          =   315
               Left            =   9450
               TabIndex        =   43
               Top             =   1680
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
            Begin Threed.SSPanel SSPanel14 
               Height          =   60
               Left            =   30
               TabIndex        =   100
               Top             =   1560
               Width           =   10965
               _Version        =   65536
               _ExtentX        =   19341
               _ExtentY        =   106
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
            End
            Begin VB.Label Label6 
               Caption         =   "Tiene Azotea"
               Height          =   315
               Left            =   -70980
               TabIndex        =   139
               Top             =   390
               Width           =   1485
            End
            Begin VB.Label Label48 
               Caption         =   "Valor Realización:"
               Height          =   285
               Left            =   7680
               TabIndex        =   135
               Top             =   1680
               Width           =   1485
            End
            Begin VB.Label Label47 
               Caption         =   "Valor Comercial:"
               Height          =   285
               Left            =   4020
               TabIndex        =   134
               Top             =   1680
               Width           =   1485
            End
            Begin VB.Label Label46 
               Caption         =   "Valor Areas Comunes:"
               Height          =   285
               Left            =   7680
               TabIndex        =   133
               Top             =   2010
               Width           =   1725
            End
            Begin VB.Label Label45 
               Caption         =   "Valor Edificación:"
               Height          =   285
               Left            =   4020
               TabIndex        =   132
               Top             =   2010
               Width           =   1485
            End
            Begin VB.Label Label44 
               Caption         =   "Valor Terreno:"
               Height          =   285
               Left            =   90
               TabIndex        =   131
               Top             =   2010
               Width           =   1485
            End
            Begin VB.Label Label43 
               Caption         =   "Suma Asegurada:"
               Height          =   285
               Left            =   90
               TabIndex        =   130
               Top             =   1680
               Width           =   1485
            End
            Begin VB.Label Label42 
               Caption         =   "Area Construcción:"
               Height          =   285
               Left            =   90
               TabIndex        =   129
               Top             =   1200
               Width           =   1485
            End
            Begin VB.Label Label41 
               Caption         =   "Area Terreno:"
               Height          =   285
               Left            =   90
               TabIndex        =   128
               Top             =   870
               Width           =   1485
            End
            Begin VB.Label Label40 
               Caption         =   "Depósito:"
               Height          =   315
               Left            =   90
               TabIndex        =   127
               Top             =   390
               Width           =   1365
            End
            Begin VB.Label Label39 
               Caption         =   "Valor Realización:"
               Height          =   285
               Left            =   -67320
               TabIndex        =   126
               Top             =   1680
               Width           =   1485
            End
            Begin VB.Label Label38 
               Caption         =   "Valor Comercial:"
               Height          =   285
               Left            =   -70980
               TabIndex        =   125
               Top             =   1680
               Width           =   1485
            End
            Begin VB.Label Label37 
               Caption         =   "Valor Areas Comunes:"
               Height          =   285
               Left            =   -67320
               TabIndex        =   124
               Top             =   2010
               Width           =   1725
            End
            Begin VB.Label Label36 
               Caption         =   "Valor Edificación:"
               Height          =   285
               Left            =   -70980
               TabIndex        =   123
               Top             =   2010
               Width           =   1485
            End
            Begin VB.Label Label35 
               Caption         =   "Valor Terreno:"
               Height          =   285
               Left            =   -74910
               TabIndex        =   122
               Top             =   2010
               Width           =   1485
            End
            Begin VB.Label Label34 
               Caption         =   "Suma Asegurada:"
               Height          =   285
               Left            =   -74910
               TabIndex        =   121
               Top             =   1680
               Width           =   1485
            End
            Begin VB.Label Label33 
               Caption         =   "Area Construcción:"
               Height          =   285
               Left            =   -74910
               TabIndex        =   120
               Top             =   1200
               Width           =   1485
            End
            Begin VB.Label Label32 
               Caption         =   "Area Terreno:"
               Height          =   285
               Left            =   -74910
               TabIndex        =   119
               Top             =   870
               Width           =   1485
            End
            Begin VB.Label Label31 
               Caption         =   "Estacionamiento 2:"
               Height          =   315
               Left            =   -74910
               TabIndex        =   118
               Top             =   390
               Width           =   1365
            End
            Begin VB.Label Label30 
               Caption         =   "Valor Realización:"
               Height          =   285
               Left            =   -67320
               TabIndex        =   117
               Top             =   1680
               Width           =   1485
            End
            Begin VB.Label Label29 
               Caption         =   "Valor Comercial:"
               Height          =   285
               Left            =   -70980
               TabIndex        =   116
               Top             =   1680
               Width           =   1485
            End
            Begin VB.Label Label28 
               Caption         =   "Valor Areas Comunes:"
               Height          =   285
               Left            =   -67320
               TabIndex        =   115
               Top             =   2010
               Width           =   1725
            End
            Begin VB.Label Label27 
               Caption         =   "Valor Edificación:"
               Height          =   285
               Left            =   -70980
               TabIndex        =   114
               Top             =   2010
               Width           =   1485
            End
            Begin VB.Label Label26 
               Caption         =   "Valor Terreno:"
               Height          =   285
               Left            =   -74910
               TabIndex        =   113
               Top             =   2010
               Width           =   1485
            End
            Begin VB.Label Label25 
               Caption         =   "Suma Asegurada:"
               Height          =   285
               Left            =   -74910
               TabIndex        =   112
               Top             =   1680
               Width           =   1485
            End
            Begin VB.Label Label24 
               Caption         =   "Area Construcción:"
               Height          =   285
               Left            =   -74910
               TabIndex        =   111
               Top             =   1200
               Width           =   1485
            End
            Begin VB.Label Label10 
               Caption         =   "Area Terreno:"
               Height          =   285
               Left            =   -74910
               TabIndex        =   110
               Top             =   870
               Width           =   1485
            End
            Begin VB.Label Label9 
               Caption         =   "Estacionamiento 1:"
               Height          =   315
               Left            =   -74910
               TabIndex        =   109
               Top             =   390
               Width           =   1365
            End
            Begin VB.Label Label20 
               Caption         =   "Valor Realización:"
               Height          =   285
               Left            =   -67320
               TabIndex        =   108
               Top             =   1200
               Width           =   1485
            End
            Begin VB.Label Label19 
               Caption         =   "Valor Comercial:"
               Height          =   285
               Left            =   -70980
               TabIndex        =   107
               Top             =   1200
               Width           =   1485
            End
            Begin VB.Label Label18 
               Caption         =   "Valor Areas Comunes:"
               Height          =   285
               Left            =   -67320
               TabIndex        =   106
               Top             =   1530
               Width           =   1725
            End
            Begin VB.Label Label17 
               Caption         =   "Valor Edificación:"
               Height          =   285
               Left            =   -70980
               TabIndex        =   105
               Top             =   1530
               Width           =   1485
            End
            Begin VB.Label Label16 
               Caption         =   "Valor Terreno:"
               Height          =   285
               Left            =   -74910
               TabIndex        =   104
               Top             =   1530
               Width           =   1485
            End
            Begin VB.Label Label15 
               Caption         =   "Suma Asegurada:"
               Height          =   285
               Left            =   -74910
               TabIndex        =   103
               Top             =   1200
               Width           =   1485
            End
            Begin VB.Label Label23 
               Caption         =   "Area Construcción:"
               Height          =   285
               Left            =   -74910
               TabIndex        =   102
               Top             =   720
               Width           =   1485
            End
            Begin VB.Label Label22 
               Caption         =   "Area Terreno:"
               Height          =   285
               Left            =   -74910
               TabIndex        =   101
               Top             =   390
               Width           =   1485
            End
         End
      End
   End
End
Attribute VB_Name = "frm_ModSol_06"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_EmpPer()      As moddat_tpo_Genera
Dim l_arr_PerTas()      As moddat_tpo_Genera
Dim l_arr_ParPrd()      As moddat_tpo_Genera
Dim l_arr_CliNCo()      As modcal_g_est_CuoCli

Dim l_Arr_TNC_Cli()     As String
Dim l_Arr_TC_Cli()      As String
Dim l_Arr_TNC_Cof()     As String
Dim l_Arr_TC_Cof()      As String

Dim l_str_EmpSeg        As String
Dim l_int_TipSeg        As Integer
Dim l_dbl_TasInt        As Double
Dim l_dbl_IntGra        As Double
Dim l_int_PlaAno        As Integer
Dim l_int_PerGra        As Integer
Dim l_int_DiaPag        As Integer
Dim l_dbl_MtoPre        As Double
Dim l_dbl_CuoApr        As Double
Dim l_dbl_CuoAce        As Double
Dim l_str_CodCiu        As String
Dim l_int_TasEsp        As Integer

'variables nueva para la generacion del cronograma
Dim obj_Cronog             As Object
Dim int_Produc             As Integer
Dim int_CuoDbl             As Integer
Dim dbl_ValInm             As Double
Dim dbl_CuoIni             As Double
Dim dbl_MtoCon             As Double
Dim dbl_MtoTas             As Double
Dim int_PlaPre             As Integer
Dim dbl_TasInt             As Double
Dim dbl_TasCof             As Double
Dim dbl_ComCof             As Double
Dim dat_FecDes             As Date
Dim int_DiaVct             As Integer
Dim int_PerGra             As Integer
Dim str_PriVct             As String
Dim dbl_Portes             As Double
Dim dbl_SegViv             As Double
Dim int_TipSDe             As Integer
Dim dbl_SegDes             As Double
Dim dbl_CuoMen             As Double
Dim dbl_CuoPbp             As Double
Dim dbl_IngReq             As Double

Private Sub fs_GenExc()
Dim r_obj_Excel      As Excel.Application
Dim r_int_NroFil     As Integer
Dim r_int_Conta      As Integer

   'Generando Excel
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.DisplayAlerts = False
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   r_int_NroFil = 1
   
   With r_obj_Excel.ActiveSheet
      r_int_NroFil = 2
      .Cells(1, 4) = "FECHA IMPRESION: " & Format(date, "dd/mm/yyyy")
      .Range(.Cells(1, 4), .Cells(1, 5)).Merge
      .Cells(1, 4).HorizontalAlignment = xlHAlignCenter
      
      .Cells(r_int_NroFil, 1) = "DATOS DE TASACION"
      .Cells(r_int_NroFil, 1).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 1).Font.Bold = True
      .Range("A" & r_int_NroFil & ":E" & r_int_NroFil).Merge
      
      'muestra de titulos
      r_int_NroFil = 4
      .Cells(r_int_NroFil, 1) = "DATOS PRINCIPALES"
      .Cells(r_int_NroFil + 1, 1) = "Nro. Solicitud"
      .Cells(r_int_NroFil + 1, 4) = "Fec. Ing. Solicitud"
      .Cells(r_int_NroFil + 2, 1) = "Cliente"
      .Cells(r_int_NroFil + 3, 1) = "Empresa Peritaje"
      .Cells(r_int_NroFil + 4, 1) = "Perito Tasador"
      .Cells(r_int_NroFil + 5, 1) = "Numero Informe"
      .Cells(r_int_NroFil + 5, 4) = "Fecha Evaluacion"
      .Cells(r_int_NroFil + 6, 1) = "Año Construccion"
      .Cells(r_int_NroFil + 7, 1) = "Numero Pisos"
      .Cells(r_int_NroFil + 7, 4) = "Numero Sotanos"
      .Cells(r_int_NroFil + 8, 1) = "Tipo Inmueble"
      .Cells(r_int_NroFil + 8, 4) = "Uso Inmueble"
      .Cells(r_int_NroFil + 9, 1) = "Material Construccion"
      .Cells(r_int_NroFil + 10, 1) = "Moneda"
      .Cells(r_int_NroFil + 10, 4) = "Tipo Cambio"
            
      .Cells(r_int_NroFil + 12, 1) = "TOTALES"
      .Cells(r_int_NroFil + 13, 1) = "Area Terreno"
      .Cells(r_int_NroFil + 13, 3) = "Area Construccion"
      .Cells(r_int_NroFil + 14, 1) = "Suma Asegurada"
      .Cells(r_int_NroFil + 14, 3) = "Valor Comercial"
      .Cells(r_int_NroFil + 15, 1) = "Valor Terreno"
      .Cells(r_int_NroFil + 15, 3) = "Valor Edificacion"
      .Cells(r_int_NroFil + 16, 1) = "Valor Realizacion"
      .Cells(r_int_NroFil + 16, 3) = "Valor Areas Comunes"
      
      .Cells(r_int_NroFil + 18, 1) = "INMUEBLE"
      .Cells(r_int_NroFil + 19, 1) = "Area Terreno"
      .Cells(r_int_NroFil + 19, 3) = "Tiene Azotea"
      .Cells(r_int_NroFil + 20, 1) = "Area Construccion"
      .Cells(r_int_NroFil + 20, 3) = "Valor Comercial"
      .Cells(r_int_NroFil + 21, 1) = "Suma Asegurada"
      .Cells(r_int_NroFil + 21, 3) = "Valor Edificacion"
      .Cells(r_int_NroFil + 22, 1) = "Valor Terreno"
      .Cells(r_int_NroFil + 22, 3) = "Valor Realizacion"
      .Cells(r_int_NroFil + 23, 1) = "Valor Areas Comunes"
      
      .Cells(r_int_NroFil + 25, 1) = "ESTACIONAMIENTO 1"
      .Cells(r_int_NroFil + 26, 1) = "Estacionamiento 1"
      .Cells(r_int_NroFil + 26, 3) = "Valor Comercial"
      .Cells(r_int_NroFil + 27, 1) = "Area Terreno"
      .Cells(r_int_NroFil + 27, 3) = "Valor Edificacion"
      .Cells(r_int_NroFil + 28, 1) = "Area Construccion"
      .Cells(r_int_NroFil + 28, 3) = "Valor Realizacion"
      .Cells(r_int_NroFil + 29, 1) = "Suma Asegurada"
      .Cells(r_int_NroFil + 29, 3) = "Valor Areas Comunes"
      .Cells(r_int_NroFil + 30, 1) = "Valor Terreno"
            
      .Cells(r_int_NroFil + 32, 1) = "ESTACIONAMIENTO 2"
      .Cells(r_int_NroFil + 33, 1) = "Estacionamiento 2"
      .Cells(r_int_NroFil + 33, 3) = "Valor Comercial"
      .Cells(r_int_NroFil + 34, 1) = "Area Terreno"
      .Cells(r_int_NroFil + 34, 3) = "Valor Edificacion"
      .Cells(r_int_NroFil + 35, 1) = "Area Construccion"
      .Cells(r_int_NroFil + 35, 3) = "Valor Realizacion"
      .Cells(r_int_NroFil + 36, 1) = "Suma Asegurada"
      .Cells(r_int_NroFil + 36, 3) = "Valor Areas Comunes"
      .Cells(r_int_NroFil + 37, 1) = "Valor Terreno"
      
      .Cells(r_int_NroFil + 39, 1) = "DEPOSITO"
      .Cells(r_int_NroFil + 40, 1) = "Deposito"
      .Cells(r_int_NroFil + 40, 3) = "Valor Comercial"
      .Cells(r_int_NroFil + 41, 1) = "Area Terreno"
      .Cells(r_int_NroFil + 41, 3) = "Valor Edificacion"
      .Cells(r_int_NroFil + 42, 1) = "Area Construccion"
      .Cells(r_int_NroFil + 42, 3) = "Valor Realizacion"
      .Cells(r_int_NroFil + 43, 1) = "Suma Asegurada"
      .Cells(r_int_NroFil + 43, 3) = "Valor Areas Comunes"
      .Cells(r_int_NroFil + 44, 1) = "Valor Terreno"
            
      'muestra de informacion de datos
      .Cells(r_int_NroFil + 1, 2) = pnl_NumSol.Caption
      .Cells(r_int_NroFil + 1, 5) = "'" & pnl_FecIng.Caption
      .Cells(r_int_NroFil + 2, 2) = pnl_Client.Caption
      .Cells(r_int_NroFil + 3, 2) = cmb_EmpPer.Text
      .Cells(r_int_NroFil + 4, 2) = cmb_PerTas.Text
      .Cells(r_int_NroFil + 5, 2) = txt_NumInf.Text
      .Cells(r_int_NroFil + 5, 5) = "'" & ipp_FecEva.Text
      .Cells(r_int_NroFil + 6, 2) = ipp_AnoCon.Text
      .Cells(r_int_NroFil + 7, 2) = ipp_NumPis.Text
      .Cells(r_int_NroFil + 7, 5) = ipp_NumSot.Text
      .Cells(r_int_NroFil + 8, 2) = cmb_TipInm.Text
      .Cells(r_int_NroFil + 8, 5) = cmb_UsoInm.Text
      .Cells(r_int_NroFil + 9, 2) = cmb_MatCon.Text
      .Cells(r_int_NroFil + 10, 2) = cmb_TipMon.Text
      .Cells(r_int_NroFil + 10, 5) = ipp_TipCam.Text
      
      .Cells(r_int_NroFil + 13, 2) = pnl_AreTer.Caption
      .Cells(r_int_NroFil + 13, 4) = pnl_AreCon.Caption
      .Cells(r_int_NroFil + 14, 2) = pnl_SumAse.Caption
      .Cells(r_int_NroFil + 14, 4) = pnl_ValCom.Caption
      .Cells(r_int_NroFil + 15, 2) = pnl_ValTer.Caption
      .Cells(r_int_NroFil + 15, 4) = pnl_ValEdi.Caption
      .Cells(r_int_NroFil + 16, 2) = pnl_ValRea.Caption
      .Cells(r_int_NroFil + 16, 4) = pnl_ValACo.Caption
            
      .Cells(r_int_NroFil + 19, 2) = ipp_AreTer_Inm.Text
      .Cells(r_int_NroFil + 19, 4) = cmb_TieAzo.Text
      .Cells(r_int_NroFil + 20, 2) = ipp_AreCon_Inm.Text
      .Cells(r_int_NroFil + 20, 4) = ipp_ValCom_Inm.Text
      .Cells(r_int_NroFil + 21, 2) = ipp_SumAse_Inm.Text
      .Cells(r_int_NroFil + 21, 4) = ipp_ValEdi_Inm.Text
      .Cells(r_int_NroFil + 22, 2) = ipp_ValTer_Inm.Text
      .Cells(r_int_NroFil + 22, 4) = ipp_ValRea_Inm.Text
      .Cells(r_int_NroFil + 23, 2) = ipp_ValACo_Inm.Text
      
      .Cells(r_int_NroFil + 26, 2) = cmb_FlgEst_Es1.Text
      .Cells(r_int_NroFil + 26, 4) = ipp_ValCom_Es1.Text
      .Cells(r_int_NroFil + 27, 2) = ipp_AreTer_Es1.Text
      .Cells(r_int_NroFil + 27, 4) = ipp_ValEdi_Es1.Text
      .Cells(r_int_NroFil + 28, 2) = ipp_AreCon_Es1.Text
      .Cells(r_int_NroFil + 28, 4) = ipp_ValRea_Es1.Text
      .Cells(r_int_NroFil + 29, 2) = ipp_SumAse_Es1.Text
      .Cells(r_int_NroFil + 29, 4) = ipp_ValACo_Es1.Text
      .Cells(r_int_NroFil + 30, 2) = ipp_ValTer_Es1.Text
      
      .Cells(r_int_NroFil + 33, 2) = cmb_FlgEst_Es2.Text
      .Cells(r_int_NroFil + 33, 4) = ipp_ValCom_Es2.Text
      .Cells(r_int_NroFil + 34, 2) = ipp_AreTer_Es2.Text
      .Cells(r_int_NroFil + 34, 4) = ipp_ValEdi_Es2.Text
      .Cells(r_int_NroFil + 35, 2) = ipp_AreCon_Es2.Text
      .Cells(r_int_NroFil + 35, 4) = ipp_ValRea_Es2.Text
      .Cells(r_int_NroFil + 36, 2) = ipp_SumAse_Es2.Text
      .Cells(r_int_NroFil + 36, 4) = ipp_ValACo_Es2.Text
      .Cells(r_int_NroFil + 37, 2) = ipp_ValTer_Es2.Text
      
      .Cells(r_int_NroFil + 40, 2) = cmb_FlgEst_Dep.Text
      .Cells(r_int_NroFil + 40, 4) = ipp_AreTer_Dep.Text
      .Cells(r_int_NroFil + 41, 2) = ipp_AreCon_Dep.Text
      .Cells(r_int_NroFil + 41, 4) = ipp_SumAse_Dep.Text
      .Cells(r_int_NroFil + 42, 2) = ipp_ValCom_Dep.Text
      .Cells(r_int_NroFil + 42, 4) = ipp_ValRea_Dep.Text
      .Cells(r_int_NroFil + 43, 2) = ipp_ValTer_Dep.Text
      .Cells(r_int_NroFil + 43, 4) = ipp_ValEdi_Dep.Text
      .Cells(r_int_NroFil + 44, 2) = ipp_ValACo_Dep.Text
      
      .Columns("A").ColumnWidth = 17
      .Columns("B").ColumnWidth = 17
      .Columns("C").ColumnWidth = 21
      .Columns("D").ColumnWidth = 14
      .Columns("E").ColumnWidth = 14
      
      .Range(.Cells(4, 1), .Cells(50, 1)).Font.Bold = True
      .Range(.Cells(4, 3), .Cells(50, 3)).Font.Bold = True
      .Range(.Cells(4, 4), .Cells(14, 4)).Font.Bold = True
      
      .Columns("B").HorizontalAlignment = xlHAlignLeft
      .Columns("D").HorizontalAlignment = xlHAlignLeft
      .Columns("E").HorizontalAlignment = xlHAlignLeft
            
      For r_int_Conta = 5 To 14
         .Range(.Cells(r_int_Conta, 2), .Cells(r_int_Conta, 3)).Merge
      Next
            
      .Range(.Cells(13, 3), .Cells(13, 3)).WrapText = True
      .Range(.Cells(13, 1), .Cells(13, 2)).VerticalAlignment = xlVAlignCenter
      .Range(.Cells(13, 2), .Cells(13, 2)).RowHeight = 30
      
      .Range(.Cells(2, 1), .Cells(2, 5)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(2, 1), .Cells(2, 5)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(2, 1), .Cells(2, 5)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(2, 1), .Cells(2, 5)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(2, 1), .Cells(2, 5)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(2, 1), .Cells(2, 5)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(2, 1), .Cells(2, 5)).Borders(xlInsideVertical).LineStyle = xlContinuous

      .Range(.Cells(4, 1), .Cells(4, 1)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(4, 1), .Cells(4, 1)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(4, 1), .Cells(4, 1)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 1), .Cells(4, 1)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(4, 1), .Cells(4, 1)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(4, 1), .Cells(4, 1)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(4, 1), .Cells(4, 1)).Borders(xlInsideVertical).LineStyle = xlContinuous
      .Range(.Cells(5, 1), .Cells(14, 5)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(5, 1), .Cells(14, 5)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(5, 1), .Cells(14, 5)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(5, 1), .Cells(14, 5)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(5, 1), .Cells(14, 5)).Borders(xlInsideVertical).LineStyle = xlContinuous
      
      .Range(.Cells(16, 1), .Cells(16, 1)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(16, 1), .Cells(16, 1)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(16, 1), .Cells(16, 1)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(16, 1), .Cells(16, 1)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(16, 1), .Cells(16, 1)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(16, 1), .Cells(16, 1)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(16, 1), .Cells(16, 1)).Borders(xlInsideVertical).LineStyle = xlContinuous
      .Range(.Cells(17, 1), .Cells(20, 4)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(17, 1), .Cells(20, 4)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(17, 1), .Cells(20, 4)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(17, 1), .Cells(20, 4)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(17, 1), .Cells(20, 4)).Borders(xlInsideVertical).LineStyle = xlContinuous
      
      .Range(.Cells(22, 1), .Cells(22, 1)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(22, 1), .Cells(22, 1)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(22, 1), .Cells(22, 1)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(22, 1), .Cells(22, 1)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(22, 1), .Cells(22, 1)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(22, 1), .Cells(22, 1)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(22, 1), .Cells(22, 1)).Borders(xlInsideVertical).LineStyle = xlContinuous
      .Range(.Cells(23, 1), .Cells(27, 4)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(23, 1), .Cells(27, 4)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(23, 1), .Cells(27, 4)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(23, 1), .Cells(27, 4)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(23, 1), .Cells(27, 4)).Borders(xlInsideVertical).LineStyle = xlContinuous
      
      .Range(.Cells(29, 1), .Cells(29, 1)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(29, 1), .Cells(29, 1)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(29, 1), .Cells(29, 1)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(29, 1), .Cells(29, 1)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(29, 1), .Cells(29, 1)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(29, 1), .Cells(29, 1)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(29, 1), .Cells(29, 1)).Borders(xlInsideVertical).LineStyle = xlContinuous
      .Range(.Cells(30, 1), .Cells(34, 4)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(30, 1), .Cells(34, 4)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(30, 1), .Cells(34, 4)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(30, 1), .Cells(34, 4)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(30, 1), .Cells(34, 4)).Borders(xlInsideVertical).LineStyle = xlContinuous
      
      .Range(.Cells(36, 1), .Cells(36, 1)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(36, 1), .Cells(36, 1)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(36, 1), .Cells(36, 1)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(36, 1), .Cells(36, 1)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(36, 1), .Cells(36, 1)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(36, 1), .Cells(36, 1)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(36, 1), .Cells(36, 1)).Borders(xlInsideVertical).LineStyle = xlContinuous
      .Range(.Cells(37, 1), .Cells(41, 4)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(37, 1), .Cells(41, 4)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(37, 1), .Cells(41, 4)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(37, 1), .Cells(41, 4)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(37, 1), .Cells(41, 4)).Borders(xlInsideVertical).LineStyle = xlContinuous
      
      .Range(.Cells(43, 1), .Cells(43, 1)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(43, 1), .Cells(43, 1)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(43, 1), .Cells(43, 1)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(43, 1), .Cells(43, 1)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(43, 1), .Cells(43, 1)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(43, 1), .Cells(43, 1)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(43, 1), .Cells(43, 1)).Borders(xlInsideVertical).LineStyle = xlContinuous
      .Range(.Cells(44, 1), .Cells(48, 4)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(44, 1), .Cells(48, 4)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(44, 1), .Cells(48, 4)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(44, 1), .Cells(48, 4)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(44, 1), .Cells(48, 4)).Borders(xlInsideVertical).LineStyle = xlContinuous
      
      .Range(.Cells(1, 1), .Cells(50, 10)).Font.Name = "Arial"
      .Range(.Cells(1, 1), .Cells(50, 10)).Font.Size = 8
   End With
   r_obj_Excel.Sheets(1).Name = "Tasacion"
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub cmb_TieAzo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_SumAse_Inm)
    End If
End Sub

Private Sub cmd_Export_Click()
   'Confirmacion
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_GenExc
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Grabar_Click()
Dim r_dbl_Valmax_ValEdi    As Double
Dim r_dbl_MtoPre           As Double
Dim r_dbl_IntCap           As Double
Dim r_dbl_ComVta           As Double
Dim r_dbl_TipCam           As Double
Dim r_dbl_TCaMpr           As Double
Dim r_dbl_ValRea           As Double
Dim r_dbl_PorMax_ValGar    As Double
Dim r_dbl_PorMin_ValGrv    As Double
Dim r_dbl_ValGar           As Double
Dim r_int_FlgExc           As Integer
Dim r_int_FlgExc_2         As Integer
Dim r_dbl_ValViv           As Double
Dim r_dbl_Portes           As Double
Dim r_int_TipVal_Des       As Integer
Dim r_int_TipVal_Viv       As Integer
Dim r_dbl_Import_Des       As Double
Dim r_dbl_Import_Viv       As Double
Dim r_dbl_PorCon           As Double
Dim r_dbl_TopCon           As Double
Dim r_dbl_MtoNCo           As Double
Dim r_dbl_MtoCon           As Double
Dim r_dbl_CuoNue           As Double
Dim r_int_CuoExt           As Integer
Dim r_dbl_ApoPro           As Double
Dim r_dbl_MtoTas           As Double
Dim r_dbl_ComCof           As Double
Dim r_dbl_TasCof           As Double

   If cmb_EmpPer.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Empresa de Peritaje.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_EmpPer)
      Exit Sub
   End If
   If Len(Trim(cmb_PerTas.Text)) = 0 Then
      MsgBox "Debe seleccionar el Perito Tasador.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PerTas)
      Exit Sub
   End If
   If Len(Trim(txt_NumInf.Text)) = 0 Then
      MsgBox "Debe ingresar el Número de Informe del Perito.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NumInf)
      Exit Sub
   End If
   If CDate(ipp_FecEva.Text) > date Then
      MsgBox "La Fecha de Evaluación no puede ser mayor a la Fecha Actual.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecEva)
      Exit Sub
   End If
   If ipp_AnoCon.Value = 0 Then
      MsgBox "Debe ingresar el Año de Construcción.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_AnoCon)
      Exit Sub
   End If
   If cmb_TieAzo.ListIndex = -1 Then
      MsgBox "Debe ingresar Azotea.", vbExclamation, modgen_g_str_NomPlt
      tab_Genera.Tab = 0
      Call gs_SetFocus(cmb_TieAzo)
      Exit Sub
   End If
   If ipp_NumPis.Value = 0 Then
      MsgBox "Debe ingresar el Nro. de Pisos.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_NumPis)
      Exit Sub
   End If
   If cmb_TipInm.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Inmueble.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipInm)
      Exit Sub
   End If
   If cmb_UsoInm.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Uso del Inmueble.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_UsoInm)
      Exit Sub
   End If
   If cmb_MatCon.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Material de Construcción.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_MatCon)
      Exit Sub
   End If
   If cmb_TipMon.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Moneda de la Valuación.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipMon)
      Exit Sub
   End If
   If cmb_TipMon.ItemData(cmb_TipMon.ListIndex) <> moddat_g_int_TipMon Then
      MsgBox "La Moneda de la Tasación debe ser igual a la Moneda del Préstamo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipMon)
      Exit Sub
   End If
   If ipp_TipCam.Value = 0 Then
      MsgBox "Debe ingresar el Tipo de Cambio.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_TipCam)
      Exit Sub
   End If
   If ipp_AreTer_Inm.Value = 0 Then
      MsgBox "Debe ingresar el Area del Terreno.", vbExclamation, modgen_g_str_NomPlt
      tab_Genera.Tab = 0
      Call gs_SetFocus(ipp_AreTer_Inm)
      Exit Sub
   End If
   If ipp_AreCon_Inm.Value = 0 Then
      MsgBox "Debe ingresar el Area Construida.", vbExclamation, modgen_g_str_NomPlt
      tab_Genera.Tab = 0
      Call gs_SetFocus(ipp_AreCon_Inm)
      Exit Sub
   End If
   If ipp_SumAse_Inm.Value = 0 Then
      MsgBox "Debe ingresar la Suma Asegurada.", vbExclamation, modgen_g_str_NomPlt
      tab_Genera.Tab = 0
      Call gs_SetFocus(ipp_SumAse_Inm)
      Exit Sub
   End If
   If ipp_ValCom_Inm.Value = 0 Then
      MsgBox "Debe ingresar el Valor Comercial.", vbExclamation, modgen_g_str_NomPlt
      tab_Genera.Tab = 0
      Call gs_SetFocus(ipp_ValCom_Inm)
      Exit Sub
   End If
   If ipp_ValRea_Inm.Value = 0 Then
      MsgBox "Debe ingresar el Valor de Realización.", vbExclamation, modgen_g_str_NomPlt
      tab_Genera.Tab = 0
      Call gs_SetFocus(ipp_ValRea_Inm)
      Exit Sub
   End If
   If ipp_ValTer_Inm.Value = 0 Then
      MsgBox "Debe ingresar el Valor del Terreno.", vbExclamation, modgen_g_str_NomPlt
      tab_Genera.Tab = 0
      Call gs_SetFocus(ipp_ValTer_Inm)
      Exit Sub
   End If
   If ipp_ValEdi_Inm.Value = 0 Then
      MsgBox "Debe ingresar el Valor de Edificación.", vbExclamation, modgen_g_str_NomPlt
      tab_Genera.Tab = 0
      Call gs_SetFocus(ipp_ValEdi_Inm)
      Exit Sub
   End If
   
   'Estacionamiento 1
   If cmb_FlgEst_Es1.ListIndex = -1 Then
      MsgBox "Debe seleccionar si hay valorización por Estacionamiento.", vbExclamation, modgen_g_str_NomPlt
      tab_Genera.Tab = 1
      Call gs_SetFocus(cmb_FlgEst_Es1)
      Exit Sub
   End If
   If cmb_FlgEst_Es1.ItemData(cmb_FlgEst_Es1.ListIndex) = 1 Then
      If ipp_AreTer_Es1.Value = 0 Then
         MsgBox "Debe ingresar el Area del Terreno.", vbExclamation, modgen_g_str_NomPlt
         tab_Genera.Tab = 1
         Call gs_SetFocus(ipp_AreTer_Es1)
         Exit Sub
      End If
      If ipp_ValCom_Es1.Value = 0 Then
         MsgBox "Debe ingresar el Valor Comercial.", vbExclamation, modgen_g_str_NomPlt
         tab_Genera.Tab = 1
         Call gs_SetFocus(ipp_ValCom_Es1)
         Exit Sub
      End If
      If ipp_ValRea_Es1.Value = 0 Then
         MsgBox "Debe ingresar el Valor de Realización.", vbExclamation, modgen_g_str_NomPlt
         tab_Genera.Tab = 1
         Call gs_SetFocus(ipp_ValRea_Es1)
         Exit Sub
      End If
      If ipp_ValTer_Es1.Value = 0 Then
         MsgBox "Debe ingresar el Valor del Terreno.", vbExclamation, modgen_g_str_NomPlt
         tab_Genera.Tab = 1
         Call gs_SetFocus(ipp_ValTer_Es1)
         Exit Sub
      End If
      If ipp_ValEdi_Es1.Value = 0 Then
         MsgBox "Debe ingresar el Valor de Edificación.", vbExclamation, modgen_g_str_NomPlt
         tab_Genera.Tab = 1
         Call gs_SetFocus(ipp_ValEdi_Es1)
         Exit Sub
      End If
   End If
   
   'Estacionamiento 2
   If cmb_FlgEst_Es2.ListIndex = -1 Then
      MsgBox "Debe seleccionar si hay valorización por Estacionamiento.", vbExclamation, modgen_g_str_NomPlt
      tab_Genera.Tab = 2
      Call gs_SetFocus(cmb_FlgEst_Es2)
      Exit Sub
   End If
   If cmb_FlgEst_Es2.ItemData(cmb_FlgEst_Es2.ListIndex) = 1 Then
      If ipp_AreTer_Es2.Value = 0 Then
         MsgBox "Debe ingresar el Area del Terreno.", vbExclamation, modgen_g_str_NomPlt
         tab_Genera.Tab = 2
         Call gs_SetFocus(ipp_AreTer_Es2)
         Exit Sub
      End If
      If ipp_ValCom_Es2.Value = 0 Then
         MsgBox "Debe ingresar el Valor Comercial.", vbExclamation, modgen_g_str_NomPlt
         tab_Genera.Tab = 2
         Call gs_SetFocus(ipp_ValCom_Es2)
         Exit Sub
      End If
      If ipp_ValRea_Es2.Value = 0 Then
         MsgBox "Debe ingresar el Valor de Realización.", vbExclamation, modgen_g_str_NomPlt
         tab_Genera.Tab = 2
         Call gs_SetFocus(ipp_ValRea_Es2)
         Exit Sub
      End If
      If ipp_ValTer_Es2.Value = 0 Then
         MsgBox "Debe ingresar el Valor del Terreno.", vbExclamation, modgen_g_str_NomPlt
         tab_Genera.Tab = 2
         Call gs_SetFocus(ipp_ValTer_Es2)
         Exit Sub
      End If
      If ipp_ValEdi_Es2.Value = 0 Then
         MsgBox "Debe ingresar el Valor de Edificación.", vbExclamation, modgen_g_str_NomPlt
         tab_Genera.Tab = 2
         Call gs_SetFocus(ipp_ValEdi_Es2)
         Exit Sub
      End If
   End If
   
   'Depósito
   If cmb_FlgEst_Dep.ListIndex = -1 Then
      MsgBox "Debe seleccionar si hay valorización por el Depósito.", vbExclamation, modgen_g_str_NomPlt
      tab_Genera.Tab = 3
      Call gs_SetFocus(cmb_FlgEst_Dep)
      Exit Sub
   End If
   If cmb_FlgEst_Dep.ItemData(cmb_FlgEst_Dep.ListIndex) = 1 Then
      If ipp_AreTer_Dep.Value = 0 Then
         MsgBox "Debe ingresar el Area del Terreno.", vbExclamation, modgen_g_str_NomPlt
         tab_Genera.Tab = 3
         Call gs_SetFocus(ipp_AreTer_Dep)
         Exit Sub
      End If
      'If ipp_AreCon_Dep.Value = 0 Then
      '   MsgBox "Debe ingresar el Area Construida.", vbExclamation, modgen_g_str_NomPlt
      '   tab_Genera.Tab = 3
      '   Call gs_SetFocus(ipp_AreCon_Dep)
      '   Exit Sub
      'End If
      'If ipp_SumAse_Dep.Value = 0 Then
      '   MsgBox "Debe ingresar la Suma Asegurada.", vbExclamation, modgen_g_str_NomPlt
      '   tab_Genera.Tab = 3
      '   Call gs_SetFocus(ipp_SumAse_Dep)
      '   Exit Sub
      'End If
      If ipp_ValCom_Dep.Value = 0 Then
         MsgBox "Debe ingresar el Valor Comercial.", vbExclamation, modgen_g_str_NomPlt
         tab_Genera.Tab = 3
         Call gs_SetFocus(ipp_ValCom_Dep)
         Exit Sub
      End If
      If ipp_ValRea_Dep.Value = 0 Then
         MsgBox "Debe ingresar el Valor de Realización.", vbExclamation, modgen_g_str_NomPlt
         tab_Genera.Tab = 3
         Call gs_SetFocus(ipp_ValRea_Dep)
         Exit Sub
      End If
      If ipp_ValTer_Dep.Value = 0 Then
         MsgBox "Debe ingresar el Valor del Terreno.", vbExclamation, modgen_g_str_NomPlt
         tab_Genera.Tab = 3
         Call gs_SetFocus(ipp_ValTer_Dep)
         Exit Sub
      End If
      If ipp_ValEdi_Dep.Value = 0 Then
         MsgBox "Debe ingresar el Valor de Edificación.", vbExclamation, modgen_g_str_NomPlt
         tab_Genera.Tab = 3
         Call gs_SetFocus(ipp_ValEdi_Dep)
         Exit Sub
      End If
'      If ipp_ValACo_Dep.Value = 0 Then
'         MsgBox "Debe ingresar el Valor de Areas Comunes.", vbExclamation, modgen_g_str_NomPlt
'         tab_Genera.Tab = 3
'         Call gs_SetFocus(ipp_ValACo_Dep)
'         Exit Sub
'      End If
   End If
   
   If l_dbl_MtoPre > CDbl(pnl_SumAse.Caption) Then
      MsgBox "Atención!!! El monto del préstamo no puede ser mayor que la suma asegurada." & vbCrLf & "Favor de verificar los datos registrados.", vbExclamation, modgen_g_str_NomPlt
   End If
   If CDbl(pnl_SumAse.Caption) = CDbl(pnl_ValCom.Caption) - CDbl(pnl_ValTer.Caption) Then
      MsgBox "Atención!!! La 'Suma Asegurada' debe ser igual a la diferencia del 'Valor Comercial' menos el 'Valor del Terreno'." & vbCrLf & "Favor de verificar los datos registrados.", vbExclamation, modgen_g_str_NomPlt
   End If
   If CDbl(pnl_ValRea.Caption) < CDbl(pnl_ValCom.Caption) * 0.85 Then
      MsgBox "Atención!!! El 'Valor de Realizacion' es menor al 85% del 'Valor Comercial'." & vbCrLf & "Favor de verificar los datos registrados.", vbExclamation, modgen_g_str_NomPlt
   End If
   
   'Obteniendo Monto de Préstamo
   r_dbl_MtoPre = 0
   r_dbl_IntCap = 0
   r_dbl_ComVta = 0
   
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      r_dbl_MtoPre = g_rst_Princi!SOLMAE_MTOPRE_MPR
      r_dbl_IntCap = g_rst_Princi!SOLMAE_INTGRA
      r_int_CuoExt = g_rst_Princi!SOLMAE_CUOEXT
      
      If moddat_g_int_TipMon = 1 Then
         r_dbl_ComVta = g_rst_Princi!SOLMAE_COMVTA_SOL
         r_dbl_ApoPro = g_rst_Princi!SOLMAE_APOPRO_SOL - g_rst_Princi!SOLMAE_MTOGCI
      Else
         r_dbl_ComVta = g_rst_Princi!SOLMAE_COMVTA_DOL
         r_dbl_ApoPro = g_rst_Princi!SOLMAE_APOPRO_DOL - g_rst_Princi!SOLMAE_MTOGCI
      End If
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Obteniendo Valor Asegurable del Inmueble
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM TRA_EVATAS "
   g_str_Parame = g_str_Parame & " WHERE EVATAS_NUMSOL = '" & moddat_g_str_NumSol & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      r_dbl_MtoTas = g_rst_Princi!EVATAS_SUMASE_INM + g_rst_Princi!EVATAS_SUMASE_ES1 + g_rst_Princi!EVATAS_SUMASE_ES2 + g_rst_Princi!EVATAS_SUMASE_DEP
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   r_dbl_TCaMpr = 0
   r_dbl_TipCam = 0
   
   If cmb_TipMon.ItemData(cmb_TipMon.ListIndex) <> 1 Then
      r_dbl_TipCam = CDbl(ipp_TipCam.Text)
   End If
   
   'Validando contra Parámetros de Productos
   If InStr(moddat_g_str_AgrCRC, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Then 'moddat_g_str_CodPrd = "001" Or moddat_g_str_CodPrd = "003" Then
      r_dbl_Valmax_ValEdi = 0
   
      If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "021") Then
         r_dbl_Valmax_ValEdi = l_arr_ParPrd(1).Genera_Cantid * moddat_gf_Consulta_ParVal("001", "002")
      End If
      
      'Validando Valor de Edificación
      If cmb_TipMon.ItemData(cmb_TipMon.ListIndex) <> 1 Then
         If CDbl(pnl_ValEdi.Caption) * r_dbl_TipCam > r_dbl_Valmax_ValEdi Then
            MsgBox "El Valor de Edificación excede el permitido para el Producto.", vbExclamation, modgen_g_str_NomPlt
            Exit Sub
         End If
      Else
         If CDbl(pnl_ValEdi.Caption) > r_dbl_Valmax_ValEdi Then
            MsgBox "El Valor de Edificación excede el permitido para el Producto.", vbExclamation, modgen_g_str_NomPlt
            Exit Sub
         End If
      End If
   End If
   
   If InStr(moddat_g_str_AgrCRC, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_AgrTFMV, moddat_g_str_CodPrd) > 0 Then   '"001" "003" "004" "007" "009" "010" "012" "013" "014" "015" "016" "017" "018" "019" "021" "022" "023"
      'Validando Relación Valor de Gravamen Hipotecario (Valor Comercial) / Monto del Préstamo
      r_dbl_PorMin_ValGrv = 0
   
      If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "031") Then
         r_dbl_PorMin_ValGrv = l_arr_ParPrd(1).Genera_Cantid
      End If
   
      If Not (CDbl(pnl_ValCom.Caption) / r_dbl_MtoPre * 100 >= r_dbl_PorMin_ValGrv) Then
         MsgBox "La relación Valor Comercial / Monto de Préstamo no cumple con el Parámetro permitido para el Producto." & Format(CDbl(pnl_ValCom.Caption) / r_dbl_MtoPre * 100, "##0.00") & "%", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
   
      'Determinando Valor de la Garantía
      If CDbl(pnl_ValCom.Caption) < r_dbl_ComVta Then
         r_dbl_ValGar = CDbl(pnl_ValCom.Caption)
      Else
         r_dbl_ValGar = r_dbl_ComVta
      End If
      
      'Validando Relación Monto del Préstamo / Valor de la Garantía
      r_dbl_PorMax_ValGar = 0
   
      If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "032") Then
         r_dbl_PorMax_ValGar = l_arr_ParPrd(1).Genera_Cantid
      End If
   
      If Not (r_dbl_MtoPre / r_dbl_ValGar * 100 <= r_dbl_PorMax_ValGar) Then
         MsgBox "La relación Monto del Préstamo / Valor de la Garantía no cumple con el Parámetro permitido para el Producto." & Format(r_dbl_MtoPre / r_dbl_ValGar * 100, "##0.00") & "%", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
   
      'Validando Relación Monto del Préstamo + Interés Capitalizado / Valor de la Garantía
      r_dbl_PorMax_ValGar = 0
   
      If moddat_gf_Consulta_ParSubPrd(l_arr_ParPrd, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "033") Then
         r_dbl_PorMax_ValGar = l_arr_ParPrd(1).Genera_Cantid
      End If
   
      If Not ((r_dbl_MtoPre + r_dbl_IntCap) / r_dbl_ValGar * 100 <= r_dbl_PorMax_ValGar) Then
         MsgBox "La relación Monto del Préstamo + Intereses Capitalizados / Valor de la Garantía no cumple con el Parámetro permitido para el Producto." & Format((r_dbl_MtoPre + r_dbl_IntCap) / r_dbl_ValGar * 100, "##0.00") & "%", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
   End If
   
   'miCasita - Validando Valor de Realización contra Monto de Préstamo
   r_int_FlgExc = 1
   
   If moddat_g_str_CodPrd = "002" Then
      If modgen_g_int_TipUsu = 18200 Or modgen_g_int_TipUsu = 18220 Then
         If r_dbl_MtoPre > CDbl(pnl_ValRea.Caption) Then
            If MsgBox("El Valor de Realización es menor al Monto del Préstamo. ¿Desea excepcionar esta validación?.", vbExclamation + vbQuestion + vbYesNo, modgen_g_str_NomPlt) <> vbYes Then
               Exit Sub
            End If
            
            r_int_FlgExc = 2
         End If
      Else
         If r_dbl_MtoPre > CDbl(pnl_ValRea.Caption) Then
            MsgBox "El Valor de Realización es menor al Monto del Préstamo, no se puede otorgar este crédito.", vbExclamation, modgen_g_str_NomPlt
            Exit Sub
         End If
      End If
   End If
   
   'Recalculando Cuota a Pagar
   r_dbl_CuoNue = 0
   r_dbl_ValViv = 0
   r_dbl_MtoPre = 0
   r_dbl_Portes = 0
   
   'Obteniendo Valor Asegurable del Inmueble
   r_dbl_ValViv = CDbl(ipp_SumAse_Inm.Value) + CDbl(ipp_SumAse_Es1.Value) + CDbl(ipp_SumAse_Es2.Value) + CDbl(ipp_SumAse_Dep.Value)
   
    'Obteniendo tasa y comision de cofide
   If InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Then
      'r_dbl_TasMVi = moddat_gf_ComMVi(moddat_g_str_CodPrd, 3, moddat_g_int_TipMon, l_int_PlaAno)
      r_dbl_ComCof = moddat_gf_ComMVi(moddat_g_str_CodPrd, 4, moddat_g_int_TipMon, l_int_PlaAno)
      r_dbl_TasCof = moddat_gf_ComMVi(moddat_g_str_CodPrd, 5, moddat_g_int_TipMon, l_int_PlaAno)
   ElseIf InStr(moddat_g_str_AgrTFMV, moddat_g_str_CodPrd) > 0 Then
      r_dbl_ComCof = moddat_gf_ComMVi(moddat_g_str_CodPrd, 4, moddat_g_int_TipMon, l_int_PlaAno)
      r_dbl_TasCof = moddat_gf_ComMVi(moddat_g_str_CodPrd, 5, moddat_g_int_TipMon, l_int_PlaAno)
   End If
   
   'Obtiene Tasas
   Call moddat_gf_Consulta_ValSeg(moddat_g_str_CodPrd, moddat_g_str_CodSub, l_str_EmpSeg, Format(l_int_TipSeg, "000"), moddat_g_int_TipMon, l_dbl_MtoPre, r_int_TipVal_Des, r_dbl_Import_Des, l_int_TasEsp)
   Call moddat_gf_Consulta_ValSeg(moddat_g_str_CodPrd, moddat_g_str_CodSub, l_str_EmpSeg, 0, moddat_g_int_TipMon, r_dbl_ValViv, r_int_TipVal_Viv, r_dbl_Import_Viv, l_int_TasEsp)
   
   If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub, "002", "401") Then
      r_dbl_Portes = moddat_g_arr_Genera(1).Genera_Cantid
   End If
   
   l_dbl_IntGra = 0
   
   Select Case moddat_g_str_CodPrd > 0
      
      Case InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd)
         'Para obtener el tope TC
         r_dbl_PorCon = 0
         If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "011") Then
            r_dbl_PorCon = moddat_g_arr_Genera(1).Genera_Cantid
         End If
         r_dbl_TopCon = 0
         If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "012") Then
            r_dbl_TopCon = moddat_g_arr_Genera(1).Genera_Cantid
         End If
         
         'NUEVA rutina de generacion de cronogramas
         int_Produc = 1
         int_CuoDbl = r_int_CuoExt
         dbl_ValInm = r_dbl_ComVta
         dbl_CuoIni = r_dbl_ApoPro
         dbl_MtoCon = (r_dbl_ComVta - r_dbl_ApoPro) * (r_dbl_PorCon / 100)
         If dbl_MtoCon > r_dbl_TopCon Then dbl_MtoCon = r_dbl_TopCon
         dbl_MtoTas = r_dbl_MtoTas
         int_PlaPre = l_int_PlaAno * 12
         dbl_TasInt = l_dbl_TasInt
         dbl_TasCof = r_dbl_TasCof
         dbl_ComCof = r_dbl_ComCof
         dat_FecDes = CDate(Format(date, "dd/mm/yyyy"))
         int_DiaVct = l_int_DiaPag
         int_PerGra = l_int_PerGra
         str_PriVct = ""
         dbl_Portes = r_dbl_Portes
         dbl_SegViv = r_dbl_Import_Viv
         int_TipSDe = l_int_TipSeg - 10
         dbl_SegDes = r_dbl_Import_Des
         
         'Calculando cronogramas
         Set obj_Cronog = CreateObject("ComCronograma.ClsCronograma")
         Call obj_Cronog.Listar(l_Arr_TNC_Cli(), l_Arr_TC_Cli(), l_Arr_TNC_Cof(), l_Arr_TC_Cof(), int_Produc, int_CuoDbl, dbl_ValInm, dbl_CuoIni, dbl_MtoCon, dbl_MtoTas, int_PlaPre, dbl_TasInt, dbl_TasCof, dbl_ComCof, dat_FecDes, 0, int_DiaVct, str_PriVct, int_PerGra, dbl_Portes, dbl_SegViv, int_TipSDe, dbl_SegDes)
         
         dbl_CuoMen = 0
         dbl_CuoPbp = 0
         dbl_IngReq = 0
         Call modgen_gf_Buscar_CuotaMensual(l_Arr_TNC_Cli(), l_Arr_TC_Cli(), int_CuoDbl, int_PerGra, int_Produc, dbl_CuoMen, dbl_CuoPbp, dbl_IngReq, moddat_g_str_CodPrd, moddat_g_str_CodSub)
         
         'muestra valor cuota
         r_dbl_CuoNue = Format(dbl_CuoPbp, "###,###,##0.00") & " "
         l_dbl_IntGra = l_Arr_TNC_Cli(int_PerGra, 10) - (dbl_ValInm - dbl_CuoIni - dbl_MtoCon)
      
      Case InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd)
         'NUEVA rutina de generacion de cronogramas
         int_Produc = 2
         int_CuoDbl = r_int_CuoExt
         dbl_ValInm = r_dbl_ComVta
         dbl_CuoIni = r_dbl_ApoPro
         dbl_MtoCon = 0
         dbl_MtoTas = r_dbl_MtoTas
         int_PlaPre = l_int_PlaAno * 12
         dbl_TasInt = l_dbl_TasInt
         dbl_TasCof = 0
         dbl_ComCof = 0
         dat_FecDes = CDate(Format(date, "dd/mm/yyyy"))
         int_DiaVct = l_int_DiaPag
         int_PerGra = l_int_PerGra
         str_PriVct = ""
         dbl_Portes = CDbl(r_dbl_Portes)
         dbl_SegViv = r_dbl_Import_Viv
         int_TipSDe = l_int_TipSeg - 10
         dbl_SegDes = r_dbl_Import_Des
         
         'Calculando cronogramas
         Set obj_Cronog = CreateObject("ComCronograma.ClsCronograma")
         Call obj_Cronog.Listar(l_Arr_TNC_Cli(), l_Arr_TC_Cli(), l_Arr_TNC_Cof(), l_Arr_TC_Cof(), int_Produc, int_CuoDbl, dbl_ValInm, dbl_CuoIni, dbl_MtoCon, dbl_MtoTas, int_PlaPre, dbl_TasInt, dbl_TasCof, dbl_ComCof, dat_FecDes, 0, int_DiaVct, str_PriVct, int_PerGra, dbl_Portes, dbl_SegViv, int_TipSDe, dbl_SegDes)
         
         dbl_CuoMen = 0
         dbl_CuoPbp = 0
         dbl_IngReq = 0
         
         Call modgen_gf_Buscar_CuotaMensual(l_Arr_TNC_Cli(), l_Arr_TC_Cli(), int_CuoDbl, int_PerGra, int_Produc, dbl_CuoMen, dbl_CuoPbp, dbl_IngReq, moddat_g_str_CodPrd, moddat_g_str_CodSub)
         
         'muestra valor cuota
         r_dbl_CuoNue = Format(dbl_CuoMen, "###,###,##0.00") & " "
         l_dbl_IntGra = l_Arr_TNC_Cli(int_PerGra, 10) - (dbl_ValInm - dbl_CuoIni - dbl_MtoCon)
         
      Case InStr(moddat_g_str_Agr2MIC, moddat_g_str_CodPrd) Or InStr(moddat_g_str_AgrMIHG, moddat_g_str_CodPrd) Or InStr(moddat_g_str_Agr2FMV, moddat_g_str_CodPrd)
         'Para obtener el tope TC
         r_dbl_TopCon = 0
         If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "012") Then
            r_dbl_TopCon = moddat_g_arr_Genera(1).Genera_Cantid
         End If
         If CDbl(r_dbl_ComVta) > (50 * moddat_gf_Consulta_ParVal("001", "002")) Then
            r_dbl_TopCon = 5000
         End If
         
         'NUEVA rutina de generacion de cronogramas
         int_Produc = 1
         int_CuoDbl = r_int_CuoExt
         dbl_ValInm = r_dbl_ComVta
         dbl_CuoIni = r_dbl_ApoPro
         dbl_MtoTas = r_dbl_MtoTas
         dbl_MtoCon = r_dbl_TopCon
         int_PlaPre = l_int_PlaAno * 12
         dbl_TasInt = l_dbl_TasInt
         dbl_TasCof = r_dbl_TasCof
         dbl_ComCof = r_dbl_ComCof
         dat_FecDes = CDate(Format(date, "dd/mm/yyyy"))
         int_DiaVct = l_int_DiaPag
         int_PerGra = l_int_PerGra
         str_PriVct = ""
         dbl_Portes = r_dbl_Portes
         dbl_SegViv = r_dbl_Import_Viv
         int_TipSDe = l_int_TipSeg - 10
         dbl_SegDes = r_dbl_Import_Des
         
         'Calculando cronogramas
         Set obj_Cronog = CreateObject("ComCronograma.ClsCronograma")
         Call obj_Cronog.Listar(l_Arr_TNC_Cli(), l_Arr_TC_Cli(), l_Arr_TNC_Cof(), l_Arr_TC_Cof(), int_Produc, int_CuoDbl, dbl_ValInm, dbl_CuoIni, dbl_MtoCon, dbl_MtoTas, int_PlaPre, dbl_TasInt, dbl_TasCof, dbl_ComCof, dat_FecDes, 0, int_DiaVct, str_PriVct, int_PerGra, dbl_Portes, dbl_SegViv, int_TipSDe, dbl_SegDes)
         
         dbl_CuoMen = 0
         dbl_CuoPbp = 0
         dbl_IngReq = 0
         Call modgen_gf_Buscar_CuotaMensual(l_Arr_TNC_Cli(), l_Arr_TC_Cli(), int_CuoDbl, int_PerGra, int_Produc, dbl_CuoMen, dbl_CuoPbp, dbl_IngReq, moddat_g_str_CodPrd, moddat_g_str_CodSub)
         
         'muestra valor cuota
         r_dbl_CuoNue = Format(dbl_CuoPbp, "###,###,##0.00") & " "
         l_dbl_IntGra = l_Arr_TNC_Cli(int_PerGra, 10) - (dbl_ValInm - dbl_CuoIni - dbl_MtoCon)
         
      Case InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd)
         'NUEVA rutina de generacion de cronogramas
         int_Produc = 3
         int_CuoDbl = r_int_CuoExt
         dbl_ValInm = r_dbl_ComVta
         dbl_CuoIni = r_dbl_ApoPro
         dbl_MtoCon = 0
         dbl_MtoTas = r_dbl_MtoTas
         int_PlaPre = l_int_PlaAno * 12
         dbl_TasInt = l_dbl_TasInt
         dbl_TasCof = r_dbl_TasCof
         dbl_ComCof = r_dbl_ComCof
         dat_FecDes = CDate(Format(date, "dd/mm/yyyy"))
         int_DiaVct = l_int_DiaPag
         int_PerGra = l_int_PerGra
         str_PriVct = ""
         dbl_Portes = CDbl(r_dbl_Portes)
         dbl_SegViv = r_dbl_Import_Viv
         int_TipSDe = l_int_TipSeg - 10
         dbl_SegDes = r_dbl_Import_Des
         
         'Calculando cronogramas
         Set obj_Cronog = CreateObject("ComCronograma.ClsCronograma")
         Call obj_Cronog.Listar(l_Arr_TNC_Cli(), l_Arr_TC_Cli(), l_Arr_TNC_Cof(), l_Arr_TC_Cof(), int_Produc, int_CuoDbl, dbl_ValInm, dbl_CuoIni, dbl_MtoCon, dbl_MtoTas, int_PlaPre, dbl_TasInt, dbl_TasCof, dbl_ComCof, dat_FecDes, 0, int_DiaVct, str_PriVct, int_PerGra, dbl_Portes, dbl_SegViv, int_TipSDe, dbl_SegDes)
         
         dbl_CuoMen = 0
         dbl_CuoPbp = 0
         dbl_IngReq = 0
         Call modgen_gf_Buscar_CuotaMensual(l_Arr_TNC_Cli(), l_Arr_TC_Cli(), int_CuoDbl, int_PerGra, int_Produc, dbl_CuoMen, dbl_CuoPbp, dbl_IngReq, moddat_g_str_CodPrd, moddat_g_str_CodSub)
         
         'muestra valor cuota
         r_dbl_CuoNue = Format(dbl_CuoMen, "###,###,##0.00") & " "
         l_dbl_IntGra = l_Arr_TNC_Cli(int_PerGra, 10) - (dbl_ValInm - dbl_CuoIni - dbl_MtoCon)
   End Select
   
   r_int_FlgExc_2 = 1

   If r_dbl_CuoNue > l_dbl_CuoApr Then
      If r_dbl_CuoNue > l_dbl_CuoAce And modgen_g_int_TipUsu <> 18200 And modgen_g_int_TipUsu <> 18000 Then
         'MsgBox "La Cuota obtenida (" & Format(r_dbl_CuoNue, "##0.00") & ") es mayor a la Cuota Aprobada (" & Format(l_dbl_CuoApr, "##0.00") & ").", vbExclamation, modgen_g_str_NomPlt
         'Exit Sub
         If MsgBox("La Cuota obtenida (" & Format(r_dbl_CuoNue, "##0.00") & ") es mayor a la Cuota Aprobada (" & Format(l_dbl_CuoApr, "##0.00") & "). ¿Desea aprobar esta excepción?", vbQuestion + vbDefaultButton2 + vbYesNo, modgen_g_str_NomPlt) <> vbYes Then
            Exit Sub
         End If
         r_int_FlgExc_2 = 2
      Else
         If MsgBox("La Cuota obtenida (" & Format(r_dbl_CuoNue, "##0.00") & ") es mayor a la Cuota Aprobada (" & Format(l_dbl_CuoApr, "##0.00") & "). ¿Desea aprobar esta excepción?", vbQuestion + vbDefaultButton2 + vbYesNo, modgen_g_str_NomPlt) <> vbYes Then
            Exit Sub
         End If
         r_int_FlgExc_2 = 2
      End If
   End If
   
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   'Registrando Excepción
   If r_int_FlgExc_2 = 2 Then
      Call fs_RegExc(2)
   End If
   
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      Screen.MousePointer = 11
      Call moddat_gs_FecSis
      
      g_str_Parame = "USP_TRA_EVATAS ("
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumSol & "', "
      g_str_Parame = g_str_Parame & "'" & l_arr_EmpPer(cmb_EmpPer.ListIndex + 1).Genera_Codigo & "', "
      g_str_Parame = g_str_Parame & "'" & l_arr_PerTas(cmb_PerTas.ListIndex + 1).Genera_Nombre & "', "
      g_str_Parame = g_str_Parame & "'" & l_arr_PerTas(cmb_PerTas.ListIndex + 1).Genera_Prefij & "', "
      g_str_Parame = g_str_Parame & "'" & txt_NumInf.Text & "', "
      g_str_Parame = g_str_Parame & Format(CDate(ipp_FecEva.Text), "yyyymmdd") & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_AnoCon.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_NumPis.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_NumSot.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(cmb_TipInm.ItemData(cmb_TipInm.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & CStr(cmb_UsoInm.ItemData(cmb_UsoInm.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & CStr(cmb_MatCon.ItemData(cmb_MatCon.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & CStr(cmb_TipMon.ItemData(cmb_TipMon.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_TipCam.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_TCaMpr) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_AreTer_Inm.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_AreCon_Inm.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_SumAse_Inm.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValCom_Inm.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValRea_Inm.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValTer_Inm.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValEdi_Inm.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValACo_Inm.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(cmb_FlgEst_Es1.ItemData(cmb_FlgEst_Es1.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_AreTer_Es1.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_AreCon_Es1.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_SumAse_Es1.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValCom_Es1.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValRea_Es1.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValTer_Es1.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValEdi_Es1.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValACo_Es1.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(cmb_FlgEst_Es2.ItemData(cmb_FlgEst_Es2.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_AreTer_Es2.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_AreCon_Es2.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_SumAse_Es2.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValCom_Es2.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValRea_Es2.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValTer_Es2.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValEdi_Es2.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValACo_Es2.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(cmb_FlgEst_Dep.ItemData(cmb_FlgEst_Dep.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_AreTer_Dep.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_AreCon_Dep.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_SumAse_Dep.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValCom_Dep.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValRea_Dep.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValTer_Dep.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValEdi_Dep.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_ValACo_Dep.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(cmb_TieAzo.ItemData(cmb_TieAzo.ListIndex)) & ", "
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

   'Registrando la Excepción
   If r_int_FlgExc = 2 Then
      Call fs_RegExc(1)
   End If
   
   'Creando Nueva Ocurrencia en Detalle de Seguimiento
   If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, 41, 34, 0, "", 0, 0) Then
      Exit Sub
   End If
   
   moddat_g_int_FlgAct = 2
   Unload Me
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   pnl_NumSol.Caption = gf_Formato_NumSol(moddat_g_str_NumSol)
   pnl_FecIng.Caption = moddat_g_str_FecIng
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & " - " & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   
   Call fs_Inicia
   Call fs_Buscar_DatEva
   Call gs_CentraForm(Me)
   
   Call gs_SetFocus(cmb_EmpPer)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_LisIte(cmb_EmpPer, l_arr_EmpPer, 1, "507")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipInm, 1, "221")
   Call moddat_gs_Carga_LisIte_Combo(cmb_UsoInm, 1, "222")
   Call moddat_gs_Carga_LisIte_Combo(cmb_MatCon, 1, "223")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipMon, 1, "204")
   Call moddat_gs_Carga_LisIte_Combo(cmb_FlgEst_Es1, 1, "214")
   Call moddat_gs_Carga_LisIte_Combo(cmb_FlgEst_Es2, 1, "214")
   Call moddat_gs_Carga_LisIte_Combo(cmb_FlgEst_Dep, 1, "214")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TieAzo, 1, "214")

   cmb_EmpPer.ListIndex = -1
   cmb_PerTas.Clear
   txt_NumInf.Text = ""
   ipp_FecEva.Text = Format(date, "dd/mm/yyyy")
   cmb_TipMon.ListIndex = -1
   ipp_TipCam.Value = 0
   ipp_AreTer_Inm.Value = 0
   ipp_AreCon_Inm.Value = 0
   ipp_SumAse_Inm.Value = 0
   ipp_ValCom_Inm.Value = 0
   ipp_ValRea_Inm.Value = 0
   ipp_ValTer_Inm.Value = 0
   ipp_ValEdi_Inm.Value = 0
   ipp_ValACo_Inm.Value = 0

   cmb_FlgEst_Es1.ListIndex = -1
   ipp_AreTer_Es1.Value = 0
   ipp_AreCon_Es1.Value = 0
   ipp_SumAse_Es1.Value = 0
   ipp_ValCom_Es1.Value = 0
   ipp_ValRea_Es1.Value = 0
   ipp_ValTer_Es1.Value = 0
   ipp_ValEdi_Es1.Value = 0
   ipp_ValACo_Es1.Value = 0

   ipp_AreTer_Es1.Enabled = False
   ipp_AreCon_Es1.Enabled = False
   ipp_SumAse_Es1.Enabled = False
   ipp_ValCom_Es1.Enabled = False
   ipp_ValRea_Es1.Enabled = False
   ipp_ValTer_Es1.Enabled = False
   ipp_ValEdi_Es1.Enabled = False
   ipp_ValACo_Es1.Enabled = False

   cmb_FlgEst_Es2.ListIndex = -1
   ipp_AreTer_Es2.Value = 0
   ipp_AreCon_Es2.Value = 0
   ipp_SumAse_Es2.Value = 0
   ipp_ValCom_Es2.Value = 0
   ipp_ValRea_Es2.Value = 0
   ipp_ValTer_Es2.Value = 0
   ipp_ValEdi_Es2.Value = 0
   ipp_ValACo_Es2.Value = 0

   ipp_AreTer_Es2.Enabled = False
   ipp_AreCon_Es2.Enabled = False
   ipp_SumAse_Es2.Enabled = False
   ipp_ValCom_Es2.Enabled = False
   ipp_ValRea_Es2.Enabled = False
   ipp_ValTer_Es2.Enabled = False
   ipp_ValEdi_Es2.Enabled = False
   ipp_ValACo_Es2.Enabled = False

   cmb_FlgEst_Dep.ListIndex = -1
   ipp_AreTer_Dep.Value = 0
   ipp_AreCon_Dep.Value = 0
   ipp_SumAse_Dep.Value = 0
   ipp_ValCom_Dep.Value = 0
   ipp_ValRea_Dep.Value = 0
   ipp_ValTer_Dep.Value = 0
   ipp_ValEdi_Dep.Value = 0
   ipp_ValACo_Dep.Value = 0

   ipp_AreTer_Dep.Enabled = False
   ipp_AreCon_Dep.Enabled = False
   ipp_SumAse_Dep.Enabled = False
   ipp_ValCom_Dep.Enabled = False
   ipp_ValRea_Dep.Enabled = False
   ipp_ValTer_Dep.Enabled = False
   ipp_ValEdi_Dep.Enabled = False
   ipp_ValACo_Dep.Enabled = False

   cmb_TieAzo.ListIndex = -1
   'Obteniendo Datos para Recálculo de Cuota
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      l_dbl_TasInt = g_rst_Princi!SOLMAE_TASINT
      l_dbl_MtoPre = g_rst_Princi!SOLMAE_MTOPRE_MPR
      l_int_PlaAno = g_rst_Princi!SOLMAE_PLAANO
      l_int_PerGra = g_rst_Princi!SOLMAE_PERGRA
      l_str_EmpSeg = Trim(g_rst_Princi!SOLMAE_ESGDES)
      l_int_TipSeg = g_rst_Princi!SOLMAE_TIPSEG
      l_int_DiaPag = g_rst_Princi!SOLMAE_DIAPAG
      l_dbl_CuoApr = g_rst_Princi!SOLMAE_CUOMEN_MPR
      l_dbl_CuoAce = g_rst_Princi!SOLMAE_CUOAPR_MPR
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Buscar_DatEva()
Dim l_int_DatCyg     As Integer
   
   l_int_DatCyg = 0
   moddat_g_int_FlgGrb = 1
   moddat_g_int_CygTDo = 0
   moddat_g_str_CygNDo = ""
   
   'Datos del cliente
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM CLI_DATGEN "
   g_str_Parame = g_str_Parame & " WHERE DATGEN_TIPDOC = " & Trim(moddat_g_int_TipDoc) & " "
   g_str_Parame = g_str_Parame & "   AND DATGEN_NUMDOC = '" & Trim(moddat_g_str_NumDoc) & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      If (g_rst_Princi!DATGEN_ESTCIV = 2 And g_rst_Princi!DATGEN_REGCYG = 1) Or g_rst_Princi!DATGEN_ESTCIV = 5 Then
         l_int_DatCyg = 2
         moddat_g_int_CygTDo = g_rst_Princi!DATGEN_CYGTDO
         moddat_g_str_CygNDo = Trim(g_rst_Princi!DATGEN_CYGNDO & "")
      End If
      
      l_str_CodCiu = g_rst_Princi!DATGEN_CODCIU
      If CInt(l_str_CodCiu) = 0 Then
         MsgBox "El codigo CIIU del cliente debe estar registrado, favor de coordinarlo con Comercial.", vbExclamation, modgen_g_str_NomPlt
      End If
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Buscar datos del Cónyuge
   If l_int_DatCyg = 2 Then
      g_str_Parame = "SELECT * FROM CLI_DATGEN WHERE "
      g_str_Parame = g_str_Parame & "DATGEN_TIPDOC = " & CStr(moddat_g_int_CygTDo) & " AND "
      g_str_Parame = g_str_Parame & "DATGEN_NUMDOC = '" & moddat_g_str_CygNDo & "' "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         g_rst_Princi.MoveFirst
         If CInt(l_str_CodCiu) <> 7522 And CInt(l_str_CodCiu) <> 7523 Then
            l_str_CodCiu = g_rst_Princi!DATGEN_CODCIU
         End If
      End If
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   End If
   
   'Datos de la solicitud
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      l_int_TasEsp = g_rst_Princi!SOLMAE_TASESP
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Datos de la tasacion
   g_str_Parame = "SELECT * FROM TRA_EVATAS WHERE "
   g_str_Parame = g_str_Parame & "EVATAS_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      moddat_g_int_FlgGrb = 2
      cmb_EmpPer.ListIndex = gf_Busca_Arregl(l_arr_EmpPer, g_rst_Princi!EVATAS_CODEMP) - 1
      Call moddat_gs_Carga_PerTas(cmb_PerTas, l_arr_PerTas(), l_arr_EmpPer(cmb_EmpPer.ListIndex + 1).Genera_Codigo)
      Call gs_BuscarCombo(cmb_PerTas, Trim(g_rst_Princi!EVATAS_NOMPER & ""))
      
      txt_NumInf.Text = Trim(g_rst_Princi!EVATAS_NUMINF & "")
      ipp_FecEva.Text = gf_FormatoFecha(CStr(g_rst_Princi!EVATAS_FECEVA))
      ipp_AnoCon.Value = g_rst_Princi!EVATAS_ANOCON
      ipp_NumPis.Value = g_rst_Princi!EVATAS_NUMPIS
      ipp_NumSot.Value = g_rst_Princi!EVATAS_NUMSOT
      
      Call gs_BuscarCombo_Item(cmb_TipInm, g_rst_Princi!EVATAS_TIPINM)
      Call gs_BuscarCombo_Item(cmb_UsoInm, g_rst_Princi!EVATAS_USOINM)
      Call gs_BuscarCombo_Item(cmb_MatCon, g_rst_Princi!EVATAS_MATCON)
      Call gs_BuscarCombo_Item(cmb_TipMon, g_rst_Princi!EVATAS_TIPMON)
      Call gs_BuscarCombo_Item(cmb_TieAzo, IIf(IsNull(g_rst_Princi!EVATAS_FLGAZO_DEP), 0, g_rst_Princi!EVATAS_FLGAZO_DEP))
      
      ipp_TipCam.Value = g_rst_Princi!EVATAS_TIPCAM
      ipp_AreTer_Inm.Value = g_rst_Princi!EVATAS_ARETER_INM
      ipp_AreCon_Inm.Value = g_rst_Princi!EVATAS_ARECON_INM
      ipp_SumAse_Inm.Value = g_rst_Princi!EVATAS_SUMASE_INM
      ipp_ValCom_Inm.Value = g_rst_Princi!EVATAS_VALCOM_INM
      ipp_ValRea_Inm.Value = g_rst_Princi!EVATAS_VALREA_INM
      ipp_ValTer_Inm.Value = g_rst_Princi!EVATAS_VALTER_INM
      ipp_ValEdi_Inm.Value = g_rst_Princi!EVATAS_VALEDI_INM
      ipp_ValACo_Inm.Value = g_rst_Princi!EVATAS_VALACO_INM
      
      Call gs_BuscarCombo_Item(cmb_FlgEst_Es1, g_rst_Princi!EVATAS_FLGEST_ES1)
      If cmb_FlgEst_Es1.ItemData(cmb_FlgEst_Es1.ListIndex) = 1 Then
         ipp_AreTer_Es1.Value = g_rst_Princi!EVATAS_ARETER_ES1
         ipp_AreCon_Es1.Value = g_rst_Princi!EVATAS_ARECON_ES1
         ipp_SumAse_Es1.Value = g_rst_Princi!EVATAS_SUMASE_ES1
         ipp_ValCom_Es1.Value = g_rst_Princi!EVATAS_VALCOM_ES1
         ipp_ValRea_Es1.Value = g_rst_Princi!EVATAS_VALREA_ES1
         ipp_ValTer_Es1.Value = g_rst_Princi!EVATAS_VALTER_ES1
         ipp_ValEdi_Es1.Value = g_rst_Princi!EVATAS_VALEDI_ES1
         ipp_ValACo_Es1.Value = g_rst_Princi!EVATAS_VALACO_ES1
         ipp_AreTer_Es1.Enabled = True
         ipp_AreCon_Es1.Enabled = True
         ipp_SumAse_Es1.Enabled = True
         ipp_ValCom_Es1.Enabled = True
         ipp_ValRea_Es1.Enabled = True
         ipp_ValTer_Es1.Enabled = True
         ipp_ValEdi_Es1.Enabled = True
         ipp_ValACo_Es1.Enabled = True
      End If
   
      Call gs_BuscarCombo_Item(cmb_FlgEst_Es2, g_rst_Princi!EVATAS_FLGEST_ES2)
      If cmb_FlgEst_Es2.ItemData(cmb_FlgEst_Es2.ListIndex) = 1 Then
         ipp_AreTer_Es2.Value = g_rst_Princi!EVATAS_ARETER_ES2
         ipp_AreCon_Es2.Value = g_rst_Princi!EVATAS_ARECON_ES2
         ipp_SumAse_Es2.Value = g_rst_Princi!EVATAS_SUMASE_ES2
         ipp_ValCom_Es2.Value = g_rst_Princi!EVATAS_VALCOM_ES2
         ipp_ValRea_Es2.Value = g_rst_Princi!EVATAS_VALREA_ES2
         ipp_ValTer_Es2.Value = g_rst_Princi!EVATAS_VALTER_ES2
         ipp_ValEdi_Es2.Value = g_rst_Princi!EVATAS_VALEDI_ES2
         ipp_ValACo_Es2.Value = g_rst_Princi!EVATAS_VALACO_ES2
         ipp_AreTer_Es2.Enabled = True
         ipp_AreCon_Es2.Enabled = True
         ipp_SumAse_Es2.Enabled = True
         ipp_ValCom_Es2.Enabled = True
         ipp_ValRea_Es2.Enabled = True
         ipp_ValTer_Es2.Enabled = True
         ipp_ValEdi_Es2.Enabled = True
         ipp_ValACo_Es2.Enabled = True
      End If
      
      Call gs_BuscarCombo_Item(cmb_FlgEst_Dep, g_rst_Princi!EVATAS_FLGEST_DEP)
      If cmb_FlgEst_Dep.ItemData(cmb_FlgEst_Dep.ListIndex) = 1 Then
         ipp_AreTer_Dep.Value = g_rst_Princi!EVATAS_ARETER_DEP
         ipp_AreCon_Dep.Value = g_rst_Princi!EVATAS_ARECON_DEP
         ipp_SumAse_Dep.Value = g_rst_Princi!EVATAS_SUMASE_DEP
         ipp_ValCom_Dep.Value = g_rst_Princi!EVATAS_VALCOM_DEP
         ipp_ValRea_Dep.Value = g_rst_Princi!EVATAS_VALREA_DEP
         ipp_ValTer_Dep.Value = g_rst_Princi!EVATAS_VALTER_DEP
         ipp_ValEdi_Dep.Value = g_rst_Princi!EVATAS_VALEDI_DEP
         ipp_ValACo_Dep.Value = g_rst_Princi!EVATAS_VALACO_DEP
         ipp_AreTer_Dep.Enabled = True
         ipp_AreCon_Dep.Enabled = True
         ipp_SumAse_Dep.Enabled = True
         ipp_ValCom_Dep.Enabled = True
         ipp_ValRea_Dep.Enabled = True
         ipp_ValTer_Dep.Enabled = True
         ipp_ValEdi_Dep.Enabled = True
         ipp_ValACo_Dep.Enabled = True
      End If
      
      Call fs_Calcul
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   tab_Genera.Tab = 0
End Sub

Private Sub cmb_EmpPer_Click()
   If cmb_EmpPer.ListIndex > -1 Then
      Screen.MousePointer = 11
      Call moddat_gs_Carga_PerTas(cmb_PerTas, l_arr_PerTas(), l_arr_EmpPer(cmb_EmpPer.ListIndex + 1).Genera_Codigo)
      Screen.MousePointer = 0
   End If
   
   Call gs_SetFocus(cmb_PerTas)
End Sub

Private Sub cmb_EmpPer_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_EmpPer_Click
   End If
End Sub

Private Sub cmb_FlgEst_Dep_Click()
   If cmb_FlgEst_Dep.ListIndex > -1 Then
      If cmb_FlgEst_Dep.ItemData(cmb_FlgEst_Dep.ListIndex) = 1 Then
         ipp_AreTer_Dep.Enabled = True
         ipp_AreCon_Dep.Enabled = True
         ipp_SumAse_Dep.Enabled = True
         ipp_ValCom_Dep.Enabled = True
         ipp_ValRea_Dep.Enabled = True
         ipp_ValTer_Dep.Enabled = True
         ipp_ValEdi_Dep.Enabled = True
         ipp_ValACo_Dep.Enabled = True
         
         Call gs_SetFocus(ipp_AreTer_Dep)
      Else
         ipp_AreTer_Dep.Enabled = False
         ipp_AreCon_Dep.Enabled = False
         ipp_SumAse_Dep.Enabled = False
         ipp_ValCom_Dep.Enabled = False
         ipp_ValRea_Dep.Enabled = False
         ipp_ValTer_Dep.Enabled = False
         ipp_ValEdi_Dep.Enabled = False
         ipp_ValACo_Dep.Enabled = False
      
         ipp_AreTer_Dep.Value = 0
         ipp_AreCon_Dep.Value = 0
         ipp_SumAse_Dep.Value = 0
         ipp_ValCom_Dep.Value = 0
         ipp_ValRea_Dep.Value = 0
         ipp_ValTer_Dep.Value = 0
         ipp_ValEdi_Dep.Value = 0
         ipp_ValACo_Dep.Value = 0
         
         Call gs_SetFocus(cmd_Grabar)
      End If
   End If
End Sub

Private Sub cmb_FlgEst_Dep_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_FlgEst_Dep_Click
   End If
End Sub

Private Sub cmb_FlgEst_Es1_Click()
   If cmb_FlgEst_Es1.ListIndex > -1 Then
      If cmb_FlgEst_Es1.ItemData(cmb_FlgEst_Es1.ListIndex) = 1 Then
         ipp_AreTer_Es1.Enabled = True
         ipp_AreCon_Es1.Enabled = True
         ipp_SumAse_Es1.Enabled = True
         ipp_ValCom_Es1.Enabled = True
         ipp_ValRea_Es1.Enabled = True
         ipp_ValTer_Es1.Enabled = True
         ipp_ValEdi_Es1.Enabled = True
         ipp_ValACo_Es1.Enabled = True
         
         Call gs_SetFocus(ipp_AreTer_Es1)
      Else
         ipp_AreTer_Es1.Enabled = False
         ipp_AreCon_Es1.Enabled = False
         ipp_SumAse_Es1.Enabled = False
         ipp_ValCom_Es1.Enabled = False
         ipp_ValRea_Es1.Enabled = False
         ipp_ValTer_Es1.Enabled = False
         ipp_ValEdi_Es1.Enabled = False
         ipp_ValACo_Es1.Enabled = False
      
         ipp_AreTer_Es1.Value = 0
         ipp_AreCon_Es1.Value = 0
         ipp_SumAse_Es1.Value = 0
         ipp_ValCom_Es1.Value = 0
         ipp_ValRea_Es1.Value = 0
         ipp_ValTer_Es1.Value = 0
         ipp_ValEdi_Es1.Value = 0
         ipp_ValACo_Es1.Value = 0
         
         tab_Genera.Tab = 2
         Call gs_SetFocus(cmb_FlgEst_Es2)
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
         ipp_AreTer_Es2.Enabled = True
         ipp_AreCon_Es2.Enabled = True
         ipp_SumAse_Es2.Enabled = True
         ipp_ValCom_Es2.Enabled = True
         ipp_ValRea_Es2.Enabled = True
         ipp_ValTer_Es2.Enabled = True
         ipp_ValEdi_Es2.Enabled = True
         ipp_ValACo_Es2.Enabled = True
         
         Call gs_SetFocus(ipp_AreTer_Es2)
      Else
         ipp_AreTer_Es2.Enabled = False
         ipp_AreCon_Es2.Enabled = False
         ipp_SumAse_Es2.Enabled = False
         ipp_ValCom_Es2.Enabled = False
         ipp_ValRea_Es2.Enabled = False
         ipp_ValTer_Es2.Enabled = False
         ipp_ValEdi_Es2.Enabled = False
         ipp_ValACo_Es2.Enabled = False
      
         ipp_AreTer_Es2.Value = 0
         ipp_AreCon_Es2.Value = 0
         ipp_SumAse_Es2.Value = 0
         ipp_ValCom_Es2.Value = 0
         ipp_ValRea_Es2.Value = 0
         ipp_ValTer_Es2.Value = 0
         ipp_ValEdi_Es2.Value = 0
         ipp_ValACo_Es2.Value = 0
         
         tab_Genera.Tab = 3
         Call gs_SetFocus(cmb_FlgEst_Dep)
      End If
   End If
End Sub

Private Sub cmb_FlgEst_Es2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_FlgEst_Es2_Click
   End If
End Sub

Private Sub cmb_MatCon_Click()
   Call gs_SetFocus(cmb_TipMon)
End Sub

Private Sub cmb_MatCon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_MatCon_Click
   End If
End Sub

Private Sub cmb_PerTas_Click()
   Call gs_SetFocus(txt_NumInf)
End Sub

Private Sub cmb_PerTas_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_PerTas_Click
   End If
End Sub

Private Sub cmb_TipInm_Click()
   Call gs_SetFocus(cmb_UsoInm)
End Sub

Private Sub cmb_TipInm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipInm_Click
   End If
End Sub

Private Sub cmb_TipMon_Click()
   Call gs_SetFocus(ipp_TipCam)
End Sub

Private Sub cmb_TipMon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipMon_Click
   End If
End Sub

Private Sub cmb_UsoInm_Click()
   Call gs_SetFocus(cmb_MatCon)
End Sub

Private Sub cmb_UsoInm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_UsoInm_Click
   End If
End Sub

Private Sub ipp_AnoCon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_NumPis)
   End If
End Sub

Private Sub ipp_AreCon_Inm_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_AreCon_Inm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TieAzo) 'Call gs_SetFocus(ipp_SumAse_Inm)
   End If
End Sub

Private Sub ipp_AreTer_Inm_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_AreTer_Inm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_AreCon_Inm)
   End If
End Sub

Private Sub ipp_NumPis_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_NumSot)
   End If
End Sub

Private Sub ipp_NumSot_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipInm)
   End If
End Sub

Private Sub ipp_SumAse_Inm_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_SumAse_Inm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValCom_Inm)
   End If
End Sub

Private Sub ipp_ValCom_Inm_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValCom_Inm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValRea_Inm)
   End If
End Sub

Private Sub ipp_ValRea_Inm_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValRea_Inm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValTer_Inm)
   End If
End Sub

Private Sub ipp_ValTer_Inm_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValTer_Inm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValEdi_Inm)
   End If
End Sub

Private Sub ipp_ValEdi_Inm_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValEdi_Inm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValACo_Inm)
   End If
End Sub

Private Sub ipp_ValACo_Inm_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValACo_Inm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      tab_Genera.Tab = 1
      Call gs_SetFocus(cmb_FlgEst_Es1)
   End If
End Sub

Private Sub ipp_AreCon_Es1_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_AreCon_Es1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_SumAse_Es1)
   End If
End Sub

Private Sub ipp_AreTer_Es1_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_AreTer_Es1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_AreCon_Es1)
   End If
End Sub

Private Sub ipp_SumAse_Es1_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_SumAse_Es1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValCom_Es1)
   End If
End Sub

Private Sub ipp_ValCom_Es1_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValCom_Es1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValRea_Es1)
   End If
End Sub

Private Sub ipp_ValRea_Es1_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValRea_Es1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValTer_Es1)
   End If
End Sub

Private Sub ipp_ValTer_Es1_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValTer_Es1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValEdi_Es1)
   End If
End Sub

Private Sub ipp_ValEdi_Es1_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValEdi_Es1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValACo_Es1)
   End If
End Sub

Private Sub ipp_ValACo_Es1_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValACo_Es1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      tab_Genera.Tab = 2
      Call gs_SetFocus(cmb_FlgEst_Es2)
   End If
End Sub

Private Sub ipp_AreCon_Es2_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_AreCon_Es2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_SumAse_Es2)
   End If
End Sub

Private Sub ipp_AreTer_Es2_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_AreTer_Es2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_AreCon_Es2)
   End If
End Sub

Private Sub ipp_SumAse_Es2_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_SumAse_Es2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValCom_Es2)
   End If
End Sub

Private Sub ipp_ValCom_Es2_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValCom_Es2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValRea_Es2)
   End If
End Sub

Private Sub ipp_ValRea_Es2_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValRea_Es2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValTer_Es2)
   End If
End Sub

Private Sub ipp_ValTer_Es2_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValTer_Es2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValEdi_Es2)
   End If
End Sub

Private Sub ipp_ValEdi_Es2_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValEdi_Es2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValACo_Es2)
   End If
End Sub

Private Sub ipp_ValACo_Es2_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValACo_Es2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      tab_Genera.Tab = 1
      Call gs_SetFocus(cmb_FlgEst_Dep)
   End If
End Sub

Private Sub ipp_AreCon_Dep_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_AreCon_Dep_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_SumAse_Dep)
   End If
End Sub

Private Sub ipp_AreTer_Dep_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_AreTer_Dep_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_AreCon_Dep)
   End If
End Sub

Private Sub ipp_SumAse_Dep_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_SumAse_Dep_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValCom_Dep)
   End If
End Sub

Private Sub ipp_ValCom_Dep_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValCom_Dep_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValRea_Dep)
   End If
End Sub

Private Sub ipp_ValRea_Dep_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValRea_Dep_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValTer_Dep)
   End If
End Sub

Private Sub ipp_ValTer_Dep_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValTer_Dep_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValEdi_Dep)
   End If
End Sub

Private Sub ipp_ValEdi_Dep_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValEdi_Dep_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ValACo_Dep)
   End If
End Sub

Private Sub ipp_ValACo_Dep_Change()
   Call fs_Calcul
End Sub

Private Sub ipp_ValACo_Dep_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      tab_Genera.Tab = 1
      Call gs_SetFocus(cmd_Grabar)
   End If
End Sub

Private Sub ipp_FecEva_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_AnoCon)
   End If
End Sub

Private Sub ipp_TipCam_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      tab_Genera.Tab = 0
      Call gs_SetFocus(ipp_AreTer_Inm)
   End If
End Sub

Private Sub txt_NumInf_GotFocus()
   Call gs_SelecTodo(txt_NumInf)
End Sub

Private Sub txt_NumInf_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecEva)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-/():;:;")
   End If
End Sub

Private Sub fs_Calcul()
   pnl_AreTer.Caption = Format(CDbl(ipp_AreTer_Inm.Text) + CDbl(ipp_AreTer_Es1.Text) + CDbl(ipp_AreTer_Es2.Text) + CDbl(ipp_AreTer_Dep.Text), "###,###,##0.00") & " "
   pnl_AreCon.Caption = Format(CDbl(ipp_AreCon_Inm.Text) + CDbl(ipp_AreCon_Es1.Text) + CDbl(ipp_AreCon_Es2.Text) + CDbl(ipp_AreCon_Dep.Text), "###,###,##0.00") & " "
   pnl_SumAse.Caption = Format(CDbl(ipp_SumAse_Inm.Text) + CDbl(ipp_SumAse_Es1.Text) + CDbl(ipp_SumAse_Es2.Text) + CDbl(ipp_SumAse_Dep.Text), "###,###,##0.00") & " "
   pnl_ValCom.Caption = Format(CDbl(ipp_ValCom_Inm.Text) + CDbl(ipp_ValCom_Es1.Text) + CDbl(ipp_ValCom_Es2.Text) + CDbl(ipp_ValCom_Dep.Text), "###,###,##0.00") & " "
   pnl_ValRea.Caption = Format(CDbl(ipp_ValRea_Inm.Text) + CDbl(ipp_ValRea_Es1.Text) + CDbl(ipp_ValRea_Es2.Text) + CDbl(ipp_ValRea_Dep.Text), "###,###,##0.00") & " "
   pnl_ValTer.Caption = Format(CDbl(ipp_ValTer_Inm.Text) + CDbl(ipp_ValTer_Es1.Text) + CDbl(ipp_ValTer_Es2.Text) + CDbl(ipp_ValTer_Dep.Text), "###,###,##0.00") & " "
   pnl_ValEdi.Caption = Format(CDbl(ipp_ValEdi_Inm.Text) + CDbl(ipp_ValEdi_Es1.Text) + CDbl(ipp_ValEdi_Es2.Text) + CDbl(ipp_ValEdi_Dep.Text), "###,###,##0.00") & " "
   pnl_ValACo.Caption = Format(CDbl(ipp_ValACo_Inm.Text) + CDbl(ipp_ValACo_Es1.Text) + CDbl(ipp_ValACo_Es2.Text) + CDbl(ipp_ValACo_Dep.Text), "###,###,##0.00") & " "
End Sub

Private Sub fs_RegExc(ByVal p_TipExc As Integer)
   Dim r_int_NumExc     As Integer
   Dim r_int_NivAut     As Integer
   Dim r_str_DesExc     As String

   If modgen_g_int_TipUsu = 18200 Then
      r_int_NivAut = 13
   Else
      r_int_NivAut = 31
   End If

   'Generando Número de Excepción
   r_int_NumExc = 0
   
   g_str_Parame = "SELECT COUNT(SEGEXC_NUMSOL) AS NUMREG FROM TRA_SEGEXC WHERE "
   g_str_Parame = g_str_Parame & "SEGEXC_NUMSOL = '" & moddat_g_str_NumSol & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If

   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      r_int_NumExc = g_rst_Princi!NUMREG
   End If
      
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
      
   r_int_NumExc = r_int_NumExc + 1
   
   'Grabando en Detalle de Seguimiento
   If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, 41, 18, 0, "", 0, 0) Then
      Exit Sub
   End If
   
   'Grabando en Detalle de Excepciones
   If p_TipExc = 1 Then
      r_str_DesExc = "RATIO VALOR REALIZACION MENOR A MONTO DEL PRESTAMO."
   Else
      r_str_DesExc = "CUOTA MAYOR A LA APROBADA."
   End If
   
   If Not moddat_gf_Inserta_SegExc(moddat_g_str_NumSol, 41, r_int_NumExc, r_str_DesExc, r_int_NivAut) Then
      Exit Sub
   End If
End Sub

