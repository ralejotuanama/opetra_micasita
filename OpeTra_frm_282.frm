VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_Tra_EvaSeg_03 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   7680
   ClientLeft      =   2355
   ClientTop       =   1665
   ClientWidth     =   11250
   Icon            =   "OpeTra_frm_282.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   11250
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   7680
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   11250
      _Version        =   65536
      _ExtentX        =   19844
      _ExtentY        =   13547
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
         Height          =   1305
         Left            =   30
         TabIndex        =   11
         Top             =   6330
         Width           =   11175
         _Version        =   65536
         _ExtentX        =   19711
         _ExtentY        =   2302
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
         Begin VB.TextBox txt_ObsEva 
            Height          =   915
            Left            =   60
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   7
            Text            =   "OpeTra_frm_282.frx":000C
            Top             =   330
            Width           =   11085
         End
         Begin VB.Label Label9 
            Caption         =   "Observaciones:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   60
            TabIndex        =   12
            Top             =   60
            Width           =   3315
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   1395
         Left            =   30
         TabIndex        =   13
         Top             =   4890
         Width           =   11175
         _Version        =   65536
         _ExtentX        =   19711
         _ExtentY        =   2461
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
         Begin VB.ComboBox cmb_AplViv 
            Height          =   315
            Left            =   1860
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   690
            Width           =   2925
         End
         Begin EditLib.fpDoubleSingle ipp_FoIViv 
            Height          =   315
            Left            =   1860
            TabIndex        =   6
            Top             =   1020
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
            Text            =   "0.000000000"
            DecimalPlaces   =   9
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
         Begin EditLib.fpDateTime ipp_EvaViv 
            Height          =   315
            Left            =   1860
            TabIndex        =   4
            Top             =   360
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
         Begin VB.Label Label15 
            Caption         =   "F. Evaluación:"
            Height          =   285
            Left            =   60
            TabIndex        =   17
            Top             =   360
            Width           =   1485
         End
         Begin VB.Label Label14 
            Caption         =   "Seguro Inmueble"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   60
            TabIndex        =   16
            Top             =   60
            Width           =   3315
         End
         Begin VB.Label Label13 
            Caption         =   "Valor Aplicación:"
            Height          =   315
            Left            =   60
            TabIndex        =   15
            Top             =   1020
            Width           =   1695
         End
         Begin VB.Label Label11 
            Caption         =   "Tipo Aplicación:"
            Height          =   315
            Left            =   60
            TabIndex        =   14
            Top             =   690
            Width           =   1485
         End
      End
      Begin Threed.SSPanel SSPanel10 
         Height          =   1785
         Left            =   30
         TabIndex        =   18
         Top             =   3060
         Width           =   11175
         _Version        =   65536
         _ExtentX        =   19711
         _ExtentY        =   3149
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
         Begin VB.ComboBox cmb_SegDes 
            Height          =   315
            Left            =   1860
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   420
            Width           =   9255
         End
         Begin VB.ComboBox cmb_AplDes 
            Height          =   315
            Left            =   1860
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   1080
            Width           =   2925
         End
         Begin EditLib.fpDateTime ipp_EvaDes 
            Height          =   315
            Left            =   1860
            TabIndex        =   1
            Top             =   750
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
         Begin EditLib.fpDoubleSingle ipp_FoiDes 
            Height          =   315
            Left            =   1860
            TabIndex        =   3
            Top             =   1410
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
            Text            =   "0.000000000"
            DecimalPlaces   =   9
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
         Begin VB.Label Label7 
            Caption         =   "Seguro Desgravamen"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   60
            TabIndex        =   23
            Top             =   60
            Width           =   3315
         End
         Begin VB.Label Label6 
            Caption         =   "Tipo Seguro Desgrav.:"
            Height          =   285
            Left            =   60
            TabIndex        =   22
            Top             =   420
            Width           =   1665
         End
         Begin VB.Label Label12 
            Caption         =   "F. Evaluación:"
            Height          =   285
            Left            =   60
            TabIndex        =   21
            Top             =   750
            Width           =   1485
         End
         Begin VB.Label Label10 
            Caption         =   "Tipo Aplicación:"
            Height          =   315
            Left            =   60
            TabIndex        =   20
            Top             =   1080
            Width           =   1485
         End
         Begin VB.Label Label8 
            Caption         =   "Valor Aplicación:"
            Height          =   315
            Left            =   60
            TabIndex        =   19
            Top             =   1410
            Width           =   1695
         End
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   765
         Left            =   30
         TabIndex        =   24
         Top             =   2250
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
         Begin Threed.SSPanel pnl_EmpSeg 
            Height          =   315
            Left            =   1860
            TabIndex        =   25
            Top             =   60
            Width           =   9255
            _Version        =   65536
            _ExtentX        =   16325
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
         Begin Threed.SSPanel pnl_TipSeg 
            Height          =   315
            Left            =   1860
            TabIndex        =   26
            Top             =   390
            Width           =   9255
            _Version        =   65536
            _ExtentX        =   16325
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
         Begin VB.Label Label5 
            Caption         =   "Empresa Seguros:"
            Height          =   285
            Left            =   60
            TabIndex        =   28
            Top             =   60
            Width           =   1395
         End
         Begin VB.Label Label4 
            Caption         =   "Tipo Seguro Desgrav.:"
            Height          =   285
            Left            =   60
            TabIndex        =   27
            Top             =   390
            Width           =   1665
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   29
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
         Begin Threed.SSPanel SSPanel7 
            Height          =   285
            Left            =   630
            TabIndex        =   38
            Top             =   60
            Width           =   5415
            _Version        =   65536
            _ExtentX        =   9551
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Evaluación de Seguros"
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
         Begin Threed.SSPanel SSPanel5 
            Height          =   285
            Left            =   630
            TabIndex        =   39
            Top             =   330
            Width           =   5415
            _Version        =   65536
            _ExtentX        =   9551
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Registro de Tasas"
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
            Picture         =   "OpeTra_frm_282.frx":0010
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   765
         Left            =   30
         TabIndex        =   30
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
            TabIndex        =   31
            Top             =   60
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
         Begin Threed.SSPanel pnl_FecIng 
            Height          =   315
            Left            =   9690
            TabIndex        =   32
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
            Left            =   1860
            TabIndex        =   33
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
            Caption         =   "Nro. Solicitud"
            Height          =   315
            Left            =   60
            TabIndex        =   36
            Top             =   60
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "F. Ingreso Solicitud:"
            Height          =   315
            Left            =   8040
            TabIndex        =   35
            Top             =   60
            Width           =   1455
         End
         Begin VB.Label Label2 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   60
            TabIndex        =   34
            Top             =   390
            Width           =   1125
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   37
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
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_282.frx":031A
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10560
            Picture         =   "OpeTra_frm_282.frx":075C
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
      End
   End
End
Attribute VB_Name = "frm_Tra_EvaSeg_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_arr_CliNCo()      As modcal_g_est_CuoCli
Dim l_arr_ParPrd()      As moddat_tpo_Genera
Dim l_Arr_TNC_Cli()     As String
Dim l_Arr_TC_Cli()      As String
Dim l_Arr_TNC_Cof()     As String
Dim l_Arr_TC_Cof()      As String
Dim l_str_CodEmp        As String
Dim l_int_TipSeg        As Integer
Dim l_dbl_MtoPre        As Double
Dim l_dbl_ValViv        As Double
Dim l_dbl_MtoTas        As Double
Dim l_dbl_ApoPro        As Double
Dim l_dbl_TasInt        As Double
Dim l_int_PlaAno        As Integer
Dim l_int_PerGra        As Integer
Dim l_int_CuoDbl        As Integer
Dim l_int_DiaPag        As Integer
Dim l_dbl_CuoApr        As Double
Dim l_dbl_CuoAce        As Double
Dim l_dbl_TipCam        As Double
Dim l_str_CodCiu        As String
Dim l_int_TasEsp        As Integer
Dim l_int_TipVal        As Integer
Dim l_dbl_Import        As Double

Private Sub cmd_Grabar_Click()
Dim r_dbl_CuoNue        As Double
Dim r_dbl_Portes        As Double
Dim r_dbl_PorCon        As Double
Dim r_dbl_TopCon        As Double
Dim r_dbl_MtoNCo        As Double
Dim r_dbl_MtoCon        As Double
Dim r_int_FlgExc        As Integer
Dim r_dbl_TasMVi        As Double
Dim r_dbl_ComCof        As Double
Dim r_dbl_TasCof        As Double
Dim r_dat_Fecha         As Date

'variables nueva para la generacion del cronograma
Dim obj_Cronog          As Object
Dim int_Produc          As Integer
Dim int_CuoDbl          As Integer
Dim dbl_ValInm          As Double
Dim dbl_CuoIni          As Double
Dim dbl_MtoCon          As Double
Dim dbl_MtoTas          As Double
Dim int_PlaPre          As Integer
Dim dbl_TasInt          As Double
Dim dbl_TasCof          As Double
Dim dbl_ComCof          As Double
Dim dat_FecDes          As Date
Dim int_DiaVct          As Integer
Dim int_PerGra          As Integer
Dim str_PriVct          As String
Dim dbl_Portes          As Double
Dim dbl_SegViv          As Double
Dim int_TipSDe          As Integer
Dim dbl_SegDes          As Double
Dim dbl_CuoMen          As Double
Dim dbl_CuoPbp          As Double
Dim dbl_IngReq          As Double

   If CDate(ipp_EvaDes.Text) > date Then
      MsgBox "La Fecha de Evaluación no puede ser mayor a la Fecha Actual.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_EvaDes)
      Exit Sub
   End If
   If cmb_AplDes.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Aplicación para el Seguro.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_AplDes)
      Exit Sub
   End If
   If CDbl(ipp_FoiDes.Text) = 0 Then
      MsgBox "Debe ingresar el Valor de Aplicación para el Seguro.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_AplDes)
      Exit Sub
   End If
   If CDate(ipp_EvaViv.Text) > date Then
      MsgBox "La Fecha de Evaluación no puede ser mayor a la Fecha Actual.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_EvaViv)
      Exit Sub
   End If
   If cmb_AplViv.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Aplicación para el Seguro.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_AplViv)
      Exit Sub
   End If
   If CDbl(ipp_FoIViv.Text) = 0 Then
      MsgBox "Debe ingresar el Valor de Aplicación para el Seguro.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_AplViv)
      Exit Sub
   End If
   
   r_dat_Fecha = DateAdd("m", -6, moddat_g_str_FecSis)
   If Format(CDate(r_dat_Fecha), "YYYYMMDD") > Format(CDate(ipp_EvaDes.Text), "YYYYMMDD") Then
      MsgBox "Fecha de Evaluación del seguro de desgravamen debe ser máximo 6 meses mas antigua que la fecha actual.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_EvaDes)
      Exit Sub
   End If
   If Format(CDate(r_dat_Fecha), "YYYYMMDD") > Format(CDate(ipp_EvaViv.Text), "YYYYMMDD") Then
      MsgBox "Fecha de Evaluación del seguro del inmueble debe ser máximo 6 meses mas antigua que la fecha actual.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_EvaDes)
      Exit Sub
   End If
   
   'Recalculando Cuota a Pagar
   r_dbl_CuoNue = 0
   r_dbl_Portes = 0
   r_dbl_TasMVi = 0
   r_dbl_ComCof = 0
   r_dbl_TasCof = 0
   
   'Determina tasa y comision de cofide
   If InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Then
      r_dbl_TasMVi = moddat_gf_ComMVi(moddat_g_str_CodPrd, 3, moddat_g_int_TipMon, l_int_PlaAno)
      r_dbl_ComCof = moddat_gf_ComMVi(moddat_g_str_CodPrd, 4, moddat_g_int_TipMon, l_int_PlaAno)
      r_dbl_TasCof = moddat_gf_ComMVi(moddat_g_str_CodPrd, 5, moddat_g_int_TipMon, l_int_PlaAno)
   ElseIf InStr(moddat_g_str_AgrTFMV, moddat_g_str_CodPrd) > 0 Then
      r_dbl_ComCof = moddat_gf_ComMVi(moddat_g_str_CodPrd, 4, moddat_g_int_TipMon, l_int_PlaAno)
      r_dbl_TasCof = moddat_gf_ComMVi(moddat_g_str_CodPrd, 5, moddat_g_int_TipMon, l_int_PlaAno)
   End If
   
   'Obtiene portes
   If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub, "002", "401") Then
      r_dbl_Portes = moddat_g_arr_Genera(1).Genera_Cantid
   End If
   
   Select Case moddat_g_str_CodPrd > 0
      'Case InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd)
      '   'Para obtener porcentaje de TC
      '   r_dbl_PorCon = 0
      '   If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "011") Then
      '      r_dbl_PorCon = moddat_g_arr_Genera(1).Genera_Cantid
      '   End If
      '
      '   'Para obtener tope de TC
      '   r_dbl_TopCon = 0
      '   If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "012") Then
      '      r_dbl_TopCon = moddat_g_arr_Genera(1).Genera_Cantid
      '   End If
      '
      '   'NUEVA rutina de generacion de cronogramas
      '   int_Produc = 1
      '   int_CuoDbl = l_int_CuoDbl
      '   dbl_ValInm = l_dbl_ValViv
      '   dbl_CuoIni = l_dbl_ApoPro
      '   dbl_MtoCon = (l_dbl_ValViv - l_dbl_ApoPro) * (r_dbl_PorCon / 100)
      '   If dbl_MtoCon > r_dbl_TopCon Then dbl_MtoCon = r_dbl_TopCon
      '   dbl_MtoTas = l_dbl_MtoTas
      '   int_PlaPre = l_int_PlaAno * 12
      '   dbl_TasInt = l_dbl_TasInt
      '   dbl_TasCof = r_dbl_TasCof
      '   dbl_ComCof = r_dbl_ComCof
      '   dat_FecDes = CDate(Format(date, "dd/mm/yyyy"))
      '   int_DiaVct = l_int_DiaPag
      '   int_PerGra = l_int_PerGra
      '   str_PriVct = ""
      '   dbl_Portes = r_dbl_Portes
      '   dbl_SegViv = CDbl(ipp_FoIViv.Text)
      '   int_TipSDe = l_int_TipSeg - 10
      '   dbl_SegDes = CDbl(ipp_FoiDes.Text)
      '
      '   'Calculando cronogramas
      '   Set obj_Cronog = CreateObject("ComCronograma.ClsCronograma")
      '   Call obj_Cronog.Listar(l_Arr_TNC_Cli(), l_Arr_TC_Cli(), l_Arr_TNC_Cof(), l_Arr_TC_Cof(), int_Produc, int_CuoDbl, dbl_ValInm, dbl_CuoIni, dbl_MtoTas, dbl_MtoCon, int_PlaPre, dbl_TasInt, dbl_TasCof, dbl_ComCof, dat_FecDes, 0, int_DiaVct, str_PriVct, int_PerGra, dbl_Portes, dbl_SegViv, int_TipSDe, dbl_SegDes)
      '
      '   dbl_CuoMen = 0
      '   dbl_CuoPbp = 0
      '   dbl_IngReq = 0
      '   Call modgen_gf_Buscar_CuotaMensual(l_Arr_TNC_Cli(), l_Arr_TC_Cli(), int_CuoDbl, int_PerGra, int_Produc, dbl_CuoMen, dbl_CuoPbp, dbl_IngReq, moddat_g_str_CodPrd, moddat_g_str_CodSub)
      '
      '   'muestra valor cuota
      '   r_dbl_CuoNue = Format(dbl_CuoPbp, "###,###,##0.00") & " "
   
      Case InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd)
         'NUEVA rutina de generacion de cronogramas
         int_Produc = 2
         int_CuoDbl = l_int_CuoDbl
         dbl_ValInm = l_dbl_ValViv
         dbl_CuoIni = l_dbl_ApoPro
         dbl_MtoCon = 0
         dbl_MtoTas = l_dbl_MtoTas
         int_PlaPre = l_int_PlaAno * 12
         dbl_TasInt = l_dbl_TasInt
         dbl_TasCof = 0
         dbl_ComCof = 0
         dat_FecDes = CDate(Format(date, "dd/mm/yyyy"))
         int_DiaVct = l_int_DiaPag
         int_PerGra = l_int_PerGra
         str_PriVct = ""
         dbl_Portes = r_dbl_Portes
         dbl_SegViv = CDbl(ipp_FoIViv.Text)
         int_TipSDe = l_int_TipSeg - 10
         dbl_SegDes = CDbl(ipp_FoiDes.Text)
         
         'Calculando cronogramas
         Set obj_Cronog = CreateObject("ComCronograma.ClsCronograma")
         Call obj_Cronog.Listar(l_Arr_TNC_Cli(), l_Arr_TC_Cli(), l_Arr_TNC_Cof(), l_Arr_TC_Cof(), int_Produc, int_CuoDbl, dbl_ValInm, dbl_CuoIni, dbl_MtoCon, dbl_MtoTas, int_PlaPre, dbl_TasInt, dbl_TasCof, dbl_ComCof, dat_FecDes, 0, int_DiaVct, str_PriVct, int_PerGra, dbl_Portes, dbl_SegViv, int_TipSDe, dbl_SegDes)
         
         dbl_CuoMen = 0
         dbl_CuoPbp = 0
         dbl_IngReq = 0
         Call modgen_gf_Buscar_CuotaMensual(l_Arr_TNC_Cli(), l_Arr_TC_Cli(), int_CuoDbl, int_PerGra, int_Produc, dbl_CuoMen, dbl_CuoPbp, dbl_IngReq, moddat_g_str_CodPrd, moddat_g_str_CodSub)
         
         'muestra valor cuota
         If moddat_g_int_TipMon = 1 Then
            r_dbl_CuoNue = Format(dbl_CuoMen, "###,###,##0.00") & " "
         Else
            r_dbl_CuoNue = Format(dbl_CuoMen * l_dbl_TipCam, "###,###,##0.00") & " "
         End If
         
      Case InStr(moddat_g_str_AgrMIHG, moddat_g_str_CodPrd) Or InStr(moddat_g_str_Agr2MIC, moddat_g_str_CodPrd) Or InStr(moddat_g_str_Agr2FMV, moddat_g_str_CodPrd)
         'Para obtener Tope concesional
         r_dbl_TopCon = 0
         If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub, "051", "012") Then
            r_dbl_TopCon = moddat_g_arr_Genera(1).Genera_Cantid
         End If
         If CDbl(l_dbl_ValViv) > (50 * moddat_gf_Consulta_ParVal("001", "002")) Then
            r_dbl_TopCon = 5000
         End If
         
         'NUEVA rutina de generacion de cronogramas
         int_Produc = 1
         int_CuoDbl = l_int_CuoDbl
         dbl_ValInm = l_dbl_ValViv
         dbl_CuoIni = l_dbl_ApoPro
         dbl_MtoCon = r_dbl_TopCon
         dbl_MtoTas = l_dbl_MtoTas
         int_PlaPre = l_int_PlaAno * 12
         dbl_TasInt = l_dbl_TasInt
         dbl_TasCof = r_dbl_TasCof
         dbl_ComCof = r_dbl_ComCof
         dat_FecDes = CDate(Format(date, "dd/mm/yyyy"))
         int_DiaVct = l_int_DiaPag
         int_PerGra = l_int_PerGra
         str_PriVct = ""
         dbl_Portes = r_dbl_Portes
         dbl_SegViv = CDbl(ipp_FoIViv.Text)
         int_TipSDe = l_int_TipSeg - 10
         dbl_SegDes = CDbl(ipp_FoiDes.Text)
         
         'Calculando cronogramas
         Set obj_Cronog = CreateObject("ComCronograma.ClsCronograma")
         Call obj_Cronog.Listar(l_Arr_TNC_Cli(), l_Arr_TC_Cli(), l_Arr_TNC_Cof(), l_Arr_TC_Cof(), int_Produc, int_CuoDbl, dbl_ValInm, dbl_CuoIni, dbl_MtoCon, dbl_MtoTas, int_PlaPre, dbl_TasInt, dbl_TasCof, dbl_ComCof, dat_FecDes, 0, int_DiaVct, str_PriVct, int_PerGra, dbl_Portes, dbl_SegViv, int_TipSDe, dbl_SegDes)
         
         dbl_CuoMen = 0
         dbl_CuoPbp = 0
         dbl_IngReq = 0
         Call modgen_gf_Buscar_CuotaMensual(l_Arr_TNC_Cli(), l_Arr_TC_Cli(), int_CuoDbl, int_PerGra, int_Produc, dbl_CuoMen, dbl_CuoPbp, dbl_IngReq, moddat_g_str_CodPrd, moddat_g_str_CodSub)
         
         'muestra valor cuota
         r_dbl_CuoNue = Format(dbl_CuoPbp, "###,###,##0.00") & " "
         
      Case InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd)
         'NUEVA rutina de generacion de cronogramas
         int_Produc = 3
         int_CuoDbl = l_int_CuoDbl
         dbl_ValInm = l_dbl_ValViv
         dbl_CuoIni = l_dbl_ApoPro
         dbl_MtoCon = 0
         dbl_MtoTas = l_dbl_MtoTas
         int_PlaPre = l_int_PlaAno * 12
         dbl_TasInt = l_dbl_TasInt
         dbl_TasCof = r_dbl_TasCof
         dbl_ComCof = r_dbl_ComCof
         dat_FecDes = CDate(Format(date, "dd/mm/yyyy"))
         int_DiaVct = l_int_DiaPag
         int_PerGra = l_int_PerGra
         str_PriVct = ""
         dbl_Portes = r_dbl_Portes
         dbl_SegViv = CDbl(ipp_FoIViv.Text)
         int_TipSDe = l_int_TipSeg - 10
         dbl_SegDes = CDbl(ipp_FoiDes.Text)
         
         'Calculando cronogramas
         Set obj_Cronog = CreateObject("ComCronograma.ClsCronograma")
         Call obj_Cronog.Listar(l_Arr_TNC_Cli(), l_Arr_TC_Cli(), l_Arr_TNC_Cof(), l_Arr_TC_Cof(), int_Produc, int_CuoDbl, dbl_ValInm, dbl_CuoIni, dbl_MtoCon, dbl_MtoTas, int_PlaPre, dbl_TasInt, dbl_TasCof, dbl_ComCof, dat_FecDes, 0, int_DiaVct, str_PriVct, int_PerGra, dbl_Portes, dbl_SegViv, int_TipSDe, dbl_SegDes)
         
         dbl_CuoMen = 0
         dbl_CuoPbp = 0
         dbl_IngReq = 0
         Call modgen_gf_Buscar_CuotaMensual(l_Arr_TNC_Cli(), l_Arr_TC_Cli(), int_CuoDbl, int_PerGra, int_Produc, dbl_CuoMen, dbl_CuoPbp, dbl_IngReq, moddat_g_str_CodPrd, moddat_g_str_CodSub)
         
         'muestra valor cuota
         r_dbl_CuoNue = Format(dbl_CuoMen, "###,###,##0.00") & " "
         
   End Select
   
   r_int_FlgExc = 1
   If r_dbl_CuoNue > l_dbl_CuoApr Then
      If r_dbl_CuoNue > l_dbl_CuoAce And modgen_g_int_TipUsu <> 18200 And modgen_g_int_TipUsu <> 18000 Then
         If MsgBox(moddat_g_str_Msje01 & vbCrLf & "Cuota Obtenida: (" & Format(r_dbl_CuoNue, "##0.00") & ",  Cuota Aprobada: (" & Format(l_dbl_CuoApr, "##0.00") & "). ¿Desea aprobar esta excepción?", vbQuestion + vbDefaultButton2 + vbYesNo, modgen_g_str_NomPlt) <> vbYes Then
            Exit Sub
         End If
         r_int_FlgExc = 2
      Else
         If MsgBox(moddat_g_str_Msje01 & vbCrLf & "Cuota Obtenida: (" & Format(r_dbl_CuoNue, "##0.00") & ",  Cuota Aprobada: (" & Format(l_dbl_CuoApr, "##0.00") & "). ¿Desea aprobar esta excepción?", vbQuestion + vbDefaultButton2 + vbYesNo, modgen_g_str_NomPlt) <> vbYes Then
            Exit Sub
         End If
         r_int_FlgExc = 2
      End If
   End If
   
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   'Registrando Excepción
   If r_int_FlgExc = 2 Then
      Call fs_RegExc
   End If
   
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      Screen.MousePointer = 11
      Call moddat_gs_FecSis
      
      g_str_Parame = "USP_TRA_EVASEG ("
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumSol & "', "
      g_str_Parame = g_str_Parame & Format(CDate(ipp_EvaDes.Text), "yyyymmdd") & ", "
      g_str_Parame = g_str_Parame & Format(CDate(ipp_EvaViv.Text), "yyyymmdd") & ", "
      g_str_Parame = g_str_Parame & "'" & l_str_CodEmp & "', "
      g_str_Parame = g_str_Parame & CStr(cmb_SegDes.ItemData(cmb_SegDes.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & "'" & l_str_CodEmp & "', "
      g_str_Parame = g_str_Parame & CStr(cmb_AplDes.ItemData(cmb_AplDes.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_FoiDes.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(cmb_AplViv.ItemData(cmb_AplViv.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_FoIViv.Value) & ", "
      g_str_Parame = g_str_Parame & "'" & txt_ObsEva.Text & "', "
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
   
   'Creando Nueva Ocurrencia en Detalle de Seguimiento
   If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, 42, 35, 0, "", 0, 0) Then
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
   
   Call gs_SetFocus(cmb_SegDes)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
Dim r_int_MonVta     As Integer
Dim l_int_DatCyg     As Integer
Dim r_int_TipVal     As Integer
Dim r_dbl_Import     As Double

   Call moddat_gs_Carga_LisIte_Combo(cmb_AplDes, 1, "227")
   Call moddat_gs_Carga_LisIte_Combo(cmb_AplViv, 1, "227")
   pnl_EmpSeg.Caption = ""
   pnl_TipSeg.Caption = ""
   ipp_EvaDes.Text = Format(date, "dd/mm/yyyy")
   cmb_AplDes.ListIndex = -1
   ipp_FoiDes.Value = 0
   ipp_EvaViv.Text = Format(date, "dd/mm/yyyy")
   cmb_AplViv.ListIndex = -1
   ipp_FoIViv.Value = 0
   txt_ObsEva.Text = ""
   
   'Obteniendo Empresa de Seguros y Tipo de Seguro
   moddat_g_int_CygTDo = 0
   moddat_g_str_CygNDo = ""
   l_str_CodEmp = ""
   l_int_TipSeg = 0
   l_dbl_MtoPre = 0
   l_dbl_ValViv = 0
   r_int_MonVta = 0
   l_dbl_TipCam = 0
   l_dbl_MtoTas = 0
   l_int_DatCyg = 0
   
   'Obtiene datos de la solicitud
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM CRE_SOLMAE "
   g_str_Parame = g_str_Parame & " WHERE SOLMAE_NUMERO = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      l_str_CodEmp = Trim(g_rst_Princi!SOLMAE_ESGDES & "")
      l_int_TipSeg = g_rst_Princi!SOLMAE_TIPSEG
      l_dbl_MtoPre = g_rst_Princi!SOLMAE_MTOPRE_MPR
      l_int_PlaAno = g_rst_Princi!SOLMAE_PLAANO
      l_dbl_TasInt = g_rst_Princi!SOLMAE_TASINT
      l_int_DiaPag = g_rst_Princi!SOLMAE_DIAPAG
      l_int_PerGra = g_rst_Princi!SOLMAE_PERGRA
      l_int_CuoDbl = g_rst_Princi!SOLMAE_CUOEXT
      l_dbl_CuoApr = g_rst_Princi!SOLMAE_CUOMEN_MPR
      l_dbl_CuoAce = g_rst_Princi!SOLMAE_CUOAPR_MPR
      r_int_MonVta = g_rst_Princi!SOLMAE_TIPMON
      l_int_TasEsp = g_rst_Princi!SOLMAE_TASESP
      
      If g_rst_Princi!SOLMAE_TIPMON = 1 Then
         l_dbl_ValViv = g_rst_Princi!SOLMAE_COMVTA_SOL
         l_dbl_ApoPro = g_rst_Princi!SOLMAE_APOPRO_SOL - g_rst_Princi!SOLMAE_MTOGCI
      Else
         l_dbl_ValViv = g_rst_Princi!SOLMAE_COMVTA_DOL
         l_dbl_ApoPro = g_rst_Princi!SOLMAE_APOPRO_DOL - g_rst_Princi!SOLMAE_MTOGCI
      End If
      
      pnl_EmpSeg.Caption = moddat_gf_Consulta_ComSeg(g_rst_Princi!SOLMAE_ESGDES & "")
      pnl_TipSeg.Caption = moddat_gf_Consulta_TipSeg(g_rst_Princi!SOLMAE_ESGDES, g_rst_Princi!SOLMAE_TIPSEG)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Obteniendo Tipo de Cambio de Moneda del Préstamo
   If moddat_g_int_TipMon <> 1 Then
      l_dbl_TipCam = moddat_gf_Obtiene_TipCam(1, moddat_g_int_TipMon)
   End If
   
   'Valor de Vivienda
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM TRA_EVATAS "
   g_str_Parame = g_str_Parame & " WHERE EVATAS_NUMSOL = '" & moddat_g_str_NumSol & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      r_int_MonVta = g_rst_Princi!EVATAS_TIPMON
      l_dbl_MtoTas = g_rst_Princi!EVATAS_SUMASE_INM + g_rst_Princi!EVATAS_SUMASE_ES1 + g_rst_Princi!EVATAS_SUMASE_ES2 + g_rst_Princi!EVATAS_SUMASE_DEP
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call moddat_gs_Carga_TipSeg(cmb_SegDes, l_str_CodEmp)
   Call gs_BuscarCombo_Item(cmb_SegDes, l_int_TipSeg)
   
   'Datos del Cliente
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CLI_DATGEN "
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
   
   'Obteniendo Factores y Valores
   ipp_FoiDes.Value = moddat_gf_Consulta_AplSeg(moddat_g_str_CodPrd, moddat_g_str_CodSub, l_str_CodEmp, l_int_TipSeg, moddat_g_int_TipMon, l_dbl_MtoPre, cmb_AplDes, l_int_TasEsp)
   ipp_FoIViv.Value = moddat_gf_Consulta_AplSeg(moddat_g_str_CodPrd, moddat_g_str_CodSub, l_str_CodEmp, 0, r_int_MonVta, l_dbl_ValViv, cmb_AplViv, l_int_TasEsp)
   
   Call moddat_gf_Consulta_ValSeg(moddat_g_str_CodPrd, moddat_g_str_CodSub, l_str_CodEmp, l_int_TipSeg, moddat_g_int_TipMon, l_dbl_MtoPre, l_int_TipVal, l_dbl_Import, l_int_TasEsp)
   ipp_FoiDes.Value = l_dbl_Import
   Call moddat_gf_Consulta_ValSeg(moddat_g_str_CodPrd, moddat_g_str_CodSub, l_str_CodEmp, 0, moddat_g_int_TipMon, l_dbl_ValViv, l_int_TipVal, l_dbl_Import, l_int_TasEsp)
   ipp_FoIViv.Value = l_dbl_Import
End Sub

Private Sub fs_Buscar_DatEva()
   moddat_g_int_FlgGrb = 1
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM TRA_EVASEG "
   g_str_Parame = g_str_Parame & " WHERE EVASEG_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      moddat_g_int_FlgGrb = 2
      g_rst_Princi.MoveFirst
      ipp_EvaDes.Text = gf_FormatoFecha(g_rst_Princi!EVASEG_EVADES)
      Call gs_BuscarCombo_Item(cmb_AplDes, g_rst_Princi!EVASEG_TIPDES)
      ipp_FoiDes.Value = g_rst_Princi!EVASEG_FOIDES
      ipp_EvaViv.Text = gf_FormatoFecha(g_rst_Princi!EVASEG_EVAVIV)
      Call gs_BuscarCombo_Item(cmb_AplDes, g_rst_Princi!EVASEG_TIPVIV)
      ipp_FoIViv.Value = g_rst_Princi!EVASEG_FOIVIV
      txt_ObsEva.Text = Trim(g_rst_Princi!EVASEG_OBSERV & "")
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_RegExc()
Dim r_int_NumExc     As Integer
Dim r_int_NivAut     As Integer

   If modgen_g_int_TipUsu = 18200 Then
      r_int_NivAut = 31
   Else
      r_int_NivAut = 13
   End If

   'Generando Número de Excepción
   r_int_NumExc = 0
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT COUNT(SEGEXC_NUMSOL) AS NUMREG "
   g_str_Parame = g_str_Parame & "  FROM TRA_SEGEXC "
   g_str_Parame = g_str_Parame & " WHERE SEGEXC_NUMSOL = '" & moddat_g_str_NumSol & "' "
   
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
   If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, 42, 18, 0, "", 0, 0) Then
      Exit Sub
   End If
   
   'Grabando en Detalle de Excepciones
   If Not moddat_gf_Inserta_SegExc(moddat_g_str_NumSol, 42, r_int_NumExc, moddat_g_str_Msje02, r_int_NivAut) Then
      Exit Sub
   End If
End Sub

Private Sub cmb_AplDes_Click()
   Call gs_SetFocus(ipp_FoiDes)
End Sub

Private Sub cmb_AplDes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_AplDes_Click
   End If
End Sub

Private Sub cmb_AplViv_Click()
   Call gs_SetFocus(ipp_FoIViv)
End Sub

Private Sub cmb_AplViv_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_AplViv_Click
   End If
End Sub

Private Sub cmb_SegDes_Click()
   If cmb_SegDes.ListIndex > -1 Then
      Screen.MousePointer = 11
      Call moddat_gf_Consulta_ValSeg(moddat_g_str_CodPrd, moddat_g_str_CodSub, l_str_CodEmp, cmb_SegDes.ItemData(cmb_SegDes.ListIndex), moddat_g_int_TipMon, l_dbl_MtoPre, l_int_TipVal, l_dbl_Import, l_int_TasEsp)
      ipp_FoiDes.Value = l_dbl_Import
      'ipp_FoiDes.Value = moddat_gf_Consulta_AplSeg(moddat_g_str_CodPrd, moddat_g_str_CodSub, l_str_CodEmp, cmb_SegDes.ItemData(cmb_SegDes.ListIndex), moddat_g_int_TipMon, l_dbl_MtoPre, cmb_AplDes, l_str_CodCiu)
      Screen.MousePointer = 0
      Call gs_SetFocus(ipp_EvaDes)
   End If
End Sub

Private Sub cmb_SegDes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_SegDes_Click
   End If
End Sub

Private Sub ipp_EvaDes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_AplDes)
   End If
End Sub

Private Sub ipp_EvaViv_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_AplViv)
   End If
End Sub

Private Sub ipp_FoiDes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_EvaViv)
   End If
End Sub

Private Sub ipp_FoIViv_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_ObsEva)
   End If
End Sub

Private Sub txt_ObsEva_GotFocus()
   Call gs_SelecTodo(txt_ObsEva)
End Sub

Private Sub txt_ObsEva_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_., ;:()/&%$·!ª@#=?¿+*" & Chr(10))
   End If
End Sub
