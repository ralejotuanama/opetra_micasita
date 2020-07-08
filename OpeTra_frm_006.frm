VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frm_Cob_CuoHip_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   9495
   ClientLeft      =   1020
   ClientTop       =   825
   ClientWidth     =   12870
   Icon            =   "OpeTra_frm_006.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9495
   ScaleWidth      =   12870
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel2 
      Height          =   9495
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   12855
      _Version        =   65536
      _ExtentX        =   22675
      _ExtentY        =   16748
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
      Begin Threed.SSPanel SSPanel19 
         Height          =   1545
         Left            =   30
         TabIndex        =   58
         Top             =   4050
         Width           =   12765
         _Version        =   65536
         _ExtentX        =   22516
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
         Begin MSFlexGridLib.MSFlexGrid grd_CuoVig 
            Height          =   855
            Left            =   60
            TabIndex        =   59
            Top             =   630
            Width           =   12645
            _ExtentX        =   22304
            _ExtentY        =   1508
            _Version        =   393216
            Rows            =   21
            Cols            =   8
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   49152
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel21 
            Height          =   285
            Left            =   90
            TabIndex        =   60
            Top             =   360
            Width           =   1245
            _Version        =   65536
            _ExtentX        =   2196
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
         Begin Threed.SSPanel SSPanel22 
            Height          =   285
            Left            =   2880
            TabIndex        =   61
            Top             =   360
            Width           =   1605
            _Version        =   65536
            _ExtentX        =   2831
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
         Begin Threed.SSPanel SSPanel24 
            Height          =   285
            Left            =   4470
            TabIndex        =   62
            Top             =   360
            Width           =   1605
            _Version        =   65536
            _ExtentX        =   2831
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
         Begin Threed.SSPanel SSPanel25 
            Height          =   285
            Left            =   6060
            TabIndex        =   63
            Top             =   360
            Width           =   1605
            _Version        =   65536
            _ExtentX        =   2831
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Seg. Desg."
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
            Left            =   7650
            TabIndex        =   64
            Top             =   360
            Width           =   1605
            _Version        =   65536
            _ExtentX        =   2831
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Seg. Viv."
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
            Left            =   9240
            TabIndex        =   65
            Top             =   360
            Width           =   1605
            _Version        =   65536
            _ExtentX        =   2831
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Otros Cargos"
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
            Left            =   10830
            TabIndex        =   66
            Top             =   360
            Width           =   1605
            _Version        =   65536
            _ExtentX        =   2831
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
         Begin Threed.SSPanel SSPanel28 
            Height          =   285
            Left            =   1320
            TabIndex        =   69
            Top             =   360
            Width           =   1605
            _Version        =   65536
            _ExtentX        =   2831
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Vcto."
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
         Begin VB.Label Label8 
            Caption         =   "Lista de Cuotas x Vencer"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   90
            TabIndex        =   68
            Top             =   30
            Width           =   2925
         End
      End
      Begin Threed.SSPanel SSPanel23 
         Height          =   1185
         Left            =   30
         TabIndex        =   49
         Top             =   7440
         Width           =   12765
         _Version        =   65536
         _ExtentX        =   22516
         _ExtentY        =   2090
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
         Begin Threed.SSPanel pnl_ITFPag 
            Height          =   315
            Left            =   1620
            TabIndex        =   50
            Top             =   60
            Width           =   1635
            _Version        =   65536
            _ExtentX        =   2884
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
         Begin Threed.SSPanel pnl_NetPag 
            Height          =   315
            Left            =   1620
            TabIndex        =   52
            Top             =   390
            Width           =   1635
            _Version        =   65536
            _ExtentX        =   2884
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
         Begin Threed.SSPanel pnl_NetMPr 
            Height          =   315
            Left            =   1620
            TabIndex        =   54
            Top             =   720
            Width           =   1635
            _Version        =   65536
            _ExtentX        =   2884
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
         Begin Threed.SSPanel pnl_TCaMPr 
            Height          =   315
            Left            =   11430
            TabIndex        =   56
            Top             =   60
            Width           =   1245
            _Version        =   65536
            _ExtentX        =   2196
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00000 "
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
         Begin VB.Label Label13 
            Caption         =   "Tipo de Cambio:"
            Height          =   285
            Left            =   10080
            TabIndex        =   57
            Top             =   60
            Width           =   1275
         End
         Begin VB.Label Label20 
            Caption         =   "Imp. Neto Pagado (M. Prest.):"
            Height          =   405
            Left            =   90
            TabIndex        =   55
            Top             =   720
            Width           =   1395
         End
         Begin VB.Label Label16 
            Caption         =   "Imp. Neto Pagado:"
            Height          =   285
            Left            =   90
            TabIndex        =   53
            Top             =   390
            Width           =   1395
         End
         Begin VB.Label Label15 
            Caption         =   "ITF:"
            Height          =   285
            Left            =   90
            TabIndex        =   51
            Top             =   60
            Width           =   975
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   1755
         Left            =   30
         TabIndex        =   15
         Top             =   5640
         Width           =   12765
         _Version        =   65536
         _ExtentX        =   22516
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
         Begin VB.ComboBox cmb_CtaBan 
            Height          =   315
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   390
            Width           =   2775
         End
         Begin VB.TextBox txt_NumCom 
            Height          =   315
            Left            =   1620
            MaxLength       =   25
            TabIndex        =   11
            Text            =   "Text1"
            Top             =   1050
            Width           =   2775
         End
         Begin VB.ComboBox cmb_CodBan 
            Height          =   315
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   60
            Width           =   2775
         End
         Begin EditLib.fpDoubleSingle ipp_Import 
            Height          =   315
            Left            =   1620
            TabIndex        =   12
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
         Begin EditLib.fpDateTime ipp_FecPag 
            Height          =   315
            Left            =   1620
            TabIndex        =   10
            Top             =   720
            Width           =   1305
            _Version        =   196608
            _ExtentX        =   2302
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
         Begin VB.Label Label12 
            Caption         =   "Nro. Comprobante:"
            Height          =   285
            Left            =   90
            TabIndex        =   48
            Top             =   1050
            Width           =   1335
         End
         Begin VB.Label Label41 
            Caption         =   "Importe Pagado:"
            Height          =   285
            Left            =   90
            TabIndex        =   19
            Top             =   1380
            Width           =   1335
         End
         Begin VB.Label Label6 
            Caption         =   "Banco:"
            Height          =   315
            Left            =   90
            TabIndex        =   18
            Top             =   60
            Width           =   1455
         End
         Begin VB.Label Label7 
            Caption         =   "Nro. Cuenta:"
            Height          =   285
            Left            =   90
            TabIndex        =   17
            Top             =   390
            Width           =   1335
         End
         Begin VB.Label Label10 
            Caption         =   "Fecha de Pago:"
            Height          =   315
            Left            =   90
            TabIndex        =   16
            Top             =   720
            Width           =   1245
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   20
         Top             =   30
         Width           =   12765
         _Version        =   65536
         _ExtentX        =   22516
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
            Height          =   585
            Left            =   660
            TabIndex        =   21
            Top             =   30
            Width           =   3975
            _Version        =   65536
            _ExtentX        =   7011
            _ExtentY        =   1032
            _StockProps     =   15
            Caption         =   "Cobro de Cuotas"
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
            Picture         =   "OpeTra_frm_006.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel20 
         Height          =   795
         Left            =   30
         TabIndex        =   22
         Top             =   750
         Width           =   12765
         _Version        =   65536
         _ExtentX        =   22516
         _ExtentY        =   1402
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
         Begin VB.ComboBox cmb_TipDoc 
            Height          =   315
            Left            =   6210
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   60
            Width           =   2775
         End
         Begin VB.TextBox txt_NumDoc 
            Height          =   315
            Left            =   6210
            MaxLength       =   12
            TabIndex        =   3
            Text            =   "Text1"
            Top             =   390
            Width           =   2775
         End
         Begin VB.ComboBox cmb_TipBus 
            Height          =   315
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   2775
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   12030
            Picture         =   "OpeTra_frm_006.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Salir"
            Top             =   60
            Width           =   675
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   675
            Left            =   11310
            Picture         =   "OpeTra_frm_006.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Limpiar Datos"
            Top             =   60
            Width           =   675
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   675
            Left            =   10590
            Picture         =   "OpeTra_frm_006.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Buscar Registros"
            Top             =   60
            Width           =   675
         End
         Begin MSMask.MaskEdBox msk_NumOpe 
            Height          =   315
            Left            =   1620
            TabIndex        =   1
            Top             =   390
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   12
            Mask            =   "###-##-#####"
            PromptChar      =   " "
         End
         Begin VB.Label lbl_Numero 
            Caption         =   "Nro. Operación:"
            Height          =   285
            Left            =   90
            TabIndex        =   27
            Top             =   390
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo Doc. Ident.:"
            Height          =   315
            Left            =   4830
            TabIndex        =   26
            Top             =   60
            Width           =   1395
         End
         Begin VB.Label Label19 
            Caption         =   "Nro. Doc. Ident.:"
            Height          =   285
            Left            =   4830
            TabIndex        =   25
            Top             =   390
            Width           =   1335
         End
         Begin VB.Label Label18 
            Caption         =   "Tipo de Búsqueda:"
            Height          =   315
            Left            =   90
            TabIndex        =   24
            Top             =   60
            Width           =   1455
         End
         Begin VB.Label Label17 
            Caption         =   "Nro. Doc. Id.:"
            Height          =   285
            Left            =   60
            TabIndex        =   23
            Top             =   1740
            Width           =   1065
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   795
         Left            =   30
         TabIndex        =   28
         Top             =   1590
         Width           =   12765
         _Version        =   65536
         _ExtentX        =   22516
         _ExtentY        =   1402
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
            Left            =   1620
            TabIndex        =   29
            Top             =   60
            Width           =   1905
            _Version        =   65536
            _ExtentX        =   3360
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
         Begin Threed.SSPanel pnl_NomCli 
            Height          =   315
            Left            =   1620
            TabIndex        =   30
            Top             =   390
            Width           =   11055
            _Version        =   65536
            _ExtentX        =   19500
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
         Begin Threed.SSPanel pnl_MonPre 
            Height          =   315
            Left            =   8190
            TabIndex        =   46
            Top             =   60
            Width           =   4485
            _Version        =   65536
            _ExtentX        =   7911
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
         Begin VB.Label Label11 
            Caption         =   "Moneda Préstamo:"
            Height          =   315
            Left            =   6750
            TabIndex        =   47
            Top             =   60
            Width           =   1395
         End
         Begin VB.Label Label2 
            Caption         =   "Nro. de Operación:"
            Height          =   315
            Left            =   90
            TabIndex        =   32
            Top             =   60
            Width           =   1395
         End
         Begin VB.Label Label3 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   90
            TabIndex        =   31
            Top             =   390
            Width           =   1125
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   1575
         Left            =   30
         TabIndex        =   33
         Top             =   2430
         Width           =   12765
         _Version        =   65536
         _ExtentX        =   22516
         _ExtentY        =   2778
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
            Height          =   855
            Left            =   60
            TabIndex        =   7
            Top             =   660
            Width           =   12645
            _ExtentX        =   22304
            _ExtentY        =   1508
            _Version        =   393216
            Rows            =   21
            Cols            =   12
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   49152
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel11 
            Height          =   285
            Left            =   90
            TabIndex        =   34
            Top             =   390
            Width           =   585
            _Version        =   65536
            _ExtentX        =   1032
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
         Begin Threed.SSPanel SSPanel4 
            Height          =   285
            Left            =   1620
            TabIndex        =   36
            Top             =   390
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
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
         Begin Threed.SSPanel SSPanel8 
            Height          =   285
            Left            =   2700
            TabIndex        =   37
            Top             =   390
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
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
         Begin Threed.SSPanel SSPanel9 
            Height          =   285
            Left            =   3780
            TabIndex        =   38
            Top             =   390
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Seg. Desg."
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
         Begin Threed.SSPanel SSPanel14 
            Height          =   285
            Left            =   4860
            TabIndex        =   39
            Top             =   390
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Seg. Viv."
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
         Begin Threed.SSPanel SSPanel15 
            Height          =   285
            Left            =   5940
            TabIndex        =   40
            Top             =   390
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Otr. Cargos"
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
         Begin Threed.SSPanel SSPanel10 
            Height          =   285
            Left            =   7020
            TabIndex        =   41
            Top             =   390
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Int. Morat."
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
            Left            =   8100
            TabIndex        =   42
            Top             =   390
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Int. Comp."
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
            Left            =   9180
            TabIndex        =   43
            Top             =   390
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Gastos Cob."
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
            Left            =   11340
            TabIndex        =   44
            Top             =   390
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
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
         Begin Threed.SSPanel SSPanel18 
            Height          =   285
            Left            =   660
            TabIndex        =   45
            Top             =   390
            Width           =   975
            _Version        =   65536
            _ExtentX        =   1720
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Vcto."
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
            Left            =   10260
            TabIndex        =   70
            Top             =   390
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Otros Gastos"
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
         Begin VB.Label Label5 
            Caption         =   "Lista de Cuotas Vencidas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   90
            TabIndex        =   67
            Top             =   60
            Width           =   2925
         End
      End
      Begin Threed.SSPanel SSPanel12 
         Height          =   765
         Left            =   30
         TabIndex        =   35
         Top             =   8670
         Width           =   12765
         _Version        =   65536
         _ExtentX        =   22516
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
         Begin VB.CommandButton cmd_Grabar 
            Height          =   675
            Left            =   12060
            Picture         =   "OpeTra_frm_006.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   675
         End
      End
   End
End
Attribute VB_Name = "frm_Cob_CuoHip_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_CtaBan()   As moddat_tpo_Genera
Dim l_arr_CodBan()   As moddat_tpo_Genera
Dim l_dbl_PorITF     As Double
Dim l_int_MonDes     As Integer
Dim l_int_MonViv     As Integer
Dim l_int_MonOtr     As Integer
Dim l_dbl_TCaDol     As Double
Dim l_dbl_TCaMPr     As Double
Dim l_dbl_TCaDes     As Double
Dim l_dbl_TCaViv     As Double
Dim l_dbl_TCaOtr     As Double
Dim l_int_SitCre     As Integer
Dim l_int_SitAnt     As Integer
Dim l_int_Situac     As Integer
Dim l_int_CuoPen     As Integer

Private Sub cmb_CodBan_Click()
   If cmb_CodBan.ListIndex > -1 Then
      Screen.MousePointer = 11
      Call moddat_gs_Carga_CtaBan(l_arr_CodBan(cmb_CodBan.ListIndex + 1).Genera_Codigo, cmb_CtaBan, l_arr_CtaBan)
      Screen.MousePointer = 0
         
      Call gs_SetFocus(cmb_CtaBan)
   Else
      cmb_CtaBan.Clear
   End If
End Sub

Private Sub cmb_CodBan_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_CodBan_Click
   End If
End Sub

Private Sub cmb_CtaBan_Click()
   Call gs_SetFocus(ipp_FecPag)
End Sub

Private Sub cmb_CtaBan_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_CtaBan_Click
   End If
End Sub

Private Sub cmb_TipBus_Click()
   If cmb_TipBus.ListIndex > -1 Then
      If cmb_TipBus.ItemData(cmb_TipBus.ListIndex) = 1 Then
         cmb_TipDoc.Enabled = True
         txt_NumDoc.Enabled = True
         msk_NumOpe.Enabled = False
         
         msk_NumOpe.Mask = ""
         msk_NumOpe.Text = ""
         msk_NumOpe.Mask = "###-##-#####"
         
         Call gs_SetFocus(cmb_TipDoc)
      Else
         cmb_TipDoc.Enabled = False
         txt_NumDoc.Enabled = False
         msk_NumOpe.Enabled = True
         
         cmb_TipDoc.ListIndex = -1
         txt_NumDoc.Text = ""
         
         Call gs_SetFocus(msk_NumOpe)
      End If
   Else
      cmb_TipDoc.Enabled = False
      txt_NumDoc.Enabled = False
      
      msk_NumOpe.Enabled = False
   
      cmb_TipDoc.ListIndex = -1
      txt_NumDoc.Text = ""
      msk_NumOpe.Mask = ""
      msk_NumOpe.Text = ""
      msk_NumOpe.Mask = "###-##-#####"
   End If
End Sub

Private Sub cmb_TipBus_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipBus_Click
   End If
End Sub

Private Sub cmb_TipDoc_Click()
   If cmb_TipDoc.ListIndex > -1 Then
      Select Case cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
         Case 1:  txt_NumDoc.MaxLength = 8
         Case 2:  txt_NumDoc.MaxLength = 12
         Case 3:  txt_NumDoc.MaxLength = 12
      End Select
   End If
   Call gs_SetFocus(txt_NumDoc)
End Sub

Private Sub cmb_TipDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipDoc_Click
   End If
End Sub

Private Sub cmd_Buscar_Click()
   If cmb_TipBus.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Búsqueda.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipBus)
      Exit Sub
   End If
   
   If cmb_TipBus.ItemData(cmb_TipBus.ListIndex) = 1 Then
      If cmb_TipDoc.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Tipo de Documento de Identidad.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_TipDoc)
         Exit Sub
      End If
      
      If Len(Trim(txt_NumDoc.Text)) = 0 Then
         MsgBox "Debe ingresar el Número de Documento de Identidad.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NumDoc)
         Exit Sub
      End If
      
      If cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) = 1 Then
         txt_NumDoc.Text = Format(txt_NumDoc.Text, "00000000")
      End If
      
      moddat_g_int_TipDoc = cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
      moddat_g_str_TipDoc = cmb_TipDoc.Text
      moddat_g_str_NumDoc = txt_NumDoc.Text
   Else
      If Len(Trim(msk_NumOpe.Text)) < 10 Then
         MsgBox "Debe ingresar el Número de Operación.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(msk_NumOpe)
         Exit Sub
      End If
      
      moddat_g_str_NumOpe = msk_NumOpe.Text
   End If
   
   If cmb_TipBus.ItemData(cmb_TipBus.ListIndex) = 1 Then
      g_str_Parame = "SELECT * FROM CRE_HIPMAE WHERE "
      g_str_Parame = g_str_Parame & "HIPMAE_TDOCLI = " & CStr(moddat_g_int_TipDoc) & " AND "
      g_str_Parame = g_str_Parame & "HIPMAE_NDOCLI = '" & moddat_g_str_NumDoc & "' AND "
      g_str_Parame = g_str_Parame & "HIPMAE_SITUAC = 2 "
      
   Else
      g_str_Parame = "SELECT * FROM CRE_HIPMAE WHERE "
      g_str_Parame = g_str_Parame & "HIPMAE_NUMOPE = '" & moddat_g_str_NumOpe & "' AND "
      g_str_Parame = g_str_Parame & "HIPMAE_SITUAC = 2 "
   End If
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Call cmd_Limpia_Click
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      MsgBox "No existe ninguna Operación para amortizar para la Búsqueda deseada. ", vbExclamation, modgen_g_str_NomPlt
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Call cmd_Limpia_Click
      Exit Sub
   End If

   Call fs_Buscar_DatGen

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   pnl_NumOpe.Caption = Mid(moddat_g_str_NumOpe, 1, 3) & "-" & Mid(moddat_g_str_NumOpe, 4, 2) & "-" & Mid(moddat_g_str_NumOpe, 6, 5)
   pnl_NomCli.Caption = CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   pnl_MonPre.Caption = moddat_g_str_Moneda
   
   Call fs_Activa(False)

   'Obteniendo Tipo de Cambio
   l_dbl_TCaMPr = moddat_gf_Obtiene_TipCam(1, moddat_g_int_TipMon)
   pnl_TCaMPr.Caption = Format(l_dbl_TCaMPr, "###,###,##0.000000") & " "

   If l_dbl_TCaMPr = 0 Then
      MsgBox "No se encuentra disponible el Tipo de Cambio.", vbExclamation, modgen_g_str_NomPlt
      Call cmd_Limpia_Click
      Exit Sub
   End If

   Call fs_Buscar_Cuotas_Vencidas
   Call fs_Buscar_Cuotas_Vigentes
   
   Call gs_SetFocus(cmb_CodBan)
End Sub

Private Sub cmd_Grabar_Click()
   Dim r_str_Operac        As String
   Dim r_dbl_Deu_Capita    As Double
   Dim r_dbl_Deu_Intere    As Double
   Dim r_dbl_Deu_SegDes    As Double
   Dim r_dbl_Deu_SegViv    As Double
   Dim r_dbl_Deu_OtrCar    As Double
   Dim r_dbl_Deu_IntCom    As Double
   Dim r_dbl_Deu_IntMor    As Double
   Dim r_dbl_Deu_GasCob    As Double
   Dim r_dbl_Deu_OtrGas    As Double
   Dim r_dbl_Deu_TotCuo    As Double
   Dim r_dbl_Pag_Capita    As Double
   Dim r_dbl_Pag_Intere    As Double
   Dim r_dbl_Pag_SegDes    As Double
   Dim r_dbl_Pag_SegViv    As Double
   Dim r_dbl_Pag_OtrCar    As Double
   Dim r_dbl_Pag_IntCom    As Double
   Dim r_dbl_Pag_IntMor    As Double
   Dim r_dbl_Pag_GasCob    As Double
   Dim r_dbl_Pag_TotCuo    As Double
   Dim r_dbl_Pag_OtrGas    As Double
   Dim r_dbl_SalPag        As Double
   Dim r_lng_NumMov        As Long
   Dim r_int_Contad        As Integer
   Dim r_int_NumCuo        As Integer
   Dim r_int_TipMon        As Integer
   Dim r_int_SitCuo        As Integer
   Dim r_int_Situac        As Integer
   Dim r_str_PrxVct        As String
   Dim r_int_FlgCre        As Integer
   
   If cmb_CodBan.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Banco donde se realizo el Pago.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_CodBan)
      Exit Sub
   End If
   
   If cmb_CtaBan.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Cuenta Bancaria.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_CtaBan)
      Exit Sub
   End If
   
   If Len(Trim(txt_NumCom.Text)) = 0 Then
      MsgBox "Debe ingresar el Número de Comprobante.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NumCom)
      Exit Sub
   End If
   
   If CDbl(ipp_Import.Text) = 0 Then
      MsgBox "Debe ingresar el importe pagado.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_Import)
      Exit Sub
   End If
   
   'Si tiene Cuotas Vencidas recalcular Interés Compensatorio, Moratorio y Gastos de Cobranzas a la Fecha de Pago
   'If grd_Listad.Rows > 0 Then
   '   moddat_g_int_FlgGOK = False
   '   moddat_g_int_CntErr = 0
   '
   '   Do While moddat_g_int_FlgGOK = False
   '      g_str_Parame = "USP_CBR0001_CREMOR ("
   '      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "', "
   '      g_str_Parame = g_str_Parame & Format(CDate(ipp_FecPag.Text), "yyyymmdd") & ", "
   '      g_str_Parame = g_str_Parame & "1) "
   '
   '      If Not gf_EjecutaSQL(g_str_Parame, cbr_rst_Accion, 2) Then
   '         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
   '      Else
   '         moddat_g_int_FlgGOK = True
   '      End If
   '
   '      If moddat_g_int_CntErr = 6 Then
   '         If MsgBox("No se pudo completar el procedimiento USP_CBR0001_CREMOR. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
   '            Exit Sub
   '         Else
   '            moddat_g_int_CntErr = 0
   '         End If
   '      End If
   '   Loop
   '
   '   Call gs_LimpiaGrid(grd_Listad)
   '   Call fs_Buscar_Cuotas_Vencidas
   'End If
   
   
   If MsgBox("¿Está seguro de registrar la transacción?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   r_str_Operac = moddat_gf_Consulta_Operac(moddat_g_str_CodPrd, "032")
   
   If moddat_g_int_TipMon = 3 Then
      r_int_TipMon = 1
   Else
      r_int_TipMon = moddat_g_int_TipMon
   End If
   r_str_Operac = CStr(r_int_TipMon) & Right(r_str_Operac, 5)
   
   'Obteniendo Número de Movimiento
   r_lng_NumMov = opecaj_gf_Genera_NumMov()
   
   'Registrando Movimiento
   If Not opecaj_gf_Inserta_CajMov(modgen_g_str_CodUsu, "1102", moddat_g_str_NumOpe, "", moddat_g_int_TipDoc, moddat_g_str_NumDoc, l_arr_CodBan(cmb_CodBan.ListIndex + 1).Genera_Codigo, Format(CDate(ipp_FecPag.Text), "yyyymmdd"), l_arr_CtaBan(cmb_CtaBan.ListIndex + 1).Genera_Codigo, txt_NumCom.Text, r_int_TipMon, CDbl(pnl_NetPag.Caption), 0, modgen_g_str_CodSuc, l_dbl_TCaMPr, 0, 0, l_dbl_PorITF, CDbl(pnl_ITFPag.Caption), CDbl(ipp_Import.Text), 0, "0", r_str_Operac, r_lng_NumMov, 1, "0", "", "", "") Then
      Exit Sub
   End If
   
   'Actualizando Saldo de Caja
   If Not opecaj_gf_ActualizaSaldo(l_arr_CodBan(cmb_CodBan.ListIndex + 1).Genera_Codigo, r_int_TipMon, CDbl(ipp_Import.Text)) Then
      Exit Sub
   End If
   
   'Cambiando a Moneda de Préstamo
   r_str_Operac = CStr(moddat_g_int_TipMon) & Right(r_str_Operac, 5)
   
   'Grabando Pago
   grd_Listad.Redraw = False
   grd_CuoVig.Redraw = False
   
   r_dbl_SalPag = CDbl(Trim(pnl_NetMPr.Caption))
   
   'Cuotas Vencidas
   For r_int_Contad = 0 To grd_Listad.Rows - 1
      grd_Listad.Row = r_int_Contad
      
      grd_Listad.Col = 0:  r_int_NumCuo = CInt(grd_Listad.Text)
      
      grd_Listad.Col = 2:  r_dbl_Deu_Capita = CDbl(grd_Listad.Text)
      grd_Listad.Col = 3:  r_dbl_Deu_Intere = CDbl(grd_Listad.Text)
      grd_Listad.Col = 4:  r_dbl_Deu_SegDes = CDbl(grd_Listad.Text)
      grd_Listad.Col = 5:  r_dbl_Deu_SegViv = CDbl(grd_Listad.Text)
      grd_Listad.Col = 6:  r_dbl_Deu_OtrCar = CDbl(grd_Listad.Text)
      grd_Listad.Col = 7:  r_dbl_Deu_IntMor = CDbl(grd_Listad.Text)
      grd_Listad.Col = 8:  r_dbl_Deu_IntCom = CDbl(grd_Listad.Text)
      grd_Listad.Col = 9:  r_dbl_Deu_GasCob = CDbl(grd_Listad.Text)
      grd_Listad.Col = 10: r_dbl_Deu_OtrGas = CDbl(grd_Listad.Text)
      grd_Listad.Col = 11: r_dbl_Deu_TotCuo = CDbl(grd_Listad.Text)
      
      r_dbl_Pag_Capita = 0:   r_dbl_Pag_Intere = 0:   r_dbl_Pag_SegDes = 0:  r_dbl_Pag_SegViv = 0
      r_dbl_Pag_OtrCar = 0:   r_dbl_Pag_IntMor = 0:   r_dbl_Pag_IntCom = 0:  r_dbl_Pag_GasCob = 0
      r_dbl_Pag_TotCuo = 0
      r_dbl_Pag_OtrGas = 0
      
      r_int_FlgCre = 1
      
      'Otros Gastos
      If r_dbl_Deu_OtrGas > 0 Then
         If r_dbl_SalPag > r_dbl_Deu_OtrGas Then
            r_dbl_SalPag = r_dbl_SalPag - r_dbl_Deu_OtrGas
            r_dbl_Pag_OtrGas = r_dbl_Deu_OtrGas
         Else
            r_dbl_Pag_OtrGas = r_dbl_SalPag
            r_dbl_SalPag = 0
         End If
         r_dbl_Pag_TotCuo = r_dbl_Pag_TotCuo + r_dbl_Pag_OtrGas
      End If
      
      'Seguro Vivienda
      If r_dbl_Deu_SegViv > 0 Then
         If r_dbl_SalPag > r_dbl_Deu_SegViv Then
            r_dbl_SalPag = r_dbl_SalPag - r_dbl_Deu_SegViv
            r_dbl_Pag_SegViv = r_dbl_Deu_SegViv
         Else
            r_dbl_Pag_SegViv = r_dbl_SalPag
            r_dbl_SalPag = 0
         End If
         r_dbl_Pag_TotCuo = r_dbl_Pag_TotCuo + r_dbl_Pag_SegViv
      End If
         
      'Seguro Desgravamen
      If r_dbl_Deu_SegDes > 0 Then
         If r_dbl_SalPag > r_dbl_Deu_SegDes Then
            r_dbl_SalPag = r_dbl_SalPag - r_dbl_Deu_SegDes
            r_dbl_Pag_SegDes = r_dbl_Deu_SegDes
         Else
            r_dbl_Pag_SegDes = r_dbl_SalPag
            r_dbl_SalPag = 0
         End If
         r_dbl_Pag_TotCuo = r_dbl_Pag_TotCuo + r_dbl_Pag_SegDes
      End If
         
      'Otros Cargos
      If r_dbl_Deu_OtrCar > 0 Then
         If r_dbl_SalPag > r_dbl_Deu_OtrCar Then
            r_dbl_SalPag = r_dbl_SalPag - r_dbl_Deu_OtrCar
            r_dbl_Pag_OtrCar = r_dbl_Deu_OtrCar
         Else
            r_dbl_Pag_OtrCar = r_dbl_SalPag
            r_dbl_SalPag = 0
         End If
         r_dbl_Pag_TotCuo = r_dbl_Pag_TotCuo + r_dbl_Pag_OtrCar
      End If
      
      'Gastos de Cobranza
      If r_dbl_Deu_GasCob > 0 Then
         If r_dbl_SalPag > r_dbl_Deu_GasCob Then
            r_dbl_SalPag = r_dbl_SalPag - r_dbl_Deu_GasCob
            r_dbl_Pag_GasCob = r_dbl_Deu_GasCob
         Else
            r_dbl_Pag_GasCob = r_dbl_SalPag
            r_dbl_SalPag = 0
         End If
         r_dbl_Pag_TotCuo = r_dbl_Pag_TotCuo + r_dbl_Pag_GasCob
      End If
         
      'Interes Moratorio
      If r_dbl_Deu_IntMor > 0 Then
         If r_dbl_SalPag > r_dbl_Deu_IntMor Then
            r_dbl_SalPag = r_dbl_SalPag - r_dbl_Deu_IntMor
            r_dbl_Pag_IntMor = r_dbl_Deu_IntMor
         Else
            r_dbl_Pag_IntMor = r_dbl_SalPag
            r_dbl_SalPag = 0
         End If
         r_dbl_Pag_TotCuo = r_dbl_Pag_TotCuo + r_dbl_Pag_IntMor
      End If
         
      'Interes Compensatorio
      If r_dbl_Deu_IntCom > 0 Then
         If r_dbl_SalPag > r_dbl_Deu_IntCom Then
            r_dbl_SalPag = r_dbl_SalPag - r_dbl_Deu_IntCom
            r_dbl_Pag_IntCom = r_dbl_Deu_IntCom
         Else
            r_dbl_Pag_IntCom = r_dbl_SalPag
            r_dbl_SalPag = 0
         End If
         r_dbl_Pag_TotCuo = r_dbl_Pag_TotCuo + r_dbl_Pag_IntCom
      End If
         
      'Interés
      If r_dbl_Deu_Intere > 0 Then
         If r_dbl_SalPag > r_dbl_Deu_Intere Then
            r_dbl_SalPag = r_dbl_SalPag - r_dbl_Deu_Intere
            r_dbl_Pag_Intere = r_dbl_Deu_Intere
         Else
            r_dbl_Pag_Intere = r_dbl_SalPag
            r_dbl_SalPag = 0
         End If
         r_dbl_Pag_TotCuo = r_dbl_Pag_TotCuo + r_dbl_Pag_Intere
      End If
      
      'Capital
      If r_dbl_Deu_Capita > 0 Then
         If r_dbl_SalPag > r_dbl_Deu_Capita Then
            r_dbl_SalPag = r_dbl_SalPag - r_dbl_Deu_Capita
            r_dbl_Pag_Capita = r_dbl_Deu_Capita
         Else
            r_dbl_Pag_Capita = r_dbl_SalPag
            r_dbl_SalPag = 0
         End If
         r_dbl_Pag_TotCuo = r_dbl_Pag_TotCuo + r_dbl_Pag_Capita
      End If
      
      'Grabar Cuota Pago
      If CDbl(CStr(r_dbl_Pag_TotCuo)) = CDbl(CStr(r_dbl_Deu_TotCuo)) Then
         'Cuota Pagada
         r_int_SitCuo = 1
         
         If grd_Listad.Row = grd_Listad.Rows - 1 Then
            If grd_CuoVig.Rows > 0 Then
               grd_CuoVig.Row = 0
               grd_CuoVig.Col = 1
            
               r_str_PrxVct = Format(CDate(grd_CuoVig.Text), "yyyymmdd")
               l_int_CuoPen = l_int_CuoPen - 1
               r_int_Situac = l_int_Situac
            Else
               r_str_PrxVct = "0"
               l_int_CuoPen = 0
               r_int_Situac = 9
            End If
         Else
            'Fecha de Proximo Vcto
            grd_Listad.Row = grd_Listad.Row + 1
            grd_Listad.Col = 1
            r_str_PrxVct = Format(CDate(grd_Listad.Text), "yyyymmdd")
            
            'Cuotas Pendientes
            l_int_CuoPen = l_int_CuoPen - 1
            
            'Situación
            r_int_Situac = l_int_Situac
         End If
      Else
         r_int_SitCuo = 2
      
         'Fecha de Proximo Vcto
         grd_Listad.Col = 1
         r_str_PrxVct = Format(CDate(grd_Listad.Text), "yyyymmdd")
         
         r_int_Situac = l_int_Situac
      End If
   
      r_dbl_SalPag = CDbl(Format(r_dbl_SalPag, "#########0.00"))
      
      If r_dbl_SalPag = 0 Then
         r_int_FlgCre = 2
      End If
      
      If Not opecaj_gf_Pago_Cuotas(moddat_g_str_NumOpe, r_int_NumCuo, Format(CDate(ipp_FecPag.Text), "yyyymmdd"), r_dbl_Pag_TotCuo, r_dbl_Pag_Capita, r_dbl_Pag_Intere, r_dbl_Pag_SegDes, r_dbl_Pag_SegViv, r_dbl_Pag_OtrCar, r_dbl_Pag_IntCom, r_dbl_Pag_IntMor, r_dbl_Pag_GasCob, r_dbl_Pag_OtrGas, l_dbl_TCaDol, l_dbl_TCaMPr, l_int_SitCre, r_str_Operac, r_lng_NumMov, r_int_SitCuo, r_str_PrxVct, l_int_CuoPen, r_int_Situac, l_int_SitAnt, r_int_FlgCre) Then
      End If
         
      If r_dbl_SalPag <= 0 Then
         Exit For
      End If
   Next r_int_Contad
   
   If r_dbl_SalPag > 0 Then
      'Cuotas Vigentes
      For r_int_Contad = 0 To grd_CuoVig.Rows - 1
         grd_CuoVig.Row = r_int_Contad
         
         grd_CuoVig.Col = 0:  r_int_NumCuo = CInt(grd_CuoVig.Text)
         
         grd_CuoVig.Col = 2:  r_dbl_Deu_Capita = CDbl(grd_CuoVig.Text)
         grd_CuoVig.Col = 3:  r_dbl_Deu_Intere = CDbl(grd_CuoVig.Text)
         grd_CuoVig.Col = 4:  r_dbl_Deu_SegDes = CDbl(grd_CuoVig.Text)
         grd_CuoVig.Col = 5:  r_dbl_Deu_SegViv = CDbl(grd_CuoVig.Text)
         grd_CuoVig.Col = 6:  r_dbl_Deu_OtrCar = CDbl(grd_CuoVig.Text)
         grd_CuoVig.Col = 7:  r_dbl_Deu_TotCuo = CDbl(grd_CuoVig.Text)
         
         r_dbl_Pag_Capita = 0:   r_dbl_Pag_Intere = 0:   r_dbl_Pag_SegDes = 0:  r_dbl_Pag_SegViv = 0
         r_dbl_Pag_OtrCar = 0:   r_dbl_Pag_IntMor = 0:   r_dbl_Pag_IntCom = 0:  r_dbl_Pag_GasCob = 0
         r_dbl_Pag_TotCuo = 0
         
         r_int_FlgCre = 1
         
         'Seguro Vivienda
         If r_dbl_Deu_SegViv > 0 Then
            If r_dbl_SalPag > r_dbl_Deu_SegViv Then
               r_dbl_SalPag = r_dbl_SalPag - r_dbl_Deu_SegViv
               r_dbl_Pag_SegViv = r_dbl_Deu_SegViv
            Else
               r_dbl_Pag_SegViv = r_dbl_SalPag
               r_dbl_SalPag = 0
            End If
            r_dbl_Pag_TotCuo = r_dbl_Pag_TotCuo + r_dbl_Pag_SegViv
         End If
            
         'Seguro Desgravamen
         If r_dbl_Deu_SegDes > 0 Then
            If r_dbl_SalPag > r_dbl_Deu_SegDes Then
               r_dbl_SalPag = r_dbl_SalPag - r_dbl_Deu_SegDes
               r_dbl_Pag_SegDes = r_dbl_Deu_SegDes
            Else
               r_dbl_Pag_SegDes = r_dbl_SalPag
               r_dbl_SalPag = 0
            End If
            r_dbl_Pag_TotCuo = r_dbl_Pag_TotCuo + r_dbl_Pag_SegDes
         End If
            
         'Otros Cargos
         If r_dbl_Deu_OtrCar > 0 Then
            If r_dbl_SalPag > r_dbl_Deu_OtrCar Then
               r_dbl_SalPag = r_dbl_SalPag - r_dbl_Deu_OtrCar
               r_dbl_Pag_OtrCar = r_dbl_Deu_OtrCar
            Else
               r_dbl_Pag_OtrCar = r_dbl_SalPag
               r_dbl_SalPag = 0
            End If
            r_dbl_Pag_TotCuo = r_dbl_Pag_TotCuo + r_dbl_Pag_OtrCar
         End If
            
         'Interés
         If r_dbl_Deu_Intere > 0 Then
            If r_dbl_SalPag > r_dbl_Deu_Intere Then
               r_dbl_SalPag = r_dbl_SalPag - r_dbl_Deu_Intere
               r_dbl_Pag_Intere = r_dbl_Deu_Intere
            Else
               r_dbl_Pag_Intere = r_dbl_SalPag
               r_dbl_SalPag = 0
            End If
            r_dbl_Pag_TotCuo = r_dbl_Pag_TotCuo + r_dbl_Pag_Intere
         End If
         
         'Capital
         If r_dbl_Deu_Capita > 0 Then
            If r_dbl_SalPag > r_dbl_Deu_Capita Then
               r_dbl_SalPag = r_dbl_SalPag - r_dbl_Deu_Capita
               r_dbl_Pag_Capita = r_dbl_Deu_Capita
            Else
               r_dbl_Pag_Capita = r_dbl_SalPag
               r_dbl_SalPag = 0
            End If
            r_dbl_Pag_TotCuo = r_dbl_Pag_TotCuo + r_dbl_Pag_Capita
         End If
         
         'Grabar Cuota Pago
         If CDbl(CStr(r_dbl_Pag_TotCuo)) = CDbl(CStr(r_dbl_Deu_TotCuo)) Then
            'Cuota Pagada
            r_int_SitCuo = 1
            
            If grd_CuoVig.Row = grd_CuoVig.Rows - 1 Then
               r_str_PrxVct = "0"
               l_int_CuoPen = 0
               r_int_Situac = 9
            Else
               'Fecha de Proximo Vcto
               grd_CuoVig.Row = grd_CuoVig.Row + 1
               grd_CuoVig.Col = 1
               r_str_PrxVct = Format(CDate(grd_CuoVig.Text), "yyyymmdd")
               
               'Cuotas Pendientes
               l_int_CuoPen = l_int_CuoPen - 1
               
               'Situación
               r_int_Situac = l_int_Situac
            End If
         Else
            r_int_SitCuo = 2
         
            'Fecha de Proximo Vcto
            grd_CuoVig.Col = 1
            r_str_PrxVct = Format(CDate(grd_CuoVig.Text), "yyyymmdd")
            
            r_int_Situac = l_int_Situac
         End If
      
         r_dbl_SalPag = CDbl(Format(r_dbl_SalPag, "#########0.00"))
         
         If r_dbl_SalPag = 0 Then
            r_int_FlgCre = 2
         End If
         
         If Not opecaj_gf_Pago_Cuotas(moddat_g_str_NumOpe, r_int_NumCuo, Format(CDate(ipp_FecPag.Text), "yyyymmdd"), r_dbl_Pag_TotCuo, r_dbl_Pag_Capita, r_dbl_Pag_Intere, r_dbl_Pag_SegDes, r_dbl_Pag_SegViv, r_dbl_Pag_OtrCar, 0, 0, 0, 0, l_dbl_TCaDol, l_dbl_TCaMPr, l_int_SitCre, r_str_Operac, r_lng_NumMov, r_int_SitCuo, r_str_PrxVct, l_int_CuoPen, r_int_Situac, l_int_SitAnt, r_int_FlgCre) Then
            Exit Sub
         End If

         If r_dbl_SalPag <= 0 Then
            Exit For
         End If
      Next r_int_Contad
   End If
   
   grd_Listad.Redraw = True
   grd_CuoVig.Redraw = True
   
   'Call opecaj_gs_Imp_CuoHip_Ban(moddat_g_str_NumOpe, Format(CDate(moddat_g_str_FecSis), "yyyymmdd"), CStr(r_lng_NumMov))
   'Call gs_Imprim_ComPag
   
   Call cmd_Limpia_Click
End Sub

Private Sub cmd_Limpia_Click()
   Call fs_Limpia
   Call gs_SetFocus(cmb_TipBus)
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call cmd_Limpia_Click
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub ipp_FecPag_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NumCom)
   End If
End Sub

Private Sub ipp_Import_Change()
   pnl_ITFPag.Caption = gf_Truncar_Numero(CDbl(ipp_Import.Text) * (l_dbl_PorITF / 100), 2) & " "
   pnl_NetPag.Caption = Format(CDbl(ipp_Import.Text) - CDbl(Trim(pnl_ITFPag.Caption)), "###,###,##0.00") & " "
   
   'Si Moneda es Vac se ingresa el Importe en Soles y se convierte en VAC
   If moddat_g_int_TipMon = 3 Then
      pnl_NetMPr.Caption = Format(CDbl(pnl_NetPag.Caption) / l_dbl_TCaMPr, "###,###,##0.00") & " "
   Else
      pnl_NetMPr.Caption = Format(CDbl(pnl_NetPag.Caption), "###,###,##0.00") & " "
   End If
End Sub

Private Sub ipp_Import_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   End If
End Sub

Private Sub msk_NumOpe_GotFocus()
   Call gs_SelecTodo(msk_NumOpe)
End Sub

Private Sub msk_NumOpe_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Buscar)
   End If
End Sub

Private Sub txt_NumCom_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_Import)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & "-")
   End If
End Sub

Private Sub txt_NumCom_GotFocus()
   Call gs_SelecTodo(txt_NumCom)
End Sub

Private Sub txt_NumDoc_GotFocus()
   Call gs_SelecTodo(txt_NumDoc)
End Sub

Private Sub txt_NumDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Buscar)
   Else
      If cmb_TipDoc.ListIndex > -1 Then
         Select Case cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)
            Case 1:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
            Case 2:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-")
            Case 3:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-")
         End Select
      Else
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub fs_Buscar_DatGen()
   g_rst_Princi.MoveFirst
   
   moddat_g_int_TipDoc = g_rst_Princi!HIPMAE_TDOCLI
   moddat_g_str_NumDoc = Trim(g_rst_Princi!HIPMAE_NDOCLI)
   moddat_g_str_NumSol = Trim(g_rst_Princi!HIPMAE_NUMSOL)
   moddat_g_str_NumOpe = Trim(g_rst_Princi!HIPMAE_NUMOPE)
   
   'Obteniendo Nombre de Cliente
   moddat_g_str_NomCli = moddat_gf_Buscar_NomCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc)
   
   'Obteniendo Nombre y DOI de Cónyuge
   moddat_g_int_CygTDo = g_rst_Princi!HIPMAE_TDOCYG
   moddat_g_str_CygNDo = Trim(g_rst_Princi!HIPMAE_NDOCYG & "")
   moddat_g_str_CygNom = moddat_gf_Buscar_NomCli(moddat_g_int_CygTDo, moddat_g_str_CygNDo)
   
   'Obteniendo Descripción de Producto
   moddat_g_str_CodPrd = Trim(g_rst_Princi!HIPMAE_CODPRD)
   moddat_g_str_NomPrd = moddat_gf_Consulta_Produc(Trim(g_rst_Princi!HIPMAE_CODPRD))

   'Obeniendo Modalidad de Producto
   moddat_g_str_CodMod = Trim(g_rst_Princi!HIPMAE_CODMOD)
   moddat_g_str_DesMod = moddat_gf_Buscar_NomMod(Trim(g_rst_Princi!HIPMAE_CODPRD), moddat_g_str_CodMod)
   

   'Moneda
   moddat_g_str_Moneda = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!HIPMAE_MONEDA))
   moddat_g_int_TipMon = g_rst_Princi!HIPMAE_MONEDA
   
   'Situación de Crédito SBS
   l_int_SitCre = g_rst_Princi!HIPMAE_SITCRE

   'Situación Anterior
   l_int_SitAnt = g_rst_Princi!HIPMAE_SITANT

   'Situación de Crédito
   l_int_Situac = g_rst_Princi!HIPMAE_SITUAC

   'Cuotas Pendientes
   l_int_CuoPen = g_rst_Princi!HIPMAE_CUOPEN
   
   'Obteniendo ITF
   If g_rst_Princi!HIPMAE_INDITF = 2 Then
      l_dbl_PorITF = 0
   Else
      l_dbl_PorITF = opecaj_gf_Consulta_ITF(Format(CDate(moddat_g_str_FecSis), "yyyymmdd"), 1)
   End If
End Sub

Private Sub fs_Limpia()
   Call fs_Activa(True)
   
   cmb_TipBus.ListIndex = -1
   cmb_TipDoc.Enabled = False
   txt_NumDoc.Enabled = False
   msk_NumOpe.Enabled = False

   msk_NumOpe.Mask = ""
   msk_NumOpe.Text = ""
   msk_NumOpe.Mask = "###-##-#####"
   
   txt_NumDoc.Text = ""
   
   pnl_NumOpe.Caption = ""
   pnl_NomCli.Caption = ""
   pnl_MonPre.Caption = ""
   
   Call gs_LimpiaGrid(grd_Listad)
   Call gs_LimpiaGrid(grd_CuoVig)
   
   cmb_CodBan.ListIndex = -1
   cmb_CtaBan.Clear
   ipp_FecPag.Text = Format(Date, "dd/mm/yyyy")
   txt_NumCom.Text = ""
   ipp_Import.Value = 0
   
   pnl_ITFPag.Caption = "0.00 "
   pnl_NetPag.Caption = "0.00 "
   pnl_NetMPr.Caption = "0.00 "
   pnl_TCaMPr.Caption = "0.000000 "
End Sub

Private Sub fs_Activa(ByVal p_Habilita As Integer)
   cmb_TipBus.Enabled = p_Habilita
   cmb_TipDoc.Enabled = p_Habilita
   txt_NumDoc.Enabled = p_Habilita
   msk_NumOpe.Enabled = p_Habilita
   cmd_Buscar.Enabled = p_Habilita
   
   grd_Listad.Enabled = Not p_Habilita
   grd_CuoVig.Enabled = Not p_Habilita
   cmb_CodBan.Enabled = Not p_Habilita
   cmb_CtaBan.Enabled = Not p_Habilita
   ipp_FecPag.Enabled = Not p_Habilita
   txt_NumCom.Enabled = Not p_Habilita
   ipp_Import.Enabled = Not p_Habilita
   cmd_Grabar.Enabled = Not p_Habilita
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_FecSis
   
   grd_Listad.ColWidth(0) = 575
   grd_Listad.ColWidth(1) = 965
   grd_Listad.ColWidth(2) = 1085
   grd_Listad.ColWidth(3) = 1085
   grd_Listad.ColWidth(4) = 1085
   grd_Listad.ColWidth(5) = 1085
   grd_Listad.ColWidth(6) = 1085
   grd_Listad.ColWidth(7) = 1085
   grd_Listad.ColWidth(8) = 1085
   grd_Listad.ColWidth(9) = 1085
   grd_Listad.ColWidth(10) = 1085
   grd_Listad.ColWidth(11) = 1085
   
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignRightCenter
   grd_Listad.ColAlignment(3) = flexAlignRightCenter
   grd_Listad.ColAlignment(4) = flexAlignRightCenter
   grd_Listad.ColAlignment(5) = flexAlignRightCenter
   grd_Listad.ColAlignment(6) = flexAlignRightCenter
   grd_Listad.ColAlignment(7) = flexAlignRightCenter
   grd_Listad.ColAlignment(8) = flexAlignRightCenter
   grd_Listad.ColAlignment(9) = flexAlignRightCenter
   grd_Listad.ColAlignment(10) = flexAlignRightCenter
   grd_Listad.ColAlignment(11) = flexAlignRightCenter
   
   grd_CuoVig.ColWidth(0) = 1230
   grd_CuoVig.ColWidth(1) = 1580
   grd_CuoVig.ColWidth(2) = 1580
   grd_CuoVig.ColWidth(3) = 1595
   grd_CuoVig.ColWidth(4) = 1595
   grd_CuoVig.ColWidth(5) = 1595
   grd_CuoVig.ColWidth(6) = 1595
   grd_CuoVig.ColWidth(7) = 1595
   
   grd_CuoVig.ColAlignment(0) = flexAlignCenterCenter
   grd_CuoVig.ColAlignment(1) = flexAlignCenterCenter
   grd_CuoVig.ColAlignment(2) = flexAlignRightCenter
   grd_CuoVig.ColAlignment(3) = flexAlignRightCenter
   grd_CuoVig.ColAlignment(4) = flexAlignRightCenter
   grd_CuoVig.ColAlignment(5) = flexAlignRightCenter
   grd_CuoVig.ColAlignment(6) = flexAlignRightCenter
   grd_CuoVig.ColAlignment(7) = flexAlignRightCenter
   
   Call moddat_gs_Carga_LisIte(cmb_CodBan, l_arr_CodBan, 1, "516")
   Call modsis_gs_Carga_TipBus_1(cmb_TipBus)
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipDoc, 1, "230")
   
   'Obteniendo ITF
   l_dbl_PorITF = opecaj_gf_Consulta_ITF(Format(CDate(moddat_g_str_FecSis), "yyyymmdd"), 1)
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

Private Sub grd_CuoVig_SelChange()
   If grd_CuoVig.Rows > 2 Then
      grd_CuoVig.RowSel = grd_CuoVig.Row
   End If
End Sub

Private Sub fs_Buscar_Cuotas_Vencidas()
   Dim r_dbl_SegDes  As Double
   Dim r_dbl_SegViv  As Double
   Dim r_dbl_OtrCar  As Double
   Dim r_dbl_ValCuo  As Double

   'Cuotas Vencidas
   g_str_Parame = "SELECT * FROM CRE_HIPCUO WHERE "
   g_str_Parame = g_str_Parame & "HIPCUO_NUMOPE = '" & moddat_g_str_NumOpe & "' AND "
   g_str_Parame = g_str_Parame & "HIPCUO_TIPCRO = 1 AND "
   g_str_Parame = g_str_Parame & "HIPCUO_SITUAC = 2 AND "
   g_str_Parame = g_str_Parame & "HIPCUO_FECVCT <= " & Format(Date, "yyyymmdd")
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Call cmd_Limpia_Click
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_Listad.Redraw = False
      
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         
         r_dbl_ValCuo = 0
         
         grd_Listad.Col = 0
         grd_Listad.Text = Format(g_rst_Princi!HIPCUO_NUMCUO, "000")
      
         grd_Listad.Col = 1
         grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECVCT))
      
         'Capital
         grd_Listad.Col = 2
         grd_Listad.Text = Format(g_rst_Princi!HIPCUO_CAPITA + g_rst_Princi!HIPCUO_CAPBBP - g_rst_Princi!HIPCUO_CAPPAG, "###,###,##0.00")
         r_dbl_ValCuo = r_dbl_ValCuo + CDbl(grd_Listad.Text)
      
         'Interes
         grd_Listad.Col = 3
         grd_Listad.Text = Format(g_rst_Princi!HIPCUO_INTERE + g_rst_Princi!HIPCUO_INTBBP - g_rst_Princi!HIPCUO_INTPAG, "###,###,##0.00")
         r_dbl_ValCuo = r_dbl_ValCuo + CDbl(grd_Listad.Text)
      
         'Seguro de Desgravamen
         grd_Listad.Col = 4
         grd_Listad.Text = Format(g_rst_Princi!HIPCUO_DESORG - g_rst_Princi!HIPCUO_DESPAG, "###,###,##0.00")
         r_dbl_ValCuo = r_dbl_ValCuo + CDbl(grd_Listad.Text)
      
         'Seguro de Vivienda
         grd_Listad.Col = 5
         grd_Listad.Text = Format(g_rst_Princi!HIPCUO_VIVORG - g_rst_Princi!HIPCUO_VIVPAG, "###,###,##0.00")
         r_dbl_ValCuo = r_dbl_ValCuo + CDbl(grd_Listad.Text)
      
         'Otros Cargos
         grd_Listad.Col = 6
         grd_Listad.Text = Format(g_rst_Princi!HIPCUO_OTRORG - g_rst_Princi!HIPCUO_OTRPAG, "###,###,##0.00")
         r_dbl_ValCuo = r_dbl_ValCuo + CDbl(grd_Listad.Text)
      
         'Interes Moratorio
         grd_Listad.Col = 7
         grd_Listad.Text = Format(g_rst_Princi!HIPCUO_INTMOR - g_rst_Princi!HIPCUO_IMOPAG, "###,###,##0.00")
         r_dbl_ValCuo = r_dbl_ValCuo + CDbl(grd_Listad.Text)
      
         'Interes Compensatorio
         grd_Listad.Col = 8
         grd_Listad.Text = Format(g_rst_Princi!HIPCUO_INTCOM - g_rst_Princi!HIPCUO_ICOPAG, "###,###,##0.00")
         r_dbl_ValCuo = r_dbl_ValCuo + CDbl(grd_Listad.Text)
      
         'Gastos de Cobranza
         grd_Listad.Col = 9
         grd_Listad.Text = Format(g_rst_Princi!HIPCUO_GASCOB - g_rst_Princi!HIPCUO_GCOPAG, "###,###,##0.00")
         r_dbl_ValCuo = r_dbl_ValCuo + CDbl(grd_Listad.Text)
      
         'Otros Gastos
         grd_Listad.Col = 10
         grd_Listad.Text = Format(g_rst_Princi!HIPCUO_OTRGAS - g_rst_Princi!HIPCUO_OTGPAG, "###,###,##0.00")
         r_dbl_ValCuo = r_dbl_ValCuo + CDbl(grd_Listad.Text)
      
         'Valor Cuota
         grd_Listad.Col = 11
         grd_Listad.Text = Format(r_dbl_ValCuo, "###,###,##0.00")
      
         g_rst_Princi.MoveNext
      Loop
      grd_Listad.Redraw = True
      
      Call gs_UbiIniGrid(grd_Listad)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Buscar_Cuotas_Vigentes()
   Dim r_dbl_SegDes  As Double
   Dim r_dbl_SegViv  As Double
   Dim r_dbl_OtrCar  As Double
   Dim r_dbl_ValCuo  As Double
   
   'Cuotas x Vencer
   g_str_Parame = "SELECT * FROM CRE_HIPCUO WHERE "
   g_str_Parame = g_str_Parame & "HIPCUO_NUMOPE = '" & moddat_g_str_NumOpe & "' AND "
   g_str_Parame = g_str_Parame & "HIPCUO_TIPCRO = 1 AND "
   g_str_Parame = g_str_Parame & "HIPCUO_SITUAC = 2 AND "
   g_str_Parame = g_str_Parame & "HIPCUO_FECVCT > " & Format(Date, "yyyymmdd")
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Call cmd_Limpia_Click
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_CuoVig.Redraw = False
      
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         grd_CuoVig.Rows = grd_CuoVig.Rows + 1
         grd_CuoVig.Row = grd_CuoVig.Rows - 1
         
         r_dbl_ValCuo = 0
         
         grd_CuoVig.Col = 0
         grd_CuoVig.Text = Format(g_rst_Princi!HIPCUO_NUMCUO, "000")
      
         grd_CuoVig.Col = 1
         grd_CuoVig.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECVCT))
         
         'Capital
         grd_CuoVig.Col = 2
         grd_CuoVig.Text = Format(g_rst_Princi!HIPCUO_CAPITA + g_rst_Princi!HIPCUO_CAPBBP - g_rst_Princi!HIPCUO_CAPPAG, "###,###,##0.00")
         r_dbl_ValCuo = r_dbl_ValCuo + CDbl(grd_CuoVig.Text)
      
         'Interés
         grd_CuoVig.Col = 3
         grd_CuoVig.Text = Format(g_rst_Princi!HIPCUO_INTERE + g_rst_Princi!HIPCUO_INTBBP - g_rst_Princi!HIPCUO_INTPAG, "###,###,##0.00")
         r_dbl_ValCuo = r_dbl_ValCuo + CDbl(grd_CuoVig.Text)
      
         'Seguro de Desgravamen
         grd_CuoVig.Col = 4
         grd_CuoVig.Text = Format(g_rst_Princi!HIPCUO_DESORG - g_rst_Princi!HIPCUO_DESPAG, "###,###,##0.00")
         r_dbl_ValCuo = r_dbl_ValCuo + CDbl(grd_CuoVig.Text)
      
         'Seguro de Vivienda
         grd_CuoVig.Col = 5
         grd_CuoVig.Text = Format(g_rst_Princi!HIPCUO_VIVORG - g_rst_Princi!HIPCUO_VIVPAG, "###,###,##0.00")
         r_dbl_ValCuo = r_dbl_ValCuo + CDbl(grd_CuoVig.Text)
      
         'Otros Cargos
         grd_CuoVig.Col = 6
         grd_CuoVig.Text = Format(g_rst_Princi!HIPCUO_OTRORG - g_rst_Princi!HIPCUO_OTRPAG, "###,###,##0.00")
         r_dbl_ValCuo = r_dbl_ValCuo + CDbl(grd_CuoVig.Text)
      
         'Valor Cuota
         grd_CuoVig.Col = 7
         grd_CuoVig.Text = Format(r_dbl_ValCuo, "###,###,##0.00")
      
         g_rst_Princi.MoveNext
      Loop
   
      grd_CuoVig.Redraw = True
      Call gs_UbiIniGrid(grd_CuoVig)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub
