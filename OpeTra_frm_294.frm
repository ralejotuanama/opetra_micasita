VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_Pro_EvaPBP_04 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   10320
   ClientLeft      =   3900
   ClientTop       =   495
   ClientWidth     =   10980
   Icon            =   "OpeTra_frm_294.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10320
   ScaleWidth      =   10980
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   10335
      Left            =   0
      TabIndex        =   35
      Top             =   0
      Width           =   10995
      _Version        =   65536
      _ExtentX        =   19394
      _ExtentY        =   18230
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
      Begin Threed.SSPanel SSPanel13 
         Height          =   1995
         Left            =   30
         TabIndex        =   46
         Top             =   4950
         Width           =   10905
         _Version        =   65536
         _ExtentX        =   19235
         _ExtentY        =   3519
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
         Begin Threed.SSPanel pnl_CuoCon 
            Height          =   315
            Left            =   2640
            TabIndex        =   47
            Top             =   60
            Width           =   855
            _Version        =   65536
            _ExtentX        =   1508
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
         End
         Begin Threed.SSPanel pnl_VctCof 
            Height          =   315
            Left            =   2640
            TabIndex        =   49
            Top             =   1290
            Width           =   1365
            _Version        =   65536
            _ExtentX        =   2408
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
         End
         Begin Threed.SSPanel pnl_CuoCof 
            Height          =   315
            Left            =   2640
            TabIndex        =   51
            Top             =   1620
            Width           =   1365
            _Version        =   65536
            _ExtentX        =   2408
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_CuoCli 
            Height          =   315
            Left            =   2640
            TabIndex        =   53
            Top             =   810
            Width           =   1365
            _Version        =   65536
            _ExtentX        =   2408
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_VctCli 
            Height          =   315
            Left            =   2640
            TabIndex        =   55
            Top             =   480
            Width           =   1365
            _Version        =   65536
            _ExtentX        =   2408
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
         End
         Begin Threed.SSPanel SSPanel19 
            Height          =   45
            Left            =   30
            TabIndex        =   57
            Top             =   420
            Width           =   10845
            _Version        =   65536
            _ExtentX        =   19129
            _ExtentY        =   79
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
         Begin Threed.SSPanel SSPanel20 
            Height          =   45
            Left            =   30
            TabIndex        =   58
            Top             =   1200
            Width           =   10845
            _Version        =   65536
            _ExtentX        =   19129
            _ExtentY        =   79
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
         Begin Threed.SSPanel pnl_Tot_Capita_Cli 
            Height          =   315
            Left            =   8130
            TabIndex        =   99
            Top             =   810
            Width           =   885
            _Version        =   65536
            _ExtentX        =   1561
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_Tot_Intere_Cli 
            Height          =   315
            Left            =   9060
            TabIndex        =   100
            Top             =   810
            Width           =   885
            _Version        =   65536
            _ExtentX        =   1561
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_Tot_Capita_Cof 
            Height          =   315
            Left            =   8130
            TabIndex        =   102
            Top             =   1620
            Width           =   885
            _Version        =   65536
            _ExtentX        =   1561
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_Tot_Intere_Cof 
            Height          =   315
            Left            =   9060
            TabIndex        =   103
            Top             =   1620
            Width           =   885
            _Version        =   65536
            _ExtentX        =   1561
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_Tot_Comisi_Cof 
            Height          =   315
            Left            =   9990
            TabIndex        =   104
            Top             =   1620
            Width           =   885
            _Version        =   65536
            _ExtentX        =   1561
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
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_Pen_CuoIni 
            Height          =   315
            Left            =   8130
            TabIndex        =   108
            Top             =   60
            Width           =   885
            _Version        =   65536
            _ExtentX        =   1561
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
         End
         Begin Threed.SSPanel pnl_Pen_CuoFin 
            Height          =   315
            Left            =   9060
            TabIndex        =   109
            Top             =   60
            Width           =   885
            _Version        =   65536
            _ExtentX        =   1561
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
         End
         Begin VB.Label Label1 
            Caption         =   "Rango de Cuotas a Penalizar TNC:"
            Height          =   255
            Left            =   5250
            TabIndex        =   110
            Top             =   60
            Width           =   2595
         End
         Begin VB.Label Label10 
            Caption         =   "Capital, Interes:"
            Height          =   315
            Left            =   5250
            TabIndex        =   107
            Top             =   810
            Width           =   1875
         End
         Begin VB.Label Label9 
            Caption         =   "Capital, Interes, Comisión:"
            Height          =   315
            Left            =   5250
            TabIndex        =   106
            Top             =   1620
            Width           =   1875
         End
         Begin VB.Label lbl_SimMon 
            Alignment       =   1  'Right Justify
            Caption         =   "S/."
            Height          =   315
            Index           =   17
            Left            =   7350
            TabIndex        =   105
            Top             =   1620
            Width           =   645
         End
         Begin VB.Label lbl_SimMon 
            Alignment       =   1  'Right Justify
            Caption         =   "S/."
            Height          =   315
            Index           =   2
            Left            =   7350
            TabIndex        =   101
            Top             =   810
            Width           =   645
         End
         Begin VB.Label lbl_SimMon 
            Alignment       =   1  'Right Justify
            Caption         =   "S/."
            Height          =   315
            Index           =   1
            Left            =   1860
            TabIndex        =   61
            Top             =   810
            Width           =   645
         End
         Begin VB.Label lbl_SimMon 
            Alignment       =   1  'Right Justify
            Caption         =   "S/."
            Height          =   315
            Index           =   0
            Left            =   1860
            TabIndex        =   60
            Top             =   1620
            Width           =   645
         End
         Begin VB.Label Label8 
            Caption         =   "Fecha Vcto. (Cliente):"
            Height          =   315
            Left            =   60
            TabIndex        =   56
            Top             =   510
            Width           =   2415
         End
         Begin VB.Label Label7 
            Caption         =   "Monto Cuota:"
            Height          =   315
            Left            =   60
            TabIndex        =   54
            Top             =   1620
            Width           =   1185
         End
         Begin VB.Label Label6 
            Caption         =   "Monto Cuota (Cliente)"
            Height          =   315
            Left            =   60
            TabIndex        =   52
            Top             =   810
            Width           =   1605
         End
         Begin VB.Label Label5 
            Caption         =   "F. Vcto. (Cofide(Mivivienda):"
            Height          =   315
            Left            =   60
            TabIndex        =   50
            Top             =   1320
            Width           =   2055
         End
         Begin VB.Label Label4 
            Caption         =   "Cuota TC a evaluar:"
            Height          =   315
            Left            =   60
            TabIndex        =   48
            Top             =   60
            Width           =   1605
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   435
         Left            =   30
         TabIndex        =   97
         Top             =   6990
         Width           =   10905
         _Version        =   65536
         _ExtentX        =   19235
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
         Begin VB.ComboBox cmb_FlgPBP 
            Height          =   315
            Left            =   2640
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   1365
         End
         Begin VB.Label Label11 
            Caption         =   "Asignar PBP:"
            Height          =   255
            Left            =   60
            TabIndex        =   98
            Top             =   60
            Width           =   1035
         End
      End
      Begin Threed.SSPanel SSPanel21 
         Height          =   2805
         Left            =   30
         TabIndex        =   59
         Top             =   7470
         Width           =   10905
         _Version        =   65536
         _ExtentX        =   19235
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
         Begin EditLib.fpDoubleSingle ipp_CapPen_Cli_1 
            Height          =   315
            Left            =   2640
            TabIndex        =   1
            Top             =   420
            Width           =   1035
            _Version        =   196608
            _ExtentX        =   1826
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
         Begin EditLib.fpDoubleSingle ipp_IntPen_Cli_1 
            Height          =   315
            Left            =   3690
            TabIndex        =   2
            Top             =   420
            Width           =   1035
            _Version        =   196608
            _ExtentX        =   1826
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
         Begin EditLib.fpDoubleSingle ipp_CapPen_Cli_2 
            Height          =   315
            Left            =   2640
            TabIndex        =   3
            Top             =   750
            Width           =   1035
            _Version        =   196608
            _ExtentX        =   1826
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
         Begin EditLib.fpDoubleSingle ipp_IntPen_Cli_2 
            Height          =   315
            Left            =   3690
            TabIndex        =   4
            Top             =   750
            Width           =   1035
            _Version        =   196608
            _ExtentX        =   1826
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
         Begin EditLib.fpDoubleSingle ipp_CapPen_Cli_3 
            Height          =   315
            Left            =   2640
            TabIndex        =   5
            Top             =   1080
            Width           =   1035
            _Version        =   196608
            _ExtentX        =   1826
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
         Begin EditLib.fpDoubleSingle ipp_IntPen_Cli_3 
            Height          =   315
            Left            =   3690
            TabIndex        =   6
            Top             =   1080
            Width           =   1035
            _Version        =   196608
            _ExtentX        =   1826
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
         Begin EditLib.fpDoubleSingle ipp_CapPen_Cli_4 
            Height          =   315
            Left            =   2640
            TabIndex        =   7
            Top             =   1410
            Width           =   1035
            _Version        =   196608
            _ExtentX        =   1826
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
         Begin EditLib.fpDoubleSingle ipp_IntPen_Cli_4 
            Height          =   315
            Left            =   3690
            TabIndex        =   8
            Top             =   1410
            Width           =   1035
            _Version        =   196608
            _ExtentX        =   1826
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
         Begin EditLib.fpDoubleSingle ipp_CapPen_Cli_5 
            Height          =   315
            Left            =   2640
            TabIndex        =   9
            Top             =   1740
            Width           =   1035
            _Version        =   196608
            _ExtentX        =   1826
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
         Begin EditLib.fpDoubleSingle ipp_IntPen_Cli_5 
            Height          =   315
            Left            =   3690
            TabIndex        =   10
            Top             =   1740
            Width           =   1035
            _Version        =   196608
            _ExtentX        =   1826
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
         Begin EditLib.fpDoubleSingle ipp_CapPen_Cli_6 
            Height          =   315
            Left            =   2640
            TabIndex        =   11
            Top             =   2070
            Width           =   1035
            _Version        =   196608
            _ExtentX        =   1826
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
         Begin EditLib.fpDoubleSingle ipp_IntPen_Cli_6 
            Height          =   315
            Left            =   3690
            TabIndex        =   12
            Top             =   2070
            Width           =   1035
            _Version        =   196608
            _ExtentX        =   1826
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
         Begin Threed.SSPanel pnl_Ctr_CapPen_Cli 
            Height          =   315
            Left            =   2640
            TabIndex        =   70
            Top             =   2400
            Width           =   1035
            _Version        =   65536
            _ExtentX        =   1826
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   192
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
         Begin Threed.SSPanel pnl_Ctr_IntPen_Cli 
            Height          =   315
            Left            =   3690
            TabIndex        =   72
            Top             =   2400
            Width           =   1035
            _Version        =   65536
            _ExtentX        =   1826
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   192
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
         Begin EditLib.fpDoubleSingle ipp_CapPen_Cof_1 
            Height          =   315
            Left            =   7710
            TabIndex        =   13
            Top             =   420
            Width           =   1035
            _Version        =   196608
            _ExtentX        =   1826
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
         Begin EditLib.fpDoubleSingle ipp_IntPen_Cof_1 
            Height          =   315
            Left            =   8760
            TabIndex        =   14
            Top             =   420
            Width           =   1035
            _Version        =   196608
            _ExtentX        =   1826
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
         Begin EditLib.fpDoubleSingle ipp_CapPen_Cof_2 
            Height          =   315
            Left            =   7710
            TabIndex        =   16
            Top             =   750
            Width           =   1035
            _Version        =   196608
            _ExtentX        =   1826
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
         Begin EditLib.fpDoubleSingle ipp_IntPen_Cof_2 
            Height          =   315
            Left            =   8760
            TabIndex        =   17
            Top             =   750
            Width           =   1035
            _Version        =   196608
            _ExtentX        =   1826
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
         Begin EditLib.fpDoubleSingle ipp_CapPen_Cof_3 
            Height          =   315
            Left            =   7710
            TabIndex        =   19
            Top             =   1080
            Width           =   1035
            _Version        =   196608
            _ExtentX        =   1826
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
         Begin EditLib.fpDoubleSingle ipp_IntPen_Cof_3 
            Height          =   315
            Left            =   8760
            TabIndex        =   20
            Top             =   1080
            Width           =   1035
            _Version        =   196608
            _ExtentX        =   1826
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
         Begin EditLib.fpDoubleSingle ipp_CapPen_Cof_4 
            Height          =   315
            Left            =   7710
            TabIndex        =   22
            Top             =   1410
            Width           =   1035
            _Version        =   196608
            _ExtentX        =   1826
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
         Begin EditLib.fpDoubleSingle ipp_IntPen_Cof_4 
            Height          =   315
            Left            =   8760
            TabIndex        =   23
            Top             =   1410
            Width           =   1035
            _Version        =   196608
            _ExtentX        =   1826
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
         Begin EditLib.fpDoubleSingle ipp_CapPen_Cof_5 
            Height          =   315
            Left            =   7710
            TabIndex        =   25
            Top             =   1740
            Width           =   1035
            _Version        =   196608
            _ExtentX        =   1826
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
         Begin EditLib.fpDoubleSingle ipp_IntPen_Cof_5 
            Height          =   315
            Left            =   8760
            TabIndex        =   26
            Top             =   1740
            Width           =   1035
            _Version        =   196608
            _ExtentX        =   1826
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
         Begin EditLib.fpDoubleSingle ipp_CapPen_Cof_6 
            Height          =   315
            Left            =   7710
            TabIndex        =   28
            Top             =   2070
            Width           =   1035
            _Version        =   196608
            _ExtentX        =   1826
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
         Begin EditLib.fpDoubleSingle ipp_IntPen_Cof_6 
            Height          =   315
            Left            =   8760
            TabIndex        =   29
            Top             =   2070
            Width           =   1035
            _Version        =   196608
            _ExtentX        =   1826
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
         Begin Threed.SSPanel pnl_Ctr_CapPen_Cof 
            Height          =   315
            Left            =   7710
            TabIndex        =   81
            Top             =   2400
            Width           =   1035
            _Version        =   65536
            _ExtentX        =   1826
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   192
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
         Begin Threed.SSPanel pnl_Ctr_IntPen_Cof 
            Height          =   315
            Left            =   8760
            TabIndex        =   83
            Top             =   2400
            Width           =   1035
            _Version        =   65536
            _ExtentX        =   1826
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   192
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
         Begin EditLib.fpDoubleSingle ipp_ComPen_Cof_1 
            Height          =   315
            Left            =   9810
            TabIndex        =   15
            Top             =   420
            Width           =   1035
            _Version        =   196608
            _ExtentX        =   1826
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
         Begin EditLib.fpDoubleSingle ipp_ComPen_Cof_2 
            Height          =   315
            Left            =   9810
            TabIndex        =   18
            Top             =   750
            Width           =   1035
            _Version        =   196608
            _ExtentX        =   1826
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
         Begin EditLib.fpDoubleSingle ipp_ComPen_Cof_3 
            Height          =   315
            Left            =   9810
            TabIndex        =   21
            Top             =   1080
            Width           =   1035
            _Version        =   196608
            _ExtentX        =   1826
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
         Begin EditLib.fpDoubleSingle ipp_ComPen_Cof_4 
            Height          =   315
            Left            =   9810
            TabIndex        =   24
            Top             =   1410
            Width           =   1035
            _Version        =   196608
            _ExtentX        =   1826
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
         Begin EditLib.fpDoubleSingle ipp_ComPen_Cof_5 
            Height          =   315
            Left            =   9810
            TabIndex        =   27
            Top             =   1740
            Width           =   1035
            _Version        =   196608
            _ExtentX        =   1826
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
         Begin EditLib.fpDoubleSingle ipp_ComPen_Cof_6 
            Height          =   315
            Left            =   9810
            TabIndex        =   30
            Top             =   2070
            Width           =   1035
            _Version        =   196608
            _ExtentX        =   1826
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
         Begin Threed.SSPanel pnl_Ctr_ComPen_Cof 
            Height          =   315
            Left            =   9810
            TabIndex        =   84
            Top             =   2400
            Width           =   1035
            _Version        =   65536
            _ExtentX        =   1826
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   192
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
         Begin VB.Label lbl_SimMon 
            Alignment       =   1  'Right Justify
            Caption         =   "S/."
            Height          =   315
            Index           =   14
            Left            =   6930
            TabIndex        =   96
            Top             =   2070
            Width           =   645
         End
         Begin VB.Label lbl_SimMon 
            Alignment       =   1  'Right Justify
            Caption         =   "S/."
            Height          =   315
            Index           =   13
            Left            =   6930
            TabIndex        =   95
            Top             =   1740
            Width           =   645
         End
         Begin VB.Label lbl_SimMon 
            Alignment       =   1  'Right Justify
            Caption         =   "S/."
            Height          =   315
            Index           =   12
            Left            =   6930
            TabIndex        =   94
            Top             =   1410
            Width           =   645
         End
         Begin VB.Label lbl_SimMon 
            Alignment       =   1  'Right Justify
            Caption         =   "S/."
            Height          =   315
            Index           =   11
            Left            =   6930
            TabIndex        =   93
            Top             =   1080
            Width           =   645
         End
         Begin VB.Label lbl_SimMon 
            Alignment       =   1  'Right Justify
            Caption         =   "S/."
            Height          =   315
            Index           =   10
            Left            =   6930
            TabIndex        =   92
            Top             =   750
            Width           =   645
         End
         Begin VB.Label lbl_SimMon 
            Alignment       =   1  'Right Justify
            Caption         =   "S/."
            Height          =   315
            Index           =   9
            Left            =   6930
            TabIndex        =   91
            Top             =   420
            Width           =   645
         End
         Begin VB.Label lbl_SimMon 
            Alignment       =   1  'Right Justify
            Caption         =   "S/."
            Height          =   315
            Index           =   8
            Left            =   1860
            TabIndex        =   90
            Top             =   2070
            Width           =   645
         End
         Begin VB.Label lbl_SimMon 
            Alignment       =   1  'Right Justify
            Caption         =   "S/."
            Height          =   315
            Index           =   7
            Left            =   1860
            TabIndex        =   89
            Top             =   1740
            Width           =   645
         End
         Begin VB.Label lbl_SimMon 
            Alignment       =   1  'Right Justify
            Caption         =   "S/."
            Height          =   315
            Index           =   6
            Left            =   1860
            TabIndex        =   88
            Top             =   1410
            Width           =   645
         End
         Begin VB.Label lbl_SimMon 
            Alignment       =   1  'Right Justify
            Caption         =   "S/."
            Height          =   315
            Index           =   5
            Left            =   1860
            TabIndex        =   87
            Top             =   1110
            Width           =   645
         End
         Begin VB.Label lbl_SimMon 
            Alignment       =   1  'Right Justify
            Caption         =   "S/."
            Height          =   315
            Index           =   4
            Left            =   1860
            TabIndex        =   86
            Top             =   750
            Width           =   645
         End
         Begin VB.Label lbl_SimMon 
            Alignment       =   1  'Right Justify
            Caption         =   "S/."
            Height          =   315
            Index           =   3
            Left            =   1860
            TabIndex        =   85
            Top             =   420
            Width           =   645
         End
         Begin VB.Label lbl_SimMon 
            Alignment       =   1  'Right Justify
            Caption         =   "S/."
            Height          =   315
            Index           =   16
            Left            =   6930
            TabIndex        =   82
            Top             =   2400
            Width           =   645
         End
         Begin VB.Label Label34 
            Caption         =   "Total Aplicado:"
            Height          =   255
            Left            =   5250
            TabIndex        =   80
            Top             =   2400
            Width           =   1695
         End
         Begin VB.Label Label31 
            Caption         =   "Distribución de Penalidades PBP (Tramo Cofide/Mivivienda)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5250
            TabIndex        =   79
            Top             =   60
            Width           =   5115
         End
         Begin VB.Label lbl_NumCuo_Cof 
            Caption         =   "Cuota Nro. 012:"
            Height          =   255
            Index           =   5
            Left            =   5250
            TabIndex        =   78
            Top             =   2070
            Width           =   1695
         End
         Begin VB.Label lbl_NumCuo_Cof 
            Caption         =   "Cuota Nro. 011:"
            Height          =   255
            Index           =   4
            Left            =   5250
            TabIndex        =   77
            Top             =   1740
            Width           =   1695
         End
         Begin VB.Label lbl_NumCuo_Cof 
            Caption         =   "Cuota Nro. 010:"
            Height          =   255
            Index           =   3
            Left            =   5250
            TabIndex        =   76
            Top             =   1410
            Width           =   1695
         End
         Begin VB.Label lbl_NumCuo_Cof 
            Caption         =   "Cuota Nro. 009:"
            Height          =   255
            Index           =   2
            Left            =   5250
            TabIndex        =   75
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Label lbl_NumCuo_Cof 
            Caption         =   "Cuota Nro. 008:"
            Height          =   255
            Index           =   1
            Left            =   5250
            TabIndex        =   74
            Top             =   750
            Width           =   1695
         End
         Begin VB.Label lbl_NumCuo_Cof 
            Caption         =   "Cuota Nro. 007:"
            Height          =   255
            Index           =   0
            Left            =   5250
            TabIndex        =   73
            Top             =   420
            Width           =   1695
         End
         Begin VB.Label lbl_SimMon 
            Alignment       =   1  'Right Justify
            Caption         =   "S/."
            Height          =   315
            Index           =   15
            Left            =   1860
            TabIndex        =   71
            Top             =   2400
            Width           =   645
         End
         Begin VB.Label Label22 
            Caption         =   "Total Aplicado:"
            Height          =   255
            Left            =   60
            TabIndex        =   69
            Top             =   2400
            Width           =   1695
         End
         Begin VB.Label Label14 
            Caption         =   "Distribución de Penalidades PBP (Tramo Cliente)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   60
            TabIndex        =   68
            Top             =   60
            Width           =   4245
         End
         Begin VB.Label lbl_NumCuo_Cli 
            Caption         =   "Cuota Nro. 012:"
            Height          =   255
            Index           =   5
            Left            =   60
            TabIndex        =   67
            Top             =   2070
            Width           =   1695
         End
         Begin VB.Label lbl_NumCuo_Cli 
            Caption         =   "Cuota Nro. 011:"
            Height          =   255
            Index           =   4
            Left            =   60
            TabIndex        =   66
            Top             =   1740
            Width           =   1695
         End
         Begin VB.Label lbl_NumCuo_Cli 
            Caption         =   "Cuota Nro. 010:"
            Height          =   255
            Index           =   3
            Left            =   60
            TabIndex        =   65
            Top             =   1410
            Width           =   1695
         End
         Begin VB.Label lbl_NumCuo_Cli 
            Caption         =   "Cuota Nro. 009:"
            Height          =   255
            Index           =   2
            Left            =   60
            TabIndex        =   64
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Label lbl_NumCuo_Cli 
            Caption         =   "Cuota Nro. 008:"
            Height          =   255
            Index           =   1
            Left            =   60
            TabIndex        =   63
            Top             =   750
            Width           =   1695
         End
         Begin VB.Label lbl_NumCuo_Cli 
            Caption         =   "Cuota Nro. 007:"
            Height          =   255
            Index           =   0
            Left            =   60
            TabIndex        =   62
            Top             =   420
            Width           =   1695
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   1995
         Left            =   30
         TabIndex        =   36
         Top             =   2910
         Width           =   10905
         _Version        =   65536
         _ExtentX        =   19235
         _ExtentY        =   3519
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
         Begin Threed.SSPanel SSPanel9 
            Height          =   285
            Left            =   60
            TabIndex        =   41
            Top             =   60
            Width           =   1755
            _Version        =   65536
            _ExtentX        =   3096
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro. Cuota"
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
         Begin Threed.SSPanel SSPanel3 
            Height          =   285
            Left            =   1800
            TabIndex        =   42
            Top             =   60
            Width           =   2205
            _Version        =   65536
            _ExtentX        =   3889
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Vcto."
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
            Left            =   3990
            TabIndex        =   43
            Top             =   60
            Width           =   2205
            _Version        =   65536
            _ExtentX        =   3889
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Pago"
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
            Left            =   6180
            TabIndex        =   44
            Top             =   60
            Width           =   1725
            _Version        =   65536
            _ExtentX        =   3043
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Días Atraso"
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
         Begin Threed.SSPanel SSPanel12 
            Height          =   285
            Left            =   7890
            TabIndex        =   45
            Top             =   60
            Width           =   2895
            _Version        =   65536
            _ExtentX        =   5106
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Situación"
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
         Begin MSFlexGridLib.MSFlexGrid grd_LisCuo 
            Height          =   1605
            Left            =   30
            TabIndex        =   34
            Top             =   360
            Width           =   10815
            _ExtentX        =   19076
            _ExtentY        =   2831
            _Version        =   393216
            Rows            =   6
            Cols            =   5
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   37
         Top             =   30
         Width           =   10905
         _Version        =   65536
         _ExtentX        =   19235
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
            Height          =   555
            Left            =   570
            TabIndex        =   38
            Top             =   30
            Width           =   4605
            _Version        =   65536
            _ExtentX        =   8123
            _ExtentY        =   979
            _StockProps     =   15
            Caption         =   "Evaluación y Asignación de Premio Buen Pagador"
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
            Picture         =   "OpeTra_frm_294.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   39
         Top             =   750
         Width           =   10905
         _Version        =   65536
         _ExtentX        =   19235
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
            Picture         =   "OpeTra_frm_294.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   31
            ToolTipText     =   "Regenerar Propuesta de Asignación PBP"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10290
            Picture         =   "OpeTra_frm_294.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   32
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   1425
         Left            =   30
         TabIndex        =   40
         Top             =   1440
         Width           =   10905
         _Version        =   65536
         _ExtentX        =   19235
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
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   1335
            Left            =   60
            TabIndex        =   33
            Top             =   60
            Width           =   10815
            _ExtentX        =   19076
            _ExtentY        =   2355
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
   End
End
Attribute VB_Name = "frm_Pro_EvaPBP_04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_int_CEvIni     As Integer
Dim l_int_CEvFin     As Integer
Dim l_int_IniPen     As Integer
Dim l_int_FinPen     As Integer
Dim l_dbl_SalAct     As Double
Dim l_dbl_SalNue     As Double

Private Sub cmb_FlgPBP_Click()
   If cmb_FlgPBP.ListIndex > -1 Then
      If cmb_FlgPBP.ItemData(cmb_FlgPBP.ListIndex) = 2 Then
         pnl_Pen_CuoIni.Caption = CStr(l_int_IniPen)
         pnl_Pen_CuoFin.Caption = CStr(l_int_FinPen)
      
         'Proponiendo Valores
         ipp_CapPen_Cli_1.Value = Format(CDbl(pnl_Tot_Capita_Cli.Caption) / 6, "###,##0.00")
         ipp_CapPen_Cli_2.Value = Format(CDbl(pnl_Tot_Capita_Cli.Caption) / 6, "###,##0.00")
         ipp_CapPen_Cli_3.Value = Format(CDbl(pnl_Tot_Capita_Cli.Caption) / 6, "###,##0.00")
         ipp_CapPen_Cli_4.Value = Format(CDbl(pnl_Tot_Capita_Cli.Caption) / 6, "###,##0.00")
         ipp_CapPen_Cli_5.Value = Format(CDbl(pnl_Tot_Capita_Cli.Caption) / 6, "###,##0.00")
         ipp_CapPen_Cli_6.Value = Format(CDbl(pnl_Tot_Capita_Cli.Caption) / 6, "###,##0.00")
         
         ipp_IntPen_Cli_1.Value = Format(CDbl(pnl_Tot_Intere_Cli.Caption) / 6, "###,##0.00")
         ipp_IntPen_Cli_2.Value = Format(CDbl(pnl_Tot_Intere_Cli.Caption) / 6, "###,##0.00")
         ipp_IntPen_Cli_3.Value = Format(CDbl(pnl_Tot_Intere_Cli.Caption) / 6, "###,##0.00")
         ipp_IntPen_Cli_4.Value = Format(CDbl(pnl_Tot_Intere_Cli.Caption) / 6, "###,##0.00")
         ipp_IntPen_Cli_5.Value = Format(CDbl(pnl_Tot_Intere_Cli.Caption) / 6, "###,##0.00")
         ipp_IntPen_Cli_6.Value = Format(CDbl(pnl_Tot_Intere_Cli.Caption) / 6, "###,##0.00")
         
         ipp_CapPen_Cof_1.Value = Format(CDbl(pnl_Tot_Capita_Cof.Caption) / 6, "###,##0.00")
         ipp_CapPen_Cof_2.Value = Format(CDbl(pnl_Tot_Capita_Cof.Caption) / 6, "###,##0.00")
         ipp_CapPen_Cof_3.Value = Format(CDbl(pnl_Tot_Capita_Cof.Caption) / 6, "###,##0.00")
         ipp_CapPen_Cof_4.Value = Format(CDbl(pnl_Tot_Capita_Cof.Caption) / 6, "###,##0.00")
         ipp_CapPen_Cof_5.Value = Format(CDbl(pnl_Tot_Capita_Cof.Caption) / 6, "###,##0.00")
         ipp_CapPen_Cof_6.Value = Format(CDbl(pnl_Tot_Capita_Cof.Caption) / 6, "###,##0.00")
         
         ipp_IntPen_Cof_1.Value = Format(CDbl(pnl_Tot_Intere_Cof.Caption) / 6, "###,##0.00")
         ipp_IntPen_Cof_2.Value = Format(CDbl(pnl_Tot_Intere_Cof.Caption) / 6, "###,##0.00")
         ipp_IntPen_Cof_3.Value = Format(CDbl(pnl_Tot_Intere_Cof.Caption) / 6, "###,##0.00")
         ipp_IntPen_Cof_4.Value = Format(CDbl(pnl_Tot_Intere_Cof.Caption) / 6, "###,##0.00")
         ipp_IntPen_Cof_5.Value = Format(CDbl(pnl_Tot_Intere_Cof.Caption) / 6, "###,##0.00")
         ipp_IntPen_Cof_6.Value = Format(CDbl(pnl_Tot_Intere_Cof.Caption) / 6, "###,##0.00")
      
         ipp_ComPen_Cof_1.Value = Format(CDbl(pnl_Tot_Comisi_Cof.Caption) / 6, "###,##0.00")
         ipp_ComPen_Cof_2.Value = Format(CDbl(pnl_Tot_Comisi_Cof.Caption) / 6, "###,##0.00")
         ipp_ComPen_Cof_3.Value = Format(CDbl(pnl_Tot_Comisi_Cof.Caption) / 6, "###,##0.00")
         ipp_ComPen_Cof_4.Value = Format(CDbl(pnl_Tot_Comisi_Cof.Caption) / 6, "###,##0.00")
         ipp_ComPen_Cof_5.Value = Format(CDbl(pnl_Tot_Comisi_Cof.Caption) / 6, "###,##0.00")
         ipp_ComPen_Cof_6.Value = Format(CDbl(pnl_Tot_Comisi_Cof.Caption) / 6, "###,##0.00")
         
         ipp_CapPen_Cli_1.Enabled = True
         ipp_CapPen_Cli_2.Enabled = True
         ipp_CapPen_Cli_3.Enabled = True
         ipp_CapPen_Cli_4.Enabled = True
         ipp_CapPen_Cli_5.Enabled = True
         ipp_CapPen_Cli_6.Enabled = True
      
         ipp_IntPen_Cli_1.Enabled = True
         ipp_IntPen_Cli_2.Enabled = True
         ipp_IntPen_Cli_3.Enabled = True
         ipp_IntPen_Cli_4.Enabled = True
         ipp_IntPen_Cli_5.Enabled = True
         ipp_IntPen_Cli_6.Enabled = True
         
         ipp_CapPen_Cof_1.Enabled = True
         ipp_CapPen_Cof_2.Enabled = True
         ipp_CapPen_Cof_3.Enabled = True
         ipp_CapPen_Cof_4.Enabled = True
         ipp_CapPen_Cof_5.Enabled = True
         ipp_CapPen_Cof_6.Enabled = True

         ipp_IntPen_Cof_1.Enabled = True
         ipp_IntPen_Cof_2.Enabled = True
         ipp_IntPen_Cof_3.Enabled = True
         ipp_IntPen_Cof_4.Enabled = True
         ipp_IntPen_Cof_5.Enabled = True
         ipp_IntPen_Cof_6.Enabled = True
         
         ipp_ComPen_Cof_1.Enabled = True
         ipp_ComPen_Cof_2.Enabled = True
         ipp_ComPen_Cof_3.Enabled = True
         ipp_ComPen_Cof_4.Enabled = True
         ipp_ComPen_Cof_5.Enabled = True
         ipp_ComPen_Cof_6.Enabled = True
         
         Call gs_SetFocus(ipp_CapPen_Cli_1)
      Else
         ipp_CapPen_Cli_1.Value = 0
         ipp_CapPen_Cli_2.Value = 0
         ipp_CapPen_Cli_3.Value = 0
         ipp_CapPen_Cli_4.Value = 0
         ipp_CapPen_Cli_5.Value = 0
         ipp_CapPen_Cli_6.Value = 0
         
         ipp_IntPen_Cli_1.Value = 0
         ipp_IntPen_Cli_2.Value = 0
         ipp_IntPen_Cli_3.Value = 0
         ipp_IntPen_Cli_4.Value = 0
         ipp_IntPen_Cli_5.Value = 0
         ipp_IntPen_Cli_6.Value = 0
         
         ipp_CapPen_Cli_1.Enabled = False
         ipp_CapPen_Cli_2.Enabled = False
         ipp_CapPen_Cli_3.Enabled = False
         ipp_CapPen_Cli_4.Enabled = False
         ipp_CapPen_Cli_5.Enabled = False
         ipp_CapPen_Cli_6.Enabled = False
      
         ipp_IntPen_Cli_1.Enabled = False
         ipp_IntPen_Cli_2.Enabled = False
         ipp_IntPen_Cli_3.Enabled = False
         ipp_IntPen_Cli_4.Enabled = False
         ipp_IntPen_Cli_5.Enabled = False
         ipp_IntPen_Cli_6.Enabled = False
                  
         ipp_CapPen_Cof_1.Value = 0
         ipp_CapPen_Cof_2.Value = 0
         ipp_CapPen_Cof_3.Value = 0
         ipp_CapPen_Cof_4.Value = 0
         ipp_CapPen_Cof_5.Value = 0
         ipp_CapPen_Cof_6.Value = 0
         
         ipp_IntPen_Cof_1.Value = 0
         ipp_IntPen_Cof_2.Value = 0
         ipp_IntPen_Cof_3.Value = 0
         ipp_IntPen_Cof_4.Value = 0
         ipp_IntPen_Cof_5.Value = 0
         ipp_IntPen_Cof_6.Value = 0
         
         ipp_ComPen_Cof_1.Value = 0
         ipp_ComPen_Cof_2.Value = 0
         ipp_ComPen_Cof_3.Value = 0
         ipp_ComPen_Cof_4.Value = 0
         ipp_ComPen_Cof_5.Value = 0
         ipp_ComPen_Cof_6.Value = 0
         
         ipp_CapPen_Cof_1.Enabled = False
         ipp_CapPen_Cof_2.Enabled = False
         ipp_CapPen_Cof_3.Enabled = False
         ipp_CapPen_Cof_4.Enabled = False
         ipp_CapPen_Cof_5.Enabled = False
         ipp_CapPen_Cof_6.Enabled = False

         ipp_IntPen_Cof_1.Enabled = False
         ipp_IntPen_Cof_2.Enabled = False
         ipp_IntPen_Cof_3.Enabled = False
         ipp_IntPen_Cof_4.Enabled = False
         ipp_IntPen_Cof_5.Enabled = False
         ipp_IntPen_Cof_6.Enabled = False
         
         ipp_ComPen_Cof_1.Enabled = False
         ipp_ComPen_Cof_2.Enabled = False
         ipp_ComPen_Cof_3.Enabled = False
         ipp_ComPen_Cof_4.Enabled = False
         ipp_ComPen_Cof_5.Enabled = False
         ipp_ComPen_Cof_6.Enabled = False
         
         Call gs_SetFocus(cmd_Grabar)
      End If
   End If
End Sub

Private Sub cmb_FlgPBP_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_FlgPBP_Click
   End If
End Sub

Private Sub cmd_Grabar_Click()
   Dim r_int_EvaAsg     As Integer
   Dim r_int_EvaPer     As Integer
   Dim r_int_EvaPen     As Integer
   
   If cmb_FlgPBP.ListIndex = -1 Then
      MsgBox "Debe seleccionar si se Asigna PBP.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_FlgPBP)
      Exit Sub
   End If
   
   If cmb_FlgPBP.ItemData(cmb_FlgPBP.ListIndex) = 2 Then
      'Cuota 01 (Cliente)
      If CDbl(ipp_CapPen_Cli_1.Value) > 0 Then
         If ff_ValidaCuota(moddat_g_str_NumOpe, l_int_IniPen) Then
            MsgBox "No se puede aplicar Penalidades porque la Cuota ya ha sido pagada.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_CapPen_Cli_1)
            Exit Sub
         End If
      Else
         If Not ff_ValidaCuota(moddat_g_str_NumOpe, l_int_IniPen) Then
            MsgBox "Debe ingresar la fracción de Capital a Penalizar.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_CapPen_Cli_1)
            Exit Sub
         End If
      End If
   
      If CDbl(ipp_IntPen_Cli_1.Value) > 0 Then
         If ff_ValidaCuota(moddat_g_str_NumOpe, l_int_IniPen) Then
            MsgBox "No se puede aplicar Penalidades porque la Cuota ya ha sido pagada.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_IntPen_Cli_1)
            Exit Sub
         End If
      Else
         If Not ff_ValidaCuota(moddat_g_str_NumOpe, l_int_IniPen) Then
            MsgBox "Debe ingresar la fracción de Interés a Penalizar.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_IntPen_Cli_1)
            Exit Sub
         End If
      End If
   
      'Cuota 02 (Cliente)
      If CDbl(ipp_CapPen_Cli_2.Value) > 0 Then
         If ff_ValidaCuota(moddat_g_str_NumOpe, l_int_IniPen + 1) Then
            MsgBox "No se puede aplicar Penalidades porque la Cuota ya ha sido pagada.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_CapPen_Cli_2)
            Exit Sub
         End If
      Else
         If Not ff_ValidaCuota(moddat_g_str_NumOpe, l_int_IniPen + 1) Then
            MsgBox "Debe ingresar la fracción de Capital a Penalizar.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_CapPen_Cli_2)
            Exit Sub
         End If
      End If
   
      If CDbl(ipp_IntPen_Cli_2.Value) > 0 Then
         If ff_ValidaCuota(moddat_g_str_NumOpe, l_int_IniPen + 1) Then
            MsgBox "No se puede aplicar Penalidades porque la Cuota ya ha sido pagada.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_IntPen_Cli_2)
            Exit Sub
         End If
      Else
         If Not ff_ValidaCuota(moddat_g_str_NumOpe, l_int_IniPen + 1) Then
            MsgBox "Debe ingresar la fracción de Interés a Penalizar.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_IntPen_Cli_2)
            Exit Sub
         End If
      End If
   
      'Cuota 03 (Cliente)
      If CDbl(ipp_CapPen_Cli_3.Value) > 0 Then
         If ff_ValidaCuota(moddat_g_str_NumOpe, l_int_IniPen + 2) Then
            MsgBox "No se puede aplicar Penalidades porque la Cuota ya ha sido pagada.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_CapPen_Cli_3)
            Exit Sub
         End If
      Else
         If Not ff_ValidaCuota(moddat_g_str_NumOpe, l_int_IniPen + 2) Then
            MsgBox "Debe ingresar la fracción de Capital a Penalizar.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_CapPen_Cli_3)
            Exit Sub
         End If
      End If
   
      If CDbl(ipp_IntPen_Cli_3.Value) > 0 Then
         If ff_ValidaCuota(moddat_g_str_NumOpe, l_int_IniPen + 2) Then
            MsgBox "No se puede aplicar Penalidades porque la Cuota ya ha sido pagada.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_IntPen_Cli_3)
            Exit Sub
         End If
      Else
         If Not ff_ValidaCuota(moddat_g_str_NumOpe, l_int_IniPen + 2) Then
            MsgBox "Debe ingresar la fracción de Interés a Penalizar.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_IntPen_Cli_3)
            Exit Sub
         End If
      End If
   
      'Cuota 04 (Cliente)
      If CDbl(ipp_CapPen_Cli_4.Value) > 0 Then
         If ff_ValidaCuota(moddat_g_str_NumOpe, l_int_IniPen + 3) Then
            MsgBox "No se puede aplicar Penalidades porque la Cuota ya ha sido pagada.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_CapPen_Cli_4)
            Exit Sub
         End If
      Else
         If Not ff_ValidaCuota(moddat_g_str_NumOpe, l_int_IniPen + 3) Then
            MsgBox "Debe ingresar la fracción de Capital a Penalizar.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_CapPen_Cli_4)
            Exit Sub
         End If
      End If
   
      If CDbl(ipp_IntPen_Cli_4.Value) > 0 Then
         If ff_ValidaCuota(moddat_g_str_NumOpe, l_int_IniPen + 3) Then
            MsgBox "No se puede aplicar Penalidades porque la Cuota ya ha sido pagada.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_IntPen_Cli_4)
            Exit Sub
         End If
      Else
         If Not ff_ValidaCuota(moddat_g_str_NumOpe, l_int_IniPen + 3) Then
            MsgBox "Debe ingresar la fracción de Interés a Penalizar.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_IntPen_Cli_4)
            Exit Sub
         End If
      End If
   
      'Cuota 05 (Cliente)
      If CDbl(ipp_CapPen_Cli_5.Value) > 0 Then
         If ff_ValidaCuota(moddat_g_str_NumOpe, l_int_IniPen + 4) Then
            MsgBox "No se puede aplicar Penalidades porque la Cuota ya ha sido pagada.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_CapPen_Cli_5)
            Exit Sub
         End If
      Else
         If Not ff_ValidaCuota(moddat_g_str_NumOpe, l_int_IniPen + 4) Then
            MsgBox "Debe ingresar la fracción de Capital a Penalizar.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_CapPen_Cli_5)
            Exit Sub
         End If
      End If
   
      If CDbl(ipp_IntPen_Cli_5.Value) > 0 Then
         If ff_ValidaCuota(moddat_g_str_NumOpe, l_int_IniPen + 4) Then
            MsgBox "No se puede aplicar Penalidades porque la Cuota ya ha sido pagada.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_IntPen_Cli_5)
            Exit Sub
         End If
      Else
         If Not ff_ValidaCuota(moddat_g_str_NumOpe, l_int_IniPen + 4) Then
            MsgBox "Debe ingresar la fracción de Interés a Penalizar.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_IntPen_Cli_5)
            Exit Sub
         End If
      End If
   
      'Cuota 06 (Cliente)
      If CDbl(ipp_CapPen_Cli_6.Value) > 0 Then
         If ff_ValidaCuota(moddat_g_str_NumOpe, l_int_IniPen + 5) Then
            MsgBox "No se puede aplicar Penalidades porque la Cuota ya ha sido pagada.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_CapPen_Cli_6)
            Exit Sub
         End If
      Else
         If Not ff_ValidaCuota(moddat_g_str_NumOpe, l_int_IniPen + 5) Then
            MsgBox "Debe ingresar la fracción de Capital a Penalizar.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_CapPen_Cli_6)
            Exit Sub
         End If
      End If
   
      If CDbl(ipp_IntPen_Cli_6.Value) > 0 Then
         If ff_ValidaCuota(moddat_g_str_NumOpe, l_int_IniPen + 5) Then
            MsgBox "No se puede aplicar Penalidades porque la Cuota ya ha sido pagada.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_IntPen_Cli_6)
            Exit Sub
         End If
      Else
         If Not ff_ValidaCuota(moddat_g_str_NumOpe, l_int_IniPen + 5) Then
            MsgBox "Debe ingresar la fracción de Interés a Penalizar.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_IntPen_Cli_6)
            Exit Sub
         End If
      End If
   
      'Validando Totales de Control
      If CDbl(pnl_Ctr_CapPen_Cli.Caption) <> CDbl(pnl_Tot_Capita_Cli.Caption) Then
         MsgBox "La distribución de Capital Penalizado en el Tramo de Cliente no cuadra con el Total a Penalizar.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_CapPen_Cli_6)
         Exit Sub
      End If
   
      If CDbl(pnl_Ctr_IntPen_Cli.Caption) <> CDbl(pnl_Tot_Intere_Cli.Caption) Then
         MsgBox "La distribución de Interés Penalizado en el Tramo de Cliente no cuadra con el Total a Penalizar.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_IntPen_Cli_6)
         Exit Sub
      End If
      
      If InStr(moddat_g_str_Agr2MIC, moddat_g_str_CodPrd) = 0 Then 'moddat_g_str_CodPrd <> "006" Then
         'Cuota 01 (Cofide/Mivivienda)
         If CDbl(ipp_CapPen_Cof_1.Value) = 0 Then
            MsgBox "Debe ingresar la fracción de Capital a Penalizar.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_CapPen_Cof_1)
            Exit Sub
         End If
      
         If CDbl(ipp_IntPen_Cof_1.Value) = 0 Then
            MsgBox "Debe ingresar la fracción de Interés a Penalizar.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_IntPen_Cof_1)
            Exit Sub
         End If
         
         If InStr(moddat_g_str_AgrCRC, moddat_g_str_CodPrd) = 0 And InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) = 0 Then 'moddat_g_str_CodPrd <> "001" And moddat_g_str_CodPrd <> "003" Then
            If CDbl(ipp_ComPen_Cof_1.Value) = 0 Then
               MsgBox "Debe ingresar la fracción de Comisión a Penalizar.", vbExclamation, modgen_g_str_NomPlt
               Call gs_SetFocus(ipp_ComPen_Cof_1)
               Exit Sub
            End If
         End If
      
         'Cuota 02 (Cofide/Mivivienda)
         If CDbl(ipp_CapPen_Cof_2.Value) = 0 Then
            MsgBox "Debe ingresar la fracción de Capital a Penalizar.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_CapPen_Cof_2)
            Exit Sub
         End If
      
         If CDbl(ipp_IntPen_Cof_2.Value) = 0 Then
            MsgBox "Debe ingresar la fracción de Interés a Penalizar.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_IntPen_Cof_2)
            Exit Sub
         End If
         
         If InStr(moddat_g_str_AgrCRC, moddat_g_str_CodPrd) = 0 And InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) = 0 Then 'moddat_g_str_CodPrd <> "001" And moddat_g_str_CodPrd <> "003" Then
            If CDbl(ipp_ComPen_Cof_2.Value) = 0 Then
               MsgBox "Debe ingresar la fracción de Comisión a Penalizar.", vbExclamation, modgen_g_str_NomPlt
               Call gs_SetFocus(ipp_ComPen_Cof_2)
               Exit Sub
            End If
         End If
      
         'Cuota 03 (Cofide/Mivivienda)
         If CDbl(ipp_CapPen_Cof_3.Value) = 0 Then
            MsgBox "Debe ingresar la fracción de Capital a Penalizar.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_CapPen_Cof_3)
            Exit Sub
         End If
      
         If CDbl(ipp_IntPen_Cof_3.Value) = 0 Then
            MsgBox "Debe ingresar la fracción de Interés a Penalizar.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_IntPen_Cof_3)
            Exit Sub
         End If
         
         If InStr(moddat_g_str_AgrCRC, moddat_g_str_CodPrd) = 0 And InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) = 0 Then 'moddat_g_str_CodPrd <> "001" And moddat_g_str_CodPrd <> "003" Then
            If CDbl(ipp_ComPen_Cof_3.Value) = 0 Then
               MsgBox "Debe ingresar la fracción de Comisión a Penalizar.", vbExclamation, modgen_g_str_NomPlt
               Call gs_SetFocus(ipp_ComPen_Cof_3)
               Exit Sub
            End If
         End If
      
         'Cuota 04 (Cofide/Mivivienda)
         If CDbl(ipp_CapPen_Cof_4.Value) = 0 Then
            MsgBox "Debe ingresar la fracción de Capital a Penalizar.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_CapPen_Cof_4)
            Exit Sub
         End If
      
         If CDbl(ipp_IntPen_Cof_4.Value) = 0 Then
            MsgBox "Debe ingresar la fracción de Interés a Penalizar.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_IntPen_Cof_4)
            Exit Sub
         End If
         
         If InStr(moddat_g_str_AgrCRC, moddat_g_str_CodPrd) = 0 And InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) = 0 Then 'moddat_g_str_CodPrd <> "001" And moddat_g_str_CodPrd <> "003" Then
            If CDbl(ipp_ComPen_Cof_4.Value) = 0 Then
               MsgBox "Debe ingresar la fracción de Comisión a Penalizar.", vbExclamation, modgen_g_str_NomPlt
               Call gs_SetFocus(ipp_ComPen_Cof_4)
               Exit Sub
            End If
         End If
      
         'Cuota 05 (Cofide/Mivivienda)
         If CDbl(ipp_CapPen_Cof_5.Value) = 0 Then
            MsgBox "Debe ingresar la fracción de Capital a Penalizar.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_CapPen_Cof_5)
            Exit Sub
         End If
      
         If CDbl(ipp_IntPen_Cof_5.Value) = 0 Then
            MsgBox "Debe ingresar la fracción de Interés a Penalizar.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_IntPen_Cof_5)
            Exit Sub
         End If
         
         If InStr(moddat_g_str_AgrCRC, moddat_g_str_CodPrd) = 0 And InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) = 0 Then 'moddat_g_str_CodPrd <> "001" And moddat_g_str_CodPrd <> "003" Then
            If CDbl(ipp_ComPen_Cof_5.Value) = 0 Then
               MsgBox "Debe ingresar la fracción de Comisión a Penalizar.", vbExclamation, modgen_g_str_NomPlt
               Call gs_SetFocus(ipp_ComPen_Cof_5)
               Exit Sub
            End If
         End If
      
         'Cuota 06 (Cofide/Mivivienda)
         If CDbl(ipp_CapPen_Cof_6.Value) = 0 Then
            MsgBox "Debe ingresar la fracción de Capital a Penalizar.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_CapPen_Cof_6)
            Exit Sub
         End If
      
         If CDbl(ipp_IntPen_Cof_6.Value) = 0 Then
            MsgBox "Debe ingresar la fracción de Interés a Penalizar.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_IntPen_Cof_6)
            Exit Sub
         End If
         
         If InStr(moddat_g_str_AgrCRC, moddat_g_str_CodPrd) = 0 And InStr(moddat_g_str_AgrCRC, moddat_g_str_CodPrd) = 0 Then 'moddat_g_str_CodPrd <> "001" And moddat_g_str_CodPrd <> "003" Then
            If CDbl(ipp_ComPen_Cof_6.Value) = 0 Then
               MsgBox "Debe ingresar la fracción de Comisión a Penalizar.", vbExclamation, modgen_g_str_NomPlt
               Call gs_SetFocus(ipp_ComPen_Cof_6)
               Exit Sub
            End If
         End If
      
         'Totales de Control
         If CDbl(pnl_Ctr_CapPen_Cof.Caption) <> CDbl(pnl_Tot_Capita_Cof.Caption) Then
            MsgBox "La distribución de Capital Penalizado en el Tramo de COFIDE/Mivivienda no cuadra con el Total a Penalizar.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_CapPen_Cof_6)
            Exit Sub
         End If
      
         If CDbl(pnl_Ctr_IntPen_Cof.Caption) <> CDbl(pnl_Tot_Intere_Cof.Caption) Then
            MsgBox "La distribución de Interés Penalizado en el Tramo de COFIDE/Mivivienda no cuadra con el Total a Penalizar.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_IntPen_Cof_6)
            Exit Sub
         End If
         
         If CDbl(pnl_Ctr_ComPen_Cof.Caption) <> CDbl(pnl_Tot_Comisi_Cof.Caption) Then
            MsgBox "La distribución de Comisión Penalizado en el Tramo de COFIDE/Mivivienda no cuadra con el Total a Penalizar.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(ipp_ComPen_Cof_6)
            Exit Sub
         End If
      End If
   End If

   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If

   'Grabando en CRE_DETPBP
   g_str_Parame = "USP_CRE_DETPBP ("
   g_str_Parame = g_str_Parame & moddat_g_str_Codigo & ", "
   g_str_Parame = g_str_Parame & moddat_g_str_CodIte & ", "
   g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumOpe & "', "
   
   g_str_Parame = g_str_Parame & CStr(cmb_FlgPBP.ItemData(cmb_FlgPBP.ListIndex)) & ", "
   g_str_Parame = g_str_Parame & CStr(CInt(pnl_CuoCon.Caption)) & ", "
   
   g_str_Parame = g_str_Parame & CStr(CDbl(pnl_Tot_Capita_Cli.Caption)) & ", "
   g_str_Parame = g_str_Parame & CStr(CDbl(pnl_Tot_Intere_Cli.Caption)) & ", "
   g_str_Parame = g_str_Parame & CStr(CDbl(pnl_Tot_Capita_Cof.Caption)) & ", "
   g_str_Parame = g_str_Parame & CStr(CDbl(pnl_Tot_Intere_Cof.Caption)) & ", "
   g_str_Parame = g_str_Parame & CStr(CDbl(pnl_Tot_Comisi_Cof.Caption)) & ", "
   g_str_Parame = g_str_Parame & CStr(l_dbl_SalAct) & ", "
   g_str_Parame = g_str_Parame & CStr(l_dbl_SalNue) & ", "
   g_str_Parame = g_str_Parame & CStr(l_int_CEvIni) & ", "
   g_str_Parame = g_str_Parame & CStr(l_int_CEvFin) & ", "
   
   g_str_Parame = g_str_Parame & CStr(l_int_IniPen) & ", "
   g_str_Parame = g_str_Parame & CStr(l_int_FinPen) & ", "
   
   g_str_Parame = g_str_Parame & CStr(CDbl(ipp_CapPen_Cli_1.Value)) & ", "
   g_str_Parame = g_str_Parame & CStr(CDbl(ipp_IntPen_Cli_1.Value)) & ", "
   g_str_Parame = g_str_Parame & CStr(CDbl(ipp_CapPen_Cli_2.Value)) & ", "
   g_str_Parame = g_str_Parame & CStr(CDbl(ipp_IntPen_Cli_2.Value)) & ", "
   g_str_Parame = g_str_Parame & CStr(CDbl(ipp_CapPen_Cli_3.Value)) & ", "
   g_str_Parame = g_str_Parame & CStr(CDbl(ipp_IntPen_Cli_3.Value)) & ", "
   g_str_Parame = g_str_Parame & CStr(CDbl(ipp_CapPen_Cli_4.Value)) & ", "
   g_str_Parame = g_str_Parame & CStr(CDbl(ipp_IntPen_Cli_4.Value)) & ", "
   g_str_Parame = g_str_Parame & CStr(CDbl(ipp_CapPen_Cli_5.Value)) & ", "
   g_str_Parame = g_str_Parame & CStr(CDbl(ipp_IntPen_Cli_5.Value)) & ", "
   g_str_Parame = g_str_Parame & CStr(CDbl(ipp_CapPen_Cli_6.Value)) & ", "
   g_str_Parame = g_str_Parame & CStr(CDbl(ipp_IntPen_Cli_6.Value)) & ", "
   
   g_str_Parame = g_str_Parame & CStr(l_int_IniPen) & ", "
   g_str_Parame = g_str_Parame & CStr(l_int_FinPen) & ", "
   
   g_str_Parame = g_str_Parame & CStr(CDbl(ipp_CapPen_Cof_1.Value)) & ", "
   g_str_Parame = g_str_Parame & CStr(CDbl(ipp_IntPen_Cof_1.Value)) & ", "
   g_str_Parame = g_str_Parame & CStr(CDbl(ipp_ComPen_Cof_1.Value)) & ", "
   g_str_Parame = g_str_Parame & CStr(CDbl(ipp_CapPen_Cof_2.Value)) & ", "
   g_str_Parame = g_str_Parame & CStr(CDbl(ipp_IntPen_Cof_2.Value)) & ", "
   g_str_Parame = g_str_Parame & CStr(CDbl(ipp_ComPen_Cof_2.Value)) & ", "
   g_str_Parame = g_str_Parame & CStr(CDbl(ipp_CapPen_Cof_3.Value)) & ", "
   g_str_Parame = g_str_Parame & CStr(CDbl(ipp_IntPen_Cof_3.Value)) & ", "
   g_str_Parame = g_str_Parame & CStr(CDbl(ipp_ComPen_Cof_3.Value)) & ", "
   g_str_Parame = g_str_Parame & CStr(CDbl(ipp_CapPen_Cof_4.Value)) & ", "
   g_str_Parame = g_str_Parame & CStr(CDbl(ipp_IntPen_Cof_4.Value)) & ", "
   g_str_Parame = g_str_Parame & CStr(CDbl(ipp_ComPen_Cof_4.Value)) & ", "
   g_str_Parame = g_str_Parame & CStr(CDbl(ipp_CapPen_Cof_5.Value)) & ", "
   g_str_Parame = g_str_Parame & CStr(CDbl(ipp_IntPen_Cof_5.Value)) & ", "
   g_str_Parame = g_str_Parame & CStr(CDbl(ipp_ComPen_Cof_5.Value)) & ", "
   g_str_Parame = g_str_Parame & CStr(CDbl(ipp_CapPen_Cof_6.Value)) & ", "
   g_str_Parame = g_str_Parame & CStr(CDbl(ipp_IntPen_Cof_6.Value)) & ", "
   g_str_Parame = g_str_Parame & CStr(CDbl(ipp_ComPen_Cof_6.Value)) & ", "
   
   'Datos de Auditoria
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
   g_str_Parame = g_str_Parame & "2)"
         
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
      MsgBox "No se pudo ejecutar el procedimiento USP_CRE_DETPBP.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If

   'Obteniendo Totales
   r_int_EvaAsg = modmip_gf_TotalPBP(CInt(moddat_g_str_Codigo), CInt(moddat_g_str_CodIte), 1)
   r_int_EvaPer = modmip_gf_TotalPBP(CInt(moddat_g_str_Codigo), CInt(moddat_g_str_CodIte), 2)
   r_int_EvaPen = modmip_gf_TotalPBP(CInt(moddat_g_str_Codigo), CInt(moddat_g_str_CodIte), 3)
   
   'Grabando Cabecera
   g_str_Parame = "USP_CRE_CABPBP ("
   g_str_Parame = g_str_Parame & moddat_g_str_Codigo & ", "
   g_str_Parame = g_str_Parame & moddat_g_str_CodIte & ", "
   
   g_str_Parame = g_str_Parame & CStr(r_int_EvaAsg + r_int_EvaPer + r_int_EvaPen) & ", "
   g_str_Parame = g_str_Parame & CStr(r_int_EvaAsg) & ", "
   g_str_Parame = g_str_Parame & CStr(r_int_EvaPer) & ", "
   
   g_str_Parame = g_str_Parame & "1, "
   
   'Datos de Auditoria
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
   g_str_Parame = g_str_Parame & "2)"
         
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
      MsgBox "No se pudo ejecutar el procedimiento USP_CRE_CABPBP.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   moddat_g_int_FlgAct_1 = 2
   
   Unload Me
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   Call gs_CentraForm(Me)
   
   Call fs_Inicia
   Call fs_Limpia
  ' Call fs_Buscar_DatCre
   Call modmip_gs_DatNumOpe(moddat_g_str_NumOpe, grd_Listad)
   Call fs_Buscar_DatPBP
   Call fs_Buscar_Cuotas
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   grd_Listad.ColWidth(0) = 2500
   grd_Listad.ColWidth(1) = 9000
   grd_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
   
   grd_LisCuo.ColWidth(0) = 1745
   grd_LisCuo.ColWidth(1) = 2195
   grd_LisCuo.ColWidth(2) = 2195
   grd_LisCuo.ColWidth(3) = 1715
   grd_LisCuo.ColWidth(4) = 2975
   grd_LisCuo.ColAlignment(0) = flexAlignCenterCenter
   grd_LisCuo.ColAlignment(1) = flexAlignCenterCenter
   grd_LisCuo.ColAlignment(2) = flexAlignCenterCenter
   grd_LisCuo.ColAlignment(3) = flexAlignCenterCenter
   grd_LisCuo.ColAlignment(4) = flexAlignCenterCenter
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_FlgPBP, 1, "275")
End Sub

Private Sub fs_Limpia()
   Call gs_LimpiaGrid(grd_Listad)
   Call gs_LimpiaGrid(grd_LisCuo)
   
   pnl_CuoCon.Caption = ""
   pnl_VctCof.Caption = ""
   pnl_CuoCof.Caption = "0.00 "
   pnl_VctCli.Caption = ""
   pnl_CuoCli.Caption = "0.00 "
   cmb_FlgPBP.ListIndex = -1
   pnl_Pen_CuoIni.Caption = "0"
   pnl_Pen_CuoFin.Caption = "0"
   pnl_Tot_Capita_Cli.Caption = "0.00 "
   pnl_Tot_Intere_Cli.Caption = "0.00 "
   pnl_Tot_Capita_Cof.Caption = "0.00 "
   pnl_Tot_Intere_Cof.Caption = "0.00 "
   pnl_Tot_Comisi_Cof.Caption = "0.00 "

   ipp_CapPen_Cli_1.Value = 0:         ipp_CapPen_Cli_2.Value = 0:         ipp_CapPen_Cli_3.Value = 0:         ipp_CapPen_Cli_4.Value = 0:         ipp_CapPen_Cli_5.Value = 0:         ipp_CapPen_Cli_6.Value = 0:
   ipp_IntPen_Cli_1.Value = 0:         ipp_IntPen_Cli_2.Value = 0:         ipp_IntPen_Cli_3.Value = 0:         ipp_IntPen_Cli_4.Value = 0:         ipp_IntPen_Cli_5.Value = 0:         ipp_IntPen_Cli_6.Value = 0:
   ipp_CapPen_Cof_1.Value = 0:         ipp_CapPen_Cof_2.Value = 0:         ipp_CapPen_Cof_3.Value = 0:         ipp_CapPen_Cof_4.Value = 0:         ipp_CapPen_Cof_5.Value = 0:         ipp_CapPen_Cof_6.Value = 0:
   ipp_IntPen_Cof_1.Value = 0:         ipp_IntPen_Cof_2.Value = 0:         ipp_IntPen_Cof_3.Value = 0:         ipp_IntPen_Cof_4.Value = 0:         ipp_IntPen_Cof_5.Value = 0:         ipp_IntPen_Cof_6.Value = 0:
   ipp_ComPen_Cof_1.Value = 0:         ipp_ComPen_Cof_2.Value = 0:         ipp_ComPen_Cof_3.Value = 0:         ipp_ComPen_Cof_4.Value = 0:         ipp_ComPen_Cof_5.Value = 0:         ipp_ComPen_Cof_6.Value = 0:
   
   pnl_Ctr_CapPen_Cli.Caption = "0.00 "
   pnl_Ctr_IntPen_Cli.Caption = "0.00 "
   pnl_Ctr_CapPen_Cof.Caption = "0.00 "
   pnl_Ctr_IntPen_Cof.Caption = "0.00 "
   pnl_Ctr_ComPen_Cof.Caption = "0.00 "
End Sub

Private Sub grd_LisCuo_SelChange()
   If grd_LisCuo.Rows > 2 Then
      grd_LisCuo.RowSel = grd_LisCuo.Row
   End If
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

Private Sub fs_Buscar_DatCre()
   Dim r_str_SimMon     As String
   Dim r_int_Contad     As Integer
   Dim r_str_CodPry     As String
   Dim r_str_NomPry     As String
   Dim r_str_CodBco     As String
   
   'Buscando Información del Crédito
   g_str_Parame = "SELECT * FROM CRE_HIPMAE WHERE "
   g_str_Parame = g_str_Parame & "HIPMAE_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Exit Sub
   End If

   g_rst_Princi.MoveFirst
   
   'Almacenando en Variables Globales
   moddat_g_int_TipDoc = g_rst_Princi!HIPMAE_TDOCLI
   moddat_g_str_NumDoc = Trim(g_rst_Princi!HIPMAE_NDOCLI)
   moddat_g_str_NomCli = moddat_gf_Buscar_NomCli(g_rst_Princi!HIPMAE_TDOCLI, Trim(g_rst_Princi!HIPMAE_NDOCLI))
   moddat_g_str_NumSol = Trim(g_rst_Princi!hipmae_numsol)
   moddat_g_str_NumOpe = Trim(g_rst_Princi!HIPMAE_NUMOPE)
   
   moddat_g_str_CodPrd = Trim(g_rst_Princi!HIPMAE_CODPRD)
   moddat_g_str_NomPrd = moddat_gf_Consulta_Produc(Trim(g_rst_Princi!HIPMAE_CODPRD))
   
   moddat_g_str_Moneda = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!HIPMAE_MONEDA))
   moddat_g_int_TipMon = g_rst_Princi!HIPMAE_MONEDA

   'Obteniendo Información del Inmueble
   Call moddat_gs_Consulta_DatInm(moddat_g_str_NumSol, moddat_g_str_Direcc, moddat_g_str_Distri, r_str_CodPry, r_str_NomPry, r_str_CodBco)
   
   'Cargando en Grid
   grd_Listad.Rows = grd_Listad.Rows + 1:       grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0:                          grd_Listad.Text = "Período"
   grd_Listad.Col = 1:                          grd_Listad.Text = moddat_gf_Consulta_ParDes("033", moddat_g_str_Codigo) & " - " & moddat_g_str_CodIte
   
   grd_Listad.Rows = grd_Listad.Rows + 1:       grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0:                          grd_Listad.Text = "Número de Operación"
   grd_Listad.Col = 1:                          grd_Listad.Text = gf_Formato_NumOpe(g_rst_Princi!HIPMAE_NUMOPE)
   
   grd_Listad.Rows = grd_Listad.Rows + 1:       grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0:                          grd_Listad.Text = "Cliente"
   grd_Listad.Col = 1:                          grd_Listad.Text = CStr(g_rst_Princi!HIPMAE_TDOCLI) & " - " & Trim(g_rst_Princi!HIPMAE_NDOCLI) & " / " & moddat_g_str_NomCli
   
   grd_Listad.Rows = grd_Listad.Rows + 1:       grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0:                          grd_Listad.Text = "Producto"
   grd_Listad.Col = 1:                          grd_Listad.Text = moddat_g_str_NomPrd
   
   grd_Listad.Rows = grd_Listad.Rows + 1:       grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0:                          grd_Listad.Text = "Moneda"
   grd_Listad.Col = 1:                          grd_Listad.Text = moddat_g_str_Moneda
   
   grd_Listad.Rows = grd_Listad.Rows + 1:       grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0:                          grd_Listad.Text = "Dirección Inmueble"
   grd_Listad.Col = 1:                          grd_Listad.Text = moddat_g_str_Direcc
   
   grd_Listad.Rows = grd_Listad.Rows + 1:       grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0:                          grd_Listad.Text = "Distrito"
   grd_Listad.Col = 1:                          grd_Listad.Text = moddat_g_str_Distri
   
   grd_Listad.Rows = grd_Listad.Rows + 1:       grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0:                          grd_Listad.Text = "Fecha Desembolso"
   grd_Listad.Col = 1:                          grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECDES))
   
   If moddat_g_str_CodPrd <> "002" Then
      grd_Listad.Rows = grd_Listad.Rows + 2:    grd_Listad.Row = grd_Listad.Rows - 1
      grd_Listad.Col = 0
      
      Select Case moddat_g_str_CodPrd > 0
         Case InStr(moddat_g_str_AgrCRC, moddat_g_str_CodPrd): grd_Listad.Text = "Nro. Operación Mivivienda"   '"001"
         Case InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd): grd_Listad.Text = "Nro. Operación COFIDE"       '"003"
         Case InStr(moddat_g_str_AgrMIHG, moddat_g_str_CodPrd): grd_Listad.Text = "Nro. Operación COFIDE"      '"004"
         Case InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd) Or InStr(moddat_g_str_Agr2FMV, moddat_g_str_CodPrd): grd_Listad.Text = "Nro. Operación COFIDE" '"007", "009", "010", "013", "014", "015", "016", "017", "018", "019", "020", "021", "022", "023"
      End Select
      
      grd_Listad.Col = 1:                       grd_Listad.Text = Trim(g_rst_Princi!HIPMAE_OPEMVI & "")
      
      If InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Then 'moddat_g_str_CodPrd = "003" Then
         grd_Listad.Rows = grd_Listad.Rows + 1: grd_Listad.Row = grd_Listad.Rows - 1
         grd_Listad.Col = 0:                    grd_Listad.Text = "Nro. Operación Mivivienda"
         grd_Listad.Col = 1:                    grd_Listad.Text = Trim(g_rst_Princi!HIPMAE_OPEMV1 & "")
      End If
      
      If InStr(moddat_g_str_AgrCRC, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Then 'moddat_g_str_CodPrd = "001" Or moddat_g_str_CodPrd = "003" Then
         grd_Listad.Rows = grd_Listad.Rows + 1: grd_Listad.Row = grd_Listad.Rows - 1
         grd_Listad.Col = 0:                    grd_Listad.Text = "Tasa Interés Mivivienda"
         grd_Listad.Col = 1:                    grd_Listad.Text = Format(g_rst_Princi!HIPMAE_TASMVI, "##0.00") & " %"
      End If
      
      If InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_AgrTFMV, moddat_g_str_CodPrd) > 0 Then
      '  moddat_g_str_CodPrd = "004" Or moddat_g_str_CodPrd = "003" Or moddat_g_str_CodPrd = "007" Or moddat_g_str_CodPrd = "009" Or moddat_g_str_CodPrd = "010" Or _
         moddat_g_str_CodPrd = "013" Or moddat_g_str_CodPrd = "014" Or moddat_g_str_CodPrd = "015" Or moddat_g_str_CodPrd = "016" Or moddat_g_str_CodPrd = "017" Or _
         moddat_g_str_CodPrd = "018" Or moddat_g_str_CodPrd = "019" Or moddat_g_str_CodPrd = "020" Or moddat_g_str_CodPrd = "021" Or moddat_g_str_CodPrd = "022" Or moddat_g_str_CodPrd = "023" Then
         
         grd_Listad.Rows = grd_Listad.Rows + 1: grd_Listad.Row = grd_Listad.Rows - 1
         grd_Listad.Col = 0:                    grd_Listad.Text = "Tasa Interés COFIDE"
         grd_Listad.Col = 1:                    grd_Listad.Text = Format(g_rst_Princi!HIPMAE_TASCOF, "##0.00") & " %"
         
         grd_Listad.Rows = grd_Listad.Rows + 1: grd_Listad.Row = grd_Listad.Rows - 1
         grd_Listad.Col = 0:                    grd_Listad.Text = "Tasa Comisión COFIDE"
         grd_Listad.Col = 1:                    grd_Listad.Text = Format(g_rst_Princi!HIPMAE_COMCOF, "##0.00") & " %"
      End If
   End If
   
   grd_Listad.Rows = grd_Listad.Rows + 2:    grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0:                       grd_Listad.Text = "Plazo"
   grd_Listad.Col = 1:                       grd_Listad.Text = CStr(g_rst_Princi!HIPMAE_PLAANO) & " Años"
   
   grd_Listad.Rows = grd_Listad.Rows + 1:    grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0:                       grd_Listad.Text = "Tasa Interés"
   grd_Listad.Col = 1:                       grd_Listad.Text = Format(g_rst_Princi!HIPMAE_TASINT, "##0.00") & " %"
   
   grd_Listad.Rows = grd_Listad.Rows + 1:    grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0:                       grd_Listad.Text = "Nro. Cuotas"
   grd_Listad.Col = 1:                       grd_Listad.Text = CStr(g_rst_Princi!HIPMAE_NUMCUO)
   
   grd_Listad.Rows = grd_Listad.Rows + 1:    grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0:                       grd_Listad.Text = "Período Gracia"
   grd_Listad.Col = 1:                       grd_Listad.Text = CStr(g_rst_Princi!HIPMAE_PERGRA) & " Meses"
   
   grd_Listad.Rows = grd_Listad.Rows + 2:    grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0:                       grd_Listad.Text = "Saldo Capital"
   grd_Listad.Col = 1:                       grd_Listad.CellFontName = "Lucida Console"
   grd_Listad.CellFontSize = 8:              grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_SALCAP, 12, 2)
   
   grd_Listad.Rows = grd_Listad.Rows + 1:    grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0:                       grd_Listad.Text = "Saldo Capital (Tramo Conces.)"
   grd_Listad.Col = 1:                       grd_Listad.CellFontName = "Lucida Console"
   grd_Listad.CellFontSize = 8:              grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_SALCON, 12, 2)
   
   grd_Listad.Rows = grd_Listad.Rows + 1:    grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0:                       grd_Listad.Text = "Total Saldo Capital"
   grd_Listad.Col = 1:                       grd_Listad.CellFontName = "Lucida Console"
   grd_Listad.CellFontSize = 8:              grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_SALCAP + g_rst_Princi!HIPMAE_SALCON, 12, 2)
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call gs_UbiIniGrid(grd_Listad)
   
   r_str_SimMon = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon))
   
   For r_int_Contad = 0 To lbl_SimMon.Count - 1
      lbl_SimMon(r_int_Contad).Caption = r_str_SimMon
   Next r_int_Contad
End Sub

Private Sub fs_Buscar_DatPBP()
   Dim r_int_Contad     As Integer
   Dim r_int_ConAux     As Integer
   
   g_str_Parame = "SELECT * FROM CRE_DETPBP WHERE "
   g_str_Parame = g_str_Parame & "DETPBP_NUMOPE = '" & moddat_g_str_NumOpe & "' AND "
   g_str_Parame = g_str_Parame & "DETPBP_PERMES = " & moddat_g_str_Codigo & " AND "
   g_str_Parame = g_str_Parame & "DETPBP_PERANO = " & moddat_g_str_CodIte & " "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      l_int_CEvIni = g_rst_Princi!DETPBP_CEVAIN
      l_int_CEvFin = g_rst_Princi!DETPBP_CEVAFN
      
      l_dbl_SalAct = g_rst_Princi!DETPBP_SALACT
      l_dbl_SalNue = g_rst_Princi!DETPBP_SALNUE
      
      pnl_CuoCon.Caption = CStr(g_rst_Princi!DETPBP_CUOCON)
      pnl_CuoCof.Caption = Format(g_rst_Princi!DETPBP_CAPADE + g_rst_Princi!DETPBP_INTADE + g_rst_Princi!DETPBP_COMADE, "###,##0.00") & " "
      pnl_CuoCli.Caption = Format(g_rst_Princi!DETPBP_CAPCLI + g_rst_Princi!DETPBP_INTCLI, "###,##0.00") & " "
      
      l_int_IniPen = g_rst_Princi!DETPBP_CIPNCL
      l_int_FinPen = g_rst_Princi!DETPBP_CFPNCL
      
      pnl_Pen_CuoIni.Caption = CStr(g_rst_Princi!DETPBP_CIPNCL)
      pnl_Pen_CuoFin.Caption = CStr(g_rst_Princi!DETPBP_CFPNCL)
      
      pnl_Tot_Capita_Cli.Caption = Format(g_rst_Princi!DETPBP_CAPCLI, "###,##0.00") & " "
      pnl_Tot_Intere_Cli.Caption = Format(g_rst_Princi!DETPBP_INTCLI, "###,##0.00") & " "
      
      pnl_Tot_Capita_Cof.Caption = Format(g_rst_Princi!DETPBP_CAPADE, "###,##0.00") & " "
      pnl_Tot_Intere_Cof.Caption = Format(g_rst_Princi!DETPBP_INTADE, "###,##0.00") & " "
      pnl_Tot_Comisi_Cof.Caption = Format(g_rst_Princi!DETPBP_COMADE, "###,##0.00") & " "
      
      Call gs_BuscarCombo_Item(cmb_FlgPBP, g_rst_Princi!DETPBP_FLGPBP)
      
      If cmb_FlgPBP.ItemData(cmb_FlgPBP.ListIndex) = 1 Or cmb_FlgPBP.ItemData(cmb_FlgPBP.ListIndex) = 3 Then
         ipp_CapPen_Cli_1.Enabled = False
         ipp_CapPen_Cli_2.Enabled = False
         ipp_CapPen_Cli_3.Enabled = False
         ipp_CapPen_Cli_4.Enabled = False
         ipp_CapPen_Cli_5.Enabled = False
         ipp_CapPen_Cli_6.Enabled = False
      
         ipp_IntPen_Cli_1.Enabled = False
         ipp_IntPen_Cli_2.Enabled = False
         ipp_IntPen_Cli_3.Enabled = False
         ipp_IntPen_Cli_4.Enabled = False
         ipp_IntPen_Cli_5.Enabled = False
         ipp_IntPen_Cli_6.Enabled = False
         
         ipp_CapPen_Cof_1.Enabled = False
         ipp_CapPen_Cof_2.Enabled = False
         ipp_CapPen_Cof_3.Enabled = False
         ipp_CapPen_Cof_4.Enabled = False
         ipp_CapPen_Cof_5.Enabled = False
         ipp_CapPen_Cof_6.Enabled = False

         ipp_IntPen_Cof_1.Enabled = False
         ipp_IntPen_Cof_2.Enabled = False
         ipp_IntPen_Cof_3.Enabled = False
         ipp_IntPen_Cof_4.Enabled = False
         ipp_IntPen_Cof_5.Enabled = False
         ipp_IntPen_Cof_6.Enabled = False
         
         ipp_ComPen_Cof_1.Enabled = False
         ipp_ComPen_Cof_2.Enabled = False
         ipp_ComPen_Cof_3.Enabled = False
         ipp_ComPen_Cof_4.Enabled = False
         ipp_ComPen_Cof_5.Enabled = False
         ipp_ComPen_Cof_6.Enabled = False
         
      ElseIf cmb_FlgPBP.ItemData(cmb_FlgPBP.ListIndex) = 2 Then
         ipp_CapPen_Cli_1.Value = g_rst_Princi!DETPBP_CAPCL1
         ipp_IntPen_Cli_1.Value = g_rst_Princi!DETPBP_INPCL1
         
         ipp_CapPen_Cli_2.Value = g_rst_Princi!DETPBP_CAPCL2
         ipp_IntPen_Cli_2.Value = g_rst_Princi!DETPBP_INPCL2
         
         ipp_CapPen_Cli_3.Value = g_rst_Princi!DETPBP_CAPCL3
         ipp_IntPen_Cli_3.Value = g_rst_Princi!DETPBP_INPCL3
         
         ipp_CapPen_Cli_4.Value = g_rst_Princi!DETPBP_CAPCL4
         ipp_IntPen_Cli_4.Value = g_rst_Princi!DETPBP_INPCL4
         
         ipp_CapPen_Cli_5.Value = g_rst_Princi!DETPBP_CAPCL5
         ipp_IntPen_Cli_5.Value = g_rst_Princi!DETPBP_INPCL5
         
         ipp_CapPen_Cli_6.Value = g_rst_Princi!DETPBP_CAPCL6
         ipp_IntPen_Cli_6.Value = g_rst_Princi!DETPBP_INPCL6
         
         'Penalidad COFIDE
         ipp_CapPen_Cof_1.Value = g_rst_Princi!DETPBP_CAPCO1
         ipp_IntPen_Cof_1.Value = g_rst_Princi!DETPBP_INPCO1
         ipp_ComPen_Cof_1.Value = g_rst_Princi!DETPBP_COPCO1
         
         ipp_CapPen_Cof_2.Value = g_rst_Princi!DETPBP_CAPCO2
         ipp_IntPen_Cof_2.Value = g_rst_Princi!DETPBP_INPCO2
         ipp_ComPen_Cof_2.Value = g_rst_Princi!DETPBP_COPCO2
         
         ipp_CapPen_Cof_3.Value = g_rst_Princi!DETPBP_CAPCO3
         ipp_IntPen_Cof_3.Value = g_rst_Princi!DETPBP_INPCO3
         ipp_ComPen_Cof_3.Value = g_rst_Princi!DETPBP_COPCO3
         
         ipp_CapPen_Cof_4.Value = g_rst_Princi!DETPBP_CAPCO4
         ipp_IntPen_Cof_4.Value = g_rst_Princi!DETPBP_INPCO4
         ipp_ComPen_Cof_4.Value = g_rst_Princi!DETPBP_COPCO4
         
         ipp_CapPen_Cof_5.Value = g_rst_Princi!DETPBP_CAPCO5
         ipp_IntPen_Cof_5.Value = g_rst_Princi!DETPBP_INPCO5
         ipp_ComPen_Cof_5.Value = g_rst_Princi!DETPBP_COPCO5
         
         ipp_CapPen_Cof_6.Value = g_rst_Princi!DETPBP_CAPCO6
         ipp_IntPen_Cof_6.Value = g_rst_Princi!DETPBP_INPCO6
         ipp_ComPen_Cof_6.Value = g_rst_Princi!DETPBP_COPCO6
         
         ipp_CapPen_Cli_1.Enabled = True
         ipp_CapPen_Cli_2.Enabled = True
         ipp_CapPen_Cli_3.Enabled = True
         ipp_CapPen_Cli_4.Enabled = True
         ipp_CapPen_Cli_5.Enabled = True
         ipp_CapPen_Cli_6.Enabled = True
      
         ipp_IntPen_Cli_1.Enabled = True
         ipp_IntPen_Cli_2.Enabled = True
         ipp_IntPen_Cli_3.Enabled = True
         ipp_IntPen_Cli_4.Enabled = True
         ipp_IntPen_Cli_5.Enabled = True
         ipp_IntPen_Cli_6.Enabled = True
         
         ipp_CapPen_Cof_1.Enabled = True
         ipp_CapPen_Cof_2.Enabled = True
         ipp_CapPen_Cof_3.Enabled = True
         ipp_CapPen_Cof_4.Enabled = True
         ipp_CapPen_Cof_5.Enabled = True
         ipp_CapPen_Cof_6.Enabled = True

         ipp_IntPen_Cof_1.Enabled = True
         ipp_IntPen_Cof_2.Enabled = True
         ipp_IntPen_Cof_3.Enabled = True
         ipp_IntPen_Cof_4.Enabled = True
         ipp_IntPen_Cof_5.Enabled = True
         ipp_IntPen_Cof_6.Enabled = True
         
         ipp_ComPen_Cof_1.Enabled = True
         ipp_ComPen_Cof_2.Enabled = True
         ipp_ComPen_Cof_3.Enabled = True
         ipp_ComPen_Cof_4.Enabled = True
         ipp_ComPen_Cof_5.Enabled = True
         ipp_ComPen_Cof_6.Enabled = True
      End If
      
      If InStr(moddat_g_str_Agr2MIC, moddat_g_str_CodPrd) > 0 Then 'moddat_g_str_CodPrd = "006" Then
         l_int_IniPen = g_rst_Princi!DETPBP_CIPNCL
         l_int_FinPen = g_rst_Princi!DETPBP_CFPNCL
      
         r_int_ConAux = 0
         
         For r_int_Contad = g_rst_Princi!DETPBP_CIPNCL To g_rst_Princi!DETPBP_CFPNCL
            lbl_NumCuo_Cli(r_int_ConAux).Caption = "Cuota Nro. " & Format(r_int_Contad, "000") & ":"
            r_int_ConAux = r_int_ConAux + 1
         Next r_int_Contad
         
         ipp_CapPen_Cof_1.Enabled = False:   ipp_CapPen_Cof_2.Enabled = False:   ipp_CapPen_Cof_3.Enabled = False:   ipp_CapPen_Cof_4.Enabled = False:   ipp_CapPen_Cof_5.Enabled = False:   ipp_CapPen_Cof_6.Enabled = False
         ipp_IntPen_Cof_1.Enabled = False:   ipp_IntPen_Cof_2.Enabled = False:   ipp_IntPen_Cof_3.Enabled = False:   ipp_IntPen_Cof_4.Enabled = False:   ipp_IntPen_Cof_5.Enabled = False:   ipp_IntPen_Cof_6.Enabled = False
         ipp_ComPen_Cof_1.Enabled = False:   ipp_ComPen_Cof_2.Enabled = False:   ipp_ComPen_Cof_3.Enabled = False:   ipp_ComPen_Cof_4.Enabled = False:   ipp_ComPen_Cof_5.Enabled = False:   ipp_ComPen_Cof_6.Enabled = False
      Else
         l_int_IniPen = g_rst_Princi!DETPBP_CIPNCO
         l_int_FinPen = g_rst_Princi!DETPBP_CFPNCO
         
         r_int_ConAux = 0
         
         For r_int_Contad = g_rst_Princi!DETPBP_CIPNCO To g_rst_Princi!DETPBP_CFPNCO
            lbl_NumCuo_Cli(r_int_ConAux).Caption = "Cuota Nro. " & Format(r_int_Contad, "000") & ":"
            lbl_NumCuo_Cof(r_int_ConAux).Caption = "Cuota Nro. " & Format(r_int_Contad, "000") & ":"
            
            r_int_ConAux = r_int_ConAux + 1
         Next r_int_Contad
      End If
   End If

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Para obtener Información de Cuota Concesional Tramo Cofide/Mivivienda
   If InStr(moddat_g_str_Agr2MIC, moddat_g_str_CodPrd) = 0 Then 'moddat_g_str_CodPrd <> "006" Then
      g_str_Parame = "SELECT * FROM CRE_HIPCUO WHERE "
      g_str_Parame = g_str_Parame & "HIPCUO_NUMOPE = '" & moddat_g_str_NumOpe & "' AND "
      g_str_Parame = g_str_Parame & "HIPCUO_NUMCUO = " & pnl_CuoCon.Caption & " AND "
      g_str_Parame = g_str_Parame & "HIPCUO_TIPCRO = 4 "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
   
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         pnl_VctCof.Caption = gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECVCT))
      End If
               
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   End If
   
   'Para obtener Información de Cuota Concesional Tramo cliente
   g_str_Parame = "SELECT * FROM CRE_HIPCUO WHERE "
   g_str_Parame = g_str_Parame & "HIPCUO_NUMOPE = '" & moddat_g_str_NumOpe & "' AND "
   g_str_Parame = g_str_Parame & "HIPCUO_NUMCUO = " & pnl_CuoCon.Caption & " AND "
   g_str_Parame = g_str_Parame & "HIPCUO_TIPCRO = 2 "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If

   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      pnl_VctCli.Caption = gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECVCT))
   End If
            
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Buscar_Cuotas()
   g_str_Parame = "SELECT * FROM CRE_HIPCUO WHERE "
   g_str_Parame = g_str_Parame & "HIPCUO_NUMOPE = '" & moddat_g_str_NumOpe & "' AND "
   g_str_Parame = g_str_Parame & "HIPCUO_TIPCRO = 1  AND "
   g_str_Parame = g_str_Parame & "HIPCUO_NUMCUO >= " & CStr(l_int_CEvIni) & " AND "
   g_str_Parame = g_str_Parame & "HIPCUO_NUMCUO <= " & CStr(l_int_CEvFin) & " ORDER BY HIPCUO_NUMCUO ASC"
   
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
      grd_LisCuo.Rows = grd_LisCuo.Rows + 1
      grd_LisCuo.Row = grd_LisCuo.Rows - 1
   
      grd_LisCuo.Col = 0
      grd_LisCuo.Text = CStr(g_rst_Princi!HIPCUO_NUMCUO)
      
      grd_LisCuo.Col = 1
      grd_LisCuo.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECVCT))
      
      grd_LisCuo.Col = 2
      If g_rst_Princi!HIPCUO_FECPAG > 0 Then
         grd_LisCuo.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPCUO_FECPAG))
      End If
      
      If g_rst_Princi!HIPCUO_SITUAC = 2 Then
         grd_LisCuo.Col = 3
         
         If CInt(CDate(Format(ff_Ultimo_Dia_Mes(CInt(moddat_g_str_Codigo), CInt(moddat_g_str_CodIte)), "00") & "/" & Format(moddat_g_str_Codigo, "00") & "/" & Format(moddat_g_str_CodIte, "0000")) - CDate(gf_FormatoFecha(g_rst_Princi!HIPCUO_FECVCT))) < 0 Then
            grd_LisCuo.Text = ""
         Else
            grd_LisCuo.Text = CStr(CInt(CDate(Format(ff_Ultimo_Dia_Mes(CInt(moddat_g_str_Codigo), CInt(moddat_g_str_CodIte)), "00") & "/" & Format(moddat_g_str_Codigo, "00") & "/" & Format(moddat_g_str_CodIte, "0000")) - CDate(gf_FormatoFecha(g_rst_Princi!HIPCUO_FECVCT))))
         End If
         
         If CDate(Format(ff_Ultimo_Dia_Mes(CInt(moddat_g_str_Codigo), CInt(moddat_g_str_CodIte)), "00") & "/" & Format(moddat_g_str_Codigo, "00") & "/" & Format(moddat_g_str_CodIte, "0000")) > CDate(gf_FormatoFecha(g_rst_Princi!HIPCUO_FECVCT)) Then
            grd_LisCuo.Col = 4
            grd_LisCuo.Text = "VENCIDA"
         Else
            grd_LisCuo.Col = 4
            grd_LisCuo.Text = "X VENCER"
         End If
      Else
         If CDate(gf_FormatoFecha(g_rst_Princi!HIPCUO_FECPAG)) < CDate(gf_FormatoFecha(g_rst_Princi!HIPCUO_FECVCT)) Then
            grd_LisCuo.Col = 3
            grd_LisCuo.Text = "0"
         Else
            grd_LisCuo.Col = 3
            grd_LisCuo.Text = CStr(CInt(CDate(gf_FormatoFecha(g_rst_Princi!HIPCUO_FECPAG)) - CDate(gf_FormatoFecha(g_rst_Princi!HIPCUO_FECVCT))))
         End If
         
         grd_LisCuo.Col = 4
         grd_LisCuo.Text = "PAGADA"
      End If
   
      g_rst_Princi.MoveNext
   Loop
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call gs_UbiIniGrid(grd_LisCuo)
End Sub

Private Sub ipp_CapPen_Cli_1_Change()
   Call fs_TotCapPen_Cli
End Sub

Private Sub ipp_CapPen_Cli_2_Change()
   Call fs_TotCapPen_Cli
End Sub

Private Sub ipp_CapPen_Cli_3_Change()
   Call fs_TotCapPen_Cli
End Sub

Private Sub ipp_CapPen_Cli_4_Change()
   Call fs_TotCapPen_Cli
End Sub

Private Sub ipp_CapPen_Cli_5_Change()
   Call fs_TotCapPen_Cli
End Sub

Private Sub ipp_CapPen_Cli_6_Change()
   Call fs_TotCapPen_Cli
End Sub

Private Sub ipp_IntPen_Cli_1_Change()
   Call fs_TotIntPen_Cli
End Sub

Private Sub ipp_IntPen_Cli_2_Change()
   Call fs_TotIntPen_Cli
End Sub

Private Sub ipp_IntPen_Cli_3_Change()
   Call fs_TotIntPen_Cli
End Sub

Private Sub ipp_IntPen_Cli_4_Change()
   Call fs_TotIntPen_Cli
End Sub

Private Sub ipp_IntPen_Cli_5_Change()
   Call fs_TotIntPen_Cli
End Sub

Private Sub ipp_IntPen_Cli_6_Change()
   Call fs_TotIntPen_Cli
End Sub

Private Sub ipp_CapPen_Cof_1_Change()
   Call fs_TotCapPen_Cof
End Sub

Private Sub ipp_CapPen_Cof_2_Change()
   Call fs_TotCapPen_Cof
End Sub

Private Sub ipp_CapPen_Cof_3_Change()
   Call fs_TotCapPen_Cof
End Sub

Private Sub ipp_CapPen_Cof_4_Change()
   Call fs_TotCapPen_Cof
End Sub

Private Sub ipp_CapPen_Cof_5_Change()
   Call fs_TotCapPen_Cof
End Sub

Private Sub ipp_CapPen_Cof_6_Change()
   Call fs_TotCapPen_Cof
End Sub

Private Sub ipp_IntPen_Cof_1_Change()
   Call fs_TotIntPen_Cof
End Sub

Private Sub ipp_IntPen_Cof_2_Change()
   Call fs_TotIntPen_Cof
End Sub

Private Sub ipp_IntPen_Cof_3_Change()
   Call fs_TotIntPen_Cof
End Sub

Private Sub ipp_IntPen_Cof_4_Change()
   Call fs_TotIntPen_Cof
End Sub

Private Sub ipp_IntPen_Cof_5_Change()
   Call fs_TotIntPen_Cof
End Sub

Private Sub ipp_IntPen_Cof_6_Change()
   Call fs_TotIntPen_Cof
End Sub

Private Sub ipp_ComPen_Cof_1_Change()
   Call fs_TotComPen_Cof
End Sub

Private Sub ipp_ComPen_Cof_2_Change()
   Call fs_TotComPen_Cof
End Sub

Private Sub ipp_ComPen_Cof_3_Change()
   Call fs_TotComPen_Cof
End Sub

Private Sub ipp_ComPen_Cof_4_Change()
   Call fs_TotComPen_Cof
End Sub

Private Sub ipp_ComPen_Cof_5_Change()
   Call fs_TotComPen_Cof
End Sub

Private Sub ipp_ComPen_Cof_6_Change()
   Call fs_TotComPen_Cof
End Sub

Private Sub ipp_CapPen_Cli_1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_IntPen_Cli_1)
   End If
End Sub

Private Sub ipp_IntPen_Cli_1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_CapPen_Cli_2)
   End If
End Sub

Private Sub ipp_CapPen_Cli_2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_IntPen_Cli_2)
   End If
End Sub

Private Sub ipp_IntPen_Cli_2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_CapPen_Cli_3)
   End If
End Sub

Private Sub ipp_CapPen_Cli_3_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_IntPen_Cli_3)
   End If
End Sub

Private Sub ipp_IntPen_Cli_3_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_CapPen_Cli_4)
   End If
End Sub

Private Sub ipp_CapPen_Cli_4_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_IntPen_Cli_4)
   End If
End Sub

Private Sub ipp_IntPen_Cli_4_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_CapPen_Cli_5)
   End If
End Sub

Private Sub ipp_CapPen_Cli_5_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_IntPen_Cli_5)
   End If
End Sub

Private Sub ipp_IntPen_Cli_5_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_CapPen_Cli_6)
   End If
End Sub

Private Sub ipp_CapPen_Cli_6_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_IntPen_Cli_6)
   End If
End Sub

Private Sub ipp_IntPen_Cli_6_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If ipp_CapPen_Cof_1.Enabled Then
         Call gs_SetFocus(ipp_CapPen_Cof_1)
      Else
         Call gs_SetFocus(cmd_Grabar)
      End If
   End If
End Sub

Private Sub ipp_CapPen_Cof_1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_IntPen_Cof_1)
   End If
End Sub

Private Sub ipp_IntPen_Cof_1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ComPen_Cof_1)
   End If
End Sub

Private Sub ipp_ComPen_Cof_1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_CapPen_Cof_2)
   End If
End Sub

Private Sub ipp_CapPen_Cof_2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_IntPen_Cof_2)
   End If
End Sub

Private Sub ipp_IntPen_Cof_2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ComPen_Cof_2)
   End If
End Sub

Private Sub ipp_ComPen_Cof_2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_CapPen_Cof_3)
   End If
End Sub

Private Sub ipp_CapPen_Cof_3_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_IntPen_Cof_3)
   End If
End Sub

Private Sub ipp_IntPen_Cof_3_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ComPen_Cof_3)
   End If
End Sub

Private Sub ipp_ComPen_Cof_3_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_CapPen_Cof_4)
   End If
End Sub

Private Sub ipp_CapPen_Cof_4_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_IntPen_Cof_4)
   End If
End Sub

Private Sub ipp_IntPen_Cof_4_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ComPen_Cof_4)
   End If
End Sub

Private Sub ipp_ComPen_Cof_4_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_CapPen_Cof_5)
   End If
End Sub

Private Sub ipp_CapPen_Cof_5_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_IntPen_Cof_5)
   End If
End Sub

Private Sub ipp_IntPen_Cof_5_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ComPen_Cof_5)
   End If
End Sub

Private Sub ipp_ComPen_Cof_5_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_CapPen_Cof_6)
   End If
End Sub

Private Sub ipp_CapPen_Cof_6_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_IntPen_Cof_6)
   End If
End Sub

Private Sub ipp_IntPen_Cof_6_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ComPen_Cof_6)
   End If
End Sub

Private Sub ipp_ComPen_Cof_6_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   End If
End Sub

Private Sub fs_TotCapPen_Cli()
   pnl_Ctr_CapPen_Cli.Caption = Format(CDbl(ipp_CapPen_Cli_1.Value) + CDbl(ipp_CapPen_Cli_2.Value) + CDbl(ipp_CapPen_Cli_3.Value) + CDbl(ipp_CapPen_Cli_4.Value) + CDbl(ipp_CapPen_Cli_5.Value) + CDbl(ipp_CapPen_Cli_6.Value), "###,##0.00") & " "
End Sub

Private Sub fs_TotIntPen_Cli()
   pnl_Ctr_IntPen_Cli.Caption = Format(CDbl(ipp_IntPen_Cli_1.Value) + CDbl(ipp_IntPen_Cli_2.Value) + CDbl(ipp_IntPen_Cli_3.Value) + CDbl(ipp_IntPen_Cli_4.Value) + CDbl(ipp_IntPen_Cli_5.Value) + CDbl(ipp_IntPen_Cli_6.Value), "###,##0.00") & " "
End Sub

Private Sub fs_TotCapPen_Cof()
   pnl_Ctr_CapPen_Cof.Caption = Format(CDbl(ipp_CapPen_Cof_1.Value) + CDbl(ipp_CapPen_Cof_2.Value) + CDbl(ipp_CapPen_Cof_3.Value) + CDbl(ipp_CapPen_Cof_4.Value) + CDbl(ipp_CapPen_Cof_5.Value) + CDbl(ipp_CapPen_Cof_6.Value), "###,##0.00") & " "
End Sub

Private Sub fs_TotIntPen_Cof()
   pnl_Ctr_IntPen_Cof.Caption = Format(CDbl(ipp_IntPen_Cof_1.Value) + CDbl(ipp_IntPen_Cof_2.Value) + CDbl(ipp_IntPen_Cof_3.Value) + CDbl(ipp_IntPen_Cof_4.Value) + CDbl(ipp_IntPen_Cof_5.Value) + CDbl(ipp_IntPen_Cof_6.Value), "###,##0.00") & " "
End Sub

Private Sub fs_TotComPen_Cof()
   pnl_Ctr_ComPen_Cof.Caption = Format(CDbl(ipp_ComPen_Cof_1.Value) + CDbl(ipp_ComPen_Cof_2.Value) + CDbl(ipp_ComPen_Cof_3.Value) + CDbl(ipp_ComPen_Cof_4.Value) + CDbl(ipp_ComPen_Cof_5.Value) + CDbl(ipp_ComPen_Cof_6.Value), "###,##0.00") & " "
End Sub

Private Function ff_ValidaCuota(ByVal p_NumOpe As String, ByVal p_NumCuo As Integer) As Integer
   ff_ValidaCuota = False
   
   g_str_Parame = "SELECT * FROM CRE_HIPCUO WHERE "
   g_str_Parame = g_str_Parame & "HIPCUO_NUMOPE = '" & p_NumOpe & "' AND "
   g_str_Parame = g_str_Parame & "HIPCUO_TIPCRO = 1  AND "
   g_str_Parame = g_str_Parame & "HIPCUO_NUMCUO = " & CStr(p_NumCuo) & " "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      If g_rst_Princi!HIPCUO_SITUAC = 1 Then
         ff_ValidaCuota = True
      End If
   End If

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Function

