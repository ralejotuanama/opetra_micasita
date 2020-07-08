VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_Ges_TecPro_10 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8865
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13125
   Icon            =   "OpeTra_frm_833.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   13125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   8850
      Left            =   30
      TabIndex        =   18
      Top             =   0
      Width           =   13125
      _Version        =   65536
      _ExtentX        =   23151
      _ExtentY        =   15610
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
      Begin Threed.SSPanel SSPanel8 
         Height          =   675
         Left            =   30
         TabIndex        =   19
         Top             =   780
         Width           =   13035
         _Version        =   65536
         _ExtentX        =   22992
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
         Begin VB.CommandButton cmd_Editar 
            Height          =   585
            Left            =   720
            Picture         =   "OpeTra_frm_833.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Modificar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   1290
            Picture         =   "OpeTra_frm_833.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Agrega 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_833.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Nuevo Registro"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   12420
            Picture         =   "OpeTra_frm_833.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   20
         Top             =   60
         Width           =   13035
         _Version        =   65536
         _ExtentX        =   22992
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
            TabIndex        =   21
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
         Begin Threed.SSPanel SSPanel15 
            Height          =   315
            Left            =   630
            TabIndex        =   22
            Top             =   330
            Width           =   4215
            _Version        =   65536
            _ExtentX        =   7435
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Entidad Técnica - Histórico"
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
            Picture         =   "OpeTra_frm_833.frx":0D6C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   1305
         Left            =   30
         TabIndex        =   23
         Top             =   1500
         Width           =   13035
         _Version        =   65536
         _ExtentX        =   22992
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
         Begin Threed.SSPanel pnl_RazSoc 
            Height          =   315
            Left            =   1620
            TabIndex        =   24
            Top             =   480
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
            Left            =   1620
            TabIndex        =   25
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
            Left            =   9570
            TabIndex        =   26
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
            Left            =   1620
            TabIndex        =   27
            Top             =   840
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
            TabIndex        =   31
            Top             =   150
            Width           =   1335
         End
         Begin VB.Label lbl_NumDoc 
            Caption         =   "Nro. Documento:"
            Height          =   225
            Left            =   8100
            TabIndex        =   30
            Top             =   150
            Width           =   1335
         End
         Begin VB.Label lbl_RazSoc 
            Caption         =   "Razón Social:"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   510
            Width           =   1335
         End
         Begin VB.Label lbl_TipEmp 
            Caption         =   "Tipo Empresa:"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   870
            Width           =   1335
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   2490
         Left            =   30
         TabIndex        =   32
         Top             =   2850
         Width           =   13035
         _Version        =   65536
         _ExtentX        =   22992
         _ExtentY        =   4392
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
            Height          =   1935
            Left            =   60
            TabIndex        =   33
            Top             =   450
            Width           =   12870
            _ExtentX        =   22701
            _ExtentY        =   3413
            _Version        =   393216
            Rows            =   21
            Cols            =   7
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
            Appearance      =   0
         End
         Begin Threed.SSPanel pnl_Tit_TipGar 
            Height          =   405
            Left            =   60
            TabIndex        =   34
            Top             =   30
            Width           =   1575
            _Version        =   65536
            _ExtentX        =   2778
            _ExtentY        =   714
            _StockProps     =   15
            Caption         =   "Línea Asignada  Cred. Ind."
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
         Begin Threed.SSPanel pnl_Tit_NumRef 
            Height          =   405
            Left            =   3180
            TabIndex        =   35
            Top             =   30
            Width           =   1695
            _Version        =   65536
            _ExtentX        =   2990
            _ExtentY        =   714
            _StockProps     =   15
            Caption         =   "Retención (%)"
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
         Begin Threed.SSPanel pnl_Tit_FecEmi 
            Height          =   405
            Left            =   6450
            TabIndex        =   36
            Top             =   30
            Width           =   1695
            _Version        =   65536
            _ExtentX        =   2990
            _ExtentY        =   714
            _StockProps     =   15
            Caption         =   "F. Aprobación"
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
         Begin Threed.SSPanel SSPanel2 
            Height          =   405
            Left            =   4770
            TabIndex        =   37
            Top             =   30
            Width           =   1695
            _Version        =   65536
            _ExtentX        =   2990
            _ExtentY        =   714
            _StockProps     =   15
            Caption         =   "TEA (%)"
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
            Height          =   405
            Left            =   1620
            TabIndex        =   38
            Top             =   30
            Width           =   1575
            _Version        =   65536
            _ExtentX        =   2778
            _ExtentY        =   714
            _StockProps     =   15
            Caption         =   "Línea Asignada  Cred. Dir."
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
            Height          =   405
            Left            =   8025
            TabIndex        =   55
            Top             =   30
            Width           =   1695
            _Version        =   65536
            _ExtentX        =   2990
            _ExtentY        =   714
            _StockProps     =   15
            Caption         =   "F. Vencimiento"
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
      End
      Begin Threed.SSPanel pnl_LinAsig 
         Height          =   1665
         Left            =   30
         TabIndex        =   39
         Top             =   5385
         Width           =   13035
         _Version        =   65536
         _ExtentX        =   22992
         _ExtentY        =   2937
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
         Begin EditLib.fpDoubleSingle ipp_LinNRe_Ind 
            Height          =   315
            Left            =   2205
            TabIndex        =   2
            Top             =   1200
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
         Begin Threed.SSPanel pnl_LinAsig_Ind 
            Height          =   315
            Left            =   1830
            TabIndex        =   0
            Top             =   510
            Width           =   1935
            _Version        =   65536
            _ExtentX        =   3413
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
         Begin Threed.SSPanel pnl_LinAsig_Dir 
            Height          =   315
            Left            =   6255
            TabIndex        =   3
            Top             =   480
            Width           =   1935
            _Version        =   65536
            _ExtentX        =   3413
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
         Begin EditLib.fpDoubleSingle ipp_LinRev_Dir 
            Height          =   315
            Left            =   6615
            TabIndex        =   4
            Top             =   840
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
         Begin EditLib.fpDoubleSingle ipp_LinNRe_Dir 
            Height          =   315
            Left            =   6615
            TabIndex        =   5
            Top             =   1200
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
         Begin EditLib.fpDoubleSingle ipp_ExpLin 
            Height          =   315
            Left            =   11010
            TabIndex        =   7
            Top             =   840
            Width           =   1935
            _Version        =   196608
            _ExtentX        =   3413
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
         Begin Threed.SSPanel pnl_LinTot 
            Height          =   315
            Left            =   11010
            TabIndex        =   6
            Top             =   480
            Width           =   1935
            _Version        =   65536
            _ExtentX        =   3413
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
         Begin EditLib.fpDoubleSingle ipp_LinRev_Ind 
            Height          =   315
            Left            =   2205
            TabIndex        =   1
            Top             =   840
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
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Revolvente:"
            Height          =   195
            Left            =   330
            TabIndex        =   56
            Top             =   900
            Width           =   870
         End
         Begin VB.Label Label7 
            Caption         =   "Exposición Máx. Línea:"
            Height          =   255
            Left            =   9000
            TabIndex        =   54
            Top             =   870
            Width           =   1695
         End
         Begin VB.Label Label1 
            Caption         =   "Línea Total Asignada:"
            Height          =   255
            Left            =   9000
            TabIndex        =   53
            Top             =   540
            Width           =   1695
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Revolvente:"
            Height          =   195
            Left            =   4920
            TabIndex        =   46
            Top             =   900
            Width           =   870
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "No Revolvente:"
            Height          =   195
            Left            =   4920
            TabIndex        =   45
            Top             =   1260
            Width           =   1125
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "No Revolvente:"
            Height          =   195
            Left            =   330
            TabIndex        =   44
            Top             =   1260
            Width           =   1125
         End
         Begin VB.Label lbl_LinAsig 
            Caption         =   "Línea Asignada:"
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   540
            Width           =   1695
         End
         Begin VB.Label Label24 
            Caption         =   "Línea Asignada:"
            Height          =   255
            Left            =   4710
            TabIndex        =   42
            Top             =   510
            Width           =   1335
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            Caption         =   "Créditos Indirectos"
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
            Left            =   120
            TabIndex        =   41
            Top             =   120
            Width           =   1605
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Créditos Directos"
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
            Left            =   4710
            TabIndex        =   40
            Top             =   120
            Width           =   1470
         End
      End
      Begin Threed.SSPanel pnl_OtrInf 
         Height          =   945
         Left            =   30
         TabIndex        =   47
         Top             =   7080
         Width           =   13035
         _Version        =   65536
         _ExtentX        =   22992
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
         Begin EditLib.fpDoubleSingle ipp_PorRet 
            Height          =   315
            Left            =   2250
            TabIndex        =   8
            Top             =   120
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
         Begin EditLib.fpDateTime ipp_FecVct 
            Height          =   315
            Left            =   6615
            TabIndex        =   11
            Top             =   465
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
         Begin EditLib.fpDoubleSingle ipp_PorTEA 
            Height          =   315
            Left            =   6615
            TabIndex        =   9
            Top             =   120
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
         Begin EditLib.fpDateTime ipp_FecApr 
            Height          =   315
            Left            =   2250
            TabIndex        =   10
            Top             =   480
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
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "% TEA"
            Height          =   195
            Left            =   4710
            TabIndex        =   51
            Top             =   165
            Width           =   480
         End
         Begin VB.Label Label3 
            Caption         =   "F. Aprobación"
            Height          =   285
            Left            =   120
            TabIndex        =   50
            Top             =   480
            Width           =   1545
         End
         Begin VB.Label Label2 
            Caption         =   "F. Vencimiento:"
            Height          =   285
            Left            =   4710
            TabIndex        =   49
            Top             =   465
            Width           =   1545
         End
         Begin VB.Label Label12 
            Caption         =   "FMV - % Retención"
            Height          =   225
            Left            =   120
            TabIndex        =   48
            Top             =   165
            Width           =   1635
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   735
         Left            =   30
         TabIndex        =   52
         Top             =   8055
         Width           =   13035
         _Version        =   65536
         _ExtentX        =   22992
         _ExtentY        =   1296
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
         Begin VB.CommandButton cmd_Cancel 
            Height          =   675
            Left            =   12350
            Picture         =   "OpeTra_frm_833.frx":1076
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Cancelar"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   675
            Left            =   11665
            Picture         =   "OpeTra_frm_833.frx":1380
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   675
         End
      End
   End
End
Attribute VB_Name = "frm_Ges_TecPro_10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Agrega_Click()
   moddat_g_int_FlgGrb_1 = 1
   Call fs_Activa(False)
   Call gs_SetFocus(ipp_LinRev_Ind)
End Sub

Private Sub cmd_Cancel_Click()
   Call fs_Activa(True)
   Call fs_Limpia
   Call gs_SetFocus(grd_Listad)

   If grd_Listad.Rows = 0 Then
      cmd_Agrega.Enabled = True
      cmd_Editar.Enabled = False
      grd_Listad.Enabled = False
   End If
End Sub

Private Sub cmd_Editar_Click()
   
   If grd_Listad.Row = -1 Then Exit Sub
   
   Call fs_Limpia
   grd_Listad.Col = 0
   moddat_g_str_Codigo = grd_Listad.Text
   Call gs_RefrescaGrid(grd_Listad)
   moddat_g_int_FlgGrb_1 = 2
   
   'Obteniendo Información del Registro
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT ETEHIS_LINASI_IND, ETEHIS_LINASI_DIR, ETEHIS_PORRET    , ETEHIS_PORTEA    , ETEHIS_FECAPR , ETEHIS_FECVCT, ETEHIS_ADMFLJ , ETEHIS_IMPHIP , ETEHIS_IMPLIQ, "
   g_str_Parame = g_str_Parame & "        ETEHIS_LINREV_DIR, ETEHIS_LINNRE_DIR, ETEHIS_LINREV_IND, ETEHIS_LINNRE_IND, ETEHIS_LINEXP , ETEHIS_LINASI "
   g_str_Parame = g_str_Parame & "   FROM TPR_ETEHIS A "
   g_str_Parame = g_str_Parame & "  WHERE ETEHIS_TIPDOC = " & CStr(moddat_g_int_TipDoc) & " "
   g_str_Parame = g_str_Parame & "    AND ETEHIS_NUMDOC = '" & CStr(moddat_g_str_NumDoc) & "' "
   g_str_Parame = g_str_Parame & "    AND ETEHIS_NUMITE = " & moddat_g_str_Codigo & ""
   g_str_Parame = g_str_Parame & "  ORDER BY ETEHIS_NUMITE DESC " 'SEGFECCRE, SEGHORCRE

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Sub
   End If
   
   g_rst_Genera.MoveFirst

   If Not IsNull(g_rst_Genera!ETEHIS_LINASI_IND) Then
      pnl_LinAsig_Ind.Caption = Format(CStr(g_rst_Genera!ETEHIS_LINASI_IND), "###,###,###,##0.00") & "  "
   End If
    If Not IsNull(g_rst_Genera!ETEHIS_LINREV_IND) Then
      ipp_LinRev_Ind.Value = g_rst_Genera!ETEHIS_LINREV_IND
   End If
   If Not IsNull(g_rst_Genera!ETEHIS_LINNRE_IND) Then
      ipp_LinNRe_Ind.Value = g_rst_Genera!ETEHIS_LINNRE_IND
   End If
   If Not IsNull(g_rst_Genera!ETEHIS_LINASI_DIR) Then
      pnl_LinAsig_Dir.Caption = Format(CStr(g_rst_Genera!ETEHIS_LINASI_DIR), "###,###,###,##0.00") & "  "
   End If
   If Not IsNull(g_rst_Genera!ETEHIS_LINREV_DIR) Then
      ipp_LinRev_Dir.Value = g_rst_Genera!ETEHIS_LINREV_DIR
   End If
   If Not IsNull(g_rst_Genera!ETEHIS_LINNRE_DIR) Then
      ipp_LinNRe_Dir.Value = g_rst_Genera!ETEHIS_LINNRE_DIR
   End If

   If Not IsNull(g_rst_Genera!ETEHIS_LINEXP) Then
      ipp_ExpLin.Text = Format(CStr(g_rst_Genera!ETEHIS_LINEXP), "###,###,###,##0.00") & "  "
   End If
   If Not IsNull(g_rst_Genera!ETEHIS_LINASI) Then
      pnl_LinTot.Caption = Format(CStr(g_rst_Genera!ETEHIS_LINASI), "###,###,###,##0.00") & "  "
   End If
   If Not IsNull(g_rst_Genera!ETEHIS_PORRET) Then
      ipp_PorRet.Value = g_rst_Genera!ETEHIS_PORRET
   End If
   If Not IsNull(g_rst_Genera!ETEHIS_PORTEA) Then
      ipp_PorTEA.Value = g_rst_Genera!ETEHIS_PORTEA
   End If
   If Not IsNull(g_rst_Genera!ETEHIS_FECAPR) Then
      ipp_FecApr.Text = Format(gf_FormatoFecha(CStr(g_rst_Genera!ETEHIS_FECAPR)), "DD/MM/YYYY") & "  "
   End If
   If Not IsNull(g_rst_Genera!ETEHIS_FECVCT) Then
      ipp_FecVct.Text = Format(gf_FormatoFecha(CStr(g_rst_Genera!ETEHIS_FECVCT)), "DD/MM/YYYY") & "  "
   End If
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   
   If grd_Listad.Row = 0 Then
      Call fs_Activa(False)
   Else
      Call fs_Activa(True)
   End If

   Call gs_SetFocus(ipp_LinRev_Ind)
End Sub


Private Sub cmd_ExpExc_Click()
   'Confirmacion
   If grd_Listad.Rows > 0 Then
      If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
         Exit Sub
      End If
      Screen.MousePointer = 11
      Call fs_GenExc
   Else
      MsgBox "No existen datos a exportar", vbInformation, modgen_g_str_NomPlt
   End If
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Grabar_Click()

   If ipp_ExpLin.Value = 0 Then
'      MsgBox "El importe de Exposición de Línea, es cero.", vbExclamation, modgen_g_str_NomPlt
      MsgBox "Debe especificar el importe de Exposición de línea.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_ExpLin)
      Exit Sub
   End If
   
   If Trim(pnl_LinAsig_Ind.Caption) = 0 And ipp_LinRev_Ind.Value = 0 And ipp_LinNRe_Ind.Value = 0 Then
      MsgBox "El importe de Línea Asignada para Cred. Indirectos, es cero.", vbInformation, modgen_g_str_NomPlt
      'MsgBox "Debe especificar el importe de la línea asignada para Cred. Indirectos.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_LinRev_Ind)
      'Exit Sub
   End If
   If Trim(pnl_LinAsig_Dir.Caption) = 0 And ipp_LinRev_Dir.Value = 0 And ipp_LinNRe_Dir.Value = 0 Then
      MsgBox "El importe de Línea Asignada para Cred. Directos, es cero.", vbInformation, modgen_g_str_NomPlt
      'MsgBox "Debe especificar el importe de la línea asignada para Cred. Directos.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_LinRev_Dir)
      'Exit Sub
   End If
      
   If CDbl(Replace(ipp_PorRet.Value, "%", "")) = 0 Then
      MsgBox "El Porcentaje de Retención es cero.", vbInformation, modgen_g_str_NomPlt
      'MsgBox "Debe ingresar Porcentaje de Retención.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PorRet)
      'Exit Sub
   End If
   
   If CDbl(Replace(ipp_PorTEA.Value, "%", "")) = 0 Then
      MsgBox "TEA es cero.", vbInformation, modgen_g_str_NomPlt
      'MsgBox "Debe ingresar TEA.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PorTEA)
      'Exit Sub
   End If
   
   If Len(ipp_FecApr.Text) = 0 Then
      MsgBox "Debe ingresar una Fecha de Aprobación válida.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecApr)
      Exit Sub
   Else
      If CDate(ipp_FecApr.Value) <= Format(CDate(DateAdd("M", -1, date)), "DD/MM/YYYY") Then
         MsgBox "Debe ingresar una Fecha de Aprobación válida.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_FecApr)
         Exit Sub
      End If
      If CDate(ipp_FecApr.Text) >= CDate(ipp_FecVct.Text) Then
         MsgBox "La Fecha de Aprobación no puede ser mayor a la Fecha de Vencimiento.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_FecApr)
         Exit Sub
      End If
   End If
   
   If Len(ipp_FecVct.Text) = 0 Then
      MsgBox "Debe ingresar una Fecha de vencimiento válida.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecVct)
      Exit Sub
   Else
      If CDate(ipp_FecVct.Text) <= date Then
         MsgBox "Debe ingresar una Fecha de vencimiento válida.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_FecVct)
         Exit Sub
      End If
   End If
   
   'Valida que Línea Asignada sea menor e igual a todas las Cartas Fianzas ya registradas
   If moddat_g_int_FlgGrb = 2 Then
      If fs_Validar_MtoLinAsi = False Then
         MsgBox "Línea Aprobada no puede ser menor a Línea Utilizada.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_LinNRe_Ind)
         Exit Sub
      End If
   End If
   
   'Valida que el Monto Total de Línea Asignada, tenga como tope el 30% del Patrimonio Efectivo
   If pnl_LinTot > fs_PatEfe Then
      MsgBox "Línea Aprobada no puede ser mayor al 30% del Patrimonio Efectivo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_LinRev_Ind)
      Exit Sub
   End If
   
    If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      Screen.MousePointer = 11
      Call moddat_gs_FecSis
   
      'Grabando Información en el histórico de Entidad Técnica
      g_str_Parame = "USP_TPR_ETEHIS ("
      g_str_Parame = g_str_Parame & CStr(Trim(Mid(pnl_TipDoc.Caption, 1, 2))) & ", "
      g_str_Parame = g_str_Parame & "'" & CStr(pnl_NroDoc.Caption) & "', "
      g_str_Parame = g_str_Parame & CDbl(pnl_LinAsig_Ind.Caption) & ", "
      g_str_Parame = g_str_Parame & CDbl(ipp_LinRev_Ind.Text) & ", "
      g_str_Parame = g_str_Parame & CDbl(ipp_LinNRe_Ind.Text) & ", "
      g_str_Parame = g_str_Parame & CDbl(pnl_LinAsig_Dir.Caption) & ", "
      g_str_Parame = g_str_Parame & CDbl(ipp_LinRev_Dir.Text) & ", "
      g_str_Parame = g_str_Parame & CDbl(ipp_LinNRe_Dir.Text) & ", "
      g_str_Parame = g_str_Parame & CDbl(ipp_ExpLin.Value) & ", "
      g_str_Parame = g_str_Parame & CDbl(pnl_LinTot.Caption) & ", "
      g_str_Parame = g_str_Parame & CDbl(0) & ", "                                                 'ipp_AdmFlj.Text
      g_str_Parame = g_str_Parame & CDbl(0) & ", "                                                 'ipp_ImpHip.Text
      g_str_Parame = g_str_Parame & CDbl(0) & ", "                                                 'ipp_ImpLiq.Text
      g_str_Parame = g_str_Parame & CStr(ipp_PorRet.Value) & ", "
      g_str_Parame = g_str_Parame & CStr(ipp_PorTEA.Value) & ", "
      g_str_Parame = g_str_Parame & "'" & Format(ipp_FecApr.Text, "yyyymmdd") & "', "
      g_str_Parame = g_str_Parame & "'" & Format(ipp_FecVct.Text, "yyyymmdd") & "', "
      
      If moddat_g_int_FlgGrb_1 = 2 Then
         g_str_Parame = g_str_Parame & CStr(moddat_g_str_Codigo) & ", "
      Else
         g_str_Parame = g_str_Parame & (0) & ", "
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
   
'   If moddat_g_int_FlgGrb_1 = 1 Then
      Call fs_Activa(True)
   'End If
   frm_Ges_TecPro_02.fs_Buscar
   MsgBox "Actualización realizada satisfactoriamente.", vbInformation, modgen_g_con_PltPar
   Call fs_Buscar
   Call fs_Limpia
End Sub
Private Function fs_PatEfe() As Double
Dim r_dbl_ValExpo  As Double

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "            SELECT NVL(CONLIM_PATEFE, 0) AS PATRIMONIO_EFECTIVO"
   g_str_Parame = g_str_Parame & "              FROM CTB_CONLIM "
   g_str_Parame = g_str_Parame & "             WHERE CONLIM_CODANO = (SELECT CONLIM_CODANO "
   g_str_Parame = g_str_Parame & "                                      FROM (SELECT DISTINCT CONLIM_CODANO, CONLIM_CODMES "
   g_str_Parame = g_str_Parame & "                                              FROM CTB_CONLIM "
   g_str_Parame = g_str_Parame & "                                             ORDER BY CONLIM_CODANO DESC, CONLIM_CODMES DESC) "
   g_str_Parame = g_str_Parame & "                                     WHERE ROWNUM < 2) "
   g_str_Parame = g_str_Parame & "               AND CONLIM_CODMES = (SELECT CONLIM_CODMES "
   g_str_Parame = g_str_Parame & "                                      FROM (SELECT DISTINCT CONLIM_CODANO, CONLIM_CODMES "
   g_str_Parame = g_str_Parame & "                                              FROM CTB_CONLIM "
   g_str_Parame = g_str_Parame & "                                             ORDER BY CONLIM_CODANO DESC, CONLIM_CODMES DESC)"
   g_str_Parame = g_str_Parame & "                                     WHERE ROWNUM < 2) "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Function
   End If
   
   If g_rst_Genera.BOF And g_rst_Genera.EOF Then
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      Exit Function
   End If
   
   If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
      g_rst_Genera.MoveFirst
      
      r_dbl_ValExpo = CDbl(Trim$(g_rst_Genera!PATRIMONIO_EFECTIVO)) * 0.3
      
      fs_PatEfe = Round(r_dbl_ValExpo, 2)
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
   End If
   
End Function
Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   Call gs_CentraForm(Me)
   Call fs_Limpia
   Call fs_Inicia
   Call fs_Activa(True)
   Call fs_Buscar
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   
   grd_Listad.ColWidth(0) = 0
   grd_Listad.ColWidth(1) = 1560
   grd_Listad.ColWidth(2) = 1565
   grd_Listad.ColWidth(3) = 1675
   grd_Listad.ColWidth(4) = 1610
   grd_Listad.ColWidth(5) = 1670
   grd_Listad.ColWidth(6) = 1670
   
   grd_Listad.ColAlignment(1) = flexAlignRightCenter
   grd_Listad.ColAlignment(2) = flexAlignRightCenter
   grd_Listad.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad.ColAlignment(4) = flexAlignCenterCenter
   grd_Listad.ColAlignment(5) = flexAlignCenterCenter
   grd_Listad.ColAlignment(6) = flexAlignCenterCenter
   
   Call gs_LimpiaGrid(grd_Listad)
   
   pnl_TipDoc.Caption = moddat_gf_Consulta_ParDes("118", moddat_g_int_TipDoc)
   pnl_NroDoc.Caption = moddat_g_str_NumDoc
   pnl_RazSoc.Caption = moddat_g_str_NomCli
   pnl_TipEmp.Caption = moddat_g_str_Descri
      
End Sub
Private Sub fs_Limpia()
   
   pnl_LinTot.Caption = "0.00  "
   ipp_ExpLin.Text = 0#
   pnl_LinAsig_Ind.Caption = "0.00  "
   ipp_LinRev_Ind.Text = 0#
   ipp_LinNRe_Ind.Text = 0#
   
   pnl_LinAsig_Dir.Caption = "0.00  "
   ipp_LinRev_Dir.Text = 0#
   ipp_LinNRe_Dir.Text = 0#

   ipp_PorRet.Value = 0#
   ipp_PorTEA.Value = 0#
   ipp_FecVct.Text = ""
   ipp_FecApr.Text = ""
End Sub
Private Sub fs_Activa(ByVal p_Activa As Boolean)

   ipp_ExpLin.Enabled = Not p_Activa
   ipp_LinRev_Ind.Enabled = Not p_Activa
   ipp_LinNRe_Ind.Enabled = Not p_Activa
   ipp_LinRev_Dir.Enabled = Not p_Activa
   ipp_LinNRe_Dir.Enabled = Not p_Activa
   ipp_PorRet.Enabled = Not p_Activa
   ipp_PorTEA.Enabled = Not p_Activa
   ipp_FecApr.Enabled = Not p_Activa
   ipp_FecVct.Enabled = Not p_Activa
   
   cmd_Grabar.Enabled = Not p_Activa
   cmd_Cancel.Enabled = Not p_Activa
   
End Sub
Private Sub fs_Buscar()

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT ETEHIS_NUMITE, ETEHIS_LINASI_IND, ETEHIS_LINASI_DIR, ETEHIS_PORRET, ETEHIS_PORTEA, ETEHIS_FECAPR, ETEHIS_FECVCT "
   g_str_Parame = g_str_Parame & "   FROM TPR_ETEHIS A "
   g_str_Parame = g_str_Parame & "  WHERE ETEHIS_TIPDOC = " & CStr(moddat_g_int_TipDoc) & " "
   g_str_Parame = g_str_Parame & "    AND ETEHIS_NUMDOC = '" & CStr(moddat_g_str_NumDoc) & "' "
   g_str_Parame = g_str_Parame & "  ORDER BY ETEHIS_NUMITE DESC " 'SEGFECCRE, SEGHORCRE
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
       
   grd_Listad.Redraw = False
   Call gs_LimpiaGrid(grd_Listad)
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
     g_rst_Princi.Close
     Set g_rst_Princi = Nothing
      grd_Listad.Redraw = True
     Exit Sub
   End If

   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
          
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         
         grd_Listad.Col = 0
         grd_Listad.Text = g_rst_Princi!ETEHIS_NUMITE
         
         grd_Listad.Col = 1
         If Not IsNull(g_rst_Princi!ETEHIS_LINASI_IND) Then
            grd_Listad.Text = Format(CStr(g_rst_Princi!ETEHIS_LINASI_IND), "###,###,###,##0.00")
         Else
            grd_Listad.Text = Format(CStr(0), "###,###,###,##0.00")
         End If
         
         grd_Listad.Col = 2
         If Not IsNull(g_rst_Princi!ETEHIS_LINASI_DIR) Then
            grd_Listad.Text = Format(CStr(g_rst_Princi!ETEHIS_LINASI_DIR), "###,###,###,##0.00")
         Else
            grd_Listad.Text = Format(CStr(0), "###,###,###,##0.00")
         End If
         
         grd_Listad.Col = 3
         If Not IsNull(g_rst_Princi!ETEHIS_PORRET) Then
            grd_Listad.Text = Format(CStr(g_rst_Princi!ETEHIS_PORRET), "0.00")
         Else
            grd_Listad.Text = Format(CStr(0#), "0.00")
         End If
         
         grd_Listad.Col = 4
         If Not IsNull(g_rst_Princi!ETEHIS_PORTEA) Then
            grd_Listad.Text = Format(CStr(g_rst_Princi!ETEHIS_PORTEA), "0.00")
         Else
            grd_Listad.Text = Format(CStr(0#), "0.00")
         End If
         
         grd_Listad.Col = 5
         If Not IsNull(g_rst_Princi!ETEHIS_FECAPR) Then
            grd_Listad.Text = Format(gf_FormatoFecha(CStr(g_rst_Princi!ETEHIS_FECAPR)), "dd/mm/yyyy")
         Else
            grd_Listad.Text = ""
         End If
         
         grd_Listad.Col = 6
         If Not IsNull(g_rst_Princi!ETEHIS_FECVCT) Then
            grd_Listad.Text = Format(gf_FormatoFecha(CStr(g_rst_Princi!ETEHIS_FECVCT)), "dd/mm/yyyy")
         Else
            grd_Listad.Text = ""
         End If
         
                 
         g_rst_Princi.MoveNext
      Loop
   End If
   
   grd_Listad.Redraw = True
   Call gs_UbiIniGrid(grd_Listad)
   
End Sub

Private Sub fs_GenExc()
Dim r_obj_Excel         As Excel.Application
Dim r_str_ParAux        As String
Dim r_str_FecRpt        As String
Dim r_int_Contad        As Integer
Dim r_int_NroFil        As Integer
Dim r_int_NoFlLi        As Integer
   
    r_int_NroFil = 8
    
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Cells(2, 2) = "REPORTE DE ENTIDAD TECNICA - HISTORICO"
      .Range(.Cells(2, 2), .Cells(2, 9)).Merge
      .Range(.Cells(2, 2), .Cells(2, 9)).Font.Bold = True
      .Range(.Cells(2, 2), .Cells(2, 9)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(2, 2), .Cells(2, 9)).Font.Size = 14
      
      .Cells(4, 2) = "TIPO DE DOCUMENTO"
      .Cells(4, 3) = Trim(pnl_TipDoc.Caption)
      .Cells(5, 2) = "NRO. DOCUMENTO"
      .Cells(5, 3) = "'" & Trim(pnl_NroDoc.Caption)
      .Cells(6, 2) = "RAZÓN SOCIAL"
      .Cells(6, 3) = Trim(pnl_RazSoc.Caption)
      .Range(.Cells(3, 2), .Cells(6, 2)).Font.Bold = True
      
      .Cells(r_int_NroFil, 2) = "LINEA ASIGNADA INDIRECTA"
      .Range(.Cells(r_int_NroFil, 2), .Cells(r_int_NroFil + 1, 2)).Merge
      .Cells(r_int_NroFil, 3) = "LINEA ASIGNADA DIRECTA"
      .Range(.Cells(r_int_NroFil, 3), .Cells(r_int_NroFil + 1, 3)).Merge
      .Cells(r_int_NroFil, 4) = "RETENCION (%)"
      .Range(.Cells(r_int_NroFil, 4), .Cells(r_int_NroFil + 1, 4)).Merge
      .Cells(r_int_NroFil, 5) = "TEA (%)"
      .Range(.Cells(r_int_NroFil, 5), .Cells(r_int_NroFil + 1, 5)).Merge
      .Cells(r_int_NroFil, 6) = "FECHA VENCIMIENTO"
      .Range(.Cells(r_int_NroFil, 6), .Cells(r_int_NroFil + 1, 6)).Merge
      
      .Range(.Cells(r_int_NroFil, 2), .Cells(r_int_NroFil + 1, 6)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(r_int_NroFil, 2), .Cells(r_int_NroFil + 1, 6)).Font.Bold = True
      .Range(.Cells(r_int_NroFil, 2), .Cells(r_int_NroFil + 1, 6)).HorizontalAlignment = xlHAlignCenter
      
      .Columns("A").ColumnWidth = 1
      .Columns("B").ColumnWidth = 20
      .Columns("B").NumberFormat = "###,###,###,##0.00"
      .Columns("B").HorizontalAlignment = xlHAlignRight
      .Columns("C").ColumnWidth = 20
      .Columns("C").NumberFormat = "###,###,###,##0.00"
      .Columns("C").HorizontalAlignment = xlHAlignRight
      .Columns("D").ColumnWidth = 17
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("D").NumberFormat = "0.00"
      .Columns("E").ColumnWidth = 17
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("E").NumberFormat = "0.00"
      .Columns("F").ColumnWidth = 17
      .Columns("F").HorizontalAlignment = xlHAlignCenter
               
      With .Range(.Cells(8, 2), .Cells(9, 6))
         .HorizontalAlignment = xlCenter
         .VerticalAlignment = xlCenter
         .WrapText = True
      End With
      
      .Range(.Cells(1, 1), .Cells(99, 99)).Font.Name = "Calibri"
      .Range(.Cells(3, 1), .Cells(99, 99)).Font.Size = 11
      
      r_int_NroFil = r_int_NroFil + 2
      
      For r_int_NoFlLi = 0 To grd_Listad.Rows - 1
      
         .Cells(r_int_NroFil, 2) = grd_Listad.TextMatrix(r_int_NoFlLi, 0)
         .Cells(r_int_NroFil, 3) = grd_Listad.TextMatrix(r_int_NoFlLi, 1)
         .Cells(r_int_NroFil, 4) = grd_Listad.TextMatrix(r_int_NoFlLi, 2)
         .Cells(r_int_NroFil, 5) = grd_Listad.TextMatrix(r_int_NoFlLi, 3)
         .Cells(r_int_NroFil, 6) = "'" & grd_Listad.TextMatrix(r_int_NoFlLi, 4)
         
         r_int_NroFil = r_int_NroFil + 1
      Next r_int_NoFlLi
           
      With .Range(.Cells(8, 2), .Cells(r_int_NroFil - 1, 6))
          .Borders(xlEdgeLeft).LineStyle = xlContinuous
          .Borders(xlEdgeTop).LineStyle = xlContinuous
          .Borders(xlEdgeBottom).LineStyle = xlContinuous
          .Borders(xlEdgeRight).LineStyle = xlContinuous
          .Borders(xlInsideVertical).LineStyle = xlContinuous
          .Borders(xlInsideHorizontal).LineStyle = xlContinuous
      End With
      
      With .Range(.Cells(4, 2), .Cells(6, 5))
          .HorizontalAlignment = xlLeft
          .VerticalAlignment = xlBottom
          .WrapText = False
      End With
      
   End With
   
   r_obj_Excel.Visible = True
End Sub

Private Sub ipp_ExpLin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_PorRet)
   End If
End Sub

Private Sub ipp_FecApr_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecVct)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub ipp_FecApr_LostFocus()
   If ipp_FecApr.Value > 0 Then
      ipp_FecVct.Text = Format(CDate(DateAdd("D", 360, ipp_FecApr.Value)), "DD/MM/YYYY")
   Else
      ipp_FecVct.Value = ""
   End If
End Sub

Private Sub ipp_FecVct_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   Else
      KeyAscii = 0
   End If
End Sub

Private Sub ipp_LinNRe_Dir_Change()
   Call fs_Calcular_LinApr_Dir
End Sub

Private Sub ipp_LinNRe_Dir_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_ExpLin)
   End If
End Sub

Private Sub ipp_LinRev_Dir_Change()
   Call fs_Calcular_LinApr_Dir
End Sub

Private Sub ipp_LinRev_Dir_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_LinNRe_Dir)
   End If
End Sub

Private Sub ipp_LinNRe_Ind_Change()
   Call fs_Calcular_LinApr_Ind
End Sub

Private Sub ipp_LinNRe_Ind_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_LinRev_Dir)
   End If
End Sub

Private Sub ipp_LinRev_Ind_Change()
   Call fs_Calcular_LinApr_Ind
End Sub

Private Sub ipp_LinRev_Ind_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_LinNRe_Ind)
   End If
End Sub

Private Sub ipp_PorRet_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_PorTEA)
   End If
End Sub

Private Sub ipp_PorTEA_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecApr)
   End If
End Sub
Private Sub fs_Calcular_LinApr_Ind()
      pnl_LinAsig_Ind.Caption = Format(CDbl(ipp_LinRev_Ind.Value + CDbl(ipp_LinNRe_Ind.Value)), "###,##0.00") & " "
      If CDbl(pnl_LinAsig_Ind.Caption) <= 0 Then
        pnl_LinAsig_Ind.Caption = "0.00  "
      End If
      pnl_LinTot.Caption = Format(CDbl(pnl_LinAsig_Ind.Caption) + CDbl(pnl_LinAsig_Dir.Caption), "###,##0.00") & " "
End Sub
Private Sub fs_Calcular_LinApr_Dir()
      pnl_LinAsig_Dir.Caption = Format(CDbl(ipp_LinRev_Dir.Value) + CDbl(ipp_LinNRe_Dir.Value), "###,##0.00") & " "
      If CDbl(pnl_LinAsig_Dir.Caption) <= 0 Then
        pnl_LinAsig_Dir.Caption = "0.00  "
      End If
      pnl_LinTot.Caption = Format(CDbl(pnl_LinAsig_Ind.Caption) + CDbl(pnl_LinAsig_Dir.Caption), "###,##0.00") & " "
End Sub
Private Function fs_Validar_MtoLinAsi() As Boolean
   fs_Validar_MtoLinAsi = False

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "   SELECT (NVL(MAEETE_LINASI_IND,0) + NVL(MAEETE_LINASI_DIR,0)) AS LINASI, "
   g_str_Parame = g_str_Parame & "           NVL(SUM(CASE WHEN MAECFI_CODPRD <> '008' THEN MAECFI_GARFIA "
   g_str_Parame = g_str_Parame & "                   ELSE CASE WHEN MAECFI_CODMOD <> '002' THEN MAECFI_IMPFIA ELSE 0 END "
   g_str_Parame = g_str_Parame & "                    END),0) AS CARTA_FIANZA"
   g_str_Parame = g_str_Parame & "     FROM TPR_MAEETE"
   g_str_Parame = g_str_Parame & "          LEFT JOIN TPR_MAECFI ON MAEETE_TIPDOC = MAECFI_TIPDOC AND MAEETE_NUMDOC = MAECFI_NUMDOC AND MAECFI_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "    WHERE MAEETE_TIPDOC = " & Trim(CStr(Mid(pnl_TipDoc.Caption, 1, 2))) & " "
   g_str_Parame = g_str_Parame & "      AND MAEETE_NUMDOC = '" & Trim(CStr(pnl_NroDoc.Caption)) & "'"
   g_str_Parame = g_str_Parame & "    GROUP BY MAECFI_TIPDOC, MAECFI_NUMDOC, MAEETE_LINASI_IND, MAEETE_LINASI_DIR "
    
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
     Exit Function
   End If
   
   If g_rst_GenAux.BOF And g_rst_GenAux.EOF Then
      g_rst_GenAux.Close
      Set g_rst_GenAux = Nothing
      fs_Validar_MtoLinAsi = True
      Exit Function
   End If
     
   If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
      If moddat_g_int_FlgGrb = 1 Then
         If CDbl(g_rst_GenAux!CARTA_FIANZA) + CDbl(pnl_LinAsig_Ind.Caption) <= CDbl(g_rst_GenAux!LINASI) Then
            fs_Validar_MtoLinAsi = True
         End If
      Else
         If CDbl(pnl_LinAsig_Ind.Caption) >= CDbl(g_rst_GenAux!CARTA_FIANZA) Then
            fs_Validar_MtoLinAsi = True
         End If
      End If
   End If
End Function
