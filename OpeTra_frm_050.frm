VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_GasAdm_04 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8940
   ClientLeft      =   1845
   ClientTop       =   1155
   ClientWidth     =   10020
   Icon            =   "OpeTra_frm_050.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8940
   ScaleWidth      =   10020
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel2 
      Height          =   8925
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10005
      _Version        =   65536
      _ExtentX        =   17648
      _ExtentY        =   15743
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
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   9915
         _Version        =   65536
         _ExtentX        =   17489
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
            TabIndex        =   2
            Top             =   30
            Width           =   6435
            _Version        =   65536
            _ExtentX        =   11351
            _ExtentY        =   1032
            _StockProps     =   15
            Caption         =   "Gastos de Cierre - Conciliaci�n"
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
            Picture         =   "OpeTra_frm_050.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   1425
         Left            =   30
         TabIndex        =   3
         Top             =   750
         Width           =   9915
         _Version        =   65536
         _ExtentX        =   17489
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
         Begin Threed.SSPanel pnl_NumSol 
            Height          =   315
            Left            =   1620
            TabIndex        =   4
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
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_NomCli 
            Height          =   315
            Left            =   1620
            TabIndex        =   5
            Top             =   720
            Width           =   8235
            _Version        =   65536
            _ExtentX        =   14526
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
            Left            =   1620
            TabIndex        =   6
            Top             =   1050
            Width           =   2925
            _Version        =   65536
            _ExtentX        =   5159
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
         Begin Threed.SSPanel pnl_DocIde 
            Height          =   315
            Left            =   1620
            TabIndex        =   7
            Top             =   390
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
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_Situac 
            Height          =   315
            Left            =   6930
            TabIndex        =   38
            Top             =   60
            Width           =   2925
            _Version        =   65536
            _ExtentX        =   5159
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
         Begin VB.Label Label9 
            Caption         =   "Situaci�n Solicitud:"
            Height          =   315
            Left            =   5430
            TabIndex        =   39
            Top             =   60
            Width           =   1395
         End
         Begin VB.Label Label3 
            Caption         =   "Nombre Cliente:"
            Height          =   315
            Left            =   60
            TabIndex        =   11
            Top             =   720
            Width           =   1125
         End
         Begin VB.Label Label2 
            Caption         =   "Nro. de Solicitud:"
            Height          =   315
            Left            =   60
            TabIndex        =   10
            Top             =   60
            Width           =   1395
         End
         Begin VB.Label Label1 
            Caption         =   "Moneda:"
            Height          =   315
            Left            =   60
            TabIndex        =   9
            Top             =   1050
            Width           =   1125
         End
         Begin VB.Label Label12 
            Caption         =   "DOI Cliente:"
            Height          =   315
            Left            =   60
            TabIndex        =   8
            Top             =   390
            Width           =   1395
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   2745
         Left            =   30
         TabIndex        =   12
         Top             =   3030
         Width           =   9915
         _Version        =   65536
         _ExtentX        =   17489
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
         Begin Threed.SSPanel pnl_TotSal 
            Height          =   315
            Left            =   8490
            TabIndex        =   21
            Top             =   2370
            Width           =   1125
            _Version        =   65536
            _ExtentX        =   1984
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
            ForeColor       =   16777215
            BackColor       =   255
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_TotGas 
            Height          =   315
            Left            =   7380
            TabIndex        =   20
            Top             =   2370
            Width           =   1125
            _Version        =   65536
            _ExtentX        =   1984
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
            ForeColor       =   16777215
            BackColor       =   255
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_TotCob 
            Height          =   315
            Left            =   6270
            TabIndex        =   13
            Top             =   2370
            Width           =   1125
            _Version        =   65536
            _ExtentX        =   1984
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "0.00 "
            ForeColor       =   16777215
            BackColor       =   255
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            Alignment       =   4
         End
         Begin MSFlexGridLib.MSFlexGrid grd_GasAdm 
            Height          =   2025
            Left            =   30
            TabIndex        =   14
            Top             =   330
            Width           =   9855
            _ExtentX        =   17383
            _ExtentY        =   3572
            _Version        =   393216
            Rows            =   21
            Cols            =   5
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   49152
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel11 
            Height          =   285
            Left            =   60
            TabIndex        =   15
            Top             =   60
            Width           =   6225
            _Version        =   65536
            _ExtentX        =   10980
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Concepto"
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
            Left            =   6270
            TabIndex        =   16
            Top             =   60
            Width           =   1125
            _Version        =   65536
            _ExtentX        =   1984
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Cobrado"
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
            Left            =   7380
            TabIndex        =   18
            Top             =   60
            Width           =   1125
            _Version        =   65536
            _ExtentX        =   1984
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Gastado"
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
            Left            =   8490
            TabIndex        =   19
            Top             =   60
            Width           =   1125
            _Version        =   65536
            _ExtentX        =   1984
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Saldo"
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
            Caption         =   "Total:"
            Height          =   285
            Left            =   5250
            TabIndex        =   17
            Top             =   2370
            Width           =   975
         End
      End
      Begin Threed.SSPanel SSPanel14 
         Height          =   765
         Left            =   30
         TabIndex        =   22
         Top             =   5820
         Width           =   9915
         _Version        =   65536
         _ExtentX        =   17489
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
         Begin VB.CommandButton cmd_EdiIte 
            Height          =   675
            Left            =   9210
            Picture         =   "OpeTra_frm_050.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Edici�n de Datos"
            Top             =   30
            Width           =   675
         End
         Begin Threed.SSPanel pnl_ITFImp 
            Height          =   315
            Left            =   1620
            TabIndex        =   41
            Top             =   60
            Width           =   1665
            _Version        =   65536
            _ExtentX        =   2937
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   16777215
            BackColor       =   192
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            Alignment       =   4
         End
         Begin Threed.SSPanel pnl_TotImp 
            Height          =   315
            Left            =   1620
            TabIndex        =   42
            Top             =   390
            Width           =   1665
            _Version        =   65536
            _ExtentX        =   2937
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   16777215
            BackColor       =   192
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            Alignment       =   4
         End
         Begin VB.Label Label10 
            Caption         =   "ITF:"
            Height          =   285
            Left            =   60
            TabIndex        =   44
            Top             =   60
            Width           =   975
         End
         Begin VB.Label Label11 
            Caption         =   "Importe Devoluci�n:"
            Height          =   285
            Left            =   60
            TabIndex        =   43
            Top             =   390
            Width           =   1455
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   1425
         Left            =   30
         TabIndex        =   24
         Top             =   6630
         Width           =   9915
         _Version        =   65536
         _ExtentX        =   17489
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
         Begin EditLib.fpDoubleSingle ipp_Import 
            Height          =   315
            Left            =   1620
            TabIndex        =   25
            Top             =   720
            Width           =   1365
            _Version        =   196608
            _ExtentX        =   2408
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
         Begin Threed.SSPanel pnl_NomGas 
            Height          =   315
            Left            =   1620
            TabIndex        =   28
            Top             =   60
            Width           =   8235
            _Version        =   65536
            _ExtentX        =   14526
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
         Begin Threed.SSPanel pnl_ImpCob 
            Height          =   315
            Left            =   1620
            TabIndex        =   29
            Top             =   390
            Width           =   1365
            _Version        =   65536
            _ExtentX        =   2408
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
         Begin Threed.SSPanel pnl_SalGas 
            Height          =   315
            Left            =   1620
            TabIndex        =   32
            Top             =   1050
            Width           =   1365
            _Version        =   65536
            _ExtentX        =   2408
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
         Begin VB.Label Label7 
            Caption         =   "Saldo:"
            Height          =   285
            Left            =   60
            TabIndex        =   31
            Top             =   1050
            Width           =   1485
         End
         Begin VB.Label Label6 
            Caption         =   "Importe Gastado:"
            Height          =   285
            Left            =   60
            TabIndex        =   30
            Top             =   720
            Width           =   1485
         End
         Begin VB.Label Label4 
            Caption         =   "Concepto de Gasto:"
            Height          =   285
            Left            =   60
            TabIndex        =   27
            Top             =   60
            Width           =   1485
         End
         Begin VB.Label Label8 
            Caption         =   "Importe Cobrado:"
            Height          =   285
            Left            =   60
            TabIndex        =   26
            Top             =   390
            Width           =   1485
         End
      End
      Begin Threed.SSPanel SSPanel39 
         Height          =   765
         Left            =   30
         TabIndex        =   33
         Top             =   2220
         Width           =   9915
         _Version        =   65536
         _ExtentX        =   17489
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
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   9210
            Picture         =   "OpeTra_frm_050.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   675
            Left            =   8520
            Picture         =   "OpeTra_frm_050.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   40
            ToolTipText     =   "Edici�n de Datos"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel12 
         Height          =   765
         Left            =   30
         TabIndex        =   35
         Top             =   8100
         Width           =   9915
         _Version        =   65536
         _ExtentX        =   17489
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
            Left            =   8520
            Picture         =   "OpeTra_frm_050.frx":0EA4
            Style           =   1  'Graphical
            TabIndex        =   37
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Cancel 
            Height          =   675
            Left            =   9210
            Picture         =   "OpeTra_frm_050.frx":12E6
            Style           =   1  'Graphical
            TabIndex        =   36
            ToolTipText     =   "Cancelar"
            Top             =   30
            Width           =   675
         End
      End
   End
End
Attribute VB_Name = "frm_GasAdm_04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_dbl_PorITF     As Double
Dim l_str_CodGas     As String

Private Sub cmd_Cancel_Click()
   Call fs_LimpiaItem
   Call fs_ActivaItem(False)
   
   Call gs_SetFocus(grd_GasAdm)
End Sub

Private Sub cmd_EdiIte_Click()
   If grd_GasAdm.Rows = 0 Then
      Exit Sub
   End If
   
   grd_GasAdm.Col = 4
   l_str_CodGas = grd_GasAdm.Text
   
   grd_GasAdm.Col = 0
   pnl_NomGas.Caption = grd_GasAdm.Text
   
   grd_GasAdm.Col = 1
   pnl_ImpCob.Caption = Format(grd_GasAdm.Text, "###,##0.00") & " "
   
   grd_GasAdm.Col = 2
   ipp_Import.Value = CDbl(grd_GasAdm.Text)
   
   Call gs_RefrescaGrid(grd_GasAdm)
   
   Call fs_ActivaItem(True)
   Call gs_SetFocus(ipp_Import)
End Sub

Private Sub cmd_Grabar_Click()
   If MsgBox("�Est� seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If

   Screen.MousePointer = 11
   
   'Grabar Datos
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      Screen.MousePointer = 11
      
      Call moddat_gs_FecSis
      
      g_str_Parame = "USP_TRA_GASADM_CONCILIA ("
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumSol & "', "
      g_str_Parame = g_str_Parame & l_str_CodGas & ", "
      g_str_Parame = g_str_Parame & "2, "
      g_str_Parame = g_str_Parame & CStr(CDbl(ipp_Import.Value)) & ", "
      g_str_Parame = g_str_Parame & CStr(CDbl(pnl_SalGas.Caption)) & ", "
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
         If MsgBox("No se pudo completar la grabaci�n de los datos. �Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
      
      Screen.MousePointer = 0
   Loop
   
   Call fs_Buscar_GasAdm
   
   Call fs_LimpiaItem
   Call fs_ActivaItem(False)
   
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt

   pnl_NumSol.Caption = gf_Formato_NumSol(moddat_g_str_NumSol)
   pnl_DocIde.Caption = CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc
   pnl_NomCli.Caption = moddat_g_str_NomCli
   pnl_Moneda.Caption = moddat_gf_Consulta_ParDes("204", CStr(moddat_g_int_TipMon))
   pnl_Situac.Caption = moddat_g_str_Descri
   
   Call fs_Inicia
   
   'Buscando Gastos Administrativos
   Call fs_Buscar_GasAdm
   
   Call fs_LimpiaItem
   Call fs_ActivaItem(False)
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Dim r_int_Contad     As Integer
   
   grd_GasAdm.ColWidth(0) = 6215
   grd_GasAdm.ColWidth(1) = 1115
   grd_GasAdm.ColWidth(2) = 1115
   grd_GasAdm.ColWidth(3) = 1115
   grd_GasAdm.ColWidth(4) = 0
   
   
   grd_GasAdm.ColAlignment(0) = flexAlignLeftCenter
   grd_GasAdm.ColAlignment(1) = flexAlignRightCenter
   grd_GasAdm.ColAlignment(2) = flexAlignRightCenter
   grd_GasAdm.ColAlignment(3) = flexAlignRightCenter

   'Obteniendo ITF
   l_dbl_PorITF = opecaj_gf_Consulta_ITF(Format(CDate(moddat_g_str_FecSis), "yyyymmdd"), 1)
End Sub

Private Sub fs_Buscar_GasAdm()
   Dim r_dbl_TotCob     As Double
   Dim r_dbl_TotGas     As Double
   Dim r_dbl_TotSal     As Double
   
   r_dbl_TotCob = 0
   r_dbl_TotGas = 0
   r_dbl_TotSal = 0
   
   pnl_TotCob.Caption = "0.00 "
   pnl_TotGas.Caption = "0.00 "
   pnl_TotSal.Caption = "0.00 "
   
   Call gs_LimpiaGrid(grd_GasAdm)
   
   g_str_Parame = "SELECT * FROM TRA_GASADM WHERE "
   g_str_Parame = g_str_Parame & "GASADM_NUMSOL = '" & moddat_g_str_NumSol & "'"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      grd_GasAdm.Redraw = False
   
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         grd_GasAdm.Rows = grd_GasAdm.Rows + 1
         grd_GasAdm.Row = grd_GasAdm.Rows - 1
      
         'Buscando Descripci�n de Gastos Administrativos
         grd_GasAdm.Col = 0
         If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), moddat_g_str_CodPrd, moddat_g_str_CodSub, "007", Format(g_rst_Princi!GASADM_CODGAS, "00") & CStr(g_rst_Princi!GASADM_TIPMON)) Then
            grd_GasAdm.Text = Trim(moddat_g_arr_Genera(1).Genera_Nombre)
         End If
      
         grd_GasAdm.Col = 1
         grd_GasAdm.Text = Format(g_rst_Princi!GASADM_PAGIMP, "###,###,##0.00")
      
         grd_GasAdm.Col = 2
         grd_GasAdm.Text = Format(g_rst_Princi!GASADM_GASIMP, "###,###,##0.00")
      
         grd_GasAdm.Col = 3
         grd_GasAdm.Text = Format(g_rst_Princi!GASADM_GASSAL, "###,###,##0.00")
      
         grd_GasAdm.Col = 4
         grd_GasAdm.Text = Format(g_rst_Princi!GASADM_CODGAS, "00")
      
         r_dbl_TotCob = r_dbl_TotCob + g_rst_Princi!GASADM_PAGIMP
         r_dbl_TotGas = r_dbl_TotGas + g_rst_Princi!GASADM_GASIMP
         r_dbl_TotSal = r_dbl_TotSal + g_rst_Princi!GASADM_GASSAL
      
         g_rst_Princi.MoveNext
      Loop
      
      grd_GasAdm.Redraw = True
   End If
   
   pnl_TotCob.Caption = Format(r_dbl_TotCob, "###,###,##0.00") & " "
   pnl_TotGas.Caption = Format(r_dbl_TotGas, "###,###,##0.00") & " "
   pnl_TotSal.Caption = Format(r_dbl_TotSal, "###,###,##0.00") & " "
   
   pnl_ITFImp.Caption = Format(CDbl(gf_Truncar_Numero(CDbl(pnl_TotSal.Caption) * (l_dbl_PorITF / 100), 2)), "###,###,##0.00") & " "
   pnl_TotImp.Caption = Format(CDbl(pnl_TotSal.Caption) - CDbl(pnl_ITFImp.Caption), "###,###,##0.00") & " "
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call gs_UbiIniGrid(grd_GasAdm)
   Call gs_SetFocus(grd_GasAdm)
End Sub

Private Sub fs_ActivaItem(ByVal p_Activa As Integer)
   ipp_Import.Enabled = p_Activa
   cmd_Grabar.Enabled = p_Activa
   cmd_Cancel.Enabled = p_Activa
   
   cmd_EdiIte.Enabled = Not p_Activa
   grd_GasAdm.Enabled = Not p_Activa
   cmd_Imprim.Enabled = Not p_Activa
End Sub

Private Sub fs_LimpiaItem()
   pnl_NomGas.Caption = ""
   pnl_ImpCob.Caption = "0.00 "
   ipp_Import.Value = 0
   pnl_SalGas.Caption = "0.00 "
End Sub

Private Sub grd_GasAdm_DblClick()
   Call cmd_EdiIte_Click
End Sub

Private Sub grd_GasAdm_SelChange()
   If grd_GasAdm.Rows > 2 Then
      grd_GasAdm.RowSel = grd_GasAdm.Row
   End If
End Sub

Private Sub ipp_Import_Change()
   pnl_SalGas.Caption = Format(CDbl(pnl_ImpCob.Caption) - CDbl(ipp_Import.Text), "###,##0.00") & " "
End Sub

Private Sub ipp_Import_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   End If
End Sub
