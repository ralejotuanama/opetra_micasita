VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_Con_PrePgo_03 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form4"
   ClientHeight    =   10020
   ClientLeft      =   1785
   ClientTop       =   3600
   ClientWidth     =   11670
   Icon            =   "OpeTra_frm_330.frx":0000
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10020
   ScaleWidth      =   11670
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel5 
      Height          =   3345
      Left            =   15
      TabIndex        =   0
      Top             =   5160
      Width           =   11625
      _Version        =   65536
      _ExtentX        =   20505
      _ExtentY        =   5900
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
      Begin VB.ComboBox cmb_RedPlz 
         Height          =   315
         Left            =   12870
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "1 AÑO"
         Top             =   2025
         Visible         =   0   'False
         Width           =   1170
      End
      Begin EditLib.fpDateTime ipp_FecPre 
         Height          =   315
         Left            =   1530
         TabIndex        =   3
         Top             =   390
         Width           =   1320
         _Version        =   196608
         _ExtentX        =   2328
         _ExtentY        =   556
         Enabled         =   0   'False
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
      Begin Threed.SSPanel pnl_NuevoSaldoTNC 
         Height          =   315
         Left            =   7440
         TabIndex        =   4
         Top             =   2610
         Width           =   1230
         _Version        =   65536
         _ExtentX        =   2170
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "0.00 "
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
         RoundedCorners  =   0   'False
         Font3D          =   2
         Alignment       =   4
      End
      Begin Threed.SSPanel pnl_NuevoSaldoTC 
         Height          =   315
         Left            =   7440
         TabIndex        =   5
         Top             =   2940
         Width           =   1230
         _Version        =   65536
         _ExtentX        =   2170
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "0.00 "
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
         RoundedCorners  =   0   'False
         Font3D          =   2
         Alignment       =   4
      End
      Begin Threed.SSPanel pnl_NuevaCuota 
         Height          =   315
         Left            =   10290
         TabIndex        =   6
         Top             =   2610
         Width           =   1230
         _Version        =   65536
         _ExtentX        =   2170
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "0.00 "
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
         RoundedCorners  =   0   'False
         Font3D          =   2
         Alignment       =   4
      End
      Begin Threed.SSPanel pnl_SaldoTNC1 
         Height          =   315
         Left            =   1530
         TabIndex        =   7
         Top             =   720
         Width           =   1320
         _Version        =   65536
         _ExtentX        =   2328
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "0.00 "
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
         RoundedCorners  =   0   'False
         Font3D          =   2
         Alignment       =   4
      End
      Begin Threed.SSPanel pnl_SaldoTC1 
         Height          =   315
         Left            =   4400
         TabIndex        =   8
         Top             =   720
         Width           =   1350
         _Version        =   65536
         _ExtentX        =   2381
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "0.00 "
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
         RoundedCorners  =   0   'False
         Font3D          =   2
         Alignment       =   4
      End
      Begin Threed.SSPanel pnl_UltPagTNC 
         Height          =   315
         Left            =   1530
         TabIndex        =   9
         Top             =   1080
         Width           =   1320
         _Version        =   65536
         _ExtentX        =   2328
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "31/12/2011"
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
         RoundedCorners  =   0   'False
         Font3D          =   2
      End
      Begin Threed.SSPanel pnl_UltPagTC 
         Height          =   315
         Left            =   1530
         TabIndex        =   10
         Top             =   1410
         Width           =   1320
         _Version        =   65536
         _ExtentX        =   2328
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "31/12/2011"
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
         RoundedCorners  =   0   'False
         Font3D          =   2
      End
      Begin Threed.SSPanel pnl_DiasTNC 
         Height          =   315
         Left            =   4400
         TabIndex        =   11
         Top             =   1080
         Width           =   1350
         _Version        =   65536
         _ExtentX        =   2381
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "0 "
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
         RoundedCorners  =   0   'False
         Font3D          =   2
         Alignment       =   4
      End
      Begin Threed.SSPanel pnl_DiasTC 
         Height          =   315
         Left            =   4400
         TabIndex        =   12
         Top             =   1410
         Width           =   1350
         _Version        =   65536
         _ExtentX        =   2381
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "0 "
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
         RoundedCorners  =   0   'False
         Font3D          =   2
         Alignment       =   4
      End
      Begin Threed.SSPanel pnl_SaldoTNC2 
         Height          =   315
         Left            =   1530
         TabIndex        =   13
         Top             =   2610
         Width           =   1320
         _Version        =   65536
         _ExtentX        =   2328
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "0.00 "
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
         RoundedCorners  =   0   'False
         Font3D          =   2
         Alignment       =   4
      End
      Begin Threed.SSPanel pnl_SaldoTC2 
         Height          =   315
         Left            =   1530
         TabIndex        =   14
         Top             =   2940
         Width           =   1320
         _Version        =   65536
         _ExtentX        =   2328
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "0.00 "
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
         RoundedCorners  =   0   'False
         Font3D          =   2
         Alignment       =   4
      End
      Begin Threed.SSPanel pnl_MtoApl 
         Height          =   315
         Left            =   1530
         TabIndex        =   15
         Top             =   2070
         Width           =   1320
         _Version        =   65536
         _ExtentX        =   2328
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "0.00 "
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
         RoundedCorners  =   0   'False
         Font3D          =   2
         Alignment       =   4
      End
      Begin Threed.SSPanel pnl_Val_AsgInm 
         Height          =   315
         Left            =   7440
         TabIndex        =   16
         Top             =   720
         Width           =   1230
         _Version        =   65536
         _ExtentX        =   2170
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "0.00 "
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
         RoundedCorners  =   0   'False
         Font3D          =   2
         Alignment       =   4
      End
      Begin Threed.SSPanel pnl_Mto_Deposito 
         Height          =   330
         Left            =   10290
         TabIndex        =   48
         Top             =   375
         Width           =   1230
         _Version        =   65536
         _ExtentX        =   2170
         _ExtentY        =   582
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
      Begin Threed.SSPanel pnl_InteresTNC 
         Height          =   315
         Left            =   7440
         TabIndex        =   49
         Top             =   1080
         Width           =   1230
         _Version        =   65536
         _ExtentX        =   2170
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "0.00 "
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
         RoundedCorners  =   0   'False
         Font3D          =   2
         Alignment       =   4
      End
      Begin Threed.SSPanel pnl_InteresTC 
         Height          =   315
         Left            =   7440
         TabIndex        =   50
         Top             =   1410
         Width           =   1230
         _Version        =   65536
         _ExtentX        =   2170
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "0.00 "
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
         RoundedCorners  =   0   'False
         Font3D          =   2
         Alignment       =   4
      End
      Begin Threed.SSPanel pnl_SegDes 
         Height          =   315
         Left            =   1530
         TabIndex        =   51
         Top             =   1740
         Width           =   1320
         _Version        =   65536
         _ExtentX        =   2328
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "0.00 "
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
         RoundedCorners  =   0   'False
         Font3D          =   2
         Alignment       =   4
      End
      Begin Threed.SSPanel pnl_SegInm 
         Height          =   315
         Left            =   4400
         TabIndex        =   52
         Top             =   1740
         Width           =   1350
         _Version        =   65536
         _ExtentX        =   2381
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "0.00 "
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
         RoundedCorners  =   0   'False
         Font3D          =   2
         Alignment       =   4
      End
      Begin Threed.SSPanel pnl_MontoITF 
         Height          =   315
         Left            =   7440
         TabIndex        =   53
         Top             =   1740
         Width           =   1230
         _Version        =   65536
         _ExtentX        =   2170
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "0.00 "
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
         RoundedCorners  =   0   'False
         Font3D          =   2
         Alignment       =   4
      End
      Begin Threed.SSPanel pnl_CuoPend 
         Height          =   315
         Left            =   10290
         TabIndex        =   54
         Top             =   720
         Width           =   1230
         _Version        =   65536
         _ExtentX        =   2170
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "0.00 "
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
         RoundedCorners  =   0   'False
         Font3D          =   2
         Alignment       =   4
      End
      Begin Threed.SSPanel pnl_AplTNC 
         Height          =   315
         Left            =   4400
         TabIndex        =   55
         Top             =   2610
         Width           =   1350
         _Version        =   65536
         _ExtentX        =   2381
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "0.00 "
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
         RoundedCorners  =   0   'False
         Font3D          =   2
         Alignment       =   4
      End
      Begin Threed.SSPanel pnl_ApliTC 
         Height          =   315
         Left            =   4400
         TabIndex        =   56
         Top             =   2940
         Width           =   1350
         _Version        =   65536
         _ExtentX        =   2381
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "0.00 "
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
         RoundedCorners  =   0   'False
         Font3D          =   2
         Alignment       =   4
      End
      Begin Threed.SSPanel pnl_CapPbp 
         Height          =   315
         Left            =   4400
         TabIndex        =   61
         Top             =   2070
         Width           =   1350
         _Version        =   65536
         _ExtentX        =   2381
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "0.00 "
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
         RoundedCorners  =   0   'False
         Font3D          =   2
         Alignment       =   4
      End
      Begin Threed.SSPanel pnl_IntPbp 
         Height          =   315
         Left            =   10290
         TabIndex        =   62
         Top             =   1410
         Width           =   1230
         _Version        =   65536
         _ExtentX        =   2170
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "0.00 "
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
         RoundedCorners  =   0   'False
         Font3D          =   2
         Alignment       =   4
      End
      Begin Threed.SSPanel pnl_RedAno 
         Height          =   330
         Left            =   7440
         TabIndex        =   65
         Top             =   375
         Width           =   1230
         _Version        =   65536
         _ExtentX        =   2170
         _ExtentY        =   582
         _StockProps     =   15
         Caption         =   "0 "
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
      Begin Threed.SSPanel pnl_TipPre 
         Height          =   330
         Left            =   4400
         TabIndex        =   47
         Top             =   375
         Width           =   1350
         _Version        =   65536
         _ExtentX        =   2381
         _ExtentY        =   582
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
         RoundedCorners  =   0   'False
         Font3D          =   2
      End
      Begin VB.ComboBox cmb_TipPre 
         Height          =   315
         Left            =   4500
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "RED. MONTO"
         Top             =   390
         Visible         =   0   'False
         Width           =   1230
      End
      Begin Threed.SSPanel pnl_MtoApl_Final 
         Height          =   315
         Left            =   7440
         TabIndex        =   67
         Top             =   2070
         Width           =   1230
         _Version        =   65536
         _ExtentX        =   2170
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "0.00 "
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
         RoundedCorners  =   0   'False
         Font3D          =   2
         Alignment       =   4
      End
      Begin VB.Label Label12 
         Caption         =   "Monto Aplicar Final"
         Height          =   315
         Left            =   5910
         TabIndex        =   68
         Top             =   2130
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Reducc. Años"
         Height          =   315
         Left            =   5910
         TabIndex        =   66
         Top             =   450
         Width           =   1110
      End
      Begin VB.Label Label32 
         Caption         =   "Capital PBP"
         Height          =   315
         Left            =   3030
         TabIndex        =   64
         Top             =   2130
         Width           =   1410
      End
      Begin VB.Label Label33 
         Caption         =   "Interés PBP"
         Height          =   315
         Left            =   9180
         TabIndex        =   63
         Top             =   1470
         Width           =   1110
      End
      Begin VB.Label Label6 
         Caption         =   "Interés TNC "
         Height          =   315
         Left            =   5910
         TabIndex        =   60
         Top             =   1140
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Interés TC "
         Height          =   315
         Left            =   5910
         TabIndex        =   59
         Top             =   1470
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Monto Depositado"
         Height          =   315
         Left            =   8820
         TabIndex        =   58
         Top             =   450
         Width           =   1335
      End
      Begin VB.Label Label17 
         Caption         =   "Seguro Desgrav."
         Height          =   315
         Left            =   120
         TabIndex        =   39
         Top             =   1800
         Width           =   1410
      End
      Begin VB.Label Label3 
         Caption         =   "Monto a Aplicar"
         Height          =   315
         Left            =   120
         TabIndex        =   38
         Top             =   2130
         Width           =   1695
      End
      Begin VB.Label Label11 
         Caption         =   "Fecha Prepago"
         Height          =   315
         Left            =   120
         TabIndex        =   37
         Top             =   450
         Width           =   1515
      End
      Begin VB.Label Label13 
         Caption         =   "Aplicación TC"
         Height          =   315
         Left            =   3030
         TabIndex        =   36
         Top             =   3000
         Width           =   1410
      End
      Begin VB.Label Label9 
         Caption         =   "Cuotas Pendientes"
         Height          =   285
         Left            =   8820
         TabIndex        =   35
         Top             =   780
         Width           =   1695
      End
      Begin VB.Label Label10 
         Caption         =   "Nuevo TNC"
         Height          =   315
         Left            =   5910
         TabIndex        =   34
         Top             =   2670
         Width           =   1110
      End
      Begin VB.Label Label16 
         Caption         =   "Nuevo TC"
         Height          =   315
         Left            =   5910
         TabIndex        =   33
         Top             =   3000
         Width           =   1110
      End
      Begin VB.Label Label18 
         Caption         =   "Nueva Cuota"
         Height          =   315
         Left            =   8850
         TabIndex        =   32
         Top             =   2670
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Aplicación TNC"
         Height          =   315
         Left            =   3030
         TabIndex        =   31
         Top             =   2670
         Width           =   1410
      End
      Begin VB.Label Label5 
         Caption         =   "V. Asegurable Inm."
         Height          =   315
         Left            =   5910
         TabIndex        =   30
         Top             =   780
         Width           =   1515
      End
      Begin VB.Label Label19 
         Caption         =   "Seg. Inmueble"
         Height          =   315
         Left            =   3030
         TabIndex        =   29
         Top             =   1800
         Width           =   1110
      End
      Begin VB.Label Label20 
         Caption         =   "Saldo Actual TC"
         Height          =   315
         Left            =   3030
         TabIndex        =   28
         Top             =   780
         Width           =   1275
      End
      Begin VB.Label Label21 
         Caption         =   "Saldo Actual TNC"
         Height          =   315
         Left            =   120
         TabIndex        =   27
         Top             =   780
         Width           =   1515
      End
      Begin VB.Label Label22 
         Caption         =   "Ultimo Pago TC"
         Height          =   315
         Left            =   120
         TabIndex        =   26
         Top             =   1470
         Width           =   1410
      End
      Begin VB.Label Label23 
         Caption         =   "Ultimo Pago TNC"
         Height          =   315
         Left            =   120
         TabIndex        =   25
         Top             =   1140
         Width           =   1410
      End
      Begin VB.Label Label24 
         Caption         =   "Dias TC"
         Height          =   315
         Left            =   3030
         TabIndex        =   24
         Top             =   1470
         Width           =   1110
      End
      Begin VB.Label Label25 
         Caption         =   "Dias TNC"
         Height          =   315
         Left            =   3030
         TabIndex        =   23
         Top             =   1140
         Width           =   1110
      End
      Begin VB.Label Label26 
         Caption         =   "Monto del ITF"
         Height          =   315
         Left            =   5910
         TabIndex        =   22
         Top             =   1770
         Width           =   1035
      End
      Begin VB.Label Label27 
         Caption         =   "Saldo Actual TC"
         Height          =   315
         Left            =   120
         TabIndex        =   21
         Top             =   3000
         Width           =   1500
      End
      Begin VB.Label Label28 
         Caption         =   "Saldo Actual TNC"
         Height          =   315
         Left            =   120
         TabIndex        =   20
         Top             =   2670
         Width           =   1500
      End
      Begin VB.Label Label29 
         Caption         =   "Datos del Prepago"
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
         Left            =   60
         TabIndex        =   19
         Top             =   90
         UseMnemonic     =   0   'False
         Width           =   1875
      End
      Begin VB.Line Line1 
         X1              =   150
         X2              =   11500
         Y1              =   2490
         Y2              =   2490
      End
      Begin VB.Label Label31 
         Caption         =   "Tipo de Prepago"
         Height          =   315
         Left            =   3030
         TabIndex        =   18
         Top             =   450
         Width           =   1410
      End
      Begin VB.Label Label30 
         Caption         =   "Reducc. Años"
         Height          =   315
         Left            =   11640
         TabIndex        =   17
         Top             =   2085
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin Threed.SSPanel SSPanel4 
      Height          =   3720
      Left            =   30
      TabIndex        =   40
      Top             =   1410
      Width           =   11625
      _Version        =   65536
      _ExtentX        =   20505
      _ExtentY        =   6562
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
         Height          =   3285
         Left            =   60
         TabIndex        =   41
         Top             =   330
         Width           =   11565
         _ExtentX        =   20399
         _ExtentY        =   5794
         _Version        =   393216
         Rows            =   21
         FixedRows       =   0
         FixedCols       =   0
         BackColorSel    =   32768
         FocusRect       =   0
         ScrollBars      =   2
         SelectionMode   =   1
      End
      Begin VB.Label Label2 
         Caption         =   "Datos del Crédito"
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
         Left            =   60
         TabIndex        =   42
         Top             =   90
         Width           =   1875
      End
   End
   Begin Threed.SSPanel SSPanel3 
      Height          =   645
      Left            =   30
      TabIndex        =   43
      Top             =   720
      Width           =   11625
      _Version        =   65536
      _ExtentX        =   20505
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
      Begin VB.CommandButton cmd_ExpExc 
         Height          =   585
         Left            =   30
         Picture         =   "OpeTra_frm_330.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Exportar a Excel"
         Top             =   30
         Width           =   585
      End
      Begin VB.CommandButton cmd_Salida 
         Height          =   585
         Left            =   11010
         Picture         =   "OpeTra_frm_330.frx":0316
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Salir"
         Top             =   30
         Width           =   585
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   645
      Left            =   30
      TabIndex        =   45
      Top             =   30
      Width           =   11625
      _Version        =   65536
      _ExtentX        =   20505
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
      Begin Threed.SSPanel SSPanel7 
         Height          =   555
         Left            =   690
         TabIndex        =   46
         Top             =   60
         Width           =   4875
         _Version        =   65536
         _ExtentX        =   8599
         _ExtentY        =   979
         _StockProps     =   15
         Caption         =   "Prepago Parcial de Crédito Hipotecario - Consulta"
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
         Picture         =   "OpeTra_frm_330.frx":0758
         Top             =   90
         Width           =   480
      End
   End
   Begin Threed.SSPanel SSPanel6 
      Height          =   1450
      Left            =   15
      TabIndex        =   69
      Top             =   8520
      Width           =   11625
      _Version        =   65536
      _ExtentX        =   20505
      _ExtentY        =   2558
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.22
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      Begin VB.TextBox txt_ObsPpg 
         Enabled         =   0   'False
         Height          =   585
         Left            =   1470
         MaxLength       =   2000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   71
         Top             =   750
         Width           =   10065
      End
      Begin VB.ComboBox cmb_MotPpg 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1470
         TabIndex        =   70
         Text            =   "MOTIVO DEL PREPAGO"
         Top             =   420
         Width           =   10080
      End
      Begin VB.Label Label34 
         Caption         =   "Observaciones"
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
         TabIndex        =   74
         Top             =   90
         Width           =   2085
      End
      Begin VB.Label Label15 
         Caption         =   "Comentarios"
         Height          =   315
         Left            =   120
         TabIndex        =   73
         Top             =   780
         Width           =   1290
      End
      Begin VB.Label Label14 
         Caption         =   "Motivo Prepago"
         Height          =   315
         Left            =   120
         TabIndex        =   72
         Top             =   450
         Width           =   1290
      End
   End
End
Attribute VB_Name = "frm_Con_PrePgo_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_int_CodPrd           As Integer
Dim l_dbl_TasInt           As Double
Dim l_dbl_SegDes           As Double
Dim l_dbl_SegInm           As Double
 
Private Sub cmd_ExpExc_Click()
    If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    Me.Enabled = False
    
    'exporta liquidacion
    'If CInt(moddat_g_str_CodPrd) = 2 Or CInt(moddat_g_str_CodPrd) = 11 Or CInt(moddat_g_str_CodPrd) = 19 Or CInt(moddat_g_str_CodPrd) = 20 Then
    If InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd) > 0 Then
       Call fs_RptPar_Micasita
    Else
       Call fs_RptPar_Mivivienda
    End If
    
    Me.Enabled = True
    Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
    Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Limpia
   Call fs_Buscar
   Call fs_Buscar_Credto
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   'Inicializando Grid de Datos del Crédito
   grd_Listad.ColWidth(0) = 2900
   grd_Listad.ColWidth(1) = 8150
   grd_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
   
   cmb_TipPre.Clear
   cmb_TipPre.AddItem "RED. MONTO"
   cmb_TipPre.ItemData(cmb_TipPre.NewIndex) = 1
   cmb_TipPre.AddItem "RED. PLAZO"
   cmb_TipPre.ItemData(cmb_TipPre.NewIndex) = 2
   cmb_TipPre.ListIndex = -1
   
   cmb_RedPlz.Clear
   cmb_RedPlz.AddItem "1 AÑO"
   cmb_RedPlz.ItemData(cmb_RedPlz.NewIndex) = 1
   cmb_RedPlz.AddItem "2 AÑOS"
   cmb_RedPlz.ItemData(cmb_RedPlz.NewIndex) = 2
   cmb_RedPlz.AddItem "3 AÑOS"
   cmb_RedPlz.ItemData(cmb_RedPlz.NewIndex) = 3
   cmb_RedPlz.AddItem "4 AÑOS"
   cmb_RedPlz.ItemData(cmb_RedPlz.NewIndex) = 4
   cmb_RedPlz.AddItem "5 AÑOS"
   cmb_RedPlz.ItemData(cmb_RedPlz.NewIndex) = 5
   cmb_RedPlz.AddItem "6 AÑOS"
   cmb_RedPlz.ItemData(cmb_RedPlz.NewIndex) = 6
   cmb_RedPlz.AddItem "7 AÑOS"
   cmb_RedPlz.ItemData(cmb_RedPlz.NewIndex) = 7
   cmb_RedPlz.AddItem "8 AÑOS"
   cmb_RedPlz.ItemData(cmb_RedPlz.NewIndex) = 8
   cmb_RedPlz.AddItem "9 AÑOS"
   cmb_RedPlz.ItemData(cmb_RedPlz.NewIndex) = 9
   cmb_RedPlz.AddItem "10 AÑOS"
   cmb_RedPlz.ItemData(cmb_RedPlz.NewIndex) = 10
   cmb_RedPlz.AddItem "11 AÑOS"
   cmb_RedPlz.ItemData(cmb_RedPlz.NewIndex) = 11
   cmb_RedPlz.AddItem "12 AÑOS"
   cmb_RedPlz.ItemData(cmb_RedPlz.NewIndex) = 12
   cmb_RedPlz.AddItem "13 AÑOS"
   cmb_RedPlz.ItemData(cmb_RedPlz.NewIndex) = 13
   cmb_RedPlz.AddItem "14 AÑOS"
   cmb_RedPlz.ItemData(cmb_RedPlz.NewIndex) = 14
   cmb_RedPlz.ListIndex = -1
   cmb_RedPlz.Enabled = False
   
   pnl_Mto_Deposito.Caption = "0.00 "
   pnl_Val_AsgInm.Caption = "0.00 "
   pnl_SaldoTNC1.Caption = "0.00 "
   pnl_SaldoTC1.Caption = "0.00 "
   pnl_UltPagTNC.Caption = " "
   pnl_UltPagTC.Caption = " "
   pnl_DiasTNC.Caption = "0 "
   pnl_DiasTC.Caption = "0 "
   pnl_InteresTNC.Caption = 0
   pnl_InteresTC.Caption = 0
   pnl_SegDes.Caption = 0
   pnl_SegInm.Caption = 0
   pnl_MontoITF.Caption = 0
   pnl_CuoPend.Caption = 0
   pnl_MtoApl.Caption = "0.00 "
   pnl_MtoApl_Final.Caption = "0.00 "
   pnl_SaldoTNC2.Caption = "0.00 "
   pnl_SaldoTC2.Caption = "0.00 "
   pnl_AplTNC.Caption = 0
   pnl_ApliTC.Caption = 0
   pnl_NuevoSaldoTNC.Caption = "0.00 "
   pnl_NuevoSaldoTC.Caption = "0.00 "
   pnl_NuevaCuota.Caption = "0.00 "
   Call moddat_gs_Carga_LisIte_Combo(cmb_MotPpg, 1, "115")
End Sub

Private Sub fs_Limpia()
   Call gs_LimpiaGrid(grd_Listad)
End Sub

Private Sub fs_Buscar()
   'Información de los datos del Crédito
   Call modmip_gs_DatNumOpe(moddat_g_str_NumOpe, grd_Listad)
   l_int_CodPrd = moddat_g_str_CodPrd
   
   'Buscando Información del Prepago
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT PPGCAB_FECPPG, PPGCAB_TIPPPGPAR, PPGCAB_SLDACT_TNC, PPGCAB_SLDACT_TC, PPGCAB_SEGDES, PPGCAB_SEGINM, "
   g_str_Parame = g_str_Parame & "       PPGCAB_REDANO, PPGCAB_ULTPAG_TNC, PPGCAB_ULTPAG_TC, PPGCAB_MTOITF, PPGCAB_CUOPEN, PPGCAB_MTODEP, "
   g_str_Parame = g_str_Parame & "       PPGCAB_DIFDIA_TNC, PPGCAB_DIFDIA_TC, PPGCAB_PBPPER, PPGCAB_PBPINT, PPGCAB_MTOAPL, PPGCAB_INTCAL_TNC, "
   g_str_Parame = g_str_Parame & "       PPGCAB_INTCAL_TC, PPGCAB_APLTNC, PPGCAB_APLTC, PPGCAB_SLDACT_TNC, PPGCAB_SLDACT_TC, PPGCAB_MTOCUO, "
   g_str_Parame = g_str_Parame & "       PPGCAB_MOTPPG, PPGCAB_COMENT, HIPMAE_TASINT, HIPMAE_FOIPRE, HIPMAE_FOIVIV "
   g_str_Parame = g_str_Parame & "  FROM CRE_PPGCAB INNER JOIN CRE_HIPMAE ON HIPMAE_NUMOPE = PPGCAB_NUMOPE "
   g_str_Parame = g_str_Parame & " WHERE PPGCAB_NUMOPE = '" & moddat_g_str_NumOpe & "'  "
   g_str_Parame = g_str_Parame & "   AND PPGCAB_FECPPG =  " & moddat_g_str_FecIng & " "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If

   g_rst_Princi.MoveFirst
   ipp_FecPre.Text = gf_FormatoFecha(g_rst_Princi!PPGCAB_FECPPG)
   If g_rst_Princi!PPGCAB_TIPPPGPAR = 1 Then
      cmb_TipPre.ListIndex = 0
   Else
      cmb_TipPre.ListIndex = 1
   End If
   
    'TASAS DE INTERES
   l_dbl_TasInt = g_rst_Princi!HIPMAE_TASINT
   l_dbl_SegDes = g_rst_Princi!HIPMAE_FOIPRE
   l_dbl_SegInm = g_rst_Princi!HIPMAE_FOIVIV
   
   pnl_SaldoTNC1.Caption = Format(g_rst_Princi!PPGCAB_SLDACT_TNC, "###,##0.00") & " "
   pnl_SaldoTC1.Caption = Format(g_rst_Princi!PPGCAB_SLDACT_TC, "###,##0.00") & " "
   pnl_SegDes.Caption = Format(g_rst_Princi!PPGCAB_SEGDES, "###,##0.00") & " "
   pnl_SegInm = Format(g_rst_Princi!PPGCAB_SEGINM, "###,##0.00") & " "
   pnl_TipPre.Caption = cmb_TipPre.Text
   pnl_RedAno.Caption = Format(g_rst_Princi!PPGCAB_REDANO, "#0") & " "
   pnl_UltPagTNC.Caption = gf_FormatoFecha(g_rst_Princi!PPGCAB_ULTPAG_TNC)
   pnl_UltPagTC.Caption = gf_FormatoFecha(g_rst_Princi!PPGCAB_ULTPAG_TC)
   pnl_MontoITF.Caption = Format(g_rst_Princi!PPGCAB_MTOITF, "###,##0.00") & " "
   pnl_CuoPend.Caption = CInt(g_rst_Princi!PPGCAB_CUOPEN) & " "
   pnl_Mto_Deposito.Caption = Format(g_rst_Princi!PPGCAB_MTODEP, "###,##0.00") & " "
   pnl_DiasTNC.Caption = CInt(g_rst_Princi!PPGCAB_DIFDIA_TNC) & " "
   pnl_DiasTC.Caption = CInt(g_rst_Princi!PPGCAB_DIFDIA_TC) & " "
   If IsNull(g_rst_Princi!PPGCAB_PBPPER) Then
      pnl_CapPbp.Caption = Format(0, "###,##0.00") & " "
   Else
      pnl_CapPbp.Caption = Format(g_rst_Princi!PPGCAB_PBPPER, "###,##0.00") & " "
   End If
   If IsNull(g_rst_Princi!PPGCAB_PBPINT) Then
      pnl_IntPbp.Caption = Format(0, "###,##0.00") & " "
   Else
      pnl_IntPbp.Caption = Format(g_rst_Princi!PPGCAB_PBPINT, "###,##0.00") & " "
   End If
   pnl_MtoApl_Final.Caption = Format(IIf(IsNull(g_rst_Princi!PPGCAB_MTOAPL), 0, g_rst_Princi!PPGCAB_MTOAPL), "###,##0.00") & " "
   pnl_MtoApl.Caption = Format(CDbl(IIf(IsNull(g_rst_Princi!PPGCAB_MTOAPL), 0, g_rst_Princi!PPGCAB_MTOAPL)) + CDbl(IIf(IsNull(g_rst_Princi!PPGCAB_PBPPER), 0, g_rst_Princi!PPGCAB_PBPPER)), "###,##0.00") & " "
   pnl_InteresTNC.Caption = Format(IIf(IsNull(g_rst_Princi!PPGCAB_INTCAL_TNC), 0, g_rst_Princi!PPGCAB_INTCAL_TNC), "###,##0.00") & " "
   pnl_InteresTC.Caption = Format(g_rst_Princi!PPGCAB_INTCAL_TC, "###,##0.00") & " "
   pnl_SaldoTNC2.Caption = Format(g_rst_Princi!PPGCAB_SLDACT_TNC, "###,##0.00") & " "
   pnl_SaldoTC2.Caption = Format(g_rst_Princi!PPGCAB_SLDACT_TC, "###,##0.00") & " "
   pnl_AplTNC.Caption = Format(g_rst_Princi!PPGCAB_APLTNC, "###,##0.00") & " "
   pnl_ApliTC.Caption = Format(g_rst_Princi!PPGCAB_APLTC, "###,##0.00") & " "
   pnl_NuevoSaldoTNC.Caption = gf_FormatoNumero(g_rst_Princi!PPGCAB_SLDACT_TNC - CDbl(pnl_AplTNC.Caption), 12, 2) & " "
   pnl_NuevoSaldoTC.Caption = gf_FormatoNumero(g_rst_Princi!PPGCAB_SLDACT_TC - CDbl(pnl_ApliTC.Caption), 12, 2) & " "
   pnl_NuevaCuota.Caption = Format(g_rst_Princi!PPGCAB_MTOCUO, "###,##0.00") & " "
   If g_rst_Princi!PPGCAB_MOTPPG > 0 Then
      Call gs_BuscarCombo_Item(cmb_MotPpg, g_rst_Princi!PPGCAB_MOTPPG)
   End If
   txt_ObsPpg.Text = Trim(IIf(IsNull(g_rst_Princi!PPGCAB_COMENT), " ", g_rst_Princi!PPGCAB_COMENT))
   
   'DETERMINA SI OPERACION ES MICASITA
   If l_int_CodPrd = 2 Or l_int_CodPrd = 11 Or l_int_CodPrd = 19 Or l_int_CodPrd = 21 Or l_int_CodPrd = 22 Or l_int_CodPrd = 23 Or l_int_CodPrd = 24 Then
      Label21.Caption = "Saldo Actual"
      Label23.Caption = "Ultimo Pago"
      Label25.Caption = "Dias"
      Label6.Caption = "Interés a la fecha"
      Label20.Visible = False
      pnl_SaldoTC1.Visible = False
      Label22.Visible = False
      pnl_UltPagTC.Visible = False
      Label24.Visible = False
      pnl_DiasTC.Visible = False
      Label7.Visible = False
      pnl_InteresTC.Visible = False
      Label32.Visible = False
      pnl_CapPbp.Visible = False
      Label33.Visible = False
      pnl_IntPbp.Visible = False
      Label28.Caption = "Saldo Actual"
      Label8.Caption = "Aplicación PP"
      Label10.Caption = "Nuevo Saldo"
      Label27.Visible = False
      pnl_SaldoTC2.Visible = False
      Label13.Visible = False
      pnl_ApliTC.Visible = False
      Label16.Visible = False
      pnl_NuevoSaldoTC.Visible = False
   Else
      Label21.Caption = "Saldo Actual TNC"
      Label23.Caption = "Ultimo Pago TNC"
      Label25.Caption = "Dias TNC"
      Label6.Caption = "Interés TNC "
      Label20.Visible = True
      pnl_SaldoTC1.Visible = True
      Label22.Visible = True
      pnl_UltPagTC.Visible = True
      Label24.Visible = True
      pnl_DiasTC.Visible = True
      Label7.Visible = True
      pnl_InteresTC.Visible = True
      Label32.Visible = True
      pnl_CapPbp.Visible = True
      Label33.Visible = True
      pnl_IntPbp.Visible = True
      Label28.Caption = "Saldo Actual TNC"
      Label8.Caption = "Aplicación TNC"
      Label10.Caption = "Nuevo TNC"
      Label27.Visible = True
      pnl_SaldoTC2.Visible = True
      Label13.Visible = True
      pnl_ApliTC.Visible = True
      Label16.Visible = True
      pnl_NuevoSaldoTC.Visible = True
   End If
  
   Call gs_UbiIniGrid(grd_Listad)
End Sub

Private Sub fs_Buscar_ant()
Dim r_str_CodPry     As String
Dim r_str_NomPry     As String
Dim r_str_CodBco     As String
Dim l_int_CodPrd     As Integer
   
   'Buscando Información del Crédito
   g_str_Parame = "SELECT * FROM CRE_HIPMAE WHERE "
   g_str_Parame = g_str_Parame & "HIPMAE_NUMOPE = '" & moddat_g_str_NumOpe & "' AND "
   g_str_Parame = g_str_Parame & "(HIPMAE_SITUAC = 2 OR HIPMAE_SITUAC = 6 OR HIPMAE_SITUAC = 9)"
   
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
   moddat_g_str_NumSol = Trim(g_rst_Princi!hipmae_numsol)
   moddat_g_str_NumOpe = Trim(g_rst_Princi!HIPMAE_NUMOPE)
   
    'TASAS DE INTERES
   l_dbl_TasInt = g_rst_Princi!HIPMAE_TASINT
   l_dbl_SegDes = g_rst_Princi!HIPMAE_FOIPRE
   l_dbl_SegInm = g_rst_Princi!HIPMAE_FOIVIV
   
   'Obteniendo Nombre de Cliente
   moddat_g_str_NomCli = moddat_gf_Buscar_NomCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc)
   
   'Obteniendo Nombre y DOI de Cónyuge
   moddat_g_int_CygTDo = g_rst_Princi!HIPMAE_TDOCYG
   moddat_g_str_CygNDo = ""
   moddat_g_str_CygNom = ""
   If moddat_g_int_CygTDo > 0 Then
      moddat_g_str_CygNDo = Trim(g_rst_Princi!HIPMAE_NDOCYG & "")
      moddat_g_str_CygNom = moddat_gf_Buscar_NomCli(moddat_g_int_CygTDo, moddat_g_str_CygNDo)
   End If
   
   'Obteniendo Descripción de Producto
   moddat_g_str_CodPrd = Trim(g_rst_Princi!HIPMAE_CODPRD)
   moddat_g_str_NomPrd = moddat_gf_Consulta_Produc(Trim(g_rst_Princi!HIPMAE_CODPRD))
   moddat_g_str_CodSub = Trim(g_rst_Princi!HIPMAE_CODSUB)

   'Obeniendo Modalidad de Producto
   moddat_g_str_CodMod = Trim(g_rst_Princi!HIPMAE_CODMOD)
   moddat_g_str_DesMod = moddat_gf_Buscar_NomMod(Trim(g_rst_Princi!HIPMAE_CODPRD), moddat_g_str_CodMod)
   
   'Ejecutivo de Seguimiento
   moddat_g_str_CodEjeSeg = Trim(g_rst_Princi!HIPMAE_EJESEG & "")
   moddat_g_str_NomEjeSeg = moddat_gf_Buscar_NomEje(moddat_g_str_CodEjeSeg)

   'Consejero Hipotecario
   moddat_g_str_CodConHip = Trim(g_rst_Princi!HIPMAE_CONHIP & "")
   moddat_g_str_NomConHip = moddat_gf_Buscar_NomEje(moddat_g_str_CodConHip)

   'Moneda
   moddat_g_str_Moneda = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!HIPMAE_MONEDA))
   moddat_g_int_TipMon = g_rst_Princi!HIPMAE_MONEDA                           'Moneda Préstamo
   
   'Situación de Crédito
   moddat_g_int_Situac = g_rst_Princi!HIPMAE_SITUAC
   moddat_g_str_Situac = moddat_gf_Consulta_ParDes("027", CStr(g_rst_Princi!HIPMAE_SITUAC))
   
   'Obteniendo Información del Inmueble
   Call moddat_gs_Consulta_DatInm(moddat_g_str_NumSol, moddat_g_str_Direcc, moddat_g_str_Distri, r_str_CodPry, r_str_NomPry, r_str_CodBco)
   
   'Cargando en Grid
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.CellFontBold = True
   grd_Listad.Text = "Número de Operación"
   
   grd_Listad.Col = 1
   grd_Listad.CellFontBold = True
   grd_Listad.Text = gf_Formato_NumOpe(g_rst_Princi!HIPMAE_NUMOPE)
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.CellFontBold = True
   grd_Listad.Text = "Situación"
   
   grd_Listad.Col = 1
   grd_Listad.CellFontBold = True
   grd_Listad.Text = moddat_g_str_Situac
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.CellFontBold = True
   grd_Listad.Text = "Cliente"
   
   grd_Listad.Col = 1
   grd_Listad.CellFontBold = True
   grd_Listad.Text = CStr(g_rst_Princi!HIPMAE_TDOCLI) & " - " & Trim(g_rst_Princi!HIPMAE_NDOCLI) & " / " & moddat_g_str_NomCli
   
   If g_rst_Princi!HIPMAE_TDOCYG > 0 Then
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      grd_Listad.Col = 0
      grd_Listad.Text = "Cónyuge"
      
      grd_Listad.Col = 1
      grd_Listad.Text = CStr(g_rst_Princi!HIPMAE_TDOCYG) & " - " & Trim(g_rst_Princi!HIPMAE_NDOCYG) & " / " & moddat_g_str_CygNom
   End If
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Producto"
   
   grd_Listad.Col = 1
   grd_Listad.Text = moddat_g_str_NomPrd & " / " & moddat_gf_Consulta_SubPrd(g_rst_Princi!HIPMAE_CODPRD, g_rst_Princi!HIPMAE_CODSUB)
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Moneda Préstamo"
   
   grd_Listad.Col = 1
   grd_Listad.Text = moddat_g_str_Moneda
   
   grd_Listad.Rows = grd_Listad.Rows + 2
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Modalidad"
   
   grd_Listad.Col = 1
   grd_Listad.Text = moddat_g_str_DesMod
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Dirección Inmueble"
   
   grd_Listad.Col = 1
   grd_Listad.Text = moddat_g_str_Direcc
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Distrito"
   
   grd_Listad.Col = 1
   grd_Listad.Text = moddat_g_str_Distri
   
   If g_rst_Princi!HIPMAE_PRYMCS = 1 Or (g_rst_Princi!HIPMAE_PRYMCS = 2 And CInt(g_rst_Princi!HIPMAE_CODMOD) = 2 Or CInt(g_rst_Princi!HIPMAE_CODMOD) = 3) Then
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      grd_Listad.Col = 0
      grd_Listad.Text = "Proyecto Inmobiliario"
      
      grd_Listad.Col = 1
      grd_Listad.Text = moddat_gf_Consulta_NomPry(g_rst_Princi!HIPMAE_PRYINM & "")
      
      If g_rst_Princi!HIPMAE_PRYMCS = 2 Then
         grd_Listad.Text = grd_Listad.Text & " (" & moddat_gf_Consulta_ParDes("513", r_str_CodBco) & ")"
      End If
   End If
   
   If moddat_g_int_TipMon = 1 Then
      grd_Listad.Rows = grd_Listad.Rows + 2
      grd_Listad.Row = grd_Listad.Rows - 1
      grd_Listad.Col = 0
      grd_Listad.Text = "Valor Compra Venta"
      
      grd_Listad.Col = 1
      grd_Listad.CellFontName = "Lucida Console"
      grd_Listad.CellFontSize = 8
      grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_CVTSOL, 12, 2)
      
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      grd_Listad.Col = 0
      grd_Listad.Text = "Aporte Propio"
      
      grd_Listad.Col = 1
      grd_Listad.CellFontName = "Lucida Console"
      grd_Listad.CellFontSize = 8
      grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_APOSOL, 12, 2)
   Else
      grd_Listad.Rows = grd_Listad.Rows + 2
      grd_Listad.Row = grd_Listad.Rows - 1
      grd_Listad.Col = 0
      grd_Listad.Text = "Valor Compra Venta"
      
      grd_Listad.Col = 1
      grd_Listad.CellFontName = "Lucida Console"
      grd_Listad.CellFontSize = 8
      grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_CVTDOL, 12, 2)
   
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      grd_Listad.Col = 0
      grd_Listad.Text = "Aporte Propio"
      
      grd_Listad.Col = 1
      grd_Listad.CellFontName = "Lucida Console"
      grd_Listad.CellFontSize = 8
      grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_APODOL, 12, 2)
   End If
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Monto Desembolsado"
   
   grd_Listad.Col = 1
   grd_Listad.CellFontName = "Lucida Console"
   grd_Listad.CellFontSize = 8
   grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_IMPDES, 12, 2)
   
   grd_Listad.Rows = grd_Listad.Rows + 2
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Monto Préstamo"
   
   grd_Listad.Col = 1
   grd_Listad.CellFontName = "Lucida Console"
   grd_Listad.CellFontSize = 8
   grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_MTOPRE, 12, 2)
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Interés Capitalizado"
   
   grd_Listad.Col = 1
   grd_Listad.CellFontName = "Lucida Console"
   grd_Listad.CellFontSize = 8
   grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_INTCAP, 12, 2)
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Total Préstamo"
   
   grd_Listad.Col = 1
   grd_Listad.CellFontName = "Lucida Console"
   grd_Listad.CellFontSize = 8
   grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_TOTPRE, 12, 2)
   
   grd_Listad.Rows = grd_Listad.Rows + 2
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Fecha Activación"
   
   grd_Listad.Col = 1
   grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECACT))
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Fecha Desembolso"
   
   grd_Listad.Col = 1
   grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECDES))
   
   If g_rst_Princi!HIPMAE_FECESC > 0 Then
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      grd_Listad.Col = 0
      grd_Listad.Text = "Fecha Firma EE.PP"
      
      grd_Listad.Col = 1
      grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECESC))
   End If
   
   If moddat_g_str_CodPrd <> "002" Then
      grd_Listad.Rows = grd_Listad.Rows + 2
      grd_Listad.Row = grd_Listad.Rows - 1
      grd_Listad.Col = 0
      
      Select Case moddat_g_str_CodPrd > 0
         Case InStr(moddat_g_str_AgrCRC, moddat_g_str_CodPrd): grd_Listad.Text = "Nro. Operación Mivivienda"   '"001"
         Case InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd): grd_Listad.Text = "Nro. Operación COFIDE"       '"003"
         Case InStr(moddat_g_str_AgrTFMV, moddat_g_str_CodPrd): grd_Listad.Text = "Nro. Operación COFIDE"       '"004", "007", "009", "010", "013", "014", "015", "016", "017", "018", "019", "020", "021", "022", "023"
      End Select
      
      grd_Listad.Col = 1
      grd_Listad.Text = Trim(g_rst_Princi!HIPMAE_OPEMVI & "")
      
      If InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Then 'moddat_g_str_CodPrd = "003" Then
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         grd_Listad.Col = 0
         grd_Listad.Text = "Nro. Operación Mivivienda"
         
         grd_Listad.Col = 1
         grd_Listad.Text = Trim(g_rst_Princi!HIPMAE_OPEMV1 & "")
      End If
      
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      grd_Listad.Col = 0
      grd_Listad.Text = "Monto Préstamo (Tramo No Conces.)"
      
      grd_Listad.Col = 1
      grd_Listad.CellFontName = "Lucida Console"
      grd_Listad.CellFontSize = 8
      grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_IMPNCO, 12, 2)
   
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      grd_Listad.Col = 0
      grd_Listad.Text = "Monto Préstamo (Tramo Conces.)"
      
      grd_Listad.Col = 1
      grd_Listad.CellFontName = "Lucida Console"
      grd_Listad.CellFontSize = 8
      grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(moddat_g_int_TipMon)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_IMPCON, 12, 2)
      
      If InStr(moddat_g_str_AgrCRC, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Then 'moddat_g_str_CodPrd = "001" Or moddat_g_str_CodPrd = "003" Then
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         grd_Listad.Col = 0
         grd_Listad.Text = "Tasa de Interés Mivivienda"
      
         grd_Listad.Col = 1
         grd_Listad.Text = Format(g_rst_Princi!HIPMAE_TASMVI, "##0.00") & " %"
      End If
      
      If InStr(moddat_g_str_AgrCME, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_AgrTFMV, moddat_g_str_CodPrd) > 0 Then 'moddat_g_str_CodPrd = "004" Or moddat_g_str_CodPrd = "003" Or moddat_g_str_CodPrd = "007" Or moddat_g_str_CodPrd = "009" Or moddat_g_str_CodPrd = "010" Or moddat_g_str_CodPrd = "013" Or moddat_g_str_CodPrd = "014" Or moddat_g_str_CodPrd = "015" Or moddat_g_str_CodPrd = "016" Or moddat_g_str_CodPrd = "017" Or moddat_g_str_CodPrd = "018" Or moddat_g_str_CodPrd = "019" Or moddat_g_str_CodPrd = "020" Or moddat_g_str_CodPrd = "021" Or moddat_g_str_CodPrd = "022" Or moddat_g_str_CodPrd = "023" Then
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         grd_Listad.Col = 0
         grd_Listad.Text = "Tasa de Interés COFIDE"
      
         grd_Listad.Col = 1
         grd_Listad.Text = Format(g_rst_Princi!HIPMAE_TASCOF, "##0.00") & " %"
         
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         grd_Listad.Col = 0
         grd_Listad.Text = "Tasa de Comisión COFIDE"
         
         grd_Listad.Col = 1
         grd_Listad.Text = Format(g_rst_Princi!HIPMAE_COMCOF, "##0.00") & " %"
      End If
   End If
   
   grd_Listad.Rows = grd_Listad.Rows + 2
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Plazo"
   
   grd_Listad.Col = 1
   grd_Listad.Text = CStr(g_rst_Princi!HIPMAE_PLAANO) & " Años"
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Tasa de Interés"
   
   grd_Listad.Col = 1
   grd_Listad.Text = Format(g_rst_Princi!HIPMAE_TASINT, "##0.00") & " %"
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Nro. de Cuotas"
   
   grd_Listad.Col = 1
   grd_Listad.Text = CStr(g_rst_Princi!HIPMAE_NUMCUO)
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Período de Gracia"
   
   grd_Listad.Col = 1
   grd_Listad.Text = CStr(g_rst_Princi!HIPMAE_PERGRA) & " Meses"
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Compañía de Seguros"
   
   grd_Listad.Col = 1
   grd_Listad.Text = moddat_gf_Consulta_ComSeg(g_rst_Princi!HIPMAE_SEGPRE & "")
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Tipo de Seguro Desg."
   
   grd_Listad.Col = 1
   grd_Listad.Text = moddat_gf_Consulta_TipSeg(g_rst_Princi!HIPMAE_SEGPRE, g_rst_Princi!HIPMAE_TIPSEG)
      
   grd_Listad.Rows = grd_Listad.Rows + 2
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Tipo Garantía"
   
   grd_Listad.Col = 1
   grd_Listad.Text = moddat_gf_Consulta_ParDes("241", CStr(g_rst_Princi!HIPMAE_TIPGAR))
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Monto Garantía"
   
   grd_Listad.Col = 1
   grd_Listad.CellFontName = "Lucida Console"
   grd_Listad.CellFontSize = 8
   grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!HIPMAE_MONGAR)) & " " & gf_FormatoNumero(g_rst_Princi!HIPMAE_MTOGAR, 12, 2)
   
   grd_Listad.Rows = grd_Listad.Rows + 2
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Consejero Hipotecario"
   
   grd_Listad.Col = 1
   grd_Listad.Text = moddat_g_str_NomConHip
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Ejecutivo de Seguimiento"
   
   grd_Listad.Col = 1
   grd_Listad.Text = moddat_g_str_NomEjeSeg
   
   l_int_CodPrd = g_rst_Princi!HIPMAE_CODPRD
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing

   'Buscando Información del Prepago
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * "
   g_str_Parame = g_str_Parame & "  FROM CRE_PPGCAB "
   g_str_Parame = g_str_Parame & " WHERE PPGCAB_NUMOPE = '" & moddat_g_str_NumOpe & "'  "
   g_str_Parame = g_str_Parame & "   AND PPGCAB_FECPPG =  " & moddat_g_str_FecIng & " "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If

   g_rst_Princi.MoveFirst
   ipp_FecPre.Text = gf_FormatoFecha(g_rst_Princi!PPGCAB_FECPPG)
   If g_rst_Princi!PPGCAB_TIPPPGPAR = 1 Then
      cmb_TipPre.ListIndex = 0
   Else
      cmb_TipPre.ListIndex = 1
   End If
   
   pnl_SaldoTNC1.Caption = Format(g_rst_Princi!PPGCAB_SLDACT_TNC, "###,##0.00") & " "
   pnl_SaldoTC1.Caption = Format(g_rst_Princi!PPGCAB_SLDACT_TC, "###,##0.00") & " "
   pnl_SegDes.Caption = Format(g_rst_Princi!PPGCAB_SEGDES, "###,##0.00") & " "
   pnl_SegInm = Format(g_rst_Princi!PPGCAB_SEGINM, "###,##0.00") & " "
   pnl_TipPre.Caption = cmb_TipPre.Text
   pnl_RedAno.Caption = Format(g_rst_Princi!PPGCAB_REDANO, "#0") & " "
   pnl_UltPagTNC.Caption = gf_FormatoFecha(g_rst_Princi!PPGCAB_ULTPAG_TNC)
   pnl_UltPagTC.Caption = gf_FormatoFecha(g_rst_Princi!PPGCAB_ULTPAG_TC)
   pnl_MontoITF.Caption = Format(g_rst_Princi!PPGCAB_MTOITF, "###,##0.00") & " "
   pnl_CuoPend.Caption = CInt(g_rst_Princi!PPGCAB_CUOPEN) & " "
   pnl_Mto_Deposito.Caption = Format(g_rst_Princi!PPGCAB_MTODEP, "###,##0.00") & " "
   pnl_DiasTNC.Caption = CInt(g_rst_Princi!PPGCAB_DIFDIA_TNC) & " "
   pnl_DiasTC.Caption = CInt(g_rst_Princi!PPGCAB_DIFDIA_TC) & " "
   If IsNull(g_rst_Princi!PPGCAB_PBPPER) Then
      pnl_CapPbp.Caption = Format(0, "###,##0.00") & " "
   Else
      pnl_CapPbp.Caption = Format(g_rst_Princi!PPGCAB_PBPPER, "###,##0.00") & " "
   End If
   If IsNull(g_rst_Princi!PPGCAB_PBPINT) Then
      pnl_IntPbp.Caption = Format(0, "###,##0.00") & " "
   Else
      pnl_IntPbp.Caption = Format(g_rst_Princi!PPGCAB_PBPINT, "###,##0.00") & " "
   End If
   pnl_MtoApl.Caption = Format(g_rst_Princi!PPGCAB_MTOAPL, "###,##0.00") & " "
   '
   pnl_InteresTNC.Caption = Format(g_rst_Princi!PPGCAB_INTCAL_TNC, "###,##0.00") & " "
   pnl_InteresTC.Caption = Format(g_rst_Princi!PPGCAB_INTCAL_TC, "###,##0.00") & " "
   pnl_SaldoTNC2.Caption = Format(g_rst_Princi!PPGCAB_SLDACT_TNC, "###,##0.00") & " "
   pnl_SaldoTC2.Caption = Format(g_rst_Princi!PPGCAB_SLDACT_TC, "###,##0.00") & " "
   pnl_AplTNC.Caption = Format(g_rst_Princi!PPGCAB_APLTNC, "###,##0.00") & " "
   pnl_ApliTC.Caption = Format(g_rst_Princi!PPGCAB_APLTC, "###,##0.00") & " "
   pnl_NuevoSaldoTNC.Caption = gf_FormatoNumero(g_rst_Princi!PPGCAB_SLDACT_TNC - CDbl(pnl_AplTNC.Caption), 12, 2) & " "
   pnl_NuevoSaldoTC.Caption = gf_FormatoNumero(g_rst_Princi!PPGCAB_SLDACT_TC - CDbl(pnl_ApliTC.Caption), 12, 2) & " "
   pnl_NuevaCuota.Caption = Format(g_rst_Princi!PPGCAB_MTOCUO, "###,##0.00") & " "
   
   'DETERMINA SI OPERACION ES MICASITA
   If l_int_CodPrd = 2 Or l_int_CodPrd = 11 Or l_int_CodPrd = 19 Or l_int_CodPrd = 20 Then
      Label21.Caption = "Saldo Actual"
      Label23.Caption = "Ultimo Pago"
      Label25.Caption = "Dias"
      Label6.Caption = "Interés a la fecha"
      Label20.Visible = False
      pnl_SaldoTC1.Visible = False
      Label22.Visible = False
      pnl_UltPagTC.Visible = False
      Label24.Visible = False
      pnl_DiasTC.Visible = False
      Label7.Visible = False
      pnl_InteresTC.Visible = False
      Label32.Visible = False
      pnl_CapPbp.Visible = False
      Label33.Visible = False
      pnl_IntPbp.Visible = False
      Label28.Caption = "Saldo Actual"
      Label8.Caption = "Aplicacion PP"
      Label10.Caption = "Nuevo Saldo"
      Label27.Visible = False
      pnl_SaldoTC2.Visible = False
      Label13.Visible = False
      pnl_ApliTC.Visible = False
      Label16.Visible = False
      pnl_NuevoSaldoTC.Visible = False
   Else
      Label21.Caption = "Saldo Actual TNC"
      Label23.Caption = "Ultimo Pago TNC"
      Label25.Caption = "Dias TNC"
      Label6.Caption = "Interés TNC a la fecha"
      Label20.Visible = True
      pnl_SaldoTC1.Visible = True
      Label22.Visible = True
      pnl_UltPagTC.Visible = True
      Label24.Visible = True
      pnl_DiasTC.Visible = True
      Label7.Visible = True
      pnl_InteresTC.Visible = True
      Label32.Visible = True
      pnl_CapPbp.Visible = True
      Label33.Visible = True
      pnl_IntPbp.Visible = True
      Label28.Caption = "Saldo Actual TNC"
      Label8.Caption = "Aplicacion PP TNC"
      Label10.Caption = "Nuevo TNC"
      Label27.Visible = True
      pnl_SaldoTC2.Visible = True
      Label13.Visible = True
      pnl_ApliTC.Visible = True
      Label16.Visible = True
      pnl_NuevoSaldoTC.Visible = True
   End If
  
   Call gs_UbiIniGrid(grd_Listad)
End Sub

Private Sub fs_Buscar_Credto()
   Dim r_rst_Temp    As Recordset
   
   g_str_Parame = "SELECT EVATAS_TIPMON,EVATAS_SUMASE_INM,EVATAS_SUMASE_ES1,EVATAS_SUMASE_ES2,EVATAS_SUMASE_DEP "
   g_str_Parame = g_str_Parame & "FROM TRA_EVATAS WHERE "
   g_str_Parame = g_str_Parame & "EVATAS_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Temp, 3) Then
       Exit Sub
   End If
   
   If Not (r_rst_Temp.BOF And r_rst_Temp.EOF) Then
      r_rst_Temp.MoveFirst
      pnl_Val_AsgInm.Caption = gf_FormatoNumero(r_rst_Temp!EVATAS_SUMASE_INM + r_rst_Temp!EVATAS_SUMASE_ES1 + r_rst_Temp!EVATAS_SUMASE_ES2 + r_rst_Temp!EVATAS_SUMASE_DEP, 12, 2) & " "
   End If
   
   r_rst_Temp.Close
   Set r_rst_Temp = Nothing
End Sub

Private Function fs_Obtiene_FechaPago(ByVal p_NumOpe As String, ByVal p_TipCro As Integer, ByVal p_FecDes As String) As String
   Dim r_rst_Temp    As Recordset
   fs_Obtiene_FechaPago = p_FecDes
   
   g_str_Parame = "SELECT HIPCUO_FECVCT FROM CRE_HIPCUO WHERE "
   g_str_Parame = g_str_Parame & "HIPCUO_NUMOPE = '" & p_NumOpe & "' AND "
   g_str_Parame = g_str_Parame & "HIPCUO_SITUAC = 1 AND "
   g_str_Parame = g_str_Parame & "HIPCUO_TIPCRO = " & p_TipCro & " "
   g_str_Parame = g_str_Parame & "ORDER BY HIPCUO_FECVCT DESC"
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Temp, 3) Then
       Exit Function
   End If
   
   If Not (r_rst_Temp.BOF And r_rst_Temp.EOF) Then
      r_rst_Temp.MoveFirst
      fs_Obtiene_FechaPago = r_rst_Temp!HIPCUO_FECVCT
   End If
   
   r_rst_Temp.Close
   Set r_rst_Temp = Nothing
End Function

'******************************
' LIQUIDACION CREDITO MICASITA
'******************************
Private Sub fs_RptPar_Micasita()
   Dim r_obj_Excel      As Excel.Application
   Dim r_int_NroFil     As Integer
   Dim r_int_ColumC     As Integer
   Dim r_int_ColumD     As Integer
   Dim r_int_ColumE     As Integer
   Dim r_int_ColumF     As Integer
   Dim r_int_ColumK     As Integer
   Dim r_int_ColumL     As Integer
   Dim r_int_ColumM     As Integer
   Dim r_int_ColumN     As Integer
   Dim r_int_ColumO     As Integer
   Dim r_int_ColumP     As Integer
   Dim r_int_ColumQ     As Integer
   Dim r_int_ColumR     As Integer

   r_int_NroFil = 3
   r_int_ColumC = 3
   r_int_ColumD = 4
   r_int_ColumE = 5
   r_int_ColumF = 6
   r_int_ColumK = 11
   r_int_ColumL = 12
   r_int_ColumM = 13
   r_int_ColumN = 14
   r_int_ColumO = 15
   r_int_ColumP = 16
   r_int_ColumQ = 17
   r_int_ColumR = 18
   
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add

   With r_obj_Excel.ActiveSheet
      'CENTRADO DE LA PAGINA
      '.PageSetup.CenterHorizontally = True
      .PageSetup.Orientation = xlLandscape
      
      'MARGENES
      .PageSetup.LeftMargin = Application.CentimetersToPoints(1)
      .PageSetup.RightMargin = Application.CentimetersToPoints(1)
      .PageSetup.TopMargin = Application.CentimetersToPoints(1)
      .PageSetup.BottomMargin = Application.CentimetersToPoints(1)
      
      .Range("A1:W1").ColumnWidth = 4.57
      .Range("A8:A36").RowHeight = 13.5
      
      'BORDERS
      .Range("C2:V3").Borders(xlEdgeTop).Weight = xlMedium
      .Range("C4:V4").Borders(xlEdgeTop).Weight = xlMedium
      .Range("C6:V6").Borders(xlEdgeTop).Weight = xlMedium
      If cmb_TipPre.ItemData(cmb_TipPre.ListIndex) = 2 Then
        .Range("C32:V32").Borders(xlEdgeBottom).Weight = xlMedium
        .Range("C2:V32").Borders(xlEdgeLeft).Weight = xlMedium
        .Range("V2:V32").Borders(xlEdgeRight).Weight = xlMedium
      Else
        .Range("C31:V31").Borders(xlEdgeBottom).Weight = xlMedium
        .Range("C2:V31").Borders(xlEdgeLeft).Weight = xlMedium
        .Range("V2:V31").Borders(xlEdgeRight).Weight = xlMedium
      End If

      'Font
      .Range("A1:V44").Font.Name = "Arial"
      .Range("A1:V35").Font.Size = 9
      
      'Fecha de realizado el prepago
      .Range("O1") = "FECHA DE EMISIÓN: "
      .Range("O1").Font.Bold = True
      .Range("O1:R1").Merge
      .Range("S1:V1").Merge
      .Range("S1") = "'" & moddat_g_str_FecSis
      .Range("S1:V1").HorizontalAlignment = xlHAlignCenter
      .Range("O1:R1").HorizontalAlignment = xlHAlignCenter
      
      'Linea 3 - Titulo
      .Range("C2:V3").Merge
      .Range("C2") = "LIQUIDACION PREPAGO PARCIAL - " & moddat_g_str_NomPrd
      .Range("C2:V3").HorizontalAlignment = xlHAlignCenter
      .Range("C2:V3").VerticalAlignment = xlCenter
      .Range("C2:V3").Font.Size = 12
      .Range("C2:V3").Font.Bold = True
      r_int_NroFil = r_int_NroFil + 1
      
      'Linea 5 - Datos del cliente
      .Range("C4:V5").Merge
      .Range("C4") = "OPERACIÓN: " & moddat_g_str_NumOpe & " - CLIENTE: (DNI-" & moddat_g_str_NumDoc & ") " & moddat_g_str_NomCli
      .Range("C4:V5").HorizontalAlignment = xlHAlignCenter
      .Range("C4:V5").VerticalAlignment = xlCenter
      .Range("C4:V5").Font.Size = 11
      .Range("C4:V5").Font.Bold = True
      r_int_NroFil = r_int_NroFil + 3
      
      'Linea 9 - Saldo
      r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, r_int_ColumD) = "Saldo al"
      .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumK).HorizontalAlignment = xlHAlignRight
      .Cells(r_int_NroFil, r_int_ColumK) = "'" & Format(CDate(pnl_UltPagTNC.Caption), "dd-mm-yy")
      .Range("N" & r_int_NroFil & ":Q" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumN).Font.Bold = True
      .Cells(r_int_NroFil, r_int_ColumN).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, r_int_ColumN) = Format(CDbl(pnl_SaldoTNC2) + CDbl(pnl_SaldoTC2), "###,###.00") 'Format(moddat_g_dbl_SalCap + l_dbl_SalCon, "###,###.00")
      
      'MARCO
      .Cells(r_int_NroFil, r_int_ColumN).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumQ).Borders(xlEdgeRight).Weight = xlThin
      .Range("N" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("N" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      r_int_NroFil = r_int_NroFil + 1
      
      'Linea 10 - Fecha prepago
      .Cells(r_int_NroFil, r_int_ColumD) = "Fecha de corte (fecha del prepago)"
      .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumK) = "'" & Format(CDate(ipp_FecPre), "dd-mm-yy")
      .Cells(r_int_NroFil, r_int_ColumK).HorizontalAlignment = xlHAlignRight
      If cmb_TipPre.ItemData(cmb_TipPre.ListIndex) <> 2 Then
        .Range("D" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      End If
      
      'Si es reducción de plazo
      If cmb_TipPre.ItemData(cmb_TipPre.ListIndex) = 2 Then
        r_int_NroFil = r_int_NroFil + 1
        .Cells(r_int_NroFil, r_int_ColumD) = "Reducción de Plazo"
        .Cells(r_int_NroFil, r_int_ColumK) = cmb_RedPlz.Text
        .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Merge
        .Cells(r_int_NroFil, r_int_ColumK).HorizontalAlignment = xlHAlignRight
        .Range("D" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      End If
      r_int_NroFil = r_int_NroFil + 2
      
      'Linea 12 - Dias interes
      .Cells(r_int_NroFil, r_int_ColumD) = "Días de interés"
      .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumK) = CInt(pnl_DiasTNC.Caption)
      r_int_NroFil = r_int_NroFil + 1
      
'      'Linea  14 - Valor inmueble
'      .Cells(r_int_nrofil, r_int_ColumD) = "Valor asegurable inmueble"
'      .Range("K" & r_int_nrofil & ":L" & r_int_nrofil & "").Merge
'      .Cells(r_int_nrofil, r_int_ColumK).Font.Bold = True
'      .Cells(r_int_nrofil, r_int_ColumK) = Format(CDbl(pnl_Val_AsgInm.Caption), "###,###.00")
'      r_int_nrofil = r_int_nrofil + 2
      
      'Linea 16 - Tasa Interes anual
      .Cells(r_int_NroFil, r_int_ColumD) = "Tasa de interés anual (%)"
      .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumK).HorizontalAlignment = xlHAlignRight
      .Cells(r_int_NroFil, r_int_ColumK) = l_dbl_TasInt
      .Cells(r_int_NroFil, r_int_ColumK).Select
      r_obj_Excel.Selection.NumberFormat = "###0.0000"
      r_int_NroFil = r_int_NroFil + 1
      
      'Linea 17 - Tasa seguro inmueble
      .Cells(r_int_NroFil, r_int_ColumD) = "Tasa de Seguro Inmueble (%)"
      .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumK).HorizontalAlignment = xlHAlignRight
      .Cells(r_int_NroFil, r_int_ColumK) = l_dbl_SegInm
      .Cells(r_int_NroFil, r_int_ColumK).Select
      r_obj_Excel.Selection.NumberFormat = "###0.0000"
      r_int_NroFil = r_int_NroFil + 1
      
      'Linea 18 - Tasa seguro desgravamen
      .Cells(r_int_NroFil, r_int_ColumD) = "Tasa de Seguro Desgravamen (%)"
      .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumK).HorizontalAlignment = xlHAlignRight
      .Cells(r_int_NroFil, r_int_ColumK) = l_dbl_SegDes
      .Cells(r_int_NroFil, r_int_ColumK).Select
      r_obj_Excel.Selection.NumberFormat = "###0.0000"
      r_int_NroFil = r_int_NroFil + 2
      
      'Linea 20 - monto depositado
      .Range("D" & r_int_NroFil & ":M" & (r_int_NroFil) & "").Merge
      .Range("N" & r_int_NroFil & ":P" & (r_int_NroFil) & "").Merge
      .Cells(r_int_NroFil, r_int_ColumD) = "Monto Depositado"
      .Cells(r_int_NroFil, r_int_ColumN) = Format(CDbl(pnl_Mto_Deposito.Caption), "###,###.00")
      .Cells(r_int_NroFil, r_int_ColumD).VerticalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, r_int_ColumN).VerticalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, r_int_ColumD).Font.Bold = True
      .Cells(r_int_NroFil, r_int_ColumN).Font.Bold = True
      
       'MARCO
      .Cells(r_int_NroFil, r_int_ColumE).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, r_int_ColumO).HorizontalAlignment = xlHAlignRight
      .Cells(r_int_NroFil, r_int_ColumD).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumM).Borders(xlEdgeRight).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumQ).Borders(xlEdgeRight).Weight = xlThin
      .Range("D" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Cells(r_int_NroFil, r_int_ColumD).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumM).Borders(xlEdgeRight).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumQ).Borders(xlEdgeRight).Weight = xlThin
      .Range("D" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      r_int_NroFil = r_int_NroFil + 2
      
      'Linea 22 - interes
      .Cells(r_int_NroFil, r_int_ColumD) = "Intereses a la fecha"
      .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumK) = CDbl(pnl_InteresTNC.Caption)
      .Cells(r_int_NroFil, r_int_ColumK).Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.00"
      r_int_NroFil = r_int_NroFil + 1
      
      'Linea 23 - seguro desgravamen
      .Cells(r_int_NroFil, r_int_ColumD) = "Seguro Desgravamen"
      .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumK) = CDbl(pnl_SegDes.Caption)
      .Cells(r_int_NroFil, r_int_ColumK).Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.00"
      r_int_NroFil = r_int_NroFil + 1
      
      'Linea 24 - seguro inmueble
      .Cells(r_int_NroFil, r_int_ColumD) = "Seguro Inmueble"
      .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumK) = CDbl(pnl_SegInm.Caption)
      .Cells(r_int_NroFil, r_int_ColumK).Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.00"
      r_int_NroFil = r_int_NroFil + 1
      
'      'Linea 25 - Deuda Pendiente
'      .Cells(r_int_NroFil, r_int_ColumD) = "Deuda Pendiente"
'      .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Merge
'      .Cells(r_int_NroFil, r_int_ColumK) = CDbl(pnl_DeuPen.Caption)
'      .Cells(r_int_NroFil, r_int_ColumK).Select
'      r_obj_Excel.Selection.NumberFormat = "###,##0.00"
'      r_int_NroFil = r_int_NroFil + 1
      
      'Linea 26 - total gastos
      .Cells(r_int_NroFil, r_int_ColumD) = "Total de Interés y Seguros"
      .Range("D" & r_int_NroFil & ":M" & r_int_NroFil & "").Merge
      .Range("N" & r_int_NroFil & ":P" & r_int_NroFil & "").Merge
'      .Cells(r_int_NroFil, r_int_ColumN) = Format(CDbl(CDbl(pnl_InteresTNC.Caption) + CDbl(pnl_SegDes.Caption) + CDbl(pnl_SegInm.Caption) + CDbl(pnl_DeuPen.Caption)), "###,##0.00")
      .Cells(r_int_NroFil, r_int_ColumN) = Format(CDbl(CDbl(pnl_InteresTNC.Caption) + CDbl(pnl_SegDes.Caption) + CDbl(pnl_SegInm.Caption)), "###,##0.00")
      .Cells(r_int_NroFil, r_int_ColumD).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, r_int_ColumN).HorizontalAlignment = xlHAlignRight
      
      'MARCO
      .Cells(r_int_NroFil, r_int_ColumD).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumM).Borders(xlEdgeRight).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumQ).Borders(xlEdgeRight).Weight = xlThin
      .Range("D" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("D" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      r_int_NroFil = r_int_NroFil + 2
      
      'Linea 27 - ITF
      .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Merge
      .Range("N" & r_int_NroFil & ":P" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumD) = "ITF (%)"
      .Cells(r_int_NroFil, r_int_ColumK) = "'0.005"
      .Cells(r_int_NroFil, r_int_ColumN) = CDbl(pnl_MontoITF.Caption)
      .Cells(r_int_NroFil, r_int_ColumN).Select
      r_obj_Excel.Selection.NumberFormat = "###0.00"
      .Cells(r_int_NroFil, r_int_ColumK).HorizontalAlignment = xlHAlignRight
      .Cells(r_int_NroFil, r_int_ColumN).HorizontalAlignment = xlHAlignRight
      r_int_NroFil = r_int_NroFil + 2
      
      'Linea 29 - monto prepago
      .Cells(r_int_NroFil, r_int_ColumD) = "Monto del Prepago a Aplicar"
      .Range("D" & r_int_NroFil & ":M" & r_int_NroFil & "").Merge
      .Range("N" & r_int_NroFil & ":P" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumN) = Format(CDbl(pnl_MtoApl.Caption), "###,###.00")
      .Cells(r_int_NroFil, r_int_ColumD).Font.Bold = True
      .Cells(r_int_NroFil, r_int_ColumN).Font.Bold = True
      .Cells(r_int_NroFil, r_int_ColumD).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, r_int_ColumN).HorizontalAlignment = xlHAlignRight
      
      'MARCO
      .Cells(r_int_NroFil, r_int_ColumD).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumM).Borders(xlEdgeRight).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumQ).Borders(xlEdgeRight).Weight = xlThin
      .Range("D" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("D" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      r_int_NroFil = r_int_NroFil + 2
      
      'Linea 31 - saldo reprogramar
      .Cells(r_int_NroFil, r_int_ColumD) = "Saldo después del prepago"
      .Range("D" & r_int_NroFil & ":M" & r_int_NroFil & "").Merge
      .Range("N" & r_int_NroFil & ":P" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumN) = Format(CDbl(pnl_NuevoSaldoTNC.Caption), "###,###.00")
      .Cells(r_int_NroFil, r_int_ColumD).Font.Bold = True
      .Cells(r_int_NroFil, r_int_ColumN).Font.Bold = True
      .Cells(r_int_NroFil, r_int_ColumD).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, r_int_ColumN).HorizontalAlignment = xlHAlignRight
      
      'MARCO
      .Cells(r_int_NroFil, r_int_ColumD).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumM).Borders(xlEdgeRight).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumQ).Borders(xlEdgeRight).Weight = xlThin
      .Range("D" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("D" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      r_int_NroFil = r_int_NroFil + 2
      
      'Linea 33 - monto cuota
      .Cells(r_int_NroFil, r_int_ColumD) = "MONTO DE LA NUEVA CUOTA"
      .Range("D" & r_int_NroFil & ":M" & r_int_NroFil & "").Merge
      .Range("N" & r_int_NroFil & ":P" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumN) = CDbl(pnl_NuevaCuota.Caption)
      .Cells(r_int_NroFil, r_int_ColumD).Font.Bold = True
      .Cells(r_int_NroFil, r_int_ColumN).Font.Bold = True
      .Cells(r_int_NroFil, r_int_ColumD).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, r_int_ColumN).HorizontalAlignment = xlHAlignRight
      
      'MARCO
      .Cells(r_int_NroFil, r_int_ColumD).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumM).Borders(xlEdgeRight).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumQ).Borders(xlEdgeRight).Weight = xlThin
      .Range("D" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("D" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range("N" & r_int_NroFil & ":P" & r_int_NroFil & "").Select
      r_int_NroFil = r_int_NroFil + 2
      
      If CInt(moddat_g_int_TipMon) = 1 Then
         r_obj_Excel.Selection.NumberFormat = "$###,##0.00_);[Red]($###,##0.00)"
         If cmb_TipPre.ItemData(cmb_TipPre.ListIndex) = 2 Then
            .Cells(34, 3) = "'" & "- Esta liquidación es válida solo a la fecha de corte."
            .Cells(35, 3) = "'" & "- Cualquier gasto adicional que no contemple la presente liquidación, será informado al cliente y "
            .Cells(36, 3) = "'" & "  tendrá que ser cancelado antes de realizar el abono que consigna la presente liquidación."
            .Cells(37, 3) = "'" & "- Realizar el depósito en la cuenta Nº 0011-0369-02-00090532 del BBVA Banco Continental."
            .Cells(38, 3) = "'" & "- Realizar el depósito en la cuenta Nº 011-369-000200090532-69 del BBVA Banco Continental desde otro Banco."
            .Range(.Cells(34, 3), .Cells(38, 3)).Font.Size = 11
            .Range(.Cells(34, 3), .Cells(38, 3)).Font.Bold = True
         Else
            .Cells(33, 3) = "'" & "- Esta liquidación es válida solo a la fecha de corte."
            .Cells(34, 3) = "'" & "- Cualquier gasto adicional que no contemple la presente liquidación, será informado al cliente y "
            .Cells(35, 3) = "'" & "  tendrá que ser cancelado antes de realizar el abono que consigna la presente liquidación."
            .Cells(36, 3) = "'" & "- Realizar el depósito en la cuenta Nº 0011-0369-02-00090532 del BBVA Banco Continental."
            .Cells(37, 3) = "'" & "- Realizar el depósito en la cuenta Nº 011-369-000200090532-69 del BBVA Banco Continental desde otro Banco."
            .Range(.Cells(33, 3), .Cells(37, 3)).Font.Size = 11
            .Range(.Cells(33, 3), .Cells(37, 3)).Font.Bold = True
         End If
      Else
         r_obj_Excel.Selection.NumberFormat = "[$$]#,##0.00;[Red][$$]#,##0.00"
         If cmb_TipPre.ItemData(cmb_TipPre.ListIndex) = 2 Then
            
            .Cells(34, 3) = "'" & "- Esta liquidación es válida solo a la fecha de corte."
            .Cells(35, 3) = "'" & "- Cualquier gasto adicional que no contemple la presente liquidación, será informado al cliente y "
            .Cells(36, 3) = "'" & "  tendrá que ser cancelado antes de realizar el abono que consigna la presente liquidación."
            .Cells(37, 3) = "'" & "- Realizar el depósito en la cuenta Nº 0011-0369-02-00090540 del BBVA Banco Continental."
            .Range(.Cells(34, 3), .Cells(37, 3)).Font.Size = 11
            .Range(.Cells(34, 3), .Cells(37, 3)).Font.Bold = True
         Else
            .Cells(33, 3) = "'" & "- Esta liquidación es válida solo a la fecha de corte."
            .Cells(34, 3) = "'" & "- Cualquier gasto adicional que no contemple la presente liquidación, será informado al cliente y "
            .Cells(35, 3) = "'" & "  tendrá que ser cancelado antes de realizar el abono que consigna la presente liquidación."
            .Cells(36, 3) = "'" & "- Realizar el depósito en la cuenta Nº 0011-0369-02-00090540 del BBVA Banco Continental."
            .Range(.Cells(33, 3), .Cells(36, 3)).Font.Size = 11
            .Range(.Cells(33, 3), .Cells(36, 3)).Font.Bold = True
         End If
      End If
   End With
      
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub fs_RptPar_Micasita_ant()
    Dim r_obj_Excel      As Excel.Application
    Dim r_int_NroFil     As Integer
    Dim r_int_ColumC     As Integer
    Dim r_int_ColumD     As Integer
    Dim r_int_ColumE     As Integer
    Dim r_int_ColumF     As Integer
    Dim r_int_ColumK     As Integer
    Dim r_int_ColumL     As Integer
    Dim r_int_ColumM     As Integer
    Dim r_int_ColumN     As Integer
    Dim r_int_ColumO     As Integer
    Dim r_int_ColumP     As Integer
    Dim r_int_ColumQ     As Integer
    Dim r_int_ColumR     As Integer
     
    r_int_NroFil = 3
    r_int_ColumC = 3
    r_int_ColumD = 4
    r_int_ColumE = 5
    r_int_ColumF = 6
    r_int_ColumK = 11
    r_int_ColumL = 12
    r_int_ColumM = 13
    r_int_ColumN = 14
    r_int_ColumO = 15
    r_int_ColumP = 16
    r_int_ColumQ = 17
    r_int_ColumR = 18
    
    Set r_obj_Excel = New Excel.Application
    r_obj_Excel.SheetsInNewWorkbook = 1
    r_obj_Excel.Workbooks.Add
    
    With r_obj_Excel.ActiveSheet
       'CENTRADO DE LA PAGINA
       '.PageSetup.CenterHorizontally = True
       .PageSetup.Orientation = xlLandscape
       
       'MARGENES
       .PageSetup.LeftMargin = Application.CentimetersToPoints(1)
       .PageSetup.RightMargin = Application.CentimetersToPoints(1)
       .PageSetup.TopMargin = Application.CentimetersToPoints(1)
       .PageSetup.BottomMargin = Application.CentimetersToPoints(1)
       
       .Range("A1:W1").ColumnWidth = 4.57
       .Range("A8:A36").RowHeight = 13.5
       
       'BORDERS
       .Range("C3:T3").Borders(xlEdgeTop).Weight = xlMedium
       .Range("C4:T4").Borders(xlEdgeTop).Weight = xlMedium
       .Range("C5:T5").Borders(xlEdgeTop).Weight = xlMedium
       .Range("C6:T6").Borders(xlEdgeTop).Weight = xlMedium
       .Range("C34:T34").Borders(xlEdgeBottom).Weight = xlMedium
       .Range("C3:T34").Borders(xlEdgeLeft).Weight = xlMedium
       .Range("T3:T34").Borders(xlEdgeRight).Weight = xlMedium
       
       'Font
       .Range("A1:T44").Font.Name = "Arial"
       .Range("A1:T34").Font.Size = 9
       
       'Linea 3 - Titulo
       .Range("C3:T3").Merge
       .Range("C3") = "LIQUIDACION PREPAGO PARCIAL - " & moddat_g_str_NomPrd
       .Range("C3:T3").HorizontalAlignment = xlHAlignCenter
       .Range("C3:T3").Font.Size = 12
       .Range("C5:T5").Font.Size = 12
       .Range("C7:T7").Font.Size = 12
       .Range("C3:T3").Font.Bold = True
       .Range("C5:T5").Font.Bold = True
       .Range("C7:T7").Font.Bold = True
       r_int_NroFil = r_int_NroFil + 2
       
       'Linea 5 - Datos del cliente
       .Cells(r_int_NroFil, r_int_ColumD) = "Cliente:"
       .Cells(r_int_NroFil, r_int_ColumF) = moddat_g_str_NomCli
       .Cells(r_int_NroFil, r_int_ColumQ) = "DNI"
       .Cells(r_int_NroFil, r_int_ColumR) = "'" & moddat_g_str_NumDoc
       r_int_NroFil = r_int_NroFil + 2
       
       'Linea 7 - Operacion, moneda
       .Cells(r_int_NroFil, r_int_ColumD) = "N° de Operación: " & Mid(moddat_g_str_NumOpe, 1, 3) & "-" & Mid(moddat_g_str_NumOpe, 4, 2) & "-" & Mid(moddat_g_str_NumOpe, 6, 5)
       .Cells(r_int_NroFil, r_int_ColumD).HorizontalAlignment = xlHAlignLeft
       .Cells(r_int_NroFil, r_int_ColumD).Font.Bold = True
       .Cells(r_int_NroFil, r_int_ColumM) = "Moneda: " & moddat_g_str_Moneda
       .Cells(r_int_NroFil, r_int_ColumM).HorizontalAlignment = xlHAlignLeft
       .Cells(r_int_NroFil, r_int_ColumM).Font.Bold = True
       r_int_NroFil = r_int_NroFil + 2
       
       'Linea 9 - Saldo
       r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, r_int_ColumD) = "Saldo al"
       .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Merge
       .Cells(r_int_NroFil, r_int_ColumK).HorizontalAlignment = xlHAlignRight
       .Cells(r_int_NroFil, r_int_ColumK) = "'" & Format(CDate(pnl_UltPagTNC.Caption), "dd-mm-yy")
       .Range("N" & r_int_NroFil & ":Q" & r_int_NroFil & "").Merge
       .Cells(r_int_NroFil, r_int_ColumN).Font.Bold = True
       .Cells(r_int_NroFil, r_int_ColumN).HorizontalAlignment = xlHAlignCenter
       .Cells(r_int_NroFil, r_int_ColumN) = Format(pnl_SaldoTNC2.Caption, "###,###.00")
       
       'MARCO
       .Cells(r_int_NroFil, r_int_ColumN).Borders(xlEdgeLeft).Weight = xlThin
       .Cells(r_int_NroFil, r_int_ColumQ).Borders(xlEdgeRight).Weight = xlThin
       .Range("N" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
       .Range("N" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
       r_int_NroFil = r_int_NroFil + 1
       
       'Linea 10 - Fecha prepago
       .Cells(r_int_NroFil, r_int_ColumD) = "Fecha de corte (fecha del prepago)"
       .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Merge
       .Cells(r_int_NroFil, r_int_ColumK) = "'" & Format(CDate(ipp_FecPre), "dd-mm-yy")
       .Cells(r_int_NroFil, r_int_ColumK).HorizontalAlignment = xlHAlignRight
       .Range("D" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
       r_int_NroFil = r_int_NroFil + 2
       
       'Linea 12 - Dias interes
       .Cells(r_int_NroFil, r_int_ColumD) = "Días de interés"
       .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Merge
       .Cells(r_int_NroFil, r_int_ColumK) = CInt(pnl_DiasTNC.Caption)
       r_int_NroFil = r_int_NroFil + 2
       
       'Linea  14 - Valor inmueble
       .Cells(r_int_NroFil, r_int_ColumD) = "Valor asegurable inmueble"
       .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Merge
       .Cells(r_int_NroFil, r_int_ColumK).Font.Bold = True
       .Cells(r_int_NroFil, r_int_ColumK) = Format(CDbl(pnl_Val_AsgInm.Caption), "###,###.00")
       r_int_NroFil = r_int_NroFil + 2
       
       'Linea 16 - Tasa Interes anual
       .Cells(r_int_NroFil, r_int_ColumD) = "Tasa de interés anual (%)"
       .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Merge
       .Cells(r_int_NroFil, r_int_ColumK).HorizontalAlignment = xlHAlignRight
       .Cells(r_int_NroFil, r_int_ColumK) = l_dbl_TasInt
       .Cells(r_int_NroFil, r_int_ColumK).Select
       r_obj_Excel.Selection.NumberFormat = "###0.0000"
       r_int_NroFil = r_int_NroFil + 1
       
       'Linea 17 - Tasa seguro inmueble
       .Cells(r_int_NroFil, r_int_ColumD) = "Tasa de Seguro Inmueble (%)"
       .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Merge
       .Cells(r_int_NroFil, r_int_ColumK).HorizontalAlignment = xlHAlignRight
       .Cells(r_int_NroFil, r_int_ColumK) = l_dbl_SegInm
       .Cells(r_int_NroFil, r_int_ColumK).Select
       r_obj_Excel.Selection.NumberFormat = "###0.0000"
       r_int_NroFil = r_int_NroFil + 1
       
       'Linea 18 - Tasa seguro desgravamen
       .Cells(r_int_NroFil, r_int_ColumD) = "Tasa de Seguro Desgravamen (%)"
       .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Merge
       .Cells(r_int_NroFil, r_int_ColumK).HorizontalAlignment = xlHAlignRight
       .Cells(r_int_NroFil, r_int_ColumK) = l_dbl_SegDes
       .Cells(r_int_NroFil, r_int_ColumK).Select
       r_obj_Excel.Selection.NumberFormat = "###0.0000"
       r_int_NroFil = r_int_NroFil + 2
       
       'Linea 20 - monto depositado
       .Range("D" & r_int_NroFil & ":M" & (r_int_NroFil) & "").Merge
       .Range("N" & r_int_NroFil & ":P" & (r_int_NroFil) & "").Merge
       .Cells(r_int_NroFil, r_int_ColumD) = "Monto Depositado"
       .Cells(r_int_NroFil, r_int_ColumN) = Format(CDbl(pnl_Mto_Deposito.Caption), "###,###.00")
       .Cells(r_int_NroFil, r_int_ColumD).VerticalAlignment = xlHAlignCenter
       .Cells(r_int_NroFil, r_int_ColumN).VerticalAlignment = xlHAlignCenter
       .Cells(r_int_NroFil, r_int_ColumD).Font.Bold = True
       .Cells(r_int_NroFil, r_int_ColumN).Font.Bold = True
       
        'MARCO
       .Cells(r_int_NroFil, r_int_ColumE).HorizontalAlignment = xlHAlignCenter
       .Cells(r_int_NroFil, r_int_ColumO).HorizontalAlignment = xlHAlignRight
       .Cells(r_int_NroFil, r_int_ColumD).Borders(xlEdgeLeft).Weight = xlThin
       .Cells(r_int_NroFil, r_int_ColumM).Borders(xlEdgeRight).Weight = xlThin
       .Cells(r_int_NroFil, r_int_ColumQ).Borders(xlEdgeRight).Weight = xlThin
       .Range("D" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
       .Cells(r_int_NroFil, r_int_ColumD).Borders(xlEdgeLeft).Weight = xlThin
       .Cells(r_int_NroFil, r_int_ColumM).Borders(xlEdgeRight).Weight = xlThin
       .Cells(r_int_NroFil, r_int_ColumQ).Borders(xlEdgeRight).Weight = xlThin
       .Range("D" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
       r_int_NroFil = r_int_NroFil + 2
       
       'Linea 22 - interes
       .Cells(r_int_NroFil, r_int_ColumD) = "Intereses a la fecha"
       .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Merge
       .Cells(r_int_NroFil, r_int_ColumK) = CDbl(pnl_InteresTNC.Caption)
       .Cells(r_int_NroFil, r_int_ColumK).Select
       r_obj_Excel.Selection.NumberFormat = "###0.00"
       r_int_NroFil = r_int_NroFil + 1
       
       'Linea 23 - seguro desgravamen
       .Cells(r_int_NroFil, r_int_ColumD) = "Seguro Desgravamen"
       .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Merge
       .Cells(r_int_NroFil, r_int_ColumK) = CDbl(pnl_SegDes.Caption)
       .Cells(r_int_NroFil, r_int_ColumK).Select
       r_obj_Excel.Selection.NumberFormat = "###0.00"
       r_int_NroFil = r_int_NroFil + 1
       
       'Linea 24 - seguro inmueble
       .Cells(r_int_NroFil, r_int_ColumD) = "Seguro Inmueble"
       .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Merge
       .Cells(r_int_NroFil, r_int_ColumK) = CDbl(pnl_SegInm.Caption)
       .Cells(r_int_NroFil, r_int_ColumK).Select
       r_obj_Excel.Selection.NumberFormat = "###0.00"
       r_int_NroFil = r_int_NroFil + 1
       
       'Linea 25 - total gastos
       .Cells(r_int_NroFil, r_int_ColumD) = "Total de Interés y seguros"
       .Range("D" & r_int_NroFil & ":M" & r_int_NroFil & "").Merge
       .Range("N" & r_int_NroFil & ":P" & r_int_NroFil & "").Merge
       .Cells(r_int_NroFil, r_int_ColumN) = CDbl(CDbl(pnl_InteresTNC.Caption) + CDbl(pnl_SegDes.Caption) + CDbl(pnl_SegInm.Caption))
       .Cells(r_int_NroFil, r_int_ColumD).HorizontalAlignment = xlHAlignCenter
       .Cells(r_int_NroFil, r_int_ColumN).HorizontalAlignment = xlHAlignRight
       
       'MARCO
       .Cells(r_int_NroFil, r_int_ColumD).Borders(xlEdgeLeft).Weight = xlThin
       .Cells(r_int_NroFil, r_int_ColumM).Borders(xlEdgeRight).Weight = xlThin
       .Cells(r_int_NroFil, r_int_ColumQ).Borders(xlEdgeRight).Weight = xlThin
       .Range("D" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
       .Range("D" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
       r_int_NroFil = r_int_NroFil + 2
       
       'Linea 27 - ITF
       .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Merge
       .Range("N" & r_int_NroFil & ":P" & r_int_NroFil & "").Merge
       .Cells(r_int_NroFil, r_int_ColumD) = "ITF (%)"
       .Cells(r_int_NroFil, r_int_ColumK) = "'0.005"
       .Cells(r_int_NroFil, r_int_ColumN) = CDbl(pnl_MontoITF.Caption)
       .Cells(r_int_NroFil, r_int_ColumN).Select
       r_obj_Excel.Selection.NumberFormat = "###0.00"
       .Cells(r_int_NroFil, r_int_ColumK).HorizontalAlignment = xlHAlignRight
       .Cells(r_int_NroFil, r_int_ColumN).HorizontalAlignment = xlHAlignRight
       r_int_NroFil = r_int_NroFil + 2
       
       'Linea 29 - monto prepago
       .Cells(r_int_NroFil, r_int_ColumD) = "Monto del Prepago a Aplicar"
       .Range("D" & r_int_NroFil & ":M" & r_int_NroFil & "").Merge
       .Range("N" & r_int_NroFil & ":P" & r_int_NroFil & "").Merge
       .Cells(r_int_NroFil, r_int_ColumN) = Format(CDbl(pnl_MtoApl_Final.Caption), "###,###.00")
       .Cells(r_int_NroFil, r_int_ColumD).Font.Bold = True
       .Cells(r_int_NroFil, r_int_ColumN).Font.Bold = True
       .Cells(r_int_NroFil, r_int_ColumD).HorizontalAlignment = xlHAlignCenter
       .Cells(r_int_NroFil, r_int_ColumN).HorizontalAlignment = xlHAlignRight
       
       'MARCO
       .Cells(r_int_NroFil, r_int_ColumD).Borders(xlEdgeLeft).Weight = xlThin
       .Cells(r_int_NroFil, r_int_ColumM).Borders(xlEdgeRight).Weight = xlThin
       .Cells(r_int_NroFil, r_int_ColumQ).Borders(xlEdgeRight).Weight = xlThin
       .Range("D" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
       .Range("D" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
       r_int_NroFil = r_int_NroFil + 2
       
       'Linea 31 - saldo reprogramar
       .Cells(r_int_NroFil, r_int_ColumD) = "Saldo después del prepago"
       .Range("D" & r_int_NroFil & ":M" & r_int_NroFil & "").Merge
       .Range("N" & r_int_NroFil & ":P" & r_int_NroFil & "").Merge
       .Cells(r_int_NroFil, r_int_ColumN) = Format(CDbl(pnl_NuevoSaldoTNC.Caption), "###,###.00")
       .Cells(r_int_NroFil, r_int_ColumD).Font.Bold = True
       .Cells(r_int_NroFil, r_int_ColumN).Font.Bold = True
       .Cells(r_int_NroFil, r_int_ColumD).HorizontalAlignment = xlHAlignCenter
       .Cells(r_int_NroFil, r_int_ColumN).HorizontalAlignment = xlHAlignRight
       
       'MARCO
       .Cells(r_int_NroFil, r_int_ColumD).Borders(xlEdgeLeft).Weight = xlThin
       .Cells(r_int_NroFil, r_int_ColumM).Borders(xlEdgeRight).Weight = xlThin
       .Cells(r_int_NroFil, r_int_ColumQ).Borders(xlEdgeRight).Weight = xlThin
       .Range("D" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
       .Range("D" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
       r_int_NroFil = r_int_NroFil + 2
       
       'Linea 33 - monto cuota
       .Cells(r_int_NroFil, r_int_ColumD) = "MONTO DE LA NUEVA CUOTA"
       .Range("D" & r_int_NroFil & ":M" & r_int_NroFil & "").Merge
       .Range("N" & r_int_NroFil & ":P" & r_int_NroFil & "").Merge
       .Cells(r_int_NroFil, r_int_ColumN) = CDbl(pnl_NuevaCuota.Caption)
       .Cells(r_int_NroFil, r_int_ColumD).Font.Bold = True
       .Cells(r_int_NroFil, r_int_ColumN).Font.Bold = True
       .Cells(r_int_NroFil, r_int_ColumD).HorizontalAlignment = xlHAlignCenter
       .Cells(r_int_NroFil, r_int_ColumN).HorizontalAlignment = xlHAlignRight
       
       'MARCO
       .Cells(r_int_NroFil, r_int_ColumD).Borders(xlEdgeLeft).Weight = xlThin
       .Cells(r_int_NroFil, r_int_ColumM).Borders(xlEdgeRight).Weight = xlThin
       .Cells(r_int_NroFil, r_int_ColumQ).Borders(xlEdgeRight).Weight = xlThin
       .Range("D" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
       .Range("D" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
       .Range("N" & r_int_NroFil & ":P" & r_int_NroFil & "").Select
       r_int_NroFil = r_int_NroFil + 2
       
       If CInt(moddat_g_int_TipMon) = 1 Then
          r_obj_Excel.Selection.NumberFormat = "$###,##0.00_);[Red]($###,##0.00)"
          .Cells(36, 3) = "Realizar el depósito en la cuenta Nº 0011-0369-02-00090532 del BBVA Banco Continental."
          .Cells(36, 3).Font.Size = 12
          .Cells(36, 3).Font.Bold = True
       Else
          r_obj_Excel.Selection.NumberFormat = "[$$]#,##0.00;[Red][$$]#,##0.00"
          .Cells(36, 3) = "Realizar el depósito en la cuenta Nº 0011-0369-02-00090540 del BBVA Banco Continental."
          .Cells(36, 3).Font.Size = 12
          .Cells(36, 3).Font.Bold = True
       End If
    End With
      
    r_obj_Excel.Visible = True
    Set r_obj_Excel = Nothing
End Sub

'********************************
' LIQUIDACION CREDITO MIVIVIENDA
'********************************
Private Sub fs_RptPar_Mivivienda()
Dim r_obj_Excel      As Excel.Application
Dim r_int_NroFil     As Integer
Dim r_int_ColumC     As Integer
Dim r_int_ColumF     As Integer
Dim r_int_ColumK     As Integer
Dim r_int_ColumL     As Integer
Dim r_int_ColumM     As Integer
Dim r_int_ColumN     As Integer
Dim r_int_ColumO     As Integer
Dim r_int_ColumP     As Integer
Dim r_int_ColumQ     As Integer
Dim r_int_ColumR     As Integer
Dim r_int_ColumS     As Integer
Dim r_int_ColumT     As Integer
Dim r_int_ColumV     As Integer
Dim r_int_ColumW     As Integer
Dim r_int_ColumX     As Integer
Dim r_int_ColumY     As Integer
Dim r_int_ColumZ     As Integer

   r_int_NroFil = 2
   r_int_ColumC = 3
   r_int_ColumF = 6
   r_int_ColumK = 11
   r_int_ColumL = 12
   r_int_ColumM = 13
   r_int_ColumN = 14
   r_int_ColumO = 15
   r_int_ColumP = 16
   r_int_ColumQ = 17
   r_int_ColumR = 18
   r_int_ColumS = 19
   r_int_ColumT = 20
   r_int_ColumV = 22
   r_int_ColumW = 23
   r_int_ColumX = 24
   r_int_ColumY = 25
   r_int_ColumZ = 26
   
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add

   With r_obj_Excel.ActiveSheet
      'CENTRADO DE LA PAGINA
      '.PageSetup.CenterHorizontally = True
      .PageSetup.Orientation = xlLandscape
      
      'MARGENES
      .PageSetup.LeftMargin = Application.CentimetersToPoints(1)
      .PageSetup.RightMargin = Application.CentimetersToPoints(1)
      .PageSetup.TopMargin = Application.CentimetersToPoints(1)
      .PageSetup.BottomMargin = Application.CentimetersToPoints(1)
      
      .Range("A1:A1").ColumnWidth = 1.14
      .Range("B1:AB1").ColumnWidth = 3.57
      .Range("A2:A46").RowHeight = 12
      
      'BORDERS
      .Range("B2:AA2").Borders(xlEdgeTop).Weight = xlMedium
      .Range("B4:AA4").Borders(xlEdgeTop).Weight = xlMedium
      .Range("B6:AA6").Borders(xlEdgeTop).Weight = xlMedium
      
      If cmb_TipPre.ItemData(cmb_TipPre.ListIndex) = 2 Then
        .Range("B41:AA41").Borders(xlEdgeBottom).Weight = xlMedium
        .Range("B2:AA41").Borders(xlEdgeLeft).Weight = xlMedium
        .Range("AA2:AA41").Borders(xlEdgeRight).Weight = xlMedium
      Else
        .Range("B40:AA40").Borders(xlEdgeBottom).Weight = xlMedium
        .Range("B2:AA40").Borders(xlEdgeLeft).Weight = xlMedium
        .Range("AA2:AA40").Borders(xlEdgeRight).Weight = xlMedium
        .Range("B1:AA46").Font.Name = "Arial"
      End If
      
      'Font
      .Range("B1:AA46").Font.Name = "Arial"
      .Range("B1:AA40").Font.Size = 9
      
      'Fecha de realizado el prepago
      .Range("T1") = "FECHA DE EMISIÓN: "
      .Range("T1").Font.Bold = True
      .Range("T1:X1").Merge
      .Range("Y1:AA1").Merge
      .Range("Y1") = "'" & moddat_g_str_FecSis
      .Range("T1:X1").HorizontalAlignment = xlHAlignCenter
      .Range("Y1:AA1").HorizontalAlignment = xlHAlignCenter
      
      'Linea 1 - Titulo
      .Range("B2:AA3").Merge
      .Range("B2") = "LIQUIDACION PREPAGO PARCIAL - " & moddat_g_str_NomPrd & " - MONEDA " & moddat_g_str_Moneda
      .Range("B2:AA3").HorizontalAlignment = xlHAlignCenter
      .Range("B2:AA3").VerticalAlignment = xlCenter
      .Range("B2:AA3").Font.Size = 12
      .Range("B4:AA4").Font.Size = 12
      .Range("B6:AA6").Font.Size = 12
      .Range("B2:AA2").Font.Bold = True
      .Range("B4:AA4").Font.Bold = True
      .Range("B6:AA6").Font.Bold = True
'      .Range("B2:AA3").RowHeight = 12
      
      r_int_NroFil = r_int_NroFil + 2
      
      'Linea 2 - Datos del Cliente
      .Range("B4:AA5").Merge
      .Range("B4") = "OPERACIÓN: " & moddat_g_str_NumOpe & " - CLIENTE: (DNI-" & moddat_g_str_NumDoc & ") " & moddat_g_str_NomCli
      .Cells(r_int_NroFil, r_int_ColumC).HorizontalAlignment = xlCenter
      .Cells(r_int_NroFil, r_int_ColumC).VerticalAlignment = xlCenter
      .Range("B4:AA5").Font.Size = 11
'      .Range("B4:AA5").RowHeight = 12
       r_int_NroFil = r_int_NroFil + 3
        
      'Linea 3 - Fecha de Desembolso o última cuota TC (A)
      .Cells(r_int_NroFil, r_int_ColumC) = "Fecha de Desembolso o última cuota TC (A)"
      .Range("L" & r_int_NroFil & ":M" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumL) = "'" & Format(CDate(pnl_UltPagTC.Caption), "dd-mm-yy")
      .Cells(r_int_NroFil, r_int_ColumL).HorizontalAlignment = xlHAlignRight

      'Linea 3 - Días de interés TNC (C-B)
      .Cells(r_int_NroFil, r_int_ColumR) = "Días de interés TNC (C-B)"
      .Range("Y" & r_int_NroFil & ":Z" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumY) = CInt(pnl_DiasTNC.Caption)
      r_int_NroFil = r_int_NroFil + 1
      
      'Linea 4 - Saldo TNC al (B)
      r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, r_int_ColumC) = "Saldo TNC al (B)"
      .Range("L" & r_int_NroFil & ":M" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumL).HorizontalAlignment = xlHAlignRight
      .Cells(r_int_NroFil, r_int_ColumT).HorizontalAlignment = xlHAlignRight
      .Cells(r_int_NroFil, r_int_ColumL) = "'" & Format(CDate(pnl_UltPagTNC.Caption), "dd-mm-yy")
      
       'Linea 4 - Sumatoria pnl_SaldoTNC2 + pnl_SaldoTC2
      .Range("N" & r_int_NroFil & ":P" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumN).Font.Bold = True
      .Cells(r_int_NroFil, r_int_ColumN).HorizontalAlignment = xlHAlignRight
      .Cells(r_int_NroFil, r_int_ColumN) = Format(CDbl(pnl_SaldoTNC2) + CDbl(pnl_SaldoTC2), "###,###.00")

      'MARCO
      .Cells(r_int_NroFil, r_int_ColumN).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumP).Borders(xlEdgeRight).Weight = xlThin
      .Range("N" & r_int_NroFil & ":P" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("N" & r_int_NroFil & ":P" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous

      'Linea 4 - Días de interés TC (C-A)
      .Cells(r_int_NroFil, r_int_ColumR) = "Días de interés TC (C-A)"
      .Range("Y" & r_int_NroFil & ":Z" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumY) = CInt(pnl_DiasTC.Caption)
      r_int_NroFil = r_int_NroFil + 1
      
      'Linea 5 - Fecha de corte (fecha del prepago) (C)
      .Cells(r_int_NroFil, r_int_ColumC) = "Fecha de corte (fecha del prepago) (C)"
'      If cmb_TipPre.ItemData(cmb_TipPre.ListIndex) <> 2 Then
'        .Range("C" & r_int_nrofil & ":P" & r_int_nrofil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
'      End If
      .Range("L" & r_int_NroFil & ":M" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumL) = "'" & Format(CDate(ipp_FecPre), "dd-mm-yy")
      .Cells(r_int_NroFil, r_int_ColumL).HorizontalAlignment = xlHAlignRight
      
      'Linea 5 - Tasa de interés anual (%)
      .Cells(r_int_NroFil, r_int_ColumR) = "Tasa de interés anual (%)"
      .Range("Y" & r_int_NroFil & ":Z" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumY) = l_dbl_TasInt
      .Cells(r_int_NroFil, r_int_ColumY).Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.0000"
      r_int_NroFil = r_int_NroFil + 1
      
      'Si es reducción de plazo
      If cmb_TipPre.ItemData(cmb_TipPre.ListIndex) = 2 Then
        .Cells(r_int_NroFil, r_int_ColumC) = "Reducción de Plazo"
        .Range("L" & r_int_NroFil & ":M" & r_int_NroFil & "").Merge
        .Cells(r_int_NroFil, r_int_ColumL) = cmb_RedPlz.Text
        .Cells(r_int_NroFil, r_int_ColumL).HorizontalAlignment = xlHAlignRight
'        .Range("C" & r_int_nrofil & ":P" & r_int_nrofil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      End If
     
      'Linea 6 - Tasa de Seguro Inmueble (%)
      .Cells(r_int_NroFil, r_int_ColumR) = "Tasa de Seguro Inmueble (%)"
      .Range("Y" & r_int_NroFil & ":Z" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumY) = l_dbl_SegInm
      .Range("Y" & r_int_NroFil & ":Y" & r_int_NroFil & "").Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.0000"
      r_int_NroFil = r_int_NroFil + 1

      'Linea 7 - Tasa de Seguro Desgravamen (%)
      .Cells(r_int_NroFil, r_int_ColumR) = "Tasa de Seguro Desgravamen (%)"
      .Range("Y" & r_int_NroFil & ":Z" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumY) = l_dbl_SegDes
      .Range("Y" & r_int_NroFil & ":Y" & r_int_NroFil & "").Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.0000"
      r_int_NroFil = r_int_NroFil + 2
      
      'Linea 8 - Monto Depositado
      .Range("C" & r_int_NroFil & ":M" & (r_int_NroFil) & "").Merge
      .Range("N" & r_int_NroFil & ":P" & (r_int_NroFil) & "").Merge
      .Cells(r_int_NroFil, r_int_ColumC) = "Monto Depositado"
      .Cells(r_int_NroFil, r_int_ColumN) = Format(CDbl(pnl_Mto_Deposito.Caption), "###,###.00")
      .Cells(r_int_NroFil, r_int_ColumC).Font.Bold = True
      .Cells(r_int_NroFil, r_int_ColumN).Font.Bold = True

       'MARCO
      .Cells(r_int_NroFil, r_int_ColumC).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, r_int_ColumN).HorizontalAlignment = xlHAlignRight
      .Cells(r_int_NroFil, r_int_ColumC).VerticalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, r_int_ColumN).VerticalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, r_int_ColumC).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumP).Borders(xlEdgeRight).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumM).Borders(xlEdgeRight).Weight = xlThin
      .Range("C" & r_int_NroFil & ":P" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Cells(r_int_NroFil, r_int_ColumC).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumP).Borders(xlEdgeRight).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumM).Borders(xlEdgeRight).Weight = xlThin
      .Range("C" & r_int_NroFil & ":P" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      r_int_NroFil = r_int_NroFil + 2
      
      'Linea 9 - Intereses TNC a la fecha
      .Cells(r_int_NroFil, r_int_ColumC) = "Intereses TNC a la fecha"
      .Range("L" & r_int_NroFil & ":M" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumL) = CDbl(pnl_InteresTNC.Caption)
      .Cells(r_int_NroFil, r_int_ColumL).Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.00"
      
      r_int_NroFil = r_int_NroFil + 1
      
      'Linea 10 - Intereses TC a la fecha
      .Cells(r_int_NroFil, r_int_ColumC) = "Intereses TC a la fecha"
      .Range("L" & r_int_NroFil & ":M" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumL) = CDbl(pnl_InteresTC.Caption)
      .Cells(r_int_NroFil, r_int_ColumL).Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.00"
      
      'Linea 10 - Saldo antes del Prepago
      .Range("S" & r_int_NroFil & ":Z" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumS).Font.Bold = True
      .Cells(r_int_NroFil, r_int_ColumS) = "Saldo antes del prepago"
      
       'MARCO
      .Cells(r_int_NroFil, r_int_ColumS).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumZ).Borders(xlEdgeRight).Weight = xlThin
      .Range("S" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("S" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Cells(r_int_NroFil, r_int_ColumT).HorizontalAlignment = xlHAlignCenter

      r_int_NroFil = r_int_NroFil + 1
          
      'Linea 11 - Seguro Desgravamen
      .Cells(r_int_NroFil, r_int_ColumC) = "Seguro Desgravamen"
      .Range("L" & r_int_NroFil & ":M" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumL) = CDbl(pnl_SegDes.Caption)
      .Cells(r_int_NroFil, r_int_ColumL).Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.00"
      
      'Linea 11 - Saldo antes del Prepago - TNC
      .Range("S" & r_int_NroFil & ":V" & r_int_NroFil & "").Merge
      .Range("W" & r_int_NroFil & ":Y" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumS) = "TNC"
      .Cells(r_int_NroFil, r_int_ColumW) = Format(CDbl(pnl_SaldoTNC1.Caption), "###,###.00")
      .Cells(r_int_NroFil, r_int_ColumS).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, r_int_ColumW).HorizontalAlignment = xlHAlignRight
      
      'MARCO
      .Cells(r_int_NroFil, r_int_ColumS).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumV).Borders(xlEdgeRight).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumZ).Borders(xlEdgeRight).Weight = xlThin
      .Range("S" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("S" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      r_int_NroFil = r_int_NroFil + 1
      
      'Linea 12 - Seguro inmueble
      .Cells(r_int_NroFil, r_int_ColumC) = "Seguro Inmueble"
      .Range("L" & r_int_NroFil & ":M" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumL) = CDbl(pnl_SegInm.Caption)
      .Cells(r_int_NroFil, r_int_ColumL).Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.00"
     
      'Linea 12 - Saldo antes del Prepago - TC
      .Range("S" & r_int_NroFil & ":V" & r_int_NroFil & "").Merge
      .Range("W" & r_int_NroFil & ":Y" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumS) = "TC"
      .Cells(r_int_NroFil, r_int_ColumW) = Format(CDbl(pnl_SaldoTC1.Caption), "###,###.00")
      .Cells(r_int_NroFil, r_int_ColumS).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, r_int_ColumW).HorizontalAlignment = xlHAlignRight
      
      'MARCO
      .Cells(r_int_NroFil, r_int_ColumS).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumV).Borders(xlEdgeRight).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumZ).Borders(xlEdgeRight).Weight = xlThin
      .Range("S" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
'      If cmb_TipPre.ItemData(cmb_TipPre.ListIndex) <> 2 Then
'        .Range("C" & r_int_nrofil & ":Q" & r_int_nrofil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
'      End If
      r_int_NroFil = r_int_NroFil + 1
    
      'Linea 13 - Interes PBP
      .Cells(r_int_NroFil, r_int_ColumC) = "Interés PBP"
      .Range("L" & r_int_NroFil & ":M" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumL) = CDbl(pnl_IntPbp.Caption)
      .Cells(r_int_NroFil, r_int_ColumL).Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.00"
      
      'Linea 13 - Saldo antes del prepago - Capital PBP
      .Range("S" & r_int_NroFil & ":V" & r_int_NroFil & "").Merge
      .Range("W" & r_int_NroFil & ":Y" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumS) = "PBP"
      .Cells(r_int_NroFil, r_int_ColumW) = Format(CDbl(pnl_CapPbp.Caption), "###,###.00")
      .Cells(r_int_NroFil, r_int_ColumS).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, r_int_ColumW).HorizontalAlignment = xlHAlignRight
            
      'MARCO
      .Cells(r_int_NroFil, r_int_ColumS).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumV).Borders(xlEdgeRight).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumZ).Borders(xlEdgeRight).Weight = xlThin
      .Range("S" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("S" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      r_int_NroFil = r_int_NroFil + 1
      
'      'Linea 14 - Deuda Pendiente
'      .Cells(r_int_NroFil, r_int_ColumC) = "Deuda Pendiente"
'      .Range("L" & r_int_NroFil & ":M" & r_int_NroFil & "").Merge
'      .Cells(r_int_NroFil, r_int_ColumL) = CDbl(pnl_DeuPen.Caption)
'      .Cells(r_int_NroFil, r_int_ColumL).Select
'      r_obj_Excel.Selection.NumberFormat = "###,##0.00"

       'Linea 14 - Total Saldo antes del Prepago
      .Range("S" & r_int_NroFil & ":V" & r_int_NroFil & "").Merge
      .Range("W" & r_int_NroFil & ":Y" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumS).Font.Bold = True
      .Cells(r_int_NroFil, r_int_ColumS) = "Total"
      .Cells(r_int_NroFil, r_int_ColumW) = Format(CDbl(pnl_SaldoTNC2) + CDbl(pnl_SaldoTC2) + CDbl(pnl_CapPbp), "###,###.00")
      .Cells(r_int_NroFil, r_int_ColumS).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, r_int_ColumW).HorizontalAlignment = xlHAlignRight
      .Range("N" & r_int_NroFil & ":O" & r_int_NroFil & "").HorizontalAlignment = xlHAlignCenter
      
      'MARCO
      .Cells(r_int_NroFil, r_int_ColumS).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumV).Borders(xlEdgeRight).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumZ).Borders(xlEdgeRight).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumW).Font.Bold = True
      .Range("S" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      r_int_NroFil = r_int_NroFil + 1
      
'      If cmb_TipPre.ItemData(cmb_TipPre.ListIndex) = 2 Then
'         r_int_nrofil = r_int_nrofil + 1
'      End If
      
      'Linea 15 - Total de Interés, Seguros y Vencidos
      .Cells(r_int_NroFil, r_int_ColumC) = "Total de Interés, seguros y vencidos"
      .Range("C" & r_int_NroFil & ":M" & r_int_NroFil & "").Merge
      .Range("N" & r_int_NroFil & ":P" & r_int_NroFil & "").Merge
'      .Cells(r_int_NroFil, r_int_ColumN) = CDbl(CDbl(txt_InteresTNC.Text) + CDbl(txt_InteresTC.Text) + CDbl(txt_SegDes.Text) + CDbl(txt_SegInm.Text) + CDbl(pnl_IntPbp.Caption)) + CDbl(pnl_DeuPen.Caption)
      .Cells(r_int_NroFil, r_int_ColumN) = CDbl(CDbl(pnl_InteresTNC.Caption) + CDbl(pnl_InteresTC.Caption) + CDbl(pnl_SegDes.Caption) + CDbl(pnl_SegInm.Caption) + CDbl(pnl_IntPbp.Caption))
      .Cells(r_int_NroFil, r_int_ColumC).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, r_int_ColumN).HorizontalAlignment = xlHAlignRight

      'MARCO
      .Cells(r_int_NroFil, r_int_ColumC).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumN).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumP).Borders(xlEdgeRight).Weight = xlThin
      .Range("C" & r_int_NroFil & ":P" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C" & r_int_NroFil & ":P" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Cells(r_int_NroFil, r_int_ColumC).Font.Bold = True
      .Cells(r_int_NroFil, r_int_ColumN).Font.Bold = True
      r_int_NroFil = r_int_NroFil + 3
      
      'Linea 16 - Monto de Prepago a Aplicar
      .Range("C" & r_int_NroFil & ":M" & r_int_NroFil & "").Merge
      .Range("N" & r_int_NroFil & ":P" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumC).Font.Bold = True
      .Cells(r_int_NroFil, r_int_ColumN).Font.Bold = True
      .Cells(r_int_NroFil, r_int_ColumC) = "Monto de Prepago a Aplicar"
      .Cells(r_int_NroFil, r_int_ColumN) = Format(CDbl(pnl_MtoApl.Caption), "###,###.00")
      .Cells(r_int_NroFil, r_int_ColumC).HorizontalAlignment = xlHAlignCenter
      
      'MARCO
      .Cells(r_int_NroFil, r_int_ColumC).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumN).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumP).Borders(xlEdgeRight).Weight = xlThin
      .Range("C" & r_int_NroFil & ":P" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C" & r_int_NroFil & ":P" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      
      'Línea 16 - Cancelación de PBP x Cobrar
      .Range("S" & r_int_NroFil & ":Z" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumS).Font.Bold = True
      .Cells(r_int_NroFil, r_int_ColumS) = "Cancelación de PBP x Cobrar"
      
       'MARCO
      .Cells(r_int_NroFil, r_int_ColumS).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumZ).Borders(xlEdgeRight).Weight = xlThin
      .Range("S" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("S" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Cells(r_int_NroFil, r_int_ColumS).HorizontalAlignment = xlHAlignCenter
      
      r_int_NroFil = r_int_NroFil + 1
      
      'Línea 17 - Cancelación de PBP x Cobrar - Monto Aplicar
      .Range("S" & r_int_NroFil & ":V" & r_int_NroFil & "").Merge
      .Range("W" & r_int_NroFil & ":Y" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumS) = "Monto Aplicar"
      .Cells(r_int_NroFil, r_int_ColumW) = Format(CDbl(pnl_MtoApl.Caption), "###,###.00")
      .Cells(r_int_NroFil, r_int_ColumS).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, r_int_ColumW).HorizontalAlignment = xlHAlignRight
      
      'MARCO
      .Cells(r_int_NroFil, r_int_ColumS).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumV).Borders(xlEdgeRight).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumZ).Borders(xlEdgeRight).Weight = xlThin
      .Range("S" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("S" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      r_int_NroFil = r_int_NroFil + 1
      
      'Línea 18 - Cancelación de PBP x Cobrar - PBP
      .Range("S" & r_int_NroFil & ":V" & r_int_NroFil & "").Merge
      .Range("W" & r_int_NroFil & ":Y" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumS) = "PBP"
      .Cells(r_int_NroFil, r_int_ColumW) = Format(CDbl(pnl_CapPbp.Caption), "###,###.00")
      .Cells(r_int_NroFil, r_int_ColumS).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, r_int_ColumW).HorizontalAlignment = xlHAlignRight
      
      'MARCO
      .Cells(r_int_NroFil, r_int_ColumS).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumV).Borders(xlEdgeRight).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumZ).Borders(xlEdgeRight).Weight = xlThin
      .Range("S" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("S" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      r_int_NroFil = r_int_NroFil + 1
      
      'Linea 19 - Saldo Distribuir
      .Range("S" & r_int_NroFil & ":V" & r_int_NroFil & "").Merge
      .Range("W" & r_int_NroFil & ":Y" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumS).Font.Bold = True
      .Cells(r_int_NroFil, r_int_ColumS) = "Saldo Distribuir"
      .Cells(r_int_NroFil, r_int_ColumW) = Format(CDbl(pnl_MtoApl) - CDbl(pnl_CapPbp), "###,###.00")
      .Cells(r_int_NroFil, r_int_ColumS).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, r_int_ColumW).HorizontalAlignment = xlHAlignRight
      .Range("N" & r_int_NroFil & ":O" & r_int_NroFil & "").HorizontalAlignment = xlHAlignCenter
      
      'MARCO
      .Cells(r_int_NroFil, r_int_ColumS).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumV).Borders(xlEdgeRight).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumZ).Borders(xlEdgeRight).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumW).Font.Bold = True
      .Range("S" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      r_int_NroFil = r_int_NroFil + 2
         
      'Linea 20 - Monto de Prepago a Distribuir
      .Range("C" & r_int_NroFil & ":M" & r_int_NroFil & "").Merge
      .Range("N" & r_int_NroFil & ":P" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumC).Font.Bold = True
      .Cells(r_int_NroFil, r_int_ColumN).Font.Bold = True
      .Cells(r_int_NroFil, r_int_ColumC) = "Monto de Prepago a Distribuir"
      .Cells(r_int_NroFil, r_int_ColumN) = Format(CDbl(pnl_MtoApl_Final.Caption), "###,###.00")
      .Cells(r_int_NroFil, r_int_ColumC).HorizontalAlignment = xlHAlignCenter
      
      'MARCO
      .Cells(r_int_NroFil, r_int_ColumC).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumN).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumP).Borders(xlEdgeRight).Weight = xlThin
      .Range("C" & r_int_NroFil & ":P" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C" & r_int_NroFil & ":P" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      
       'Linea 20 - Distribución del Prepago
      .Range("S" & r_int_NroFil & ":Z" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumS).Font.Bold = True
      .Cells(r_int_NroFil, r_int_ColumS) = "Distribución del Prepago"
      
      'MARCO
      .Cells(r_int_NroFil, r_int_ColumS).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumZ).Borders(xlEdgeRight).Weight = xlThin
      .Range("S" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("S" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Cells(r_int_NroFil, r_int_ColumS).HorizontalAlignment = xlHAlignCenter
      r_int_NroFil = r_int_NroFil + 1
      
      'Linea 21 - Distribución TNC prepago
      .Range("S" & r_int_NroFil & ":V" & r_int_NroFil & "").Merge
      .Range("W" & r_int_NroFil & ":Y" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumS) = "TNC"
      .Cells(r_int_NroFil, r_int_ColumW) = Format(CDbl(pnl_AplTNC.Caption), "###,###.00")
      .Cells(r_int_NroFil, r_int_ColumS).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, r_int_ColumW).HorizontalAlignment = xlHAlignRight
      
      'MARCO
      .Cells(r_int_NroFil, r_int_ColumS).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumV).Borders(xlEdgeRight).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumZ).Borders(xlEdgeRight).Weight = xlThin
      .Range("S" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      r_int_NroFil = r_int_NroFil + 1
      
      'Linea 22 - Distribución TC prepago
      .Range("S" & r_int_NroFil & ":V" & r_int_NroFil & "").Merge
      .Range("W" & r_int_NroFil & ":Y" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumS) = "TC"
      .Cells(r_int_NroFil, r_int_ColumW) = Format(CDbl(pnl_ApliTC.Caption), "###,###.00")
      .Cells(r_int_NroFil, r_int_ColumS).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, r_int_ColumW).HorizontalAlignment = xlHAlignRight
      
      'MARCO
      .Cells(r_int_NroFil, r_int_ColumS).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumV).Borders(xlEdgeRight).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumZ).Borders(xlEdgeRight).Weight = xlThin
      .Range("S" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      r_int_NroFil = r_int_NroFil + 1
      
      'Linea 23 - Distribución total prepago
      .Range("S" & r_int_NroFil & ":V" & r_int_NroFil & "").Merge
      .Range("W" & r_int_NroFil & ":Y" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumS).Font.Bold = True
      .Cells(r_int_NroFil, r_int_ColumW).Font.Bold = True
      .Cells(r_int_NroFil, r_int_ColumS) = "Total"
      .Cells(r_int_NroFil, r_int_ColumW) = Format(CDbl(pnl_MtoApl_Final.Caption), "###,###.00")
      .Cells(r_int_NroFil, r_int_ColumS).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, r_int_ColumW).HorizontalAlignment = xlHAlignRight
      
      'MARCO
      .Cells(r_int_NroFil, r_int_ColumS).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumV).Borders(xlEdgeRight).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumZ).Borders(xlEdgeRight).Weight = xlThin
      .Range("S" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      r_int_NroFil = r_int_NroFil + 2
      
      'Linea 24 - Saldo después prepago
      .Range("C" & r_int_NroFil & ":M" & r_int_NroFil & "").Merge
      .Range("N" & r_int_NroFil & ":P" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumC).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, r_int_ColumC).Font.Bold = True
      .Cells(r_int_NroFil, r_int_ColumN).Font.Bold = True
      .Cells(r_int_NroFil, r_int_ColumC) = "Saldo después del prepago"
      .Cells(r_int_NroFil, r_int_ColumN) = Format(CDbl(pnl_NuevoSaldoTNC.Caption) + CDbl(pnl_NuevoSaldoTC.Caption), "###,###.00")
      
      'MARCO
      .Cells(r_int_NroFil, r_int_ColumC).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumN).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumP).Borders(xlEdgeRight).Weight = xlThin
      .Range("C" & r_int_NroFil & ":P" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C" & r_int_NroFil & ":P" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      
      'Linea 24 - Saldo después prepago
      .Range("S" & r_int_NroFil & ":Z" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumS).Font.Bold = True
      .Cells(r_int_NroFil, r_int_ColumS) = "Saldo después del prepago"
      
      'MARCO
      .Cells(r_int_NroFil, r_int_ColumS).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumZ).Borders(xlEdgeRight).Weight = xlThin
      .Range("S" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("S" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Cells(r_int_NroFil, r_int_ColumS).HorizontalAlignment = xlHAlignCenter
      r_int_NroFil = r_int_NroFil + 1
      
      'Linea 25 - Saldo TNC después prepago
      .Range("S" & r_int_NroFil & ":V" & r_int_NroFil & "").Merge
      .Range("W" & r_int_NroFil & ":Y" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumS) = "TNC"
      .Cells(r_int_NroFil, r_int_ColumW) = Format(CDbl(pnl_NuevoSaldoTNC.Caption), "###,###.00")
      .Cells(r_int_NroFil, r_int_ColumS).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, r_int_ColumW).HorizontalAlignment = xlHAlignRight
      
      'MARCO
      .Cells(r_int_NroFil, r_int_ColumS).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumV).Borders(xlEdgeRight).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumZ).Borders(xlEdgeRight).Weight = xlThin
      .Range("S" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      r_int_NroFil = r_int_NroFil + 1
      
      'Linea 26 - Saldo TC después prepago
      .Range("S" & r_int_NroFil & ":V" & r_int_NroFil & "").Merge
      .Range("W" & r_int_NroFil & ":Y" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumS) = "TC"
      .Cells(r_int_NroFil, r_int_ColumW) = Format(CDbl(pnl_NuevoSaldoTC.Caption), "###,###.00")
      .Cells(r_int_NroFil, r_int_ColumS).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, r_int_ColumW).HorizontalAlignment = xlHAlignRight
      
      'MARCO
      .Cells(r_int_NroFil, r_int_ColumS).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumV).Borders(xlEdgeRight).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumZ).Borders(xlEdgeRight).Weight = xlThin
      .Range("S" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      r_int_NroFil = r_int_NroFil + 1
     
      'Linea 27 - Total Saldo después prepago
      .Range("S" & r_int_NroFil & ":V" & r_int_NroFil & "").Merge
      .Range("W" & r_int_NroFil & ":Y" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumS).Font.Bold = True
      .Cells(r_int_NroFil, r_int_ColumW).Font.Bold = True
      .Cells(r_int_NroFil, r_int_ColumS) = "Total"
      .Cells(r_int_NroFil, r_int_ColumW) = Format(CDbl(pnl_NuevoSaldoTNC.Caption) + CDbl(pnl_NuevoSaldoTC.Caption), "###,###.00")
      .Cells(r_int_NroFil, r_int_ColumS).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, r_int_ColumW).HorizontalAlignment = xlHAlignRight
      
      'MARCO
      .Cells(r_int_NroFil, r_int_ColumS).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumV).Borders(xlEdgeRight).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumZ).Borders(xlEdgeRight).Weight = xlThin
      .Range("S" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      r_int_NroFil = r_int_NroFil + 2
      
      'Linea 28 - Importe nueva cuota
      .Cells(r_int_NroFil, r_int_ColumC) = "IMPORTE NUEVA CUOTA"
      .Range("C" & r_int_NroFil & ":M" & r_int_NroFil & "").Merge
      .Range("N" & r_int_NroFil & ":P" & r_int_NroFil & "").Merge
      .Cells(r_int_NroFil, r_int_ColumN) = Format(CDbl(pnl_NuevaCuota.Caption), "###,###.00")
      .Cells(r_int_NroFil, r_int_ColumC).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, r_int_ColumN).HorizontalAlignment = xlHAlignRight
      .Cells(r_int_NroFil, r_int_ColumC).Font.Bold = True
      .Cells(r_int_NroFil, r_int_ColumN).Font.Bold = True
'      .Cells(r_int_nrofil, r_int_ColumC).RowHeight = 12
      
      'MARCO
      .Cells(r_int_NroFil, r_int_ColumC).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumN).Borders(xlEdgeLeft).Weight = xlThin
      .Cells(r_int_NroFil, r_int_ColumP).Borders(xlEdgeRight).Weight = xlThin
      .Range("C" & r_int_NroFil & ":P" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range("C" & r_int_NroFil & ":P" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Cells(r_int_NroFil, r_int_ColumN).Select

'      'Linea - Valor inmueble
'      .Cells(r_int_nrofil, r_int_ColumC) = "Valor asegurable inmueble"
'      .Range("K" & r_int_nrofil & ":M" & r_int_nrofil & "").Merge
'      .Cells(r_int_nrofil, r_int_ColumK).Font.Bold = True
'      .Cells(r_int_nrofil, r_int_ColumK) = Format(CDbl(pnl_Val_AsgInm.Caption), "###,###.00")
'      r_int_nrofil = r_int_nrofil + 2

      If CInt(moddat_g_int_TipMon) = 1 Then
         r_obj_Excel.Selection.NumberFormat = "$###,##0.00_);[Red]($###,##0.00)"
         If cmb_TipPre.ItemData(cmb_TipPre.ListIndex) = 2 Then
            .Cells(43, 2) = "'" & "- Esta liquidación es válida solo a la fecha de corte."
            .Cells(44, 2) = "'" & "- Cualquier gasto adicional que no contemple la presente liquidación, será informado al cliente y tendrá que ser cancelado"
            .Cells(45, 2) = "'" & "  antes de realizar el abono que consigna la presente liquidación."
            .Cells(46, 2) = "'" & "- Realizar el depósito en la cuenta Nº 0011-0369-02-00090532 del BBVA Banco Continental."
            .Cells(47, 2) = "'" & "- Realizar el depósito en la cuenta Nº 011-369-000200090532-69 del BBVA Banco Continental desde otro Banco."
            .Range(.Cells(43, 2), .Cells(47, 2)).Font.Size = 11
            .Range(.Cells(43, 2), .Cells(47, 2)).Font.Bold = True
         Else
            .Cells(42, 2) = "'" & "- Esta liquidación es válida solo a la fecha de corte."
            .Cells(43, 2) = "'" & "- Cualquier gasto adicional que no contemple la presente liquidación, será informado al cliente y tendrá que ser cancelado"
            .Cells(44, 2) = "'" & "  antes de realizar el abono que consigna la presente liquidación."
            .Cells(45, 2) = "'" & "- Realizar el depósito en la cuenta Nº 0011-0369-02-00090532 del BBVA Banco Continental."
            .Cells(46, 2) = "'" & "- Realizar el depósito en la cuenta Nº 011-369-000200090532-69 del BBVA Banco Continental desde otro Banco."
            .Range(.Cells(42, 2), .Cells(46, 2)).Font.Size = 11
            .Range(.Cells(42, 2), .Cells(46, 2)).Font.Bold = True
         End If
      Else
         r_obj_Excel.Selection.NumberFormat = "[$$]#,##0.00;[Red][$$]#,##0.00"
         If cmb_TipPre.ItemData(cmb_TipPre.ListIndex) = 2 Then
            .Cells(43, 2) = "'" & "- Esta liquidación es válida solo a la fecha de corte."
            .Cells(44, 2) = "'" & "- Cualquier gasto adicional que no contemple la presente liquidación, será informado al cliente y tendrá que ser cancelado"
            .Cells(45, 2) = "'" & "  antes de realizar el abono que consigna la presente liquidación."
            .Cells(46, 2) = "'" & "- Realizar el depósito en la cuenta Nº 0011-0369-02-00090540 del BBVA Banco Continental."
            .Range(.Cells(43, 2), .Cells(46, 2)).Font.Size = 11
            .Range(.Cells(43, 2), .Cells(46, 2)).Font.Bold = True
         Else
            .Cells(42, 2) = "'" & "- Esta liquidación es válida solo a la fecha de corte."
            .Cells(43, 2) = "'" & "- Cualquier gasto adicional que no contemple la presente liquidación, será informado al cliente y tendrá que ser cancelado"
            .Cells(44, 2) = "'" & "  antes de realizar el abono que consigna la presente liquidación."
            .Cells(45, 2) = "'" & "- Realizar el depósito en la cuenta Nº 0011-0369-02-00090540 del BBVA Banco Continental."
            .Range(.Cells(42, 2), .Cells(45, 2)).Font.Size = 11
            .Range(.Cells(42, 2), .Cells(45, 2)).Font.Bold = True
         End If
      End If
   End With

   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub fs_RptPar_Mivivienda_ant()
Dim r_obj_Excel      As Excel.Application
Dim r_int_NroFil     As Integer
Dim r_int_ColumC     As Integer
Dim r_int_ColumF     As Integer
Dim r_int_ColumK     As Integer
Dim r_int_ColumL     As Integer
Dim r_int_ColumM     As Integer
Dim r_int_ColumN     As Integer
Dim r_int_ColumO     As Integer
Dim r_int_ColumQ     As Integer
Dim r_int_ColumT     As Integer
Dim r_int_ColumV     As Integer
Dim r_int_ColumW     As Integer
Dim r_int_ColumX     As Integer
Dim r_int_ColumZ     As Integer
    
    r_int_NroFil = 2
    r_int_ColumC = 3
    r_int_ColumF = 6
    r_int_ColumK = 11
    r_int_ColumL = 12
    r_int_ColumM = 13
    r_int_ColumN = 14
    r_int_ColumO = 15
    r_int_ColumQ = 17
    r_int_ColumT = 20
    r_int_ColumV = 22
    r_int_ColumW = 23
    r_int_ColumX = 24
    r_int_ColumZ = 26
    
    Set r_obj_Excel = New Excel.Application
    r_obj_Excel.SheetsInNewWorkbook = 1
    r_obj_Excel.Workbooks.Add
    
    With r_obj_Excel.ActiveSheet
        'CENTRADO DE LA PAGINA
        '.PageSetup.CenterHorizontally = True
        .PageSetup.Orientation = xlLandscape
        
        'MARGENES
        .PageSetup.LeftMargin = Application.CentimetersToPoints(1)
        .PageSetup.RightMargin = Application.CentimetersToPoints(1)
        .PageSetup.TopMargin = Application.CentimetersToPoints(1)
        .PageSetup.BottomMargin = Application.CentimetersToPoints(1)
        .Range("A1:A1").ColumnWidth = 1.14
        .Range("B1:AB1").ColumnWidth = 3.57
        .Range("A8:A41").RowHeight = 12
        
        'BORDERS
        .Range("B2:AA2").Borders(xlEdgeTop).Weight = xlMedium
        .Range("B3:AA3").Borders(xlEdgeTop).Weight = xlMedium
        .Range("B4:AA4").Borders(xlEdgeTop).Weight = xlMedium
        .Range("B5:AA5").Borders(xlEdgeTop).Weight = xlMedium
        .Range("B40:AA40").Borders(xlEdgeBottom).Weight = xlMedium
        .Range("B2:AA40").Borders(xlEdgeLeft).Weight = xlMedium
        .Range("AA2:AA40").Borders(xlEdgeRight).Weight = xlMedium
        
        'Font
        .Range("B1:AA44").Font.Name = "Arial"
        .Range("B1:AA40").Font.Size = 9
        
        'Linea 2 - Titulo
        .Range("B2:AA2").Merge
        .Range("B2") = "LIQUIDACION PREPAGO PARCIAL - " & moddat_g_str_NomPrd
        .Range("B2:AA2").HorizontalAlignment = xlHAlignCenter
        .Range("B2:AA2").Font.Size = 12
        .Range("B4:AA4").Font.Size = 12
        .Range("B6:AA6").Font.Size = 12
        .Range("B2:AA2").Font.Bold = True
        .Range("B4:AA4").Font.Bold = True
        .Range("B6:AA6").Font.Bold = True
        r_int_NroFil = r_int_NroFil + 2
        
        'Linea 4 - Datos Cliente
        .Cells(r_int_NroFil, r_int_ColumC) = "Cliente:"
        .Cells(r_int_NroFil, r_int_ColumF) = moddat_g_str_NomCli
        .Cells(r_int_NroFil, r_int_ColumV) = "DNI:"
        .Cells(r_int_NroFil, r_int_ColumX) = "'" & moddat_g_str_NumDoc
        .Cells(r_int_NroFil, r_int_ColumX).HorizontalAlignment = xlHAlignLeft
        r_int_NroFil = r_int_NroFil + 2
        
        'Linea 6 - Numero de Operacion
        .Range("C" & r_int_NroFil & ":M" & (r_int_NroFil) & "").Merge
        .Cells(r_int_NroFil, r_int_ColumC) = "N° de Operación: " & Mid(moddat_g_str_NumOpe, 1, 3) & "-" & Mid(moddat_g_str_NumOpe, 4, 2) & "-" & Mid(moddat_g_str_NumOpe, 6, 5)
        .Cells(r_int_NroFil, r_int_ColumC).Font.Bold = True
        .Cells(r_int_NroFil, r_int_ColumO) = "Moneda: " & moddat_g_str_Moneda
        .Cells(r_int_NroFil, r_int_ColumO).Font.Bold = True
        r_int_NroFil = r_int_NroFil + 2
        
        'Linea 8 - Ultimo pago TC, subtitulo saldo
        .Cells(r_int_NroFil, r_int_ColumC) = "Fecha de Desembolso o última cuota TC (A)"
        .Range("L" & r_int_NroFil & ":M" & r_int_NroFil & "").Merge
        .Cells(r_int_NroFil, r_int_ColumL) = "'" & Format(CDate(pnl_UltPagTC.Caption), "dd-mm-yy")
        .Cells(r_int_NroFil, r_int_ColumL).HorizontalAlignment = xlHAlignRight
        .Range("T" & r_int_NroFil & ":Z" & r_int_NroFil & "").Merge
        .Cells(r_int_NroFil, r_int_ColumT).Font.Bold = True
        .Cells(r_int_NroFil, r_int_ColumT) = "Saldo antes del prepago"
        
        'MARCO
        .Cells(r_int_NroFil, r_int_ColumT).Borders(xlEdgeLeft).Weight = xlThin
        .Cells(r_int_NroFil, r_int_ColumZ).Borders(xlEdgeRight).Weight = xlThin
        .Range("T" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
        .Range("T" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Cells(r_int_NroFil, r_int_ColumT).HorizontalAlignment = xlHAlignCenter
        r_int_NroFil = r_int_NroFil + 1
        
        'Linea 9 - Ultimo Pago TNC, Saldo Total, Saldo TNC
        r_obj_Excel.ActiveSheet.Cells(r_int_NroFil, r_int_ColumC) = "Saldo al (B)"
        .Range("L" & r_int_NroFil & ":M" & r_int_NroFil & "").Merge
        .Cells(r_int_NroFil, r_int_ColumL).HorizontalAlignment = xlHAlignRight
        .Cells(r_int_NroFil, r_int_ColumT).HorizontalAlignment = xlHAlignRight
        .Cells(r_int_NroFil, r_int_ColumL) = "'" & Format(CDate(pnl_UltPagTNC.Caption), "dd-mm-yy")
        
        .Range("N" & r_int_NroFil & ":P" & r_int_NroFil & "").Merge
        .Cells(r_int_NroFil, r_int_ColumN).Font.Bold = True
        .Cells(r_int_NroFil, r_int_ColumN).HorizontalAlignment = xlHAlignRight
        .Cells(r_int_NroFil, r_int_ColumN) = Format(CDbl(pnl_SaldoTNC2) + CDbl(pnl_SaldoTC2), "###,###.00")
 
        'MARCO
        .Cells(r_int_NroFil, r_int_ColumN).Borders(xlEdgeLeft).Weight = xlThin
        .Cells(r_int_NroFil, r_int_ColumQ).Borders(xlEdgeRight).Weight = xlThin
        .Range("N" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
        .Range("N" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
        
        .Range("T" & r_int_NroFil & ":V" & r_int_NroFil & "").Merge
        .Range("W" & r_int_NroFil & ":Y" & r_int_NroFil & "").Merge
        .Cells(r_int_NroFil, r_int_ColumT) = "TNC"
        .Cells(r_int_NroFil, r_int_ColumW) = Format(CDbl(pnl_SaldoTNC1.Caption), "###,###.00")
        .Cells(r_int_NroFil, r_int_ColumT).HorizontalAlignment = xlHAlignCenter
        .Cells(r_int_NroFil, r_int_ColumW).HorizontalAlignment = xlHAlignRight
        
        'MARCO
        .Cells(r_int_NroFil, r_int_ColumT).Borders(xlEdgeLeft).Weight = xlThin
        .Cells(r_int_NroFil, r_int_ColumV).Borders(xlEdgeRight).Weight = xlThin
        .Cells(r_int_NroFil, r_int_ColumZ).Borders(xlEdgeRight).Weight = xlThin
        .Range("T" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
        .Range("T" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
        r_int_NroFil = r_int_NroFil + 1
        
        'Linea 10 - Fecha Prepago, saldo TC
        .Cells(r_int_NroFil, r_int_ColumC) = "Fecha de corte (fecha del prepago) (C)"
        .Range("C" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range("L" & r_int_NroFil & ":M" & r_int_NroFil & "").Merge
        .Cells(r_int_NroFil, r_int_ColumL) = "'" & Format(CDate(ipp_FecPre), "dd-mm-yy")
        .Cells(r_int_NroFil, r_int_ColumL).HorizontalAlignment = xlHAlignRight
        
        .Range("T" & r_int_NroFil & ":V" & r_int_NroFil & "").Merge
        .Range("W" & r_int_NroFil & ":Y" & r_int_NroFil & "").Merge
        .Cells(r_int_NroFil, r_int_ColumT) = "TC"
        .Cells(r_int_NroFil, r_int_ColumW) = Format(CDbl(pnl_SaldoTC1.Caption), "###,###.00")
        .Cells(r_int_NroFil, r_int_ColumT).HorizontalAlignment = xlHAlignCenter
        .Cells(r_int_NroFil, r_int_ColumW).HorizontalAlignment = xlHAlignRight
        
        'MARCO
        .Cells(r_int_NroFil, r_int_ColumT).Borders(xlEdgeLeft).Weight = xlThin
        .Cells(r_int_NroFil, r_int_ColumV).Borders(xlEdgeRight).Weight = xlThin
        .Cells(r_int_NroFil, r_int_ColumZ).Borders(xlEdgeRight).Weight = xlThin
        .Range("T" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range("C" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
        r_int_NroFil = r_int_NroFil + 1
        
        'Linea 11 - Total Saldo
        .Range("T" & r_int_NroFil & ":V" & r_int_NroFil & "").Merge
        .Range("W" & r_int_NroFil & ":Y" & r_int_NroFil & "").Merge
        .Cells(r_int_NroFil, r_int_ColumT).Font.Bold = True
        .Cells(r_int_NroFil, r_int_ColumT) = "Total"
        .Cells(r_int_NroFil, r_int_ColumW) = Format(CDbl(pnl_SaldoTNC2) + CDbl(pnl_SaldoTC2), "###,###.00")
        .Cells(r_int_NroFil, r_int_ColumT).HorizontalAlignment = xlHAlignCenter
        .Cells(r_int_NroFil, r_int_ColumW).HorizontalAlignment = xlHAlignRight
        .Range("N" & r_int_NroFil & ":O" & r_int_NroFil & "").HorizontalAlignment = xlHAlignCenter
        
        'MARCO
        .Cells(r_int_NroFil, r_int_ColumT).Borders(xlEdgeLeft).Weight = xlThin
        .Cells(r_int_NroFil, r_int_ColumV).Borders(xlEdgeRight).Weight = xlThin
        .Cells(r_int_NroFil, r_int_ColumZ).Borders(xlEdgeRight).Weight = xlThin
        .Cells(r_int_NroFil, r_int_ColumW).Font.Bold = True
        .Range("T" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
        r_int_NroFil = r_int_NroFil + 1
        
        'Linea 12 - Dias TNC
        .Cells(r_int_NroFil, r_int_ColumC) = "Días de interés TNC (C-B)"
        .Range("L" & r_int_NroFil & ":M" & r_int_NroFil & "").Merge
        .Cells(r_int_NroFil, r_int_ColumL) = CInt(pnl_DiasTNC.Caption)
        r_int_NroFil = r_int_NroFil + 1
        
        'Linea 13 - Dias TC
        .Cells(r_int_NroFil, r_int_ColumC) = "Días de interés TC (C-A)"
        .Range("L" & r_int_NroFil & ":M" & r_int_NroFil & "").Merge
        .Cells(r_int_NroFil, r_int_ColumL) = CInt(pnl_DiasTC.Caption)
        r_int_NroFil = r_int_NroFil + 2
        
        'Linea 15 - Valor inmueble
        .Cells(r_int_NroFil, r_int_ColumC) = "Valor asegurable inmueble"
        .Range("K" & r_int_NroFil & ":M" & r_int_NroFil & "").Merge
        .Cells(r_int_NroFil, r_int_ColumK).Font.Bold = True
        .Cells(r_int_NroFil, r_int_ColumK) = Format(CDbl(pnl_Val_AsgInm.Caption), "###,###.00")
        r_int_NroFil = r_int_NroFil + 2
        
        'Linea 17 - Tasa interes anual
        .Cells(r_int_NroFil, r_int_ColumC) = "Tasa de interés anual (%)"
        .Range("L" & r_int_NroFil & ":M" & r_int_NroFil & "").Merge
        .Cells(r_int_NroFil, r_int_ColumL).HorizontalAlignment = xlHAlignRight
        .Cells(r_int_NroFil, r_int_ColumL) = l_dbl_TasInt
        .Cells(r_int_NroFil, r_int_ColumL).Select
        r_obj_Excel.Selection.NumberFormat = "###0.0000"
        r_int_NroFil = r_int_NroFil + 1
        
        'Linea 18 - Tasa seguro inmueble
        .Cells(r_int_NroFil, r_int_ColumC) = "Tasa de Seguro Inmueble (%)"
        .Range("L" & r_int_NroFil & ":M" & r_int_NroFil & "").Merge
        .Cells(r_int_NroFil, r_int_ColumL).HorizontalAlignment = xlHAlignRight
        .Cells(r_int_NroFil, r_int_ColumL) = l_dbl_SegInm
        .Range("L" & r_int_NroFil & ":L" & r_int_NroFil & "").Select
        r_obj_Excel.Selection.NumberFormat = "###0.0000"
        r_int_NroFil = r_int_NroFil + 1
        
        'Linea 19 - Tasa seguro desgravamen
        .Cells(r_int_NroFil, r_int_ColumC) = "Tasa de Seguro Desgravamen (%)"
        .Range("L" & r_int_NroFil & ":M" & r_int_NroFil & "").Merge
        .Cells(r_int_NroFil, r_int_ColumL).HorizontalAlignment = xlHAlignRight
        .Cells(r_int_NroFil, r_int_ColumL) = l_dbl_SegDes
        .Range("L" & r_int_NroFil & ":L" & r_int_NroFil & "").Select
        r_obj_Excel.Selection.NumberFormat = "###0.0000"
        r_int_NroFil = r_int_NroFil + 2
        
        'Linea 21 - Monto Deposito
        .Range("C" & r_int_NroFil & ":M" & (r_int_NroFil) & "").Merge
        .Range("N" & r_int_NroFil & ":P" & (r_int_NroFil) & "").Merge
        .Cells(r_int_NroFil, r_int_ColumC) = "Monto Depositado"
        .Cells(r_int_NroFil, r_int_ColumN) = Format(CDbl(pnl_Mto_Deposito.Caption), "###,###.00")
        .Cells(r_int_NroFil, r_int_ColumC).Font.Bold = True
        .Cells(r_int_NroFil, r_int_ColumN).Font.Bold = True
        
         'MARCO
         .Cells(r_int_NroFil, r_int_ColumC).HorizontalAlignment = xlHAlignCenter
         .Cells(r_int_NroFil, r_int_ColumN).HorizontalAlignment = xlHAlignRight
         .Cells(r_int_NroFil, r_int_ColumC).VerticalAlignment = xlHAlignCenter
         .Cells(r_int_NroFil, r_int_ColumN).VerticalAlignment = xlHAlignCenter
         .Cells(r_int_NroFil, r_int_ColumC).Borders(xlEdgeLeft).Weight = xlThin
         .Cells(r_int_NroFil, r_int_ColumQ).Borders(xlEdgeRight).Weight = xlThin
         .Cells(r_int_NroFil, r_int_ColumM).Borders(xlEdgeRight).Weight = xlThin
         .Range("C" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
         .Cells(r_int_NroFil, r_int_ColumC).Borders(xlEdgeLeft).Weight = xlThin
         .Cells(r_int_NroFil, r_int_ColumQ).Borders(xlEdgeRight).Weight = xlThin
         .Cells(r_int_NroFil, r_int_ColumM).Borders(xlEdgeRight).Weight = xlThin
         .Range("C" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
         r_int_NroFil = r_int_NroFil + 1
         
         'Linea 22 - Interes TNC
         .Cells(r_int_NroFil, r_int_ColumC) = "Intereses TNC a la fecha"
         .Range("L" & r_int_NroFil & ":M" & r_int_NroFil & "").Merge
         .Cells(r_int_NroFil, r_int_ColumL) = CDbl(pnl_InteresTNC.Caption)
         .Cells(r_int_NroFil, r_int_ColumL).Select
         r_obj_Excel.Selection.NumberFormat = "###0.00"
         r_int_NroFil = r_int_NroFil + 1
         
         'Linea 23 - Interes TC
         .Cells(r_int_NroFil, r_int_ColumC) = "Intereses TC a la fecha"
         .Range("L" & r_int_NroFil & ":M" & r_int_NroFil & "").Merge
         .Cells(r_int_NroFil, r_int_ColumL) = CDbl(pnl_InteresTC.Caption)
         .Cells(r_int_NroFil, r_int_ColumL).Select
         r_obj_Excel.Selection.NumberFormat = "###0.00"
         r_int_NroFil = r_int_NroFil + 1
         
         'Linea 24 - Seguro desgravamen
         .Cells(r_int_NroFil, r_int_ColumC) = "Seguro Desgravamen"
         .Range("L" & r_int_NroFil & ":M" & r_int_NroFil & "").Merge
         .Cells(r_int_NroFil, r_int_ColumL) = CDbl(pnl_SegDes.Caption)
         .Cells(r_int_NroFil, r_int_ColumL).Select
         r_obj_Excel.Selection.NumberFormat = "###0.00"
         r_int_NroFil = r_int_NroFil + 1
         
         'Linea 25 - Seguro inmueble
         .Cells(r_int_NroFil, r_int_ColumC) = "Seguro Inmueble"
         .Range("L" & r_int_NroFil & ":M" & r_int_NroFil & "").Merge
         .Cells(r_int_NroFil, r_int_ColumL) = CDbl(pnl_SegInm.Caption)
         .Cells(r_int_NroFil, r_int_ColumL).Select
         r_obj_Excel.Selection.NumberFormat = "###0.00"
         r_int_NroFil = r_int_NroFil + 1
         
         'Linea 26 - Capital PBP
         .Cells(r_int_NroFil, r_int_ColumC) = "Capital PBP"
         .Range("L" & r_int_NroFil & ":M" & r_int_NroFil & "").Merge
         .Cells(r_int_NroFil, r_int_ColumL) = CDbl(pnl_CapPbp.Caption)
         .Cells(r_int_NroFil, r_int_ColumL).Select
         r_obj_Excel.Selection.NumberFormat = "###0.00"
         r_int_NroFil = r_int_NroFil + 1
         
         'Linea 27 - Capital PBP
         .Cells(r_int_NroFil, r_int_ColumC) = "Interes PBP"
         .Range("L" & r_int_NroFil & ":M" & r_int_NroFil & "").Merge
         .Cells(r_int_NroFil, r_int_ColumL) = CDbl(pnl_IntPbp.Caption)
         .Cells(r_int_NroFil, r_int_ColumL).Select
         r_obj_Excel.Selection.NumberFormat = "###0.00"
         r_int_NroFil = r_int_NroFil + 1
         
         'Linea 28 - Total gastos
         .Cells(r_int_NroFil, r_int_ColumC) = "Total de Interés y seguros"
         .Range("C" & r_int_NroFil & ":M" & r_int_NroFil & "").Merge
         .Range("N" & r_int_NroFil & ":P" & r_int_NroFil & "").Merge
         .Cells(r_int_NroFil, r_int_ColumN) = CDbl(CDbl(pnl_InteresTNC.Caption) + CDbl(pnl_InteresTC.Caption) + CDbl(pnl_SegDes.Caption) + CDbl(pnl_SegInm.Caption) + CDbl(pnl_CapPbp.Caption) + CDbl(pnl_IntPbp.Caption))
         .Cells(r_int_NroFil, r_int_ColumC).HorizontalAlignment = xlHAlignCenter
         .Cells(r_int_NroFil, r_int_ColumN).HorizontalAlignment = xlHAlignRight
         
         'MARCO
         .Cells(r_int_NroFil, r_int_ColumC).Borders(xlEdgeLeft).Weight = xlThin
         .Cells(r_int_NroFil, r_int_ColumN).Borders(xlEdgeLeft).Weight = xlThin
         .Cells(r_int_NroFil, r_int_ColumQ).Borders(xlEdgeRight).Weight = xlThin
         .Range("C" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
         .Range("C" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
         r_int_NroFil = r_int_NroFil + 2
         
         'Linea 31 - Prepago a Aplicar
         .Range("C" & r_int_NroFil & ":M" & r_int_NroFil & "").Merge
         .Range("N" & r_int_NroFil & ":P" & r_int_NroFil & "").Merge
         .Cells(r_int_NroFil, r_int_ColumC).Font.Bold = True
         .Cells(r_int_NroFil, r_int_ColumN).Font.Bold = True
         .Cells(r_int_NroFil, r_int_ColumC) = "Monto de Prepago a Aplicar"
         .Cells(r_int_NroFil, r_int_ColumN) = Format(CDbl(pnl_MtoApl_Final.Caption), "###,###.00")
         .Cells(r_int_NroFil, r_int_ColumC).HorizontalAlignment = xlHAlignCenter
         
         'MARCO
         .Cells(r_int_NroFil, r_int_ColumC).Borders(xlEdgeLeft).Weight = xlThin
         .Cells(r_int_NroFil, r_int_ColumN).Borders(xlEdgeLeft).Weight = xlThin
         .Cells(r_int_NroFil, r_int_ColumQ).Borders(xlEdgeRight).Weight = xlThin
         .Range("C" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
         .Range("C" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
         
         .Range("T" & r_int_NroFil & ":Z" & r_int_NroFil & "").Merge
         .Cells(r_int_NroFil, r_int_ColumT).Font.Bold = True
         .Cells(r_int_NroFil, r_int_ColumT) = "Distribucion del Prepago"
         
         'MARCO
         .Cells(r_int_NroFil, r_int_ColumT).Borders(xlEdgeLeft).Weight = xlThin
         .Cells(r_int_NroFil, r_int_ColumZ).Borders(xlEdgeRight).Weight = xlThin
         .Range("T" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
         .Range("T" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
         .Cells(r_int_NroFil, r_int_ColumT).HorizontalAlignment = xlHAlignCenter
         r_int_NroFil = r_int_NroFil + 1
         
         'Linea 32 - Distibucion TNC prepago
         .Range("T" & r_int_NroFil & ":V" & r_int_NroFil & "").Merge
         .Range("W" & r_int_NroFil & ":Y" & r_int_NroFil & "").Merge
         .Cells(r_int_NroFil, r_int_ColumT) = "TNC"
         .Cells(r_int_NroFil, r_int_ColumW) = Format(CDbl(pnl_AplTNC.Caption), "###,###.00")
         .Cells(r_int_NroFil, r_int_ColumT).HorizontalAlignment = xlHAlignCenter
         .Cells(r_int_NroFil, r_int_ColumW).HorizontalAlignment = xlHAlignRight
         
         'MARCO
         .Cells(r_int_NroFil, r_int_ColumT).Borders(xlEdgeLeft).Weight = xlThin
         .Cells(r_int_NroFil, r_int_ColumV).Borders(xlEdgeRight).Weight = xlThin
         .Cells(r_int_NroFil, r_int_ColumZ).Borders(xlEdgeRight).Weight = xlThin
         .Range("T" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
         r_int_NroFil = r_int_NroFil + 1
         
         'Linea 33 - Distibucion TC prepago
         .Range("T" & r_int_NroFil & ":V" & r_int_NroFil & "").Merge
         .Range("W" & r_int_NroFil & ":Y" & r_int_NroFil & "").Merge
         .Cells(r_int_NroFil, r_int_ColumT) = "TC"
         .Cells(r_int_NroFil, r_int_ColumW) = Format(CDbl(pnl_ApliTC.Caption), "###,###.00")
         .Cells(r_int_NroFil, r_int_ColumT).HorizontalAlignment = xlHAlignCenter
         .Cells(r_int_NroFil, r_int_ColumW).HorizontalAlignment = xlHAlignRight
         
         'MARCO
         .Cells(r_int_NroFil, r_int_ColumT).Borders(xlEdgeLeft).Weight = xlThin
         .Cells(r_int_NroFil, r_int_ColumV).Borders(xlEdgeRight).Weight = xlThin
         .Cells(r_int_NroFil, r_int_ColumZ).Borders(xlEdgeRight).Weight = xlThin
         .Range("T" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
         r_int_NroFil = r_int_NroFil + 1
         
         'Linea 34 - Distibucion total prepago
         .Range("T" & r_int_NroFil & ":V" & r_int_NroFil & "").Merge
         .Range("W" & r_int_NroFil & ":Y" & r_int_NroFil & "").Merge
         .Cells(r_int_NroFil, r_int_ColumT).Font.Bold = True
         .Cells(r_int_NroFil, r_int_ColumW).Font.Bold = True
         .Cells(r_int_NroFil, r_int_ColumT) = "Total"
         .Cells(r_int_NroFil, r_int_ColumW) = Format(CDbl(pnl_MtoApl_Final.Caption), "###,###.00")
         .Cells(r_int_NroFil, r_int_ColumT).HorizontalAlignment = xlHAlignCenter
         .Cells(r_int_NroFil, r_int_ColumW).HorizontalAlignment = xlHAlignRight
         
         'MARCO
         .Cells(r_int_NroFil, r_int_ColumT).Borders(xlEdgeLeft).Weight = xlThin
         .Cells(r_int_NroFil, r_int_ColumV).Borders(xlEdgeRight).Weight = xlThin
         .Cells(r_int_NroFil, r_int_ColumZ).Borders(xlEdgeRight).Weight = xlThin
         .Range("T" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
         r_int_NroFil = r_int_NroFil + 2
         
         'Linea 36 - Saldo despues prepago
         .Range("C" & r_int_NroFil & ":M" & r_int_NroFil & "").Merge
         .Range("N" & r_int_NroFil & ":P" & r_int_NroFil & "").Merge
         .Cells(r_int_NroFil, r_int_ColumC).HorizontalAlignment = xlHAlignCenter
         .Cells(r_int_NroFil, r_int_ColumC).Font.Bold = True
         .Cells(r_int_NroFil, r_int_ColumN).Font.Bold = True
         .Cells(r_int_NroFil, r_int_ColumC) = "Saldo después del prepago"
         .Cells(r_int_NroFil, r_int_ColumN) = Format(CDbl(pnl_NuevoSaldoTNC.Caption) + CDbl(pnl_NuevoSaldoTC.Caption), "###,###.00")
         
         'MARCO
         .Cells(r_int_NroFil, r_int_ColumC).Borders(xlEdgeLeft).Weight = xlThin
         .Cells(r_int_NroFil, r_int_ColumN).Borders(xlEdgeLeft).Weight = xlThin
         .Cells(r_int_NroFil, r_int_ColumQ).Borders(xlEdgeRight).Weight = xlThin
         .Range("C" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
         .Range("C" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
         
         .Range("T" & r_int_NroFil & ":Z" & r_int_NroFil & "").Merge
         .Cells(r_int_NroFil, r_int_ColumT).Font.Bold = True
         .Cells(r_int_NroFil, r_int_ColumT) = "Saldo despues del prepago"
         
         'MARCO
         .Cells(r_int_NroFil, r_int_ColumT).Borders(xlEdgeLeft).Weight = xlThin
         .Cells(r_int_NroFil, r_int_ColumZ).Borders(xlEdgeRight).Weight = xlThin
         .Range("T" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
         .Range("T" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
         .Cells(r_int_NroFil, r_int_ColumT).HorizontalAlignment = xlHAlignCenter
         r_int_NroFil = r_int_NroFil + 1
         
         'Linea 37 - Saldo TNC despues prepago
         .Range("T" & r_int_NroFil & ":V" & r_int_NroFil & "").Merge
         .Range("W" & r_int_NroFil & ":Y" & r_int_NroFil & "").Merge
         .Cells(r_int_NroFil, r_int_ColumT) = "TNC"
         .Cells(r_int_NroFil, r_int_ColumW) = Format(CDbl(pnl_NuevoSaldoTNC.Caption), "###,###.00")
         .Cells(r_int_NroFil, r_int_ColumT).HorizontalAlignment = xlHAlignCenter
         .Cells(r_int_NroFil, r_int_ColumW).HorizontalAlignment = xlHAlignRight
         
         'MARCO
         .Cells(r_int_NroFil, r_int_ColumT).Borders(xlEdgeLeft).Weight = xlThin
         .Cells(r_int_NroFil, r_int_ColumV).Borders(xlEdgeRight).Weight = xlThin
         .Cells(r_int_NroFil, r_int_ColumZ).Borders(xlEdgeRight).Weight = xlThin
         .Range("T" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
         r_int_NroFil = r_int_NroFil + 1
         
         'Linea 38 - Saldo TC despues prepago
         .Range("T" & r_int_NroFil & ":V" & r_int_NroFil & "").Merge
         .Range("W" & r_int_NroFil & ":Y" & r_int_NroFil & "").Merge
         .Cells(r_int_NroFil, r_int_ColumT) = "TC"
         .Cells(r_int_NroFil, r_int_ColumW) = Format(CDbl(pnl_NuevoSaldoTC.Caption), "###,###.00")
         .Cells(r_int_NroFil, r_int_ColumT).HorizontalAlignment = xlHAlignCenter
         .Cells(r_int_NroFil, r_int_ColumW).HorizontalAlignment = xlHAlignRight
         
         'MARCO
         .Cells(r_int_NroFil, r_int_ColumT).Borders(xlEdgeLeft).Weight = xlThin
         .Cells(r_int_NroFil, r_int_ColumV).Borders(xlEdgeRight).Weight = xlThin
         .Cells(r_int_NroFil, r_int_ColumZ).Borders(xlEdgeRight).Weight = xlThin
         .Range("T" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
         r_int_NroFil = r_int_NroFil + 1
         
         'Linea 39 - Saldo despues prepago
         .Range("T" & r_int_NroFil & ":V" & r_int_NroFil & "").Merge
         .Range("W" & r_int_NroFil & ":Y" & r_int_NroFil & "").Merge
         .Cells(r_int_NroFil, r_int_ColumT).Font.Bold = True
         .Cells(r_int_NroFil, r_int_ColumW).Font.Bold = True
         .Cells(r_int_NroFil, r_int_ColumT) = "Total"
         .Cells(r_int_NroFil, r_int_ColumW) = Format(CDbl(pnl_NuevoSaldoTNC.Caption) + CDbl(pnl_NuevoSaldoTC.Caption), "###,###.00")
         .Cells(r_int_NroFil, r_int_ColumT).HorizontalAlignment = xlHAlignCenter
         .Cells(r_int_NroFil, r_int_ColumW).HorizontalAlignment = xlHAlignRight
         
         'MARCO
         .Cells(r_int_NroFil, r_int_ColumT).Borders(xlEdgeLeft).Weight = xlThin
         .Cells(r_int_NroFil, r_int_ColumV).Borders(xlEdgeRight).Weight = xlThin
         .Cells(r_int_NroFil, r_int_ColumZ).Borders(xlEdgeRight).Weight = xlThin
         .Range("T" & r_int_NroFil & ":Z" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
         r_int_NroFil = r_int_NroFil + 1
         
         'Linea 40 - Importe nueva cuota
         .Cells(r_int_NroFil, r_int_ColumC) = "IMPORTE NUEVA CUOTA"
         .Range("C" & r_int_NroFil & ":M" & r_int_NroFil & "").Merge
         .Range("N" & r_int_NroFil & ":P" & r_int_NroFil & "").Merge
         .Cells(r_int_NroFil, r_int_ColumN) = Format(CDbl(pnl_NuevaCuota.Caption), "###,###.00")
         .Cells(r_int_NroFil, r_int_ColumC).HorizontalAlignment = xlHAlignCenter
         .Cells(r_int_NroFil, r_int_ColumN).HorizontalAlignment = xlHAlignRight
         .Cells(r_int_NroFil, r_int_ColumC).Font.Bold = True
         .Cells(r_int_NroFil, r_int_ColumN).Font.Bold = True
         .Cells(r_int_NroFil, r_int_ColumC).RowHeight = 12
         
         'MARCO
         .Cells(r_int_NroFil, r_int_ColumC).Borders(xlEdgeLeft).Weight = xlThin
         .Cells(r_int_NroFil, r_int_ColumN).Borders(xlEdgeLeft).Weight = xlThin
         .Cells(r_int_NroFil, r_int_ColumQ).Borders(xlEdgeRight).Weight = xlThin
         .Range("C" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
         .Range("C" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
         .Cells(r_int_NroFil, r_int_ColumN).Select
        
       If CInt(moddat_g_int_TipMon) = 1 Then
          r_obj_Excel.Selection.NumberFormat = "$###,##0.00_);[Red]($###,##0.00)"
          .Cells(42, 2) = "Realizar el depósito en la cuenta Nº 0011-0369-02-00090532 del BBVA Banco Continental."
          .Cells(42, 2).Font.Size = 12
          .Cells(42, 2).Font.Bold = True
       Else
          r_obj_Excel.Selection.NumberFormat = "[$$]#,##0.00;[Red][$$]#,##0.00"
          .Cells(42, 2) = "Realizar el depósito en la cuenta Nº 0011-0369-02-00090540 del BBVA Banco Continental."
          .Cells(42, 2).Font.Size = 12
          .Cells(42, 2).Font.Bold = True
       End If
    End With

    r_obj_Excel.Visible = True
    Set r_obj_Excel = Nothing
End Sub

