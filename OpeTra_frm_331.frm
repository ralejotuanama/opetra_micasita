VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_Con_PrePgo_04 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form4"
   ClientHeight    =   9720
   ClientLeft      =   2295
   ClientTop       =   2805
   ClientWidth     =   11685
   Icon            =   "OpeTra_frm_331.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9720
   ScaleWidth      =   11685
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel111 
      Height          =   9735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11685
      _Version        =   65536
      _ExtentX        =   20611
      _ExtentY        =   17171
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
         Height          =   1935
         Left            =   30
         TabIndex        =   1
         Top             =   7770
         Width           =   11625
         _Version        =   65536
         _ExtentX        =   20505
         _ExtentY        =   3413
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
         Begin VB.ComboBox cmb_MotPpg 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1470
            TabIndex        =   46
            Text            =   "MOTIVO DEL PREPAGO"
            Top             =   420
            Width           =   10080
         End
         Begin VB.TextBox txt_ObsPpg 
            Enabled         =   0   'False
            Height          =   1065
            Left            =   1470
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   2
            Top             =   750
            Width           =   10065
         End
         Begin VB.Label Label31 
            Caption         =   "Motivo Prepago"
            Height          =   315
            Left            =   120
            TabIndex        =   48
            Top             =   450
            Width           =   1290
         End
         Begin VB.Label Label9 
            Caption         =   "Comentarios"
            Height          =   315
            Left            =   120
            TabIndex        =   47
            Top             =   780
            Width           =   1290
         End
         Begin VB.Label Label8 
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
            TabIndex        =   3
            Top             =   90
            Width           =   2085
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   2985
         Left            =   30
         TabIndex        =   4
         Top             =   4770
         Width           =   11625
         _Version        =   65536
         _ExtentX        =   20505
         _ExtentY        =   5265
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
         Begin EditLib.fpDateTime ipp_FecPre 
            Height          =   315
            Left            =   1470
            TabIndex        =   5
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
         Begin Threed.SSPanel pnl_SaldoTNC1 
            Height          =   315
            Left            =   1470
            TabIndex        =   6
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
            Left            =   1470
            TabIndex        =   7
            Top             =   1050
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
         Begin Threed.SSPanel pnl_UltPagTNC 
            Height          =   315
            Left            =   4350
            TabIndex        =   8
            Top             =   720
            Width           =   1440
            _Version        =   65536
            _ExtentX        =   2540
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
            Left            =   4350
            TabIndex        =   9
            Top             =   1050
            Width           =   1440
            _Version        =   65536
            _ExtentX        =   2540
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
            Left            =   7140
            TabIndex        =   10
            Top             =   720
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
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
         Begin Threed.SSPanel pnl_DiasTC 
            Height          =   315
            Left            =   7140
            TabIndex        =   11
            Top             =   1050
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
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
            Left            =   1470
            TabIndex        =   12
            Top             =   2580
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
            Left            =   4350
            TabIndex        =   13
            Top             =   390
            Width           =   1440
            _Version        =   65536
            _ExtentX        =   2540
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
         Begin Threed.SSPanel pnl_MtoPrepagar 
            Height          =   315
            Left            =   7140
            TabIndex        =   14
            Top             =   2580
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
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
            Left            =   10230
            TabIndex        =   39
            Top             =   1050
            Width           =   1260
            _Version        =   65536
            _ExtentX        =   2222
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
         Begin Threed.SSPanel pnl_InteresTNC 
            Height          =   315
            Left            =   10230
            TabIndex        =   40
            Top             =   720
            Width           =   1260
            _Version        =   65536
            _ExtentX        =   2222
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
         Begin Threed.SSPanel pnl_MontoPortes 
            Height          =   315
            Left            =   7140
            TabIndex        =   41
            Top             =   1380
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
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
            Left            =   4350
            TabIndex        =   42
            Top             =   1380
            Width           =   1440
            _Version        =   65536
            _ExtentX        =   2540
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
            Left            =   1470
            TabIndex        =   43
            Top             =   1380
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
         Begin Threed.SSPanel pnl_MontoITF 
            Height          =   315
            Left            =   4350
            TabIndex        =   44
            Top             =   2580
            Width           =   1440
            _Version        =   65536
            _ExtentX        =   2540
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
         Begin Threed.SSPanel pnl_CapPbpPerdido 
            Height          =   315
            Left            =   1470
            TabIndex        =   49
            Top             =   2250
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
         Begin Threed.SSPanel pnl_IntPbpPerdido 
            Height          =   315
            Left            =   4350
            TabIndex        =   50
            Top             =   2250
            Width           =   1440
            _Version        =   65536
            _ExtentX        =   2540
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
         Begin Threed.SSPanel pnl_IntLeg 
            Height          =   315
            Left            =   4350
            TabIndex        =   53
            Top             =   1710
            Width           =   1440
            _Version        =   65536
            _ExtentX        =   2540
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
         Begin Threed.SSPanel pnl_DevPbp 
            Height          =   315
            Left            =   1470
            TabIndex        =   54
            Top             =   1710
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
         Begin VB.Label Label14 
            Caption         =   "Devolucion PBP"
            Height          =   315
            Left            =   120
            TabIndex        =   56
            Top             =   1770
            Width           =   1335
         End
         Begin VB.Label Label13 
            Caption         =   "Interés Legal"
            Height          =   315
            Left            =   2970
            TabIndex        =   55
            Top             =   1770
            Width           =   1335
         End
         Begin VB.Label Label10 
            Caption         =   "Capital PBP"
            Height          =   315
            Left            =   120
            TabIndex        =   52
            Top             =   2310
            Width           =   1185
         End
         Begin VB.Label Label12 
            Caption         =   "Interés PBP"
            Height          =   315
            Left            =   2970
            TabIndex        =   51
            Top             =   2310
            Width           =   1335
         End
         Begin VB.Label Label6 
            Caption         =   "Interés TNC a la fecha"
            Height          =   315
            Left            =   8520
            TabIndex        =   31
            Top             =   780
            Width           =   1695
         End
         Begin VB.Label Label7 
            Caption         =   "Interés TC a la fecha"
            Height          =   315
            Left            =   8520
            TabIndex        =   30
            Top             =   1110
            Width           =   1695
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
            Left            =   90
            TabIndex        =   29
            Top             =   90
            UseMnemonic     =   0   'False
            Width           =   1875
         End
         Begin VB.Label Label26 
            Caption         =   "Monto ITF"
            Height          =   315
            Left            =   2970
            TabIndex        =   28
            Top             =   2640
            Width           =   1335
         End
         Begin VB.Label Label25 
            Caption         =   "Dias TNC"
            Height          =   315
            Left            =   6000
            TabIndex        =   27
            Top             =   780
            Width           =   1020
         End
         Begin VB.Label Label24 
            Caption         =   "Dias TC"
            Height          =   315
            Left            =   6000
            TabIndex        =   26
            Top             =   1110
            Width           =   1020
         End
         Begin VB.Label Label23 
            Caption         =   "Ultimo Pago TNC"
            Height          =   315
            Left            =   2970
            TabIndex        =   25
            Top             =   780
            Width           =   1335
         End
         Begin VB.Label Label22 
            Caption         =   "Ultimo Pago TC"
            Height          =   315
            Left            =   2970
            TabIndex        =   24
            Top             =   1110
            Width           =   1335
         End
         Begin VB.Label Label21 
            Caption         =   "Saldo Actual TNC"
            Height          =   315
            Left            =   120
            TabIndex        =   23
            Top             =   780
            Width           =   1335
         End
         Begin VB.Label Label20 
            Caption         =   "Saldo Actual TC"
            Height          =   315
            Left            =   120
            TabIndex        =   22
            Top             =   1110
            Width           =   1335
         End
         Begin VB.Label Label19 
            Caption         =   "Seguro Inmueble"
            Height          =   315
            Left            =   2970
            TabIndex        =   21
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label Label5 
            Caption         =   "Valor Asegur. Inm."
            Height          =   315
            Left            =   2970
            TabIndex        =   20
            Top             =   450
            Width           =   1335
         End
         Begin VB.Label Label11 
            Caption         =   "Fecha Prepago"
            Height          =   315
            Left            =   120
            TabIndex        =   19
            Top             =   450
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "Total Antes ITF"
            Height          =   315
            Left            =   120
            TabIndex        =   18
            Top             =   2640
            Width           =   1335
         End
         Begin VB.Label Label17 
            Caption         =   "Seguro Desgrav."
            Height          =   315
            Left            =   120
            TabIndex        =   17
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Monto Portes"
            Height          =   315
            Left            =   6000
            TabIndex        =   16
            Top             =   1440
            Width           =   1020
         End
         Begin VB.Label Label4 
            Caption         =   "Total Prepagar"
            Height          =   315
            Left            =   6000
            TabIndex        =   15
            Top             =   2640
            Width           =   1335
         End
         Begin VB.Line Line1 
            X1              =   120
            X2              =   11470
            Y1              =   2130
            Y2              =   2130
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   3405
         Left            =   30
         TabIndex        =   32
         Top             =   1350
         Width           =   11625
         _Version        =   65536
         _ExtentX        =   20505
         _ExtentY        =   6006
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
            Height          =   2985
            Left            =   30
            TabIndex        =   33
            Top             =   360
            Width           =   11565
            _ExtentX        =   20399
            _ExtentY        =   5265
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
            TabIndex        =   34
            Top             =   90
            Width           =   1875
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   645
         Left            =   30
         TabIndex        =   35
         Top             =   690
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
            Picture         =   "OpeTra_frm_331.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   45
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   11010
            Picture         =   "OpeTra_frm_331.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   36
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   645
         Left            =   30
         TabIndex        =   37
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
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   10950
            Top             =   150
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
         Begin Threed.SSPanel SSPanel7 
            Height          =   555
            Left            =   690
            TabIndex        =   38
            Top             =   60
            Width           =   4575
            _Version        =   65536
            _ExtentX        =   8070
            _ExtentY        =   979
            _StockProps     =   15
            Caption         =   "Prepago Total de Crédito Hipotecario - Consulta"
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
            Picture         =   "OpeTra_frm_331.frx":0758
            Top             =   90
            Width           =   480
         End
      End
   End
End
Attribute VB_Name = "frm_Con_PrePgo_04"
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
    'exporta liquidacion
    If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    Me.Enabled = False
    If InStr(moddat_g_str_Agr1MIC, moddat_g_str_CodPrd) > 0 Or InStr(moddat_g_str_Agr1FMV, moddat_g_str_CodPrd) > 0 Then
        Call fs_Report_Micasita
    Else
        Call fs_Report_Mivivienda
    End If
    Me.Enabled = True
    Screen.MousePointer = 0
End Sub

Private Sub fs_Report_Micasita()
    Dim r_obj_Excel      As Excel.Application
    Dim r_int_NroFil     As Integer
    Dim r_int_ColumB     As Integer
    Dim r_int_ColumC     As Integer
    Dim r_int_ColumD     As Integer
    Dim r_int_ColumE     As Integer
    Dim r_int_ColumF     As Integer
    Dim r_int_ColumH     As Integer
    Dim r_int_ColumI     As Integer
    Dim r_int_ColumJ     As Integer
    Dim r_int_ColumK     As Integer
    Dim r_int_ColumL     As Integer
    Dim r_int_ColumM     As Integer
    Dim r_int_ColumO     As Integer
    Dim r_int_ColumP     As Integer
    
    r_int_NroFil = 3
    r_int_ColumB = 2
    r_int_ColumC = 3
    r_int_ColumD = 4
    r_int_ColumE = 5
    r_int_ColumF = 6
    r_int_ColumH = 8
    r_int_ColumI = 9
    r_int_ColumJ = 10
    r_int_ColumK = 11
    r_int_ColumL = 12
    r_int_ColumM = 13
    r_int_ColumO = 15
    r_int_ColumP = 16

    ''********************************************
    Set r_obj_Excel = New Excel.Application
    r_obj_Excel.SheetsInNewWorkbook = 1
    r_obj_Excel.Workbooks.Add
      
    With r_obj_Excel.ActiveSheet
       'CENTRADO DE LA PAGINA
       '.PageSetup.CenterHorizontally = True
       .PageSetup.Orientation = xlLandscape
    
       'WIDTH
       .Columns("B").ColumnWidth = 50
       .Columns("C").ColumnWidth = 15
      
       'MARGENES
       .PageSetup.LeftMargin = Application.CentimetersToPoints(1)
       .PageSetup.RightMargin = Application.CentimetersToPoints(1)
       .PageSetup.TopMargin = Application.CentimetersToPoints(1)
       .PageSetup.BottomMargin = Application.CentimetersToPoints(1)
    
       .Range("A1:Y1").ColumnWidth = 6
       .Range("A8:A32").RowHeight = 12
    
       'BORDERS
       .Range("B3:R3").Borders(xlEdgeTop).Weight = xlMedium
       .Range("B4:R4").Borders(xlEdgeTop).Weight = xlMedium
       .Range("B5:R5").Borders(xlEdgeTop).Weight = xlMedium
       .Range("B6:R6").Borders(xlEdgeTop).Weight = xlMedium
       .Range("B32:R32").Borders(xlEdgeBottom).Weight = xlMedium
       .Range("B3:B32").Borders(xlEdgeLeft).Weight = xlMedium
       .Range("R3:R32").Borders(xlEdgeRight).Weight = xlMedium
    
       'Font
       .Range("A1:T45").Font.Name = "Arial"
       .Range("A1:T34").Font.Size = 9
    
       'Fila 3 - Titulo
       .Range("B3:R3").Merge
       .Range("B3") = "LIQUIDACION PREPAGO TOTAL - " & moddat_g_str_NomPrd
       .Range("B3:R3").HorizontalAlignment = xlHAlignCenter
       .Range("B3:R3").Font.Size = 12
       .Range("B5:R5").Font.Size = 12
       .Range("B7:R7").Font.Size = 12
       .Range("B3:R3").Font.Bold = True
       .Range("B5:R5").Font.Bold = True
       .Range("B7:R7").Font.Bold = True
       r_int_NroFil = r_int_NroFil + 2
       
       'Fila 5 - Datos Cliente
       .Cells(r_int_NroFil, r_int_ColumC) = "Cliente:"
       .Cells(r_int_NroFil, r_int_ColumC).HorizontalAlignment = xlHAlignLeft
       .Cells(r_int_NroFil, r_int_ColumE) = moddat_g_str_NomCli
       .Cells(r_int_NroFil, r_int_ColumO) = "DNI:"
       .Cells(r_int_NroFil, r_int_ColumO).HorizontalAlignment = xlHAlignLeft
       .Cells(r_int_NroFil, r_int_ColumP) = "'" & moddat_g_str_NumDoc
       .Cells(r_int_NroFil, r_int_ColumP).HorizontalAlignment = xlHAlignLeft
       r_int_NroFil = r_int_NroFil + 2
       
       'Fila 7 - Numero de Operacion, Moneda
       .Cells(r_int_NroFil, r_int_ColumC) = "N° de Operación:"
       .Cells(r_int_NroFil, r_int_ColumF) = Mid(moddat_g_str_NumOpe, 1, 3) & "-" & Mid(moddat_g_str_NumOpe, 4, 2) & "-" & Mid(moddat_g_str_NumOpe, 6, 5)
       .Range("C" & r_int_NroFil & ":N" & r_int_NroFil & "").Font.Bold = True
       .Range("C" & r_int_NroFil & ":N" & r_int_NroFil & "").Font.Size = 11
       .Cells(r_int_NroFil, r_int_ColumK) = "Moneda:"
       .Cells(r_int_NroFil, r_int_ColumM) = moddat_g_str_Moneda
       r_int_NroFil = r_int_NroFil + 2
       
       'Fila 9 - Ultimo Pago, Saldo
       .Cells(r_int_NroFil, r_int_ColumC) = "Saldo al"
       .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
       .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Font.Bold = True
       .Cells(r_int_NroFil, r_int_ColumI).HorizontalAlignment = xlHAlignRight
       .Cells(r_int_NroFil, r_int_ColumI) = "'" & Format(CDate(pnl_UltPagTNC.Caption), "dd-mm-yy")
       
       .Cells(r_int_NroFil, r_int_ColumK) = Format(moddat_g_dbl_SalCap, "###,###.00")
       .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Merge
       .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").HorizontalAlignment = xlHAlignRight
       
       'MARCO
       .Cells(r_int_NroFil, r_int_ColumK).Borders(xlEdgeLeft).Weight = xlThin
       .Cells(r_int_NroFil, r_int_ColumM).Borders(xlEdgeRight).Weight = xlThin
       .Range("K" & r_int_NroFil & ":M" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
       .Range("K" & r_int_NroFil & ":M" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
       r_int_NroFil = r_int_NroFil + 1
       
       'Fila 10 - Fecha de Prepago
       .Cells(r_int_NroFil, r_int_ColumC) = "Fecha de corte"
       .Range("C" & r_int_NroFil & ":M" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
       .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
       .Cells(r_int_NroFil, r_int_ColumI) = "'" & Format(CDate(ipp_FecPre), "dd-mm-yy")
       .Cells(r_int_NroFil, r_int_ColumI).HorizontalAlignment = xlHAlignRight
       r_int_NroFil = r_int_NroFil + 2
       
       'Fila 12 - Dias Interes
       .Cells(r_int_NroFil, r_int_ColumC) = "Días de interés"
       .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
       .Cells(r_int_NroFil, r_int_ColumI) = CInt(pnl_DiasTNC.Caption)
       r_int_NroFil = r_int_NroFil + 2
       
       'Fila 14 - Valor inmueble
       .Cells(r_int_NroFil, r_int_ColumC) = "Valor asegurable inmueble"
       .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
       .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Font.Bold = True
       .Cells(r_int_NroFil, r_int_ColumI) = Format(CDbl(pnl_Val_AsgInm.Caption), "###,###.00")
       r_int_NroFil = r_int_NroFil + 2
       
       'Fila 16 - Tasa interes anual
       .Cells(r_int_NroFil, r_int_ColumC) = "Tasa de interés anual (%)"
       .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
       .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").HorizontalAlignment = xlHAlignRight
       .Cells(r_int_NroFil, r_int_ColumI) = l_dbl_TasInt
       .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Select
       r_obj_Excel.Selection.NumberFormat = "###,##0.0000"
       r_int_NroFil = r_int_NroFil + 1
       
       'Fila 17 - Tasa seguro inmueble
       .Cells(r_int_NroFil, r_int_ColumC) = "Tasa de Seguro Inmueble (%)"
       .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
       .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").HorizontalAlignment = xlHAlignRight
       .Cells(r_int_NroFil, r_int_ColumI) = l_dbl_SegInm
       .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Select
       r_obj_Excel.Selection.NumberFormat = "###,##0.0000"
       r_int_NroFil = r_int_NroFil + 1
       
       'Fila 18 - Tasa seguro desgravamen
       .Cells(r_int_NroFil, r_int_ColumC) = "Tasa de Seguro Desgravamen (%)"
       .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
       .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").HorizontalAlignment = xlHAlignRight
       .Cells(r_int_NroFil, r_int_ColumI) = l_dbl_SegDes
       .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Select
       r_obj_Excel.Selection.NumberFormat = "###,##0.0000"
       r_int_NroFil = r_int_NroFil + 2
       
       'Fila 20 - Subtitulo gastos
       .Cells(r_int_NroFil, r_int_ColumC) = "Intereses y Seguros"
       .Range("C" & r_int_NroFil & ":M" & r_int_NroFil & "").Merge
        
        'MARCO
       .Cells(r_int_NroFil, r_int_ColumC).Borders(xlEdgeLeft).Weight = xlThin
       .Cells(r_int_NroFil, r_int_ColumM).Borders(xlEdgeRight).Weight = xlThin
       .Range("C" & r_int_NroFil & ":M" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
       .Range("C" & r_int_NroFil & ":M" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
       .Range("C" & r_int_NroFil & ":M" & r_int_NroFil & "").HorizontalAlignment = xlHAlignCenter
       r_int_NroFil = r_int_NroFil + 1
       
       'Fila 21 - Portes
       .Cells(r_int_NroFil, r_int_ColumC) = "Portes"
       .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
       .Cells(r_int_NroFil, r_int_ColumI) = CDbl(pnl_MontoPortes.Caption)
       .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Select
       r_obj_Excel.Selection.NumberFormat = "###,##0.00"
       r_int_NroFil = r_int_NroFil + 1
       
       'Fila 22 - Intereses
       .Cells(r_int_NroFil, r_int_ColumC) = "Intereses a la fecha"
       .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
       .Cells(r_int_NroFil, r_int_ColumI) = CDbl(pnl_InteresTNC.Caption)
       .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Select
       r_obj_Excel.Selection.NumberFormat = "###,##0.00"
       r_int_NroFil = r_int_NroFil + 1
       
       'Fila 23 - Seguro desgravamen
       .Cells(r_int_NroFil, r_int_ColumC) = "Seguro Desgravamen"
       .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
       .Cells(r_int_NroFil, r_int_ColumI) = CDbl(pnl_SegDes.Caption)
       .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Select
       r_obj_Excel.Selection.NumberFormat = "###,##0.00"
       r_int_NroFil = r_int_NroFil + 1
       
       'Fila 24 - Seguro inmueble
       .Cells(r_int_NroFil, r_int_ColumC) = "Seguro Inmueble"
       .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
       .Cells(r_int_NroFil, r_int_ColumI) = CDbl(pnl_SegInm.Caption)
       .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Select
       r_obj_Excel.Selection.NumberFormat = "###,##0.00"
       r_int_NroFil = r_int_NroFil + 1
       
       'Fila 25 - Devolucion BBP
       .Cells(r_int_NroFil, r_int_ColumC) = "Devolución Bono"
       .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
       .Cells(r_int_NroFil, r_int_ColumI) = CDbl(pnl_DevPbp.Caption)
       .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Select
       r_obj_Excel.Selection.NumberFormat = "###,##0.00"
       r_int_NroFil = r_int_NroFil + 1
       
       'Fila 26 - Interes Legal
       .Cells(r_int_NroFil, r_int_ColumC) = "Interes Legal"
       .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
       .Cells(r_int_NroFil, r_int_ColumI) = CDbl(pnl_IntLeg.Caption)
       .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Select
       r_obj_Excel.Selection.NumberFormat = "###,##0.00"
       r_int_NroFil = r_int_NroFil + 1
       
       'Fila 27 - Total gastos
       .Cells(r_int_NroFil, r_int_ColumC) = "Total a aplicar al saldo"
       .Range("C" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
       .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Merge
       .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Font.Bold = True
       .Cells(r_int_NroFil, r_int_ColumK) = CDbl(CDbl(pnl_InteresTNC.Caption) + CDbl(pnl_MontoPortes.Caption) + CDbl(pnl_SegDes.Caption) + CDbl(pnl_SegInm.Caption) + CDbl(pnl_DevPbp.Caption) + CDbl(pnl_IntLeg.Caption))
       .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Select
       r_obj_Excel.Selection.NumberFormat = "###,##0.00"
       .Range("C" & r_int_NroFil & ":J" & r_int_NroFil & "").HorizontalAlignment = xlHAlignCenter
       .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").HorizontalAlignment = xlHAlignRight
       
       'MARCO
       .Cells(r_int_NroFil, r_int_ColumC).Borders(xlEdgeLeft).Weight = xlThin
       .Cells(r_int_NroFil, r_int_ColumK).Borders(xlEdgeLeft).Weight = xlThin
       .Cells(r_int_NroFil, r_int_ColumM).Borders(xlEdgeRight).Weight = xlThin
       .Range("C" & r_int_NroFil & ":M" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
       .Range("C" & r_int_NroFil & ":M" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
       r_int_NroFil = r_int_NroFil + 2
       
       'Fila 29 - Total antes del ITF
       .Cells(r_int_NroFil, r_int_ColumC) = "Monto a prepagar antes de ITF"
       .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Merge
       .Range("C" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
       .Range("K" & r_int_NroFil & ":M" & r_int_NroFil & "").Font.Bold = True
       .Range("C" & r_int_NroFil & ":J" & r_int_NroFil & "").HorizontalAlignment = xlHAlignCenter
       .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").HorizontalAlignment = xlHAlignRight
       .Cells(r_int_NroFil, r_int_ColumK) = Format(CDbl(pnl_MtoApl.Caption), "###,###.00")
       
       'MARCO
       .Cells(r_int_NroFil, r_int_ColumC).Borders(xlEdgeLeft).Weight = xlThin
       .Cells(r_int_NroFil, r_int_ColumK).Borders(xlEdgeLeft).Weight = xlThin
       .Cells(r_int_NroFil, r_int_ColumM).Borders(xlEdgeRight).Weight = xlThin
       .Range("C" & r_int_NroFil & ":M" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
       .Range("C" & r_int_NroFil & ":M" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
       r_int_NroFil = r_int_NroFil + 2
       
       'Fila 31 - Monto ITF
       .Cells(r_int_NroFil, r_int_ColumC) = "ITF (%)"
       .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Merge
       .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
       .Cells(r_int_NroFil, r_int_ColumI) = "'0.0005"
       .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").HorizontalAlignment = xlHAlignRight
       .Cells(r_int_NroFil, r_int_ColumK) = Format$(CDbl(pnl_MontoITF.Caption), "##.00")
       .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").HorizontalAlignment = xlHAlignRight
       .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Font.Bold = True
       r_int_NroFil = r_int_NroFil + 2
       
       'Fila 33 - Total Prepago
       .Cells(r_int_NroFil, r_int_ColumC) = "Monto total del prepago"
       .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Merge
       .Range("C" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
       .Range("C" & r_int_NroFil & ":J" & r_int_NroFil & "").Font.Bold = True
       .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Font.Bold = True
       .Range("C" & r_int_NroFil & ":J" & r_int_NroFil & "").HorizontalAlignment = xlHAlignCenter
       .Cells(r_int_NroFil, r_int_ColumK) = Format$(CDbl(pnl_MtoPrepagar.Caption), "###,###.00")
       .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Select
       
       If CInt(moddat_g_int_TipMon) = 1 Then
          r_obj_Excel.Selection.NumberFormat = "$###,##0.00_);[Red]($###,##0.00)"
          .Cells(34, 2) = "Realizar el depósito en la cuenta Nº 0011-0369-02-00090532 del BBVA Banco Continental."
          .Cells(35, 2) = "Realizar el depósito en la cuenta Nº 011-369-000200090532-69 del BBVA Banco Continental desde otro Banco."
          .Range(.Cells(34, 2), .Cells(35, 2)).Font.Size = 12
          .Range(.Cells(34, 2), .Cells(35, 2)).Font.Bold = True
       Else
          r_obj_Excel.Selection.NumberFormat = "[$$]#,##0.00;[Red][$$]#,##0.00"
          .Cells(34, 2) = "Realizar el depósito en la cuenta Nº 0011-0369-02-00090540 del BBVA Banco Continental."
          .Cells(34, 2).Font.Size = 12
          .Cells(34, 2).Font.Bold = True
       End If
       
       .Cells(r_int_NroFil, r_int_ColumC).Borders(xlEdgeLeft).Weight = xlThin
       .Cells(r_int_NroFil, r_int_ColumK).Borders(xlEdgeLeft).Weight = xlThin
       .Cells(r_int_NroFil, r_int_ColumM).Borders(xlEdgeRight).Weight = xlThin
       .Range("C" & r_int_NroFil & ":M" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
       .Range("C" & r_int_NroFil & ":M" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
       r_int_NroFil = r_int_NroFil + 2
    End With
   
    r_obj_Excel.Visible = True
    Set r_obj_Excel = Nothing
 
End Sub

Private Sub fs_Report_Mivivienda()
    Dim r_obj_Excel      As Excel.Application
    Dim r_int_NroFil     As Integer
    Dim r_int_ColumC     As Integer
    Dim r_int_ColumD     As Integer
    Dim r_int_ColumE     As Integer
    Dim r_int_ColumH     As Integer
    Dim r_int_ColumI     As Integer
    Dim r_int_ColumJ     As Integer
    Dim r_int_ColumK     As Integer
    Dim r_int_ColumL     As Integer
    Dim r_int_ColumM     As Integer
    Dim r_int_ColumN     As Integer
    Dim r_int_ColumO     As Integer
    Dim r_int_ColumP     As Integer
    Dim r_int_ColumQ     As Integer
    
    r_int_NroFil = 3
    r_int_ColumC = 3
    r_int_ColumD = 4
    r_int_ColumE = 5
    r_int_ColumH = 8
    r_int_ColumI = 9
    r_int_ColumJ = 10
    r_int_ColumK = 11
    r_int_ColumL = 12
    r_int_ColumM = 13
    r_int_ColumN = 14
    r_int_ColumO = 15
    r_int_ColumP = 16
    r_int_ColumQ = 17
    
    ''********************************************
    Set r_obj_Excel = New Excel.Application
    r_obj_Excel.SheetsInNewWorkbook = 1
    r_obj_Excel.Workbooks.Add
    
    With r_obj_Excel.ActiveSheet
       'CENTRADO DE LA PAGINA
       '.PageSetup.CenterHorizontally = True
       .PageSetup.Orientation = xlLandscape
       
       'WIDTH
       .Columns("B").ColumnWidth = 50
       .Columns("C").ColumnWidth = 15
             
       'MARGENES
       .PageSetup.LeftMargin = Application.CentimetersToPoints(1)
       .PageSetup.RightMargin = Application.CentimetersToPoints(1)
       .PageSetup.TopMargin = Application.CentimetersToPoints(1)
       .PageSetup.BottomMargin = Application.CentimetersToPoints(1)
       
       .Range("A1:A1").ColumnWidth = 2
       .Range("B1:Y1").ColumnWidth = 6
       .Range("A8:A33").RowHeight = 12.75
       
       'BORDERS
       .Range("B3:R3").Borders(xlEdgeTop).Weight = xlMedium
       .Range("B4:R4").Borders(xlEdgeTop).Weight = xlMedium
       .Range("B5:R5").Borders(xlEdgeTop).Weight = xlMedium
       .Range("B6:R6").Borders(xlEdgeTop).Weight = xlMedium
       .Range("B31:R32").Borders(xlEdgeBottom).Weight = xlMedium
       .Range("B3:B32").Borders(xlEdgeLeft).Weight = xlMedium
       .Range("R3:R32").Borders(xlEdgeRight).Weight = xlMedium
       
       'Font
       .Range("A1:T43").Font.Name = "Arial"
       .Range("A1:T33").Font.Size = 9
       
       'Fila 3 - Titulo
       .Range("B3:R3").Merge
       .Range("B3") = "LIQUIDACION PREPAGO TOTAL - " & moddat_g_str_NomPrd '& " " & moddat_g_str_Moneda
       .Range("B3:R3").HorizontalAlignment = xlHAlignCenter
       .Range("B3:R3").Font.Size = 12
       .Range("B5:R5").Font.Size = 12
       .Range("B7:R7").Font.Size = 12
       .Range("B3:R3").Font.Bold = True
       .Range("B5:R5").Font.Bold = True
       .Range("B7:R7").Font.Bold = True
       r_int_NroFil = r_int_NroFil + 2
       
       'Fila 5 - Datos Cliente
       .Cells(r_int_NroFil, r_int_ColumC) = "Cliente:"
       .Cells(r_int_NroFil, r_int_ColumC).HorizontalAlignment = xlHAlignLeft
       .Cells(r_int_NroFil, r_int_ColumE) = moddat_g_str_NomCli
       .Cells(r_int_NroFil, r_int_ColumO) = "DNI:"
       .Cells(r_int_NroFil, r_int_ColumO).HorizontalAlignment = xlHAlignLeft
       .Cells(r_int_NroFil, r_int_ColumP) = "'" & moddat_g_str_NumDoc
       .Cells(r_int_NroFil, r_int_ColumP).HorizontalAlignment = xlHAlignLeft
       r_int_NroFil = r_int_NroFil + 2
       
       'Fila 7 - Numero Operacion
       .Cells(r_int_NroFil, r_int_ColumC) = "Operación:"
       .Cells(r_int_NroFil, r_int_ColumE) = Mid(moddat_g_str_NumOpe, 1, 3) & "-" & Mid(moddat_g_str_NumOpe, 4, 2) & "-" & Mid(moddat_g_str_NumOpe, 6, 5)
       .Cells(r_int_NroFil, r_int_ColumK) = "Moneda:"
       .Cells(r_int_NroFil, r_int_ColumM) = moddat_g_str_Moneda
       r_int_NroFil = r_int_NroFil + 2
       
       'Fila 9 - Ultimo cuota TC, subtitulo saldo
       .Cells(r_int_NroFil, r_int_ColumC) = "Fecha de Desembolso o última cuota TC (A)"
       .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
       .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Font.Bold = True
       .Cells(r_int_NroFil, r_int_ColumI) = "'" & Format(CDate(pnl_UltPagTC.Caption), "dd-mm-yy")
       .Cells(r_int_NroFil, r_int_ColumI).HorizontalAlignment = xlHAlignRight
       .Range("N" & r_int_NroFil & ":Q" & r_int_NroFil & "").Merge
       .Range("N" & r_int_NroFil & ":Q" & r_int_NroFil & "").Font.Bold = True
       .Cells(r_int_NroFil, r_int_ColumN) = "Saldo antes del prepago"
       
       'MARCO
       .Cells(r_int_NroFil, r_int_ColumN).Borders(xlEdgeLeft).Weight = xlThin
       .Cells(r_int_NroFil, r_int_ColumQ).Borders(xlEdgeRight).Weight = xlThin
       .Range("N" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
       .Range("N" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
       .Range("N" & r_int_NroFil & ":M" & r_int_NroFil & "").HorizontalAlignment = xlHAlignCenter
       r_int_NroFil = r_int_NroFil + 1
       
       'Fila 11 - Ultimo Pago TNC, Saldo Deuda, Saldo TNC
       .Cells(r_int_NroFil, r_int_ColumC) = "Saldo al (B)"
       .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
       .Cells(r_int_NroFil, r_int_ColumI).HorizontalAlignment = xlHAlignRight
       .Cells(r_int_NroFil, r_int_ColumI) = "'" & Format(CDate(pnl_UltPagTNC.Caption), "dd-mm-yy")
       
       .Cells(r_int_NroFil, r_int_ColumL) = Format(CDbl(pnl_SaldoTNC1.Caption) + CDbl(pnl_SaldoTC1.Caption), "###,###.00")
       .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Merge
       .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").HorizontalAlignment = xlHAlignRight
       
       .Range("N" & r_int_NroFil & ":O" & r_int_NroFil & "").Merge
       .Range("P" & r_int_NroFil & ":Q" & r_int_NroFil & "").Merge
       .Cells(r_int_NroFil, r_int_ColumN) = "TNC"
       .Cells(r_int_NroFil, r_int_ColumP) = Format(pnl_SaldoTNC1.Caption, "###,###.00")
       .Range("N" & r_int_NroFil & ":O" & r_int_NroFil & "").HorizontalAlignment = xlHAlignCenter
       .Cells(r_int_NroFil, r_int_ColumK).Font.Bold = True
       
       'MARCO
       .Cells(r_int_NroFil, r_int_ColumK).Borders(xlEdgeLeft).Weight = xlThin
       .Cells(r_int_NroFil, r_int_ColumL).Borders(xlEdgeRight).Weight = xlThin
       .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
       .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
       .Cells(r_int_NroFil, r_int_ColumN).Borders(xlEdgeLeft).Weight = xlThin
       .Cells(r_int_NroFil, r_int_ColumO).Borders(xlEdgeRight).Weight = xlThin
       .Cells(r_int_NroFil, r_int_ColumQ).Borders(xlEdgeRight).Weight = xlThin
       .Range("N" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
       .Range("N" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
       r_int_NroFil = r_int_NroFil + 1
       
       'Fila 12 - Fecha Prepago, Saldo TC
       .Cells(r_int_NroFil, r_int_ColumC) = "Fecha de corte (fecha del prepago) (C)"
       .Range("C" & r_int_NroFil & ":L" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
       .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
       .Cells(r_int_NroFil, r_int_ColumI) = "'" & Format(CDate(ipp_FecPre), "dd-mm-yy")
       .Cells(r_int_NroFil, r_int_ColumI).HorizontalAlignment = xlHAlignRight
       
       .Range("N" & r_int_NroFil & ":O" & r_int_NroFil & "").Merge
       .Range("P" & r_int_NroFil & ":Q" & r_int_NroFil & "").Merge
       .Cells(r_int_NroFil, r_int_ColumN) = "TC"
       .Cells(r_int_NroFil, r_int_ColumP) = Format(pnl_SaldoTC1.Caption, "###,###.00")
       .Range("N" & r_int_NroFil & ":O" & r_int_NroFil & "").HorizontalAlignment = xlHAlignCenter
       
       'MARCO
       .Cells(r_int_NroFil, r_int_ColumN).Borders(xlEdgeLeft).Weight = xlThin
       .Cells(r_int_NroFil, r_int_ColumO).Borders(xlEdgeRight).Weight = xlThin
       .Cells(r_int_NroFil, r_int_ColumQ).Borders(xlEdgeRight).Weight = xlThin
       .Range("N" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
       r_int_NroFil = r_int_NroFil + 1
             
       .Range("N" & r_int_NroFil & ":O" & r_int_NroFil & "").Merge
       .Range("P" & r_int_NroFil & ":Q" & r_int_NroFil & "").Merge
       .Range("P" & r_int_NroFil & ":Q" & r_int_NroFil & "").Font.Bold = True
       .Cells(r_int_NroFil, r_int_ColumN) = "Total"
       .Cells(r_int_NroFil, r_int_ColumP) = Format(CDbl(pnl_SaldoTNC1.Caption) + CDbl(pnl_SaldoTC1.Caption), "###,###.00")
       .Range("N" & r_int_NroFil & ":O" & r_int_NroFil & "").HorizontalAlignment = xlHAlignCenter
       
       .Cells(r_int_NroFil, r_int_ColumN).Borders(xlEdgeLeft).Weight = xlThin
       .Cells(r_int_NroFil, r_int_ColumO).Borders(xlEdgeRight).Weight = xlThin
       .Cells(r_int_NroFil, r_int_ColumQ).Borders(xlEdgeRight).Weight = xlThin
       .Range("N" & r_int_NroFil & ":Q" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
       r_int_NroFil = r_int_NroFil + 1
       
       'Fila 14 - Dias TNC
       .Cells(r_int_NroFil, r_int_ColumC) = "Días de interés TNC (C-B)"
       .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
       .Cells(r_int_NroFil, r_int_ColumI) = CInt(pnl_DiasTNC.Caption)
       r_int_NroFil = r_int_NroFil + 1
       
       'Fila 15 - Dias TC
       .Cells(r_int_NroFil, r_int_ColumC) = "Días de interés TC (C-A)"
       .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
       .Cells(r_int_NroFil, r_int_ColumI) = CInt(pnl_DiasTC.Caption)
       r_int_NroFil = r_int_NroFil + 2
       
       'Fila 17 - Valor inmueble
       .Cells(r_int_NroFil, r_int_ColumC) = "Valor asegurable inmueble"
       .Cells(r_int_NroFil, r_int_ColumI) = Format(CDbl(pnl_Val_AsgInm.Caption), "###,###.00")
       .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
       .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Font.Bold = True
       r_int_NroFil = r_int_NroFil + 2
       
       'Fila 19,20,21 - Tasas
       .Cells(r_int_NroFil, r_int_ColumC) = "Tasa de interés anual (%)"
       .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
       .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").HorizontalAlignment = xlHAlignRight
       .Cells(r_int_NroFil, r_int_ColumI) = l_dbl_TasInt
       .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Select
       r_obj_Excel.Selection.NumberFormat = "###0.0000"
       r_int_NroFil = r_int_NroFil + 1
       
       .Cells(r_int_NroFil, r_int_ColumC) = "Tasa de Seguro Inmueble (%)"
       .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
       .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").HorizontalAlignment = xlHAlignRight
       .Cells(r_int_NroFil, r_int_ColumI) = l_dbl_SegInm
       .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Select
       r_obj_Excel.Selection.NumberFormat = "###0.0000"
       r_int_NroFil = r_int_NroFil + 1
       
       .Cells(r_int_NroFil, r_int_ColumC) = "Tasa de Seguro Desgravamen (%)"
       .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
       .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").HorizontalAlignment = xlHAlignRight
       .Cells(r_int_NroFil, r_int_ColumI) = l_dbl_SegDes
       .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Select
       r_obj_Excel.Selection.NumberFormat = "###0.0000"
       r_int_NroFil = r_int_NroFil + 2
       
       'Fila 23 - Subtitulo gastos
       .Cells(r_int_NroFil, r_int_ColumC) = "Intereses y Seguros"
       .Range("C" & r_int_NroFil & ":L" & r_int_NroFil & "").Merge
       
       'MARCO
       .Cells(r_int_NroFil, r_int_ColumC).Borders(xlEdgeLeft).Weight = xlThin
       .Cells(r_int_NroFil, r_int_ColumL).Borders(xlEdgeRight).Weight = xlThin
       .Range("C" & r_int_NroFil & ":L" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
       .Range("C" & r_int_NroFil & ":L" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
       .Range("C" & r_int_NroFil & ":L" & r_int_NroFil & "").HorizontalAlignment = xlHAlignCenter
       r_int_NroFil = r_int_NroFil + 1
       
       'Fila 24 - Portes
       .Cells(r_int_NroFil, r_int_ColumC) = "Portes"
       .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
       .Cells(r_int_NroFil, r_int_ColumI) = CDbl(pnl_MontoPortes.Caption)
       .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Select
       r_obj_Excel.Selection.NumberFormat = "###0.00"
       r_int_NroFil = r_int_NroFil + 1
       
       'Fila 25 - Interes TNC
       .Cells(r_int_NroFil, r_int_ColumC) = "Intereses TNC a la fecha"
       .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
       .Cells(r_int_NroFil, r_int_ColumI) = CDbl(pnl_InteresTNC.Caption)
       .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Select
       r_obj_Excel.Selection.NumberFormat = "###0.00"
       r_int_NroFil = r_int_NroFil + 1
       
       'Fila 26 - Intereses TC
       .Cells(r_int_NroFil, r_int_ColumC) = "Intereses TC a la fecha"
       .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
       .Cells(r_int_NroFil, r_int_ColumI) = CDbl(pnl_InteresTC.Caption)
       .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Select
       r_obj_Excel.Selection.NumberFormat = "###0.00"
       r_int_NroFil = r_int_NroFil + 1
       
       'Fila 27 - Seg. Desgravamen
       .Cells(r_int_NroFil, r_int_ColumC) = "Seguro Desgravamen"
       .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
       .Cells(r_int_NroFil, r_int_ColumI) = CDbl(pnl_SegDes.Caption)
       .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Select
       r_obj_Excel.Selection.NumberFormat = "###0.00"
       r_int_NroFil = r_int_NroFil + 1
       
       'Fila 28 - Seg. Inmueble
       .Cells(r_int_NroFil, r_int_ColumC) = "Seguro Inmueble"
       .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
       .Cells(r_int_NroFil, r_int_ColumI) = CDbl(pnl_SegInm.Caption)
       .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Select
       r_obj_Excel.Selection.NumberFormat = "###0.00"
       r_int_NroFil = r_int_NroFil + 1
       
       'Fila ## - PBP Perdido
       If CDbl(pnl_CapPbpPerdido.Caption) > 0 Then
         .Cells(r_int_NroFil, r_int_ColumC) = "Capital PBP Pendiente"
         .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
         .Cells(r_int_NroFil, r_int_ColumI) = CDbl(pnl_CapPbpPerdido.Caption)
         .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Select
         r_obj_Excel.Selection.NumberFormat = "#,##0.00"
         r_int_NroFil = r_int_NroFil + 1
       End If
       
       'Fila ## - PBP Perdido
       If CDbl(pnl_IntPbpPerdido.Caption) > 0 Then
         .Cells(r_int_NroFil, r_int_ColumC) = "Interes PBP Pendiente"
         .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
         .Cells(r_int_NroFil, r_int_ColumI) = CDbl(pnl_IntPbpPerdido.Caption)
         .Range("I" & r_int_NroFil & ":J" & r_int_NroFil & "").Select
         r_obj_Excel.Selection.NumberFormat = "#,##0.00"
         r_int_NroFil = r_int_NroFil + 1
       End If
       
       'Fila 29 - Total Gastos
       .Cells(r_int_NroFil, r_int_ColumC) = "Total de Interés y Seguros"
       .Range("C" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
       .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Merge
       .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Font.Bold = True
       .Cells(r_int_NroFil, r_int_ColumK) = CDbl(CDbl(pnl_InteresTNC.Caption) + CDbl(pnl_InteresTC.Caption) + CDbl(pnl_MontoPortes.Caption) + CDbl(pnl_SegDes.Caption) + CDbl(pnl_SegInm.Caption) + CDbl(pnl_CapPbpPerdido.Caption) + CDbl(pnl_IntPbpPerdido.Caption))
       .Range("C" & r_int_NroFil & ":J" & r_int_NroFil & "").HorizontalAlignment = xlHAlignCenter
       .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").HorizontalAlignment = xlHAlignRight
       
       'MARCO
       .Cells(r_int_NroFil, r_int_ColumC).Borders(xlEdgeLeft).Weight = xlThin
       .Cells(r_int_NroFil, r_int_ColumK).Borders(xlEdgeLeft).Weight = xlThin
       .Cells(r_int_NroFil, r_int_ColumL).Borders(xlEdgeRight).Weight = xlThin
       .Range("C" & r_int_NroFil & ":L" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
       .Range("C" & r_int_NroFil & ":L" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
       r_int_NroFil = r_int_NroFil + 2
       
       'Fila 31 - Total Prepago
       .Cells(r_int_NroFil, r_int_ColumC) = "Monto del Prepago Total"
       .Range("C" & r_int_NroFil & ":J" & r_int_NroFil & "").Merge
       .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Merge
       .Range("C" & r_int_NroFil & ":J" & r_int_NroFil & "").Font.Bold = True
       .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Font.Bold = True
       .Range("C" & r_int_NroFil & ":J" & r_int_NroFil & "").HorizontalAlignment = xlHAlignCenter
       .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").HorizontalAlignment = xlHAlignRight
       .Cells(r_int_NroFil, r_int_ColumK) = Format(CDbl(pnl_MtoApl.Caption), "###,###.00")
       .Range("K" & r_int_NroFil & ":L" & r_int_NroFil & "").Select
       
       If CInt(moddat_g_int_TipMon) = 1 Then
          r_obj_Excel.Selection.NumberFormat = "$###,##0.00_);[Red]($###,##0.00)"
          .Cells(34, 2) = "Realizar el depósito en la cuenta Nº 0011-0369-02-00090532 del BBVA Banco Continental."
          .Cells(35, 2) = "Realizar el depósito en la cuenta Nº 011-369-000200090532-69 del BBVA Banco Continental desde otro Banco."
          .Range(.Cells(34, 2), .Cells(35, 2)).Font.Size = 12
          .Range(.Cells(34, 2), .Cells(35, 2)).Font.Bold = True
       Else
          r_obj_Excel.Selection.NumberFormat = "[$$]#,##0.00;[Red][$$]#,##0.00"
          .Cells(34, 2) = "Realizar el depósito en la cuenta Nº 0011-0369-02-00090540 del BBVA Banco Continental."
          .Cells(34, 2).Font.Size = 12
          .Cells(34, 2).Font.Bold = True
       End If
       
       'MARCO
       .Cells(r_int_NroFil, r_int_ColumC).Borders(xlEdgeLeft).Weight = xlThin
       .Cells(r_int_NroFil, r_int_ColumK).Borders(xlEdgeLeft).Weight = xlThin
       .Cells(r_int_NroFil, r_int_ColumL).Borders(xlEdgeRight).Weight = xlThin
       .Range("C" & r_int_NroFil & ":L" & r_int_NroFil & "").Borders(xlEdgeTop).LineStyle = xlContinuous
       .Range("C" & r_int_NroFil & ":L" & r_int_NroFil & "").Borders(xlEdgeBottom).LineStyle = xlContinuous
       r_int_NroFil = r_int_NroFil + 2
    End With
    
    r_obj_Excel.Visible = True
    Set r_obj_Excel = Nothing
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
    Call fs_Buscar_Credito
    Call gs_CentraForm(Me)
    
    Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   'Inicializando Grid de Datos del Crédito
   grd_Listad.ColWidth(0) = 2900
   grd_Listad.ColWidth(1) = 8150
   grd_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
   
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
   pnl_DevPbp.Caption = 0
   pnl_IntLeg.Caption = 0
   pnl_CapPbpPerdido.Caption = "0.00 "
   pnl_IntPbpPerdido.Caption = "0.00 "
   pnl_MontoITF.Caption = 0
   pnl_MtoApl.Caption = "0.00 "
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_MotPpg, 1, "115")
   cmb_MotPpg.ListIndex = -1
End Sub

Private Sub fs_Limpia()
   Call gs_LimpiaGrid(grd_Listad)
End Sub

Private Sub fs_Buscar()
   
   'Obtiene Datos del Crédito
   Call modmip_gs_DatNumOpe(moddat_g_str_NumOpe, grd_Listad)
   l_int_CodPrd = moddat_g_str_CodPrd
   
   'DATOS DE LA OPERACION DEL PREPAGO
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
   pnl_SaldoTNC1.Caption = Format(g_rst_Princi!PPGCAB_SLDACT_TNC, "###,##0.00") & " "
   pnl_SaldoTC1.Caption = Format(g_rst_Princi!PPGCAB_SLDACT_TC, "###,##0.00") & " "
   pnl_SegDes.Caption = Format(g_rst_Princi!PPGCAB_SEGDES, "###,##0.00") & " "
   pnl_UltPagTNC.Caption = gf_FormatoFecha(g_rst_Princi!PPGCAB_ULTPAG_TNC)
   pnl_UltPagTC.Caption = gf_FormatoFecha(g_rst_Princi!PPGCAB_ULTPAG_TC)
   pnl_SegInm.Caption = Format(g_rst_Princi!PPGCAB_SEGINM, "###,##0.00") & " "
   pnl_DevPbp.Caption = Format(g_rst_Princi!PPGCAB_DEVBBP, "###,##0.00") & " "
   pnl_IntLeg.Caption = Format(g_rst_Princi!PPGCAB_INTLEG, "###,##0.00") & " "
   pnl_DiasTNC.Caption = CInt(g_rst_Princi!PPGCAB_DIFDIA_TNC) & " "
   pnl_DiasTC.Caption = CInt(g_rst_Princi!PPGCAB_DIFDIA_TC) & " "
   pnl_MontoPortes.Caption = Format(g_rst_Princi!PPGCAB_MTOPOR, "###,##0.00") & " "
   pnl_InteresTNC.Caption = Format(g_rst_Princi!PPGCAB_INTCAL_TNC, "###,##0.00") & " "
   pnl_InteresTC.Caption = Format(g_rst_Princi!PPGCAB_INTCAL_TC, "###,##0.00") & " "
   If Not IsNull(g_rst_Princi!PPGCAB_PBPPER) Then
      pnl_CapPbpPerdido.Caption = Format(g_rst_Princi!PPGCAB_PBPPER, "###,##0.00") & " "
   End If
   If Not IsNull(g_rst_Princi!PPGCAB_PBPINT) Then
      pnl_IntPbpPerdido.Caption = Format(g_rst_Princi!PPGCAB_PBPINT, "###,##0.00") & " "
   End If
   pnl_MtoApl.Caption = gf_FormatoNumero(CDbl(pnl_SaldoTNC1.Caption) + CDbl(pnl_SaldoTC1.Caption) + CDbl(pnl_InteresTNC.Caption) + CDbl(pnl_InteresTC.Caption) + CDbl(pnl_SegDes.Caption) + CDbl(pnl_SegInm.Caption) + CDbl(pnl_MontoPortes.Caption) + CDbl(pnl_CapPbpPerdido.Caption) + CDbl(pnl_IntPbpPerdido.Caption) + CDbl(pnl_DevPbp.Caption) + CDbl(pnl_IntLeg.Caption), 12, 2) & " "
   pnl_MontoITF.Caption = Format(g_rst_Princi!PPGCAB_MTOITF, "###,##0.00") & " "
   pnl_MtoPrepagar.Caption = Format(g_rst_Princi!PPGCAB_MTOTOT, "###,##0.00") & " "
   If g_rst_Princi!PPGCAB_MOTPPG > 0 Then
      Call gs_BuscarCombo_Item(cmb_MotPpg, g_rst_Princi!PPGCAB_MOTPPG)
   End If
   txt_ObsPpg.Text = Trim(IIf(IsNull(g_rst_Princi!PPGCAB_COMENT), " ", g_rst_Princi!PPGCAB_COMENT))
   
   'DETERMINA SI OPERACION ES MICASITA
   If l_int_CodPrd = 2 Or l_int_CodPrd = 11 Then
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
      Label10.Visible = False
      pnl_CapPbpPerdido.Visible = False
      pnl_IntPbpPerdido.Visible = False
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
      Label10.Visible = True
      pnl_CapPbpPerdido.Visible = True
      pnl_IntPbpPerdido.Visible = True
   End If
   
   Call gs_UbiIniGrid(grd_Listad)
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

Private Sub fs_Buscar_Credito()
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

Private Sub txt_ObsPpg_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub
