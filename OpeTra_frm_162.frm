VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_MntCli_52 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   9660
   ClientLeft      =   3420
   ClientTop       =   2265
   ClientWidth     =   11610
   Icon            =   "OpeTra_frm_162.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9660
   ScaleWidth      =   11610
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   9675
      Left            =   0
      TabIndex        =   49
      Top             =   0
      Width           =   11655
      _Version        =   65536
      _ExtentX        =   20558
      _ExtentY        =   17066
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
      Begin Threed.SSPanel SSPanel10 
         Height          =   435
         Left            =   30
         TabIndex        =   94
         Top             =   6570
         Width           =   11565
         _Version        =   65536
         _ExtentX        =   20399
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
         Begin VB.TextBox txt_TelFij 
            Height          =   315
            Left            =   8160
            MaxLength       =   25
            TabIndex        =   28
            Text            =   "Text1"
            Top             =   60
            Width           =   3315
         End
         Begin VB.ComboBox cmb_PaiRes 
            Height          =   315
            Left            =   1920
            TabIndex        =   27
            Text            =   "cmb_PaiRes"
            Top             =   60
            Width           =   3315
         End
         Begin VB.Label Label30 
            Caption         =   "Teléfono Fijo:"
            Height          =   285
            Left            =   6180
            TabIndex        =   96
            Top             =   60
            Width           =   1485
         End
         Begin VB.Label Label32 
            Caption         =   "País Residencia:"
            Height          =   315
            Left            =   60
            TabIndex        =   95
            Top             =   60
            Width           =   1455
         End
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   1755
         Left            =   30
         TabIndex        =   80
         Top             =   7050
         Width           =   11565
         _Version        =   65536
         _ExtentX        =   20399
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
         Begin VB.ComboBox cmb_TipVia 
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   60
            Width           =   3315
         End
         Begin VB.TextBox txt_NomVia 
            Height          =   315
            Left            =   1920
            MaxLength       =   120
            TabIndex        =   30
            Text            =   "Text1"
            Top             =   390
            Width           =   3315
         End
         Begin VB.TextBox txt_NumVia 
            Height          =   315
            Left            =   8160
            MaxLength       =   30
            TabIndex        =   31
            Text            =   "Text1"
            Top             =   390
            Width           =   1640
         End
         Begin VB.TextBox txt_IntDpt 
            Height          =   315
            Left            =   9840
            MaxLength       =   30
            TabIndex        =   32
            Text            =   "Text1"
            Top             =   390
            Width           =   1640
         End
         Begin VB.ComboBox cmb_TipZon 
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   720
            Width           =   3315
         End
         Begin VB.TextBox txt_NomZon 
            Height          =   315
            Left            =   8160
            MaxLength       =   120
            TabIndex        =   34
            Text            =   "Text1"
            Top             =   720
            Width           =   3315
         End
         Begin VB.ComboBox cmb_DptDir 
            Height          =   315
            Left            =   1920
            TabIndex        =   35
            Text            =   "cmb_DptDir"
            Top             =   1050
            Width           =   3315
         End
         Begin VB.ComboBox cmb_PrvDir 
            Height          =   315
            Left            =   8160
            TabIndex        =   36
            Text            =   "cmb_PrvDir"
            Top             =   1050
            Width           =   3315
         End
         Begin VB.ComboBox cmb_DstDir 
            Height          =   315
            Left            =   1920
            TabIndex        =   37
            Text            =   "cmb_DstDir"
            Top             =   1380
            Width           =   3315
         End
         Begin VB.TextBox txt_Refere 
            Height          =   315
            Left            =   8160
            MaxLength       =   250
            TabIndex        =   38
            Text            =   "Text1"
            Top             =   1380
            Width           =   3315
         End
         Begin VB.Label Label19 
            Caption         =   "Tipo de Vía:"
            Height          =   315
            Left            =   60
            TabIndex        =   89
            Top             =   60
            Width           =   1905
         End
         Begin VB.Label Label20 
            Caption         =   "Nombre Vía:"
            Height          =   285
            Left            =   60
            TabIndex        =   88
            Top             =   390
            Width           =   1485
         End
         Begin VB.Label Label21 
            Caption         =   "Nro - Int/Dpto/Mza/Lote:"
            Height          =   285
            Left            =   6180
            TabIndex        =   87
            Top             =   390
            Width           =   2055
         End
         Begin VB.Label Label22 
            Caption         =   "Tipo de Zona:"
            Height          =   315
            Left            =   60
            TabIndex        =   86
            Top             =   720
            Width           =   1905
         End
         Begin VB.Label Label23 
            Caption         =   "Nombre Zona:"
            Height          =   285
            Left            =   6180
            TabIndex        =   85
            Top             =   720
            Width           =   1485
         End
         Begin VB.Label Label24 
            Caption         =   "Departamento:"
            Height          =   315
            Left            =   60
            TabIndex        =   84
            Top             =   1050
            Width           =   1905
         End
         Begin VB.Label Label25 
            Caption         =   "Provincia:"
            Height          =   315
            Left            =   6180
            TabIndex        =   83
            Top             =   1050
            Width           =   1905
         End
         Begin VB.Label Label26 
            Caption         =   "Distrito:"
            Height          =   315
            Left            =   60
            TabIndex        =   82
            Top             =   1380
            Width           =   1905
         End
         Begin VB.Label Label28 
            Caption         =   "Referencia:"
            Height          =   285
            Left            =   6180
            TabIndex        =   81
            Top             =   1380
            Width           =   1485
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   795
         Left            =   30
         TabIndex        =   76
         Top             =   1950
         Width           =   11565
         _Version        =   65536
         _ExtentX        =   20399
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
         Begin VB.TextBox txt_NDoAlt 
            Height          =   315
            Left            =   8160
            MaxLength       =   12
            TabIndex        =   2
            Text            =   "Text1"
            Top             =   390
            Width           =   3315
         End
         Begin VB.ComboBox cmb_TDoAlt 
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   390
            Width           =   3315
         End
         Begin VB.ComboBox cmb_DocAlt 
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   1065
         End
         Begin VB.Label Label33 
            Caption         =   "Nro. Doc. Id.:"
            Height          =   285
            Left            =   6180
            TabIndex        =   79
            Top             =   390
            Width           =   1065
         End
         Begin VB.Label Label34 
            Caption         =   "Tipo Docum. Identidad:"
            Height          =   315
            Left            =   60
            TabIndex        =   78
            Top             =   390
            Width           =   1845
         End
         Begin VB.Label Label35 
            Caption         =   "Personal FF.AA / FF.PP:"
            Height          =   315
            Left            =   60
            TabIndex        =   77
            Top             =   60
            Width           =   1845
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   435
         Left            =   30
         TabIndex        =   50
         Top             =   1470
         Width           =   11565
         _Version        =   65536
         _ExtentX        =   20399
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
         Begin Threed.SSPanel pnl_DocIde 
            Height          =   315
            Left            =   1950
            TabIndex        =   51
            Top             =   60
            Width           =   9555
            _Version        =   65536
            _ExtentX        =   16854
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "1 - 07522154"
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
         Begin VB.Label Label1 
            Caption         =   "Docum. de Identidad:"
            Height          =   315
            Left            =   60
            TabIndex        =   52
            Top             =   60
            Width           =   1725
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   53
         Top             =   30
         Width           =   11565
         _Version        =   65536
         _ExtentX        =   20399
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
         Begin Threed.SSPanel pnl_RelAcc 
            Height          =   285
            Left            =   5700
            TabIndex        =   98
            Top             =   330
            Width           =   5835
            _Version        =   65536
            _ExtentX        =   10292
            _ExtentY        =   503
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
         End
         Begin Threed.SSPanel pnl_RelLab 
            Height          =   285
            Left            =   5700
            TabIndex        =   97
            Top             =   60
            Width           =   5835
            _Version        =   65536
            _ExtentX        =   10292
            _ExtentY        =   503
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
         End
         Begin Threed.SSPanel SSPanel16 
            Height          =   255
            Left            =   600
            TabIndex        =   54
            Top             =   60
            Width           =   4845
            _Version        =   65536
            _ExtentX        =   8546
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "Mantenimiento de Clientes"
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
         Begin Threed.SSPanel SSPanel7 
            Height          =   255
            Left            =   600
            TabIndex        =   55
            Top             =   330
            Width           =   4785
            _Version        =   65536
            _ExtentX        =   8440
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "Datos Generales"
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
            Picture         =   "OpeTra_frm_162.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   3735
         Left            =   30
         TabIndex        =   56
         Top             =   2790
         Width           =   11565
         _Version        =   65536
         _ExtentX        =   20399
         _ExtentY        =   6588
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
         Begin VB.ComboBox cmb_TipAfp 
            Height          =   315
            Left            =   8160
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   2340
            Width           =   3315
         End
         Begin VB.CheckBox chk_DirEle 
            Caption         =   "Autoriz. Corresp."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   9870
            TabIndex        =   20
            Top             =   3030
            Width           =   1485
         End
         Begin VB.TextBox txt_DirEle 
            Height          =   315
            Left            =   8160
            MaxLength       =   120
            TabIndex        =   19
            Text            =   "Text1"
            Top             =   3030
            Width           =   1665
         End
         Begin VB.TextBox txt_Celula 
            Height          =   315
            Left            =   1920
            MaxLength       =   25
            TabIndex        =   18
            Text            =   "Text1"
            Top             =   3030
            Width           =   3315
         End
         Begin VB.ComboBox cmb_Profes 
            Height          =   315
            Left            =   1920
            TabIndex        =   17
            Text            =   "cmb_Profes"
            Top             =   2700
            Width           =   9555
         End
         Begin VB.ComboBox cmb_NivEst 
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   2370
            Width           =   3315
         End
         Begin VB.ComboBox cmb_RegCyg 
            Height          =   315
            Left            =   8160
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   2040
            Width           =   3315
         End
         Begin VB.ComboBox cmb_EstCiv 
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   2040
            Width           =   3315
         End
         Begin VB.ComboBox cmb_DstNac 
            Height          =   315
            Left            =   8160
            TabIndex        =   12
            Text            =   "cmb_DstNac"
            Top             =   1710
            Width           =   3315
         End
         Begin VB.ComboBox cmb_PrvNac 
            Height          =   315
            Left            =   1920
            TabIndex        =   11
            Text            =   "cmb_PrvNac"
            Top             =   1710
            Width           =   3315
         End
         Begin VB.ComboBox cmb_DptNac 
            Height          =   315
            Left            =   8160
            TabIndex        =   10
            Text            =   "cmb_DptNac"
            Top             =   1380
            Width           =   3315
         End
         Begin VB.ComboBox cmb_Paises 
            Height          =   315
            Left            =   1920
            TabIndex        =   9
            Text            =   "cmb_Paises"
            Top             =   1380
            Width           =   3315
         End
         Begin VB.ComboBox cmb_CodSex 
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   1050
            Width           =   3315
         End
         Begin VB.TextBox txt_Nombre 
            Height          =   315
            Left            =   1920
            MaxLength       =   30
            TabIndex        =   6
            Text            =   "Text1"
            Top             =   720
            Width           =   3315
         End
         Begin VB.TextBox txt_ApePat 
            Height          =   315
            Left            =   1920
            MaxLength       =   30
            TabIndex        =   3
            Text            =   "Text1"
            Top             =   60
            Width           =   3315
         End
         Begin VB.TextBox txt_ApeMat 
            Height          =   315
            Left            =   1920
            MaxLength       =   30
            TabIndex        =   4
            Text            =   "Text1"
            Top             =   390
            Width           =   3315
         End
         Begin VB.TextBox txt_ApeCas 
            Height          =   315
            Left            =   8160
            MaxLength       =   30
            TabIndex        =   5
            Text            =   "Text1"
            Top             =   390
            Width           =   3315
         End
         Begin EditLib.fpLongInteger ipp_DepEc1 
            Height          =   315
            Left            =   8160
            TabIndex        =   22
            Top             =   3360
            Width           =   630
            _Version        =   196608
            _ExtentX        =   1111
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
            MaxValue        =   "99"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
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
         Begin EditLib.fpDateTime ipp_FecNac 
            Height          =   315
            Left            =   8160
            TabIndex        =   8
            Top             =   1050
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
         Begin EditLib.fpLongInteger ipp_DepEc2 
            Height          =   315
            Left            =   8790
            TabIndex        =   23
            Top             =   3360
            Width           =   630
            _Version        =   196608
            _ExtentX        =   1111
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
            MaxValue        =   "99"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
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
         Begin EditLib.fpLongInteger ipp_DepEc3 
            Height          =   315
            Left            =   9450
            TabIndex        =   24
            Top             =   3360
            Width           =   660
            _Version        =   196608
            _ExtentX        =   1164
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
            MaxValue        =   "99"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
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
         Begin EditLib.fpLongInteger ipp_DepEc4 
            Height          =   315
            Left            =   10110
            TabIndex        =   25
            Top             =   3360
            Width           =   630
            _Version        =   196608
            _ExtentX        =   1111
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
            MaxValue        =   "99"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
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
         Begin EditLib.fpLongInteger ipp_NumDep 
            Height          =   315
            Left            =   1920
            TabIndex        =   21
            Top             =   3360
            Width           =   735
            _Version        =   196608
            _ExtentX        =   1296
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
            MaxValue        =   "99"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
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
         Begin EditLib.fpLongInteger ipp_DepEc5 
            Height          =   315
            Left            =   10740
            TabIndex        =   26
            Top             =   3360
            Width           =   630
            _Version        =   196608
            _ExtentX        =   1111
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
            MaxValue        =   "99"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
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
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Afp:"
            Height          =   195
            Left            =   6180
            TabIndex        =   99
            Top             =   2340
            Width           =   285
         End
         Begin VB.Label Label38 
            Caption         =   "Nro. Depend. Econom.:"
            Height          =   285
            Left            =   60
            TabIndex        =   74
            Top             =   3360
            Width           =   1815
         End
         Begin VB.Label Label18 
            Caption         =   "Edades Depend. Econom.:"
            Height          =   285
            Left            =   6180
            TabIndex        =   73
            Top             =   3360
            Width           =   2055
         End
         Begin VB.Label Label17 
            Caption         =   "E-mail:"
            Height          =   285
            Left            =   6180
            TabIndex        =   72
            Top             =   3030
            Width           =   1485
         End
         Begin VB.Label Label16 
            Caption         =   "Teléfono Celular:"
            Height          =   285
            Left            =   60
            TabIndex        =   71
            Top             =   3030
            Width           =   1485
         End
         Begin VB.Label Label15 
            Caption         =   "Profesión o Actividad:"
            Height          =   315
            Left            =   60
            TabIndex        =   70
            Top             =   2700
            Width           =   1905
         End
         Begin VB.Label Label14 
            Caption         =   "Nivel de Estudio:"
            Height          =   315
            Left            =   60
            TabIndex        =   69
            Top             =   2370
            Width           =   1905
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Régimen Conyugal:"
            Height          =   195
            Left            =   6180
            TabIndex        =   68
            Top             =   2040
            Width           =   1380
         End
         Begin VB.Label Label12 
            Caption         =   "Estado Civil:"
            Height          =   315
            Left            =   60
            TabIndex        =   67
            Top             =   2040
            Width           =   1905
         End
         Begin VB.Label Label11 
            Caption         =   "Distrito Nacimiento:"
            Height          =   315
            Left            =   6180
            TabIndex        =   66
            Top             =   1710
            Width           =   1905
         End
         Begin VB.Label Label10 
            Caption         =   "Provincia Nacimiento:"
            Height          =   315
            Left            =   60
            TabIndex        =   65
            Top             =   1710
            Width           =   1905
         End
         Begin VB.Label Label9 
            Caption         =   "Dpto. Nacimiento:"
            Height          =   315
            Left            =   6180
            TabIndex        =   64
            Top             =   1380
            Width           =   1905
         End
         Begin VB.Label Label8 
            Caption         =   "Nacionalidad:"
            Height          =   315
            Left            =   60
            TabIndex        =   63
            Top             =   1380
            Width           =   1905
         End
         Begin VB.Label Label7 
            Caption         =   "Fecha de Nacimiento:"
            Height          =   315
            Left            =   6180
            TabIndex        =   62
            Top             =   1050
            Width           =   1905
         End
         Begin VB.Label Label6 
            Caption         =   "Sexo:"
            Height          =   315
            Left            =   60
            TabIndex        =   61
            Top             =   1050
            Width           =   1905
         End
         Begin VB.Label Label5 
            Caption         =   "Nombres:"
            Height          =   285
            Left            =   60
            TabIndex        =   60
            Top             =   720
            Width           =   1485
         End
         Begin VB.Label Label3 
            Caption         =   "Apellido Paterno:"
            Height          =   285
            Left            =   60
            TabIndex        =   59
            Top             =   60
            Width           =   1485
         End
         Begin VB.Label Label4 
            Caption         =   "Apellido Materno:"
            Height          =   285
            Left            =   60
            TabIndex        =   58
            Top             =   390
            Width           =   1485
         End
         Begin VB.Label Label29 
            Caption         =   "Apellido Casada:"
            Height          =   285
            Left            =   6180
            TabIndex        =   57
            Top             =   390
            Width           =   1485
         End
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   675
         Left            =   30
         TabIndex        =   75
         Top             =   750
         Width           =   11565
         _Version        =   65536
         _ExtentX        =   20399
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
         Begin VB.CommandButton cmd_Cancel 
            Height          =   585
            Left            =   3030
            Picture         =   "OpeTra_frm_162.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   47
            ToolTipText     =   "Cancelar Modificación de Datos Generales"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_DatApo 
            Height          =   585
            Left            =   1830
            Picture         =   "OpeTra_frm_162.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   45
            ToolTipText     =   "Datos del Apoderado"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Editar 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_162.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   42
            ToolTipText     =   "Modificar Datos del Cliente"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10950
            Picture         =   "OpeTra_frm_162.frx":0C34
            Style           =   1  'Graphical
            TabIndex        =   48
            ToolTipText     =   "Salir de la Ventana"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   2430
            Picture         =   "OpeTra_frm_162.frx":1076
            Style           =   1  'Graphical
            TabIndex        =   46
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_DatCyg 
            Height          =   585
            Left            =   1230
            Picture         =   "OpeTra_frm_162.frx":14B8
            Style           =   1  'Graphical
            TabIndex        =   44
            ToolTipText     =   "Datos del Cónyuge"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ActEco 
            Height          =   585
            Left            =   630
            Picture         =   "OpeTra_frm_162.frx":162F
            Style           =   1  'Graphical
            TabIndex        =   43
            ToolTipText     =   "Actividades Económicas"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   765
         Left            =   30
         TabIndex        =   90
         Top             =   8850
         Width           =   11565
         _Version        =   65536
         _ExtentX        =   20399
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
         Begin VB.TextBox txt_Direcc 
            Height          =   315
            Left            =   1920
            MaxLength       =   250
            TabIndex        =   39
            Text            =   "Text1"
            Top             =   60
            Width           =   9555
         End
         Begin VB.ComboBox cmb_PrvEst 
            Height          =   315
            Left            =   1920
            TabIndex        =   40
            Text            =   "cmb_DptDir"
            Top             =   390
            Width           =   3315
         End
         Begin VB.TextBox txt_CodPos 
            Height          =   315
            Left            =   8160
            MaxLength       =   250
            TabIndex        =   41
            Text            =   "Text1"
            Top             =   390
            Width           =   1640
         End
         Begin VB.Label Label36 
            Caption         =   "Dirección:"
            Height          =   285
            Left            =   60
            TabIndex        =   93
            Top             =   60
            Width           =   1485
         End
         Begin VB.Label Label31 
            Caption         =   "Provincia / Estado:"
            Height          =   315
            Left            =   60
            TabIndex        =   92
            Top             =   390
            Width           =   1905
         End
         Begin VB.Label Label2 
            Caption         =   "Código Postal:"
            Height          =   285
            Left            =   6180
            TabIndex        =   91
            Top             =   390
            Width           =   1485
         End
      End
   End
End
Attribute VB_Name = "frm_MntCli_52"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_Paises()   As moddat_tpo_Genera
Dim l_arr_PaiRes()   As moddat_tpo_Genera
Dim l_arr_Profes()   As moddat_tpo_Genera
Dim l_arr_PrvEst()   As moddat_tpo_Genera
Dim l_str_Paises     As String
Dim l_str_DptNac     As String
Dim l_str_PrvNac     As String
Dim l_str_DstNac     As String
Dim l_str_Profes     As String
Dim l_str_PaiRes     As String
Dim l_str_DptDir     As String
Dim l_str_PrvDir     As String
Dim l_str_DstDir     As String
Dim l_str_PrvEst     As String
Dim l_int_FlgCmb     As Integer
Dim l_int_RelLab     As Integer
Dim l_int_VinTDo     As Integer
Dim l_str_VinNDo     As String
Dim l_int_VinTip     As Integer
Dim l_int_RelAcc     As Integer
Dim l_int_AccTDo     As Integer
Dim l_str_AccNDo     As String
Dim l_int_AccVin     As Integer
Dim l_int_TipVia     As Integer
Dim l_str_NomVia     As String
Dim l_str_NumVia     As String
Dim l_str_IntDpt     As String
Dim l_int_TipZon     As Integer
Dim l_str_NomZon     As String
Dim l_str_UbiGeo     As String
Dim l_str_Refere     As String

Private Sub cmb_CodSex_Click()
   Call gs_SetFocus(ipp_FecNac)
End Sub

Private Sub cmb_CodSex_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_CodSex_Click
   End If
End Sub

Private Sub cmb_DocAlt_Click()
   If cmb_DocAlt.ListIndex = -1 Then
      cmb_TDoAlt.ListIndex = -1
      txt_NDoAlt.Text = ""
      
      cmb_TDoAlt.Enabled = False
      txt_NDoAlt.Enabled = False
   Else
      If cmb_DocAlt.ItemData(cmb_DocAlt.ListIndex) = 1 Then
         cmb_TDoAlt.Enabled = True
         txt_NDoAlt.Enabled = True
         
         Call gs_SetFocus(cmb_TDoAlt)
      Else
         cmb_TDoAlt.ListIndex = -1
         txt_NDoAlt.Text = ""
         
         cmb_TDoAlt.Enabled = False
         txt_NDoAlt.Enabled = False
      
         Call gs_SetFocus(txt_ApePat)
      End If
   End If
End Sub

Private Sub cmb_DocAlt_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_DocAlt_Click
   End If
End Sub

Private Sub cmb_DptDir_Change()
   l_str_DptDir = cmb_DptDir.Text
End Sub

Private Sub cmb_DptDir_Click()
   If cmb_DptDir.ListIndex > -1 Then
      If l_int_FlgCmb Then
         cmb_PrvDir.Clear
         cmb_DstDir.Clear
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Provin(cmb_PrvDir, Format(cmb_DptDir.ItemData(cmb_DptDir.ListIndex), "00"))
         Screen.MousePointer = 0
         
         Call gs_SetFocus(cmb_PrvDir)
      End If
   End If
End Sub

Private Sub cmb_DptDir_GotFocus()
   l_int_FlgCmb = True
   l_str_DptDir = cmb_DptDir.Text
End Sub

Private Sub cmb_DptDir_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_DptDir, l_str_DptDir)
      l_int_FlgCmb = True
      
      cmb_PrvDir.Clear
      cmb_DstDir.Clear
      If cmb_DptDir.ListIndex > -1 Then
         l_str_DptDir = ""
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Provin(cmb_PrvDir, Format(cmb_DptDir.ItemData(cmb_DptDir.ListIndex), "00"))
         Screen.MousePointer = 0
      End If
      
      Call gs_SetFocus(cmb_PrvDir)
   End If
End Sub

Private Sub cmb_DptNac_Change()
   l_str_DptNac = cmb_DptNac.Text
End Sub

Private Sub cmb_DptNac_Click()
   If cmb_DptNac.ListIndex > -1 Then
      If l_int_FlgCmb Then
         cmb_PrvNac.Clear
         cmb_DstNac.Clear
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Provin(cmb_PrvNac, Format(cmb_DptNac.ItemData(cmb_DptNac.ListIndex), "00"))
         Screen.MousePointer = 0
         
         Call gs_SetFocus(cmb_PrvNac)
      End If
   End If
End Sub

Private Sub cmb_DptNac_GotFocus()
   l_int_FlgCmb = True
   l_str_DptNac = cmb_DptNac.Text
End Sub

Private Sub cmb_DptNac_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_DptNac, l_str_DptNac)
      l_int_FlgCmb = True
      
      cmb_PrvNac.Clear
      cmb_DstNac.Clear
      If cmb_DptNac.ListIndex > -1 Then
         l_str_DptNac = ""
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Provin(cmb_PrvNac, Format(cmb_DptNac.ItemData(cmb_DptNac.ListIndex), "00"))
         Screen.MousePointer = 0
      End If
      
      Call gs_SetFocus(cmb_PrvNac)
   End If
End Sub

Private Sub cmb_DstDir_Change()
   l_str_DstDir = cmb_DstDir.Text
End Sub

Private Sub cmb_DstDir_Click()
   If cmb_DstDir.ListIndex > -1 Then
      If l_int_FlgCmb Then
         Call gs_SetFocus(txt_Refere)
      End If
   End If
End Sub

Private Sub cmb_DstDir_GotFocus()
   l_int_FlgCmb = True
   l_str_DstDir = cmb_DstDir.Text
End Sub

Private Sub cmb_DstDir_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_DstDir, l_str_DstDir)
      l_int_FlgCmb = True
      
      If cmb_DstDir.ListIndex > -1 Then
         l_str_DstDir = ""
      End If
      
      Call gs_SetFocus(txt_Refere)
   End If
End Sub

Private Sub cmb_EstCiv_Click()
   If cmb_EstCiv.ListIndex > -1 Then
      If cmb_EstCiv.ItemData(cmb_EstCiv.ListIndex) = 2 Then
         cmb_RegCyg.Enabled = True
         Call gs_SetFocus(cmb_RegCyg)
      ElseIf cmb_EstCiv.ItemData(cmb_EstCiv.ListIndex) = 5 Then
         Call gs_SetFocus(cmb_NivEst)
      Else
         cmb_RegCyg.ListIndex = -1
         cmb_RegCyg.Enabled = False
         
         Call gs_SetFocus(cmb_NivEst)
      End If
   Else
      cmb_RegCyg.ListIndex = -1
      cmb_RegCyg.Enabled = False
   End If
End Sub

Private Sub cmb_EstCiv_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_EstCiv_Click
   End If
End Sub

Private Sub cmb_NivEst_Click()
   Call gs_SetFocus(cmb_TipAfp)
End Sub

Private Sub cmb_NivEst_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_NivEst_Click
   End If
End Sub

Private Sub cmb_Paises_Change()
   l_str_Paises = cmb_Paises.Text
   
   cmb_Paises.SelLength = Len(l_str_Paises)
End Sub

Private Sub cmb_Paises_Click()
   If cmb_Paises.ListIndex > -1 Then
      If l_int_FlgCmb Then
         cmb_DptNac.Enabled = True
         cmb_PrvNac.Enabled = True
         cmb_DstNac.Enabled = True
         
         If l_arr_Paises(cmb_Paises.ListIndex + 1).Genera_Codigo <> "004028" Then
            cmb_DptNac.ListIndex = -1
            cmb_PrvNac.Clear
            cmb_DstNac.Clear
            
            cmb_DptNac.Enabled = False
            cmb_PrvNac.Enabled = False
            cmb_DstNac.Enabled = False
         
            Call gs_SetFocus(cmb_EstCiv)
         Else
            Call gs_SetFocus(cmb_DptNac)
         End If
      End If
   Else
      cmb_DptNac.ListIndex = -1
      cmb_PrvNac.Clear
      cmb_DstNac.Clear
      
      cmb_DptNac.Enabled = False
      cmb_PrvNac.Enabled = False
      cmb_DstNac.Enabled = False
   
      Call gs_SetFocus(cmb_EstCiv)
   End If
End Sub

Private Sub cmb_Paises_GotFocus()
   l_int_FlgCmb = True
End Sub

Private Sub cmb_Paises_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_Paises, l_str_Paises)
      l_int_FlgCmb = True
      
      cmb_DptNac.Enabled = True
      cmb_PrvNac.Enabled = True
      cmb_DstNac.Enabled = True
      
      If cmb_Paises.ListIndex > -1 Then
         l_str_Paises = ""
         
         If l_arr_Paises(cmb_Paises.ListIndex + 1).Genera_Codigo <> "004028" Then
            cmb_DptNac.ListIndex = -1
            cmb_PrvNac.Clear
            cmb_DstNac.Clear
            
            cmb_DptNac.Enabled = False
            cmb_PrvNac.Enabled = False
            cmb_DstNac.Enabled = False
         
            Call gs_SetFocus(cmb_EstCiv)
         Else
            Call gs_SetFocus(cmb_DptNac)
         End If
      Else
         cmb_DptNac.ListIndex = -1
         cmb_PrvNac.Clear
         cmb_DstNac.Clear
      
         cmb_DptNac.Enabled = False
         cmb_PrvNac.Enabled = False
         cmb_DstNac.Enabled = False
   
         Call gs_SetFocus(cmb_EstCiv)
      End If
   End If
End Sub

Private Sub cmb_PaiRes_Change()
   l_str_PaiRes = cmb_PaiRes.Text
   
   cmb_PaiRes.SelLength = Len(l_str_PaiRes)
End Sub

Private Sub cmb_PaiRes_Click()
   If cmb_PaiRes.ListIndex > -1 Then
      If l_int_FlgCmb Then
         If l_arr_PaiRes(cmb_PaiRes.ListIndex + 1).Genera_Codigo <> "004028" Then
            cmb_TipVia.ListIndex = -1
            txt_NomVia.Text = ""
            txt_NumVia.Text = ""
            txt_IntDpt.Text = ""
            cmb_TipZon.ListIndex = -1
            txt_NomZon.Text = ""
            cmb_DptDir.ListIndex = -1
            cmb_PrvDir.Clear
            cmb_DstDir.Clear
            txt_Refere.Text = ""
            
            cmb_TipVia.Enabled = False
            txt_NomVia.Enabled = False
            txt_NumVia.Enabled = False
            txt_IntDpt.Enabled = False
            cmb_TipZon.Enabled = False
            txt_NomZon.Enabled = False
            cmb_DptDir.Enabled = False
            cmb_PrvDir.Enabled = False
            cmb_DstDir.Enabled = False
            txt_Refere.Enabled = False
            
            txt_Direcc.Enabled = True
            cmb_PrvEst.Enabled = True
            txt_CodPos.Enabled = True
            
            'Cargar Provincia / Estado segñun País seleccionado
            Call modmip_gs_Carga_CiuExt(l_arr_PrvEst, cmb_PrvEst, l_arr_PaiRes(cmb_PaiRes.ListIndex + 1).Genera_Codigo)
         Else
            cmb_TipVia.Enabled = True
            txt_NomVia.Enabled = True
            txt_NumVia.Enabled = True
            txt_IntDpt.Enabled = True
            cmb_TipZon.Enabled = True
            txt_NomZon.Enabled = True
            cmb_DptDir.Enabled = True
            cmb_PrvDir.Enabled = True
            cmb_DstDir.Enabled = True
            txt_Refere.Enabled = True
         
            txt_Direcc.Text = ""
            cmb_PrvEst.Clear
            txt_CodPos.Text = ""
            
            txt_Direcc.Enabled = False
            cmb_PrvEst.Enabled = False
            txt_CodPos.Enabled = False
         End If
         
         Call gs_SetFocus(txt_TelFij)
      End If
   Else
      cmb_TipVia.ListIndex = -1
      txt_NomVia.Text = ""
      txt_NumVia.Text = ""
      txt_IntDpt.Text = ""
      cmb_TipZon.ListIndex = -1
      txt_NomZon.Text = ""
      cmb_DptDir.ListIndex = -1
      cmb_PrvDir.Clear
      cmb_DstDir.Clear
      txt_Refere.Text = ""
      
      cmb_TipVia.Enabled = False
      txt_NomVia.Enabled = False
      txt_NumVia.Enabled = False
      txt_IntDpt.Enabled = False
      cmb_TipZon.Enabled = False
      txt_NomZon.Enabled = False
      cmb_DptDir.Enabled = False
      cmb_PrvDir.Enabled = False
      cmb_DstDir.Enabled = False
      txt_Refere.Enabled = False
   
      txt_Direcc.Text = ""
      cmb_PrvEst.ListIndex = -1
      txt_CodPos.Text = ""
      
      txt_Direcc.Enabled = False
      cmb_PrvEst.Enabled = False
      txt_CodPos.Enabled = False
      
      Call gs_SetFocus(txt_TelFij)
   End If
End Sub

Private Sub cmb_PaiRes_GotFocus()
   l_int_FlgCmb = True
End Sub

Private Sub cmb_PaiRes_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_PaiRes, l_str_PaiRes)
      l_int_FlgCmb = True
      
      If cmb_PaiRes.ListIndex > -1 Then
         l_str_PaiRes = ""
         
         If l_arr_PaiRes(cmb_PaiRes.ListIndex + 1).Genera_Codigo <> "004028" Then
            cmb_TipVia.ListIndex = -1
            txt_NomVia.Text = ""
            txt_NumVia.Text = ""
            txt_IntDpt.Text = ""
            cmb_TipZon.ListIndex = -1
            txt_NomZon.Text = ""
            cmb_DptDir.ListIndex = -1
            cmb_PrvDir.Clear
            cmb_DstDir.Clear
            txt_Refere.Text = ""
            
            cmb_TipVia.Enabled = False
            txt_NomVia.Enabled = False
            txt_NumVia.Enabled = False
            txt_IntDpt.Enabled = False
            cmb_TipZon.Enabled = False
            txt_NomZon.Enabled = False
            cmb_DptDir.Enabled = False
            cmb_PrvDir.Enabled = False
            cmb_DstDir.Enabled = False
            txt_Refere.Enabled = False
            
            txt_Direcc.Enabled = True
            cmb_PrvEst.Enabled = True
            txt_CodPos.Enabled = True
            
            'Cargar Provincia/estado
            Call modmip_gs_Carga_CiuExt(l_arr_PrvEst, cmb_PrvEst, l_arr_PaiRes(cmb_PaiRes.ListIndex + 1).Genera_Codigo)
         Else
            cmb_TipVia.Enabled = True
            txt_NomVia.Enabled = True
            txt_NumVia.Enabled = True
            txt_IntDpt.Enabled = True
            cmb_TipZon.Enabled = True
            txt_NomZon.Enabled = True
            cmb_DptDir.Enabled = True
            cmb_PrvDir.Enabled = True
            cmb_DstDir.Enabled = True
            txt_Refere.Enabled = True
         
            txt_Direcc.Text = ""
            cmb_PrvEst.Clear
            txt_CodPos.Text = ""
            
            txt_Direcc.Enabled = False
            cmb_PrvEst.Enabled = False
            txt_CodPos.Enabled = False
         End If
         
         Call gs_SetFocus(txt_TelFij)
      Else
         cmb_TipVia.ListIndex = -1
         txt_NomVia.Text = ""
         txt_NumVia.Text = ""
         txt_IntDpt.Text = ""
         cmb_TipZon.ListIndex = -1
         txt_NomZon.Text = ""
         cmb_DptDir.ListIndex = -1
         cmb_PrvDir.Clear
         cmb_DstDir.Clear
         txt_Refere.Text = ""
         
         cmb_TipVia.Enabled = False
         txt_NomVia.Enabled = False
         txt_NumVia.Enabled = False
         txt_IntDpt.Enabled = False
         cmb_TipZon.Enabled = False
         txt_NomZon.Enabled = False
         cmb_DptDir.Enabled = False
         cmb_PrvDir.Enabled = False
         cmb_DstDir.Enabled = False
         txt_Refere.Enabled = False
      
         txt_Direcc.Text = ""
         cmb_PrvEst.ListIndex = -1
         txt_CodPos.Text = ""
         
         txt_Direcc.Enabled = False
         cmb_PrvEst.Enabled = False
         txt_CodPos.Enabled = False
         
         Call gs_SetFocus(txt_TelFij)
      End If
   End If
End Sub

Private Sub cmb_Profes_Change()
   l_str_Profes = cmb_Profes.Text
End Sub

Private Sub cmb_Profes_Click()
   If cmb_Profes.ListIndex > -1 Then
      If l_int_FlgCmb Then
         Call gs_SetFocus(txt_Celula)
      End If
   End If
End Sub

Private Sub cmb_Profes_GotFocus()
   l_int_FlgCmb = True
   l_str_Profes = cmb_Profes.Text
End Sub

Private Sub cmb_Profes_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./<>*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_Profes, l_str_Profes)
      l_int_FlgCmb = True
      
      If cmb_Profes.ListIndex > -1 Then
         l_str_Profes = ""
      End If
      
      Call gs_SetFocus(txt_Celula)
   End If
End Sub

Private Sub cmb_PrvDir_Change()
   l_str_PrvDir = cmb_PrvDir.Text
End Sub

Private Sub cmb_PrvDir_Click()
   If cmb_PrvDir.ListIndex > -1 Then
      If l_int_FlgCmb Then
         cmb_DstDir.Clear
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Distri(cmb_DstDir, Format(cmb_DptDir.ItemData(cmb_DptDir.ListIndex), "00"), Format(cmb_PrvDir.ItemData(cmb_PrvDir.ListIndex), "00"))
         Screen.MousePointer = 0
         
         Call gs_SetFocus(cmb_DstDir)
      End If
   End If
End Sub

Private Sub cmb_PrvDir_GotFocus()
   l_int_FlgCmb = True
   l_str_PrvDir = cmb_PrvDir.Text
End Sub

Private Sub cmb_PrvDir_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_PrvDir, l_str_PrvDir)
      l_int_FlgCmb = True
      
      cmb_DstDir.Clear
      If cmb_PrvDir.ListIndex > -1 Then
         l_str_DstDir = ""
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Distri(cmb_DstDir, Format(cmb_DptDir.ItemData(cmb_DptDir.ListIndex), "00"), Format(cmb_PrvDir.ItemData(cmb_PrvDir.ListIndex), "00"))
         Screen.MousePointer = 0
      End If
      
      Call gs_SetFocus(cmb_DstDir)
   End If
End Sub

Private Sub cmb_PrvEst_Change()
   l_str_PrvEst = cmb_PrvEst.Text
End Sub

Private Sub cmb_PrvEst_Click()
   If cmb_PrvEst.ListIndex > -1 Then
      If l_int_FlgCmb Then
         Call gs_SetFocus(txt_CodPos)
      End If
   End If
End Sub

Private Sub cmb_PrvEst_GotFocus()
   l_int_FlgCmb = True
   l_str_PrvEst = cmb_PrvEst.Text
End Sub

Private Sub cmb_PrvEst_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_PrvEst, l_str_PrvEst)
      l_int_FlgCmb = True
      
      If cmb_PrvEst.ListIndex > -1 Then
         l_str_PrvEst = ""
      End If
      
      Call gs_SetFocus(txt_CodPos)
   End If
End Sub

Private Sub cmb_PrvNac_Change()
   l_str_PrvNac = cmb_PrvNac.Text
End Sub

Private Sub cmb_PrvNac_Click()
   If cmb_PrvNac.ListIndex > -1 Then
      If l_int_FlgCmb Then
         cmb_DstNac.Clear
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Distri(cmb_DstNac, Format(cmb_DptNac.ItemData(cmb_DptNac.ListIndex), "00"), Format(cmb_PrvNac.ItemData(cmb_PrvNac.ListIndex), "00"))
         Screen.MousePointer = 0
         
         Call gs_SetFocus(cmb_DstNac)
      End If
   End If
End Sub

Private Sub cmb_PrvNac_GotFocus()
   l_int_FlgCmb = True
   l_str_PrvNac = cmb_PrvNac.Text
End Sub

Private Sub cmb_PrvNac_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_PrvNac, l_str_PrvNac)
      l_int_FlgCmb = True
      
      cmb_DstNac.Clear
      If cmb_PrvNac.ListIndex > -1 Then
         l_str_DstNac = ""
         
         Screen.MousePointer = 11
         Call moddat_gs_Carga_Distri(cmb_DstNac, Format(cmb_DptNac.ItemData(cmb_DptNac.ListIndex), "00"), Format(cmb_PrvNac.ItemData(cmb_PrvNac.ListIndex), "00"))
         Screen.MousePointer = 0
      End If
      
      Call gs_SetFocus(cmb_DstNac)
   End If
End Sub

Private Sub cmb_DstNac_Change()
   l_str_DstNac = cmb_DstNac.Text
End Sub

Private Sub cmb_DstNac_Click()
   If cmb_DstNac.ListIndex > -1 Then
      If l_int_FlgCmb Then
         Call gs_SetFocus(cmb_EstCiv)
      End If
   End If
End Sub

Private Sub cmb_DstNac_GotFocus()
   l_int_FlgCmb = True
   l_str_DstNac = cmb_DstNac.Text
End Sub

Private Sub cmb_DstNac_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_DstNac, l_str_DstNac)
      l_int_FlgCmb = True
      
      If cmb_DstNac.ListIndex > -1 Then
         l_str_DstNac = ""
      End If
      
      Call gs_SetFocus(cmb_EstCiv)
   End If
End Sub

Private Sub cmb_RegCyg_Click()
   Call gs_SetFocus(cmb_NivEst)
End Sub

Private Sub cmb_RegCyg_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_RegCyg_Click
   End If
End Sub

Private Sub cmb_TDoAlt_Click()
   If cmb_TDoAlt.ListIndex > -1 Then
      Select Case cmb_TDoAlt.ItemData(cmb_TDoAlt.ListIndex)
         Case 1:  txt_NDoAlt.MaxLength = 8
         Case Else:  txt_NDoAlt.MaxLength = 12
      End Select
   End If
   
   Call gs_SetFocus(txt_NDoAlt)
End Sub

Private Sub cmb_TDoAlt_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TDoAlt_Click
   End If
End Sub

Private Sub cmb_TipAfp_Click()
   Call gs_SetFocus(cmb_Profes)
End Sub

Private Sub cmb_TipAfp_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipAfp_Click
   End If
End Sub

Private Sub cmb_TipVia_Click()
   Call gs_SetFocus(txt_NomVia)
End Sub

Private Sub cmb_TipVia_KeyPress(KeyAscii As Integer)
   Call cmb_TipVia_Click
End Sub

Private Sub cmb_TipZon_Click()
   Call gs_SetFocus(txt_NomZon)
End Sub

Private Sub cmb_TipZon_KeyPress(KeyAscii As Integer)
   Call cmb_TipZon_Click
End Sub

Private Sub cmd_ActEco_Click()
   moddat_g_str_NomCli = txt_ApePat.Text & " " & txt_ApeMat.Text & " " & txt_Nombre.Text
   modmip_g_int_PaiRes = CInt(l_arr_PaiRes(cmb_PaiRes.ListIndex + 1).Genera_Codigo)
   modmip_g_int_TipCli = 1
   
   frm_MntCli_56.Show 1
End Sub

Private Sub cmd_Cancel_Click()
   Call fs_Limpia
   Call fs_Cargar_Datos
   Call fs_Activa(False)
End Sub

Private Sub cmd_DatApo_Click()
   moddat_g_str_NomCli = txt_ApePat.Text & " " & txt_ApeMat.Text & " " & txt_Nombre.Text
   modmip_g_int_PaiRes = CInt(l_arr_PaiRes(cmb_PaiRes.ListIndex + 1).Genera_Codigo)
   
   If modmip_g_int_PaiRes = 4028 Then
      MsgBox "Sólo se puede ingresar información del Apoderado para clientes que tienen residencia en el extranjero.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   frm_MntCli_53.Show 1
End Sub

Private Sub cmd_DatCyg_Click()
   moddat_g_str_NomCli = txt_ApePat.Text & " " & txt_ApeMat.Text & " " & txt_Nombre.Text
   modmip_g_int_PaiRes = CInt(l_arr_PaiRes(cmb_PaiRes.ListIndex + 1).Genera_Codigo)
   
   If cmb_EstCiv.ItemData(cmb_EstCiv.ListIndex) = 1 Or cmb_EstCiv.ItemData(cmb_EstCiv.ListIndex) = 3 Or cmb_EstCiv.ItemData(cmb_EstCiv.ListIndex) = 4 Then
      MsgBox "El cliente no presenta Estado Civil Casado o Conviviente.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   frm_MntCli_54.Show 1
End Sub

Private Sub cmd_Editar_Click()
   moddat_g_int_FlgGrb = 2
   
   Call fs_Activa(True)
   
   If cmb_DocAlt.ItemData(cmb_DocAlt.ListIndex) = 2 Then
      cmb_TDoAlt.Enabled = False
      txt_NDoAlt.Enabled = False
   End If
   
   If l_arr_Paises(cmb_Paises.ListIndex + 1).Genera_Codigo <> "004028" Then
      cmb_DptNac.Enabled = False
      cmb_PrvNac.Enabled = False
      cmb_DstNac.Enabled = False
   End If
         
   If cmb_EstCiv.ItemData(cmb_EstCiv.ListIndex) <> 2 Then
      cmb_RegCyg.Enabled = False
   End If
   
   If l_arr_PaiRes(cmb_PaiRes.ListIndex + 1).Genera_Codigo <> "004028" Then
      cmb_TipVia.Enabled = False
      txt_NomVia.Enabled = False
      txt_NumVia.Enabled = False
      txt_IntDpt.Enabled = False
      cmb_TipZon.Enabled = False
      txt_NomZon.Enabled = False
      cmb_DptDir.Enabled = False
      cmb_PrvDir.Enabled = False
      cmb_DstDir.Enabled = False
      txt_Refere.Enabled = False
   Else
      txt_Direcc.Enabled = False
      cmb_PrvEst.Enabled = False
      txt_CodPos.Enabled = False
   End If
   
   Call gs_SetFocus(cmb_DocAlt)
End Sub

Private Sub cmd_Grabar_Click()
   Dim r_int_EdaMin     As Integer
   Dim r_int_EdaMax     As Integer
   Dim r_int_EdaAct     As Integer

   If cmb_DocAlt.ListIndex = -1 Then
      MsgBox "Debe seleccionar si el Cliente es miembro de las FF.AA o FF.PP.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_DocAlt)
      Exit Sub
   End If
   
   If cmb_DocAlt.ItemData(cmb_DocAlt.ListIndex) = 1 Then
      If cmb_TDoAlt.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Tipo de Documento de Identidad.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_TDoAlt)
         Exit Sub
      End If
      
      If Len(Trim(txt_NDoAlt.Text)) = 0 Then
         MsgBox "Debe ingresar el Número de Documento de Identidad.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NDoAlt)
         Exit Sub
      End If
   End If
   
   If Len(Trim(txt_ApePat.Text)) = 0 Then
      MsgBox "Debe ingresar el Apellido Paterno.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_ApePat)
      Exit Sub
   End If
   
   If Len(Trim(txt_ApeMat.Text)) = 0 Then
      MsgBox "Debe ingresar el Apellido Materno.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_ApeMat)
      Exit Sub
   End If
   
   If Len(Trim(txt_Nombre.Text)) = 0 Then
      MsgBox "Debe ingresar el Nombre.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Nombre)
      Exit Sub
   End If
   
   If cmb_CodSex.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Sexo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_CodSex)
      Exit Sub
   End If
   
   'Si es Masculino
   If cmb_CodSex.ItemData(cmb_CodSex.ListIndex) = 1 Then
      If Len(Trim(txt_ApeCas.Text)) > 0 Then
         MsgBox "El cliente no puede presentar Apellido de Casada.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_ApeCas)
         Exit Sub
      End If
   End If
   
   If Not IsDate(ipp_FecNac.Text) Then
      MsgBox "La Fecha de Nacimiento no es válida.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecNac)
      Exit Sub
   End If
   
   If CDate(ipp_FecNac.Text) > date Then
      MsgBox "Debe ingresar una Fecha de Nacimiento valida.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecNac)
      Exit Sub
   End If

   Call moddat_gs_FecSis
   
   'Rango de Edades del Cliente
   If Len(Trim(moddat_g_str_CodPrd)) > 0 Then
      If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera, moddat_g_str_CodPrd, moddat_g_str_CodSub, "001", "011") Then
         r_int_EdaMin = moddat_g_arr_Genera(1).Genera_ValMin
         r_int_EdaMax = moddat_g_arr_Genera(1).Genera_ValMax
      End If
      
      r_int_EdaAct = CInt(Left(gs_CalcularEdad(CDate(ipp_FecNac.Text), date), 2))
      
      If Not (r_int_EdaAct >= r_int_EdaMin And r_int_EdaAct <= r_int_EdaMax) Then
         MsgBox "El Cliente no cumple con los requisitos de Edad requeridos. Tiene " & CStr(r_int_EdaAct) & " años.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_FecNac)
         Exit Sub
      End If
   End If
   
   If Len(Trim(txt_DirEle.Text)) = 0 Then
      MsgBox "Debe ingresar el E-mail del cliente.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_DirEle)
      Exit Sub
   End If
   If Not gf_ValidarEmail(txt_DirEle.Text) Then
      MsgBox "El E-mail del cliente no tiene el formato correcto.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_DirEle)
      Exit Sub
   End If
   
   If cmb_Paises.ListIndex = -1 Then
      MsgBox "Debe seleccionar el País de Nacimiento.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Paises)
      Exit Sub
   End If
   
   If CInt(l_arr_Paises(cmb_Paises.ListIndex + 1).Genera_Codigo) = 4028 Then
      If cmb_DptNac.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Departamento de Nacimiento.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_DptNac)
         Exit Sub
      End If
      
      If cmb_PrvNac.ListIndex = -1 Then
         MsgBox "Debe seleccionar la Provincia de Nacimiento.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_PrvNac)
         Exit Sub
      End If
      
      If cmb_DstNac.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Distrito de Nacimiento.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_DstNac)
         Exit Sub
      End If
   End If
   
   If cmb_EstCiv.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Estado Civil.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_EstCiv)
      Exit Sub
   End If
   
   If cmb_EstCiv.ItemData(cmb_EstCiv.ListIndex) = 2 Then
      If cmb_RegCyg.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Régimen Conyugal.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_RegCyg)
         Exit Sub
      End If
   End If
   
   If cmb_NivEst.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Nivel de Estudio.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_NivEst)
      Exit Sub
   End If
   
   If cmb_Profes.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Profesión u Oficio.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Profes)
      Exit Sub
   End If
   
   If cmb_PaiRes.ListIndex = -1 Then
      MsgBox "Debe seleccionar el País de residencia.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PaiRes)
      Exit Sub
   End If
   
   If CInt(l_arr_PaiRes(cmb_PaiRes.ListIndex + 1).Genera_Codigo) = 4028 Then
      If cmb_TipVia.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Tipo de Vía.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_TipVia)
         Exit Sub
      End If
      
      If Len(Trim(txt_NomVia.Text)) = 0 Then
         MsgBox "Debe ingresar el Nombre de Vía.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NomVia)
         Exit Sub
      End If
      
      If Len(Trim(txt_NumVia.Text)) = 0 Then
         MsgBox "Debe ingresar el Número.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NumVia)
         Exit Sub
      End If
      
      If cmb_TipZon.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Tipo de Zona.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_TipZon)
         Exit Sub
      End If
      
      If cmb_TipZon.ItemData(cmb_TipZon.ListIndex) <> 12 Then
         If Len(Trim(txt_NomZon.Text)) = 0 Then
            MsgBox "Debe ingresar el Nombre de Zona.", vbExclamation, modgen_g_str_NomPlt
            Call gs_SetFocus(txt_NomZon)
            Exit Sub
         End If
      End If
      
      If cmb_DptDir.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Departamento de la Dirección.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_DptDir)
         Exit Sub
      End If
      
      If cmb_PrvDir.ListIndex = -1 Then
         MsgBox "Debe seleccionar la Provincia de la Dirección.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_PrvDir)
         Exit Sub
      End If
      
      If cmb_DstDir.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Distrito de la Dirección.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_DstDir)
         Exit Sub
      End If
   Else
      If Len(Trim(txt_Direcc.Text)) = 0 Then
         MsgBox "Debe ingresar la Dirección.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_Direcc)
         Exit Sub
      End If
      
      If cmb_PrvEst.ListIndex = -1 Then
         MsgBox "Debe seleccionar la Provincia / Estado.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_PrvEst)
         Exit Sub
      End If
      
      If Len(Trim(txt_CodPos.Text)) = 0 Then
         MsgBox "Debe ingresar el Código Postal.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_CodPos)
         Exit Sub
      End If
   End If
   
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   'Grabando Información del Cliente
   If moddat_g_int_FlgGrb = 1 Then
      g_str_Parame = "USP_CLI_DATGEN_NUEVO ("
      g_str_Parame = g_str_Parame & "1, "
   Else
      g_str_Parame = "USP_CLI_DATGEN_MODIFICA ("
   End If
   
   g_str_Parame = g_str_Parame & CStr(moddat_g_int_TipDoc) & ", "
   g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumDoc & "', "
   g_str_Parame = g_str_Parame & CStr(cmb_DocAlt.ItemData(cmb_DocAlt.ListIndex)) & ", "
   
   If cmb_DocAlt.ItemData(cmb_DocAlt.ListIndex) = 1 Then
      g_str_Parame = g_str_Parame & CStr(cmb_TDoAlt.ItemData(cmb_TDoAlt.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & "'" & txt_NDoAlt.Text & "', "
   Else
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "'', "
   End If
   
   g_str_Parame = g_str_Parame & "'" & txt_ApePat & "', "
   g_str_Parame = g_str_Parame & "'" & txt_ApeMat & "', "
   g_str_Parame = g_str_Parame & "'" & txt_ApeCas & "', "
   g_str_Parame = g_str_Parame & "'" & txt_Nombre & "', "
   g_str_Parame = g_str_Parame & CStr(cmb_EstCiv.ItemData(cmb_EstCiv.ListIndex)) & ", "
   
   If cmb_EstCiv.ItemData(cmb_EstCiv.ListIndex) = 2 Then
      g_str_Parame = g_str_Parame & CStr(cmb_RegCyg.ItemData(cmb_RegCyg.ListIndex)) & ", "
   Else
      g_str_Parame = g_str_Parame & "0, "
   End If
   
   g_str_Parame = g_str_Parame & CStr(cmb_NivEst.ItemData(cmb_NivEst.ListIndex)) & ", "
   g_str_Parame = g_str_Parame & "'" & l_arr_Profes(cmb_Profes.ListIndex + 1).Genera_Codigo & "', "
   g_str_Parame = g_str_Parame & CStr(cmb_CodSex.ItemData(cmb_CodSex.ListIndex)) & ", "
   g_str_Parame = g_str_Parame & Format(CDate(ipp_FecNac.Text), "yyyymmdd") & ", "
   
   If CInt(l_arr_Paises(cmb_Paises.ListIndex + 1).Genera_Codigo) = 4028 Then
      g_str_Parame = g_str_Parame & "'" & Format(cmb_DptNac.ItemData(cmb_DptNac.ListIndex), "00") & Format(cmb_PrvNac.ItemData(cmb_PrvNac.ListIndex), "00") & Format(cmb_DstNac.ItemData(cmb_DstNac.ListIndex), "00") & "', "
   Else
      g_str_Parame = g_str_Parame & "'000000', "
   End If
   g_str_Parame = g_str_Parame & "'" & l_arr_Paises(cmb_Paises.ListIndex + 1).Genera_Codigo & "', "
   
   g_str_Parame = g_str_Parame & CStr(ipp_NumDep.Value) & ", "
   g_str_Parame = g_str_Parame & CStr(ipp_DepEc1.Value) & ", "
   g_str_Parame = g_str_Parame & CStr(ipp_DepEc2.Value) & ", "
   g_str_Parame = g_str_Parame & CStr(ipp_DepEc3.Value) & ", "
   g_str_Parame = g_str_Parame & CStr(ipp_DepEc4.Value) & ", "
   g_str_Parame = g_str_Parame & CStr(ipp_DepEc5.Value) & ", "
   
   If CInt(l_arr_PaiRes(cmb_PaiRes.ListIndex + 1).Genera_Codigo) = 4028 Then
      g_str_Parame = g_str_Parame & CStr(cmb_TipVia.ItemData(cmb_TipVia.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & "'" & txt_NomVia.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_NumVia.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_IntDpt.Text & "', "
      g_str_Parame = g_str_Parame & CStr(cmb_TipZon.ItemData(cmb_TipZon.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & "'" & txt_NomZon.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_Refere.Text & "', "
      g_str_Parame = g_str_Parame & "'" & Format(cmb_DptDir.ItemData(cmb_DptDir.ListIndex), "00") & Format(cmb_PrvDir.ItemData(cmb_PrvDir.ListIndex), "00") & Format(cmb_DstDir.ItemData(cmb_DstDir.ListIndex), "00") & "', "
   Else
      If l_int_TipVia > 0 Then
         g_str_Parame = g_str_Parame & CStr(l_int_TipVia) & ", "
         g_str_Parame = g_str_Parame & "'" & l_str_NomVia & "', "
         g_str_Parame = g_str_Parame & "'" & l_str_NumVia & "', "
         g_str_Parame = g_str_Parame & "'" & l_str_IntDpt & "', "
         g_str_Parame = g_str_Parame & CStr(l_int_TipZon) & ", "
         g_str_Parame = g_str_Parame & "'" & l_str_NomZon & "', "
         g_str_Parame = g_str_Parame & "'" & l_str_Refere & "', "
         g_str_Parame = g_str_Parame & "'" & l_str_UbiGeo & "', "
      Else
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "0, "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'', "
         g_str_Parame = g_str_Parame & "'', "
      End If
   End If
   
   g_str_Parame = g_str_Parame & "'" & txt_Celula.Text & "', "
   g_str_Parame = g_str_Parame & "'" & txt_TelFij.Text & "', "
   
   g_str_Parame = g_str_Parame & "'" & txt_DirEle.Text & "', "
   
   If chk_DirEle.Value = 1 Then
      g_str_Parame = g_str_Parame & "1, "
   ElseIf chk_DirEle.Value = 0 Then
      g_str_Parame = g_str_Parame & "2, "
   End If
   
   If CInt(l_arr_Paises(cmb_Paises.ListIndex + 1).Genera_Codigo) = 4028 Then
      g_str_Parame = g_str_Parame & "'1', "
   Else
      g_str_Parame = g_str_Parame & "'0', "
   End If
   
   g_str_Parame = g_str_Parame & "'" & CStr(l_int_RelAcc) & "', "
   g_str_Parame = g_str_Parame & "'" & CStr(l_int_RelLab) & "', "
   
   g_str_Parame = g_str_Parame & "'" & l_arr_PaiRes(cmb_PaiRes.ListIndex + 1).Genera_Codigo & "', "
   
   If CInt(l_arr_PaiRes(cmb_PaiRes.ListIndex + 1).Genera_Codigo) <> 4028 Then
      g_str_Parame = g_str_Parame & "'" & txt_Direcc.Text & "', "
      g_str_Parame = g_str_Parame & "'" & l_arr_PrvEst(cmb_PrvEst.ListIndex + 1).Genera_Codigo & "', "
      g_str_Parame = g_str_Parame & "'" & txt_CodPos.Text & "', "
   Else
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "'', "
      g_str_Parame = g_str_Parame & "'', "
   End If
   
   g_str_Parame = g_str_Parame & CStr(l_int_VinTDo) & ", "
   g_str_Parame = g_str_Parame & "'" & l_str_VinNDo & "', "
   g_str_Parame = g_str_Parame & CStr(l_int_VinTip) & ", "
   g_str_Parame = g_str_Parame & CStr(l_int_AccTDo) & ", "
   g_str_Parame = g_str_Parame & "'" & l_str_AccNDo & "', "
   g_str_Parame = g_str_Parame & CStr(l_int_AccVin) & ", "
   g_str_Parame = g_str_Parame & "1, "
   
   If (cmb_TipAfp.ListIndex = -1) Then
       g_str_Parame = g_str_Parame & "NULL, "
   Else
       g_str_Parame = g_str_Parame & CStr(cmb_TipAfp.ItemData(cmb_TipAfp.ListIndex)) & ", "
   End If

   'Datos de Auditoria
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
      If moddat_g_int_FlgGrb = 1 Then
         MsgBox "Error al ejecutar el Procedimiento USP_CLI_DATGEN_NUEVO.", vbCritical, modgen_g_str_NomPlt
         Exit Sub
      Else
         MsgBox "Error al ejecutar el Procedimiento USP_CLI_DATGEN_MODIFICA.", vbCritical, modgen_g_str_NomPlt
         Exit Sub
      End If
   End If
   
   moddat_g_int_FlgAct = 2
   
   MsgBox "Los datos se grabaron correctamente.", vbInformation, modgen_g_str_NomPlt
   
   Call fs_Activa(False)
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Dim r_str_NomVin     As String
   Dim r_str_NomAcc     As String
   
   Screen.MousePointer = 11
   Call gs_CentraForm(Me)
   Me.Caption = modgen_g_str_NomPlt
   pnl_DocIde.Caption = moddat_g_str_TipDoc & " - " & moddat_g_str_NumDoc
   pnl_RelLab.Visible = False
   pnl_RelAcc.Visible = False
   
   Call fs_Inicio
   
   If moddat_g_int_FlgGrb = 1 Then
      Call fs_Activa(True)
      cmd_Cancel.Enabled = False
      Call fs_Limpia
   Else
      Call fs_Limpia
      Call fs_Cargar_Datos
      Call fs_Activa(False)
   End If
   
   'Verificando Relación Laboral
   Call modmip_gs_RelLab(moddat_g_int_TipDoc, moddat_g_str_NumDoc, l_int_RelLab, l_int_VinTDo, l_str_VinNDo, l_int_VinTip)
   
   If l_int_RelLab > 0 Then
      pnl_RelLab.Visible = True
      
      If l_int_VinTip = 1 Then
         pnl_RelLab.Caption = "Cliente es Trabajador de miCasita"
      ElseIf l_int_VinTip = 2 Or l_int_VinTip = 3 Then
         pnl_RelLab.Caption = "Cliente Vinculado (" & modmip_gf_Consulta_NomTra(l_int_VinTDo, l_str_VinNDo) & ")"
      ElseIf l_int_VinTip = 4 Then
         pnl_RelLab.Caption = "Cliente es Funcionario de miCasita"
      ElseIf l_int_VinTip = 5 Then
         pnl_RelLab.Caption = "Cliente Vinculado (" & modmip_gf_Consulta_NomOtrFun(l_int_VinTDo, l_str_VinNDo) & ")"
      Else
         pnl_RelLab.Caption = ""
      End If
   End If
   
   'Verificando Relación con Accionistas
   Call modmip_gs_RelAcc(moddat_g_int_TipDoc, moddat_g_str_NumDoc, l_int_RelAcc, l_int_AccTDo, l_str_AccNDo, l_int_AccVin)

   If l_int_RelAcc > 0 Then
      pnl_RelAcc.Visible = True
      
      If l_int_AccVin = 1 Then
         pnl_RelAcc.Caption = "Cliente es Accionista"
      ElseIf l_int_AccVin = 2 Then
         pnl_RelAcc.Caption = "Relación con Accionista (" & modmip_gf_Consulta_NomAcc(l_int_AccTDo, l_str_AccNDo) & ")"
      End If
   End If
   
   Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Si registro es nuevo solicita confirmacion de salir
   If moddat_g_int_FlgGrb = 1 Then
      If MsgBox("¿Está seguro de salir sin grabar?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) = vbYes Then
         Exit Sub
      End If
   End If
   
   'Verificar que las Actividades Económicas se hayan ingresado
   g_str_Parame = "SELECT * FROM CLI_ACTECO WHERE "
   g_str_Parame = g_str_Parame & "ACTECO_CLITDO = " & CStr(moddat_g_int_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "ACTECO_CLINDO = '" & moddat_g_str_NumDoc & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
      Exit Sub
   End If
   
   If g_rst_Listas.BOF And g_rst_Listas.EOF Then
      MsgBox "No se encuentran registradas las Actividades Económicas del Cliente.", vbExclamation, modgen_g_str_NomPlt
      Cancel = True
   Else
      g_rst_Listas.Close
      Set g_rst_Listas = Nothing
   End If
   
   'Verificar que el Cónyuge se haya ingresado (Casado o Conviviente)
   If cmb_EstCiv.ListIndex > -1 Then
      If cmb_EstCiv.ItemData(cmb_EstCiv.ListIndex) = 2 Or cmb_EstCiv.ItemData(cmb_EstCiv.ListIndex) = 5 And Cancel = False Then
         If moddat_g_int_CygTDo <> 0 And moddat_g_str_CygNDo <> "" Then
            g_str_Parame = "SELECT * FROM CLI_DATGEN WHERE "
            g_str_Parame = g_str_Parame & "DATGEN_TIPDOC= " & CStr(moddat_g_int_CygTDo) & " AND "
            g_str_Parame = g_str_Parame & "DATGEN_NUMDOC = '" & moddat_g_str_CygNDo & "' "
            
            If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
               Exit Sub
            End If
            
            If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
            
               'Verificar que las Actividades Económicas del Cónyuge se hayan ingresado (Si Trabaja) cli_acteco
               If g_rst_Listas!DATGEN_ACTECO = 1 Then
                  g_str_Parame = "SELECT * FROM CLI_ACTECO WHERE "
                  g_str_Parame = g_str_Parame & "ACTECO_CLITDO = " & CStr(moddat_g_int_CygTDo) & " AND "
                  g_str_Parame = g_str_Parame & "ACTECO_CLINDO = '" & moddat_g_str_CygNDo & "' "
                  
                  If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
                     Exit Sub
                  End If
                  
                  If g_rst_Genera.BOF And g_rst_Genera.EOF Then
                     MsgBox "No se encuentran registradas las Actividades Económicas del Cónyuge.", vbExclamation, modgen_g_str_NomPlt
                     Cancel = True
                  Else
                     g_rst_Genera.Close
                     Set g_rst_Genera = Nothing
                  End If
               
               End If
            
            End If
            g_rst_Listas.Close
            Set g_rst_Listas = Nothing
         
         Else
            MsgBox "No se encuentra registrado el Cónyuge o Conviviente.", vbExclamation, modgen_g_str_NomPlt
            Cancel = True
         End If
      End If
   End If
   
   'Verificar que la información del Apoderado se haya ingresado
   If cmb_PaiRes.ListIndex > -1 Then
      modmip_g_int_PaiRes = CInt(l_arr_PaiRes(cmb_PaiRes.ListIndex + 1).Genera_Codigo)
      
      If modmip_g_int_PaiRes <> 4028 Then
         g_str_Parame = "SELECT * FROM CLI_DATGEN WHERE "
         g_str_Parame = g_str_Parame & "DATGEN_TIPDOC= " & CStr(moddat_g_int_TipDoc) & " AND "
         g_str_Parame = g_str_Parame & "DATGEN_NUMDOC = '" & moddat_g_str_NumDoc & "' "
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Listas, 3) Then
            Exit Sub
         End If
         
         If Not (g_rst_Listas.BOF And g_rst_Listas.EOF) Then
            If g_rst_Listas!DATGEN_APOTDO = 0 Or IIf(IsNull(g_rst_Listas!DATGEN_APONDO) = True, "", Trim(g_rst_Listas!DATGEN_APONDO)) = "" Then
               MsgBox "No se encuentra registrado el Apoderado.", vbExclamation, modgen_g_str_NomPlt
               Cancel = True
            End If
         End If
         
         g_rst_Listas.Close
         Set g_rst_Listas = Nothing
      End If
   End If
End Sub

Private Sub ipp_DepEc1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_DepEc2)
   End If
End Sub

Private Sub ipp_DepEc2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_DepEc3)
   End If
End Sub

Private Sub ipp_DepEc3_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_DepEc4)
   End If
End Sub

Private Sub ipp_DepEc4_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_DepEc5)
   End If
End Sub

Private Sub ipp_DepEc5_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_PaiRes)
   End If
End Sub

Private Sub ipp_FecNac_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Paises)
   End If
End Sub

Private Sub ipp_NumDep_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_DepEc1)
   End If
End Sub

Private Sub txt_ApePat_GotFocus()
   Call gs_SelecTodo(txt_ApePat)
End Sub

Private Sub txt_ApePat_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_ApeMat)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "- '")
   End If
End Sub

Private Sub txt_ApeMat_GotFocus()
   Call gs_SelecTodo(txt_ApeMat)
End Sub

Private Sub txt_ApeMat_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_ApeCas)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "- '")
   End If
End Sub

Private Sub txt_ApeCas_GotFocus()
   Call gs_SelecTodo(txt_ApeCas)
End Sub

Private Sub txt_ApeCas_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Nombre)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "- '")
   End If
End Sub

Private Sub txt_Celula_GotFocus()
   Call gs_SelecTodo(txt_Celula)
End Sub

Private Sub txt_Celula_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_DirEle)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_CodPos_GotFocus()
   Call gs_SelecTodo(txt_CodPos)
End Sub

Private Sub txt_CodPos_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-( )%$.,;:@_?¿º")
   End If
End Sub

Private Sub txt_Direcc_GotFocus()
   Call gs_SelecTodo(txt_Direcc)
End Sub

Private Sub txt_Direcc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_PrvEst)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-( )%$.,;:@_?¿º")
   End If
End Sub

Private Sub txt_DirEle_Change()
   If Len(Trim(txt_DirEle)) > 0 Then
      chk_DirEle.Enabled = True
   Else
      chk_DirEle.Value = 0
      chk_DirEle.Enabled = False
   End If
End Sub

Private Sub txt_DirEle_GotFocus()
   Call gs_SelecTodo(txt_DirEle)
End Sub

Private Sub txt_DirEle_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_NumDep)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-@_.")
   End If
End Sub

Private Sub txt_IntDpt_GotFocus()
   Call gs_SelecTodo(txt_IntDpt)
End Sub

Private Sub txt_IntDpt_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipZon)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_NDoAlt_GotFocus()
   Call gs_SelecTodo(txt_NDoAlt)
End Sub

Private Sub txt_NDoAlt_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_ApePat)
   Else
      If cmb_TDoAlt.ListIndex > -1 Then
         Select Case cmb_TDoAlt.ItemData(cmb_TDoAlt.ListIndex)
            Case 1:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
            Case 2:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-")
            Case 3:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-")
            Case 4:  KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-")
         End Select
      Else
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txt_Nombre_GotFocus()
   Call gs_SelecTodo(txt_Nombre)
End Sub

Private Sub txt_Nombre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_CodSex)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "- '")
   End If
End Sub

Private Sub txt_NomVia_GotFocus()
   Call gs_SelecTodo(txt_NomVia)
End Sub

Private Sub txt_NomVia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NumVia)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_NomZon_GotFocus()
   Call gs_SelecTodo(txt_NomZon)
End Sub

Private Sub txt_NomZon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_DptDir)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_NumVia_GotFocus()
   Call gs_SelecTodo(txt_NumVia)
End Sub

Private Sub txt_NumVia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_IntDpt)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_Refere_GotFocus()
   Call gs_SelecTodo(txt_Refere)
End Sub

Private Sub txt_Refere_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-( )%$.,;:@_?¿º")
   End If
End Sub

Private Sub txt_TelFij_GotFocus()
   Call gs_SelecTodo(txt_TelFij)
End Sub

Private Sub txt_TelFij_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmb_TipVia.Enabled Then
         Call gs_SetFocus(cmb_TipVia)
      Else
         Call gs_SetFocus(txt_Direcc)
      End If
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub fs_Inicio()
   Call moddat_gs_Carga_LisIte_Combo(cmb_DocAlt, 1, "214")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TDoAlt, 1, "231")
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_CodSex, 1, "207")
   Call moddat_gs_Carga_LisIte_Combo(cmb_EstCiv, 1, "205")
   Call moddat_gs_Carga_LisIte_Combo(cmb_RegCyg, 1, "206")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipVia, 1, "201")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipZon, 1, "202")
   Call moddat_gs_Carga_LisIte_Combo(cmb_NivEst, 1, "209")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipAfp, 1, "517")
      
   Call moddat_gs_Carga_LisIte(cmb_Paises, l_arr_Paises, 1, "500")
   Call moddat_gs_Carga_LisIte(cmb_PaiRes, l_arr_PaiRes, 1, "500")
   Call moddat_gs_Carga_LisIte(cmb_Profes, l_arr_Profes, 1, "501")
      
   Call moddat_gs_Carga_Depart(cmb_DptNac)
   Call moddat_gs_Carga_Depart(cmb_DptDir)
End Sub
   
Private Sub fs_Activa(ByVal p_Habilita As Integer)
   cmb_DocAlt.Enabled = p_Habilita
   cmb_TDoAlt.Enabled = p_Habilita
   txt_NDoAlt.Enabled = p_Habilita
   
   txt_ApePat.Enabled = p_Habilita
   txt_ApeMat.Enabled = p_Habilita
   txt_ApeCas.Enabled = p_Habilita
   txt_Nombre.Enabled = p_Habilita
   cmb_CodSex.Enabled = p_Habilita
   ipp_FecNac.Enabled = p_Habilita
   cmb_Paises.Enabled = p_Habilita
   cmb_DptNac.Enabled = p_Habilita
   cmb_PrvNac.Enabled = p_Habilita
   cmb_DstNac.Enabled = p_Habilita
   cmb_EstCiv.Enabled = p_Habilita
   cmb_RegCyg.Enabled = p_Habilita
   cmb_NivEst.Enabled = p_Habilita
   cmb_TipAfp.Enabled = p_Habilita
   cmb_Profes.Enabled = p_Habilita
   txt_Celula.Enabled = p_Habilita
   txt_DirEle.Enabled = p_Habilita
   chk_DirEle.Enabled = p_Habilita
   ipp_NumDep.Enabled = p_Habilita
   ipp_DepEc1.Enabled = p_Habilita
   ipp_DepEc2.Enabled = p_Habilita
   ipp_DepEc3.Enabled = p_Habilita
   ipp_DepEc4.Enabled = p_Habilita
   ipp_DepEc5.Enabled = p_Habilita
   cmb_PaiRes.Enabled = p_Habilita
   txt_TelFij.Enabled = p_Habilita
   cmb_TipVia.Enabled = p_Habilita
   txt_NomVia.Enabled = p_Habilita
   txt_NumVia.Enabled = p_Habilita
   txt_IntDpt.Enabled = p_Habilita
   cmb_TipZon.Enabled = p_Habilita
   txt_NomZon.Enabled = p_Habilita
   cmb_DptDir.Enabled = p_Habilita
   cmb_PrvDir.Enabled = p_Habilita
   cmb_DstDir.Enabled = p_Habilita
   txt_Refere.Enabled = p_Habilita
   txt_Direcc.Enabled = p_Habilita
   cmb_PrvEst.Enabled = p_Habilita
   txt_CodPos.Enabled = p_Habilita
   
   cmd_Grabar.Enabled = p_Habilita
   cmd_Cancel.Enabled = p_Habilita
   
   cmd_Editar.Enabled = Not p_Habilita
   cmd_ActEco.Enabled = Not p_Habilita
   cmd_DatCyg.Enabled = Not p_Habilita
   cmd_DatApo.Enabled = Not p_Habilita
End Sub

Private Sub fs_Limpia()
   cmb_DocAlt.ListIndex = -1
   cmb_TDoAlt.ListIndex = -1
   txt_NDoAlt.Text = ""
   
   cmb_TDoAlt.Enabled = False
   txt_NDoAlt.Enabled = False
   
   txt_ApePat.Text = ""
   txt_ApeMat.Text = ""
   txt_ApeCas.Text = ""
   txt_Nombre.Text = ""
   
   cmb_CodSex.ListIndex = -1
   ipp_FecNac.Text = Format(date, "dd/mm/yyyy")
   cmb_Paises.ListIndex = -1
   cmb_DptNac.ListIndex = -1
   cmb_PrvNac.Clear
   cmb_DstNac.Clear
   cmb_DptNac.Enabled = False
   cmb_PrvNac.Enabled = False
   cmb_DstNac.Enabled = False
   cmb_EstCiv.ListIndex = -1
   cmb_RegCyg.ListIndex = -1
   cmb_RegCyg.Enabled = False
   cmb_NivEst.ListIndex = -1
   cmb_TipAfp.ListIndex = -1
   cmb_Profes.ListIndex = -1
   txt_Celula.Text = ""
   txt_DirEle.Text = ""
   chk_DirEle.Value = 0
   chk_DirEle.Enabled = False
   
   ipp_NumDep.Value = 0
   ipp_DepEc1.Value = 0
   ipp_DepEc2.Value = 0
   ipp_DepEc3.Value = 0
   ipp_DepEc4.Value = 0
   ipp_DepEc5.Value = 0
   
   cmb_PaiRes.ListIndex = -1
   txt_TelFij.Text = ""
   
   cmb_TipVia.ListIndex = -1
   txt_NomVia.Text = ""
   txt_NumVia.Text = ""
   txt_IntDpt.Text = ""
   cmb_TipZon.ListIndex = -1
   txt_NomZon.Text = ""
   cmb_DptDir.ListIndex = -1
   cmb_PrvDir.Clear
   cmb_DstDir.Clear
   txt_Refere.Text = ""
   cmb_TipVia.Enabled = False
   txt_NomVia.Enabled = False
   txt_NumVia.Enabled = False
   txt_IntDpt.Enabled = False
   cmb_TipZon.Enabled = False
   txt_NomZon.Enabled = False
   cmb_DptDir.Enabled = False
   cmb_PrvDir.Enabled = False
   cmb_DstDir.Enabled = False
   txt_Refere.Enabled = False
   
   txt_Direcc.Text = ""
   cmb_PrvEst.Clear
   txt_CodPos.Text = ""
   
   txt_Direcc.Enabled = False
   cmb_PrvEst.Enabled = False
   txt_CodPos.Enabled = False
End Sub

Private Sub fs_Cargar_Datos()
   l_int_TipVia = 0
   l_str_NomVia = ""
   l_str_NumVia = ""
   l_str_IntDpt = ""
   l_int_TipZon = 0
   l_str_NomZon = ""
   l_str_UbiGeo = ""
   l_str_Refere = ""
   
   g_str_Parame = "SELECT * FROM CLI_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_TIPDOC = " & CStr(moddat_g_int_TipDoc) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_NUMDOC = '" & moddat_g_str_NumDoc & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      Call gs_BuscarCombo_Item(cmb_DocAlt, g_rst_Princi!DatGen_FLGDOA)
      
      If cmb_DocAlt.ItemData(cmb_DocAlt.ListIndex) = 1 Then
         Call gs_BuscarCombo_Item(cmb_TDoAlt, g_rst_Princi!DatGen_TIPDOA)
         txt_NDoAlt.Text = Trim(g_rst_Princi!DatGen_NUMDOA)
         
         cmb_TDoAlt.Enabled = True
         txt_NDoAlt.Enabled = True
      End If
      
      txt_ApePat.Text = Trim(g_rst_Princi!DATGEN_APEPAT & "")
      txt_ApeMat.Text = Trim(g_rst_Princi!DATGEN_APEMAT & "")
      txt_ApeCas.Text = Trim(g_rst_Princi!DatGen_ApeCas & "")
      txt_Nombre.Text = Trim(g_rst_Princi!DATGEN_NOMBRE & "")
      
      Call gs_BuscarCombo_Item(cmb_CodSex, g_rst_Princi!DatGen_CodSex)
      ipp_FecNac.Text = Right(CStr(g_rst_Princi!DATGEN_NACFEC), 2) & "/" & Mid(CStr(g_rst_Princi!DATGEN_NACFEC), 5, 2) & "/" & Left(CStr(g_rst_Princi!DATGEN_NACFEC), 4)
      
      cmb_Paises.ListIndex = gf_Busca_Arregl(l_arr_Paises, g_rst_Princi!DATGEN_NACPAI) - 1
      
      If l_arr_Paises(cmb_Paises.ListIndex + 1).Genera_Codigo = "004028" Then
         Call gs_BuscarCombo_Item(cmb_DptNac, CInt(Left(g_rst_Princi!DATGEN_NACLUG, 2)))
         Call moddat_gs_Carga_Provin(cmb_PrvNac, Left(g_rst_Princi!DATGEN_NACLUG, 2))
         Call gs_BuscarCombo_Item(cmb_PrvNac, CInt(Mid(g_rst_Princi!DATGEN_NACLUG, 3, 2)))
         Call moddat_gs_Carga_Distri(cmb_DstNac, Left(g_rst_Princi!DATGEN_NACLUG, 2), Mid(g_rst_Princi!DATGEN_NACLUG, 3, 2))
         Call gs_BuscarCombo_Item(cmb_DstNac, CInt(Right(g_rst_Princi!DATGEN_NACLUG, 2)))
         
         cmb_DptNac.Enabled = True
         cmb_PrvNac.Enabled = True
         cmb_DstNac.Enabled = True
      End If
      
      Call gs_BuscarCombo_Item(cmb_EstCiv, g_rst_Princi!DATGEN_ESTCIV)
      
      If cmb_EstCiv.ItemData(cmb_EstCiv.ListIndex) = 2 Then
         Call gs_BuscarCombo_Item(cmb_RegCyg, g_rst_Princi!DATGEN_REGCYG)
         
         cmb_RegCyg.Enabled = True
      End If
      
      Call gs_BuscarCombo_Item(cmb_NivEst, g_rst_Princi!DatGen_NivEst)
      If (Not IsNull(g_rst_Princi!DATGEN_TIPAFP)) Then
          Call gs_BuscarCombo_Item(cmb_TipAfp, g_rst_Princi!DATGEN_TIPAFP)
      End If
      cmb_Profes.ListIndex = gf_Busca_Arregl(l_arr_Profes, g_rst_Princi!DatGen_Profes) - 1
      
      txt_Celula.Text = Trim(g_rst_Princi!DATGEN_NUMCEL & "")
      txt_DirEle.Text = Trim(g_rst_Princi!DatGen_DirEle & "")
      
      If g_rst_Princi!DATGEN_AUTENV = 1 Then
         chk_DirEle.Value = 1
         chk_DirEle.Enabled = True
      End If
      
      ipp_NumDep.Value = g_rst_Princi!DatGen_DepEco
      
      ipp_DepEc1.Value = g_rst_Princi!DatGen_EDAD01
      ipp_DepEc2.Value = g_rst_Princi!DatGen_EDAD02
      ipp_DepEc3.Value = g_rst_Princi!DatGen_EDAD03
      ipp_DepEc4.Value = g_rst_Princi!DatGen_EDAD04
      ipp_DepEc5.Value = g_rst_Princi!DatGen_EDAD05
      
      cmb_PaiRes.ListIndex = gf_Busca_Arregl(l_arr_PaiRes, g_rst_Princi!DATGEN_PAIRES) - 1
      txt_TelFij.Text = Trim(g_rst_Princi!DatGen_Telefo & "")
      
      If l_arr_PaiRes(cmb_PaiRes.ListIndex + 1).Genera_Codigo = "004028" Then
         Call gs_BuscarCombo_Item(cmb_TipVia, g_rst_Princi!DatGen_TipVia)
         txt_NomVia.Text = Trim(g_rst_Princi!DatGen_NomVia & "")
         txt_NumVia.Text = Trim(g_rst_Princi!DatGen_Numero & "")
         txt_IntDpt.Text = Trim(g_rst_Princi!DATGEN_INTDPT & "")
         
         Call gs_BuscarCombo_Item(cmb_TipZon, g_rst_Princi!DatGen_TipZon)
         txt_NomZon.Text = Trim(g_rst_Princi!DatGen_NomZon & "")
      
         Call gs_BuscarCombo_Item(cmb_DptDir, CInt(Left(g_rst_Princi!DatGen_Ubigeo, 2)))
         Call moddat_gs_Carga_Provin(cmb_PrvDir, Left(g_rst_Princi!DatGen_Ubigeo, 2))
         Call gs_BuscarCombo_Item(cmb_PrvDir, CInt(Mid(g_rst_Princi!DatGen_Ubigeo, 3, 2)))
         Call moddat_gs_Carga_Distri(cmb_DstDir, Left(g_rst_Princi!DatGen_Ubigeo, 2), Mid(g_rst_Princi!DatGen_Ubigeo, 3, 2))
         Call gs_BuscarCombo_Item(cmb_DstDir, CInt(Right(g_rst_Princi!DatGen_Ubigeo, 2)))
         
         txt_Refere.Text = Trim(g_rst_Princi!DATGEN_REFERE & "")
      Else
         txt_Direcc.Text = Trim(g_rst_Princi!DATGEN_EXTDIR & "")
         
         Call modmip_gs_Carga_CiuExt(l_arr_PrvEst, cmb_PrvEst, l_arr_PaiRes(cmb_PaiRes.ListIndex + 1).Genera_Codigo)
         cmb_PrvEst.ListIndex = gf_Busca_Arregl(l_arr_PrvEst, g_rst_Princi!DATGEN_EXTCIU) - 1
         
         txt_CodPos.Text = Trim(g_rst_Princi!DATGEN_EXTCPO & "")
         
         If g_rst_Princi!DATGEN_APOTDO > 0 Then
            l_int_TipVia = g_rst_Princi!DatGen_TipVia
            l_str_NomVia = Trim(g_rst_Princi!DatGen_NomVia & "")
            l_str_NumVia = Trim(g_rst_Princi!DatGen_Numero & "")
            l_str_IntDpt = Trim(g_rst_Princi!DATGEN_INTDPT & "")
            
            l_int_TipZon = g_rst_Princi!DatGen_TipZon
            l_str_NomZon = Trim(g_rst_Princi!DatGen_NomZon & "")
         
            l_str_UbiGeo = Trim(g_rst_Princi!DatGen_Ubigeo & "")
            l_str_Refere = Trim(g_rst_Princi!DATGEN_REFERE & "")
         End If
      End If
      
      moddat_g_int_CygTDo = g_rst_Princi!DATGEN_CYGTDO
      moddat_g_str_CygNDo = Trim(g_rst_Princi!DATGEN_CYGNDO & "")
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub


