VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_Cob_GasAdm_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   8535
   ClientLeft      =   1965
   ClientTop       =   1170
   ClientWidth     =   11820
   Icon            =   "OpeTra_frm_003.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   11820
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel2 
      Height          =   8535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11835
      _Version        =   65536
      _ExtentX        =   20876
      _ExtentY        =   15055
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
      Begin Threed.SSPanel SSPanel5 
         Height          =   1755
         Left            =   30
         TabIndex        =   1
         Top             =   5910
         Width           =   11715
         _Version        =   65536
         _ExtentX        =   20664
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
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Top             =   390
            Width           =   2775
         End
         Begin VB.TextBox txt_NumCom 
            Height          =   315
            Left            =   1590
            MaxLength       =   25
            TabIndex        =   28
            Text            =   "Text1"
            Top             =   1050
            Width           =   2775
         End
         Begin VB.ComboBox cmb_CodBan 
            Height          =   315
            Left            =   1590
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   60
            Width           =   2775
         End
         Begin EditLib.fpDoubleSingle ipp_Import 
            Height          =   315
            Left            =   1590
            TabIndex        =   31
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
            Left            =   1590
            TabIndex        =   27
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
         Begin VB.Label Label11 
            Caption         =   "Nro. Comprobante:"
            Height          =   285
            Left            =   90
            TabIndex        =   43
            Top             =   1050
            Width           =   1335
         End
         Begin VB.Label Label10 
            Caption         =   "Fecha de Pago:"
            Height          =   315
            Left            =   90
            TabIndex        =   42
            Top             =   720
            Width           =   1245
         End
         Begin VB.Label Label7 
            Caption         =   "Nro. Cuenta:"
            Height          =   285
            Left            =   90
            TabIndex        =   41
            Top             =   390
            Width           =   1335
         End
         Begin VB.Label Label6 
            Caption         =   "Banco:"
            Height          =   315
            Left            =   90
            TabIndex        =   40
            Top             =   60
            Width           =   1455
         End
         Begin VB.Label Label41 
            Caption         =   "Importe:"
            Height          =   285
            Left            =   90
            TabIndex        =   2
            Top             =   1380
            Width           =   1125
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   3
         Top             =   30
         Width           =   11715
         _Version        =   65536
         _ExtentX        =   20664
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
            TabIndex        =   4
            Top             =   30
            Width           =   7215
            _Version        =   65536
            _ExtentX        =   12726
            _ExtentY        =   1032
            _StockProps     =   15
            Caption         =   "Operaciones por Bancos - Pago de Gastos de Cierre"
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
            Picture         =   "OpeTra_frm_003.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel20 
         Height          =   795
         Left            =   30
         TabIndex        =   5
         Top             =   750
         Width           =   11715
         _Version        =   65536
         _ExtentX        =   20664
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
         Begin VB.CommandButton cmd_Buscar 
            Height          =   675
            Left            =   9570
            Picture         =   "OpeTra_frm_003.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Buscar Registros"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   675
            Left            =   10290
            Picture         =   "OpeTra_frm_003.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Limpiar Datos"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   11010
            Picture         =   "OpeTra_frm_003.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   675
         End
         Begin VB.ComboBox cmb_TipBus 
            Height          =   315
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   60
            Width           =   2775
         End
         Begin VB.TextBox txt_NumDoc 
            Height          =   315
            Left            =   6210
            MaxLength       =   12
            TabIndex        =   7
            Text            =   "Text1"
            Top             =   390
            Width           =   2775
         End
         Begin VB.ComboBox cmb_TipDoc 
            Height          =   315
            Left            =   6210
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   60
            Width           =   2775
         End
         Begin MSMask.MaskEdBox msk_NumSol 
            Height          =   315
            Left            =   1620
            TabIndex        =   12
            Top             =   390
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Mask            =   "###-###-##-####"
            PromptChar      =   " "
         End
         Begin VB.Label Label17 
            Caption         =   "Nro. Doc. Id.:"
            Height          =   285
            Left            =   60
            TabIndex        =   17
            Top             =   1740
            Width           =   1065
         End
         Begin VB.Label Label18 
            Caption         =   "Tipo de Búsqueda:"
            Height          =   315
            Left            =   90
            TabIndex        =   16
            Top             =   60
            Width           =   1455
         End
         Begin VB.Label Label19 
            Caption         =   "Nro. Doc. Ident.:"
            Height          =   285
            Left            =   4830
            TabIndex        =   15
            Top             =   390
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo Doc. Ident.:"
            Height          =   315
            Left            =   4830
            TabIndex        =   14
            Top             =   60
            Width           =   1395
         End
         Begin VB.Label lbl_Numero 
            Caption         =   "Nro. Solicitud:"
            Height          =   285
            Left            =   90
            TabIndex        =   13
            Top             =   390
            Width           =   1335
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   795
         Left            =   30
         TabIndex        =   18
         Top             =   1590
         Width           =   11715
         _Version        =   65536
         _ExtentX        =   20664
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
         Begin Threed.SSPanel pnl_NumSol 
            Height          =   315
            Left            =   1620
            TabIndex        =   19
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
         Begin Threed.SSPanel pnl_Client 
            Height          =   315
            Left            =   1620
            TabIndex        =   20
            Top             =   390
            Width           =   10035
            _Version        =   65536
            _ExtentX        =   17701
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
         Begin VB.Label Label3 
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   90
            TabIndex        =   22
            Top             =   390
            Width           =   1125
         End
         Begin VB.Label Label2 
            Caption         =   "Nro. de Solicitud:"
            Height          =   315
            Left            =   90
            TabIndex        =   21
            Top             =   60
            Width           =   1395
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   3435
         Left            =   30
         TabIndex        =   23
         Top             =   2430
         Width           =   11715
         _Version        =   65536
         _ExtentX        =   20664
         _ExtentY        =   6059
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
         Begin Threed.SSPanel pnl_Import 
            Height          =   315
            Left            =   9660
            TabIndex        =   24
            Top             =   2370
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
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   2025
            Left            =   30
            TabIndex        =   25
            Top             =   330
            Width           =   11595
            _ExtentX        =   20452
            _ExtentY        =   3572
            _Version        =   393216
            Rows            =   21
            Cols            =   4
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel11 
            Height          =   285
            Left            =   60
            TabIndex        =   29
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
         Begin Threed.SSPanel SSPanel13 
            Height          =   285
            Left            =   6270
            TabIndex        =   30
            Top             =   60
            Width           =   3405
            _Version        =   65536
            _ExtentX        =   6006
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Moneda"
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
            Left            =   9660
            TabIndex        =   33
            Top             =   60
            Width           =   1665
            _Version        =   65536
            _ExtentX        =   2937
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Importe"
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
         Begin Threed.SSPanel pnl_ITFImp 
            Height          =   315
            Left            =   9660
            TabIndex        =   34
            Top             =   2700
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
            Left            =   9660
            TabIndex        =   35
            Top             =   3030
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
         Begin VB.Label Label5 
            Caption         =   "Sub-Total:"
            Height          =   285
            Left            =   8640
            TabIndex        =   38
            Top             =   2370
            Width           =   975
         End
         Begin VB.Label Label8 
            Caption         =   "ITF:"
            Height          =   285
            Left            =   8640
            TabIndex        =   37
            Top             =   2700
            Width           =   975
         End
         Begin VB.Label Label9 
            Caption         =   "Total:"
            Height          =   285
            Left            =   8640
            TabIndex        =   36
            Top             =   3030
            Width           =   975
         End
      End
      Begin Threed.SSPanel SSPanel12 
         Height          =   765
         Left            =   30
         TabIndex        =   39
         Top             =   7710
         Width           =   11715
         _Version        =   65536
         _ExtentX        =   20664
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
            Left            =   11010
            Picture         =   "OpeTra_frm_003.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   32
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   675
         End
         Begin MSComDlg.CommonDialog dlg_Guarda 
            Left            =   0
            Top             =   0
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
      End
   End
End
Attribute VB_Name = "frm_Cob_GasAdm_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_dbl_PorITF     As Double
Dim l_int_TipMon     As Integer
Dim l_arr_CodBan()   As moddat_tpo_Genera
Dim l_arr_CtaBan()   As moddat_tpo_Genera

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

Private Sub cmb_TipMon_Click()
   Call gs_SetFocus(ipp_Import)
End Sub

Private Sub cmb_TipMon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipMon_Click
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
      If Len(Trim(msk_NumSol.Text)) < 12 Then
         MsgBox "Debe ingresar el Número de Documento de Identidad.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NumDoc)
         Exit Sub
      End If
      
      moddat_g_str_NumSol = msk_NumSol.Text
   End If
   
   If cmb_TipBus.ItemData(cmb_TipBus.ListIndex) = 1 Then
      g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
      g_str_Parame = g_str_Parame & "SOLMAE_TITTDO = " & CStr(moddat_g_int_TipDoc) & " AND "
      g_str_Parame = g_str_Parame & "SOLMAE_TITNDO = '" & moddat_g_str_NumDoc & "' AND "
      g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 "
   Else
      g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
      g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = '" & moddat_g_str_NumSol & "' AND "
      g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 "
   End If
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Call cmd_Limpia_Click
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      MsgBox "No existe Solicitud en Trámite para la Selección de Búsqueda. ", vbExclamation, modgen_g_str_NomPlt
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Call cmd_Limpia_Click
      Exit Sub
   End If

   moddat_g_int_TipDoc = g_rst_Princi!SOLMAE_TITTDO
   moddat_g_str_NumDoc = Trim(g_rst_Princi!SOLMAE_TITNDO)
   moddat_g_str_NumSol = Trim(g_rst_Princi!SOLMAE_NUMERO)
   moddat_g_str_CodPrd = Trim(g_rst_Princi!SOLMAE_CODPRD)
   moddat_g_str_CodSub = Trim(g_rst_Princi!SOLMAE_CODSUB)
   
   'Obteniendo Nombre de Cliente
   moddat_g_str_NomCli = moddat_gf_Buscar_NomCli(moddat_g_int_TipDoc, moddat_g_str_NumDoc)

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   pnl_NumSol.Caption = Mid(moddat_g_str_NumSol, 1, 3) & "-" & Mid(moddat_g_str_NumSol, 4, 3) & "-" & Mid(moddat_g_str_NumSol, 7, 2) & "-" & Mid(moddat_g_str_NumSol, 9, 4)
   pnl_Client.Caption = CStr(moddat_g_int_TipDoc) & "-" & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   
   'Buscando Gastos Administrativos
   If Not ff_Buscar_GasAdm() Then
      Call cmd_Limpia_Click
      Exit Sub
   End If
   
   Call fs_Activa(False)
   Call gs_SetFocus(cmb_CodBan)
End Sub

Private Sub cmd_Grabar_Click()
   Dim r_int_Contad     As Integer
   Dim r_int_CodGas     As Integer
   Dim r_dbl_Import     As Double
   Dim r_dbl_ImpITF     As Double
   Dim r_dbl_SubTot     As Double
   Dim r_str_Operac     As String
   Dim r_lng_NumMov     As Long
   
   'On Error GoTo Error_Imp

   
   If cmb_CodBan.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Banco.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_CodBan)
      Exit Sub
   End If
   
   If cmb_CtaBan.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Número de Cuenta.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_CtaBan)
      Exit Sub
   End If
   
   If CDate(ipp_FecPag.Text) > Date Then
      MsgBox "Debe ingresar una fecha correcta.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecPag)
      Exit Sub
   End If
   
   If Len(Trim(txt_NumCom.Text)) = 0 Then
      MsgBox "Debe ingresar el Número de Comprobante.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NumCom)
      Exit Sub
   End If
   
   r_dbl_SubTot = CDbl(pnl_Import.Caption)
   r_dbl_ImpITF = CDbl(pnl_ITFImp.Caption)
   
   If MsgBox("¿Está seguro de registrar la transacción?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   r_str_Operac = moddat_gf_Consulta_Operac(moddat_g_str_CodPrd, "210")
   r_str_Operac = CStr(l_int_TipMon) & Right(r_str_Operac, 5)
   
   'Obteniendo Número de Movimiento
   r_lng_NumMov = opecaj_gf_Genera_NumMov()
   
   'Registrando Movimiento
   If Not opecaj_gf_Inserta_CajMov(modgen_g_str_CodUsu, "1101", moddat_g_str_NumSol, "", moddat_g_int_TipDoc, moddat_g_str_NumDoc, l_arr_CodBan(cmb_CodBan.ListIndex + 1).Genera_Codigo, Format(CDate(ipp_FecPag.Text), "yyyymmdd"), l_arr_CtaBan(cmb_CtaBan.ListIndex + 1).Genera_Codigo, txt_NumCom.Text, l_int_TipMon, CDbl(r_dbl_SubTot), 0, modgen_g_str_CodSuc, 0, 0, 0, l_dbl_PorITF, r_dbl_ImpITF, CDbl(ipp_Import.Text), 0, "0", r_str_Operac, r_lng_NumMov, 1, "0", "", "", "") Then
      Exit Sub
   End If
   
   'Actualizando Saldo de Caja
   If Not opecaj_gf_ActualizaSaldo(l_arr_CodBan(cmb_CodBan.ListIndex + 1).Genera_Codigo, l_int_TipMon, CDbl(ipp_Import.Text)) Then
      Exit Sub
   End If
   
   'Actualizando Pago en Tabla de Gasto Administrativo
   grd_Listad.Redraw = False
   For r_int_Contad = 0 To grd_Listad.Rows - 1
      grd_Listad.Row = r_int_Contad
      
      grd_Listad.Col = 3
      r_int_CodGas = CInt(grd_Listad.Text)
      
      grd_Listad.Col = 2
      r_dbl_Import = CDbl(grd_Listad.Text)
      
      r_str_Operac = moddat_gf_Consulta_Operac(moddat_g_str_CodPrd, "2" & Format(r_int_CodGas, "00"))
      r_str_Operac = CStr(l_int_TipMon) & Right(r_str_Operac, 5)
      
      If Not opecaj_gf_Pago_GasAdm(moddat_g_str_NumSol, r_int_CodGas, l_int_TipMon, r_dbl_Import, l_dbl_PorITF, Format(CDate(ipp_FecPag.Text), "yyyymmdd"), r_str_Operac) Then
         Exit Sub
      End If
   Next r_int_Contad
   grd_Listad.Redraw = True
   
   'Actualizando en Seguimiento de Tasacion Pago de Gastos Administrativos
   If Not moddat_gf_Inserta_SegDet(moddat_g_str_NumSol, 32, 25, 0, "", 0, 0) Then
      Exit Sub
   End If
   
   'dlg_Guarda.CancelError = True
   'dlg_Guarda.ShowPrinter
   
   'Impresión de Voucher
   'Call opecaj_gs_Imp_GasAdm_Ban(Format(CDate(moddat_g_str_FecSis), "yyyymmdd"), CStr(r_lng_NumMov))
   'Call gs_Imprim_ComPag

   Call cmd_Limpia_Click
   Exit Sub
   
Error_Imp:
   Call cmd_Limpia_Click
   Exit Sub

End Sub

Private Sub cmd_Limpia_Click()
   Call fs_Activa(True)
   
   cmb_TipBus.ListIndex = -1
   cmb_TipDoc.Enabled = False
   txt_NumDoc.Enabled = False
   msk_NumSol.Enabled = False

   msk_NumSol.Mask = ""
   msk_NumSol.Text = ""
   msk_NumSol.Mask = "###-###-##-####"
   
   txt_NumDoc.Text = ""
   
   pnl_Client.Caption = ""
   pnl_NumSol.Caption = ""
   
   Call gs_LimpiaGrid(grd_Listad)
   
   pnl_Import.Caption = "0.00 "
   pnl_ITFImp.Caption = "0.00 "
   pnl_TotImp.Caption = "0.00 "
   
   cmb_CodBan.ListIndex = -1
   cmb_CtaBan.Clear
   ipp_FecPag.Text = Format(Date, "dd/mm/yyyy")
   txt_NumCom.Text = ""
   ipp_Import.Value = 0
   
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

Private Sub fs_Activa(ByVal p_Habilita As Integer)
   cmb_TipBus.Enabled = p_Habilita
   cmb_TipDoc.Enabled = p_Habilita
   txt_NumDoc.Enabled = p_Habilita
   msk_NumSol.Enabled = p_Habilita
   cmd_Buscar.Enabled = p_Habilita
   
   grd_Listad.Enabled = Not p_Habilita
   cmb_CodBan.Enabled = Not p_Habilita
   cmb_CtaBan.Enabled = Not p_Habilita
   ipp_FecPag.Enabled = Not p_Habilita
   txt_NumCom.Enabled = Not p_Habilita
   ipp_Import.Enabled = Not p_Habilita
   cmd_Grabar.Enabled = Not p_Habilita
End Sub

Private Sub cmb_TipBus_Click()
   If cmb_TipBus.ListIndex > -1 Then
      If cmb_TipBus.ItemData(cmb_TipBus.ListIndex) = 1 Then
         cmb_TipDoc.Enabled = True
         txt_NumDoc.Enabled = True
         msk_NumSol.Enabled = False
         
         msk_NumSol.Mask = ""
         msk_NumSol.Text = ""
         msk_NumSol.Mask = "###-###-##-####"
         
         Call gs_SetFocus(cmb_TipDoc)
      Else
         cmb_TipDoc.Enabled = False
         txt_NumDoc.Enabled = False
         msk_NumSol.Enabled = True
         
         cmb_TipDoc.ListIndex = -1
         txt_NumDoc.Text = ""
         
         Call gs_SetFocus(msk_NumSol)
      End If
   Else
      cmb_TipDoc.Enabled = False
      txt_NumDoc.Enabled = False
      
      msk_NumSol.Enabled = False
   
      cmb_TipDoc.ListIndex = -1
      txt_NumDoc.Text = ""
      msk_NumSol.Mask = ""
      msk_NumSol.Text = ""
      msk_NumSol.Mask = "###-###-##-####"
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

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

Private Sub ipp_FecPag_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NumCom)
   End If
End Sub

Private Sub ipp_Import_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   End If
End Sub

Private Sub msk_NumSol_GotFocus()
   Call gs_SelecTodo(msk_NumSol)
End Sub

Private Sub msk_NumSol_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Buscar)
   End If
End Sub

Private Sub txt_NumCta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecPag)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_NumDoc_GotFocus()
   Call gs_SelecTodo(txt_NumDoc)
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

Private Function ff_Buscar_GasAdm() As Integer
   Dim r_dbl_Import     As Double
   
   ff_Buscar_GasAdm = False
   
   l_int_TipMon = 0
   r_dbl_Import = 0
   
   Call gs_LimpiaGrid(grd_Listad)
   
   g_str_Parame = "SELECT COUNT(*) AS VS_CONTAD FROM TRA_GASADM WHERE "
   g_str_Parame = g_str_Parame & "GASADM_NUMSOL = '" & moddat_g_str_NumSol & "' AND "
   g_str_Parame = g_str_Parame & "GASADM_SITUAC = 2"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Function
   End If
   
   If g_rst_Princi!VS_CONTAD = 0 Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      MsgBox "No tiene Gastos Administrativos para cancelar.", vbExclamation, modgen_g_con_PltPar
      
      Exit Function
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   
   'Lista de Gastos Administrativos
   g_str_Parame = "SELECT * FROM TRA_GASADM WHERE "
   g_str_Parame = g_str_Parame & "GASADM_NUMSOL = '" & moddat_g_str_NumSol & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Function
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
     g_rst_Princi.Close
     Set g_rst_Princi = Nothing
     
     MsgBox "No se han asignado Gastos Administrativos.", vbExclamation, modgen_g_con_PltPar
     Exit Function
   End If
   
   grd_Listad.Redraw = False
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      
      grd_Listad.Row = grd_Listad.Rows - 1
      
      'Buscando Descripción de Gastos Administrativos
      grd_Listad.Col = 0
      If moddat_gf_Consulta_ParSubPrd(moddat_g_arr_Genera(), moddat_g_str_CodPrd, moddat_g_str_CodSub, "007", Format(g_rst_Princi!GASADM_CODGAS, "00") & Format(g_rst_Princi!GASADM_TIPMON, "0")) Then
         grd_Listad.Text = Trim(moddat_g_arr_Genera(1).Genera_Nombre)
      End If
      
      grd_Listad.Col = 3
      grd_Listad.Text = g_rst_Princi!GASADM_CODGAS
      
      'Tipo de Moneda
      grd_Listad.Col = 1
      grd_Listad.Text = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!GASADM_TIPMON))
      
      l_int_TipMon = g_rst_Princi!GASADM_TIPMON
      
      'Importe
      grd_Listad.Col = 2
      grd_Listad.Text = Format(g_rst_Princi!GASADM_IMPORT, "###,###,##0.00")
      
      r_dbl_Import = r_dbl_Import + g_rst_Princi!GASADM_IMPORT
      
      g_rst_Princi.MoveNext
   Loop
      
   grd_Listad.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   pnl_Import.Caption = Format(r_dbl_Import, "###,###,##0.00") & " "
   
   pnl_ITFImp.Caption = gf_Truncar_Numero(CDbl(pnl_Import.Caption) * (l_dbl_PorITF / 100), 2) & " "
   pnl_TotImp.Caption = Format(CDbl(pnl_Import.Caption) + CDbl(Trim(pnl_ITFImp.Caption)), "###,###,##0.00") & " "
   
   Call gs_UbiIniGrid(grd_Listad)
   
   ff_Buscar_GasAdm = True
End Function

Private Sub fs_Inicia()
   Call moddat_gs_FecSis
   
   grd_Listad.ColWidth(0) = 6210
   grd_Listad.ColWidth(1) = 3390
   grd_Listad.ColWidth(2) = 1650
   grd_Listad.ColWidth(3) = 0
   
   grd_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignRightCenter

   Call moddat_gs_Carga_LisIte(cmb_CodBan, l_arr_CodBan, 1, "516")
   Call modsis_gs_Carga_TipBus(cmb_TipBus)
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipDoc, 1, "230")
   
   'Obteniendo ITF
   l_dbl_PorITF = opecaj_gf_Consulta_ITF(Format(CDate(moddat_g_str_FecSis), "yyyymmdd"), 1)
End Sub

