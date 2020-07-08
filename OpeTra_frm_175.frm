VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_MntCli_64 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7350
   ClientLeft      =   2865
   ClientTop       =   2775
   ClientWidth     =   11640
   Icon            =   "OpeTra_frm_175.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   7365
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   11685
      _Version        =   65536
      _ExtentX        =   20611
      _ExtentY        =   12991
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
      Begin Threed.SSPanel SSPanel4 
         Height          =   795
         Left            =   30
         TabIndex        =   25
         Top             =   3750
         Width           =   11595
         _Version        =   65536
         _ExtentX        =   20452
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
         Begin VB.ComboBox cmb_MonIng 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   60
            Width           =   3315
         End
         Begin VB.ComboBox cmb_ConLoc 
            Height          =   315
            Left            =   8220
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   450
            Width           =   765
         End
         Begin EditLib.fpDoubleSingle ipp_IngNet 
            Height          =   315
            Left            =   8220
            TabIndex        =   10
            Top             =   120
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
         Begin EditLib.fpDateTime ipp_IniAct 
            Height          =   315
            Left            =   2040
            TabIndex        =   11
            Top             =   420
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
         Begin VB.Label lbl_General 
            Caption         =   "Ingreso Declarado:"
            Height          =   285
            Index           =   61
            Left            =   6210
            TabIndex        =   29
            Top             =   120
            Width           =   1755
         End
         Begin VB.Label lbl_General 
            Caption         =   "Fecha Inicio Actividades:"
            Height          =   315
            Index           =   58
            Left            =   120
            TabIndex        =   28
            Top             =   420
            Width           =   1845
         End
         Begin VB.Label lbl_General 
            Caption         =   "Moneda de Ingresos:"
            Height          =   285
            Index           =   0
            Left            =   90
            TabIndex        =   27
            Top             =   60
            Width           =   1665
         End
         Begin VB.Label Label11 
            Caption         =   "Contrato Locación:"
            Height          =   285
            Left            =   6210
            TabIndex        =   26
            Top             =   450
            Width           =   1785
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   1755
         Left            =   30
         TabIndex        =   30
         Top             =   1950
         Width           =   11595
         _Version        =   65536
         _ExtentX        =   20452
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
         Begin VB.CommandButton cmd_DirCas 
            Caption         =   "="
            Height          =   315
            Left            =   1560
            TabIndex        =   23
            ToolTipText     =   "Obtener Dirección de Domicilio"
            Top             =   390
            Width           =   435
         End
         Begin VB.TextBox txt_CodPos 
            Height          =   315
            Left            =   8220
            MaxLength       =   250
            TabIndex        =   4
            Text            =   "Text1"
            Top             =   720
            Width           =   1640
         End
         Begin VB.ComboBox cmb_PrvEst 
            Height          =   315
            Left            =   2010
            TabIndex        =   3
            Text            =   "cmb_DptDir"
            Top             =   720
            Width           =   3345
         End
         Begin VB.TextBox txt_Direcc 
            Height          =   315
            Left            =   2010
            MaxLength       =   250
            TabIndex        =   2
            Text            =   "Text1"
            Top             =   390
            Width           =   9555
         End
         Begin VB.TextBox txt_Telef1 
            Height          =   315
            Left            =   2010
            MaxLength       =   25
            TabIndex        =   5
            Text            =   "Text1"
            Top             =   1050
            Width           =   1640
         End
         Begin VB.TextBox txt_NumDoc 
            Height          =   315
            Left            =   8220
            MaxLength       =   11
            TabIndex        =   1
            Text            =   "Text1"
            Top             =   60
            Width           =   3345
         End
         Begin VB.ComboBox cmb_TipDoc 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   3315
         End
         Begin VB.TextBox txt_NumFax 
            Height          =   315
            Left            =   8220
            MaxLength       =   25
            TabIndex        =   7
            Text            =   "Text1"
            Top             =   1050
            Width           =   1640
         End
         Begin VB.TextBox txt_Telef2 
            Height          =   315
            Left            =   3660
            MaxLength       =   25
            TabIndex        =   6
            Text            =   "Text1"
            Top             =   1050
            Width           =   1640
         End
         Begin VB.ComboBox cmb_CodCiu 
            Height          =   315
            Left            =   2010
            TabIndex        =   8
            Text            =   "cmb_DptDir"
            Top             =   1380
            Width           =   9555
         End
         Begin VB.Label Label20 
            Caption         =   "Dirección:"
            Height          =   285
            Left            =   90
            TabIndex        =   54
            Top             =   390
            Width           =   1485
         End
         Begin VB.Label Label28 
            Caption         =   "Código Postal:"
            Height          =   285
            Left            =   6210
            TabIndex        =   53
            Top             =   720
            Width           =   1485
         End
         Begin VB.Label Label24 
            Caption         =   "Provincia / Estado:"
            Height          =   315
            Left            =   90
            TabIndex        =   52
            Top             =   720
            Width           =   1905
         End
         Begin VB.Label lbl_General 
            Caption         =   "CIIU:"
            Height          =   285
            Index           =   39
            Left            =   90
            TabIndex        =   35
            Top             =   1380
            Width           =   1365
         End
         Begin VB.Label lbl_General 
            Caption         =   "Número Docum. Ident.:"
            Height          =   285
            Index           =   48
            Left            =   6210
            TabIndex        =   34
            Top             =   60
            Width           =   1845
         End
         Begin VB.Label lbl_General 
            Caption         =   "Tipo Docum. Ident.:"
            Height          =   285
            Index           =   36
            Left            =   90
            TabIndex        =   33
            Top             =   60
            Width           =   1635
         End
         Begin VB.Label lbl_General 
            Caption         =   "Teléfono (s):"
            Height          =   285
            Index           =   46
            Left            =   90
            TabIndex        =   32
            Top             =   1050
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Fax:"
            Height          =   285
            Index           =   55
            Left            =   6210
            TabIndex        =   31
            Top             =   1050
            Width           =   1485
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   36
         Top             =   30
         Width           =   11595
         _Version        =   65536
         _ExtentX        =   20452
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
         Begin Threed.SSPanel SSPanel16 
            Height          =   255
            Left            =   690
            TabIndex        =   37
            Top             =   60
            Width           =   6165
            _Version        =   65536
            _ExtentX        =   10874
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
            Left            =   690
            TabIndex        =   38
            Top             =   330
            Width           =   6165
            _Version        =   65536
            _ExtentX        =   10874
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "Actividades Económicas - Profesional Independiente"
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
            Picture         =   "OpeTra_frm_175.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   435
         Left            =   30
         TabIndex        =   39
         Top             =   1470
         Width           =   11595
         _Version        =   65536
         _ExtentX        =   20452
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
         Begin Threed.SSPanel pnl_NomCli 
            Height          =   315
            Left            =   2010
            TabIndex        =   40
            Top             =   60
            Width           =   9525
            _Version        =   65536
            _ExtentX        =   16801
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "1-07522154 / IKEHARA PUNK MIGUEL ANGEL (1-07521154 / IKEHARA PUNK MIGUEL ANGEL)"
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
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   90
            TabIndex        =   41
            Top             =   60
            Width           =   1725
         End
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   675
         Left            =   30
         TabIndex        =   42
         Top             =   750
         Width           =   11595
         _Version        =   65536
         _ExtentX        =   20452
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
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10980
            Picture         =   "OpeTra_frm_175.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_175.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel10 
         Height          =   1425
         Left            =   30
         TabIndex        =   43
         Top             =   5400
         Width           =   11595
         _Version        =   65536
         _ExtentX        =   20452
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
            TabIndex        =   18
            Top             =   60
            Width           =   11475
            _ExtentX        =   20241
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
      Begin Threed.SSPanel SSPanel8 
         Height          =   765
         Left            =   30
         TabIndex        =   44
         Top             =   4590
         Width           =   11595
         _Version        =   65536
         _ExtentX        =   20452
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
         Begin VB.CommandButton cmd_Editar 
            Height          =   585
            Left            =   10980
            Picture         =   "OpeTra_frm_175.frx":0B9A
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   60
            Width           =   585
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   10380
            Picture         =   "OpeTra_frm_175.frx":0EA4
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Limpiar Datos"
            Top             =   60
            Width           =   585
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   585
            Left            =   9780
            Picture         =   "OpeTra_frm_175.frx":11AE
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Buscar Datos"
            Top             =   60
            Width           =   585
         End
         Begin VB.TextBox txt_NDoEmp 
            Height          =   315
            Left            =   1980
            MaxLength       =   11
            TabIndex        =   14
            Text            =   "Text1"
            Top             =   390
            Width           =   3315
         End
         Begin VB.ComboBox cmb_TDoEmp 
            Height          =   315
            Left            =   1980
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   60
            Width           =   3315
         End
         Begin VB.Label lbl_General 
            Caption         =   "Número Docum. Ident.:"
            Height          =   285
            Index           =   1
            Left            =   60
            TabIndex        =   46
            Top             =   390
            Width           =   1845
         End
         Begin VB.Label lbl_General 
            Caption         =   "Tipo Docum. Ident.:"
            Height          =   285
            Index           =   2
            Left            =   60
            TabIndex        =   45
            Top             =   60
            Width           =   1635
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   435
         Left            =   30
         TabIndex        =   47
         Top             =   6870
         Width           =   11595
         _Version        =   65536
         _ExtentX        =   20452
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
         Begin VB.TextBox txt_NomCar 
            Height          =   315
            Left            =   8190
            MaxLength       =   250
            TabIndex        =   20
            Text            =   "Text1"
            Top             =   60
            Width           =   3315
         End
         Begin VB.CommandButton cmd_BusEmp_Ant 
            Caption         =   "..."
            Height          =   315
            Left            =   10620
            TabIndex        =   48
            ToolTipText     =   "Obtener Dirección de Domicilio"
            Top             =   6600
            Width           =   435
         End
         Begin EditLib.fpDateTime ipp_FecIng 
            Height          =   315
            Left            =   1980
            TabIndex        =   19
            Top             =   60
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
         Begin Threed.SSPanel pnl_FlgEmp_Ant 
            Height          =   315
            Left            =   11100
            TabIndex        =   49
            Top             =   6600
            Width           =   435
            _Version        =   65536
            _ExtentX        =   767
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "NR"
            ForeColor       =   255
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderWidth     =   0
         End
         Begin VB.Label lbl_General 
            Caption         =   "Fecha de Ingreso:"
            Height          =   315
            Index           =   5
            Left            =   60
            TabIndex        =   51
            Top             =   60
            Width           =   1365
         End
         Begin VB.Label lbl_General 
            Caption         =   "Cargo:"
            Height          =   285
            Index           =   57
            Left            =   6180
            TabIndex        =   50
            Top             =   60
            Width           =   1665
         End
      End
   End
End
Attribute VB_Name = "frm_MntCli_64"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_PrvEst()   As moddat_tpo_Genera
Dim l_str_CodCiu     As String
Dim l_str_PrvEst     As String
Dim l_int_FlgCmb     As Integer

Private Sub cmb_CodCiu_Change()
   l_str_CodCiu = cmb_CodCiu.Text
End Sub

Private Sub cmb_CodCiu_Click()
   If cmb_CodCiu.ListIndex > -1 Then
      If l_int_FlgCmb Then
         Call gs_SetFocus(cmb_MonIng)
      End If
   End If
End Sub

Private Sub cmb_CodCiu_GotFocus()
   l_int_FlgCmb = True
   l_str_CodCiu = cmb_CodCiu.Text
End Sub

Private Sub cmb_CodCiu_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + modgen_g_con_NUMERO + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_CodCiu, l_str_CodCiu)
      l_int_FlgCmb = True
      
      If cmb_CodCiu.ListIndex > -1 Then
         l_str_CodCiu = ""
      End If
      
      Call gs_SetFocus(cmb_MonIng)
   End If
End Sub

Private Sub cmb_ConLoc_Click()
   If cmb_ConLoc.ListIndex > -1 Then
      If cmb_ConLoc.ItemData(cmb_ConLoc.ListIndex) = 1 Then
         cmb_TDoEmp.Enabled = True
         txt_NDoEmp.Enabled = True
         ipp_FecIng.Enabled = True
         txt_NomCar.Enabled = True
         grd_Listad.Enabled = True
         
         cmd_Buscar.Enabled = True
         cmd_Limpia.Enabled = True
         cmd_Editar.Enabled = True
         
         Call fs_Activa_Emp(True)
         
         Call gs_SetFocus(cmb_TDoEmp)
      Else
         cmb_TDoEmp.ListIndex = -1
         txt_NDoEmp.Text = ""
         ipp_FecIng.Text = Format(date, "dd/mm/yyyy")
         txt_NomCar.Text = ""
         
         Call gs_LimpiaGrid(grd_Listad)
         
         cmb_TDoEmp.Enabled = False
         txt_NDoEmp.Enabled = False
         ipp_FecIng.Enabled = False
         txt_NomCar.Enabled = False
         grd_Listad.Enabled = False
         
         cmd_Buscar.Enabled = False
         cmd_Limpia.Enabled = False
         cmd_Editar.Enabled = False
         
         Call gs_SetFocus(cmd_Grabar)
      End If
   Else
      cmb_TDoEmp.ListIndex = -1
      txt_NDoEmp.Text = ""
      ipp_FecIng.Text = Format(date, "dd/mm/yyyy")
      txt_NomCar.Text = ""
      
      Call gs_LimpiaGrid(grd_Listad)
      
      cmb_TDoEmp.Enabled = False
      txt_NDoEmp.Enabled = False
      ipp_FecIng.Enabled = False
      txt_NomCar.Enabled = False
      grd_Listad.Enabled = False
      
      Call gs_SetFocus(cmd_Grabar)
   End If
End Sub

Private Sub cmb_MonIng_Click()
   Call gs_SetFocus(ipp_IngNet)
End Sub

Private Sub cmb_MonIng_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_MonIng_Click
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

Private Sub cmb_TDoEmp_Click()
   Call gs_SetFocus(txt_NDoEmp)
End Sub

Private Sub cmb_TDoEmp_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TDoEmp_Click
   End If
End Sub

Private Sub cmb_TipDoc_Click()
   Call gs_SetFocus(txt_NumDoc)
End Sub

Private Sub cmb_TipDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipDoc_Click
   End If
End Sub

Private Sub cmd_Buscar_Click()
   If cmb_TDoEmp.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Documento.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TDoEmp)
      Exit Sub
   End If
   
   If Len(Trim(txt_NDoEmp.Text)) = 0 Then
      MsgBox "Debe ingresar el Número de Documento.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NDoEmp)
      Exit Sub
   End If
   
   'Buscar Empresa en Maestro de Empresas
   modmip_g_int_TDoEmp = CStr(cmb_TDoEmp.ItemData(cmb_TDoEmp.ListIndex))
   modmip_g_str_TDoEmp = cmb_TipDoc.Text
   modmip_g_str_NDoEmp = Trim(txt_NDoEmp.Text)
   
   g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(modmip_g_int_TDoEmp) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & modmip_g_str_NDoEmp & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   
      cmd_Editar.Enabled = False
   
      modmip_g_int_FlgGrb = 1
      modmip_g_int_FlgAct_2 = 1
      
      frm_MntEmp_52.Show 1
      
      If modmip_g_int_FlgAct_2 = 1 Then
         Exit Sub
      End If
   Else
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   End If
   
   Call fs_BusEmp
End Sub

Private Sub cmd_DirCas_Click()
   txt_Direcc.Text = frm_MntCli_52.txt_Direcc.Text
   cmb_PrvEst.ListIndex = frm_MntCli_52.cmb_PrvEst.ListIndex
   txt_CodPos.Text = frm_MntCli_52.txt_CodPos.Text
   
   txt_Telef1.Text = frm_MntCli_52.txt_TelFij.Text
   
   Call gs_SetFocus(cmb_CodCiu)
End Sub

Private Sub cmd_Editar_Click()
   modmip_g_int_FlgGrb = 2
   modmip_g_int_FlgAct_2 = 1
   
   frm_MntEmp_52.Show 1
   
   If modmip_g_int_FlgAct_2 = 2 Then
      Screen.MousePointer = 11
      
      Call fs_BusEmp
      
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmd_Grabar_Click()
   If cmb_TipDoc.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Documento.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipDoc)
      
      Exit Sub
   End If

   If Len(Trim(txt_NumDoc.Text)) = 0 Then
      MsgBox "Debe ingresar el Número de Documento.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NumDoc)
      Exit Sub
   End If
   
   If Len(Trim(txt_Direcc.Text)) = 0 Then
      MsgBox "Debe ingresar la Dirección.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_Direcc)
      Exit Sub
   End If
   
   If cmb_PrvEst.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Provincia o Estado.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PrvEst)
      Exit Sub
   End If
   
   If Len(Trim(txt_CodPos.Text)) = 0 Then
      MsgBox "Debe ingresar el Código Postal.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_CodPos)
      Exit Sub
   End If
   
   If cmb_CodCiu.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Código de CIIU.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_CodCiu)
      Exit Sub
   End If
   
   If cmb_MonIng.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Moneda de Ingresos.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_MonIng)
      Exit Sub
   End If
   
   If ipp_IngNet.Value = 0 Then
      MsgBox "Debe ingresar el Ingreso Neto.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_IngNet)
      Exit Sub
   End If
   
   If cmb_ConLoc.ListIndex = -1 Then
      MsgBox "Debe ingresar si el cliente tiene Contrato de Locación.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_ConLoc)
      Exit Sub
   End If
   
   If cmb_ConLoc.ItemData(cmb_ConLoc.ListIndex) = 1 Then
      If cmb_TDoEmp.ListIndex = -1 Then
         MsgBox "Debe ingresar el Tipo de Documento de la Empresa.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_TDoEmp)
         Exit Sub
      End If
      
      If Len(Trim(txt_NDoEmp.Text)) = 0 Then
         MsgBox "Debe ingresar el Número de Documento de la Empresa.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NDoEmp)
         Exit Sub
      End If
      
      If grd_Listad.Rows = 0 Then
         MsgBox "Debe ingresar la información de la Empresa.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_TDoEmp)
         Exit Sub
      End If
   
      If Len(Trim(txt_NomCar.Text)) = 0 Then
         MsgBox "Debe ingresar el Cargo.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NomCar)
         Exit Sub
      End If
   End If
   
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   
   If modmip_g_int_FlgGrb_1 = 2 Then
      'Borrar Actividad Económica
      g_str_Parame = "DELETE FROM CLI_ACTECO WHERE "
      
      If modmip_g_int_TipCli = 1 Then
         g_str_Parame = g_str_Parame & "ACTECO_CLITDO = " & CStr(moddat_g_int_TipDoc) & " AND "
         g_str_Parame = g_str_Parame & "ACTECO_CLINDO = '" & CStr(moddat_g_str_NumDoc) & "' AND "
      Else
         g_str_Parame = g_str_Parame & "ACTECO_CLITDO = " & CStr(moddat_g_int_CygTDo) & " AND "
         g_str_Parame = g_str_Parame & "ACTECO_CLINDO = '" & CStr(moddat_g_str_CygNDo) & "' AND "
      End If
      
      g_str_Parame = g_str_Parame & "ACTECO_ORDACT = " & CStr(modmip_g_int_OrdAct) & " "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
          Exit Sub
      End If
   End If
   
   'Insertando Actividad Económica
   g_str_Parame = "USP_CLI_ACTECO_AGREGA ("
   
   If modmip_g_int_TipCli = 1 Then
      g_str_Parame = g_str_Parame & CStr(moddat_g_int_TipDoc) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumDoc & "', "
   Else
      g_str_Parame = g_str_Parame & CStr(moddat_g_int_CygTDo) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_CygNDo & "', "
   End If
   
   g_str_Parame = g_str_Parame & CStr(modmip_g_int_OrdAct) & ", "
   g_str_Parame = g_str_Parame & "21, "                                                      'Código Actividad Económica (Independiente)
   
   'Dependiente
   g_str_Parame = g_str_Parame & "0, "                                                       'Tipo DOI Empleador
   g_str_Parame = g_str_Parame & "'', "                                                      'Número DOI Empleador
   g_str_Parame = g_str_Parame & "'', "                                                      'Razón Social
   g_str_Parame = g_str_Parame & "'', "                                                      'Nombre Comercial
   g_str_Parame = g_str_Parame & "0, "                                                       'Tipo Oficina
   g_str_Parame = g_str_Parame & "0, "                                                       'Situación trabajador
   g_str_Parame = g_str_Parame & "0, "                                                       'Tipo de Via
   g_str_Parame = g_str_Parame & "'', "                                                      'Nombre de Vía
   g_str_Parame = g_str_Parame & "'', "                                                      'Número de Vía
   g_str_Parame = g_str_Parame & "'', "                                                      'Interior / Dpto.
   g_str_Parame = g_str_Parame & "0, "                                                       'Tipo de Zona
   g_str_Parame = g_str_Parame & "'', "                                                      'Nombre de Zona
   g_str_Parame = g_str_Parame & "'', "                                                      'Ubicación Geográfica
   g_str_Parame = g_str_Parame & "'', "                                                      'Referencia
   g_str_Parame = g_str_Parame & "'', "                                                      'Teléfono 1
   g_str_Parame = g_str_Parame & "'', "                                                      'Teléfono 2
   g_str_Parame = g_str_Parame & "'', "                                                      'Fax
   g_str_Parame = g_str_Parame & "0, "                                                       'Código CIIU
   g_str_Parame = g_str_Parame & "'', "                                                      'Telefono RR.HH
   g_str_Parame = g_str_Parame & "'', "                                                      'Anexo RR.HH
   g_str_Parame = g_str_Parame & "0, "                                                       'Ingreso Neto
   g_str_Parame = g_str_Parame & "0, "                                                       'Frecuencia de Haberes
   g_str_Parame = g_str_Parame & "0, "                                                       'Fecha de Ingreso
   g_str_Parame = g_str_Parame & "'', "                                                      'Código de Cargo
   g_str_Parame = g_str_Parame & "'', "                                                      'Nombre de Cargo
   g_str_Parame = g_str_Parame & "'', "                                                      'Nombre de Area
   g_str_Parame = g_str_Parame & "'', "                                                      'Número de Anexo
   g_str_Parame = g_str_Parame & "'', "                                                      'Teléfono Directo
   g_str_Parame = g_str_Parame & "'', "                                                      'Celular del Trabajo
   g_str_Parame = g_str_Parame & "'', "                                                      'E-mail del Trabajo
   g_str_Parame = g_str_Parame & "2, "                                                       'Flag de Trabajo Anterior
   g_str_Parame = g_str_Parame & "0, "                                                       'Tipo DOI Empleador Anterior
   g_str_Parame = g_str_Parame & "'', "                                                      'Número DOI Empleador Anterior
   g_str_Parame = g_str_Parame & "'', "                                                      'Razón Social Empleador Anterior
   g_str_Parame = g_str_Parame & "'', "                                                      'Nombre Comercial Empleador Anterior
   g_str_Parame = g_str_Parame & "'', "                                                      'Teléfono 1 Empleador Anterior
   g_str_Parame = g_str_Parame & "'', "                                                      'Teléfono 2 Empleador Anterior
   g_str_Parame = g_str_Parame & "0, "                                                       'Fecha Ingreso Empleador Anterior
   g_str_Parame = g_str_Parame & "0, "                                                       'Fecha Cese Empleador Anterior
   
   'Independiente
   g_str_Parame = g_str_Parame & CStr(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)) & ", "   'Tipo DOI
   g_str_Parame = g_str_Parame & "'" & txt_NumDoc.Text & "', "                            'Número DOI
   g_str_Parame = g_str_Parame & "0, "                                                    'Tipo Vía
   g_str_Parame = g_str_Parame & "'', "                                                   'Nombre Vía
   g_str_Parame = g_str_Parame & "'', "                                                   'Número Vía
   g_str_Parame = g_str_Parame & "'', "                                                   'Interior
   g_str_Parame = g_str_Parame & "0, "                                                    'Tipo Zona
   g_str_Parame = g_str_Parame & "'', "                                                   'Nombre zona
   g_str_Parame = g_str_Parame & "'', "                                                   'Ubicación Geográfica
   g_str_Parame = g_str_Parame & "'', "                                                   'Referencia
   g_str_Parame = g_str_Parame & "'" & txt_Telef1.Text & "', "                            'Teléfono 1
   g_str_Parame = g_str_Parame & "'" & txt_Telef2.Text & "', "                            'Teléfono 2
   g_str_Parame = g_str_Parame & "'" & txt_NumFax.Text & "', "                            'Fax
   
   g_str_Parame = g_str_Parame & CStr(cmb_CodCiu.ItemData(cmb_CodCiu.ListIndex)) & ", "   'CIIU
   g_str_Parame = g_str_Parame & CStr(ipp_IngNet.Value) & ","                             'Ingreso Neto
   g_str_Parame = g_str_Parame & Format(CDate(ipp_IniAct.Text), "yyyymmdd") & ", "        'Inicio de Actividad
   g_str_Parame = g_str_Parame & CStr(cmb_ConLoc.ItemData(cmb_ConLoc.ListIndex)) & ", "   'Contrato Locación
   
   If cmb_ConLoc.ItemData(cmb_ConLoc.ListIndex) = 1 Then
      g_str_Parame = g_str_Parame & CStr(cmb_TDoEmp.ItemData(cmb_TDoEmp.ListIndex)) & ", "   'Tipo DOI Empleador
      g_str_Parame = g_str_Parame & "'" & txt_NDoEmp.Text & "', "                            'Número DOI Empleador
      g_str_Parame = g_str_Parame & "'', "                                                   'Razón Social Empleador
      g_str_Parame = g_str_Parame & "'', "                                                   'Nombre Comercial Empleador
      g_str_Parame = g_str_Parame & "'', "                                                   'Teléfono 1 Empleador
      g_str_Parame = g_str_Parame & "'', "                                                   'Teléfono 2 Empleador
      g_str_Parame = g_str_Parame & Format(CDate(ipp_FecIng.Text), "yyyymmdd") & ", "        'Fecha de Ingreso Empleador
      g_str_Parame = g_str_Parame & "'999999', "                                             'Código Cargo
      g_str_Parame = g_str_Parame & "'" & txt_NomCar.Text & "', "                            'Nombre Cargo
   Else
      g_str_Parame = g_str_Parame & "'', "                                                   'Tipo DOI Empleador
      g_str_Parame = g_str_Parame & "'', "                                                   'Número DOI Empleador
      g_str_Parame = g_str_Parame & "'', "                                                   'Razón Social Empleador
      g_str_Parame = g_str_Parame & "'', "                                                   'Nombre Comercial Empleador
      g_str_Parame = g_str_Parame & "'', "                                                   'Teléfono 1 Empleador
      g_str_Parame = g_str_Parame & "'', "                                                   'Teléfono 2 Empleador
      g_str_Parame = g_str_Parame & "0, "                                                    'Fecha Ingreso Empleador
      g_str_Parame = g_str_Parame & "'', "                                                   'Código Cargo
      g_str_Parame = g_str_Parame & "'', "                                                   'Nombre Cargo
   End If
   
   'Comerciante
   g_str_Parame = g_str_Parame & "0, "                                                    'Tipo DOI
   g_str_Parame = g_str_Parame & "'', "                                                   'Número DOI
   g_str_Parame = g_str_Parame & "'', "                                                   'Razón Social Empleador
   g_str_Parame = g_str_Parame & "'', "                                                   'Nombre Comercial Empleador
   g_str_Parame = g_str_Parame & "0, "                                                    'Tipo Vía
   g_str_Parame = g_str_Parame & "'', "                                                   'Nombre Vía
   g_str_Parame = g_str_Parame & "'', "                                                   'Número Vía
   g_str_Parame = g_str_Parame & "'', "                                                   'Interior
   g_str_Parame = g_str_Parame & "0, "                                                    'Tipo Zona
   g_str_Parame = g_str_Parame & "'', "                                                   'Nombre zona
   g_str_Parame = g_str_Parame & "'', "                                                   'Ubicación Geográfica
   g_str_Parame = g_str_Parame & "'', "                                                   'Referencia
   g_str_Parame = g_str_Parame & "'', "                                                   'Teléfono 1
   g_str_Parame = g_str_Parame & "'', "                                                   'Teléfono 2
   g_str_Parame = g_str_Parame & "'', "                                                   'Fax
   g_str_Parame = g_str_Parame & "0, "                                                    'CIIU
   g_str_Parame = g_str_Parame & "'', "                                                   'Giro Comercial
   g_str_Parame = g_str_Parame & "0, "                                                    'Ingreso Neto
   g_str_Parame = g_str_Parame & "0, "                                                    'Ventas Mensuales
   g_str_Parame = g_str_Parame & "0, "                                                    'Inicio de Operaciones
   g_str_Parame = g_str_Parame & "'', "                                                   'Código Cargo
   g_str_Parame = g_str_Parame & "'', "                                                   'Nombre Cargo
   g_str_Parame = g_str_Parame & "0, "                                                    'Régimen Tributario
   g_str_Parame = g_str_Parame & "0, "                                                    'Porcentaje Participación
   g_str_Parame = g_str_Parame & "0, "                                                    'Tipo Local
   g_str_Parame = g_str_Parame & "0, "                                                    'Alquiler Mensual
   g_str_Parame = g_str_Parame & "'', "                                                   'Nombre Arrendador
   g_str_Parame = g_str_Parame & "'', "                                                   'Teléfono Arrendador
   
   'Accionista
   g_str_Parame = g_str_Parame & "0, "                                                    'Tipo DOI
   g_str_Parame = g_str_Parame & "'', "                                                   'Número DOI
   g_str_Parame = g_str_Parame & "'', "                                                   'Razón Social Empleador
   g_str_Parame = g_str_Parame & "'', "                                                   'Nombre Comercial Empleador
   g_str_Parame = g_str_Parame & "0, "                                                    'Tipo Vía
   g_str_Parame = g_str_Parame & "'', "                                                   'Nombre Vía
   g_str_Parame = g_str_Parame & "'', "                                                   'Número Vía
   g_str_Parame = g_str_Parame & "'', "                                                   'Interior
   g_str_Parame = g_str_Parame & "0, "                                                    'Tipo Zona
   g_str_Parame = g_str_Parame & "'', "                                                   'Nombre zona
   g_str_Parame = g_str_Parame & "'', "                                                   'Ubicación Geográfica
   g_str_Parame = g_str_Parame & "'', "                                                   'Referencia
   g_str_Parame = g_str_Parame & "'', "                                                   'Teléfono 1
   g_str_Parame = g_str_Parame & "'', "                                                   'Teléfono 2
   g_str_Parame = g_str_Parame & "'', "                                                   'Fax
   g_str_Parame = g_str_Parame & "0, "                                                    'CIIU
   g_str_Parame = g_str_Parame & "0, "                                                    'Ingreso Neto
   g_str_Parame = g_str_Parame & "0, "                                                    'Porcentaje Participación
   g_str_Parame = g_str_Parame & "0, "                                                    'Fecha Antigüedad
   
   'Rentista
   g_str_Parame = g_str_Parame & "0, "                                                    'Ingreso Neto
   g_str_Parame = g_str_Parame & "'', "                                                   'Dirección 1
   g_str_Parame = g_str_Parame & "'', "                                                   'Nombre 1
   g_str_Parame = g_str_Parame & "0, "                                                    'Inicio Alquiler 1
   g_str_Parame = g_str_Parame & "'', "                                                   'Teléfono 1 - 1
   g_str_Parame = g_str_Parame & "'', "                                                   'Teléfono 2 - 1
   g_str_Parame = g_str_Parame & "0, "                                                    'Monto Alquiler 1
   g_str_Parame = g_str_Parame & "0, "                                                    'Segunda Propiedad
   g_str_Parame = g_str_Parame & "'', "                                                   'Dirección 2
   g_str_Parame = g_str_Parame & "'', "                                                   'Nombre 2
   g_str_Parame = g_str_Parame & "0, "                                                    'Inicio Alquiler 2
   g_str_Parame = g_str_Parame & "'', "                                                   'Teléfono 1 - 2
   g_str_Parame = g_str_Parame & "'', "                                                   'Teléfono 2 - 2
   g_str_Parame = g_str_Parame & "0, "                                                    'Monto Alquiler 2
   
   'Otros
   g_str_Parame = g_str_Parame & "0, "                                                    'Ingreso Neto
   g_str_Parame = g_str_Parame & "'', "                                                   'Actividad
   g_str_Parame = g_str_Parame & "0, "                                                    'CIIU
   g_str_Parame = g_str_Parame & "'', "                                                   'Observaciones
   
   'Dependiente
   g_str_Parame = g_str_Parame & "0, "                                                    'Moneda Ingresos
   g_str_Parame = g_str_Parame & "'', "                                                   'Dirección
   g_str_Parame = g_str_Parame & "'', "                                                   'Ciudad
   g_str_Parame = g_str_Parame & "'', "                                                   'Código Postal
   
   'Independiente
   g_str_Parame = g_str_Parame & CStr(cmb_MonIng.ItemData(cmb_MonIng.ListIndex)) & ", "               'Moneda Ingresos
   g_str_Parame = g_str_Parame & "'" & txt_Direcc.Text & "', "                                        'Dirección
   g_str_Parame = g_str_Parame & "'" & l_arr_PrvEst(cmb_PrvEst.ListIndex + 1).Genera_Codigo & "', "   'Ciudad
   g_str_Parame = g_str_Parame & "'" & txt_CodPos.Text & "', "                                        'Código Postal
   
   'Comerciante
   g_str_Parame = g_str_Parame & "0, "                                                    'Moneda Ingresos
   g_str_Parame = g_str_Parame & "'', "                                                   'Dirección
   g_str_Parame = g_str_Parame & "'', "                                                   'Ciudad
   g_str_Parame = g_str_Parame & "'', "                                                   'Código Postal
   
   'Accionista
   g_str_Parame = g_str_Parame & "0, "                                                    'Moneda Ingresos
   
   'Rentista
   g_str_Parame = g_str_Parame & "0, "                                                    'Moneda Ingresos
   
   'Otros
   g_str_Parame = g_str_Parame & "0, "                                                    'Moneda Ingresos
   
   'Datos de Auditoria
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
      MsgBox "Error al ejecutar el Procedimiento USP_CLI_ACTECO_AGREGA.", vbCritical, modgen_g_str_NomPlt
      
      Exit Sub
   End If
   
   If modmip_g_int_OrdAct = 1 Then
      'Actualizar en Maestro de Clientes
      g_str_Parame = "USP_CLI_DATGEN_ACTECOPRI ("
      
      If modmip_g_int_TipCli = 1 Then
         g_str_Parame = g_str_Parame & CStr(moddat_g_int_TipDoc) & ", "
         g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumDoc & "', "
      Else
         g_str_Parame = g_str_Parame & CStr(moddat_g_int_CygTDo) & ", "
         g_str_Parame = g_str_Parame & "'" & moddat_g_str_CygNDo & "', "
      End If
      
      g_str_Parame = g_str_Parame & "21, "
      g_str_Parame = g_str_Parame & CStr(cmb_CodCiu.ItemData(cmb_CodCiu.ListIndex)) & ", "   'CIIU
      g_str_Parame = g_str_Parame & "0, "
      g_str_Parame = g_str_Parame & "'', "
      
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "
         
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
         MsgBox "Error al ejecutar el Procedimiento USP_CLI_ACTECO_AGREGA.", vbCritical, modgen_g_str_NomPlt
         
         Exit Sub
      End If
   End If
   
   modmip_g_int_FlgAct_1 = 2
   moddat_g_int_FlgAct = 2
   
   Screen.MousePointer = 0
   
   Unload Me
End Sub

Private Sub cmd_Limpia_Click()
   cmb_TDoEmp.ListIndex = -1
   txt_NDoEmp.Text = ""
   
   Call gs_LimpiaGrid(grd_Listad)
   
   Call fs_Activa_Emp(True)
   
   Call gs_SetFocus(cmb_TDoEmp)
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   If modmip_g_int_TipCli = 1 Then
      pnl_NomCli.Caption = CStr(moddat_g_int_TipDoc) & " - " & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   Else
      pnl_NomCli.Caption = CStr(moddat_g_int_TipDoc) & " - " & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli & " (" & CStr(moddat_g_int_CygTDo) & " - " & moddat_g_str_CygNDo & " / " & moddat_g_str_CygNom & ")"
   End If
   
   Call fs_Inicio
   
   Call fs_Activa_Emp(True)
   Call fs_Limpia
   
   If modmip_g_int_FlgGrb_1 = 2 Then
      g_str_Parame = "SELECT * FROM CLI_ACTECO WHERE "
      
      If modmip_g_int_TipCli = 1 Then
         g_str_Parame = g_str_Parame & "ACTECO_CLITDO = " & CStr(moddat_g_int_TipDoc) & " AND "
         g_str_Parame = g_str_Parame & "ACTECO_CLINDO = '" & CStr(moddat_g_str_NumDoc) & "' AND "
      Else
         g_str_Parame = g_str_Parame & "ACTECO_CLITDO = " & CStr(moddat_g_int_CygTDo) & " AND "
         g_str_Parame = g_str_Parame & "ACTECO_CLINDO = '" & CStr(moddat_g_str_CygNDo) & "' AND "
      End If
      
      g_str_Parame = g_str_Parame & "ACTECO_ORDACT = " & CStr(modmip_g_int_OrdAct) & " "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      Call gs_BuscarCombo_Item(cmb_TipDoc, g_rst_Princi!ActEco_Ind_TipDoc)
      txt_NumDoc.Text = Trim(g_rst_Princi!ActEco_Ind_NumDoc)
      
      txt_Direcc.Text = Trim(g_rst_Princi!ACTECO_IND_EXTDIR & "")
      cmb_PrvEst.ListIndex = gf_Busca_Arregl(l_arr_PrvEst, g_rst_Princi!ACTECO_IND_EXTCIU) - 1
      txt_CodPos.Text = Trim(g_rst_Princi!ACTECO_IND_EXTCPO & "")
      
      txt_Telef1.Text = Trim(g_rst_Princi!ActEco_Ind_Telef1 & "")
      txt_Telef2.Text = Trim(g_rst_Princi!ActEco_Ind_Telef2 & "")
      txt_NumFax.Text = Trim(g_rst_Princi!ActEco_Ind_NumFax & "")
      
      Call gs_BuscarCombo_Item(cmb_CodCiu, g_rst_Princi!ActEco_Ind_CodCiu)
      
      Call gs_BuscarCombo_Item(cmb_MonIng, g_rst_Princi!ActEco_ind_MonIng)
      
      ipp_IngNet.Value = g_rst_Princi!ActEco_Ind_IngNet
      ipp_IniAct.Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Ind_IniAct))
      
      Call gs_BuscarCombo_Item(cmb_ConLoc, g_rst_Princi!ActEco_Ind_ConLoc)
      
      If cmb_ConLoc.ItemData(cmb_ConLoc.ListIndex) = 1 Then
         Call gs_BuscarCombo_Item(cmb_TDoEmp, g_rst_Princi!ActEco_Ind_TipDoc_Emp)
         txt_NDoEmp.Text = Trim(g_rst_Princi!ActEco_Ind_NumDoc_Emp)
         
         modmip_g_int_TDoEmp = CStr(cmb_TDoEmp.ItemData(cmb_TDoEmp.ListIndex))
         modmip_g_str_TDoEmp = cmb_TDoEmp.Text
         modmip_g_str_NDoEmp = Trim(txt_NDoEmp.Text)
         
         Call fs_BusEmp
      
         ipp_FecIng.Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Ind_FecIng_Emp))
         txt_NomCar.Text = Trim(g_rst_Princi!ActEco_Ind_NomCar & "")
      End If
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   End If
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicio()
   grd_Listad.ColWidth(0) = 3105
   grd_Listad.ColWidth(1) = 7995
   
   grd_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipDoc, 1, "232")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TDoEmp, 1, "232")
   
   Call moddat_gs_Carga_CdCIIU(cmb_CodCiu)
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_MonIng, 1, "113")
   Call moddat_gs_Carga_LisIte_Combo(cmb_ConLoc, 1, "214")

   Call modmip_gs_Carga_CiuExt(l_arr_PrvEst, cmb_PrvEst, Format(modmip_g_int_PaiRes, "000000"))
End Sub

Private Sub fs_Limpia()
   cmb_TipDoc.ListIndex = -1
   txt_NumDoc.Text = ""
      
   txt_Direcc.Text = ""
   cmb_PrvEst.ListIndex = -1
   txt_CodPos.Text = ""
   
   txt_Telef1.Text = ""
   txt_Telef2.Text = ""
   txt_NumFax.Text = ""
   
   cmb_CodCiu.ListIndex = -1
   cmb_MonIng.ListIndex = -1
   ipp_IngNet.Value = 0
   ipp_IniAct.Text = Format(date, "dd/mm/yyyy")
   cmb_ConLoc.ListIndex = -1
   
   cmb_TDoEmp.ListIndex = -1
   txt_NDoEmp.Text = ""
   
   Call gs_LimpiaGrid(grd_Listad)
   
   ipp_FecIng.Text = Format(date, "dd/mm/yyyy")
   txt_NomCar.Text = ""
   
   cmb_TDoEmp.Enabled = False
   txt_NDoEmp.Enabled = False
   cmd_Buscar.Enabled = False
   cmd_Limpia.Enabled = False
   cmd_Editar.Enabled = False
   grd_Listad.Enabled = False
   ipp_FecIng.Enabled = False
   txt_NomCar.Enabled = False
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

Private Sub ipp_FecIng_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NomCar)
   End If
End Sub

Private Sub ipp_IngNet_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_IniAct)
   End If
End Sub

Private Sub ipp_IniAct_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_ConLoc)
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

Private Sub txt_NDoEmp_GotFocus()
   Call gs_SelecTodo(txt_NDoEmp)
End Sub

Private Sub txt_NDoEmp_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Buscar)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_NomCar_GotFocus()
   Call gs_SelecTodo(txt_NomCar)
End Sub

Private Sub txt_NomCar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/º")
   End If
End Sub

Private Sub txt_NumDoc_GotFocus()
   Call gs_SelecTodo(txt_NumDoc)
End Sub

Private Sub txt_NumDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Direcc)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO & modgen_g_con_LETRAS & "-")
   End If
End Sub

Private Sub txt_Telef1_GotFocus()
   Call gs_SelecTodo(txt_Telef1)
End Sub

Private Sub txt_Telef1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Telef2)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_Telef2_GotFocus()
   Call gs_SelecTodo(txt_Telef2)
End Sub

Private Sub txt_Telef2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NumFax)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_NumFax_GotFocus()
   Call gs_SelecTodo(txt_NumFax)
End Sub

Private Sub txt_NumFax_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_CodCiu)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub fs_Activa_Emp(ByVal p_Activa As Integer)
   cmb_TDoEmp.Enabled = p_Activa
   txt_NDoEmp.Enabled = p_Activa
   
   cmd_Buscar.Enabled = p_Activa
      
   grd_Listad.Enabled = Not p_Activa
   cmd_Editar.Enabled = Not p_Activa
End Sub

Private Sub fs_BusEmp()
   Dim r_rst_Princi     As ADODB.Recordset
   
   Call gs_LimpiaGrid(grd_Listad)
   
   g_str_Parame = "SELECT * FROM EMP_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_EMPTDO = " & CStr(modmip_g_int_TDoEmp) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_EMPNDO = '" & modmip_g_str_NDoEmp & "' "

   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Princi, 3) Then
      Exit Sub
   End If
      
   Call fs_Activa_Emp(False)
      
   cmd_Editar.Enabled = True
   
   grd_Listad.Redraw = False
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Documento de Identidad"
   
   grd_Listad.Col = 1
   grd_Listad.Text = moddat_gf_Consulta_ParDes("203", CStr(modmip_g_int_TDoEmp)) & " - " & Trim(modmip_g_str_NDoEmp & "")

   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Razón Social"
   
   grd_Listad.Col = 1
   grd_Listad.Text = Trim(r_rst_Princi!DATGEN_RAZSOC & "")

   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Nombre Comercial"

   grd_Listad.Col = 1
   grd_Listad.Text = Trim(r_rst_Princi!DATGEN_NOMCOM & "")
         
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "País"
         
   grd_Listad.Col = 1
   grd_Listad.Text = moddat_gf_Consulta_ParDes("500", Trim(r_rst_Princi!DATGEN_PAIRES))
         
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Dirección"
         
   grd_Listad.Col = 1
   grd_Listad.Text = Trim(r_rst_Princi!DATGEN_EXTDIR & "")
         
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Provincia / Estado"

   grd_Listad.Col = 1
   grd_Listad.Text = modmip_gf_Consulta_CiuExt(r_rst_Princi!DATGEN_PAIRES, Trim(r_rst_Princi!DATGEN_EXTCIU))
   
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Código Postal"
   
   grd_Listad.Col = 1
   grd_Listad.Text = Trim(r_rst_Princi!DATGEN_EXTCPO & "")
         
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Teléfonos"
   
   grd_Listad.Col = 1
   grd_Listad.Text = Trim(r_rst_Princi!DATGEN_TELEF1 & "") & IIf(Len(Trim(Trim(r_rst_Princi!DATGEN_TELEF2 & ""))) > 0, " / " & Trim(r_rst_Princi!DATGEN_TELEF2 & ""), "")
   
   If Len(Trim(r_rst_Princi!DATGEN_NUMFAX & "")) > 0 Then
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      grd_Listad.Col = 0
      grd_Listad.Text = "Fax"
      
      grd_Listad.Col = 1
      grd_Listad.Text = Trim(r_rst_Princi!DATGEN_NUMFAX & "")
   End If
   
   If Len(Trim(r_rst_Princi!DATGEN_TELERH & "")) > 0 Then
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      grd_Listad.Col = 0
      grd_Listad.Text = "Teléfono RR.HH"
      
      grd_Listad.Col = 1
      grd_Listad.Text = Trim(r_rst_Princi!DATGEN_TELERH & "")
   End If
   
   If Len(Trim(r_rst_Princi!DATGEN_ANEXRH & "")) > 0 Then
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      grd_Listad.Col = 0
      grd_Listad.Text = "Anexo RR.HH"
      
      grd_Listad.Col = 1
      grd_Listad.Text = Trim(r_rst_Princi!DATGEN_ANEXRH & "")
   End If
   
   If Len(Trim(r_rst_Princi!DATGEN_PAGWEB & "")) > 0 Then
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      grd_Listad.Col = 0
      grd_Listad.Text = "Página Web"
      
      grd_Listad.Col = 1
      grd_Listad.Text = Trim(r_rst_Princi!DATGEN_PAGWEB & "")
   End If
   
   r_rst_Princi.Close
   Set r_rst_Princi = Nothing
   
   grd_Listad.Redraw = True
   
   Call gs_UbiIniGrid(grd_Listad)
End Sub



