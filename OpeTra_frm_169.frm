VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_MntCli_59 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6570
   ClientLeft      =   6525
   ClientTop       =   3525
   ClientWidth     =   11655
   Icon            =   "OpeTra_frm_169.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   11655
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   6585
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   11685
      _Version        =   65536
      _ExtentX        =   20611
      _ExtentY        =   11615
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
         Height          =   3735
         Left            =   30
         TabIndex        =   27
         Top             =   1950
         Width           =   11595
         _Version        =   65536
         _ExtentX        =   20452
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
         Begin VB.ComboBox cmb_CodCiu 
            Height          =   315
            Left            =   2010
            TabIndex        =   17
            Text            =   "cmb_DptDir"
            Top             =   3030
            Width           =   9525
         End
         Begin VB.TextBox txt_NomCom 
            Height          =   315
            Left            =   2010
            MaxLength       =   250
            TabIndex        =   3
            Text            =   "Text1"
            Top             =   720
            Width           =   9525
         End
         Begin VB.TextBox txt_Telef2 
            Height          =   315
            Left            =   3660
            MaxLength       =   12
            TabIndex        =   15
            Text            =   "Text1"
            Top             =   2700
            Width           =   1640
         End
         Begin VB.ComboBox cmb_TipVia 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   1050
            Width           =   3315
         End
         Begin VB.TextBox txt_NumVia 
            Height          =   315
            Left            =   8220
            MaxLength       =   15
            TabIndex        =   6
            Text            =   "Text1"
            Top             =   1380
            Width           =   1640
         End
         Begin VB.TextBox txt_IntDpt 
            Height          =   315
            Left            =   9870
            MaxLength       =   15
            TabIndex        =   7
            Text            =   "Text1"
            Top             =   1380
            Width           =   1665
         End
         Begin VB.TextBox txt_NomZon 
            Height          =   315
            Left            =   8220
            MaxLength       =   120
            TabIndex        =   9
            Text            =   "Text1"
            Top             =   1710
            Width           =   3315
         End
         Begin VB.ComboBox cmb_PrvDir 
            Height          =   315
            Left            =   8220
            TabIndex        =   11
            Text            =   "cmb_PrvDir"
            Top             =   2040
            Width           =   3315
         End
         Begin VB.TextBox txt_Refere 
            Height          =   315
            Left            =   8220
            MaxLength       =   250
            TabIndex        =   13
            Text            =   "Text1"
            Top             =   2370
            Width           =   3315
         End
         Begin VB.TextBox txt_NumFax 
            Height          =   315
            Left            =   8220
            MaxLength       =   12
            TabIndex        =   16
            Text            =   "Text1"
            Top             =   2700
            Width           =   1640
         End
         Begin VB.TextBox txt_RazSoc 
            Height          =   315
            Left            =   2010
            MaxLength       =   250
            TabIndex        =   2
            Text            =   "Text1"
            Top             =   390
            Width           =   9525
         End
         Begin VB.ComboBox cmb_TipDoc 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   3315
         End
         Begin VB.TextBox txt_NumDoc 
            Height          =   315
            Left            =   8220
            MaxLength       =   11
            TabIndex        =   1
            Text            =   "Text1"
            Top             =   60
            Width           =   3315
         End
         Begin VB.TextBox txt_NomVia 
            Height          =   315
            Left            =   2010
            MaxLength       =   120
            TabIndex        =   5
            Text            =   "Text1"
            Top             =   1380
            Width           =   3315
         End
         Begin VB.ComboBox cmb_TipZon 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   1710
            Width           =   3315
         End
         Begin VB.ComboBox cmb_DptDir 
            Height          =   315
            Left            =   2010
            TabIndex        =   10
            Text            =   "cmb_DptDir"
            Top             =   2040
            Width           =   3315
         End
         Begin VB.ComboBox cmb_DstDir 
            Height          =   315
            Left            =   2010
            TabIndex        =   12
            Text            =   "cmb_DstDir"
            Top             =   2370
            Width           =   3315
         End
         Begin VB.TextBox txt_Telef1 
            Height          =   315
            Left            =   2010
            MaxLength       =   12
            TabIndex        =   14
            Text            =   "Text1"
            Top             =   2700
            Width           =   1640
         End
         Begin VB.ComboBox cmb_GirCom 
            Height          =   315
            Left            =   2010
            TabIndex        =   18
            Text            =   "cmb_DptDir"
            Top             =   3360
            Width           =   9525
         End
         Begin VB.CommandButton cmd_DirCas 
            Caption         =   "="
            Height          =   315
            Left            =   1530
            TabIndex        =   25
            ToolTipText     =   "Obtener Dirección de Domicilio"
            Top             =   1050
            Width           =   435
         End
         Begin VB.Label lbl_General 
            Caption         =   "Nombre Comercial:"
            Height          =   285
            Index           =   49
            Left            =   90
            TabIndex        =   44
            Top             =   720
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Tipo de Vía:"
            Height          =   285
            Index           =   41
            Left            =   90
            TabIndex        =   43
            Top             =   1050
            Width           =   1545
         End
         Begin VB.Label lbl_General 
            Caption         =   "Nro - Int/Dpto/Mza/Lote:"
            Height          =   285
            Index           =   51
            Left            =   6210
            TabIndex        =   42
            Top             =   1380
            Width           =   1935
         End
         Begin VB.Label lbl_General 
            Caption         =   "Nombre Zona:"
            Height          =   285
            Index           =   52
            Left            =   6210
            TabIndex        =   41
            Top             =   1710
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Provincia:"
            Height          =   315
            Index           =   53
            Left            =   6210
            TabIndex        =   40
            Top             =   2040
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Referencia:"
            Height          =   285
            Index           =   54
            Left            =   6210
            TabIndex        =   39
            Top             =   2370
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Fax:"
            Height          =   285
            Index           =   55
            Left            =   6210
            TabIndex        =   38
            Top             =   2700
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Nombre Vía:"
            Height          =   285
            Index           =   42
            Left            =   90
            TabIndex        =   37
            Top             =   1380
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Tipo de Zona:"
            Height          =   315
            Index           =   43
            Left            =   90
            TabIndex        =   36
            Top             =   1710
            Width           =   1455
         End
         Begin VB.Label lbl_General 
            Caption         =   "Departamento:"
            Height          =   315
            Index           =   44
            Left            =   90
            TabIndex        =   35
            Top             =   2040
            Width           =   1425
         End
         Begin VB.Label lbl_General 
            Caption         =   "Distrito:"
            Height          =   315
            Index           =   45
            Left            =   90
            TabIndex        =   34
            Top             =   2370
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Teléfono (s):"
            Height          =   285
            Index           =   46
            Left            =   90
            TabIndex        =   33
            Top             =   2700
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Tipo Docum. Ident.:"
            Height          =   285
            Index           =   36
            Left            =   90
            TabIndex        =   32
            Top             =   60
            Width           =   1635
         End
         Begin VB.Label lbl_General 
            Caption         =   "Número Docum. Ident.:"
            Height          =   285
            Index           =   48
            Left            =   6210
            TabIndex        =   31
            Top             =   60
            Width           =   1845
         End
         Begin VB.Label lbl_General 
            Caption         =   "Razón Social:"
            Height          =   285
            Index           =   37
            Left            =   90
            TabIndex        =   30
            Top             =   390
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "CIIU:"
            Height          =   285
            Index           =   39
            Left            =   90
            TabIndex        =   29
            Top             =   3030
            Width           =   1365
         End
         Begin VB.Label lbl_General 
            Caption         =   "Giro Comercial:"
            Height          =   285
            Index           =   7
            Left            =   90
            TabIndex        =   28
            Top             =   3360
            Width           =   1365
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   45
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
            Left            =   660
            TabIndex        =   50
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
            Left            =   660
            TabIndex        =   51
            Top             =   330
            Width           =   6165
            _Version        =   65536
            _ExtentX        =   10874
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "Actividades Económicas - Comerciante"
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
            Picture         =   "OpeTra_frm_169.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   435
         Left            =   30
         TabIndex        =   46
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
            TabIndex        =   47
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
            TabIndex        =   48
            Top             =   60
            Width           =   1725
         End
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   675
         Left            =   30
         TabIndex        =   49
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
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_169.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10980
            Picture         =   "OpeTra_frm_169.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   795
         Left            =   30
         TabIndex        =   52
         Top             =   5730
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
         Begin VB.CommandButton cmd_BusEmp_Ant 
            Caption         =   "..."
            Height          =   315
            Left            =   10620
            TabIndex        =   53
            ToolTipText     =   "Obtener Dirección de Domicilio"
            Top             =   6600
            Width           =   435
         End
         Begin VB.TextBox txt_NomCar 
            Height          =   315
            Left            =   8220
            MaxLength       =   250
            TabIndex        =   22
            Text            =   "Text1"
            Top             =   420
            Width           =   3315
         End
         Begin VB.ComboBox cmb_MonIng 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   90
            Width           =   3315
         End
         Begin EditLib.fpDoubleSingle ipp_IngNet 
            Height          =   315
            Left            =   8220
            TabIndex        =   20
            Top             =   90
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
         Begin EditLib.fpDateTime ipp_FecIng 
            Height          =   315
            Left            =   2010
            TabIndex        =   21
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
         Begin Threed.SSPanel pnl_FlgEmp_Ant 
            Height          =   315
            Left            =   11100
            TabIndex        =   54
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
            Caption         =   "Cargo:"
            Height          =   285
            Index           =   57
            Left            =   6210
            TabIndex        =   58
            Top             =   420
            Width           =   1665
         End
         Begin VB.Label lbl_General 
            Caption         =   "Fecha de Ingreso:"
            Height          =   315
            Index           =   58
            Left            =   90
            TabIndex        =   57
            Top             =   420
            Width           =   1365
         End
         Begin VB.Label lbl_General 
            Caption         =   "Ingreso Declarado:"
            Height          =   285
            Index           =   61
            Left            =   6210
            TabIndex        =   56
            Top             =   90
            Width           =   1755
         End
         Begin VB.Label lbl_General 
            Caption         =   "Moneda de Ingresos:"
            Height          =   285
            Index           =   0
            Left            =   90
            TabIndex        =   55
            Top             =   90
            Width           =   1665
         End
      End
   End
End
Attribute VB_Name = "frm_MntCli_59"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_GirCom()   As moddat_tpo_Genera
Dim l_str_GirCom     As String
Dim l_str_CodCiu     As String
Dim l_str_DptDir     As String
Dim l_str_PrvDir     As String
Dim l_str_DstDir     As String
Dim l_int_FlgCmb     As Integer

Private Sub cmb_CodCiu_Change()
   l_str_CodCiu = cmb_CodCiu.Text
End Sub

Private Sub cmb_CodCiu_Click()
   If cmb_CodCiu.ListIndex > -1 Then
      If l_int_FlgCmb Then
         Call gs_SetFocus(cmb_GirCom)
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
      
      Call gs_SetFocus(cmb_GirCom)
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

Private Sub cmb_GirCom_Change()
   l_str_GirCom = cmb_GirCom.Text
End Sub

Private Sub cmb_GirCom_Click()
   If cmb_GirCom.ListIndex > -1 Then
      If l_int_FlgCmb Then
         Call gs_SetFocus(cmb_MonIng)
      End If
   End If
End Sub

Private Sub cmb_GirCom_GotFocus()
   l_int_FlgCmb = True
   l_str_GirCom = cmb_GirCom.Text
End Sub

Private Sub cmb_GirCom_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS + modgen_g_con_NUMERO + "-_ ./*+#,()" + Chr(34))
   Else
      l_int_FlgCmb = False
      Call gs_BuscarCombo(cmb_GirCom, l_str_GirCom)
      l_int_FlgCmb = True
      
      If cmb_GirCom.ListIndex > -1 Then
         l_str_GirCom = ""
      End If
      
      Call gs_SetFocus(cmb_MonIng)
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

Private Sub cmb_TipDoc_Click()
   Call gs_SetFocus(txt_NumDoc)
End Sub

Private Sub cmb_TipDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipDoc_Click
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

Private Sub cmd_DirCas_Click()
   cmb_TipVia.ListIndex = frm_MntCli_52.cmb_TipVia.ListIndex
   txt_NomVia.Text = frm_MntCli_52.txt_NomVia.Text
   txt_NumVia.Text = frm_MntCli_52.txt_NumVia.Text
   txt_IntDpt.Text = frm_MntCli_52.txt_IntDpt.Text
   cmb_TipZon.ListIndex = frm_MntCli_52.cmb_TipZon.ListIndex
   txt_NomZon.Text = frm_MntCli_52.txt_NomZon.Text
   
   cmb_DptDir.ListIndex = frm_MntCli_52.cmb_DptDir.ListIndex
   
   Call moddat_gs_Carga_Provin(cmb_PrvDir, Format(cmb_DptDir.ItemData(cmb_DptDir.ListIndex), "00"))
   cmb_PrvDir.ListIndex = frm_MntCli_52.cmb_PrvDir.ListIndex
         
   Call moddat_gs_Carga_Distri(cmb_DstDir, Format(cmb_DptDir.ItemData(cmb_DptDir.ListIndex), "00"), Format(cmb_PrvDir.ItemData(cmb_PrvDir.ListIndex), "00"))
   cmb_DstDir.ListIndex = frm_MntCli_52.cmb_DstDir.ListIndex
   
   txt_Refere.Text = frm_MntCli_52.txt_Refere.Text
   
   txt_Telef1.Text = frm_MntCli_52.txt_TelFij.Text
   
   Call gs_SetFocus(cmb_CodCiu)
End Sub

Private Sub cmd_Grabar_Click()
   If cmb_TipDoc.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Documento.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipDoc)
      Exit Sub
   End If
   
   If cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) = 7 Then
      If Len(Trim(txt_NumDoc.Text)) <> 11 Then
         MsgBox "El Número de Ruc debe tener 11 digitos.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NumDoc)
         Exit Sub
      End If
      If Not gf_Valida_RUC(Trim(txt_NumDoc.Text), Mid(Trim(txt_NumDoc.Text), Len(Trim(txt_NumDoc.Text)), 1)) Then
         MsgBox "El Número de RUC no es valido.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NumDoc)
         Exit Sub
      End If
   Else
      If Len(Trim(txt_NumDoc.Text)) = 0 Then
         MsgBox "Debe ingresar el Número de Documento.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NumDoc)
         Exit Sub
      End If
   End If
   
   If Len(Trim(txt_RazSoc.Text)) = 0 Then
      MsgBox "Debe ingresar la Razón Social.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_RazSoc)
      Exit Sub
   End If
   
   If Len(Trim(txt_NomCom.Text)) = 0 Then
      MsgBox "Debe ingresar el Nombre Comercial.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NomCom)
      Exit Sub
   End If
   
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
   
   If cmb_CodCiu.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Código de CIIU.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_CodCiu)
      Exit Sub
   End If
   
   If cmb_GirCom.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Giro Comercial.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_GirCom)
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
   
   If Len(Trim(txt_NomCar.Text)) = 0 Then
      MsgBox "Debe ingresar el Cargo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NomCar)
      Exit Sub
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
   g_str_Parame = g_str_Parame & "31, "                                                      'Código Actividad Económica (Comerciante)
   
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
   g_str_Parame = g_str_Parame & "0, "                                                    'Tipo DOI
   g_str_Parame = g_str_Parame & "'', "                                                   'Número DOI
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
   g_str_Parame = g_str_Parame & "0, "                                                    'Inicio de Actividad
   g_str_Parame = g_str_Parame & "0, "                                                    'Contrato Locación
   g_str_Parame = g_str_Parame & "0, "                                                    'Tipo DOI Empleador
   g_str_Parame = g_str_Parame & "'', "                                                   'Número DOI Empleador
   g_str_Parame = g_str_Parame & "'', "                                                   'Razón Social Empleador
   g_str_Parame = g_str_Parame & "'', "                                                   'Nombre Comercial Empleador
   g_str_Parame = g_str_Parame & "'', "                                                   'Teléfono 1 Empleador
   g_str_Parame = g_str_Parame & "'', "                                                   'Teléfono 2 Empleador
   g_str_Parame = g_str_Parame & "0, "                                                    'Fecha Ingreso Empleador
   g_str_Parame = g_str_Parame & "'', "                                                   'Código Cargo
   g_str_Parame = g_str_Parame & "'', "                                                   'Nombre Cargo
   
   'Comerciante
   g_str_Parame = g_str_Parame & CStr(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)) & ", "   'Tipo DOI
   g_str_Parame = g_str_Parame & "'" & txt_NumDoc.Text & "', "                            'Número DOI
   g_str_Parame = g_str_Parame & "'" & txt_RazSoc.Text & "', "                            'Razón Social
   g_str_Parame = g_str_Parame & "'" & txt_NomCom.Text & "', "                            'Nombre Comercial
   g_str_Parame = g_str_Parame & CStr(cmb_TipVia.ItemData(cmb_TipVia.ListIndex)) & ", "   'Tipo Vía
   g_str_Parame = g_str_Parame & "'" & txt_NomVia.Text & "', "                            'Nombre Vía
   g_str_Parame = g_str_Parame & "'" & txt_NumVia.Text & "', "                            'Número Vía
   g_str_Parame = g_str_Parame & "'" & txt_IntDpt.Text & "', "                            'Interior Dpto
   g_str_Parame = g_str_Parame & CStr(cmb_TipZon.ItemData(cmb_TipZon.ListIndex)) & ", "   'Tipo Zona
   g_str_Parame = g_str_Parame & "'" & txt_NomZon.Text & "', "                            'Nombre Zona
   
   g_str_Parame = g_str_Parame & "'" & Format(cmb_DptDir.ItemData(cmb_DptDir.ListIndex), "00") & Format(cmb_PrvDir.ItemData(cmb_PrvDir.ListIndex), "00") & Format(cmb_DstDir.ItemData(cmb_DstDir.ListIndex), "00") & "', "
   
   g_str_Parame = g_str_Parame & "'" & txt_Refere.Text & "', "                            'Referencia
   g_str_Parame = g_str_Parame & "'" & txt_Telef1.Text & "', "                            'Teléfono 1
   g_str_Parame = g_str_Parame & "'" & txt_Telef2.Text & "', "                            'Teléfono 2
   g_str_Parame = g_str_Parame & "'" & txt_NumFax.Text & "', "                            'Fax
   
   g_str_Parame = g_str_Parame & CStr(cmb_CodCiu.ItemData(cmb_CodCiu.ListIndex)) & ", "   'CIIU
   
   g_str_Parame = g_str_Parame & "'" & l_arr_GirCom(cmb_GirCom.ListIndex + 1).Genera_Codigo & "', "      'Giro comercial
   g_str_Parame = g_str_Parame & CStr(ipp_IngNet.Value) & ","                             'Ingreso Neto
   g_str_Parame = g_str_Parame & "0, "                                                    'Ventas Mensuales
   g_str_Parame = g_str_Parame & Format(CDate(ipp_FecIng.Text), "yyyymmdd") & ", "        'Fecha de Inicio Operaciones
   g_str_Parame = g_str_Parame & "'999999', "                                             'Código Cargo
   g_str_Parame = g_str_Parame & "'" & txt_NomCar.Text & "', "                            'Nombre Cargo
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
   g_str_Parame = g_str_Parame & "0, "                                                    'Moneda Ingresos
   g_str_Parame = g_str_Parame & "'', "                                                   'Dirección
   g_str_Parame = g_str_Parame & "'', "                                                   'Ciudad
   g_str_Parame = g_str_Parame & "'', "                                                   'Código Postal
   
   'Comerciante
   g_str_Parame = g_str_Parame & CStr(cmb_MonIng.ItemData(cmb_MonIng.ListIndex)) & ", "   'Moneda Ingresos
   g_str_Parame = g_str_Parame & "'', "                                                   'Dirección
   g_str_Parame = g_str_Parame & "'', "                                                   'Ciudad
   g_str_Parame = g_str_Parame & "'', "                                                   'Código Postal
   
   'Accionista
   g_str_Parame = g_str_Parame & "0, "                                                    'Moneda Ingresos
   
   'Rentista
   g_str_Parame = g_str_Parame & "0, "                                                    'Moneda Ingresos
   
   'Otros
   g_str_Parame = g_str_Parame & "0, "                                                    'Moneda Ingresos
   
   'Dependiente
   g_str_Parame = g_str_Parame & "0, "    'TIPO IDENTIFICACION(ACTECO_IND_TIPIDE)
   
   'Independiente
   g_str_Parame = g_str_Parame & "0, "   'TIPO IDENTIFICACION(ACTECO_DEP_TIPIDE)
   
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
      
      g_str_Parame = g_str_Parame & "31, "
      g_str_Parame = g_str_Parame & CStr(cmb_CodCiu.ItemData(cmb_CodCiu.ListIndex)) & ", "   'CIIU
      g_str_Parame = g_str_Parame & CStr(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)) & ", "   'Tipo DOI
      g_str_Parame = g_str_Parame & "'" & txt_NumDoc.Text & "', "                            'Número DOI
      
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
      
      Call gs_BuscarCombo_Item(cmb_TipDoc, g_rst_Princi!ActEco_Com_TipDoc)
      txt_NumDoc.Text = Trim(g_rst_Princi!ActEco_Com_NumDoc)
      
      txt_RazSoc.Text = Trim(g_rst_Princi!ActEco_Com_RazSoc)
      txt_NomCom.Text = Trim(g_rst_Princi!ActEco_Com_NomCom)
      
      Call gs_BuscarCombo_Item(cmb_TipVia, g_rst_Princi!ActEco_Com_TipVia)
      
      txt_NomVia.Text = Trim(g_rst_Princi!ActEco_Com_NomVia & "")
      txt_NumVia.Text = Trim(g_rst_Princi!ActEco_Com_NumVia & "")
      txt_IntDpt.Text = Trim(g_rst_Princi!ActEco_Com_IntDpt & "")
      
      Call gs_BuscarCombo_Item(cmb_TipZon, g_rst_Princi!ActEco_Com_TipZon)
      txt_NomZon.Text = Trim(g_rst_Princi!ActEco_Com_NomZon & "")
   
      Call gs_BuscarCombo_Item(cmb_DptDir, CInt(Left(g_rst_Princi!ActEco_Com_UbiGeo, 2)))
      Call moddat_gs_Carga_Provin(cmb_PrvDir, Left(g_rst_Princi!ActEco_Com_UbiGeo, 2))
      Call gs_BuscarCombo_Item(cmb_PrvDir, CInt(Mid(g_rst_Princi!ActEco_Com_UbiGeo, 3, 2)))
      Call moddat_gs_Carga_Distri(cmb_DstDir, Left(g_rst_Princi!ActEco_Com_UbiGeo, 2), Mid(g_rst_Princi!ActEco_Com_UbiGeo, 3, 2))
      Call gs_BuscarCombo_Item(cmb_DstDir, CInt(Right(g_rst_Princi!ActEco_Com_UbiGeo, 2)))
      
      txt_Refere.Text = Trim(g_rst_Princi!ActEco_Com_Refere & "")
      txt_Telef1.Text = Trim(g_rst_Princi!ActEco_Com_Telef1 & "")
      txt_Telef2.Text = Trim(g_rst_Princi!ActEco_Com_Telef2 & "")
      txt_NumFax.Text = Trim(g_rst_Princi!ActEco_Com_NumFax & "")
      
      Call gs_BuscarCombo_Item(cmb_CodCiu, g_rst_Princi!ActEco_Com_CodCiu)
      
      cmb_GirCom.ListIndex = gf_Busca_Arregl(l_arr_GirCom, g_rst_Princi!ActEco_Com_GirCom) - 1
      
      Call gs_BuscarCombo_Item(cmb_MonIng, g_rst_Princi!ActEco_com_MonIng)
      
      ipp_IngNet.Value = g_rst_Princi!ActEco_Com_IngNet
      ipp_FecIng.Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Com_IniOpe))
      
      txt_NomCar.Text = Trim(g_rst_Princi!ActEco_Com_NomCar & "")
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   End If
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicio()
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipDoc, 1, "232")
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipVia, 1, "201")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipZon, 1, "202")
   
   Call moddat_gs_Carga_Depart(cmb_DptDir)
   Call moddat_gs_Carga_CdCIIU(cmb_CodCiu)
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_MonIng, 1, "113")
   Call moddat_gs_Carga_GirCom(cmb_GirCom, l_arr_GirCom())
End Sub

Private Sub fs_Limpia()
   cmb_TipDoc.ListIndex = -1
   txt_NumDoc.Text = ""
   txt_RazSoc.Text = ""
   txt_NomCom.Text = ""
   
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
   txt_Telef1.Text = ""
   txt_Telef2.Text = ""
   txt_NumFax.Text = ""
   
   cmb_CodCiu.ListIndex = -1
   cmb_GirCom.ListIndex = -1
   cmb_MonIng.ListIndex = -1
   ipp_IngNet.Value = 0
   ipp_FecIng.Text = Format(date, "dd/mm/yyyy")
   txt_NomCar.Text = ""
End Sub

Private Sub ipp_FecIng_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NomCar)
   End If
End Sub

Private Sub ipp_IngNet_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecIng)
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

Private Sub txt_NomCom_GotFocus()
   Call gs_SelecTodo(txt_NomCom)
End Sub

Private Sub txt_NomCom_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipVia)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "- =)(/&%$#@.;_,:")
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

Private Sub txt_RazSoc_GotFocus()
   Call gs_SelecTodo(txt_RazSoc)
End Sub

Private Sub txt_RazSoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NomCom)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "- =)(/&%$#@.;_,:")
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

Private Sub txt_NumDoc_GotFocus()
   Call gs_SelecTodo(txt_NumDoc)
End Sub

Private Sub txt_NumDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_RazSoc)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
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


