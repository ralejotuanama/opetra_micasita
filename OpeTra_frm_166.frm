VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_MntCli_57 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7800
   ClientLeft      =   4335
   ClientTop       =   2160
   ClientWidth     =   11655
   Icon            =   "OpeTra_frm_166.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   11655
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   7815
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   11685
      _Version        =   65536
      _ExtentX        =   20611
      _ExtentY        =   13785
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
         Height          =   1065
         Left            =   30
         TabIndex        =   30
         Top             =   6690
         Width           =   11595
         _Version        =   65536
         _ExtentX        =   20452
         _ExtentY        =   1879
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
            Left            =   1980
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   360
            Width           =   3315
         End
         Begin VB.TextBox txt_TelDir 
            Height          =   315
            Left            =   1980
            MaxLength       =   25
            TabIndex        =   21
            Text            =   "Text1"
            Top             =   60
            Width           =   1640
         End
         Begin VB.TextBox txt_NomCar 
            Height          =   315
            Left            =   8190
            MaxLength       =   250
            TabIndex        =   26
            Text            =   "Text1"
            Top             =   720
            Width           =   3315
         End
         Begin VB.TextBox txt_NumAnx 
            Height          =   315
            Left            =   8190
            MaxLength       =   5
            TabIndex        =   22
            Text            =   "Text1"
            Top             =   60
            Width           =   1640
         End
         Begin VB.CommandButton cmd_BusEmp_Ant 
            Caption         =   "..."
            Height          =   315
            Left            =   10620
            TabIndex        =   31
            ToolTipText     =   "Obtener Direcci�n de Domicilio"
            Top             =   6600
            Width           =   435
         End
         Begin EditLib.fpDoubleSingle ipp_IngNet 
            Height          =   315
            Left            =   8190
            TabIndex        =   24
            Top             =   390
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
            Left            =   1980
            TabIndex        =   25
            Top             =   690
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
            TabIndex        =   32
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
            Caption         =   "Moneda de Ingresos:"
            Height          =   285
            Index           =   0
            Left            =   60
            TabIndex        =   64
            Top             =   360
            Width           =   1665
         End
         Begin VB.Label lbl_General 
            Caption         =   "Tel�fono Directo:"
            Height          =   285
            Index           =   59
            Left            =   60
            TabIndex        =   63
            Top             =   60
            Width           =   1575
         End
         Begin VB.Label lbl_General 
            Caption         =   "Ingreso Declarado:"
            Height          =   285
            Index           =   61
            Left            =   6180
            TabIndex        =   36
            Top             =   390
            Width           =   1755
         End
         Begin VB.Label lbl_General 
            Caption         =   "Fecha de Ingreso:"
            Height          =   315
            Index           =   58
            Left            =   60
            TabIndex        =   35
            Top             =   690
            Width           =   1365
         End
         Begin VB.Label lbl_General 
            Caption         =   "Anexo:"
            Height          =   285
            Index           =   64
            Left            =   6180
            TabIndex        =   34
            Top             =   60
            Width           =   1575
         End
         Begin VB.Label lbl_General 
            Caption         =   "Cargo:"
            Height          =   285
            Index           =   57
            Left            =   6180
            TabIndex        =   33
            Top             =   720
            Width           =   1665
         End
      End
      Begin Threed.SSPanel SSPanel11 
         Height          =   2415
         Left            =   30
         TabIndex        =   49
         Top             =   4230
         Width           =   11595
         _Version        =   65536
         _ExtentX        =   20452
         _ExtentY        =   4260
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
         Begin VB.ComboBox cmb_Tipo 
            Height          =   315
            Left            =   8190
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   60
            Width           =   3315
         End
         Begin VB.TextBox txt_Telef1 
            Height          =   315
            Left            =   1980
            MaxLength       =   25
            TabIndex        =   18
            Text            =   "Text1"
            Top             =   2040
            Width           =   1640
         End
         Begin VB.ComboBox cmb_DstDir 
            Height          =   315
            Left            =   1980
            TabIndex        =   16
            Text            =   "cmb_DstDir"
            Top             =   1710
            Width           =   3315
         End
         Begin VB.ComboBox cmb_DptDir 
            Height          =   315
            Left            =   1980
            TabIndex        =   14
            Text            =   "cmb_DptDir"
            Top             =   1380
            Width           =   3315
         End
         Begin VB.ComboBox cmb_TipZon 
            Height          =   315
            Left            =   1980
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   1050
            Width           =   3315
         End
         Begin VB.TextBox txt_NomVia 
            Height          =   315
            Left            =   1980
            MaxLength       =   120
            TabIndex        =   9
            Text            =   "Text1"
            Top             =   720
            Width           =   3315
         End
         Begin VB.TextBox txt_NumFax 
            Height          =   315
            Left            =   8190
            MaxLength       =   25
            TabIndex        =   20
            Text            =   "Text1"
            Top             =   2040
            Width           =   1640
         End
         Begin VB.TextBox txt_Refere 
            Height          =   315
            Left            =   8190
            MaxLength       =   250
            TabIndex        =   17
            Text            =   "Text1"
            Top             =   1710
            Width           =   3315
         End
         Begin VB.ComboBox cmb_PrvDir 
            Height          =   315
            Left            =   8190
            TabIndex        =   15
            Text            =   "cmb_PrvDir"
            Top             =   1380
            Width           =   3315
         End
         Begin VB.TextBox txt_NomZon 
            Height          =   315
            Left            =   8190
            MaxLength       =   120
            TabIndex        =   13
            Text            =   "Text1"
            Top             =   1050
            Width           =   3315
         End
         Begin VB.TextBox txt_IntDpt 
            Height          =   315
            Left            =   9840
            MaxLength       =   30
            TabIndex        =   11
            Text            =   "Text1"
            Top             =   720
            Width           =   1665
         End
         Begin VB.TextBox txt_NumVia 
            Height          =   315
            Left            =   8190
            MaxLength       =   30
            TabIndex        =   10
            Text            =   "Text1"
            Top             =   720
            Width           =   1640
         End
         Begin VB.ComboBox cmb_TipVia 
            Height          =   315
            Left            =   8190
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   390
            Width           =   3315
         End
         Begin VB.TextBox txt_Telef2 
            Height          =   315
            Left            =   3630
            MaxLength       =   25
            TabIndex        =   19
            Text            =   "Text1"
            Top             =   2040
            Width           =   1640
         End
         Begin VB.ComboBox cmb_TipOfi 
            Height          =   315
            Left            =   1980
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   390
            Width           =   3315
         End
         Begin VB.ComboBox cmb_SitTra 
            Height          =   315
            Left            =   1980
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   60
            Width           =   3315
         End
         Begin VB.Label lbl_General 
            Caption         =   "Tipo:"
            Height          =   285
            Index           =   1
            Left            =   6180
            TabIndex        =   65
            Top             =   60
            Width           =   1515
         End
         Begin VB.Label lbl_General 
            Caption         =   "Tel�fono (s):"
            Height          =   285
            Index           =   46
            Left            =   60
            TabIndex        =   62
            Top             =   2040
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Distrito:"
            Height          =   315
            Index           =   45
            Left            =   60
            TabIndex        =   61
            Top             =   1710
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Departamento:"
            Height          =   315
            Index           =   44
            Left            =   60
            TabIndex        =   60
            Top             =   1380
            Width           =   1425
         End
         Begin VB.Label lbl_General 
            Caption         =   "Tipo de Zona:"
            Height          =   315
            Index           =   43
            Left            =   60
            TabIndex        =   59
            Top             =   1050
            Width           =   1455
         End
         Begin VB.Label lbl_General 
            Caption         =   "Nombre V�a:"
            Height          =   285
            Index           =   42
            Left            =   60
            TabIndex        =   58
            Top             =   720
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Fax:"
            Height          =   285
            Index           =   55
            Left            =   6180
            TabIndex        =   57
            Top             =   2040
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Referencia:"
            Height          =   285
            Index           =   54
            Left            =   6180
            TabIndex        =   56
            Top             =   1710
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Provincia:"
            Height          =   315
            Index           =   53
            Left            =   6180
            TabIndex        =   55
            Top             =   1380
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Nombre Zona:"
            Height          =   285
            Index           =   52
            Left            =   6180
            TabIndex        =   54
            Top             =   1050
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Nro - Int/Dpto/Mza/Lote:"
            Height          =   285
            Index           =   51
            Left            =   6180
            TabIndex        =   53
            Top             =   720
            Width           =   1935
         End
         Begin VB.Label lbl_General 
            Caption         =   "Tipo de Oficina:"
            Height          =   285
            Index           =   40
            Left            =   60
            TabIndex        =   52
            Top             =   390
            Width           =   1485
         End
         Begin VB.Label lbl_General 
            Caption         =   "Tipo de V�a:"
            Height          =   285
            Index           =   41
            Left            =   6180
            TabIndex        =   51
            Top             =   390
            Width           =   1545
         End
         Begin VB.Label lbl_General 
            Caption         =   "Situaci�n Trabajador:"
            Height          =   285
            Index           =   8
            Left            =   60
            TabIndex        =   50
            Top             =   60
            Width           =   1515
         End
      End
      Begin Threed.SSPanel SSPanel10 
         Height          =   1425
         Left            =   30
         TabIndex        =   47
         Top             =   2760
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
            TabIndex        =   48
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
         Top             =   1950
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
            Picture         =   "OpeTra_frm_166.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   60
            Width           =   585
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   10380
            Picture         =   "OpeTra_frm_166.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Limpiar Datos"
            Top             =   60
            Width           =   585
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   585
            Left            =   9780
            Picture         =   "OpeTra_frm_166.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Buscar Datos"
            Top             =   60
            Width           =   585
         End
         Begin VB.TextBox txt_NumDoc 
            Height          =   315
            Left            =   1980
            MaxLength       =   11
            TabIndex        =   1
            Text            =   "Text1"
            Top             =   390
            Width           =   3315
         End
         Begin VB.ComboBox cmb_TipDoc 
            Height          =   315
            Left            =   1980
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   3315
         End
         Begin VB.Label lbl_General 
            Caption         =   "N�mero Docum. Ident.:"
            Height          =   285
            Index           =   48
            Left            =   60
            TabIndex        =   46
            Top             =   390
            Width           =   1845
         End
         Begin VB.Label lbl_General 
            Caption         =   "Tipo Docum. Ident.:"
            Height          =   285
            Index           =   36
            Left            =   60
            TabIndex        =   45
            Top             =   60
            Width           =   1635
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   37
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
            TabIndex        =   42
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
            TabIndex        =   43
            Top             =   330
            Width           =   6165
            _Version        =   65536
            _ExtentX        =   10874
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "Actividades Econ�micas - Dependiente o Pensionista"
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
            Picture         =   "OpeTra_frm_166.frx":092A
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   435
         Left            =   30
         TabIndex        =   38
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
            Left            =   1980
            TabIndex        =   39
            Top             =   60
            Width           =   9555
            _Version        =   65536
            _ExtentX        =   16854
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
            Left            =   60
            TabIndex        =   40
            Top             =   60
            Width           =   1725
         End
      End
      Begin Threed.SSPanel SSPanel9 
         Height          =   675
         Left            =   30
         TabIndex        =   41
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
            Picture         =   "OpeTra_frm_166.frx":0C34
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_166.frx":1076
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   30
            Width           =   585
         End
      End
   End
End
Attribute VB_Name = "frm_MntCli_57"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_str_DptDir     As String
Dim l_str_PrvDir     As String
Dim l_str_DstDir     As String
Dim l_int_FlgCmb     As Integer

Dim l_str_RazSoc     As String
Dim l_str_NomCom     As String
Dim l_int_TipVia     As Integer
Dim l_str_NomVia     As String
Dim l_str_NumVia     As String
Dim l_str_IntDpt     As String
Dim l_int_TipZon     As Integer
Dim l_str_NomZon     As String
Dim l_str_UbiGeo     As String
Dim l_str_Refere     As String
Dim l_str_Telef1     As String
Dim l_str_Telef2     As String
Dim l_str_NumFax     As String
Dim l_str_TeleRH     As String
Dim l_str_AnexRH     As String
Dim l_int_CodCiu     As Integer

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

Private Sub cmb_SitTra_Click()
   Call gs_SetFocus(cmb_Tipo)
End Sub

Private Sub cmb_SitTra_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_SitTra_Click
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

Private Sub cmb_Tipo_Click()
   Call gs_SetFocus(cmb_TipOfi)
End Sub

Private Sub cmb_Tipo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Tipo_Click
   End If
End Sub

Private Sub cmb_TipOfi_Click()
   If cmb_TipOfi.ListIndex > -1 Then
      If cmb_TipOfi.ItemData(cmb_TipOfi.ListIndex) = 2 Then
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
         txt_Telef1.Enabled = True
         txt_Telef2.Enabled = True
         txt_NumFax.Enabled = True
         
         Call gs_SetFocus(cmb_TipVia)
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
         txt_Telef1.Text = ""
         txt_Telef2.Text = ""
         txt_NumFax.Text = ""
         
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
         txt_Telef1.Enabled = False
         txt_Telef2.Enabled = False
         txt_NumFax.Enabled = False
         
         Call gs_SetFocus(txt_TelDir)
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
      txt_Telef1.Text = ""
      txt_Telef2.Text = ""
      txt_NumFax.Text = ""
      
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
      txt_Telef1.Enabled = False
      txt_Telef2.Enabled = False
      txt_NumFax.Enabled = False
      
      Call gs_SetFocus(txt_TelDir)
   End If
End Sub

Private Sub cmb_TipOfi_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipOfi_Click
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

Private Sub cmd_Buscar_Click()
   If cmb_TipDoc.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Documento.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipDoc)
      Exit Sub
   End If
   If cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex) = 7 Then
      If Len(Trim(txt_NumDoc.Text)) <> 11 Then
         MsgBox "El N�mero de Ruc debe tener 11 digitos.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NumDoc)
         Exit Sub
      End If
      If Not gf_Valida_RUC(Trim(txt_NumDoc.Text), Mid(Trim(txt_NumDoc.Text), Len(Trim(txt_NumDoc.Text)), 1)) Then
         MsgBox "El N�mero de RUC no es valido.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NumDoc)
         Exit Sub
      End If
   Else
      If Len(Trim(txt_NumDoc.Text)) = 0 Then
         MsgBox "Debe ingresar el N�mero de Documento.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NumDoc)
         Exit Sub
      End If
   End If
   
   'Buscar Empresa en Maestro de Empresas
   modmip_g_int_TDoEmp = CStr(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex))
   modmip_g_str_TDoEmp = cmb_TipDoc.Text
   modmip_g_str_NDoEmp = Trim(txt_NumDoc.Text)
   
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
      MsgBox "Debe ingresar el N�mero de Documento.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(txt_NumDoc)
      Exit Sub
   End If
   
   If grd_Listad.Rows = 0 Then
      MsgBox "Debe ingresar la informaci�n de la Empresa.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipDoc)
      Exit Sub
   End If
   
   If cmb_SitTra.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Situaci�n del Trabajador.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_SitTra)
      Exit Sub
   End If
   
   If cmb_Tipo.ListIndex = -1 Then
      MsgBox "Debe seleccionar un tipo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_Tipo)
      Exit Sub
   End If
   
   If cmb_TipOfi.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Oficina.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipOfi)
      Exit Sub
   End If
   
   If cmb_TipOfi.ItemData(cmb_TipOfi.ListIndex) = 2 Then
      If cmb_TipVia.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Tipo de V�a.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_TipVia)
         Exit Sub
      End If
      
      If Len(Trim(txt_NomVia.Text)) = 0 Then
         MsgBox "Debe ingresar el Nombre de V�a.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(txt_NomVia)
         Exit Sub
      End If
      
      If Len(Trim(txt_NumVia.Text)) = 0 Then
         MsgBox "Debe ingresar el N�mero.", vbExclamation, modgen_g_str_NomPlt
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
         MsgBox "Debe seleccionar el Departamento de la Direcci�n.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_DptDir)
         Exit Sub
      End If
      
      If cmb_PrvDir.ListIndex = -1 Then
         MsgBox "Debe seleccionar la Provincia de la Direcci�n.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_PrvDir)
         Exit Sub
      End If
      
      If cmb_DstDir.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Distrito de la Direcci�n.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_DstDir)
         Exit Sub
      End If
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
   
   If MsgBox("�Est� seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   
   If modmip_g_int_FlgGrb_1 = 2 Then
      'Borrar Actividad Econ�mica
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
   
   'Insertando Actividad Econ�mica
   g_str_Parame = "USP_CLI_ACTECO_AGREGA ("
   
   If modmip_g_int_TipCli = 1 Then
      g_str_Parame = g_str_Parame & CStr(moddat_g_int_TipDoc) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumDoc & "', "
   Else
      g_str_Parame = g_str_Parame & CStr(moddat_g_int_CygTDo) & ", "
      g_str_Parame = g_str_Parame & "'" & moddat_g_str_CygNDo & "', "
   End If
   
   g_str_Parame = g_str_Parame & CStr(modmip_g_int_OrdAct) & ", "
   g_str_Parame = g_str_Parame & "11, "                                                      'C�digo Actividad Econ�mica (Dependiente)
   
   'Dependiente
   g_str_Parame = g_str_Parame & CStr(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)) & ", "      'Tipo DOI Empleador
   g_str_Parame = g_str_Parame & "'" & txt_NumDoc.Text & "', "                               'N�mero DOI Empleador
   g_str_Parame = g_str_Parame & "'" & l_str_RazSoc & "', "                                  'Raz�n Social
   g_str_Parame = g_str_Parame & "'" & l_str_NomCom & "', "                                  'Nombre Comercial
   g_str_Parame = g_str_Parame & CStr(cmb_TipOfi.ItemData(cmb_TipOfi.ListIndex)) & ", "      'Tipo de Oficina
   g_str_Parame = g_str_Parame & CStr(cmb_SitTra.ItemData(cmb_SitTra.ListIndex)) & ", "      'Situaci�n Trabajador
   
   If cmb_TipOfi.ItemData(cmb_TipOfi.ListIndex) = 1 Then
      g_str_Parame = g_str_Parame & CStr(l_int_TipVia) & ", "                                'Tipo de Via
      g_str_Parame = g_str_Parame & "'" & l_str_NomVia & "', "                               'Nombre de V�a
      g_str_Parame = g_str_Parame & "'" & l_str_NumVia & "', "                               'N�mero de V�a
      g_str_Parame = g_str_Parame & "'" & l_str_IntDpt & "', "                               'Interior / Dpto.
      g_str_Parame = g_str_Parame & CStr(l_int_TipZon) & ", "                                'Tipo de Zona
      g_str_Parame = g_str_Parame & "'" & l_str_NomZon & "', "                               'Nombre de Zona
      g_str_Parame = g_str_Parame & "'" & l_str_UbiGeo & "', "                               'Ubicaci�n Geogr�fica
      g_str_Parame = g_str_Parame & "'" & l_str_Refere & "', "                               'Referencia
      g_str_Parame = g_str_Parame & "'" & l_str_Telef1 & "', "                               'Tel�fono 1
      g_str_Parame = g_str_Parame & "'" & l_str_Telef2 & "', "                               'Tel�fono 2
      g_str_Parame = g_str_Parame & "'" & l_str_NumFax & "', "                               'N�mero de Fax
   Else
      g_str_Parame = g_str_Parame & CStr(cmb_TipVia.ItemData(cmb_TipVia.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & "'" & txt_NomVia.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_NumVia.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_IntDpt.Text & "', "
      g_str_Parame = g_str_Parame & CStr(cmb_TipZon.ItemData(cmb_TipZon.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & "'" & txt_NomZon.Text & "', "
      g_str_Parame = g_str_Parame & "'" & Format(cmb_DptDir.ItemData(cmb_DptDir.ListIndex), "00") & Format(cmb_PrvDir.ItemData(cmb_PrvDir.ListIndex), "00") & Format(cmb_DstDir.ItemData(cmb_DstDir.ListIndex), "00") & "', "
      g_str_Parame = g_str_Parame & "'" & txt_Refere.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_Telef1.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_Telef2.Text & "', "
      g_str_Parame = g_str_Parame & "'" & txt_NumFax.Text & "', "
   End If
   
   g_str_Parame = g_str_Parame & CStr(l_int_CodCiu) & ", "                                'C�digo CIIU
   g_str_Parame = g_str_Parame & "'" & l_str_TeleRH & "', "                               'Telefono RR.HH
   g_str_Parame = g_str_Parame & "'" & l_str_AnexRH & "', "                               'Anexo RR.HH
   g_str_Parame = g_str_Parame & CDbl(ipp_IngNet.Text) & ", "                             'Ingreso Neto
   g_str_Parame = g_str_Parame & "1, "                                                    'Frecuencia de Haberes
   g_str_Parame = g_str_Parame & Format(CDate(ipp_FecIng.Text), "yyyymmdd") & ", "        'Fecha de Ingreso
   g_str_Parame = g_str_Parame & "'999999', "                                             'C�digo de Cargo
   g_str_Parame = g_str_Parame & "'" & txt_NomCar.Text & "', "                            'Nombre de Cargo
   g_str_Parame = g_str_Parame & "'', "                                                   'Nombre de Area
   g_str_Parame = g_str_Parame & "'" & txt_NumAnx.Text & "', "                            'N�mero de Anexo
   g_str_Parame = g_str_Parame & "'" & txt_TelDir.Text & "', "                            'Tel�fono Directo
   g_str_Parame = g_str_Parame & "'', "                                                   'Celular del Trabajo
   g_str_Parame = g_str_Parame & "'', "                                                   'E-mail del Trabajo
   g_str_Parame = g_str_Parame & "2, "                                                    'Flag de Trabajo Anterior
   g_str_Parame = g_str_Parame & "0, "                                                    'Tipo DOI Empleador Anterior
   g_str_Parame = g_str_Parame & "'', "                                                   'N�mero DOI Empleador Anterior
   g_str_Parame = g_str_Parame & "'', "                                                   'Raz�n Social Empleador Anterior
   g_str_Parame = g_str_Parame & "'', "                                                   'Nombre Comercial Empleador Anterior
   g_str_Parame = g_str_Parame & "'', "                                                   'Tel�fono 1 Empleador Anterior
   g_str_Parame = g_str_Parame & "'', "                                                   'Tel�fono 2 Empleador Anterior
   g_str_Parame = g_str_Parame & "0, "                                                    'Fecha Ingreso Empleador Anterior
   g_str_Parame = g_str_Parame & "0, "                                                    'Fecha Cese Empleador Anterior
   
   
   'Independiente
   g_str_Parame = g_str_Parame & "0, "                                                    'Tipo DOI
   g_str_Parame = g_str_Parame & "'', "                                                   'N�mero DOI
   g_str_Parame = g_str_Parame & "0, "                                                    'Tipo V�a
   g_str_Parame = g_str_Parame & "'', "                                                   'Nombre V�a
   g_str_Parame = g_str_Parame & "'', "                                                   'N�mero V�a
   g_str_Parame = g_str_Parame & "'', "                                                   'Interior
   g_str_Parame = g_str_Parame & "0, "                                                    'Tipo Zona
   g_str_Parame = g_str_Parame & "'', "                                                   'Nombre zona
   g_str_Parame = g_str_Parame & "'', "                                                   'Ubicaci�n Geogr�fica
   g_str_Parame = g_str_Parame & "'', "                                                   'Referencia
   g_str_Parame = g_str_Parame & "'', "                                                   'Tel�fono 1
   g_str_Parame = g_str_Parame & "'', "                                                   'Tel�fono 2
   g_str_Parame = g_str_Parame & "'', "                                                   'Fax
   g_str_Parame = g_str_Parame & "0, "                                                    'CIIU
   g_str_Parame = g_str_Parame & "0, "                                                    'Ingreso Neto
   g_str_Parame = g_str_Parame & "0, "                                                    'Inicio de Actividad
   g_str_Parame = g_str_Parame & "0, "                                                    'Contrato Locaci�n
   g_str_Parame = g_str_Parame & "0, "                                                    'Tipo DOI Empleador
   g_str_Parame = g_str_Parame & "'', "                                                   'N�mero DOI Empleador
   g_str_Parame = g_str_Parame & "'', "                                                   'Raz�n Social Empleador
   g_str_Parame = g_str_Parame & "'', "                                                   'Nombre Comercial Empleador
   g_str_Parame = g_str_Parame & "'', "                                                   'Tel�fono 1 Empleador
   g_str_Parame = g_str_Parame & "'', "                                                   'Tel�fono 2 Empleador
   g_str_Parame = g_str_Parame & "0, "                                                    'Fecha Ingreso Empleador
   g_str_Parame = g_str_Parame & "'', "                                                   'C�digo Cargo
   g_str_Parame = g_str_Parame & "'', "                                                   'Nombre Cargo
   
   'Comerciante
   g_str_Parame = g_str_Parame & "0, "                                                    'Tipo DOI
   g_str_Parame = g_str_Parame & "'', "                                                   'N�mero DOI
   g_str_Parame = g_str_Parame & "'', "                                                   'Raz�n Social Empleador
   g_str_Parame = g_str_Parame & "'', "                                                   'Nombre Comercial Empleador
   g_str_Parame = g_str_Parame & "0, "                                                    'Tipo V�a
   g_str_Parame = g_str_Parame & "'', "                                                   'Nombre V�a
   g_str_Parame = g_str_Parame & "'', "                                                   'N�mero V�a
   g_str_Parame = g_str_Parame & "'', "                                                   'Interior
   g_str_Parame = g_str_Parame & "0, "                                                    'Tipo Zona
   g_str_Parame = g_str_Parame & "'', "                                                   'Nombre zona
   g_str_Parame = g_str_Parame & "'', "                                                   'Ubicaci�n Geogr�fica
   g_str_Parame = g_str_Parame & "'', "                                                   'Referencia
   g_str_Parame = g_str_Parame & "'', "                                                   'Tel�fono 1
   g_str_Parame = g_str_Parame & "'', "                                                   'Tel�fono 2
   g_str_Parame = g_str_Parame & "'', "                                                   'Fax
   g_str_Parame = g_str_Parame & "0, "                                                    'CIIU
   g_str_Parame = g_str_Parame & "'', "                                                   'Giro Comercial
   g_str_Parame = g_str_Parame & "0, "                                                    'Ingreso Neto
   g_str_Parame = g_str_Parame & "0, "                                                    'Ventas Mensuales
   g_str_Parame = g_str_Parame & "0, "                                                    'Inicio de Operaciones
   g_str_Parame = g_str_Parame & "'', "                                                   'C�digo Cargo
   g_str_Parame = g_str_Parame & "'', "                                                   'Nombre Cargo
   g_str_Parame = g_str_Parame & "0, "                                                    'R�gimen Tributario
   g_str_Parame = g_str_Parame & "0, "                                                    'Porcentaje Participaci�n
   g_str_Parame = g_str_Parame & "0, "                                                    'Tipo Local
   g_str_Parame = g_str_Parame & "0, "                                                    'Alquiler Mensual
   g_str_Parame = g_str_Parame & "'', "                                                   'Nombre Arrendador
   g_str_Parame = g_str_Parame & "'', "                                                   'Tel�fono Arrendador
   
   'Accionista
   g_str_Parame = g_str_Parame & "0, "                                                    'Tipo DOI
   g_str_Parame = g_str_Parame & "'', "                                                   'N�mero DOI
   g_str_Parame = g_str_Parame & "'', "                                                   'Raz�n Social Empleador
   g_str_Parame = g_str_Parame & "'', "                                                   'Nombre Comercial Empleador
   g_str_Parame = g_str_Parame & "0, "                                                    'Tipo V�a
   g_str_Parame = g_str_Parame & "'', "                                                   'Nombre V�a
   g_str_Parame = g_str_Parame & "'', "                                                   'N�mero V�a
   g_str_Parame = g_str_Parame & "'', "                                                   'Interior
   g_str_Parame = g_str_Parame & "0, "                                                    'Tipo Zona
   g_str_Parame = g_str_Parame & "'', "                                                   'Nombre zona
   g_str_Parame = g_str_Parame & "'', "                                                   'Ubicaci�n Geogr�fica
   g_str_Parame = g_str_Parame & "'', "                                                   'Referencia
   g_str_Parame = g_str_Parame & "'', "                                                   'Tel�fono 1
   g_str_Parame = g_str_Parame & "'', "                                                   'Tel�fono 2
   g_str_Parame = g_str_Parame & "'', "                                                   'Fax
   g_str_Parame = g_str_Parame & "0, "                                                    'CIIU
   g_str_Parame = g_str_Parame & "0, "                                                    'Ingreso Neto
   g_str_Parame = g_str_Parame & "0, "                                                    'Porcentaje Participaci�n
   g_str_Parame = g_str_Parame & "0, "                                                    'Fecha Antig�edad
   
   'Rentista
   g_str_Parame = g_str_Parame & "0, "                                                    'Ingreso Neto
   g_str_Parame = g_str_Parame & "'', "                                                   'Direcci�n 1
   g_str_Parame = g_str_Parame & "'', "                                                   'Nombre 1
   g_str_Parame = g_str_Parame & "0, "                                                    'Inicio Alquiler 1
   g_str_Parame = g_str_Parame & "'', "                                                   'Tel�fono 1 - 1
   g_str_Parame = g_str_Parame & "'', "                                                   'Tel�fono 2 - 1
   g_str_Parame = g_str_Parame & "0, "                                                    'Monto Alquiler 1
   g_str_Parame = g_str_Parame & "0, "                                                    'Segunda Propiedad
   g_str_Parame = g_str_Parame & "'', "                                                   'Direcci�n 2
   g_str_Parame = g_str_Parame & "'', "                                                   'Nombre 2
   g_str_Parame = g_str_Parame & "0, "                                                    'Inicio Alquiler 2
   g_str_Parame = g_str_Parame & "'', "                                                   'Tel�fono 1 - 2
   g_str_Parame = g_str_Parame & "'', "                                                   'Tel�fono 2 - 2
   g_str_Parame = g_str_Parame & "0, "                                                    'Monto Alquiler 2
   
   'Otros
   g_str_Parame = g_str_Parame & "0, "                                                    'Ingreso Neto
   g_str_Parame = g_str_Parame & "'', "                                                   'Actividad
   g_str_Parame = g_str_Parame & "0, "                                                    'CIIU
   g_str_Parame = g_str_Parame & "'', "                                                   'Observaciones
   
   'Dependiente
   g_str_Parame = g_str_Parame & CStr(cmb_MonIng.ItemData(cmb_MonIng.ListIndex)) & ", "   'Moneda Ingresos
   g_str_Parame = g_str_Parame & "'', "                                                   'Direcci�n
   g_str_Parame = g_str_Parame & "'', "                                                   'Ciudad
   g_str_Parame = g_str_Parame & "'', "                                                   'C�digo Postal
   
   'Independiente
   g_str_Parame = g_str_Parame & "0, "                                                    'Moneda Ingresos
   g_str_Parame = g_str_Parame & "'', "                                                   'Direcci�n
   g_str_Parame = g_str_Parame & "'', "                                                   'Ciudad
   g_str_Parame = g_str_Parame & "'', "                                                   'C�digo Postal
   
   'Comerciante
   g_str_Parame = g_str_Parame & "0, "                                                    'Moneda Ingresos
   g_str_Parame = g_str_Parame & "'', "                                                   'Direcci�n
   g_str_Parame = g_str_Parame & "'', "                                                   'Ciudad
   g_str_Parame = g_str_Parame & "'', "                                                   'C�digo Postal
   
   'Accionista
   g_str_Parame = g_str_Parame & "0, "                                                    'Moneda Ingresos
   
   'Rentista
   g_str_Parame = g_str_Parame & "0, "                                                    'Moneda Ingresos
   
   'Otros
   g_str_Parame = g_str_Parame & "0, "                                                    'Moneda Ingresos
   
   'Dependiente
   g_str_Parame = g_str_Parame & CStr(cmb_Tipo.ItemData(cmb_Tipo.ListIndex)) & ", "   'TIPO IDENTIFICACION(ACTECO_DEP_TIPIDE)
   
   'Independiente
   g_str_Parame = g_str_Parame & "0, "    'TIPO IDENTIFICACION(ACTECO_IND_TIPIDE)
   
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
      
      g_str_Parame = g_str_Parame & "11, "
      g_str_Parame = g_str_Parame & CStr(l_int_CodCiu) & ", "
      g_str_Parame = g_str_Parame & CStr(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & "'" & txt_NumDoc.Text & "', "
      
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
   cmb_TipDoc.ListIndex = -1
   txt_NumDoc.Text = ""
   
   Call gs_LimpiaGrid(grd_Listad)
   
   Call fs_Activa_Emp(True)
   
   Call gs_SetFocus(cmb_TipDoc)
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
   Call fs_Activa_Emp(True)
   
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
      
      Call gs_BuscarCombo_Item(cmb_TipDoc, g_rst_Princi!ActEco_Dep_TipDoc)
      txt_NumDoc.Text = Trim(g_rst_Princi!ActEco_Dep_NumDoc)
      
      modmip_g_int_TDoEmp = CStr(cmb_TipDoc.ItemData(cmb_TipDoc.ListIndex))
      modmip_g_str_TDoEmp = cmb_TipDoc.Text
      modmip_g_str_NDoEmp = Trim(txt_NumDoc.Text)
         
      Call fs_BusEmp
         
      Call gs_BuscarCombo_Item(cmb_SitTra, g_rst_Princi!ActEco_Dep_SitTra)
      Call gs_BuscarCombo_Item(cmb_TipOfi, g_rst_Princi!ActEco_Dep_TipOfi)
      If Trim(g_rst_Princi!ACTECO_DEP_TIPIDE & "") <> "" Then
         Call gs_BuscarCombo_Item(cmb_Tipo, g_rst_Princi!ACTECO_DEP_TIPIDE)
      End If
      If cmb_TipOfi.ItemData(cmb_TipOfi.ListIndex) = 2 Then
         Call gs_BuscarCombo_Item(cmb_TipVia, g_rst_Princi!ActEco_Dep_TipVia)
         
         txt_NomVia.Text = Trim(g_rst_Princi!ActEco_Dep_NomVia & "")
         txt_NumVia.Text = Trim(g_rst_Princi!ActEco_Dep_NumVia & "")
         txt_IntDpt.Text = Trim(g_rst_Princi!ActEco_Dep_IntDpt & "")
         
         Call gs_BuscarCombo_Item(cmb_TipZon, g_rst_Princi!ActEco_Dep_TipZon)
         txt_NomZon.Text = Trim(g_rst_Princi!ActEco_Dep_NomZon & "")
      
         Call gs_BuscarCombo_Item(cmb_DptDir, CInt(Left(g_rst_Princi!ActEco_Dep_UbiGeo, 2)))
         Call moddat_gs_Carga_Provin(cmb_PrvDir, Left(g_rst_Princi!ActEco_Dep_UbiGeo, 2))
         Call gs_BuscarCombo_Item(cmb_PrvDir, CInt(Mid(g_rst_Princi!ActEco_Dep_UbiGeo, 3, 2)))
         Call moddat_gs_Carga_Distri(cmb_DstDir, Left(g_rst_Princi!ActEco_Dep_UbiGeo, 2), Mid(g_rst_Princi!ActEco_Dep_UbiGeo, 3, 2))
         Call gs_BuscarCombo_Item(cmb_DstDir, CInt(Right(g_rst_Princi!ActEco_Dep_UbiGeo, 2)))
         
         txt_Refere.Text = Trim(g_rst_Princi!ActEco_Dep_Refere & "")
         txt_Telef1.Text = Trim(g_rst_Princi!ActEco_Dep_Telef1 & "")
         txt_Telef2.Text = Trim(g_rst_Princi!ActEco_Dep_Telef2 & "")
         txt_NumFax.Text = Trim(g_rst_Princi!ActEco_Dep_NumFax & "")
      End If
      
      txt_TelDir.Text = Trim(g_rst_Princi!ActEco_Dep_TelDir & "")
      txt_NumAnx.Text = Trim(g_rst_Princi!ActEco_Dep_NumAnx & "")
         
      Call gs_BuscarCombo_Item(cmb_MonIng, g_rst_Princi!ActEco_Dep_MonIng)
      
      ipp_IngNet.Value = g_rst_Princi!ActEco_Dep_IngNet
      ipp_FecIng.Text = gf_FormatoFecha(CStr(g_rst_Princi!ActEco_Dep_FecIng))
      
      txt_NomCar.Text = Trim(g_rst_Princi!ActEco_Dep_NomCar & "")
      
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
   Call moddat_gs_Carga_LisIte_Combo(cmb_SitTra, 1, "235")
   Call moddat_gs_Carga_LisIte_Combo(cmb_Tipo, 1, "380")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipOfi, 1, "234")
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipVia, 1, "201")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipZon, 1, "202")
   
   Call moddat_gs_Carga_Depart(cmb_DptDir)
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_MonIng, 1, "113")
End Sub

Private Sub fs_Limpia()
   cmb_TipDoc.ListIndex = -1
   txt_NumDoc.Text = ""
   
   Call gs_LimpiaGrid(grd_Listad)
   
   cmb_SitTra.ListIndex = -1
   cmb_Tipo.ListIndex = -1
   cmb_TipOfi.ListIndex = -1
   
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
   txt_Telef1.Enabled = False
   txt_Telef2.Enabled = False
   txt_NumFax.Enabled = False
   
   txt_TelDir.Text = ""
   txt_NumAnx.Text = ""
   cmb_MonIng.ListIndex = -1
   ipp_IngNet.Value = 0
   ipp_FecIng.Text = Format(date, "dd/mm/yyyy")
   txt_NomCar.Text = ""

   l_str_RazSoc = ""
   l_str_NomCom = ""
   l_int_TipVia = 0
   l_str_NomVia = ""
   l_str_NumVia = ""
   l_str_IntDpt = ""
   l_int_TipZon = 0
   l_str_NomZon = ""
   l_str_UbiGeo = ""
   l_str_Refere = ""
   l_str_Telef1 = ""
   l_str_Telef2 = ""
   l_str_NumFax = ""
   l_str_TeleRH = ""
   l_str_AnexRH = ""
   l_int_CodCiu = 0
   modmip_g_int_FlgGrb = 0
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
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/�")
   End If
End Sub

Private Sub txt_NomCar_GotFocus()
   Call gs_SelecTodo(txt_NomCar)
End Sub

Private Sub txt_NomCar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Grabar)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/�")
   End If
End Sub

Private Sub txt_NomVia_GotFocus()
   Call gs_SelecTodo(txt_NomVia)
End Sub

Private Sub txt_NomVia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NumVia)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/�")
   End If
End Sub

Private Sub txt_NomZon_GotFocus()
   Call gs_SelecTodo(txt_NomZon)
End Sub

Private Sub txt_NomZon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_DptDir)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/�")
   End If
End Sub

Private Sub txt_NumVia_GotFocus()
   Call gs_SelecTodo(txt_NumVia)
End Sub

Private Sub txt_NumVia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_IntDpt)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-_. ,;:()/�")
   End If
End Sub

Private Sub txt_Refere_GotFocus()
   Call gs_SelecTodo(txt_Refere)
End Sub

Private Sub txt_Refere_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_Telef1)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-( )%$.,;:@_?��")
   End If
End Sub

Private Sub txt_NumDoc_GotFocus()
   Call gs_SelecTodo(txt_NumDoc)
End Sub

Private Sub txt_NumDoc_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Buscar)
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
      Call gs_SetFocus(txt_TelDir)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_TelDir_GotFocus()
   Call gs_SelecTodo(txt_TelDir)
End Sub

Private Sub txt_TelDir_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txt_NumAnx)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub txt_NumAnx_GotFocus()
   Call gs_SelecTodo(txt_NumAnx)
End Sub

Private Sub txt_NumAnx_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_MonIng)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
   End If
End Sub

Private Sub fs_BusEmp()
   Dim r_rst_Princi     As ADODB.Recordset
   
   l_str_RazSoc = ""
   l_str_NomCom = ""
   l_int_TipVia = 0
   l_str_NomVia = ""
   l_str_NumVia = ""
   l_str_IntDpt = ""
   l_int_TipZon = 0
   l_str_NomZon = ""
   l_str_UbiGeo = ""
   l_str_Refere = ""
   l_str_Telef1 = ""
   l_str_Telef2 = ""
   l_str_NumFax = ""
   l_str_TeleRH = ""
   l_str_AnexRH = ""
   l_int_CodCiu = 0
   
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
   grd_Listad.Text = "Raz�n Social"
   
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
   grd_Listad.Text = "Direcci�n"
   
   grd_Listad.Col = 1
   grd_Listad.Text = moddat_gf_Consulta_ParDes("201", CStr(r_rst_Princi!DatGen_TipVia)) & _
                               " " & Trim(r_rst_Princi!DatGen_NomVia) & " " & Trim(r_rst_Princi!DatGen_numVia) & _
                               IIf(Len(Trim(r_rst_Princi!DATGEN_INTDPT)) > 0, " (" & Trim(r_rst_Princi!DATGEN_INTDPT) & ")", "") & _
                               IIf(Len(Trim(r_rst_Princi!DatGen_NomZon)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(r_rst_Princi!DatGen_TipZon)) & " " & Trim(r_rst_Princi!DatGen_NomZon), "")
         
   If Len(Trim(r_rst_Princi!DATGEN_REFERE & "")) > 0 Then
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      grd_Listad.Col = 0
      grd_Listad.Text = "Referencia"
      
      grd_Listad.Col = 1
      grd_Listad.Text = Trim(r_rst_Princi!DATGEN_REFERE & "")
   End If
         
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Dpto. / Provin. / Dist."

   grd_Listad.Col = 1
   grd_Listad.Text = moddat_gf_Consulta_ParDes("101", Left(r_rst_Princi!DatGen_Ubigeo & "", 2) & "0000") & _
                               " - " & moddat_gf_Consulta_ParDes("101", Left(r_rst_Princi!DatGen_Ubigeo & "", 4) & "00") & _
                               " - " & moddat_gf_Consulta_ParDes("101", Trim(r_rst_Princi!DatGen_Ubigeo & ""))
         
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.Col = 0
   grd_Listad.Text = "Tel�fonos"
   
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
      grd_Listad.Text = "Tel�fono RR.HH"
      
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
      grd_Listad.Text = "P�gina Web"
      
      grd_Listad.Col = 1
      grd_Listad.Text = Trim(r_rst_Princi!DATGEN_PAGWEB & "")
   End If
   
   
   l_str_RazSoc = Trim(r_rst_Princi!DATGEN_RAZSOC & "")
   l_str_NomCom = Trim(r_rst_Princi!DATGEN_NOMCOM & "")
   l_int_TipVia = r_rst_Princi!DatGen_TipVia
   l_str_NomVia = Trim(r_rst_Princi!DatGen_NomVia & "")
   l_str_NumVia = Trim(r_rst_Princi!DatGen_numVia & "")
   l_str_IntDpt = Trim(r_rst_Princi!DATGEN_INTDPT & "")
   l_int_TipZon = r_rst_Princi!DatGen_TipZon
   l_str_NomZon = Trim(r_rst_Princi!DatGen_NomZon & "")
   l_str_UbiGeo = Trim(r_rst_Princi!DatGen_Ubigeo & "")
   l_str_Refere = Trim(r_rst_Princi!DATGEN_REFERE & "")
   l_str_Telef1 = Trim(r_rst_Princi!DATGEN_TELEF1 & "")
   l_str_Telef2 = Trim(r_rst_Princi!DATGEN_TELEF2 & "")
   l_str_NumFax = Trim(r_rst_Princi!DATGEN_NUMFAX & "")
   l_str_TeleRH = Trim(r_rst_Princi!DATGEN_TELERH & "")
   l_str_AnexRH = Trim(r_rst_Princi!DATGEN_ANEXRH & "")
   l_int_CodCiu = r_rst_Princi!DATGEN_CODCIU
   
   r_rst_Princi.Close
   Set r_rst_Princi = Nothing
   
   grd_Listad.Redraw = True
   
   Call gs_UbiIniGrid(grd_Listad)
End Sub

Private Sub fs_Activa_Emp(ByVal p_Activa As Integer)
   cmb_TipDoc.Enabled = p_Activa
   txt_NumDoc.Enabled = p_Activa
   
   cmd_Buscar.Enabled = p_Activa
      
   grd_Listad.Enabled = Not p_Activa
   cmd_Editar.Enabled = Not p_Activa
End Sub

