VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_MntCli_55 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6720
   ClientLeft      =   3840
   ClientTop       =   2775
   ClientWidth     =   11610
   Icon            =   "OpeTra_frm_164.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   11610
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   6765
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   11655
      _Version        =   65536
      _ExtentX        =   20558
      _ExtentY        =   11933
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
      Begin Threed.SSPanel SSPanel11 
         Height          =   765
         Left            =   30
         TabIndex        =   48
         Top             =   2490
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
         Begin VB.ComboBox cmb_DocAlt 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   1065
         End
         Begin VB.ComboBox cmb_TDoAlt 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   390
            Width           =   3315
         End
         Begin VB.TextBox txt_NDoAlt 
            Height          =   315
            Left            =   8220
            MaxLength       =   12
            TabIndex        =   2
            Text            =   "Text1"
            Top             =   390
            Width           =   3315
         End
         Begin VB.Label Label35 
            Caption         =   "Personal FF.AA / FF.PP:"
            Height          =   315
            Left            =   90
            TabIndex        =   51
            Top             =   60
            Width           =   1845
         End
         Begin VB.Label Label34 
            Caption         =   "Tipo Docum. Identidad:"
            Height          =   315
            Left            =   90
            TabIndex        =   50
            Top             =   390
            Width           =   1845
         End
         Begin VB.Label Label33 
            Caption         =   "Nro. Doc. Id.:"
            Height          =   285
            Left            =   6240
            TabIndex        =   49
            Top             =   390
            Width           =   1065
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   25
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
         Begin Threed.SSPanel SSPanel16 
            Height          =   255
            Left            =   630
            TabIndex        =   46
            Top             =   30
            Width           =   2955
            _Version        =   65536
            _ExtentX        =   5212
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
            Left            =   630
            TabIndex        =   47
            Top             =   300
            Width           =   2685
            _Version        =   65536
            _ExtentX        =   4736
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "Datos del Cónyuge"
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
         Begin Threed.SSPanel pnl_RelLab 
            Height          =   285
            Left            =   5730
            TabIndex        =   54
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
         Begin Threed.SSPanel pnl_RelAcc 
            Height          =   285
            Left            =   5730
            TabIndex        =   55
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
         Begin VB.Image Image1 
            Height          =   480
            Left            =   60
            Picture         =   "OpeTra_frm_164.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   3405
         Left            =   30
         TabIndex        =   26
         Top             =   3300
         Width           =   11565
         _Version        =   65536
         _ExtentX        =   20399
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
         Begin VB.ComboBox cmb_TipAfp 
            Height          =   315
            Left            =   8190
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   2040
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
            Left            =   9930
            TabIndex        =   17
            Top             =   2730
            Width           =   1485
         End
         Begin VB.TextBox txt_DirEle 
            Height          =   315
            Left            =   8190
            MaxLength       =   120
            TabIndex        =   16
            Text            =   "Text1"
            Top             =   2700
            Width           =   1665
         End
         Begin VB.TextBox txt_Celula 
            Height          =   315
            Left            =   2010
            MaxLength       =   9
            TabIndex        =   15
            Text            =   "Text1"
            Top             =   2700
            Width           =   3315
         End
         Begin VB.ComboBox cmb_Profes 
            Height          =   315
            Left            =   2010
            TabIndex        =   14
            Text            =   "cmb_Profes"
            Top             =   2370
            Width           =   9525
         End
         Begin VB.ComboBox cmb_NivEst 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   2040
            Width           =   3315
         End
         Begin VB.ComboBox cmb_DstNac 
            Height          =   315
            Left            =   8190
            TabIndex        =   11
            Text            =   "cmb_DstNac"
            Top             =   1710
            Width           =   3315
         End
         Begin VB.ComboBox cmb_PrvNac 
            Height          =   315
            Left            =   2010
            TabIndex        =   10
            Text            =   "cmb_PrvNac"
            Top             =   1710
            Width           =   3315
         End
         Begin VB.ComboBox cmb_DptNac 
            Height          =   315
            Left            =   8190
            TabIndex        =   9
            Text            =   "cmb_DptNac"
            Top             =   1380
            Width           =   3315
         End
         Begin VB.ComboBox cmb_Paises 
            Height          =   315
            Left            =   2010
            TabIndex        =   8
            Text            =   "cmb_Paises"
            Top             =   1380
            Width           =   3315
         End
         Begin VB.TextBox txt_Nombre 
            Height          =   315
            Left            =   2010
            MaxLength       =   30
            TabIndex        =   6
            Text            =   "Text1"
            Top             =   720
            Width           =   3315
         End
         Begin VB.TextBox txt_ApePat 
            Height          =   315
            Left            =   2010
            MaxLength       =   30
            TabIndex        =   3
            Text            =   "Text1"
            Top             =   60
            Width           =   3315
         End
         Begin VB.TextBox txt_ApeMat 
            Height          =   315
            Left            =   2010
            MaxLength       =   30
            TabIndex        =   4
            Text            =   "Text1"
            Top             =   390
            Width           =   3315
         End
         Begin VB.TextBox txt_ApeCas 
            Height          =   315
            Left            =   8190
            MaxLength       =   30
            TabIndex        =   5
            Text            =   "Text1"
            Top             =   390
            Width           =   3315
         End
         Begin VB.ComboBox cmb_ActEco 
            Height          =   315
            Left            =   2010
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   3030
            Width           =   765
         End
         Begin EditLib.fpDateTime ipp_FecNac 
            Height          =   315
            Left            =   2010
            TabIndex        =   7
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
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Afp:"
            Height          =   195
            Left            =   6210
            TabIndex        =   56
            Top             =   2040
            Width           =   285
         End
         Begin VB.Label Label17 
            Caption         =   "E-mail:"
            Height          =   285
            Left            =   6210
            TabIndex        =   40
            Top             =   2700
            Width           =   1485
         End
         Begin VB.Label Label16 
            Caption         =   "Teléfono Celular:"
            Height          =   285
            Left            =   90
            TabIndex        =   39
            Top             =   2700
            Width           =   1485
         End
         Begin VB.Label Label15 
            Caption         =   "Profesión o Actividad:"
            Height          =   315
            Left            =   90
            TabIndex        =   38
            Top             =   2370
            Width           =   1905
         End
         Begin VB.Label Label14 
            Caption         =   "Nivel de Estudio:"
            Height          =   315
            Left            =   90
            TabIndex        =   37
            Top             =   2040
            Width           =   1905
         End
         Begin VB.Label Label11 
            Caption         =   "Distrito Nacimiento:"
            Height          =   315
            Left            =   6210
            TabIndex        =   36
            Top             =   1710
            Width           =   1905
         End
         Begin VB.Label Label10 
            Caption         =   "Provincia Nacimiento:"
            Height          =   315
            Left            =   90
            TabIndex        =   35
            Top             =   1710
            Width           =   1905
         End
         Begin VB.Label Label9 
            Caption         =   "Dpto. Nacimiento:"
            Height          =   315
            Left            =   6210
            TabIndex        =   34
            Top             =   1380
            Width           =   1905
         End
         Begin VB.Label Label8 
            Caption         =   "Nacionalidad:"
            Height          =   315
            Left            =   90
            TabIndex        =   33
            Top             =   1380
            Width           =   1905
         End
         Begin VB.Label Label7 
            Caption         =   "Fecha de Nacimiento:"
            Height          =   315
            Left            =   90
            TabIndex        =   32
            Top             =   1050
            Width           =   1905
         End
         Begin VB.Label Label5 
            Caption         =   "Nombres:"
            Height          =   285
            Left            =   90
            TabIndex        =   31
            Top             =   720
            Width           =   1485
         End
         Begin VB.Label Label3 
            Caption         =   "Apellido Paterno:"
            Height          =   285
            Left            =   90
            TabIndex        =   30
            Top             =   60
            Width           =   1485
         End
         Begin VB.Label Label4 
            Caption         =   "Apellido Materno:"
            Height          =   285
            Left            =   90
            TabIndex        =   29
            Top             =   390
            Width           =   1485
         End
         Begin VB.Label Label29 
            Caption         =   "Apellido Casada:"
            Height          =   285
            Left            =   6210
            TabIndex        =   28
            Top             =   390
            Width           =   1485
         End
         Begin VB.Label Label6 
            Caption         =   "Registra Activ. Econ.:"
            Height          =   285
            Left            =   90
            TabIndex        =   27
            Top             =   3060
            Width           =   1785
         End
      End
      Begin Threed.SSPanel SSPanel8 
         Height          =   735
         Left            =   30
         TabIndex        =   41
         Top             =   750
         Width           =   11565
         _Version        =   65536
         _ExtentX        =   20399
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
         Begin VB.CommandButton cmd_ActEco 
            Height          =   585
            Left            =   1830
            Picture         =   "OpeTra_frm_164.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Actividades Económicas"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Cancel 
            Height          =   585
            Left            =   1230
            Picture         =   "OpeTra_frm_164.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Cancelar Modificación de Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Editar 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_164.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Modificar Datos del Cónyuge"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10950
            Picture         =   "OpeTra_frm_164.frx":0C34
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Salir de la Ventana"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Grabar 
            Height          =   585
            Left            =   630
            Picture         =   "OpeTra_frm_164.frx":1076
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Grabar Datos"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   435
         Left            =   30
         TabIndex        =   42
         Top             =   1530
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
         Begin Threed.SSPanel pnl_NomCli 
            Height          =   315
            Left            =   2010
            TabIndex        =   43
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
            Caption         =   "Cliente:"
            Height          =   315
            Left            =   90
            TabIndex        =   44
            Top             =   60
            Width           =   1575
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   435
         Left            =   30
         TabIndex        =   45
         Top             =   2010
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
            Left            =   2010
            TabIndex        =   52
            Top             =   60
            Width           =   9525
            _Version        =   65536
            _ExtentX        =   16801
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
         Begin VB.Label Label2 
            Caption         =   "Docum. de Identidad:"
            Height          =   315
            Left            =   90
            TabIndex        =   53
            Top             =   60
            Width           =   1725
         End
      End
   End
End
Attribute VB_Name = "frm_MntCli_55"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_Paises()   As moddat_tpo_Genera
Dim l_arr_Profes()   As moddat_tpo_Genera
Dim l_arr_PrvEst()   As moddat_tpo_Genera
Dim l_str_Paises     As String
Dim l_str_DptNac     As String
Dim l_str_PrvNac     As String
Dim l_str_DstNac     As String
Dim l_str_Profes     As String
Dim l_int_FlgCmb     As Integer
Dim l_int_RelLab     As Integer
Dim l_int_VinTDo     As Integer
Dim l_str_VinNDo     As String
Dim l_int_VinTip     As Integer
Dim l_int_RelAcc     As Integer
Dim l_int_AccTDo     As Integer
Dim l_str_AccNDo     As String
Dim l_int_AccVin     As Integer

Private Sub cmb_ActEco_Click()
   Call gs_SetFocus(cmd_Grabar)
End Sub

Private Sub cmb_ActEco_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_ActEco_Click
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
         
            Call gs_SetFocus(cmb_NivEst)
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
   
      Call gs_SetFocus(cmb_NivEst)
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
         
            Call gs_SetFocus(cmb_NivEst)
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
   
         Call gs_SetFocus(cmb_NivEst)
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
         Call gs_SetFocus(cmb_NivEst)
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
      
      Call gs_SetFocus(cmb_NivEst)
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

Private Sub cmd_ActEco_Click()
   If cmb_ActEco.ItemData(cmb_ActEco.ListIndex) = 2 Then
      MsgBox "No se pueden ingresar Actividades Económicas para el Cónyuge porque no labora.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   moddat_g_str_CygNom = txt_ApePat.Text & " " & txt_ApeMat.Text & " " & txt_Nombre.Text
   modmip_g_int_TipCli = 2
   
   frm_MntCli_56.Show 1
End Sub

Private Sub cmd_Cancel_Click()
   Call fs_Limpia
   Call fs_Cargar_Datos
   Call fs_Activa(False)
End Sub

Private Sub cmd_Editar_Click()
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
   
   
   'Si Cliente Titular es Femenino
   If frm_MntCli_52.cmb_CodSex.ItemData(frm_MntCli_52.cmb_CodSex.ListIndex) = 2 Then
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
   
   If cmb_ActEco.ListIndex = -1 Then
      MsgBox "Debe seleccionar si el Cónyuge desarrolla Actividad Económica.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_ActEco)
      Exit Sub
   End If
   
   
   'Rango de Edades del Cliente
   If Len(Trim(moddat_g_str_CodPrd)) > 0 Then
      If cmb_ActEco.ItemData(cmb_ActEco.ListIndex) = 1 Then
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
   End If
   
   If MsgBox("¿Está seguro de grabar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   'Grabando Información del Cliente
   If modmip_g_int_FlgGrb_2 = 1 Then
      g_str_Parame = "USP_CLI_DATGEN_NUEVO ("
      g_str_Parame = g_str_Parame & "2, "
   Else
      g_str_Parame = "USP_CLI_DATGEN_MODIFICA ("
   End If
   
   g_str_Parame = g_str_Parame & CStr(moddat_g_int_CygTDo) & ", "
   g_str_Parame = g_str_Parame & "'" & moddat_g_str_CygNDo & "', "
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
   g_str_Parame = g_str_Parame & CStr(frm_MntCli_52.cmb_EstCiv.ItemData(frm_MntCli_52.cmb_EstCiv.ListIndex)) & ", "
   
   If frm_MntCli_52.cmb_EstCiv.ItemData(frm_MntCli_52.cmb_EstCiv.ListIndex) = 2 Then
      g_str_Parame = g_str_Parame & CStr(frm_MntCli_52.cmb_RegCyg.ItemData(frm_MntCli_52.cmb_RegCyg.ListIndex)) & ", "
   Else
      g_str_Parame = g_str_Parame & "0, "
   End If
   
   g_str_Parame = g_str_Parame & CStr(cmb_NivEst.ItemData(cmb_NivEst.ListIndex)) & ", "
   g_str_Parame = g_str_Parame & "'" & l_arr_Profes(cmb_Profes.ListIndex + 1).Genera_Codigo & "', "
   
   If frm_MntCli_52.cmb_CodSex.ItemData(frm_MntCli_52.cmb_CodSex.ListIndex) = 1 Then
      g_str_Parame = g_str_Parame & "2, "
   Else
      g_str_Parame = g_str_Parame & "1, "
   End If
   
   g_str_Parame = g_str_Parame & Format(CDate(ipp_FecNac.Text), "yyyymmdd") & ", "
   
   If CInt(l_arr_Paises(cmb_Paises.ListIndex + 1).Genera_Codigo) = 4028 Then
      g_str_Parame = g_str_Parame & "'" & Format(cmb_DptNac.ItemData(cmb_DptNac.ListIndex), "00") & Format(cmb_PrvNac.ItemData(cmb_PrvNac.ListIndex), "00") & Format(cmb_DstNac.ItemData(cmb_DstNac.ListIndex), "00") & "', "
   Else
      g_str_Parame = g_str_Parame & "'000000', "
   End If
   g_str_Parame = g_str_Parame & "'" & l_arr_Paises(cmb_Paises.ListIndex + 1).Genera_Codigo & "', "
   
   g_str_Parame = g_str_Parame & CStr(frm_MntCli_52.ipp_NumDep.Value) & ", "
   g_str_Parame = g_str_Parame & CStr(frm_MntCli_52.ipp_DepEc1.Value) & ", "
   g_str_Parame = g_str_Parame & CStr(frm_MntCli_52.ipp_DepEc2.Value) & ", "
   g_str_Parame = g_str_Parame & CStr(frm_MntCli_52.ipp_DepEc3.Value) & ", "
   g_str_Parame = g_str_Parame & CStr(frm_MntCli_52.ipp_DepEc4.Value) & ", "
   g_str_Parame = g_str_Parame & CStr(frm_MntCli_52.ipp_DepEc5.Value) & ", "
   
   If CInt(l_arr_Paises(frm_MntCli_52.cmb_PaiRes.ListIndex + 1).Genera_Codigo) = 4028 Then
      g_str_Parame = g_str_Parame & CStr(frm_MntCli_52.cmb_TipVia.ItemData(frm_MntCli_52.cmb_TipVia.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & "'" & frm_MntCli_52.txt_NomVia.Text & "', "
      g_str_Parame = g_str_Parame & "'" & frm_MntCli_52.txt_NumVia.Text & "', "
      g_str_Parame = g_str_Parame & "'" & frm_MntCli_52.txt_IntDpt.Text & "', "
      g_str_Parame = g_str_Parame & CStr(frm_MntCli_52.cmb_TipZon.ItemData(frm_MntCli_52.cmb_TipZon.ListIndex)) & ", "
      g_str_Parame = g_str_Parame & "'" & frm_MntCli_52.txt_NomZon.Text & "', "
      g_str_Parame = g_str_Parame & "'" & frm_MntCli_52.txt_Refere.Text & "', "
      g_str_Parame = g_str_Parame & "'" & Format(frm_MntCli_52.cmb_DptDir.ItemData(frm_MntCli_52.cmb_DptDir.ListIndex), "00") & Format(frm_MntCli_52.cmb_PrvDir.ItemData(frm_MntCli_52.cmb_PrvDir.ListIndex), "00") & Format(frm_MntCli_52.cmb_DstDir.ItemData(frm_MntCli_52.cmb_DstDir.ListIndex), "00") & "', "
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
   
   g_str_Parame = g_str_Parame & "'" & txt_Celula.Text & "', "
   g_str_Parame = g_str_Parame & "'" & frm_MntCli_52.txt_TelFij.Text & "', "
   
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
   
   g_str_Parame = g_str_Parame & "'" & l_arr_Paises(frm_MntCli_52.cmb_PaiRes.ListIndex + 1).Genera_Codigo & "', "
   
   If CInt(l_arr_Paises(frm_MntCli_52.cmb_PaiRes.ListIndex + 1).Genera_Codigo) <> 4028 Then
      g_str_Parame = g_str_Parame & "'" & frm_MntCli_52.txt_Direcc.Text & "', "
      g_str_Parame = g_str_Parame & "'" & l_arr_PrvEst(frm_MntCli_52.cmb_PrvEst.ListIndex + 1).Genera_Codigo & "', "
      g_str_Parame = g_str_Parame & "'" & frm_MntCli_52.txt_CodPos.Text & "', "
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
   g_str_Parame = g_str_Parame & CStr(cmb_ActEco.ItemData(cmb_ActEco.ListIndex)) & ", "
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
      If modmip_g_int_FlgGrb_2 = 1 Then
         MsgBox "Error al ejecutar el Procedimiento USP_CLI_DATGEN_NUEVO.", vbCritical, modgen_g_str_NomPlt
         Exit Sub
      Else
         MsgBox "Error al ejecutar el Procedimiento USP_CLI_DATGEN_MODIFICA.", vbCritical, modgen_g_str_NomPlt
         Exit Sub
      End If
   End If
   
   'Actualizar Datos del Cónyuge sobre cliente
   g_str_Parame = "USP_CLI_DATGEN_CONYUGE ("
   
   g_str_Parame = g_str_Parame & CStr(moddat_g_int_TipDoc) & ", "
   g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumDoc & "', "
   g_str_Parame = g_str_Parame & CStr(moddat_g_int_CygTDo) & ", "
   g_str_Parame = g_str_Parame & "'" & moddat_g_str_CygNDo & "', "
   
   'Datos de Auditoria
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
      MsgBox "Error al ejecutar el Procedimiento USP_CLI_ACTECO_CONYUGE.", vbCritical, modgen_g_str_NomPlt
      
      Exit Sub
   End If
   
   'Actualizar Datos del Cónyuge sobre cliente
   g_str_Parame = "USP_CLI_DATGEN_CONYUGE ("
   
   g_str_Parame = g_str_Parame & CStr(moddat_g_int_CygTDo) & ", "
   g_str_Parame = g_str_Parame & "'" & moddat_g_str_CygNDo & "', "
   g_str_Parame = g_str_Parame & CStr(moddat_g_int_TipDoc) & ", "
   g_str_Parame = g_str_Parame & "'" & moddat_g_str_NumDoc & "', "
   
   'Datos de Auditoria
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "') "
      
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
      MsgBox "Error al ejecutar el Procedimiento USP_CLI_ACTECO_CONYUGE.", vbCritical, modgen_g_str_NomPlt
      
      Exit Sub
   End If
   
   modmip_g_int_FlgAct_3 = 2
   moddat_g_int_FlgAct = 2
   
   MsgBox "Los datos se grabaron correctamente.", vbInformation, modgen_g_str_NomPlt
   
   Call fs_Activa(False)
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Call gs_CentraForm(Me)
   Me.Caption = modgen_g_str_NomPlt
   pnl_NomCli.Caption = CStr(moddat_g_int_TipDoc) & " - " & moddat_g_str_NumDoc & " / " & moddat_g_str_NomCli
   pnl_DocIde.Caption = moddat_g_str_CygTDo & " - " & moddat_g_str_CygNDo
   pnl_RelLab.Visible = False
   pnl_RelAcc.Visible = False
   
   Call fs_Inicio
   
   If modmip_g_int_FlgGrb_2 = 1 Then
      Call fs_Activa(True)
      cmd_Cancel.Enabled = False
      Call fs_Limpia
   Else
      Call fs_Limpia
      Call fs_Cargar_Datos
      Call fs_Activa(False)
   End If
   
   'Verificando Relación Laboral
   Call modmip_gs_RelLab(moddat_g_int_CygTDo, moddat_g_str_CygNDo, l_int_RelLab, l_int_VinTDo, l_str_VinNDo, l_int_VinTip)
   
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
   Call modmip_gs_RelAcc(moddat_g_int_CygTDo, moddat_g_str_CygNDo, l_int_RelAcc, l_int_AccTDo, l_str_AccNDo, l_int_AccVin)

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

Private Sub fs_Inicio()
   Call moddat_gs_Carga_LisIte_Combo(cmb_DocAlt, 1, "214")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TDoAlt, 1, "231")
   Call moddat_gs_Carga_LisIte_Combo(cmb_NivEst, 1, "209")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipAfp, 1, "517")
   Call moddat_gs_Carga_LisIte(cmb_Paises, l_arr_Paises, 1, "500")
   Call moddat_gs_Carga_LisIte(cmb_Profes, l_arr_Profes, 1, "501")
   Call moddat_gs_Carga_LisIte_Combo(cmb_ActEco, 1, "214")
   
   Call moddat_gs_Carga_Depart(cmb_DptNac)
   Call modmip_gs_Carga_CiuExt_Arregl(l_arr_PrvEst, l_arr_Paises(frm_MntCli_52.cmb_PaiRes.ListIndex + 1).Genera_Codigo)
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
   ipp_FecNac.Text = Format(date, "dd/mm/yyyy")
   cmb_Paises.ListIndex = -1
   cmb_DptNac.ListIndex = -1
   cmb_PrvNac.Clear
   cmb_DstNac.Clear
   cmb_DptNac.Enabled = False
   cmb_PrvNac.Enabled = False
   cmb_DstNac.Enabled = False
   cmb_NivEst.ListIndex = -1
   cmb_TipAfp.ListIndex = -1
   cmb_Profes.ListIndex = -1
   txt_Celula.Text = ""
   txt_DirEle.Text = ""
   chk_DirEle.Value = 0
   chk_DirEle.Enabled = False
End Sub

Private Sub ipp_FecNac_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_Paises)
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
      Call gs_SetFocus(cmb_ActEco)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "-@_.")
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
      Call gs_SetFocus(ipp_FecNac)
   Else
      KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & "- '")
   End If
End Sub

Private Sub fs_Activa(ByVal p_Habilita As Integer)
   cmb_DocAlt.Enabled = p_Habilita
   cmb_TDoAlt.Enabled = p_Habilita
   txt_NDoAlt.Enabled = p_Habilita
   
   txt_ApePat.Enabled = p_Habilita
   txt_ApeMat.Enabled = p_Habilita
   txt_ApeCas.Enabled = p_Habilita
   txt_Nombre.Enabled = p_Habilita
   ipp_FecNac.Enabled = p_Habilita
   cmb_Paises.Enabled = p_Habilita
   cmb_DptNac.Enabled = p_Habilita
   cmb_PrvNac.Enabled = p_Habilita
   cmb_DstNac.Enabled = p_Habilita
   cmb_NivEst.Enabled = p_Habilita
   cmb_TipAfp.Enabled = p_Habilita
   cmb_Profes.Enabled = p_Habilita
   txt_Celula.Enabled = p_Habilita
   txt_DirEle.Enabled = p_Habilita
   chk_DirEle.Enabled = p_Habilita
   cmb_ActEco.Enabled = p_Habilita
   
   cmd_Grabar.Enabled = p_Habilita
   cmd_Cancel.Enabled = p_Habilita
   cmd_Editar.Enabled = Not p_Habilita
   cmd_ActEco.Enabled = Not p_Habilita
End Sub

Private Sub fs_Cargar_Datos()
   g_str_Parame = "SELECT * FROM CLI_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "DATGEN_TIPDOC = " & CStr(moddat_g_int_CygTDo) & " AND "
   g_str_Parame = g_str_Parame & "DATGEN_NUMDOC = '" & moddat_g_str_CygNDo & "' "
   
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
      
      Call gs_BuscarCombo_Item(cmb_ActEco, CInt(Left(g_rst_Princi!DATGEN_ACTECO, 2)))
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub



