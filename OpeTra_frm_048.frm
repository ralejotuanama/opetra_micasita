VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_Caj_MovDia_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   10110
   ClientLeft      =   3225
   ClientTop       =   2130
   ClientWidth     =   15150
   Icon            =   "OpeTra_frm_048.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10110
   ScaleWidth      =   15150
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   10095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15135
      _Version        =   65536
      _ExtentX        =   26696
      _ExtentY        =   17806
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
         Height          =   1095
         Left            =   30
         TabIndex        =   1
         Top             =   780
         Width           =   15045
         _Version        =   65536
         _ExtentX        =   26538
         _ExtentY        =   1931
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
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   390
            Width           =   3795
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   675
            Left            =   12870
            Picture         =   "OpeTra_frm_048.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Buscar Registros"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   675
            Left            =   13590
            Picture         =   "OpeTra_frm_048.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Limpiar Datos"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   14310
            Picture         =   "OpeTra_frm_048.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   675
         End
         Begin VB.ComboBox cmb_CodBan 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   60
            Width           =   10995
         End
         Begin EditLib.fpDateTime ipp_FecIni 
            Height          =   315
            Left            =   1560
            TabIndex        =   6
            Top             =   720
            Width           =   1395
            _Version        =   196608
            _ExtentX        =   2461
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
         Begin EditLib.fpDateTime ipp_FecFin 
            Height          =   315
            Left            =   2970
            TabIndex        =   25
            Top             =   720
            Width           =   1395
            _Version        =   196608
            _ExtentX        =   2461
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
         Begin VB.Label Label7 
            Caption         =   "Nro. Cuenta:"
            Height          =   285
            Left            =   90
            TabIndex        =   24
            Top             =   390
            Width           =   1185
         End
         Begin VB.Label Label3 
            Caption         =   "Banco:"
            Height          =   315
            Left            =   90
            TabIndex        =   8
            Top             =   60
            Width           =   765
         End
         Begin VB.Label Label10 
            Caption         =   "Rango de Fechas:"
            Height          =   315
            Left            =   90
            TabIndex        =   7
            Top             =   690
            Width           =   1365
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   705
         Left            =   30
         TabIndex        =   9
         Top             =   30
         Width           =   15045
         _Version        =   65536
         _ExtentX        =   26538
         _ExtentY        =   1244
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
            Left            =   630
            TabIndex        =   10
            Top             =   30
            Width           =   6795
            _Version        =   65536
            _ExtentX        =   11986
            _ExtentY        =   1032
            _StockProps     =   15
            Caption         =   "Cobro por Banco - Movimientos Diarios"
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
            Picture         =   "OpeTra_frm_048.frx":0A62
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   7245
         Left            =   30
         TabIndex        =   11
         Top             =   1920
         Width           =   15045
         _Version        =   65536
         _ExtentX        =   26538
         _ExtentY        =   12779
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
            Height          =   6855
            Left            =   60
            TabIndex        =   12
            Top             =   360
            Width           =   14925
            _ExtentX        =   26326
            _ExtentY        =   12091
            _Version        =   393216
            Rows            =   30
            Cols            =   9
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   49152
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_Tit_NumMov 
            Height          =   285
            Left            =   90
            TabIndex        =   13
            Top             =   60
            Width           =   1035
            _Version        =   65536
            _ExtentX        =   1826
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro. Movim."
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
         Begin Threed.SSPanel pnl_Tit_TipMov 
            Height          =   285
            Left            =   2280
            TabIndex        =   14
            Top             =   60
            Width           =   3675
            _Version        =   65536
            _ExtentX        =   6482
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Tipo de Movimiento"
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
         Begin Threed.SSPanel pnl_Tit_Import 
            Height          =   285
            Left            =   13710
            TabIndex        =   15
            Top             =   60
            Width           =   1005
            _Version        =   65536
            _ExtentX        =   1773
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
         Begin Threed.SSPanel pnl_Tit_NumRef 
            Height          =   285
            Left            =   5940
            TabIndex        =   16
            Top             =   60
            Width           =   1575
            _Version        =   65536
            _ExtentX        =   2778
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Número Referencia"
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
         Begin Threed.SSPanel pnl_Tit_DoiCli 
            Height          =   285
            Left            =   7500
            TabIndex        =   20
            Top             =   60
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "ID Cliente"
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
         Begin Threed.SSPanel pnl_Tit_NomCli 
            Height          =   285
            Left            =   8700
            TabIndex        =   21
            Top             =   60
            Width           =   4185
            _Version        =   65536
            _ExtentX        =   7382
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nombre Cliente"
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
         Begin Threed.SSPanel pnl_Tit_Moneda 
            Height          =   285
            Left            =   12870
            TabIndex        =   22
            Top             =   60
            Width           =   855
            _Version        =   65536
            _ExtentX        =   1508
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
         Begin Threed.SSPanel pnl_Tit_FecMov 
            Height          =   285
            Left            =   1110
            TabIndex        =   26
            Top             =   60
            Width           =   1185
            _Version        =   65536
            _ExtentX        =   2090
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Movim."
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
      End
      Begin Threed.SSPanel SSPanel12 
         Height          =   795
         Left            =   30
         TabIndex        =   17
         Top             =   9210
         Width           =   15045
         _Version        =   65536
         _ExtentX        =   26538
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
         Begin VB.CommandButton cmd_Imprim 
            Height          =   675
            Left            =   60
            Picture         =   "OpeTra_frm_048.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Imprimir Listado"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_VerCom 
            Height          =   675
            Left            =   14310
            Picture         =   "OpeTra_frm_048.frx":11AE
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Ver Comprobante"
            Top             =   30
            Width           =   675
         End
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   2160
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
      End
   End
End
Attribute VB_Name = "frm_Caj_MovDia_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_CodBan()  As moddat_tpo_Genera
Dim l_arr_CtaBan()  As moddat_tpo_Genera

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
   Call gs_SetFocus(ipp_FecIni)
End Sub

Private Sub cmb_CtaBan_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_CtaBan_Click
   End If
End Sub

Private Sub cmd_Buscar_Click()
   If cmb_CodBan.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Banco del que desea ver las Operaciones.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_CodBan)
      Exit Sub
   End If
   
   If cmb_CtaBan.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Cuenta de la que desea ver las Operaciones.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_CtaBan)
      Exit Sub
   End If
   
   If CDate(ipp_FecFin.Text) < CDate(ipp_FecIni.Text) Then
      MsgBox "La Fecha de Fin no puede ser mayor a la Fecha de Inicio.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecFin)
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_Buscar
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Imprim_Click()
   If MsgBox("¿Está seguro de registrar el pago?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If

   Screen.MousePointer = 11
   
   'Borrando Spool Local
   moddat_g_bdt_Report.Execute "DELETE FROM RPT_CABGEN"
   moddat_g_bdt_Report.Execute "DELETE FROM RPT_COBBC1"
   DoEvents
   
   'Grabando en DAO (Cabecera)
   moddat_g_str_CadDAO = "SELECT * FROM RPT_CABGEN WHERE CABGEN_PRODUC = ' '"
   Set moddat_g_rst_RecDAO = moddat_g_bdt_Report.OpenRecordset(moddat_g_str_CadDAO, dbOpenDynaset)
   
   moddat_g_rst_RecDAO.AddNew
      
   moddat_g_rst_RecDAO("CABGEN_NOMBCO") = cmb_CodBan.Text
   moddat_g_rst_RecDAO("CABGEN_NUMCTA") = l_arr_CtaBan(cmb_CtaBan.ListIndex + 1).Genera_Codigo
   moddat_g_rst_RecDAO("CABGEN_FECINI") = ipp_FecIni.Text
   moddat_g_rst_RecDAO("CABGEN_FECFIN") = ipp_FecFin.Text
   
   moddat_g_rst_RecDAO.Update
   DoEvents
   
   moddat_g_rst_RecDAO.Close
   DoEvents
      
   'Procesando Detalle
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM OPE_CAJMOV WHERE "
   g_str_Parame = g_str_Parame & "CAJMOV_SUCMOV = '" & modgen_g_str_CodSuc & "' AND "
   g_str_Parame = g_str_Parame & "CAJMOV_CODBAN = '" & opecaj_g_str_CodBan & "' AND "
   g_str_Parame = g_str_Parame & "CAJMOV_NUMCTA = '" & opecaj_g_str_CtaBan & "' AND "
   g_str_Parame = g_str_Parame & "CAJMOV_FECMOV >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND "
   g_str_Parame = g_str_Parame & "CAJMOV_FECMOV <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd")
   g_str_Parame = g_str_Parame & "ORDER BY CAJMOV_NUMMOV ASC, CAJMOV_FECMOV ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If

   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         'Grabando en DAO (Detalle)
         moddat_g_str_CadDAO = "SELECT * FROM RPT_COBBC1 WHERE COBBC1_FECMOV = ' '"
         Set moddat_g_rst_RecDAO = moddat_g_bdt_Report.OpenRecordset(moddat_g_str_CadDAO, dbOpenDynaset)
         
         moddat_g_rst_RecDAO.AddNew
            
         moddat_g_rst_RecDAO("COBBC1_FECMOV") = gf_FormatoFecha(CStr(g_rst_Princi!CAJMOV_FECMOV))
         moddat_g_rst_RecDAO("COBBC1_NUMMOV") = Format(g_rst_Princi!CAJMOV_NUMMOV, "00000")
         moddat_g_rst_RecDAO("COBBC1_TIPMOV") = moddat_gf_Consulta_ParDes("301", CStr(g_rst_Princi!CAJMOV_TIPMOV))
         moddat_g_rst_RecDAO("COBBC1_TIPREG") = moddat_gf_Consulta_ParDes("239", CStr(g_rst_Princi!CAJMOV_TIPREG))
         
         If g_rst_Princi!CAJMOV_TIPMOV = 1101 Then
            moddat_g_rst_RecDAO("COBBC1_NUMREF") = gf_Formato_NumSol(Trim(g_rst_Princi!CAJMOV_NUMOPE & ""))
         Else
            moddat_g_rst_RecDAO("COBBC1_NUMREF") = gf_Formato_NumOpe(Trim(g_rst_Princi!CAJMOV_NUMOPE & ""))
         End If
         
         moddat_g_rst_RecDAO("COBBC1_DOCIDE") = CStr(g_rst_Princi!CAJMOV_TIPDOC) & "-" & Trim(g_rst_Princi!CAJMOV_NUMDOC & "")
         moddat_g_rst_RecDAO("COBBC1_NOMCLI") = moddat_gf_Buscar_NomCli(g_rst_Princi!CAJMOV_TIPDOC, Trim(g_rst_Princi!CAJMOV_NUMDOC & ""))
         moddat_g_rst_RecDAO("COBBC1_MONEDA") = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!CAJMOV_MONPAG))
         
         moddat_g_rst_RecDAO("COBBC1_IMPORT") = g_rst_Princi!CAJMOV_IMPTOT
         moddat_g_rst_RecDAO("COBBC1_IMPITF") = g_rst_Princi!CAJMOV_ITFIMP
         moddat_g_rst_RecDAO("COBBC1_IMPNET") = g_rst_Princi!CAJMOV_IMPPAG
         
         moddat_g_rst_RecDAO("COBBC1_FECPAG") = gf_FormatoFecha(CStr(g_rst_Princi!CAJMOV_FECDEP))
         
         If g_rst_Princi!CAJMOV_FECREC > 0 Then
            moddat_g_rst_RecDAO("COBBC1_FECREC") = gf_FormatoFecha(CStr(g_rst_Princi!CAJMOV_FECREC))
         End If
         
         moddat_g_rst_RecDAO.Update
         DoEvents
         
         moddat_g_rst_RecDAO.Close
         DoEvents
      
         g_rst_Princi.MoveNext
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Screen.MousePointer = 0

   'Generando Reporte
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_COBBCO_02.RPT"
   crp_Imprim.Action = 1
End Sub

Private Sub cmd_Limpia_Click()
   cmb_CodBan.ListIndex = -1
   cmb_CtaBan.Clear
   ipp_FecIni.Text = Format(Date - CDate(30), "dd/mm/yyyy")
   ipp_FecFin.Text = Format(Date, "dd/mm/yyyy")
   
   Call gs_LimpiaGrid(grd_Listad)
   
   Call fs_Activa(True)
   Call gs_SetFocus(cmb_CodBan)
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub cmd_VerCom_Click()
   Dim r_int_TipOpe As String
   
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   grd_Listad.Col = 0
   opecaj_g_str_NumMov = grd_Listad.Text
   
   grd_Listad.Col = 8
   opecaj_g_str_FecMov = grd_Listad.Text
   
   Call gs_RefrescaGrid(grd_Listad)
   
   frm_Caj_MovDia_02.Show 1
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call cmd_Limpia_Click
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Activa(ByVal p_Activa As Integer)
   cmb_CodBan.Enabled = p_Activa
   cmb_CtaBan.Enabled = p_Activa
   ipp_FecIni.Enabled = p_Activa
   ipp_FecFin.Enabled = p_Activa
   
   cmd_Buscar.Enabled = p_Activa
   
   grd_Listad.Enabled = Not p_Activa
   cmd_Imprim.Enabled = Not p_Activa
   cmd_VerCom.Enabled = Not p_Activa
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_FecSis
   
   grd_Listad.ColWidth(0) = 1025
   grd_Listad.ColWidth(1) = 1175
   grd_Listad.ColWidth(2) = 3665
   grd_Listad.ColWidth(3) = 1575
   grd_Listad.ColWidth(4) = 1205
   grd_Listad.ColWidth(5) = 4165
   grd_Listad.ColWidth(6) = 835
   grd_Listad.ColWidth(7) = 985
   grd_Listad.ColWidth(8) = 0
   
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignLeftCenter
   grd_Listad.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad.ColAlignment(4) = flexAlignCenterCenter
   grd_Listad.ColAlignment(5) = flexAlignLeftCenter
   grd_Listad.ColAlignment(6) = flexAlignCenterCenter
   grd_Listad.ColAlignment(7) = flexAlignRightCenter

   Call moddat_gs_Carga_LisIte(cmb_CodBan, l_arr_CodBan, 1, "516")
End Sub

Private Sub fs_Buscar()
   Call gs_LimpiaGrid(grd_Listad)

   opecaj_g_str_CodBan = l_arr_CodBan(cmb_CodBan.ListIndex + 1).Genera_Codigo
   opecaj_g_str_CtaBan = l_arr_CtaBan(cmb_CtaBan.ListIndex + 1).Genera_Codigo
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM OPE_CAJMOV WHERE "
   g_str_Parame = g_str_Parame & "CAJMOV_SUCMOV = '" & modgen_g_str_CodSuc & "' AND "
   g_str_Parame = g_str_Parame & "CAJMOV_CODBAN = '" & opecaj_g_str_CodBan & "' AND "
   g_str_Parame = g_str_Parame & "CAJMOV_NUMCTA = '" & opecaj_g_str_CtaBan & "' AND "
   g_str_Parame = g_str_Parame & "CAJMOV_FECMOV >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND "
   g_str_Parame = g_str_Parame & "CAJMOV_FECMOV <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd")
   g_str_Parame = g_str_Parame & "ORDER BY CAJMOV_NUMMOV ASC, CAJMOV_FECMOV ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If

   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
     g_rst_Princi.Close
     Set g_rst_Princi = Nothing
     
     MsgBox "No se han encontrado registros.", vbExclamation, modgen_g_con_PltPar
     Call gs_SetFocus(cmb_CodBan)
     Exit Sub
   End If
   
   Call fs_Activa(False)
   
   grd_Listad.Redraw = False
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      
      grd_Listad.Row = grd_Listad.Rows - 1
      
      grd_Listad.Col = 0
      grd_Listad.Text = Format(g_rst_Princi!CAJMOV_NUMMOV, "00000")
      
      grd_Listad.Col = 1
      grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!CAJMOV_FECMOV))
      
      grd_Listad.Col = 2
      grd_Listad.Text = CStr(g_rst_Princi!CAJMOV_TIPMOV) & " - " & moddat_gf_Consulta_ParDes("301", Format(g_rst_Princi!CAJMOV_TIPMOV, "000000"))
      
      grd_Listad.Col = 3
      
      If g_rst_Princi!CAJMOV_TIPMOV = 1101 Then
         grd_Listad.Text = gf_Formato_NumSol(Trim(g_rst_Princi!CAJMOV_NUMOPE & ""))
      Else
        grd_Listad.Text = gf_Formato_NumOpe(Trim(g_rst_Princi!CAJMOV_NUMOPE & ""))
      End If
      
      If g_rst_Princi!CAJMOV_TIPDOC > 0 Then
         grd_Listad.Col = 4
         grd_Listad.Text = CStr(g_rst_Princi!CAJMOV_TIPDOC) & "-" & Trim(g_rst_Princi!CAJMOV_NUMDOC & "")
         
         grd_Listad.Col = 5
         grd_Listad.Text = moddat_gf_Buscar_NomCli(g_rst_Princi!CAJMOV_TIPDOC, Trim(g_rst_Princi!CAJMOV_NUMDOC & ""))
      End If
      
      grd_Listad.Col = 6
      grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!CAJMOV_MONPAG))
      
      grd_Listad.Col = 7
      grd_Listad.Text = Format(g_rst_Princi!CAJMOV_IMPTOT, "###,###,##0.00")
      
      grd_Listad.Col = 8
      grd_Listad.Text = CStr(g_rst_Princi!CAJMOV_FECMOV)
      
      g_rst_Princi.MoveNext
   Loop
   
   grd_Listad.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   If grd_Listad.Rows > 0 Then
      cmd_Imprim.Enabled = True
      cmd_VerCom.Enabled = True
      grd_Listad.Enabled = True
   End If
   
   Call gs_UbiIniGrid(grd_Listad)
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub grd_Listad_DblClick()
   Call cmd_VerCom_Click
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

Private Sub ipp_FecFin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Buscar)
   End If
End Sub

Private Sub ipp_FecIni_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecFin)
   End If
End Sub

Private Sub pnl_Tit_DoiCli_Click()
   If Len(Trim(pnl_Tit_DoiCli.Tag)) = 0 Or pnl_Tit_DoiCli.Tag = "D" Then
      pnl_Tit_DoiCli.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 4, "C")
   Else
      pnl_Tit_DoiCli.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 4, "C-")
   End If
End Sub

Private Sub pnl_Tit_FecMov_Click()
   If Len(Trim(pnl_Tit_FecMov.Tag)) = 0 Or pnl_Tit_FecMov.Tag = "D" Then
      pnl_Tit_FecMov.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 8, "C")
   Else
      pnl_Tit_FecMov.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 8, "N-")
   End If
End Sub

Private Sub pnl_Tit_Import_Click()
   If Len(Trim(pnl_Tit_Import.Tag)) = 0 Or pnl_Tit_Import.Tag = "D" Then
      pnl_Tit_Import.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 7, "N")
   Else
      pnl_Tit_Import.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 7, "N-")
   End If
End Sub

Private Sub pnl_Tit_Moneda_Click()
   If Len(Trim(pnl_Tit_Moneda.Tag)) = 0 Or pnl_Tit_Moneda.Tag = "D" Then
      pnl_Tit_Moneda.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 6, "C")
   Else
      pnl_Tit_Moneda.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 6, "C-")
   End If
End Sub

Private Sub pnl_Tit_NomCli_Click()
   If Len(Trim(pnl_Tit_NomCli.Tag)) = 0 Or pnl_Tit_NomCli.Tag = "D" Then
      pnl_Tit_NomCli.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 5, "C")
   Else
      pnl_Tit_NomCli.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 5, "C-")
   End If
End Sub

Private Sub pnl_Tit_NumMov_Click()
   If Len(Trim(pnl_Tit_NumMov.Tag)) = 0 Or pnl_Tit_NumMov.Tag = "D" Then
      pnl_Tit_NumMov.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 0, "C")
   Else
      pnl_Tit_NumMov.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 0, "C-")
   End If
End Sub

Private Sub pnl_Tit_NumRef_Click()
   If Len(Trim(pnl_Tit_NumRef.Tag)) = 0 Or pnl_Tit_NumRef.Tag = "D" Then
      pnl_Tit_NumRef.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 3, "C")
   Else
      pnl_Tit_NumRef.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 3, "C-")
   End If
End Sub

Private Sub pnl_Tit_TipMov_Click()
   If Len(Trim(pnl_Tit_TipMov.Tag)) = 0 Or pnl_Tit_TipMov.Tag = "D" Then
      pnl_Tit_TipMov.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 2, "C")
   Else
      pnl_Tit_TipMov.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 2, "C-")
   End If
End Sub
