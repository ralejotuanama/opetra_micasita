VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_Con_OpeFin_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   9750
   ClientLeft      =   330
   ClientTop       =   2370
   ClientWidth     =   15120
   Icon            =   "OpeTra_frm_128.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9750
   ScaleWidth      =   15120
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   9765
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   15135
      _Version        =   65536
      _ExtentX        =   26696
      _ExtentY        =   17224
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
         Height          =   645
         Left            =   30
         TabIndex        =   24
         Top             =   780
         Width           =   15045
         _Version        =   65536
         _ExtentX        =   26538
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
            Left            =   3030
            Picture         =   "OpeTra_frm_128.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_BusCli 
            Height          =   585
            Left            =   2430
            Picture         =   "OpeTra_frm_128.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Buscar Movimiento"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   14430
            Picture         =   "OpeTra_frm_128.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_VerCom 
            Height          =   585
            Left            =   1830
            Picture         =   "OpeTra_frm_128.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Ver Comprobante"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   1230
            Picture         =   "OpeTra_frm_128.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Imprimir Listado"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   630
            Picture         =   "OpeTra_frm_128.frx":11AE
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Limpiar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_128.frx":14B8
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Buscar Registros"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   1095
         Left            =   30
         TabIndex        =   10
         Top             =   1470
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
         Begin VB.ComboBox cmb_SucAge 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   13425
         End
         Begin EditLib.fpDateTime ipp_FecIni 
            Height          =   315
            Left            =   1560
            TabIndex        =   1
            Top             =   390
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
            Left            =   1560
            TabIndex        =   2
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
         Begin VB.Label Label3 
            Caption         =   "Sucursal:"
            Height          =   225
            Left            =   60
            TabIndex        =   25
            Top             =   60
            Width           =   945
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha Fin:"
            Height          =   315
            Left            =   60
            TabIndex        =   23
            Top             =   690
            Width           =   1365
         End
         Begin VB.Label Label10 
            Caption         =   "Fecha Inicio:"
            Height          =   315
            Left            =   60
            TabIndex        =   11
            Top             =   390
            Width           =   1365
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   705
         Left            =   30
         TabIndex        =   12
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
            TabIndex        =   13
            Top             =   30
            Width           =   6795
            _Version        =   65536
            _ExtentX        =   11986
            _ExtentY        =   1032
            _StockProps     =   15
            Caption         =   "Consulta de Operaciones Financieras"
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
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   14190
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
         Begin VB.Image Image1 
            Height          =   480
            Left            =   60
            Picture         =   "OpeTra_frm_128.frx":17C2
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   7095
         Left            =   30
         TabIndex        =   14
         Top             =   2610
         Width           =   15045
         _Version        =   65536
         _ExtentX        =   26538
         _ExtentY        =   12515
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
            Height          =   6675
            Left            =   60
            TabIndex        =   3
            Top             =   360
            Width           =   14925
            _ExtentX        =   26326
            _ExtentY        =   11774
            _Version        =   393216
            Rows            =   30
            Cols            =   10
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_Tit_NumMov 
            Height          =   285
            Left            =   90
            TabIndex        =   15
            Top             =   60
            Width           =   1035
            _Version        =   65536
            _ExtentX        =   1826
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro. Movim."
            ForeColor       =   16777215
            BackColor       =   16384
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
            TabIndex        =   16
            Top             =   60
            Width           =   3675
            _Version        =   65536
            _ExtentX        =   6482
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Tipo de Movimiento"
            ForeColor       =   16777215
            BackColor       =   16384
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
            TabIndex        =   17
            Top             =   60
            Width           =   1005
            _Version        =   65536
            _ExtentX        =   1773
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Importe"
            ForeColor       =   16777215
            BackColor       =   16384
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
            TabIndex        =   18
            Top             =   60
            Width           =   1575
            _Version        =   65536
            _ExtentX        =   2778
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Número Referencia"
            ForeColor       =   16777215
            BackColor       =   16384
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
            TabIndex        =   19
            Top             =   60
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "ID Cliente"
            ForeColor       =   16777215
            BackColor       =   16384
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
            TabIndex        =   20
            Top             =   60
            Width           =   4185
            _Version        =   65536
            _ExtentX        =   7382
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nombre Cliente"
            ForeColor       =   16777215
            BackColor       =   16384
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
            TabIndex        =   21
            Top             =   60
            Width           =   855
            _Version        =   65536
            _ExtentX        =   1508
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Moneda"
            ForeColor       =   16777215
            BackColor       =   16384
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
            TabIndex        =   22
            Top             =   60
            Width           =   1185
            _Version        =   65536
            _ExtentX        =   2090
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Movim."
            ForeColor       =   16777215
            BackColor       =   16384
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
   End
End
Attribute VB_Name = "frm_Con_OpeFin_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_arr_SucAge()      As moddat_tpo_Genera
Dim l_str_Existe        As String
Dim l_int_MsjErr        As Integer

Private Sub cmd_Buscar_Click()
   If cmb_SucAge.ListIndex = -1 Then
      MsgBox "Debe seleccionar la Sucursal o Agencia.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_SucAge)
      Exit Sub
   End If
   If CDate(ipp_FecFin.Text) < CDate(ipp_FecIni.Text) Then
      MsgBox "La Fecha de Fin es menor a la Fecha de Inicio.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecFin)
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_Activa(False)
   Call fs_Buscar
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Limpia_Click()
   Call fs_Limpia
   Call fs_Activa(True)
   Call gs_SetFocus(cmb_SucAge)
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Imprim_Click()
Dim r_int_NumPag     As Integer
Dim r_int_NumIte     As Integer
   
   If MsgBox("¿Está seguro de Imprimir los Comprobantes?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   'Borrar Spool de PC (Cabecera)
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "DELETE FROM RPT_COMPGC "
   g_str_Parame = g_str_Parame & " WHERE COMPGC_CODTER = '" & modgen_g_str_NombPC & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
       Exit Sub
   End If
   
   'Borrar Spool de PC (Detalle)
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "DELETE FROM RPT_COMPGD "
   g_str_Parame = g_str_Parame & " WHERE COMPGD_CODTER = '" & modgen_g_str_NombPC & "' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 2) Then
       Exit Sub
   End If
   
   Screen.MousePointer = 11
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM OPE_CAJMOV "
   g_str_Parame = g_str_Parame & " WHERE CAJMOV_SUCMOV = '" & moddat_g_str_CodGrp & "' "
   g_str_Parame = g_str_Parame & "   AND CAJMOV_FECMOV >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " "
   g_str_Parame = g_str_Parame & "   AND CAJMOV_FECMOV <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd")
   g_str_Parame = g_str_Parame & " ORDER BY CAJMOV_NUMMOV ASC, CAJMOV_FECMOV ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      r_int_NumPag = 1
      r_int_NumIte = 1
      
      Do While Not g_rst_Princi.EOF
         Call opecaj_gs_ComPago(moddat_g_str_CodGrp, Format(g_rst_Princi!CAJMOV_NUMMOV, "00000"), CStr(g_rst_Princi!CAJMOV_FECMOV), r_int_NumPag, r_int_NumIte)
         g_rst_Princi.MoveNext
         DoEvents
         
         If r_int_NumIte = 1 Then
            r_int_NumIte = 2
         ElseIf r_int_NumIte = 2 Then
            r_int_NumIte = 1
            r_int_NumPag = r_int_NumPag + 1
         End If
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   Screen.MousePointer = 0
   
   'Se conecta al crystal report
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat

   'Se envia las tablas correspondientes en el orden que fueron utilizadas
   crp_Imprim.DataFiles(0) = "RPT_COMPGC"
   crp_Imprim.DataFiles(1) = "RPT_COMPGD"
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_COMPAG_01.RPT"
   crp_Imprim.SelectionFormula = "{RPT_COMPGC.COMPGC_CODTER} = '" & modgen_g_str_NombPC & "'"
   crp_Imprim.WindowShowPrintSetupBtn = True
   crp_Imprim.Destination = crptToWindow
   crp_Imprim.Action = 1
End Sub

Private Sub cmd_VerCom_Click()
Dim r_int_TipOpe As String
   
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   grd_Listad.Col = 9
   opecaj_g_str_NumMov = grd_Listad.Text
   grd_Listad.Col = 8
   opecaj_g_str_FecMov = grd_Listad.Text
   
   Call gs_RefrescaGrid(grd_Listad)
   
   frm_Con_OpeFin_02.Show 1
End Sub

Private Sub cmd_BusCli_Click()
   frm_Con_OpeFin_03.Show 1
End Sub

Private Sub cmd_ExpExc_Click()
   If grd_Listad.Rows = 0 Then
      MsgBox "No existe ninguna operación financiera.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   'Confirmacion
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
      
   l_int_MsjErr = 0
   Screen.MousePointer = 11
    
   '1102 - PAGO CUOTA CRED. HIPOT.
   Call fs_GenExc_CuoCHp
   
   '1101 - PAGO GASTOS DE CIERRE CRED. HIPOT.
   Call fs_GenExc_GtoCie
   
   '1103 - DESEMBOLSO CRED. HIPOT.
   Call fs_GenExc_Desemb
   
   '2101 - REVERSA PAGO GASTOS DE CIERRE
   Call fs_GenExc_Extrno
   
   Screen.MousePointer = 0
   If l_int_MsjErr > 0 Then
      MsgBox "No se han encontrado registros.", vbExclamation, modgen_g_str_NomPlt
   End If
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Limpia
   Call fs_Activa(True)
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   grd_Listad.ColWidth(0) = 1025
   grd_Listad.ColWidth(1) = 1175
   grd_Listad.ColWidth(2) = 3665
   grd_Listad.ColWidth(3) = 1575
   grd_Listad.ColWidth(4) = 1205
   grd_Listad.ColWidth(5) = 4165
   grd_Listad.ColWidth(6) = 835
   grd_Listad.ColWidth(7) = 985
   grd_Listad.ColWidth(8) = 0
   grd_Listad.ColWidth(9) = 0
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignLeftCenter
   grd_Listad.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad.ColAlignment(4) = flexAlignCenterCenter
   grd_Listad.ColAlignment(5) = flexAlignLeftCenter
   grd_Listad.ColAlignment(6) = flexAlignCenterCenter
   grd_Listad.ColAlignment(7) = flexAlignRightCenter
   
   moddat_g_str_Codigo = "000001"
   Call moddat_gs_Carga_SucAge(cmb_SucAge, l_arr_SucAge, moddat_g_str_Codigo)
End Sub

Private Sub fs_Limpia()
   cmb_SucAge.ListIndex = -1
   ipp_FecIni.Text = Format(date, "dd/mm/yyyy")
   ipp_FecFin.Text = Format(date, "dd/mm/yyyy")
   Call gs_LimpiaGrid(grd_Listad)
End Sub

Private Sub fs_Activa(ByVal p_Activa As Integer)
   cmb_SucAge.Enabled = p_Activa
   ipp_FecIni.Enabled = p_Activa
   ipp_FecFin.Enabled = p_Activa
   cmd_Buscar.Enabled = p_Activa
   grd_Listad.Enabled = Not p_Activa
   cmd_Imprim.Enabled = Not p_Activa
   cmd_VerCom.Enabled = Not p_Activa
End Sub

Private Sub cmb_SucAge_Click()
   Call gs_SetFocus(ipp_FecIni)
End Sub

Private Sub cmb_SucAge_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_SucAge_Click
   End If
End Sub

Private Sub ipp_FecIni_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecFin)
   End If
End Sub

Private Sub ipp_FecFin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Buscar)
   End If
End Sub

Private Sub fs_Buscar()
   moddat_g_str_CodGrp = l_arr_SucAge(cmb_SucAge.ListIndex + 1).Genera_Codigo
   moddat_g_str_DesGrp = l_arr_SucAge(cmb_SucAge.ListIndex + 1).Genera_Nombre
   Call gs_LimpiaGrid(grd_Listad)
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM OPE_CAJMOV "
   g_str_Parame = g_str_Parame & " WHERE CAJMOV_SUCMOV = '" & moddat_g_str_CodGrp & "' "
   g_str_Parame = g_str_Parame & "   AND CAJMOV_FECMOV >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " "
   g_str_Parame = g_str_Parame & "   AND CAJMOV_FECMOV <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd")
   g_str_Parame = g_str_Parame & " ORDER BY CAJMOV_NUMMOV ASC, CAJMOV_FECMOV ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se han encontrado registros.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecIni)
      Exit Sub
   End If
   
   Call fs_Activa(False)
   grd_Listad.Redraw = False
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      
      grd_Listad.Col = 0
      grd_Listad.Text = Mid(CStr(g_rst_Princi!CAJMOV_FECMOV), 3, 2) & Format(g_rst_Princi!CAJMOV_NUMMOV, "00000")
      
      grd_Listad.Col = 1
      grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!CAJMOV_FECMOV))
      
      grd_Listad.Col = 2
      grd_Listad.Text = CStr(g_rst_Princi!CAJMOV_TIPMOV) & " - " & moddat_gf_Consulta_ParDes("301", Format(g_rst_Princi!CAJMOV_TIPMOV, "000000"))
      
      grd_Listad.Col = 3
      If g_rst_Princi!CAJMOV_TIPMOV = 1101 Or g_rst_Princi!CAJMOV_TIPMOV = 2101 Then
         grd_Listad.Text = gf_Formato_NumSol(Trim(g_rst_Princi!CAJMOV_NUMOPE & ""))
      ElseIf g_rst_Princi!CAJMOV_TIPMOV = 1105 Then
         grd_Listad.Text = CStr(Trim(g_rst_Princi!CAJMOV_NUMOPE & ""))
      Else
         grd_Listad.Text = gf_Formato_NumOpe(Trim(g_rst_Princi!CAJMOV_NUMOPE & ""))
      End If
      
      If g_rst_Princi!CAJMOV_TIPDOC > 0 Then
         grd_Listad.Col = 4
         grd_Listad.Text = CStr(g_rst_Princi!CAJMOV_TIPDOC) & "-" & Trim(g_rst_Princi!CAJMOV_NUMDOC & "")
         
         grd_Listad.Col = 5
         If g_rst_Princi!CAJMOV_TIPMOV = 1105 Then
            grd_Listad.Text = moddat_gf_Buscar_NomCli_PlanAhorro(g_rst_Princi!CAJMOV_TIPDOC, Trim(g_rst_Princi!CAJMOV_NUMDOC & ""))
         Else
            grd_Listad.Text = moddat_gf_Buscar_NomCli(g_rst_Princi!CAJMOV_TIPDOC, Trim(g_rst_Princi!CAJMOV_NUMDOC & ""))
         End If
      End If
      
      grd_Listad.Col = 6
      grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!CAJMOV_MONPAG))
      
      grd_Listad.Col = 7
      grd_Listad.Text = Format(g_rst_Princi!CAJMOV_IMPTOT, "###,###,##0.00")
      
      grd_Listad.Col = 8
      grd_Listad.Text = CStr(g_rst_Princi!CAJMOV_FECMOV)
      
      grd_Listad.Col = 9
      grd_Listad.Text = CStr(g_rst_Princi!CAJMOV_NUMMOV)
      
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

'************************************
'1102 - PAGO CUOTA CRED. HIPOT.
Private Sub fs_GenExc_CuoCHp()
Dim r_obj_Excel      As Excel.Application
Dim r_int_NroFil     As Integer
Dim r_int_nroaux     As Integer
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT TRIM(TO_CHAR(C.CAJMOV_SUCMOV,'000')) ||'-' ||  SUBSTR(C.CAJMOV_FECMOV,3,2) || TRIM(TO_CHAR(C.CAJMOV_NUMMOV,'00000')) AS COMPRB "
   g_str_Parame = g_str_Parame & "        ,TO_DATE(C.CAJMOV_FECMOV, 'YYYY/MM/DD') AS FCHAMOV "
   g_str_Parame = g_str_Parame & "        ,TRIM(DATGEN_APEPAT) ||' ' || TRIM(DATGEN_APEMAT) ||' '|| TRIM(DATGEN_NOMBRE) AS NOMBRE  "
   g_str_Parame = g_str_Parame & "        ,TRIM(DNI.PARDES_DESCRI) ||' - ' || C.CAJMOV_NUMDOC AS DNI  "
   g_str_Parame = g_str_Parame & "        ,SUBSTRC(C.CAJMOV_NUMOPE,1,3)||'-'||SUBSTRC(C.CAJMOV_NUMOPE,4,2) ||'-'||SUBSTRC(C.CAJMOV_NUMOPE,6,5) AS NUMOPE  "
   g_str_Parame = g_str_Parame & "        ,TO_DATE(C.CAJMOV_FECDEP, 'YYYY/MM/DD') AS FCHAPAGO "
   g_str_Parame = g_str_Parame & "        ,DECODE(C.CAJMOV_CODBAN,0,'',TRIM(BCO.PARDES_DESCRI) ||' - CTA: '|| C.CAJMOV_NUMCTA) AS BANCO "
   g_str_Parame = g_str_Parame & "        ,DECODE(C.CAJMOV_MONPAG,1,'SOLES','DOLARES') AS MONEDA "
   g_str_Parame = g_str_Parame & "        ,HIPPAG_CAPITA"
   g_str_Parame = g_str_Parame & "        ,HIPPAG_INTERE"
   g_str_Parame = g_str_Parame & "        ,HIPPAG_DESORG"
   g_str_Parame = g_str_Parame & "        ,HIPPAG_VIVORG"
   g_str_Parame = g_str_Parame & "        ,HIPPAG_OTRORG"
   g_str_Parame = g_str_Parame & "        ,HIPPAG_CAPBBP"
   g_str_Parame = g_str_Parame & "        ,HIPPAG_INTBBP"
   g_str_Parame = g_str_Parame & "        ,HIPPAG_INTMOR"
   g_str_Parame = g_str_Parame & "        ,HIPPAG_INTCOM"
   g_str_Parame = g_str_Parame & "        ,HIPPAG_GASCOB"
   g_str_Parame = g_str_Parame & "        ,HIPPAG_OTRGAS "
   g_str_Parame = g_str_Parame & "        ,HIPPAG_PAGMPR"
   g_str_Parame = g_str_Parame & "        ,(SELECT CAJMOV_ITFIMP "
   g_str_Parame = g_str_Parame & "            FROM OPE_CAJMOV "
   g_str_Parame = g_str_Parame & "           WHERE C.CAJMOV_SUCMOV = CAJMOV_SUCMOV"
   g_str_Parame = g_str_Parame & "             AND C.CAJMOV_FECMOV = CAJMOV_FECMOV"
   g_str_Parame = g_str_Parame & "             AND C.CAJMOV_NUMOPE = CAJMOV_NUMOPE"
   g_str_Parame = g_str_Parame & "             AND C.CAJMOV_NUMMOV = CAJMOV_NUMMOV) AS ITFIMP"
   g_str_Parame = g_str_Parame & "        ,(SELECT CAJMOV_IMPTOT  "
   g_str_Parame = g_str_Parame & "            FROM OPE_CAJMOV "
   g_str_Parame = g_str_Parame & "           WHERE C.CAJMOV_SUCMOV = CAJMOV_SUCMOV"
   g_str_Parame = g_str_Parame & "             AND C.CAJMOV_FECMOV = CAJMOV_FECMOV"
   g_str_Parame = g_str_Parame & "             AND C.CAJMOV_NUMOPE = CAJMOV_NUMOPE"
   g_str_Parame = g_str_Parame & "             AND C.CAJMOV_NUMMOV = CAJMOV_NUMMOV) AS IMPTOT"
   g_str_Parame = g_str_Parame & "        ,CAJMOV_TIPMOV ||' - ' || MOV.PARDES_DESCRI AS TIPMOV  "
   g_str_Parame = g_str_Parame & "   FROM (SELECT  CAJMOV_FECMOV, CAJMOV_NUMOPE, CAJMOV_NUMMOV"
   g_str_Parame = g_str_Parame & "               ,CAJMOV_SUCMOV, CAJMOV_CODBAN, CAJMOV_TIPDOC"
   g_str_Parame = g_str_Parame & "               ,CAJMOV_TIPMOV, CAJMOV_NUMDOC, CAJMOV_MONPAG"
   g_str_Parame = g_str_Parame & "               ,CAJMOV_NUMCTA,CAJMOV_FECDEP"
   g_str_Parame = g_str_Parame & "               , SUM(HIPPAG_CAPITA) AS HIPPAG_CAPITA"
   g_str_Parame = g_str_Parame & "               , SUM(HIPPAG_INTERE) AS HIPPAG_INTERE"
   g_str_Parame = g_str_Parame & "               , SUM(HIPPAG_DESORG) AS HIPPAG_DESORG"
   g_str_Parame = g_str_Parame & "               , SUM(HIPPAG_VIVORG) AS HIPPAG_VIVORG"
   g_str_Parame = g_str_Parame & "               , SUM(HIPPAG_OTRORG) AS HIPPAG_OTRORG"
   g_str_Parame = g_str_Parame & "               , SUM(HIPPAG_CAPBBP) AS HIPPAG_CAPBBP"
   g_str_Parame = g_str_Parame & "               , SUM(HIPPAG_INTBBP) AS HIPPAG_INTBBP"
   g_str_Parame = g_str_Parame & "               , SUM(HIPPAG_INTMOR) AS HIPPAG_INTMOR"
   g_str_Parame = g_str_Parame & "               , SUM(HIPPAG_INTCOM) AS HIPPAG_INTCOM"
   g_str_Parame = g_str_Parame & "               , SUM(HIPPAG_GASCOB) AS HIPPAG_GASCOB"
   g_str_Parame = g_str_Parame & "               , SUM(HIPPAG_OTRGAS) AS HIPPAG_OTRGAS "
   g_str_Parame = g_str_Parame & "               , SUM(HIPPAG_PAGMPR) AS HIPPAG_PAGMPR "
   g_str_Parame = g_str_Parame & "          FROM OPE_CAJMOV C"
   g_str_Parame = g_str_Parame & "         INNER JOIN CRE_HIPPAG ON (CAJMOV_SUCMOV=HIPPAG_SUCMOV AND CAJMOV_NUMMOV=HIPPAG_NUMMOV "
   g_str_Parame = g_str_Parame & "                           AND CAJMOV_FECMOV=HIPPAG_FECMOV AND CAJMOV_NUMOPE=HIPPAG_NUMOPE)"
   g_str_Parame = g_str_Parame & "         WHERE CAJMOV_SUCMOV = '001'  "
   g_str_Parame = g_str_Parame & "           AND CAJMOV_FECMOV >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " "
   g_str_Parame = g_str_Parame & "           AND CAJMOV_FECMOV <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & " "
   g_str_Parame = g_str_Parame & "           AND CAJMOV_TIPMOV = '1102'"
   g_str_Parame = g_str_Parame & "         GROUP BY  CAJMOV_FECMOV, CAJMOV_NUMOPE, CAJMOV_NUMMOV, CAJMOV_SUCMOV "
   g_str_Parame = g_str_Parame & "                  ,CAJMOV_CODBAN, CAJMOV_TIPDOC, CAJMOV_TIPMOV, CAJMOV_NUMDOC "
   g_str_Parame = g_str_Parame & "                  ,CAJMOV_MONPAG, CAJMOV_NUMCTA, CAJMOV_FECDEP "
   g_str_Parame = g_str_Parame & "        ) C"
   g_str_Parame = g_str_Parame & "  INNER JOIN MNT_PARDES BCO ON( BCO.PARDES_CODGRP = 516 AND BCO.PARDES_CODITE = C.CAJMOV_CODBAN) "
   g_str_Parame = g_str_Parame & "  INNER JOIN MNT_PARDES DNI ON( DNI.PARDES_CODGRP = 203 AND DNI.PARDES_CODITE = C.CAJMOV_TIPDOC) "
   g_str_Parame = g_str_Parame & "  INNER JOIN MNT_PARDES MOV ON( MOV.PARDES_CODGRP = 301 AND MOV.PARDES_CODITE = C.CAJMOV_TIPMOV) "
   g_str_Parame = g_str_Parame & "  INNER JOIN CLI_DATGEN ON (DATGEN_TIPDOC = C.CAJMOV_TIPDOC AND DATGEN_NUMDOC = C.CAJMOV_NUMDOC) "
   g_str_Parame = g_str_Parame & "  WHERE C.CAJMOV_FECMOV >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " "
   g_str_Parame = g_str_Parame & "    AND C.CAJMOV_FECMOV <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & " "
   g_str_Parame = g_str_Parame & "  ORDER BY C.CAJMOV_NUMMOV ASC, C.CAJMOV_FECMOV ASC  "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If

   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      l_int_MsjErr = l_int_MsjErr + 1
      Exit Sub
   End If
   
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   r_int_NroFil = 2
   
   With r_obj_Excel.ActiveSheet
   
   'IMAGEN
      On Local Error Resume Next
      
      'Unir celdas
      .Range("A" & r_int_NroFil & ":V" & r_int_NroFil & "").Merge
      .Range("A" & r_int_NroFil & "") = "REPORTE DE PAGO DE CUOTA - CRÉDITO HIPOTECARIO"
      .Range("A" & r_int_NroFil & "").Font.Underline = True
      .Range("A" & r_int_NroFil & "").Font.Bold = True
      .Range("A" & r_int_NroFil & "").HorizontalAlignment = xlHAlignCenter
         
      r_int_NroFil = r_int_NroFil + 1
      .Range("V" & r_int_NroFil & "") = "FECHA: " & Format(date, "DD/MM/YYYY")
      .Range("V" & r_int_NroFil & "").HorizontalAlignment = xlHAlignRight
      
      r_int_NroFil = r_int_NroFil + 1
      
      .Range("V" & r_int_NroFil & "") = "HORA: " & Format(Now, "HH:MM:SS")
      .Range("V" & r_int_NroFil & "").HorizontalAlignment = xlHAlignRight
      
      r_int_NroFil = r_int_NroFil + 2
   
      .Cells(r_int_NroFil, 1) = "No COMPROB.":               .Columns("A").ColumnWidth = 13
      .Cells(r_int_NroFil, 2) = "F. MOV.":                   .Columns("B").ColumnWidth = 10
      .Cells(r_int_NroFil, 3) = "CLIENTE":                   .Columns("C").ColumnWidth = 40
      .Cells(r_int_NroFil, 4) = "DOC. IDENTIDAD":            .Columns("D").ColumnWidth = 15
      .Cells(r_int_NroFil, 5) = "No OPERACIÓN":              .Columns("E").ColumnWidth = 14
      .Cells(r_int_NroFil, 6) = "FECHA PAGO":                .Columns("F").ColumnWidth = 11
      .Cells(r_int_NroFil, 7) = "BANCO":                     .Columns("G").ColumnWidth = 50
      .Cells(r_int_NroFil, 8) = "MONEDA":                    .Columns("H").ColumnWidth = 9
      .Cells(r_int_NroFil, 9) = "CAPITAL":                   .Columns("I").ColumnWidth = 9
      .Cells(r_int_NroFil, 10) = "INTERÉS":                  .Columns("J").ColumnWidth = 9
      .Cells(r_int_NroFil, 11) = "SEG.DESG.":                .Columns("K").ColumnWidth = 9
      .Cells(r_int_NroFil, 12) = "SEG.INM.":                 .Columns("L").ColumnWidth = 9
      .Cells(r_int_NroFil, 13) = "PORTES":                   .Columns("M").ColumnWidth = 8
      .Cells(r_int_NroFil, 14) = "CAP.PBP":                  .Columns("N").ColumnWidth = 10
      .Cells(r_int_NroFil, 15) = "INT.PBP":                  .Columns("O").ColumnWidth = 10
      .Cells(r_int_NroFil, 16) = "INT.MOR.":                 .Columns("P").ColumnWidth = 10
      .Cells(r_int_NroFil, 17) = "INT.COM.":                 .Columns("Q").ColumnWidth = 10
      .Cells(r_int_NroFil, 18) = "GTOS.COB.":                .Columns("R").ColumnWidth = 10
      .Cells(r_int_NroFil, 19) = "OTR.GTOS.":                .Columns("S").ColumnWidth = 10
      .Cells(r_int_NroFil, 20) = "SUB TOTAL":                .Columns("T").ColumnWidth = 11
      .Cells(r_int_NroFil, 21) = "ITF":                      .Columns("U").ColumnWidth = 0
      .Cells(r_int_NroFil, 22) = "TOTAL":                    .Columns("V").ColumnWidth = 11
            
      .Columns("A").HorizontalAlignment = xlHAlignCenter
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("H").HorizontalAlignment = xlHAlignCenter
            
      .Range("A" & r_int_NroFil & ":V" & r_int_NroFil & "").HorizontalAlignment = xlHAlignCenter
      .Range("A" & r_int_NroFil & ":V" & r_int_NroFil & "").Font.Bold = True
      r_int_NroFil = r_int_NroFil + 1
      
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         g_rst_Princi.MoveFirst
         
         Do While Not g_rst_Princi.EOF
            .Cells(r_int_NroFil, 1) = g_rst_Princi!COMPRB
            .Cells(r_int_NroFil, 2) = "'" & Trim(g_rst_Princi!FCHAMOV)
            .Cells(r_int_NroFil, 3) = Trim(g_rst_Princi!NOMBRE)
            .Cells(r_int_NroFil, 4) = g_rst_Princi!DNI
            .Cells(r_int_NroFil, 5) = g_rst_Princi!NUMOPE
            .Cells(r_int_NroFil, 6) = "'" & Trim(g_rst_Princi!FCHAPAGO)
            .Cells(r_int_NroFil, 7) = Trim(g_rst_Princi!BANCO)
            .Cells(r_int_NroFil, 8) = Trim(g_rst_Princi!MONEDA)
            .Cells(r_int_NroFil, 9) = g_rst_Princi!HIPPAG_CAPITA
            .Cells(r_int_NroFil, 10) = g_rst_Princi!HIPPAG_INTERE
            .Cells(r_int_NroFil, 11) = g_rst_Princi!HIPPAG_DESORG
            .Cells(r_int_NroFil, 12) = g_rst_Princi!HIPPAG_VIVORG
            .Cells(r_int_NroFil, 13) = g_rst_Princi!HIPPAG_OTRORG
            .Cells(r_int_NroFil, 14) = g_rst_Princi!HIPPAG_CAPBBP
            .Cells(r_int_NroFil, 15) = g_rst_Princi!HIPPAG_INTBBP
            .Cells(r_int_NroFil, 16) = g_rst_Princi!HIPPAG_INTMOR
            .Cells(r_int_NroFil, 17) = g_rst_Princi!HIPPAG_INTCOM
            .Cells(r_int_NroFil, 18) = g_rst_Princi!HIPPAG_GASCOB
            .Cells(r_int_NroFil, 19) = g_rst_Princi!HIPPAG_OTRGAS
            .Cells(r_int_NroFil, 20) = g_rst_Princi!HIPPAG_PAGMPR
            .Cells(r_int_NroFil, 21) = g_rst_Princi!ITFIMP
            .Cells(r_int_NroFil, 22) = g_rst_Princi!IMPTOT
            r_int_NroFil = r_int_NroFil + 1
         
            g_rst_Princi.MoveNext
         Loop
      End If
       
      .Range(.Cells(2, 9), .Cells(r_int_NroFil, 22)).Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.00"
      .Cells(1, 1).Select
   End With
     
   g_rst_Princi.Close
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

'*****************************************
'1101 - PAGO GASTOS DE CIERRE CRED. HIPOT.
Private Sub fs_GenExc_GtoCie()
Dim r_obj_Excel      As Excel.Application
Dim r_int_NroFil     As Integer
Dim r_int_nroaux     As Integer
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT  TRIM(TO_CHAR(CAJMOV_SUCMOV,'000')) ||'-' ||  SUBSTR(CAJMOV_FECMOV,3,2) || TRIM(TO_CHAR(CAJMOV_NUMMOV,'00000')) AS COMPRB "
   g_str_Parame = g_str_Parame & "       ,TO_DATE(CAJMOV_FECMOV, 'YYYY/MM/DD') AS FCHAMOV "
   g_str_Parame = g_str_Parame & "       ,TRIM(DATGEN_APEPAT) ||' ' || TRIM(DATGEN_APEMAT) ||' '|| TRIM(DATGEN_NOMBRE) AS NOMBRE  "
   g_str_Parame = g_str_Parame & "       ,TRIM(DNI.PARDES_DESCRI) ||' - ' || CAJMOV_NUMDOC AS DNI  "
   g_str_Parame = g_str_Parame & "       ,SUBSTRC(CAJMOV_NUMOPE,1,3)||'-'||SUBSTRC(CAJMOV_NUMOPE,4,3) ||'-'||SUBSTRC(CAJMOV_NUMOPE,7,2) ||'-'||SUBSTRC(CAJMOV_NUMOPE,9,4)AS NUMOPE "
   g_str_Parame = g_str_Parame & "       ,TO_DATE(CAJMOV_FECDEP, 'YYYY/MM/DD') AS FCHAPAGO "
   g_str_Parame = g_str_Parame & "       ,DECODE(CAJMOV_CODBAN,0,'',TRIM(BCO.PARDES_DESCRI) ||' - CTA: '|| CAJMOV_NUMCTA) AS BANCO "
   g_str_Parame = g_str_Parame & "       ,DECODE(CAJMOV_MONPAG,1,'SOLES','DOLARES') AS MONEDA "
   g_str_Parame = g_str_Parame & "       ,CAJMOV_IMPPAG "
   g_str_Parame = g_str_Parame & "       ,CAJMOV_ITFIMP "
   g_str_Parame = g_str_Parame & "       ,CAJMOV_IMPTOT  "
   g_str_Parame = g_str_Parame & "       ,CAJMOV_TIPMOV ||' - ' || MOV.PARDES_DESCRI AS TIPMOV "
   g_str_Parame = g_str_Parame & "       ,CAJMOV_NUMOPE "
   g_str_Parame = g_str_Parame & "  FROM OPE_CAJMOV  "
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES BCO ON( BCO.PARDES_CODGRP = 516 AND BCO.PARDES_CODITE = CAJMOV_CODBAN) "
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES DNI ON( DNI.PARDES_CODGRP = 203 AND DNI.PARDES_CODITE = CAJMOV_TIPDOC) "
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES MOV ON( MOV.PARDES_CODGRP = 301 AND MOV.PARDES_CODITE = CAJMOV_TIPMOV) "
   g_str_Parame = g_str_Parame & " INNER JOIN CLI_DATGEN ON (DATGEN_TIPDOC = CAJMOV_TIPDOC AND DATGEN_NUMDOC = CAJMOV_NUMDOC) "
   g_str_Parame = g_str_Parame & " WHERE CAJMOV_SUCMOV = '001' "
   g_str_Parame = g_str_Parame & "   AND CAJMOV_FECMOV >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " "
   g_str_Parame = g_str_Parame & "   AND CAJMOV_FECMOV <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & " "
   g_str_Parame = g_str_Parame & "   AND CAJMOV_TIPMOV = 1101 "
   g_str_Parame = g_str_Parame & " ORDER BY CAJMOV_NUMMOV ASC, CAJMOV_FECMOV ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If

   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      l_int_MsjErr = l_int_MsjErr + 1
      Exit Sub
   End If
   
   '***
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT SOLMAE_NUMERO, GASADM_IMPORT, PARPRD_DESCRI,GASADM_CODGAS "
   g_str_Parame = g_str_Parame & "  FROM CRE_SOLMAE "
   g_str_Parame = g_str_Parame & " INNER JOIN OPE_CAJMOV  ON (SOLMAE_NUMERO = CAJMOV_NUMOPE AND SOLMAE_TITTDO =CAJMOV_TIPDOC AND SOLMAE_TITNDO = CAJMOV_NUMDOC) "
   g_str_Parame = g_str_Parame & " INNER JOIN TRA_GASADM  ON (SOLMAE_NUMERO = GASADM_NUMSOL) "
   g_str_Parame = g_str_Parame & " INNER JOIN CRE_PARPRD  ON (PARPRD_CODPRD = SOLMAE_CODPRD AND PARPRD_CODSUB = SOLMAE_CODSUB  AND "
   g_str_Parame = g_str_Parame & "                            PARPRD_CODGRP = '007' AND PARPRD_CODITE = GASADM_CODGAS||GASADM_TIPMON) "
   g_str_Parame = g_str_Parame & " WHERE CAJMOV_SUCMOV = '001' "
   g_str_Parame = g_str_Parame & "   AND CAJMOV_FECMOV >=" & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " "
   g_str_Parame = g_str_Parame & "   AND CAJMOV_FECMOV <=" & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & " "
   g_str_Parame = g_str_Parame & "   AND CAJMOV_TIPMOV = 1101 "
   g_str_Parame = g_str_Parame & " ORDER BY CAJMOV_NUMMOV ASC, CAJMOV_FECMOV ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
       Exit Sub
   End If

   If g_rst_GenAux.BOF And g_rst_GenAux.EOF Then
      g_rst_GenAux.Close
      Set g_rst_GenAux = Nothing
      l_int_MsjErr = l_int_MsjErr + 1
      Exit Sub
   End If
   
   '*********************************
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   r_int_NroFil = 2
   
   With r_obj_Excel.ActiveSheet
      'IMAGEN
      On Local Error Resume Next
      
      'Unir celdas
      .Range("A" & r_int_NroFil & ":X" & r_int_NroFil & "").Merge
      .Range("A" & r_int_NroFil & "") = "REPORTE DE PAGO DE GASTOS DE CIERRE - CRÉDITO HIPOTECARIO"
      .Range("A" & r_int_NroFil & "").Font.Underline = True
      .Range("A" & r_int_NroFil & "").Font.Bold = True
      .Range("A" & r_int_NroFil & "").HorizontalAlignment = xlHAlignCenter
         
      r_int_NroFil = r_int_NroFil + 1
      .Range("X" & r_int_NroFil & "") = "FECHA: " & Format(date, "DD/MM/YYYY")
      .Range("X" & r_int_NroFil & "").HorizontalAlignment = xlHAlignRight
      
      r_int_NroFil = r_int_NroFil + 1
      
      .Range("X" & r_int_NroFil & "") = "HORA: " & Format(Now, "HH:MM:SS")
      .Range("X" & r_int_NroFil & "").HorizontalAlignment = xlHAlignRight
      
      r_int_NroFil = r_int_NroFil + 2
   
      .Cells(r_int_NroFil, 1) = "No COMPR.":                 .Columns("A").ColumnWidth = 12
      .Cells(r_int_NroFil, 2) = "F. MOV.":                   .Columns("B").ColumnWidth = 10
      .Cells(r_int_NroFil, 3) = "CLIENTE":                   .Columns("C").ColumnWidth = 40
      .Cells(r_int_NroFil, 4) = "DOC. IDENT.":               .Columns("D").ColumnWidth = 13
      .Cells(r_int_NroFil, 5) = "No SOLICITUD":              .Columns("E").ColumnWidth = 15
      .Cells(r_int_NroFil, 6) = "F. PAGO":                   .Columns("F").ColumnWidth = 10
      .Cells(r_int_NroFil, 7) = "BANCO":                     .Columns("G").ColumnWidth = 50
      .Cells(r_int_NroFil, 8) = "MONEDA":                    .Columns("H").ColumnWidth = 9
      .Cells(r_int_NroFil, 9) = "GTO.TAS.":                  .Columns("I").ColumnWidth = 9
      .Cells(r_int_NroFil, 10) = "GTO.NOT.":                 .Columns("J").ColumnWidth = 9
      .Cells(r_int_NroFil, 11) = "ITF":                      .Columns("K").ColumnWidth = 7
      .Cells(r_int_NroFil, 12) = "COM.EST.TIT.":             .Columns("L").ColumnWidth = 12
      .Cells(r_int_NroFil, 13) = "COM.EVA.CRE.":             .Columns("M").ColumnWidth = 13
      .Cells(r_int_NroFil, 14) = "GTO./BLQ.REG.":            .Columns("N").ColumnWidth = 13
      .Cells(r_int_NroFil, 15) = "MIN. CPRA/VTA":            .Columns("O").ColumnWidth = 14
      .Cells(r_int_NroFil, 16) = "INS.GAR.":                 .Columns("P").ColumnWidth = 10
      .Cells(r_int_NroFil, 17) = "GTO.DES.COF.":             .Columns("Q").ColumnWidth = 13
      .Cells(r_int_NroFil, 18) = "ADM/CTRL TAS.":            .Columns("R").ColumnWidth = 14
      .Cells(r_int_NroFil, 19) = "COM.RED.CONT.":            .Columns("S").ColumnWidth = 15
      .Cells(r_int_NroFil, 20) = "COM.GES.CRED.":            .Columns("T").ColumnWidth = 15
      .Cells(r_int_NroFil, 21) = "CTRL. GAR.":               .Columns("U").ColumnWidth = 10
      .Cells(r_int_NroFil, 22) = "SUB TOTAL":                .Columns("V").ColumnWidth = 11
      .Cells(r_int_NroFil, 23) = "ITF":                      .Columns("W").ColumnWidth = 6
      .Cells(r_int_NroFil, 24) = "TOTAL":                    .Columns("X").ColumnWidth = 10
            
      .Columns("A").HorizontalAlignment = xlHAlignCenter
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("H").HorizontalAlignment = xlHAlignCenter
             
      .Range("A" & r_int_NroFil & ":X" & r_int_NroFil & "").HorizontalAlignment = xlHAlignCenter
      .Range("A" & r_int_NroFil & ":X" & r_int_NroFil & "").Font.Bold = True
      
      r_int_NroFil = r_int_NroFil + 1
      
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         g_rst_Princi.MoveFirst
         
         Do While Not g_rst_Princi.EOF
            .Cells(r_int_NroFil, 1) = g_rst_Princi!COMPRB
            .Cells(r_int_NroFil, 2) = "'" & Trim(g_rst_Princi!FCHAMOV)
            .Cells(r_int_NroFil, 3) = Trim(g_rst_Princi!NOMBRE)
            .Cells(r_int_NroFil, 4) = Trim(g_rst_Princi!DNI)
            .Cells(r_int_NroFil, 5) = g_rst_Princi!NUMOPE
            .Cells(r_int_NroFil, 6) = "'" & Trim(g_rst_Princi!FCHAPAGO)
            .Cells(r_int_NroFil, 7) = Trim(g_rst_Princi!BANCO)
            .Cells(r_int_NroFil, 8) = Trim(g_rst_Princi!MONEDA)
            .Cells(r_int_NroFil, 9) = 0
            .Cells(r_int_NroFil, 10) = 0
            .Cells(r_int_NroFil, 11) = 0
            .Cells(r_int_NroFil, 12) = 0
            .Cells(r_int_NroFil, 13) = 0
            .Cells(r_int_NroFil, 14) = 0
            .Cells(r_int_NroFil, 15) = 0
            .Cells(r_int_NroFil, 16) = 0
            .Cells(r_int_NroFil, 17) = 0
            .Cells(r_int_NroFil, 18) = 0
            .Cells(r_int_NroFil, 19) = 0
            .Cells(r_int_NroFil, 20) = 0
            .Cells(r_int_NroFil, 21) = 0
            .Cells(r_int_NroFil, 22) = g_rst_Princi!CAJMOV_IMPPAG
            .Cells(r_int_NroFil, 23) = g_rst_Princi!CAJMOV_ITFIMP
            .Cells(r_int_NroFil, 24) = g_rst_Princi!CAJMOV_IMPTOT
                                                                        
            '***********************************
            If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
               g_rst_GenAux.MoveFirst
               Do While Not g_rst_GenAux.EOF
               
                  If CDbl(g_rst_Princi!CAJMOV_NUMOPE) = CDbl(g_rst_GenAux!SOLMAE_NUMERO) Then
                     Select Case g_rst_GenAux!GASADM_CODGAS
                        Case 11
                           .Cells(r_int_NroFil, 9) = g_rst_GenAux!GASADM_IMPORT
                        Case 12
                           .Cells(r_int_NroFil, 10) = g_rst_GenAux!GASADM_IMPORT
                        Case 13
                           .Cells(r_int_NroFil, 11) = g_rst_GenAux!GASADM_IMPORT
                        Case 14
                           .Cells(r_int_NroFil, 12) = g_rst_GenAux!GASADM_IMPORT
                        Case 15
                           .Cells(r_int_NroFil, 13) = g_rst_GenAux!GASADM_IMPORT
                        Case 16
                           .Cells(r_int_NroFil, 14) = g_rst_GenAux!GASADM_IMPORT
                        Case 17
                           .Cells(r_int_NroFil, 15) = g_rst_GenAux!GASADM_IMPORT
                        Case 18
                           .Cells(r_int_NroFil, 16) = g_rst_GenAux!GASADM_IMPORT
                        Case 19
                           .Cells(r_int_NroFil, 17) = g_rst_GenAux!GASADM_IMPORT
                        Case 20
                           .Cells(r_int_NroFil, 18) = g_rst_GenAux!GASADM_IMPORT
                        Case 21
                           .Cells(r_int_NroFil, 19) = g_rst_GenAux!GASADM_IMPORT
                        Case 22
                           .Cells(r_int_NroFil, 20) = g_rst_GenAux!GASADM_IMPORT
                        Case 23
                           .Cells(r_int_NroFil, 21) = g_rst_GenAux!GASADM_IMPORT
                     End Select
                  End If
                  
                  g_rst_GenAux.MoveNext
               Loop
               
            End If
            '***************
            
            r_int_NroFil = r_int_NroFil + 1
            g_rst_Princi.MoveNext
         Loop
      End If
       
      .Range(.Cells(2, 9), .Cells(r_int_NroFil, 21)).Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.00"
      .Cells(1, 1).Select
   End With
     
   g_rst_Princi.Close
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

'***********************************
'1103 - DESEMBOLSO CRED. HIPOT.
Private Sub fs_GenExc_Desemb()
Dim r_obj_Excel      As Excel.Application
Dim r_int_NroFil     As Integer
Dim r_int_nroaux     As Integer
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT TRIM(TO_CHAR(CAJMOV_SUCMOV,'000')) ||'-' ||  SUBSTR(CAJMOV_FECMOV,3,2) || TRIM(TO_CHAR(CAJMOV_NUMMOV,'00000')) AS COMPRB "
   g_str_Parame = g_str_Parame & "       ,TO_DATE(CAJMOV_FECMOV, 'YYYY/MM/DD') AS FCHAMOV "
   g_str_Parame = g_str_Parame & "       ,TRIM(DATGEN_APEPAT) ||' ' || TRIM(DATGEN_APEMAT) ||' '|| TRIM(DATGEN_NOMBRE) AS NOMBRE  "
   g_str_Parame = g_str_Parame & "       ,TRIM(DNI.PARDES_DESCRI) ||' - ' || CAJMOV_NUMDOC AS DNI  "
   g_str_Parame = g_str_Parame & "       ,SUBSTRC(CAJMOV_NUMOPE,1,3)||'-'||SUBSTRC(CAJMOV_NUMOPE,4,2) ||'-'||SUBSTRC(CAJMOV_NUMOPE,6,5) AS NUMOPE "
   g_str_Parame = g_str_Parame & "       ,TO_DATE(CAJMOV_FECDEP, 'YYYY/MM/DD') AS FCHAPAGO "
   g_str_Parame = g_str_Parame & "       ,DECODE(CAJMOV_CODBAN,0,'',TRIM(BCO.PARDES_DESCRI) ||' - NRO DE CUENTA: '|| CAJMOV_NUMCTA) AS BANCO "
   g_str_Parame = g_str_Parame & "       ,DECODE(CAJMOV_MONPAG,1,'SOLES','DOLARES') AS MONEDA "
   g_str_Parame = g_str_Parame & "       ,CAJMOV_IMPPAG "
   g_str_Parame = g_str_Parame & "       ,CAJMOV_ITFIMP "
   g_str_Parame = g_str_Parame & "       ,CAJMOV_IMPTOT  "
   g_str_Parame = g_str_Parame & "       ,CAJMOV_TIPMOV ||' - ' || MOV.PARDES_DESCRI AS TIPMOV "
   g_str_Parame = g_str_Parame & "  FROM OPE_CAJMOV  "
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES BCO ON( BCO.PARDES_CODGRP = 516 AND BCO.PARDES_CODITE = CAJMOV_CODBAN) "
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES DNI ON( DNI.PARDES_CODGRP = 203 AND DNI.PARDES_CODITE = CAJMOV_TIPDOC) "
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES MOV ON( MOV.PARDES_CODGRP = 301 AND MOV.PARDES_CODITE = CAJMOV_TIPMOV) "
   g_str_Parame = g_str_Parame & " INNER JOIN CLI_DATGEN ON (DATGEN_TIPDOC = CAJMOV_TIPDOC AND DATGEN_NUMDOC = CAJMOV_NUMDOC) "
   g_str_Parame = g_str_Parame & " WHERE CAJMOV_SUCMOV = '001' "
   g_str_Parame = g_str_Parame & "   AND CAJMOV_FECMOV >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " "
   g_str_Parame = g_str_Parame & "   AND CAJMOV_FECMOV <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & " "
   g_str_Parame = g_str_Parame & "   AND CAJMOV_TIPMOV = 1103 "
   g_str_Parame = g_str_Parame & " ORDER BY CAJMOV_NUMMOV ASC, CAJMOV_FECMOV ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If

   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      l_int_MsjErr = l_int_MsjErr + 1
      Exit Sub
   End If
   
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   r_int_NroFil = 2
   
   With r_obj_Excel.ActiveSheet
      On Local Error Resume Next
      
      'Unir celdas
      .Range("A" & r_int_NroFil & ":K" & r_int_NroFil & "").Merge
      .Range("A" & r_int_NroFil & "") = "REPORTE DE DESEMBOLSO - CRÉDITO HIPOTECARIO"
      .Range("A" & r_int_NroFil & "").Font.Underline = True
      .Range("A" & r_int_NroFil & "").Font.Bold = True
      .Range("A" & r_int_NroFil & "").HorizontalAlignment = xlHAlignCenter
         
      r_int_NroFil = r_int_NroFil + 1
      .Range("K" & r_int_NroFil & "") = "FECHA: " & Format(date, "DD/MM/YYYY")
      .Range("K" & r_int_NroFil & "").HorizontalAlignment = xlHAlignRight
      
      r_int_NroFil = r_int_NroFil + 1
      
      .Range("K" & r_int_NroFil & "") = "HORA: " & Format(Now, "HH:MM:SS")
      .Range("K" & r_int_NroFil & "").HorizontalAlignment = xlHAlignRight
      
      r_int_NroFil = r_int_NroFil + 2
   
      .Cells(r_int_NroFil, 1) = "No COMPROBANTE":          .Columns("A").ColumnWidth = 18
      .Cells(r_int_NroFil, 2) = "F. MOVIMIENTO":             .Columns("B").ColumnWidth = 14.8
      .Cells(r_int_NroFil, 3) = "CLIENTE":                   .Columns("C").ColumnWidth = 45
      .Cells(r_int_NroFil, 4) = "DOC. IDENTIDAD":            .Columns("D").ColumnWidth = 17
      .Cells(r_int_NroFil, 5) = "NRO OPERACIÓN":             .Columns("E").ColumnWidth = 17
      .Cells(r_int_NroFil, 6) = "FECHA PAGO":                .Columns("F").ColumnWidth = 13
      .Cells(r_int_NroFil, 7) = "MONEDA":                    .Columns("G").ColumnWidth = 10
      .Cells(r_int_NroFil, 8) = "DESEMBOLSO":                .Columns("H").ColumnWidth = 13
      .Cells(r_int_NroFil, 9) = "SUB TOTAL":                 .Columns("I").ColumnWidth = 13
      .Cells(r_int_NroFil, 10) = "ITF":                      .Columns("J").ColumnWidth = 10
      .Cells(r_int_NroFil, 11) = "TOTAL":                    .Columns("K").ColumnWidth = 13
            
      .Columns("A").HorizontalAlignment = xlHAlignCenter
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("G").HorizontalAlignment = xlHAlignCenter
                  
      .Range("A" & r_int_NroFil & ":K" & r_int_NroFil & "").HorizontalAlignment = xlHAlignCenter
      .Range("A" & r_int_NroFil & ":K" & r_int_NroFil & "").Font.Bold = True
      r_int_NroFil = r_int_NroFil + 1
      
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         g_rst_Princi.MoveFirst
         Do While Not g_rst_Princi.EOF
            .Cells(r_int_NroFil, 1) = g_rst_Princi!COMPRB
            .Cells(r_int_NroFil, 2) = "'" & Trim(g_rst_Princi!FCHAMOV)
            .Cells(r_int_NroFil, 3) = Trim(g_rst_Princi!NOMBRE)
            .Cells(r_int_NroFil, 4) = g_rst_Princi!DNI
            .Cells(r_int_NroFil, 5) = g_rst_Princi!NUMOPE
            .Cells(r_int_NroFil, 6) = "'" & Trim(g_rst_Princi!FCHAPAGO)
            .Cells(r_int_NroFil, 7) = Trim(g_rst_Princi!MONEDA)
            .Cells(r_int_NroFil, 8) = g_rst_Princi!CAJMOV_IMPPAG
            .Cells(r_int_NroFil, 9) = g_rst_Princi!CAJMOV_IMPPAG
            .Cells(r_int_NroFil, 10) = g_rst_Princi!CAJMOV_ITFIMP
            .Cells(r_int_NroFil, 11) = g_rst_Princi!CAJMOV_IMPTOT
            r_int_NroFil = r_int_NroFil + 1
         
            g_rst_Princi.MoveNext
         Loop
      End If
       
      .Range(.Cells(2, 8), .Cells(r_int_NroFil, 12)).Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.00"
      .Cells(1, 1).Select
   End With
     
   g_rst_Princi.Close
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

'***********************************
'2101 - EXTORNO PAGO GASTOS DE CIERRE
Private Sub fs_GenExc_Extrno()
Dim r_obj_Excel      As Excel.Application
Dim r_int_NroFil     As Integer
Dim r_int_nroaux     As Integer
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT  TRIM(TO_CHAR(CAJMOV_SUCMOV,'000')) ||'-' ||  SUBSTR(CAJMOV_FECMOV,3,2) || TRIM(TO_CHAR(CAJMOV_NUMMOV,'00000')) AS COMPRB "
   g_str_Parame = g_str_Parame & "       ,TO_DATE(CAJMOV_FECMOV, 'YYYY/MM/DD') AS FCHAMOV "
   g_str_Parame = g_str_Parame & "       ,TRIM(DATGEN_APEPAT) ||' ' || TRIM(DATGEN_APEMAT) ||' '|| TRIM(DATGEN_NOMBRE) AS NOMBRE  "
   g_str_Parame = g_str_Parame & "       ,TRIM(DNI.PARDES_DESCRI) ||' - ' || CAJMOV_NUMDOC AS DNI  "
   g_str_Parame = g_str_Parame & "       ,SUBSTRC(CAJMOV_NUMOPE,1,3)||'-'||SUBSTRC(CAJMOV_NUMOPE,4,3) ||'-'||SUBSTRC(CAJMOV_NUMOPE,7,2) ||'-'||SUBSTRC(CAJMOV_NUMOPE,9,4)AS NUMOPE"
   g_str_Parame = g_str_Parame & "       ,TO_DATE(CAJMOV_FECDEP, 'YYYY/MM/DD') AS FCHAPAGO "
   g_str_Parame = g_str_Parame & "       ,DECODE(CAJMOV_CODBAN,0,'',TRIM(BCO.PARDES_DESCRI) ||' - NRO DE CUENTA: '|| CAJMOV_NUMCTA) AS BANCO "
   g_str_Parame = g_str_Parame & "       ,DECODE(CAJMOV_MONPAG,1,'SOLES','DOLARES') AS MONEDA "
   g_str_Parame = g_str_Parame & "       ,CAJMOV_IMPPAG "
   g_str_Parame = g_str_Parame & "       ,CAJMOV_ITFIMP "
   g_str_Parame = g_str_Parame & "       ,CAJMOV_IMPTOT  "
   g_str_Parame = g_str_Parame & "  FROM OPE_CAJMOV  "
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES BCO ON( BCO.PARDES_CODGRP = 516 AND BCO.PARDES_CODITE = CAJMOV_CODBAN) "
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES DNI ON( DNI.PARDES_CODGRP = 203 AND DNI.PARDES_CODITE = CAJMOV_TIPDOC) "
   g_str_Parame = g_str_Parame & " INNER JOIN CLI_DATGEN ON (DATGEN_TIPDOC = CAJMOV_TIPDOC AND DATGEN_NUMDOC = CAJMOV_NUMDOC) "
   g_str_Parame = g_str_Parame & " WHERE CAJMOV_SUCMOV = '001'  "
   g_str_Parame = g_str_Parame & "   AND CAJMOV_FECMOV >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " "
   g_str_Parame = g_str_Parame & "   AND CAJMOV_FECMOV <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & " "
   g_str_Parame = g_str_Parame & "   AND CAJMOV_TIPMOV = 2101 "
   g_str_Parame = g_str_Parame & " ORDER BY CAJMOV_FECMOV ASC, CAJMOV_NUMOPE ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      l_int_MsjErr = l_int_MsjErr + 1
      Exit Sub
   End If
   
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   r_int_NroFil = 2
   
   With r_obj_Excel.ActiveSheet
      On Local Error Resume Next
   
      'Unir celdas
      .Range("A" & r_int_NroFil & ":K" & r_int_NroFil & "").Merge
      .Range("A" & r_int_NroFil & "") = "REPORTE DE EXTORNO DE PAGO GASTOS DE CIERRE - CRÉDITO HIPOTECARIO"
      .Range("A" & r_int_NroFil & "").Font.Underline = True
      .Range("A" & r_int_NroFil & "").Font.Bold = True
      .Range("A" & r_int_NroFil & "").HorizontalAlignment = xlHAlignCenter
         
      r_int_NroFil = r_int_NroFil + 1
      .Range("K" & r_int_NroFil & "") = "FECHA: " & Format(date, "DD/MM/YYYY")
      .Range("K" & r_int_NroFil & "").HorizontalAlignment = xlHAlignRight
      
      r_int_NroFil = r_int_NroFil + 1
      
      .Range("K" & r_int_NroFil & "") = "HORA: " & Format(Now, "HH:MM:SS")
      .Range("K" & r_int_NroFil & "").HorizontalAlignment = xlHAlignRight
      
      r_int_NroFil = r_int_NroFil + 2
         
      .Cells(r_int_NroFil, 1) = "NRO. COMPROBANTE":          .Columns("A").ColumnWidth = 20
      .Cells(r_int_NroFil, 2) = "FECHA MOVIMIENTO":          .Columns("B").ColumnWidth = 19
      .Cells(r_int_NroFil, 3) = "CLIENTE":                   .Columns("C").ColumnWidth = 45
      .Cells(r_int_NroFil, 4) = "DOC. IDENTIDAD":            .Columns("D").ColumnWidth = 16
      .Cells(r_int_NroFil, 5) = "NRO OPERACIÓN":             .Columns("E").ColumnWidth = 16
      .Cells(r_int_NroFil, 6) = "FECHA PAGO":                .Columns("F").ColumnWidth = 13
      .Cells(r_int_NroFil, 7) = "MONEDA":                    .Columns("G").ColumnWidth = 11
      .Cells(r_int_NroFil, 8) = "EXTORNO":                   .Columns("H").ColumnWidth = 13
      .Cells(r_int_NroFil, 9) = "SUB TOTAL":                 .Columns("I").ColumnWidth = 13
      .Cells(r_int_NroFil, 10) = "ITF":                      .Columns("J").ColumnWidth = 10
      .Cells(r_int_NroFil, 11) = "TOTAL":                    .Columns("K").ColumnWidth = 13
            
      .Columns("A").HorizontalAlignment = xlHAlignCenter
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("G").HorizontalAlignment = xlHAlignCenter
            
      .Range("A" & r_int_NroFil & ":K" & r_int_NroFil & "").HorizontalAlignment = xlHAlignCenter
      .Range("A" & r_int_NroFil & ":K" & r_int_NroFil & "").Font.Bold = True
      
      r_int_NroFil = r_int_NroFil + 1
      
      If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
         g_rst_Princi.MoveFirst
         Do While Not g_rst_Princi.EOF
            .Cells(r_int_NroFil, 1) = g_rst_Princi!COMPRB
            .Cells(r_int_NroFil, 2) = "'" & Trim(g_rst_Princi!FCHAMOV)
            .Cells(r_int_NroFil, 3) = Trim(g_rst_Princi!NOMBRE)
            .Cells(r_int_NroFil, 4) = g_rst_Princi!DNI
            .Cells(r_int_NroFil, 5) = g_rst_Princi!NUMOPE
            .Cells(r_int_NroFil, 6) = "'" & Trim(g_rst_Princi!FCHAPAGO)
            .Cells(r_int_NroFil, 7) = Trim(g_rst_Princi!MONEDA)
            .Cells(r_int_NroFil, 8) = g_rst_Princi!CAJMOV_IMPPAG
            .Cells(r_int_NroFil, 9) = g_rst_Princi!CAJMOV_IMPPAG
            .Cells(r_int_NroFil, 10) = g_rst_Princi!CAJMOV_ITFIMP
            .Cells(r_int_NroFil, 11) = g_rst_Princi!CAJMOV_IMPTOT
            r_int_NroFil = r_int_NroFil + 1
         
            g_rst_Princi.MoveNext
         Loop
      End If
        
      .Columns("H").Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.00"
      .Columns("I").Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.00"
      .Columns("J").Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.00"
      .Columns("K").Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.00"
      .Cells(1, 1).Select
   End With
     
   g_rst_Princi.Close
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub
