VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_Con_PreCon_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form4"
   ClientHeight    =   9570
   ClientLeft      =   1335
   ClientTop       =   2715
   ClientWidth     =   13620
   Icon            =   "OpeTra_frm_329.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9570
   ScaleWidth      =   13620
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   9570
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   13635
      _Version        =   65536
      _ExtentX        =   24051
      _ExtentY        =   16880
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
         Left            =   60
         TabIndex        =   8
         Top             =   810
         Width           =   13515
         _Version        =   65536
         _ExtentX        =   23839
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
         Begin VB.CommandButton cmd_Buscar 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_329.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Buscar Registros"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   630
            Picture         =   "OpeTra_frm_329.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Limpiar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   12915
            Picture         =   "OpeTra_frm_329.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   1230
            Picture         =   "OpeTra_frm_329.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   840
         Left            =   60
         TabIndex        =   9
         Top             =   1500
         Width           =   13515
         _Version        =   65536
         _ExtentX        =   23839
         _ExtentY        =   1482
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
         Begin VB.ComboBox cmb_TipPre 
            Height          =   315
            Left            =   5370
            TabIndex        =   23
            Text            =   "RED. MONTO"
            Top             =   90
            Width           =   2280
         End
         Begin EditLib.fpDateTime ipp_FecIni 
            Height          =   315
            Left            =   1575
            TabIndex        =   0
            Top             =   90
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
            Left            =   1575
            TabIndex        =   1
            Top             =   420
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
         Begin VB.Label Label31 
            Caption         =   "Tipo de Prepago"
            Height          =   315
            Left            =   3930
            TabIndex        =   24
            Top             =   150
            Width           =   1500
         End
         Begin VB.Label Label10 
            Caption         =   "Fecha Inicio:"
            Height          =   315
            Left            =   135
            TabIndex        =   11
            Top             =   150
            Width           =   1365
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha Fin:"
            Height          =   315
            Left            =   135
            TabIndex        =   10
            Top             =   480
            Width           =   1365
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   705
         Left            =   60
         TabIndex        =   12
         Top             =   60
         Width           =   13515
         _Version        =   65536
         _ExtentX        =   23839
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
         Begin Threed.SSPanel SSPanel5 
            Height          =   315
            Left            =   660
            TabIndex        =   22
            Top             =   60
            Width           =   4995
            _Version        =   65536
            _ExtentX        =   8811
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Operaciones Financieras"
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
            Height          =   315
            Left            =   660
            TabIndex        =   13
            Top             =   360
            Width           =   4995
            _Version        =   65536
            _ExtentX        =   8811
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Consulta de Prepagos"
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
            Picture         =   "OpeTra_frm_329.frx":0D6C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   7095
         Left            =   60
         TabIndex        =   14
         Top             =   2385
         Width           =   13530
         _Version        =   65536
         _ExtentX        =   23865
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
            Height          =   6630
            Left            =   45
            TabIndex        =   6
            Top             =   405
            Width           =   13470
            _ExtentX        =   23760
            _ExtentY        =   11695
            _Version        =   393216
            Rows            =   26
            Cols            =   20
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_Tit_NumOpe 
            Height          =   285
            Left            =   90
            TabIndex        =   15
            Top             =   90
            Width           =   1230
            _Version        =   65536
            _ExtentX        =   2170
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro. Operación"
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
         Begin Threed.SSPanel pnl_Tit_TipPpg 
            Height          =   285
            Left            =   6615
            TabIndex        =   16
            Top             =   90
            Width           =   2325
            _Version        =   65536
            _ExtentX        =   4101
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Tipo de Prepago"
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
            Left            =   11310
            TabIndex        =   17
            Top             =   90
            Width           =   1860
            _Version        =   65536
            _ExtentX        =   3281
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
         Begin Threed.SSPanel pnl_Tit_FecPro 
            Height          =   285
            Left            =   10125
            TabIndex        =   18
            Top             =   90
            Width           =   1185
            _Version        =   65536
            _ExtentX        =   2090
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Proceso"
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
            Left            =   1320
            TabIndex        =   19
            Top             =   90
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
            Left            =   2535
            TabIndex        =   20
            Top             =   90
            Width           =   4080
            _Version        =   65536
            _ExtentX        =   7197
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
         Begin Threed.SSPanel pnl_Tit_FecPpg 
            Height          =   285
            Left            =   8940
            TabIndex        =   21
            Top             =   90
            Width           =   1185
            _Version        =   65536
            _ExtentX        =   2090
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "F. Prepago"
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
Attribute VB_Name = "frm_Con_PreCon_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private l_rst_Prepagos     As ADODB.Recordset

Private Sub cmd_Buscar_Click()
   If CDate(ipp_FecFin.Text) < CDate(ipp_FecIni.Text) Then
      MsgBox "La Fecha de Fin es menor a la Fecha de Inicio.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecFin)
      Exit Sub
   End If
   If cmb_TipPre.ListIndex = -1 Then
      MsgBox "Debe seleccionar el tipo de prepago.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipPre)
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_Activa(False)
   Call fs_Buscar
   Screen.MousePointer = 0
End Sub

Private Sub cmd_ExpExc_Click()
   If CDate(ipp_FecFin.Text) < CDate(ipp_FecIni.Text) Then
      MsgBox "La Fecha de Fin no puede ser menor a la Fecha de Inicio.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecFin)
      Exit Sub
   End If
   
   'Confirmacion
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
      
   Screen.MousePointer = 11
   Call fs_GenExc
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Limpia_Click()
   Call fs_Limpia
   Call fs_Activa(True)
   Call gs_SetFocus(ipp_FecIni)
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
   cmb_TipPre.Clear
   cmb_TipPre.AddItem "- TODOS -"
   cmb_TipPre.ItemData(cmb_TipPre.NewIndex) = 0
   cmb_TipPre.AddItem "PREPAGO PARCIAL"
   cmb_TipPre.ItemData(cmb_TipPre.NewIndex) = 1
   cmb_TipPre.AddItem "PREPAGO TOTAL"
   cmb_TipPre.ItemData(cmb_TipPre.NewIndex) = 2
   cmb_TipPre.ListIndex = 0
   
   grd_Listad.ColWidth(0) = 1225
   grd_Listad.ColWidth(1) = 1210
   grd_Listad.ColWidth(2) = 4075
   grd_Listad.ColWidth(3) = 2320
   grd_Listad.ColWidth(4) = 1180
   grd_Listad.ColWidth(5) = 1180
   grd_Listad.ColWidth(6) = 1855
   grd_Listad.ColWidth(7) = 0
   grd_Listad.ColWidth(8) = 0
   grd_Listad.ColWidth(9) = 0
   grd_Listad.ColWidth(10) = 0
   grd_Listad.ColWidth(11) = 0
   grd_Listad.ColWidth(12) = 0
   grd_Listad.ColWidth(13) = 0
   grd_Listad.ColWidth(14) = 0
   grd_Listad.ColWidth(15) = 0
   grd_Listad.ColWidth(16) = 0
   grd_Listad.ColWidth(17) = 0
   grd_Listad.ColWidth(18) = 0
   grd_Listad.ColWidth(19) = 0
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignLeftCenter
   grd_Listad.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad.ColAlignment(4) = flexAlignCenterCenter
   grd_Listad.ColAlignment(5) = flexAlignCenterCenter
   grd_Listad.ColAlignment(6) = flexAlignRightCenter
End Sub

Private Sub fs_Limpia()
   ipp_FecIni.Text = "01/01/" & Format(Year(date), "0000")
   ipp_FecFin.Text = Format(date, "DD/MM/YYYY")
   cmb_TipPre.ListIndex = 0
   Call gs_LimpiaGrid(grd_Listad)
End Sub

Private Sub fs_Activa(ByVal p_Activa As Integer)
   ipp_FecIni.Enabled = p_Activa
   ipp_FecFin.Enabled = p_Activa
   cmb_TipPre.Enabled = p_Activa
   grd_Listad.Enabled = Not p_Activa
   cmd_ExpExc.Enabled = Not p_Activa
End Sub

Private Sub fs_Buscar()
   Call gs_LimpiaGrid(grd_Listad)
   g_str_Parame = ""
   
   g_str_Parame = g_str_Parame & "SELECT PP.PPGCAB_NUMOPE, CH.HIPMAE_TDOCLI, CH.HIPMAE_NDOCLI, CH.HIPMAE_MONEDA, PP.PPGCAB_TIPPPG, "
   g_str_Parame = g_str_Parame & "       PP.PPGCAB_FECPRO, PP.PPGCAB_FECPPG, PP.PPGCAB_MTODEP, PP.PPGCAB_MTOTOT, PP.PPGCAB_TIPPPGPAR, "
   g_str_Parame = g_str_Parame & "       TRIM(CL.DATGEN_APEPAT)||' '||TRIM(CL.DATGEN_APEMAT)||' '||TRIM(DATGEN_NOMBRE) AS CLIENTE, "
   g_str_Parame = g_str_Parame & "       TRIM(PR.PRODUC_DESCRI) AS PRODUCTO, CH.HIPMAE_FECDES AS DESEMBOLSO, SH.SOLMAE_PLAANO_CAL*12 AS NRO_CUOTAS, "
   g_str_Parame = g_str_Parame & "       CH.HIPMAE_MTOPRE AS MONTO_CREDITO, PPGCAB_SLDACT_TNC-PPGCAB_APLTNC AS SALDO_TNC, PP.PPGCAB_MOTPPG "
   g_str_Parame = g_str_Parame & "  FROM CRE_PPGCAB PP "
   g_str_Parame = g_str_Parame & " INNER JOIN CRE_HIPMAE CH ON CH.HIPMAE_NUMOPE = PP.PPGCAB_NUMOPE "
   g_str_Parame = g_str_Parame & " INNER JOIN CLI_DATGEN CL ON CL.DATGEN_TIPDOC = CH.HIPMAE_TDOCLI AND CL.DATGEN_NUMDOC = CH.HIPMAE_NDOCLI "
   g_str_Parame = g_str_Parame & " INNER JOIN CRE_PRODUC PR ON PR.PRODUC_CODIGO = CH.HIPMAE_CODPRD"
   g_str_Parame = g_str_Parame & " INNER JOIN CRE_SOLMAE SH ON SH.SOLMAE_NUMERO = CH.HIPMAE_NUMSOL"
   g_str_Parame = g_str_Parame & " WHERE PP.PPGCAB_FECPPG >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " "
   g_str_Parame = g_str_Parame & "   AND PP.PPGCAB_FECPPG <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd")
   If cmb_TipPre.ListIndex <> 0 Then
      g_str_Parame = g_str_Parame & "   AND PP.PPGCAB_TIPPPG = " & cmb_TipPre.ItemData(cmb_TipPre.ListIndex)
   End If
   g_str_Parame = g_str_Parame & " ORDER BY PP.PPGCAB_NUMOPE ASC, PP.PPGCAB_FECPPG ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, l_rst_Prepagos, 3) Then
      Exit Sub
   End If

   If l_rst_Prepagos.BOF And l_rst_Prepagos.EOF Then
      l_rst_Prepagos.Close
      Set l_rst_Prepagos = Nothing
      MsgBox "No se han encontrado registros.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecIni)
      Exit Sub
   End If
   
   Call fs_Activa(False)
   grd_Listad.Redraw = False
   
   l_rst_Prepagos.MoveFirst
   Do While Not l_rst_Prepagos.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
      
      'numero operacion (formateado)
      grd_Listad.Col = 0
      grd_Listad.Text = gf_Formato_NumOpe(Trim(l_rst_Prepagos!PPGCAB_NUMOPE & ""))
      
      'tipo de documento
      grd_Listad.Col = 1
      grd_Listad.Text = CStr(l_rst_Prepagos!HIPMAE_TDOCLI) & "-" & Trim(l_rst_Prepagos!HIPMAE_NDOCLI & "")
      
      'nombre del cliente
      grd_Listad.Col = 2
      grd_Listad.Text = Trim(l_rst_Prepagos!CLIENTE & "")
      
      'tipo de prepago
      grd_Listad.Col = 3
      If l_rst_Prepagos!PPGCAB_TIPPPG = 1 Then
         If l_rst_Prepagos!PPGCAB_TIPPPGPAR = 1 Then
            grd_Listad.Text = "PARCIAL - RED MONTO"
         Else
            grd_Listad.Text = "PARCIAL - RED PLAZO"
         End If
      Else
        grd_Listad.Text = "TOTAL"
      End If
      
      'fecha del prepago (formateado)
      grd_Listad.Col = 4
      grd_Listad.Text = gf_FormatoFecha(CStr(l_rst_Prepagos!PPGCAB_FECPPG))
      
      'fecha de proceso (formateado)
      grd_Listad.Col = 5
      grd_Listad.Text = gf_FormatoFecha(CStr(l_rst_Prepagos!PPGCAB_FECPRO))
      
      'importe del prepago (formateado)
      grd_Listad.Col = 6
      If l_rst_Prepagos!PPGCAB_TIPPPG = 1 Then
         grd_Listad.Text = Format(l_rst_Prepagos!PPGCAB_MTODEP, "###,###,##0.00")
      Else
         grd_Listad.Text = Format(l_rst_Prepagos!PPGCAB_MTOTOT, "###,###,##0.00")
      End If
      
      ' Tipo de prepago (parcial o total)
      grd_Listad.Col = 7
      grd_Listad.Text = Trim(l_rst_Prepagos!PPGCAB_TIPPPG)
      
      ' Tipo de prepago Parcial (monto o tiempo)
      grd_Listad.Col = 8
      grd_Listad.Text = Trim(l_rst_Prepagos!PPGCAB_TIPPPGPAR & "")
      
      'numero de operacion
      grd_Listad.Col = 9
      grd_Listad.Text = Trim(l_rst_Prepagos!PPGCAB_NUMOPE & "")
      
      'fecha de prepago
      grd_Listad.Col = 10
      grd_Listad.Text = CStr(l_rst_Prepagos!PPGCAB_FECPPG)
      
      'fecha de proceso
      grd_Listad.Col = 11
      grd_Listad.Text = CStr(l_rst_Prepagos!PPGCAB_FECPRO)
      
      'importe del prepago
      grd_Listad.Col = 12
      grd_Listad.Text = l_rst_Prepagos!PPGCAB_MTOTOT
      
      'moneda del prestamos
      grd_Listad.Col = 13
      If l_rst_Prepagos!HIPMAE_MONEDA = 1 Then
         grd_Listad.Text = "SOLES"
      Else
         grd_Listad.Text = "DOLARES"
      End If
      
      'Producto
      grd_Listad.Col = 14
      grd_Listad.Text = l_rst_Prepagos!PRODUCTO
      
      'Fecha de desembolso
      grd_Listad.Col = 15
      grd_Listad.Text = gf_FormatoFecha(CStr(l_rst_Prepagos!DESEMBOLSO))
      
      'Monto del credito original
      grd_Listad.Col = 16
      grd_Listad.Text = Format(l_rst_Prepagos!MONTO_CREDITO, "###,###,##0.00")
      
      'Numero de Cuotas del prestamo original
      grd_Listad.Col = 17
      grd_Listad.Text = l_rst_Prepagos!NRO_CUOTAS
            
      'Saldo TNC luego del prepago
      grd_Listad.Col = 18
      grd_Listad.Text = Format(l_rst_Prepagos!SALDO_TNC, "###,###,##0.00")
      
      'Motivo del prepago
      grd_Listad.Col = 19
      'grd_Listad.Text = IIf((CStr(l_rst_Prepagos!PPGCAB_MOTPPG) = 0) Or IsNull(l_rst_Prepagos!PPGCAB_MOTPPG), "", moddat_gf_Consulta_ParDes("115", CStr(l_rst_Prepagos!PPGCAB_MOTPPG)))
      If IsNull(l_rst_Prepagos!PPGCAB_MOTPPG) Then
         grd_Listad.Text = ""
      Else
         If CStr(l_rst_Prepagos!PPGCAB_MOTPPG) = 0 Then
            grd_Listad.Text = ""
         Else
            grd_Listad.Text = moddat_gf_Consulta_ParDes("115", CStr(l_rst_Prepagos!PPGCAB_MOTPPG))
         End If
      End If
      
      l_rst_Prepagos.MoveNext
   Loop
   grd_Listad.Redraw = True
   
   If grd_Listad.Rows > 0 Then
      grd_Listad.Enabled = True
   End If
   
   Call gs_UbiIniGrid(grd_Listad)
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub fs_GenExc()
Dim r_obj_Excel      As Excel.Application
Dim r_int_NroFil     As Integer
Dim r_int_nroaux     As Integer

   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      r_int_NroFil = 2
      .Cells(r_int_NroFil, 1) = "REPORTE DE PREPAGOS (" & CStr(ipp_FecIni.Text) & " - " & CStr(ipp_FecFin.Text) & ")"
      .Range(.Cells(r_int_NroFil, 1), .Cells(r_int_NroFil, 15)).Font.Bold = True
      .Range(.Cells(r_int_NroFil, 1), .Cells(r_int_NroFil, 15)).Merge
      .Range(.Cells(r_int_NroFil, 1), .Cells(r_int_NroFil, 15)).HorizontalAlignment = xlHAlignCenter
      
      r_int_NroFil = 4
      .Cells(r_int_NroFil, 1) = "ITEM"
      .Cells(r_int_NroFil, 2) = "PRODUCTO"
      .Cells(r_int_NroFil, 3) = "FEC. DESEMB."
      .Cells(r_int_NroFil, 4) = "MONTO CREDITO"
      .Cells(r_int_NroFil, 5) = "NRO. CUOTAS"
      .Cells(r_int_NroFil, 6) = "OPERACION"
      .Cells(r_int_NroFil, 7) = "DOC. IDENTIDAD"
      .Cells(r_int_NroFil, 8) = "NOMBRE CLIENTE"
      .Cells(r_int_NroFil, 9) = "TIPO PREPAGO"
      .Cells(r_int_NroFil, 10) = "F. PREPAGO"
      .Cells(r_int_NroFil, 11) = "F. PROCESO"
      .Cells(r_int_NroFil, 12) = "MONEDA"
      .Cells(r_int_NroFil, 13) = "MTO. PREPAGO"
      .Cells(r_int_NroFil, 14) = "SALDO CREDITO"
      .Cells(r_int_NroFil, 15) = "MOTIVO DEL PREPAGO"
      .Range(.Cells(r_int_NroFil, 1), .Cells(r_int_NroFil, 15)).Font.Bold = True
      .Range(.Cells(r_int_NroFil, 1), .Cells(r_int_NroFil, 15)).HorizontalAlignment = xlHAlignCenter
       
      .Columns("A").ColumnWidth = 6
      .Columns("B").ColumnWidth = 45
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 13
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 16
      .Columns("D").HorizontalAlignment = xlHAlignRight
      .Columns("E").ColumnWidth = 13
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").ColumnWidth = 13
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("G").ColumnWidth = 15
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      .Columns("H").ColumnWidth = 50
      .Columns("H").HorizontalAlignment = xlHAlignLeft
      .Columns("I").ColumnWidth = 22
      .Columns("I").HorizontalAlignment = xlHAlignCenter
      .Columns("J").ColumnWidth = 13
      .Columns("J").HorizontalAlignment = xlHAlignCenter
      .Columns("K").ColumnWidth = 13
      .Columns("K").HorizontalAlignment = xlHAlignCenter
      .Columns("L").ColumnWidth = 13
      .Columns("L").HorizontalAlignment = xlHAlignCenter
      .Columns("M").ColumnWidth = 15
      .Columns("M").HorizontalAlignment = xlHAlignRight
      .Columns("N").ColumnWidth = 15
      .Columns("N").HorizontalAlignment = xlHAlignRight
      .Columns("O").ColumnWidth = 50
      .Columns("O").HorizontalAlignment = xlHAlignLeft
      
      r_int_NroFil = r_int_NroFil + 1
      For r_int_nroaux = 0 To grd_Listad.Rows - 1
         .Cells(r_int_NroFil, 1) = Format(r_int_NroFil - 4, "00#")
         .Cells(r_int_NroFil, 2) = grd_Listad.TextMatrix(r_int_nroaux, 14)
         .Cells(r_int_NroFil, 3) = "'" & grd_Listad.TextMatrix(r_int_nroaux, 15)
         .Cells(r_int_NroFil, 4) = grd_Listad.TextMatrix(r_int_nroaux, 16)
         .Cells(r_int_NroFil, 5) = grd_Listad.TextMatrix(r_int_nroaux, 17)
         .Cells(r_int_NroFil, 6) = grd_Listad.TextMatrix(r_int_nroaux, 0)
         .Cells(r_int_NroFil, 7) = grd_Listad.TextMatrix(r_int_nroaux, 1)
         .Cells(r_int_NroFil, 8) = grd_Listad.TextMatrix(r_int_nroaux, 2)
         .Cells(r_int_NroFil, 9) = grd_Listad.TextMatrix(r_int_nroaux, 3)
         .Cells(r_int_NroFil, 10) = "'" & grd_Listad.TextMatrix(r_int_nroaux, 4)
         .Cells(r_int_NroFil, 11) = "'" & grd_Listad.TextMatrix(r_int_nroaux, 5)
         .Cells(r_int_NroFil, 12) = grd_Listad.TextMatrix(r_int_nroaux, 13)
         .Cells(r_int_NroFil, 13) = grd_Listad.TextMatrix(r_int_nroaux, 6)
         .Cells(r_int_NroFil, 14) = grd_Listad.TextMatrix(r_int_nroaux, 18)
         .Cells(r_int_NroFil, 15) = grd_Listad.TextMatrix(r_int_nroaux, 19)
         r_int_NroFil = r_int_NroFil + 1
      Next
      .Cells(1, 1).Select
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub grd_Listad_DblClick()
Dim r_int_TipPpg    As Integer
Dim r_int_PpgPar    As Integer

   If grd_Listad.Rows > 0 Then
      grd_Listad.Col = 7
      r_int_TipPpg = CInt(grd_Listad.Text)
      
      grd_Listad.Col = 8
      r_int_PpgPar = CInt(grd_Listad.Text)
      
      grd_Listad.Col = 9
      moddat_g_str_NumOpe = Trim(grd_Listad.Text)
      
      grd_Listad.Col = 10
      moddat_g_str_FecIng = Trim(grd_Listad.Text)
      
      If r_int_TipPpg = 1 Then
         moddat_g_int_FlgPre = 1
         frm_Con_PrePgo_03.Show 1
      Else
         moddat_g_int_FlgPre = 2
         frm_Con_PrePgo_04.Show 1
      End If
   End If
End Sub

Private Sub ipp_FecIni_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecFin)
   End If
End Sub

Private Sub ipp_FecFin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipPre)
   End If
End Sub

Private Sub cmb_TipPre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Buscar)
   End If
End Sub

'Reordenar
Private Sub pnl_Tit_NumOpe_Click()
   If Len(Trim(pnl_Tit_NumOpe.Tag)) = 0 Or pnl_Tit_NumOpe.Tag = "D" Then
      pnl_Tit_NumOpe.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 0, "C")
   Else
      pnl_Tit_NumOpe.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 0, "C-")
   End If
End Sub

Private Sub pnl_Tit_DoiCli_Click()
   If Len(Trim(pnl_Tit_DoiCli.Tag)) = 0 Or pnl_Tit_DoiCli.Tag = "D" Then
      pnl_Tit_DoiCli.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 1, "C")
   Else
      pnl_Tit_DoiCli.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 1, "C-")
   End If
End Sub

Private Sub pnl_Tit_NomCli_Click()
   If Len(Trim(pnl_Tit_NomCli.Tag)) = 0 Or pnl_Tit_NomCli.Tag = "D" Then
      pnl_Tit_NomCli.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 2, "C")
   Else
      pnl_Tit_NomCli.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 2, "C-")
   End If
End Sub

Private Sub pnl_Tit_TipPpg_Click()
   If Len(Trim(pnl_Tit_TipPpg.Tag)) = 0 Or pnl_Tit_TipPpg.Tag = "D" Then
      pnl_Tit_TipPpg.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 3, "C")
   Else
      pnl_Tit_TipPpg.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 3, "C-")
   End If
End Sub

Private Sub pnl_Tit_FecPpg_Click()
   If Len(Trim(pnl_Tit_FecPpg.Tag)) = 0 Or pnl_Tit_FecPpg.Tag = "D" Then
      pnl_Tit_FecPpg.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 10, "C")
   Else
      pnl_Tit_FecPpg.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 10, "C-")
   End If
End Sub

Private Sub pnl_Tit_FecPro_Click()
   If Len(Trim(pnl_Tit_FecPro.Tag)) = 0 Or pnl_Tit_FecPro.Tag = "D" Then
      pnl_Tit_FecPro.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 11, "C")
   Else
      pnl_Tit_FecPro.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 11, "C-")
   End If
End Sub

Private Sub pnl_Tit_Import_Click()
   If Len(Trim(pnl_Tit_Import.Tag)) = 0 Or pnl_Tit_Import.Tag = "D" Then
      pnl_Tit_Import.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 12, "N")
   Else
      pnl_Tit_Import.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 12, "N-")
   End If
End Sub
