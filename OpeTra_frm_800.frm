VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_Rpt_CreDes_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3615
   ClientLeft      =   7845
   ClientTop       =   5910
   ClientWidth     =   7170
   Icon            =   "OpeTra_frm_800.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   7170
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   3615
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   7185
      _Version        =   65536
      _ExtentX        =   12674
      _ExtentY        =   6376
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
         Height          =   2115
         Left            =   30
         TabIndex        =   13
         Top             =   1440
         Width           =   7095
         _Version        =   65536
         _ExtentX        =   12515
         _ExtentY        =   3731
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
         Begin VB.ComboBox cmb_TipPro 
            Height          =   315
            Left            =   1140
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   750
            Width           =   5895
         End
         Begin VB.CheckBox chk_TipPro 
            Caption         =   "Todos los Productos"
            Height          =   285
            Left            =   1140
            TabIndex        =   3
            Top             =   1080
            Width           =   1995
         End
         Begin VB.CheckBox chk_Empres 
            Caption         =   "Todos las Empresas"
            Height          =   285
            Left            =   1140
            TabIndex        =   1
            Top             =   420
            Width           =   1995
         End
         Begin VB.ComboBox cmb_Empres 
            Height          =   315
            Left            =   1140
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   5895
         End
         Begin EditLib.fpDateTime ipp_FecIni 
            Height          =   315
            Left            =   1140
            TabIndex        =   4
            Top             =   1380
            Width           =   1425
            _Version        =   196608
            _ExtentX        =   2514
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
            Left            =   1140
            TabIndex        =   5
            Top             =   1740
            Width           =   1425
            _Version        =   196608
            _ExtentX        =   2514
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
         Begin VB.Label Label2 
            Caption         =   "Fecha Inicio:"
            Height          =   315
            Left            =   60
            TabIndex        =   17
            Top             =   1380
            Width           =   975
         End
         Begin VB.Label Label3 
            Caption         =   "Fecha Fin:"
            Height          =   285
            Left            =   60
            TabIndex        =   16
            Top             =   1770
            Width           =   885
         End
         Begin VB.Label Label1 
            Caption         =   "Producto:"
            Height          =   255
            Left            =   60
            TabIndex        =   15
            Top             =   750
            Width           =   795
         End
         Begin VB.Label Label4 
            Caption         =   "Empresa:"
            Height          =   255
            Left            =   60
            TabIndex        =   14
            Top             =   60
            Width           =   915
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   10
         Top             =   30
         Width           =   7095
         _Version        =   65536
         _ExtentX        =   12515
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
            Height          =   570
            Left            =   630
            TabIndex        =   11
            Top             =   45
            Width           =   4875
            _Version        =   65536
            _ExtentX        =   8599
            _ExtentY        =   1005
            _StockProps     =   15
            Caption         =   "Reporte de Créditos Hipotecarios Desembolsados"
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
            Left            =   6600
            Top             =   30
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
            Left            =   90
            Picture         =   "OpeTra_frm_800.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   12
         Top             =   750
         Width           =   7095
         _Version        =   65536
         _ExtentX        =   12515
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
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   6480
            Picture         =   "OpeTra_frm_800.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_800.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   630
            Picture         =   "OpeTra_frm_800.frx":0B9A
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
      End
   End
End
Attribute VB_Name = "frm_Rpt_CreDes_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_arr_Produc()      As moddat_tpo_Genera
Dim l_arr_Empres()      As moddat_tpo_Genera
Dim l_str_Fecha         As String
Dim l_str_Hora          As String

Private Sub chk_Empres_Click()
   If chk_Empres.Value = 1 Then
      cmb_Empres.ListIndex = -1
      cmb_Empres.Enabled = False
      If cmb_TipPro.Enabled Then
         Call gs_SetFocus(cmb_TipPro)
      Else
         Call gs_SetFocus(ipp_FecIni)
      End If
   
   ElseIf chk_Empres.Value = 0 Then
      cmb_Empres.Enabled = True
      Call gs_SetFocus(cmb_Empres)
   End If
End Sub

Private Sub chk_TipPro_Click()
   If chk_TipPro.Value = 1 Then
      cmb_TipPro.ListIndex = -1
      cmb_TipPro.Enabled = False
      Call gs_SetFocus(ipp_FecIni)
      
   ElseIf chk_TipPro.Value = 0 Then
      cmb_TipPro.Enabled = True
      Call gs_SetFocus(cmb_TipPro)
   End If
End Sub

Private Sub cmb_Empres_Click()
   If cmb_TipPro.Enabled Then
      Call gs_SetFocus(cmb_TipPro)
   Else
      Call gs_SetFocus(ipp_FecIni)
   End If
End Sub

Private Sub cmb_Empres_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_Empres_Click
   End If
End Sub

Private Sub cmb_TipPro_Click()
   Call gs_SetFocus(ipp_FecIni)
End Sub

Private Sub cmb_TipPro_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipPro_Click
   End If
End Sub

Private Sub cmd_ExpExc_Click()
   If chk_Empres.Value = 0 Then
      If cmb_Empres.ListIndex = -1 Then
         MsgBox "Debe seleccionar la Empresa.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_Empres)
         Exit Sub
      End If
   End If
   If chk_TipPro.Value = 0 Then
      If cmb_TipPro.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Producto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_TipPro)
         Exit Sub
      End If
   End If
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

Private Sub cmd_Imprim_Click()
Dim r_str_TIPMON As String
      
   If chk_Empres.Value = 0 Then
      If cmb_Empres.ListIndex = -1 Then
         MsgBox "Debe seleccionar la Empresa.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_Empres)
         Exit Sub
      End If
   End If
   If chk_TipPro.Value = 0 Then
      If cmb_TipPro.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Producto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_TipPro)
         Exit Sub
      End If
   End If
   If CDate(ipp_FecFin.Text) < CDate(ipp_FecIni.Text) Then
      MsgBox "La Fecha de Fin no puede ser menor a la Fecha de Inicio.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecFin)
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de imprimir el Reporte?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If

   'Proceso
   Screen.MousePointer = 11
   
   'LLenamos las variables con la fecha y hora del sistema
   l_str_Fecha = Format(date, "yyyymmdd")
   l_str_Hora = Format(Time, "hhmmss")
      
   'Eliminamos el contenido de la tabla si es q existiera
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "DELETE FROM RPT_CREDES "
   g_str_Parame = g_str_Parame & " WHERE CREDES_NOMRPT = 'CTB_RPTSOL_01.RPT' "
   g_str_Parame = g_str_Parame & "   AND CREDES_TERCRE = '" & modgen_g_str_NombPC & "'"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
      Exit Sub
   End If
   
   'Leyendo Tabla de solicitudes
   g_str_Parame = "SELECT * FROM CRE_HIPMAE, CLI_DATGEN WHERE "
   g_str_Parame = g_str_Parame & "HIPMAE_TDOCLI = DATGEN_TIPDOC AND "
   g_str_Parame = g_str_Parame & "HIPMAE_NDOCLI = DATGEN_NUMDOC AND "
   
   If chk_Empres.Value = 0 Then
      g_str_Parame = g_str_Parame & "HIPMAE_PROCRE = '" & l_arr_Empres(cmb_Empres.ListIndex + 1).Genera_Codigo & "' AND "
   End If
   
   If chk_TipPro.Value = 0 Then
      g_str_Parame = g_str_Parame & "HIPMAE_CODPRD = '" & l_arr_Produc(cmb_TipPro.ListIndex + 1).Genera_Codigo & "' AND "
   End If
   
   g_str_Parame = g_str_Parame & "HIPMAE_FECDES >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND "
   g_str_Parame = g_str_Parame & "HIPMAE_FECDES <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & " AND "
   g_str_Parame = g_str_Parame & "HIPMAE_SITUAC = 2 "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         'Para obtener Descripción de Ultima Ocurrencia (Situación de Instancia)
         r_str_TIPMON = moddat_gf_Consulta_ParDes("204", CStr(g_rst_Princi!HIPMAE_MONEDA))
        
         'Insertando Registro
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & "INSERT INTO RPT_CREDES("
         g_str_Parame = g_str_Parame & "CREDES_NOMRPT, "
         g_str_Parame = g_str_Parame & "CREDES_FECCRE, "
         g_str_Parame = g_str_Parame & "CREDES_HORCRE, "
         g_str_Parame = g_str_Parame & "CREDES_TERCRE, "
         g_str_Parame = g_str_Parame & "CREDES_NUMOPE, "
         g_str_Parame = g_str_Parame & "CREDES_FECINI, "
         g_str_Parame = g_str_Parame & "CREDES_FECFIN, "
         g_str_Parame = g_str_Parame & "CREDES_CODPRD, "
         g_str_Parame = g_str_Parame & "CREDES_TIPMON, "
         g_str_Parame = g_str_Parame & "CREDES_TIPDOC, "
         g_str_Parame = g_str_Parame & "CREDES_NUMDOC, "
         g_str_Parame = g_str_Parame & "CREDES_APEPAT, "
         g_str_Parame = g_str_Parame & "CREDES_APEMAT, "
         g_str_Parame = g_str_Parame & "CREDES_NOMBRE, "
         g_str_Parame = g_str_Parame & "CREDES_MTOPRE, "
         g_str_Parame = g_str_Parame & "CREDES_TOTPRE, "
         g_str_Parame = g_str_Parame & "CREDES_INTCAP, "
         g_str_Parame = g_str_Parame & "CREDES_TASINT, "
         g_str_Parame = g_str_Parame & "CREDES_NUMCUO, "
         g_str_Parame = g_str_Parame & "CREDES_PERGRA, "
         g_str_Parame = g_str_Parame & "CREDES_FECDES, "
         g_str_Parame = g_str_Parame & "CREDES_FECACT, "
         g_str_Parame = g_str_Parame & "CREDES_IMPNCO, "
         g_str_Parame = g_str_Parame & "CREDES_IMPCON, "
         g_str_Parame = g_str_Parame & "CREDES_IMPDES, "
         g_str_Parame = g_str_Parame & "CREDES_COSEFE, "
         g_str_Parame = g_str_Parame & "CREDES_EMPRES, "
         g_str_Parame = g_str_Parame & "CREDES_MTOCVT, "
         g_str_Parame = g_str_Parame & "CREDES_APOPRO, "
         g_str_Parame = g_str_Parame & "CREDES_PORINI) "
                           
         g_str_Parame = g_str_Parame & "VALUES ("
         g_str_Parame = g_str_Parame & "'CTB_RPTSOL_01.RPT', "
         g_str_Parame = g_str_Parame & l_str_Fecha & ", "
         g_str_Parame = g_str_Parame & l_str_Hora & ", "
         g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!HIPMAE_NUMOPE & "', "
         g_str_Parame = g_str_Parame & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & ", "
         g_str_Parame = g_str_Parame & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & ", "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!HIPMAE_CODPRD & "', "
         g_str_Parame = g_str_Parame & "'" & r_str_TIPMON & "', "
         g_str_Parame = g_str_Parame & g_rst_Princi!HIPMAE_TDOCLI & ", "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!HIPMAE_NDOCLI & "', "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!DatGen_ApePat & "', "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!DatGen_ApeMat & "', "
         g_str_Parame = g_str_Parame & "'" & g_rst_Princi!DatGen_Nombre & "', "
         g_str_Parame = g_str_Parame & g_rst_Princi!HIPMAE_MTOPRE & ", "
         g_str_Parame = g_str_Parame & g_rst_Princi!HIPMAE_TOTPRE & ", "
         g_str_Parame = g_str_Parame & g_rst_Princi!HIPMAE_INTCAP & ", "
         g_str_Parame = g_str_Parame & g_rst_Princi!HIPMAE_TASINT & ", "
         g_str_Parame = g_str_Parame & g_rst_Princi!HIPMAE_NUMCUO & ", "
         g_str_Parame = g_str_Parame & g_rst_Princi!HIPMAE_PERGRA & ", "
         g_str_Parame = g_str_Parame & g_rst_Princi!HIPMAE_FECDES & ", "
         g_str_Parame = g_str_Parame & g_rst_Princi!HIPMAE_FECACT & ", "
         g_str_Parame = g_str_Parame & g_rst_Princi!HIPMAE_IMPNCO & ", "
         g_str_Parame = g_str_Parame & g_rst_Princi!HIPMAE_IMPCON & ", "
         g_str_Parame = g_str_Parame & g_rst_Princi!HIPMAE_IMPDES & ","
         g_str_Parame = g_str_Parame & g_rst_Princi!HIPMAE_COSEFE & ", "
         g_str_Parame = g_str_Parame & "'" & moddat_gf_ConsultaEmpGrp(g_rst_Princi!HIPMAE_PROCRE) & "',"
      
         If g_rst_Princi!HIPMAE_MONEDA = 2 Then
            g_str_Parame = g_str_Parame & g_rst_Princi!HIPMAE_CVTDOL & ", "
            g_str_Parame = g_str_Parame & g_rst_Princi!HIPMAE_APODOL & ", "
            g_str_Parame = g_str_Parame & Format(g_rst_Princi!HIPMAE_APODOL / g_rst_Princi!HIPMAE_CVTDOL * 100, "##0.00") & ") "
         Else
            g_str_Parame = g_str_Parame & g_rst_Princi!HIPMAE_CVTSOL & ", "
            g_str_Parame = g_str_Parame & g_rst_Princi!HIPMAE_APOSOL & ", "
            g_str_Parame = g_str_Parame & Format(g_rst_Princi!HIPMAE_APOSOL / g_rst_Princi!HIPMAE_CVTSOL * 100, "##0.00") & ") "
         End If
                
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
            Exit Sub
         End If
         
         g_rst_Princi.MoveNext
      Loop
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   Else
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Screen.MousePointer = 0
      MsgBox "No se encontraron Solicitudes registradas.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
      
   'Se envia la cadena de conexión
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
   crp_Imprim.DataFiles(0) = "CRE_PRODUC"
   crp_Imprim.DataFiles(1) = "RPT_CREDES"
            
   crp_Imprim.SelectionFormula = "{RPT_CREDES.CREDES_NOMRPT} = 'CTB_RPTSOL_01.RPT' " & " AND "
   crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{RPT_CREDES.CREDES_TERCRE} = '" & modgen_g_str_NombPC & "'"
      
   'Se hace la invocación y llamado del Reporte en la ubicación correspondiente
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "CTB_RPTSOL_01.RPT"
      
   'Se le envia el destino a una ventana de crystal report
   crp_Imprim.WindowShowPrintSetupBtn = True
   crp_Imprim.Destination = crptToWindow
   crp_Imprim.Action = 1
   
   'El puntero del mouse regresa al estado normal
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call gs_CentraForm(Me)
   
   Call gs_SetFocus(cmb_Empres)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_EmpGrp(cmb_Empres, l_arr_Empres)
   Call moddat_gs_Carga_Produc(cmb_TipPro, l_arr_Produc, 4)
   ipp_FecIni.Text = Format(date, "dd/mm/yyyy")
   ipp_FecFin.Text = Format(date, "dd/mm/yyyy")
End Sub

Private Sub fs_Limpia()
   cmb_Empres.ListIndex = -1
   chk_Empres.Value = 0
   cmb_TipPro.ListIndex = -1
   chk_TipPro.Value = 0
   ipp_FecIni.Text = Format(date - CDate(30), "dd/mm/yyyy")
   ipp_FecFin.Text = Format(date, "dd/mm/yyyy")
End Sub

Private Sub fs_GenExc()
Dim r_obj_Excel      As Excel.Application
Dim r_int_ConVer     As Integer
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT A.HIPMAE_CODPRD, A.HIPMAE_NUMOPE, A.HIPMAE_TDOCLI, A.HIPMAE_NDOCLI, A.HIPMAE_FECACT, A.HIPMAE_FECDES, "
   g_str_Parame = g_str_Parame & "       A.HIPMAE_MONEDA, A.HIPMAE_IMPDES, A.HIPMAE_MTOPRE, A.HIPMAE_INTCAP, A.HIPMAE_TOTPRE, A.HIPMAE_IMPNCO, "
   g_str_Parame = g_str_Parame & "       A.HIPMAE_IMPCON, A.HIPMAE_TASINT, A.HIPMAE_COSEFE, A.HIPMAE_NUMCUO, A.HIPMAE_PERGRA, A.HIPMAE_CVTDOL, "
   g_str_Parame = g_str_Parame & "       A.HIPMAE_APODOL, A.HIPMAE_CVTSOL, A.HIPMAE_APOSOL, C.PRODUC_DESCRI, A.HIPMAE_PRYMCS,  "
   g_str_Parame = g_str_Parame & "       DECODE(A.HIPMAE_MONEDA, 1, F.SOLMAE_APOPRO_SOL, F.SOLMAE_APOPRO_DOL) - F.SOLMAE_PBPMTO - F.SOLMAE_FMVBBP - F.SOLMAE_BMSMTO - F.SOLMAE_AFPMTO AS AP_PROPIO, "
   g_str_Parame = g_str_Parame & "       F.SOLMAE_PBPMTO AS MONTO_PBP, F.SOLMAE_FMVBBP AS MONTO_BBP, F.SOLMAE_BMSMTO AS MONTO_BMS, F.SOLMAE_AFPMTO AS MONTO_AFP, "
   g_str_Parame = g_str_Parame & "       TRIM(B.DATGEN_APEPAT)||' '||TRIM(B.DATGEN_APEMAT)||' '||TRIM(B.DATGEN_NOMBRE) AS NOMBRE_CLIENTE, "
   g_str_Parame = g_str_Parame & "       TRIM(E.PARDES_DESCRI) AS TIPO_GARANTIA, TRIM(D.PARDES_DESCRI) AS MONEDA, SOLMAE_MTOGCI, SOLMAE_PREMTO, "
   g_str_Parame = g_str_Parame & "       TRIM(Y.PARDES_DESCRI) AS ESTADO_CREDITO, TRIM(J.SUBPRD_DESCRI) AS SUB_PRODUCTO,  "
   g_str_Parame = g_str_Parame & "       TRIM(NVL(DECODE(G.SOLINM_PRYCOD, 1, G.SOLINM_PRYNOM, DECODE(G.SOLINM_PRYCOD, NULL, G.SOLINM_PRYNOM, I.DATGEN_TITULO)), '-')) AS PROYECTO, "
   g_str_Parame = g_str_Parame & "       NVL(CASE WHEN SOLINM_TIPDOC_PRO = 7 THEN TRIM(X.DATGEN_RAZSOC) ELSE  TRIM(SOLINM_RAZSOC_PRO) END, '-') AS PROMOTOR, "
   g_str_Parame = g_str_Parame & "       NVL(CASE WHEN SOLINM_TIPDOC_CON = 0 THEN '-' WHEN SOLINM_TIPDOC_CON = 1 THEN TRIM(SOLINM_RAZSOC_CON) ELSE TRIM(Z.DATGEN_RAZSOC) END, '-') AS CONSTRUCTOR "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPMAE A "
   g_str_Parame = g_str_Parame & " INNER JOIN CLI_DATGEN B ON B.DATGEN_TIPDOC = A.HIPMAE_TDOCLI AND B.DATGEN_NUMDOC = A.HIPMAE_NDOCLI "
   g_str_Parame = g_str_Parame & " INNER JOIN CRE_PRODUC C ON C.PRODUC_CODIGO = A.HIPMAE_CODPRD "
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES D ON D.PARDES_CODGRP = 204 AND D.PARDES_CODITE = A.HIPMAE_MONEDA "
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES E ON E.PARDES_CODGRP = 241 AND E.PARDES_CODITE = A.HIPMAE_TIPGAR "
   g_str_Parame = g_str_Parame & " INNER JOIN CRE_SOLMAE F ON F.SOLMAE_NUMERO = A.HIPMAE_NUMSOL "
   g_str_Parame = g_str_Parame & "  LEFT JOIN CRE_SOLINM G ON G.SOLINM_NUMSOL = A.HIPMAE_NUMSOL "
   g_str_Parame = g_str_Parame & "  LEFT JOIN CLI_DATGEN H ON H.DATGEN_TIPDOC = F.SOLMAE_TITTDO AND H.DATGEN_NUMDOC = F.SOLMAE_TITNDO "
   g_str_Parame = g_str_Parame & "  LEFT JOIN PRY_DATGEN I ON I.DATGEN_CODIGO = G.SOLINM_PRYCOD "
   g_str_Parame = g_str_Parame & "  LEFT JOIN EMP_DATGEN X ON X.DATGEN_EMPTDO = G.SOLINM_TIPDOC_PRO AND X.DATGEN_EMPNDO = G.SOLINM_NUMDOC_PRO "
   g_str_Parame = g_str_Parame & "  LEFT JOIN EMP_DATGEN Z ON Z.DATGEN_EMPTDO = G.SOLINM_TIPDOC_CON AND Z.DATGEN_EMPNDO = G.SOLINM_NUMDOC_CON "
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES Y ON Y.PARDES_CODGRP = 027 AND Y.PARDES_CODITE = A.HIPMAE_SITUAC "
   g_str_Parame = g_str_Parame & " INNER JOIN CRE_SUBPRD J ON J.SUBPRD_CODPRD = A.HIPMAE_CODPRD AND J.SUBPRD_CODSUB = A.HIPMAE_CODSUB "
   g_str_Parame = g_str_Parame & " WHERE A.HIPMAE_SITUAC IN (2,6,9) "
   g_str_Parame = g_str_Parame & "   AND A.HIPMAE_FECDES >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " "
   g_str_Parame = g_str_Parame & "   AND A.HIPMAE_FECDES <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & " "
   If chk_Empres.Value = 0 Then
      g_str_Parame = g_str_Parame & "   AND A.HIPMAE_PROCRE = '" & l_arr_Empres(cmb_Empres.ListIndex + 1).Genera_Codigo & "' "
   End If
   If chk_TipPro.Value = 0 Then
      g_str_Parame = g_str_Parame & "   AND A.HIPMAE_CODPRD = '" & l_arr_Produc(cmb_TipPro.ListIndex + 1).Genera_Codigo & "' "
   End If
   g_str_Parame = g_str_Parame & "ORDER BY A.HIPMAE_CODPRD ASC, B.DATGEN_APEPAT ASC, B.DATGEN_APEMAT ASC, B.DATGEN_NOMBRE ASC "
  
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se encontraron Solicitudes registradas.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Cells(1, 1) = "ITEM"
      .Cells(1, 2) = "PRODUCTO"
      .Cells(1, 3) = "SUB PRODUCTO"
      .Cells(1, 4) = "OPERACION"
      .Cells(1, 5) = "DOC. IDENTIDAD"
      .Cells(1, 6) = "NOMBRE CLIENTE"
      .Cells(1, 7) = "F. ACTIV."
      .Cells(1, 8) = "F. DESEMB."
      .Cells(1, 9) = "SIT. CREDITO"
      .Cells(1, 10) = "TIPO DE MONEDA"
      .Cells(1, 11) = "V. COMPRA-VENTA"
      .Cells(1, 12) = "CUOTA INICIAL"
      .Cells(1, 13) = "% INICIAL"
      .Cells(1, 14) = "AP. PROPIO"
      .Cells(1, 15) = "MONTO PBP"
      .Cells(1, 16) = "MONTO BBP"
      .Cells(1, 17) = "MONTO BMS"
      .Cells(1, 18) = "MONTO AFP"
      .Cells(1, 19) = "MONTO PRESTAMO"
      .Cells(1, 20) = "GASTOS CIERRE"
      .Cells(1, 21) = "INT. CAPIT."
      .Cells(1, 22) = "TOTAL PREST."
      .Cells(1, 23) = "M. PREST. TNC"
      .Cells(1, 24) = "M. PREST. TC"
      .Cells(1, 25) = "TASA INT."
      .Cells(1, 26) = "COSTO EFECTIVO"
      .Cells(1, 27) = "CUOTAS"
      .Cells(1, 28) = "P. GRACIA"
      .Cells(1, 29) = "TIPO DE GARANTIA"
      .Cells(1, 30) = "VINCULADO"
      .Cells(1, 31) = "NOMBRE DE PROYECTO"
      .Cells(1, 32) = "NOMBRE PROMOTOR"
      .Cells(1, 33) = "NOMBRE CONSTRUCTOR"
      .Range(.Cells(1, 1), .Cells(1, 33)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 33)).HorizontalAlignment = xlHAlignCenter
      
      .Columns("A").ColumnWidth = 6
      .Columns("A").HorizontalAlignment = xlHAlignCenter
      .Columns("B").ColumnWidth = 50
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 70
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 15
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").ColumnWidth = 15
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").ColumnWidth = 44
      .Columns("G").ColumnWidth = 12
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      .Columns("H").ColumnWidth = 12
      .Columns("H").HorizontalAlignment = xlHAlignCenter
      .Columns("I").ColumnWidth = 19
      .Columns("I").HorizontalAlignment = xlHAlignCenter
      .Columns("J").ColumnWidth = 18
      .Columns("J").HorizontalAlignment = xlHAlignCenter
      .Columns("K").ColumnWidth = 20
      .Columns("L").ColumnWidth = 16
      .Columns("M").ColumnWidth = 16
      .Columns("N").ColumnWidth = 16
      .Columns("O").ColumnWidth = 16
      .Columns("P").ColumnWidth = 16
      .Columns("Q").ColumnWidth = 16
      .Columns("R").ColumnWidth = 16
      .Columns("S").ColumnWidth = 17
      .Columns("T").ColumnWidth = 18
      .Columns("U").ColumnWidth = 16
      .Columns("V").ColumnWidth = 16
      .Columns("W").ColumnWidth = 16
      .Columns("X").ColumnWidth = 16
      .Columns("Y").ColumnWidth = 12
      .Columns("Z").ColumnWidth = 16
      .Columns("AA").ColumnWidth = 12
      .Columns("AA").HorizontalAlignment = xlHAlignCenter
      .Columns("AB").ColumnWidth = 12
      .Columns("AB").HorizontalAlignment = xlHAlignCenter
      .Columns("AC").ColumnWidth = 30
      .Columns("AC").HorizontalAlignment = xlHAlignCenter
      .Columns("AD").ColumnWidth = 12
      .Columns("AD").HorizontalAlignment = xlHAlignCenter
      .Columns("AE").ColumnWidth = 50
      .Columns("AE").HorizontalAlignment = xlHAlignCenter
      .Columns("AF").ColumnWidth = 60
      .Columns("AF").HorizontalAlignment = xlHAlignCenter
      .Columns("AG").ColumnWidth = 60
      .Columns("AG").HorizontalAlignment = xlHAlignCenter
   End With
   
   g_rst_Princi.MoveFirst
   r_int_ConVer = 2
   
   Do While Not g_rst_Princi.EOF
      'Buscando datos de la Garantía en Registro de Hipotecas
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 1
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = Trim(g_rst_Princi!PRODUC_DESCRI)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = Trim(g_rst_Princi!SUB_PRODUCTO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = gf_Formato_NumOpe(g_rst_Princi!HIPMAE_NUMOPE)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = CStr(g_rst_Princi!HIPMAE_TDOCLI) & "-" & Trim(g_rst_Princi!HIPMAE_NDOCLI)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = Trim(g_rst_Princi!NOMBRE_CLIENTE)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECACT)))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECDES)))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = Trim(g_rst_Princi!ESTADO_CREDITO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = Trim(g_rst_Princi!Moneda)
      
      If g_rst_Princi!HIPMAE_MONEDA = 2 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = Format(g_rst_Princi!HIPMAE_CVTDOL, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = Format(g_rst_Princi!HIPMAE_APODOL, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = Format(g_rst_Princi!HIPMAE_APODOL / g_rst_Princi!HIPMAE_CVTDOL * 100, "##0.00") & "%"
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = Format(g_rst_Princi!HIPMAE_CVTSOL, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = Format(g_rst_Princi!HIPMAE_APOSOL, "###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = Format(g_rst_Princi!HIPMAE_APOSOL / g_rst_Princi!HIPMAE_CVTSOL * 100, "##0.00") & "%"
      End If
      
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14) = Format(g_rst_Princi!AP_PROPIO, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = Format(g_rst_Princi!MONTO_PBP, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 16) = Format(g_rst_Princi!MONTO_BBP, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 17) = Format(g_rst_Princi!MONTO_BMS, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 18) = Format(g_rst_Princi!MONTO_AFP, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 19) = Format(g_rst_Princi!HIPMAE_MTOPRE, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 20) = Format(g_rst_Princi!SOLMAE_MTOGCI, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 21) = Format(g_rst_Princi!HIPMAE_INTCAP, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 22) = Format(g_rst_Princi!HIPMAE_TOTPRE, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 23) = Format(g_rst_Princi!HIPMAE_IMPNCO, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 24) = Format(g_rst_Princi!HIPMAE_IMPCON, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 25) = g_rst_Princi!HIPMAE_TASINT
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 26) = g_rst_Princi!HIPMAE_COSEFE
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 27) = g_rst_Princi!HIPMAE_NUMCUO
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 28) = g_rst_Princi!HIPMAE_PERGRA
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 29) = Trim(g_rst_Princi!TIPO_GARANTIA)
      If g_rst_Princi!HIPMAE_PRYMCS = 1 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 30) = "SI"
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 30) = "NO"
      End If
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 31) = Trim(g_rst_Princi!PROYECTO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 32) = Trim(g_rst_Princi!PROMOTOR)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 33) = Trim(g_rst_Princi!CONSTRUCTOR)
      
      r_int_ConVer = r_int_ConVer + 1
      g_rst_Princi.MoveNext
      DoEvents
   Loop
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub ipp_FecFin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(cmd_Imprim)
    End If
End Sub

Private Sub ipp_FecIni_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call gs_SetFocus(ipp_FecFin)
    End If
End Sub
