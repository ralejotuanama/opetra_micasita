VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_Rpt_EvaTas_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   2310
   ClientLeft      =   1470
   ClientTop       =   4275
   ClientWidth     =   5460
   Icon            =   "OpeTra_frm_226.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   2295
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   5445
      _Version        =   65536
      _ExtentX        =   9604
      _ExtentY        =   4048
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
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   6
         Top             =   30
         Width           =   5355
         _Version        =   65536
         _ExtentX        =   9446
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
            Height          =   285
            Left            =   660
            TabIndex        =   7
            Top             =   30
            Width           =   4425
            _Version        =   65536
            _ExtentX        =   7805
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Reporte de Solicitudes en Tasaci�n del Inmueble"
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
         Begin Threed.SSPanel SSPanel5 
            Height          =   285
            Left            =   660
            TabIndex        =   8
            Top             =   300
            Width           =   3345
            _Version        =   65536
            _ExtentX        =   5900
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Consejero Hipotecario"
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
            Picture         =   "OpeTra_frm_226.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   9
         Top             =   750
         Width           =   5355
         _Version        =   65536
         _ExtentX        =   9446
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
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_226.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   4740
            Picture         =   "OpeTra_frm_226.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   630
            Picture         =   "OpeTra_frm_226.frx":0B9A
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   3540
            Top             =   90
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            WindowTitle     =   "Presentaci�n Preliminar"
            WindowControlBox=   -1  'True
            WindowMaxButton =   -1  'True
            WindowMinButton =   -1  'True
            WindowState     =   2
            PrintFileLinesPerPage=   60
            WindowShowPrintSetupBtn=   -1  'True
            WindowShowRefreshBtn=   -1  'True
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   795
         Left            =   30
         TabIndex        =   10
         Top             =   1440
         Width           =   5355
         _Version        =   65536
         _ExtentX        =   9446
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
         Begin VB.ComboBox cmb_ConHip 
            Height          =   315
            Left            =   1020
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   4275
         End
         Begin VB.CheckBox chk_ConHip 
            Caption         =   "Todos los Consejeros Hipotecarios"
            Height          =   315
            Left            =   1020
            TabIndex        =   1
            Top             =   420
            Width           =   3555
         End
         Begin VB.Label Label1 
            Caption         =   "Consejero:"
            Height          =   255
            Left            =   60
            TabIndex        =   11
            Top             =   60
            Width           =   765
         End
      End
   End
End
Attribute VB_Name = "frm_Rpt_EvaTas_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_ConHip()      As moddat_tpo_Genera

Private Sub chk_ConHip_Click()
   If chk_ConHip.Value = 1 Then
      cmb_ConHip.ListIndex = -1
      cmb_ConHip.Enabled = False
   ElseIf chk_ConHip.Value = 0 Then
      cmb_ConHip.Enabled = True
      
      Call gs_SetFocus(cmb_ConHip)
   End If
End Sub

Private Sub cmb_ConHip_Click()
   Call gs_SetFocus(cmd_Imprim)
End Sub

Private Sub cmb_ConHip_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_ConHip_Click
   End If
End Sub

Private Sub cmd_ExpExc_Click()
   Dim r_obj_Excel      As excel.Application
   Dim r_int_ConVer     As Integer
   
   Dim r_rst_SolMae     As ADODB.Recordset
   Dim r_rst_Seguim     As ADODB.Recordset
   Dim r_rst_SegDet     As ADODB.Recordset
   
   Dim r_str_Cadena     As String
   Dim r_str_DetObs     As String
   Dim r_int_TpoObs     As Integer
   Dim r_int_TpoIns     As Integer
   
   If chk_ConHip.Value = 0 Then
      If cmb_ConHip.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Consejero Hipotecario.", vbExclamation, modgen_g_str_NomPlt
         
         Call gs_SetFocus(cmb_ConHip)
         Exit Sub
      End If
   End If
   
   If MsgBox("�Est� seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If

   Screen.MousePointer = 11
   
   If chk_ConHip.Value = 0 Then
      Call modmip_gs_Exc_Tramit_Dbl(41, 41, 2, l_arr_ConHip(cmb_ConHip.ListIndex + 1).Genera_Codigo, "")
   Else
      Call modmip_gs_Exc_Tramit_Dbl(41, 41, 2, "", "")
   End If

   Screen.MousePointer = 0
End Sub

Private Sub cmd_Imprim_Click()
   Dim r_rst_SolMae     As ADODB.Recordset
   Dim r_rst_Seguim     As ADODB.Recordset
   Dim r_rst_SegDet     As ADODB.Recordset
   
   Dim r_str_Cadena     As String
   Dim r_str_DetObs     As String
   Dim r_int_TpoObs     As Integer
   Dim r_int_TpoIns     As Integer

   If chk_ConHip.Value = 0 Then
      If cmb_ConHip.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Consejero Hipotecario.", vbExclamation, modgen_g_str_NomPlt
         
         Call gs_SetFocus(cmb_ConHip)
         Exit Sub
      End If
   End If
   
   If MsgBox("�Est� seguro de imprimir el reporte?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
      
   If chk_ConHip.Value = 0 Then
      Call modmip_gs_Rpt_EvaIns_Dbl("CRE_EVAHIP_02.RPT", 2, l_arr_ConHip(cmb_ConHip.ListIndex + 1).Genera_Codigo, 41, 41, "")
   Else
      Call modmip_gs_Rpt_EvaIns_Dbl("CRE_EVAHIP_02.RPT", 2, "", 41, 41, "")
   End If

   Screen.MousePointer = 0

   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
   
   crp_Imprim.DataFiles(0) = "RPT_EVAHIP"
   
   crp_Imprim.SelectionFormula = "{RPT_EVAHIP.EVAHIP_NOMTER} = '" & modgen_g_str_NombPC & "' AND {RPT_EVAHIP.EVAHIP_NOMRPT} = 'CRE_EVAHIP_02.RPT'"
   
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "CRE_EVAHIP_02.RPT"
   
   crp_Imprim.Action = 1
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Me.Caption = modgen_g_str_NomPlt
   
   Call gs_CentraForm(Me)
   
   Call moddat_gs_Carga_EjecMC(cmb_ConHip, l_arr_ConHip, 121)
End Sub



