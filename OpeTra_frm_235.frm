VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_Rpt_TraCof_03 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   3150
   ClientLeft      =   5475
   ClientTop       =   3945
   ClientWidth     =   6030
   Icon            =   "OpeTra_frm_235.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   3135
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   6015
      _Version        =   65536
      _ExtentX        =   10610
      _ExtentY        =   5530
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
         TabIndex        =   8
         Top             =   30
         Width           =   5955
         _Version        =   65536
         _ExtentX        =   10504
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
            TabIndex        =   9
            Top             =   30
            Width           =   4425
            _Version        =   65536
            _ExtentX        =   7805
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Reporte de Solicitudes en Trámites COFIDE"
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
            TabIndex        =   10
            Top             =   300
            Width           =   3345
            _Version        =   65536
            _ExtentX        =   5900
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Proyecto Inmobiliario"
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
            Picture         =   "OpeTra_frm_235.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   11
         Top             =   750
         Width           =   5955
         _Version        =   65536
         _ExtentX        =   10504
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
            Picture         =   "OpeTra_frm_235.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   5340
            Picture         =   "OpeTra_frm_235.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   630
            Picture         =   "OpeTra_frm_235.frx":0B9A
            Style           =   1  'Graphical
            TabIndex        =   5
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
      Begin Threed.SSPanel SSPanel3 
         Height          =   795
         Left            =   30
         TabIndex        =   12
         Top             =   2280
         Width           =   5955
         _Version        =   65536
         _ExtentX        =   10504
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
         Begin VB.ComboBox cmb_CodPry 
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   60
            Width           =   4575
         End
         Begin VB.CheckBox chk_CodPry 
            Caption         =   "Todos los Proyectos"
            Height          =   315
            Left            =   1320
            TabIndex        =   3
            Top             =   420
            Width           =   3555
         End
         Begin VB.Label Label1 
            Caption         =   "Proyecto:"
            Height          =   255
            Left            =   60
            TabIndex        =   13
            Top             =   60
            Width           =   765
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   795
         Left            =   30
         TabIndex        =   14
         Top             =   1440
         Width           =   5955
         _Version        =   65536
         _ExtentX        =   10504
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
         Begin VB.CheckBox chk_TipPry 
            Caption         =   "Todos los Tipos de Proyectos"
            Height          =   315
            Left            =   1320
            TabIndex        =   1
            Top             =   420
            Width           =   3555
         End
         Begin VB.ComboBox cmb_TipPry 
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   4575
         End
         Begin VB.Label Label2 
            Caption         =   "Tipo Proyecto:"
            Height          =   255
            Left            =   60
            TabIndex        =   15
            Top             =   60
            Width           =   1185
         End
      End
   End
End
Attribute VB_Name = "frm_Rpt_TraCof_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_CodPry()      As moddat_tpo_Genera

Private Sub chk_CodPry_Click()
   If chk_CodPry.Value = 1 Then
      cmb_CodPry.ListIndex = -1
      cmb_CodPry.Enabled = False
      
      Call gs_SetFocus(cmd_Imprim)
   ElseIf chk_CodPry.Value = 0 Then
      cmb_CodPry.Enabled = True
      chk_CodPry.Enabled = True
      
      Call gs_SetFocus(cmb_CodPry)
   End If
End Sub

Private Sub chk_TipPry_Click()
   If chk_TipPry.Value = 1 Then
      cmb_TipPry.ListIndex = -1
      cmb_TipPry.Enabled = False
      
      cmb_CodPry.ListIndex = -1
      
      cmb_CodPry.Enabled = False
      chk_CodPry.Enabled = False
      
      Call gs_SetFocus(cmd_Imprim)
   ElseIf chk_TipPry.Value = 0 Then
      cmb_TipPry.Enabled = True
      
      cmb_CodPry.Enabled = True
      chk_CodPry.Enabled = True
      
      Call gs_SetFocus(cmb_TipPry)
   End If
End Sub

Private Sub cmb_CodPry_Click()
   Call gs_SetFocus(cmd_Imprim)
End Sub

Private Sub cmb_CodPry_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_CodPry_Click
   End If
End Sub

Private Sub cmb_TipPry_Click()
   If cmb_TipPry.ListIndex > -1 Then
      If cmb_TipPry.ItemData(cmb_TipPry.ListIndex) = 1 Then
         cmb_CodPry.ListIndex = -1
         cmb_CodPry.Enabled = False
         
         chk_CodPry.Value = 0
         chk_CodPry.Enabled = False
         
         Call gs_SetFocus(cmd_Imprim)
      Else
         Screen.MousePointer = 11
         
         If chk_CodPry.Value = 0 Then
            cmb_CodPry.Enabled = True
         End If
         
         chk_CodPry.Enabled = True
         
         If cmb_TipPry.ItemData(cmb_TipPry.ListIndex) = 2 Then
            Call modmip_gs_Carga_PryInm_Combo(cmb_CodPry, l_arr_CodPry, 2)
         Else
            Call modmip_gs_Carga_PryInm_Combo(cmb_CodPry, l_arr_CodPry, 1)
         End If
         
         Screen.MousePointer = 0
         
         Call gs_SetFocus(cmb_CodPry)
      End If
   End If
End Sub

Private Sub cmb_TipPry_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipPry_Click
   End If
End Sub

Private Sub cmd_ExpExc_Click()
   Dim r_str_CodMod     As String
   Dim r_str_PrySel     As String

   If chk_TipPry.Value = 0 Then
      If cmb_TipPry.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Tipo de Proyecto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_TipPry)
         Exit Sub
      End If
   End If
   
   If chk_CodPry.Value = 0 And chk_CodPry.Enabled Then
      If cmb_CodPry.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Proyecto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_CodPry)
         Exit Sub
      End If
   End If

   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If

   Screen.MousePointer = 11
   
   r_str_PrySel = ""
   
   If chk_CodPry.Value = 0 And chk_CodPry.Enabled Then
      r_str_PrySel = l_arr_CodPry(cmb_CodPry.ListIndex + 1).Genera_Codigo
   End If
   
   r_str_CodMod = ""
   If chk_TipPry.Value = 0 Then
      r_str_CodMod = Format(cmb_TipPry.ItemData(cmb_TipPry.ListIndex), "00")
   End If
   
   Call modmip_gs_Exc_Tramit_Dbl(61, 62, 3, r_str_CodMod, r_str_PrySel)

   Screen.MousePointer = 0
End Sub

Private Sub cmd_Imprim_Click()
   If chk_TipPry.Value = 0 Then
      If cmb_TipPry.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Tipo de Proyecto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_TipPry)
         Exit Sub
      End If
   End If
   
   If chk_CodPry.Value = 0 And chk_CodPry.Enabled Then
      If cmb_CodPry.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Proyecto.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_CodPry)
         Exit Sub
      End If
   End If

   If MsgBox("¿Está seguro de imprimir el reporte?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If

   Screen.MousePointer = 11
      
   If chk_TipPry.Value = 0 Then
      If chk_CodPry.Value = 0 And chk_CodPry.Enabled = True Then
         Call modmip_gs_Rpt_EvaIns_Dbl("CRE_EVAHIP_03.RPT", 3, Format(cmb_TipPry.ItemData(cmb_TipPry.ListIndex), "00"), 61, 62, l_arr_CodPry(cmb_CodPry.ListIndex + 1).Genera_Codigo)
      Else
         Call modmip_gs_Rpt_EvaIns_Dbl("CRE_EVAHIP_03.RPT", 3, Format(cmb_TipPry.ItemData(cmb_TipPry.ListIndex), "00"), 61, 62, "")
      End If
   Else
      Call modmip_gs_Rpt_EvaIns_Dbl("CRE_EVAHIP_03.RPT", 1, "", 61, 62, "")
   End If

   Screen.MousePointer = 0

   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
   
   crp_Imprim.DataFiles(0) = "RPT_EVAHIP"
   
   crp_Imprim.SelectionFormula = "{RPT_EVAHIP.EVAHIP_NOMTER} = '" & modgen_g_str_NombPC & "' AND {RPT_EVAHIP.EVAHIP_NOMRPT} = 'CRE_EVAHIP_03.RPT'"
   
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "CRE_EVAHIP_03.RPT"
   
   crp_Imprim.Action = 1
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Me.Caption = modgen_g_str_NomPlt
   
   Call gs_CentraForm(Me)

   Call moddat_gs_Carga_LisIte_Combo(cmb_TipPry, 1, "272")
End Sub




