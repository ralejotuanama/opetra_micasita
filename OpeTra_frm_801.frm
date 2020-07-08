VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_Rpt_CreSal_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3930
   ClientLeft      =   5535
   ClientTop       =   4470
   ClientWidth     =   7170
   Icon            =   "OpeTra_frm_801.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   7170
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   3975
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   7185
      _Version        =   65536
      _ExtentX        =   12674
      _ExtentY        =   7011
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
         TabIndex        =   11
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
            TabIndex        =   12
            Top             =   45
            Width           =   4275
            _Version        =   65536
            _ExtentX        =   7541
            _ExtentY        =   1005
            _StockProps     =   15
            Caption         =   "Reporte de Saldos de Créditos Hipotecarios"
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
            Left            =   90
            Picture         =   "OpeTra_frm_801.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   13
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
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   630
            Picture         =   "OpeTra_frm_801.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_801.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   6480
            Picture         =   "OpeTra_frm_801.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   1230
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
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   2445
         Left            =   30
         TabIndex        =   14
         Top             =   1440
         Width           =   7095
         _Version        =   65536
         _ExtentX        =   12515
         _ExtentY        =   4313
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
         Begin VB.CheckBox Chk_FecAct 
            Caption         =   "A la Fecha"
            Height          =   285
            Left            =   1140
            TabIndex        =   6
            Top             =   2100
            Width           =   1995
         End
         Begin VB.ComboBox cmb_Permes 
            Height          =   315
            Left            =   1140
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   1350
            Width           =   5895
         End
         Begin VB.ComboBox cmb_Empres 
            Height          =   315
            Left            =   1140
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   5895
         End
         Begin VB.CheckBox chk_Empres 
            Caption         =   "Todas las Empresas"
            Height          =   285
            Left            =   1140
            TabIndex        =   1
            Top             =   420
            Width           =   1995
         End
         Begin VB.ComboBox cmb_TipPro 
            Height          =   315
            Left            =   1140
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   720
            Width           =   5895
         End
         Begin VB.CheckBox chk_TipPro 
            Caption         =   "Todos los Productos"
            Height          =   285
            Left            =   1140
            TabIndex        =   3
            Top             =   1050
            Width           =   1995
         End
         Begin EditLib.fpDoubleSingle ipp_PerAno 
            Height          =   315
            Left            =   1140
            TabIndex        =   5
            Top             =   1740
            Width           =   825
            _Version        =   196608
            _ExtentX        =   1455
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
            ButtonStyle     =   1
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
            Text            =   "0"
            DecimalPlaces   =   0
            DecimalPoint    =   "."
            FixedPoint      =   0   'False
            LeadZero        =   0
            MaxValue        =   "9999"
            MinValue        =   "1900"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ","
            UseSeparator    =   0   'False
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
         Begin VB.Label Label5 
            Caption         =   "Año:"
            Height          =   255
            Left            =   90
            TabIndex        =   18
            Top             =   1800
            Width           =   795
         End
         Begin VB.Label Label3 
            Caption         =   "Mes:"
            Height          =   255
            Left            =   60
            TabIndex        =   17
            Top             =   1410
            Width           =   795
         End
         Begin VB.Label Label4 
            Caption         =   "Empresa:"
            Height          =   255
            Left            =   60
            TabIndex        =   16
            Top             =   60
            Width           =   915
         End
         Begin VB.Label Label1 
            Caption         =   "Producto:"
            Height          =   255
            Left            =   60
            TabIndex        =   15
            Top             =   720
            Width           =   795
         End
      End
   End
End
Attribute VB_Name = "frm_Rpt_CreSal_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_Produc()      As moddat_tpo_Genera
Dim l_arr_Empres()      As moddat_tpo_Genera

Private Sub cmd_Imprim_Click()
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
   If Chk_FecAct.Value = 0 Then
      If cmb_PerMes.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Mes.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_PerMes)
         Exit Sub
      End If
      If ipp_PerAno.Text = 0 Then
         MsgBox "Debe seleccionar un Año.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_PerAno)
         Exit Sub
      End If
   End If
   
   If MsgBox("¿Está seguro de imprimir el Reporte?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If

   Screen.MousePointer = 11
   If cmb_PerMes.ListIndex = -1 Then
      crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
      crp_Imprim.DataFiles(0) = "CRE_HIPMAE"
      crp_Imprim.DataFiles(1) = "CLI_DATGEN"
      crp_Imprim.DataFiles(2) = "CRE_PRODUC"
      crp_Imprim.SelectionFormula = "{CRE_HIPMAE.HIPMAE_SITUAC} = 2 "
      
      If chk_TipPro.Value = 0 Then
         crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "AND {CRE_HIPMAE.HIPMAE_CODPRD} = '" & l_arr_Produc(cmb_TipPro.ListIndex + 1).Genera_Codigo & "'"
      End If
      
      crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "CTB_RPTSOL_02.RPT"
      crp_Imprim.WindowShowPrintSetupBtn = True
      crp_Imprim.Action = 1
   Else
      crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
      crp_Imprim.DataFiles(0) = "CLI_DATGEN"
      crp_Imprim.DataFiles(1) = "CRE_PRODUC"
      crp_Imprim.DataFiles(2) = "CRE_HIPCIE"
      crp_Imprim.SelectionFormula = "{CRE_HIPCIE.HIPCIE_PERMES} = " & cmb_PerMes.ItemData(cmb_PerMes.ListIndex) & " AND "
      crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{CRE_HIPCIE.HIPCIE_PERANO} = " & Format(ipp_PerAno.Text, "0000") & ""
      
      If chk_TipPro.Value = 0 Then
         crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "AND {CRE_HIPCIE.HIPCIE_CODPRD} = '" & l_arr_Produc(cmb_TipPro.ListIndex + 1).Genera_Codigo & "'"
      End If
      
      crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "CTB_RPTSOL_03.RPT"
      crp_Imprim.WindowShowPrintSetupBtn = True
      crp_Imprim.Action = 1
   End If
   Screen.MousePointer = 0
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
   If Chk_FecAct.Value = 0 Then
      If cmb_PerMes.ListIndex = -1 Then
         MsgBox "Debe seleccionar el Mes.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(cmb_PerMes)
         Exit Sub
      End If
      If ipp_PerAno.Text = 0 Then
         MsgBox "Debe seleccionar un Año.", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_PerAno)
         Exit Sub
      End If
   End If
   
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   cmd_ExpExc.Enabled = False
   If Chk_FecAct.Value = 0 Then
      Call fs_GenExc_Period
   Else
      Call fs_GenExc_FecAct
   End If
   cmd_ExpExc.Enabled = True
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Limpia
   Call gs_CentraForm(Me)
   
   Call gs_SetFocus(cmb_Empres)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_Produc(cmb_TipPro, l_arr_Produc, 4)
   Call moddat_gs_Carga_EmpGrp(cmb_Empres, l_arr_Empres)
   Call moddat_gs_Carga_LisIte_Combo(cmb_PerMes, 1, "033")
End Sub

Private Sub fs_Limpia()
   cmb_Empres.ListIndex = -1
   chk_Empres.Value = 0
   cmb_TipPro.ListIndex = -1
   chk_TipPro.Value = 0
   ipp_PerAno.Text = Year(date)
End Sub

Private Sub chk_Empres_Click()
   If chk_Empres.Value = 1 Then
      cmb_Empres.ListIndex = -1
      cmb_Empres.Enabled = False
      If cmb_TipPro.Enabled Then
         Call gs_SetFocus(cmb_TipPro)
      Else
         Call gs_SetFocus(cmb_PerMes)
      End If
   
   ElseIf chk_Empres.Value = 0 Then
      cmb_Empres.Enabled = True
      Call gs_SetFocus(cmb_Empres)
   End If
End Sub

Private Sub Chk_FecAct_Click()
   If Chk_FecAct.Value = 1 Then
      cmb_PerMes.ListIndex = -1
      cmb_PerMes.Enabled = False
      ipp_PerAno.Value = 0
      ipp_PerAno.Enabled = False
      Call gs_SetFocus(cmd_Imprim)
   ElseIf Chk_FecAct.Value = 0 Then
      cmb_PerMes.Enabled = True
      ipp_PerAno.Enabled = True
      Call gs_SetFocus(cmb_PerMes)
   End If
End Sub

Private Sub chk_TipPro_Click()
   If chk_TipPro.Value = 1 Then
      cmb_TipPro.ListIndex = -1
      cmb_TipPro.Enabled = False
      Call gs_SetFocus(cmb_PerMes)
   ElseIf chk_TipPro.Value = 0 Then
      cmb_TipPro.Enabled = True
      Call gs_SetFocus(cmb_TipPro)
   End If
End Sub

Private Sub cmb_TipPro_Click()
   Call gs_SetFocus(cmd_Imprim)
End Sub

Private Sub cmb_Empres_Click()
   If cmb_TipPro.Enabled Then
      Call gs_SetFocus(cmb_TipPro)
   Else
      Call gs_SetFocus(cmb_PerMes)
   End If
End Sub

Private Sub cmb_TipPro_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_PerMes)
   End If
End Sub

Private Sub cmb_PerMes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_PerAno)
   End If
End Sub

Private Function ff_Calcula_PBPPerdido(ByVal p_NumOpe As String) As Double
Dim r_rst_Genera        As ADODB.Recordset

   ff_Calcula_PBPPerdido = 0
   
   g_str_Parame = "SELECT SUM(HIPCUO_CAPBBP) - SUM(HIPCUO_CBPPAG) AS TOTAL FROM CRE_HIPCUO WHERE "
   g_str_Parame = g_str_Parame & "HIPCUO_NUMOPE = '" & p_NumOpe & "' AND "
   g_str_Parame = g_str_Parame & "HIPCUO_TIPCRO = 1 AND "
   g_str_Parame = g_str_Parame & "HIPCUO_CAPBBP > 0 AND "
   g_str_Parame = g_str_Parame & "HIPCUO_SITUAC = 2 "
   
   If Not gf_EjecutaSQL(g_str_Parame, r_rst_Genera, 3) Then
      Exit Function
   End If
   
   If r_rst_Genera.BOF And r_rst_Genera.EOF Then
      r_rst_Genera.Close
      Set r_rst_Genera = Nothing
      Exit Function
   End If

   r_rst_Genera.MoveFirst
   If IsNull(r_rst_Genera!Total) Then
      ff_Calcula_PBPPerdido = 0
   Else
      ff_Calcula_PBPPerdido = r_rst_Genera!Total
   End If
   
   r_rst_Genera.Close
   Set r_rst_Genera = Nothing
End Function

Private Sub fs_GenExc_FecAct()
Dim r_obj_Excel      As Excel.Application
Dim r_int_ConVer     As Integer
Dim r_dbl_PBPPer     As Double

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT HIPMAE_CODPRD,HIPMAE_NUMOPE,HIPMAE_TDOCLI,HIPMAE_NDOCLI,"
   g_str_Parame = g_str_Parame & "       DATGEN_APEPAT,DATGEN_APEMAT,DATGEN_NOMBRE,"
   g_str_Parame = g_str_Parame & "       DATGEN_NACFEC,HIPMAE_FECDES,HIPMAE_MTOPRE,D.PARDES_DESCRI MONEDA,"
   g_str_Parame = g_str_Parame & "       HIPMAE_INTCAP,HIPMAE_TOTPRE,HIPMAE_TASINT,HIPMAE_PLAANO,"
   g_str_Parame = g_str_Parame & "       HIPMAE_SALCAP,HIPMAE_SALCON,HIPMAE_PRXVCT,"
   g_str_Parame = g_str_Parame & "       HIPMAE_UlTVCT,HIPMAE_ULTPAG,HIPMAE_VCTANT,HIPMAE_DIAMOR,"
   g_str_Parame = g_str_Parame & "       HIPMAE_MONGAR,HIPMAE_MTOGAR,HIPMAE_TIPGAR,HIPMAE_MONEDA,"
   g_str_Parame = g_str_Parame & "       HIPMAE_ACUDIF,PRODUC_DESCRI,E.PARDES_DESCRI GARANTIA,"
   g_str_Parame = g_str_Parame & "       (SELECT TRIM(PARDES_DESCRI) FROM MNT_PARDES WHERE PARDES_CODGRP = 509 AND PARDES_CODITE = F.EVALEG_CODNOT) AS NOTARIA "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPMAE A "
   g_str_Parame = g_str_Parame & "        INNER JOIN CLI_DATGEN B ON A.HIPMAE_TDOCLI = B.DATGEN_TIPDOC AND A.HIPMAE_NDOCLI = B.DATGEN_NUMDOC"
   g_str_Parame = g_str_Parame & "        INNER JOIN CRE_PRODUC C ON A.HIPMAE_CODPRD = C.PRODUC_CODIGO"
   g_str_Parame = g_str_Parame & "        INNER JOIN MNT_PARDES D ON D.PARDES_CODGRP = 204 AND D.PARDES_CODITE = A.HIPMAE_MONEDA"
   g_str_Parame = g_str_Parame & "        INNER JOIN MNT_PARDES E ON E.PARDES_CODGRP = 241 AND E.PARDES_CODITE = A.HIPMAE_TIPGAR"
   g_str_Parame = g_str_Parame & "        INNER JOIN TRA_EVALEG F ON F.EVALEG_NUMSOL = A.HIPMAE_NUMSOL "
   g_str_Parame = g_str_Parame & " WHERE HIPMAE_SITUAC = 2 "
   If chk_TipPro.Value = 0 Then
      g_str_Parame = g_str_Parame & " AND HIPMAE_CODPRD = '" & l_arr_Produc(cmb_TipPro.ListIndex + 1).Genera_Codigo & "' "
   End If
   g_str_Parame = g_str_Parame & "ORDER BY HIPMAE_CODPRD ASC, HIPMAE_MONEDA ASC, DATGEN_APEPAT ASC, DATGEN_APEMAT ASC, DATGEN_NOMBRE ASC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Cells(1, 1) = "ITEM"
      .Cells(1, 2) = "PRODUCTO"
      .Cells(1, 3) = "OPERACION"
      .Cells(1, 4) = "DOC. IDENTIDAD"
      .Cells(1, 5) = "NOMBRE CLIENTE"
      .Cells(1, 6) = "FEC. NACIMIENTO"
      .Cells(1, 7) = "F. DESEMBOLSO"
      .Cells(1, 8) = "MONEDA"
      .Cells(1, 9) = "MTO. PRESTAMO"
      .Cells(1, 10) = "INT. CAPIT."
      .Cells(1, 11) = "TOTAL PRESTAMO"
      .Cells(1, 12) = "T. INTERES"
      .Cells(1, 13) = "PLAZO"
      .Cells(1, 14) = "SALDO CAPITAL"
      .Cells(1, 15) = "SALDO TC"
      .Cells(1, 16) = "SALDO PBP"
      .Cells(1, 17) = "TOTAL SALDO"
      .Cells(1, 18) = "F. PROX. VCTO."
      .Cells(1, 19) = "F. ULT. VCTO."
      .Cells(1, 20) = "F. ULT. PAGO"
      .Cells(1, 21) = "F. VCTO ANT."
      .Cells(1, 22) = "DIA ATR."
      .Cells(1, 23) = "TIPO GARANTIA"
      .Cells(1, 24) = "GARANTIA S/."
      .Cells(1, 25) = "GARANTIA US$"
      .Cells(1, 26) = "INT. DIFERIDO"
      .Cells(1, 27) = "PROYECTO MI CASITA"
      .Cells(1, 28) = "NOMBRE DEL PROYECTO"
      .Cells(1, 29) = "DIRECCIÓN"
      .Cells(1, 30) = "NOTARIA"
      
      .Range(.Cells(1, 1), .Cells(1, 30)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 30)).HorizontalAlignment = xlHAlignCenter
      .Columns("A").ColumnWidth = 8
      .Columns("B").ColumnWidth = 30
      .Columns("C").ColumnWidth = 15
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 15
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").ColumnWidth = 40
      .Columns("F").ColumnWidth = 16
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("G").ColumnWidth = 20
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      .Columns("H").ColumnWidth = 30
      .Columns("I").ColumnWidth = 20
      .Columns("J").ColumnWidth = 20
      .Columns("K").ColumnWidth = 20
      .Columns("L").ColumnWidth = 20
      .Columns("M").ColumnWidth = 20
      .Columns("N").ColumnWidth = 20
      .Columns("O").ColumnWidth = 20
      .Columns("P").ColumnWidth = 20
      .Columns("Q").ColumnWidth = 20
      .Columns("R").ColumnWidth = 20
      .Columns("R").HorizontalAlignment = xlHAlignCenter
      .Columns("S").ColumnWidth = 20
      .Columns("S").HorizontalAlignment = xlHAlignCenter
      .Columns("T").ColumnWidth = 20
      .Columns("T").HorizontalAlignment = xlHAlignCenter
      .Columns("U").ColumnWidth = 20
      .Columns("U").HorizontalAlignment = xlHAlignCenter
      .Columns("V").ColumnWidth = 15
      .Columns("V").HorizontalAlignment = xlHAlignCenter
      .Columns("W").ColumnWidth = 30
      .Columns("W").HorizontalAlignment = xlHAlignCenter
      .Columns("X").ColumnWidth = 20
      .Columns("Y").ColumnWidth = 20
      .Columns("Z").ColumnWidth = 17
      .Columns("AA").HorizontalAlignment = xlHAlignCenter
      .Columns("AA").ColumnWidth = 20
      .Columns("AB").ColumnWidth = 40
      .Columns("AC").ColumnWidth = 120
      .Columns("AD").ColumnWidth = 50
   End With
   
   g_rst_Princi.MoveFirst
   r_int_ConVer = 2
   
   Do While Not g_rst_Princi.EOF
      'Buscando datos de la Garantía en Registro de Hipotecas
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 1
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = Trim(g_rst_Princi!PRODUC_DESCRI)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = gf_Formato_NumOpe(g_rst_Princi!HIPMAE_NUMOPE)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = CStr(g_rst_Princi!HIPMAE_TDOCLI) & "-" & Trim(g_rst_Princi!HIPMAE_NDOCLI)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = Trim(g_rst_Princi!DatGen_ApePat) & " " & Trim(g_rst_Princi!DatGen_ApeMat) & " " & Trim(g_rst_Princi!DatGen_Nombre)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!DATGEN_NACFEC)))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_FECDES)))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = Trim(g_rst_Princi!MONEDA)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = Format(g_rst_Princi!HIPMAE_MTOPRE, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = Format(g_rst_Princi!HIPMAE_INTCAP, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = Format(g_rst_Princi!HIPMAE_TOTPRE, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = Format(g_rst_Princi!HIPMAE_TASINT, "##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = CStr(g_rst_Princi!HIPMAE_PLAANO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14) = Format(g_rst_Princi!HIPMAE_SALCAP, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = Format(g_rst_Princi!HIPMAE_SALCON, "###,###,##0.00")
      r_dbl_PBPPer = ff_Calcula_PBPPerdido(g_rst_Princi!HIPMAE_NUMOPE)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 16) = Format(r_dbl_PBPPer, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 17) = Format(g_rst_Princi!HIPMAE_SALCAP + g_rst_Princi!HIPMAE_SALCON, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 18) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_PRXVCT)))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 19) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_UlTVCT)))
      If g_rst_Princi!HIPMAE_ULTPAG > 0 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 20) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_ULTPAG)))
      End If
      If g_rst_Princi!HIPMAE_VCTANT > 0 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 21) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!HIPMAE_VCTANT)))
      End If
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 22) = g_rst_Princi!HIPMAE_DIAMOR
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 23) = Trim(g_rst_Princi!GARANTIA)
      If g_rst_Princi!HIPMAE_MONGAR = 1 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 24) = Format(g_rst_Princi!HIPMAE_MTOGAR, "###,###,##0.00")
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 25) = Format(g_rst_Princi!HIPMAE_MTOGAR, "###,###,##0.00")
      End If
      If g_rst_Princi!HIPMAE_TIPGAR = 3 Or g_rst_Princi!HIPMAE_TIPGAR = 6 Then
         If g_rst_Princi!HIPMAE_MONEDA = 1 Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 24) = Format(g_rst_Princi!HIPMAE_SALCAP + g_rst_Princi!HIPMAE_SALCON, "###,###,##0.00")
         Else
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 25) = Format(g_rst_Princi!HIPMAE_SALCAP + g_rst_Princi!HIPMAE_SALCON, "###,###,##0.00")
         End If
      End If
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 26) = Format(g_rst_Princi!HIPMAE_ACUDIF, "###,###,##0.00")
      
      'OBTENER LOS DATOS DEL PROYECTO MI CASITA
      g_str_Parame = "SELECT * FROM CRE_SOLINM "
      g_str_Parame = g_str_Parame & "JOIN CRE_HIPMAE ON (HIPMAE_NUMSOL = SOLINM_NUMSOL) "
      g_str_Parame = g_str_Parame & "WHERE HIPMAE_NUMOPE = '" & g_rst_Princi!HIPMAE_NUMOPE & "' "
       
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
         Exit Sub
      End If
      
      If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 27) = moddat_gf_Consulta_ParDes("214", g_rst_GenAux!SOLINM_PRYMCS)
         
         If g_rst_GenAux!SOLINM_TABPRY = 2 Then
            If Len(Trim(g_rst_GenAux!SOLINM_PRYCOD)) > 0 Then
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 28) = moddat_gf_Consulta_NomPry(g_rst_GenAux!SOLINM_PRYCOD)
            Else
               If Len(Trim(g_rst_GenAux!SOLINM_PRYNOM)) > 0 Then
                  r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 28) = Trim(g_rst_GenAux!SOLINM_PRYNOM & "")
               End If
            End If
         Else
            If Len(Trim(g_rst_GenAux!SOLINM_PRYCOD & "")) > 0 Then
                r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 28) = moddat_gf_Consulta_NomPry(g_rst_GenAux!SOLINM_PRYCOD)
                
            End If
         End If
      
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 29) = moddat_gf_Consulta_ParDes("201", CStr(g_rst_GenAux!SOLINM_TIPVIA)) & _
                        " " & Trim(g_rst_GenAux!SOLINM_NOMVIA) & " " & Trim(g_rst_GenAux!SOLINM_NUMVIA) & _
                        IIf(Len(Trim(g_rst_GenAux!SOLINM_INTDPT)) > 0, " (" & Trim(g_rst_GenAux!SOLINM_INTDPT) & ")", "") & _
                        IIf(Len(Trim(g_rst_GenAux!SOLINM_NOMZON)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_GenAux!SOLINM_TIPZON)) & " " & Trim(g_rst_GenAux!SOLINM_NOMZON), "")
      End If
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 30) = Trim(g_rst_Princi!NOTARIA)
       
      r_int_ConVer = r_int_ConVer + 1
      g_rst_Princi.MoveNext
      DoEvents
   Loop
   
   g_rst_GenAux.Close
   Set g_rst_GenAux = Nothing
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Screen.MousePointer = 0
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub fs_GenExc_Period()
Dim r_obj_Excel      As Excel.Application
Dim r_int_ConVer     As Integer
Dim r_str_NumSol     As String
Dim r_int_TipMon     As Integer
Dim r_str_FecCam     As String
Dim r_str_TipCam     As String
Dim r_str_CheCgo     As String
Dim r_str_CtaCgo     As String
Dim r_str_BanCgo     As String

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT HIPCIE_CODPRD,HIPCIE_NUMOPE,HIPCIE_TDOCLI,HIPCIE_NDOCLI,"
   g_str_Parame = g_str_Parame & "       DATGEN_APEPAT,DATGEN_APEMAT,DATGEN_NOMBRE,"
   g_str_Parame = g_str_Parame & "       DATGEN_NACFEC,HIPCIE_FECDES,HIPCIE_MTOPRE,D.PARDES_DESCRI MONEDA,"
   g_str_Parame = g_str_Parame & "       HIPCIE_INTCAP,HIPCIE_TOTPRE,HIPCIE_TASINT,HIPCIE_PLAANO,"
   g_str_Parame = g_str_Parame & "       HIPCIE_SALCAP,HIPCIE_SALCON,HIPCIE_PERPBP,HIPCIE_PRXVCT,"
   g_str_Parame = g_str_Parame & "       HIPCIE_UlTVCT,HIPCIE_ULTPAG,HIPCIE_VCTANT,HIPCIE_DIAMOR,"
   g_str_Parame = g_str_Parame & "       HIPCIE_MONGAR,HIPCIE_MTOGAR,HIPCIE_TIPGAR,HIPCIE_TIPMON,"
   g_str_Parame = g_str_Parame & "       HIPCIE_INTDIF,PRODUC_DESCRI,E.PARDES_DESCRI GARANTIA,"
   g_str_Parame = g_str_Parame & "       (SELECT TRIM(PARDES_DESCRI) FROM MNT_PARDES WHERE PARDES_CODGRP = 509 AND PARDES_CODITE = G.EVALEG_CODNOT) AS NOTARIA "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPCIE A "
   g_str_Parame = g_str_Parame & "        INNER JOIN CLI_DATGEN B ON A.HIPCIE_TDOCLI = B.DATGEN_TIPDOC AND A.HIPCIE_NDOCLI = B.DATGEN_NUMDOC"
   g_str_Parame = g_str_Parame & "        INNER JOIN CRE_PRODUC C ON A.HIPCIE_CODPRD = C.PRODUC_CODIGO"
   g_str_Parame = g_str_Parame & "        INNER JOIN MNT_PARDES D ON D.PARDES_CODGRP = 204 AND D.PARDES_CODITE = A.HIPCIE_TIPMON"
   g_str_Parame = g_str_Parame & "        INNER JOIN MNT_PARDES E ON E.PARDES_CODGRP = 241 AND E.PARDES_CODITE = A.HIPCIE_TIPGAR"
   g_str_Parame = g_str_Parame & "        INNER JOIN CRE_HIPMAE F ON A.HIPCIE_NUMOPE = F.HIPMAE_NUMOPE"
   g_str_Parame = g_str_Parame & "        INNER JOIN TRA_EVALEG G ON G.EVALEG_NUMSOL = F.HIPMAE_NUMSOL "
   g_str_Parame = g_str_Parame & " WHERE HIPCIE_PERMES = " & cmb_PerMes.ItemData(cmb_PerMes.ListIndex) & ""
   g_str_Parame = g_str_Parame & "   AND HIPCIE_PERANO = " & Format(ipp_PerAno.Text, "0000") & " "
   If chk_TipPro.Value = 0 Then
      g_str_Parame = g_str_Parame & "AND HIPCIE_CODPRD = '" & l_arr_Produc(cmb_TipPro.ListIndex + 1).Genera_Codigo & "' "
   End If
   g_str_Parame = g_str_Parame & "ORDER BY HIPCIE_CODPRD ASC, HIPCIE_TIPMON ASC, DATGEN_APEPAT ASC, DATGEN_APEMAT ASC, DATGEN_NOMBRE ASC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Cells(1, 1) = "ITEM"
      .Cells(1, 2) = "PRODUCTO"
      .Cells(1, 3) = "OPERACION"
      .Cells(1, 4) = "DOC. IDENTIDAD"
      .Cells(1, 5) = "NOMBRE CLIENTE"
      .Cells(1, 6) = "FEC. NACIMIENTO"
      .Cells(1, 7) = "F. DESEMBOLSO"
      .Cells(1, 8) = "MONEDA"
      .Cells(1, 9) = "MTO. PRESTAMO"
      .Cells(1, 10) = "INT. CAPIT."
      .Cells(1, 11) = "TOTAL PRESTAMO"
      .Cells(1, 12) = "T. INTERES"
      .Cells(1, 13) = "PLAZO"
      .Cells(1, 14) = "SALDO CAPITAL"
      .Cells(1, 15) = "SALDO TC"
      .Cells(1, 16) = "SALDO PBP"
      .Cells(1, 17) = "TOTAL SALDO"
      .Cells(1, 18) = "F. PROX. VCTO."
      .Cells(1, 19) = "F. ULT. VCTO."
      .Cells(1, 20) = "F. ULT. PAGO"
      .Cells(1, 21) = "F. VCTO ANT."
      .Cells(1, 22) = "DIA ATR."
      .Cells(1, 23) = "TIPO GARANTIA"
      .Cells(1, 24) = "GARANTIA S/."
      .Cells(1, 25) = "GARANTIA US$."
      .Cells(1, 26) = "INT. DIFERIDO"
      .Cells(1, 27) = "PROYECTO MI CASITA"
      .Cells(1, 28) = "NOMBRE DEL PROYECTO"
      .Cells(1, 29) = "DIRECCIÓN"
      .Cells(1, 30) = "NOTARIA"
       
      .Range(.Cells(1, 1), .Cells(1, 30)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 30)).HorizontalAlignment = xlHAlignCenter
      .Columns("A").ColumnWidth = 8
      .Columns("B").ColumnWidth = 30
      .Columns("C").ColumnWidth = 15
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 15
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").ColumnWidth = 40
      .Columns("F").ColumnWidth = 16
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("G").ColumnWidth = 20
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      .Columns("H").ColumnWidth = 30
      .Columns("I").ColumnWidth = 20
      .Columns("J").ColumnWidth = 20
      .Columns("K").ColumnWidth = 20
      .Columns("L").ColumnWidth = 20
      .Columns("M").ColumnWidth = 20
      .Columns("N").ColumnWidth = 20
      .Columns("O").ColumnWidth = 20
      .Columns("P").ColumnWidth = 20
      .Columns("Q").ColumnWidth = 20
      .Columns("R").ColumnWidth = 20
      .Columns("R").HorizontalAlignment = xlHAlignCenter
      .Columns("S").ColumnWidth = 20
      .Columns("S").HorizontalAlignment = xlHAlignCenter
      .Columns("T").ColumnWidth = 20
      .Columns("T").HorizontalAlignment = xlHAlignCenter
      .Columns("U").ColumnWidth = 20
      .Columns("U").HorizontalAlignment = xlHAlignCenter
      .Columns("V").ColumnWidth = 15
      .Columns("V").HorizontalAlignment = xlHAlignCenter
      .Columns("W").ColumnWidth = 30
      .Columns("W").HorizontalAlignment = xlHAlignCenter
      .Columns("X").ColumnWidth = 20
      .Columns("Y").ColumnWidth = 20
      .Columns("Z").ColumnWidth = 17
      .Columns("AA").HorizontalAlignment = xlHAlignCenter
      .Columns("AA").ColumnWidth = 20
      .Columns("AB").ColumnWidth = 40
      .Columns("AC").ColumnWidth = 120
      .Columns("AD").ColumnWidth = 50
   End With
   
   g_rst_Princi.MoveFirst
   r_int_ConVer = 2
   
   Do While Not g_rst_Princi.EOF
      'Buscando datos de la Garantía en Registro de Hipotecas
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 1
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = Trim(g_rst_Princi!PRODUC_DESCRI)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = gf_Formato_NumOpe(g_rst_Princi!HIPCIE_NUMOPE)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = CStr(g_rst_Princi!HIPCIE_TDOCLI) & "-" & Trim(g_rst_Princi!HIPCIE_NDOCLI)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = Trim(g_rst_Princi!DatGen_ApePat) & " " & Trim(g_rst_Princi!DatGen_ApeMat) & " " & Trim(g_rst_Princi!DatGen_Nombre)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!DATGEN_NACFEC)))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!HIPCIE_FECDES)))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = Trim(g_rst_Princi!MONEDA)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = Format(g_rst_Princi!HIPCIE_MTOPRE, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = Format(g_rst_Princi!HIPCIE_INTCAP, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = Format(g_rst_Princi!HIPCIE_TOTPRE, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = Format(g_rst_Princi!HIPCIE_TASINT, "##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = CStr(g_rst_Princi!HIPCIE_PLAANO)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14) = Format(g_rst_Princi!HIPCIE_SALCAP, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = Format(g_rst_Princi!HIPCIE_SALCON, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 16) = Format(g_rst_Princi!HIPCIE_PERPBP, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 17) = Format(g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON, "###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 18) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!HIPCIE_PRXVCT)))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 19) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!HIPCIE_UlTVCT)))
      If g_rst_Princi!HIPCIE_ULTPAG > 0 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 20) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!HIPCIE_ULTPAG)))
      End If
      If g_rst_Princi!HIPCIE_VCTANT > 0 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 21) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!HIPCIE_VCTANT)))
      End If
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 22) = g_rst_Princi!HIPCIE_DIAMOR
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 23) = Trim(g_rst_Princi!GARANTIA)
      If g_rst_Princi!HIPCIE_MONGAR = 1 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 24) = Format(g_rst_Princi!HIPCIE_MTOGAR, "###,###,##0.00")
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 25) = Format(g_rst_Princi!HIPCIE_MTOGAR, "###,###,##0.00")
      End If
      
      If g_rst_Princi!HIPCIE_TIPGAR = 3 Or g_rst_Princi!HIPCIE_TIPGAR = 6 Then
         If g_rst_Princi!HIPCIE_TIPMON = 1 Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 24) = Format(g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON, "###,###,##0.00")
         Else
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 25) = Format(g_rst_Princi!HIPCIE_SALCAP + g_rst_Princi!HIPCIE_SALCON, "###,###,##0.00")
         End If
      End If
      
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 26) = Format(g_rst_Princi!HIPCIE_INTDIF, "###,###,##0.00")
      
      'OBTENER LOS DATOS DEL PROYECTO MI CASITA
      g_str_Parame = "SELECT * FROM CRE_SOLINM "
      g_str_Parame = g_str_Parame & "JOIN CRE_HIPMAE ON (HIPMAE_NUMSOL =SOLINM_NUMSOL) "
      g_str_Parame = g_str_Parame & "WHERE HIPMAE_NUMOPE ='" & g_rst_Princi!HIPCIE_NUMOPE & "' "
       
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
         Exit Sub
      End If
      
      If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 27) = moddat_gf_Consulta_ParDes("214", g_rst_GenAux!SOLINM_PRYMCS)
         
         If g_rst_GenAux!SOLINM_TABPRY = 2 Then
            If Len(Trim(g_rst_GenAux!SOLINM_PRYCOD)) > 0 Then
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 28) = moddat_gf_Consulta_NomPry(g_rst_GenAux!SOLINM_PRYCOD)
            Else
               If Len(Trim(g_rst_GenAux!SOLINM_PRYNOM)) > 0 Then
                  r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 28) = Trim(g_rst_GenAux!SOLINM_PRYNOM & "")
               End If
            End If
         Else
            If Len(Trim(g_rst_GenAux!SOLINM_PRYCOD & "")) > 0 Then
                r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 28) = moddat_gf_Consulta_NomPry(g_rst_GenAux!SOLINM_PRYCOD)
            End If
         End If
      
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 29) = moddat_gf_Consulta_ParDes("201", CStr(g_rst_GenAux!SOLINM_TIPVIA)) & _
                        " " & Trim(g_rst_GenAux!SOLINM_NOMVIA) & " " & Trim(g_rst_GenAux!SOLINM_NUMVIA) & _
                        IIf(Len(Trim(g_rst_GenAux!SOLINM_INTDPT)) > 0, " (" & Trim(g_rst_GenAux!SOLINM_INTDPT) & ")", "") & _
                        IIf(Len(Trim(g_rst_GenAux!SOLINM_NOMZON)) > 0, " - " & moddat_gf_Consulta_ParDes("202", CStr(g_rst_GenAux!SOLINM_TIPZON)) & " " & Trim(g_rst_GenAux!SOLINM_NOMZON), "")
      End If
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 30) = Trim(g_rst_Princi!NOTARIA)
      
      r_int_ConVer = r_int_ConVer + 1
      g_rst_Princi.MoveNext
      DoEvents
   Loop
   
   g_rst_GenAux.Close
   Set g_rst_GenAux = Nothing

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Screen.MousePointer = 0
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub
