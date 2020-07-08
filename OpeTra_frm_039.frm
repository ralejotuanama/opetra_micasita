VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_RptFia_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form3"
   ClientHeight    =   3210
   ClientLeft      =   3540
   ClientTop       =   3390
   ClientWidth     =   8295
   Icon            =   "OpeTra_frm_039.frx":0000
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   8295
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   3200
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8295
      _Version        =   65536
      _ExtentX        =   14631
      _ExtentY        =   5644
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
      Begin Threed.SSPanel SSPanel3 
         Height          =   1665
         Left            =   30
         TabIndex        =   3
         Top             =   1470
         Width           =   8205
         _Version        =   65536
         _ExtentX        =   14473
         _ExtentY        =   2937
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
         Begin VB.ComboBox cmb_Permes 
            Height          =   315
            Left            =   1500
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   540
            Width           =   6495
         End
         Begin VB.CheckBox Chk_FecAct 
            Caption         =   "A la Fecha"
            Height          =   285
            Left            =   1500
            TabIndex        =   10
            Top             =   1275
            Width           =   1995
         End
         Begin VB.ComboBox cmb_TipRep 
            Height          =   315
            Left            =   1500
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   210
            Width           =   6495
         End
         Begin EditLib.fpDoubleSingle ipp_PerAno 
            Height          =   315
            Left            =   1500
            TabIndex        =   11
            Top             =   915
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
         Begin VB.Label Label3 
            Caption         =   "Mes:"
            Height          =   255
            Left            =   720
            TabIndex        =   14
            Top             =   570
            Width           =   435
         End
         Begin VB.Label Label5 
            Caption         =   "Año:"
            Height          =   255
            Left            =   750
            TabIndex        =   12
            Top             =   945
            Width           =   435
         End
         Begin VB.Label Label6 
            Caption         =   "Tipo Reporte:"
            Height          =   315
            Left            =   90
            TabIndex        =   5
            Top             =   210
            Width           =   1395
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   680
         Left            =   30
         TabIndex        =   6
         Top             =   750
         Width           =   8205
         _Version        =   65536
         _ExtentX        =   14473
         _ExtentY        =   1199
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
            Height          =   615
            Left            =   650
            Picture         =   "OpeTra_frm_039.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   615
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   615
            Left            =   30
            Picture         =   "OpeTra_frm_039.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   615
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   615
            Left            =   7590
            Picture         =   "OpeTra_frm_039.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   615
         End
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   5400
            Top             =   0
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            WindowTitle     =   "Presentación Preliminar"
            WindowControlBox=   -1  'True
            WindowMaxButton =   -1  'True
            WindowMinButton =   -1  'True
            WindowState     =   2
            PrintFileLinesPerPage=   60
            WindowShowRefreshBtn=   -1  'True
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   8205
         _Version        =   65536
         _ExtentX        =   14473
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
            Height          =   495
            Left            =   660
            TabIndex        =   2
            Top             =   60
            Width           =   3795
            _Version        =   65536
            _ExtentX        =   6694
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Reporte de Cartas Fianza"
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
            Picture         =   "OpeTra_frm_039.frx":0B9A
            Top             =   90
            Width           =   480
         End
         Begin VB.Image Image2 
            Height          =   480
            Left            =   3150
            Picture         =   "OpeTra_frm_039.frx":0EA4
            Top             =   90
            Width           =   480
         End
      End
   End
End
Attribute VB_Name = "frm_RptFia_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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

Private Sub cmb_PerMes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     Call gs_SetFocus(ipp_PerAno)
   End If
End Sub

Private Sub cmb_TipRep_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     Call gs_SetFocus(cmb_PerMes)
   End If
End Sub

Private Sub cmd_ExpExc_Click()
   If cmb_TipRep.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Reporte.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipRep)
      Exit Sub
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
   If MsgBox("¿Está seguro de exportar a Excel el Reporte?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Me.Enabled = False
   
   Call fs_GenExc
   
   Me.Enabled = True
   Screen.MousePointer = 0
 
End Sub
Private Sub cmd_Imprim_Click()
   If cmb_TipRep.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Reporte.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipRep)
      Exit Sub
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
   Call fs_Imp_Genera
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
   Screen.MousePointer = 0
End Sub
Private Sub fs_Inicia()
   cmb_TipRep.Clear
   cmb_TipRep.AddItem "TODAS LAS CARTAS FIANZA"
   cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = 1
   cmb_TipRep.AddItem "CARTAS FIANZA VIGENTES"
   cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = 2
   cmb_TipRep.AddItem "CARTAS FIANZA VENCIDAS"
   cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = 3
   cmb_TipRep.AddItem "CARTAS FIANZA REQUERIDAS"
   cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = 4
   cmb_TipRep.AddItem "CARTAS FIANZA EJECUTADAS"
   cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = 5
   cmb_TipRep.AddItem "CARTAS FIANZA LIBERADAS"
   cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = 6
   cmb_TipRep.AddItem "CARTAS FIANZA VIGENTES (POR BANCO)"
   cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = 7
   
   Call moddat_gs_Carga_LisIte_Combo(cmb_PerMes, 1, "033")
End Sub

Private Sub fs_Limpia()
   cmb_TipRep.ListIndex = -1
   ipp_PerAno.Text = Year(date)
End Sub
Private Sub fs_Imp_Genera()
Dim r_str_NumOpe     As String
Dim r_str_NumDoc     As String
Dim r_int_TipDoc     As Integer
Dim r_str_FecIni     As String
Dim r_str_FecFin     As String

   If cmb_PerMes.ListIndex = -1 Then
      r_str_FecIni = "01" & Format(Month(date), "00") & Format(Year(date), "0000")
      r_str_FecFin = Format(date, "DD/MM/YYYY")
   Else
      r_str_FecIni = "01" & "/" & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") & "/" & Format(ipp_PerAno.Text, "0000")
      r_str_FecFin = Format(ff_Ultimo_Dia_Mes(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), ipp_PerAno.Text), "00") & "/" & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") & "/" & Format(ipp_PerAno.Text, "0000")
   End If
   
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
   
   'Borrando Spool Local
   If Not gf_EjecutaSQL("DELETE FROM RPT_CARFIA", g_rst_Genera, 2) Then
      Exit Sub
   End If
   
   'Generando Reporte
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " USP_RPT_CARTA_FIANZA ("
   
   If cmb_PerMes.ListIndex = -1 Then
     g_str_Parame = g_str_Parame & "" & Month(Now) & ", "
     g_str_Parame = g_str_Parame & "" & Format(Year(Now), "0000") & ", "
   Else
     g_str_Parame = g_str_Parame & "" & cmb_PerMes.ItemData(cmb_PerMes.ListIndex) & ", "
     g_str_Parame = g_str_Parame & "" & Format(ipp_PerAno.Text, "0000") & ", "
   End If
   
   g_str_Parame = g_str_Parame & "'REPORTE CARTAS FIANZAS', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   g_str_Parame = g_str_Parame & Chk_FecAct.Value & ", "
   g_str_Parame = g_str_Parame & cmb_TipRep.ItemData(cmb_TipRep.ListIndex) & ") "
        
   DoEvents: DoEvents: DoEvents
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   DoEvents: DoEvents: DoEvents
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se encontraron registros.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
          r_str_NumOpe = Left(g_rst_Princi!NUMOPE, 3) & "-" & Mid(g_rst_Princi!NUMOPE, 4, 2) & "-" & Right(g_rst_Princi!NUMOPE, 5)

            g_str_Parame = ""
            g_str_Parame = g_str_Parame & "INSERT INTO RPT_CARFIA("
            g_str_Parame = g_str_Parame & "CARFIA_NUMOPE, "
            g_str_Parame = g_str_Parame & "CARFIA_DOCIDE, "
            g_str_Parame = g_str_Parame & "CARFIA_NOMCLI, "
            g_str_Parame = g_str_Parame & "CARFIA_NOMBCO, "
            g_str_Parame = g_str_Parame & "CARFIA_NUMFIA, "
            g_str_Parame = g_str_Parame & "CARFIA_MONFIA, "
            g_str_Parame = g_str_Parame & "CARFIA_MTOFIA, "
            g_str_Parame = g_str_Parame & "CARFIA_FECEMI, "
            g_str_Parame = g_str_Parame & "CARFIA_FECVCT, "
            g_str_Parame = g_str_Parame & "CARFIA_FECREQ, "
            g_str_Parame = g_str_Parame & "CARFIA_FECEJE, "
            g_str_Parame = g_str_Parame & "CARFIA_SITUAC, "
            g_str_Parame = g_str_Parame & "CARFIA_FECLIB, "
            g_str_Parame = g_str_Parame & "CARFIA_FECINI, "
            g_str_Parame = g_str_Parame & "CARFIA_FECFIN) "
            
            g_str_Parame = g_str_Parame & "VALUES ("
            g_str_Parame = g_str_Parame & "'" & r_str_NumOpe & "', "
            g_str_Parame = g_str_Parame & "'" & g_rst_Princi!DOCIDE & "', "
            g_str_Parame = g_str_Parame & "'" & g_rst_Princi!NomCli & "', "
            g_str_Parame = g_str_Parame & "'" & g_rst_Princi!NOMBCO & "', "
            g_str_Parame = g_str_Parame & "'" & g_rst_Princi!NUMFIA & "', "
            g_str_Parame = g_str_Parame & "'" & g_rst_Princi!MONFIA & "', "
            g_str_Parame = g_str_Parame & "" & g_rst_Princi!MTOFIA & ", "
            g_str_Parame = g_str_Parame & "'" & gf_FormatoFecha(CStr(g_rst_Princi!FECEMI)) & "', "
            g_str_Parame = g_str_Parame & "'" & gf_FormatoFecha(CStr(g_rst_Princi!FECVCT)) & "', "
            g_str_Parame = g_str_Parame & "'" & gf_FormatoFecha(CStr(g_rst_Princi!FECREQ)) & "', "
            g_str_Parame = g_str_Parame & "'" & gf_FormatoFecha(CStr(g_rst_Princi!FECEJE)) & "', "
            g_str_Parame = g_str_Parame & "'" & g_rst_Princi!SITUACION & "', "
            g_str_Parame = g_str_Parame & "'" & gf_FormatoFecha(CStr(g_rst_Princi!FECLIB)) & "', "
            g_str_Parame = g_str_Parame & "'" & CStr(r_str_FecIni) & "',"
            g_str_Parame = g_str_Parame & "'" & CStr(r_str_FecFin) & "')"
            
            If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
               Exit Sub
            End If
         
         g_rst_Princi.MoveNext
         DoEvents
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Select Case cmb_TipRep.ItemData(cmb_TipRep.ListIndex)
      Case 1:  crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_CARFIA_01.RPT"
      Case 2:  crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_CARFIA_02.RPT"
      Case 3:  crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_CARFIA_03.RPT"
      Case 4:  crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_CARFIA_04.RPT"
      Case 5:  crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_CARFIA_05.RPT"
      Case 6:  crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_CARFIA_06.RPT"
      Case 7:  crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_CARFIA_08.RPT"
   End Select
   
   crp_Imprim.WindowShowPrintSetupBtn = True
   crp_Imprim.Action = 1
End Sub

Private Sub fs_GenExc() '_Period()
Dim r_obj_Excel         As Excel.Application
Dim r_int_ConVer        As Integer
Dim r_str_NumOpe        As String
Dim r_str_NumOpeAnt     As String
Dim r_int_TipRep        As Integer
Dim r_str_NomBco        As String
Dim r_str_TIPMON        As String
Dim r_str_FecIni        As String
    
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " USP_RPT_CARTA_FIANZA ("
   
   If cmb_PerMes.ListIndex = -1 Then
     g_str_Parame = g_str_Parame & "" & Month(Now) & ", "
     g_str_Parame = g_str_Parame & "" & Format(Year(Now), "0000") & ", "
   Else
     g_str_Parame = g_str_Parame & "" & cmb_PerMes.ItemData(cmb_PerMes.ListIndex) & ", "
     g_str_Parame = g_str_Parame & "" & Format(ipp_PerAno.Text, "0000") & ", "
   End If
   
   g_str_Parame = g_str_Parame & "'REPORTE CARTAS FIANZA', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   g_str_Parame = g_str_Parame & Chk_FecAct.Value & ", "
   g_str_Parame = g_str_Parame & cmb_TipRep.ItemData(cmb_TipRep.ListIndex) & ") "
        
   DoEvents: DoEvents: DoEvents
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   DoEvents: DoEvents: DoEvents
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se encontraron registros.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   r_int_TipRep = cmb_TipRep.ItemData(cmb_TipRep.ListIndex)
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      Set r_obj_Excel = New Excel.Application
      r_obj_Excel.SheetsInNewWorkbook = 1
      r_obj_Excel.Workbooks.Add
        
      Call fs_Cabecera(r_obj_Excel, r_int_TipRep)
        
      g_rst_Princi.MoveFirst
       
      Select Case r_int_TipRep
          Case 7:     r_int_ConVer = 8
          Case Else:  r_int_ConVer = 5
      End Select
        
      If Not g_rst_Princi.BOF And Not g_rst_Princi.EOF Then
          g_rst_Princi.MoveFirst
          r_str_NomBco = Trim(g_rst_Princi!NOMBCO)
          r_str_TIPMON = Trim(g_rst_Princi!MONFIA)
          If r_int_TipRep = 7 Then r_obj_Excel.Sheets(1).Name = IIf(IsNull(Trim(g_rst_Princi!NOMBCO)), "SIN BANCO", Trim(g_rst_Princi!NOMBCO)) & "_" & IIf(Trim(g_rst_Princi!MONFIA) = "S/.", "SOLES", IIf(Trim(g_rst_Princi!MONFIA) = "$US", "DOLARES", "SIN MONEDA"))
      End If
        
      Do While Not g_rst_Princi.EOF
         If IsNull(g_rst_Princi!NUMOPE) Then
             r_str_NumOpe = Left(g_rst_Princi!NUMOPE_CIERRE, 3) & "-" & Mid(g_rst_Princi!NUMOPE_CIERRE, 4, 2) & "-" & Right(g_rst_Princi!NUMOPE_CIERRE, 5)
         Else
             r_str_NumOpe = Left(g_rst_Princi!NUMOPE, 3) & "-" & Mid(g_rst_Princi!NUMOPE, 4, 2) & "-" & Right(g_rst_Princi!NUMOPE, 5)
         End If
         Select Case r_int_TipRep
            Case 7:
               If IIf(IsNull(g_rst_Princi!NOMBCO), "", g_rst_Princi!NOMBCO) <> r_str_NomBco Then
                  If IIf(IsNull(g_rst_Princi!MONFIA), "", g_rst_Princi!MONFIA) = r_str_TIPMON Then
Llenar:
                     r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = "Total por Moneda ->"
                     r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5).FormulaR1C1 = "=SUM(R[-" & r_int_ConVer - 8 & "]C:R[-1]C)"
                     r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(4, 2), r_obj_Excel.ActiveSheet.Cells(5, 2)).Font.Bold = True
                     r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4), r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5)).Font.Bold = True
                     r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(4, 2), r_obj_Excel.ActiveSheet.Cells(5, 2)).HorizontalAlignment = xlHAlignLeft
                   
                     r_str_NomBco = IIf(IsNull(g_rst_Princi!NOMBCO), "", Trim(g_rst_Princi!NOMBCO))
                     r_str_TIPMON = IIf(IsNull(g_rst_Princi!MONFIA), "", Trim(g_rst_Princi!MONFIA))
                               
                     r_obj_Excel.Sheets(r_obj_Excel.Sheets.Count).Select
                     r_obj_Excel.Sheets(r_obj_Excel.Sheets.Count).Cells(8, 1).Select
                     r_obj_Excel.ActiveWindow.FreezePanes = True
      
                     'AGREGA UNA NUEVA HOJA
                     r_obj_Excel.Sheets.Add After:=r_obj_Excel.Sheets(r_obj_Excel.Sheets.Count)
                     r_obj_Excel.Sheets(r_obj_Excel.Sheets.Count).Name = IIf(IsNull(Trim(g_rst_Princi!NOMBCO)), "SIN BANCO", Trim(g_rst_Princi!NOMBCO)) & "_" & IIf(Trim(g_rst_Princi!MONFIA) = "S/.", "SOLES", IIf(Trim(g_rst_Princi!MONFIA) = "US$", "DOLARES", "SIN MONEDA"))
                     Call fs_Cabecera(r_obj_Excel, r_int_TipRep)
                     r_obj_Excel.ActiveSheet.Cells(4, 2) = r_str_NomBco
                     r_obj_Excel.ActiveSheet.Cells(5, 2) = r_str_TIPMON
                     r_int_ConVer = 8
                  Else
                     r_str_NomBco = IIf(IsNull(g_rst_Princi!NOMBCO), "", Trim(g_rst_Princi!NOMBCO))
                  End If
               Else
                  If IIf(IsNull(g_rst_Princi!MONFIA), "", g_rst_Princi!MONFIA) = r_str_TIPMON Then
                      If r_obj_Excel.ActiveSheet.Cells(4, 2) = "" And r_obj_Excel.ActiveSheet.Cells(5, 2) = "" Then
                          r_obj_Excel.ActiveSheet.Cells(4, 2) = r_str_NomBco
                          r_obj_Excel.ActiveSheet.Cells(5, 2) = r_str_TIPMON
                      End If
                  Else
                      GoTo Llenar
                  End If
               End If
                   
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_str_NumOpe
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = Trim(g_rst_Princi!DOCIDE)
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = Trim(g_rst_Princi!NomCli)
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = Trim(g_rst_Princi!NUMFIA)
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = Format(g_rst_Princi!MTOFIA, "###,###,##0.00")
               If gf_FormatoFecha(CStr(g_rst_Princi!FECEMI)) <> "" Then
                    r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!FECEMI)))
               End If
               If gf_FormatoFecha(CStr(g_rst_Princi!FECVCT)) <> "" Then
                    r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!FECVCT)))
               End If
                   
            Case Else
           
               If r_str_NumOpeAnt = r_str_NumOpe And r_str_NumOpeAnt <> "" Then
                  r_int_ConVer = r_int_ConVer - 1
               End If
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_str_NumOpe
                   
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = Trim(g_rst_Princi!DOCIDE)
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = Trim(g_rst_Princi!NomCli)
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = Trim(g_rst_Princi!NOMBCO)
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = Trim(g_rst_Princi!NUMFIA)
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = Trim(g_rst_Princi!MONFIA)
               r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = Format(g_rst_Princi!MTOFIA, "###,###,##0.00")
               If gf_FormatoFecha(CStr(g_rst_Princi!FECEMI)) <> "" Then
                    r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!FECEMI)))
               End If
               If gf_FormatoFecha(CStr(g_rst_Princi!FECVCT)) <> "" Then
                    r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!FECVCT)))
               End If
                       
               If r_int_TipRep = 1 Then
                   If gf_FormatoFecha(CStr(g_rst_Princi!FECREQ)) <> "" Then
                       r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!FECREQ)))
                   End If
                   If gf_FormatoFecha(CStr(g_rst_Princi!FECEJE)) <> "" Then
                       r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!FECEJE)))
                   End If
                   If gf_FormatoFecha(CStr(g_rst_Princi!FECLIB)) <> "" Then
                       r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!FECLIB)))
                   End If
                   r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = Trim(g_rst_Princi!SITUACION)
               
               ElseIf r_int_TipRep = 6 Then
                   If gf_FormatoFecha(CStr(g_rst_Princi!FECLIB)) <> "" Then
                       r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!FECLIB)))
                   End If
               End If
                   
               r_str_NumOpeAnt = r_str_NumOpe
         End Select
         
         r_int_ConVer = r_int_ConVer + 1
         g_rst_Princi.MoveNext
         DoEvents
      Loop
        
      If r_int_TipRep = 7 Then
      'Última hoja
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = "Total por Moneda ->"
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5).FormulaR1C1 = "=SUM(R[-" & r_int_ConVer - 8 & "]C:R[-1]C)"
         r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(4, 2), r_obj_Excel.ActiveSheet.Cells(5, 2)).Font.Bold = True
         r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4), r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5)).Font.Bold = True
         r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(4, 2), r_obj_Excel.ActiveSheet.Cells(5, 2)).HorizontalAlignment = xlHAlignLeft
          
         r_obj_Excel.Sheets(r_obj_Excel.Sheets.Count).Select
         r_obj_Excel.Sheets(r_obj_Excel.Sheets.Count).Cells(8, 1).Select
         r_obj_Excel.ActiveWindow.FreezePanes = True
      End If
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      Screen.MousePointer = 0
      r_obj_Excel.Visible = True
      Set r_obj_Excel = Nothing
   End If
End Sub
Private Sub fs_Cabecera(ByRef r_Excel As Excel.Application, ByVal r_int_TipRep As Integer)
   With r_Excel.ActiveSheet
      Select Case r_int_TipRep
         Case 1, 2, 3, 4, 5, 6:
            .Cells(4, 1) = "NRO. OPERACION"
            .Cells(4, 2) = "DOC. IDENTIDAD"
            .Cells(4, 3) = "NOMBRE CLIENTE"
            .Cells(4, 4) = "BANCO"
            .Cells(4, 5) = "NRO. CARTA FIANZA"
            .Range(.Cells(4, 6), .Cells(4, 7)).Merge
            .Cells(4, 6) = "MTO. CARTA FIANZA"
            .Cells(4, 8) = "F. EMISION"
            .Cells(4, 9) = "F. VENCIMIENTO"
               
            If r_int_TipRep = 1 Then
               If cmb_PerMes.ListIndex = -1 Then
                  .Cells(2, 1) = "REPORTE GENERAL DE CARTAS FIANZA AL " & Format(date, "DD/MM/YYYY")
               Else
                  .Cells(2, 1) = "REPORTE GENERAL DE CARTAS FIANZA AL " & Format(ff_Ultimo_Dia_Mes(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), ipp_PerAno.Text), "00") & "/" & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") & "/" & Format(ipp_PerAno.Text, "0000")
               End If
               .Range(.Cells(2, 1), .Cells(2, 13)).Merge

               .Cells(4, 10) = "F. REQUERIMIENTO"
               .Cells(4, 11) = "F. EJECUCION"
               .Cells(4, 12) = "F. LIBERACION"
               .Cells(4, 13) = "SITUACION"
                    
               .Range(.Cells(1, 1), .Cells(4, 13)).Font.Bold = True
               .Range(.Cells(1, 1), .Cells(4, 13)).HorizontalAlignment = xlHAlignCenter
                    
               .Columns("J").ColumnWidth = 17.5
               .Columns("J").HorizontalAlignment = xlHAlignCenter
               .Columns("K").ColumnWidth = 15
               .Columns("K").HorizontalAlignment = xlHAlignCenter
               .Columns("L").ColumnWidth = 15
               .Columns("L").HorizontalAlignment = xlHAlignCenter
               .Columns("M").ColumnWidth = 15
               .Columns("M").HorizontalAlignment = xlHAlignCenter
                
            ElseIf r_int_TipRep = 6 Then
               If cmb_PerMes.ListIndex = -1 Then
                  .Cells(2, 1) = "REPORTE DE " & UCase(cmb_TipRep.Text) & " AL " & Format(date, "DD/MM/YYYY")
               Else
                  .Cells(2, 1) = "REPORTE DE " & UCase(cmb_TipRep.Text) & " AL " & Format(ff_Ultimo_Dia_Mes(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), ipp_PerAno.Text), "00") & "/" & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") & "/" & Format(ipp_PerAno.Text, "0000")
               End If
               
               .Range(.Cells(2, 1), .Cells(2, 10)).Merge
               .Cells(4, 10) = "F. LIBERACION"
               
               .Range(.Cells(1, 1), .Cells(4, 10)).Font.Bold = True
               .Range(.Cells(1, 1), .Cells(4, 10)).HorizontalAlignment = xlHAlignCenter
               
               .Columns("J").ColumnWidth = 15
               .Columns("J").HorizontalAlignment = xlHAlignCenter
            Else
               If cmb_PerMes.ListIndex = -1 Then
                  .Cells(2, 1) = "REPORTE AL " & UCase(cmb_TipRep.Text) & " AL " & Format(date, "DD/MM/YYYY")
               Else
                  .Cells(2, 1) = "REPORTE AL " & UCase(cmb_TipRep.Text) & " AL " & Format(ff_Ultimo_Dia_Mes(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), ipp_PerAno.Text), "00") & "/" & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") & "/" & Format(ipp_PerAno.Text, "0000")
               End If
               
               .Range(.Cells(2, 1), .Cells(2, 9)).Merge
               .Range(.Cells(1, 1), .Cells(4, 9)).Font.Bold = True
               .Range(.Cells(1, 1), .Cells(4, 9)).HorizontalAlignment = xlHAlignCenter
            End If
                
         Case 7:
            .Cells(4, 1) = "BANCO"
            .Cells(5, 1) = "MONEDA"
            .Cells(7, 1) = "NRO. OPERACION"
            .Cells(7, 2) = "DOC. IDENTIDAD"
            .Cells(7, 3) = "NOMBRE CLIENTE"
            .Cells(7, 4) = "NRO. CARTA FIANZA"
            .Cells(7, 5) = "MTO. CARTA FIANZA"
            .Cells(7, 6) = "F. EMISION"
            .Cells(7, 7) = "F. VENCIMIENTO"
            If cmb_PerMes.ListIndex = -1 Then
              .Cells(2, 1) = "REPORTE DE " & UCase(cmb_TipRep.Text) & " AL " & Format(date, "DD/MM/YYYY")
            Else
              .Cells(2, 1) = "REPORTE DE " & UCase(cmb_TipRep.Text) & " AL " & Format(ff_Ultimo_Dia_Mes(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), ipp_PerAno.Text), "00") & "/" & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") & "/" & Format(ipp_PerAno.Text, "0000")
            End If
            
            .Range(.Cells(2, 1), .Cells(2, 7)).Merge
            .Range(.Cells(1, 1), .Cells(7, 7)).Font.Bold = True
            .Range(.Cells(1, 1), .Cells(7, 7)).HorizontalAlignment = xlHAlignCenter
                
      End Select
        
      .Columns("A").ColumnWidth = 16
      .Columns("A").HorizontalAlignment = xlHAlignCenter
      .Columns("B").ColumnWidth = 15
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 42
      If r_int_TipRep = 7 Then
           .Columns("D").ColumnWidth = 21.8
           .Columns("D").NumberFormat = "@"
           .Columns("D").HorizontalAlignment = xlHAlignCenter
           .Columns("E").ColumnWidth = 19
           .Columns("E").NumberFormat = "#,##0.00"
           .Columns("E").HorizontalAlignment = xlHAlignRight
           .Columns("F").ColumnWidth = 15
           .Columns("F").HorizontalAlignment = xlHAlignCenter
           .Columns("G").HorizontalAlignment = xlHAlignCenter
      Else
           .Columns("D").ColumnWidth = 26
           .Columns("D").NumberFormat = "@"
           .Columns("E").ColumnWidth = 23
           .Columns("E").HorizontalAlignment = xlHAlignCenter
           .Columns("F").ColumnWidth = 9.5
           .Columns("F").HorizontalAlignment = xlHAlignRight
           .Columns("F").NumberFormat = "#,##0.00"
      End If
      .Columns("G").ColumnWidth = 15
      .Columns("H").ColumnWidth = 15
      .Columns("H").HorizontalAlignment = xlHAlignCenter
      .Columns("I").ColumnWidth = 15
      .Columns("I").HorizontalAlignment = xlHAlignCenter
              
   End With
   
   Select Case r_int_TipRep
      Case 1, 2, 3, 4, 5, 6
            r_Excel.Sheets(1).Select
            r_Excel.Sheets(1).Cells(5, 1).Select
            r_Excel.ActiveWindow.FreezePanes = True
   End Select
End Sub

Private Sub ipp_PerAno_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_ExpExc)
    End If
End Sub
