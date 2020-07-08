VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_Pro_EvaPBP_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form3"
   ClientHeight    =   2265
   ClientLeft      =   5235
   ClientTop       =   3690
   ClientWidth     =   8040
   Icon            =   "OpeTra_frm_291.frx":0000
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   2265
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   8055
      _Version        =   65536
      _ExtentX        =   14208
      _ExtentY        =   3995
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
         Height          =   765
         Left            =   30
         TabIndex        =   5
         Top             =   1440
         Width           =   7965
         _Version        =   65536
         _ExtentX        =   14049
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
         Begin VB.ComboBox cmb_PerMes 
            Height          =   315
            Left            =   1260
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   6645
         End
         Begin EditLib.fpDoubleSingle ipp_PerAno 
            Height          =   315
            Left            =   1260
            TabIndex        =   1
            Top             =   390
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
         Begin VB.Label Label2 
            Caption         =   "Año:"
            Height          =   255
            Left            =   60
            TabIndex        =   10
            Top             =   390
            Width           =   915
         End
         Begin VB.Label Label3 
            Caption         =   "Mes:"
            Height          =   255
            Left            =   60
            TabIndex        =   9
            Top             =   60
            Width           =   915
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   6
         Top             =   30
         Width           =   7965
         _Version        =   65536
         _ExtentX        =   14049
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
            Height          =   555
            Left            =   570
            TabIndex        =   7
            Top             =   30
            Width           =   4605
            _Version        =   65536
            _ExtentX        =   8123
            _ExtentY        =   979
            _StockProps     =   15
            Caption         =   "Evaluación y Asignación de Premio Buen Pagador"
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
            Picture         =   "OpeTra_frm_291.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   8
         Top             =   750
         Width           =   7965
         _Version        =   65536
         _ExtentX        =   14049
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
         Begin VB.CommandButton cmd_Proces 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_291.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Nueva Evaluación"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   7350
            Picture         =   "OpeTra_frm_291.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
      End
   End
End
Attribute VB_Name = "frm_Pro_EvaPBP_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmb_PerMes_Click()
   Call gs_SetFocus(ipp_PerAno)
End Sub

Private Sub cmb_PerMes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_PerMes_Click
   End If
End Sub

Private Sub cmd_Proces_Click()
   If cmb_PerMes.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Mes.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PerMes)
      Exit Sub
   End If
   
   'Verificar si ya se ha ejecutado el Proceso
   g_str_Parame = "SELECT * FROM CRE_CABPBP WHERE "
   g_str_Parame = g_str_Parame & "CABPBP_PERMES = " & CStr(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)) & " AND "
   g_str_Parame = g_str_Parame & "CABPBP_PERANO = " & Format(ipp_PerAno.Text, "0000") & " "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      MsgBox "No se pudo leer la tabla CRE_CABPBP.", vbCritical, modgen_g_str_NomPlt
      Exit Sub
   End If

   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      MsgBox "El Proceso de Evaluación ya ha sido generado.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PerMes)
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de ejecutar el Proceso?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_GeneraPBP(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), CInt(ipp_PerAno.Text))
   Screen.MousePointer = 0

   moddat_g_int_FlgAct = 2
   MsgBox "El Proceso ha terminado correctamente.", vbInformation, modgen_g_str_NomPlt
   Unload Me
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   Call gs_CentraForm(Me)
   Call fs_Inicia
   Call fs_Limpia
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_LisIte_Combo(cmb_PerMes, 1, "033")
End Sub

Private Sub fs_Limpia()
   cmb_PerMes.ListIndex = -1
   ipp_PerAno.Text = Format(Year(date), "0000")
End Sub

Private Sub ipp_PerAno_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Proces)
   End If
End Sub

Private Sub fs_GeneraPBP(ByVal p_PerMes As Integer, ByVal p_PerAno As Integer)
   Dim r_arr_ParPrd()   As moddat_tpo_Genera
   Dim r_arr_DiaAtr(6)  As Integer
   Dim r_str_NumOpe     As String
   Dim r_int_NumCuo     As Integer
   Dim r_int_CuoIni     As Integer
   Dim r_int_CuoFin     As Integer
   Dim r_int_AtrMax     As Integer
   Dim r_int_DiaAtr     As Integer
   Dim r_str_FecVct     As String
   Dim r_str_FecPag     As String
   Dim r_rst_Cuotas     As ADODB.Recordset
   Dim r_rst_CuoCon     As ADODB.Recordset
   Dim r_int_Contad     As Integer
   Dim r_int_FlgApl     As Integer
   Dim r_int_FlgPBP     As Integer
   
   Dim r_dbl_CapCof     As Double
   Dim r_dbl_IntCof     As Double
   Dim r_dbl_ComCof     As Double
   
   Dim r_dbl_CapPen_Cli          As Double
   Dim r_dbl_IntPen_Cli          As Double
   Dim r_dbl_CapPen_Ult_Cli      As Double
   Dim r_dbl_IntPen_Ult_Cli      As Double
   
   Dim r_dbl_CapPen_Cli_1        As Double
   Dim r_dbl_CapPen_Cli_2        As Double
   Dim r_dbl_CapPen_Cli_3        As Double
   Dim r_dbl_CapPen_Cli_4        As Double
   Dim r_dbl_CapPen_Cli_5        As Double
   Dim r_dbl_CapPen_Cli_6        As Double
   
   Dim r_dbl_IntPen_Cli_1        As Double
   Dim r_dbl_IntPen_Cli_2        As Double
   Dim r_dbl_IntPen_Cli_3        As Double
   Dim r_dbl_IntPen_Cli_4        As Double
   Dim r_dbl_IntPen_Cli_5        As Double
   Dim r_dbl_IntPen_Cli_6        As Double
   
   Dim r_dbl_CapPen_Cof          As Double
   Dim r_dbl_IntPen_Cof          As Double
   Dim r_dbl_ComPen_Cof          As Double
   Dim r_dbl_CapPen_Ult_Cof      As Double
   Dim r_dbl_IntPen_Ult_Cof      As Double
   Dim r_dbl_ComPen_Ult_Cof      As Double
   
   Dim r_int_CuoDis              As Integer
   Dim r_int_DiaMax              As Integer
   Dim r_str_FecIni              As String
   Dim r_str_FecFin              As String
   Dim r_dbl_CapCli              As Double
   Dim r_dbl_IntCli              As Double
   
   Dim r_int_NumEva              As Integer
   Dim r_int_EvaAsg              As Integer
   Dim r_int_EvaPer              As Integer
   
   Dim r_int_CuoIni_Eval         As Integer
   Dim r_int_CuoFin_Eval         As Integer
   Dim r_int_CuoIni_Cast         As Integer
   Dim r_int_CuoFin_Cast         As Integer
   
   Screen.MousePointer = 11
   
   r_int_NumEva = 0
   r_int_EvaAsg = 0
   r_int_EvaPer = 0
   r_int_CuoIni_Eval = 0
   r_int_CuoFin_Eval = 0
   r_int_CuoIni_Cast = 0
   r_int_CuoFin_Cast = 0
   
   r_str_FecIni = Format(CDate("01/" & Format(p_PerMes, "00") & "/" & Format(p_PerAno, "0000")), "yyyymmdd")
   r_str_FecFin = Format(CDate(Format(ff_Ultimo_Dia_Mes(p_PerMes, p_PerAno), "00") & "/" & Format(p_PerMes, "00") & "/" & Format(p_PerAno, "0000")), "yyyymmdd")
   
   'Creando Cursor Principal
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT HIPMAE_CODPRD, HIPMAE_CODSUB, HIPCUO_NUMOPE, HIPMAE_SALCON, HIPCUO_NUMCUO, HIPCUO_CAPITA, HIPCUO_INTERE, HIPCUO_COMCOF, HIPCUO_SALCAP "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPCUO A, CRE_HIPMAE B "
   g_str_Parame = g_str_Parame & " WHERE HIPCUO_NUMOPE = HIPMAE_NUMOPE "
   g_str_Parame = g_str_Parame & "   AND (HIPMAE_SITUAC = 2 OR (HIPMAE_SITUAC = 6 AND SUBSTR(HIPMAE_FECCAN,1,6) = " & Mid(r_str_FecIni, 1, 6) & ")) "
   g_str_Parame = g_str_Parame & "   AND HIPMAE_CUOPEN > 0 "
   g_str_Parame = g_str_Parame & "   AND HIPMAE_JUDICI = 0 "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_FECVCT >= " & r_str_FecIni & " "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_FECVCT <= " & r_str_FecFin & " "
   g_str_Parame = g_str_Parame & "   AND ((HIPMAE_CODPRD = '001' AND HIPCUO_TIPCRO = 4) OR "
   g_str_Parame = g_str_Parame & "        (HIPMAE_CODPRD = '003' AND HIPCUO_TIPCRO = 4) OR "
   g_str_Parame = g_str_Parame & "        (HIPMAE_CODPRD = '004' AND HIPCUO_TIPCRO = 4) OR "
   g_str_Parame = g_str_Parame & "        (HIPMAE_CODPRD = '006' AND HIPCUO_TIPCRO = 2) OR "
   g_str_Parame = g_str_Parame & "        (HIPMAE_CODPRD = '007' AND HIPCUO_TIPCRO = 4) OR "
   g_str_Parame = g_str_Parame & "        (HIPMAE_CODPRD = '009' AND HIPCUO_TIPCRO = 4) OR "
   g_str_Parame = g_str_Parame & "        (HIPMAE_CODPRD = '010' AND HIPCUO_TIPCRO = 4) OR "
   g_str_Parame = g_str_Parame & "        (HIPMAE_CODPRD = '013' AND HIPCUO_TIPCRO = 4) OR "
   g_str_Parame = g_str_Parame & "        (HIPMAE_CODPRD = '014' AND HIPCUO_TIPCRO = 4) OR "
   g_str_Parame = g_str_Parame & "        (HIPMAE_CODPRD = '015' AND HIPCUO_TIPCRO = 4) OR "
   g_str_Parame = g_str_Parame & "        (HIPMAE_CODPRD = '016' AND HIPCUO_TIPCRO = 4) OR "
   g_str_Parame = g_str_Parame & "        (HIPMAE_CODPRD = '017' AND HIPCUO_TIPCRO = 4) OR "
   g_str_Parame = g_str_Parame & "        (HIPMAE_CODPRD = '018' AND HIPCUO_TIPCRO = 4)) "
   
   g_str_Parame = g_str_Parame & " UNION "
   
   g_str_Parame = g_str_Parame & "SELECT HIPMAE_CODPRD, HIPMAE_CODSUB, HIPCUO_NUMOPE, HIPMAE_SALCON, HIPCUO_NUMCUO, HIPCUO_CAPITA, HIPCUO_INTERE, HIPCUO_COMCOF, HIPCUO_SALCAP "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPCUO A "
   g_str_Parame = g_str_Parame & " INNER JOIN CRE_HIPMAE B ON B.HIPMAE_NUMOPE = A.HIPCUO_NUMOPE "
   g_str_Parame = g_str_Parame & " WHERE HIPCUO_NUMOPE IN (SELECT HIPMAE_NUMOPE FROM CRE_HIPMAE "
   g_str_Parame = g_str_Parame & "                         Where HIPMAE_SITUAC = 9 "
   g_str_Parame = g_str_Parame & "                           AND HIPMAE_FECCAN >= " & r_str_FecIni & " AND HIPMAE_FECCAN <= " & r_str_FecFin & " AND HIPMAE_CUOPEN = 0 AND HIPMAE_CODPRD IN ('001','003','004','006','007','009','010','013','014','015','016','017','018')) "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_TIPCRO = 4 "
   g_str_Parame = g_str_Parame & "   AND HIPCUO_FECVCT >= " & r_str_FecIni
   g_str_Parame = g_str_Parame & "   AND HIPCUO_FECVCT <= " & r_str_FecFin
   g_str_Parame = g_str_Parame & " ORDER BY HIPCUO_NUMOPE ASC "
 
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      'Para leer Parámetro de Días de Atraso Máximo para Aplicación de PBP
      r_int_AtrMax = 0
      
      If moddat_gf_Consulta_ParSubPrd(r_arr_ParPrd, g_rst_Princi!HIPMAE_CODPRD, g_rst_Princi!HIPMAE_CODSUB, "051", "041") Then
         r_int_AtrMax = r_arr_ParPrd(1).Genera_Cantid
      End If
      
      r_str_NumOpe = g_rst_Princi!HIPCUO_NUMOPE
      r_int_NumCuo = CInt(g_rst_Princi!HIPCUO_NUMCUO)
      
      'Cuotas de Evaluacion
      r_int_CuoIni_Eval = (r_int_NumCuo * 6) - 6
      r_int_CuoFin_Eval = (r_int_NumCuo * 6) - 1
      If r_int_NumCuo = 1 Then
         r_int_CuoIni_Eval = 1
      End If
      
      'Para leer Cuotas de TNC de cada Operación y determinar Días de Atraso de cada cuota
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT HIPCUO_NUMCUO, HIPCUO_FECVCT, HIPCUO_FECPAG, HIPCUO_SITUAC "
      g_str_Parame = g_str_Parame & "  FROM CRE_HIPCUO "
      g_str_Parame = g_str_Parame & " WHERE HIPCUO_NUMOPE = '" & g_rst_Princi!HIPCUO_NUMOPE & "' "
      g_str_Parame = g_str_Parame & "   AND HIPCUO_TIPCRO = 1 "
      g_str_Parame = g_str_Parame & "   AND HIPCUO_NUMCUO >= " & CStr(r_int_CuoIni_Eval) & " "
      g_str_Parame = g_str_Parame & "   AND HIPCUO_NUMCUO <= " & CStr(r_int_CuoFin_Eval) & " "
      g_str_Parame = g_str_Parame & "ORDER BY HIPCUO_NUMCUO ASC"
   
      If Not gf_EjecutaSQL(g_str_Parame, r_rst_Cuotas, 3) Then
         Exit Sub
      End If
   
      If Not (r_rst_Cuotas.BOF And r_rst_Cuotas.EOF) Then
         r_int_Contad = 1
         r_int_FlgApl = 0
         r_int_DiaMax = 0
         r_rst_Cuotas.MoveFirst
         
         Do While Not r_rst_Cuotas.EOF
            If r_rst_Cuotas!HIPCUO_SITUAC = 2 Then
               r_str_FecVct = gf_FormatoFecha(CStr(r_rst_Cuotas!HIPCUO_FECVCT))
               
               r_int_DiaAtr = 0
               If CInt((date - CDate(1)) - CDate(r_str_FecVct)) > r_int_AtrMax Then
                  r_int_DiaAtr = CInt((date - CDate(1)) - CDate(r_str_FecVct))
               Else
                  r_int_FlgApl = 1
               End If
            Else
               r_str_FecVct = gf_FormatoFecha(CStr(r_rst_Cuotas!HIPCUO_FECVCT))
               r_str_FecPag = gf_FormatoFecha(CStr(r_rst_Cuotas!HIPCUO_FECPAG))
               
               If CDate(r_str_FecPag) > CDate(r_str_FecVct) Then
                  r_int_DiaAtr = CInt(CDate(r_str_FecPag) - CDate(r_str_FecVct))
               Else
                  r_int_DiaAtr = 0
               End If
            End If
            
            r_arr_DiaAtr(r_int_Contad) = r_int_DiaAtr
            If r_int_DiaAtr > r_int_DiaMax Then
               r_int_DiaMax = r_int_DiaAtr
            End If
            
            r_int_Contad = r_int_Contad + 1
            r_rst_Cuotas.MoveNext
            DoEvents
         Loop
      End If
   
      r_rst_Cuotas.Close
      Set r_rst_Cuotas = Nothing
         
      'Para determinar si se asigna PBP
      r_int_FlgPBP = 3
      
      If (r_int_FlgApl = 0) Or (r_int_FlgApl = 1 And r_int_DiaMax > r_int_DiaAtr) Then
         r_int_FlgPBP = 1
         If r_int_CuoIni_Eval = 1 Then
            For r_int_Contad = 1 To 5
               If r_arr_DiaAtr(r_int_Contad) > r_int_AtrMax Then
                  r_int_FlgPBP = 2
               End If
            Next r_int_Contad
         Else
            For r_int_Contad = 1 To 6
               If r_arr_DiaAtr(r_int_Contad) > r_int_AtrMax Then
                  r_int_FlgPBP = 2
               End If
            Next r_int_Contad
         End If
      End If
      
      'Leyendo Capital, Interes y Comisión de Cronograma Cliente y Cofide/Mivivienda
      r_dbl_CapCli = 0
      r_dbl_IntCli = 0
      r_dbl_CapCof = 0
      r_dbl_IntCof = 0
      r_dbl_ComCof = 0
      
      If g_rst_Princi!HIPMAE_CODPRD = "006" Then
         r_dbl_CapCli = g_rst_Princi!HIPCUO_CAPITA
         r_dbl_IntCli = g_rst_Princi!HIPCUO_INTERE
      Else
         r_dbl_CapCof = g_rst_Princi!HIPCUO_CAPITA
         r_dbl_IntCof = g_rst_Princi!HIPCUO_INTERE
         r_dbl_ComCof = g_rst_Princi!HIPCUO_COMCOF
         r_dbl_CapCli = r_dbl_CapCof
         
         'Leer Cuota de TC (Cliente)
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & "SELECT HIPCUO_INTERE "
         g_str_Parame = g_str_Parame & "  FROM CRE_HIPCUO "
         g_str_Parame = g_str_Parame & " WHERE HIPCUO_NUMOPE = '" & g_rst_Princi!HIPCUO_NUMOPE & "' "
         g_str_Parame = g_str_Parame & "   AND HIPCUO_TIPCRO = 2 "
         g_str_Parame = g_str_Parame & "   AND HIPCUO_NUMCUO = " & CStr(r_int_NumCuo) & " "
         
         If Not gf_EjecutaSQL(g_str_Parame, r_rst_CuoCon, 3) Then
            Exit Sub
         End If
         
         r_rst_CuoCon.MoveFirst
         r_dbl_IntCli = r_rst_CuoCon!HIPCUO_INTERE
         r_rst_CuoCon.Close
         Set r_rst_CuoCon = Nothing
      End If
      
      'Obteniendo Nro. de Cuotas a Distribuir Penalidad
      r_int_CuoDis = 6
      r_int_CuoIni_Cast = (r_int_NumCuo * 6) + 1
      r_int_CuoFin_Cast = (r_int_NumCuo * 6) + 6
      
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT COUNT(*) AS TOTCUO FROM CRE_HIPCUO "
      g_str_Parame = g_str_Parame & " WHERE HIPCUO_NUMOPE = '" & g_rst_Princi!HIPCUO_NUMOPE & "' "
      g_str_Parame = g_str_Parame & "   AND HIPCUO_NUMCUO >= " & CStr(r_int_CuoIni_Cast) & " "
      g_str_Parame = g_str_Parame & "   AND HIPCUO_NUMCUO <= " & CStr(r_int_CuoFin_Cast) & " "
      g_str_Parame = g_str_Parame & "   AND HIPCUO_TIPCRO = 1 "
      g_str_Parame = g_str_Parame & "   AND HIPCUO_SITUAC = 2 "
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
         Exit Sub
      End If
      
      r_int_CuoDis = g_rst_Genera!TOTCUO
      If r_int_CuoDis = 0 Then r_int_FlgPBP = 1
      
      g_rst_Genera.Close
      Set g_rst_Genera = Nothing
      
      'Aplicar Penalidad PBP
      r_dbl_CapPen_Cli = 0:      r_dbl_IntPen_Cli = 0
      r_dbl_CapPen_Ult_Cli = 0:  r_dbl_IntPen_Ult_Cli = 0
      
      r_dbl_CapPen_Cli_1 = 0: r_dbl_CapPen_Cli_2 = 0: r_dbl_CapPen_Cli_3 = 0: r_dbl_CapPen_Cli_4 = 0: r_dbl_CapPen_Cli_5 = 0: r_dbl_CapPen_Cli_6 = 0
      r_dbl_IntPen_Cli_1 = 0: r_dbl_IntPen_Cli_2 = 0: r_dbl_IntPen_Cli_3 = 0: r_dbl_IntPen_Cli_4 = 0: r_dbl_IntPen_Cli_5 = 0: r_dbl_IntPen_Cli_6 = 0
      
      r_dbl_CapPen_Cof = 0:      r_dbl_IntPen_Cof = 0:      r_dbl_ComPen_Cof = 0
      r_dbl_CapPen_Ult_Cof = 0:  r_dbl_IntPen_Ult_Cof = 0:  r_dbl_ComPen_Ult_Cof = 0
         
      If r_int_FlgPBP = 2 Then
         r_int_EvaPer = r_int_EvaPer + 1
      
         r_dbl_CapPen_Cli = CDbl(Format(r_dbl_CapCli / r_int_CuoDis, "######0.00"))
         r_dbl_IntPen_Cli = CDbl(Format(r_dbl_IntCli / r_int_CuoDis, "#####0.00"))
         r_dbl_CapPen_Cof = CDbl(Format(r_dbl_CapCof / 6, "######0.00"))
         r_dbl_IntPen_Cof = CDbl(Format(r_dbl_IntCof / 6, "#####0.00"))
         r_dbl_ComPen_Cof = CDbl(Format(r_dbl_ComCof / 6, "#####0.00"))
         
         'Ajustando Capital Ultima Cuota
         If r_dbl_CapPen_Cli * r_int_CuoDis < r_dbl_CapCli Then
            r_dbl_CapPen_Ult_Cli = r_dbl_CapPen_Cli + (r_dbl_CapCli - (r_dbl_CapPen_Cli * r_int_CuoDis))
         Else
            r_dbl_CapPen_Ult_Cli = r_dbl_CapPen_Cli - ((r_dbl_CapPen_Cli * r_int_CuoDis) - r_dbl_CapCli)
         End If
         
         'Ajustando Interes Ultima Cuota
         If r_dbl_IntPen_Cli * r_int_CuoDis < r_dbl_IntCli Then
            r_dbl_IntPen_Ult_Cli = r_dbl_IntPen_Cli + (r_dbl_IntCli - (r_dbl_IntPen_Cli * r_int_CuoDis))
         Else
            r_dbl_IntPen_Ult_Cli = r_dbl_IntPen_Cli - ((r_dbl_IntPen_Cli * r_int_CuoDis) - r_dbl_IntCli)
         End If
         
         'Ajustando Capital COFIDE Ultima Cuota
         If r_dbl_CapPen_Cof * r_int_CuoDis < r_dbl_CapCof Then
            r_dbl_CapPen_Ult_Cof = r_dbl_CapPen_Cof + (r_dbl_CapCof - (r_dbl_CapPen_Cof * 6))
         Else
            r_dbl_CapPen_Ult_Cof = r_dbl_CapPen_Cof - ((r_dbl_CapPen_Cof * 6) - r_dbl_CapCof)
         End If
         
         'Ajustando Interes COFIDE Ultima Cuota
         If r_dbl_IntPen_Cof * r_int_CuoDis < r_dbl_IntCof Then
            r_dbl_IntPen_Ult_Cof = r_dbl_IntPen_Cof + (r_dbl_IntCof - (r_dbl_IntPen_Cof * 6))
         Else
            r_dbl_IntPen_Ult_Cof = r_dbl_IntPen_Cof - ((r_dbl_IntPen_Cof * 6) - r_dbl_IntCof)
         End If
         
         'Ajustando Interes COFIDE Ultima Cuota
         If r_dbl_ComPen_Cof * r_int_CuoDis < r_dbl_ComCof Then
            r_dbl_ComPen_Ult_Cof = r_dbl_ComPen_Cof + (r_dbl_ComCof - (r_dbl_ComPen_Cof * 6))
         Else
            r_dbl_ComPen_Ult_Cof = r_dbl_ComPen_Cof - ((r_dbl_ComPen_Cof * 6) - r_dbl_ComCof)
         End If
         
         If r_int_CuoDis = 6 Then
            r_dbl_CapPen_Cli_1 = r_dbl_CapPen_Cli:    r_dbl_CapPen_Cli_2 = r_dbl_CapPen_Cli:    r_dbl_CapPen_Cli_3 = r_dbl_CapPen_Cli:    r_dbl_CapPen_Cli_4 = r_dbl_CapPen_Cli:    r_dbl_CapPen_Cli_5 = r_dbl_CapPen_Cli:    r_dbl_CapPen_Cli_6 = r_dbl_CapPen_Ult_Cli
            r_dbl_IntPen_Cli_1 = r_dbl_IntPen_Cli:    r_dbl_IntPen_Cli_2 = r_dbl_IntPen_Cli:    r_dbl_IntPen_Cli_3 = r_dbl_IntPen_Cli:    r_dbl_IntPen_Cli_4 = r_dbl_IntPen_Cli:    r_dbl_IntPen_Cli_5 = r_dbl_IntPen_Cli:    r_dbl_IntPen_Cli_6 = r_dbl_IntPen_Ult_Cli
         ElseIf r_int_CuoDis = 5 Then
            r_dbl_CapPen_Cli_2 = r_dbl_CapPen_Cli:    r_dbl_CapPen_Cli_3 = r_dbl_CapPen_Cli:    r_dbl_CapPen_Cli_4 = r_dbl_CapPen_Cli:    r_dbl_CapPen_Cli_5 = r_dbl_CapPen_Cli:    r_dbl_CapPen_Cli_6 = r_dbl_CapPen_Ult_Cli
            r_dbl_IntPen_Cli_2 = r_dbl_IntPen_Cli:    r_dbl_IntPen_Cli_3 = r_dbl_IntPen_Cli:    r_dbl_IntPen_Cli_4 = r_dbl_IntPen_Cli:    r_dbl_IntPen_Cli_5 = r_dbl_IntPen_Cli:    r_dbl_IntPen_Cli_6 = r_dbl_IntPen_Ult_Cli
         ElseIf r_int_CuoDis = 4 Then
            r_dbl_CapPen_Cli_3 = r_dbl_CapPen_Cli:    r_dbl_CapPen_Cli_4 = r_dbl_CapPen_Cli:    r_dbl_CapPen_Cli_5 = r_dbl_CapPen_Cli:    r_dbl_CapPen_Cli_6 = r_dbl_CapPen_Ult_Cli
            r_dbl_IntPen_Cli_3 = r_dbl_IntPen_Cli:    r_dbl_IntPen_Cli_4 = r_dbl_IntPen_Cli:    r_dbl_IntPen_Cli_5 = r_dbl_IntPen_Cli:    r_dbl_IntPen_Cli_6 = r_dbl_IntPen_Ult_Cli
         ElseIf r_int_CuoDis = 3 Then
            r_dbl_CapPen_Cli_4 = r_dbl_CapPen_Cli:    r_dbl_CapPen_Cli_5 = r_dbl_CapPen_Cli:    r_dbl_CapPen_Cli_6 = r_dbl_CapPen_Ult_Cli
            r_dbl_IntPen_Cli_4 = r_dbl_IntPen_Cli:    r_dbl_IntPen_Cli_5 = r_dbl_IntPen_Cli:    r_dbl_IntPen_Cli_6 = r_dbl_IntPen_Ult_Cli
         ElseIf r_int_CuoDis = 2 Then
            r_dbl_CapPen_Cli_5 = r_dbl_CapPen_Cli:    r_dbl_CapPen_Cli_6 = r_dbl_CapPen_Ult_Cli
            r_dbl_IntPen_Cli_5 = r_dbl_IntPen_Cli:    r_dbl_IntPen_Cli_6 = r_dbl_IntPen_Ult_Cli
         ElseIf r_int_CuoDis = 1 Then
            r_dbl_CapPen_Cli_6 = r_dbl_CapPen_Ult_Cli
            r_dbl_IntPen_Cli_6 = r_dbl_IntPen_Ult_Cli
         End If
      ElseIf r_int_FlgPBP = 1 Then
         r_int_EvaAsg = r_int_EvaAsg + 1
      End If
      
      r_int_NumEva = r_int_NumEva + 1
         
      'Grabando en CRE_DETPBP
      g_str_Parame = "USP_CRE_DETPBP ("
      g_str_Parame = g_str_Parame & CStr(p_PerMes) & ", "
      g_str_Parame = g_str_Parame & CStr(p_PerAno) & ", "
      g_str_Parame = g_str_Parame & "'" & g_rst_Princi!HIPCUO_NUMOPE & "', "
      g_str_Parame = g_str_Parame & CStr(r_int_FlgPBP) & ", "
      g_str_Parame = g_str_Parame & CStr(r_int_NumCuo) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_CapCli) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_IntCli) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_CapCof) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_IntCof) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_ComCof) & ", "
      g_str_Parame = g_str_Parame & CStr(g_rst_Princi!HIPCUO_SALCAP + g_rst_Princi!HIPCUO_CAPITA) & ", "
      g_str_Parame = g_str_Parame & CStr(g_rst_Princi!HIPCUO_SALCAP) & ", "
      g_str_Parame = g_str_Parame & CStr(r_int_CuoIni_Eval) & ", "
      g_str_Parame = g_str_Parame & CStr(r_int_CuoFin_Eval) & ", "
      g_str_Parame = g_str_Parame & CStr(r_int_CuoIni_Cast) & ", "
      g_str_Parame = g_str_Parame & CStr(r_int_CuoFin_Cast) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_CapPen_Cli_1) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_IntPen_Cli_1) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_CapPen_Cli_2) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_IntPen_Cli_2) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_CapPen_Cli_3) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_IntPen_Cli_3) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_CapPen_Cli_4) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_IntPen_Cli_4) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_CapPen_Cli_5) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_IntPen_Cli_5) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_CapPen_Cli_6) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_IntPen_Cli_6) & ", "
      g_str_Parame = g_str_Parame & CStr(r_int_CuoIni_Cast) & ", "
      g_str_Parame = g_str_Parame & CStr(r_int_CuoFin_Cast) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_CapPen_Cof) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_IntPen_Cof) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_ComPen_Cof) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_CapPen_Cof) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_IntPen_Cof) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_ComPen_Cof) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_CapPen_Cof) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_IntPen_Cof) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_ComPen_Cof) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_CapPen_Cof) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_IntPen_Cof) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_ComPen_Cof) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_CapPen_Cof) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_IntPen_Cof) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_ComPen_Cof) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_CapPen_Ult_Cof) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_IntPen_Ult_Cof) & ", "
      g_str_Parame = g_str_Parame & CStr(r_dbl_ComPen_Ult_Cof) & ", "
      
      'Datos de Auditoria
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
      g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
      g_str_Parame = g_str_Parame & "1)"
            
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         MsgBox "No se pudo ejecutar el procedimiento USP_CRE_DETPBP.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
      
      g_rst_Princi.MoveNext
   Loop
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Grabando Cabecera
   g_str_Parame = "USP_CRE_CABPBP ("
   g_str_Parame = g_str_Parame & CStr(p_PerMes) & ", "
   g_str_Parame = g_str_Parame & CStr(p_PerAno) & ", "
   g_str_Parame = g_str_Parame & CStr(r_int_NumEva) & ", "
   g_str_Parame = g_str_Parame & CStr(r_int_EvaAsg) & ", "
   g_str_Parame = g_str_Parame & CStr(r_int_EvaPer) & ", "
   g_str_Parame = g_str_Parame & "1, "
   
   'Datos de Auditoria
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "', "
   g_str_Parame = g_str_Parame & "'" & UCase(App.EXEName) & "', "
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodSuc & "', "
   g_str_Parame = g_str_Parame & "1)"
         
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
      MsgBox "No se pudo ejecutar el procedimiento USP_CRE_CABPBP.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
End Sub
