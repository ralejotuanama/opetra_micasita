VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_Caj_GenArc_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3960
   ClientLeft      =   6000
   ClientTop       =   3495
   ClientWidth     =   7875
   Icon            =   "OpeTra_frm_064.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   7875
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   3975
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   7905
      _Version        =   65536
      _ExtentX        =   13944
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
      Begin Threed.SSPanel SSPanel3 
         Height          =   645
         Left            =   30
         TabIndex        =   6
         Top             =   780
         Width           =   7815
         _Version        =   65536
         _ExtentX        =   13785
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
         Begin VB.CommandButton cmd_Import 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_064.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   7200
            Picture         =   "OpeTra_frm_064.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   2445
         Left            =   30
         TabIndex        =   7
         Top             =   1470
         Width           =   7815
         _Version        =   65536
         _ExtentX        =   13785
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
         Begin VB.DirListBox dir_LisCar 
            Height          =   1665
            Left            =   1560
            TabIndex        =   2
            Top             =   720
            Width           =   6195
         End
         Begin VB.DriveListBox drv_LisUni 
            Height          =   315
            Left            =   1560
            TabIndex        =   1
            Top             =   390
            Width           =   6195
         End
         Begin VB.ComboBox cmb_CodBan 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   6195
         End
         Begin VB.Label Label3 
            Caption         =   "Banco:"
            Height          =   315
            Left            =   60
            TabIndex        =   9
            Top             =   60
            Width           =   765
         End
         Begin VB.Label Label1 
            Caption         =   "Carpeta a guardar archivos:"
            Height          =   615
            Left            =   60
            TabIndex        =   8
            Top             =   390
            Width           =   1365
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   705
         Left            =   30
         TabIndex        =   10
         Top             =   30
         Width           =   7815
         _Version        =   65536
         _ExtentX        =   13785
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
            Height          =   315
            Left            =   660
            TabIndex        =   11
            Top             =   60
            Width           =   3315
            _Version        =   65536
            _ExtentX        =   5847
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
         Begin Threed.SSPanel SSPanel4 
            Height          =   315
            Left            =   660
            TabIndex        =   12
            Top             =   360
            Width           =   4005
            _Version        =   65536
            _ExtentX        =   7064
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Generación de Archivo de Recaudación"
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
         Begin MSComDlg.CommonDialog dlg_Guarda 
            Left            =   7200
            Top             =   90
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   6720
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
            WindowShowRefreshBtn=   -1  'True
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   60
            Picture         =   "OpeTra_frm_064.frx":0890
            Top             =   60
            Width           =   480
         End
      End
   End
End
Attribute VB_Name = "frm_Caj_GenArc_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_arr_CodBan()      As moddat_tpo_Genera
Dim l_dbl_PorITF        As Double
Dim l_lng_Contad        As Long
Dim l_lng_ConTem        As Long
Dim l_lng_ConAux        As Long
Dim l_lng_TemAux        As Long
Dim l_int_CanSol        As Integer
Dim l_int_CanDol        As Integer
Dim l_str_CodCuo()      As String
Dim l_lng_NumOpe        As Long
Dim l_str_DebAut()      As String
Dim l_str_NomArc        As String
Dim l_str_CodUni        As String
Dim l_str_CodRub        As String
Dim l_str_CodEmp        As String
Dim l_str_CodSer        As String
Dim l_str_CodSol        As String
Dim l_bln_FlgPro        As Boolean

Private Sub cmb_CodBan_Click()
   Call gs_SetFocus(drv_LisUni)
End Sub

Private Sub cmb_CodBan_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_CodBan_Click
   End If
End Sub

Private Sub cmd_Import_Click()
   Dim r_int_FlgSol     As Integer
   Dim r_int_FlgDol     As Integer

   If cmb_CodBan.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Banco.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_CodBan)
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de ejecutar el proceso?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Select Case l_arr_CodBan(cmb_CodBan.ListIndex + 1).Genera_Codigo
      Case "000002"  'BBVA Banco Continental
         Screen.MousePointer = 11
         r_int_FlgSol = ff_GenArc_BBVA(dir_LisCar.Path, 1)
         r_int_FlgDol = ff_GenArc_BBVA(dir_LisCar.Path, 2)
         Screen.MousePointer = 0
         
         If r_int_FlgSol = True And r_int_FlgDol = True Then
            MsgBox "Se generaron los archivos para Recaudación en SOLES y DOLARES AMERICANOS.", vbInformation, modgen_g_str_NomPlt
         ElseIf r_int_FlgSol = True And r_int_FlgDol = False Then
            MsgBox "Se generó el archivo para Recaudación en SOLES.", vbInformation, modgen_g_str_NomPlt
         ElseIf r_int_FlgSol = False And r_int_FlgDol = True Then
            MsgBox "Se generó el archivo para Recaudación en DOLARES AMERICANOS.", vbInformation, modgen_g_str_NomPlt
         Else
            MsgBox "No se encontraron datos para generar Archivos de Recaudación.", vbInformation, modgen_g_str_NomPlt
         End If
         
      Case "000004"  'Banco Interbank
         l_str_NomArc = "0727501"         'Nombre Archivo Plano
         l_str_CodUni = "0011347855"      'Código único
         l_str_CodRub = "07"              'Código de rubro
         l_str_CodEmp = "275"             'Código de empresa
         l_str_CodSer = "01"              'Código de servicio
         l_str_CodSol = "01"              'Código de solicitud
   
         Screen.MousePointer = 11
         l_bln_FlgPro = ff_GenArc_Interbank(dir_LisCar.Path)
         Screen.MousePointer = 0
         
         If l_int_CanSol > 0 And l_int_CanDol > 0 Then
            MsgBox "Se generaron los archivos para Recaudación en SOLES y DOLARES AMERICANOS.", vbInformation, modgen_g_str_NomPlt
         ElseIf l_int_CanSol > 0 And l_int_CanDol = 0 Then
            MsgBox "Se generó el archivo para Recaudación en SOLES.", vbInformation, modgen_g_str_NomPlt
         ElseIf l_int_CanSol = 0 And l_int_CanDol > 0 Then
            MsgBox "Se generó el archivo para Recaudación en DOLARES AMERICANOS.", vbInformation, modgen_g_str_NomPlt
         Else
            MsgBox "No se encontraron datos para generar Archivos de Recaudación.", vbInformation, modgen_g_str_NomPlt
         End If
         
   End Select
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub drv_LisUni_Change()
   dir_LisCar.Path = drv_LisUni.Drive
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_LisIte(cmb_CodBan, l_arr_CodBan, 1, "516")
   
   'Obteniendo ITF
   l_dbl_PorITF = opecaj_gf_Consulta_ITF(Format(CDate(moddat_g_str_FecSis), "yyyymmdd"), 1)
   
   drv_LisUni.Drive = "C:"
End Sub

Private Function ff_GenArc_BBVA(ByVal p_RutArc As String, ByVal p_TipMon As Integer) As Integer
   Dim r_int_NumFil     As Integer
   Dim r_str_NumSol     As String
   Dim r_dbl_ImpITF     As Double
   Dim r_dbl_TotGas     As Double
   Dim r_str_NomCli     As String
   Dim r_str_NumRef     As String
   Dim r_str_NumOpe     As String
   Dim r_int_NumCuo     As Integer
   Dim r_dbl_TotCuo     As Double
   Dim r_int_ConLin     As Double
   Dim r_dbl_ImpMax     As Double
   Dim r_dbl_ImpMin     As Double
   Dim r_str_CodCla     As String
   Dim r_int_FlgDat     As Integer
   
   r_int_ConLin = 0
   r_dbl_ImpMax = 0
   r_dbl_ImpMin = 0
   r_int_FlgDat = 0
   r_int_NumFil = FreeFile
   
   If p_TipMon = 2 Then
      r_str_CodCla = "421"
   Else
      r_str_CodCla = "420"
   End If
   
   'Crea Archivo de Pagos (Cabecera)
   Open p_RutArc & "\R" & r_str_CodCla & Format(date, "mmdd") & ".TXT" For Output As r_int_NumFil
   Print #r_int_NumFil, "01" & "20511904162" & r_str_CodCla & IIf(p_TipMon = 1, "PEN", "USD") & Format(date, "yyyymmdd") & "000" & Space(330)
   
   '***********************************************
   'OBTIENE LOS GASTOS DE CIERRE PENDIENTES DE PAGO
   g_str_Parame = "SELECT GASADM_NUMSOL, SUM(GASADM_IMPORT) AS TOTGAS FROM TRA_GASADM WHERE "
   g_str_Parame = g_str_Parame & "GASADM_TIPMON = " & CStr(p_TipMon) & " AND "
   g_str_Parame = g_str_Parame & "GASADM_SITUAC = 2 "
   g_str_Parame = g_str_Parame & "GROUP BY GASADM_NUMSOL "
   g_str_Parame = g_str_Parame & "ORDER BY GASADM_NUMSOL ASC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      r_int_FlgDat = 1
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         r_str_NumSol = g_rst_Princi!GASADM_NUMSOL
         
         'Verificando que Solicitud este en Trámite
         g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
         g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = '" & r_str_NumSol & "' AND "
         g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1"
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
            Exit Function
         End If
         
         If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
            g_rst_Genera.MoveFirst
            r_dbl_ImpITF = gf_NueImp_Numero(gf_Truncar_Numero(g_rst_Princi!TOTGAS * (l_dbl_PorITF / 100), 2))
            r_dbl_TotGas = CDbl(Format(g_rst_Princi!TOTGAS + r_dbl_ImpITF, "###,###,##0.00"))
            r_str_NomCli = moddat_gf_Buscar_NomCli(g_rst_Genera!SOLMAE_TITTDO, g_rst_Genera!SOLMAE_TITNDO) & Space(30)
            r_str_NumRef = gf_Formato_NumSol(g_rst_Genera!SOLMAE_NUMERO)
            
            Print #r_int_NumFil, "02" & _
                                 Mid(r_str_NomCli, 1, 30) & _
                                 Format(CStr(g_rst_Genera!SOLMAE_TITTDO) & Trim(g_rst_Genera!SOLMAE_TITNDO), "0000000000000") & _
                                 Mid(g_rst_Genera!SOLMAE_NUMERO & Space(20), 1, 20) & "01" & _
                                 "001" & Space(10) & _
                                 Format(date, "yyyymmdd") & Format(date + CDate(4), "yyyymmdd") & "00" & _
                                 Left(Format(r_dbl_TotGas, "0000000000000.00"), 13) & Right(Format(r_dbl_TotGas, "0000000000000.00"), 2) & _
                                 Left(Format(r_dbl_TotGas, "0000000000000.00"), 13) & Right(Format(r_dbl_TotGas, "0000000000000.00"), 2) & _
                                 String(32, "0") & _
                                 "00" & String(14, "0") & _
                                 "00" & String(14, "0") & _
                                 "00" & String(14, "0") & _
                                 "00" & String(14, "0") & _
                                 "05" & Left(Format(r_dbl_TotGas, "000000000000.00"), 12) & Right(Format(r_dbl_TotGas, "000000000000.00"), 2) & _
                                 "00" & String(14, "0") & _
                                 "00" & String(14, "0") & _
                                 "00" & String(14, "0") & _
                                 String(20, "0") & _
                                 IIf(g_rst_Genera!SOLMAE_TITTDO = 1, "L", "E") & Mid(Trim(g_rst_Genera!SOLMAE_TITNDO) & Space(15), 1, 15) & _
                                 Space(36)
            
            r_int_ConLin = r_int_ConLin + 1
            r_dbl_ImpMax = r_dbl_ImpMax + r_dbl_TotGas
            r_dbl_ImpMin = r_dbl_ImpMin + r_dbl_TotGas
         End If
         
         g_rst_Genera.Close
         Set g_rst_Genera = Nothing
         g_rst_Princi.MoveNext
         DoEvents
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   '*****************************************************************
   'OBTIENE LAS CUOTAS PENDIENTE DE PAGO DE LOS CREDITOS HIPOTECARIOS
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM CRE_HIPMAE "
   g_str_Parame = g_str_Parame & " WHERE HIPMAE_SITUAC = 2 "
   g_str_Parame = g_str_Parame & "   AND HIPMAE_JUDICI = 0 "
   g_str_Parame = g_str_Parame & "   AND HIPMAE_MONEDA = " & CStr(p_TipMon) & " "
   g_str_Parame = g_str_Parame & "   AND (HIPMAE_ENVCUO = 1 OR HIPMAE_ENVCUO IS NULL) "
   g_str_Parame = g_str_Parame & "ORDER BY HIPMAE_NUMOPE ASC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      r_int_FlgDat = 1
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         r_str_NumOpe = g_rst_Princi!HIPMAE_NUMOPE
         g_str_Parame = "SELECT * FROM CRE_HIPCUO WHERE "
         g_str_Parame = g_str_Parame & "HIPCUO_NUMOPE = '" & r_str_NumOpe & "' AND "
         g_str_Parame = g_str_Parame & "HIPCUO_TIPCRO = 1 AND "
         g_str_Parame = g_str_Parame & "HIPCUO_SITUAC = 2 "
         g_str_Parame = g_str_Parame & "ORDER BY HIPCUO_NUMCUO ASC "
      
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
            Exit Function
         End If
         
         If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
            g_rst_Genera.MoveFirst
            r_str_NumRef = gf_Formato_NumOpe(g_rst_Genera!HIPCUO_NUMOPE)
            r_str_NomCli = moddat_gf_Buscar_NomCli(g_rst_Princi!HIPMAE_TDOCLI, g_rst_Princi!HIPMAE_NDOCLI) & Space(30)
            r_str_NomCli = Replace((r_str_NomCli), "Ü", "U")
            r_int_NumCuo = 0
            r_dbl_TotGas = 0
            
            Do While Not g_rst_Genera.EOF
               r_dbl_TotCuo = g_rst_Genera!HIPCUO_CAPITA + g_rst_Genera!HIPCUO_INTERE + g_rst_Genera!HIPCUO_DESORG + g_rst_Genera!HIPCUO_VIVORG + g_rst_Genera!HIPCUO_OTRORG
               r_dbl_TotCuo = r_dbl_TotCuo + g_rst_Genera!HIPCUO_INTCOM + g_rst_Genera!HIPCUO_INTMOR + g_rst_Genera!HIPCUO_GASCOB + g_rst_Genera!HIPCUO_OTRGAS
               r_dbl_TotCuo = r_dbl_TotCuo + g_rst_Genera!HIPCUO_CAPBBP + g_rst_Genera!HIPCUO_INTBBP
               r_dbl_TotCuo = r_dbl_TotCuo - g_rst_Genera!HIPCUO_IMPPAG
            
               'No toma en cuentas las cuotas que son periodo de gracia.
               If r_dbl_TotCuo > 0 Then
                  If g_rst_Princi!HIPMAE_INDITF = 1 Then
                     r_dbl_ImpITF = gf_Truncar_Numero(r_dbl_TotCuo * (l_dbl_PorITF / 100), 2)
                  Else
                     r_dbl_ImpITF = 0
                  End If
                  r_dbl_TotCuo = CDbl(Format(r_dbl_TotCuo + r_dbl_ImpITF, "###,###,##0.00"))
                  r_dbl_TotGas = r_dbl_TotGas + r_dbl_TotCuo
               
                  Print #r_int_NumFil, "02" & _
                                       Mid(r_str_NomCli, 1, 30) & _
                                       Format(CStr(g_rst_Princi!HIPMAE_TDOCLI) & Trim(g_rst_Princi!HIPMAE_NDOCLI), "0000000000000") & _
                                       Mid(g_rst_Princi!HIPMAE_NUMOPE & Space(20), 1, 20) & "02" & _
                                       Format(g_rst_Genera!HIPCUO_NUMCUO, "000") & Space(10) & _
                                       CStr(g_rst_Genera!HIPCUO_FECVCT) & _
                                       IIf(CDate(gf_FormatoFecha(CStr(g_rst_Genera!HIPCUO_FECVCT))) < date, Format(date + CDate(4), "yyyymmdd"), Format(CDate(gf_FormatoFecha(CStr(g_rst_Genera!HIPCUO_FECVCT))) + CDate(4), "yyyymmdd")) & _
                                       Mid(CStr(g_rst_Genera!HIPCUO_FECVCT), 5, 2) & _
                                       Left(Format(r_dbl_TotCuo, "0000000000000.00"), 13) & Right(Format(r_dbl_TotCuo, "0000000000000.00"), 2) & _
                                       Left(Format(r_dbl_TotCuo, "0000000000000.00"), 13) & Right(Format(r_dbl_TotCuo, "0000000000000.00"), 2) & _
                                       String(32, "0") & _
                                       "01" & Left(Format(r_dbl_TotCuo, "000000000000.00"), 12) & Right(Format(r_dbl_TotCuo, "000000000000.00"), 2) & _
                                       "00" & String(14, "0") & _
                                       "00" & String(14, "0") & _
                                       "00" & String(14, "0") & _
                                       "00" & String(14, "0") & _
                                       "00" & String(14, "0") & _
                                       "00" & String(14, "0") & _
                                       "00" & String(14, "0") & _
                                       String(20, "0") & _
                                       IIf(g_rst_Princi!HIPMAE_TDOCLI = 1, "L", "E") & Mid(Trim(g_rst_Princi!HIPMAE_NDOCLI) & Space(15), 1, 15) & _
                                       Space(36)
               
                  If Trim(g_rst_Genera!HIPCUO_FECVCT) >= Format(date, "YYYYMMDD") Then
                     r_int_NumCuo = r_int_NumCuo + 1
                  End If
                  
                  r_int_ConLin = r_int_ConLin + 1
                  r_dbl_ImpMax = r_dbl_ImpMax + r_dbl_TotCuo
                  r_dbl_ImpMin = r_dbl_ImpMin + r_dbl_TotCuo
               End If
               
               g_rst_Genera.MoveNext
               If r_int_NumCuo = 2 Then
                  Exit Do
               End If
            Loop
         End If
      
         g_rst_Genera.Close
         Set g_rst_Genera = Nothing
         g_rst_Princi.MoveNext
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   '********************************************************
   'OBTIENE LAS CUOTAS PENDIENTE DE PAGO DEL PLAN DE AHORROS
   g_str_Parame = "SELECT A.AHOMAE_NUMERO,C.AHOCLI_TIPDOC,C.AHOCLI_NUMDOC,C.AHOCLI_APEPAT,C.AHOCLI_APEMAT,C.AHOCLI_NOMBRE  "
   g_str_Parame = g_str_Parame & "FROM CRE_AHOMAE A, CRE_AHOCLI C WHERE "
   g_str_Parame = g_str_Parame & "A.AHOMAE_TIPDOC = C.AHOCLI_TIPDOC AND A.AHOMAE_NUMDOC = C.AHOCLI_NUMDOC AND "
   g_str_Parame = g_str_Parame & "A.AHOMAE_SITUAC = '2' AND "
   g_str_Parame = g_str_Parame & "A.AHOMAE_MONAHO = " & CStr(p_TipMon) & " "
   g_str_Parame = g_str_Parame & "ORDER BY A.AHOMAE_NUMERO "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      r_int_FlgDat = 1
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
      
           g_str_Parame = "SELECT * FROM CRE_AHOCUO WHERE "
           g_str_Parame = g_str_Parame & "TRIM(AHOCUO_NUMERO) = '" & Trim(g_rst_Princi!AHOMAE_NUMERO) & "' "
           g_str_Parame = g_str_Parame & "AND AHOCUO_SITUAC = '2' "
           g_str_Parame = g_str_Parame & "ORDER BY AHOCUO_NUMCUO "
        
           If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
              Exit Function
           End If
           
           If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
              g_rst_Genera.MoveFirst
              r_str_NomCli = Trim(g_rst_Princi!AHOCLI_APEPAT) & " " & Trim(g_rst_Princi!AHOCLI_APEMAT) & " " & Trim(g_rst_Princi!AHOCLI_NOMBRE) & Space(30)
              r_int_NumCuo = 0
              r_dbl_TotGas = 0
              
              Do While Not g_rst_Genera.EOF
                 r_str_NumOpe = Trim(g_rst_Princi!AHOMAE_NUMERO)
                 r_dbl_TotCuo = g_rst_Genera!AHOCUO_CAPITA
                 r_dbl_ImpITF = 0
                 r_dbl_TotCuo = CDbl(Format(r_dbl_TotCuo + r_dbl_ImpITF, "###,###,##0.00"))
                 r_dbl_TotGas = r_dbl_TotGas + r_dbl_TotCuo
                 
                 Print #r_int_NumFil, "02" & _
                                      Mid(r_str_NomCli, 1, 30) & _
                                      Format(CStr(g_rst_Princi!AHOCLI_TIPDOC) & Trim(g_rst_Princi!AHOCLI_NUMDOC), "0000000000000") & _
                                      Mid(r_str_NumOpe & Space(20), 1, 20) & "05" & _
                                      Format(g_rst_Genera!AHOCUO_NUMCUO, "000") & Space(10) & _
                                      CStr(Trim(g_rst_Genera!AHOCUO_FECVCT)) & _
                                      IIf(CDate(gf_FormatoFecha(CStr(Trim(g_rst_Genera!AHOCUO_FECVCT)))) < date, Format(date + CDate(4), "yyyymmdd"), Format(CDate(gf_FormatoFecha(CStr(Trim(g_rst_Genera!AHOCUO_FECVCT)))) + CDate(4), "yyyymmdd")) & _
                                      Mid(CStr(g_rst_Genera!AHOCUO_FECVCT), 5, 2) & _
                                      Left(Format(r_dbl_TotCuo, "0000000000000.00"), 13) & Right(Format(r_dbl_TotCuo, "0000000000000.00"), 2) & _
                                      Left(Format(r_dbl_TotCuo, "0000000000000.00"), 13) & Right(Format(r_dbl_TotCuo, "0000000000000.00"), 2) & _
                                      String(32, "0") & _
                                      "01" & Left(Format(r_dbl_TotCuo, "000000000000.00"), 12) & Right(Format(r_dbl_TotCuo, "000000000000.00"), 2) & _
                                      "00" & String(14, "0") & _
                                      "00" & String(14, "0") & _
                                      "00" & String(14, "0") & _
                                      "00" & String(14, "0") & _
                                      "00" & String(14, "0") & _
                                      "00" & String(14, "0") & _
                                      "00" & String(14, "0") & _
                                      String(20, "0") & _
                                      IIf(g_rst_Princi!AHOCLI_TIPDOC = 1, "L", "E") & Mid(Trim(g_rst_Princi!AHOCLI_NUMDOC) & Space(15), 1, 15) & _
                                      Space(36)
              
                 r_int_ConLin = r_int_ConLin + 1
                 r_dbl_ImpMax = r_dbl_ImpMax + r_dbl_TotCuo
                 r_dbl_ImpMin = r_dbl_ImpMin + r_dbl_TotCuo
                 g_rst_Genera.MoveNext
                 r_int_NumCuo = r_int_NumCuo + 1
                 
                 If r_int_NumCuo = 6 Then
                    Exit Do
                 End If
              Loop
           End If
        
           g_rst_Genera.Close
           Set g_rst_Genera = Nothing
           
      g_rst_Princi.MoveNext
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Cerrando Archivo de Pagos
   Print #r_int_NumFil, "03" & Format(r_int_ConLin, "000000000") & _
         Left(Format(r_dbl_ImpMax, "0000000000000000.00"), 16) & Right(Format(r_dbl_ImpMax, "0000000000000000.00"), 2) & _
         Left(Format(r_dbl_ImpMin, "0000000000000000.00"), 16) & Right(Format(r_dbl_ImpMin, "0000000000000000.00"), 2) & _
         String(18, "0") & Space(295)
   
   Close #r_int_NumFil
   
   If r_int_FlgDat = 0 Then
      Kill p_RutArc & "\R" & r_str_CodCla & Format(date, "mmdd") & ".TXT"
      ff_GenArc_BBVA = False
   Else
      ff_GenArc_BBVA = True
   End If
End Function

Private Function ff_GenArc_Interbank(ByVal p_RutArc As String) As Integer
   Dim r_int_NumFil     As Integer
   Dim r_str_NumSol     As String
   Dim r_dbl_ImpITF     As Double
   Dim r_dbl_TotGas     As Double
   Dim r_str_NomCli     As String
   Dim r_str_NumRef     As String
   Dim r_str_NumOpe     As String
   Dim r_int_NumCuo     As Integer
   Dim r_dbl_TotCuo     As Double
   Dim r_int_ConLin     As Double
   Dim r_dbl_ImpMax     As Double
   Dim r_dbl_ImpMin     As Double
   Dim r_str_CodCla     As String
   Dim r_int_FlgDat     As Integer
   Dim r_dbl_MtoSol     As Double
   Dim r_dbl_MtoDol     As Double
   Dim r_int_TipMon     As Integer
   Dim r_str_Concep     As String
   
   r_int_ConLin = 0
   r_dbl_ImpMax = 0
   r_dbl_ImpMin = 0
   r_int_FlgDat = 0
   l_lng_Contad = -1
   l_lng_ConTem = 21
   l_lng_ConAux = 0
   l_int_CanSol = 0
   l_int_CanDol = 0
   r_int_FlgDat = 0
   r_int_NumFil = FreeFile
         
   'Obteniendo Gastos de Cierre Pendientes de Pago
   g_str_Parame = "SELECT GASADM_NUMSOL, SUM(GASADM_IMPORT) AS TOTGAS FROM TRA_GASADM WHERE "
   g_str_Parame = g_str_Parame & "GASADM_SITUAC = 2 "
   g_str_Parame = g_str_Parame & "GROUP BY GASADM_NUMSOL "
   g_str_Parame = g_str_Parame & "ORDER BY GASADM_NUMSOL ASC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      r_int_FlgDat = 1
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         r_str_NumSol = g_rst_Princi!GASADM_NUMSOL
         
         'Verificando que Solicitud este en Trámite
         g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
         g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = '" & r_str_NumSol & "' AND "
         g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1"
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
            Exit Function
         End If
         
         If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
            g_rst_Genera.MoveFirst
            
            r_dbl_TotGas = 0
            l_lng_TemAux = 0
            r_dbl_ImpITF = gf_NueImp_Numero(gf_Truncar_Numero(g_rst_Princi!TOTGAS * (l_dbl_PorITF / 100), 2))
            r_dbl_TotGas = CDbl(Format(g_rst_Princi!TOTGAS + r_dbl_ImpITF, "###,###,##0.00"))
            r_str_NomCli = moddat_gf_Buscar_NomCli(g_rst_Genera!SOLMAE_TITTDO, g_rst_Genera!SOLMAE_TITNDO)
            'r_str_NumRef = gf_Formato_NumSol(g_rst_Genera!SOLMAE_NUMERO)
                                                         
            If l_lng_Contad = -1 Then
               ReDim Preserve l_str_CodCuo(0) As String
               l_str_CodCuo(0) = Left(Format(date + CDate(4), "yyyymmdd"), 6) & Trim(g_rst_Genera!SOLMAE_TIPMON) & "5"
            Else
               For l_lng_ConAux = 0 To UBound(l_str_CodCuo) Step 1
                  If l_str_CodCuo(l_lng_ConAux) = Left(Format(date + CDate(4), "yyyymmdd"), 6) & Trim(g_rst_Genera!SOLMAE_TIPMON) & "5" Then
                     l_lng_TemAux = 1
                  End If
               Next
               
               If l_lng_TemAux = 0 Then
                  ReDim Preserve l_str_CodCuo(UBound(l_str_CodCuo) + 1) As String
                  l_str_CodCuo(UBound(l_str_CodCuo)) = Left(Format(date + CDate(4), "yyyymmdd"), 6) & Trim(g_rst_Genera!SOLMAE_TIPMON) & "5"
               End If
            End If
         
            ReDim Preserve l_str_DebAut(l_lng_ConTem + l_lng_NumOpe) As String
            l_str_DebAut(l_lng_Contad + 1) = "13"
            l_str_DebAut(l_lng_Contad + 2) = "" & gs_modsec_Genera(Format(Trim(g_rst_Genera!SOLMAE_TITTDO) & Trim(g_rst_Genera!SOLMAE_TITNDO), "########0"), 2, " ", 20)
            l_str_DebAut(l_lng_Contad + 3) = "" & gs_modsec_Genera(Left(Format(date + CDate(4), "yyyymmdd"), 4) & "-" & Mid(Format(date + CDate(4), "yyyymmdd"), 5, 2) & "5", 2, " ", 8)
            l_str_DebAut(l_lng_Contad + 4) = "" & gs_modsec_Genera(Left(r_str_NomCli, 30), 2, " ", 30)
            l_str_DebAut(l_lng_Contad + 5) = "" & String(10, "0")
            l_str_DebAut(l_lng_Contad + 6) = "" & String(10, "0")
            l_str_DebAut(l_lng_Contad + 7) = "" & Format(date, "yyyymmdd")
            l_str_DebAut(l_lng_Contad + 8) = "" & Format(date + CDate(4), "yyyymmdd")
            l_str_DebAut(l_lng_Contad + 9) = "" & gs_modsec_Genera(Format(Trim(g_rst_Genera!SOLMAE_NUMERO), "###########0"), 1, "0", 15)
            l_str_DebAut(l_lng_Contad + 10) = "" & IIf(CInt(g_rst_Genera!SOLMAE_TIPMON) = 1, "01", "10")
            l_str_DebAut(l_lng_Contad + 11) = "" & String(9, "0")
            l_str_DebAut(l_lng_Contad + 12) = "" & String(9, "0")
            l_str_DebAut(l_lng_Contad + 13) = "" & String(9, "0")
            l_str_DebAut(l_lng_Contad + 14) = "" & String(9, "0")
            l_str_DebAut(l_lng_Contad + 15) = "" & gs_modsec_Genera(Format(r_dbl_TotGas, "########0.00"), 1, "0", 9)
            l_str_DebAut(l_lng_Contad + 16) = "" & String(9, "0")
            l_str_DebAut(l_lng_Contad + 17) = "" & String(9, "0")
            l_str_DebAut(l_lng_Contad + 18) = "" & String(2, "  ")
            l_str_DebAut(l_lng_Contad + 19) = "" & "A" '& "M" & "E" & "D"
            l_str_DebAut(l_lng_Contad + 20) = "" & String(13, "  ")
            l_str_DebAut(l_lng_Contad + 21) = "" & String(8, "0")
            l_str_DebAut(l_lng_Contad + 22) = "" & CInt(g_rst_Genera!SOLMAE_TIPMON)
                                 
            l_lng_Contad = l_lng_Contad + 22
            l_lng_ConTem = l_lng_ConTem + 21
            
            If g_rst_Genera!SOLMAE_TIPMON = 1 Then
               l_int_CanSol = l_int_CanSol + 1
               r_dbl_MtoSol = r_dbl_MtoSol + r_dbl_TotGas
            ElseIf g_rst_Genera!SOLMAE_TIPMON = 2 Then
               l_int_CanDol = l_int_CanDol + 1
               r_dbl_MtoDol = r_dbl_MtoDol + r_dbl_TotGas
            End If
            
            l_lng_NumOpe = l_lng_NumOpe + 1
         End If
         
         g_rst_Genera.Close
         Set g_rst_Genera = Nothing
         
         g_rst_Princi.MoveNext
         DoEvents
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Cuotas Pendientes de Pago
   g_str_Parame = "SELECT * FROM CRE_HIPMAE WHERE "
   g_str_Parame = g_str_Parame & "HIPMAE_SITUAC = 2 "
   g_str_Parame = g_str_Parame & "ORDER BY HIPMAE_NUMOPE ASC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      r_int_FlgDat = 1
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         r_str_NumOpe = g_rst_Princi!HIPMAE_NUMOPE
         
         g_str_Parame = "SELECT * FROM CRE_HIPCUO WHERE "
         g_str_Parame = g_str_Parame & "HIPCUO_NUMOPE = '" & r_str_NumOpe & "' AND "
         g_str_Parame = g_str_Parame & "HIPCUO_TIPCRO = 1 AND "
         g_str_Parame = g_str_Parame & "HIPCUO_SITUAC = 2 "
         g_str_Parame = g_str_Parame & "ORDER BY HIPCUO_NUMCUO ASC "
      
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
            Exit Function
         End If
         
         If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
            g_rst_Genera.MoveFirst
            
            'r_str_NumRef = gf_Formato_NumOpe(g_rst_Genera!HIPCUO_NUMOPE)
            r_str_NomCli = moddat_gf_Buscar_NomCli(g_rst_Princi!HIPMAE_TDOCLI, g_rst_Princi!HIPMAE_NDOCLI)
            r_int_NumCuo = 0
            r_dbl_TotGas = 0
            l_lng_TemAux = 0
            
            Do While Not g_rst_Genera.EOF
               r_dbl_TotCuo = g_rst_Genera!HIPCUO_CAPITA + g_rst_Genera!HIPCUO_INTERE + g_rst_Genera!HIPCUO_DESORG + g_rst_Genera!HIPCUO_VIVORG + g_rst_Genera!HIPCUO_OTRORG
               r_dbl_TotCuo = r_dbl_TotCuo + g_rst_Genera!HIPCUO_INTCOM + g_rst_Genera!HIPCUO_INTMOR + g_rst_Genera!HIPCUO_GASCOB + g_rst_Genera!HIPCUO_OTRGAS
               r_dbl_TotCuo = r_dbl_TotCuo + g_rst_Genera!HIPCUO_CAPBBP + g_rst_Genera!HIPCUO_INTBBP
               r_dbl_TotCuo = r_dbl_TotCuo - g_rst_Genera!HIPCUO_IMPPAG
            
               If g_rst_Princi!HIPMAE_INDITF = 1 Then
                  r_dbl_ImpITF = gf_Truncar_Numero(r_dbl_TotCuo * (l_dbl_PorITF / 100), 2)
               Else
                  r_dbl_ImpITF = 0
               End If
               
               r_dbl_TotCuo = CDbl(Format(r_dbl_TotCuo + r_dbl_ImpITF, "###,###,##0.00"))
               r_dbl_TotGas = r_dbl_TotGas + r_dbl_TotCuo
                              
               If l_lng_Contad = -1 Then
                  ReDim Preserve l_str_CodCuo(0) As String
                  l_str_CodCuo(0) = Left(Trim(g_rst_Genera!HIPCUO_FECVCT), 6) & Trim(g_rst_Princi!HIPMAE_MONEDA) & "1"
               Else
                  For l_lng_ConAux = 0 To UBound(l_str_CodCuo) Step 1
                     If l_str_CodCuo(l_lng_ConAux) = Left(Trim(g_rst_Genera!HIPCUO_FECVCT), 6) & Trim(g_rst_Princi!HIPMAE_MONEDA) & "1" Then
                        l_lng_TemAux = 1
                     End If
                  Next
                  
                  If l_lng_TemAux = 0 Then
                     ReDim Preserve l_str_CodCuo(UBound(l_str_CodCuo) + 1) As String
                     l_str_CodCuo(UBound(l_str_CodCuo)) = Left(Trim(g_rst_Genera!HIPCUO_FECVCT), 6) & Trim(g_rst_Princi!HIPMAE_MONEDA) & "1"
                  End If
               End If

               ReDim Preserve l_str_DebAut(l_lng_ConTem + l_lng_NumOpe) As String
               l_str_DebAut(l_lng_Contad + 1) = "13"
               l_str_DebAut(l_lng_Contad + 2) = "" & gs_modsec_Genera(Format(Trim(g_rst_Princi!HIPMAE_TDOCLI) & Trim(g_rst_Princi!HIPMAE_NDOCLI), "########0"), 2, " ", 20)
               l_str_DebAut(l_lng_Contad + 3) = "" & gs_modsec_Genera(Left(Trim(g_rst_Genera!HIPCUO_FECVCT), 4) & "-" & Mid(Trim(g_rst_Genera!HIPCUO_FECVCT), 5, 2) & "1", 2, " ", 8)
               l_str_DebAut(l_lng_Contad + 4) = "" & gs_modsec_Genera(Left(r_str_NomCli, 30), 2, " ", 30)
               l_str_DebAut(l_lng_Contad + 5) = "" & String(10, "0")
               l_str_DebAut(l_lng_Contad + 6) = "" & String(10, "0")
               l_str_DebAut(l_lng_Contad + 7) = "" & Format(date, "yyyymmdd")
               l_str_DebAut(l_lng_Contad + 8) = "" & Trim(g_rst_Genera!HIPCUO_FECVCT)
               l_str_DebAut(l_lng_Contad + 9) = "" & gs_modsec_Genera(Format(Trim(g_rst_Princi!HIPMAE_NUMOPE), "###########0"), 1, "0", 15)
               l_str_DebAut(l_lng_Contad + 10) = "" & IIf(CInt(g_rst_Princi!HIPMAE_MONEDA) = 1, "01", "10")
               l_str_DebAut(l_lng_Contad + 11) = "" & gs_modsec_Genera(Format(r_dbl_TotGas, "########0.00"), 1, "0", 9)
               l_str_DebAut(l_lng_Contad + 12) = "" & String(9, "0")
               l_str_DebAut(l_lng_Contad + 13) = "" & String(9, "0")
               l_str_DebAut(l_lng_Contad + 14) = "" & String(9, "0")
               l_str_DebAut(l_lng_Contad + 15) = "" & String(9, "0")
               l_str_DebAut(l_lng_Contad + 16) = "" & String(9, "0")
               l_str_DebAut(l_lng_Contad + 17) = "" & String(9, "0")
               l_str_DebAut(l_lng_Contad + 18) = "" & String(2, "  ")
               l_str_DebAut(l_lng_Contad + 19) = "" & "M" '& "A" & "E" & "D"
               l_str_DebAut(l_lng_Contad + 20) = "" & String(13, "  ")
               l_str_DebAut(l_lng_Contad + 21) = "" & String(8, "0")
               l_str_DebAut(l_lng_Contad + 22) = "" & CInt(g_rst_Princi!HIPMAE_MONEDA)
               l_lng_Contad = l_lng_Contad + 22
               l_lng_ConTem = l_lng_ConTem + 21
               
               If g_rst_Princi!HIPMAE_MONEDA = 1 Then
                  l_int_CanSol = l_int_CanSol + 1
                  r_dbl_MtoSol = r_dbl_MtoSol + r_dbl_TotGas
               ElseIf g_rst_Princi!HIPMAE_MONEDA = 2 Then
                  l_int_CanDol = l_int_CanDol + 1
                  r_dbl_MtoDol = r_dbl_MtoDol + r_dbl_TotGas
               End If
               
               l_lng_NumOpe = l_lng_NumOpe + 1
               g_rst_Genera.MoveNext
               r_int_NumCuo = r_int_NumCuo + 1
               
               If r_int_NumCuo = 12 Then
                  Exit Do
               End If
            Loop
         End If
      
         g_rst_Genera.Close
         Set g_rst_Genera = Nothing
                        
         g_rst_Princi.MoveNext
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call Ordenar_Matriz(l_str_CodCuo, LBound(l_str_CodCuo), UBound(l_str_CodCuo))
   
   For l_lng_ConAux = 0 To UBound(l_str_CodCuo) Step 1
      l_str_CodCuo(l_lng_ConAux) = Left(l_str_CodCuo(l_lng_ConAux), 4) & "-" & Mid(l_str_CodCuo(l_lng_ConAux), 5, 4) '& Right(l_str_CodCuo(l_lng_ConAux), 1)
   Next
   
   l_str_NomArc = "0727501"
   Open p_RutArc & "\C" & Format(l_str_NomArc, "0000000") & ".TXT" For Output As r_int_NumFil
      
   Print #r_int_NumFil, "11" & "21" & Format(l_str_CodUni, "0000000000") & Format(l_str_CodRub, "00") & Format(l_str_CodEmp, "000") & Format(l_str_CodSer, "00") & "002" & Format(l_str_CodSol, "00") & gs_modsec_Genera("PAGO DE CUOTAS ", 2, " ", 30) & "1" & "M" & Format(l_int_CanSol + l_int_CanDol, "00000000") & _
   gs_modsec_Genera(Format(r_dbl_MtoSol, "########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(r_dbl_MtoDol, "########0.00"), 1, "0", 15) & Format(Now, "YYYYMMDD") & Space(88) & "00000000"
   
   l_lng_Contad = 0
   
   For l_lng_ConAux = 0 To UBound(l_str_CodCuo) Step 1
         
      If Left(Right(l_str_CodCuo(l_lng_ConAux), 2), 1) = 1 Then
         If Right(Right(l_str_CodCuo(l_lng_ConAux), 2), 1) = 1 Then
             r_str_Concep = gs_modsec_Genera("CUOTA " & UCase(Left(GenMes(Mid(l_str_CodCuo(l_lng_ConAux), 6, 2)), 4)), 2, " ", 10)
         ElseIf Right(Right(l_str_CodCuo(l_lng_ConAux), 2), 1) = 5 Then
             r_str_Concep = gs_modsec_Genera(Left("GASTOS CIERRE", 10), 2, " ", 10)
         Else
             r_str_Concep = gs_modsec_Genera(String(10, " "), 2, " ", 10)
         End If
         
         Print #r_int_NumFil, "12" & gs_modsec_Genera(Left(l_str_CodCuo(l_lng_ConAux), 7) & Right(l_str_CodCuo(l_lng_ConAux), 1), 2, " ", 8) & "1" & r_str_Concep & Space(60) & Space(111) & String(8, "0")
      End If
           
   Next
      
   l_lng_Contad = 0
      
   For l_lng_ConTem = 0 To UBound(l_str_DebAut) - 1 Step 22
      If Int(l_str_DebAut(l_lng_Contad + 21)) = 1 Then
         Print #r_int_NumFil, l_str_DebAut(l_lng_Contad + 0) & _
                              l_str_DebAut(l_lng_Contad + 1) & _
                              l_str_DebAut(l_lng_Contad + 2) & _
                              l_str_DebAut(l_lng_Contad + 3) & _
                              l_str_DebAut(l_lng_Contad + 4) & _
                              l_str_DebAut(l_lng_Contad + 5) & _
                              l_str_DebAut(l_lng_Contad + 6) & _
                              l_str_DebAut(l_lng_Contad + 7) & _
                              l_str_DebAut(l_lng_Contad + 8) & _
                              l_str_DebAut(l_lng_Contad + 9) & _
                              l_str_DebAut(l_lng_Contad + 10) & _
                              l_str_DebAut(l_lng_Contad + 11) & _
                              l_str_DebAut(l_lng_Contad + 12) & _
                              l_str_DebAut(l_lng_Contad + 13) & _
                              l_str_DebAut(l_lng_Contad + 14) & _
                              l_str_DebAut(l_lng_Contad + 15) & _
                              l_str_DebAut(l_lng_Contad + 16) & _
                              l_str_DebAut(l_lng_Contad + 17) & _
                              l_str_DebAut(l_lng_Contad + 18) & _
                              l_str_DebAut(l_lng_Contad + 19) & _
                              l_str_DebAut(l_lng_Contad + 20) '& _
                              'l_str_DebAut(l_lng_Contad + 21)
                        
      End If
      l_lng_Contad = l_lng_Contad + 22
   Next
      
   l_lng_Contad = 0
      
   For l_lng_ConTem = 0 To UBound(l_str_DebAut) - 1 Step 22
      If Int(l_str_DebAut(l_lng_Contad + 21)) = 2 Then
         Print #r_int_NumFil, l_str_DebAut(l_lng_Contad + 0) & _
                              l_str_DebAut(l_lng_Contad + 1) & _
                              l_str_DebAut(l_lng_Contad + 2) & _
                              l_str_DebAut(l_lng_Contad + 3) & _
                              l_str_DebAut(l_lng_Contad + 4) & _
                              l_str_DebAut(l_lng_Contad + 5) & _
                              l_str_DebAut(l_lng_Contad + 6) & _
                              l_str_DebAut(l_lng_Contad + 7) & _
                              l_str_DebAut(l_lng_Contad + 8) & _
                              l_str_DebAut(l_lng_Contad + 9) & _
                              l_str_DebAut(l_lng_Contad + 10) & _
                              l_str_DebAut(l_lng_Contad + 11) & _
                              l_str_DebAut(l_lng_Contad + 12) & _
                              l_str_DebAut(l_lng_Contad + 13) & _
                              l_str_DebAut(l_lng_Contad + 14) & _
                              l_str_DebAut(l_lng_Contad + 15) & _
                              l_str_DebAut(l_lng_Contad + 16) & _
                              l_str_DebAut(l_lng_Contad + 17) & _
                              l_str_DebAut(l_lng_Contad + 18) & _
                              l_str_DebAut(l_lng_Contad + 19) & _
                              l_str_DebAut(l_lng_Contad + 20) '& _
                              'l_str_DebAut(l_lng_Contad + 21)
                        
      End If
      l_lng_Contad = l_lng_Contad + 22
   Next
      
   'Cerrando Archivo de Pagos
   Close #r_int_NumFil
      
   If r_int_FlgDat = 0 Then
      Kill p_RutArc & "\0XXX" & Format(date, "MMDD") & ".TXT"
      ff_GenArc_Interbank = False
   Else
      ff_GenArc_Interbank = True
   End If
End Function

Private Function ff_GenArc_Interbank_old(ByVal p_RutArc As String) As Integer
   Dim r_int_NumFil     As Integer
   Dim r_str_NumSol     As String
   Dim r_dbl_ImpITF     As Double
   Dim r_dbl_TotGas     As Double
   Dim r_str_NomCli     As String
   Dim r_str_NumRef     As String
   Dim r_str_NumOpe     As String
   Dim r_int_NumCuo     As Integer
   Dim r_dbl_TotCuo     As Double
   Dim r_int_ConLin     As Double
   Dim r_dbl_ImpMax     As Double
   Dim r_dbl_ImpMin     As Double
   Dim r_str_CodCla     As String
   Dim r_int_FlgDat     As Integer
   Dim r_dbl_MtoSol     As Double
   Dim r_dbl_MtoDol     As Double
   Dim r_int_TipMon     As Integer
   
   r_int_ConLin = 0
   r_dbl_ImpMax = 0
   r_dbl_ImpMin = 0
   r_int_FlgDat = 0
   l_lng_Contad = -1
   l_lng_ConTem = 21
   l_lng_ConAux = 0
   l_int_CanSol = 0
   l_int_CanDol = 0
   r_int_FlgDat = 0
   r_int_NumFil = FreeFile
         
   'Obteniendo Gastos de Cierre Pendientes de Pago
   g_str_Parame = "SELECT GASADM_NUMSOL, SUM(GASADM_IMPORT) AS TOTGAS FROM TRA_GASADM WHERE "
   'g_str_Parame = g_str_Parame & "GASADM_TIPMON = " & CStr(p_TipMon) & " AND "
   g_str_Parame = g_str_Parame & "GASADM_SITUAC = 2 "
   g_str_Parame = g_str_Parame & "GROUP BY GASADM_NUMSOL "
   g_str_Parame = g_str_Parame & "ORDER BY GASADM_NUMSOL ASC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      r_int_FlgDat = 1
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         r_str_NumSol = g_rst_Princi!GASADM_NUMSOL
         
         'Verificando que Solicitud este en Trámite
         g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
         g_str_Parame = g_str_Parame & "SOLMAE_NUMERO = '" & r_str_NumSol & "' AND "
         g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1"
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
            Exit Function
         End If
         
         If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
            g_rst_Genera.MoveFirst
            
            r_dbl_TotGas = 0
            l_lng_TemAux = 0
            r_dbl_ImpITF = gf_NueImp_Numero(gf_Truncar_Numero(g_rst_Princi!TOTGAS * (l_dbl_PorITF / 100), 2))
            r_dbl_TotGas = CDbl(Format(g_rst_Princi!TOTGAS + r_dbl_ImpITF, "###,###,##0.00"))
            r_str_NomCli = moddat_gf_Buscar_NomCli(g_rst_Genera!SOLMAE_TITTDO, g_rst_Genera!SOLMAE_TITNDO)
            'r_str_NumRef = gf_Formato_NumSol(g_rst_Genera!SOLMAE_NUMERO)
                                                         
            If l_lng_Contad = -1 Then
               ReDim Preserve l_str_CodCuo(0) As String
               l_str_CodCuo(0) = Left(Format(date + CDate(4), "yyyymmdd"), 6) & Trim(g_rst_Genera!SOLMAE_TIPMON) & "5"
               
            Else
               For l_lng_ConAux = 0 To UBound(l_str_CodCuo) Step 1
                  If l_str_CodCuo(l_lng_ConAux) = Left(Format(date + CDate(4), "yyyymmdd"), 6) & Trim(g_rst_Genera!SOLMAE_TIPMON) & "5" Then
                     l_lng_TemAux = 1
                  End If
               Next
               
               If l_lng_TemAux = 0 Then
                  ReDim Preserve l_str_CodCuo(UBound(l_str_CodCuo) + 1) As String
                  l_str_CodCuo(UBound(l_str_CodCuo)) = Left(Format(date + CDate(4), "yyyymmdd"), 6) & Trim(g_rst_Genera!SOLMAE_TIPMON) & "5"
               End If
            
            End If
         
            ReDim Preserve l_str_DebAut(l_lng_ConTem + l_lng_NumOpe) As String
                        
            l_str_DebAut(l_lng_Contad + 1) = "13"
            l_str_DebAut(l_lng_Contad + 2) = "" & gs_modsec_Genera(Format(Trim(g_rst_Genera!SOLMAE_TITTDO) & Trim(g_rst_Genera!SOLMAE_TITNDO), "########0"), 2, " ", 20)
            l_str_DebAut(l_lng_Contad + 3) = "" & gs_modsec_Genera(Left(Format(date + CDate(4), "yyyymmdd"), 4) & "-" & Mid(Format(date + CDate(4), "yyyymmdd"), 5, 2) & "5", 2, " ", 8)
            l_str_DebAut(l_lng_Contad + 4) = "" & gs_modsec_Genera(Left(r_str_NomCli, 30), 2, " ", 30)
            l_str_DebAut(l_lng_Contad + 5) = "" & String(10, "0")
            l_str_DebAut(l_lng_Contad + 6) = "" & String(10, "0")
            l_str_DebAut(l_lng_Contad + 7) = "" & Format(date, "yyyymmdd")
            l_str_DebAut(l_lng_Contad + 8) = "" & Format(date + CDate(4), "yyyymmdd")
            l_str_DebAut(l_lng_Contad + 9) = "" & gs_modsec_Genera(Format(Trim(g_rst_Genera!SOLMAE_NUMERO), "###########0"), 1, "0", 15)
            l_str_DebAut(l_lng_Contad + 10) = "" & IIf(CInt(g_rst_Genera!SOLMAE_TIPMON) = 1, "01", "10")
            l_str_DebAut(l_lng_Contad + 11) = "" & String(9, "0")
            l_str_DebAut(l_lng_Contad + 12) = "" & String(9, "0")
            l_str_DebAut(l_lng_Contad + 13) = "" & String(9, "0")
            l_str_DebAut(l_lng_Contad + 14) = "" & String(9, "0")
            l_str_DebAut(l_lng_Contad + 15) = "" & gs_modsec_Genera(Format(r_dbl_TotGas, "########0.00"), 1, "0", 9)
            l_str_DebAut(l_lng_Contad + 16) = "" & String(9, "0")
            l_str_DebAut(l_lng_Contad + 17) = "" & String(9, "0")
            l_str_DebAut(l_lng_Contad + 18) = "" & String(2, "  ")
            l_str_DebAut(l_lng_Contad + 19) = "" & "A" '& "M" & "E" & "D"
            l_str_DebAut(l_lng_Contad + 20) = "" & String(13, "  ")
            l_str_DebAut(l_lng_Contad + 21) = "" & String(8, "0")
            l_str_DebAut(l_lng_Contad + 22) = "" & CInt(g_rst_Genera!SOLMAE_TIPMON)
                                 
            l_lng_Contad = l_lng_Contad + 22
            l_lng_ConTem = l_lng_ConTem + 21
            
            If g_rst_Genera!SOLMAE_TIPMON = 1 Then
               l_int_CanSol = l_int_CanSol + 1
               r_dbl_MtoSol = r_dbl_MtoSol + r_dbl_TotGas
            ElseIf g_rst_Genera!SOLMAE_TIPMON = 2 Then
               l_int_CanDol = l_int_CanDol + 1
               r_dbl_MtoDol = r_dbl_MtoDol + r_dbl_TotGas
            End If
            
            l_lng_NumOpe = l_lng_NumOpe + 1
            
         End If
         
         g_rst_Genera.Close
         Set g_rst_Genera = Nothing
         
         g_rst_Princi.MoveNext
         DoEvents
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Cuotas Pendientes de Pago
   g_str_Parame = "SELECT * FROM CRE_HIPMAE WHERE "
   g_str_Parame = g_str_Parame & "HIPMAE_SITUAC = 2 "
   'g_str_Parame = g_str_Parame & "AND HIPMAE_MONEDA = " & CStr(p_TipMon) & " "
   g_str_Parame = g_str_Parame & "ORDER BY HIPMAE_NUMOPE ASC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Function
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      r_int_FlgDat = 1
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         r_str_NumOpe = g_rst_Princi!HIPMAE_NUMOPE
         
         g_str_Parame = "SELECT * FROM CRE_HIPCUO WHERE "
         g_str_Parame = g_str_Parame & "HIPCUO_NUMOPE = '" & r_str_NumOpe & "' AND "
         g_str_Parame = g_str_Parame & "HIPCUO_TIPCRO = 1 AND "
         g_str_Parame = g_str_Parame & "HIPCUO_SITUAC = 2 "
         g_str_Parame = g_str_Parame & "ORDER BY HIPCUO_NUMCUO ASC "
      
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
            Exit Function
         End If
         
         If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
            g_rst_Genera.MoveFirst
            
            'r_str_NumRef = gf_Formato_NumOpe(g_rst_Genera!HIPCUO_NUMOPE)
            r_str_NomCli = moddat_gf_Buscar_NomCli(g_rst_Princi!HIPMAE_TDOCLI, g_rst_Princi!HIPMAE_NDOCLI)
            r_int_NumCuo = 0
            r_dbl_TotGas = 0
            l_lng_TemAux = 0
            
            Do While Not g_rst_Genera.EOF
               r_dbl_TotCuo = g_rst_Genera!HIPCUO_CAPITA + g_rst_Genera!HIPCUO_INTERE + g_rst_Genera!HIPCUO_DESORG + g_rst_Genera!HIPCUO_VIVORG + g_rst_Genera!HIPCUO_OTRORG
               r_dbl_TotCuo = r_dbl_TotCuo + g_rst_Genera!HIPCUO_INTCOM + g_rst_Genera!HIPCUO_INTMOR + g_rst_Genera!HIPCUO_GASCOB + g_rst_Genera!HIPCUO_OTRGAS
               r_dbl_TotCuo = r_dbl_TotCuo + g_rst_Genera!HIPCUO_CAPBBP + g_rst_Genera!HIPCUO_INTBBP
               r_dbl_TotCuo = r_dbl_TotCuo - g_rst_Genera!HIPCUO_IMPPAG
            
               If g_rst_Princi!HIPMAE_INDITF = 1 Then
                  r_dbl_ImpITF = gf_Truncar_Numero(r_dbl_TotCuo * (l_dbl_PorITF / 100), 2)
               Else
                  r_dbl_ImpITF = 0
               End If
               
               r_dbl_TotCuo = CDbl(Format(r_dbl_TotCuo + r_dbl_ImpITF, "###,###,##0.00"))
               r_dbl_TotGas = r_dbl_TotGas + r_dbl_TotCuo
                              
               If l_lng_Contad = -1 Then
                  ReDim Preserve l_str_CodCuo(0) As String
                  l_str_CodCuo(0) = Left(Trim(g_rst_Genera!HIPCUO_FECVCT), 6) & Trim(g_rst_Princi!HIPMAE_MONEDA) & "1"
                  
               Else
                  For l_lng_ConAux = 0 To UBound(l_str_CodCuo) Step 1
                     If l_str_CodCuo(l_lng_ConAux) = Left(Trim(g_rst_Genera!HIPCUO_FECVCT), 6) & Trim(g_rst_Princi!HIPMAE_MONEDA) & "1" Then
                        l_lng_TemAux = 1
                     End If
                  Next
                  
                  If l_lng_TemAux = 0 Then
                     ReDim Preserve l_str_CodCuo(UBound(l_str_CodCuo) + 1) As String
                     l_str_CodCuo(UBound(l_str_CodCuo)) = Left(Trim(g_rst_Genera!HIPCUO_FECVCT), 6) & Trim(g_rst_Princi!HIPMAE_MONEDA) & "1"
                  End If
               
               End If

               ReDim Preserve l_str_DebAut(l_lng_ConTem + l_lng_NumOpe) As String
                              
               l_str_DebAut(l_lng_Contad + 1) = "13"
               l_str_DebAut(l_lng_Contad + 2) = "" & gs_modsec_Genera(Format(Trim(g_rst_Princi!HIPMAE_TDOCLI) & Trim(g_rst_Princi!HIPMAE_NDOCLI), "########0"), 2, " ", 20)
               l_str_DebAut(l_lng_Contad + 3) = "" & gs_modsec_Genera(Left(Trim(g_rst_Genera!HIPCUO_FECVCT), 4) & "-" & Mid(Trim(g_rst_Genera!HIPCUO_FECVCT), 5, 2) & "1", 2, " ", 8)
               l_str_DebAut(l_lng_Contad + 4) = "" & gs_modsec_Genera(Left(r_str_NomCli, 30), 2, " ", 30)
               l_str_DebAut(l_lng_Contad + 5) = "" & String(10, "0")
               l_str_DebAut(l_lng_Contad + 6) = "" & String(10, "0")
               l_str_DebAut(l_lng_Contad + 7) = "" & Format(date, "yyyymmdd")
               l_str_DebAut(l_lng_Contad + 8) = "" & Trim(g_rst_Genera!HIPCUO_FECVCT)
               l_str_DebAut(l_lng_Contad + 9) = "" & gs_modsec_Genera(Format(Trim(g_rst_Princi!HIPMAE_NUMOPE), "###########0"), 1, "0", 15)
               l_str_DebAut(l_lng_Contad + 10) = "" & IIf(CInt(g_rst_Princi!HIPMAE_MONEDA) = 1, "01", "10")
               l_str_DebAut(l_lng_Contad + 11) = "" & gs_modsec_Genera(Format(r_dbl_TotGas, "########0.00"), 1, "0", 9)
               l_str_DebAut(l_lng_Contad + 12) = "" & String(9, "0")
               l_str_DebAut(l_lng_Contad + 13) = "" & String(9, "0")
               l_str_DebAut(l_lng_Contad + 14) = "" & String(9, "0")
               l_str_DebAut(l_lng_Contad + 15) = "" & String(9, "0")
               l_str_DebAut(l_lng_Contad + 16) = "" & String(9, "0")
               l_str_DebAut(l_lng_Contad + 17) = "" & String(9, "0")
               l_str_DebAut(l_lng_Contad + 18) = "" & String(2, "  ")
               l_str_DebAut(l_lng_Contad + 19) = "" & "M" '& "A" & "E" & "D"
               l_str_DebAut(l_lng_Contad + 20) = "" & String(13, "  ")
               l_str_DebAut(l_lng_Contad + 21) = "" & String(8, "0")
               l_str_DebAut(l_lng_Contad + 22) = "" & CInt(g_rst_Princi!HIPMAE_MONEDA)
                                    
               l_lng_Contad = l_lng_Contad + 22
               l_lng_ConTem = l_lng_ConTem + 21
               
               If g_rst_Princi!HIPMAE_MONEDA = 1 Then
                  l_int_CanSol = l_int_CanSol + 1
                  r_dbl_MtoSol = r_dbl_MtoSol + r_dbl_TotGas
               ElseIf g_rst_Princi!HIPMAE_MONEDA = 2 Then
                  l_int_CanDol = l_int_CanDol + 1
                  r_dbl_MtoDol = r_dbl_MtoDol + r_dbl_TotGas
               End If
               
               l_lng_NumOpe = l_lng_NumOpe + 1
               g_rst_Genera.MoveNext
               r_int_NumCuo = r_int_NumCuo + 1
               
               If r_int_NumCuo = 12 Then
                  Exit Do
               End If
            Loop
         End If
      
         g_rst_Genera.Close
         Set g_rst_Genera = Nothing
                        
         g_rst_Princi.MoveNext
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Call Ordenar_Matriz(l_str_CodCuo, LBound(l_str_CodCuo), UBound(l_str_CodCuo))
   
   For l_lng_ConAux = 0 To UBound(l_str_CodCuo) Step 1
      l_str_CodCuo(l_lng_ConAux) = Left(l_str_CodCuo(l_lng_ConAux), 4) & "-" & Mid(l_str_CodCuo(l_lng_ConAux), 5, 4) '& Right(l_str_CodCuo(l_lng_ConAux), 1)
   Next
         
   For r_int_TipMon = 1 To 2 Step 1
      'Archivo de Recaudo (Cabecera)
      If l_int_CanSol > 0 And r_int_TipMon = 1 Then
         'Nombre Archivo Plano en Soles
         l_str_NomArc = "0727501"
         'l_str_CodSol = 1 'SOLES
      ElseIf l_int_CanDol > 0 And r_int_TipMon = 2 Then
         'Nombre Archivo Plano en Dolares
         l_str_NomArc = "0727502"
         'l_str_CodSol = 2 'DOLARES
      End If
      
      Open p_RutArc & "\C" & Format(l_str_NomArc, "0000000") & ".TXT" For Output As r_int_NumFil
      Print #r_int_NumFil, "11" & "21" & Format(l_str_CodUni, "0000000000") & Format(l_str_CodRub, "00") & Format(l_str_CodEmp, "000") & Format(l_str_CodSer, "00") & "002" & Format(l_str_CodSol, "00") & gs_modsec_Genera("PAGO DE CUOTAS EN " & IIf(r_int_TipMon = 1, "SOLES", "DOLARES "), 2, " ", 30) & "1" & "M" & IIf(r_int_TipMon = 1, Format(l_int_CanSol, "00000000"), Format(l_int_CanDol, "00000000")) & _
      IIf(r_int_TipMon = 1, gs_modsec_Genera(Format(r_dbl_MtoSol, "########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(0, "########0.00"), 1, "0", 15), gs_modsec_Genera(Format(0, "########0.00"), 1, "0", 15) & gs_modsec_Genera(Format(r_dbl_MtoDol, "########0.00"), 1, "0", 15)) & Format(Now, "YYYYMMDD") & Space(88) & "00000000"
      
      
      l_lng_Contad = 0
      
      For l_lng_ConAux = 0 To UBound(l_str_CodCuo) Step 1
         
         If Left(Right(l_str_CodCuo(l_lng_ConAux), 2), 1) = r_int_TipMon Then
            'Print #r_int_NumFil, "12" & gs_modsec_Genera(Left(l_str_CodCuo(l_lng_ConAux), 7), 2, " ", 8) & "1" & UCase(gs_modsec_Genera(Left(GenMes(Mid(l_str_CodCuo(l_lng_ConAux), 6, 2)), 10), 2, " ", 10)) & Space(171) & "00000000"
            Print #r_int_NumFil, "12" & gs_modsec_Genera(Left(l_str_CodCuo(l_lng_ConAux), 7) & Right(l_str_CodCuo(l_lng_ConAux), 1), 2, " ", 8) & "1" & gs_modsec_Genera(IIf(Right(Right(l_str_CodCuo(l_lng_ConAux), 2), 1) = 1, "CUOTA " & UCase(Left(GenMes(Mid(l_str_CodCuo(l_lng_ConAux), 6, 2)), 4)), String(10, " ")), 2, " ", 10) & Space(30) & gs_modsec_Genera(IIf(Right(Right(l_str_CodCuo(l_lng_ConAux), 2), 1) = 5, Left("GASTOS CIERRE", 10), String(10, " ")), 2, " ", 10) & Space(131) & String(8, "0")
         End If
               
      Next
      
      l_lng_Contad = 0
      
      'For l_lng_ConTem = 0 To UBound(l_str_DebAut) - 1 Step 23
      For l_lng_ConTem = 0 To UBound(l_str_DebAut) - 1 Step 22
         If r_int_TipMon = Int(l_str_DebAut(l_lng_Contad + 21)) Then
            Print #r_int_NumFil, l_str_DebAut(l_lng_Contad + 0) & _
                                 l_str_DebAut(l_lng_Contad + 1) & _
                                 l_str_DebAut(l_lng_Contad + 2) & _
                                 l_str_DebAut(l_lng_Contad + 3) & _
                                 l_str_DebAut(l_lng_Contad + 4) & _
                                 l_str_DebAut(l_lng_Contad + 5) & _
                                 l_str_DebAut(l_lng_Contad + 6) & _
                                 l_str_DebAut(l_lng_Contad + 7) & _
                                 l_str_DebAut(l_lng_Contad + 8) & _
                                 l_str_DebAut(l_lng_Contad + 9) & _
                                 l_str_DebAut(l_lng_Contad + 10) & _
                                 l_str_DebAut(l_lng_Contad + 11) & _
                                 l_str_DebAut(l_lng_Contad + 12) & _
                                 l_str_DebAut(l_lng_Contad + 13) & _
                                 l_str_DebAut(l_lng_Contad + 14) & _
                                 l_str_DebAut(l_lng_Contad + 15) & _
                                 l_str_DebAut(l_lng_Contad + 16) & _
                                 l_str_DebAut(l_lng_Contad + 17) & _
                                 l_str_DebAut(l_lng_Contad + 18) & _
                                 l_str_DebAut(l_lng_Contad + 19) & _
                                 l_str_DebAut(l_lng_Contad + 20) '& _
                                 'l_str_DebAut(l_lng_Contad + 21)
                           
         End If
         l_lng_Contad = l_lng_Contad + 22
      
      Next
      
      'Cerrando Archivo de Pagos
      Close #r_int_NumFil
      
      If r_int_FlgDat = 0 Then
         Kill p_RutArc & "\0XXX" & Format(date, "MMDD") & ".TXT"
         ff_GenArc_Interbank_old = False
      Else
'         modgen_g_str_Mail_Asunto = "miCasita hipotecaria - Archivo Recaudación Interbank - " & IIf(r_int_TipMon = 1, "Soles ", "Dólares ") & Format(Date, "dd/mm/yyyy")
'         modgen_g_str_Mail_Mensaj = "Señores Interbank: <br><br>"
'         modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "Adjunto archivo de recaudo en " & IIf(r_int_TipMon = 1, "Soles ", "Dólares ") & "para que nuestros clientes puedan realizar sus pagos. <br><br>"
'         modgen_g_str_Mail_Mensaj = modgen_g_str_Mail_Mensaj & "Saludos Cordiales <br>"
'
'         Call fs_Envia_CorEle(modgen_g_str_Mail_Asunto, modgen_g_str_Mail_Mensaj, p_RutArc & "\C" & Format(l_str_CodEmp, "000000") & ".TXT", l_str_Destin)
'         Kill l_str_RutZip
'         r_int_FlgZip = r_int_FlgZip + 1
         
         ff_GenArc_Interbank_old = True
      End If
   
   Next
   
End Function

Sub Ordenar_Matriz(El_Vector() As String, Limite_Inferior As Long, Limite_Superior As Long)
  
Dim i As Long, j As Long, x As Variant, y As Variant
  
   i = Limite_Inferior
   j = Limite_Superior
   x = El_Vector((Limite_Inferior + Limite_Superior) / 2)
  
   While i <= j
      While (El_Vector(i) < x) And (i < Limite_Superior)
          i = i + 1
      Wend
       
      While (x < El_Vector(j)) And (j > Limite_Inferior)
          j = j - 1
      Wend
       
      If i <= j Then
          y = El_Vector(i)
          El_Vector(i) = El_Vector(j)
          El_Vector(j) = y
          i = i + 1
          j = j - 1
      End If
   Wend
  
   If Limite_Inferior < j Then Ordenar_Matriz El_Vector(), Limite_Inferior, j
   If i < Limite_Superior Then Ordenar_Matriz El_Vector(), i, Limite_Superior
End Sub

Private Function GenMes(ByVal p_Mes As Integer) As String
   Select Case (p_Mes)
      Case 1:         GenMes = "Enero"
      Case 2:     GenMes = "Febrero"
      Case 3:     GenMes = "Marzo"
      Case 4:     GenMes = "Abril"
      Case 5:     GenMes = "Mayo"
      Case 6:     GenMes = "Junio"
      Case 7:     GenMes = "Julio"
      Case 8:     GenMes = "Agosto"
      Case 9:     GenMes = "Septiembre"
      Case 10:    GenMes = "Octubre"
      Case 11:    GenMes = "Noviembre"
      Case 12:    GenMes = "Diciembre"
   End Select
End Function







