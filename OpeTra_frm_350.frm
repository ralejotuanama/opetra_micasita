VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_Rpt_Cofide_04 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4965
   Icon            =   "OpeTra_frm_350.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   4965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   2415
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   4965
      _Version        =   65536
      _ExtentX        =   8758
      _ExtentY        =   4260
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
         TabIndex        =   5
         Top             =   60
         Width           =   4875
         _Version        =   65536
         _ExtentX        =   8599
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
            Height          =   300
            Left            =   630
            TabIndex        =   6
            Top             =   30
            Width           =   4065
            _Version        =   65536
            _ExtentX        =   7170
            _ExtentY        =   529
            _StockProps     =   15
            Caption         =   "Reporte de Comparativo de Cobranza Mensual"
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
         Begin Threed.SSPanel SSPanel2 
            Height          =   270
            Left            =   630
            TabIndex        =   7
            Top             =   315
            Width           =   2295
            _Version        =   65536
            _ExtentX        =   4048
            _ExtentY        =   476
            _StockProps     =   15
            Caption         =   "miCasita - Cofide"
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
            Picture         =   "OpeTra_frm_350.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   8
         Top             =   780
         Width           =   4875
         _Version        =   65536
         _ExtentX        =   8599
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
            Left            =   4260
            Picture         =   "OpeTra_frm_350.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_350.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   885
         Left            =   30
         TabIndex        =   9
         Top             =   1470
         Width           =   4875
         _Version        =   65536
         _ExtentX        =   8599
         _ExtentY        =   1561
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
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   120
            Width           =   2775
         End
         Begin EditLib.fpLongInteger ipp_PerAno 
            Height          =   315
            Left            =   1530
            TabIndex        =   1
            Top             =   480
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
            MaxValue        =   "9999"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
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
         Begin VB.Label Label1 
            Caption         =   "Periodo:"
            Height          =   315
            Left            =   180
            TabIndex        =   11
            Top             =   150
            Width           =   1245
         End
         Begin VB.Label Label3 
            Caption         =   "Año:"
            Height          =   285
            Left            =   180
            TabIndex        =   10
            Top             =   510
            Width           =   1065
         End
      End
   End
End
Attribute VB_Name = "frm_Rpt_Cofide_04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim l_obj_Excel            As Excel.Application
Dim l_str_PerMes           As String
Dim l_str_PerAno           As String
Dim l_str_TipArc           As String
Dim l_int_numhoja          As Integer

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Limpia
   Call gs_CentraForm(Me)
   
   Call gs_SetFocus(cmb_PerMes)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   cmb_PerMes.Clear
   Call moddat_gs_Carga_LisIte_Combo(cmb_PerMes, 1, "033")
   ipp_PerAno = Mid(date, 7, 4)
End Sub

Private Sub fs_Limpia()
   cmb_PerMes.ListIndex = -1
End Sub

Private Sub cmd_ExpExc_Click()
Dim r_int_AntMes      As Integer
Dim r_int_AntAno      As Integer
Dim r_str_AntIni      As String
Dim r_str_AntFin      As String
Dim r_str_ActIni      As String
Dim r_str_ActFin      As String
Dim r_str_fchaux      As String
Dim r_int_Filaux      As Integer
Dim r_int_numhoja     As Integer
Dim r_obj_Excel       As Excel.Application
Dim r_int_FilExl      As Integer
Dim r_int_FilItm      As Integer
Dim r_str_Cadena      As String
Dim r_dbl_ImpAux      As Double

'Variables totales por grupo
Dim r_dbl_SalCof      As Double
Dim r_dbl_SalMic      As Double
Dim r_dbl_SalDif      As Double
Dim r_dbl_CapCof      As Double
Dim r_dbl_CapMic      As Double
Dim r_dbl_CapDif      As Double
Dim r_dbl_IntCof      As Double
Dim r_dbl_IntMic      As Double
Dim r_dbl_IntDif      As Double
Dim r_dbl_TasDif      As Double
Dim r_dbl_ComCof      As Double
Dim r_dbl_ComMic      As Double
Dim r_dbl_ComDif      As Double
Dim r_dbl_TotCof      As Double
Dim r_dbl_TotMic      As Double
Dim r_dbl_TotDif      As Double
Dim r_dbl_FinCof      As Double
Dim r_dbl_FinMic      As Double
Dim r_dbl_FinDif      As Double

'Variables totales generales
Dim r_dbl_Tot_SalCof  As Double
Dim r_dbl_Tot_SalMic  As Double
Dim r_dbl_Tot_SalDif  As Double
Dim r_dbl_Tot_CapCof  As Double
Dim r_dbl_Tot_CapMic  As Double
Dim r_dbl_Tot_CapDif  As Double
Dim r_dbl_Tot_IntCof  As Double
Dim r_dbl_Tot_IntMic  As Double
Dim r_dbl_Tot_IntDif  As Double
Dim r_dbl_Tot_TasDif  As Double
Dim r_dbl_Tot_ComCof  As Double
Dim r_dbl_Tot_ComMic  As Double
Dim r_dbl_Tot_ComDif  As Double
Dim r_dbl_Tot_TotCof  As Double
Dim r_dbl_Tot_TotMic  As Double
Dim r_dbl_Tot_TotDif  As Double

Dim r_dbl_Tot_FinCof  As Double
Dim r_dbl_Tot_FinMic  As Double
Dim r_dbl_Tot_FinDif  As Double

   If cmb_PerMes.ListIndex = -1 Then
      MsgBox "Debe seleccionar un Periodo.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PerMes)
      Exit Sub
   End If
   If ipp_PerAno.Text = 0 Then
      MsgBox "Debe seleccionar un Año.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_PerAno)
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   l_str_PerMes = cmb_PerMes.ItemData(cmb_PerMes.ListIndex)
   l_str_PerAno = ipp_PerAno.Text
   cmd_ExpExc.Enabled = False
   
   r_str_fchaux = Format(DateAdd("m", -1, CDate("01/" & l_str_PerMes & "/" & l_str_PerAno)), "yyyymmdd")
   r_int_AntMes = Mid(r_str_fchaux, 5, 2)
   r_int_AntAno = Mid(r_str_fchaux, 1, 4)
   'fecha anterior
   r_str_AntIni = r_int_AntAno & Format(r_int_AntMes, "00") & "01"
   r_str_AntFin = r_int_AntAno & Format(r_int_AntMes, "00") & "31"
   'fecha actual
   r_str_ActIni = l_str_PerAno & Format(l_str_PerMes, "00") & "01"
   r_str_ActFin = l_str_PerAno & Format(l_str_PerMes, "00") & "31"

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT SUBSTR(TB.NUM_OPERACION,1,3)||'-'|| SUBSTR(TB.NUM_OPERACION,4,2)||'-'|| SUBSTR(TB.NUM_OPERACION,6,5) AS NUM_OPERACION, "
   g_str_Parame = g_str_Parame & "       HM.HIPMAE_TDOCLI ||'-'|| HM.HIPMAE_NDOCLI AS NUM_DOCUMENTO, ARCMEN_NUMCTR, ARCMEN_IDCIPR,ARCMEN_NOMPRO, "
   g_str_Parame = g_str_Parame & "       ARCMEN_CODCOF, ARCMEN_NOMCLI   , TB.ARCMEN_NUMALT, ARCMEN_TIPMON, ARCMEN_EXPINI, ARCMEN_PRINCI, ARCMEN_IMPINT   , "
   g_str_Parame = g_str_Parame & "       ARCMEN_IMTASA, TB.ARCMEN_COMSIN, ARCMEN_IMPTOT   , ARCMEN_EXPFIN, ARCMEN_BUENPG, ARCMEN_MALPAG, HM.HIPMAE_NUMOPE, HM.HIPMAE_IMPNCO, "
   g_str_Parame = g_str_Parame & "       (SELECT HC.HIPCIE_TASCOF "
   g_str_Parame = g_str_Parame & "          FROM CRE_HIPCIE HC "
   g_str_Parame = g_str_Parame & "         WHERE HC.HIPCIE_NUMOPE = TB.NUM_OPERACION "
   g_str_Parame = g_str_Parame & "           AND HC.HIPCIE_PERMES = " & r_int_AntMes
   g_str_Parame = g_str_Parame & "           AND HC.HIPCIE_PERANO = " & r_int_AntAno & ") AS HIPCIE_TASCOF, "
   g_str_Parame = g_str_Parame & "       CASE WHEN HM.HIPMAE_CODPRD = '003' "
   g_str_Parame = g_str_Parame & "            THEN (SELECT C.HIPCUO_SALCAP FROM CRE_HIPCUO C "
   g_str_Parame = g_str_Parame & "                   WHERE C.HIPCUO_NUMOPE = HM.HIPMAE_NUMOPE And C.HIPCUO_TIPCRO = 5   "
   g_str_Parame = g_str_Parame & "                     AND C.HIPCUO_FECVCT >= " & r_str_AntIni & " AND C.HIPCUO_FECVCT <= " & r_str_AntFin & ") "
   g_str_Parame = g_str_Parame & "            ELSE (SELECT C.HIPCUO_SALCAP FROM CRE_HIPCUO C "
   g_str_Parame = g_str_Parame & "                   WHERE C.HIPCUO_NUMOPE = HM.HIPMAE_NUMOPE And C.HIPCUO_TIPCRO = 3   "
   g_str_Parame = g_str_Parame & "                     AND C.HIPCUO_FECVCT >= " & r_str_AntIni & " AND C.HIPCUO_FECVCT <= " & r_str_AntFin & ") "
   g_str_Parame = g_str_Parame & "       END AS SALDO_ANTEXPINI, "
   g_str_Parame = g_str_Parame & "       CASE WHEN HM.HIPMAE_CODPRD = '003' "
   g_str_Parame = g_str_Parame & "            THEN (SELECT D.HIPCUO_CAPITA FROM CRE_HIPCUO D "
   g_str_Parame = g_str_Parame & "                   WHERE D.HIPCUO_NUMOPE = HM.HIPMAE_NUMOPE And D.HIPCUO_TIPCRO = 5   "
   g_str_Parame = g_str_Parame & "                     AND D.HIPCUO_FECVCT >= " & r_str_ActIni & " AND D.HIPCUO_FECVCT <= " & r_str_ActFin & ") "
   g_str_Parame = g_str_Parame & "            ELSE (SELECT D.HIPCUO_CAPITA FROM CRE_HIPCUO D "
   g_str_Parame = g_str_Parame & "                   WHERE D.HIPCUO_NUMOPE = HM.HIPMAE_NUMOPE And D.HIPCUO_TIPCRO = 3   "
   g_str_Parame = g_str_Parame & "                     AND D.HIPCUO_FECVCT >= " & r_str_ActIni & " AND D.HIPCUO_FECVCT <= " & r_str_ActFin & ") "
   g_str_Parame = g_str_Parame & "       END  AS CAPITAL_PRINCI, "
   g_str_Parame = g_str_Parame & "       CASE WHEN HM.HIPMAE_CODPRD = '003' "
   g_str_Parame = g_str_Parame & "            THEN (SELECT E.HIPCUO_INTERE FROM CRE_HIPCUO E "
   g_str_Parame = g_str_Parame & "                   WHERE E.HIPCUO_NUMOPE = HM.HIPMAE_NUMOPE And E.HIPCUO_TIPCRO = 5   "
   g_str_Parame = g_str_Parame & "                     AND E.HIPCUO_FECVCT >= " & r_str_ActIni & " AND E.HIPCUO_FECVCT <= " & r_str_ActFin & ") "
   g_str_Parame = g_str_Parame & "            ELSE (SELECT E.HIPCUO_INTERE FROM CRE_HIPCUO E "
   g_str_Parame = g_str_Parame & "                   WHERE E.HIPCUO_NUMOPE = HM.HIPMAE_NUMOPE And E.HIPCUO_TIPCRO = 3   "
   g_str_Parame = g_str_Parame & "                     AND E.HIPCUO_FECVCT >= " & r_str_ActIni & " AND E.HIPCUO_FECVCT <= " & r_str_ActFin & ") "
   g_str_Parame = g_str_Parame & "       END AS INTERES_IMPINT, "
   g_str_Parame = g_str_Parame & "       CASE WHEN HM.HIPMAE_CODPRD = '003' "
   g_str_Parame = g_str_Parame & "            THEN (SELECT F.HIPCUO_COMCOF FROM CRE_HIPCUO F "
   g_str_Parame = g_str_Parame & "                   WHERE F.HIPCUO_NUMOPE = HM.HIPMAE_NUMOPE And F.HIPCUO_TIPCRO = 5   "
   g_str_Parame = g_str_Parame & "                     AND F.HIPCUO_FECVCT >= " & r_str_ActIni & " AND F.HIPCUO_FECVCT <= " & r_str_ActFin & ") "
   g_str_Parame = g_str_Parame & "            ELSE (SELECT F.HIPCUO_COMCOF FROM CRE_HIPCUO F "
   g_str_Parame = g_str_Parame & "                   Where F.HIPCUO_NUMOPE = HM.HIPMAE_NUMOPE And F.HIPCUO_TIPCRO = 3   "
   g_str_Parame = g_str_Parame & "                     AND F.HIPCUO_FECVCT >= " & r_str_ActIni & " AND F.HIPCUO_FECVCT <= " & r_str_ActFin & ") "
   g_str_Parame = g_str_Parame & "            END AS COMISION, "
   g_str_Parame = g_str_Parame & "       CASE WHEN HM.HIPMAE_CODPRD = '003' "
   g_str_Parame = g_str_Parame & "            THEN (SELECT G.HIPCUO_SALCAP FROM CRE_HIPCUO G "
   g_str_Parame = g_str_Parame & "                   WHERE G.HIPCUO_NUMOPE = HM.HIPMAE_NUMOPE And G.HIPCUO_TIPCRO = 5   "
   g_str_Parame = g_str_Parame & "                     AND G.HIPCUO_FECVCT >= " & r_str_ActIni & " AND G.HIPCUO_FECVCT <= " & r_str_ActFin & ") "
   g_str_Parame = g_str_Parame & "            ELSE (SELECT G.HIPCUO_SALCAP FROM CRE_HIPCUO G "
   g_str_Parame = g_str_Parame & "                   WHERE G.HIPCUO_NUMOPE = HM.HIPMAE_NUMOPE And G.HIPCUO_TIPCRO = 3   "
   g_str_Parame = g_str_Parame & "                     AND G.HIPCUO_FECVCT >= " & r_str_ActIni & " AND G.HIPCUO_FECVCT <= " & r_str_ActFin & ") "
   g_str_Parame = g_str_Parame & "       END AS SALDO_EXPFIN  "
   g_str_Parame = g_str_Parame & "  FROM  "
   g_str_Parame = g_str_Parame & "       (SELECT AM.*, (SELECT X.HIPMAE_NUMOPE "
   g_str_Parame = g_str_Parame & "                        FROM CRE_HIPMAE X "
   g_str_Parame = g_str_Parame & "                       WHERE AM.ARCMEN_CODCOF = X.HIPMAE_CODCOF AND X.HIPMAE_SITUAC IN (2,6,9) AND ROWNUM = 1) AS  NUM_OPERACION "
   g_str_Parame = g_str_Parame & "          FROM CRE_ARCMEN AM  "
   g_str_Parame = g_str_Parame & "         ORDER BY ARCMEN_IDCIPR ASC, ARCMEN_NUMCTR ASC) TB       "
   g_str_Parame = g_str_Parame & "  LEFT JOIN CRE_HIPMAE HM ON TB.NUM_OPERACION = HM.HIPMAE_NUMOPE "
   g_str_Parame = g_str_Parame & " WHERE TB.ARCMEN_PERANO = " & l_str_PerAno
   g_str_Parame = g_str_Parame & "   AND TB.ARCMEN_PERMES = " & l_str_PerMes
   g_str_Parame = g_str_Parame & " ORDER BY TB.ARCMEN_IDCIPR ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      MsgBox "No se encontró información para el período seleccionado.", vbExclamation, modgen_g_str_NomPlt
   Else
      Set r_obj_Excel = New Excel.Application
      r_obj_Excel.SheetsInNewWorkbook = 6
      r_obj_Excel.Workbooks.Add
      r_int_FilExl = 3
      r_int_FilItm = 1
      r_str_Cadena = ""
      r_int_numhoja = 1
      
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         'If (r_int_FilExl = 3) Then
             'r_obj_Excel.SheetsInNewWorkbook = r_int_numhoja
         'End If
         With r_obj_Excel.Sheets(r_int_numhoja)
              If (r_int_FilExl = 3) Then
                  'inicializando variables
                  r_dbl_SalCof = 0: r_dbl_SalMic = 0: r_dbl_SalDif = 0: r_dbl_CapCof = 0
                  r_dbl_CapMic = 0: r_dbl_CapDif = 0: r_dbl_IntCof = 0: r_dbl_IntMic = 0
                  r_dbl_IntDif = 0: r_dbl_TasDif = 0: r_dbl_ComCof = 0: r_dbl_ComMic = 0
                  r_dbl_ComDif = 0: r_dbl_TotCof = 0: r_dbl_TotMic = 0: r_dbl_TotDif = 0
                  r_dbl_FinCof = 0: r_dbl_FinMic = 0: r_dbl_FinDif = 0
                    
                  'ancho en las columnas
                  .Columns("A").ColumnWidth = 4.57
                  .Columns("B").ColumnWidth = 11.57
                  .Columns("C").ColumnWidth = 10.14
                  .Columns("D").ColumnWidth = 10.57
                  .Columns("E").ColumnWidth = 3.5
                  .Columns("F").ColumnWidth = 17.86
                  .Columns("G").ColumnWidth = 13.43
                  .Columns("H").ColumnWidth = 26
                  .Columns("I").ColumnWidth = 14
                  .Columns("J").ColumnWidth = 16
                  .Columns("K").ColumnWidth = 13
                  .Columns("L").ColumnWidth = 13
                  .Columns("M").ColumnWidth = 13
                  .Columns("N").ColumnWidth = 13
                  .Columns("O").ColumnWidth = 13
                  .Columns("P").ColumnWidth = 13
                  .Columns("Q").ColumnWidth = 13
                  .Columns("R").ColumnWidth = 0
                  .Columns("S").ColumnWidth = 0
                  .Columns("T").ColumnWidth = 0
                  .Columns("U").ColumnWidth = 13
                  .Columns("V").ColumnWidth = 15
                  .Columns("W").ColumnWidth = 13
                  .Columns("X").ColumnWidth = 13
                  .Columns("Y").ColumnWidth = 12
                  .Columns("Z").ColumnWidth = 13
                  .Columns("AA").ColumnWidth = 14
                  .Columns("AB").ColumnWidth = 16
                  .Columns("AC").ColumnWidth = 13
                  .Columns("AD").ColumnWidth = 10.86
                  .Columns("AE").ColumnWidth = 10.43
                  
                  'en negrita las diferencias
                  .Columns("K").Font.Bold = True
                  .Columns("N").Font.Bold = True
                  .Columns("Q").Font.Bold = True
                  .Columns("T").Font.Bold = True
                  .Columns("W").Font.Bold = True
                  .Columns("Z").Font.Bold = True
                  .Columns("AC").Font.Bold = True
                       
                  .Range("A3:AE3").Font.Bold = True
                  .Range("A3:AE3").Font.Size = 8
                  .Range("A3:AE3").Interior.Color = RGB(146, 208, 80)
                  .Range("A3:AE3").HorizontalAlignment = xlHAlignCenter
                  .Columns("AD").HorizontalAlignment = xlHAlignCenter
                  .Columns("AE").HorizontalAlignment = xlHAlignCenter
                  .Range("A3:AE3").Borders(xlEdgeLeft).LineStyle = xlContinuous
                  .Range("A3:AE3").Borders(xlEdgeTop).LineStyle = xlContinuous
                  .Range("A3:AE3").Borders(xlEdgeBottom).LineStyle = xlContinuous
                  .Range("A3:AE3").Borders(xlEdgeRight).LineStyle = xlContinuous
                  .Range("A3:AE3").Borders(xlInsideVertical).LineStyle = xlContinuous
                  'titulo en el excel
                  .Cells(r_int_FilExl, 1) = "ITEM"
                  .Cells(r_int_FilExl, 2) = "NRO. OPERACION"
                  .Cells(r_int_FilExl, 3) = "DOI. CLIENTE"
                  .Cells(r_int_FilExl, 4) = "COD. COFIDE"
                  .Cells(r_int_FilExl, 5) = "CIPR"
                  .Cells(r_int_FilExl, 6) = "PRODUCTO"
                  .Cells(r_int_FilExl, 7) = "NRO. CONTRATO"
                  .Cells(r_int_FilExl, 8) = "NOMBRES"
                  .Cells(r_int_FilExl, 9) = "SALDO ANT. COFIDE"      'ARCMEN_EXPINI
                  .Cells(r_int_FilExl, 10) = "SALDO ANT. MICASITA"   'SALDO_ANTEXPINI
                  .Cells(r_int_FilExl, 11) = "DIFERENCIA"
                  .Cells(r_int_FilExl, 12) = "CAPITAL COFIDE"        'ARCMEN_PRINCI
                  .Cells(r_int_FilExl, 13) = "CAPITAL MICASITA"      'CAPITAL_PRINCI
                  .Cells(r_int_FilExl, 14) = "DIFERENCIA"
                  .Cells(r_int_FilExl, 15) = "INTERES COFIDE"        'ARCMEN_IMPINT
                  .Cells(r_int_FilExl, 16) = "INTERES MICASITA"      'INTERES_IMPINT
                  .Cells(r_int_FilExl, 17) = "DIFERENCIA"
                  .Cells(r_int_FilExl, 18) = "TASA COFIDE"           'ARCMEN_IMTASA
                  .Cells(r_int_FilExl, 19) = "TASA MICASITA"         'HIPCIE_TASCOF
                  .Cells(r_int_FilExl, 20) = "DIFERENCIA"
                  .Cells(r_int_FilExl, 21) = "COMISION COFIDE"       'ARCMEN_COMSIN
                  .Cells(r_int_FilExl, 22) = "COMISION MICASITA"     'COMISION
                  .Cells(r_int_FilExl, 23) = "DIFERENCIA"
                  .Cells(r_int_FilExl, 24) = "TOTAL COFIDE"          'ARCMEN_IMPTOT
                  .Cells(r_int_FilExl, 25) = "TOTAL MICASITA"        'ARCMEN_IMPTOT
                  .Cells(r_int_FilExl, 26) = "DIFERENCIA"
                  .Cells(r_int_FilExl, 27) = "SALDO FINAL COFIDE"    'ARCMEN_EXPFIN
                  .Cells(r_int_FilExl, 28) = "SALDO FINAL MICASITA"  'SALDO_EXPFIN
                  .Cells(r_int_FilExl, 29) = "DIFERENCIA"
                  .Cells(r_int_FilExl, 30) = "BUEN PAGADOR"
                  .Cells(r_int_FilExl, 31) = "MAL PAGADOR"
                      
                  r_obj_Excel.Sheets(r_int_numhoja).Name = g_rst_Princi!ARCMEN_NOMPRO
                  r_int_FilExl = r_int_FilExl + 1
              End If
                                  
              .Cells(r_int_FilExl, 1) = r_int_FilItm
              .Cells(r_int_FilExl, 2) = g_rst_Princi!NUM_OPERACION
              .Cells(r_int_FilExl, 3) = g_rst_Princi!NUM_DOCUMENTO
              .Cells(r_int_FilExl, 4) = g_rst_Princi!ARCMEN_CODCOF
              .Cells(r_int_FilExl, 5) = g_rst_Princi!ARCMEN_IDCIPR
              r_str_Cadena = g_rst_Princi!ARCMEN_IDCIPR
              .Cells(r_int_FilExl, 6) = g_rst_Princi!ARCMEN_NOMPRO
              .Cells(r_int_FilExl, 7) = g_rst_Princi!ARCMEN_NUMCTR
              .Cells(r_int_FilExl, 8) = g_rst_Princi!ARCMEN_NOMCLI
               
              'evaluar si es buen pagador o mal pagador
              If (IsNull(g_rst_Princi!ARCMEN_BUENPG) = True) Then
                 .Cells(r_int_FilExl, 30) = "Si"
              ElseIf (UCase(Left(Trim(g_rst_Princi!ARCMEN_BUENPG), 1)) = UCase(Trim("N"))) Then
                 .Cells(r_int_FilExl, 30) = "No"
                 Else
                 .Cells(r_int_FilExl, 30) = "Si"
              End If
              If (IsNull(g_rst_Princi!ARCMEN_MALPAG) = True) Then
                 .Cells(r_int_FilExl, 31) = "No"
              ElseIf (UCase(Left(Trim(g_rst_Princi!ARCMEN_MALPAG), 1)) = UCase(Trim("N"))) Then
                 .Cells(r_int_FilExl, 31) = "No"
              Else
                 .Cells(r_int_FilExl, 31) = "Si"
              End If
               
              'validacion de clientes morosos(si es buen pagador o mal pagador)
              If (UCase(Left(Trim(.Cells(r_int_FilExl, 30)), 1)) = UCase(Trim("N"))) Then
                  r_str_fchaux = DateAdd("m", -1, CDate("01/" & l_str_PerMes & "/" & l_str_PerAno))
                  .Cells(r_int_FilExl, 10) = 0: .Cells(r_int_FilExl, 13) = 0: .Cells(r_int_FilExl, 16) = 0
                  .Cells(r_int_FilExl, 22) = 0: .Cells(r_int_FilExl, 28) = 0
                  g_str_Parame = ""
                  g_str_Parame = g_str_Parame & "SELECT *  "
                  g_str_Parame = g_str_Parame & "  FROM (SELECT HIPCUO_NUMCUO, HIPCUO_FECVCT, HIPCUO_CAPITA, "
                  g_str_Parame = g_str_Parame & "               HIPCUO_INTERE , HIPCUO_SALCAP, HIPCUO_COMCOF "
                  g_str_Parame = g_str_Parame & "          FROM CRE_HIPCUO C "
                  g_str_Parame = g_str_Parame & "         WHERE HIPCUO_NUMOPE = '" & g_rst_Princi!HIPMAE_NUMOPE & "' "
                  g_str_Parame = g_str_Parame & "           AND HIPCUO_TIPCRO = 4 "
                  g_str_Parame = g_str_Parame & "           AND HIPCUO_FECVCT < " & r_str_ActFin
                  g_str_Parame = g_str_Parame & "         ORDER BY HIPCUO_NUMCUO DESC) TAB "
                  g_str_Parame = g_str_Parame & " WHERE ROWNUM = 1 "
                      
                  r_dbl_ImpAux = 0
                  If gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
                     If Not g_rst_Genera.BOF And Not g_rst_Genera.EOF Then
                        .Cells(r_int_FilExl, 13) = IIf(IsNull(g_rst_Genera!HIPCUO_CAPITA) = True, 0, g_rst_Genera!HIPCUO_CAPITA)
                        .Cells(r_int_FilExl, 13) = .Cells(r_int_FilExl, 13) / 6
                        .Cells(r_int_FilExl, 16) = IIf(IsNull(g_rst_Genera!HIPCUO_INTERE) = True, 0, g_rst_Genera!HIPCUO_INTERE)
                        .Cells(r_int_FilExl, 16) = .Cells(r_int_FilExl, 16) / 6
                        .Cells(r_int_FilExl, 22) = IIf(IsNull(g_rst_Genera!HIPCUO_COMCOF) = True, 0, g_rst_Genera!HIPCUO_COMCOF)
                        .Cells(r_int_FilExl, 22) = .Cells(r_int_FilExl, 22) / 6
                                                                                    
                        r_str_fchaux = "01/" & Mid(g_rst_Genera!HIPCUO_FECVCT, 5, 2) & "/" & Left(g_rst_Genera!HIPCUO_FECVCT, 4)
                        r_dbl_ImpAux = 6 - DateDiff("M", r_str_fchaux, "01/" & r_int_AntMes & "/" & r_int_AntAno)
                        r_dbl_ImpAux = .Cells(r_int_FilExl, 13) * r_dbl_ImpAux
                        .Cells(r_int_FilExl, 10) = Round(g_rst_Genera!HIPCUO_SALCAP + r_dbl_ImpAux, 2)
                             
                        r_dbl_ImpAux = 0
                        r_str_fchaux = "01/" & Mid(g_rst_Genera!HIPCUO_FECVCT, 5, 2) & "/" & Left(g_rst_Genera!HIPCUO_FECVCT, 4)
                        r_dbl_ImpAux = 6 - DateDiff("M", r_str_fchaux, "01/" & l_str_PerMes & "/" & l_str_PerAno)
                        r_dbl_ImpAux = .Cells(r_int_FilExl, 13) * r_dbl_ImpAux
                        .Cells(r_int_FilExl, 28) = Round(g_rst_Genera!HIPCUO_SALCAP + r_dbl_ImpAux, 2)
                             
                        .Cells(r_int_FilExl, 13) = Round(.Cells(r_int_FilExl, 13), 2)
                        .Cells(r_int_FilExl, 16) = Round(.Cells(r_int_FilExl, 16), 2)
                        .Cells(r_int_FilExl, 22) = Round(.Cells(r_int_FilExl, 22), 2)
                        .Cells(r_int_FilExl, 10) = Round(.Cells(r_int_FilExl, 10), 2)
                        .Cells(r_int_FilExl, 28) = Round(.Cells(r_int_FilExl, 28), 2)
                        g_rst_Genera.Close
                        Set g_rst_Genera = Nothing
                     Else
                        g_rst_Genera.Close
                        Set g_rst_Genera = Nothing
                     End If
                  Else
                     g_rst_Genera.Close
                     Set g_rst_Genera = Nothing
                  End If
              Else
                  'si es el primer pago se asigna la cuota TNC
                  .Cells(r_int_FilExl, 10) = IIf(IsNull(g_rst_Princi!SALDO_ANTEXPINI) = True, 0, g_rst_Princi!SALDO_ANTEXPINI)
                  
                  If (.Cells(r_int_FilExl, 10) = 0) Then
                      g_str_Parame = ""
                      g_str_Parame = g_str_Parame & "SELECT * "
                      g_str_Parame = g_str_Parame & "  FROM (SELECT HIPCUO_NUMCUO, HIPCUO_NUMOPE "
                      g_str_Parame = g_str_Parame & "          FROM CRE_HIPCUO C "
                      g_str_Parame = g_str_Parame & "         WHERE HIPCUO_NUMOPE = '" & g_rst_Princi!HIPMAE_NUMOPE & "' "
                      g_str_Parame = g_str_Parame & "           AND HIPCUO_TIPCRO = 1 "
                      g_str_Parame = g_str_Parame & "           AND HIPCUO_FECVCT < " & r_str_ActFin
                      g_str_Parame = g_str_Parame & "         ORDER BY HIPCUO_NUMCUO DESC) TAB "
                      g_str_Parame = g_str_Parame & " Where ROWNUM = 1 "
                      If gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
                         If Not g_rst_Genera.BOF And Not g_rst_Genera.EOF Then
                            If (g_rst_Genera!HIPCUO_NUMCUO = 1) Then
                               .Cells(r_int_FilExl, 10) = IIf(IsNull(g_rst_Princi!HIPMAE_IMPNCO) = True, 0, g_rst_Princi!HIPMAE_IMPNCO)
                            End If
                            g_rst_Genera.Close
                            Set g_rst_Genera = Nothing
                         Else
                            g_rst_Genera.Close
                            Set g_rst_Genera = Nothing
                         End If
                      End If
                  End If
                  .Cells(r_int_FilExl, 13) = IIf(IsNull(g_rst_Princi!CAPITAL_PRINCI) = True, 0, g_rst_Princi!CAPITAL_PRINCI)
                  .Cells(r_int_FilExl, 16) = IIf(IsNull(g_rst_Princi!INTERES_IMPINT) = True, 0, g_rst_Princi!INTERES_IMPINT)
                  .Cells(r_int_FilExl, 22) = IIf(IsNull(g_rst_Princi!COMISION) = True, 0, g_rst_Princi!COMISION)
                  .Cells(r_int_FilExl, 28) = IIf(IsNull(g_rst_Princi!SALDO_EXPFIN) = True, 0, g_rst_Princi!SALDO_EXPFIN)
              End If
                
              'diferencia saldo anterior
              .Cells(r_int_FilExl, 9) = IIf(IsNull(g_rst_Princi!ARCMEN_EXPINI) = True, 0, g_rst_Princi!ARCMEN_EXPINI)
              .Cells(r_int_FilExl, 11) = .Cells(r_int_FilExl, 10) - .Cells(r_int_FilExl, 9)
              r_dbl_SalCof = r_dbl_SalCof + .Cells(r_int_FilExl, 9)
              r_dbl_SalMic = r_dbl_SalMic + .Cells(r_int_FilExl, 10)
              r_dbl_SalDif = r_dbl_SalDif + .Cells(r_int_FilExl, 11)
               
              'diferencia capital
              .Cells(r_int_FilExl, 12) = IIf(IsNull(g_rst_Princi!ARCMEN_PRINCI) = True, 0, g_rst_Princi!ARCMEN_PRINCI)
              .Cells(r_int_FilExl, 14) = .Cells(r_int_FilExl, 13) - .Cells(r_int_FilExl, 12)
              r_dbl_CapCof = r_dbl_CapCof + .Cells(r_int_FilExl, 12)
              r_dbl_CapMic = r_dbl_CapMic + .Cells(r_int_FilExl, 13)
              r_dbl_CapDif = r_dbl_CapDif + .Cells(r_int_FilExl, 14)
               
              'diferencia interes
              .Cells(r_int_FilExl, 15) = IIf(IsNull(g_rst_Princi!ARCMEN_IMPINT) = True, 0, g_rst_Princi!ARCMEN_IMPINT)
              .Cells(r_int_FilExl, 17) = .Cells(r_int_FilExl, 16) - .Cells(r_int_FilExl, 15)
              r_dbl_IntCof = r_dbl_IntCof + .Cells(r_int_FilExl, 15)
              r_dbl_IntMic = r_dbl_IntMic + .Cells(r_int_FilExl, 16)
              r_dbl_IntDif = r_dbl_IntDif + .Cells(r_int_FilExl, 17)
               
              'diferencia tasacion
              .Cells(r_int_FilExl, 18) = IIf(IsNull(g_rst_Princi!ARCMEN_IMTASA) = True, 0, g_rst_Princi!ARCMEN_IMTASA)
              .Cells(r_int_FilExl, 19) = IIf(IsNull(g_rst_Princi!HIPCIE_TASCOF) = True, 0, g_rst_Princi!HIPCIE_TASCOF)
              .Cells(r_int_FilExl, 20) = .Cells(r_int_FilExl, 19) - .Cells(r_int_FilExl, 18)
              r_dbl_TasDif = r_dbl_TasDif + .Cells(r_int_FilExl, 20)
               
              'diferencia comision
              .Cells(r_int_FilExl, 21) = IIf(IsNull(g_rst_Princi!ARCMEN_COMSIN) = True, 0, g_rst_Princi!ARCMEN_COMSIN)
              .Cells(r_int_FilExl, 23) = .Cells(r_int_FilExl, 22) - .Cells(r_int_FilExl, 21)
              r_dbl_ComCof = r_dbl_ComCof + .Cells(r_int_FilExl, 21)
              r_dbl_ComMic = r_dbl_ComMic + .Cells(r_int_FilExl, 22)
              r_dbl_ComDif = r_dbl_ComDif + .Cells(r_int_FilExl, 23)
               
              'diferencia totales
              .Cells(r_int_FilExl, 24) = .Cells(r_int_FilExl, 12) + .Cells(r_int_FilExl, 15) + .Cells(r_int_FilExl, 21)
              .Cells(r_int_FilExl, 25) = .Cells(r_int_FilExl, 13) + .Cells(r_int_FilExl, 16) + .Cells(r_int_FilExl, 22)
              .Cells(r_int_FilExl, 26) = .Cells(r_int_FilExl, 25) - .Cells(r_int_FilExl, 24)
              r_dbl_TotCof = r_dbl_TotCof + .Cells(r_int_FilExl, 24)
              r_dbl_TotMic = r_dbl_TotMic + .Cells(r_int_FilExl, 25)
              r_dbl_TotDif = r_dbl_TotDif + .Cells(r_int_FilExl, 26)
               
              'diferencia saldo final
              .Cells(r_int_FilExl, 27) = IIf(IsNull(g_rst_Princi!ARCMEN_EXPFIN) = True, 0, g_rst_Princi!ARCMEN_EXPFIN)
              .Cells(r_int_FilExl, 29) = .Cells(r_int_FilExl, 28) - .Cells(r_int_FilExl, 27)
              r_dbl_FinCof = r_dbl_FinCof + .Cells(r_int_FilExl, 27)
              r_dbl_FinMic = r_dbl_FinMic + .Cells(r_int_FilExl, 28)
              r_dbl_FinDif = r_dbl_FinDif + .Cells(r_int_FilExl, 29)
                              
              r_int_FilExl = r_int_FilExl + 1
              r_int_FilItm = r_int_FilItm + 1
              g_rst_Princi.MoveNext
              DoEvents
               
              If (Not g_rst_Princi.EOF) Then
                  If (r_str_Cadena <> g_rst_Princi!ARCMEN_IDCIPR) Then
                     .Range("R" & r_int_FilExl & ":S" & r_int_FilExl).Merge
                     .Range("I" & r_int_FilExl & ":AC" & r_int_FilExl).Interior.Color = RGB(146, 208, 80)
                     
                     'escribir totales por grupo
                     .Cells(r_int_FilExl, 9) = r_dbl_SalCof: .Cells(r_int_FilExl, 10) = r_dbl_SalMic: .Cells(r_int_FilExl, 11) = r_dbl_SalDif
                     .Cells(r_int_FilExl, 12) = r_dbl_CapCof: .Cells(r_int_FilExl, 13) = r_dbl_CapMic: .Cells(r_int_FilExl, 14) = r_dbl_CapDif
                     .Cells(r_int_FilExl, 15) = r_dbl_IntCof: .Cells(r_int_FilExl, 16) = r_dbl_IntMic: .Cells(r_int_FilExl, 17) = r_dbl_IntDif
                     .Cells(r_int_FilExl, 20) = r_dbl_TasDif:
                     .Cells(r_int_FilExl, 21) = r_dbl_ComCof: .Cells(r_int_FilExl, 22) = r_dbl_ComMic: .Cells(r_int_FilExl, 23) = r_dbl_ComDif
                     .Cells(r_int_FilExl, 24) = r_dbl_TotCof: .Cells(r_int_FilExl, 25) = r_dbl_TotMic: .Cells(r_int_FilExl, 26) = r_dbl_TotDif
                     .Cells(r_int_FilExl, 27) = r_dbl_FinCof: .Cells(r_int_FilExl, 28) = r_dbl_FinMic: .Cells(r_int_FilExl, 29) = r_dbl_FinDif
                     
                     'inicializar totales por grupo
                     r_dbl_SalCof = 0: r_dbl_SalMic = 0: r_dbl_SalDif = 0: r_dbl_CapCof = 0
                     r_dbl_CapMic = 0: r_dbl_CapDif = 0: r_dbl_IntCof = 0: r_dbl_IntMic = 0
                     r_dbl_IntDif = 0: r_dbl_TasDif = 0: r_dbl_ComCof = 0: r_dbl_ComMic = 0
                     r_dbl_ComDif = 0: r_dbl_TotCof = 0: r_dbl_TotMic = 0: r_dbl_TotDif = 0
                     r_dbl_FinCof = 0: r_dbl_FinMic = 0: r_dbl_FinDif = 0
                        
                     .Range("I4:AC" & r_int_FilExl).NumberFormat = "###,###,##0.00"
                     .Range("A4:AE" & r_int_FilExl + 10).Font.Size = 8
                     .Rows("4:" & r_int_FilExl + 10).RowHeight = 12
                        
                     r_int_FilItm = 1
                     r_int_FilExl = 3
                     r_int_numhoja = r_int_numhoja + 1
                  End If
              End If
         End With
      Loop
      
      With r_obj_Excel.Sheets(r_int_numhoja)
           .Range("R" & r_int_FilExl & ":S" & r_int_FilExl).Merge
           .Range("I" & r_int_FilExl & ":AC" & r_int_FilExl).Interior.Color = RGB(146, 208, 80)
           
           'escribir totales ultimo por grupo
           .Cells(r_int_FilExl, 9) = r_dbl_SalCof:  .Cells(r_int_FilExl, 10) = r_dbl_SalMic: .Cells(r_int_FilExl, 11) = r_dbl_SalDif
           .Cells(r_int_FilExl, 12) = r_dbl_CapCof: .Cells(r_int_FilExl, 13) = r_dbl_CapMic: .Cells(r_int_FilExl, 14) = r_dbl_CapDif
           .Cells(r_int_FilExl, 15) = r_dbl_IntCof: .Cells(r_int_FilExl, 16) = r_dbl_IntMic: .Cells(r_int_FilExl, 17) = r_dbl_IntDif
           .Cells(r_int_FilExl, 20) = r_dbl_TasDif:
           .Cells(r_int_FilExl, 21) = r_dbl_ComCof: .Cells(r_int_FilExl, 22) = r_dbl_ComMic: .Cells(r_int_FilExl, 23) = r_dbl_ComDif
           .Cells(r_int_FilExl, 24) = r_dbl_TotCof: .Cells(r_int_FilExl, 25) = r_dbl_TotMic: .Cells(r_int_FilExl, 26) = r_dbl_TotDif
           .Cells(r_int_FilExl, 27) = r_dbl_FinCof: .Cells(r_int_FilExl, 28) = r_dbl_FinMic: .Cells(r_int_FilExl, 29) = r_dbl_FinDif

           .Range("I4:AC" & r_int_FilExl).NumberFormat = "###,###,##0.00"
           .Range("A4:AE" & r_int_FilExl + 10).Font.Size = 8
           .Rows("4:" & r_int_FilExl + 10).RowHeight = 12
      End With
      r_obj_Excel.Visible = True
      Set r_obj_Excel = Nothing
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   Screen.MousePointer = 0
   cmd_ExpExc.Enabled = True
End Sub

Private Sub cmb_PerMes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_PerAno)
   End If
End Sub

Private Sub ipp_PerAno_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_ExpExc)
   End If
End Sub
