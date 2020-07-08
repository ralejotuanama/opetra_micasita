VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_Rpt_MviCof_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   2280
   ClientLeft      =   5940
   ClientTop       =   4005
   ClientWidth     =   5475
   Icon            =   "OpeTra_frm_295.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   2265
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   5475
      _Version        =   65536
      _ExtentX        =   9657
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
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   6
         Top             =   30
         Width           =   5385
         _Version        =   65536
         _ExtentX        =   9499
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
            Left            =   630
            TabIndex        =   7
            Top             =   30
            Width           =   4605
            _Version        =   65536
            _ExtentX        =   8123
            _ExtentY        =   979
            _StockProps     =   15
            Caption         =   "Reporte de Asignación de Premio Buen Pagador"
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
            Picture         =   "OpeTra_frm_295.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   8
         Top             =   750
         Width           =   5385
         _Version        =   65536
         _ExtentX        =   9499
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
            Picture         =   "OpeTra_frm_295.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_295.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   4770
            Picture         =   "OpeTra_frm_295.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   2040
            Top             =   120
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
         Height          =   765
         Left            =   30
         TabIndex        =   9
         Top             =   1440
         Width           =   5355
         _Version        =   65536
         _ExtentX        =   9446
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
            Width           =   4035
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
         Begin VB.Label Label1 
            Caption         =   "Mes:"
            Height          =   255
            Left            =   60
            TabIndex        =   11
            Top             =   60
            Width           =   915
         End
         Begin VB.Label Label2 
            Caption         =   "Año:"
            Height          =   255
            Left            =   60
            TabIndex        =   10
            Top             =   390
            Width           =   915
         End
      End
   End
End
Attribute VB_Name = "frm_Rpt_MviCof_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_ExpExc_Click()
   If cmb_PerMes.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Mes.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PerMes)
      Exit Sub
   End If

   g_str_Parame = "SELECT * FROM CRE_CABPBP WHERE "
   g_str_Parame = g_str_Parame & "CABPBP_PERMES = " & CStr(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)) & " AND "
   g_str_Parame = g_str_Parame & "CABPBP_PERANO = " & ipp_PerAno.Text & " "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      MsgBox "No se pudo leer la tabla CRE_CABPBP.", vbCritical, modgen_g_str_NomPlt
      Exit Sub
   End If

   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "El Período no ha sido Evaluado.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   Else
      If g_rst_Princi!CABPBP_SITUAC = 1 Then
         g_rst_Princi.Close
         Set g_rst_Princi = Nothing
         
         MsgBox "El Período se encuentra en Evaluación.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   End If

   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_ExpExcel_01
   Call fs_ExpExcel_02
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Imprim_Click()
   If cmb_PerMes.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Mes.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_PerMes)
      Exit Sub
   End If

   g_str_Parame = "SELECT * FROM CRE_CABPBP WHERE "
   g_str_Parame = g_str_Parame & "CABPBP_PERMES = " & CStr(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)) & " AND "
   g_str_Parame = g_str_Parame & "CABPBP_PERANO = " & ipp_PerAno.Text & " "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      MsgBox "No se pudo leer la tabla CRE_CABPBP.", vbCritical, modgen_g_str_NomPlt
      Exit Sub
   End If

   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
     g_rst_Princi.Close
     Set g_rst_Princi = Nothing
     
     MsgBox "El Período no ha sido Evaluado.", vbExclamation, modgen_g_str_NomPlt
     
     Exit Sub
   Else
      If g_rst_Princi!CABPBP_SITUAC = 1 Then
         g_rst_Princi.Close
         Set g_rst_Princi = Nothing
         
         MsgBox "El Período se encuentra en Evaluación.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   End If

   If MsgBox("¿Está seguro de imprimir el reporte?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If

   'Se conecta al crystal report
   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
   
   'Se envia las tablas correspondientes en el orden que fueron utilizadas
   crp_Imprim.DataFiles(0) = "CRE_DETPBP"
   crp_Imprim.DataFiles(1) = "CRE_HIPMAE"
   crp_Imprim.DataFiles(2) = "CLI_DATGEN"
   crp_Imprim.DataFiles(3) = "CRE_PRODUC"
   
      
   'Se pone la llamada del nombre del reporte y se escoge donde se destinara el reporte
   crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "OPE_EVAPBP_11.RPT"
   crp_Imprim.SelectionFormula = "{CRE_DETPBP.DETPBP_PERMES} = " & CStr(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)) & " AND "
   crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{CRE_DETPBP.DETPBP_PERANO} = " & ipp_PerAno.Text & " "
   crp_Imprim.WindowShowPrintSetupBtn = True
   crp_Imprim.Action = 1
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
   Call moddat_gs_Carga_LisIte_Combo(cmb_PerMes, 1, "033")
End Sub

Private Sub fs_Limpia()
   cmb_PerMes.ListIndex = -1
   ipp_PerAno.Text = Format(Year(date), "0000")
End Sub

Private Sub fs_ExpExcel_01()
   Dim r_obj_Excel         As Excel.Application
   Dim r_obj_NomArc        As New Excel.Workbook
   Dim r_obj_NomHoj        As New Excel.Worksheet
   Dim r_str_Parame        As String
   Dim r_int_ConVer        As Integer
   Dim r_rst_Princi        As ADODB.Recordset
   Dim r_rst_Genera        As ADODB.Recordset
   Dim r_int_PerMes        As Integer
   Dim r_str_FecIni        As String
   Dim r_str_FecFin        As String
   
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & "SELECT DETPBP_PERMES, DETPBP_PERANO, HIPMAE_CODPRD, DETPBP_FLGPBP, HIPMAE_NUMOPE, HIPMAE_TDOCLI, HIPMAE_NDOCLI, HIPMAE_OPEMVI,"
   r_str_Parame = r_str_Parame & "       DETPBP_CUOCON, DETPBP_CAPCLI, DETPBP_INTCLI, DETPBP_CAPADE, DETPBP_INTADE, DETPBP_COMADE, DETPBP_CEVAIN, DETPBP_CEVAFN,"
   r_str_Parame = r_str_Parame & "       DETPBP_CIPNCL, DETPBP_CFPNCL, TRIM(E.PARDES_DESCRI) AS SITUACION_PBP, TRIM(D.PRODUC_DESCRI) AS PRODUCTO,"
   r_str_Parame = r_str_Parame & "       TRIM(C.DATGEN_APEPAT)||' '||TRIM(C.DATGEN_APEMAT)||' '||TRIM(C.DATGEN_NOMBRE) AS NOMBRE_CLIENTE, HIPMAE_OPEMV1"
   r_str_Parame = r_str_Parame & "  FROM CRE_DETPBP A"
   r_str_Parame = r_str_Parame & " INNER JOIN CRE_HIPMAE B ON DETPBP_NUMOPE = HIPMAE_NUMOPE"
   r_str_Parame = r_str_Parame & " INNER JOIN CLI_DATGEN C ON C.DATGEN_TIPDOC = B.HIPMAE_TDOCLI AND C.DATGEN_NUMDOC = B.HIPMAE_NDOCLI"
   r_str_Parame = r_str_Parame & " INNER JOIN CRE_PRODUC D ON PRODUC_CODIGO = B.HIPMAE_CODPRD"
   r_str_Parame = r_str_Parame & " INNER JOIN MNT_PARDES E ON E.PARDES_CODGRP = 275 AND E.PARDES_CODITE = DETPBP_FLGPBP"
   r_str_Parame = r_str_Parame & " Where DETPBP_PERMES = " & CStr(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   r_str_Parame = r_str_Parame & "   AND DETPBP_PERANO = " & ipp_PerAno.Text
   r_str_Parame = r_str_Parame & " ORDER BY PRODUCTO ASC, NOMBRE_CLIENTE ASC"
   
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Princi, 3) Then
       Exit Sub
   End If
   
   If r_rst_Princi.BOF And r_rst_Princi.EOF Then
      r_rst_Princi.Close
      Set r_rst_Princi = Nothing
      MsgBox "No se encontraron datos a reportar.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If

   r_int_PerMes = CStr(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   r_str_FecIni = Format(CDate("01/" & Format(r_int_PerMes, "00") & "/" & Format(ipp_PerAno.Text, "0000")), "yyyymmdd")
   r_str_FecFin = Format(CDate(Format(ff_Ultimo_Dia_Mes(r_int_PerMes, ipp_PerAno.Text), "00") & "/" & Format(r_int_PerMes, "00") & "/" & Format(ipp_PerAno.Text, "0000")), "yyyymmdd")
      
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & "SELECT A.HIPCUO_NUMOPE FROM CRE_HIPCUO A "
   r_str_Parame = r_str_Parame & " WHERE HIPCUO_NUMOPE IN (SELECT HIPMAE_NUMOPE FROM CRE_HIPMAE "
   r_str_Parame = r_str_Parame & "                          WHERE HIPMAE_SITUAC = 9 "
   r_str_Parame = r_str_Parame & "                            AND HIPMAE_FECCAN >= " & r_str_FecIni & " AND HIPMAE_FECCAN <= " & r_str_FecFin & " AND HIPMAE_CUOPEN = 0 AND HIPMAE_CODPRD IN ('003','004','007','009','010','012','013','014','015','016','017','018')) "
   r_str_Parame = r_str_Parame & "   AND HIPCUO_TIPCRO = 4 "
   r_str_Parame = r_str_Parame & "   AND HIPCUO_FECVCT >= " & r_str_FecIni
   r_str_Parame = r_str_Parame & "   AND HIPCUO_FECVCT <= " & r_str_FecFin
   
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Genera, 3) Then
      Exit Sub
   End If
            
   'Generando Excel
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.DisplayAlerts = False
   Set r_obj_NomArc = r_obj_Excel.Workbooks.Add
   Set r_obj_NomHoj = r_obj_NomArc.Worksheets("Hoja1")
   
   With r_obj_NomArc.ActiveSheet
      .Cells(1, 1) = "ITEM":                             .Columns("A").ColumnWidth = 8
      .Cells(1, 2) = "PERIODO":                          .Columns("B").ColumnWidth = 11:        .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 3) = "PRODUCTO":                         .Columns("C").ColumnWidth = 32:        .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 4) = "SITUACION PBP":                    .Columns("D").ColumnWidth = 12:        .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 5) = "NRO. OPERACION":                   .Columns("E").ColumnWidth = 14:        .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 6) = "CLIENTE":                          .Columns("F").ColumnWidth = 40
      .Cells(1, 7) = "NRO. OPERACION MIVIVIENDA":        .Columns("G").ColumnWidth = 24:        .Columns("G").HorizontalAlignment = xlHAlignCenter:      .Columns("G").NumberFormat = "@"
      .Cells(1, 8) = "NRO. OPERACION COFIDE":            .Columns("H").ColumnWidth = 21:        .Columns("H").HorizontalAlignment = xlHAlignCenter:      .Columns("H").NumberFormat = "@"
      .Cells(1, 9) = "CUOTA TC":                         .Columns("I").ColumnWidth = 10:        .Columns("I").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 10) = "CAPITAL TRAMO CLIENTE":           .Columns("J").ColumnWidth = 20:        .Columns("J").NumberFormat = "###,##0.00"
      .Cells(1, 11) = "INTERES TRAMO CLIENTE":           .Columns("K").ColumnWidth = 21:        .Columns("K").NumberFormat = "###,##0.00"
      .Cells(1, 12) = "TOTAL CUOTA TRAMO CLIENTE":       .Columns("L").ColumnWidth = 25:        .Columns("L").NumberFormat = "###,##0.00"
      .Cells(1, 13) = "CAPITAL TRAMO COFIDE/MVI":        .Columns("M").ColumnWidth = 25:        .Columns("M").NumberFormat = "###,##0.00"
      .Cells(1, 14) = "INTERES TRAMO COFIDE/MVI":        .Columns("N").ColumnWidth = 24:        .Columns("N").NumberFormat = "###,##0.00"
      .Cells(1, 15) = "COMISION TRAMO COFIDE/MVI":       .Columns("O").ColumnWidth = 25:        .Columns("O").NumberFormat = "###,##0.00"
      .Cells(1, 16) = "TOTAL CUOTA TRAMO COFIDE/MVI":    .Columns("P").ColumnWidth = 28:        .Columns("P").NumberFormat = "###,##0.00"
      .Cells(1, 17) = "CUOTA INICIO EVALUACION":         .Columns("Q").ColumnWidth = 23:        .Columns("Q").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 18) = "CUOTA FIN EVALUACION":            .Columns("R").ColumnWidth = 20:        .Columns("R").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 19) = "CUOTA INICIO PENALIDAD":          .Columns("S").ColumnWidth = 21:        .Columns("S").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 20) = "CUOTA FIN PENALIDAD":             .Columns("T").ColumnWidth = 19:        .Columns("T").HorizontalAlignment = xlHAlignCenter
      
      .Range(.Cells(1, 1), .Cells(1, 20)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 20)).HorizontalAlignment = xlHAlignCenter
   End With
   
   r_rst_Princi.MoveFirst
   r_int_ConVer = 2
   Do While Not r_rst_Princi.EOF
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 1
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = Format(r_rst_Princi!DETPBP_PERMES, "00") & " - " & Format(r_rst_Princi!DETPBP_PERANO, "0000")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = r_rst_Princi!PRODUCTO
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = r_rst_Princi!SITUACION_PBP
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = gf_Formato_NumOpe(r_rst_Princi!HIPMAE_NUMOPE)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = r_rst_Princi!NOMBRE_CLIENTE
      
      If r_rst_Princi!HIPMAE_CODPRD = "001" Or r_rst_Princi!HIPMAE_CODPRD = "003" Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = Trim(r_rst_Princi!HIPMAE_OPEMVI & "")
      End If
      
      If r_rst_Princi!HIPMAE_CODPRD = "003" Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = Trim(r_rst_Princi!HIPMAE_OPEMV1 & "")
      ElseIf r_rst_Princi!HIPMAE_CODPRD <> "001" And r_rst_Princi!HIPMAE_CODPRD <> "006" Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = Trim(r_rst_Princi!HIPMAE_OPEMVI & "")
      End If
      
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = CStr(r_rst_Princi!DETPBP_CUOCON)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = r_rst_Princi!DETPBP_CAPCLI
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = r_rst_Princi!DETPBP_INTCLI
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = r_rst_Princi!DETPBP_CAPCLI + r_rst_Princi!DETPBP_INTCLI
      
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = r_rst_Princi!DETPBP_CAPADE
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14) = r_rst_Princi!DETPBP_INTADE
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = r_rst_Princi!DETPBP_COMADE
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 16) = r_rst_Princi!DETPBP_CAPADE + r_rst_Princi!DETPBP_INTADE + r_rst_Princi!DETPBP_COMADE
      
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 17) = r_rst_Princi!DETPBP_CEVAIN
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 18) = r_rst_Princi!DETPBP_CEVAFN
      
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 19) = r_rst_Princi!DETPBP_CIPNCL
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 20) = r_rst_Princi!DETPBP_CFPNCL
      
      '-----------------------------
      If Not (r_rst_Genera.BOF And r_rst_Genera.EOF) Then
         r_rst_Genera.MoveFirst
         Do While Not r_rst_Genera.EOF
            If Not (r_rst_Genera.BOF And r_rst_Genera.EOF) Then
               If r_rst_Genera!HIPCUO_NUMOPE = r_rst_Princi!HIPMAE_NUMOPE Then
                  r_obj_Excel.Sheets(1).Range("A" & r_int_ConVer & ":T" & r_int_ConVer).Interior.Color = RGB(146, 208, 80)
                  Exit Do
               End If
               r_rst_Genera.MoveNext
            End If
         Loop
      End If
      '-----------------------------
      r_int_ConVer = r_int_ConVer + 1
      
      r_rst_Princi.MoveNext
   Loop
      
   r_rst_Princi.Close
   Set r_rst_Princi = Nothing
      
      
   r_rst_Genera.Close
   Set r_rst_Genera = Nothing
      
   'Ordenando por Producto y Cliente
   r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(1, 1), r_obj_Excel.ActiveSheet.Cells(r_int_ConVer - 1, 20)).Font.Size = 9
   'r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(1, 2), r_obj_Excel.Cells(r_int_ConVer - 1, 20)).Sort r_obj_Excel.Range("C1"), xlAscending, r_obj_Excel.Range("D1"), , xlAscending, r_obj_Excel.Range("F1"), , xlAscending, , , xlYes
   r_obj_Excel.Visible = True
   
   Set r_obj_Excel = Nothing
End Sub

Private Sub fs_ExpExcel_02()
   Dim r_obj_Excel         As Excel.Application
   Dim r_obj_NomArc        As New Excel.Workbook
   Dim r_obj_NomHoj        As New Excel.Worksheet
   Dim r_str_Parame        As String
   Dim r_int_ConVer        As Integer
   Dim r_rst_Princi        As ADODB.Recordset
   Dim r_rst_Genera        As ADODB.Recordset
   Dim r_int_PerMes        As Integer
   Dim r_int_PerAno        As Integer
   Dim r_int_UltDia        As Integer
   Dim r_str_CadAux        As String
   Dim r_str_Cadena        As String
   
   On Error GoTo Error_Excel
   
   r_str_Parame = ""
   r_str_Parame = r_str_Parame & " SELECT DETPBP_PERMES, DETPBP_PERANO, '10005339' AS CODIGO_IFI, 'EDPYME MICASITA S.A.' AS NOMBRE_IFI, SUBSTR(HIPMAE_OPEMVI,1,3) AS CLASE_PRODUCTO,"
   r_str_Parame = r_str_Parame & "        TRIM(DECODE(INSTR(HIPMAE_OPEMVI,'-',1),0, HIPMAE_OPEMVI, SUBSTR(HIPMAE_OPEMVI,1,INSTR(HIPMAE_OPEMVI,'-',1)-1))) AS NRO_PRESTAMO, TRIM(SUBSTR(HIPMAE_OPEMVI,INSTR(HIPMAE_OPEMVI,'-',1) + 1)) AS CODIGO, "
   r_str_Parame = r_str_Parame & "        TRIM(HIPMAE_CODCOF) AS CODIGO_PRESTAMO_FINAL, TRIM(C.DATGEN_NOMBRE) AS NOMBRE_CLIENTE, TRIM(C.DATGEN_APEPAT)||' '||TRIM(C.DATGEN_APEMAT) AS APELLIDO_CLIENTE,"
   r_str_Parame = r_str_Parame & "        (CASE WHEN B.HIPMAE_TDOCLI = 5 THEN 'PE0' /*Pasaporte*/"
   r_str_Parame = r_str_Parame & "              WHEN B.HIPMAE_TDOCLI = 7 THEN 'PE1' /*Perú: RUC*/"
   r_str_Parame = r_str_Parame & "              WHEN B.HIPMAE_TDOCLI = 1 THEN 'PE2' /*DNI*/"
   r_str_Parame = r_str_Parame & "              WHEN B.HIPMAE_TDOCLI = 3 OR B.HIPMAE_TDOCLI = 4 THEN 'PE3' /*Carnet de FFPP*/"
   r_str_Parame = r_str_Parame & "              WHEN B.HIPMAE_TDOCLI = 2 THEN 'PE4'/*Carnet de Extranjería*/"
   r_str_Parame = r_str_Parame & "         END) AS TIPO_DOCUMENTO,"
   r_str_Parame = r_str_Parame & "        HIPMAE_NDOCLI AS NRO_DOCUMENTO,"
   r_str_Parame = r_str_Parame & "        DECODE(TRIM(E.PARDES_DESCRI),'NO','X','') AS PREMIO"
   r_str_Parame = r_str_Parame & "   FROM CRE_DETPBP A"
   r_str_Parame = r_str_Parame & "  INNER JOIN CRE_HIPMAE B ON DETPBP_NUMOPE = HIPMAE_NUMOPE"
   r_str_Parame = r_str_Parame & "  INNER JOIN CLI_DATGEN C ON C.DATGEN_TIPDOC = B.HIPMAE_TDOCLI AND C.DATGEN_NUMDOC = B.HIPMAE_NDOCLI"
   r_str_Parame = r_str_Parame & "  INNER JOIN CRE_PRODUC D ON PRODUC_CODIGO = B.HIPMAE_CODPRD"
   r_str_Parame = r_str_Parame & "  INNER JOIN MNT_PARDES E ON E.PARDES_CODGRP = 275 AND E.PARDES_CODITE = DETPBP_FLGPBP"
   r_str_Parame = r_str_Parame & "  Where DETPBP_PERMES = " & CStr(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   r_str_Parame = r_str_Parame & "    AND DETPBP_PERANO = " & ipp_PerAno.Text
   r_str_Parame = r_str_Parame & "  ORDER BY APELLIDO_CLIENTE ASC, NOMBRE_CLIENTE ASC"
   
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Princi, 3) Then
       Exit Sub
   End If
               
               
   'Generando Excel
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.DisplayAlerts = False
   Set r_obj_NomArc = r_obj_Excel.Workbooks.Add
   Set r_obj_NomHoj = r_obj_NomArc.Worksheets("Hoja1")
   
   With r_obj_NomArc.ActiveSheet
     '.Cells(1, 1) = "ITEM":                        .Columns("A").ColumnWidth = 9:    .Columns("A").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 1) = "PERIODO ANUAL":               .Columns("A").ColumnWidth = 13:   .Columns("A").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 2) = "PERIODO MENSUAL":             .Columns("B").ColumnWidth = 15:   .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 3) = "CODIGO IFI":                  .Columns("C").ColumnWidth = 12:   .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 4) = "NOMBRE IFI":                  .Columns("D").ColumnWidth = 25:   .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 5) = "CLASE PRODUCTO":              .Columns("E").ColumnWidth = 15:   .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 6) = "NRO. PRESTAMO":               .Columns("F").ColumnWidth = 22:   .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 7) = "CODIGO PRESTAMO FINAL":       .Columns("G").ColumnWidth = 22:   .Columns("G").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 8) = "NOMBRES":                     .Columns("H").ColumnWidth = 25:   .Columns("H").HorizontalAlignment = xlHAlignLeft
      .Cells(1, 9) = "APELLIDO PAT. Y MATERNO":     .Columns("I").ColumnWidth = 25:   .Columns("I").HorizontalAlignment = xlHAlignLeft
      .Cells(1, 10) = "TIPO DOCUMENTO":             .Columns("J").ColumnWidth = 16:   .Columns("J").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 11) = "NRO DOCUMENTO":              .Columns("K").ColumnWidth = 18:   .Columns("K").HorizontalAlignment = xlHAlignCenter
      .Cells(1, 12) = "PREMIO":                     .Columns("L").ColumnWidth = 9:    .Columns("L").HorizontalAlignment = xlHAlignCenter
      
      .Range(.Cells(1, 1), .Cells(1, 12)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 12)).HorizontalAlignment = xlHAlignCenter
   End With
   
   r_rst_Princi.MoveFirst
   r_int_ConVer = 2
   Do While Not r_rst_Princi.EOF
      'r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 1
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = "'" & Format(r_rst_Princi!DETPBP_PERANO, "0000")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = "'" & Format(r_rst_Princi!DETPBP_PERMES, "00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = "'" & r_rst_Princi!CODIGO_IFI
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = "'" & r_rst_Princi!NOMBRE_IFI
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = "'" & r_rst_Princi!CLASE_PRODUCTO
      
      r_str_Cadena = Mid(r_rst_Princi!NRO_PRESTAMO, 1, Len(r_rst_Princi!NRO_PRESTAMO) - Len(r_rst_Princi!Codigo))
      r_str_Cadena = r_str_Cadena & r_rst_Princi!Codigo
      
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = "'" & r_str_Cadena 'r_rst_Princi!NRO_PRESTAMO
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = "'" & r_rst_Princi!CODIGO_PRESTAMO_FINAL
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = "'" & r_rst_Princi!NOMBRE_CLIENTE
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = "'" & r_rst_Princi!APELLIDO_CLIENTE
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = "'" & r_rst_Princi!TIPO_DOCUMENTO
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = "'" & r_rst_Princi!NRO_DOCUMENTO
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = "'" & r_rst_Princi!PREMIO
      
      r_int_ConVer = r_int_ConVer + 1
      r_rst_Princi.MoveNext
   Loop
      
   r_rst_Princi.Close
   Set r_rst_Princi = Nothing
            
   'Ordenando por Producto y Cliente
   r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(1, 1), r_obj_Excel.ActiveSheet.Cells(r_int_ConVer - 1, 12)).Font.Size = 9
   r_int_PerMes = CStr(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   r_int_PerAno = ipp_PerAno.Text
   r_int_UltDia = ff_Ultimo_Dia_Mes(r_int_PerMes, r_int_PerAno)
   r_str_CadAux = moddat_g_str_RutLoc & "\" & "10005339-" & Format(r_int_UltDia, "00") & Format(r_int_PerMes, "00") & Format(r_int_PerAno, "0000") & ".XLS"
   r_obj_Excel.ActiveWorkbook.SaveAs (r_str_CadAux)
   
   r_obj_Excel.Application.Quit
   Set r_obj_Excel = Nothing
   MsgBox "Archivo excel generado correctamente: " & r_str_CadAux, vbInformation, modgen_g_str_NomPlt
   Exit Sub
   
Error_Excel:
   MsgBox "Error al generar archivo: " & Err.Description, vbExclamation, modgen_g_str_NomPlt
End Sub


