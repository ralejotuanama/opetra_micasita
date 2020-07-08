VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_Rpt_Cofide_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form4"
   ClientHeight    =   2670
   ClientLeft      =   10905
   ClientTop       =   8640
   ClientWidth     =   4950
   Icon            =   "OpeTra_frm_321.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   2685
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   4965
      _Version        =   65536
      _ExtentX        =   8758
      _ExtentY        =   4736
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
            TabIndex        =   7
            Top             =   30
            Width           =   3615
            _Version        =   65536
            _ExtentX        =   6376
            _ExtentY        =   529
            _StockProps     =   15
            Caption         =   "Reporte de Saldos de Cuentas x Pagar"
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
            TabIndex        =   8
            Top             =   315
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   476
            _StockProps     =   15
            Caption         =   "COFIDE"
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
            Picture         =   "OpeTra_frm_321.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   9
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
            Picture         =   "OpeTra_frm_321.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_321.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Exportar a Excel"
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
         Height          =   1155
         Left            =   30
         TabIndex        =   10
         Top             =   1470
         Width           =   4875
         _Version        =   65536
         _ExtentX        =   8599
         _ExtentY        =   2037
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
         Begin VB.ComboBox cmb_CodPrd 
            Height          =   315
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   90
            Width           =   2775
         End
         Begin VB.ComboBox cmb_PerMes 
            Height          =   315
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   420
            Width           =   2775
         End
         Begin EditLib.fpLongInteger ipp_PerAno 
            Height          =   315
            Left            =   1530
            TabIndex        =   2
            Top             =   750
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
            Caption         =   "Producto:"
            Height          =   315
            Left            =   120
            TabIndex        =   13
            Top             =   120
            Width           =   1245
         End
         Begin VB.Label Label3 
            Caption         =   "Año:"
            Height          =   285
            Left            =   120
            TabIndex        =   12
            Top             =   810
            Width           =   1065
         End
         Begin VB.Label Label2 
            Caption         =   "Periodo:"
            Height          =   315
            Left            =   120
            TabIndex        =   11
            Top             =   450
            Width           =   1245
         End
      End
   End
End
Attribute VB_Name = "frm_Rpt_Cofide_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_str_PerMes           As String
Dim l_str_PerAno           As String

Private Sub cmd_ExpExc_Click()
   If cmb_CodPrd.ListIndex = -1 Then
      MsgBox "Debe seleccionar un Producto.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_CodPrd)
      Exit Sub
   End If
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
   Call fs_GenExc(l_str_PerMes, l_str_PerAno)
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
   
   Call gs_SetFocus(cmb_CodPrd)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   cmb_PerMes.Clear
   Call moddat_gs_Carga_LisIte_Combo(cmb_PerMes, 1, "033")
   ipp_PerAno = Mid(date, 7, 4)
         
   cmb_CodPrd.AddItem "CREDITO CME"
   cmb_CodPrd.ItemData(cmb_CodPrd.NewIndex) = CInt(3)
   cmb_CodPrd.AddItem "CREDITO MIHOGAR"
   cmb_CodPrd.ItemData(cmb_CodPrd.NewIndex) = CInt(4)
   cmb_CodPrd.AddItem "CREDITO MIVIVIENDA"
   cmb_CodPrd.ItemData(cmb_CodPrd.NewIndex) = CInt(7)
   cmb_CodPrd.AddItem "CREDITO MICASA MAS"
   cmb_CodPrd.ItemData(cmb_CodPrd.NewIndex) = CInt(19)
   cmb_CodPrd.AddItem "CREDITO COFICASA"
   cmb_CodPrd.ItemData(cmb_CodPrd.NewIndex) = CInt(20)
   cmb_CodPrd.AddItem "CREDITO MIVIVIENDA MAS BBP"
   cmb_CodPrd.ItemData(cmb_CodPrd.NewIndex) = CInt(21)
   cmb_CodPrd.AddItem "CREDITO MIVIVIENDA BBP COMPLEMENTO INICIAL"
   cmb_CodPrd.ItemData(cmb_CodPrd.NewIndex) = CInt(22)
End Sub

Private Sub fs_Limpia()
   cmb_PerMes.ListIndex = -1
End Sub

Private Sub fs_GenExc(ByVal p_PerMes As String, ByVal p_PerAno As String)
Dim r_obj_Excel      As Excel.Application
Dim r_rst_Genera     As ADODB.Recordset
Dim r_int_ConVer     As Integer
Dim r_str_EmpSeg     As String
Dim r_int_FlagTC     As Integer
Dim r_int_ConPbp     As Integer
Dim r_int_CuoPbp     As Integer
Dim r_dbl_SumPbp     As Double
   
   r_int_FlagTC = 0
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT A.HIPCIE_NUMOPE AS OPERACION, "
   g_str_Parame = g_str_Parame & "       TRIM(C.DATGEN_TIPDOC)||'-'||TRIM(C.DATGEN_NUMDOC) AS DOC_CLIENTE, "
   g_str_Parame = g_str_Parame & "       TRIM(B.HIPMAE_OPEMVI) AS OPER_MIVIVIENDA, "
   g_str_Parame = g_str_Parame & "       TRIM(B.HIPMAE_CODCOF) AS CLIENTE_COFIDE, "
   g_str_Parame = g_str_Parame & "       TRIM(C.DATGEN_APEPAT) AS APE_PATERNO, "
   g_str_Parame = g_str_Parame & "       TRIM(C.DATGEN_APEMAT) AS APE_MATERNO, "
   g_str_Parame = g_str_Parame & "       TRIM(C.DATGEN_NOMBRE) AS NOM_CLIENTE, "
   g_str_Parame = g_str_Parame & "       TRIM(D.PARDES_DESCRI) AS TIPO_MONEDA, "
   g_str_Parame = g_str_Parame & "       DECODE(NVL(E.HIPCUO_SALCAP, 0), 0, A.HIPCIE_PRENCO, E.HIPCUO_SALCAP) AS MONTO_CRONOG3, "
   g_str_Parame = g_str_Parame & "       F.HIPCUO_SALCAP AS MONTO_CRONOG5, "
   g_str_Parame = g_str_Parame & "       A.HIPCIE_PERPBP AS MONTO_PBP, "
   g_str_Parame = g_str_Parame & "       A.HIPCIE_CLACLI AS CLASIFICACION "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPCIE A "
   g_str_Parame = g_str_Parame & " INNER JOIN CRE_HIPMAE B ON B.HIPMAE_NUMOPE = A.HIPCIE_NUMOPE "
   g_str_Parame = g_str_Parame & " INNER JOIN CLI_DATGEN C ON C.DATGEN_TIPDOC = A.HIPCIE_TDOCLI AND C.DATGEN_NUMDOC = A.HIPCIE_NDOCLI "
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES D ON D.PARDES_CODGRP = 204 AND D.PARDES_CODITE = A.HIPCIE_TIPMON "
   g_str_Parame = g_str_Parame & "  LEFT JOIN CRE_HIPCUO E ON E.HIPCUO_NUMOPE = A.HIPCIE_NUMOPE AND E.HIPCUO_TIPCRO = 3 AND E.HIPCUO_FECVCT >= " & l_str_PerAno & Format(l_str_PerMes, "00") & "01" & " AND E.HIPCUO_FECVCT <= " & l_str_PerAno & Format(l_str_PerMes, "00") & Format(ff_Ultimo_Dia_Mes(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), CInt(ipp_PerAno.Text)), "00") & " "
   g_str_Parame = g_str_Parame & "  LEFT JOIN CRE_HIPCUO F ON F.HIPCUO_NUMOPE = A.HIPCIE_NUMOPE AND F.HIPCUO_TIPCRO = 5 AND F.HIPCUO_FECVCT >= " & l_str_PerAno & Format(l_str_PerMes, "00") & "01" & " AND F.HIPCUO_FECVCT <= " & l_str_PerAno & Format(l_str_PerMes, "00") & Format(ff_Ultimo_Dia_Mes(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), CInt(ipp_PerAno.Text)), "00") & " "
   g_str_Parame = g_str_Parame & " WHERE A.HIPCIE_PERANO = " & CStr(l_str_PerAno) & " "
   g_str_Parame = g_str_Parame & "   AND A.HIPCIE_PERMES = " & CStr(l_str_PerMes) & " "
   If cmb_CodPrd.ItemData(cmb_CodPrd.ListIndex) = 3 Then
      r_int_FlagTC = 0
      g_str_Parame = g_str_Parame & "  AND HIPCIE_CODPRD IN (" & moddat_g_str_AgrCME & ") "
   ElseIf cmb_CodPrd.ItemData(cmb_CodPrd.ListIndex) = 4 Then
      r_int_FlagTC = 1
      g_str_Parame = g_str_Parame & "  AND HIPCIE_CODPRD IN (" & moddat_g_str_AgrMIHG & ") "
   ElseIf cmb_CodPrd.ItemData(cmb_CodPrd.ListIndex) = 7 Then
      r_int_FlagTC = 1
      g_str_Parame = g_str_Parame & "  AND HIPCIE_CODPRD IN (" & moddat_g_str_Agr1FMV & ") "
   ElseIf cmb_CodPrd.ItemData(cmb_CodPrd.ListIndex) = 19 Then
      r_int_FlagTC = 2
      g_str_Parame = g_str_Parame & "  AND HIPCIE_CODPRD IN (" & moddat_g_str_Agr2FMV & ") "
   End If
   g_str_Parame = g_str_Parame & " ORDER BY A.HIPCIE_NUMOPE "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      MsgBox "No se encontro información para el período seleccionado.", vbExclamation, modgen_g_str_NomPlt
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      r_int_ConVer = 1
      .Range(.Cells(r_int_ConVer, 1), .Cells(r_int_ConVer, 1)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(r_int_ConVer, 1), .Cells(r_int_ConVer, 1)).Font.Bold = True
      .Range(.Cells(r_int_ConVer, 1), .Cells(r_int_ConVer, 1)).Font.Underline = xlUnderlineStyleSingle
      .Range("A" & r_int_ConVer & ":M" & r_int_ConVer & "").Merge
      .Cells(r_int_ConVer, 1) = "SALDOS POR PAGAR A CODIFDE  -  " & Trim(cmb_PerMes.Text) & " / " & ipp_PerAno.Text
      
      r_int_ConVer = 3
      .Cells(r_int_ConVer, 1) = "ITEM"
      .Cells(r_int_ConVer, 2) = "NRO. OPERACION"
      .Cells(r_int_ConVer, 3) = "DOI CLIENTE"
      .Cells(r_int_ConVer, 4) = "COD. OPER. MIVIVIENDA"
      .Cells(r_int_ConVer, 5) = "COD. CLIENTE COFIDE"
      .Cells(r_int_ConVer, 6) = "APELLIDO PATERNO"
      .Cells(r_int_ConVer, 7) = "APELLIDO MATERNO"
      .Cells(r_int_ConVer, 8) = "NOMBRE"
      .Cells(r_int_ConVer, 9) = "TIPO MONEDA"
      .Cells(r_int_ConVer, 10) = "MTO. COFIDE TNC"
      .Cells(r_int_ConVer, 11) = "MTO. COFIDE TC"
      .Cells(r_int_ConVer, 12) = "MTO. PBP PERDIDO"
      .Cells(r_int_ConVer, 13) = "CLASIFICACION"
      
      .Range(.Cells(r_int_ConVer, 1), .Cells(r_int_ConVer, 13)).Font.Bold = True
      .Range(.Cells(r_int_ConVer, 1), .Cells(r_int_ConVer, 13)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(r_int_ConVer, 1), .Cells(r_int_ConVer, 13)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(r_int_ConVer, 1), .Cells(r_int_ConVer, 13)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(r_int_ConVer, 1), .Cells(r_int_ConVer, 13)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(r_int_ConVer, 1), .Cells(r_int_ConVer, 13)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(r_int_ConVer, 1), .Cells(r_int_ConVer, 13)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(r_int_ConVer, 1), .Cells(r_int_ConVer, 13)).Borders(xlInsideVertical).LineStyle = xlContinuous
      
      .Columns("A").ColumnWidth = 4
      .Columns("B").ColumnWidth = 14
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 12
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 18
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("D").NumberFormat = "@"
      .Columns("E").ColumnWidth = 18
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").ColumnWidth = 20
      .Columns("G").ColumnWidth = 20
      .Columns("H").ColumnWidth = 20
      .Columns("I").ColumnWidth = 18
      .Columns("I").HorizontalAlignment = xlHAlignCenter
      .Columns("J").ColumnWidth = 16
      .Columns("J").NumberFormat = "###,###,##0.00"
      .Columns("K").ColumnWidth = 16
      .Columns("K").NumberFormat = "###,###,##0.00"
      .Columns("L").ColumnWidth = 16
      .Columns("L").NumberFormat = "###,###,##0.00"
      .Columns("M").ColumnWidth = 14
      .Columns("M").HorizontalAlignment = xlHAlignCenter
      
      g_rst_Princi.MoveFirst
      r_int_ConVer = 4
      Do While Not g_rst_Princi.EOF
         .Cells(r_int_ConVer, 1) = r_int_ConVer - 3
         .Cells(r_int_ConVer, 2) = gf_Formato_NumOpe(g_rst_Princi!OPERACION)
         .Cells(r_int_ConVer, 3) = Trim(g_rst_Princi!DOC_CLIENTE)
         .Cells(r_int_ConVer, 4) = Trim(g_rst_Princi!OPER_MIVIVIENDA)
         .Cells(r_int_ConVer, 5) = Trim(g_rst_Princi!CLIENTE_COFIDE)
         .Cells(r_int_ConVer, 6) = Trim(g_rst_Princi!APE_PATERNO)
         .Cells(r_int_ConVer, 7) = Trim(g_rst_Princi!APE_MATERNO)
         .Cells(r_int_ConVer, 8) = Trim(g_rst_Princi!NOM_CLIENTE)
         .Cells(r_int_ConVer, 9) = Trim(g_rst_Princi!TIPO_MONEDA)
         If r_int_FlagTC = 0 Then
            .Cells(r_int_ConVer, 10) = gf_FormatoNumero(g_rst_Princi!MONTO_CRONOG5, 12, 2)
         ElseIf r_int_FlagTC = 1 Then
            .Cells(r_int_ConVer, 10) = gf_FormatoNumero(g_rst_Princi!MONTO_CRONOG3, 12, 2)
         ElseIf r_int_FlagTC = 2 Then
            .Cells(r_int_ConVer, 10) = gf_FormatoNumero(g_rst_Princi!MONTO_CRONOG3, 12, 2)
         Else
            .Cells(r_int_ConVer, 10) = "0.00"
         End If
         .Cells(r_int_ConVer, 11) = gf_FormatoNumero(0, 12, 2)
         .Cells(r_int_ConVer, 12) = gf_FormatoNumero(0, 12, 2)
         .Cells(r_int_ConVer, 13) = Trim(g_rst_Princi!CLASIFICACION)
         
         'Carga Saldo TC
         If r_int_FlagTC = 1 Then
            g_str_Parame = ""
            g_str_Parame = g_str_Parame & "SELECT MAX(HIPCUO_SALCAP + HIPCUO_CAPITA) AS CAPITA_TC, MAX(HIPCUO_CUOBBP) AS CUOTA_BBP "
            g_str_Parame = g_str_Parame & "  FROM CRE_HIPCUO "
            g_str_Parame = g_str_Parame & " WHERE HIPCUO_NUMOPE = '" & g_rst_Princi!OPERACION & "' "
            g_str_Parame = g_str_Parame & "   AND HIPCUO_TIPCRO = 4 "
            g_str_Parame = g_str_Parame & "   AND HIPCUO_FECVCT >= " & l_str_PerAno & Format(l_str_PerMes, "00") & "01"
            
            If gf_EjecutaSQL(g_str_Parame, r_rst_Genera, 3) Then
               If Not (r_rst_Genera.BOF And r_rst_Genera.EOF) Then
                  .Cells(r_int_ConVer, 11) = gf_FormatoNumero(r_rst_Genera!CAPITA_TC, 12, 2)
               End If
            End If
            
            r_rst_Genera.Close
            Set r_rst_Genera = Nothing
            
            'Carga PBP Perdido
            g_str_Parame = ""
            g_str_Parame = g_str_Parame & "SELECT HIPCUO_NUMOPE, HIPCUO_FECVCT, HIPCUO_CUOBBP, HIPCUO_CAPBBP "
            g_str_Parame = g_str_Parame & "  FROM CRE_HIPCUO "
            g_str_Parame = g_str_Parame & " WHERE HIPCUO_NUMOPE = '" & g_rst_Princi!OPERACION & "' "
            g_str_Parame = g_str_Parame & "   AND HIPCUO_TIPCRO = 1 "
            g_str_Parame = g_str_Parame & "   AND HIPCUO_FECVCT >= " & l_str_PerAno & Format(l_str_PerMes, "00") & Format(ff_Ultimo_Dia_Mes(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), CInt(ipp_PerAno.Text)), "00")
            g_str_Parame = g_str_Parame & "   AND HIPCUO_CUOBBP > 0 "
            g_str_Parame = g_str_Parame & " ORDER BY HIPCUO_FECVCT "
            
            If gf_EjecutaSQL(g_str_Parame, r_rst_Genera, 3) Then
               If Not (r_rst_Genera.BOF And r_rst_Genera.EOF) Then
                  r_rst_Genera.MoveFirst
                  r_int_ConPbp = 0
                  r_dbl_SumPbp = 0
                  r_int_CuoPbp = r_rst_Genera!HIPCUO_CUOBBP
                  
                  Do While Not r_rst_Genera.EOF
                     If r_int_CuoPbp <> r_rst_Genera!HIPCUO_CUOBBP Then
                        Exit Do
                     End If
                     r_dbl_SumPbp = r_dbl_SumPbp + r_rst_Genera!HIPCUO_CAPBBP
                     r_int_ConPbp = r_int_ConPbp + 1
                     r_rst_Genera.MoveNext
                  Loop
                  
                  If r_int_ConPbp > 0 And r_int_ConPbp < 6 Then
                     .Cells(r_int_ConVer, 12) = gf_FormatoNumero(r_dbl_SumPbp, 12, 2)
                  End If
               End If
            End If
            
            r_rst_Genera.Close
            Set r_rst_Genera = Nothing
         End If
         
         'Siguiente registro
         r_int_ConVer = r_int_ConVer + 1
         g_rst_Princi.MoveNext
         DoEvents
      Loop
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      .Range(.Cells(1, 1), .Cells(r_int_ConVer, 13)).Font.Name = "Arial"
      .Range(.Cells(1, 1), .Cells(r_int_ConVer, 13)).Font.Size = 8
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

