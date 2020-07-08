VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_Rpt_Cofide_03 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4965
   Icon            =   "OpeTra_frm_341.frx":0000
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
      TabIndex        =   0
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
         TabIndex        =   1
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
            TabIndex        =   2
            Top             =   30
            Width           =   3615
            _Version        =   65536
            _ExtentX        =   6376
            _ExtentY        =   529
            _StockProps     =   15
            Caption         =   "Reporte Comparativo de Saldos"
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
            TabIndex        =   3
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
            Picture         =   "OpeTra_frm_341.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   30
         TabIndex        =   4
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
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_341.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   4260
            Picture         =   "OpeTra_frm_341.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   885
         Left            =   30
         TabIndex        =   7
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
            ItemData        =   "OpeTra_frm_341.frx":0A62
            Left            =   1530
            List            =   "OpeTra_frm_341.frx":0A64
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   120
            Width           =   2775
         End
         Begin EditLib.fpLongInteger ipp_PerAno 
            Height          =   315
            Left            =   1530
            TabIndex        =   9
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
         Begin VB.Label Label3 
            Caption         =   "Año:"
            Height          =   285
            Left            =   180
            TabIndex        =   11
            Top             =   510
            Width           =   1065
         End
         Begin VB.Label Label1 
            Caption         =   "Periodo:"
            Height          =   315
            Left            =   180
            TabIndex        =   10
            Top             =   150
            Width           =   1245
         End
      End
   End
End
Attribute VB_Name = "frm_Rpt_Cofide_03"
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

Private Type g_tpo_PeriodoTC
   PeriodoTC_Col1     As String
   PeriodoTC_Col2     As String
End Type

Private Sub cmd_ExpExc_Click()
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
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT ARCCOF_TIPARC "
   g_str_Parame = g_str_Parame & "   FROM CRE_ARCCOF "
   g_str_Parame = g_str_Parame & "  WHERE ARCCOF_PERMES = " & l_str_PerMes & ""
   g_str_Parame = g_str_Parame & "    AND ARCCOF_PERANO = " & l_str_PerAno & ""
   g_str_Parame = g_str_Parame & "  GROUP BY ARCCOF_TIPARC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
      Exit Sub
   End If
   
   If g_rst_GenAux.BOF And g_rst_GenAux.EOF Then
      MsgBox "No se encontró información para el período seleccionado.", vbExclamation, modgen_g_str_NomPlt
   Else
      g_rst_GenAux.MoveFirst
      Do While Not g_rst_GenAux.EOF
         l_int_numhoja = l_int_numhoja + 1
         g_rst_GenAux.MoveNext
      Loop
      
      Set l_obj_Excel = New Excel.Application
      l_obj_Excel.SheetsInNewWorkbook = l_int_numhoja
      l_obj_Excel.Workbooks.Add
   
      g_rst_GenAux.MoveFirst
      l_int_numhoja = 1
      Do While Not g_rst_GenAux.EOF
         l_str_TipArc = g_rst_GenAux!ARCCOF_TIPARC
         Call fs_GenExc(l_str_PerMes, l_str_PerAno, l_str_TipArc)
         g_rst_GenAux.MoveNext
      Loop
      
      l_obj_Excel.Sheets(1).Select
      l_obj_Excel.Visible = True
      Set l_obj_Excel = Nothing
   End If
   
   g_rst_GenAux.Close
   Set g_rst_GenAux = Nothing
   
   Screen.MousePointer = 0
   cmd_ExpExc.Enabled = True
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

Private Sub fs_GenExc(ByVal p_PerMes As String, ByVal p_PerAno As String, ByVal p_TipArc As String)
Dim r_int_ConVer     As Integer
Dim r_int_FlagTC     As Integer
Dim r_dbl_CapBBP     As Double
Dim r_str_MesBBP     As String
Dim r_arr_Matriz()   As g_tpo_PeriodoTC
Dim r_int_Contad     As Integer


   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT B.HIPMAE_NUMOPE AS OPERACION, "
   g_str_Parame = g_str_Parame & "       TRIM(B.HIPMAE_TDOCLI)||'-'||TRIM(B.HIPMAE_NDOCLI) AS DOC_CLIENTE, "
   g_str_Parame = g_str_Parame & "       TRIM(A.ARCCOF_CODCLI) AS CLIENTE_COFIDE, "
   g_str_Parame = g_str_Parame & "       TRIM(A.ARCCOF_NOMCLI) AS NOMBRE_CLIENTE, "
   g_str_Parame = g_str_Parame & "       DECODE(NVL(C.CUOCIE_SALCAP, 0), 0, 0 , C.CUOCIE_SALCAP) AS MONTO_CRONOG3, "
   g_str_Parame = g_str_Parame & "       E.CUOCIE_SALCAP AS MONTO_CRONOG5, "
   g_str_Parame = g_str_Parame & "       A.ARCCOF_SALDTN AS MONTO_TNC_COFIDE, "
   g_str_Parame = g_str_Parame & "       A.ARCCOF_SALDTC AS MONTO_TC_COFIDE "
   g_str_Parame = g_str_Parame & "  FROM CRE_ARCCOF A "
   g_str_Parame = g_str_Parame & "  LEFT JOIN CRE_HIPMAE B ON B.HIPMAE_CODCOF = A.ARCCOF_CODCLI AND B.HIPMAE_SITUAC IN (2,6,9) "
   g_str_Parame = g_str_Parame & "  LEFT JOIN CRE_CUOCIE C ON C.CUOCIE_PERMES = A.ARCCOF_PERMES AND C.CUOCIE_PERANO = A.ARCCOF_PERANO AND C.CUOCIE_NUMOPE = B.HIPMAE_NUMOPE AND C.CUOCIE_TIPCRO = 3 "
   g_str_Parame = g_str_Parame & "                        AND C.CUOCIE_FECVCT >= " & l_str_PerAno & Format(l_str_PerMes, "00") & "01" & " AND C.CUOCIE_FECVCT <= " & l_str_PerAno & Format(l_str_PerMes, "00") & "31" & " "
   g_str_Parame = g_str_Parame & "  LEFT JOIN CRE_CUOCIE E ON E.CUOCIE_PERMES = A.ARCCOF_PERMES AND E.CUOCIE_PERANO = A.ARCCOF_PERANO AND E.CUOCIE_NUMOPE = B.HIPMAE_NUMOPE AND E.CUOCIE_TIPCRO = 5 "
   g_str_Parame = g_str_Parame & "                        AND E.CUOCIE_FECVCT >= " & l_str_PerAno & Format(l_str_PerMes, "00") & "01" & " AND E.CUOCIE_FECVCT <= " & l_str_PerAno & Format(l_str_PerMes, "00") & "31" & " "
   g_str_Parame = g_str_Parame & " WHERE A.ARCCOF_PERMES = " & p_PerMes & ""
   g_str_Parame = g_str_Parame & "   AND A.ARCCOF_PERANO = " & p_PerAno & ""
   g_str_Parame = g_str_Parame & "   AND A.ARCCOF_CODCLI > 0 "
   g_str_Parame = g_str_Parame & "   AND A.ARCCOF_TIPARC = " & p_TipArc & ""
   g_str_Parame = g_str_Parame & " ORDER BY B.HIPMAE_NUMOPE "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      MsgBox "No se encontró información para el período seleccionado.", vbExclamation, modgen_g_str_NomPlt
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
   
   l_obj_Excel.Sheets(l_int_numhoja).Name = "SALDOS_IFI_10005339_PRD " & Right("000" & p_TipArc, 3)
   With l_obj_Excel.Sheets(l_int_numhoja)
      r_int_ConVer = 1
      .Range(.Cells(r_int_ConVer, 1), .Cells(r_int_ConVer, 1)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(r_int_ConVer, 1), .Cells(r_int_ConVer, 1)).Font.Bold = True
      .Range(.Cells(r_int_ConVer, 1), .Cells(r_int_ConVer, 1)).Font.Underline = xlUnderlineStyleSingle
      .Range("A" & r_int_ConVer & ":K" & r_int_ConVer & "").Merge
      
      r_int_ConVer = 3
      .Cells(r_int_ConVer, 1) = "ITEM"
      .Cells(r_int_ConVer, 2) = "NRO. OPERACION"
      .Cells(r_int_ConVer, 3) = "DOI CLIENTE"
      .Cells(r_int_ConVer, 4) = "COD. CLIENTE COFIDE"
      .Cells(r_int_ConVer, 5) = "NOMBRES"
      .Cells(r_int_ConVer, 6) = "MTO. MICASITA TNC"
      .Cells(r_int_ConVer, 7) = "MTO. COFIDE TNC"
      .Cells(r_int_ConVer, 8) = "DIFERENCIA TNC"
      .Cells(r_int_ConVer, 9) = "MTO. MICASITA TC"
      .Cells(r_int_ConVer, 10) = "MTO. COFIDE TC"
      .Cells(r_int_ConVer, 11) = "DIFERENCIA TC"
      
      .Range(.Cells(r_int_ConVer, 1), .Cells(r_int_ConVer, 11)).Font.Bold = True
      .Range(.Cells(r_int_ConVer, 1), .Cells(r_int_ConVer, 11)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(r_int_ConVer, 1), .Cells(r_int_ConVer, 11)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(r_int_ConVer, 1), .Cells(r_int_ConVer, 11)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(r_int_ConVer, 1), .Cells(r_int_ConVer, 11)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(r_int_ConVer, 1), .Cells(r_int_ConVer, 11)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(r_int_ConVer, 1), .Cells(r_int_ConVer, 11)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(r_int_ConVer, 1), .Cells(r_int_ConVer, 11)).Borders(xlInsideVertical).LineStyle = xlContinuous
      
      .Columns("A").ColumnWidth = 4
      .Columns("B").ColumnWidth = 14
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 12
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 18
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("D").NumberFormat = "@"
      .Columns("E").ColumnWidth = 40
      .Columns("E").HorizontalAlignment = xlHAlignLeft
      .Columns("F").ColumnWidth = 16
      .Columns("F").NumberFormat = "###,###,##0.00"
      .Columns("G").ColumnWidth = 16
      .Columns("G").NumberFormat = "###,###,##0.00"
      .Columns("H").ColumnWidth = 16
      .Columns("H").NumberFormat = "###,###,##0.00"
      .Columns("I").ColumnWidth = 16
      .Columns("I").NumberFormat = "###,###,##0.00"
      .Columns("J").ColumnWidth = 16
      .Columns("J").NumberFormat = "###,###,##0.00"
      .Columns("K").ColumnWidth = 16
      .Columns("K").NumberFormat = "###,###,##0.00"
      
      g_rst_Princi.MoveFirst
      r_int_ConVer = 4
      
      Do While Not g_rst_Princi.EOF
         .Cells(r_int_ConVer, 1) = r_int_ConVer - 3
         
         If Not IsNull(g_rst_Princi!OPERACION) Then
            .Cells(r_int_ConVer, 2) = gf_Formato_NumOpe(g_rst_Princi!OPERACION)
            
            If Mid(g_rst_Princi!OPERACION, 1, 3) = moddat_g_str_AgrCME Then '"003"
               r_int_FlagTC = 0
            ElseIf Mid(g_rst_Princi!OPERACION, 1, 3) = moddat_g_str_AgrMIHG Then '"004"
               r_int_FlagTC = 1
            ElseIf InStr(moddat_g_str_Agr2FMV, Mid(g_rst_Princi!OPERACION, 1, 3)) > 0 Then
'            ElseIf Mid(g_rst_Princi!OPERACION, 1, 3) = "007" Or _
'                   Mid(g_rst_Princi!OPERACION, 1, 3) = "009" Or _
'                   Mid(g_rst_Princi!OPERACION, 1, 3) = "010" Or _
'                   Mid(g_rst_Princi!OPERACION, 1, 3) = "013" Or _
'                   Mid(g_rst_Princi!OPERACION, 1, 3) = "014" Or _
'                   Mid(g_rst_Princi!OPERACION, 1, 3) = "015" Or _
'                   Mid(g_rst_Princi!OPERACION, 1, 3) = "016" Or _
'                   Mid(g_rst_Princi!OPERACION, 1, 3) = "017" Or _
'                   Mid(g_rst_Princi!OPERACION, 1, 3) = "018" Then
               r_int_FlagTC = 1
            ElseIf InStr(moddat_g_str_Agr1FMV, Mid(g_rst_Princi!OPERACION, 1, 3)) > 0 Then
               r_int_FlagTC = 2
'            ElseIf Mid(g_rst_Princi!OPERACION, 1, 3) = "019" Then
'               r_int_FlagTC = 2
'            ElseIf Mid(g_rst_Princi!OPERACION, 1, 3) = "021" Then
'               r_int_FlagTC = 2
'            ElseIf Mid(g_rst_Princi!OPERACION, 1, 3) = "022" Then
'               r_int_FlagTC = 2
'            ElseIf Mid(g_rst_Princi!OPERACION, 1, 3) = "023" Then
'               r_int_FlagTC = 2
'            ElseIf Mid(g_rst_Princi!OPERACION, 1, 3) = "025" Then
'               r_int_FlagTC = 2
            End If
         Else
            .Range(.Cells(r_int_ConVer, 1), .Cells(r_int_ConVer, 11)).Font.Color = vbRed
            .Cells(r_int_ConVer, 2) = "-"
         End If
         
         .Cells(r_int_ConVer, 3) = Trim(g_rst_Princi!DOC_CLIENTE)
         .Cells(r_int_ConVer, 4) = Trim(g_rst_Princi!CLIENTE_COFIDE)
         .Cells(r_int_ConVer, 5) = Trim(Replace(g_rst_Princi!NOMBRE_CLIENTE, "  ", " "))
        
         If Not IsNull(g_rst_Princi!OPERACION) Then
            
            If r_int_FlagTC = 0 Then
               If g_rst_Princi!MONTO_CRONOG5 = "" Then
                  .Cells(r_int_ConVer, 6) = gf_FormatoNumero(0, 12, 2)
               Else
                  If Not IsNull(g_rst_Princi!MONTO_CRONOG5) Then
                     .Cells(r_int_ConVer, 6) = gf_FormatoNumero(g_rst_Princi!MONTO_CRONOG5, 12, 2)
                  Else
                     .Cells(r_int_ConVer, 6) = gf_FormatoNumero(0, 12, 2)
                  End If
               End If
            
            ElseIf r_int_FlagTC = 1 Or r_int_FlagTC = 2 Then
            
               If g_rst_Princi!MONTO_CRONOG3 = 0 Then

                  g_str_Parame = ""
                  g_str_Parame = g_str_Parame & " SELECT HIPCIE_PRECON, HIPCIE_PRENCO "
                  g_str_Parame = g_str_Parame & "   FROM CRE_HIPCIE "
                  g_str_Parame = g_str_Parame & "  WHERE HIPCIE_PERMES  = '" & p_PerMes & "'"
                  g_str_Parame = g_str_Parame & "    AND HIPCIE_PERANO  = '" & p_PerAno & "'"
                  g_str_Parame = g_str_Parame & "    AND HIPCIE_NUMOPE  = '" & g_rst_Princi!OPERACION & "'"

                  If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
                     Exit Sub
                  End If

                  If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
                    If Not IsNull(g_rst_Genera!HIPCIE_PRENCO) Then
                        .Cells(r_int_ConVer, 6) = gf_FormatoNumero(g_rst_Genera!HIPCIE_PRENCO, 12, 2)
                    Else
                        .Cells(r_int_ConVer, 6) = gf_FormatoNumero(0, 12, 2)
                    End If
                  Else
                    .Cells(r_int_ConVer, 6) = gf_FormatoNumero(0, 12, 2)
                  End If
                  g_rst_Genera.Close
                  Set g_rst_Genera = Nothing
                 
               Else
                  .Cells(r_int_ConVer, 6) = gf_FormatoNumero(g_rst_Princi!MONTO_CRONOG3, 12, 2)
               End If

            Else
               .Cells(r_int_ConVer, 6) = gf_FormatoNumero(0, 12, 2)
            End If
         Else
            .Cells(r_int_ConVer, 6) = gf_FormatoNumero(0, 12, 2)
         End If
         
         .Cells(r_int_ConVer, 7) = gf_FormatoNumero(g_rst_Princi!MONTO_TNC_COFIDE, 12, 2)
         .Cells(r_int_ConVer, 8) = .Cells(r_int_ConVer, 6).Value - .Cells(r_int_ConVer, 7).Value
         
         If Not IsNull(g_rst_Princi!OPERACION) Then
            
            If r_int_FlagTC = 1 Then
                  r_str_MesBBP = 0
                  ReDim r_arr_Matriz(0)
                  'MES DE EVALUACIÓN
                  g_str_Parame = ""
                  g_str_Parame = g_str_Parame & " SELECT ROWNUM AS ID, SUBSTR(CUOCIE_FECVCT,5,2) AS MEVA"
                  g_str_Parame = g_str_Parame & "   FROM CRE_CUOCIE"
                  g_str_Parame = g_str_Parame & "  WHERE CUOCIE_PERMES  = '" & p_PerMes & "'"
                  g_str_Parame = g_str_Parame & "    AND CUOCIE_PERANO  = '" & p_PerAno & "'"
                  g_str_Parame = g_str_Parame & "    AND CUOCIE_NUMOPE  = '" & g_rst_Princi!OPERACION & "'"
                  g_str_Parame = g_str_Parame & "    AND CUOCIE_TIPCRO  = 4 "
                  g_str_Parame = g_str_Parame & "    AND ROWNUM <= 2 "
                  g_str_Parame = g_str_Parame & "  ORDER BY ROWNUM"
                  
                  If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
                     Exit Sub
                  End If
                  
                  If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
                     g_rst_Genera.MoveFirst
                     
                     Do Until g_rst_Genera.EOF
                        For r_int_Contad = 1 To 6
                           If g_rst_Genera!ID = 1 Then
                              ReDim Preserve r_arr_Matriz(UBound(r_arr_Matriz) + 1)
                              If r_int_Contad = 1 Then
                                 r_arr_Matriz(UBound(r_arr_Matriz)).PeriodoTC_Col1 = CInt(g_rst_Genera!MEVA) + r_int_Contad
                              Else
                                 r_arr_Matriz(UBound(r_arr_Matriz)).PeriodoTC_Col1 = IIf(CInt(r_arr_Matriz(UBound(r_arr_Matriz) - 1).PeriodoTC_Col1) = 12, 1, CInt(r_arr_Matriz(UBound(r_arr_Matriz) - 1).PeriodoTC_Col1) + 1)
                              End If
                              
                            Else
                              If r_int_Contad = 1 Then
                                 r_arr_Matriz(r_int_Contad).PeriodoTC_Col2 = CInt(g_rst_Genera!MEVA) + r_int_Contad
                              Else
                                 r_arr_Matriz(r_int_Contad).PeriodoTC_Col2 = IIf(CInt(r_arr_Matriz(r_int_Contad - 1).PeriodoTC_Col2) = 12, 1, CInt(r_arr_Matriz(r_int_Contad - 1).PeriodoTC_Col2) + 1)
                              End If
                            End If
                        Next r_int_Contad
                        g_rst_Genera.MoveNext
                        
                     Loop
                  End If
                  g_rst_Genera.Close
                  Set g_rst_Genera = Nothing
                  
                  'BÚSQUEDA EN QUE RANGO SE ENCUENTRA EL PERIODO SELECCIONADO
                  For r_int_Contad = 1 To UBound(r_arr_Matriz)
                     If r_arr_Matriz(r_int_Contad).PeriodoTC_Col1 = l_str_PerMes Then
                        r_str_MesBBP = r_arr_Matriz(UBound(r_arr_Matriz)).PeriodoTC_Col1
                        Exit For
                     ElseIf r_arr_Matriz(r_int_Contad).PeriodoTC_Col2 = l_str_PerMes Then
                        r_str_MesBBP = r_arr_Matriz(UBound(r_arr_Matriz)).PeriodoTC_Col2
                        Exit For
                     End If
                  Next r_int_Contad
                  
                  g_str_Parame = ""
                  g_str_Parame = g_str_Parame & "  SELECT SUM(CUOCIE_CAPBBP) AS CAPBBP "
                  g_str_Parame = g_str_Parame & "    FROM CRE_CUOCIE "
                  g_str_Parame = g_str_Parame & "   WHERE CUOCIE_NUMOPE  = '" & g_rst_Princi!OPERACION & "'"
                  g_str_Parame = g_str_Parame & "     AND CUOCIE_PERMES = '" & p_PerMes & "'"
                  g_str_Parame = g_str_Parame & "     AND CUOCIE_PERANO = '" & p_PerAno & "'"
                  g_str_Parame = g_str_Parame & "     AND CUOCIE_TIPCRO = 1 AND CUOCIE_CUOBBP > 0 "
                  g_str_Parame = g_str_Parame & "     AND CUOCIE_FECVCT > " & l_str_PerAno & Format(l_str_PerMes, "00") & "31" & " AND CUOCIE_FECVCT <= " & l_str_PerAno & Format(r_str_MesBBP, "00") & "31" & " "
                  
                  If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
                     Exit Sub
                  End If
                  If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
                    If Not IsNull(g_rst_Genera!CAPBBP) Then
                        r_dbl_CapBBP = gf_FormatoNumero(g_rst_Genera!CAPBBP, 12, 2)
                    Else
                        r_dbl_CapBBP = gf_FormatoNumero(0, 12, 2)
                    End If
                  Else
                    r_dbl_CapBBP = gf_FormatoNumero(0, 12, 2)
                  End If
                  g_rst_Genera.Close
                  Set g_rst_Genera = Nothing
                  
                  'MONTO MICASITA TC
                  g_str_Parame = ""
                  g_str_Parame = g_str_Parame & " SELECT MAX(CUOCIE_SALCAP + CUOCIE_CAPITA) AS CAPITA_TC "
                  g_str_Parame = g_str_Parame & "   FROM CRE_CUOCIE "
                  g_str_Parame = g_str_Parame & "  WHERE CUOCIE_PERMES  = '" & p_PerMes & "'"
                  g_str_Parame = g_str_Parame & "    AND CUOCIE_PERANO  = '" & p_PerAno & "'"
                  g_str_Parame = g_str_Parame & "    AND CUOCIE_NUMOPE  = '" & g_rst_Princi!OPERACION & "'"
                  g_str_Parame = g_str_Parame & "    AND CUOCIE_TIPCRO  = 4 "
                  g_str_Parame = g_str_Parame & "    AND CUOCIE_FECVCT  >= " & l_str_PerAno & Format(l_str_PerMes, "00") & "01" & " "
                  
                  If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
                     Exit Sub
                  End If
                  
                  If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
                    If Not IsNull(g_rst_Genera!CAPITA_TC) Then
                        .Cells(r_int_ConVer, 9) = gf_FormatoNumero(g_rst_Genera!CAPITA_TC, 12, 2) + r_dbl_CapBBP
                    Else
                        .Cells(r_int_ConVer, 9) = gf_FormatoNumero(0, 12, 2)
                    End If
                  Else
                    .Cells(r_int_ConVer, 9) = gf_FormatoNumero(0, 12, 2)
                  End If
                  g_rst_Genera.Close
                  Set g_rst_Genera = Nothing
            Else
               .Cells(r_int_ConVer, 9) = gf_FormatoNumero(0, 12, 2)
            End If
         Else
           .Cells(r_int_ConVer, 9) = gf_FormatoNumero(0, 12, 2)
         End If
         .Cells(r_int_ConVer, 10) = gf_FormatoNumero(g_rst_Princi!MONTO_TC_COFIDE, 12, 2)
         .Cells(r_int_ConVer, 11) = .Cells(r_int_ConVer, 9).Value - .Cells(r_int_ConVer, 10).Value
            
         'Siguiente registro
         r_int_ConVer = r_int_ConVer + 1
         g_rst_Princi.MoveNext
         DoEvents
      Loop
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      
      .Range(.Cells(1, 1), .Cells(r_int_ConVer, 11)).Font.Name = "Arial"
      .Range(.Cells(1, 1), .Cells(r_int_ConVer, 11)).Font.Size = 8
      
      .Range(.Cells(r_int_ConVer, 6), .Cells(r_int_ConVer, 11)).FormulaR1C1 = "=SUM(R[-" & r_int_ConVer - 2 & "]C:R[-1]C)"
      .Range(.Cells(r_int_ConVer, 6), .Cells(r_int_ConVer, 11)).Font.Bold = True
      .Range(.Cells(r_int_ConVer, 6), .Cells(r_int_ConVer, 11)).HorizontalAlignment = xlVAlignCenter
      .Range(.Cells(r_int_ConVer, 6), .Cells(r_int_ConVer, 11)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(r_int_ConVer, 6), .Cells(r_int_ConVer, 11)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(r_int_ConVer, 6), .Cells(r_int_ConVer, 11)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(r_int_ConVer, 6), .Cells(r_int_ConVer, 11)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(r_int_ConVer, 6), .Cells(r_int_ConVer, 11)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(r_int_ConVer, 6), .Cells(r_int_ConVer, 11)).Borders(xlInsideVertical).LineStyle = xlContinuous
   End With
   
   l_obj_Excel.Sheets(l_int_numhoja).Select
   l_obj_Excel.Sheets(l_int_numhoja).Cells(4, 1).Select
   l_obj_Excel.ActiveWindow.FreezePanes = True
   l_int_numhoja = l_int_numhoja + 1
End Sub

Private Sub ipp_FecIni_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_ExpExc)
   End If
End Sub
