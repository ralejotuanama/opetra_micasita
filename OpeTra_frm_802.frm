VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_Pro_MViPag_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form4"
   ClientHeight    =   2970
   ClientLeft      =   6015
   ClientTop       =   2835
   ClientWidth     =   7110
   Icon            =   "OpeTra_frm_802.frx":0000
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel6 
      Height          =   675
      Left            =   0
      TabIndex        =   6
      Top             =   0
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
         TabIndex        =   7
         Top             =   45
         Width           =   4275
         _Version        =   65536
         _ExtentX        =   7541
         _ExtentY        =   1005
         _StockProps     =   15
         Caption         =   "Reporte de Pagos de Seguros"
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
         Picture         =   "OpeTra_frm_802.frx":000C
         Top             =   90
         Width           =   480
      End
   End
   Begin Threed.SSPanel SSPanel4 
      Height          =   645
      Left            =   0
      TabIndex        =   8
      Top             =   720
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
      Begin VB.CommandButton cmd_Salida 
         Height          =   585
         Left            =   6480
         Picture         =   "OpeTra_frm_802.frx":0316
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Salir de la Opción"
         Top             =   30
         Width           =   585
      End
      Begin VB.CommandButton cmd_ExpExc 
         Height          =   585
         Left            =   30
         Picture         =   "OpeTra_frm_802.frx":0758
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Exportar Información de Seguros"
         Top             =   30
         Width           =   585
      End
   End
   Begin Threed.SSPanel SSPanel3 
      Height          =   1560
      Left            =   0
      TabIndex        =   9
      Top             =   1410
      Width           =   7095
      _Version        =   65536
      _ExtentX        =   12515
      _ExtentY        =   2752
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
      Begin VB.ComboBox cmb_TipRep 
         Height          =   315
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   180
         Width           =   5775
      End
      Begin VB.TextBox txtFactor 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6405
         TabIndex        =   3
         Text            =   "1.03"
         Top             =   990
         Width           =   570
      End
      Begin VB.ComboBox cmb_Permes 
         Height          =   315
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   600
         Width           =   5775
      End
      Begin EditLib.fpDoubleSingle ipp_PerAno 
         Height          =   315
         Left            =   1230
         TabIndex        =   2
         Top             =   990
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
         Caption         =   "Tipo Reporte:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label Label1 
         Caption         =   "Factor:"
         Height          =   255
         Left            =   5625
         TabIndex        =   12
         Top             =   1050
         Width           =   795
      End
      Begin VB.Label Label3 
         Caption         =   "Mes:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   660
         Width           =   795
      End
      Begin VB.Label Label5 
         Caption         =   "Año:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1050
         Width           =   795
      End
   End
End
Attribute VB_Name = "frm_Pro_MViPag_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_dbl_TasaAnual_Cia    As Double
Dim l_dbl_TasaAnual_Cli    As Double
Dim l_dbl_TasaMensual_Cia  As Double
Dim l_dbl_TasaMensual_Cli  As Double

Private Sub cmd_ExpExc_Click()
   If cmb_TipRep.ListIndex = -1 Then
      MsgBox "Debe seleccionar el tipo de reporte.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipRep)
      Exit Sub
   End If
   If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) <> 3 Then
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
   Select Case cmb_TipRep.ItemData(cmb_TipRep.ListIndex)
      Case 1: Call fs_GenExc_SegDes
      Case 2: Call fs_GenExc_SegImb
      Case 3: Call fs_GenExc_Endoso
   End Select
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Salida_Click()
    Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call gs_CentraForm(Me)
   
   Call gs_SetFocus(cmb_TipRep)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   Call moddat_gs_Carga_LisIte_Combo(cmb_PerMes, 1, "033")
   ipp_PerAno.Text = Year(date)
   l_dbl_TasaAnual_Cia = Format(0.1589 / 100, "##0.000000")
   l_dbl_TasaAnual_Cli = Format(0.227102 / 100, "##0.000000")
   l_dbl_TasaMensual_Cia = Format((1 + l_dbl_TasaAnual_Cia) ^ (1 / 12) - 1, "##0.000000")
   l_dbl_TasaMensual_Cli = Format((1 + l_dbl_TasaAnual_Cli) ^ (1 / 12) - 1, "##0.000000")
   
   cmb_TipRep.Clear
   cmb_TipRep.AddItem "REPORTE - SEGUROS DE DESGRAVAMEN"
   cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = 1
   cmb_TipRep.AddItem "REPORTE - SEGUROS DEL INMUEBLE"
   cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = 2
   cmb_TipRep.AddItem "REPORTE - ENDOSADOS"
   cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = 3
   cmb_TipRep.ListIndex = -1
End Sub

Private Sub fs_GenExc_SegDes()
Dim r_obj_Excel      As Excel.Application
Dim r_int_Contad     As Integer
Dim r_int_ConVer     As Integer
Dim r_dbl_PBPPer     As Double
Dim r_str_Fecha      As String
Dim r_str_FecIni     As String
Dim r_str_FecFin     As String
Dim r_int_Index      As Integer
Dim r_dbl_Monto      As Double
Dim r_dbl_MtoInd     As Double
Dim r_dbl_MtoMan     As Double
Dim r_int_ItmReg     As Integer
Dim r_str_PerMes     As Integer
Dim r_str_PerAno     As Integer
Dim r_int_ConTem     As Integer
Dim r_int_ConAux     As Integer

   r_str_PerMes = CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   r_str_PerAno = CInt(ipp_PerAno.Text)
   r_str_FecIni = Format(ipp_PerAno.Text, "0000") & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") & "25"
   If cmb_PerMes.ItemData(cmb_PerMes.ListIndex) = 12 Then
      r_str_FecFin = Format(ipp_PerAno.Text + 1, "0000") & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex) + 1, "00") & "02"
   Else
      r_str_FecFin = Format(ipp_PerAno.Text, "0000") & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex) + 1, "00") & "02"
   End If
   r_str_Fecha = "01/" & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") & "/" & Format(ipp_PerAno.Text, "0000")
   
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 2
   r_obj_Excel.Workbooks.Add
   r_obj_Excel.Sheets(1).Name = "DETALLE"
   
   r_dbl_Monto = 0
   With r_obj_Excel.ActiveSheet
      .Cells(1, 1) = ObtieneNomMes(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)) & " - " & Format(ipp_PerAno.Text, "0000")
      .Range(.Cells(1, 1), .Cells(1, 26)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(1, 1), .Cells(1, 26)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 26)).Merge
      .Range(.Cells(1, 1), .Cells(1, 26)).Font.Size = 18
      .Cells(3, 1) = "MONEDA: SOLES"
      .Range(.Cells(3, 1), .Cells(3, 3)).HorizontalAlignment = xlHAlignLeft
      .Range(.Cells(3, 1), .Cells(3, 3)).Font.Bold = True
      .Range(.Cells(3, 1), .Cells(3, 3)).Merge
      
      For r_int_Contad = 1 To 26 Step 1
         .Range(.Cells(5, 1), .Cells(5, r_int_Contad)).Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range(.Cells(5, 1), .Cells(5, r_int_Contad)).Borders(xlEdgeTop).LineStyle = xlContinuous
         .Range(.Cells(5, 1), .Cells(5, r_int_Contad)).Borders(xlEdgeBottom).LineStyle = xlContinuous
         .Range(.Cells(5, 1), .Cells(5, r_int_Contad)).Borders(xlEdgeRight).LineStyle = xlContinuous
         .Range(.Cells(5, 1), .Cells(5, r_int_Contad)).Borders(xlInsideVertical).LineStyle = xlContinuous
      Next
      
      .Cells(5, 1) = "ITEM"
      .Cells(5, 2) = "NRO. PRESTAMO"
      .Cells(5, 3) = "CODIGO"
      .Cells(5, 4) = "TIPO PERSONA"
      .Cells(5, 5) = "TIPO DOCUMENTO"
      .Cells(5, 6) = "DOC. IDENTIDAD"
      .Cells(5, 7) = "APE_PATERNO"
      .Cells(5, 8) = "APE_MATERNO"
      .Cells(5, 9) = "NOM_CLIENTE"
      .Cells(5, 10) = "F. NACIMIENTO"
      .Cells(5, 11) = "SEXO"
      .Cells(5, 12) = "CORREO"
      .Cells(5, 13) = "TELEF_FIJO"
      .Cells(5, 14) = "TELEF_CELULAR"
      .Cells(5, 15) = "EMP. SEGUROS"
      .Cells(5, 16) = "F. APERTURA"
      .Cells(5, 17) = "DURACION"
      .Cells(5, 18) = "F. PAGO"
      .Cells(5, 19) = "IMP. PRESTAMO"
      .Cells(5, 20) = "SALDO PRESTAMO"
      .Cells(5, 21) = "TIPOSEGURO"
      .Cells(5, 22) = "COBERTURA"
      .Cells(5, 23) = "FACTOR APLICACION"
      .Cells(5, 24) = "% RECARGO"
      .Cells(5, 25) = "MONTO"
      .Cells(5, 26) = "MONEDA"
      
      .Range(.Cells(5, 1), .Cells(5, 26)).Font.Bold = True
      .Range(.Cells(5, 1), .Cells(5, 26)).HorizontalAlignment = xlHAlignCenter
      
      .Columns("A").ColumnWidth = 6
      .Columns("A").HorizontalAlignment = xlHAlignCenter
      .Columns("B").ColumnWidth = 16
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 16
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 16
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").ColumnWidth = 16
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").ColumnWidth = 16
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("G").ColumnWidth = 25
      .Columns("G").HorizontalAlignment = xlHAlignLeft
      .Columns("H").ColumnWidth = 25
      .Columns("H").HorizontalAlignment = xlHAlignLeft
      .Columns("I").ColumnWidth = 30
      .Columns("I").HorizontalAlignment = xlHAlignLeft
      .Columns("J").ColumnWidth = 16
      .Columns("J").HorizontalAlignment = xlHAlignCenter
      .Columns("K").ColumnWidth = 13
      .Columns("K").HorizontalAlignment = xlHAlignCenter
      .Columns("L").ColumnWidth = 45
      .Columns("L").HorizontalAlignment = xlHAlignLeft
      .Columns("M").ColumnWidth = 16
      .Columns("M").HorizontalAlignment = xlHAlignCenter
      .Columns("N").ColumnWidth = 16
      .Columns("N").HorizontalAlignment = xlHAlignCenter
      .Columns("O").ColumnWidth = 50
      .Columns("O").HorizontalAlignment = xlHAlignCenter
      .Columns("P").ColumnWidth = 20
      .Columns("P").HorizontalAlignment = xlHAlignCenter
      .Columns("Q").ColumnWidth = 13
      .Columns("Q").HorizontalAlignment = xlHAlignRight
      .Columns("R").ColumnWidth = 13
      .Columns("R").HorizontalAlignment = xlHAlignCenter
      .Columns("S").ColumnWidth = 18
      .Columns("S").NumberFormat = "###,###,##0.00"
      .Columns("S").HorizontalAlignment = xlHAlignRight
      .Columns("T").ColumnWidth = 18
      .Columns("T").NumberFormat = "###,###,##0.00"
      .Columns("T").HorizontalAlignment = xlHAlignRight
      .Columns("U").ColumnWidth = 15
      .Columns("U").HorizontalAlignment = xlHAlignCenter
      .Columns("V").ColumnWidth = 25
      .Columns("V").HorizontalAlignment = xlHAlignCenter
      .Columns("W").ColumnWidth = 19
      .Columns("W").NumberFormat = "###,##0.000000"
      .Columns("W").HorizontalAlignment = xlHAlignCenter
      .Columns("X").ColumnWidth = 16
      .Columns("X").HorizontalAlignment = xlHAlignCenter
      .Columns("Y").ColumnWidth = 16
      .Columns("Y").NumberFormat = "###,###,##0.00"
      .Columns("Z").ColumnWidth = 22
      .Columns("Z").HorizontalAlignment = xlHAlignCenter
   End With
   
   r_int_ConVer = 6
   For r_int_Index = 1 To 2 Step 1
      If r_int_Index = 2 Then
         r_dbl_Monto = 9
         r_int_ConVer = r_int_ConVer + 4
         
         r_obj_Excel.Range(r_obj_Excel.Cells(5, 1), r_obj_Excel.Cells(5, 26)).HorizontalAlignment = xlHAlignCenter
               
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = "MONEDA: DOLARES"
         r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1), r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3)).HorizontalAlignment = xlHAlignLeft
         r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1), r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3)).Font.Bold = True
         r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1), r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3)).Merge
         r_int_ConVer = r_int_ConVer + 2
         
         For r_int_Contad = 1 To 26 Step 1
            r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1), r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, r_int_Contad)).Borders(xlEdgeLeft).LineStyle = xlContinuous
            r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1), r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, r_int_Contad)).Borders(xlEdgeTop).LineStyle = xlContinuous
            r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1), r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, r_int_Contad)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1), r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, r_int_Contad)).Borders(xlEdgeRight).LineStyle = xlContinuous
            r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1), r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, r_int_Contad)).Borders(xlInsideVertical).LineStyle = xlContinuous
         Next
         
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = "ITEM"
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = "NRO. PRESTAMO"
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = "CODIGO"
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = "TIPO PERSONA"
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = "TIPO DOCUMENTO"
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = "DOC. IDENTIDAD"
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = "APE_PATERNO"
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = "APE_MATERNO"
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = "NOM_CLIENTE"
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = "F. NACIMIENTO"
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = "SEXO"
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = "CORREO"
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = "TELEF_FIJO"
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14) = "TELEF_CELULAR"
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = "EMP. SEGUROS"
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 16) = "F. APERTURA"
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 17) = "DURACION"
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 18) = "F. PAGO"
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 19) = "IMP. PRESTAMO"
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 20) = "SALDO PRESTAMO"
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 21) = "TIPOSEGURO"
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 22) = "COBERTURA"
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 23) = "FACTOR APLICACION"
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 24) = "% RECARGO"
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 25) = "MONTO"
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 26) = "MONEDA"
             
         r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1), r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 26)).Font.Bold = True
         r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1), r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 26)).HorizontalAlignment = xlHAlignCenter
         r_int_ConVer = r_int_ConVer + 1
      End If
      
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT HIPCIE_NUMOPE AS NROPRESTAMO, "
      g_str_Parame = g_str_Parame & "       TRIM(HIPCIE_TDOCLI)||'-'||TRIM(HIPCIE_NDOCLI) AS CODIGO, "
      g_str_Parame = g_str_Parame & "       (CASE WHEN TRIM(HIPCIE_TDOCLI) = 1 OR TRIM(HIPCIE_TDOCLI) = 4 OR TRIM(HIPCIE_TDOCLI) = 7 THEN 'N' "
      g_str_Parame = g_str_Parame & "        ELSE CASE WHEN TRIM(HIPCIE_TDOCLI) = 6 THEN 'J' END "
      g_str_Parame = g_str_Parame & "         END) AS TIPO_PERSONA,"
      g_str_Parame = g_str_Parame & "       TRIM(HIPCIE_TDOCLI) AS TIPO_DOCUMENTO, "
      g_str_Parame = g_str_Parame & "       TRIM(HIPCIE_NDOCLI) AS DOCIDENTIDAD, "
      g_str_Parame = g_str_Parame & "       TRIM(DATGEN_APEPAT) AS APELLIDO_PATERNO, "
      g_str_Parame = g_str_Parame & "       TRIM(DATGEN_APEMAT) AS APELLIDO_MATERNO, "
      g_str_Parame = g_str_Parame & "       TRIM(DATGEN_NOMBRE) AS NOM_CLIENTE, "
      g_str_Parame = g_str_Parame & "       SUBSTR(DATGEN_NACFEC,7,2)||'/'||SUBSTR(DATGEN_NACFEC,5,2)||'/'||SUBSTR(DATGEN_NACFEC,1,4) AS FECHANAC, "
      g_str_Parame = g_str_Parame & "       DECODE(DATGEN_CODSEX, 1, 'MASCULINO', 'FEMENINO')  AS SEXO, "
      g_str_Parame = g_str_Parame & "       TRIM(DATGEN_DIRELE)                 AS CORREO,"
      g_str_Parame = g_str_Parame & "       TRIM(DATGEN_TELEFO)                 AS TELEF_FIJO,"
      g_str_Parame = g_str_Parame & "       TRIM(DATGEN_NUMCEL)                 AS TELEF_CELULAR,"
      g_str_Parame = g_str_Parame & "       TRIM(SEGEMP_RAZSOC) AS EMPSEGUROS, "
      g_str_Parame = g_str_Parame & "       SUBSTR(POLIZA_FEMDES,7,2)||'/'||SUBSTR(POLIZA_FEMDES,5,2)||'/'||SUBSTR(POLIZA_FEMDES,1,4) AS FECHAAPERTURA, "
      g_str_Parame = g_str_Parame & "       HIPMAE_PLAANO AS DURACION, "
      g_str_Parame = g_str_Parame & "       TO_CHAR(TO_DATE('" & r_str_Fecha & "','DD/MM/YYYY'), 'MONTH', 'NLS_DATE_LANGUAGE=SPANISH') AS FECPAGO, "
      g_str_Parame = g_str_Parame & "       HIPCIE_MTOPRE AS IMPPRESTAMO, "
      g_str_Parame = g_str_Parame & "       HIPCIE_SALCAP+HIPCIE_SALCON AS SALDOPRESTAMO, "
      g_str_Parame = g_str_Parame & "       EVASEG_TIPSEG AS TIPOSEGURO, "
      g_str_Parame = g_str_Parame & "       TRIM(SEGTIP_DESCRI) AS COBERTURA, "
      g_str_Parame = g_str_Parame & "       ROUND(HIPCIE_FOIPRE, 5) AS FACTORAPLICA, "
      g_str_Parame = g_str_Parame & "       ROUND(((HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_FOIPRE)/100,2) AS MONTO, "
      g_str_Parame = g_str_Parame & "       '0%' AS RECARGO, "
      g_str_Parame = g_str_Parame & "       TRIM(PARDES_DESCRI) AS MONEDA "
      g_str_Parame = g_str_Parame & "  FROM CRE_HIPCIE "
      g_str_Parame = g_str_Parame & " INNER JOIN CLI_DATGEN ON DATGEN_TIPDOC = HIPCIE_TDOCLI AND DATGEN_NUMDOC = HIPCIE_NDOCLI "
      g_str_Parame = g_str_Parame & " INNER JOIN CRE_HIPMAE ON HIPMAE_NUMOPE = HIPCIE_NUMOPE "
      g_str_Parame = g_str_Parame & " INNER JOIN TRA_POLIZA ON POLIZA_NUMSOL = HIPMAE_NUMSOL "
      g_str_Parame = g_str_Parame & " INNER JOIN TRA_EVASEG ON EVASEG_NUMSOL = HIPMAE_NUMSOL AND EVASEG_TIPSEG <> 13"
      g_str_Parame = g_str_Parame & " INNER JOIN MNT_SEGEMP ON SEGEMP_CODIGO = EVASEG_ESGDES "
      g_str_Parame = g_str_Parame & " INNER JOIN MNT_SEGTIP ON SEGTIP_CODIGO = EVASEG_ESGDES AND SEGTIP_TIPSEG = EVASEG_TIPSEG "
      g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES ON PARDES_CODGRP = 204 AND PARDES_CODITE = HIPCIE_TIPMON"
      g_str_Parame = g_str_Parame & " WHERE HIPCIE_PERMES = " & cmb_PerMes.ItemData(cmb_PerMes.ListIndex) & " "
      g_str_Parame = g_str_Parame & "   AND HIPCIE_PERANO = " & Format(ipp_PerAno.Text, "0000") & " "
      g_str_Parame = g_str_Parame & "   AND HIPCIE_TIPMON = " & r_int_Index & " "
      g_str_Parame = g_str_Parame & "UNION "
      g_str_Parame = g_str_Parame & "SELECT HIPMAE_NUMOPE AS NROPRESTAMO, "
      g_str_Parame = g_str_Parame & "       TRIM(HIPMAE_TDOCLI)||'-'||TRIM(HIPMAE_NDOCLI) AS CODIGO, "
      g_str_Parame = g_str_Parame & "       (CASE WHEN TRIM(HIPMAE_TDOCLI) = 1 OR TRIM(HIPMAE_TDOCLI) = 4 OR TRIM(HIPMAE_TDOCLI) = 7 THEN 'N' "
      g_str_Parame = g_str_Parame & "        ELSE CASE WHEN TRIM(HIPMAE_TDOCLI) = 6 THEN 'J' END "
      g_str_Parame = g_str_Parame & "         END) AS TIPO_PERSONA,"
      g_str_Parame = g_str_Parame & "       TRIM(HIPMAE_TDOCLI) AS TIPO_DOCUMENTO, "
      g_str_Parame = g_str_Parame & "       TRIM(HIPMAE_NDOCLI) AS DOCIDENTIDAD, "
      g_str_Parame = g_str_Parame & "       TRIM(DATGEN_APEPAT) AS APELLIDO_PATERNO, "
      g_str_Parame = g_str_Parame & "       TRIM(DATGEN_APEMAT) AS APELLIDO_MATERNO, "
      g_str_Parame = g_str_Parame & "       TRIM(DATGEN_NOMBRE) AS NOM_CLIENTE, "
      g_str_Parame = g_str_Parame & "       SUBSTR(DATGEN_NACFEC,7,2)||'/'||SUBSTR(DATGEN_NACFEC,5,2)||'/'||SUBSTR(DATGEN_NACFEC,1,4) AS FECHANAC, "
      g_str_Parame = g_str_Parame & "       DECODE(DATGEN_CODSEX, 1, 'MASCULINO', 'FEMENINO')  AS SEXO, "
      g_str_Parame = g_str_Parame & "       TRIM(DATGEN_DIRELE)                 AS CORREO, "
      g_str_Parame = g_str_Parame & "       TRIM(DATGEN_TELEFO)                 AS TELEF_FIJO, "
      g_str_Parame = g_str_Parame & "       TRIM(DATGEN_NUMCEL)                 AS TELEF_CELULAR, "
      g_str_Parame = g_str_Parame & "       TRIM(SEGEMP_RAZSOC) AS EMPSEGUROS, "
      g_str_Parame = g_str_Parame & "       SUBSTR(POLIZA_FEMDES,7,2)||'/'||SUBSTR(POLIZA_FEMDES,5,2)||'/'||SUBSTR(POLIZA_FEMDES,1,4) AS FECHAAPERTURA, "
      g_str_Parame = g_str_Parame & "       HIPMAE_PLAANO AS DURACION, "
      g_str_Parame = g_str_Parame & "       TO_CHAR(TO_DATE('" & r_str_Fecha & "','DD/MM/YYYY'), 'MONTH', 'NLS_DATE_LANGUAGE=SPANISH') AS FECPAGO, "
      g_str_Parame = g_str_Parame & "       HIPMAE_MTOPRE AS IMPPRESTAMO, "
      g_str_Parame = g_str_Parame & "       HIPMAE_SALCAP+HIPMAE_SALCON AS SALDOPRESTAMO, "
      g_str_Parame = g_str_Parame & "       EVASEG_TIPSEG AS TIPOSEGURO, "
      g_str_Parame = g_str_Parame & "       TRIM(SEGTIP_DESCRI) AS COBERTURA, "
      g_str_Parame = g_str_Parame & "       ROUND(HIPMAE_FOIPRE, 5) AS FACTORAPLICA, "
      g_str_Parame = g_str_Parame & "       ROUND(((HIPMAE_SALCAP+HIPMAE_SALCON)*HIPMAE_FOIPRE)/100,2) AS MONTO, "
      g_str_Parame = g_str_Parame & "       '0%' AS RECARGO, "
      g_str_Parame = g_str_Parame & "       TRIM(PARDES_DESCRI) AS MONEDA "
      g_str_Parame = g_str_Parame & "  FROM CRE_HIPMAE "
      g_str_Parame = g_str_Parame & " INNER JOIN CLI_DATGEN ON DATGEN_TIPDOC = HIPMAE_TDOCLI AND DATGEN_NUMDOC = HIPMAE_NDOCLI "
      g_str_Parame = g_str_Parame & " INNER JOIN TRA_POLIZA ON POLIZA_NUMSOL = HIPMAE_NUMSOL "
      g_str_Parame = g_str_Parame & " INNER JOIN TRA_EVASEG ON EVASEG_NUMSOL = HIPMAE_NUMSOL AND EVASEG_TIPSEG <> 13"
      g_str_Parame = g_str_Parame & " INNER JOIN MNT_SEGEMP ON SEGEMP_CODIGO = EVASEG_ESGDES "
      g_str_Parame = g_str_Parame & " INNER JOIN MNT_SEGTIP ON SEGTIP_CODIGO = EVASEG_ESGDES AND SEGTIP_TIPSEG = EVASEG_TIPSEG "
      g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES ON PARDES_CODGRP = 204 AND PARDES_CODITE = HIPMAE_MONEDA "
      g_str_Parame = g_str_Parame & " WHERE HIPMAE_SITUAC = 6 "
      g_str_Parame = g_str_Parame & "   AND HIPMAE_MONEDA = " & r_int_Index & " "
      g_str_Parame = g_str_Parame & "   AND HIPMAE_FECCAN >= " & r_str_FecIni
      g_str_Parame = g_str_Parame & "   AND HIPMAE_FECCAN <= " & r_str_FecFin
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      If g_rst_Princi.BOF And g_rst_Princi.EOF Then
         MsgBox "No se encontraron registrios.", vbInformation, "Mensaje"
         g_rst_Princi.Close
         Set g_rst_Princi = Nothing
         Exit Sub
      End If
      
      r_dbl_MtoInd = 0
      r_dbl_MtoMan = 0
      r_int_ItmReg = 1
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ItmReg
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = gf_Formato_NumOpe(g_rst_Princi!NROPRESTAMO)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = Trim(g_rst_Princi!Codigo)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = Trim(g_rst_Princi!TIPO_PERSONA)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = "'" & CStr(g_rst_Princi!TIPO_DOCUMENTO)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = "'" & CStr(g_rst_Princi!DOCIDENTIDAD)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = "'" & CStr(g_rst_Princi!APELLIDO_PATERNO)
         If Not IsNull(g_rst_Princi!APELLIDO_MATERNO) Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = "'" & CStr(g_rst_Princi!APELLIDO_MATERNO)
         End If
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = "'" & CStr(g_rst_Princi!NOM_CLIENTE)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = "'" & CStr(g_rst_Princi!FECHANAC)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = "'" & CStr(g_rst_Princi!SEXO)
         If Not IsNull(g_rst_Princi!CORREO) Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = "'" & CStr(g_rst_Princi!CORREO)
         End If
         If Not IsNull(g_rst_Princi!TELEF_FIJO) Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = "'" & CStr(g_rst_Princi!TELEF_FIJO)
         End If
         If Not IsNull(g_rst_Princi!TELEF_CELULAR) Then
            r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14) = "'" & CStr(g_rst_Princi!TELEF_CELULAR)
         End If
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = Trim(g_rst_Princi!EMPSEGUROS)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 16) = CStr(g_rst_Princi!FECHAAPERTURA)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 17) = Trim(g_rst_Princi!DURACION)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 18) = "'" & Trim(g_rst_Princi!FECPAGO)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 19) = Format(g_rst_Princi!IMPPRESTAMO, "###,###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 20) = Format(g_rst_Princi!SALDOPRESTAMO, "###,###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 21) = Trim(g_rst_Princi!TIPOSEGURO)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 22) = Trim(g_rst_Princi!COBERTURA)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 23) = Format(g_rst_Princi!FACTORAPLICA, "###,###,###,##0.0000")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 24) = g_rst_Princi!RECARGO
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 25) = Format(g_rst_Princi!MONTO, "###,###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 26) = Trim(g_rst_Princi!Moneda)
         
         r_int_ConVer = r_int_ConVer + 1
         r_int_ItmReg = r_int_ItmReg + 1
         g_rst_Princi.MoveNext
         DoEvents
      Loop
      
      If r_int_Index = 1 Then
         r_int_ConAux = r_int_ConVer
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 25).FormulaR1C1 = "=SUM(R[-" & r_int_ConVer - 5 & "]C:R[-1]C)"
      Else
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 25).FormulaR1C1 = "=SUM(R[-" & r_int_ConVer - r_int_ConAux - 5 & "]C:R[-1]C)"
      End If
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 25).Font.Bold = True
            
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   Next
   
   r_obj_Excel.Sheets(2).Name = "RESUMEN"
   With r_obj_Excel.Sheets(2)
      .Range(.Cells(4, 1), .Cells(4, 25)).Font.Name = "Calibri"
      .Range(.Cells(4, 1), .Cells(4, 25)).Font.Size = 9
      .Range(.Cells(4, 1), .Cells(4, 25)).Font.Bold = True
      
      .Range(.Cells(4, 2), .Cells(4, 6)).Merge
      .Range(.Cells(4, 2), .Cells(4, 6)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(4, 2), .Cells(4, 6)).Font.Bold = True
      
      .Range(.Cells(4, 8), .Cells(4, 14)).Merge
      .Range(.Cells(4, 8), .Cells(4, 14)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(4, 8), .Cells(4, 14)).Font.Bold = True
      
      .Range(.Cells(4, 1), .Cells(4, 1)).Merge
      .Range(.Cells(4, 1), .Cells(4, 1)).WrapText = True
      .Range(.Cells(4, 1), .Cells(4, 1)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(4, 1), .Cells(4, 1)).ColumnWidth = 17
      
      .Range(.Cells(4, 2), .Cells(4, 2)).Merge
      .Range(.Cells(4, 2), .Cells(4, 2)).WrapText = True
      .Range(.Cells(4, 2), .Cells(4, 2)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(4, 2), .Cells(4, 2)).ColumnWidth = 11
      
      .Range(.Cells(4, 3), .Cells(4, 3)).Merge
      .Range(.Cells(4, 3), .Cells(4, 3)).WrapText = True
      .Range(.Cells(4, 3), .Cells(4, 3)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(4, 3), .Cells(4, 3)).ColumnWidth = 11
      
      .Range(.Cells(4, 4), .Cells(4, 4)).Merge
      .Range(.Cells(4, 4), .Cells(4, 4)).WrapText = True
      .Range(.Cells(4, 4), .Cells(4, 4)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(4, 4), .Cells(4, 4)).ColumnWidth = 11
      
      .Range(.Cells(4, 5), .Cells(4, 5)).Merge
      .Range(.Cells(4, 5), .Cells(4, 5)).WrapText = True
      .Range(.Cells(4, 5), .Cells(4, 5)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(4, 5), .Cells(4, 5)).ColumnWidth = 11
      
      .Range(.Cells(4, 6), .Cells(4, 6)).Merge
      .Range(.Cells(4, 6), .Cells(4, 6)).WrapText = True
      .Range(.Cells(4, 6), .Cells(4, 6)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(4, 6), .Cells(4, 6)).ColumnWidth = 11
      
      .Range(.Cells(4, 7), .Cells(4, 7)).ColumnWidth = 4
      
      .Range(.Cells(4, 8), .Cells(4, 8)).Merge
      .Range(.Cells(4, 8), .Cells(4, 8)).WrapText = True
      .Range(.Cells(4, 8), .Cells(4, 8)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(4, 8), .Cells(4, 8)).ColumnWidth = 11
      
      .Range(.Cells(4, 9), .Cells(4, 9)).Merge
      .Range(.Cells(4, 9), .Cells(4, 9)).WrapText = True
      .Range(.Cells(4, 9), .Cells(4, 9)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(4, 9), .Cells(4, 9)).ColumnWidth = 11
      
      .Range(.Cells(4, 10), .Cells(4, 10)).Merge
      .Range(.Cells(4, 10), .Cells(4, 10)).WrapText = True
      .Range(.Cells(4, 10), .Cells(4, 10)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(4, 10), .Cells(4, 10)).ColumnWidth = 11
      
      .Range(.Cells(4, 11), .Cells(4, 11)).Merge
      .Range(.Cells(4, 11), .Cells(4, 11)).WrapText = True
      .Range(.Cells(4, 11), .Cells(4, 11)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(4, 11), .Cells(4, 11)).ColumnWidth = 11
      
      .Range(.Cells(4, 12), .Cells(4, 12)).Merge
      .Range(.Cells(4, 12), .Cells(4, 12)).WrapText = True
      .Range(.Cells(4, 12), .Cells(4, 12)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(4, 12), .Cells(4, 12)).ColumnWidth = 11
      
      .Range(.Cells(4, 13), .Cells(4, 13)).Merge
      .Range(.Cells(4, 13), .Cells(4, 13)).WrapText = True
      .Range(.Cells(4, 13), .Cells(4, 13)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(4, 13), .Cells(4, 13)).ColumnWidth = 11
      
      .Range(.Cells(4, 14), .Cells(4, 14)).Merge
      .Range(.Cells(4, 14), .Cells(4, 14)).WrapText = True
      .Range(.Cells(4, 14), .Cells(4, 14)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(4, 14), .Cells(4, 14)).ColumnWidth = 11
      
      .Range(.Cells(4, 15), .Cells(4, 15)).Merge
      .Range(.Cells(4, 15), .Cells(4, 15)).WrapText = True
      .Range(.Cells(4, 15), .Cells(4, 15)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(4, 15), .Cells(4, 15)).ColumnWidth = 11
      
      .Range(.Cells(2, 1), .Cells(2, 2)).Font.Name = "Calibri"
      .Range(.Cells(2, 1), .Cells(2, 2)).Font.Size = 9
      .Range(.Cells(2, 1), .Cells(2, 2)).Font.Bold = True
      
      .Cells(2, 1) = "TIPO DE MONEDA"
      .Cells(2, 2) = "SOLES"
      .Range(.Cells(2, 1), .Cells(2, 1)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(2, 2), .Cells(2, 2)).HorizontalAlignment = xlHAlignCenter
      
      .Cells(4, 1) = "DESCRIPCION"
      .Cells(4, 2) = "IND"
      .Cells(4, 8) = "MAN"
      .Cells(4, 15) = "TOTAL"
      
      .Range(.Cells(4, 1), .Cells(4, 15)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(4, 1), .Cells(4, 15)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 1), .Cells(4, 15)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(4, 1), .Cells(4, 15)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(4, 1), .Cells(4, 15)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(4, 1), .Cells(4, 15)).Borders(xlInsideVertical).LineStyle = xlContinuous
      
      g_str_Parame = ""
      g_str_Parame = "USP_RPT_SEGDESG ("
      g_str_Parame = g_str_Parame & "" & r_str_PerMes & ", "
      g_str_Parame = g_str_Parame & "" & r_str_PerAno & ","
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "',"
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "',"
      g_str_Parame = g_str_Parame & "'REPORTE DE DESGRAVAMEN',"
      g_str_Parame = g_str_Parame & "'1')"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         MsgBox "Error al ejecutar procedimiento de Desgravamen.", vbCritical, modgen_g_str_NomPlt
         Exit Sub
      End If
   
      'Soles
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT * FROM RPT_TABLA_TEMP "
      g_str_Parame = g_str_Parame + " WHERE RPT_PERMES = '" & r_str_PerMes & "' "
      g_str_Parame = g_str_Parame + "   AND RPT_PERANO = '" & r_str_PerAno & "' "
      g_str_Parame = g_str_Parame + "   AND RPT_TERCRE = '" & modgen_g_str_NombPC & "' "
      g_str_Parame = g_str_Parame + "   AND RPT_USUCRE = '" & modgen_g_str_CodUsu & "' "
      g_str_Parame = g_str_Parame + "   AND TRIM(RPT_NOMBRE) = 'REPORTE DE DESGRAVAMEN' "
      g_str_Parame = g_str_Parame + "   AND RPT_MONEDA = '1' "
      g_str_Parame = g_str_Parame + " ORDER BY TO_NUMBER(RPT_CODIGO)"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      If g_rst_Princi.BOF And g_rst_Princi.EOF Then
         g_rst_Princi.Close
         Set g_rst_Princi = Nothing
         MsgBox "No se encontraron registros.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
      
      r_int_ConTem = 5
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         .Range(.Cells(r_int_ConTem, 1), .Cells(r_int_ConTem, 25)).Font.Name = "Calibri"
         .Range(.Cells(r_int_ConTem, 1), .Cells(r_int_ConTem, 25)).Font.Size = 9
         
         If Val(g_rst_Princi!RPT_CODIGO) >= "3" Then
            .Range(.Cells(r_int_ConTem, 1), .Cells(r_int_ConTem, 1)).NumberFormat = "###,##0.00"
            .Range(.Cells(r_int_ConTem, 2), .Cells(r_int_ConTem, 2)).NumberFormat = "###,##0.00"
            .Range(.Cells(r_int_ConTem, 3), .Cells(r_int_ConTem, 3)).NumberFormat = "###,##0.00"
            .Range(.Cells(r_int_ConTem, 4), .Cells(r_int_ConTem, 4)).NumberFormat = "###,##0.00"
            .Range(.Cells(r_int_ConTem, 5), .Cells(r_int_ConTem, 5)).NumberFormat = "###,##0.00"
            .Range(.Cells(r_int_ConTem, 6), .Cells(r_int_ConTem, 6)).NumberFormat = "###,##0.00"
            .Range(.Cells(r_int_ConTem, 8), .Cells(r_int_ConTem, 8)).NumberFormat = "###,##0.00"
            .Range(.Cells(r_int_ConTem, 9), .Cells(r_int_ConTem, 9)).NumberFormat = "###,##0.00"
            .Range(.Cells(r_int_ConTem, 10), .Cells(r_int_ConTem, 10)).NumberFormat = "###,##0.00"
            .Range(.Cells(r_int_ConTem, 11), .Cells(r_int_ConTem, 11)).NumberFormat = "###,##0.00"
            .Range(.Cells(r_int_ConTem, 12), .Cells(r_int_ConTem, 12)).NumberFormat = "###,##0.00"
            .Range(.Cells(r_int_ConTem, 13), .Cells(r_int_ConTem, 13)).NumberFormat = "###,##0.00"
            .Range(.Cells(r_int_ConTem, 14), .Cells(r_int_ConTem, 14)).NumberFormat = "###,##0.00"
            .Range(.Cells(r_int_ConTem, 15), .Cells(r_int_ConTem, 15)).NumberFormat = "###,##0.00"
         End If
         
         If Val(g_rst_Princi!RPT_CODIGO) = "3" Then
            .Range(.Cells(r_int_ConTem, 1), .Cells(r_int_ConTem, 15)).Interior.Color = RGB(255, 255, 0)
         End If
         
         If Val(g_rst_Princi!RPT_CODIGO) = "9" Or Val(g_rst_Princi!RPT_CODIGO) = "10" Then
            .Cells(r_int_ConTem, 1) = "       " & Trim(g_rst_Princi!RPT_DESCRI)
         Else
            .Cells(r_int_ConTem, 1) = Trim(g_rst_Princi!RPT_DESCRI)
         End If
                  
         .Cells(r_int_ConTem, 2) = IIf(IsNull(g_rst_Princi!RPT_VALNUM01), 0, g_rst_Princi!RPT_VALNUM01)
         .Cells(r_int_ConTem, 3) = IIf(IsNull(g_rst_Princi!RPT_VALNUM02), 0, g_rst_Princi!RPT_VALNUM02)
         .Cells(r_int_ConTem, 4) = IIf(IsNull(g_rst_Princi!RPT_VALNUM03), 0, g_rst_Princi!RPT_VALNUM03)
         .Cells(r_int_ConTem, 5) = IIf(IsNull(g_rst_Princi!RPT_VALNUM04), 0, g_rst_Princi!RPT_VALNUM04)
         .Cells(r_int_ConTem, 6) = IIf(IsNull(g_rst_Princi!RPT_VALNUM05), 0, g_rst_Princi!RPT_VALNUM05)
         .Cells(r_int_ConTem, 8) = IIf(IsNull(g_rst_Princi!RPT_VALNUM06), 0, g_rst_Princi!RPT_VALNUM06)
         .Cells(r_int_ConTem, 9) = IIf(IsNull(g_rst_Princi!RPT_VALNUM07), 0, g_rst_Princi!RPT_VALNUM07)
         .Cells(r_int_ConTem, 10) = IIf(IsNull(g_rst_Princi!RPT_VALNUM08), 0, g_rst_Princi!RPT_VALNUM08)
         .Cells(r_int_ConTem, 11) = IIf(IsNull(g_rst_Princi!RPT_VALNUM09), 0, g_rst_Princi!RPT_VALNUM09)
         .Cells(r_int_ConTem, 12) = IIf(IsNull(g_rst_Princi!RPT_VALNUM10), 0, g_rst_Princi!RPT_VALNUM10)
         .Cells(r_int_ConTem, 13) = IIf(IsNull(g_rst_Princi!RPT_VALNUM11), 0, g_rst_Princi!RPT_VALNUM11)
         .Cells(r_int_ConTem, 14) = IIf(IsNull(g_rst_Princi!RPT_VALNUM12), 0, g_rst_Princi!RPT_VALNUM12)
         .Cells(r_int_ConTem, 15) = IIf(IsNull(g_rst_Princi!RPT_VALNUM13), 0, g_rst_Princi!RPT_VALNUM13)
         .Range(.Cells(r_int_ConTem, 7), .Cells(r_int_ConTem, 7)).Interior.Color = RGB(255, 255, 0)
         
         For r_int_Contad = 1 To 16
            .Range(.Cells(r_int_ConTem, r_int_Contad), .Cells(r_int_ConTem, r_int_Contad)).Borders(xlEdgeLeft).LineStyle = xlContinuous
         Next
         .Range(.Cells(r_int_ConTem, 1), .Cells(r_int_ConTem, 15)).Borders(xlEdgeBottom).LineStyle = xlContinuous

         r_int_ConTem = r_int_ConTem + 1
         g_rst_Princi.MoveNext
      Loop
      
      .Cells(16, 15).Formula = "=SUM(O13:O14)"
      .Range(.Cells(17, 1), .Cells(17, 2)).Font.Name = "Calibri"
      .Range(.Cells(17, 1), .Cells(17, 2)).Font.Size = 9
      .Range(.Cells(17, 1), .Cells(17, 2)).Font.Bold = True
      
      .Cells(17, 1) = "TIPO DE MONEDA"
      .Cells(17, 2) = "DOLARES"
      .Range(.Cells(17, 1), .Cells(17, 1)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(17, 2), .Cells(17, 2)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(19, 1), .Cells(19, 1)).HorizontalAlignment = xlHAlignCenter
      
      .Range(.Cells(19, 2), .Cells(19, 6)).Merge
      .Range(.Cells(19, 2), .Cells(19, 6)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(19, 2), .Cells(19, 6)).Font.Bold = True
      
      .Range(.Cells(19, 8), .Cells(19, 14)).Merge
      .Range(.Cells(19, 8), .Cells(19, 14)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(19, 8), .Cells(19, 14)).Font.Bold = True
      
      .Range(.Cells(19, 15), .Cells(19, 15)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(19, 1), .Cells(19, 25)).Font.Name = "Calibri"
      .Range(.Cells(19, 1), .Cells(19, 25)).Font.Size = 9
      .Range(.Cells(19, 1), .Cells(19, 25)).Font.Bold = True
      
      .Cells(19, 1) = "DESCRIPCION"
      .Cells(19, 2) = "IND"
      .Cells(19, 8) = "MAN"
      .Cells(19, 15) = "TOTAL"
      
      .Range(.Cells(19, 1), .Cells(19, 15)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(19, 1), .Cells(19, 15)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(19, 1), .Cells(19, 15)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(19, 1), .Cells(19, 15)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(19, 1), .Cells(19, 15)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(19, 1), .Cells(19, 15)).Borders(xlInsideVertical).LineStyle = xlContinuous

      'Dolares
      g_str_Parame = ""
      g_str_Parame = "USP_RPT_SEGDESG ("
      g_str_Parame = g_str_Parame & "" & r_str_PerMes & ", "
      g_str_Parame = g_str_Parame & "" & r_str_PerAno & ","
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "',"
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "',"
      g_str_Parame = g_str_Parame & "'REPORTE DE DESGRAVAMEN',"
      g_str_Parame = g_str_Parame & "'2')"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         MsgBox "Error al ejecutar procedimiento de Desgravamen.", vbCritical, modgen_g_str_NomPlt
         Exit Sub
      End If
      
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT * FROM RPT_TABLA_TEMP "
      g_str_Parame = g_str_Parame + " WHERE RPT_PERMES = '" & r_str_PerMes & "' "
      g_str_Parame = g_str_Parame + "   AND RPT_PERANO = '" & r_str_PerAno & "' "
      g_str_Parame = g_str_Parame + "   AND RPT_TERCRE = '" & modgen_g_str_NombPC & "' "
      g_str_Parame = g_str_Parame + "   AND RPT_USUCRE = '" & modgen_g_str_CodUsu & "' "
      g_str_Parame = g_str_Parame + "   AND TRIM(RPT_NOMBRE) = 'REPORTE DE DESGRAVAMEN' "
      g_str_Parame = g_str_Parame + "   AND RPT_MONEDA = '2' "
      g_str_Parame = g_str_Parame + " ORDER BY TO_NUMBER(RPT_CODIGO)"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      If g_rst_Princi.BOF And g_rst_Princi.EOF Then
         g_rst_Princi.Close
         Set g_rst_Princi = Nothing
         MsgBox "No se encontraron registros.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
   
      r_int_ConTem = 20
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         .Range(.Cells(r_int_ConTem, 1), .Cells(r_int_ConTem, 25)).Font.Name = "Calibri"
         .Range(.Cells(r_int_ConTem, 1), .Cells(r_int_ConTem, 25)).Font.Size = 9
         
         If Val(g_rst_Princi!RPT_CODIGO) >= "3" Then
            .Range(.Cells(r_int_ConTem, 1), .Cells(r_int_ConTem, 1)).NumberFormat = "###,##0.00"
            .Range(.Cells(r_int_ConTem, 2), .Cells(r_int_ConTem, 2)).NumberFormat = "###,##0.00"
            .Range(.Cells(r_int_ConTem, 3), .Cells(r_int_ConTem, 3)).NumberFormat = "###,##0.00"
            .Range(.Cells(r_int_ConTem, 4), .Cells(r_int_ConTem, 4)).NumberFormat = "###,##0.00"
            .Range(.Cells(r_int_ConTem, 5), .Cells(r_int_ConTem, 5)).NumberFormat = "###,##0.00"
            .Range(.Cells(r_int_ConTem, 6), .Cells(r_int_ConTem, 6)).NumberFormat = "###,##0.00"
            .Range(.Cells(r_int_ConTem, 8), .Cells(r_int_ConTem, 8)).NumberFormat = "###,##0.00"
            .Range(.Cells(r_int_ConTem, 9), .Cells(r_int_ConTem, 9)).NumberFormat = "###,##0.00"
            .Range(.Cells(r_int_ConTem, 10), .Cells(r_int_ConTem, 10)).NumberFormat = "###,##0.00"
            .Range(.Cells(r_int_ConTem, 11), .Cells(r_int_ConTem, 11)).NumberFormat = "###,##0.00"
            .Range(.Cells(r_int_ConTem, 12), .Cells(r_int_ConTem, 12)).NumberFormat = "###,##0.00"
            .Range(.Cells(r_int_ConTem, 13), .Cells(r_int_ConTem, 13)).NumberFormat = "###,##0.00"
            .Range(.Cells(r_int_ConTem, 14), .Cells(r_int_ConTem, 14)).NumberFormat = "###,##0.00"
            .Range(.Cells(r_int_ConTem, 15), .Cells(r_int_ConTem, 15)).NumberFormat = "###,##0.00"
         End If

         If Val(g_rst_Princi!RPT_CODIGO) = "3" Then
            .Range(.Cells(r_int_ConTem, 1), .Cells(r_int_ConTem, 15)).Interior.Color = RGB(255, 255, 0)
         End If
         If Val(g_rst_Princi!RPT_CODIGO) = "9" Or Val(g_rst_Princi!RPT_CODIGO) = "10" Then
            .Cells(r_int_ConTem, 1) = "       " & Trim(g_rst_Princi!RPT_DESCRI)
         Else
            .Cells(r_int_ConTem, 1) = Trim(g_rst_Princi!RPT_DESCRI)
         End If
         
         .Cells(r_int_ConTem, 2) = IIf(IsNull(g_rst_Princi!RPT_VALNUM01), 0, g_rst_Princi!RPT_VALNUM01)
         .Cells(r_int_ConTem, 3) = IIf(IsNull(g_rst_Princi!RPT_VALNUM02), 0, g_rst_Princi!RPT_VALNUM02)
         .Cells(r_int_ConTem, 4) = IIf(IsNull(g_rst_Princi!RPT_VALNUM03), 0, g_rst_Princi!RPT_VALNUM03)
         .Cells(r_int_ConTem, 5) = IIf(IsNull(g_rst_Princi!RPT_VALNUM04), 0, g_rst_Princi!RPT_VALNUM04)
         .Cells(r_int_ConTem, 6) = IIf(IsNull(g_rst_Princi!RPT_VALNUM05), 0, g_rst_Princi!RPT_VALNUM05)
         .Cells(r_int_ConTem, 8) = IIf(IsNull(g_rst_Princi!RPT_VALNUM06), 0, g_rst_Princi!RPT_VALNUM06)
         .Cells(r_int_ConTem, 9) = IIf(IsNull(g_rst_Princi!RPT_VALNUM07), 0, g_rst_Princi!RPT_VALNUM07)
         .Cells(r_int_ConTem, 10) = IIf(IsNull(g_rst_Princi!RPT_VALNUM08), 0, g_rst_Princi!RPT_VALNUM08)
         .Cells(r_int_ConTem, 11) = IIf(IsNull(g_rst_Princi!RPT_VALNUM09), 0, g_rst_Princi!RPT_VALNUM09)
         .Cells(r_int_ConTem, 12) = IIf(IsNull(g_rst_Princi!RPT_VALNUM10), 0, g_rst_Princi!RPT_VALNUM10)
         .Cells(r_int_ConTem, 13) = IIf(IsNull(g_rst_Princi!RPT_VALNUM11), 0, g_rst_Princi!RPT_VALNUM11)
         .Cells(r_int_ConTem, 14) = IIf(IsNull(g_rst_Princi!RPT_VALNUM12), 0, g_rst_Princi!RPT_VALNUM12)
         .Cells(r_int_ConTem, 15) = IIf(IsNull(g_rst_Princi!RPT_VALNUM13), 0, g_rst_Princi!RPT_VALNUM13)
         
         .Range(.Cells(r_int_ConTem, 7), .Cells(r_int_ConTem, 7)).Interior.Color = RGB(255, 255, 0)
         
         For r_int_Contad = 1 To 16
            .Range(.Cells(r_int_ConTem, r_int_Contad), .Cells(r_int_ConTem, r_int_Contad)).Borders(xlEdgeLeft).LineStyle = xlContinuous
         Next
         .Range(.Cells(r_int_ConTem, 1), .Cells(r_int_ConTem, 15)).Borders(xlEdgeBottom).LineStyle = xlContinuous

         r_int_ConTem = r_int_ConTem + 1
         g_rst_Princi.MoveNext
      Loop
      
      .Cells(31, 15).Formula = "=SUM(O28:O29)"
   End With

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub
Private Sub fs_GenExc_SegDes_old()
Dim r_obj_Excel      As Excel.Application
Dim r_int_Contad     As Integer
Dim r_int_ConVer     As Integer
Dim r_dbl_PBPPer     As Double
Dim r_str_Fecha      As String
Dim r_str_FecIni     As String
Dim r_str_FecFin     As String
Dim r_int_Index      As Integer
Dim r_dbl_Monto      As Double
Dim r_dbl_MtoInd     As Double
Dim r_dbl_MtoMan     As Double
Dim r_int_ItmReg     As Integer
Dim r_str_PerMes     As Integer
Dim r_str_PerAno     As Integer
Dim r_int_ConTem     As Integer

   r_str_PerMes = CInt(cmb_PerMes.ItemData(cmb_PerMes.ListIndex))
   r_str_PerAno = CInt(ipp_PerAno.Text)
   r_str_FecIni = Format(ipp_PerAno.Text, "0000") & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") & "25"
   If cmb_PerMes.ItemData(cmb_PerMes.ListIndex) = 12 Then
      r_str_FecFin = Format(ipp_PerAno.Text + 1, "0000") & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex) + 1, "00") & "02"
   Else
      r_str_FecFin = Format(ipp_PerAno.Text, "0000") & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex) + 1, "00") & "02"
   End If
   r_str_Fecha = "01/" & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") & "/" & Format(ipp_PerAno.Text, "0000")
   
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 2
   r_obj_Excel.Workbooks.Add
   r_obj_Excel.Sheets(1).Name = "DETALLE"
   
   r_dbl_Monto = 0
   With r_obj_Excel.ActiveSheet
      .Cells(1, 1) = ObtieneNomMes(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)) & " - " & Format(ipp_PerAno.Text, "0000")
      .Range(.Cells(1, 1), .Cells(1, 16)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(1, 1), .Cells(1, 16)).Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 16)).Merge
      .Range(.Cells(1, 1), .Cells(1, 16)).Font.Size = 18
      .Cells(3, 1) = "MONEDA: SOLES"
      .Range(.Cells(3, 1), .Cells(3, 3)).HorizontalAlignment = xlHAlignLeft
      .Range(.Cells(3, 1), .Cells(3, 3)).Font.Bold = True
      .Range(.Cells(3, 1), .Cells(3, 3)).Merge
      
      For r_int_Contad = 1 To 18 Step 1
         .Range(.Cells(5, 1), .Cells(5, r_int_Contad)).Borders(xlEdgeLeft).LineStyle = xlContinuous
         .Range(.Cells(5, 1), .Cells(5, r_int_Contad)).Borders(xlEdgeTop).LineStyle = xlContinuous
         .Range(.Cells(5, 1), .Cells(5, r_int_Contad)).Borders(xlEdgeBottom).LineStyle = xlContinuous
         .Range(.Cells(5, 1), .Cells(5, r_int_Contad)).Borders(xlEdgeRight).LineStyle = xlContinuous
         .Range(.Cells(5, 1), .Cells(5, r_int_Contad)).Borders(xlInsideVertical).LineStyle = xlContinuous
      Next
      
      .Cells(5, 1) = "ITEM"
      .Cells(5, 2) = "NRO. PRESTAMO"
      .Cells(5, 3) = "MONEDA"
      .Cells(5, 4) = "CODIGO"
      .Cells(5, 5) = "NOMBRE CLIENTE"
      .Cells(5, 6) = "F. NACIMIENTO"
      .Cells(5, 7) = "TIPO DOCUMENTO"
      .Cells(5, 8) = "DOC. IDENTIDAD"
      .Cells(5, 9) = "EMP. SEGURO"
      .Cells(5, 10) = "F. APERTURA"
      .Cells(5, 11) = "PRESTAMO (AÑOS)"
      .Cells(5, 12) = "F. PAGO"
      .Cells(5, 13) = "IMP. PRESTAMO"
      .Cells(5, 14) = "SALDO PRESTAMO"
      .Cells(5, 15) = "COBERTURA"
      .Cells(5, 16) = "FACTOR APLICACION"
      .Cells(5, 17) = "MONTO INDIV."
      .Cells(5, 18) = "MONTO MANCO."
      
      .Range(.Cells(5, 1), .Cells(5, 18)).Font.Bold = True
      .Range(.Cells(5, 1), .Cells(5, 18)).HorizontalAlignment = xlHAlignCenter
      
      .Columns("A").ColumnWidth = 6
      .Columns("A").HorizontalAlignment = xlHAlignCenter
      .Columns("B").ColumnWidth = 16
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 22
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 12
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").ColumnWidth = 40
      .Columns("E").HorizontalAlignment = xlHAlignLeft
      .Columns("F").ColumnWidth = 23
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("G").ColumnWidth = 18
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      .Columns("H").ColumnWidth = 15
      .Columns("H").HorizontalAlignment = xlHAlignCenter
      .Columns("I").ColumnWidth = 49
      .Columns("I").HorizontalAlignment = xlHAlignCenter
      .Columns("J").ColumnWidth = 13
      .Columns("J").HorizontalAlignment = xlHAlignCenter
      .Columns("K").ColumnWidth = 18
      .Columns("K").HorizontalAlignment = xlHAlignCenter
      .Columns("L").ColumnWidth = 12
      .Columns("L").HorizontalAlignment = xlHAlignCenter
      .Columns("M").ColumnWidth = 18
      .Columns("M").NumberFormat = "###,###,##0.00"
      .Columns("N").ColumnWidth = 18
      .Columns("N").NumberFormat = "###,###,##0.00"
      .Columns("O").ColumnWidth = 25
      .Columns("O").HorizontalAlignment = xlHAlignCenter
      .Columns("P").ColumnWidth = 20
      .Columns("P").NumberFormat = "###,##0.000000"
      .Columns("Q").ColumnWidth = 16
      .Columns("Q").NumberFormat = "###,###,##0.00"
      .Columns("R").ColumnWidth = 16
      .Columns("R").NumberFormat = "###,###,##0.00"
   End With
   
   r_int_ConVer = 6
   For r_int_Index = 1 To 2 Step 1
      If r_int_Index = 2 Then
         r_dbl_Monto = 0
         r_int_ConVer = r_int_ConVer + 4
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = "MONEDA: DOLARES"
         r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1), r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3)).HorizontalAlignment = xlHAlignLeft
         r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1), r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3)).Font.Bold = True
         r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1), r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3)).Merge
         r_int_ConVer = r_int_ConVer + 2
         
         For r_int_Contad = 1 To 18 Step 1
            r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1), r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, r_int_Contad)).Borders(xlEdgeLeft).LineStyle = xlContinuous
            r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1), r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, r_int_Contad)).Borders(xlEdgeTop).LineStyle = xlContinuous
            r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1), r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, r_int_Contad)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1), r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, r_int_Contad)).Borders(xlEdgeRight).LineStyle = xlContinuous
            r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1), r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, r_int_Contad)).Borders(xlInsideVertical).LineStyle = xlContinuous
         Next
         
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = "ITEM"
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = "NRO. PRESTAMO"
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = "MONEDA"
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = "CODIGO"
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = "NOMBRE CLIENTE"
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = "F. NACIMIENTO"
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = "TIPO DOCUMENTO IDENTIDAD"
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = "DOC. IDENTIDAD"
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = "EMP. SEGURO"
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = "F. APERTURA"
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = "PRESTAMO(AÑOS)"
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = "F. PAGO"
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = "IMP. PRESTAMO"
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14) = "SALDO PRESTAMO"
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = "COBERTURA"
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 16) = "FACTOR APLICACION"
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 17) = "MONTO INDIV."
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 18) = "MONTO MANCO."
         
         r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1), r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 18)).Font.Bold = True
         r_obj_Excel.ActiveSheet.Range(r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1), r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 18)).HorizontalAlignment = xlHAlignCenter
         r_int_ConVer = r_int_ConVer + 1
      End If
      
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "SELECT HIPCIE_NUMOPE AS NROPRESTAMO, "
      g_str_Parame = g_str_Parame & "       TRIM(HIPCIE_TDOCLI)||'-'||TRIM(HIPCIE_NDOCLI) AS CODIGO, "
      g_str_Parame = g_str_Parame & "       TRIM(DATGEN_APEPAT)||' '||TRIM(DATGEN_APEMAT)||' '||TRIM(DATGEN_NOMBRE) AS NOMCLIENTE, "
      g_str_Parame = g_str_Parame & "       SUBSTR(DATGEN_NACFEC,7,2)||'/'||SUBSTR(DATGEN_NACFEC,5,2)||'/'||SUBSTR(DATGEN_NACFEC,1,4) AS FECHANAC, "
      g_str_Parame = g_str_Parame & "       TRIM(HIPCIE_NDOCLI) AS DOCIDENTIDAD, "
      g_str_Parame = g_str_Parame & "       TRIM(SEGEMP_RAZSOC) AS EMPSEGUROS, "
      g_str_Parame = g_str_Parame & "       SUBSTR(POLIZA_FEMDES,7,2)||'/'||SUBSTR(POLIZA_FEMDES,5,2)||'/'||SUBSTR(POLIZA_FEMDES,1,4) AS FECHAAPERTURA, "
      g_str_Parame = g_str_Parame & "       HIPMAE_PLAANO AS DURACION, "
      g_str_Parame = g_str_Parame & "       TO_CHAR(TO_DATE('" & r_str_Fecha & "','DD/MM/YYYY'), 'MONTH', 'NLS_DATE_LANGUAGE=SPANISH') AS FECPAGO, "
      g_str_Parame = g_str_Parame & "       HIPCIE_MTOPRE AS IMPPRESTAMO, "
      g_str_Parame = g_str_Parame & "       HIPCIE_SALCAP+HIPCIE_SALCON AS SALDOPRESTAMO, "
      g_str_Parame = g_str_Parame & "       EVASEG_TIPSEG AS TIPOSEGURO, "
      g_str_Parame = g_str_Parame & "       TRIM(SEGTIP_DESCRI) AS COBERTURA, "
      g_str_Parame = g_str_Parame & "       ROUND(HIPCIE_FOIPRE, 5) AS FACTORAPLICA, "
      g_str_Parame = g_str_Parame & "       ROUND(((HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_FOIPRE)/100,2) AS MONTO, HIPCIE_TDOCLI, TRIM(PARDES_DESCRI) AS MONEDA "
      g_str_Parame = g_str_Parame & "  FROM CRE_HIPCIE "
      g_str_Parame = g_str_Parame & " INNER JOIN CLI_DATGEN ON DATGEN_TIPDOC = HIPCIE_TDOCLI AND DATGEN_NUMDOC = HIPCIE_NDOCLI "
      g_str_Parame = g_str_Parame & " INNER JOIN CRE_HIPMAE ON HIPMAE_NUMOPE = HIPCIE_NUMOPE "
      g_str_Parame = g_str_Parame & " INNER JOIN TRA_POLIZA ON POLIZA_NUMSOL = HIPMAE_NUMSOL "
      g_str_Parame = g_str_Parame & " INNER JOIN TRA_EVASEG ON EVASEG_NUMSOL = HIPMAE_NUMSOL AND EVASEG_TIPSEG <> 13"
      g_str_Parame = g_str_Parame & " INNER JOIN MNT_SEGEMP ON SEGEMP_CODIGO = EVASEG_ESGDES "
      g_str_Parame = g_str_Parame & " INNER JOIN MNT_SEGTIP ON SEGTIP_CODIGO = EVASEG_ESGDES AND SEGTIP_TIPSEG = EVASEG_TIPSEG "
      g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES ON PARDES_CODGRP = 204 AND PARDES_CODITE = HIPCIE_TIPMON"
      g_str_Parame = g_str_Parame & " WHERE HIPCIE_PERMES = " & cmb_PerMes.ItemData(cmb_PerMes.ListIndex) & " "
      g_str_Parame = g_str_Parame & "   AND HIPCIE_PERANO = " & Format(ipp_PerAno.Text, "0000") & " "
      g_str_Parame = g_str_Parame & "   AND HIPCIE_TIPMON = " & r_int_Index & " "
      g_str_Parame = g_str_Parame & "UNION "
      g_str_Parame = g_str_Parame & "SELECT HIPMAE_NUMOPE AS NROPRESTAMO, "
      g_str_Parame = g_str_Parame & "       TRIM(HIPMAE_TDOCLI)||'-'||TRIM(HIPMAE_NDOCLI) AS CODIGO, "
      g_str_Parame = g_str_Parame & "       TRIM(DATGEN_APEPAT)||' '||TRIM(DATGEN_APEMAT)||' '||TRIM(DATGEN_NOMBRE) AS NOMCLIENTE, "
      g_str_Parame = g_str_Parame & "       SUBSTR(DATGEN_NACFEC,7,2)||'/'||SUBSTR(DATGEN_NACFEC,5,2)||'/'||SUBSTR(DATGEN_NACFEC,1,4) AS FECHANAC, "
      g_str_Parame = g_str_Parame & "       TRIM(HIPMAE_NDOCLI) AS DOCIDENTIDAD, "
      g_str_Parame = g_str_Parame & "       TRIM(SEGEMP_RAZSOC) AS EMPSEGUROS, "
      g_str_Parame = g_str_Parame & "       SUBSTR(POLIZA_FEMDES,7,2)||'/'||SUBSTR(POLIZA_FEMDES,5,2)||'/'||SUBSTR(POLIZA_FEMDES,1,4) AS FECHAAPERTURA, "
      g_str_Parame = g_str_Parame & "       HIPMAE_PLAANO AS DURACION, "
      g_str_Parame = g_str_Parame & "       TO_CHAR(TO_DATE('" & r_str_Fecha & "','DD/MM/YYYY'), 'MONTH', 'NLS_DATE_LANGUAGE=SPANISH') AS FECPAGO, "
      g_str_Parame = g_str_Parame & "       HIPMAE_MTOPRE AS IMPPRESTAMO, "
      g_str_Parame = g_str_Parame & "       HIPMAE_SALCAP+HIPMAE_SALCON AS SALDOPRESTAMO, "
      g_str_Parame = g_str_Parame & "       EVASEG_TIPSEG AS TIPOSEGURO, "
      g_str_Parame = g_str_Parame & "       TRIM(SEGTIP_DESCRI) AS COBERTURA, "
      g_str_Parame = g_str_Parame & "       ROUND(HIPMAE_FOIPRE, 5) AS FACTORAPLICA, "
      g_str_Parame = g_str_Parame & "       ROUND(((HIPMAE_SALCAP+HIPMAE_SALCON)*HIPMAE_FOIPRE)/100,2) AS MONTO, HIPMAE_TDOCLI AS HIPCIE_TDOCLI, TRIM(PARDES_DESCRI) AS MONEDA "
      g_str_Parame = g_str_Parame & "  FROM CRE_HIPMAE "
      g_str_Parame = g_str_Parame & " INNER JOIN CLI_DATGEN ON DATGEN_TIPDOC = HIPMAE_TDOCLI AND DATGEN_NUMDOC = HIPMAE_NDOCLI "
      g_str_Parame = g_str_Parame & " INNER JOIN TRA_POLIZA ON POLIZA_NUMSOL = HIPMAE_NUMSOL "
      g_str_Parame = g_str_Parame & " INNER JOIN TRA_EVASEG ON EVASEG_NUMSOL = HIPMAE_NUMSOL AND EVASEG_TIPSEG <> 13"
      g_str_Parame = g_str_Parame & " INNER JOIN MNT_SEGEMP ON SEGEMP_CODIGO = EVASEG_ESGDES "
      g_str_Parame = g_str_Parame & " INNER JOIN MNT_SEGTIP ON SEGTIP_CODIGO = EVASEG_ESGDES AND SEGTIP_TIPSEG = EVASEG_TIPSEG "
      g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES ON PARDES_CODGRP = 204 AND PARDES_CODITE = HIPMAE_MONEDA "
      g_str_Parame = g_str_Parame & " WHERE HIPMAE_SITUAC = 6 "
      g_str_Parame = g_str_Parame & "   AND HIPMAE_MONEDA = " & r_int_Index & " "
      g_str_Parame = g_str_Parame & "   AND HIPMAE_FECCAN >= " & r_str_FecIni
      g_str_Parame = g_str_Parame & "   AND HIPMAE_FECCAN <= " & r_str_FecFin
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      If g_rst_Princi.BOF And g_rst_Princi.EOF Then
         g_rst_Princi.Close
         Set g_rst_Princi = Nothing
         Exit Sub
      End If
      
      r_dbl_MtoInd = 0
      r_dbl_MtoMan = 0
      r_int_ItmReg = 1
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ItmReg
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = gf_Formato_NumOpe(g_rst_Princi!NROPRESTAMO)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = Trim(g_rst_Princi!Moneda)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = Trim(g_rst_Princi!Codigo)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = Trim(g_rst_Princi!NOMCLIENTE)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = "'" & CStr(g_rst_Princi!FECHANAC)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = "'" & CStr(g_rst_Princi!HIPCIE_TDOCLI)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = "'" & CStr(g_rst_Princi!DOCIDENTIDAD)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = Trim(g_rst_Princi!EMPSEGUROS)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = CStr(g_rst_Princi!FECHAAPERTURA)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = Trim(g_rst_Princi!DURACION)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = "'" & Trim(g_rst_Princi!FECPAGO)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = Format(g_rst_Princi!IMPPRESTAMO, "###,###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14) = Format(g_rst_Princi!SALDOPRESTAMO, "###,###,###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = Trim(g_rst_Princi!COBERTURA)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 16) = Format(g_rst_Princi!FACTORAPLICA, "###,###,###,##0.0000")
         If CInt(g_rst_Princi!TIPOSEGURO) = 11 Then
             r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 17) = Format(g_rst_Princi!MONTO, "###,###,###,##0.00")
             r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 18) = Format(0, "###,###,###,##0.00")
             r_dbl_MtoInd = r_dbl_MtoInd + g_rst_Princi!MONTO
         Else
             r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 17) = Format(0, "###,###,###,##0.00")
             r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 18) = Format(g_rst_Princi!MONTO, "###,###,###,##0.00")
             r_dbl_MtoMan = r_dbl_MtoMan + g_rst_Princi!MONTO
         End If
         
         r_int_ConVer = r_int_ConVer + 1
         r_int_ItmReg = r_int_ItmReg + 1
         g_rst_Princi.MoveNext
         DoEvents
      Loop
      
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 17) = Format(r_dbl_MtoInd, "###,###,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 18) = Format(r_dbl_MtoMan, "###,###,###,##0.00")
      
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   Next
   
   r_obj_Excel.Sheets(2).Name = "RESUMEN"
   With r_obj_Excel.Sheets(2)
      .Range(.Cells(4, 1), .Cells(4, 25)).Font.Name = "Calibri"
      .Range(.Cells(4, 1), .Cells(4, 25)).Font.Size = 9
      .Range(.Cells(4, 1), .Cells(4, 25)).Font.Bold = True
      
      .Range(.Cells(4, 2), .Cells(4, 6)).Merge
      .Range(.Cells(4, 2), .Cells(4, 6)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(4, 2), .Cells(4, 6)).Font.Bold = True
      
      .Range(.Cells(4, 8), .Cells(4, 14)).Merge
      .Range(.Cells(4, 8), .Cells(4, 14)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(4, 8), .Cells(4, 14)).Font.Bold = True
      
      .Range(.Cells(4, 1), .Cells(4, 1)).Merge
      .Range(.Cells(4, 1), .Cells(4, 1)).WrapText = True
      .Range(.Cells(4, 1), .Cells(4, 1)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(4, 1), .Cells(4, 1)).ColumnWidth = 17
      
      .Range(.Cells(4, 2), .Cells(4, 2)).Merge
      .Range(.Cells(4, 2), .Cells(4, 2)).WrapText = True
      .Range(.Cells(4, 2), .Cells(4, 2)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(4, 2), .Cells(4, 2)).ColumnWidth = 11
      
      .Range(.Cells(4, 3), .Cells(4, 3)).Merge
      .Range(.Cells(4, 3), .Cells(4, 3)).WrapText = True
      .Range(.Cells(4, 3), .Cells(4, 3)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(4, 3), .Cells(4, 3)).ColumnWidth = 11
      
      .Range(.Cells(4, 4), .Cells(4, 4)).Merge
      .Range(.Cells(4, 4), .Cells(4, 4)).WrapText = True
      .Range(.Cells(4, 4), .Cells(4, 4)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(4, 4), .Cells(4, 4)).ColumnWidth = 11
      
      .Range(.Cells(4, 5), .Cells(4, 5)).Merge
      .Range(.Cells(4, 5), .Cells(4, 5)).WrapText = True
      .Range(.Cells(4, 5), .Cells(4, 5)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(4, 5), .Cells(4, 5)).ColumnWidth = 11
      
      .Range(.Cells(4, 6), .Cells(4, 6)).Merge
      .Range(.Cells(4, 6), .Cells(4, 6)).WrapText = True
      .Range(.Cells(4, 6), .Cells(4, 6)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(4, 6), .Cells(4, 6)).ColumnWidth = 11
      
      .Range(.Cells(4, 7), .Cells(4, 7)).ColumnWidth = 4
      
      .Range(.Cells(4, 8), .Cells(4, 8)).Merge
      .Range(.Cells(4, 8), .Cells(4, 8)).WrapText = True
      .Range(.Cells(4, 8), .Cells(4, 8)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(4, 8), .Cells(4, 8)).ColumnWidth = 11
      
      .Range(.Cells(4, 9), .Cells(4, 9)).Merge
      .Range(.Cells(4, 9), .Cells(4, 9)).WrapText = True
      .Range(.Cells(4, 9), .Cells(4, 9)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(4, 9), .Cells(4, 9)).ColumnWidth = 11
      
      .Range(.Cells(4, 10), .Cells(4, 10)).Merge
      .Range(.Cells(4, 10), .Cells(4, 10)).WrapText = True
      .Range(.Cells(4, 10), .Cells(4, 10)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(4, 10), .Cells(4, 10)).ColumnWidth = 11
      
      .Range(.Cells(4, 11), .Cells(4, 11)).Merge
      .Range(.Cells(4, 11), .Cells(4, 11)).WrapText = True
      .Range(.Cells(4, 11), .Cells(4, 11)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(4, 11), .Cells(4, 11)).ColumnWidth = 11
      
      .Range(.Cells(4, 12), .Cells(4, 12)).Merge
      .Range(.Cells(4, 12), .Cells(4, 12)).WrapText = True
      .Range(.Cells(4, 12), .Cells(4, 12)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(4, 12), .Cells(4, 12)).ColumnWidth = 11
      
      .Range(.Cells(4, 13), .Cells(4, 13)).Merge
      .Range(.Cells(4, 13), .Cells(4, 13)).WrapText = True
      .Range(.Cells(4, 13), .Cells(4, 13)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(4, 13), .Cells(4, 13)).ColumnWidth = 11
      
      .Range(.Cells(4, 14), .Cells(4, 14)).Merge
      .Range(.Cells(4, 14), .Cells(4, 14)).WrapText = True
      .Range(.Cells(4, 14), .Cells(4, 14)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(4, 14), .Cells(4, 14)).ColumnWidth = 11
      
      .Range(.Cells(4, 15), .Cells(4, 15)).Merge
      .Range(.Cells(4, 15), .Cells(4, 15)).WrapText = True
      .Range(.Cells(4, 15), .Cells(4, 15)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(4, 15), .Cells(4, 15)).ColumnWidth = 11
      
      .Range(.Cells(2, 1), .Cells(2, 2)).Font.Name = "Calibri"
      .Range(.Cells(2, 1), .Cells(2, 2)).Font.Size = 9
      .Range(.Cells(2, 1), .Cells(2, 2)).Font.Bold = True
      
      .Cells(2, 1) = "TIPO DE MONEDA"
      .Cells(2, 2) = "SOLES"
      .Range(.Cells(2, 1), .Cells(2, 1)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(2, 2), .Cells(2, 2)).HorizontalAlignment = xlHAlignCenter
      
      .Cells(4, 1) = "DESCRIPCION"
      .Cells(4, 2) = "IND"
      .Cells(4, 8) = "MAN"
      .Cells(4, 15) = "TOTAL"
      
      .Range(.Cells(4, 1), .Cells(4, 15)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(4, 1), .Cells(4, 15)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(4, 1), .Cells(4, 15)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(4, 1), .Cells(4, 15)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(4, 1), .Cells(4, 15)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(4, 1), .Cells(4, 15)).Borders(xlInsideVertical).LineStyle = xlContinuous
      
      g_str_Parame = ""
      g_str_Parame = "USP_RPT_SEGDESG ("
      g_str_Parame = g_str_Parame & "" & r_str_PerMes & ", "
      g_str_Parame = g_str_Parame & "" & r_str_PerAno & ","
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "',"
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "',"
      g_str_Parame = g_str_Parame & "'REPORTE DE DESGRAVAMEN',"
      g_str_Parame = g_str_Parame & "'1')"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         MsgBox "Error al ejecutar procedimiento de Desgravamen.", vbCritical, modgen_g_str_NomPlt
         Exit Sub
      End If
   
      'Soles
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT * FROM RPT_TABLA_TEMP "
      g_str_Parame = g_str_Parame + " WHERE RPT_PERMES = '" & r_str_PerMes & "' "
      g_str_Parame = g_str_Parame + "   AND RPT_PERANO = '" & r_str_PerAno & "' "
      g_str_Parame = g_str_Parame + "   AND RPT_TERCRE = '" & modgen_g_str_NombPC & "' "
      g_str_Parame = g_str_Parame + "   AND RPT_USUCRE = '" & modgen_g_str_CodUsu & "' "
      g_str_Parame = g_str_Parame + "   AND TRIM(RPT_NOMBRE) = 'REPORTE DE DESGRAVAMEN' "
      g_str_Parame = g_str_Parame + "   AND RPT_MONEDA = '1' "
      g_str_Parame = g_str_Parame + " ORDER BY TO_NUMBER(RPT_CODIGO)"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      If g_rst_Princi.BOF And g_rst_Princi.EOF Then
         g_rst_Princi.Close
         Set g_rst_Princi = Nothing
         MsgBox "No se encontraron registros.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
      
      r_int_ConTem = 5
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         .Range(.Cells(r_int_ConTem, 1), .Cells(r_int_ConTem, 25)).Font.Name = "Calibri"
         .Range(.Cells(r_int_ConTem, 1), .Cells(r_int_ConTem, 25)).Font.Size = 9
         
         If Val(g_rst_Princi!RPT_CODIGO) >= "3" Then
            .Range(.Cells(r_int_ConTem, 1), .Cells(r_int_ConTem, 1)).NumberFormat = "###,##0.00"
            .Range(.Cells(r_int_ConTem, 2), .Cells(r_int_ConTem, 2)).NumberFormat = "###,##0.00"
            .Range(.Cells(r_int_ConTem, 3), .Cells(r_int_ConTem, 3)).NumberFormat = "###,##0.00"
            .Range(.Cells(r_int_ConTem, 4), .Cells(r_int_ConTem, 4)).NumberFormat = "###,##0.00"
            .Range(.Cells(r_int_ConTem, 5), .Cells(r_int_ConTem, 5)).NumberFormat = "###,##0.00"
            .Range(.Cells(r_int_ConTem, 6), .Cells(r_int_ConTem, 6)).NumberFormat = "###,##0.00"
            .Range(.Cells(r_int_ConTem, 8), .Cells(r_int_ConTem, 8)).NumberFormat = "###,##0.00"
            .Range(.Cells(r_int_ConTem, 9), .Cells(r_int_ConTem, 9)).NumberFormat = "###,##0.00"
            .Range(.Cells(r_int_ConTem, 10), .Cells(r_int_ConTem, 10)).NumberFormat = "###,##0.00"
            .Range(.Cells(r_int_ConTem, 11), .Cells(r_int_ConTem, 11)).NumberFormat = "###,##0.00"
            .Range(.Cells(r_int_ConTem, 12), .Cells(r_int_ConTem, 12)).NumberFormat = "###,##0.00"
            .Range(.Cells(r_int_ConTem, 13), .Cells(r_int_ConTem, 13)).NumberFormat = "###,##0.00"
            .Range(.Cells(r_int_ConTem, 14), .Cells(r_int_ConTem, 14)).NumberFormat = "###,##0.00"
            .Range(.Cells(r_int_ConTem, 15), .Cells(r_int_ConTem, 15)).NumberFormat = "###,##0.00"
         End If
         
         If Val(g_rst_Princi!RPT_CODIGO) = "3" Then
            .Range(.Cells(r_int_ConTem, 1), .Cells(r_int_ConTem, 15)).Interior.Color = RGB(255, 255, 0)
         End If
         
         If Val(g_rst_Princi!RPT_CODIGO) = "9" Or Val(g_rst_Princi!RPT_CODIGO) = "10" Then
            .Cells(r_int_ConTem, 1) = "       " & Trim(g_rst_Princi!RPT_DESCRI)
         Else
            .Cells(r_int_ConTem, 1) = Trim(g_rst_Princi!RPT_DESCRI)
         End If
                  
         .Cells(r_int_ConTem, 2) = IIf(IsNull(g_rst_Princi!RPT_VALNUM01), 0, g_rst_Princi!RPT_VALNUM01)
         .Cells(r_int_ConTem, 3) = IIf(IsNull(g_rst_Princi!RPT_VALNUM02), 0, g_rst_Princi!RPT_VALNUM02)
         .Cells(r_int_ConTem, 4) = IIf(IsNull(g_rst_Princi!RPT_VALNUM03), 0, g_rst_Princi!RPT_VALNUM03)
         .Cells(r_int_ConTem, 5) = IIf(IsNull(g_rst_Princi!RPT_VALNUM04), 0, g_rst_Princi!RPT_VALNUM04)
         .Cells(r_int_ConTem, 6) = IIf(IsNull(g_rst_Princi!RPT_VALNUM05), 0, g_rst_Princi!RPT_VALNUM05)
         .Cells(r_int_ConTem, 8) = IIf(IsNull(g_rst_Princi!RPT_VALNUM06), 0, g_rst_Princi!RPT_VALNUM06)
         .Cells(r_int_ConTem, 9) = IIf(IsNull(g_rst_Princi!RPT_VALNUM07), 0, g_rst_Princi!RPT_VALNUM07)
         .Cells(r_int_ConTem, 10) = IIf(IsNull(g_rst_Princi!RPT_VALNUM08), 0, g_rst_Princi!RPT_VALNUM08)
         .Cells(r_int_ConTem, 11) = IIf(IsNull(g_rst_Princi!RPT_VALNUM09), 0, g_rst_Princi!RPT_VALNUM09)
         .Cells(r_int_ConTem, 12) = IIf(IsNull(g_rst_Princi!RPT_VALNUM10), 0, g_rst_Princi!RPT_VALNUM10)
         .Cells(r_int_ConTem, 13) = IIf(IsNull(g_rst_Princi!RPT_VALNUM11), 0, g_rst_Princi!RPT_VALNUM11)
         .Cells(r_int_ConTem, 14) = IIf(IsNull(g_rst_Princi!RPT_VALNUM12), 0, g_rst_Princi!RPT_VALNUM12)
         .Cells(r_int_ConTem, 15) = IIf(IsNull(g_rst_Princi!RPT_VALNUM13), 0, g_rst_Princi!RPT_VALNUM13)
         .Range(.Cells(r_int_ConTem, 7), .Cells(r_int_ConTem, 7)).Interior.Color = RGB(255, 255, 0)
         
         For r_int_Contad = 1 To 16
            .Range(.Cells(r_int_ConTem, r_int_Contad), .Cells(r_int_ConTem, r_int_Contad)).Borders(xlEdgeLeft).LineStyle = xlContinuous
         Next
         .Range(.Cells(r_int_ConTem, 1), .Cells(r_int_ConTem, 15)).Borders(xlEdgeBottom).LineStyle = xlContinuous

         r_int_ConTem = r_int_ConTem + 1
         g_rst_Princi.MoveNext
      Loop
      
      .Cells(16, 15).Formula = "=SUM(O13:O14)"
      .Range(.Cells(17, 1), .Cells(17, 2)).Font.Name = "Calibri"
      .Range(.Cells(17, 1), .Cells(17, 2)).Font.Size = 9
      .Range(.Cells(17, 1), .Cells(17, 2)).Font.Bold = True
      
      .Cells(17, 1) = "TIPO DE MONEDA"
      .Cells(17, 2) = "DOLARES"
      .Range(.Cells(17, 1), .Cells(17, 1)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(17, 2), .Cells(17, 2)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(19, 1), .Cells(19, 1)).HorizontalAlignment = xlHAlignCenter
      
      .Range(.Cells(19, 2), .Cells(19, 6)).Merge
      .Range(.Cells(19, 2), .Cells(19, 6)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(19, 2), .Cells(19, 6)).Font.Bold = True
      
      .Range(.Cells(19, 8), .Cells(19, 14)).Merge
      .Range(.Cells(19, 8), .Cells(19, 14)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(19, 8), .Cells(19, 14)).Font.Bold = True
      
      .Range(.Cells(19, 15), .Cells(19, 15)).HorizontalAlignment = xlHAlignCenter
      .Range(.Cells(19, 1), .Cells(19, 25)).Font.Name = "Calibri"
      .Range(.Cells(19, 1), .Cells(19, 25)).Font.Size = 9
      .Range(.Cells(19, 1), .Cells(19, 25)).Font.Bold = True
      
      .Cells(19, 1) = "DESCRIPCION"
      .Cells(19, 2) = "IND"
      .Cells(19, 8) = "MAN"
      .Cells(19, 15) = "TOTAL"
      
      .Range(.Cells(19, 1), .Cells(19, 15)).Interior.Color = RGB(146, 208, 80)
      .Range(.Cells(19, 1), .Cells(19, 15)).Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Range(.Cells(19, 1), .Cells(19, 15)).Borders(xlEdgeTop).LineStyle = xlContinuous
      .Range(.Cells(19, 1), .Cells(19, 15)).Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Range(.Cells(19, 1), .Cells(19, 15)).Borders(xlEdgeRight).LineStyle = xlContinuous
      .Range(.Cells(19, 1), .Cells(19, 15)).Borders(xlInsideVertical).LineStyle = xlContinuous

      'Dolares
      g_str_Parame = ""
      g_str_Parame = "USP_RPT_SEGDESG ("
      g_str_Parame = g_str_Parame & "" & r_str_PerMes & ", "
      g_str_Parame = g_str_Parame & "" & r_str_PerAno & ","
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "',"
      g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "',"
      g_str_Parame = g_str_Parame & "'REPORTE DE DESGRAVAMEN',"
      g_str_Parame = g_str_Parame & "'2')"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         MsgBox "Error al ejecutar procedimiento de Desgravamen.", vbCritical, modgen_g_str_NomPlt
         Exit Sub
      End If
      
      g_str_Parame = ""
      g_str_Parame = g_str_Parame + "SELECT * FROM RPT_TABLA_TEMP "
      g_str_Parame = g_str_Parame + " WHERE RPT_PERMES = '" & r_str_PerMes & "' "
      g_str_Parame = g_str_Parame + "   AND RPT_PERANO = '" & r_str_PerAno & "' "
      g_str_Parame = g_str_Parame + "   AND RPT_TERCRE = '" & modgen_g_str_NombPC & "' "
      g_str_Parame = g_str_Parame + "   AND RPT_USUCRE = '" & modgen_g_str_CodUsu & "' "
      g_str_Parame = g_str_Parame + "   AND TRIM(RPT_NOMBRE) = 'REPORTE DE DESGRAVAMEN' "
      g_str_Parame = g_str_Parame + "   AND RPT_MONEDA = '2' "
      g_str_Parame = g_str_Parame + " ORDER BY TO_NUMBER(RPT_CODIGO)"
      
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         Exit Sub
      End If
      
      If g_rst_Princi.BOF And g_rst_Princi.EOF Then
         g_rst_Princi.Close
         Set g_rst_Princi = Nothing
         MsgBox "No se encontraron registros.", vbExclamation, modgen_g_str_NomPlt
         Exit Sub
      End If
   
      r_int_ConTem = 20
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         .Range(.Cells(r_int_ConTem, 1), .Cells(r_int_ConTem, 25)).Font.Name = "Calibri"
         .Range(.Cells(r_int_ConTem, 1), .Cells(r_int_ConTem, 25)).Font.Size = 9
         
         If Val(g_rst_Princi!RPT_CODIGO) >= "3" Then
            .Range(.Cells(r_int_ConTem, 1), .Cells(r_int_ConTem, 1)).NumberFormat = "###,##0.00"
            .Range(.Cells(r_int_ConTem, 2), .Cells(r_int_ConTem, 2)).NumberFormat = "###,##0.00"
            .Range(.Cells(r_int_ConTem, 3), .Cells(r_int_ConTem, 3)).NumberFormat = "###,##0.00"
            .Range(.Cells(r_int_ConTem, 4), .Cells(r_int_ConTem, 4)).NumberFormat = "###,##0.00"
            .Range(.Cells(r_int_ConTem, 5), .Cells(r_int_ConTem, 5)).NumberFormat = "###,##0.00"
            .Range(.Cells(r_int_ConTem, 6), .Cells(r_int_ConTem, 6)).NumberFormat = "###,##0.00"
            .Range(.Cells(r_int_ConTem, 8), .Cells(r_int_ConTem, 8)).NumberFormat = "###,##0.00"
            .Range(.Cells(r_int_ConTem, 9), .Cells(r_int_ConTem, 9)).NumberFormat = "###,##0.00"
            .Range(.Cells(r_int_ConTem, 10), .Cells(r_int_ConTem, 10)).NumberFormat = "###,##0.00"
            .Range(.Cells(r_int_ConTem, 11), .Cells(r_int_ConTem, 11)).NumberFormat = "###,##0.00"
            .Range(.Cells(r_int_ConTem, 12), .Cells(r_int_ConTem, 12)).NumberFormat = "###,##0.00"
            .Range(.Cells(r_int_ConTem, 13), .Cells(r_int_ConTem, 13)).NumberFormat = "###,##0.00"
            .Range(.Cells(r_int_ConTem, 14), .Cells(r_int_ConTem, 14)).NumberFormat = "###,##0.00"
            .Range(.Cells(r_int_ConTem, 15), .Cells(r_int_ConTem, 15)).NumberFormat = "###,##0.00"
         End If

         If Val(g_rst_Princi!RPT_CODIGO) = "3" Then
            .Range(.Cells(r_int_ConTem, 1), .Cells(r_int_ConTem, 15)).Interior.Color = RGB(255, 255, 0)
         End If
         If Val(g_rst_Princi!RPT_CODIGO) = "9" Or Val(g_rst_Princi!RPT_CODIGO) = "10" Then
            .Cells(r_int_ConTem, 1) = "       " & Trim(g_rst_Princi!RPT_DESCRI)
         Else
            .Cells(r_int_ConTem, 1) = Trim(g_rst_Princi!RPT_DESCRI)
         End If
         
         .Cells(r_int_ConTem, 2) = IIf(IsNull(g_rst_Princi!RPT_VALNUM01), 0, g_rst_Princi!RPT_VALNUM01)
         .Cells(r_int_ConTem, 3) = IIf(IsNull(g_rst_Princi!RPT_VALNUM02), 0, g_rst_Princi!RPT_VALNUM02)
         .Cells(r_int_ConTem, 4) = IIf(IsNull(g_rst_Princi!RPT_VALNUM03), 0, g_rst_Princi!RPT_VALNUM03)
         .Cells(r_int_ConTem, 5) = IIf(IsNull(g_rst_Princi!RPT_VALNUM04), 0, g_rst_Princi!RPT_VALNUM04)
         .Cells(r_int_ConTem, 6) = IIf(IsNull(g_rst_Princi!RPT_VALNUM05), 0, g_rst_Princi!RPT_VALNUM05)
         .Cells(r_int_ConTem, 8) = IIf(IsNull(g_rst_Princi!RPT_VALNUM06), 0, g_rst_Princi!RPT_VALNUM06)
         .Cells(r_int_ConTem, 9) = IIf(IsNull(g_rst_Princi!RPT_VALNUM07), 0, g_rst_Princi!RPT_VALNUM07)
         .Cells(r_int_ConTem, 10) = IIf(IsNull(g_rst_Princi!RPT_VALNUM08), 0, g_rst_Princi!RPT_VALNUM08)
         .Cells(r_int_ConTem, 11) = IIf(IsNull(g_rst_Princi!RPT_VALNUM09), 0, g_rst_Princi!RPT_VALNUM09)
         .Cells(r_int_ConTem, 12) = IIf(IsNull(g_rst_Princi!RPT_VALNUM10), 0, g_rst_Princi!RPT_VALNUM10)
         .Cells(r_int_ConTem, 13) = IIf(IsNull(g_rst_Princi!RPT_VALNUM11), 0, g_rst_Princi!RPT_VALNUM11)
         .Cells(r_int_ConTem, 14) = IIf(IsNull(g_rst_Princi!RPT_VALNUM12), 0, g_rst_Princi!RPT_VALNUM12)
         .Cells(r_int_ConTem, 15) = IIf(IsNull(g_rst_Princi!RPT_VALNUM13), 0, g_rst_Princi!RPT_VALNUM13)
         
         .Range(.Cells(r_int_ConTem, 7), .Cells(r_int_ConTem, 7)).Interior.Color = RGB(255, 255, 0)
         
         For r_int_Contad = 1 To 16
            .Range(.Cells(r_int_ConTem, r_int_Contad), .Cells(r_int_ConTem, r_int_Contad)).Borders(xlEdgeLeft).LineStyle = xlContinuous
         Next
         .Range(.Cells(r_int_ConTem, 1), .Cells(r_int_ConTem, 15)).Borders(xlEdgeBottom).LineStyle = xlContinuous

         r_int_ConTem = r_int_ConTem + 1
         g_rst_Princi.MoveNext
      Loop
      
      .Cells(31, 15).Formula = "=SUM(O28:O29)"
   End With

   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub fs_GenExc_SegImb()
Dim r_obj_Excel      As Excel.Application
Dim r_str_Fecha      As String
Dim r_str_FecIni     As String
Dim r_str_FecFin     As String
Dim r_dbl_Monto      As Double
Dim r_dbl_PorIGV     As Double
Dim r_dbl_Factor     As Double
Dim r_int_Contad     As Integer
Dim r_int_TemAux     As Integer
Dim r_int_ConVer     As Integer
Dim r_int_Index      As Integer

   r_str_FecIni = Format(ipp_PerAno.Text, "0000") & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") & "25"
   If cmb_PerMes.ItemData(cmb_PerMes.ListIndex) = 12 Then
      r_str_FecFin = Format(ipp_PerAno.Text + 1, "0000") & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex) + 1, "00") & "02"
   Else
      r_str_FecFin = Format(ipp_PerAno.Text, "0000") & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex) + 1, "00") & "02"
   End If
   
   '-- Obtiene el IGV
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM MNT_PARVAL "
   g_str_Parame = g_str_Parame & " WHERE PARVAL_CODGRP = '001' "
   g_str_Parame = g_str_Parame & "   AND PARVAL_CODITE = '001' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
      Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   r_dbl_PorIGV = Format(g_rst_Genera!PARVAL_CANTID, "###,###,###,##0.000000")
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   
   '-- Prepara excel
   r_str_Fecha = "01/" & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") & "/" & Format(ipp_PerAno.Text, "0000")
   r_dbl_Monto = 0
   
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 2
   r_obj_Excel.Workbooks.Add
   
   For r_int_TemAux = 1 To 2 Step 1
      Select Case r_int_TemAux
         Case 1: r_obj_Excel.Sheets(r_int_TemAux).Name = "Soles"
         Case 2: r_obj_Excel.Sheets(r_int_TemAux).Name = "Dolares"
      End Select
      
      With r_obj_Excel.Sheets(r_int_TemAux)
         .Cells(2, 1) = "MES : " & ObtieneNomMes(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)) & " - " & Format(ipp_PerAno.Text, "0000")
         .Range(.Cells(2, 1), .Cells(2, 15)).HorizontalAlignment = xlHAlignCenter
         .Range(.Cells(2, 1), .Cells(2, 15)).Font.Bold = True
         .Range(.Cells(2, 1), .Cells(2, 15)).Merge
         .Range(.Cells(2, 1), .Cells(2, 15)).Font.Size = 18
         
         .Cells(3, 1) = "MONEDA: " & IIf(r_int_TemAux = 1, "SOLES", "DOLARES")
         .Range(.Cells(3, 1), .Cells(3, 3)).HorizontalAlignment = xlHAlignLeft
         .Range(.Cells(3, 1), .Cells(3, 3)).Font.Bold = True
         .Range(.Cells(3, 1), .Cells(3, 3)).Merge
         
         .Cells(5, 1) = "DECLARACION 0"
         .Range(.Cells(5, 1), .Cells(5, 8)).Merge
         
         .Cells(5, 9) = "CONTRATANTE (1)"
         .Range(.Cells(5, 9), .Cells(5, 34)).Merge
         
         .Cells(5, 35) = "ASEGURADO  (2)"
         .Range(.Cells(5, 35), .Cells(5, 60)).Merge
         
         .Cells(5, 61) = "MATERIA ASEGURADA  (4)"
         .Range(.Cells(5, 61), .Cells(5, 84)).Merge
         
         .Range(.Cells(5, 1), .Cells(5, 84)).HorizontalAlignment = xlHAlignCenter
         .Range(.Cells(5, 1), .Cells(5, 84)).Font.Bold = True
         
         For r_int_Contad = 1 To 84 Step 1
            .Range(.Cells(5, 1), .Cells(5, r_int_Contad)).Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Range(.Cells(5, 1), .Cells(5, r_int_Contad)).Borders(xlEdgeTop).LineStyle = xlContinuous
            .Range(.Cells(5, 1), .Cells(5, r_int_Contad)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Range(.Cells(5, 1), .Cells(5, r_int_Contad)).Borders(xlEdgeRight).LineStyle = xlContinuous
            .Range(.Cells(5, 1), .Cells(5, r_int_Contad)).Borders(xlInsideVertical).LineStyle = xlContinuous
         Next
         
         For r_int_Contad = 1 To 84 Step 1
            .Range(.Cells(6, 1), .Cells(6, r_int_Contad)).Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Range(.Cells(6, 1), .Cells(6, r_int_Contad)).Borders(xlEdgeTop).LineStyle = xlContinuous
            .Range(.Cells(6, 1), .Cells(6, r_int_Contad)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Range(.Cells(6, 1), .Cells(6, r_int_Contad)).Borders(xlEdgeRight).LineStyle = xlContinuous
            .Range(.Cells(6, 1), .Cells(6, r_int_Contad)).Borders(xlInsideVertical).LineStyle = xlContinuous
         Next
         
'         For r_int_Contad = 85 To 85 Step 1
'            .Range(.Cells(7, r_int_Contad), .Cells(7, r_int_Contad)).Borders(xlEdgeLeft).LineStyle = xlContinuous
'            .Range(.Cells(7, r_int_Contad), .Cells(7, r_int_Contad)).Borders(xlEdgeTop).LineStyle = xlContinuous
'            .Range(.Cells(7, r_int_Contad), .Cells(7, r_int_Contad)).Borders(xlEdgeBottom).LineStyle = xlContinuous
'            .Range(.Cells(7, r_int_Contad), .Cells(7, r_int_Contad)).Borders(xlEdgeRight).LineStyle = xlContinuous
'            .Range(.Cells(7, r_int_Contad), .Cells(7, r_int_Contad)).Borders(xlInsideVertical).LineStyle = xlContinuous
'         Next
         
         .Cells(6, 1) = "NRO POLIZA"
         .Cells(6, 2) = "N° DE CERTIFICADO"
         .Cells(6, 3) = "VIGENCIA INICIAL"
         .Cells(6, 4) = "VIGENCIA FINAL"
         .Cells(6, 5) = "SUMA ASEGURADA"
         .Cells(6, 6) = "TASA NETA MENSUAL CIA"
         .Cells(6, 7) = "PRIMA NETA MENSUAL CIA"
         .Cells(6, 8) = "PLAN"
        
         'CONTRATANTE
         .Cells(6, 9) = "TIPO PERSONA"
         .Cells(6, 10) = "TIPO DOCUMENTO"
         .Cells(6, 11) = "N° DE DOCUMENTO"
         .Cells(6, 12) = "AP. PATERNO"
         .Cells(6, 13) = "AP. MATERNO"
         .Cells(6, 14) = "NOMBRES"
         .Cells(6, 15) = "FECHA NACIMIENTO"
         .Cells(6, 16) = "SEXO"
         .Cells(6, 17) = "TIPO VIA DIRECCION"
         .Cells(6, 18) = "DIRECCION"
         .Cells(6, 19) = "DIRECCION NRO"
         .Cells(6, 20) = "DIRECCION KILOMETRO"
         .Cells(6, 21) = "DIRECCION MANZANA"
         .Cells(6, 22) = "DIRECCION LOTE"
         .Cells(6, 23) = "DIRECCION INTERIOR"
         .Cells(6, 24) = "DIRECCION DEPARTAMENTO"
         .Cells(6, 25) = "DIRECCION PISO"
         .Cells(6, 26) = "DIRECCION URBANIZACION"
         .Cells(6, 27) = "DIRECCION REFERENCIA"
         .Cells(6, 28) = "DEPARTAMENTO"
         .Cells(6, 29) = "PROVINCIA"
         .Cells(6, 30) = "DISTRITO"
         .Cells(6, 31) = "TELEFONO"
         .Cells(6, 32) = "CELULAR"
         .Cells(6, 33) = "ESTADO CIVIL"
         .Cells(6, 34) = "CORREO ELECTRONICO"

         'ASEGURADO
         .Cells(6, 35) = "TIPO PERSONA"
         .Cells(6, 36) = "TIPO DOCUMENTO"
         .Cells(6, 37) = "N° DE DOCUMENTO"
         .Cells(6, 38) = "AP. PATERNO"
         .Cells(6, 39) = "AP. MATERNO"
         .Cells(6, 40) = "NOMBRES"
         .Cells(6, 41) = "FECHA NACIMIENTO"
         .Cells(6, 42) = "SEXO"
         .Cells(6, 43) = "TIPO VIA DIRECCION"
         .Cells(6, 44) = "DIRECCION"
         .Cells(6, 45) = "DIRECCION NRO"
         .Cells(6, 46) = "DIRECCION KILOMETRO"
         .Cells(6, 47) = "DIRECCION MANZANA"
         .Cells(6, 48) = "DIRECCION LOTE"
         .Cells(6, 49) = "DIRECCION INTERIOR"
         .Cells(6, 50) = "DIRECCION DEPARTAMENTO"
         .Cells(6, 51) = "DIRECCION PISO"
         .Cells(6, 52) = "DIRECCION URBANIZACION"
         .Cells(6, 53) = "DIRECCION REFERENCIA"
         .Cells(6, 54) = "DEPARTAMENTO"
         .Cells(6, 55) = "PROVINCIA"
         .Cells(6, 56) = "DISTRITO"
         .Cells(6, 57) = "TELEFONO"
         .Cells(6, 58) = "CELULAR"
         .Cells(6, 59) = "ESTADO CIVIL"
         .Cells(6, 60) = "CORREO ELECTRONICO"
         
         'MATERIA ASEGURADA
         .Cells(6, 61) = "TIPO NEGOCIO"
         .Cells(6, 62) = "CAPITAL BASICO"
         .Cells(6, 63) = "TIPO ESTRUCTURA"
         .Cells(6, 64) = "USO DE EDIFICACION"
         .Cells(6, 65) = "TIPO DE EDIFICACION"
         .Cells(6, 66) = "AÑO DE CONSTRUCCION"
         .Cells(6, 67) = "ESTADO DE PREDIO"
         .Cells(6, 68) = "DETALLE"
         .Cells(6, 69) = "TIPO VIA DIRECCION"
         .Cells(6, 70) = "DIRECCION"
         .Cells(6, 71) = "DIRECCION NRO"
         .Cells(6, 72) = "DIRECCION KILOMETRO"
         .Cells(6, 73) = "DIRECCION MANZANA"
         .Cells(6, 74) = "DIRECCION LOTE"
         .Cells(6, 75) = "DIRECCION INTERIOR"
         .Cells(6, 76) = "DIRECCION DEPARTAMENTO"
         .Cells(6, 77) = "DIRECCION PISO"
         .Cells(6, 78) = "DIRECCION URBANIZACION"
         .Cells(6, 79) = "DIRECCION REFERENCIA"
         .Cells(6, 80) = "DEPARTAMENTO"
         .Cells(6, 81) = "PROVINCIA"
         .Cells(6, 82) = "DISTRITO"
         .Cells(6, 83) = "NRO PISOS"
         .Cells(6, 84) = "NRO SOTANOS"
         
        
         .Range(.Cells(6, 1), .Cells(6, 84)).Font.Bold = True
         .Range(.Cells(6, 1), .Cells(6, 84)).HorizontalAlignment = xlHAlignCenter
         
         .Columns("A").ColumnWidth = 18
         .Columns("A").HorizontalAlignment = xlHAlignCenter
         .Columns("B").ColumnWidth = 22
         .Columns("B").HorizontalAlignment = xlHAlignCenter
         .Columns("C").ColumnWidth = 14
         .Columns("C").HorizontalAlignment = xlHAlignCenter
         .Columns("D").ColumnWidth = 14
         .Columns("D").HorizontalAlignment = xlHAlignCenter
         .Columns("E").ColumnWidth = 18
         .Columns("E").HorizontalAlignment = xlHAlignRight
         .Columns("F").ColumnWidth = 14
         .Columns("F").HorizontalAlignment = xlHAlignRight
         .Columns("G").ColumnWidth = 14
         .Columns("G").HorizontalAlignment = xlHAlignRight
         .Columns("H").ColumnWidth = 14
         .Columns("H").HorizontalAlignment = xlHAlignLeft
         .Columns("I").ColumnWidth = 13
         .Columns("I").HorizontalAlignment = xlHAlignCenter
         .Columns("J").ColumnWidth = 13
         .Columns("J").HorizontalAlignment = xlHAlignCenter
         .Columns("K").ColumnWidth = 13
         .Columns("K").HorizontalAlignment = xlHAlignCenter
         .Columns("L").ColumnWidth = 13
         .Columns("L").HorizontalAlignment = xlHAlignCenter
         .Columns("M").ColumnWidth = 13
         .Columns("M").HorizontalAlignment = xlHAlignCenter
         .Columns("N").ColumnWidth = 13
         .Columns("N").HorizontalAlignment = xlHAlignCenter
         .Columns("O").ColumnWidth = 13
         .Columns("O").HorizontalAlignment = xlHAlignLeft
         .Columns("P").ColumnWidth = 13
         .Columns("P").HorizontalAlignment = xlHAlignCenter
         .Columns("Q").ColumnWidth = 13
         .Columns("Q").HorizontalAlignment = xlHAlignCenter
         .Columns("R").ColumnWidth = 55
         .Columns("R").HorizontalAlignment = xlHAlignCenter
         .Columns("S").ColumnWidth = 13
         .Columns("S").HorizontalAlignment = xlHAlignCenter
         .Columns("T").ColumnWidth = 13
         .Columns("T").HorizontalAlignment = xlHAlignLeft
         .Columns("U").ColumnWidth = 13
         .Columns("U").HorizontalAlignment = xlHAlignCenter
         .Columns("V").ColumnWidth = 13
         .Columns("V").HorizontalAlignment = xlHAlignCenter
         .Columns("W").ColumnWidth = 13
         .Columns("W").HorizontalAlignment = xlHAlignCenter
         .Columns("X").ColumnWidth = 16
         .Columns("X").HorizontalAlignment = xlHAlignCenter
         .Columns("Y").ColumnWidth = 13
         .Columns("Y").HorizontalAlignment = xlHAlignCenter
         .Columns("Z").ColumnWidth = 14
         .Columns("Z").HorizontalAlignment = xlHAlignCenter
         .Columns("AA").ColumnWidth = 13
         .Columns("AA").HorizontalAlignment = xlHAlignCenter
         
         .Columns("AB").ColumnWidth = 15
         .Columns("AB").HorizontalAlignment = xlHAlignCenter
         .Columns("AC").ColumnWidth = 13
         .Columns("AC").HorizontalAlignment = xlHAlignCenter
         .Columns("AD").ColumnWidth = 13
         .Columns("AD").HorizontalAlignment = xlHAlignCenter
         .Columns("AE").ColumnWidth = 13
         .Columns("AE").HorizontalAlignment = xlHAlignCenter
         .Columns("AF").ColumnWidth = 13
         .Columns("AF").HorizontalAlignment = xlHAlignCenter
         .Columns("AG").ColumnWidth = 13
         .Columns("AG").HorizontalAlignment = xlHAlignCenter
         .Columns("AH").ColumnWidth = 13
         .Columns("AH").HorizontalAlignment = xlHAlignCenter
         .Columns("AI").ColumnWidth = 11
         .Columns("AI").HorizontalAlignment = xlHAlignCenter
         .Columns("AJ").ColumnWidth = 13
         .Columns("AJ").HorizontalAlignment = xlHAlignCenter
         .Columns("AK").ColumnWidth = 14
         .Columns("AK").HorizontalAlignment = xlHAlignCenter
         .Columns("AL").ColumnWidth = 18
         .Columns("AL").HorizontalAlignment = xlHAlignLeft
         .Columns("AM").ColumnWidth = 18
         .Columns("AM").HorizontalAlignment = xlHAlignLeft
         .Columns("AN").ColumnWidth = 18
         .Columns("AN").HorizontalAlignment = xlHAlignLeft
         .Columns("AO").ColumnWidth = 14
         .Columns("AO").HorizontalAlignment = xlHAlignCenter
         .Columns("AP").ColumnWidth = 13
         .Columns("AP").HorizontalAlignment = xlHAlignCenter
         
         .Columns("AQ").ColumnWidth = 13
         .Columns("AR").ColumnWidth = 100
         .Columns("AR").HorizontalAlignment = xlHAlignLeft
         
         .Columns("AS").ColumnWidth = 13
         .Columns("AT").ColumnWidth = 13
         .Columns("AU").ColumnWidth = 13
         .Columns("AV").ColumnWidth = 13
         .Columns("AW").ColumnWidth = 13
         .Columns("AX").ColumnWidth = 16
         .Columns("AY").ColumnWidth = 13
         .Columns("AZ").ColumnWidth = 15
         .Columns("BA").ColumnWidth = 13
         
         .Columns("BB").ColumnWidth = 16
         .Columns("BB").HorizontalAlignment = xlHAlignCenter
         .Columns("BC").ColumnWidth = 13
         .Columns("BC").HorizontalAlignment = xlHAlignCenter
         .Columns("BD").ColumnWidth = 13
         .Columns("BD").HorizontalAlignment = xlHAlignCenter
         .Columns("BE").ColumnWidth = 13
         .Columns("BE").HorizontalAlignment = xlHAlignCenter
         .Columns("BF").ColumnWidth = 13
         .Columns("BF").HorizontalAlignment = xlHAlignCenter
         .Columns("BG").ColumnWidth = 13
         .Columns("BG").HorizontalAlignment = xlHAlignCenter
         
         .Columns("BH").ColumnWidth = 47
         .Columns("BI").ColumnWidth = 15
         .Columns("BI").HorizontalAlignment = xlHAlignCenter
         .Columns("BJ").ColumnWidth = 13
         .Columns("BK").ColumnWidth = 61
         .Columns("BL").ColumnWidth = 13
         .Columns("BL").HorizontalAlignment = xlHAlignCenter
         .Columns("BM").ColumnWidth = 20
         .Columns("BN").ColumnWidth = 15
         .Columns("BN").HorizontalAlignment = xlHAlignCenter
         .Columns("BO").ColumnWidth = 13
         .Columns("BP").ColumnWidth = 13
         .Columns("BQ").ColumnWidth = 13
         .Columns("BR").ColumnWidth = 100
         .Columns("BS").ColumnWidth = 13
         .Columns("BT").ColumnWidth = 13
         .Columns("BU").ColumnWidth = 13
         .Columns("BV").ColumnWidth = 13
         .Columns("BW").ColumnWidth = 13
         .Columns("BX").ColumnWidth = 16
         .Columns("BY").ColumnWidth = 13
         .Columns("BZ").ColumnWidth = 15
         
         .Columns("CA").ColumnWidth = 13
         .Columns("CB").ColumnWidth = 16
         .Columns("CB").HorizontalAlignment = xlHAlignCenter
         .Columns("CC").ColumnWidth = 13
         .Columns("CC").HorizontalAlignment = xlHAlignCenter
         .Columns("CD").ColumnWidth = 13
         .Columns("CD").HorizontalAlignment = xlHAlignCenter
         .Columns("CE").ColumnWidth = 13
         .Columns("CE").HorizontalAlignment = xlHAlignCenter
         .Columns("CF").ColumnWidth = 13
         .Columns("CF").HorizontalAlignment = xlHAlignCenter
         
         For r_int_Index = 1 To 84 Step 1
           .Range(.Cells(6, r_int_Index), .Cells(6, r_int_Index)).WrapText = True
           .Range(.Cells(6, r_int_Index), .Cells(6, r_int_Index)).VerticalAlignment = xlCenter
           .Range(.Cells(6, r_int_Index), .Cells(6, r_int_Index)).HorizontalAlignment = xlHAlignCenter
         Next
         
         r_int_ConVer = 7
'         r_dbl_Factor = CDbl(txtFactor.Text) * ((r_dbl_PorIGV / 100) + 1)
'         .Cells(r_int_ConVer, 85) = Format(r_dbl_Factor, "##0.000000") '39
'         .Cells(r_int_ConVer, 86) = "Factor Actual" '40
'         .Range(.Cells(r_int_ConVer, 85), .Cells(r_int_ConVer, 86)).Font.Bold = True
'
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & "SELECT "
         g_str_Parame = g_str_Parame & "       DECODE(HIPCIE_TIPMON, 1, '60031833', '60031696')                         AS POLIZA, "
         g_str_Parame = g_str_Parame & "       TRIM(POLIZA_NUMVIV)                                                      AS NOCERTIFICADO, "
         g_str_Parame = g_str_Parame & "       POLIZA_FEMVIV                                                            AS FECHAAFILIACION, "
         g_str_Parame = g_str_Parame & "       HIPMAE_ULTVCT                                                            AS FECHA_VENCIMIENTO, "
         g_str_Parame = g_str_Parame & "       EVATAS_SUMASE_INM+EVATAS_SUMASE_ES1+EVATAS_SUMASE_ES2+EVATAS_SUMASE_DEP  AS SUMAASEGURADA,"
         g_str_Parame = g_str_Parame & "       ROUND(HIPCIE_FOIVIV, 5)                                                  AS TASA, "
         g_str_Parame = g_str_Parame & "       ROUND(((HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_FOIVIV)/100,2)               AS PRIMENETA, "
         g_str_Parame = g_str_Parame & "       ''                                                                       AS PLAN, "
         g_str_Parame = g_str_Parame & "       'J'                                                                      AS TIPPER_CONTRATANTE, "
         g_str_Parame = g_str_Parame & "       '6'                                                                      AS TIPDOC_CONTRATANTE, "
         g_str_Parame = g_str_Parame & "       (SELECT X.EMPGRP_NUMRUC"
         g_str_Parame = g_str_Parame & "          FROM MNT_EMPGRP X )                                                   AS NUMDOC_CONTRATANTE, "
         g_str_Parame = g_str_Parame & "       '' AS APEPAT_CONTRATANTE,"
         g_str_Parame = g_str_Parame & "       '' AS APEMAT_CONTRATANTE,"
         g_str_Parame = g_str_Parame & "       (SELECT X.EMPGRP_RAZSOC"
         g_str_Parame = g_str_Parame & "          FROM MNT_EMPGRP X )                                                   AS NOMBRE_CONTRATANTE, "
         g_str_Parame = g_str_Parame & "       (SELECT TRIM(Y.PARDES_DESCRI) || ' ' || TRIM(X.EMPGRP_NOMVIA) || ' ' || TRIM(X.EMPGRP_NUMVIA) || ' ' || TRIM(X.EMPGRP_INTDPT) || ' ' || TRIM(Z.PARDES_DESCRI) || ' ' || TRIM(X.EMPGRP_NOMZON)"
         g_str_Parame = g_str_Parame & "          FROM MNT_EMPGRP X"
         g_str_Parame = g_str_Parame & "               LEFT JOIN MNT_PARDES Y ON (EMPGRP_TIPVIA = Y.PARDES_CODITE AND  Y.PARDES_CODGRP = 201 )"
         g_str_Parame = g_str_Parame & "               LEFT JOIN MNT_PARDES Z ON (EMPGRP_TIPZON = Z.PARDES_CODITE AND  Z.PARDES_CODGRP = 202 )"
         g_str_Parame = g_str_Parame & "          )  AS DIREC_CONTRATANTE,"
         g_str_Parame = g_str_Parame & "       (SELECT SUBSTR(X.EMPGRP_UBIGEO,1,2)"
         g_str_Parame = g_str_Parame & "          FROM MNT_EMPGRP X )                                                   AS DEPART_CONTRATANTE, "
         g_str_Parame = g_str_Parame & "       (SELECT SUBSTR(X.EMPGRP_UBIGEO,3,2)"
         g_str_Parame = g_str_Parame & "          FROM MNT_EMPGRP X )                                                   AS PROVIN_CONTRATANTE, "
         g_str_Parame = g_str_Parame & "       (SELECT SUBSTR(X.EMPGRP_UBIGEO,5,2)"
         g_str_Parame = g_str_Parame & "          FROM MNT_EMPGRP X )                                                   AS DISTRI_CONTRATANTE, "
         g_str_Parame = g_str_Parame & "       (SELECT X.EMPGRP_TELEF1"
         g_str_Parame = g_str_Parame & "          FROM MNT_EMPGRP X )                                                   AS TELEFO_CONTRATANTE, "
         g_str_Parame = g_str_Parame & "       '1' AS ESTCIV_CONTRATANTE,"
         g_str_Parame = g_str_Parame & "       (CASE WHEN TRIM(HIPMAE_TDOCLI) = 1 OR TRIM(HIPMAE_TDOCLI) = 4 OR TRIM(HIPMAE_TDOCLI) = 7 THEN 'N'"
         g_str_Parame = g_str_Parame & "        ELSE CASE WHEN TRIM(HIPMAE_TDOCLI) = 6 THEN 'J' END"
         g_str_Parame = g_str_Parame & "         END)                                                                   AS TIPPER_ASEGURADO, "
         g_str_Parame = g_str_Parame & "       TRIM(HIPCIE_TDOCLI)                                                      AS TIPDOC_ASEGURADO, "
         g_str_Parame = g_str_Parame & "       TRIM(HIPCIE_NDOCLI)                                                      AS NUMDOC_ASEGURADO, "
         g_str_Parame = g_str_Parame & "       TRIM(DATGEN_APEPAT)                                                      AS APEPAT_ASEGURADO,"
         g_str_Parame = g_str_Parame & "       TRIM(DATGEN_APEMAT)                                                      AS APEMAT_ASEGURADO, "
         g_str_Parame = g_str_Parame & "       TRIM(DATGEN_NOMBRE)                                                      AS NOMCLI_ASEGURADO, "
         g_str_Parame = g_str_Parame & "       DATGEN_NACFEC                                                            AS FECNAC_ASEGURADO, "
         g_str_Parame = g_str_Parame & "       DECODE(DATGEN_CODSEX,1,'M','F')                                          AS SEXO_ASEGURADO, "
         g_str_Parame = g_str_Parame & "       TRIM(DECODE(DATGEN_TIPVIA, 12, '', TRIM(H.PARDES_DESCRI)))                 AS TIPVIA_ASEGURADO, "
         g_str_Parame = g_str_Parame & "       TRIM(TRIM(DATGEN_NOMVIA)||' '||TRIM(DATGEN_NUMERO)||' '||DECODE(NVL(LENGTH(TRIM(DATGEN_INTDPT)), 0), 0, '', '('||TRIM(DATGEN_INTDPT)||')')||' '||DECODE(NVL(LENGTH(TRIM(DATGEN_NOMZON)),0), 0, '', ' - '||DECODE(DATGEN_TIPZON, 12, '', TRIM(I.PARDES_DESCRI))||' '||TRIM(DATGEN_NOMZON))) AS DIRECC_ASEGURADO, "
         g_str_Parame = g_str_Parame & "       DECODE(DATGEN_TIPZON, 12, '', TRIM(I.PARDES_DESCRI)) || ' ' || TRIM(DATGEN_NOMZON) AS TIPO_DEPT_INT_ASEGURADO, "
         g_str_Parame = g_str_Parame & "       TRIM(DATGEN_REFERE)                                                      AS REFERE_ASEGURADO, "
         g_str_Parame = g_str_Parame & "       SUBSTR(DATGEN_UBIGEO,1,2)                                                AS DEPART_ASEGURADO, "
         g_str_Parame = g_str_Parame & "       SUBSTR(DATGEN_UBIGEO,3,2)                                                AS PROVIN_ASEGURADO, "
         g_str_Parame = g_str_Parame & "       SUBSTR(DATGEN_UBIGEO,5,2)                                                AS DISTRT_ASEGURADO, "
         g_str_Parame = g_str_Parame & "       DATGEN_UBIGEO,"
         g_str_Parame = g_str_Parame & "       TRIM(DATGEN_TELEFO)                                                      AS TELFIJ_ASEGURADO, "
         g_str_Parame = g_str_Parame & "       TRIM(DATGEN_NUMCEL)                                                      AS CELULA_ASEGURADO, "
         g_str_Parame = g_str_Parame & "       TRIM(DATGEN_ESTCIV) || '-' || TRIM(EC.PARDES_DESCRI)                     AS ESTCIV_ASEGURADO, "
         g_str_Parame = g_str_Parame & "       TRIM(DATGEN_DIRELE)                                                      AS CORREO_ASEGURADO, "
         g_str_Parame = g_str_Parame & "       ''                                                                       AS TIPO_NEGOCIO, "      'TRIM(NM.PARDES_DESCRI)
         g_str_Parame = g_str_Parame & "       ''                                                                       AS CAPITAL_BASICO, "
         g_str_Parame = g_str_Parame & "       TRIM(MC.PARDES_DESCRI)                                                   AS TIPO_ESTRUCTURA, "
         g_str_Parame = g_str_Parame & "       TRIM(UI.PARDES_DESCRI)                                                   AS USO_INMUEBLE, "
         g_str_Parame = g_str_Parame & "       TRIM(TI.PARDES_DESCRI)                                                   AS TIPO_EDIFICACION, "
         g_str_Parame = g_str_Parame & "       EVATAS_ANOCON                                                            AS ANO_CONSTRUCCION, "
         g_str_Parame = g_str_Parame & "       'BUENO'                                                                  AS ESTADO_PREDIO, "
         g_str_Parame = g_str_Parame & "       ''                                                                       AS DETALLE, "
         g_str_Parame = g_str_Parame & "       TRIM(IV.PARDES_DESCRI)                                                   AS TIPOVIA_MATERIA ,"
         g_str_Parame = g_str_Parame & "       TRIM(SOLINM_NOMVIA) ||' '|| TRIM(SOLINM_NUMVIA) ||' '|| TRIM(SOLINM_INTDPT) ||' '|| TRIM(ZN.PARDES_DESCRI) ||' '|| TRIM(SOLINM_NOMZON) AS DIRECCION_MATERIA,"
         g_str_Parame = g_str_Parame & "       SUBSTR(SOLINM_UBIGEO,1,2)                                                AS DEPART_MATERIA, "
         g_str_Parame = g_str_Parame & "       SUBSTR(SOLINM_UBIGEO,3,2)                                                AS PROVIN_MATERIA, "
         g_str_Parame = g_str_Parame & "       SUBSTR(SOLINM_UBIGEO,5,2)                                                AS DISTRT_MATERIA, "
         g_str_Parame = g_str_Parame & "       SOLINM_UBIGEO,"
         g_str_Parame = g_str_Parame & "       EVATAS_NUMPIS                                                            AS NUMERO_PISOS, "
         g_str_Parame = g_str_Parame & "       EVATAS_NUMSOT                                                            AS NUMERO_SOTANOS "
         g_str_Parame = g_str_Parame & "  FROM CRE_HIPCIE "
         g_str_Parame = g_str_Parame & " INNER JOIN CLI_DATGEN    ON (DATGEN_TIPDOC = HIPCIE_TDOCLI AND DATGEN_NUMDOC = HIPCIE_NDOCLI) "
         g_str_Parame = g_str_Parame & " INNER JOIN CRE_HIPMAE    ON (HIPMAE_NUMOPE = HIPCIE_NUMOPE) "
         g_str_Parame = g_str_Parame & " INNER JOIN TRA_POLIZA    ON (POLIZA_NUMSOL = HIPMAE_NUMSOL) "
         g_str_Parame = g_str_Parame & " INNER JOIN TRA_EVASEG    ON (EVASEG_NUMSOL = HIPMAE_NUMSOL) "
         g_str_Parame = g_str_Parame & " INNER JOIN TRA_EVATAS    ON (EVATAS_NUMSOL = HIPMAE_NUMSOL) "
         g_str_Parame = g_str_Parame & " INNER JOIN MNT_SEGEMP    ON (SEGEMP_CODIGO = EVASEG_ESGDES) "
         g_str_Parame = g_str_Parame & " INNER JOIN MNT_SEGTIP    ON (SEGTIP_CODIGO = EVASEG_ESGDES AND SEGTIP_TIPSEG = EVASEG_TIPSEG) "
         g_str_Parame = g_str_Parame & " INNER JOIN CRE_SOLINM    ON (SOLINM_NUMSOL = HIPMAE_NUMSOL) "
         g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES II ON (DATGEN_UBIGEO = II.PARDES_CODITE AND  II.PARDES_CODGRP = 101 ) "
         g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES IV ON (DATGEN_TIPVIA = IV.PARDES_CODITE AND  IV.PARDES_CODGRP = 201 ) "
         g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES ZN ON (SOLINM_TIPZON = ZN.PARDES_CODITE AND  ZN.PARDES_CODGRP = 202 ) "
         g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES IT ON (IT.PARDES_CODGRP = 101 AND RPAD(SUBSTR(SOLINM_UBIGEO,1,2),6,0)= IT.PARDES_CODITE) "
         g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES IP ON (IP.PARDES_CODGRP = 101 AND RPAD(SUBSTR(SOLINM_UBIGEO,1,4),6,0)= IP.PARDES_CODITE) "
         g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES NM ON (NM.PARDES_CODGRP = 217 AND SOLINM_TIPINM = NM.PARDES_CODITE) "
         g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES EC ON (EC.PARDES_CODGRP = '205' AND EC.PARDES_CODITE = DATGEN_ESTCIV) "
         g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES TD ON (TD.PARDES_CODGRP = '203' AND TD.PARDES_CODITE = HIPCIE_TDOCLI) "
         g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES GR ON (GR.PARDES_CODGRP = '241' AND GR.PARDES_CODITE = HIPCIE_TIPGAR) "
         g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES UI ON (UI.PARDES_CODGRP = '222' AND UI.PARDES_CODITE = EVATAS_USOINM) "
         g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES TI ON (TI.PARDES_CODGRP = '221' AND TI.PARDES_CODITE = EVATAS_TIPINM) "
         g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES MC ON (MC.PARDES_CODGRP = '223' AND MC.PARDES_CODITE = EVATAS_MATCON) "
         g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES MD ON (MD.PARDES_CODGRP = '204' AND MD.PARDES_CODITE = HIPCIE_TIPMON) "
         g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES H  ON H.PARDES_CODGRP = 201 AND H.PARDES_CODITE = DATGEN_TIPVIA "
         g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES I  ON I.PARDES_CODGRP = 202 AND I.PARDES_CODITE = DATGEN_TIPZON "
         g_str_Parame = g_str_Parame & " WHERE HIPCIE_PERMES = " & cmb_PerMes.ItemData(cmb_PerMes.ListIndex) & " "
         g_str_Parame = g_str_Parame & "   AND HIPCIE_PERANO = " & Format(ipp_PerAno.Text, "0000") & " "
         g_str_Parame = g_str_Parame & "   AND HIPCIE_TIPMON = " & r_int_TemAux & " "
         
         g_str_Parame = g_str_Parame & "UNION "
         
         g_str_Parame = g_str_Parame & "SELECT "
         g_str_Parame = g_str_Parame & "       DECODE(HIPMAE_MONEDA, 1, '60031833', '60031696')                         AS POLIZA, "
         g_str_Parame = g_str_Parame & "       TRIM(POLIZA_NUMVIV)                                                      AS NOCERTIFICADO, "
         g_str_Parame = g_str_Parame & "       POLIZA_FEMVIV                                                            AS FECHAAFILIACION, "
         g_str_Parame = g_str_Parame & "       HIPMAE_ULTVCT                                                            AS FECHA_VENCIMIENTO, "
         g_str_Parame = g_str_Parame & "       EVATAS_SUMASE_INM+EVATAS_SUMASE_ES1+EVATAS_SUMASE_ES2+EVATAS_SUMASE_DEP  AS SUMAASEGURADA, "
         g_str_Parame = g_str_Parame & "       ROUND(HIPMAE_FOIVIV, 5)                                                  AS TASA, "
         g_str_Parame = g_str_Parame & "       ROUND(((HIPMAE_SALCAP+HIPMAE_SALCON)*HIPMAE_FOIVIV)/100,2)               AS PRIMENETA, "
         g_str_Parame = g_str_Parame & "        ''                                                                      AS PLAN, "
         g_str_Parame = g_str_Parame & "       'J'                                                                      AS TIPPER_CONTRATANTE, "
         g_str_Parame = g_str_Parame & "       '6'                                                                      AS TIPDOC_CONTRATANTE, "
         g_str_Parame = g_str_Parame & "       (SELECT X.EMPGRP_NUMRUC"
         g_str_Parame = g_str_Parame & "          FROM MNT_EMPGRP X )                                                   AS NUMDOC_CONTRATANTE, "
         g_str_Parame = g_str_Parame & "       ''                                                                       AS APEPAT_CONTRATANTE, "
         g_str_Parame = g_str_Parame & "       ''                                                                       AS APEMAT_CONTRATANTE, "
         g_str_Parame = g_str_Parame & "       (SELECT X.EMPGRP_RAZSOC"
         g_str_Parame = g_str_Parame & "          FROM MNT_EMPGRP X )                                                   AS NOMBRE_CONTRATANTE, "
       
         g_str_Parame = g_str_Parame & "       (SELECT TRIM(Y.PARDES_DESCRI) || ' ' || TRIM(X.EMPGRP_NOMVIA) || ' ' || TRIM(X.EMPGRP_NUMVIA) || ' ' || TRIM(X.EMPGRP_INTDPT) || ' ' || TRIM(Z.PARDES_DESCRI) || ' ' || TRIM(X.EMPGRP_NOMZON)"
         g_str_Parame = g_str_Parame & "          FROM MNT_EMPGRP X"
         g_str_Parame = g_str_Parame & "          LEFT JOIN MNT_PARDES Y ON (EMPGRP_TIPVIA = Y.PARDES_CODITE AND  Y.PARDES_CODGRP = 201 )"
         g_str_Parame = g_str_Parame & "          LEFT JOIN MNT_PARDES Z ON (EMPGRP_TIPZON = Z.PARDES_CODITE AND  Z.PARDES_CODGRP = 202 )"
         g_str_Parame = g_str_Parame & "          )  AS DIREC_CONTRATANTE,"
                   
         g_str_Parame = g_str_Parame & "       (SELECT SUBSTR(X.EMPGRP_UBIGEO,1,2)"
         g_str_Parame = g_str_Parame & "          FROM MNT_EMPGRP X )                                                   AS DEPART_CONTRATANTE, "
         g_str_Parame = g_str_Parame & "       (SELECT SUBSTR(X.EMPGRP_UBIGEO,3,2)"
         g_str_Parame = g_str_Parame & "          FROM MNT_EMPGRP X )                                                   AS PROVIN_CONTRATANTE, "
         g_str_Parame = g_str_Parame & "       (SELECT SUBSTR(X.EMPGRP_UBIGEO,5,2)"
         g_str_Parame = g_str_Parame & "          FROM MNT_EMPGRP X )                                                   AS DISTRI_CONTRATANTE, "
         g_str_Parame = g_str_Parame & "       (SELECT X.EMPGRP_TELEF1"
         g_str_Parame = g_str_Parame & "          FROM MNT_EMPGRP X )                                                   AS TELEFO_CONTRATANTE, "
         g_str_Parame = g_str_Parame & "       '1' AS ESTCIV_CONTRATANTE,"
         g_str_Parame = g_str_Parame & "       (CASE WHEN TRIM(HIPMAE_TDOCLI) = 1 OR TRIM(HIPMAE_TDOCLI) = 4 OR TRIM(HIPMAE_TDOCLI) = 7 THEN 'N'"
         g_str_Parame = g_str_Parame & "        ELSE CASE WHEN TRIM(HIPMAE_TDOCLI) = 6 THEN 'J' END"
         g_str_Parame = g_str_Parame & "         END)                                                                   AS TIPPER_ASEGURADO, "
         g_str_Parame = g_str_Parame & "       TRIM(HIPMAE_TDOCLI)                                                      AS TIPDOC_ASEGURADO, "
         g_str_Parame = g_str_Parame & "       TRIM(HIPMAE_NDOCLI)                                                      AS NUMDOC_ASEGURADO, "
         g_str_Parame = g_str_Parame & "       TRIM(DATGEN_APEPAT)                                                      AS APEPAT_ASEGURADO, "
         g_str_Parame = g_str_Parame & "       TRIM(DATGEN_APEMAT)                                                      AS APEMAT_ASEGURADO, "
         g_str_Parame = g_str_Parame & "       TRIM(DATGEN_NOMBRE)                                                      AS NOMCLI_ASEGURADO, "
         g_str_Parame = g_str_Parame & "       DATGEN_NACFEC                                                            AS FECNAC_ASEGURADO, "
         g_str_Parame = g_str_Parame & "       DECODE(DATGEN_CODSEX,1,'M','F')                                          AS SEXO_ASEGURADO, "
         g_str_Parame = g_str_Parame & "       TRIM(DECODE(DATGEN_TIPVIA, 12, '', TRIM(H.PARDES_DESCRI)))               AS TIPVIA_ASEGURADO, "
         g_str_Parame = g_str_Parame & "       TRIM(TRIM(DATGEN_NOMVIA)||' '||TRIM(DATGEN_NUMERO)||' '||DECODE(NVL(LENGTH(TRIM(DATGEN_INTDPT)), 0), 0, '', '('||TRIM(DATGEN_INTDPT)||')')||' '||DECODE(NVL(LENGTH(TRIM(DATGEN_NOMZON)),0), 0, '', ' - '||DECODE(DATGEN_TIPZON, 12, '', TRIM(I.PARDES_DESCRI))||' '||TRIM(DATGEN_NOMZON))) AS DIRECCION_ASEGURADO, "
         g_str_Parame = g_str_Parame & "       DECODE(DATGEN_TIPZON, 12, '', TRIM(I.PARDES_DESCRI)) || ' ' || TRIM(DATGEN_NOMZON) AS TIPO_DEPT_INT_ASEGURADO, "
         g_str_Parame = g_str_Parame & "       TRIM(DATGEN_REFERE)                                                      AS REFERENCIA_ASEGURADO, "
         g_str_Parame = g_str_Parame & "       SUBSTR(DATGEN_UBIGEO,1,2)                                                AS DEPART_ASEGURADO, "
         g_str_Parame = g_str_Parame & "       SUBSTR(DATGEN_UBIGEO,3,2)                                                AS PROVIN_ASEGURADO, "
         g_str_Parame = g_str_Parame & "       SUBSTR(DATGEN_UBIGEO,5,2)                                                AS DISTRT_ASEGURADO, "
         g_str_Parame = g_str_Parame & "       DATGEN_UBIGEO,"
         g_str_Parame = g_str_Parame & "       TRIM(DATGEN_TELEFO)                                                      AS TELFIJ_ASEGURADO, "
         g_str_Parame = g_str_Parame & "       TRIM(DATGEN_NUMCEL)                                                      AS CELULA_ASEGURADO, "
         g_str_Parame = g_str_Parame & "       TRIM(DATGEN_ESTCIV) || '-' || TRIM(EC.PARDES_DESCRI)                     AS ESTCIV_ASEGURADO, "
         g_str_Parame = g_str_Parame & "       TRIM(DATGEN_DIRELE)                                                      AS CORREO_ASEGURADO, "
         g_str_Parame = g_str_Parame & "       ''                                                                       AS TIPO_NEGOCIO, " 'TRIM(NM.PARDES_DESCRI)
         g_str_Parame = g_str_Parame & "       ''                                                                       AS CAPITAL_BASICO, "
         g_str_Parame = g_str_Parame & "       TRIM(MC.PARDES_DESCRI)                                                   AS TIPO_ESTRUCTURA, "
         g_str_Parame = g_str_Parame & "       TRIM(UI.PARDES_DESCRI)                                                   AS USO_INMUEBLE, "
         g_str_Parame = g_str_Parame & "       TRIM(TI.PARDES_DESCRI)                                                   AS TIPO_EDIFICACION, "
         g_str_Parame = g_str_Parame & "       EVATAS_ANOCON                                                            AS ANO_CONSTRUCCION, "
         g_str_Parame = g_str_Parame & "       'BUENO'                                                                  AS ESTADO_PREDIO, "
         g_str_Parame = g_str_Parame & "       ''                                                                       AS DETALLE, "
         g_str_Parame = g_str_Parame & "       TRIM(IV.PARDES_DESCRI)                                                   AS TIPOVIA_MATERIA ,"
         g_str_Parame = g_str_Parame & "       TRIM(SOLINM_NOMVIA) ||' '|| TRIM(SOLINM_NUMVIA) ||' '|| TRIM(SOLINM_INTDPT) ||' '|| TRIM(ZN.PARDES_DESCRI) ||' '|| TRIM(SOLINM_NOMZON) AS DIRECCION_MATERIA,"
         g_str_Parame = g_str_Parame & "       SUBSTR(SOLINM_UBIGEO,1,2)                                                AS DEPART_MATERIA, "
         g_str_Parame = g_str_Parame & "       SUBSTR(SOLINM_UBIGEO,3,2)                                                AS PROVIN_MATERIA, "
         g_str_Parame = g_str_Parame & "       SUBSTR(SOLINM_UBIGEO,5,2)                                                AS DISTRT_MATERIA, "
         g_str_Parame = g_str_Parame & "       SOLINM_UBIGEO,"
         g_str_Parame = g_str_Parame & "       EVATAS_NUMPIS                                                            AS NUMERO_PISOS, "
         g_str_Parame = g_str_Parame & "       EVATAS_NUMSOT                                                            As NUMERO_SOTANOS "
         g_str_Parame = g_str_Parame & "  FROM CRE_HIPMAE "
         g_str_Parame = g_str_Parame & " INNER JOIN CLI_DATGEN    ON (DATGEN_TIPDOC = HIPMAE_TDOCLI AND DATGEN_NUMDOC = HIPMAE_NDOCLI) "
         g_str_Parame = g_str_Parame & " INNER JOIN TRA_POLIZA    ON (POLIZA_NUMSOL = HIPMAE_NUMSOL) "
         g_str_Parame = g_str_Parame & " INNER JOIN TRA_EVASEG    ON (EVASEG_NUMSOL = HIPMAE_NUMSOL) "
         g_str_Parame = g_str_Parame & " INNER JOIN TRA_EVATAS    ON (EVATAS_NUMSOL = HIPMAE_NUMSOL) "
         g_str_Parame = g_str_Parame & " INNER JOIN MNT_SEGEMP    ON (SEGEMP_CODIGO = EVASEG_ESGDES) "
         g_str_Parame = g_str_Parame & " INNER JOIN MNT_SEGTIP    ON (SEGTIP_CODIGO = EVASEG_ESGDES AND SEGTIP_TIPSEG = EVASEG_TIPSEG) "
         g_str_Parame = g_str_Parame & " INNER JOIN CRE_SOLINM    ON (SOLINM_NUMSOL = HIPMAE_NUMSOL) "
         g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES II ON (DATGEN_UBIGEO = II.PARDES_CODITE AND  II.PARDES_CODGRP = 101 ) "
         g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES IV ON (DATGEN_TIPVIA = IV.PARDES_CODITE AND  IV.PARDES_CODGRP = 201 ) "
         g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES ZN ON (SOLINM_TIPZON = ZN.PARDES_CODITE AND  ZN.PARDES_CODGRP = 202 ) "
         g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES IT ON (IT.PARDES_CODGRP = 101 AND RPAD(SUBSTR(SOLINM_UBIGEO,1,2),6,0)= IT.PARDES_CODITE) "
         g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES IP ON (IP.PARDES_CODGRP = 101 AND RPAD(SUBSTR(SOLINM_UBIGEO,1,4),6,0)= IP.PARDES_CODITE) "
         g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES NM ON (NM.PARDES_CODGRP = 217 AND SOLINM_TIPINM = NM.PARDES_CODITE) "
         g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES EC ON (EC.PARDES_CODGRP = '205' AND EC.PARDES_CODITE = DATGEN_ESTCIV) "
         g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES TD ON (TD.PARDES_CODGRP = '203' AND TD.PARDES_CODITE = HIPMAE_TDOCLI) "
         g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES GR ON (GR.PARDES_CODGRP = '241' AND GR.PARDES_CODITE = HIPMAE_TIPGAR) "
         g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES UI ON (UI.PARDES_CODGRP = '222' AND UI.PARDES_CODITE = EVATAS_USOINM) "
         g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES TI ON (TI.PARDES_CODGRP = '221' AND TI.PARDES_CODITE = EVATAS_TIPINM) "
         g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES MC ON (MC.PARDES_CODGRP = '223' AND MC.PARDES_CODITE = EVATAS_MATCON) "
         g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES MD ON (MD.PARDES_CODGRP = '204' AND MD.PARDES_CODITE = HIPMAE_MONEDA) "
         g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES H  ON H.PARDES_CODGRP = 201 AND H.PARDES_CODITE = DATGEN_TIPVIA "
         g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES I  ON I.PARDES_CODGRP = 202 AND I.PARDES_CODITE = DATGEN_TIPZON "
         g_str_Parame = g_str_Parame & " WHERE HIPMAE_SITUAC = 6 "
         g_str_Parame = g_str_Parame & "   AND HIPMAE_MONEDA = " & r_int_TemAux & " "
         g_str_Parame = g_str_Parame & "   AND HIPMAE_FECCAN >= " & r_str_FecIni
         g_str_Parame = g_str_Parame & "   AND HIPMAE_FECCAN <= " & r_str_FecFin
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
            Exit Sub
         End If
         
         If g_rst_Princi.BOF And g_rst_Princi.EOF Then
            MsgBox "No se encontraron registros.", vbInformation, "Mensaje"
            g_rst_Princi.Close
            Set g_rst_Princi = Nothing
            Exit Sub
         End If
         
         g_rst_Princi.MoveFirst
         Do While Not g_rst_Princi.EOF
            If Len(Trim(g_rst_Princi!NOCERTIFICADO)) > 0 Then
               .Cells(r_int_ConVer, 1) = Trim(g_rst_Princi!POLIZA)
               .Cells(r_int_ConVer, 2) = "'" & Trim(g_rst_Princi!NOCERTIFICADO)
               .Cells(r_int_ConVer, 3) = "'" & CDate(gf_FormatoFecha(CStr(g_rst_Princi!FECHAAFILIACION)))
               .Cells(r_int_ConVer, 4) = "'" & CDate(gf_FormatoFecha(CStr(g_rst_Princi!FECHA_VENCIMIENTO)))
               .Cells(r_int_ConVer, 5) = Format(g_rst_Princi!SUMAASEGURADA, "###,###,###,##0.00")
               .Cells(r_int_ConVer, 6) = Format(l_dbl_TasaMensual_Cia, "##0.000000")
               .Cells(r_int_ConVer, 7) = g_rst_Princi!PRIMENETA
               .Cells(r_int_ConVer, 8) = g_rst_Princi!PLAN
               .Cells(r_int_ConVer, 9) = g_rst_Princi!TIPPER_CONTRATANTE
               .Cells(r_int_ConVer, 10) = g_rst_Princi!TIPDOC_CONTRATANTE
               .Cells(r_int_ConVer, 11) = g_rst_Princi!NUMDOC_CONTRATANTE
               .Cells(r_int_ConVer, 12) = g_rst_Princi!APEPAT_CONTRATANTE
               .Cells(r_int_ConVer, 13) = g_rst_Princi!APEMAT_CONTRATANTE
               .Cells(r_int_ConVer, 14) = g_rst_Princi!NOMBRE_CONTRATANTE
               
               .Cells(r_int_ConVer, 18) = g_rst_Princi!DIREC_CONTRATANTE
               .Cells(r_int_ConVer, 28) = "'" & CStr(g_rst_Princi!DEPART_CONTRATANTE)
               .Cells(r_int_ConVer, 29) = "'" & CStr(g_rst_Princi!PROVIN_CONTRATANTE)
               .Cells(r_int_ConVer, 30) = "'" & CStr(g_rst_Princi!DISTRI_CONTRATANTE)
               .Cells(r_int_ConVer, 31) = "'" & CStr(g_rst_Princi!TELEFO_CONTRATANTE)
               
               .Cells(r_int_ConVer, 33) = "'" & CStr(g_rst_Princi!ESTCIV_CONTRATANTE)
               
               .Cells(r_int_ConVer, 35) = Trim(g_rst_Princi!TIPPER_ASEGURADO)
               .Cells(r_int_ConVer, 36) = Trim(g_rst_Princi!TIPDOC_ASEGURADO)
               .Cells(r_int_ConVer, 37) = "'" & CStr(g_rst_Princi!NUMDOC_ASEGURADO)
               .Cells(r_int_ConVer, 38) = Trim(g_rst_Princi!APEPAT_ASEGURADO)
               .Cells(r_int_ConVer, 39) = Trim(g_rst_Princi!APEMAT_ASEGURADO)
               .Cells(r_int_ConVer, 40) = Trim(g_rst_Princi!NOMCLI_ASEGURADO)
               .Cells(r_int_ConVer, 41) = "'" & CDate(gf_FormatoFecha(CStr(g_rst_Princi!FECNAC_ASEGURADO)))
               .Cells(r_int_ConVer, 42) = CStr(g_rst_Princi!SEXO_ASEGURADO)
               If Not IsNull(g_rst_Princi!TIPVIA_ASEGURADO) Then
                  .Cells(r_int_ConVer, 43) = CStr(g_rst_Princi!TIPVIA_ASEGURADO)
               End If
               .Cells(r_int_ConVer, 44) = CStr(g_rst_Princi!DIRECC_ASEGURADO)
               
               .Cells(r_int_ConVer, 54) = "'" & CStr(g_rst_Princi!DEPART_ASEGURADO)
               .Cells(r_int_ConVer, 55) = "'" & CStr(g_rst_Princi!PROVIN_ASEGURADO)
               .Cells(r_int_ConVer, 56) = "'" & CStr(g_rst_Princi!DISTRT_ASEGURADO)
               If Not IsNull(g_rst_Princi!TELFIJ_ASEGURADO) Then
                  .Cells(r_int_ConVer, 57) = "'" & CStr(g_rst_Princi!TELFIJ_ASEGURADO)
               End If
               If Not IsNull(g_rst_Princi!CELULA_ASEGURADO) Then
                  .Cells(r_int_ConVer, 58) = "'" & CStr(g_rst_Princi!CELULA_ASEGURADO)
               End If
               .Cells(r_int_ConVer, 59) = Trim(g_rst_Princi!ESTCIV_ASEGURADO)
               .Cells(r_int_ConVer, 60) = Trim(g_rst_Princi!CORREO_ASEGURADO)
               
               
               .Cells(r_int_ConVer, 61) = Trim(g_rst_Princi!TIPO_NEGOCIO)
               .Cells(r_int_ConVer, 62) = Trim(g_rst_Princi!CAPITAL_BASICO)
               .Cells(r_int_ConVer, 63) = Trim(g_rst_Princi!TIPO_ESTRUCTURA)
               .Cells(r_int_ConVer, 64) = CStr(g_rst_Princi!USO_INMUEBLE)
               .Cells(r_int_ConVer, 65) = CStr(g_rst_Princi!TIPO_EDIFICACION)
               .Cells(r_int_ConVer, 66) = CStr(g_rst_Princi!ANO_CONSTRUCCION)
               If Not IsNull(g_rst_Princi!ESTADO_PREDIO) Then
                  .Cells(r_int_ConVer, 67) = CStr(g_rst_Princi!ESTADO_PREDIO)
               End If
               If Not IsNull(g_rst_Princi!DETALLE) Then
                  .Cells(r_int_ConVer, 68) = CStr(g_rst_Princi!DETALLE)
               End If
               .Cells(r_int_ConVer, 69) = CStr(g_rst_Princi!TIPOVIA_MATERIA)
               .Cells(r_int_ConVer, 70) = CStr(g_rst_Princi!DIRECCION_MATERIA)
               .Cells(r_int_ConVer, 80) = "'" & CStr(g_rst_Princi!DEPART_MATERIA)
               .Cells(r_int_ConVer, 81) = "'" & CStr(g_rst_Princi!PROVIN_MATERIA)
               .Cells(r_int_ConVer, 82) = "'" & CStr(g_rst_Princi!DISTRT_MATERIA)
               .Cells(r_int_ConVer, 83) = CStr(g_rst_Princi!NUMERO_PISOS)
               .Cells(r_int_ConVer, 84) = CStr(g_rst_Princi!NUMERO_SOTANOS)
               
               r_int_ConVer = r_int_ConVer + 1
            End If
            
            g_rst_Princi.MoveNext
            DoEvents
         Loop
               
         g_rst_Princi.Close
         Set g_rst_Princi = Nothing
      End With
   Next
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub
Private Sub fs_GenExc_SegImb_old()
Dim r_obj_Excel      As Excel.Application
Dim r_str_Fecha      As String
Dim r_str_FecIni     As String
Dim r_str_FecFin     As String
Dim r_dbl_Monto      As Double
Dim r_dbl_PorIGV     As Double
Dim r_dbl_Factor     As Double
Dim r_int_Contad     As Integer
Dim r_int_TemAux     As Integer
Dim r_int_ConVer     As Integer
Dim r_int_Index      As Integer

   r_str_FecIni = Format(ipp_PerAno.Text, "0000") & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") & "25"
   If cmb_PerMes.ItemData(cmb_PerMes.ListIndex) = 12 Then
      r_str_FecFin = Format(ipp_PerAno.Text + 1, "0000") & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex) + 1, "00") & "02"
   Else
      r_str_FecFin = Format(ipp_PerAno.Text, "0000") & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex) + 1, "00") & "02"
   End If
   
   '-- Obtiene el IGV
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM MNT_PARVAL "
   g_str_Parame = g_str_Parame & " WHERE PARVAL_CODGRP = '001' "
   g_str_Parame = g_str_Parame & "   AND PARVAL_CODITE = '001' "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
       Exit Sub
   End If
   
   g_rst_Genera.MoveFirst
   r_dbl_PorIGV = Format(g_rst_Genera!PARVAL_CANTID, "###,###,###,##0.000000")
   
   g_rst_Genera.Close
   Set g_rst_Genera = Nothing
   
   '-- Prepara excel
   r_str_Fecha = "01/" & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") & "/" & Format(ipp_PerAno.Text, "0000")
   r_dbl_Monto = 0
   
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 2
   r_obj_Excel.Workbooks.Add
   
   For r_int_TemAux = 1 To 2 Step 1
      Select Case r_int_TemAux
         Case 1: r_obj_Excel.Sheets(r_int_TemAux).Name = "Soles"
         Case 2: r_obj_Excel.Sheets(r_int_TemAux).Name = "Dolares"
      End Select
      
      With r_obj_Excel.Sheets(r_int_TemAux)
         .Cells(2, 1) = "MES : " & ObtieneNomMes(cmb_PerMes.ItemData(cmb_PerMes.ListIndex)) & " - " & Format(ipp_PerAno.Text, "0000")
         .Range(.Cells(2, 1), .Cells(2, 15)).HorizontalAlignment = xlHAlignCenter
         .Range(.Cells(2, 1), .Cells(2, 15)).Font.Bold = True
         .Range(.Cells(2, 1), .Cells(2, 15)).Merge
         .Range(.Cells(2, 1), .Cells(2, 15)).Font.Size = 18
         
         .Cells(3, 1) = "MONEDA: " & IIf(r_int_TemAux = 1, "SOLES", "DOLARES")
         .Range(.Cells(3, 1), .Cells(3, 3)).HorizontalAlignment = xlHAlignLeft
         .Range(.Cells(3, 1), .Cells(3, 3)).Font.Bold = True
         .Range(.Cells(3, 1), .Cells(3, 3)).Merge
         
         .Cells(5, 1) = "DOMICILIARIO"
         .Range(.Cells(5, 1), .Cells(5, 13)).Merge
         
         .Cells(5, 13) = "CORRESPONDENCIA"
         .Range(.Cells(5, 14), .Cells(5, 17)).Merge
         
         .Cells(5, 17) = "DATOS DEL INMUEBLE POR ASEGURAR"
         .Range(.Cells(5, 18), .Cells(5, 29)).Merge
         
         .Range(.Cells(5, 1), .Cells(5, 37)).HorizontalAlignment = xlHAlignCenter
         .Range(.Cells(5, 1), .Cells(5, 37)).Font.Bold = True
         
         For r_int_Contad = 1 To 29 Step 1
            .Range(.Cells(5, 1), .Cells(5, r_int_Contad)).Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Range(.Cells(5, 1), .Cells(5, r_int_Contad)).Borders(xlEdgeTop).LineStyle = xlContinuous
            .Range(.Cells(5, 1), .Cells(5, r_int_Contad)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Range(.Cells(5, 1), .Cells(5, r_int_Contad)).Borders(xlEdgeRight).LineStyle = xlContinuous
            .Range(.Cells(5, 1), .Cells(5, r_int_Contad)).Borders(xlInsideVertical).LineStyle = xlContinuous
         Next
         
         For r_int_Contad = 1 To 38 Step 1
            .Range(.Cells(6, 1), .Cells(6, r_int_Contad)).Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Range(.Cells(6, 1), .Cells(6, r_int_Contad)).Borders(xlEdgeTop).LineStyle = xlContinuous
            .Range(.Cells(6, 1), .Cells(6, r_int_Contad)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Range(.Cells(6, 1), .Cells(6, r_int_Contad)).Borders(xlEdgeRight).LineStyle = xlContinuous
            .Range(.Cells(6, 1), .Cells(6, r_int_Contad)).Borders(xlInsideVertical).LineStyle = xlContinuous
         Next
         
         For r_int_Contad = 39 To 40 Step 1
            .Range(.Cells(7, r_int_Contad), .Cells(7, r_int_Contad)).Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Range(.Cells(7, r_int_Contad), .Cells(7, r_int_Contad)).Borders(xlEdgeTop).LineStyle = xlContinuous
            .Range(.Cells(7, r_int_Contad), .Cells(7, r_int_Contad)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Range(.Cells(7, r_int_Contad), .Cells(7, r_int_Contad)).Borders(xlEdgeRight).LineStyle = xlContinuous
            .Range(.Cells(7, r_int_Contad), .Cells(7, r_int_Contad)).Borders(xlInsideVertical).LineStyle = xlContinuous
         Next
         
         .Cells(6, 1) = "OPERADORA"
         .Cells(6, 2) = "FECHA AFILIACION"
         .Cells(6, 3) = "FECHA VENCIMIENTO"
         .Cells(6, 4) = "N° DE CERTIFICADO"
         .Cells(6, 5) = "MONEDA"
         .Cells(6, 6) = "TIPO DOCUMENTO"
         .Cells(6, 7) = "N° DE DOCUMENTO"
         .Cells(6, 8) = "AP. PATERNO"
         .Cells(6, 9) = "AP. MATERNO"
         .Cells(6, 10) = "NOMBRES"
         .Cells(6, 11) = "ESTADO CIVIL"
         .Cells(6, 12) = "FECHA NACIMIENTO"
         .Cells(6, 13) = "SEXO"
         .Cells(6, 14) = "EDAD"
         .Cells(6, 15) = "DIRECCION"
         .Cells(6, 16) = "DEPARTAMENTO"
         .Cells(6, 17) = "PROVINCIA"
         .Cells(6, 18) = "DISTRITO"
         .Cells(6, 19) = "TIPO INMUEBLE"
         .Cells(6, 20) = "DIRECCION"
         .Cells(6, 21) = "DEPARTAMENTO"
         .Cells(6, 22) = "PROVINCIA"
         .Cells(6, 23) = "DISTRITO"
         .Cells(6, 24) = "USO DEL BIEN"
         .Cells(6, 25) = "TIPO DE EDIFICACION"
         .Cells(6, 26) = "AÑO DE CONSTRUCCION"
         .Cells(6, 27) = "MATERIAL DE CONSTRUCCION"
         .Cells(6, 28) = "NUMERO DE PISOS"
         .Cells(6, 29) = "NUMERO DE SOTANOS"
         .Cells(6, 30) = "TIPO DE GARANTIA"
         .Cells(6, 31) = "SUMA ASEGURADA"
         .Cells(6, 32) = "TASA NETA MENSUAL CIA"
         .Cells(6, 33) = "PRIMA NETA MENSUAL CIA"
         .Cells(6, 34) = "TASA NETA MENSUAL CLIENTE"
         .Cells(6, 35) = "PRIMA NETA MENSUAL CLIENTE"
         .Cells(6, 36) = "PRIMA TOTAL MENSUAL CLIENTE"
         .Cells(6, 37) = "COMISION"
         .Cells(6, 38) = "CUOTA MES"
         
         .Range(.Cells(6, 1), .Cells(6, 38)).Font.Bold = True
         .Range(.Cells(6, 1), .Cells(6, 38)).HorizontalAlignment = xlHAlignCenter
         
         .Columns("A").ColumnWidth = 22
         .Columns("A").HorizontalAlignment = xlHAlignCenter
         .Columns("B").ColumnWidth = 14
         .Columns("B").HorizontalAlignment = xlHAlignCenter
         .Columns("C").ColumnWidth = 14
         .Columns("C").HorizontalAlignment = xlHAlignCenter
         .Columns("D").ColumnWidth = 22
         .Columns("D").HorizontalAlignment = xlHAlignCenter
         .Columns("E").ColumnWidth = 22
         .Columns("E").HorizontalAlignment = xlHAlignCenter
         .Columns("F").ColumnWidth = 14
         .Columns("F").HorizontalAlignment = xlHAlignCenter
         .Columns("G").ColumnWidth = 14
         .Columns("G").HorizontalAlignment = xlHAlignCenter
         .Columns("H").ColumnWidth = 18
         .Columns("H").HorizontalAlignment = xlHAlignLeft
         .Columns("I").ColumnWidth = 18
         .Columns("I").HorizontalAlignment = xlHAlignLeft
         .Columns("J").ColumnWidth = 24
         .Columns("J").HorizontalAlignment = xlHAlignLeft
         .Columns("K").ColumnWidth = 16
         .Columns("K").HorizontalAlignment = xlHAlignCenter
         .Columns("L").ColumnWidth = 14
         .Columns("L").HorizontalAlignment = xlHAlignCenter
         .Columns("M").ColumnWidth = 7
         .Columns("M").HorizontalAlignment = xlHAlignCenter
         .Columns("N").ColumnWidth = 7
         .Columns("N").HorizontalAlignment = xlHAlignCenter
         .Columns("O").ColumnWidth = 80
         .Columns("O").HorizontalAlignment = xlHAlignLeft
         .Columns("P").ColumnWidth = 16
         .Columns("P").HorizontalAlignment = xlHAlignCenter
         .Columns("Q").ColumnWidth = 16
         .Columns("Q").HorizontalAlignment = xlHAlignCenter
         .Columns("R").ColumnWidth = 28
         .Columns("R").HorizontalAlignment = xlHAlignCenter
         .Columns("S").ColumnWidth = 18
         .Columns("S").HorizontalAlignment = xlHAlignCenter
         .Columns("T").ColumnWidth = 80
         .Columns("T").HorizontalAlignment = xlHAlignLeft
         .Columns("U").ColumnWidth = 16
         .Columns("U").HorizontalAlignment = xlHAlignCenter
         .Columns("V").ColumnWidth = 16
         .Columns("V").HorizontalAlignment = xlHAlignCenter
         .Columns("W").ColumnWidth = 28
         .Columns("W").HorizontalAlignment = xlHAlignCenter
         .Columns("X").ColumnWidth = 30
         .Columns("X").HorizontalAlignment = xlHAlignCenter
         .Columns("Y").ColumnWidth = 25
         .Columns("Y").HorizontalAlignment = xlHAlignCenter
         .Columns("Z").ColumnWidth = 15
         .Columns("Z").HorizontalAlignment = xlHAlignCenter
         .Columns("AA").ColumnWidth = 60
         .Columns("AA").HorizontalAlignment = xlHAlignCenter
         .Columns("AB").ColumnWidth = 12
         .Columns("AB").HorizontalAlignment = xlHAlignCenter
         .Columns("AC").ColumnWidth = 12
         .Columns("AC").HorizontalAlignment = xlHAlignCenter
         .Columns("AD").ColumnWidth = 30
         .Columns("AD").HorizontalAlignment = xlHAlignCenter
         .Columns("AE").ColumnWidth = 14
         .Columns("AE").NumberFormat = "###,###,###,##0.00"
         .Columns("AF").ColumnWidth = 14
         .Columns("AF").NumberFormat = "###,###,###,##0.000000"
         .Columns("AG").ColumnWidth = 14
         .Columns("AG").NumberFormat = "###,###,###,##0.00"
         .Columns("AH").ColumnWidth = 14
         .Columns("AH").NumberFormat = "###,###,###,##0.000000"
         .Columns("AI").ColumnWidth = 14
         .Columns("AI").NumberFormat = "###,###,###,##0.00"
         .Columns("AJ").ColumnWidth = 14
         .Columns("AJ").NumberFormat = "###,###,###,##0.00"
         .Columns("AK").ColumnWidth = 14
         .Columns("AK").NumberFormat = "###,###,###,##0.00"
         .Columns("AL").ColumnWidth = 13
         .Columns("AL").NumberFormat = "###,###,###,##0.00"
         .Columns("AM").ColumnWidth = 12
         .Columns("AN").ColumnWidth = 12
         
         For r_int_Index = 1 To 39 Step 1
           .Range(.Cells(6, r_int_Index), .Cells(6, r_int_Index)).WrapText = True
           .Range(.Cells(6, r_int_Index), .Cells(6, r_int_Index)).VerticalAlignment = xlCenter
           .Range(.Cells(6, r_int_Index), .Cells(6, r_int_Index)).HorizontalAlignment = xlHAlignCenter
         Next
         
         r_int_ConVer = 7
         r_dbl_Factor = CDbl(txtFactor.Text) * ((r_dbl_PorIGV / 100) + 1)
         .Cells(r_int_ConVer, 39) = Format(r_dbl_Factor, "##0.000000")
         .Cells(r_int_ConVer, 40) = "Factor Actual"
         .Range(.Cells(r_int_ConVer, 39), .Cells(r_int_ConVer, 40)).Font.Bold = True
         
         g_str_Parame = ""
         g_str_Parame = g_str_Parame & "SELECT 'EDPYME MICASITA S.A.' AS EMP, "
         g_str_Parame = g_str_Parame & "       POLIZA_FEMVIV AS FECHAAFILIACION, "
         g_str_Parame = g_str_Parame & "       HIPMAE_ULTVCT AS FECHA_VENCIMIENTO, "
         g_str_Parame = g_str_Parame & "       TRIM(POLIZA_NUMVIV) AS NOCERTIFICADO, "
         g_str_Parame = g_str_Parame & "       TRIM(HIPCIE_TDOCLI)||'-'||TRIM(HIPCIE_NDOCLI) AS CODIGO, "
         g_str_Parame = g_str_Parame & "       TRIM(HIPCIE_TDOCLI) AS TDOCLI, "
         g_str_Parame = g_str_Parame & "       TRIM(HIPCIE_NDOCLI) AS NDOCLI, "
         g_str_Parame = g_str_Parame & "       TRIM(TD.PARDES_DESCRI) AS TIPODOCUMENTO, "
         g_str_Parame = g_str_Parame & "       TRIM(DATGEN_APEPAT) AS APEPAT, "
         g_str_Parame = g_str_Parame & "       TRIM(DATGEN_APEMAT) AS APEMAT, "
         g_str_Parame = g_str_Parame & "       TRIM(DATGEN_NOMBRE) AS NOMCLIENTE, "
         g_str_Parame = g_str_Parame & "       TRIM(EC.PARDES_DESCRI) AS ESTCIVIL, "
         g_str_Parame = g_str_Parame & "       DATGEN_NACFEC AS FECNACIMIENTO, "
         g_str_Parame = g_str_Parame & "       DECODE(DATGEN_CODSEX,1,'M','F') AS SEXO, "
         g_str_Parame = g_str_Parame & "       TRUNC((SYSDATE - TO_DATE(DATGEN_NACFEC, 'YYYY/MM/DD'))/365,0) AS EDAD, "
         g_str_Parame = g_str_Parame & "       TRIM(NM.PARDES_DESCRI) AS TIPOINMUEBLE, "
         g_str_Parame = g_str_Parame & "       TRIM(IV.PARDES_DESCRI) ||' '|| TRIM(SOLINM_NOMVIA) ||' '|| TRIM(SOLINM_NUMVIA) ||' '|| TRIM(SOLINM_INTDPT) ||' '|| TRIM(ZN.PARDES_DESCRI) ||' '|| TRIM(SOLINM_NOMZON) AS DIRECCION, "
         g_str_Parame = g_str_Parame & "       TRIM(IT.PARDES_DESCRI) AS DEPARTAMENTO, "
         g_str_Parame = g_str_Parame & "       TRIM(IP.PARDES_DESCRI) AS PROVINCIA, "
         g_str_Parame = g_str_Parame & "       TRIM(II.PARDES_DESCRI) AS DISTRITO, "
         g_str_Parame = g_str_Parame & "       TRIM(SEGEMP_RAZSOC) AS EMPSEGUROS, "
         g_str_Parame = g_str_Parame & "       EVATAS_SUMASE_INM+EVATAS_SUMASE_ES1+EVATAS_SUMASE_ES2+EVATAS_SUMASE_DEP AS SUMAASEGURADA, "
         g_str_Parame = g_str_Parame & "       ROUND(HIPCIE_FOIVIV, 5) AS TASA, "
         g_str_Parame = g_str_Parame & "       ROUND(((HIPCIE_SALCAP+HIPCIE_SALCON)*HIPCIE_FOIVIV)/100,2) AS PRIMENETA, "
         g_str_Parame = g_str_Parame & "       TRIM(GR.PARDES_DESCRI) AS TIPO_GARANTIA, "
         g_str_Parame = g_str_Parame & "       TRIM(UI.PARDES_DESCRI) AS USO_INMUEBLE, "
         g_str_Parame = g_str_Parame & "       TRIM(TI.PARDES_DESCRI) AS TIPO_EDIFICACION, "
         g_str_Parame = g_str_Parame & "       TRIM(MC.PARDES_DESCRI) AS MATERIAL_CONSTRUCCION, "
         g_str_Parame = g_str_Parame & "       EVATAS_ANOCON AS ANO_CONSTRUCCION, "
         g_str_Parame = g_str_Parame & "       EVATAS_NUMPIS AS NUMERO_PISOS, "
         g_str_Parame = g_str_Parame & "       EVATAS_NUMSOT AS NUMERO_SOTANOS, TRIM(MD.PARDES_DESCRI) AS MONEDA, "
         g_str_Parame = g_str_Parame & "       NVL((SELECT SUM(HIPCUO_VIVORG) FROM CRE_HIPCUO WHERE HIPCUO_NUMOPE = HIPCIE_NUMOPE AND HIPCUO_TIPCRO = 1 "
         g_str_Parame = g_str_Parame & "               AND HIPCUO_FECVCT >= " & Format(ipp_PerAno.Text, "0000") & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") & "01 "
         g_str_Parame = g_str_Parame & "               AND HIPCUO_FECVCT <= " & Format(ipp_PerAno.Text, "0000") & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") & "31),0) AS CUOTA_MES "
         g_str_Parame = g_str_Parame & "  FROM CRE_HIPCIE "
         g_str_Parame = g_str_Parame & " INNER JOIN CLI_DATGEN    ON (DATGEN_TIPDOC = HIPCIE_TDOCLI AND DATGEN_NUMDOC = HIPCIE_NDOCLI) "
         g_str_Parame = g_str_Parame & " INNER JOIN CRE_HIPMAE    ON (HIPMAE_NUMOPE = HIPCIE_NUMOPE) "
         g_str_Parame = g_str_Parame & " INNER JOIN TRA_POLIZA    ON (POLIZA_NUMSOL = HIPMAE_NUMSOL) "
         g_str_Parame = g_str_Parame & " INNER JOIN TRA_EVASEG    ON (EVASEG_NUMSOL = HIPMAE_NUMSOL) "
         g_str_Parame = g_str_Parame & " INNER JOIN TRA_EVATAS    ON (EVATAS_NUMSOL = HIPMAE_NUMSOL) "
         g_str_Parame = g_str_Parame & " INNER JOIN MNT_SEGEMP    ON (SEGEMP_CODIGO = EVASEG_ESGDES) "
         g_str_Parame = g_str_Parame & " INNER JOIN MNT_SEGTIP    ON (SEGTIP_CODIGO = EVASEG_ESGDES AND SEGTIP_TIPSEG = EVASEG_TIPSEG) "
         g_str_Parame = g_str_Parame & " INNER JOIN CRE_SOLINM    ON (SOLINM_NUMSOL = HIPMAE_NUMSOL) "
         g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES II ON (DATGEN_UBIGEO = II.PARDES_CODITE AND  II.PARDES_CODGRP = 101 ) "
         g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES IV ON (DATGEN_TIPVIA = IV.PARDES_CODITE AND  IV.PARDES_CODGRP = 201 ) "
         g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES ZN ON (SOLINM_TIPZON = ZN.PARDES_CODITE AND  ZN.PARDES_CODGRP = 202 ) "
         g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES IT ON (IT.PARDES_CODGRP = 101 AND RPAD(SUBSTR(SOLINM_UBIGEO,1,2),6,0)= IT.PARDES_CODITE) "
         g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES IP ON (IP.PARDES_CODGRP = 101 AND RPAD(SUBSTR(SOLINM_UBIGEO,1,4),6,0)= IP.PARDES_CODITE) "
         g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES NM ON (NM.PARDES_CODGRP = 217 AND SOLINM_TIPINM = NM.PARDES_CODITE) "
         g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES EC ON (EC.PARDES_CODGRP = '205' AND EC.PARDES_CODITE = DATGEN_ESTCIV) "
         g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES TD ON (TD.PARDES_CODGRP = '203' AND TD.PARDES_CODITE = HIPCIE_TDOCLI) "
         g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES GR ON (GR.PARDES_CODGRP = '241' AND GR.PARDES_CODITE = HIPCIE_TIPGAR) "
         g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES UI ON (UI.PARDES_CODGRP = '222' AND UI.PARDES_CODITE = EVATAS_USOINM) "
         g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES TI ON (TI.PARDES_CODGRP = '221' AND TI.PARDES_CODITE = EVATAS_TIPINM) "
         g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES MC ON (MC.PARDES_CODGRP = '223' AND MC.PARDES_CODITE = EVATAS_MATCON) "
         g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES MD ON (MD.PARDES_CODGRP = '204' AND MD.PARDES_CODITE = HIPCIE_TIPMON) "
         g_str_Parame = g_str_Parame & " WHERE HIPCIE_PERMES = " & cmb_PerMes.ItemData(cmb_PerMes.ListIndex) & " "
         g_str_Parame = g_str_Parame & "   AND HIPCIE_PERANO = " & Format(ipp_PerAno.Text, "0000") & " "
         g_str_Parame = g_str_Parame & "   AND HIPCIE_TIPMON = " & r_int_TemAux & " "
         g_str_Parame = g_str_Parame & "UNION "
         g_str_Parame = g_str_Parame & "SELECT 'EDPYME MICASITA S.A.' AS EMP, "
         g_str_Parame = g_str_Parame & "       POLIZA_FEMVIV AS FECHAAFILIACION, "
         g_str_Parame = g_str_Parame & "       HIPMAE_ULTVCT AS FECHA_VENCIMIENTO, "
         g_str_Parame = g_str_Parame & "       TRIM(POLIZA_NUMVIV) AS NOCERTIFICADO, "
         g_str_Parame = g_str_Parame & "       TRIM(HIPMAE_TDOCLI)||'-'||TRIM(HIPMAE_NDOCLI) AS CODIGO, "
         g_str_Parame = g_str_Parame & "       TRIM(HIPMAE_TDOCLI) AS TDOCLI, "
         g_str_Parame = g_str_Parame & "       TRIM(HIPMAE_NDOCLI) AS NDOCLI, "
         g_str_Parame = g_str_Parame & "       TRIM(TD.PARDES_DESCRI) AS TIPODOCUMENTO, "
         g_str_Parame = g_str_Parame & "       TRIM(DATGEN_APEPAT) AS APEPAT, "
         g_str_Parame = g_str_Parame & "       TRIM(DATGEN_APEMAT) AS APEMAT, "
         g_str_Parame = g_str_Parame & "       TRIM(DATGEN_NOMBRE) AS NOMCLIENTE, "
         g_str_Parame = g_str_Parame & "       TRIM(EC.PARDES_DESCRI) AS ESTCIVIL, "
         g_str_Parame = g_str_Parame & "       DATGEN_NACFEC AS FECNACIMIENTO, "
         g_str_Parame = g_str_Parame & "       DECODE(DATGEN_CODSEX,1,'M','F') AS SEXO, "
         g_str_Parame = g_str_Parame & "       TRUNC((SYSDATE - TO_DATE(DATGEN_NACFEC, 'YYYY/MM/DD'))/365,0) AS EDAD, "
         g_str_Parame = g_str_Parame & "       TRIM(NM.PARDES_DESCRI) AS TIPOINMUEBLE, "
         g_str_Parame = g_str_Parame & "       TRIM(IV.PARDES_DESCRI) ||' '|| TRIM(SOLINM_NOMVIA) ||' '|| TRIM(SOLINM_NUMVIA) ||' '|| TRIM(SOLINM_INTDPT) ||' '|| TRIM(ZN.PARDES_DESCRI) ||' '|| TRIM(SOLINM_NOMZON) AS DIRECCION, "
         g_str_Parame = g_str_Parame & "       TRIM(IT.PARDES_DESCRI) AS DEPARTAMENTO, "
         g_str_Parame = g_str_Parame & "       TRIM(IP.PARDES_DESCRI) AS PROVINCIA, "
         g_str_Parame = g_str_Parame & "       TRIM(II.PARDES_DESCRI) AS DISTRITO, "
         g_str_Parame = g_str_Parame & "       TRIM(SEGEMP_RAZSOC) AS EMPSEGUROS, "
         g_str_Parame = g_str_Parame & "       EVATAS_SUMASE_INM+EVATAS_SUMASE_ES1+EVATAS_SUMASE_ES2+EVATAS_SUMASE_DEP AS SUMAASEGURADA, "
         g_str_Parame = g_str_Parame & "       ROUND(HIPMAE_FOIVIV, 5) AS TASA, "
         g_str_Parame = g_str_Parame & "       ROUND(((HIPMAE_SALCAP+HIPMAE_SALCON)*HIPMAE_FOIVIV)/100,2) AS PRIMENETA, "
         g_str_Parame = g_str_Parame & "       TRIM(GR.PARDES_DESCRI) AS TIPO_GARANTIA, "
         g_str_Parame = g_str_Parame & "       TRIM(UI.PARDES_DESCRI) AS USO_INMUEBLE, "
         g_str_Parame = g_str_Parame & "       TRIM(TI.PARDES_DESCRI) AS TIPO_EDIFICACION, "
         g_str_Parame = g_str_Parame & "       TRIM(MC.PARDES_DESCRI) AS MATERIAL_CONSTRUCCION, "
         g_str_Parame = g_str_Parame & "       EVATAS_ANOCON AS ANO_CONSTRUCCION, "
         g_str_Parame = g_str_Parame & "       EVATAS_NUMPIS AS NUMERO_PISOS, "
         g_str_Parame = g_str_Parame & "       EVATAS_NUMSOT As NUMERO_SOTANOS, TRIM(MD.PARDES_DESCRI) AS MONEDA, "
         g_str_Parame = g_str_Parame & "       (SELECT SUM(HIPCUO_VIVORG) FROM CRE_HIPCUO WHERE HIPCUO_NUMOPE = HIPMAE_NUMOPE AND HIPCUO_TIPCRO = 1 "
         g_str_Parame = g_str_Parame & "           AND HIPCUO_FECVCT >= " & Format(ipp_PerAno.Text, "0000") & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") & "01 "
         g_str_Parame = g_str_Parame & "           AND HIPCUO_FECVCT <= " & Format(ipp_PerAno.Text, "0000") & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") & "31) AS CUOTA_MES "
         g_str_Parame = g_str_Parame & "  FROM CRE_HIPMAE "
         g_str_Parame = g_str_Parame & " INNER JOIN CLI_DATGEN    ON (DATGEN_TIPDOC = HIPMAE_TDOCLI AND DATGEN_NUMDOC = HIPMAE_NDOCLI) "
         g_str_Parame = g_str_Parame & " INNER JOIN TRA_POLIZA    ON (POLIZA_NUMSOL = HIPMAE_NUMSOL) "
         g_str_Parame = g_str_Parame & " INNER JOIN TRA_EVASEG    ON (EVASEG_NUMSOL = HIPMAE_NUMSOL) "
         g_str_Parame = g_str_Parame & " INNER JOIN TRA_EVATAS    ON (EVATAS_NUMSOL = HIPMAE_NUMSOL) "
         g_str_Parame = g_str_Parame & " INNER JOIN MNT_SEGEMP    ON (SEGEMP_CODIGO = EVASEG_ESGDES) "
         g_str_Parame = g_str_Parame & " INNER JOIN MNT_SEGTIP    ON (SEGTIP_CODIGO = EVASEG_ESGDES AND SEGTIP_TIPSEG = EVASEG_TIPSEG) "
         g_str_Parame = g_str_Parame & " INNER JOIN CRE_SOLINM    ON (SOLINM_NUMSOL = HIPMAE_NUMSOL) "
         g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES II ON (DATGEN_UBIGEO = II.PARDES_CODITE AND  II.PARDES_CODGRP = 101 ) "
         g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES IV ON (DATGEN_TIPVIA = IV.PARDES_CODITE AND  IV.PARDES_CODGRP = 201 ) "
         g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES ZN ON (SOLINM_TIPZON = ZN.PARDES_CODITE AND  ZN.PARDES_CODGRP = 202 ) "
         g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES IT ON (IT.PARDES_CODGRP = 101 AND RPAD(SUBSTR(SOLINM_UBIGEO,1,2),6,0)= IT.PARDES_CODITE) "
         g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES IP ON (IP.PARDES_CODGRP = 101 AND RPAD(SUBSTR(SOLINM_UBIGEO,1,4),6,0)= IP.PARDES_CODITE) "
         g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES NM ON (NM.PARDES_CODGRP = 217 AND SOLINM_TIPINM = NM.PARDES_CODITE) "
         g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES EC ON (EC.PARDES_CODGRP = '205' AND EC.PARDES_CODITE = DATGEN_ESTCIV) "
         g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES TD ON (TD.PARDES_CODGRP = '203' AND TD.PARDES_CODITE = HIPMAE_TDOCLI) "
         g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES GR ON (GR.PARDES_CODGRP = '241' AND GR.PARDES_CODITE = HIPMAE_TIPGAR) "
         g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES UI ON (UI.PARDES_CODGRP = '222' AND UI.PARDES_CODITE = EVATAS_USOINM) "
         g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES TI ON (TI.PARDES_CODGRP = '221' AND TI.PARDES_CODITE = EVATAS_TIPINM) "
         g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES MC ON (MC.PARDES_CODGRP = '223' AND MC.PARDES_CODITE = EVATAS_MATCON) "
         g_str_Parame = g_str_Parame & "  LEFT JOIN MNT_PARDES MD ON (MD.PARDES_CODGRP = '204' AND MD.PARDES_CODITE = HIPMAE_MONEDA) "
         g_str_Parame = g_str_Parame & " WHERE HIPMAE_SITUAC = 6 "
         g_str_Parame = g_str_Parame & "   AND HIPMAE_MONEDA = " & r_int_TemAux & " "
         g_str_Parame = g_str_Parame & "   AND HIPMAE_FECCAN >= " & r_str_FecIni
         g_str_Parame = g_str_Parame & "   AND HIPMAE_FECCAN <= " & r_str_FecFin
         
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
            If Len(Trim(g_rst_Princi!NOCERTIFICADO)) > 0 Then
               .Cells(r_int_ConVer, 1) = Trim(g_rst_Princi!EMP)
               .Cells(r_int_ConVer, 2) = "'" & CDate(gf_FormatoFecha(CStr(g_rst_Princi!FECHAAFILIACION)))
               .Cells(r_int_ConVer, 3) = "'" & CDate(gf_FormatoFecha(CStr(g_rst_Princi!FECHA_VENCIMIENTO)))
               .Cells(r_int_ConVer, 4) = "'" & Trim(g_rst_Princi!NOCERTIFICADO)
               .Cells(r_int_ConVer, 5) = "'" & Trim(g_rst_Princi!Moneda)
               .Cells(r_int_ConVer, 6) = Trim(g_rst_Princi!TDOCLI)
               .Cells(r_int_ConVer, 7) = "'" & CStr(g_rst_Princi!NDOCLI)
               .Cells(r_int_ConVer, 8) = Trim(g_rst_Princi!ApePat)
               .Cells(r_int_ConVer, 9) = Trim(g_rst_Princi!ApeMat)
               .Cells(r_int_ConVer, 10) = Trim(g_rst_Princi!NOMCLIENTE)
               .Cells(r_int_ConVer, 11) = Trim(g_rst_Princi!ESTCIVIL)
               .Cells(r_int_ConVer, 12) = "'" & CDate(gf_FormatoFecha(CStr(g_rst_Princi!FECNACIMIENTO)))
               .Cells(r_int_ConVer, 13) = CStr(g_rst_Princi!SEXO)
               .Cells(r_int_ConVer, 14) = Trim(g_rst_Princi!EDAD)
               .Cells(r_int_ConVer, 15) = CStr(g_rst_Princi!Direccion)
               .Cells(r_int_ConVer, 16) = CStr(g_rst_Princi!DEPARTAMENTO)
               .Cells(r_int_ConVer, 17) = CStr(g_rst_Princi!PROVINCIA)
               .Cells(r_int_ConVer, 18) = CStr(g_rst_Princi!DISTRITO)
               .Cells(r_int_ConVer, 19) = Trim(g_rst_Princi!tipoInmueble)
               .Cells(r_int_ConVer, 20) = CStr(g_rst_Princi!Direccion)
               .Cells(r_int_ConVer, 21) = CStr(g_rst_Princi!DEPARTAMENTO)
               .Cells(r_int_ConVer, 22) = CStr(g_rst_Princi!PROVINCIA)
               .Cells(r_int_ConVer, 23) = CStr(g_rst_Princi!DISTRITO)
               .Cells(r_int_ConVer, 24) = CStr(g_rst_Princi!USO_INMUEBLE)
               .Cells(r_int_ConVer, 25) = CStr(g_rst_Princi!TIPO_EDIFICACION)
               .Cells(r_int_ConVer, 26) = CStr(g_rst_Princi!ANO_CONSTRUCCION)
               .Cells(r_int_ConVer, 27) = CStr(g_rst_Princi!MATERIAL_CONSTRUCCION)
               .Cells(r_int_ConVer, 28) = CStr(g_rst_Princi!NUMERO_PISOS)
               .Cells(r_int_ConVer, 29) = CStr(g_rst_Princi!NUMERO_SOTANOS)
               .Cells(r_int_ConVer, 30) = CStr(g_rst_Princi!TIPO_GARANTIA)
               .Cells(r_int_ConVer, 31) = Format(g_rst_Princi!SUMAASEGURADA, "###,###,###,##0.00")
               .Cells(r_int_ConVer, 32) = Format(l_dbl_TasaMensual_Cia, "##0.000000")
               .Cells(r_int_ConVer, 33) = Format(g_rst_Princi!SUMAASEGURADA * l_dbl_TasaMensual_Cia, "###,##0.00")
               .Cells(r_int_ConVer, 34) = Format(l_dbl_TasaMensual_Cli, "##0.000000")
               .Cells(r_int_ConVer, 35) = Format(g_rst_Princi!SUMAASEGURADA * l_dbl_TasaMensual_Cli, "###,##0.00")
               .Cells(r_int_ConVer, 36) = Format(g_rst_Princi!SUMAASEGURADA * l_dbl_TasaMensual_Cli * r_dbl_Factor, "###,##0.00")
               .Cells(r_int_ConVer, 37) = Format((g_rst_Princi!SUMAASEGURADA * l_dbl_TasaMensual_Cli) - (g_rst_Princi!SUMAASEGURADA * l_dbl_TasaMensual_Cia), "###,##0.00")
               .Cells(r_int_ConVer, 38) = Format(g_rst_Princi!CUOTA_MES, "###,##0.00")
               r_int_ConVer = r_int_ConVer + 1
            End If
            
            g_rst_Princi.MoveNext
            DoEvents
         Loop
         
         r_int_ConVer = r_int_ConVer + 1
         .Cells(r_int_ConVer, 35) = "Prima Neta Mensual Cliente " & IIf(r_int_TemAux = 1, "S/.", "US$")
         .Cells(r_int_ConVer, 37).Formula = "=SUM(AI7" & ":" & "AI" & r_int_ConVer - 2 & ")"
         .Range(.Cells(r_int_ConVer, 35), .Cells(r_int_ConVer, 37)).Font.Bold = True
         
         r_int_ConVer = r_int_ConVer + 1
         .Cells(r_int_ConVer, 35) = "Prima Total Mensual Cliente " & IIf(r_int_TemAux = 1, "S/.", "US$")
         .Cells(r_int_ConVer, 37).Formula = "=SUM(AJ7" & ":" & "AJ" & r_int_ConVer - 3 & ")"
         .Range(.Cells(r_int_ConVer, 35), .Cells(r_int_ConVer, 37)).Font.Bold = True
         
         g_rst_Princi.Close
         Set g_rst_Princi = Nothing
      End With
   Next
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub
Private Sub fs_GenExc_Endoso()
Dim r_obj_Excel      As Excel.Application
Dim r_int_NroFil     As Integer
Dim r_str_NumCuo     As String
Dim r_str_FecVct     As String
Dim r_str_FecIni     As String
Dim r_str_FecFin     As String

   r_str_FecIni = Format(ipp_PerAno.Text, "0000") & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex), "00") & "25"
   If cmb_PerMes.ItemData(cmb_PerMes.ListIndex) = 12 Then
      r_str_FecFin = Format(ipp_PerAno.Text + 1, "0000") & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex) + 1, "00") & "02"
   Else
      r_str_FecFin = Format(ipp_PerAno.Text, "0000") & Format(cmb_PerMes.ItemData(cmb_PerMes.ListIndex) + 1, "00") & "02"
   End If
   
   'Generando Excel
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.DisplayAlerts = False
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   r_int_NroFil = 1
   
   With r_obj_Excel.ActiveSheet
      .Cells(r_int_NroFil, 1) = "ITEM":                     .Columns("A").ColumnWidth = 7
      .Cells(r_int_NroFil, 2) = "OPERACION":                .Columns("B").ColumnWidth = 14
      .Cells(r_int_NroFil, 3) = "TIPO DE SEGURO":           .Columns("C").ColumnWidth = 22
      .Cells(r_int_NroFil, 4) = "TIPO ACCION":              .Columns("D").ColumnWidth = 15
      .Cells(r_int_NroFil, 5) = "ESTADO":                   .Columns("E").ColumnWidth = 14
      .Cells(r_int_NroFil, 6) = "NOMBRE DEL CLIENTE":       .Columns("F").ColumnWidth = 45
      .Cells(r_int_NroFil, 7) = "CUOTA ENDOSO":             .Columns("G").ColumnWidth = 15
      .Cells(r_int_NroFil, 8) = "VCTO. CUOTA":              .Columns("H").ColumnWidth = 15
      .Cells(r_int_NroFil, 9) = "SALDO CAPITAL":            .Columns("I").ColumnWidth = 15:           .Columns("I").NumberFormat = "#,##0.00"
      .Cells(r_int_NroFil, 10) = "SUMA ASEGURADA":          .Columns("J").ColumnWidth = 18:           .Columns("J").NumberFormat = "#,##0.00"
      .Cells(r_int_NroFil, 11) = "TIPO DE COBERTURA":       .Columns("K").ColumnWidth = 25
      .Cells(r_int_NroFil, 12) = "TASA ORIGINAL":           .Columns("L").ColumnWidth = 15
      .Cells(r_int_NroFil, 13) = "COMPAÑIA DE SEGUROS":     .Columns("M").ColumnWidth = 50
      .Cells(r_int_NroFil, 14) = "FECHA APROB. LEGAL":      .Columns("N").ColumnWidth = 20
      .Cells(r_int_NroFil, 15) = "FECHA ENDOSO":            .Columns("O").ColumnWidth = 20
      .Cells(r_int_NroFil, 16) = "NUEVO NRO POLIZA":        .Columns("P").ColumnWidth = 25
      .Cells(r_int_NroFil, 17) = "NUEVA CIA. DE SEGUROS":   .Columns("Q").ColumnWidth = 50
      .Cells(r_int_NroFil, 18) = "NUEVO MONTO DE POLIZA":   .Columns("R").ColumnWidth = 25:           .Columns("R").NumberFormat = "#,##0.00"
      .Cells(r_int_NroFil, 19) = "BANCO TRANSFERIDO":       .Columns("S").ColumnWidth = 25
      .Cells(r_int_NroFil, 20) = "FECHA TRANSFERENCIA":     .Columns("T").ColumnWidth = 25
      .Range(.Cells(r_int_NroFil, 1), .Cells(r_int_NroFil, 20)).Font.Bold = True
      .Range(.Cells(r_int_NroFil, 1), .Cells(r_int_NroFil, 20)).HorizontalAlignment = xlHAlignCenter
      
      .Columns("A").HorizontalAlignment = xlHAlignCenter
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").HorizontalAlignment = xlHAlignLeft
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      .Columns("H").HorizontalAlignment = xlHAlignCenter
      .Columns("K").HorizontalAlignment = xlHAlignCenter
      .Columns("L").HorizontalAlignment = xlHAlignRight
      .Columns("M").HorizontalAlignment = xlHAlignLeft
      .Columns("N").HorizontalAlignment = xlHAlignCenter
      .Columns("O").HorizontalAlignment = xlHAlignCenter
      .Columns("P").HorizontalAlignment = xlHAlignCenter
      .Columns("S").HorizontalAlignment = xlHAlignCenter
      .Columns("T").HorizontalAlignment = xlHAlignCenter
           
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & "  SELECT SEGEND_NUMOPE AS OPERACION, "
      g_str_Parame = g_str_Parame & "         TRIM(K.PARDES_DESCRI) AS TIPO_SEGURO, "
      g_str_Parame = g_str_Parame & "         TRIM(J.PARDES_DESCRI) AS TIPO_ACCION, "
      g_str_Parame = g_str_Parame & "         TRIM(C.PARDES_DESCRI) AS SITUACION, "
      g_str_Parame = g_str_Parame & "         TRIM(DATGEN_APEPAT)||' '||TRIM(DATGEN_APEMAT)||' '||TRIM(DATGEN_NOMBRE) AS NOMBRE_CLIENTE, "
      g_str_Parame = g_str_Parame & "         SEGEND_CUOEND         AS CUOTA_ENDOSO, "
      g_str_Parame = g_str_Parame & "         SEGEND_FECVCT         AS VCTO_CUOTA, "
      g_str_Parame = g_str_Parame & "         SEGEND_SALPRE         AS SALDO_CAPITAL, "
      g_str_Parame = g_str_Parame & "         SEGEND_SUMASG         AS SUMA_ASEGURADA, "
      g_str_Parame = g_str_Parame & "         TRIM(SEGTIP_DESCRI)   AS TIPO_COBERTURA, "
      g_str_Parame = g_str_Parame & "         A.SEGEND_TASSEG       AS TASA_ORIGINAL, "
      g_str_Parame = g_str_Parame & "         TRIM(H.SEGEMP_RAZSOC) AS EMPRESA_SEGURO, "
      g_str_Parame = g_str_Parame & "         A.SEGEND_FECAPR       AS FEC_APROB_LEGAL, "
      g_str_Parame = g_str_Parame & "         A.SEGEND_FECEND       AS FEC_ENDOSO, "
      g_str_Parame = g_str_Parame & "         TRIM(A.SEGEND_NUMPOL) AS NUMERO_POLIZA, "
      g_str_Parame = g_str_Parame & "         TRIM(I.SEGEMP_RAZSOC) AS NUEVA_EMPRESA_SEGURO, "
      g_str_Parame = g_str_Parame & "         A.SEGEND_MTOPOL       AS MONTO_POLIZA, "
      g_str_Parame = g_str_Parame & "         TRIM(SUBSTR(L.PARDES_DESCRI,5)) AS BANCO_TRANFERIDO, "
      g_str_Parame = g_str_Parame & "         A.SEGEND_FECTRF       AS FECHA_TRANSFERENCIA, "
      g_str_Parame = g_str_Parame & "         A.SEGEND_TIPMON       AS MONEDA_POLIZA, "
      g_str_Parame = g_str_Parame & "         B.HIPMAE_MONEDA       AS MONEDA_CREDITO "
      g_str_Parame = g_str_Parame & "    FROM CRE_SEGEND A "
      g_str_Parame = g_str_Parame & "   INNER JOIN CRE_HIPMAE B ON B.HIPMAE_NUMOPE = A.SEGEND_NUMOPE "
      g_str_Parame = g_str_Parame & "   INNER JOIN MNT_PARDES C ON C.PARDES_CODGRP = 027 AND C.PARDES_CODITE = A.SEGEND_SITUAC "
      g_str_Parame = g_str_Parame & "   INNER JOIN CLI_DATGEN D ON D.DATGEN_TIPDOC = B.HIPMAE_TDOCLI AND D.DATGEN_NUMDOC = B.HIPMAE_NDOCLI "
      g_str_Parame = g_str_Parame & "    LEFT JOIN MNT_SEGTIP F ON F.SEGTIP_CODIGO = A.SEGEND_ESGEND AND F.SEGTIP_TIPSEG = A.SEGEND_TIPCOB "
      g_str_Parame = g_str_Parame & "    LEFT JOIN MNT_SEGEMP H ON H.SEGEMP_CODIGO = B.HIPMAE_SEGPRE "
      g_str_Parame = g_str_Parame & "    LEFT JOIN MNT_SEGEMP I ON I.SEGEMP_CODIGO = A.SEGEND_ESGEND "
      g_str_Parame = g_str_Parame & "   INNER JOIN MNT_PARDES J ON J.PARDES_CODGRP = '534' AND J.PARDES_CODITE = SEGEND_TIPACC "
      g_str_Parame = g_str_Parame & "   INNER JOIN MNT_PARDES K ON K.PARDES_CODGRP = '533' AND K.PARDES_CODITE = SEGEND_TIPSEG "
      g_str_Parame = g_str_Parame & "    LEFT JOIN MNT_PARDES L ON L.PARDES_CODGRP = '122' AND L.PARDES_CODITE = SEGEND_CODBAN "
      'g_str_Parame = g_str_Parame & "   WHERE SEGEND_FECEND >= " & r_str_FecIni & ""
      'g_str_Parame = g_str_Parame & "     AND SEGEND_FECEND <= " & r_str_FecFin & ""
      
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
         r_int_NroFil = r_int_NroFil + 1
         .Cells(r_int_NroFil, 1) = Format(r_int_NroFil - 1, "000")
         .Cells(r_int_NroFil, 2) = g_rst_Princi!OPERACION
         .Cells(r_int_NroFil, 3) = g_rst_Princi!TIPO_SEGURO
         .Cells(r_int_NroFil, 4) = g_rst_Princi!TIPO_ACCION
         .Cells(r_int_NroFil, 5) = g_rst_Princi!SITUACION
         .Cells(r_int_NroFil, 6) = g_rst_Princi!NOMBRE_CLIENTE
         
         'r_str_NumCuo = ""
         'r_str_FecVct = ""
         'Call fs_ObtienePeriodoEndoso(g_rst_Princi!OPERACION, r_str_NumCuo, r_str_FecVct)
         .Cells(r_int_NroFil, 7) = g_rst_Princi!CUOTA_ENDOSO 'r_str_NumCuo
         If Not IsNull(g_rst_Princi!VCTO_CUOTA) Then
            .Cells(r_int_NroFil, 8) = gf_FormatoFecha(CStr(g_rst_Princi!VCTO_CUOTA))  '"'" & r_str_FecVct
         End If
         
         If g_rst_Princi!MONEDA_CREDITO = 1 Then
            .Cells(r_int_NroFil, 9) = g_rst_Princi!SALDO_CAPITAL
            .Cells(r_int_NroFil, 10) = g_rst_Princi!SUMA_ASEGURADA
            .Range(.Cells(r_int_NroFil, 9), .Cells(r_int_NroFil, 10)).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
         Else
            .Cells(r_int_NroFil, 9) = g_rst_Princi!SALDO_CAPITAL
            .Cells(r_int_NroFil, 10) = g_rst_Princi!SUMA_ASEGURADA
            .Range(.Cells(r_int_NroFil, 9), .Cells(r_int_NroFil, 10)).NumberFormat = "_-[$$-80A]* #,##0.00_-;-[$$-80A]* #,##0.00_-;_-[$$-80A]* ""-""??_-;_-@_-"
         End If
         
         .Cells(r_int_NroFil, 11) = g_rst_Princi!TIPO_COBERTURA
         .Cells(r_int_NroFil, 12) = Format(g_rst_Princi!TASA_ORIGINAL, "0.000000")
         .Cells(r_int_NroFil, 13) = g_rst_Princi!EMPRESA_SEGURO
         
         .Cells(r_int_NroFil, 14) = g_rst_Princi!FEC_APROB_LEGAL
         .Cells(r_int_NroFil, 15) = g_rst_Princi!FEC_ENDOSO
         .Cells(r_int_NroFil, 16) = "'" & g_rst_Princi!NUMERO_POLIZA
         .Cells(r_int_NroFil, 17) = g_rst_Princi!NUEVA_EMPRESA_SEGURO
         
         If g_rst_Princi!MONEDA_POLIZA = 1 Then
            .Cells(r_int_NroFil, 18) = g_rst_Princi!MONTO_POLIZA
            .Cells(r_int_NroFil, 18).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
         Else
            .Cells(r_int_NroFil, 18) = g_rst_Princi!MONTO_POLIZA
            .Cells(r_int_NroFil, 18).NumberFormat = "_-[$$-80A]* #,##0.00_-;-[$$-80A]* #,##0.00_-;_-[$$-80A]* ""-""??_-;_-@_-"
         End If
         .Cells(r_int_NroFil, 19) = Trim(g_rst_Princi!BANCO_TRANFERIDO)
         If Not IsNull(g_rst_Princi!FECHA_TRANSFERENCIA) Then
            .Cells(r_int_NroFil, 20) = gf_FormatoFecha(CStr(g_rst_Princi!FECHA_TRANSFERENCIA))
         End If
         g_rst_Princi.MoveNext
      Loop
   
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub fs_ObtienePeriodoEndoso(ByVal p_NumOpe As String, ByRef p_NumCuo As String, ByRef p_FecVct As String)
Dim r_str_Parame     As String
Dim r_rst_Endoso     As ADODB.Recordset

   r_str_Parame = ""
   r_str_Parame = r_str_Parame & "SELECT MIN(HIPCUO_NUMCUO) AS NUMERO_CUOTA, MIN(HIPCUO_FECVCT) AS FECHA_VCTO "
   r_str_Parame = r_str_Parame & "  FROM CRE_HIPCUO "
   r_str_Parame = r_str_Parame & " WHERE HIPCUO_NUMOPE = '" & Trim(p_NumOpe) & "' "
   r_str_Parame = r_str_Parame & "   AND HIPCUO_TIPCRO = 1 "
   r_str_Parame = r_str_Parame & "   AND HIPCUO_DESORG = 0 "
   
   If Not gf_EjecutaSQL(r_str_Parame, r_rst_Endoso, 3) Then
      Exit Sub
   End If
   
   If r_rst_Endoso.BOF And r_rst_Endoso.EOF Then
      r_rst_Endoso.Close
      Set r_rst_Endoso = Nothing
      Exit Sub
   End If
   
   r_rst_Endoso.MoveFirst
   p_NumCuo = r_rst_Endoso!NUMERO_CUOTA
   p_FecVct = gf_FormatoFecha(CStr(r_rst_Endoso!FECHA_VCTO))

   r_rst_Endoso.Close
   Set r_rst_Endoso = Nothing
End Sub

Private Function ObtieneNomMes(ByVal mes As Integer) As String
    Select Case mes
         Case 1: ObtieneNomMes = "ENERO"
         Case 2: ObtieneNomMes = "FEBRERO"
         Case 3: ObtieneNomMes = "MARZO"
         Case 4: ObtieneNomMes = "ABRIL"
         Case 5: ObtieneNomMes = "MAYO"
         Case 6: ObtieneNomMes = "JUNIO"
         Case 7: ObtieneNomMes = "JULIO"
         Case 8: ObtieneNomMes = "AGOSTO"
         Case 9: ObtieneNomMes = "SETIEMBRE"
         Case 10: ObtieneNomMes = "OCTUBRE"
         Case 11: ObtieneNomMes = "NOVIEMBRE"
         Case 12: ObtieneNomMes = "DICIEMBRE"
      End Select
End Function

Private Sub cmb_TipRep_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmb_TipRep.ListIndex > -1 Then
         Call gs_SetFocus(cmb_PerMes)
      End If
   End If
End Sub

Private Sub cmb_PerMes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If cmb_PerMes.ListIndex > -1 Then
         Call gs_SetFocus(ipp_PerAno)
      End If
   End If
End Sub

Private Sub ipp_PerAno_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(txtFactor)
   End If
End Sub

Private Sub txtFactor_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_ExpExc)
   End If
End Sub
