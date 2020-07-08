VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frm_RptSol_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   3690
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   6255
      _Version        =   65536
      _ExtentX        =   11033
      _ExtentY        =   6509
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
      Begin Threed.SSPanel SSPanel5 
         Height          =   675
         Left            =   60
         TabIndex        =   8
         Top             =   1470
         Width           =   6135
         _Version        =   65536
         _ExtentX        =   10821
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
         Begin VB.ComboBox cmb_TipRep 
            Height          =   315
            Left            =   1470
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   180
            Width           =   4485
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo de Reporte:"
            Height          =   255
            Left            =   60
            TabIndex        =   9
            Top             =   210
            Width           =   1380
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   60
         TabIndex        =   10
         Top             =   60
         Width           =   6135
         _Version        =   65536
         _ExtentX        =   10821
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
         Begin Threed.SSPanel ssp_TipRep 
            Height          =   495
            Left            =   660
            TabIndex        =   11
            Top             =   135
            Width           =   5145
            _Version        =   65536
            _ExtentX        =   9075
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Reporte de Solicitudes"
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
         Begin Threed.SSPanel ssp_TipRep1 
            Height          =   315
            Left            =   4980
            TabIndex        =   12
            Top             =   300
            Visible         =   0   'False
            Width           =   1005
            _Version        =   65536
            _ExtentX        =   1773
            _ExtentY        =   556
            _StockProps     =   15
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
            Picture         =   "OpeTra_frm_811.frx":0000
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   60
         TabIndex        =   13
         Top             =   780
         Width           =   6135
         _Version        =   65536
         _ExtentX        =   10821
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
            Left            =   5520
            Picture         =   "OpeTra_frm_811.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   585
            Left            =   630
            Picture         =   "OpeTra_frm_811.frx":074C
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Imprimir Reporte"
            Top             =   30
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_811.frx":0B8E
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin Crystal.CrystalReport crp_Imprim 
            Left            =   1500
            Top             =   150
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
         Height          =   510
         Left            =   60
         TabIndex        =   14
         Top             =   2205
         Width           =   6135
         _Version        =   65536
         _ExtentX        =   10821
         _ExtentY        =   900
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
         Begin VB.ComboBox cmb_CodIns 
            Height          =   315
            Left            =   1470
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   90
            Width           =   4485
         End
         Begin VB.Label lbl_TipCon 
            Caption         =   "Instancia:"
            Height          =   255
            Left            =   60
            TabIndex        =   15
            Top             =   120
            Width           =   1005
         End
      End
      Begin Threed.SSPanel pnlfecha 
         Height          =   870
         Left            =   60
         TabIndex        =   16
         Top             =   2760
         Width           =   6135
         _Version        =   65536
         _ExtentX        =   10821
         _ExtentY        =   1535
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
         Enabled         =   0   'False
         Begin EditLib.fpDateTime ipp_FecIni 
            Height          =   315
            Left            =   1470
            TabIndex        =   2
            Top             =   90
            Width           =   1965
            _Version        =   196608
            _ExtentX        =   3466
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
            ButtonStyle     =   3
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   0
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
            Text            =   "28/09/2004"
            DateCalcMethod  =   0
            DateTimeFormat  =   0
            UserDefinedFormat=   ""
            DateMax         =   "00000000"
            DateMin         =   "00000000"
            TimeMax         =   "000000"
            TimeMin         =   "000000"
            TimeString1159  =   ""
            TimeString2359  =   ""
            DateDefault     =   "00000000"
            TimeDefault     =   "000000"
            TimeStyle       =   0
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   2
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            PopUpType       =   0
            DateCalcY2KSplit=   60
            CaretPosition   =   0
            IncYear         =   1
            IncMonth        =   1
            IncDay          =   1
            IncHour         =   1
            IncMinute       =   1
            IncSecond       =   1
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpDateTime ipp_FecFin 
            Height          =   315
            Left            =   1470
            TabIndex        =   3
            Top             =   450
            Width           =   1965
            _Version        =   196608
            _ExtentX        =   3466
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
            ButtonStyle     =   3
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   0
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
            Text            =   "28/09/2004"
            DateCalcMethod  =   0
            DateTimeFormat  =   0
            UserDefinedFormat=   ""
            DateMax         =   "00000000"
            DateMin         =   "00000000"
            TimeMax         =   "000000"
            TimeMin         =   "000000"
            TimeString1159  =   ""
            TimeString2359  =   ""
            DateDefault     =   "00000000"
            TimeDefault     =   "000000"
            TimeStyle       =   0
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   2
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            PopUpType       =   0
            DateCalcY2KSplit=   60
            CaretPosition   =   0
            IncYear         =   1
            IncMonth        =   1
            IncDay          =   1
            IncHour         =   1
            IncMinute       =   1
            IncSecond       =   1
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin VB.Label Label2 
            Caption         =   "Fecha Inicio:"
            Height          =   255
            Left            =   60
            TabIndex        =   18
            Top             =   150
            Width           =   1065
         End
         Begin VB.Label Label3 
            Caption         =   "Fecha Fin:"
            Height          =   225
            Left            =   60
            TabIndex        =   17
            Top             =   480
            Width           =   1035
         End
      End
   End
End
Attribute VB_Name = "frm_RptSol_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmb_CodIns_Click()
   Call gs_SetFocus(ipp_FecIni)
   If cmb_TipRep.ListIndex > -1 Then
      ssp_TipRep.Caption = "Reporte de Solicitudes " & Trim(Replace(LCase(cmb_TipRep.Text), "solicitudes", "")) & " en " & LCase(cmb_CodIns.Text)
   End If
End Sub

Private Sub cmb_TipIns_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_ExpExc)
   End If
End Sub
Private Sub cmb_TipRep_Click()
   If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 4 Or cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 5 Then
      pnlfecha.Enabled = True
   Else
      pnlfecha.Enabled = False
   End If
   If cmb_TipRep.ListIndex > -1 Then
      ssp_TipRep.Caption = "Reporte de Solicitudes " & Trim(Replace(LCase(cmb_TipRep.Text), "solicitudes", ""))
   End If
   cmb_CodIns.ListIndex = -1
   Call gs_SetFocus(cmb_CodIns)
End Sub
Private Sub cmb_TipRep_KeyPress(KeyAscii As Integer)
   Call gs_SetFocus(cmb_CodIns)
End Sub
Private Sub cmd_ExpExc_Click()
   If cmb_TipRep.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Reporte.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipRep)
      Exit Sub
   End If

   If cmb_CodIns.ListIndex = -1 Then
      MsgBox "Debe seleccionar Instancia.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_CodIns)
      Exit Sub
   End If

   If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 4 Or cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 5 Then
      If CDate(ipp_FecIni.Text) > CDate(ipp_FecFin.Text) Then
         MsgBox "Fecha de Inicio no puede ser mayor a la Fecha Final", vbExclamation, modgen_g_str_NomPlt
         Call gs_SetFocus(ipp_FecIni)
         Exit Sub
      End If
   End If
   
   'Confirmación
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   g_str_Parame = ""
   g_str_Parame = "USP_RPT_REPSOL ( "
   If cmb_CodIns.ItemData(cmb_CodIns.ListIndex) = 41 Or cmb_CodIns.ItemData(cmb_CodIns.ListIndex) = 42 Then
      g_str_Parame = g_str_Parame & "'41',"
   Else
      g_str_Parame = g_str_Parame & "'61',"
   End If
   g_str_Parame = g_str_Parame & "" & cmb_CodIns.ItemData(cmb_CodIns.ListIndex) & ","
   g_str_Parame = g_str_Parame & "" & cmb_TipRep.ItemData(cmb_TipRep.ListIndex) & ","
   g_str_Parame = g_str_Parame & "'" & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & "',"
   g_str_Parame = g_str_Parame & "'" & Format(CDate(ipp_FecFin.Text), "yyyymmdd") & "',"
   g_str_Parame = g_str_Parame & "" & Month(Now) & ", "
   g_str_Parame = g_str_Parame & "" & Year(Now) & ","
   g_str_Parame = g_str_Parame & "'REPORTE DE SOLICITUDES'" & ","
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_NombPC & "',"
   g_str_Parame = g_str_Parame & "'" & modgen_g_str_CodUsu & "')"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
         MsgBox "Error al ejecutar procedimiento USP_RPT_REPSOL.", vbCritical, modgen_g_str_NomPlt
         Exit Sub
   End If
  
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      Call fs_GenExc(g_rst_Princi, cmb_TipRep.ItemData(cmb_TipRep.ListIndex))
   End If
    
End Sub
Private Sub fs_GenExc(ByVal g_rst_Princi As Object, ByVal TipRep As Integer)
Dim r_obj_Excel      As Excel.Application
Dim r_int_ConVer     As Integer

   'Generando Excel
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.DisplayAlerts = False
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      If TipRep = 1 Or TipRep = 2 Then
         .Cells(1, 1) = "ITEM":                          .Columns("A").ColumnWidth = 8
         .Cells(1, 2) = "PRODUCTO":                      .Columns("B").ColumnWidth = 40
         .Cells(1, 3) = "SOLICITUD":                     .Columns("C").ColumnWidth = 15:        .Columns("C").HorizontalAlignment = xlHAlignCenter
         .Cells(1, 4) = "DOC. IDENTIDAD":                .Columns("D").ColumnWidth = 15:        .Columns("D").HorizontalAlignment = xlHAlignCenter
         .Cells(1, 5) = "NOMBRE CLIENTE":                .Columns("E").ColumnWidth = 40
         .Cells(1, 6) = "F. SOLICITUD":                  .Columns("F").ColumnWidth = 15:        .Columns("F").HorizontalAlignment = xlHAlignCenter
         .Cells(1, 7) = "F. INGRESO INSTANCIA":          .Columns("G").ColumnWidth = 21:        .Columns("G").HorizontalAlignment = xlHAlignCenter
         .Cells(1, 8) = "TASA":                          .Columns("H").ColumnWidth = 15
         .Cells(1, 9) = "PLAZO":                         .Columns("I").ColumnWidth = 15
         .Cells(1, 10) = "NRO. DE ACTA":                 .Columns("J").ColumnWidth = 20:        .Columns("J").HorizontalAlignment = xlHAlignCenter
         .Cells(1, 11) = "TIPO EVALUACION":              .Columns("K").ColumnWidth = 40
         .Cells(1, 12) = "CONSEJERO HIPOTECARIO":        .Columns("L").ColumnWidth = 30:        .Columns("L").HorizontalAlignment = xlHAlignCenter
         .Cells(1, 13) = "INSTANCIA":                    .Columns("M").ColumnWidth = 30:        .Columns("M").HorizontalAlignment = xlHAlignCenter
         .Cells(1, 14) = "TIEMPO INSTANCIA":             .Columns("N").ColumnWidth = 20
         .Cells(1, 15) = "TIEMPO OBSERVADO":             .Columns("O").ColumnWidth = 20
         .Cells(1, 16) = "TIEMPO EVALUACION":            .Columns("P").ColumnWidth = 20
         .Cells(1, 17) = "MONEDA":                       .Columns("Q").ColumnWidth = 15:        .Columns("Q").HorizontalAlignment = xlHAlignCenter
         .Cells(1, 18) = "V. INMUEBLE":                  .Columns("R").ColumnWidth = 20
         .Cells(1, 19) = "CUOTA INICIAL":                .Columns("S").ColumnWidth = 20
         .Cells(1, 20) = "PORC. INICIAL":                .Columns("T").ColumnWidth = 20
         .Cells(1, 21) = "MTO. CREDITO S/.":             .Columns("U").ColumnWidth = 20
         .Cells(1, 22) = "MTO. CREDITO US$":             .Columns("V").ColumnWidth = 20
         .Cells(1, 23) = "SITUACION INSTANCIA":          .Columns("W").ColumnWidth = 30
         .Cells(1, 24) = "MODALIDAD":                    .Columns("X").ColumnWidth = 60
         .Cells(1, 25) = "PROYECTO INMOBILIARIO":        .Columns("Y").ColumnWidth = 120
         .Cells(1, 26) = "OBSERVACION":                  .Columns("Z").ColumnWidth = 200
         
         .Range(.Cells(1, 1), .Cells(1, 26)).Font.Bold = True
         .Range(.Cells(1, 1), .Cells(1, 26)).HorizontalAlignment = xlHAlignCenter
      ElseIf TipRep = 3 Then
         .Cells(1, 1) = "ITEM":                          .Columns("A").ColumnWidth = 8
         .Cells(1, 2) = "PRODUCTO":                      .Columns("B").ColumnWidth = 40
         .Cells(1, 3) = "SUB-PRODUCTO":                  .Columns("C").ColumnWidth = 70
         .Cells(1, 4) = "SOLICITUD":                     .Columns("D").ColumnWidth = 15:        .Columns("D").HorizontalAlignment = xlHAlignCenter
         .Cells(1, 5) = "DOC. IDENTIDAD":                .Columns("E").ColumnWidth = 15:        .Columns("E").HorizontalAlignment = xlHAlignCenter
         .Cells(1, 6) = "NOMBRE CLIENTE":                .Columns("F").ColumnWidth = 40
         .Cells(1, 7) = "F. SOLICITUD":                  .Columns("G").ColumnWidth = 20:        .Columns("G").HorizontalAlignment = xlHAlignCenter
         .Cells(1, 8) = "F. APROB. CONDIC.":             .Columns("H").ColumnWidth = 20:        .Columns("H").HorizontalAlignment = xlHAlignCenter
         .Cells(1, 9) = "TASA":                          .Columns("I").ColumnWidth = 15
         .Cells(1, 10) = "PLAZO":                        .Columns("J").ColumnWidth = 15
         .Cells(1, 11) = "NRO. DE ACTA":                 .Columns("K").ColumnWidth = 20:        .Columns("K").HorizontalAlignment = xlHAlignCenter
         .Cells(1, 12) = "TIPO EVALUACION":              .Columns("L").ColumnWidth = 40
         .Cells(1, 13) = "CONSEJERO HIPOTECARIO":        .Columns("M").ColumnWidth = 30:        .Columns("M").HorizontalAlignment = xlHAlignCenter
         .Cells(1, 14) = "INSTANCIA":                    .Columns("N").ColumnWidth = 30:        .Columns("N").HorizontalAlignment = xlHAlignCenter
         .Cells(1, 15) = "TIEMPO APROBAC. CONDIC.":      .Columns("O").ColumnWidth = 25
         .Cells(1, 16) = "MONEDA":                       .Columns("P").ColumnWidth = 15:        .Columns("P").HorizontalAlignment = xlHAlignCenter
         .Cells(1, 17) = "V. INMUEBLE":                  .Columns("Q").ColumnWidth = 20
         .Cells(1, 18) = "CUOTA INICIAL":                .Columns("R").ColumnWidth = 20
         .Cells(1, 19) = "PORC. INICIAL":                .Columns("S").ColumnWidth = 20
         .Cells(1, 20) = "MTO. CREDITO S/.":             .Columns("T").ColumnWidth = 20
         .Cells(1, 21) = "MTO. CREDITO US$":             .Columns("U").ColumnWidth = 20
         .Cells(1, 22) = "CONDICIONES":                  .Columns("V").ColumnWidth = 200
         
         .Range(.Cells(1, 1), .Cells(1, 22)).Font.Bold = True
         .Range(.Cells(1, 1), .Cells(1, 22)).HorizontalAlignment = xlHAlignCenter
      ElseIf TipRep = 4 Then
         .Cells(1, 1) = "ITEM":                          .Columns("A").ColumnWidth = 8
         .Cells(1, 2) = "PRODUCTO":                      .Columns("B").ColumnWidth = 40
         .Cells(1, 3) = "SOLICITUD":                     .Columns("C").ColumnWidth = 15:        .Columns("C").HorizontalAlignment = xlHAlignCenter
         .Cells(1, 4) = "OPERACION":                     .Columns("D").ColumnWidth = 15:        .Columns("D").HorizontalAlignment = xlHAlignCenter
         .Cells(1, 5) = "DOC. IDENTIDAD":                .Columns("E").ColumnWidth = 15:        .Columns("E").HorizontalAlignment = xlHAlignCenter
         .Cells(1, 6) = "NOMBRE CLIENTE":                .Columns("F").ColumnWidth = 40
         .Cells(1, 7) = "F. SOLICITUD":                  .Columns("G").ColumnWidth = 20:        .Columns("G").HorizontalAlignment = xlHAlignCenter
         .Cells(1, 8) = "F. EXCEPCION":                  .Columns("H").ColumnWidth = 20:        .Columns("H").HorizontalAlignment = xlHAlignCenter
         .Cells(1, 9) = "TASA":                          .Columns("I").ColumnWidth = 15
         .Cells(1, 10) = "PLAZO":                        .Columns("J").ColumnWidth = 15
         .Cells(1, 11) = "NRO. DE ACTA":                 .Columns("K").ColumnWidth = 20:        .Columns("K").HorizontalAlignment = xlHAlignCenter
         .Cells(1, 12) = "TIPO EVALUACION":              .Columns("L").ColumnWidth = 40
         .Cells(1, 13) = "CONSEJERO HIPOTECARIO":        .Columns("M").ColumnWidth = 30:        .Columns("M").HorizontalAlignment = xlHAlignCenter
         .Cells(1, 14) = "INSTANCIA":                    .Columns("N").ColumnWidth = 30:        .Columns("N").HorizontalAlignment = xlHAlignCenter
         .Cells(1, 15) = "MONEDA":                       .Columns("O").ColumnWidth = 15:        .Columns("O").HorizontalAlignment = xlHAlignCenter
         .Cells(1, 16) = "V. INMUEBLE":                  .Columns("P").ColumnWidth = 20
         .Cells(1, 17) = "CUOTA INICIAL":                .Columns("Q").ColumnWidth = 20
         .Cells(1, 18) = "PORC. INICIAL":                .Columns("R").ColumnWidth = 20
         .Cells(1, 19) = "MTO. CREDITO S/.":             .Columns("S").ColumnWidth = 20
         .Cells(1, 20) = "MTO. CREDITO US$":             .Columns("T").ColumnWidth = 20
         .Cells(1, 21) = "SITUAC. SOLIC.":               .Columns("U").ColumnWidth = 20
         .Cells(1, 22) = "AUTORIZACION":                 .Columns("V").ColumnWidth = 40
         .Cells(1, 23) = "MOTIVO EXCEPCION":             .Columns("W").ColumnWidth = 50
         .Cells(1, 24) = "DESCRIPCION EXCEPCION":        .Columns("X").ColumnWidth = 200
         
         .Range(.Cells(1, 1), .Cells(1, 24)).Font.Bold = True
         .Range(.Cells(1, 1), .Cells(1, 24)).HorizontalAlignment = xlHAlignCenter
         
      ElseIf TipRep = 5 Then
         .Cells(1, 1) = "ITEM":                          .Columns("A").ColumnWidth = 6:         .Columns("A").HorizontalAlignment = xlHAlignCenter
         .Cells(1, 2) = "PRODUCTO":                      .Columns("B").ColumnWidth = 40:        .Columns("B").HorizontalAlignment = xlHAlignCenter
         .Cells(1, 3) = "SUB-PRODUCTO":                  .Columns("C").ColumnWidth = 65:        .Columns("C").HorizontalAlignment = xlHAlignCenter
         .Cells(1, 4) = "SOLICITUD":                     .Columns("D").ColumnWidth = 15:        .Columns("D").HorizontalAlignment = xlHAlignCenter
         .Cells(1, 5) = "DOC. IDENTIDAD":                .Columns("E").ColumnWidth = 15:        .Columns("E").HorizontalAlignment = xlHAlignCenter
         .Cells(1, 6) = "NOMBRE CLIENTE":                .Columns("F").ColumnWidth = 44
         .Cells(1, 7) = "F. SOLICITUD":                  .Columns("G").ColumnWidth = 12:        .Columns("G").HorizontalAlignment = xlHAlignCenter
         .Cells(1, 8) = "TASA":                          .Columns("H").ColumnWidth = 10:
         .Cells(1, 9) = "PLAZO":                         .Columns("I").ColumnWidth = 10:        .Columns("I").HorizontalAlignment = xlHAlignCenter
         .Cells(1, 10) = "NRO. DE ACTA":                 .Columns("J").ColumnWidth = 13:        .Columns("J").HorizontalAlignment = xlHAlignCenter
         .Cells(1, 11) = "TIPO EVALUACION":              .Columns("K").ColumnWidth = 22:        .Columns("K").HorizontalAlignment = xlHAlignCenter
         .Cells(1, 12) = "CONSEJERO HIPOTECARIO":        .Columns("L").ColumnWidth = 24:        .Columns("L").HorizontalAlignment = xlHAlignCenter
         .Cells(1, 13) = "INSTANCIA":                    .Columns("M").ColumnWidth = 34:        .Columns("M").HorizontalAlignment = xlHAlignCenter
         .Cells(1, 14) = "RESULTADO EVALUACION":         .Columns("N").ColumnWidth = 24:        .Columns("N").HorizontalAlignment = xlHAlignCenter
         .Cells(1, 15) = "INGRESO INSTANCIA":            .Columns("O").ColumnWidth = 18:        .Columns("O").HorizontalAlignment = xlHAlignCenter
         .Cells(1, 16) = "SALIDA INSTANCIA":             .Columns("P").ColumnWidth = 18:        .Columns("P").HorizontalAlignment = xlHAlignCenter
         .Cells(1, 17) = "TIEMPO INSTANCIA":             .Columns("Q").ColumnWidth = 18:        .Columns("Q").HorizontalAlignment = xlHAlignCenter
         .Cells(1, 18) = "TIEMPO OBSERVADO":             .Columns("R").ColumnWidth = 18:        .Columns("R").HorizontalAlignment = xlHAlignCenter
         .Cells(1, 19) = "TIEMPO EVALUACION":            .Columns("S").ColumnWidth = 20:        .Columns("S").HorizontalAlignment = xlHAlignCenter
         .Cells(1, 20) = "MONEDA":                       .Columns("T").ColumnWidth = 10:        .Columns("T").HorizontalAlignment = xlHAlignCenter
         .Cells(1, 21) = "V. INMUEBLE":                  .Columns("U").ColumnWidth = 12
         .Cells(1, 22) = "CUOTA INICIAL":                .Columns("V").ColumnWidth = 14
         .Cells(1, 23) = "PORC. INICIAL":                .Columns("W").ColumnWidth = 13
         .Cells(1, 24) = "MTO. CREDITO S/.":             .Columns("X").ColumnWidth = 18
         .Cells(1, 25) = "MTO. CREDITO US$":             .Columns("Y").ColumnWidth = 18
         .Cells(1, 26) = "MODALIDAD":                    .Columns("Z").ColumnWidth = 50
         .Cells(1, 27) = "PROYECTO INMOBILIARIO":        .Columns("AA").ColumnWidth = 50
         .Cells(1, 28) = "MOTIVO RECHAZO":               .Columns("AB").ColumnWidth = 140
         .Cells(1, 29) = "COD. EXCEP.":                  .Columns("AC").ColumnWidth = 15:       .Columns("AC").HorizontalAlignment = xlHAlignCenter
         .Cells(1, 30) = "EXCEPCION":                    .Columns("AD").ColumnWidth = 80
         .Cells(1, 31) = "OBSERVACION EXCEPCION":        .Columns("AE").ColumnWidth = 100
         .Cells(1, 32) = "CONDICIONES":                  .Columns("AF").ColumnWidth = 100
         .Cells(1, 33) = "SITUACION ACTUAL":             .Columns("AG").ColumnWidth = 50:       .Columns("AG").HorizontalAlignment = xlHAlignCenter
         
         .Range(.Cells(1, 1), .Cells(1, 33)).Font.Bold = True
         .Range(.Cells(1, 1), .Cells(1, 33)).HorizontalAlignment = xlHAlignCenter
      End If
         
   End With
     
   r_int_ConVer = 2
   
   g_rst_Princi.MoveFirst
   
   Do While Not g_rst_Princi.EOF
      If TipRep = 1 Or TipRep = 2 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 1
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = Trim(g_rst_Princi!PRODUCTO)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = gf_Formato_NumSol(g_rst_Princi!SOLICITUD)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = CStr(g_rst_Princi!DOC_IDENTIDAD)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = Trim(g_rst_Princi!NOMBRE_CLIENTE)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!FEC_SOLICITUD)))
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!FEC_INGRESO_INSTANCIA)))
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = Trim(g_rst_Princi!TASA)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = Trim(g_rst_Princi!PLAZO)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = "'" & Trim(g_rst_Princi!NRO_ACTA)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = Trim(g_rst_Princi!TIPO_EVALUACION)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = Trim(g_rst_Princi!CONSEJERO)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = Trim(g_rst_Princi!INSTANCIA)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14) = Trim(g_rst_Princi!TIEMPO_INSTANCIA)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = Trim(g_rst_Princi!TIEMPO_OBSERVADO)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 16) = Trim(g_rst_Princi!TIEMPO_EVALUACION)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 17) = Trim(g_rst_Princi!moneda)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 18) = Format(g_rst_Princi!V_INMUEBLE, "###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 19) = Format(g_rst_Princi!CUOTA_INICIAL, "###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 20) = Format(g_rst_Princi!PORC_INICIAL, "##0.00") & "%"
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 21) = Format(g_rst_Princi!MTO_CRED_SOL, "###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 22) = Format(g_rst_Princi!MTO_CRED_DOL, "###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 23) = Trim(g_rst_Princi!SITUACION_INSTANCIA)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 24) = Trim(g_rst_Princi!MODALIDAD)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 25) = Trim(g_rst_Princi!PROYECTO_INMOBILIARIO)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 26) = Trim(g_rst_Princi!OBSERVACION)
         
      ElseIf TipRep = 3 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 1
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = Trim(g_rst_Princi!PRODUCTO)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = Trim(g_rst_Princi!SUB_PRODUCTO)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = gf_Formato_NumSol(g_rst_Princi!SOLICITUD)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = CStr(g_rst_Princi!DOC_IDENTIDAD)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = Trim(g_rst_Princi!NOMBRE_CLIENTE)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!FEC_SOLICITUD)))
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!FEC_APROB_CONDIC)))
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = Trim(g_rst_Princi!TASA)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = Trim(g_rst_Princi!PLAZO)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = Trim(g_rst_Princi!NRO_ACTA)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = Trim(g_rst_Princi!TIPO_EVALUACION)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = Trim(g_rst_Princi!CONSEJERO)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14) = Trim(g_rst_Princi!INSTANCIA)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = Trim(g_rst_Princi!TIEMPO_APROB_COND)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 16) = Trim(g_rst_Princi!moneda)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 17) = Format(g_rst_Princi!V_INMUEBLE, "###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 18) = Format(g_rst_Princi!CUOTA_INICIAL, "###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 19) = Format(g_rst_Princi!PORC_INICIAL, "##0.00") & "%"
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 20) = Format(g_rst_Princi!MTO_CRED_SOL, "###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 21) = Format(g_rst_Princi!MTO_CRED_DOL, "###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 22) = Trim(g_rst_Princi!CONDICION)
   
      ElseIf TipRep = 4 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 1
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = Trim(g_rst_Princi!PRODUCTO)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = gf_Formato_NumSol(g_rst_Princi!SOLICITUD)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = Trim(g_rst_Princi!OPERACION)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = CStr(g_rst_Princi!DOC_IDENTIDAD)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = Trim(g_rst_Princi!NOMBRE_CLIENTE)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!FEC_SOLICITUD)))
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!FEC_EXCEPCION)))
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = Trim(g_rst_Princi!TASA)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = Trim(g_rst_Princi!PLAZO)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = Trim(g_rst_Princi!NRO_ACTA)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = Trim(g_rst_Princi!TIPO_EVALUACION)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = Trim(g_rst_Princi!CONSEJERO)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14) = Trim(g_rst_Princi!INSTANCIA)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = Trim(g_rst_Princi!moneda)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 16) = Format(g_rst_Princi!V_INMUEBLE, "###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 17) = Format(g_rst_Princi!CUOTA_INICIAL, "###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 18) = Format(g_rst_Princi!PORC_INICIAL, "##0.00") & "%"
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 19) = Format(g_rst_Princi!MTO_CRED_SOL, "###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 20) = Format(g_rst_Princi!MTO_CRED_DOL, "###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 21) = Trim(g_rst_Princi!SITUAC_SOLIC)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 22) = Trim(g_rst_Princi!AUTORIZACION)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 23) = Trim(g_rst_Princi!MOTIVO_EXCEPCION)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 24) = Trim(g_rst_Princi!DESCRIPCION_EXCEPCION)
         
      ElseIf TipRep = 5 Then
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_ConVer - 1
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = Trim(g_rst_Princi!PRODUCTO)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = Trim(g_rst_Princi!SUB_PRODUCTO)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = gf_Formato_NumSol(g_rst_Princi!SOLICITUD)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = CStr(g_rst_Princi!DOC_IDENTIDAD)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = Trim(g_rst_Princi!NOMBRE_CLIENTE)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!FEC_SOLICITUD)))
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = Trim(g_rst_Princi!TASA)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = Trim(g_rst_Princi!PLAZO)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 10) = Trim(g_rst_Princi!NRO_ACTA)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 11) = Trim(g_rst_Princi!TIPO_EVALUACION)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 12) = Trim(g_rst_Princi!CONSEJERO)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 13) = Trim(g_rst_Princi!INSTANCIA)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 14) = Trim(g_rst_Princi!RESULTADO_EVAL)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 15) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!INGRESO_INSTANCIA)))
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 16) = CDate(gf_FormatoFecha(CStr(g_rst_Princi!SALIDA_INSTANCIA)))
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 17) = Trim(g_rst_Princi!TIEMPO_INSTANCIA)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 18) = Trim(g_rst_Princi!TIEMPO_OBSERVADO)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 19) = Trim(g_rst_Princi!TIEMPO_EVALUACION)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 20) = Trim(g_rst_Princi!moneda)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 21) = Format(g_rst_Princi!V_INMUEBLE, "###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 22) = Format(g_rst_Princi!CUOTA_INICIAL, "###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 23) = Format(g_rst_Princi!PORC_INICIAL, "##0.00") & "%"
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 24) = Format(g_rst_Princi!MTO_CRED_SOL, "###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 25) = Format(g_rst_Princi!MTO_CRED_DOL, "###,##0.00")
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 26) = Trim(g_rst_Princi!MODALIDAD)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 27) = g_rst_Princi!PROYECTO_INMOBILIARIO
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 28) = Trim(g_rst_Princi!MOT_RECHAZO)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 29) = Trim(g_rst_Princi!CODEXP)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 30) = Trim(g_rst_Princi!EXCEPCION)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 31) = Trim(g_rst_Princi!OBSERV_EXCEP)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 32) = Trim(g_rst_Princi!CONDICION)
         r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 33) = Trim(g_rst_Princi!SITUACION_ACTUAL)
      End If
      r_int_ConVer = r_int_ConVer + 1
      g_rst_Princi.MoveNext
   Loop
   
   g_rst_Princi.Close
   
   Set g_rst_Princi = Nothing
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub
Private Sub cmd_Salida_Click()
   Unload Me
End Sub
Private Sub Form_Load()
   Me.Caption = modgen_g_str_NomPlt
   Call fs_Limpia
   
   cmb_TipRep.AddItem "SOLICITUDES VIGENTES"
   cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = 1
   cmb_TipRep.AddItem "SOLICITUDES CON OBSERVACION PENDIENTE"
   cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = 2
   cmb_TipRep.AddItem "SOLICITUDES CON APROBACION CONDICIONADA"
   cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = 3
   cmb_TipRep.AddItem "SOLICITUDES CON EXCEPCION APROBADA"
   cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = 4
   cmb_TipRep.AddItem "SOLICITUDES EVALUADAS"
   cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = 5
   cmb_TipRep.ListIndex = -1
   
   cmb_CodIns.AddItem "TASACION DEL INMUEBLE"
   cmb_CodIns.ItemData(cmb_CodIns.NewIndex) = 41
   cmb_CodIns.AddItem "EVALUACION DE SEGUROS"
   cmb_CodIns.ItemData(cmb_CodIns.NewIndex) = 42
   cmb_CodIns.AddItem "POLIZAS DE SEGURO"
   cmb_CodIns.ItemData(cmb_CodIns.NewIndex) = 61
   cmb_CodIns.AddItem "TRAMITES COFIDE"
   cmb_CodIns.ItemData(cmb_CodIns.NewIndex) = 62
   cmb_CodIns.ListIndex = -1

   ipp_FecIni.Text = Format(date - CDate(30), "dd/mm/yyyy")
   ipp_FecFin.Text = Format(date, "dd/mm/yyyy")
   
   Call gs_SetFocus(cmb_TipRep)
   Call gs_CentraForm(Me)
   
End Sub
Private Sub ipp_FecFin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Imprim)
   End If
End Sub
Private Sub ipp_FecIni_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecFin)
   End If
End Sub
Private Sub fs_Limpia()
   ipp_FecIni.Text = Format(date, "dd/mm/yyyy")
   ipp_FecFin.Text = Format(date, "dd/mm/yyyy")
   cmb_TipRep.Clear
   cmb_CodIns.Clear
End Sub

'Private Sub cmd_Imprim_Click()
'   If cmb_TipRep.ListIndex = -1 Then
'      MsgBox "Debe seleccionar el Tipo de Reporte.", vbExclamation, modgen_g_str_NomPlt
'      Call gs_SetFocus(cmb_TipRep)
'      Exit Sub
'   End If
'
'   If cmb_TipIns.ListIndex = -1 Then
'      MsgBox "Debe seleccionar Instancia.", vbExclamation, modgen_g_str_NomPlt
'      Call gs_SetFocus(cmb_TipIns)
'      Exit Sub
'   End If
'
'   If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 4 Or cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 5 Then
'      If CDate(ipp_FecIni.Text) > CDate(ipp_FecFin.Text) Then
'         MsgBox "Fecha de Inicio no puede ser mayor a la Fecha Final", vbExclamation, modgen_g_str_NomPlt
'         Call gs_SetFocus(ipp_FecIni)
'         Exit Sub
'      End If
'   End If
'
'   If MsgBox("¿Está seguro de imprimir el reporte?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
'      Exit Sub
'   End If
'
'   Screen.MousePointer = 11
'
'   If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 1 Then
'      If cmb_TipIns.ItemData(cmb_TipIns.ListIndex) = 1 Then
'         Call modmip_gs_Rpt_EvaIns_Dbl("CRE_EVAHIP_01.RPT", 1, "", 41, 41, "")
'      ElseIf cmb_TipIns.ItemData(cmb_TipIns.ListIndex) = 2 Then
'         Call modmip_gs_Rpt_EvaIns_Dbl("CRE_EVAHIP_01.RPT", 1, "", 41, 42, "")
'      ElseIf cmb_TipIns.ItemData(cmb_TipIns.ListIndex) = 3 Then
'         Call modmip_gs_Rpt_EvaIns_Dbl("CRE_EVAHIP_01.RPT", 1, "", 61, 61, "")
'      ElseIf cmb_TipIns.ItemData(cmb_TipIns.ListIndex) = 4 Then
'         Call modmip_gs_Rpt_EvaIns_Dbl("CRE_EVAHIP_01.RPT", 1, "", 61, 62, "")
'      End If
'
'   ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 2 Then
'      If cmb_TipIns.ItemData(cmb_TipIns.ListIndex) = 1 Then
'         Call modmip_gs_Rpt_EvaObs_Dbl("CRE_EVAHIP_04.RPT", 1, "", 41, 41, "")
'      ElseIf cmb_TipIns.ItemData(cmb_TipIns.ListIndex) = 2 Then
'         Call modmip_gs_Rpt_EvaObs_Dbl("CRE_EVAHIP_04.RPT", 1, "", 41, 42, "")
'      ElseIf cmb_TipIns.ItemData(cmb_TipIns.ListIndex) = 3 Then
'         Call modmip_gs_Rpt_EvaObs_Dbl("CRE_EVAHIP_04.RPT", 1, "", 61, 61, "")
'      ElseIf cmb_TipIns.ItemData(cmb_TipIns.ListIndex) = 4 Then
'         Call modmip_gs_Rpt_EvaObs_Dbl("CRE_EVAHIP_04.RPT", 1, "", 61, 62, "")
'      End If
'
'   ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 3 Then
'      If cmb_TipIns.ItemData(cmb_TipIns.ListIndex) = 1 Then
'         Call modmip_gs_Rpt_AprCon("CRE_EVAHIP_06.RPT", 1, "", 41, "")
'      ElseIf cmb_TipIns.ItemData(cmb_TipIns.ListIndex) = 2 Then
'         Call modmip_gs_Rpt_AprCon("CRE_EVAHIP_06.RPT", 1, "", 42, "")
'      ElseIf cmb_TipIns.ItemData(cmb_TipIns.ListIndex) = 3 Then
'         Call modmip_gs_Rpt_AprCon("CRE_EVAHIP_06.RPT", 1, "", 61, "")
'      ElseIf cmb_TipIns.ItemData(cmb_TipIns.ListIndex) = 4 Then
'         Call modmip_gs_Rpt_AprCon("CRE_EVAHIP_06.RPT", 1, "", 62, "")
'      End If
'
'   'ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 4 Then
'
'   ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 5 Then
'      If cmb_TipIns.ItemData(cmb_TipIns.ListIndex) = 1 Then
'         Call modmip_gs_Rpt_SolEva("CRE_EVAHIP_10.RPT", 1, "", 41, "", 0, ipp_FecIni.Text, ipp_FecFin.Text)
'      ElseIf cmb_TipIns.ItemData(cmb_TipIns.ListIndex) = 2 Then
'         Call modmip_gs_Rpt_SolEva("CRE_EVAHIP_10.RPT", 1, "", 42, "", 0, ipp_FecIni.Text, ipp_FecFin.Text)
'      ElseIf cmb_TipIns.ItemData(cmb_TipIns.ListIndex) = 3 Then
'         Call modmip_gs_Rpt_SolEva("CRE_EVAHIP_10.RPT", 1, "", 61, "", 0, ipp_FecIni.Text, ipp_FecFin.Text)
'      ElseIf cmb_TipIns.ItemData(cmb_TipIns.ListIndex) = 4 Then
'         Call modmip_gs_Rpt_SolEva("CRE_EVAHIP_10.RPT", 1, "", 62, "", 0, ipp_FecIni.Text, ipp_FecFin.Text)
'      End If
'   End If
'
'   Screen.MousePointer = 0
'
'   crp_Imprim.Connect = "DSN=" & moddat_g_str_NomEsq & "; UID=" & moddat_g_str_EntDat & "; PWD=" & moddat_g_str_ClaDat
'
'   If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 4 Then
'      crp_Imprim.DataFiles(0) = "CRE_SOLMAE"
'      crp_Imprim.DataFiles(1) = "CLI_DATGEN"
'      crp_Imprim.DataFiles(2) = "TRA_SEGEXC"
'      crp_Imprim.DataFiles(3) = "CRE_PRODUC"
'   Else
'      crp_Imprim.DataFiles(0) = "RPT_EVAHIP"
'      crp_Imprim.DataFiles(1) = ""
'      crp_Imprim.DataFiles(2) = ""
'      crp_Imprim.DataFiles(3) = ""
'   End If
'
'   If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 1 Then
'      crp_Imprim.SelectionFormula = "{RPT_EVAHIP.EVAHIP_NOMTER} = '" & modgen_g_str_NombPC & "' AND {RPT_EVAHIP.EVAHIP_NOMRPT} = 'CRE_EVAHIP_01.RPT'"
'      crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "CRE_EVAHIP_14.RPT"
'
'   ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 2 Then
'      crp_Imprim.SelectionFormula = "{RPT_EVAHIP.EVAHIP_NOMTER} = '" & modgen_g_str_NombPC & "' AND {RPT_EVAHIP.EVAHIP_NOMRPT} = 'CRE_EVAHIP_04.RPT'"
'      crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "CRE_EVAHIP_15.RPT"
'
'   ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 3 Then
'      crp_Imprim.SelectionFormula = "{RPT_EVAHIP.EVAHIP_NOMTER} = '" & modgen_g_str_NombPC & "' AND {RPT_EVAHIP.EVAHIP_NOMRPT} = 'CRE_EVAHIP_06.RPT'"
'      crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "CRE_EVAHIP_16.RPT"
'
'   ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 4 Then
'      If cmb_TipIns.ItemData(cmb_TipIns.ListIndex) = 1 Then
'         crp_Imprim.SelectionFormula = "{TRA_SEGEXC.SEGEXC_CODINS} = 41 AND "
'      ElseIf cmb_TipIns.ItemData(cmb_TipIns.ListIndex) = 2 Then
'         crp_Imprim.SelectionFormula = "{TRA_SEGEXC.SEGEXC_CODINS} = 42 AND "
'      ElseIf cmb_TipIns.ItemData(cmb_TipIns.ListIndex) = 3 Then
'         crp_Imprim.SelectionFormula = "{TRA_SEGEXC.SEGEXC_CODINS} = 61 AND "
'      ElseIf cmb_TipIns.ItemData(cmb_TipIns.ListIndex) = 4 Then
'         crp_Imprim.SelectionFormula = "{TRA_SEGEXC.SEGEXC_CODINS} = 62 AND "
'      End If
'      crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{TRA_SEGEXC.SEGFECCRE} >= " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " AND "
'      crp_Imprim.SelectionFormula = crp_Imprim.SelectionFormula & "{TRA_SEGEXC.SEGFECCRE} <= " & Format(CDate(ipp_FecFin.Text), "yyyymmdd")
'      crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "CRE_EVAHIP_17.RPT"
'
'   ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 5 Then
'      crp_Imprim.SelectionFormula = "{RPT_EVAHIP.EVAHIP_NOMTER} = '" & modgen_g_str_NombPC & "' AND {RPT_EVAHIP.EVAHIP_NOMRPT} = 'CRE_EVAHIP_18.RPT'"
'      crp_Imprim.ReportFileName = g_str_RutRpt & "\" & "CRE_EVAHIP_18.RPT"
'   End If
'
'   crp_Imprim.Action = 1
'End Sub
