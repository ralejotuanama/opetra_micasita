VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_Rpt_CuoPen_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5550
   Icon            =   "OpeTra_frm_342.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   2265
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5565
      _Version        =   65536
      _ExtentX        =   9816
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
         Left            =   45
         TabIndex        =   4
         Top             =   60
         Width           =   5445
         _Version        =   65536
         _ExtentX        =   9604
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
            TabIndex        =   5
            Top             =   45
            Width           =   4155
            _Version        =   65536
            _ExtentX        =   7329
            _ExtentY        =   1005
            _StockProps     =   15
            Caption         =   "Reporte de Cuotas Pendientes de Pago"
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
            Picture         =   "OpeTra_frm_342.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   45
         TabIndex        =   6
         Top             =   780
         Width           =   5445
         _Version        =   65536
         _ExtentX        =   9604
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
            Left            =   4830
            Picture         =   "OpeTra_frm_342.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Salir de la Opción"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_342.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel5 
         Height          =   735
         Left            =   45
         TabIndex        =   7
         Top             =   1470
         Width           =   5445
         _Version        =   65536
         _ExtentX        =   9604
         _ExtentY        =   1296
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
         Begin EditLib.fpDateTime ipp_FecIni 
            Height          =   315
            Left            =   1200
            TabIndex        =   0
            Top             =   210
            Width           =   1425
            _Version        =   196608
            _ExtentX        =   2514
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
            Caption         =   "A la Fecha :"
            Height          =   315
            Left            =   240
            TabIndex        =   8
            Top             =   240
            Width           =   975
         End
      End
   End
End
Attribute VB_Name = "frm_Rpt_CuoPen_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call gs_CentraForm(Me)
   
   Call gs_SetFocus(ipp_FecIni)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   ipp_FecIni.Text = Format(date, "dd/mm/yyyy")
End Sub

Private Sub cmd_ExpExc_Click()
   'Confirmacion
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_GenExc
   Screen.MousePointer = 0
End Sub

Private Sub fs_GenExc()
Dim r_obj_Excel      As Excel.Application
Dim r_int_ConVer     As Integer
Dim r_int_Cont       As Integer
Dim r_int_Cont1      As Integer
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT A.HIPCUO_NUMOPE AS OPERACION, "
   g_str_Parame = g_str_Parame & "       TRIM(C.DATGEN_APEPAT)||' '||TRIM(C.DATGEN_APEMAT)||' '||TRIM(C.DATGEN_NOMBRE) AS NOMBRE_CLIENTE,"
   g_str_Parame = g_str_Parame & "       A.HIPCUO_NUMCUO   AS NRO_CUOTA, "
   g_str_Parame = g_str_Parame & "       A.HIPCUO_FECVCT   AS FECHA_VCTO,"
   g_str_Parame = g_str_Parame & "       A.HIPCUO_CAPITA+A.HIPCUO_INTERE+A.HIPCUO_DESORG+A.HIPCUO_VIVORG+A.HIPCUO_OTRORG AS CUOTA_SIN_CARGOS,"
   g_str_Parame = g_str_Parame & "       A.HIPCUO_CAPITA+A.HIPCUO_INTERE+A.HIPCUO_DESORG+A.HIPCUO_VIVORG+A.HIPCUO_OTRORG+A.HIPCUO_CAPBBP+A.HIPCUO_INTBBP AS CUOTA_INC_PBP, "
   g_str_Parame = g_str_Parame & "       A.HIPCUO_CAPITA+A.HIPCUO_INTERE+A.HIPCUO_DESORG+A.HIPCUO_VIVORG+A.HIPCUO_OTRORG+A.HIPCUO_INTCOM+"
   g_str_Parame = g_str_Parame & "       A.HIPCUO_INTMOR+A.HIPCUO_GASCOB+A.HIPCUO_OTRGAS+HIPCUO_CAPBBP+HIPCUO_INTBBP AS CUOTA_AL_DIA"
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPCUO A "
   g_str_Parame = g_str_Parame & " INNER JOIN CRE_HIPMAE B ON B.HIPMAE_NUMOPE = A.HIPCUO_NUMOPE AND B.HIPMAE_SITUAC = 2"
   g_str_Parame = g_str_Parame & " INNER JOIN CLI_DATGEN C ON C.DATGEN_TIPDOC = B.HIPMAE_TDOCLI AND C.DATGEN_NUMDOC = B.HIPMAE_NDOCLI"
   g_str_Parame = g_str_Parame & " WHERE A.HIPCUO_TIPCRO = 1 AND A.HIPCUO_SITUAC = 2"
   g_str_Parame = g_str_Parame & "   AND A.HIPCUO_FECVCT < " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " "
   g_str_Parame = g_str_Parame & "ORDER BY A.HIPCUO_NUMOPE, A.HIPCUO_NUMCUO"
  
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se encontraron Operaciones registradas.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Cells(1, 1) = "REPORTE DE CUOTAS PENDIENTES DE PAGO AL " & UCase(Format(ipp_FecIni.Text, "Long Date"))
      .Range("A1:H1").Select
      .Range("A1:H1").HorizontalAlignment = xlHAlignCenter
      .Range("A1:H1").Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 11)).Font.Size = 14
      r_obj_Excel.Selection.MergeCells = True
      
      .Cells(4, 1) = "ITEM"
      .Cells(4, 2) = "OPERACION"
      .Cells(4, 3) = "NOMBRE DE CLIENTE"
      .Cells(4, 4) = "Nº CUOTA"
      .Cells(4, 5) = "FECHA VCTO."
      .Cells(4, 6) = "CUOTA SIN CARGO"
      .Cells(4, 7) = "CUOTA INC.PBP"
      .Cells(4, 8) = "CUOTA AL DIA"
            
      .Range(.Cells(4, 1), .Cells(4, 25)).Font.Bold = True
      .Range(.Cells(4, 1), .Cells(4, 27)).HorizontalAlignment = xlHAlignCenter
      
      .Columns("A").ColumnWidth = 6
      .Columns("B").ColumnWidth = 13
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 40
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 10
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").ColumnWidth = 12
      .Columns("F").ColumnWidth = 17
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("G").ColumnWidth = 16
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      .Columns("H").ColumnWidth = 16
      .Columns("H").HorizontalAlignment = xlHAlignCenter
   End With
   
   g_rst_Princi.MoveFirst
   r_int_ConVer = 5
   r_int_Cont = 1
   
   Do While Not g_rst_Princi.EOF
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_Cont
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = "'" & Trim(g_rst_Princi!OPERACION)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = Trim(g_rst_Princi!NOMBRE_CLIENTE)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = Trim(g_rst_Princi!NRO_CUOTA)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = "'" & gf_FormatoFecha(CStr(g_rst_Princi!FECHA_VCTO))
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = Format(g_rst_Princi!CUOTA_SIN_CARGOS, "#,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = Format(g_rst_Princi!CUOTA_INC_PBP, "#,###,##0.00")
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = Format(g_rst_Princi!CUOTA_AL_DIA, "#,###,##0.00")
      
      r_obj_Excel.Range(r_obj_Excel.Cells(r_int_ConVer, 3), r_obj_Excel.Cells(r_int_ConVer, 3)).HorizontalAlignment = xlHAlignLeft
      r_obj_Excel.Range(r_obj_Excel.Cells(r_int_ConVer, 5), r_obj_Excel.Cells(r_int_ConVer, 5)).HorizontalAlignment = xlHAlignCenter
      r_obj_Excel.Range(r_obj_Excel.Cells(r_int_ConVer, 6), r_obj_Excel.Cells(r_int_ConVer, 6)).HorizontalAlignment = xlHAlignRight
      r_obj_Excel.Range(r_obj_Excel.Cells(r_int_ConVer, 7), r_obj_Excel.Cells(r_int_ConVer, 7)).HorizontalAlignment = xlHAlignRight
      r_obj_Excel.Range(r_obj_Excel.Cells(r_int_ConVer, 8), r_obj_Excel.Cells(r_int_ConVer, 8)).HorizontalAlignment = xlHAlignRight
      
      r_int_ConVer = r_int_ConVer + 1
      r_int_Cont = r_int_Cont + 1
      g_rst_Princi.MoveNext
      DoEvents
   Loop
   
   r_obj_Excel.Range(r_obj_Excel.Cells(4, 1), r_obj_Excel.Cells(4, 8)).Select
   r_obj_Excel.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
   r_obj_Excel.Range(r_obj_Excel.Cells(5, 1), r_obj_Excel.Cells(5, 8)).Select
   r_obj_Excel.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
   r_obj_Excel.ActiveSheet.Range("A4:H4").Interior.Color = RGB(146, 208, 80)
   r_obj_Excel.Range(r_obj_Excel.Cells(4, 1), r_obj_Excel.Cells(4, 8)).Select
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub ipp_FecIni_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_ExpExc)
   End If
End Sub
