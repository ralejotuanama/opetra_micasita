VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_Rpt_RepVar_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12405
   Icon            =   "OpeTra_frm_819.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   12405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   8565
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   12435
      _Version        =   65536
      _ExtentX        =   21934
      _ExtentY        =   15108
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
         Left            =   60
         TabIndex        =   8
         Top             =   60
         Width           =   12315
         _Version        =   65536
         _ExtentX        =   21722
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
            Height          =   285
            Left            =   660
            TabIndex        =   9
            Top             =   180
            Width           =   6855
            _Version        =   65536
            _ExtentX        =   12091
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Reporte Varios"
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
            Picture         =   "OpeTra_frm_819.frx":000C
            Top             =   90
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   60
         TabIndex        =   10
         Top             =   780
         Width           =   12315
         _Version        =   65536
         _ExtentX        =   21722
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
            Picture         =   "OpeTra_frm_819.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   11700
            Picture         =   "OpeTra_frm_819.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_819.frx":0A62
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Buscar Operaciones"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   1245
         Left            =   60
         TabIndex        =   11
         Top             =   1485
         Width           =   12315
         _Version        =   65536
         _ExtentX        =   21722
         _ExtentY        =   2196
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
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   120
            Width           =   6135
         End
         Begin EditLib.fpDateTime ipp_FecIni 
            Height          =   315
            Left            =   1440
            TabIndex        =   1
            Top             =   480
            Width           =   1545
            _Version        =   196608
            _ExtentX        =   2725
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
            Text            =   "01/01/2008"
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
            Left            =   1440
            TabIndex        =   2
            Top             =   810
            Width           =   1545
            _Version        =   196608
            _ExtentX        =   2725
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
            Text            =   "01/01/2008"
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
         Begin VB.Label Label20 
            Caption         =   "Fecha de Fin:"
            Height          =   195
            Left            =   60
            TabIndex        =   15
            Top             =   870
            Width           =   1125
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Reporte:"
            Height          =   195
            Left            =   60
            TabIndex        =   14
            Top             =   180
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "Fecha de Inicio:"
            Height          =   225
            Left            =   60
            TabIndex        =   12
            Top             =   525
            Width           =   1215
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   5715
         Left            =   60
         TabIndex        =   13
         Top             =   2775
         Width           =   12315
         _Version        =   65536
         _ExtentX        =   21722
         _ExtentY        =   10081
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
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   5685
            Left            =   30
            TabIndex        =   3
            Top             =   30
            Width           =   12270
            _ExtentX        =   21643
            _ExtentY        =   10028
            _Version        =   393216
            Rows            =   10
            Cols            =   11
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            SelectionMode   =   1
         End
      End
   End
End
Attribute VB_Name = "frm_Rpt_RepVar_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmb_TipRep_Click()
   Call fs_Limpia
   If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 2 Then
      ipp_FecIni.Enabled = True
      ipp_FecFin.Enabled = False
   ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 1 Or cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 4 Then
      ipp_FecIni.Enabled = False
      ipp_FecFin.Enabled = False
   Else
      ipp_FecIni.Enabled = True
      ipp_FecFin.Enabled = True
   End If
End Sub

Private Sub cmb_TipRep_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecIni)
   End If
End Sub

Private Sub cmd_Buscar_Click()
   If CDate(ipp_FecFin.Text) < CDate(ipp_FecIni.Text) Then
      MsgBox "La fecha de fin no puede ser menor a la fecha de inicio.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecIni)
      Exit Sub
   End If
   If cmb_TipRep.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Reporte.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipRep)
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Select Case cmb_TipRep.ItemData(cmb_TipRep.ListIndex)
      Case 1: Call fs_BusCli_CreRef
      Case 2: Call fs_BusCli_CuoPen
      Case 3: Call fs_BusCli_CreTra
      Case 4: Call fs_BusCli_AlmExt
   End Select
   Screen.MousePointer = 0
End Sub

Private Sub cmd_ExpExc_Click()
   If grd_Listad.Rows = 0 Then
      MsgBox "No existe información a exportar.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Select Case cmb_TipRep.ItemData(cmb_TipRep.ListIndex)
      Case 1: Call fs_Exportar_CreRef
      Case 2: Call fs_Exportar_CuoPen
      Case 3: Call fs_Exportar_CreTra
      Case 4: Call fs_Exportar_AlmExt
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
   Call fs_Limpia
   Call gs_CentraForm(Me)
   
   Call gs_SetFocus(cmb_TipRep)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   cmb_TipRep.Clear
   cmb_TipRep.AddItem "REPORTE DE CREDITOS REFINANCIADOS"
   cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = CInt(1)
   cmb_TipRep.AddItem "REPORTE DE CUOTAS PENDIENTES DE PAGO"
   cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = CInt(2)
   cmb_TipRep.AddItem "REPORTE DE CREDITOS TRANSFERIDOS"
   cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = CInt(3)
   cmb_TipRep.AddItem "REPORTE DE ALMACEN EXTERNO"
   cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = CInt(4)
   cmb_TipRep.ListIndex = -1
End Sub

Private Sub fs_Config()
   If cmb_TipRep.ListIndex > -1 Then
      If cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 1 Then
         grd_Listad.ColWidth(0) = 400
         grd_Listad.ColWidth(1) = 1320
         grd_Listad.ColWidth(2) = 1140
         grd_Listad.ColWidth(3) = 3650
         grd_Listad.ColWidth(4) = 1220
         grd_Listad.ColWidth(5) = 2020
         grd_Listad.ColWidth(6) = 1200
         grd_Listad.ColWidth(7) = 1200
         grd_Listad.ColWidth(8) = 0
         grd_Listad.ColWidth(9) = 0
         grd_Listad.ColWidth(10) = 0
         grd_Listad.ColAlignment(7) = flexAlignCenterCenter
         
      ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 2 Then
         grd_Listad.ColWidth(0) = 400
         grd_Listad.ColWidth(1) = 1220
         grd_Listad.ColWidth(2) = 1120
         grd_Listad.ColWidth(3) = 3350
         grd_Listad.ColWidth(4) = 800
         grd_Listad.ColWidth(5) = 1230
         grd_Listad.ColWidth(6) = 1300
         grd_Listad.ColWidth(7) = 1230
         grd_Listad.ColWidth(8) = 1230
         grd_Listad.ColWidth(9) = 0
         grd_Listad.ColWidth(10) = 0
         grd_Listad.ColAlignment(7) = flexAlignRightCenter
         
      ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 3 Then
         grd_Listad.ColWidth(0) = 400
         grd_Listad.ColWidth(1) = 1320
         grd_Listad.ColWidth(2) = 1140
         grd_Listad.ColWidth(3) = 3750
         grd_Listad.ColWidth(4) = 3750
         grd_Listad.ColWidth(5) = 1500
         grd_Listad.ColWidth(6) = 0
         grd_Listad.ColWidth(7) = 0
         grd_Listad.ColWidth(8) = 0
         grd_Listad.ColWidth(9) = 0
         grd_Listad.ColWidth(10) = 0
         
      ElseIf cmb_TipRep.ItemData(cmb_TipRep.ListIndex) = 4 Then
         grd_Listad.ColWidth(0) = 600
         grd_Listad.ColWidth(1) = 1320
         grd_Listad.ColWidth(2) = 1140
         grd_Listad.ColWidth(3) = 3950
         grd_Listad.ColWidth(4) = 1300
         grd_Listad.ColWidth(5) = 2200
         grd_Listad.ColWidth(6) = 1300
         grd_Listad.ColWidth(7) = 1200
         grd_Listad.ColWidth(8) = 1200
         grd_Listad.ColWidth(9) = 900
         grd_Listad.ColWidth(10) = 1000
         
         grd_Listad.ColAlignment(7) = flexAlignCenterCenter
         grd_Listad.ColAlignment(8) = flexAlignCenterCenter
         grd_Listad.ColAlignment(9) = flexAlignCenterCenter
         grd_Listad.ColAlignment(10) = flexAlignCenterCenter
      End If
   End If
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignLeftCenter
   grd_Listad.ColAlignment(3) = flexAlignLeftCenter
   grd_Listad.ColAlignment(4) = flexAlignCenterCenter
   grd_Listad.ColAlignment(5) = flexAlignCenterCenter
   grd_Listad.ColAlignment(6) = flexAlignRightCenter
End Sub

Private Sub fs_Limpia()
   Call gs_LimpiaGrid(grd_Listad)
   ipp_FecIni.Text = Format(date - CDate(30), "dd/mm/yyyy")
   ipp_FecFin.Text = Format(date, "dd/mm/yyyy")
End Sub

Private Sub fs_BusCli_CreRef()
Dim r_dbl_TotPag  As Double

   Call fs_Config
   grd_Listad.Redraw = False
   
   'Buscando Información del Crédito
   Call gs_LimpiaGrid(grd_Listad)
   
   'CABECERA
   grd_Listad.Rows = grd_Listad.Rows + 2
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.FixedRows = 1
   
   grd_Listad.Row = 0
   grd_Listad.Col = 0:        grd_Listad.Text = "Item":                    grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 1:        grd_Listad.Text = "Nro Operación":           grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 2:        grd_Listad.Text = "DNI":                     grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 3:        grd_Listad.Text = "Apellidos y Nombres":     grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 4:        grd_Listad.Text = "Fecha Desemb.":           grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 5:        grd_Listad.Text = "Moneda Desembolso":       grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 6:        grd_Listad.Text = "Monto Desemb.":           grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 7:        grd_Listad.Text = "Situación":               grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Rows = grd_Listad.Rows - 1
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT ROWNUM, A.HIPMAE_TDOCLI || ' - ' || TRIM(A.HIPMAE_NDOCLI) AS DNI, "
   g_str_Parame = g_str_Parame & "       TRIM(B.DATGEN_APEPAT) || ' ' || TRIM(B.DATGEN_APEMAT) || ' '|| TRIM(B.DATGEN_NOMBRE) AS NOMBRE, "
   g_str_Parame = g_str_Parame & "       SUBSTR(A.HIPMAE_NUMOPE,1,3) ||'-' || SUBSTR(A.HIPMAE_NUMOPE,4,2) ||'-' ||SUBSTR(A.HIPMAE_NUMOPE,6,5) AS NUMOPE, "
   g_str_Parame = g_str_Parame & "       TO_DATE(A.HIPMAE_FECDES,'YYYY/MM/DD') AS FECHA, TRIM(C.PARDES_DESCRI) AS MONEDA, "
   g_str_Parame = g_str_Parame & "       ROUND(A.HIPMAE_IMPDES, 2) AS MONTO, TRIM(D.PARDES_DESCRI) AS ESTADO "
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPMAE A "
   g_str_Parame = g_str_Parame & " INNER JOIN CLI_DATGEN B ON (B.DATGEN_NUMDOC = A.HIPMAE_NDOCLI AND B.DATGEN_TIPDOC = A.HIPMAE_TDOCLI) "
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES C ON (C.PARDES_CODGRP = '204' AND C.PARDES_CODITE = A.HIPMAE_MONEDA) "
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES D ON (D.PARDES_CODGRP = '027' AND D.PARDES_CODITE = A.HIPMAE_SITUAC) "
   g_str_Parame = g_str_Parame & " WHERE HIPMAE_REFINA = 1 "
   g_str_Parame = g_str_Parame & " ORDER BY NUMOPE "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "El cliente no cuenta con una clasificación.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      r_dbl_TotPag = 0
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         
         grd_Listad.Col = 0
         grd_Listad.Text = CStr(g_rst_Princi!ROWNUM)
                         
         grd_Listad.Col = 1
         grd_Listad.Text = CStr(g_rst_Princi!NUMOPE)
           
         grd_Listad.Col = 2
         grd_Listad.Text = CStr(g_rst_Princi!DNI)
         
         grd_Listad.Col = 3
         grd_Listad.Text = CStr(g_rst_Princi!NOMBRE)
         
         grd_Listad.Col = 4
         grd_Listad.Text = CStr(g_rst_Princi!fecha)
         
         grd_Listad.Col = 5
         grd_Listad.Text = CStr(g_rst_Princi!MONEDA)
         
         grd_Listad.Col = 6
         grd_Listad.Text = Format(g_rst_Princi!MONTO, "###,###,##0.00")
         
         grd_Listad.Col = 7
         grd_Listad.Text = CStr(g_rst_Princi!estado)
         g_rst_Princi.MoveNext
      Loop
      
      Call gs_UbiIniGrid(grd_Listad)
      grd_Listad.Redraw = True
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_BusCli_CreTra()
Dim r_int_Contad     As Integer

   Call fs_Config
   grd_Listad.Redraw = False
   
   'Buscando Información del Crédito
   Call gs_LimpiaGrid(grd_Listad)
   
   'CABECERA
   grd_Listad.Rows = grd_Listad.Rows + 2
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.FixedRows = 1
   r_int_Contad = 0
   
   grd_Listad.Row = 0
   grd_Listad.Col = 0:   grd_Listad.Text = "Item":                  grd_Listad.ColAlignment(0) = flexAlignCenterCenter:   grd_Listad.ColWidth(0) = 600
   grd_Listad.Col = 1:   grd_Listad.Text = "Nª Operación":          grd_Listad.ColAlignment(1) = flexAlignCenterCenter:   grd_Listad.ColWidth(1) = 1200
   grd_Listad.Col = 2:   grd_Listad.Text = "DNI":                   grd_Listad.ColAlignment(2) = flexAlignCenterCenter:   grd_Listad.ColWidth(2) = 1200
   grd_Listad.Col = 3:   grd_Listad.Text = "Apellidos y Nombres":   grd_Listad.ColAlignment(3) = flexAlignLeftCenter:     grd_Listad.ColWidth(3) = 3600
   grd_Listad.Col = 4:   grd_Listad.Text = "Producto":              grd_Listad.ColAlignment(4) = flexAlignLeftCenter:     grd_Listad.ColWidth(4) = 3500
   grd_Listad.Col = 5:   grd_Listad.Text = "Banco":                 grd_Listad.ColAlignment(5) = flexAlignCenterCenter:   grd_Listad.ColWidth(5) = 2500
   grd_Listad.Col = 6:   grd_Listad.Text = "Fecha Transf.":         grd_Listad.ColAlignment(6) = flexAlignCenterCenter:   grd_Listad.ColWidth(6) = 1200
   grd_Listad.Col = 7:   grd_Listad.Text = "Monto Activo TNC":      grd_Listad.ColAlignment(7) = flexAlignRightCenter:    grd_Listad.ColWidth(7) = 1500
   grd_Listad.Col = 8:   grd_Listad.Text = "Monto Activo TC":       grd_Listad.ColAlignment(8) = flexAlignRightCenter:    grd_Listad.ColWidth(8) = 1400
   grd_Listad.Col = 9:   grd_Listad.Text = "Monto Pasivo TNC":      grd_Listad.ColAlignment(9) = flexAlignRightCenter:    grd_Listad.ColWidth(9) = 1500
   grd_Listad.Col = 10:  grd_Listad.Text = "Monto Pasivo TC":       grd_Listad.ColAlignment(10) = flexAlignRightCenter:   grd_Listad.ColWidth(10) = 1400
   grd_Listad.Rows = grd_Listad.Rows - 1
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT TRIM(A.SALMIG_NUMOPE) AS OPERACION, "
   g_str_Parame = g_str_Parame & "       TRIM(B.HIPMAE_TDOCLI)||'-'||TRIM(B.HIPMAE_NDOCLI)AS TIPO_DOCUM, "
   g_str_Parame = g_str_Parame & "       TRIM(DATGEN_APEPAT)||' '||TRIM(DATGEN_APEMAT)||' '||TRIM(DATGEN_NOMBRE) AS NOMBRE_CLIENTE, "
   g_str_Parame = g_str_Parame & "       TRIM(D.PRODUC_DESCRI) AS PRODUCTO, "
   g_str_Parame = g_str_Parame & "       TRIM(E.PARDES_DESCRI) AS BANCO, "
   g_str_Parame = g_str_Parame & "       SUBSTR(A.SALMIG_FECMIG,7,2)||'/'||SUBSTR(A.SALMIG_FECMIG,5,2)||'/'||SUBSTR(A.SALMIG_FECMIG,1,4) AS FECHA_TRANSF, "
   g_str_Parame = g_str_Parame & "       A.SALMIG_ACTTNC       AS MONTO_ACT_TNC, "
   g_str_Parame = g_str_Parame & "       A.SALMIG_ACTTC        AS MONTO_ACT_TC, "
   g_str_Parame = g_str_Parame & "       A.SALMIG_PASTNC       AS MONTO_PAS_TNC, "
   g_str_Parame = g_str_Parame & "       A.SALMIG_PASTC        AS MONTO_PAS_TC "
   g_str_Parame = g_str_Parame & "  FROM CRE_SALMIG A "
   g_str_Parame = g_str_Parame & " INNER JOIN CRE_HIPMAE B ON B.HIPMAE_NUMOPE = A.SALMIG_NUMOPE "
   g_str_Parame = g_str_Parame & " INNER JOIN CLI_DATGEN C ON C.DATGEN_TIPDOC = B.HIPMAE_TDOCLI AND C.DATGEN_NUMDOC = B.HIPMAE_NDOCLI "
   g_str_Parame = g_str_Parame & " INNER JOIN CRE_PRODUC D ON D.PRODUC_CODIGO = B.HIPMAE_CODPRD "
   g_str_Parame = g_str_Parame & " INNER JOIN MNT_PARDES E ON E.PARDES_CODGRP = '122' AND E.PARDES_CODITE = A.SALMIG_CODBAN "
   g_str_Parame = g_str_Parame & " WHERE A.SALMIG_FECMIG >= '" & Format(ipp_FecIni.Text, "YYYYMMDD") & "' "
   g_str_Parame = g_str_Parame & "   AND A.SALMIG_FECMIG <= '" & Format(ipp_FecFin.Text, "YYYYMMDD") & "' "
   g_str_Parame = g_str_Parame & " ORDER BY A.SALMIG_FECMIG, A.SALMIG_NUMOPE "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se encontraron operaciones.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecIni)
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         r_int_Contad = r_int_Contad + 1
         
         grd_Listad.Col = 0
         grd_Listad.Text = Format(r_int_Contad, "00000")
         
         grd_Listad.Col = 1
         grd_Listad.Text = CStr(g_rst_Princi!OPERACION)
           
         grd_Listad.Col = 2
         grd_Listad.Text = CStr(g_rst_Princi!TIPO_DOCUM)
         
         grd_Listad.Col = 3
         grd_Listad.Text = CStr(g_rst_Princi!NOMBRE_CLIENTE)
         
         grd_Listad.Col = 4
         grd_Listad.Text = CStr(g_rst_Princi!PRODUCTO)
         
         grd_Listad.Col = 5
         grd_Listad.Text = CStr(g_rst_Princi!BANCO)
         
         grd_Listad.Col = 6
         grd_Listad.Text = CStr(g_rst_Princi!FECHA_TRANSF)
         
         grd_Listad.Col = 7
         grd_Listad.Text = Format(g_rst_Princi!MONTO_ACT_TNC, "###,##0.00")
         
         grd_Listad.Col = 8
         grd_Listad.Text = Format(g_rst_Princi!MONTO_ACT_TC, "###,##0.00")
         
         grd_Listad.Col = 9
         grd_Listad.Text = Format(g_rst_Princi!MONTO_PAS_TNC, "###,##0.00")
         
         grd_Listad.Col = 10
         grd_Listad.Text = Format(g_rst_Princi!MONTO_PAS_TC, "###,##0.00")
         
         g_rst_Princi.MoveNext
      Loop
      
      Call gs_UbiIniGrid(grd_Listad)
      grd_Listad.Redraw = True
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_BusCli_CuoPen()
Dim r_int_Cont As Integer

   Call fs_Config
   grd_Listad.Redraw = False
   
   'Buscando Información del Crédito
   Call gs_LimpiaGrid(grd_Listad)
   
   'CABECERA
   grd_Listad.Rows = grd_Listad.Rows + 2
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.FixedRows = 1
   
   grd_Listad.Row = 0
   grd_Listad.Col = 0:        grd_Listad.Text = "Item":                    grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 1:        grd_Listad.Text = "Nro Operación":           grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 2:        grd_Listad.Text = "DNI":                     grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 3:        grd_Listad.Text = "Apellidos y Nombres":     grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 4:        grd_Listad.Text = "N_Cuota":                 grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 5:        grd_Listad.Text = "Fecha Vencim.":           grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 6:        grd_Listad.Text = "Cuota Sin Cargos":        grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 7:        grd_Listad.Text = "Cuota Inc. PBP":          grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 8:        grd_Listad.Text = "Cuota al Día":            grd_Listad.CellAlignment = flexAlignCenterCenter
   
   grd_Listad.Rows = grd_Listad.Rows - 1
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT SUBSTR(HIPCUO_NUMOPE,1,3) || '-' || SUBSTR(HIPCUO_NUMOPE,4,2) || '-' ||SUBSTR(HIPCUO_NUMOPE,6,5) AS OPERACION, "
   g_str_Parame = g_str_Parame & "       (TRIM(B.HIPMAE_TDOCLI) || ' - ' || TRIM(B.HIPMAE_NDOCLI)) AS DOC_IDENTIDAD, "
   g_str_Parame = g_str_Parame & "       TRIM(C.DATGEN_APEPAT)||' '||TRIM(C.DATGEN_APEMAT)||' '||TRIM(C.DATGEN_NOMBRE) AS NOMBRE_CLIENTE,"
   g_str_Parame = g_str_Parame & "       A.HIPCUO_NUMCUO   AS NRO_CUOTA, "
   g_str_Parame = g_str_Parame & "       TO_DATE(A.HIPCUO_FECVCT ,'YYYY/MM/DD') AS FECHA_VCTO, "
   g_str_Parame = g_str_Parame & "       A.HIPCUO_CAPITA+A.HIPCUO_INTERE+A.HIPCUO_DESORG+A.HIPCUO_VIVORG+A.HIPCUO_OTRORG AS CUOTA_SIN_CARGOS, "
   g_str_Parame = g_str_Parame & "       A.HIPCUO_CAPITA+A.HIPCUO_INTERE+A.HIPCUO_DESORG+A.HIPCUO_VIVORG+A.HIPCUO_OTRORG+A.HIPCUO_CAPBBP+A.HIPCUO_INTBBP AS CUOTA_INC_PBP, "
   g_str_Parame = g_str_Parame & "       A.HIPCUO_CAPITA+A.HIPCUO_INTERE+A.HIPCUO_DESORG+A.HIPCUO_VIVORG+A.HIPCUO_OTRORG+A.HIPCUO_INTCOM+ "
   g_str_Parame = g_str_Parame & "       A.HIPCUO_INTMOR+A.HIPCUO_GASCOB+A.HIPCUO_OTRGAS+HIPCUO_CAPBBP+HIPCUO_INTBBP AS CUOTA_AL_DIA"
   g_str_Parame = g_str_Parame & "  FROM CRE_HIPCUO A "
   g_str_Parame = g_str_Parame & " INNER JOIN CRE_HIPMAE B ON B.HIPMAE_NUMOPE = A.HIPCUO_NUMOPE AND B.HIPMAE_SITUAC = 2 "
   g_str_Parame = g_str_Parame & " INNER JOIN CLI_DATGEN C ON C.DATGEN_TIPDOC = B.HIPMAE_TDOCLI AND C.DATGEN_NUMDOC = B.HIPMAE_NDOCLI "
   g_str_Parame = g_str_Parame & " WHERE A.HIPCUO_TIPCRO = 1 AND A.HIPCUO_SITUAC = 2"
   g_str_Parame = g_str_Parame & "   AND A.HIPCUO_FECVCT < " & Format(CDate(ipp_FecIni.Text), "yyyymmdd") & " "
   g_str_Parame = g_str_Parame & " ORDER BY A.HIPCUO_NUMOPE, A.HIPCUO_NUMCUO"
  
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se encontraron Operaciones registradas.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   r_int_Cont = 1
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         
         grd_Listad.Col = 0
         grd_Listad.Text = r_int_Cont
         
         grd_Listad.Col = 1
         grd_Listad.Text = CStr(g_rst_Princi!OPERACION)
           
         grd_Listad.Col = 2
         grd_Listad.Text = CStr(g_rst_Princi!DOC_IDENTIDAD)
         
         grd_Listad.Col = 3
         grd_Listad.Text = CStr(g_rst_Princi!NOMBRE_CLIENTE)
         
         grd_Listad.Col = 4
         grd_Listad.Text = CStr(g_rst_Princi!NRO_CUOTA)
         
         grd_Listad.Col = 5
         grd_Listad.Text = CStr(g_rst_Princi!FECHA_VCTO)
         
         grd_Listad.Col = 6
         grd_Listad.Text = Format(g_rst_Princi!CUOTA_SIN_CARGOS, "#,###,##0.00")
         
         grd_Listad.Col = 7
         grd_Listad.Text = Format(g_rst_Princi!CUOTA_INC_PBP, "#,###,##0.00")
         
         grd_Listad.Col = 8
         grd_Listad.Text = Format(g_rst_Princi!CUOTA_AL_DIA, "#,###,##0.00")
         
         g_rst_Princi.MoveNext
         r_int_Cont = r_int_Cont + 1
      Loop
      
      Call gs_UbiIniGrid(grd_Listad)
      grd_Listad.Redraw = True
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_BusCli_AlmExt()
Dim r_int_Cont As Integer

   Call fs_Config
   grd_Listad.Redraw = False
   
   'Buscando Información del Crédito
   Call gs_LimpiaGrid(grd_Listad)
   
   'CABECERA
   grd_Listad.Rows = grd_Listad.Rows + 2
   grd_Listad.Row = grd_Listad.Rows - 1
   grd_Listad.FixedRows = 1
   
   grd_Listad.Row = 0
   grd_Listad.Col = 0:        grd_Listad.Text = "Item":                    grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 1:        grd_Listad.Text = "Nro Operación":           grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 2:        grd_Listad.Text = "DNI":                     grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 3:        grd_Listad.Text = "Apellidos y Nombres":     grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 4:        grd_Listad.Text = "Fecha Desemb.":           grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 5:        grd_Listad.Text = "Moneda":                  grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 6:        grd_Listad.Text = "Monto Desemb.":           grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 7:        grd_Listad.Text = "Estado":                  grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 8:        grd_Listad.Text = "Fecha Inscrip.":          grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 9:        grd_Listad.Text = "Vinculado":               grd_Listad.CellAlignment = flexAlignCenterCenter
   grd_Listad.Col = 10:       grd_Listad.Text = "Código Caja":             grd_Listad.CellAlignment = flexAlignCenterCenter

   grd_Listad.Rows = grd_Listad.Rows - 1
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "   SELECT SUBSTR(A.HIPMAE_NUMOPE,1,3) || '-' || SUBSTR(A.HIPMAE_NUMOPE,4,2) || '-' ||SUBSTR(A.HIPMAE_NUMOPE,6,5) AS OPERACION, "
   g_str_Parame = g_str_Parame & "          (TRIM(A.HIPMAE_TDOCLI) || ' - ' || TRIM(A.HIPMAE_NDOCLI)) AS DOC_IDENTIDAD, "
   g_str_Parame = g_str_Parame & "          TRIM(B.DATGEN_APEPAT)||' '||TRIM(B.DATGEN_APEMAT)||' '||TRIM(B.DATGEN_NOMBRE) AS NOMBRE_CLIENTE, "
   g_str_Parame = g_str_Parame & "          TO_DATE(A.HIPMAE_FECDES ,'YYYY/MM/DD') AS FECHA_DESEMBOLSO, TRIM(C.PARDES_DESCRI) AS MONEDA, ROUND(A.HIPMAE_IMPDES, 2) AS MONTO_DESEMBOLSO, "
   g_str_Parame = g_str_Parame & "          TRIM(D.PARDES_DESCRI) AS ESTADO, TO_DATE(HIPGAR_FECINS ,'YYYY/MM/DD') AS FECHA_INSCRIPCION, "
   g_str_Parame = g_str_Parame & "          CASE WHEN LENGTH(TRIM(H.PARDES_DESCRI)) > 0 THEN TRIM(H.PARDES_DESCRI) ELSE 'NO' END AS VINCULADO, "
   g_str_Parame = g_str_Parame & "          CASE WHEN LENGTH(TRIM(A.HIPMAE_CODCUS)) > 0 THEN TRIM(A.HIPMAE_CODCUS) ELSE '0' END AS CODIGO_CAJA "
   g_str_Parame = g_str_Parame & "     FROM CRE_HIPMAE A "
   g_str_Parame = g_str_Parame & "          INNER JOIN CLI_DATGEN B ON B.DATGEN_TIPDOC = A.HIPMAE_TDOCLI AND B.DATGEN_NUMDOC = A.HIPMAE_NDOCLI "
   g_str_Parame = g_str_Parame & "          INNER JOIN MNT_PARDES C ON (C.PARDES_CODGRP = '204' AND C.PARDES_CODITE = A.HIPMAE_MONEDA) "
   g_str_Parame = g_str_Parame & "          INNER JOIN MNT_PARDES D ON (D.PARDES_CODGRP = '027' AND D.PARDES_CODITE = A.HIPMAE_SITUAC) "
   g_str_Parame = g_str_Parame & "           LEFT JOIN CRE_HIPGAR E ON E.HIPGAR_NUMOPE = A.HIPMAE_NUMOPE AND E.HIPGAR_BIEGAR = 1 "
   g_str_Parame = g_str_Parame & "           LEFT JOIN CRE_SOLINM F ON F.SOLINM_NUMSOL = A.HIPMAE_NUMSOL"
   g_str_Parame = g_str_Parame & "           LEFT JOIN PRY_DATGEN G ON G.DATGEN_CODIGO = F.SOLINM_PRYCOD"
   g_str_Parame = g_str_Parame & "           LEFT JOIN MNT_PARDES H ON H.PARDES_CODGRP = 214 AND H.PARDES_CODITE = G.DATGEN_PRYMCS"
   g_str_Parame = g_str_Parame & "    ORDER BY A.HIPMAE_NUMOPE "
  
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   r_int_Cont = 1
   
   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      MsgBox "No se encontraron Operaciones registradas.", vbInformation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         
         grd_Listad.Col = 0
         grd_Listad.Text = r_int_Cont
         
         grd_Listad.Col = 1
         grd_Listad.Text = CStr(g_rst_Princi!OPERACION)
           
         grd_Listad.Col = 2
         grd_Listad.Text = CStr(g_rst_Princi!DOC_IDENTIDAD)
         
         grd_Listad.Col = 3
         grd_Listad.Text = CStr(g_rst_Princi!NOMBRE_CLIENTE)
    
         grd_Listad.Col = 4
         grd_Listad.Text = CStr(g_rst_Princi!FECHA_DESEMBOLSO)
         
         grd_Listad.Col = 5
         grd_Listad.Text = CStr(g_rst_Princi!MONEDA)
         
         grd_Listad.Col = 6
         grd_Listad.Text = Format(g_rst_Princi!MONTO_DESEMBOLSO, "#,###,##0.00")
         
         grd_Listad.Col = 7
         grd_Listad.Text = CStr(g_rst_Princi!estado)
         
         grd_Listad.Col = 8
         If IsNull(g_rst_Princi!FECHA_INSCRIPCION) Then
            grd_Listad.Text = ""
         Else
            grd_Listad.Text = CStr(g_rst_Princi!FECHA_INSCRIPCION)
         End If
         
         grd_Listad.Col = 9
         grd_Listad.Text = CStr(g_rst_Princi!VINCULADO)
         
         grd_Listad.Col = 10
         grd_Listad.Text = IIf(IsNull(g_rst_Princi!CODIGO_CAJA), 0, CStr(g_rst_Princi!CODIGO_CAJA))
         
         g_rst_Princi.MoveNext
         r_int_Cont = r_int_Cont + 1
      Loop
      
      Call gs_UbiIniGrid(grd_Listad)
      grd_Listad.Redraw = True
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Exportar_CreRef()
   Dim r_obj_Excel      As Excel.Application
   Dim r_int_NroFil     As Integer
   Dim r_int_nroaux     As Integer
   
   'Generando Excel
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.DisplayAlerts = False
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   r_int_NroFil = 4
   
   With r_obj_Excel.ActiveSheet
      .Cells(2, 1) = "REPORTE DE CREDITOS REFINANCIADOS DEL " & ipp_FecIni.Text & " AL " & ipp_FecFin.Text
      .Range("A2:H2").HorizontalAlignment = xlHAlignCenter
      .Range("A2:H2").Font.Bold = True
      .Range(.Cells(2, 1), .Cells(2, 8)).Font.Size = 14
      .Range("A2:H2").MergeCells = True
      
      .Cells(r_int_NroFil, 1) = "ITEM":                  .Columns("A").ColumnWidth = 6
      .Cells(r_int_NroFil, 2) = "NRO DE OPERACION":      .Columns("B").ColumnWidth = 17
      .Cells(r_int_NroFil, 3) = "DNI":                   .Columns("C").ColumnWidth = 12
      .Cells(r_int_NroFil, 4) = "APELLIDOS Y NOMBRES":   .Columns("D").ColumnWidth = 45
      .Cells(r_int_NroFil, 5) = "FECHA DESEMBOLSO":      .Columns("E").ColumnWidth = 17
      .Cells(r_int_NroFil, 6) = "MONEDA DESEMBOLSO":     .Columns("F").ColumnWidth = 20
      .Cells(r_int_NroFil, 7) = "MONTO DESEMBOLSO":      .Columns("G").ColumnWidth = 18
      .Cells(r_int_NroFil, 8) = "SITUACION":             .Columns("H").ColumnWidth = 14
      
      .Range("A4:H4").Interior.Color = RGB(146, 208, 80)
      
      '.Columns("A").HorizontalAlignment = xlHAlignCenter
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").HorizontalAlignment = xlHAlignLeft
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("H").HorizontalAlignment = xlHAlignCenter
      
      .Cells(r_int_NroFil, 1).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 2).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 3).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 4).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 5).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 6).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 7).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 8).HorizontalAlignment = xlHAlignCenter
      
      .Range(.Cells(r_int_NroFil, 1), .Cells(r_int_NroFil, 11)).Font.Bold = True
      r_int_NroFil = r_int_NroFil + 1
      
      For r_int_nroaux = 1 To grd_Listad.Rows - 1
         .Cells(r_int_NroFil, 1) = grd_Listad.TextMatrix(r_int_nroaux, 0)
         .Cells(r_int_NroFil, 2) = grd_Listad.TextMatrix(r_int_nroaux, 1)
         .Cells(r_int_NroFil, 3) = grd_Listad.TextMatrix(r_int_nroaux, 2)
         .Cells(r_int_NroFil, 4) = grd_Listad.TextMatrix(r_int_nroaux, 3)
         .Cells(r_int_NroFil, 5) = grd_Listad.TextMatrix(r_int_nroaux, 4)
         .Cells(r_int_NroFil, 6) = grd_Listad.TextMatrix(r_int_nroaux, 5)
         .Cells(r_int_NroFil, 7) = grd_Listad.TextMatrix(r_int_nroaux, 6)
         .Cells(r_int_NroFil, 8) = grd_Listad.TextMatrix(r_int_nroaux, 7)
         r_int_NroFil = r_int_NroFil + 1
      Next
      
      '.Columns("F").Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.00"
      r_int_NroFil = r_int_NroFil + 2
       
      '.Range(.Cells(4, 8), .Cells(r_int_NroFil, 8)).Font.Name = "Arial"
      .Range(.Cells(4, 1), .Cells(r_int_NroFil, 8)).Font.Size = 9
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub fs_Exportar_CuoPen()
Dim r_obj_Excel      As Excel.Application
Dim r_int_ConVer     As Integer
Dim r_int_Contad     As Integer
Dim r_int_nroaux     As Integer
   
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   
   With r_obj_Excel.ActiveSheet
      .Cells(1, 1) = "REPORTE DE CUOTAS PENDIENTES DE PAGO AL " & UCase(Format(ipp_FecIni.Text, "Long Date"))
      .Range("A1:I1").Select
      .Range("A1:I1").HorizontalAlignment = xlHAlignCenter
      .Range("A1:I1").Font.Bold = True
      .Range(.Cells(1, 1), .Cells(1, 11)).Font.Size = 14
      r_obj_Excel.Selection.MergeCells = True
      
      .Cells(4, 1) = "ITEM"
      .Cells(4, 2) = "NRO DE OPERACION"
      .Cells(4, 3) = "DNI"
      .Cells(4, 4) = "APELLIDOS Y NOMBRES"
      .Cells(4, 5) = "Nº CUOTA"
      .Cells(4, 6) = "FECHA VCTO."
      .Cells(4, 7) = "CUOTA SIN CARGO"
      .Cells(4, 8) = "CUOTA INC.PBP"
      .Cells(4, 9) = "CUOTA AL DIA"
            
      .Range(.Cells(4, 1), .Cells(4, 25)).Font.Bold = True
      .Range(.Cells(4, 1), .Cells(4, 27)).HorizontalAlignment = xlHAlignCenter
      
      .Columns("A").ColumnWidth = 6
      .Columns("B").ColumnWidth = 17
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").ColumnWidth = 13
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("D").ColumnWidth = 40
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").ColumnWidth = 10
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").ColumnWidth = 12
      .Columns("G").ColumnWidth = 17
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      .Columns("H").ColumnWidth = 16
      .Columns("H").HorizontalAlignment = xlHAlignCenter
      .Columns("I").ColumnWidth = 16
      .Columns("I").HorizontalAlignment = xlHAlignCenter
   End With
   
   r_int_ConVer = 5
   r_int_Contad = 1
   
   For r_int_nroaux = 1 To grd_Listad.Rows - 1
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 1) = r_int_Contad
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 2) = "'" & grd_Listad.TextMatrix(r_int_nroaux, 1)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 3) = grd_Listad.TextMatrix(r_int_nroaux, 2)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 4) = grd_Listad.TextMatrix(r_int_nroaux, 3)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 5) = grd_Listad.TextMatrix(r_int_nroaux, 4)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 6) = grd_Listad.TextMatrix(r_int_nroaux, 5)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 7) = grd_Listad.TextMatrix(r_int_nroaux, 6)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 8) = grd_Listad.TextMatrix(r_int_nroaux, 7)
      r_obj_Excel.ActiveSheet.Cells(r_int_ConVer, 9) = grd_Listad.TextMatrix(r_int_nroaux, 8)
      
      r_obj_Excel.Range(r_obj_Excel.Cells(r_int_ConVer, 3), r_obj_Excel.Cells(r_int_ConVer, 3)).HorizontalAlignment = xlHAlignCenter
      r_obj_Excel.Range(r_obj_Excel.Cells(r_int_ConVer, 4), r_obj_Excel.Cells(r_int_ConVer, 4)).HorizontalAlignment = xlHAlignLeft
      r_obj_Excel.Range(r_obj_Excel.Cells(r_int_ConVer, 5), r_obj_Excel.Cells(r_int_ConVer, 5)).HorizontalAlignment = xlHAlignCenter
      r_obj_Excel.Range(r_obj_Excel.Cells(r_int_ConVer, 6), r_obj_Excel.Cells(r_int_ConVer, 6)).HorizontalAlignment = xlHAlignRight
      r_obj_Excel.Range(r_obj_Excel.Cells(r_int_ConVer, 7), r_obj_Excel.Cells(r_int_ConVer, 7)).HorizontalAlignment = xlHAlignRight
      r_obj_Excel.Range(r_obj_Excel.Cells(r_int_ConVer, 8), r_obj_Excel.Cells(r_int_ConVer, 8)).HorizontalAlignment = xlHAlignRight
      r_obj_Excel.Range(r_obj_Excel.Cells(r_int_ConVer, 9), r_obj_Excel.Cells(r_int_ConVer, 9)).HorizontalAlignment = xlHAlignRight
      
      r_int_ConVer = r_int_ConVer + 1
      r_int_Contad = r_int_Contad + 1
      DoEvents
   Next
   
   r_obj_Excel.Range(r_obj_Excel.Cells(4, 1), r_obj_Excel.Cells(4, 9)).Select
   r_obj_Excel.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
   r_obj_Excel.Range(r_obj_Excel.Cells(5, 1), r_obj_Excel.Cells(5, 9)).Select
   r_obj_Excel.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
   r_obj_Excel.ActiveSheet.Range("A4:I4").Interior.Color = RGB(146, 208, 80)
   r_obj_Excel.Range(r_obj_Excel.Cells(4, 1), r_obj_Excel.Cells(4, 9)).Select
   
   r_obj_Excel.Range(r_obj_Excel.Cells(4, 11), r_obj_Excel.Cells(r_int_ConVer, 11)).Font.Name = "Arial"
   r_obj_Excel.Range(r_obj_Excel.Cells(4, 1), r_obj_Excel.Cells(r_int_ConVer, 11)).Font.Size = 9
      
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub fs_Exportar_CreTra()
Dim r_obj_Excel      As Excel.Application
Dim r_int_NroFil     As Integer
Dim r_int_nroaux     As Integer
   
   'Generando Excel
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.DisplayAlerts = False
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   r_int_NroFil = 4
   
   With r_obj_Excel.ActiveSheet
      .Cells(2, 1) = "REPORTE DE CREDITOS TRANSFERIDOS DEL " & ipp_FecIni.Text & " AL " & ipp_FecFin.Text
      .Range("A2:K2").HorizontalAlignment = xlHAlignCenter
      .Range("A2:K2").Font.Bold = True
      .Range(.Cells(2, 1), .Cells(2, 11)).Font.Size = 14
      .Range("A2:K2").MergeCells = True
      
      .Cells(r_int_NroFil, 1) = "ITEM":                    .Columns("A").ColumnWidth = 6
      .Cells(r_int_NroFil, 2) = "Nº OPERACION":            .Columns("B").ColumnWidth = 14
      .Cells(r_int_NroFil, 3) = "DNI":                     .Columns("C").ColumnWidth = 12
      .Cells(r_int_NroFil, 4) = "APELLIDOS Y NOMBRES":     .Columns("D").ColumnWidth = 40
      .Cells(r_int_NroFil, 5) = "PRODUCTO":                .Columns("E").ColumnWidth = 35
      .Cells(r_int_NroFil, 6) = "BANCO":                   .Columns("F").ColumnWidth = 30
      .Cells(r_int_NroFil, 7) = "FEC.TRANSF.":             .Columns("G").ColumnWidth = 12
      .Cells(r_int_NroFil, 8) = "MONTO ACT. TNC":          .Columns("H").ColumnWidth = 15
      .Cells(r_int_NroFil, 9) = "MONTO ACT. TC":           .Columns("I").ColumnWidth = 15
      .Cells(r_int_NroFil, 10) = "MONTO PAS. TNC":         .Columns("J").ColumnWidth = 15
      .Cells(r_int_NroFil, 11) = "MONTO PAS. TC":          .Columns("K").ColumnWidth = 15
      
      .Range("A4:K4").Interior.Color = RGB(146, 208, 80)
      
      .Cells(r_int_NroFil, 1).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 2).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 3).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 4).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 5).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 6).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 7).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 8).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 9).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 10).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 11).HorizontalAlignment = xlHAlignCenter
      
      .Columns("A").HorizontalAlignment = xlHAlignCenter
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("G").HorizontalAlignment = xlHAlignCenter
      .Columns("H").HorizontalAlignment = xlHAlignRight
      .Columns("I").HorizontalAlignment = xlHAlignRight
      .Columns("J").HorizontalAlignment = xlHAlignRight
      .Columns("K").HorizontalAlignment = xlHAlignRight
      
      .Cells(r_int_NroFil, 1).Font.Bold = True
      .Cells(r_int_NroFil, 2).Font.Bold = True
      .Cells(r_int_NroFil, 3).Font.Bold = True
      .Cells(r_int_NroFil, 4).Font.Bold = True
      .Cells(r_int_NroFil, 5).Font.Bold = True
      .Cells(r_int_NroFil, 6).Font.Bold = True
      .Cells(r_int_NroFil, 7).Font.Bold = True
      .Cells(r_int_NroFil, 8).Font.Bold = True
      .Cells(r_int_NroFil, 9).Font.Bold = True
      .Cells(r_int_NroFil, 10).Font.Bold = True
      .Cells(r_int_NroFil, 11).Font.Bold = True
 
      r_int_NroFil = r_int_NroFil + 1
      For r_int_nroaux = 1 To grd_Listad.Rows - 1
         .Cells(r_int_NroFil, 1) = grd_Listad.TextMatrix(r_int_nroaux, 0)
         .Cells(r_int_NroFil, 2) = "'" & (grd_Listad.TextMatrix(r_int_nroaux, 1))
         .Cells(r_int_NroFil, 3) = grd_Listad.TextMatrix(r_int_nroaux, 2)
         .Cells(r_int_NroFil, 4) = grd_Listad.TextMatrix(r_int_nroaux, 3)
         .Cells(r_int_NroFil, 5) = grd_Listad.TextMatrix(r_int_nroaux, 4)
         .Cells(r_int_NroFil, 6) = grd_Listad.TextMatrix(r_int_nroaux, 5)
         .Cells(r_int_NroFil, 7) = "'" & grd_Listad.TextMatrix(r_int_nroaux, 6)
         
         .Cells(r_int_NroFil, 8) = grd_Listad.TextMatrix(r_int_nroaux, 7)
         .Cells(r_int_NroFil, 9) = grd_Listad.TextMatrix(r_int_nroaux, 8)
         .Cells(r_int_NroFil, 10) = grd_Listad.TextMatrix(r_int_nroaux, 9)
         .Cells(r_int_NroFil, 11) = grd_Listad.TextMatrix(r_int_nroaux, 10)
         r_int_NroFil = r_int_NroFil + 1
      Next

      .Range(.Cells(4, 8), .Cells(r_int_NroFil, 12)).Font.Name = "Arial"
      .Range(.Cells(4, 1), .Cells(r_int_NroFil, 12)).Font.Size = 9
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub fs_Exportar_AlmExt()
Dim r_obj_Excel      As Excel.Application
Dim r_int_NroFil     As Integer
Dim r_int_nroaux     As Integer
   
   'Generando Excel
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.DisplayAlerts = False
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   r_int_NroFil = 4
   
   With r_obj_Excel.ActiveSheet
      .Cells(2, 1) = "REPORTE DE ALMACEN EXTERNO AL " & UCase(Format(date, "Long Date"))
      .Range("A2:K2").HorizontalAlignment = xlHAlignCenter
      .Range("A2:K2").Font.Bold = True
      .Range(.Cells(2, 1), .Cells(2, 11)).Font.Size = 14
      .Range("A2:K2").MergeCells = True
      
      .Cells(r_int_NroFil, 1) = "ITEM":                        .Columns("A").ColumnWidth = 6
      .Cells(r_int_NroFil, 2) = "NRO DE OPERACION":            .Columns("B").ColumnWidth = 17
      .Cells(r_int_NroFil, 3) = "DNI":                         .Columns("C").ColumnWidth = 12
      .Cells(r_int_NroFil, 4) = "APELLIDOS Y NOMBRES":         .Columns("D").ColumnWidth = 45
      .Cells(r_int_NroFil, 5) = "FECHA DESEMBOLSO":            .Columns("E").ColumnWidth = 16
      .Cells(r_int_NroFil, 6) = "TIPO DE MONEDA":              .Columns("F").ColumnWidth = 20
      .Cells(r_int_NroFil, 7) = "MONTO  DESEMBOLSO":           .Columns("G").ColumnWidth = 20
      .Cells(r_int_NroFil, 8) = "ESTADO":                      .Columns("H").ColumnWidth = 20
      .Cells(r_int_NroFil, 9) = "FECHA INSCRIPCION":           .Columns("I").ColumnWidth = 20
      .Cells(r_int_NroFil, 10) = "VINCULADO":                  .Columns("J").ColumnWidth = 20
      .Cells(r_int_NroFil, 11) = "CODIGO CAJA":                .Columns("K").ColumnWidth = 20
      
      .Range("A4:K4").Interior.Color = RGB(146, 208, 80)
      
      .Cells(r_int_NroFil, 1).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 2).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 3).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 4).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 5).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 6).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 7).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 8).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 9).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 10).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 11).HorizontalAlignment = xlHAlignCenter
      
      .Columns("B").HorizontalAlignment = xlHAlignCenter
      .Columns("C").HorizontalAlignment = xlHAlignCenter
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("H").HorizontalAlignment = xlHAlignCenter
      .Columns("I").HorizontalAlignment = xlHAlignCenter
      .Columns("J").HorizontalAlignment = xlHAlignCenter
      .Columns("K").HorizontalAlignment = xlHAlignCenter
      
      .Range(.Cells(r_int_NroFil, 1), .Cells(r_int_NroFil, 11)).Font.Bold = True
      r_int_NroFil = r_int_NroFil + 1
      
      For r_int_nroaux = 1 To grd_Listad.Rows - 1
         .Cells(r_int_NroFil, 1) = grd_Listad.TextMatrix(r_int_nroaux, 0)
         .Cells(r_int_NroFil, 2) = grd_Listad.TextMatrix(r_int_nroaux, 1)
         .Cells(r_int_NroFil, 3) = grd_Listad.TextMatrix(r_int_nroaux, 2)
         .Cells(r_int_NroFil, 4) = grd_Listad.TextMatrix(r_int_nroaux, 3)
         .Cells(r_int_NroFil, 5) = grd_Listad.TextMatrix(r_int_nroaux, 4)
         .Cells(r_int_NroFil, 6) = grd_Listad.TextMatrix(r_int_nroaux, 5)
         .Cells(r_int_NroFil, 7) = grd_Listad.TextMatrix(r_int_nroaux, 6)
         .Cells(r_int_NroFil, 8) = grd_Listad.TextMatrix(r_int_nroaux, 7)
         .Cells(r_int_NroFil, 9) = grd_Listad.TextMatrix(r_int_nroaux, 8)
         .Cells(r_int_NroFil, 10) = grd_Listad.TextMatrix(r_int_nroaux, 9)
         .Cells(r_int_NroFil, 11) = grd_Listad.TextMatrix(r_int_nroaux, 10)
         r_int_NroFil = r_int_NroFil + 1
      Next

      '.Range(.Cells(1, 11), .Cells(r_int_NroFil, 11)).Font.Name = "Arial"
      .Range(.Cells(4, 1), .Cells(r_int_NroFil, 11)).Font.Size = 9
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub
