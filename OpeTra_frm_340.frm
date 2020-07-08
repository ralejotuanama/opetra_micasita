VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_Rpt_PagPrv_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form4"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11145
   Icon            =   "OpeTra_frm_340.frx":0000
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   11145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   8565
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   11160
      _Version        =   65536
      _ExtentX        =   19685
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
         Width           =   11040
         _Version        =   65536
         _ExtentX        =   19473
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
            Caption         =   "Reporte de Gastos Cierre - Pago a Proveedores"
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
            Picture         =   "OpeTra_frm_340.frx":000C
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   645
         Left            =   60
         TabIndex        =   10
         Top             =   780
         Width           =   11040
         _Version        =   65536
         _ExtentX        =   19473
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
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   630
            Picture         =   "OpeTra_frm_340.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Limpiar Datos"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExc 
            Height          =   585
            Left            =   1230
            Picture         =   "OpeTra_frm_340.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   10410
            Picture         =   "OpeTra_frm_340.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_340.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Buscar pagos"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel24 
         Height          =   555
         Left            =   60
         TabIndex        =   11
         Top             =   1470
         Width           =   11040
         _Version        =   65536
         _ExtentX        =   19473
         _ExtentY        =   979
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
            Left            =   7890
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   120
            Width           =   1575
         End
         Begin EditLib.fpDateTime ipp_FecIni 
            Height          =   315
            Left            =   1470
            TabIndex        =   0
            Top             =   120
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
            Left            =   4590
            TabIndex        =   1
            Top             =   120
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
         Begin VB.Label Label5 
            Caption         =   "Tipo de Formato:"
            Height          =   285
            Left            =   6540
            TabIndex        =   15
            Top             =   180
            Width           =   1365
         End
         Begin VB.Label Label20 
            Caption         =   "Fecha de Fin:"
            Height          =   195
            Left            =   3450
            TabIndex        =   13
            Top             =   180
            Width           =   1125
         End
         Begin VB.Label Label2 
            Caption         =   "Fecha de Inicio:"
            Height          =   225
            Left            =   150
            TabIndex        =   12
            Top             =   180
            Width           =   1215
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   6405
         Left            =   60
         TabIndex        =   14
         Top             =   2070
         Width           =   11040
         _Version        =   65536
         _ExtentX        =   19473
         _ExtentY        =   11298
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
            Height          =   6285
            Left            =   30
            TabIndex        =   6
            Top             =   60
            Width           =   10980
            _ExtentX        =   19368
            _ExtentY        =   11086
            _Version        =   393216
            Rows            =   25
            Cols            =   7
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            SelectionMode   =   1
         End
      End
   End
End
Attribute VB_Name = "frm_Rpt_PagPrv_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Buscar_Click()
   'valida datos
   If CDate(ipp_FecFin.Text) < CDate(ipp_FecIni.Text) Then
      MsgBox "La fecha de fin no puede ser menor a la fecha de inicio.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(ipp_FecIni)
      Exit Sub
   End If
   If cmb_TipRep.ListIndex = -1 Then
      MsgBox "Debe seleccionar el tipo de reporte.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipRep)
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de procesar reporte?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   Call fs_BusPagos
   Call fs_Habilita(True)
   Screen.MousePointer = 0
End Sub

Private Sub cmd_ExpExc_Click()
   'valida datos
   If grd_Listad.Rows = 0 Then
      MsgBox "No existe información a exportar.", vbExclamation, modgen_g_str_NomPlt
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de exportar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   If cmb_TipRep.ListIndex = 0 Then
      Call fs_Exportar
   Else
      Call fs_Exportar2
   End If
   Screen.MousePointer = 0
End Sub

Private Sub cmd_Limpia_Click()
   Call fs_Inicia
   Call fs_Limpia
   Call fs_Habilita(False)
   Call gs_SetFocus(ipp_FecIni)
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
   Call fs_Habilita(False)
   
   Call gs_SetFocus(ipp_FecIni)
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   grd_Listad.Cols = 5
   
   cmb_TipRep.Clear
   cmb_TipRep.AddItem "GENERAL"
   cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = 0
   cmb_TipRep.AddItem "DETALLADO"
   cmb_TipRep.ItemData(cmb_TipRep.NewIndex) = 1
   cmb_TipRep.ListIndex = -1
   
   'Inicializando Controles
   cmb_TipRep.ListIndex = 0
End Sub

Private Sub fs_Limpia()
   Call gs_LimpiaGrid(grd_Listad)
   ipp_FecIni.Text = Format(date - CDate(30), "dd/mm/yyyy")
   ipp_FecFin.Text = Format(date, "dd/mm/yyyy")
End Sub

Private Sub fs_Habilita(ByVal p_Habilita As Boolean)
   cmd_Buscar.Enabled = Not p_Habilita
   cmd_ExpExc.Enabled = p_Habilita
   ipp_FecIni.Enabled = Not p_Habilita
   ipp_FecFin.Enabled = Not p_Habilita
   cmb_TipRep.Enabled = Not p_Habilita
   grd_Listad.Enabled = p_Habilita
End Sub

Private Sub fs_BusPagos()
Dim str_FecIni       As String
Dim str_FecFin       As String

   Call gs_LimpiaGrid(grd_Listad)
   grd_Listad.Rows = grd_Listad.Rows + 1
   grd_Listad.Row = grd_Listad.Rows - 1
   str_FecIni = Format(CDate(ipp_FecIni.Text), "yyyymmdd")
   str_FecFin = Format(CDate(ipp_FecFin.Text), "yyyymmdd")
   
   g_str_Parame = "USP_TRA_ADMGAS_CONSULTA ("
   g_str_Parame = g_str_Parame & " '" & Format(CDate(DateAdd("D", -1, ipp_FecIni.Text)), "yyyymmdd") & "' , "
   g_str_Parame = g_str_Parame & " '" & str_FecIni & "' , "
   g_str_Parame = g_str_Parame & " '" & str_FecFin & "' , "
   g_str_Parame = g_str_Parame & " " & cmb_TipRep.ItemData(cmb_TipRep.ListIndex) & " , "
   g_str_Parame = g_str_Parame & "  ) "
    
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   If (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.Close
      Set g_rst_Princi = Nothing
      Exit Sub
   End If
      
   If cmb_TipRep.ListIndex = 0 Then
      grd_Listad.Cols = 5
      grd_Listad.ColAlignment(1) = flexAlignCenterCenter
      grd_Listad.ColAlignment(2) = flexAlignCenterCenter
      grd_Listad.ColAlignment(3) = flexAlignCenterCenter
      grd_Listad.ColAlignment(4) = flexAlignCenterCenter
      grd_Listad.ColWidth(0) = 0
      grd_Listad.ColWidth(1) = 2500
      grd_Listad.ColWidth(2) = 2500
      grd_Listad.ColWidth(3) = 2500
      grd_Listad.ColWidth(4) = 2500
      
      grd_Listad.Col = 1
      grd_Listad.Text = "Saldo Inicial"
      grd_Listad.Col = 2
      grd_Listad.Text = "Pagos Cliente"
      grd_Listad.Col = 3
      grd_Listad.Text = "Pagos Proveedor"
      grd_Listad.Col = 4
      grd_Listad.Text = "Saldo Final"
      
      grd_Listad.Rows = grd_Listad.Rows + 1
      grd_Listad.Row = grd_Listad.Rows - 1
            
      grd_Listad.Col = 1
      grd_Listad.Text = Format(g_rst_Princi!dbl_SldIni, "###,###,##0.00")
      grd_Listad.Col = 2
      grd_Listad.Text = Format(g_rst_Princi!dbl_PagCli, "###,###,##0.00")
      grd_Listad.Col = 3
      grd_Listad.Text = Format(g_rst_Princi!dbl_PagPrv, "###,###,##0.00")
      grd_Listad.Col = 4
      grd_Listad.Text = Format(g_rst_Princi!dbl_SldFin, "###,###,##0.00")
   
   Else
   
      grd_Listad.Cols = 14
      grd_Listad.ColAlignment(1) = flexAlignCenterCenter
      grd_Listad.ColAlignment(2) = flexAlignCenterCenter
      grd_Listad.ColAlignment(3) = flexAlignLeftCenter
      grd_Listad.ColAlignment(4) = flexAlignLeftCenter
      grd_Listad.ColAlignment(5) = flexAlignCenterCenter
      grd_Listad.ColAlignment(6) = flexAlignCenterCenter
      grd_Listad.ColAlignment(8) = flexAlignCenterCenter
      grd_Listad.ColAlignment(9) = flexAlignRightCenter
      grd_Listad.ColAlignment(10) = flexAlignCenterCenter
      grd_Listad.ColAlignment(11) = flexAlignCenterCenter
      grd_Listad.ColAlignment(12) = flexAlignCenterCenter
      grd_Listad.ColAlignment(13) = flexAlignCenterCenter
      
      grd_Listad.ColWidth(0) = 0
      grd_Listad.ColWidth(1) = 1500
      grd_Listad.ColWidth(2) = 1200
      grd_Listad.ColWidth(3) = 3800
      grd_Listad.ColWidth(4) = 3500
      grd_Listad.ColWidth(5) = 1800
      grd_Listad.ColWidth(6) = 1500
      grd_Listad.ColWidth(7) = 1600
      grd_Listad.ColWidth(8) = 1800
      grd_Listad.ColWidth(9) = 1800
      grd_Listad.ColWidth(10) = 1800
      grd_Listad.ColWidth(11) = 1800
      grd_Listad.ColWidth(12) = 3000
      grd_Listad.ColWidth(13) = 1800
      
      grd_Listad.Col = 1
      grd_Listad.Text = "N° Solicitud"
      grd_Listad.Col = 2
      grd_Listad.Text = "N° DNI"
      grd_Listad.Col = 3
      grd_Listad.Text = "Nombres y Apellidos"
      grd_Listad.Col = 4
      grd_Listad.Text = "Concepto de Gasto"
      grd_Listad.Col = 5
      grd_Listad.Text = "Tipo de Moneda"
      grd_Listad.Col = 6
      grd_Listad.Text = "Fecha Pago Cliente"
      grd_Listad.Col = 7
      grd_Listad.Text = "Monto Pago Cliente"
      grd_Listad.Col = 8
      grd_Listad.Text = "Fecha Pago Proveedor"
      grd_Listad.Col = 9
      grd_Listad.Text = "Monto Pago Proveedor"
      grd_Listad.Col = 10
      grd_Listad.Text = "Tipo de Pago"
      grd_Listad.Col = 11
      grd_Listad.Text = "Numero de Documento"
      grd_Listad.Col = 12
      grd_Listad.Text = "Tipo de Garantía"
      grd_Listad.Col = 13
      grd_Listad.Text = "Situación"
      
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         
         grd_Listad.Col = 1
         grd_Listad.Text = gf_Formato_NumSol(g_rst_Princi!Numsol)
         
         grd_Listad.Col = 2
         grd_Listad.Text = CStr(g_rst_Princi!numdoc)
         
         grd_Listad.Col = 3
         grd_Listad.Text = g_rst_Princi!nomcli
         
         grd_Listad.Col = 4
         grd_Listad.Text = g_rst_Princi!concepto
         
         grd_Listad.Col = 5
         grd_Listad.Text = g_rst_Princi!moneda
         
         If (CStr(g_rst_Princi!fecpag_cli) >= CStr(str_FecIni)) And (CStr(g_rst_Princi!fecpag_cli) <= CStr(str_FecFin)) Then
            grd_Listad.Col = 6
            If Not IsNull(g_rst_Princi!fecpag_cli) Then
               grd_Listad.Text = (gf_FormatoFecha(g_rst_Princi!fecpag_cli))
            Else
               grd_Listad.Text = ""
            End If
            
            grd_Listad.Col = 7
            grd_Listad.Text = Format(g_rst_Princi!PagCli, "###,###,##0.00")
         End If
         
         If Not IsNull(g_rst_Princi!fecpag_prv) Then
            If (CStr(g_rst_Princi!fecpag_prv) >= CStr(str_FecIni)) And (CStr(g_rst_Princi!fecpag_prv) <= CStr(str_FecFin)) Then
               grd_Listad.Col = 8
               grd_Listad.Text = gf_FormatoFecha(g_rst_Princi!fecpag_prv)
               
               grd_Listad.Col = 9
               grd_Listad.Text = Format(g_rst_Princi!PagPrv, "###,###,##0.00")
               
               grd_Listad.Col = 10
               If IsNull(g_rst_Princi!TIPO_PAGO) Then
                  grd_Listad.Text = ""
               Else
                  grd_Listad.Text = Trim(g_rst_Princi!TIPO_PAGO)
               End If
               
               grd_Listad.Col = 11
               If IsNull(g_rst_Princi!COMPROBANTE) Then
                  grd_Listad.Text = ""
               Else
                  grd_Listad.Text = Trim(g_rst_Princi!COMPROBANTE)
               End If
            End If
         End If
         
         grd_Listad.Col = 12
         If IsNull(g_rst_Princi!tipgar) Then
            grd_Listad.Text = ""
         Else
            grd_Listad.Text = Trim(g_rst_Princi!tipgar)
         End If
         
         If IsNull(g_rst_Princi!situac) Then
            grd_Listad.Col = 13
            grd_Listad.Text = ""
         Else
            grd_Listad.Col = 13
            grd_Listad.Text = Trim(g_rst_Princi!situac)
         End If
         g_rst_Princi.MoveNext
      Loop
      Call gs_UbiIniGrid(grd_Listad)
   End If
      
   grd_Listad.FixedRows = 1
   Call gs_UbiIniGrid(grd_Listad)
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub fs_Exportar()
   Dim r_obj_Excel      As Excel.Application
   Dim r_int_NroFil     As Integer
   Dim r_int_nroaux     As Integer
   
   'Generando Excel
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.DisplayAlerts = False
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   r_int_NroFil = 6
   
   With r_obj_Excel.ActiveSheet
      .Range("A2:D2").Merge
      .Range("A2:D2").HorizontalAlignment = xlHAlignCenter

      .Cells(2, 1) = "REPORTE GENERAL DE PAGO DE GASTOS DE CIERRE A PROVEEDORES"
      .Cells(2, 1).Font.Bold = True
      .Cells(4, 1) = "FECHA INICIO :"
      .Cells(4, 2) = ipp_FecIni.Text
      .Cells(4, 3) = "FECHA FINAL :"
      .Cells(4, 4) = ipp_FecFin.Text
      
      .Cells(r_int_NroFil, 1).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 2).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 3).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 4).HorizontalAlignment = xlHAlignCenter
      
      .Cells(r_int_NroFil, 1) = "SALDO INICIAL":         .Columns("A").ColumnWidth = 20
      .Cells(r_int_NroFil, 2) = "SALDO CLIENTE":         .Columns("B").ColumnWidth = 20
      .Cells(r_int_NroFil, 3) = "SALDO PROVEEDOR":       .Columns("C").ColumnWidth = 20
      .Cells(r_int_NroFil, 4) = "SALDO FINAL":           .Columns("D").ColumnWidth = 20
      
      .Cells(r_int_NroFil, 1).Font.Bold = True
      .Cells(r_int_NroFil, 2).Font.Bold = True
      .Cells(r_int_NroFil, 3).Font.Bold = True
      .Cells(r_int_NroFil, 4).Font.Bold = True
       
      r_int_NroFil = r_int_NroFil + 1
      For r_int_nroaux = 1 To grd_Listad.Rows - 1
         .Cells(r_int_NroFil, 1) = grd_Listad.TextMatrix(r_int_nroaux, 1)
         .Cells(r_int_NroFil, 2) = grd_Listad.TextMatrix(r_int_nroaux, 2)
         .Cells(r_int_NroFil, 3) = grd_Listad.TextMatrix(r_int_nroaux, 3)
         .Cells(r_int_NroFil, 4) = grd_Listad.TextMatrix(r_int_nroaux, 4)
         r_int_NroFil = r_int_NroFil + 1
      Next
      
      .Cells(r_int_NroFil, 1).HorizontalAlignment = xlHAlignRight
      .Cells(r_int_NroFil, 2).HorizontalAlignment = xlHAlignRight
      .Cells(r_int_NroFil, 3).HorizontalAlignment = xlHAlignRight
      .Cells(r_int_NroFil, 4).HorizontalAlignment = xlHAlignRight
      
      .Range(.Cells(6, 1), .Cells(9, 1)).Select
      r_obj_Excel.Selection.NumberFormat = "###,##0.00"

      r_int_NroFil = r_int_NroFil + 2
       
      .Range(.Cells(1, 8), .Cells(r_int_NroFil, 8)).Font.Name = "Arial"
      .Range(.Cells(1, 1), .Cells(r_int_NroFil, 8)).Font.Size = 9
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub fs_Exportar2()
Dim r_obj_Excel      As Excel.Application
Dim r_int_NroFil     As Integer
Dim r_int_nroaux     As Integer
   
   'Generando Excel
   Set r_obj_Excel = New Excel.Application
   r_obj_Excel.DisplayAlerts = False
   r_obj_Excel.SheetsInNewWorkbook = 1
   r_obj_Excel.Workbooks.Add
   r_int_NroFil = 6
   
   With r_obj_Excel.ActiveSheet
      .Range("A2:M2").Merge
      .Range("A2:M2").HorizontalAlignment = xlHAlignCenter
      .Cells(2, 1) = "REPORTE DETALLADO DE PAGO DE GASTOS DE CIERRE A PROVEEDORES"
      .Cells(2, 1).Font.Bold = True
      
      .Range("B4:C4").Merge
      .Range("B4:C4").HorizontalAlignment = xlHAlignCenter
      .Cells(4, 2) = "FECHA INICIO :  " + ipp_FecIni.Text
      
      .Range("E4:F4").Merge
      .Range("E4:F4").HorizontalAlignment = xlHAlignCenter
      .Cells(4, 5) = "FECHA FINAL :  " + ipp_FecFin.Text
      
      .Range("A6:M6").HorizontalAlignment = xlHAlignCenter
      .Columns("D").HorizontalAlignment = xlHAlignCenter
      .Columns("E").HorizontalAlignment = xlHAlignCenter
      .Columns("F").HorizontalAlignment = xlHAlignCenter
      .Columns("H").HorizontalAlignment = xlHAlignCenter
      .Columns("H").HorizontalAlignment = xlHAlignCenter
      .Columns("J").HorizontalAlignment = xlHAlignCenter
      .Columns("K").HorizontalAlignment = xlHAlignCenter
      .Columns("L").HorizontalAlignment = xlHAlignCenter
      .Columns("M").HorizontalAlignment = xlHAlignCenter
      
      .Cells(r_int_NroFil, 1) = "SOLICITUD":                   .Columns("A").ColumnWidth = 14
      .Cells(r_int_NroFil, 2) = "DNI":                         .Columns("B").ColumnWidth = 10
      .Cells(r_int_NroFil, 3) = "APELLIDOS Y NOMBRES":         .Columns("C").ColumnWidth = 36
      .Cells(r_int_NroFil, 4) = "CONCEPTO DE GASTO":           .Columns("D").ColumnWidth = 40
      .Cells(r_int_NroFil, 5) = "TIPO DE MONEDA":              .Columns("E").ColumnWidth = 15
      .Cells(r_int_NroFil, 6) = "FEC. PAGO CLIENTE":           .Columns("F").ColumnWidth = 15
      .Cells(r_int_NroFil, 7) = "MONTO PAGO CLIENTE":          .Columns("G").ColumnWidth = 17
      .Cells(r_int_NroFil, 8) = "FEC. PAGO PROVEEDOR":         .Columns("H").ColumnWidth = 17
      .Cells(r_int_NroFil, 9) = "MONTO PAGO PROVEEDOR":        .Columns("I").ColumnWidth = 20
      .Cells(r_int_NroFil, 10) = "TIPO DE PAGO":               .Columns("J").ColumnWidth = 15
      .Cells(r_int_NroFil, 11) = "NUMERO DOCUMENTO":           .Columns("K").ColumnWidth = 17
      .Cells(r_int_NroFil, 12) = "TIPO DE GARANTIA":           .Columns("L").ColumnWidth = 30
      .Cells(r_int_NroFil, 13) = "SITUACION":                  .Columns("M").ColumnWidth = 18
      
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
      .Cells(r_int_NroFil, 12).Font.Bold = True
      .Cells(r_int_NroFil, 13).Font.Bold = True
      
      r_int_NroFil = r_int_NroFil + 1
      For r_int_nroaux = 1 To grd_Listad.Rows - 1
          .Cells(r_int_NroFil, 2).Select
          r_obj_Excel.Selection.NumberFormat = "@"
          .Cells(r_int_NroFil, 6).Select
          r_obj_Excel.Selection.NumberFormat = "DD/MM/YYYY"
         .Cells(r_int_NroFil, 7).Select
          r_obj_Excel.Selection.NumberFormat = "###,##0.00"
          .Cells(r_int_NroFil, 8).Select
          r_obj_Excel.Selection.NumberFormat = "DD/MM/YYYY"
         .Cells(r_int_NroFil, 9).Select
          r_obj_Excel.Selection.NumberFormat = "###,##0.00"
          
         .Cells(r_int_NroFil, 1) = CStr(grd_Listad.TextMatrix(r_int_nroaux, 1))
         .Cells(r_int_NroFil, 2) = grd_Listad.TextMatrix(r_int_nroaux, 2)
         .Cells(r_int_NroFil, 3) = grd_Listad.TextMatrix(r_int_nroaux, 3)
         .Cells(r_int_NroFil, 4) = grd_Listad.TextMatrix(r_int_nroaux, 4)
         .Cells(r_int_NroFil, 5) = grd_Listad.TextMatrix(r_int_nroaux, 5)
         
         If Len(grd_Listad.TextMatrix(r_int_nroaux, 6)) = 10 Then
            .Cells(r_int_NroFil, 6) = CDate(grd_Listad.TextMatrix(r_int_nroaux, 6))
         Else
            .Cells(r_int_NroFil, 6) = ""
         End If
         .Cells(r_int_NroFil, 7) = grd_Listad.TextMatrix(r_int_nroaux, 7)
         If Len(Trim(grd_Listad.TextMatrix(r_int_nroaux, 8))) = 10 Then
            .Cells(r_int_NroFil, 8) = CDate(grd_Listad.TextMatrix(r_int_nroaux, 8))
         Else
            .Cells(r_int_NroFil, 8) = ""
         End If
         
         .Cells(r_int_NroFil, 9) = grd_Listad.TextMatrix(r_int_nroaux, 9)
         .Cells(r_int_NroFil, 10) = grd_Listad.TextMatrix(r_int_nroaux, 10)
         .Cells(r_int_NroFil, 11) = grd_Listad.TextMatrix(r_int_nroaux, 11)
         .Cells(r_int_NroFil, 12) = grd_Listad.TextMatrix(r_int_nroaux, 12)
         .Cells(r_int_NroFil, 13) = grd_Listad.TextMatrix(r_int_nroaux, 13)
         r_int_NroFil = r_int_NroFil + 1
      Next
      
      .Range(.Cells(7, 1), .Cells(r_int_NroFil, 2)).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 3).HorizontalAlignment = xlHAlignLeft
      .Cells(r_int_NroFil, 4).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 5).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 6).HorizontalAlignment = xlHAlignRight
      .Cells(r_int_NroFil, 7).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 8).HorizontalAlignment = xlHAlignRight
      .Cells(r_int_NroFil, 7).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 10).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 11).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 12).HorizontalAlignment = xlHAlignCenter
      .Cells(r_int_NroFil, 13).HorizontalAlignment = xlHAlignCenter
      
      r_int_NroFil = r_int_NroFil + 2
       
      .Range(.Cells(1, 14), .Cells(r_int_NroFil, 14)).Font.Name = "Arial"
      .Range(.Cells(1, 1), .Cells(r_int_NroFil, 14)).Font.Size = 9
   End With
   
   r_obj_Excel.Visible = True
   Set r_obj_Excel = Nothing
End Sub

Private Sub ipp_FecIni_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(ipp_FecFin)
   End If
End Sub

Private Sub ipp_FecFin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmb_TipRep)
   End If
End Sub

Private Sub cmb_TipRep_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Buscar)
   End If
End Sub

