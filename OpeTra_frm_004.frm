VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frm_Cob_MovDia_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   9735
   ClientLeft      =   2295
   ClientTop       =   630
   ClientWidth     =   10005
   Icon            =   "OpeTra_frm_004.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9735
   ScaleWidth      =   10005
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   9735
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   10005
      _Version        =   65536
      _ExtentX        =   17648
      _ExtentY        =   17171
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
      Begin Threed.SSPanel SSPanel2 
         Height          =   1095
         Left            =   30
         TabIndex        =   10
         Top             =   780
         Width           =   9915
         _Version        =   65536
         _ExtentX        =   17489
         _ExtentY        =   1931
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
         Begin VB.ComboBox cmb_CodBan 
            Height          =   315
            Left            =   1410
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   5715
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   9150
            Picture         =   "OpeTra_frm_004.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Salir"
            Top             =   60
            Width           =   675
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   675
            Left            =   8430
            Picture         =   "OpeTra_frm_004.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Limpiar Datos"
            Top             =   60
            Width           =   675
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   675
            Left            =   7710
            Picture         =   "OpeTra_frm_004.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Buscar Registros"
            Top             =   60
            Width           =   675
         End
         Begin VB.ComboBox cmb_TipMon 
            Height          =   315
            Left            =   1410
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   390
            Width           =   5715
         End
         Begin EditLib.fpDateTime ipp_FecPag 
            Height          =   315
            Left            =   1410
            TabIndex        =   2
            Top             =   720
            Width           =   1335
            _Version        =   196608
            _ExtentX        =   2355
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
         Begin VB.Label Label10 
            Caption         =   "Fecha Movim.:"
            Height          =   315
            Left            =   90
            TabIndex        =   24
            Top             =   720
            Width           =   1185
         End
         Begin VB.Label Label3 
            Caption         =   "Banco:"
            Height          =   315
            Left            =   90
            TabIndex        =   12
            Top             =   60
            Width           =   765
         End
         Begin VB.Label Label1 
            Caption         =   "Moneda:"
            Height          =   315
            Left            =   90
            TabIndex        =   11
            Top             =   390
            Width           =   765
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   705
         Left            =   30
         TabIndex        =   13
         Top             =   30
         Width           =   9915
         _Version        =   65536
         _ExtentX        =   17489
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
            Height          =   585
            Left            =   630
            TabIndex        =   14
            Top             =   30
            Width           =   6795
            _Version        =   65536
            _ExtentX        =   11986
            _ExtentY        =   1032
            _StockProps     =   15
            Caption         =   "Operaciones por Bancos - Movimiento Diario"
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
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
            Picture         =   "OpeTra_frm_004.frx":0A62
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   6915
         Left            =   30
         TabIndex        =   15
         Top             =   1920
         Width           =   9915
         _Version        =   65536
         _ExtentX        =   17489
         _ExtentY        =   12197
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
            Height          =   6135
            Left            =   60
            TabIndex        =   6
            Top             =   360
            Width           =   9795
            _ExtentX        =   17277
            _ExtentY        =   10821
            _Version        =   393216
            Rows            =   21
            Cols            =   5
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel SSPanel11 
            Height          =   285
            Left            =   90
            TabIndex        =   16
            Top             =   90
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro. Movim."
            ForeColor       =   16777215
            BackColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel SSPanel4 
            Height          =   285
            Left            =   1230
            TabIndex        =   17
            Top             =   90
            Width           =   1155
            _Version        =   65536
            _ExtentX        =   2037
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Hora Movim."
            ForeColor       =   16777215
            BackColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel SSPanel5 
            Height          =   285
            Left            =   2370
            TabIndex        =   18
            Top             =   90
            Width           =   3675
            _Version        =   65536
            _ExtentX        =   6482
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Tipo de Movimiento"
            ForeColor       =   16777215
            BackColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel SSPanel8 
            Height          =   285
            Left            =   7950
            TabIndex        =   19
            Top             =   90
            Width           =   1575
            _Version        =   65536
            _ExtentX        =   2778
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Importe"
            ForeColor       =   16777215
            BackColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel SSPanel10 
            Height          =   285
            Left            =   6030
            TabIndex        =   20
            Top             =   90
            Width           =   1935
            _Version        =   65536
            _ExtentX        =   3413
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Número Referencia"
            ForeColor       =   16777215
            BackColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
         End
         Begin Threed.SSPanel pnl_TotDia 
            Height          =   315
            Left            =   7950
            TabIndex        =   21
            Top             =   6510
            Width           =   1575
            _Version        =   65536
            _ExtentX        =   2778
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   16777215
            BackColor       =   192
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            Outline         =   -1  'True
            Alignment       =   4
         End
         Begin VB.Label Label5 
            Caption         =   "Totales ==>"
            Height          =   285
            Left            =   6930
            TabIndex        =   22
            Top             =   6540
            Width           =   975
         End
      End
      Begin Threed.SSPanel SSPanel12 
         Height          =   795
         Left            =   30
         TabIndex        =   23
         Top             =   8880
         Width           =   9915
         _Version        =   65536
         _ExtentX        =   17489
         _ExtentY        =   1402
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
         Begin VB.CommandButton cmd_VerCom 
            Height          =   675
            Left            =   9180
            Picture         =   "OpeTra_frm_004.frx":0D6C
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Ver Comprobante"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Imprim 
            Height          =   675
            Left            =   60
            Picture         =   "OpeTra_frm_004.frx":1076
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Imprimir Listado"
            Top             =   30
            Width           =   675
         End
      End
   End
End
Attribute VB_Name = "frm_Cob_MovDia_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_arr_CodBan()   As moddat_tpo_Genera

Private Sub cmb_TipMon_Click()
   Call gs_SetFocus(ipp_FecPag)
End Sub

Private Sub cmb_TipMon_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_TipMon_Click
   End If
End Sub

Private Sub cmb_CodBan_Click()
   Call gs_SetFocus(cmb_TipMon)
End Sub

Private Sub cmb_CodBan_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmb_CodBan_Click
   End If
End Sub

Private Sub cmd_Buscar_Click()
   If cmb_CodBan.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Banco del que desea ver las Operaciones.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_CodBan)
      Exit Sub
   End If

   If cmb_TipMon.ListIndex = -1 Then
      MsgBox "Debe seleccionar el Tipo de Moneda del que desea ver las Operaciones.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(cmb_TipMon)
      Exit Sub
   End If
   
   Call fs_Buscar
End Sub

Private Sub cmd_Limpia_Click()
   cmb_CodBan.ListIndex = -1
   cmb_TipMon.ListIndex = -1
   ipp_FecPag.Text = Format(Date, "dd/mm/yyyy")
   
   pnl_TotDia.Caption = "0.00 "
   Call gs_LimpiaGrid(grd_Listad)
   
   Call fs_Activa(True)
   Call gs_SetFocus(cmb_CodBan)
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub cmd_VerCom_Click()
   Dim r_int_TipOpe As String
   
   If grd_Listad.Rows > 0 Then
   End If
   
   grd_Listad.Col = 2
   
   grd_Listad.Col = 0
   opecaj_g_str_NumMov = grd_Listad.Text
   
   Call gs_RefrescaGrid(grd_Listad)
   
   opecaj_g_int_FlgAct = 1
   frm_Cob_MovDia_02.Show 1
   
   If opecaj_g_int_FlgAct = 2 Then
      Call cmd_Buscar_Click
   End If
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call cmd_Limpia_Click
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Activa(ByVal p_Activa As Integer)
   cmb_CodBan.Enabled = p_Activa
   cmb_TipMon.Enabled = p_Activa
   ipp_FecPag.Enabled = p_Activa
   
   cmd_Buscar.Enabled = p_Activa
   
   grd_Listad.Enabled = Not p_Activa
   cmd_Imprim.Enabled = Not p_Activa
   cmd_VerCom.Enabled = Not p_Activa
End Sub

Private Sub fs_Inicia()
   'Inicializando Rejilla
   grd_Listad.ColWidth(0) = 1145
   grd_Listad.ColWidth(1) = 1145
   grd_Listad.ColWidth(2) = 3665
   grd_Listad.ColWidth(3) = 1925
   grd_Listad.ColWidth(4) = 1565
   
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignLeftCenter
   grd_Listad.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad.ColAlignment(4) = flexAlignRightCenter

   Call moddat_gs_FecSis
   
   Call moddat_gs_Carga_LisIte(cmb_CodBan, l_arr_CodBan, 1, "516")
   Call moddat_gs_Carga_LisIte_Combo(cmb_TipMon, 1, "204")
End Sub

Private Sub fs_Buscar()
   Dim r_str_HorMov     As String
   Dim r_dbl_Import     As Double
   
   r_dbl_Import = 0
   
   Call moddat_gs_FecSis
   Call gs_LimpiaGrid(grd_Listad)

   opecaj_g_str_CodBan = l_arr_CodBan(cmb_CodBan.ListIndex + 1).Genera_Codigo
   opecaj_g_str_FecMov = Format(CDate(ipp_FecPag.Text), "yyyymmdd")
   
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT * FROM OPE_CAJMOV WHERE "
   g_str_Parame = g_str_Parame & "CAJMOV_SUCMOV = '" & modgen_g_str_CodSuc & "' AND "
   g_str_Parame = g_str_Parame & "CAJMOV_MONPAG = " & CStr(cmb_TipMon.ItemData(cmb_TipMon.ListIndex)) & " AND "
   g_str_Parame = g_str_Parame & "CAJMOV_CODBAN = '" & opecaj_g_str_CodBan & "' AND "
   g_str_Parame = g_str_Parame & "CAJMOV_FECMOV = " & opecaj_g_str_FecMov
   g_str_Parame = g_str_Parame & "ORDER BY CAJMOV_NUMMOV ASC "

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
       Exit Sub
   End If

   If g_rst_Princi.BOF And g_rst_Princi.EOF Then
     g_rst_Princi.Close
     Set g_rst_Princi = Nothing
     
     MsgBox "No se han encontrado registros.", vbExclamation, modgen_g_con_PltPar
     Call gs_SetFocus(cmb_CodBan)
     Exit Sub
   End If
   
   Call fs_Activa(False)
   
   grd_Listad.Redraw = False
   
   g_rst_Princi.MoveFirst
   Do While Not g_rst_Princi.EOF
      grd_Listad.Rows = grd_Listad.Rows + 1
      
      grd_Listad.Row = grd_Listad.Rows - 1
      
      grd_Listad.Col = 0
      grd_Listad.Text = Format(g_rst_Princi!CAJMOV_NUMMOV, "00000")
      
      r_str_HorMov = Format(g_rst_Princi!CAJMOV_HORMOV, "000000")
      r_str_HorMov = Mid(r_str_HorMov, 1, 2) & ":" & Mid(r_str_HorMov, 3, 2) & ":" & Mid(r_str_HorMov, 5, 2)
      
      grd_Listad.Col = 1
      grd_Listad.Text = r_str_HorMov
      
      grd_Listad.Col = 2
      grd_Listad.Text = CStr(g_rst_Princi!CAJMOV_TIPMOV) & " - " & moddat_gf_Consulta_ParDes("301", Format(g_rst_Princi!CAJMOV_TIPMOV, "000000"))
      
      grd_Listad.Col = 3
      grd_Listad.Text = Trim(g_rst_Princi!CAJMOV_NUMOPE & "")
      
      grd_Listad.Col = 4
      grd_Listad.Text = Format(g_rst_Princi!CAJMOV_IMPTOT, "###,###,##0.00")
      
      If Left(CStr(g_rst_Princi!CAJMOV_TIPMOV), 1) = "2" Then
         r_dbl_Import = r_dbl_Import - g_rst_Princi!CAJMOV_IMPTOT
      Else
         r_dbl_Import = r_dbl_Import + g_rst_Princi!CAJMOV_IMPTOT
      End If
      
      g_rst_Princi.MoveNext
   Loop
   
   pnl_TotDia.Caption = Format(r_dbl_Import, "###,###,##0.00") & " "
   
   grd_Listad.Redraw = True
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   If grd_Listad.Rows > 0 Then
      cmd_Imprim.Enabled = True
      cmd_VerCom.Enabled = True
      grd_Listad.Enabled = True
   End If
   
   Call gs_UbiIniGrid(grd_Listad)
   Call gs_SetFocus(grd_Listad)
End Sub

Private Sub grd_Listad_DblClick()
   Call cmd_VerCom_Click
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

Private Sub ipp_FecPag_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call gs_SetFocus(cmd_Buscar)
   End If
End Sub
