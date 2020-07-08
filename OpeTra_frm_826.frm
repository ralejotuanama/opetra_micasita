VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_Ges_TecPro_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8820
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   19140
   Icon            =   "OpeTra_frm_826.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8820
   ScaleWidth      =   19140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Height          =   8775
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   19125
      _Version        =   65536
      _ExtentX        =   33734
      _ExtentY        =   15478
      _StockProps     =   15
      BackColor       =   14215660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txt_Buscar 
         Height          =   315
         Left            =   5700
         MaxLength       =   100
         TabIndex        =   15
         Top             =   8310
         Width           =   6975
      End
      Begin VB.ComboBox cmb_Buscar 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   8280
         Width           =   2595
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   30
         TabIndex        =   3
         Top             =   750
         Width           =   19035
         _Version        =   65536
         _ExtentX        =   33576
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
         Begin VB.CommandButton cmd_Editar 
            Height          =   585
            Left            =   1800
            Picture         =   "OpeTra_frm_826.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Modificar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Agrega 
            Height          =   585
            Left            =   1200
            Picture         =   "OpeTra_frm_826.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Nuevo"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Limpia 
            Height          =   585
            Left            =   600
            Picture         =   "OpeTra_frm_826.frx":0620
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Limpiar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Buscar 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_826.frx":092A
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Buscar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_NueGar 
            Height          =   585
            Left            =   3600
            Picture         =   "OpeTra_frm_826.frx":0C34
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Garantía"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Evalua 
            Height          =   585
            Left            =   3000
            Picture         =   "OpeTra_frm_826.frx":0F3E
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Registrar Informe de Tasación"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Borrar 
            Height          =   585
            Left            =   2400
            Picture         =   "OpeTra_frm_826.frx":1808
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Borrar"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ResEte 
            Height          =   585
            Left            =   5400
            Picture         =   "OpeTra_frm_826.frx":1B12
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Posición Consolidada"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_ExpExcRes 
            Height          =   585
            Left            =   4800
            Picture         =   "OpeTra_frm_826.frx":1E1C
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Exportar a Excel"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_CarFia 
            Height          =   585
            Left            =   4200
            Picture         =   "OpeTra_frm_826.frx":2126
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Detalle"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   585
            Left            =   18420
            Picture         =   "OpeTra_frm_826.frx":29F0
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Salida"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   4
         Top             =   30
         Width           =   19035
         _Version        =   65536
         _ExtentX        =   33576
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
            Height          =   585
            Left            =   660
            TabIndex        =   5
            Top             =   30
            Width           =   3975
            _Version        =   65536
            _ExtentX        =   7011
            _ExtentY        =   1032
            _StockProps     =   15
            Caption         =   "Techo Propio - Entidades Técnicas"
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
            Picture         =   "OpeTra_frm_826.frx":2E32
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   6765
         Left            =   30
         TabIndex        =   6
         Top             =   1440
         Width           =   19035
         _Version        =   65536
         _ExtentX        =   33576
         _ExtentY        =   11933
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
            Left            =   30
            TabIndex        =   1
            Top             =   540
            Width           =   18975
            _ExtentX        =   33470
            _ExtentY        =   10821
            _Version        =   393216
            Rows            =   30
            Cols            =   18
            FixedRows       =   0
            FixedCols       =   0
            ForeColor       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            SelectionMode   =   1
         End
         Begin MSFlexGridLib.MSFlexGrid grd_Listad_Cab 
            Height          =   615
            Left            =   30
            TabIndex        =   9
            Top             =   0
            Width           =   18705
            _ExtentX        =   32994
            _ExtentY        =   1085
            _Version        =   393216
            Rows            =   30
            Cols            =   18
            FixedRows       =   0
            FixedCols       =   0
            ForeColor       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   0
            SelectionMode   =   1
         End
      End
      Begin VB.Label lbl_NomEti 
         AutoSize        =   -1  'True
         Caption         =   "Columna a Buscar:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   8370
         Width           =   1335
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Buscar Por:"
         Height          =   195
         Left            =   4710
         TabIndex        =   16
         Top             =   8370
         Width           =   825
      End
   End
End
Attribute VB_Name = "frm_Ges_TecPro_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmb_Buscar_Click()
   If (cmb_Buscar.ListIndex = 0 Or cmb_Buscar.ListIndex = -1) Then
      txt_Buscar.Enabled = False
      Call gs_SetFocus(cmd_Buscar)
   Else
      txt_Buscar.Enabled = True
      Call gs_SetFocus(txt_Buscar)
   End If
   txt_Buscar.Text = ""
End Sub

Private Sub cmb_Buscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If (txt_Buscar.Enabled = False) Then
          Call gs_SetFocus(cmd_Buscar)
      Else
          Call gs_SetFocus(txt_Buscar)
      End If
   End If
End Sub

Private Sub cmd_Agrega_Click()
   moddat_g_int_FlgGrb = 1
   frm_Ges_TecPro_02.Show 1
End Sub

Private Sub cmd_Borrar_Click()
   
   If fs_Validar_MovEte = True Then
      MsgBox "Verifique que el registro no tenga Carta Fianzas, Adendas o Cartas Seriedad Oferta, registradas.", vbExclamation, modgen_g_str_NomPlt
      Call gs_SetFocus(grd_Listad)
      Exit Sub
   End If
   
   If MsgBox("¿Está seguro de eliminar los datos?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_str_NomPlt) <> vbYes Then
      Exit Sub
   End If
   
   moddat_g_int_FlgGOK = False
   moddat_g_int_CntErr = 0
   
   Do While moddat_g_int_FlgGOK = False
      Screen.MousePointer = 11
   
      'Grabando Información de Carta Fianza
      g_str_Parame = ""
      g_str_Parame = g_str_Parame & " USP_TPR_MAEETE_ELIMINA ("
      g_str_Parame = g_str_Parame & CStr(Trim(Mid(grd_Listad.TextMatrix(grd_Listad.Row, 0), 1, 2))) & ", "
      g_str_Parame = g_str_Parame & "'" & CStr(Trim(Mid(grd_Listad.TextMatrix(grd_Listad.Row, 0), 4))) & "') "
                  
      If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 2) Then
         moddat_g_int_CntErr = moddat_g_int_CntErr + 1
      Else
         moddat_g_int_FlgGOK = True
      End If
      
      If moddat_g_int_CntErr = 6 Then
         If MsgBox("No se pudo completar la eliminación de los datos. ¿Desea seguir intentando?", vbQuestion + vbYesNo + vbDefaultButton2, modgen_g_con_PltPar) <> vbYes Then
            Exit Sub
         Else
            moddat_g_int_CntErr = 0
         End If
      End If
      
      Screen.MousePointer = 0
   Loop
   
   'Actualiza la Grilla
   Call fs_Buscar
   
   If Me.grd_Listad.Rows = 0 Then
      Call fs_Activa(True)
   End If
End Sub
Private Function fs_Validar_MovEte() As Boolean
   fs_Validar_MovEte = False

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "  SELECT COUNT(*) AS CONTADOR "
   g_str_Parame = g_str_Parame & "    FROM TPR_MAECFI "
   g_str_Parame = g_str_Parame & "   WHERE MAECFI_TIPDOC = " & CStr(Trim(Mid(grd_Listad.TextMatrix(grd_Listad.Row, 0), 1, 2))) & " "
   g_str_Parame = g_str_Parame & "     AND MAECFI_NUMDOC = '" & CStr(Trim(Mid(grd_Listad.TextMatrix(grd_Listad.Row, 0), 4))) & "'"

   If Not gf_EjecutaSQL(g_str_Parame, g_rst_GenAux, 3) Then
     Exit Function
   End If
   
   If g_rst_GenAux.BOF And g_rst_GenAux.EOF Then
      g_rst_GenAux.Close
      Set g_rst_GenAux = Nothing
      Exit Function
   End If
     
   If Not (g_rst_GenAux.BOF And g_rst_GenAux.EOF) Then
      If (g_rst_GenAux!CONTADOR) > 0 Then
         fs_Validar_MovEte = True
      End If
   End If
End Function

Private Sub cmd_Buscar_Click()
   Screen.MousePointer = 11
   Call fs_Buscar
   Screen.MousePointer = 0
End Sub

Private Sub cmd_CarFia_Click()
Dim r_int_NumCFi  As Integer
Dim r_int_NumAD   As Integer
Dim r_int_NumCSO  As Integer
Dim r_int_NumLC   As Integer
Dim r_int_NumCP   As Integer

   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   grd_Listad.Col = 0
   moddat_g_int_TipDoc = Left(grd_Listad.Text, 1)
   moddat_g_str_NumDoc = Trim(Mid(grd_Listad.Text, 4))
         
   'RAZÓN SOCIAL - ET
   grd_Listad.Col = 1
   moddat_g_str_NomCli = grd_Listad.Text
   
   'LINEA ASIGNADA
'   grd_Listad.Col = 2
'   moddat_g_dbl_LinAsi = grd_Listad.Text
   
   'NÚMERO DE CARTAS FIANZAS
   grd_Listad.Col = 3
   r_int_NumCFi = grd_Listad.Text
   
   'NÚMERO DE ADENDAS
   grd_Listad.Col = 4
   r_int_NumAD = grd_Listad.Text
   
   'NÚMERO DE CARTAS DE SERIEDAD OFERTA
   grd_Listad.Col = 5
   r_int_NumCSO = grd_Listad.Text
   
   'NÚMERO DE CREDITOS INDIRECTOS
   grd_Listad.Col = 11
   r_int_NumLC = grd_Listad.Text
   
   'NÚMERO DE CREDITO PUNTUAL
   grd_Listad.Col = 12
   r_int_NumCP = grd_Listad.Text
   
   moddat_g_int_NumCuo = r_int_NumCFi + r_int_NumAD + r_int_NumCSO + r_int_NumLC + r_int_NumCP
   
   'TIPO DE EMPRESA
   grd_Listad.Col = 10
   moddat_g_str_Descri = moddat_gf_Consulta_ParDes("526", grd_Listad.Text)
   moddat_g_int_TipCli = grd_Listad.Text
   
   Call gs_RefrescaGrid(grd_Listad)
      
   moddat_g_int_FlgAct = 1
   
   frm_Ges_TecPro_03.Show 1
   
   If moddat_g_int_FlgAct = 2 Then
      Screen.MousePointer = 11
      Call fs_Buscar
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmd_Editar_Click()

   moddat_g_int_FlgGrb = 2    'Actualiza
   If fs_Validar = True Then
      frm_Ges_TecPro_02.Show 1
   End If
End Sub
Private Function fs_Validar() As Boolean
   fs_Validar = False
   
   If grd_Listad.Row >= 0 Then
      moddat_g_int_TipDoc = Left(grd_Listad.TextMatrix(grd_Listad.Row, 0), 1)
      moddat_g_str_NumDoc = Trim(Mid(grd_Listad.TextMatrix(grd_Listad.Row, 0), 4))
      moddat_g_str_NomCli = Trim(grd_Listad.TextMatrix(grd_Listad.Row, 1))
      moddat_g_str_Descri = moddat_gf_Consulta_ParDes("526", grd_Listad.TextMatrix(grd_Listad.Row, 10))
   End If

   fs_Validar = True
End Function


Private Sub cmd_Evalua_Click()
   moddat_g_int_FlgGrb_1 = 5
   If fs_Validar = True Then
      frm_Ges_TecPro_15.Show 1
   End If
End Sub

Private Sub cmd_ExpExcRes_Click()
   frm_Ges_TecPro_13.Show 1
End Sub

Private Sub cmd_Limpia_Click()
   Call fs_Limpia
End Sub

Private Sub cmd_ResEte_Click()
   If fs_Validar = True Then
      frm_Ges_TecPro_16.Show 1
   End If
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Obtiene_Cabecera
   Call fs_Buscar
   Call fs_Activa(True)
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   
   Call gs_LimpiaGrid(grd_Listad)
   Call gs_LimpiaGrid(grd_Listad_Cab)
   
   'Búsqueda
   cmb_Buscar.Clear
   cmb_Buscar.AddItem "NINGUNA"
   cmb_Buscar.AddItem "DOCUMENTO"
   cmb_Buscar.AddItem "RAZON SOCIAL"
   cmb_Buscar.ListIndex = 0
   
   'Inicializando Rejilla
   grd_Listad.ColWidth(0) = 1510       'Tipo Doc
   grd_Listad.ColWidth(1) = 5450       'Razon Social
   grd_Listad.ColWidth(2) = 1450       'Linea Asignada
   grd_Listad.ColWidth(3) = 670        'Nro CF
   grd_Listad.ColWidth(4) = 670        'Nro AD
   grd_Listad.ColWidth(5) = 670        'Nro CSO
   grd_Listad.ColWidth(6) = 1370       'Garantia
   grd_Listad.ColWidth(7) = 1380       'Linea Util CF
   grd_Listad.ColWidth(8) = 1380       'Linea Util AD
   grd_Listad.ColWidth(9) = 1380       'Linea Util CSO
   grd_Listad.ColWidth(10) = 0         'Tipo Empresa
   grd_Listad.ColWidth(11) = 670       'Nro LC
   grd_Listad.ColWidth(12) = 670       'Nro CP
   grd_Listad.ColWidth(13) = 1370      'Garantia
   grd_Listad.ColWidth(14) = 1400      'Línea Útil Línea de Crédito
   grd_Listad.ColWidth(15) = 1400      'Línea Útil Crédito Puntual
   grd_Listad.ColWidth(16) = 1400      'Saldo
   grd_Listad.ColWidth(17) = 1600      'Usuario
   
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
   grd_Listad.ColAlignment(2) = flexAlignRightCenter
   grd_Listad.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad.ColAlignment(4) = flexAlignCenterCenter
   grd_Listad.ColAlignment(5) = flexAlignCenterCenter
   grd_Listad.ColAlignment(6) = flexAlignRightCenter
   grd_Listad.ColAlignment(7) = flexAlignRightCenter
   grd_Listad.ColAlignment(8) = flexAlignRightCenter
   grd_Listad.ColAlignment(9) = flexAlignRightCenter
   grd_Listad.ColAlignment(11) = flexAlignCenterCenter
   grd_Listad.ColAlignment(12) = flexAlignCenterCenter
   grd_Listad.ColAlignment(13) = flexAlignRightCenter
   grd_Listad.ColAlignment(14) = flexAlignRightCenter
   grd_Listad.ColAlignment(15) = flexAlignRightCenter
   grd_Listad.ColAlignment(16) = flexAlignRightCenter
   grd_Listad.ColAlignment(17) = flexAlignCenterCenter
   
   
   'Inicializando Rejilla Cabecera
   grd_Listad_Cab.ColWidth(0) = 1510       'Tipo Doc
   grd_Listad_Cab.ColWidth(1) = 5450       'Razon Social
   grd_Listad_Cab.ColWidth(2) = 1450       'Linea Asignada
   grd_Listad_Cab.ColWidth(3) = 670        'Nro CF
   grd_Listad_Cab.ColWidth(4) = 670        'Nro AD
   grd_Listad_Cab.ColWidth(5) = 670        'Nro CSO
   grd_Listad_Cab.ColWidth(6) = 1370       'Garantia
   grd_Listad_Cab.ColWidth(7) = 1380       'Linea Util CF
   grd_Listad_Cab.ColWidth(8) = 1380       'Linea Util AD
   grd_Listad_Cab.ColWidth(9) = 1380       'Linea Util CSO
   grd_Listad_Cab.ColWidth(10) = 0         'Tipo Empresa
   grd_Listad_Cab.ColWidth(11) = 670       'Nro LC
   grd_Listad_Cab.ColWidth(12) = 670       'Nro CP
   grd_Listad_Cab.ColWidth(13) = 1370      'Garantia
   grd_Listad_Cab.ColWidth(14) = 1400      'Línea Útil Línea de Crédito
   grd_Listad_Cab.ColWidth(15) = 1400      'Línea Útil Crédito Puntual
   grd_Listad_Cab.ColWidth(16) = 1400      'Saldo
   grd_Listad_Cab.ColWidth(17) = 1600      'Usuario
   
   grd_Listad_Cab.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad_Cab.ColAlignment(1) = flexAlignLeftCenter
   grd_Listad_Cab.ColAlignment(2) = flexAlignRightCenter
   grd_Listad_Cab.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad_Cab.ColAlignment(4) = flexAlignCenterCenter
   grd_Listad_Cab.ColAlignment(5) = flexAlignCenterCenter
   grd_Listad_Cab.ColAlignment(6) = flexAlignRightCenter
   grd_Listad_Cab.ColAlignment(7) = flexAlignRightCenter
   grd_Listad_Cab.ColAlignment(8) = flexAlignRightCenter
   grd_Listad_Cab.ColAlignment(9) = flexAlignRightCenter
   grd_Listad_Cab.ColAlignment(11) = flexAlignCenterCenter
   grd_Listad_Cab.ColAlignment(12) = flexAlignCenterCenter
   grd_Listad_Cab.ColAlignment(13) = flexAlignRightCenter
   grd_Listad_Cab.ColAlignment(14) = flexAlignRightCenter
   grd_Listad_Cab.ColAlignment(15) = flexAlignRightCenter
   grd_Listad_Cab.ColAlignment(16) = flexAlignRightCenter
   grd_Listad_Cab.ColAlignment(17) = flexAlignCenterCenter
End Sub
Public Sub fs_Limpia()
   Call gs_LimpiaGrid(grd_Listad)
   cmb_Buscar.ListIndex = 0
   txt_Buscar.Text = Empty
End Sub
Private Sub fs_Obtiene_Cabecera()
Dim r_int_ConFil As Integer
Dim r_int_ConCol As Integer
   
   grd_Listad_Cab.Redraw = False
'   Call gs_LimpiaGrid(grd_Listad_Cab)
   
   'Primera Linea
   grd_Listad_Cab.Rows = grd_Listad_Cab.Rows + 1
   grd_Listad_Cab.Row = grd_Listad_Cab.Rows - 1
   grd_Listad_Cab.Row = 0:   grd_Listad_Cab.Text = ""
   grd_Listad_Cab.Col = 0:   grd_Listad_Cab.Text = "TIPO - NRO.DOC":            grd_Listad_Cab.CellAlignment = flexAlignCenterCenter
   grd_Listad_Cab.Col = 1:   grd_Listad_Cab.Text = "RAZON SOCIAL":              grd_Listad_Cab.CellAlignment = flexAlignCenterCenter
   grd_Listad_Cab.Col = 2:   grd_Listad_Cab.Text = "LINEA ASIGNADA":            grd_Listad_Cab.CellAlignment = flexAlignCenterCenter
   grd_Listad_Cab.Col = 3:   grd_Listad_Cab.Text = "CREDICTOS INDIRECTOS":      grd_Listad_Cab.CellAlignment = flexAlignCenterCenter
   grd_Listad_Cab.Col = 4:   grd_Listad_Cab.Text = "CREDICTOS INDIRECTOS":      grd_Listad_Cab.CellAlignment = flexAlignCenterCenter
   grd_Listad_Cab.Col = 5:   grd_Listad_Cab.Text = "CREDICTOS INDIRECTOS":      grd_Listad_Cab.CellAlignment = flexAlignCenterCenter
   grd_Listad_Cab.Col = 6:   grd_Listad_Cab.Text = "CREDICTOS INDIRECTOS":      grd_Listad_Cab.CellAlignment = flexAlignCenterCenter
   grd_Listad_Cab.Col = 7:   grd_Listad_Cab.Text = "CREDICTOS INDIRECTOS":      grd_Listad_Cab.CellAlignment = flexAlignCenterCenter
   grd_Listad_Cab.Col = 8:   grd_Listad_Cab.Text = "CREDICTOS INDIRECTOS":      grd_Listad_Cab.CellAlignment = flexAlignCenterCenter
   grd_Listad_Cab.Col = 9:   grd_Listad_Cab.Text = "CREDICTOS INDIRECTOS":      grd_Listad_Cab.CellAlignment = flexAlignCenterCenter
   grd_Listad_Cab.Col = 11:  grd_Listad_Cab.Text = "CREDICTOS DIRECTOS":        grd_Listad_Cab.CellAlignment = flexAlignCenterCenter
   grd_Listad_Cab.Col = 12:  grd_Listad_Cab.Text = "CREDICTOS DIRECTOS":        grd_Listad_Cab.CellAlignment = flexAlignCenterCenter
   grd_Listad_Cab.Col = 13:  grd_Listad_Cab.Text = "CREDICTOS DIRECTOS":        grd_Listad_Cab.CellAlignment = flexAlignCenterCenter
   grd_Listad_Cab.Col = 14:  grd_Listad_Cab.Text = "CREDICTOS DIRECTOS":        grd_Listad_Cab.CellAlignment = flexAlignCenterCenter
   grd_Listad_Cab.Col = 15:  grd_Listad_Cab.Text = "CREDICTOS DIRECTOS":        grd_Listad_Cab.CellAlignment = flexAlignCenterCenter
   grd_Listad_Cab.Col = 16:  grd_Listad_Cab.Text = "SALDO":                     grd_Listad_Cab.CellAlignment = flexAlignCenterCenter
   grd_Listad_Cab.Col = 17:  grd_Listad_Cab.Text = "USUARIO":                   grd_Listad_Cab.CellAlignment = flexAlignCenterCenter

   'Segunda linea
   grd_Listad_Cab.Rows = grd_Listad_Cab.Rows + 1
   grd_Listad_Cab.Row = grd_Listad_Cab.Rows - 1
   grd_Listad_Cab.Col = 0:   grd_Listad_Cab.Text = "TIPO - NRO.DOC":            grd_Listad_Cab.CellAlignment = flexAlignCenterCenter
   grd_Listad_Cab.Col = 1:   grd_Listad_Cab.Text = "RAZON SOCIAL":              grd_Listad_Cab.CellAlignment = flexAlignCenterCenter
   grd_Listad_Cab.Col = 2:   grd_Listad_Cab.Text = "LINEA ASIGNADA":            grd_Listad_Cab.CellAlignment = flexAlignCenterCenter
   grd_Listad_Cab.Col = 3:   grd_Listad_Cab.Text = "N° CF":                     grd_Listad_Cab.CellAlignment = flexAlignCenterCenter
   grd_Listad_Cab.Col = 4:   grd_Listad_Cab.Text = "N° AD":                     grd_Listad_Cab.CellAlignment = flexAlignCenterCenter
   grd_Listad_Cab.Col = 5:   grd_Listad_Cab.Text = "N° CSO":                    grd_Listad_Cab.CellAlignment = flexAlignCenterCenter
   grd_Listad_Cab.Col = 6:   grd_Listad_Cab.Text = "GARANTIA":                  grd_Listad_Cab.CellAlignment = flexAlignCenterCenter
   grd_Listad_Cab.Col = 7:   grd_Listad_Cab.Text = "LINEA UTIL.CF":             grd_Listad_Cab.CellAlignment = flexAlignCenterCenter
   grd_Listad_Cab.Col = 8:   grd_Listad_Cab.Text = "LINEA UTIL.AD":             grd_Listad_Cab.CellAlignment = flexAlignCenterCenter
   grd_Listad_Cab.Col = 9:   grd_Listad_Cab.Text = "LINEA UTIL.CSO":            grd_Listad_Cab.CellAlignment = flexAlignCenterCenter
   grd_Listad_Cab.Col = 11:  grd_Listad_Cab.Text = "N° LC":                     grd_Listad_Cab.CellAlignment = flexAlignCenterCenter
   grd_Listad_Cab.Col = 12:  grd_Listad_Cab.Text = "N° CP":                     grd_Listad_Cab.CellAlignment = flexAlignCenterCenter
   grd_Listad_Cab.Col = 13:  grd_Listad_Cab.Text = "GARANTIA":                  grd_Listad_Cab.CellAlignment = flexAlignCenterCenter
   grd_Listad_Cab.Col = 14:  grd_Listad_Cab.Text = "LINEA UTIL.LC":             grd_Listad_Cab.CellAlignment = flexAlignCenterCenter
   grd_Listad_Cab.Col = 15:  grd_Listad_Cab.Text = "LINEA UTIL.CP":             grd_Listad_Cab.CellAlignment = flexAlignCenterCenter
   grd_Listad_Cab.Col = 16:  grd_Listad_Cab.Text = "SALDO":                     grd_Listad_Cab.CellAlignment = flexAlignCenterCenter
   grd_Listad_Cab.Col = 17:  grd_Listad_Cab.Text = "USUARIO":                   grd_Listad_Cab.CellAlignment = flexAlignCenterCenter

   grd_Listad_Cab.Rows = grd_Listad_Cab.Rows + 1
   grd_Listad_Cab.Row = grd_Listad_Cab.Rows - 1
   
   With grd_Listad_Cab
      .MergeCells = flexMergeFree
      .MergeRow(0) = True
      .MergeCol(0) = True
      .MergeCol(1) = True
      .MergeCol(2) = True
      .MergeCol(16) = True
      .MergeCol(17) = True
      .FixedRows = 2
      .FixedCols = 2
   End With
   
   With grd_Listad
'      .FixedRows = 2
      .FixedCols = 2
   End With

   grd_Listad_Cab.Rows = grd_Listad_Cab.Rows - 1

   For r_int_ConFil = 0 To grd_Listad_Cab.Rows - 1
      For r_int_ConCol = 0 To grd_Listad_Cab.Cols - 1
         grd_Listad_Cab.Col = r_int_ConCol
         grd_Listad_Cab.Row = r_int_ConFil
         grd_Listad_Cab.CellBackColor = &H4000&
         grd_Listad_Cab.ForeColorFixed = &HFFFFFF
      Next r_int_ConCol
   Next r_int_ConFil
   
   grd_Listad_Cab.Redraw = True
End Sub
Public Sub fs_Buscar()

   g_str_Parame = ""
   g_str_Parame = g_str_Parame & " SELECT (MAEETE_TIPDOC || ' - ' || MAEETE_NUMDOC) AS DOCUMENTO, MAEPRV_RAZSOC AS RAZON_SOCIAL, "
   g_str_Parame = g_str_Parame & "        SUM(NVL(MAEETE_LINASI_IND,0) + NVL(MAEETE_LINASI_DIR,0)) LINEA_ASIGNADA, MAEETE_TIPEMP, TPR_MAEETE.SEGUSUCRE, "
   
   g_str_Parame = g_str_Parame & "        NVL(( SELECT COUNT(MAECFI_NUMREF)  "
   g_str_Parame = g_str_Parame & "                FROM TPR_MAECFI "
   g_str_Parame = g_str_Parame & "               WHERE MAECFI_TIPDOC = MAEETE_TIPDOC AND MAECFI_NUMDOC = MAEETE_NUMDOC AND MAECFI_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                 AND MAECFI_CODPRD IN ('026','027') AND MAECFI_CODMOD <> '005' AND MAECFI_CODMOD <> '008' "
   g_str_Parame = g_str_Parame & "               GROUP BY MAECFI_TIPDOC, MAECFI_NUMDOC ),0) AS NRO_CARTA_CF, "
   
   g_str_Parame = g_str_Parame & "        NVL(( SELECT COUNT(MAECFI_NUMREF)  "
   g_str_Parame = g_str_Parame & "                FROM TPR_MAECFI "
   g_str_Parame = g_str_Parame & "               WHERE MAECFI_TIPDOC = MAEETE_TIPDOC AND MAECFI_NUMDOC = MAEETE_NUMDOC AND MAECFI_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                 AND MAECFI_CODPRD = '026' AND MAECFI_CODMOD = '005'"
   g_str_Parame = g_str_Parame & "               GROUP BY MAECFI_TIPDOC, MAECFI_NUMDOC ),0) AS NRO_CARTA_AD, "
   
   g_str_Parame = g_str_Parame & "        NVL(( SELECT COUNT(MAECFI_NUMREF)  "
   g_str_Parame = g_str_Parame & "                FROM TPR_MAECFI "
   g_str_Parame = g_str_Parame & "               WHERE MAECFI_TIPDOC = MAEETE_TIPDOC AND MAECFI_NUMDOC = MAEETE_NUMDOC AND MAECFI_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                 AND MAECFI_CODPRD IN ('026','027') AND MAECFI_CODMOD = '008'"
   g_str_Parame = g_str_Parame & "               GROUP BY MAECFI_TIPDOC, MAECFI_NUMDOC ),0) AS NRO_CARTA_CSO, "
   
   g_str_Parame = g_str_Parame & "        NVL(( SELECT NVL(SUM(MAECFI_GARFIA),0) "
   g_str_Parame = g_str_Parame & "                FROM TPR_MAECFI "
   g_str_Parame = g_str_Parame & "               WHERE MAECFI_TIPDOC = MAEETE_TIPDOC AND MAECFI_NUMDOC = MAEETE_NUMDOC AND MAECFI_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                 AND MAECFI_CODPRD IN ('026','027') AND MAECFI_CODMOD <> '005' AND MAECFI_CODMOD <> '008' "
   g_str_Parame = g_str_Parame & "               GROUP BY MAECFI_TIPDOC, MAECFI_NUMDOC),0) AS LINEA_UTILIZADA_CF, "
   
   g_str_Parame = g_str_Parame & "        NVL(( SELECT NVL(SUM(MAECFI_GARFIA),0) "
   g_str_Parame = g_str_Parame & "                FROM TPR_MAECFI "
   g_str_Parame = g_str_Parame & "               WHERE MAECFI_TIPDOC = MAEETE_TIPDOC AND MAECFI_NUMDOC = MAEETE_NUMDOC AND MAECFI_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                 AND MAECFI_CODPRD = '026' AND MAECFI_CODMOD = '005'"
   g_str_Parame = g_str_Parame & "               GROUP BY MAECFI_TIPDOC, MAECFI_NUMDOC),0) AS LINEA_UTILIZADA_AD, "
   
   g_str_Parame = g_str_Parame & "        NVL(( SELECT NVL(SUM(MAECFI_GARFIA),0) "
   g_str_Parame = g_str_Parame & "                FROM TPR_MAECFI "
   g_str_Parame = g_str_Parame & "               WHERE MAECFI_TIPDOC = MAEETE_TIPDOC AND MAECFI_NUMDOC = MAEETE_NUMDOC AND MAECFI_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                 AND MAECFI_CODPRD IN ('026','027') AND MAECFI_CODMOD = '008'"
   g_str_Parame = g_str_Parame & "               GROUP BY MAECFI_TIPDOC, MAECFI_NUMDOC),0) AS LINEA_UTILIZADA_CSO, "
   
   g_str_Parame = g_str_Parame & "        NVL(( SELECT NVL(SUM(MAECFI_GARFIA),0) "
   g_str_Parame = g_str_Parame & "                FROM TPR_MAECFI "
   g_str_Parame = g_str_Parame & "               WHERE MAECFI_TIPDOC = MAEETE_TIPDOC AND MAECFI_NUMDOC = MAEETE_NUMDOC AND MAECFI_CODPRD IN ('026','027') AND MAECFI_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "               GROUP BY MAECFI_TIPDOC, MAECFI_NUMDOC),0) AS LINEA_UTILIZADA_IND, "
   
   g_str_Parame = g_str_Parame & "        NVL(( SELECT NVL(SUM(NVL(MAEGAR_MTOGAR_INM,0) + NVL(MAEGAR_MTOGAR_ES1,0) + NVL(MAEGAR_MTOGAR_ES2,0) + NVL(MAEGAR_MTOGAR_DE1,0) + NVL(MAEGAR_MTOGAR_DE2,0)),0)"
   g_str_Parame = g_str_Parame & "                FROM TPR_MAEGAR INNER JOIN TPR_MAECFI ON MAECFI_TIPDOC = MAEGAR_TIPDOC AND MAECFI_NUMDOC = MAEGAR_NUMDOC AND TRIM(MAECFI_NUMREF) = TRIM (MAEGAR_NUMREF) AND MAECFI_CODPRD IN ('026','027')"
   g_str_Parame = g_str_Parame & "               WHERE MAEGAR_TIPDOC = MAEETE_TIPDOC And MAEGAR_NUMDOC = MAEETE_NUMDOC And MAEGAR_SITUAC = 1 And MAEGAR_TIPGAR = 1"
   g_str_Parame = g_str_Parame & "               GROUP BY MAEGAR_TIPDOC, MAEGAR_NUMDOC),0) AS GARANTIA_LIQUIDA_IND, "
                     
   g_str_Parame = g_str_Parame & "        NVL(( SELECT NVL(SUM(NVL(MAEGAR_MTOGAR_INM,0) + NVL(MAEGAR_MTOGAR_ES1,0) + NVL(MAEGAR_MTOGAR_ES2,0) + NVL(MAEGAR_MTOGAR_DE1,0) + NVL(MAEGAR_MTOGAR_DE2,0)),0)"
   g_str_Parame = g_str_Parame & "                FROM TPR_MAEGAR"
   g_str_Parame = g_str_Parame & "               WHERE MAEGAR_TIPDOC = MAEETE_TIPDOC And MAEGAR_NUMDOC = MAEETE_NUMDOC And MAEGAR_SITUAC = 1 And MAEGAR_TIPGAR = 2"
   g_str_Parame = g_str_Parame & "               GROUP BY MAEGAR_TIPDOC, MAEGAR_NUMDOC),0) AS GARANTIA_HIPOTECARIO_IND, "
   
'   g_str_Parame = g_str_Parame & "        NVL(( SELECT NVL(SUM(NVL(MAEGAR_MTOGAR_INM,0) + NVL(MAEGAR_MTOGAR_ES1,0) + NVL(MAEGAR_MTOGAR_ES2,0) + NVL(MAEGAR_MTOGAR_DE1,0) + NVL(MAEGAR_MTOGAR_DE2,0)),0) "
'   g_str_Parame = g_str_Parame & "                FROM TPR_MAEGAR "
'   g_str_Parame = g_str_Parame & "               WHERE MAEGAR_TIPDOC = MAEETE_TIPDOC AND MAEGAR_NUMDOC = MAEETE_NUMDOC AND MAEGAR_SITUAC = 1 "
'   g_str_Parame = g_str_Parame & "               GROUP BY MAEGAR_TIPDOC, MAEGAR_NUMDOC),0) AS GARANTIA_CRED_IND "
   
   g_str_Parame = g_str_Parame & "        NVL(( SELECT COUNT(MAECFI_NUMREF)  "
   g_str_Parame = g_str_Parame & "                FROM TPR_MAECFI "
   g_str_Parame = g_str_Parame & "               WHERE MAECFI_TIPDOC = MAEETE_TIPDOC AND MAECFI_NUMDOC = MAEETE_NUMDOC AND MAECFI_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                 AND MAECFI_CODPRD = '008' AND MAECFI_CODSUB = '008' AND MAECFI_CODMOD = '001' "
   g_str_Parame = g_str_Parame & "               GROUP BY MAECFI_TIPDOC, MAECFI_NUMDOC ),0) AS NRO_DIR_LIN_CREDITO, "
   
   g_str_Parame = g_str_Parame & "        NVL(( SELECT COUNT(MAECFI_NUMREF)  "
   g_str_Parame = g_str_Parame & "                FROM TPR_MAECFI "
   g_str_Parame = g_str_Parame & "               WHERE MAECFI_TIPDOC = MAEETE_TIPDOC AND MAECFI_NUMDOC = MAEETE_NUMDOC AND MAECFI_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                 AND MAECFI_CODPRD = '008' AND MAECFI_CODSUB = '008' AND MAECFI_CODMOD = '002' "
   g_str_Parame = g_str_Parame & "               GROUP BY MAECFI_TIPDOC, MAECFI_NUMDOC ),0) AS NRO_DIR_CRED_PUNTUAL, "
   
   g_str_Parame = g_str_Parame & "        NVL(( SELECT NVL(SUM(MAECFI_IMPFIA),0) "
   g_str_Parame = g_str_Parame & "                FROM TPR_MAECFI "
   g_str_Parame = g_str_Parame & "               WHERE MAECFI_TIPDOC = MAEETE_TIPDOC AND MAECFI_NUMDOC = MAEETE_NUMDOC AND MAECFI_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                 AND MAECFI_CODPRD = '008' AND MAECFI_CODSUB = '008' AND MAECFI_CODMOD = '001' "
   g_str_Parame = g_str_Parame & "               GROUP BY MAECFI_TIPDOC, MAECFI_NUMDOC),0) AS LINEA_UTIL_DIR_LIN_CREDITO, "
   
   g_str_Parame = g_str_Parame & "        NVL(( SELECT NVL(SUM(MAECFI_IMPFIA),0) "
   g_str_Parame = g_str_Parame & "                FROM TPR_MAECFI "
   g_str_Parame = g_str_Parame & "               WHERE MAECFI_TIPDOC = MAEETE_TIPDOC AND MAECFI_NUMDOC = MAEETE_NUMDOC AND MAECFI_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "                 AND MAECFI_CODPRD = '008' AND MAECFI_CODSUB = '008' AND MAECFI_CODMOD = '002' "
   g_str_Parame = g_str_Parame & "               GROUP BY MAECFI_TIPDOC, MAECFI_NUMDOC),0) AS LINEA_UTIL_DIR_CRE_PUNTUAL, "
     
   g_str_Parame = g_str_Parame & "        NVL(( SELECT NVL(SUM(MAECFI_IMPFIA),0) "
   g_str_Parame = g_str_Parame & "                FROM TPR_MAECFI "
   g_str_Parame = g_str_Parame & "               WHERE MAECFI_TIPDOC = MAEETE_TIPDOC AND MAECFI_NUMDOC = MAEETE_NUMDOC AND MAECFI_CODPRD = '008' AND MAECFI_SITUAC = 1 "
   g_str_Parame = g_str_Parame & "               GROUP BY MAECFI_TIPDOC, MAECFI_NUMDOC),0) AS LINEA_UTILIZADA_DIR, "
   
   g_str_Parame = g_str_Parame & "        NVL(( SELECT NVL(SUM(NVL(MAEGAR_MTOGAR_INM,0) + NVL(MAEGAR_MTOGAR_ES1,0) + NVL(MAEGAR_MTOGAR_ES2,0) + NVL(MAEGAR_MTOGAR_DE1,0) + NVL(MAEGAR_MTOGAR_DE2,0)),0)"
   g_str_Parame = g_str_Parame & "                FROM TPR_MAEGAR INNER JOIN TPR_MAECFI ON MAECFI_TIPDOC = MAEGAR_TIPDOC AND MAECFI_NUMDOC = MAEGAR_NUMDOC AND TRIM(MAECFI_NUMREF) = TRIM (MAEGAR_NUMREF) AND MAECFI_CODPRD IN ('008')"
   g_str_Parame = g_str_Parame & "               WHERE MAEGAR_TIPDOC = MAEETE_TIPDOC And MAEGAR_NUMDOC = MAEETE_NUMDOC And MAEGAR_SITUAC = 1 And MAEGAR_TIPGAR = 1"
   g_str_Parame = g_str_Parame & "               GROUP BY MAEGAR_TIPDOC, MAEGAR_NUMDOC),0) AS GARANTIA_LIQUIDA_DIR "
     
   g_str_Parame = g_str_Parame & "   FROM TPR_MAEETE "
   g_str_Parame = g_str_Parame & "        INNER JOIN CNTBL_MAEPRV ON MAEETE_TIPDOC = MAEPRV_TIPDOC AND MAEETE_NUMDOC = MAEPRV_NUMDOC "
   
   g_str_Parame = g_str_Parame & "  WHERE MAEETE_TIPDOC > 0 "
   g_str_Parame = g_str_Parame & "    AND MAEETE_NUMDOC IS NOT NULL "
   
   If cmb_Buscar.ListIndex > 0 Then
      If cmb_Buscar.ListIndex = 1 Then    'NRO DOCUMENTO
         g_str_Parame = g_str_Parame & "   AND MAEETE_NUMDOC = '" & Trim(txt_Buscar.Text) & "' "
      ElseIf cmb_Buscar.ListIndex = 2 Then 'RAZON SOCIAL
         g_str_Parame = g_str_Parame & "   AND TRIM(MAEPRV_RAZSOC) like '%" & UCase(Trim(txt_Buscar.Text)) & "%' "
      End If
   End If
   
   g_str_Parame = g_str_Parame & "  GROUP BY MAEETE_TIPDOC, MAEETE_NUMDOC, MAEPRV_RAZSOC, MAEETE_TIPEMP, TPR_MAEETE.SEGUSUCRE "
   g_str_Parame = g_str_Parame & "  ORDER BY MAEPRV_RAZSOC ASC "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   grd_Listad.Redraw = False
   
   Call gs_LimpiaGrid(grd_Listad)
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
          
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
         
         grd_Listad.Col = 0
         grd_Listad.Text = CStr(g_rst_Princi!DOCUMENTO)
         grd_Listad.CellForeColor = &H80000012
         
         grd_Listad.Col = 1
         grd_Listad.Text = CStr(g_rst_Princi!RAZON_SOCIAL)
         grd_Listad.CellForeColor = &H80000012
                 
         grd_Listad.Col = 2
         grd_Listad.Text = Format(CStr(g_rst_Princi!LINEA_ASIGNADA), "###,###,###,##0.00")
         
         grd_Listad.Col = 3
         grd_Listad.Text = CStr(g_rst_Princi!NRO_CARTA_CF)
         
         grd_Listad.Col = 4
         grd_Listad.Text = CStr(g_rst_Princi!NRO_CARTA_AD)
         
         grd_Listad.Col = 5
         grd_Listad.Text = CStr(g_rst_Princi!NRO_CARTA_CSO)
         
         grd_Listad.Col = 6
         grd_Listad.Text = Format(CStr(g_rst_Princi!GARANTIA_LIQUIDA_IND + g_rst_Princi!GARANTIA_HIPOTECARIO_IND), "###,###,###,##0.00")
         
         grd_Listad.Col = 7
         grd_Listad.Text = Format(CStr(g_rst_Princi!LINEA_UTILIZADA_CF), "###,###,###,##0.00")
         
         grd_Listad.Col = 8
         grd_Listad.Text = Format(CStr(g_rst_Princi!LINEA_UTILIZADA_AD), "###,###,###,##0.00")
         
         grd_Listad.Col = 9
         grd_Listad.Text = Format(CStr(g_rst_Princi!LINEA_UTILIZADA_CSO), "###,###,###,##0.00")
         
         grd_Listad.Col = 10
         grd_Listad.Text = CStr(g_rst_Princi!MAEETE_TIPEMP)
         
         grd_Listad.Col = 11
         grd_Listad.Text = CStr(g_rst_Princi!NRO_DIR_LIN_CREDITO)
         
         grd_Listad.Col = 12
         grd_Listad.Text = CStr(g_rst_Princi!NRO_DIR_CRED_PUNTUAL)
         
         grd_Listad.Col = 13
         grd_Listad.Text = Format(CStr(g_rst_Princi!GARANTIA_LIQUIDA_DIR), "###,###,###,##0.00")
                        
         grd_Listad.Col = 14
         grd_Listad.Text = Format(CStr(g_rst_Princi!LINEA_UTIL_DIR_LIN_CREDITO), "###,###,###,##0.00")
         
         grd_Listad.Col = 15
         grd_Listad.Text = Format(CStr(g_rst_Princi!LINEA_UTIL_DIR_CRE_PUNTUAL), "###,###,###,##0.00")
         
         grd_Listad.Col = 16
         grd_Listad.Text = Format(CDbl(g_rst_Princi!LINEA_ASIGNADA) - CDbl(g_rst_Princi!LINEA_UTILIZADA_IND) - CDbl(g_rst_Princi!LINEA_UTILIZADA_DIR) + CDbl(g_rst_Princi!LINEA_UTIL_DIR_CRE_PUNTUAL), "###,###,###,##0.00")

         grd_Listad.Col = 17
         grd_Listad.Text = Trim(g_rst_Princi!SEGUSUCRE & "")
         
         g_rst_Princi.MoveNext
      Loop
   End If
     
   grd_Listad.Redraw = True
      
   If grd_Listad.Rows = 0 Then
      Call fs_Activa(True)
            
      MsgBox "No se encontraron Entidades Técnicas.", vbInformation, modgen_g_str_NomPlt ' con Cartas Fianzas
   Else
      'Ordenando por Nombre de Cliente
'      pnl_RazSoc.Tag = "A"
'      Call gs_SorteaGrid(grd_Listad, 1, "C")
   
      Call gs_UbiIniGrid(grd_Listad)
      Call fs_Activa(True)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
End Sub
Public Sub fs_Activa(ByVal p_Activa As Integer)
   cmd_Agrega.Enabled = p_Activa
   
   If grd_Listad.Rows = 0 Then
      cmd_Editar.Enabled = Not p_Activa
      cmd_Borrar.Enabled = Not p_Activa
      cmd_CarFia.Enabled = Not p_Activa
      cmd_ExpExcRes.Enabled = Not p_Activa
      cmd_NueGar.Enabled = Not p_Activa
      cmd_Evalua.Enabled = Not p_Activa
      cmd_ResEte.Enabled = Not p_Activa
   Else
      cmd_Editar.Enabled = p_Activa
      cmd_Borrar.Enabled = p_Activa
      cmd_CarFia.Enabled = p_Activa
      cmd_ExpExcRes.Enabled = p_Activa
      cmd_NueGar.Enabled = p_Activa
      cmd_Evalua.Enabled = p_Activa
      cmd_ResEte.Enabled = p_Activa
   End If
End Sub

Private Sub grd_Listad_DblClick()
   moddat_g_int_FlgAct_2 = 0
   Call cmd_CarFia_Click
End Sub

Private Sub grd_Listad_Scroll()
   grd_Listad_Cab.Height = 615
   grd_Listad_Cab.ScrollBars = flexScrollBarHorizontal
   grd_Listad_Cab.LeftCol = grd_Listad.LeftCol
   grd_Listad_Cab.ScrollBars = flexScrollBarNone
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

'Private Sub pnl_DocIde_Click()
'   If Len(Trim(pnl_DocIde.Tag)) = 0 Or pnl_DocIde.Tag = "D" Then
'      pnl_DocIde.Tag = "A"
'      Call gs_SorteaGrid(grd_Listad, 0, "C")
'   Else
'      pnl_DocIde.Tag = "D"
'      Call gs_SorteaGrid(grd_Listad, 0, "C-")
'   End If
'End Sub
'
'Private Sub pnl_LinAsg_Click()
'   If Len(Trim(pnl_LinAsg.Tag)) = 0 Or pnl_LinAsg.Tag = "D" Then
'      pnl_LinAsg.Tag = "A"
'      Call gs_SorteaGrid(grd_Listad, 2, "N")
'   Else
'      pnl_LinAsg.Tag = "D"
'      Call gs_SorteaGrid(grd_Listad, 2, "N-")
'   End If
'End Sub
'
'Private Sub pnl_LinUtiCF_Click()
'   If Len(Trim(pnl_LinUtiCF.Tag)) = 0 Or pnl_LinUtiCF.Tag = "D" Then
'      pnl_LinUtiCF.Tag = "A"
'      Call gs_SorteaGrid(grd_Listad, 7, "N")
'   Else
'      pnl_LinUtiCF.Tag = "D"
'      Call gs_SorteaGrid(grd_Listad, 7, "N-")
'   End If
'End Sub
'
'Private Sub pnl_LinUtilAD_Click()
'   If Len(Trim(pnl_LinUtilAD.Tag)) = 0 Or pnl_LinUtilAD.Tag = "D" Then
'      pnl_LinUtilAD.Tag = "A"
'      Call gs_SorteaGrid(grd_Listad, 8, "N")
'   Else
'      pnl_LinUtilAD.Tag = "D"
'      Call gs_SorteaGrid(grd_Listad, 8, "N-")
'   End If
'End Sub
'Private Sub pnl_LinUtilCSO_Click()
'   If Len(Trim(pnl_LinUtilCSO.Tag)) = 0 Or pnl_LinUtilCSO.Tag = "D" Then
'      pnl_LinUtilCSO.Tag = "A"
'      Call gs_SorteaGrid(grd_Listad, 9, "N")
'   Else
'      pnl_LinUtilCSO.Tag = "D"
'      Call gs_SorteaGrid(grd_Listad, 9, "N-")
'   End If
'End Sub
'Private Sub pnl_MToGar_Click()
'   If Len(Trim(pnl_MToGar.Tag)) = 0 Or pnl_MToGar.Tag = "D" Then
'      pnl_MToGar.Tag = "A"
'      Call gs_SorteaGrid(grd_Listad, 6, "N")
'   Else
'      pnl_MToGar.Tag = "D"
'      Call gs_SorteaGrid(grd_Listad, 6, "N-")
'   End If
'End Sub
'
'Private Sub pnl_NumAd_Click()
'   If Len(Trim(pnl_NumAd.Tag)) = 0 Or pnl_NumAd.Tag = "D" Then
'      pnl_NumAd.Tag = "A"
'      Call gs_SorteaGrid(grd_Listad, 4, "N")
'   Else
'      pnl_NumAd.Tag = "D"
'      Call gs_SorteaGrid(grd_Listad, 4, "N-")
'   End If
'End Sub
'Private Sub pnl_NumCso_Click()
'   If Len(Trim(pnl_NumCso.Tag)) = 0 Or pnl_NumCso.Tag = "D" Then
'      pnl_NumCso.Tag = "A"
'      Call gs_SorteaGrid(grd_Listad, 5, "N")
'   Else
'      pnl_NumCso.Tag = "D"
'      Call gs_SorteaGrid(grd_Listad, 5, "N-")
'   End If
'End Sub
'Private Sub pnl_NumFia_Click()
'   If Len(Trim(pnl_NumFia.Tag)) = 0 Or pnl_NumFia.Tag = "D" Then
'      pnl_NumFia.Tag = "A"
'      Call gs_SorteaGrid(grd_Listad, 3, "N")
'   Else
'      pnl_NumFia.Tag = "D"
'      Call gs_SorteaGrid(grd_Listad, 3, "N-")
'   End If
'End Sub
'
'Private Sub pnl_RazSoc_Click()
'   If Len(Trim(pnl_RazSoc.Tag)) = 0 Or pnl_RazSoc.Tag = "D" Then
'      pnl_RazSoc.Tag = "A"
'      Call gs_SorteaGrid(grd_Listad, 1, "C")
'   Else
'      pnl_RazSoc.Tag = "D"
'      Call gs_SorteaGrid(grd_Listad, 1, "C-")
'   End If
'End Sub
'
'Private Sub pnl_Saldo_Click()
'   If Len(Trim(pnl_Saldo.Tag)) = 0 Or pnl_Saldo.Tag = "D" Then
'      pnl_Saldo.Tag = "A"
'      Call gs_SorteaGrid(grd_Listad, 10, "N")
'   Else
'      pnl_Saldo.Tag = "D"
'      Call gs_SorteaGrid(grd_Listad, 10, "N-")
'   End If
'End Sub
Private Sub cmd_NueGar_Click()
   moddat_g_int_FlgGrb_1 = 4
   If fs_Validar = True Then
      frm_Ges_TecPro_05.Show 1
   End If
End Sub

'Private Sub pnl_Usuario_Click()
'   If Len(Trim(pnl_Usuario.Tag)) = 0 Or pnl_Usuario.Tag = "D" Then
'      pnl_Usuario.Tag = "A"
'      Call gs_SorteaGrid(grd_Listad, 12, "C")
'   Else
'      pnl_Usuario.Tag = "D"
'      Call gs_SorteaGrid(grd_Listad, 12, "C-")
'   End If
'End Sub

Private Sub txt_Buscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Screen.MousePointer = 11
      Call fs_Buscar
      Screen.MousePointer = 0
   Else
      If cmb_Buscar.ListIndex = 1 Then
         KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_NUMERO)
      Else
         KeyAscii = gf_ValidaCaracter(KeyAscii, modgen_g_con_LETRAS & modgen_g_con_NUMERO & "- ()?¿)(/&%$·#@_.,;:")
      End If
   End If
End Sub
