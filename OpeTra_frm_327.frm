VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_Ges_CreHip_19 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form4"
   ClientHeight    =   7725
   ClientLeft      =   1650
   ClientTop       =   5010
   ClientWidth     =   8355
   Icon            =   "OpeTra_frm_327.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   8355
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   7725
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8400
      _Version        =   65536
      _ExtentX        =   14817
      _ExtentY        =   13626
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
         Height          =   645
         Left            =   30
         TabIndex        =   1
         Top             =   750
         Width           =   8295
         _Version        =   65536
         _ExtentX        =   14631
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
            Left            =   7680
            Picture         =   "OpeTra_frm_327.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   3
         Top             =   30
         Width           =   8295
         _Version        =   65536
         _ExtentX        =   14631
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
            TabIndex        =   4
            Top             =   60
            Width           =   6345
            _Version        =   65536
            _ExtentX        =   11192
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Créditos Hipotecarios"
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
         Begin Threed.SSPanel SSPanel4 
            Height          =   285
            Left            =   660
            TabIndex        =   5
            Top             =   360
            Width           =   6345
            _Version        =   65536
            _ExtentX        =   11192
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Clasificación de Clientes"
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
            Picture         =   "OpeTra_frm_327.frx":044E
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   6255
         Left            =   30
         TabIndex        =   6
         Top             =   1440
         Width           =   8295
         _Version        =   65536
         _ExtentX        =   14631
         _ExtentY        =   11033
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
         Begin Threed.SSPanel pnl_Tit_Produc 
            Height          =   285
            Left            =   60
            TabIndex        =   7
            Top             =   60
            Width           =   1000
            _Version        =   65536
            _ExtentX        =   1764
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Año"
            ForeColor       =   16777215
            BackColor       =   16384
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
         Begin Threed.SSPanel pnl_Tit_NumSol 
            Height          =   285
            Left            =   1050
            TabIndex        =   8
            Top             =   60
            Width           =   1740
            _Version        =   65536
            _ExtentX        =   3069
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Mes"
            ForeColor       =   16777215
            BackColor       =   16384
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
         Begin Threed.SSPanel pnl_Tit_DocIde 
            Height          =   285
            Left            =   2760
            TabIndex        =   9
            Top             =   60
            Width           =   2600
            _Version        =   65536
            _ExtentX        =   4586
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Clasificación Interna"
            ForeColor       =   16777215
            BackColor       =   16384
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
         Begin Threed.SSPanel pnl_Tit_FecSol 
            Height          =   285
            Left            =   5340
            TabIndex        =   10
            Top             =   60
            Width           =   2600
            _Version        =   65536
            _ExtentX        =   4586
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Clasificación Alineada"
            ForeColor       =   16777215
            BackColor       =   16384
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
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   5805
            Left            =   30
            TabIndex        =   11
            Top             =   360
            Width           =   8265
            _ExtentX        =   14579
            _ExtentY        =   10239
            _Version        =   393216
            Rows            =   30
            Cols            =   5
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
      End
   End
End
Attribute VB_Name = "frm_Ges_CreHip_19"
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
   Call fs_Limpia
   Call fs_Buscar
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   'Inicializando Grid de Datos del Cliente
   grd_Listad.ColWidth(0) = 0
   grd_Listad.ColWidth(1) = 1000
   grd_Listad.ColWidth(2) = 1700
   grd_Listad.ColWidth(3) = 2600
   grd_Listad.ColWidth(4) = 2600
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignCenterCenter
   grd_Listad.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad.ColAlignment(4) = flexAlignCenterCenter
End Sub

Private Sub fs_Limpia()
   Call gs_LimpiaGrid(grd_Listad)
End Sub

Private Sub fs_Buscar()
   'Buscando Información del Crédito
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT  HIPCIE_PERANO "
   g_str_Parame = g_str_Parame & ", HIPCIE_PERMES "
   g_str_Parame = g_str_Parame & ", TRIM(TO_CHAR(TO_DATE('2012/'||TRIM(HIPCIE_PERMES)|| '/01', 'yyyy/mm/dd'), 'MONTH','NLS_DATE_LANGUAGE=SPANISH' ) ) AS MES "
   g_str_Parame = g_str_Parame & ", TRIM(C.TIPCLA_DESCRI) AS CLACLI "
   g_str_Parame = g_str_Parame & ", TRIM(A.TIPCLA_DESCRI) AS CLAALI "
   g_str_Parame = g_str_Parame & ", HIPCIE_CLACLI, HIPCIE_CLAALI, HIPCIE_CLAPRV "
   g_str_Parame = g_str_Parame & "FROM CRE_HIPCIE "
   g_str_Parame = g_str_Parame & "INNER JOIN CTB_TIPCLA C ON (C.TIPCLA_CODIGO = HIPCIE_CLACLI) "
   g_str_Parame = g_str_Parame & "INNER JOIN CTB_TIPCLA A ON (A.TIPCLA_CODIGO = HIPCIE_CLAPRV) "
   g_str_Parame = g_str_Parame & "WHERE HIPCIE_NUMOPE = '" & moddat_g_str_NumOpe & "' "
   g_str_Parame = g_str_Parame & "AND C.TIPCLA_TIPCRE = '13' "
   g_str_Parame = g_str_Parame & "AND A.TIPCLA_TIPCRE = '13' "
   g_str_Parame = g_str_Parame & "ORDER BY HIPCIE_PERANO DESC, HIPCIE_PERMES DESC "
   
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
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
                         
         grd_Listad.Col = 1
         grd_Listad.Text = CStr(g_rst_Princi!HIPCIE_PERANO)
           
         grd_Listad.Col = 2
         grd_Listad.Text = CStr(g_rst_Princi!mes)
         
         grd_Listad.Col = 3
         grd_Listad.Text = CStr(g_rst_Princi!CLACLI)
         
         grd_Listad.Col = 4
         grd_Listad.Text = CStr(g_rst_Princi!CLAALI)
         g_rst_Princi.MoveNext
      Loop
      
      Call gs_UbiIniGrid(grd_Listad)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub
 
