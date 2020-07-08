VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_Tas_ActReg_02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6360
   Icon            =   "OpeTra_frm_336.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel3 
      Height          =   4995
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6345
      _Version        =   65536
      _ExtentX        =   11192
      _ExtentY        =   8811
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
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   30
         TabIndex        =   4
         Top             =   720
         Width           =   6285
         _Version        =   65536
         _ExtentX        =   11086
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
            Left            =   5670
            Picture         =   "OpeTra_frm_336.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   585
         End
         Begin Threed.SSPanel pnl_NumSol 
            Height          =   315
            Left            =   1290
            TabIndex        =   0
            Top             =   180
            Width           =   2115
            _Version        =   65536
            _ExtentX        =   3731
            _ExtentY        =   556
            _StockProps     =   15
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            Font3D          =   2
         End
         Begin VB.Label Label1 
            Caption         =   "Nro. Solicitud"
            Height          =   285
            Left            =   150
            TabIndex        =   5
            Top             =   180
            Width           =   1125
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   3585
         Left            =   30
         TabIndex        =   6
         Top             =   1380
         Width           =   6285
         _Version        =   65536
         _ExtentX        =   11086
         _ExtentY        =   6324
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
            Height          =   3165
            Left            =   30
            TabIndex        =   2
            Top             =   390
            Width           =   6195
            _ExtentX        =   10927
            _ExtentY        =   5583
            _Version        =   393216
            Rows            =   30
            Cols            =   3
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_Tit_DocIde 
            Height          =   285
            Left            =   60
            TabIndex        =   7
            Top             =   90
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Fecha Registro"
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
         Begin Threed.SSPanel pnl_Tit_NomCli 
            Height          =   285
            Left            =   1245
            TabIndex        =   8
            Top             =   90
            Width           =   4665
            _Version        =   65536
            _ExtentX        =   8229
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Perito Tasador"
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
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   9
         Top             =   30
         Width           =   6285
         _Version        =   65536
         _ExtentX        =   11086
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
            Left            =   600
            TabIndex        =   10
            Top             =   30
            Width           =   5415
            _Version        =   65536
            _ExtentX        =   9551
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Tasación del Inmueble"
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
         Begin Threed.SSPanel SSPanel15 
            Height          =   285
            Left            =   600
            TabIndex        =   11
            Top             =   300
            Width           =   5415
            _Version        =   65536
            _ExtentX        =   9551
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Historico de Tasaciones"
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
            Picture         =   "OpeTra_frm_336.frx":044E
            Top             =   60
            Width           =   480
         End
      End
   End
End
Attribute VB_Name = "frm_Tas_ActReg_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub fs_Inicia()
   grd_Listad.ColWidth(0) = 1200
   grd_Listad.ColWidth(1) = 4600
   grd_Listad.ColWidth(2) = 0
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignLeftCenter
   grd_Listad.ColAlignment(2) = flexAlignLeftCenter
End Sub

Private Sub fs_Buscar_Historico()
   moddat_g_int_CntErr = 1
   Call gs_LimpiaGrid(grd_Listad)
   
   'Buscando Actual
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT EVATAS_FECEVA, EVATAS_NOMPER "
   g_str_Parame = g_str_Parame & "  FROM TRA_EVATAS "
   g_str_Parame = g_str_Parame & " WHERE EVATAS_NUMSOL = '" & moddat_g_str_NumSol & "' "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      moddat_g_int_CntErr = 2
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
            
         grd_Listad.Col = 0
         grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!EVATAS_FECEVA))
            
         grd_Listad.Col = 1
         grd_Listad.Text = Trim(g_rst_Princi!EVATAS_NOMPER)
          
         grd_Listad.Col = 2
         grd_Listad.Text = "A"
          
         g_rst_Princi.MoveNext
      Loop
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
   
   'Buscando Historico
   g_str_Parame = ""
   g_str_Parame = g_str_Parame & "SELECT EVATAS_FECEVA, EVATAS_NOMPER "
   g_str_Parame = g_str_Parame & "  FROM HIS_EVATAS "
   g_str_Parame = g_str_Parame & " WHERE EVATAS_NUMSOL = '" & moddat_g_str_NumSol & "' "
   g_str_Parame = g_str_Parame & " ORDER BY EVATAS_NUMSOL "
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      moddat_g_int_CntErr = 2
      Exit Sub
   End If
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      Do While Not g_rst_Princi.EOF
         grd_Listad.Rows = grd_Listad.Rows + 1
         grd_Listad.Row = grd_Listad.Rows - 1
            
         grd_Listad.Col = 0
         grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!EVATAS_FECEVA))
            
         grd_Listad.Col = 1
         grd_Listad.Text = Trim(g_rst_Princi!EVATAS_NOMPER)
          
         grd_Listad.Col = 2
         grd_Listad.Text = "H"
         
         g_rst_Princi.MoveNext
      Loop
   End If
   
   Call gs_UbiIniGrid(grd_Listad)
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Me.Caption = modgen_g_str_NomPlt
   pnl_NumSol.Caption = gf_Formato_NumSol(moddat_g_str_NumSol)
   
   Call fs_Inicia
   Call fs_Buscar_Historico
   
   Call gs_CentraForm(Me)
   Screen.MousePointer = 0
End Sub

Private Sub grd_Listad_DblClick()
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   frm_Tas_ActReg_03.l_str_FecTas = grd_Listad.TextMatrix(grd_Listad.Row, 0)
   frm_Tas_ActReg_03.l_str_TipReg = grd_Listad.TextMatrix(grd_Listad.Row, 2)
   frm_Tas_ActReg_03.Show 1
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub
