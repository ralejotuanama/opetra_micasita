VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_TraMVi_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8835
   ClientLeft      =   900
   ClientTop       =   1125
   ClientWidth     =   14130
   Icon            =   "OpeTra_frm_041.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8835
   ScaleWidth      =   14130
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   8835
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14115
      _Version        =   65536
      _ExtentX        =   24897
      _ExtentY        =   15584
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
         Height          =   735
         Left            =   30
         TabIndex        =   1
         Top             =   8040
         Width           =   14025
         _Version        =   65536
         _ExtentX        =   24739
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
         Begin VB.CommandButton cmd_Evalua 
            Height          =   675
            Left            =   12630
            Picture         =   "OpeTra_frm_041.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Registrar Evaluación Mivivienda"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmd_Salida 
            Height          =   675
            Left            =   13320
            Picture         =   "OpeTra_frm_041.frx":08D6
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Salir"
            Top             =   30
            Width           =   675
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   4
         Top             =   30
         Width           =   14025
         _Version        =   65536
         _ExtentX        =   24739
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
            Height          =   495
            Left            =   630
            TabIndex        =   5
            Top             =   60
            Width           =   5685
            _Version        =   65536
            _ExtentX        =   10028
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Trámites Mivivienda"
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
            Picture         =   "OpeTra_frm_041.frx":0D18
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   7245
         Left            =   30
         TabIndex        =   6
         Top             =   750
         Width           =   14025
         _Version        =   65536
         _ExtentX        =   24739
         _ExtentY        =   12779
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
         Begin Threed.SSPanel pnl_Tit_NumSol 
            Height          =   285
            Left            =   60
            TabIndex        =   7
            Top             =   60
            Width           =   3705
            _Version        =   65536
            _ExtentX        =   6535
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Producto"
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
            Left            =   3750
            TabIndex        =   8
            Top             =   60
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro. Solicitud"
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
            Left            =   5250
            TabIndex        =   9
            Top             =   60
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "ID Cliente"
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
            Left            =   6750
            TabIndex        =   10
            Top             =   60
            Width           =   5385
            _Version        =   65536
            _ExtentX        =   9499
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Apellidos y Nombres"
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
         Begin Threed.SSPanel SSPanel9 
            Height          =   285
            Left            =   12120
            TabIndex        =   11
            Top             =   60
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Fecha Registro"
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
         Begin MSFlexGridLib.MSFlexGrid grd_Listad 
            Height          =   6855
            Left            =   30
            TabIndex        =   12
            Top             =   360
            Width           =   13935
            _ExtentX        =   24580
            _ExtentY        =   12091
            _Version        =   393216
            Rows            =   30
            Cols            =   9
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   49152
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
      End
   End
End
Attribute VB_Name = "frm_TraMVi_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Evalua_Click()
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If
   
   grd_Listad.Col = 0
   moddat_g_str_NomPrd = grd_Listad.Text

   grd_Listad.Col = 1
   moddat_g_str_NumSol = Mid(grd_Listad.Text, 1, 3) & Mid(grd_Listad.Text, 5, 3) & Mid(grd_Listad.Text, 9, 2) & Mid(grd_Listad.Text, 12, 4)
   
   grd_Listad.Col = 2
   moddat_g_int_TipDoc = Left(grd_Listad.Text, 1)
   moddat_g_str_NumDoc = Mid(grd_Listad.Text, 3)
         
   grd_Listad.Col = 3
   moddat_g_str_NomCli = grd_Listad.Text
   
   grd_Listad.Col = 4
   moddat_g_str_FecIng = grd_Listad.Text
   
   grd_Listad.Col = 5
   moddat_g_str_CodPrd = grd_Listad.Text
   
   grd_Listad.Col = 6
   moddat_g_str_CodSub = grd_Listad.Text
   
   grd_Listad.Col = 8
   moddat_g_int_TipMon = CInt(grd_Listad.Text)
   
   Call gs_RefrescaGrid(grd_Listad)
   
   moddat_g_int_FlgAct = 1
   
   frm_TraMVi_02.Show 1
   
   If moddat_g_int_FlgAct = 2 Then
      Screen.MousePointer = 11
      
      Call fs_Buscar
      
      Screen.MousePointer = 0
   End If
End Sub

Private Sub cmd_Salida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11

   Me.Caption = modgen_g_str_NomPlt
   
   Call fs_Inicia
   Call fs_Buscar
   
   Call gs_CentraForm(Me)
   
   Screen.MousePointer = 0
End Sub

Private Sub fs_Inicia()
   'Inicializando Rejilla
   grd_Listad.ColWidth(0) = 3675
   grd_Listad.ColWidth(1) = 1515
   grd_Listad.ColWidth(2) = 1505
   grd_Listad.ColWidth(3) = 5365
   grd_Listad.ColWidth(4) = 1505
   grd_Listad.ColWidth(5) = 0
   grd_Listad.ColWidth(6) = 0
   grd_Listad.ColWidth(7) = 0
   grd_Listad.ColWidth(8) = 0
   
   grd_Listad.ColAlignment(0) = flexAlignLeftCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignCenterCenter
   grd_Listad.ColAlignment(3) = flexAlignLeftCenter
   grd_Listad.ColAlignment(4) = flexAlignCenterCenter
End Sub

Private Sub fs_Buscar()
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 AND "
   g_str_Parame = g_str_Parame & "SOLMAE_CODINS = 61 AND "
   g_str_Parame = g_str_Parame & "SOLMAE_CODPRD = '001' "
   g_str_Parame = g_str_Parame & "ORDER BY SOLMAE_NUMERO ASC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   Call gs_LimpiaGrid(grd_Listad)
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         g_str_Parame = "SELECT * FROM TRA_SEGUIM WHERE "
         g_str_Parame = g_str_Parame & "SEGUIM_NUMSOL = '" & g_rst_Princi!SOLMAE_NUMERO & "' AND "
         g_str_Parame = g_str_Parame & "SEGUIM_CODINS = 62 AND "
         g_str_Parame = g_str_Parame & "SEGUIM_SITUAC = 9 "
         
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
            Exit Sub
         End If
         
         If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) Then
            g_rst_Genera.MoveFirst
      
            grd_Listad.Rows = grd_Listad.Rows + 1
            grd_Listad.Row = grd_Listad.Rows - 1
            
            grd_Listad.Col = 0
            grd_Listad.Text = moddat_gf_Consulta_Produc(g_rst_Princi!SOLMAE_CODPRD)
            
            grd_Listad.Col = 1
            grd_Listad.Text = Left(g_rst_Princi!SOLMAE_NUMERO, 3) & "-" & Mid(g_rst_Princi!SOLMAE_NUMERO, 4, 3) & "-" & Mid(g_rst_Princi!SOLMAE_NUMERO, 7, 2) & "-" & Right(g_rst_Princi!SOLMAE_NUMERO, 4)
            
            grd_Listad.Col = 2
            grd_Listad.Text = CStr(g_rst_Princi!SOLMAE_TITTDO) & "-" & Trim(g_rst_Princi!SOLMAE_TITNDO)
            
            grd_Listad.Col = 3
            grd_Listad.Text = moddat_gf_Buscar_NomCli(CStr(g_rst_Princi!SOLMAE_TITTDO), Trim(g_rst_Princi!SOLMAE_TITNDO))
            
            grd_Listad.Col = 4
            grd_Listad.Text = gf_FormatoFecha(CStr(g_rst_Princi!SOLMAE_FECSOL))
            
            grd_Listad.Col = 5
            grd_Listad.Text = Trim(g_rst_Princi!SOLMAE_CODPRD & "")
            
            grd_Listad.Col = 6
            grd_Listad.Text = Trim(g_rst_Princi!SOLMAE_CODSUB & "")
            
            grd_Listad.Col = 7
            grd_Listad.Text = g_rst_Princi!SOLMAE_FECSOL
            
            grd_Listad.Col = 8
            grd_Listad.Text = CStr(g_rst_Princi!SOLMAE_TIPMON)
         End If
            
         g_rst_Genera.Close
         Set g_rst_Genera = Nothing
         
         g_rst_Princi.MoveNext
      Loop
   End If
   
   If grd_Listad.Rows = 0 Then
      cmd_Evalua.Enabled = False
      
      MsgBox "No se encontraron Solicitudes Pendientes de Trámites Mivivienda.", vbInformation, modgen_g_str_NomPlt
   Else
      Call gs_UbiIniGrid(grd_Listad)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub grd_Listad_DblClick()
   Call cmd_Evalua_Click
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub







