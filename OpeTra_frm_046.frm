VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_Caj_GasCie_01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8730
   ClientLeft      =   3540
   ClientTop       =   1740
   ClientWidth     =   11355
   Icon            =   "OpeTra_frm_046.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11355
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   8745
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11355
      _Version        =   65536
      _ExtentX        =   20029
      _ExtentY        =   15425
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
         Width           =   11295
         _Version        =   65536
         _ExtentX        =   19923
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
            Left            =   10680
            Picture         =   "OpeTra_frm_046.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Salida"
            Top             =   30
            Width           =   585
         End
         Begin VB.CommandButton cmd_GasAdm 
            Height          =   585
            Left            =   30
            Picture         =   "OpeTra_frm_046.frx":044E
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Evaluar Solicitud"
            Top             =   30
            Width           =   585
         End
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   675
         Left            =   30
         TabIndex        =   4
         Top             =   30
         Width           =   11295
         _Version        =   65536
         _ExtentX        =   19923
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
            Height          =   255
            Left            =   630
            TabIndex        =   5
            Top             =   60
            Width           =   7875
            _Version        =   65536
            _ExtentX        =   13891
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "Operaciones Financieras"
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
            Height          =   255
            Left            =   630
            TabIndex        =   13
            Top             =   330
            Width           =   7875
            _Version        =   65536
            _ExtentX        =   13891
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "Créditos Hipotecarios - Cobro de Gastos de Cierre"
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
            Picture         =   "OpeTra_frm_046.frx":0D18
            Top             =   60
            Width           =   480
         End
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   7245
         Left            =   30
         TabIndex        =   6
         Top             =   1440
         Width           =   11295
         _Version        =   65536
         _ExtentX        =   19923
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
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Nro. Solicitud"
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
            Left            =   1560
            TabIndex        =   8
            Top             =   60
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "ID Cliente"
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
            Left            =   3060
            TabIndex        =   9
            Top             =   60
            Width           =   5385
            _Version        =   65536
            _ExtentX        =   9499
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Apellidos y Nombres"
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
         Begin Threed.SSPanel pnl_Tit_Import 
            Height          =   285
            Left            =   9390
            TabIndex        =   10
            Top             =   60
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Importe"
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
            Height          =   6855
            Left            =   30
            TabIndex        =   11
            Top             =   360
            Width           =   11235
            _ExtentX        =   19817
            _ExtentY        =   12091
            _Version        =   393216
            Rows            =   30
            Cols            =   12
            FixedRows       =   0
            FixedCols       =   0
            BackColorSel    =   32768
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin Threed.SSPanel pnl_Tit_Moneda 
            Height          =   285
            Left            =   8430
            TabIndex        =   12
            Top             =   60
            Width           =   975
            _Version        =   65536
            _ExtentX        =   1720
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Moneda"
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
   End
End
Attribute VB_Name = "frm_Caj_GasCie_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_dbl_PorITF     As Double

Private Sub cmd_GasAdm_Click()
   If grd_Listad.Rows = 0 Then
      Exit Sub
   End If

   grd_Listad.Col = 7
   moddat_g_str_NumSol = grd_Listad.Text
   
   grd_Listad.Col = 1
   moddat_g_int_TipDoc = Left(grd_Listad.Text, 1)
   moddat_g_str_NumDoc = Mid(grd_Listad.Text, 3)
         
   grd_Listad.Col = 2
   moddat_g_str_NomCli = grd_Listad.Text
   
   grd_Listad.Col = 5
   moddat_g_str_CodPrd = grd_Listad.Text
   
   grd_Listad.Col = 6
   moddat_g_str_CodSub = grd_Listad.Text
   
   grd_Listad.Col = 8
   moddat_g_int_TipMon = CInt(grd_Listad.Text)
   
   grd_Listad.Col = 9
   moddat_g_str_CodConHip = grd_Listad.Text
   
   grd_Listad.Col = 10
   moddat_g_str_CodEjeSeg = grd_Listad.Text
   
   grd_Listad.Col = 11
   moddat_g_int_CodIns = grd_Listad.Text
   
   Call gs_RefrescaGrid(grd_Listad)
   
   moddat_g_int_FlgAct = 1
   
   frm_Caj_GasCie_02.Show 1
   
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
   grd_Listad.ColWidth(0) = 1505
   grd_Listad.ColWidth(1) = 1515
   grd_Listad.ColWidth(2) = 5375
   grd_Listad.ColWidth(3) = 975
   grd_Listad.ColWidth(4) = 1515
   grd_Listad.ColWidth(5) = 0
   grd_Listad.ColWidth(6) = 0
   grd_Listad.ColWidth(7) = 0
   grd_Listad.ColWidth(8) = 0
   grd_Listad.ColWidth(9) = 0
   grd_Listad.ColWidth(10) = 0
   grd_Listad.ColWidth(11) = 0
   
   grd_Listad.ColAlignment(0) = flexAlignCenterCenter
   grd_Listad.ColAlignment(1) = flexAlignCenterCenter
   grd_Listad.ColAlignment(2) = flexAlignLeftCenter
   grd_Listad.ColAlignment(3) = flexAlignCenterCenter
   grd_Listad.ColAlignment(4) = flexAlignRightCenter

   'Obteniendo ITF
   l_dbl_PorITF = opecaj_gf_Consulta_ITF(Format(CDate(moddat_g_str_FecSis), "yyyymmdd"), 1)
End Sub

Private Sub fs_Buscar()
   Dim r_dbl_ITFGas     As Double
   
   g_str_Parame = "SELECT * FROM CRE_SOLMAE WHERE "
   g_str_Parame = g_str_Parame & "SOLMAE_SITUAC = 1 AND "
   g_str_Parame = g_str_Parame & "SOLMAE_CODINS >= 31 AND "
   g_str_Parame = g_str_Parame & "SOLMAE_CODINS <= 41 "
   g_str_Parame = g_str_Parame & "ORDER BY SOLMAE_NUMERO ASC"
   
   If Not gf_EjecutaSQL(g_str_Parame, g_rst_Princi, 3) Then
      Exit Sub
   End If
   
   grd_Listad.Redraw = False
   
   Call gs_LimpiaGrid(grd_Listad)
   
   If Not (g_rst_Princi.BOF And g_rst_Princi.EOF) Then
      g_rst_Princi.MoveFirst
      
      Do While Not g_rst_Princi.EOF
         'Verificando que los Gastos Administrativos no hayan sido pagados
         g_str_Parame = "SELECT SUM(GASADM_IMPORT) AS TOTGAS FROM TRA_GASADM WHERE "
         g_str_Parame = g_str_Parame & "GASADM_NUMSOL = '" & g_rst_Princi!SOLMAE_NUMERO & "' AND "
         g_str_Parame = g_str_Parame & "GASADM_SITUAC = 2"
      
         If Not gf_EjecutaSQL(g_str_Parame, g_rst_Genera, 3) Then
             Exit Sub
         End If
         
         If Not (g_rst_Genera.BOF And g_rst_Genera.EOF) And g_rst_Genera!TOTGAS > 0 Then
            grd_Listad.Rows = grd_Listad.Rows + 1
            grd_Listad.Row = grd_Listad.Rows - 1
            
            grd_Listad.Col = 0
            grd_Listad.Text = gf_Formato_NumSol(g_rst_Princi!SOLMAE_NUMERO)
            
            grd_Listad.Col = 1
            grd_Listad.Text = CStr(g_rst_Princi!SOLMAE_TITTDO) & "-" & Trim(g_rst_Princi!SOLMAE_TITNDO)
            
            grd_Listad.Col = 2
            grd_Listad.Text = moddat_gf_Buscar_NomCli(CStr(g_rst_Princi!SOLMAE_TITTDO), Trim(g_rst_Princi!SOLMAE_TITNDO))
            
            grd_Listad.Col = 3
            grd_Listad.Text = moddat_gf_Consulta_ParDes("229", CStr(g_rst_Princi!SOLMAE_TIPMON))
            
            r_dbl_ITFGas = CDbl(gf_NueImp_Numero(gf_Truncar_Numero(g_rst_Genera!TOTGAS * (l_dbl_PorITF / 100), 2)))
            
            grd_Listad.Col = 4
            grd_Listad.Text = Format(g_rst_Genera!TOTGAS + r_dbl_ITFGas, "###,###,##0.00")
            
            grd_Listad.Col = 5
            grd_Listad.Text = Trim(g_rst_Princi!SOLMAE_CODPRD & "")
            
            grd_Listad.Col = 6
            grd_Listad.Text = Trim(g_rst_Princi!SOLMAE_CODSUB & "")
            
            grd_Listad.Col = 7
            grd_Listad.Text = Trim(g_rst_Princi!SOLMAE_NUMERO & "")
            
            grd_Listad.Col = 8
            grd_Listad.Text = CStr(g_rst_Princi!SOLMAE_TIPMON)
         
            grd_Listad.Col = 9
            grd_Listad.Text = CStr(g_rst_Princi!SOLMAE_CONHIP)
         
            grd_Listad.Col = 10
            grd_Listad.Text = CStr(g_rst_Princi!SOLMAE_EJESEG)
         
            grd_Listad.Col = 11
            grd_Listad.Text = g_rst_Princi!SOLMAE_CODINS
         End If
      
         g_rst_Genera.Close
         Set g_rst_Genera = Nothing
         
         g_rst_Princi.MoveNext
      Loop
   End If
   
   grd_Listad.Redraw = True
   
   If grd_Listad.Rows = 0 Then
      cmd_GasAdm.Enabled = False
      
      MsgBox "No se encontraron Solicitudes Pendientes de Asignación de Gastos Administrativos.", vbInformation, modgen_g_str_NomPlt
   Else
      'Ordenando por Nombres de Clientes
      pnl_Tit_NomCli.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 2, "C")
   
      Call gs_UbiIniGrid(grd_Listad)
   End If
   
   g_rst_Princi.Close
   Set g_rst_Princi = Nothing
End Sub

Private Sub grd_Listad_DblClick()
   Call cmd_GasAdm_Click
End Sub

Private Sub grd_Listad_SelChange()
   If grd_Listad.Rows > 2 Then
      grd_Listad.RowSel = grd_Listad.Row
   End If
End Sub

Private Sub pnl_Tit_NumSol_Click()
   If Len(Trim(pnl_Tit_NumSol.Tag)) = 0 Or pnl_Tit_NumSol.Tag = "D" Then
      pnl_Tit_NumSol.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 0, "C")
   Else
      pnl_Tit_NumSol.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 0, "C-")
   End If
End Sub

Private Sub pnl_Tit_DocIde_Click()
   If Len(Trim(pnl_Tit_DocIde.Tag)) = 0 Or pnl_Tit_DocIde.Tag = "D" Then
      pnl_Tit_DocIde.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 1, "C")
   Else
      pnl_Tit_DocIde.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 1, "C-")
   End If
End Sub

Private Sub pnl_Tit_NomCli_Click()
   If Len(Trim(pnl_Tit_NomCli.Tag)) = 0 Or pnl_Tit_NomCli.Tag = "D" Then
      pnl_Tit_NomCli.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 2, "C")
   Else
      pnl_Tit_NomCli.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 2, "C-")
   End If
End Sub

Private Sub pnl_Tit_Moneda_Click()
   If Len(Trim(pnl_Tit_Moneda.Tag)) = 0 Or pnl_Tit_Moneda.Tag = "D" Then
      pnl_Tit_Moneda.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 3, "C")
   Else
      pnl_Tit_Moneda.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 3, "C-")
   End If
End Sub

Private Sub pnl_Tit_Import_Click()
   If Len(Trim(pnl_Tit_Import.Tag)) = 0 Or pnl_Tit_Import.Tag = "D" Then
      pnl_Tit_Import.Tag = "A"
      Call gs_SorteaGrid(grd_Listad, 4, "N")
   Else
      pnl_Tit_Import.Tag = "D"
      Call gs_SorteaGrid(grd_Listad, 4, "N-")
   End If
End Sub



